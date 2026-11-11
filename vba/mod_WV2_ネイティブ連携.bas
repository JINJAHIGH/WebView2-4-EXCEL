Attribute VB_Name = "mod_WV2_ネイティブ連携"
Option Explicit

' -------------
' ポインターの基本値
' -------------
Private Const 空ポインター As LongLong = 0^
Private Const ポインター長 As Long = 8

' -------------
' ウィンドウ操作
' -------------
Public Const 最上位親を取得         As Long = 2
Public Const ウィンドウスタイル位置 As Long = -16
Public Const サイズ変更枠           As Long = &H40000
Public Const サイズ変更なし         As Long = &H1&
Public Const 位置変更なし           As Long = &H2&
Public Const 重なり順変更なし       As Long = &H4&
Public Const フレーム更新           As Long = &H20&

' -------------
' COM呼び出しと文字コード
' -------------
Public Const UTF8コードページ          As Long = 65001
Private Const 正常終了                 As Long = 0
Private Const 対応インターフェースなし As Long = &H80004002
Private Const 標準呼出規約             As Long = 4

' -------------
' メモリー操作
' -------------
Private Const メモリー実体確保 As Long = &H1000&
Private Const メモリー領域予約 As Long = &H2000&
Private Const メモリー領域解放 As Long = &H8000&
Private Const 未使用領域       As Long = &H10000
Private Const 実行読書可能     As Long = &H40&
Private Const 読書可能         As Long = &H4&

' -------------
' 前回領域の追跡
' -------------
Private Const 前回領域の保存変数名   As String = "WV2_VBA_LastRegion"
Private Const 前回領域アドレス文字数 As Long = 32

' -------------
' 中継領域の配置（バイト単位）
' -------------
Private Const 遅延呼出差込位置        As Long = 55
Private Const 中継元命令長            As Long = 91
Private Const 引数変換命令長          As Long = 74
Private Const 引数変換命令位置        As Long = 96
Public Const 仮想COMオブジェクト位置 As Long = 176
Private Const スロットバイト数        As Long = 224
Private Const スロット数              As Long = 512
Private Const 共通ヘッダー長          As Long = 64
Private Const 中継領域バイト数        As Long = 共通ヘッダー長 + スロットバイト数 * スロット数
Private Const 命令作成領域長          As Long = 80

Private Enum 配列属性
    配列属性自動 = &H1
    配列属性固定長 = &H10
End Enum

Private Type 配列境界
    要素数 As Long
    下限値 As Long
End Type

Private Type 一次元配列情報
    次元数 As Integer
    属性値 As Integer
    要素バイト数 As Long
    ロック数 As Long
    データアドレス As LongPtr
    先頭次元境界 As 配列境界
End Type

Private Type ポインター参照情報
    参照配列() As LongPtr
    配列情報 As 一次元配列情報
End Type

Public Declare PtrSafe Function WebView2環境生成API Lib "WebView2Loader.dll" Alias "CreateCoreWebView2EnvironmentWithOptions" (ByVal ブラウザー実行環境フォルダー As LongPtr, ByVal ユーザーデータフォルダー As LongPtr, ByVal 追加ブラウザー引数一覧 As LongPtr, _
    ByVal 環境作成済みハンドラー As LongPtr) As Long
    
Private Declare PtrSafe Function 仮想メモリー確保API Lib "kernel32" Alias "VirtualAlloc" (ByVal 対象アドレス As LongPtr, ByVal 要求バイト数 As LongPtr, ByVal 確保方式 As Long, ByVal メモリー保護属性 As Long) As LongPtr

Private Declare PtrSafe Function 仮想メモリー解放API Lib "kernel32" Alias "VirtualFree" (ByVal 対象アドレス As LongPtr, ByVal 要求バイト数 As LongPtr, ByVal 解放方式 As Long) As Long

Private Declare PtrSafe Function COM関数呼出API Lib "oleaut32" Alias "DispCallFunc" (ByVal インスタンスアドレス As LongPtr, ByVal 仮想関数オフセット As LongPtr, ByVal 呼出規約 As Long, ByVal 戻り値の型 As Integer, ByVal 実引数数 As Long, ByRef 引数型配列 As _
    Any, ByRef 引数アドレス配列 As Any, ByRef 戻り値領域 As Any) As Long
    
Public Declare PtrSafe Function クライアント矩形取得API Lib "user32" Alias "GetClientRect" (ByVal ウィンドウハンドル As LongPtr, ByRef 矩形出力先 As 矩形) As Long

Public Declare PtrSafe Function 親ウィンドウ取得API Lib "user32" Alias "GetAncestor" (ByVal ウィンドウハンドル As LongPtr, ByVal 親取得方式 As Long) As LongPtr

Public Declare PtrSafe Function ウィンドウ属性取得API Lib "user32" Alias "GetWindowLongPtrW" (ByVal ウィンドウハンドル As LongPtr, ByVal 属性番号 As Long) As LongPtr

Public Declare PtrSafe Function ウィンドウ属性設定API Lib "user32" Alias "SetWindowLongPtrW" (ByVal ウィンドウハンドル As LongPtr, ByVal 属性番号 As Long, ByVal 設定属性値 As LongPtr) As LongPtr

Public Declare PtrSafe Function ウィンドウ位置設定API Lib "user32" Alias "SetWindowPos" (ByVal ウィンドウハンドル As LongPtr, ByVal 重なり順基準ハンドル As LongPtr, ByVal 横位置 As Long, ByVal 縦位置 As Long, ByVal 横幅 As Long, ByVal 高さ As Long, ByVal _
    位置設定フラグ As Long) As Long
    
Public Declare PtrSafe Function 文字列長取得API Lib "kernel32" Alias "lstrlenW" (ByVal 文字列アドレス As LongPtr) As Long

Public Declare PtrSafe Sub COMメモリー解放API Lib "ole32" Alias "CoTaskMemFree" (ByVal 引数領域 As LongPtr)

Public Declare PtrSafe Sub メモリー転送API Lib "kernel32" Alias "RtlMoveMemory" (ByVal 転送先 As LongPtr, ByVal 元データ As LongPtr, ByVal 長さ As LongPtr)

Public Declare PtrSafe Function Unicode文字列変換API Lib "kernel32" Alias "MultiByteToWideChar" (ByVal コードページ As Long, ByVal 変換フラグ As Long, ByVal 入力バイト列アドレス As LongPtr, ByVal 入力バイト数 As Long, ByVal 出力文字列アドレス As LongPtr, ByVal _
 _
    出力可能文字数 As Long) As Long
    
Private Declare PtrSafe Function 環境変数取得API Lib "kernel32" Alias "GetEnvironmentVariableW" (ByVal 変数名アドレス As LongPtr, ByVal 出力先アドレス As LongPtr, ByVal 出力領域文字数 As Long) As Long

Private Declare PtrSafe Function 環境変数設定API Lib "kernel32" Alias "SetEnvironmentVariableW" (ByVal 変数名アドレス As LongPtr, ByVal 設定文字列アドレス As LongPtr) As Long

Private Type メモリー基本情報
    領域先頭       As LongPtr
    確保先頭    As LongPtr
    確保時保護属性 As Long
    整列領域1               As Long
    領域サイズ        As LongPtr
    状態             As Long
    保護属性           As Long
    領域種別             As Long
    整列領域2               As Long
End Type

Private Declare PtrSafe Function メモリー情報取得API Lib "kernel32" Alias "VirtualQuery" (ByVal 対象アドレス As LongPtr, ByRef 出力先アドレス As メモリー基本情報, ByVal 構造体バイト数 As LongPtr) As LongPtr

Public Type 矩形
    左端   As Long
    上端    As Long
    右端  As Long
    下端 As Long
End Type

Public Type GUID情報
    先頭値 As Long
    第2値 As Integer
    第3値 As Integer
    後半値(0 To 7) As Byte
End Type

Public Enum 通知区分
    通知なし = 0
    環境生成完了通知 = 1
    表示生成完了通知 = 2
    遷移開始通知 = 3
    遷移完了通知 = 4
    Webメッセージ通知 = 5
    タイトル変更通知 = 6
    新規ウィンドウ通知 = 7
    スクリプト完了通知 = 8
    履歴変更通知 = 9
    DOM読込通知 = 10
    通信要求通知 = 11
    通信応答通知 = 12
    応答本文取得通知 = 13
End Enum

Private 照会関数アドレス       As LongPtr
Private 参照追加関数アドレス   As LongPtr
Private 参照解放関数アドレス  As LongPtr

Private 中継領域先頭 As LongPtr
Private 未使用スロット先頭    As Long
Private 次の未使用スロット()  As Long
Private 使用中スロット数       As Long

Private 登録ハンドラー(0 To スロット数 - 1) As mod_WV2_COM通知ハンドラー

Private 通知識別子一覧(通知なし To 応答本文取得通知) As GUID情報

Private IUnknown識別子 As GUID情報

' =========================================================================
' 処理概要       : ネイティブ呼び出しの中継先アドレスを確保
' =========================================================================
Private Sub ネイティブ中継入口(): End Sub

' =========================================================================
' 処理概要       : 通知先を設定したCOMコールバックハンドラーを取得
' 引数           : 通知種別 (通知区分) - コールバックまたはイベントの種別
'                  所有元 (Object) - コールバックを通知する所有元
' 戻り値         : mod_WV2_COM通知ハンドラー - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Public Function 通知先を設定したCOMコールバックハンドラーを取得(ByVal 通知種別 As 通知区分, ByVal 所有元 As Object) As mod_WV2_COM通知ハンドラー
    Dim 通知ハンドラー As mod_WV2_COM通知ハンドラー
    Dim ポインターハンドラー呼出先 As LongPtr
    Dim ポインタースロット As LongPtr
    
    If 中継領域先頭 = 0 Then
        If Not COMコールバックを中継するネイティブ領域を初期化() Then Exit Function
    End If
    
    Set 通知ハンドラー = New mod_WV2_COM通知ハンドラー
    
    ポインターハンドラー呼出先 = クラスの固定vTable位置からメソッドアドレスを取得(通知ハンドラー, 7)
    If ポインターハンドラー呼出先 = 0 Then Exit Function
    
    ポインタースロット = ハンドラー用の中継スロットを割り当て(通知ハンドラー, ObjPtr(通知ハンドラー), ポインターハンドラー呼出先)
    If ポインタースロット = 0 Then Exit Function
    
    通知ハンドラー.初期化 通知種別, 所有元, ポインタースロット
    
    Set 通知先を設定したCOMコールバックハンドラーを取得 = 通知ハンドラー
End Function

' =========================================================================
' 処理概要       : 必要な保存先フォルダーを作成
' 引数           : パス (String) - 作成するフォルダーのパス
' =========================================================================
Public Sub 必要な保存先フォルダーを作成(ByVal パス As String)
    Dim ファイル操作 As Object, 親パス As String
    
    If Len(パス) = 0 Then Exit Sub
    
    Set ファイル操作 = CreateObject("Scripting.FileSystemObject")
    If ファイル操作.FolderExists(パス) Then Exit Sub
    
    親パス = ファイル操作.GetParentFolderName(パス)
    If Len(親パス) > 0 Then
        If Not ファイル操作.FolderExists(親パス) Then 必要な保存先フォルダーを作成 親パス
    End If
    ファイル操作.CreateFolder パス
End Sub

' =========================================================================
' 処理概要       : ネイティブ中継領域を終了
' =========================================================================
Public Sub ネイティブ中継領域を終了()
    生存フラグを無効化してハンドラーと中継領域を解放
End Sub

' =========================================================================
' 処理概要       : vTable上のCOMメソッドを引数配列で呼び出し
' 引数           : ポインターインターフェース (LongPtr) - 呼び出すCOMインターフェースのポインター
'                  仮想関数位置番号 (Long) - 呼び出すCOMメソッドのvTable位置
'                  関数名称 (String) - 呼び出すCOMメソッドの識別名
'                  引数一覧 (Variant()) - COMメソッドへ渡す可変長引数
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function COMメソッド呼出(ByVal ポインターインターフェース As LongPtr, ByVal 仮想関数位置番号 As Long, ByVal 関数名称 As String, ParamArray 引数一覧() As Variant) As Long
    Dim 引数数 As Long
    Dim 実行結果値 As Variant
    Dim 実行結果 As Long
    Dim 引数型一覧() As Integer
    Dim メソッドアドレス() As LongPtr
    Dim 値一覧() As Variant
    Dim 添字 As Long
    
    If ポインターインターフェース = 0 Then
    
        COMメソッド呼出 = &H80004003
        Exit Function
    End If
    
    引数数 = UBound(引数一覧) - LBound(引数一覧) + 1
    If 引数数 < 0 Then 引数数 = 0
    
    If 引数数 = 0 Then
    
        実行結果 = COM関数呼出API(ポインターインターフェース, 仮想関数位置番号 * ポインター長, 標準呼出規約, vbLong, 0, ByVal 0&, ByVal 0&, 実行結果値)
    Else
    
        ReDim 引数型一覧(0 To 引数数 - 1)
        ReDim メソッドアドレス(0 To 引数数 - 1)
        ReDim 値一覧(0 To 引数数 - 1)
        
        For 添字 = 0 To 引数数 - 1
            値一覧(添字) = 引数一覧(LBound(引数一覧) + 添字)
            Select Case VarType(値一覧(添字))
                Case vbLong:     引数型一覧(添字) = vbLong
                Case vbLongLong: 引数型一覧(添字) = vbLongLong
                Case vbDouble:   引数型一覧(添字) = vbDouble
                Case Else:       引数型一覧(添字) = vbLongLong
            End Select
            メソッドアドレス(添字) = VarPtr(値一覧(添字))
        Next 添字
        
        実行結果 = COM関数呼出API(ポインターインターフェース, 仮想関数位置番号 * ポインター長, 標準呼出規約, vbLong, 引数数, 引数型一覧(0), メソッドアドレス(0), 実行結果値)
    End If
    
    If 実行結果 <> 0 Then
    
        COMメソッド呼出 = 実行結果
    Else
    
        COMメソッド呼出 = CLng(実行結果値)
    End If
End Function

' =========================================================================
' 処理概要       : COMインターフェースの参照カウントを減算
' 引数           : ポインターインターフェース (LongPtr) - 呼び出すCOMインターフェースのポインター
' 戻り値         : Long - 更新後のCOM参照カウント
' =========================================================================
Public Function COM参照解放(ByVal ポインターインターフェース As LongPtr) As Long
    If ポインターインターフェース <> 0 Then COM参照解放 = COMメソッド呼出(ポインターインターフェース, 2, "Release")
End Function

' =========================================================================
' 処理概要       : COMインターフェースの参照カウントを加算
' 引数           : ポインターインターフェース (LongPtr) - 呼び出すCOMインターフェースのポインター
' 戻り値         : Long - 更新後のCOM参照カウント
' =========================================================================
Public Function COM参照追加(ByVal ポインターインターフェース As LongPtr) As Long
    If ポインターインターフェース <> 0 Then COM参照追加 = COMメソッド呼出(ポインターインターフェース, 1, "AddRef")
End Function

' =========================================================================
' 処理概要       : UTF-16文字列ポインターをVBA文字列へ変換
' 引数           : 対象参照 (LongPtr) - 読み取る文字列またはメモリーのアドレス
' 戻り値         : String - UTF-16文字列ポインターをVBA文字列
' =========================================================================
Public Function UTF16文字列ポインターをVBA文字列へ変換(ByVal 対象参照 As LongPtr) As String
    Dim 文字数 As Long
    
    If 対象参照 = 0 Then Exit Function
    
    文字数 = 文字列長取得API(対象参照)
    If 文字数 = 0 Then Exit Function
    
    UTF16文字列ポインターをVBA文字列へ変換 = String$(文字数, vbNullChar)
    メモリー転送API StrPtr(UTF16文字列ポインターをVBA文字列へ変換), 対象参照, CLngPtr(文字数 * 2)
End Function

' =========================================================================
' 処理概要       : UTF-8のバイト配列をVBA文字列へ変換
' 引数           : 入力データ (Byte()) - 変換対象のバイト配列
'                  バイト数 (Long) - 読み取るバイト数
' 戻り値         : String - UTF-8のバイト配列をVBA文字列
' =========================================================================
Public Function UTF8のバイト配列をVBA文字列へ変換(ByRef 入力データ() As Byte, ByVal バイト数 As Long) As String
    Dim 文字数 As Long
    Dim 出力文字列 As String
    
    If バイト数 <= 0 Then Exit Function
    
    文字数 = Unicode文字列変換API(UTF8コードページ, 0, VarPtr(入力データ(LBound(入力データ))), バイト数, 0, 0)
    If 文字数 <= 0 Then Exit Function
    
    出力文字列 = String$(文字数, vbNullChar)
    文字数 = Unicode文字列変換API(UTF8コードページ, 0, VarPtr(入力データ(LBound(入力データ))), バイト数, StrPtr(出力文字列), 文字数)
    If 文字数 <= 0 Then Exit Function
    
    UTF8のバイト配列をVBA文字列へ変換 = Left$(出力文字列, 文字数)
End Function

' =========================================================================
' 処理概要       : COMメソッドから文字列プロパティを取得
' 引数           : ポインターインターフェース (LongPtr) - 呼び出すCOMインターフェースのポインター
'                  仮想関数位置番号 (Long) - 呼び出すCOMメソッドのvTable位置
'                  関数名称 (String) - 呼び出すCOMメソッドの識別名（省略時: ""）
' 戻り値         : String - COMメソッドから文字列プロパティ
' =========================================================================
Public Function COMメソッドから文字列プロパティを取得(ByVal ポインターインターフェース As LongPtr, ByVal 仮想関数位置番号 As Long, Optional ByVal 関数名称 As String = "") As String
    Dim ポインター文字列 As LongPtr
    Dim 実行結果 As Long
    
    If ポインターインターフェース = 0 Then Exit Function
    
    実行結果 = COMメソッド呼出(ポインターインターフェース, 仮想関数位置番号, 関数名称, VarPtr(ポインター文字列))
    If 実行結果 = 0 And ポインター文字列 <> 0 Then
        COMメソッドから文字列プロパティを取得 = UTF16文字列ポインターをVBA文字列へ変換(ポインター文字列)
        COMメモリー解放API ポインター文字列
    End If
End Function

' =========================================================================
' 処理概要       : COMメソッドから真偽値プロパティを取得
' 引数           : ポインターインターフェース (LongPtr) - 呼び出すCOMインターフェースのポインター
'                  仮想関数位置番号 (Long) - 呼び出すCOMメソッドのvTable位置
'                  関数名称 (String) - 呼び出すCOMメソッドの識別名（省略時: ""）
' 戻り値         : Boolean - 取得した真偽値。取得に失敗した場合 False
' =========================================================================
Public Function COMメソッドから真偽値プロパティを取得(ByVal ポインターインターフェース As LongPtr, ByVal 仮想関数位置番号 As Long, Optional ByVal 関数名称 As String = "") As Boolean
    Dim 入力値 As Long
    Dim 実行結果 As Long
    
    If ポインターインターフェース = 0 Then Exit Function
    
    実行結果 = COMメソッド呼出(ポインターインターフェース, 仮想関数位置番号, 関数名称, VarPtr(入力値))
    If 実行結果 = 0 Then COMメソッドから真偽値プロパティを取得 = (入力値 <> 0)
End Function

' =========================================================================
' 処理概要       : COMメソッドから32ビット整数プロパティを取得
' 引数           : ポインターインターフェース (LongPtr) - 呼び出すCOMインターフェースのポインター
'                  仮想関数位置番号 (Long) - 呼び出すCOMメソッドのvTable位置
'                  関数名称 (String) - 呼び出すCOMメソッドの識別名（省略時: ""）
' 戻り値         : Long - COMメソッドから32ビット整数プロパティ
' =========================================================================
Public Function COMメソッドから32ビット整数プロパティを取得(ByVal ポインターインターフェース As LongPtr, ByVal 仮想関数位置番号 As Long, Optional ByVal 関数名称 As String = "") As Long
    Dim 入力値 As Long
    Dim 実行結果 As Long
    
    If ポインターインターフェース = 0 Then Exit Function
    
    実行結果 = COMメソッド呼出(ポインターインターフェース, 仮想関数位置番号, 関数名称, VarPtr(入力値))
    If 実行結果 = 0 Then COMメソッドから32ビット整数プロパティを取得 = 入力値
End Function

' =========================================================================
' 処理概要       : COMメソッドから64ビット整数プロパティを取得
' 引数           : ポインターインターフェース (LongPtr) - 呼び出すCOMインターフェースのポインター
'                  仮想関数位置番号 (Long) - 呼び出すCOMメソッドのvTable位置
'                  関数名称 (String) - 呼び出すCOMメソッドの識別名（省略時: ""）
' 戻り値         : LongLong - COMメソッドから64ビット整数プロパティ
' =========================================================================
Public Function COMメソッドから64ビット整数プロパティを取得(ByVal ポインターインターフェース As LongPtr, ByVal 仮想関数位置番号 As Long, Optional ByVal 関数名称 As String = "") As LongLong
    Dim 入力値 As LongLong
    Dim 実行結果 As Long
    
    If ポインターインターフェース = 0 Then Exit Function
    
    実行結果 = COMメソッド呼出(ポインターインターフェース, 仮想関数位置番号, 関数名称, VarPtr(入力値))
    If 実行結果 = 0 Then COMメソッドから64ビット整数プロパティを取得 = 入力値
End Function

' =========================================================================
' 処理概要       : 前回の中継領域アドレスをプロセス環境変数から取得
' 戻り値         : LongPtr - 対象のアドレス。取得できない場合 0
' =========================================================================
Private Function 前回の中継領域アドレスをプロセス環境変数から取得() As LongPtr
    Dim 受信領域 As String
    Dim 件数 As Long
    Dim 作成文字列 As String
    
    前回の中継領域アドレスをプロセス環境変数から取得 = 0^
    
    受信領域 = String$(前回領域アドレス文字数, vbNullChar)
    
    件数 = 環境変数取得API(StrPtr(前回領域の保存変数名), StrPtr(受信領域), 前回領域アドレス文字数)
    
    If 件数 = 0 Then Exit Function
    
    If 件数 >= 前回領域アドレス文字数 Then Exit Function
    
    作成文字列 = Left$(受信領域, 件数)
    
    On Error Resume Next
    前回の中継領域アドレスをプロセス環境変数から取得 = CLngLng(作成文字列)
    On Error GoTo 0
End Function

' =========================================================================
' 処理概要       : 中継領域のアドレスをプロセス環境変数へ保存
' 引数           : アドレス (LongPtr) - 読み書きするメモリーアドレス
' =========================================================================
Private Sub 中継領域のアドレスをプロセス環境変数へ保存(ByVal アドレス As LongPtr)
    Dim 作成文字列 As String
    
    作成文字列 = CStr(アドレス)
    環境変数設定API StrPtr(前回領域の保存変数名), StrPtr(作成文字列)
End Sub

' =========================================================================
' 処理概要       : 保存済みの中継領域アドレスを削除
' =========================================================================
Private Sub 保存済みの中継領域アドレスを削除()
    環境変数設定API StrPtr(前回領域の保存変数名), 0
End Sub

' =========================================================================
' 処理概要       : 前回の中継領域が残っている場合に安全性を確認して解放
' =========================================================================
Private Sub 前回の中継領域が残っている場合に安全性を確認して解放()
    Dim 前回基準 As LongPtr
    Dim メモリー情報 As メモリー基本情報
    Dim 照会サイズ As LongPtr
    Dim 領域有効 As Boolean
    Dim 未使用結果 As Long
    
    前回基準 = 前回の中継領域アドレスをプロセス環境変数から取得()
    If 前回基準 = 0 Then Exit Sub
    
    領域有効 = False
    
    照会サイズ = メモリー情報取得API(前回基準, メモリー情報, LenB(メモリー情報))
    If Not (照会サイズ = 0) Then
    
        領域有効 = (メモリー情報.状態 = メモリー実体確保) And (メモリー情報.確保先頭 = 前回基準) And メモリー保護属性が読み書き可能か判定(メモリー情報.保護属性)
    End If
    
    If Not 領域有効 Then
    
        保存済みの中継領域アドレスを削除
        Exit Sub
    End If
    
    ' -------------
    ' コールバックの生存フラグを無効化して旧領域を解放
    ' -------------
    指定アドレスにポインター値を書き込み(前回基準) = 0^
    
    未使用結果 = 仮想メモリー解放API(前回基準, 0, メモリー領域解放)
    
    保存済みの中継領域アドレスを削除
End Sub

' =========================================================================
' 処理概要       : メモリー保護属性が読み書き可能か判定
' 引数           : 保護属性 (Long) - 判定するメモリー保護属性
' 戻り値         : Boolean - 読み書き可能な保護属性の場合 True
' =========================================================================
Private Function メモリー保護属性が読み書き可能か判定(ByVal 保護属性 As Long) As Boolean
    Dim 基準 As Long
    
    If (保護属性 And &H100&) <> 0 Then
        メモリー保護属性が読み書き可能か判定 = False
        Exit Function
    End If
    
    基準 = 保護属性 And Not &H100& And Not &H200& And Not &H400&
    Select Case 基準
        Case 読書可能, 実行読書可能
            メモリー保護属性が読み書き可能か判定 = True
        Case Else
            メモリー保護属性が読み書き可能か判定 = False
    End Select
End Function

' =========================================================================
' 処理概要       : COMコールバックを中継するネイティブ領域を初期化
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function COMコールバックを中継するネイティブ領域を初期化() As Boolean
    Dim ポインター中継元元 As LongPtr
    Dim 参照キー As Long
    Dim 添字 As Long, ポインタースロット As LongPtr, 仮想COMオブジェクト As LongPtr, 関数アドレス一覧 As LongPtr
    
    If 中継領域先頭 <> 0 Then
        COMコールバックを中継するネイティブ領域を初期化 = True
        Exit Function
    End If
    
    前回の中継領域が残っている場合に安全性を確認して解放
    
    ネイティブ中継入口
    
    ポインター中継元元 = VBA.Int(AddressOf ネイティブ中継入口)
    If ポインター中継元元 = 0 Then Exit Function
    
    照会関数アドレス = 関数アドレスを取得(AddressOf 要求されたCOMインターフェースに対応するポインターを返却)
    参照追加関数アドレス = 関数アドレスを取得(AddressOf 中継オブジェクトの参照カウントを加算)
    参照解放関数アドレス = 関数アドレスを取得(AddressOf 中継オブジェクトの参照カウントを減算)
    If 照会関数アドレス = 0 Or 参照追加関数アドレス = 0 Or 参照解放関数アドレス = 0 Then
        Exit Function
    End If
    
    ' 中継命令とCOMオブジェクトを格納するメモリーを確保する
    中継領域先頭 = 仮想メモリー確保API(0, 中継領域バイト数, メモリー実体確保 Or メモリー領域予約, 実行読書可能)
    If 中継領域先頭 = 0 Then Exit Function
    
    中継領域のアドレスをプロセス環境変数へ保存 中継領域先頭
    
    For 参照キー = 0 To 共通ヘッダー長 - 1 Step 8
        指定アドレスにポインター値を書き込み(中継領域先頭 + 参照キー) = 0^
    Next 参照キー
    
    指定アドレスにポインター値を書き込み(中継領域先頭) = 1^
    
    ' 各スロットに中継命令とCOMの関数アドレスを設定する
    For 添字 = 0 To スロット数 - 1
        ポインタースロット = 管理番号に対応するスロットアドレスを取得(添字)
        
        For 参照キー = 0 To 95 Step 8
            指定アドレスにポインター値を書き込み(ポインタースロット + 参照キー) = 指定アドレスからポインター値を取得(ポインター中継元元 + 参照キー)
        Next 参照キー
        
        仮想COMオブジェクト = ポインタースロット + 仮想COMオブジェクト位置
        関数アドレス一覧 = 仮想COMオブジェクト + ポインター長
        
        指定アドレスにポインター値を書き込み(仮想COMオブジェクト) = 関数アドレス一覧
        指定アドレスにポインター値を書き込み(関数アドレス一覧 + 0 * ポインター長) = 照会関数アドレス
        指定アドレスにポインター値を書き込み(関数アドレス一覧 + 1 * ポインター長) = 参照追加関数アドレス
        指定アドレスにポインター値を書き込み(関数アドレス一覧 + 2 * ポインター長) = 参照解放関数アドレス
        指定アドレスにポインター値を書き込み(関数アドレス一覧 + 3 * ポインター長) = ポインタースロット
    Next 添字
    
    ' 空きスロットを管理番号順につなぎ、割り当てに備える
    ReDim 次の未使用スロット(0 To スロット数 - 1)
    For 添字 = 0 To スロット数 - 2
        次の未使用スロット(添字) = 添字 + 1
    Next 添字
    次の未使用スロット(スロット数 - 1) = -1
    未使用スロット先頭 = 0
    使用中スロット数 = 0
    
    For 添字 = 0 To スロット数 - 1
        Set 登録ハンドラー(添字) = Nothing
    Next 添字
    
    ハンドラー種別とCOMインターフェースIDの対応表を初期化
    
    COMコールバックを中継するネイティブ領域を初期化 = True
End Function

' =========================================================================
' 処理概要       : ハンドラー用の中継スロットを割り当て
' 引数           : ハンドラー (mod_WV2_COM通知ハンドラー) - スロットに割り当てるハンドラー
'                  ポインター自身オブジェクト (LongPtr) - ハンドラーのVBAオブジェクトアドレス
'                  ポインター対象関数 (LongPtr) - 中継先メソッドのアドレス
' 戻り値         : LongPtr - 対象のアドレス。取得できない場合 0
' =========================================================================
Public Function ハンドラー用の中継スロットを割り当て(ByVal ハンドラー As mod_WV2_COM通知ハンドラー, ByVal ポインター自身オブジェクト As LongPtr, ByVal ポインター対象関数 As LongPtr) As LongPtr
    Dim 対象添字 As Long
    Dim ポインタースロット As LongPtr
    
    If 中継領域先頭 = 0 Then Exit Function
    If 未使用スロット先頭 < 0 Then Exit Function
    If ハンドラー Is Nothing Then Exit Function
    
    対象添字 = 未使用スロット先頭
    未使用スロット先頭 = 次の未使用スロット(対象添字)
    次の未使用スロット(対象添字) = -1
    
    ポインタースロット = 管理番号に対応するスロットアドレスを取得(対象添字)
    
    引数を変換する中継機械語をスロットへ書き込み ポインタースロット + 引数変換命令位置, ポインター自身オブジェクト, ポインター対象関数, 中継領域先頭
    
    指定アドレスにポインター値を書き込み(ポインタースロット + 遅延呼出差込位置) = ポインタースロット + 引数変換命令位置
    
    Set 登録ハンドラー(対象添字) = ハンドラー
    
    使用中スロット数 = 使用中スロット数 + 1
    ハンドラー用の中継スロットを割り当て = ポインタースロット
End Function

' =========================================================================
' 処理概要       : ハンドラーの中継スロットを未使用へ戻す
' 引数           : ポインタースロット (LongPtr) - 中継スロットの先頭アドレス
' =========================================================================
Public Sub ハンドラーの中継スロットを未使用へ戻す(ByVal ポインタースロット As LongPtr)
    Dim 対象添字 As Long
    
    If 中継領域先頭 = 0 Then Exit Sub
    If ポインタースロット = 0 Then Exit Sub
    
    対象添字 = スロットのアドレスから管理番号を取得(ポインタースロット)
    If 対象添字 < 0 Then Exit Sub
    If 次の未使用スロット(対象添字) <> -1 Then Exit Sub
    
    Set 登録ハンドラー(対象添字) = Nothing
    
    次の未使用スロット(対象添字) = 未使用スロット先頭
    未使用スロット先頭 = 対象添字
    
    使用中スロット数 = 使用中スロット数 - 1
End Sub

' =========================================================================
' 処理概要       : スロットのアドレスから管理番号を取得
' 引数           : ポインタースロット (LongPtr) - 中継スロットの先頭アドレス
' 戻り値         : Long - 対応する配列添字。該当しない場合 -1
' =========================================================================
Private Function スロットのアドレスから管理番号を取得(ByVal ポインタースロット As LongPtr) As Long
    Dim 位置 As LongPtr
    Dim 対象添字 As Long
    
    スロットのアドレスから管理番号を取得 = -1
    If 中継領域先頭 = 0 Then Exit Function
    
    位置 = ポインタースロット - (中継領域先頭 + 共通ヘッダー長)
    If 位置 < 0 Then Exit Function
    If (位置 Mod スロットバイト数) <> 0 Then Exit Function
    
    対象添字 = CLng(位置 \ スロットバイト数)
    If 対象添字 < 0 Or 対象添字 >= スロット数 Then Exit Function
    
    スロットのアドレスから管理番号を取得 = 対象添字
End Function

' =========================================================================
' 処理概要       : COMオブジェクトのアドレスから管理番号を取得
' 引数           : 仮想COMオブジェクト (LongPtr) - 中継COMオブジェクトのアドレス
' 戻り値         : Long - 対応する配列添字。該当しない場合 -1
' =========================================================================
Private Function COMオブジェクトのアドレスから管理番号を取得(ByVal 仮想COMオブジェクト As LongPtr) As Long
    Dim 位置 As LongPtr
    Dim 対象添字 As Long
    
    COMオブジェクトのアドレスから管理番号を取得 = -1
    If 中継領域先頭 = 0 Then Exit Function
    
    位置 = 仮想COMオブジェクト - (中継領域先頭 + 共通ヘッダー長 + 仮想COMオブジェクト位置)
    If 位置 < 0 Then Exit Function
    If (位置 Mod スロットバイト数) <> 0 Then Exit Function
    
    対象添字 = CLng(位置 \ スロットバイト数)
    If 対象添字 < 0 Or 対象添字 >= スロット数 Then Exit Function
    
    COMオブジェクトのアドレスから管理番号を取得 = 対象添字
End Function

' =========================================================================
' 処理概要       : 生存フラグを無効化してハンドラーと中継領域を解放
' =========================================================================
Public Sub 生存フラグを無効化してハンドラーと中継領域を解放()
    Dim 添字 As Long
    Dim 終了未使用結果 As Long
    
    If 中継領域先頭 = 0 Then Exit Sub
    
    ' -------------
    ' コールバックの生存フラグを無効化して中継領域を解放
    ' -------------
    指定アドレスにポインター値を書き込み(中継領域先頭) = 0^
    
    ' 各スロットに中継命令とCOMの関数アドレスを設定する
    For 添字 = 0 To スロット数 - 1
        If Not (登録ハンドラー(添字) Is Nothing) Then
            登録ハンドラー(添字).通知先オブジェクトへの参照を解除
            Set 登録ハンドラー(添字) = Nothing
        End If
    Next 添字
    
    終了未使用結果 = 仮想メモリー解放API(中継領域先頭, 0, メモリー領域解放)
    
    中継領域先頭 = 0
    未使用スロット先頭 = -1
    使用中スロット数 = 0
    Erase 次の未使用スロット
    
    保存済みの中継領域アドレスを削除
End Sub

' =========================================================================
' 処理概要       : 管理番号に対応するスロットアドレスを取得
' 引数           : 対象添字 (Long) - 対象の配列添字
' 戻り値         : LongPtr - 対象のアドレス。取得できない場合 0
' =========================================================================
Private Function 管理番号に対応するスロットアドレスを取得(ByVal 対象添字 As Long) As LongPtr
    管理番号に対応するスロットアドレスを取得 = 中継領域先頭 + 共通ヘッダー長 + CLngLng(対象添字) * スロットバイト数
End Function

' =========================================================================
' 処理概要       : 引数を変換する中継機械語をスロットへ書き込み
' 引数           : アドレス (LongPtr) - 読み書きするメモリーアドレス
'                  ポインター自身オブジェクト (LongPtr) - ハンドラーのVBAオブジェクトアドレス
'                  ポインター対象関数 (LongPtr) - 中継先メソッドのアドレス
'                  ポインター領域基準 (LongPtr) - ネイティブ中継領域の先頭アドレス
' =========================================================================
Private Sub 引数を変換する中継機械語をスロットへ書き込み(ByVal アドレス As LongPtr, ByVal ポインター自身オブジェクト As LongPtr, ByVal ポインター対象関数 As LongPtr, ByVal ポインター領域基準 As LongPtr)
    Dim 入力データ(0 To 命令作成領域長 - 1) As Byte, 添字 As Long
    Dim 参照キー As Long
    
    添字 = 0
    
    入力データ(添字) = &H48: 入力データ(添字 + 1) = &HB8: 添字 = 添字 + 2
    指定アドレスにポインター値を書き込み(VarPtr(入力データ(添字))) = ポインター領域基準: 添字 = 添字 + 8
    
    入力データ(添字) = &H80: 入力データ(添字 + 1) = &H38: 入力データ(添字 + 2) = &H1: 添字 = 添字 + 3
    
    入力データ(添字) = &H74: 入力データ(添字 + 1) = &H3: 添字 = 添字 + 2
    
    入力データ(添字) = &H33: 入力データ(添字 + 1) = &HC0: 添字 = 添字 + 2
    
    入力データ(添字) = &HC3: 添字 = 添字 + 1
    
    入力データ(添字) = &H48: 入力データ(添字 + 1) = &H83: 入力データ(添字 + 2) = &HEC: 入力データ(添字 + 3) = &H38: 添字 = 添字 + 4
    
    入力データ(添字) = &H4D: 入力データ(添字 + 1) = &H89: 入力データ(添字 + 2) = &HC1: 添字 = 添字 + 3
    
    入力データ(添字) = &H49: 入力データ(添字 + 1) = &H89: 入力データ(添字 + 2) = &HD0: 添字 = 添字 + 3
    
    入力データ(添字) = &H48: 入力データ(添字 + 1) = &H89: 入力データ(添字 + 2) = &HCA: 添字 = 添字 + 3
    
    入力データ(添字) = &H48: 入力データ(添字 + 1) = &HB9: 添字 = 添字 + 2
    指定アドレスにポインター値を書き込み(VarPtr(入力データ(添字))) = ポインター自身オブジェクト: 添字 = 添字 + 8
    
    入力データ(添字) = &H48: 入力データ(添字 + 1) = &H8D: 入力データ(添字 + 2) = &H44: 入力データ(添字 + 3) = &H24: 入力データ(添字 + 4) = &H28: 添字 = 添字 + 5
    
    入力データ(添字) = &H48: 入力データ(添字 + 1) = &H89: 入力データ(添字 + 2) = &H44: 入力データ(添字 + 3) = &H24: 入力データ(添字 + 4) = &H20: 添字 = 添字 + 5
    
    入力データ(添字) = &H48: 入力データ(添字 + 1) = &HB8: 添字 = 添字 + 2
    指定アドレスにポインター値を書き込み(VarPtr(入力データ(添字))) = ポインター対象関数: 添字 = 添字 + 8
    
    入力データ(添字) = &HFF: 入力データ(添字 + 1) = &HD0: 添字 = 添字 + 2
    
    入力データ(添字) = &H8B: 入力データ(添字 + 1) = &H44: 入力データ(添字 + 2) = &H24: 入力データ(添字 + 3) = &H28: 添字 = 添字 + 4
    
    入力データ(添字) = &H48: 入力データ(添字 + 1) = &H83: 入力データ(添字 + 2) = &HC4: 入力データ(添字 + 3) = &H38: 添字 = 添字 + 4
    
    入力データ(添字) = &HC3: 添字 = 添字 + 1
    
    入力データ(添字) = &HCC: 添字 = 添字 + 1
    入力データ(添字) = &HCC: 添字 = 添字 + 1
    
    For 参照キー = 0 To 命令作成領域長 - 1 Step 8
        指定アドレスにポインター値を書き込み(アドレス + 参照キー) = バイト配列からポインター値を取得(入力データ, 参照キー)
    Next 参照キー
End Sub

' =========================================================================
' 処理概要       : バイト配列からポインター値を取得
' 引数           : 入力データ (Byte()) - 変換対象のバイト配列
'                  位置 (Long) - バイト配列の読み取り開始位置
' 戻り値         : LongPtr - 対象のアドレス。取得できない場合 0
' =========================================================================
Private Function バイト配列からポインター値を取得(ByRef 入力データ() As Byte, ByVal 位置 As Long) As LongPtr
    バイト配列からポインター値を取得 = 指定アドレスからポインター値を取得(VarPtr(入力データ(位置)))
End Function

' =========================================================================
' 処理概要       : クラスの固定vTable位置からメソッドアドレスを取得
' 引数           : 対象インスタンス (Object) - メソッドアドレスを取得するクラスのインスタンス
'                  スロット位置番号 (Long) - ハンドラースロットまたはvTableの位置
' 戻り値         : LongPtr - 対象のアドレス。取得できない場合 0
' =========================================================================
Private Function クラスの固定vTable位置からメソッドアドレスを取得(ByVal 対象インスタンス As Object, ByVal スロット位置番号 As Long) As LongPtr
    Dim ポインターオブジェクト As LongPtr
    Dim ポインター値対応表 As LongPtr
    
    If 対象インスタンス Is Nothing Then Exit Function
    
    ポインターオブジェクト = ObjPtr(対象インスタンス)
    If ポインターオブジェクト = 0 Then Exit Function
    
    ポインター値対応表 = 指定アドレスからポインター値を取得(ポインターオブジェクト)
    If ポインター値対応表 = 0 Then Exit Function
    
    クラスの固定vTable位置からメソッドアドレスを取得 = 指定アドレスからポインター値を取得(ポインター値対応表 + スロット位置番号 * ポインター長)
End Function

' =========================================================================
' 処理概要       : AddressOfで取得したアドレスを返却
' 引数           : アドレス (LongPtr) - 読み書きするメモリーアドレス
' 戻り値         : LongPtr - 対象のアドレス。取得できない場合 0
' =========================================================================
Private Function 関数アドレスを取得(ByVal アドレス As LongPtr) As LongPtr
    関数アドレスを取得 = アドレス
End Function

' =========================================================================
' 処理概要       : 要求されたCOMインターフェースに対応するポインターを返却
' 引数           : 呼出元アドレス (LongPtr) - 呼び出された中継COMオブジェクトのアドレス
'                  インターフェース識別子 (LongPtr) - 要求されたインターフェースIDのポインター
'                  インターフェース出力先 (LongPtr) - インターフェースポインターの出力先
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Private Function 要求されたCOMインターフェースに対応するポインターを返却(ByVal 呼出元アドレス As LongPtr, ByVal インターフェース識別子 As LongPtr, ByRef インターフェース出力先 As LongPtr) As Long
    Dim 対象添字 As Long
    Dim 通知種別 As 通知区分
    
    If インターフェース識別子 = 0 Then
        インターフェース出力先 = 0
        要求されたCOMインターフェースに対応するポインターを返却 = &H80004003
        Exit Function
    End If
    
    対象添字 = COMオブジェクトのアドレスから管理番号を取得(呼出元アドレス)
    If 対象添字 < 0 Then
        インターフェース出力先 = 0
        要求されたCOMインターフェースに対応するポインターを返却 = 対応インターフェースなし
        Exit Function
    End If
    If 登録ハンドラー(対象添字) Is Nothing Then
        インターフェース出力先 = 0
        要求されたCOMインターフェースに対応するポインターを返却 = 対応インターフェースなし
        Exit Function
    End If
    
    If 指定アドレスのGUIDと比較対象のGUIDが一致するか判定(インターフェース識別子, IUnknown識別子) Then
        インターフェース出力先 = 呼出元アドレス
        スロット内のハンドラー参照カウントを加算 呼出元アドレス
        要求されたCOMインターフェースに対応するポインターを返却 = 正常終了
        Exit Function
    End If
    
    通知種別 = 登録ハンドラー(対象添字).通知種別
    If 指定アドレスのGUIDと比較対象のGUIDが一致するか判定(インターフェース識別子, 通知識別子一覧(通知種別)) Then
        インターフェース出力先 = 呼出元アドレス
        スロット内のハンドラー参照カウントを加算 呼出元アドレス
        要求されたCOMインターフェースに対応するポインターを返却 = 正常終了
        Exit Function
    End If
    
    インターフェース出力先 = 0
    要求されたCOMインターフェースに対応するポインターを返却 = 対応インターフェースなし
End Function

' =========================================================================
' 処理概要       : 中継オブジェクトの参照カウントを加算
' 引数           : 呼出元アドレス (LongPtr) - 呼び出された中継COMオブジェクトのアドレス
' 戻り値         : Long - 更新後のCOM参照カウント
' =========================================================================
Private Function 中継オブジェクトの参照カウントを加算(ByVal 呼出元アドレス As LongPtr) As Long
    中継オブジェクトの参照カウントを加算 = スロット内のハンドラー参照カウントを加算(呼出元アドレス)
End Function

' =========================================================================
' 処理概要       : 中継オブジェクトの参照カウントを減算
' 引数           : 呼出元アドレス (LongPtr) - 呼び出された中継COMオブジェクトのアドレス
' 戻り値         : Long - 更新後のCOM参照カウント
' =========================================================================
Private Function 中継オブジェクトの参照カウントを減算(ByVal 呼出元アドレス As LongPtr) As Long
    中継オブジェクトの参照カウントを減算 = 参照カウントを減算し未参照のスロットを解放(呼出元アドレス)
End Function

' =========================================================================
' 処理概要       : 指定アドレスのGUIDと比較対象のGUIDが一致するか判定
' 引数           : ポインターインターフェース識別子 (LongPtr) - 比較対象のGUIDへのポインター
'                  参照GUID (GUID情報) - 比較するGUID構造体
' 戻り値         : Boolean - GUIDが一致する場合 True
' =========================================================================
Private Function 指定アドレスのGUIDと比較対象のGUIDが一致するか判定(ByVal ポインターインターフェース識別子 As LongPtr, ByRef 参照GUID As GUID情報) As Boolean
    Dim ポインター参照 As LongPtr
    
    ポインター参照 = VarPtr(参照GUID)
    
    If 指定アドレスからポインター値を取得(ポインターインターフェース識別子) <> 指定アドレスからポインター値を取得(ポインター参照) Then Exit Function
    If 指定アドレスからポインター値を取得(ポインターインターフェース識別子 + 8) <> 指定アドレスからポインター値を取得(ポインター参照 + 8) Then Exit Function
    
    指定アドレスのGUIDと比較対象のGUIDが一致するか判定 = True
End Function

' =========================================================================
' 処理概要       : ポインター値の下位32ビットを符号付き整数として取得
' 引数           : 処理値 (LongPtr) - 設定する値
' 戻り値         : Long - 下位32ビットの符号付き整数値
' =========================================================================
Private Function ポインター値の下位32ビットを符号付き整数として取得(ByVal 処理値 As LongPtr) As Long
    Dim 対象URL As LongLong
    
    対象URL = CLngLng(処理値) And &HFFFFFFFF^
    If 対象URL > &H7FFFFFFF^ Then
        ポインター値の下位32ビットを符号付き整数として取得 = CLng(対象URL - &H100000000^)
    Else
        ポインター値の下位32ビットを符号付き整数として取得 = CLng(対象URL)
    End If
End Function

' =========================================================================
' 処理概要       : ハンドラー種別とCOMインターフェースIDの対応表を初期化
' =========================================================================
Private Sub ハンドラー種別とCOMインターフェースIDの対応表を初期化()
    文字列表記のGUIDをGUID構造体へ設定 IUnknown識別子, "00000000-0000-0000-C000-000000000046"
    
    文字列表記のGUIDをGUID構造体へ設定 通知識別子一覧(環境生成完了通知), "4e8a3389-c9d8-4bd2-b6b5-124fee6cc14d"
    
    文字列表記のGUIDをGUID構造体へ設定 通知識別子一覧(表示生成完了通知), "6c4819f3-c9b7-4260-8127-c9f5bde7f68c"
    
    文字列表記のGUIDをGUID構造体へ設定 通知識別子一覧(遷移開始通知), "9adbe429-f36d-432b-9ddc-f8881fbd76e3"
    
    文字列表記のGUIDをGUID構造体へ設定 通知識別子一覧(遷移完了通知), "d33a35bf-1c49-4f98-93ab-006e0533fe1c"
    
    文字列表記のGUIDをGUID構造体へ設定 通知識別子一覧(Webメッセージ通知), "57213f19-00e6-49fa-8e07-898ea01ecbd2"
    
    文字列表記のGUIDをGUID構造体へ設定 通知識別子一覧(タイトル変更通知), "f5f2b923-953e-4042-9f95-f3a118e1afd4"
    
    文字列表記のGUIDをGUID構造体へ設定 通知識別子一覧(新規ウィンドウ通知), "d4c185fe-c81c-4989-97af-2d3fa7ab5651"
    
    文字列表記のGUIDをGUID構造体へ設定 通知識別子一覧(スクリプト完了通知), "49511172-cc67-4bca-9923-137112f4c4cc"
    
    文字列表記のGUIDをGUID構造体へ設定 通知識別子一覧(履歴変更通知), "c79a420c-efd9-4058-9295-3e8b4bcab645"
    
    文字列表記のGUIDをGUID構造体へ設定 通知識別子一覧(DOM読込通知), "4bac7e9c-199e-49ed-87ed-249303acf019"
    
    文字列表記のGUIDをGUID構造体へ設定 通知識別子一覧(通信要求通知), "ab00b74c-15f1-4646-80e8-e76341d25d71"
    
    文字列表記のGUIDをGUID構造体へ設定 通知識別子一覧(通信応答通知), "7de9898a-24f5-40c3-a2de-d4f458e69828"
    
    文字列表記のGUIDをGUID構造体へ設定 通知識別子一覧(応答本文取得通知), "875738e1-9fa2-40e3-8b74-2e8972dd6fe7"
End Sub

' =========================================================================
' 処理概要       : 文字列表記のGUIDをGUID構造体へ設定
' 引数           : GUID値 (GUID情報) - 設定するGUID構造体
'                  作成文字列 (String) - 変換対象の文字列
' =========================================================================
Public Sub 文字列表記のGUIDをGUID構造体へ設定(ByRef GUID値 As GUID情報, ByVal 作成文字列 As String)
    GUID値.先頭値 = 数値16進文字列をLongに変換(Mid$(作成文字列, 1, 8))
    GUID値.第2値 = 数値16進文字列をIntegerに変換(Mid$(作成文字列, 10, 4))
    GUID値.第3値 = 数値16進文字列をIntegerに変換(Mid$(作成文字列, 15, 4))
    
    GUID値.後半値(0) = CByte("&H" & Mid$(作成文字列, 20, 2))
    GUID値.後半値(1) = CByte("&H" & Mid$(作成文字列, 22, 2))
    GUID値.後半値(2) = CByte("&H" & Mid$(作成文字列, 25, 2))
    GUID値.後半値(3) = CByte("&H" & Mid$(作成文字列, 27, 2))
    GUID値.後半値(4) = CByte("&H" & Mid$(作成文字列, 29, 2))
    GUID値.後半値(5) = CByte("&H" & Mid$(作成文字列, 31, 2))
    GUID値.後半値(6) = CByte("&H" & Mid$(作成文字列, 33, 2))
    GUID値.後半値(7) = CByte("&H" & Mid$(作成文字列, 35, 2))
End Sub

' =========================================================================
' 処理概要       : 8桁の16進文字列を32ビット整数へ変換
' 引数           : 作成文字列 (String) - 変換対象の文字列
' 戻り値         : Long - 8桁の16進文字列を32ビット整数
' =========================================================================
Private Function 数値16進文字列をLongに変換(ByVal 作成文字列 As String) As Long
    Dim 処理値 As LongLong
    
    処理値 = CLngLng("&H" & 作成文字列)
    If 処理値 >= &H80000000^ Then
        数値16進文字列をLongに変換 = CLng(処理値 - &H100000000^)
    Else
        数値16進文字列をLongに変換 = CLng(処理値)
    End If
End Function

' =========================================================================
' 処理概要       : 4桁の16進文字列を16ビット整数へ変換
' 引数           : 作成文字列 (String) - 変換対象の文字列
' 戻り値         : Integer - 4桁の16進文字列を16ビット整数
' =========================================================================
Private Function 数値16進文字列をIntegerに変換(ByVal 作成文字列 As String) As Integer
    Dim 処理値 As Long
    
    処理値 = CLng("&H" & 作成文字列)
    If 処理値 >= &H8000& Then
        数値16進文字列をIntegerに変換 = CInt(処理値 - &H10000)
    Else
        数値16進文字列をIntegerに変換 = CInt(処理値)
    End If
End Function

' =========================================================================
' 処理概要       : スロット内のハンドラー参照カウントを加算
' 引数           : 呼出元アドレス (LongPtr) - 呼び出された中継COMオブジェクトのアドレス
' 戻り値         : Long - 更新後のCOM参照カウント
' =========================================================================
Private Function スロット内のハンドラー参照カウントを加算(ByVal 呼出元アドレス As LongPtr) As Long
    Dim 対象添字 As Long
    Dim 件数 As Long
    
    対象添字 = COMオブジェクトのアドレスから管理番号を取得(呼出元アドレス)
    If 対象添字 < 0 Then Exit Function
    If 登録ハンドラー(対象添字) Is Nothing Then Exit Function
    
    件数 = 登録ハンドラー(対象添字).参照数 + 1
    登録ハンドラー(対象添字).参照数 = 件数
    
    スロット内のハンドラー参照カウントを加算 = 件数
End Function

' =========================================================================
' 処理概要       : 参照カウントを減算し未参照のスロットを解放
' 引数           : 呼出元アドレス (LongPtr) - 呼び出された中継COMオブジェクトのアドレス
' 戻り値         : Long - 更新後のCOM参照カウント
' =========================================================================
Private Function 参照カウントを減算し未参照のスロットを解放(ByVal 呼出元アドレス As LongPtr) As Long
    Dim 対象添字 As Long
    Dim 件数 As Long
    Dim 通知ハンドラー As mod_WV2_COM通知ハンドラー
    Dim ポインタースロット As LongPtr
    
    対象添字 = COMオブジェクトのアドレスから管理番号を取得(呼出元アドレス)
    If 対象添字 < 0 Then Exit Function
    If 登録ハンドラー(対象添字) Is Nothing Then Exit Function
    
    件数 = 登録ハンドラー(対象添字).参照数 - 1
    If 件数 < 0 Then 件数 = 0
    登録ハンドラー(対象添字).参照数 = 件数
    
    If 件数 = 0 Then
    
        Set 通知ハンドラー = 登録ハンドラー(対象添字)
        
        ポインタースロット = 通知ハンドラー.スロットアドレス
        
        通知ハンドラー.通知先オブジェクトへの参照を解除
        
        Set 登録ハンドラー(対象添字) = Nothing
        
        ハンドラーの中継スロットを未使用へ戻す ポインタースロット
        
        Set 通知ハンドラー = Nothing
        
    End If
    
    参照カウントを減算し未参照のスロットを解放 = 件数
End Function

' =========================================================================
' 処理概要       : 指定アドレスにポインター値を書き込み
' 引数           : アドレス (LongPtr) - 読み書きするメモリーアドレス
'                  新規値 (LongPtr) - 設定する値
' =========================================================================
Private Property Let 指定アドレスにポインター値を書き込み(ByVal アドレス As LongPtr, ByVal 新規値 As LongPtr)
    Dim ポインター参照(0 To 0) As ポインター参照情報
    
    With ポインター参照(0)
        .配列情報.次元数 = 1
        .配列情報.ロック数 = 1
        .配列情報.属性値 = 配列属性自動 Or 配列属性固定長
        .配列情報.要素バイト数 = ポインター長
        .配列情報.データアドレス = アドレス
        .配列情報.先頭次元境界.要素数 = 1
        ポインター配列の先頭にアドレスを書き込み ポインター参照, VarPtr(.配列情報)
        .参照配列(0) = 新規値
        .配列情報.先頭次元境界.要素数 = 0
        .配列情報.データアドレス = 空ポインター
    End With
End Property

' =========================================================================
' 処理概要       : 指定アドレスからポインター値を取得
' 引数           : アドレス (LongPtr) - 読み書きするメモリーアドレス
' 戻り値         : LongPtr - 対象のアドレス。取得できない場合 0
' =========================================================================
Private Function 指定アドレスからポインター値を取得(ByVal アドレス As LongPtr) As LongPtr
    Dim ポインター参照(0 To 0) As ポインター参照情報
    
    With ポインター参照(0)
        .配列情報.次元数 = 1
        .配列情報.ロック数 = 1
        .配列情報.属性値 = 配列属性自動 Or 配列属性固定長
        .配列情報.要素バイト数 = ポインター長
        .配列情報.データアドレス = アドレス
        .配列情報.先頭次元境界.要素数 = 1
        ポインター配列の先頭にアドレスを書き込み ポインター参照, VarPtr(.配列情報)
        指定アドレスからポインター値を取得 = .参照配列(0)
        .配列情報.先頭次元境界.要素数 = 0
        .配列情報.データアドレス = 空ポインター
    End With
End Function

' =========================================================================
' 処理概要       : ポインター配列の先頭にアドレスを書き込み
' 引数           : ポインター配列 (LONG_PTR()) - 書き込み対象のポインター配列
'                  書込ポインター (LongPtr) - 書き込むポインター値
' =========================================================================
Private Sub ポインター配列の先頭にアドレスを書き込み(ByRef ポインター配列() As LONG_PTR, ByVal 書込ポインター As LongPtr)
    ポインター配列(0) = 書込ポインター
End Sub
