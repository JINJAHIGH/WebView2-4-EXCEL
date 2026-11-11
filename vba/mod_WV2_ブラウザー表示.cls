VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "mod_WV2_ブラウザー表示"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' -------------
' 定数
' -------------
Private Const 新規タブ作成方式       As Long = 0
Private Const ポップアップ委譲方式   As Long = 1
Private Const 予備タブ利用方式       As Long = 2
Private Const スクリプト結果取得済み As Long = 1
Private Const スクリプト結果破棄済み As Long = 2
Private Const 状態確認間隔秒         As Single = 0.1
Private Const 画面監視版             As Long = 3
Private Const 要素数警告閾値         As Long = 5000
Private Const 通信本文読込単位       As Long = 16384
Private Const 通信16進先頭           As Long = 32
Private Const ストリーム先頭位置     As Long = 0
Private Const データなし             As Long = &H800700E8

' -------------
' 待機時間
' -------------
Private Const スクリプト結果待機中 As Long = 0
Private Const 既定スクリプト待機秒 As Single = 5
Private Const 既定要素待機秒       As Single = 10
Private Const 既定安定待機ミリ秒   As Long = 500
Private Const 既定長時間通信ミリ秒 As Long = 10000

' -------------
' 記録と取得の上限
' -------------
Private Const 既定要素取得上限     As Long = 200
Private Const 通信詳細保持上限既定 As Long = 50
Private Const 通信詳細保持上限最大 As Long = 2000
Private Const 通信本文最大既定     As Long = 65536
Private Const 通信本文最大上限     As Long = 8388608
Private Const 通信ヘッダー最大     As Long = 40
Private Const 通信応答本文最大既定 As Long = 65536
Private Const 通信保持上限既定     As Long = 500
Private Const 通信保持上限最大     As Long = 50000
Private Const 通信リソース種別既定 As String = "DOCUMENT,XHR,FETCH"

Public Event スクリプト実行完了(ByVal 通知識別番号 As Long, ByVal エラーコード As Long, ByVal 結果JSON As String)

Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

Public Enum 表示初期化状態
    表示作成済み = 0
    表示準備待機中 = 1
    表示準備済み = 2
    表示初期化失敗 = 3
End Enum

Private 保持ハンドラー As mod_WV2_COM通知ハンドラー
Private 保持ポインターコントローラー   As LongPtr
Private 保持ポインターWebView   As LongPtr

Private 保持ポインター設定 As LongPtr

Private 保持ポインターWebView2 As LongPtr

Private 保持ポインターWebView3 As LongPtr

Private 保持ブラウザー As mod_WV2_ブラウザー管理
Private 保持状態   As 表示初期化状態
Private 保持最終エラー As Long

Private 保持永続ハンドラー一覧(遷移開始通知 To 通信応答通知) As mod_WV2_COM通知ハンドラー
Private 保持永続登録トークン一覧(遷移開始通知 To 通信応答通知) As LongLong

Private 保持ページ読込済み As Boolean

Private 保持実行スクリプト連番 As Long
Private 保持終了済み As Boolean

Private 保持タイトル As String
Private 保持Webメッセージ通知先 As Object
Private 保持判定設定 As Boolean

Private 保持スクリプト評価待機中   As Collection
Private 保持最終スクリプト評価成功    As Boolean
Private 保持最終スクリプト評価エラー As String
Private 保持内部通知    As Long
Private 保持スクリプト評価階層数     As Long

Private 保持暗黙待機秒 As Single

Private 保持全件打切り済み As Boolean

Private 保持除外通信  As Collection
Private 保持除外選択条件  As String
Private 保持自動待機   As Boolean
Private 保持最終安定待機 As String

Private 保持シグナルDOM As Boolean
Private 保持シグナル通信 As Boolean
Private 保持シグナルURL As Boolean

Private 保持参照失効ミリ秒 As Long

Private 保持詳細有効         As Boolean
Private 保持詳細保持上限        As Long
Private 保持詳細先頭       As Long
Private 保持詳細件数      As Long
Private 保持詳細破棄件数    As Long
Private 保持詳細本文最大    As Long
Private 保持詳細機密値除外     As Boolean
Private 保持詳細機密値除外一覧 As String
Private 保持詳細URIフィルター  As String
Private 保持詳細連番()     As Long
Private 保持詳細メソッド()  As String
Private 保持詳細URI()     As String
Private 保持詳細ヘッダー一覧() As String
Private 保持詳細件数型()   As String
Private 保持詳細バイト数()   As Long
Private 保持詳細打切り()   As Boolean
Private 保持詳細判定文字列()  As Boolean
Private 保持詳細本文()    As String
Private 保持詳細16進()     As String
Private 保持詳細応答ヘッダー一覧() As String
Private 保持詳細状態()      As Long
Private 保持詳細理由()      As String
Private 保持詳細応答本文()    As String
Private 保持詳細応答バイト数()   As Long
Private 保持詳細応答打切り()   As Boolean
Private 保持詳細応答判定文字列()  As Boolean
Private 保持詳細応答文字コード()     As String
Private 保持詳細応答16進()     As String
Private 保持詳細応答取得済み()     As Boolean

Private 保持通信有効      As Boolean
Private 保持通信保持上限     As Long
Private 保持通信先頭    As Long
Private 保持通信件数   As Long
Private 保持通信総数   As Long
Private 保持通信破棄件数 As Long
Private 保持通信計測開始0      As Long
Private 保持通信応答有効    As Boolean
Private 保持通信未対応件数 As Long

Private 保持応答本文有効     As Boolean
Private 保持応答本文最大    As Long
Private 保持応答待機中    As Collection
Private 保持応答本文対応なし件数 As Long

Private 保持通信フィルター一覧 As Collection
Private 保持通信状態()  As Long
Private 保持通信理由()  As String
Private 保持通信連番()    As Long
Private 保持通信ミリ秒()     As Long
Private 保持通信リソース種別()    As Long
Private 保持通信メソッド() As String
Private 保持通信URI()    As String

Private 保持DOM世代 As String

' =========================================================================
' 処理概要       : インスタンスの内部状態を初期化
' =========================================================================
Private Sub Class_Initialize()
    保持参照失効ミリ秒 = 既定長時間通信ミリ秒
    保持状態 = 表示作成済み
    保持ポインターコントローラー = 0
    保持ポインターWebView = 0
    保持ポインターWebView2 = 0
    保持ポインターWebView3 = 0
    Set 保持ブラウザー = Nothing
    保持最終エラー = 0
    保持ページ読込済み = False
    保持実行スクリプト連番 = 0
    保持終了済み = False
    保持タイトル = ""
    Set 保持Webメッセージ通知先 = Nothing
    保持判定設定 = False
    
    Set 保持スクリプト評価待機中 = New Collection
    保持最終スクリプト評価成功 = False
    保持最終スクリプト評価エラー = ""
    保持内部通知 = 0
    保持スクリプト評価階層数 = 0
    
    保持DOM世代 = ""
    
    保持通信有効 = False
    保持通信保持上限 = 通信保持上限既定
    Set 保持通信フィルター一覧 = New Collection
    通信一覧のリングバッファーを確保
    
    保持詳細有効 = False
    保持詳細保持上限 = 通信詳細保持上限既定
    保持詳細本文最大 = 通信本文最大既定
    保持詳細機密値除外 = True
    保持詳細機密値除外一覧 = 伏せ字にする既定のヘッダー名一覧を取得()
    保持詳細URIフィルター = ""
    通信詳細のリングバッファーを確保
    
    保持応答本文有効 = False
    保持応答本文最大 = 通信応答本文最大既定
    Set 保持応答待機中 = New Collection
    保持応答本文対応なし件数 = 0
End Sub

' =========================================================================
' 処理概要       : 保持するリソースと参照を解放
' =========================================================================
Private Sub Class_Terminate()
    登録済みイベントを解除してハンドラーを解放
    
    If 保持ポインター設定 <> 0 Then
    
        COM参照解放 保持ポインター設定
        保持ポインター設定 = 0
    End If
    
    If 保持ポインターWebView3 <> 0 Then
    
        COM参照解放 保持ポインターWebView3
        保持ポインターWebView3 = 0
    End If
    
    If 保持ポインターWebView2 <> 0 Then
    
        COM参照解放 保持ポインターWebView2
        保持ポインターWebView2 = 0
    End If
    
    If 保持ポインターWebView <> 0 Then
    
        COM参照解放 保持ポインターWebView
        保持ポインターWebView = 0
    End If
    
    If 保持ポインターコントローラー <> 0 Then
    
        COM参照解放 保持ポインターコントローラー
        保持ポインターコントローラー = 0
    End If
    
    If Not (保持ハンドラー Is Nothing) Then
    
        Set 保持ハンドラー = Nothing
    End If
    
End Sub

' =========================================================================
' 処理概要       : 登録済みイベントを解除してハンドラーを解放
' =========================================================================
Private Sub 登録済みイベントを解除してハンドラーを解放()
    Dim 参照キー先頭 As Long
    Dim 参照キー As Long
    Dim 解除位置番号 As Long
    Dim イベント解除インターフェース As LongPtr
    
    If 保持ポインターWebView = 0 Then
    
        For 参照キー先頭 = LBound(保持永続ハンドラー一覧) To UBound(保持永続ハンドラー一覧)
            Set 保持永続ハンドラー一覧(参照キー先頭) = Nothing
            保持永続登録トークン一覧(参照キー先頭) = 0^
        Next 参照キー先頭
        Exit Sub
    End If
    
    For 参照キー = LBound(保持永続ハンドラー一覧) To UBound(保持永続ハンドラー一覧)
        If Not (保持永続ハンドラー一覧(参照キー) Is Nothing) Then
            解除位置番号 = イベント解除メソッドのvTable位置を取得(CLng(参照キー))
            If 解除位置番号 > 0 Then
            
                イベント解除インターフェース = イベント解除に使用するインターフェースを取得(CLng(参照キー))
                If イベント解除インターフェース <> 0 Then
                
                    COMメソッド呼出 イベント解除インターフェース, 解除位置番号, "View.remove_Event", 保持永続登録トークン一覧(参照キー)
                End If
            End If
            Set 保持永続ハンドラー一覧(参照キー) = Nothing
            保持永続登録トークン一覧(参照キー) = 0^
        End If
    Next 参照キー
End Sub

' =========================================================================
' 処理概要       : イベント解除に使用するインターフェースを取得
' 引数           : 通知種別 (Long) - コールバックまたはイベントの種別
' 戻り値         : LongPtr - 対象のアドレス。取得できない場合 0
' =========================================================================
Private Function イベント解除に使用するインターフェースを取得(ByVal 通知種別 As Long) As LongPtr
    Select Case 通知種別
        Case DOM読込通知, 通信応答通知
        
            イベント解除に使用するインターフェースを取得 = 保持ポインターWebView2
        Case Else
        
            イベント解除に使用するインターフェースを取得 = 保持ポインターWebView
    End Select
End Function

' =========================================================================
' 処理概要       : イベント解除メソッドのvTable位置を取得
' 引数           : 通知種別 (Long) - コールバックまたはイベントの種別
' 戻り値         : Long - イベント解除メソッドのvTable位置
' =========================================================================
Private Function イベント解除メソッドのvTable位置を取得(ByVal 通知種別 As Long) As Long
    Select Case 通知種別
        Case 遷移開始通知:   イベント解除メソッドのvTable位置を取得 = 8
        Case 遷移完了通知:  イベント解除メソッドのvTable位置を取得 = 16
        Case Webメッセージ通知:   イベント解除メソッドのvTable位置を取得 = 35
        Case タイトル変更通知: イベント解除メソッドのvTable位置を取得 = 47
        Case 新規ウィンドウ通知:   イベント解除メソッドのvTable位置を取得 = 45
        Case 履歴変更通知:       イベント解除メソッドのvTable位置を取得 = 14
        Case DOM読込通知:     イベント解除メソッドのvTable位置を取得 = 65
        Case 通信要求通知: イベント解除メソッドのvTable位置を取得 = 56
        Case 通信応答通知: イベント解除メソッドのvTable位置を取得 = 62
        Case Else:                    イベント解除メソッドのvTable位置を取得 = 0
    End Select
End Function

' =========================================================================
' 処理概要       : 指定ウィンドウにコントローラーの非同期生成を開始
' 引数           : 環境 (mod_WV2_ブラウザー環境) - コントローラーを生成するWebView2環境
'                  ウィンドウハンドル (LongPtr) - WebView2を配置するウィンドウハンドル
' =========================================================================
Public Sub 初期化(ByVal 環境 As mod_WV2_ブラウザー環境, ByVal ウィンドウハンドル As LongPtr)
    Dim ポインター環境 As LongPtr
    Dim 仮想COMオブジェクト As LongPtr
    Dim 実行結果 As Long
    
    If 保持状態 <> 表示作成済み Then
        MsgBox "ブラウザー表示.Init は二重呼び出し不可", vbCritical
        Exit Sub
    End If
    
    If 環境 Is Nothing Then
        MsgBox "ブラウザー表示.Init: env が Nothing", vbCritical
        保持状態 = 表示初期化失敗
        Exit Sub
    End If
    
    If Not 環境.利用可能 Then
        MsgBox "ブラウザー表示.Init: env がまだ Ready でない (state=" & 環境.状態 & ")", vbCritical
        保持状態 = 表示初期化失敗
        Exit Sub
    End If
    
    ポインター環境 = 環境.環境ポインター
    If ポインター環境 = 0 Then
        MsgBox "ブラウザー表示.Init: env.EnvironmentPtr が 0", vbCritical
        保持状態 = 表示初期化失敗
        Exit Sub
    End If
    
    Set 保持ハンドラー = 通知先を設定したCOMコールバックハンドラーを取得(表示生成完了通知, Me)
    If 保持ハンドラー Is Nothing Then
        MsgBox "ブラウザー表示.Init: ハンドラの Acquire に失敗", vbCritical
        保持状態 = 表示初期化失敗
        Exit Sub
    End If
    
    仮想COMオブジェクト = 保持ハンドラー.スロットアドレス + 仮想COMオブジェクト位置
    
    保持状態 = 表示準備待機中
    
    実行結果 = COMメソッド呼出(ポインター環境, 3, "Env.CreateCoreWebView2Controller", ウィンドウハンドル, 仮想COMオブジェクト)
    
    If 実行結果 <> 0 Then
        MsgBox "CreateCoreWebView2Controller failed. HRESULT: &H" & Hex(実行結果), vbCritical
        保持状態 = 表示初期化失敗
        保持最終エラー = 実行結果
        
        Exit Sub
    End If
End Sub

' =========================================================================
' 処理概要       : コントローラーの生成結果を保持して状態を更新
' 引数           : エラーコード (Long) - 非同期処理のHRESULT
'                  ポインターコントローラー (LongPtr) - 生成されたコントローラーのポインター
' =========================================================================
Public Sub コントローラーの生成結果を保持して状態を更新(ByVal エラーコード As Long, ByVal ポインターコントローラー As LongPtr)
    保持最終エラー = エラーコード
    
    If エラーコード = 0 And ポインターコントローラー <> 0 Then
    
        保持ポインターコントローラー = ポインターコントローラー
        COM参照追加 保持ポインターコントローラー
        保持状態 = 表示準備済み
    Else
        保持ポインターコントローラー = 0
        保持状態 = 表示初期化失敗
    End If
    
    Set 保持ハンドラー = Nothing
End Sub

' =========================================================================
' 処理概要       : 初期化処理の状態を取得
' 戻り値         : 表示初期化状態 - 初期化処理の状態
' =========================================================================
Public Property Get 状態() As 表示初期化状態
    状態 = 保持状態
End Property

' =========================================================================
' 処理概要       : 初期化が完了して利用できるか判定
' 戻り値         : Boolean - 初期化完了時 True
' =========================================================================
Public Property Get 利用可能() As Boolean
    利用可能 = (保持状態 = 表示準備済み)
End Property

' =========================================================================
' 処理概要       : 初期化が失敗したか判定
' 戻り値         : Boolean - 初期化失敗時 True
' =========================================================================
Public Property Get 初期化失敗() As Boolean
    初期化失敗 = (保持状態 = 表示初期化失敗)
End Property

' =========================================================================
' 処理概要       : 直前の処理で発生したエラーを取得
' 戻り値         : Long - 直前の処理で発生したエラー
' =========================================================================
Public Property Get 最終エラー() As Long
    最終エラー = 保持最終エラー
End Property

' =========================================================================
' 処理概要       : コントローラーのインターフェースポインターを取得
' 戻り値         : LongPtr - 対象のアドレス。取得できない場合 0
' =========================================================================
Public Property Get コントローラーポインター() As LongPtr
    コントローラーポインター = 保持ポインターコントローラー
End Property

' =========================================================================
' 処理概要       : WebView2のインターフェースポインターを取得
' 戻り値         : LongPtr - 対象のアドレス。取得できない場合 0
' =========================================================================
Public Property Get 表示ポインター() As LongPtr
    表示ポインター = 保持ポインターWebView
End Property

' =========================================================================
' 処理概要       : ページの読み込み完了状態を取得
' 戻り値         : Boolean - 読み込み完了時 True
' =========================================================================
Public Property Get 読込完了() As Boolean
    読込完了 = 保持ページ読込済み
End Property

' =========================================================================
' 処理概要       : 保持するページタイトルを取得
' 戻り値         : String - 保持するページタイトル
' =========================================================================
Public Property Get 保持するページタイトルを取得() As String
    保持するページタイトルを取得 = 保持タイトル
End Property

' =========================================================================
' 処理概要       : 設定タブとして使用するペインか取得
' 戻り値         : Boolean - 設定用ペインの場合 True
' =========================================================================
Public Property Get 設定画面() As Boolean
    設定画面 = 保持判定設定
End Property

' =========================================================================
' 処理概要       : 設定タブとして使用するペインか設定
' 引数           : 入力値 (Boolean) - 設定する値
' =========================================================================
Public Property Let 設定画面(ByVal 入力値 As Boolean)
    保持判定設定 = 入力値
End Property

' =========================================================================
' 処理概要       : コントローラーの表示領域を設定
' 引数           : 表示領域 (矩形) - コントローラーの表示領域（ピクセル）
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function コントローラーの表示領域を設定(ByRef 表示領域 As 矩形) As Long
    If 保持ポインターコントローラー = 0 Then
        コントローラーの表示領域を設定 = &H80004003
        Exit Function
    End If
    コントローラーの表示領域を設定 = COMメソッド呼出(保持ポインターコントローラー, 6, "Ctrl.put_Bounds", VarPtr(表示領域))
End Function

' =========================================================================
' 処理概要       : コントローラーの表示状態を設定
' 引数           : 表示状態 (Boolean) - 表示する場合 True
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function コントローラーの表示状態を設定(ByVal 表示状態 As Boolean) As Long
    Dim 入力データ As Long
    
    If 保持ポインターコントローラー = 0 Then
        コントローラーの表示状態を設定 = &H80004003
        Exit Function
    End If
    If 表示状態 Then 入力データ = 1 Else 入力データ = 0
    コントローラーの表示状態を設定 = COMメソッド呼出(保持ポインターコントローラー, 4, "Ctrl.put_IsVisible", 入力データ)
End Function

' =========================================================================
' 処理概要       : コントローラーからWebView2インターフェースを取得して保持
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function コントローラーからWebView2インターフェースを取得して保持() As Long
    Dim 引数領域 As LongPtr
    Dim 実行結果 As Long
    
    If 保持ポインターコントローラー = 0 Then
        コントローラーからWebView2インターフェースを取得して保持 = &H80004003
        Exit Function
    End If
    
    If 保持ポインターWebView <> 0 Then
        COM参照解放 保持ポインターWebView
        保持ポインターWebView = 0
    End If
    
    実行結果 = COMメソッド呼出(保持ポインターコントローラー, 25, "Ctrl.get_CoreWebView2", VarPtr(引数領域))
    If 実行結果 = 0 Then
        保持ポインターWebView = 引数領域
    End If
    コントローラーからWebView2インターフェースを取得して保持 = 実行結果
End Function

' =========================================================================
' 処理概要       : 指定URLへの移動を開始
' 引数           : URL (String) - 移動先または通知対象のURL
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function 指定URLへの移動を開始(ByVal URL As String) As Long
    If 保持ポインターWebView = 0 Then
        指定URLへの移動を開始 = &H80004003
        Exit Function
    End If
    
    保持ページ読込済み = False
    指定URLへの移動を開始 = COMメソッド呼出(保持ポインターWebView, 5, "View.Navigate", StrPtr(URL))
End Function

' =========================================================================
' 処理概要       : ページ遷移開始イベントを登録
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function ページ遷移開始イベントを登録() As Long
    Dim 通知ハンドラー As mod_WV2_COM通知ハンドラー
    Dim 仮想COMオブジェクト As LongPtr
    Dim 登録トークン As LongLong
    Dim 実行結果 As Long
    
    If 保持ポインターWebView = 0 Then
        ページ遷移開始イベントを登録 = &H80004003
        Exit Function
    End If
    
    If Not (保持永続ハンドラー一覧(遷移開始通知) Is Nothing) Then
    
        ページ遷移開始イベントを登録 = 0
        Exit Function
    End If
    
    Set 通知ハンドラー = 通知先を設定したCOMコールバックハンドラーを取得(遷移開始通知, Me)
    If 通知ハンドラー Is Nothing Then
    
        ページ遷移開始イベントを登録 = &H8000FFFF
        Exit Function
    End If
    
    仮想COMオブジェクト = 通知ハンドラー.スロットアドレス + 仮想COMオブジェクト位置
    
    実行結果 = COMメソッド呼出(保持ポインターWebView, 7, "View.add_NavigationStarting", 仮想COMオブジェクト, VarPtr(登録トークン))
    
    If 実行結果 <> 0 Then
    
        Set 通知ハンドラー = Nothing
        ページ遷移開始イベントを登録 = 実行結果
        Exit Function
    End If
    
    保持永続登録トークン一覧(遷移開始通知) = 登録トークン
    Set 保持永続ハンドラー一覧(遷移開始通知) = 通知ハンドラー
    
    ページ遷移開始イベントを登録 = 0
End Function

' =========================================================================
' 処理概要       : ページ遷移開始イベントを解除
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function ページ遷移開始イベントを解除() As Long
    Dim 実行結果 As Long
    
    If 保持ポインターWebView = 0 Then
        ページ遷移開始イベントを解除 = &H80004003
        Exit Function
    End If
    
    If 保持永続ハンドラー一覧(遷移開始通知) Is Nothing Then
        ページ遷移開始イベントを解除 = 0
        Exit Function
    End If
    
    実行結果 = COMメソッド呼出(保持ポインターWebView, 8, "View.remove_NavigationStarting", 保持永続登録トークン一覧(遷移開始通知))
    
    Set 保持永続ハンドラー一覧(遷移開始通知) = Nothing
    保持永続登録トークン一覧(遷移開始通知) = 0^
    
    ページ遷移開始イベントを解除 = 実行結果
End Function

' =========================================================================
' 処理概要       : ページ遷移開始イベントを受け付けてコールバック状態を管理
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
'                  ポインター引数一覧 (LongPtr) - イベント引数のCOMポインター
' =========================================================================
Public Sub ページ遷移開始イベントを受け付けてコールバック状態を管理(ByVal ポインター通知元 As LongPtr, ByVal ポインター引数一覧 As LongPtr)

    ' コールバック中であることを記録し、同期処理の再入を防ぐ
    保持内部通知 = 保持内部通知 + 1
    On Error GoTo ErrorHanler
    ページ遷移開始イベントの内容を処理 ポインター通知元, ポインター引数一覧
    GoTo EndProc
    
ErrorHanler:
    MsgBox "ブラウザーのイベント処理でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "ブラウザー"
    Resume EndProc
    
EndProc:

    ' 通知処理の終了時にコールバック階層を戻す
    If 保持内部通知 > 0 Then 保持内部通知 = 保持内部通知 - 1
End Sub

' =========================================================================
' 処理概要       : ページ遷移開始イベントの内容を処理
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
'                  ポインター引数一覧 (LongPtr) - イベント引数のCOMポインター
' =========================================================================
Private Sub ページ遷移開始イベントの内容を処理(ByVal ポインター通知元 As LongPtr, ByVal ポインター引数一覧 As LongPtr)
    Dim URL As String
    
    If ポインター引数一覧 = 0 Then
    
        Exit Sub
    End If
    
    URL = COMメソッドから文字列プロパティを取得(ポインター引数一覧, 3, "NavStartingArgs.get_Uri")
    
End Sub

' =========================================================================
' 処理概要       : ページ遷移完了イベントを登録
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function ページ遷移完了イベントを登録() As Long
    Dim 通知ハンドラー As mod_WV2_COM通知ハンドラー
    Dim 仮想COMオブジェクト As LongPtr
    Dim 登録トークン As LongLong
    Dim 実行結果 As Long
    
    If 保持ポインターWebView = 0 Then
        ページ遷移完了イベントを登録 = &H80004003
        Exit Function
    End If
    
    If Not (保持永続ハンドラー一覧(遷移完了通知) Is Nothing) Then
    
        ページ遷移完了イベントを登録 = 0
        Exit Function
    End If
    
    Set 通知ハンドラー = 通知先を設定したCOMコールバックハンドラーを取得(遷移完了通知, Me)
    If 通知ハンドラー Is Nothing Then
    
        ページ遷移完了イベントを登録 = &H8000FFFF
        Exit Function
    End If
    
    仮想COMオブジェクト = 通知ハンドラー.スロットアドレス + 仮想COMオブジェクト位置
    
    実行結果 = COMメソッド呼出(保持ポインターWebView, 15, "View.add_NavigationCompleted", 仮想COMオブジェクト, VarPtr(登録トークン))
    
    If 実行結果 <> 0 Then
    
        Set 通知ハンドラー = Nothing
        ページ遷移完了イベントを登録 = 実行結果
        Exit Function
    End If
    
    保持永続登録トークン一覧(遷移完了通知) = 登録トークン
    Set 保持永続ハンドラー一覧(遷移完了通知) = 通知ハンドラー
    
    ページ遷移完了イベントを登録 = 0
End Function

' =========================================================================
' 処理概要       : ページ遷移完了イベントを解除
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function ページ遷移完了イベントを解除() As Long
    Dim 実行結果 As Long
    
    If 保持ポインターWebView = 0 Then
        ページ遷移完了イベントを解除 = &H80004003
        Exit Function
    End If
    
    If 保持永続ハンドラー一覧(遷移完了通知) Is Nothing Then
        ページ遷移完了イベントを解除 = 0
        Exit Function
    End If
    
    実行結果 = COMメソッド呼出(保持ポインターWebView, 16, "View.remove_NavigationCompleted", 保持永続登録トークン一覧(遷移完了通知))
    
    Set 保持永続ハンドラー一覧(遷移完了通知) = Nothing
    保持永続登録トークン一覧(遷移完了通知) = 0^
    
    ページ遷移完了イベントを解除 = 実行結果
End Function

' =========================================================================
' 処理概要       : ページ遷移完了イベントを受け付けてコールバック状態を管理
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
'                  ポインター引数一覧 (LongPtr) - イベント引数のCOMポインター
' =========================================================================
Public Sub ページ遷移完了イベントを受け付けてコールバック状態を管理(ByVal ポインター通知元 As LongPtr, ByVal ポインター引数一覧 As LongPtr)

    ' コールバック中であることを記録し、同期処理の再入を防ぐ
    保持内部通知 = 保持内部通知 + 1
    On Error GoTo ErrorHanler
    ページ遷移完了イベントの内容を処理 ポインター通知元, ポインター引数一覧
    GoTo EndProc
    
ErrorHanler:
    MsgBox "ブラウザーのイベント処理でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "ブラウザー"
    Resume EndProc
    
EndProc:

    ' 通知処理の終了時にコールバック階層を戻す
    If 保持内部通知 > 0 Then 保持内部通知 = 保持内部通知 - 1
End Sub

' =========================================================================
' 処理概要       : ページ遷移完了イベントの内容を処理
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
'                  ポインター引数一覧 (LongPtr) - イベント引数のCOMポインター
' =========================================================================
Private Sub ページ遷移完了イベントの内容を処理(ByVal ポインター通知元 As LongPtr, ByVal ポインター引数一覧 As LongPtr)
    Dim 判定成功 As Boolean
    
    If ポインター引数一覧 = 0 Then
    
        保持ページ読込済み = True
        
        現在のURLと履歴の状態をブラウザーへ通知 ポインター通知元
        Exit Sub
    End If
    
    判定成功 = COMメソッドから真偽値プロパティを取得(ポインター引数一覧, 3, "NavCompletedArgs.get_IsSuccess")
    
    保持ページ読込済み = True
    
    現在のURLと履歴の状態をブラウザーへ通知 ポインター通知元
End Sub

' =========================================================================
' 処理概要       : 現在のURLと履歴の状態をブラウザーへ通知
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
' =========================================================================
Private Sub 現在のURLと履歴の状態をブラウザーへ通知(ByVal ポインター通知元 As LongPtr)
    Dim URL As String
    Dim 可能戻る As Boolean
    Dim 可能進む As Boolean
    
    If 保持ブラウザー Is Nothing Then Exit Sub
    If ポインター通知元 = 0 Then Exit Sub
    
    URL = COMメソッドから文字列プロパティを取得(ポインター通知元, 4, "View.get_Source")
    可能戻る = COMメソッドから真偽値プロパティを取得(ポインター通知元, 38, "View.get_CanGoBack")
    可能進む = COMメソッドから真偽値プロパティを取得(ポインター通知元, 39, "View.get_CanGoForward")
    
    保持ブラウザー.ペインのURLと履歴状態をタブのイベントとして通知 Me, URL, 可能戻る, 可能進む
End Sub

' =========================================================================
' 処理概要       : ページタイトル変更イベントを登録
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function ページタイトル変更イベントを登録() As Long
    Dim 通知ハンドラー As mod_WV2_COM通知ハンドラー
    Dim 仮想COMオブジェクト As LongPtr
    Dim 登録トークン As LongLong
    Dim 実行結果 As Long
    
    If 保持ポインターWebView = 0 Then
        ページタイトル変更イベントを登録 = &H80004003
        Exit Function
    End If
    
    If Not (保持永続ハンドラー一覧(タイトル変更通知) Is Nothing) Then
    
        ページタイトル変更イベントを登録 = 0
        Exit Function
    End If
    
    Set 通知ハンドラー = 通知先を設定したCOMコールバックハンドラーを取得(タイトル変更通知, Me)
    If 通知ハンドラー Is Nothing Then
    
        ページタイトル変更イベントを登録 = &H8000FFFF
        Exit Function
    End If
    
    仮想COMオブジェクト = 通知ハンドラー.スロットアドレス + 仮想COMオブジェクト位置
    
    実行結果 = COMメソッド呼出(保持ポインターWebView, 46, "View.add_DocumentTitleChanged", 仮想COMオブジェクト, VarPtr(登録トークン))
    
    If 実行結果 <> 0 Then
    
        Set 通知ハンドラー = Nothing
        ページタイトル変更イベントを登録 = 実行結果
        Exit Function
    End If
    
    保持永続登録トークン一覧(タイトル変更通知) = 登録トークン
    Set 保持永続ハンドラー一覧(タイトル変更通知) = 通知ハンドラー
    
    ページタイトル変更イベントを登録 = 0
End Function

' =========================================================================
' 処理概要       : ページタイトル変更イベントを解除
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function ページタイトル変更イベントを解除() As Long
    Dim 実行結果 As Long
    
    If 保持ポインターWebView = 0 Then
        ページタイトル変更イベントを解除 = &H80004003
        Exit Function
    End If
    
    If 保持永続ハンドラー一覧(タイトル変更通知) Is Nothing Then
        ページタイトル変更イベントを解除 = 0
        Exit Function
    End If
    
    実行結果 = COMメソッド呼出(保持ポインターWebView, 47, "View.remove_DocumentTitleChanged", 保持永続登録トークン一覧(タイトル変更通知))
    
    Set 保持永続ハンドラー一覧(タイトル変更通知) = Nothing
    保持永続登録トークン一覧(タイトル変更通知) = 0^
    
    ページタイトル変更イベントを解除 = 実行結果
End Function

' =========================================================================
' 処理概要       : ページタイトル変更イベントを受け付けてコールバック状態を管理
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
'                  ポインター引数一覧 (LongPtr) - イベント引数のCOMポインター
' =========================================================================
Public Sub ページタイトル変更イベントを受け付けてコールバック状態を管理(ByVal ポインター通知元 As LongPtr, ByVal ポインター引数一覧 As LongPtr)

    ' コールバック中であることを記録し、同期処理の再入を防ぐ
    保持内部通知 = 保持内部通知 + 1
    On Error GoTo ErrorHanler
    ページタイトル変更イベントの内容を処理 ポインター通知元, ポインター引数一覧
    GoTo EndProc
    
ErrorHanler:
    MsgBox "ブラウザーのイベント処理でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "ブラウザー"
    Resume EndProc
    
EndProc:

    ' 通知処理の終了時にコールバック階層を戻す
    If 保持内部通知 > 0 Then 保持内部通知 = 保持内部通知 - 1
End Sub

' =========================================================================
' 処理概要       : ページタイトル変更イベントの内容を処理
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
'                  ポインター引数一覧 (LongPtr) - イベント引数のCOMポインター
' =========================================================================
Private Sub ページタイトル変更イベントの内容を処理(ByVal ポインター通知元 As LongPtr, ByVal ポインター引数一覧 As LongPtr)
    Dim タイトル As String
    
    タイトル = COMメソッドから文字列プロパティを取得(ポインター通知元, 48, "View.get_DocumentTitle")
    
    保持タイトル = タイトル
    
    If Not 保持ブラウザー Is Nothing Then
        保持ブラウザー.ペインのタイトル変更をタブのイベントとして通知 Me, タイトル
    End If
End Sub

' =========================================================================
' 処理概要       : Webメッセージ受信イベントを登録
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function Webメッセージ受信イベントを登録() As Long
    Dim 通知ハンドラー As mod_WV2_COM通知ハンドラー
    Dim 仮想COMオブジェクト As LongPtr
    Dim 登録トークン As LongLong
    Dim 実行結果 As Long
    
    If 保持ポインターWebView = 0 Then
        Webメッセージ受信イベントを登録 = &H80004003
        Exit Function
    End If
    
    If Not (保持永続ハンドラー一覧(Webメッセージ通知) Is Nothing) Then
    
        Webメッセージ受信イベントを登録 = 0
        Exit Function
    End If
    
    Set 通知ハンドラー = 通知先を設定したCOMコールバックハンドラーを取得(Webメッセージ通知, Me)
    If 通知ハンドラー Is Nothing Then
    
        Webメッセージ受信イベントを登録 = &H8000FFFF
        Exit Function
    End If
    
    仮想COMオブジェクト = 通知ハンドラー.スロットアドレス + 仮想COMオブジェクト位置
    
    実行結果 = COMメソッド呼出(保持ポインターWebView, 34, "View.add_WebMessageReceived", 仮想COMオブジェクト, VarPtr(登録トークン))
    
    If 実行結果 <> 0 Then
    
        Set 通知ハンドラー = Nothing
        Webメッセージ受信イベントを登録 = 実行結果
        Exit Function
    End If
    
    保持永続登録トークン一覧(Webメッセージ通知) = 登録トークン
    Set 保持永続ハンドラー一覧(Webメッセージ通知) = 通知ハンドラー
    
    Webメッセージ受信イベントを登録 = 0
End Function

' =========================================================================
' 処理概要       : Webメッセージ受信イベントを解除
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function Webメッセージ受信イベントを解除() As Long
    Dim 実行結果 As Long
    
    If 保持ポインターWebView = 0 Then
        Webメッセージ受信イベントを解除 = &H80004003
        Exit Function
    End If
    
    If 保持永続ハンドラー一覧(Webメッセージ通知) Is Nothing Then
        Webメッセージ受信イベントを解除 = 0
        Exit Function
    End If
    
    実行結果 = COMメソッド呼出(保持ポインターWebView, 35, "View.remove_WebMessageReceived", 保持永続登録トークン一覧(Webメッセージ通知))
    
    Set 保持永続ハンドラー一覧(Webメッセージ通知) = Nothing
    保持永続登録トークン一覧(Webメッセージ通知) = 0^
    
    Webメッセージ受信イベントを解除 = 実行結果
End Function

' =========================================================================
' 処理概要       : Webメッセージ受信イベントを受け付けてコールバック状態を管理
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
'                  ポインター引数一覧 (LongPtr) - イベント引数のCOMポインター
' =========================================================================
Public Sub Webメッセージ受信イベントを受け付けてコールバック状態を管理(ByVal ポインター通知元 As LongPtr, ByVal ポインター引数一覧 As LongPtr)

    ' コールバック中であることを記録し、同期処理の再入を防ぐ
    保持内部通知 = 保持内部通知 + 1
    On Error GoTo ErrorHanler
    Webメッセージ受信イベントの内容を処理 ポインター通知元, ポインター引数一覧
    GoTo EndProc
    
ErrorHanler:
    MsgBox "ブラウザーのイベント処理でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "ブラウザー"
    Resume EndProc
    
EndProc:

    ' 通知処理の終了時にコールバック階層を戻す
    If 保持内部通知 > 0 Then 保持内部通知 = 保持内部通知 - 1
End Sub

' =========================================================================
' 処理概要       : Webメッセージ受信イベントの内容を処理
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
'                  ポインター引数一覧 (LongPtr) - イベント引数のCOMポインター
' =========================================================================
Private Sub Webメッセージ受信イベントの内容を処理(ByVal ポインター通知元 As LongPtr, ByVal ポインター引数一覧 As LongPtr)
    Dim 元データ As String
    Dim メッセージ As String
    
    If ポインター引数一覧 = 0 Then
    
        Exit Sub
    End If
    
    元データ = COMメソッドから文字列プロパティを取得(ポインター引数一覧, 3, "WebMsgArgs.get_Source")
    
    メッセージ = COMメソッドから文字列プロパティを取得(ポインター引数一覧, 5, "WebMsgArgs.TryGetWebMessageAsString")
    
    If Not 保持Webメッセージ通知先 Is Nothing Then
        保持Webメッセージ通知先.画面から受信したコマンドを振り分けて実行 Me, 元データ, メッセージ
    End If
End Sub

' =========================================================================
' 処理概要       : 新規ウィンドウ要求イベントを登録
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function 新規ウィンドウ要求イベントを登録() As Long
    Dim 通知ハンドラー As mod_WV2_COM通知ハンドラー
    Dim 仮想COMオブジェクト As LongPtr
    Dim 登録トークン As LongLong
    Dim 実行結果 As Long
    
    If 保持ポインターWebView = 0 Then
        新規ウィンドウ要求イベントを登録 = &H80004003
        Exit Function
    End If
    
    If Not (保持永続ハンドラー一覧(新規ウィンドウ通知) Is Nothing) Then
    
        新規ウィンドウ要求イベントを登録 = 0
        Exit Function
    End If
    
    Set 通知ハンドラー = 通知先を設定したCOMコールバックハンドラーを取得(新規ウィンドウ通知, Me)
    If 通知ハンドラー Is Nothing Then
    
        新規ウィンドウ要求イベントを登録 = &H8000FFFF
        Exit Function
    End If
    
    仮想COMオブジェクト = 通知ハンドラー.スロットアドレス + 仮想COMオブジェクト位置
    
    実行結果 = COMメソッド呼出(保持ポインターWebView, 44, "View.add_NewWindowRequested", 仮想COMオブジェクト, VarPtr(登録トークン))
    
    If 実行結果 <> 0 Then
    
        Set 通知ハンドラー = Nothing
        新規ウィンドウ要求イベントを登録 = 実行結果
        Exit Function
    End If
    
    保持永続登録トークン一覧(新規ウィンドウ通知) = 登録トークン
    Set 保持永続ハンドラー一覧(新規ウィンドウ通知) = 通知ハンドラー
    
    新規ウィンドウ要求イベントを登録 = 0
End Function

' =========================================================================
' 処理概要       : 新規ウィンドウ要求イベントを解除
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function 新規ウィンドウ要求イベントを解除() As Long
    Dim 実行結果 As Long
    
    If 保持ポインターWebView = 0 Then
        新規ウィンドウ要求イベントを解除 = &H80004003
        Exit Function
    End If
    
    If 保持永続ハンドラー一覧(新規ウィンドウ通知) Is Nothing Then
        新規ウィンドウ要求イベントを解除 = 0
        Exit Function
    End If
    
    実行結果 = COMメソッド呼出(保持ポインターWebView, 45, "View.remove_NewWindowRequested", 保持永続登録トークン一覧(新規ウィンドウ通知))
    
    Set 保持永続ハンドラー一覧(新規ウィンドウ通知) = Nothing
    保持永続登録トークン一覧(新規ウィンドウ通知) = 0^
    
    新規ウィンドウ要求イベントを解除 = 実行結果
End Function

' =========================================================================
' 処理概要       : 新規ウィンドウ要求イベントを受け付けてコールバック状態を管理
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
'                  ポインター引数一覧 (LongPtr) - イベント引数のCOMポインター
' =========================================================================
Public Sub 新規ウィンドウ要求イベントを受け付けてコールバック状態を管理(ByVal ポインター通知元 As LongPtr, ByVal ポインター引数一覧 As LongPtr)

    ' コールバック中であることを記録し、同期処理の再入を防ぐ
    保持内部通知 = 保持内部通知 + 1
    On Error GoTo ErrorHanler
    新規ウィンドウ要求イベントの内容を処理 ポインター通知元, ポインター引数一覧
    GoTo EndProc
    
ErrorHanler:
    MsgBox "ブラウザーのイベント処理でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "ブラウザー"
    Resume EndProc
    
EndProc:

    ' 通知処理の終了時にコールバック階層を戻す
    If 保持内部通知 > 0 Then 保持内部通知 = 保持内部通知 - 1
End Sub

' =========================================================================
' 処理概要       : 新規ウィンドウ要求イベントの内容を処理
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
'                  ポインター引数一覧 (LongPtr) - イベント引数のCOMポインター
' =========================================================================
Private Sub 新規ウィンドウ要求イベントの内容を処理(ByVal ポインター通知元 As LongPtr, ByVal ポインター引数一覧 As LongPtr)
    Dim URL As String
    Dim 新規ウィンドウ方式 As Long
    Dim ポインター新規 As LongPtr
    Dim 新規ウィンドウ割当結果 As Long
    Dim 実行結果 As Long
    
    If ポインター引数一覧 = 0 Then
    
        Exit Sub
    End If
    
    URL = COMメソッドから文字列プロパティを取得(ポインター引数一覧, 3, "NewWindowReqArgs.get_Uri")
    
    新規ウィンドウ方式 = 新規タブ作成方式
    If Not (保持ブラウザー Is Nothing) Then 新規ウィンドウ方式 = 保持ブラウザー.新規ウィンドウ方式
    
    If 新規ウィンドウ方式 = ポップアップ委譲方式 Then
    
        Exit Sub
        
    ElseIf 新規ウィンドウ方式 = 予備タブ利用方式 Then
    
        ポインター新規 = 保持ブラウザー.事前生成したペインを予約してWebView2ポインターを取得()
        
        If ポインター新規 = 0 Then
        
            保持ブラウザー.予備ペインの生成を予約
            
        Else
            新規ウィンドウ割当結果 = COMメソッド呼出(ポインター引数一覧, 4, "NewWindowReqArgs.put_NewWindow", ポインター新規)
            
            If 新規ウィンドウ割当結果 <> 0 Then
            
                保持ブラウザー.事前生成ペインの利用予約を解除
                
            Else
            
                If 保持ブラウザー.新規ウィンドウ割当後に処理済みにする Then
                    新規ウィンドウ割当結果 = COMメソッド呼出(ポインター引数一覧, 6, "NewWindowReqArgs.put_Handled", 1&)
                End If
                
                保持ブラウザー.予備ペインのタブへの取り込みを予約 URL
                Exit Sub
            End If
        End If
    End If
    
    If URL = "" Then
    
        Exit Sub
    End If
    
    実行結果 = COMメソッド呼出(ポインター引数一覧, 6, "NewWindowReqArgs.put_Handled", 1&)
    
    If 保持ブラウザー Is Nothing Then
    
        実行結果 = 指定URLへの移動を開始(URL)
        
        Exit Sub
    End If
    
    保持ブラウザー.新規タブの作成をタイマー経由で予約 URL
    
End Sub

' =========================================================================
' 処理概要       : ペインの所有元ブラウザーを設定
' 引数           : ブラウザー (mod_WV2_ブラウザー管理) - 操作対象のブラウザー
' =========================================================================
Public Sub ペインの所有元ブラウザーを設定(ByVal ブラウザー As mod_WV2_ブラウザー管理)
    Set 保持ブラウザー = ブラウザー
End Sub

' =========================================================================
' 処理概要       : 所有元ブラウザーへの参照を解除
' =========================================================================
Public Sub 所有元ブラウザーへの参照を解除()
    Set 保持ブラウザー = Nothing
    Set 保持Webメッセージ通知先 = Nothing
End Sub

' =========================================================================
' 処理概要       : Webメッセージの通知先を設定
' 引数           : 通知先 (Object) - Webメッセージの通知先オブジェクト
' =========================================================================
Public Sub Webメッセージの通知先を設定(ByVal 通知先 As Object)
    Set 保持Webメッセージ通知先 = 通知先
End Sub

' =========================================================================
' 処理概要       : Webメッセージの通知先を解除
' =========================================================================
Public Sub Webメッセージの通知先を解除()
    Set 保持Webメッセージ通知先 = Nothing
End Sub

' =========================================================================
' 処理概要       : 履歴変更イベントを登録
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function 履歴変更イベントを登録() As Long
    Dim 通知ハンドラー As mod_WV2_COM通知ハンドラー
    Dim 仮想COMオブジェクト As LongPtr
    Dim 登録トークン As LongLong
    Dim 実行結果 As Long
    
    If 保持ポインターWebView = 0 Then
        履歴変更イベントを登録 = &H80004003
        Exit Function
    End If
    
    If Not (保持永続ハンドラー一覧(履歴変更通知) Is Nothing) Then
    
        履歴変更イベントを登録 = 0
        Exit Function
    End If
    
    Set 通知ハンドラー = 通知先を設定したCOMコールバックハンドラーを取得(履歴変更通知, Me)
    If 通知ハンドラー Is Nothing Then
    
        履歴変更イベントを登録 = &H8000FFFF
        Exit Function
    End If
    
    仮想COMオブジェクト = 通知ハンドラー.スロットアドレス + 仮想COMオブジェクト位置
    
    実行結果 = COMメソッド呼出(保持ポインターWebView, 13, "View.add_HistoryChanged", 仮想COMオブジェクト, VarPtr(登録トークン))
    
    If 実行結果 <> 0 Then
    
        Set 通知ハンドラー = Nothing
        履歴変更イベントを登録 = 実行結果
        Exit Function
    End If
    
    保持永続登録トークン一覧(履歴変更通知) = 登録トークン
    Set 保持永続ハンドラー一覧(履歴変更通知) = 通知ハンドラー
    
    履歴変更イベントを登録 = 0
End Function

' =========================================================================
' 処理概要       : 履歴変更イベントを解除
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function 履歴変更イベントを解除() As Long
    Dim 実行結果 As Long
    
    If 保持ポインターWebView = 0 Then
        履歴変更イベントを解除 = &H80004003
        Exit Function
    End If
    
    If 保持永続ハンドラー一覧(履歴変更通知) Is Nothing Then
        履歴変更イベントを解除 = 0
        Exit Function
    End If
    
    実行結果 = COMメソッド呼出(保持ポインターWebView, 14, "View.remove_HistoryChanged", 保持永続登録トークン一覧(履歴変更通知))
    
    Set 保持永続ハンドラー一覧(履歴変更通知) = Nothing
    保持永続登録トークン一覧(履歴変更通知) = 0^
    
    履歴変更イベントを解除 = 実行結果
End Function

' =========================================================================
' 処理概要       : 履歴変更イベントを受け付けてコールバック状態を管理
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
'                  ポインター引数一覧 (LongPtr) - イベント引数のCOMポインター
' =========================================================================
Public Sub 履歴変更イベントを受け付けてコールバック状態を管理(ByVal ポインター通知元 As LongPtr, ByVal ポインター引数一覧 As LongPtr)

    ' コールバック中であることを記録し、同期処理の再入を防ぐ
    保持内部通知 = 保持内部通知 + 1
    On Error GoTo ErrorHanler
    履歴変更イベントの内容を処理 ポインター通知元, ポインター引数一覧
    GoTo EndProc
    
ErrorHanler:
    MsgBox "ブラウザーのイベント処理でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "ブラウザー"
    Resume EndProc
    
EndProc:

    ' 通知処理の終了時にコールバック階層を戻す
    If 保持内部通知 > 0 Then 保持内部通知 = 保持内部通知 - 1
End Sub

' =========================================================================
' 処理概要       : 履歴変更イベントの内容を処理
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
'                  ポインター引数一覧 (LongPtr) - イベント引数のCOMポインター
' =========================================================================
Private Sub 履歴変更イベントの内容を処理(ByVal ポインター通知元 As LongPtr, ByVal ポインター引数一覧 As LongPtr)
    Dim URL As String
    Dim 可能戻る As Boolean
    Dim 可能進む As Boolean
    
    URL = COMメソッドから文字列プロパティを取得(ポインター通知元, 4, "View.get_Source")
    可能戻る = COMメソッドから真偽値プロパティを取得(ポインター通知元, 38, "View.get_CanGoBack")
    可能進む = COMメソッドから真偽値プロパティを取得(ポインター通知元, 39, "View.get_CanGoForward")
    
    If Not 保持ブラウザー Is Nothing Then
        保持ブラウザー.ペインのURLと履歴状態をタブのイベントとして通知 Me, URL, 可能戻る, 可能進む
    End If
End Sub

' =========================================================================
' 処理概要       : DOM読み込み完了イベントを登録
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function DOM読み込み完了イベントを登録() As Long
    Dim 通知ハンドラー As mod_WV2_COM通知ハンドラー
    Dim 仮想COMオブジェクト As LongPtr
    Dim 登録トークン As LongLong
    Dim 実行結果 As Long
    
    If Not ICoreWebView2_2を取得して保持() Then
    
        DOM読み込み完了イベントを登録 = &H80004003
        Exit Function
    End If
    
    If Not (保持永続ハンドラー一覧(DOM読込通知) Is Nothing) Then
    
        DOM読み込み完了イベントを登録 = 0
        Exit Function
    End If
    
    Set 通知ハンドラー = 通知先を設定したCOMコールバックハンドラーを取得(DOM読込通知, Me)
    If 通知ハンドラー Is Nothing Then
    
        DOM読み込み完了イベントを登録 = &H8000FFFF
        Exit Function
    End If
    
    仮想COMオブジェクト = 通知ハンドラー.スロットアドレス + 仮想COMオブジェクト位置
    
    実行結果 = COMメソッド呼出(保持ポインターWebView2, 64, "View2.add_DOMContentLoaded", 仮想COMオブジェクト, VarPtr(登録トークン))
    
    If 実行結果 <> 0 Then
    
        Set 通知ハンドラー = Nothing
        DOM読み込み完了イベントを登録 = 実行結果
        Exit Function
    End If
    
    保持永続登録トークン一覧(DOM読込通知) = 登録トークン
    Set 保持永続ハンドラー一覧(DOM読込通知) = 通知ハンドラー
    
    DOM読み込み完了イベントを登録 = 0
End Function

' =========================================================================
' 処理概要       : DOM読み込み完了イベントを解除
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function DOM読み込み完了イベントを解除() As Long
    Dim 実行結果 As Long
    
    If Not ICoreWebView2_2を取得して保持() Then
    
        DOM読み込み完了イベントを解除 = &H80004003
        Exit Function
    End If
    
    If 保持永続ハンドラー一覧(DOM読込通知) Is Nothing Then
        DOM読み込み完了イベントを解除 = 0
        Exit Function
    End If
    
    実行結果 = COMメソッド呼出(保持ポインターWebView2, 65, "View2.remove_DOMContentLoaded", 保持永続登録トークン一覧(DOM読込通知))
    
    Set 保持永続ハンドラー一覧(DOM読込通知) = Nothing
    保持永続登録トークン一覧(DOM読込通知) = 0^
    
    DOM読み込み完了イベントを解除 = 実行結果
End Function

' =========================================================================
' 処理概要       : DOM読み込み完了イベントを受け付けてコールバック状態を管理
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
'                  ポインター引数一覧 (LongPtr) - イベント引数のCOMポインター
' =========================================================================
Public Sub DOM読み込み完了イベントを受け付けてコールバック状態を管理(ByVal ポインター通知元 As LongPtr, ByVal ポインター引数一覧 As LongPtr)

    ' コールバック中であることを記録し、同期処理の再入を防ぐ
    保持内部通知 = 保持内部通知 + 1
    On Error GoTo ErrorHanler
    DOM読み込み完了イベントの内容を処理 ポインター通知元, ポインター引数一覧
    GoTo EndProc
    
ErrorHanler:
    MsgBox "ブラウザーのイベント処理でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "ブラウザー"
    Resume EndProc
    
EndProc:

    ' 通知処理の終了時にコールバック階層を戻す
    If 保持内部通知 > 0 Then 保持内部通知 = 保持内部通知 - 1
End Sub

' =========================================================================
' 処理概要       : DOM読み込み完了イベントの内容を処理
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
'                  ポインター引数一覧 (LongPtr) - イベント引数のCOMポインター
' =========================================================================
Private Sub DOM読み込み完了イベントの内容を処理(ByVal ポインター通知元 As LongPtr, ByVal ポインター引数一覧 As LongPtr)
    Dim 遷移識別番号 As LongLong
    
    If ポインター引数一覧 = 0 Then
    
        Exit Sub
    End If
    
    遷移識別番号 = COMメソッドから64ビット整数プロパティを取得(ポインター引数一覧, 3, "DOMContentLoadedArgs.get_NavigationId")
    
End Sub

' =========================================================================
' 処理概要       : ページのJavaScriptへ文字列メッセージを送信
' 引数           : Webメッセージ (String) - ページへ送信する文字列メッセージ
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function ページのJavaScriptへ文字列メッセージを送信(ByVal Webメッセージ As String) As Long
    If 保持ポインターWebView = 0 Then
        ページのJavaScriptへ文字列メッセージを送信 = &H80004003
        Exit Function
    End If
    ページのJavaScriptへ文字列メッセージを送信 = COMメソッド呼出(保持ポインターWebView, 33, "View.PostWebMessageAsString", StrPtr(Webメッセージ))
End Function

' =========================================================================
' 処理概要       : 指定HTML文字列をページとして表示
' 引数           : HTMLコンテンツ (String) - 表示するHTML文字列
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function 指定HTML文字列をページとして表示(ByVal HTMLコンテンツ As String) As Long
    If 保持ポインターWebView = 0 Then
        指定HTML文字列をページとして表示 = &H80004003
        Exit Function
    End If
    
    保持ページ読込済み = False
    指定HTML文字列をページとして表示 = COMメソッド呼出(保持ポインターWebView, 6, "View.NavigateToString", StrPtr(HTMLコンテンツ))
End Function

' =========================================================================
' 処理概要       : VBAオブジェクトをJavaScriptへ公開
' 引数           : オブジェクト名称 (String) - JavaScriptへ公開するオブジェクト名
'                  インスタンス (Object) - JavaScriptへ公開するVBAオブジェクト
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function VBAオブジェクトをJavaScriptへ公開(ByVal オブジェクト名称 As String, ByVal インスタンス As Object) As Long
    Dim 値オブジェクト As Variant
    
    If 保持ポインターWebView = 0 Then
        VBAオブジェクトをJavaScriptへ公開 = &H80004003
        Exit Function
    End If
    Set 値オブジェクト = インスタンス
    VBAオブジェクトをJavaScriptへ公開 = COMメソッド呼出(保持ポインターWebView, 49, "View.AddHostObjectToScript", StrPtr(オブジェクト名称), VarPtr(値オブジェクト))
End Function

' =========================================================================
' 処理概要       : 前の履歴へ戻る
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function 前の履歴へ戻る() As Long
    If 保持ポインターWebView = 0 Then
        前の履歴へ戻る = &H80004003
        Exit Function
    End If
    前の履歴へ戻る = COMメソッド呼出(保持ポインターWebView, 40, "View.GoBack")
End Function

' =========================================================================
' 処理概要       : 次の履歴へ進む
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function 次の履歴へ進む() As Long
    If 保持ポインターWebView = 0 Then
        次の履歴へ進む = &H80004003
        Exit Function
    End If
    次の履歴へ進む = COMメソッド呼出(保持ポインターWebView, 41, "View.GoForward")
End Function

' =========================================================================
' 処理概要       : 表示中のページのURLを取得
' 戻り値         : String - 表示中のページのURL
' =========================================================================
Public Function 表示中のページのURLを取得() As String
    If 保持ポインターWebView = 0 Then Exit Function
    
    表示中のページのURLを取得 = COMメソッドから文字列プロパティを取得(保持ポインターWebView, 4, "View.get_Source")
End Function

' =========================================================================
' 処理概要       : 表示中のページのタイトルを取得
' 戻り値         : String - 表示中のページのタイトル
' =========================================================================
Public Function 表示中のページのタイトルを取得() As String
    If 保持ポインターWebView = 0 Then Exit Function
    
    表示中のページのタイトルを取得 = COMメソッドから文字列プロパティを取得(保持ポインターWebView, 48, "View.get_DocumentTitle")
End Function

' =========================================================================
' 処理概要       : ページを再読み込み
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function ページを再読み込み() As Long
    If 保持ポインターWebView = 0 Then
        ページを再読み込み = &H80004003
        Exit Function
    End If
    
    保持ページ読込済み = False
    ページを再読み込み = COMメソッド呼出(保持ポインターWebView, 31, "View.Reload")
End Function

' =========================================================================
' 処理概要       : ページの読み込みを停止
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function ページの読み込みを停止() As Long
    If 保持ポインターWebView = 0 Then
        ページの読み込みを停止 = &H80004003
        Exit Function
    End If
    ページの読み込みを停止 = COMメソッド呼出(保持ポインターWebView, 43, "View.Stop")
End Function

' =========================================================================
' 処理概要       : 前の履歴へ戻れるか判定
' 戻り値         : Boolean - 前の履歴へ戻れる場合 True
' =========================================================================
Public Function 前の履歴へ戻れるか判定() As Boolean
    If 保持ポインターWebView = 0 Then Exit Function
    
    前の履歴へ戻れるか判定 = COMメソッドから真偽値プロパティを取得(保持ポインターWebView, 38, "View.get_CanGoBack")
End Function

' =========================================================================
' 処理概要       : 次の履歴へ進めるか判定
' 戻り値         : Boolean - 次の履歴へ進める場合 True
' =========================================================================
Public Function 次の履歴へ進めるか判定() As Boolean
    If 保持ポインターWebView = 0 Then Exit Function
    
    次の履歴へ進めるか判定 = COMメソッドから真偽値プロパティを取得(保持ポインターWebView, 39, "View.get_CanGoForward")
End Function

' =========================================================================
' 処理概要       : ブラウザープロセスのIDを取得
' 戻り値         : Long - ブラウザープロセスのID
' =========================================================================
Public Function ブラウザープロセスのIDを取得() As Long
    If 保持ポインターWebView = 0 Then Exit Function
    
    ブラウザープロセスのIDを取得 = COMメソッドから32ビット整数プロパティを取得(保持ポインターWebView, 37, "View.get_BrowserProcessId")
End Function

' =========================================================================
' 処理概要       : WebView2設定インターフェースを取得して保持
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Private Function WebView2設定インターフェースを取得して保持() As Boolean
    Dim ポインター設定 As LongPtr
    Dim 実行結果 As Long
    
    If 保持ポインター設定 <> 0 Then
        WebView2設定インターフェースを取得して保持 = True
        Exit Function
    End If
    If 保持ポインターWebView = 0 Then
    
        Exit Function
    End If
    
    実行結果 = COMメソッド呼出(保持ポインターWebView, 3, "View.get_Settings", VarPtr(ポインター設定))
    If 実行結果 <> 0 Then
    
        Exit Function
    End If
    If ポインター設定 = 0 Then
    
        Exit Function
    End If
    
    保持ポインター設定 = ポインター設定
    
    WebView2設定インターフェースを取得して保持 = True
End Function

' =========================================================================
' 処理概要       : JavaScript実行の有効状態を設定
' 引数           : 処理値 (Boolean) - 設定する値
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function JavaScript実行の有効状態を設定(ByVal 処理値 As Boolean) As Long
    Dim バイト値 As Long
    
    If Not WebView2設定インターフェースを取得して保持() Then
        JavaScript実行の有効状態を設定 = &H80004003
        Exit Function
    End If
    If 処理値 Then バイト値 = 1 Else バイト値 = 0
    JavaScript実行の有効状態を設定 = COMメソッド呼出(保持ポインター設定, 4, "Settings.put_IsScriptEnabled", バイト値)
End Function

' =========================================================================
' 処理概要       : 標準コンテキストメニューの有効状態を設定
' 引数           : 処理値 (Boolean) - 設定する値
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function 標準コンテキストメニューの有効状態を設定(ByVal 処理値 As Boolean) As Long
    Dim バイト値 As Long
    
    If Not WebView2設定インターフェースを取得して保持() Then
        標準コンテキストメニューの有効状態を設定 = &H80004003
        Exit Function
    End If
    If 処理値 Then バイト値 = 1 Else バイト値 = 0
    標準コンテキストメニューの有効状態を設定 = COMメソッド呼出(保持ポインター設定, 14, "Settings.put_AreDefaultContextMenusEnabled", バイト値)
End Function

' =========================================================================
' 処理概要       : 開発者ツールの有効状態を設定
' 引数           : 処理値 (Boolean) - 設定する値
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function 開発者ツールの有効状態を設定(ByVal 処理値 As Boolean) As Long
    Dim バイト値 As Long
    
    If Not WebView2設定インターフェースを取得して保持() Then
        開発者ツールの有効状態を設定 = &H80004003
        Exit Function
    End If
    If 処理値 Then バイト値 = 1 Else バイト値 = 0
    開発者ツールの有効状態を設定 = COMメソッド呼出(保持ポインター設定, 12, "Settings.put_AreDevToolsEnabled", バイト値)
End Function

' =========================================================================
' 処理概要       : ICoreWebView2_2を取得して保持
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Private Function ICoreWebView2_2を取得して保持() As Boolean
    Dim 表示拡張2識別子 As GUID情報
    Dim 引数領域 As LongPtr
    Dim 実行結果 As Long
    
    If 保持ポインターWebView2 <> 0 Then
        ICoreWebView2_2を取得して保持 = True
        Exit Function
    End If
    If 保持ポインターWebView = 0 Then
    
        Exit Function
    End If
    
    文字列表記のGUIDをGUID構造体へ設定 表示拡張2識別子, "9E8F0CF8-E670-4B5E-B2BC-73E061E3184C"
    
    実行結果 = COMメソッド呼出(保持ポインターWebView, 0, "View.QueryInterface(ICoreWebView2_2)", VarPtr(表示拡張2識別子), VarPtr(引数領域))
    If 実行結果 <> 0 Then
    
        Exit Function
    End If
    If 引数領域 = 0 Then
    
        Exit Function
    End If
    
    保持ポインターWebView2 = 引数領域
    
    ICoreWebView2_2を取得して保持 = True
End Function

' =========================================================================
' 処理概要       : ICoreWebView2_2が利用できるか確認
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function ICoreWebView2_2が利用できるか確認() As Boolean
    ICoreWebView2_2が利用できるか確認 = ICoreWebView2_2を取得して保持()
End Function

' =========================================================================
' 処理概要       : ICoreWebView2_3を取得して保持
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Private Function ICoreWebView2_3を取得して保持() As Boolean
    Dim 表示拡張3識別子 As GUID情報
    Dim 引数領域 As LongPtr
    Dim 実行結果 As Long
    
    If 保持ポインターWebView3 <> 0 Then
        ICoreWebView2_3を取得して保持 = True
        Exit Function
    End If
    If 保持ポインターWebView = 0 Then
    
        Exit Function
    End If
    
    文字列表記のGUIDをGUID構造体へ設定 表示拡張3識別子, "A0D6DF20-3B92-416D-AA0C-437A9C727857"
    
    実行結果 = COMメソッド呼出(保持ポインターWebView, 0, "View.QueryInterface(ICoreWebView2_3)", VarPtr(表示拡張3識別子), VarPtr(引数領域))
    If 実行結果 <> 0 Then
    
        Exit Function
    End If
    If 引数領域 = 0 Then
    
        Exit Function
    End If
    
    保持ポインターWebView3 = 引数領域
    
    ICoreWebView2_3を取得して保持 = True
End Function

' =========================================================================
' 処理概要       : ICoreWebView2_3が利用できるか確認
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function ICoreWebView2_3が利用できるか確認() As Boolean
    ICoreWebView2_3が利用できるか確認 = ICoreWebView2_3を取得して保持()
End Function

' =========================================================================
' 処理概要       : 仮想ホスト名にローカルフォルダーを割り当て
' 引数           : ホスト名称 (String) - ローカルフォルダーへ割り当てる仮想ホスト名
'                  フォルダーパス (String) - 対象の保存先またはHTMLフォルダー
'                  アクセス種別 (Long) - 仮想ホストのリソースアクセス方式
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function 仮想ホスト名にローカルフォルダーを割り当て(ByVal ホスト名称 As String, ByVal フォルダーパス As String, ByVal アクセス種別 As Long) As Long
    Dim 実行結果 As Long
    
    If Not ICoreWebView2_3を取得して保持() Then
    
        仮想ホスト名にローカルフォルダーを割り当て = &H80004003
        Exit Function
    End If
    
    実行結果 = COMメソッド呼出(保持ポインターWebView3, 71, "View3.SetVirtualHostNameToFolderMapping", StrPtr(ホスト名称), StrPtr(フォルダーパス), アクセス種別)
    
    仮想ホスト名にローカルフォルダーを割り当て = 実行結果
End Function

' =========================================================================
' 処理概要       : イベントとコントローラーを閉じて関連参照を解放
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function イベントとコントローラーを閉じて関連参照を解放() As Long
    Dim 実行結果 As Long
    Dim 参照キー As Long
    
    If 保持終了済み Then
    
        イベントとコントローラーを閉じて関連参照を解放 = 0
        Exit Function
    End If
    If 保持ポインターコントローラー = 0 Then
        イベントとコントローラーを閉じて関連参照を解放 = &H80004003
        Exit Function
    End If
    
    実行結果 = COMメソッド呼出(保持ポインターコントローラー, 24, "Ctrl.Close")
    
    If 保持ポインター設定 <> 0 Then
    
        COM参照解放 保持ポインター設定
        保持ポインター設定 = 0
    End If
    If 保持ポインターWebView3 <> 0 Then
    
        COM参照解放 保持ポインターWebView3
        保持ポインターWebView3 = 0
    End If
    If 保持ポインターWebView2 <> 0 Then
    
        COM参照解放 保持ポインターWebView2
        保持ポインターWebView2 = 0
    End If
    If 保持ポインターWebView <> 0 Then
    
        COM参照解放 保持ポインターWebView
        保持ポインターWebView = 0
    End If
    If 保持ポインターコントローラー <> 0 Then
    
        COM参照解放 保持ポインターコントローラー
        保持ポインターコントローラー = 0
    End If
    
    For 参照キー = LBound(保持永続ハンドラー一覧) To UBound(保持永続ハンドラー一覧)
        保持永続登録トークン一覧(参照キー) = 0^
        Set 保持永続ハンドラー一覧(参照キー) = Nothing
    Next 参照キー
    
    保持終了済み = True
    保持状態 = 表示初期化失敗
    
    イベントとコントローラーを閉じて関連参照を解放 = 実行結果
End Function

' =========================================================================
' 処理概要       : JavaScriptを非同期で実行
' 引数           : JavaScriptコード (String) - 実行するJavaScript
' 戻り値         : Long - 受付成功時は正のコールバックID / 失敗時は負のエラー値
' =========================================================================
Public Function JavaScriptを非同期で実行(ByVal JavaScriptコード As String) As Long
    Dim バイト数識別番号 As Long
    Dim 通知ハンドラー As mod_WV2_COM通知ハンドラー
    Dim 仮想COMオブジェクト As LongPtr
    Dim 実行結果 As Long
    
    If 保持終了済み Then
    
        JavaScriptを非同期で実行 = -1
        Exit Function
    End If
    If 保持ポインターWebView = 0 Then
    
        JavaScriptを非同期で実行 = -2
        Exit Function
    End If
    
    保持実行スクリプト連番 = 保持実行スクリプト連番 + 1
    バイト数識別番号 = 保持実行スクリプト連番
    
    Set 通知ハンドラー = 通知先を設定したCOMコールバックハンドラーを取得(スクリプト完了通知, Me)
    If 通知ハンドラー Is Nothing Then
    
        JavaScriptを非同期で実行 = -3
        Exit Function
    End If
    
    通知ハンドラー.通知識別番号 = バイト数識別番号
    
    仮想COMオブジェクト = 通知ハンドラー.スロットアドレス + 仮想COMオブジェクト位置
    
    実行結果 = COMメソッド呼出(保持ポインターWebView, 29, "View.ExecuteScript", StrPtr(JavaScriptコード), 仮想COMオブジェクト)
    
    Set 通知ハンドラー = Nothing
    
    If 実行結果 <> 0 Then
    
        JavaScriptを非同期で実行 = -4
        Exit Function
    End If
    
    JavaScriptを非同期で実行 = バイト数識別番号
End Function

' =========================================================================
' 処理概要       : 例外情報をJSONへ変換するJavaScriptを実行
' 引数           : ユーザーコード (String) - 例外を捕捉して実行するJavaScript
' 戻り値         : Long - 受付成功時は正のコールバックID / 失敗時は負のエラー値
' =========================================================================
Public Function 例外情報をJSONへ変換するJavaScriptを実行(ByVal ユーザーコード As String) As Long
    Dim 例外保護JavaScript As String
    
    例外保護JavaScript = "(function(){try{var __r=(" & ユーザーコード & ");" & "return {ok:true,value:__r};}" & "catch(e){return {ok:false,errorName:e.name," & "errorMessage:e.message,stack:e.stack};}})()"
    
    例外情報をJSONへ変換するJavaScriptを実行 = JavaScriptを非同期で実行(例外保護JavaScript)
End Function

' =========================================================================
' 処理概要       : 非同期JavaScriptの結果を待機中の処理またはフォームへ通知
' 引数           : 通知識別番号 (Long) - 非同期処理を識別するID
'                  エラーコード (Long) - 非同期処理のHRESULT
'                  結果JSON (String) - JavaScriptの実行結果を表すJSON
' =========================================================================
Public Sub 非同期JavaScriptの結果を待機中の処理またはフォームへ通知(ByVal 通知識別番号 As Long, ByVal エラーコード As Long, ByVal 結果JSON As String)
    Dim 参照キー As String
    Dim 結果項目 As Variant
    Dim 検出済み As Boolean
    
    参照キー = "cb" & 通知識別番号
    If Not 保持スクリプト評価待機中 Is Nothing Then
        検出済み = False
        On Error Resume Next
        結果項目 = 保持スクリプト評価待機中(参照キー)
        検出済み = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
        
        If 検出済み Then
            If 結果項目(0) = スクリプト結果破棄済み Then
            
                保持スクリプト評価待機中.Remove 参照キー
            Else
                保持スクリプト評価待機中.Remove 参照キー
                保持スクリプト評価待機中.Add Array(スクリプト結果取得済み, エラーコード, 結果JSON), 参照キー
            End If
        End If
    End If
    
    ' コールバック中であることを記録し、同期処理の再入を防ぐ
    保持内部通知 = 保持内部通知 + 1
    On Error GoTo ErrorHanler
    RaiseEvent スクリプト実行完了(通知識別番号, エラーコード, 結果JSON)
    If Not 保持ブラウザー Is Nothing Then 保持ブラウザー.ペインのJavaScript実行結果をブラウザーのイベントとして通知 Me, 通知識別番号, エラーコード, 結果JSON
    
EndProc:
    保持内部通知 = 保持内部通知 - 1
    Exit Sub
    
ErrorHanler:
    MsgBox "処理を完了できません。" & vbCrLf & Err.Description, vbExclamation, "OnExecuteScriptCompleted"
    mod_WV2_ログ管理.エラーレベルのログを記録 "ブラウザー表示.ScriptCompleted: 通知先の処理でエラーが発生しました。" & Err.Description
    Resume EndProc
End Sub

' =========================================================================
' 処理概要       : JavaScript式を評価して完了まで待機
' 引数           : JavaScript式 (String) - 同期評価するJavaScript式
'                  待機上限秒 (Single) - 待機の上限秒数（省略時: 5）
' 戻り値         : String - JavaScriptの評価結果JSON。成否はLastEvalOkで確認
' =========================================================================
Public Function JavaScript式を評価して完了まで待機(ByVal JavaScript式 As String, Optional ByVal 待機上限秒 As Single = 5) As String
    Dim 実行結果値 As String
    Dim バイト数識別番号 As Long
    Dim 参照キー As String
    Dim 開始時刻 As Single
    Dim 結果項目 As Variant
    Dim 時間出力 As Boolean
    Dim エラーコード As Long
    Dim 結果JSON As String
    Dim 処理値 As String
    Dim 例外名称 As String
    Dim 例外メッセージ As String
    
    実行結果値 = ""
    保持最終スクリプト評価成功 = False
    保持最終スクリプト評価エラー = ""
    
    If 保持終了済み Then
        保持最終スクリプト評価エラー = "closed"
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.EvalSync: Close 済みのため拒否"
        JavaScript式を評価して完了まで待機 = ""
        Exit Function
    End If
    
    If 保持ポインターWebView = 0 Then
        保持最終スクリプト評価エラー = "no-view"
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.EvalSync: m_pView が未取得"
        JavaScript式を評価して完了まで待機 = ""
        Exit Function
    End If
    
    If 保持内部通知 > 0 Then
    
        保持最終スクリプト評価エラー = "in-callback"
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.EvalSync: コールバック処理中は同期評価できません。安全タイマーから実行してください。"
        JavaScript式を評価して完了まで待機 = ""
        Exit Function
    End If
    
    If 待機上限秒 <= 0 Then 待機上限秒 = 既定スクリプト待機秒
    If 保持スクリプト評価待機中 Is Nothing Then Set 保持スクリプト評価待機中 = New Collection
    
    ' -------------
    ' 評価要求を発行し、完了通知をコールバックIDで対応付け
    ' -------------
    バイト数識別番号 = 例外情報をJSONへ変換するJavaScriptを実行("JSON.stringify(" & JavaScript式 & ")")
    If バイト数識別番号 <= 0 Then
        保持最終スクリプト評価エラー = "execute-failed(" & バイト数識別番号 & ")"
        mod_WV2_ログ管理.エラーレベルのログを記録 "ブラウザー表示.EvalSync: ExecuteScript の発行に失敗 rc=" & バイト数識別番号
        JavaScript式を評価して完了まで待機 = ""
        Exit Function
    End If
    
    参照キー = "cb" & バイト数識別番号
    保持スクリプト評価待機中.Add Array(スクリプト結果待機中, 0&, ""), 参照キー
    
    保持スクリプト評価階層数 = 保持スクリプト評価階層数 + 1
    
    時間出力 = False
    開始時刻 = Timer
    Do
        DoEvents
        結果項目 = 保持スクリプト評価待機中(参照キー)
        If 結果項目(0) = スクリプト結果取得済み Then Exit Do
        If (Timer - 開始時刻) > 待機上限秒 Then
            時間出力 = True
            
            保持スクリプト評価待機中.Remove 参照キー
            保持スクリプト評価待機中.Add Array(スクリプト結果破棄済み, 0&, ""), 参照キー
            Exit Do
        End If
    Loop
    
    保持スクリプト評価階層数 = 保持スクリプト評価階層数 - 1
    
    If 時間出力 Then
        保持最終スクリプト評価エラー = "timeout(" & 待機上限秒 & "s)"
        mod_WV2_ログ管理.エラーレベルのログを記録 "ブラウザー表示.EvalSync: タイムアウト " & 参照キー & " を破棄済みにした " & "(遅れて完了通知が来ても捨てる)"
        JavaScript式を評価して完了まで待機 = ""
        Exit Function
    End If
    
    エラーコード = 結果項目(1)
    結果JSON = 結果項目(2)
    保持スクリプト評価待機中.Remove 参照キー
    
    If エラーコード <> 0 Then
        保持最終スクリプト評価エラー = "hresult(&H" & Hex$(エラーコード) & ")"
        mod_WV2_ログ管理.エラーレベルのログを記録 "ブラウザー表示.EvalSync: ExecuteScript が失敗 " & 保持最終スクリプト評価エラー
        JavaScript式を評価して完了まで待機 = ""
        Exit Function
    End If
    
    If InStr(1, 結果JSON, """ok"":true") > 0 Then
        If mod_WV2_JSON共通.指定キーの文字列を復号してJSONから取得(結果JSON, "value", 処理値) Then
            保持最終スクリプト評価成功 = True
            実行結果値 = 処理値
        Else
        
            保持最終スクリプト評価成功 = True
            実行結果値 = "undefined"
        End If
        
    ElseIf InStr(1, 結果JSON, """ok"":false") > 0 Then
        mod_WV2_JSON共通.指定キーの文字列を復号してJSONから取得 結果JSON, "errorName", 例外名称
        mod_WV2_JSON共通.指定キーの文字列を復号してJSONから取得 結果JSON, "errorMessage", 例外メッセージ
        保持最終スクリプト評価エラー = 例外名称 & ": " & 例外メッセージ
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.EvalSync: JS 例外 " & 保持最終スクリプト評価エラー
        
    Else
    
        保持最終スクリプト評価エラー = "syntax-error-or-null (resultJson=" & Left$(結果JSON, 60) & ")"
        mod_WV2_ログ管理.エラーレベルのログを記録 "ブラウザー表示.EvalSync: " & 保持最終スクリプト評価エラー
    End If
    
    JavaScript式を評価して完了まで待機 = 実行結果値
End Function

' =========================================================================
' 処理概要       : 直前の同期JavaScript評価が成功したか取得
' 戻り値         : Boolean - 直前の同期評価が成功した場合 True
' =========================================================================
Public Property Get 直前の同期JavaScript評価が成功したか取得() As Boolean
    直前の同期JavaScript評価が成功したか取得 = 保持最終スクリプト評価成功
End Property

' =========================================================================
' 処理概要       : 直前の同期JavaScript評価のエラーを取得
' 戻り値         : String - 直前の同期JavaScript評価のエラー
' =========================================================================
Public Property Get 直前の同期JavaScript評価のエラーを取得() As String
    直前の同期JavaScript評価のエラーを取得 = 保持最終スクリプト評価エラー
End Property

' =========================================================================
' 処理概要       : イベントコールバックの呼び出し階層数を取得
' 戻り値         : Long - イベントコールバックの呼び出し階層数
' =========================================================================
Public Property Get イベントコールバックの呼び出し階層数を取得() As Long
    イベントコールバックの呼び出し階層数を取得 = 保持内部通知
End Property

' =========================================================================
' 処理概要       : イベントコールバックの階層数を初期化
' =========================================================================
Public Sub イベントコールバックの階層数を初期化()
    保持内部通知 = 0
End Sub

' =========================================================================
' 処理概要       : 文字列をJavaScriptの文字列リテラルへ変換
' 引数           : 作成文字列 (String) - 変換対象の文字列
' 戻り値         : String - 文字列をJavaScriptの文字列リテラル
' =========================================================================
Friend Function JavaScript文字列に変換(ByVal 作成文字列 As String) As String
    JavaScript文字列に変換 = Chr$(34) & mod_WV2_JSON共通.文字列をJSONに埋め込める形式へ変換(作成文字列) & Chr$(34)
End Function

' =========================================================================
' 処理概要       : DOM要素レジストリを初期化するJavaScriptを生成
' 戻り値         : String - DOM要素レジストリを初期化するJavaScript
' =========================================================================
Private Function DOM要素レジストリを初期化するJavaScriptを生成() As String
    DOM要素レジストリを初期化するJavaScriptを生成 = "if(typeof window.__wv2==='undefined'||!window.__wv2){" & "window.__wv2={gen:'g'+Date.now()+'-'+Math.random(),els:[]};}" & "var R=window.__wv2;"
End Function

' =========================================================================
' 処理概要       : IDでDOM要素を検索
' 引数           : 要素識別番号 (String) - 検索するDOM要素のID
' 戻り値         : mod_WV2_HTML要素 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Public Function IDでDOM要素を検索(ByVal 要素識別番号 As String) As mod_WV2_HTML要素
    Set IDでDOM要素を検索 = タイムアウトまでDOM要素の検索を繰り返す("document.getElementById(" & JavaScript文字列に変換(要素識別番号) & ")", "GetElementById", 要素識別番号, 保持暗黙待機秒)
End Function

' =========================================================================
' 処理概要       : CSSセレクターで最初のDOM要素を検索
' 引数           : セレクター (String) - 検索対象のCSSセレクター
' 戻り値         : mod_WV2_HTML要素 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Public Function CSSセレクターで最初のDOM要素を検索(ByVal セレクター As String) As mod_WV2_HTML要素
    Set CSSセレクターで最初のDOM要素を検索 = タイムアウトまでDOM要素の検索を繰り返す("document.querySelector(" & JavaScript文字列に変換(セレクター) & ")", "QuerySelector", セレクター, 保持暗黙待機秒)
End Function

' =========================================================================
' 処理概要       : DOM要素を検索してレジストリへ登録
' 引数           : 要素検索式 (String) - DOM要素を検索するJavaScript式
'                  API名称 (String) - 呼び出し元の操作名
'                  引数 (String) - 検索条件の表示用文字列
'                  静止 (Boolean) - 要素未検出時の通知を抑える場合 True（省略時: False）
' 戻り値         : mod_WV2_HTML要素 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Private Function DOM要素を検索してレジストリへ登録(ByVal 要素検索式 As String, ByVal API名称 As String, ByVal 引数 As String, Optional ByVal 静止 As Boolean = False) As mod_WV2_HTML要素
    Dim 実行スクリプト As String
    Dim 実行結果値 As String
    Dim 世代 As String
    Dim 通知ハンドラー As Long
    Dim 対象要素 As mod_WV2_HTML要素
    
    Set DOM要素を検索してレジストリへ登録 = Nothing
    
    実行スクリプト = "(function(){" & DOM要素レジストリを初期化するJavaScriptを生成() & "var e=" & 要素検索式 & ";" & "if(!e){return {g:R.gen,found:false,h:0};}" & "R.els.push(e);" & "return {g:R.gen,found:true,h:R.els.length-1};})()"
    
    実行結果値 = JavaScript式を評価して完了まで待機(実行スクリプト)
    
    If Not 保持最終スクリプト評価成功 Then
    
        Exit Function
    End If
    
    If mod_WV2_JSON共通.指定キーの文字列を復号してJSONから取得(実行結果値, "g", 世代) Then
        保持DOM世代 = 世代
    End If
    
    If InStr(1, 実行結果値, """found"":true") = 0 Then
    
        Exit Function
    End If
    
    通知ハンドラー = mod_WV2_JSON共通.JSONから指定キーの数値を取得(実行結果値, "h")
    
    ' 取得した要素をVBAから操作できるよう、参照を作成する
    Set 対象要素 = New mod_WV2_HTML要素
    対象要素.DOM要素の参照情報を設定 Me, 通知ハンドラー, 世代
    Set DOM要素を検索してレジストリへ登録 = 対象要素
    
End Function

' =========================================================================
' 処理概要       : 現在のDOM要素レジストリのページ世代を取得
' 戻り値         : String - 現在のDOM要素レジストリのページ世代
' =========================================================================
Public Property Get 現在のDOM要素レジストリのページ世代を取得() As String
    現在のDOM要素レジストリのページ世代を取得 = 保持DOM世代
End Property

' =========================================================================
' 処理概要       : 取得済みDOM要素を破棄してレジストリを再作成
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 取得済みDOM要素を破棄してレジストリを再作成() As Boolean
    Dim 実行スクリプト As String
    Dim 実行結果値 As String
    Dim 世代 As String
    
    実行スクリプト = "(function(){" & "window.__wv2={gen:'g'+Date.now()+'-'+Math.random(),els:[]};" & "return window.__wv2.gen;})()"
    
    実行結果値 = JavaScript式を評価して完了まで待機(実行スクリプト)
    
    If Not 保持最終スクリプト評価成功 Then
        mod_WV2_ログ管理.エラーレベルのログを記録 "ブラウザー表示.ClearElementRegistry: 失敗 err=" & 保持最終スクリプト評価エラー
        取得済みDOM要素を破棄してレジストリを再作成 = False
        Exit Function
    End If
    
    世代 = mod_WV2_JSON共通.JSONのエスケープ表現を通常の文字列へ復号(実行結果値)
    保持DOM世代 = 世代
    
    取得済みDOM要素を破棄してレジストリを再作成 = True
End Function

' =========================================================================
' 処理概要       : ページ世代を確認してDOM要素の値を取得
' 引数           : 要素ハンドル (Long) - DOM要素のレジストリ番号
'                  要素世代 (String) - 要素を取得したページ世代
'                  値式 (String) - DOM要素から値を読み取るJavaScript式
'                  出力値 (String) - ［出力］取得した要素値の出力先
'                  出力状態 (String) - ［出力］処理状態の出力先
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Friend Function ページ世代を確認してDOM要素の値を取得(ByVal 要素ハンドル As Long, ByVal 要素世代 As String, ByVal 値式 As String, ByRef 出力値 As String, ByRef 出力状態 As String) As Boolean
    Dim 実行スクリプト As String
    Dim 実行結果値 As String
    Dim 世代 As String
    
    出力値 = ""
    出力状態 = ""
    ページ世代を確認してDOM要素の値を取得 = False
    
    If Len(保持DOM世代) > 0 Then
        If 要素世代 <> 保持DOM世代 Then
            出力状態 = "stale"
            Exit Function
        End If
    End If
    
    実行スクリプト = "(function(){" & DOM要素レジストリを初期化するJavaScriptを生成() & "if(R.gen!==" & JavaScript文字列に変換(要素世代) & "){return {s:'gen',g:R.gen};}" & "var e=R.els[" & CStr(要素ハンドル) & "];" & "if(!e){return {s:'nil',g:R.gen};}" & _
        "var x=" & 値式 & ";" & "if(x===null||x===undefined){return {s:'ok',g:R.gen,has:false,v:''};}" & "return {s:'ok',g:R.gen,has:true,v:String(x)};})()"
        
    実行結果値 = JavaScript式を評価して完了まで待機(実行スクリプト)
    
    If Not 保持最終スクリプト評価成功 Then
        出力状態 = 保持最終スクリプト評価エラー
        Exit Function
    End If
    
    If mod_WV2_JSON共通.指定キーの文字列を復号してJSONから取得(実行結果値, "g", 世代) Then
        保持DOM世代 = 世代
    End If
    
    If InStr(1, 実行結果値, """s"":""ok""") = 0 Then
        出力状態 = "stale"
        Exit Function
    End If
    
    If InStr(1, 実行結果値, """has"":true") = 0 Then
    
        出力値 = ""
        ページ世代を確認してDOM要素の値を取得 = True
        Exit Function
    End If
    
    mod_WV2_JSON共通.指定キーの文字列を復号してJSONから取得 実行結果値, "v", 出力値
    ページ世代を確認してDOM要素の値を取得 = True
End Function

' =========================================================================
' 処理概要       : DOM要素のページ世代と参照の有効性を確認
' 引数           : 要素ハンドル (Long) - DOM要素のレジストリ番号
'                  要素世代 (String) - 要素を取得したページ世代
'                  出力状態 (String) - ［出力］処理状態の出力先
' 戻り値         : Boolean - 参照が失効した場合または確認に失敗した場合 True
' =========================================================================
Friend Function DOM要素のページ世代と参照の有効性を確認(ByVal 要素ハンドル As Long, ByVal 要素世代 As String, ByRef 出力状態 As String) As Boolean
    Dim 実行スクリプト As String
    Dim 実行結果値 As String
    Dim 世代 As String
    
    出力状態 = ""
    DOM要素のページ世代と参照の有効性を確認 = True
    
    If Len(保持DOM世代) > 0 Then
        If 要素世代 <> 保持DOM世代 Then Exit Function
    End If
    
    実行スクリプト = "(function(){" & DOM要素レジストリを初期化するJavaScriptを生成() & "var ok=(R.gen===" & JavaScript文字列に変換(要素世代) & ")&&(!!R.els[" & CStr(要素ハンドル) & "]);" & "return {s:ok?'ok':'stale',g:R.gen};})()"
    
    実行結果値 = JavaScript式を評価して完了まで待機(実行スクリプト)
    
    If Not 保持最終スクリプト評価成功 Then
        出力状態 = 保持最終スクリプト評価エラー
        Exit Function
    End If
    
    If mod_WV2_JSON共通.指定キーの文字列を復号してJSONから取得(実行結果値, "g", 世代) Then
        保持DOM世代 = 世代
    End If
    
    DOM要素のページ世代と参照の有効性を確認 = (InStr(1, 実行結果値, """s"":""ok""") = 0)
End Function

' =========================================================================
' 処理概要       : ページ世代を確認してDOM要素の操作を実行
' 引数           : 要素ハンドル (Long) - DOM要素のレジストリ番号
'                  要素世代 (String) - 要素を取得したページ世代
'                  操作実行文 (String) - 要素に対して実行するJavaScript文
'                  出力値 (String) - ［出力］取得した要素値の出力先
'                  出力状態 (String) - ［出力］処理状態の出力先
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Friend Function ページ世代を確認してDOM要素の操作を実行(ByVal 要素ハンドル As Long, ByVal 要素世代 As String, ByVal 操作実行文 As String, ByRef 出力値 As String, ByRef 出力状態 As String) As Boolean
    Dim 実行スクリプト As String
    Dim 実行結果値 As String
    Dim 世代 As String
    
    出力値 = ""
    出力状態 = ""
    ページ世代を確認してDOM要素の操作を実行 = False
    
    If Len(保持DOM世代) > 0 Then
        If 要素世代 <> 保持DOM世代 Then
            出力状態 = "stale"
            Exit Function
        End If
    End If
    
    実行スクリプト = "(function(){" & DOM要素レジストリを初期化するJavaScriptを生成() & "if(R.gen!==" & JavaScript文字列に変換(要素世代) & "){return {s:'gen',g:R.gen};}" & "var e=R.els[" & CStr(要素ハンドル) & "];" & "if(!e){return {s:'nil',g:R.gen};}" & _
        "var rv='';" & 操作実行文 & "return {s:'ok',g:R.gen,v:String(rv)};})()"
        
    実行結果値 = JavaScript式を評価して完了まで待機(実行スクリプト)
    
    If Not 保持最終スクリプト評価成功 Then
        出力状態 = 保持最終スクリプト評価エラー
        Exit Function
    End If
    
    If mod_WV2_JSON共通.指定キーの文字列を復号してJSONから取得(実行結果値, "g", 世代) Then
        保持DOM世代 = 世代
    End If
    
    If InStr(1, 実行結果値, """s"":""ok""") = 0 Then
        出力状態 = "stale"
        Exit Function
    End If
    
    mod_WV2_JSON共通.指定キーの文字列を復号してJSONから取得 実行結果値, "v", 出力値
    ページ世代を確認してDOM要素の操作を実行 = True
End Function

' =========================================================================
' 処理概要       : 要素検索時の暗黙の待機秒数を取得
' 戻り値         : Single - 要素検索時の暗黙の待機秒数
' =========================================================================
Public Property Get 暗黙待機秒数() As Single
    暗黙待機秒数 = 保持暗黙待機秒
End Property

' =========================================================================
' 処理概要       : 要素検索時の暗黙の待機秒数を設定
' 引数           : 待機秒 (Single) - 待機時間（秒）
' =========================================================================
Public Property Let 暗黙待機秒数(ByVal 待機秒 As Single)
    If 待機秒 < 0 Then 待機秒 = 0
    保持暗黙待機秒 = 待機秒
    
End Property

' =========================================================================
' 処理概要       : 指定要素が見つかるまで待機
' 引数           : セレクター (String) - 検索対象のCSSセレクター
'                  待機上限秒 (Single) - 待機の上限秒数（省略時: 既定要素待機秒）
' 戻り値         : mod_WV2_HTML要素 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Public Function 指定要素が見つかるまで待機(ByVal セレクター As String, Optional ByVal 待機上限秒 As Single = 既定要素待機秒) As mod_WV2_HTML要素
    If 待機上限秒 <= 0 Then 待機上限秒 = 既定要素待機秒
    
    Set 指定要素が見つかるまで待機 = タイムアウトまでDOM要素の検索を繰り返す("document.querySelector(" & JavaScript文字列に変換(セレクター) & ")", "WaitFor", セレクター, 待機上限秒)
End Function

' =========================================================================
' 処理概要       : 指定要素が存在しなくなるまで待機
' 引数           : セレクター (String) - 検索対象のCSSセレクター
'                  待機上限秒 (Single) - 待機の上限秒数（省略時: 既定要素待機秒）
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 指定要素が存在しなくなるまで待機(ByVal セレクター As String, Optional ByVal 待機上限秒 As Single = 既定要素待機秒) As Boolean
    Dim 待機開始時刻 As Single
    Dim 検出済み As Boolean
    Dim 成功 As Boolean
    
    指定要素が存在しなくなるまで待機 = False
    If 待機上限秒 <= 0 Then 待機上限秒 = 既定要素待機秒
    
    待機開始時刻 = Timer
    Do
        検出済み = 指定要素の存在と検索処理の成否を取得(セレクター, 成功)
        
        If Not 成功 Then
        
            Exit Function
        End If
        
        If Not 検出済み Then
            指定要素が存在しなくなるまで待機 = True
            Exit Function
        End If
        
        If 日付の切り替わりを考慮して経過秒数を取得(待機開始時刻) >= 待機上限秒 Then
        
            Exit Function
        End If
        
        次回の状態確認までイベント処理を進める
    Loop
End Function

' =========================================================================
' 処理概要       : タイムアウトまでDOM要素の検索を繰り返す
' 引数           : 要素検索式 (String) - DOM要素を検索するJavaScript式
'                  API名称 (String) - 呼び出し元の操作名
'                  引数 (String) - 検索条件の表示用文字列
'                  待機秒 (Single) - 待機時間（秒）
' 戻り値         : mod_WV2_HTML要素 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Private Function タイムアウトまでDOM要素の検索を繰り返す(ByVal 要素検索式 As String, ByVal API名称 As String, ByVal 引数 As String, ByVal 待機秒 As Single) As mod_WV2_HTML要素
    Dim 対象要素 As mod_WV2_HTML要素
    Dim 待機開始時刻 As Single
    
    待機開始時刻 = Timer
    Do
        Set 対象要素 = DOM要素を検索してレジストリへ登録(要素検索式, API名称, 引数, (待機秒 > 0))
        
        If Not 対象要素 Is Nothing Then
            Set タイムアウトまでDOM要素の検索を繰り返す = 対象要素
            Exit Function
        End If
        
        If Not 保持最終スクリプト評価成功 Then Exit Function
        If 待機秒 <= 0 Then Exit Function
        If 日付の切り替わりを考慮して経過秒数を取得(待機開始時刻) >= 待機秒 Then
        
            Exit Function
        End If
        
        次回の状態確認までイベント処理を進める
    Loop
End Function

' =========================================================================
' 処理概要       : 指定要素の存在と検索処理の成否を取得
' 引数           : セレクター (String) - 検索対象のCSSセレクター
'                  出力成功 (Boolean) - ［出力］検索処理の成否の出力先
' 戻り値         : Boolean - 要素が存在する場合 True。検索の成否はoutOkへ設定
' =========================================================================
Private Function 指定要素の存在と検索処理の成否を取得(ByVal セレクター As String, ByRef 出力成功 As Boolean) As Boolean
    Dim 実行スクリプト As String
    Dim 実行結果値 As String
    Dim 世代 As String
    
    指定要素の存在と検索処理の成否を取得 = False
    出力成功 = False
    
    実行スクリプト = "(function(){" & DOM要素レジストリを初期化するJavaScriptを生成() & "var e=document.querySelector(" & JavaScript文字列に変換(セレクター) & ");" & "return {g:R.gen,found:!!e};})()"
    
    実行結果値 = JavaScript式を評価して完了まで待機(実行スクリプト)
    If Not 保持最終スクリプト評価成功 Then Exit Function
    
    If mod_WV2_JSON共通.指定キーの文字列を復号してJSONから取得(実行結果値, "g", 世代) Then
        保持DOM世代 = 世代
    End If
    
    出力成功 = True
    指定要素の存在と検索処理の成否を取得 = (InStr(1, 実行結果値, """found"":true") > 0)
End Function

' =========================================================================
' 処理概要       : 次回の状態確認までイベント処理を進める
' =========================================================================
Private Sub 次回の状態確認までイベント処理を進める()
    Dim 待機開始時刻 As Single
    
    待機開始時刻 = Timer
    Do
        DoEvents
        If 日付の切り替わりを考慮して経過秒数を取得(待機開始時刻) >= 状態確認間隔秒 Then Exit Do
    Loop
End Sub

' =========================================================================
' 処理概要       : 日付の切り替わりを考慮して経過秒数を取得
' 引数           : 開始タイマー (Single) - 計測を開始したTimer値
' 戻り値         : Single - 日付の切り替わりを考慮して経過秒数
' =========================================================================
Private Function 日付の切り替わりを考慮して経過秒数を取得(ByVal 開始タイマー As Single) As Single
    Dim 後続文字 As Single
    
    後続文字 = Timer - 開始タイマー
    If 後続文字 < 0 Then 後続文字 = 後続文字 + 86400
    日付の切り替わりを考慮して経過秒数を取得 = 後続文字
End Function

' =========================================================================
' 処理概要       : SPAの通信とDOM更新の監視状態を取得
' 戻り値         : String - SPAの通信とDOM更新の監視状態
' =========================================================================
Public Function SPAの通信とDOM更新の監視状態を取得() As String
    Dim 実行スクリプト As String
    
    実行スクリプト = "(function(){" & SPAの通信とDOM更新を監視するJavaScriptを生成() & 監視の除外条件をJavaScriptに反映するコードを生成() & "return {v:P.v,h:HP(P),g:P.gen,rp:P.repairs," & "f:P.fetchN,x:P.xhrN," & _
        "q:Math.round(N()-P.lastMut),qn:Math.round(N()-P.lastNet)," & "m:P.mutN,sm:P.skipMut,n:P.netN,sn:P.skipNet," & "up:Math.round(N()-P.since),ob:!!P.ob," & _
        "sda:!!P.sdArmed,sdh:P.sdHits,sna:!!P.snArmed,snh:P.snHits," & "sua:!!P.suArmed," & "suh:(P.suArmed&&String(location.href).indexOf(P.suPat)>=0)?1:0," & "lgN:P.log.length,lgOn:!!P.logOn," & _
        "il:P.ifl.map(function(o){" & "return o.w+' '+Math.round(N()-o.t)+'ms '+o.u;}).slice(0,5).join(' # ')," & "ifn:P.ifl.filter(function(o){return (N()-o.t)<" & 通信監視の有効期限設定をJavaScriptへ変換() & ";}).length," & _
        "ifo:P.ifl.filter(function(o){return (N()-o.t)>=" & 通信監視の有効期限設定をJavaScriptへ変換() & ";}).length};})()"
        
    SPAの通信とDOM更新の監視状態を取得 = JavaScript式を評価して完了まで待機(実行スクリプト)
    
    If Not 保持最終スクリプト評価成功 Then
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.SpaProbeState: 失敗 err=" & 保持最終スクリプト評価エラー
        SPAの通信とDOM更新の監視状態を取得 = ""
    End If
End Function

' =========================================================================
' 処理概要       : SPAの監視機構が動作しているか判定
' 戻り値         : Boolean - 監視機構が正常な場合 True
' =========================================================================
Public Function SPAの監視機構が動作しているか判定() As Boolean
    Dim 作業情報 As String
    
    作業情報 = SPAの通信とDOM更新の監視状態を取得()
    If Not 保持最終スクリプト評価成功 Then Exit Function
    
    SPAの監視機構が動作しているか判定 = (InStr(1, 作業情報, """h"":true") > 0)
End Function

' =========================================================================
' 処理概要       : SPAの監視機構を再作成
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function SPAの監視機構を再作成() As Boolean
    Dim 実行スクリプト As String
    
    実行スクリプト = "(function(){var o=window.__wv2p;if(o){o.v=-1;}" & SPAの通信とDOM更新を監視するJavaScriptを生成() & "return {g:P.gen,h:HP(P)};})()"
    
    JavaScript式を評価して完了まで待機 実行スクリプト
    SPAの監視機構を再作成 = 保持最終スクリプト評価成功
    
    If Not 保持最終スクリプト評価成功 Then
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.SpaProbeReset: 失敗 err=" & 保持最終スクリプト評価エラー
    End If
End Function

' =========================================================================
' 処理概要       : SPAの通信とDOM更新を監視するJavaScriptを生成
' 戻り値         : String - SPAの通信とDOM更新を監視するJavaScript
' =========================================================================
Private Function SPAの通信とDOM更新を監視するJavaScriptを生成() As String
    Dim 実行スクリプト As String
    
    実行スクリプト = "var V=" & CStr(画面監視版) & ",N=function(){return performance.now();};"
    実行スクリプト = 実行スクリプト & "var XP=(window.XMLHttpRequest?XMLHttpRequest.prototype:null);"
    実行スクリプト = 実行スクリプト & "var HP=function(p){return !!p&&p.v===V" & "&&(typeof window.fetch!=='function'||window.fetch===p.fw)" & "&&(!XP||(XP.open===p.xo&&XP.send===p.xs))&&!!p.ob;};"
    実行スクリプト = 実行スクリプト & "var P=window.__wv2p;"
    実行スクリプト = 実行スクリプト & "if(!HP(P)){var prev=P;"
    
    実行スクリプト = 実行スクリプト & "var q={v:V,gen:(prev?((prev.gen||0)+1):1)," & "repairs:(prev?((prev.repairs||0)+1):0)," & "since:N(),lastMut:N(),lastNet:N()," & "fetchN:0,xhrN:0,mutN:0,netN:0,skipMut:0,skipNet:0," & _
        "sdArmed:false,sdSel:'',sdHits:0,sdObs:null," & "snArmed:false,snPat:'',snHits:0," & "suArmed:false,suPat:''," & "ifl:[],log:[],logOn:(prev&&prev.logOn?true:false)," & _
        "ignoreNet:(prev&&prev.ignoreNet?prev.ignoreNet:[])," & "ignoreSel:(prev&&prev.ignoreSel?prev.ignoreSel:'')};"
        
    実行スクリプト = 実行スクリプト & "var ig=function(u){var a=q.ignoreNet,i;" & "for(i=0;i<a.length;i++){if(u.indexOf(a[i])>=0){return true;}}return false;};"
    実行スクリプト = 実行スクリプト & "var sk=function(t){if(!q.ignoreSel){return false;}" & "try{var e=(t&&t.nodeType===1)?t:(t?t.parentElement:null);" & "return !!(e&&e.closest(q.ignoreSel));}catch(x){return false;}};"
    
    実行スクリプト = 実行スクリプト & "var lg=function(k,d){if(!q.logOn){return;}" & "q.log.push(Math.round(N())+' '+k+' '+String(d).slice(0,120));" & "if(q.log.length>200){q.log.shift();}};q.lg=lg;"
    実行スクリプト = 実行スクリプト & "var ifA=function(u,w){var o={u:String(u).slice(0,150),t:N(),w:w};" & "q.ifl.push(o);lg('NET+',w+' '+o.u);return o;};"
    実行スクリプト = 実行スクリプト & "var ifD=function(o){var i=q.ifl.indexOf(o);" & "if(i>=0){q.ifl.splice(i,1);}" & "lg('NET-',Math.round(N()-o.t)+'ms '+o.u);};"
    
    実行スクリプト = 実行スクリプト & "var sn=function(u){if(q.snArmed&&q.snPat&&" & "String(u).indexOf(q.snPat)>=0){q.snHits++;q.lastMut=N();" & "lg('SIG','net hit '+String(u).slice(0,80));}};q.sn=sn;"
    実行スクリプト = 実行スクリプト & "q.armUrl=function(pat){q.suPat=pat;q.suArmed=true;return 0;};"
    実行スクリプト = 実行スクリプト & "q.armDom=function(sel){var n;" & "try{n=document.querySelector(sel);}catch(e){return -2;}" & "if(!n){return -1;}" & "if(q.sdObs){try{q.sdObs.disconnect();}catch(e){}}" & _
        "q.sdSel=sel;q.sdHits=0;q.sdArmed=true;" & "q.sdObs=new MutationObserver(function(ms){" & "q.sdHits+=ms.length;q.lastMut=N();" & "lg('SIG','dom hit x'+ms.length);});" & "q.sdObs.observe(n," & _
        "{childList:true,characterData:true,subtree:true,attributes:true});" & "return 0;};"
    実行スクリプト = 実行スクリプト & "q.armNet=function(pat){q.snPat=pat;q.snHits=0;q.snArmed=true;return 0;};"
    実行スクリプト = 実行スクリプト & "q.disarm=function(){if(q.sdObs){try{q.sdObs.disconnect();}catch(e){}}" & "q.sdObs=null;q.sdArmed=false;q.snArmed=false;q.suArmed=false;" & "return 0;};"
    
    実行スクリプト = 実行スクリプト & "if(typeof window.fetch==='function'){"
    実行スクリプト = 実行スクリプト & "var of=(prev&&prev.of&&prev.fw===window.fetch)?prev.of:window.fetch;q.of=of;"
    
    実行スクリプト = 実行スクリプト & "q.fw=function(u,o){" & "var r=(typeof u==='string')?u:((u&&u.url)?u.url:'');" & "var sp=ig(r);if(sp){q.skipNet++;}else{q.netN++;q.fetchN++;}" & "var en=sp?null:ifA(r,'fetch');" & _
        "var d=function(){if(!sp){q.fetchN--;q.lastNet=N();ifD(en);}sn(r);};" & "return of.apply(this,arguments).then(" & "function(v){d();return v;},function(e){d();throw e;});};"
    実行スクリプト = 実行スクリプト & "window.fetch=q.fw;}"
    
    実行スクリプト = 実行スクリプト & "if(XP){"
    実行スクリプト = 実行スクリプト & "var oo=(prev&&prev.oo&&prev.xo===XP.open)?prev.oo:XP.open;"
    実行スクリプト = 実行スクリプト & "var os=(prev&&prev.os&&prev.xs===XP.send)?prev.os:XP.send;q.oo=oo;q.os=os;"
    実行スクリプト = 実行スクリプト & "q.xo=function(m,u){try{this.__wv2u=String(u);}catch(e){}" & "return oo.apply(this,arguments);};"
    実行スクリプト = 実行スクリプト & "q.xs=function(){var s=this;try{var u=s.__wv2u||'';" & "var sp=ig(u);if(sp){q.skipNet++;}else{q.netN++;q.xhrN++;}" & "var en=sp?null:ifA(u,'xhr');" & "s.addEventListener('loadend',function(){" & _
        "if(!s.__wv2d){s.__wv2d=true;" & "if(!sp){q.xhrN--;q.lastNet=N();ifD(en);}sn(u);}});}catch(e){}" & "return os.apply(this,arguments);};"
    実行スクリプト = 実行スクリプト & "XP.open=q.xo;XP.send=q.xs;}"
    
    実行スクリプト = 実行スクリプト & "try{q.ob=new MutationObserver(function(ms){" & "var i,c=0,t;for(i=0;i<ms.length;i++){t=ms[i].target;" & "if(ms[i].type==='characterData'){t=(t&&t.parentElement)?t.parentElement:t;}" & _
        "if(sk(t)){q.skipMut++;}else{c++;}}" & "if(c>0){q.mutN+=c;q.lastMut=N();lg('DOM','x'+c);}});"
    実行スクリプト = 実行スクリプト & "q.ob.observe(document," & "{childList:true,characterData:true,subtree:true,attributes:true});" & "}catch(e){q.ob=null;}"
    
    実行スクリプト = 実行スクリプト & "window.__wv2p=q;P=q;}"
    
    SPAの通信とDOM更新を監視するJavaScriptを生成 = 実行スクリプト
End Function

' =========================================================================
' 処理概要       : SPAの通信とDOM更新が落ち着くまで待機
' 引数           : 待機上限秒 (Single) - 待機の上限秒数（省略時: 既定要素待機秒）
'                  安定ミリ秒 (Long) - 安定状態の継続時間（ミリ秒）（省略時: 既定安定待機ミリ秒）
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function SPAの通信とDOM更新が落ち着くまで待機(Optional ByVal 待機上限秒 As Single = 既定要素待機秒, Optional ByVal 安定ミリ秒 As Long = 既定安定待機ミリ秒) As Boolean
    Dim 待機開始時刻 As Single
    Dim 作業情報 As String
    Dim 処理中通信 As Long
    Dim 静止ミリ秒 As Long
    Dim 静止有効値 As Long
    Dim 経過ミリ秒 As Long
    Dim 状態確認回数 As Long
    Dim シグナル成功 As Boolean
    Dim 処理中通信一覧 As String
    Dim 参照失効 As Long
    
    SPAの通信とDOM更新が落ち着くまで待機 = False
    保持最終安定待機 = ""
    If 待機上限秒 <= 0 Then 待機上限秒 = 既定要素待機秒
    If 安定ミリ秒 < 0 Then 安定ミリ秒 = 既定安定待機ミリ秒
    
    待機開始時刻 = Timer
    Do
        状態確認回数 = 状態確認回数 + 1
        作業情報 = SPAの通信とDOM更新の監視状態を取得()
        
        If Not 保持最終スクリプト評価成功 Then
        
            保持最終安定待機 = "error " & 保持最終スクリプト評価エラー
            mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.WaitSettled: 失敗 err=" & 保持最終スクリプト評価エラー
            保持シグナルDOM = False
            保持シグナル通信 = False
            保持シグナルURL = False
            Exit Function
        End If
        
        処理中通信 = mod_WV2_JSON共通.JSONから指定キーの数値を取得(作業情報, "ifn")
        参照失効 = mod_WV2_JSON共通.JSONから指定キーの数値を取得(作業情報, "ifo")
        静止ミリ秒 = mod_WV2_JSON共通.JSONから指定キーの数値を取得(作業情報, "q")
        
        If 保持シグナルDOM Or 保持シグナル通信 Or 保持シグナルURL Then
            If (保持シグナルDOM And InStr(1, 作業情報, """sda"":true") = 0) Or (保持シグナル通信 And InStr(1, 作業情報, """sna"":true") = 0) Or (保持シグナルURL And InStr(1, 作業情報, """sua"":true") = 0) Then
                保持最終安定待機 = "signal-lost (遷移などで arm が消えた)"
                mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.WaitSettled: signal-lost " & "arm したシグナルがページ側から消えている。" & "遷移をまたいだ可能性がある (arm し直すこと)"
                登録済みの待機シグナルを解除
                Exit Function
            End If
        End If
        
        シグナル成功 = True
        If 保持シグナルDOM Then
            If mod_WV2_JSON共通.JSONから指定キーの数値を取得(作業情報, "sdh") = 0 Then シグナル成功 = False
        End If
        If 保持シグナル通信 Then
            If mod_WV2_JSON共通.JSONから指定キーの数値を取得(作業情報, "snh") = 0 Then シグナル成功 = False
        End If
        If 保持シグナルURL Then
            If mod_WV2_JSON共通.JSONから指定キーの数値を取得(作業情報, "suh") = 0 Then シグナル成功 = False
        End If
        
        経過ミリ秒 = CLng(日付の切り替わりを考慮して経過秒数を取得(待機開始時刻) * 1000)
        静止有効値 = 静止ミリ秒
        If 静止有効値 > 経過ミリ秒 Then 静止有効値 = 経過ミリ秒
        
        If 処理中通信 = 0 And 静止有効値 >= 安定ミリ秒 And シグナル成功 Then
            保持最終安定待機 = "stable q=" & 静止ミリ秒 & "ms(eff=" & 静止有効値 & ") slack=" & (静止有効値 - 安定ミリ秒) & "ms polls=" & 状態確認回数 & " elapsed=" & Format$(日付の切り替わりを考慮して経過秒数を取得(待機開始時刻), "0.00") & "s" & IIf(参照失効 > 0, " (居座り " & 参照失効 & _
                " 件は除外)", "") & シグナルの状態をJSONから取得(作業情報)
            SPAの通信とDOM更新が落ち着くまで待機 = True
            登録済みの待機シグナルを解除
            Exit Function
        End If
        
        If 日付の切り替わりを考慮して経過秒数を取得(待機開始時刻) >= 待機上限秒 Then
        
            mod_WV2_JSON共通.指定キーの文字列を復号してJSONから取得 作業情報, "il", 処理中通信一覧
            
            保持最終安定待機 = "timeout inflight=" & 処理中通信 & IIf(参照失効 > 0, "(居座り " & 参照失効 & " 件は除外)", "") & " q=" & 静止ミリ秒 & "ms (stableMs=" & 安定ミリ秒 & ") polls=" & 状態確認回数 & シグナルの状態をJSONから取得(作業情報)
            If Len(処理中通信一覧) > 0 Then
                保持最終安定待機 = 保持最終安定待機 & " | 実行中: " & 処理中通信一覧
            End If
            
            登録済みの待機シグナルを解除
            Exit Function
        End If
        
        次回の状態確認までイベント処理を進める
    Loop
End Function

' =========================================================================
' 処理概要       : 直前の安定待機処理の状態情報を取得
' 戻り値         : String - 直前の安定待機処理の状態情報
' =========================================================================
Public Property Get 直前の安定待機処理の状態情報を取得() As String
    直前の安定待機処理の状態情報を取得 = 保持最終安定待機
End Property

' =========================================================================
' 処理概要       : 要素操作後に画面の安定を待つ設定を取得
' 戻り値         : Boolean - 有効時 True / 無効時 False
' =========================================================================
Public Property Get 操作後に安定待機() As Boolean
    操作後に安定待機 = 保持自動待機
End Property

' =========================================================================
' 処理概要       : 要素操作後に画面の安定を待つ設定を更新
' 引数           : 有効無効 (Boolean) - 機能を有効にする場合 True
' =========================================================================
Public Property Let 操作後に安定待機(ByVal 有効無効 As Boolean)
    保持自動待機 = 有効無効
    
End Property

' =========================================================================
' 処理概要       : 安定判定から除外する通信URLパターンを追加
' 引数           : URLパターン (String) - 対象の通信URLパターン
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 安定判定から除外する通信URLパターンを追加(ByVal URLパターン As String) As Boolean
    Dim 処理値 As Variant
    
    If Len(URLパターン) = 0 Then Exit Function
    
    If 保持除外通信 Is Nothing Then Set 保持除外通信 = New Collection
    
    For Each 処理値 In 保持除外通信
        If CStr(処理値) = URLパターン Then
            安定判定から除外する通信URLパターンを追加 = True
            Exit Function
        End If
    Next 処理値
    
    保持除外通信.Add URLパターン
    
    安定判定から除外する通信URLパターンを追加 = True
End Function

' =========================================================================
' 処理概要       : 安定判定から除外する通信URLパターンを初期化
' =========================================================================
Public Sub 安定判定から除外する通信URLパターンを初期化()
    Set 保持除外通信 = New Collection
    
End Sub

' =========================================================================
' 処理概要       : 安定判定から除外するDOMセレクターを取得
' 戻り値         : String - 安定判定から除外するDOMセレクター
' =========================================================================
Public Property Get 除外セレクター() As String
    除外セレクター = 保持除外選択条件
End Property

' =========================================================================
' 処理概要       : 安定判定から除外するDOMセレクターを設定
' 引数           : CSSセレクター一覧 (String) - 安定判定から除外するCSSセレクター
' =========================================================================
Public Property Let 除外セレクター(ByVal CSSセレクター一覧 As String)
    保持除外選択条件 = CSSセレクター一覧
    
End Property

' =========================================================================
' 処理概要       : 監視の除外条件をJavaScriptに反映するコードを生成
' 戻り値         : String - 監視の除外条件をJavaScriptに反映するコード
' =========================================================================
Private Function 監視の除外条件をJavaScriptに反映するコードを生成() As String
    Dim 実行スクリプト As String
    Dim 参照配列 As String
    Dim 処理値 As Variant
    
    If Not 保持除外通信 Is Nothing Then
        For Each 処理値 In 保持除外通信
            If Len(参照配列) > 0 Then 参照配列 = 参照配列 & ","
            参照配列 = 参照配列 & JavaScript文字列に変換(CStr(処理値))
        Next 処理値
    End If
    
    実行スクリプト = "P.ignoreNet=[" & 参照配列 & "];P.ignoreSel=" & JavaScript文字列に変換(保持除外選択条件) & ";"
    
    監視の除外条件をJavaScriptに反映するコードを生成 = 実行スクリプト
End Function

' =========================================================================
' 処理概要       : 指定要素の更新を待つシグナルを登録
' 引数           : セレクター (String) - 検索対象のCSSセレクター
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 指定要素の更新を待つシグナルを登録(ByVal セレクター As String) As Boolean
    Dim 実行スクリプト As String
    Dim 実行結果値 As String
    
    実行スクリプト = "(function(){" & SPAの通信とDOM更新を監視するJavaScriptを生成() & 監視の除外条件をJavaScriptに反映するコードを生成() & "return {r:P.armDom(" & JavaScript文字列に変換(セレクター) & ")};})()"
    
    実行結果値 = JavaScript式を評価して完了まで待機(実行スクリプト)
    If Not 保持最終スクリプト評価成功 Then
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.ArmContentSignal: 失敗 err=" & 保持最終スクリプト評価エラー
        Exit Function
    End If
    
    If InStr(1, 実行結果値, """r"":0") = 0 Then
        If InStr(1, 実行結果値, "-2") > 0 Then
            保持最終スクリプト評価エラー = "bad-selector"
        Else
            保持最終スクリプト評価エラー = "not-found"
        End If
        保持最終スクリプト評価成功 = False
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.ArmContentSignal: arm できない (" & セレクター & ") 理由=" & 保持最終スクリプト評価エラー
        Exit Function
    End If
    
    保持シグナルDOM = True
    
    指定要素の更新を待つシグナルを登録 = True
End Function

' =========================================================================
' 処理概要       : 指定URLの通信を待つシグナルを登録
' 引数           : URLパターン (String) - 対象の通信URLパターン
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 指定URLの通信を待つシグナルを登録(ByVal URLパターン As String) As Boolean
    Dim 実行スクリプト As String
    
    If Len(URLパターン) = 0 Then Exit Function
    
    実行スクリプト = "(function(){" & SPAの通信とDOM更新を監視するJavaScriptを生成() & 監視の除外条件をJavaScriptに反映するコードを生成() & "return {r:P.armNet(" & JavaScript文字列に変換(URLパターン) & ")};})()"
    
    JavaScript式を評価して完了まで待機 実行スクリプト
    If Not 保持最終スクリプト評価成功 Then
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.ArmNetworkSignal: 失敗 err=" & 保持最終スクリプト評価エラー
        Exit Function
    End If
    
    保持シグナル通信 = True
    
    指定URLの通信を待つシグナルを登録 = True
End Function

' =========================================================================
' 処理概要       : 登録済みの待機シグナルを解除
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 登録済みの待機シグナルを解除() As Boolean
    保持シグナルDOM = False
    保持シグナル通信 = False
    保持シグナルURL = False
    
    JavaScript式を評価して完了まで待機 "(function(){var p=window.__wv2p;" & "if(p&&p.disarm){p.disarm();}return 1;})()"
    
    登録済みの待機シグナルを解除 = 保持最終スクリプト評価成功
End Function

' =========================================================================
' 処理概要       : シグナルの状態をJSONから取得
' 引数           : 作業情報 (String) - JSON内の状態文字列
' 戻り値         : String - シグナルの状態をJSONから取得
' =========================================================================
Private Function シグナルの状態をJSONから取得(ByVal 作業情報 As String) As String
    Dim 行番号 As String
    
    If 保持シグナルDOM Then
        行番号 = 行番号 & " dom:"
        If mod_WV2_JSON共通.JSONから指定キーの数値を取得(作業情報, "sdh") > 0 Then 行番号 = 行番号 & "hit" Else 行番号 = 行番号 & "miss"
    End If
    
    If 保持シグナル通信 Then
        行番号 = 行番号 & " net:"
        If mod_WV2_JSON共通.JSONから指定キーの数値を取得(作業情報, "snh") > 0 Then 行番号 = 行番号 & "hit" Else 行番号 = 行番号 & "miss"
    End If
    
    If 保持シグナルURL Then
        行番号 = 行番号 & " url:"
        If mod_WV2_JSON共通.JSONから指定キーの数値を取得(作業情報, "suh") > 0 Then 行番号 = 行番号 & "hit" Else 行番号 = 行番号 & "miss"
    End If
    
    If Len(行番号) > 0 Then 行番号 = " signal=" & Trim$(行番号)
    シグナルの状態をJSONから取得 = 行番号
End Function

' =========================================================================
' 処理概要       : URL変更を待つシグナルを登録
' 引数           : URL部分文字列 (String) - 待機対象のURL部分文字列
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function URL変更を待つシグナルを登録(ByVal URL部分文字列 As String) As Boolean
    Dim 実行スクリプト As String
    
    If Len(URL部分文字列) = 0 Then Exit Function
    
    実行スクリプト = "(function(){" & SPAの通信とDOM更新を監視するJavaScriptを生成() & 監視の除外条件をJavaScriptに反映するコードを生成() & "return {r:P.armUrl(" & JavaScript文字列に変換(URL部分文字列) & ")};})()"
    
    JavaScript式を評価して完了まで待機 実行スクリプト
    If Not 保持最終スクリプト評価成功 Then
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.ArmUrlSignal: 失敗 err=" & 保持最終スクリプト評価エラー
        Exit Function
    End If
    
    保持シグナルURL = True
    
    URL変更を待つシグナルを登録 = True
End Function

' =========================================================================
' 処理概要       : SPA監視の履歴記録設定を取得
' 戻り値         : Boolean - 有効時 True / 無効時 False
' =========================================================================
Public Property Get 画面監視履歴記録() As Boolean
    Dim 作業情報 As String
    
    作業情報 = SPAの通信とDOM更新の監視状態を取得()
    If Not 保持最終スクリプト評価成功 Then Exit Property
    
    画面監視履歴記録 = (InStr(1, 作業情報, """lgOn"":true") > 0)
End Property

' =========================================================================
' 処理概要       : SPA監視の履歴記録設定を更新
' 引数           : 有効無効 (Boolean) - 機能を有効にする場合 True
' =========================================================================
Public Property Let 画面監視履歴記録(ByVal 有効無効 As Boolean)
    Dim 実行スクリプト As String
    
    実行スクリプト = "(function(){" & SPAの通信とDOM更新を監視するJavaScriptを生成() & "P.logOn=" & IIf(有効無効, "true", "false") & ";" & "if(!P.logOn){P.log=[];}return {r:1};})()"
    
    JavaScript式を評価して完了まで待機 実行スクリプト
    
End Property

' =========================================================================
' 処理概要       : SPA監視の履歴をログへ出力して消去
' 戻り値         : Long - ログへ出力した件数
' =========================================================================
Public Function SPA監視の履歴をログへ出力して消去() As Long
    Dim 実行スクリプト As String
    Dim 実行結果値 As String
    Dim 本文文字列 As String
    Dim 行一覧 As Variant
    Dim 添字 As Long
    
    実行スクリプト = "(function(){" & SPAの通信とDOM更新を監視するJavaScriptを生成() & "var a=P.log.splice(0,60);" & "return {n:a.length,s:a.join(String.fromCharCode(1))};})()"
    
    実行結果値 = JavaScript式を評価して完了まで待機(実行スクリプト, 10)
    If Not 保持最終スクリプト評価成功 Then
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.SpaProbeDrainLog: 失敗 err=" & 保持最終スクリプト評価エラー
        Exit Function
    End If
    
    SPA監視の履歴をログへ出力して消去 = mod_WV2_JSON共通.JSONから指定キーの数値を取得(実行結果値, "n")
    If SPA監視の履歴をログへ出力して消去 = 0 Then Exit Function
    
    mod_WV2_JSON共通.指定キーの文字列を復号してJSONから取得 実行結果値, "s", 本文文字列
    行一覧 = Split(本文文字列, ChrW$(1))
    
    mod_WV2_ログ管理.情報レベルのログを記録 "  --- 診断ログ (" & SPA監視の履歴をログへ出力して消去 & " 件) ---"
    For 添字 = LBound(行一覧) To UBound(行一覧)
        mod_WV2_ログ管理.情報レベルのログを記録 "    " & 行一覧(添字)
    Next 添字
End Function

' =========================================================================
' 処理概要       : 長時間継続する通信を安定判定から外す閾値を取得
' 戻り値         : Long - 長時間継続する通信を安定判定から外す閾値
' =========================================================================
Public Property Get 長時間通信の除外閾値() As Long
    長時間通信の除外閾値 = 保持参照失効ミリ秒
End Property

' =========================================================================
' 処理概要       : 長時間継続する通信を安定判定から外す閾値を設定
' 引数           : 経過ミリ秒 (Long) - 通信記録開始からの経過ミリ秒
' =========================================================================
Public Property Let 長時間通信の除外閾値(ByVal 経過ミリ秒 As Long)
    If 経過ミリ秒 < 0 Then 経過ミリ秒 = 0
    保持参照失効ミリ秒 = 経過ミリ秒
    
End Property

' =========================================================================
' 処理概要       : 通信監視の有効期限設定をJavaScriptへ変換
' 戻り値         : String - 通信監視の有効期限設定をJavaScript
' =========================================================================
Private Function 通信監視の有効期限設定をJavaScriptへ変換() As String
    If 保持参照失効ミリ秒 <= 0 Then
        通信監視の有効期限設定をJavaScriptへ変換 = "1e15"
    Else
        通信監視の有効期限設定をJavaScriptへ変換 = CStr(保持参照失効ミリ秒)
    End If
End Function

' =========================================================================
' 処理概要       : CSSセレクターに一致するDOM要素を上限件数まで取得
' 引数           : セレクター (String) - 検索対象のCSSセレクター
'                  最大件数 (Long) - 一度に取得または出力する最大件数（省略時: 既定要素取得上限）
' 戻り値         : Collection - 取得した要素の一覧
' =========================================================================
Public Function CSSセレクターに一致するDOM要素を上限件数まで取得(ByVal セレクター As String, Optional ByVal 最大件数 As Long = 既定要素取得上限) As Collection
    Dim 実行スクリプト As String
    Dim 実行結果値 As String
    Dim 世代 As String
    Dim 要素番号一覧 As String
    Dim 分割値一覧 As Variant
    Dim 添字 As Long
    Dim 対象要素 As mod_WV2_HTML要素
    Dim 出力列 As Collection
    
    Set 出力列 = New Collection
    Set CSSセレクターに一致するDOM要素を上限件数まで取得 = 出力列
    保持全件打切り済み = False
    
    If 最大件数 <= 0 Then 最大件数 = 既定要素取得上限
    
    実行スクリプト = "(function(){" & DOM要素レジストリを初期化するJavaScriptを生成() & "var q=document.querySelectorAll(" & JavaScript文字列に変換(セレクター) & ");" & "var mx=" & CStr(最大件数) & ",h=[],i;" & _
        "for(i=0;i<q.length&&i<mx;i++){R.els.push(q[i]);h.push(R.els.length-1);}" & "return {g:R.gen,n:q.length,k:h.length,cut:(q.length>mx)," & "h:h.join(','),tot:R.els.length};})()"
        
    実行結果値 = JavaScript式を評価して完了まで待機(実行スクリプト)
    
    If Not 保持最終スクリプト評価成功 Then
    
        Exit Function
    End If
    
    If mod_WV2_JSON共通.指定キーの文字列を復号してJSONから取得(実行結果値, "g", 世代) Then
        保持DOM世代 = 世代
    End If
    
    保持全件打切り済み = (InStr(1, 実行結果値, """cut"":true") > 0)
    If 保持全件打切り済み Then
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.QuerySelectorAll: 打ち切った (" & セレクター & ") ページには " & mod_WV2_JSON共通.JSONから指定キーの数値を取得(実行結果値, "n") & " 件あるが " & 最大件数 & " 件で止めた"
    End If
    
    mod_WV2_JSON共通.指定キーの文字列を復号してJSONから取得 実行結果値, "h", 要素番号一覧
    If Len(要素番号一覧) = 0 Then
    
        DOM要素の登録件数が閾値を超えた場合に警告を記録 mod_WV2_JSON共通.JSONから指定キーの数値を取得(実行結果値, "tot")
        Exit Function
    End If
    
    分割値一覧 = Split(要素番号一覧, ",")
    For 添字 = LBound(分割値一覧) To UBound(分割値一覧)
    
        ' 取得した要素をVBAから操作できるよう、参照を作成する
        Set 対象要素 = New mod_WV2_HTML要素
        対象要素.DOM要素の参照情報を設定 Me, CLng(分割値一覧(添字)), 世代
        出力列.Add 対象要素
    Next 添字
    
    DOM要素の登録件数が閾値を超えた場合に警告を記録 mod_WV2_JSON共通.JSONから指定キーの数値を取得(実行結果値, "tot")
    
End Function

' =========================================================================
' 処理概要       : 直前の複数要素取得が上限で打ち切られたか取得
' 戻り値         : Boolean - 取得件数の上限に達した場合 True
' =========================================================================
Public Property Get 直前の複数要素取得が上限で打ち切られたか取得() As Boolean
    直前の複数要素取得が上限で打ち切られたか取得 = 保持全件打切り済み
End Property

' =========================================================================
' 処理概要       : DOM要素レジストリの登録件数を取得
' 戻り値         : Long - DOM要素レジストリの登録件数
' =========================================================================
Public Property Get DOM要素レジストリの登録件数を取得() As Long
    Dim 実行結果値 As String
    
    実行結果値 = JavaScript式を評価して完了まで待機("(function(){" & DOM要素レジストリを初期化するJavaScriptを生成() & "return {n:R.els.length};})()")
    
    If Not 保持最終スクリプト評価成功 Then
        DOM要素レジストリの登録件数を取得 = -1
        Exit Property
    End If
    
    DOM要素レジストリの登録件数を取得 = mod_WV2_JSON共通.JSONから指定キーの数値を取得(実行結果値, "n")
End Property

' =========================================================================
' 処理概要       : DOM要素の登録件数が閾値を超えた場合に警告を記録
' 引数           : 総数 (Long) - 記録件数
' =========================================================================
Private Sub DOM要素の登録件数が閾値を超えた場合に警告を記録(ByVal 総数 As Long)
    If 総数 < 要素数警告閾値 Then Exit Sub
    
    mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示: 要素レジストリに " & 総数 & " 個積まれている " & "(掴んだ要素は解放しない設計)。要らなければ ClearElementRegistry で" & "作り直せる。ただし手元の HTML要素 はすべて stale になる。"
End Sub

' =========================================================================
' 処理概要       : 指定条件の通信記録を開始
' 引数           : 対象リソース種別 (String) - 対象の通信リソース種別一覧（省略時: ""）
'                  URIフィルター (String) - 詳細記録の対象URI条件（省略時: "*"）
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 指定条件の通信記録を開始(Optional ByVal 対象リソース種別 As String = "", Optional ByVal URIフィルター As String = "*") As Boolean
    Dim 分割値一覧 As Variant
    Dim 添字     As Long
    Dim リソース種別   As Long
    Dim 追加件数 As Long
    Dim 実行結果    As Long
    Dim 対象条件  As String
    
    If 保持終了済み Then
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.NetCaptureStart: このペインは既に閉じている"
        Exit Function
    End If
    If 保持ポインターWebView = 0 Then
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.NetCaptureStart: m_pView が未取得 (先にタブを作ること)"
        Exit Function
    End If
    
    If 保持通信有効 Then 通信イベントを解除して記録を停止
    
    対象条件 = Trim$(対象リソース種別)
    If LenB(対象条件) = 0 Then 対象条件 = 通信リソース種別既定
    If LenB(Trim$(URIフィルター)) = 0 Then URIフィルター = "*"
    
    実行結果 = 通信要求イベントを登録()
    If 実行結果 <> 0 Then
        mod_WV2_ログ管理.エラーレベルのログを記録 "ブラウザー表示.NetCaptureStart: add_WebResourceRequested 失敗 hr=&H" & Hex(実行結果)
        Exit Function
    End If
    
    分割値一覧 = Split(対象条件, ",")
    For 添字 = LBound(分割値一覧) To UBound(分割値一覧)
        リソース種別 = 通信リソース種別の名前を数値へ変換(CStr(分割値一覧(添字)))
        If リソース種別 < 0 Then
            mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.NetCaptureStart: 種別 [" & Trim$(CStr(分割値一覧(添字))) & "] が分からないので飛ばす"
        Else
            実行結果 = 通信要求の対象フィルターを追加(URIフィルター, リソース種別)
            If 実行結果 <> 0 Then
                mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.NetCaptureStart: フィルタ登録に失敗 " & 通信リソース種別を表示名へ変換(リソース種別) & " hr=&H" & Hex(実行結果)
            Else
                保持通信フィルター一覧.Add URIフィルター & vbTab & リソース種別
                追加件数 = 追加件数 + 1
            End If
        End If
    Next 添字
    
    If 追加件数 = 0 Then
        mod_WV2_ログ管理.エラーレベルのログを記録 "ブラウザー表示.NetCaptureStart: フィルタが 1 つも張れなかった " & "この状態ではイベントが 1 件も来ないので、ハンドラも外して中止する。"
        通信要求イベントを解除
        Exit Function
    End If
    
    保持通信応答有効 = False
    実行結果 = 通信応答イベントを登録()
    If 実行結果 = 0 Then
        保持通信応答有効 = True
    Else
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.NetCaptureStart: 応答イベントを張れなかった hr=&H" & Hex(実行結果) & " 一覧のステータス欄は空欄のままになる。"
    End If
    
    通信一覧の記録を消去
    保持通信有効 = True
    指定条件の通信記録を開始 = True
    
    mod_WV2_ログ管理.情報レベルのログを記録 "ブラウザー表示.NetCaptureStart: 捕捉開始 uri=" & URIフィルター & " 種別=" & 対象条件 & " (フィルタ " & 追加件数 & " 本)"
End Function

' =========================================================================
' 処理概要       : 通信イベントを解除して記録を停止
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 通信イベントを解除して記録を停止() As Boolean
    Dim 処理値     As Variant
    Dim 分割値一覧 As Variant
    Dim 実行結果    As Long
    Dim 全件成功 As Boolean
    
    If Not 保持通信有効 Then
        通信イベントを解除して記録を停止 = True
        Exit Function
    End If
    
    全件成功 = True
    
    If Not (保持終了済み Or 保持ポインターWebView = 0) Then
        For Each 処理値 In 保持通信フィルター一覧
            分割値一覧 = Split(CStr(処理値), vbTab)
            実行結果 = 通信要求の対象フィルターを解除(CStr(分割値一覧(0)), CLng(分割値一覧(1)))
            If 実行結果 <> 0 Then
                mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.NetCaptureStop: フィルタ解除に失敗 " & CStr(分割値一覧(0)) & " / " & 通信リソース種別を表示名へ変換(CLng(分割値一覧(1))) & " hr=&H" & Hex(実行結果)
                全件成功 = False
            End If
        Next 処理値
    End If
    Set 保持通信フィルター一覧 = New Collection
    
    実行結果 = 通信要求イベントを解除()
    If 実行結果 <> 0 Then 全件成功 = False
    
    実行結果 = 通信応答イベントを解除()
    If 実行結果 <> 0 Then 全件成功 = False
    保持通信応答有効 = False
    
    応答本文の取得待ちハンドラーを解放
    
    保持通信有効 = False
    通信イベントを解除して記録を停止 = 全件成功
    
    mod_WV2_ログ管理.情報レベルのログを記録 "ブラウザー表示.NetCaptureStop: 記録停止 (累計記録 " & 保持通信総数 & " 件 / " & "保持 " & 保持通信件数 & " 件 / 破棄 " & 保持通信破棄件数 & " 件)"
End Function

' =========================================================================
' 処理概要       : 通信記録が有効か取得
' 戻り値         : Boolean - 有効時 True / 無効時 False
' =========================================================================
Public Property Get 通信記録が有効か取得() As Boolean
    通信記録が有効か取得 = 保持通信有効
End Property

' =========================================================================
' 処理概要       : 通信一覧に保持している件数を取得
' 戻り値         : Long - 通信一覧に保持している件数
' =========================================================================
Public Property Get 通信一覧に保持している件数を取得() As Long
    通信一覧に保持している件数を取得 = 保持通信件数
End Property

' =========================================================================
' 処理概要       : 通信一覧の累計記録件数を取得
' 戻り値         : Long - 通信一覧の累計記録件数
' =========================================================================
Public Property Get 通信一覧の累計記録件数を取得() As Long
    通信一覧の累計記録件数を取得 = 保持通信総数
End Property

' =========================================================================
' 処理概要       : 通信一覧の上限超過で破棄した件数を取得
' 戻り値         : Long - 通信一覧の上限超過で破棄した件数
' =========================================================================
Public Property Get 通信一覧の上限超過で破棄した件数を取得() As Long
    通信一覧の上限超過で破棄した件数を取得 = 保持通信破棄件数
End Property

' =========================================================================
' 処理概要       : 通信一覧の最大保持件数を取得
' 戻り値         : Long - 通信一覧の最大保持件数
' =========================================================================
Public Property Get 通信一覧の最大保持件数を取得() As Long
    通信一覧の最大保持件数を取得 = 保持通信保持上限
End Property

' =========================================================================
' 処理概要       : 通信一覧の最大保持件数を設定
' 引数           : 件数 (Long) - 処理対象の件数
' =========================================================================
Public Property Let 通信一覧の最大保持件数を取得(ByVal 件数 As Long)
    If 件数 < 1 Then 件数 = 1
    If 件数 > 通信保持上限最大 Then
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.NetLogCapacity: " & 件数 & " は上限 " & 通信保持上限最大 & " を超えたため上限値を使用"
        件数 = 通信保持上限最大
    End If
    保持通信保持上限 = 件数
    通信一覧のリングバッファーを確保
    
End Property

' =========================================================================
' 処理概要       : 通信一覧の記録を消去
' =========================================================================
Public Sub 通信一覧の記録を消去()
    通信一覧のリングバッファーを確保
End Sub

' =========================================================================
' 処理概要       : 指定位置の通信一覧を文字列で取得
' 引数           : 対象添字 (Long) - 通信記録の位置（1始まり）
' 戻り値         : String - 指定位置の通信一覧を文字列で取得
' =========================================================================
Public Property Get 指定位置の通信一覧を文字列で取得(ByVal 対象添字 As Long) As String
    Dim 参照キー As Long
    
    参照キー = 通信一覧の表示位置を配列添字に変換(対象添字)
    If 参照キー < 0 Then Exit Property
    
    指定位置の通信一覧を文字列で取得 = 保持通信連番(参照キー) & vbTab & 保持通信ミリ秒(参照キー) & vbTab & 保持通信メソッド(参照キー) & vbTab & 通信リソース種別を表示名へ変換(保持通信リソース種別(参照キー)) & vbTab & 保持通信URI(参照キー) & vbTab & 通信状態とHTTPステータスを表示用文字列へ変換(保持通信状態(参照キー)) & vbTab & _
        保持通信理由(参照キー)
End Property

' =========================================================================
' 処理概要       : 通信一覧をログへ出力して消去
' 戻り値         : Long - ログへ出力した件数
' =========================================================================
Public Function 通信一覧をログへ出力して消去() As Long
    Dim 添字 As Long
    Dim 件数 As Long
    
    件数 = 保持通信件数
    If 件数 = 0 Then
        mod_WV2_ログ管理.情報レベルのログを記録 "  --- ネットワーク要求: 保持データなし (累計記録 " & 保持通信総数 & " 件) ---"
        Exit Function
    End If
    
    mod_WV2_ログ管理.情報レベルのログを記録 "  --- ネットワーク要求 " & 件数 & " 件 (累計記録 " & 保持通信総数 & " 件) ---"
    mod_WV2_ログ管理.情報レベルのログを記録 "        #   経過ms  メソッド 種別       状態  URL"
    For 添字 = 1 To 件数
        mod_WV2_ログ管理.情報レベルのログを記録 "      " & 通信一覧の指定行を文字列に整形(添字)
    Next 添字
    
    If 保持通信破棄件数 > 0 Then
    
        mod_WV2_ログ管理.警告レベルのログを記録 "  捕捉開始から " & 保持通信破棄件数 & " 件が溢れて捨てられている " & "NetLogCapacity を増やすか、こまめに NetLogDrain すること。"
    End If
    
    If 保持通信未対応件数 > 0 Then
        mod_WV2_ログ管理.情報レベルのログを記録 "  (一覧に対応が無かった応答 " & 保持通信未対応件数 & " 件 ― " & "画像や CSS など捕捉対象外のものが大半で、これは正常)"
    End If
    
    通信一覧をログへ出力して消去 = 件数
    
    通信一覧のリングバッファーを初期化
End Function

' =========================================================================
' 処理概要       : 通信要求イベントを登録
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function 通信要求イベントを登録() As Long
    Dim 通知ハンドラー As mod_WV2_COM通知ハンドラー
    Dim 仮想COMオブジェクト As LongPtr
    Dim 登録トークン As LongLong
    Dim 実行結果 As Long
    
    If 保持終了済み Then
        通信要求イベントを登録 = &H80004003
        Exit Function
    End If
    If 保持ポインターWebView = 0 Then
    
        通信要求イベントを登録 = &H80004003
        Exit Function
    End If
    
    If Not (保持永続ハンドラー一覧(通信要求通知) Is Nothing) Then
    
        通信要求イベントを登録 = 0
        Exit Function
    End If
    
    Set 通知ハンドラー = 通知先を設定したCOMコールバックハンドラーを取得(通信要求通知, Me)
    If 通知ハンドラー Is Nothing Then
    
        通信要求イベントを登録 = &H8000FFFF
        Exit Function
    End If
    
    仮想COMオブジェクト = 通知ハンドラー.スロットアドレス + 仮想COMオブジェクト位置
    
    実行結果 = COMメソッド呼出(保持ポインターWebView, 55, "View.add_WebResourceRequested", 仮想COMオブジェクト, VarPtr(登録トークン))
    
    If 実行結果 <> 0 Then
    
        Set 通知ハンドラー = Nothing
        通信要求イベントを登録 = 実行結果
        Exit Function
    End If
    
    保持永続登録トークン一覧(通信要求通知) = 登録トークン
    Set 保持永続ハンドラー一覧(通信要求通知) = 通知ハンドラー
    
    通信要求イベントを登録 = 0
End Function

' =========================================================================
' 処理概要       : 通信要求イベントを解除
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function 通信要求イベントを解除() As Long
    Dim 実行結果 As Long
    
    If 保持永続ハンドラー一覧(通信要求通知) Is Nothing Then
        通信要求イベントを解除 = 0
        Exit Function
    End If
    
    If 保持ポインターWebView <> 0 Then
        実行結果 = COMメソッド呼出(保持ポインターWebView, 56, "View.remove_WebResourceRequested", 保持永続登録トークン一覧(通信要求通知))
    End If
    
    Set 保持永続ハンドラー一覧(通信要求通知) = Nothing
    保持永続登録トークン一覧(通信要求通知) = 0^
    
    通信要求イベントを解除 = 実行結果
End Function

' =========================================================================
' 処理概要       : 通信要求の対象フィルターを追加
' 引数           : URIフィルター (String) - 詳細記録の対象URI条件
'                  リソース種別番号 (Long) - 通信リソースの種別番号
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function 通信要求の対象フィルターを追加(ByVal URIフィルター As String, ByVal リソース種別番号 As Long) As Long
    If 保持ポインターWebView = 0 Then
        通信要求の対象フィルターを追加 = &H80004003
        Exit Function
    End If
    通信要求の対象フィルターを追加 = COMメソッド呼出(保持ポインターWebView, 57, "View.AddWebResourceRequestedFilter", StrPtr(URIフィルター), リソース種別番号)
End Function

' =========================================================================
' 処理概要       : 通信要求の対象フィルターを解除
' 引数           : URIフィルター (String) - 詳細記録の対象URI条件
'                  リソース種別番号 (Long) - 通信リソースの種別番号
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function 通信要求の対象フィルターを解除(ByVal URIフィルター As String, ByVal リソース種別番号 As Long) As Long
    If 保持ポインターWebView = 0 Then
        通信要求の対象フィルターを解除 = &H80004003
        Exit Function
    End If
    通信要求の対象フィルターを解除 = COMメソッド呼出(保持ポインターWebView, 58, "View.RemoveWebResourceRequestedFilter", StrPtr(URIフィルター), リソース種別番号)
End Function

' =========================================================================
' 処理概要       : 通信要求イベントを受け付けてコールバック状態を管理
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
'                  ポインター引数一覧 (LongPtr) - イベント引数のCOMポインター
' =========================================================================
Public Sub 通信要求イベントを受け付けてコールバック状態を管理(ByVal ポインター通知元 As LongPtr, ByVal ポインター引数一覧 As LongPtr)
    Dim エラー番号 As Long
    Dim エラー説明 As String
    
    ' コールバック中であることを記録し、同期処理の再入を防ぐ
    保持内部通知 = 保持内部通知 + 1
    On Error GoTo ErrorHanler
    通信要求イベントの内容を処理 ポインター通知元, ポインター引数一覧
    
    GoTo EndProc
    
ErrorHanler:
    エラー番号 = Err.Number
    エラー説明 = Err.Description
    MsgBox "通信イベントの処理でエラーが発生しました。" & vbCrLf & エラー説明, vbExclamation, "ブラウザー"
    Resume EndProc
    
EndProc:

    ' 通知処理の終了時にコールバック階層を戻す
    If 保持内部通知 > 0 Then 保持内部通知 = 保持内部通知 - 1
    If エラー番号 <> 0 Then
        On Error Resume Next
        通信一覧のリングバッファーに要求を追加 "(err)", -1, "通知処理でエラー " & エラー番号 & " " & エラー説明 & ""
    End If
End Sub

' =========================================================================
' 処理概要       : 通信要求イベントの内容を処理
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
'                  ポインター引数一覧 (LongPtr) - イベント引数のCOMポインター
' =========================================================================
Private Sub 通信要求イベントの内容を処理(ByVal ポインター通知元 As LongPtr, ByVal ポインター引数一覧 As LongPtr)
    Dim ポインター要求 As LongPtr
    Dim 実行結果   As Long
    Dim URI  As String
    Dim HTTPメソッド  As String
    Dim リソース種別  As Long
    
    If ポインター引数一覧 = 0 Then
        通信一覧のリングバッファーに要求を追加 "(err)", -1, "args が 0"
        Exit Sub
    End If
    
    リソース種別 = COMメソッドから32ビット整数プロパティを取得(ポインター引数一覧, 7, "")
    
    実行結果 = COMメソッド呼出(ポインター引数一覧, 3, "", VarPtr(ポインター要求))
    If 実行結果 <> 0 Or ポインター要求 = 0 Then
        通信一覧のリングバッファーに要求を追加 "(err)", リソース種別, "get_Request 失敗 hr=&H" & Hex(実行結果) & ""
        Exit Sub
    End If
    
    URI = COMメソッドから文字列プロパティを取得(ポインター要求, 3, "")
    HTTPメソッド = COMメソッドから文字列プロパティを取得(ポインター要求, 5, "")
    
    ' 要求のメソッドとURIを通信一覧へ記録する
    通信一覧のリングバッファーに要求を追加 HTTPメソッド, リソース種別, URI
    
    If 保持詳細有効 Then 要求のヘッダーと本文を読み取り詳細に保存 ポインター要求, HTTPメソッド, URI
    
    COM参照解放 ポインター要求
End Sub

' =========================================================================
' 処理概要       : 要求のヘッダーと本文を読み取り詳細に保存
' 引数           : ポインター要求 (LongPtr) - 要求インターフェースのポインター
'                  HTTPメソッド (String) - HTTPメソッド
'                  URI (String) - 通信先URI
' =========================================================================
Private Sub 要求のヘッダーと本文を読み取り詳細に保存(ByVal ポインター要求 As LongPtr, ByVal HTTPメソッド As String, ByVal URI As String)
    Dim ヘッダー一覧  As String
    Dim ContentType値    As String
    Dim 件数バイト数   As Long
    Dim 打切り    As Boolean
    Dim 判定文字列   As Boolean
    Dim 本文文字列 As String
    Dim 数値16進先頭  As String
    Dim 作業文字列()    As Byte
    
    If LenB(保持詳細URIフィルター) > 0 Then
        If InStr(1, URI, 保持詳細URIフィルター, vbTextCompare) = 0 Then Exit Sub
    End If
    
    ヘッダー一覧 = COMインターフェースからヘッダーを取得(ポインター要求)
    ContentType値 = ヘッダー文字列から指定ヘッダーの値を取得(ヘッダー一覧, "content-type")
    
    If UCase$(HTTPメソッド) <> "GET" Then
        If 要求本文を上限バイト数まで取得(ポインター要求, 件数バイト数, 打切り, 作業文字列) Then
            If 件数バイト数 > 0 Then
                判定文字列 = ContentTypeがテキスト形式か判定(ContentType値)
                If 判定文字列 Then
                    本文文字列 = UTF8のバイト配列をVBA文字列へ変換(作業文字列, 件数バイト数)
                Else
                    数値16進先頭 = バイト配列の先頭を16進文字列へ変換(作業文字列, 件数バイト数)
                End If
            End If
        End If
    End If
    
    通信詳細をリングバッファーに追加 保持通信総数, HTTPメソッド, URI, ヘッダー一覧, ContentType値, 件数バイト数, 打切り, 判定文字列, 本文文字列, 数値16進先頭
End Sub

' =========================================================================
' 処理概要       : 通信一覧のリングバッファーを確保
' =========================================================================
Private Sub 通信一覧のリングバッファーを確保()
    If 保持通信保持上限 < 1 Then 保持通信保持上限 = 通信保持上限既定
    
    ReDim 保持通信連番(0 To 保持通信保持上限 - 1)
    ReDim 保持通信ミリ秒(0 To 保持通信保持上限 - 1)
    ReDim 保持通信リソース種別(0 To 保持通信保持上限 - 1)
    ReDim 保持通信メソッド(0 To 保持通信保持上限 - 1)
    ReDim 保持通信URI(0 To 保持通信保持上限 - 1)
    ReDim 保持通信状態(0 To 保持通信保持上限 - 1)
    ReDim 保持通信理由(0 To 保持通信保持上限 - 1)
    
    保持通信先頭 = 0
    保持通信件数 = 0
    保持通信総数 = 0
    保持通信破棄件数 = 0
    保持通信未対応件数 = 0
    保持通信計測開始0 = GetTickCount()
End Sub

' =========================================================================
' 処理概要       : 通信一覧のリングバッファーを初期化
' =========================================================================
Private Sub 通信一覧のリングバッファーを初期化()
    保持通信先頭 = 0
    保持通信件数 = 0
End Sub

' =========================================================================
' 処理概要       : 通信一覧のリングバッファーに要求を追加
' 引数           : HTTPメソッド (String) - HTTPメソッド
'                  リソース種別 (Long) - 通信リソースの種別番号
'                  URI (String) - 通信先URI
' =========================================================================
Private Sub 通信一覧のリングバッファーに要求を追加(ByVal HTTPメソッド As String, ByVal リソース種別 As Long, ByVal URI As String)
    保持通信総数 = 保持通信総数 + 1
    
    保持通信連番(保持通信先頭) = 保持通信総数
    保持通信ミリ秒(保持通信先頭) = 通信記録の開始時点からの経過ミリ秒を取得()
    保持通信リソース種別(保持通信先頭) = リソース種別
    保持通信メソッド(保持通信先頭) = HTTPメソッド
    保持通信URI(保持通信先頭) = URI
    
    保持通信状態(保持通信先頭) = 0
    保持通信理由(保持通信先頭) = ""
    
    保持通信先頭 = 保持通信先頭 + 1
    If 保持通信先頭 >= 保持通信保持上限 Then 保持通信先頭 = 0
    
    If 保持通信件数 < 保持通信保持上限 Then
        保持通信件数 = 保持通信件数 + 1
    Else
        保持通信破棄件数 = 保持通信破棄件数 + 1
    End If
End Sub

' =========================================================================
' 処理概要       : 通信記録の開始時点からの経過ミリ秒を取得
' 戻り値         : Long - 通信記録の開始時点からの経過ミリ秒
' =========================================================================
Private Function 通信記録の開始時点からの経過ミリ秒を取得() As Long
    Dim 後続文字 As Double
    
    後続文字 = CDbl(GetTickCount()) - CDbl(保持通信計測開始0)
    If 後続文字 < 0 Then 後続文字 = 後続文字 + 4294967296#
    If 後続文字 > 2147483647# Then 後続文字 = 2147483647#
    通信記録の開始時点からの経過ミリ秒を取得 = CLng(後続文字)
End Function

' =========================================================================
' 処理概要       : 通信一覧の表示位置を配列添字に変換
' 引数           : 対象添字 (Long) - 通信記録の位置（1始まり）
' 戻り値         : Long - 対応する配列添字。該当しない場合 -1
' =========================================================================
Private Function 通信一覧の表示位置を配列添字に変換(ByVal 対象添字 As Long) As Long
    Dim 参照キー As Long
    
    If 対象添字 < 1 Or 対象添字 > 保持通信件数 Then
        通信一覧の表示位置を配列添字に変換 = -1
        Exit Function
    End If
    
    参照キー = 保持通信先頭 - 保持通信件数 + (対象添字 - 1)
    Do While 参照キー < 0
        参照キー = 参照キー + 保持通信保持上限
    Loop
    通信一覧の表示位置を配列添字に変換 = 参照キー Mod 保持通信保持上限
End Function

' =========================================================================
' 処理概要       : 通信一覧の指定行を文字列に整形
' 引数           : 対象添字 (Long) - 通信記録の位置（1始まり）
' 戻り値         : String - 通信一覧の指定行を文字列に整形
' =========================================================================
Private Function 通信一覧の指定行を文字列に整形(ByVal 対象添字 As Long) As String
    Dim 参照キー As Long
    
    参照キー = 通信一覧の表示位置を配列添字に変換(対象添字)
    If 参照キー < 0 Then Exit Function
    
    通信一覧の指定行を文字列に整形 = Right$("     " & 保持通信連番(参照キー), 5) & "  " & Right$("       " & 保持通信ミリ秒(参照キー), 7) & "  " & Left$(保持通信メソッド(参照キー) & "        ", 8) & " " & Left$(通信リソース種別を表示名へ変換(保持通信リソース種別(参照キー)) & "          ", 10) & _
        " " & Right$("     " & 通信状態とHTTPステータスを表示用文字列へ変換(保持通信状態(参照キー)), 5) & " " & 保持通信URI(参照キー)
End Function

' =========================================================================
' 処理概要       : 通信状態とHTTPステータスを表示用文字列へ変換
' 引数           : 作業情報 (Long) - HTTPステータスコード
' 戻り値         : String - 通信状態とHTTPステータスを表示用文字列
' =========================================================================
Private Function 通信状態とHTTPステータスを表示用文字列へ変換(ByVal 作業情報 As Long) As String
    If 作業情報 = 0 Then Exit Function
    
    If 作業情報 < 0 Then 通信状態とHTTPステータスを表示用文字列へ変換 = "(読めず)": Exit Function
    通信状態とHTTPステータスを表示用文字列へ変換 = CStr(作業情報)
End Function

' =========================================================================
' 処理概要       : 通信リソース種別を表示名へ変換
' 引数           : リソース種別 (Long) - 通信リソースの種別番号
' 戻り値         : String - 通信リソース種別を表示名
' =========================================================================
Private Function 通信リソース種別を表示名へ変換(ByVal リソース種別 As Long) As String
    Select Case リソース種別
        Case 0:  通信リソース種別を表示名へ変換 = "ALL"
        Case 1:  通信リソース種別を表示名へ変換 = "DOCUMENT"
        Case 2:  通信リソース種別を表示名へ変換 = "STYLESHEET"
        Case 3:  通信リソース種別を表示名へ変換 = "IMAGE"
        Case 4:  通信リソース種別を表示名へ変換 = "MEDIA"
        Case 5:  通信リソース種別を表示名へ変換 = "FONT"
        Case 6:  通信リソース種別を表示名へ変換 = "SCRIPT"
        Case 7:  通信リソース種別を表示名へ変換 = "XHR"
        Case 8:  通信リソース種別を表示名へ変換 = "FETCH"
        Case 9:  通信リソース種別を表示名へ変換 = "TEXTTRACK"
        Case 10: 通信リソース種別を表示名へ変換 = "EVENTSRC"
        Case 11: 通信リソース種別を表示名へ変換 = "WEBSOCKET"
        Case 12: 通信リソース種別を表示名へ変換 = "MANIFEST"
        Case 13: 通信リソース種別を表示名へ変換 = "SIGNEDEX"
        Case 14: 通信リソース種別を表示名へ変換 = "PING"
        Case 15: 通信リソース種別を表示名へ変換 = "CSPREPORT"
        Case 16: 通信リソース種別を表示名へ変換 = "OTHER"
        Case Else: 通信リソース種別を表示名へ変換 = "CTX" & リソース種別
    End Select
End Function

' =========================================================================
' 処理概要       : 通信リソース種別の名前を数値へ変換
' 引数           : リソース種別登録トークン (String) - 通信リソースの種別名
' 戻り値         : Long - 通信リソース種別の番号。解決できない場合 -1
' =========================================================================
Private Function 通信リソース種別の名前を数値へ変換(ByVal リソース種別登録トークン As String) As Long
    Dim 開始時刻 As String
    
    開始時刻 = UCase$(Trim$(リソース種別登録トークン))
    開始時刻 = Replace$(開始時刻, "_", "")
    開始時刻 = Replace$(開始時刻, " ", "")
    
    If LenB(開始時刻) = 0 Then
        通信リソース種別の名前を数値へ変換 = -1
        Exit Function
    End If
    
    If IsNumeric(開始時刻) Then
        通信リソース種別の名前を数値へ変換 = CLng(開始時刻)
        Exit Function
    End If
    
    Select Case 開始時刻
        Case "ALL":                             通信リソース種別の名前を数値へ変換 = 0
        Case "DOCUMENT", "DOC":                 通信リソース種別の名前を数値へ変換 = 1
        Case "STYLESHEET", "CSS":               通信リソース種別の名前を数値へ変換 = 2
        Case "IMAGE", "IMG":                    通信リソース種別の名前を数値へ変換 = 3
        Case "MEDIA":                           通信リソース種別の名前を数値へ変換 = 4
        Case "FONT":                            通信リソース種別の名前を数値へ変換 = 5
        Case "SCRIPT", "JS":                    通信リソース種別の名前を数値へ変換 = 6
        Case "XHR", "XMLHTTPREQUEST":           通信リソース種別の名前を数値へ変換 = 7
        Case "FETCH":                           通信リソース種別の名前を数値へ変換 = 8
        Case "TEXTTRACK":                       通信リソース種別の名前を数値へ変換 = 9
        Case "EVENTSOURCE", "EVENTSRC":         通信リソース種別の名前を数値へ変換 = 10
        Case "WEBSOCKET", "WS":                 通信リソース種別の名前を数値へ変換 = 11
        Case "MANIFEST":                        通信リソース種別の名前を数値へ変換 = 12
        Case "SIGNEDEXCHANGE", "SIGNEDEX":      通信リソース種別の名前を数値へ変換 = 13
        Case "PING":                            通信リソース種別の名前を数値へ変換 = 14
        Case "CSPVIOLATIONREPORT", "CSPREPORT": 通信リソース種別の名前を数値へ変換 = 15
        Case "OTHER":                           通信リソース種別の名前を数値へ変換 = 16
        Case Else:                              通信リソース種別の名前を数値へ変換 = -1
    End Select
End Function

' =========================================================================
' 処理概要       : 要求の詳細記録設定を取得
' 戻り値         : Boolean - 有効時 True / 無効時 False
' =========================================================================
Public Property Get 要求の詳細記録設定を取得() As Boolean
    要求の詳細記録設定を取得 = 保持詳細有効
End Property

' =========================================================================
' 処理概要       : 要求の詳細記録設定を更新
' 引数           : 有効無効 (Boolean) - 機能を有効にする場合 True
' =========================================================================
Public Property Let 要求の詳細記録設定を取得(ByVal 有効無効 As Boolean)
    保持詳細有効 = 有効無効
    
End Property

' =========================================================================
' 処理概要       : 通信詳細の最大保持件数を取得
' 戻り値         : Long - 通信詳細の最大保持件数
' =========================================================================
Public Property Get 通信詳細の最大保持件数を取得() As Long
    通信詳細の最大保持件数を取得 = 保持詳細保持上限
End Property

' =========================================================================
' 処理概要       : 通信詳細の最大保持件数を設定
' 引数           : 件数 (Long) - 処理対象の件数
' =========================================================================
Public Property Let 通信詳細の最大保持件数を取得(ByVal 件数 As Long)
    If 件数 < 1 Then 件数 = 1
    If 件数 > 通信詳細保持上限最大 Then
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.NetDetailCapacity: " & 件数 & " は上限 " & 通信詳細保持上限最大 & " を超えたため上限値を使用"
        件数 = 通信詳細保持上限最大
    End If
    保持詳細保持上限 = 件数
    通信詳細のリングバッファーを確保
    
End Property

' =========================================================================
' 処理概要       : 要求本文の最大取得バイト数を取得
' 戻り値         : Long - 要求本文の最大取得バイト数
' =========================================================================
Public Property Get 要求本文の最大取得バイト数を取得() As Long
    要求本文の最大取得バイト数を取得 = 保持詳細本文最大
End Property

' =========================================================================
' 処理概要       : 要求本文の最大取得バイト数を設定
' 引数           : 件数 (Long) - 処理対象の件数
' =========================================================================
Public Property Let 要求本文の最大取得バイト数を取得(ByVal 件数 As Long)
    If 件数 < 0 Then 件数 = 0
    If 件数 > 通信本文最大上限 Then 件数 = 通信本文最大上限
    保持詳細本文最大 = 件数
    
End Property

' =========================================================================
' 処理概要       : 機密ヘッダーを伏せ字にする設定を取得
' 戻り値         : Boolean - 有効時 True / 無効時 False
' =========================================================================
Public Property Get 機密ヘッダーを伏せ字にする設定を取得() As Boolean
    機密ヘッダーを伏せ字にする設定を取得 = 保持詳細機密値除外
End Property

' =========================================================================
' 処理概要       : 機密ヘッダーを伏せ字にする設定を更新
' 引数           : 有効無効 (Boolean) - 機能を有効にする場合 True
' =========================================================================
Public Property Let 機密ヘッダーを伏せ字にする設定を取得(ByVal 有効無効 As Boolean)
    保持詳細機密値除外 = 有効無効
    If Not 有効無効 Then
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.NetRedact = False Cookie などがログに平文で残る " & "ログを人に渡す前に必ず確かめること。"
    End If
End Property

' =========================================================================
' 処理概要       : 伏せ字にするヘッダー名を追加
' 引数           : ヘッダー名称 (String) - 対象のHTTPヘッダー名
' =========================================================================
Public Sub 伏せ字にするヘッダー名を追加(ByVal ヘッダー名称 As String)
    Dim 開始時刻 As String
    
    開始時刻 = LCase$(Trim$(ヘッダー名称))
    If LenB(開始時刻) = 0 Then Exit Sub
    If InStr(1, 保持詳細機密値除外一覧, "|" & 開始時刻 & "|") > 0 Then Exit Sub
    
    保持詳細機密値除外一覧 = 保持詳細機密値除外一覧 & 開始時刻 & "|"
    
End Sub

' =========================================================================
' 処理概要       : 伏せ字にするヘッダー名の一覧を消去
' =========================================================================
Public Sub 伏せ字にするヘッダー名の一覧を消去()
    保持詳細機密値除外一覧 = "|"
    mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.ClearRedactHeaders: 伏せる対象が 1 つも無くなった " & "Cookie も平文で出る。"
End Sub

' =========================================================================
' 処理概要       : 詳細記録の対象URL条件を取得
' 戻り値         : String - 詳細記録の対象URL条件
' =========================================================================
Public Property Get 詳細記録の対象URL条件を取得() As String
    詳細記録の対象URL条件を取得 = 保持詳細URIフィルター
End Property

' =========================================================================
' 処理概要       : 詳細記録の対象URL条件を設定
' 引数           : パターン (String) - 通信記録の対象URLパターン
' =========================================================================
Public Property Let 詳細記録の対象URL条件を取得(ByVal パターン As String)
    保持詳細URIフィルター = パターン
    
End Property

' =========================================================================
' 処理概要       : 保持する通信詳細の件数を取得
' 戻り値         : Long - 保持する通信詳細の件数
' =========================================================================
Public Property Get 保持する通信詳細の件数を取得() As Long
    保持する通信詳細の件数を取得 = 保持詳細件数
End Property

' =========================================================================
' 処理概要       : 詳細の上限超過で破棄した件数を取得
' 戻り値         : Long - 詳細の上限超過で破棄した件数
' =========================================================================
Public Property Get 詳細の上限超過で破棄した件数を取得() As Long
    詳細の上限超過で破棄した件数を取得 = 保持詳細破棄件数
End Property

' =========================================================================
' 処理概要       : 指定位置の通信詳細の概要を取得
' 引数           : 対象添字 (Long) - 通信記録の位置（1始まり）
' 戻り値         : String - 指定位置の通信詳細の概要
' =========================================================================
Public Property Get 指定位置の通信詳細の概要を取得(ByVal 対象添字 As Long) As String
    Dim 参照キー As Long
    
    参照キー = 通信詳細の表示位置を配列添字に変換(対象添字)
    If 参照キー < 0 Then Exit Property
    
    指定位置の通信詳細の概要を取得 = 保持詳細連番(参照キー) & vbTab & 保持詳細メソッド(参照キー) & vbTab & 保持詳細URI(参照キー) & vbTab & 保持詳細バイト数(参照キー) & vbTab & CStr(保持詳細打切り(参照キー)) & vbTab & 保持詳細件数型(参照キー)
End Property

' =========================================================================
' 処理概要       : 指定位置の要求ヘッダーを取得
' 引数           : 対象添字 (Long) - 通信記録の位置（1始まり）
' 戻り値         : String - 指定位置の要求ヘッダー
' =========================================================================
Public Property Get 指定位置の要求ヘッダーを取得(ByVal 対象添字 As Long) As String
    Dim 参照キー As Long
    
    参照キー = 通信詳細の表示位置を配列添字に変換(対象添字)
    If 参照キー < 0 Then Exit Property
    
    指定位置の要求ヘッダーを取得 = 対象ヘッダーの値を伏せ字へ変換(保持詳細ヘッダー一覧(参照キー))
End Property

' =========================================================================
' 処理概要       : 指定位置の要求本文を取得
' 引数           : 対象添字 (Long) - 通信記録の位置（1始まり）
' 戻り値         : String - 指定位置の要求本文
' =========================================================================
Public Property Get 指定位置の要求本文を取得(ByVal 対象添字 As Long) As String
    Dim 参照キー As Long
    
    参照キー = 通信詳細の表示位置を配列添字に変換(対象添字)
    If 参照キー < 0 Then Exit Property
    
    指定位置の要求本文を取得 = 保持詳細本文(参照キー)
End Property

' =========================================================================
' 処理概要       : 指定位置の通信詳細をログへ出力
' 引数           : 対象添字 (Long) - 通信記録の位置（1始まり）
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 指定位置の通信詳細をログへ出力(ByVal 対象添字 As Long) As Boolean
    Dim 参照キー As Long
    Dim 添字 As Long
    Dim 行一覧 As Variant
    
    参照キー = 通信詳細の表示位置を配列添字に変換(対象添字)
    If 参照キー < 0 Then Exit Function
    
    mod_WV2_ログ管理.情報レベルのログを記録 "  --- 詳細 #" & 保持詳細連番(参照キー) & " ---"
    mod_WV2_ログ管理.情報レベルのログを記録 "      " & 保持詳細メソッド(参照キー) & " " & 保持詳細URI(参照キー)
    
    If LenB(保持詳細ヘッダー一覧(参照キー)) > 0 Then
        行一覧 = Split(対象ヘッダーの値を伏せ字へ変換(保持詳細ヘッダー一覧(参照キー)), vbLf)
        mod_WV2_ログ管理.情報レベルのログを記録 "      [ヘッダ " & (UBound(行一覧) - LBound(行一覧) + 1) & " 個]"
        For 添字 = LBound(行一覧) To UBound(行一覧)
            mod_WV2_ログ管理.情報レベルのログを記録 "        " & 行一覧(添字)
        Next 添字
    Else
        mod_WV2_ログ管理.情報レベルのログを記録 "      [ヘッダ] 取れていない"
    End If
    
    If 保持詳細バイト数(参照キー) = 0 Then
        mod_WV2_ログ管理.情報レベルのログを記録 "      [ボディ] 無し"
    ElseIf 保持詳細判定文字列(参照キー) Then
        mod_WV2_ログ管理.情報レベルのログを記録 "      [ボディ] " & 保持詳細バイト数(参照キー) & " バイト (テキスト" & IIf(保持詳細打切り(参照キー), "・上限で切った", "") & ")"
        行一覧 = Split(Replace$(保持詳細本文(参照キー), vbCrLf, vbLf), vbLf)
        For 添字 = LBound(行一覧) To UBound(行一覧)
            mod_WV2_ログ管理.情報レベルのログを記録 "        " & 行一覧(添字)
        Next 添字
    Else
        mod_WV2_ログ管理.情報レベルのログを記録 "      [ボディ] " & 保持詳細バイト数(参照キー) & " バイト (バイナリ" & IIf(保持詳細打切り(参照キー), "・上限で切った", "") & ") 先頭: " & 保持詳細16進(参照キー)
    End If
    
    If 保持詳細状態(参照キー) = 0 Then
        mod_WV2_ログ管理.情報レベルのログを記録 "      [応答] まだ来ていない"
    Else
        mod_WV2_ログ管理.情報レベルのログを記録 "      [応答] " & 通信状態とHTTPステータスを表示用文字列へ変換(保持詳細状態(参照キー)) & " " & 保持詳細理由(参照キー)
        If LenB(保持詳細応答ヘッダー一覧(参照キー)) > 0 Then
            行一覧 = Split(対象ヘッダーの値を伏せ字へ変換(保持詳細応答ヘッダー一覧(参照キー)), vbLf)
            mod_WV2_ログ管理.情報レベルのログを記録 "      [応答ヘッダ " & (UBound(行一覧) - LBound(行一覧) + 1) & " 個]"
            For 添字 = LBound(行一覧) To UBound(行一覧)
                mod_WV2_ログ管理.情報レベルのログを記録 "        " & 行一覧(添字)
            Next 添字
        End If
        
        If 保持詳細応答バイト数(参照キー) > 0 Or LenB(保持詳細応答文字コード(参照キー)) > 0 Then
            mod_WV2_ログ管理.情報レベルのログを記録 "      [応答本文] " & 保持詳細応答バイト数(参照キー) & " バイト" & IIf(保持詳細応答打切り(参照キー), " 上限で切った", "") & IIf(LenB(保持詳細応答文字コード(参照キー)) > 0, " encoding=" & 保持詳細応答文字コード(参照キー), "") & IIf(保持詳細応答判定文字列(参照キー), " (テキスト)", _
 _
                " (バイナリ)")
            mod_WV2_ログ管理.情報レベルのログを記録 "        先頭: " & 保持詳細応答16進(参照キー)
            If LenB(保持詳細応答本文(参照キー)) > 0 Then
                行一覧 = Split(Replace$(Left$(保持詳細応答本文(参照キー), 2000), vbCrLf, vbLf), vbLf)
                For 添字 = LBound(行一覧) To UBound(行一覧)
                    mod_WV2_ログ管理.情報レベルのログを記録 "        " & 行一覧(添字)
                Next 添字
            End If
        End If
    End If
    
    指定位置の通信詳細をログへ出力 = True
End Function

' =========================================================================
' 処理概要       : 保持する通信詳細をまとめてログへ出力
' 戻り値         : Long - ログへ出力した件数
' =========================================================================
Public Function 保持する通信詳細をまとめてログへ出力() As Long
    Dim 添字 As Long
    
    If 保持詳細件数 = 0 Then
        mod_WV2_ログ管理.情報レベルのログを記録 "  --- 詳細: 保持データなし ---"
        Exit Function
    End If
    
    mod_WV2_ログ管理.情報レベルのログを記録 "  --- 詳細 " & 保持詳細件数 & " 件 ---"
    If 保持詳細破棄件数 > 0 Then
        mod_WV2_ログ管理.警告レベルのログを記録 "  詳細リングから溢れて捨てたものが " & 保持詳細破棄件数 & " 件ある " & "NetDetailCapacity を増やすか、NetDetailUriFilter で絞ること。"
    End If
    For 添字 = 1 To 保持詳細件数
        指定位置の通信詳細をログへ出力 添字
    Next 添字
    保持する通信詳細をまとめてログへ出力 = 保持詳細件数
End Function

' =========================================================================
' 処理概要       : 通信詳細の記録を消去
' =========================================================================
Public Sub 通信詳細の記録を消去()
    通信詳細のリングバッファーを確保
End Sub

' =========================================================================
' 処理概要       : COMインターフェースからヘッダーを取得
' 引数           : ポインター要求 (LongPtr) - 要求インターフェースのポインター
' 戻り値         : String - COMインターフェースからヘッダー
' =========================================================================
Private Function COMインターフェースからヘッダーを取得(ByVal ポインター要求 As LongPtr) As String
    Dim ポインターヘッダー一覧 As LongPtr
    Dim 実行結果       As Long
    
    If ポインター要求 = 0 Then Exit Function
    
    実行結果 = COMメソッド呼出(ポインター要求, 9, "", VarPtr(ポインターヘッダー一覧))
    If 実行結果 <> 0 Or ポインターヘッダー一覧 = 0 Then Exit Function
    
    COMインターフェースからヘッダーを取得 = イテレーターからヘッダーを順に取得(ポインターヘッダー一覧, 8)
    COM参照解放 ポインターヘッダー一覧
End Function

' =========================================================================
' 処理概要       : イテレーターからヘッダーを順に取得
' 引数           : ポインターヘッダー一覧 (LongPtr) - ヘッダーインターフェースのポインター
'                  列挙位置番号 (Long) - ヘッダーイテレーターのvTable位置
' 戻り値         : String - イテレーターからヘッダーを順に取得
' =========================================================================
Private Function イテレーターからヘッダーを順に取得(ByVal ポインターヘッダー一覧 As LongPtr, ByVal 列挙位置番号 As Long) As String
    Dim ポインター列挙    As LongPtr
    Dim 実行結果       As Long
    Dim 件数        As Long
    Dim ポインター名称    As LongPtr
    Dim ポインター値   As LongPtr
    Dim 続き状態 As Long
    Dim 出力文字列       As String
    
    If ポインターヘッダー一覧 = 0 Then
        イテレーターからヘッダーを順に取得 = "(診断: ヘッダのコレクションが 0)"
        Exit Function
    End If
    
    実行結果 = COMメソッド呼出(ポインターヘッダー一覧, 列挙位置番号, "", VarPtr(ポインター列挙))
    If 実行結果 <> 0 Then
        イテレーターからヘッダーを順に取得 = "(診断: GetIterator が失敗 hr=&H" & Hex(実行結果) & " index=" & 列挙位置番号 & ")"
        Exit Function
    End If
    If ポインター列挙 = 0 Then
        イテレーターからヘッダーを順に取得 = "(診断: GetIterator は S_OK だが iter が 0 index=" & 列挙位置番号 & ")"
        Exit Function
    End If
    
    ' 上限件数までヘッダーを順に取得する
    Do While 件数 < 通信ヘッダー最大
        If Not COMメソッドから真偽値プロパティを取得(ポインター列挙, 4, "") Then Exit Do
        
        ポインター名称 = 0
        ポインター値 = 0
        実行結果 = COMメソッド呼出(ポインター列挙, 3, "", VarPtr(ポインター名称), VarPtr(ポインター値))
        If 実行結果 = 0 Then
            If LenB(出力文字列) > 0 Then 出力文字列 = 出力文字列 & vbLf
            出力文字列 = 出力文字列 & UTF16文字列ポインターをVBA文字列へ変換(ポインター名称) & ": " & UTF16文字列ポインターをVBA文字列へ変換(ポインター値)
            件数 = 件数 + 1
        End If
        If ポインター名称 <> 0 Then COMメモリー解放API ポインター名称
        If ポインター値 <> 0 Then COMメモリー解放API ポインター値
        If 実行結果 <> 0 Then Exit Do
        
        続き状態 = 0
        実行結果 = COMメソッド呼出(ポインター列挙, 5, "", VarPtr(続き状態))
        If 実行結果 <> 0 Or 続き状態 = 0 Then Exit Do
    Loop
    
    If 件数 >= 通信ヘッダー最大 Then
        出力文字列 = 出力文字列 & vbLf & "(" & 通信ヘッダー最大 & " 個で打ち切った)"
    End If
    If 件数 = 0 Then
    
        出力文字列 = "(診断: ヘッダが 0 個。HasCurrentHeader が最初から False" & " index=" & 列挙位置番号 & ")"
    End If
    
    COM参照解放 ポインター列挙
    イテレーターからヘッダーを順に取得 = 出力文字列
End Function

' =========================================================================
' 処理概要       : 要求本文を上限バイト数まで取得
' 引数           : ポインター要求 (LongPtr) - 要求インターフェースのポインター
'                  出力バイト数 (Long) - ［出力］出力先のバイト配列または取得バイト数
'                  出力打切り (Boolean) - ［出力］上限による打ち切り有無の出力先
'                  出力バッファー (Byte()) - ［出力］出力先のバイト配列
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Private Function 要求本文を上限バイト数まで取得(ByVal ポインター要求 As LongPtr, ByRef 出力バイト数 As Long, ByRef 出力打切り As Boolean, ByRef 出力バッファー() As Byte) As Boolean
    Dim ポインターストリーム As LongPtr
    Dim 実行結果      As Long
    Dim 保持上限     As Long
    Dim 取得済み     As Long
    Dim 読込単位   As Long
    Dim バイト数読込  As Long
    
    出力バイト数 = 0
    出力打切り = False
    
    If ポインター要求 = 0 Then Exit Function
    
    実行結果 = COMメソッド呼出(ポインター要求, 7, "", VarPtr(ポインターストリーム))
    If 実行結果 <> 0 Or ポインターストリーム = 0 Then Exit Function
    
    保持上限 = 保持詳細本文最大
    If 保持上限 < 1 Then
    
        COM参照解放 ポインターストリーム
        要求本文を上限バイト数まで取得 = True
        Exit Function
    End If
    
    If Not COMストリームを上限バイト数まで読み取り(ポインターストリーム, 保持上限, 取得済み, 出力打切り, 出力バッファー) Then
        COM参照解放 ポインターストリーム
        Exit Function
    End If
    
    実行結果 = COMメソッド呼出(ポインターストリーム, 5, "", 0^, ストリーム先頭位置, 0^)
    If 実行結果 <> 0 Then
    
        出力打切り = True
        出力バイト数 = 取得済み
        COM参照解放 ポインターストリーム
        要求本文を上限バイト数まで取得 = True
        Exit Function
    End If
    
    COM参照解放 ポインターストリーム
    出力バイト数 = 取得済み
    要求本文を上限バイト数まで取得 = True
End Function

' =========================================================================
' 処理概要       : COMストリームを上限バイト数まで読み取り
' 引数           : ポインターストリーム (LongPtr) - 読み取るCOMストリームのポインター
'                  保持上限 (Long) - 読み取る最大バイト数
'                  出力バイト数 (Long) - ［出力］出力先のバイト配列または取得バイト数
'                  出力打切り (Boolean) - ［出力］上限による打ち切り有無の出力先
'                  出力バッファー (Byte()) - ［出力］出力先のバイト配列
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Private Function COMストリームを上限バイト数まで読み取り(ByVal ポインターストリーム As LongPtr, ByVal 保持上限 As Long, ByRef 出力バイト数 As Long, ByRef 出力打切り As Boolean, ByRef 出力バッファー() As Byte) As Boolean
    Dim 実行結果     As Long
    Dim 取得済み    As Long
    Dim 読込単位  As Long
    Dim バイト数読込 As Long
    Dim 画面監視(0 To 0) As Byte
    
    出力バイト数 = 0
    出力打切り = False
    If ポインターストリーム = 0 Or 保持上限 < 1 Then Exit Function
    
    ReDim 出力バッファー(0 To 保持上限 - 1)
    
    ' 保存できる容量まで本文を順に読み取る
    Do While 取得済み < 保持上限
        読込単位 = 保持上限 - 取得済み
        If 読込単位 > 通信本文読込単位 Then 読込単位 = 通信本文読込単位
        バイト数読込 = 0
        実行結果 = COMメソッド呼出(ポインターストリーム, 3, "", VarPtr(出力バッファー(取得済み)), 読込単位, VarPtr(バイト数読込))
        If 実行結果 <> 0 Then Exit Do
        If バイト数読込 <= 0 Then Exit Do
        取得済み = 取得済み + バイト数読込
    Loop
    
    If 取得済み >= 保持上限 Then
        バイト数読込 = 0
        実行結果 = COMメソッド呼出(ポインターストリーム, 3, "", VarPtr(画面監視(0)), 1&, VarPtr(バイト数読込))
        If 実行結果 = 0 And バイト数読込 > 0 Then 出力打切り = True
    End If
    
    出力バイト数 = 取得済み
    COMストリームを上限バイト数まで読み取り = True
End Function

' =========================================================================
' 処理概要       : 通信詳細をリングバッファーに追加
' 引数           : 連番番号 (Long) - 通信記録の連番
'                  HTTPメソッド (String) - HTTPメソッド
'                  URI (String) - 通信先URI
'                  ヘッダー一覧 (String) - HTTPヘッダーの文字列
'                  ContentType値 (String) - HTTPのContent-Type
'                  件数バイト数 (Long) - 取得した本文のバイト数
'                  打切り (Boolean) - 本文を上限で打ち切った場合 True
'                  判定文字列 (Boolean) - 本文がテキスト形式かどうか
'                  本文文字列 (String) - 記録する本文
'                  数値16進先頭 (String) - バイナリ本文の先頭を表す16進文字列
' =========================================================================
Private Sub 通信詳細をリングバッファーに追加(ByVal 連番番号 As Long, ByVal HTTPメソッド As String, ByVal URI As String, ByVal ヘッダー一覧 As String, ByVal ContentType値 As String, ByVal 件数バイト数 As Long, ByVal 打切り As Boolean, ByVal 判定文字列 As Boolean, _
    ByVal 本文文字列 As String, ByVal 数値16進先頭 As String)
    保持詳細連番(保持詳細先頭) = 連番番号
    保持詳細メソッド(保持詳細先頭) = HTTPメソッド
    保持詳細URI(保持詳細先頭) = URI
    保持詳細ヘッダー一覧(保持詳細先頭) = ヘッダー一覧
    保持詳細件数型(保持詳細先頭) = ContentType値
    保持詳細バイト数(保持詳細先頭) = 件数バイト数
    保持詳細打切り(保持詳細先頭) = 打切り
    保持詳細判定文字列(保持詳細先頭) = 判定文字列
    保持詳細本文(保持詳細先頭) = 本文文字列
    保持詳細16進(保持詳細先頭) = 数値16進先頭
    
    保持詳細状態(保持詳細先頭) = 0
    保持詳細理由(保持詳細先頭) = ""
    保持詳細応答ヘッダー一覧(保持詳細先頭) = ""
    
    保持詳細応答本文(保持詳細先頭) = ""
    保持詳細応答バイト数(保持詳細先頭) = 0
    保持詳細応答打切り(保持詳細先頭) = False
    保持詳細応答判定文字列(保持詳細先頭) = False
    保持詳細応答文字コード(保持詳細先頭) = ""
    保持詳細応答16進(保持詳細先頭) = ""
    保持詳細応答取得済み(保持詳細先頭) = False
    
    保持詳細先頭 = 保持詳細先頭 + 1
    If 保持詳細先頭 >= 保持詳細保持上限 Then 保持詳細先頭 = 0
    If 保持詳細件数 < 保持詳細保持上限 Then
        保持詳細件数 = 保持詳細件数 + 1
    Else
        保持詳細破棄件数 = 保持詳細破棄件数 + 1
    End If
End Sub

' =========================================================================
' 処理概要       : 通信詳細のリングバッファーを確保
' =========================================================================
Private Sub 通信詳細のリングバッファーを確保()
    If 保持詳細保持上限 < 1 Then 保持詳細保持上限 = 通信詳細保持上限既定
    
    ReDim 保持詳細連番(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細メソッド(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細URI(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細ヘッダー一覧(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細件数型(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細バイト数(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細打切り(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細判定文字列(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細本文(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細16進(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細応答ヘッダー一覧(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細状態(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細理由(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細応答本文(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細応答バイト数(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細応答打切り(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細応答判定文字列(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細応答文字コード(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細応答16進(0 To 保持詳細保持上限 - 1)
    ReDim 保持詳細応答取得済み(0 To 保持詳細保持上限 - 1)
    
    保持詳細先頭 = 0
    保持詳細件数 = 0
    保持詳細破棄件数 = 0
End Sub

' =========================================================================
' 処理概要       : 通信詳細の表示位置を配列添字に変換
' 引数           : 対象添字 (Long) - 通信記録の位置（1始まり）
' 戻り値         : Long - 対応する配列添字。該当しない場合 -1
' =========================================================================
Private Function 通信詳細の表示位置を配列添字に変換(ByVal 対象添字 As Long) As Long
    Dim 参照キー As Long
    
    If 対象添字 < 1 Or 対象添字 > 保持詳細件数 Then
        通信詳細の表示位置を配列添字に変換 = -1
        Exit Function
    End If
    
    参照キー = 保持詳細先頭 - 保持詳細件数 + (対象添字 - 1)
    Do While 参照キー < 0
        参照キー = 参照キー + 保持詳細保持上限
    Loop
    通信詳細の表示位置を配列添字に変換 = 参照キー Mod 保持詳細保持上限
End Function

' =========================================================================
' 処理概要       : 対象ヘッダーの値を伏せ字へ変換
' 引数           : ヘッダー一覧 (String) - HTTPヘッダーの文字列
' 戻り値         : String - 対象ヘッダーの値を伏せ字
' =========================================================================
Private Function 対象ヘッダーの値を伏せ字へ変換(ByVal ヘッダー一覧 As String) As String
    Dim 行一覧 As Variant
    Dim 添字     As Long
    Dim 対象参照     As Long
    Dim 場所名    As String
    Dim 出力   As String
    
    If LenB(ヘッダー一覧) = 0 Then Exit Function
    
    If Not 保持詳細機密値除外 Then
        対象ヘッダーの値を伏せ字へ変換 = ヘッダー一覧
        Exit Function
    End If
    
    行一覧 = Split(ヘッダー一覧, vbLf)
    For 添字 = LBound(行一覧) To UBound(行一覧)
        対象参照 = InStr(1, CStr(行一覧(添字)), ":")
        If 対象参照 > 1 Then
            場所名 = LCase$(Trim$(Left$(CStr(行一覧(添字)), 対象参照 - 1)))
            If InStr(1, 保持詳細機密値除外一覧, "|" & 場所名 & "|") > 0 Then
                行一覧(添字) = Left$(CStr(行一覧(添字)), 対象参照) & " <伏せた " & (Len(CStr(行一覧(添字))) - 対象参照 - 1) & " 文字>"
            End If
        End If
        If LenB(出力) > 0 Then 出力 = 出力 & vbLf
        出力 = 出力 & CStr(行一覧(添字))
    Next 添字
    
    対象ヘッダーの値を伏せ字へ変換 = 出力
End Function

' =========================================================================
' 処理概要       : Content-Typeがテキスト形式か判定
' 引数           : ContentType値 (String) - HTTPのContent-Type
' 戻り値         : Boolean - テキスト形式の場合 True
' =========================================================================
Private Function ContentTypeがテキスト形式か判定(ByVal ContentType値 As String) As Boolean
    Dim 開始時刻 As String
    
    開始時刻 = LCase$(ContentType値)
    If LenB(開始時刻) = 0 Then Exit Function
    
    If InStr(1, 開始時刻, "text/") > 0 Then ContentTypeがテキスト形式か判定 = True: Exit Function
    If InStr(1, 開始時刻, "json") > 0 Then ContentTypeがテキスト形式か判定 = True: Exit Function
    If InStr(1, 開始時刻, "xml") > 0 Then ContentTypeがテキスト形式か判定 = True: Exit Function
    If InStr(1, 開始時刻, "x-www-form-urlencoded") > 0 Then ContentTypeがテキスト形式か判定 = True: Exit Function
    If InStr(1, 開始時刻, "javascript") > 0 Then ContentTypeがテキスト形式か判定 = True: Exit Function
    If InStr(1, 開始時刻, "charset=") > 0 Then ContentTypeがテキスト形式か判定 = True: Exit Function
End Function

' =========================================================================
' 処理概要       : バイト配列の先頭を16進文字列へ変換
' 引数           : 入力データ (Byte()) - 変換対象のバイト配列
'                  バイト数 (Long) - 読み取るバイト数
' 戻り値         : String - バイト配列の先頭を16進文字列
' =========================================================================
Private Function バイト配列の先頭を16進文字列へ変換(ByRef 入力データ() As Byte, ByVal バイト数 As Long) As String
    Dim 添字 As Long
    Dim 件数 As Long
    Dim 作成文字列 As String
    
    件数 = バイト数
    If 件数 > 通信16進先頭 Then 件数 = 通信16進先頭
    For 添字 = 0 To 件数 - 1
        作成文字列 = 作成文字列 & Right$("0" & Hex$(入力データ(添字)), 2) & " "
    Next 添字
    If バイト数 > 件数 Then 作成文字列 = 作成文字列 & "..."
    バイト配列の先頭を16進文字列へ変換 = RTrim$(作成文字列)
End Function

' =========================================================================
' 処理概要       : ヘッダー文字列から指定ヘッダーの値を取得
' 引数           : ヘッダー一覧 (String) - HTTPヘッダーの文字列
'                  ヘッダー名称 (String) - 対象のHTTPヘッダー名
' 戻り値         : String - ヘッダー文字列から指定ヘッダーの値
' =========================================================================
Private Function ヘッダー文字列から指定ヘッダーの値を取得(ByVal ヘッダー一覧 As String, ByVal ヘッダー名称 As String) As String
    Dim 行一覧 As Variant
    Dim 添字     As Long
    Dim 対象参照     As Long
    Dim 場所名    As String
    
    If LenB(ヘッダー一覧) = 0 Then Exit Function
    
    行一覧 = Split(ヘッダー一覧, vbLf)
    For 添字 = LBound(行一覧) To UBound(行一覧)
        対象参照 = InStr(1, CStr(行一覧(添字)), ":")
        If 対象参照 > 1 Then
            場所名 = LCase$(Trim$(Left$(CStr(行一覧(添字)), 対象参照 - 1)))
            If 場所名 = LCase$(ヘッダー名称) Then
                ヘッダー文字列から指定ヘッダーの値を取得 = Trim$(Mid$(CStr(行一覧(添字)), 対象参照 + 1))
                Exit Function
            End If
        End If
    Next 添字
End Function

' =========================================================================
' 処理概要       : 伏せ字にする既定のヘッダー名一覧を取得
' 戻り値         : String - 伏せ字にする既定のヘッダー名一覧
' =========================================================================
Private Function 伏せ字にする既定のヘッダー名一覧を取得() As String
    伏せ字にする既定のヘッダー名一覧を取得 = "|cookie|set-cookie|authorization|proxy-authorization|" & "x-api-key|x-auth-token|x-csrf-token|api-key|"
End Function

' =========================================================================
' 処理概要       : HTTPステータスを含めて通信一覧の行を取得
' 引数           : 対象添字 (Long) - 通信記録の位置（1始まり）
' 戻り値         : Long - HTTPステータスを含めて通信一覧の行
' =========================================================================
Public Property Get HTTPステータスを含めて通信一覧の行を取得(ByVal 対象添字 As Long) As Long
    Dim 参照キー As Long
    
    参照キー = 通信一覧の表示位置を配列添字に変換(対象添字)
    If 参照キー < 0 Then Exit Property
    
    HTTPステータスを含めて通信一覧の行を取得 = 保持通信状態(参照キー)
End Property

' =========================================================================
' 処理概要       : 応答イベントの記録状態を取得
' 戻り値         : Boolean - 有効時 True / 無効時 False
' =========================================================================
Public Property Get 応答イベントの記録状態を取得() As Boolean
    応答イベントの記録状態を取得 = 保持通信応答有効
End Property

' =========================================================================
' 処理概要       : 要求と対応付けられなかった応答の件数を取得
' 戻り値         : Long - 要求と対応付けられなかった応答の件数
' =========================================================================
Public Property Get 要求と対応付けられなかった応答の件数を取得() As Long
    要求と対応付けられなかった応答の件数を取得 = 保持通信未対応件数
End Property

' =========================================================================
' 処理概要       : 指定位置の応答ヘッダーを取得
' 引数           : 対象添字 (Long) - 通信記録の位置（1始まり）
' 戻り値         : String - 指定位置の応答ヘッダー
' =========================================================================
Public Property Get 指定位置の応答ヘッダーを取得(ByVal 対象添字 As Long) As String
    Dim 参照キー As Long
    
    参照キー = 通信詳細の表示位置を配列添字に変換(対象添字)
    If 参照キー < 0 Then Exit Property
    
    指定位置の応答ヘッダーを取得 = 対象ヘッダーの値を伏せ字へ変換(保持詳細応答ヘッダー一覧(参照キー))
End Property

' =========================================================================
' 処理概要       : 通信応答イベントを登録
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function 通信応答イベントを登録() As Long
    Dim 通知ハンドラー As mod_WV2_COM通知ハンドラー
    Dim 仮想COMオブジェクト As LongPtr
    Dim 登録トークン As LongLong
    Dim 実行結果 As Long
    
    If 保持終了済み Then
        通信応答イベントを登録 = &H80004003
        Exit Function
    End If
    If Not ICoreWebView2_2を取得して保持() Then
    
        通信応答イベントを登録 = &H80004003
        Exit Function
    End If
    
    If Not (保持永続ハンドラー一覧(通信応答通知) Is Nothing) Then
    
        通信応答イベントを登録 = 0
        Exit Function
    End If
    
    Set 通知ハンドラー = 通知先を設定したCOMコールバックハンドラーを取得(通信応答通知, Me)
    If 通知ハンドラー Is Nothing Then
    
        通信応答イベントを登録 = &H8000FFFF
        Exit Function
    End If
    
    仮想COMオブジェクト = 通知ハンドラー.スロットアドレス + 仮想COMオブジェクト位置
    
    実行結果 = COMメソッド呼出(保持ポインターWebView2, 61, "View2.add_WebResourceResponseReceived", 仮想COMオブジェクト, VarPtr(登録トークン))
    
    If 実行結果 <> 0 Then
    
        Set 通知ハンドラー = Nothing
        通信応答イベントを登録 = 実行結果
        Exit Function
    End If
    
    保持永続登録トークン一覧(通信応答通知) = 登録トークン
    Set 保持永続ハンドラー一覧(通信応答通知) = 通知ハンドラー
    
    通信応答イベントを登録 = 0
End Function

' =========================================================================
' 処理概要       : 通信応答イベントを解除
' 戻り値         : Long - HRESULT（成功時 0 / 失敗時はエラーコード）
' =========================================================================
Public Function 通信応答イベントを解除() As Long
    Dim 実行結果 As Long
    
    If 保持永続ハンドラー一覧(通信応答通知) Is Nothing Then
        通信応答イベントを解除 = 0
        Exit Function
    End If
    
    If 保持ポインターWebView2 <> 0 Then
        実行結果 = COMメソッド呼出(保持ポインターWebView2, 62, "View2.remove_WebResourceResponseReceived", 保持永続登録トークン一覧(通信応答通知))
    End If
    
    Set 保持永続ハンドラー一覧(通信応答通知) = Nothing
    保持永続登録トークン一覧(通信応答通知) = 0^
    
    通信応答イベントを解除 = 実行結果
End Function

' =========================================================================
' 処理概要       : 通信応答イベントを受け付けてコールバック状態を管理
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
'                  ポインター引数一覧 (LongPtr) - イベント引数のCOMポインター
' =========================================================================
Public Sub 通信応答イベントを受け付けてコールバック状態を管理(ByVal ポインター通知元 As LongPtr, ByVal ポインター引数一覧 As LongPtr)
    Dim エラー番号 As Long
    Dim エラー説明 As String
    
    ' コールバック中であることを記録し、同期処理の再入を防ぐ
    保持内部通知 = 保持内部通知 + 1
    On Error GoTo ErrorHanler
    通信応答イベントの内容を処理 ポインター通知元, ポインター引数一覧
    
    GoTo EndProc
    
ErrorHanler:
    エラー番号 = Err.Number
    エラー説明 = Err.Description
    MsgBox "通信イベントの処理でエラーが発生しました。" & vbCrLf & エラー説明, vbExclamation, "ブラウザー"
    Resume EndProc
    
EndProc:

    ' 通知処理の終了時にコールバック階層を戻す
    If 保持内部通知 > 0 Then 保持内部通知 = 保持内部通知 - 1
    If エラー番号 <> 0 Then
        On Error Resume Next
        保持通信未対応件数 = 保持通信未対応件数 + 1
    End If
End Sub

' =========================================================================
' 処理概要       : 通信応答イベントの内容を処理
' 引数           : ポインター通知元 (LongPtr) - イベント通知元のCOMポインター
'                  ポインター引数一覧 (LongPtr) - イベント引数のCOMポインター
' =========================================================================
Private Sub 通信応答イベントの内容を処理(ByVal ポインター通知元 As LongPtr, ByVal ポインター引数一覧 As LongPtr)
    Dim ポインター要求  As LongPtr
    Dim ポインター応答 As LongPtr
    Dim 実行結果    As Long
    Dim URI   As String
    Dim HTTPメソッド   As String
    Dim 作業情報    As Long
    Dim 理由 As String
    Dim 参照キー     As Long
    
    If ポインター引数一覧 = 0 Then Exit Sub
    
    実行結果 = COMメソッド呼出(ポインター引数一覧, 3, "", VarPtr(ポインター要求))
    If 実行結果 <> 0 Or ポインター要求 = 0 Then Exit Sub
    
    URI = COMメソッドから文字列プロパティを取得(ポインター要求, 3, "")
    HTTPメソッド = COMメソッドから文字列プロパティを取得(ポインター要求, 5, "")
    COM参照解放 ポインター要求
    
    参照キー = 応答に対応する要求の記録位置を検索(URI, HTTPメソッド)
    If 参照キー < 0 Then
        保持通信未対応件数 = 保持通信未対応件数 + 1
        Exit Sub
    End If
    
    実行結果 = COMメソッド呼出(ポインター引数一覧, 4, "", VarPtr(ポインター応答))
    If 実行結果 <> 0 Or ポインター応答 = 0 Then
        保持通信状態(参照キー) = -1
        保持通信理由(参照キー) = "(応答の取得に失敗)"
        Exit Sub
    End If
    
    作業情報 = COMメソッドから32ビット整数プロパティを取得(ポインター応答, 4, "")
    理由 = COMメソッドから文字列プロパティを取得(ポインター応答, 5, "")
    
    保持通信状態(参照キー) = 作業情報
    保持通信理由(参照キー) = 理由
    
    If 保持詳細有効 Then 要求に対応する応答ヘッダーを詳細へ保存 保持通信連番(参照キー), ポインター応答, 作業情報, 理由
    
    If 保持応答本文有効 And 保持詳細有効 Then 応答本文の非同期取得を開始 保持通信連番(参照キー), ポインター応答
    
    COM参照解放 ポインター応答
End Sub

' =========================================================================
' 処理概要       : 応答に対応する要求の記録位置を検索
' 引数           : URI (String) - 通信先URI
'                  HTTPメソッド (String) - HTTPメソッド
' 戻り値         : Long - 対応する配列添字。該当しない場合 -1
' =========================================================================
Private Function 応答に対応する要求の記録位置を検索(ByVal URI As String, ByVal HTTPメソッド As String) As Long
    Dim 添字 As Long
    Dim 参照キー As Long
    
    応答に対応する要求の記録位置を検索 = -1
    If 保持通信件数 = 0 Then Exit Function
    
    For 添字 = 1 To 保持通信件数
        参照キー = 通信一覧の表示位置を配列添字に変換(添字)
        If 参照キー >= 0 Then
            If 保持通信状態(参照キー) = 0 Then
                If 保持通信URI(参照キー) = URI Then
                    If LenB(HTTPメソッド) = 0 Or 保持通信メソッド(参照キー) = HTTPメソッド Then
                        応答に対応する要求の記録位置を検索 = 参照キー
                        Exit Function
                    End If
                End If
            End If
        End If
    Next 添字
End Function

' =========================================================================
' 処理概要       : 要求に対応する応答ヘッダーを詳細へ保存
' 引数           : 連番番号 (Long) - 通信記録の連番
'                  ポインター応答 (LongPtr) - 応答インターフェースのポインター
'                  作業情報 (Long) - HTTPステータスコード
'                  理由 (String) - 通信状態の補足内容
' =========================================================================
Private Sub 要求に対応する応答ヘッダーを詳細へ保存(ByVal 連番番号 As Long, ByVal ポインター応答 As LongPtr, ByVal 作業情報 As Long, ByVal 理由 As String)
    Dim 添字 As Long
    Dim 参照キー As Long
    Dim ポインターヘッダー一覧 As LongPtr
    Dim 実行結果 As Long
    
    For 添字 = 保持詳細件数 To 1 Step -1
        参照キー = 通信詳細の表示位置を配列添字に変換(添字)
        If 参照キー >= 0 Then
            If 保持詳細連番(参照キー) = 連番番号 Then
                保持詳細状態(参照キー) = 作業情報
                保持詳細理由(参照キー) = 理由
                
                実行結果 = COMメソッド呼出(ポインター応答, 3, "", VarPtr(ポインターヘッダー一覧))
                If 実行結果 <> 0 Then
                    保持詳細応答ヘッダー一覧(参照キー) = "(診断: get_Headers が失敗 hr=&H" & Hex(実行結果) & ")"
                ElseIf ポインターヘッダー一覧 = 0 Then
                    保持詳細応答ヘッダー一覧(参照キー) = "(診断: get_Headers は S_OK だが 0 が返った)"
                Else
                
                    保持詳細応答ヘッダー一覧(参照キー) = イテレーターからヘッダーを順に取得(ポインターヘッダー一覧, 7)
                    COM参照解放 ポインターヘッダー一覧
                End If
                Exit Sub
            End If
        End If
    Next 添字
End Sub

' =========================================================================
' 処理概要       : 応答本文の取得設定を取得
' 戻り値         : Boolean - 有効時 True / 無効時 False
' =========================================================================
Public Property Get 応答本文の取得設定を取得() As Boolean
    応答本文の取得設定を取得 = 保持応答本文有効
End Property

' =========================================================================
' 処理概要       : 応答本文の取得設定を更新
' 引数           : 有効無効 (Boolean) - 機能を有効にする場合 True
' =========================================================================
Public Property Let 応答本文の取得設定を取得(ByVal 有効無効 As Boolean)
    保持応答本文有効 = 有効無効
    If 有効無効 And Not 保持詳細有効 Then
        mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー表示.NetRespBodyOn = True だが NetDetailOn が False なので " & "本文の置き場が無く、1 件も取れない。NetDetailOn = True も要る。"
    End If
End Property

' =========================================================================
' 処理概要       : 応答本文の最大取得バイト数を取得
' 戻り値         : Long - 応答本文の最大取得バイト数
' =========================================================================
Public Property Get 応答本文の最大取得バイト数を取得() As Long
    応答本文の最大取得バイト数を取得 = 保持応答本文最大
End Property

' =========================================================================
' 処理概要       : 応答本文の最大取得バイト数を設定
' 引数           : 件数 (Long) - 処理対象の件数
' =========================================================================
Public Property Let 応答本文の最大取得バイト数を取得(ByVal 件数 As Long)
    If 件数 < 0 Then 件数 = 0
    If 件数 > 通信本文最大上限 Then 件数 = 通信本文最大上限
    保持応答本文最大 = 件数
    
End Property

' =========================================================================
' 処理概要       : 対応する詳細が失われた応答本文の件数を取得
' 戻り値         : Long - 対応する詳細が失われた応答本文の件数
' =========================================================================
Public Property Get 対応する詳細が失われた応答本文の件数を取得() As Long
    対応する詳細が失われた応答本文の件数を取得 = 保持応答本文対応なし件数
End Property

' =========================================================================
' 処理概要       : 指定位置の応答本文を取得
' 引数           : 対象添字 (Long) - 通信記録の位置（1始まり）
' 戻り値         : String - 指定位置の応答本文
' =========================================================================
Public Property Get 指定位置の応答本文を取得(ByVal 対象添字 As Long) As String
    Dim 参照キー As Long
    
    参照キー = 通信詳細の表示位置を配列添字に変換(対象添字)
    If 参照キー < 0 Then Exit Property
    
    指定位置の応答本文を取得 = 保持詳細応答本文(参照キー)
End Property

' =========================================================================
' 処理概要       : 指定位置の応答本文の概要を取得
' 引数           : 対象添字 (Long) - 通信記録の位置（1始まり）
' 戻り値         : String - 指定位置の応答本文の概要
' =========================================================================
Public Property Get 指定位置の応答本文の概要を取得(ByVal 対象添字 As Long) As String
    Dim 参照キー As Long
    
    参照キー = 通信詳細の表示位置を配列添字に変換(対象添字)
    If 参照キー < 0 Then Exit Property
    
    指定位置の応答本文の概要を取得 = 保持詳細応答バイト数(参照キー) & vbTab & CStr(保持詳細応答打切り(参照キー)) & vbTab & CStr(保持詳細応答判定文字列(参照キー)) & vbTab & 保持詳細応答文字コード(参照キー) & vbTab & 保持詳細応答16進(参照キー) & vbTab & CStr(保持詳細応答取得済み(参照キー))
End Property

' =========================================================================
' 処理概要       : 応答本文の非同期取得を開始
' 引数           : 連番番号 (Long) - 通信記録の連番
'                  ポインター応答 (LongPtr) - 応答インターフェースのポインター
' =========================================================================
Private Sub 応答本文の非同期取得を開始(ByVal 連番番号 As Long, ByVal ポインター応答 As LongPtr)
    Dim 通知ハンドラー  As mod_WV2_COM通知ハンドラー
    Dim 実行結果 As Long
    Dim 項目キー As String
    Dim 未使用値 As Variant
    
    If ポインター応答 = 0 Then Exit Sub
    
    If 通信の連番に対応する詳細の記録位置を検索(連番番号) < 0 Then Exit Sub
    
    項目キー = "r" & 連番番号
    
    On Error Resume Next
    未使用値 = 保持応答待機中(項目キー)
    If Err.Number = 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0
    
    Set 通知ハンドラー = 通知先を設定したCOMコールバックハンドラーを取得(応答本文取得通知, Me)
    If 通知ハンドラー Is Nothing Then Exit Sub
    
    通知ハンドラー.通知識別番号 = 連番番号
    
    COM参照追加 ポインター応答
    保持応答待機中.Add ポインター応答, 項目キー
    
    実行結果 = COMメソッド呼出(ポインター応答, 6, "", 通知ハンドラー.スロットアドレス + 仮想COMオブジェクト位置)
    If 実行結果 <> 0 Then
    
        COM参照解放 ポインター応答
        保持応答待機中.Remove 項目キー
        Set 通知ハンドラー = Nothing
    End If
End Sub

' =========================================================================
' 処理概要       : 応答本文の取得結果を受け付けてコールバック状態を管理
' 引数           : バイト数識別番号 (Long) - 非同期コールバックの識別ID
'                  エラーコード (Long) - 非同期処理のHRESULT
'                  ポインターストリーム (LongPtr) - 読み取るCOMストリームのポインター
' =========================================================================
Public Sub 応答本文の取得結果を受け付けてコールバック状態を管理(ByVal バイト数識別番号 As Long, ByVal エラーコード As Long, ByVal ポインターストリーム As LongPtr)
    Dim エラー説明 As String
    Dim エラー番号 As Long
    
    ' コールバック中であることを記録し、同期処理の再入を防ぐ
    保持内部通知 = 保持内部通知 + 1
    On Error GoTo ErrorHanler
    取得した応答本文を対応する通信詳細に保存 バイト数識別番号, エラーコード, ポインターストリーム
    
    GoTo EndProc
    
ErrorHanler:
    エラー番号 = Err.Number
    エラー説明 = Err.Description
    MsgBox "通信イベントの処理でエラーが発生しました。" & vbCrLf & エラー説明, vbExclamation, "ブラウザー"
    Resume EndProc
    
EndProc:

    ' 通知処理の終了時にコールバック階層を戻す
    If 保持内部通知 > 0 Then 保持内部通知 = 保持内部通知 - 1
    If エラー番号 <> 0 Then
        On Error Resume Next
        保持応答本文対応なし件数 = 保持応答本文対応なし件数 + 1
    End If
End Sub

' =========================================================================
' 処理概要       : 取得した応答本文を対応する通信詳細に保存
' 引数           : バイト数識別番号 (Long) - 非同期コールバックの識別ID
'                  エラーコード (Long) - 非同期処理のHRESULT
'                  ポインターストリーム (LongPtr) - 読み取るCOMストリームのポインター
' =========================================================================
Private Sub 取得した応答本文を対応する通信詳細に保存(ByVal バイト数識別番号 As Long, ByVal エラーコード As Long, ByVal ポインターストリーム As LongPtr)
    Dim 参照キー       As Long
    Dim 件数バイト数  As Long
    Dim 打切り   As Boolean
    Dim 作業文字列()   As Byte
    Dim 文字コード名     As String
    Dim ContentType値   As String
    Dim 項目キー      As String
    Dim ポインター保持対象 As LongPtr
    
    項目キー = "r" & バイト数識別番号
    On Error Resume Next
    ポインター保持対象 = CLngPtr(保持応答待機中(項目キー))
    If ポインター保持対象 <> 0 Then
        COM参照解放 ポインター保持対象
        保持応答待機中.Remove 項目キー
    End If
    Err.Clear
    On Error GoTo 0
    
    参照キー = 通信の連番に対応する詳細の記録位置を検索(バイト数識別番号)
    If 参照キー < 0 Then
    
        保持応答本文対応なし件数 = 保持応答本文対応なし件数 + 1
        Exit Sub
    End If
    
    保持詳細応答取得済み(参照キー) = True
    
    If エラーコード = データなし Then
    
        保持詳細応答本文(参照キー) = ""
        保持詳細応答バイト数(参照キー) = 0
        Exit Sub
    ElseIf エラーコード <> 0 Then
        保持詳細応答本文(参照キー) = ""
        保持詳細応答文字コード(参照キー) = "(取得失敗 hr=&H" & Hex(エラーコード) & ")"
        Exit Sub
    End If
    
    If ポインターストリーム = 0 Then
    
        保持詳細応答バイト数(参照キー) = 0
        Exit Sub
    End If
    
    If Not COMストリームを上限バイト数まで読み取り(ポインターストリーム, 保持応答本文最大, 件数バイト数, 打切り, 作業文字列) Then Exit Sub
    
    保持詳細応答バイト数(参照キー) = 件数バイト数
    保持詳細応答打切り(参照キー) = 打切り
    If 件数バイト数 <= 0 Then Exit Sub
    
    文字コード名 = ヘッダー文字列から指定ヘッダーの値を取得(保持詳細応答ヘッダー一覧(参照キー), "content-encoding")
    ContentType値 = ヘッダー文字列から指定ヘッダーの値を取得(保持詳細応答ヘッダー一覧(参照キー), "content-type")
    保持詳細応答文字コード(参照キー) = 文字コード名
    保持詳細応答16進(参照キー) = バイト配列の先頭を16進文字列へ変換(作業文字列, 件数バイト数)
    保持詳細応答判定文字列(参照キー) = ContentTypeがテキスト形式か判定(ContentType値)
    If 保持詳細応答判定文字列(参照キー) Then
        保持詳細応答本文(参照キー) = UTF8のバイト配列をVBA文字列へ変換(作業文字列, 件数バイト数)
    End If
End Sub

' =========================================================================
' 処理概要       : 通信の連番に対応する詳細の記録位置を検索
' 引数           : 連番番号 (Long) - 通信記録の連番
' 戻り値         : Long - 対応する配列添字。該当しない場合 -1
' =========================================================================
Private Function 通信の連番に対応する詳細の記録位置を検索(ByVal 連番番号 As Long) As Long
    Dim 添字 As Long
    Dim 参照キー As Long
    
    通信の連番に対応する詳細の記録位置を検索 = -1
    For 添字 = 保持詳細件数 To 1 Step -1
        参照キー = 通信詳細の表示位置を配列添字に変換(添字)
        If 参照キー >= 0 Then
            If 保持詳細連番(参照キー) = 連番番号 Then
                通信の連番に対応する詳細の記録位置を検索 = 参照キー
                Exit Function
            End If
        End If
    Next 添字
End Function

' =========================================================================
' 処理概要       : 応答本文の取得待ちハンドラーを解放
' =========================================================================
Private Sub 応答本文の取得待ちハンドラーを解放()
    Dim 処理値 As Variant
    
    If 保持応答待機中 Is Nothing Then Exit Sub
    
    On Error Resume Next
    For Each 処理値 In 保持応答待機中
        If CLngPtr(処理値) <> 0 Then COM参照解放 CLngPtr(処理値)
    Next 処理値
    Err.Clear
    On Error GoTo 0
    
    Set 保持応答待機中 = New Collection
End Sub
