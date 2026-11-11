VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "mod_WV2_ブラウザー管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' -------------
' 待機時間
' -------------
Private Const 初期化待機上限秒   As Single = 10#
Private Const 予備生成待機ミリ秒 As Long = 1500

' -------------
' URL設定
' -------------
Private Const 既定検索URL As String = "https://www.google.com/search?q="

Public Event スクリプト実行完了(ByVal 表示 As mod_WV2_ブラウザー表示, ByVal 通知識別番号 As Long, ByVal エラーコード As Long, ByVal 結果JSON As String)

Public Enum ブラウザー状態
    ブラウザー作成済み = 0
    ブラウザー環境準備済み = 1
    ブラウザー初期化失敗 = 2
End Enum

Private 保持環境      As mod_WV2_ブラウザー環境
Private 保持タブ一覧     As Collection
Private 保持ウィンドウハンドルホスト As LongPtr
Private 保持状態    As ブラウザー状態

Private 保持検索テンプレート As String

Private 保持選択位置番号 As Long

Private 保持最終コンテンツ矩形 As 矩形
Private コンテンツ表示領域取得済み  As Boolean

Private WithEvents 保持新規タブタイマー As mod_WV2_安全タイマー
Attribute 保持新規タブタイマー.VB_VarHelpID = -1
Private 保持待機中URL一覧 As Collection

Private 保持静止有効終了 As Boolean

Private WithEvents 保持設定タイマー As mod_WV2_安全タイマー
Attribute 保持設定タイマー.VB_VarHelpID = -1

Private WithEvents 保持同期タイマー As mod_WV2_安全タイマー
Attribute 保持同期タイマー.VB_VarHelpID = -1
Private 保持同期待機中 As Boolean
Private 保持設定待機中 As Boolean
Private 保持設定設定連携 As mod_WV2_検索設定連携

Private 保持新規ウィンドウ方式 As Long
Private 保持設定処理済み付加新規ウィンドウ As Boolean

Private 保持予備生成表示 As mod_WV2_ブラウザー表示
Private 保持予備生成予約済み As Boolean

Private WithEvents 保持予備生成タイマー As mod_WV2_安全タイマー
Attribute 保持予備生成タイマー.VB_VarHelpID = -1
Private 保持予備生成待機中 As Boolean

Private WithEvents 保持取込タイマー As mod_WV2_安全タイマー
Attribute 保持取込タイマー.VB_VarHelpID = -1
Private 保持取込待機中 As Boolean
Private 保持取込URL As String

Private 保持作成中表示 As Long

Public Event タブ追加(ByVal 位置番号 As Long)
Public Event タブ終了(ByVal 位置番号 As Long)
Public Event 選択タブ変更(ByVal 位置番号 As Long)
Public Event タブタイトル変更(ByVal 位置番号 As Long, ByVal タイトル As String)

Public Event タブURL変更(ByVal 位置番号 As Long, ByVal URL As String, ByVal 可能進む戻る As Boolean, ByVal 可能進む進む As Boolean)

' =========================================================================
' 処理概要       : インスタンスの内部状態を初期化
' =========================================================================
Private Sub Class_Initialize()
    Dim 初期化設定連携 As mod_WV2_検索設定連携
    Dim 保存済みテンプレート As String
    
    保持静止有効終了 = True
    Set 保持タブ一覧 = New Collection
    Set 保持環境 = Nothing
    保持ウィンドウハンドルホスト = 0
    保持状態 = ブラウザー作成済み
    保持選択位置番号 = 0
    Set 保持待機中URL一覧 = New Collection
    Set 保持新規タブタイマー = Nothing
    コンテンツ表示領域取得済み = False
    保持検索テンプレート = 既定検索URL
    Set 保持設定タイマー = Nothing
    保持設定待機中 = False
    Set 保持設定設定連携 = Nothing
    
    保持新規ウィンドウ方式 = 2
    保持設定処理済み付加新規ウィンドウ = True
    Set 保持予備生成表示 = Nothing
    保持予備生成予約済み = False
    Set 保持予備生成タイマー = Nothing
    保持予備生成待機中 = False
    Set 保持取込タイマー = Nothing
    保持取込待機中 = False
    保持取込URL = ""
    保持作成中表示 = 0
    
    Set 初期化設定連携 = New mod_WV2_検索設定連携
    保存済みテンプレート = 初期化設定連携.保存済み設定を検索URLテンプレートへ変換()
    If Len(保存済みテンプレート) > 0 Then
        保持検索テンプレート = 保存済みテンプレート
        
    End If
    Set 初期化設定連携 = Nothing
End Sub

' =========================================================================
' 処理概要       : 指定ウィンドウで使用するWebView2環境を初期化
' 引数           : ウィンドウハンドルホスト (LongPtr) - WebView2を配置するウィンドウハンドル
'                  ユーザーデータパス (String) - WebView2プロファイルの保存先
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 初期化(ByVal ウィンドウハンドルホスト As LongPtr, ByVal ユーザーデータパス As String) As Boolean
    If 保持状態 <> ブラウザー作成済み Then
    
        初期化 = (保持状態 = ブラウザー環境準備済み)
        Exit Function
    End If
    
    保持ウィンドウハンドルホスト = ウィンドウハンドルホスト
    
    ' タブで共有するWebView2環境を作成する
    Set 保持環境 = New mod_WV2_ブラウザー環境
    保持環境.初期化 ユーザーデータパス
    
    ' 環境の作成結果を確認してからタブを作成できる状態にする
    If Not WebView2環境の初期化完了を待機() Then
    
        保持状態 = ブラウザー初期化失敗
        Exit Function
    End If
    
    保持状態 = ブラウザー環境準備済み
    
    ' 次回のタブ作成に備えてペインを確保する
    予備ペインの生成を予約
    
    初期化 = True
End Function

' =========================================================================
' 処理概要       : 新しいタブを作成して選択
' 戻り値         : mod_WV2_ブラウザー表示 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Public Function 新しいタブを作成して選択() As mod_WV2_ブラウザー表示
    Dim 表示 As mod_WV2_ブラウザー表示
    
    Set 表示 = ペイン作成中の状態を管理してペインを生成()
    If 表示 Is Nothing Then Exit Function
    
    ' 作成したペインを管理対象に登録する
    保持タブ一覧.Add 表示
    
    RaiseEvent タブ追加(保持タブ一覧.Count)
    
    指定タブを表示して選択状態を更新 保持タブ一覧.Count
    
    予備ペインがない場合に生成を予約
    
    Set 新しいタブを作成して選択 = 表示
End Function

' =========================================================================
' 処理概要       : ペイン作成中の状態を管理してペインを生成
' 戻り値         : mod_WV2_ブラウザー表示 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Private Function ペイン作成中の状態を管理してペインを生成() As mod_WV2_ブラウザー表示
    保持作成中表示 = 保持作成中表示 + 1
    Set ペイン作成中の状態を管理してペインを生成 = コントローラーを作成して標準イベントを登録()
    保持作成中表示 = 保持作成中表示 - 1
    If 保持作成中表示 < 0 Then 保持作成中表示 = 0
End Function

' =========================================================================
' 処理概要       : コントローラーを作成して標準イベントを登録
' 戻り値         : mod_WV2_ブラウザー表示 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Private Function コントローラーを作成して標準イベントを登録() As mod_WV2_ブラウザー表示
    Dim 表示 As mod_WV2_ブラウザー表示
    Dim 表示矩形 As 矩形
    Dim 実行結果 As Long
    
    If 保持状態 <> ブラウザー環境準備済み Then
    
        Exit Function
    End If
    
    Set 表示 = New mod_WV2_ブラウザー表示
    表示.初期化 保持環境, 保持ウィンドウハンドルホスト
    
    If Not ペインの初期化完了を待機(表示) Then
        Exit Function
    End If
    
    ' 保存済みの表示領域があれば再利用する
    If コンテンツ表示領域取得済み Then
        表示矩形 = 保持最終コンテンツ矩形
        
    Else
        クライアント矩形取得API 保持ウィンドウハンドルホスト, 表示矩形
        
    End If
    実行結果 = 表示.コントローラーの表示領域を設定(表示矩形)
    
    実行結果 = 表示.コントローラーからWebView2インターフェースを取得して保持()
    If 実行結果 <> 0 Then
    
        Exit Function
    End If
    If 表示.表示ポインター = 0 Then
    
        Exit Function
    End If
    
    ' 遷移やタイトル変更を受信できるよう、基本イベントを登録する
    ブラウザーの基本イベントを登録 表示
    
    表示.ペインの所有元ブラウザーを設定 Me
    
    Set コントローラーを作成して標準イベントを登録 = 表示
End Function

' =========================================================================
' 処理概要       : ブラウザーの基本イベントを登録
' 引数           : 表示 (mod_WV2_ブラウザー表示) - 対象のブラウザーペイン
' =========================================================================
Private Sub ブラウザーの基本イベントを登録(ByVal 表示 As mod_WV2_ブラウザー表示)
    Dim 実行結果 As Long
    
    実行結果 = 表示.ページ遷移開始イベントを登録()
    
    実行結果 = 表示.ページ遷移完了イベントを登録()
    
    実行結果 = 表示.ページタイトル変更イベントを登録()
    
    実行結果 = 表示.新規ウィンドウ要求イベントを登録()
    
    実行結果 = 表示.履歴変更イベントを登録()
    
End Sub

' =========================================================================
' 処理概要       : SPAの履歴変更とDOM読み込みイベントを登録
' 引数           : 表示 (mod_WV2_ブラウザー表示) - 対象のブラウザーペイン
' =========================================================================
Public Sub SPAの履歴変更とDOM読み込みイベントを登録(ByVal 表示 As mod_WV2_ブラウザー表示)
    Dim 実行結果 As Long
    
    If 表示 Is Nothing Then
    
        Exit Sub
    End If
    
    実行結果 = 表示.DOM読み込み完了イベントを登録()
    
End Sub

' =========================================================================
' 処理概要       : 新しいタブを作成して指定URLを表示
' 引数           : URL (String) - 移動先または通知対象のURL
' 戻り値         : mod_WV2_ブラウザー表示 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Public Function 新しいタブを作成して指定URLを表示(ByVal URL As String) As mod_WV2_ブラウザー表示
    Dim 表示 As mod_WV2_ブラウザー表示
    Dim 実行結果 As Long
    
    Set 表示 = 新しいタブを作成して選択()
    If 表示 Is Nothing Then Exit Function
    
    実行結果 = 表示.指定URLへの移動を開始(URL)
    
    Set 新しいタブを作成して指定URLを表示 = 表示
End Function

' =========================================================================
' 処理概要       : 新しいタブを作成して指定HTMLを表示
' 引数           : HTMLコンテンツ (String) - 表示するHTML文字列
' 戻り値         : mod_WV2_ブラウザー表示 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Public Function 新しいタブを作成して指定HTMLを表示(ByVal HTMLコンテンツ As String) As mod_WV2_ブラウザー表示
    Dim 表示 As mod_WV2_ブラウザー表示
    Dim 実行結果 As Long
    
    Set 表示 = 新しいタブを作成して選択()
    If 表示 Is Nothing Then Exit Function
    
    実行結果 = 表示.指定HTML文字列をページとして表示(HTMLコンテンツ)
    
    Set 新しいタブを作成して指定HTMLを表示 = 表示
End Function

' =========================================================================
' 処理概要       : 仮想ホストを設定してローカルHTMLのタブを作成
' 引数           : ホスト名称 (String) - ローカルフォルダーへ割り当てる仮想ホスト名
'                  フォルダーパス (String) - 対象の保存先またはHTMLフォルダー
'                  ファイル名称 (String) - 表示または保存するHTMLファイル名
' 戻り値         : mod_WV2_ブラウザー表示 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Public Function 仮想ホストを設定してローカルHTMLのタブを作成(ByVal ホスト名称 As String, ByVal フォルダーパス As String, ByVal ファイル名称 As String) As mod_WV2_ブラウザー表示
    Dim 表示 As mod_WV2_ブラウザー表示
    Dim 実行結果 As Long
    Dim URL As String
    
    Set 表示 = 新しいタブを作成して選択()
    If 表示 Is Nothing Then Exit Function
    
    実行結果 = 表示.仮想ホスト名にローカルフォルダーを割り当て(ホスト名称, フォルダーパス, 1)
    If 実行結果 <> 0 Then
    
        Set 仮想ホストを設定してローカルHTMLのタブを作成 = 表示
        Exit Function
    End If
    
    URL = "https://" & ホスト名称 & "/" & ファイル名称
    実行結果 = 表示.指定URLへの移動を開始(URL)
    
    SPAの履歴変更とDOM読み込みイベントを登録 表示
    
    Set 仮想ホストを設定してローカルHTMLのタブを作成 = 表示
End Function

' =========================================================================
' 処理概要       : 保持するタブ数を取得
' 戻り値         : Long - 保持するタブ数
' =========================================================================
Public Property Get タブ数() As Long
    タブ数 = 保持タブ一覧.Count
End Property

' =========================================================================
' 処理概要       : 選択中のペインを取得
' 戻り値         : mod_WV2_ブラウザー表示 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Public Property Get 選択中の表示() As mod_WV2_ブラウザー表示
    If 保持選択位置番号 < 1 Or 保持選択位置番号 > 保持タブ一覧.Count Then
    
        Set 選択中の表示 = Nothing
        Exit Property
    End If
    Set 選択中の表示 = 保持タブ一覧.Item(保持選択位置番号)
End Property

' =========================================================================
' 処理概要       : 共有するWebView2環境を取得
' 戻り値         : mod_WV2_ブラウザー環境 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Public Property Get 共有環境() As mod_WV2_ブラウザー環境
    Set 共有環境 = 保持環境
End Property

' =========================================================================
' 処理概要       : 選択中のタブ番号を取得
' 戻り値         : Long - 選択中のタブ番号
' =========================================================================
Public Property Get 選択タブ番号() As Long
    選択タブ番号 = 保持選択位置番号
End Property

' =========================================================================
' 処理概要       : 指定タブを表示して選択状態を更新
' 引数           : 位置番号 (Long) - 対象の位置番号（公開タブ・ログは1始まり）
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 指定タブを表示して選択状態を更新(ByVal 位置番号 As Long) As Boolean
    Dim 添字 As Long
    Dim 表示 As mod_WV2_ブラウザー表示
    Dim 実行結果 As Long
    
    If 位置番号 < 1 Or 位置番号 > 保持タブ一覧.Count Then
    
        Exit Function
    End If
    
    For 添字 = 1 To 保持タブ一覧.Count
        Set 表示 = 保持タブ一覧.Item(添字)
        実行結果 = 表示.コントローラーの表示状態を設定(添字 = 位置番号)
    Next 添字
    
    保持選択位置番号 = 位置番号
    
    RaiseEvent 選択タブ変更(位置番号)
    指定タブを表示して選択状態を更新 = True
End Function

' =========================================================================
' 処理概要       : 指定タブを閉じて選択状態を更新
' 引数           : 位置番号 (Long) - 対象の位置番号（公開タブ・ログは1始まり）
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 指定タブを閉じて選択状態を更新(ByVal 位置番号 As Long) As Boolean
    Dim 表示 As mod_WV2_ブラウザー表示
    Dim 実行結果 As Long
    Dim 新規選択 As Long
    
    If 位置番号 < 1 Or 位置番号 > 保持タブ一覧.Count Then
    
        Exit Function
    End If
    
    Set 表示 = 保持タブ一覧.Item(位置番号)
    
    表示.所有元ブラウザーへの参照を解除
    
    実行結果 = 表示.イベントとコントローラーを閉じて関連参照を解放()
    
    Set 表示 = Nothing
    保持タブ一覧.Remove 位置番号
    
    If 保持タブ一覧.Count = 0 Then
    
        保持選択位置番号 = 0
        
    ElseIf 位置番号 < 保持選択位置番号 Then
    
        新規選択 = 保持選択位置番号 - 1
        指定タブを表示して選択状態を更新 新規選択
    ElseIf 位置番号 = 保持選択位置番号 Then
    
        If 位置番号 <= 保持タブ一覧.Count Then
            新規選択 = 位置番号
        Else
            新規選択 = 保持タブ一覧.Count
        End If
        指定タブを表示して選択状態を更新 新規選択
    Else
    
        指定タブを表示して選択状態を更新 保持選択位置番号
    End If
    
    RaiseEvent タブ終了(位置番号)
    指定タブを閉じて選択状態を更新 = True
End Function

' =========================================================================
' 処理概要       : 指定タブの並び順を変更
' 引数           : 移動元位置番号 (Long) - 移動元のタブ番号（1始まり）
'                  移動先位置番号 (Long) - 移動先のタブ番号（1始まり）
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 指定タブの並び順を変更(ByVal 移動元位置番号 As Long, ByVal 移動先位置番号 As Long) As Boolean
    Dim 件数 As Long
    Dim 選択表示参照 As mod_WV2_ブラウザー表示
    Dim 移動対象 As mod_WV2_ブラウザー表示
    Dim 添字 As Long
    
    件数 = 保持タブ一覧.Count
    
    If 移動元位置番号 < 1 Or 移動元位置番号 > 件数 Then
    
        Exit Function
    End If
    If 移動先位置番号 < 1 Or 移動先位置番号 > 件数 Then
    
        Exit Function
    End If
    If 移動元位置番号 = 移動先位置番号 Then
    
        指定タブの並び順を変更 = True
        Exit Function
    End If
    
    If 保持選択位置番号 >= 1 And 保持選択位置番号 <= 件数 Then
        Set 選択表示参照 = 保持タブ一覧.Item(保持選択位置番号)
    End If
    
    Set 移動対象 = 保持タブ一覧.Item(移動元位置番号)
    保持タブ一覧.Remove 移動元位置番号
    
    If 移動先位置番号 = 1 Then
    
        保持タブ一覧.Add 移動対象, , 1
    Else
    
        保持タブ一覧.Add 移動対象, , , 移動先位置番号 - 1
    End If
    
    If Not 選択表示参照 Is Nothing Then
        For 添字 = 1 To 保持タブ一覧.Count
            If 保持タブ一覧.Item(添字) Is 選択表示参照 Then
                保持選択位置番号 = 添字
                Exit For
            End If
        Next 添字
    End If
    
    If 保持選択位置番号 >= 1 And 保持選択位置番号 <= 保持タブ一覧.Count Then
        指定タブを表示して選択状態を更新 保持選択位置番号
    End If
    
    指定タブの並び順を変更 = True
End Function

' =========================================================================
' 処理概要       : ペインのタイトル変更をタブのイベントとして通知
' 引数           : 通知元表示 (mod_WV2_ブラウザー表示) - 変更を通知したペイン
'                  タイトル (String) - 通知するページタイトル
' =========================================================================
Public Sub ペインのタイトル変更をタブのイベントとして通知(ByVal 通知元表示 As mod_WV2_ブラウザー表示, ByVal タイトル As String)
    Dim 添字 As Long
    
    For 添字 = 1 To 保持タブ一覧.Count
        If 保持タブ一覧.Item(添字) Is 通知元表示 Then
            RaiseEvent タブタイトル変更(添字, タイトル)
            Exit Sub
        End If
    Next 添字
    
End Sub

' =========================================================================
' 処理概要       : ペインのURLと履歴状態をタブのイベントとして通知
' 引数           : 通知元表示 (mod_WV2_ブラウザー表示) - 変更を通知したペイン
'                  URL (String) - 移動先または通知対象のURL
'                  可能進む戻る (Boolean) - 前の履歴へ戻れる状態
'                  可能進む進む (Boolean) - 次の履歴へ進める状態
' =========================================================================
Public Sub ペインのURLと履歴状態をタブのイベントとして通知(ByVal 通知元表示 As mod_WV2_ブラウザー表示, ByVal URL As String, ByVal 可能進む戻る As Boolean, ByVal 可能進む進む As Boolean)
    Dim 添字 As Long
    
    For 添字 = 1 To 保持タブ一覧.Count
        If 保持タブ一覧.Item(添字) Is 通知元表示 Then
            RaiseEvent タブURL変更(添字, URL, 可能進む戻る, 可能進む進む)
            Exit Sub
        End If
    Next 添字
    
End Sub

' =========================================================================
' 処理概要       : 入力をURLに解決して選択中のタブへ移動
' 引数           : URL (String) - 移動先または通知対象のURL
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 入力をURLに解決して選択中のタブへ移動(ByVal URL As String) As Boolean
    Dim 表示 As mod_WV2_ブラウザー表示
    Dim 対象 As String
    Dim 実行結果 As Long
    
    Set 表示 = 選択中の表示
    If 表示 Is Nothing Then
    
        Exit Function
    End If
    
    対象 = 入力文字列を移動先URLまたは検索URLに変換(URL)
    If Len(対象) = 0 Then
    
        Exit Function
    End If
    
    実行結果 = 表示.指定URLへの移動を開始(対象)
    If 実行結果 <> 0 Then
    
        Exit Function
    End If
    
    入力をURLに解決して選択中のタブへ移動 = True
End Function

' =========================================================================
' 処理概要       : 選択中のタブで前の履歴へ戻る
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 選択中のタブで前の履歴へ戻る() As Boolean
    Dim 表示 As mod_WV2_ブラウザー表示
    Dim 実行結果 As Long
    
    Set 表示 = 選択中の表示
    If 表示 Is Nothing Then Exit Function
    
    If Not 表示.前の履歴へ戻れるか判定() Then
    
        Exit Function
    End If
    
    実行結果 = 表示.前の履歴へ戻る()
    If 実行結果 <> 0 Then
    
        Exit Function
    End If
    
    選択中のタブで前の履歴へ戻る = True
End Function

' =========================================================================
' 処理概要       : 選択中のタブで次の履歴へ進む
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 選択中のタブで次の履歴へ進む() As Boolean
    Dim 表示 As mod_WV2_ブラウザー表示
    Dim 実行結果 As Long
    
    Set 表示 = 選択中の表示
    If 表示 Is Nothing Then Exit Function
    
    If Not 表示.次の履歴へ進めるか判定() Then
    
        Exit Function
    End If
    
    実行結果 = 表示.次の履歴へ進む()
    If 実行結果 <> 0 Then
    
        Exit Function
    End If
    
    選択中のタブで次の履歴へ進む = True
End Function

' =========================================================================
' 処理概要       : 選択中のタブを再読み込み
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 選択中のタブを再読み込み() As Boolean
    Dim 表示 As mod_WV2_ブラウザー表示
    Dim 実行結果 As Long
    
    Set 表示 = 選択中の表示
    If 表示 Is Nothing Then Exit Function
    
    実行結果 = 表示.ページを再読み込み()
    If 実行結果 <> 0 Then
    
        Exit Function
    End If
    
    選択中のタブを再読み込み = True
End Function

' =========================================================================
' 処理概要       : 選択中のタブの読み込みを停止
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 選択中のタブの読み込みを停止() As Boolean
    Dim 表示 As mod_WV2_ブラウザー表示
    Dim 実行結果 As Long
    
    Set 表示 = 選択中の表示
    If 表示 Is Nothing Then Exit Function
    
    実行結果 = 表示.ページの読み込みを停止()
    If 実行結果 <> 0 Then
    
        Exit Function
    End If
    
    選択中のタブの読み込みを停止 = True
End Function

' =========================================================================
' 処理概要       : 選択中のタブのURLを取得
' 戻り値         : String - 選択中のタブのURL
' =========================================================================
Public Function 選択中のタブのURLを取得() As String
    Dim 表示 As mod_WV2_ブラウザー表示
    
    Set 表示 = 選択中の表示
    If 表示 Is Nothing Then Exit Function
    If Not 表示.利用可能 Then Exit Function
    
    選択中のタブのURLを取得 = 表示.表示中のページのURLを取得()
End Function

' =========================================================================
' 処理概要       : 選択中のタブで前の履歴へ戻れるか判定
' 戻り値         : Boolean - 前の履歴へ戻れる場合 True
' =========================================================================
Public Function 選択中のタブで前の履歴へ戻れるか判定() As Boolean
    Dim 表示 As mod_WV2_ブラウザー表示
    
    Set 表示 = 選択中の表示
    If 表示 Is Nothing Then Exit Function
    If Not 表示.利用可能 Then Exit Function
    
    選択中のタブで前の履歴へ戻れるか判定 = 表示.前の履歴へ戻れるか判定()
End Function

' =========================================================================
' 処理概要       : 選択中のタブで次の履歴へ進めるか判定
' 戻り値         : Boolean - 次の履歴へ進める場合 True
' =========================================================================
Public Function 選択中のタブで次の履歴へ進めるか判定() As Boolean
    Dim 表示 As mod_WV2_ブラウザー表示
    
    Set 表示 = 選択中の表示
    If 表示 Is Nothing Then Exit Function
    If Not 表示.利用可能 Then Exit Function
    
    選択中のタブで次の履歴へ進めるか判定 = 表示.次の履歴へ進めるか判定()
End Function

' =========================================================================
' 処理概要       : 入力文字列を移動先URLまたは検索URLに変換
' 引数           : 作成文字列 (String) - 変換対象の文字列
' 戻り値         : String - 入力文字列を移動先URLまたは検索URLに変換
' =========================================================================
Private Function 入力文字列を移動先URLまたは検索URLに変換(ByVal 作成文字列 As String) As String
    Dim 開始時刻 As String
    Dim コロン位置 As Long
    Dim 先頭 As String
    
    開始時刻 = Trim$(作成文字列)
    If Len(開始時刻) = 0 Then Exit Function
    
    If InStr(開始時刻, "://") > 0 Then
        入力文字列を移動先URLまたは検索URLに変換 = 開始時刻
        Exit Function
    End If
    
    コロン位置 = InStr(開始時刻, ":")
    If コロン位置 > 1 Then
        先頭 = Left$(開始時刻, コロン位置 - 1)
        
        If mod_WV2_URL共通.文字列が英字のみで構成されるか判定(先頭) Then
            入力文字列を移動先URLまたは検索URLに変換 = 開始時刻
            Exit Function
        End If
    End If
    
    If InStr(開始時刻, ".") > 0 And InStr(開始時刻, " ") = 0 Then
        入力文字列を移動先URLまたは検索URLに変換 = "https://" & 開始時刻
    Else
        入力文字列を移動先URLまたは検索URLに変換 = 検索語をエンコードして検索URLを生成(開始時刻)
    End If
End Function

' =========================================================================
' 処理概要       : 現在の検索設定を使って入力文字列をURLに変換
' 引数           : 作成文字列 (String) - 変換対象の文字列
' 戻り値         : String - 現在の検索設定を使って入力文字列をURLに変換
' =========================================================================
Public Function 移動先URLを解決(ByVal 作成文字列 As String) As String
    移動先URLを解決 = 入力文字列を移動先URLまたは検索URLに変換(作成文字列)
End Function

' =========================================================================
' 処理概要       : 検索語をエンコードして検索URLを生成
' 引数           : 検索語 (String) - URLに反映する検索語
' 戻り値         : String - 検索語をエンコードして検索URL
' =========================================================================
Private Function 検索語をエンコードして検索URLを生成(ByVal 検索語 As String) As String
    検索語をエンコードして検索URLを生成 = 保持検索テンプレート & mod_WV2_URL共通.文字列をUTF8でURLエンコード(検索語)
End Function

' =========================================================================
' 処理概要       : 新規ウィンドウ要求の処理方式を取得
' 戻り値         : Long - 新規ウィンドウ要求の処理方式
' =========================================================================
Public Property Get 新規ウィンドウ方式() As Long
    新規ウィンドウ方式 = 保持新規ウィンドウ方式
End Property

' =========================================================================
' 処理概要       : 新規ウィンドウ要求の処理方式を設定
' 引数           : 入力値 (Long) - 設定する値
' =========================================================================
Public Property Let 新規ウィンドウ方式(ByVal 入力値 As Long)
    Dim 処理値 As Long
    
    処理値 = 入力値
    If 処理値 < 0 Or 処理値 > 2 Then
    
        処理値 = 0
    End If
    保持新規ウィンドウ方式 = 処理値
    
    Select Case 処理値
        Case 0
        
        Case 1
        
        Case 2
        
            If Not 事前生成したペインを利用できるか判定 Then
            
                ' 次回のタブ作成に備えてペインを確保する
                予備ペインの生成を予約
            End If
    End Select
End Property

' =========================================================================
' 処理概要       : 新規ウィンドウの割り当て時に処理済みとする設定を取得
' 戻り値         : Boolean - 有効時 True / 無効時 False
' =========================================================================
Public Property Get 新規ウィンドウ割当後に処理済みにする() As Boolean
    新規ウィンドウ割当後に処理済みにする = 保持設定処理済み付加新規ウィンドウ
End Property

' =========================================================================
' 処理概要       : 新規ウィンドウの割り当て時に処理済みとする設定を更新
' 引数           : 入力値 (Boolean) - 設定する値
' =========================================================================
Public Property Let 新規ウィンドウ割当後に処理済みにする(ByVal 入力値 As Boolean)
    保持設定処理済み付加新規ウィンドウ = 入力値
End Property

' =========================================================================
' 処理概要       : 新規ウィンドウ要求をWebView2へ委譲する設定を取得
' 戻り値         : Boolean - 有効時 True / 無効時 False
' =========================================================================
Public Property Get ポップアップ委譲() As Boolean
    ポップアップ委譲 = (保持新規ウィンドウ方式 = 1)
End Property

' =========================================================================
' 処理概要       : 新規ウィンドウ要求をWebView2へ委譲する設定を更新
' 引数           : 入力値 (Boolean) - 設定する値
' =========================================================================
Public Property Let ポップアップ委譲(ByVal 入力値 As Boolean)
    If 入力値 Then
        新規ウィンドウ方式 = 1
    Else
        新規ウィンドウ方式 = 0
    End If
End Property

' =========================================================================
' 処理概要       : 検索URLのテンプレートを取得
' 戻り値         : String - 検索URLのテンプレート
' =========================================================================
Public Property Get 検索URL() As String
    検索URL = 保持検索テンプレート
End Property

' =========================================================================
' 処理概要       : 検索URLのテンプレートを設定
' 引数           : テンプレート (String) - 検索語を末尾に連結するURLテンプレート
' =========================================================================
Public Property Let 検索URL(ByVal テンプレート As String)
    Dim 開始時刻 As String
    
    開始時刻 = Trim$(テンプレート)
    If Len(開始時刻) = 0 Then
        保持検索テンプレート = 既定検索URL
        
    Else
        保持検索テンプレート = 開始時刻
        
    End If
    
    設定画面の同期をタイマー経由で予約
End Property

' =========================================================================
' 処理概要       : 全タブの表示領域を更新
' 引数           : 表示矩形 (矩形) - コンテンツの表示領域（ピクセル）
' =========================================================================
Public Sub 全タブの表示領域を更新(ByRef 表示矩形 As 矩形)
    Dim 添字 As Long
    Dim 表示 As mod_WV2_ブラウザー表示
    Dim 実行結果 As Long
    
    保持最終コンテンツ矩形 = 表示矩形
    コンテンツ表示領域取得済み = True
    
    For 添字 = 1 To 保持タブ一覧.Count
        Set 表示 = 保持タブ一覧.Item(添字)
        If 表示.利用可能 Then
            実行結果 = 表示.コントローラーの表示領域を設定(表示矩形)
        End If
    Next 添字
End Sub

' =========================================================================
' 処理概要       : 指定タブのペインを取得
' 引数           : 位置番号 (Long) - 対象の位置番号（公開タブ・ログは1始まり）
' 戻り値         : mod_WV2_ブラウザー表示 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Public Property Get 指定タブのペインを取得(ByVal 位置番号 As Long) As mod_WV2_ブラウザー表示
    If 位置番号 < 1 Or 位置番号 > 保持タブ一覧.Count Then
        Set 指定タブのペインを取得 = Nothing
        Exit Property
    End If
    Set 指定タブのペインを取得 = 保持タブ一覧.Item(位置番号)
End Property

' =========================================================================
' 処理概要       : 初期化処理の状態を取得
' 戻り値         : ブラウザー状態 - 初期化処理の状態
' =========================================================================
Public Property Get 状態() As ブラウザー状態
    状態 = 保持状態
End Property

' =========================================================================
' 処理概要       : WebView2環境の初期化完了を待機
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Private Function WebView2環境の初期化完了を待機() As Boolean
    Dim 開始時刻 As Single
    
    開始時刻 = Timer
    Do While (Not 保持環境.利用可能) And (Not 保持環境.初期化失敗) And ((Timer - 開始時刻) < 初期化待機上限秒)
        DoEvents
    Loop
    WebView2環境の初期化完了を待機 = 保持環境.利用可能
End Function

' =========================================================================
' 処理概要       : ペインの初期化完了を待機
' 引数           : 表示 (mod_WV2_ブラウザー表示) - 対象のブラウザーペイン
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Private Function ペインの初期化完了を待機(ByVal 表示 As mod_WV2_ブラウザー表示) As Boolean
    Dim 開始時刻 As Single
    
    開始時刻 = Timer
    Do While (Not 表示.利用可能) And (Not 表示.初期化失敗) And ((Timer - 開始時刻) < 初期化待機上限秒)
        DoEvents
    Loop
    ペインの初期化完了を待機 = 表示.利用可能
End Function

' =========================================================================
' 処理概要       : 新規タブの作成をタイマー経由で予約
' 引数           : URL (String) - 移動先または通知対象のURL
' =========================================================================
Public Sub 新規タブの作成をタイマー経由で予約(ByVal URL As String)
    If 保持待機中URL一覧 Is Nothing Then Set 保持待機中URL一覧 = New Collection
    保持待機中URL一覧.Add URL
    
    Set 保持新規タブタイマー = New mod_WV2_安全タイマー
    保持新規タブタイマー.指定間隔のタイマーを開始または停止 1
    
End Sub

' =========================================================================
' 処理概要       : 設定タブの表示をタイマー経由で予約
' =========================================================================
Public Sub 設定タブの表示をタイマー経由で予約()
    保持設定待機中 = True
    
    Set 保持設定タイマー = New mod_WV2_安全タイマー
    保持設定タイマー.指定間隔のタイマーを開始または停止 1
    
End Sub

' =========================================================================
' 処理概要       : 予約済みの設定タブ表示を実行
' 引数           : 通知時刻 (Long) - タイマー通知時刻
'                  任意データ (Variant) - タイマー通知へ引き渡すデータ
' =========================================================================
Private Sub 保持設定タイマー_タイマー発生(ByVal 通知時刻 As Long, ByRef 任意データ As Variant)
    If Not 保持設定タイマー Is Nothing Then 保持設定タイマー.指定間隔のタイマーを開始または停止 0
    
    If Not 保持設定待機中 Then Exit Sub
    
    保持設定待機中 = False
    
    既存の設定タブを選択または新規作成
End Sub

' =========================================================================
' 処理概要       : 指定したWebView2が設定タブか判定
' 引数           : 表示 (mod_WV2_ブラウザー表示) - 対象のブラウザーペイン
' 戻り値         : Boolean - 設定タブに一致した場合 True
' =========================================================================
Private Function 指定したWebView2が設定タブか判定(ByVal 表示 As mod_WV2_ブラウザー表示) As Boolean
    Dim 実行スクリプト As String
    Dim 実行結果値 As String
    
    If 表示 Is Nothing Then Exit Function
    
    実行スクリプト = "(function(){try{return {s:(window.__wv2Settings===true)?'yes':'no'};}" & "catch(e){return {s:'no'};}})()"
    
    実行結果値 = 表示.JavaScript式を評価して完了まで待機(実行スクリプト, 2)
    
    If Not 表示.直前の同期JavaScript評価が成功したか取得 Then
    
        Exit Function
    End If
    
    指定したWebView2が設定タブか判定 = (InStr(1, 実行結果値, """s"":""yes""") > 0)
End Function

' =========================================================================
' 処理概要       : 設定画面の同期をタイマー経由で予約
' =========================================================================
Public Sub 設定画面の同期をタイマー経由で予約()
    保持同期待機中 = True
    
    Set 保持同期タイマー = New mod_WV2_安全タイマー
    保持同期タイマー.指定間隔のタイマーを開始または停止 1
End Sub

' =========================================================================
' 処理概要       : 予約済みの設定画面同期を実行
' 引数           : 通知時刻 (Long) - タイマー通知時刻
'                  任意データ (Variant) - タイマー通知へ引き渡すデータ
' =========================================================================
Private Sub 保持同期タイマー_タイマー発生(ByVal 通知時刻 As Long, ByRef 任意データ As Variant)
    If Not 保持同期タイマー Is Nothing Then 保持同期タイマー.指定間隔のタイマーを開始または停止 0
    
    If Not 保持同期待機中 Then Exit Sub
    
    保持同期待機中 = False
    
    開いている設定タブへ現在の設定を反映
End Sub

' =========================================================================
' 処理概要       : 開いている設定タブへ現在の設定を反映
' =========================================================================
Private Sub 開いている設定タブへ現在の設定を反映()
    Dim 同期名称 As String
    Dim 同期添字 As Long
    Dim 同期表示 As mod_WV2_ブラウザー表示
    Dim 同期JavaScript As String
    Dim 実行結果値 As String
    Dim 一致位置 As Long
    
    If 保持タブ一覧 Is Nothing Then Exit Sub
    If 保持タブ一覧.Count = 0 Then Exit Sub
    
    If 保持設定設定連携 Is Nothing Then Exit Sub
    
    同期名称 = 保持設定設定連携.現在の検索エンジン名を取得()
    If Len(同期名称) = 0 Then Exit Sub
    
    For 同期添字 = 1 To 保持タブ一覧.Count
        Set 同期表示 = 保持タブ一覧.Item(同期添字)
        If 指定したWebView2が設定タブか判定(同期表示) Then
            同期JavaScript = "(function(){try{return {r:applyFromHost(" & 同期表示.JavaScript文字列に変換(同期名称) & ")};}catch(e){return {r:'err'};}})()"
            実行結果値 = 同期表示.JavaScript式を評価して完了まで待機(同期JavaScript, 2)
            If 同期表示.直前の同期JavaScript評価が成功したか取得 Then
                If InStr(1, 実行結果値, """r"":""ok""") > 0 Then
                    一致位置 = 一致位置 + 1
                Else
                    mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー管理.SyncSettingsViewsNow: index " & 同期添字 & " の適用が失敗 (" & 実行結果値 & ")"
                End If
            Else
                mod_WV2_ログ管理.警告レベルのログを記録 "ブラウザー管理.SyncSettingsViewsNow: index " & 同期添字 & " へ届かなかった (" & 同期表示.直前の同期JavaScript評価のエラーを取得 & ")"
            End If
        End If
    Next 同期添字
    
End Sub

' =========================================================================
' 処理概要       : 既存の設定タブを選択または新規作成
' =========================================================================
Public Sub 既存の設定タブを選択または新規作成()
    Dim 副添字 As Long
    Dim 表示 As mod_WV2_ブラウザー表示
    Dim 実行結果 As Long
    Dim 添字 As Long
    
    For 副添字 = 1 To 保持タブ一覧.Count
        If 指定したWebView2が設定タブか判定(保持タブ一覧.Item(副添字)) Then
        
            指定タブを表示して選択状態を更新 副添字
            Exit Sub
        End If
    Next 副添字
    
    Set 表示 = 新しいタブを作成して選択()
    If 表示 Is Nothing Then
    
        Exit Sub
    End If
    
    表示.設定画面 = True
    
    Set 保持設定設定連携 = New mod_WV2_検索設定連携
    保持設定設定連携.設定の読み書き対象となるブラウザーを設定 Me
    
    実行結果 = 表示.VBAオブジェクトをJavaScriptへ公開("Settings", 保持設定設定連携)
    
    実行結果 = 表示.指定HTML文字列をページとして表示(検索エンジン設定画面のHTMLを生成())
    
    For 添字 = 1 To 保持タブ一覧.Count
        If 保持タブ一覧.Item(添字) Is 表示 Then
            指定タブを表示して選択状態を更新 添字
            Exit For
        End If
    Next 添字
End Sub

' =========================================================================
' 処理概要       : 検索エンジン設定画面のHTMLを生成
' 戻り値         : String - 検索エンジン設定画面のHTML
' =========================================================================
Private Function 検索エンジン設定画面のHTMLを生成() As String
    Dim 作成文字列 As String
    
    作成文字列 = "<!DOCTYPE html>" & vbLf
    作成文字列 = 作成文字列 & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    作成文字列 = 作成文字列 & "<title>設定</title>" & vbLf
    作成文字列 = 作成文字列 & "<style>" & vbLf
    作成文字列 = 作成文字列 & "  *{box-sizing:border-box;margin:0;padding:0;}" & vbLf
    作成文字列 = 作成文字列 & "  html,body{height:100%;font-family:'Segoe UI','Meiryo',sans-serif;" & "color:#e8eaed;background:radial-gradient(1200px 600px at 20% -10%,#2b3550 0%,#151a26 55%,#0e121b 100%);}" & vbLf
    作成文字列 = 作成文字列 & "  .wrap{max-width:840px;margin:0 auto;padding:48px 28px 64px;}" & vbLf
    作成文字列 = 作成文字列 & "  .eyebrow{letter-spacing:.22em;font-size:11px;color:#8ea2c8;text-transform:uppercase;}" & vbLf
    作成文字列 = 作成文字列 & "  h1{font-size:26px;font-weight:650;margin:6px 0 4px;}" & vbLf
    作成文字列 = 作成文字列 & "  .lead{color:#9aa7bd;font-size:13.5px;margin-bottom:28px;}" & vbLf
    作成文字列 = 作成文字列 & "  .grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:16px;}" & vbLf
    作成文字列 = 作成文字列 & "  .card{position:relative;border:1px solid rgba(255,255,255,.08);border-radius:14px;" & "padding:18px 18px 16px;cursor:pointer;background:rgba(255,255,255,.03);" & _
        "transition:transform .12s ease,border-color .12s ease,background .12s ease;overflow:hidden;}" & vbLf
    作成文字列 = 作成文字列 & "  .card:hover{transform:translateY(-2px);border-color:rgba(120,160,255,.5);background:rgba(255,255,255,.06);}" & vbLf
    作成文字列 = 作成文字列 & "  .card.sel{border-color:#6ea8fe;background:linear-gradient(180deg,rgba(110,168,254,.16),rgba(110,168,254,.05));" & _
        "box-shadow:0 0 0 1px rgba(110,168,254,.35),0 10px 30px -12px rgba(110,168,254,.5);}" & vbLf
    作成文字列 = 作成文字列 & "  .card .dot{position:absolute;top:16px;right:16px;width:16px;height:16px;border-radius:50%;" & "border:2px solid rgba(255,255,255,.3);}" & vbLf
    作成文字列 = 作成文字列 & "  .card.sel .dot{border-color:#6ea8fe;background:#6ea8fe;box-shadow:0 0 10px #6ea8fe;}" & vbLf
    作成文字列 = 作成文字列 & "  .card .name{font-size:17px;font-weight:620;margin-bottom:2px;}" & vbLf
    作成文字列 = 作成文字列 & "  .card .host{font-size:11.5px;color:#8ea2c8;word-break:break-all;}" & vbLf
    作成文字列 = 作成文字列 & "  .badge{display:inline-block;margin-top:10px;font-size:10.5px;letter-spacing:.08em;" & "color:#6ea8fe;border:1px solid rgba(110,168,254,.4);border-radius:999px;padding:2px 9px;opacity:0;" & _
        "transition:opacity .12s ease;}" & vbLf
    作成文字列 = 作成文字列 & "  .card.sel .badge{opacity:1;}" & vbLf
    作成文字列 = 作成文字列 & "  .preview{margin-top:26px;border-top:1px solid rgba(255,255,255,.08);padding-top:16px;}" & vbLf
    作成文字列 = 作成文字列 & "  .preview .plabel{font-size:11px;letter-spacing:.14em;color:#8ea2c8;text-transform:uppercase;margin-bottom:6px;}" & vbLf
    作成文字列 = 作成文字列 & "  .preview .purl{font-family:'Consolas','Courier New',monospace;font-size:12.5px;color:#cfe0ff;" & _
        "background:rgba(0,0,0,.25);border-radius:8px;padding:10px 12px;word-break:break-all;min-height:38px;}" & vbLf
    作成文字列 = 作成文字列 & "  .sample{color:#7d8aa0;}" & vbLf
    
    作成文字列 = 作成文字列 & "  .custom{margin-top:22px;border:1px solid rgba(255,255,255,.08);border-radius:14px;" & "padding:16px 18px;background:rgba(255,255,255,.03);}" & vbLf
    作成文字列 = 作成文字列 & "  .custom.sel{border-color:#6ea8fe;" & "background:linear-gradient(180deg,rgba(110,168,254,.16),rgba(110,168,254,.05));" & "box-shadow:0 0 0 1px rgba(110,168,254,.35);}" & vbLf
    作成文字列 = 作成文字列 & "  .chead{display:flex;align-items:center;gap:10px;}" & vbLf
    作成文字列 = 作成文字列 & "  .cname{font-size:15px;font-weight:620;}" & vbLf
    作成文字列 = 作成文字列 & "  .cbadge{font-size:10.5px;letter-spacing:.08em;color:#6ea8fe;" & "border:1px solid rgba(110,168,254,.4);border-radius:999px;padding:2px 9px;" & "opacity:0;transition:opacity .12s ease;}" & vbLf
    作成文字列 = 作成文字列 & "  .custom.sel .cbadge{opacity:1;}" & vbLf
    作成文字列 = 作成文字列 & "  .chint{font-size:11.5px;color:#8ea2c8;margin:6px 0 12px;}" & vbLf
    作成文字列 = 作成文字列 & "  .crow{display:flex;gap:10px;}" & vbLf
    作成文字列 = 作成文字列 & "  .crow input{flex:1;min-width:0;font-family:'Consolas','Courier New',monospace;" & "font-size:12.5px;color:#e8eaed;background:rgba(0,0,0,.25);" & _
        "border:1px solid rgba(255,255,255,.12);border-radius:8px;padding:9px 11px;outline:none;}" & vbLf
    作成文字列 = 作成文字列 & "  .crow input:focus{border-color:#6ea8fe;}" & vbLf
    作成文字列 = 作成文字列 & "  .crow button{font-size:13px;color:#0e121b;background:#6ea8fe;border:0;" & "border-radius:8px;padding:9px 18px;cursor:pointer;font-weight:620;}" & vbLf
    作成文字列 = 作成文字列 & "  .crow button:hover{background:#8fbcff;}" & vbLf
    作成文字列 = 作成文字列 & "  .cmsg{margin-top:10px;font-family:'Consolas','Courier New',monospace;font-size:12px;" & "color:#a9c6ff;word-break:break-all;min-height:16px;}" & vbLf
    作成文字列 = 作成文字列 & "  .cmsg.err{color:#ff9d9d;}" & vbLf
    作成文字列 = 作成文字列 & "</style></head><body>" & vbLf
    作成文字列 = 作成文字列 & "<div class=""wrap"">" & vbLf
    作成文字列 = 作成文字列 & "  <div class=""eyebrow"">Search Engine</div>" & vbLf
    作成文字列 = 作成文字列 & "  <h1>検索エンジンを選ぶ</h1>" & vbLf
    作成文字列 = 作成文字列 & "  <div class=""lead"">アドレスバーに URL でない言葉を入力したとき、どの検索エンジンに送るかを選びます。" & "カードにマウスを乗せると、生成される検索 URL をその場で確認できます。</div>" & vbLf
    作成文字列 = 作成文字列 & "  <div class=""grid"" id=""grid""></div>" & vbLf
    
    作成文字列 = 作成文字列 & "  <div class=""custom"" id=""custom"">" & vbLf
    作成文字列 = 作成文字列 & "    <div class=""chead""><span class=""cname"">カスタム (プリセット以外)</span>" & "<span class=""cbadge"">使用中</span></div>" & vbLf
    作成文字列 = 作成文字列 & "    <div class=""chint"">検索語を末尾に付け足す形の URL を入力します。" & "例: https://example.com/search?q=</div>" & vbLf
    作成文字列 = 作成文字列 & "    <div class=""crow"">" & vbLf
    作成文字列 = 作成文字列 & "      <input id=""ctpl"" type=""text"" spellcheck=""false"" " & "placeholder=""https://example.com/search?q="">" & vbLf
    作成文字列 = 作成文字列 & "      <button id=""capply"">適用</button>" & vbLf
    作成文字列 = 作成文字列 & "    </div>" & vbLf
    作成文字列 = 作成文字列 & "    <div class=""cmsg"" id=""cmsg""></div>" & vbLf
    作成文字列 = 作成文字列 & "  </div>" & vbLf
    作成文字列 = 作成文字列 & "  <div class=""preview"">" & vbLf
    作成文字列 = 作成文字列 & "    <div class=""plabel"">Preview</div>" & vbLf
    作成文字列 = 作成文字列 & "    <div class=""purl"" id=""purl""><span class=""sample"">現在のエンジンでのプレビューを表示します</span></div>" & vbLf
    作成文字列 = 作成文字列 & "  </div>" & vbLf
    作成文字列 = 作成文字列 & "</div>" & vbLf
    作成文字列 = 作成文字列 & "<script>" & vbLf
    
    作成文字列 = 作成文字列 & "  var S = (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects" & " && window.chrome.webview.hostObjects.sync) ? window.chrome.webview.hostObjects.sync.Settings : null;" & _
        vbLf
    作成文字列 = 作成文字列 & "  var SAMPLE = '日本 天気';" & vbLf
    作成文字列 = 作成文字列 & "  var current = '';" & vbLf
    
    作成文字列 = 作成文字列 & "  var hoverId = null;" & vbLf
    
    作成文字列 = 作成文字列 & "  window.__wv2Settings = true;" & vbLf
    
    作成文字列 = 作成文字列 & "  function loadState(){" & vbLf
    作成文字列 = 作成文字列 & "    if(!S){ document.getElementById('grid').textContent='ホストオブジェクト未接続'; return null; }" & vbLf
    作成文字列 = 作成文字列 & "    try{ current = S.現在の検索エンジン名を取得(); }catch(e){ current=''; }" & vbLf
    作成文字列 = 作成文字列 & "    try{ return JSON.parse(S.検索エンジンのプリセット一覧をJSONで取得()); }catch(e){ return []; }" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    
    作成文字列 = 作成文字列 & "  function hostOf(tpl){ try{ return (tpl.split('/')[2]||tpl); }catch(e){ return tpl; } }" & vbLf
    作成文字列 = 作成文字列 & "  function markSel(){" & vbLf
    作成文字列 = 作成文字列 & "    var cs=document.querySelectorAll('.card');" & vbLf
    作成文字列 = 作成文字列 & "    for(var i=0;i<cs.length;i++){ cs[i].classList.toggle('sel', cs[i].getAttribute('data-id')===current); }" & vbLf
    
    作成文字列 = 作成文字列 & "    markCustom();" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    作成文字列 = 作成文字列 & "  function showPreview(id){" & vbLf
    作成文字列 = 作成文字列 & "    var el=document.getElementById('purl'); if(!S){ return; }" & vbLf
    作成文字列 = 作成文字列 & "    var u=''; try{ u=S.指定プリセットで検索URLを生成(id, SAMPLE); }catch(e){ u=''; }" & vbLf
    作成文字列 = 作成文字列 & "    el.innerHTML = '<span class=""sample"">' + SAMPLE + '</span> &nbsp;&rarr;&nbsp; ' + (u||'(取得失敗)');" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    
    作成文字列 = 作成文字列 & "  function showCurrentPreview(){" & vbLf
    作成文字列 = 作成文字列 & "    if(!S){ return; }" & vbLf
    作成文字列 = 作成文字列 & "    if(current){ showPreview(current); }" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    作成文字列 = 作成文字列 & "  function build(list){" & vbLf
    作成文字列 = 作成文字列 & "    var g=document.getElementById('grid'); g.innerHTML='';" & vbLf
    作成文字列 = 作成文字列 & "    for(var i=0;i<list.length;i++){ (function(item){" & vbLf
    作成文字列 = 作成文字列 & "      var c=document.createElement('div'); c.className='card'; c.setAttribute('data-id',item.id);" & vbLf
    作成文字列 = 作成文字列 & "      c.innerHTML='<span class=""dot""></span>' +" & vbLf
    作成文字列 = 作成文字列 & "        '<div class=""name"">'+item.label+'</div>' +" & vbLf
    作成文字列 = 作成文字列 & "        '<div class=""host"">'+hostOf(item.template)+'</div>' +" & vbLf
    作成文字列 = 作成文字列 & "        '<div class=""badge"">選択中</div>';" & vbLf
    
    作成文字列 = 作成文字列 & "      c.onmouseenter=function(){ hoverId=item.id; showPreview(item.id); };" & vbLf
    
    作成文字列 = 作成文字列 & "      c.onmouseleave=function(){ hoverId=null; setTimeout(function(){ if(hoverId===null){ showCurrentPreview(); } },0); };" & vbLf
    作成文字列 = 作成文字列 & "      c.onclick=function(){" & vbLf
    作成文字列 = 作成文字列 & "        if(!S) return;" & vbLf
    作成文字列 = 作成文字列 & "        var res=''; try{ res=S.指定プリセットの検索エンジンを適用して保存(item.id); }catch(e){ res=''; }" & vbLf
    作成文字列 = 作成文字列 & "        if(res){ current=res; } else { current=item.id; }" & vbLf
    作成文字列 = 作成文字列 & "        markSel(); showPreview(item.id);" & vbLf
    作成文字列 = 作成文字列 & "      };" & vbLf
    作成文字列 = 作成文字列 & "      g.appendChild(c);" & vbLf
    作成文字列 = 作成文字列 & "    })(list[i]); }" & vbLf
    作成文字列 = 作成文字列 & "    markSel();" & vbLf
    
    作成文字列 = 作成文字列 & "    showCurrentPreview();" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    
    作成文字列 = 作成文字列 & "  function markCustom(){" & vbLf
    作成文字列 = 作成文字列 & "    var box=document.getElementById('custom'); if(!box){ return; }" & vbLf
    作成文字列 = 作成文字列 & "    box.classList.toggle('sel', current==='custom');" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    作成文字列 = 作成文字列 & "  function setMsg(text, isErr){" & vbLf
    作成文字列 = 作成文字列 & "    var m=document.getElementById('cmsg'); if(!m){ return; }" & vbLf
    作成文字列 = 作成文字列 & "    m.textContent = text || '';" & vbLf
    作成文字列 = 作成文字列 & "    m.className = isErr ? 'cmsg err' : 'cmsg';" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    
    作成文字列 = 作成文字列 & "  function customPreview(){" & vbLf
    作成文字列 = 作成文字列 & "    var ct=document.getElementById('ctpl'); if(!ct||!S){ return; }" & vbLf
    作成文字列 = 作成文字列 & "    var v=ct.value;" & vbLf
    作成文字列 = 作成文字列 & "    if(!v){ setMsg('', false); return; }" & vbLf
    作成文字列 = 作成文字列 & "    var u=''; try{ u=S.指定テンプレートで検索URLを生成(v, SAMPLE); }catch(e){ u=''; }" & vbLf
    作成文字列 = 作成文字列 & "    if(u){ setMsg(SAMPLE + '  →  ' + u, false); }" & vbLf
    作成文字列 = 作成文字列 & "    else { setMsg('http:// または https:// で始まる URL を入力してください', true); }" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    作成文字列 = 作成文字列 & "  function applyCustom(){" & vbLf
    作成文字列 = 作成文字列 & "    var ct=document.getElementById('ctpl'); if(!ct||!S){ return; }" & vbLf
    作成文字列 = 作成文字列 & "    var res=''; try{ res=S.独自の検索URLを適用して保存(ct.value); }catch(e){ res=''; }" & vbLf
    作成文字列 = 作成文字列 & "    if(!res){ setMsg('適用できませんでした。http:// または https:// で始まる URL を" & "入力してください', true); return; }" & vbLf
    作成文字列 = 作成文字列 & "    current=res; markSel(); showCurrentPreview();" & vbLf
    作成文字列 = 作成文字列 & "    if(res==='custom'){ setMsg('適用しました (カスタム)', false); }" & vbLf
    作成文字列 = 作成文字列 & "    else { setMsg('入力がプリセット (' + res + ') と同じだったので、" & "そちらを選択しました', false); }" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    作成文字列 = 作成文字列 & "  function initCustom(){" & vbLf
    作成文字列 = 作成文字列 & "    var ct=document.getElementById('ctpl'); if(!ct){ return; }" & vbLf
    作成文字列 = 作成文字列 & "    if(S){ try{ ct.value = S.独自の検索URLテンプレートを取得(); }catch(e){ } }" & vbLf
    作成文字列 = 作成文字列 & "    ct.oninput = function(){ customPreview(); };" & vbLf
    作成文字列 = 作成文字列 & "    ct.onkeydown = function(e){ if(e.key==='Enter'){ applyCustom(); } };" & vbLf
    作成文字列 = 作成文字列 & "    var btn=document.getElementById('capply');" & vbLf
    作成文字列 = 作成文字列 & "    if(btn){ btn.onclick=function(){ applyCustom(); }; }" & vbLf
    作成文字列 = 作成文字列 & "    markCustom();" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    
    作成文字列 = 作成文字列 & "  function applyFromHost(name){" & vbLf
    作成文字列 = 作成文字列 & "    if(!name){ return 'ng'; }" & vbLf
    作成文字列 = 作成文字列 & "    current = name;" & vbLf
    作成文字列 = 作成文字列 & "    markSel();" & vbLf
    作成文字列 = 作成文字列 & "    showCurrentPreview();" & vbLf
    作成文字列 = 作成文字列 & "    return 'ok';" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    作成文字列 = 作成文字列 & "  var list=loadState(); if(list){ build(list); }" & vbLf
    作成文字列 = 作成文字列 & "  initCustom();" & vbLf
    作成文字列 = 作成文字列 & "</script>" & vbLf
    作成文字列 = 作成文字列 & "</body></html>"
    検索エンジン設定画面のHTMLを生成 = 作成文字列
End Function

' =========================================================================
' 処理概要       : 事前生成したペインを利用できるか判定
' 戻り値         : Boolean - 予備ペインを利用できる場合 True
' =========================================================================
Public Property Get 事前生成したペインを利用できるか判定() As Boolean
    If 保持予備生成表示 Is Nothing Then Exit Property
    If 保持予備生成予約済み Then Exit Property
    If Not 保持予備生成表示.利用可能 Then Exit Property
    
    事前生成したペインを利用できるか判定 = (保持予備生成表示.表示ポインター <> 0)
End Property

' =========================================================================
' 処理概要       : 事前生成したペインを予約してWebView2ポインターを取得
' 戻り値         : LongPtr - 対象のアドレス。取得できない場合 0
' =========================================================================
Public Function 事前生成したペインを予約してWebView2ポインターを取得() As LongPtr
    Dim 引数領域 As LongPtr
    
    If Not 事前生成したペインを利用できるか判定 Then
    
        Exit Function
    End If
    
    引数領域 = 保持予備生成表示.表示ポインター
    保持予備生成予約済み = True
    事前生成したペインを予約してWebView2ポインターを取得 = 引数領域
    
End Function

' =========================================================================
' 処理概要       : 事前生成ペインの利用予約を解除
' =========================================================================
Public Sub 事前生成ペインの利用予約を解除()
    If Not 保持予備生成予約済み Then Exit Sub
    
    保持予備生成予約済み = False
    
End Sub

' =========================================================================
' 処理概要       : 予備ペインの生成を予約
' =========================================================================
Public Sub 予備ペインの生成を予約()
    If Not 保持予備生成表示 Is Nothing Then
    
        Exit Sub
    End If
    
    保持予備生成待機中 = True
    Set 保持予備生成タイマー = New mod_WV2_安全タイマー
    保持予備生成タイマー.指定間隔のタイマーを開始または停止 予備生成待機ミリ秒
    
End Sub

' =========================================================================
' 処理概要       : 予備ペインがない場合に生成を予約
' =========================================================================
Private Sub 予備ペインがない場合に生成を予約()
    If Not 保持予備生成表示 Is Nothing Then Exit Sub
    If 保持予備生成待機中 Then Exit Sub
    
    If 保持作成中表示 > 0 Then
    
        Exit Sub
    End If
    
    ' 次回のタブ作成に備えてペインを確保する
    予備ペインの生成を予約
End Sub

' =========================================================================
' 処理概要       : 予約済みの予備ペイン生成を実行
' 引数           : 通知時刻 (Long) - タイマー通知時刻
'                  任意データ (Variant) - タイマー通知へ引き渡すデータ
' =========================================================================
Private Sub 保持予備生成タイマー_タイマー発生(ByVal 通知時刻 As Long, ByRef 任意データ As Variant)
    If Not 保持予備生成タイマー Is Nothing Then 保持予備生成タイマー.指定間隔のタイマーを開始または停止 0
    
    If Not 保持予備生成待機中 Then Exit Sub
    
    If 保持作成中表示 > 0 Then
        保持予備生成待機中 = False
        
        Exit Sub
    End If
    
    保持予備生成待機中 = False
    新規ウィンドウ要求に備えて予備ペインを生成
End Sub

' =========================================================================
' 処理概要       : 新規ウィンドウ要求に備えて予備ペインを生成
' =========================================================================
Private Sub 新規ウィンドウ要求に備えて予備ペインを生成()
    Dim 表示 As mod_WV2_ブラウザー表示
    Dim 実行結果 As Long
    
    If Not 保持予備生成表示 Is Nothing Then
    
        Exit Sub
    End If
    If 保持状態 <> ブラウザー環境準備済み Then
    
        Exit Sub
    End If
    
    Set 表示 = ペイン作成中の状態を管理してペインを生成()
    If 表示 Is Nothing Then
    
        Exit Sub
    End If
    
    実行結果 = 表示.コントローラーの表示状態を設定(False)
    
    Set 保持予備生成表示 = 表示
    保持予備生成予約済み = False
    
End Sub

' =========================================================================
' 処理概要       : 予備ペインのタブへの取り込みを予約
' 引数           : URL (String) - 移動先または通知対象のURL
' =========================================================================
Public Sub 予備ペインのタブへの取り込みを予約(ByVal URL As String)
    保持取込待機中 = True
    保持取込URL = URL
    
    Set 保持取込タイマー = New mod_WV2_安全タイマー
    保持取込タイマー.指定間隔のタイマーを開始または停止 1
    
End Sub

' =========================================================================
' 処理概要       : 予約済みの予備ペイン取り込みを実行
' 引数           : 通知時刻 (Long) - タイマー通知時刻
'                  任意データ (Variant) - タイマー通知へ引き渡すデータ
' =========================================================================
Private Sub 保持取込タイマー_タイマー発生(ByVal 通知時刻 As Long, ByRef 任意データ As Variant)
    If Not 保持取込タイマー Is Nothing Then 保持取込タイマー.指定間隔のタイマーを開始または停止 0
    
    If Not 保持取込待機中 Then Exit Sub
    
    保持取込待機中 = False
    
    予備ペインを通常のタブとして取り込み
    新規ウィンドウ要求に備えて予備ペインを生成
End Sub

' =========================================================================
' 処理概要       : 予備ペインを通常のタブとして取り込み
' =========================================================================
Private Sub 予備ペインを通常のタブとして取り込み()
    Dim 表示 As mod_WV2_ブラウザー表示
    
    If 保持予備生成表示 Is Nothing Then
    
        保持予備生成予約済み = False
        Exit Sub
    End If
    
    Set 表示 = 保持予備生成表示
    Set 保持予備生成表示 = Nothing
    保持予備生成予約済み = False
    
    ' 作成したペインを管理対象に登録する
    保持タブ一覧.Add 表示
    
    RaiseEvent タブ追加(保持タブ一覧.Count)
    指定タブを表示して選択状態を更新 保持タブ一覧.Count
End Sub

' =========================================================================
' 処理概要       : 予約されたURLを順に新しいタブで表示
' 引数           : 通知時刻 (Long) - タイマー通知時刻
'                  任意データ (Variant) - タイマー通知へ引き渡すデータ
' =========================================================================
Private Sub 保持新規タブタイマー_タイマー発生(ByVal 通知時刻 As Long, ByRef 任意データ As Variant)
    Dim URL As String
    Dim 表示 As mod_WV2_ブラウザー表示
    
    ' 新規タブの生成を待機している場合は、予約を解除する
    If Not 保持新規タブタイマー Is Nothing Then 保持新規タブタイマー.指定間隔のタイマーを開始または停止 0
    
    If 保持待機中URL一覧 Is Nothing Then Exit Sub
    
    Do While 保持待機中URL一覧.Count > 0
        URL = 保持待機中URL一覧(1)
        保持待機中URL一覧.Remove 1
        
        Set 表示 = 新しいタブを作成して指定URLを表示(URL)
    Loop
End Sub

' =========================================================================
' 処理概要       : 終了前に各タブを空白ページへ移動する設定を取得
' 戻り値         : Boolean - 有効時 True / 無効時 False
' =========================================================================
Public Property Get 終了前に空白ページへ移動() As Boolean
    終了前に空白ページへ移動 = 保持静止有効終了
End Property

' =========================================================================
' 処理概要       : 終了前に各タブを空白ページへ移動する設定を更新
' 引数           : 新規値 (Boolean) - 設定する値
' =========================================================================
Public Property Let 終了前に空白ページへ移動(ByVal 新規値 As Boolean)
    保持静止有効終了 = 新規値
    
End Property

' =========================================================================
' 処理概要       : 終了前に各タブを空白ページへ移動して処理を静止
' =========================================================================
Private Sub 終了前に各タブを空白ページへ移動して処理を静止()
    Dim 表示 As mod_WV2_ブラウザー表示
    Dim 件数 As Long
    Dim 待機開始時刻 As Single
    
    待機開始時刻 = Timer
    
    If Not 保持予備生成表示 Is Nothing Then
        保持予備生成表示.指定URLへの移動を開始 "about:blank"
        件数 = 件数 + 1
    End If
    
    If Not 保持タブ一覧 Is Nothing Then
        For Each 表示 In 保持タブ一覧
            If Not 表示 Is Nothing Then
                表示.指定URLへの移動を開始 "about:blank"
                件数 = 件数 + 1
            End If
        Next 表示
    End If
    
End Sub

' =========================================================================
' 処理概要       : タイマーとペインを停止して循環参照を解除
' =========================================================================
Public Sub タイマーとペインを停止して循環参照を解除()
    Dim 表示 As mod_WV2_ブラウザー表示
    
    If 保持静止有効終了 Then 終了前に各タブを空白ページへ移動して処理を静止
    
    ' 新規タブの生成を待機している場合は、予約を解除する
    If Not 保持新規タブタイマー Is Nothing Then
        保持新規タブタイマー.指定間隔のタイマーを開始または停止 0
        Set 保持新規タブタイマー = Nothing
    End If
    
    If Not 保持設定タイマー Is Nothing Then
        保持設定タイマー.指定間隔のタイマーを開始または停止 0
        Set 保持設定タイマー = Nothing
    End If
    保持設定待機中 = False
    Set 保持設定設定連携 = Nothing
    
    If Not 保持予備生成タイマー Is Nothing Then
        保持予備生成タイマー.指定間隔のタイマーを開始または停止 0
        Set 保持予備生成タイマー = Nothing
    End If
    If Not 保持取込タイマー Is Nothing Then
        保持取込タイマー.指定間隔のタイマーを開始または停止 0
        Set 保持取込タイマー = Nothing
    End If
    保持予備生成待機中 = False
    保持取込待機中 = False
    If Not 保持予備生成表示 Is Nothing Then
        保持予備生成表示.所有元ブラウザーへの参照を解除
        Set 保持予備生成表示 = Nothing
        保持予備生成予約済み = False
        
    End If
    
    If 保持タブ一覧 Is Nothing Then Exit Sub
    
    For Each 表示 In 保持タブ一覧
        If Not 表示 Is Nothing Then 表示.所有元ブラウザーへの参照を解除
    Next 表示
    
End Sub

' =========================================================================
' 処理概要       : 保持するリソースと参照を解放
' =========================================================================
Private Sub Class_Terminate()
    Dim 添字 As Long
    
    If Not 保持予備生成表示 Is Nothing Then
        保持予備生成表示.所有元ブラウザーへの参照を解除
        Set 保持予備生成表示 = Nothing
        
    End If
    
    For 添字 = 保持タブ一覧.Count To 1 Step -1
        保持タブ一覧.Remove 添字
    Next 添字
    Set 保持タブ一覧 = Nothing
    
    Set 保持環境 = Nothing
    
    保持ウィンドウハンドルホスト = 0
    保持状態 = ブラウザー作成済み
    
End Sub

' =========================================================================
' 処理概要       : ペインのJavaScript実行結果をブラウザーのイベントとして通知
' 引数           : 表示 (mod_WV2_ブラウザー表示) - 対象のブラウザーペイン
'                  通知識別番号 (Long) - 非同期処理を識別するID
'                  エラーコード (Long) - 非同期処理のHRESULT
'                  結果JSON (String) - JavaScriptの実行結果を表すJSON
' =========================================================================
Public Sub ペインのJavaScript実行結果をブラウザーのイベントとして通知(ByVal 表示 As mod_WV2_ブラウザー表示, ByVal 通知識別番号 As Long, ByVal エラーコード As Long, ByVal 結果JSON As String)
    RaiseEvent スクリプト実行完了(表示, 通知識別番号, エラーコード, 結果JSON)
End Sub
