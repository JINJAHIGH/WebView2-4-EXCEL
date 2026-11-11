VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_WV2_ブラウザーフォーム
   Caption         =   "WebView2"
   ClientHeight    =   8850.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12375
   OleObjectBlob   =   "frm_WV2_ブラウザーフォーム.frx":0000
   StartUpPosition =   1
End
Attribute VB_Name = "frm_WV2_ブラウザーフォーム"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' -------------
' 定数
' -------------
Private Const タブバー高さ As Long = 36
Private Const ナビバー高さ As Long = 34

Public Event スクリプト実行完了(ByVal 通知識別番号 As Long, ByVal エラーコード As Long, ByVal 結果JSON As String)

Private ハンドルウィンドウ表示枠 As LongPtr

Private ハンドルウィンドウフォーム    As LongPtr

Private 保持余白左    As Single
Private 保持余白右   As Single
Private 保持余白上     As Single
Private 保持余白下  As Single

Private 保持サイズ変更準備完了 As Boolean

Private WithEvents 保持ブラウザー As mod_WV2_ブラウザー管理
Attribute 保持ブラウザー.VB_VarHelpID = -1
Private 保持タブバー As mod_WV2_タブバー
Private 保持遷移バー As mod_WV2_ナビゲーションバー

' =========================================================================
' 処理概要       : 表示領域のハンドルとフォームの余白を取得
' =========================================================================
Private Sub UserForm_Initialize()
    Dim スタイル As LongPtr
    
    ' WebView2の配置先となるフレームのウィンドウハンドルを取得する
    ハンドルウィンドウ表示枠 = 表示枠.[_GethWnd]
    
    ハンドルウィンドウフォーム = 親ウィンドウ取得API(ハンドルウィンドウ表示枠, 最上位親を取得)
    
    スタイル = ウィンドウ属性取得API(ハンドルウィンドウフォーム, ウィンドウスタイル位置)
    ウィンドウ属性設定API ハンドルウィンドウフォーム, ウィンドウスタイル位置, スタイル Or サイズ変更枠
    ウィンドウ位置設定API ハンドルウィンドウフォーム, 0, 0, 0, 0, 0, 位置変更なし Or サイズ変更なし Or 重なり順変更なし Or フレーム更新
    
    ' サイズ変更時にも維持する四辺の余白を記録する
    保持余白左 = 表示枠.Left
    保持余白上 = 表示枠.Top
    保持余白右 = Me.InsideWidth - (表示枠.Left + 表示枠.width)
    保持余白下 = Me.InsideHeight - (表示枠.Top + 表示枠.Height)
    
    保持サイズ変更準備完了 = True
End Sub

' =========================================================================
' 処理概要       : フォームの大きさに合わせて各表示領域を配置
' =========================================================================
Private Sub UserForm_Resize()
    Dim 表示矩形 As 矩形
    Dim 縦位置 As Long
    Dim バー矩形 As 矩形
    Dim 遷移矩形 As 矩形
    Dim コンテンツ矩形 As 矩形
    
    If Not 保持サイズ変更準備完了 Then Exit Sub
    
    表示枠.width = Me.InsideWidth - 保持余白左 - 保持余白右
    表示枠.Height = Me.InsideHeight - 保持余白上 - 保持余白下
    
    If 保持ブラウザー Is Nothing Then Exit Sub
    
    クライアント矩形取得API ハンドルウィンドウ表示枠, 表示矩形
    
    ' -------------
    ' 表示領域を上からタブバー、mod_WV2_ナビゲーションバー、コンテンツへ割り当て
    ' -------------
    縦位置 = 表示矩形.上端
    
    If Not 保持タブバー Is Nothing Then
        バー矩形.左端 = 表示矩形.左端
        バー矩形.上端 = 縦位置
        バー矩形.右端 = 表示矩形.右端
        バー矩形.下端 = 縦位置 + タブバー高さ
        If バー矩形.下端 > 表示矩形.下端 Then バー矩形.下端 = 表示矩形.下端
        保持タブバー.画面部品の表示領域を更新 バー矩形
        縦位置 = バー矩形.下端
    End If
    
    If Not 保持遷移バー Is Nothing Then
        遷移矩形.左端 = 表示矩形.左端
        遷移矩形.上端 = 縦位置
        遷移矩形.右端 = 表示矩形.右端
        遷移矩形.下端 = 縦位置 + ナビバー高さ
        If 遷移矩形.下端 > 表示矩形.下端 Then 遷移矩形.下端 = 表示矩形.下端
        保持遷移バー.画面部品の表示領域を更新 遷移矩形
        縦位置 = 遷移矩形.下端
    End If
    
    コンテンツ矩形.左端 = 表示矩形.左端
    コンテンツ矩形.上端 = 縦位置
    コンテンツ矩形.右端 = 表示矩形.右端
    コンテンツ矩形.下端 = 表示矩形.下端
    If コンテンツ矩形.上端 > コンテンツ矩形.下端 Then コンテンツ矩形.上端 = コンテンツ矩形.下端
    保持ブラウザー.全タブの表示領域を更新 コンテンツ矩形
End Sub

' =========================================================================
' 処理概要       : 表示中のページへ文字列メッセージを送信
' 引数           : 文字列 (String) - ページへ送信する文字列
' =========================================================================
Public Sub 表示中のページへ文字列メッセージを送信(ByVal 文字列 As String)
    Dim 表示 As mod_WV2_ブラウザー表示
    Dim 実行結果 As Long
    
    Set 表示 = 表示中のペインを取得
    If 表示 Is Nothing Then
    
        Exit Sub
    End If
    If Not 表示.利用可能 Then
    
        Exit Sub
    End If
    
    実行結果 = 表示.ページのJavaScriptへ文字列メッセージを送信(文字列)
End Sub

' =========================================================================
' 処理概要       : ページの読み込み完了状態を取得
' 戻り値         : Boolean - 読み込み完了時 True
' =========================================================================
Public Property Get 読込完了() As Boolean
    Dim 表示 As mod_WV2_ブラウザー表示
    
    Set 表示 = 表示中のペインを取得
    If 表示 Is Nothing Then
        読込完了 = False
    Else
        読込完了 = 表示.読込完了
    End If
End Property

' =========================================================================
' 処理概要       : 表示中のブラウザーペインを取得
' 戻り値         : mod_WV2_ブラウザー表示 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Public Property Get 表示中のペインを取得() As mod_WV2_ブラウザー表示
    If 保持ブラウザー Is Nothing Then
    
        Set 表示中のペインを取得 = Nothing
        Exit Property
    End If
    Set 表示中のペインを取得 = 保持ブラウザー.選択中の表示
End Property

' =========================================================================
' 処理概要       : JavaScriptの実行結果を利用側へ通知
' 引数           : 通知識別番号 (Long) - 非同期処理を識別するID
'                  エラーコード (Long) - 非同期処理のHRESULT
'                  結果JSON (String) - JavaScriptの実行結果を表すJSON
' =========================================================================
Public Sub JavaScriptの実行結果を利用側へ通知(ByVal 通知識別番号 As Long, ByVal エラーコード As Long, ByVal 結果JSON As String)
    RaiseEvent スクリプト実行完了(通知識別番号, エラーコード, 結果JSON)
End Sub

' =========================================================================
' 処理概要       : WebView2を配置するフレームのウィンドウハンドルを取得
' 戻り値         : LongPtr - 対象のアドレス。取得できない場合 0
' =========================================================================
Public Property Get 表示枠ハンドル() As LongPtr
    表示枠ハンドル = ハンドルウィンドウ表示枠
End Property

' =========================================================================
' 処理概要       : 指定フォルダーのHTMLを仮想ホスト経由で表示
' 引数           : フォルダーパス (String) - 対象の保存先またはHTMLフォルダー
'                  ホスト名称 (String) - ローカルフォルダーへ割り当てる仮想ホスト名（省略時: "appassets.local"）
'                  ファイル名称 (String) - 表示または保存するHTMLファイル名（省略時: "位置番号.HTML"）
'                  データフォルダー (String) - プロファイルと設定の保存先フォルダー（省略時: ""）
' =========================================================================
Public Sub ローカルHTMLを表示(ByVal フォルダーパス As String, Optional ByVal ホスト名称 As String = "appassets.local", Optional ByVal ファイル名称 As String = "index.html", Optional ByVal データフォルダー As String = "")
    Dim 表示 As mod_WV2_ブラウザー表示, エラー文字列 As String
    
    On Error GoTo ErrorHanler
    If Not 保持ブラウザー Is Nothing Then MsgBox "ブラウザーは既に起動しています。", vbInformation, "ブラウザー": GoTo EndProc
    If Len(データフォルダー) > 0 Then mod_WV2_共通設定.保存先フォルダー = データフォルダー
    If Len(フォルダーパス) = 0 Then MsgBox "表示するHTMLのフォルダーを指定してください。", vbExclamation, "ブラウザー": GoTo EndProc
    
    ' フォームに配置するブラウザーを初期化する
    Set 保持ブラウザー = New mod_WV2_ブラウザー管理
    If Not 保持ブラウザー.初期化(ハンドルウィンドウ表示枠, mod_WV2_共通設定.プロファイル保存先) Then エラー文字列 = "ブラウザーの初期化に失敗しました。": GoTo ErrorHanler
    Set 表示 = 保持ブラウザー.仮想ホストを設定してローカルHTMLのタブを作成(ホスト名称, フォルダーパス, ファイル名称)
    If 表示 Is Nothing Then エラー文字列 = "ローカルHTMLを表示できません。": GoTo ErrorHanler
    UserForm_Resize
    GoTo EndProc
    
ErrorHanler:
    If Len(エラー文字列) = 0 Then エラー文字列 = Err.Description
    MsgBox エラー文字列, vbExclamation, "ブラウザー"
    On Error Resume Next
    表示部品とブラウザーを依存関係の順に終了
    GoTo EndProc
    
EndProc:
    Set 表示 = Nothing
End Sub

' =========================================================================
' 処理概要       : フォーム終了時にブラウザーを解放
' =========================================================================
Private Sub UserForm_Terminate()
    表示部品とブラウザーを依存関係の順に終了
End Sub

' =========================================================================
' 処理概要       : 指定URLをタブバーとナビゲーションバー付きで表示
' 引数           : 初期URL (String) - 最初に表示するURL（省略時: ""）
'                  データフォルダー (String) - プロファイルと設定の保存先フォルダー（省略時: ""）
' =========================================================================
Public Sub ブラウザー画面を開始(Optional ByVal 初期URL As String = "", Optional ByVal データフォルダー As String = "")
    Dim 表示 As mod_WV2_ブラウザー表示, エラー文字列 As String
    
    On Error GoTo ErrorHanler
    If Not 保持ブラウザー Is Nothing Then MsgBox "ブラウザーは既に起動しています。", vbInformation, "ブラウザー": GoTo EndProc
    If Len(データフォルダー) > 0 Then mod_WV2_共通設定.保存先フォルダー = データフォルダー
    If Len(初期URL) = 0 Then 初期URL = mod_WV2_共通設定.初期URL
    
    ' フォームに配置するブラウザーを初期化する
    Set 保持ブラウザー = New mod_WV2_ブラウザー管理
    If Not 保持ブラウザー.初期化(ハンドルウィンドウ表示枠, mod_WV2_共通設定.プロファイル保存先) Then エラー文字列 = "ブラウザーの初期化に失敗しました。": GoTo ErrorHanler
    Set 表示 = 保持ブラウザー.新しいタブを作成して指定URLを表示(初期URL)
    If 表示 Is Nothing Then エラー文字列 = "初期タブを作成できません。": GoTo ErrorHanler
    Set 保持タブバー = New mod_WV2_タブバー
    If Not 保持タブバー.初期化(保持ブラウザー, ハンドルウィンドウ表示枠, mod_WV2_共通設定.画面部品の保存先("tabbar")) Then エラー文字列 = "タブバーを初期化できません。": GoTo ErrorHanler
    Set 保持遷移バー = New mod_WV2_ナビゲーションバー
    If Not 保持遷移バー.初期化(保持ブラウザー, ハンドルウィンドウ表示枠, mod_WV2_共通設定.画面部品の保存先("navbar")) Then エラー文字列 = "ナビゲーションバーを初期化できません。": GoTo ErrorHanler
    UserForm_Resize
    GoTo EndProc
    
ErrorHanler:
    If Len(エラー文字列) = 0 Then エラー文字列 = Err.Description
    MsgBox エラー文字列, vbExclamation, "ブラウザー"
    On Error Resume Next
    表示部品とブラウザーを依存関係の順に終了
    GoTo EndProc
    
EndProc:
    Set 表示 = Nothing
End Sub

' =========================================================================
' 処理概要       : フォームが保持するブラウザーを取得
' 戻り値         : mod_WV2_ブラウザー管理 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Public Property Get 起動中のブラウザー() As mod_WV2_ブラウザー管理
    Set 起動中のブラウザー = 保持ブラウザー
End Property

' =========================================================================
' 処理概要       : 検索エンジンのプリセットを適用
' 引数           : 検索エンジン名称 (String) - 検索エンジンのプリセット名
' =========================================================================
Public Sub 検索エンジンのプリセットを適用(ByVal 検索エンジン名称 As String)
    Dim 設定連携 As mod_WV2_検索設定連携
    Dim 適用結果 As String
    
    If 保持ブラウザー Is Nothing Then
    
        Exit Sub
    End If
    
    Set 設定連携 = New mod_WV2_検索設定連携
    設定連携.設定の読み書き対象となるブラウザーを設定 保持ブラウザー
    
    適用結果 = 設定連携.指定プリセットの検索エンジンを適用して保存(検索エンジン名称)
    
End Sub

' =========================================================================
' 処理概要       : 表示部品とブラウザーを依存関係の順に終了
' =========================================================================
Private Sub 表示部品とブラウザーを依存関係の順に終了()
    ' -------------
    ' 画面部品の参照を解除してからブラウザーを終了
    ' -------------
    If Not 保持遷移バー Is Nothing Then
        保持遷移バー.タイマーとペインを停止して循環参照を解除
        Set 保持遷移バー = Nothing
    End If
    If Not 保持タブバー Is Nothing Then
        保持タブバー.タイマーとペインを停止して循環参照を解除
        Set 保持タブバー = Nothing
    End If
    If Not 保持ブラウザー Is Nothing Then 保持ブラウザー.タイマーとペインを停止して循環参照を解除
    Set 保持ブラウザー = Nothing
End Sub

' =========================================================================
' 処理概要       : JavaScriptの実行結果をフォームの利用側へ通知
' 引数           : 表示 (mod_WV2_ブラウザー表示) - 対象のブラウザーペイン
'                  通知識別番号 (Long) - 非同期処理を識別するID
'                  エラーコード (Long) - 非同期処理のHRESULT
'                  結果JSON (String) - JavaScriptの実行結果を表すJSON
' =========================================================================
Private Sub 保持ブラウザー_スクリプト実行完了(ByVal 表示 As mod_WV2_ブラウザー表示, ByVal 通知識別番号 As Long, ByVal エラーコード As Long, ByVal 結果JSON As String)
    JavaScriptの実行結果を利用側へ通知 通知識別番号, エラーコード, 結果JSON
End Sub
