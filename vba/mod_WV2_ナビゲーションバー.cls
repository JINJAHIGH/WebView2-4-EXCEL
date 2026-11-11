VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "mod_WV2_ナビゲーションバー"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' -------------
' 定数
' -------------
Private Const ナビバー仮想ホスト名 As String = "appassets.navbar"
Private Const ナビバーHTML名       As String = "navbar.html"

' -------------
' 待機時間
' -------------
Private Const 初期化待機上限秒 As Single = 10#

Private WithEvents 保持ブラウザー As mod_WV2_ブラウザー管理
Attribute 保持ブラウザー.VB_VarHelpID = -1
Private 保持バー表示    As mod_WV2_ブラウザー表示
Private 保持ウィンドウハンドルホスト   As LongPtr
Private 保持フォルダーパス As String
Private 保持準備完了      As Boolean

' =========================================================================
' 処理概要       : インスタンスの内部状態を初期化
' =========================================================================
Private Sub Class_Initialize()
    Set 保持ブラウザー = Nothing
    Set 保持バー表示 = Nothing
    保持ウィンドウハンドルホスト = 0
    保持フォルダーパス = ""
    保持準備完了 = False
End Sub

' =========================================================================
' 処理概要       : 共有環境にナビゲーションバー用のペインを生成
' 引数           : ブラウザー (mod_WV2_ブラウザー管理) - 操作対象のブラウザー
'                  ウィンドウハンドル (LongPtr) - WebView2を配置するウィンドウハンドル
'                  フォルダーパス (String) - 対象の保存先またはHTMLフォルダー
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 初期化(ByVal ブラウザー As mod_WV2_ブラウザー管理, ByVal ウィンドウハンドル As LongPtr, ByVal フォルダーパス As String) As Boolean
    Dim 環境 As mod_WV2_ブラウザー環境
    Dim 実行結果 As Long
    Dim ホスト公開結果 As Long
    Dim URL As String
    
    If ブラウザー Is Nothing Then
    
        Exit Function
    End If
    
    Set 環境 = ブラウザー.共有環境
    If 環境 Is Nothing Then
    
        Exit Function
    End If
    
    Set 保持ブラウザー = ブラウザー
    保持ウィンドウハンドルホスト = ウィンドウハンドル
    保持フォルダーパス = フォルダーパス
    
    If Not ナビゲーション画面のHTMLをUTF8で保存() Then
    
        Exit Function
    End If
    
    Set 保持バー表示 = New mod_WV2_ブラウザー表示
    保持バー表示.初期化 環境, 保持ウィンドウハンドルホスト
    
    If Not ペインの初期化完了を待機(保持バー表示) Then
        Exit Function
    End If
    
    実行結果 = 保持バー表示.コントローラーからWebView2インターフェースを取得して保持()
    If 実行結果 <> 0 Or 保持バー表示.表示ポインター = 0 Then
    
        Exit Function
    End If
    
    実行結果 = 保持バー表示.仮想ホスト名にローカルフォルダーを割り当て(ナビバー仮想ホスト名, 保持フォルダーパス, 1)
    If 実行結果 <> 0 Then
    
        Exit Function
    End If
    
    ホスト公開結果 = 保持バー表示.VBAオブジェクトをJavaScriptへ公開("NavBar", Me)
    
    URL = "https://" & ナビバー仮想ホスト名 & "/" & ナビバーHTML名
    実行結果 = 保持バー表示.指定URLへの移動を開始(URL)
    If 実行結果 <> 0 Then
    
        Exit Function
    End If
    
    保持バー表示.Webメッセージの通知先を設定 Me
    実行結果 = 保持バー表示.Webメッセージ受信イベントを登録()
    If 実行結果 <> 0 Then
    
        Exit Function
    End If
    
    保持準備完了 = True
    初期化 = True
    
End Function

' =========================================================================
' 処理概要       : 画面部品の表示領域を更新
' 引数           : バー矩形 (矩形) - 画面部品の表示領域（ピクセル）
' =========================================================================
Public Sub 画面部品の表示領域を更新(ByRef バー矩形 As 矩形)
    Dim 実行結果 As Long
    
    If 保持バー表示 Is Nothing Then Exit Sub
    
    実行結果 = 保持バー表示.コントローラーの表示領域を設定(バー矩形)
    保持バー表示.コントローラーの表示状態を設定 True
End Sub

' =========================================================================
' 処理概要       : タイマーとペインを停止して循環参照を解除
' =========================================================================
Public Sub タイマーとペインを停止して循環参照を解除()
    If Not 保持バー表示 Is Nothing Then
        保持バー表示.Webメッセージの通知先を解除
        保持バー表示.イベントとコントローラーを閉じて関連参照を解放
        Set 保持バー表示 = Nothing
    End If
    Set 保持ブラウザー = Nothing
    保持準備完了 = False
    
End Sub

' =========================================================================
' 処理概要       : 初期化が完了して利用できるか判定
' 戻り値         : Boolean - 初期化完了時 True
' =========================================================================
Public Property Get 利用可能() As Boolean
    利用可能 = 保持準備完了
End Property

' =========================================================================
' 処理概要       : タブのURL変更をナビゲーション画面に反映
' 引数           : 位置番号 (Long) - 対象の位置番号（公開タブ・ログは1始まり）
'                  URL (String) - 移動先または通知対象のURL
'                  可能進む戻る (Boolean) - 前の履歴へ戻れる状態
'                  可能進む進む (Boolean) - 次の履歴へ進める状態
' =========================================================================
Private Sub 保持ブラウザー_タブURL変更(ByVal 位置番号 As Long, ByVal URL As String, ByVal 可能進む戻る As Boolean, ByVal 可能進む進む As Boolean)
    If 保持ブラウザー Is Nothing Then Exit Sub
    
    If 位置番号 <> 保持ブラウザー.選択タブ番号 Then
    
        Exit Sub
    End If
    
    URLと履歴状態をJSONとしてナビゲーション画面へ送信 URL, 可能進む戻る, 可能進む進む
End Sub

' =========================================================================
' 処理概要       : 選択タブの変更を画面部品に反映
' 引数           : 位置番号 (Long) - 対象の位置番号（公開タブ・ログは1始まり）
' =========================================================================
Private Sub 保持ブラウザー_選択タブ変更(ByVal 位置番号 As Long)
    表示中のタブのURLと履歴状態を画面に同期
End Sub

' =========================================================================
' 処理概要       : タブの削除を画面部品に反映
' 引数           : 位置番号 (Long) - 対象の位置番号（公開タブ・ログは1始まり）
' =========================================================================
Private Sub 保持ブラウザー_タブ終了(ByVal 位置番号 As Long)
    表示中のタブのURLと履歴状態を画面に同期
End Sub

' =========================================================================
' 処理概要       : 画面から受信したコマンドを振り分けて実行
' 引数           : 表示 (mod_WV2_ブラウザー表示) - 対象のブラウザーペイン
'                  元データ (String) - 解析するJSON文字列
'                  メッセージ (String) - 画面から受信したメッセージ
' =========================================================================
Public Sub 画面から受信したコマンドを振り分けて実行(ByVal 表示 As mod_WV2_ブラウザー表示, ByVal 元データ As String, ByVal メッセージ As String)
    Dim 受信コマンド As String
    Dim URL As String
    
    If 保持ブラウザー Is Nothing Then
    
        Exit Sub
    End If
    
    受信コマンド = JSONから指定キーの文字列を取得(メッセージ, "cmd")
    
    Select Case 受信コマンド
    
        Case "ready"
        
            表示中のタブのURLと履歴状態を画面に同期
            
        Case "back"
        
            画面からの要求で前の履歴へ戻る
            
        Case "forward"
        
            画面からの要求で次の履歴へ進む
            
        Case "reload"
        
            画面からの要求でページを再読み込み
            
        Case "navigate"
        
            URL = JSONから指定キーの文字列を取得(メッセージ, "url")
            
            画面からの要求で指定URLまたは検索結果へ移動 URL
            
        Case ""
        
        Case Else
        
    End Select
End Sub

' =========================================================================
' 処理概要       : 画面からの要求で前の履歴へ戻る
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 画面からの要求で前の履歴へ戻る() As Boolean
    Dim 成功 As Boolean
    
    If 保持ブラウザー Is Nothing Then
    
        画面からの要求で前の履歴へ戻る = False
        Exit Function
    End If
    成功 = 保持ブラウザー.選択中のタブで前の履歴へ戻る()
    画面からの要求で前の履歴へ戻る = 成功
End Function

' =========================================================================
' 処理概要       : 画面からの要求で次の履歴へ進む
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 画面からの要求で次の履歴へ進む() As Boolean
    Dim 成功 As Boolean
    
    If 保持ブラウザー Is Nothing Then
    
        画面からの要求で次の履歴へ進む = False
        Exit Function
    End If
    成功 = 保持ブラウザー.選択中のタブで次の履歴へ進む()
    画面からの要求で次の履歴へ進む = 成功
End Function

' =========================================================================
' 処理概要       : 画面からの要求でページを再読み込み
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 画面からの要求でページを再読み込み() As Boolean
    Dim 成功 As Boolean
    
    If 保持ブラウザー Is Nothing Then
    
        画面からの要求でページを再読み込み = False
        Exit Function
    End If
    成功 = 保持ブラウザー.選択中のタブを再読み込み()
    画面からの要求でページを再読み込み = 成功
End Function

' =========================================================================
' 処理概要       : 画面からの要求で指定URLまたは検索結果へ移動
' 引数           : URL (String) - 移動先または通知対象のURL
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 画面からの要求で指定URLまたは検索結果へ移動(ByVal URL As String) As Boolean
    Dim 成功 As Boolean
    
    If 保持ブラウザー Is Nothing Then
    
        画面からの要求で指定URLまたは検索結果へ移動 = False
        Exit Function
    End If
    If Len(Trim$(URL)) = 0 Then
    
        画面からの要求で指定URLまたは検索結果へ移動 = False
        Exit Function
    End If
    成功 = 保持ブラウザー.入力をURLに解決して選択中のタブへ移動(URL)
    画面からの要求で指定URLまたは検索結果へ移動 = 成功
End Function

' =========================================================================
' 処理概要       : 表示中のタブのURLと履歴状態を画面に同期
' =========================================================================
Private Sub 表示中のタブのURLと履歴状態を画面に同期()
    Dim URL As String
    Dim 可能戻る As Boolean
    Dim 可能進む As Boolean
    
    If 保持ブラウザー Is Nothing Then Exit Sub
    
    If 保持ブラウザー.タブ数 = 0 Or 保持ブラウザー.選択タブ番号 < 1 Then
    
        URL = ""
        可能戻る = False
        可能進む = False
    Else
        URL = 保持ブラウザー.選択中のタブのURLを取得()
        可能戻る = 保持ブラウザー.選択中のタブで前の履歴へ戻れるか判定()
        可能進む = 保持ブラウザー.選択中のタブで次の履歴へ進めるか判定()
    End If
    
    URLと履歴状態をJSONとしてナビゲーション画面へ送信 URL, 可能戻る, 可能進む
End Sub

' =========================================================================
' 処理概要       : URLと履歴状態をJSONとしてナビゲーション画面へ送信
' 引数           : URL (String) - 移動先または通知対象のURL
'                  可能戻る (Boolean) - 前の履歴へ戻れる状態
'                  可能進む (Boolean) - 次の履歴へ進める状態
' =========================================================================
Private Sub URLと履歴状態をJSONとしてナビゲーション画面へ送信(ByVal URL As String, ByVal 可能戻る As Boolean, ByVal 可能進む As Boolean)
    Dim 保持済みタブ As Boolean
    Dim 表示済み As String
    Dim JSON As String
    Dim 実行結果 As Long
    
    If 保持バー表示 Is Nothing Then Exit Sub
    If Not 保持準備完了 Then Exit Sub
    
    If Not 保持ブラウザー Is Nothing Then
        保持済みタブ = (保持ブラウザー.タブ数 > 0 And 保持ブラウザー.選択タブ番号 >= 1)
    End If
    
    If URL = "about:blank" Then
        表示済み = ""
    Else
        表示済み = URL
    End If
    
    JSON = "{""type"":""navsync"",""url"":""" & 文字列をJSONに埋め込める形式へ変換(表示済み) & """" & ",""canBack"":" & Boolean値をJSONの真偽値文字列へ変換(可能戻る) & ",""canForward"":" & Boolean値をJSONの真偽値文字列へ変換(可能進む) & ",""hasTab"":" & _
        Boolean値をJSONの真偽値文字列へ変換(保持済みタブ) & "}"
        
    実行結果 = 保持バー表示.ページのJavaScriptへ文字列メッセージを送信(JSON)
End Sub

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
' 処理概要       : ナビゲーション画面のHTMLをUTF-8で保存
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Private Function ナビゲーション画面のHTMLをUTF8で保存() As Boolean
    Dim 完全パス As String
    Dim HTML As String
    Dim バイナリストリーム As Object
    Dim バイト数() As Byte
    Dim 出力ストリーム As Object
    
    On Error GoTo ErrorHanler
    
    If Dir(保持フォルダーパス, vbDirectory) = "" Then
        MkDir 保持フォルダーパス
        
    End If
    
    完全パス = 保持フォルダーパス
    If Right$(完全パス, 1) <> "\" Then 完全パス = 完全パス & "\"
    完全パス = 完全パス & ナビバーHTML名
    
    HTML = ナビゲーション画面のHTMLを生成()
    
    Set バイナリストリーム = CreateObject("ADODB.Stream")
    バイナリストリーム.Type = 2
    バイナリストリーム.Charset = "UTF-8"
    バイナリストリーム.Open
    バイナリストリーム.WriteText HTML
    
    バイナリストリーム.Position = 0
    バイナリストリーム.Type = 1
    バイナリストリーム.Position = 3
    バイト数 = バイナリストリーム.Read
    バイナリストリーム.Close
    
    Set 出力ストリーム = CreateObject("ADODB.Stream")
    出力ストリーム.Type = 1
    出力ストリーム.Open
    出力ストリーム.Write バイト数
    出力ストリーム.SaveToFile 完全パス, 2
    出力ストリーム.Close
    
    ナビゲーション画面のHTMLをUTF8で保存 = True
    Exit Function
    
ErrorHanler:
    MsgBox "処理を完了できません。" & vbCrLf & Err.Description, vbExclamation, "ナビゲーションバー"
    
    ナビゲーション画面のHTMLをUTF8で保存 = False
    GoTo EndProc
    
EndProc:
End Function

' =========================================================================
' 処理概要       : ナビゲーション画面のHTMLを生成
' 戻り値         : String - ナビゲーション画面のHTML
' =========================================================================
Private Function ナビゲーション画面のHTMLを生成() As String
    Dim 作成文字列 As String
    
    作成文字列 = "<!DOCTYPE html>" & vbLf
    作成文字列 = 作成文字列 & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    作成文字列 = 作成文字列 & "<style>" & vbLf
    作成文字列 = 作成文字列 & "  html,body{margin:0;padding:0;height:100%;box-sizing:border-box;" & "font-family:'Segoe UI','Meiryo',sans-serif;}" & vbLf
    作成文字列 = 作成文字列 & "  *{box-sizing:border-box;}" & vbLf
    作成文字列 = 作成文字列 & "  body{display:flex;align-items:center;gap:4px;padding:0 6px;" & "overflow:hidden;background:#fafafa;border-bottom:1px solid #ccc;}" & vbLf
    作成文字列 = 作成文字列 & "  button{width:30px;height:26px;border:1px solid transparent;" & "border-radius:4px;background:transparent;cursor:pointer;" & "font-size:15px;color:#333;line-height:1;padding:0;flex:none;}" & vbLf
    作成文字列 = 作成文字列 & "  button:hover:not(:disabled){background:#e6e6e6;}" & vbLf
    作成文字列 = 作成文字列 & "  button:disabled{color:#bbb;cursor:default;}" & vbLf
    作成文字列 = 作成文字列 & "  #url{flex:1;height:26px;padding:0 10px;border:1px solid #ccc;" & "border-radius:13px;background:#fff;font-size:13px;color:#222;" & "outline:none;min-width:0;}" & vbLf
    作成文字列 = 作成文字列 & "  #url:focus{border-color:#0078d4;box-shadow:0 0 0 1px #0078d4;}" & vbLf
    作成文字列 = 作成文字列 & "  #url:disabled{background:#f0f0f0;color:#999;}" & vbLf
    作成文字列 = 作成文字列 & "</style></head><body>" & vbLf
    作成文字列 = 作成文字列 & "  <button id=""back""    title=""戻る""       disabled>&#8592;</button>" & vbLf
    作成文字列 = 作成文字列 & "  <button id=""forward"" title=""進む""       disabled>&#8594;</button>" & vbLf
    作成文字列 = 作成文字列 & "  <button id=""reload""  title=""再読み込み"" disabled>&#10227;</button>" & vbLf
    作成文字列 = 作成文字列 & "  <input  id=""url"" type=""text"" placeholder=""URL を入力"" spellcheck=""false"" disabled>" & vbLf
    作成文字列 = 作成文字列 & "<script>" & vbLf
    作成文字列 = 作成文字列 & "  var W = window.chrome && window.chrome.webview;" & vbLf
    作成文字列 = 作成文字列 & "  function postCmd(obj){ if(W){ W.postMessage(JSON.stringify(obj)); } }" & vbLf
    
    作成文字列 = 作成文字列 & "  function HO(){ return (W && W.hostObjects && W.hostObjects.sync) ? W.hostObjects.sync : null; }" & vbLf
    作成文字列 = 作成文字列 & "  function navHost(msgObj, callHost){" & vbLf
    作成文字列 = 作成文字列 & "    var ho = HO();" & vbLf
    作成文字列 = 作成文字列 & "    if(ho && ho.NavBar){" & vbLf
    作成文字列 = 作成文字列 & "      try{ callHost(ho.NavBar); return; }" & vbLf
    作成文字列 = 作成文字列 & "      catch(e){ postCmd(msgObj); return; }" & vbLf
    作成文字列 = 作成文字列 & "    }" & vbLf
    作成文字列 = 作成文字列 & "    postCmd(msgObj);" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    
    作成文字列 = 作成文字列 & "  var editing = false;" & vbLf
    
    作成文字列 = 作成文字列 & "  var lastUrl = '';" & vbLf
    
    作成文字列 = 作成文字列 & "  document.getElementById('back').onclick = function(){ navHost({cmd:'back'}, function(nb){ nb.画面からの要求で前の履歴へ戻る(); }); };" & vbLf
    作成文字列 = 作成文字列 & "  document.getElementById('forward').onclick = function(){ navHost({cmd:'forward'}, function(nb){ nb.画面からの要求で次の履歴へ進む(); }); };" & vbLf
    作成文字列 = 作成文字列 & "  document.getElementById('reload').onclick = function(){ navHost({cmd:'reload'}, function(nb){ nb.画面からの要求でページを再読み込み(); }); };" & vbLf
    
    作成文字列 = 作成文字列 & "  document.getElementById('url').onkeydown = function(ev){" & vbLf
    作成文字列 = 作成文字列 & "    if(ev.key === 'Enter'){" & vbLf
    作成文字列 = 作成文字列 & "      var v = this.value;" & vbLf
    作成文字列 = 作成文字列 & "      editing = false;" & vbLf
    作成文字列 = 作成文字列 & "      if(v && v.length){ navHost({cmd:'navigate', url:v}, function(nb){ nb.画面からの要求で指定URLまたは検索結果へ移動(v); }); }" & vbLf
    作成文字列 = 作成文字列 & "    } else if(ev.key === 'Escape'){" & vbLf
    作成文字列 = 作成文字列 & "      editing = false;" & vbLf
    作成文字列 = 作成文字列 & "      this.value = lastUrl;   " & vbLf
    作成文字列 = 作成文字列 & "      this.blur();" & vbLf
    作成文字列 = 作成文字列 & "    }" & vbLf
    作成文字列 = 作成文字列 & "  };" & vbLf
    
    作成文字列 = 作成文字列 & "  function applyNavSync(st){" & vbLf
    作成文字列 = 作成文字列 & "    var u = document.getElementById('url');" & vbLf
    
    作成文字列 = 作成文字列 & "    lastUrl = st.url || '';" & vbLf
    作成文字列 = 作成文字列 & "    var b = document.getElementById('back');" & vbLf
    作成文字列 = 作成文字列 & "    var f = document.getElementById('forward');" & vbLf
    作成文字列 = 作成文字列 & "    var r = document.getElementById('reload');" & vbLf
    
    作成文字列 = 作成文字列 & "    b.disabled = !st.canBack;" & vbLf
    作成文字列 = 作成文字列 & "    f.disabled = !st.canForward;" & vbLf
    作成文字列 = 作成文字列 & "    r.disabled = !st.hasTab;" & vbLf
    
    作成文字列 = 作成文字列 & "    if(!editing){" & vbLf
    作成文字列 = 作成文字列 & "      var want = !st.hasTab;" & vbLf
    作成文字列 = 作成文字列 & "      if(u.disabled !== want){ u.disabled = want; }" & vbLf
    作成文字列 = 作成文字列 & "      u.value = st.url || '';" & vbLf
    作成文字列 = 作成文字列 & "    }" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    作成文字列 = 作成文字列 & "  (function(){" & vbLf
    作成文字列 = 作成文字列 & "    var u = document.getElementById('url');" & vbLf
    作成文字列 = 作成文字列 & "    u.addEventListener('input', function(){ editing = true; });" & vbLf
    作成文字列 = 作成文字列 & "    u.addEventListener('blur',  function(){ editing = false; });" & vbLf
    作成文字列 = 作成文字列 & "  })();" & vbLf
    作成文字列 = 作成文字列 & "  if(W){" & vbLf
    作成文字列 = 作成文字列 & "    W.addEventListener('message', function(ev){" & vbLf
    作成文字列 = 作成文字列 & "      var data = ev.data;" & vbLf
    作成文字列 = 作成文字列 & "      if(typeof data === 'string'){ try{ data = JSON.parse(data); }catch(e){ return; } }" & vbLf
    作成文字列 = 作成文字列 & "      if(data && data.type === 'navsync'){ applyNavSync(data); }" & vbLf
    作成文字列 = 作成文字列 & "    });" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    
    作成文字列 = 作成文字列 & "  postCmd({cmd:'ready'});" & vbLf
    作成文字列 = 作成文字列 & "</script>" & vbLf
    作成文字列 = 作成文字列 & "</body></html>"
    ナビゲーション画面のHTMLを生成 = 作成文字列
End Function
