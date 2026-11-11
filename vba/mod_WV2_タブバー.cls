VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "mod_WV2_タブバー"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' -------------
' 定数
' -------------
Private Const タブバー仮想ホスト名 As String = "appassets.tabbar"
Private Const タブバーHTML名       As String = "tabbar.html"

' -------------
' 待機時間
' -------------
Private Const 初期化待機上限秒 As Single = 10#

Private WithEvents 保持ブラウザー As mod_WV2_ブラウザー管理
Attribute 保持ブラウザー.VB_VarHelpID = -1
Private 保持バー表示   As mod_WV2_ブラウザー表示
Private 保持ウィンドウハンドルホスト  As LongPtr
Private 保持フォルダーパス As String
Private 保持準備完了     As Boolean

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
' 処理概要       : 共有環境にタブバー用のペインを生成
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
    
    If Not タブバーのHTMLをUTF8で保存() Then
    
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
    
    実行結果 = 保持バー表示.仮想ホスト名にローカルフォルダーを割り当て(タブバー仮想ホスト名, 保持フォルダーパス, 1)
    If 実行結果 <> 0 Then
    
        Exit Function
    End If
    
    ホスト公開結果 = 保持バー表示.VBAオブジェクトをJavaScriptへ公開("TabBar", Me)
    
    URL = "https://" & タブバー仮想ホスト名 & "/" & タブバーHTML名
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
' 処理概要       : タブの追加をタブバーに反映
' 引数           : 位置番号 (Long) - 対象の位置番号（公開タブ・ログは1始まり）
' =========================================================================
Private Sub 保持ブラウザー_タブ追加(ByVal 位置番号 As Long)
    タブ一覧と選択状態をタブバーへ同期
End Sub

' =========================================================================
' 処理概要       : タブの削除を画面部品に反映
' 引数           : 位置番号 (Long) - 対象の位置番号（公開タブ・ログは1始まり）
' =========================================================================
Private Sub 保持ブラウザー_タブ終了(ByVal 位置番号 As Long)
    タブ一覧と選択状態をタブバーへ同期
End Sub

' =========================================================================
' 処理概要       : 選択タブの変更を画面部品に反映
' 引数           : 位置番号 (Long) - 対象の位置番号（公開タブ・ログは1始まり）
' =========================================================================
Private Sub 保持ブラウザー_選択タブ変更(ByVal 位置番号 As Long)
    タブ一覧と選択状態をタブバーへ同期
End Sub

' =========================================================================
' 処理概要       : タブのタイトル変更をタブバーに反映
' 引数           : 位置番号 (Long) - 対象の位置番号（公開タブ・ログは1始まり）
'                  タイトル (String) - 通知するページタイトル
' =========================================================================
Private Sub 保持ブラウザー_タブタイトル変更(ByVal 位置番号 As Long, ByVal タイトル As String)
    タブ一覧と選択状態をタブバーへ同期
End Sub

' =========================================================================
' 処理概要       : 画面から受信したコマンドを振り分けて実行
' 引数           : 表示 (mod_WV2_ブラウザー表示) - 対象のブラウザーペイン
'                  元データ (String) - 解析するJSON文字列
'                  メッセージ (String) - 画面から受信したメッセージ
' =========================================================================
Public Sub 画面から受信したコマンドを振り分けて実行(ByVal 表示 As mod_WV2_ブラウザー表示, ByVal 元データ As String, ByVal メッセージ As String)
    Dim 受信コマンド As String
    Dim 対象添字 As Long
    Dim 移動元添字 As Long, 移動先添字 As Long
    
    If 保持ブラウザー Is Nothing Then
    
        Exit Sub
    End If
    
    受信コマンド = JSONから指定キーの文字列を取得(メッセージ, "cmd")
    
    Select Case 受信コマンド
    
        Case "ready"
        
            タブ一覧と選択状態をタブバーへ同期
            
        Case "activate"
        
            対象添字 = JSONから指定キーの数値を取得(メッセージ, "index")
            
            画面からの要求で指定タブを選択 対象添字
            
        Case "close"
        
            対象添字 = JSONから指定キーの数値を取得(メッセージ, "index")
            
            画面からの要求で指定タブを閉じる 対象添字
            
        Case "newtab"
        
            保持ブラウザー.新規タブの作成をタイマー経由で予約 mod_WV2_共通設定.初期URL
            
        Case "settings"
        
            保持ブラウザー.設定タブの表示をタイマー経由で予約
            
        Case "reorder"
        
            移動元添字 = JSONから指定キーの数値を取得(メッセージ, "from")
            移動先添字 = JSONから指定キーの数値を取得(メッセージ, "to")
            
            画面からの要求でタブの並び順を変更 移動元添字, 移動先添字
            
        Case ""
        
        Case Else
        
    End Select
End Sub

' =========================================================================
' 処理概要       : 画面からの要求で指定タブを選択
' 引数           : 位置番号 (Long) - 対象の位置番号（公開タブ・ログは1始まり）
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 画面からの要求で指定タブを選択(ByVal 位置番号 As Long) As Boolean
    Dim 成功 As Boolean
    
    If 保持ブラウザー Is Nothing Then
    
        画面からの要求で指定タブを選択 = False
        Exit Function
    End If
    成功 = 保持ブラウザー.指定タブを表示して選択状態を更新(位置番号)
    画面からの要求で指定タブを選択 = 成功
End Function

' =========================================================================
' 処理概要       : 画面からの要求で指定タブを閉じる
' 引数           : 位置番号 (Long) - 対象の位置番号（公開タブ・ログは1始まり）
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 画面からの要求で指定タブを閉じる(ByVal 位置番号 As Long) As Boolean
    Dim 成功 As Boolean
    
    If 保持ブラウザー Is Nothing Then
    
        画面からの要求で指定タブを閉じる = False
        Exit Function
    End If
    成功 = 保持ブラウザー.指定タブを閉じて選択状態を更新(位置番号)
    画面からの要求で指定タブを閉じる = 成功
End Function

' =========================================================================
' 処理概要       : 画面からの要求でタブの並び順を変更
' 引数           : 移動元添字 (Long) - 移動元のタブ番号（1始まり）
'                  移動先添字 (Long) - 移動先のタブ番号（1始まり）
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 画面からの要求でタブの並び順を変更(ByVal 移動元添字 As Long, ByVal 移動先添字 As Long) As Boolean
    Dim 成功 As Boolean
    
    If 保持ブラウザー Is Nothing Then
    
        画面からの要求でタブの並び順を変更 = False
        Exit Function
    End If
    成功 = 保持ブラウザー.指定タブの並び順を変更(移動元添字, 移動先添字)
    画面からの要求でタブの並び順を変更 = 成功
End Function

' =========================================================================
' 処理概要       : タブ一覧と選択状態をタブバーへ同期
' =========================================================================
Private Sub タブ一覧と選択状態をタブバーへ同期()
    Dim 件数 As Long
    Dim JSON As String
    Dim 添字 As Long
    Dim 表示 As mod_WV2_ブラウザー表示
    Dim タイトル As String
    Dim 実行結果 As Long
    
    If 保持バー表示 Is Nothing Then Exit Sub
    If 保持ブラウザー Is Nothing Then Exit Sub
    If Not 保持準備完了 Then Exit Sub
    
    件数 = 保持ブラウザー.タブ数
    
    JSON = "{""type"":""sync"",""tabs"":["
    
    For 添字 = 1 To 件数
        Set 表示 = 保持ブラウザー.指定タブのペインを取得(添字)
        If 表示 Is Nothing Then
            タイトル = ""
        Else
            タイトル = 表示.保持するページタイトルを取得
        End If
        If 添字 > 1 Then JSON = JSON & ","
        JSON = JSON & "{""title"":""" & 文字列をJSONに埋め込める形式へ変換(タイトル) & """}"
    Next 添字
    
    JSON = JSON & "],""active"":" & 保持ブラウザー.選択タブ番号 & "}"
    
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
' 処理概要       : タブバーのHTMLをUTF-8で保存
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Private Function タブバーのHTMLをUTF8で保存() As Boolean
    Dim 完全パス As String
    Dim HTML As String
    Dim 作業情報 As Object
    Dim バイナリストリーム As Object
    Dim バイト数() As Byte
    Dim 出力ストリーム As Object
    
    On Error GoTo ErrorHanler
    
    If Dir(保持フォルダーパス, vbDirectory) = "" Then
        MkDir 保持フォルダーパス
        
    End If
    
    完全パス = 保持フォルダーパス
    If Right$(完全パス, 1) <> "\" Then 完全パス = 完全パス & "\"
    完全パス = 完全パス & タブバーHTML名
    
    HTML = タブバーのHTMLを生成()
    
    Set 作業情報 = CreateObject("ADODB.Stream")
    作業情報.Type = 2
    作業情報.Charset = "UTF-8"
    作業情報.Open
    作業情報.WriteText HTML
    
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
    
    作業情報.Close
    
    タブバーのHTMLをUTF8で保存 = True
    Exit Function
    
ErrorHanler:
    MsgBox "処理を完了できません。" & vbCrLf & Err.Description, vbExclamation, "タブバー"
    
    タブバーのHTMLをUTF8で保存 = False
    GoTo EndProc
    
EndProc:
End Function

' =========================================================================
' 処理概要       : タブバーのHTMLを生成
' 戻り値         : String - タブバーのHTML
' =========================================================================
Private Function タブバーのHTMLを生成() As String
    Dim 作成文字列 As String
    
    作成文字列 = "<!DOCTYPE html>" & vbLf
    作成文字列 = 作成文字列 & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    作成文字列 = 作成文字列 & "<style>" & vbLf
    作成文字列 = 作成文字列 & "  html,body{margin:0;padding:0;height:100%;box-sizing:border-box;" & "font-family:'Segoe UI','Meiryo',sans-serif;}" & vbLf
    作成文字列 = 作成文字列 & "  *{box-sizing:border-box;}" & vbLf
    作成文字列 = 作成文字列 & "  body{display:flex;align-items:stretch;overflow:hidden;" & "background:#f3f3f3;border-bottom:1px solid #ccc;}" & vbLf
    作成文字列 = 作成文字列 & "  #tabs{display:flex;align-items:stretch;height:100%;overflow-x:auto;}" & vbLf
    
    作成文字列 = 作成文字列 & "  #tabs::-webkit-scrollbar{display:none;}" & vbLf
    
    作成文字列 = 作成文字列 & "  .tab{display:flex;align-items:center;max-width:200px;min-width:80px;" & "height:100%;padding:0 6px 0 12px;border-right:1px solid #ccc;" & _
        "background:#e7e7e7;cursor:pointer;font-size:13px;color:#333;" & "white-space:nowrap;transition:background-color .12s;}" & vbLf
    作成文字列 = 作成文字列 & "  .tab.active{background:#fff;color:#000;font-weight:600;}" & vbLf
    
    作成文字列 = 作成文字列 & "  .tab:not(.active):hover{background:#efefef;}" & vbLf
    作成文字列 = 作成文字列 & "  .tab .title{overflow:hidden;text-overflow:ellipsis;flex:1;}" & vbLf
    
    作成文字列 = 作成文字列 & "  .tab .close{margin-left:6px;width:18px;height:18px;line-height:18px;" & "text-align:center;border-radius:3px;color:#888;visibility:hidden;}" & vbLf
    
    作成文字列 = 作成文字列 & "  .tab.active .close,.tab:hover .close{visibility:visible;}" & vbLf
    作成文字列 = 作成文字列 & "  .tab .close:hover{background:#ccc;color:#000;}" & vbLf
    作成文字列 = 作成文字列 & "  .tab.dragging{opacity:0.45;}" & vbLf
    作成文字列 = 作成文字列 & "  .tab.dragover{box-shadow:inset 3px 0 0 #4a90e2;}" & vbLf
    作成文字列 = 作成文字列 & "  .tab.dragover-after{box-shadow:inset -3px 0 0 #4a90e2;}" & vbLf
    作成文字列 = 作成文字列 & "  #newtab{display:flex;align-items:center;justify-content:center;" & "width:32px;height:100%;cursor:pointer;font-size:18px;color:#555;}" & vbLf
    作成文字列 = 作成文字列 & "  #newtab:hover{background:#e0e0e0;}" & vbLf
    
    作成文字列 = 作成文字列 & "  #settings{display:flex;align-items:center;justify-content:center;" & "width:34px;height:100%;cursor:pointer;font-size:16px;color:#666;margin-left:auto;}" & vbLf
    作成文字列 = 作成文字列 & "  #settings:hover{background:#e0e0e0;color:#333;}" & vbLf
    作成文字列 = 作成文字列 & "  #empty{padding:0 12px;display:flex;align-items:center;color:#999;font-size:13px;}" & vbLf
    作成文字列 = 作成文字列 & "</style></head><body>" & vbLf
    作成文字列 = 作成文字列 & "  <div id=""tabs""></div>" & vbLf
    作成文字列 = 作成文字列 & "  <div id=""newtab"" title=""新しいタブ"">+</div>" & vbLf
    
    作成文字列 = 作成文字列 & "  <div id=""settings"" title=""設定"">&#9881;</div>" & vbLf
    作成文字列 = 作成文字列 & "<script>" & vbLf
    作成文字列 = 作成文字列 & "  var W = window.chrome && window.chrome.webview;" & vbLf
    作成文字列 = 作成文字列 & "  function postCmd(obj){ if(W){ W.postMessage(JSON.stringify(obj)); } }" & vbLf
    
    作成文字列 = 作成文字列 & "  function HO(){ return (W && W.hostObjects && W.hostObjects.sync) ? W.hostObjects.sync : null; }" & vbLf
    
    作成文字列 = 作成文字列 & "  function markActive(el){" & vbLf
    作成文字列 = 作成文字列 & "    var xs = document.querySelectorAll('.tab.active');" & vbLf
    作成文字列 = 作成文字列 & "    for(var k=0;k<xs.length;k++){ xs[k].classList.remove('active'); }" & vbLf
    作成文字列 = 作成文字列 & "    if(el){ el.classList.add('active'); }" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    
    作成文字列 = 作成文字列 & "  function afterPaint(fn){" & vbLf
    作成文字列 = 作成文字列 & "    if(window.requestAnimationFrame){" & vbLf
    作成文字列 = 作成文字列 & "      requestAnimationFrame(function(){ requestAnimationFrame(fn); });" & vbLf
    作成文字列 = 作成文字列 & "    } else {" & vbLf
    作成文字列 = 作成文字列 & "      setTimeout(fn, 0);" & vbLf
    作成文字列 = 作成文字列 & "    }" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    
    作成文字列 = 作成文字列 & "  function sendActivate(n){" & vbLf
    作成文字列 = 作成文字列 & "    var ho = HO();" & vbLf
    作成文字列 = 作成文字列 & "    if(ho && ho.TabBar){" & vbLf
    作成文字列 = 作成文字列 & "      try{ ho.TabBar.画面からの要求で指定タブを選択(n); }" & vbLf
    作成文字列 = 作成文字列 & "      catch(e){ postCmd({cmd:'activate', index:n}); }" & vbLf
    作成文字列 = 作成文字列 & "    } else {" & vbLf
    作成文字列 = 作成文字列 & "      postCmd({cmd:'activate', index:n});" & vbLf
    作成文字列 = 作成文字列 & "    }" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    
    作成文字列 = 作成文字列 & "  var lastActive = -1;" & vbLf
    作成文字列 = 作成文字列 & "  var dragSrc = null;" & vbLf
    作成文字列 = 作成文字列 & "  function clearMarks(){" & vbLf
    作成文字列 = 作成文字列 & "    var xs = document.querySelectorAll('.dragover, .dragover-after');" & vbLf
    作成文字列 = 作成文字列 & "    for(var k=0;k<xs.length;k++){ xs[k].classList.remove('dragover'); xs[k].classList.remove('dragover-after'); }" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    
    作成文字列 = 作成文字列 & "  function render(state){" & vbLf
    作成文字列 = 作成文字列 & "    var tabsEl = document.getElementById('tabs');" & vbLf
    
    作成文字列 = 作成文字列 & "    var keepLeft = tabsEl.scrollLeft;" & vbLf
    作成文字列 = 作成文字列 & "    tabsEl.innerHTML = '';" & vbLf
    作成文字列 = 作成文字列 & "    var list = (state && state.tabs) ? state.tabs : [];" & vbLf
    作成文字列 = 作成文字列 & "    var active = (state && state.active) ? state.active : 0;" & vbLf
    作成文字列 = 作成文字列 & "    if(list.length === 0){" & vbLf
    作成文字列 = 作成文字列 & "      var e = document.createElement('div');" & vbLf
    作成文字列 = 作成文字列 & "      e.id = 'empty'; e.textContent = '(タブなし)';" & vbLf
    作成文字列 = 作成文字列 & "      tabsEl.appendChild(e); return;" & vbLf
    作成文字列 = 作成文字列 & "    }" & vbLf
    作成文字列 = 作成文字列 & "    for(var i=0; i<list.length; i++){" & vbLf
    作成文字列 = 作成文字列 & "      var idx = i + 1;" & vbLf
    作成文字列 = 作成文字列 & "      var t = list[i] || {};" & vbLf
    作成文字列 = 作成文字列 & "      var title = (t.title && t.title.length) ? t.title : '(読み込み中)';" & vbLf
    作成文字列 = 作成文字列 & "      var tab = document.createElement('div');" & vbLf
    作成文字列 = 作成文字列 & "      tab.className = 'tab' + (idx === active ? ' active' : '');" & vbLf
    作成文字列 = 作成文字列 & "      tab.setAttribute('data-index', idx);" & vbLf
    作成文字列 = 作成文字列 & "      tab.draggable = true;" & vbLf
    作成文字列 = 作成文字列 & "      var ti = document.createElement('span');" & vbLf
    作成文字列 = 作成文字列 & "      ti.className = 'title'; ti.textContent = title; ti.title = title;" & vbLf
    作成文字列 = 作成文字列 & "      var cl = document.createElement('span');" & vbLf
    作成文字列 = 作成文字列 & "      cl.className = 'close'; cl.textContent = '×';" & vbLf
    
    作成文字列 = 作成文字列 & "      tab.onclick = function(ev){" & vbLf
    作成文字列 = 作成文字列 & "        if(ev.target && ev.target.className === 'close') return;" & vbLf
    作成文字列 = 作成文字列 & "        var n = parseInt(this.getAttribute('data-index'), 10);" & vbLf
    作成文字列 = 作成文字列 & "        markActive(this);" & vbLf
    作成文字列 = 作成文字列 & "        afterPaint(function(){ sendActivate(n); });" & vbLf
    作成文字列 = 作成文字列 & "      };" & vbLf
    
    作成文字列 = 作成文字列 & "      cl.onclick = function(ev){" & vbLf
    作成文字列 = 作成文字列 & "        ev.stopPropagation();" & vbLf
    作成文字列 = 作成文字列 & "        var n = parseInt(this.parentNode.getAttribute('data-index'), 10);" & vbLf
    作成文字列 = 作成文字列 & "        var ho = HO();" & vbLf
    作成文字列 = 作成文字列 & "        if(ho && ho.TabBar){" & vbLf
    作成文字列 = 作成文字列 & "          try{ ho.TabBar.画面からの要求で指定タブを閉じる(n); }" & vbLf
    作成文字列 = 作成文字列 & "          catch(e){ postCmd({cmd:'close', index:n}); }" & vbLf
    作成文字列 = 作成文字列 & "        } else {" & vbLf
    作成文字列 = 作成文字列 & "          postCmd({cmd:'close', index:n});" & vbLf
    作成文字列 = 作成文字列 & "        }" & vbLf
    作成文字列 = 作成文字列 & "      };" & vbLf
    
    作成文字列 = 作成文字列 & "      tab.ondragstart = function(ev){" & vbLf
    作成文字列 = 作成文字列 & "        dragSrc = parseInt(this.getAttribute('data-index'),10);" & vbLf
    作成文字列 = 作成文字列 & "        this.classList.add('dragging');" & vbLf
    作成文字列 = 作成文字列 & "        if(ev.dataTransfer){ ev.dataTransfer.effectAllowed='move'; try{ev.dataTransfer.setData('text/plain', String(dragSrc));}catch(e){} }" & vbLf
    作成文字列 = 作成文字列 & "      };" & vbLf
    作成文字列 = 作成文字列 & "      tab.ondragover = function(ev){" & vbLf
    作成文字列 = 作成文字列 & "        ev.preventDefault();" & vbLf
    作成文字列 = 作成文字列 & "        if(ev.dataTransfer){ ev.dataTransfer.dropEffect='move'; }" & vbLf
    作成文字列 = 作成文字列 & "        clearMarks();" & vbLf
    作成文字列 = 作成文字列 & "        var r = this.getBoundingClientRect();" & vbLf
    作成文字列 = 作成文字列 & "        var after = (ev.clientX - r.left) > (r.width/2);" & vbLf
    作成文字列 = 作成文字列 & "        this.classList.add(after ? 'dragover-after' : 'dragover');" & vbLf
    作成文字列 = 作成文字列 & "      };" & vbLf
    作成文字列 = 作成文字列 & "      tab.ondragleave = function(ev){ this.classList.remove('dragover'); this.classList.remove('dragover-after'); };" & vbLf
    作成文字列 = 作成文字列 & "      tab.ondrop = function(ev){" & vbLf
    作成文字列 = 作成文字列 & "        ev.preventDefault();" & vbLf
    作成文字列 = 作成文字列 & "        var target = parseInt(this.getAttribute('data-index'),10);" & vbLf
    作成文字列 = 作成文字列 & "        var r = this.getBoundingClientRect();" & vbLf
    作成文字列 = 作成文字列 & "        var after = (ev.clientX - r.left) > (r.width/2);" & vbLf
    作成文字列 = 作成文字列 & "        clearMarks();" & vbLf
    作成文字列 = 作成文字列 & "        if(dragSrc == null) return;" & vbLf
    
    作成文字列 = 作成文字列 & "        var dest = after ? target + 1 : target;" & vbLf
    作成文字列 = 作成文字列 & "        if(dragSrc < dest){ dest = dest - 1; }" & vbLf
    作成文字列 = 作成文字列 & "        if(dest < 1) dest = 1;" & vbLf
    
    作成文字列 = 作成文字列 & "        if(dragSrc !== dest){" & vbLf
    作成文字列 = 作成文字列 & "          var ho = HO();" & vbLf
    作成文字列 = 作成文字列 & "          if(ho && ho.TabBar){" & vbLf
    作成文字列 = 作成文字列 & "            try{ ho.TabBar.画面からの要求でタブの並び順を変更(dragSrc, dest); }" & vbLf
    作成文字列 = 作成文字列 & "            catch(e){ postCmd({cmd:'reorder', from:dragSrc, to:dest}); }" & vbLf
    作成文字列 = 作成文字列 & "          } else {" & vbLf
    作成文字列 = 作成文字列 & "            postCmd({cmd:'reorder', from:dragSrc, to:dest});" & vbLf
    作成文字列 = 作成文字列 & "          }" & vbLf
    作成文字列 = 作成文字列 & "        }" & vbLf
    作成文字列 = 作成文字列 & "        dragSrc = null;" & vbLf
    作成文字列 = 作成文字列 & "      };" & vbLf
    作成文字列 = 作成文字列 & "      tab.ondragend = function(ev){ this.classList.remove('dragging'); clearMarks(); dragSrc = null; };" & vbLf
    作成文字列 = 作成文字列 & "      tab.appendChild(ti); tab.appendChild(cl);" & vbLf
    作成文字列 = 作成文字列 & "      tabsEl.appendChild(tab);" & vbLf
    作成文字列 = 作成文字列 & "    }" & vbLf
    
    作成文字列 = 作成文字列 & "    tabsEl.scrollLeft = keepLeft;" & vbLf
    作成文字列 = 作成文字列 & "    if(active !== lastActive){" & vbLf
    作成文字列 = 作成文字列 & "      var act = tabsEl.querySelector('.tab.active');" & vbLf
    作成文字列 = 作成文字列 & "      if(act && act.scrollIntoView){" & vbLf
    作成文字列 = 作成文字列 & "        try{ act.scrollIntoView({block:'nearest', inline:'nearest'}); }catch(e){}" & vbLf
    作成文字列 = 作成文字列 & "      }" & vbLf
    作成文字列 = 作成文字列 & "    }" & vbLf
    作成文字列 = 作成文字列 & "    lastActive = active;" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    
    作成文字列 = 作成文字列 & "  var tabsRoot = document.getElementById('tabs');" & vbLf
    作成文字列 = 作成文字列 & "  if(tabsRoot){" & vbLf
    作成文字列 = 作成文字列 & "    tabsRoot.addEventListener('wheel', function(ev){" & vbLf
    作成文字列 = 作成文字列 & "      if(tabsRoot.scrollWidth <= tabsRoot.clientWidth) return;" & vbLf
    作成文字列 = 作成文字列 & "      var d = ev.deltaY || 0;" & vbLf
    作成文字列 = 作成文字列 & "      if(d === 0) return;" & vbLf
    作成文字列 = 作成文字列 & "      tabsRoot.scrollLeft += d;" & vbLf
    作成文字列 = 作成文字列 & "      ev.preventDefault();" & vbLf
    作成文字列 = 作成文字列 & "    }, {passive:false});" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    作成文字列 = 作成文字列 & "  document.getElementById('newtab').onclick = function(){" & vbLf
    作成文字列 = 作成文字列 & "    postCmd({cmd:'newtab'});" & vbLf
    作成文字列 = 作成文字列 & "  };" & vbLf
    
    作成文字列 = 作成文字列 & "  document.getElementById('settings').onclick = function(){" & vbLf
    作成文字列 = 作成文字列 & "    postCmd({cmd:'settings'});" & vbLf
    作成文字列 = 作成文字列 & "  };" & vbLf
    
    作成文字列 = 作成文字列 & "  if(W){" & vbLf
    作成文字列 = 作成文字列 & "    W.addEventListener('message', function(ev){" & vbLf
    作成文字列 = 作成文字列 & "      var data = ev.data;" & vbLf
    作成文字列 = 作成文字列 & "      if(typeof data === 'string'){ try{ data = JSON.parse(data); }catch(e){ return; } }" & vbLf
    作成文字列 = 作成文字列 & "      if(data && data.type === 'sync'){ render(data); }" & vbLf
    作成文字列 = 作成文字列 & "    });" & vbLf
    作成文字列 = 作成文字列 & "  }" & vbLf
    
    作成文字列 = 作成文字列 & "  render({tabs:[], active:0});" & vbLf
    作成文字列 = 作成文字列 & "  postCmd({cmd:'ready'});" & vbLf
    作成文字列 = 作成文字列 & "</script>" & vbLf
    作成文字列 = 作成文字列 & "</body></html>"
    タブバーのHTMLを生成 = 作成文字列
End Function
