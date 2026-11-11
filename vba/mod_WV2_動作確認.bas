Attribute VB_Name = "mod_WV2_動作確認"
Option Explicit

' =========================================================================
' 処理概要       : example.comへアクセスしてページのタイトルを表示
' =========================================================================
Public Sub サンプルサイトのタイトルを表示()
    Dim 対象ブラウザー As mod_WV2_ブラウザー管理
    Dim 対象表示 As mod_WV2_ブラウザー表示
    Dim ページタイトル As String
    Dim 開始時刻 As Single
    Dim 経過秒数 As Single
    
    On Error GoTo ErrorHanler
    
    ' -------------
    ' サンプルサイトを表示して読込完了を待機
    ' -------------
    mod_WV2_ブラウザー起動.ブラウザーを開く "https://example.com/"
    Set 対象ブラウザー = mod_WV2_ブラウザー起動.起動中のブラウザー
    If 対象ブラウザー Is Nothing Then GoTo EndProc
    
    Set 対象表示 = 対象ブラウザー.選択中の表示
    If 対象表示 Is Nothing Then GoTo EndProc
    
    開始時刻 = Timer
    Do
        DoEvents
        If 対象表示.利用可能 And 対象表示.読込完了 Then
            If 対象表示.表示中のページのURLを取得() = "https://example.com/" Then
                ページタイトル = 対象表示.表示中のページのタイトルを取得()
                If Len(ページタイトル) > 0 Then Exit Do
            End If
        End If
        経過秒数 = Timer - 開始時刻
        If 経過秒数 < 0 Then 経過秒数 = 経過秒数 + 86400
        If 経過秒数 >= 30 Then
            MsgBox "ページのタイトルを取得できませんでした。", vbExclamation, "動作確認"
            GoTo EndProc
        End If
    Loop
    MsgBox ページタイトル, vbInformation, "ページタイトル"
    GoTo EndProc
    
ErrorHanler:
    MsgBox "タイトル取得でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "動作確認"
    Resume EndProc
    
EndProc:
    Set 対象表示 = Nothing
    Set 対象ブラウザー = Nothing
End Sub

' =========================================================================
' 処理概要       : ページだけを表示する画面でexample.comを開き、現在のURLと説明文を表示する
' =========================================================================
Public Sub サンプルサイトのURLと本文を表示()
    Const 接続先 As String = "https://example.com/"
    Const 読込待機上限秒 As Single = 30
    Dim 確認画面 As frm_WV2_ページ表示
    Dim 対象表示 As mod_WV2_ブラウザー表示
    Dim 本文要素 As mod_WV2_HTML要素
    Dim 現在URL As String
    Dim 本文 As String
    Dim 開始時刻 As Single
    Dim 経過秒数 As Single
    
    On Error GoTo ErrorHanler
    
    ' -------------
    ' ページ表示専用フォームを開く
    ' -------------
    Set 確認画面 = New frm_WV2_ページ表示
    確認画面.Show vbModeless
    Set 対象表示 = 確認画面.ページを開く(接続先)
    If 対象表示 Is Nothing Then GoTo EndProc
    
    ' -------------
    ' 遷移完了を待ち、WebView2から現在のURLを取得する
    ' -------------
    開始時刻 = Timer
    Do
        DoEvents
        If Not 確認画面.visible Then GoTo EndProc
        If 対象表示.読込完了 Then
            現在URL = 対象表示.表示中のページのURLを取得()
            If 現在URL = 接続先 Then Exit Do
        End If
        
        ' 日付が変わるとTimerが0に戻るため、日付をまたいだ分を補正する
        経過秒数 = Timer - 開始時刻
        If 経過秒数 < 0 Then 経過秒数 = 経過秒数 + 86400
        If 経過秒数 >= 読込待機上限秒 Then
            MsgBox "指定したページの読み込みを確認できませんでした。", vbExclamation, "WebView2 動作確認"
            GoTo EndProc
        End If
    Loop
    
    ' -------------
    ' 説明文をDOMから取得し、現在のURLと一緒に表示する
    ' -------------
    Set 本文要素 = 対象表示.CSSセレクターで最初のDOM要素を検索("body > p")
    If Not 本文要素 Is Nothing Then 本文 = Trim$(本文要素.DOM要素の表示テキストを取得)
    If Len(本文) = 0 Then
        MsgBox "ページの説明文を取得できませんでした。" & vbCrLf & 現在URL, vbExclamation, "WebView2 動作確認"
        GoTo EndProc
    End If
    
    MsgBox "現在のURL" & vbCrLf & 現在URL & vbCrLf & vbCrLf & "本文" & vbCrLf & 本文, vbInformation, "WebView2 動作確認"
    GoTo EndProc
    
ErrorHanler:
    MsgBox "動作確認でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "WebView2 動作確認"
    Resume EndProc
    
EndProc:
    ' DOM要素、ペイン、フォームの順に参照を解放する
    Set 本文要素 = Nothing
    Set 対象表示 = Nothing
    If Not 確認画面 Is Nothing Then Unload 確認画面
    Set 確認画面 = Nothing
End Sub
