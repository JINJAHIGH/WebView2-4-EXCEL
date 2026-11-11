VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_WV2_ページ表示
   Caption         =   "WebView2 動作確認"
   ClientHeight    =   8850.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12375
   OleObjectBlob   =   "frm_WV2_ページ表示.frx":0000
   StartUpPosition =   1
End
Attribute VB_Name = "frm_WV2_ページ表示"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private 保持ブラウザー As mod_WV2_ブラウザー管理

' =========================================================================
' 処理概要       : フォームの表示枠にWebView2を配置して指定URLを開く
' 引数           : URL (String) - 表示するページのURL
' 戻り値         : mod_WV2_ブラウザー表示 - 作成した表示。起動できない場合 Nothing
' =========================================================================
Public Function ページを開く(ByVal URL As String) As mod_WV2_ブラウザー表示
    Dim エラー内容 As String
    
    On Error GoTo ErrorHanler
    
    ' フレームをフォーム全体に広げ、ページを表示する領域を準備する
    表示枠.BorderStyle = fmBorderStyleNone
    表示枠.Caption = ""
    UserForm_Resize
    
    ' フレームのウィンドウハンドルをWebView2の親ウィンドウとして使用する
    Set 保持ブラウザー = New mod_WV2_ブラウザー管理
    If Not 保持ブラウザー.初期化(表示枠.[_GethWnd], mod_WV2_共通設定.プロファイル保存先) Then
        エラー内容 = "WebView2を初期化できませんでした。"
        GoTo ErrorHanler
    End If
    
    Set ページを開く = 保持ブラウザー.新しいタブを作成して指定URLを表示(URL)
    If ページを開く Is Nothing Then
        エラー内容 = "ページの表示を開始できませんでした。"
        GoTo ErrorHanler
    End If
    GoTo EndProc
    
ErrorHanler:
    If Len(エラー内容) = 0 Then エラー内容 = Err.Description
    MsgBox エラー内容, vbExclamation, "WebView2 動作確認"
    
EndProc:
End Function

' =========================================================================
' 処理概要       : フォームの内側全体に表示枠とページを配置する
' =========================================================================
Private Sub UserForm_Resize()
    Dim 表示領域 As 矩形
    
    ' フォームの内寸に合わせてフレームを配置する
    表示枠.Move 0, 0, Me.InsideWidth, Me.InsideHeight
    
    ' WebView2にはピクセル単位の表示領域を渡す
    If 保持ブラウザー Is Nothing Then Exit Sub
    
    If クライアント矩形取得API(表示枠.[_GethWnd], 表示領域) <> 0 Then 保持ブラウザー.全タブの表示領域を更新 表示領域
End Sub

' =========================================================================
' 処理概要       : 画面を閉じる前にタイマーとブラウザーの参照を解放する
' 引数           : Cancel (Integer) - Trueを設定すると画面の終了を中止する
'                  CloseMode (Integer) - フォームが閉じられた方法
' =========================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrorHanler
    
    ' 表示枠が破棄される前にタイマーを停止し、ペインからブラウザーへの参照を解除する
    If Not 保持ブラウザー Is Nothing Then 保持ブラウザー.タイマーとペインを停止して循環参照を解除
    GoTo EndProc
    
ErrorHanler:
    MsgBox "ブラウザーの終了処理でエラーが発生しました。" & vbCrLf & Err.Description, vbExclamation, "WebView2 動作確認"
    Resume EndProc
    
EndProc:
    Set 保持ブラウザー = Nothing
End Sub
