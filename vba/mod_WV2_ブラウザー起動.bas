Attribute VB_Name = "mod_WV2_ブラウザー起動"
Option Explicit

' =========================================================================
' 処理概要       : 既定の設定でブラウザー画面を開く
' =========================================================================
Public Sub 起動()
    ブラウザーを開く
End Sub

' =========================================================================
' 処理概要       : ブラウザー画面を開き初期URLを表示
' 引数           : 初期URL (String) - 最初に表示するURL（省略時: ""）
'                  データフォルダー (String) - プロファイルと設定の保存先フォルダー（省略時: ""）
' =========================================================================
Public Sub ブラウザーを開く(Optional ByVal 初期URL As String = "", Optional ByVal データフォルダー As String = "")
    On Error GoTo ErrorHanler
    
    ' Excelの操作を続けられるよう、画面をモードレスで表示する
    frm_WV2_ブラウザーフォーム.Show vbModeless
    
    ' 起動状態に応じて、画面の初期化または表示中ページの移動を行う
    If frm_WV2_ブラウザーフォーム.起動中のブラウザー Is Nothing Then
        frm_WV2_ブラウザーフォーム.ブラウザー画面を開始 初期URL, データフォルダー
        If frm_WV2_ブラウザーフォーム.起動中のブラウザー Is Nothing Then Unload frm_WV2_ブラウザーフォーム
    ElseIf Len(初期URL) > 0 Then
        frm_WV2_ブラウザーフォーム.起動中のブラウザー.入力をURLに解決して選択中のタブへ移動 初期URL
    End If
    GoTo EndProc
    
ErrorHanler:
    MsgBox "ブラウザーを開けません。" & vbCrLf & Err.Description, vbExclamation, "ブラウザー"
    On Error Resume Next
    Unload frm_WV2_ブラウザーフォーム
    GoTo EndProc
    
EndProc:
End Sub

' =========================================================================
' 処理概要       : 起動中のブラウザー画面とログファイルを閉じる
' =========================================================================
Public Sub ブラウザーを閉じる()
    Dim フォーム As Object
    
    ' 読み込み済みのフォームから対象を探す
    For Each フォーム In VBA.UserForms
        If TypeName(フォーム) = "frm_WV2_ブラウザーフォーム" Then
            Unload フォーム
            Exit For
        End If
    Next フォーム
    mod_WV2_ログ管理.ログファイルを閉じる
End Sub

' =========================================================================
' 処理概要       : 起動中のブラウザーを取得（フォームの新規生成は行わない）
' 戻り値         : mod_WV2_ブラウザー管理 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Public Property Get 起動中のブラウザー() As mod_WV2_ブラウザー管理
    Dim フォーム As Object
    
    ' 読み込み済みのフォームから対象を探す
    For Each フォーム In VBA.UserForms
        If TypeName(フォーム) = "frm_WV2_ブラウザーフォーム" Then
            Set 起動中のブラウザー = フォーム.起動中のブラウザー
            Exit Property
        End If
    Next フォーム
End Property
