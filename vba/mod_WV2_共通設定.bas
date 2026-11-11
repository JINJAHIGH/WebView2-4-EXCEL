Attribute VB_Name = "mod_WV2_共通設定"
Option Explicit

Private 保持データルート As String
Private 保持初期URL As String

' =========================================================================
' 処理概要       : プロファイルと設定の保存先フォルダーを取得
' 戻り値         : String - プロファイルと設定の保存先フォルダー
' =========================================================================
Public Property Get 保存先フォルダー() As String
    Dim 基準パス As String, ブック名称 As String, 無効文字 As Variant
    
    On Error GoTo ErrorHanler
    If Len(保持データルート) = 0 Then
    
        ' 保存先の指定がなければ、ユーザーごとの既定フォルダーを使用する
        基準パス = Environ$("LOCALAPPDATA")
        If Len(基準パス) = 0 Then 基準パス = Environ$("TEMP")
        If Len(基準パス) = 0 Then MsgBox "保存先フォルダーを指定してください。", vbExclamation, "保存先設定": GoTo EndProc
        ブック名称 = ThisWorkbook.name
        
        ' ブック名をフォルダー名に使える形式へ変換する
        For Each 無効文字 In Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
            ブック名称 = Replace(ブック名称, CStr(無効文字), "_")
        Next 無効文字
        保持データルート = 基準パス & "\WebView2ForExcel\" & ブック名称
    End If
    保存先フォルダー = 保持データルート
    GoTo EndProc
    
ErrorHanler:
    MsgBox "保存先を設定できません。" & vbCrLf & Err.Description, vbExclamation, "保存先設定"
    GoTo EndProc
    
EndProc:
End Property

' =========================================================================
' 処理概要       : プロファイルと設定の保存先フォルダーを設定
' 引数           : フォルダーパス (String) - 対象の保存先またはHTMLフォルダー
' =========================================================================
Public Property Let 保存先フォルダー(ByVal フォルダーパス As String)
    Dim ファイル操作 As Object, 解決済みパス As String
    
    On Error GoTo ErrorHanler
    If Len(Trim$(フォルダーパス)) = 0 Then MsgBox "保存先フォルダーを指定してください。", vbExclamation, "保存先設定": GoTo EndProc
    
    ' 相対パスを絶対パスへ変換し、末尾の区切り文字を整える
    Set ファイル操作 = CreateObject("Scripting.FileSystemObject")
    解決済みパス = ファイル操作.GetAbsolutePathName(フォルダーパス)
    If Len(解決済みパス) > 3 And Right$(解決済みパス, 1) = "\" Then 解決済みパス = Left$(解決済みパス, Len(解決済みパス) - 1)
    保持データルート = 解決済みパス
    GoTo EndProc
    
ErrorHanler:
    MsgBox "保存先を設定できません。" & vbCrLf & Err.Description, vbExclamation, "保存先設定"
    GoTo EndProc
    
EndProc:
End Property

' =========================================================================
' 処理概要       : WebView2プロファイルの保存先を取得
' 戻り値         : String - WebView2プロファイルの保存先
' =========================================================================
Public Property Get プロファイル保存先() As String
    プロファイル保存先 = 保存先フォルダー & "\profile"
End Property

' =========================================================================
' 処理概要       : 画面部品のHTML保存先を作成して取得
' 引数           : 区分 (String) - 画面部品の区分（tabbar / navbar）
' 戻り値         : String - 画面部品のHTML保存先を作成して取得
' =========================================================================
Public Function 画面部品の保存先(ByVal 区分 As String) As String
    On Error GoTo ErrorHanler
    Select Case 区分
        Case "tabbar", "navbar"
        Case Else: MsgBox "保存先の区分が不正です。", vbExclamation, "保存先設定": GoTo EndProc
    End Select
    画面部品の保存先 = 保存先フォルダー & "\assets\" & 区分
    mod_WV2_ネイティブ連携.必要な保存先フォルダーを作成 画面部品の保存先
    GoTo EndProc
    
ErrorHanler:
    MsgBox "保存先を設定できません。" & vbCrLf & Err.Description, vbExclamation, "保存先設定"
    GoTo EndProc
    
EndProc:
End Function

' =========================================================================
' 処理概要       : 起動時と新規タブに表示するURLを取得
' 戻り値         : String - 起動時と新規タブに表示するURL
' =========================================================================
Public Property Get 初期URL() As String
    If Len(保持初期URL) = 0 Then 保持初期URL = "about:blank"
    初期URL = 保持初期URL
End Property

' =========================================================================
' 処理概要       : 起動時と新規タブに表示するURLを設定
' 引数           : URL (String) - 移動先または通知対象のURL
' =========================================================================
Public Property Let 初期URL(ByVal URL As String)
    保持初期URL = Trim$(URL)
    If Len(保持初期URL) = 0 Then 保持初期URL = "about:blank"
End Property
