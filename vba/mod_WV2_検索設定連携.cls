VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "mod_WV2_検索設定連携"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' -------------
' URL設定
' -------------
Private Const Google検索URL     As String = "https://www.google.com/search?q="
Private Const Bing検索URL       As String = "https://www.bing.com/search?q="
Private Const DuckDuckGo検索URL As String = "https://duckduckgo.com/?q="
Private Const Yahoo検索URL      As String = "https://search.yahoo.com/search?p="
Private Const 検索URL設定キー   As String = "template"

' -------------
' 定数
' -------------
Private Const 設定ファイル名       As String = "settings.txt"
Private Const 検索エンジン設定キー As String = "engine"

Private 保持ブラウザー As mod_WV2_ブラウザー管理

' =========================================================================
' 処理概要       : 設定の読み書き対象となるブラウザーを設定
' 引数           : 対象ブラウザー (mod_WV2_ブラウザー管理) - 設定を反映するブラウザー
' =========================================================================
Public Sub 設定の読み書き対象となるブラウザーを設定(ByVal 対象ブラウザー As mod_WV2_ブラウザー管理)
    Set 保持ブラウザー = 対象ブラウザー
End Sub

' =========================================================================
' 処理概要       : 検索エンジン名をURLテンプレートへ変換
' 引数           : 検索エンジン名称 (String) - 検索エンジンのプリセット名
' 戻り値         : String - 検索エンジン名をURLテンプレート
' =========================================================================
Private Function 検索エンジン名をURLテンプレートへ変換(ByVal 検索エンジン名称 As String) As String
    Dim キー As String
    
    キー = LCase$(Trim$(検索エンジン名称))
    Select Case キー
        Case "google":            検索エンジン名をURLテンプレートへ変換 = Google検索URL
        Case "bing":              検索エンジン名をURLテンプレートへ変換 = Bing検索URL
        Case "duckduckgo", "ddg": 検索エンジン名をURLテンプレートへ変換 = DuckDuckGo検索URL
        Case "yahoo":             検索エンジン名をURLテンプレートへ変換 = Yahoo検索URL
        Case Else:                検索エンジン名をURLテンプレートへ変換 = ""
    End Select
End Function

' =========================================================================
' 処理概要       : URLテンプレートに対応する検索エンジン名を取得
' 引数           : テンプレート (String) - 検索語を末尾に連結するURLテンプレート
' 戻り値         : String - URLテンプレートに対応する検索エンジン名
' =========================================================================
Private Function URLテンプレートに対応する検索エンジン名を取得(ByVal テンプレート As String) As String
    Select Case テンプレート
        Case Google検索URL:     URLテンプレートに対応する検索エンジン名を取得 = "google"
        Case Bing検索URL:       URLテンプレートに対応する検索エンジン名を取得 = "bing"
        Case DuckDuckGo検索URL: URLテンプレートに対応する検索エンジン名を取得 = "duckduckgo"
        Case Yahoo検索URL:      URLテンプレートに対応する検索エンジン名を取得 = "yahoo"
        Case Else:          URLテンプレートに対応する検索エンジン名を取得 = "custom"
    End Select
End Function

' =========================================================================
' 処理概要       : 現在の検索エンジン名を取得
' 戻り値         : String - 現在の検索エンジン名
' =========================================================================
Public Function 現在の検索エンジン名を取得() As String
    If 保持ブラウザー Is Nothing Then
        現在の検索エンジン名を取得 = ""
        Exit Function
    End If
    現在の検索エンジン名を取得 = URLテンプレートに対応する検索エンジン名を取得(保持ブラウザー.検索URL)
End Function

' =========================================================================
' 処理概要       : 指定プリセットの検索エンジンを適用して保存
' 引数           : 検索エンジン名称 (String) - 検索エンジンのプリセット名
' 戻り値         : String - 指定プリセットの検索エンジンを適用して保存
' =========================================================================
Public Function 指定プリセットの検索エンジンを適用して保存(ByVal 検索エンジン名称 As String) As String
    Dim テンプレート As String
    
    If 保持ブラウザー Is Nothing Then
    
        指定プリセットの検索エンジンを適用して保存 = ""
        Exit Function
    End If
    
    テンプレート = 検索エンジン名をURLテンプレートへ変換(検索エンジン名称)
    
    If Len(テンプレート) = 0 Then
    
        テンプレート = Google検索URL
    End If
    
    保持ブラウザー.検索URL = テンプレート
    指定プリセットの検索エンジンを適用して保存 = URLテンプレートに対応する検索エンジン名を取得(保持ブラウザー.検索URL)
    
    If 指定プリセットの検索エンジンを適用して保存 <> "custom" And Len(指定プリセットの検索エンジンを適用して保存) > 0 Then
        検索エンジン名を設定ファイルへ保存 指定プリセットの検索エンジンを適用して保存
    End If
End Function

' =========================================================================
' 処理概要       : 独自検索URLがHTTPまたはHTTPS形式か確認
' 引数           : テンプレート (String) - 検索語を末尾に連結するURLテンプレート
' 戻り値         : String - 独自検索URLがHTTPまたはHTTPS形式か確認
' =========================================================================
Private Function 独自検索URLがHTTPまたはHTTPS形式か確認(ByVal テンプレート As String) As String
    Dim 開始時刻 As String
    Dim 経度値 As String
    
    独自検索URLがHTTPまたはHTTPS形式か確認 = ""
    
    開始時刻 = Trim$(テンプレート)
    If Len(開始時刻) = 0 Then Exit Function
    
    経度値 = LCase$(開始時刻)
    If Left$(経度値, 7) = "http://" Or Left$(経度値, 8) = "https://" Then
        独自検索URLがHTTPまたはHTTPS形式か確認 = 開始時刻
    End If
End Function

' =========================================================================
' 処理概要       : 独自の検索URLを適用して保存
' 引数           : テンプレート (String) - 検索語を末尾に連結するURLテンプレート
' 戻り値         : String - 独自の検索URLを適用して保存
' =========================================================================
Public Function 独自の検索URLを適用して保存(ByVal テンプレート As String) As String
    Dim 開始時刻 As String
    Dim 場所名 As String
    
    独自の検索URLを適用して保存 = ""
    
    If 保持ブラウザー Is Nothing Then
    
        Exit Function
    End If
    
    開始時刻 = 独自検索URLがHTTPまたはHTTPS形式か確認(テンプレート)
    If Len(開始時刻) = 0 Then
    
        Exit Function
    End If
    
    保持ブラウザー.検索URL = 開始時刻
    
    場所名 = URLテンプレートに対応する検索エンジン名を取得(保持ブラウザー.検索URL)
    
    If 場所名 = "custom" Then
        検索エンジン名と独自検索URLを保存 "custom", 開始時刻
    Else
    
        検索エンジン名と独自検索URLを保存 場所名, ""
        
    End If
    
    独自の検索URLを適用して保存 = 場所名
    
End Function

' =========================================================================
' 処理概要       : 独自の検索URLテンプレートを取得
' 戻り値         : String - 独自の検索URLテンプレート
' =========================================================================
Public Function 独自の検索URLテンプレートを取得() As String
    Dim 開始時刻 As String
    Dim 現在値 As String
    
    開始時刻 = ""
    
    If Not 保持ブラウザー Is Nothing Then
        現在値 = 保持ブラウザー.検索URL
        If URLテンプレートに対応する検索エンジン名を取得(現在値) = "custom" Then 開始時刻 = 現在値
    End If
    
    If Len(開始時刻) = 0 Then 開始時刻 = 保存済みの独自検索URLを取得()
    
    独自の検索URLテンプレートを取得 = 開始時刻
End Function

' =========================================================================
' 処理概要       : 現在の設定で検索URLを生成
' 引数           : 検索語 (String) - URLに反映する検索語
' 戻り値         : String - 現在の設定で検索URL
' =========================================================================
Public Function 現在の設定で検索URLを生成(ByVal 検索語 As String) As String
    If 保持ブラウザー Is Nothing Then
        現在の設定で検索URLを生成 = ""
        Exit Function
    End If
    現在の設定で検索URLを生成 = 保持ブラウザー.移動先URLを解決(検索語)
End Function

' =========================================================================
' 処理概要       : 指定プリセットで検索URLを生成
' 引数           : 検索エンジン名称 (String) - 検索エンジンのプリセット名
'                  検索語 (String) - URLに反映する検索語
' 戻り値         : String - 指定プリセットで検索URL
' =========================================================================
Public Function 指定プリセットで検索URLを生成(ByVal 検索エンジン名称 As String, ByVal 検索語 As String) As String
    Dim テンプレート As String
    
    If 保持ブラウザー Is Nothing Then
        指定プリセットで検索URLを生成 = ""
        Exit Function
    End If
    
    テンプレート = 検索エンジン名をURLテンプレートへ変換(検索エンジン名称)
    If Len(テンプレート) = 0 Then テンプレート = Google検索URL
    
    指定プリセットで検索URLを生成 = テンプレート & mod_WV2_URL共通.文字列をUTF8でURLエンコード(検索語)
End Function

' =========================================================================
' 処理概要       : 指定テンプレートで検索URLを生成
' 引数           : テンプレート (String) - 検索語を末尾に連結するURLテンプレート
'                  検索語 (String) - URLに反映する検索語
' 戻り値         : String - 指定テンプレートで検索URL
' =========================================================================
Public Function 指定テンプレートで検索URLを生成(ByVal テンプレート As String, ByVal 検索語 As String) As String
    Dim 開始時刻 As String
    
    指定テンプレートで検索URLを生成 = ""
    
    開始時刻 = 独自検索URLがHTTPまたはHTTPS形式か確認(テンプレート)
    If Len(開始時刻) = 0 Then Exit Function
    
    指定テンプレートで検索URLを生成 = 開始時刻 & mod_WV2_URL共通.文字列をUTF8でURLエンコード(検索語)
End Function

' =========================================================================
' 処理概要       : 検索エンジンのプリセット一覧をJSONで取得
' 戻り値         : String - 検索エンジンのプリセット一覧をJSONで取得
' =========================================================================
Public Function 検索エンジンのプリセット一覧をJSONで取得() As String
    Dim 作成文字列 As String
    
    作成文字列 = "["
    作成文字列 = 作成文字列 & "{""id"":""google"",""label"":""Google"",""template"":""" & Google検索URL & """},"
    作成文字列 = 作成文字列 & "{""id"":""bing"",""label"":""Bing"",""template"":""" & Bing検索URL & """},"
    作成文字列 = 作成文字列 & "{""id"":""duckduckgo"",""label"":""DuckDuckGo"",""template"":""" & DuckDuckGo検索URL & """},"
    作成文字列 = 作成文字列 & "{""id"":""yahoo"",""label"":""Yahoo"",""template"":""" & Yahoo検索URL & """}"
    作成文字列 = 作成文字列 & "]"
    検索エンジンのプリセット一覧をJSONで取得 = 作成文字列
End Function

' =========================================================================
' 処理概要       : 検索設定ファイルのパスを取得
' 戻り値         : String - 検索設定ファイルのパス
' =========================================================================
Private Function 検索設定ファイルのパスを取得() As String
    検索設定ファイルのパスを取得 = mod_WV2_共通設定.保存先フォルダー & "\" & 設定ファイル名
End Function

' =========================================================================
' 処理概要       : 必要な保存先フォルダーを作成
' =========================================================================
Private Sub 必要な保存先フォルダーを作成()
    mod_WV2_ネイティブ連携.必要な保存先フォルダーを作成 mod_WV2_共通設定.保存先フォルダー
End Sub

' =========================================================================
' 処理概要       : 保存済みの検索エンジン名を取得
' 戻り値         : String - 保存済みの検索エンジン名
' =========================================================================
Public Function 保存済みの検索エンジン名を取得() As String
    保存済みの検索エンジン名を取得 = 設定ファイルから指定キーの値を取得(検索エンジン設定キー)
End Function

' =========================================================================
' 処理概要       : 保存済みの独自検索URLを取得
' 戻り値         : String - 保存済みの独自検索URL
' =========================================================================
Public Function 保存済みの独自検索URLを取得() As String
    保存済みの独自検索URLを取得 = 独自検索URLがHTTPまたはHTTPS形式か確認(設定ファイルから指定キーの値を取得(検索URL設定キー))
End Function

' =========================================================================
' 処理概要       : 設定ファイルから指定キーの値を取得
' 引数           : キー名称 (String) - 設定ファイルのキー名
' 戻り値         : String - 設定ファイルから指定キーの値
' =========================================================================
Private Function 設定ファイルから指定キーの値を取得(ByVal キー名称 As String) As String
    Dim パス As String
    Dim ファイル番号 As Integer
    Dim 読込行 As String, 等号位置 As Long, 参照キー As String, 処理値 As String
    
    設定ファイルから指定キーの値を取得 = ""
    
    パス = 検索設定ファイルのパスを取得()
    If Len(パス) = 0 Then Exit Function
    
    On Error GoTo ErrorHanler
    If Len(Dir$(パス)) = 0 Then Exit Function
    
    ファイル番号 = FreeFile
    
    ' 保存済みファイルから指定した設定項目を読み込む
    Open パス For Input As #ファイル番号
    
    Do Until EOF(ファイル番号)
        Line Input #ファイル番号, 読込行
        等号位置 = InStr(読込行, "=")
        If 等号位置 > 0 Then
            参照キー = LCase$(Trim$(Left$(読込行, 等号位置 - 1)))
            処理値 = Trim$(Mid$(読込行, 等号位置 + 1))
            If 参照キー = LCase$(キー名称) Then
                設定ファイルから指定キーの値を取得 = 処理値
                
            End If
        End If
    Loop
    
    Close #ファイル番号
    Exit Function
    
ErrorHanler:
    MsgBox "処理を完了できません。" & vbCrLf & Err.Description, vbExclamation, "検索設定の読込"
    
    On Error Resume Next
    Close #ファイル番号
    On Error GoTo 0
    設定ファイルから指定キーの値を取得 = ""
    GoTo EndProc
    
EndProc:
End Function

' =========================================================================
' 処理概要       : 保存済み設定を検索URLテンプレートへ変換
' 戻り値         : String - 保存済み設定を検索URLテンプレート
' =========================================================================
Public Function 保存済み設定を検索URLテンプレートへ変換() As String
    Dim 名称 As String
    
    名称 = 保存済みの検索エンジン名を取得()
    If Len(名称) = 0 Then
        保存済み設定を検索URLテンプレートへ変換 = ""
        Exit Function
    End If
    
    If LCase$(Trim$(名称)) = "custom" Then
        保存済み設定を検索URLテンプレートへ変換 = 保存済みの独自検索URLを取得()
        Exit Function
    End If
    
    保存済み設定を検索URLテンプレートへ変換 = 検索エンジン名をURLテンプレートへ変換(名称)
End Function

' =========================================================================
' 処理概要       : 検索エンジン名を設定ファイルへ保存
' 引数           : 検索エンジン名称 (String) - 検索エンジンのプリセット名
' =========================================================================
Private Sub 検索エンジン名を設定ファイルへ保存(ByVal 検索エンジン名称 As String)
    検索エンジン名と独自検索URLを保存 検索エンジン名称, ""
End Sub

' =========================================================================
' 処理概要       : 検索エンジン名と独自検索URLを保存
' 引数           : 検索エンジン名称 (String) - 検索エンジンのプリセット名
'                  テンプレート (String) - 検索語を末尾に連結するURLテンプレート
' =========================================================================
Private Sub 検索エンジン名と独自検索URLを保存(ByVal 検索エンジン名称 As String, ByVal テンプレート As String)
    Dim パス As String
    Dim ファイル番号 As Integer
    
    パス = 検索設定ファイルのパスを取得()
    If Len(パス) = 0 Then Exit Sub
    
    必要な保存先フォルダーを作成
    
    On Error GoTo ErrorHanler
    ファイル番号 = FreeFile
    
    ' 検索エンジンと独自URLを設定ファイルへ書き込む
    Open パス For Output As #ファイル番号
    Print #ファイル番号, 検索エンジン設定キー & "=" & 検索エンジン名称
    If Len(テンプレート) > 0 Then
        Print #ファイル番号, 検索URL設定キー & "=" & テンプレート
    End If
    Close #ファイル番号
    
    Exit Sub
    
ErrorHanler:
    MsgBox "処理を完了できません。" & vbCrLf & Err.Description, vbExclamation, "検索設定の保存"
    On Error Resume Next
    Close #ファイル番号
    On Error GoTo 0
    
    GoTo EndProc
    
EndProc:
End Sub
