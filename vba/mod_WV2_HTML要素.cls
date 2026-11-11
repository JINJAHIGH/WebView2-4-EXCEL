VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "mod_WV2_HTML要素"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private 保持表示   As mod_WV2_ブラウザー表示
Private 保持ハンドル As Long
Private 保持世代    As String

Private 保持最終成功    As Boolean
Private 保持最終エラー As String

Private 保持最終情報  As String

' =========================================================================
' 処理概要       : インスタンスの内部状態を初期化
' =========================================================================
Private Sub Class_Initialize()
    Set 保持表示 = Nothing
    保持ハンドル = -1
    保持世代 = ""
    保持最終成功 = False
    保持最終エラー = ""
    保持最終情報 = ""
End Sub

' =========================================================================
' 処理概要       : 保持するリソースと参照を解放
' =========================================================================
Private Sub Class_Terminate()
    Set 保持表示 = Nothing
End Sub

' =========================================================================
' 処理概要       : DOM要素の参照情報を設定
' 引数           : 所有元表示 (mod_WV2_ブラウザー表示) - 要素を保持するペイン
'                  要素ハンドル (Long) - DOM要素のレジストリ番号
'                  要素世代 (String) - 要素を取得したページ世代
' =========================================================================
Friend Sub DOM要素の参照情報を設定(ByVal 所有元表示 As mod_WV2_ブラウザー表示, ByVal 要素ハンドル As Long, ByVal 要素世代 As String)
    Set 保持表示 = 所有元表示
    保持ハンドル = 要素ハンドル
    保持世代 = 要素世代
    保持最終成功 = True
    保持最終エラー = ""
    保持最終情報 = ""
End Sub

' =========================================================================
' 処理概要       : DOM要素のレジストリ番号を取得
' 戻り値         : Long - DOM要素のレジストリ番号
' =========================================================================
Public Property Get 要素番号() As Long
    要素番号 = 保持ハンドル
End Property

' =========================================================================
' 処理概要       : DOM要素を取得したページ世代を取得
' 戻り値         : String - DOM要素を取得したページ世代
' =========================================================================
Public Property Get ページ世代() As String
    ページ世代 = 保持世代
End Property

' =========================================================================
' 処理概要       : 直前の要素操作が成功したか取得
' 戻り値         : Boolean - 直前の要素操作が成功した場合 True
' =========================================================================
Public Property Get 直前操作成功() As Boolean
    直前操作成功 = 保持最終成功
End Property

' =========================================================================
' 処理概要       : 直前の処理で発生したエラーを取得
' 戻り値         : String - 直前の処理で発生したエラー
' =========================================================================
Public Property Get 最終エラー() As String
    最終エラー = 保持最終エラー
End Property

' =========================================================================
' 処理概要       : 直前の要素操作の状態情報を取得
' 戻り値         : String - 直前の要素操作の状態情報
' =========================================================================
Public Property Get 直前操作情報() As String
    直前操作情報 = 保持最終情報
End Property

' =========================================================================
' 処理概要       : DOM要素のタグ名を取得
' 戻り値         : String - DOM要素のタグ名
' =========================================================================
Public Property Get DOM要素のタグ名を取得() As String
    DOM要素のタグ名を取得 = DOM要素のプロパティを取得して操作結果を記録("e.tagName")
End Property

' =========================================================================
' 処理概要       : DOM要素の表示テキストを取得
' 戻り値         : String - DOM要素の表示テキスト
' =========================================================================
Public Property Get DOM要素の表示テキストを取得() As String
    DOM要素の表示テキストを取得 = DOM要素のプロパティを取得して操作結果を記録("e.innerText")
End Property

' =========================================================================
' 処理概要       : DOM要素内のHTMLを取得
' 戻り値         : String - DOM要素内のHTML
' =========================================================================
Public Property Get DOM要素内のHTMLを取得() As String
    DOM要素内のHTMLを取得 = DOM要素のプロパティを取得して操作結果を記録("e.innerHTML")
End Property

' =========================================================================
' 処理概要       : DOM要素の入力値を取得
' 戻り値         : String - DOM要素の入力値
' =========================================================================
Public Property Get 入力値() As String
    入力値 = DOM要素のプロパティを取得して操作結果を記録("e.value")
End Property

' =========================================================================
' 処理概要       : DOM要素の指定属性を取得
' 引数           : 属性名称 (String) - DOM属性名
' 戻り値         : String - DOM要素の指定属性
' =========================================================================
Public Function DOM要素の指定属性を取得(ByVal 属性名称 As String) As String
    If 保持表示 Is Nothing Then
        保持最終成功 = False
        保持最終エラー = "no-pane"
        DOM要素の指定属性を取得 = ""
        Exit Function
    End If
    DOM要素の指定属性を取得 = DOM要素のプロパティを取得して操作結果を記録("e.getAttribute(" & 保持表示.JavaScript文字列に変換(属性名称) & ")")
End Function

' =========================================================================
' 処理概要       : DOM要素の入力値を設定
' 引数           : 新規値 (String) - DOM要素へ設定する入力値
' =========================================================================
Public Property Let 入力値(ByVal 新規値 As String)
    Dim JavaScript値 As String
    Dim 実行スクリプト As String
    
    If 保持表示 Is Nothing Then
        保持最終成功 = False
        保持最終エラー = "no-pane"
        保持最終情報 = ""
        Exit Property
    End If
    
    JavaScript値 = 保持表示.JavaScript文字列に変換(新規値)
    
    実行スクリプト = "var d=Object.getOwnPropertyDescriptor(e.constructor.prototype,'value');" & "if(d&&d.set){d.set.call(e," & JavaScript値 & ");rv='setter';}" & "else{e.value=" & JavaScript値 & ";rv='direct';}" & _
        "e.dispatchEvent(new Event('input',{bubbles:true}));" & "e.dispatchEvent(new Event('change',{bubbles:true}));"
        
    DOM要素の操作を実行して結果を記録 実行スクリプト, "Value"
End Property

' =========================================================================
' 処理概要       : DOM要素をクリック
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function DOM要素をクリック() As Boolean
    Dim 実行スクリプト As String
    
    実行スクリプト = "if(typeof e.click==='function'){e.click();rv='click';}" & "else{e.dispatchEvent(new MouseEvent('click'," & "{bubbles:true,cancelable:true,view:window}));rv='dispatch';}"
    
    DOM要素をクリック = DOM要素の操作を実行して結果を記録(実行スクリプト, "Click")
End Function

' =========================================================================
' 処理概要       : DOM要素の属性値を設定
' 引数           : 属性名称 (String) - DOM属性名
'                  属性値 (String) - 設定するDOM属性値
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function DOM要素の属性値を設定(ByVal 属性名称 As String, ByVal 属性値 As String) As Boolean
    Dim 実行スクリプト As String
    
    DOM要素の属性値を設定 = False
    
    If 保持表示 Is Nothing Then
        保持最終成功 = False
        保持最終エラー = "no-pane"
        保持最終情報 = ""
        Exit Function
    End If
    
    実行スクリプト = "e.setAttribute(" & 保持表示.JavaScript文字列に変換(属性名称) & "," & 保持表示.JavaScript文字列に変換(属性値) & ");rv='ok';"
    
    DOM要素の属性値を設定 = DOM要素の操作を実行して結果を記録(実行スクリプト, "SetAttribute")
End Function

' =========================================================================
' 処理概要       : DOM要素の参照が失効しているか判定
' 戻り値         : Boolean - 参照が失効している場合 True
' =========================================================================
Public Property Get DOM要素の参照が失効しているか判定() As Boolean
    Dim 作業情報 As String
    Dim 参照失効 As Boolean
    
    保持最終成功 = False
    保持最終エラー = ""
    保持最終情報 = ""
    
    If 保持表示 Is Nothing Then
        保持最終エラー = "no-pane"
        DOM要素の参照が失効しているか判定 = True
        Exit Property
    End If
    
    参照失効 = 保持表示.DOM要素のページ世代と参照の有効性を確認(保持ハンドル, 保持世代, 作業情報)
    
    If Len(作業情報) = 0 Then
        保持最終成功 = True
    Else
        保持最終エラー = 作業情報
    End If
    
    DOM要素の参照が失効しているか判定 = 参照失効
End Property

' =========================================================================
' 処理概要       : DOM要素のプロパティを取得して操作結果を記録
' 引数           : 値式 (String) - DOM要素から値を読み取るJavaScript式
' 戻り値         : String - DOM要素のプロパティを取得して操作結果を記録
' =========================================================================
Private Function DOM要素のプロパティを取得して操作結果を記録(ByVal 値式 As String) As String
    Dim 処理値 As String
    Dim 作業情報 As String
    Dim 成功 As Boolean
    
    保持最終成功 = False
    保持最終エラー = ""
    保持最終情報 = ""
    DOM要素のプロパティを取得して操作結果を記録 = ""
    
    If 保持表示 Is Nothing Then
        保持最終エラー = "no-pane"
        Exit Function
    End If
    
    成功 = 保持表示.ページ世代を確認してDOM要素の値を取得(保持ハンドル, 保持世代, 値式, 処理値, 作業情報)
    
    If 成功 Then
        保持最終成功 = True
        DOM要素のプロパティを取得して操作結果を記録 = 処理値
    Else
        保持最終エラー = 作業情報
    End If
End Function

' =========================================================================
' 処理概要       : DOM要素の操作を実行して結果を記録
' 引数           : 操作実行文 (String) - 要素に対して実行するJavaScript文
'                  API名称 (String) - 呼び出し元の操作名
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Private Function DOM要素の操作を実行して結果を記録(ByVal 操作実行文 As String, ByVal API名称 As String) As Boolean
    Dim 処理値 As String
    Dim 作業情報 As String
    
    保持最終成功 = False
    保持最終エラー = ""
    保持最終情報 = ""
    DOM要素の操作を実行して結果を記録 = False
    
    If 保持表示 Is Nothing Then
        保持最終エラー = "no-pane"
        Exit Function
    End If
    
    If 保持表示.ページ世代を確認してDOM要素の操作を実行(保持ハンドル, 保持世代, 操作実行文, 処理値, 作業情報) Then
        保持最終成功 = True
        保持最終情報 = 処理値
        DOM要素の操作を実行して結果を記録 = True
        
        If 保持表示.操作後に安定待機 Then 保持表示.SPAの通信とDOM更新が落ち着くまで待機
    Else
        保持最終エラー = 作業情報
        mod_WV2_ログ管理.警告レベルのログを記録 "HTML要素." & API名称 & ": 失敗 (handle=" & 保持ハンドル & ") err=" & 作業情報
    End If
End Function
