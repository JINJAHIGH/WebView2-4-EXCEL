Attribute VB_Name = "mod_WV2_地図検索"
Option Explicit

' -------------
' 定数
' -------------
Private Const Escキー番号 As Long = &H1B&

' -------------
' 画面要素の検索条件
' -------------
Private Const 地図候補セレクター       As String = "a[href*='/maps/place/']" & vbLf & "div[role='feed'] a[href*='/place/']" & vbLf & "a.hfpxzc"
Private Const 地図検索欄セレクター     As String = "input[name='q']" & vbLf & "#searchboxinput" & vbLf & "form input[type='text']"
Private Const 地図検索ボタンセレクター As String = "button[aria-label='検索']" & vbLf & "#searchbox-searchbutton" & vbLf & "button[aria-label='Search']"

Private 保持最終エラー As String

Private 保持最終候補採用済み As Boolean

Private Declare PtrSafe Function キー状態取得API Lib "user32" Alias "GetAsyncKeyState" (ByVal 値キー As Long) As Integer

Private 保持中断要求 As Boolean

Private 保持中断済み As Boolean

Private 保持内部一括処理 As Boolean

Private 保持保存済み中断要求キー As Long

' =========================================================================
' 処理概要       : Googleマップを開いて検索可能な状態にする
' 引数           : 入力データ (mod_WV2_ブラウザー管理) - 操作対象のブラウザー
'                  待機上限秒 (Single) - 待機の上限秒数（省略時: 30）
' 戻り値         : mod_WV2_ブラウザー表示 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Public Function Googleマップを開いて検索可能な状態にする(ByVal 入力データ As mod_WV2_ブラウザー管理, Optional ByVal 待機上限秒 As Single = 30) As mod_WV2_ブラウザー表示
    Dim 対象参照 As mod_WV2_ブラウザー表示, 単独実行 As Boolean
    
    On Error GoTo ErrorHanler
    単独実行 = Not 保持内部一括処理
    
    ' 単独実行の場合は、中断操作を受け付ける準備を行う
    If 単独実行 Then
        地図処理の中断状態を初期化
        Escによる中断を捕捉する設定へ切り替え
    End If
    Set 対象参照 = 地図タブの生成と検索画面の待機を実行(入力データ, 待機上限秒)
    GoTo EndProc
    
ErrorHanler:
    保持中断要求 = True
    保持中断済み = True
    If Err.Number = 18 Then
        保持最終エラー = "canceled"
    Else
        保持最終エラー = Err.Description
        MsgBox "地図を開けません。" & vbCrLf & 保持最終エラー, vbExclamation, "地図検索"
    End If
    Set 対象参照 = Nothing
    
    ' 実行結果を保持し、中断設定の復元へ進む
    Resume EndProc
    
EndProc:
    If 単独実行 Then Escによる中断の設定を元に戻す
    Set Googleマップを開いて検索可能な状態にする = 対象参照
End Function

' =========================================================================
' 処理概要       : 地図タブの生成と検索画面の待機を実行
' 引数           : 入力データ (mod_WV2_ブラウザー管理) - 操作対象のブラウザー
'                  待機上限秒 (Single) - 待機の上限秒数
' 戻り値         : mod_WV2_ブラウザー表示 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Private Function 地図タブの生成と検索画面の待機を実行(ByVal 入力データ As mod_WV2_ブラウザー管理, ByVal 待機上限秒 As Single) As mod_WV2_ブラウザー表示
    Dim 対象参照 As mod_WV2_ブラウザー表示
    Dim 対象要素 As mod_WV2_HTML要素
    Dim 待機開始時刻 As Single
    
    保持最終エラー = ""
    
    If 入力データ Is Nothing Then
        保持最終エラー = "no-browser"
        Exit Function
    End If
    
    Set 対象参照 = 入力データ.新しいタブを作成して指定URLを表示("https://www.google.com/maps")
    
    If 対象参照 Is Nothing Then
        保持最終エラー = "no-pane"
        Exit Function
    End If
    
    待機開始時刻 = Timer
    
    Do
    
        対象参照.JavaScript式を評価して完了まで待機 "1", 3
        If 対象参照.直前の同期JavaScript評価が成功したか取得 Then Exit Do
        If 中断要求とEscキーを確認() Then
            保持最終エラー = "canceled"
            保持中断済み = True
            Exit Function
        End If
        If 日付の切り替わりを考慮して経過秒数を取得(待機開始時刻) >= 待機上限秒 Then
            保持最終エラー = "pane-not-ready"
            Exit Function
        End If
        キャンセルを確認しながらイベント処理を進める 0.3
    Loop
    
    Set 対象要素 = 複数のセレクターから最初に見つかった要素を取得(対象参照, 地図検索欄セレクター, 待機上限秒)
    
    If 対象要素 Is Nothing Then
        保持最終エラー = "no-searchbox"
        Exit Function
    End If
    
    待機開始時刻 = Timer
    
    Do
        If InStr(1, 対象参照.表示中のページのURLを取得(), "/@") > 0 Then Exit Do
        If 中断要求とEscキーを確認() Then Exit Do
        If 日付の切り替わりを考慮して経過秒数を取得(待機開始時刻) >= 8 Then
            mod_WV2_ログ管理.警告レベルのログを記録 "地図検索.MapsOpen: URL に座標が入らない (" & 対象参照.表示中のページのURLを取得() & ")"
            Exit Do
        End If
        キャンセルを確認しながらイベント処理を進める 0.3
    Loop
    
    対象参照.安定判定から除外する通信URLパターンを追加 "/search?tbm=map"
    
    mod_WV2_ログ管理.情報レベルのログを記録 "地図検索.MapsOpen: 準備できた (" & 対象参照.表示中のページのURLを取得() & ")"
    Set 地図タブの生成と検索画面の待機を実行 = 対象参照
End Function

' =========================================================================
' 処理概要       : 住所を検索して緯度と経度を取得
' 引数           : 対象参照 (mod_WV2_ブラウザー表示) - 対象のブラウザーペイン
'                  住所文字列 (String) - 検索する住所
'                  出力緯度 (Double) - ［出力］取得した緯度の出力先
'                  出力経度 (Double) - ［出力］取得した経度の出力先
'                  出力名称 (String) - ［出力］取得した場所名の出力先
'                  待機上限秒 (Single) - 待機の上限秒数（省略時: 20）
'                  複数候補時に先頭を採用 (Boolean) - 複数候補時に先頭候補を採用する場合 True（省略時: False）
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 住所を検索して緯度と経度を取得(ByVal 対象参照 As mod_WV2_ブラウザー表示, ByVal 住所文字列 As String, ByRef 出力緯度 As Double, ByRef 出力経度 As Double, ByRef 出力名称 As String, Optional ByVal 待機上限秒 As Single = 20, Optional ByVal 複数候補時に先頭を採用 As Boolean _
    = False) As Boolean
    Dim 行番号 As Boolean, 単独実行 As Boolean
    
    On Error GoTo ErrorHanler
    単独実行 = Not 保持内部一括処理
    
    ' 単独実行の場合は、中断操作を受け付ける準備を行う
    If 単独実行 Then
        地図処理の中断状態を初期化
        Escによる中断を捕捉する設定へ切り替え
    End If
    行番号 = 住所検索と座標抽出を実行(対象参照, 住所文字列, 出力緯度, 出力経度, 出力名称, 待機上限秒, 複数候補時に先頭を採用)
    GoTo EndProc
    
ErrorHanler:
    保持中断要求 = True
    保持中断済み = True
    If Err.Number = 18 Then
        保持最終エラー = "canceled"
    Else
        保持最終エラー = Err.Description
        MsgBox "住所を検索できません。" & vbCrLf & 保持最終エラー, vbExclamation, "地図検索"
    End If
    行番号 = False
    
    ' 実行結果を保持し、中断設定の復元へ進む
    Resume EndProc
    
EndProc:
    If 単独実行 Then Escによる中断の設定を元に戻す
    住所を検索して緯度と経度を取得 = 行番号
End Function

' =========================================================================
' 処理概要       : 住所検索と座標抽出を実行
' 引数           : 対象参照 (mod_WV2_ブラウザー表示) - 対象のブラウザーペイン
'                  住所文字列 (String) - 検索する住所
'                  出力緯度 (Double) - ［出力］取得した緯度の出力先
'                  出力経度 (Double) - ［出力］取得した経度の出力先
'                  出力名称 (String) - ［出力］取得した場所名の出力先
'                  待機上限秒 (Single) - 待機の上限秒数
'                  複数候補時に先頭を採用 (Boolean) - 複数候補時に先頭候補を採用する場合 True
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Private Function 住所検索と座標抽出を実行(ByVal 対象参照 As mod_WV2_ブラウザー表示, ByVal 住所文字列 As String, ByRef 出力緯度 As Double, ByRef 出力経度 As Double, ByRef 出力名称 As String, ByVal 待機上限秒 As Single, ByVal 複数候補時に先頭を採用 As Boolean) As Boolean
    Dim 対象要素 As mod_WV2_HTML要素
    Dim 前回URL As String
    Dim 新規URL As String
    Dim 候補一覧 As Collection
    
    保持最終エラー = ""
    保持最終候補採用済み = False
    出力緯度 = 0
    出力経度 = 0
    出力名称 = ""
    
    If 対象参照 Is Nothing Then
        保持最終エラー = "no-pane"
        Exit Function
    End If
    
    前回URL = 対象参照.表示中のページのURLを取得()
    
    Set 対象要素 = 複数のセレクターから最初に見つかった要素を取得(対象参照, 地図検索欄セレクター, 10)
    If 中断要求とEscキーを確認() Then GoTo Canceled
    If 対象要素 Is Nothing Then
        保持最終エラー = "no-searchbox"
        Exit Function
    End If
    
    対象要素.入力値 = 住所文字列
    If Not 対象要素.直前操作成功 Then
        保持最終エラー = "write-failed: " & 対象要素.最終エラー
        Exit Function
    End If
    
    Set 対象要素 = 複数のセレクターから最初に見つかった要素を取得(対象参照, 地図検索ボタンセレクター, 3)
    If 対象要素 Is Nothing Then
    
        対象参照.JavaScript式を評価して完了まで待機 "(function(){var e=document.querySelector('input[name=""q""]');" & "if(!e){return 0;}e.focus();" & "e.dispatchEvent(new KeyboardEvent('keydown'," & _
            "{key:'Enter',code:'Enter',keyCode:13,which:13,bubbles:true}));" & "return 1;})()"
        If Not 対象参照.直前の同期JavaScript評価が成功したか取得 Then
            保持最終エラー = "no-search-trigger"
            Exit Function
        End If
    Else
        If Not 対象要素.DOM要素をクリック() Then
            保持最終エラー = "click-failed: " & 対象要素.最終エラー
            Exit Function
        End If
    End If
    
    If Not 検索前と異なる場所のURLへ移動するまで待機(対象参照, 前回URL, 待機上限秒, 新規URL) Then
    
        If 中断要求とEscキーを確認() Then GoTo Canceled
        
        If InStr(1, 新規URL, "/@") = 0 Then
            保持最終エラー = "timeout-url"
            Exit Function
        End If
        
        Set 候補一覧 = 地図の検索結果から候補のリンクを取得(対象参照)
        mod_WV2_ログ管理.情報レベルのログを記録 "地図検索.MapsGeocode: 候補リンク " & 候補一覧.Count & " 件"
        
        If 候補一覧.Count = 0 Then
            保持最終エラー = "not-found"
            mod_WV2_ログ管理.警告レベルのログを記録 "地図検索.MapsGeocode: 見つからない (" & 住所文字列 & ") 候補一覧も出ていないので座標は返さない"
            Exit Function
        End If
        
        If 複数候補時に先頭を採用 Then
            If 先頭の検索候補を選択して場所のURLを取得(対象参照, 候補一覧, 待機上限秒, 新規URL) Then
                保持最終候補採用済み = True
                保持最終エラー = ""
                mod_WV2_ログ管理.情報レベルのログを記録 "地図検索.MapsGeocode: 候補の 1 件目を採った (" & 住所文字列 & ")"
                GoTo Resolved
            End If
            mod_WV2_ログ管理.警告レベルのログを記録 "地図検索.MapsGeocode: 候補を選べなかった (" & 住所文字列 & ")"
            If 中断要求とEscキーを確認() Then GoTo Canceled
        End If
        
        地図URLから緯度と経度を抽出 新規URL, 出力緯度, 出力経度
        出力名称 = 地図のページタイトルから場所名を取得(対象参照)
        保持最終エラー = "ambiguous"
        mod_WV2_ログ管理.警告レベルのログを記録 "地図検索.MapsGeocode: 候補が複数 (" & 住所文字列 & ") 中心座標を返す"
        Exit Function
    End If
    
Resolved:

    対象参照.SPAの通信とDOM更新が落ち着くまで待機 10
    
    出力名称 = 地図のページタイトルから場所名を取得(対象参照)
    
    If Not 地図URLから緯度と経度を抽出(新規URL, 出力緯度, 出力経度) Then
        保持最終エラー = "no-coords"
        Exit Function
    End If
    
    住所検索と座標抽出を実行 = True
    Exit Function
    
Canceled:
    出力緯度 = 0
    出力経度 = 0
    出力名称 = ""
    保持最終エラー = "canceled"
    保持中断済み = True
    mod_WV2_ログ管理.警告レベルのログを記録 "地図検索.MapsGeocodeCore: 中断された (" & 住所文字列 & ")"
End Function

' =========================================================================
' 処理概要       : 地図のページタイトルから場所名を取得
' 引数           : 対象参照 (mod_WV2_ブラウザー表示) - 対象のブラウザーペイン
' 戻り値         : String - 地図のページタイトルから場所名
' =========================================================================
Private Function 地図のページタイトルから場所名を取得(ByVal 対象参照 As mod_WV2_ブラウザー表示) As String
    Dim 開始時刻 As String
    
    開始時刻 = 対象参照.表示中のページのタイトルを取得()
    If InStr(1, 開始時刻, " - ") > 0 Then
        開始時刻 = Left$(開始時刻, InStrRev(開始時刻, " - ") - 1)
    End If
    
    地図のページタイトルから場所名を取得 = 開始時刻
End Function

' =========================================================================
' 処理概要       : 直前の検索で先頭候補を採用したか取得
' 戻り値         : Boolean - 先頭の検索候補を採用した場合 True
' =========================================================================
Public Property Get 直前の検索で先頭候補を採用したか取得() As Boolean
    直前の検索で先頭候補を採用したか取得 = 保持最終候補採用済み
End Property

' =========================================================================
' 処理概要       : 地図の検索結果から候補のリンクを取得
' 引数           : 対象参照 (mod_WV2_ブラウザー表示) - 対象のブラウザーペイン
' 戻り値         : Collection - 取得した要素の一覧
' =========================================================================
Private Function 地図の検索結果から候補のリンクを取得(ByVal 対象参照 As mod_WV2_ブラウザー表示) As Collection
    Dim セレクター一覧 As Variant
    Dim 添字 As Long
    Dim 要素一覧 As Collection
    
    Set 地図の検索結果から候補のリンクを取得 = New Collection
    セレクター一覧 = Split(地図候補セレクター, vbLf)
    
    For 添字 = LBound(セレクター一覧) To UBound(セレクター一覧)
        Set 要素一覧 = 対象参照.CSSセレクターに一致するDOM要素を上限件数まで取得(CStr(セレクター一覧(添字)), 20)
        If 要素一覧.Count > 0 Then
        
            Set 地図の検索結果から候補のリンクを取得 = 要素一覧
            Exit Function
        End If
    Next 添字
End Function

' =========================================================================
' 処理概要       : 先頭の検索候補を選択して場所のURLを取得
' 引数           : 対象参照 (mod_WV2_ブラウザー表示) - 対象のブラウザーペイン
'                  候補一覧 (Collection) - 検索候補の要素コレクション
'                  待機上限秒 (Single) - 待機の上限秒数
'                  出力URL (String) - ［出力］取得した場所URLの出力先
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Private Function 先頭の検索候補を選択して場所のURLを取得(ByVal 対象参照 As mod_WV2_ブラウザー表示, ByVal 候補一覧 As Collection, ByVal 待機上限秒 As Single, ByRef 出力URL As String) As Boolean
    Dim 対象要素 As mod_WV2_HTML要素
    Dim 前回URL As String
    
    If 候補一覧 Is Nothing Then Exit Function
    If 候補一覧.Count = 0 Then Exit Function
    
    前回URL = 対象参照.表示中のページのURLを取得()
    Set 対象要素 = 候補一覧(1)
    
    If Not 対象要素.DOM要素をクリック() Then
        mod_WV2_ログ管理.警告レベルのログを記録 "地図検索.MapsPickFirst: クリックできない err=" & 対象要素.最終エラー
        Exit Function
    End If
    
    先頭の検索候補を選択して場所のURLを取得 = 検索前と異なる場所のURLへ移動するまで待機(対象参照, 前回URL, 待機上限秒, 出力URL)
End Function

' =========================================================================
' 処理概要       : 直前の地図処理のエラー内容を取得
' 戻り値         : String - 直前の地図処理のエラー内容
' =========================================================================
Public Property Get 直前の地図処理のエラー内容を取得() As String
    直前の地図処理のエラー内容を取得 = 保持最終エラー
End Property

' =========================================================================
' 処理概要       : 検索前と異なる場所のURLへ移動するまで待機
' 引数           : 対象参照 (mod_WV2_ブラウザー表示) - 対象のブラウザーペイン
'                  前回URL (String) - 検索前のURL
'                  待機上限秒 (Single) - 待機の上限秒数
'                  出力URL (String) - ［出力］取得した場所URLの出力先
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Private Function 検索前と異なる場所のURLへ移動するまで待機(ByVal 対象参照 As mod_WV2_ブラウザー表示, ByVal 前回URL As String, ByVal 待機上限秒 As Single, ByRef 出力URL As String) As Boolean
    Dim 待機開始時刻 As Single
    Dim 対象URL As String
    
    待機開始時刻 = Timer
    Do
        対象URL = 対象参照.表示中のページのURLを取得()
        
        If 中断要求とEscキーを確認() Then
            出力URL = 対象URL
            Exit Function
        End If
        
        If 対象URL <> 前回URL And 地図URLが場所の詳細ページか判定(対象URL) Then
            出力URL = 対象URL
            検索前と異なる場所のURLへ移動するまで待機 = True
            Exit Function
        End If
        
        If 日付の切り替わりを考慮して経過秒数を取得(待機開始時刻) >= 待機上限秒 Then
            出力URL = 対象URL
            
            検索前と異なる場所のURLへ移動するまで待機 = 地図URLが場所の詳細ページか判定(対象URL)
            Exit Function
        End If
        
        キャンセルを確認しながらイベント処理を進める 0.2
    Loop
End Function

' =========================================================================
' 処理概要       : 地図URLが場所の詳細ページか判定
' 引数           : 対象URL (String) - 判定する地図URL
' 戻り値         : Boolean - 場所の詳細URLの場合 True
' =========================================================================
Private Function 地図URLが場所の詳細ページか判定(ByVal 対象URL As String) As Boolean
    If InStr(1, 対象URL, "/place/") = 0 Then Exit Function
    If InStr(1, 対象URL, "!3d") = 0 Then Exit Function
    
    地図URLが場所の詳細ページか判定 = True
End Function

' =========================================================================
' 処理概要       : 地図URLから緯度と経度を抽出
' 引数           : URL (String) - 移動先または通知対象のURL
'                  出力緯度 (Double) - ［出力］取得した緯度の出力先
'                  出力経度 (Double) - ［出力］取得した経度の出力先
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Private Function 地図URLから緯度と経度を抽出(ByVal URL As String, ByRef 出力緯度 As Double, ByRef 出力経度 As Double) As Boolean
    Dim 添字 As Long
    Dim 副添字 As Long
    
    添字 = InStr(1, URL, "!3d")
    副添字 = InStr(1, URL, "!4d")
    If 添字 > 0 And 副添字 > 0 Then
        出力緯度 = Val(Mid$(URL, 添字 + 3))
        出力経度 = Val(Mid$(URL, 副添字 + 3))
        If 緯度と経度が有効範囲にあるか判定(出力緯度, 出力経度) Then
            地図URLから緯度と経度を抽出 = True
            Exit Function
        End If
    End If
    
    添字 = InStr(1, URL, "/@")
    If 添字 > 0 Then
        出力緯度 = Val(Mid$(URL, 添字 + 2))
        副添字 = InStr(添字, URL, ",")
        If 副添字 > 0 Then 出力経度 = Val(Mid$(URL, 副添字 + 1))
        If 緯度と経度が有効範囲にあるか判定(出力緯度, 出力経度) Then
            地図URLから緯度と経度を抽出 = True
            Exit Function
        End If
    End If
    
    出力緯度 = 0
    出力経度 = 0
End Function

' =========================================================================
' 処理概要       : 緯度と経度が有効範囲にあるか判定
' 引数           : 緯度値 (Double) - 判定する緯度
'                  経度値 (Double) - 判定する経度
' 戻り値         : Boolean - 緯度と経度が有効範囲内の場合 True
' =========================================================================
Private Function 緯度と経度が有効範囲にあるか判定(ByVal 緯度値 As Double, ByVal 経度値 As Double) As Boolean
    If 緯度値 = 0 And 経度値 = 0 Then Exit Function
    If 緯度値 < -90 Or 緯度値 > 90 Then Exit Function
    If 経度値 < -180 Or 経度値 > 180 Then Exit Function
    
    緯度と経度が有効範囲にあるか判定 = True
End Function

' =========================================================================
' 処理概要       : 複数のセレクターから最初に見つかった要素を取得
' 引数           : 対象参照 (mod_WV2_ブラウザー表示) - 対象のブラウザーペイン
'                  セレクター一覧 (String) - 改行区切りのCSSセレクター一覧
'                  待機上限秒 (Single) - 待機の上限秒数
' 戻り値         : mod_WV2_HTML要素 - 対象オブジェクト。取得できない場合 Nothing
' =========================================================================
Private Function 複数のセレクターから最初に見つかった要素を取得(ByVal 対象参照 As mod_WV2_ブラウザー表示, ByVal セレクター一覧 As String, ByVal 待機上限秒 As Single) As mod_WV2_HTML要素
    Dim 候補一覧 As Variant
    Dim 添字 As Long
    Dim 対象要素 As mod_WV2_HTML要素
    
    候補一覧 = Split(セレクター一覧, vbLf)
    
    For 添字 = LBound(候補一覧) To UBound(候補一覧)
        If 添字 = LBound(候補一覧) Then
            Set 対象要素 = 対象参照.指定要素が見つかるまで待機(CStr(候補一覧(添字)), 待機上限秒)
        Else
            Set 対象要素 = 対象参照.CSSセレクターで最初のDOM要素を検索(CStr(候補一覧(添字)))
        End If
        
        If Not 対象要素 Is Nothing Then
            Set 複数のセレクターから最初に見つかった要素を取得 = 対象要素
            Exit Function
        End If
    Next 添字
    
    mod_WV2_ログ管理.警告レベルのログを記録 "地図検索.MapsFindFirst: どの候補も外れた [" & Replace$(セレクター一覧, vbLf, " / ") & "]"
End Function

' =========================================================================
' 処理概要       : キャンセルを確認しながらイベント処理を進める
' 引数           : 待機秒 (Single) - 待機時間（秒）
' =========================================================================
Private Sub キャンセルを確認しながらイベント処理を進める(ByVal 待機秒 As Single)
    Dim 待機開始時刻 As Single
    
    待機開始時刻 = Timer
    Do
        DoEvents
        
        If 中断要求とEscキーを確認() Then Exit Do
        If 日付の切り替わりを考慮して経過秒数を取得(待機開始時刻) >= 待機秒 Then Exit Do
    Loop
End Sub

' =========================================================================
' 処理概要       : 日付の切り替わりを考慮して経過秒数を取得
' 引数           : 開始タイマー (Single) - 計測を開始したTimer値
' 戻り値         : Single - 日付の切り替わりを考慮して経過秒数
' =========================================================================
Private Function 日付の切り替わりを考慮して経過秒数を取得(ByVal 開始タイマー As Single) As Single
    Dim 後続文字 As Single
    
    後続文字 = Timer - 開始タイマー
    If 後続文字 < 0 Then 後続文字 = 後続文字 + 86400
    日付の切り替わりを考慮して経過秒数を取得 = 後続文字
End Function

' =========================================================================
' 処理概要       : 地図処理の中断要求を取得
' 戻り値         : Boolean - 中断が要求されている場合 True
' =========================================================================
Public Property Get 地図処理中断要求() As Boolean
    地図処理中断要求 = 保持中断要求
End Property

' =========================================================================
' 処理概要       : 地図処理の中断要求を設定
' 引数           : 新規値 (Boolean) - 設定する値
' =========================================================================
Public Property Let 地図処理中断要求(ByVal 新規値 As Boolean)
    保持中断要求 = 新規値
End Property

' =========================================================================
' 処理概要       : 地図処理が中断されたか取得
' 戻り値         : Boolean - 中断済みの場合 True
' =========================================================================
Public Property Get 地図処理中断済み() As Boolean
    地図処理中断済み = 保持中断済み
End Property

' =========================================================================
' 処理概要       : 指定シートの住所列にある処理対象行数を取得
' 引数           : 対象シート (Object) - 住所と検索結果を読み書きするシート
'                  先頭行 (Long) - 処理を開始する行番号（省略時: 2）
'                  アドレス列 (Long) - 住所を読み取る列番号（省略時: 1）
' 戻り値         : Long - 指定シートの住所列にある処理対象行数
' =========================================================================
Public Function 指定シートの住所列にある処理対象行数を取得(ByVal 対象シート As Object, Optional ByVal 先頭行 As Long = 2, Optional ByVal アドレス列 As Long = 1) As Long
    Dim 行番号 As Long
    Dim 件数 As Long
    
    If 対象シート Is Nothing Then Exit Function
    
    行番号 = 先頭行
    Do
        If Len(Trim$(CStr(対象シート.Cells(行番号, アドレス列).value))) = 0 Then Exit Do
        件数 = 件数 + 1
        行番号 = 行番号 + 1
    Loop
    
    指定シートの住所列にある処理対象行数を取得 = 件数
End Function

' =========================================================================
' 処理概要       : 中断要求とEscキーを確認
' 戻り値         : Boolean - 中断要求を検出した場合 True
' =========================================================================
Private Function 中断要求とEscキーを確認() As Boolean
    Dim 参照キー As Integer
    
    If 保持中断要求 Then
        中断要求とEscキーを確認 = True
        Exit Function
    End If
    
    参照キー = キー状態取得API(Escキー番号)
    If 参照キー <> 0 Then
        保持中断要求 = True
        中断要求とEscキーを確認 = True
        mod_WV2_ログ管理.情報レベルのログを記録 "地図検索: Esc で中断が要求された"
    End If
End Function

' =========================================================================
' 処理概要       : 地図処理の中断状態を初期化
' =========================================================================
Private Sub 地図処理の中断状態を初期化()
    Dim 参照キー As Integer
    
    保持中断要求 = False
    保持中断済み = False
    参照キー = キー状態取得API(Escキー番号)
End Sub

' =========================================================================
' 処理概要       : Escによる中断を捕捉する設定へ切り替え
' =========================================================================
Private Sub Escによる中断を捕捉する設定へ切り替え()
    保持保存済み中断要求キー = Application.EnableCancelKey
    Application.EnableCancelKey = xlDisabled
End Sub

' =========================================================================
' 処理概要       : Escによる中断の設定を元に戻す
' =========================================================================
Private Sub Escによる中断の設定を元に戻す()
    If 保持保存済み中断要求キー = 0 Then 保持保存済み中断要求キー = xlInterrupt
    Application.EnableCancelKey = 保持保存済み中断要求キー
    保持保存済み中断要求キー = 0
End Sub

' =========================================================================
' 処理概要       : 地図処理の進捗をステータスバーへ表示
' 引数           : 完了行数 (Long) - 処理済み行数
'                  総数行数 (Long) - 処理対象の総行数
'                  成功行数 (Long) - 成功した行数
'                  失敗行数 (Long) - 失敗した行数
' =========================================================================
Private Sub 地図処理の進捗をステータスバーへ表示(ByVal 完了行数 As Long, ByVal 総数行数 As Long, ByVal 成功行数 As Long, ByVal 失敗行数 As Long)
    Application.StatusBar = "住所→座標 " & 完了行数 & "/" & 総数行数 & " (ok " & 成功行数 & " / 失敗 " & 失敗行数 & ")" & "  ―  中止は Esc"
End Sub

' =========================================================================
' 処理概要       : 指定シートの住所を順に検索して座標を書き込み
' 引数           : 対象シート (Object) - 住所と検索結果を読み書きするシート
'                  先頭行 (Long) - 処理を開始する行番号（省略時: 2）
'                  アドレス列 (Long) - 住所を読み取る列番号（省略時: 1）
'                  出力列 (Long) - 結果を書き込む先頭列番号（省略時: 2）
'                  スキップ完了 (Boolean) - 結果記入済みの行を処理しない場合 True（省略時: True）
'                  採用先頭 (Boolean) - 複数候補時に先頭候補を採用する場合 True（省略時: False）
'                  ブラウザー (mod_WV2_ブラウザー管理) - 操作対象のブラウザー（省略時: Nothing）
' 戻り値         : Long - 座標の取得に成功した行数
' =========================================================================
Public Function 指定シートの住所を順に検索して座標を書き込み(ByVal 対象シート As Object, Optional ByVal 先頭行 As Long = 2, Optional ByVal アドレス列 As Long = 1, Optional ByVal 出力列 As Long = 2, Optional ByVal スキップ完了 As Boolean = True, Optional ByVal _
    採用先頭 As Boolean = False, Optional ByVal ブラウザー As mod_WV2_ブラウザー管理 = Nothing) As Long
    Dim 対象参照 As mod_WV2_ブラウザー表示
    Dim 行番号 As Long
    Dim アドレス As String
    Dim 緯度 As Double
    Dim 経度 As Double
    Dim 場所名 As String
    Dim 成功 As Boolean
    Dim 完了件数 As Long
    Dim 総数 As Long
    Dim 総数行数 As Long
    
    保持最終エラー = ""
    
    If 対象シート Is Nothing Then
        保持最終エラー = "no-sheet"
        Exit Function
    End If
    
    地図処理の中断状態を初期化
    
    総数行数 = 指定シートの住所列にある処理対象行数を取得(対象シート, 先頭行, アドレス列)
    If 総数行数 = 0 Then
        保持最終エラー = "no-rows"
        mod_WV2_ログ管理.警告レベルのログを記録 "地図検索.MapsGeocodeSheet: 処理する行が無い"
        Exit Function
    End If
    mod_WV2_ログ管理.情報レベルのログを記録 "地図検索.MapsGeocodeSheet: 全 " & 総数行数 & " 件 (中止は Esc)"
    
    保持内部一括処理 = True
    Escによる中断を捕捉する設定へ切り替え
    On Error GoTo ErrorHanler
    
    Set 対象参照 = Googleマップを開いて検索可能な状態にする(ブラウザー)
    If 対象参照 Is Nothing Then
        mod_WV2_ログ管理.エラーレベルのログを記録 "地図検索.MapsGeocodeSheet: Maps を開けない (" & 保持最終エラー & ")"
        GoTo EndProc
    End If
    
    行番号 = 先頭行
    Do
    
        If 中断要求とEscキーを確認() Then
            保持中断済み = True
            mod_WV2_ログ管理.警告レベルのログを記録 "地図検索.MapsGeocodeSheet: 中断された (" & 行番号 & " 行目の手前まで)"
            Exit Do
        End If
        
        アドレス = Trim$(CStr(対象シート.Cells(行番号, アドレス列).value))
        If Len(アドレス) = 0 Then Exit Do
        
        総数 = 総数 + 1
        地図処理の進捗をステータスバーへ表示 総数, 総数行数, 完了件数, (総数 - 1) - 完了件数
        
        If スキップ完了 And Left$(CStr(対象シート.Cells(行番号, 出力列 + 3).value), 2) = "ok" Then
            mod_WV2_ログ管理.情報レベルのログを記録 "  [" & 行番号 & "] 済みなので飛ばす: " & アドレス
            完了件数 = 完了件数 + 1
        Else
            成功 = 住所を検索して緯度と経度を取得(対象参照, アドレス, 緯度, 経度, 場所名, 20, 採用先頭)
            
            If 保持中断済み Then
                総数 = 総数 - 1
                mod_WV2_ログ管理.警告レベルのログを記録 "地図検索.MapsGeocodeSheet: 中断された (" & 行番号 & " 行目の処理中)"
                Exit Do
            End If
            
            If 成功 Then
                対象シート.Cells(行番号, 出力列).value = 緯度
                対象シート.Cells(行番号, 出力列 + 1).value = 経度
                対象シート.Cells(行番号, 出力列 + 2).value = 場所名
                
                If 直前の検索で先頭候補を採用したか取得 Then
                    対象シート.Cells(行番号, 出力列 + 3).value = "ok(候補1)"
                Else
                    対象シート.Cells(行番号, 出力列 + 3).value = "ok"
                End If
                完了件数 = 完了件数 + 1
            Else
            
                If 緯度 <> 0 Or 経度 <> 0 Then
                    対象シート.Cells(行番号, 出力列).value = 緯度
                    対象シート.Cells(行番号, 出力列 + 1).value = 経度
                End If
                If Len(場所名) > 0 Then 対象シート.Cells(行番号, 出力列 + 2).value = 場所名
                対象シート.Cells(行番号, 出力列 + 3).value = 直前の地図処理のエラー内容を取得
            End If
            
            mod_WV2_ログ管理.情報レベルのログを記録 "  [" & 行番号 & "] " & アドレス & " → " & IIf(成功, 緯度 & "," & 経度, "" & 直前の地図処理のエラー内容を取得 & "")
        End If
        
        DoEvents
        行番号 = 行番号 + 1
    Loop
        GoTo EndProc
        
ErrorHanler:
    保持中断要求 = True
    保持中断済み = True
    If Err.Number = 18 Then
        保持最終エラー = "canceled"
    Else
        保持最終エラー = Err.Description
        MsgBox "シートの住所検索を続行できません。" & vbCrLf & 保持最終エラー, vbExclamation, "地図検索"
    End If
    
    ' 実行結果を保持し、中断設定の復元へ進む
    Resume EndProc
    
EndProc:
    Escによる中断の設定を元に戻す
    保持内部一括処理 = False
    
    Application.StatusBar = Empty
    If 保持中断済み Then 保持最終エラー = "canceled"
    
    mod_WV2_ログ管理.情報レベルのログを記録 "地図検索.MapsGeocodeSheet: " & 総数 & " 行中 " & 完了件数 & " 行が ok" & IIf(保持中断済み, " 中断した", "")
    指定シートの住所を順に検索して座標を書き込み = 完了件数
End Function
