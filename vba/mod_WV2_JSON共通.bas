Attribute VB_Name = "mod_WV2_JSON共通"
Option Explicit

' =========================================================================
' 処理概要       : 文字列をJSONに埋め込める形式へ変換
' 引数           : 作成文字列 (String) - 変換対象の文字列
' 戻り値         : String - 文字列をJSONに埋め込める形式
' =========================================================================
Public Function 文字列をJSONに埋め込める形式へ変換(ByVal 作成文字列 As String) As String
    Dim 出力 As String
    Dim 添字 As Long
    Dim 文字コード As Long
    Dim 対象文字 As String
    
    出力 = ""
    For 添字 = 1 To Len(作成文字列)
        対象文字 = Mid$(作成文字列, 添字, 1)
        文字コード = AscW(対象文字)
        
        ' 引用符、バックスラッシュ、制御文字をJSONのエスケープ表現へ置き換える
        Select Case 文字コード
            Case 34
                出力 = 出力 & "\"""
            Case 92
                出力 = 出力 & "\\"
            Case 8
                出力 = 出力 & "\b"
            Case 9
                出力 = 出力 & "\t"
            Case 10
                出力 = 出力 & "\n"
            Case 12
                出力 = 出力 & "\f"
            Case 13
                出力 = 出力 & "\r"
            Case Else
            
                If 文字コード >= 0 And 文字コード < 32 Then
                    出力 = 出力 & "\u" & Right$("0000" & Hex$(文字コード), 4)
                Else
                    出力 = 出力 & 対象文字
                End If
        End Select
    Next 添字
    
    文字列をJSONに埋め込める形式へ変換 = 出力
End Function

' =========================================================================
' 処理概要       : JSONから指定キーの文字列を取得
' 引数           : JSON (String) - 読み取るJSON文字列
'                  キー (String) - JSONから読み取るキー名
' 戻り値         : String - JSONから指定キーの文字列
' =========================================================================
Public Function JSONから指定キーの文字列を取得(ByVal JSON As String, ByVal キー As String) As String
    Dim 検索パターン As String
    Dim 対象参照 As Long
    Dim 作業情報 As Long
    Dim 終端位置 As Long
    
    検索パターン = """" & キー & """:"""
    
    対象参照 = InStr(JSON, 検索パターン)
    If 対象参照 = 0 Then
        JSONから指定キーの文字列を取得 = ""
        Exit Function
    End If
    
    作業情報 = 対象参照 + Len(検索パターン)
    
    終端位置 = InStr(作業情報, JSON, """")
    If 終端位置 = 0 Then
        JSONから指定キーの文字列を取得 = ""
        Exit Function
    End If
    
    JSONから指定キーの文字列を取得 = Mid$(JSON, 作業情報, 終端位置 - 作業情報)
End Function

' =========================================================================
' 処理概要       : JSONから指定キーの数値を取得
' 引数           : JSON (String) - 読み取るJSON文字列
'                  キー (String) - JSONから読み取るキー名
' 戻り値         : Long - 指定キーの非負整数値。取得できない場合 0
' =========================================================================
Public Function JSONから指定キーの数値を取得(ByVal JSON As String, ByVal キー As String) As Long
    Dim 検索パターン As String
    Dim 対象参照 As Long
    Dim 作業情報 As Long
    Dim 作業文字列 As String
    Dim 対象文字 As String
    
    検索パターン = """" & キー & """:"
    
    対象参照 = InStr(JSON, 検索パターン)
    If 対象参照 = 0 Then
        JSONから指定キーの数値を取得 = 0
        Exit Function
    End If
    
    作業情報 = 対象参照 + Len(検索パターン)
    
    Do While 作業情報 <= Len(JSON)
        If Mid$(JSON, 作業情報, 1) <> " " Then Exit Do
        作業情報 = 作業情報 + 1
    Loop
    
    Do While 作業情報 <= Len(JSON)
        対象文字 = Mid$(JSON, 作業情報, 1)
        If 対象文字 < "0" Or 対象文字 > "9" Then Exit Do
        作業文字列 = 作業文字列 & 対象文字
        作業情報 = 作業情報 + 1
    Loop
    
    If Len(作業文字列) = 0 Then
        JSONから指定キーの数値を取得 = 0
    Else
        JSONから指定キーの数値を取得 = CLng(作業文字列)
    End If
End Function

' =========================================================================
' 処理概要       : Boolean値をJSONの真偽値文字列へ変換
' 引数           : 入力データ (Boolean) - 変換する真偽値
' 戻り値         : String - Boolean値をJSONの真偽値文字列
' =========================================================================
Public Function Boolean値をJSONの真偽値文字列へ変換(ByVal 入力データ As Boolean) As String
    If 入力データ Then
        Boolean値をJSONの真偽値文字列へ変換 = "true"
    Else
        Boolean値をJSONの真偽値文字列へ変換 = "false"
    End If
End Function

' =========================================================================
' 処理概要       : 指定位置のJSON文字列を復号
' 引数           : JSON (String) - 読み取るJSON文字列
'                  開始位置 (Long) - JSON文字列内の読み取り開始位置
'                  出力値 (String) - ［出力］復号したJSON文字列の出力先
' 戻り値         : Long - 終端の引用符の位置（1始まり）。復号できない場合 0
' =========================================================================
Public Function 指定位置のJSON文字列を復号(ByVal JSON As String, ByVal 開始位置 As Long, ByRef 出力値 As String) As Long
    Dim 添字 As Long
    Dim 作業文字列 As String
    Dim 対象文字 As String
    Dim 後続文字 As String
    
    出力値 = ""
    指定位置のJSON文字列を復号 = 0
    
    If 開始位置 < 1 Then Exit Function
    
    添字 = 開始位置
    Do While 添字 <= Len(JSON)
        対象文字 = Mid$(JSON, 添字, 1)
        
        If 対象文字 = "\" Then
            後続文字 = Mid$(JSON, 添字 + 1, 1)
            
            ' バックスラッシュに続く文字を確認し、元の文字へ戻す
            Select Case 後続文字
                Case """"
                    作業文字列 = 作業文字列 & """"
                Case "\"
                    作業文字列 = 作業文字列 & "\"
                Case "/"
                    作業文字列 = 作業文字列 & "/"
                Case "b"
                    作業文字列 = 作業文字列 & Chr$(8)
                Case "f"
                    作業文字列 = 作業文字列 & Chr$(12)
                Case "n"
                    作業文字列 = 作業文字列 & vbLf
                Case "r"
                    作業文字列 = 作業文字列 & vbCr
                Case "t"
                    作業文字列 = 作業文字列 & vbTab
                Case "u"
                
                    作業文字列 = 作業文字列 & ChrW$(CLng("&H" & Mid$(JSON, 添字 + 2, 4)))
                    添字 = 添字 + 4
                Case Else
                    作業文字列 = 作業文字列 & 後続文字
            End Select
            添字 = 添字 + 2
            
        ElseIf 対象文字 = """" Then
            出力値 = 作業文字列
            指定位置のJSON文字列を復号 = 添字
            Exit Function
            
        Else
            作業文字列 = 作業文字列 & 対象文字
            添字 = 添字 + 1
        End If
    Loop
End Function

' =========================================================================
' 処理概要       : JSONのエスケープ表現を通常の文字列へ復号
' 引数           : 未加工値 (String) - 解析する入力文字列
' 戻り値         : String - JSONのエスケープ表現を通常の文字列へ復号
' =========================================================================
Public Function JSONのエスケープ表現を通常の文字列へ復号(ByVal 未加工値 As String) As String
    Dim 処理値 As String
    Dim 終端位置 As Long
    
    If Len(未加工値) = 0 Then
        JSONのエスケープ表現を通常の文字列へ復号 = ""
        Exit Function
    End If
    
    If Left$(未加工値, 1) <> """" Then
        JSONのエスケープ表現を通常の文字列へ復号 = 未加工値
        Exit Function
    End If
    
    終端位置 = 指定位置のJSON文字列を復号(未加工値, 2, 処理値)
    If 終端位置 > 0 Then
        JSONのエスケープ表現を通常の文字列へ復号 = 処理値
    Else
        JSONのエスケープ表現を通常の文字列へ復号 = 未加工値
    End If
End Function

' =========================================================================
' 処理概要       : 指定キーの文字列を復号してJSONから取得
' 引数           : JSON (String) - 読み取るJSON文字列
'                  キー (String) - JSONから読み取るキー名
'                  出力値 (String) - ［出力］復号したJSON文字列の出力先
' 戻り値         : Boolean - 処理成功時 True / 失敗時 False
' =========================================================================
Public Function 指定キーの文字列を復号してJSONから取得(ByVal JSON As String, ByVal キー As String, ByRef 出力値 As String) As Boolean
    Dim 検索パターン As String
    Dim 対象参照 As Long
    Dim 終端位置 As Long
    
    出力値 = ""
    指定キーの文字列を復号してJSONから取得 = False
    
    検索パターン = """" & キー & """:"""
    対象参照 = InStr(1, JSON, 検索パターン)
    If 対象参照 = 0 Then Exit Function
    
    終端位置 = 指定位置のJSON文字列を復号(JSON, 対象参照 + Len(検索パターン), 出力値)
    指定キーの文字列を復号してJSONから取得 = (終端位置 > 0)
    If Not 指定キーの文字列を復号してJSONから取得 Then 出力値 = ""
End Function
