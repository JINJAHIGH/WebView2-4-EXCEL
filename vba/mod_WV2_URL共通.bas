Attribute VB_Name = "mod_WV2_URL共通"
Option Explicit

' =========================================================================
' 処理概要       : 文字列が英字のみで構成されるか判定
' 引数           : 作成文字列 (String) - 変換対象の文字列
' 戻り値         : Boolean - 空でなく英字のみの場合 True
' =========================================================================
Public Function 文字列が英字のみで構成されるか判定(ByVal 作成文字列 As String) As Boolean
    Dim 添字 As Long
    Dim 対象文字 As String
    
    If Len(作成文字列) = 0 Then Exit Function
    
    For 添字 = 1 To Len(作成文字列)
        対象文字 = LCase$(Mid$(作成文字列, 添字, 1))
        If 対象文字 < "a" Or 対象文字 > "z" Then Exit Function
    Next 添字
    文字列が英字のみで構成されるか判定 = True
End Function

' =========================================================================
' 処理概要       : 文字列をUTF-8でURLエンコード
' 引数           : 作成文字列 (String) - 変換対象の文字列
' 戻り値         : String - 文字列をUTF-8でURLエンコード
' =========================================================================
Public Function 文字列をUTF8でURLエンコード(ByVal 作成文字列 As String) As String
    Dim バイト数() As Byte
    Dim 配列下限 As Long, 配列上限 As Long
    Dim 出力文字列 As String
    Dim 添字 As Long
    Dim 入力データ As Integer
    Dim 対象文字 As String
    
    If Len(作成文字列) = 0 Then Exit Function
    
    バイト数 = 文字列をUTF8のバイト配列へ変換(作成文字列)
    
    On Error Resume Next
    配列下限 = LBound(バイト数)
    配列上限 = UBound(バイト数)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    
    ' UTF-8の各バイトを確認し、予約されていない文字以外をパーセントエンコードする
    For 添字 = 配列下限 To 配列上限
        入力データ = バイト数(添字)
        対象文字 = Chr$(入力データ)
        If (入力データ >= 65 And 入力データ <= 90) Or (入力データ >= 97 And 入力データ <= 122) Or (入力データ >= 48 And 入力データ <= 57) Or 入力データ = 45 Or 入力データ = 95 Or 入力データ = 46 Or 入力データ = 126 Then
        
            出力文字列 = 出力文字列 & 対象文字
        Else
            出力文字列 = 出力文字列 & "%" & Right$("0" & Hex$(入力データ), 2)
        End If
    Next 添字
    
    文字列をUTF8でURLエンコード = 出力文字列
End Function

' =========================================================================
' 処理概要       : 文字列をUTF-8のバイト配列へ変換
' 引数           : 作成文字列 (String) - 変換対象の文字列
' 戻り値         : Byte() - UTF-8のバイト配列
' =========================================================================
Public Function 文字列をUTF8のバイト配列へ変換(ByVal 作成文字列 As String) As Byte()
    Dim 作業情報 As Object
    Dim 未加工値() As Byte
    Dim 配列下限 As Long, 配列上限 As Long
    Dim 出力() As Byte
    Dim 副添字 As Long
    
    Set 作業情報 = CreateObject("ADODB.Stream")
    作業情報.Type = 2
    作業情報.Charset = "utf-8"
    作業情報.Open
    作業情報.WriteText 作成文字列
    作業情報.Position = 0
    作業情報.Type = 1
    
    未加工値 = 作業情報.Read
    作業情報.Close
    Set 作業情報 = Nothing
    
    On Error Resume Next
    配列下限 = LBound(未加工値)
    配列上限 = UBound(未加工値)
    If Err.Number <> 0 Then
    
        On Error GoTo 0
        文字列をUTF8のバイト配列へ変換 = 未加工値
        Exit Function
    End If
    On Error GoTo 0
    
    ' UTF-8のBOMが付いている場合は、URLに含めないよう先頭3バイトを除く
    If (配列上限 - 配列下限) >= 2 Then
        If 未加工値(配列下限) = &HEF And 未加工値(配列下限 + 1) = &HBB And 未加工値(配列下限 + 2) = &HBF Then
            ReDim 出力(0 To (配列上限 - 配列下限) - 3)
            For 副添字 = 配列下限 + 3 To 配列上限
                出力(副添字 - (配列下限 + 3)) = 未加工値(副添字)
            Next 副添字
            文字列をUTF8のバイト配列へ変換 = 出力
            Exit Function
        End If
    End If
    
    文字列をUTF8のバイト配列へ変換 = 未加工値
End Function
