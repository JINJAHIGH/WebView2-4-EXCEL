VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "mod_WV2_安全タイマー"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function タイマー設定API Lib "user32" Alias "SetTimer" (ByVal ウィンドウハンドル As LongPtr, ByVal 件数識別番号イベント As LongPtr, ByVal 値間隔 As Long, ByVal タイマー呼出先 As LongPtr) As LongPtr

Private 保持フォーム As New frm_WV2_タイマー保持フォーム
Private 保持値間隔 As Long
Private 保持データ As Variant

Public Event タイマー発生(ByVal 通知時刻 As Long, ByRef 任意データ As Variant)

' =========================================================================
' 処理概要       : 指定間隔のタイマーを開始または停止
' 引数           : 値間隔 (Long) - タイマー間隔（ミリ秒）。0で停止
'                  任意データ (Variant) - タイマー通知へ引き渡すデータ
' =========================================================================
Public Sub 指定間隔のタイマーを開始または停止(ByVal 値間隔 As Long, Optional ByRef 任意データ As Variant)
    If 値間隔 = 保持値間隔 Then Exit Sub
    
    ' 間隔が変更された場合は、タイマーを保持するフォームを解放する
    Set 保持フォーム = Nothing
    If 値間隔 = 0 Then Exit Sub
    
    保持値間隔 = 値間隔
    If IsObject(任意データ) Then Set 保持データ = 任意データ Else 保持データ = 任意データ
    
    ' フォームの表示枠を通知先に指定し、Windowsタイマーを登録する
    タイマー設定API 保持フォーム.Controls.Add("Forms.Frame.1").[_GethWnd], ObjPtr(Me), 値間隔, 安全タイマーの中継関数アドレスを取得(Me)
End Sub

' =========================================================================
' 処理概要       : タイマーの通知をイベントとして転送
' 引数           : 通知メッセージ (Long) - タイマーのWindowsメッセージ番号
'                  件数識別番号イベント (Long) - タイマーの識別ID
'                  通知時刻 (Long) - タイマー通知時刻
' =========================================================================
Private Sub タイマーの通知をイベントとして転送(ByVal 通知メッセージ As Long, ByVal 件数識別番号イベント As Long, ByVal 通知時刻 As Long)
    On Error Resume Next
    RaiseEvent タイマー発生(通知時刻, 保持データ)
    On Error GoTo 0
End Sub
