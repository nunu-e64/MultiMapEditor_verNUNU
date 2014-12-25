Attribute VB_Name = "General"
'マルチマップエディター

Option Explicit

'マップデータの部分コピー用変数
Public CopyMap() As Byte
Public eCopyMap() As Byte
Public eCopy As Boolean

Public ChipNow As Integer
Public Const a As Long = 1
Public Const INIFILENAME As String = "MapEditor.ini"

Sub Wait(Wait_Time As Long)
'ＡＰＩ版ウェイト関数
    
    '使用する変数の定義
    Dim Start_Time As Long
    
    'Wait開始時の時間を取得
    Start_Time = timeGetTime()
    Do
        DoEvents    '他の処理を実行
        
        '設定時間到達のチェック
        If timeGetTime() - Start_Time > Wait_Time Then
            '到達したらループを抜ける
            Exit Do
        End If
    Loop

End Sub

Public Function Min(a, b)
    
    Min = IIf(a < b, a, b)
    
End Function
Public Function Max(a, b)
    
    Max = IIf(a > b, a, b)
    
End Function

Public Function GetINIValue(KEY As String) As String

    Dim Value As String * 255
    Call GetPrivateProfileString("SYSTEM", KEY, "ERROR", Value, Len(Value), App.Path & "\" & INIFILENAME)
    GetINIValue = Left$(Value, InStr(1, Value, vbNullChar) - 1)

End Function
Public Function SetINIValue(Value As String, KEY As String) As Boolean

    Dim Ret As Long
    Ret = WritePrivateProfileString("SYSTEM", KEY, Value, App.Path & "\" & INIFILENAME)
    SetINIValue = CBool(Ret)
End Function

