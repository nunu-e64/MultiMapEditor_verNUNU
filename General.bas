Attribute VB_Name = "General"
'�}���`�}�b�v�G�f�B�^�[

Option Explicit

'�}�b�v�f�[�^�̕����R�s�[�p�ϐ�
Public CopyMap() As Byte
Public eCopyMap() As Byte
Public eCopy As Boolean

Public ChipNow As Integer
Public Const a As Long = 1
Public Const INIFILENAME As String = "MapEditor.ini"

Sub Wait(Wait_Time As Long)
'�`�o�h�ŃE�F�C�g�֐�
    
    '�g�p����ϐ��̒�`
    Dim Start_Time As Long
    
    'Wait�J�n���̎��Ԃ��擾
    Start_Time = timeGetTime()
    Do
        DoEvents    '���̏��������s
        
        '�ݒ莞�ԓ��B�̃`�F�b�N
        If timeGetTime() - Start_Time > Wait_Time Then
            '���B�����烋�[�v�𔲂���
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

