VERSION 5.00
Begin VB.Form Map 
   AutoRedraw      =   -1  'True
   Caption         =   "Map [Sample.map] X:00 Y:00"
   ClientHeight    =   2610
   ClientLeft      =   6180
   ClientTop       =   4365
   ClientWidth     =   2760
   BeginProperty Font 
      Name            =   "�l�r �o�S�V�b�N"
      Size            =   15.75
      Charset         =   128
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Map.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   174
   ScaleMode       =   3  '�߸��
   ScaleWidth      =   184
   Begin VB.PictureBox SelectPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  '�Ȃ�
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   600
      ScaleHeight     =   35.31
      ScaleMode       =   0  'հ�ް
      ScaleWidth      =   32
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Chip 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  '�Ȃ�
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   1080
      ScaleHeight     =   29
      ScaleMode       =   3  '�߸��
      ScaleWidth      =   29
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox Chip 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  '�Ȃ�
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   600
      ScaleHeight     =   29
      ScaleMode       =   3  '�߸��
      ScaleWidth      =   29
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox SelectPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  '�Ȃ�
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   0
      ScaleHeight     =   35.31
      ScaleMode       =   0  'հ�ް
      ScaleWidth      =   32
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.HScrollBar X_Scroll 
      CausesValidation=   0   'False
      Height          =   240
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1740
      Width           =   795
   End
   Begin VB.VScrollBar Y_Scroll 
      CausesValidation=   0   'False
      Height          =   975
      Left            =   2340
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   240
   End
   Begin VB.PictureBox Chip 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  '�Ȃ�
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  '�߸��
      ScaleWidth      =   29
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox Crt 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '�Ȃ�
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  '�߸��
      ScaleWidth      =   29
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�}���`�}�b�v�G�f�B�^�[
Option Explicit

'�ҏW���̃}�b�v�t�@�C�������^�C�g���Ƃ��Ċi�[����
Private Title As String

'�\�����Ă���}�b�v���W
Public x As Long, y As Long

'�f�[�^�͈͑I�����̃}�b�v���W
Public Select_SX As Integer, Select_SY As Integer
Public Select_EX As Integer, Select_EY As Integer

'�_�u���N���b�N����p�}�b�v���W
Private MouseX As Single, MouseY As Single

'�}�b�v���i�[����ϐ�
Private Map() As Byte
Private RedoMap() As Byte
Private UndoMap() As Byte
Private eMap() As Byte
Private OldMap() As Byte
Private OldeMap() As Byte

Public MapSize As Long      '�}�b�v�̈�ӂ̃T�C�Y

'�ҏW���̃}�b�v���ۑ��p�ϐ�
Public SaveFileName As String
Public eSaveFileName As String
Public OpenTime As String
Public eOpenTime As String

'�`�b�v�I��ԍ�
Public LeftNo As Long
Public RightNo As Long
Public LeftDraw As Long
Public RightDraw As Long

'�c�[���̑I�����
Public Tool As String

'�ҏW���̃f�[�^�̏�ԁiTrue �ύX�FFalse ���ύX�j
Public DataChanged As Boolean
Public eDataChanged As Boolean

'INI�t�@�C���p
Private MyKey As String
Private eMyKey As String


Private Sub Crt_DblClick()
    Call Crt_MouseDown(1, 1, MouseX, MouseY)
End Sub

Private Sub Crt_MouseDown(Button As Integer, Shift As Integer, MX As Single, MY As Single)
'�}�b�v�̒u������
        
    Dim Ret As Integer
    On Error Resume Next
    
    MouseX = MX
    MouseY = MY
    
    If Shift = 0 Then
        Select Case Tool
        
            Case "Pen"
                '�`�b�v�z�u����
    
                '���{�^���̏���
                If Button = 1 Then
                    If Map(ChipNow, (x + (MX \ 32)) And MapSize, (y + (MY \ 32)) And MapSize) <> LeftNo Then
                        DataChanged = True
                        UndoSet
                    End If
                    
                    Map(ChipNow, (x + (MX \ 32)) And MapSize, (y + (MY \ 32)) And MapSize) = LeftNo
                    LeftDraw = 1
                    RightDraw = 0
                End If
                
                '�E�{�^���̏���
                If Button = 2 Then
                    If Map(ChipNow, (x + (MX \ 32)) And MapSize, (y + (MY \ 32)) And MapSize) <> RightNo Then
                        DataChanged = True
                        UndoSet
                    End If
                    
                    Map(ChipNow, (x + (MX \ 32)) And MapSize, (y + (MY \ 32)) And MapSize) = RightNo
                    LeftDraw = 0
                    RightDraw = 1
                End If
                '�}�b�v���ĕ`��
                MapShow
                
            Case "Syringe"
                '�X�|�C�g����
                
                '���{�^���̏���
                If Button = 1 Then
                    LeftNo = Map(ChipNow, (x + (MX \ 32)) And MapSize, (y + (MY \ 32)) And MapSize)
                End If
                '�E�{�^���̏���
                If Button = 2 Then
                    RightNo = Map(ChipNow, (x + (MX \ 32)) And MapSize, (y + (MY \ 32)) And MapSize)
                End If
                '�z���o�����ԍ��ŕ\���̕ύX
                ToolChipShow
                
            Case "Cursor"
                '�f�[�^�̑I������
                
                If Button = 1 Then
                    LeftDraw = 1
                    eCopy = False
                ElseIf Button = 2 Then
                    RightDraw = 1
                    eCopy = True
                End If
                
                    Select_SX = x + MX \ 32
                    Select_SY = y + MY \ 32
                    Select_EX = Select_SX
                    Select_EY = Select_SY
                    MapShow
                    'SelectShow
                    
                    'ToolTipText�N���b�N�����ꏊ�̃}�b�v�f�[�^�ԍ�
                    Me.Crt.ToolTipText = Map(0, x + MX \ 32, y + MY \ 32) & "-" & Map(1, x + MX \ 32, y + MY \ 32)

                
            Case "Paint"
                '�h��ׂ�����
                Ret = MsgBox("�I������Ă���`�b�v�œh��ׂ��܂�", vbOKCancel + vbQuestion, "MapEditor")
                If Ret = vbOK Then
                    UndoSet
                    DataChanged = True
                    MapPaint IIf(Button <> 2, LeftNo, RightNo)
                End If
        
        End Select
    
    Else
        On Error Resume Next
        Dim tmp As String
        
        tmp = InputBox("�C�x���g�ԍ�����͂��Ă��������i0~255�j", "�C�x���g�ԍ�����[X:" & (x + (MouseX \ 32)) Mod (MapSize + 1) & " Y:" & (y + (MouseY \ 32)) Mod (MapSize + 1) & "]", CStr(eMap((x + (MouseX \ 32)) Mod (MapSize + 1), (y + (MouseY \ 32)) Mod (MapSize + 1))))
        If tmp <> "" Then
            If CLng(tmp) < 0 Or CLng(tmp) > 255 Then
                Call MsgBox("�C�x���g�ԍ���0~255�̐����œ��͂��Ă�������", vbOKOnly, "�C�x���g�ԍ�����[X:" & x + (MouseX \ 32) & " Y:" & y + (MouseY \ 32) & "]")
                Call Crt_DblClick
            ElseIf CLng(tmp) <> eMap((x + (MouseX \ 32)) Mod (MapSize + 1), (y + (MouseY \ 32)) Mod (MapSize + 1)) Then
                eMap((x + (MouseX \ 32)) Mod (MapSize + 1), (y + (MouseY \ 32)) Mod (MapSize + 1)) = CLng(tmp)
                '�f�[�^�̕ύX���L������
                eDataChanged = True
                MapShow
            End If
        End If
    
    End If
    
End Sub

Private Sub Crt_MouseMove(Button As Integer, Shift As Integer, MX As Single, MY As Single)
'�}�E�X�̈ړ����̏���
On Error Resume Next

    Select Case Tool
    
        Case "Pen"
            '�A���f�[�^�z�u����
            If LeftDraw = 1 And (Crt.Width > MX And Crt.Height > MY) Then
                Map(ChipNow, (x + (MX \ 32)) And MapSize, (y + (MY \ 32)) And MapSize) = LeftNo
                MapShow
            End If
            If RightDraw = 1 And (Crt.Width > MX And Crt.Height > MY) Then
                Map(ChipNow, (x + (MX \ 32)) And MapSize, (y + (MY \ 32)) And MapSize) = RightNo
                MapShow
            End If
            
        Case "Cursor"
            '�I��͈͊g�又��
            If (LeftDraw = 1 Or RightDraw = 1) And (Crt.Width > MX And Crt.Height > MY) Then
                Select_EX = x + MX \ 32
                Select_EY = y + MY \ 32
                MapShow
                'SelectShow
            
            ElseIf (Crt.Width > MX And Crt.Height > MY) Then
                'ToolTipText�N���b�N�����ꏊ�̃}�b�v�f�[�^�ԍ�
                Me.Crt.ToolTipText = "" '��x���Z�b�g���邱�Ƃŕ\���ӏ����|�C���^�ʒu��
                Me.Crt.ToolTipText = Map(0, x + MX \ 32, y + MY \ 32) & "-" & Map(1, x + MX \ 32, y + MY \ 32)
            End If
                
    End Select
    

    
End Sub

Private Sub Crt_MouseUp(Button As Integer, Shift As Integer, MX As Single, MY As Single)
'�{�^���������ꂽ�ꍇ�̏���

    Select Case Tool
        Case "Cursor", "Pen"
            If Button = 1 Then
                LeftDraw = 0
            End If
            If Button = 2 Then
                RightDraw = 0
            End If
    End Select

End Sub


Private Sub Form_Activate()
Dim i As Long

'�A�N�e�B�u�ɂȂ������ɂl�c�h�t�H�[���̃`�b�v��؂肩����
    ChipBarShow (0)
    ChipBarShow (1)
    ChipBarShow (2)
    ToolForm.Tool1.Buttons(Tool).Value = tbrPressed
    ToolChipShow
    
    For i = 1 To 4
        MainForm.Menu101(i).Checked = False
    Next i
    Select Case MapSize
    Case 255
        MainForm.Menu101(1).Checked = True
    Case 127
        MainForm.Menu101(2).Checked = True
    Case 63
        MainForm.Menu101(3).Checked = True
    Case Else
        MainForm.Menu101(4).Checked = True
    End Select
    
End Sub

Public Sub KeyToKey(KeyCode As Integer, Shift As Integer)

    Call Form_KeyDown(KeyCode, Shift)

End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    '�E�L�[�̏���
    Case vbKeyRight, vbKeyD
        X_Scroll.Value = (X_Scroll.Value + 1 + MapSize + 1) Mod (MapSize + 1)
                                   
    '���L�[�̏���
    Case vbKeyLeft, vbKeyA
        X_Scroll.Value = (X_Scroll.Value - 1 + MapSize + 1) Mod (MapSize + 1)
        
    '��L�[�̏���
    Case vbKeyUp, vbKeyW
        Y_Scroll.Value = (Y_Scroll.Value - 1 + MapSize + 1) Mod (MapSize + 1)
                
    '���L�[�̏���
    Case vbKeyDown, vbKeyS
        Y_Scroll.Value = (Y_Scroll.Value + 1 + MapSize + 1) Mod (MapSize + 1)
                   
    '�I��͈͂�0��
    Case vbKeyDelete And Tool = "Cursor"
        MapDelete (Shift)
            
    Case vbKeyReturn And Tool = "Cursor"
        MapSelectSet
    
    '�c�[���̐؂�ւ�
    Case vbKey1 And ToolForm.Tool1.Buttons(1).Enabled
        ToolForm.Tool1.Buttons(1).Value = tbrPressed
        Call ToolForm.Tool1_ButtonClick(ToolForm.Tool1.Buttons(1))
        
    Case vbKey2 And ToolForm.Tool1.Buttons(2).Enabled
        ToolForm.Tool1.Buttons(2).Value = tbrPressed
        Call ToolForm.Tool1_ButtonClick(ToolForm.Tool1.Buttons(2))
        
    Case vbKey3 And ToolForm.Tool1.Buttons(3).Enabled
        ToolForm.Tool1.Buttons(3).Value = tbrPressed
        Call ToolForm.Tool1_ButtonClick(ToolForm.Tool1.Buttons(3))
        
    Case vbKey4 And ToolForm.Tool1.Buttons(4).Enabled
        ToolForm.Tool1.Buttons(4).Value = tbrPressed
        Call ToolForm.Tool1_ButtonClick(ToolForm.Tool1.Buttons(4))
        
    Case vbKeyZ And Shift = 2 And ToolForm.Tool1.Buttons("Undo").Enabled
        Call ToolForm.Tool1_ButtonClick(ToolForm.Tool1.Buttons("Undo"))
    
    Case vbKeyY And Shift = 2 And ToolForm.Tool1.Buttons("Redo").Enabled
        Call ToolForm.Tool1_ButtonClick(ToolForm.Tool1.Buttons("Redo"))
    
    Case vbKeyC And Shift = 2 And ToolForm.Tool1.Buttons("Copy").Enabled
        Call ToolForm.Tool1_ButtonClick(ToolForm.Tool1.Buttons("Copy"))

    Case vbKeyV And Shift = 2 And ToolForm.Tool1.Buttons("Paste").Enabled
        Call ToolForm.Tool1_ButtonClick(ToolForm.Tool1.Buttons("Paste"))

    'ChipChange
    Case vbKeyTab 'And Shift = 2
        If MainForm.Top_bar.Buttons("ChipChange").Enabled Then Call MainForm.Top_bar_ButtonClick(MainForm.Top_bar.Buttons("ChipChange"))
    
    '����
    Case vbKeyF And Shift = 2
        Call Search(False)
    Case vbKeyF And Shift = 3
        Call Search(True)
        
    '�u��
    Case vbKeyH And Shift = 2
        Call Replace(False)
    Case vbKeyH And Shift = 3
        Call Replace(True)
    
    
    End Select
    
End Sub

Private Sub Search(IsEvent As Boolean)

    Dim Counting As Integer
    Dim SearchNum As String
    Dim i As Integer, j As Integer

    On Error Resume Next
    
    If Not IsEvent Then
        SearchNum = InputBox(IIf(ChipNow = 0, "���w", "��w") & "���猟���������f�[�^�ԍ�����͂��Ă�������(0~255)")
        If SearchNum <> "" Then
            If CInt(SearchNum) < 0 Or CInt(SearchNum) > 255 Then
                Call Search(False)
                Exit Sub
            Else
                For i = 0 To MapSize
                For j = 0 To MapSize
                    If Map(ChipNow, i, j) = CInt(SearchNum) Then Counting = Counting + 1
                Next
                Next
                
                MsgBox SearchNum & "�Ԃ�" & CStr(Counting) & "������܂���"
            End If
        End If
    Else
        SearchNum = InputBox("�C�x���g�}�b�v���猟���������f�[�^�ԍ�����͂��Ă�������(0~255)")
        If SearchNum <> "" Then
            If SearchNum < 0 Or SearchNum > 255 Then
                Call Search(True)
                Exit Sub
            Else
                For i = 0 To MapSize
                For j = 0 To MapSize
                    If eMap(i, j) = CInt(SearchNum) Then Counting = Counting + 1
                Next
                Next
                
                MsgBox SearchNum & "�Ԃ�" & CStr(Counting) & "������܂���"
            End If
        End If
    End If
    

End Sub

Private Sub Replace(IsEvent As Boolean)
    Dim Counting As Integer
    Dim SearchNum As String
    Dim ReplaceNum As String
    Dim i As Integer, j As Integer
    
    On Error Resume Next
    
    If Not IsEvent Then
        SearchNum = InputBox(IIf(ChipNow = 0, "���w", "��w") & "����u���������f�[�^�ԍ�����͂��Ă�������(0~255)")
        If SearchNum <> "" Then
            If CInt(SearchNum) < 0 Or CInt(SearchNum) > 255 Then
                Call Replace(False)
                Exit Sub
            Else
                ReplaceNum = InputBox(IIf(ChipNow = 0, "���w��", "��w��") & SearchNum & "�Ԃ����Ԃɒu���������̂ł���(0~255)")
                        
                If ReplaceNum <> "" Then
                    If CInt(ReplaceNum) < 0 Or CInt(ReplaceNum) > 255 Then
                        Call Replace(False)
                        Exit Sub
                    Else
                        UndoSet
                        
                        For i = 0 To MapSize
                        For j = 0 To MapSize
                            If Map(ChipNow, i, j) = CInt(SearchNum) Then
                                Map(ChipNow, i, j) = CInt(ReplaceNum)
                                Counting = Counting + 1
                            End If
                        Next
                        Next
                        
                        MsgBox CStr(Counting) & "��" & SearchNum & "��" & ReplaceNum & "�ɒu������܂���"
                        DataChanged = True
                        
                    End If
                End If
                
            End If
        End If
    
    Else
        
        SearchNum = InputBox("�C�x���g�}�b�v����u���������f�[�^�ԍ�����͂��Ă�������(0~255)")
        If SearchNum <> "" Then
            If CInt(SearchNum) < 0 Or CInt(SearchNum) > 255 Then
                Call Replace(True)
                Exit Sub
            Else
                ReplaceNum = InputBox("�C�x���g�}�b�v" & SearchNum & "�Ԃ����Ԃɒu���������̂ł���(0~255)")
                        
                If ReplaceNum <> "" Then
                    If CInt(ReplaceNum) < 0 Or CInt(ReplaceNum) > 255 Then
                        Call Replace(True)
                        Exit Sub
                    Else
                        For i = 0 To MapSize
                        For j = 0 To MapSize
                            If eMap(i, j) = CInt(SearchNum) Then
                                Counting = Counting + 1
                                eMap(i, j) = CInt(ReplaceNum)
                            End If
                        Next
                        Next
                        
                        MsgBox CStr(Counting) & "��" & SearchNum & "��" & ReplaceNum & "�ɒu������܂���"
                        eDataChanged = True
                        
                    End If
                End If
                
            End If
        End If
        
    End If

    MapShow
End Sub

Private Sub Form_Load()
'�}�b�v�z�u�p�t�H�[���̃��[�h�C�x���g
        
    '�}�b�v�T�C�Y�̐ݒ�
    Select Case True
    Case MainForm.Menu101(1).Checked
        MapSize = 256 - 1
    Case MainForm.Menu101(2).Checked
        MapSize = 128 - 1
    Case MainForm.Menu101(3).Checked
        MapSize = 64 - 1
    Case Else
        MapSize = 128 - 1
    End Select
    
    ReDim Map(1, 0 To 255, 0 To 255)    '$
    ReDim eMap(0 To 255, 0 To 255)      '$
    
    '�}�b�v�\���p�̃s�N�`���{�b�N�X�̈ʒu�̏�����
    Crt.Top = 0
    Crt.Left = 0
    
    Chip(0).Width = 512
    Chip(0).Height = 512
    Chip(1).Width = 512
    Chip(1).Height = 512
    
    MapReSize
    
    x = 0: y = 0
    Title = "NewMap(NoName)"

    '�c�[���o�[�̈ꕔ��L���ɂ���
    With MainForm
        .Top_bar.Buttons("Chip").Enabled = True
        .Top_bar.Buttons("Chip2").Enabled = True
        .Top_bar.Buttons("Map").Enabled = True
        .Top_bar.Buttons("Save").Enabled = True
        .Top_bar.Buttons("eSave").Enabled = True
        .Top_bar.Buttons("ChipChange").Enabled = True
    End With
    MainForm.MenuTrue
    
    X_Scroll.Max = MapSize
    Y_Scroll.Max = MapSize
    X_Scroll.Value = x
    Y_Scroll.Value = y
    
    Tool = "Pen"
    
End Sub

Private Sub MapReSize()
'�t�H�[���T�C�Y�Ƀs�N�`���{�b�N�X�̃T�C�Y�����킹��

    On Error Resume Next    '���̃��[�`�����̃G���[�𖳌��ɂ���B

    '�}�b�v�\���p�̃s�N�`���{�b�N�X�̃T�C�Y����
    Crt.Width = Me.ScaleWidth - 16
    Crt.Height = Me.ScaleHeight - 16
    
    '�X�N���[���o�[�̃T�C�Y����
    Y_Scroll.Top = 0
    Y_Scroll.Left = Me.ScaleWidth - 16
    Y_Scroll.Height = Me.ScaleHeight - 16
    
    X_Scroll.Top = Me.ScaleHeight - 16
    X_Scroll.Left = 0
    X_Scroll.Width = Me.ScaleWidth - 16

    MapShow
    
    
End Sub

Private Sub Form_Resize()
'�t�H�[���̑傫����ύX���ꂽ�ꍇ�̏���

    MapReSize
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'�E�B���h�E����鎞�̏���

    Dim Ret As Integer
    
    If DataChanged = True Then
        
        Ret = MsgBox("�ҏW���̃}�b�v�f�[�^�͕ύX����Ă��܂��B" & vbCrLf & "�}�b�v�f�[�^��ۑ����܂����H", vbYesNoCancel + vbExclamation, "MapEditor")
        Select Case Ret
            '�L�����Z���{�^���Ȃ�I��������߂�
            Case vbCancel
                Cancel = True
                Exit Sub
            '�n�j�Ȃ�t�@�C���Z�[�u���[�`�������s�A�A�������ŃL�����Z�����ꂽ���͂�I���͂��Ȃ�
            Case vbYes
                If MainForm.MapSave = False Then
                    Cancel = True
                    Exit Sub
                End If
            Case vbNo
        End Select
    End If
    
    If eDataChanged = True Then
        
        Ret = MsgBox("�ҏW���̃C�x���g�}�b�v�f�[�^�͕ύX����Ă��܂��B" & vbCrLf & "�C�x���g�}�b�v�f�[�^��ۑ����܂����H", vbYesNoCancel + vbExclamation, "MapEditor")
        Select Case Ret
            '�L�����Z���{�^���Ȃ�I��������߂�
            Case vbCancel
                Cancel = True
                Exit Sub
            '�n�j�Ȃ�t�@�C���Z�[�u���[�`�������s�A�A�������ŃL�����Z�����ꂽ���͂�I���͂��Ȃ�
            Case vbYes
                If MainForm.eMapSave = False Then
                    Cancel = True
                    Exit Sub
                End If
            Case vbNo
        End Select
    End If

    '�J���Ă���t�H�[���̐������炷
    MainForm.FormCounter = MainForm.FormCounter - 1

    '�t�H�[���ɕt������`�b�v�̕\���Ȃǂ��N���A����
    MainForm.ShowChip(0).Cls
    MainForm.ShowChip(1).Cls
    MainForm.ShowChip(2).Cls
    ToolForm.LeftPic.Cls
    ToolForm.RightPic.Cls
    
    '�������t�H�[�����Ōォ�ǂ������ׂ�
    If MainForm.FormCounter = 0 Then
    
        '�c�[���o�[�̈ꕔ�𖳌��ɂ���
        With MainForm.Top_bar
            .Buttons("Chip").Enabled = False
            .Buttons("Chip2").Enabled = False
            .Buttons("Map").Enabled = False
            .Buttons("Save").Enabled = False
            .Buttons("eSave").Enabled = False
            .Buttons("ChipChange").Enabled = False
        End With
        MainForm.MenuFalse
        
    End If
    
    If MyKey <> "" Then Call SetINIValue("", MyKey)
    If eMyKey <> "" Then Call SetINIValue("", eMyKey)
    
End Sub

Public Sub ChipLoad(FileName As String, i As Integer)
'�w�肳�ꂽ�t�@�C�����Ń}�b�v�`�b�v�����[�h����
    If Right(FileName, 4) = ".png" Or Right(FileName, 4) = ".PNG" Then
        Chip(i).Picture = LoadPNG(FileName)
    Else
        Chip(i).Picture = LoadPicture(FileName)
    End If
    
    ChipBarShow (i)
    ToolChipShow
    If (i = 0 Or i = 2) Then MapShow

End Sub
Public Sub MapLoad(FileName As String, Optional ForE As Boolean = False)
Dim i As Long
'�w�肳�ꂽ�t�@�C�����Ń}�b�v�����[�h����
'On Error Resume Next
    
    Do While i < 100
        If FileName = GetINIValue("NowEditFile" & CStr(i) & "-0") Or FileName = GetINIValue("NowEditFile" & CStr(i) & "-1") Then
            Call MsgBox("���̃t�@�C���͂��łɊJ���Ă���\��������܂��B" & vbCrLf & FileName & vbCrLf & vbCrLf & "�i��肪�Ȃ��ɂ�������炸���̃��b�Z�[�W���o��ꍇ�A" & vbCrLf & "�u�t�@�C��(F)->MapEditor�̏ꏊ���J���v����MapEditor.ini���C�����邩�폜���Ă��������B" & vbCrLf & "�܂����̍ہANUNU�܂Ńo�O�񍐂���������΍K���ł��j")
            Exit Sub
        End If
        i = i + 1
    Loop
    
    'FileName���o�C�i���|���[�h�ŃI�[�v�����Ă��̂܂ܕϐ��ɓǂݍ���
    Open FileName For Binary Access Read As #1
        If Not ForE Then
            ReDim Map(1, 0 To MapSize, 0 To MapSize)
            Get #1, , Map
            SaveFileName = FileName
            
            Title = FileName
            ReDim OldMap(1, 0 To MapSize, 0 To MapSize)
            OldMap = Map
            DataChanged = False
            OpenTime = Format(Now(), "YYMMDD_HHMM")
        Else
            ReDim eMap(0 To MapSize, 0 To MapSize)
            Get #1, , eMap
            eSaveFileName = FileName
            
            ReDim OldeMap(0 To MapSize, 0 To MapSize)
            OldeMap = eMap
            eDataChanged = False
            eOpenTime = Format(Now(), "YYMMDD_HHMM")
        End If
    Close #1
    
    If Not (ForE) Then
        If MyKey = "" Then
            i = 0
            Do
                If GetINIValue("NowEditFile" & CStr(i) & "-0") = "" Or GetINIValue("NowEditFile" & CStr(i) & "-0") = "ERROR" Then
                    MyKey = "NowEditFile" & CStr(i) & "-0"
                    Exit Do
                End If
                i = i + 1
            Loop
        End If
        Call SetINIValue(FileName, MyKey)
    Else
        If eMyKey = "" Then
            i = 0
            Do
                If GetINIValue("NowEditFile" & CStr(i) & "-1") = "" Or GetINIValue("NowEditFile" & CStr(i) & "-1") = "ERROR" Then
                    eMyKey = "NowEditFile" & CStr(i) & "-1"
                    Exit Do
                End If
                i = i + 1
            Loop
        End If
        Call SetINIValue(FileName, eMyKey)
    End If
    
    MapShow
    
End Sub
Public Sub MapSave(FileName As String, Optional ForE As Boolean = False)
On Error GoTo SaveError

Dim i As Long
Dim BackUpName As String
Dim BackUpSave As Boolean
        If Dir(FileName) <> "" Then
            BackUpSave = True
            BackUpName = Left(Dir(FileName), Len(Dir(FileName)) - InStr(1, StrReverse(Dir(FileName)), ".")) & "_" & IIf(ForE, eOpenTime, OpenTime) & Right(FileName, InStr(StrReverse(FileName), "."))
            BackUpName = App.Path & "\tmp\" & BackUpName
            
            If Dir(App.Path & "\tmp", vbDirectory) = "" Then
                MkDir App.Path & "\tmp"
            End If
            
            If Dir(BackUpName) <> "" Then BackUpSave = False    '�o�b�N�A�b�v�͈��
        End If
            
    'FILE���o�C�i���|���[�h�ŃI�[�v�����ĕϐ������̂܂܏�����
    Open FileName For Binary Access Write As #1
    If BackUpSave Then Open BackUpName For Binary Access Write As #2
    
    If Not ForE Then
        Put #1, , Map
        If DataChanged And BackUpSave Then Put #2, , OldMap
        DataChanged = False
        SaveFileName = FileName
        Title = FileName
        
        If MyKey = "" Then
            i = 0
            Do
                If GetINIValue("NowEditFile" & CStr(i) & "-0") = "" Or GetINIValue("NowEditFile" & CStr(i) & "-0") = "ERROR" Then
                    MyKey = "NowEditFile" & CStr(i) & "-0"
                    Exit Do
                    End If
            Loop
        End If
        Call SetINIValue(FileName, MyKey)

    Else
        Put #1, , eMap
        If eDataChanged And BackUpSave Then Put #2, , OldeMap
        eDataChanged = False
        eSaveFileName = FileName
        
        If eMyKey = "" Then
            i = 0
            Do
                If GetINIValue("NowEditFile" & CStr(i) & "-1") = "" Or GetINIValue("NowEditFile" & CStr(i) & "-1") = "ERROR" Then
                    eMyKey = "NowEditFile" & CStr(i) & "-1"
                    Exit Do
                End If
            Loop
        End If
        Call SetINIValue(FileName, eMyKey)
    End If
    
    Close #1
    Close #2
    
    
    '�}�b�v�̍ĕ`��
    MapShow
    Exit Sub

SaveError:

    MsgBox "Error->MapSave->" & Err.Description

   
End Sub
Public Sub ChangeMapSize(Size As Long)
'�}�b�v�T�C�Y�̕ύX
    
    If Size < 16 Or Size > 256 Then
        Call MsgBox("�}�b�v�T�C�Y��16�ȏ�256�ȉ��Ŏw�肵�Ă�������", vbOKOnly, "�G���[�F�}�b�v�T�C�Y�̕ύX")
        Exit Sub
    End If
    
    MapSize = Size - 1
    'ReDim Preserve eMap(0 To MapSize, 0 To MapSize)    $
    'ReDim Preserve Map(1, 0 To MapSize, 0 To MapSize)  $
    
    If MyKey <> "" Then Call SetINIValue("", MyKey)
    If eMyKey <> "" Then Call SetINIValue("", eMyKey)
    
    If SaveFileName <> "" Then Call MapLoad(SaveFileName, False)
    If eSaveFileName <> "" Then Call MapLoad(eSaveFileName, True)
    
    X_Scroll.Max = MapSize
    Y_Scroll.Max = MapSize
    X_Scroll.Value = x
    Y_Scroll.Value = y
    MapShow

End Sub
Public Sub MapShow()
'�}�b�v�̕\�����s��
    On Error Resume Next
    Dim i As Long, j As Long
    Dim HX As Long, HY As Long
    Dim ShowX As Long, ShowY As Long
    
    ShowX = Crt.Width \ 32
    ShowY = Crt.Height \ 32
    
    '���w�̕`��
        For i = 0 To ShowY
            For j = 0 To ShowX
                HX = (Map(0, ((x + j) And MapSize), ((y + i) And MapSize)) And 8 * a - 1) * 32
                HY = (Map(0, ((x + j) And MapSize), ((y + i) And MapSize)) And (&HF8 - 8 * (a - 1))) * 4
                BitBlt Me.Crt.hdc, j * 32, i * 32, 32, 32, MainForm.ShowChip(0).hdc, HX, HY, SrcCopy
            Next j
        Next i
    
    '��w�̕`��
        For i = 0 To ShowY
            For j = 0 To ShowX
                HX = (Map(1, (x + j) And MapSize, (y + i) And MapSize) And 8 * a - 1) * 32
                HY = (Map(1, (x + j) And MapSize, (y + i) And MapSize) And (&HF8 - 8 * (a - 1))) * 4
                If HX + HY <> 0 Then
                    Call BitBlt(Me.Crt.hdc, j * 32, i * 32, 32, 32, MainForm.ShowChip(2).hdc, HX, HY, SrcAnd)
                    Call BitBlt(Me.Crt.hdc, j * 32, i * 32, 32, 32, MainForm.ShowChip(1).hdc, HX, HY, SrcPaint)
                End If
            Next j
        Next i
        
    '�C�x���g�}�b�v�̕����`��
        Dim Tmp_Rect As RECT
        Dim Tmp_subRect As RECT
        With Tmp_Rect
        For i = 0 To ShowY
            .Top = i * 32 + (32 - 14) / 2
            .Bottom = (i + 1) * 32 - (32 - 14) / 2
            Tmp_subRect.Top = .Top + 1
            Tmp_subRect.Bottom = .Bottom + 1
            
            For j = 0 To ShowX
                If eMap((x + j) And MapSize, (y + i) And MapSize) <> 0 Then
                    .Left = j * 32
                    .Right = .Left + 32
                    Tmp_subRect.Left = .Left + 1
                    Tmp_subRect.Right = .Right + 1
                    
                    Me.Crt.ForeColor = RGB(0, 0, 0)
                    Call DrawText(Me.Crt.hdc, CStr(eMap((x + j) And MapSize, (y + i) And MapSize)), -1, Tmp_subRect, DT_CENTER)
                    Me.Crt.ForeColor = RGB(255, 0, 0)
                    Call DrawText(Me.Crt.hdc, CStr(eMap((x + j) And MapSize, (y + i) And MapSize)), -1, Tmp_Rect, DT_CENTER)
                End If
            Next j
        Next i
        End With
    
    '�L���v�V�����Ɍ��݂̍��W��\������
        Me.Caption = IIf(DataChanged, "*", "") & IIf(eDataChanged, "^", "") & IIf(Len(Title) > 20, "..." & Right(Title, 20), Title) & "[X:" & x & " Y:" & y & "] " & Map(0, x, y) & "-" & Map(1, x, y)
        'Me.Caption = Title & "[X:" & Hex(X) & " Y:" & Hex(Y) & "]"     16�i���\�L
    
    If Tool = "Cursor" Then SelectShow
    
    Crt.Refresh
    
End Sub
Public Sub SelectShow()
    
    Dim i As Integer, j As Integer
    Dim D_X As Integer, D_Y As Integer
    
    Dim StartX As Integer:  StartX = Select_SX
    Dim StartY As Integer:  StartY = Select_SY
    Dim EndX As Integer:  EndX = Select_EX
    Dim EndY As Integer:  EndY = Select_EY
    
    
    '�I��͈͂��}�C�i�X�����̏ꍇ�J�n�n�_�ƏI���n�_����ꊷ����
    If StartX > EndX Then
        D_X = StartX
        StartX = EndX
        EndX = D_X
    End If
    If StartY > EndY Then
        D_Y = StartY
        StartY = EndY
        EndY = D_Y
    End If
    
    '�I��͈֖͂Ԋ|����`�悷��i���ۂ͂����̂n�q�]���j
    For i = 0 To EndY - StartY
        For j = 0 To EndX - StartX
            BitBlt Me.Crt.hdc, (j + (StartX - x)) * 32, (i + (StartY - y)) * 32, 32, 32, IIf(eCopy, SelectPic(1).hdc, SelectPic(0).hdc), 0, 0, SrcPaint
        Next j
    Next i

    
    '�ĕ`����s��
    Crt.Refresh
    
End Sub

Public Sub MapSelectSet()
'�I�𕔕��̔ԍ����ꊇ�ύX����

    Dim h As Integer, i As Integer, j As Integer
    Dim StartX As Integer, StartY As Integer
    Dim EndX As Integer, EndY As Integer
    
    Dim tmp As String
    
    '�I��͈͂��}�C�i�X�����̏ꍇ�J�n�n�_�ƏI���n�_����ꊷ����
    If Select_SX > Select_EX Then
        StartX = Select_EX
        EndX = Select_SX
    Else
        StartX = Select_SX
        EndX = Select_EX
    End If
    If Select_SY > Select_EY Then
        StartY = Select_EY
        EndY = Select_SY
    Else
        StartY = Select_SY
        EndY = Select_EY
    End If
    
    
    If Not eCopy Then
            
        On Error Resume Next
        
        tmp = InputBox("�}�b�v�`�b�v�f�[�^�ԍ�����͂��Ă��������i0~255�j", IIf(ChipNow = 0, "���w", "��w") & "�I��͈͈ꊇ�ύX")
        If tmp <> "" Then
            If CLng(tmp) < 0 Or CLng(tmp) > 255 Then
                Call MsgBox("�ԍ���0~255�̐����œ��͂��Ă�������", vbOKOnly, IIf(ChipNow = 0, "���w", "��w") & "�I��͈͈ꊇ�ύX")
                MapSelectSet
                Exit Sub
            Else
                ReDim CopyMap(0 To EndX - StartX, 0 To EndY - StartY)
                Call UndoSet
                For i = 0 To EndY - StartY
                    For j = 0 To EndX - StartX
                        Map(ChipNow, (j + StartX) Mod (MapSize + 1), (i + StartY) Mod (MapSize + 1)) = CLng(tmp)
                    Next j
                Next i
                DataChanged = True
            End If
        End If
        
    Else
        On Error Resume Next
        
        tmp = InputBox("�C�x���g�ԍ�����͂��Ă��������i0~255�j", "�C�x���g�ԍ��ꊇ����")
        If tmp <> "" Then
            If CLng(tmp) < 0 Or CLng(tmp) > 255 Then
                Call MsgBox("�C�x���g�ԍ���0~255�̐����œ��͂��Ă�������", vbOKOnly, "�C�x���g�ԍ��ꊇ����")
                MapSelectSet
                Exit Sub
            Else
                ReDim eCopyMap(0 To EndX - StartX, 0 To EndY - StartY)
                For i = 0 To EndY - StartY
                    For j = 0 To EndX - StartX
                        eMap((j + StartX) Mod (MapSize + 1), (i + StartY) Mod (MapSize + 1)) = CLng(tmp)
                    Next j
                Next i
                eDataChanged = True
            End If
        End If
    End If
        
    MapShow

End Sub


Public Sub MapDelete(Optional Shift As Integer)
'�ҏW���̑I�𕔕����폜����

    Dim h As Integer, i As Integer, j As Integer
    Dim StartX As Integer, StartY As Integer
    Dim EndX As Integer, EndY As Integer
    
    '�I��͈͂��}�C�i�X�����̏ꍇ�J�n�n�_�ƏI���n�_����ꊷ����
    If Select_SX > Select_EX Then
        StartX = Select_EX
        EndX = Select_SX
    Else
        StartX = Select_SX
        EndX = Select_EX
    End If
    If Select_SY > Select_EY Then
        StartY = Select_EY
        EndY = Select_SY
    Else
        StartY = Select_SY
        EndY = Select_EY
    End If
    
    
    If Not eCopy Then
        Call UndoSet
        
        ReDim CopyMap(1, 0 To EndX - StartX, 0 To EndY - StartY)
        For i = 0 To EndY - StartY
            For j = 0 To EndX - StartX
                Map(ChipNow, (j + StartX) Mod (MapSize + 1), (i + StartY) Mod (MapSize + 1)) = IIf(Shift = 1, 0, RightNo) '���ܕҏW���̑w��RightNo�ɂ��܂��B�V�t�g�L�[�����Ȃ���Ȃ�0�ɂ��܂�
            Next j
        Next i
        DataChanged = True
        
    Else
        ReDim eCopyMap(0 To EndX - StartX, 0 To EndY - StartY)
        For i = 0 To EndY - StartY
            For j = 0 To EndX - StartX
                eMap((j + StartX) Mod (MapSize + 1), (i + StartY) Mod (MapSize + 1)) = 0
            Next j
        Next i
        eCopy = True
        eDataChanged = True
    End If
        
    MapShow

End Sub
Public Sub MapCopy()
'�ҏW���̑I�𕔕����R�s�[����

    Dim h As Integer, i As Integer, j As Integer
    Dim StartX As Integer, StartY As Integer
    Dim EndX As Integer, EndY As Integer
    
    '�I��͈͂��}�C�i�X�����̏ꍇ�J�n�n�_�ƏI���n�_����ꊷ����
    If Select_SX > Select_EX Then
        StartX = Select_EX
        EndX = Select_SX
    Else
        StartX = Select_SX
        EndX = Select_EX
    End If
    If Select_SY > Select_EY Then
        StartY = Select_EY
        EndY = Select_SY
    Else
        StartY = Select_SY
        EndY = Select_EY
    End If
    
    If Not eCopy Then
        ReDim CopyMap(1, 0 To EndX - StartX, 0 To EndY - StartY)
        For h = 0 To 1
        For i = 0 To EndY - StartY
            For j = 0 To EndX - StartX
                CopyMap(h, j, i) = Map(h, (j + StartX) Mod (MapSize + 1), (i + StartY) Mod (MapSize + 1))
            Next j
        Next i
        Next
        eCopy = False
    Else
        ReDim eCopyMap(0 To EndX - StartX, 0 To EndY - StartY)
        For i = 0 To EndY - StartY
            For j = 0 To EndX - StartX
                eCopyMap(j, i) = eMap((j + StartX) Mod (MapSize + 1), (i + StartY) Mod (MapSize + 1))
            Next j
        Next i
        eCopy = True
    End If
        
End Sub

Public Sub MapPast()
'�R�s�[�����}�b�v�f�[�^��\��t����
On Error GoTo MapPastEnd
    Dim h As Integer, i As Integer, j As Integer
    
    For h = 0 To 1
    For i = 0 To UBound(CopyMap, 3)
        For j = 0 To UBound(CopyMap, 2)
            Map(h, (j + Select_SX) And MapSize, (i + Select_SY) And MapSize) = CopyMap(h, j, i)
        Next j
    Next i
    Next h
    
    Select_EX = Select_SX + UBound(CopyMap, 2)
    Select_EY = Select_SY + UBound(CopyMap, 3)
    
    MapShow
    
MapPastEnd:

End Sub
Public Sub eMapPast()
'�R�s�[����e�}�b�v�f�[�^��\��t����
On Error GoTo eMapPastEnd
    Dim i As Integer, j As Integer
    
    For i = 0 To UBound(eCopyMap, 2)
        For j = 0 To UBound(eCopyMap, 1)
            eMap((j + Select_SX) Mod (MapSize + 1), (i + Select_SY) Mod (MapSize + 1)) = eCopyMap(j, i)
        Next j
    Next i

    Select_EX = Select_SX + UBound(eCopyMap, 1)
    Select_EY = Select_SY + UBound(eCopyMap, 2)
    
    MapShow
    
eMapPastEnd:

End Sub
Public Sub Undo()
'�A���f�D�����s

    If ToolForm.Tool1.Buttons("Undo").Enabled Then
        ReDim RedoMap(0 To MapSize, 0 To MapSize)
        RedoMap = Map
        Map = UndoMap
        ToolForm.Tool1.Buttons("Redo").Enabled = True
        ToolForm.Tool1.Buttons("Undo").Enabled = False
        MapShow
    End If
    
End Sub
Public Sub Redo()
'���h�D�����s

    If ToolForm.Tool1.Buttons("Redo").Enabled Then
        ReDim UndoMap(0 To MapSize, 0 To MapSize)
        UndoMap = Map
        Map = RedoMap
        ToolForm.Tool1.Buttons("Redo").Enabled = False
        ToolForm.Tool1.Buttons("Undo").Enabled = True
        MapShow
    End If
    
End Sub
Public Sub UndoSet()
'�ύX�O�̃f�[�^��ۑ�����

    ReDim UndoMap(0 To MapSize, 0 To MapSize)
    UndoMap = Map
    ToolForm.Tool1.Buttons("Redo").Enabled = False
    ToolForm.Tool1.Buttons("Undo").Enabled = True

End Sub


Public Sub MapPaint(ByVal Num As Integer)
'�w�肳�ꂽ�`�b�v�ԍ��Ń}�b�v��h��ׂ�

    Dim i As Integer, j As Integer

    For i = 0 To MapSize
        For j = 0 To MapSize
            Map(ChipNow, i, j) = Num
        Next j
    Next i
    
    '�}�b�v�̍ĕ`��
    MapShow

End Sub
Public Sub ChipBarShow(Index As Integer)
'�l�h�c�t�H�[���̃`�b�v�p�s�N�`���{�b�N�X�Ƀ`�b�v���Ĕz�u�\������

    Dim i As Long, j As Long
    Dim HX As Long, HY As Long
    Dim tmp As Long
    
    MainForm.ShowChip(Index).Cls

        For j = 0 To 31
            For i = 0 To 7
                tmp = j * 8 + i
                HX = tmp Mod (Me.Chip(Index).Width / 32)
                HY = tmp \ (Me.Chip(Index).Width / 32)
                BitBlt MainForm.ShowChip(Index).hdc, i * 32, j * 32, 32, 32, Me.Chip(Index).hdc, HX * 32, HY * 32, SrcCopy
            Next i
        Next j
        
        '�`��ł��Ȃ��Ƃ����x���ׂ̈ɓ_�`�Ń`�b�v���Â�����
        If Index = 1 Then
            For i = 0 To 31 Step 2
            For j = 0 To 31
                Call SetPixelV(MainForm.ShowChip(Index).hdc, i + j Mod 2, j, RGB(0, 0, 0))
            Next
            Next
        End If

        MainForm.ShowChip(Index).Refresh

End Sub
Public Sub ToolChipShow()

Dim Tmp_Rect As RECT
Dim Tmp_subRect As RECT

    With Tmp_Rect
        .Top = 9
        .Bottom = 32 - 9
        .Left = 0
        .Right = 32
    End With
    With Tmp_subRect
        .Top = Tmp_Rect.Top + 1
        .Bottom = Tmp_Rect.Bottom + 1
        .Left = Tmp_Rect.Left + 1
        .Right = Tmp_Rect.Right + 1
    End With


    '�c�[���o�[�̑I���`�b�v��ύX����
    BitBlt ToolForm.LeftPic.hdc, 0, 0, 32, 32, MainForm.ShowChip(ChipNow).hdc, (LeftNo And 8 - 1) * 32, (LeftNo And (&HF8 - 8 * (a - 1))) * 4, SrcCopy
    BitBlt ToolForm.RightPic.hdc, 0, 0, 32, 32, MainForm.ShowChip(ChipNow).hdc, (RightNo And 8 - 1) * 32, (RightNo And (&HF8 - 8 * (a - 1))) * 4, SrcCopy
    
    ToolForm.LeftPic.ForeColor = RGB(0, 0, 0)
    ToolForm.RightPic.ForeColor = RGB(0, 0, 0)
    Call DrawText(ToolForm.LeftPic.hdc, CStr(LeftNo), -1, Tmp_subRect, DT_CENTER)
    Call DrawText(ToolForm.RightPic.hdc, CStr(RightNo), -1, Tmp_subRect, DT_CENTER)
    ToolForm.LeftPic.ForeColor = IIf(ChipNow = 0, RGB(255, 0, 255), RGB(0, 255, 255))
    ToolForm.RightPic.ForeColor = IIf(ChipNow = 0, RGB(255, 0, 255), RGB(0, 255, 255))
    Call DrawText(ToolForm.LeftPic.hdc, CStr(LeftNo), -1, Tmp_Rect, DT_CENTER)
    Call DrawText(ToolForm.RightPic.hdc, CStr(RightNo), -1, Tmp_Rect, DT_CENTER)
    
    ToolForm.LeftPic.Refresh
    ToolForm.RightPic.Refresh

End Sub

Private Sub X_Scroll_Change()
'�w�����̃X�N���[���o�[�̏���

    x = X_Scroll.Value
    MapShow
    
End Sub

Private Sub Y_Scroll_Change()
'�x�����̃X�N���[���o�[�̏���

    y = Y_Scroll.Value
    MapShow
    
End Sub
