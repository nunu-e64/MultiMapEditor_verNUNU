VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MainForm 
   BackColor       =   &H8000000C&
   Caption         =   "MultiMapEditor_NUNU"
   ClientHeight    =   7320
   ClientLeft      =   3825
   ClientTop       =   2910
   ClientWidth     =   11145
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  '�蓮
   Begin MSComctlLib.Toolbar Top_bar 
      Align           =   1  '�㑵��
      Height          =   705
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   1244
      ButtonWidth     =   1984
      ButtonHeight    =   1085
      Appearance      =   1
      ImageList       =   "Image1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "NewMap"
            Key             =   "New"
            Object.ToolTipText     =   "�V�����ҏW�f�[�^���쐬����"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ChipOpen"
            Key             =   "Chip"
            Object.ToolTipText     =   "�}�b�v�`�b�v�i���w�j��I������"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Chip2Open"
            Key             =   "Chip2"
            Object.ToolTipText     =   "�}�b�v�`�b�v�i��w�j��I������"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "MapOpen"
            Key             =   "Map"
            Object.ToolTipText     =   "�ҏW���̃}�b�v�������}�b�v�ƌ�������"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "MapSave"
            Key             =   "Save"
            Object.ToolTipText     =   "���O�����ă}�b�v��ۑ�"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "EventSave"
            Key             =   "eSave"
            Object.ToolTipText     =   "���O�����ăC�x���g�}�b�v��ۑ�"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ChipChange"
            Key             =   "ChipChange"
            Object.ToolTipText     =   "�\���`�b�v�Z�b�g�̏�w�Ɖ��w��؂肩����"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Image1 
      Left            =   840
      Top             =   900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":0DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MainForm.frx":0F0A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1500
      Top             =   1140
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1440
      Top             =   660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  '������
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   7005
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13996
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2014/02/28"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "2:54"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Left_Bar 
      Align           =   4  '�E����
      Height          =   6300
      Left            =   8955
      ScaleHeight     =   416
      ScaleMode       =   3  '�߸��
      ScaleWidth      =   142
      TabIndex        =   0
      Top             =   705
      Width           =   2190
      Begin VB.PictureBox ShowChip 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000007&
         BorderStyle     =   0  '�Ȃ�
         Height          =   435
         Index           =   2
         Left            =   1440
         ScaleHeight     =   29
         ScaleMode       =   3  '�߸��
         ScaleWidth      =   29
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox ShowChip 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000007&
         BorderStyle     =   0  '�Ȃ�
         Height          =   435
         Index           =   1
         Left            =   840
         ScaleHeight     =   29
         ScaleMode       =   3  '�߸��
         ScaleWidth      =   29
         TabIndex        =   5
         Top             =   60
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox ShowChip 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000007&
         BorderStyle     =   0  '�Ȃ�
         Height          =   435
         Index           =   0
         Left            =   240
         ScaleHeight     =   29
         ScaleMode       =   3  '�߸��
         ScaleWidth      =   29
         TabIndex        =   4
         Top             =   60
         Width           =   435
      End
      Begin VB.VScrollBar ChipBar 
         Height          =   615
         Left            =   0
         TabIndex        =   3
         Top             =   60
         Width           =   255
      End
   End
   Begin VB.Menu Menu000 
      Caption         =   "�t�@�C��(&F)"
      Begin VB.Menu Menu001 
         Caption         =   "�V�����}�b�v���쐬����"
         Shortcut        =   ^N
      End
      Begin VB.Menu Menu002 
         Caption         =   "�����}�b�v���J��"
         Shortcut        =   ^O
      End
      Begin VB.Menu Menu003 
         Caption         =   "-"
      End
      Begin VB.Menu Menu004 
         Caption         =   "�}�b�v�`�b�v�i���w�j�̓Ǎ���"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu Menu004 
         Caption         =   "�}�b�v�`�b�v�i��w�j�̓Ǎ���"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu Menu005 
         Caption         =   "�}�b�v�f�[�^�̓Ǎ���"
         Enabled         =   0   'False
      End
      Begin VB.Menu Menu006 
         Caption         =   "-"
      End
      Begin VB.Menu Menu007 
         Caption         =   "���O��t���ă}�b�v��ۑ�"
         Enabled         =   0   'False
      End
      Begin VB.Menu MenueSave 
         Caption         =   "���O��t���ăC�x���g�}�b�v�̕ۑ�"
         Enabled         =   0   'False
      End
      Begin VB.Menu MenuOverWrite 
         Caption         =   "�}�b�v���C�x���g�̏㏑���ۑ�"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu blank0 
         Caption         =   "-"
      End
      Begin VB.Menu OpenFolder 
         Caption         =   "MapEditor�̏ꏊ���J��"
      End
   End
   Begin VB.Menu Menu100 
      Caption         =   "�}�b�v�T�C�Y"
      Begin VB.Menu Menu101 
         Caption         =   "256�~256"
         Index           =   1
      End
      Begin VB.Menu Menu101 
         Caption         =   "128�~128"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu Menu101 
         Caption         =   "64�~64"
         Index           =   3
      End
      Begin VB.Menu Menu101 
         Caption         =   "�蓮�ݒ�"
         Index           =   4
      End
   End
   Begin VB.Menu Menu200 
      Caption         =   "�E�B���h�E(&W)"
      WindowList      =   -1  'True
      Begin VB.Menu Menu201 
         Caption         =   "�d�˂ĕ\��"
      End
      Begin VB.Menu Menu202 
         Caption         =   "���ׂĕ\��"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'�}���`�}�b�v�G�f�B�^�[

Option Explicit

'���݊J���Ă���`���C���h�E�B���h�E�̐�
Public FormCounter As Long

Private Sub ChipBar_Change()
    
    ShowChip(0).Top = ChipBar.Value * 32 * -1
    ShowChip(1).Top = ChipBar.Value * 32 * -1
    
End Sub

Private Sub Left_Bar_Resize()
    
    ChipReSize
    
End Sub


Private Sub MDIForm_Load()
'�l�c�h�t�H�[�����[�h�C�x���g

    '�e�R���g���[���̏�������
    ChipReSize
    '�R�����_�C�A���O�̃L�����Z�����ɃG���[�Ƃ���
    CommonDialog1.CancelError = True
    
    '�����\���ʒu����ʏ�̃Z���^�[�ֈړ�
    If WindowState = 0 Then
        Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    End If

    'Me.Caption = Me.Caption & Space(2) & "ver1.00"

    '�c�[���o�[��\������
    Load ToolForm
    ToolForm.Show , Me
    
    '�c�[���o�[�̈ꕔ�𖳌��ɂ���
    With Top_bar
        .Buttons("Chip").Enabled = False
        .Buttons("Chip2").Enabled = False
        .Buttons("Map").Enabled = False
        .Buttons("Save").Enabled = False
        .Buttons("eSave").Enabled = False
        .Buttons("ChipChange").Enabled = False
    End With
    MenuFalse
    
    
    If Command <> "" Then
        Dim Comstr As String
        Dim Tmp_Name As String
        Comstr = Command
        
        If Left(Comstr, 1) = """" And Right(Comstr, 1) = """" Then
            Comstr = Mid(Comstr, 2, Len(Comstr) - 2)
        End If
                
        If Right(Comstr, 5) = ".eMap" Then Comstr = Left(Comstr, Len(Comstr) - 5) + ".Map2"
            
        If Right(Comstr, 5) = ".Map2" Then
            Call Menu001_Click
            ActiveForm.MapLoad Comstr
            Tmp_Name = Comstr
            Tmp_Name = Left(Tmp_Name, Len(Tmp_Name) - 5) + ".eMap"
            Call ActiveForm.MapLoad(Tmp_Name, True)
        End If
        
    End If
    
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Comstr As String
    Dim Tmp_Name As String
    Dim i As Long
    On Error GoTo DragExit
    
    For i = 1 To Data.Files.Count
        Comstr = Data.Files(i)
            
        If Right(Comstr, 5) = ".eMap" Then Comstr = Left(Comstr, Len(Comstr) - 5) + ".Map2"
            
        If Right(Comstr, 5) = ".Map2" Then
            Call Menu001_Click
            ActiveForm.MapLoad Comstr
            Tmp_Name = Comstr
            Tmp_Name = Left(Tmp_Name, Len(Tmp_Name) - 5) + ".eMap"
            Call ActiveForm.MapLoad(Tmp_Name, True)
        End If
    Next
    
DragExit:
End Sub
Private Sub ChipReSize()
'�t�H�[����̃R���g���[���̃T�C�Y�ύX

    Left_Bar.Width = (20 + 32 * 8 * a) * Screen.TwipsPerPixelX
    
    '�`�b�v�p�̃X�N���[���o�[�̃T�C�Y�ύX
    With ChipBar
        .Top = 0
        .Left = 0
        .Height = Left_Bar.ScaleHeight
        .Width = 16
        .Max = ((32 * 64) - Left_Bar.ScaleHeight) \ 32
    End With
    
    '�`�b�v�p�̃s�N�`���{�b�N�X�̃T�C�Y�ύX
    With ShowChip(0)
        .Top = 0
        .Left = ChipBar.Width
        .Width = 32 * 8
        .Height = 32 * 32
    End With
    With ShowChip(1)
        .Top = 0
        .Left = ChipBar.Width
        .Width = 32 * 8
        .Height = 32 * 32
    End With
    With ShowChip(2)
        .Top = 0
        .Left = ChipBar.Width
        .Width = 32 * 8
        .Height = 32 * 32
    End With
    
End Sub
Public Sub MenuTrue()
'�ҏW���̃t�H�[���������Ǝ����o���Ȃ��R���g���[���̗L����

    '���j���[���̗L����
    Menu004(0).Enabled = True
    Menu004(1).Enabled = True
    Menu005.Enabled = True
    Menu007.Enabled = True
    Menu100.Enabled = True
    Menu200.Enabled = True
    MenuOverWrite.Enabled = True
    MenueSave.Enabled = True
    
    '�c�[���{�b�N�X�̗L����
    With ToolForm.Tool1
        .Buttons("Cursor").Enabled = True
        .Buttons("Pen").Enabled = True
        .Buttons("Syringe").Enabled = True
        .Buttons("Paint").Enabled = True
    End With
    
End Sub
Public Sub MenuFalse()
'�ҏW���̃t�H�[���������Ǝ����o���Ȃ��R���g���[���̖���
      
    '���j���[���̖�����
    Menu004(0).Enabled = False
    Menu004(1).Enabled = False
    Menu005.Enabled = False
    Menu007.Enabled = False
    Menu100.Enabled = False
    Menu200.Enabled = False
    MenuOverWrite.Enabled = False
    MenueSave.Enabled = False

    '�c�[���{�b�N�X�̖�����
    With ToolForm.Tool1
        .Buttons("Cursor").Enabled = False
        .Buttons("Pen").Enabled = False
        .Buttons("Syringe").Enabled = False
        .Buttons("Paint").Enabled = False
    
        .Buttons("Copy").Enabled = False
        .Buttons("Past").Enabled = False
    
        .Buttons("Undo").Enabled = False
        .Buttons("Redo").Enabled = False
    
        '�I���c�[���̏�����
        .Buttons("Cursor").Value = 0
        .Buttons("Pen").Value = 0
        .Buttons("Syringe").Value = 0
        .Buttons("Paint").Value = 0
    End With

End Sub
Public Function MapSave(Optional OverWrite As Boolean = False) As Boolean
'�}�b�v�t�@�C���̕ۑ�
    
    On Error Resume Next    '���̃��[�`�����̃G���[�𖳌��ɂ���B
    
    If OverWrite And ActiveForm.SaveFileName <> "" Then
        ActiveForm.MapSave (ActiveForm.SaveFileName)
        MapSave = True
    Else
        
        If ActiveForm.SaveFileName = "" Then
            CommonDialog1.FileName = ""
        Else
            CommonDialog1.FileName = ActiveForm.SaveFileName
        End If
        
        With CommonDialog1
            .DialogTitle = "���O��t���ăt�@�C���̕ۑ�"
            .Filter = "Pictures(*.Map2)|*.Map2"
            .Flags = &H2
            .ShowSave   '���O��t���ĕۑ��p���޲�۸ނ��J��
        End With
        
        DoEvents
            
        If Err <> cdlCancel Then    ' հ�ް��[��ݾ�]��I�����܂����B 32755=��ݾٺ���
            ActiveForm.MapSave CommonDialog1.FileName
            MapSave = True
        Else
            MapSave = False
        End If
    End If
    
End Function
Public Function eMapSave(Optional OverWrite As Boolean = False) As Boolean
'�}�b�v�t�@�C���̕ۑ�
    
    On Error Resume Next    '���̃��[�`�����̃G���[�𖳌��ɂ���B
    
    If OverWrite And ActiveForm.eSaveFileName <> "" Then
        Call ActiveForm.MapSave(ActiveForm.eSaveFileName, True)
        eMapSave = True
    Else
        
        If ActiveForm.eSaveFileName = "" Then
            CommonDialog1.FileName = Left(ActiveForm.SaveFileName, Len(ActiveForm.SaveFileName) - 5) + ".eMap"
        Else
            CommonDialog1.FileName = ActiveForm.eSaveFileName
        End If
        
        With CommonDialog1
            .DialogTitle = "���O��t���ăC�x���g�}�b�v�̕ۑ�"
            .Filter = "Pictures(*.eMap)|*.eMap"
            .Flags = &H2
            .ShowSave   '���O��t���ĕۑ��p���޲�۸ނ��J��
        End With
        
        DoEvents
            
        If Err <> cdlCancel Then    ' հ�ް��[��ݾ�]��I�����܂����B 32755=��ݾٺ���
            Call ActiveForm.MapSave(CommonDialog1.FileName, True)
            eMapSave = True
        Else
            eMapSave = False
        End If
    End If
    
End Function


Private Sub Menu001_Click()
'�V�����`���C���h�E�B���h�E���J��

    FormCounter = FormCounter + 1
    Dim MapForm As New Map
    MapForm.Tag = FormCounter
    MapForm.Show

End Sub

Private Sub Menu002_Click()
'�V�����t�H�[�����J���ă}�b�v��Ǎ���
    Dim Tmp_Name As String
    
    On Error Resume Next    '���̃��[�`�����̃G���[�𖳌��ɂ���B
    With CommonDialog1
        .DialogTitle = "�}�b�v�f�[�^�̓ǂݍ���"
        .FileName = ""
        .Filter = "Map2�t�@�C��(*.Map2)|*.Map2"
        .ShowOpen     '̧�ٵ���ݗp���޲�۸ނ��J��
    End With
    
    DoEvents
        
    If Err <> cdlCancel Then    ' հ�ް��[��ݾ�]��I�����܂����B 32755=��ݾٺ���

        FormCounter = FormCounter + 1
        Dim MapForm As New Map
        MapForm.Tag = FormCounter
        ActiveForm.MapLoad CommonDialog1.FileName
        Tmp_Name = CommonDialog1.FileName
        
        '�����̃}�b�v���J�����ꍇ�A���������}�b�v�`�b�v�̑I�����s��
        Menu004_Click 0
        Menu004_Click 1
        
        Tmp_Name = Left(Tmp_Name, Len(Tmp_Name) - 5) + ".eMap"
        Call ActiveForm.MapLoad(Tmp_Name, True)
        
    End If
    
    


End Sub

Public Sub Menu004_Click(Index As Integer)
'�}�b�v�`�b�v�̓ǂݍ���

    On Error Resume Next    '���̃��[�`�����̃G���[�𖳌��ɂ���B

    With CommonDialog1
        .DialogTitle = IIf(Index = 0, "���w�}�b�v�`�b�v�̓Ǎ���", "��w�}�b�v�`�b�v�̓Ǎ���")
        .FileName = ""
        .Filter = "Pictures(*.bmp;*.gif)|*.bmp;*.gif"
        .ShowOpen   '̧�ٵ���ݗp���޲�۸ނ��J��
    End With
    
    DoEvents
        
    If Err <> cdlCancel Then    ' հ�ް��[��ݾ�]��I�����܂����B 32755=��ݾٺ���
        Call ActiveForm.ChipLoad(CommonDialog1.FileName, Index)
        
        If Index = 1 Then
            With CommonDialog1
                .DialogTitle = "��w�}�b�v�`�b�v�̃}�X�N�̓Ǎ���"
                .FileName = ""
                .Filter = "Pictures(*.bmp;*.gif)|*.bmp;*.gif"
                .ShowOpen   '̧�ٵ���ݗp���޲�۸ނ��J��
            End With
            If Err <> cdlCancel Then    ' հ�ް��[��ݾ�]��I�����܂����B 32755=��ݾٺ���
                Call ActiveForm.ChipLoad(CommonDialog1.FileName, 2)
            Else
                Call ActiveForm.MapShow
            End If
        End If
    End If
    
End Sub
Public Sub Menu005_Click()
'�}�b�v�t�@�C����I�����ēǍ���
Dim Tmp_Name As String

    On Error Resume Next    '���̃��[�`�����̃G���[�𖳌��ɂ���B
    With CommonDialog1
        .DialogTitle = "�}�b�v�f�[�^�̓ǂݍ���"
        .FileName = ""
        .Filter = "Pictures(*.Map2)|*.Map2"
        .ShowOpen   '̧�ٵ���ݗp���޲�۸ނ��J��
    End With
    
    DoEvents
        
    If Err <> cdlCancel Then    ' հ�ް��[��ݾ�]��I�����܂����B 32755=��ݾٺ���
        ActiveForm.MapLoad CommonDialog1.FileName
        Tmp_Name = CommonDialog1.FileName
        Tmp_Name = Left(Tmp_Name, Len(Tmp_Name) - 5) + ".eMap"
        Call ActiveForm.MapLoad(Tmp_Name, True)
    End If

End Sub
Public Sub Menu007_Click()
'�}�b�v�t�@�C���̕ۑ�
    
    Dim Ret As Boolean
    Ret = MapSave

End Sub

Private Sub MenueSave_Click()
    eMapSave
End Sub

Private Sub MenuOverWrite_Click()
    MapSave (True)
    eMapSave (True)
End Sub


Private Sub Menu101_Click(Index As Integer)
'�}�b�v�T�C�Y�̕ύX���j���[

    Dim i As Long
    Dim Tmp_Size As Long
    Dim Tmp_strSize As String
    
    Tmp_Size = ActiveForm.MapSize + 1
    Select Case Index
    Case 1
        Tmp_Size = 256
    Case 2
         Tmp_Size = 128
    Case 3
        Tmp_Size = 64
    Case Else
        Tmp_strSize = InputBox("�}�b�v�T�C�Y����͂��Ă��������i16~256�j", "�}�b�v�T�C�Y�̕ύX ", ActiveForm.MapSize + 1)
        If Tmp_strSize <> "" Then Tmp_Size = CLng(Tmp_strSize)
        
    End Select
    
        
    '��Ԃ̕ύX�����邩�ǂ������`�F�b�N
    If ActiveForm.MapSize = Tmp_Size - 1 Then Exit Sub
    
    If Tmp_Size < 16 Or Tmp_Size > 256 Then
        Call MsgBox("�}�b�v�T�C�Y��16�ȏ�256�ȉ��Ŏw�肵�Ă�������", vbOKOnly, "�G���[�F�}�b�v�T�C�Y�̕ύX")
        Exit Sub
    End If
    
    '���b�Z�[�W�{�b�N�X�ɂĊm�F��\��
    If MsgBox("�}�b�v�T�C�Y�̕ύX���s���܂����H�F" & ActiveForm.MapSize + 1 & "*" & ActiveForm.MapSize + 1 & "��" & Tmp_Size & "*" & Tmp_Size, vbOKCancel, "�}�b�v�T�C�Y�̕ύX") <> 1 Then
        Exit Sub
    End If
    
    '���ׂẴ`�F�b�N���\���ɂ���
    For i = 1 To 4
        Menu101(i).Checked = False
    Next i
    '�I�����ꂽ���j���[�̃`�F�b�N��\���ɂ���
    Menu101(Index).Checked = True
    
    '�I�����ꂽ���j���[�ɏ]���ă}�b�v�T�C�Y�̕ύX���s��
    ActiveForm.ChangeMapSize Tmp_Size
    
    
End Sub
Private Sub Menu201_Click()
'���݂̃E�B���h�E���d�˂Đ���

    Arrange vbCascade

End Sub

Private Sub Menu202_Click()
'���݂̃E�B���h�E����ׂĐ���
    
    Arrange vbTileVertical

End Sub


Private Sub OpenFolder_Click()
    
    Shell "rundll32.exe url.dll,FileProtocolHandler " & App.Path, vbNormalFocus     '140222nunu �Ȃ�����̕��@�ł�exe�������イ�܂��@�\���Ă���Ȃ������̂�
    
End Sub

Private Sub ShowChip_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'�}�b�v�`�b�v�̑I��

    '�ҏW���̃t�H�[�������邩�ǂ������`�F�b�N
    If FormCounter <> 0 Then
    
        '���{�^���������ꂽ�ꍇ�̏���
        If Button = 1 Then
            Me.ActiveForm.LeftNo = (X \ 32) + ((Y \ 32) * 8 * a)
        End If
        '�E�{�^���������ꂽ�ꍇ�̏���
        If Button = 2 Then
            Me.ActiveForm.RightNo = (X \ 32) + ((Y \ 32) * 8 * a)
        End If
        Me.ActiveForm.ToolChipShow
        
        If ToolForm.Tool1.Buttons("Pen").Value = tbrUnpressed And ToolForm.Tool1.Buttons("Paint").Value = tbrUnpressed Then
            ToolForm.Tool1.Buttons("Pen").Value = tbrPressed
            MainForm.ActiveForm.Tool = "Pen"
            ToolForm.Tool1.Buttons("Past").Enabled = False
            ToolForm.Tool1.Buttons("Copy").Enabled = False
            MainForm.ActiveForm.LeftDraw = 0
            MainForm.ActiveForm.RightDraw = 0
            MainForm.ActiveForm.MapShow
        End If
    
    End If

End Sub

Private Sub Timer1_Timer()
'�^�C�}�[���荞�݂ɂă}�b�v�̃X�N���[�����s��

    If 0 And FormCounter <> 0 And WindowState <> 1 Then
        
        '�E�L�[�̏���
        If GetAsyncKeyState(vbKeyRight) Then
            Me.ActiveForm.X_Scroll.Value = (Me.ActiveForm.X_Scroll.Value + 1 + Me.ActiveForm.MapSize + 1) Mod (Me.ActiveForm.MapSize + 1)
        End If
        
        '���L�[�̏���
        If GetAsyncKeyState(vbKeyLeft) Then
            Me.ActiveForm.X_Scroll.Value = (Me.ActiveForm.X_Scroll.Value - 1 + Me.ActiveForm.MapSize + 1) Mod (Me.ActiveForm.MapSize + 1)
        End If
        
        '��L�[�̏���
        If GetAsyncKeyState(vbKeyUp) Then
            Me.ActiveForm.Y_Scroll.Value = (Me.ActiveForm.Y_Scroll.Value - 1 + Me.ActiveForm.MapSize + 1) Mod (Me.ActiveForm.MapSize + 1)
        End If
        
        '���L�[�̏���
        If GetAsyncKeyState(vbKeyDown) Then
            Me.ActiveForm.Y_Scroll.Value = (Me.ActiveForm.Y_Scroll.Value + 1 + Me.ActiveForm.MapSize + 1) Mod (Me.ActiveForm.MapSize + 1)
        End If
            
        If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyZ) Then
            MainForm.ActiveForm.Undo
            
        ElseIf GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyY) Then
            MainForm.ActiveForm.Redo
        
        ElseIf GetAsyncKeyState(vbKeyDelete) Then
            MainForm.ActiveForm.MapDelete
            
        End If
        
    End If

End Sub

Private Sub Top_bar_ButtonClick(ByVal Button As MSComctlLib.Button)
'�c�[���o�[�̏���

    Select Case Button.KEY
    
        Case "New"
            Call Menu001_Click
        Case "Chip"
            Call Menu004_Click(0)
        Case "Chip2"
            Call Menu004_Click(1)
        Case "Map"
            Call Menu005_Click
        Case "Save"
            Call Menu007_Click
        Case "eSave"
            Call MenueSave_Click
        Case "ChipChange"
            ShowChip(0).Visible = Not (ShowChip(0).Visible)
            ShowChip(1).Visible = Not (ShowChip(1).Visible)
            ChipNow = (ChipNow + 1) And 1
            ActiveForm.ToolChipShow
    End Select

End Sub
