VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form ToolForm 
   BorderStyle     =   4  '固定ﾂｰﾙ ｳｨﾝﾄﾞｳ
   Caption         =   "ToolBox"
   ClientHeight    =   2910
   ClientLeft      =   2190
   ClientTop       =   2580
   ClientWidth     =   1560
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   194
   ScaleMode       =   3  'ﾋﾟｸｾﾙ
   ScaleWidth      =   104
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   60
      TabIndex        =   0
      Top             =   1500
      Width           =   675
      Begin VB.PictureBox RightPic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   60
         ScaleHeight     =   32
         ScaleMode       =   3  'ﾋﾟｸｾﾙ
         ScaleWidth      =   32
         TabIndex        =   2
         Top             =   780
         Width           =   540
      End
      Begin VB.PictureBox LeftPic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   60
         ScaleHeight     =   32
         ScaleMode       =   3  'ﾋﾟｸｾﾙ
         ScaleWidth      =   32
         TabIndex        =   1
         Top             =   180
         Width           =   540
      End
   End
   Begin MSComctlLib.ImageList Image01 
      Left            =   900
      Top             =   2220
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ToolForm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ToolForm.frx":0118
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ToolForm.frx":0230
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ToolForm.frx":0344
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ToolForm.frx":0458
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ToolForm.frx":0570
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ToolForm.frx":0684
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ToolForm.frx":0798
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Tool1 
      Height          =   2040
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   3598
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "Image01"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cursor"
            Object.ToolTipText     =   "データの選択を行います"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pen"
            Object.ToolTipText     =   "選択されたマップチップを配置します"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paint"
            Object.ToolTipText     =   "選択されたマップチップで塗り潰します"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Syringe"
            Object.ToolTipText     =   "マップ上のチップを抽出します"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "選択されたデータをコピーします"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "コピーしたデータを貼り付けます"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "編集を一つ前の状態にします"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Object.ToolTipText     =   "アンドゥでやり直した作業を再度実行します"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ToolForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'マルチマップエディター

Option Explicit

Private Sub Form_Activate()
    If Not MainForm.ActiveForm Is Nothing Then
        MainForm.ActiveForm.Crt.ToolTipText = ""
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Call MainForm.ActiveForm.KeyToKey(KeyCode, Shift)
End Sub

Private Sub Form_Load()
    
    Me.Left = MainForm.Left
    Me.Top = MainForm.Top + Screen.TwipsPerPixelY * 100
    Me.Width = 59 * Screen.TwipsPerPixelX

End Sub

Public Sub Tool1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    '選択されたツール情報を編集中のフォームへ伝える
    Select Case Button.KEY
        Case "Cursor"
            With MainForm.ActiveForm
                .Tool = Button.KEY
                Tool1.Buttons("Copy").Enabled = True
                .Select_SX = .x
                .Select_SY = .y
                .Select_EX = .Select_SX
                .Select_EY = .Select_SY
                .SelectShow
            End With
        
        Case "Pen"
            MainForm.ActiveForm.Tool = Button.KEY
            Tool1.Buttons("Paste").Enabled = False
            Tool1.Buttons("Copy").Enabled = False
        Case "Paint"
            MainForm.ActiveForm.Tool = Button.KEY
            Tool1.Buttons("Paste").Enabled = False
            Tool1.Buttons("Copy").Enabled = False
        Case "Syringe"
            MainForm.ActiveForm.Tool = Button.KEY
            Tool1.Buttons("Paste").Enabled = False
            Tool1.Buttons("Copy").Enabled = False
        Case "Copy"
        
            With MainForm.ActiveForm
                .MapCopy
                .MapShow
            End With
            Tool1.Buttons("Paste").Enabled = True
            Exit Sub
        Case "Paste"
            If (eCopy) Then
                MainForm.ActiveForm.eDataChanged = True
                MainForm.ActiveForm.eMapPast
            Else
                MainForm.ActiveForm.DataChanged = True
                MainForm.ActiveForm.UndoSet
                MainForm.ActiveForm.MapPast
            End If
            Exit Sub
        Case "Undo"
            MainForm.ActiveForm.Undo
            Exit Sub
        Case "Redo"
            MainForm.ActiveForm.Redo
            Exit Sub
    End Select
    
    'ツールに依存する変数をクリアする
        MainForm.ActiveForm.LeftDraw = 0
        MainForm.ActiveForm.RightDraw = 0
        MainForm.ActiveForm.MapShow

End Sub

