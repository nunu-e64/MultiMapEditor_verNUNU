VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows の既定値
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   6360
      TabIndex        =   6
      Text            =   "***に"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   5520
      TabIndex        =   5
      Text            =   "***を"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "変更"
      Height          =   375
      Left            =   7320
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   4680
      TabIndex        =   3
      Text            =   "***層の"
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Map2Open"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Map2Save"
      Height          =   495
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MapOpen Change"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Map(0 To 255, 0 To 255) As Byte
Dim Map2(1, 0 To 255, 0 To 255) As Byte

Private Sub Command1_Click(i As Integer)
Dim OpenFile As String

    On Error Resume Next    'このルーチン内のエラーを無効にする。
        
    If i = 0 Then
    
        With CommonDialog1
            .DialogTitle = "マップデータの読み込み"
            .FileName = ""
            .Filter = "Pictures(*.Map)|*.Map"
            .ShowOpen   'ﾌｧｲﾙｵｰﾌﾟﾝ用のﾀﾞｲｱﾛｸﾞを開く
        End With
        
        DoEvents
            
        If (Err <> cdlCancel) And (CommonDialog1.FileName <> "") Then ' ﾕｰｻﾞｰが[ｷｬﾝｾﾙ]を選択しました。 32755=ｷｬﾝｾﾙｺｰﾄﾞ
            MapLoad CommonDialog1.FileName
            MapMove
        End If
        
    Else
        
        With CommonDialog1
            .DialogTitle = "名前を付けてファイルの保存"
            .Filter = "Pictures(*.Map2)|*.Map2"
            .FileName = "*.Map2"
            .Flags = &H2
            .ShowSave   '名前を付けて保存用のﾀﾞｲｱﾛｸﾞを開く
        End With
        
        DoEvents
            
        If (Err <> cdlCancel) And (CommonDialog1.FileName <> "") Then   ' ﾕｰｻﾞｰが[ｷｬﾝｾﾙ]を選択しました。 32755=ｷｬﾝｾﾙｺｰﾄﾞ
            MapSave2 CommonDialog1.FileName
        End If
    
    End If
End Sub

Private Sub Command2_Click()
    
        With CommonDialog1
            .DialogTitle = "マップデータの読み込み"
            .FileName = ""
            .Filter = "Pictures(*.Map2)|*.Map2"
            .ShowOpen   'ﾌｧｲﾙｵｰﾌﾟﾝ用のﾀﾞｲｱﾛｸﾞを開く
        End With
        
        DoEvents
            
        If (Err <> cdlCancel) And (CommonDialog1.FileName <> "") Then ' ﾕｰｻﾞｰが[ｷｬﾝｾﾙ]を選択しました。 32755=ｷｬﾝｾﾙｺｰﾄﾞ
            MapLoad2 CommonDialog1.FileName
        End If
        
End Sub

Private Sub MapLoad(FileName As String)

    Open FileName For Binary Access Read As 1
        Get #1, , Map
    Close #1
    Me.Caption = FileName

End Sub
Private Sub MapLoad2(FileName As String)

    Open FileName For Binary Access Read As 1
        Get #1, , Map2
    Close #1
    Me.Caption = FileName

End Sub

Private Sub MapSave2(FileName As String)

    Open FileName For Binary Access Write As 1
        Put #1, , Map2
    Close #1
    Me.Caption = FileName

End Sub
Private Sub MapMove()
Dim i As Long, j As Long

    For i = 0 To UBound(Map2, 2)
        For j = 0 To UBound(Map2, 3)
            Map2(0, i, j) = Map(i, j)
        Next
    Next
End Sub

Private Sub Command3_Click()
Dim i As Long, j As Long
    Dim ChangeCount As Long
    
    For i = 0 To UBound(Map2, 2)
        For j = 0 To UBound(Map2, 3)
            If Map2(CLng(Text1(0).Text), i, j) = CLng(Text1(1).Text) Then
                Map2(CLng(Text1(0).Text), i, j) = CLng(Text1(2).Text)
                ChangeCount = ChangeCount + 1
            End If
        Next
    Next
    
    call msgbox(cstr(changecount) & "個をおきかえました"
    
End Sub

