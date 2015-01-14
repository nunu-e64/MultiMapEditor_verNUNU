VERSION 5.00
Begin VB.Form Map 
   AutoRedraw      =   -1  'True
   Caption         =   "Map [Sample.map] X:00 Y:00"
   ClientHeight    =   2610
   ClientLeft      =   6180
   ClientTop       =   4365
   ClientWidth     =   2760
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
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
   ScaleMode       =   3  'ﾋﾟｸｾﾙ
   ScaleWidth      =   184
   Begin VB.PictureBox SelectPic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'なし
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      ScaleMode       =   0  'ﾕｰｻﾞｰ
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
      BorderStyle     =   0  'なし
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      ScaleMode       =   3  'ﾋﾟｸｾﾙ
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
      BorderStyle     =   0  'なし
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      ScaleMode       =   3  'ﾋﾟｸｾﾙ
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
      BorderStyle     =   0  'なし
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      ScaleMode       =   0  'ﾕｰｻﾞｰ
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
      BorderStyle     =   0  'なし
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      ScaleMode       =   3  'ﾋﾟｸｾﾙ
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
      BorderStyle     =   0  'なし
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
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
      ScaleMode       =   3  'ﾋﾟｸｾﾙ
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
'マルチマップエディター
Option Explicit

'編集中のマップファイル名をタイトルとして格納する
Private Title As String

'表示しているマップ座標
Public x As Long, y As Long

'データ範囲選択時のマップ座標
Public Select_SX As Integer, Select_SY As Integer
Public Select_EX As Integer, Select_EY As Integer

'ダブルクリック判定用マップ座標
Private MouseX As Single, MouseY As Single

'マップを格納する変数
Private Map() As Byte
Private RedoMap() As Byte
Private UndoMap() As Byte
Private eMap() As Byte
Private OldMap() As Byte
Private OldeMap() As Byte

Public MapSize As Long      'マップの一辺のサイズ

'編集中のマップ名保存用変数
Public SaveFileName As String
Public eSaveFileName As String
Public OpenTime As String
Public eOpenTime As String

'チップ選択番号
Public LeftNo As Long
Public RightNo As Long
Public LeftDraw As Long
Public RightDraw As Long

'ツールの選択状態
Public Tool As String

'編集中のデータの状態（True 変更：False 未変更）
Public DataChanged As Boolean
Public eDataChanged As Boolean

'INIファイル用
Private MyKey As String
Private eMyKey As String


Private Sub Crt_DblClick()
    Call Crt_MouseDown(1, 1, MouseX, MouseY)
End Sub

Private Sub Crt_MouseDown(Button As Integer, Shift As Integer, MX As Single, MY As Single)
'マップの置き換え
        
    Dim Ret As Integer
    On Error Resume Next
    
    MouseX = MX
    MouseY = MY
    
    If Shift = 0 Then
        Select Case Tool
        
            Case "Pen"
                'チップ配置処理
    
                '左ボタンの処理
                If Button = 1 Then
                    If Map(ChipNow, (x + (MX \ 32)) And MapSize, (y + (MY \ 32)) And MapSize) <> LeftNo Then
                        DataChanged = True
                        UndoSet
                    End If
                    
                    Map(ChipNow, (x + (MX \ 32)) And MapSize, (y + (MY \ 32)) And MapSize) = LeftNo
                    LeftDraw = 1
                    RightDraw = 0
                End If
                
                '右ボタンの処理
                If Button = 2 Then
                    If Map(ChipNow, (x + (MX \ 32)) And MapSize, (y + (MY \ 32)) And MapSize) <> RightNo Then
                        DataChanged = True
                        UndoSet
                    End If
                    
                    Map(ChipNow, (x + (MX \ 32)) And MapSize, (y + (MY \ 32)) And MapSize) = RightNo
                    LeftDraw = 0
                    RightDraw = 1
                End If
                'マップを再描画
                MapShow
                
            Case "Syringe"
                'スポイト処理
                
                '左ボタンの処理
                If Button = 1 Then
                    LeftNo = Map(ChipNow, (x + (MX \ 32)) And MapSize, (y + (MY \ 32)) And MapSize)
                End If
                '右ボタンの処理
                If Button = 2 Then
                    RightNo = Map(ChipNow, (x + (MX \ 32)) And MapSize, (y + (MY \ 32)) And MapSize)
                End If
                '吸い出した番号で表示の変更
                ToolChipShow
                
            Case "Cursor"
                'データの選択処理
                
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
                    
                    'ToolTipTextクリックした場所のマップデータ番号
                    Me.Crt.ToolTipText = Map(0, x + MX \ 32, y + MY \ 32) & "-" & Map(1, x + MX \ 32, y + MY \ 32)

                
            Case "Paint"
                '塗り潰し処理
                Ret = MsgBox("選択されているチップで塗り潰します", vbOKCancel + vbQuestion, "MapEditor")
                If Ret = vbOK Then
                    UndoSet
                    DataChanged = True
                    MapPaint IIf(Button <> 2, LeftNo, RightNo)
                End If
        
        End Select
    
    Else
        On Error Resume Next
        Dim tmp As String
        
        tmp = InputBox("イベント番号を入力してください（0~255）", "イベント番号入力[X:" & (x + (MouseX \ 32)) Mod (MapSize + 1) & " Y:" & (y + (MouseY \ 32)) Mod (MapSize + 1) & "]", CStr(eMap((x + (MouseX \ 32)) Mod (MapSize + 1), (y + (MouseY \ 32)) Mod (MapSize + 1))))
        If tmp <> "" Then
            If CLng(tmp) < 0 Or CLng(tmp) > 255 Then
                Call MsgBox("イベント番号は0~255の数字で入力してください", vbOKOnly, "イベント番号入力[X:" & x + (MouseX \ 32) & " Y:" & y + (MouseY \ 32) & "]")
                Call Crt_DblClick
            ElseIf CLng(tmp) <> eMap((x + (MouseX \ 32)) Mod (MapSize + 1), (y + (MouseY \ 32)) Mod (MapSize + 1)) Then
                eMap((x + (MouseX \ 32)) Mod (MapSize + 1), (y + (MouseY \ 32)) Mod (MapSize + 1)) = CLng(tmp)
                'データの変更を記憶する
                eDataChanged = True
                MapShow
            End If
        End If
    
    End If
    
End Sub

Private Sub Crt_MouseMove(Button As Integer, Shift As Integer, MX As Single, MY As Single)
'マウスの移動時の処理
On Error Resume Next

    Select Case Tool
    
        Case "Pen"
            '連続データ配置処理
            If LeftDraw = 1 And (Crt.Width > MX And Crt.Height > MY) Then
                Map(ChipNow, (x + (MX \ 32)) And MapSize, (y + (MY \ 32)) And MapSize) = LeftNo
                MapShow
            End If
            If RightDraw = 1 And (Crt.Width > MX And Crt.Height > MY) Then
                Map(ChipNow, (x + (MX \ 32)) And MapSize, (y + (MY \ 32)) And MapSize) = RightNo
                MapShow
            End If
            
        Case "Cursor"
            '選択範囲拡大処理
            If (LeftDraw = 1 Or RightDraw = 1) And (Crt.Width > MX And Crt.Height > MY) Then
                Select_EX = x + MX \ 32
                Select_EY = y + MY \ 32
                MapShow
                'SelectShow
            
            ElseIf (Crt.Width > MX And Crt.Height > MY) Then
                'ToolTipTextクリックした場所のマップデータ番号
                Me.Crt.ToolTipText = "" '一度リセットすることで表示箇所をポインタ位置に
                Me.Crt.ToolTipText = Map(0, x + MX \ 32, y + MY \ 32) & "-" & Map(1, x + MX \ 32, y + MY \ 32)
            End If
                
    End Select
    

    
End Sub

Private Sub Crt_MouseUp(Button As Integer, Shift As Integer, MX As Single, MY As Single)
'ボタンが離された場合の処理

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

'アクティブになった時にＭＤＩフォームのチップを切りかえる
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
    '右キーの処理
    Case vbKeyRight, vbKeyD
        X_Scroll.Value = (X_Scroll.Value + 1 + MapSize + 1) Mod (MapSize + 1)
                                   
    '左キーの処理
    Case vbKeyLeft, vbKeyA
        X_Scroll.Value = (X_Scroll.Value - 1 + MapSize + 1) Mod (MapSize + 1)
        
    '上キーの処理
    Case vbKeyUp, vbKeyW
        Y_Scroll.Value = (Y_Scroll.Value - 1 + MapSize + 1) Mod (MapSize + 1)
                
    '下キーの処理
    Case vbKeyDown, vbKeyS
        Y_Scroll.Value = (Y_Scroll.Value + 1 + MapSize + 1) Mod (MapSize + 1)
                   
    '選択範囲の0化
    Case vbKeyDelete And Tool = "Cursor"
        MapDelete (Shift)
            
    Case vbKeyReturn And Tool = "Cursor"
        MapSelectSet
    
    'ツールの切り替え
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
    
    '検索
    Case vbKeyF And Shift = 2
        Call Search(False)
    Case vbKeyF And Shift = 3
        Call Search(True)
        
    '置換
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
        SearchNum = InputBox(IIf(ChipNow = 0, "下層", "上層") & "から検索したいデータ番号を入力してください(0~255)")
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
                
                MsgBox SearchNum & "番は" & CStr(Counting) & "個見つかりました"
            End If
        End If
    Else
        SearchNum = InputBox("イベントマップから検索したいデータ番号を入力してください(0~255)")
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
                
                MsgBox SearchNum & "番は" & CStr(Counting) & "個見つかりました"
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
        SearchNum = InputBox(IIf(ChipNow = 0, "下層", "上層") & "から置換したいデータ番号を入力してください(0~255)")
        If SearchNum <> "" Then
            If CInt(SearchNum) < 0 Or CInt(SearchNum) > 255 Then
                Call Replace(False)
                Exit Sub
            Else
                ReplaceNum = InputBox(IIf(ChipNow = 0, "下層の", "上層の") & SearchNum & "番を何番に置換したいのですか(0~255)")
                        
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
                        
                        MsgBox CStr(Counting) & "個の" & SearchNum & "が" & ReplaceNum & "に置換されました"
                        DataChanged = True
                        
                    End If
                End If
                
            End If
        End If
    
    Else
        
        SearchNum = InputBox("イベントマップから置換したいデータ番号を入力してください(0~255)")
        If SearchNum <> "" Then
            If CInt(SearchNum) < 0 Or CInt(SearchNum) > 255 Then
                Call Replace(True)
                Exit Sub
            Else
                ReplaceNum = InputBox("イベントマップ" & SearchNum & "番を何番に置換したいのですか(0~255)")
                        
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
                        
                        MsgBox CStr(Counting) & "個の" & SearchNum & "が" & ReplaceNum & "に置換されました"
                        eDataChanged = True
                        
                    End If
                End If
                
            End If
        End If
        
    End If

    MapShow
End Sub

Private Sub Form_Load()
'マップ配置用フォームのロードイベント
        
    'マップサイズの設定
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
    
    'マップ表示用のピクチャボックスの位置の初期化
    Crt.Top = 0
    Crt.Left = 0
    
    Chip(0).Width = 512
    Chip(0).Height = 512
    Chip(1).Width = 512
    Chip(1).Height = 512
    
    MapReSize
    
    x = 0: y = 0
    Title = "NewMap(NoName)"

    'ツールバーの一部を有効にする
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
'フォームサイズにピクチャボックスのサイズを合わせる

    On Error Resume Next    'このルーチン内のエラーを無効にする。

    'マップ表示用のピクチャボックスのサイズ調整
    Crt.Width = Me.ScaleWidth - 16
    Crt.Height = Me.ScaleHeight - 16
    
    'スクロールバーのサイズ調整
    Y_Scroll.Top = 0
    Y_Scroll.Left = Me.ScaleWidth - 16
    Y_Scroll.Height = Me.ScaleHeight - 16
    
    X_Scroll.Top = Me.ScaleHeight - 16
    X_Scroll.Left = 0
    X_Scroll.Width = Me.ScaleWidth - 16

    MapShow
    
    
End Sub

Private Sub Form_Resize()
'フォームの大きさを変更された場合の処理

    MapReSize
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'ウィンドウを閉じる時の処理

    Dim Ret As Integer
    
    If DataChanged = True Then
        
        Ret = MsgBox("編集中のマップデータは変更されています。" & vbCrLf & "マップデータを保存しますか？", vbYesNoCancel + vbExclamation, "MapEditor")
        Select Case Ret
            'キャンセルボタンなら終了を取りやめる
            Case vbCancel
                Cancel = True
                Exit Sub
            'ＯＫならファイルセーブルーチンを実行、但しそこでキャンセルされたらやはり終了はしない
            Case vbYes
                If MainForm.MapSave = False Then
                    Cancel = True
                    Exit Sub
                End If
            Case vbNo
        End Select
    End If
    
    If eDataChanged = True Then
        
        Ret = MsgBox("編集中のイベントマップデータは変更されています。" & vbCrLf & "イベントマップデータを保存しますか？", vbYesNoCancel + vbExclamation, "MapEditor")
        Select Case Ret
            'キャンセルボタンなら終了を取りやめる
            Case vbCancel
                Cancel = True
                Exit Sub
            'ＯＫならファイルセーブルーチンを実行、但しそこでキャンセルされたらやはり終了はしない
            Case vbYes
                If MainForm.eMapSave = False Then
                    Cancel = True
                    Exit Sub
                End If
            Case vbNo
        End Select
    End If

    '開いているフォームの数を減らす
    MainForm.FormCounter = MainForm.FormCounter - 1

    'フォームに付随するチップの表示などをクリアする
    MainForm.ShowChip(0).Cls
    MainForm.ShowChip(1).Cls
    MainForm.ShowChip(2).Cls
    ToolForm.LeftPic.Cls
    ToolForm.RightPic.Cls
    
    '今閉じたフォームが最後かどうか調べる
    If MainForm.FormCounter = 0 Then
    
        'ツールバーの一部を無効にする
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
'指定されたファイル名でマップチップをロードする
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
'指定されたファイル名でマップをロードする
'On Error Resume Next
    
    Do While i < 100
        If FileName = GetINIValue("NowEditFile" & CStr(i) & "-0") Or FileName = GetINIValue("NowEditFile" & CStr(i) & "-1") Then
            Call MsgBox("このファイルはすでに開いている可能性があります。" & vbCrLf & FileName & vbCrLf & vbCrLf & "（問題がないにもかかわらずこのメッセージが出る場合、" & vbCrLf & "「ファイル(F)->MapEditorの場所を開く」からMapEditor.iniを修正するか削除してください。" & vbCrLf & "またその際、NUNUまでバグ報告いただければ幸いです）")
            Exit Sub
        End If
        i = i + 1
    Loop
    
    'FileNameをバイナリ－モードでオープンしてそのまま変数に読み込む
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
            
            If Dir(BackUpName) <> "" Then BackUpSave = False    'バックアップは一回
        End If
            
    'FILEをバイナリ－モードでオープンして変数をそのまま書込む
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
    
    
    'マップの再描画
    MapShow
    Exit Sub

SaveError:

    MsgBox "Error->MapSave->" & Err.Description

   
End Sub
Public Sub ChangeMapSize(Size As Long)
'マップサイズの変更
    
    If Size < 16 Or Size > 256 Then
        Call MsgBox("マップサイズは16以上256以下で指定してください", vbOKOnly, "エラー：マップサイズの変更")
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
'マップの表示を行う
    On Error Resume Next
    Dim i As Long, j As Long
    Dim HX As Long, HY As Long
    Dim ShowX As Long, ShowY As Long
    
    ShowX = Crt.Width \ 32
    ShowY = Crt.Height \ 32
    
    '下層の描画
        For i = 0 To ShowY
            For j = 0 To ShowX
                HX = (Map(0, ((x + j) And MapSize), ((y + i) And MapSize)) And 8 * a - 1) * 32
                HY = (Map(0, ((x + j) And MapSize), ((y + i) And MapSize)) And (&HF8 - 8 * (a - 1))) * 4
                BitBlt Me.Crt.hdc, j * 32, i * 32, 32, 32, MainForm.ShowChip(0).hdc, HX, HY, SrcCopy
            Next j
        Next i
    
    '上層の描画
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
        
    'イベントマップの文字描画
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
    
    'キャプションに現在の座標を表示する
        Me.Caption = IIf(DataChanged, "*", "") & IIf(eDataChanged, "^", "") & IIf(Len(Title) > 20, "..." & Right(Title, 20), Title) & "[X:" & x & " Y:" & y & "] " & Map(0, x, y) & "-" & Map(1, x, y)
        'Me.Caption = Title & "[X:" & Hex(X) & " Y:" & Hex(Y) & "]"     16進数表記
    
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
    
    
    '選択範囲がマイナス方向の場合開始地点と終了地点を入れ換える
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
    
    '選択範囲へ網掛けを描画する（実際はただのＯＲ転送）
    For i = 0 To EndY - StartY
        For j = 0 To EndX - StartX
            BitBlt Me.Crt.hdc, (j + (StartX - x)) * 32, (i + (StartY - y)) * 32, 32, 32, IIf(eCopy, SelectPic(1).hdc, SelectPic(0).hdc), 0, 0, SrcPaint
        Next j
    Next i

    
    '再描画を行う
    Crt.Refresh
    
End Sub

Public Sub MapSelectSet()
'選択部分の番号を一括変更する

    Dim h As Integer, i As Integer, j As Integer
    Dim StartX As Integer, StartY As Integer
    Dim EndX As Integer, EndY As Integer
    
    Dim tmp As String
    
    '選択範囲がマイナス方向の場合開始地点と終了地点を入れ換える
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
        
        tmp = InputBox("マップチップデータ番号を入力してください（0~255）", IIf(ChipNow = 0, "下層", "上層") & "選択範囲一括変更")
        If tmp <> "" Then
            If CLng(tmp) < 0 Or CLng(tmp) > 255 Then
                Call MsgBox("番号は0~255の数字で入力してください", vbOKOnly, IIf(ChipNow = 0, "下層", "上層") & "選択範囲一括変更")
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
        
        tmp = InputBox("イベント番号を入力してください（0~255）", "イベント番号一括入力")
        If tmp <> "" Then
            If CLng(tmp) < 0 Or CLng(tmp) > 255 Then
                Call MsgBox("イベント番号は0~255の数字で入力してください", vbOKOnly, "イベント番号一括入力")
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
'編集中の選択部分を削除する

    Dim h As Integer, i As Integer, j As Integer
    Dim StartX As Integer, StartY As Integer
    Dim EndX As Integer, EndY As Integer
    
    '選択範囲がマイナス方向の場合開始地点と終了地点を入れ換える
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
                Map(ChipNow, (j + StartX) Mod (MapSize + 1), (i + StartY) Mod (MapSize + 1)) = IIf(Shift = 1, 0, RightNo) 'いま編集中の層のRightNoにします。シフトキー押しながらなら0にします
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
'編集中の選択部分をコピーする

    Dim h As Integer, i As Integer, j As Integer
    Dim StartX As Integer, StartY As Integer
    Dim EndX As Integer, EndY As Integer
    
    '選択範囲がマイナス方向の場合開始地点と終了地点を入れ換える
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
'コピーしたマップデータを貼り付ける
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
'コピーしたeマップデータを貼り付ける
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
'アンデゥを実行

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
'リドゥを実行

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
'変更前のデータを保存する

    ReDim UndoMap(0 To MapSize, 0 To MapSize)
    UndoMap = Map
    ToolForm.Tool1.Buttons("Redo").Enabled = False
    ToolForm.Tool1.Buttons("Undo").Enabled = True

End Sub


Public Sub MapPaint(ByVal Num As Integer)
'指定されたチップ番号でマップを塗り潰す

    Dim i As Integer, j As Integer

    For i = 0 To MapSize
        For j = 0 To MapSize
            Map(ChipNow, i, j) = Num
        Next j
    Next i
    
    'マップの再描画
    MapShow

End Sub
Public Sub ChipBarShow(Index As Integer)
'ＭＩＤフォームのチップ用ピクチャボックスにチップを再配置表示する

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
        
        '描画できないという警告の為に点描でチップを暗くする
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


    'ツールバーの選択チップを変更する
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
'Ｘ方向のスクロールバーの処理

    x = X_Scroll.Value
    MapShow
    
End Sub

Private Sub Y_Scroll_Change()
'Ｙ方向のスクロールバーの処理

    y = Y_Scroll.Value
    MapShow
    
End Sub
