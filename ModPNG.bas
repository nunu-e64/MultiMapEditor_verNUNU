Attribute VB_Name = "ModPNG"
Option Explicit

'ModPNG.bas
'mpng.dll 1.10付属　PNG形式の画像ファイル読み書き用モジュール
'CopyRight (C) minutes 2002-2003


'ユーザーが変更可能な定数

'画面のカラーモードが32ビットのときでSavePNGのbppを省略またはColorModeで
'呼び出したとき、24ビット画像で保存するかどうか指定します。
'Trueにすると24ビットで保存されます。Falseだと32ビットになります。
'初期値ではTrueにセットしてあります。
Private Const REVISE_32_to_24 As Boolean = True








'==============================================================================
'CreateBitmapPictureで使う構造体と関数の定義
'GUIDDEF.Hを参照
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
'OLECTL.Hを参照
Private Type PICTDESC_BMP
    Size As Long
    Type As Long
    hbmp As Long
    hPal As Long
    Reserved As Long
End Type
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PICTDESC_BMP, riid As GUID, ByVal fOwn As Long, lplpvObj As IPicture) As Long

'==============================================================================
'CreatePictureFromDIBで使う構造体と関数の定義
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type

Private Declare Function CreateDIBitmap Lib "gdi32" (ByVal hdc As Long, lpbmih As BITMAPINFOHEADER, ByVal fdwInit As Long, lpbInit As Any, lpbmi As BITMAPINFO, ByVal fuUsage As Long) As Long
Private Const CBM_INIT = &H4
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

'==============================================================================
'LoadPNGで使う関数の定義

'mpng.dll API
Private Declare Function mLoadPNG Lib "mpng.dll" (ByVal strfilename As String, hPng As Long, bminfo As BITMAPINFO, length As Long) As Long
Private Declare Function mEndPNG Lib "mpng.dll" (hPng As Long) As Long
Private Declare Function mGetPNGData Lib "mpng.dll" (hPng As Long, buf As Byte) As Long


'==============================================================================
'SavePNGで使う関数・定数・構造体の定義
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Public Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020
Private Const WHITENESS = &HFF0062
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Const BI_RGB As Long = 0
Private Const DIB_RGB_COLORS = 0

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const IMAGE_BITMAP = 0
Private Const LR_COPYRETURNORG = 4

Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSX = 88
Private Const LOGPIXELSY = 90
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Declare Function PlayMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hMF As Long) As Long
Private Declare Function PlayEnhMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hemf As Long, lpRect As RECT) As Long
Private Declare Function CreateEnhMetaFile Lib "gdi32" Alias "CreateEnhMetaFileA" (ByVal hdcRef As Long, ByVal lpFileName As String, lpRect As Any, ByVal lpDescription As String) As Long
Private Declare Function CloseEnhMetaFile Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteEnhMetaFile Lib "gdi32" (ByVal hemf As Long) As Long
Private Declare Function SetWinMetaFileBits Lib "gdi32" (ByVal cbBuffer As Long, lpbBuffer As Byte, ByVal hdcRef As Long, lpmfp As Any) As Long
Private Declare Function GetMetaFileBitsEx Lib "gdi32" (ByVal hMF As Long, ByVal nSize As Long, lpvData As Any) As Long


'mpng.dll API
Private Declare Function mWritePNG Lib "mpng.dll" (ByVal strfilename As String, lpdat As Byte, bminfo As BITMAPINFO, ByVal Interlace As Long) As Long

Public Enum ModPNGColorTypeConstants
    ColorMode = 0
    PALETTE_8bit = 8
    RGB_24bit = 24
    RGB_ALPHA_32bit = 32
End Enum

















'hBmp(DDB)からstdPictre型のピクチャーを生成する
Private Function CreateBitmapPicture(ByVal hbmp As Long, Optional ByVal hPal As Long = 0) As StdPicture

    Dim ret As Long
    Dim PicInfo As PICTDESC_BMP
    Dim sPic As StdPicture
    Dim IID_IPicture As GUID
    
    'IPictureインターフェイスのIDを設定
    'OCIDL.Hより7BF80980-BF32-101A-8BBB-00AA00300CAB : IPicture
    'こっち↓の方がいいかもしれない
    'OAIDL.Hより00020400-0000-0000-C000-000000000046 : IDispatch
    With IID_IPicture
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
    
    'ピクチャーの情報を設定
    With PicInfo
       .Size = Len(PicInfo)      '構造体サイズ
       .Type = vbPicTypeBitmap   'ピクチャーの種類 (ビットマップだけ対応するつもり)
       .hbmp = hbmp              'ビットマップへのポインタ
       .hPal = hPal              'パレットへのポインタ（いらない場合はNULLでもよい）
    End With
    
    'stdPictureオブジェクトを作成
    If OleCreatePictureIndirect(PicInfo, IID_IPicture, 1, sPic) Then Exit Function
    
    'オブジェクトを返す
    Set CreateBitmapPicture = sPic
End Function



'DIBからstdPictureオブジェクトを作る
Private Function CreatePictureFromDIB(dib As Byte, Info As BITMAPINFO) As StdPicture
    Dim hdc As Long, hbmp As Long
    
    '現在の画面と互換のＤＣを確保
    hdc = GetDC(0)
    If hdc = 0 Then Exit Function
    
    'DIBからDDBを作成
    hbmp = CreateDIBitmap(hdc, Info.bmiHeader, CBM_INIT, dib, Info, 0)
    If hbmp = 0 Then Exit Function
    
    'hBmpからstdPictureオブジェクトを生成
    Set CreatePictureFromDIB = CreateBitmapPicture(hbmp)
    
    'DCの解放。hBmpはオブジェクトと共に破棄されるので処理する必要は無い（と思うのですが）
    ReleaseDC 0, hdc
End Function




'PNG形式の画像を読み込む
Public Function LoadPNG(ByVal strfilename As String) As StdPicture
    Dim hPng As Long
    Dim ret As Long
    Dim bmi As BITMAPINFO
    Dim length As Long
    Dim dib() As Byte
    
    On Error GoTo Er

    'step1 ファイルをロードし情報を取得する
    ret = mLoadPNG(strfilename, hPng, bmi, length)
    If ret <> 1 Then
        Debug.Print "Error! code: " & ret
        Exit Function
    End If
    
    'step2 データを受け取る
    ReDim dib(length - 1) As Byte
    ret = mGetPNGData(hPng, dib(0))
    If ret <> 1 Then
        Debug.Print "Error! code: " & ret
        mEndPNG hPng
        Exit Function
    End If
    
    'step3 ハンドルを解放
    mEndPNG hPng
    
    'ピクチャーを取得
    Set LoadPNG = CreatePictureFromDIB(dib(0), bmi)
    Exit Function
Er:
    If Err.Number Then Debug.Print "Error " & Err.Number & " : " & Err.Description
    Set LoadPNG = Nothing
End Function



Public Function SavePNG(pict As StdPicture, ByVal strfilename As String, Optional ByVal bpp As ModPNGColorTypeConstants = ColorMode, Optional ByVal Interlace As Boolean = False) As Boolean
    Dim bminfo As BITMAPINFO
    Dim dib() As Byte
    Dim ret As Long
    
    If strfilename = "" Then Exit Function
    If pict Is Nothing Then Exit Function

    'DIBを取得する
    Select Case pict.Type
    Case vbPicTypeNone
        Exit Function
    Case vbPicTypeBitmap
        'ビットマップ
        If GetDIBfromBitmap(pict.handle, dib, bminfo, bpp) = False Then
            Exit Function
        End If
        
    Case vbPicTypeIcon
        'アイコン
        If GetDIBfromIcon(pict.handle, dib, bminfo) = False Then
            Exit Function
        End If
    
    Case vbPicTypeMetafile, vbPicTypeEMetafile
        'メタファイル及び拡張メタファイル
        If GetDIBfromMF(pict, dib, bminfo, bpp) = False Then
            Exit Function
        End If

    End Select
    
    
    'PNGとして保存
    If bminfo.bmiHeader.biHeight < 0 Then
        bminfo.bmiHeader.biHeight = -bminfo.bmiHeader.biHeight
    End If
    ret = mWritePNG(strfilename, dib(0), bminfo, -CLng(Interlace))
    If ret Then
        SavePNG = True
    Else
        SavePNG = False
    End If
End Function



'ここから下の関数はほぼコピペで作ってあります。長くて読みにくいですがご容赦を。
Private Function GetDIBfromBitmap(ByVal handle As Long, dib() As Byte, bminfo As BITMAPINFO, bpp As Long) As Boolean
    Dim ret As Long
    Dim hdc As Long, tmpdc As Long
    Dim hbmp As Long, oldbmp As Long
    Dim bmp As BITMAP
    
    
    'hbmpを作る
    hbmp = CopyImage(handle, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
    If hbmp = 0 Then Exit Function
    
    'DCを作る
    tmpdc = GetDC(0)
    If tmpdc = 0 Then GoTo Failed
    hdc = CreateCompatibleDC(tmpdc)
    ReleaseDC 0, tmpdc
    If hdc = 0 Then GoTo Failed
    
    'DCとhbmpを関連付けする
    oldbmp = SelectObject(hdc, hbmp)
    If oldbmp = 0 Then GoTo Failed
    
    'bminfoを準備
    ret = GetObject(hbmp, LenB(bmp), bmp)
    If ret = 0 Then GoTo Failed
    With bminfo.bmiHeader
        .biSize = LenB(bminfo.bmiHeader)
        .biWidth = bmp.bmWidth
        .biHeight = -bmp.bmHeight
        If bpp = ColorMode Then
            Select Case bmp.bmBitsPixel
            Case Is <= 8
                .biBitCount = 8
            Case Is <= 24
                .biBitCount = 24
            Case Is > 24
                .biBitCount = 32
                If REVISE_32_to_24 Then
                    .biBitCount = 24
                End If
            End Select
        Else
            .biBitCount = bpp
        End If
        .biPlanes = 1
    End With
    
    'メモリを確保してGetDIBitsを実行。DIBを取得
    ReDim dib(((bmp.bmWidth * bminfo.bmiHeader.biBitCount \ 8 + 3) And &H7FFFFFFC) * bmp.bmHeight) As Byte
    ret = GetDIBits(hdc, hbmp, 0, bmp.bmHeight, dib(0), bminfo, DIB_RGB_COLORS)
    If ret = 0 Then GoTo Failed
    
    '後片付け
    SelectObject hdc, oldbmp
    DeleteObject hbmp
    DeleteDC hdc
    
    GetDIBfromBitmap = True
    Exit Function
    
Failed:
    If hbmp <> 0 Then DeleteObject hbmp
    If hdc <> 0 Then DeleteDC hdc
End Function

Private Function GetDIBfromIcon(ByVal handle As Long, dib() As Byte, bminfo As BITMAPINFO) As Boolean
    Dim ret As Long
    Dim hdc As Long, tmpdc As Long
    Dim hbmp As Long, oldbmp As Long
    Dim bmp As BITMAP, iinfo As ICONINFO
    
    '情報を取得
    ret = GetIconInfo(handle, iinfo)
    If ret = 0 Then GoTo Failed
    ret = GetObject(iinfo.hbmColor, LenB(bmp), bmp)
    If ret = 0 Then GoTo Failed
    
    'DCとhbmpを作る
    tmpdc = GetDC(0)
    If tmpdc = 0 Then Exit Function
    hdc = CreateCompatibleDC(tmpdc)
    hbmp = CreateCompatibleBitmap(tmpdc, bmp.bmWidth, bmp.bmHeight)
    ReleaseDC 0, tmpdc
    If hdc = 0 Or hbmp = 0 Then GoTo Failed
    
    
    'DCとhbmpを関連付けする
    oldbmp = SelectObject(hdc, hbmp)
    If oldbmp = 0 Then GoTo Failed
    
    'アイコンを描画
    BitBlt hdc, 0, 0, bmp.bmWidth, bmp.bmHeight, hdc, 0, 0, WHITENESS
    ret = DrawIcon(hdc, 0, 0, handle)
    If ret = 0 Then GoTo Failed
    
    'bminfoを準備
    ret = GetObject(hbmp, LenB(bmp), bmp)
    If ret = 0 Then GoTo Failed
    With bminfo.bmiHeader
        .biSize = LenB(bminfo.bmiHeader)
        .biWidth = bmp.bmWidth
        .biHeight = -bmp.bmHeight
        .biBitCount = 8
        .biPlanes = 1
    End With

    'メモリを確保してGetDIBitsを実行。DIBを取得
    ReDim dib(((bmp.bmWidth * bminfo.bmiHeader.biBitCount \ 8 + 3) And &H7FFFFFFC) * bmp.bmHeight) As Byte
    ret = GetDIBits(hdc, hbmp, 0, bmp.bmHeight, dib(0), bminfo, DIB_RGB_COLORS)
    If ret = 0 Then GoTo Failed
    
    '後片付け
    SelectObject hdc, oldbmp
    DeleteObject hbmp
    DeleteDC hdc
    
    GetDIBfromIcon = True
    Exit Function
Failed:
    If hbmp <> 0 Then DeleteObject hbmp
    If hdc <> 0 Then DeleteDC hdc
End Function

Private Function GetDIBfromMF(pict As StdPicture, dib() As Byte, bminfo As BITMAPINFO, ByVal bpp As Long) As Boolean
    Dim ret As Long
    Dim hdc As Long, tmpdc As Long
    Dim hbmp As Long, oldbmp As Long
    Dim bmp As BITMAP
    Dim width As Long, height As Long
    Dim hemf As Long
    Dim arysize As Long, ary() As Byte
    Dim r As RECT
    
    Const HIMETRIC = 2540
    
    'DCとhbmpの作成及び情報の取得
    tmpdc = GetDC(0)
    If tmpdc = 0 Then Exit Function
    width = MulDiv(pict.width, GetDeviceCaps(tmpdc, LOGPIXELSX), HIMETRIC)
    height = MulDiv(pict.height, GetDeviceCaps(tmpdc, LOGPIXELSY), HIMETRIC)
    hdc = CreateCompatibleDC(tmpdc)
    hbmp = CreateCompatibleBitmap(tmpdc, width, height)
    ret = GetObject(hbmp, LenB(bmp), bmp)
    ReleaseDC 0, tmpdc
    If ret = 0 Then GoTo Failed
    If hdc = 0 Or hbmp = 0 Then GoTo Failed
    
    'DCとhbmpを関連付けする
    oldbmp = SelectObject(hdc, hbmp)
    If oldbmp = 0 Then GoTo Failed
    
    If pict.Type = vbPicTypeMetafile Then
        'メタファイルを拡張メタファイルに変換
        arysize = GetMetaFileBitsEx(pict.handle, 1, ByVal 0&)
        If arysize = 0 Then GoTo Failed
        ReDim ary(arysize - 1) As Byte
        ret = GetMetaFileBitsEx(pict.handle, arysize, ary(0))
        If ret = 0 Then GoTo Failed
        hemf = SetWinMetaFileBits(arysize, ary(0), hdc, ByVal 0&)
        If hemf = 0 Then GoTo Failed
        '変換した拡張メタファイルで描画
        r.Right = width
        r.Bottom = height
        BitBlt hdc, 0, 0, bmp.bmWidth, bmp.bmHeight, hdc, 0, 0, WHITENESS
        PlayEnhMetaFile hdc, hemf, r
        DeleteEnhMetaFile hemf
    ElseIf pict.Type = vbPicTypeEMetafile Then
        '拡張メタファイルを描画
        With r
            .Right = bmp.bmWidth
            .Bottom = bmp.bmHeight
        End With
        BitBlt hdc, 0, 0, bmp.bmWidth, bmp.bmHeight, hdc, 0, 0, WHITENESS
        PlayEnhMetaFile hdc, pict.handle, r
    End If
    
    'bminfoを準備
    With bminfo.bmiHeader
        .biSize = LenB(bminfo.bmiHeader)
        .biWidth = width
        .biHeight = -height
        If bpp = ColorMode Then
            Select Case bmp.bmBitsPixel
            Case Is <= 8
                .biBitCount = 8
            Case Is <= 24
                .biBitCount = 24
            Case Is > 24
                .biBitCount = 32
                If REVISE_32_to_24 Then
                    .biBitCount = 24
                End If
            End Select
        Else
            .biBitCount = bpp
        End If
        .biPlanes = 1
    End With
    
    'メモリを確保してGetDIBitsを実行。DIBを取得
    ReDim dib(((bmp.bmWidth * bminfo.bmiHeader.biBitCount \ 8 + 3) And &H7FFFFFFC) * bmp.bmHeight) As Byte
    ret = GetDIBits(hdc, hbmp, 0, bmp.bmHeight, dib(0), bminfo, DIB_RGB_COLORS)
    If ret = 0 Then GoTo Failed
    
    '後片付け
    SelectObject hdc, oldbmp
    DeleteObject hbmp
    DeleteDC hdc
    
    GetDIBfromMF = True
    Exit Function
Failed:
    If hbmp <> 0 Then DeleteObject hbmp
    If hdc <> 0 Then DeleteDC hdc
End Function
   
