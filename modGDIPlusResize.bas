Attribute VB_Name = "modGDIPlusResize"
Option Explicit

Private Type GUID
   data1    As Long
   data2    As Integer
   data3    As Integer
   data4(7) As Byte
End Type

Private Type PICTDESC
   Size     As Long
   Type     As Long
   hbmp     As Long
   hPal     As Long
   reserved As Long
End Type

Private Type IID
    data1 As Long
    data2 As Integer
    data3 As Integer
    data4(0 To 7)  As Byte
End Type
Private Type GdiplusStartupInput
    GdiplusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type PWMFRect16
    Left   As Integer
    top    As Integer
    Right  As Integer
    Bottom As Integer
End Type

Private Type wmfPlaceableFileHeader
    Key         As Long
    hMf         As Integer
    BoundingBox As PWMFRect16
    Inch        As Integer
    reserved    As Long
    CheckSum    As Integer
End Type
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal addr As Long, retval As Byte)
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal addr As Long, ByVal NewVal As Byte)
' GDI Functions
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Sub OleCreatePictureIndirect2 Lib "OleAut32.dll" Alias "OleCreatePictureIndirect" _
    (lpPictDesc As PICTDESC, riid As IID, ByVal fOwn As Boolean, _
    lplpvObj As Object)
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
' GDI+ functions
Private Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal hWnd As Long, Graphics As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mSmoothingMode As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal Graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal order As Long) As Long
Private Declare Function GdipDisposeImageAttributes Lib "gdiplus" (ByVal imgAttr As Long) As Long
Private Declare Function GdipCreateImageAttributes Lib "gdiplus" (ByRef imgAttr As Long) As Long
Private Declare Function GdipSetImageAttributesColorMatrix Lib "gdiplus" (ByVal imgAttr As Long, ByVal clrAdjust As Long, ByVal clrAdjustEnabled As Long, ByRef clrMatrix As Any, ByRef grayMatrix As Any, ByVal clrMatrixFlags As Long) As Long
Private Declare Function GdipSetImageAttributesColorKeys Lib "GdiPlus.dll" (ByVal mImageattr As Long, ByVal mType As Long, ByVal mEnableFlag As Long, ByVal mColorLow As Long, ByVal mColorHigh As Long) As Long
Private Declare Function GdipSetPixelOffsetMode Lib "GdiPlus.dll" (ByVal Graphics As Long, ByVal PixelOffsetMode As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "GdiPlus.dll" (ByVal Graphics As Long, ByVal angle As Single, ByVal order As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "GdiPlus.dll" (ByVal FileName As Long, GpImage As Long) As Long
Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (Token As Long, gdipInput As GdiplusStartupInput, GdiplusStartupOutput As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal hDC As Long, GpGraphics As Long) As Long
Private Declare Function GdipSetInterpolationMode Lib "GdiPlus.dll" (ByVal Graphics As Long, ByVal InterMode As Long) As Long
Private Declare Function GdipDrawImageRectI Lib "GdiPlus.dll" (ByVal Graphics As Long, ByVal Img As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "GdiPlus.dll" (ByVal Graphics As Long, ByVal Img As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal srcX As Long, ByVal srcY As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal Callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal Graphics As Long) As Long
Private Declare Function GdipDisposeImage Lib "GdiPlus.dll" (ByVal Image As Long) As Long
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GdiPlus.dll" (ByVal hbmp As Long, ByVal hPal As Long, GpBitmap As Long) As Long
Private Declare Function GdipGetImageWidth Lib "GdiPlus.dll" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "GdiPlus.dll" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipCreateMetafileFromWmf Lib "GdiPlus.dll" (ByVal hWmf As Long, ByVal deleteWmf As Long, WmfHeader As wmfPlaceableFileHeader, Metafile As Long) As Long
Private Declare Function GdipCreateMetafileFromEmf Lib "GdiPlus.dll" (ByVal hEmf As Long, ByVal deleteEmf As Long, Metafile As Long) As Long
Private Declare Function GdipCreateBitmapFromHICON Lib "GdiPlus.dll" (ByVal hIcon As Long, GpBitmap As Long) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal Token As Long)
Private Declare Function GdipDrawLineI Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "GdiPlus.dll" (ByVal mColor As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal mBrush As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipSetPenEndCap Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mCap As Long) As Long
Private Declare Function GdipSetPenStartCap Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mCap As Long) As Long
Private Declare Function GdipDrawLinesI Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, ByRef pPoints As Any, ByVal count As Long) As Long
Private Declare Function GdipSetPenLineJoin Lib "gdiplus" (ByVal pen As Long, ByVal lnJoin As Long) As Long
Private Declare Function GdipFillPolygon2I Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Points As Any, ByVal count As Long) As Long
Private Declare Function GdipCreateHatchBrush Lib "GdiPlus.dll" (ByVal mHatchStyle As Long, ByVal mForecol As Long, ByVal mBackcol As Long, ByRef mBrush As Long) As Long
Private Declare Function GdipSetPenDashStyle Lib "GdiPlus.dll" (ByVal mPen As Long, ByVal mDashStyle As Long) As Long
Private Declare Function GdipDrawCurveI Lib "gdiplus" (ByVal Graphics As Long, ByVal mPen As Long, Points As Any, ByVal count As Long) As Long
Public Declare Function GdipDrawCurve2I Lib "gdiplus" (ByVal Graphics As Long, ByVal mPen As Long, Points As Any, ByVal count As Long, ByVal tension As Single) As Long
Private Declare Function GdipFillClosedCurveI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Points As Any, ByVal count As Long) As Long
Private Declare Function GdipFillClosedCurve2I Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, Points As Any, ByVal count As Long, ByVal tension As Single, ByVal FillMd As Long) As Long
Private Declare Function GdipDrawBeziersI Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, Points As Any, ByVal count As Long) As Long
Private Declare Function GdipCreatePath Lib "gdiplus" (ByVal brushmode As Long, path As Long) As Long
Private Declare Function GdipAddPathBeziersI Lib "gdiplus" (ByVal path As Long, Points As Any, ByVal count As Long) As Long
Private Declare Function GdipFillPath Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal path As Long) As Long
Private Declare Function GdipDrawPath Lib "gdiplus" (ByVal Graphics As Long, ByVal mPen As Long, ByVal path As Long) As Long
Private Declare Function GdipDeletePath Lib "gdiplus" (ByVal path As Long) As Long
Private Declare Function GdipDrawEllipseI Lib "gdiplus" (ByVal Graphics As Long, ByVal mPen As Long, ByVal X As Long, ByVal Y As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipFillEllipseI Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal X As Long, ByVal Y As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipDrawPieI Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipFillPie Lib "gdiplus" (ByVal Graphics As Long, ByVal Brush As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipDrawArcI Lib "gdiplus" (ByVal Graphics As Long, ByVal pen As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipFlush Lib "gdiplus" (ByVal Graphics As Long, ByVal intention As Long) As Long
Private Declare Function GdipResetPageTransform Lib "gdiplus" (ByVal Graphics As Long) As Long
Private Declare Sub CreateStreamOnHGlobal Lib "ole32.dll" _
    (ByVal hGlobal As Any, ByVal fDeleteOnRelease As Long, _
    ByRef ppstm As Any)
    ' ----==== GDI+ Enums ====----
Private Enum Status 'GDI+ Status
    ok = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
    ProfileNotFound = 21
End Enum
Private Declare Function GdipLoadImageFromStream Lib "gdiplus" _
    (ByVal Stream As Any, ByRef Image As Long) As Status
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" _
    (ByVal BITMAP As Long, ByRef hbmReturn As Long, _
    ByVal Background As Long) As Status
' GDI and GDI+ constants
Private Const PLANES = 14            '  Number of planes
Private Const BITSPIXEL = 12         '  Number of bits per pixel
Private Const PATCOPY = &HF00021     ' (DWORD) dest = pattern
Private Const PICTYPE_BITMAP = 1     ' Bitmap type
Private Const InterpolationModeHighQualityBicubic = 7
Private Const GDIP_WMF_PLACEABLEKEY = &H9AC6CDD7
Private Const UnitPixel = 2
Public InitOk As Boolean, myToken As Long

' Initialises GDI Plus
Private Sub SetTokenNow()
Exit Sub
If InitOk = 0 Then
myToken = InitGDIPlus()
InitOk = 1
'Else
'InitOk = InitOk + 1
End If
End Sub
Private Sub ResetTokenNow()
Exit Sub
If InitOk > 0 Then
InitOk = InitOk - 1
If InitOk = 0 Then FreeGDIPlus myToken
End If
End Sub
Public Sub ResetTokenFinal()
Exit Sub
If InitOk > 0 Then
InitOk = 0
FreeGDIPlus myToken
End If
End Sub
Public Function InitGDIPlus() As Long
    Dim Token    As Long
    On Error GoTo err1
    Dim gdipInit As GdiplusStartupInput
    
    gdipInit.GdiplusVersion = 1
    GdiplusStartup Token, gdipInit, ByVal 0&
    InitGDIPlus = Token
    Exit Function
err1:
End Function

' Frees GDI Plus
Public Sub FreeGDIPlus(Token As Long)
    GdiplusShutdown Token
End Sub

' Loads the picture (optionally resized)
Public Function LoadPictureGDIPlus(picFile As String, Optional Width As Long = -1, Optional Height As Long = -1, Optional ByVal backcolor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture
    Dim hDC     As Long
    Dim hBitmap As Long
    Dim Img     As Long

    ' Load the image
    If GdipLoadImageFromFile(StrPtr(picFile), Img) <> 0 Then
        
        'Err.Raise 999, "GDI+ Module", "Error loading picture " & PicFile
        MyEr "GDI+ - can't load picture", "GDI+ - δεν μπορώ να φορώτσω την εικόνα"
        Exit Function
    End If
    
    ' Calculate picture's width and height if not specified
    If Width = -1 Or Height = -1 Then
        GdipGetImageWidth Img, Width
        GdipGetImageHeight Img, Height
    End If
    
    ' Initialise the hDC
    InitDC hDC, hBitmap, backcolor, Width, Height

    ' Resize the picture
    gdipResize Img, hDC, Width, Height, RetainRatio
    GdipDisposeImage Img
    
    ' Get the bitmap back
    GetBitmap hDC, hBitmap

    ' Create the picture
    Set LoadPictureGDIPlus = CreatePicture(hBitmap)
    
End Function
' Initialises the hDC to draw
Private Sub InitDC(hDC As Long, hBitmap As Long, backcolor As Long, Width As Long, Height As Long)
    Dim hBrush As Long
        
    ' Create a memory DC and select a bitmap into it, fill it in with the backcolor
    hDC = CreateCompatibleDC(ByVal 0&)
    hBitmap = CreateBitmap(Width, Height, GetDeviceCaps(hDC, PLANES), GetDeviceCaps(hDC, BITSPIXEL), ByVal 0&)
    hBitmap = SelectObject(hDC, hBitmap)
    hBrush = CreateSolidBrush(backcolor)
    hBrush = SelectObject(hDC, hBrush)
    PatBlt hDC, 0, 0, Width, Height, PATCOPY
    DeleteObject SelectObject(hDC, hBrush)
End Sub
Public Sub DrawLineGdi(hDC As Long, PenColor As Long, ByVal penwidth As Long, DashStyle As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long)
Dim mPen As Long, Graphics As Long
If DashStyle = 5 Then Exit Sub

GdipCreateFromHDC hDC, Graphics
GdipSetSmoothingMode Graphics, 4

If penwidth <= 1 Then penwidth = 1
If GdiPlusExec(GdipCreatePen1(GDIP_ARGB1(255, PenColor), penwidth, UnitPixel, mPen)) = ok Then
    GdipSetPenEndCap mPen, 2
    GdipSetPenStartCap mPen, 2
    GdipSetPenDashStyle mPen, DashStyle
    GdipDrawLineI Graphics, mPen, x1, y1, x2, y2
    GdipDeletePen mPen
End If
GdipDeleteGraphics Graphics

End Sub
Public Sub DrawArcPieGdi(hDC As Long, PenColor As Long, backcolor As Long, ByVal fillstyle As Long, ByVal penwidth As Long, DashStyle As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, ByVal startAngle As Single, ByVal endAngle As Single)
Dim mPen As Long, Graphics As Long, mBrush As Long, swap
If DashStyle = 5 Then Exit Sub

endAngle = MyMod(endAngle / 1.745329E-02!, 360)
startAngle = MyMod(startAngle / 1.745329E-02!, 360)
If endAngle < 0 Then
    endAngle = 360! + endAngle
End If
If startAngle < 0 Then
    startAngle = 360! + startAngle
End If
If startAngle < endAngle Then
    swap = 360! - endAngle + startAngle
Else
    swap = startAngle - endAngle
End If
startAngle = endAngle
endAngle = swap




GdipCreateFromHDC hDC, Graphics
GdipSetSmoothingMode Graphics, 4
fillstyle = fillstyle - 2
If DashStyle = 5 Then PenColor = -1
If penwidth <= 1 Then penwidth = 1
If backcolor < 0 Or fillstyle = -1 Then
If PenColor >= 0 Then
    If GdiPlusExec(GdipCreatePen1(GDIP_ARGB1(255, PenColor), penwidth, UnitPixel, mPen)) = ok Then
        GdipSetPenEndCap mPen, 2
        GdipSetPenStartCap mPen, 2
        GdipSetPenDashStyle mPen, DashStyle
        GdipDrawArcI Graphics, mPen, x1, y1, x2, y2, startAngle, endAngle
        GdipDeletePen mPen
    End If
End If
Else
If fillstyle = -2 Then
    If GdiPlusExec(GdipCreateSolidFill(GDIP_ARGB1(255, backcolor), mBrush)) = ok Then
        If PenColor >= 0 Then
            If GdiPlusExec(GdipCreatePen1(GDIP_ARGB1(255, PenColor), penwidth, UnitPixel, mPen)) = ok Then
                GdipSetPenEndCap mPen, 2
                GdipSetPenStartCap mPen, 2
                GdipSetPenLineJoin mPen, 2
                GdipSetPenDashStyle mPen, DashStyle
                GdipFillPie Graphics, mBrush, x1, y1, x2, y2, startAngle, endAngle
                GdipDrawPieI Graphics, mPen, x1, y1, x2, y2, startAngle, endAngle
                GdipDeletePen mPen
            End If
        Else
                GdipFillEllipseI Graphics, mBrush, x1, y1, x2, y2
        End If
        GdipDeleteBrush mBrush
    End If
Else
If GdiPlusExec(GdipCreateHatchBrush(fillstyle, GDIP_ARGB1(255, backcolor), GDIP_ARGB1(0, backcolor), mBrush)) = ok Then
    If PenColor >= 0 Then
        If GdiPlusExec(GdipCreatePen1(GDIP_ARGB1(255, PenColor), penwidth, UnitPixel, mPen)) = ok Then
            GdipSetPenEndCap mPen, 2
            GdipSetPenStartCap mPen, 2
            GdipSetPenLineJoin mPen, 2
            GdipSetPenDashStyle mPen, DashStyle
            GdipFillPie Graphics, mBrush, x1, y1, x2, y2, startAngle, endAngle
            GdipDrawPieI Graphics, mPen, x1, y1, x2, y2, startAngle, endAngle
            GdipDeletePen mPen
        End If
    Else
            GdipDrawPieI Graphics, mBrush, x1, y1, x2, y2, startAngle, endAngle
    End If
    GdipDeleteBrush mBrush
End If

End If
End If
GdipDeleteGraphics Graphics

End Sub
Public Sub DrawEllipseGdi(hDC As Long, PenColor As Long, backcolor As Long, ByVal fillstyle As Long, ByVal penwidth As Long, DashStyle As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long)
Dim mPen As Long, Graphics As Long, mBrush As Long
If DashStyle = 5 Then Exit Sub

GdipCreateFromHDC hDC, Graphics
GdipSetSmoothingMode Graphics, 4
fillstyle = fillstyle - 2
If DashStyle = 5 Then PenColor = -1
If penwidth <= 1 Then penwidth = 1
If backcolor < 0 Or fillstyle = -1 Then
If PenColor >= 0 Then
    If GdiPlusExec(GdipCreatePen1(GDIP_ARGB1(255, PenColor), penwidth, UnitPixel, mPen)) = ok Then
        GdipSetPenEndCap mPen, 2
        GdipSetPenStartCap mPen, 2
        GdipSetPenDashStyle mPen, DashStyle
        GdipDrawEllipseI Graphics, mPen, x1, y1, x2, y2
        GdipDeletePen mPen
    End If
End If
Else
If fillstyle = -2 Then
    If GdiPlusExec(GdipCreateSolidFill(GDIP_ARGB1(255, backcolor), mBrush)) = ok Then
        If PenColor >= 0 Then
            If GdiPlusExec(GdipCreatePen1(GDIP_ARGB1(255, PenColor), penwidth, UnitPixel, mPen)) = ok Then
                GdipSetPenEndCap mPen, 2
                GdipSetPenStartCap mPen, 2
                GdipSetPenLineJoin mPen, 2
                GdipSetPenDashStyle mPen, DashStyle
                GdipFillEllipseI Graphics, mBrush, x1, y1, x2, y2
                GdipDrawEllipseI Graphics, mPen, x1, y1, x2, y2
                GdipDeletePen mPen
            End If
        Else
                GdipFillEllipseI Graphics, mBrush, x1, y1, x2, y2
        End If
        GdipDeleteBrush mBrush
    End If
Else
If GdiPlusExec(GdipCreateHatchBrush(fillstyle, GDIP_ARGB1(255, backcolor), GDIP_ARGB1(0, backcolor), mBrush)) = ok Then
    If PenColor >= 0 Then
        If GdiPlusExec(GdipCreatePen1(GDIP_ARGB1(255, PenColor), penwidth, UnitPixel, mPen)) = ok Then
            GdipSetPenEndCap mPen, 2
            GdipSetPenStartCap mPen, 2
            GdipSetPenLineJoin mPen, 2
            GdipSetPenDashStyle mPen, DashStyle
            GdipFillEllipseI Graphics, mBrush, x1, y1, x2, y2
            GdipDrawEllipseI Graphics, mPen, x1, y1, x2, y2
            GdipDeletePen mPen
        End If
    Else
            GdipFillEllipseI Graphics, mBrush, x1, y1, x2, y2
    End If
    GdipDeleteBrush mBrush
End If

End If
End If
GdipDeleteGraphics Graphics

End Sub
Public Sub DrawLinesGdi(hDC As Long, PenColor As Long, ByVal penwidth As Long, Points() As POINTAPI, count As Long)
Dim mPen As Long, Graphics As Long

GdipCreateFromHDC hDC, Graphics
GdipSetSmoothingMode Graphics, 4

If penwidth <= 1 Then penwidth = 1
If GdiPlusExec(GdipCreatePen1(GDIP_ARGB1(255, PenColor), penwidth, UnitPixel, mPen)) = ok Then
    GdipSetPenEndCap mPen, 2
    GdipSetPenStartCap mPen, 2
    GdipSetPenLineJoin mPen, 2
    GdipDrawLinesI Graphics, mPen, ByVal VarPtr(Points(0)), count ' graphics, mPen, x1, y1, x2, y2
    
    GdipDeletePen mPen
End If
GdipDeleteGraphics Graphics

End Sub
'
Public Sub DrawBezierGdi(hDC As Long, ByVal PenColor As Long, backcolor As Long, ByVal fillstyle As Long, ByVal penwidth As Long, DashStyle As Long, Points() As POINTAPI, count As Long)
Dim mPen As Long, Graphics As Long, mBrush As Long, mPath As Long

GdipCreateFromHDC hDC, Graphics
GdipSetSmoothingMode Graphics, 4
fillstyle = fillstyle - 2
If DashStyle = 5 Then PenColor = -1
If penwidth <= 1 Then penwidth = 1
If backcolor < 0 Or fillstyle = -1 Then
If PenColor >= 0 Then
    If GdiPlusExec(GdipCreatePen1(GDIP_ARGB1(255, PenColor), penwidth, UnitPixel, mPen)) = ok Then
        GdipSetPenEndCap mPen, 2
        GdipSetPenStartCap mPen, 2
        GdipSetPenLineJoin mPen, 2
        GdipSetPenDashStyle mPen, DashStyle
        GdipDrawBeziersI Graphics, mPen, ByVal VarPtr(Points(0)), count
        GdipDeletePen mPen
    End If
End If
Else
If fillstyle = -2 Then
If GdiPlusExec(GdipCreateSolidFill(GDIP_ARGB1(255, backcolor), mBrush)) = ok Then
    If PenColor >= 0 Then
        If GdiPlusExec(GdipCreatePen1(GDIP_ARGB1(255, PenColor), penwidth, UnitPixel, mPen)) = ok Then
            GdipSetPenEndCap mPen, 2
            GdipSetPenStartCap mPen, 2
            GdipSetPenLineJoin mPen, 2
            GdipSetPenDashStyle mPen, DashStyle
            If GdiPlusExec(GdipCreatePath(1, mPath)) = ok Then
                GdipAddPathBeziersI mPath, ByVal VarPtr(Points(0)), count
                GdipFillPath Graphics, mBrush, mPath
                GdipDrawPath Graphics, mPen, mPath
                GdipDeletePath mPath
            End If
            GdipDeletePen mPen
        End If
    Else
            If GdiPlusExec(GdipCreatePath(1, mPath)) = ok Then
                GdipAddPathBeziersI mPath, ByVal VarPtr(Points(0)), count
                GdipFillPath Graphics, mBrush, mPath
                GdipDeletePath mPath
            End If
    End If
    GdipDeleteBrush mBrush
End If
Else
If GdiPlusExec(GdipCreateHatchBrush(fillstyle, GDIP_ARGB1(255, backcolor), GDIP_ARGB1(0, backcolor), mBrush)) = ok Then
    If PenColor >= 0 Then
        If GdiPlusExec(GdipCreatePen1(GDIP_ARGB1(255, PenColor), penwidth, UnitPixel, mPen)) = ok Then
            GdipSetPenEndCap mPen, 2
            GdipSetPenStartCap mPen, 2
            GdipSetPenLineJoin mPen, 2
            GdipSetPenDashStyle mPen, DashStyle
            If GdiPlusExec(GdipCreatePath(1, mPath)) = ok Then
                GdipAddPathBeziersI mPath, ByVal VarPtr(Points(0)), count
                GdipFillPath Graphics, mBrush, mPath
                GdipDrawPath Graphics, mPen, mPath
                GdipDeletePath mPath
            End If
            GdipDeletePen mPen
        End If
    Else
            If GdiPlusExec(GdipCreatePath(1, mPath)) = ok Then
                GdipAddPathBeziersI mPath, ByVal VarPtr(Points(0)), count
                GdipFillPath Graphics, mBrush, mPath
                GdipDeletePath mPath
            End If
    End If
    GdipDeleteBrush mBrush
End If

End If
End If
GdipDeleteGraphics Graphics

End Sub
Public Sub DrawPolygonGdi(hDC As Long, ByVal PenColor As Long, backcolor As Long, ByVal fillstyle As Long, ByVal penwidth As Long, DashStyle As Long, Points() As POINTAPI, count As Long)
Dim mPen As Long, Graphics As Long, mBrush As Long

GdipCreateFromHDC hDC, Graphics
GdipSetSmoothingMode Graphics, 4
fillstyle = fillstyle - 2
If DashStyle = 5 Then PenColor = -1
If penwidth <= 1 Then penwidth = 1
If backcolor < 0 Or fillstyle = -1 Then
If PenColor >= 0 Then
    If GdiPlusExec(GdipCreatePen1(GDIP_ARGB1(255, PenColor), penwidth, UnitPixel, mPen)) = ok Then
        GdipSetPenEndCap mPen, 2
        GdipSetPenStartCap mPen, 2
        GdipSetPenLineJoin mPen, 2
        GdipSetPenDashStyle mPen, DashStyle
        GdipDrawLinesI Graphics, mPen, ByVal VarPtr(Points(0)), count
        GdipDeletePen mPen
    End If
End If
Else
If fillstyle = -2 Then
If GdiPlusExec(GdipCreateSolidFill(GDIP_ARGB1(255, backcolor), mBrush)) = ok Then
    If PenColor >= 0 Then
        If GdiPlusExec(GdipCreatePen1(GDIP_ARGB1(255, PenColor), penwidth, UnitPixel, mPen)) = ok Then
            GdipSetPenEndCap mPen, 2
            GdipSetPenStartCap mPen, 2
            GdipSetPenLineJoin mPen, 2
            GdipSetPenDashStyle mPen, DashStyle
            GdipFillPolygon2I Graphics, mBrush, ByVal VarPtr(Points(0)), count  ' graphics, mPen, x1, y1, x2, y2
            GdipDrawLinesI Graphics, mPen, ByVal VarPtr(Points(0)), count
            GdipDeletePen mPen
        End If
    Else
             GdipFillPolygon2I Graphics, mBrush, ByVal VarPtr(Points(0)), count  ' graphics, mPen, x1, y1, x2, y2
    End If
    GdipDeleteBrush mBrush
End If
Else
If GdiPlusExec(GdipCreateHatchBrush(fillstyle, GDIP_ARGB1(255, backcolor), GDIP_ARGB1(0, backcolor), mBrush)) = ok Then
    If PenColor >= 0 Then
        If GdiPlusExec(GdipCreatePen1(GDIP_ARGB1(255, PenColor), penwidth, UnitPixel, mPen)) = ok Then
            GdipSetPenEndCap mPen, 2
            GdipSetPenStartCap mPen, 2
            GdipSetPenLineJoin mPen, 2
            GdipSetPenDashStyle mPen, DashStyle
            GdipFillPolygon2I Graphics, mBrush, ByVal VarPtr(Points(0)), count  ' graphics, mPen, x1, y1, x2, y2
            GdipDrawLinesI Graphics, mPen, ByVal VarPtr(Points(0)), count
            GdipDeletePen mPen
        End If
    Else
            GdipFillPolygon2I Graphics, mBrush, ByVal VarPtr(Points(0)), count  ' graphics, mPen, x1, y1, x2, y2
    End If
    GdipDeleteBrush mBrush
End If

End If
End If
GdipDeleteGraphics Graphics

End Sub
' Resize the picture using GDI plus
Private Sub gdipResize(Img As Long, hDC As Long, Width As Long, Height As Long, Optional RetainRatio As Boolean = False)
    Dim Graphics   As Long      ' Graphics Object Pointer
    Dim OrWidth    As Long      ' Original Image Width
    Dim OrHeight   As Long      ' Original Image Height
    Dim OrRatio    As Double    ' Original Image Ratio
    Dim DesRatio   As Double    ' Destination rect Ratio
    Dim DestX      As Long      ' Destination image X
    Dim DestY      As Long      ' Destination image Y
    Dim DestWidth  As Long      ' Destination image Width
    Dim DestHeight As Long      ' Destination image Height
    
    GdipCreateFromHDC hDC, Graphics
    
   GdipSetInterpolationMode Graphics, InterpolationModeHighQualityBicubic
   ' GdipSetInterpolationMode Graphics, 0
    If RetainRatio Then
        GdipGetImageWidth Img, OrWidth
        GdipGetImageHeight Img, OrHeight
        
        OrRatio = OrWidth / OrHeight
        DesRatio = Width / Height
        
        ' Calculate destination coordinates
        DestWidth = IIf(DesRatio < OrRatio, Width, Height * OrRatio)
        DestHeight = IIf(DesRatio < OrRatio, Width / OrRatio, Height)


        DestX = 0
        DestY = 0

        GdipDrawImageRectRectI Graphics, Img, DestX, DestY, DestWidth, DestHeight, 0, 0, OrWidth, OrHeight, UnitPixel, 0, 0, 0
    Else
        GdipDrawImageRectI Graphics, Img, 0, 0, Width, Height
    End If
    GdipDeleteGraphics Graphics
End Sub
Private Sub gdipResizeToXYsimple(Img As Long, hDC As Long, DestX As Long, DestY As Long, Width As Long, Height As Long, Optional RetainRatio As Boolean = False)
    Dim Graphics   As Long      ' Graphics Object Pointer
    GdipCreateFromHDC hDC, Graphics
   ' GdipSetInterpolationMode graphics, InterpolationModeHighQualityBicubic
    GdipSetPixelOffsetMode Graphics, 4
    GdipDrawImageRectI Graphics, Img, DestX, DestY, Width, Height
    GdipDeleteGraphics Graphics
End Sub
Private Function MyMod(r1 As Single, po As Single) As Single
MyMod = r1 - Fix(r1 / po) * po
End Function
Private Sub gdipResizeToXY(bstack As basetask, Img As Long, angle!, zoomfactor As Single, Alpha!, Optional backcolor As Long = -1)
    Dim clrMatrix(0 To 4, 0 To 4) As Single
    Dim hDC As Long, DestX As Long, DestY As Long
    
    
    Dim Graphics   As Long      ' Graphics Object Pointer
    Dim Width As Long
    Dim Height As Long
    
    Dim OrWidth    As Long      ' Original Image Width
    Dim OrHeight   As Long      ' Original Image Height
    Dim m_Attr As Long
    
       Const Pi = 3.14159!
    angle! = -MyMod(angle!, 360!)
    If angle! < 0 Then angle! = angle! + 360!
If zoomfactor <= 1 Then zoomfactor = 1
zoomfactor = zoomfactor / 100#

     Const ColorAdjustTypeBitmap As Long = &H1&
    
   
    Dim prive As Long, Scr As Object
    Set Scr = bstack.Owner
    If Alpha! <> 0! Or backcolor <> -1 Then Call GdipCreateImageAttributes(m_Attr)
    

    If Alpha! <> 0! Then
            If clrMatrix(4, 4) = 0! Then
                clrMatrix(0, 0) = 1!: clrMatrix(1, 1) = 1!: clrMatrix(2, 2) = 1!
                clrMatrix(3, 3) = CSng((100! - Alpha!) / 100!) ' global blending; value between 0 & 1
                clrMatrix(4, 4) = 1! ' required; cannot be anything else
            End If
            If GdipSetImageAttributesColorMatrix(m_Attr, ColorAdjustTypeBitmap, 1&, clrMatrix(0, 0), clrMatrix(0, 0), 0&) Then
                    GdipDisposeImageAttributes m_Attr
                    m_Attr = 0&
            End If
    End If
    If m_Attr And backcolor >= 0 Then
     GdipSetImageAttributesColorKeys m_Attr, 1&, 1&, GDIP_ARGB1(0, backcolor), GDIP_ARGB1(255, backcolor)
    End If
    GdipGetImageWidth Img, OrWidth
    GdipGetImageHeight Img, OrHeight
    Height = OrHeight * zoomfactor
    Width = OrWidth * zoomfactor
    
   On Error Resume Next
   
    GdipCreateFromHDC bstack.Owner.hDC, Graphics
    GdipSetPixelOffsetMode Graphics, 2
    GdipRotateWorldTransform Graphics, angle!, 1
    With players(GetCode(Scr))
    GdipTranslateWorldTransform Graphics, Scr.ScaleX(.XGRAPH, 1, 3), Scr.ScaleY(.YGRAPH, 1, 3), 1
    End With
    
    GdipDrawImageRectRectI Graphics, Img, -Width \ 2, -Height \ 2, Width, Height, 0, 0, OrWidth, OrHeight, UnitPixel, m_Attr
    
    If m_Attr Then GdipDisposeImageAttributes m_Attr
    
    GdipDeleteGraphics Graphics
End Sub


' Replaces the old bitmap of the hDC, Returns the bitmap and Deletes the hDC
Private Sub GetBitmap(hDC As Long, hBitmap As Long)
    hBitmap = SelectObject(hDC, hBitmap)
    DeleteDC hDC
End Sub

' Creates a Picture Object from a handle to a bitmap
Private Function CreatePicture(hBitmap As Long) As IPicture
    Dim IID_IDispatch As GUID
    Dim pic           As PICTDESC
    Dim IPic          As IPicture
    
    ' Fill in OLE IDispatch Interface ID
    IID_IDispatch.data1 = &H20400
    IID_IDispatch.data4(0) = &HC0
    IID_IDispatch.data4(7) = &H46
        
    ' Fill Pic with necessary parts
    pic.Size = Len(pic)        ' Length of structure
    pic.Type = PICTYPE_BITMAP  ' Type of Picture (bitmap)
    pic.hbmp = hBitmap         ' Handle to bitmap

    ' Create the picture
    OleCreatePictureIndirect pic, IID_IDispatch, True, IPic
    Set CreatePicture = IPic
End Function

' Returns a resized version of the picture
Public Function Resize(handle As Long, picType As PictureTypeConstants, Width As Long, Height As Long, Optional backcolor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture
    Dim Img       As Long
    Dim hDC       As Long
    Dim hBitmap   As Long
    Dim WmfHeader As wmfPlaceableFileHeader
    
    ' Determine pictyre type
    Select Case picType
    Case vbPicTypeBitmap
         GdipCreateBitmapFromHBITMAP handle, ByVal 0&, Img
    Case vbPicTypeMetafile
         FillInWmfHeader WmfHeader, Width, Height
         GdipCreateMetafileFromWmf handle, False, WmfHeader, Img
    Case vbPicTypeEMetafile
         GdipCreateMetafileFromEmf handle, False, Img
    Case vbPicTypeIcon
         ' Does not return a valid Image object
         GdipCreateBitmapFromHICON handle, Img
    End Select
    
    ' Continue with resizing only if we have a valid image object
    If Img Then
        InitDC hDC, hBitmap, backcolor, Width, Height
        gdipResize Img, hDC, Width, Height, RetainRatio
        GdipDisposeImage Img
        GetBitmap hDC, hBitmap
        Set Resize = CreatePicture(hBitmap)
    End If
End Function

' Fills in the wmfPlacable header
Private Sub FillInWmfHeader(WmfHeader As wmfPlaceableFileHeader, Width As Long, Height As Long)
    WmfHeader.BoundingBox.Right = Width
    WmfHeader.BoundingBox.Bottom = Height
    WmfHeader.Inch = 1440
    WmfHeader.Key = GDIP_WMF_PLACEABLEKEY
End Sub
Public Function ReadSizeImageFromBuffer(ResData() As Byte, Width As Long, Height) As Boolean
    
    On Error GoTo PROC_ERR
    Dim Stream As IUnknown
    Dim hDC As Long
    Dim Img As Long
    Dim hBitmap As Long
    
    Call CreateStreamOnHGlobal(VarPtr(ResData(0)), _
    False, Stream)
    If Not (Stream Is Nothing) Then
        If GdiPlusExec(GdipLoadImageFromStream( _
        Stream, Img)) = ok Then
            
                    GdipGetImageWidth Img, Width
                    GdipGetImageHeight Img, Height
            
            

    
            ReadSizeImageFromBuffer = True
        End If
    End If
    
PROC_EXIT:
    Set Stream = Nothing
    
    Exit Function
    
PROC_ERR:
Dim er$
er$ = "GDI+: " & Err.Number & ". " & Err.Description
    MyEr er$, er$
    Err.clear
    Resume PROC_EXIT

End Function
Public Function LoadImageFromBuffer2(ResData() As Byte, Optional Width As Long = -1, Optional Height As Long = -1, Optional ByVal backcolor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture
    
    On Error GoTo PROC_ERR
    Dim Stream As IUnknown
    Dim hDC As Long
    Dim Img As Long
    Dim hBitmap As Long
    
    ' Ressource in ByteArray speichern
    
    ' Stream erzeugen
    Call CreateStreamOnHGlobal(VarPtr(ResData(0)), _
    False, Stream)
    
    ' ist ein Stream vorhanden
    If Not (Stream Is Nothing) Then
        
        ' GDI+ Bitmapobjekt vom Stream erstellen
        If GdiPlusExec(GdipLoadImageFromStream( _
        Stream, Img)) = ok Then
            
            
                  If Width = -1 Or Height = -1 Then
                    GdipGetImageWidth Img, Width
                    GdipGetImageHeight Img, Height
                End If
                 ' Initialise the hDC
                  InitDC hDC, hBitmap, backcolor, Width, Height

                ' Resize the picture
                gdipResize Img, hDC, Width, Height, RetainRatio
            GdipDisposeImage Img
    
    ' Get the bitmap back
    GetBitmap hDC, hBitmap

    ' Create the picture
    Set LoadImageFromBuffer2 = CreatePicture(hBitmap)
         
            
            
            
        End If
    End If
    
PROC_EXIT:
    Set Stream = Nothing
    
    Exit Function
    
PROC_ERR:
Dim er$
er$ = "GDI+: " & Err.Number & ". " & Err.Description
    MyEr er$, er$
    Err.clear
    Resume PROC_EXIT

End Function
Public Function DrawImageFromBuffer(ResData() As Byte, hDC As Long, Optional X As Long = 0&, Optional Y As Long = 0&, Optional Width As Long = -1, Optional Height As Long = -1) As Boolean
    
    On Error GoTo PROC_ERR
    Dim Stream As IUnknown
    Dim Img As Long
    Dim hBitmap As Long
    
    Call CreateStreamOnHGlobal(VarPtr(ResData(0)), False, Stream)
    
    If Not (Stream Is Nothing) Then
        If GdiPlusExec(GdipLoadImageFromStream(Stream, Img)) = ok Then
            
    Dim OldWidth As Long
    If Width = -1 Or Height = -1 Then
        If Width = -1 Then
            GdipGetImageWidth Img, Width
            GdipGetImageHeight Img, Height
            
        Else
            GdipGetImageWidth Img, OldWidth
            GdipGetImageHeight Img, Height
            Height = Height * Width / OldWidth
            End If
        End If
        gdipResizeToXYsimple Img, hDC, X, Y, Width, Height
        GdipDisposeImage Img
        End If
    End If
    
PROC_EXIT:
    Set Stream = Nothing
    
    Exit Function
    
PROC_ERR:
Dim er$
er$ = "GDI+: " & Err.Number & ". " & Err.Description
    MyEr er$, er$
    Err.clear

    Resume PROC_EXIT

End Function
Public Function DrawSpriteFromBuffer(bstack As basetask, ResData() As Byte, sprt As Boolean, angle!, zoomfactor!, blend!, Optional backcolor As Long = -1) As Boolean
    
    On Error GoTo PROC_ERR
    Dim Stream As IUnknown
    Dim Img As Long
    Dim hBitmap As Long, hDC As Long
    
    Call CreateStreamOnHGlobal(VarPtr(ResData(0)), False, Stream)
    Dim Width As Long, Height As Long
    If Not (Stream Is Nothing) Then
        If GdiPlusExec(GdipLoadImageFromStream(Stream, Img)) = ok Then
            GdipGetImageWidth Img, Width
            GdipGetImageHeight Img, Height
            If sprt Then GetBackSprite bstack, Width, Height, angle!, zoomfactor
            gdipResizeToXY bstack, Img, angle!, zoomfactor!, blend!, backcolor
            GdipDisposeImage Img
            
        End If
    End If
    
PROC_EXIT:
    Set Stream = Nothing
    
    Exit Function
    
PROC_ERR:
Dim er$
er$ = "GDI+: " & Err.Number & ". " & Err.Description
    MyEr er$, er$
    Err.clear

    Resume PROC_EXIT

End Function

Public Function LoadImageFromBuffer( _
 ResData() As Byte) As StdPicture
    
    On Error GoTo PROC_ERR
    Dim Stream As IUnknown
    Dim lBitmap As Long
    Dim hBitmap As Long
    
    ' Ressource in ByteArray speichern
    
    ' Stream erzeugen
    Call CreateStreamOnHGlobal(VarPtr(ResData(0)), _
    False, Stream)
    
    ' ist ein Stream vorhanden
    If Not (Stream Is Nothing) Then
        
        ' GDI+ Bitmapobjekt vom Stream erstellen
        If GdiPlusExec(GdipLoadImageFromStream( _
        Stream, lBitmap)) = ok Then
            
            ' Handle des Bitmapobjektes ermitteln
            If GdiPlusExec(GdipCreateHBITMAPFromBitmap( _
            lBitmap, hBitmap, 0)) = ok Then
                
                ' StdPicture Objekt erstellen
                Set LoadImageFromBuffer = _
                HandleToPicture(hBitmap, vbPicTypeBitmap)
                
            End If
            
            ' Bitmapobjekt lφschen
            Call GdiPlusExec(GdipDisposeImage(lBitmap))
        End If
    End If
    
PROC_EXIT:
    Set Stream = Nothing
    Exit Function
    
PROC_ERR:
Dim er$
er$ = "GDI+: " & Err.Number & ". " & Err.Description
    MyEr er$, er$
    Err.clear

    Resume PROC_EXIT

End Function
Private Function GdiErrorString(ByVal lError As Status) As String
    Dim s As String
    
    Select Case lError
    Case GenericError:              s = "Generic Error."
    Case InvalidParameter:          s = "Invalid Parameter."
    Case OutOfMemory:               s = "Out Of Memory."
    Case ObjectBusy:                s = "Object Busy."
    Case InsufficientBuffer:        s = "Insufficient Buffer."
    Case NotImplemented:            s = "Not Implemented."
    Case Win32Error:                s = "Win32 Error."
    Case WrongState:                s = "Wrong State."
    Case Aborted:                   s = "Aborted."
    Case FileNotFound:              s = "File Not Found."
    Case ValueOverflow:             s = "Value Overflow."
    Case AccessDenied:              s = "Access Denied."
    Case UnknownImageFormat:        s = "Unknown Image Format."
    Case FontFamilyNotFound:        s = "FontFamily Not Found."
    Case FontStyleNotFound:         s = "FontStyle Not Found."
    Case NotTrueTypeFont:           s = "Not TrueType Font."
    Case UnsupportedGdiplusVersion: s = "Unsupported Gdiplus Version."
    Case GdiplusNotInitialized:     s = "Gdiplus Not Initialized."
    Case PropertyNotFound:          s = "Property Not Found."
    Case PropertyNotSupported:      s = "Property Not Supported."
    Case Else:                      s = "Unknown GDI+ Error."
    End Select
    
    GdiErrorString = s
End Function
Private Function GdiPlusExec(ByVal lReturn As Status) As Status
    Dim lCurErr As Status
    If lReturn = Status.ok Then
        lCurErr = Status.ok
    Else
        lCurErr = lReturn
        Dim er$
    er$ = "GDI+: " & GdiErrorString(lReturn) & " GDI+ Error:" & lReturn
    MyEr er$, er$
    Err.clear
    End If
    GdiPlusExec = lCurErr
End Function
Private Function HandleToPicture(ByVal hGDIHandle As Long, _
    ByVal ObjectType As PictureTypeConstants, _
    Optional ByVal hPal As Long = 0) As StdPicture
    
    Dim tPictDesc As PICTDESC
    Dim IID_IPicture As IID
    Dim oPicture As IPicture
    
    ' Initialisiert die PICTDESC Structur
    With tPictDesc
         .Size = Len(tPictDesc)
        .Type = ObjectType
        .hbmp = hGDIHandle
        .hPal = hPal
    End With
    
    ' Initialisiert das IPicture Interface ID
    With IID_IPicture
        .data1 = &H7BF80981
        .data2 = &HBF32
        .data3 = &H101A
        .data4(0) = &H8B
        .data4(1) = &HBB
        .data4(3) = &HAA
        .data4(5) = &H30
        .data4(6) = &HC
        .data4(7) = &HAB
    End With
    
    ' Erzeugen des Objekts
    OleCreatePictureIndirect2 tPictDesc, _
    IID_IPicture, True, oPicture
    
    ' Rόckgabe des Pictureobjekts
    Set HandleToPicture = oPicture
    
End Function
Function GDIP_ARGB(Alpha As Long, red As Long, green As Long, blue As Long) As Long
Dim b As Byte
GetMem1 VarPtr(Alpha), b
PutMem1 VarPtr(GDIP_ARGB) + 3, b
GetMem1 VarPtr(red), b
PutMem1 VarPtr(GDIP_ARGB) + 2, b
GetMem1 VarPtr(green), b
PutMem1 VarPtr(GDIP_ARGB) + 1, b
GetMem1 VarPtr(blue), b
PutMem1 VarPtr(GDIP_ARGB), b
End Function
Function GDIP_ARGB1(Alpha As Long, color As Long) As Long
Dim b As Byte
GetMem1 VarPtr(Alpha), b
PutMem1 VarPtr(GDIP_ARGB1) + 3, b
GetMem1 VarPtr(color) + 2, b
PutMem1 VarPtr(GDIP_ARGB1), b
GetMem1 VarPtr(color) + 1, b
PutMem1 VarPtr(GDIP_ARGB1) + 1, b
GetMem1 VarPtr(color), b
PutMem1 VarPtr(GDIP_ARGB1) + 2, b

End Function
