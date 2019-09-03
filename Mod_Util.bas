Attribute VB_Name = "Module2"
Option Explicit
Dim FontList As FastCollection
Public Trush() As VarItem
Public TrushCount As Long, TrushWait As Boolean
Const b123 = vbCr + "'\"
Const b1234 = vbCr + "'\:"
Public k1 As Long, Kform As Boolean
Private Const doc = "Document"
Public tracecode As String, lasttracecode As Long
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public stackshowonly As Boolean, NoBackFormFirstUse As Boolean
Private Declare Function IsBadCodePtr Lib "KERNEL32" (ByVal lpfn As Long) As Long
Public Pow2(33) As Currency, Pow2minusOne(33) As Currency
Public Enum Ftypes
    FnoUse
    Finput
    Foutput
    Fappend
    Frandom
End Enum
Public FLEN(512) As Long, FKIND(512) As Ftypes
Public Type Counters
    k1 As Long
    RRCOUNTER As Long
End Type
Public Type basket
    used As Long
    X As Long  ' for hotspot
    Y As Long  '
    XGRAPH As Long  ' graphic cursor
    YGRAPH As Long
    MAXXGRAPH As Long
    MAXYGRAPH As Long
    dv15 As Long  ' not used
    curpos As Long   ' text cursor
    currow As Long
    mypen As Long
    mysplit As Long
    Paper As Long
    italics As Boolean  ' removed from process, only in current
    bold As Boolean
    double As Boolean
    osplit As Long  '(for double size letters)
    Column As Long
    OCOLUMN As Long
    pageframe As Long
    basicpageframe As Long
    MineLineSpace As Long
    uMineLineSpace As Long
    LastReportLines As Double
    SZ As Single
    UseDouble As Single
    Xt As Long
    Yt As Long
    mx As Long
    My As Long
    FontName As String
    charset As Long
    FTEXT As Long
    FTXT As String
    lastprint As Boolean  ' if true then we have to place letters using currentX
    ' gdi drawing enabled Smooth On, disabled with Smooth Of
    NoGDI As Boolean
    pathgdi As Long  ' only for gdi+
    pathcolor As Long ' only for gdi+
    pathfillstyle As Integer
    LastIcon As Integer  ' 1..   / 99 loaded
    LastIconPic As StdPicture
    HideIcon As Boolean
    ReportTab As Long
End Type
Private stopwatch As Long
Private Const myArray = "mArray"
Private Const LOCALE_SYSTEM_DEFAULT As Long = &H800
Private Const LOCALE_USER_DEFAULT As Long = &H800
Private Const C3_DIACRITIC As Long = &H2
Private Const CT_CTYPE3 As Byte = &H4
Private Declare Function GetStringTypeExW Lib "kernel32.dll" (ByVal Locale As Long, ByVal dwInfoType As Long, ByVal lpSrcStr As Long, ByVal cchSrc As Long, ByRef lpCharType As Byte) As Long
Private Declare Function SetTextCharacterExtra Lib "gdi32" (ByVal hDC As Long, ByVal nCharExtra As Long) As Long
Private Declare Function WideCharToMultiByte Lib "KERNEL32" (ByVal codepage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function GdiFlush Lib "gdi32" () As Long
Public iamactive As Boolean
Declare Function MultiByteToWideChar& Lib "KERNEL32" (ByVal codepage&, ByVal dwFlags&, MultiBytes As Any, ByVal cBytes&, ByVal pWideChars&, ByVal cWideChars&)
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

Public Type RECT
        Left As Long
        top As Long
        Right As Long
        Bottom As Long
End Type
Public Const LOCALE_SDECIMAL = &HE&
Public Const LOCALE_SGROUPING = &H10&
Public Const LOCALE_STHOUSAND = &HF&
Public Const LOCALE_SMONDECIMALSEP = &H16&
Public Const LOCALE_SMONTHOUSANDSEP = &H17&
Public Const LOCALE_SMONGROUPING = &H18&
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Const DT_BOTTOM As Long = &H8&
Private Const DT_CALCRECT As Long = &H400&
Private Const DT_CENTER As Long = &H1&
Private Const DT_EDITCONTROL As Long = &H2000&
Private Const DT_END_ELLIPSIS As Long = &H8000&
Private Const DT_EXPANDTABS As Long = &H40&
Private Const DT_EXTERNALLEADING As Long = &H200&
Private Const DT_HIDEPREFIX As Long = &H100000
Private Const DT_INTERNAL As Long = &H1000&
Private Const DT_LEFT As Long = &H0&
Private Const DT_MODIFYSTRING As Long = &H10000
Private Const DT_NOCLIP As Long = &H100&
Private Const DT_NOFULLWIDTHCHARBREAK As Long = &H80000
Private Const DT_NOPREFIX As Long = &H800&
Private Const DT_PATH_ELLIPSIS As Long = &H4000&
Private Const DT_PREFIXONLY As Long = &H200000
Private Const DT_RIGHT As Long = &H2&
Private Const DT_SINGLELINE As Long = &H20&
Private Const DT_TABSTOP As Long = &H80&
Private Const DT_TOP As Long = &H0&
Private Const DT_VCENTER As Long = &H4&
Private Const DT_WORDBREAK As Long = &H10&
Private Const DT_WORD_ELLIPSIS As Long = &H40000
Public Declare Function DestroyCaret Lib "user32" () As Long
Public Declare Function CreateCaret Lib "user32" (ByVal hWnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function ShowCaret Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function SetCaretPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function HideCaret Lib "user32" (ByVal hWnd As Long) As Long
Const dv = 0.877551020408163
Public QUERYLIST As String
Public LASTQUERYLIST As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public releasemouse As Boolean
Public LASTPROG$
Public NORUN1 As Boolean
Public UseEnter As Boolean
Public dv20 As Single  ' = 24.5
Public dv15 As Long
Public mHelp As Boolean
Public abt As Boolean
Public vH_title$
Public vH_doc$
Public vH_x As Long
Public vH_y As Long
Public ttl As Boolean
Public Const SRCCOPY = &HCC0020
Public Release As Boolean
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function RoundRect Lib "gdi32" (ByVal hDC As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal y3 As Long) As Long
Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ScrollDC Lib "user32" (ByVal hDC As Long, ByVal dX As Long, ByVal dY As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
Public LastErName As String
Public LastErNameGR As String
Public LastErNum As Long
Public LastErNum1 As Long, LastErNum2 As Long
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal addr As Long, ByVal NewVal As Byte)

Type POINTAPI
        X As Long
        Y As Long
End Type
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32W" (ByVal hDC As Long, ByVal lpsz As Long, ByVal cbString As Long, lpSize As POINTAPI) As Long

Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function PaintDesktop Lib "user32" (ByVal hDC As Long) As Long
Declare Function SelectClipPath Lib "gdi32" (ByVal hDC As Long, ByVal iMode As Long) As Long
  Public Const RGN_AND = 1
    Public Const RGN_COPY = 5
    Public Const RGN_DIFF = 4
    Public Const RGN_MAX = RGN_COPY
    Public Const RGN_MIN = RGN_AND
    Public Const RGN_OR = 2
    Public Const RGN_XOR = 3
Declare Function StrokePath Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Declare Function PolyBezier Lib "gdi32.dll" (ByVal hDC As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long
Declare Function PolyBezierTo Lib "gdi32.dll" (ByVal hDC As Long, lppt As POINTAPI, ByVal cCount As Long) As Long
Declare Function BeginPath Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function EndPath Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function FillPath Lib "gdi32" (ByVal hDC As Long) As Long
Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hDC As Long) As Long

Public PLG() As POINTAPI
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Public lckfrm As Long
Public NERR As Boolean
Public moux As Single, mouy As Single, MOUB As Long
Public mouxb As Single, mouyb As Single, MOUBb As Long
Public vol As Long
Public MYFONT As String, myCharSet As Integer, myBold As Boolean
Public FFONT As String

Public escok As Boolean
Public NOEDIT As Boolean
Public CancelEDIT As Boolean

Global Const HWND_TOP = 0

Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_SHOWWINDOW = &H40
Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long)
Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Public Const FLOODFILLSURFACE = 1
Public Const FLOODFILLBORDER = 0

Public avifile As String
Public BigPi As Variant
Public Const Pi = 3.14159265358979
Public Const PI2 = 6.28318530717958
Public EditTabWidth As Long, ReportTabWidth As Long
Public Result As Long
Public mcd As String
Public NOEXECUTION As Boolean, RoundDouble As Boolean
Public QRY As Boolean, GFQRY As Boolean
Public nomore As Boolean
Private Declare Function CallWindowProc _
 Lib "user32.dll" Alias "CallWindowProcW" ( _
 ByVal lpPrevWndFunc As Long, _
 ByVal hWnd As Long, _
 ByVal Msg As Long, _
 ByVal wParam As Long, _
 ByVal lParam As Long) As Long

'== MCI Wave API Declarations ================================================
Public ExTarget As Boolean
''Public pageframe As Long
''Public basicpageframe As Long

Public q() As target
Public Targets As Boolean
Public SzOne As Single
Public PenOne As Long
Public NoAction As Boolean
Public StartLine As Boolean
Public www&
Public WWX&, ins&
Public INK$, MINK$
Public MKEY$
Public Type target
    Comm As String
    Tag As String ' specified by id
    Id As Long ' function id
    ' THIS IS POINTS AT CHARACTER RESOLUTION
    SZ As Single
    ' SO WE NEED SZ
    Lx As Long
    ly As Long
    tx As Long
    ty As Long
    back As Long 'background fill color' -1 no fill
    fore As Long 'border line ' -1 no line
    Enable As Boolean ' in use
    pen As Long
    layer As Long
    Xt As Long
    Yt As Long
    sUAddTwipsTop As Long
End Type

Public here$, PaperOne As Long
Const PROOF_QUALITY = 2
Const NONANTIALIASED_QUALITY = 3
Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
' lfFaceName(LF_FACESIZE) As Byte 'THIS WAS DEFINED IN API-CHANGES MY OWN
  lfFaceName As String * 33
End Type
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal y3 As Long) As Long

Private Declare Function PathToRegion Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
' OCTOBER 2000
Public dstyle As Long
' Jule 2001
Const DC_ACTIVE = &H1
Const DC_ICON = &H4
Const DC_TEXT = &H8
Const BDR_SUNKENOUTER = &H2
Const BDR_RAISEDINNER = &H4
Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Const BF_BOTTOM = &H8
Const BF_LEFT = &H1
Const BF_RIGHT = &H4
Const BF_TOP = &H2
Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Const DFC_BUTTON = 4
Const DFC_POPUPMENU = 5            'Only Win98/2000 !!
Const DFCS_BUTTON3STATE = &H10
Const DC_GRADIENT = &H20          'Only Win98/2000 !!

Private Declare Function DrawCaption Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long, pcRect As RECT, ByVal un As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExW" (ByVal hDC As Long, ByVal lpsz As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long, ByVal lpDrawTextParams As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
''API declarations
' old api..
Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceW" (ByVal lpFileName As Long) As Long
Private Declare Function RemoveFontResource Lib "gdi32" Alias "RemoveFontResourceW" (ByVal lpFileName As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" _
    (ByVal vKey As Long) As Long
Public TextEditLineHeight As Long
Public LablelEditLineHeight As Long
Private Const Utf8CodePage As Long = 65001
Public Type DRAWTEXTPARAMS
     cbSize As Long
     iTabLength As Long
     iLeftMargin As Long
     iRightMargin As Long
     uiLengthDrawn As Long
End Type
Public tParam As DRAWTEXTPARAMS
Function CheckItemType(bstackstr As basetask, v As Variant, a$, r$, Optional ByVal wasarr As Boolean = False) As Boolean
Dim useHandler As mHandler, fastcol As FastCollection, pppp As mArray, w1 As Long, p As Variant, s$
CheckItemType = True
Dim vv
If MyIsObject(v) Then
    Set vv = v
Else
   r$ = Typename(v)
   CheckItemType = FastSymbol(a$, ")")
   Exit Function
End If
againtype:
        r$ = Typename(vv)
        If r$ = "mHandler" Then
            Set useHandler = vv
            Select Case useHandler.t1
            Case 1
                Set fastcol = useHandler.objref
                If FastSymbol(a$, ",") Then
contHandler:
                    If IsExp(bstackstr, a$, p) Then
                        If Not fastcol.Find(p) Then GoTo keynotexist
                            If fastcol.IsObj Then
                                Set vv = fastcol.ValueObj
                                GoTo againtype
                            Else
                                wasarr = True
                                GoTo checkit
                            End If
                        ElseIf IsStrExp(bstackstr, a$, s$) Then
                            If fastcol.IsObj Then
                                Set vv = fastcol.ValueObj
                                GoTo againtype
                            Else
                            If fastcol.StructLen > 0 Then GoTo checkit
                                r$ = Typename(fastcol.Value)
                            End If
                        Else
                            MissParam a$
                            CheckItemType = False
                            Exit Function
keynotexist:
                            indexout a$
                            CheckItemType = False
                            Exit Function
                    End If
                ElseIf FastSymbol(a$, ")(", , 2) Then
                    GoTo contHandler
   
                Else
                    ' new
checkit:
                    If fastcol.StructLen > 0 Then
                                    Select Case fastcol.sValue
                                    Case Is < 0
                                        r$ = "String"
                                    Case 1
                                        r$ = "Byte"
                                    Case 2
                                        r$ = "Integer"
                                    Case 4
                                        r$ = "Long"
                                    Case 8
                                        r$ = "LongLong"  ' can be double or two longs or ...etc
                                    Case Else
                                        r$ = "Structure"
                                    End Select
                    ElseIf wasarr Then
                    r$ = Typename(fastcol.Value)
                    
                    ElseIf FastSymbol(a$, "!") Then
                        If fastcol.IsQueue Then
                            r$ = "Queue"
                        Else
                            r$ = "List"
                        End If
                    Else
                        r$ = "Inventory"
                    End If
                End If
            Case 2
                r$ = "Buffer"
                
                
            Case 3
                w1 = useHandler.indirect
                If w1 > -1 And w1 <= var2used Then
                                r$ = Typename(var(w1))
                                If r$ = "mHandler" Then Set vv = var(w1): GoTo againtype
                    Else
                            r$ = Typename(useHandler.objref)
                                       If FastSymbol(a$, ",") Then
contarr0:
                                        If r$ = "mArray" Then

                                            Set pppp = useHandler.objref
                                                If IsExp(bstackstr, a$, p) Then
                                                   pppp.index = p
                                                    If MyIsObject(pppp.Value) Then
                                                         Set vv = pppp.Value
                                                         wasarr = False
                                                         GoTo againtype
                                                    Else
                                                        r$ = Typename(pppp.Value)
                                                    End If
                                                Else
                                                MissParam a$
                                                CheckItemType = False
                                                Exit Function
                                            End If
                                        ElseIf FastSymbol(a$, ")(", , 2) Then
                                        GoTo contarr0
            
                                        Else
                                                MyEr "Use STACKTYPE$() ", " Χρησιμοποίησε την ΣΩΡΟΥΤΥΠΟΣ$()"
                                                CheckItemType = False
                                                Exit Function
                                        End If
                                        
                                        End If
                                        
                                        End If
                                    
            Case 4
                    r$ = useHandler.objref.EnumName
            Case Else
                r$ = Typename(vv.objref)
            End Select
        ElseIf Typename(vv) = "PropReference" Then
            r$ = Typename$(vv.Value)
        ElseIf Typename(vv) = "mArray" Then
        
         If FastSymbol(a$, ",") Then
contarr1:
            Set pppp = vv

            If IsExp(bstackstr, a$, p) Then

                pppp.index = p
                If MyIsObject(pppp.Value) Then
                     Set vv = pppp.Value
                     wasarr = False
                     GoTo againtype
                Else
                    r$ = Typename(pppp.Value)
                End If
                Else
                MissParam a$
                CheckItemType = False
                Exit Function
                End If
            ElseIf FastSymbol(a$, ")(", , 2) Then
          GoTo contarr1
            Else
            r$ = "mArray"
            End If
        ElseIf Typename(vv) = "lambda" Then
        If FastSymbol(a$, ")(", , 2) Then
            Set bstackstr.lastobj = vv
            Set bstackstr.lastpointer = Nothing
            s$ = BlockParam(a$)
            If Len(s$) > 0 Then Mid$(a$, 1, Len(s$)) = space$(Len(s$))
            s$ = s$ + ")"
            If CallLambdaASAP(bstackstr, s$, p, False) Then
                If bstackstr.lastobj Is Nothing Then
                    r$ = Typename$(p)
                Else
                    Set vv = bstackstr.lastobj
                    Set bstackstr.lastobj = Nothing
                    GoTo againtype
                End If
            Else
            Exit Function
            End If
            Else
            r$ = "lambda"
            End If
        ElseIf Typename(vv) = "Group" Then
        If FastSymbol(a$, ")(", , 2) Then
            s$ = BlockParam(a$)
            If Len(s$) > 0 Then Mid$(a$, 1, Len(s$)) = space$(Len(s$))
            If FastSymbol(s$, "@") Then
            s$ = NLtrim(s$)
                    If Len(s$) > 0 Then
                        Set pppp = New mArray
                        pppp.Arr = False
                        Set pppp.GroupRef = vv
                        Set vv = bstackstr.soros
                        Set bstackstr.Sorosref = New mStiva
                        
                        SpeedGroup bstackstr, pppp, "FOR", "", "{Push type$(" + String$(bstackstr.ForLevel + 1, ".") + s$ + ")}", -2
                        If bstackstr.soros.IsEmpty Then
                           Set bstackstr.Sorosref = vv
                           Exit Function
                        Else
                           r$ = bstackstr.soros.PopStr
                        End If
                        Set bstackstr.Sorosref = vv
                     Else
                        SyntaxError
                        Exit Function
                End If
                ' check member
            ElseIf vv.IamApointer Then
                r$ = "Group"   ' not decide yet
            ElseIf vv.HasStrValue Then
                r$ = "String"
            ElseIf vv.HasValue Then
            Set pppp = New mArray
            pppp.Arr = False
            Set pppp.GroupRef = vv
             If SpeedGroup(bstackstr, pppp, "VAL", "", s$ + ")", -2) = 1 Then
                If bstackstr.lastobj Is Nothing Then
                    r$ = Typename$(bstackstr.LastValue)
                Else
                    Set vv = bstackstr.lastobj
                    Set bstackstr.lastobj = Nothing
                    GoTo againtype
                End If

             End If

            End If
        End If
        End If
        Set bstackstr.lastobj = Nothing
        Set bstackstr.lastpointer = Nothing
        While FastSymbol(a$, "!")
        Wend
        CheckItemType = FastSymbol(a$, ")", True)
End Function

Public Function Utf16toUtf8(s As String) As Byte()
    ' code from vbforum
    ' UTF-8 returned to VB6 as a byte array (zero based) because it's pretty useless to VB6 as anything else.
    Dim iLen As Long
    Dim bbBuf() As Byte
    '
    iLen = WideCharToMultiByte(Utf8CodePage, 0, StrPtr(s), Len(s), 0, 0, 0, 0)
    ReDim bbBuf(0 To iLen - 1) ' Will be initialized as all &h00.
    iLen = WideCharToMultiByte(Utf8CodePage, 0, StrPtr(s), Len(s), VarPtr(bbBuf(0)), iLen, 0, 0)
    Utf16toUtf8 = bbBuf
End Function
Public Function KeyPressedLong(ByVal VirtKeyCode As Long) As Long
On Error GoTo KEXIT
If Not Screen.ActiveForm Is Nothing Then
If GetForegroundWindow = Screen.ActiveForm.hWnd Then
KeyPressedLong = GetAsyncKeyState(VirtKeyCode)
End If
End If
KEXIT:
End Function
Public Function KeyPressed2(ByVal VirtKeyCode As Long, ByVal VirtKeyCode2 As Long) As Boolean
On Error GoTo KEXIT
If Not Screen.ActiveForm Is Nothing Then
If GetForegroundWindow = Screen.ActiveForm.hWnd Then
KeyPressed2 = CBool((GetAsyncKeyState(VirtKeyCode) And &H8000&) = &H8000&) And CBool((GetAsyncKeyState(VirtKeyCode2) And &H8000&) = &H8000&)
End If
End If
KEXIT:
End Function
Public Function KeyPressed(ByVal VirtKeyCode As Long) As Boolean
On Error GoTo KEXIT
If Not Screen.ActiveForm Is Nothing Then
If GetForegroundWindow = Screen.ActiveForm.hWnd Then
KeyPressed = CBool((GetAsyncKeyState(VirtKeyCode) And &H8000&) = &H8000&)
End If
End If
KEXIT:
End Function
Public Function mouse2() As Long
On Error GoTo MEXIT
If Not Screen.ActiveForm Is Nothing Then
If GetForegroundWindow = Screen.ActiveForm.hWnd Then

mouse2 = (UINT(GetAsyncKeyState((1))) And &HFF) + (UINT(GetAsyncKeyState((2))) And &HFF) * 2 + (UINT(GetAsyncKeyState((4))) And &HFF) * 4
End If
End If
MEXIT:
End Function
Public Function mouse() As Long
On Error GoTo MEXIT
If Not Screen.ActiveForm Is Nothing Then
If GetForegroundWindow = Screen.ActiveForm.hWnd Then
''If Screen.ActiveForm Is Form1 Then If Form1.lockme Then Exit Function

mouse = -1 * CBool((GetAsyncKeyState(1) And &H8000&) = &H8000&) - 2 * CBool((GetAsyncKeyState(2) And &H8000&) = &H8000&) - 4 * CBool((GetAsyncKeyState(4) And &H8000&) = &H8000&)
End If
End If
MEXIT:
End Function

Public Function MOUSEX(Optional offset As Long = 0) As Long
Static X As Long
On Error GoTo MOUSEX
Dim tp As POINTAPI
MOUSEX = X
If Not Screen.ActiveForm Is Nothing Then
If GetForegroundWindow = Screen.ActiveForm.hWnd Then
   GetCursorPos tp
   X = tp.X * dv15 - offset
  MOUSEX = X
  End If
End If
MOUSEX:
End Function
Public Function MOUSEY(Optional offset As Long = 0) As Long
Static Y As Long
On Error GoTo MOUSEY
Dim tp As POINTAPI
MOUSEY = Y
If Not Screen.ActiveForm Is Nothing Then
If GetForegroundWindow = Screen.ActiveForm.hWnd Then
   GetCursorPos tp
   Y = tp.Y * dv15 - offset
   MOUSEY = Y
  End If
End If
MOUSEY:
End Function
Public Sub OnlyInAGroup()
    MyEr "Only in a group", "Μόνο σε μια ομάδα"
End Sub
Public Sub WrongOperator()
MyEr "Wrong operator", "λάθος τελεστής"
End Sub
Public Sub NoOperatorForThatObject(ss$)
If ss$ = "g" Then ss$ = "<="
    MyEr "Object not support operator " + ss$, "Το αντικείμενο δεν υποστηρίζει το τελεστή " + ss$
End Sub
Public Sub NoStackObjectFound(a$)
    MyErMacro a$, "Not stack object found", "Δεν βρήκα αντικείμενο σωρού"
End Sub
Public Sub NoStackObjectToMerge()
    MyEr "Not stack object to merge", "Δεν βρήκα αντικείμενο σωρού να ενώσω"
End Sub
Public Sub Unsignlongnegative(a$)
    MyErMacro a$, "Unsigned long can't be negative", "Ο ακέραιος χωρίς προσημο δεν μπορεί να είναι αρνητικός"
End Sub
Public Sub Unsignlongfailed(a$)
MyErMacro a$, "Unsigned long to sign failed", "Η μετατροπή ακέραιου χωρίς πρόσημο σε ακέραιο με πρόσημο, απέτυχε"
End Sub
Public Sub NoProperObject()
MyEr "This object not supported", "Αυτό το αντικείμενο δεν υποστηρίζεται"
End Sub

Public Sub MyEr(er$, ergr$)
If Left$(LastErName, 1) = Chr(0) Then
    LastErName = vbNullString
    LastErNameGR = vbNullString
End If
If er$ = vbNullString Then
LastErNum = 0
LastErNum1 = 0
LastErName = vbNullString
LastErNameGR = vbNullString
Else
er$ = Split(er$, ChrW(&H1FFF))(0)
ergr$ = Split(ergr$, ChrW(&H1FFF))(0)
If rinstr(er$, " ") = 0 Then
LastErNum = 1001
Else

LastErNum = val(" " & Mid$(er$, rinstr(er$, " ")) + ".0")
End If
If LastErNum = 0 Then LastErNum = -1 ': Debug.Print er$, ergr$: Stop
LastErNum1 = LastErNum

If InStr("*" + LastErName, NLtrim$(er$)) = 0 Then
LastErName = RTrim$(LastErName) & " " & NLtrim$(er$)
LastErNameGR = RTrim$(LastErNameGR) & " " & NLtrim$(ergr$)
End If
End If
End Sub
Sub UnknownVariable1(a$, v$)
Dim i As Long
i = rinstr(v$, "." + ChrW(8191))
If i > 0 Then
    i = rinstr(v$, ".")
    MyErMacro a$, "Unknown Variable " & Mid$(v$, i), "’γνωστη μεταβλητή " & Mid$(v$, i)
Else
    i = rinstr(v$, "].")
    If i > 0 Then
        MyErMacro a$, "Unknown Variable " & Mid$(v$, i + 2), "’γνωστη μεταβλητή " & Mid$(v$, i + 2)
    Else
        i = rinstr(v$, ChrW(8191))
    If i > 0 Then
        i = InStr(i + 1, v$, ".")
        If i > 0 Then
            MyErMacro a$, "Unknown Variable " & Mid$(v$, i + 1), "’γνωστη μεταβλητή " & Mid$(v$, i + 1)
        Else
            MyErMacro a$, "Unknown Variable", "’γνωστη μεταβλητή"
        End If
    Else
        MyErMacro a$, "Unknown Variable " & v$, "’γνωστη μεταβλητή " & v$
    End If
    End If
End If

End Sub
Sub UnknownProperty1(a$, v$)
MyErMacro a$, "Unknown Property " & v$, "’γνωστη ιδιότητα " & v$
End Sub
Sub UnknownMethod1(a$, v$)
 MyErMacro a$, "unknown method/array  " & v$, "’γνωστη μέθοδος/πίνακας " & v$
End Sub
Sub UnknownFunction1(a$, v$)
 MyErMacro a$, "unknown function/array " & v$, "’γνωστη συνάρτηση/πίνακας " & v$
End Sub

Sub InternalError()
 MyEr "Internal Error", "Εσωτερικό Πρόβλημα"
End Sub
Public Function LoadFont(ByVal FntFileName As String) As Boolean
    Dim FntRC As Long
    If FontList Is Nothing Then
    Set FontList = New FastCollection
    End If
    FntFileName = mylcasefILE(FntFileName)
    If FontList.ExistKey(FntFileName) Then
        LoadFont = True
    Else
        FntRC = AddFontResource(StrPtr(FntFileName))
        If FntRC = 0 Then 'no success
         LoadFont = False
        Else 'success
        FontList.AddKey FntFileName
     LoadFont = True
    End If
        End If
End Function
'FntFileName includes also path
Public Function RemoveFont(ByVal FntFileName As String) As Boolean
     Dim rc As Long, Inc As Integer
     If FontList Is Nothing Then Exit Function
        FntFileName = mylcasefILE(FntFileName)
     If FontList.ExistKey(FntFileName) Then
     Do
       rc = RemoveFontResource(StrPtr(FntFileName))
       Inc = Inc + 1
     Loop Until rc = 0 Or Inc > 10
     If rc = 0 Then
        FontList.Remove (FntFileName)
        RemoveFont = True
     End If
    End If
End Function
Public Sub RemoveAllFonts()
Dim i As Long, FntFileName As String, Inc As Integer, rc As Long
If FontList Is Nothing Then Exit Sub
For i = 0 To FontList.count - 1
    FontList.index = i
    FntFileName = FontList.KeyToString
    Inc = 0
    FntFileName = mylcasefILE(FntFileName)
    Do
      rc = RemoveFontResource(StrPtr(FntFileName))
      Inc = Inc + 1
    Loop Until rc = 0 Or Inc > 10
Next i
Set FontList = Nothing
End Sub



Sub myform(m As Object, X As Long, Y As Long, x1 As Long, y1 As Long, Optional t As Boolean = False, Optional factor As Single = 1)
Dim hRgn As Long
m.Move X, Y, x1, y1
If Int(25 * factor) > 2 Then
m.ScaleMode = vbPixels

hRgn = CreateRoundRectRgn(0, 0, m.ScaleX(x1, 1, 3), m.ScaleY(y1, 1, 3), 25 * factor, 25 * factor)
SetWindowRgn m.hWnd, hRgn, t
DeleteObject hRgn
m.ScaleMode = vbTwips

m.Line (0, 0)-(m.ScaleWidth - dv15, m.ScaleHeight - dv15), m.backcolor, BF
End If
End Sub

Sub MyRect(m As Object, mb As basket, x1 As Long, y1 As Long, way As Long, par As Variant, Optional zoom As Long = 0)
Dim r As RECT, b$
With mb
Dim x0&, y0&, X As Long, Y As Long
GetXYb m, mb, x0&, y0&
X = m.ScaleX(x0& * .Xt - DXP, 1, 3)
Y = m.ScaleY(y0& * .Yt - DYP, 1, 3)
If x1 >= .mx Then x1 = m.ScaleX(m.ScaleWidth, 1, 3) Else x1 = m.ScaleX(x1 * .Xt, 1, 3)
If y1 >= .My Then y1 = m.ScaleY(m.ScaleHeight, 1, 3) Else y1 = m.ScaleY(y1 * .Yt + .Yt, 1, 3)

SetRect r, X + zoom, Y + zoom, x1 - zoom, y1 - zoom
Select Case way
Case 0
DrawEdge m.hDC, r, CLng(par) Mod 256, CLng(par) \ 256
Case 1
DrawCaption m.hWnd, m.hDC, r, CLng(par)
Case 2
DrawEdge m.hDC, r, CLng(par), BF_RECT
Case 3
DrawFocusRect m.hDC, r
Case 4
DrawFrameControl m.hDC, r, DFC_BUTTON, DFCS_BUTTON3STATE
Case 5
b$ = Replace(CStr(par), ChrW(&HFFFFF8FB), ChrW(&H2007))
DrawText m.hDC, StrPtr(b$), Len(CStr(par)), r, DT_CENTER
Case 6
DrawFrameControl m.hDC, r, CLng(par) Mod 256, CLng(par) \ 256
Case Else
k1 = 0
MyDoEvents1 Form1
End Select
LCTbasket m, mb, y0&, x0&
End With
End Sub
Sub MyFill(m As Object, x1 As Long, y1 As Long, way As Long, par As Variant, Optional zoom As Long = 0)
Dim r As RECT, b$
Dim X As Long, Y As Long
With players(GetCode(m))
x1 = .XGRAPH + x1
y1 = .YGRAPH + y1
x1 = m.ScaleX(x1, 1, 3)
y1 = m.ScaleY(y1, 1, 3)
X = m.ScaleX(.XGRAPH, 1, 3)
Y = m.ScaleY(.YGRAPH, 1, 3)
SetRect r, X + zoom, Y + zoom, x1 - zoom, y1 - zoom
Select Case way
Case 0
DrawEdge m.hDC, r, CLng(par) Mod 256, CLng(par) \ 256
Case 1
DrawCaption m.hWnd, m.hDC, r, CLng(par)
Case 2
DrawEdge m.hDC, r, CLng(par), BF_RECT
Case 3
DrawFocusRect m.hDC, r
Case 4
DrawFrameControl m.hDC, r, DFC_BUTTON, DFCS_BUTTON3STATE
Case 5
b$ = Replace(CStr(par), ChrW(&HFFFFF8FB), ChrW(&H2007))
DrawText m.hDC, StrPtr(b$), Len(CStr(par)), r, DT_CENTER
Case 6
DrawFrameControl m.hDC, r, CLng(par) Mod 256, CLng(par) \ 256
Case Else
k1 = 0
MyDoEvents1 Form1
End Select
End With
End Sub
' ***************


Public Sub TextColor(d As Object, tc As Long)
d.ForeColor = tc
End Sub
Public Sub TextColorB(d As Object, mb As basket, tc As Long)
d.ForeColor = tc
mb.mypen = d.ForeColor
End Sub

Public Sub LCTNo(DqQQ As Object, ByVal Y As Long, ByVal X As Long)

''DqQQ.CurrentX = x * Xt
''DqQQ.CurrentY = y * Yt + UAddTwipsTop
''xPos = x
''yPos = y
End Sub

Public Sub LCTbasketCur(DqQQ As Object, mybasket As basket)
With mybasket
DqQQ.CurrentX = .curpos * .Xt
DqQQ.CurrentY = .currow * .Yt + .uMineLineSpace

End With
End Sub
Public Sub LCTbasket(DqQQ As Object, mybasket As basket, ByVal Y As Long, ByVal X As Long)
DqQQ.CurrentX = X * mybasket.Xt
DqQQ.CurrentY = Y * mybasket.Yt + mybasket.uMineLineSpace
mybasket.curpos = X
mybasket.currow = Y
End Sub
Public Sub nomoveLCTC(dqq As Object, mb As basket, Y As Long, X As Long, t&)
Dim oldx&, oldy&
With mb
oldx& = dqq.CurrentX
oldy& = dqq.CurrentY
dqq.DrawMode = vbXorPen
If t& = 1 Then
dqq.Line (X * .Xt, Int(Y * .Yt + .uMineLineSpace))-(X * .Xt + .Xt - DXP, Y * .Yt - .uMineLineSpace + .Yt - DYP), (mycolor(.mypen) Xor dqq.backcolor), BF
Else
dqq.Line (X * .Xt, Int((Y + 1) * .Yt - .uMineLineSpace - .Yt \ 6 - DYP))-(X * .Xt + .Xt - DXP, (Y + 1) * .Yt - .uMineLineSpace - DYP), (mycolor(.mypen) Xor dqq.backcolor), BF
End If
dqq.DrawMode = vbCopyPen
dqq.CurrentX = oldx&
dqq.CurrentY = oldy&
End With
End Sub

Public Sub oldLCTCB(dqq As Object, mb As basket, t&)

dqq.DrawMode = vbXorPen
With mb
'QRY = Not QRY
If IsWine Then
If t& = 1 Then
dqq.Line (.curpos * .Xt, .currow * .Yt + .uMineLineSpace)-(.curpos * .Xt + .Xt, .currow * .Yt - .uMineLineSpace + .Yt), (mycolor(.mypen) Xor dqq.backcolor), BF
Else
dqq.Line (.curpos * .Xt, (dqq.ScaleY((.currow + 1) * .Yt - .uMineLineSpace, 1, 3) - .Yt \ DYP \ 6 - 1) * DYP)-(.curpos * .Xt + .Xt - DXP, (.currow + 1) * .Yt - .uMineLineSpace - DYP), (mycolor(.mypen) Xor dqq.backcolor), BF

End If
Else
If t& = 1 Then
dqq.Line (.curpos * .Xt, .currow * .Yt + .uMineLineSpace)-(.curpos * .Xt + .Xt, .currow * .Yt - .uMineLineSpace + .Yt), &HFFFFFF, BF
Else
dqq.Line (.curpos * .Xt, (dqq.ScaleY((.currow + 1) * .Yt - .uMineLineSpace, 1, 3) - .Yt \ DYP \ 6 - 1) * DYP)-(.curpos * .Xt + .Xt - DXP, (.currow + 1) * .Yt - .uMineLineSpace - DYP), &HFFFFFF, BF
End If
End If
End With
dqq.DrawMode = vbCopyPen
End Sub
Public Sub LCTCnew(dqq As Object, mb As basket, Y As Long, X As Long)
DestroyCaret
With mb
CreateCaret dqq.hWnd, 0, dqq.ScaleX(.Xt, 1, 3), dqq.ScaleY((.Yt - .uMineLineSpace * 2) * 0.2, 1, 3)
SetCaretPos dqq.ScaleX(X * .Xt, 1, 3), dqq.ScaleY((Y + 0.8) * .Yt, 1, 3)
End With
End Sub
Public Sub LCTCB(dqq As Object, mb As basket, t&)
With mb
If t& = -1 Or Not Form1.ActiveControl Is dqq Then
        If Not t& = -1 Then
        
        Else
        If Form1.ActiveControl Is Nothing Then
        Else
            CreateCaret Form1.ActiveControl.hWnd, 0, -1, 0
            End If
            CreateCaret dqq.hWnd, 0, -1, 0
        End If
        Exit Sub
End If

If t& = 1 Then
       ' CreateCaret dqq.hWnd, 0, dqq.ScaleX(.Xt, 1, 3), dqq.ScaleY((.Yt - .uMineLineSpace * 2), 1, 3)
       CreateCaret dqq.hWnd, 0, dqq.ScaleX(.Xt, 1, 3), dqq.ScaleY(.Yt - .uMineLineSpace * 2, 1, 3)
        SetCaretPos dqq.ScaleX(.curpos * .Xt, 1, 3), dqq.ScaleY(.currow * .Yt + .uMineLineSpace, 1, 3)
        On Error Resume Next
        If Not extreme Then If INK$ = vbNullString Then dqq.Refresh
Else
    CreateCaret dqq.hWnd, 0, dqq.ScaleX(.Xt, 1, 3), .Yt \ DYP \ 6 + 1
        
            SetCaretPos dqq.ScaleX(.curpos * .Xt, 1, 3), dqq.ScaleY((.currow + 1) * .Yt - .uMineLineSpace, 1, 3) - .Yt \ DYP \ 6 - 1
        On Error Resume Next
        If Not extreme Then If INK$ = vbNullString Then dqq.Refresh
End If
dqq.DrawMode = vbCopyPen
dqq.CurrentX = .curpos * .Xt
dqq.CurrentY = .currow * .Yt + .uMineLineSpace
End With
End Sub
Public Sub SetDouble(dq As Object)

SetTextSZ dq, players(GetCode(dq)).SZ, 2


End Sub

Public Sub SetNormal(dq As Object)
SetTextSZ dq, players(GetCode(dq)).SZ, 1
End Sub


Sub BOXbasket(dqq As Object, mybasket As basket, b$, c As Long)
With mybasket
    dqq.Line (.X * .Xt - DXP, .Y * .Yt - DYP)-((.X + Len(b$)) * .Xt, .Y * .Yt + .Yt), mycolor(c), B
End With
End Sub

Sub BoxBigNew(dqq As Object, mb As basket, x1&, y1&, c As Long)
With mb
dqq.Line (.curpos * .Xt - DXP, .currow * .Yt - DYP)-(x1& * .Xt - DXP + .Xt, y1& * .Yt + .Yt - DYP), mycolor(c), B
End With

End Sub
Sub CircleBig(dqq As Object, mb As basket, x1&, y1&, c As Long, el As Boolean)
Dim X&, Y&
With mb
X& = .curpos
Y& = .currow
dqq.FillColor = mycolor(c)
dqq.fillstyle = vbFSSolid
If el Then
dqq.Circle (((X& + x1& + 1) / 2 * .Xt) - DXP, ((Y& + y1& + 1) / 2 * .Yt) - DYP), RMAX((x1& - X& + 1) * .Xt, (y1& - Y& + 1) * .Yt) / 2 - DYP, mycolor(c), , , ((y1& - Y& + 1) * .Yt - DYP) / ((x1& - X& + 1) * .Xt - DXP)
Else
dqq.Circle (((X& + x1& + 1) / 2 * .Xt) - DXP, ((Y& + y1& + 1) / 2 * .Yt) - DYP), (RMIN((x1& - X& + 1) * .Xt, (y1& - Y& + 1) * .Yt) / 2 - DYP), mycolor(c)

End If
dqq.fillstyle = vbFSTransparent
End With
End Sub
Sub Ffill(dqq As Object, x1 As Long, y1 As Long, c As Long, v As Boolean)
Dim osm
With players(GetCode(dqq))
osm = dqq.ScaleMode
dqq.ScaleMode = vbPixels
dqq.FillColor = mycolor(c)
dqq.fillstyle = vbFSSolid
If v Then
ExtFloodFill dqq.hDC, dqq.ScaleX(x1, 1, 3), dqq.ScaleY(y1, 1, 3), dqq.Point(dqq.ScaleX(x1, 1, 3), dqq.ScaleY(y1, 1, 3)), FLOODFILLSURFACE
Else
ExtFloodFill dqq.hDC, dqq.ScaleX(x1, 1, 3), dqq.ScaleY(y1, 1, 3), mycolor(.mypen), FLOODFILLBORDER
End If
dqq.ScaleMode = osm
dqq.fillstyle = vbFSTransparent
End With
'LCT Dqq, y&, x&
End Sub

Sub BoxColorNew(dqq As Object, mb As basket, x1&, y1&, c As Long)
Dim addpixels As Long
With mb
If InternalLeadingSpace() = 0 And .MineLineSpace = 0 Then
addpixels = 0
Else
addpixels = 2
End If

dqq.Line (.curpos * .Xt, .currow * .Yt)-(x1& * .Xt + .Xt - 2 * DXP, y1& * .Yt + .Yt - addpixels * DYP), mycolor(c), BF
End With
End Sub
Sub BoxImage(d1 As Object, mb As basket, x1&, y1&, F As String, df&, s As Boolean)
'
Dim p As Picture, scl As Double, x2&, dib As Object, aPic As StdPicture

If df& > 0 Then
df& = df& * DXP '* 20

Else

df& = 0
End If
With mb
x1& = .curpos + x1& - 1
x2& = x1&
y1& = .currow + y1& - 1
On Error Resume Next
 If (Left$(F$, 4) = "cDIB" And Len(F$) > 12) Then
   Set dib = New cDIBSection
  If Not cDib(F$, dib) Then
    dib.Create x1&, y1&
    dib.Cls d1.backcolor
  End If
      Set p = dib.Picture
    Set dib = Nothing
 Else
        If ExtractType(F, 0) = vbNullString Then
        F = F + ".bmp"
        End If
        FixPath F
        
    If CFname(F) <> "" Then
    F = CFname(F)
    Set aPic = LoadMyPicture(GetDosPath(F$))
     If aPic Is Nothing Then Exit Sub
    Set p = aPic
                                            

    Else
    Set dib = New cDIBSection
    dib.Create x1&, y1&
    dib.Cls d1.backcolor
    Set p = dib.Picture
    Set dib = Nothing
    End If
End If

If Err.Number > 0 Then Exit Sub

If s Then
scl = (y1& - .currow + 1) * .Yt - df&
If p.Type = vbPicTypeBitmap Then
d1.PaintPicture p, .curpos * .Xt, .currow * .Yt, (x1& - .curpos + 1) * .Xt - df&, scl, , , , , vbSrcCopy
Else
d1.PaintPicture p, .curpos * .Xt, .currow * .Yt, (x1& - .curpos + 1) * .Xt - df&, scl
End If
Else
scl = p.Height * ((x1& - .curpos + 1) * .Xt - df&) / p.Width
If p.Type = vbPicTypeBitmap Then
d1.PaintPicture p, .curpos * .Xt, .currow * .Yt, (x1& - .curpos + 1) * .Xt - df&, scl, , , , , vbSrcCopy
Else
d1.PaintPicture p, .curpos * .Xt, .currow * .Yt, (x1& - .curpos + 1) * .Xt - df&, scl
End If
End If
y1& = -Int(-((scl) / .Yt))
Set p = Nothing
''LCT d1, .currow, .curpos
End With
End Sub

Sub sprite(bstack As basetask, ByVal F As String, rst As String)

On Error GoTo SPerror
Dim d1 As Object, amask$, aPic As StdPicture
Set d1 = bstack.Owner
Dim raster As New cDIBSection
Dim p As Double, i As Long, ROT As Double, sp As Double
Dim Pcw As Long, Pch As Long, blend As Double, NoUseBack As Boolean

If Not cDib(F, raster) Then
    If CFname(F) <> "" Then
        F = CFname(F)
        Set aPic = LoadMyPicture(GetDosPath(F$))
        If aPic Is Nothing Then Exit Sub
        raster.CreateFromPicture aPic
        If raster.bitsPerPixel <> 24 Then
            Conv24 raster
        Else
            CheckOrientation raster, F
        End If
    Else
        
        BACKSPRITE = vbNullString
        Exit Sub
    End If
End If
If raster.Width = 0 Then
    BACKSPRITE = vbNullString
    Set raster = Nothing
    Set d1 = Nothing
    Exit Sub
End If
i = -1
sp = 100!
blend = 100!
If FastSymbol(rst$, ",") Then
    If IsExp(bstack, rst$, p, , True) Then i = CLng(p) Else i = -players(GetCode(d1)).Paper
    If FastSymbol(rst$, ",") Then
        If IsExp(bstack, rst$, p, , True) Then ROT = p
        If FastSymbol(rst$, ",") Then
            If Not IsExp(bstack, rst$, sp) Then sp = 100!
            If FastSymbol(rst$, ",") Then
                If IsExp(bstack, rst$, blend) Then
                    blend = Abs(Int(blend)) Mod 101
                    If FastSymbol(rst$, ",") Then GoTo cont0
                ElseIf IsStrExp(bstack, rst$, amask$) Then
                    blend = 100!
                    If FastSymbol(rst$, ",") Then GoTo cont0
                ElseIf FastSymbol(rst$, ",") Then
                blend = 100!
cont0:
                    If Not IsExp(bstack, rst$, p, , True) Then
                            MyEr "missing parameter", "λείπει παράμετρος"
                            Exit Sub
                    End If
                    NoUseBack = CBool(p)
                Else
                    MyEr "missing parameter", "λείπει παράμετρος"
                End If
                
                
            End If
            End If
        End If
Else
        Pcw = raster.Width \ 2
        Pch = raster.Height \ 2
        With players(GetCode(d1))
        raster.PaintPicture d1.hDC, Int(d1.ScaleX(.XGRAPH, 1, 3) - Pcw), Int(d1.ScaleX(.YGRAPH, 1, 3) - Pch)
        End With
    GoTo cont1
End If
If sp <= 0 Then sp = 0
If i > 0 Then i = QBColor(i) Else i = -i
RotateDib bstack, raster, ROT, sp, i, NoUseBack, (blend), amask$
Pcw = raster.Width \ 2
Pch = raster.Height \ 2
With players(GetCode(d1))
raster.PaintPicture d1.hDC, Int(d1.ScaleX(.XGRAPH, 1, 3) - Pcw), Int(d1.ScaleX(.YGRAPH, 1, 3) - Pch)
End With
cont1:
If Not bstack.toprinter Then
GdiFlush
End If
Set raster = Nothing
'MyDoEvents1 d1
Set d1 = Nothing
Exit Sub
SPerror:
 BACKSPRITE = vbNullString
Set raster = Nothing
End Sub
Sub spriteGDI(bstack As basetask, rst As String)
Dim NoUseBack As Boolean
If bstack.lastobj Is Nothing Then
err1:
    MyEr "Expecting a memory Buffer", "Περίμενα διάρθρωση μνήμης"
    Exit Sub
End If
If Not TypeOf bstack.lastobj Is mHandler Then GoTo err1
If Not bstack.lastobj.t1 = 2 Then GoTo err1
Dim d1 As Object
Set d1 = bstack.Owner
Dim p, i As Long, mem As MemBlock, blend, sp, ROT As Single
Set mem = bstack.lastobj.objref
i = -1
sp = 100!
blend = 0!
If FastSymbol(rst$, ",") Then
    If IsExp(bstack, rst$, p, , True) Then i = CLng(p) Else i = -players(GetCode(d1)).Paper
    If FastSymbol(rst$, ",") Then
        If IsExp(bstack, rst$, p, , True) Then ROT = p
        If FastSymbol(rst$, ",") Then
            If Not IsExp(bstack, rst$, sp, , True) Then sp = 100!
            If FastSymbol(rst$, ",") Then
                If IsExp(bstack, rst$, blend) Then blend = 100 - Abs(Int(blend)) Mod 101
                If FastSymbol(rst$, ",") Then
                    If Not IsExp(bstack, rst$, p) Then
                        MyEr "missing parameter", "λείπει παράμετρος"
                        Exit Sub
                    End If
                    NoUseBack = Not CBool(p)
                End If
            End If
        End If
    End If
End If
If sp <= 0 Then sp = 0
If i > 0 Then i = QBColor(i) Else i = -i
If Not bstack.toprinter Then
GdiFlush
End If
mem.DrawSpriteToHdc bstack, NoUseBack, ROT, (sp), (blend), i

'MyDoEvents1 d1
Set d1 = Nothing
Set bstack.lastobj = Nothing
Exit Sub
SPerror:
Set bstack.lastobj = Nothing
 BACKSPRITE = vbNullString
End Sub

Sub ThumbImage(d1 As Object, x1 As Long, y1 As Long, F As String, border As Long, tpp As Long, ttl1$)
On Error Resume Next
With players(GetCode(d1))
If Left$(F, 4) = "cDIB" And Len(F) > 12 Then
Dim ph As New cDIBSection
If cDib(F, ph) Then
ph.ThumbnailPartPaint d1, x1 / tpp, y1 / tpp, 0, 0, border <> 0, , ttl1$, .XGRAPH / tpp, .YGRAPH / tpp
End If
End If
End With
End Sub
Sub ThumbImageDib(d1 As Object, x1 As Long, y1 As Long, ph As Object, border As Long, tpp As Long, ttl1$)
On Error Resume Next
Dim pointer2dib As cDIBSection
Set pointer2dib = ph
With players(GetCode(d1))
    pointer2dib.ThumbnailPartPaint d1, x1 / tpp, y1 / tpp, 0, 0, border <> 0, , ttl1$, .XGRAPH / tpp, .YGRAPH / tpp
End With
Set pointer2dib = Nothing
End Sub
Sub SImage(d1 As Object, x1 As Long, y1 As Long, F As String)
'
Dim p As Picture, aPic As StdPicture
On Error Resume Next
With players(GetCode(d1))
If Left$(F, 4) = "cDIB" And Len(F) > 12 Then
Dim ph As New cDIBSection
If cDib(F, ph) Then
If x1 = 0 Then
ph.PaintPicture d1.hDC, CLng(d1.ScaleX(.XGRAPH, 1, 3)), CLng(d1.ScaleX(.YGRAPH, 1, 3))
Exit Sub
Else
If y1 = 0 Then y1 = Abs(ph.Height * x1 / ph.Width)
ph.StretchPictureH d1.hDC, CLng(d1.ScaleX(.XGRAPH, 1, 3)), CLng(d1.ScaleX(.YGRAPH, 1, 3)), CLng(d1.ScaleX(x1, 1, 3)), CLng(d1.ScaleX(y1, 1, 3))
Exit Sub
End If
End If
ElseIf CFname(F) <> "" Then
    F = CFname(F)
     Set aPic = LoadMyPicture(GetDosPath(F$), , , True)
     If aPic Is Nothing Then Exit Sub
     Set p = aPic
Else
If y1 = 0 Then y1 = x1
d1.Line (.XGRAPH, .YGRAPH)-(x1, y1), .Paper, BF
d1.CurrentX = .XGRAPH
d1.CurrentY = .YGRAPH
Exit Sub
End If
If x1 = 0 Then
x1 = d1.ScaleX(p.Width, vbHimetric, vbTwips)

If y1 = 0 Then y1 = p.Height * d1.ScaleX(p.Width, vbHimetric, vbTwips) / p.Width
Else
If y1 = 0 Then y1 = p.Height * x1 / p.Width
End If
If Err.Number > 0 Then Exit Sub

If p.Type = vbPicTypeBitmap Then
d1.PaintPicture p, .XGRAPH, .YGRAPH, x1, y1, , , , , vbSrcCopy
Else
d1.PaintPicture p, .XGRAPH, .YGRAPH, x1, y1
End If
Set p = Nothing
End With
' UpdateWindow d1.hwnd
End Sub
Public Function LoadMyPicture(s1$, Optional useback As Boolean = False, Optional bcolor As Variant = 0&, Optional includeico As Boolean = False) As StdPicture
Dim s As String
Err.clear
   On Error Resume Next
                    If s1$ <> vbNullString Then
                        s$ = UCase(ExtractType(s1$))
                        If LenB(s$) = 0 Then s$ = "Bmp": s1$ = s1$ + ".bmp"
                        Select Case s
                        Case "JPG", "BMP", "WMF", "EMF", "ICO", "DIB"
                        
                           Set LoadMyPicture = LoadPicture(s1$)
                           If Err.Number > 0 Then
                           Err.clear
                           If useback Then
                              Set LoadMyPicture = LoadPictureGDIPlus(s1$, , , bcolor, True)
                           Else
                              Set LoadMyPicture = LoadPictureGDIPlus(s1$, , , , True)
                            End If
                           End If
                           If Err.Number > 0 Then
                           Err.clear
                           
                           Set LoadMyPicture = LoadPicture("")
                           End If
                           If LoadMyPicture Is Nothing Then
                           Set LoadMyPicture = LoadPicture("")
                           End If
                        Case Else
                            If includeico And Not useback Then
                            Set LoadMyPicture = LoadPicture(s1$)
                                If Err.Number > 0 Then
                                    Err.clear
                                    GoTo conthere
                                End If
                            Else
conthere:
                          If useback Then
                              Set LoadMyPicture = LoadPictureGDIPlus(s1$, , , bcolor, True)
                           Else
                              Set LoadMyPicture = LoadPictureGDIPlus(s1$, , , , True)
                            End If
                            End If
                            If Err.Number > 0 Then
                           Err.clear
                          
                           Set LoadMyPicture = LoadPicture("")
                           End If
                           If LoadMyPicture Is Nothing Then
                           Set LoadMyPicture = LoadPicture("")
                           End If
                        End Select
                    End If
                          
End Function

Public Function GetTextWidth(dd As Object, c As String, r As RECT) As Long
' using current.x and current.y to define r


End Function

Public Sub CalcRect(mHdc As Long, c As String, r As RECT)
r.top = 0
r.Left = 0
DrawTextEx mHdc, StrPtr(c), -1, r, DT_CALCRECT Or DT_NOPREFIX Or DT_SINGLELINE Or DT_NOCLIP Or DT_EXPANDTABS Or DT_TABSTOP, VarPtr(tParam)
End Sub

Public Sub PrintLineControlSingle(mHdc As Long, c As String, r As RECT)
    DrawTextEx mHdc, StrPtr(c), -1, r, DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS Or DT_TABSTOP, VarPtr(tParam)
    End Sub
'
Public Sub MyPrintNew(ddd As Object, UAddTwipsTop, s$, Optional cr As Boolean = False, Optional fake As Boolean = False)

Dim nr As RECT, nl As Long, mytop As Long
mytop = ddd.CurrentY
If s$ = vbNullString Then
nr.Left = 0: nr.Right = 0: nr.top = 0: nr.Bottom = 0
CalcRect ddd.hDC, " ", nr
nr.Left = ddd.CurrentX / dv15
nr.Right = nr.Right + nr.Left
nr.top = ddd.CurrentY / dv15
nr.Bottom = nr.top + nr.Bottom
nl = (nr.Bottom + 1) * dv15
If cr Then
ddd.CurrentY = (nr.Bottom + 1) * dv15 + UAddTwipsTop ''2
ddd.CurrentX = 0
Else
ddd.CurrentX = nr.Right * dv15
End If
Else
nr.Left = 0: nr.Right = 0: nr.top = 0: nr.Bottom = 0
CalcRect ddd.hDC, s$, nr
nr.Left = ddd.CurrentX / dv15
nr.Right = nr.Right + nr.Left
nr.top = ddd.CurrentY / dv15
nr.Bottom = nr.top + nr.Bottom
nl = (nr.Bottom + 1) * dv15
If Not fake Then
If nr.Left * dv15 < ddd.Width Then PrintLineControlSingle ddd.hDC, s$, nr
End If
If cr Then
ddd.CurrentY = nl + UAddTwipsTop ''* 2
ddd.CurrentX = 0
Else
ddd.CurrentY = mytop
ddd.CurrentX = nr.Right * dv15
End If
End If

End Sub
Public Sub MyPrint(ddd As Object, s$)
Dim nr As RECT, nl As Long
If s$ = vbNullString Then
    nr.Left = 0: nr.Right = 0: nr.top = 0: nr.Bottom = 0
    CalcRect ddd.hDC, " ", nr
    nr.Left = ddd.CurrentX / dv15
    nr.Right = nr.Right + nr.Left
    nr.top = ddd.CurrentY / dv15
    nr.Bottom = nr.top + nr.Bottom
    nl = (nr.Bottom + 1) * dv15
    ddd.CurrentY = (nr.Bottom + 1) * dv15
    ddd.CurrentX = 0
Else
nr.Left = 0: nr.Right = 0: nr.top = 0: nr.Bottom = 0
CalcRect ddd.hDC, s$, nr
nr.Left = ddd.CurrentX / dv15
nr.Right = nr.Right + nr.Left
nr.top = ddd.CurrentY / dv15
nr.Bottom = nr.top + nr.Bottom
nl = (nr.Bottom + 1) * dv15
If nr.Left * dv15 < ddd.Width Then PrintLineControlSingle ddd.hDC, s$, nr
ddd.CurrentY = nl
ddd.CurrentX = 0
End If
End Sub

Public Function TextWidth(ddd As Object, a$) As Long
Dim nr As RECT
CalcRect ddd.hDC, a$, nr
TextWidth = nr.Right * dv15
End Function
Private Function TextHeight(ddd As Object, a$) As Long
Dim nr As RECT
CalcRect ddd.hDC, a$, nr

TextHeight = nr.Bottom * dv15
End Function

Public Sub PrintLine(dd As Object, c As String, r As RECT)
DrawText dd.hDC, StrPtr(c), -1, r, DT_CENTER
End Sub
Public Sub PrintUnicodeStandardWidthAddXT(dd As Object, c As String, r As RECT)
'Dim m As Long
'm = dd.CurrentX + r.Left

DrawText dd.hDC, StrPtr(c), -1, r, DT_SINGLELINE Or DT_CENTER Or DT_NOPREFIX
'dd.CurrentX = m
End Sub

Public Sub PlainOLD(ddd As Object, mb As basket, ByVal what As String, Optional ONELINE As Boolean = False, Optional nocr As Boolean = False, Optional plusone As Long = 2)
Dim PX As Long, PY As Long, r As Long, p$, c$, LEAVEME As Boolean, nr As RECT, nr2 As RECT
Dim p2 As Long
With mb
p2 = .uMineLineSpace \ dv15 * 2
LEAVEME = False
 PX = .curpos
 PY = .currow
Dim pixX As Long, pixY As Long
pixX = .Xt / dv15
pixY = .Yt / dv15
Dim rTop As Long, rBottom As Long
 With nr
 .Left = PX * pixX
 .Right = .Left + pixX
 .top = PY * pixY + mb.uMineLineSpace \ dv15
 
 .Bottom = .top + pixY - mb.uMineLineSpace \ dv15 * 2
 End With
rTop = PY * pixY
rBottom = rTop + pixY - plusone
Do While Len(what) >= .mx - PX And (.mx - PX) > 0
 p$ = Left$(what, .mx - PX)
 
  With nr2
 .Left = PX * pixX
 
 .Right = (PX + Len(p$)) * pixX + 1
 .top = rTop
 .Bottom = rBottom
 
 End With
 
 If ddd.FontTransparent = False Then
 FillBack ddd.hDC, nr2, ddd.backcolor
 End If
 For r = 0 To Len(p$) - 1
If ONELINE And nocr And PX > .mx Then what = vbNullString: Exit For
 c$ = Mid$(p$, r + 1, 1)

If nounder32(c$) Then ddd.CurrentX = ddd.CurrentX + .Xt: PrintUnicodeStandardWidthAddXT ddd, c$, nr
 With nr
 .Left = .Right
 .Right = .Left + pixX
 End With

  Next r
 LCTbasket ddd, mb, PY, PX + r
   
   
what = Mid$(what, .mx - PX + 1)

If Not ONELINE Then PX = 0

If nocr Then Exit Do Else PY = PY + 1

If PY >= .My And Not ONELINE Then

If ddd.name = "PrinterDocument1" Then
getnextpage
 With nr
 .top = PY * pixY + mb.uMineLineSpace
  .Bottom = .top + pixY - p2
 End With
PY = 1
Else
ScrollUpNew ddd, mb
End If

PY = PY - 1
End If
If ONELINE Then
LCTbasket ddd, mb, PY, PX
LEAVEME = True
Exit Do
Else
 With nr
 .Left = PX * pixX
 .Right = .Left + pixX
 .top = PY * pixY + mb.uMineLineSpace
 .Bottom = .top + pixY - p2
 End With
rTop = PY * pixY
rBottom = rTop + pixY - plusone
End If
Loop
If LEAVEME Then Exit Sub

 If ddd.FontTransparent = False Then
     With nr2
 .Left = PX * pixX
 .Right = (PX + Len(what$)) * pixX + 1
 .top = rTop
 .Bottom = rBottom
 
 End With
 FillBack ddd.hDC, nr2, ddd.backcolor
 End If
 
If what$ <> "" Then
.currow = PY
.curpos = PX
LCTbasketCur ddd, mb
  For r = 0 To Len(what$) - 1
 c$ = Mid$(what$, r + 1, 1)
 If nounder32(c$) Then ddd.CurrentX = ddd.CurrentX + .Xt: PrintUnicodeStandardWidthAddXT ddd, c$, nr
 With nr
 .Left = .Right
 .Right = .Left + pixX
 End With
 
  Next r
  LCTbasket ddd, mb, PY, PX + r
End If

GetXYb ddd, mb, .curpos, .currow
End With
End Sub


Public Sub PlainBaSket(ddd As Object, mybasket As basket, ByVal what As String, Optional ONELINE As Boolean = False, Optional nocr As Boolean = False, Optional plusone As Long = 2, Optional clearline As Boolean = False, Optional processcr As Boolean = False)
Dim PX As Long, PY As Long, r As Long, p$, c$, LEAVEME As Boolean, nr As RECT, nr2 As RECT
Dim p2 As Long, mUAddPixelsTop As Long
Dim pixX As Long, pixY As Long
Dim rTop As Long, rBottom As Long
Dim lenw&, realR&, realstop&, r1 As Long, WHAT1$

Dim a() As Byte, a1() As Byte
'' LEAVEME = False -  NOT NEEDED
again:
nr.Left = 0
realR& = 0

With mybasket
    mUAddPixelsTop = mybasket.uMineLineSpace \ dv15  ' for now
    PX = .curpos
    PY = .currow
    p2 = mUAddPixelsTop * 2
    pixX = .Xt / dv15
    pixY = .Yt / dv15
    With nr
        .Left = PX * pixX
        .Right = .Left + pixX
        .top = PY * pixY + mUAddPixelsTop
         .Bottom = .top + pixY - mUAddPixelsTop * 2
    End With
    
    rTop = PY * pixY
    rBottom = rTop + pixY - plusone
    lenw& = Len(what)
    WHAT1$ = what + " "
     ReDim a(Len(WHAT1$) * 2 + 20)
       ReDim a1(Len(WHAT1$) * 2 + 20)
     
     Dim skip As Boolean
     
     skip = GetStringTypeExW(&HB, 1, StrPtr(WHAT1$), Len(WHAT1$), a(0)) = 0  ' Or IsWine
     skip = GetStringTypeExW(&HB, 4, StrPtr(WHAT1$), Len(WHAT1$), a1(0)) = 0 Or skip
        Do While (lenw& - r) >= .mx - PX And (.mx - PX) > 0
        

        With nr2
                .Left = PX * pixX
                 .Right = mybasket.mx * pixX + 1
                .top = rTop
                .Bottom = rBottom
        End With
        If ddd.FontTransparent = False Then FillBack ddd.hDC, nr2, .Paper
        ddd.CurrentX = PX * .Xt
        ddd.CurrentY = PY * .Yt + .uMineLineSpace
        r1 = .mx - PX - 1 + r
        If ddd.CurrentX = 0 And clearline Then ddd.Line (0&, PY * .Yt)-((.mx - 1) * .Xt + .Xt * 2, (PY) * .Yt + .Yt - 1 * DYP), .Paper, BF
            Do
           '  If ddd.CurrentX = 0 And clearline Then ddd.Line (0&, PY * .Yt)-((.mx - 1) * .Xt + .Xt * 2, (PY) * .Yt + .Yt - 1 * DYP), .Paper, BF

            If ONELINE And nocr And PX > .mx Then what = vbNullString: Exit Do
            c$ = Mid$(WHAT1$, r + 1, 1)
      
                If nounder32(c$) Then
            
                If Not skip Then
                    If a(r * 2 + 2) = 0 And a(r * 2 + 3) <> 0 And a1(r * 2 + 2) < 8 Then
                          Do
                                p$ = Mid$(WHAT1$, r + 2, 1)
                                If ideographs(p$) Then Exit Do
                                If Not nounder32(p$) Then Mid$(WHAT1$, r + 2, 1) = " ": Exit Do
                                c$ = c$ + p$
                                r = r + 1
                                If r >= r1 Then Exit Do
                         Loop Until a(r * 2 + 2) <> 0 Or a(r * 2 + 3) = 0
                     End If
                 End If
                 DrawText ddd.hDC, StrPtr(c$), -1, nr, DT_SINGLELINE Or DT_CENTER Or DT_NOPREFIX
              Else
              
            If c$ = Chr$(7) Then Beep: r = r + 1: realR = realR - 1: GoTo cont0

        If processcr Then
            realR& = realR + 1
            If c$ = ChrW(9) Then
            
            what$ = space$(.Column - (PX + realR - 1) Mod (.Column + 1)) + Mid$(WHAT1$, r + 2)
            r = 0
            .curpos = PX + realR
            If Len(what$) > 0 Then what$ = Mid$(what$, 1, Len(what$) - 1)
            GoTo again
            ElseIf c$ = ChrW(13) Then
               If Mid$(WHAT1$, r + 2, 1) = ChrW(10) Then
                            r = r + 1
               End If
              .curpos = 0
                If PY + 1 >= .My Then
                    If ddd.name = "PrinterDocument1" Then
                        getnextpage
                         With nr
                         .top = PY * pixY + mUAddPixelsTop
                          .Bottom = .top + pixY - p2
                         End With
                        PY = 0
                        .currow = 0
                        Else
                        
                        ScrollUpNew ddd, mybasket
                        End If
                Else
                .currow = PY + 1
                End If
                what$ = Mid$(WHAT1$, r + 2)
               If Len(what$) > 0 Then what$ = Mid$(what$, 1, Len(what$) - 1)
               r = 0
                    GoTo again
        
            ElseIf c$ = ChrW(10) Then
                .curpos = 0
                If PY + 1 = .My Then
                    If ddd.name = "PrinterDocument1" Then
                        getnextpage
                         With nr
                         .top = PY * pixY + mUAddPixelsTop
                          .Bottom = .top + pixY - p2
                         End With
                        PY = 0
                        .currow = 0
                        Else
                        
                        ScrollUpNew ddd, mybasket
                        End If
                    
                    
                Else
                .currow = PY + 1
                End If
                what$ = Mid$(WHAT1$, r + 2)
               If Len(what$) > 0 Then what$ = Mid$(what$, 1, Len(what$) - 1)
               r = 0
                   GoTo again
            End If
        
        End If
  
              
              
              
            End If
           r = r + 1
            With nr
            .Left = .Right
            .Right = .Left + pixX
            End With
cont0:
           ddd.CurrentX = (PX + realR) * .Xt
        realR = realR + 1
     
        If r >= lenw& Then
         r = lenw& + 1
        lenw& = lenw& - 1
        Exit Do
        End If
        If realR > .mx - PX - 1 Then Exit Do
    
         Loop
        .curpos = PX + realR
 
        If Not ONELINE Then PX = 0
        
        If nocr Then Exit Sub Else PY = PY + 1
        
        If PY >= .My And Not ONELINE Then
        
        If ddd.name = "PrinterDocument1" Then
        getnextpage
         With nr
         .top = PY * pixY + mUAddPixelsTop
          .Bottom = .top + pixY - p2
         End With
        PY = 0
        .currow = 0
        Else
        
        ScrollUpNew ddd, mybasket
        End If
        
        PY = PY - 1
       
        End If
        If ONELINE Then

            LEAVEME = True
            Exit Do
        Else
            With nr
               .Left = PX * pixX
               .Right = .Left + pixX
               .top = PY * pixY + mUAddPixelsTop
               .Bottom = .top + pixY - p2
            End With
            rTop = PY * pixY
            rBottom = rTop + pixY - plusone
   

        End If
        realR& = 0
    Loop
    If LEAVEME Then
                With mybasket
                .curpos = PX
                .currow = PY
            End With
    Exit Sub
    End If
     If ddd.FontTransparent = False Then
        With nr2
            .Left = PX * pixX
            .Right = (PX + Len(what$)) * pixX + 1
            .top = rTop
            .Bottom = rBottom
        End With
        FillBack ddd.hDC, nr2, mybasket.Paper
    End If
realR& = 0
    If Len(what$) > r Then

       ddd.CurrentX = PX * .Xt
    
    ddd.CurrentY = PY * .Yt + .uMineLineSpace
        If ddd.CurrentX = 0 And clearline Then ddd.Line (0&, PY * .Yt)-((.mx - 1) * .Xt + .Xt * 2, (PY) * .Yt + .Yt - 1 * DYP), .Paper, BF

r1 = Len(what$) - 1
    For r = r To r1
        c$ = Mid$(WHAT1$, r + 1, 1)
        If nounder32(c$) Then
       ' skip = True
             If Not skip Then
           If a(r * 2 + 2) = 0 And a(r * 2 + 3) <> 0 And a1(r * 2 + 2) < 8 Then
            Do
                p$ = Mid$(WHAT1$, r + 2, 1)
                If ideographs(p$) Then Exit Do
                If Not nounder32(p$) Then Mid$(WHAT1$, r + 2, 1) = " ": Exit Do
                c$ = c$ + p$
                r = r + 1
                If r >= r1 Then Exit Do
            Loop Until a(r * 2 + 2) <> 0 Or a(r * 2 + 3) = 0
            End If
         End If
               
      ddd.CurrentX = ddd.CurrentX + .Xt
        
    Else
        If c$ = Chr$(7) Then Beep: GoTo cont1
        If processcr Then
            realR& = realR + 1
            If c$ = ChrW(9) Then
            
            what$ = space$(.Column - (PX + realR - 1) Mod (.Column + 1)) + Mid$(WHAT1$, r + 2)
            r = 0
            .curpos = PX + realR
            If Len(what$) > 0 Then what$ = Mid$(what$, 1, Len(what$) - 1)
            GoTo again
            ElseIf c$ = ChrW(13) Then
               If Mid$(WHAT1$, r + 2, 1) = ChrW(10) Then
                    r = r + 1
                End If
                    .curpos = 0
                If PY + 1 = .My Then
                    If ddd.name = "PrinterDocument1" Then
                        getnextpage
                         With nr
                         .top = PY * pixY + mUAddPixelsTop
                          .Bottom = .top + pixY - p2
                         End With
                        PY = 0
                        .currow = 0
                        Else
                        
                        ScrollUpNew ddd, mybasket
                        End If
                Else
                .currow = PY + 1
                End If
                what$ = Mid$(WHAT1$, r + 2)
               If Len(what$) > 0 Then what$ = Mid$(what$, 1, Len(what$) - 1)
               r = 0
                    
                    GoTo again
                
            ElseIf c$ = ChrW(10) Then
                .curpos = 0
                If PY + 1 >= .My Then
                    If ddd.name = "PrinterDocument1" Then
                        getnextpage
                         With nr
                         .top = PY * pixY + mUAddPixelsTop
                          .Bottom = .top + pixY - p2
                         End With
                        PY = 0
                        .currow = 0
                        Else
                        
                        ScrollUpNew ddd, mybasket
                        End If
                    
                    
                Else
                .currow = PY + 1
                
                End If
                what$ = Mid$(WHAT1$, r + 2)
               If Len(what$) > 0 Then what$ = Mid$(what$, 1, Len(what$) - 1)
               r = 0
                GoTo again
            End If
        
        End If
    End If
        
    DrawText ddd.hDC, StrPtr(c$), -1, nr, DT_SINGLELINE Or DT_CENTER Or DT_NOPREFIX
    realR& = realR + 1
    With nr
       .Left = .Right
       .Right = .Left + pixX
    End With
cont1:
    Next r
     .curpos = PX + realR
     .currow = PY
     Exit Sub
    End If

  .curpos = PX
 .currow = PY
  End With
End Sub


Public Function nTextY(basestack As basetask, ByVal what As String, ByVal Font As String, ByVal Size As Single, Optional ByVal degree As Double = 0#)
Dim ddd As Object
Set ddd = basestack.Owner
Dim PX As Long, PY As Long, OLDFONT As String, OLDSIZE As String, DE#
Dim F As LOGFONT, hPrevFont As Long, hFont As Long
Dim BFONT As String
Dim prive As Long
prive = GetCode(ddd)
On Error Resume Next
With players(prive)
BFONT = ddd.Font.name
If Font <> "" Then
If Size = 0 Then Size = ddd.FontSize
StoreFont Font, Size, .charset
ddd.Font.charset = 0
ddd.FontSize = 9
ddd.FontName = .FontName
ddd.Font.charset = .charset
ddd.FontSize = Size
Else
Font = .FontName
End If

DE# = (degree) * 180# / Pi
   F.lfItalic = Abs(.italics)
F.lfWeight = Abs(.bold) * 800
  F.lfEscapement = CLng(10 * DE#)
  F.lfFaceName = Left$(Font, 30) + Chr$(0)
  F.lfCharSet = .charset
  F.lfQuality = 3 ' PROOF_QUALITY
  F.lfHeight = (Size * -20) / DYP

  hFont = CreateFontIndirect(F)
  hPrevFont = SelectObject(ddd.hDC, hFont)
nTextY = Int(TextWidth(ddd, what$) * Sin(degree) + TextHeight(ddd, what$) * Cos(degree))





  hFont = SelectObject(ddd.hDC, hPrevFont)
  DeleteObject hFont

End With
PlaceBasket ddd, players(prive)

End Function
Public Function nText(basestack As basetask, ByVal what As String, ByVal Font As String, ByVal Size As Single, Optional ByVal degree As Double = 0#)
Dim ddd As Object
Set ddd = basestack.Owner
Dim PX As Long, PY As Long, OLDFONT As String, OLDSIZE As String, DE#
Dim F As LOGFONT, hPrevFont As Long, hFont As Long
Dim BFONT As String
Dim prive As Long
prive = GetCode(ddd)
On Error Resume Next
With players(prive)
BFONT = ddd.Font.name
If Font <> "" Then
If Size = 0 Then Size = ddd.FontSize
StoreFont Font, Size, .charset
ddd.Font.charset = 0
ddd.FontSize = 9
ddd.FontName = .FontName
ddd.Font.charset = .charset
ddd.FontSize = Size
Else
Font = .FontName
End If

DE# = (degree) * 180# / Pi
   F.lfItalic = Abs(.italics)
F.lfWeight = Abs(.bold) * 800
  F.lfEscapement = CLng(10 * DE#)
  F.lfFaceName = Left$(Font, 30) + Chr$(0)
  F.lfCharSet = .charset
  F.lfQuality = 3 ' PROOF_QUALITY
  F.lfHeight = (Size * -20) / DYP

  hFont = CreateFontIndirect(F)
  hPrevFont = SelectObject(ddd.hDC, hFont)
nText = Int(TextWidth(ddd, what$) * Cos(degree) + TextHeight(ddd, what$) * Sin(degree))


  hFont = SelectObject(ddd.hDC, hPrevFont)
  DeleteObject hFont

End With
PlaceBasket ddd, players(prive)


End Function
Public Sub fullPlain(dd As Object, mb As basket, ByVal wh$, ByVal wi, Optional fake As Boolean = False, Optional nocr As Boolean = False)
Dim whNoSpace$, Displ As Long, DisplLeft As Long, i As Long, whSpace$, INTD As Long, MinDispl As Long, some As Long
Dim st As Long
st = DXP
MinDispl = (TextWidth(dd, "A") \ 2) \ st
If MinDispl <= 1 Then MinDispl = 3
MinDispl = st * MinDispl
INTD = TextWidth(dd, space$(MyTrimL3Len(wh$)))
dd.CurrentX = dd.CurrentX + INTD

wi = wi - INTD
wh$ = NLTrim2$(wh$)
INTD = wi + dd.CurrentX

whNoSpace$ = ReplaceStr(" ", "", wh$)
Dim magicratio As Double, whsp As Long, whl As Double


If whNoSpace$ = wh$ Then
MyPrintNew dd, mb.uMineLineSpace, wh$, Not nocr, fake

    'dd.Print wh$
Else
 If Len(whNoSpace$) > 0 Then
   whSpace$ = space$(Len(Trim$(wh$)) - Len(whNoSpace$))
   
        Displ = st * ((wi - TextWidth(dd, whNoSpace)) \ (Len(whSpace)) \ st)
        some = (wi - TextWidth(dd, whNoSpace) - Len(whSpace) * Displ) \ st  ' ((Displ - MinDispl) * Len(whSpace)) \ st
        magicratio = some / Len(whNoSpace)
        whsp = Len(whSpace)
                whNoSpace$ = vbNullString
                
        For i = 1 To Len(wh$)
            If Mid$(wh$, i, 1) = " " Then
            whsp = whsp - 1
            
               If whNoSpace$ <> "" Then
               whl = Len(whNoSpace$) * magicratio + whl
                    MyPrintNew dd, mb.uMineLineSpace, whNoSpace$, , fake
                whNoSpace$ = vbNullString
                End If
                If some > 0 Then
                '
                some = some - whl
                dd.CurrentX = ((dd.CurrentX + Displ) \ st) * st + CLng(whl) * st
                whl = whl - CLng(whl)
                Else
              dd.CurrentX = ((dd.CurrentX + Displ) \ st) * st
              End If
              
            Else
                whNoSpace$ = whNoSpace$ & Mid$(wh$, i, 1)
            End If
        Next i

          whl = Len(whNoSpace$) * magicratio + whl
      dd.CurrentX = dd.CurrentX + CLng(whl) * st
      
                   MyPrintNew dd, mb.uMineLineSpace, whNoSpace$, , fake
    Else

            MyPrintNew dd, mb.uMineLineSpace, wh$, Not nocr, fake
    End If
End If
End Sub
Public Sub fullPlainWhere(dd As Object, mb As basket, ByVal wh$, ByVal wi As Long, whr As Long, Optional fake As Boolean = False, Optional nocr As Boolean = False)
Dim whNoSpace$, Displ As Long, DisplLeft As Long, i As Long, whSpace$, INTD As Long, MinDispl As Long
MinDispl = (TextWidth(dd, "A") \ 2) \ DXP
If MinDispl <= 1 Then MinDispl = 3
MinDispl = DXP * MinDispl
If whr = 3 Or whr = 0 Then INTD = TextWidth(dd, space$(MyTrimL3Len(wh$)))
dd.CurrentX = dd.CurrentX + INTD
wi = wi - INTD
wh$ = NLTrim2$(wh$)
INTD = wi + dd.CurrentX
whNoSpace$ = ReplaceStr(" ", "", wh$)
If whr = 2 Then
wh$ = Trim$(wh$)
whNoSpace$ = ReplaceStr(" ", "", wh$)
dd.CurrentX = dd.CurrentX + ((wi - TextWidth(dd, whNoSpace) - (Len(wh$) - Len(whNoSpace)) * MinDispl)) / 2
ElseIf whr = 1 Then
dd.CurrentX = dd.CurrentX + (wi - TextWidth(dd, whNoSpace) - (Len(wh$) - Len(whNoSpace)) * MinDispl)
Else
INTD = (wi - TextWidth(dd, whNoSpace)) * 0.2 + dd.CurrentX

End If
If whNoSpace$ = wh$ Then
 MyPrintNew dd, mb.uMineLineSpace, wh$, Not nocr, fake
Else
 If Len(whNoSpace$) > 0 Then
   whSpace$ = space$(Len(Trim$(wh$)) - Len(whNoSpace$))
   INTD = TextWidth(dd, whSpace$) + dd.CurrentX
   
   wh$ = Trim$(wh$)
   Displ = MinDispl
   If Displ * Len(whSpace$) + TextWidth(dd, whNoSpace$) > wi Then
   Displ = (wi - TextWidth(dd, whNoSpace$)) / (Len(wh$))
   
   End If
     
    
                whNoSpace$ = vbNullString
        For i = 1 To Len(wh$)
            If Mid$(wh$, i, 1) = " " Then
            whSpace$ = Mid$(whSpace$, 2)
            
               If whNoSpace$ <> "" Then
                 MyPrintNew dd, mb.uMineLineSpace, whNoSpace$, , fake
                whNoSpace$ = vbNullString
                
                End If
              dd.CurrentX = dd.CurrentX + Displ
 
              
            Else
                whNoSpace$ = whNoSpace$ & Mid$(wh$, i, 1)
            End If
        Next i
        If whNoSpace$ <> "" Then

        End If
          MyPrintNew dd, mb.uMineLineSpace, whNoSpace$, Not nocr, fake
    Else
    MyPrintNew dd, mb.uMineLineSpace, wh$, Not nocr, fake
    
    End If
End If
End Sub

Public Sub wPlain(ddd As Object, mb As basket, ByVal what As String, ByVal wi&, ByVal Hi&, Optional nocr As Boolean = False)
Dim PX As Long, PY As Long, ttt As Long, ruller&
Dim buf$, b$, npy As Long ', npx As long

With mb
PlaceBasket ddd, mb
tParam.iTabLength = .ReportTab
If what = vbNullString Then Exit Sub
PX = .curpos
PY = .currow
If .mx - PX < wi& Then wi& = .mx - PX
If .My - PY < Hi& Then Hi& = .My - PY
If wi& = 0 Or Hi& < 0 Then Exit Sub
npy = PY
ruller& = wi&
For ttt = 1 To Len(what)
    b$ = Mid$(what, ttt, 1)
   ' If nounder32(b$) Then
   
   If Not (b$ = vbCr Or b$ = vbLf) Then
    If TextWidth(ddd, buf$ & b$) <= (wi& * .Xt) Then
    buf$ = buf$ & b$
    End If
    ElseIf b$ = vbCr Then
    
    If nocr Then Exit For
    MyPrintNew ddd, mb.uMineLineSpace, buf$, Not nocr
    
    
    buf$ = vbNullString
    Hi& = Hi& - 1
    npy = npy + 1
    LCTbasket ddd, mb, npy, PX
    End If
    If Hi& < 0 Then Exit For
Next ttt
If Hi& >= 0 And buf$ <> "" Then MyPrintNew ddd, mb.uMineLineSpace, buf$, Not nocr
If Not nocr Then LCTbasket ddd, mb, PY, PX
End With
End Sub
Public Sub wwPlain(bstack As basetask, mybasket As basket, ByVal what As String, ByVal wi As Long, ByVal Hi As Long, Optional scrollme As Boolean = False, Optional nosettext As Boolean = False, Optional frmt As Long = 0, Optional ByVal skip As Long = 0, Optional res As Long, Optional isAcolumn As Boolean = False, Optional collectit As Boolean = False, Optional nonewline As Boolean)
Dim ddd As Object, mDoc As Object
If collectit Then Set mDoc = New Document
Set ddd = bstack.Owner
Dim PX As Long, PY As Long, ttt As Long, ruller&, last As Boolean, INTD As Long, nowait As Boolean
Dim nopage As Boolean
Dim buf$, b$, npy As Long, kk&, lCount As Long, SCRnum2stop As Long, itnd As Long
Dim nopr As Boolean, nohi As Long, spcc As Long
Dim dv2x15 As Long
dv2x15 = dv15 * 2
If what = vbNullString Then Exit Sub
With mybasket
tParam.iTabLength = .ReportTab
If Not nosettext Then
PX = .curpos
PY = .currow
If PX >= .mx Then
nowait = True
PX = 0
End If
LCTbasket ddd, mybasket, PY, PX
Else
PX = .curpos
PY = .currow
End If
If PX > .mx Then nowait = True
If wi = 0 Then
If nowait Then wi = .Xt * (.mx - PX) Else wi = .mx * .Xt

Else
If wi <= .mx Then wi = wi * .Xt
End If

wi = wi - CLng(dv2x15)

ddd.CurrentX = ddd.CurrentX + dv2x15
If Not scrollme Then
If Hi >= 0 Then
If (.My - PY) * .Yt < Hi Then Hi = (.My - PY) * .Yt
End If
Else

If Hi > 1 Then
If .pageframe <> 0 Then
lCount = holdcontrol(ddd, mybasket)
.pageframe = 0
End If
SCRnum2stop = holdcontrol(ddd, mybasket)
End If
End If
If wi = 0 Then Exit Sub
npy = PY
Dim w2 As Long, kkl As Long, MinDispl As Long, OverDispl As Long
MinDispl = (TextWidth(ddd, "A") \ 2) \ DXP
If MinDispl <= 1 Then MinDispl = 3
MinDispl = DXP * MinDispl
 w2 = wi '- TextWidth( ddd, "i") +  dv2x15
 If w2 < 0 Then Exit Sub
 If Left$(what, 1) = " " Then INTD = 1
 Dim kku&
 OverDispl = MinDispl
If Hi < 0 Then
Hi = -Hi - 2
nohi = Hi
nopr = True
End If
Dim paragr As Boolean, help1 As Long, help2 As Long, hstr$
nopr = nopr Or collectit
paragr = True
If bstack.IamThread Then nopage = True
For ttt = 1 To Len(what)
If NOEXECUTION Then Exit For
b$ = Mid$(what, ttt, 1)
If paragr Then INTD = MyTrimL3Len(buf$ & b$)
If b$ = Chr$(0) Or b$ = vbLf Then
ElseIf Not b$ = vbCr Then
spcc = (Len(buf$ & b$) - Len(ReplaceStr(" ", "", Trim$(buf$ & b$))))

kkl = spcc * OverDispl
hstr$ = ReplaceStr(" ", "", buf$ & b$)
help1 = TextWidth(ddd, space(INTD) + hstr$)
kk& = (help1 + help2) < (w2 - kkl)
    If kk& Then '- 15 * Len(buf$) Then
        buf$ = buf$ & b$
    Else
         kk& = rinstr(Mid$(buf$, INTD + 1), " ") + INTD
         kku& = rinstr(Mid$(buf$, INTD + 1), "_") + INTD
         If kku& > kk& Then kk& = kku&
         If kk& = INTD Then kk& = Len(buf$) + 1
         If CDbl((Len(buf$) - INTD)) > 0 Then
         If (kk& - INTD) / CDbl((Len(buf$) - INTD)) > 0.5 And kkl / wi > 0.2 Then
         If InStr(Mid$(what, ttt), " ") < (Len(buf$) - kk&) Then
                                kk& = Len(buf$) + 1
                            If OverDispl > 5 * DXP Then
                                   OverDispl = MinDispl - 2 * DXP
                   
                              End If
                      buf$ = buf$ & b$
                      GoTo thmagic
                       ElseIf InStr(Mid$(what, ttt), "_") < (Len(buf$) - kk&) And InStr(Mid$(what, ttt), "_") <> 0 Then
      kk& = Len(buf$) + 1
                    If OverDispl > 5 * DXP Then
                         OverDispl = MinDispl - 2 * DXP
                   
                    End If
                      buf$ = buf$ & b$
                       GoTo thmagic
                       
               End If
         End If
         paragr = False: INTD = 0
         If b$ = "." Or b$ = "_" Or b$ = "," Then
         kk& = Len(buf$) + 1
       buf$ = buf$ & b$
       b$ = vbNullString
         End If
       End If
        If kk& > 0 And kk& < Len(buf$) Then
            b$ = Mid$(buf$, kk& + 1) + b$
                If last Then
                buf$ = Trim$(Left$(buf$, kk&))
                Else
            
                buf$ = Left$(buf$, kk&)
                
                End If
                End If
 
          skip = skip - 1
        If skip < 0 Then
        
            If last Then
             If frmt > 0 Then
                    If Not nopr Then fullPlainWhere ddd, mybasket, Trim$(buf$), w2, frmt, nowait, nonewline
               Else
                    If Not nopr Then fullPlain ddd, mybasket, Trim$(buf$), w2, nowait, nonewline   'DDD.Width ' w2
                 End If
                 If collectit Then
                 mDoc.AppendParagraphOneLine Trim$(buf$)
                 End If
            Else
                If frmt > 0 Then
                    If Not nopr Then fullPlainWhere ddd, mybasket, RTrim$(buf$), w2, frmt, nowait, nonewline ' rtrim
                Else
                    If Not nopr Then fullPlain ddd, mybasket, RTrim$(buf$), w2, nowait, nonewline
                    ' npy
                          End If
              If collectit Then
                 mDoc.AppendParagraphOneLine RTrim$(buf$)
                 End If
            End If
        End If
        If isAcolumn Then Exit Sub
        last = True
        buf$ = b$
        If skip < 0 Or scrollme Then
            Hi = Hi - 1
            lCount = lCount + 1
            npy = npy + 1
            
            If npy >= .My And scrollme Then
            If Not nopr Then
                If SCRnum2stop > 0 Then
                    If lCount >= SCRnum2stop Then
                      If Not bstack.toprinter Then
                       If Not nowait Then
                    
                    If Not nopage Then
                     ddd.Refresh
                        Do
   
                            mywait bstack, 10
                       
                        Loop Until INKEY$ <> "" Or mouse <> 0 Or NOEXECUTION
                        End If
                        End If
                        End If
                        SCRnum2stop = .pageframe
                        lCount = 1
                    
                    End If
                End If
                           If Not bstack.toprinter Then
                                ddd.Refresh
                                ScrollUpNew ddd, mybasket
                              ''If Not isAcolumn Then
                               ''    ddd.CurrentY = .My * .Yt - .Yt
                             '' End If
                            Else
                              getnextpage
                              npy = 1
                          End If
                End If
                npy = npy - 1
                      ''
         ElseIf npy >= .My Then
         
        If Not nopr Then crNew bstack, mybasket
               npy = npy - 1
              
          
      End If
If Not nopr Then LCTbasket ddd, mybasket, npy, PX: ddd.CurrentX = ddd.CurrentX + dv2x15
  End If
    End If
'ElseIf b$ = vbCr Then
Else
If nonewline Then Exit For
paragr = True
 skip = skip - 1
 
        If skip < 0 Or scrollme Then
        
If last Then
    If frmt > 0 Then
        If Not nopr Then fullPlainWhere ddd, mybasket, Trim$(buf$), w2, frmt, nowait, nonewline
    Else
    
        If Not nopr Then fullPlainWhere ddd, mybasket, Trim$(buf$), w2, 3, nowait, nonewline
    End If
        If collectit Then
                 mDoc.AppendParagraphOneLine Trim$(buf$)
                 End If
Else
If frmt > 0 Then
If Not nopr Then fullPlainWhere ddd, mybasket, RTrim$(buf$), w2, frmt, nowait, nonewline 'rtrim
Else

If Not nopr Then fullPlainWhere ddd, mybasket, RTrim$(buf$), w2, 3, nowait, nonewline ' rtrim
End If
    If collectit Then
                 mDoc.AppendParagraphOneLine RTrim$(buf$)
                 End If
End If
End If
last = False

buf$ = vbNullString
'''''''''''''''''''''''''
If isAcolumn Then Exit Sub
If skip < 0 Or scrollme Then
lCount = lCount + 1
    Hi = Hi - 1
    npy = npy + 1
    If npy >= .My And scrollme Then
    If Not nopr Then
            If SCRnum2stop > 0 Then
                If lCount >= SCRnum2stop Then
                     If Not bstack.toprinter Then
                     If Not nowait Then
                     If Not nopage Then
                     ddd.Refresh
                        Do
      
                            mywait bstack, 10
                        Loop Until INKEY$ <> "" Or mouse <> 0 Or NOEXECUTION
                         End If
                         End If
                         End If
                                    SCRnum2stop = .pageframe
                        lCount = 1
                End If
            End If
            
                  If Not bstack.toprinter Then
                            ddd.Refresh
                            ScrollUpNew ddd, mybasket
                ''     If Not isAcolumn Then
                        ddd.CurrentY = .My * .Yt - .Yt
                         ''  End If
                          Else
                          getnextpage
                          npy = 1
                          End If
            End If
            npy = npy - 1
    ElseIf npy >= .My Then
            
If Not nopr Then crNew bstack, mybasket
            ' 1ST
If Not nopr Then ddd.CurrentY = ddd.CurrentY - mybasket.Yt:   npy = npy - 1
    End If
' If Not nopr Then GetXYb2 ddd, mybasket, ruller&, npy
If Not nopr Then
    If nonewline Then npy = npy + 1
    
    ruller& = ddd.CurrentX \ mybasket.Xt
End If
conthere:
If Not nopr Then LCTbasket ddd, mybasket, npy, PX: ddd.CurrentX = ddd.CurrentX + dv2x15
End If
End If

If Hi < 0 Then
' Exit For
'
skip = 1000
scrollme = False
End If
 OverDispl = MinDispl
thmagic:
Next ttt
If Hi >= 0 And buf$ <> "" Then
 skip = skip - 1
        If skip < 0 Then
If frmt = 2 Then
If Not nopr Then fullPlainWhere ddd, mybasket, RTrim$(buf$), w2, frmt, nowait, nonewline
            If collectit Then
                 mDoc.AppendParagraphOneLine RTrim$(buf$)
                 End If
Else
If Hi = 0 And frmt = 0 And Not scrollme Then
If Not nopr Then

MyPrintNew ddd, mybasket.uMineLineSpace, buf$, , nowait     ';   '************************************************************************************

res = ddd.CurrentX
        If Trim$(buf$) = vbNullString Then
        ddd.CurrentX = ((ddd.CurrentX + .Xt \ 2) \ .Xt) * .Xt
        Else
        ddd.CurrentX = ((ddd.CurrentX + .Xt \ 1.2) \ .Xt) * .Xt
        End If
End If
            If collectit Then
                 mDoc.AppendParagraphOneLine buf$
                 End If


Exit Sub
Else
If Not nopr Then
fullPlainWhere ddd, mybasket, RTrim$(buf$), w2, frmt, nowait, nonewline
End If
    If collectit Then
                 mDoc.AppendParagraphOneLine buf$
                 End If
End If
End If
End If
If skip < 0 Or scrollme Then
    Hi = Hi - 1
    lCount = lCount + 1
   If Not isAcolumn Then npy = npy + 1
        If npy >= .My And scrollme Then

            If Not nopr Then  ' NOPT -> NOPR
                If SCRnum2stop > 0 Then
                    If lCount >= SCRnum2stop Then
                      If Not bstack.toprinter Then
                      If Not nowait Then
                      If Not nopage Then
                     ddd.Refresh
                    Do
            
                            mywait bstack, 10
                    Loop Until INKEY$ <> "" Or mouse <> 0 Or NOEXECUTION
                    End If
                    End If
                    End If
                                SCRnum2stop = .pageframe
                        lCount = 1
                    End If
                End If
                      If Not bstack.toprinter Then
                            ddd.Refresh
                            
                          ScrollUpNew ddd, mybasket
                             
                                   ddd.CurrentY = .My * .Yt - .Yt
                             
                          Else
                          getnextpage
                          npy = 1
                          End If
            End If
            npy = npy - 1
         ElseIf npy >= .My Then
          
            If npy >= .My Then
            
           If Not (nopr Or isAcolumn) Then crNew bstack, mybasket
            npy = npy - 1
            End If
        End If
    If Not nopr Then LCTbasket ddd, mybasket, npy, PX: ddd.CurrentX = ddd.CurrentX + dv2x15
    End If
End If
If scrollme Then

HoldReset lCount, mybasket
End If
res = nohi - Hi

wi = ddd.CurrentX
    If collectit Then
    Dim aa As Document
   bstack.soros.PushStr mDoc.textDoc
        Set mDoc = Nothing
                 End If
''GetXYb ddd, mybasket, .curpos, .currow
End With
End Sub

Public Sub FeedFont2Stack(basestack As basetask, ok As Boolean)
Dim mS As New mStiva
If ok Then
mS.PushVal CDbl(ReturnBold)
mS.PushVal CDbl(ReturnItalic)
mS.PushVal CDbl(ReturnCharset)
mS.PushVal CDbl(ReturnSize)
mS.PushStr ReturnFontName
mS.PushVal CDbl(1)
Else
mS.PushVal CDbl(0)
End If
basestack.soros.MergeTop mS
End Sub
Public Sub nPlain(basestack As basetask, ByVal what As String, ByVal Font As String, ByVal Size As Single, Optional ByVal degree As Double = 0#, Optional ByVal JUSTIFY As Long = 0, Optional ByVal qual As Boolean = True, Optional ByVal ExtraWidth As Long = 0)
Dim ddd As Object
Set ddd = basestack.Owner
Dim PX As Long, PY As Long, OLDFONT As String, OLDSIZE As Long, DEGR As Double
Dim F As LOGFONT, hPrevFont As Long, hFont As Long, fline$, ruler As Long
Dim BFONT As String
On Error Resume Next
BFONT = ddd.Font.name
If ExtraWidth <> 0 Then
SetTextCharacterExtra ddd.hDC, ExtraWidth
End If
Dim icx As Long, icy As Long, X As Long, Y As Long, icH As Long

If JUSTIFY < 0 Then degree = 0
DEGR = (degree) * 180# / Pi
  F.lfItalic = Abs(basestack.myitalic)
  F.lfWeight = Abs(basestack.myBold) * 800
  F.lfEscapement = 0
  F.lfFaceName = Left$(Font, 30) + Chr$(0)
  F.lfCharSet = basestack.myCharSet
  If qual Then
  F.lfQuality = PROOF_QUALITY 'NONANTIALIASED_QUALITY '
  Else
  F.lfQuality = NONANTIALIASED_QUALITY
  End If
  F.lfHeight = (Size * -20) / DYP
  hFont = CreateFontIndirect(F)
  hPrevFont = SelectObject(ddd.hDC, hFont)
    icH = TextHeight(ddd, "fq")
  hFont = SelectObject(ddd.hDC, hPrevFont)
  DeleteObject hFont
 F.lfItalic = Abs(basestack.myitalic)
  F.lfWeight = Abs(basestack.myBold) * 800
F.lfEscapement = CLng(10 * DEGR)
  F.lfFaceName = Left$(Font, 30) + Chr$(0)
  F.lfCharSet = basestack.myCharSet
  If qual Then
  F.lfQuality = PROOF_QUALITY 'NONANTIALIASED_QUALITY '
  Else
  F.lfQuality = NONANTIALIASED_QUALITY
  End If
  F.lfHeight = (Size * -20) / DYP
  

  
    hFont = CreateFontIndirect(F)
  hPrevFont = SelectObject(ddd.hDC, hFont)



icy = CLng(Cos(degree) * icH)
icx = CLng(Sin(degree) * icH)

With players(GetCode(ddd))
If JUSTIFY < 0 Then
JUSTIFY = Abs(JUSTIFY) - 1
If JUSTIFY = 0 Then
Y = .YGRAPH - icy
X = .XGRAPH - icx * 2
ElseIf JUSTIFY = 1 Then
Y = .YGRAPH
X = .XGRAPH
Else
Y = .YGRAPH - icy / 2
X = .XGRAPH - icx
End If
Else
Y = .YGRAPH - icy
X = .XGRAPH - icx

End If
End With
what$ = ReplaceStr(vbCrLf, vbCr, what) + vbCr
Dim textmetrics As POINTAPI
Do While what$ <> ""
If Left$(what$, 1) = vbCr Then
fline$ = vbNullString
what$ = Mid$(what$, 2)
Else
fline$ = GetStrUntil(vbCr, what$)
End If
textmetrics.X = 0
textmetrics.Y = 0
GetTextExtentPoint32 ddd.hDC, StrPtr(fline$), Len(fline$), textmetrics

X = X + icx
Y = Y + icy
If JUSTIFY = 1 Then
    ddd.CurrentX = X - Int((textmetrics.X * Cos(degree) + textmetrics.Y * Sin(degree)) * dv15)
    ddd.CurrentY = Y + Int((textmetrics.X * Sin(degree) - textmetrics.Y * Cos(degree)) * dv15)
ElseIf JUSTIFY = 2 Then
     ddd.CurrentX = X - Int((textmetrics.X * Cos(degree) + textmetrics.Y * Sin(degree)) * dv15) \ 2
    ddd.CurrentY = Y + Int((textmetrics.X * Sin(degree) - textmetrics.Y * Cos(degree)) * dv15) \ 2
Else
    ddd.CurrentX = X
    ddd.CurrentY = Y
End If
MyPrint ddd, fline$
Loop
  hFont = SelectObject(ddd.hDC, hPrevFont)
  DeleteObject hFont
If ExtraWidth <> 0 Then SetTextCharacterExtra ddd.hDC, 0
End Sub

Public Sub nForm(bstack As basetask, TheSize As Single, nW As Long, nH As Long, myLineSpace As Long)
    On Error Resume Next
    StoreFont bstack.Owner.Font.name, TheSize, bstack.myCharSet
    nH = fonttest.TextHeight("Wq") + myLineSpace * 2
    nW = fonttest.TextWidth("W") + dv15
End Sub

Sub crNew(bstack As basetask, mb As basket)
Dim d As Object
Set d = bstack.Owner
With mb
Dim PX As Long, PY As Long, r As Long
PX = .curpos
PY = .currow
PX = 0
PY = PY + 1
If PY >= .My Then

If Not bstack.toprinter Then
ScrollUpNew d, mb
PY = .My - 1
Else
PY = 0
PX = 0
getnextpage
End If
End If
.curpos = PX
.currow = PY

End With
End Sub

Public Sub CdESK()
Dim X, Y, ff As Form, useform1 As Boolean
If Form1.Visible Then
    If Form5.Visible Then
    Set ff = Form5
    Form5.RestoreSizePos
    Form5.backcolor = 0
    useform1 = True
    Else
    Set ff = Form1
    End If
    X = ff.Left / DXP
    Y = ff.top / DYP
    If useform1 Then Form1.Visible = False
    ff.Hide
    Sleep 50
    MyDoEvents1 ff, True
    
    Dim aa As New cDIBSection
    aa.CreateFromPicture hDCToPicture(GetDC(0), X, Y, ff.Width / DXP, ff.Height / DYP)
    aa.ThumbnailPaint ff
    GdiFlush
      ff.Visible = True
      
    If useform1 Then Form1.Visible = True
    
End If
Set ff = Nothing
End Sub
Private Sub FillBack(thathDC As Long, there As RECT, bgcolor As Long)
' create brush
Dim my_brush As Long
my_brush = CreateSolidBrush(bgcolor)
FillRect thathDC, there, my_brush
DeleteObject my_brush
End Sub

Public Sub ScrollUpNew(d As Object, mb As basket)
Dim ar As RECT, r As Long
Dim p As Long
With mb
ar.Left = 0
ar.Bottom = d.Height / dv15
ar.Right = d.Width / dv15
ar.top = .mysplit * .Yt / dv15
p = .Yt / dv15
r = BitBlt(d.hDC, CLng(ar.Left), CLng(ar.top), CLng(ar.Right), CLng(ar.Bottom - p), d.hDC, CLng(ar.Left), CLng(ar.top + p), SRCCOPY)

 ar.top = ar.Bottom - p
FillBack d.hDC, ar, .Paper
.curpos = 0
.currow = .My - 1
End With
GdiFlush
End Sub
Public Sub ScrollDownNew(d As Object, mb As basket)
Dim ar As RECT, r As Long
Dim p As Long
With mb
ar.Left = 0
ar.Bottom = d.ScaleY(d.Height, 1, 3)
ar.Right = d.ScaleX(d.Width, 1, 3)
ar.top = d.ScaleY(.mysplit * .Yt, 1, 3)
p = d.ScaleY(.Yt, 1, 3)
r = BitBlt(d.hDC, CLng(ar.Left), CLng(ar.top + p), CLng(ar.Right), CLng(ar.Bottom - p), d.hDC, CLng(ar.Left), CLng(ar.top), SRCCOPY)
d.Line (0, .mysplit * .Yt)-(d.ScaleWidth, .mysplit * .Yt + .Yt), .Paper, BF
.currow = .mysplit
.curpos = 0
End With
End Sub





Public Sub SetText(dq As Object, Optional alinespace As Long = -1, Optional ResetColumns As Boolean = False)
' can be used for first time also
Dim mymul As Long
On Error Resume Next
With players(GetCode(dq))
If .FontName = vbNullString Or alinespace = -2 Then
' we have to make it
If alinespace = -2 Then alinespace = 0
ResetColumns = True
.FontName = dq.FontName
.charset = dq.Font.charset
.SZ = dq.FontSize
Else
If Not (fonttest.FontName = .FontName And fonttest.Font.charset = dq.Font.charset And fonttest.Font.Size = .SZ) Then
fonttest.Font.charset = .charset
If fonttest.Font.charset = .charset Then
StoreFont .FontName, .SZ, .charset
dq.Font.charset = 0
dq.FontSize = 9
dq.FontName = .FontName
dq.Font.charset = .charset
dq.FontSize = .SZ
End If
End If
End If
If alinespace <> -1 Then
If .uMineLineSpace = .MineLineSpace * 2 And .MineLineSpace <> 0 Then
.MineLineSpace = alinespace
.uMineLineSpace = alinespace * 2
Else
.MineLineSpace = alinespace
.uMineLineSpace = alinespace ' so now we have normal
End If
End If
.SZ = dq.FontSize
.Xt = fonttest.TextWidth("W") + dv15
.Yt = fonttest.TextHeight("fj")
.mx = Int(dq.Width / .Xt)
.My = Int(dq.Height / (.Yt + .uMineLineSpace * 2))
''.Paper = dq.BackColor
If .My <= 0 Then .My = 1
If .mx <= 0 Then .mx = 1
.Yt = .Yt + .uMineLineSpace * 2
If ResetColumns Then
mymul = Int(.mx / 8)
If mymul = 1 Then mymul = 2
If mymul = 0 Then
.Column = .mx \ 2 - 1
Else
.Column = Int(.mx / mymul)
While (.mx Mod .Column) > 0 And (.mx / .Column >= 3)
.Column = .Column + 1
Wend
End If
If .Column = 0 Then .Column = .mx
.Column = .Column - 1
If .Column < 4 Then .Column = 4
End If
.MAXXGRAPH = dq.Width
.MAXYGRAPH = dq.Height
End With

End Sub

Public Sub SetTextSZ(dq As Object, mSz As Single, Optional factor As Single = 1, Optional AddTwipsTop As Long = -1)
' Used for making specific basket
On Error Resume Next
With players(GetCode(dq))
If AddTwipsTop < 0 Then
    If .double And factor = 1 Then
    .mysplit = .osplit
    .Column = .OCOLUMN
    .currow = (.currow + 1) * 2 - 2
    .curpos = .curpos * 2
    mSz = .SZ / 2
    .uMineLineSpace = .MineLineSpace
    .double = False
    ElseIf factor = 2 And Not .double Then
     .osplit = .mysplit
     .OCOLUMN = .Column
     .Column = .Column / 2
     .mysplit = .mysplit / 2
     .currow = (.currow + 1) / 2
     .curpos = .curpos / 2
     mSz = .SZ * 2
    .uMineLineSpace = .MineLineSpace * 2
    .double = True
    End If
Else

mSz = mSz * factor
.MineLineSpace = AddTwipsTop
.uMineLineSpace = AddTwipsTop * factor
.double = factor <> 1
End If
dq.FontSize = mSz

StoreFont dq.Font.name, mSz, dq.Font.charset
If .double Then
    Dim nowtextheight As Long
    nowtextheight = fonttest.TextHeight("fj")
    If .MineLineSpace = 0 Then
    Else
    If (.Yt - .MineLineSpace * 2) * 2 <> nowtextheight Then
    .uMineLineSpace = Int((.MAXYGRAPH - nowtextheight * .My / 2) / .My)
    End If
    
    End If
End If
SetText dq



If .My <= 0 Then .My = 1
If .mx <= 0 Then .mx = 1
.SZ = dq.FontSize
.MAXXGRAPH = dq.Width
.MAXYGRAPH = dq.Height
End With

End Sub

Public Sub SetTextBasketBack(dq As Object, mb As basket)
' set minimum display parameters for current object
' need an already filled basket
On Error Resume Next
With mb

If Not (dq.FontName = .FontName And dq.Font.charset = .charset And dq.Font.Size = .SZ) Then

StoreFont .FontName, .SZ, .charset
dq.Font.charset = 0
dq.FontSize = 9
dq.FontName = .FontName
dq.Font.charset = .charset
dq.FontSize = .SZ
End If
dq.ForeColor = .mypen

If Not dq.backcolor = .Paper Then
   dq.backcolor = .Paper
End If
End With
End Sub

Function gf$(bstack As basetask, ByVal Y&, ByVal X&, ByVal a$, c&, F&, Optional STAR As Boolean = False)
On Error Resume Next
Dim cLast&, b$, cc$, dq As Object, ownLinespace, oldrefresh As Double, noinp As Double
oldrefresh = REFRESHRATE
Dim mybasket As basket, addpixels As Long
GFQRY = True
Set dq = bstack.Owner
SetText dq
mybasket = players(GetCode(dq))

With mybasket
If InternalLeadingSpace() = 0 And .MineLineSpace = 0 Then
addpixels = 0
Else
addpixels = 2
End If
If dq.Visible = False Then dq.Visible = True
If exWnd = 0 Then dq.SetFocus
dq.FontTransparent = False
LCTbasket dq, mybasket, Y&, X&
Dim o$
o$ = a$
If a$ = vbNullString Then a$ = " "
INK$ = vbNullString

Dim XX&
XX& = X&

X& = X& - 1

cLast& = Len(a$)
'*****************
If cLast& + X& >= .mx Then
MyDoEvents
If dq.Font.charset = 161 Then
b$ = InputBoxN("Εισαγωγή Τιμής Πεδίου", MesTitle$, a$, noinp)
Else
b$ = InputBoxN("Input Field Value", MesTitle$, a$, noinp)
End If
If noinp <> 1 Then b$ = a$
If MyTrim(b$) < "A" Then b$ = Right$(String$(cLast&, " ") + b$, cLast&) Else b$ = Left$(b$ + String$(cLast&, " "), cLast&)
gf$ = b$
If XX& < .mx Then
dq.FontTransparent = False
If STAR Then
PlainBaSket dq, mybasket, StarSTR(Left$(b$, .mx - X&)), True, , addpixels
Else
PlainBaSket dq, mybasket, Left$(b$, .mx - X&), True, , addpixels
End If
End If
GoTo GFEND
Else
dq.FontTransparent = False
If STAR Then
PlainBaSket dq, mybasket, StarSTR(a$), True, , addpixels
Else
PlainBaSket dq, mybasket, a$, True, , addpixels
End If
End If

'************
b$ = a$
.currow = Y&
.curpos = c& + X&
LCTCB dq, mybasket, ins&

Do
MyDoEvents1 Form1, , True
If bstack.IamThread Then If myexit(bstack) Then GoTo contgfhere
If Not TaskMaster Is Nothing Then
If TaskMaster.QueueCount > 0 Then
dq.FontTransparent = True
TaskMaster.RestEnd1
TaskMasterTick
End If
End If
 cc$ = INKEY$
 If cc$ <> "" Then
If Not TaskMaster Is Nothing Then TaskMaster.rest
SetTextBasketBack dq, mybasket
 Else
If Not TaskMaster Is Nothing Then TaskMaster.RestEnd
SetTextBasketBack dq, mybasket
        If iamactive Then
           If Screen.ActiveForm Is Nothing Then
                            DestroyCaret
                      nomoveLCTC dq, mybasket, Y&, c& + X&, ins&
                      iamactive = False
           Else
                If Not (GetForegroundWindow = Screen.ActiveForm.hWnd And Screen.ActiveForm.name = "Form1") Then
                 
                      DestroyCaret
                      nomoveLCTC dq, mybasket, Y&, c& + X&, ins&
                      iamactive = False
             Else
                         If ShowCaret(dq.hWnd) = 0 Then
                                   HideCaret dq.hWnd
                                   .currow = Y&
                                   .curpos = c& + X&
                                   LCTCB dq, mybasket, ins&
                                   ShowCaret dq.hWnd
                         End If
                End If
                End If
     Else
  If Not Screen.ActiveForm Is Nothing Then
            If GetForegroundWindow = Screen.ActiveForm.hWnd And Screen.ActiveForm.name = "Form1" Then
           
                          nomoveLCTC dq, mybasket, Y&, c& + X&, ins&
                             iamactive = True
                              If ShowCaret(dq.hWnd) = 0 And Screen.ActiveForm.name = "Form1" Then
                                   HideCaret dq.hWnd
                                   .currow = Y&
                                   .curpos = c& + X&
                                   LCTCB dq, mybasket, ins&
                                   ShowCaret dq.hWnd
                         End If
                         End If
            End If
     End If

 End If

 
        If NOEXECUTION Then
        If KeyPressed(&H1B) Then
                       F& = 99 'ESC  ****************
                        c& = 1
                        gf$ = o$
                        b$ = o$
                                          NOEXECUTION = False
                                         BLOCKkey = True
                                    While KeyPressed(&H1B)
                                    If Not TaskMaster Is Nothing Then
                             If TaskMaster.Processing Then
                                                TaskMaster.RestEnd1
                                                TaskMaster.TimerTick
                                                TaskMaster.rest
                                                MyDoEvents1 dq
                                                Else
                                                MyDoEvents
                                                
                                                End If
                                                Else
                                                DoEvents
                                                End If
'''sleepwait 1
                                    Wend
                                                                        BLOCKkey = False
                                                                        End If
                 Exit Do
        End If
        Select Case Len(cc$)
        Case 0
        If FKey > 0 Then
        If FK$(FKey) <> "" And FKey <> 13 Then
            cc$ = FK$(FKey)
            interpret basestack1, cc$
        
        End If
        FKey = 0
        Else
        
        End If
        
        Case 1
        If STAR And cc$ = " " Then cc$ = Chr$(127)
                Select Case AscW(cc$)
                Case 8
                        If c& > 1 Then
                        Mid$(b$, c& - 1) = Mid$(b$, c&) & " "
                         c& = c& - 1
                         dq.FontTransparent = False
                                   .currow = Y&
                                   .curpos = c& + X&
                                   LCTCB dq, mybasket, ins&
                        If STAR Then
                        PlainBaSket dq, mybasket, StarSTR(Mid$(b$, c&)), True, , addpixels
                        Else
                        PlainBaSket dq, mybasket, Mid$(b$, c&), True, , addpixels
                        End If
                         dq.Refresh
                                   .currow = Y&
                                   .curpos = c& + X&
                                   LCTCB dq, mybasket, ins&
                        End If
                Case 6
                F& = -1
                 gf$ = b$
                Exit Do
                Case 13, 9
                F& = 1 'NEXT  *************
                gf$ = b$
                Exit Do

                Case 27
                        F& = 99 'ESC  ****************
                        c& = 1
                        gf$ = o$
                        b$ = o$
                                    NOEXECUTION = False
                                    BLOCKkey = True
                                    While KeyPressed(&H1B)
                                    If Not TaskMaster Is Nothing Then
                                    If TaskMaster.Processing Then
                                            TaskMaster.RestEnd1
                                            TaskMaster.TimerTick
                                            TaskMaster.rest
                                            MyDoEvents1 dq
                                            Else
                                            MyDoEvents
                                            
                                            End If
                                            Else
                                            DoEvents
                                            End If
                                    ''''MyDoEvents
                                    Wend
                                                                        BLOCKkey = False
                        NOEXECUTION = False
                        Exit Do
                       Case 32 To 126, Is > 128
           
                        .currow = Y&
                        .curpos = c& + X&
                        LCTCB dq, mybasket, ins&
                        If ins& = 1 Then
                          If AscW(cc$) = 32 And STAR Then
                If AscW(Mid$(b$, c& + 1)) > 32 Then
                 Mid$(b$, c&) = Mid$(b$, c& + 1) & " "
                End If
                
                
                Else
                        
                                                
                        Mid$(b$, c&, 1) = cc$
                        dq.FontTransparent = False
                        If STAR Then
                        PlainBaSket dq, mybasket, StarSTR(Mid$(b$, c&)), True, , addpixels
                        Else
                        PlainBaSket dq, mybasket, Mid$(b$, c&), True, , addpixels
                        End If
                         dq.Refresh
                        End If
                        If c& < Len(b$) Then c& = c& + 1
                                   .currow = Y&
                                   .curpos = c& + X&
                                   LCTCB dq, mybasket, ins&
                        Else
                                 If AscW(cc$) = 32 And STAR Then
            
                
                
                Else
                     
                        LSet b$ = Left$(b$, c& - 1) + cc$ & Mid$(b$, c&)
                        dq.FontTransparent = False
                        If STAR Then
                        PlainBaSket dq, mybasket, StarSTR(Mid$(b$, c&)), True, , addpixels
                        Else
                        PlainBaSket dq, mybasket, Mid$(b$, c&), True, , addpixels
                        End If
                         dq.Refresh
                        'LCTC Dq, Y&, X& + C& + 1, INS&
                        End If
                        If c& < cLast& Then c& = c& + 1
                                .currow = Y&
                                .curpos = c& + X&
                                LCTCB dq, mybasket, ins&
                        End If
                End Select
        Case 2
                Select Case AscW(Right$(cc$, 1))
                Case 81
                F& = 10 ' exit - pagedown ***************
                gf$ = b$
                Exit Do
                Case 73
                F& = -10 ' exit - pageup
                gf$ = b$
                Exit Do
                Case 79
                F& = 20 ' End
                gf$ = b$
                Exit Do
                Case 71
                F& = -20 ' exit - home
                gf$ = b$
                Exit Do
                Case 75 'LEFT
                        If c& > 1 Then
                                   .currow = Y&
                                .curpos = c& + X&
                                LCTCB dq, mybasket, ins&
                        c& = c& - 1:
                        .currow = Y&
                        .curpos = c& + X&
                        LCTCB dq, mybasket, ins&
                        End If
                Case 77 'RIGHT
                        If c& < cLast& Then
                      
                If Not (AscW(Mid$(b$, c&)) = 32 And STAR) Then
                
             
                                    .currow = Y&
                                .curpos = c& + X&
                                LCTCB dq, mybasket, ins&
                        c& = c& + 1:
                        .currow = Y&
                                .curpos = c& + X&
                                LCTCB dq, mybasket, ins&
                        End If
                        End If
                Case 72 ' EXIT UP
                F& = -1 ' PREVIUS ***************
                gf$ = b$
                Exit Do
                Case 80 'EXIT DOWN OR ENTER OR TAB
                F& = 1 'NEXT  *************
                gf$ = b$
                Exit Do
                Case 82
                            .currow = Y&
                                .curpos = c& + X&
                                LCTCB dq, mybasket, ins&
                ins& = 1 - ins&
                           .currow = Y&
                                .curpos = c& + X&
                                LCTCB dq, mybasket, ins&
                Case 83
                        Mid$(b$, c&) = Mid$(b$, c& + 1) & " "
                        dq.FontTransparent = False
                        LCTbasket dq, mybasket, Y&, c& + X&
                        If STAR Then
                        PlainBaSket dq, mybasket, StarSTR(Mid$(b$, c&)), True, , addpixels
                        Else
                        PlainBaSket dq, mybasket, Mid$(b$, c&), True, , addpixels
                        End If
                               .currow = Y&
                                .curpos = c& + X&
                                LCTCB dq, mybasket, ins&
                     dq.Refresh
                End Select
        End Select
      
Loop

GFEND:
REFRESHRATE = oldrefresh
LCTbasket dq, mybasket, Y&, X& + 1
If X& < .mx And Not XX& > .mx Then
If STAR Then
 PlainBaSket dq, mybasket, StarSTR(b$), True, , addpixels
Else
PlainBaSket dq, mybasket, b$, True, , addpixels
End If
contgfhere:
 dq.Refresh
If Not TaskMaster Is Nothing Then If TaskMaster.QueueCount > 0 Then TaskMaster.RestEnd
End If
dq.FontTransparent = True
 DestroyCaret
Set dq = Nothing
TaskMaster.RestEnd1
GFQRY = False
End With
End Function

Public Sub ResetPrefresh()
Dim i As Long
For i = -2 To 131
    Prefresh(i).k1 = 0
    Prefresh(i).RRCOUNTER = 0
Next i

End Sub

Sub original(bstack As basetask, COM$)
Dim d As Object, b$

If COM$ <> "" Then QUERYLIST = vbNullString
If Form1.Visible Then REFRESHRATE = 25: ResetPrefresh
If bstack.toprinter Then
bstack.toprinter = False
Form1.PrinterDocument1.Cls
Set d = bstack.Owner
Else
Set d = bstack.Owner
End If
On Error Resume Next
Dim basketcode As Long
basketcode = GetCode(d)


Form1.IEUP ""
Form1.KeyPreview = True
Dim dummy As Boolean, rs As String, mPen As Long, ICO As Long, BAR As Long, bar2 As Long
BAR = 1
Form1.DIS.Visible = True
GDILines = False  ' reset to normal ' use Smooth on to change this to true
If COM$ <> "" Then d.Visible = False
ClrSprites
mPen = PenOne
d.Font.bold = bstack.myBold
d.Font.Italic = bstack.myitalic
GetMonitorsNow
Console = FindPrimary
With ScrInfo(Console)
If SzOne < 4 Then SzOne = 4
    'Form1.Visible = False
   ' If IsWine Then
    Sleep 30
    .Width = .Width - dv15 - 1
    .Height = .Height - dv15 - 1
   ' End If
    If Not Form1.WindowState = 0 Then Form1.WindowState = 0
    Sleep 10
    If Form1.WindowState = 0 Then
        Form1.Move .Left, .top, .Width - 1, .Height - 1
        If Form1.top <> .Left Or Form1.Left <> .top Then
            Form1.Cls
            Form1.Move .Left, .top, .Width - 1, .Height - 1
        End If
    Else
        Sleep 100
        On Error Resume Next
        Form1.WindowState = 0
        Form1.Move .Left, .top, .Width - 1, .Height - 1
        If Form1.top <> .top Or Form1.Left <> .Left Then
        Form1.Cls
        Form1.Move .Left, .top, .Width - 1, .Height - 1
        End If
    End If
NoBackFormFirstUse = False
If players(-1).MAXXGRAPH <> 0 Then ClearScrNew Form1, players(-1), 0&
Form1.DIS.Visible = True
FrameText d, SzOne, (.Width + .Left - 1 - Form1.Left), (.Height + .top - 1 - Form1.top), PaperOne
End With
Form1.DIS.backcolor = mycolor(PaperOne)
If lckfrm = 0 Then
SetText d
bstack.Owner.Font.charset = bstack.myCharSet
StoreFont bstack.Owner.Font.name, SzOne, bstack.myCharSet
 
 With players(basketcode)
.mypen = PenOne
.XGRAPH = 0
.YGRAPH = 0
.bold = bstack.myBold '' I have to change that
.italics = bstack.myitalic
.FontName = bstack.Owner.FontName
.SZ = SzOne
.charset = bstack.myCharSet
.MAXXGRAPH = Form1.Width
.MAXYGRAPH = Form1.Height
.Paper = bstack.Owner.backcolor
.mypen = mycolor(PenOne)
End With


 
' check to see if
Dim ss$, skipthat As Boolean
If Not IsSupervisor Then
    ss$ = ReadUnicodeOrANSI(userfiles & "desktop.inf")
    LastErNum = 0
    If ss$ <> "" Then
     skipthat = interpret(bstack, ss$)
     If mycolor(PenOne) <> d.ForeColor Then
     PenOne = -d.ForeColor
     End If
    End If
End If
If SzOne < 36 And d.Height / SzOne > 250 Then SetDouble d: BAR = BAR + 1
If SzOne < 83 Then

If bstack.myCharSet = 161 Then
b$ = "ΠΕΡΙΒΑΛΛΟΝ "
Else
b$ = "ENVIRONMENT "
End If
d.ForeColor = mycolor(PenOne)
LCTbasket d, players(DisForm), 0, 0
wwPlain bstack, players(DisForm), b$ & "M2000", d.Width, 0, 0 '',True
ICO = TextWidth(d, b$ & "M2000") + 100
' draw graphic'
Dim IX As Long, IY As Long
With players(DisForm)
IX = (.Xt \ 25) * 25
IY = Form1.icon.Height * IX / Form1.icon.Width
If IsWine Then
Form1.DIS.PaintPicture Form1.icon, ICO, (.Yt - IY) / 2, IX, IY
Form1.DIS.PaintPicture Form1.icon, ICO, (.Yt - IY) / 2, IX, IY
Else
Dim myico As New cDIBSection
myico.backcolor = Form1.DIS.backcolor
myico.CreateFromPicture Form1.icon
Form1.DIS.PaintPicture myico.Picture(1), ICO, (.Yt - IY) / 2, IX, IY
End If
End With

' ********
SetNormal d
   Dim osbit As String
   If Is64bit Then osbit = " (64-bit)" Else osbit = " (32-bit)"
        LCTbasket d, players(basketcode), BAR, 0
        rs = vbNullString
            If bstack.myCharSet = 161 Then
            If Revision = 0 Then
            wwPlain bstack, players(DisForm), "Έκδοση Διερμηνευτή: " & CStr(VerMajor) & "." & CStr(VerMinor), d.Width, 0, True
            Else
                    wwPlain bstack, players(DisForm), "Έκδοση Διερμηνευτή: " & CStr(VerMajor) & "." & Left$(CStr(VerMinor), 1) & " (" & CStr(Revision) & ")", d.Width, 0, True
                End If
                   wwPlain bstack, players(DisForm), "Λειτουργικό Σύστημα: " & os & osbit, d.Width, 0, True
            
                      wwPlain bstack, players(DisForm), "Όνομα Χρήστη: " & Tcase(Originalusername), d.Width, 0, True
                
            Else
             If Revision = 0 Then
              wwPlain bstack, players(DisForm), "Interpreter Version: " & CStr(VerMajor) & "." & CStr(VerMinor), d.Width, 0, True
             Else
                    wwPlain bstack, players(DisForm), "Interpreter Version: " & CStr(VerMajor) & "." & Left$(CStr(VerMinor), 1) & " rev. (" & CStr(Revision) & ")", d.Width, 0, True
                 End If
              
                      wwPlain bstack, players(DisForm), "Operating System: " & os & osbit, d.Width, 0, True
                
                   wwPlain bstack, players(DisForm), "User Name: " & Tcase(Originalusername), d.Width, 0, True
        
                 End If
                        '    cr bstack
            GetXYb d, players(basketcode), bar2, BAR
             players(basketcode).curpos = bar2
            players(basketcode).currow = BAR
           BAR = BAR + 1
            If BAR >= players(basketcode).My Then ScrollUpNew d, players(basketcode)
                    LCTbasket d, players(basketcode), BAR, 0
                    players(basketcode).curpos = 0
            players(basketcode).currow = BAR
    End If
If Not skipthat Then
 dummy = interpret(bstack, "PEN " & CStr(mPen) & ":CLS ," & CStr(BAR))
End If
End If
If Not skipthat Then
dummy = interpret(bstack, COM$)
End If
'cr bstack
End Sub
Sub ClearScr(d As Object, c1 As Long)
Dim aa As Long
With players(GetCode(d))
.Paper = c1
.curpos = 0
.currow = 0
.lastprint = False
End With
d.Line (0, 0)-(d.ScaleWidth - dv15, d.ScaleHeight - dv15), c1, BF
d.CurrentX = 0
d.CurrentY = 0

End Sub
Sub ClearScrNew(d As Object, mb As basket, c1 As Long)
Dim im As New StdPicture, spl As Long
With mb
spl = .mysplit * .Yt
Set im = d.Image
.Paper = c1

If TypeOf d Is GuiM2000 Then
If .mysplit = 0 Then
    If Not d.backcolor = c1 Then d.backcolor = c1
    d.Cls
Else
    d.Line (0, spl)-(d.ScaleWidth - dv15, d.ScaleHeight - dv15), .Paper, BF
    End If
    .currow = .mysplit
ElseIf d.name = "Form1" Or mb.used Then
d.Line (0, spl)-(d.ScaleWidth - dv15, d.ScaleHeight - dv15), .Paper, BF
.curpos = 0
.currow = .mysplit
Else
d.backcolor = c1
If spl > 0 Then d.PaintPicture im, 0, 0, d.Width, spl, 0, 0, d.Width, spl, vbSrcCopy
.curpos = 0
.currow = .mysplit

End If
.lastprint = False
d.CurrentX = 0
d.CurrentY = 0
End With
End Sub
Function iText(bb As basetask, ByVal v$, wi&, Hi&, aTitle$, n As Long, Optional NumberOnly As Boolean = False, Optional UseIntOnly As Boolean = False) As String
Dim X&, Y&, dd As Object, wh&, shiftlittle As Long, OLDV$
Set dd = bb.Owner
With players(GetCode(dd))
If .lastprint Then
X& = (dd.CurrentX + .Xt - dv15) \ .Xt
Y& = dd.CurrentY \ .Yt
shiftlittle = X& * .Xt - dd.CurrentX
If Y& > .mx Then
Y& = .mx - 1
crNew bb, players(GetCode(dd))

End If
Else
X& = .curpos
Y& = .currow
End If
If .mx - X& - 1 < wi& Then wi& = .mx - X&
If .My - Y& - 1 < Hi& Then Hi& = .My - Y& - 1
If wi& = 0 Or Hi& < 0 Then
iText = v$
Exit Function
End If
wi& = wi& + X&
Hi& = Hi& + Y&
Form1.EditTextWord = True
wh& = -1
Dim oldshow As Boolean
With Form1.TEXT1
     oldshow = .showparagraph
    .showparagraph = False
    
    If n <= 0 Then .Title = aTitle$ + " ": wh& = Abs(n - 1)
    If NumberOnly Then
     .glistN.UseTab = False
        .NumberOnly = True
        .NumberIntOnly = UseIntOnly
        OLDV$ = v$
        ScreenEdit bb, v$, X&, Y&, wi& - 1, Hi&, wh&, , n, shiftlittle
        If Result = 99 Then v$ = OLDV$
        .NumberIntOnly = False
        .NumberOnly = False
    Else
    .glistN.UseTab = True
        OLDV$ = v$
        ScreenEdit bb, v$, X&, Y&, wi& - 1, Hi&, wh&, , n, shiftlittle
        If Result = 99 And Hi& = wi& Then v$ = OLDV$
    End If
    .showparagraph = oldshow
    .glistN.UseTab = UseTabInForm1Text1
End With
iText = v$
End With
End Function
Sub ScreenEditDOC(bstack As basetask, aaa As Variant, X&, Y&, x1&, y1&, Optional l As Long = 0, Optional usecol As Boolean = False, Optional col As Long)
On Error Resume Next
Dim ot As Boolean, back As New Document, i As Long, d As Object
Dim prive As basket
Set d = bstack.Owner
prive = players(GetCode(d))
With prive
Dim oldesc As Boolean
oldesc = escok
escok = False
' we have a limit here
If Not aaa.IsEmpty Then
For i = 1 To aaa.DocParagraphs
back.AppendParagraph aaa.TextParagraph(i)
Next i
End If
i = back.LastSelStart
Dim aaaa As Document, tcol As Long, trans As Boolean
If usecol Then tcol = mycolor(col) Else tcol = d.backcolor
If Not Form1.Visible Then newshow basestack1

'd.Enabled = False
If Not bstack.toback Then d.TabStop = False
If d Is Form1 Then
d.lockme = True
Else
d.Parent.lockme = True
End If
If y1& - Y& = 0 Then Y& = Y& - 1: If y1& < 0 Then Y& = Y& + 1: y1& = y1& + 1
TextEditLineHeight = y1& - Y& + 1

With Form1.TEXT1
'MyDoEvents
ProcTask2 bstack
.glistN.UseTab = True
Hook Form1.hWnd, Nothing '.glistN
.AutoNumber = Not Form1.EditTextWord

.UsedAsTextBox = False
.glistN.LeftMarginPixels = 10
.glistN.maxchar = 0
If d.ForeColor = tcol Then
Set Form1.Point2Me = d
If d.name = "Form1" Then
.glistN.SkipForm = False
Else
.glistN.SkipForm = True
End If
Form1.TEXT1.glistN.BackStyle = 1
End If
Dim scope As Long
scope = ChooseByHue(d.ForeColor, rgb(16, 12, 8), rgb(253, 245, 232))
If d.backcolor = ChooseByHue(scope, d.backcolor, rgb(128, 128, 128)) Then
If lightconv(scope) > 192 Then
scope = lightconv(scope) - 128
.glistN.CapColor = rgb(128 + scope / 2, 128 + scope / 2, 128 + scope / 2)
Else
.glistN.CapColor = scope
End If
Else
scope = lightconv(scope) - 128

If scope > 0 Then
.glistN.CapColor = rgb(128 + scope / 2, 128 + scope / 2, 128 + scope / 2)
Else
.glistN.CapColor = rgb(128, 128, 128)
End If
End If
.SelectionColor = .glistN.CapColor
.glistN.addpixels = 2 * prive.uMineLineSpace / dv15
.EditDoc = True
.enabled = True
.glistN.ZOrder 0

.backcolor = tcol

.ForeColor = d.ForeColor
Form1.SetText1
.glistN.overrideTextHeight = fonttest.TextHeight("fj")
.Font.name = d.Font.name
.Font.Size = d.Font.Size ' SZ 'Int(d.font.Size) Why
.Font.charset = d.Font.charset
.Font.Italic = d.Font.Italic
.Font.bold = d.Font.bold
.Font.name = d.Font.name
.Font.charset = d.Font.charset
.Font.Size = prive.SZ
With prive
If bstack.toback Then

Form1.TEXT1.Move X& * .Xt, Y& * .Yt, (x1& - X&) * .Xt + .Xt, (y1& - Y&) * .Yt + .Yt
Else
Form1.TEXT1.Move X& * .Xt + d.Left, Y& * .Yt + d.top, (x1& - X&) * .Xt + .Xt, (y1& - Y&) * .Yt + .Yt
End If
End With
If d.ForeColor = tcol Then
Form1.TEXT1.glistN.RepaintFromOut d.Image, d.Left, d.top
End If

Set .mDoc = aaa
.mDoc.ColorEvent = True
.nowrap = False


With Form1.TEXT1
.Form1mn1Enabled = False
.Form1mn2Enabled = False
.Form1mn3Enabled = Clipboard.GetFormat(13) Or Clipboard.GetFormat(1)
End With

Form1.KeyPreview = False
NOEDIT = False

.WrapAll
.Render

.Visible = True
.SetFocus
If l <> 0 Then
    If l > 0 Then
        If aaa.SizeCRLF < l Then l = aaa.SizeCRLF
        
        .SelStart = l
        Else
        .SelStart = 0
    End If
Else
If aaa.SizeCRLF < .LastSelStart Then
.SelStart = 1
Else
 .SelStart = .LastSelStart
End If
End If
    .ResetUndoRedo

End With
''MyDoEvents
ProcTask2 bstack
CancelEDIT = False
Do
BLOCKkey = False

 If bstack.IamThread Then If myexit(bstack) Then GoTo contScreenEditThere1

ProcTask2 bstack


'End If

Loop Until NOEDIT
 NOEXECUTION = False
 BLOCKkey = True
While KeyPressed(&H1B)
ProcTask2 bstack

Wend
BLOCKkey = False
contScreenEditThere1:
TaskMaster.RestEnd1
If Form1.TEXT1.Visible Then Form1.TEXT1.Visible = False
 l = Form1.TEXT1.LastSelStart


If d Is Form1 Then
d.lockme = False
Else
d.Parent.lockme = False
End If
If Not CancelEDIT Then

Else
Set aaa = back
back.LastSelStart = i
End If
Set Form1.TEXT1.mDoc = New Document
Form1.TEXT1.glistN.UseTab = UseTabInForm1Text1
Form1.TEXT1.glistN.BackStyle = 0
Set Form1.Point2Me = Nothing
UnHook Form1.hWnd
Form1.KeyPreview = True

INK$ = vbNullString
escok = oldesc
Set d = Nothing
End With
End Sub
Sub ScreenEdit(bstack As basetask, a$, X&, Y&, x1&, y1&, Optional l As Long = 0, Optional changelinefeeds As Long = 0, Optional maxchar As Long = 0, Optional ExcludeThisLeft As Long = 0, Optional internal As Boolean = False)
On Error Resume Next
' allways a$ enter with crlf,but exit with crlf or cr or lf depents from changelinefeeds
Dim oldesc As Boolean, d As Object
Set d = bstack.Owner

''SetTextSZ d, Sz

Dim prive As basket
prive = players(GetCode(d))
oldesc = escok
escok = False
Dim ot As Boolean

If Not bstack.toback Then
d.TabStop = False
d.Parent.lockme = True
Else
d.lockme = True
End If
If Not Form1.Visible Then newshow basestack1
d.Visible = True
If d.Visible Then d.SetFocus
With Form1.TEXT1
'MyDoEvents
ProcTask2 bstack
Hook Form1.hWnd, Nothing
'.Filename = VbNullString
.AutoNumber = Not Form1.EditTextWord

If maxchar > 0 Then
ot = .glistN.DragEnabled
 .glistN.DragEnabled = True
y1& = Y&
TextEditLineHeight = 1
.glistN.BorderStyle = 0
.glistN.BackStyle = 1
Set Form1.Point2Me = d
If d.name = "Form1" Then
.glistN.SkipForm = False
Else
.glistN.SkipForm = True
End If

.glistN.HeadLine = vbNullString
.glistN.HeadLine = vbNullString
.glistN.LeftMarginPixels = 1
.glistN.maxchar = maxchar
.nowrap = True
If Len(a$) > maxchar Then
a$ = Left$(a$, maxchar)
End If

l = Len(a$)


.UsedAsTextBox = True

Else
.glistN.BorderStyle = 0
.glistN.BackStyle = 0

If y1& - Y& = 0 Then Y& = Y& - 1: If y1& < 0 Then Y& = Y& + 1: y1& = y1& + 1
TextEditLineHeight = y1& - Y& + 1
.UsedAsTextBox = False
.glistN.LeftMarginPixels = 10
.glistN.maxchar = 0

End If

If Form1.EditTextWord Then
.glistN.WordCharLeft = ConCat(":", "{", "}", "[", "]", ",", "(", ")", "!", ";", "=", ">", "<", """", " ", "+", "-", "/", "*", "^", "$", "%", "_", "@")
.glistN.WordCharRight = ConCat(".", ":", "{", "}", "[", "]", ",", ")", "!", ";", "=", ">", "<", """", " ", "+", "-", "/", "*", "^", "$", "%", "_")
.glistN.WordCharRightButIncluded = vbNullString
.glistN.WordCharLeftButIncluded = vbNullString

Else
.glistN.WordCharLeft = ConCat(":", "{", "}", "[", "]", ",", "(", ")", "!", ";", "=", ">", "<", "'", """", " ", "+", "-", "/", "*", "^", "@", Chr$(9), "#", "%", "&")
.glistN.WordCharRight = ConCat(":", "{", "}", "[", "]", ",", ")", "!", ";", "=", ">", "<", "'", """", " ", "+", "-", "/", "*", "^", Chr$(9), "#")
.glistN.WordCharRightButIncluded = "(" ' so aaa(sdd) give aaa( as word
.glistN.WordCharLeftButIncluded = "#"
End If

Dim scope As Long
scope = ChooseByHue(d.ForeColor, rgb(16, 12, 8), rgb(253, 245, 232))
If d.backcolor = ChooseByHue(scope, d.backcolor, rgb(128, 128, 128)) Then
If lightconv(scope) > 192 Then
scope = lightconv(scope) - 128
.glistN.CapColor = rgb(128 + scope / 2, 128 + scope / 2, 128 + scope / 2)
Else
.glistN.CapColor = scope
End If
Else
scope = lightconv(scope) - 128

If scope > 0 Then
.glistN.CapColor = rgb(128 + scope / 2, 128 + scope / 2, 128 + scope / 2)
Else
.glistN.CapColor = rgb(128, 128, 128)
End If
End If
.SelectionColor = .glistN.CapColor
.glistN.addpixels = 2 * prive.uMineLineSpace / dv15
.EditDoc = True
.enabled = True
'.glistN.AddPixels = 0
.glistN.ZOrder 0
.backcolor = d.backcolor
.ForeColor = d.ForeColor
.Font.name = d.Font.name
Form1.SetText1
.glistN.overrideTextHeight = fonttest.TextHeight("fj")
.Font.Size = d.Font.Size ' SZ 'Int(d.font.Size) Why
.Font.charset = d.Font.charset
.Font.Italic = d.Font.Italic
.Font.bold = d.Font.bold

.Font.name = d.Font.name

.Font.charset = d.Font.charset
.Font.Size = prive.SZ 'Int(d.font.Size)
If bstack.toback Then
If maxchar > 0 Then

.Move X& * prive.Xt - ExcludeThisLeft, Y& * prive.Yt, (x1& - X&) * prive.Xt + prive.Xt, (y1& - Y&) * prive.Yt + prive.Yt
.glistN.RepaintFromOut d.Image, 0, 0
Else
.Move X& * prive.Xt, Y& * prive.Yt, (x1& - X&) * prive.Xt + prive.Xt, (y1& - Y&) * prive.Yt + prive.Yt
End If
Else
If maxchar > 0 Then
.Move X& * prive.Xt + d.Left - ExcludeThisLeft, Y& * prive.Yt + d.top, (x1& - X&) * prive.Xt + prive.Xt, (y1& - Y&) * prive.Yt + prive.Yt
.glistN.RepaintFromOut d.Image, d.Left, d.top
Else
.Move X& * prive.Xt + d.Left, Y& * prive.Yt + d.top, (x1& - X&) * prive.Xt + prive.Xt, (y1& - Y&) * prive.Yt + prive.Yt
End If
End If
If a$ <> "" Then
If .Text <> a$ Then .LastSelStart = 0
If internal Then
.Text2 = a$
Else
.Text = a$
End If
Else
.Text = vbNullString
.LastSelStart = 0
End If
'.glistN.NoFreeMoveUpDown = True

'With Form1.TEXT1
.Form1mn1Enabled = False
.Form1mn2Enabled = False
.Form1mn3Enabled = Clipboard.GetFormat(13) Or Clipboard.GetFormat(1)
'End With

Form1.KeyPreview = False

NOEDIT = False

If maxchar = 0 Then
If .nowrap Then
.nowrap = False
End If
.Charpos = 0
If Len(a$) < 100000 Then .Render
Else
.Render
End If

.Visible = True
''MyDoEvents
ProcTask2 bstack
.SetFocus

If l <> 0 Then
    If l > 0 Then
        If Len(a$) < l Then l = Len(a$)
        .SelStart = l
                Else
        .SelStart = 0
    End If
Else
If Len(a$) < .LastSelStart Then
.SelStart = 1
l = Len(a$)
Else
    .SelStart = .LastSelStart
End If
End If
    .ResetUndoRedo



End With
'MyDoEvents
ProcTask2 bstack
CancelEDIT = False
Dim timeOut As Long


Do
BLOCKkey = False

 If bstack.IamThread Then If myexit(bstack) Then GoTo contScreenEditThere

ProcTask2 bstack

 Loop Until NOEDIT
 NOEXECUTION = False
 BLOCKkey = True
While KeyPressed(&H1B)
'
ProcTask2 bstack


Wend
BLOCKkey = False
contScreenEditThere:
TaskMaster.RestEnd1
If Form1.TEXT1.Visible Then Form1.TEXT1.Visible = False

 l = Form1.TEXT1.LastSelStart

If bstack.toback Then
d.lockme = False
Else
d.Parent.lockme = False
End If
If Not CancelEDIT Then

If changelinefeeds > 10 Then
a$ = Form1.TEXT1.TextFormatBreak(vbCr)
ElseIf changelinefeeds > 9 Then
a$ = Form1.TEXT1.TextFormatBreak(vbLf)
Else
If changelinefeeds = -1 Then changelinefeeds = 0
a$ = Form1.TEXT1.Text
End If
Else
changelinefeeds = -1
End If

Form1.KeyPreview = True
If maxchar > 0 Then Form1.TEXT1.glistN.DragEnabled = ot

UnHook Form1.hWnd
INK$ = vbNullString
Form1.TEXT1.glistN.UseTab = False
escok = oldesc
Set d = Nothing
End Sub

Function blockCheck(ByVal s$, ByVal Lang As Long, countlines As Long, Optional ByVal sbname$ = vbNullString, Optional Column As Long) As Boolean
If s$ = vbNullString Then blockCheck = True: Exit Function
Dim i As Long, j As Long, c As Long, b$, resp&
Dim openpar As Long, oldi As Long, lastlabel$, oldjump As Boolean, st As Long, stc As Long
Dim paren As New mStiva2
countlines = 1
Column = 0
Lang = Not Lang
Dim a1 As Boolean
Dim jump As Boolean
If Trim$(s$) = vbNullString Then Exit Function
c = Len(s$)
a1 = True
i = 1
Do
Column = Column + 1
'Debug.Print Mid$(s$, i, 1)
Select Case AscW(Mid$(s$, i, 1))
Case 10
Column = 0
Case 13
lastlabel$ = ""
If openpar <> 0 Then
GoTo pareprob
End If
oldjump = False
jump = False
If Len(s$) > i + 1 Then countlines = countlines + 1
Column = 0
Case 58
lastlabel$ = ""
oldjump = False
jump = False
Case 32, 160, 9
If Len(lastlabel$) > 0 Then
lastlabel$ = myUcase(lastlabel$)
If Not ismine1(lastlabel$) Then
If Not ismine2(lastlabel$) Then
If Not ismine22(lastlabel$) Then
    jump = Not oldjump
Else
    oldjump = True
    jump = False
End If
Else
oldjump = True
jump = False
End If
Else
oldjump = False
jump = False
End If
lastlabel$ = ""
End If
Case 34
lastlabel$ = ""
oldi = i
Do While i < c
i = i + 1
Select Case AscW(Mid$(s$, i, 1))
Case 34
Exit Do
Case 13

checkit:
    If Not Lang Then
        b$ = sbname$ + "Problem in string in paragraph " + CStr(countlines)
    Else
        b$ = sbname$ + "Πρόβλημα με το αλφαριθμητικό στη παράγραφο " + CStr(countlines)
    End If
    resp& = ask(b$, True)
If resp& <> 4 Then
blockCheck = True
End If
Exit Function
End Select

Loop
If oldi <> i Then
Else
i = oldi + 1
GoTo checkit
End If

Case 40
lastlabel$ = ""
jump = True
openpar = openpar + 1
paren.PushVal countlines
Case 41
lastlabel$ = ""
openpar = openpar - 1
If openpar = 0 Then jump = False
If openpar < 0 Then Exit Do
paren.drop 1

Case 39, 92
lastlabel$ = ""
Do While i < c
i = i + 1
If Mid$(s$, i, 2) = vbCrLf Then Exit Do
Loop
countlines = countlines + 1
If openpar > 0 Then Exit Do
Case 61, 43, 44
lastlabel$ = ""
jump = True
Case 123
If Len(lastlabel$) > 0 Then
lastlabel$ = myUcase(lastlabel$)
If Not ismine1(lastlabel$) Then
If Not ismine2(lastlabel$) Then
If Not ismine22(lastlabel$) Then
    jump = Not oldjump
Else
    oldjump = True
    jump = False
End If
Else
oldjump = True
jump = False
End If
Else
oldjump = False
jump = False
End If
lastlabel$ = ""
End If

If jump Then
jump = False
' we have a multiline text
Dim target As Long
target = j
st = countlines
stc = Column
    Do
    Select Case AscW(Mid$(s$, i, 1))
            Case 34
            Do While i < c
            i = i + 1
            If AscW(Mid$(s$, i, 1)) = 34 Then Exit Do
            If AscW(Mid$(s$, i, 1)) = 13 Then GoTo checkit
            Loop
        Case 13
        countlines = countlines + 1
        Case 123
        j = j - 1
        Case 125
        j = j + 1: If j = target Then Exit Do
    End Select
    i = i + 1
    Loop Until i > c
    If j <> target Then
    countlines = st
    Column = st
    Exit Do
    End If
    Else
j = j - 1
oldjump = False
End If
Case 13

Case 125
If openpar <> 0 And j > 0 Then
pareprob:
If paren.count > 0 Then countlines = paren.PopVal
If Not Lang Then
        b$ = sbname$ + "Problem in parenthesis in paragraph" + Str$(countlines)
    Else
        b$ = sbname$ + "Πρόβλημα με τις παρενθέσεις στη παράγραφο" + Str$(countlines)
    End If
    resp& = ask(b$, True)
If resp& <> 4 Then
blockCheck = True
End If
    Exit Function

End If
j = j + 1: If j = 1 Then Exit Do
Case 65 To 93, 97 To 122, Is > 127
jump = False
lastlabel$ = lastlabel$ + Mid$(s$, i, 1)
Case 46
jump = False
lastlabel$ = lastlabel$ + Mid$(s$, i, 1)

Case 48 To 57, 95
jump = False
If Len(lastlabel$) > 0 Then lastlabel$ = lastlabel$ + Mid$(s$, i, 1)
Case Else
jump = False
lastlabel$ = ""
End Select
i = i + 1
Loop Until i > c
If openpar <> 0 Then
GoTo pareprob
End If
If j = 0 Then

ElseIf j < 0 Then
    If Not Lang Then
        b$ = sbname$ + "Problem in blocks - look } are less " + CStr(Abs(j))
    Else
        b$ = sbname$ + "Πρόβλημα με τα τμήματα - δες τα } είναι λιγότερα " + CStr(Abs(j))
    End If
resp& = ask(b$, True)
Else
If Not Lang Then
b$ = sbname$ + "Problem in blocks - look { are less " + CStr(j)
Else
b$ = sbname$ + "Πρόβλημα με τα τμήματα - δες τα { είναι λιγότερα " + CStr(j)
End If
resp& = ask(b$, True)
End If
If resp& <> 4 Then
blockCheck = True
End If

End Function

Sub ListChoise(bstack As basetask, a$, X&, Y&, x1&, y1&)
On Error Resume Next
Dim d As Object, oldh As Long
Dim s$, prive As basket
If NOEXECUTION Then Exit Sub
Set d = bstack.Owner
prive = players(GetCode(d))
Hook Form1.hWnd, Form1.List1
Dim ot As Boolean, drop
With Form1.List1
.Font.name = d.Font.name
Form1.Font.charset = d.Font.charset
Form1.Font.Strikethrough = False
.Font.Size = d.Font.Size
.Font.name = d.Font.name
Form1.Font.charset = d.Font.charset
.Font.Size = d.Font.Size
If LEVCOLMENU < 2 Then .backcolor = d.ForeColor
If LEVCOLMENU < 3 Then .ForeColor = d.backcolor
.Font.bold = d.Font.bold
.Font.Italic = d.Font.Italic
.addpixels = 2 * prive.uMineLineSpace / dv15
.VerticalCenterText = True
If d.Visible = False Then d.Visible = True
.StickBar = True
s$ = .HeadLine
.HeadLine = vbNullString
.HeadLine = s$
.enabled = False
If .Visible Then
If .BorderStyle = 0 Then

Else
End If

Else

If .BorderStyle = 0 Then
.Move X& * prive.Xt + d.Left, Y& * prive.Yt + d.top, (x1& - X&) * prive.Xt + prive.Xt, (y1& - Y&) * prive.Yt + prive.Yt + .HeadlineHeight * dv15
Else
.Move X& * prive.Xt - dv15 + d.Left, Y& * prive.Yt - dv15 + d.top, (x1& - X&) * prive.Xt + prive.Xt + 2 * dv15, (y1& - Y&) * prive.Yt + prive.Yt + 2 * dv15 + .HeadlineHeight * dv15
End If
End If
.enabled = True
.ShowBar = False

If .LeaveonChoose Then
.CalcAndShowBar
Exit Sub
End If



ot = Targets
Targets = False

.PanPos = 0

If .ListIndex < 0 Then
.ShowThis 1
Else
.ShowThis .ListIndex + 1
End If
.Visible = True
.ZOrder 0
NOEDIT = False
.Tag = a$

If a$ = vbNullString Then
    drop = mouse
    MyDoEvents
    ' Form1.KeyPreview = False
    .enabled = True
    .SetFocus
    .LeaveonChoose = True
    If .HeadLine <> "" Then
    oldh = 0
    Else
    oldh = .HeadlineHeight
    End If
    Else
        .enabled = True
    .SetFocus
    .LeaveonChoose = False
    
    End If
    .ShowMe
            If bstack.TaskMain Or TaskMaster.Processing Then
            If TaskMaster.QueueCount > 0 Then
            mywait bstack, 100
              Else
            MyDoEvents
            End If
        Else
         DoEvents
         Sleep 1
         End If

    If .HeadlineHeight <> oldh Then
    If .BorderStyle = 0 Then
    If ((y1& - Y&) * prive.Yt + prive.Yt + 2 * dv15 + .HeadlineHeight * dv15) + .top > ScrY() Then
    .Move .Left, .top - (((y1& - Y&) * prive.Yt + prive.Yt + 2 * dv15 + .HeadlineHeight * dv15) + .top - ScrY()), (x1& - X&) * prive.Xt + prive.Xt, (y1& - Y&) * prive.Yt + prive.Yt + .HeadlineHeight * dv15
    Else
.Move .Left, .top, (x1& - X&) * prive.Xt + prive.Xt, (y1& - Y&) * prive.Yt + prive.Yt + .HeadlineHeight * dv15
End If
Else
If ((y1& - Y&) * prive.Yt + prive.Yt + 2 * dv15 + .HeadlineHeight * dv15) + .top > ScrY() Then
.Move .Left, .top - (((y1& - Y&) * prive.Yt + prive.Yt + 2 * dv15 + .HeadlineHeight * dv15) + .top - ScrY()), (x1& - X&) * prive.Xt + prive.Xt + 2 * dv15, (y1& - Y&) * prive.Yt + prive.Yt + 2 * dv15 + .HeadlineHeight * dv15
Else
.Move .Left, .top, (x1& - X&) * prive.Xt + prive.Xt + 2 * dv15, (y1& - Y&) * prive.Yt + prive.Yt + 2 * dv15 + .HeadlineHeight * dv15
End If
End If
  
oldh = .HeadlineHeight
    End If
    .FloatLimitTop = Form1.Height - prive.Yt * 2
     .FloatLimitLeft = Form1.Width - prive.Xt * 2
    MyDoEvents
    End With
If a$ = vbNullString Then
    Do
        If bstack.TaskMain Or TaskMaster.Processing Then
            If TaskMaster.QueueCount > 0 Then
          mywait bstack, 2
             TaskMaster.RestEnd1
   TaskMaster.TimerTick
TaskMaster.rest
''SleepWait 1
  Sleep 1
              Else
            MyDoEvents
            End If
        Else
         DoEvents
                  Sleep 1
         End If
    
    Loop Until Form1.List1.Visible = False
    If Not NOEXECUTION Then MOUT = False
    Do
    drop = mouse
    MyDoEvents
    Loop Until drop = 0 Or MOUT
    MOUT = False
    While KeyPressed(&H1B)
ProcTask2 bstack
Wend
MOUT = False: NOEXECUTION = False
    If Form1.List1.ListIndex >= 0 Then
    a$ = Form1.List1.list(Form1.List1.ListIndex)
    Else
    a$ = vbNullString
    End If
   Form1.List1.enabled = False
    Else
        Form1.List1.enabled = True
    
  If a$ = vbNullString Then
  Form1.List1.SetFocus
  Form1.List1.LeaveonChoose = True
  Else
  d.TabStop = True
  End If
  End If
NOEDIT = True
Set d = Nothing
UnHook Form1.hWnd
Form1.KeyPreview = True
Targets = ot
End Sub
Private Sub mywait11(bstack As basetask, PP As Double)
Dim p As Boolean, e As Boolean
On Error Resume Next
If bstack.Process Is Nothing Then
''If extreme Then MyDoEvents
If PP = 0 Then Exit Sub
Else

Err.clear
p = bstack.Process.Done
If Err.Number = 0 Then
e = True
If p <> 0 Then
Exit Sub
End If
End If
End If
PP = PP + CDbl(timeGetTime)

Do


If TaskMaster.Processing And Not bstack.TaskMain Then
        If Not bstack.toprinter Then bstack.Owner.Refresh
        'If TaskMaster.tickdrop > 0 Then TaskMaster.tickdrop
        TaskMaster.TimerTick  'Now
       ' SleepWait 1
       MyDoEvents
       
Else
        ' SleepWait 1
        MyDoEvents
        End If
If e Then
p = bstack.Process.Done
If Err.Number = 0 Then
If p <> 0 Then
Exit Do
End If
End If
End If
Loop Until PP <= CDbl(timeGetTime) Or NOEXECUTION

                       If exWnd <> 0 Then
                MyTitle$ bstack
                End If
End Sub
Public Sub WaitDialog(bstack As basetask)
Dim oldesc As Boolean
oldesc = escok
escok = False
Dim d As Object
Set d = bstack.Owner
Dim ot As Boolean, drop
ot = Targets
Targets = False  ' do not use targets for now
'NOEDIT = False
    drop = mouse
    ''SleepWait3 100
    Sleep 1
    If bstack.ThreadsNumber = 0 Then
    If Not (bstack.toback Or bstack.toprinter) Then If bstack.Owner.Visible Then bstack.Owner.Refresh

    End If
    Dim mycode As Double, oldcodeid As Double
mycode = Rnd * 1233312231
oldcodeid = Modalid
Dim X As Form, zz As Form
Set zz = Screen.ActiveForm
For Each X In Forms
        If X.Visible And X.name = "GuiM2000" Then
                                   If X.Enablecontrol Then
                                        X.Modal = mycode
                                        X.Enablecontrol = False
                                    End If
        End If
Next X
On Error Resume Next
If zz.enabled Then zz.SetFocus
Set zz = Nothing
      Do
   

            mywait11 bstack, 5
      Sleep 1
    
    Loop Until loadfileiamloaded = False Or LastErNum <> 0
    Modalid = mycode
    MOUT = False
    Do
    drop = mouse Or KeyPressed(&H1B)
    MyDoEvents

    Loop Until drop = 0 Or MOUT Or LastErNum <> 0
 ' NOEDIT = True
 BLOCKkey = True

While KeyPressed(&H1B)

ProcTask2 bstack
NOEXECUTION = False
Wend
Dim z As Form
Set z = Nothing

           For Each X In Forms
            If X.Visible And X.name = "GuiM2000" Then
                If Not X.Enablecontrol Then
                        X.TestModal mycode
                End If
            End If
            Next X
          Modalid = oldcodeid

BLOCKkey = False
escok = oldesc
INK$ = vbNullString
If Form1.Visible Then Form1.KeyPreview = Not Form1.gList1.Visible
Targets = ot
 mywait11 bstack, 5
End Sub

Public Sub FrameText(dd As Object, ByVal Size As Single, X As Long, Y As Long, cc As Long, Optional myCut As Boolean = False)
Dim i As Long, mymul As Long

If dd Is Form1.PrinterDocument1 Then
' check this please
dd.Width = X
dd.Height = Y
Pr_Back dd, Size
Exit Sub
End If


Dim basketcode As Long
basketcode = GetCode(dd)
With players(basketcode)
.curpos = 0
.currow = 0
.XGRAPH = 0
.YGRAPH = 0
If X = 0 Then
X = dd.Width
Y = dd.Height
End If

.mysplit = 0

''dd.BackColor = 0 '' mycolor(cc)    ' check if paper...

.Paper = mycolor(cc)
dd.CurrentX = 0
dd.CurrentY = 0

''ClearScreenNew dd, mybasket, cc
dd.CurrentY = 0
dd.Font.Size = Size
Size = dd.Font.Size

''Sleep 1  '' USED TO GIVE TIME TO LOAD FONT
If fonttest.FontName = dd.Font.name And dd.Font.Size = fonttest.Font.Size Then
Else
StoreFont dd.Font.name, Size, dd.Font.charset
End If
.Yt = fonttest.TextHeight("fj")
.Xt = fonttest.TextWidth("W")

While TextHeight(fonttest, "fj") / (.Yt / 2 + dv15) < dv
Size = Size + 0.2
fonttest.Font.Size = Size
Wend
dd.Font.Size = Size
.Yt = TextHeight(fonttest, "fj")
.Xt = fonttest.TextWidth("W") + dv15

.mx = Int(X / .Xt)
.My = Int(Y / (.Yt + .MineLineSpace * 2))
.Yt = .Yt + .MineLineSpace * 2
If .mx < 2 Then .mx = 2: X = 2 * .Xt
If .My < 2 Then .My = 2: Y = 2 * .Yt
If (.mx Mod 2) = 1 And .mx > 1 Then
.mx = .mx - 1
End If
mymul = Int(.mx / 8)
If mymul = 1 Then mymul = 2
If mymul = 0 Then
.Column = .mx \ 2 - 1
Else
.Column = Int(.mx / mymul)

While (.mx Mod .Column) > 0 And (.mx / .Column >= 3)
.Column = .Column + 1
Wend
End If
If .Column = 0 Then .Column = .mx
' second stage
If .mx Mod .Column > 0 Then


If .mx Mod 4 <> 0 Then .mx = 4 * (.mx \ 4)
If .mx < 4 Then .mx = 4
'.My = Int(y / (.Yt + .MineLineSpace * 2))
'.Yt = .Yt + .MineLineSpace * 2
If .mx < 2 Then .mx = 2: X = 2 * .Xt
If .My < 2 Then .My = 2: Y = 2 * .Yt
If (.mx Mod 2) = 1 And .mx > 1 Then
.mx = .mx - 1
End If
mymul = Int(.mx / 8)
If mymul = 1 Then mymul = 2
If mymul = 0 Then
.Column = .mx \ 2 - 1
Else
.Column = Int(.mx / mymul)

While (.mx Mod .Column) > 0 And (.mx / .Column >= 3)
.Column = .Column + 1
Wend
End If
If .Column = 0 Then .Column = .mx

End If

.Column = .Column - 1 ' FOR PRINT 0 TO COLUMN-1

If .Column < 4 Then .Column = 4


.SZ = Size

If dd.name = "Form1" Then
' no change
Else
If dd.name <> "dSprite" And Typename(dd) <> "GuiM2000" Then
Dim mmxx As Long, mmyy As Long, XX As Long, YY As Long
mmxx = .mx * CLng(.Xt)
mmyy = .My * CLng(.Yt)
XX = (dd.Parent.ScaleWidth - mmxx) \ 2
YY = (dd.Parent.ScaleHeight - mmyy) \ 2
dd.Move XX, YY, mmxx, mmyy
ElseIf myCut Then
Dim mmxx1, mmyy1
mmxx1 = .mx * .Xt
mmyy1 = .My * .Yt
dd.Move dd.Left, dd.top, mmxx1, mmyy1
'dd.width = .mx * .Xt
'dd.Height = .My * .Yt
End If

End If

.MAXXGRAPH = dd.Width
.MAXYGRAPH = dd.Height
.FTEXT = 0
.FTXT = vbNullString

Form1.MY_BACK.ClearUp
If dd.Visible Then
ClearScr dd, .Paper
Else
dd.backcolor = .Paper
End If
End With



End Sub

Sub Pr_Back(dd As Object, Optional msize As Single = 0)

SetText dd
If msize <> 0 Then players(GetCode(dd)).SZ = msize
If msize > 0 Then
SetTextSZ dd, msize
End If

End Sub
Function INKEY$()
' αυτή η συνάρτηση θα αδειάσει τον προσωρινό χώρο πληκτρολογίσεων, που μπορεί να είναι πολλά πλήκτρα..
' θα επιστρέψει το πρώτο από αυτά ή τίποτα.
' Χρησιμοποιείται παντού όπου διαβάζουμε το πληκτρολόγιο

If MKEY$ <> "" Then ' κοιτάει να αδειάσει τον προσωρινό χώρο MKEY$
' αν έχει κάτι τότε το λαμβάνει τον αδείαζε βάζοντας τον στο τέλος του INK$
' και αδείαζουμε το MKEY$
    INK$ = MKEY$ & INK$
    MKEY$ = vbNullString
End If
' τώρα θα ασχοληθούμε με το INK$ αν έχει τίποτα
If INK$ <> "" Then
' ειδική περίπτωση αν έχουμε 0 στο πρώτο Byte, έχουμε ειδικό κ
    If Asc(INK$) = 0 Then
        INKEY$ = Left$(INK$, 2)
        INK$ = Mid$(INK$, 3)
    Else
    ' αλλιώς σηκώνουμε ένα χαρακτήρα με ότι έχει ακόμα
    INKEY$ = PopOne(INK$)
    
   
        
    End If
Else
    'Αν δεν έχουμε τίποτα...δεν κάνουμε τίποτα...γυρίζουμε το τίποτα!
    INKEY$ = vbNullString
End If

End Function
Function UINKEY$()
' mink$ used for reinput keystrokes
' MINK$ = MINK$ & UINK$
If UKEY$ <> "" Then MINK$ = MINK$ + UKEY$: UKEY$ = vbNullString
If MINK$ <> "" Then
If AscW(MINK$) = 0 Then
    UINKEY$ = Left$(MINK$, 2)
    MINK$ = Mid$(MINK$, 3)
Else
    UINKEY$ = Left$(MINK$, 1)
    MINK$ = Mid$(MINK$, 2)
End If
Else
    UINKEY$ = vbNullString
End If

End Function

Function QUERY(bstack As basetask, Prompt$, s$, m&, Optional USELIST As Boolean = True, Optional endchars As String = vbCr, Optional excludechars As String = vbNullString, Optional checknumber As Boolean = False) As String
'NoAction = True
On Error Resume Next
Dim dX As Long, dY As Long, safe$, oldREFRESHRATE As Double
oldREFRESHRATE = REFRESHRATE

If excludechars = vbNullString Then excludechars = Chr$(0)
If QUERYLIST = vbNullString Then QUERYLIST = Chr$(13): LASTQUERYLIST = 1
Dim q1 As Long, sp$, once As Boolean, dq As Object
 
Set dq = bstack.Owner
SetText dq
Dim basketcode As Long, prive As basket
prive = players(GetCode(dq))
With prive
If .currow >= .My Or .lastprint Then crNew bstack, prive: .lastprint = False
LCTbasketCur dq, prive
ins& = 0
Dim fr1 As Long, fr2 As Long, p As Double
UseEnter = False
If dq.name = "DIS" Then
If Form1.Visible = False Then
    If Not Form3.Visible Then
        Form1.Hide: Sleep 100
    Else
        'Form3.PREPARE
    End If

    If Form1.WindowState = vbMinimized Then Form1.WindowState = vbNormal
    Form1.Show , Form5
    If ttl Then
    If Form3.Visible Then
    If Not Form3.WindowState = 0 Then
    Form3.skiptimer = True: Form3.WindowState = 0
    End If
    End If
    End If
    MyDoEvents
    Sleep 100
    End If
Else
    Console = FindFormSScreen(Form1)
If Form1.top >= VirtualScreenHeight() Then Form1.Move ScrInfo(Console).Left, ScrInfo(Console).top
End If
If dq.Visible = False Then dq.Visible = True
If exWnd = 0 Then Form1.KeyPreview = True
QRY = True
If GetForegroundWindow = Form1.hWnd Then
If exWnd = 0 Then dq.SetFocus
End If


Dim DE$

PlainBaSket dq, prive, Prompt$, , , 0
dq.Refresh

 

INK$ = vbNullString
dq.FontTransparent = False

Dim a$
s$ = vbNullString
oldLCTCB dq, prive, 0
Do
If Not once Then
If USELIST Then
 DoEvents
  If Not iamactive Then
  Sleep 1
  Else
  If Not (bstack.IamChild Or bstack.IamAnEvent) Then Sleep 1
  End If
 ''If MKEY$ = VbNullString Then Dq.refresh
Else
If Not bstack.IamThread Then

 If Not iamactive Then
 If Not Form1.Visible Then
 If Form1.WindowState = 1 Then Form1.WindowState = 0
 If Form1.top > VirtualScreenHeight() - 100 Then Form1.top = ScrInfo(Console).top
 Form1.Visible = True
 If Form3.Visible Then Form3.skiptimer = True: Form3.WindowState = 0
 End If
 k1 = 0: MyDoEvents1 Form1, , True
 End If
If LastErNum <> 0 Then
      LCTCB dq, prive, -1: DestroyCaret
 oldLCTCB dq, prive, 0
Exit Do
End If
 Else
 
LCTbasketCur dq, prive                       ' here
 End If
 End If
 End If
If Not QRY Then HideCaret dq.hWnd:   Exit Do

 BLOCKkey = False
 If USELIST Then

 If Not once Then
 once = True

 If QUERYLIST <> "" Then  ' up down
 
    If INK = vbNullString Then MyDoEvents
If clickMe = 38 Then

 If Len(QUERYLIST) < LASTQUERYLIST Then LASTQUERYLIST = 2
  q1 = InStr(LASTQUERYLIST, QUERYLIST, vbCr)
         If q1 < 2 Or q1 <= LASTQUERYLIST Then
         q1 = 1: LASTQUERYLIST = 1
         End If
        MKEY$ = vbNullString
        INK = String$(RealLen(s$), 8) + Mid$(QUERYLIST, LASTQUERYLIST, q1 - LASTQUERYLIST)
        LASTQUERYLIST = q1 + 1

    ElseIf clickMe = 40 Then
    
    If LASTQUERYLIST < 3 Then LASTQUERYLIST = Len(QUERYLIST)
    q1 = InStrRev(QUERYLIST, vbCr, LASTQUERYLIST - 2)
         If q1 < 2 Then
                   q1 = Len(QUERYLIST)
         End If
         If q1 > 1 Then
         LASTQUERYLIST = InStrRev(QUERYLIST, vbCr, q1 - 1) + 1
         If LASTQUERYLIST < 2 Then LASTQUERYLIST = 2
         
        MKEY$ = vbNullString
        INK = String$(RealLen(s$), 8) + Mid$(QUERYLIST, LASTQUERYLIST, q1 - LASTQUERYLIST)
   LASTQUERYLIST = q1 + 1

      End If
 End If
 clickMe = -2
 End If
 
 ElseIf INK <> "" Then
 MKEY$ = vbNullString
 Else
 clickMe = 0
 once = False
 End If
 End If

  
againquery:
 a$ = INKEY$
 
If a$ = vbNullString Then
If TaskMaster Is Nothing Then Set TaskMaster = New TaskMaster
    If TaskMaster.QueueCount > 0 Then
  ProcTask2 bstack
  If Not NOEDIT Or Not QRY Then
  LCTCB dq, prive, -1: DestroyCaret
   oldLCTCB dq, prive, 0
  Exit Do
  End If
  SetText dq

LCTbasket dq, prive, .currow, .curpos
    Else
  
   End If
      If iamactive Then
 If ShowCaret(dq.hWnd) = 0 Then
 
   LCTCB dq, prive, 0
  End If
If Not bstack.IamThread Then

MyDoEvents1 Form1, , True
End If

 If Screen.ActiveForm Is Nothing Then
 iamactive = False:  If ShowCaret(dq.hWnd) <> 0 Then HideCaret dq.hWnd
Else
 
    If Not GetForegroundWindow = Screen.ActiveForm.hWnd Then
    iamactive = False:  If ShowCaret(dq.hWnd) <> 0 Then HideCaret dq.hWnd
  
    End If
    End If
    End If

  End If
    If bstack Is Nothing Then
    Set bstack = basestack1
    NOEXECUTION = True
    MOUT = True
     Modalid = 0
                         ShutEnabledGuiM2000
                         MyDoEvents
                         GoTo contqueryhere
    End If
   If bstack.IamThread Then If myexit(bstack) Then GoTo contqueryhere

If Screen.ActiveForm Is Nothing Then
iamactive = False
Else
If Screen.ActiveForm.name <> "Form1" Then
iamactive = False
Else
iamactive = GetForegroundWindow = Screen.ActiveForm.hWnd
End If
End If
If FKey > 0 Then
If FK$(FKey) <> "" Then
s$ = FK$(FKey)
FKey = 0
             ''  here
      LCTCB dq, prive, -1: DestroyCaret
 oldLCTCB dq, prive, 0
 Exit Do
End If
End If


dq.FontTransparent = False
If RealLen(a$) = 1 Or Len(a$) = 1 Or (RealLen(a$) = 0 And Len(a$) = 1 And Len(s$) > 1) Then
   '
   
   If Len(a$) = 1 Then
    If InStr(endchars, a$) > 0 Then
     If a$ = vbCr Then
     If a$ <> Left$(endchars, 1) Then
    
    a$ = Left$(endchars, 1)
     Else
      LCTCB dq, prive, -1: DestroyCaret
 oldLCTCB dq, prive, 0

        Exit Do
End If
     End If
     End If
     ElseIf a$ = vbCr Then
     a$ = Left$(endchars, 1)
     End If
    If Asc(a$) = 27 And escok Then
        
      LCTCB dq, prive, -1: DestroyCaret
 oldLCTCB dq, prive, 0
    s$ = vbNullString
    'If ExTarget Then End

    Exit Do
ElseIf Asc(a$) = 27 Then
a$ = Chr$(0)
End If
If a$ = Chr(8) Then
DE$ = " "
    If Len(s$) > 0 Then
    ExcludeOne s$

             LCTCB dq, prive, -1: DestroyCaret
            oldLCTCB dq, prive, 0

        
        .curpos = .curpos - 1
        If .curpos < 0 Then
            .curpos = .mx - 1: .currow = .currow - 1

            If .currow < .mysplit Then
                ScrollDownNew dq, prive
                PlainBaSket dq, prive, Right$(Prompt$ & s$, .mx - 1), , , 0
                DE$ = vbNullString
            End If
        End If

       LCTbasketCur dq, prive
        dX = .curpos
        dY = .currow
       PlainBaSket dq, prive, DE$, , , 0
       .curpos = dX
       .currow = dY
         
         
            oldLCTCB dq, prive, 0
            
    End If
End If
If safe$ <> "" Then
        a$ = 65
End If
If AscW(a$) > 31 And (RealLen(s$) < m& Or RealLen(a$, True) = 0) Then
If RealLen(a$, True) = 0 Then
If Asc(a$) = 63 And s$ <> "" Then
s$ = s$ & a$: a$ = s$: ExcludeOne s$: a$ = Mid$(a$, Len(s$) + 1)
s$ = s$ + a$
MKEY$ = vbNullString
'UINK = VbNullString
safe$ = a$
INK = Chr$(8)
Else
If s$ = vbNullString Then a$ = " "
GoTo cont12345
End If
Else
cont12345:
    If InStr(excludechars, a$) > 0 Then

    Else
            If checknumber Then
                    fr1 = 1
                    If (s$ = vbNullString And a$ = "-") Or IsNumberQuery(s$ + a$, fr1, p, fr2) Then
                            If fr2 - 1 = RealLen(s$) + 1 Or (s$ = vbNullString And a$ = "-") Then
   If ShowCaret(dq.hWnd) <> 0 Then DestroyCaret
                If a$ = "." Then
                If Not NoUseDec Then
                    If OverideDec Then
                    PlainBaSket dq, prive, NowDec$, , , 0
                    Else
                    PlainBaSket dq, prive, ".", , , 0
                    End If
                Else
                    PlainBaSket dq, prive, QueryDecString, , , 0
                End If
                Else
                   PlainBaSket dq, prive, a$, , , 0
                   End If
                   s$ = s$ & a$
                 
              oldLCTCB dq, prive, 0
                  LCTCB dq, prive, 0
GdiFlush
                            End If
                    
                    End If
            Else
            If ShowCaret(dq.hWnd) <> 0 Then DestroyCaret
                   If safe$ <> "" Then
        a$ = safe$: safe$ = vbNullString
End If
 If InStr(endchars, a$) = 0 Then PlainBaSket dq, prive, a$, , , 0: s$ = s$ & a$
              If .curpos >= .mx Then
                                .curpos = 0
                                .currow = .currow + 1
                            End If
              oldLCTCB dq, prive, 0
                  LCTCB dq, prive, 0
                  GdiFlush
                
            End If
    End If
End If
If InStr(endchars, a$) > 0 Then
    If a$ >= " " Then
                     PlainBaSket dq, prive, a$, , , 0
              
      LCTCB dq, prive, -1: DestroyCaret
                                GdiFlush
                                End If
QUERY = a$
Exit Do
End If
 .pageframe = 0
 End If
End If
If Not QRY Then
      LCTCB dq, prive, -1: DestroyCaret
 oldLCTCB dq, prive, 0
Exit Do
''HideCaret dq.hWnd:


End If
Loop


 
If Not QRY Then s$ = vbNullString
dq.FontTransparent = True
QRY = False

Call mouse

If s$ <> "" And USELIST Then
q1 = InStr(QUERYLIST, Chr$(13) + s$ & Chr$(13))
If q1 = 1 Then ' same place
ElseIf q1 > 1 Then ' reorder
sp$ = Mid$(QUERYLIST, q1 + RealLen(s$) + 1)
QUERYLIST = Chr$(13) + s$ & Mid$(QUERYLIST, 1, q1 - 1) + sp$
Else ' insert
QUERYLIST = Chr$(13) + s$ & QUERYLIST
End If
LASTQUERYLIST = 2
End If
End With
contqueryhere:
If Not bstack.IamThread Then
MyDoEvents1 Form1, , True
End If
REFRESHRATE = oldREFRESHRATE
If TaskMaster Is Nothing Then Exit Function
If TaskMaster.QueueCount > 0 Then TaskMaster.RestEnd
players(GetCode(dq)) = prive
Set dq = Nothing
TaskMaster.RestEnd1

End Function


Public Sub GetXYb(dd As Object, mb As basket, X As Long, Y As Long)
With mb
If dd.CurrentY Mod .Yt <= dv15 Then
Y = (dd.CurrentY) \ .Yt
Else
Y = (dd.CurrentY - .uMineLineSpace) \ .Yt
End If
X = dd.CurrentX \ .Xt

''
End With
End Sub
Public Sub GetXYb2(dd As Object, mb As basket, X As Long, Y As Long)
With mb
X = dd.CurrentX \ .Xt
Y = Int((dd.CurrentY / .Yt) + 0.5)
End With
End Sub
Sub Gradient(TheObject As Object, ByVal F&, ByVal t&, ByVal xx1&, ByVal xx2&, ByVal yy1&, ByVal yy2&, ByVal hor As Boolean, ByVal all As Boolean)
    Dim Redval&, Greenval&, Blueval&
    Dim r1&, g1&, b1&, sr&, SG&, sb&
    F& = F& Mod &H1000000
    t& = t& Mod &H1000000
    Redval& = F& And &H10000FF
    Greenval& = (F& And &H100FF00) / &H100
    Blueval& = (F& And &HFF0000) / &H10000
    r1& = t& And &H10000FF
    g1& = (t& And &H100FF00) / &H100
    b1& = (t& And &HFF0000) / &H10000
    sr& = (r1& - Redval&) * 1000 / 127
    SG& = (g1& - Greenval&) * 1000 / 127
    sb& = (b1& - Blueval&) * 1000 / 127
    Redval& = Redval& * 1000
    
    Greenval& = Greenval& * 1000
    Blueval& = Blueval& * 1000
    Dim Step&, Reps&, FillTop As Single, FillLeft As Single, FillRight As Single, FillBottom As Single
    If hor Then
    yy2& = TheObject.Height - yy2&
    If all Then
    Step = ((yy2& - yy1&) / 127)
    Else
    Step = (TheObject.Height / 127)
    End If
    If all Then
    FillTop = yy1&
    Else
    FillTop = 0
    End If
    FillLeft = xx1&
    FillRight = TheObject.Width - xx2&
    FillBottom = FillTop + Step * 2
    Else ' vertical
    
        xx2& = TheObject.Width - xx2&
    If all Then
    Step = ((xx2& - xx1&) / 127)
    Else
    Step = (TheObject.Width / 127)
    End If
    If all Then
    FillLeft = xx1&
    Else
    FillLeft = 0
    End If
    FillTop = yy1&
    FillBottom = TheObject.Height - yy2&
    FillRight = FillLeft + Step * 2
    
    End If
    For Reps = 1 To 127
    If hor Then
        If FillTop <= yy2& And FillBottom >= yy1& Then
        TheObject.Line (FillLeft, RMAX(FillTop, yy1&))-(FillRight, RMIN(FillBottom, yy2&)), rgb(Redval& / 1000, Greenval& / 1000, Blueval& / 1000), BF
        End If
        Redval& = Redval& + sr&
        Greenval& = Greenval& + SG&
        Blueval& = Blueval& + sb&
        FillTop = FillBottom
        FillBottom = FillTop + Step
    Else
        If FillLeft <= xx2& And FillRight >= xx1& Then
        TheObject.Line (RMAX(FillLeft, xx1&), FillTop)-(RMIN(FillRight, xx2&), FillBottom), rgb(Redval& / 1000, Greenval& / 1000, Blueval& / 1000), BF
        End If
        Redval& = Redval& + sr&
        Greenval& = Greenval& + SG&
        Blueval& = Blueval& + sb&
        FillLeft = FillRight
        FillRight = FillRight + Step
    End If
    Next
    
End Sub
Function mycolor(q)
If Abs(q) > 2147483392# Then
If q < 0 Then
mycolor = GetSysColor(q And &HFF) And &HFFFFFF
Else
mycolor = GetSysColor((q - 4294967296#) And &HFF) And &HFFFFFF
End If
Exit Function
End If
If q = 0 Then
mycolor = 0
ElseIf q < 0 Or q > 15 Then

 mycolor = Abs(q) And &HFFFFFF
Else
mycolor = QBColor(q Mod 16)
End If
End Function




Sub ICOPY(d1 As Object, x1 As Long, y1 As Long, w As Long, h As Long)
Dim sV As Long
With players(GetCode(d1))
sV = BitBlt(d1.hDC, CLng(d1.ScaleX(x1, 1, 3)), CLng(d1.ScaleY(y1, 1, 3)), CLng(d1.ScaleX(w, 1, 3)), CLng(d1.ScaleY(h, 1, 3)), d1.hDC, CLng(d1.ScaleX(.XGRAPH, 1, 3)), CLng(d1.ScaleY(.YGRAPH, 1, 3)), SRCCOPY)
'sv = UpdateWindow(d1.hwnd)
End With
End Sub

Sub sHelp(Title$, doc$, X As Long, Y As Long)
vH_title$ = Title$
vH_doc$ = doc$
vH_x = X
vH_y = Y
End Sub

Sub vHelp(Optional ByVal bypassshow As Boolean = False)
Dim huedif As Long
Dim UAddPixelsTop As Long, monitor As Long

If abt Then
If vH_title$ = lastAboutHTitle Then Exit Sub
vH_title$ = lastAboutHTitle
vH_doc$ = LastAboutText
Else
If vH_title$ = vbNullString Then Exit Sub
End If
If bypassshow Then
monitor = FindMonitorFromMouse
Else
monitor = FindFormSScreen(Form4)
End If
If Not Form4.Visible Then Form4.Show , Form1: bypassshow = True

If bypassshow Then
myform Form4, ScrInfo(monitor).Width - vH_x * Helplastfactor + ScrInfo(monitor).Left, ScrInfo(monitor).Height - vH_y * Helplastfactor + ScrInfo(monitor).top, vH_x * Helplastfactor, vH_y * Helplastfactor, True, Helplastfactor
Else
If Screen.Width <= Form4.Left - ScrInfo(monitor).Left Then
myform Form4, Screen.Width - vH_x * Helplastfactor + ScrInfo(monitor).Left, Form4.top, vH_x * Helplastfactor, vH_y * Helplastfactor, True, Helplastfactor
Else
myform Form4, Form4.Left, Form4.top, vH_x * Helplastfactor, vH_y * Helplastfactor, True, Helplastfactor
End If
End If
Form4.moveMe

If Form1.Visible Then
If Form1.DIS.Visible Then
  ''  If Abs(Val(hueconvSpecial(mycolor(uintnew(&H80000018)))) - Val(hueconvSpecial(-Paper))) > Abs(Val(hueconvSpecial(mycolor(uintnew(&H80000003)))) - Val(hueconvSpecial(-Paper))) Then
  If Abs(hueconv(mycolor(uintnew(&H80000018))) - val(hueconv(players(0).Paper))) > 10 And Not Abs(lightconv(mycolor(uintnew(&H80000018))) - val(lightconv(players(0).Paper))) < 50 Then
    Form4.backcolor = &H80000018
    Form4.label1.backcolor = &H80000018
    
    Else
    
    Form4.backcolor = &H80000003
    Form4.label1.backcolor = &H80000003
    End If

Else
''If Abs(Val(hueconvSpecial(mycolor(&H80000018))) - Val(hueconvSpecial(Form1.BackColor))) > Abs(Val(hueconvSpecial(mycolor(&H80000003))) - Val(hueconvSpecial(Form1.BackColor))) Then
     If Abs(hueconv(mycolor(uintnew(&H80000018))) - val(hueconv(Form1.backcolor))) > 10 And Not Abs(lightconv(mycolor(uintnew(&H80000018))) - val(lightconv(Form1.backcolor))) < 50 Then

    Form4.backcolor = &H80000018
    Form4.label1.backcolor = &H80000018
    Else
    
    Form4.backcolor = &H80000003
    Form4.label1.backcolor = &H80000003
    End If
End If
End If
With Form4.label1
.Visible = True
.enabled = False
.Text = vH_doc$
.SetRowColumn 1, 0
.EditDoc = False
.NoMark = True
If abt Then
.glistN.WordCharLeft = "["
.glistN.WordCharRight = "]"
.glistN.WordCharRightButIncluded = vbNullString
.glistN.WordCharLeftButIncluded = vbNullString
Else
.glistN.WordCharRightButIncluded = ChrW(160) + "("
.glistN.WordCharLeft = ConCat(":", "{", "}", "[", "]", ",", "(", ")", "!", ";", "=", ">", "<", "'", """", " ", "+", "-", "/", "*", "^", "@", Chr$(9), "#", "%", "&", "$")
.glistN.WordCharRight = ConCat(":", "{", "}", "[", "]", ",", ")", "!", ";", "=", ">", "<", "'", """", " ", "+", "-", "/", "*", "^", Chr$(9), "#")
.glistN.WordCharLeftButIncluded = "#$@~"

End If
.enabled = True
.NewTitle vH_title$, (4 + UAddPixelsTop) * Helplastfactor
.glistN.ShowMe
End With


'Form4.ZOrder
Form4.label1.glistN.DragEnabled = Not abt
If exWnd = 0 Then If Form1.Visible Then Form1.SetFocus
End Sub

Function FileNameType(extension As String) As String
Dim i As Long, fs, b
 strTemp = String(200, Chr$(0))
    'Get
    GetTempPath 200, StrPtr(strTemp)
    strTemp = LONGNAME(mylcasefILE(Left$(strTemp, InStr(strTemp, Chr(0)) - 1)))
    If strTemp = vbNullString Then
     strTemp = mylcasefILE(Left$(strTemp, InStr(strTemp, Chr(0)) - 1))
     If Right$(strTemp, 1) <> "\" Then strTemp = strTemp & "\"
    End If
    
    i = FreeFile
    Open strTemp & "dummy." & extension For Output As i
    Print #i, "test"
    Close #i
    Sleep 10
    Set fs = CreateObject("Scripting.FileSystemObject")
  Set b = fs.GetFile(strTemp & "dummy." & extension)
    FileNameType = b.Type
    KillFile strTemp & "dummy." & extension
End Function
Function mylcasefILE(ByVal a$) As String
If a$ = vbNullString Then Exit Function
If casesensitive Then
' no case change
mylcasefILE = a$
Else
 mylcasefILE = LCase(a$)
 End If

End Function

Function myUcase(ByVal a$, Optional convert As Boolean = False) As String
Dim i As Long
If a$ = vbNullString Then Exit Function
 If AscW(a$) > 255 Or convert Then
 For i = 1 To Len(a$)
 Select Case AscW(Mid$(a$, i, 1))
Case 902
Mid$(a$, i, 1) = ChrW(913)
Case 904
Mid$(a$, i, 1) = ChrW(917)
Case 906
Mid$(a$, i, 1) = ChrW(921)
Case 912
Mid$(a$, i, 1) = ChrW(921)
Case 905
Mid$(a$, i, 1) = ChrW(919)
Case 908
Mid$(a$, i, 1) = ChrW(927)
Case 911
Mid$(a$, i, 1) = ChrW(937)
Case 910
Mid$(a$, i, 1) = ChrW(933)
Case 940
Mid$(a$, i, 1) = ChrW(913)
Case 941
Mid$(a$, i, 1) = ChrW(917)
Case 943
Mid$(a$, i, 1) = ChrW(921)
Case 942
Mid$(a$, i, 1) = ChrW(919)
Case 972
Mid$(a$, i, 1) = ChrW(927)
Case 974
Mid$(a$, i, 1) = ChrW(937)
Case 973
Mid$(a$, i, 1) = ChrW(933)
Case 962
Mid$(a$, i, 1) = ChrW(931)
End Select
Next i
End If
myUcase = UCase(a$)
End Function

Function myLcase(ByVal a$) As String
If a$ = vbNullString Then Exit Function
a$ = LCase(a$)
If a$ = vbNullString Then Exit Function
 If AscW(a$) > 255 Then
a$ = a$ & Chr(0)
' Here are greek letters for proper case conversion
a$ = Replace(a$, "σ" & Chr(0), "ς")
a$ = Replace(a$, Chr(0), "")
a$ = Replace(a$, "σ ", "ς ")
a$ = Replace(a$, "σ$", "ς$")
a$ = Replace(a$, "σ&", "ς&")
a$ = Replace(a$, "σ.", "ς.")
a$ = Replace(a$, "σ(", "ς(")
a$ = Replace(a$, "σ_", "ς_")
a$ = Replace(a$, "σ/", "ς/")
a$ = Replace(a$, "σ\", "ς\")
a$ = Replace(a$, "σ-", "ς-")
a$ = Replace(a$, "σ+", "ς+")
a$ = Replace(a$, "σ*", "ς*")
a$ = Replace(a$, "σ" & vbCr, "ς" & vbCr)
a$ = Replace(a$, "σ" & vbLf, "ς" & vbLf)
End If

myLcase = a$
End Function
Function MesTitle$()
On Error Resume Next
If ttl Then
If Form1.Caption = vbNullString Then
If here$ = vbNullString Then
MesTitle$ = "M2000"
' IDE
Else
If LASTPROG$ <> "" Then
MesTitle$ = ExtractNameOnly(LASTPROG$)
Else
MesTitle$ = "M2000"
End If
End If
Else
MesTitle$ = Form1.Caption
End If
Else
If Typename$(Screen.ActiveForm) = "GuiM2000" Then
MesTitle$ = Screen.ActiveForm.Title
Else
If here$ = vbNullString Or LASTPROG$ = vbNullString Then
MesTitle$ = "M2000"
Else
MesTitle$ = ExtractNameOnly(LASTPROG$) & " " & here$
End If
End If
End If
End Function
Public Function holdcontrol(wh As Object, mb As basket) As Long
Dim x1 As Long, y1 As Long
With mb
If .pageframe = 0 Then
''GetXYb wh, mb, X1, y1
If .mysplit > 0 Then .pageframe = (.My - .mysplit) * 4 / 5 Else .pageframe = Fix(.My * 4 / 5)
If .pageframe < 1 Then .pageframe = 1
.basicpageframe = .pageframe
holdcontrol = .pageframe
Else
holdcontrol = .basicpageframe
End If
End With
End Function
Public Sub HoldReset(col As Long, mb As basket)
With mb
.basicpageframe = col
If .basicpageframe <= 0 Then .basicpageframe = .pageframe
End With
End Sub
Public Sub gsb_file(Optional assoc As Boolean = True)
   Dim cd As String
     cd = App.path
        AddDirSep cd

        If assoc Then
          associate ".gsb", "M2000 Ver" & Str$(VerMajor) & "." & CStr(VerMinor \ 100) & " User Module", cd & "M2000.EXE"
        Else
      deassociate ".gsb", "M2000 Ver" & Str$(VerMajor) & "." & CStr(VerMinor \ 100) & " User Module", cd & "M2000.EXE"
   End If
End Sub
Public Sub Switches(s$, Optional fornow As Boolean = False)
Dim cc As cRegistry
Set cc = New cRegistry
cc.temp = fornow
cc.ClassKey = HKEY_CURRENT_USER
    cc.SectionKey = basickey
Dim d$, w$, p As Long, b As Long
If s$ <> "" Then
    Do While FastSymbol(s$, "-")
            If IsLabel(basestack1, s$, d$) > 0 Then
            d$ = UCase(d$)
            If d$ = "TEST" Then
                STq = False
                STEXIT = False
                STbyST = True
                Form2.Show , Form1
                Form2.label1(0) = vbNullString
                Form2.label1(1) = vbNullString
                Form2.label1(2) = vbNullString
                 TestShowSub = vbNullString
 TestShowStart = 0
   
                stackshow basestack1
                Form1.Show , Form5
                If Form3.Visible Then Form3.skiptimer = True: Form3.WindowState = 0
                trace = True
            ElseIf d$ = "NORUN" Then
                If ttl Then Form3.WindowState = vbNormal Else Form1.Show , Form5
                NORUN1 = True
            ElseIf d$ = "FONT" Then
            ' + LOAD NEW
                cc.ValueKey = "FONT"
                    cc.ValueType = REG_SZ
               cc.Value = "Monospac821Greek BT"
            ElseIf d$ = "SEC" Then
                    cc.ValueKey = "NEWSECURENAMES"
                cc.ValueType = REG_DWORD
                cc.Value = 0
                SecureNames = False
            ElseIf d$ = "DIV" Then
                cc.ValueKey = "DIV"
                    cc.ValueType = REG_DWORD
                  cc.Value = 0
                  UseIntDiv = False
            ElseIf d$ = "LINESPACE" Then
                cc.ValueKey = "LINESPACE"
                    cc.ValueType = REG_DWORD
               
                  cc.Value = 0
            ElseIf d$ = "SIZE" Then
                cc.ValueKey = "SIZE"
                    cc.ValueType = REG_DWORD
               
                  cc.Value = 15
                 
                 
            ElseIf d$ = "PEN" Then
                cc.ValueKey = "PEN"
                    cc.ValueType = REG_DWORD
                  cc.Value = 0
                      cc.ValueKey = "PAPER"
                    cc.ValueType = REG_DWORD
                  cc.Value = 7
                  
            ElseIf d$ = "BOLD" Then
             cc.ValueKey = "BOLD"
                   cc.ValueType = REG_DWORD
                 
                  cc.Value = 0
                 
            
            ElseIf d$ = "PAPER" Then
                cc.ValueKey = "PAPER"
                    cc.ValueType = REG_DWORD
                  cc.Value = 7
                   cc.ValueKey = "PEN"
                    cc.ValueType = REG_DWORD
                  cc.Value = 0
                   
            ElseIf d$ = "GREEK" Then
            cc.ValueKey = "COMMAND"
                 cc.ValueType = REG_SZ
                    cc.Value = "LATIN"
                    If fornow Then pagio$ = "LATIN"
            ElseIf d$ = "DARK" Then
            cc.ValueKey = "HTML"
                 cc.ValueType = REG_SZ
                    cc.Value = "BRIGHT"
            ElseIf d$ = "CASESENSITIVE" Then
            cc.ValueKey = "CASESENSITIVE"
             cc.ValueType = REG_SZ
                    cc.Value = "NO"
            If fornow Then
                casesensitive = False
            End If
            ElseIf d$ = "RDB" Then
            RoundDouble = False
            ElseIf d$ = "EXT" Then
            wide = False
            ElseIf d$ = "TAB" Then
            UseTabInForm1Text1 = False
            ElseIf d$ = "SBL" Then
            ShowBooleanAsString = False
            ElseIf d$ = "DIM" Then
            DimLikeBasic = False
            ElseIf d$ = "FOR" Then
           ' cc.ValueKey = "FOR-LIKE-BASIC"
           ' cc.ValueType = REG_DWORD
            'cc.Value = CLng(0)
            ForLikeBasic = False
            ElseIf d$ = "PRI" Then
            cc.ValueKey = "PRIORITY-OR"
            cc.ValueType = REG_DWORD
            cc.Value = CLng(0)  ' FALSE IS WRONG VALUE HERE
            priorityOr = False
            ElseIf d$ = "REG" Then
            gsb_file False
            ElseIf d$ = "DEC" Then
            cc.ValueKey = "DEC"
             cc.ValueType = REG_DWORD
                    cc.Value = CLng(0)
                    mNoUseDec = False
                    CheckDec
            ElseIf d$ = "TXT" Then
            cc.ValueKey = "TEXTCOMPARE"
             cc.ValueType = REG_DWORD
                    cc.Value = CLng(0)
                    mTextCompare = False
                    
            ElseIf d$ = "REC" Then
               cc.ValueKey = "FUNCDEEP"  ' RESET
             cc.ValueType = REG_DWORD
                    cc.Value = 300
                    If m_bInIDE Then funcdeep = 128
                    ' funcdeep not used - but functionality stay there for old dll's
                ClaimStack
                If findstack - 100000 > 0 Then
                    stacksize = findstack - 100000
                End If
            Else
            s$ = "-" & d$ & s$
            Exit Do
            End If
            Else
        Exit Do
        End If
        Sleep 2
    Loop
Do While FastSymbol(s$, "+")
If IsLabel(basestack1, s$, d$) > 0 Then
            d$ = UCase(d$)
    If d$ = "TEST" Then
            STq = False
            STEXIT = False
            STbyST = True
            Form2.Show , Form1
            Form2.label1(0) = vbNullString
            Form2.label1(1) = vbNullString
            Form2.label1(2) = vbNullString
             TestShowSub = vbNullString
 TestShowStart = 0

            stackshow basestack1
            
            Form1.Show , Form5
            If Form3.Visible Then Form3.skiptimer = True: Form3.WindowState = 0
            trace = True
        ElseIf d$ = "REG" Then
        gsb_file
        ElseIf d$ = "FONT" Then
    ' + LOAD NEW
        cc.ValueKey = "FONT"
            cc.ValueType = REG_SZ
            If ISSTRINGA(s$, w$) Then cc.Value = w$
            ElseIf d$ = "SEC" Then
                    cc.ValueKey = "NEWSECURENAMES"
                cc.ValueType = REG_DWORD
                cc.Value = -1
                SecureNames = True
            ElseIf d$ = "DIV" Then
                cc.ValueKey = "DIV"
                    cc.ValueType = REG_DWORD
                  cc.Value = -1
                  UseIntDiv = True
        ElseIf d$ = "LINESPACE" Then
            cc.ValueKey = "LINESPACE"
                cc.ValueType = REG_DWORD
            If IsNumberLabel(s$, w$) Then If val(w$) >= 0 And val(w$) <= 60 * dv15 Then cc.Value = CLng(val(w$) * 2)
               
        ElseIf d$ = "SIZE" Then
            cc.ValueKey = "SIZE"
            cc.ValueType = REG_DWORD
            If IsNumberLabel(s$, w$) Then If val(w$) >= 8 And val(w$) <= 48 Then cc.Value = CLng(val(w$))
          
        ElseIf d$ = "PEN" Then
            cc.ValueKey = "PAPER"
            cc.ValueType = REG_DWORD
            p = cc.Value
            cc.ValueKey = "PEN"
            cc.ValueType = REG_DWORD
            If IsNumberLabel(s$, w$) Then
                If p = val(w$) Then p = 16 - p Else p = val(w$) Mod 16
                cc.Value = CLng(val(p))
            End If
        ElseIf d$ = "BOLD" Then
                cc.ValueKey = "BOLD"
                cc.ValueType = REG_DWORD
                If IsNumberLabel(s$, w$) Then cc.Value = CLng(val(w$) Mod 16)
                
        ElseIf d$ = "PAPER" Then
                cc.ValueKey = "PEN"
                cc.ValueType = REG_DWORD
                p = cc.Value
                cc.ValueKey = "PAPER"
                cc.ValueType = REG_DWORD
                If IsNumberLabel(s$, w$) Then
                If p = val(w$) Then p = 16 - p Else p = val(w$) Mod 16
                    cc.Value = CLng(val(p))
                End If
        ElseIf d$ = "GREEK" Then
                cc.ValueKey = "COMMAND"
                cc.ValueType = REG_SZ
                cc.Value = "GREEK"
                If fornow Then pagio$ = "GREEK"
        ElseIf d$ = "DARK" Then
            cc.ValueKey = "HTML"
                 cc.ValueType = REG_SZ
                    cc.Value = "DARK"
        ElseIf d$ = "CASESENSITIVE" Then
                cc.ValueKey = "CASESENSITIVE"
                cc.ValueType = REG_SZ
                cc.Value = "YES"
                If fornow Then
                     casesensitive = True
                End If
          ElseIf d$ = "RDB" Then
            RoundDouble = True
            ElseIf d$ = "EXT" Then
            wide = True
           ElseIf d$ = "TAB" Then
            UseTabInForm1Text1 = True
           ElseIf d$ = "SBL" Then
            ShowBooleanAsString = True
         ElseIf d$ = "DIM" Then
            DimLikeBasic = True
         ElseIf d$ = "FOR" Then
          '  cc.ValueKey = "FOR-LIKE-BASIC"
           ' cc.ValueType = REG_DWORD
           ' cc.Value = CLng(True)
             ForLikeBasic = True
        ElseIf d$ = "PRI" Then
        cc.ValueKey = "PRIORITY-OR"
                cc.ValueType = REG_DWORD
                cc.Value = CLng(True)
            priorityOr = True
            ElseIf d$ = "TXT" Then
            cc.ValueKey = "TEXTCOMPARE"
             cc.ValueType = REG_DWORD
                    cc.Value = CLng(True)
                    mTextCompare = True
        ElseIf d$ = "DEC" Then
            cc.ValueKey = "DEC"
             cc.ValueType = REG_DWORD
                    cc.Value = CLng(True)
                    mNoUseDec = True
                    CheckDec
        ElseIf d$ = "REC" Then
               cc.ValueKey = "FUNCDEEP"  ' RESET
             cc.ValueType = REG_DWORD
             funcdeep = 3260
                    cc.Value = 3260 ' SET REVISION DEFAULT
        ClaimStack
                If findstack - 100000 > 0 Then
                    stacksize = findstack - 100000
                End If
        Else
            s$ = "+" & d$ & s$
            Exit Do
        End If
    Else
    Exit Do
    End If
Sleep 2
Loop

End If
End Sub
Function blockStringPOS(s$, pos As Long) As Boolean
Dim i As Long, j As Long, c As Long
Dim a1 As Boolean
c = Len(s$)
a1 = True
i = pos
If i > Len(s$) Then Exit Function
Do
Select Case AscW(Mid$(s$, i, 1))
Case 34
Do While i < c
i = i + 1
If AscW(Mid$(s$, i, 1)) = 34 Then Exit Do
Loop
Case 123
j = j - 1
Case 125
j = j + 1: If j = 1 Then Exit Do
End Select
i = i + 1
Loop Until i > c
If j = 1 Then
blockStringPOS = True
pos = i
Else
pos = Len(s$)
End If

End Function
Function BlockParam2(s$, pos As Long) As Boolean
' need to be open
Dim i As Long, j As Long, ii As Long
j = 1
For i = pos To Len(s$)
Select Case AscW(Mid$(s$, i, 1))
Case 0
Exit For
Case 34
again:
ii = InStr(i + 1, s$, """")
If ii = 0 Then Exit Function
 i = ii
If Mid$(s$, ii - 1, 1) = "`" Then GoTo again

Case 40
j = j + 1
Case 41
j = j - 1
If j = 0 Then Exit For
Case 123
i = i + 1
If blockStringPOS(s$, i) Then
Else
i = 0
End If
If i = 0 Then Exit Function
End Select
Next i
If j = 0 Then pos = i: BlockParam2 = True
End Function
Function BlockParam3(s$, pos As Long) As Boolean
' need to be open
Dim i As Long, j As Long, ii As Long
j = 1
For i = pos To Len(s$)
Select Case AscW(Mid$(s$, i, 1))
Case 0
Exit For
Case 34
again:
ii = InStr(i + 1, s$, """")
If ii = 0 Then Exit Function
 i = ii
If Mid$(s$, ii - 1, 1) = "`" Then GoTo again

Case 40
j = j + 1
Case 41
j = j - 1
If j = 0 Then Exit For
Case 123
i = i + 1
If blockStringPOS(s$, i) Then
Else
i = 0
End If
If i = 0 Then Exit Function
End Select
Next i
If j = 0 Then pos = i: BlockParam3 = True
End Function
Public Sub GetCodePart(a$, pos As Long)
Dim pos2 As Long, w$
If a$ = vbNullString Then Exit Sub
Dim v1 As Long
If pos = 0 Then pos = 1
Do While pos <= Len(a$)
    w$ = Mid$(a$, pos, 1)
    v1 = AscW(w$)
    If Abs(v1) > 9 Then
    If w$ = """" Then
        pos = pos + 1
        Do While pos <= Len(a$)
        If Mid$(a$, pos, 1) = """" Then Exit Do
        If AscW(Mid$(a$, pos, 1)) < 32 Then Exit Do
        pos = pos + 1
        Loop
    ElseIf w$ = "(" Then
again22:
      pos = pos + 1
        If Not BlockParam2(a$, pos) Then Exit Do
        If Mid$(a$, pos + 1, 1) = "#" Then
        pos = pos + 1
        GoTo conthere
        ElseIf Mid$(a$, pos + 1, 1) = "(" Then
        pos = pos + 1: GoTo again22
        End If
    ElseIf w$ = "{" Then
       If pos <= Len(a$) Then
        'If Not blockStringAhead(a$, pos) Then Exit Do
        pos = blockLen2(a$, pos + 1)
        If pos = 0 Then pos = Len(a$): Exit Do
        End If
    Else
        Select Case w$
        Case ")", "}", Is < " ", "'", "\", vbLf
        Exit Do
        End Select
        End If
End If
        pos = pos + 1
        
  
conthere:
  
Loop

End Sub
Function blockLen2(s$, pos) As Long
Dim i As Long, j As Long, c As Long
Dim a1 As Boolean
Dim jump As Boolean
If Trim$(s$) = vbNullString Then Exit Function
c = Len(s$)
a1 = True
i = pos
Do
Select Case AscW(Mid$(s$, i, 1))
Case 32
' nothing
Case 34
Do While i < c
i = i + 1
If AscW(Mid$(s$, i, 1)) = 34 Then Exit Do
Loop
Case 39, 92
Do While i < c
i = i + 1
If Mid$(s$, i, 2) = vbCrLf Then Exit Do
Loop
Case 61
jump = True
Case 123
If jump Then
jump = False
Dim target As Long
target = j
    Do
    Select Case AscW(Mid$(s$, i, 1))
    Case 34
    Do While i < c
    i = i + 1
    If AscW(Mid$(s$, i, 1)) = 34 Then Exit Do
    Loop
    Case 123
    j = j - 1
    Case 125
    j = j + 1: If j = target Then Exit Do
    End Select
    i = i + 1
    Loop Until i > c
    If j <> target Then Exit Do
    Else
j = j - 1
End If


Case 125
j = j + 1: If j = 1 Then Exit Do
Case Else
jump = False

End Select
i = i + 1
Loop Until i > c
If j = 1 Then
blockLen2 = i
Else
blockLen2 = 0
End If



End Function
Public Sub aheadstatusSkipParam(a$, pos As Long)
Dim pos2 As Long, w$
If a$ = vbNullString Then Exit Sub
Dim v1 As Long
If pos = 0 Then pos = 1
Do While pos <= Len(a$)
    w$ = Mid$(a$, pos, 1)
    v1 = AscW(w$)
    If Abs(v1) > 9 Then
    If w$ = """" Then
        pos = pos + 1
        Do While pos <= Len(a$)
        If Mid$(a$, pos, 1) = """" Then Exit Do
        If AscW(Mid$(a$, pos, 1)) < 32 Then Exit Do
        pos = pos + 1
        Loop
    ElseIf w$ = "(" Then
again22:
      pos = pos + 1
        If Not BlockParam2(a$, pos) Then Exit Do
        If Mid$(a$, pos + 1, 1) = "#" Then
        pos = pos + 1
        GoTo conthere
        ElseIf Mid$(a$, pos + 1, 1) = "(" Then
        pos = pos + 1: GoTo again22
        End If
    ElseIf w$ = "{" Then
       If pos <= Len(a$) Then
        'If Not blockStringAhead(a$, pos) Then Exit Do
        pos = blockLen2(a$, pos + 1)
        If pos = 0 Then pos = Len(a$): Exit Do
        End If
    Else
        Select Case w$
        Case ":", ")", "}", Is < " ", "'", "\", vbLf
        Exit Do
        End Select
        End If
End If
        pos = pos + 1
        
  
conthere:
  
Loop

End Sub
Public Sub aheadstatusSkipParam2(a$, pos As Long)
' no block  ' for for next
Dim pos2 As Long, w$
If a$ = vbNullString Then Exit Sub
Dim v1 As Long
If pos = 0 Then pos = 1
Do While pos <= Len(a$)
    w$ = Mid$(a$, pos, 1)
    v1 = AscW(w$)
    If Abs(v1) > 9 Then
    If w$ = """" Then
        pos = pos + 1
        Do While pos <= Len(a$)
        If Mid$(a$, pos, 1) = """" Then Exit Do
        If AscW(Mid$(a$, pos, 1)) < 32 Then Exit Do
        pos = pos + 1
        Loop
    ElseIf w$ = "(" Then
again22:
      pos = pos + 1
        If Not BlockParam2(a$, pos) Then Exit Do
        If Mid$(a$, pos + 1, 1) = "#" Then
        pos = pos + 1
        GoTo conthere
        ElseIf Mid$(a$, pos + 1, 1) = "(" Then
        pos = pos + 1: GoTo again22
        End If
    Else
        Select Case w$
        Case ":", ")", "{", "}", Is < " ", "'", "\", vbLf
        Exit Do
        End Select
        End If
End If
        pos = pos + 1
        
  
conthere:
  
Loop

End Sub

Public Sub aheadstatusNext(a$, pos As Long, Lang As Long, flag As Boolean)
Dim pos2 As Long, what$, w$, lenA As Long, level2 As Integer
Const second1$ = "ΓΙΑ", len1 = 3
Const second2$ = "FOR", len2 = 3
flag = False
If a$ = vbNullString Then Exit Sub
Dim v1 As Long
If pos = 0 Then pos = 1
lenA = Len(a$)
Do While pos <= lenA
    w$ = Mid$(a$, pos, 1)
    v1 = AscW(w$)
    If Abs(v1) > 9 Then
    
    If w$ = """" Then
            If Len(what$) > 0 Then what$ = vbNullString
        pos = pos + 1
        Do While pos <= lenA
        If Mid$(a$, pos, 1) = """" Then Exit Do
        If AscW(Mid$(a$, pos, 1)) < 32 Then Exit Do
        pos = pos + 1
        Loop
    ElseIf w$ = "(" Then
        If Len(what$) > 0 Then what$ = vbNullString
again22:
      pos = pos + 1

        If Not BlockParam2(a$, pos) Then Exit Do
        If Mid$(a$, pos + 1, 1) = "#" Then
        pos = pos + 1
        GoTo conthere
        ElseIf Mid$(a$, pos + 1, 1) = "(" Then
        pos = pos + 1: GoTo again22
        End If
    ElseIf w$ = "{" Then
       If Len(what$) > 0 Then what$ = vbNullString
       If pos <= lenA Then
        'If Not blockStringAhead(a$, pos) Then Exit Do
        pos = blockLen2(a$, pos + 1)
        If pos = 0 Then pos = lenA: Exit Do
        End If
    Else
        Select Case w$
        Case "%", "$", "0" To "9", vbLf
            If Len(what$) > 0 Then what$ = vbNullString
        Case " ", ChrW(160), vbCr, ":"
again:
            If Len(what$) > 2 Then
                If Lang = 0 Then
                    If Len(what$) = 7 Or Len(what$) = len1 Then
                        what$ = myUcase(what$)
                        If what$ = "ΕΠΟΜΕΝΟ" Then
                               If MyTrim$(w$) = "" Then aheadstatusSkipParam a$, pos
                                If level2 = 0 Then
                                    flag = True
                                    Exit Sub
                                Else
                                    level2 = level2 - 1
                                End If
                        ElseIf what$ = second1$ Then
                            aheadstatusSkipParam2 a$, pos
                            If MaybeIsSymbol3(a$, "{", pos) Then
                                pos = blockLen2(a$, pos + 1)
                                If pos = 0 Then pos = Len(a$): Exit Sub
                            ElseIf MaybeIsSymbol3lot(a$, b1234, pos) Then
                                level2 = level2 + 1
                            End If
                        Else
                            aheadstatusSkipParam a$, pos
                        End If
                    Else
                        aheadstatusSkipParam a$, pos
                    End If
                Else
                    If Len(what$) = 4 Or Len(what$) = len2 Then
                        what$ = myUcase(what$)
                        If what$ = "NEXT" Then
                            If MyTrim$(w$) = "" Then aheadstatusSkipParam a$, pos
                            If level2 = 0 Then
                                flag = True
                                Exit Sub
                            Else
                                level2 = level2 - 1
                            End If
                        ElseIf what$ = second2$ Then
                            aheadstatusSkipParam2 a$, pos
                            If MaybeIsSymbol3(a$, "{", pos) Then
                                pos = blockLen2(a$, pos + 1)
                                If pos = 0 Then pos = Len(a$): Exit Sub
                            ElseIf MaybeIsSymbol3lot(a$, b1234, pos) Then
                                level2 = level2 + 1
            
                              
                            End If
                        Else
                            aheadstatusSkipParam a$, pos
                        End If
                    Else
                        aheadstatusSkipParam a$, pos
                    End If
                End If
          Else
          If Len(what$) > 0 Then aheadstatusSkipParam a$, pos
          End If
                what$ = vbNullString
                
        pos = pos + 1
        Do
        pos2 = pos + 1
        Select Case Mid$(a$, pos, 1)
        Case " ", Chr$(160), vbTab
            pos = pos + 1
        End Select
        Loop Until pos2 > pos
        pos = pos - 1

        Case "'", "\"
            If Len(what$) > 0 Then what$ = vbNullString
        Do
        pos = pos + 1
        
        Loop While pos < lenA And Not Mid$(a$, pos, 1) = vbLf
        
        Case ")", "}", Is < " ", "'", "\"
        
        Exit Do
        Case ".", "A" To "Z", "a" To "z", Is >= "Α"
        
        what$ = what$ + w$
        
        Case Else
            If Len(what$) > 0 Then what$ = vbNullString
        End Select
        End If
End If
        pos = pos + 1
        
  
conthere:
  
Loop
If Len(what$) > 2 Then GoTo again
pos = lenA + 2

End Sub

Public Sub aheadstatusDO(a$, pos As Long, Lang As Long, flag As Boolean)
Dim pos2 As Long, what$, w$, lenA As Long, level2 As Integer
Const second1$ = "ΕΠΑΝΕΛΑΒΕ", len1 = 9
Const second11$ = "ΕΠΑΝΑΛΑΒΕ", len11 = 9
Const second2$ = "DO", len2 = 2
Const second22$ = "REPEAT", len22 = 6
flag = False
If a$ = vbNullString Then Exit Sub
Dim v1 As Long
If pos = 0 Then pos = 1
lenA = Len(a$)
Do While pos <= lenA
    w$ = Mid$(a$, pos, 1)
    v1 = AscW(w$)
    If Abs(v1) > 9 Then
    
    If w$ = """" Then
            If Len(what$) > 0 Then what$ = vbNullString
        pos = pos + 1
        Do While pos <= lenA
        If Mid$(a$, pos, 1) = """" Then Exit Do
        If AscW(Mid$(a$, pos, 1)) < 32 Then Exit Do
        pos = pos + 1
        Loop
    ElseIf w$ = "(" Then
        If Len(what$) > 0 Then what$ = vbNullString
again22:
      pos = pos + 1

        If Not BlockParam2(a$, pos) Then Exit Do
        If Mid$(a$, pos + 1, 1) = "#" Then
        pos = pos + 1
        GoTo conthere
        ElseIf Mid$(a$, pos + 1, 1) = "(" Then
        pos = pos + 1: GoTo again22
        End If
    ElseIf w$ = "{" Then
       If Len(what$) > 0 Then what$ = vbNullString
       If pos <= lenA Then
        'If Not blockStringAhead(a$, pos) Then Exit Do
                pos = blockLen2(a$, pos + 1)
        If pos = 0 Then pos = lenA: Exit Do
        End If
    Else
        Select Case w$
        Case "%", "$", "0" To "9", vbLf
            If Len(what$) > 0 Then what$ = vbNullString
        Case " ", ChrW(160), vbCr
            If Len(what$) > 1 Then
                If Lang = 0 Then
                    v1 = Len(what$)
                    If Len(what$) = 5 Or v1 = len1 Then
                        what$ = myUcase(what$)
                        If what$ = "ΜΕΧΡΙ" Or what$ = "ΠΑΝΤΑ" Then
                                
                                If level2 = 0 Then
                                    pos = pos - v1
                                    flag = True
                                    Exit Sub
                                Else
                                    If Left$(what$, 1) = "Μ" Then aheadstatusSkipParam a$, pos
                                    level2 = level2 - 1
                                End If
                        ElseIf what$ = second1$ Or what$ = second11$ Then
                            If MaybeIsSymbol3(a$, "{", pos) Then
                                aheadstatusSTRUCT a$, pos
                            'ElseIf MaybeIsSymbol3lot(a$, b1234, pos) Then
                            Else
                                level2 = level2 + 1
                            End If
                        Else
                            aheadstatusSkipParam a$, pos
                        End If
                    Else
                        aheadstatusSkipParam a$, pos
                    End If
                Else
                    v1 = Len(what$)
                    If v1 = 4 Or v1 = 5 Or v1 = len2 Or v1 = len22 Then
                        what$ = UCase(what$)
                        If what$ = "UNTIL" Or what$ = "ALWAYS" Then
                            If level2 = 0 Then
                                pos = pos - v1
                                flag = True
                                Exit Sub
                            Else
                                If Left$(what$, 1) = "U" Then aheadstatusSkipParam a$, pos
                                level2 = level2 - 1
                            End If
                        ElseIf what$ = second2$ Or what$ = second22$ Then
                            If MaybeIsSymbol3(a$, "{", pos) Then
                                aheadstatusSTRUCT a$, pos
                          '  ElseIf MaybeIsSymbol3lot(a$, b1234, pos) Then
                          Else
                                level2 = level2 + 1
                            End If
                        Else
                            aheadstatusSkipParam a$, pos
                        End If
                    Else
                        aheadstatusSkipParam a$, pos
                    End If
                End If
          Else
          If Len(what$) > 0 Then aheadstatusSkipParam a$, pos
          End If
                  pos = pos + 1
        Do
        pos2 = pos + 1
        Select Case Mid$(a$, pos, 1)
        Case " ", Chr$(160), vbCr, vbLf, vbTab
            pos = pos + 1
        End Select
        Loop Until pos2 > pos
        pos = pos - 1

                what$ = vbNullString
        Case "'", "\"
            If Len(what$) > 0 Then what$ = vbNullString
        Do
        pos = pos + 1
        
        Loop While pos < lenA And Not Mid$(a$, pos, 1) = vbLf
        
        Case ")", "}", Is < " ", "'", "\"
        
        Exit Do
        Case ".", "A" To "Z", "a" To "z", Is >= "Α"
        
        what$ = what$ + w$
        
        Case Else
            If Len(what$) > 0 Then what$ = vbNullString
        End Select
        End If
End If
        pos = pos + 1
        
  
conthere:
  
Loop
If flag = False Then
If Len(what$) > 0 Then
If level2 = 0 Then
    what$ = UCase(what$)
    If what$ = "ALWAYS" Then
        pos = pos - Len(what$)
        flag = True
        Exit Sub
    End If
End If
End If
End If
pos = lenA + 2

End Sub
Public Sub aheadstatusANY(a$, pos As Long)
Dim pos2 As Long, w$
If a$ = vbNullString Then Exit Sub
Dim v1 As Long
If pos = 0 Then pos = 1
Do While pos <= Len(a$)
    w$ = Mid$(a$, pos, 1)
    v1 = AscW(w$)
    If Abs(v1) > 9 Then
    If w$ = """" Then
        pos = pos + 1
        Do While pos <= Len(a$)
        If Mid$(a$, pos, 1) = """" Then Exit Do
        If AscW(Mid$(a$, pos, 1)) < 32 Then Exit Do
        pos = pos + 1
        Loop
    ElseIf w$ = "(" Then
again22:
      pos = pos + 1
        If Not BlockParam2(a$, pos) Then Exit Do
        If Mid$(a$, pos + 1, 1) = "#" Then
        pos = pos + 1
        GoTo conthere
        ElseIf Mid$(a$, pos + 1, 1) = "(" Then
        pos = pos + 1: GoTo again22
        End If
    ElseIf w$ = "{" Then
       If pos <= Len(a$) Then
        'If Not blockStringAhead(a$, pos) Then Exit Do
        pos = blockLen2(a$, pos + 1)
        If pos = 0 Then pos = Len(a$): Exit Do
        End If
    Else
        Select Case w$
        Case ")", "}", Is < " ", "'", "\"
        Exit Do
        End Select
        End If
End If
        pos = pos + 1
        
  
conthere:
  
Loop

End Sub
Public Sub aheadstatusSTRUCT(a$, pos As Long)
Dim pos2 As Long, w$, lenA As Long
If a$ = vbNullString Then Exit Sub
Dim v1 As Long
lenA = Len(a$)
If pos = 0 Then pos = 1
Do While pos <= lenA
    w$ = Mid$(a$, pos, 1)
    v1 = AscW(w$)
    If Abs(v1) > 9 Then
    If w$ = """" Then
        pos = pos + 1
        Do While pos <= lenA
        If Mid$(a$, pos, 1) = """" Then Exit Do
        If AscW(Mid$(a$, pos, 1)) < 32 Then Exit Do
        pos = pos + 1
        Loop
    ElseIf w$ = "(" Then
again22:
      pos = pos + 1
        If Not BlockParam2(a$, pos) Then Exit Do
        If Mid$(a$, pos + 1, 1) = "#" Then
        pos = pos + 1
        GoTo conthere
        ElseIf Mid$(a$, pos + 1, 1) = "(" Then
        pos = pos + 1: GoTo again22
        End If
    ElseIf w$ = "{" Then
       If pos <= lenA Then
        'If Not blockStringAhead(a$, pos) Then Exit Do
        pos = blockLen2(a$, pos + 1)
        If pos = 0 Then pos = lenA: Exit Do
        If MaybeIsSymbol3(a$, ":", pos) Then pos = pos + 1: Exit Do
        End If
    Else
        Select Case w$
        Case ")", "}", Is < " ", "'", "\"
        Exit Do
        End Select
        End If
End If
        pos = pos + 1
        
  
conthere:
  
Loop
' skip line
    Do
        pos = pos + 1
        
        Loop While pos < lenA And Not Mid$(a$, pos, 1) = vbLf

End Sub
Public Sub aheadstatusIFthen(a$, pos As Long, Lang As Long, w$)
Dim pos2 As Long, what$

If a$ = vbNullString Then Exit Sub
Dim v1 As Long
If pos = 0 Then pos = 1
Do While pos <= Len(a$)
    w$ = Mid$(a$, pos, 1)
    v1 = AscW(w$)
    If Abs(v1) > 9 Then
    
    If w$ = """" Then
            If Len(what$) > 0 Then what$ = vbNullString
        pos = pos + 1
        Do While pos <= Len(a$)
        If Mid$(a$, pos, 1) = """" Then Exit Do
        If AscW(Mid$(a$, pos, 1)) < 32 Then Exit Do
        pos = pos + 1
        Loop
    ElseIf w$ = "(" Then
        If Len(what$) > 0 Then what$ = vbNullString
again22:
      pos = pos + 1

        If Not BlockParam2(a$, pos) Then Exit Do
        If Mid$(a$, pos + 1, 1) = "#" Then
        pos = pos + 1
        GoTo conthere
        ElseIf Mid$(a$, pos + 1, 1) = "(" Then
        pos = pos + 1: GoTo again22
        End If
    Else
        Select Case w$
        Case ",", ":"
        Exit Do
        Case "%", "$", "0" To "9"
            If Len(what$) > 0 Then what$ = vbNullString
        Case " ", ChrW(160), vbCr, "{"
                If Len(what$) > 3 Then
                    If Len(what$) > 0 Then
                        what$ = myUcase(what$)
                        If Lang = 0 Then
                             If what$ = "ΤΟΤΕ" Or what$ = "ΑΛΛΙΩΣ" Then
                             pos2 = pos
                                    If MaybeIsSymbol3lot(a$, b123, pos2) Then
                                       
                                        w$ = what$
                                        Exit Sub
                                    Else
                                        aheadstatusSkipParam a$, pos
                                    End If
                                    Exit Do
                            End If
                        Else
                            If what$ = "THEN" Or what$ = "ELSE" Then
                                    pos2 = pos
                                    If MaybeIsSymbol3lot(a$, b123, pos2) Then
                                       
                                        w$ = what$
                                        Exit Sub
                                    Else
                                        aheadstatusSkipParam a$, pos
                                    End If
                                    Exit Do
                            End If
                        End If
                    End If
                    End If
                    If w$ = vbCr Then Exit Do
                    what$ = vbNullString
            If w$ = "{" Then
                     If pos <= Len(a$) Then
                    pos = blockLen2(a$, pos + 1)
                    If pos = 0 Then pos = Len(a$): Exit Do
                    End If
            Else
            pos = pos + 1
        Do
        pos2 = pos + 1
        Select Case Mid$(a$, pos, 1)
        Case " ", Chr$(160), vbTab
            pos = pos + 1
        End Select
        Loop Until pos2 > pos
        pos = pos - 1
        End If
         
        Case ")", "}", Is < " ", "'", "\"
        Exit Do
        Case Else
        
        what$ = what$ + w$
        
        
        End Select
        End If
End If
        pos = pos + 1
        
  
conthere:
  
Loop
w$ = vbNullString
If Len(what$) > 0 Then
      what$ = myUcase(what$)
      If Lang = 0 Then
         If what$ = "ΤΟΤΕ" Or what$ = "ΑΛΛΙΩΣ" Then
            w$ = what$
        ElseIf what$ = "THEN" Or what$ = "ELSE" Then
            w$ = what$
        End If
      End If
        
End If
End Sub
Public Sub aheadstatusIF(a$, pos As Long, Lang As Long, w$)
Dim pos2 As Long, what$

If a$ = vbNullString Then Exit Sub
Dim v1 As Long
If pos = 0 Then pos = 1
Do While pos <= Len(a$)
    w$ = Mid$(a$, pos, 1)
    v1 = AscW(w$)
    If Abs(v1) > 9 Then
    
    If w$ = """" Then
            If Len(what$) > 0 Then what$ = vbNullString
        pos = pos + 1
        Do While pos <= Len(a$)
        If Mid$(a$, pos, 1) = """" Then Exit Do
        If AscW(Mid$(a$, pos, 1)) < 32 Then Exit Do
        pos = pos + 1
        Loop
    ElseIf w$ = "(" Then
        If Len(what$) > 0 Then what$ = vbNullString
again22:
      pos = pos + 1

        If Not BlockParam2(a$, pos) Then Exit Do
        If Mid$(a$, pos + 1, 1) = "#" Then
        pos = pos + 1
        GoTo conthere
        ElseIf Mid$(a$, pos + 1, 1) = "(" Then
        pos = pos + 1: GoTo again22
        End If
    Else
        Select Case w$
        Case ",", ":"
        Exit Do
        Case "%", "$", "0" To "9"
            If Len(what$) > 0 Then what$ = vbNullString
        Case " ", ChrW(160), "{"
                If Len(what$) > 3 Then
                    If Len(what$) > 0 Then
                        what$ = myUcase(what$)
                        If Lang = 0 Then
                             If what$ = "ΤΟΤΕ" Or what$ = "ΑΛΛΙΩΣ" Then
                                w$ = what$
                                Exit Sub
                            End If
                        Else
                            If what$ = "THEN" Or what$ = "ELSE" Then
                                w$ = what$
                                Exit Sub
                            End If
                        End If
                    End If
                    End If
                    what$ = vbNullString
                    If w$ = "{" Then
                     If pos <= Len(a$) Then
                    pos = blockLen2(a$, pos + 1)
                    If pos = 0 Then pos = Len(a$): Exit Do
                    End If
                    Else
                 pos = pos + 1
        Do
        pos2 = pos + 1
        Select Case Mid$(a$, pos, 1)
        Case " ", Chr$(160), vbTab
            pos = pos + 1
        End Select
        Loop Until pos2 > pos
        pos = pos - 1
        End If
        Case ")", "}", Is < " ", "'", "\"
        Exit Do
        Case Else
        
        what$ = what$ + w$
        
        
        End Select
        End If
End If
        pos = pos + 1
        
  
conthere:
  
Loop
w$ = vbNullString
If Len(what$) > 0 Then
      what$ = myUcase(what$)
      If Lang = 0 Then
         If what$ = "ΤΟΤΕ" Or what$ = "ΑΛΛΙΩΣ" Then
            w$ = what$
        ElseIf what$ = "THEN" Or what$ = "ELSE" Then
            w$ = what$
        End If
      End If
        
End If
End Sub
Public Sub aheadstatusELSE(a$, pos As Long, Lang As Long, w$)
Dim pos2 As Long, what$

If a$ = vbNullString Then Exit Sub
Dim v1 As Long
If pos = 0 Then pos = 1
Do While pos <= Len(a$)
    w$ = Mid$(a$, pos, 1)
    v1 = AscW(w$)
    If Abs(v1) > 9 Then
    
    If w$ = """" Then
            If Len(what$) > 0 Then what$ = vbNullString
        pos = pos + 1
        Do While pos <= Len(a$)
        If Mid$(a$, pos, 1) = """" Then Exit Do
        If AscW(Mid$(a$, pos, 1)) < 32 Then Exit Do
        pos = pos + 1
        Loop
    ElseIf w$ = "(" Then
        If Len(what$) > 0 Then what$ = vbNullString
again22:
      pos = pos + 1

        If Not BlockParam2(a$, pos) Then Exit Do
        If Mid$(a$, pos + 1, 1) = "#" Then
        pos = pos + 1
        GoTo conthere
        ElseIf Mid$(a$, pos + 1, 1) = "(" Then
        pos = pos + 1: GoTo again22
        End If
    Else
        Select Case w$
        Case "%", "$", "0" To "9"
            If Len(what$) > 0 Then what$ = vbNullString
        Case " ", ChrW(160), "0" To "9", "{"
                If Len(what$) > 1 Then
                    If Len(what$) > 0 Then
                        what$ = myUcase(what$)
                        If Lang = 0 Then
                        If what$ = "ΑΝ" Or what$ = "ΑΛΛΙΩΣ" Or what$ = "ΑΛΛΙΩΣ.ΑΝ" Then 'Or what$ = "ΤΟΤΕ"
                            w$ = what$
                                Exit Sub
                            End If
                        Else
                            If what$ = "IF" Or what$ = "ELSE" Or what$ = "ELSE.IF" Then  'Or what$ = "THEN"
                                w$ = what$
                                Exit Sub
                            End If
                        End If
                    End If
                    End If
                    what$ = vbNullString
       
                 If w$ = "{" Then
                     If pos <= Len(a$) Then
                    pos = blockLen2(a$, pos + 1)
                    If pos = 0 Then pos = Len(a$): Exit Do
                    End If
                 Else
                pos = pos + 1
        Do
        pos2 = pos + 1
        Select Case Mid$(a$, pos, 1)
        Case " ", Chr$(160), vbTab
            pos = pos + 1
        End Select
        Loop Until pos2 > pos
        pos = pos - 1
        End If
        Case ")", "}", Is < " ", "'", "\"
        Exit Do
        Case Else
        If w$ = ":" Then
        what$ = ""
        Else
        what$ = what$ + w$
        End If
        
        End Select
        End If
End If
        pos = pos + 1
        
  
conthere:
  
Loop
w$ = vbNullString
If Len(what$) > 1 Then
      what$ = myUcase(what$)
      If Lang = 0 Then
                        If what$ = "ΑΝ" Or what$ = "ΤΟΤΕ" Or what$ = "ΑΛΛΙΩΣ" Or what$ = "ΑΛΛΙΩΣ.ΑΝ" Then
                            w$ = what$
                                Exit Sub
                            End If
                        Else
                            If what$ = "IF" Or what$ = "THEN" Or what$ = "ELSE" Or what$ = "ELSE.IF" Then
                                w$ = what$
                                Exit Sub
                            End If
                        End If
        
End If
End Sub
Public Sub aheadstatusThen(a$, pos As Long, Lang As Long, w$)
Dim pos2 As Long, what$

If a$ = vbNullString Then Exit Sub
Dim v1 As Long
If pos = 0 Then pos = 1
Do While pos <= Len(a$)
    w$ = Mid$(a$, pos, 1)
    v1 = AscW(w$)
    If Abs(v1) > 9 Then
    
    If w$ = """" Then
            If Len(what$) > 0 Then what$ = vbNullString
        pos = pos + 1
        Do While pos <= Len(a$)
        If Mid$(a$, pos, 1) = """" Then Exit Do
        If AscW(Mid$(a$, pos, 1)) < 32 Then Exit Do
        pos = pos + 1
        Loop
    ElseIf w$ = "(" Then
        If Len(what$) > 0 Then what$ = vbNullString
again22:
      pos = pos + 1

        If Not BlockParam2(a$, pos) Then Exit Do
        If Mid$(a$, pos + 1, 1) = "#" Then
        pos = pos + 1
        GoTo conthere
        ElseIf Mid$(a$, pos + 1, 1) = "(" Then
        pos = pos + 1: GoTo again22
        End If
    ElseIf w$ = "{" Then
       If Len(what$) > 0 Then what$ = vbNullString
       If pos <= Len(a$) Then
        'If Not blockStringAhead(a$, pos) Then Exit Do
        pos = blockLen2(a$, pos + 1)
        If pos = 0 Then pos = Len(a$): Exit Do
        End If
    Else
        Select Case w$
        Case "%", "$", "0" To "9"
            If Len(what$) > 0 Then what$ = vbNullString
        Case " ", ChrW(160), "0" To "9"
                If Len(what$) > 1 Then
                    If Len(what$) > 0 Then
                        what$ = myUcase(what$)
                        If Lang = 0 Then
                        If what$ = "ΤΟΤΕ" Or what$ = "ΑΛΛΙΩΣ" Then
                            w$ = what$
                                Exit Sub
                            End If
                        Else
                            If what$ = "THEN" Or what$ = "ELSE" Then
                                w$ = what$
                                Exit Sub
                            End If
                        End If
                    End If
                    End If
                    what$ = vbNullString
                 pos = pos + 1
        Do
        pos2 = pos + 1
        Select Case Mid$(a$, pos, 1)
        Case " ", Chr$(160), vbTab
            pos = pos + 1
        End Select
        Loop Until pos2 > pos
        pos = pos - 1

        Case ")", "}", Is < " ", "'", "\"
        Exit Do
        Case Else
        
        what$ = what$ + w$
        
        
        End Select
        End If
End If
        pos = pos + 1
        
  
conthere:
  
Loop
w$ = vbNullString
If Len(what$) > 1 Then
      what$ = myUcase(what$)
      If Lang = 0 Then
                        If what$ = "ΤΟΤΕ" Or what$ = "ΑΛΛΙΩΣ" Then
                            w$ = what$
                                Exit Sub
                            End If
                        Else
                            If what$ = "THEN" Or what$ = "ELSE" Then
                                w$ = what$
                                Exit Sub
                            End If
                        End If
        
End If
End Sub
Public Sub aheadstatusELSEIF(a$, pos As Long, Lang As Long, jump As Boolean, IFCTRL As Long, flag As Boolean)
Dim what$, w$, lenA As Long, level2 As Integer, pos3 As Long
Const second1$ = "ΑΝ", len1 = 2
Const second2$ = "IF", len2 = 2
flag = True
If a$ = vbNullString Then Exit Sub
Dim v1 As Long
If pos = 0 Then pos = 1
lenA = Len(a$)
Do While pos <= lenA
    w$ = Mid$(a$, pos, 1)
    v1 = AscW(w$)
    If Abs(v1) > 9 Then
    
    If w$ = """" Then
            If Len(what$) > 0 Then what$ = vbNullString
        pos = pos + 1
        Do While pos <= lenA
        If Mid$(a$, pos, 1) = """" Then Exit Do
        If AscW(Mid$(a$, pos, 1)) < 32 Then Exit Do
        pos = pos + 1
        Loop
    ElseIf w$ = "(" Then
        If Len(what$) > 0 Then what$ = vbNullString
again22:
      pos = pos + 1

        If Not BlockParam2(a$, pos) Then Exit Do
        If Mid$(a$, pos + 1, 1) = "#" Then
        pos = pos + 1
        GoTo conthere
        ElseIf Mid$(a$, pos + 1, 1) = "(" Then
        pos = pos + 1: GoTo again22
        End If
    ElseIf w$ = "{" Then
       If Len(what$) > 0 Then what$ = vbNullString
       If pos <= lenA Then
        'If Not blockStringAhead(a$, pos) Then Exit Do
        pos = blockLen2(a$, pos + 1)
        If pos = 0 Then pos = lenA: Exit Do
        End If
    Else
        Select Case w$
        Case "%", "$", "0" To "9", vbLf
            If Len(what$) > 0 Then what$ = vbNullString
        
        
        Case " ", ChrW(160), vbCr ',  Check Else and IF
                If Len(what$) > 1 Then
                        
                        If Lang = 0 Then
                        
                        
                            what$ = myUcase(what$)
                            If what$ = "ΑΛΛΙΩΣ" Then
                                                                    
                                pos3 = pos
                                If MaybeIsSymbol3lot(a$, b123, pos3) Then
                                    If level2 = 0 Then pos = pos - 6: Exit Sub
                                Else
                                    aheadstatusSkipParam a$, pos
                                End If
                            ElseIf what$ = "ΤΕΛΟΣ" Then
                                    
                                If FastSymbolAt(pos, a$, second1$, len1) Then
                                    If level2 = 0 Then IFCTRL = 0: Exit Sub
                                    level2 = level2 - 1
                                Else
                                    aheadstatusSkipParam a$, pos
                                End If
                            ElseIf what$ = "ΑΛΛΙΩΣ.ΑΝ" Then
                                    If (Not jump) Or IFCTRL = 2 Then
                                        aheadstatusSkipParam a$, pos
                                    Else
                                    
                                    If level2 = 0 Then
                                    pos = pos - 9: Exit Sub
                                    Else
                                        aheadstatusSkipParam a$, pos
                                    End If
                                    End If
                            ElseIf what$ = second1$ Then  ' skip any nested IF
                                aheadstatusThen a$, pos, 0, what$
                                If what$ <> vbNullString Then
                                    If MaybeIsSymbol3(a$, "{", pos) Then
                                    aheadstatusSTRUCT a$, pos
                                    
                                    ElseIf MaybeIsSymbol3lot(a$, b123, pos) Then
                                    level2 = level2 + 1
                                    Else 'skip line
                                    Do
                                        pos = pos + 1
        
                                    Loop While pos < lenA And Not Mid$(a$, pos, 1) = vbLf
                                    End If
                                Else
                                flag = False
                                SyntaxError
                                Exit Sub
                                End If
                            Else
                            aheadstatusSkipParam a$, pos
                            End If
                          
                        Else
                                what$ = UCase(what$)
                                If what$ = "ELSE" Then
                                    pos3 = pos
                                    If MaybeIsSymbol3lot(a$, b123, pos3) Then
                                        If level2 = 0 Then pos = pos - 4: Exit Sub
                                    Else
                                        aheadstatusSkipParam a$, pos
                                    End If
                                ElseIf what$ = "END" Then
                            
                                    If FastSymbolAt(pos, a$, second2$, len2) Then
                                        If level2 = 0 Then IFCTRL = 0: Exit Sub
                                        level2 = level2 - 1
                                    Else
                                        aheadstatusSkipParam a$, pos
                                    End If
                                ElseIf what$ = "ELSE.IF" Then
                                    If (Not jump) Or IFCTRL = 2 Then
                                        aheadstatusSkipParam a$, pos
                                    Else
                                    
                                    If level2 = 0 Then
                                    pos = pos - 7: Exit Sub
                                    Else
                                        aheadstatusSkipParam a$, pos
                                    End If
                                    End If
                                ElseIf what$ = second2$ Then
                                aheadstatusThen a$, pos, 1, what$
                                If what$ <> vbNullString Then
                                    If MaybeIsSymbol3(a$, "{", pos) Then
                                    aheadstatusSTRUCT a$, pos
                                    
                                    ElseIf MaybeIsSymbol3lot(a$, b123, pos) Then
                                    level2 = level2 + 1
                                    Else 'skip line
                                    Do
                                        pos = pos + 1
        
                                    Loop While pos < lenA And Not Mid$(a$, pos, 1) = vbLf
                                    End If
                                Else
                                flag = False
                                SyntaxError
                                Exit Sub
                                End If
                                
                                Else
                                aheadstatusSkipParam a$, pos
                                End If
                            
                        End If
                    End If
                    what$ = vbNullString
                            pos = pos + 1
        Do
        pos3 = pos + 1
        Select Case Mid$(a$, pos, 1)
        Case " ", Chr$(160), vbCr, vbLf, vbTab
            pos = pos + 1
        End Select
        Loop Until pos3 > pos
        pos = pos - 1

        Case "'", "\"
            If Len(what$) > 0 Then what$ = vbNullString
        Do
        pos = pos + 1
        
        Loop While pos < lenA And Not Mid$(a$, pos, 1) = vbLf
        
        Case ")", "}", Is < " ", "'", "\"
        
        Exit Do
         Case ".", "A" To "Z", "a" To "z", Is >= "Α"
        
        what$ = what$ + w$
        
        Case Else
            If Len(what$) > 0 Then what$ = vbNullString
        End Select
        End If
End If
        pos = pos + 1
        
  
conthere:
  
Loop
pos = lenA + 2

End Sub

'
Public Sub aheadstatusENDIF(a$, pos As Long, Lang As Long, flag As Boolean)
Dim pos2 As Long, what$, w$, lenA As Long, level2 As Integer
Const second1$ = "ΑΝ", len1 = 2
Const second2$ = "IF", len2 = 2
flag = True
If a$ = vbNullString Then Exit Sub
Dim v1 As Long
If pos = 0 Then pos = 1
lenA = Len(a$)
Do While pos <= lenA
    w$ = Mid$(a$, pos, 1)
    v1 = AscW(w$)
    If Abs(v1) > 9 Then
    
    If w$ = """" Then
            If Len(what$) > 0 Then what$ = vbNullString
        pos = pos + 1
        Do While pos <= lenA
        If Mid$(a$, pos, 1) = """" Then Exit Do
        If AscW(Mid$(a$, pos, 1)) < 32 Then Exit Do
        pos = pos + 1
        Loop
    ElseIf w$ = "(" Then
        If Len(what$) > 0 Then what$ = vbNullString
again22:
      pos = pos + 1

        If Not BlockParam2(a$, pos) Then Exit Do
        If Mid$(a$, pos + 1, 1) = "#" Then
        pos = pos + 1
        GoTo conthere
        ElseIf Mid$(a$, pos + 1, 1) = "(" Then
        pos = pos + 1: GoTo again22
        End If
    ElseIf w$ = "{" Then
       If Len(what$) > 0 Then what$ = vbNullString
       If pos <= lenA Then
        'If Not blockStringAhead(a$, pos) Then Exit Do
        pos = blockLen2(a$, pos + 1)
        If pos = 0 Then pos = lenA: Exit Do
        End If
    Else
        Select Case w$
        Case "%", "$", "0" To "9", vbCr, vbLf
            If Len(what$) > 0 Then what$ = vbNullString
        Case " ", ChrW(160)
                If Len(what$) > 1 Then
                        
                        If Lang = 0 Then
                        If Len(what$) = 5 Or Len(what$) = len1 Then
                            what$ = myUcase(what$)
                            If what$ = "ΤΕΛΟΣ" Then
                                    
                                If FastSymbolAt(pos, a$, second1$, len1) Then
                                If level2 = 0 Then Exit Sub
                                level2 = level2 - 1
                                Else
                                aheadstatusSkipParam a$, pos
                                End If
                            ElseIf what$ = second1$ Then
                                aheadstatusThen a$, pos, 0, what$
                                If what$ <> vbNullString Then
                                  If MaybeIsSymbol3(a$, "{", pos) Then
                                    aheadstatusSTRUCT a$, pos
                                    
                                    ElseIf MaybeIsSymbol3lot(a$, b123, pos) Then
                                    level2 = level2 + 1
                                    Else 'skip line
                                    Do
                                        pos = pos + 1
        
                                    Loop While pos < lenA And Not Mid$(a$, pos, 1) = vbLf
                                    End If
                                Else
                                flag = False
                                SyntaxError
                                Exit Sub
                                End If
                            Else
                            aheadstatusSkipParam a$, pos
                            End If
                            Else
                            aheadstatusSkipParam a$, pos
                            End If
                          
                        Else
                            If Len(what$) = 3 Or Len(what$) = len2 Then
                                what$ = UCase(what$)
                                If what$ = "END" Then
                            
                                If FastSymbolAt(pos, a$, second2$, len2) Then
                                    If level2 = 0 Then Exit Sub
                                    level2 = level2 - 1
                                 Else
                                 aheadstatusSkipParam a$, pos
                                End If
                                ElseIf what$ = second2$ Then
                                aheadstatusThen a$, pos, 1, what$
                                If what$ <> vbNullString Then
                                    If MaybeIsSymbol3(a$, "{", pos) Then
                                    aheadstatusSTRUCT a$, pos
                                    
                                    ElseIf MaybeIsSymbol3lot(a$, b123, pos) Then
                                    level2 = level2 + 1
                                    Else 'skip line
                                    Do
                                        pos = pos + 1
        
                                    Loop While pos < lenA And Not Mid$(a$, pos, 1) = vbLf
                                    End If
                                Else
                                flag = False
                                SyntaxError
                                Exit Sub
                                End If
                                
                                Else
                                aheadstatusSkipParam a$, pos
                                End If
                            Else
                            aheadstatusSkipParam a$, pos
                            
                            End If
                            
                        End If
                    End If
                    what$ = vbNullString
                        pos = pos + 1
                        Do
                        pos2 = pos + 1
                        Select Case Mid$(a$, pos, 1)
                        Case " ", Chr$(160), vbTab
                            pos = pos + 1
                        End Select
                        Loop Until pos2 > pos
                        pos = pos - 1

        Case "'", "\"
            If Len(what$) > 0 Then what$ = vbNullString
        Do
        pos = pos + 1
        
        Loop While pos < lenA And Not Mid$(a$, pos, 1) = vbLf
        
        Case ")", "}", Is < " ", "'", "\"
        
        Exit Do
        Case ".", "A" To "Z", "a" To "z", Is >= "Α"
        
        what$ = what$ + w$
        
        Case Else
            If Len(what$) > 0 Then what$ = vbNullString
        End Select
        End If
End If
        pos = pos + 1
        
  
conthere:
  
Loop
pos = lenA + 2

End Sub


Public Sub aheadstatusEND(a$, pos As Long, Lang As Long, second1$, len1 As Long, second2$, len2 As Long, flag As Boolean, v1 As Long)
Dim pos2 As Long, what$, w$, lenA As Long, level2 As Integer
flag = False

If a$ = vbNullString Then Exit Sub
If pos = 0 Then pos = 1
lenA = Len(a$)
Do While pos <= lenA
    w$ = Mid$(a$, pos, 1)
    v1 = AscW(w$)
    If Abs(v1) > 9 Then
    
    If w$ = """" Then
            If Len(what$) > 0 Then what$ = vbNullString
        pos = pos + 1
        Do While pos <= lenA
        If Mid$(a$, pos, 1) = """" Then Exit Do
        If AscW(Mid$(a$, pos, 1)) < 32 Then Exit Do
        pos = pos + 1
        Loop
    ElseIf w$ = "(" Then
        If Len(what$) > 0 Then what$ = vbNullString
again22:
      pos = pos + 1

        If Not BlockParam2(a$, pos) Then Exit Do
        If Mid$(a$, pos + 1, 1) = "#" Then
        pos = pos + 1
        GoTo conthere
        ElseIf Mid$(a$, pos + 1, 1) = "(" Then
        pos = pos + 1: GoTo again22
        End If
    ElseIf w$ = "{" Then
       If Len(what$) > 0 Then what$ = vbNullString
       If pos <= lenA Then
        'If Not blockStringAhead(a$, pos) Then Exit Do
        pos = blockLen2(a$, pos + 1)
        If pos = 0 Then pos = lenA: Exit Do
        End If
    Else
        Select Case w$
        Case "%", "$", "0" To "9", vbCr, vbLf
            If Len(what$) > 0 Then what$ = vbNullString
        Case " ", ChrW(160)
                If Len(what$) > 1 Then
                        
                        If Lang = 0 Then
                        If Len(what$) = 5 Or Len(what$) = len1 Then
                            what$ = myUcase(what$)
                            If what$ = "ΤΕΛΟΣ" Then
                                v1 = pos - 5
                                If FastSymbolAt(pos, a$, second1$, len1) Then
                                If level2 = 0 Then flag = True: Exit Sub
                                level2 = level2 - 1
                                Else
                                aheadstatusSkipParam a$, pos
                                End If
                            ElseIf what$ = second1$ Then
                                pos = pos - 1
                                Do
                                    pos = pos + 1
                                    pos2 = pos
                                    aheadstatusNew a$, pos, flag
                                Loop While flag And MaybeIsSymbol3(a$, ",", pos)
                                If MaybeIsSymbol3(a$, "{", pos) Then
                                    aheadstatusSTRUCT a$, pos
                                ElseIf MaybeIsSymbol3lot(a$, b1234, pos) Then
                                    level2 = level2 + 1
                                    pos = pos + 1
                                End If
                            
                            Else
                            aheadstatusSkipParam a$, pos
                            End If
                            Else
                            aheadstatusSkipParam a$, pos
                            End If
                          
                        Else
                            If Len(what$) = 3 Or Len(what$) = len2 Then
                                what$ = UCase(what$)
                                If what$ = "END" Then
                                v1 = pos - 3
                                If FastSymbolAt(pos, a$, second2$, len2) Then
                                    If level2 = 0 Then flag = True: Exit Sub
                                    level2 = level2 - 1
                                 Else
                                 aheadstatusSkipParam a$, pos
                                End If
                                ElseIf what$ = second2$ Then
                                    pos = pos - 1
                                    Do
                                        pos = pos + 1
                                        pos2 = pos
                                        aheadstatusNew a$, pos, flag
                                    Loop While flag And MaybeIsSymbol3(a$, ",", pos)
                                    
                                    If MaybeIsSymbol3(a$, "{", pos) Then
                                        aheadstatusSTRUCT a$, pos
                                    ElseIf MaybeIsSymbol3lot(a$, b1234, pos) Then
                                        level2 = level2 + 1
                                        pos = pos + 1
                                    End If
                                Else
                                aheadstatusSkipParam a$, pos
                                End If
                            Else
                            aheadstatusSkipParam a$, pos
                            
                            End If
                            
                        End If
                    
                    End If
                    what$ = vbNullString
                            pos = pos + 1
                            Do
                            pos2 = pos + 1
                            Select Case Mid$(a$, pos, 1)
                            Case " ", Chr$(160), vbTab
                                pos = pos + 1
                            End Select
                            Loop Until pos2 > pos
                            pos = pos - 1
        Case "'", "\"
            If Len(what$) > 0 Then what$ = vbNullString
        Do
        pos = pos + 1
        
        Loop While pos < lenA And Not Mid$(a$, pos, 1) = vbLf
        
        Case ")", "}", Is < " ", "'", "\"
        
        Exit Do
        Case "A" To "Z", "a" To "z", Is >= "Α"
        
        what$ = what$ + w$
        
        Case Else
            If Len(what$) > 0 Then what$ = vbNullString
        End Select
        End If
End If
        pos = pos + 1
        
  
conthere:
  
Loop
pos = lenA + 2

End Sub
Sub dumpModule(no As Long)
Dim a$, i As Long, GUARD As Long, oldi As Long, ok As Boolean, r As Long
Dim part$, trimright
a$ = sbf(no).sb
GUARD = Len(a$)
i = 1
While i <= GUARD
i = MyTrimLi(a$, i)
While Mid$(a$, i, 2) = vbCrLf And i <= GUARD
i = i + 2
Wend
While MaybeIsSymbol3lot(a$, "'\#", i)
i = i + 1
While Not Mid$(a$, i, 2) = vbCrLf And i <= GUARD
i = i + 1
Wend
While Mid$(a$, i, 2) = vbCrLf And i <= GUARD
i = i + 2
Wend
i = MyTrimLi(a$, i)
Wend

If i <= GUARD Then
    oldi = i
 
    
    GetCodePart a$, i
    
    trimright = MyTrimRfrom(a$, oldi, i)
    If IsNumberLabel2(a$, part$, oldi, i) Then
    Debug.Print "<<LABEL>> " & part$
    oldi = MyTrimLi(a$, oldi)
    End If
   If oldi < trimright Then
    Debug.Print "<<Part>>"
    Debug.Print Mid$(a$, oldi, trimright - oldi)
    
    End If
    If i = 0 Then Exit Sub
    If Mid$(a$, i) = vbCr Then i = i + 2 Else i = i + 1
End If
Wend
'Debug.Print a$
End Sub
Sub test2(Optional a$)
Dim i As Long, j As Long, s$, ok As Boolean
a$ = "a>15  {" + vbCrLf + " Print 1000 " + vbCrLf + "} 1 2 3"
a$ = "nmn nmmn {asdsad}+{dsdfsf}>a${hjkh}"
a$ = "{ddg}+b$ {hkjh}"
a$ = "not {a}>{aa}+aaa$ and 10>len({3}) and {z}=basdas$ gghjh{asdas}"
a$ = " 1>2 and  {alfa}={beta} and 1>2"
a$ = "a$+b$ {}" ' + vbCr
a$ = "              a$> b$"
a$ = "list:=1,2,3,4 : Print 12"
i = 1
aheadstatusNew a$, i, ok
Debug.Print Left$(a$, i), i
Debug.Print ok
End Sub
'
Public Sub aheadstatusNew(a$, pos As Long, flag As Boolean)
Dim b$, part$, w$, pos2 As Long
flag = False
If a$ = vbNullString Then Exit Sub

Dim v1 As Long
If pos = 0 Then pos = 1
Do While pos <= Len(a$)
    w$ = Mid$(a$, pos, 1)
    v1 = AscW(w$)
    If Abs(v1) > 9 Then
    If part$ = vbNullString And w$ = "0" Then
        If pos + 2 <= Len(a$) Then
            If LCase(Mid$(a$, pos, 2)) Like "0[xχ]" Then
                pos = pos + 2
                Do While pos <= Len(a$)
                If Not Mid$(a$, pos, 1) Like "[0-9a-fA-F]" Then Exit Do
                pos = pos + 1
                Loop
                b$ = b$ & "N"
                If pos <= Len(a$) Then
                    w$ = Mid$(a$, pos, 1)
                Else
                    Exit Do
                End If
            End If
        End If
    End If

    If w$ = """" Then
        If part$ <> "" Then
        b$ = b$ & part$
        End If
        part$ = "S"
        pos = pos + 1
        Do While pos <= Len(a$)
        If Mid$(a$, pos, 1) = """" Then Exit Do
    If AscW(Mid$(a$, pos, 1)) < 32 Then Exit Do
   
        pos = pos + 1
        Loop

    ElseIf w$ = "(" Then
again:
        If part$ <> "" Then
            ' after
            If part$ = "S" Then
            '
             If Mid$(a$, pos + 1, 1) = ")" Then pos = pos + 2: GoTo conthere
             
            End If
            ElseIf Right$(b$, 1) = "a" Then
            b$ = Left$(b$, Len(b$) - 1)
            part$ = vbNullString
            Else
            part$ = "N"
              
        End If
again22:
      pos = pos + 1

        If Not BlockParam2(a$, pos) Then Exit Do
        If Mid$(a$, pos + 1, 1) = "#" Then
        b$ = vbNullString
        part$ = "N"
        pos = pos + 1
        GoTo conthere
        ElseIf Mid$(a$, pos + 1, 1) = "(" Then
        pos = pos + 1: GoTo again22
        End If
       If Mid$(a$, pos + 1, 1) <> "." And Mid$(a$, pos + 1, 2) <> "=>" Then
       b$ = b$ & part$
       End If
        part$ = vbNullString
        
    ElseIf w$ = "{" Then

Select Case Left$(Right$(b$, 2), 1)
    Case "l"
        Exit Do
    Case "o"
    If Right$(b$, 1) = "N" Then Exit Do
    Case "S"
    If Right$(b$, 1) = "a" Then
    Select Case Left$(Right$(b$, 3), 1)
    Case "l"
         Exit Do
    Case "o"
   Exit Do

    End Select
    
       If Left$(Right$(b$, 3), 1) = "l" Then Exit Do
    End If
    End Select
    If part$ <> "" Then
        b$ = b$ & part$
        End If
        part$ = "S"
        
        
            If pos <= Len(a$) Then
            pos2 = pos
        If Not blockStringAhead(a$, pos) Then Exit Do
        If Right$(b$, 1) = "N" Then
        If Not MaybeIsSymbol3lot(a$, "+<>=~", (pos + 1)) Then
        pos = pos2
        Exit Do
        End If
        End If
        End If
      

      

    Else
        Select Case w$
        Case ","  ' bye bye
        Exit Do
        Case "%"
            If part$ = vbNullString Then
            End If
        Case "$"
            If part$ = vbNullString Then
                If b$ = vbNullString Then
                    part$ = "N"
                ElseIf Right$(b$, 1) = "o" Then
                    part$ = "N"
                Else
                    flag = Len(b$) > 0
                    Exit Sub
                End If
            ElseIf part$ = "N" Then
                    b$ = b$ & "Sa"
                    If Mid$(a$, pos + 1, 1) = "." Then pos = pos + 1
                    part$ = vbNullString
            End If
        Case "+", "-", "|"
                    b$ = b$ & part$
                    If b$ = vbNullString Then
                    Else
                    
                part$ = "o"
                End If
        Case "*", "/", "^"
            If part$ <> "o" Then
            b$ = b$ & part$
            End If
            part$ = "o"
        Case " ", ChrW(160)
            If part$ <> "" Then
            b$ = b$ & part$
            part$ = vbNullString
            Else
            'skip
            End If
        pos = pos + 1
        Do
        pos2 = pos + 1
        Select Case Mid$(a$, pos, 1)
        Case " ", Chr$(160), vbTab
            pos = pos + 1
        End Select
        Loop Until pos2 > pos
        pos = pos - 1
        Case "0" To "9", "."
            If part$ = "N" Then
            If Len(a$) < pos Then
                If Mid$(a$, pos + 1, 1) Like "[&@#%~]" Then pos = pos + 1
            End If
            
            ElseIf part$ = "S" Then
            
            Else
            
            b$ = b$ & part$
            part$ = "N"
            End If
        Case "&"
        If part$ = vbNullString Then
        part$ = "S"
        ElseIf part$ = "N" Then
        b$ = b$ + part$
        part$ = vbNullString
        Else
        b$ = part$
        part$ = "S"
        End If
        Case "e", "E", "ε", "Ε"
            If part$ = "N" Then

            ElseIf part$ = "S" Then
            
            
            Else
            b$ = b$ & part$
            part$ = "N"
            End If
         Case ">", "<", "~"
            If Len(a$) >= pos + 1 Then
            If Mid$(a$, pos, 2) = Mid$(a$, pos, 1) Then
                b$ = b$ & part$
                If b$ = vbNullString Then
                        Else
                        
                    part$ = "o"
                    pos = pos + 1
                    End If
                ElseIf w$ = ">" And pos > 1 Then
                    If Mid$(a$, pos - 1, 2) = "->" Then ' "->"
                   If Right$(b$, 1) = "S" Then
                    b$ = b$ + part$
                    part$ = "N"
                    Else
                      '  part$ = vbNullString
                    End If
                        
                    End If
                End If
            End If
            GoTo there1
         Case "="
            If Mid$(a$, pos + 1, 1) = ">" Then
                pos = pos + 2
                GoTo conthere
                End If
there1:
                If b$ & part$ <> "" Then
               
                w$ = Replace(b$ & part$, "a", "")
            part$ = vbNullString
               
               If Len(b$) > 1 Then If Left$(b$, Len(b$) - 1) <> "l" Then part$ = "l"
                Else
                Exit Do
                End If

        Case ")", "}", Is < " ", ":", ";", "'", "\"
        Exit Do
        Case Else
        If part$ = "N" Then
        ElseIf part$ = "S" Then
        Else
        
     b$ = b$ & part$
     part$ = "N"

            End If
        End Select
        End If
End If
        pos = pos + 1
        
conthere:
  
Loop


    flag = Len(b$) <> 0




End Sub




Public Function aheadstatusFast(a$) As String 'ok
Dim b$, part$, w$, pos2 As Long, pos As Long

If a$ = vbNullString Then Exit Function
Dim v1 As Integer
pos = 1
Do While pos <= Len(a$)
    w$ = Mid$(a$, pos, 1)
    v1 = AscW(w$)
    If v1 = 2 Then
        If part$ <> "" Then
        b$ = b$ & part$
        End If
        part$ = "S"
        pos = pos + CLng("&H" & Mid$(a$, pos + 1, 8)) + 8
        w$ = """"
   
    
    ElseIf Abs(v1) > 9 Then
    If part$ = vbNullString And w$ = "0" Then
        If pos + 2 <= Len(a$) Then
            If LCase(Mid$(a$, pos, 2)) Like "0[xχ]" Then
            'hexadecimal literal number....
                pos = pos + 2
                Do While pos <= Len(a$)
                If Not Mid$(a$, pos, 1) Like "[0-9a-fA-F]" Then Exit Do
                pos = pos + 1
                Loop
                b$ = b$ & "N"
                If pos <= Len(a$) Then
                    w$ = Mid$(a$, pos, 1)
                Else
                    Exit Do
                End If
            End If
        End If
    End If

    If w$ = """" Then
        If part$ <> "" Then
        b$ = b$ & part$
        End If
        part$ = "S"
        pos = pos + 1
        Do While pos <= Len(a$)
        If Mid$(a$, pos, 1) = """" Then Exit Do
    If AscW(Mid$(a$, pos, 1)) < 32 Then Exit Do
   
        pos = pos + 1
        Loop

    ElseIf w$ = "(" Then
again:
        If part$ <> "" Then
            ' after
            If part$ = "S" Then
            '
             If Mid$(a$, pos + 1, 1) = ")" Then pos = pos + 2: GoTo conthere
             
            End If
            ElseIf Right$(b$, 1) = "a" Then
            b$ = Left$(b$, Len(b$) - 1)
            part$ = vbNullString
            Else
            part$ = "N"
              
        End If
again22:
      pos = pos + 1

        If Not BlockParam2(a$, pos) Then Exit Do
        If Mid$(a$, pos + 1, 1) = "#" Then
        b$ = vbNullString
        part$ = "N"
        pos = pos + 1
        GoTo conthere
        ElseIf Mid$(a$, pos + 1, 1) = "(" Then
        pos = pos + 1: GoTo again22
        End If
       If Mid$(a$, pos + 1, 1) <> "." And Mid$(a$, pos + 1, 2) <> "=>" Then
       b$ = b$ & part$
       End If
        part$ = vbNullString
        
    ElseIf w$ = "{" Then
Select Case Left$(Right$(b$, 2), 1)
    Case "l"
       Exit Do
    Case "o"
    If Right$(b$, 1) = "N" Then Exit Do
    Case "S"
    If Right$(b$, 1) = "a" Then
    Select Case Left$(Right$(b$, 3), 1)
    Case "l"
        Exit Do
    Case "o"
    Exit Do

    End Select
    
       If Left$(Right$(b$, 3), 1) = "l" Then Exit Do
    End If
    End Select
         
    If part$ <> "" Then
        b$ = b$ & part$
        End If
        part$ = "S"
        
If pos <= Len(a$) Then
            pos2 = pos
        If Not blockStringAhead(a$, pos) Then Exit Do
        If Right$(b$, 1) = "N" Then
        If Not MaybeIsSymbol3lot(a$, "+<>=~", (pos + 1)) Then
        pos = pos2
        Exit Do
        End If
        End If
        End If
      

    Else
        Select Case w$
        Case ","  ' bye bye
        Exit Do
        Case "%"
            If part$ = vbNullString Then
            End If
        Case "$"
            If part$ = vbNullString Then
                If b$ = vbNullString Then
                    part$ = "N"
                ElseIf Right$(b$, 1) = "o" Then
                    part$ = "N"
                Else
                    aheadstatusFast = b$
                    Exit Function
                End If
            ElseIf part$ = "N" Then
                    b$ = b$ & "Sa"
                    If Mid$(a$, pos + 1, 1) = "." Then pos = pos + 1
                    part$ = vbNullString
            End If
        Case "+", "-", "|"
                    b$ = b$ & part$
                    If b$ = vbNullString Then
                    Else
                    
                part$ = "o"
                End If
        Case "*", "/", "^"
            If part$ <> "o" Then
            b$ = b$ & part$
            End If
            part$ = "o"
        Case " ", ChrW(160)
            If part$ <> "" Then
            b$ = b$ & part$
            part$ = vbNullString
            Else
            'skip
            End If
        
        Case "0" To "9", "."
            If part$ = "N" Then
            If Len(a$) < pos Then
                If Mid$(a$, pos + 1, 1) Like "[&@#%~]" Then pos = pos + 1
            End If
            
            ElseIf part$ = "S" Then
            
            Else
            
            b$ = b$ & part$
            part$ = "N"
            End If
        Case "&"
        If part$ = vbNullString Then
        part$ = "S"
        ElseIf part$ = "N" Then
        b$ = b$ + part$
        part$ = vbNullString
        Else
        b$ = part$
        part$ = "S"
        End If
        Case "e", "E", "ε", "Ε"
            If part$ = "N" Then

            ElseIf part$ = "S" Then
            
            
            Else
            b$ = b$ & part$
            part$ = "N"
            End If
         Case ">", "<", "~"
            If Len(a$) >= pos + 1 Then
            If Mid$(a$, pos, 2) = Mid$(a$, pos, 1) Then
                b$ = b$ & part$
                If b$ = vbNullString Then
                        Else
                        
                    part$ = "o"
                    pos = pos + 1
                    End If
                ElseIf w$ = ">" And pos > 1 Then
                    If Mid$(a$, pos - 1, 2) = "->" Then ' "->"
                   If Right$(b$, 1) = "S" Then
                    b$ = b$ + part$
                    part$ = "N"
                    Else
                      '  part$ = vbNullString
                    End If
                        
                    End If
                End If
            End If
            GoTo there1
         Case "="
            If Mid$(a$, pos + 1, 1) = ">" Then
                pos = pos + 2
                GoTo conthere
                End If
there1:
                If b$ & part$ <> "" Then
               
               b$ = Replace(b$ & part$, "a", "")
                part$ = vbNullString
               
                If Left$(b$, Len(b$) - 1) <> "l" Then part$ = "l": Exit Do
                
                Else
                Exit Do
                End If

        Case ")", "}", Is < " ", ":", ";", "'", "\"
        Exit Do
        Case Else
        If part$ = "N" Then
        ElseIf part$ = "S" Then
        Else
        
     b$ = b$ & part$
     part$ = "N"

            End If
        End Select
        End If
End If
        pos = pos + 1
        
conthere:
  
Loop
 
   
   


    aheadstatusFast = b$ & part$




End Function


Public Function aheadstatus(a$, Optional srink As Boolean = True, Optional pos As Long = 1) As String 'ok
Dim b$, part$, w$, pos2 As Long

If a$ = vbNullString Then Exit Function
Dim v1 As Long
If pos = 0 Then pos = 1
Do While pos <= Len(a$)
    w$ = Mid$(a$, pos, 1)
    v1 = AscW(w$)
    If v1 = 2 Then
        If part$ <> "" Then
        b$ = b$ & part$
        End If
        part$ = "S"
        pos = pos + CLng("&H" & Mid$(a$, pos + 1, 8)) + 8
        w$ = """"
   
    
    ElseIf Abs(v1) > 9 Then
    If part$ = vbNullString And w$ = "0" Then
        If pos + 2 <= Len(a$) Then
            If LCase(Mid$(a$, pos, 2)) Like "0[xχ]" Then
            'hexadecimal literal number....
                pos = pos + 2
                Do While pos <= Len(a$)
                If Not Mid$(a$, pos, 1) Like "[0-9a-fA-F]" Then Exit Do
                pos = pos + 1
                Loop
                b$ = b$ & "N"
                If pos <= Len(a$) Then
                    w$ = Mid$(a$, pos, 1)
                Else
                    Exit Do
                End If
            End If
        End If
    End If

    If w$ = """" Then
        If part$ <> "" Then
        b$ = b$ & part$
        End If
        part$ = "S"
        pos = pos + 1
        Do While pos <= Len(a$)
        If Mid$(a$, pos, 1) = """" Then Exit Do
    If AscW(Mid$(a$, pos, 1)) < 32 Then Exit Do
   
        pos = pos + 1
        Loop

    ElseIf w$ = "(" Then
again:
        If part$ <> "" Then
            ' after
            If part$ = "S" Then
            '
             If Mid$(a$, pos + 1, 1) = ")" Then pos = pos + 2: GoTo conthere
             
            End If
            ElseIf Right$(b$, 1) = "a" Then
            b$ = Left$(b$, Len(b$) - 1)
            part$ = vbNullString
            Else
            part$ = "N"
              
        End If
again22:
      pos = pos + 1

        If Not BlockParam2(a$, pos) Then Exit Do
        If Mid$(a$, pos + 1, 1) = "#" Then
        b$ = vbNullString
        part$ = "N"
        pos = pos + 1
        GoTo conthere
        ElseIf Mid$(a$, pos + 1, 1) = "(" Then
        pos = pos + 1: GoTo again22
        End If
       If Mid$(a$, pos + 1, 1) <> "." And Mid$(a$, pos + 1, 2) <> "=>" Then
       b$ = b$ & part$
       End If
        part$ = vbNullString
        
    ElseIf w$ = "{" Then
Select Case Left$(Right$(b$, 2), 1)
    Case "l"
       Exit Do
    Case "o"
    If Right$(b$, 1) = "N" Then Exit Do
    Case "S"
    If Right$(b$, 1) = "a" Then
    Select Case Left$(Right$(b$, 3), 1)
    Case "l"
        Exit Do
    Case "o"
    Exit Do

    End Select
    
       If Left$(Right$(b$, 3), 1) = "l" Then Exit Do
    End If
    End Select
         
    If part$ <> "" Then
        b$ = b$ & part$
        End If
        part$ = "S"
        
If pos <= Len(a$) Then
            pos2 = pos
        If Not blockStringAhead(a$, pos) Then Exit Do
        If Right$(b$, 1) = "N" Then
        If Not MaybeIsSymbol3lot(a$, "+<>=~", (pos + 1)) Then
        pos = pos2
        Exit Do
        End If
        End If
        End If
      

    Else
        Select Case w$
        Case ","  ' bye bye
        Exit Do
        Case "%"
            If part$ = vbNullString Then
            End If
        Case "$"
            If part$ = vbNullString Then
                If b$ = vbNullString Then
                    part$ = "N"
                ElseIf Right$(b$, 1) = "o" Then
                    part$ = "N"
                Else
                    aheadstatus = b$
                    Exit Function
                End If
            ElseIf part$ = "N" Then
                    b$ = b$ & "Sa"
                    If Mid$(a$, pos + 1, 1) = "." Then pos = pos + 1
                    part$ = vbNullString
            End If
        Case "+", "-", "|"
                    b$ = b$ & part$
                    If b$ = vbNullString Then
                    Else
                    
                part$ = "o"
                End If
        Case "*", "/", "^"
            If part$ <> "o" Then
            b$ = b$ & part$
            End If
            part$ = "o"
        Case " ", ChrW(160)
            If part$ <> "" Then
            b$ = b$ & part$
            part$ = vbNullString
            Else
            'skip
            End If
        
        Case "0" To "9", "."
            If part$ = "N" Then
            If Len(a$) < pos Then
                If Mid$(a$, pos + 1, 1) Like "[&@#%~]" Then pos = pos + 1
            End If
            
            ElseIf part$ = "S" Then
            
            Else
            
            b$ = b$ & part$
            part$ = "N"
            End If
        Case "&"
        If part$ = vbNullString Then
        part$ = "S"
        ElseIf part$ = "N" Then
        b$ = b$ + part$
        part$ = vbNullString
        Else
        b$ = part$
        part$ = "S"
        End If
        Case "e", "E", "ε", "Ε"
            If part$ = "N" Then

            ElseIf part$ = "S" Then
            
            
            Else
            b$ = b$ & part$
            part$ = "N"
            End If
         Case ">", "<", "~"
            If Len(a$) >= pos + 1 Then
            If Mid$(a$, pos, 2) = Mid$(a$, pos, 1) Then
                b$ = b$ & part$
                If b$ = vbNullString Then
                        Else
                        
                    part$ = "o"
                    pos = pos + 1
                    End If
                ElseIf w$ = ">" And pos > 1 Then
                    If Mid$(a$, pos - 1, 2) = "->" Then ' "->"
                   If Right$(b$, 1) = "S" Then
                    b$ = b$ + part$
                    part$ = "N"
                    Else
                      '  part$ = vbNullString
                    End If
                        
                    End If
                End If
            End If
            GoTo there1
         Case "="
            If Mid$(a$, pos + 1, 1) = ">" Then
                pos = pos + 2
                GoTo conthere
                End If
there1:
                If b$ & part$ <> "" Then
               
                w$ = Replace(b$ & part$, "a", "")
            part$ = vbNullString
               If srink Then
                  Do
                b$ = w$
                w$ = Replace(b$, "NN", "N")
                Loop While w$ <> b$
                         Do
                        b$ = w$
                          w$ = Replace(b$, "SlS", "N")
                          Loop While w$ <> b$
                            Do
                          b$ = w$
                          w$ = Replace(b$, "NlN", "N")
                          Loop While w$ <> b$
    
                Do
                b$ = w$
                w$ = Replace(b$, "NoN", "N")
                Loop While w$ <> b$
                
                Do
                b$ = w$
                w$ = Replace(b$, "SoS", "S")
                Loop While w$ <> b$
                Else
              b$ = w$
               End If
               
                If Left$(b$, Len(b$) - 1) <> "l" Then part$ = "l"
                Else
                Exit Do
                End If

        Case ")", "}", Is < " ", ":", ";", "'", "\"
        Exit Do
        Case Else
        If part$ = "N" Then
        ElseIf part$ = "S" Then
        Else
        
     b$ = b$ & part$
     part$ = "N"

            End If
        End Select
        End If
End If
        pos = pos + 1
        
conthere:
  
Loop

    w$ = Replace(b$ & part$, "a", "")
    
    b$ = w$
If srink Then
         Do
  b$ = w$

    w$ = Replace(b$, "SlS", "N")
    Loop While w$ <> b$
      Do
    b$ = w$
    w$ = Replace(b$, "NlN", "N")
    Loop While w$ <> b$
    
    Do
    b$ = w$
    w$ = Replace(b$, "NoN", "N")
    Loop While w$ <> b$
    
    Do
    b$ = w$
    w$ = Replace(b$, "SoS", "S")
    Loop While w$ <> b$
End If
   
   
   
   


    aheadstatus = b$




End Function

Public Function aheadstatusStr(a$) As Long 'ok
Dim w$, pos2 As Long, pos As Long
If a$ = vbNullString Then Exit Function
Dim v1 As Integer
 pos = 1
Do While pos <= Len(a$)
    w$ = Mid$(a$, pos, 1)
    v1 = AscW(w$)
    
    If v1 = 2 Then
    If pos2 = 0 Then aheadstatusStr = True
    Exit Function
    
    ElseIf Abs(v1) > 9 Then
    
    If v1 = 34 Then
        If pos2 = 0 Then aheadstatusStr = True
        Exit Function

    ElseIf w$ = "(" Then
        
again22:
        pos = pos + 1

        If Not BlockParam2(a$, pos) Then Exit Do
        If Mid$(a$, pos + 1, 1) = "#" Then
        pos = pos + 1
        GoTo conthere
        ElseIf Mid$(a$, pos + 1, 1) = "(" Then
        pos = pos + 1: GoTo again22
        End If
       If Mid$(a$, pos + 1, 1) <> "." And Mid$(a$, pos + 1, 2) <> "=>" Then
       
       End If
       
        
    ElseIf w$ = "{" Then
         
        aheadstatusStr = True
        Exit Function
    Else
        Select Case w$
        Case " ", ChrW(160)
        If pos2 > 0 Then Exit Function
        Case "%"
            If Not Mid$(a$, pos + 1, 1) = "(" Then Exit Function
        Case "$"
            aheadstatusStr = True
            Exit Function
        Case "="
            If Not Mid$(a$, pos + 1, 1) = ">" Then Exit Function
            pos = pos + 1
        Case "-"
            If Not Mid$(a$, pos + 1, 1) = ">" Then Exit Function
            pos = pos + 1
        Case "."
         pos2 = pos2 + 1
        Case "0" To "9"
        If pos2 = 0 Then Exit Function
        Case "&"
        If pos2 = 0 Then aheadstatusStr = True: Exit Function
       
         Case Is < " ", ")", "}", ":", ";", "'", "\", "<", "~", ",", ">", "+", "*", "/", "'", "~", "|"
         Exit Function
        Exit Do
        Case Else
        pos2 = pos2 + 1
        End Select
        End If
End If
        pos = pos + 1
        
conthere:
  
Loop



End Function

Function blockStringAhead(s$, pos1 As Long) As Long
Dim i As Long, j As Long, c As Long
c = Len(s$)
i = pos1
If i > c Then blockStringAhead = c: Exit Function
Do

Select Case AscW(Mid$(s$, i, 1))
Case 34
Do While i < c
i = i + 1
If AscW(Mid$(s$, i, 1)) = 34 Then Exit Do

Loop
Case 123
j = j - 1
Case 125
j = j + 1: If j = 0 Then Exit Do
End Select
i = i + 1
Loop Until i > c
If j = 0 Then
pos1 = i
blockStringAhead = True
Else
blockStringAhead = False
End If


End Function
Public Function CleanStr(sStr As String, noValidcharList As String) As String
Dim a$, i As Long '', ddt As Boolean
If noValidcharList <> "" Then
''If Len(sStr) > 20000 Then ddt = True
If Len(sStr) > 0 Then
For i = 1 To Len(sStr)
''If ddt Then If i Mod 321 = 0 Then Sleep 20
If InStr(noValidcharList, Mid$(sStr, i, 1)) = 0 Then a$ = a$ & Mid$(sStr, i, 1)

Next i
End If
Else
a$ = sStr
End If
CleanStr = a$
End Function
Public Sub ResCounter()
k1 = 0
End Sub

Public Function CheckStackObj(bstack As basetask, anything As Object, Optional counter As Long) As Boolean
If TypeOf bstack.lastobj Is mHandler Then
        If bstack.lastobj.t1 <> 3 Then Exit Function
        counter = bstack.lastobj.index_cursor + 1
        Set anything = bstack.lastobj
        Set bstack.lastobj = Nothing
        If CheckDeepAny(anything) Then CheckStackObj = True
End If
        
End Function
Sub myesc(b$)
MyErMacro b$, "Escape", "Διακοπή εκτέλεσης"
End Sub
Sub wrongsizeOrposition(a$)
    MyErMacro a$, "Wrong Size-Position for reading buffer", "Λάθος Μέγεθος-θέση, για διάβασμα Διάρθρωσης"
End Sub
Sub wrongweakref(a$)
MyErMacro a$, "Wrong weak reference", "λάθος ισχνής αναφοράς"
End Sub
Sub negsqrt(a$)
MyErMacro a$, "negative number for root", "αρνητικός σε ρίζα"
End Sub
Sub expecteddecimal(a$)
MyErMacro a$, "Expected decimal separator char", "Περίμενα χαρακτήρα διαχωρισμού δεκαδικών"
End Sub
Sub wrongexprinstring(a$)
MyErMacro a$, "Wrong expression in string", "λάθος μαθηματική έκφραση στο αλφαριθμητικό"
End Sub
Sub unknownoffset(a$, s$)
MyErMacro a$, "Unknown Offset " & s$, "’γνωστη Μετάθεση " & s$
End Sub
Sub wronguseofenum(a$)
MyErMacro a$, "Wrong use of enumerator", "λάθος χρήση απαριθμητή"
End Sub
Sub nosuchfile()
MyEr "No such file", "Δεν υπάρχει τέτοιο αρχείο"
End Sub
Public Function MyDoEvents()
On Error GoTo there
If TaskMaster Is Nothing Then
DoEvents
Exit Function
ElseIf Not TaskMaster.Processing And TaskMaster.QueueCount = 0 Then
        DoEvents
Exit Function
Else
If TaskMaster.PlayMusic Then
                  TaskMaster.OnlyMusic = True
                      TaskMaster.TimerTick
                    TaskMaster.OnlyMusic = False
                 End If
        TaskMaster.StopProcess
         TaskMaster.TimerTick
         DoEvents
         TaskMaster.StartProcess
If TaskMaster Is Nothing Then Exit Function

End If
Exit Function
there:
If Not TaskMaster Is Nothing Then TaskMaster.RestEnd1
End Function

Public Function ContainsUTF16(ByRef Source() As Byte, Optional maxsearch As Long = -1) As Long
  Dim i As Long, lUBound As Long, lUBound2 As Long, lUBound3 As Long
  Dim CurByte As Byte, CurByte1 As Byte
  Dim CurBytes As Long, CurBytes1 As Long
    lUBound = UBound(Source)
    If lUBound > 4 Then
    CurByte = Source(0)
    CurByte1 = Source(1)
    If maxsearch = -1 Then
    maxsearch = lUBound - 1
    ElseIf maxsearch < 8 Or maxsearch > lUBound - 1 Then
    maxsearch = lUBound - 1
    End If
    
    
    
    For i = 2 To maxsearch Step 2
        If CurByte1 = 0 And CurByte < 31 Then CurBytes1 = CurBytes1 + 1
        If CurByte = 0 And CurByte1 < 31 Then CurBytes = CurBytes + 1
        If Source(i) = CurByte Then
            CurBytes = CurBytes + 1
        Else
            CurByte = Source(i)
        End If
        If Source(i + 1) = CurByte1 Then
            CurBytes1 = CurBytes1 + 1
        Else
            CurByte1 = Source(i + 1)
        End If
        
    Next i
    End If
    If CurBytes1 = CurBytes And CurBytes1 * 3 >= lUBound Then
    ContainsUTF16 = 0
    Else
    If CurBytes1 * 3 >= lUBound Then
    ContainsUTF16 = 1
    ElseIf CurBytes * 3 >= lUBound Then
    ContainsUTF16 = 2
    Else
    ContainsUTF16 = 0
    End If
    End If
End Function
Public Function ContainsUTF8(ByRef Source() As Byte) As Boolean
  Dim i As Long, lUBound As Long, lUBound2 As Long, lUBound3 As Long
  Dim CurByte As Byte
    lUBound = UBound(Source)
    lUBound2 = lUBound - 2
    lUBound3 = lUBound - 3
    If lUBound > 2 Then
    
    For i = 0 To lUBound - 1
      CurByte = Source(i)
        If (CurByte And &HE0) = &HC0 Then
        If (Source(i + 1) And &HC0) = &H80 Then
            ContainsUTF8 = ContainsUTF8 Or True
             i = i + 1
             Else
                ContainsUTF8 = False
                Exit For
            End If
        

        ElseIf (CurByte And &HF0) = &HE0 Then
        ' 2 bytes
        If (Source(i + 1) And &HC0) = &H80 Then
            i = i + 1
            If i < lUBound2 Then
            If (Source(i + 1) And &HC0) = &H80 Then
                ContainsUTF8 = ContainsUTF8 Or True
                i = i + 1
            Else
                ContainsUTF8 = False
                Exit For
            End If
                Else
                ContainsUTF8 = False
                Exit For
            End If
        Else
            ContainsUTF8 = False
            Exit For
        End If
        ElseIf (CurByte And &HF8) = &HF0 Then
        ' 2 bytes
        If (Source(i + 1) And &HC0) = &H80 Then
            i = i + 1
            If i < lUBound2 Then
               If (Source(i + 1) And &HC0) = &H80 Then
                    ContainsUTF8 = ContainsUTF8 Or True
                    i = i + 1
                    If i < lUBound3 Then
                       If (Source(i + 1) And &HC0) = &H80 Then
                            ContainsUTF8 = ContainsUTF8 Or True
                            i = i + 1
                        Else
                            ContainsUTF8 = False
                            Exit For
                        End If
                        
                    Else
                        ContainsUTF8 = False
                        Exit For
                    End If
                Else
                    ContainsUTF8 = False
                    Exit For
                End If
                
            Else
                ContainsUTF8 = False
                Exit For
            End If
        Else
            ContainsUTF8 = False
            Exit For
        End If
        
        
        End If
        
    Next i
    End If
    

End Function
Function ReadUnicodeOrANSI(FileName As String, Optional ByVal EnsureWinLFs As Boolean, Optional feedback As Long) As String
Dim i&, FNr&, BLen&, WChars&, BOM As Integer, BTmp As Byte, b() As Byte
Dim mLof As Long, nobom As Long
nobom = 1
' code from Schmidt, member of vbforums
If FileName = vbNullString Then Exit Function
On Error Resume Next
If GetDosPath(FileName) = vbNullString Then MissFile: Exit Function
 On Error GoTo ErrHandler
  BLen = FileLen(GetDosPath(FileName))
'  If Err.Number = 53 Then missfile: Exit Function
 
  If BLen = 0 Then Exit Function
  
  FNr = FreeFile
  Open GetDosPath(FileName) For Binary Access Read As FNr
      Get FNr, , BOM
    Select Case BOM
      Case &HFEFF, &HFFFE 'one of the two possible 16 Bit BOMs
        If BLen >= 3 Then
          ReDim b(0 To BLen - 3): Get FNr, 3, b 'read the Bytes
utf16conthere:
          feedback = 0
          If BOM = &HFFFE Then 'big endian, so lets swap the byte-pairs
          feedback = 1
            For i = 0 To UBound(b) Step 2
              BTmp = b(i): b(i) = b(i + 1): b(i + 1) = BTmp
            Next
          End If
          ReadUnicodeOrANSI = b
        End If
      Case &HBBEF 'the start of a potential UTF8-BOM
        Get FNr, , BTmp
        If BTmp = &HBF Then 'it's indeed the UTF8-BOM
        feedback = 2
          If BLen >= 4 Then
            ReDim b(0 To BLen - 4): Get FNr, 4, b 'read the Bytes
            WChars = MultiByteToWideChar(65001, 0, b(0), BLen - 3, 0, 0)
            ReadUnicodeOrANSI = space$(WChars)
            MultiByteToWideChar 65001, 0, b(0), BLen - 3, StrPtr(ReadUnicodeOrANSI), WChars
          End If
        Else 'not an UTF8-BOM, so read the whole Text as ANSI
        feedback = 3
        
          ReadUnicodeOrANSI = StrConv(space$(BLen), vbFromUnicode)
          Get FNr, 1, ReadUnicodeOrANSI
        End If
        
      Case Else 'no BOM was detected, so read the whole Text as ANSI
        feedback = 3
       mLof = LOF(FNr)
       Dim buf() As Byte
       If mLof > 1000 Then
       ReDim buf(1000)
       Else
       ReDim buf(mLof)
       End If
       Get FNr, 1, buf()
       Seek FNr, 1
       Dim notok As Boolean
      If ContainsUTF8(buf()) Then 'maybe is utf-8
      feedback = 2
      nobom = -1
        ReDim b(0 To BLen - 1): Get FNr, 1, b
            WChars = MultiByteToWideChar(65001, 0, b(0), BLen, 0, 0)
            ReadUnicodeOrANSI = space$(WChars)
            MultiByteToWideChar 65001, 0, b(0), BLen, StrPtr(ReadUnicodeOrANSI), WChars
        Else
        notok = True
        
        
            Select Case ContainsUTF16(buf())
        Case 1
            nobom = -1
            BOM = &HFEFF
            ReDim b(0 To BLen - 1): Get FNr, 1, b 'read the Bytes
            GoTo utf16conthere
        Case 2
            nobom = -1
            BOM = &HFEFF
            ReDim b(0 To BLen - 1): Get FNr, 1, b 'read the Bytes
            GoTo utf16conthere
        End Select
        End If
        If notok Then
        ReDim b(0 To BLen - 1): Get FNr, 1, b
        If BLen Mod 2 = 1 Then
        ReadUnicodeOrANSI = StrConv(space$(BLen), vbFromUnicode)
        Else
        ReadUnicodeOrANSI = space$(BLen \ 2)
        End If
         CopyMemory ByVal StrPtr(ReadUnicodeOrANSI), b(0), BLen
         
         cLid = FoundLocaleId(Left$(ReadUnicodeOrANSI, 500))
         
         
         
        ReadUnicodeOrANSI = StrConv(ReadUnicodeOrANSI, vbUnicode, cLid)
        'End If
        End If
    End Select
    
    If InStr(ReadUnicodeOrANSI, vbCrLf) = 0 Then
      If InStr(ReadUnicodeOrANSI, vbLf) Then
      feedback = feedback + 10
   If EnsureWinLFs Then ReadUnicodeOrANSI = Replace(ReadUnicodeOrANSI, vbLf, vbCrLf)
      ElseIf InStr(ReadUnicodeOrANSI, vbCr) Then
      feedback = feedback + 20
      
    If EnsureWinLFs Then ReadUnicodeOrANSI = Replace(ReadUnicodeOrANSI, vbCr, vbCrLf)
      End If
    End If
    feedback = nobom * feedback
ErrHandler:
If FNr Then Close FNr
If Err Then
'MyEr Err.Description, Err.Description
Err.Raise Err.Number, Err.Source & ".ReadUnicodeOrANSI", Err.Description
End If
End Function

Public Function SaveUnicode(ByVal FileName As String, ByVal buf As String, mode2save As Long, Optional Append As Boolean = False) As Boolean
' using doc as extension you can read it from word...with automatic conversion to unicode
' OVERWRITE ALWAYS
Dim w As Long, a() As Byte, F$, i As Long, bb As Byte, yesswap As Boolean
On Error GoTo t12345
If Not Append Then
If Not NeoUnicodeFile(FileName) Then Exit Function
Else
If Not CanKillFile(FileName$) Then Exit Function
End If
F$ = GetDosPath(FileName)
If Err.Number > 0 Or F$ = vbNullString Then Exit Function
w = FreeFile
MyDoEvents
Open F$ For Binary As w
' mode2save
' 0 is utf-le
If Append Then Seek #w, LOF(w) + 1
mode2save = mode2save Mod 10
If mode2save = 0 Then
a() = ChrW(&HFEFF)
Put #w, , a()

ElseIf mode2save = 1 Then
a() = ChrW(&HFFFE) ' big endian...need swap
If Not Append Then Put #w, , a()
yesswap = True
ElseIf Abs(mode2save) = 2 Then  'utf8
If mode2save > 0 And Not Append Then

        Put #w, , CByte(&HEF)
        Put #w, , CByte(&HBB)
        Put #w, , CByte(&HBF)
        End If
        Put #w, , Utf16toUtf8(buf)
        Close w
    SaveUnicode = True
        Exit Function
ElseIf mode2save = 3 Then ' ascii
Dim buf1() As Byte
buf1 = StrConv(buf, vbFromUnicode, cLid)
Put #w, , buf1()
      Close w
    SaveUnicode = True
        Exit Function
End If

Dim maxmw As Long, iPos As Long
iPos = 1
maxmw = 32000 ' check it with maxmw 20 OR 1
If yesswap Then
For iPos = 1 To Len(buf) Step maxmw
a() = Mid$(buf, iPos, maxmw)
For i = 0 To UBound(a()) - 1 Step 2
bb = a(i): a(i) = a(i + 1): a(i + 1) = bb
Next i
Put #w, 3, a()
Next iPos
Else
For iPos = 1 To Len(buf) Step maxmw
a() = Mid$(buf, iPos, maxmw)
Put #w, , a()
Next iPos
End If
Close w
SaveUnicode = True
t12345:
End Function
Public Sub getUniString(F As Long, s As String)
Dim a() As Byte
a() = s
Get #F, , a()
s = a()
End Sub
Public Function getUniStringNoUTF8(F As Long, s As String) As Boolean
Dim a() As Byte
a() = s
Get #F, , a()
If UBound(a) > 4 Then If Not ContainsUTF16(a(), 256) = 1 Then MyEr "No UTF16LE", "Δεν βρήκα UTF16LE": Exit Function
s = a()
getUniStringNoUTF8 = True
End Function
Public Sub putUniString(F As Long, s As String)
Dim a() As Byte
a() = s

Put #F, , a()
End Sub
Public Sub putANSIString(F As Long, s As String)
Dim a() As Byte
a() = StrConv(s, vbFromUnicode, cLid)

Put #F, , a()
End Sub
Public Function getUniStringlINE(F As Long, s As String) As Boolean
' 2 bytes a time... stop to line end and advance to next line

Dim a() As Byte, s1 As String, ss As Long, lbreak As String
a = " "
On Error GoTo a11
Do While Not (LOF(F) < Seek(F))
Get #F, , a()

s1 = a()
If s1 <> vbCr And s1 <> vbLf Then
s = s + s1
'If Asc(s1) = 63 And (AscW(a()) <> 63 And AscW(a()) <> -257) Then
'If AscW(a()) < &H4000 Then Exit Function
''End If
Else
If Not (LOF(F) < Seek(F)) Then
ss = Seek(F)
lbreak = s1
Get #F, , a()
s1 = a()
If s1 <> vbCr And s1 <> vbLf Or lbreak = s1 Then
Seek #F, ss  ' restore it
End If
End If
Exit Do
End If
Loop
getUniStringlINE = True
a11:
End Function

Public Sub getAnsiStringlINE(F As Long, s As String)
' 2 bytes a time... stop to line end and advance to next line
Dim a As Byte, s1 As String, ss As Long, lbreak As String
'a = " "
On Error GoTo a11
Do While Not (LOF(F) < Seek(F))
Get #F, , a

s1 = ChrW(AscW(ChrW(AscW(StrConv(ChrW(a), vbUnicode, cLid)))))
If s1 <> vbCr And s1 <> vbLf Then
s = s + s1
Else
If Not (LOF(F) < Seek(F)) Then
ss = Seek(F)
Get #F, , a
lbreak = s1
s1 = ChrW(AscW(ChrW(AscW(StrConv(ChrW(a), vbUnicode, cLid)))))

If s1 <> vbCr And s1 <> vbLf Or lbreak = s1 Then
Seek #F, ss  ' restore it
End If
End If
Exit Do
End If
Loop
'S = StrConv(S, vbUnicode)
a11:
End Sub
Public Sub getUniStringComma(F As Long, s As String, Optional nochar34 As Boolean)
' sring must be in quotes
' 2 bytes a time... stop to line end and advance to next line
' use numbers with . as decimal not ,
Dim a() As Byte, s1 As String, ss As Long, inside As Boolean
s = vbNullString

a = " "
On Error GoTo a1115

Do While Not (LOF(F) < Seek(F))
    Get #F, , a()
    s1 = a()
    If s1 <> " " Then
    If nochar34 Then s = s1: Exit Do
    If s1 = """" Then inside = True: Exit Do
    End If
Loop
' we throw the first
If Not nochar34 Then If s1 <> """" Then Exit Sub

Do While Not (LOF(F) < Seek(F))
    Get #F, , a()
    
    s1 = a()
    If s1 <> vbCr And s1 <> vbLf And nochar34 And Not s1 = inpcsvsep$ Then
        s = s + s1
    ElseIf s1 <> vbCr And s1 <> vbLf And s1 <> """" And Not nochar34 Then
        s = s + s1
    Else
        If nochar34 Then
        GoTo there
        ElseIf s1 = """" Then
            If s = vbNullString Then ' is the first we have empty string
                inside = False
            Else
            ' look if we have one  more
                If Not (LOF(F) < Seek(F)) Then
                    ss = Seek(F)
                    Get #F, , a()
                    If a(0) = 34 Then
                        s = s + Chr(34)
                        GoTo nn1
                    Else
                        Seek #F, ss
                    End If
                End If
            End If
            inside = False
            Do While Not (LOF(F) < Seek(F))
            Get #F, , a()
            s1 = a()
            
            If s1 = vbCr Or s1 = vbLf Or s1 = inpcsvsep$ Then Exit Do
            Loop
there:
            If s1 = inpcsvsep$ Then Exit Do
        End If
        If s1 <> inpcsvsep$ And (Not (LOF(F) < Seek(F))) And (Not inside) Then
            ss = Seek(F)
            Get #F, , a()
            s1 = a()
            If s1 <> vbCr And s1 <> vbLf Then Seek #F, ss             ' restore it
        End If
        If Not inside Then Exit Do Else s = s + s1
    End If
nn1:
Loop
a1115:
End Sub
Public Sub getAnsiStringComma(F As Long, s As String, Optional nochar34 As Boolean)
' sring must be in quotes
' 2 bytes a time... stop to line end and advance to next line
' use numbers with . as decimal not ,
Dim a As Byte, s1 As String, ss As Long, inside As Boolean
s = vbNullString

On Error GoTo a1111

Do While Not (LOF(F) < Seek(F))
Get #F, , a
s1 = ChrW(AscW(StrConv(ChrW(a), vbUnicode, cLid)))
If s1 <> " " Then
If nochar34 Then s = s1: Exit Do
If s1 = """" Then inside = True: Exit Do

End If
Loop
' we throw the first
If Not nochar34 Then If s1 <> """" Then Exit Sub

Do While Not (LOF(F) < Seek(F))
Get #F, , a

s1 = ChrW(AscW(StrConv(ChrW(a), vbUnicode, cLid)))
If s1 <> vbCr And s1 <> vbLf And nochar34 And Not s1 = inpcsvsep$ Then
    s = s + s1
ElseIf s1 <> vbCr And s1 <> vbLf And s1 <> """" And Not nochar34 Then
    s = s + s1
Else
If nochar34 Then
        GoTo there
        ElseIf s1 = """" Then
If s = vbNullString Then ' is the first we have empty string
inside = False
Else
' look if we have one  more
If Not (LOF(F) < Seek(F)) Then
ss = Seek(F)

Get #F, , a
If a = 34 Then
s = s + Chr(34)
GoTo nn1
Else
Seek #F, ss
End If
End If

End If
inside = False
Do While Not (LOF(F) < Seek(F))
Get #F, , a
s1 = ChrW(AscW(StrConv(ChrW(a), vbUnicode, cLid)))

If s1 = vbCr Or s1 = vbLf Or s1 = inpcsvsep$ Then Exit Do

Loop
there:
If s1 = inpcsvsep$ Then Exit Do
End If
If s1 <> inpcsvsep$ And (Not (LOF(F) < Seek(F))) And (Not inside) Then
    ss = Seek(F)
    Get #F, , a
    s1 = ChrW(AscW(StrConv(ChrW(a), vbUnicode, cLid)))
    If s1 <> vbCr And s1 <> vbLf Then
    Seek #F, ss  ' restore it
    End If
    End If
If Not inside Then Exit Do Else s = s + s1

End If
nn1:
Loop

a1111:
End Sub
Public Sub getUniRealComma(F As Long, s$)
' 2 bytes a time... stop to line end and advance to next line
' use numbers with . as decimal not ,
Dim a() As Byte, s1 As String, ss As Long
s$ = vbNullString
a = " "
On Error GoTo a111
Do While Not LOF(F) < Seek(F)
Get #F, , a()

s1 = a()
If s1 <> vbCr And s1 <> vbLf And s1 <> inpcsvsep$ Then
s = s + s1
Else
If s1 <> inpcsvsep$ And Not (LOF(F) < Seek(F)) Then
    ss = Seek(F)
    Get #F, , a()
    s1 = a()
    If s1 <> vbCr And s1 <> vbLf Then
    Seek #F, ss  ' restore it
    End If
End If
Exit Do
End If
Loop
s$ = MyTrim$(s$)
If LenB(s$) = 0 Then s$ = "0"
a111:


End Sub
Public Sub getAnsiRealComma(F As Long, s$)
' 2 bytes a time... stop to line end and advance to next line
' use numbers with . as decimal not ,
Dim a As Byte, s1 As String, ss As Long
s$ = vbNullString


On Error GoTo a112
Do While Not LOF(F) < Seek(F)
Get #F, , a

s1 = ChrW(AscW(StrConv(ChrW(a), vbUnicode, cLid)))
If s1 <> vbCr And s1 <> vbLf And s1 <> inpcsvsep$ Then
s = s + s1
Else
If s1 <> inpcsvsep$ And Not (LOF(F) < Seek(F)) Then
    ss = Seek(F)
    Get #F, , a
    s1 = ChrW(AscW(StrConv(ChrW(a), vbUnicode, cLid)))
    If s1 <> vbCr And s1 <> vbLf Then
    Seek #F, ss  ' restore it
    End If
End If
Exit Do
End If
Loop
s$ = MyTrim$(s$)
If LenB(s$) = 0 Then s$ = "0"
a112:


End Sub
Public Function RealLenOLD(s$, Optional checkone As Boolean = False) As Long
Dim a() As Byte, ctype As Long, s1$, i As Long, LL As Long, ii As Long
If IsWine Then
RealLenOLD = Len(s$)
Else
ctype = CT_CTYPE3
LL = Len(s$)
   If LL Then
      ReDim a(Len(s$) * 2 + 20)
      If GetStringTypeExW(&HB, ctype, StrPtr(s$), Len(s$), a(0)) <> 0 Then
      ii = 0
      For i = 1 To Len(s$) * 2 - 1 Step 2
      ii = ii + 1
      If a(i - 1) > 0 Then
      If a(i) = 0 Then
      If ii > 1 Then If a(i - 1) < 8 Then LL = LL - 1
      End If
      ElseIf a(i) = 0 Then
      LL = LL - 1
      End If
      
          Next i
      End If
   End If
RealLenOLD = LL
End If
End Function
Public Function RealLen(s$, Optional checkone As Boolean = False) As Long
Dim a() As Byte, a1() As Byte, s1$, i As Long, LL As Long, ii As Long, l$, LLL$
LL = Len(s$)
   If LL Then
      ReDim a(Len(s$) * 2 + 20), a1(Len(s$) * 2 + 20)
         If GetStringTypeExW(&HB, 1, StrPtr(s$), Len(s$), a(0)) <> 0 And GetStringTypeExW(&HB, 4, StrPtr(s$), Len(s$), a1(0)) <> 0 Then
         
        ii = 0
      For i = 1 To Len(s$) * 2 - 1 Step 2
        ii = ii + 1
        If a(i - 1) = 0 Then
        If a(i) = 2 And a1(2) < 8 Then
        
                 If ii > 1 Then
                    s1$ = Mid$(s$, ii, 1)
                    
                    If (AscW(s1$) And &HFFFF0000) = &HFFFF0000 Then
                    Else
                    If l$ = s1$ Then
                        If LLL$ = vbNullString Then LL = LL + 1
                        LLL$ = l$
                    Else
                        l$ = Mid$(s$, ii, 1)
                        LL = LL - 1
                    End If
                    End If
                 Else
                 If checkone Then LL = LL - 1
                 End If
            
        Else
        LLL$ = vbNullString
        End If
       
        
        End If
           l$ = Mid$(s$, ii, 1)
          Next i
      End If
   End If
RealLen = LL
End Function
Public Function PopOne(s$) As String
Dim a() As Byte, ctype As Long, s1$, i As Long, LL As Long, mm As Long
ctype = CT_CTYPE3
Dim one As Boolean
LL = Len(s$)
mm = LL
   If LL Then
      ReDim a(Len(s$) * 2 + 20)
      If GetStringTypeExW(&HB, ctype, StrPtr(s$), Len(s$), a(0)) <> 0 Then
      For i = 1 To Len(s$) * 2 - 1 Step 2
      If a(i - 1) > 0 Then
            If a(i) = 0 Then
            
            If a(i - 1) < 8 Then LL = LL - 1
            Else
            If Not one Then Exit For
            
            End If
            Else
            If one Then Exit For
            one = Not one
            End If
      Next i
      End If
        LL = LL - 1
      mm = mm - LL
   End If
If LL < 0 Then
PopOne = s$
s$ = vbNullString
ElseIf mm > 0 Then
    PopOne = Left$(s$, mm)
    s$ = Right$(s$, LL)
End If

End Function
Public Sub ExcludeOne(s$)
Dim a() As Byte, ctype As Long, s1$, i As Long, LL As Long
LL = Len(s$)
ctype = CT_CTYPE3
   If LL > 1 Then
      ReDim a(Len(s$) * 2 + 20)
      If GetStringTypeExW(&HB, ctype, StrPtr(s$), -1, a(0)) <> 0 Then
      For i = LL * 2 - 1 To 1 Step -2
      If a(i) = 0 Then
      If a(i - 1) > 0 Then
      If a(i - 1) < 8 Then LL = LL - 1
      Else
      Exit For
      End If
      Else
      Exit For
      End If
          Next i
      End If
       LL = LL - 1
       If LL <= 0 Then
       s$ = vbNullString
       Else
       
        s$ = Left$(s$, LL)
        End If
      Else
      s$ = vbNullString
      
   End If
End Sub
Function Tcase(s$) As String
Dim a() As String, i As Long
If s$ = vbNullString Then Exit Function
a() = Split(s$, " ")
For i = 0 To UBound(a())
a(i) = myUcase(Left$(a(i), 1), True) + Mid$(myLcase(a(i)), 2)
Next i
If UBound(a()) > 0 Then
Tcase = Join(a(), " ")
Else
Tcase = a(0)
End If
End Function
Public Sub choosenext()
Dim catchit As Boolean
On Error Resume Next
If Not Screen.ActiveForm Is Nothing Then

    Dim X As Form
     For Each X In Forms
     If X.name = "Form1" Or X.name = "GuiM2000" Or X.name = "Form2" Or X.name = "Form4" Then
         If X.Visible And X.enabled Then
             If catchit Then X.SetFocus: Exit Sub
             If X.hWnd = GetForegroundWindow Then
             catchit = True
             End If
         End If
    End If
         
     Next X
     Set X = Nothing
     For Each X In Forms
     If X.name = "Form1" Or X.name = "GuiM2000" Or X.name = "Form2" Or X.name = "Form4" Then
         If X.Visible And X.enabled Then X.SetFocus: Exit Sub
             
             
         End If
     Next X
     Set X = Nothing
    End If

End Sub
Public Function CheckIsmArray(obj As Object) As Boolean
Dim oldobj As Object
If obj Is Nothing Then Exit Function
Set oldobj = obj

Dim kk As Long
again:
If kk > 20 Then Set obj = oldobj: Exit Function
If TypeOf obj Is mHandler Then
    If obj.t1 = 3 Then
        If obj.indirect >= 0 And obj.indirect <= var2used Then
                Set obj = var(obj.indirect)
                kk = kk + 1
                GoTo again
        Else
                Set obj = obj.objref
        End If

    End If
    
End If
If Not obj Is Nothing Then
If TypeOf obj Is mArray Then If obj.Arr Then CheckIsmArray = True: Set oldobj = Nothing: Exit Function
End If
Set obj = oldobj
End Function
Public Function CheckIsmArrayOrStackOrCollection(obj As Object) As Boolean
Dim oldobj As Object
If obj Is Nothing Then Exit Function
Set oldobj = obj

Dim kk As Long
again:
If kk > 20 Then Set obj = oldobj: Exit Function
If TypeOf obj Is mHandler Then
    If obj.t1 <> 2 Then
        If obj.indirect >= 0 And obj.indirect <= var2used Then
                Set obj = var(obj.indirect)
                kk = kk + 1
                GoTo again
        Else
                Set obj = obj.objref
        End If
   
    End If
    
End If
If Not obj Is Nothing Then
If TypeOf obj Is mArray Then If obj.Arr Then CheckIsmArrayOrStackOrCollection = True: Set oldobj = Nothing: Exit Function
If TypeOf obj Is mStiva Then CheckIsmArrayOrStackOrCollection = True: Set oldobj = Nothing: Exit Function
If TypeOf obj Is FastCollection Then CheckIsmArrayOrStackOrCollection = True: Set oldobj = Nothing: Exit Function
End If
Set obj = oldobj
End Function
Public Function CheckDeepAny(obj As Object) As Boolean
Dim oldobj As Object
If obj Is Nothing Then Exit Function
Set oldobj = obj

Dim kk As Long
again:
If kk > 20 Then Set obj = oldobj: Exit Function
If TypeOf obj Is mHandler Then
    If obj.t1 = 3 Then
        If obj.indirect >= 0 And obj.indirect <= var2used Then
                Set obj = var(obj.indirect)
                kk = kk + 1
                GoTo again
        Else
                Set obj = obj.objref
        End If

    End If
    
End If
If Not obj Is Nothing Then Set oldobj = Nothing: CheckDeepAny = True: Exit Function
Set obj = oldobj
End Function
Public Function CheckLastHandler(obj As Object) As Boolean
Dim oldobj As Object, first As Object
If obj Is Nothing Then Exit Function
Set first = obj

Dim kk As Long
again:
If kk > 20 Then Set obj = first: Exit Function
If TypeOf obj Is mHandler Then
    'If obj.t1 = 3 Then
        If obj.indirect >= 0 And obj.indirect <= var2used Then
                Set oldobj = obj
                Set obj = var(obj.indirect)
                kk = kk + 1
                GoTo again
        Else
                kk = kk + 1
                Set oldobj = obj
                Set obj = obj.objref
                GoTo again
        End If

    'End If
    
End If
If Not oldobj Is Nothing Then Set obj = oldobj: Set oldobj = Nothing: CheckLastHandler = True: Exit Function
Set obj = first
End Function
Public Function CheckLastHandlerVariant(obj) As Boolean
Dim oldobj As Object, first As Object
If obj Is Nothing Then Exit Function
Set first = obj

Dim kk As Long
again:
If kk > 20 Then Set obj = first: Exit Function
If obj Is Nothing Then Exit Function
If TypeOf obj Is mHandler Then
    'If obj.t1 = 3 Then
        If obj.indirect >= 0 And obj.indirect <= var2used Then
                Set oldobj = obj
                Set obj = var(obj.indirect)
                kk = kk + 1
                GoTo again
        Else
                kk = kk + 1
                Set oldobj = obj
                Set obj = obj.objref
                GoTo again
        End If

    'End If
    
End If
If Not oldobj Is Nothing Then Set obj = oldobj: Set oldobj = Nothing: CheckLastHandlerVariant = True: Exit Function
Set obj = first
End Function
Public Function CheckLastHandlerOrIterator(obj As Object, lastindex As Long) As Boolean
Dim oldobj As Object, first As Object
If obj Is Nothing Then Exit Function
Set first = obj
lastindex = -1
Dim kk As Long
again:
If kk > 20 Then Set obj = first: Exit Function
If TypeOf obj Is mHandler Then
        If obj.UseIterator Then lastindex = obj.index_cursor
        If obj.indirect >= 0 And obj.indirect <= var2used Then
                Set oldobj = obj
                Set obj = var(obj.indirect)
                kk = kk + 1
                GoTo again
        Else
                kk = kk + 1
                Set oldobj = obj
                Set obj = obj.objref
                GoTo again
        End If

End If
    

If Not oldobj Is Nothing Then Set obj = oldobj: Set oldobj = Nothing: CheckLastHandlerOrIterator = True: Exit Function
Set obj = first
End Function
Public Function IfierVal()
If LastErNum <> 0 Then LastErNum = 0: IfierVal = True
End Function
Public Sub OutOfLimit()
  MyEr "Out of limit", "Εκτός ορίου"
End Sub
Public Sub stackproblem()
MyEr "Problem in return stack", "Πρόβλημα στον σωρό επιστροφής"
End Sub
Public Sub PlaceAcommaBefore()
MyEr "Place a comma before", "Βάλε ένα κόμμα πριν"
End Sub
Public Sub unknownid(b$, w$)
MyErMacro b$, "unknown identifier " & w$, "’γνωστο αναγνωριστικό " & w$
End Sub
Public Sub MissCdib()
  MyEr "Missing IMAGE", "Λείπει εικόνα"
End Sub
Public Sub MissFile()
 MyEr "File not found", "Δεν βρέθηκε ο αρχείο"
End Sub
Public Sub BadObjectDecl()
  MyEr "Bad object declaration - use Clear Command for Gui Elements", "Λάθος όρισμα αντικειμένου - χρησιμοποίησε Καθαρό για να καθαρίσεις τυχόν στοιχεία του γραφικού περιβάλλοντος"
End Sub
Public Sub NoEnumaretor()
  MyEr " - No enumarator found for this object", " - Δεν βρήκα δρομέα συλλογής για αυτό το αντικείμενο"
End Sub
Public Sub AssigntoNothing()
  MyEr "Bad object declaration - use Declare command", "Λάθος όρισμα αντικειμένου - χρησιμοποίησε την Όρισε"
End Sub
Public Sub Overflow()
 MyEr "Overflow", "υπερχείλιση"
End Sub
Public Sub MissCdibStr()
  MyEr "Missing IMAGE in string", "Λείπει εικόνα στο αλφαριθμητικό"
End Sub
Public Sub MissStackStr()
  MyEr "Missing string value from stack", "Λείπει αλφαριθμητικό από το σωρό"
End Sub
Public Sub WrongFileHandler()
MyEr "Wrong File Handler", "Λάθος Χειριστής Αρχείου"
End Sub

Public Sub MissStackItem()
 MyEr "Missing item from stack", "Λείπει κάτι από το σωρό"
End Sub
Public Sub MissStackNumber()
 MyEr "Missing number value from stack", "Λείπει αριθμός από το σωρό"
End Sub
Public Sub missNumber()
MyEr "Only number allowed", "Μόνο αριθμός επιτρέπεται"
End Sub
Public Sub MissNumExpr()
MyEr "Missing number expression", "Λείπει αριθμητική παράσταση"
End Sub
Public Sub MissLicence()
MyEr "Missing Licence", "Λείπει ’δεια"
End Sub
Public Sub MissStringExpr()
MyEr "Missing string expression", "Λείπει αλφαριθμητική παράσταση"
End Sub
Public Sub MissString()
MyEr "Missing string", "Λείπει αλφαριθμητικό"
End Sub
Public Sub MissStringNumber()
MyEr "Missing string or number", "Λείπει αλφαριθμητικό ή αριθμός"
End Sub

Public Sub NoCreateFile()
    MyEr "Can't create file", "Δεν μπορώ να φτιάξω αρχείο"
End Sub
Public Sub BadFilename()
MyEr "Bad filename", "Λάθος στο όνομα αρχείου"
End Sub
Public Sub ReadOnly()
MyEr "Read Only", "Μόνο για ανάγνωση"
End Sub
Public Sub MissDir()
MyEr "Missing directory name", "Λείπει όνομα φακέλου"
End Sub
Public Sub MissType()
MyEr "Wrong data type", "’λλος τύπος μεταβλητής"
End Sub

Public Sub BadPath()
MyEr "Bad Path name", "Λάθος στο όνομα φακέλου (τόπο)"
End Sub
Public Sub BadReBound()
MyEr "Can't commit a reference here", "Δεν μπορώ να αναθέσω εδώ μια αναφορά"
End Sub
Public Sub oxiforPrinter()
MyEr "Not allowed this command for printer", "Δεν επιτρέπεται αυτή η εντολή για τον εκτυπωτή"
End Sub
Public Sub ResourceLimit()
MyEr "No more Graphic Resource for forms - 100 Max", "Δεν έχω άλλο χώρο για γραφικά σε φόρμες - 100 Μεγιστο"
End Sub
Public Sub oxiforforms()
MyEr "Not allowed this command for forms", "Δεν επιτρέπεται αυτή η εντολή για φόρμες"
End Sub
Public Sub SyntaxError()
If LastErName = vbNullString Then
MyEr "Syntax Error", "Συντακτικό Λάθος"
Else
If LastErNum = 0 Then LastErNum = -1 ' general
LastErNum1 = LastErNum
End If
End Sub
Public Sub MissingnumVar()
MyEr "missing numeric variable", "λείπει αριθμητική μεταβλητή"
End Sub
Public Sub BadGraphic()
MyEr "Can't operate graphic", "δεν μπορώ να χειριστώ το γραφικό"
End Sub
Public Sub SelectorInUse()
MyEr "File/Folder Selector in Use", "Η φόρμα επιλογής αρχείων/φακέλων είναι σε χρήση"
End Sub
Public Sub MissingDoc()  ' this is for identifier or execute part
MyEr "missing document type variable", "λείπει μεταβλητή τύπου εγγράφου"
End Sub
Public Sub MissingLabel()
MyEr "Missing label/Number line", "Λείπει Ετικέτα/Αριθμός γραμμής"
End Sub
Public Sub MissFuncParammeterdOCVar(ar$)
MyEr "Not a Document variable " + ar$, "Δεν είναι μεταβλητή τύπου εγγράφου " + ar$
End Sub
Public Sub MissingBlock()  ' this is for identifier or execute part
MyEr "missing block {} or string expression", "λείπει κώδικας σε {} η αλφαριθμητική έκφραση"
End Sub
Public Sub MissingBlockCode()
MyEr "missing block {}", "λείπει κώδικας σε μπλοκ {}"
End Sub
Public Sub OnlyOneLineAllowed()
MyEr "Use block {} in starting line only", "Χρησιμοποίησε μπλοκ {} στην αρχική γραμμή"
End Sub
Public Function CheckBlock(once As Boolean) As Long
                                    If once Then
                                        OnlyOneLineAllowed
                                    Else
                                        MissingBlockCode
                                    End If
End Function

Public Sub MissingEnumBlock()
MyEr "missing block {} for enumeration constants", "λείπει μπλοκ {} για σταθερές απαρίθμησης "
End Sub
Public Sub MissingCodeBlock()
MyEr "missing block {}", "λείπει μπλοκ κώδικα σε {}"
End Sub
Public Sub MissingArray(w$)
MyEr "Can't find array " & w$ & ")", "Δεν βρίσκω πίνακα " & w$ & ")"
End Sub
Public Sub ErrNum()
MyEr "Error in number", "Λάθος στον αριθμό"
End Sub
Public Sub CantAssignValue()
MyEr "Can't assign value to constant", "Δεν μπορώ να βάλω τιμή σε σταθερά"
End Sub
Public Sub ExpectedEnumType()
 MyEr "Expected Enumaration Type", "Περίμενα τύπο απαρίθμησης"
End Sub

Public Sub ExpectedVariable()
 MyEr "Expected variable", "Περίμενα μεταβλητή"
End Sub
Public Sub Expected(w1$, w2$)
 MyEr "Expected object type " + w1$, "Περίμενα αντικείμενο τύπου " + w2$
End Sub
Public Sub ExpectedCaseorElseorEnd2()
MyEr "Expected Case or Else or End Select", "Περίμενα Με ή Αλλιώς ή Τέλος Επιλογής"
End Sub
Public Sub ExpectedCaseorElseorEnd()
 MyEr "Expected Case or Else or End Select, for two or more commands use {}", "Περίμενα Με ή Αλλιώς ή Τέλος Επιλογής, για δυο ή περισσότερες εντολές χρησιμοποίησε { }"
End Sub
Public Sub ExpectedCommentsOnly()
 MyEr "Expected comments (using ' or \) or new line", "Περίμενα σημειώσεις (με ' ή \) ή αλλαγή γραμής"
End Sub

Public Sub ExpectedEndSelect()
 MyEr "Expected Εnd Select", "Περίμενα Τέλος Επιλογής"
End Sub
Public Sub ExpectedEndSelect2()
 MyEr "Expected Εnd Select, for two or more commands use {}", "Περίμενα Τέλος Επιλογής, για δυο ή περισσότερες εντολές χρησιμοποίησε { }"
End Sub
Public Sub LocalAndGlobal()
MyEr "Global and local together;", "Γενική και τοπική μαζί!"
End Sub
Public Sub UnknownProperty(w$)
MyEr "Unknown Property " & w$, "’γνωστη ιδιότητα " & w$
End Sub
Public Sub UnknownVariable(v$)
Dim i As Long
i = rinstr(v$, "." + ChrW(8191))
If i > 0 Then
    i = rinstr(v$, ".")
    MyEr "Unknown Variable " & Mid$(v$, i), "’γνωστη μεταβλητή " & Mid$(v$, i)
Else
    i = rinstr(v$, "].")
    If i > 0 Then
        MyEr "Unknown Variable " & Mid$(v$, i + 2), "’γνωστη μεταβλητή " & Mid$(v$, i + 2)
    Else
        i = rinstr(v$, ChrW(8191))
    If i > 0 Then
        i = InStr(i + 1, v$, ".")
        If i > 0 Then
            MyEr "Unknown Variable " & Mid$(v$, i + 1), "’γνωστη μεταβλητή " & Mid$(v$, i + 1)
        Else
            MyEr "Unknown Variable", "’γνωστη μεταβλητή"
        End If
    Else
        MyEr "Unknown Variable " & v$, "’γνωστη μεταβλητή " & v$
    End If
    End If
End If
End Sub
Sub indexout(a$)
MyErMacro a$, "Index out of limits", "Δείκτης εκτός ορίων"
End Sub

Sub wrongfilenumber(a$)
 MyErMacro a$, "not valid file number", "λάθος αριθμός αρχείου"
End Sub
Public Sub WrongArgument(a$)
MyErMacro a$, Err.Description, "Λάθος όρισμα"
End Sub
Public Sub UnKnownWeak(w$)
 MyEr "Unknown Weak " & w$, "’γνωστη ισχνή " & w$
End Sub
Public Sub InternalEror()
MyEr "Internal error", "Εσωτερικό λάθος"
End Sub
Sub NegativeIindex(a$)
MyErMacro a$, "negative index", "αρνητικός δείκτη"
End Sub
Sub joypader(a$, r)
MyErMacro a$, "Joypad number " & CStr(r) & " isn't ready", "Το νούμερο Λαβής " & CStr(r) & " δεν είναι έτοιμο"
End Sub
Sub noImage(a$)
MyErMacro a$, "Νο image in string", "Δεν υπάρχει εικόνα στο αλφαριθμητικό"
End Sub
Sub noImageInBuffer(a$)
MyErMacro a$, "No Image in Buffer", "Δεν έχει εικόνα η Διάρθρωση"
End Sub

Sub WrongJoypadNumber(a$)
MyErMacro a$, "Joypad number 0 to 15", "Αριθμός Λαβής από 0 έως 15"
End Sub
Sub CantFindArray(a$, s$)
MyErMacro a$, "Can't find array " & s$, "Δεν βρίσκω πίνακα " & s$
End Sub
Sub CantReadDimension(a$, s$)
 MyErMacro a$, "Can't read dimension index from array " & s$, "Δεν μπορώ να διαβάσω τον δείκτη διάστασης του πίνακα " & s$

End Sub
Sub cantreadlib(a$)
MyErMacro a$, "Can't Read TypeLib", "Δεν μπορώ να διαβάσω τους τύπους των παραμέτρων"
End Sub
Public Sub NotForArray()
MyEr "not for array items", "όχι για στοιχεία πίνακα"
End Sub
Public Sub NotArray()  ' this is for identifier or execute part
MyEr "Expected Array", "Περίμενα πίνακα"
End Sub
Public Sub NotExistArray()  ' this is for identifier or execute part
MyEr "Array not exist", "Δεν υπάρχει τέτοιος πίνακας"
End Sub
Public Sub MissingGroup()  ' this is for identifier or execute part
MyEr "missing group type variable", "λείπει μεταβλητή τύπου ομάδας"
End Sub
Public Sub MissingGroupExp()  ' this is for identifier or execute part
MyEr "missing group type expression", "λείπει έκφραση τύπου ομάδας"
End Sub
Public Sub BadGroupHandle()  ' this is for identifier or execute part
MyEr "group isn't variable", "η ομάδα δεν είναι μεταβλητή"
End Sub
Public Sub MissingDocRef()  ' this is for identifier or execute part
MyEr "invalid document pointer", "μη έγκυρος δείκτης εγγράφου"
End Sub
Public Sub MissingObjReturn()
MyEr "Missing Object", "Δεν βρήκα αντικείμενο"
End Sub
Public Sub NoNewLambda()
    MyEr "No New statement for lambda", "Όχι δήλωση νέου για λαμδα"
End Sub
Public Sub ExpectedObj(nn$)
MyEr "Expected object type " + nn$, "Περίμενα αντικείμενο τύπου " + nn$
End Sub
Public Sub MisOperatror(ss$)
MyEr "Group not support operator " + ss$, "Η ομάδα δεν υποστηρίζει το τελεστή " + ss$
End Sub
Public Sub CantReadFileTimeStap(a$)
MyErMacro a$, "Can't Read File TimeStamp", "Δεν μπορώ να διαβάσω την Χρονοσήμανση του αρχείου"
End Sub

Public Sub ExpectedObjInline(nn$)
MyErMacro nn$, "Expected Object", "Περίμενα αντικείμενο"
End Sub
Public Sub MissingObj()
MyEr "missing object type variable", "λείπει μεταβλητή τύπου αντικειμένου"
End Sub
Public Sub BadGetProp()
MyEr "Can't Get Property", "Δεν μπορώ να διαβάσω αυτή την ιδιότητα"
End Sub
Public Sub BadLetProp()
MyEr "Can't Let Property", "Δεν μπορώ να γράψω αυτή την ιδιότητα"
End Sub
Public Sub NoNumberAssign()
MyEr "Can't assign number to object", "Δεν μπορώ να δώσω αριθμό στο αντικείμενο"
End Sub
Public Sub NoAssignThere()
MyEr "Use Return Object to change items", "Χρησιμοποίησε την Επιστροφή αντικείμενο για να επιστρέψεις τιμές"
End Sub
Public Sub NoObjectpAssignTolong()
MyEr "Can't assign object to long", "Δεν μπορώ να δώσω αντικείμενο στον μακρυ"
End Sub
Public Sub NoObjectpAssignToInteger()
MyEr "Can't assign object to Integer", "Δεν μπορώ να δώσω αντικείμενο στον ακέραιο"
End Sub
Public Sub NoObjectAssign()
MyEr "Can't assign object", "Δεν μπορώ να δώσω αντικείμενο"
End Sub
Public Sub NoNewStatFor(w1$, w2$)
MyEr "No New statement for " + w1$, "Όχι δήλωση νέου για " + w2$
End Sub
Public Sub NoThatOperator(ss$)
    MyEr ss$ + " operator not allowed in group definition", " Ο τελεστής " + ss$ + " δεν επιτρεπεται σε ορισμό ομάδας"
End Sub
Public Sub MissingObjRef()
MyEr "invalid object pointer", "μη έγκυρος δείκτης αντικειμένου"
End Sub
Public Sub MissingStrVar()  ' this is for identifier or execute part
MyEr "missing string variable", "λείπει αλφαριθμητική μεταβλητή"
End Sub
Public Sub NoSwap(nameOfvar$)
MyEr "Can't swap ", "Δεν μπορώ να αλλάξω τιμές "
End Sub
Public Sub Nosuchvariable(nameOfvar$)
MyEr "No such variable " + nameOfvar$, "δεν υπάρχει τέτοια μεταβλητή " + nameOfvar$
End Sub
Public Sub NoValueForVar(w$)
If LastErNum = 0 Then
MyEr "No value for variable " & w$, "Χωρίς τιμή η μεταβλητή " & w$
End If
End Sub
Public Sub NoReference()
   MyEr "No reference exist", "Δεν υπάρχει αναφορά"
End Sub
Public Sub NoCommandOrBlock()
MyEr "Expected in Select Case a Block or a Command", "Περίμενα στην Επίλεξε Με μια εντολή ή ένα μπλοκ εντολών)"
End Sub

Public Sub NoSecReF()
MyEr "No reference allowed - use new variable", "Δεν δέχεται αναφορά - χρησιμοποίησε νέα μεταβλητή"
End Sub
Public Sub MissSymbolMyEr(wht$)   ' not the macro one
MyEr "missing " & wht$, "λείπει " & wht$
End Sub
Public Sub MissTHENELSE()
    MyEr "missing THEN or ELSE", "δεν βρήκα ΤΟΤΕ ή ΑΛΛΙΩΣ"
End Sub

Public Sub MissENDIF()
    MyEr "missing END IF", "δεν βρήκα ΤΕΛΟΣ ΑΝ"
End Sub
Public Sub MissIF()
    MyEr "No IF for END IF", "δεν βρήκα ΑΝ για την ΤΕΛΟΣ ΑΝ"
End Sub

Public Sub BadCommand()
 MyEr "Command for supervisor rights", "Εντολή μόνο για επόπτη"
End Sub
Public Sub NoClauseInThread()
MyEr "can't find ERASE or HOLD or RESTART or INTERVAL clause", "Δεν μπορώ να βρω όρο όπως το ΣΒΗΣΕ ή το ΚΡΑΤΑ ή το ΞΕΚΙΝΑ ή το ΚΑΘΕ"
End Sub
Public Sub NoThisInThread()
MyEr "Clause This can't used outside a thread", "Ο όρος ΑΥΤΟ δεν μπορεί να χρησιμοποιηθεί έξω από ένα νήμα"
End Sub
Public Sub MisInterval()
MyEr "Expected number for interval, miliseconds", "Περίμενα αριθμό για ορισμό τακτικού διαστήματος εκκίνησης νήματος (χρόνο σε χιλιοστά δευτερολέπτου)"
End Sub
Public Sub NoRef2()
MyEr "No with reference in left side of assignment", "Όχι με αναφορά στην εκχώρηση τιμής"
End Sub
Public Sub WrongObject()
MyEr "Wrong object type", "λάθος τύπος αντικειμένου"
End Sub
Public Sub NullObject()
MyEr "object type is Nothing", "O τύπος αντικειμένου είναι Τίποτα"
End Sub
Public Sub WrongType()
MyEr "Wrong type", "λάθος τύπος"
End Sub
Public Sub GroupWrongUse()
MyEr "Something wrong with group", "Κάτι πάει στραβά με την ομάδα"
End Sub
Public Sub GroupCantSetValue()
    MyEr "Group can't set value", "Η ομάδα δεν μπορεί να θέσει τιμή"
End Sub
Public Sub PropCantChange()
MyEr "Property can't change", "Η ιδιότητα δεν μπορεί να αλλάξει"
End Sub
Public Sub NeedAGroupFromExpression()
MyEr "Need a group from expression", "Χρειάζομαι μια ομάδα από την έκφραση"
End Sub
Public Sub NeedAGroupInRightExpression()
MyEr "Need a group from right expression", "Χρειάζομαι μια ομάδα από την δεξιά έκφραση"
End Sub
Public Sub NotAfter(a$)
MyErMacro a$, "not an expression after not operator", "δεν υπάρχει παράσταση δεξιά τού τελεστή όχι"
End Sub
Public Sub EmptyArray()
MyEr "Empty Array", "’δειος Πίνακας"
End Sub
Public Sub EmptyStack(a$)
 MyErMacro a$, "Stack is empty", "O σωρός είναι άδειος"
End Sub
Public Sub StackTopNotArray(a$)
 MyErMacro a$, "Stack top isn't array", "Η κορυφή του σωρού δεν είναι πίνακας"
End Sub

Public Sub StackTopNotGroup(a$)
MyErMacro a$, "Stack top isn't group", "Η κορυφή του σωρού δεν είναι ομάδα"
End Sub
Public Sub StackTopNotNumber(a$)
MyErMacro a$, "Stack top isn't number", "Η κορυφή του σωρού δεν είναι αριθμός"
End Sub
Public Sub NeedAnArray(a$)
MyErMacro a$, "Need an Array", "Χρειάζομαι ένα πίνακα"
End Sub
Public Sub noref()
MyEr "No with reference (&)", "Όχι με αναφορά (&)"
End Sub
Public Sub NoMoreDeep(deep As Variant)
MyEr "No more" + Str(deep) + " levels gosub allowed", "Δεν επιτρέπονται πάνω από" + Str(deep) + " επίπεδα για εντολή ΔΙΑΜΕΣΟΥ"
End Sub
Public Sub CantFind(w$)
MyEr "Can't find " + w$ + " or type name", "Δεν μπορώ να βρω το " + w$ + " ή όνομα τύπου"
End Sub
Public Sub OverflowLong(Optional b As Boolean = False)
If b Then
MyEr "OverFlow Integer", "Yπερχείλιση ακεραίου"
Else
MyEr "OverFlow Long", "Yπερχείλιση μακρύ"
End If
End Sub
Public Sub BadUseofReturn()
MyEr "Wrong Use of Return", "Κακή χρήση της επιστροφής"
End Sub
Public Sub DevZero()
    MyEr "division by zero", "διαίρεση με το μηδέν"
End Sub
Public Sub DevZeroMacro(aa$)
    MyErMacro aa$, "division by zero", "διαίρεση με το μηδέν"
End Sub
Public Sub ErrInExponet(a$)
MyErMacro a$, "Error in exponet", "Λάθος στον εκθέτη"
End Sub

Public Sub LambdaOnly(a$)
MyErMacro a$, "Only in lambda function", "Μόνο σε λάμδα συνάρτηση"
End Sub
Public Sub FilePathNotForUser()
MyEr "Filepath is not valid for user", "Ο τόπος του αρχείου δεν είναι έγκυρος για τον χρήστη"
End Sub

' used to isnumber
Public Sub MyErMacro(wher$, en$, gr$)
If stackshowonly Then
LastErNum = -2
wher$ = " : ERROR -2" & Sput(en$) + Sput(gr$) + wher$
Else
MyEr en$, gr$
End If
End Sub
Public Sub MyErMacroStr(wher$, en$, gr$)
If stackshowonly Then
LastErNum = -2
wher$ = " : ERROR -2" & Sput(en$) + Sput(gr$) + wher$
Else
MyEr en$, gr$
End If
End Sub
Public Sub ZeroParam(ar$)   ' we use MyErMacro in isNumber and isString
MyErMacro ar$, "Empty parameter", "Μηδενική παράμετρος"
End Sub
Public Sub MissPar()
MyEr "missing parameter", "λείπει παράμετρος"
End Sub
Public Sub MissModuleName()
MyEr "Missing module name", "Λείπει όνομα τμήματος"
End Sub
Public Sub nonext()
MyEr "NEXT without FOR", "ΕΠΟΜΕΝΟ χωρίς ΓΙΑ"
End Sub
Public Sub MissWhile()
MyEr "Missing the End While", "Έχασα το Τέλος Ενώ"
End Sub

Public Sub MissUntil()
MyEr "Missing the Until or Always", "Έχασα το Μέχρι ή το Πάντα"
End Sub

Public Sub MissNext()
MyEr "Missing the right NEXT", "Έχασα το σωστό ΕΠΟΜΕΝΟ"
End Sub
Public Sub MissVarName()
MyEr "Missing variable name", "Λείπει όνομα μεταβλητής"
End Sub
Public Sub MissParamref(ar$)
MyErMacro ar$, "missing by reference parameter", "λείπει με αναφορά παράμετρος"
End Sub
Public Sub MissParam(ar$)
MyErMacro ar$, "missing parameter", "λείπει παράμετρος"
End Sub
Public Sub MissFuncParameterStringVar()
MyEr "Not a string variable", "Δεν είναι αλφαριθμητική μεταβλητή"
End Sub
Public Sub MissFuncParameterStringVarMacro(ar$)
MyErMacro ar$, "Not a string variable", "Δεν είναι αλφαριθμητική μεταβλητή"
End Sub
Public Sub NoSuchFolder()
MyEr "No such folder", "Δεν υπάρχει τέτοιος φάκελος"
End Sub
Public Sub MissSymbol(wht$)
MyEr "missing " & wht$, "λείπει " & wht$
End Sub
Public Sub ClearSpace(nm$)
Dim i As Long
Do
    i = 1
    If FastOperator(nm$, vbCrLf, i, 2, False) Then
        SetNextLine nm$
    ElseIf FastOperator(nm$, "\", i) Then
        SetNextLine nm$
    ElseIf FastOperator(nm$, "'", i) Then
        SetNextLine nm$
    Else
    Exit Do
    End If
Loop
End Sub
Public Function StringToEscapeStr(RHS As String, Optional json As Boolean = False) As String
Dim i As Long, cursor As Long, ch As String
cursor = 0
Dim DEL As String
Dim H9F As String
DEL = ChrW(127)
H9F = ChrW(&H9F)
For i = 1 To Len(RHS)
                ch = Mid$(RHS, i, 1)
                cursor = cursor + 1
                Select Case ch
                    Case "\":        ch = "\\"
                   ' Case """":       ch = "\"""
                    Case """"
                    If json Then
                        ch = "\"""
                    Else
                        ch = "\u0022"
                    End If
                    Case vbLf:       ch = "\n"
                    Case vbCr:       ch = "\r"
                    Case vbTab:      ch = "\t"
                    Case vbBack:     ch = "\b"
                    Case vbFormFeed: ch = "\f"
                    Case Is < " ", DEL To H9F
                        ch = "\u" & Right$("000" & Hex$(AscW(ch)), 4)
                End Select
                If cursor + Len(ch) > Len(StringToEscapeStr) Then StringToEscapeStr = StringToEscapeStr + space$(500)
                Mid$(StringToEscapeStr, cursor, Len(ch)) = ch
                cursor = cursor + Len(ch) - 1
Next
If cursor > 0 Then StringToEscapeStr = Left$(StringToEscapeStr, cursor)

End Function
Public Function EscapeStrToString(RHS As String) As String
Dim i As Long, cursor As Long, ch As String
     For cursor = 1 To Len(RHS)
        ch = Mid$(RHS, cursor, 1)
        i = i + 1
        Select Case ch
            Case """": GoTo ok1
            Case "\":
                cursor = cursor + 1
                ch = Mid$(RHS, cursor, 1)
                Select Case LCase$(ch) 'We'll make this forgiving though lowercase is proper.
                    Case "\", "/": ch = ch
                    Case """":      ch = """"
                    Case "a":       ch = Chr$(7)
                    Case "n":      ch = vbLf
                    Case "r":      ch = vbCr
                    Case "t":      ch = vbTab
                    Case "b":      ch = vbBack
                    Case "f":      ch = vbFormFeed
                    Case "u":      ch = ParseHexChar(RHS, cursor, Len(RHS))
                End Select
        End Select
                If i + Len(ch) > Len(EscapeStrToString) Then EscapeStrToString = EscapeStrToString + space$(500)
                Mid$(EscapeStrToString, i, Len(ch)) = ch
                i = i + Len(ch) - 1
    Next
ok1:
    If i > 0 Then EscapeStrToString = Left$(EscapeStrToString, i)
End Function

Private Function ParseHexChar( _
    ByRef Text As String, _
    ByRef cursor As Long, _
    ByVal LenOfText As Long) As String
    
    Const ASCW_OF_ZERO As Long = &H30&
    Dim Length As Long
    Dim ch As String
    Dim DigitValue As Long
    Dim Value As Long

    For cursor = cursor + 1 To LenOfText
        ch = Mid$(Text, cursor, 1)
        Select Case ch
            Case "0" To "9", "A" To "F", "a" To "f"
                Length = Length + 1
                If Length > 4 Then Exit For
                If ch > "9" Then
                    DigitValue = (AscW(ch) And &HF&) + 9
                Else
                    DigitValue = AscW(ch) - ASCW_OF_ZERO
                End If
                Value = Value * &H10& + DigitValue
            Case Else
                Exit For
        End Select
    Next
    If Length = 0 Then Err.Raise 5 'No hex digits at all.
    cursor = cursor - 1
    ParseHexChar = ChrW$(Value)
End Function

Public Function ReplaceSpace(a$) As String
Dim i As Long, j As Long
i = 1
Do
i = InStr(i, a$, "[")
If i > 0 Then
    i = i + 1
    j = InStr(i, a$, "]")
    If j > 0 Then
    j = j - i
    Mid$(a$, i, j) = Replace(Mid$(a$, i, j), " ", ChrW(160))
    i = i + j
    End If
Else
    Exit Do
End If
Loop
ReplaceSpace = a$
End Function
Function GetReturnArray(bstack As basetask, x1 As Long, b$, p As Variant, ss$, pppp As mArray) As Boolean ' true is error

Do
        If IsExp(bstack, b$, p) Then
        If x1 = 0 Then If MaybeIsSymbol(b$, ",") Then x1 = 1: Set pppp = New mArray: pppp.PushDim (1): pppp.PushEnd
        If x1 = 0 Then
                If Len(bstack.OriginalName$) > 3 Then
                        If Mid$(bstack.OriginalName$, Len(bstack.OriginalName$) - 2, 1) = "$" Then
                            MissStringExpr
                            Exit Do
                        End If
                    End If
                 If Right$(bstack.OriginalName$, 3) = "%()" Then p = MyRound(p)
                 Set bstack.FuncObj = bstack.lastobj
                 Set bstack.lastobj = Nothing
                 bstack.FuncValue = p
        Else
                pppp.SerialItem 0, x1 + 1, 9
                If bstack.lastobj Is Nothing Then
                    pppp.item(x1 - 1) = p
                Else
                    Set pppp.item(x1 - 1) = bstack.lastobj
                    Set bstack.lastobj = Nothing
                End If
                bstack.FuncValue = p
                x1 = x1 + 1
                             
        End If
        ElseIf IsStrExp(bstack, b$, ss$) Then
            If x1 = 0 Then If MaybeIsSymbol(b$, ",") Then x1 = 1: Set pppp = New mArray: pppp.PushDim (1): pppp.PushEnd
            If x1 = 0 Then
                If Len(bstack.OriginalName$) > 3 Then
                    If Mid$(bstack.OriginalName$, Len(bstack.OriginalName$) - 2, 1) <> "$" Then
                         MissNumExpr
                         GetReturnArray = True
                         Exit Function
                    End If
                Else
                    MissNumExpr
                    GetReturnArray = True
                    Exit Function
                End If
                Set bstack.FuncObj = bstack.lastobj
                Set bstack.lastobj = Nothing
                bstack.FuncValue = ss$
            Else
                pppp.SerialItem 0, x1 + 1, 9
                If bstack.lastobj Is Nothing Then
                    pppp.item(x1 - 1) = ss$
                Else
                    Set pppp.item(x1 - 1) = bstack.lastobj
                    Set bstack.lastobj = Nothing
                End If
                x1 = x1 + 1
                bstack.FuncValue = ss$
                            
            End If
        End If
        Loop Until Not FastSymbol(b$, ",")
        If x1 > 0 Then
         pppp.SerialItem 0, x1, 9
         Set bstack.FuncObj = pppp
         Set pppp = New mArray
         Set bstack.lastobj = Nothing
         If VarType(bstack.FuncValue) = 5 Then
         bstack.FuncValue = 0
         Else
         bstack.FuncValue = vbNullString
         End If
        End If
        x1 = 0
End Function
Function AssignTypeNumeric2(v, i As Long) As Boolean
If VarType(v) = i Then AssignTypeNumeric2 = True: Exit Function
On Error GoTo there
If VarType(v) = vbString Then v = Format$(v)
Select Case i
Case vbBoolean
v = CBool(v)
Case vbCurrency
v = CCur(v)
Case vbDecimal
v = CDec(v)
Case vbLong
v = CLng(v)
Case vbSingle
v = CSng(v)
Case vbInteger
v = CInt(v)
Case Else
v = CDbl(v)
End Select
AssignTypeNumeric2 = True
Exit Function
there:
End Function
Function AssignTypeNumeric(v, i As Long) As Boolean
If VarType(v) = i Then AssignTypeNumeric = True: Exit Function
On Error GoTo there
If VarType(v) = vbString Then v = Format$(v)
Select Case i
Case vbBoolean
v = CBool(v)
Case vbCurrency
v = CCur(v)
Case vbDecimal
v = CDec(v)
Case vbLong
v = CLng(v)
Case vbSingle
v = CSng(v)
Case vbInteger
v = CInt(v)
Case Else
v = CDbl(v)
End Select
AssignTypeNumeric = True
Exit Function
there:
If Err = 6 Then
Err.clear
OverflowLong i = vbInteger
Exit Function
End If
MyEr "Can't convert value", "Δεν μπορώ να μετατρέψω την τιμή"
End Function
Function MergeOperators(ByVal a$, ByVal b$) As String
If a$ = vbNullString Then MergeOperators = b$: Exit Function
If b$ = vbNullString Then MergeOperators = a$: Exit Function
If a$ = b$ Then MergeOperators = a$: Exit Function
Dim BR() As String, i As Long
If Len(a$) > Len(b$) Then
BR() = Split("[]" + b$ + "[]", "][")
For i = 1 To UBound(BR) - 1
If InStr(a$, "[" + BR(i) + "]") = 0 Then a$ = a$ + "[" + BR(i) + "]"
Next i
MergeOperators = a$
Else
BR() = Split("[]" + a$ + "[]", "][")
For i = 1 To UBound(BR) - 1
If InStr(b$, "[" + BR(i) + "]") = 0 Then b$ = b$ + "[" + BR(i) + "]"
Next i
MergeOperators = b$
End If
End Function
Public Sub GarbageFlush()
ReDim Trush(500) As VarItem
Dim i As Long
For i = 1 To 500
   Set Trush(i) = New VarItem
Next i
TrushCount = 500
End Sub
Public Sub GarbageFlush2()
ReDim Trush(500) As VarItem
Dim i As Long
For i = 1 To 500
   Set Trush(i) = New VarItem
Next i
TrushCount = 500
End Sub
Function PointPos(F$) As Long
Dim er As Long, er2 As Long
While FastSymbol(F$, Chr(34))
F$ = GetStrUntil(Chr(34), F$)
Wend
Dim i As Long, j As Long, oj As Long
If F$ = vbNullString Then
PointPos = 1
Else
er = 3
er2 = 3
For i = 1 To Len(F$)
er = er + 1
er2 = er2 + 1
Select Case Mid$(F$, i, 1)
Case "."
oj = j: j = i
Case "\", "/", ":", Is = Chr(34)
If er = 2 Then oj = 0: j = i - 2: Exit For
er2 = 1
oj = j: j = 0
If oj = 0 Then oj = i - 1: If oj < 0 Then oj = 0
Case " ", ChrW(160)
If j > 0 Then Exit For
If er2 = 2 Then oj = 0: j = i - 1: Exit For
er = 1
Case "|", "'"
j = i - 1
Exit For
Case Is > " "

If j > 0 Then oj = j Else oj = 0
Case Else
If oj <> 0 Then j = oj Else j = i
Exit For
End Select
Next i
If j = 0 Then
If oj = 0 Then
j = Len(F$) + 1
Else
j = oj
End If
End If
While Mid$(F$, j, i) = " "
j = j - 1
Wend
PointPos = j
End If
End Function
Public Function ExtractType(F$, Optional JJ As Long = 0) As String
Dim i As Long, j As Long, d$
If FastSymbol(F$, Chr(34)) Then F$ = GetStrUntil(Chr(34), F$)
If F$ = vbNullString Then ExtractType = vbNullString: Exit Function
If JJ > 0 Then
j = JJ
Else


j = PointPos(F$)
End If
d$ = F$ & " "
If j < Len(d$) Then
For i = j To Len(d$)
Select Case Mid$(d$, i, 1)
Case "/", "|", "'", " ", Is = Chr(34)
i = i + 1
Exit For
End Select
Next i
If (i - j - 2) < 1 Then
ExtractType = vbNullString
Else
ExtractType = mylcasefILE(Mid$(d$, j + 1, i - j - 2))
End If
Else
ExtractType = vbNullString
End If
End Function


Public Function CFname(a$, Optional TS As Variant, Optional createtime As Variant) As String
If Len(a$) > 2000 Then Exit Function
Dim b$
Dim mDir As New recDir
If Not IsMissing(createtime) Then
mDir.UseUTC = createtime <= 0
End If
Sleep 1
If a$ <> "" Then
On Error GoTo 1
b$ = mDir.Dir1(a$, GetCurDir)
If b$ = vbNullString Then b$ = mDir.Dir1(a$, mDir.GetLongName(App.path))
If b$ <> "" Then
CFname = mylcasefILE(b$)
If Not IsMissing(TS) Then
If Not IsMissing(createtime) Then
If Abs(createtime) = 1 Then
TS = CDbl(mDir.lastTimeStamp2)
Else
TS = CDbl(mDir.lastTimeStamp)
End If
Else
TS = CDbl(mDir.lastTimeStamp)
End If
End If
End If
Exit Function
End If
1:
CFname = vbNullString
End Function

Public Function LONGNAME(Spath As String) As String
LONGNAME = ExtractPath(Spath, , True)
End Function
Public Function ExpEnvirStr(strInput) As String
Dim Result As Long
Dim strOutput As String
'' Two calls required, one to get expansion buffer length first then do expansion
strOutput = space$(1000)
Result = ExpandEnvironmentStrings(StrPtr(strInput), StrPtr(strOutput), Result)
strOutput = space$(Result)
Result = ExpandEnvironmentStrings(StrPtr(strInput), StrPtr(strOutput), Result)
ExpEnvirStr = StripTerminator(strOutput)
End Function

Public Function ExtractPath(ByVal F$, Optional Slash As Boolean = True, Optional existonly As Boolean = False) As String
If F$ = vbNullString Then Exit Function
Dim i As Long, j As Long, test$
test$ = F$ & " \/:": i = InStr(test$, " "): j = InStr(test$, "\")
If i < j Then j = InStr(test$, "/"): If i < j Then j = InStr(test$, ":"): If i < j Then Exit Function
If Right(F$, 1) = "\" Or Right(F$, 1) = "/" Then F$ = F$ & " a"
j = PointPos(F$)
If Mid$(F$, j, 1) = "." Then j = j - 1
If Len(F$) < j Then
If ExtractType(Mid$(F$, j) & "\.10") = "10" Then j = j - 1 Else Exit Function
Else

End If

j = j - Len(ExtractNameOnly(F$))
If j <= 3 Then
If Mid$(F$, 2, 1) = ":" Then
If Slash Then
ExtractPath = mylcasefILE(Left$(F$, 2)) & "\"
Else
ExtractPath = mylcasefILE(Left$(F$, 2))
End If
Else
ExtractPath = vbNullString
End If
Else
If Slash Then
ExtractPath = mylcasefILE(Left$(F$, j))
Else
ExtractPath = mylcasefILE(Left$(F$, j - 1))
End If
End If

If existonly Then
ExtractPath = mylcasefILE(StripTerminator(GetLongName(ExpEnvirStr(ExtractPath))))
Else
ExtractPath = ExpEnvirStr(ExtractPath)
End If
Dim ccc() As String, c$
ccc() = Split(ExtractPath, "\..")
If UBound(ccc()) > LBound(ccc()) Then
c$ = vbNullString
For i = LBound(ccc()) To UBound(ccc()) - 1
If ccc(i) = vbNullString Then
c$ = ExtractPath(ExtractPath(c$, False))
Else
c$ = c$ & ExtractPath(ccc(i), True)
End If

Next i
If Left$(ccc(i), 1) = "\" Then
ExtractPath = c$ & Mid$(ccc(i), 2)
Else
ExtractPath = c$ & ccc(i)
End If
End If
End Function
Public Function ExtractName(F$) As String
Dim i As Long, j As Long, k$
If F$ = vbNullString Then Exit Function
j = PointPos(F$)
If Mid$(F$, j, 1) = "." Then
k$ = ExtractType(F$, j)
Else
j = Len(F$)
End If
For i = j To 1 Step -1
Select Case Mid$(F$, i, 1)
Case Is < " ", "\", "/", ":"
Exit For
End Select
Next i
If k$ = vbNullString Then
If Mid$(F$, i + j - i, 1) = "." Then
ExtractName = mylcasefILE(Mid$(F$, i + 1, j - i - 1))
Else
ExtractName = mylcasefILE(Mid$(F$, i + 1, j - i))

End If
Else
ExtractName = mylcasefILE(Mid$(F$, i + 1, j - i)) + k$
End If

'ExtractName = mylcasefILE(Trim$(Mid$(f$, I + 1, j - I)))

End Function
Public Function ExtractNameOnly(ByVal F$) As String
Dim i As Long, j As Long
If F$ = vbNullString Then Exit Function
j = PointPos(F$)
If j > Len(F$) Then j = Len(F$)
For i = j To 1 Step -1
Select Case Mid$(F$, i, 1)
Case Is < " ", "\", "/", ":"
Exit For
End Select
Next i
If Mid$(F$, i + j - i, 1) = "." Then
ExtractNameOnly = mylcasefILE(Mid$(F$, i + 1, j - i - 1))
Else
ExtractNameOnly = mylcasefILE(Mid$(F$, i + 1, j - i))
End If
End Function
Public Function GetCurDir(Optional AppPath As Boolean = False) As String
Dim a$, cd As String

If AppPath Then
cd = App.path
AddDirSep cd
a$ = mylcasefILE(cd)
Else
AddDirSep mcd
a$ = mylcasefILE(mcd)

End If
'If Right$(a$, 1) <> "\" Then a$ = a$ & "\"
GetCurDir = a$
End Function
Sub MakeGroupPointer(bstack As basetask, v, Optional usethisname As String = "", Optional glob As Boolean)
Dim varv As New Group
    With varv
        .IamGlobal = v.IamGlobal
        .IamApointer = True
        .BeginFloat 2
        Set .Sorosref = v.soros
        If Not v.IamFloatGroup Then
       If Len(usethisname) > 0 Then
       If glob Then
         .IamGlobal = True
       Else
        .lasthere = here$
        
        End If
        .GroupName = usethisname
       Else
       If Not .IamGlobal Then
       
        .lasthere = here$
        End If
       
        If Len(v.GroupName) > 1 Then
            .GroupName = Mid$(v.GroupName, 1, Len(v.GroupName) - 1)
        End If
        End If
        End If
    End With
     Set varv.LinkRef = v
Set bstack.lastpointer = varv
Set bstack.lastobj = varv
End Sub
Function PreparePointer(bstack As basetask) As Boolean
Dim a As Group, pppp As mArray
    If bstack.lastpointer Is Nothing Then
    
    Else
        Set a = bstack.lastpointer
        
            Set pppp = New mArray
            pppp.PushDim 1
            pppp.PushEnd
            pppp.Arr = True
            Set pppp.item(0) = a
            Set bstack.lastpointer = pppp
            PreparePointer = True
  
    End If
    
End Function
Function BoxGroupVar(aGroup As Variant) As mArray
            Set BoxGroupVar = New mArray
            BoxGroupVar.PushDim 1
            BoxGroupVar.PushEnd
            BoxGroupVar.Arr = True
            Set BoxGroupVar.item(0) = aGroup
End Function

Function BoxGroupObj(aGroup As Object) As mArray
            Set BoxGroupObj = New mArray
            BoxGroupObj.PushDim 1
            BoxGroupObj.PushEnd
            BoxGroupObj.Arr = True
            Set BoxGroupObj.item(0) = aGroup
End Function

Sub monitor(bstack As basetask, prive As basket, Lang As Long)
    Dim ss$, di As Object
    Set di = bstack.Owner
    If Lang = 0 Then
        wwPlain bstack, prive, "Εξ ορισμού κωδικοσελίδα: " & GetACP, bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Φάκελος εφαρμογής", bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, PathFromApp("m2000"), bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Καταχώρηση gsb", bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, myRegister("gsb"), bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Φάκελος προσωρινών αρχείων", bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, LONGNAME(strTemp), bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Τρέχον φάκελος", bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, mcd, bstack.Owner.Width, 1000, True
        If m_bInIDE Then
        wwPlain bstack, prive, "Όριο Αναδρομής για Συναρτήσεις " + CStr(stacksize \ 2948 - 1), bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Όριο Αναδρομής Συναρτήσεων/Τμημάτων με την Κάλεσε " + CStr(stacksize \ 1772 - 1), bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Όριο κλήσεων για Τμήματα " + CStr(stacksize \ 1254 - 1), bstack.Owner.Width, 1000, True
        Else
        wwPlain bstack, prive, "Όριο Αναδρομής για Συναρτήσεις " + CStr(stacksize \ 9832 - 1), bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Όριο Αναδρομής Συναρτήσεων/Τμημάτων με την Κάλεσε " + CStr(stacksize \ 5864), bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Όριο κλήσεων για Τμήματα  " + CStr(stacksize \ 5004), bstack.Owner.Width, 1000, True
        End If
        If OverideDec Then wwPlain bstack, prive, "Αλλαγή Τοπικού " + CStr(cLid), bstack.Owner.Width, 1000, True
        If UseIntDiv Then ss$ = "+DIV" Else ss$ = "-DIV"
        If priorityOr Then ss$ = ss$ + " +PRI" Else ss$ = ss$ + " -PRI"
        If Not mNoUseDec Then ss$ = ss$ + " -DEC" Else ss$ = ss$ + " +DEC"
        If mNoUseDec <> NoUseDec Then ss$ = ss$ + "(παράκαμψη)"
        If mTextCompare Then ss$ = ss$ + " +TXT" Else ss$ = ss$ + " -TXT"
        If ForLikeBasic Then ss$ = ss$ + " +FOR" Else ss$ = ss$ + " -FOR"
        If DimLikeBasic Then ss$ = ss$ + " +DIM" Else ss$ = ss$ + " -DIM"
        If ShowBooleanAsString Then ss$ = ss$ + " +SBL" Else ss$ = ss$ + " -SBL"
        If wide Then ss$ = ss$ + " +EXT" Else ss$ = ss$ + " -EXT"
        If RoundDouble Then ss$ = ss$ + " +RDB" Else ss$ = ss$ + " -RDB"
        If SecureNames Then ss$ = ss$ + " +SEC" Else ss$ = ss$ + " -SEC"
        If UseTabInForm1Text1 Then ss$ = ss$ + " +TAB" Else ss$ = ss$ + " -TAB"
        wwPlain bstack, prive, "Διακόπτες " + ss$, bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Περί διακοπτών: χρησιμοποίησε την εντολή Βοήθεια Διακόπτες", bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Οθόνες:" + Str$(DisplayMonitorCount()) + "  η βασική :" + Str$(FindPrimary + 1), bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Αυτή η φόρμα είναι στην οθόνη:" + Str$(FindFormSScreen(di) + 1), bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Η κονσόλα είναι στην οθόνη:" + Str$(Console + 1), bstack.Owner.Width, 1000, True

    Else
        wwPlain bstack, prive, "Default Code Page:" & GetACP, bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "App Path", bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, PathFromApp("m2000"), bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Register gsb", bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, myRegister("gsb"), bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Temporary", bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, LONGNAME(strTemp), bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Current directory", bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, mcd, bstack.Owner.Width, 1000, True
        If m_bInIDE Then
        wwPlain bstack, prive, "Max Limit for Function Recursion " + CStr(stacksize \ 2948 - 1), bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Max Limit for Function/Module Recursion using Call " + CStr(stacksize \ 1772 - 1), bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Max Limit for calling modules in depth " + CStr(stacksize \ 1254 - 1), bstack.Owner.Width, 1000, True
        Else
        wwPlain bstack, prive, "Max Limit for Function Recursion " + CStr(stacksize \ 9832 - 1), bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Max Limit for Function/Module Recursion using Call " + CStr(stacksize \ 5864), bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Max Limit for calling modules in depth " + CStr(stacksize \ 5004), bstack.Owner.Width, 1000, True
        End If
        If OverideDec Then wwPlain bstack, prive, "Locale Overide " + CStr(cLid), bstack.Owner.Width, 1000, True
        If UseIntDiv Then ss$ = "+DIV" Else ss$ = "-DIV"
        If priorityOr Then ss$ = ss$ + " +PRI" Else ss$ = ss$ + " -PRI"
        If Not mNoUseDec Then ss$ = ss$ + " -DEC" Else ss$ = ss$ + " +DEC"
        If mNoUseDec <> NoUseDec Then ss$ = ss$ + "(bypass)"
        If mTextCompare Then ss$ = ss$ + " +TXT" Else ss$ = ss$ + " -TXT"
        If ForLikeBasic Then ss$ = ss$ + " +FOR" Else ss$ = ss$ + " -FOR"
        If DimLikeBasic Then ss$ = ss$ + " +DIM" Else ss$ = ss$ + " -DIM"
        If ShowBooleanAsString Then ss$ = ss$ + " +SBL" Else ss$ = ss$ + " -SBL"
        If wide Then ss$ = ss$ + " +EXT" Else ss$ = ss$ + " -EXT"
        If RoundDouble Then ss$ = ss$ + " +RDB" Else ss$ = ss$ + " -RDB"
        If SecureNames Then ss$ = ss$ + " +SEC" Else ss$ = ss$ + " -SEC"
        If UseTabInForm1Text1 Then ss$ = ss$ + " +TAB" Else ss$ = ss$ + " -TAB"
        wwPlain bstack, prive, "Switches " + ss$, bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "About Switches: use command Help Switches", bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Screens:" + Str$(DisplayMonitorCount()) + "  Primary is:" + Str$(FindPrimary + 1), bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "This form is in screen:" + Str$(FindFormSScreen(di) + 1), bstack.Owner.Width, 1000, True
        wwPlain bstack, prive, "Console is in screen:" + Str$(Console + 1), bstack.Owner.Width, 1000, True
    End If
End Sub
Sub NeoSwap(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MySwap(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoComm(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyRead(3, ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoRef(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyRead(2, ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoRead(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyRead(1, ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoReport(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyReport(ObjFromPtr(basestackLP), rest$, Lang)
End Sub

Sub NeoDeclare(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyDeclare(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoMethod(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyMethod(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoWith(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyWith(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoSprite(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
Dim s$, p
If IsStrExp(ObjFromPtr(basestackLP), rest$, s$) Then
sprite ObjFromPtr(basestackLP), s$, rest$
ElseIf IsExp(ObjFromPtr(basestackLP), rest$, p) Then
spriteGDI ObjFromPtr(basestackLP), rest$
End If
resp = LastErNum1 = 0
End Sub

Sub NeoPlayer(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcPlayer(ObjFromPtr(basestackLP), rest$, Lang)
End Sub

Sub NeoPrinter(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcPrinter(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoPage(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
ProcPage ObjFromPtr(basestackLP), rest$, Lang
resp = True
End Sub
Sub NeoCompact(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
BaseCompact ObjFromPtr(basestackLP), rest$
resp = True
End Sub
Sub NeoLayer(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcLayer(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoOrder(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
MyOrder ObjFromPtr(basestackLP), rest$
resp = True
End Sub

Sub NeoDelete(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = DELfields(ObjFromPtr(basestackLP), rest$)
'resp = True  '' maybe this can be change
End Sub
Sub NeoAppend(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
Dim s$, p As Variant
resp = True
If IsExp(ObjFromPtr(basestackLP), rest$, p) Then
resp = AddInventory(ObjFromPtr(basestackLP), rest$)
ElseIf IsStrExp(ObjFromPtr(basestackLP), rest$, s$) Then
append_table ObjFromPtr(basestackLP), s$, rest$, False
Else
SyntaxError
resp = False
End If
End Sub
Sub NeoSearch(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
getrow ObjFromPtr(basestackLP), rest$, , "", Lang
resp = True
End Sub
Sub NeoRetr(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
getrow ObjFromPtr(basestackLP), rest$, , , Lang
resp = True
End Sub
Sub NeoExecute(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
If IsLabelSymbolNew(rest$, "ΚΩΔΙΚΑ", "CODE", Lang) Then
 resp = ExecCode(ObjFromPtr(basestackLP), rest$)
 Else
CommExecAndTimeOut ObjFromPtr(basestackLP), rest$
resp = True
End If

End Sub

Sub NeoTable(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
NewTable ObjFromPtr(basestackLP), rest$
resp = True
End Sub
Sub NeoBase(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
NewBase ObjFromPtr(basestackLP), rest$
resp = True
End Sub
Sub NeoHold(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcHold(ObjFromPtr(basestackLP))
End Sub
Sub NeoRelease(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcRelease(ObjFromPtr(basestackLP))
End Sub
Sub NeoSuperClass(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcClass(ObjFromPtr(basestackLP), rest$, Lang, True)
End Sub
Sub NeoClass(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcClass(ObjFromPtr(basestackLP), rest$, Lang, False)
End Sub
Sub NeoDIM(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyDim(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoPathDraw(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcPath(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoDrawings(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyDrawings(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoFill(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcFill(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoFloodFill(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcFLOODFILL(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoTextCursor(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyCursor(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoMouseIcon(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
i3MouseIcon ObjFromPtr(basestackLP), rest$, Lang
resp = True
End Sub
Sub NeoDouble(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
Dim bstack As basetask
Set bstack = ObjFromPtr(basestackLP)
SetDouble bstack.Owner
Set bstack = Nothing
resp = True
End Sub
Sub NeoNormal(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
Dim bstack As basetask
Set bstack = ObjFromPtr(basestackLP)
SetNormal bstack.Owner
Set bstack = Nothing
resp = True
End Sub
Sub NeoSort(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcSort(ObjFromPtr(basestackLP), rest$, Lang)
End Sub

Sub NeoImage(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcImage(ObjFromPtr(basestackLP), rest$, Lang)
End Sub

Sub NeoBitmaps(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyBitmaps(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoDef(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcDef(ObjFromPtr(basestackLP), rest$, Lang)
End Sub

Sub NeoMovies(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyMovies(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoSounds(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MySounds(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoPen(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcPen(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoCls(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcCls(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoStructure(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = myStructure(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoInput(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyInput(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoEvent(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = myEvent(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoProto(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcProto(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoEnum(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcEnum(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoPset(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyPset(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoModule(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyModule(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoModules(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyModules(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoGroup(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcGroup(0, ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoBack(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
ProcBackGround ObjFromPtr(basestackLP), rest$, Lang, resp
End Sub
Sub NeoOver(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcOver(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoDrop(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcDrop(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoShift(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcShift(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoShiftBack(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcShiftBack(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoLoad(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcLoad(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoText(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcText(ObjFromPtr(basestackLP), False, rest$)
End Sub
Sub NeoHtml(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcText(ObjFromPtr(basestackLP), True, rest$)
End Sub

Sub NeoCurve(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcCurve(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoPoly(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcPoly(ObjFromPtr(basestackLP), rest$, Lang)
End Sub

Sub NeoCircle(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcCircle(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoNew(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyNew(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoTitle(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcTitle(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoDraw(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcDraw(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoWidth(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcDrawWidth(ObjFromPtr(basestackLP), rest$)
End Sub

Sub NeoMove(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcMove(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoStep(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcStep(ObjFromPtr(basestackLP), rest$, Lang)
End Sub

Sub NeoPrint(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = RevisionPrint(ObjFromPtr(basestackLP), rest$, 0, Lang)
End Sub
Sub NeoCopy(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyCopy(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoPrinthEX(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = RevisionPrint(ObjFromPtr(basestackLP), rest$, 1, Lang)
End Sub
Sub NeoRem(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
    Dim i As Long
    If FastSymbol(rest$, "{") Then
    i = blockLen(rest$)
    If i > 0 Then rest$ = Mid$(rest$, i + 1) Else rest$ = vbNullString
    Else
    SetNextLineNL rest$
    End If
    resp = True
End Sub
Sub NeoPush(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyPush(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoData(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyData(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoClear(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyClear(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoLinespace(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = procLineSpace(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoSet(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
Dim i As Long, s$
aheadstatusANY rest$, i
s$ = Left$(rest$, i - 1)
resp = interpret(ObjFromPtr(basestackLP), s$)
If resp Then
rest$ = Mid$(rest$, i)
Else
rest$ = s$ + Mid$(rest$, i)
End If
End Sub


Sub NeoBold(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
ProcBold ObjFromPtr(basestackLP), rest$
resp = True
End Sub
Sub NeoChooseObj(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
    resp = ProcChooseObj(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoChooseFont(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
    ProcChooseFont ObjFromPtr(basestackLP), Lang
    resp = True
End Sub
Sub NeoFont(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
    ProcChooseFont ObjFromPtr(basestackLP), Lang
    resp = True
End Sub
Sub NeoScore(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyScore(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoPlayScore(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyPlayScore(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoMode(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcMode(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoGradient(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcGradient(ObjFromPtr(basestackLP), rest$)
End Sub
Sub NeoFunction(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyFunction(0, ObjFromPtr(basestackLP), rest$, Lang)
End Sub

Sub NeoFiles(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcFiles(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoCat(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = ProcCat(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Sub NeoLet(basestackLP As Long, rest$, Lang As Long, resp As Boolean)
resp = MyLet(ObjFromPtr(basestackLP), rest$, Lang)
End Sub
Function GetArrayReference(bstack As basetask, a$, v$, PP, Result As mArray, index As Long) As Boolean
Dim dn As Long, dd As Long, p, w3, w2 As Long, pppp As mArray
If PP Is Nothing Then Exit Function
If Not TypeOf PP Is mArray Then
If TypeOf PP Is mHandler Then If PP.t1 = 3 Then If Not PP.objref Is Nothing Then If TypeOf PP.objref Is mArray Then Set PP = PP.objref: GoTo cont
Exit Function
End If
cont:
Set pppp = PP

If pppp.Arr Then
dn = 0

pppp.SerialItem (0), dd, 5
dd = dd - 1
If dd < 0 Then If Typename(pppp.GroupRef) = "PropReference" Then Exit Function
            
            
p = 0
    GetArrayReference = True
    w2 = 0



        Do While dn <= dd
                    pppp.SerialItem w3, dn, 6
                    
                        If IsExp(bstack, a$, p, , True) Then
                        If dn < dd Then
                            If Not FastSymbol(a$, ",") Then: MyErMacro a$, "need index for " & v$ & ")", "χρειάζομαι δείκτη για το πίνακα " & v$ & ")": GetArrayReference = False: Exit Function
                           
                            Else
                         If FastSymbol(a$, ",") Then
                        GetArrayReference = False
                        MyErMacro a$, "too many indexes for array " & v$ & ")", "πολλοί δείκτες για το πίνακα " & v$ & ")"
                        Exit Function
                         
                         End If
                            If Not FastSymbol(a$, ")") Then: MissSymbol ")": GetArrayReference = False: Exit Function
                            
                         
                        End If
                            On Error Resume Next
                            If p < -pppp.myarrbase Then
                            GetArrayReference = False
                              MyErMacro a$, "index too low for array " & v$ & ")", "αρνητικός δείκτης στο πίνακα " & v$ & ")"
                            Exit Function
                            End If
                            
                        If Not pppp.PushOffset(w2, dn, CLng(Fix(p))) Then
                                GetArrayReference = False
                                MyErMacro a$, "index too high for array " & v$ & ")", "δείκτης υψηλός για το πίνακα " & v$ & ")"
                                GetArrayReference = False
                            Exit Function
                            End If
                            On Error GoTo 0
                        Else
                        
                         GetArrayReference = False
                        If LastErNum = -2 Then
                        Else
                        
                        MyErMacro a$, "missing index for array " & v$ & ")", "χάθηκε δείκτης για το πίνακα " & v$ & ")"
                        End If
                        Exit Function
                        End If
                    dn = dn + 1
                    Loop
                    
                    
                        Set Result = pppp
                        index = w2
    End If
End Function
Function ProcessArray(bstack As basetask, a$, v$, PP, Result) As Boolean
Dim dn As Long, dd As Long, p, w3, w2 As Long, pppp As mArray
If Not Typename$(PP) = "mArray" Then Exit Function
Set pppp = PP

If pppp.Arr Then
dn = 0

pppp.SerialItem (0), dd, 5
dd = dd - 1
If dd < 0 Then If Typename(pppp.GroupRef) = "PropReference" Then Exit Function
            
            
p = 0
    ProcessArray = True
    w2 = 0



        Do While dn <= dd
                    pppp.SerialItem w3, dn, 6
                    
                        If IsExp(bstack, a$, p, , True) Then
                        If dn < dd Then
                            If Not FastSymbol(a$, ",") Then: MyErMacro a$, "need index for " & v$ & ")", "χρειάζομαι δείκτη για το πίνακα " & v$ & ")": ProcessArray = False: Exit Function
                           
                            Else
                         If FastSymbol(a$, ",") Then
                        ProcessArray = False
                        MyErMacro a$, "too many indexes for array " & v$ & ")", "πολλοί δείκτες για το πίνακα " & v$ & ")"
                        Exit Function
                         
                         End If
                            If Not FastSymbol(a$, ")") Then: MissSymbol ")": ProcessArray = False: Exit Function
                            
                         
                        End If
                            On Error Resume Next
                            If p < -pppp.myarrbase Then
                            ProcessArray = False
                              MyErMacro a$, "index too low for array " & v$ & ")", "αρνητικός δείκτης στο πίνακα " & v$ & ")"
                            Exit Function
                            End If
                            
                        If Not pppp.PushOffset(w2, dn, CLng(Fix(p))) Then
                                ProcessArray = False
                                MyErMacro a$, "index too high for array " & v$ & ")", "δείκτης υψηλός για το πίνακα " & v$ & ")"
                                ProcessArray = False
                            Exit Function
                            End If
                            On Error GoTo 0
                        Else
                        
                         ProcessArray = False
                        If LastErNum = -2 Then
                        Else
                        
                        MyErMacro a$, "missing index for array " & v$ & ")", "χάθηκε δείκτης για το πίνακα " & v$ & ")"
                        End If
                        Exit Function
                        End If
                    dn = dn + 1
                    Loop
                    If MyIsObject(pppp.item(w2)) Then
                        Set Result = pppp.item(w2)
                    Else
                        Result = pppp.item(w2)
                    End If
    End If
End Function
Function ReplaceCRLFSPACE(a$) As Boolean
Dim i As Long
For i = 1 To Len(a$)
Select Case AscW(Mid$(a$, i, 1))
Case 13
ReplaceCRLFSPACE = True
Case 32, 10, 160, 9
Case Else
Exit For
End Select
Next i
If i = 1 Then Exit Function
If i > Len(a$) Then a$ = vbNullString: Exit Function
Mid$(a$, 1, i - 1) = String$(i - 1, Chr(7))
End Function
Function CallAsk(bstack As basetask, a$, v$) As Boolean
If UCase(v$) = "ASK(" Then
DialogSetupLang 1
Else
DialogSetupLang 0
End If
If AskText$ = vbNullString Then: ZeroParam a$: Exit Function
If FastSymbol(a$, ",") Then IsStrExp bstack, a$, AskTitle$
If FastSymbol(a$, ",") Then IsStrExp bstack, a$, AskOk$
If FastSymbol(a$, ",") Then IsStrExp bstack, a$, AskCancel$
If FastSymbol(a$, ",") Then IsStrExp bstack, a$, AskDIB$
If FastSymbol(a$, ",") Then IsStrExp bstack, a$, AskStrInput$: AskInput = True

olamazi
CallAsk = True
End Function
Public Sub olamazi()
If Form4.Visible Then
Form4.Visible = False
If Form1.Visible Then
   
   ' If Form2.Visible Then Form2.ZOrder
    If Form1.TEXT1.Visible Then
        Form1.TEXT1.SetFocus
    Else
        Form1.SetFocus
    End If
    End If
    End If
End Sub
Sub GetGuiM2000(r$)
Dim aaa As GuiM2000
If TypeOf Screen.ActiveForm Is GuiM2000 Then
Set aaa = Screen.ActiveForm
                  If aaa.index > -1 Then
                  r$ = myUcase(aaa.MyName$ + "(" + CStr(aaa.index) + ")", True)
                  Else
                  r$ = myUcase(aaa.MyName$, True)
                  End If
Else
                r$ = vbNullString
End If

End Sub
Public Function IsSupervisor() As Boolean

Dim ss$
                 ss$ = UCase(userfiles)
                    DropLeft "\M2000_USER\", ss$
IsSupervisor = ss$ = vbNullString
End Function


Public Function UserPath() As String

Dim ss$
                 ss$ = UCase(userfiles)
                    DropLeft "\M2000_USER\", ss$
        If ss$ <> "" Then
        If CanKillFile(mcd) Then
        DropLeft "\", ss$
UserPath = Mid$(mcd, Len(userfiles) - Len(ss$) + 1)
If UserPath = vbNullString Then
UserPath = "."
End If
Else
UserPath = mcd
End If
Else
UserPath = mcd
End If
End Function
Public Function UserPath2() As String

Dim ss$
                 ss$ = UCase(userfiles)
                    DropLeft "\M2000_USER\", ss$
        If ss$ <> "" Then
        If CanKillFile(mcd) Then
        DropLeft "\", ss$
UserPath2 = Mid$(mcd, Len(userfiles) - Len(ss$) + 1)
If UserPath2 = vbNullString Then
UserPath2 = "."
End If
Else
UserPath2 = mcd
End If
Else
UserPath2 = mcd
End If
If Right$(UserPath2, 1) = "\" Then UserPath2 = Left$(UserPath2$, Len(UserPath2$) - 1)


End Function
Function Fast2Label(a$, c$, cl As Long, d$, dl As Long, ahead&) As Boolean
Dim i As Long, Pad$, j As Long
j = Len(a$)
If j = 0 Then Exit Function
i = MyTrimL(a$)
If i > j Then Exit Function
Pad$ = myUcase(Mid$(a$, i, ahead& + 1)) + " "
If j - i >= cl - 1 Then
If InStr(c$, Left$(Pad$, cl)) > 0 Then
If Mid$(Pad$, cl + 1, 1) Like "[0-9+.\( @-]" Then
a$ = Mid$(a$, MyTrimLi(a$, i + cl))
Fast2Label = True
End If
Exit Function
End If
End If
If j - i >= dl - 1 Then
If InStr(d$, Left$(Pad$, dl)) > 0 Then
If Mid$(Pad$, dl + 1, 1) Like "[0-9+.\( @-]" Then
a$ = Mid$(a$, MyTrimLi(a$, i + dl))
Fast2Label = True
End If
End If
End If
End Function
Function Fast2Symbol(a$, c$, k As Long, d$, l As Long) As Boolean
Dim i As Long, j As Long
j = Len(a$)
If j = 0 Then Exit Function
i = MyTrimL(a$)
If i > j Then Exit Function
If j - i >= k - 1 Then
    If InStr(c$, Mid$(a$, i, k)) > 0 Then
    a$ = Mid$(a$, MyTrimLi(a$, i + k))
    Fast2Symbol = True
    Exit Function
    End If
End If
'If j - i >= Len(d$) - 1 Then
If j - i >= l - 1 Then
    If InStr(d$, Mid$(a$, i, l)) > 0 Then
    a$ = Mid$(a$, MyTrimLi(a$, i + l))
    Fast2Symbol = True
    Exit Function
    End If

End If
End Function
Function FastOperator2(a$, c$, i As Long) As Boolean
If Mid$(a$, i, 1) = c$ Then
Mid$(a$, i, 1) = " "
FastOperator2 = True
End If
End Function
Function FastOperator2char(a$, c$, i As Long) As Boolean
If Mid$(a$, i, 2) = c$ Then
Mid$(a$, i, 2) = "  "
FastOperator2char = True
End If
End Function
Function FastOperator(a$, c$, i As Long, Optional cl As Long = 1, Optional Remove As Boolean = True) As Boolean
Dim j As Long
If i <= 0 Then i = 1
j = Len(a$)
If j = 0 Then Exit Function
i = MyTrimLi(a$, i)
If i > j Then i = 1 ' no spaces
If j - i < cl - 1 Then Exit Function
If c$ = Mid$(a$, i, cl) Then
'If InStr(c$, Mid$(a$, i, cl)) > 0 Then
If Remove Then Mid$(a$, i, cl) = space$(cl)
FastOperator = True
End If
End Function
Function FastType(a$, c$) As Boolean
Dim i As Long, j As Long, cl, part$
cl = Len(c$)
j = Len(a$)
If j = 0 Then Exit Function
i = MyTrimL(a$)
If i > j Then Exit Function  ' this is not good
If j - i < cl - 1 Then
Exit Function
End If
If IsLabelOnly(Mid$(a$, i, cl + 1), part$) = 1 Then

If c$ = part$ Then
a$ = Mid$(a$, MyTrimLi(a$, i + cl))
FastType = True
End If
End If
End Function
Function FastSymbol(a$, c$, Optional mis As Boolean = False, Optional cl As Long = 1) As Boolean
Dim i As Long, j As Long
'If Len(c$) <> cl Then Stop  ; only for check
j = Len(a$)
If j = 0 Then Exit Function
i = MyTrimL(a$)
If i > j Then Exit Function  ' this is not good
If j - i < cl - 1 Then
If mis Then MyEr "missing " & c$, "λείπει " & c$
Exit Function
End If
If c$ = Mid$(a$, i, cl) Then
'If InStr(c$, Mid$(a$, i, cl)) > 0 Then
a$ = Mid$(a$, MyTrimLi(a$, i + cl))
'Mid$(a$, i, cl) = Space$(cl)
FastSymbol = True
ElseIf mis Then
MyEr "missing " & c$, "λείπει " & c$
End If
End Function
Function FastSymbolNoTrimAfter(a$, c$) As Boolean
Dim i As Long, j As Long
j = Len(a$)
If j = 0 Then Exit Function
i = MyTrimL(a$)
If i > j Then Exit Function  ' this is not good
If j - i < 0 Then Exit Function
If c$ = Mid$(a$, i, 1) Then
a$ = Mid$(a$, i + 1)
FastSymbolNoTrimAfter = True
End If
End Function
Function FastSymbolAt(i As Long, a$, c$, Optional cl As Long = 1) As Boolean
Dim j As Long
j = Len(a$)
If j = 0 Then Exit Function
i = MyTrimLi(a$, i)
If i > j Then Exit Function
If j - i < cl - 1 Then Exit Function
If c$ = myUcase(Mid$(a$, i, cl)) Then i = i + cl: FastSymbolAt = True
End Function
Function FastSymbolAtNoSpace(i As Long, a$, c$, Optional cl As Long = 1) As Boolean
Dim j As Long
j = Len(a$)
If j = 0 Then Exit Function
If i > j Then Exit Function
If j - i < cl - 1 Then Exit Function
If c$ = myUcase(Mid$(a$, i, cl)) Then i = i + cl: FastSymbolAtNoSpace = True
End Function

Function NocharsInLine(a$) As Boolean
Dim i As Long, j As Long
j = Len(a$)
If j = 0 Then NocharsInLine = True: Exit Function
i = MyTrimL(a$)
If i > j Then NocharsInLine = True: Exit Function

End Function
Sub DropCommentOrLine(a$)
Dim i As Long, j As Long
again:
j = Len(a$)
If j = 0 Then a$ = vbNullString:  Exit Sub
i = MyTrimL(a$)
If i > j Then a$ = vbNullString: Exit Sub
Select Case AscW(Mid$(a$, i, 1))
Case 39, 92
' drop line
i = InStr(i, a$, vbLf)
If i = 0 Then a$ = vbNullString Else a$ = Mid$(a$, i + 1): GoTo again
Case 13
' drop one line
Mid$(a$, 1, i + 1) = space$(i + 1)
GoTo again
Case Else
If i > 1 Then Mid$(a$, 1, i - 1) = space$(i - 1)
End Select


End Sub
Function MaybeIsTwoSymbol(a$, c$, Optional l As Long = 2) As Boolean
Dim i As Long
If a$ = vbNullString Then Exit Function
i = MyTrimL(a$)
If i > Len(a$) Then Exit Function
MaybeIsTwoSymbol = c$ = Mid$(a$, i, 2)

End Function
Function MaybeIsSymbolReplace(a$, c$, d$) As Boolean
Dim i As Long
If a$ = vbNullString Then Exit Function
i = MyTrimL(a$)
If i > Len(a$) Then Exit Function
If c$ = Mid$(a$, i, 1) Then
Mid$(a$, i, 1) = d$
MaybeIsSymbolReplace = True
End If
End Function

Function MaybeIsSymbol(a$, c$) As Boolean
Dim i As Long
If a$ = vbNullString Then Exit Function
i = MyTrimL(a$)
If i > Len(a$) Then Exit Function
MaybeIsSymbol = InStr(c$, Mid$(a$, i, 1)) > 0
End Function
Function MaybeIsSymbol2(a$, c$, i As Long) As Boolean
'' for isnumber
If a$ = vbNullString Then Exit Function
i = MyTrimL(a$)
If i > Len(a$) Then Exit Function
MaybeIsSymbol2 = InStr(c$, Mid$(a$, i, 1)) > 0
End Function
Function MaybeIsSymbol3lot(a$, c$, i As Long) As Boolean
If a$ = vbNullString Then Exit Function
i = MyTrimLi(a$, IIf(i, i, 1))
If i > Len(a$) Then Exit Function
MaybeIsSymbol3lot = InStr(c$, Mid$(a$, i, 1)) > 0
End Function
Function MaybeIsSymbol3(a$, c$, i As Long) As Boolean
If a$ = vbNullString Then Exit Function
i = MyTrimLi(a$, IIf(i, i, 1))
If i > Len(a$) Then Exit Function
MaybeIsSymbol3 = c$ = Mid$(a$, i, 1)
End Function

Function MaybeIsSymbolNoSpace(a$, c$) As Boolean
MaybeIsSymbolNoSpace = Left$(a$, 1) Like c$
End Function
Function IsLabelSymbolNew(a$, gre$, Eng$, code As Long, Optional mis As Boolean = False, Optional ByVal ByPass As Boolean = False, Optional checkonly As Boolean = False, Optional Free As Boolean = True) As Boolean
' code 2  gre or eng, set new value to code 1 or 0
' 0 for gre
' 1 for eng
' return true if we have label
Dim what As Boolean, drop$
Select Case code
Case 0
IsLabelSymbolNew = IsLabelSymbol3(1032, a$, gre$, drop$, mis, ByPass, checkonly, Free)
Case 1
IsLabelSymbolNew = IsLabelSymbol3(1033, a$, Eng$, drop$, mis, ByPass, checkonly, Free)
Case 2
what = IsLabelSymbol3(1032, a$, gre$, drop$, mis, ByPass, checkonly, Free)
If what Then
code = 0
IsLabelSymbolNew = what
Exit Function
End If
what = IsLabelSymbol3(1033, a$, Eng$, drop$, mis, ByPass, checkonly, Free)
If what Then code = 1
IsLabelSymbolNew = what
End Select
End Function
Function IsLabelSymbolNewExp(a$, gre$, Eng$, code As Long, usethis$) As Boolean
' code 2  gre or eng, set new value to code 1 or 0
' 0 for gre
' 1 for eng
' return true if we have label
If Len(usethis$) = 0 Then
Dim what As Boolean
Select Case code
Case 0
IsLabelSymbolNewExp = IsLabelSymbol3(1032, a$, gre$, usethis$, False, False, False, True)
Case 1
IsLabelSymbolNewExp = IsLabelSymbol3(1033, a$, Eng$, usethis$, False, False, False, True)
Case 2
what = IsLabelSymbol3(1032, a$, gre$, usethis$, False, False, False, True)
If what Then
code = 0
IsLabelSymbolNewExp = what
Exit Function
End If
what = IsLabelSymbol3(1033, a$, Eng$, usethis$, False, False, False, True)
If what Then code = 1
IsLabelSymbolNewExp = what
End Select
Else
Select Case code
Case 0, 2
IsLabelSymbolNewExp = gre$ = usethis$
Case 1
IsLabelSymbolNewExp = Eng$ = usethis$
End Select
If IsLabelSymbolNewExp Then a$ = Mid$(a$, MyTrimL(a$) + Len(usethis$))
End If
If IsLabelSymbolNewExp Then
usethis$ = vbNullString
End If
End Function


Function IsLabelSymbol3(ByVal code As Double, a$, c$, useth$, Optional mis As Boolean = False, Optional ByVal ByPass As Boolean = False, Optional checkonly As Boolean = False, Optional needspace As Boolean = False) As Boolean
Dim test$, what$, pass As Long
If ByPass Then Exit Function

If a$ <> "" And c$ <> "" Then
    test$ = a$
    If Right$(c$, 1) <= "9" Then
        If FastSymbol(test$, c$, , Len(c$)) Then
            If needspace Then
                If test$ = vbNullString Then
                ElseIf AscW(test$) < 36 Then
                ElseIf InStr(":;\',", Left$(test$, 1)) > 0 Then ' : ; ,
                Else
                    Exit Function
                End If
            End If
            If Not checkonly Then a$ = test$
            IsLabelSymbol3 = True
        Else
            If mis Then MyEr "missing " & c$, "λείπει " & c$
        End If
        Exit Function
    Else
        pass = 1000 ' maximum
        IsLabelSymbol3 = IsLabelSYMB33(test$, what$, pass)
   
      If Len(what$) <> Len(c$) Then
               If code = 1032 Then
                useth$ = myUcase(what$, True)
            Else
                useth$ = UCase(what$)
            End If
      IsLabelSymbol3 = False
         If mis Then GoTo theremiss
        Exit Function
      End If
    End If
    If what$ = vbNullString Then
    
        If mis Then GoTo theremiss
        Exit Function
    End If
    If code = 1032 Then
        what$ = myUcase(what$, True)
    Else
        what$ = UCase(what$)
    End If
    If what$ = c$ Then
    
        test$ = Mid$(test$, pass)
        If needspace Then
            If test$ = vbNullString Then
            ElseIf AscW(test$) < 36 Then
            ElseIf InStr(":;\',", Left$(test$, 1)) > 0 Then
            ' : ; ,
            Else
                IsLabelSymbol3 = False
                Exit Function
            End If
        End If
        If checkonly Then
          '  A$ = what$ & TEST$
          Else
           a$ = test$
        End If
  
       Else
             If mis Then
theremiss:
           ''  MyErMacro a$, "missing " & c$, "λείπει " & c$
                 MyEr "missing " & c$, "λείπει " & c$
                 Else
                 useth$ = what$
              End If
            IsLabelSymbol3 = False
            End If
Else
If mis Then GoTo theremiss
End If
End Function
Function IsLabelSymbol(a$, c$, Optional mis As Boolean = False, Optional ByVal ByPass As Boolean = False, Optional checkonly As Boolean = False) As Boolean
Dim test$, what$, pass As Long
If ByPass Then Exit Function

  If a$ <> "" And c$ <> "" Then
test$ = a$
pass = Len(c$)

IsLabelSymbol = IsLabelSYMB33(test$, what$, pass)
If Len(what$) <> Len(c$) Then IsLabelSymbol = False
If Not IsLabelSymbol Then
     If mis Then
                 MyEr "missing " & c$, "λείπει " & c$
              End If
Exit Function
End If

        If myUcase(what$) = c$ Then
        If checkonly Then
     '   A$ = what$ & " " & TEST$
        Else
                    a$ = Mid$(test$, pass)
          End If
  
             Else
             If mis Then
                 MyEr "missing " & c$, "λείπει " & c$
              End If
            IsLabelSymbol = False
            End If

End If
End Function
Function IsLabelSymbolLatin(a$, c$, Optional mis As Boolean = False, Optional ByVal ByPass As Boolean = False, Optional checkonly As Boolean = False) As Boolean
Dim test$, what$, pass As Long
If ByPass Then Exit Function

  If a$ <> "" And c$ <> "" Then
test$ = a$
pass = Len(c$)
IsLabelSymbolLatin = IsLabelSYMB33(test$, what$, pass)
If Len(what$) <> Len(c$) Then IsLabelSymbolLatin = False
If Not IsLabelSymbolLatin Then
             If mis Then
                 MyEr "missing " & c$, "λείπει " & c$
              End If
            Exit Function
End If
        If UCase(what$) = c$ Then
        If checkonly Then
      '  A$ = what$ & " " & TEST$
        Else
                    a$ = Mid$(test$, pass)
          End If
  
             Else
             If mis Then
                 MyEr "missing " & c$, "λείπει " & c$
              End If
            IsLabelSymbolLatin = False
            End If

End If
End Function

Function GetRes(bstack As basetask, b$, Lang As Long, data$) As Boolean
Dim w$, x1 As Long, label1$, useHandler As mHandler, par As Boolean, pppp As mArray, p As Variant
If IsLabelSymbolNew(b$, "ΩΣ", "AS", Lang) Then
            w$ = Funcweak(bstack, b$, x1, label1$)
            If LastErNum1 = -1 And x1 < 5 Then Exit Function
            If LenB(w$) = 0 Then
            If Len(bstack.UseGroupname) > 0 Then
                If Len(label1$) > Len(bstack.UseGroupname) Then
                    If bstack.UseGroupname = Left$(label1$, Len(bstack.UseGroupname)) Then
                        MyEr "No such member in this group", "Δεν υπάρχει τέτοιο μέλος σε αυτή την ομάδα"
                        Exit Function
                    End If
                End If
            ElseIf x1 = 1 Then
contvar1:
            x1 = globalvar(label1$, 0#)
            Set useHandler = New mHandler
                useHandler.t1 = 2
        If FastSymbol(b$, ",") Then
        If IsExp(bstack, b$, p, , True) Then
         Set useHandler.objref = Decode64toMemBloc(data$, par, CBool(p))
        Else
        GetRes = True
        MissParam data$: Exit Function
        End If
        Else
                Set useHandler.objref = Decode64toMemBloc(data$, par)
                End If
                If par Then
                    Set var(x1) = useHandler
                    GetRes = True
            
                Else
                    GoTo err1
                End If
                Exit Function
            ElseIf x1 = 3 Then
                x1 = globalvar(label1$, vbNullString)
                var(x1) = Decode64(data$, par)
                If Not par Then GoTo err1
                GetRes = True
                Exit Function
            ElseIf x1 = 5 Then
                If GetVar(bstack, label1$, x1) Then
                    If GetArrayReference(bstack, b$, label1$, var(x1), pppp, x1) Then
                        Set useHandler = New mHandler
                        useHandler.t1 = 2
                        If Not par Then GoTo err1
                        If FastSymbol(b$, ",") Then
        If IsExp(bstack, b$, p, , True) Then
         Set useHandler.objref = Decode64toMemBloc(data$, par, CBool(p))
        Else
        GetRes = True
        MissParam data$: Exit Function
        End If
        Else
                        Set useHandler.objref = Decode64toMemBloc(data$, par)
                        End If
                    
                        Set pppp.item(x1) = useHandler
                        GetRes = True
                    End If
                    Exit Function
            
                Else
                    MyEr "", ""
                    MyEr "Array not defined", "Ο πίνακας δεν έχει οριστεί"
                    Exit Function
                End If
            ElseIf x1 = 6 Then
contstr1:
                If GetVar(bstack, label1$, x1) Then
                    If GetArrayReference(bstack, b$, label1$, var(x1), pppp, x1) Then
                        pppp.item(x1) = Decode64(data$, par)
                        If Not par Then GoTo err1
                        GetRes = True
                    End If
                    Exit Function
                Else
                    MyEr "", ""
                    MyEr "Array not defined", "Ο πίνακας δεν έχει οριστεί"
                    Exit Function
                End If
            End If
            End If

            If x1 = 1 Then
            If GetVar(bstack, label1$, x1) Then
            Set useHandler = New mHandler
                useHandler.t1 = 2
        
                Set useHandler.objref = Decode64toMemBloc(data$, par)
                If par Then
                    Set var(x1) = useHandler
                    GetRes = True
            
                Else
err1:
                    MyEr "Can't decode this resource", "Δεν μπορών να αποκωδικοποιήσω αυτό το πόρο"
                End If
                Exit Function
            Else
                GoTo contvar1
            End If
                ElseIf x1 = 3 Then
                
                If GetVar(bstack, label1$, x1) Then
                var(x1) = Decode64(data$, par)
                If Not par Then GoTo err1
                GetRes = True
                Exit Function
                End If
                ElseIf x1 = 5 Then
                If GetVar(bstack, label1$, x1) Then
                      DropLeft "(", w$
                    If GetArrayReference(bstack, w$, label1$, var(x1), pppp, x1) Then
                        Set useHandler = New mHandler
                        useHandler.t1 = 2
                        Set useHandler.objref = Decode64toMemBloc(data$, par)
                        If Not par Then GoTo err1
                        Set pppp.item(x1) = useHandler
                        GetRes = True
                    End If
                    Exit Function
            
                Else
                    MyEr "", ""
                    MyEr "Array not defined", "Ο πίνακας δεν έχει οριστεί"
                    Exit Function
                End If
                            ElseIf x1 = 6 Then
                               If GetVar(bstack, label1$, x1) Then
                            DropLeft "(", w$
                    If GetArrayReference(bstack, w$, label1$, var(x1), pppp, x1) Then
                        pppp.item(x1) = Decode64(data$, par)
                        If Not par Then GoTo err1
                        GetRes = True
                    End If
                    Exit Function
                Else
                    MyEr "", ""
                    MyEr "Array not defined", "Ο πίνακας δεν έχει οριστεί"
                    Exit Function
                End If
            End If
        
        End If
End Function

Function IsHILOWWORD(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
    Dim p As Variant
    If IsExp(bstack, a$, r, , True) Then
        If FastSymbol(a$, ",") Then
              If IsExp(bstack, a$, p) Then
                    r = SG * (r * &H10000 + p)
                    
                     IsHILOWWORD = FastSymbol(a$, ")", True)
                  Else
                     
                    MissParam a$
                End If
        Else
             
             MissParam a$
        End If
     Else
             
             MissParam a$
      End If
     
End Function
Function IsBinaryNot(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
  If IsExp(bstack, a$, r, , True) Then
            On Error Resume Next
    If r < 0 Then r = r And &H7FFFFFFF
             r = SG * uintnew1(Not signlong2(r))
        If Err.Number > 0 Then
            
            WrongArgument a$
          
            Exit Function
            End If
    On Error GoTo 0
    
        IsBinaryNot = FastSymbol(a$, ")", True)
    Else
           MissParam a$
    
    End If
End Function
Function IsBinaryNeg(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
  If IsExp(bstack, a$, r, , True) Then
            On Error Resume Next
       
             r = SG * CDbl(Pow2minusOne(32) - uintnew(r))
        If Err.Number > 0 Then
        
            WrongArgument a$
        
            Exit Function
            End If
    On Error GoTo 0
    
        IsBinaryNeg = FastSymbol(a$, ")", True)
    Else
           MissParam a$
    
    End If
End Function
Function IsBinaryOr(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
        Dim p As Variant
     If IsExp(bstack, a$, r, , True) Then
        If FastSymbol(a$, ",") Then
        If IsExp(bstack, a$, p) Then
            r = SG * uintnew1(signlong2(r) Or signlong2(p))
         IsBinaryOr = FastSymbol(a$, ")", True)
           Else
                
                MissParam a$
        End If
          Else
                MissParam a$
       End If
         Else
                MissParam a$
       End If
End Function
Function IsBinaryAdd(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
    Dim p As Variant
    If IsExp(bstack, a$, r, , True) Then
            If FastSymbol(a$, ",") Then
                If IsExp(bstack, a$, p, , True) Then
                    r = add32b(r, p)
                    
                    While FastSymbol(a$, ",")
                    If Not IsExp(bstack, a$, p, , True) Then MissNumExpr: Exit Function
                    r = add32b(r, p)
                    Wend
                    If SG < 0 Then r = -r
                    IsBinaryAdd = FastSymbol(a$, ")", True)
                Else
                    
                    MissParam a$
                End If
            Else
                
                MissParam a$
            End If
        Else
            
            MissParam a$
       
       End If
End Function
Function IsBinaryAnd(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
    Dim p As Variant
    If IsExp(bstack, a$, r, , True) Then
            If FastSymbol(a$, ",") Then
                If IsExp(bstack, a$, p, , True) Then
                    r = SG * uintnew1(signlong2(r) And signlong2(p))
                    
                    IsBinaryAnd = FastSymbol(a$, ")", True)
                Else
                    
                    MissParam a$
                End If
            Else
                
                MissParam a$
            End If
        Else
            
            MissParam a$
       
       End If
End Function
Function IsBinaryXor(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
    Dim p As Variant
        If IsExp(bstack, a$, r, True) Then
            If FastSymbol(a$, ",") Then
                If IsExp(bstack, a$, p) Then
                    r = SG * uintnew1(signlong2(r) Xor signlong2(p))
                    
                    IsBinaryXor = FastSymbol(a$, ")", True)
                Else
                    
                    MissParam a$
                End If
            Else
                
                MissParam a$
            End If
        Else
            
            MissParam a$
       
       End If
End Function
Function IsBinaryShift(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
Dim p As Variant
   If IsExp(bstack, a$, r, , True) Then
  
            If FastSymbol(a$, ",") Then
                    If IsExp(bstack, a$, p) Then
                         If p > 31 Or p < -31 Then
                         
                         MyErMacro a$, "Shift from -31 to 31", "Ολίσθηση από -31 ως 31"
                         IsBinaryShift = False: Exit Function
                         Else
                               If p > 0 Then
                              
                                 r = SG * CCur((signlong(r) And signlong(Pow2minusOne(32 - p))) * Pow2(p))
                              
                              ElseIf p = 0 Then
                              If SG < 0 Then r = -CCur(r) Else r = CCur(r)
                              Else
                                    
                                 r = SG * CCur(Int(CCur(r) / Pow2(-p)))
                              End If
                              
                            IsBinaryShift = FastSymbol(a$, ")", True)
                    Exit Function
                         End If
                    Else
                          
                        MissParam a$
                    End If
            Else
                
                MissParam a$
            End If
    Else
            
            MissParam a$
   End If

End Function
Function IsBinaryRotate(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
Dim p As Variant
        If IsExp(bstack, a$, r, , True) Then
             If FastSymbol(a$, ",") Then
                 If IsExp(bstack, a$, p) Then
                        If p > 31 Or p < -31 Then
                            
                              MyErMacro a$, "Rotation from -31 to 31", "Περιστοφή από -31 ως 31"
                             IsBinaryRotate = False: Exit Function
                        Else
                             If p > 0 Then
                          
                                 r = SG * CCur((signlong(r) And signlong(Pow2minusOne(32 - p))) * Pow2(p) + Int(CCur(r) / Pow2(32 - p)))
                             ElseIf p = 0 Then
                                 If SG < 0 Then r = -CCur(r) Else r = CCur(r)
                             Else
                          
                                 r = SG * CCur((signlong(r) And signlong(Pow2minusOne(-p))) * Pow2(32 + p) + Int(CCur(r) / Pow2(-p)))
                             End If
                        End If
                     
                  Else
                    
                    MissParam a$
                 End If
             Else
                
                MissParam a$
            End If
        IsBinaryRotate = FastSymbol(a$, ")", True)
        Else
            
            MissParam a$
        End If
End Function
Function IsSin(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
   If IsExp(bstack, a$, r, , True) Then
    r = Sin(r * 1.74532925199433E-02)
    ''r = Sgn(r) * Int(Abs(r) * 10000000000000#) / 10000000000000#
    If Abs(r) < 1E-16 Then r = 0
    If SG < 0 Then r = -r
    
    
 IsSin = FastSymbol(a$, ")", True)
    Else
                
                MissParam a$
    
    End If
End Function
Function IsAbs(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
If IsExp(bstack, a$, r, , True) Then
    r = Abs(r)
    If SG < 0 Then r = -r
    
 IsAbs = FastSymbol(a$, ")", True)
    Else
                MissParam a$
    End If
End Function

Function IsCos(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
  If IsExp(bstack, a$, r, , True) Then

    r = Cos(r * 1.74532925199433E-02)
 
    If Abs(r) < 1E-16 Then r = 0
    If SG < 0 Then r = -r
    
    
  IsCos = FastSymbol(a$, ")", True)
    Else
                
                MissParam a$
    
    End If
End Function
Function IsTan(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
If IsExp(bstack, a$, r, , True) Then
     
     If r = Int(r) Then
        If r Mod 90 = 0 And r Mod 180 <> 0 Then
        MyErMacro a$, "Wrong Tan Parameter", "Λάθος παράμετρος εφαπτομένης"
        IsTan = False: Exit Function
        End If
        End If
    r = Sgn(r) * Tan(r * 1.74532925199433E-02)

     If Abs(r) < 1E-16 Then r = 0
     If Abs(r) < 1 And Abs(r) + 0.0000000000001 >= 1 Then r = Sgn(r)
   If SG < 0 Then r = -r
    
IsTan = FastSymbol(a$, ")", True)
     Else
                
                MissParam a$
    
    End If
End Function
Function IsAtan(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
 If IsExp(bstack, a$, r, , True) Then
     
     r = SG * Atn(r) * 180# / Pi
        
IsAtan = FastSymbol(a$, ")", True)
     Else
                
                MissParam a$
    
    End If
End Function
Function IsLn(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
  If IsExp(bstack, a$, r, , True) Then
    If r <= 0 Then
       MyErMacro a$, "Only > zero parameter", "Μόνο >0 παράμετρος"
        IsLn = False: Exit Function
    Else
    r = SG * Log(r)
    
    End If
    
 IsLn = FastSymbol(a$, ")", True)
     Else
                
                MissParam a$
    
    End If
End Function
Function IsLog(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
If IsExp(bstack, a$, r, , True) Then
        If r <= 0 Then
       MyErMacro a$, "Only > zero parameter", "Μόνο >0 παράμετρος"
        IsLog = False: Exit Function
    Else
    r = SG * Log(r) / 2.30258509299405
    
    End If
   IsLog = FastSymbol(a$, ")", True)
    Else
                
                MissParam a$
    
    End If
End Function
Function IsFreq(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
Dim p As Variant
    If IsExp(bstack, a$, r, , True) Then
           If FastSymbol(a$, ",") Then
                If IsExp(bstack, a$, p) Then
                    r = SG * GetFrequency(CInt(r), CInt(p))
                    
                    IsFreq = FastSymbol(a$, ")", True)
                    Else
                
                MissParam a$
                End If
            Else
                
                MissParam a$
            End If
     Else
                
                MissParam a$
     End If
End Function
Function IsSqrt(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
    If IsExp(bstack, a$, r, , True) Then
    
    If r < 0 Then
    negsqrt a$
    Exit Function
   
    End If
  
    r = Sqr(r)
    If SG < 0 Then r = -r
    
   IsSqrt = FastSymbol(a$, ")", True)
    Else
                
                MissParam a$
    
    End If
End Function
Function GiveForm() As Form
Set GiveForm = Form1
End Function
Function IsNumberD(a$, d As Double) As Boolean
Dim a1 As Long
If a$ <> "" Then
For a1 = 1 To Len(a$) + 1
Select Case Mid$(a$, a1, 1)
Case " ", ",", ChrW(160)
If a1 > 1 Then Exit For
Case Is = Chr(2)
If a1 = 1 Then Exit Function
Exit For
End Select
Next a1
If a1 > Len(a$) Then a1 = Len(a$) + 1
d = CDbl(val("0" & Left$(a$, a1 - 1)))
a$ = Mid$(a$, a1)
IsNumberD = True
Else
IsNumberD = False
End If
End Function
Function IsNumberLabel2(a$, Label$, a1 As Long, ByVal LI As Long) As Boolean
Dim A2 As Long
If LI > 0 Then
A2 = a1
If a1 > LI Then Exit Function
If LI > 5 + A2 Then LI = 4 + A2
If Mid$(a$, a1, 1) Like "[0-9]" Then
Do While a1 <= LI
a1 = a1 + 1
If Not Mid$(a$, a1, 1) Like "[0-9]" Then Exit Do

Loop
Label$ = Mid$(a$, A2, a1 - A2)
IsNumberLabel2 = True
End If
End If
End Function
Function IsNumberLabel(a$, Label$) As Boolean
Dim a1 As Long, LI As Long, A2 As Long
LI = Len(a$)

If LI > 0 Then

a1 = MyTrimL(a$)

A2 = a1
If a1 > LI Then a$ = vbNullString: Exit Function
If LI > 5 + A2 Then LI = 4 + A2
If Mid$(a$, a1, 1) Like "[0-9]" Then
Do While a1 <= LI
a1 = a1 + 1
If Not Mid$(a$, a1, 1) Like "[0-9]" Then Exit Do

Loop
Label$ = Mid$(a$, A2, a1 - A2): a$ = Mid$(a$, a1)
IsNumberLabel = True
End If

End If
End Function
Function IsNumberQuery(a$, fr As Long, r As Double, lr As Long) As Boolean
Dim SG As Long, sng As Long, n$, ig$, DE$, sg1 As Long, ex$, rr As Double
' ti kanei to e$
If a$ = vbNullString Then IsNumberQuery = False: Exit Function
SG = 1
sng = fr - 1
    Do While sng < Len(a$)
    sng = sng + 1
    Select Case Mid$(a$, sng, 1)
    Case " ", "+", ChrW(160)
    Case "-"
    SG = -SG
    Case Else
    Exit Do
    End Select
    Loop
n$ = Mid$(a$, sng)

If val("0" & Mid$(a$, sng, 1)) = 0 And Left(Mid$(a$, sng, 1), sng) <> "0" And Left(Mid$(a$, sng, 1), sng) <> "." Then
IsNumberQuery = False

Else
'compute ig$
    If Mid$(a$, sng, 1) = "." Then
    ' no long part
    ig$ = "0"
    DE$ = "."

    Else
    Do While sng <= Len(a$)
        
        Select Case Mid$(a$, sng, 1)
        Case "0" To "9"
        ig$ = ig$ & Mid$(a$, sng, 1)
        Case "."
        DE$ = "."
        Exit Do
        Case Else
        Exit Do
        End Select
       sng = sng + 1
    Loop
    End If
    ' compute decimal part
    If DE$ <> "" Then
      sng = sng + 1
        Do While sng <= Len(a$)
       
        Select Case Mid$(a$, sng, 1)
        Case " ", ChrW(160)
        If Not (sg1 And Len(ex$) = 1) Then
        Exit Do
        End If
        Case "0" To "9"
        If sg1 Then
        ex$ = ex$ & Mid$(a$, sng, 1)
        Else
        DE$ = DE$ & Mid$(a$, sng, 1)
        End If
        Case "E", "e", "Ε", "ε" ' ************check it
             If ex$ = vbNullString Then
               sg1 = True
        ex$ = "E"
        Else
        Exit Do
        End If
        
        Case "+", "-"
        If sg1 And Len(ex$) = 1 Then
         ex$ = ex$ & Mid$(a$, sng, 1)
        Else
        Exit Do
        End If
        Case Else
        Exit Do
        End Select
         sng = sng + 1
        Loop
        If sg1 Then
            If Len(ex$) < 3 Then
                If ex$ = "E" Then
                    ex$ = " "
                ElseIf ex$ = "E-" Or ex$ = "E+" Then
                    ex$ = "  "
                End If
            End If
        End If
    End If
    If ig$ = vbNullString Then
    IsNumberQuery = False
    lr = 1
    Else
    If SG < 0 Then ig$ = "-" & ig$
    Err.clear
    On Error Resume Next
    n$ = ig$ & DE$ & ex$
    sng = Len(ig$ & DE$ & ex$)
    rr = val(ig$ & DE$ & ex$)
    If Err.Number > 0 Then
         lr = 0
    Else
        r = rr
       lr = sng - fr + 2
       IsNumberQuery = True
    End If
    
       
    
    End If
End If
End Function


Function IsNumberOnly(a$, fr As Long, r As Variant, lr As Long, Optional useRtypeOnly As Boolean = False, Optional usespecial As Boolean = False) As Boolean
Dim SG As Long, sng As Long, ig$, DE$, sg1 As Long, ex$, foundsign As Boolean
' ti kanei to e$
If a$ = vbNullString Then IsNumberOnly = False: Exit Function
SG = 1
sng = fr - 1
    Do While sng < Len(a$)
    sng = sng + 1
    Select Case Mid$(a$, sng, 1)
    Case " ", ChrW(160)
    Case "+"
    foundsign = True
    Case "-"
    SG = -SG
    foundsign = True
    Case Else
    Exit Do
    End Select
    Loop
If LCase(Mid$(a$, sng, 2)) Like "0[xχ]" Then
    If foundsign Then
    MyEr "no sign for hex values", "όχι πρόσημο για δεκαεξαδικούς"
    IsNumberOnly = False
    GoTo er111
    End If
    ig$ = vbNullString
    DE$ = vbNullString
    sng = sng + 1
    Do While MaybeIsSymbolNoSpace(Mid$(a$, sng + 1, 1), "[0-9A-Fa-f]")
    DE$ = DE$ + Mid$(a$, sng + 1, 1)
    sng = sng + 1
    If Len(DE$) = 8 Then Exit Do
    Loop
    sng = sng + 1
    SG = 1 ' no sign
    If LenB(DE$) = 0 Then
    MyEr "ivalid hex values", "λάθος όρισμα για δεκαεξαδικό"
    IsNumberOnly = False
    GoTo er111
    End If
    If MaybeIsSymbolNoSpace(Mid$(a$, sng, 1), "[&%]") Then
    
        sng = sng + 1
        ig$ = "&H" + DE$
        DE$ = vbNullString
        If Mid$(a$, sng - 1, 1) = "%" Then
        If Len(ig$) > 6 Then
        OverflowLong True
        IsNumberOnly = False
        GoTo er111
        Else
        r = CInt(0)
        End If
        Else
        r = CLng(0)
        End If
        GoTo conthere1
    ElseIf useRtypeOnly Then
        If VarType(r) = vbLong Or VarType(r) = vbInteger Then
        ig$ = "&H" + DE$
        DE$ = vbNullString
        GoTo conthere1
        End If
    End If
        DE$ = Right$("00000000" & DE$, 8)
        r = CDbl(UNPACKLNG(Left$(DE$, 4)) * 65536#) + CDbl(UNPACKLNG(Right$(DE$, 4)))
        GoTo contfinal
  
ElseIf val("0" & Mid$(a$, sng, 1)) = 0 And Left(Mid$(a$, sng, 1), sng) <> "0" And Left(Mid$(a$, sng, 1), sng) <> "." Then
IsNumberOnly = False

Else
'compute ig$
    If Mid$(a$, sng, 1) = "." Then
    ' no long part
    ig$ = "0"
    DE$ = "."

    Else
    Do While sng <= Len(a$)
        
        Select Case Mid$(a$, sng, 1)
        Case "0" To "9"
        ig$ = ig$ & Mid$(a$, sng, 1)
        Case "."
        DE$ = "."
        Exit Do
        Case Else
        Exit Do
        End Select
       sng = sng + 1
    Loop
    End If
    ' compute decimal part
    If DE$ <> "" Then
      sng = sng + 1
        Do While sng <= Len(a$)
       
        Select Case Mid$(a$, sng, 1)
        Case " ", ChrW(160)
        If Not (sg1 And Len(ex$) = 1) Then
        Exit Do
        End If
        Case "0" To "9"
        If sg1 Then
        ex$ = ex$ & Mid$(a$, sng, 1)
        Else
        DE$ = DE$ & Mid$(a$, sng, 1)
        End If
        Case "E", "e", "Ε", "ε"  ' ************check it
            If ex$ = vbNullString Then
               sg1 = True
                ex$ = "E"
            Else
                Exit Do
            End If
        Case "+", "-"
            If sg1 And Len(ex$) = 1 Then
             ex$ = ex$ & Mid$(a$, sng, 1)
            Else
                Exit Do
            End If
        Case Else
            Exit Do
        End Select
        sng = sng + 1
        Loop
        If Len(ex$) < 3 Then
                If ex$ = "E" Then
                ex$ = "0"
                sng = sng + 1
                ElseIf ex$ = "E-" Or ex$ = "E+" Then
                ex$ = "00"
                sng = sng + 2
                End If
                End If
    End If
    If ig$ = vbNullString Then
    IsNumberOnly = False
    lr = 1
    Else
    If SG < 0 Then ig$ = "-" & ig$
    On Error GoTo er111
     If useRtypeOnly Then GoTo conthere1
    If sng <= Len(a$) Then
    If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = cdecimaldot$
    Select Case Mid$(a$, sng, 1)
    Case "@"
    r = CDec(ig$ & DE$)
    sng = sng + 1
    Case "&"
    r = CLng(ig$)
    sng = sng + 1
    Case "%"
    r = CInt(ig$)
    sng = sng + 1
    Case "~"
    r = CSng(ig$ & DE$ & ex$)
    sng = sng + 1
    Case "#"
    r = CCur(ig$ & DE$)
    sng = sng + 1
    Case Else
GoTo conthere
    End Select
    Else
conthere:
        If useRtypeOnly Then
conthere1:
        If usespecial Then
       If sng <= Len(a$) Then
            Select Case Mid$(a$, sng, 1)
            Case "@"
                r = CDec(0)
                sng = sng + 1
            Case "&"
                r = CLng(0)
                sng = sng + 1
            Case "~"
                r = CSng(0)
                sng = sng + 1
            Case "#"
                r = CCur(0)
                sng = sng + 1
            Case "%"
                r = CInt(0)
                sng = sng + 1
        End Select
        End If
        End If
         If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = cdecimaldot$
        Select Case VarType(r)
        Case vbDecimal
        r = CDec(ig$ & DE$)
        Case vbLong
        r = CLng(ig$)
        Case vbInteger
        r = CInt(ig$)
        Case vbSingle
        r = CSng(ig$ & DE$ & ex$)
        Case vbCurrency
        r = CCur(ig$ & DE$)
        Case vbBoolean
        r = CBool(ig$ & DE$)
        Case Else
        r = CDbl(ig$ & DE$ & ex$)
        End Select
        Else
        If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = "."
        r = val(ig$ & DE$ & ex$)
        End If
    End If
contfinal:
    lr = sng - fr + 1
    
    IsNumberOnly = True
    Exit Function
    End If
End If
er111:
    lr = sng - fr + 1
    Err.clear
Exit Function

End Function


Function IsNumberD2(a$, d As Variant, Optional noendtypes As Boolean = False, Optional exceptspecial As Boolean) As Boolean
' for inline stacitems
If VarType(d) = vbEmpty Then d = 0#
Dim a1 As Long
If a$ <> "" Then
For a1 = 1 To Len(a$) + 1
Select Case Mid$(a$, a1, 1)
Case " ", ChrW(160)
If a1 > 1 Then Exit For
Case Is = Chr(2)
If a1 = 1 Then Exit Function
Exit For
End Select
Next a1
If a1 > Len(a$) Then a1 = Len(a$) + 1
    If IsNumberOnly(a$, 1, d, a1, noendtypes, exceptspecial) Then
        a$ = Mid$(a$, a1)
        IsNumberD2 = True
    ElseIf MaybeIsSymbol(a$, "ΑαΨψTtFf") Then
        If Fast3NoSpace(a$, "ΑΛΗΘΕΣ", 6, "ΑΛΗΘΗΣ", 6, "TRUE", 4, 6) Then
            d = True
            IsNumberD2 = True
        ElseIf Fast3NoSpace(a$, "ΨΕΥΔΕΣ", 6, "ΨΕΥΔΗΣ", 6, "FALSE", 5, 5) Then
            d = False
            IsNumberD2 = True
        Else
            IsNumberD2 = False
        End If
    Else
    IsNumberD2 = False
    End If
Else
    IsNumberD2 = False
End If

End Function

Function IsNumberD3(a$, fr As Long, a1 As Long) As Boolean
' for inline stacitems
Dim d As Double
If a$ <> "" Then
For a1 = fr To Len(a$) + 1
Select Case Mid$(a$, a1, 1)
Case " ", ChrW(160)
If a1 > fr Then Exit For
Case Is = Chr(2)
If a1 = fr Then Exit Function
Exit For
End Select
Next a1
If a1 > Len(a$) Then a1 = Len(a$) + 1
If IsNumberOnly(a$, fr, d, a1) Then
IsNumberD3 = True
ElseIf Fast3NoSpaceCheck(fr, a$, "ΑΛΗΘΕΣ", 6, "ΑΛΗΘΗΣ", 6, "TRUE", 4, 6) Then
d = True
IsNumberD3 = True
ElseIf Fast3NoSpaceCheck(fr, a$, "ΨΕΥΔΕΣ", 6, "ΨΕΥΔΗΣ", 6, "FALSE", 5, 5) Then
d = False
IsNumberD3 = True
Else
a1 = fr
IsNumberD3 = False
End If
Else
a1 = fr
IsNumberD3 = False
End If

End Function

Sub tsekme()
Dim b$, l As Double
b$ = " 12323 45.44545 -2345.343 .345 345.E-45 34.53 434 534 534 534 345"
'b$ = VbNullString
Debug.Print b$
While IsNumberD2(b$, l)
Debug.Print l
Wend
End Sub
Function IsNumberCheck(a$, r As Variant, Optional mydec$ = " ") As Boolean
Dim sng&, SG As Variant, ig$, DE$, sg1 As Boolean, ex$, s$
If mydec$ = " " Then mydec$ = "."
SG = 1
Do While sng& < Len(a$)
sng& = sng& + 1
Select Case Mid$(a$, sng&, 1)
Case "#"
    If Len(a$) > sng& Then
    If MaybeIsSymbolNoSpace(Mid$(a$, sng& + 1, 1), "[0-9A-Fa-f]") Then
    s$ = "0x00" + Mid$(a$, sng& + 1, 6)
    If Len(s$) < 10 Then Exit Function
        If IsNumberCheck(s$, r) Then
        If s$ <> "" Then
          
             
        Else
            s$ = Right$("00000000" & Mid$(a$, sng& + 1, 6), 8)
            a$ = Mid$(a$, sng& + 7)
   r = SG * -(CDbl(UNPACKLNG(Right$(s$, 2)) * 65536#) + CDbl(UNPACKLNG(Mid$(s$, 5, 2)) * 256#) + CDbl(UNPACKLNG(Mid$(s$, 3, 2))))
   IsNumberCheck = True
   Exit Function
        End If
        End If
        Else
        
    End If
    Else

    '' out
    End If
    Exit Function
Case " ", "+", ChrW(160)
Case "-"
SG = -SG
Case Else
Exit Do
End Select
Loop
a$ = Mid$(a$, sng&)
sng& = 1
If val("0" & Mid$(Replace(a$, mydec$, "."), sng&, 1)) = 0 And Left(Mid$(a$, sng&, 1), sng&) <> "0" And Left(Mid$(a$, sng&, 1), sng&) <> mydec$ Then
IsNumberCheck = False
Else

    If Mid$(a$, sng&, 1) = mydec$ Then

    ig$ = "0"
    DE$ = mydec$
    ElseIf LCase(Mid$(a$, sng&, 2)) Like "0[xχ]" Then
    ig$ = "0"
    DE$ = "0x"
  sng& = sng& + 1
Else
    Do While sng& <= Len(a$)
        
        Select Case Mid$(a$, sng&, 1)
        Case "0" To "9"
        ig$ = ig$ & Mid$(a$, sng&, 1)
        Case mydec$
        DE$ = mydec$
        Exit Do
        Case Else
        Exit Do
        End Select
       sng& = sng& + 1
    Loop
    End If
    ' compute decimal part
    If DE$ <> "" Then
      sng& = sng& + 1
        Do While sng& <= Len(a$)
       
        Select Case Mid$(a$, sng&, 1)
        Case " ", ChrW(160)
        If Not (sg1 And Len(ex$) = 1) Then
        Exit Do
        End If
        Case "A" To "D", "a" To "d", "F", "f"
        If Left$(DE$, 2) = "0x" Then
        DE$ = DE$ & Mid$(a$, sng&, 1)
        End If
        Case "0" To "9"
        If sg1 Then
        ex$ = ex$ & Mid$(a$, sng&, 1)
        Else
        DE$ = DE$ & Mid$(a$, sng&, 1)
        End If
        Case "E", "e"
         If Left$(DE$, 2) = "0x" Then
         DE$ = DE$ & Mid$(a$, sng&, 1)
         Else
              If ex$ = vbNullString Then
               sg1 = True
        ex$ = "E"
        Else
        Exit Do
        End If
        End If
        Case "Ε", "ε"
 If ex$ = vbNullString Then
          sg1 = True
        ex$ = "E"
        Else
        Exit Do
        End If
        
        Case "+", "-"
        If sg1 And Len(ex$) = 1 Then
         ex$ = ex$ & Mid$(a$, sng&, 1)
        Else
        Exit Do
        End If
        Case Else
        Exit Do
        End Select
         sng& = sng& + 1
        Loop
        If Len(ex$) < 3 Then
                If ex$ = "E" Then
                ex$ = "0"
                sng = sng + 1
                ElseIf ex$ = "E-" Or ex$ = "E+" Then
                ex$ = "00"
                sng = sng + 2
                End If
                End If
    End If
    If ig$ = vbNullString Then
    IsNumberCheck = False
    Else

    If Left$(DE$, 2) = "0x" Then

            If Mid$(DE$, 3) = vbNullString Then
            r = 0
            Else
            DE$ = Right$("00000000" & Mid$(DE$, 3), 8)
            r = CDbl(UNPACKLNG(Left$(DE$, 4)) * 65536#) + CDbl(UNPACKLNG(Right$(DE$, 4)))
            End If
    Else
        If SG < 0 Then ig$ = "-" & ig$
                   On Error Resume Next
                        If ex$ <> "" Then
                        If Len(ex$) < 3 Then
                                If ex$ = "E" Then
                                ex$ = "0"
                                ElseIf ex$ = "E-" Or ex$ = "E+" Then
                                ex$ = "00"
                                End If
                                End If
                               If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = "."
                               If val(Mid$(ex$, 2)) > 308 Or val(Mid$(ex$, 2)) < -324 Then
                               
                                   r = val(ig$ & DE$)
                                   sng = sng - Len(ex$)
                                   ex$ = vbNullString
                                   
                               Else
                                   r = val(ig$ & DE$ & ex$)
                               End If
                           Else
                       If sng <= Len(a$) Then
            Select Case Asc(Mid$(a$, sng, 1))
            Case 64
                Mid$(a$, sng, 1) = " "
                If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = cdecimaldot$
                r = CDec(ig$ & DE$)
                If Err.Number = 6 Then
                Err.clear
                If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = "."
                r = val(ig$ & DE$)
                End If
            Case 35
            Mid$(a$, sng, 1) = " "
                If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = cdecimaldot$
                r = CCur(ig$ & DE$)
                If Err.Number = 6 Then
                Err.clear
                If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = "."
                r = val(ig$ & DE$)
                End If
           Case 37
                Mid$(a$, sng, 1) = " "
                If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = cdecimaldot$
                r = CInt(ig$)
                If Err.Number = 6 Then
                Err.clear
                r = val(ig$)
                End If
           Case 38
                Mid$(a$, sng, 1) = " "
                r = CLng(ig$)
                If Err.Number = 6 Then
                    Err.clear
                    r = val(ig$)
                End If
            Case 126
                Mid$(a$, sng, 1) = " "
                If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = cdecimaldot$
                r = CSng(ig$ & DE$)
                If Err.Number = 6 Then
                Err.clear
                If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = "."
                r = val(ig$ & DE$)
                End If
            Case Else
                If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = "."
                r = val(ig$ & DE$)
            End Select
            Else
            If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = "."
            r = val(ig$ & DE$)
            End If
                           End If
                     If Err.Number = 6 Then
                         If Len(ex$) > 2 Then
                             ex$ = Left$(ex$, Len(ex$) - 1)
                             sng = sng - 1
                             Err.clear
                             If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = "."
                             r = val(ig$ & DE$ & ex$)
                             If Err.Number = 6 Then
                                 sng = sng - Len(ex$)
                                 If DE$ <> vbNullString Then Mid$(DE$, 1, 1) = "."
                                  r = val(ig$ & DE$)
                             End If
                         End If
                       MyEr "Error in exponet", "Λάθος στον εκθέτη"
                       IsNumberCheck = False
                       Exit Function
                     End If
           
         End If
           a$ = Mid$(a$, sng&)
           IsNumberCheck = True
End If
End If
End Function
Function utf8encode(a$) As String
Dim bOut() As Byte, lPos As Long
If LenB(a$) = 0 Then Exit Function
bOut() = Utf16toUtf8(a$)
lPos = UBound(bOut()) + 1
If lPos Mod 2 = 1 Then
    utf8encode = StrConv(String$(lPos, Chr(0)), vbFromUnicode)
Else
    utf8encode = String$((lPos + 1) \ 2, Chr(0))
    End If
    CopyMemory ByVal StrPtr(utf8encode), bOut(0), LenB(utf8encode)
End Function
Function utf8decode(a$) As String
Dim b() As Byte, BLen As Long, WChars As Long
BLen = LenB(a$)
            ReDim b(0 To BLen - 1)
            CopyMemory b(0), ByVal StrPtr(a$), BLen
            WChars = MultiByteToWideChar(65001, 0, b(0), (BLen), 0, 0)
            utf8decode = space$(WChars)
            MultiByteToWideChar 65001, 0, b(0), (BLen), StrPtr(utf8decode), WChars
End Function
Sub test(a$)

Dim pos1 As Long
pos1 = 1

Debug.Print aheadstatus(a$, False)
'Debug.Print pos1
'Debug.Print Left$(a$, pos1)
End Sub
Public Function ideographs(c$) As Boolean
Dim code As Long
If c$ = vbNullString Then Exit Function
code = AscW(c$)  '
ideographs = (code And &H7FFF) >= &H4E00 Or (-code > 24578) Or (code >= &H3400& And code <= &HEDBF&) Or (code >= -1792 And code <= -1281)
End Function
Public Function nounder32(c$) As Boolean
nounder32 = AscW(c$) > 31 Or AscW(c$) < 0
End Function

Function GetImageX(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
Dim w1 As Long, s$, w2 As Long, pppp As mArray, useHandler As mHandler
GetImageX = False
If IsExp(bstack, a$, r) Then
      GetImageX = FastSymbol(a$, ")", True)
        If Not bstack.lastobj Is Nothing Then
           If TypeOf bstack.lastobj Is mHandler Then
              Set useHandler = bstack.lastobj
              Set bstack.lastobj = Nothing
              If useHandler.t1 = 2 Then
                  If useHandler.objref.ReadImageSizeX(r) Then
                  r = SG * bstack.Owner.ScaleX(r, 3, 1)
                          Set useHandler = Nothing
                      Exit Function
                  End If
              End If
           End If
        End If
            noImageInBuffer a$
            GetImageX = False
            r = 0#
    
Else
w1 = Abs(IsLabel(bstack, a$, s$))
        If w1 = 3 Then
            If GetVar(bstack, s$, w1) Then
                If Typename(var(w1)) <> "String" Then MissString: Exit Function
                If Left$(var(w1), 4) = "cDIB" And Len(var(w1)) > 12 Then
                    r = cDIBwidth1(var(w1)) * DXP
                    If SG < 0 Then r = -r
                    GetImageX = FastSymbol(a$, ")", True)
                Else
                    noImage a$
                    Exit Function
                End If
            Else
                MissFuncParameterStringVarMacro a$
            End If
        ElseIf w1 = 6 Then
            If neoGetArray(bstack, s$, pppp) Then
                If Not NeoGetArrayItem(pppp, bstack, s$, w2, a$) Then Exit Function
                If Not pppp.IsStringItem(w2) Then MissString: Exit Function
                Dim sV As Variant
                pppp.SwapItem w2, sV
          
                If Left$(sV, 4) = "cDIB" And Len(sV) > 12 Then
                    r = SG * cDIBwidth1(sV) * DXP
                    If SG < 0 Then r = -r
                    pppp.SwapItem w2, sV
                    GetImageX = FastSymbol(a$, ")", True)
                Else
                    pppp.SwapItem w2, sV
                    noImage a$
                End If
    
        Else
            MissParam a$
        End If
End If
End If
    
 
End Function
Function GetImageY(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
Dim w1 As Long, s$, w2 As Long, pppp As mArray, useHandler As mHandler
GetImageY = False
If IsExp(bstack, a$, r) Then
      GetImageY = FastSymbol(a$, ")", True)
        If Not bstack.lastobj Is Nothing Then
           If TypeOf bstack.lastobj Is mHandler Then
              Set useHandler = bstack.lastobj
              Set bstack.lastobj = Nothing
              If useHandler.t1 = 2 Then
                  If useHandler.objref.ReadImageSizeY(r) Then
                  r = SG * bstack.Owner.ScaleY(r, 3, 1)
                          Set useHandler = Nothing
                      Exit Function
                  End If
              End If
           End If
        End If
            noImageInBuffer a$
            GetImageY = False
            r = 0#
    
Else
w1 = Abs(IsLabel(bstack, a$, s$))
        If w1 = 3 Then
            If GetVar(bstack, s$, w1) Then
                If Typename(var(w1)) <> "String" Then MissString: Exit Function
                If Left$(var(w1), 4) = "cDIB" And Len(var(w1)) > 12 Then
                    r = cDIBheight1(var(w1)) * DXP
                    If SG < 0 Then r = -r
                    GetImageY = FastSymbol(a$, ")", True)
                Else
                    noImage a$
                    Exit Function
                End If
            Else
                MissFuncParameterStringVarMacro a$
            End If
        ElseIf w1 = 6 Then
            If neoGetArray(bstack, s$, pppp) Then
                If Not NeoGetArrayItem(pppp, bstack, s$, w2, a$) Then Exit Function
                If Not pppp.IsStringItem(w2) Then MissString: Exit Function
                Dim sV As Variant
                pppp.SwapItem w2, sV
          
                If Left$(sV, 4) = "cDIB" And Len(sV) > 12 Then
                    r = SG * cDIBheight1(sV) * DXP
                    If SG < 0 Then r = -r
                    pppp.SwapItem w2, sV
                    GetImageY = FastSymbol(a$, ")", True)
                Else
                    pppp.SwapItem w2, sV
                    noImage a$
                End If
    
        Else
            MissParam a$
        End If
End If
End If
    
 
End Function
Function GetImageXpixels(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
Dim w1 As Long, s$, w2 As Long, pppp As mArray, useHandler As mHandler
GetImageXpixels = False
If IsExp(bstack, a$, r) Then
      GetImageXpixels = FastSymbol(a$, ")", True)
        If Not bstack.lastobj Is Nothing Then
           If TypeOf bstack.lastobj Is mHandler Then
              Set useHandler = bstack.lastobj
              Set bstack.lastobj = Nothing
              If useHandler.t1 = 2 Then
                  If useHandler.objref.ReadImageSizeX(r) Then
                  r = SG * r
                          Set useHandler = Nothing
                      Exit Function
                  End If
              End If
           End If
        End If
            noImageInBuffer a$
            GetImageXpixels = False
            r = 0#
    
Else
w1 = Abs(IsLabel(bstack, a$, s$))
        If w1 = 3 Then
            If GetVar(bstack, s$, w1) Then
                If Typename(var(w1)) <> "String" Then MissString: Exit Function
                If Left$(var(w1), 4) = "cDIB" And Len(var(w1)) > 12 Then
                    r = cDIBwidth1(var(w1))
                    If SG < 0 Then r = -r
                    GetImageXpixels = FastSymbol(a$, ")", True)
                Else
                    noImage a$
                    Exit Function
                End If
            Else
                MissFuncParameterStringVarMacro a$
            End If
        ElseIf w1 = 6 Then
            If neoGetArray(bstack, s$, pppp) Then
                If Not NeoGetArrayItem(pppp, bstack, s$, w2, a$) Then Exit Function
                If Not pppp.IsStringItem(w2) Then MissString: Exit Function
                Dim sV As Variant
                pppp.SwapItem w2, sV
          
                If Left$(sV, 4) = "cDIB" And Len(sV) > 12 Then
                    r = SG * cDIBwidth1(sV)
                    If SG < 0 Then r = -r
                    pppp.SwapItem w2, sV
                    GetImageXpixels = FastSymbol(a$, ")", True)
                Else
                    pppp.SwapItem w2, sV
                    noImage a$
                End If
    
        Else
            MissParam a$
        End If
End If
End If
    
 
End Function
Function GetImageYpixels(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
Dim w1 As Long, s$, w2 As Long, pppp As mArray, useHandler As mHandler
GetImageYpixels = False
If IsExp(bstack, a$, r) Then
      GetImageYpixels = FastSymbol(a$, ")", True)
        If Not bstack.lastobj Is Nothing Then
           If TypeOf bstack.lastobj Is mHandler Then
              Set useHandler = bstack.lastobj
              Set bstack.lastobj = Nothing
              If useHandler.t1 = 2 Then
                  If useHandler.objref.ReadImageSizeY(r) Then
                  r = SG * r
                          Set useHandler = Nothing
                      Exit Function
                  End If
              End If
           End If
        End If
            noImageInBuffer a$
            GetImageYpixels = False
            r = 0#
    
Else
w1 = Abs(IsLabel(bstack, a$, s$))
        If w1 = 3 Then
            If GetVar(bstack, s$, w1) Then
                If Typename(var(w1)) <> "String" Then MissString: Exit Function
                If Left$(var(w1), 4) = "cDIB" And Len(var(w1)) > 12 Then
                    r = cDIBheight1(var(w1))
                    If SG < 0 Then r = -r
                    GetImageYpixels = FastSymbol(a$, ")", True)
                Else
                    noImage a$
                    Exit Function
                End If
            Else
                MissFuncParameterStringVarMacro a$
            End If
        ElseIf w1 = 6 Then
            If neoGetArray(bstack, s$, pppp) Then
                If Not NeoGetArrayItem(pppp, bstack, s$, w2, a$) Then Exit Function
                If Not pppp.IsStringItem(w2) Then MissString: Exit Function
                Dim sV As Variant
                pppp.SwapItem w2, sV
          
                If Left$(sV, 4) = "cDIB" And Len(sV) > 12 Then
                    r = SG * cDIBheight1(sV)
                    If SG < 0 Then r = -r
                    pppp.SwapItem w2, sV
                    GetImageYpixels = FastSymbol(a$, ")", True)
                Else
                    pppp.SwapItem w2, sV
                    noImage a$
                End If
    
        Else
            MissParam a$
        End If
End If
End If
    
 
End Function

Function enthesi(bstack As basetask, rest$) As String
'first is the string "label {0} other {1}
Dim counter As Long, pat$, final$, pat1$, pl1 As Long, pl2 As Long, pl3 As Long
Dim q$, p As Variant, p1 As Integer, pd$
If IsStrExp(bstack, rest$, final$) Then
  If FastSymbol(rest$, ",") Then
    Do
                pl2 = 1
                    pat$ = "{" + CStr(counter)
                   pat1$ = pat$ + ":"
                    pat$ = pat$ + "}"
                    If IsExp(bstack, rest$, p, , True) Then
                    If VarType(p) = vbBoolean Then q$ = Format$(p, DefBooleanString): GoTo fromboolean
again1:
                    pl2 = InStr(pl2, final$, pat1$)
                    If pl2 > 0 Then
                    pl1 = InStr(pl2, final$, "}")
                    If Mid$(final$, pl2 + Len(pat1$), 1) = ":" Then
                    p1 = 0
                    pl3 = val(Mid$(final$, pl2 + Len(pat1$) + 1) + "}")
                    Else
                    p1 = val("0" + Mid$(final$, pl2 + Len(pat1$)))
                    
                    pl3 = val(Mid$(final$, pl2 + Len(pat1$) + Len(Str$(p1))) + "}")
                    If p1 < 0 Then p1 = 13 '22
                    If p1 > 13 Then p1 = 13
                  p = MyRound(p, p1)
                  End If
                  pd$ = LTrim$(Str(p))
                  
                  If InStr(pd$, "E") > 0 Or InStr(pd$, "e") > 0 Then '' we can change e to greek ε
                  pd$ = Format$(p, "0." + String$(p1, "0") + "E+####")
                       If Not NoUseDec Then
                               If OverideDec Then
                                pd$ = Replace$(pd$, GetDeflocaleString(LOCALE_SDECIMAL), Chr(2))
                                pd$ = Replace$(pd$, GetDeflocaleString(LOCALE_STHOUSAND), Chr(3))
                                pd$ = Replace$(pd$, Chr(2), NowDec$)
                                pd$ = Replace$(pd$, Chr(3), NowThou$)
                                
                            ElseIf InStr(pd$, NowDec$) > 0 Then
                            pd$ = Replace$(pd$, NowDec$, Chr(2))
                            pd$ = Replace$(pd$, NowThou$, Chr(3))
                            pd$ = Replace$(pd$, Chr(2), ".")
                            pd$ = Replace$(pd$, Chr(3), ",")
                            
                            End If
                        End If
                  ElseIf p1 <> 0 Then
                   pd$ = Format$(p, "0." + String$(p1, "0"))
                           If Not NoUseDec Then
                            If OverideDec Then
                                pd$ = Replace$(pd$, GetDeflocaleString(LOCALE_SDECIMAL), Chr(2))
                                pd$ = Replace$(pd$, GetDeflocaleString(LOCALE_STHOUSAND), Chr(3))
                                pd$ = Replace$(pd$, Chr(2), NowDec$)
                                pd$ = Replace$(pd$, Chr(3), NowThou$)
                            ElseIf InStr(pd$, NowDec$) > 0 Then
                            pd$ = Replace$(pd$, NowDec$, Chr(2))
                            pd$ = Replace$(pd$, NowThou$, Chr(3))
                            pd$ = Replace$(pd$, Chr(2), ".")
                            pd$ = Replace$(pd$, Chr(3), ",")
                            
                            End If
                        End If
                  End If
               
                  If pl3 <> 0 Then
                    If pl3 > 0 Then
                        pd$ = Left$(pd$ + space$(pl3), pl3)
                        Else
                        pd$ = Right$(space$(Abs(pl3)) + pd$, Abs(pl3))
                        End If
                  End If
                        final$ = Replace$(final$, Mid$(final$, pl2, pl1 - pl2 + 1), pd$)
                        GoTo again1
                    Else
                    
                    If NoUseDec Then
                        final$ = Replace$(final$, pat$, CStr(p))
                    Else
                    pd$ = LTrim$(Str$(p))
                     If Left$(pd$, 1) = "." Then
                    pd$ = "0" + pd$
                    ElseIf Left$(pd$, 2) = "-." Then pd$ = "-0" + Mid$(pd$, 2)
                    End If
                    If OverideDec Then
                    final$ = Replace$(final$, pat$, Replace(pd$, ".", NowDec$))
                    Else
                    final$ = Replace$(final$, pat$, pd$)
                    End If
                    End If
                    
                    
                        End If
                        If Not FastSymbol(rest$, ",") Then Exit Do
                    
                    ElseIf IsStrExp(bstack, rest$, q$) Then
fromboolean:
                        final$ = Replace$(final$, pat$, q$)
AGAIN0:
                    pl2 = InStr(pl2, final$, pat1$)
                      If pl2 > 0 Then
                       pl1 = InStr(pl2, final$, "}")
                       pl3 = val(Mid$(final$, pl2 + Len(pat1$)) + "}")
                       If pl3 <> 0 Then
                    If pl3 > 0 Then
                        pd$ = Left$(q$ + space$(pl3), pl3)
                        Else
                        pd$ = Right$(space$(Abs(pl3)) + q$, Abs(pl3))
                        End If
                  End If
                        final$ = Replace$(final$, Mid$(final$, pl2, pl1 - pl2 + 1), pd$)
                        GoTo AGAIN0
                      End If
                        If Not FastSymbol(rest$, ",") Then Exit Do
                    Else
                        Exit Do
                    End If
                    counter = counter + 1
    Loop
    Else
    enthesi = EscapeStrToString(final$)
    Exit Function
    End If
End If
enthesi = final$
End Function

Public Function GetDeflocaleString(ByVal this As Long) As String
On Error GoTo 1234
    Dim Buffer As String, ret&, r&
    Buffer = String$(514, 0)
      
        ret = GetLocaleInfoW(0, this, StrPtr(Buffer), Len(Buffer))
    GetDeflocaleString = Left$(Buffer, ret - 1)
    
1234:
    
End Function
Function RevisionPrint(basestack As basetask, rest1$, xa As Long, Lang As Long) As Boolean
Dim Scr As Object, oldCol As Long, oldFTEXT As Long, oldFTXT As String, oldpen As Long
Dim par As Boolean, i As Long, F As Long, p As Variant, w4 As Boolean, pn&, s$, dlen As Long
Dim o As Long, w3 As Long, x1 As Long, y1 As Long, X As Double, ColOffset As Long
Dim work As Boolean, work2 As Long, skiplast As Boolean, ss$, ls As Long, myobject As Object, counter As Long, Counterend As Long, countDir As Long
Dim bck$, clearline As Boolean, ihavecoma As Boolean, isboolean As Boolean
Set Scr = basestack.Owner
Dim rest$, where As Long, final As Boolean

w3 = -1
Dim basketcode As Long, prive As basket
basketcode = GetCode(Scr)
prive = players(basketcode)
With prive
If .MAXXGRAPH = 0 Then MyEr "No form to print", "δεν υπάρχει φόρμα για εκτύπωση": Exit Function
PlaceBasketPrive Scr, prive
Scr.FontTransparent = True
On Error GoTo 0
Dim opn&
where = 1
aheadstatusANY rest1$, where
rest$ = Left$(rest1$, where - 1)


par = True
If MaybeIsSymbol3(rest$, "#", F) Then
   If Mid$(rest$, F + 1, 6) Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]" Then
   
   Else
  Mid$(rest$, 1, F) = space$(F)
        If IsExp(basestack, rest$, p, , True) Then
        If p < 0 Then
        If p < -1 Then
        
                     .lastprint = False
                     par = False
        End If
        F = p
                     If Not FastSymbol(rest$, ",") Then
                     s$ = vbNullString
                     pn& = 2
                     GoTo isAstring
                     End If
        Else
                     F = CLng(MyMod(p, 512))
                     If FKIND(F) = FnoUse Or FKIND(F) = Finput Or FKIND(F) = Frandom Then MyEr "Wrong File Handler", "Λάθος Χειριστής Αρχείου": RevisionPrint = False: GoTo exit2
                     Dim clearprive As basket
                     prive = clearprive
                     .lastprint = False
                     par = False
                     If Not FastSymbol(rest$, ",") Then
                     s$ = vbNullString
                     pn& = 2
                     GoTo isAstring
                     End If
                     
                     
            End If
             
       Else
       MyEr "expected file number", "περίμενα αριθμό αρχείου"
       End If
    End If
Else
ss$ = Left$(rest$, MyTrimL(rest$) + 5)
ls = Len(ss$)
If Not IsLabelSYMB3(ss$, s$) Then
                    F = 0
                    
Else
Select Case Lang
Case 1
If Len(s$) > 3 Then
If InStr("BOUP", UCase(Left$(s$, 1))) > 0 Then

Select Case UCase(s$)
        Case "BACK"
        Mid$(rest$, 1, ls - Len(ss$)) = space$(ls - Len(ss$))
        F = 4
        Case "OVER"
        F = 1
        Mid$(rest$, 1, ls - Len(ss$)) = space$(ls - Len(ss$))
        Case "UNDER"
        F = 2
         Mid$(rest$, 1, ls - Len(ss$)) = space$(ls - Len(ss$))
        Case "PART"
        F = 3
        Mid$(rest$, 1, ls - Len(ss$)) = space$(ls - Len(ss$))
        Case Else
        ''rest$ = s$ + rest$
        F = 0
        End Select
        Else
        F = 0
        End If
        Else
        F = 0
End If
Case 0, 2
If Len(s$) > 2 Then
If InStr("ΦΠΥΜ", myUcase(Left$(s$, 1))) > 0 Then
        Select Case myUcase(s$, True)
        Case "ΦΟΝΤΟ"
        Mid$(rest$, 1, ls - Len(ss$)) = space$(ls - Len(ss$))
        F = 4
        Case "ΠΑΝΩ"
        Mid$(rest$, 1, ls - Len(ss$)) = space$(ls - Len(ss$))
        F = 1
        Case "ΥΠΟ"
        Mid$(rest$, 1, ls - Len(ss$)) = space$(ls - Len(ss$))
        F = 2
        Case "ΜΕΡΟΣ"
        Mid$(rest$, 1, ls - Len(ss$)) = space$(ls - Len(ss$))
        F = 3
        Case Else
        F = 0
        End Select
        Else
        F = 0
        End If
        Else
        F = 0
End If
Case -1   '' this is for ?
If Len(s$) > 2 Then
If InStr("BOUPΦΠΥΜ", myUcase(Left$(s$, 1))) > 0 Then
Select Case myUcase(s$)
        Case "ΦΟΝΤΟ", "BACK"
        Mid$(rest$, 1, ls - Len(ss$)) = space$(ls - Len(ss$))
        F = 4
        Case "ΠΑΝΩ", "OVER"
        Mid$(rest$, 1, ls - Len(ss$)) = space$(ls - Len(ss$))
        F = 1
        Case "ΥΠΟ", "UNDER"
        Mid$(rest$, 1, ls - Len(ss$)) = space$(ls - Len(ss$))
        F = 2
        Case "ΜΕΡΟΣ", "PART"
        Mid$(rest$, 1, ls - Len(ss$)) = space$(ls - Len(ss$))
        F = 3
        Case Else
        F = 0
        End Select
        Else
        F = 0
        End If
        Else
        F = 0
        End If
        Lang = 0
        End Select
        
        If F > 0 And .lastprint Then
        .lastprint = False
        
        GetXYb Scr, prive, x1&, y1&
        If F <> 2 Then If x1& > 0 Or y1& >= .mx Then crNew basestack, prive
        End If
If F = 1 Then  ''
    work = True
    oldCol = .Column
    Scr.Line (0&, .currow * .Yt)-((.mx - 1) * .Xt + .Xt * 2, (.currow) * .Yt + .Yt - 2 * DYP), .Paper, BF
    LCTbasket Scr, prive, .currow, 0&
    .Column = .mx - 1
    w4 = True
    oldFTEXT = .FTEXT
    oldFTXT = .FTXT
    oldpen = .mypen
    pn& = 2
    .FTEXT = 4
ElseIf F = 2 Then
    work = True
    oldCol = .Column
    Scr.Line (0&, (.currow) * .Yt + .Yt - DYP)-((.mx - 1) * .Xt + .Xt * 2, (.currow) * .Yt + .Yt - DYP), .mypen, BF
    crNew basestack, prive
    LCTbasketCur Scr, prive
    w4 = True
    oldFTEXT = .FTEXT
    oldFTXT = .FTXT
    oldpen = .mypen
    .FTEXT = 4
    pn& = 2
ElseIf F = 3 Then
' we print in a line with lost chars, so controling the start of printing
' we can render text, like from a view port Some columns are hidden because they went out of screen;
work = True
oldCol = .Column
LCTbasket Scr, prive, .currow, 0&
w4 = True
oldFTEXT = .FTEXT
oldFTXT = .FTXT
.FTEXT = 4
oldpen = .mypen
ElseIf F = 4 Then
    work = True
    clearline = True
    ' LCTbasketCur scr, prive
    If .curpos > 0 Then
    crNew basestack, prive
    LCTbasketCur Scr, prive
    End If
    Scr.Line (0&, .currow * .Yt)-((.mx - 1) * .Xt + .Xt * 2, (.currow) * .Yt + .Yt - 1 * DYP), .Paper, BF
   ' scr.Line (0&, (.currow) * .Yt + .Yt - DYP)-((.mx - 1) * .Xt + .Xt * 2, (.currow) * .Yt + .Yt - 1 * DYP), .mypen, BF
    LCTbasketCur Scr, prive
    pn& = 2
End If

F = 0
End If

End If
If w4 Then pn& = 2 Else pn& = 0

s$ = vbNullString
If par Then
    If .FTEXT > 3 And .curpos >= .mx And Not w4 Then
    crNew basestack, prive
    w3 = 0
End If
End If
If par Then
If FastSymbol(rest$, ";") Then

            If .lastprint Then
            .lastprint = False
            LCTbasketCur Scr, prive
            crNew basestack, prive
            End If
         
ElseIf .lastprint Then
If .FTEXT > 3 Then pn& = 7: GoTo newntrance

End If
End If


Do
             If final Then
             If myobject Is Nothing Then GoTo there1
            End If
   If FastSymbol(rest$, "~(", , 2) Then ' means combine
        ' get the color and then look for @( parameters)
        w3 = -1
    If par Then  ' par is false when we print in files, we can't use color;
   
                 If IsExp(basestack, rest$, p, , True) Then .mypen = CLng(mycolor(p))
                 TextColor Scr, .mypen
                 
                     If FastSymbol(rest$, ",") Then
                     
                                If w4 Or Not work Then
                                  If prive.lastprint Then
                                   prive.lastprint = False
                                   GetXYb Scr, prive, .curpos, .currow
                                                   If work Then
                       .curpos = .curpos - ColOffset
                      If (.curpos Mod (.Column + 1)) <> 0 Then
                      .curpos = .curpos + (.Column + 1) - (.curpos Mod (.Column + 1)) + ColOffset
                      Else
                       .curpos = .curpos + ColOffset
                      End If
                 If w4 Then LCTbasketCur Scr, prive
                       End If
                                  End If
                               
                              LCTbasketCur Scr, prive
                             
                                Else
                                 If work Then
                       .curpos = .curpos - ColOffset
                      If (.curpos Mod (.Column + 1)) <> 0 Then
                      .curpos = .curpos + (.Column + 1) - (.curpos Mod (.Column + 1)) + ColOffset
                      Else
                       .curpos = .curpos + ColOffset
                      End If
                 If w4 Then LCTbasketCur Scr, prive
                       End If
                               '' LCTbasketCur scr, prive
                                End If
                                
                                
'                                         GetXYb scr, prive, .curpos, .currow
                   ''  LCTbasketCur scr, prive
                x1 = .Column + .curpos + 1
                y1 = .currow + 1
                
                                pn& = 99
                             GoTo pthere   ' background and border and or images
            
            
                 End If
                         If Not FastSymbol(rest$, ")") Then RevisionPrint = False: Set Scr = Nothing: GoTo exit2
                         pn& = 99
    End If
    ElseIf FastSymbol(rest$, "@(", , 2) Then
    clearline = False
    w3 = -1
               'If Not par Then RevisionPrint = False: Set scr = Nothing: Exit Function
                If IsExp(basestack, rest$, p, , True) Then

                If par Then .curpos = CLng(Fix(p))
                End If
                
                If FastSymbol(rest$, ",") Then
                If IsExp(basestack, rest$, p, , True) Then
                If CLng(Fix(p)) >= .My Then
                If par Then .currow = .My - 1
                Else
                If par Then .currow = CLng(Fix(p))
                End If
                End If
                End If

                If FastSymbol(rest$, ",") Then
                
                If IsExp(basestack, rest$, p, , True) Then x1 = CLng(Fix(p))
                Else
                x1 = 1
                End If
                
                If FastSymbol(rest$, ",") Then
                If IsExp(basestack, rest$, p, , True) Then y1 = CLng(Fix(p))
                Else
                y1 = 1
                End If

                If FastSymbol(rest$, ",") Then
             '   On Error Resume Next
pthere:
                   
                If par Then LCTbasketCur Scr, prive
                If IsStrExp(basestack, rest$, s$) Then
                p = 0
                    If FastSymbol(rest$, ",") Then
                        If IsExp(basestack, rest$, p, , True) Then
                            If p <> 0 Then p = True
                        Else
                        p = True
                        End If
                    End If
             
                    x1 = Abs(x1 - .curpos)
                    y1 = Abs(y1 - .currow)
                    
                    If par Then BoxImage Scr, prive, x1, y1, s$, 0, (p)
                    'If P <> 0 Then .currow = y1 + .currow
                ElseIf IsExp(basestack, rest$, p, , True) Then
         
                    If par Then BoxColorNew Scr, prive, x1 - 1, y1 - 1, (p)
                    If FastSymbol(rest$, ",") Then
                        If IsExp(basestack, rest$, X, , True) Then
                            If par Then BoxBigNew Scr, prive, x1 - 1, y1 - 1, (X)
                            
                            
                            
                        Else
                            RevisionPrint = False
                            Set Scr = Nothing
                            GoTo exit2
                        End If
                    End If
                Else
                    RevisionPrint = False
                    Set Scr = Nothing
                    GoTo exit2
                
                End If

                End If
             If par Then LCTbasket Scr, prive, .currow, .curpos
                
        If Not FastSymbol(rest$, ")") Then
        RevisionPrint = False
        Set Scr = Nothing
        GoTo exit2
        End If
        work = False
        pn& = 99
        ElseIf LastErNum <> 0 Then
      RevisionPrint = LastErNum = -2
      Set Scr = Nothing
    GoTo exit2
    
    ElseIf FastSymbol(rest$, "$(", , 2) Then
conthere:
w3 = -1
        If IsExp(basestack, rest$, p, , True) Then
        If Not par Then p = 0
            .FTEXT = Abs(p) Mod 10
            ' 0 STANDARD LEFT chars before typed beyond the line are directed to the next line
            ' 1  RIGHT
            ' 2 CENTER
            ' 3 LEFT
            ' 4 LEFT PROP....expand to next .Column......
            ' 5 RIGHT PROP
            ' 6 CENTER PROP
            ' 7 LEFT PROP
            ' 8 left and right justify
            ' 9 New in version 8 Left justify(like 7) without word wrap (cut excess)
        ElseIf IsStrExp(basestack, rest$, s$) Then
            .FTXT = s$
        End If
        
        
        If FastSymbol(rest$, ",") Then
                If IsExp(basestack, rest$, p, , True) Then
                    If par Then
                        p = p - 1
                        If Abs(Int(p Mod (.mx + 1))) < 2 Then
                            MyEr ".Column minimum width is 4 chars", "Μικρότερο μέγεθος στήλης είναι οι τέσσερις χαρακτήρες"
                        Else
                            If w4 Or Not work Then
                                LCTbasketCur Scr, prive
                            Else
                                GetXYb Scr, prive, .curpos, .currow
                            End If
                            If w4 Then ColOffset = .curpos    ' now we have columns from offset ColOffset
                                .Column = Abs(Int(p Mod (.mx + 1)))
                            End If
                    End If
                    
                Else
                    RevisionPrint = False
                    Set Scr = Nothing
                    GoTo exit2
                End If
         End If
      
            If Not FastSymbol(rest$, ")") Then
            RevisionPrint = False
            Set Scr = Nothing
            GoTo exit2
            End If
        
        
        If par Then pn& = 99
        ElseIf LastErNum <> 0 Then
       RevisionPrint = LastErNum <> -2
       Set Scr = Nothing
    GoTo exit2
    ElseIf Not myobject Is Nothing Then
takeone:
    '' for arrays only
    If countDir >= 0 Then
    
    If counter = myobject.count Or (counter > Counterend And Counterend > -1) Or countDir = 0 Then
        Set myobject = Nothing
              SwapStrings rest$, bck$
            '  rest$ = bck$
            ' bck$ = vbNullString
        GoTo taketwo
    End If
Else
        If counter < 0 Or (counter < Counterend And Counterend > -1) Then
        Set myobject = Nothing
            SwapStrings rest$, bck$
             ' rest$ = bck$
            ' bck$ = vbNullString
        GoTo taketwo
    End If
    End If
    
    myobject.index = counter
    If myobject.IsEmpty Then
        s$ = " "
        counter = counter + countDir
        GoTo isAstring
    Else
            If Not IsNumeric(myobject.Value) Then
                If myobject.IsObj Then
                    If myobject.IsEnum(p) Then
                    counter = counter + countDir
                    GoTo isanumber
                    Else
                    s$ = " "
                    End If
                Else
                    s$ = myobject.Value
                End If

                
                counter = counter + countDir
                GoTo isAstring
            Else
                If TypeOf myobject Is Enumeration Then
                p = myobject.Value
                Else
                On Error Resume Next
                If Not myobject.IsEnum(p) Then
                p = myobject.Value
                End If
                If Err.Number > 0 Then p = myobject.Value
                End If
                counter = counter + countDir
                GoTo isanumber
            End If
    End If

    
    ElseIf IsExp(basestack, rest$, p) Then
            If Not basestack.lastobj Is Nothing Then
                
                If Typename(basestack.lastobj) = myArray Then
                Set myobject = basestack.lastobj
                Set basestack.lastobj = Nothing
                Counterend = -1
                counter = 0
                countDir = 1
                bck$ = vbNullString
                SwapStrings rest$, bck$
                'bck$ = rest$
                'rest$ = vbNullString
                GoTo takeone
                ElseIf Typename(basestack.lastobj) = "mHandler" Then
                    Set myobject = basestack.lastobj
                    Set basestack.lastobj = Nothing
                    With myobject
                    If myobject.UseIterator Then
                        If TypeOf myobject.objref Is Enumeration Then
                        p = myobject.objref.Value
                        Set myobject = Nothing
                        GoTo isanumber
                        Else
                        Counterend = myobject.index_End
                        If Counterend = -1 Then
                        Set myobject = Nothing
                        GoTo isAstring
                        Else
                        counter = myobject.index_start
                        
                        
                        If counter <= Counterend Then countDir = 1 Else countDir = -1
                        End If
                        End If
                    Else
                        Counterend = -1
                        counter = 0
                        countDir = 1
                       If myobject.t1 = 4 Then
                       
                       ' p = myobject.index_cursor * myobject.Sign
                   
                    Set myobject = Nothing
                    GoTo isanumber
                        ElseIf Not CheckIsmArrayOrStackOrCollection(myobject) Then
                            Set myobject = Nothing
                            
                        Else
                                SwapStrings rest$, bck$
                                'bck$ = rest$
                                'rest$ = vbNullString
                                GoTo takeone
                        End If
                        
                    End If
                    End With
                    If Not CheckLastHandler(myobject) Then
                    NoProperObject
                    rest$ = bck$: RevisionPrint = False: GoTo exit2
                    End If

                    If Typename(myobject.objref) = "FastCollection" Then
                             Set myobject = myobject.objref
                             SwapStrings rest$, bck$
                        'bck$ = rest$
                        'rest$ = vbNullString
                        GoTo takeone
                    ElseIf Typename(myobject.objref) = "mStiva" Then
                        Set myobject = myobject.objref
                        SwapStrings rest$, bck$
                        'bck$ = rest$
                        'rest$ = vbNullString
                        GoTo takeone

                    ElseIf Typename(myobject.objref) = myArray Then
                        If myobject.objref.Arr Then
                            Set myobject = myobject.objref
                            SwapStrings rest$, bck$
                        'bck$ = rest$
                        'rest$ = vbNullString
                        GoTo takeone
                        End If
                   ElseIf Typename(myobject.objref) = "Enumeration" Then
                 
                            Set myobject = myobject.objref
                 
                            SwapStrings rest$, bck$
                       GoTo takeone
                        End If
                ElseIf TypeOf basestack.lastobj Is VarItem Then
                    p = basestack.lastobj.ItemVariant
                End If
                Set basestack.lastobj = Nothing
            ElseIf VarType(p) = vbBoolean Then
            isboolean = True
            End If
isanumber:
        If par Then
            If .lastprint Then opn& = 5

            pn& = 1
            If .Column = 1 Then
            
            pn& = 6
            End If
            Else
            .lastprint = False
            pn& = 1
           End If
    ElseIf LastErNum <> 0 Then
            .lastprint = False
            RevisionPrint = LastErNum = -2
            Set Scr = Nothing
            GoTo exit2
    ElseIf IsStrExp(basestack, rest$, s$, Len(basestack.tmpstr) = 0) Then
    ' special not good for every day...is 255 char in greek codepage
   '   If InStr(s$, ChrW(&HFFFFF8FB)) > 0 Then s$ = Replace(s$, ChrW(&HFFFFF8FB), ChrW(&H2007))
     If Not basestack.lastobj Is Nothing Then
                If Typename(basestack.lastobj) = myArray Then
                Set myobject = basestack.lastobj
                Set basestack.lastobj = Nothing
                Counterend = -1
                counter = 0
                countDir = 1
                SwapStrings rest$, bck$
                'bck$ = rest$
                'rest$ = vbNullString
                GoTo takeone
                End If
                
            End If
isAstring:
If par Then

    If .lastprint Then opn& = 5
            pn& = 2
            
      If .Column = 1 Then
            
            pn& = 7
            End If
            Else
             .lastprint = False
            pn& = 2
            End If
    ElseIf LastErNum <> 0 Then
             RevisionPrint = LastErNum = -2
             Set Scr = Nothing
                GoTo exit2
    Else
there1:
    If pn& <> 0 And pn& < 5 And Not .lastprint Then
        If par Then
            If Not w4 Then
          '' GetXYb scr, prive, .curpos, .currow

If Not (.curpos = 0) Then
GetXYb Scr, prive, .curpos, .currow
If pn& = 1 Then
crNew basestack, prive: skiplast = True
ElseIf pn& = 2 Then

If Abs(w3) = 1 And .curpos = 0 And Not (.FTEXT = 9 Or .FTEXT = 5 Or .FTEXT = 6) Then
If .FTEXT = 7 Then
crNew basestack, prive: skiplast = True
End If
Else
crNew basestack, prive: skiplast = True
End If
End If
End If


            End If
        Else
        If F < 0 Then
            crNew basestack, prive
        ElseIf uni(F) Then
            putUniString F, vbCrLf
            Else
            putANSIString F, vbCrLf
            'Print #f,
            End If
        End If
    End If
 
        Exit Do
    End If
conthere2:
If .lastprint And opn& > 4 Then .lastprint = False
    If FastSymbol(rest$, ";") Then
'' LEAVE W3
If par Then
   If opn& = 0 And (Not work) And (Not .lastprint) Then

   LCTbasket Scr, prive, .currow, .curpos
   End If
  
   ' IF  WORK THEN opn&=5
   opn& = 5
  End If
newntrance:
work = True
.lastprint = True
        
         Do While FastSymbol(rest$, ";")
         Loop
    ElseIf Not FastSymbol(rest$, ",") Then
    
    pn& = pn& + opn&
  opn& = 0
  rest$ = NLtrim$(rest$)
   '  final = True   ERROR - WHEN myobject is an array/inventory to iterate
    final = myobject Is Nothing
    Else
If par Then
ihavecoma = True ' 'rest$ = "," & rest$
End If
    End If
    pn& = pn& + opn&
    Select Case pn&
    Case 0
    Exit Do
    Case 1
        If .FTXT = vbNullString Then
        If xa Then
        s$ = PACKLNG2$(p)
        Else
If NoUseDec Then
    If isboolean Then
        If ShowBooleanAsString Then
            If cLid = 1032 Then
                s$ = Format$(p, ";\Α\λ\η\θ\έ\ς;\Ψ\ε\υ\δ\έ\ς")
            ElseIf cLid = 1033 Then
                s$ = Format$(p, ";\T\r\u\e;\F\a\l\s\e")
            Else
                s$ = Format$(p, DefBooleanString)
            End If
            isboolean = False
            GoTo contboolean2
        Else
            p = p * 1
            isboolean = False
            GoTo contboolean
        End If
    Else
        On Error Resume Next
        s$ = CStr(p)
        If Err.Number > 0 Then
                If Typename(p) = "Null" Then
                    s$ = "NULL"
                    Err.clear
                Else
                    s$ = Typename(p)
                    Err.clear
                End If
        End If
        On Error GoTo 0
    End If
    
Else

If isboolean Then
    If ShowBooleanAsString Then
        If cLid = 1032 Then
            s$ = Format$(p, ";\Α\λ\η\θ\έ\ς;\Ψ\ε\υ\δ\έ\ς")
        ElseIf cLid = 1033 Then
            s$ = Format$(p, ";\T\r\u\e;\F\a\l\s\e")
        Else
            s$ = Format$(p, DefBooleanString)
        End If
        isboolean = False
        GoTo contboolean2
    Else
        p = p * 1
        isboolean = False
        GoTo contboolean
    End If
Else
contboolean:
On Error Resume Next
 s$ = LTrim$(Str(p))
 If Err Then s$ = Typename$(p): Err.clear
    If Left$(s$, 1) = "." Then
    s$ = "0" + s$
    ElseIf Left$(s$, 2) = "-." Then s$ = "-0" + Mid$(s$, 2)
 End If
 
 If OverideDec Then s$ = Replace$(s$, ".", NowDec$)
 End If
End If
contboolean2:
      If .FTEXT < 4 Then
            If InStr(s$, ".") > 0 Then
                 If InStr(s$, ".") <= .Column Then
                        If RealLen(s$) > .Column + 1 Then
                                 If .FTEXT > 0 Then s$ = Left$(s$, .Column + 1)
                        End If
                End If
            ElseIf .FTEXT > 0 Then
                 If RealLen(s$) > .Column + 1 Then s$ = String$(.Column, "?")
            End If
          End If
    End If
        Else
        s$ = Format$(p, .FTXT)
        If Not NoUseDec Then
            If OverideDec Then
                s$ = Replace$(s$, GetDeflocaleString(LOCALE_SDECIMAL), Chr(2))
                s$ = Replace$(s$, GetDeflocaleString(LOCALE_STHOUSAND), Chr(3))
                s$ = Replace$(s$, Chr(2), NowDec$)
                s$ = Replace$(s$, Chr(3), NowThou$)
                
            ElseIf InStr(s$, NowDec$) > 0 And InStr(.FTXT, ".") > 0 Then
                ElseIf InStr(s$, NowDec$) > 0 Then
                s$ = Replace$(s$, NowDec$, Chr(2))
                s$ = Replace$(s$, NowThou$, Chr(3))
                s$ = Replace$(s$, Chr(2), ".")
                s$ = Replace$(s$, Chr(3), ",")
            
            End If
            End If
        End If
     If par Then
        If .Column > 2 Then   ' .Column 3 means 4 chars width
        If opn& < 5 Then
    '                    ensure that we are align in .Column  (.Column is based zero...)
    skiplast = False
               If .currow >= .My Then
               If Not w4 Then crNew basestack, prive: skiplast = True
               End If
        
                        If work Then
                       .curpos = .curpos - ColOffset
                      If (.curpos Mod (.Column + 1)) <> 0 Then
                      .curpos = .curpos + (.Column + 1) - (.curpos Mod (.Column + 1)) + ColOffset
                      Else
                       .curpos = .curpos + ColOffset
                      End If
                 If w4 Then LCTbasketCur Scr, prive
                       End If
                       work = True
    End If
            If .curpos >= .mx Then
    '' ???
                    Else
            If clearline And .curpos = 0 Then Scr.Line (0&, .currow * .Yt)-((.mx - 1) * .Xt + .Xt * 2, (.currow) * .Yt + .Yt - 1 * DYP), .Paper, BF
            Select Case .FTEXT
            Case 0
            
                          
                       PlainBaSket Scr, prive, space$(.Column - (RealLen(s$) - 1) Mod (.Column + 1)) + s$, w4, w4, , clearline
                       
            Case 3
                        PlainBaSket Scr, prive, Right$(space$(.Column - (RealLen(s$) - 1) Mod (.Column + 1)) + Left$(s$, .Column + 1), .Column + 1), w4, w4, , clearline
            Case 2
                        If RealLen(s$) > .Column + 1 Then s$ = "????"
                        PlainBaSket Scr, prive, Left$(space$((.Column + 1 - RealLen(s$)) \ 2) + Left$(s$, .Column + 1) & space$(.Column), .Column + 1), w4, w4, , clearline
            Case 1
                        PlainBaSket Scr, prive, Left$(s$ & space$(.Column), .Column + 1), w4, w4, , clearline
            Case 5
                        x1 = .curpos
                        y1 = .currow
                        If Not (.mx - 1 <= .curpos And w4 <> 0) Then
                        LCTbasketCur Scr, prive
                        Scr.CurrentX = Scr.CurrentX + (.Xt - TextWidth(Scr, Left$(s$, 1))) \ 2
                        wwPlain basestack, prive, s$, .Column * .Xt + .Xt - (.Xt - TextWidth(Scr, Left$(s$, 1))) \ 2, 0, , True, 0, , CBool(w4), True, , True
                        .currow = y1
        

                        .curpos = x1 + .Column + 1
                        
                        End If
                     If .curpos >= .mx And Not w4 Then
                   
                         .currow = .currow + 1
                         .curpos = 0
                         End If
              If .lastprint Then
     
                 If .curpos = 0 Then
                 If .currow >= .My Then crNew basestack, prive Else LCTbasketCur Scr, prive
                 End If
                 
     Scr.CurrentX = .curpos * .Xt
                
                  Scr.CurrentY = .currow * .Yt + .uMineLineSpace
             
         
                   End If
            Case 4, 7, 8
                         wwPlain basestack, prive, s$ & vbCrLf, .Column * .Xt + .Xt - (.Xt - TextWidth(Scr, Right$(s$, 1))) \ 2, 0, , , 1, , , pn& < 5, , True
                        .curpos = .curpos + .Column + 1
                        If .curpos >= .mx And Not w4 Then
                                .curpos = 0
                                .currow = .currow + 1

                        End If
                        If .lastprint Then
                            If .curpos = 0 Then
                                If .currow >= .My Then
                                crNew basestack, prive
                                
                             
                              
                                Else
                                LCTbasketCur Scr, prive

                                End If
                            End If
                            If .curpos > 0 Then Scr.CurrentX = .curpos * .Xt - (.Xt - TextWidth(Scr, Right$(s$, 1))) \ 2 Else Scr.CurrentX = .curpos * .Xt
                            Scr.CurrentY = .currow * .Yt + .uMineLineSpace
                        End If
            Case 6
                            
                        wwPlain basestack, prive, s$, .Column * .Xt + .Xt, 0, , False, 2, , , pn& < 5, , True
                        .curpos = .curpos + .Column + 1
                        If .curpos >= .mx And Not w4 Then
                            .curpos = 0
                            .currow = .currow + 1
                        End If
                        If .lastprint Then
                            If .curpos = 0 Then
                                If .currow >= .My Then crNew basestack, prive Else LCTbasketCur Scr, prive
                            End If
                            Scr.CurrentX = .curpos * .Xt
                            Scr.CurrentY = .currow * .Yt + .uMineLineSpace
                        End If
                            
            Case 9
                            LCTbasketCur Scr, prive
                            wPlain Scr, prive, s$, 1000, 0, True
                             GetXYb Scr, prive, .curpos, .currow
                           .curpos = .curpos + 1
                            If (.curpos Mod (.Column + 1)) <> 0 Then
                     .curpos = .curpos + (.Column + 1) - (.curpos Mod (.Column + 1)) + ColOffset
                      Else
                       .curpos = .curpos + ColOffset
                      End If
                             '     .curpos = .curpos + .Column + 1
                            If .curpos >= .mx And Not w4 Then
                                .curpos = 0
                                .currow = .currow + 1
                            End If
                                                               If .lastprint Then
     
                 If .curpos = 0 Then
                 If .currow >= .My Then crNew basestack, prive Else LCTbasketCur Scr, prive
                 End If
                If .curpos > 0 Then Scr.CurrentX = .curpos * .Xt - (.Xt - TextWidth(Scr, Right$(s$, 1))) \ 2 Else Scr.CurrentX = .curpos * .Xt
                  Scr.CurrentY = .currow * .Yt + .uMineLineSpace
             
         
                   End If
            End Select
End If
            
            
            
        Else
        ' no way to use this any more...7 rev 20
        PlainBaSket Scr, prive, s$
        End If
 
        Else
          If F < 0 Then
            PlainBaSket Scr, prive, s$
        ElseIf uni(F) Then
            putUniString F, s$
            Else
            putANSIString F, s$
        'Print #f, S$;
        End If
        End If
    Case 2
    '' for string.....................................................................................................................
        If .FTXT <> "" Then
        s$ = Format$(s$, .FTXT)
        End If
        If par Then
        If .Column > 0 Then
                             x1 = .curpos: y1 = .currow
                skiplast = False
                                If .currow >= .My And Not w4 Then
                                crNew basestack, prive
                                skiplast = True
                                End If
                        If work Then
                       .curpos = .curpos - ColOffset
                      If (.curpos Mod (.Column + 1)) <> 0 Then
                      .curpos = .curpos + (.Column + 1) - (.curpos Mod (.Column + 1)) + ColOffset
                      Else
                       .curpos = .curpos + ColOffset
                     
                      End If
                      '' LCTbasket scr, prive,   y1, X1
                       If w4 Then LCTbasketCur Scr, prive
                       End If
                       work = True
          If s$ = vbNullString Then s$ = " "
          
                 If .curpos >= .mx Then
                 y1 = 1
                    Else
                               If clearline And .curpos = 0 Then Scr.Line (0&, .currow * .Yt)-((.mx - 1) * .Xt + .Xt * 2, (.currow) * .Yt + .Yt - 1 * DYP), .Paper, BF

            Select Case .FTEXT
                Case 1
                           '' GetXY scr, X1, y1
                          ''  If s$ = VbNullString Then s$ = " "
                          dlen = RealLen(s$)
                          PlainBaSket Scr, prive, Left$(s$ & space$(Len(s$) - dlen + .Column - (dlen - 1) Mod (.Column + 1)), .Column + 1 + Len(s$) - dlen), w4, w4, , clearline
                Case 2
                            dlen = RealLen(s$)
                            If dlen > (.Column + 1 + Len(s$) - dlen) Then s$ = Left$(s$, .Column + 1 + Len(s$) - dlen):  dlen = RealLen(s$)
                            
                            PlainBaSket Scr, prive, Left$(space$((.Column + 1 + Len(s$) - dlen - dlen) \ 2) + s$ & space$(.Column), .Column + 1 + Len(s$) - dlen), w4, w4, , clearline
                Case 3
                            dlen = RealLen(s$)
                            PlainBaSket Scr, prive, Right$(space$(.Column + Len(s$) - dlen - (dlen - 1) Mod (.Column + 1)) & s$, .Column + 1 + Len(s$) - dlen), w4, w4, , clearline
                Case 0
                           '' If s$ = VbNullString Then s$ = " "
                        
                            PlainBaSket Scr, prive, s$ + space$(.Column - (RealLen(s$) - 1) Mod (.Column + 1)), w4, w4, , clearline
                       
                Case 4
                            
                            LCTbasketCur Scr, prive
                            Scr.CurrentX = Scr.CurrentX + (.Xt - TextWidth(Scr, Left$(s$, 1))) \ 2
                            
                            w3 = 0
                            wwPlain basestack, prive, s$, Scr.Width, 0, , True, 0, , w3, True
                            w3 = w3 \ .Xt + 1
                            ' go to next .Column...
                            
                            .curpos = (.Column + 1) * ((w3 + .Column + 1) \ (.Column + 1))
                        If .curpos >= .mx And Not w4 Then
                                .curpos = 0
                                .currow = .currow + 1
                            End If
                Case 5
                           '' GetXY scr, X1, y1
                            LCTbasketCur Scr, prive
                            Scr.CurrentX = Scr.CurrentX + (.Xt - TextWidth(Scr, Left$(s$, 1))) \ 2
                            wwPlain basestack, prive, s$, .Column * .Xt + .Xt - (.Xt - TextWidth(Scr, Left$(s$, 1))) \ 2, 0, , True, 3, , , True
                            .curpos = .curpos + .Column + 1
                            If .curpos >= .mx And Not w4 Then
                                .curpos = 0
                                .currow = .currow + 1
                            End If
                Case 6
                        ''    LCTbasketCur scr, prive
                            wwPlain basestack, prive, s$, .Column * .Xt + .Xt, 0, , False, 2, , , True
                                        .curpos = .curpos + .Column + 1
                            If .curpos >= .mx And Not w4 Then
                                .curpos = 0
                                .currow = .currow + 1
                             End If
                Case 7
                            
                            LCTbasketCur Scr, prive
                    work2 = Scr.CurrentY
                            
                            wwPlain basestack, prive, s$ & vbCrLf, .Column * .Xt + .Xt - (.Xt - TextWidth(Scr, Right$(s$, 1))) \ 2, 0, , True, 1, , , True, , True
                       Scr.CurrentY = work2
                            .curpos = .curpos + .Column + 1
                            If .curpos >= .mx And Not w4 Then
                                .curpos = 0
                                .currow = .currow + 1
                            End If
                Case 8
                            LCTbasketCur Scr, prive
                            Scr.CurrentX = Scr.CurrentX + (.Xt - TextWidth(Scr, Left$(s$, 1))) \ 2
                            If Not (.mx - 1 <= x1 And w4 <> 0) Then
                                    wwPlain basestack, prive, s$, .Column * .Xt + .Xt - (.Xt - TextWidth(Scr, Left$(s$, 1))) \ 2, 0, , True, 0, , , True
                            End If
                            .curpos = .curpos + .Column + 1
                            If .curpos >= .mx And Not w4 Then
                                .curpos = 0
                                .currow = .currow + 1
                            End If
                Case 9
                            LCTbasketCur Scr, prive

              wPlain Scr, prive, s$, .Column + 1, 0, True
                GetXYb Scr, prive, .curpos, .currow
                          .curpos = .curpos + 1
                            If (.curpos Mod (.Column + 1)) <> 0 Then
                     .curpos = .curpos + (.Column + 1) - (.curpos Mod (.Column + 1)) + ColOffset
                      Else
                       .curpos = .curpos + ColOffset
                      End If
                            If .curpos >= .mx And Not w4 Then
                                .curpos = 0
                                .currow = .currow + 1
                            End If
                End Select
                End If
        Else
            PlainBaSket Scr, prive, s$
        
        End If
        Else
              If F < 0 Then
            PlainBaSket Scr, prive, s$, , , , , True
        ElseIf uni(F) Then
            putUniString F, s$
            Else
            putANSIString F, s$
        'Print #f, S$;
        End If
        End If
    Case 6
        If par Then
                If .FTEXT > 3 Then
            w3 = 0
             x1 = .curpos
             y1 = .currow
                        If .FTXT <> "" Then
                                       s$ = Format$(p, .FTXT)
            If Not NoUseDec Then
               If OverideDec Then
                s$ = Replace$(s$, GetDeflocaleString(LOCALE_SDECIMAL), Chr(2))
                s$ = Replace$(s$, GetDeflocaleString(LOCALE_STHOUSAND), Chr(3))
                s$ = Replace$(s$, Chr(2), NowDec$)
                s$ = Replace$(s$, Chr(3), NowThou$)
                
            ElseIf InStr(s$, NowDec$) > 0 And InStr(.FTXT, ".") > 0 Then
                s$ = Replace$(s$, NowDec$, Chr(2))
                s$ = Replace$(s$, NowThou$, Chr(3))
                s$ = Replace$(s$, Chr(2), ".")
                s$ = Replace$(s$, Chr(3), ",")
            
            End If
            End If
                               
                                If .FTEXT > 4 And Not work Then Scr.CurrentX = Scr.CurrentX + (.Xt - TextWidth(Scr, Left$(s$, 1))) \ 2
                                If Scr.CurrentX < .mx * .Xt Then
                            
                                wwPlain basestack, prive, s$, Scr.Width, 0, , True, 0, , w3
                                
                                End If
                                
                        Else
                                 If xa Then
                                        s$ = PACKLNG2$(p)
                                Else
                                If NoUseDec Then
                                       s$ = CStr(p)
                                    Else
                                     s$ = LTrim$(Str(p))
                                      If Left$(s$, 1) = "." Then
                                        s$ = "0" + s$
                                        ElseIf Left$(s$, 2) = "-." Then s$ = "-0" + Mid$(s$, 2)
                                        End If
                                     If OverideDec Then s$ = Replace$(s$, ".", NowDec$)
                                    End If
                                End If

                                If .FTEXT > 4 And Not work Then Scr.CurrentX = Scr.CurrentX + (.Xt - TextWidth(Scr, Left$(s$, 1))) \ 2
                                      If Scr.CurrentX < 0 Then
                             
                                
                                
                                End If
                                wwPlain basestack, prive, s$, Scr.Width, 0, , True, 0, , w3
                                work = True
                                Scr.CurrentX = w3
                         
                                            
                        End If
                      '' Then LCTbasket scr, prive, y1, W3 \ .Xt + 1
                Else
                        If .FTXT = vbNullString Then
                      
                                If xa Then
                                    PlainBaSket Scr, prive, PACKLNG2$(p)
                                Else
                                  If NoUseDec Then
                                    s$ = CStr(p)
                                        Else
                                            s$ = LTrim$(Str(p))
                                                If Left$(s$, 1) = "." Then
                                                s$ = "0" + s$
                                                ElseIf Left$(s$, 2) = "-." Then s$ = "-0" + Mid$(s$, 2)
                                                End If
                                            If OverideDec Then s$ = Replace$(s$, ".", NowDec$)
                                        End If
                                    PlainBaSket Scr, prive, s$
                                End If
                        Else
                      s$ = Format$(p, .FTXT)
            If Not NoUseDec Then
                If OverideDec Then
                    s$ = Replace$(s$, GetDeflocaleString(LOCALE_SDECIMAL), Chr(2))
                    s$ = Replace$(s$, GetDeflocaleString(LOCALE_STHOUSAND), Chr(3))
                    s$ = Replace$(s$, Chr(2), NowDec$)
                    s$ = Replace$(s$, Chr(3), NowThou$)
                ElseIf InStr(s$, NowDec$) > 0 And InStr(.FTXT, ".") > 0 Then
                    s$ = Replace$(s$, NowDec$, Chr(2))
                    s$ = Replace$(s$, NowThou$, Chr(3))
                    s$ = Replace$(s$, Chr(2), ".")
                    s$ = Replace$(s$, Chr(3), ",")
                
                End If
            End If
      
                            PlainBaSket Scr, prive, s$
                        End If
                End If
        Else
              If F < 0 Then
            PlainBaSket Scr, prive, s$
        ElseIf uni(F) Then
            putUniString F, s$
            Else
            putANSIString F, s$
        ' Print #f, S$;
        End If
        End If
    Case 7
        If par Then
        If s$ <> "" Then
           If .FTEXT > 3 Then
            w3 = 0
             x1 = .curpos
             y1 = .currow
            If Not work Then LCTbasketCur Scr, prive
              If .FTXT <> "" Then s$ = Format$(s$, .FTXT)
                        If .FTEXT > 4 And Not work Then Scr.CurrentX = Scr.CurrentX + (.Xt - TextWidth(Scr, Left$(s$, 1))) \ 2
                        wwPlain basestack, prive, s$, Scr.Width, 0, , True, 0, , w3
                        work = True
                       Scr.CurrentX = w3
            Else
                If .FTXT <> "" Then
                PlainBaSket Scr, prive, Format$(s$, .FTXT), , , , clearline
                Else
                PlainBaSket Scr, prive, s$, , , , clearline
                End If
                
            End If
        Else

          
        End If
  
            
        Else
              If F < 0 Then
            PlainBaSket Scr, prive, s$
        ElseIf uni(F) Then
            putUniString F, s$
            Else
            putANSIString F, s$
        ' Print #f, S$;
        End If
        End If
    End Select
taketwo:
If ihavecoma Then
ihavecoma = False
GoTo cont12344
    ElseIf FastSymbol(rest$, ",") Then
cont12344:
        w3 = 1
        pn& = 0
      ''  skiplast = False
        If opn& > 4 Then
            Scr.CurrentX = Scr.CurrentX + .Xt - dv15
            GetXYb Scr, prive, .curpos, .currow
            If work Then
                .curpos = .curpos - ColOffset
                If (.curpos Mod (.Column + 1)) <> 0 Then
                    .curpos = .curpos + (.Column + 1) - (.curpos Mod (.Column + 1)) + ColOffset
                Else
                    .curpos = .curpos + ColOffset
                End If
                If w4 Then LCTbasketCur Scr, prive
            End If
            work = True
        Else
            work = False
        End If
        opn& = 0
        Do While FastSymbol(rest$, ",")
            If par Then
            ' ok I want that
            If .Column > .mx And .FTEXT < 4 Then
            Else
                If Not w4 Then
                    If Not skiplast Then crNew basestack, prive
                End If
            End If
            Else
                If F < 0 Then
                    crNew basestack, prive
                    
                ElseIf uni(F) Then
                    putUniString F, vbCrLf
                Else
                    putANSIString F, vbCrLf
            'Print #f,
                End If
            End If

        Loop
    End If
If par Or F < 0 Then players(basketcode) = prive

Loop
there:
If w4 <> 0 And par Then
        .FTEXT = oldFTEXT
        .FTXT = oldFTXT
        .Column = oldCol
        If .mypen <> oldpen Then .mypen = oldpen: TextColor Scr, oldpen
        ElseIf par Then
        If pn& > 4 And opn& = 0 Then
        
                 If pn& < 99 Then
                 If work Then
                 .lastprint = False
                 End If
                 If Not skiplast Then crNew basestack, prive
                 End If
        ElseIf .currow >= .My Or (w3 < 0 And pn& = 0) Then
              crNew basestack, prive
              LCTbasketCur Scr, prive
        ElseIf pn& > 4 Then
       
        End If

End If
exitnow:
If basestack.IamThread Then
' let thread do the refresh
ElseIf par Or F < 0 Then
    If Not extreme Then
    PrintRefresh basestack, Scr
    End If
End If
RevisionPrint = True
If par Or F < 0 Then players(basketcode) = prive

End With

exit2:
'If Len(rest$) > 0 Then RevisionPrint = False
If Len(rest$) > 0 Then
    If Len(rest$) < where Then
        If where - Len(rest$) = 1 Then
            Mid$(rest1$, 1, Len(rest$)) = rest$
        Else
            Mid$(rest1$, where - Len(rest$), Len(rest$)) = rest$
            rest1$ = Mid$(rest1$, where - Len(rest$))
        End If
    Else
        rest1$ = rest$ + rest1$
    End If
Else
    rest1$ = Mid$(rest1$, where)
End If
If SLOW Then Call myexit(basestack)
End Function
Function RetM2000array(var As Variant) As Variant
Dim ar As New mArray, v(), manydim As Long, probe As Long, probelow As Long
Dim j As Long
v() = var
On Error GoTo ma100
For j = 1 To 60
    probe = UBound(v, j)
    If Err Then Exit For
Next j
manydim = j - 1
On Error Resume Next
For j = manydim To 1 Step -1
    
    probe = UBound(v, j)
    If Err Then Exit For
    probelow = LBound(v, j)
    ar.PushDim probe - probelow + 1
Next j
ar.PushEnd
ar.RevOrder = True
ar.CopySerialize v()
ma100:
Set RetM2000array = ar

End Function

Private Function MyMod(r1, po) As Variant
MyMod = r1 - Fix(r1 / po) * po
End Function
Sub dset()

'USING the temporary path
    strTemp = String(MAX_FILENAME_LEN, Chr$(0))
    'Get
    GetTempPath MAX_FILENAME_LEN, StrPtr(strTemp)
    strTemp = LONGNAME(mylcasefILE(Left$(strTemp, InStr(strTemp, Chr(0)) - 1)))
    If strTemp = vbNullString Then
     strTemp = mylcasefILE(Left$(strTemp, InStr(strTemp, Chr(0)) - 1))
    End If
' NOW COPY
' for mcd
Dim cd As String, dummy As Long, q$

''cd = App.Path
''AddDirSep cd
''mcd = mylcasefILE(cd)

' Return to standrad path...for all users
userfiles = GetSpecialfolder(CLng(26)) & "\M2000"
AddDirSep userfiles
If Not isdir(userfiles) Then
MkDir userfiles
End If

mcd = userfiles
DefaultDec$ = GetDeflocaleString(LOCALE_SDECIMAL)
If NowDec$ <> "" Then
ElseIf OverideDec Then
NowDec$ = GetlocaleString(LOCALE_SDECIMAL)
NowThou$ = GetlocaleString(LOCALE_STHOUSAND)
Else
NowDec$ = DefaultDec$
NowThou$ = GetDeflocaleString(LOCALE_STHOUSAND)
End If
CheckDec
cdecimaldot$ = GetDeflocaleString(LOCALE_SDECIMAL)
End Sub
Public Sub CheckDec()
OverideDec = False
NowDec$ = GetDeflocaleString(LOCALE_SDECIMAL)
NowThou$ = GetDeflocaleString(LOCALE_STHOUSAND)
If NowDec$ = "." Then
NoUseDec = False
Else
NoUseDec = mNoUseDec
End If
End Sub
Function ProcEnumGroup(bstack As basetask, rest$, Optional glob As Boolean = False) As Boolean

    Dim s$, w1$, v As Long, enumvalue As Long, myenum As Enumeration, mh As mHandler, v1 As Long
    enumvalue = 0
    If IsLabelOnly(rest$, w1$) = 1 Then
       ' w1$ = myUcase$(w1$)
        v = globalvar(bstack.GroupName + myUcase$(w1$), v, , glob)
        Set myenum = New Enumeration
        
        myenum.EnumName = w1$
        Else
        MyEr "No proper name for enumeration", "μη κανονικό όνομα για απαρίθμηση"
        Exit Function
    End If
    If FastSymbol(rest$, "{") Then
        s$ = block(rest$)
        
        Do
        If FastSymbol(s$, vbCrLf, , 2) Then
        While FastSymbol(s$, vbCrLf, , 2)
        Wend
        ElseIf IsLabelOnly(s$, w1$) = 1 Then
            'w1 = myUcase(w1$)
            If FastSymbol(s$, "=") Then
            If IsExp(bstack, s$, enumvalue) Then
                If Not bstack.lastobj Is Nothing Then
                    MyEr "No Object allowed as enumeration value", "Δεν επιτρέπεται αντικείμενο για τιμή απαριθμητή"
                    Exit Function
                    End If
                End If
            Else
                    enumvalue = enumvalue + 1
            End If
            myenum.addone w1$, enumvalue
            Set mh = New mHandler
            Set mh.objref = myenum
            mh.t1 = 4
            mh.ReadOnly = True
            mh.index_cursor = enumvalue
            mh.index_start = myenum.count - 1
             v1 = globalvar(bstack.GroupName + myUcase(w1$), v1, , glob)
             Set var(v1) = mh
            ProcEnumGroup = True
        Else
            Exit Do
        End If
        If FastSymbol(s$, ",") Then ProcEnumGroup = False
        Loop
        If v1 > v Then Set var(v) = var(v1) Else MyEr "Empty Enumeration", "’δεια Απαρίθμηση": Exit Function
        ProcEnumGroup = FastSymbol(rest$, "}", True)
    Else
        MissingEnumBlock
        Exit Function
    End If
    
    
End Function
Function ProcEnum(bstack As basetask, rest$, Optional glob As Boolean = False) As Boolean

    Dim s$, w1$, v As Long, enumvalue As Variant, myenum As Enumeration, mh As mHandler, v1 As Long, i As Long
    enumvalue = 0#
    If IsLabelOnly(rest$, w1$) = 1 Then
       ' w1$ = myUcase$(w1$)
        v = globalvar(myUcase$(w1$), v, , glob)
        Set myenum = New Enumeration
        
        myenum.EnumName = w1$
        Else
        MyEr "No proper name for enumeration", "μη κανονικό όνομα για απαρίθμηση"
        Exit Function
    End If
    If FastSymbol(rest$, "{") Then
        s$ = block(rest$)
        
        Do
        If FastSymbol(s$, vbCrLf, , 2) Then
        While FastSymbol(s$, vbCrLf, , 2)
        Wend
        ElseIf MaybeIsSymbol(s$, "\'") Then
        
        SetNextLine s$
        ElseIf IsLabelOnly(s$, w1$) = 1 Then
            'w1 = myUcase(w1$)
            If FastSymbol(s$, "=") Then
            If IsExp(bstack, s$, enumvalue) Then
                If Not bstack.lastobj Is Nothing Then
                    MyEr "No Object allowed as enumeration value", "Δεν επιτρέπεται αντικείμενο για τιμή απαριθμητή"
                    Exit Function
                   End If
            Else
                    MyEr "No String allowed as enumeration value", "Δεν επιτρέπεται αλφαριθμητικό για τιμή απαριθμητή"
                    Exit Function
            Exit Function
                End If
            Else
                    enumvalue = enumvalue + 1
            End If
            myenum.addone w1$, enumvalue
            w1$ = myUcase(w1$, True)
            If numid.Find(w1$, i) Then If i > 0 Then numid.ItemCreator2 w1$, -1
            
            Set mh = New mHandler
            Set mh.objref = myenum
            mh.t1 = 4
            mh.ReadOnly = True
            mh.index_cursor = enumvalue
            mh.index_start = myenum.count - 1
            
             v1 = globalvar(w1$, v1, , glob)
             Set var(v1) = mh
            ProcEnum = True
        Else
            Exit Do
        End If
        If FastSymbol(s$, ",") Then ProcEnum = False
        Loop
        If v1 > v Then Set var(v) = var(v1) Else MyEr "Empty Enumeration", "’δεια Απαρίθμηση": Exit Function
        ProcEnum = FastSymbol(rest$, "}", True)
    Else
        MissingEnumBlock
        Exit Function
    End If
    
    
End Function
Function CallLambdaASAP(bstack As basetask, a$, r, Optional forstring As Boolean = False) As Long
Dim w2 As Long, w1 As Long, nbstack As basetask
PushStage bstack, False
w2 = var2used
If forstring Then
w1 = globalvarGroup("A_" + CStr(w2) + "$", 0#)
 Set var(w1) = bstack.lastobj
 Set bstack.lastobj = Nothing
  If here$ = vbNullString Then
            GlobalSub "A_" + CStr(Abs(w2)) + "$()", "", , , w1
        Else
            GlobalSub here$ & "." & bstack.GroupName & "A_" + CStr(Abs(w2)) + "$()", "", , , w1
    End If
 Set nbstack = New basetask
    nbstack.reflimit = varhash.count
    Set nbstack.Parent = bstack
    If bstack.IamThread Then Set nbstack.Process = bstack.Process
    Set nbstack.Owner = bstack.Owner
    nbstack.OriginalCode = 0
    nbstack.UseGroupname = vbNullString
 CallLambdaASAP = GoFunc(nbstack, "A_" + CStr(Abs(w2)) + "$()", a$, r)

Else
w1 = globalvarGroup("A_" + CStr(w2), 0#)
 Set var(w1) = bstack.lastobj
 Set bstack.lastobj = Nothing
  If here$ = vbNullString Then
            GlobalSub "A_" + CStr(Abs(w2)) + "()", "", , , w1
        Else
            GlobalSub here$ & "." & bstack.GroupName & "A_" + CStr(Abs(w2)) + "()", "", , , w1
    End If
     Set nbstack = New basetask
    nbstack.reflimit = varhash.count
    Set nbstack.Parent = bstack
    If bstack.IamThread Then Set nbstack.Process = bstack.Process
    Set nbstack.Owner = bstack.Owner
    nbstack.OriginalCode = 0
    nbstack.UseGroupname = vbNullString
 CallLambdaASAP = GoFunc(nbstack, "A_" + CStr(Abs(w2)) + "()", a$, r)
End If


                 
PopStage bstack
End Function

Function ProcText(basestack As basetask, isHtml As Boolean, rest$) As Boolean
Dim x1 As Long, frm$, pa$, s$
ProcText = True
If IsSymbol(rest$, "UTF-8", 5) Then
x1 = 2
ElseIf IsSymbol(rest$, "UTF-16", 6) Then
x1 = 0 ' only little endian (but if something convert it to big we can read...)
Else
x1 = 3
End If

s$ = vbNullString
If Not IsStrExp(basestack, rest$, s$) Then
If Not Abs(IsLabelOnly(rest$, s$)) = 1 Then
    ProcText = False
    Exit Function
End If
End If
FastSymbol rest$, ","
If s$ <> "" Then

If FastSymbol(rest$, "+") Then pa$ = vbNullString Else pa$ = "new"
If FastSymbol(rest$, "{") Then frm$ = NLTrim2$(blockString(rest$, 125))
If frm$ <> "" Then
If isHtml Then
If ExtractType(s$) = vbNullString Then s$ = s$ & ".html"
End If
 textPUT basestack, mylcasefILE(s$), frm$, pa$, x1
Else
 textDel (mylcasefILE(s$))
 ProcText = True
 Exit Function
End If
ProcText = FastSymbol(rest$, "}")
End If
Exit Function

End Function
Private Function textPUT(bstack As basetask, ByVal ThisFile As String, THISBODY As String, c$, mode2save As Long) As Boolean
Dim chk As String, b$, j As Long, PREPARE$, VR$, s$, v As Double, buf$, i As Long
ThisFile = strTemp + ThisFile
chk = GetDosPath(ThisFile)
If chk <> "" And c$ = "new" Then KillFile GetDosPath(chk)
On Error GoTo HM
textPUT = True
Do
j = InStr(THISBODY, "##")
If j = 0 Then PREPARE$ = PREPARE$ & THISBODY: Exit Do
If j > 1 Then PREPARE$ = PREPARE$ & Mid$(THISBODY, 1, InStr(THISBODY, "##") - 1)
THISBODY = Mid$(THISBODY, j + 2)
j = InStr(THISBODY, "##")
If j = 0 Then PREPARE$ = PREPARE$ & THISBODY: Exit Do
If j > 1 Then VR$ = Mid$(THISBODY, 1, InStr(THISBODY, "##") - 1)
THISBODY = Mid$(THISBODY, j + 2)
'
If IsExp(bstack, VR$, v, , True) Then
buf$ = Trim$(Str$(v))
ElseIf IsStrExp(bstack, VR$, s$) Then
buf$ = s$
Else
buf$ = VR$
End If
PREPARE$ = PREPARE$ & buf$
Loop
           If Not WeCanWrite(ThisFile) Then GoTo HM

textPUT = SaveUnicode(ThisFile, PREPARE$, mode2save, Not (c$ = "new"))
Exit Function
HM:
textPUT = False
End Function
Private Function textDel(ByVal ThisFile As String) As Boolean
Dim chk As String
ThisFile = strTemp + ThisFile
chk = CFname(ThisFile)
textDel = (chk <> "")
If chk <> "" Then KillFile chk
End Function
Function MyPset(bstack As basetask, rest$) As Boolean
Dim prive As Long, X As Double, p As Variant, Y As Double, col As Long
Dim Scr As Object, ss$
Set Scr = bstack.Owner
prive = GetCode(Scr)
With players(prive)
    col = players(prive).mypen
    If IsExp(bstack, rest$, p, , True) Then col = mycolor(p)
    If FastSymbol(rest$, ",") Then
        If IsExp(bstack, rest$, X, , True) Then
            If FastSymbol(rest$, ",") Then
                If IsExp(bstack, rest$, Y, , True) Then
                    Scr.PSet (X, Y), col: MyPset = True
                Else
                    MissPar
                End If
            End If
        Else
            MissPar
        End If
    Else
        Scr.PSet (.XGRAPH, .YGRAPH), col
        MyPset = True
    End If
End With
MyDoEvents1 Scr
Set Scr = Nothing
End Function
Function Matrix(bstack As basetask, a$, Arr As Variant, res As Variant) As Boolean
Dim Pad$, cut As Long, pppp As mArray, pppp1 As mArray, st1 As mStiva, anything As Object, w3 As Long, useHandler As mHandler, r As Variant, p As Variant
Dim cur As Long, w2 As Long, w4 As Long, retresonly As Boolean
Dim multi As Boolean, original As Long
Set anything = Arr
If Not CheckLastHandlerOrIterator(anything, w3) Then Exit Function
Pad$ = myUcase(Left$(a$, 20))  ' 20??
cut = InStr(Pad$, "(")

If cut <= 1 Then Exit Function
Mid$(a$, 1, cut) = space$(cut)
Set useHandler = anything
If TypeOf useHandler.objref Is mArray Then
Set pppp = useHandler.objref
If Left$(Pad$, 1) = Chr$(1) Then LSet Pad$ = Mid$(Pad$, 2): cut = cut - 1
Do
multi = False
Select Case Left$(Pad$, cut - 1)
Case "SUM", "ΑΘΡ"
res = 0
For w3 = 0 To pppp.count - 1
If pppp.MyIsNumeric(pppp.item(w3)) Then res = res + pppp.item(w3)
Next w3
Case "MIN", "ΜΙΚ"
res = 0
w4 = -1
If pppp.count > 0 Then
For w3 = 0 To pppp.count - 1
If pppp.itemnumeric(w3) Then res = pppp.itemnumeric(w3): w4 = w3: Exit For
Next w3

For w3 = w3 To pppp.count - 1
If pppp.MyIsNumeric(pppp.item(w3)) Then If pppp.item(w3) < res Then res = pppp.item(w3): w4 = w3
Next w3
End If
If Not FastSymbol(a$, ")") Then
    bstack.soros.PushVal w4
    If Not getone(bstack, a$) Then Exit Function
    Else
    Matrix = True
    Exit Function
End If
Case "MIN$", "ΜΙΚ$"
res = vbNullString
w4 = -1
If pppp.count > 0 Then
For w3 = 0 To pppp.count - 1
If pppp.IsStringItem(w3) Then res = pppp.item(w3): w4 = w3: Exit For
Next w3

For w3 = w3 To pppp.count - 1
If pppp.IsStringItem(w3) Then If pppp.item(w3) < res Then res = pppp.item(w3): w4 = w3
Next w3
End If
If Not FastSymbol(a$, ")") Then
    bstack.soros.PushVal w4
    If Not getone(bstack, a$) Then Exit Function
Else
    Matrix = True
    Exit Function
End If

Case "MAX$", "ΜΕΓ$"
res = vbNullString
w4 = -1
If pppp.count > 0 Then
For w3 = 0 To pppp.count - 1
If pppp.IsStringItem(w3) Then res = pppp.item(w3): w4 = w3: Exit For
Next w3

For w3 = w3 To pppp.count - 1
If pppp.IsStringItem(w3) Then If pppp.item(w3) > res Then res = pppp.item(w3): w4 = w3
Next w3
End If
If Not FastSymbol(a$, ")") Then
    bstack.soros.PushVal w4
    If Not getone(bstack, a$) Then Exit Function
Else
    Matrix = True
    Exit Function
End If
Case "MAX", "ΜΕΓ"
res = 0
w4 = -1
If pppp.count > 0 Then
For w3 = 0 To pppp.count - 1
If pppp.MyIsNumeric(pppp.item(w3)) Then res = pppp.itemnumeric(w3): w4 = w3: Exit For
Next w3

For w3 = w3 To pppp.count - 1
If pppp.MyIsNumeric(pppp.item(w3)) Then If pppp.item(w3) > res Then res = pppp.item(w3): w4 = w3
Next w3
End If
If Not FastSymbol(a$, ")") Then
    
    bstack.soros.PushVal w4
    If Not getone(bstack, a$) Then Exit Function
    
Else
    Matrix = True
    Exit Function
End If
Case "VAL", "ΤΙΜΗ", "VAL$", "ΤΙΜΗ$"
If IsExp(bstack, a$, p, , True) Then
    w2 = CLng(p)
Else
    w2 = 0
End If
If w2 < 0 Or w2 >= pppp.count Then
MyEr "offset out of limits", "Δείκτης εκτός ορίων"
Matrix = False
Exit Function
Else
If pppp.MyIsObject(pppp.item(w2)) Then
Set bstack.lastobj = pppp.item(w2)
res = 0
If Not bstack.lastobj Is Nothing Then
If TypeOf bstack.lastobj Is mHandler Then
Set useHandler = bstack.lastobj
If useHandler.t1 = 3 Then
If TypeOf useHandler.objref Is mArray Then
Set pppp = useHandler.objref
multi = True
End If
End If
End If
End If
Else
res = pppp.item(w2)
End If
End If
Case "SLICE", "ΜΕΡΟΣ"
If IsExp(bstack, a$, p, , True) Then
If p < 0 Or p >= pppp.count Then
    MyEr "start offset out of limits", "Δείκτης αρχής εκτός ορίων"
    Matrix = False
    Exit Function
End If
Else
p = 0
End If
If FastSymbol(a$, ",") Then
If IsExp(bstack, a$, r, , True) Then
    If r >= pppp.count Or r < p Then
    MyEr "end offset out of limits", "Δείκτης τέλους εκτός ορίων"
    Matrix = False
    Exit Function
    End If
Else
r = pppp.count - 1
End If
Else
r = pppp.count - 1
End If
If original > 0 Then
pppp.CopyArraySliceFast pppp1, CLng(p), CLng(r)
Else
pppp.CopyArraySlice pppp1, CLng(p), CLng(r)
End If
original = original + 1
Set pppp = pppp1
Set pppp1 = Nothing
multi = True
Matrix = True
Set useHandler = New mHandler
useHandler.t1 = 3
Set useHandler.objref = pppp
Set bstack.lastobj = useHandler


Case "FOLD", "ΠΑΚ", "FOLD$", "ΠΑΚ$"
If IsExp(bstack, a$, p) Then
    If Not bstack.lastobj Is Nothing Then
        Set anything = bstack.lastobj
        If FastSymbol(a$, ",") Then
            If IsExp(bstack, a$, p) Then
                 res = p
            ElseIf IsStrExp(bstack, a$, Pad$) Then
                res = Pad$
            Else
               MissParam a$
                Matrix = False
                Exit Function
            End If
            
        End If
        Set bstack.lastobj = Nothing
        CallLambdaArrayFold bstack, pppp, anything, res
    Else
        MyEr "missing a lambda function", "λείπει μια λάμδα συνάρτηση"
        Matrix = False
        Exit Function
    End If
Else
    MyEr "missing a lambda function", "λείπει μια λάμδα συνάρτηση"
    Matrix = False
    Exit Function
End If
Case "REV", "ΑΝΑΠ"
Set pppp1 = New mArray
If original > 0 Then
    pppp.CopyArrayRevFast pppp1
Else
    pppp.CopyArrayRev pppp1
End If
original = original + 1
Set pppp = pppp1
Set pppp1 = Nothing
res = 0
multi = True
Matrix = True
Set useHandler = New mHandler
useHandler.t1 = 3
Set useHandler.objref = pppp
Set bstack.lastobj = useHandler

Case "MAP", "ΑΝΤ"
againmap:
If IsExp(bstack, a$, p) Then
    If Not bstack.lastobj Is Nothing Then
    CallLambdaArrayMap bstack, pppp, bstack.lastobj
    If FastSymbol(a$, ",") Then GoTo againmap
    Else
    MyEr "missing a lambda function", "λείπει μια λάμδα συνάρτηση"
    Matrix = False
    Exit Function
End If
Else
Set pppp1 = New mArray
pppp.CopyArray pppp1
Set pppp = pppp1
original = original + 1
End If
res = 0
If FastSymbol(a$, ",") Then
    If pppp.count = 0 Then
        If IsExp(bstack, a$, p) Then
             res = p: retresonly = True
        ElseIf IsStrExp(bstack, a$, Pad$) Then
            res = Pad$: retresonly = True
        Else
           MyEr "No value", "Χωρίς τιμή"
            Matrix = False
            Exit Function
        End If
    Else
        w2 = 1
        aheadstatus a$, , w2
        If w2 > 1 Then Mid$(a$, 1, w2 - 1) = space$(w2)
    End If
End If
Matrix = True
If Not retresonly Then
Set useHandler = New mHandler
useHandler.t1 = 3
Set useHandler.objref = pppp
Set bstack.lastobj = useHandler
End If
multi = True

Case "FILTER", "ΦΙΛΤΡΟ"
again:
If IsExp(bstack, a$, p) Then
    If Not bstack.lastobj Is Nothing Then
    CallLambdaArray bstack, pppp, bstack.lastobj
    Else
    MyEr "missing a lambda function", "λείπει μια λάμδα συνάρτηση"
    Matrix = False
    Exit Function
End If
Else
Set pppp1 = New mArray
pppp.CopyArray pppp1
Set pppp = pppp1
original = original + 1
End If
res = 0
If FastSymbol(a$, ",") Then
    If pppp.count = 0 Then
        If IsExp(bstack, a$, p) Then
            If Not bstack.lastobj Is Nothing Then
            If TypeOf bstack.lastobj Is mHandler Then
            Set useHandler = bstack.lastobj
            If useHandler.t1 = 3 Then
            If TypeOf useHandler.objref Is mArray Then
                Set pppp = useHandler.objref
                GoTo again
           ' End If
            End If
            End If
            End If
            End If
             res = p: retresonly = True
        ElseIf IsStrExp(bstack, a$, Pad$) Then
            res = Pad$: retresonly = True
        Else
           MyEr "No value", "Χωρίς τιμή"
            Matrix = False
            Exit Function
        End If
    Else
        w2 = 1
        aheadstatus a$, , w2
        If w2 > 1 Then Mid$(a$, 1, w2 - 1) = space$(w2)
    End If
End If
Matrix = True
If Not retresonly Then
Set useHandler = New mHandler
useHandler.t1 = 3
Set useHandler.objref = pppp
Set bstack.lastobj = useHandler
End If
multi = True
Case "POS", "ΘΕΣΗ"
    res = -1
    cur = 0
    Dim st() As String, sn() As Variant
    If IsExp(bstack, a$, p) Then
        p = Int(p)
        If p < 0 Then p = 0
again1:
        If FastSymbol(a$, "->", , 2) Then
            If IsExp(bstack, a$, r) Then
                If bstack.lastobj Is Nothing Then
                    ReDim sn(0 To 4) As Variant
                    sn(0) = r
                Else
dothis:
                    Set anything = bstack.lastobj
                    Set bstack.lastobj = Nothing
                    If Not CheckLastHandlerOrIterator(anything, w3) Then Exit Function
                    Set useHandler = anything
                    If Not TypeOf useHandler.objref Is mArray Then
                        If useHandler.t1 = 4 Then Set useHandler = Nothing: Set anything = Nothing: GoTo again1
                        Exit Function
                    End If
                    Set pppp1 = useHandler.objref
            
                    sn() = pppp1.GetCopy()
                    If pppp1.count > 0 Then
                        ReDim Preserve sn(0 To pppp1.count - 1)
                    End If
                    cur = pppp1.count - 1
                    w3 = p
                End If
            ElseIf IsStrExp(bstack, a$, Pad$) Then
                GoTo there
            Else
                MissParam a$: Exit Function
            End If
        Else
            If bstack.lastobj Is Nothing Then
                r = p
                p = 0
                ReDim sn(0 To 4) As Variant
                sn(0) = r
            Else
                GoTo dothis
            End If
        End If
        
        If pppp.count > 0 Then
            res = -1
            If Not pppp1 Is Nothing Then
                While res = -1 And cur >= 0 And w3 < pppp.count - cur - 1
                    For w3 = w3 To pppp.count - cur - 1
                        If pppp.MyIsObject(pppp.item(w3)) Then
                             If pppp.MyIsObject(sn(0)) Then GoTo inside
                        Else
                            If pppp.MyIsObject(sn(0)) Then GoTo inside
                            If pppp.item(w3) = sn(0) Then
inside:
                                res = w3
                                w4 = w3 + 1
                                For w2 = 1 To cur
                                    If w4 < pppp.count Then
                                        If pppp.MyIsObject(pppp.item(w4)) Then
                                            If Not pppp.MyIsObject(sn(w2)) Then
                                                res = -1
                                                Exit For
                                            End If
                                        Else
                                            If pppp.MyIsObject(sn(w2)) Then
                                                res = -1
                                                Exit For
                                            Else
                                                If pppp.item(w4) <> sn(w2) Then res = -1: Exit For
                                            End If
                                        End If
                                    End If
                                    w4 = w4 + 1
                                Next w2
                                If w2 > cur Then Exit For
                            End If
                        End If
                    Next w3
                Wend
            Else
                For w3 = p To pppp.count - 1
                    If pppp.MyIsNumeric(pppp.item(w3)) Then
                        If pppp.item(w3) = r Then res = w3: Exit For
                    Else
                        If Typename(pppp.item(w3)) = "mHandler" Then
                            Set useHandler = pppp.item(w3)
                            If useHandler.t1 = 4 Then
                                If useHandler.index_cursor * useHandler.sign = r Then res = w3: Exit For
                            End If
                        End If
                    End If
                Next w3
                w2 = 1
                Do While FastSymbol(a$, ",")
                    If cur = UBound(sn()) Then ReDim Preserve sn(0 To cur * 2 - 1) As Variant
                    cur = cur + 1
                    If IsExp(bstack, a$, sn(cur), , True) Then
                        If res > -1 Then
                            w3 = w3 + 1
                            If w3 < pppp.count Then
                                If pppp.MyIsNumeric(pppp.item(w3)) Then
                                    If pppp.item(w3) <> sn(cur) Then w2 = -1
                                ElseIf Typename(pppp.item(w3)) = "mHandler" Then
                                    Set useHandler = pppp.item(w3)
                                    If useHandler.t1 = 4 Then
                                        If useHandler.index_cursor * useHandler.sign <> sn(cur) Then w2 = -1
                                    Else
                                        w2 = -1
                                    End If
                                Else
                                    w2 = -1
                                End If
                            Else
                                w2 = -1
                            End If
                        End If
                    End If
                Loop
                If w2 = -1 Then
                    w3 = res + 1
                    res = -1
                    While res = -1 And cur > 0 And w3 < pppp.count - cur - 1
                        For w3 = w3 To pppp.count - cur - 1
                            If pppp.MyIsNumeric(pppp.item(w3)) Then
                                If pppp.item(w3) = sn(0) Then
                                    res = w3
                                    w4 = w3 + 1
                                    For w2 = 1 To cur
                                        If w4 < pppp.count Then
                                            If pppp.MyIsNumeric(pppp.itemnumeric(w4)) Then
                                                If pppp.item(w4) <> sn(w2) Then res = -1: Exit For
                                            ElseIf Typename(pppp.item(w4)) = "mHandler" Then
                                                Set useHandler = pppp.item(w4)
                                                If useHandler.t1 = 4 Then
                                                    If useHandler.index_cursor * useHandler.sign <> sn(w2) Then res = -1
                                                Else
                                                    res = -1
                                                End If
                                            Else
                                                res = -1
                                            End If
                                        End If
                                        w4 = w4 + 1
                                    Next w2
                                    If w2 > cur Then Exit For
                                End If
                            End If
                        Next w3
                    Wend
                End If
            End If
        End If
    ElseIf IsStrExp(bstack, a$, Pad$) Then
        
there:
        ReDim st(0 To 4) As String
        st(0) = Pad$
        If pppp.count > 0 Then
            res = -1
            For w3 = p To pppp.count - 1
                If pppp.IsStringItem(w3) Then
                    If pppp.item(w3) = st(0) Then res = w3: Exit For
                End If
            Next w3
            w2 = 1
            Do While FastSymbol(a$, ",")
                If cur = UBound(st()) Then ReDim Preserve st(0 To cur * 2 - 1) As String
                cur = cur + 1
                If IsStrExp(bstack, a$, st(cur)) Then
                    If res > -1 Then
                        w3 = w3 + 1
                        If w3 < pppp.count Then
                            If pppp.IsStringItem(w3) Then
                                If pppp.item(w3) <> st(cur) Then w2 = -1
                            Else
                                w2 = -1
                            End If
                        Else
                            w2 = -1
                        End If
                    End If
                Else
                    w2 = -1
                End If
            Loop
            If w2 = -1 Then
                w3 = res + 1
                res = -1
                While res = -1 And cur > 0 And w3 < pppp.count - cur - 1
                    For w3 = w3 To pppp.count - cur - 1
                        If pppp.IsStringItem(w3) Then
                            If pppp.item(w3) = st(0) Then
                                res = w3
                                w4 = w3 + 1
                                For w2 = 1 To cur
                                    If w4 < pppp.count Then
                                        If pppp.IsStringItem(w4) Then
                                            If pppp.item(w4) <> st(w2) Then res = -1: Exit For
                                        Else
                                            res = -1
                                        End If
                                    End If
                                    w4 = w4 + 1
                                Next w2
                                If w2 > cur Then Exit For
                            End If
                        End If
                    Next w3
                Wend
            End If
        End If
    Else
        MissParam a$
        Exit Function
    End If
End Select
Matrix = FastSymbol(a$, ")")
If Not multi Then Exit Do
If Matrix = False Then Exit Do
If Not IsOperator(a$, "#") Then Exit Do
If pppp.count = 0 Then
        Do
        w2 = 1
        aheadstatus a$, , w2
        If w2 > 1 Then Mid$(a$, 1, w2 - 1) = space$(w2)
        If Not FastSymbol(a$, ")") Then Matrix = False: Exit Function
        Loop Until Not IsOperator(a$, "#")
Exit Function
End If
Pad$ = myUcase(Left$(a$, 20))
cut = InStr(Pad$, "(")
If cut <= 1 Then Exit Do
Mid$(a$, 1, cut) = space$(cut)
Set bstack.lastobj = Nothing

Loop
Else
WrongObject
End If
End Function
Function getone(bstack As basetask, rest$) As Boolean
Dim what$, ss$, x1 As Long
getone = True
FastSymbol rest$, "&"
   x1 = Abs(IsLabelBig(bstack, rest$, what$))
    
    If x1 <> 0 Then
            If x1 > 4 Then
                    ss$ = BlockParam(rest$)
                    what$ = what$ + ss$ + ")"
                    'rest$ = Mid$(rest$, Len(ss$) + 2)
                    Mid$(rest$, 1, Len(ss$) + 1) = space(Len(ss$) + 1)
                    Do While IsSymbol(rest$, ".")
                    x1 = IsLabel(bstack, rest$, ss$)
                    If x1 > 0 Then what$ = what$ + "." + ss$ Else Exit Do
                            If x1 > 4 Then
                            ss$ = BlockParam(rest$)
                            what$ = what$ + ss$ + ")"
                            'rest$ = Mid$(rest$, Len(ss$) + 2)
                            Mid$(rest$, 1, Len(ss$) + 1) = space(Len(ss$) + 1)
                            End If
                    Loop
            End If
    
            



              
              getone = MyRead(6, bstack, (what$), 1, what$, x1, True)

             Else
             MissParamref rest$
             Exit Function
             End If
  



End Function
Sub CallLambdaArray(bstack As basetask, ByRef pppp As mArray, mylambda As lambda)
Dim w2 As Long, w1 As Long, nbstack As basetask
PushStage bstack, False
w2 = var2used
w1 = globalvarGroup("A_" + CStr(w2), 0#)
 Set var(w1) = mylambda
 Set bstack.lastobj = Nothing
  If here$ = vbNullString Then
            GlobalSub "A_" + CStr(Abs(w2)) + "()", "", , w1
        Else
            GlobalSub here$ & "." & bstack.GroupName & "A_" + CStr(Abs(w2)) + "()", "", , , w1
    End If
     Set nbstack = New basetask
    nbstack.reflimit = varhash.count
    Set nbstack.Parent = bstack
    If bstack.IamThread Then Set nbstack.Process = bstack.Process
    Set nbstack.Owner = bstack.Owner
    nbstack.OriginalCode = 0
    nbstack.UseGroupname = vbNullString
Dim aa As Object, oldsoros As mStiva, tempsoros As New mStiva, finalpppp As New mArray
finalpppp.StartResize: finalpppp.PushDim pppp.count: finalpppp.PushEnd
Set oldsoros = bstack.soros
Set bstack.Sorosref = tempsoros
Dim r, what As Long, where As Long
For w1 = 0 To pppp.count - 1
 
  If pppp.IsStringItem(w1) Then
    tempsoros.PushStrVariant pppp.item(w1)
    what = 1
  ElseIf pppp.MyIsObject(pppp.item(w1)) Then
    tempsoros.PushObj pppp.item(w1)
    what = 2
  Else
      tempsoros.PushVal pppp.item(w1)
      what = 3
  End If
  If Not GoFunc(nbstack, "A_" + CStr(Abs(w2)) + "()", vbNullString, r, w2, , , True) Then Exit For
  If CBool(r) Then
  If what <> 2 Then
  finalpppp.item(where) = pppp.item(w1)
 
  Else
   Set finalpppp.item(where) = pppp.item(w1)
  End If
  where = where + 1

  End If
  tempsoros.Flush
Next w1
Set bstack.Sorosref = oldsoros
PopStage bstack
finalpppp.StartResize: finalpppp.PushDim where: finalpppp.PushEnd
Set pppp = finalpppp

End Sub

Sub CallLambdaArrayMap(bstack As basetask, ByRef pppp As mArray, mylambda As lambda)
Dim w2 As Long, w1 As Long, nbstack As basetask
PushStage bstack, False
w2 = var2used
w1 = globalvarGroup("A_" + CStr(w2), 0#)
 Set var(w1) = mylambda
 Set bstack.lastobj = Nothing
  If here$ = vbNullString Then
            GlobalSub "A_" + CStr(Abs(w2)) + "()", "", , , w1
        Else
            GlobalSub here$ & "." & bstack.GroupName & "A_" + CStr(Abs(w2)) + "()", "", , , w1
    End If
     Set nbstack = New basetask
    nbstack.reflimit = varhash.count
    Set nbstack.Parent = bstack
    If bstack.IamThread Then Set nbstack.Process = bstack.Process
    Set nbstack.Owner = bstack.Owner
    nbstack.OriginalCode = 0
    nbstack.UseGroupname = vbNullString
Dim aa As Object, oldsoros As mStiva, tempsoros As New mStiva, finalpppp As New mArray
finalpppp.StartResize: finalpppp.PushDim pppp.count: finalpppp.PushEnd
Set oldsoros = bstack.soros
Set bstack.Sorosref = tempsoros
Dim r, what As Long, where As Long
For w1 = 0 To pppp.count - 1
 
  If pppp.IsStringItem(w1) Then
    tempsoros.PushStrVariant pppp.item(w1)
  ElseIf pppp.MyIsObject(pppp.item(w1)) Then
    tempsoros.PushObj pppp.item(w1)
  Else
      tempsoros.PushVal pppp.item(w1)
  End If
  If Not GoFunc(nbstack, "A_" + CStr(Abs(w2)) + "()", vbNullString, r, w2, , , True) Then Exit For
  If tempsoros.count > 0 Then
  If tempsoros.StackItemTypeIsObject(1) Then
  Set finalpppp.item(w1) = tempsoros.PopObj
  Else
   finalpppp.item(w1) = tempsoros.PopAnyNoObject
  End If
    End If
  
  tempsoros.Flush
Next w1
Set bstack.Sorosref = oldsoros
PopStage bstack
Set pppp = finalpppp

End Sub
Sub CallLambdaArrayFold(bstack As basetask, pppp As mArray, mylambda As lambda, res As Variant)
Dim w2 As Long, w1 As Long, nbstack As basetask
PushStage bstack, False
w2 = var2used
w1 = globalvarGroup("A_" + CStr(w2), 0#)
 Set var(w1) = mylambda
 Set bstack.lastobj = Nothing
  If here$ = vbNullString Then
            GlobalSub "A_" + CStr(Abs(w2)) + "()", "", , , w1
        Else
            GlobalSub here$ & "." & bstack.GroupName & "A_" + CStr(Abs(w2)) + "()", "", , , w1
    End If
     Set nbstack = New basetask
    nbstack.reflimit = varhash.count
    Set nbstack.Parent = bstack
    If bstack.IamThread Then Set nbstack.Process = bstack.Process
    Set nbstack.Owner = bstack.Owner
    nbstack.OriginalCode = 0
    nbstack.UseGroupname = vbNullString
Dim aa As Object, oldsoros As mStiva, tempsoros As New mStiva
Set oldsoros = bstack.soros
Set bstack.Sorosref = tempsoros
Dim r, what As Long, where As Long
If pppp.MyIsNumeric(res) Then
tempsoros.PushVal res
ElseIf pppp.MyIsObject(res) Then
Set aa = res
tempsoros.PushObj aa
Else
tempsoros.PushStrVariant res
End If

For w1 = 0 To pppp.count - 1

  If pppp.IsStringItem(w1) Then
    tempsoros.PushStrVariant pppp.item(w1)
  ElseIf pppp.MyIsObject(pppp.item(w1)) Then
    tempsoros.PushObj pppp.item(w1)
  Else
      tempsoros.PushVal pppp.item(w1)
  End If
  If Not GoFunc(nbstack, "A_" + CStr(Abs(w2)) + "()", vbNullString, r, w2, , , True) Then Exit For
  
Next w1
  If tempsoros.count > 0 Then
  If tempsoros.StackItemTypeIsObject(1) Then
        Set bstack.lastobj = tempsoros.PopObj
        res = 0
  Else
        res = tempsoros.PopAnyNoObject
  End If
    End If
Set bstack.Sorosref = oldsoros
PopStage bstack
End Sub

Function ChangeValues(bstack As basetask, rest$) As Boolean
Dim aa As mHandler, bb As FastCollection, ah As String, p As Variant, s$, lastindex As Long
Set aa = bstack.lastobj
Set bstack.lastobj = Nothing
Set bb = aa.objref
If bb.StructLen > 0 Then
MyEr "Structure members are ReadOnly", "Τα μέλη της δομής είναι μόνο για ανάγνωση"
Exit Function
End If
If bb.Done And FastSymbol(rest$, ":=", , 2) Then
        ' change one value
        ah = aheadstatus(rest$, False) + " "
        If Left$(ah, 1) = "N" Or InStr(ah, "l") > 0 Then
            If Not IsExp(bstack, rest$, p) Then
                ChangeValues = False
                GoTo there
            End If
            ChangeValues = True
            If Not bstack.lastobj Is Nothing Then
                Set bb.ValueObj = bstack.lastobj
                Set bstack.lastobj = Nothing
            Else
                bb.Value = p
            End If
            
        ElseIf Left$(ah, 1) = "S" Then
            If Not IsStrExp(bstack, rest$, s$) Then
                ChangeValues = False
                GoTo there
            End If
            ChangeValues = True
            If Not bstack.lastobj Is Nothing Then
                Set bb.ValueObj = bstack.lastobj
                Set bstack.lastobj = Nothing
            Else
                bb.Value = s$
            End If
        Else
                MyEr "No Data found", "Δεν βρέθηκαν στοιχεία"
                ChangeValues = False
        End If
        GoTo there

ElseIf MaybeIsSymbol(rest$, ",") Then
    Do While FastSymbol(rest$, ",")
        ChangeValues = True
        ah = aheadstatus(rest$, False) + " "
        If InStr(ah, "l") Then
                MyEr "Found logical expression", "Βρήκα λογική έκφραση"
                ChangeValues = False
        Else
                If Left$(ah, 1) = "N" Then
                    If Not IsExp(bstack, rest$, p) Then
                        ChangeValues = False
                        GoTo there
                    End If
                        If VarType(p) = vbBoolean Then p = CLng(p)
                    If Not bstack.lastobj Is Nothing Then
                        MyEr "No Object Allowed for Key", "Δεν επιτρέπεται αντικείμενο για κλειδί"
                        ChangeValues = False
                        GoTo there
                    End If
                    If Not bb.Find(p) Then
                         MyEr "No Key found", "Δεν βρέθηκε κλειδί"
                        ChangeValues = False
                        GoTo there
                    End If
                    
                ElseIf Left$(ah, 1) = "S" Then
                    If Not IsStrExp(bstack, rest$, s$) Then
                        ChangeValues = False
                        GoTo there
                    End If
                    If Not bstack.lastobj Is Nothing Then
                        MyEr "No Object Allowed for Key", "Δεν επιτρέπεται αντικείμενο για κλειδί"
                        ChangeValues = False
                        GoTo there
                    End If
                    If Not bb.Find(s$) Then
                          MyEr "No Key found", "Δεν βρέθηκε κλειδί"
                        ChangeValues = False
                        GoTo there
                    End If
                Else
                        MyEr "No Key found", "Δεν βρέθηκε κλειδί"
                        ChangeValues = False
                        GoTo there
                
                End If
lastindex = bb.index
                
                If FastSymbol(rest$, ":=", , 2) Then
                    ah = aheadstatus(rest$, False) + " "
                    If Left$(ah, 1) = "N" Or InStr(ah, "l") > 0 Then
                        If Not IsExp(bstack, rest$, p) Then
                            ChangeValues = False
                            GoTo there
                        End If
                        ChangeValues = True
                        bb.index = lastindex
                        If Not bstack.lastobj Is Nothing Then
                            Set bb.ValueObj = bstack.lastobj
                            Set bstack.lastobj = Nothing
                        Else
                            bb.Value = p
                        End If
                ElseIf Left$(ah, 1) = "S" Then
                        If Not IsStrExp(bstack, rest$, s$) Then
                            ChangeValues = False
                            GoTo there
                        End If
                        ChangeValues = True
                        bb.index = lastindex
                        If Not bstack.lastobj Is Nothing Then
                            Set bb.ValueObj = bstack.lastobj
                            Set bstack.lastobj = Nothing
                        Else
                            bb.Value = s$
                        End If
                Else
                        MyEr "No Data found", "Δεν βρέθηκαν στοιχεία"
                        ChangeValues = False
                End If
                End If
        End If
    Loop
ElseIf bb.Done Then
If Not bstack.soros.IsEmpty Then
With bstack.soros
If .StackItemTypeIsObject(1) Then
Set bb.ValueObj = .PopObj

Else
bb.Value = .StackItem(1)
.drop 1
End If
End With
ChangeValues = True
Else

End If
Set bstack.lastobj = Nothing
End If


there:
Set bb = Nothing
Set aa = Nothing
End Function
Function ChangeValuesArray(bstack As basetask, rest$) As Boolean
Dim aa As mHandler, p As Variant, pppp As mArray, w As Long, s$, ah As String, stiva As mStiva
Dim bs As Long
Set aa = bstack.lastobj
Dim anything As Object
Set anything = aa
If CheckIsmArray(anything) Then
    Set pppp = anything
    Set anything = Nothing
    If Not pppp.Arr Then
        NotArray
        Exit Function
    End If
    bs = pppp.myarrbase
    FastSymbol rest$, ","
    Do
    If IsExp(bstack, rest$, p, , True) Then
    On Error Resume Next
    w = CLng(Fix(p)) + bs
    If Err Then
        Err.clear
        OutOfLimit
        Exit Function
    End If
    On Error GoTo 0
    If w < 0 Then w = pppp.count - w + bs
    If w > (pppp.count + bs) Then GoTo outlimit
    If w < 0 Then
outlimit:
    MyEr "Index out of limits", "Ο δείκτης είναι εκτός ορίων"
    Exit Function
    End If
    
    If FastSymbol(rest$, ":=", , 2) Then
            
            ah = aheadstatus(rest$, False) + " "
            If Left$(ah, 1) = "N" Or InStr(ah, "l") > 0 Then
                If Not IsExp(bstack, rest$, p) Then
                    ChangeValuesArray = False
                    GoTo there
                End If
                ChangeValuesArray = True
                If Not bstack.lastobj Is Nothing Then
                    Set pppp.item(w) = bstack.lastobj
                    Set bstack.lastobj = Nothing
                Else
                    pppp.item(w) = p
                End If
                
            ElseIf Left$(ah, 1) = "S" Then
                If Not IsStrExp(bstack, rest$, s$) Then
                    ChangeValuesArray = False
                    GoTo there
                End If
                ChangeValuesArray = True
                If Not bstack.lastobj Is Nothing Then
                    Set pppp.item(w) = bstack.lastobj
                    Set bstack.lastobj = Nothing
                Else
                    pppp.item(w) = s$
                End If
            Else
                    MyEr "No Data found", "Δεν βρέθηκαν στοιχεία"
                    ChangeValuesArray = False
            End If
    End If
    End If
    Loop Until Not FastSymbol(rest$, ",")
Else
Set bstack.lastobj = aa
If CheckStackObj(bstack, anything) Then
    Set stiva = anything
    Set anything = Nothing
    FastSymbol rest$, ","
    Do
    If IsExp(bstack, rest$, p, , True) Then
    On Error Resume Next
    w = CLng(Fix(p))
    If Err Then
        Err.clear
        OutOfLimit
        Exit Function
    End If
    On Error GoTo 0
    If w < 0 Then w = stiva.count - w + 1
    If w > stiva.count Then GoTo outlimit
    If w < 0 Then GoTo outlimit
   
    If FastSymbol(rest$, ":=", , 2) Then
            
            ah = aheadstatus(rest$, False) + " "
            If Left$(ah, 1) = "N" Or InStr(ah, "l") > 0 Then
                If Not IsExp(bstack, rest$, p) Then
                    ChangeValuesArray = False
                    GoTo there
                End If
                ChangeValuesArray = True
                If Not bstack.lastobj Is Nothing Then
                    stiva.MakeTopItem w
                    stiva.drop 1
                    stiva.PushObj bstack.lastobj
                    stiva.MakeTopItemBack w
                    Set bstack.lastobj = Nothing
                Else
                    stiva.MakeTopItem w
                    stiva.drop 1
                    stiva.PushVal p
                    stiva.MakeTopItemBack w
                End If
                
            ElseIf Left$(ah, 1) = "S" Then
                If Not IsStrExp(bstack, rest$, s$) Then
                    ChangeValuesArray = False
                    GoTo there
                End If
                ChangeValuesArray = True
                If Not bstack.lastobj Is Nothing Then
                    stiva.MakeTopItem w
                    stiva.drop 1
                    stiva.PushObj bstack.lastobj
                    stiva.MakeTopItemBack w
                    Set bstack.lastobj = Nothing
                Else
                    stiva.MakeTopItem w
                    stiva.drop 1
                    stiva.PushStr s$
                    stiva.MakeTopItemBack w
                End If
            Else
                    MyEr "No Data found", "Δεν βρέθηκαν στοιχεία"
                    ChangeValuesArray = False
            End If
    End If
    End If
    Loop Until Not FastSymbol(rest$, ",")
End If
'CheckStackObj
End If
there:
Set anything = Nothing
Set aa = Nothing
End Function
Sub FeedCopyInOut(bstack As basetask, var$, where As Long, Arr$)
Dim a As New CopyInOut
a.actualvar = var$
a.ArrArg = Arr$
a.localvar = where
If bstack.CopyInOutCol Is Nothing Then Set bstack.CopyInOutCol = New Collection
bstack.CopyInOutCol.Add a
End Sub
Sub CopyBack(bstack As basetask)
Dim a As CopyInOut, aa As Object, x1 As Long, what$, rest$, oldhere$, w2 As Long
Dim pppp As mArray, s As String
oldhere$ = here$
here$ = vbNullString
For Each a In bstack.CopyInOutCol

If Len(a.ArrArg) > 0 Then
x1 = rinstr(a.actualvar, a.ArrArg) + Len(a.ArrArg) - 1
   
    If neoGetArray(bstack, Left$(a.actualvar, x1), pppp) Then
        If Not pppp.Arr Then GoTo cont123
        If Not NeoGetArrayItem(pppp, bstack, Left$(a.actualvar, x1), w2, Mid$(a.actualvar, x1 + 1)) Then GoTo cont123
        If MyIsObject(var(a.localvar)) Then
        If TypeOf pppp.item(w2) Is Group Then
        If pppp.item(w2).IamApointer Then
            Set pppp.item(w2) = var(a.localvar)
        Else
        Set pppp.item(w2) = CopyGroupObj(var(a.localvar), Not pppp.GroupRef Is Nothing)
        End If
        Else
            Set pppp.item(w2) = var(a.localvar)
            End If
        Else
            pppp.item(w2) = var(a.localvar)
        End If
    End If
Else
    If bstack.ExistVar2(a.actualvar) Then
        
        If MyIsObject(var(a.localvar)) Then
            bstack.SetVarobJ a.actualvar, var(a.localvar)
        Else
            bstack.SetVar a.actualvar, var(a.localvar)
        End If
    End If
    ' generic not used
        'If MyIsObject(var(a.localvar)) Then
          '  Set aa = var(a.localvar)
         '   bstack.soros.PushObj aa
        'ElseIf IsNumeric(var(a.localvar)) Then
         '   bstack.soros.PushVal var(a.localvar)
        'Else
         '   bstack.soros.PushStrVariant var(a.localvar)
        'End If
        'MyRead 1, bstack, a.actualvar, 1, , , True
End If
cont123:
Next
here$ = oldhere$
Set bstack.CopyInOutCol = Nothing
End Sub
Function GetOneAsString(bstack As basetask, rest$, what$, x1 As Long) As Boolean
Dim ss$
            If x1 > 4 Then
                    ss$ = BlockParam(rest$)
                    If Mid$(rest$, Len(ss$) + 1, 1) <> ")" Then Exit Function
                    what$ = what$ + ss$ + ")"
                   rest$ = Mid$(rest$, Len(ss$) + 2)
                    GetOneAsString = True
            End If
    
End Function
Function NewVarItem() As VarItem
    If TrushCount = 0 Then
    Set NewVarItem = New VarItem
      Exit Function
    End If
    Set NewVarItem = Trush(TrushCount)
    Set Trush(TrushCount) = Nothing
    TrushCount = TrushCount - 1
End Function
Function ExpMatrix(bstack As basetask, a$, r) As Boolean
Dim useHandler As mHandler
 If Not bstack.lastobj Is Nothing Then
                                If Typename(bstack.lastobj) = "mHandler" Then
                                    Set useHandler = bstack.lastobj
                                    Set bstack.lastobj = Nothing
                                    ExpMatrix = Matrix(bstack, a$, useHandler, r)
                                    If MyIsObject(r) Then r = CDbl(0)
                                   ' If SG < 0 Then r = -r
                                    Exit Function
                                ElseIf Typename(bstack.lastobj) = "mArray" Then
                                Set useHandler = New mHandler
                                useHandler.t1 = 3
                                Set useHandler.objref = bstack.lastobj
                                Set bstack.lastobj = Nothing
                                    ExpMatrix = Matrix(bstack, a$, useHandler, r)
                                    If MyIsObject(r) Then r = CDbl(0)
                                   ' If SG < 0 Then r = -r
                                    Exit Function
                                End If
                            End If
                                SyntaxError
                                ExpMatrix = False
                                Exit Function
End Function
Sub targetsMyExec(MyExec As Long, b$, bb$, v As Long, di As Object, w$, bstack As basetask, VarStat As Boolean, temphere$)
Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long, SBB$, nd&, p As Variant

If Abs(IsLabel(bstack, b$, w$)) = 1 Then
                    If Not GetVar(bstack, w$, v) Then
                     v = globalvar(w$, 0#, , VarStat, temphere$)

                               
                    End If
                Else
                    MyExec = 0
                    Exit Sub
                End If
                If Not FastSymbol(b$, ",") Then
                    MyExec = 0
                    Exit Sub
                ElseIf IsStrExp(bstack, b$, bb$) Then
                If NocharsInLine(bb$) Then MyExec = 0: Exit Sub
                With players(GetCode(di))
               '' SetTextSZ di, Sz
               '' LCT di, yPos, xPos
                x1 = 1
                y1 = 1
                x2 = -1
                y2 = -1
                nd& = 0
                SBB$ = vbNullString
                On Error GoTo err123
                If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p, , True) Then x1 = Abs(p) Mod (.mx + 1)
                If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p, , True) Then y1 = Abs(p) Mod (.My + 1)
                If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p, , True) Then x2 = CLng(Fix(p))
                If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p, , True) Then y2 = CLng(Fix(p))
                If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p, , True) Then nd& = Abs(p)
                If FastSymbol(b$, ",") Then If Not IsStrExp(bstack, b$, SBB$) Then MyExec = 0: Exit Sub
err123:
                        If Err.Number = 6 Then
                                Overflow
                                MyExec = 0
                                Exit Sub
                                End If
                Targets = False
                MyDoEvents1 Form1
       
                ReDim Preserve q(UBound(q()) + 1)
                q(UBound(q()) - 1) = BoxTarget(bstack, x1, y1, x2, y2, SBB$, nd&, bb$, .Xt, .Yt, .uMineLineSpace)
                End With
                var(v) = UBound(q()) - 1
                Targets = True
                ElseIf IsExp(bstack, b$, p, , True) Then
                  q(var(v)).Enable = Not (p = 0)
                  RTarget bstack, q(var(v))
                Else
                  MyExec = 0
                  Exit Sub
                End If
End Sub

Function ProcUSE(basestack As basetask, rest$, Lang As Long) As Boolean
Dim ss$, ML As Long, X As Double, pa$, s$, stac1$, p As Variant, frm$, i As Long, w$, pppp As mArray
Dim it As Long, what$
If IsStrExp(basestack, rest$, ss$) Then   'gsb
ElseIf Not Abs(IsLabel(basestack, rest$, ss$)) = 1 Then ' WITHOUT " .gsb"
SyntaxError
Exit Function
End If
ML = 0
If UCase(ss$) = "PIPE" Or UCase(ss$) = "ΑΥΛΟΥ" Then
ML = 1
End If

stac1$ = vbNullString
If FastSymbol(rest$, "!") And ML <> 1 Then
If VALIDATEpart(rest$, s$) Then
Do While s$ <> ""
    If ISSTRINGA(s$, pa$) Then
        basestack.soros.DataStr pa$
    ElseIf IsNumberD2(s$, X) Then
        basestack.soros.DataVal X
        X = vbEmpty
    Else
        Exit Do
    End If
Loop
Else
SyntaxError
Exit Function
End If
Else
If ML <> 1 Then
    Do
        If IsExp(basestack, rest$, p) Then
        stac1$ = stac1$ & Str$(p)
        ElseIf IsStrExp(basestack, rest$, s$) Then
         stac1$ = stac1$ & Sput(s$)
        Else
        Exit Do
        End If
        If Not FastSymbol(rest$, ",") Then Exit Do
    Loop
    pa$ = ExtractPath(ss$)
    para$ = RTrim$(".gsb " & Mid$(ss$, Len(ExtractPath(ss$) + ExtractName(ss$)) + 1))
    If pa$ = vbNullString Then pa$ = mcd
    frm$ = ExtractNameOnly(ss$)
    End If
End If

If Not IsLabelSymbolNew(rest$, "ΣΤΟ", "TO", Lang) Then

w$ = "S" & CStr(Int(Rnd(12) * 100000))

Else

Select Case Abs(IsLabel(basestack, rest$, ss$))

Case 3
    If GetVar(basestack, ss$, i) Then
       w$ = "V" & CStr(i)
       s$ = frm$
       frm$ = var(i)
       var(i) = vbNullString
      Else
     i = globalvar(ss$, "")
             If i <> 0 Then
              w$ = "V" & CStr(i)
              
            var(i) = vbNullString
            End If
                        
     End If
Case 6
   
     If neoGetArray(basestack, ss$, pppp) Then
            If Not NeoGetArrayItem(pppp, basestack, ss$, it, rest$) Then
        MyEr "Not such index for array", "Περίμενα σωστούς δείκτες για πίνακα"
        
        Exit Function
        End If
     Else
     MyEr "Not such array, need to DIM fisrt", "Περίμενα πίνακα, πρέπει να ορίσεις έναν"
: Exit Function
     End If
    
    w$ = "A" & CopyArrayItems(basestack, ss$) + Str(it) ''''''''''εδω για τον νεο πίνακα πρέπει να δώσω το mArray???
    s$ = frm$
    frm$ = pppp.item(it)
    If Typename(pppp.item(it)) = doc Then
    Set pppp.item(it) = New Document
    Else
     pppp.item(it) = vbNullString
    End If
   
    Case Else
    SyntaxError
: Exit Function
   End Select
   If Left$(w$, 1) <> "S" Then

p = GetTaskId + 10000 ' starts from 10000
If Not IsLabelSymbolNew(rest$, "ΩΣ", "AS", Lang) Then
's$ = validpipename(ss$)

If frm$ <> "" Then

ss$ = frm$
Else

ss$ = "M" & CStr(p)
End If

Thing w$, validpipename(ss$)
sThread CLng(p), 0, ss$, w$
TaskMaster.Message CLng(p), 3, CLng(100)
Exit Function
Else
Select Case Abs(IsLabel(basestack, rest$, what$))
Case 0 ' TAKE A NUMBER
If IsNumberLabel(rest$, what$) Then
frm$ = "S" + Right$("0000" + what$, 5)
p = val(what$)
s$ = frm$
If Left$(w$, 1) = "V" Then var(val(Mid$(w$, 2))) = validpipename(frm$)
Else
MyEr "No number found (5 digits)", "Δεν βρήκα αριθμό (5 ψηφία)"
Exit Function
End If
Case 1
    If GetVar(basestack, what, i) Then
    If var(i) < 10000 Then var(i) = p Else p = var(i)
      Else
      globalvar what, p
                             
     End If
Case 5, 7
   
     If neoGetArray(basestack, what, pppp) Then
        If Not NeoGetArrayItem(pppp, basestack, ss$, it, rest$) Then
        MyEr "Not such index for array", "Περίμενα σωστούς δείκτες για πίνακα"
      
        Exit Function
        End If
     Else
     MyEr "Not such array, need to DIM fisrt", "Περίμενα πίνακα, πρέπει να ορίσεις έναν"
     
      Exit Function
     End If
     If pppp.item(it) < 10000 Then pppp.item(it) = p Else p = pppp.item(it)
    Case Else
    MyEr "Wrong parameter", "Λάθος παράμετρος"
     Exit Function
   End Select
   End If
End If
'ss$ = validpipename("M" & CStr(p))
'stac1$ = Sput(ss$) + stac1$
If frm$ <> "" Then
ss$ = frm$
Else
ss$ = "M" & CStr(p)
End If
frm$ = s$
sThread CLng(p), 0, ss$, w$
TaskMaster.Message CLng(p), 3, CLng(100)
ss$ = validpipename(ss$)
stac1$ = Sput(ss$) + stac1$
ss$ = "M" & CStr(p)
End If
If ML <> 1 Then
If stac1$ = vbNullString And Left$(s$, 1) = "S" Then
s$ = App.path
AddDirSep s$
s$ = s$ & "M2000.EXE "
If Shell(s$ & Chr(34) + pa$ & frm$ & ".gsb" & para$ & Chr(34), vbNormalFocus) > 0 Then
End If
End If
If Left$(w$, 1) = "V" Then
ss$ = GetTag$ & ".gsb"
Else
ss$ = w$ & ".gsb"
End If
i = FreeFile
On Error Resume Next
 If Not NeoUnicodeFile(strTemp + ss$) Then
 MyEr "can't save " + strTemp + ss$, "δεν μπορώ να σώσω " + strTemp + ss$
what$ = vbNullString
 Exit Function
End If

Open GetDosPath(strTemp + ss$) For Output As i
If Err.Number > 0 Then
InternalEror
what$ = vbNullString
Exit Function
End If
If stac1$ <> "" Then

' look for unicode...
Print #i, "STACK !" & stac1$ & ": DIR " & Chr(34) + pa$ & Chr(34) & " : LOAD " & Chr(34) + frm$ & para$ & Chr(34)

Else
Print #i, "DIR " & Chr(34) + pa$ & Chr(34) & " : LOAD " & Chr(34) + frm$ & para$ & Chr(34)

End If
Close i
tempList2delete = Sput(strTemp + ss$) + tempList2delete
s$ = App.path
AddDirSep s$
s$ = s$ & "M2000.EXE "
LastUse = MyShell(s$ & Chr(34) + strTemp + ss$ & Chr(34), vbNormalFocus - 4 * (ML <> 0 Or IsSymbol(rest$, ";")))
Sleep 1
If LastUse <> 0 Then

If ML = 0 Then
If IsSymbol(rest$, ";") Then
Else
'AppActivate LastUse
End If
End If
'killfile strTemp + ss$
End If
End If
ProcUSE = True
End Function

Function AddInventory(bstack As basetask, rest$, Optional ret2logical As Boolean = False) As Boolean
Dim p As Variant, s$, pppp As mArray, lastindex As Long
If Not bstack.lastobj Is Nothing Then
If Typename(bstack.lastobj) = "mHandler" Then
Dim aa As mHandler
Set aa = bstack.lastobj
Set bstack.lastobj = Nothing
If Not aa.objref Is Nothing Then
If TypeOf aa.objref Is FastCollection Then
Dim bb As FastCollection
Set bb = aa.objref
If bb.StructLen > 0 Then
MyEr "Structure members are ReadOnly", "Τα μέλη της δομής είναι μόνο για ανάγνωση"
Exit Function
End If

Dim ah As String
FastSymbol rest$, ","
again:
AddInventory = True
ah = aheadstatus(rest$, False) + " "
If InStr(ah, "l") Then
If ret2logical Then ret2logical = False: Exit Function
MyEr "No logical expression", "Όχι λογική έκφραση"
AddInventory = False
Else
If Left$(ah, 1) = "N" Then
    If Not IsExp(bstack, rest$, p) Then
        AddInventory = False
        GoTo there
    End If
    If VarType(p) = vbBoolean Then p = CLng(p)
    If Not bstack.lastobj Is Nothing Then
        If TypeOf bstack.lastobj Is mHandler Then
        If bstack.lastobj.t1 = 4 Then
        Set bstack.lastobj = Nothing
        GoTo noenum
        End If
        End If
        MyEr "No Object Allowed for Key", "Δεν επιτρέπεται αντικείμενο για κλειδί"
        AddInventory = False
        GoTo there
    End If
noenum:
    If bb.ExistKey0(p) Then
        MyEr "Key exist, must be unique", "Το κλειδί υπάρχει, πρέπει να είναι μοναδικό"
        AddInventory = False
        GoTo there
    End If
    bb.AddKey p
ElseIf Left$(ah, 1) = "S" Then
    If Not IsStrExp(bstack, rest$, s$) Then
        AddInventory = False
        GoTo there
    End If
    If Not bstack.lastobj Is Nothing Then
        MyEr "No Object Allowed for Key", "Δεν επιτρέπεται αντικείμενο για κλειδί"
        AddInventory = False
        GoTo there
    End If
    If bb.ExistKey0(s$) Then
        MyEr "Key exist, must be unique", "Το κλειδί υπάρχει, πρέπει να είναι μοναδικό"
        AddInventory = False
        GoTo there
    End If
    bb.AddKey s$

Else
        MyEr "No Key found", "Δεν βρέθηκε κλειδί"
        AddInventory = False
        GoTo there

End If
lastindex = bb.index
If FastSymbol(rest$, ":=", , 2) Then
ah = aheadstatus(rest$, False) + " "
If Left$(ah, 1) = "N" Or InStr(ah, "l") > 0 Then
    If Not IsExp(bstack, rest$, p) Then
        AddInventory = False
        GoTo there
    End If
    bb.index = lastindex
    If Not bstack.lastobj Is Nothing Then
    If TypeOf bstack.lastobj Is mArray Then
        Set pppp = New mArray
        bstack.lastobj.CopyArray pppp
        Set bb.ValueObj = pppp
        Set pppp = Nothing
        
    Else
      Set bb.ValueObj = bstack.lastobj
    End If
        Set bstack.lastobj = Nothing
    Else
        bb.Value = p
    End If
    
ElseIf Left$(ah, 1) = "S" Then
    If Not IsStrExp(bstack, rest$, s$) Then
        AddInventory = False
        GoTo there
    End If
    bb.index = lastindex
    If Not bstack.lastobj Is Nothing Then
    If TypeOf bstack.lastobj Is mArray Then
        Set pppp = New mArray
        bstack.lastobj.CopyArray pppp
        Set bb.ValueObj = pppp
        Set pppp = Nothing
    Else
    
        Set bb.ValueObj = bstack.lastobj
    End If
        Set bstack.lastobj = Nothing
    Else
        bb.Value = s$
    End If

Else
        MyEr "No Data found", "Δεν βρέθηκαν στοιχεία"
        AddInventory = False
        GoTo there

End If


End If


End If
If FastSymbol(rest$, ",") Then GoTo again
there:
Set bb = Nothing
Set aa = Nothing

Exit Function
ElseIf TypeOf aa.objref Is mArray Then
While IsSymbol(rest$, ",")
If Not IsExp(bstack, rest$, p) Then
    MyEr "Expected Array", "Περίμενα Πίνακα"
    Set aa = Nothing
    Set bstack.lastobj = Nothing
    Exit Function
End If
Dim myobject As Object
Set myobject = bstack.lastobj
Set bstack.lastobj = Nothing
If CheckIsmArray(myobject) Then
Set pppp = myobject
pppp.AppendArray aa.objref
Else
    MyEr "Expected Array", "Περίμενα Πίνακα"
    Set aa = Nothing
    Set bstack.lastobj = Nothing
    Exit Function
End If
Wend
Set aa = Nothing
AddInventory = True
Exit Function
End If
End If
End If
MyEr "Wrong type of object (not Inventory or pointer to Array)", "Λάθος τύπος αντικειμένου (όχι Κατάσταση ή δείκτης σε Πίνακα)"
Set aa = Nothing
End If
Set bstack.lastobj = Nothing
End Function

Function IsEnumAs(bstack As basetask, b$, p) As Boolean
Dim aaa As mHandler, useHandler As mHandler, ss$, i As Long, that
If MaybeIsSymbol(b$, ".") Then
            If IsNumber(bstack, b$, that) Then
                If bstack.lastobj Is Nothing Then
                    GoTo aa2
                Else
                    If TypeOf bstack.lastobj Is mHandler Then
                        Set aaa = bstack.lastobj
                        Set bstack.lastobj = Nothing
                        GoTo conthere1001
                    End If
                End If
            End If
            GoTo aa2
            ElseIf IsLabelOnly(b$, ss$) = 1 Then
           
            If GetVar(bstack, myUcase(ss$), i) Then
            If MyIsObject(var(i)) Then
                If TypeOf var(i) Is mHandler Then
                    Set aaa = var(i)
conthere1001:
                    If aaa.t1 = 4 And aaa.IamEnum = False Then
                        Set useHandler = New mHandler
                        useHandler.t1 = 4
                        Set useHandler.objref = aaa.objref
                        useHandler.index_start = 0
                        useHandler.index_cursor = aaa.objref.ZeroValue
                        useHandler.sign = 1
                        Set p = useHandler
                        
                        If FastSymbol(b$, "=") Then
                            If MaybeIsSymbol(b$, ".") Then
                                If IsNumber(bstack, b$, that) Then
                                    If Not bstack.lastobj Is Nothing Then
                                        If TypeOf bstack.lastobj Is mHandler Then
                                            Set aaa = bstack.lastobj
                                            Set bstack.lastobj = Nothing
                                            GoTo conthere1002
                                        End If
                                    End If
                                End If
                                GoTo aa2
                            ElseIf IsLabelOnly(b$, ss$) = 1 Then
                                If GetVar(bstack, myUcase(ss$), i) Then
                                    If TypeOf var(i) Is mHandler Then
                                        Set aaa = var(i)
conthere1002:
                                        If aaa.t1 = 4 Then
                                            If aaa.objref.EnumName = useHandler.objref.EnumName Then
                                                useHandler.index_start = aaa.index_start
                                                useHandler.index_cursor = aaa.index_cursor
                                                useHandler.sign = 1
                                            Else
                                                GoTo aa2
                                            End If
                                        Else
                                            GoTo aa2
                                        End If
                                    Else
                                        GoTo aa2
                                    End If
                                Else
                                    GoTo aa2
                                End If
                            Else
                                GoTo aa2
                            End If
                        End If
                       
                        IsEnumAs = True
                        End If
                    End If
                End If
            
            End If
            Else
aa2:
                ExpectedEnumType
                End If
End Function
Function NewInventory(bstack As basetask, rest$, r, Queue As Boolean) As Boolean
            Dim serr As Boolean
            
                    MakeitObjectInventory r, Queue
                    If Queue Then r.objref.AllowAnyKey
                    Set bstack.lastobj = r
                    If FastSymbol(rest$, ":=", , 2) Then
                    If AddInventory(bstack, rest$, serr) Then
                            Set bstack.lastobj = r
                        r = 0
                        NewInventory = True
                    End If
                    Else
                        Set bstack.lastobj = r
                        r = 0
                        NewInventory = True
                    End If
                    
End Function
Function IsCdate(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
Dim PP As Variant, par As Boolean, r2 As Variant, r3 As Variant, r4 As Variant
   If IsExp(bstack, a$, r, , True) Then
    PP = Abs(r) - Fix(Abs(r))
    r = Abs(r Mod 2958466)
    par = True
    If FastSymbol(a$, ",") Then
    par = IsExp(bstack, a$, r2, , True)
    If FastSymbol(a$, ",") Then
    par = IsExp(bstack, a$, r3, , True) And par
    If FastSymbol(a$, ",") Then
    par = IsExp(bstack, a$, r4, , True) And par
    
    End If
    End If
    End If
    
    If Not par Then
     MissParam a$
     Exit Function
                End If
                On Error Resume Next
     r = CDbl(DateSerial(Year(r) + r2, Month(r) + r3, Day(r) + r4) + PP)
     If SG < 0 Then r = -r
              If Err.Number > 0 Then
    WrongArgument a$
    Err.clear
    Exit Function
    End If
    On Error GoTo 0
 IsCdate = FastSymbol(a$, ")", True)
   Else
   
     MissParam a$
    
    End If
    
End Function
Function IsTimeVal(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
Dim s$
    If IsStrExp(bstack, a$, s$) Then
    On Error Resume Next
    If s$ = "UTC" Then
    r = CDbl(GetUTCTime)
    r = r - Int(r)
    Else
    r = CDbl(CDate(TimeValue(s$)))
    End If
    If SG < 0 Then r = -r
         If Err.Number > 0 Then
    
    WrongArgument a$
    Err.clear
    Exit Function
    End If
        On Error GoTo 0
    
    
    Else
     Dim useHandler As mHandler
     Set useHandler = New mHandler
     useHandler.t1 = 1
     useHandler.ReadOnly = True
     Set useHandler.objref = zones
        Set bstack.lastobj = useHandler
     r = r - r
    End If
IsTimeVal = FastSymbol(a$, ")", True)
End Function
Function IsDataVal(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
 Dim s$, p
    If IsStrExp(bstack, a$, s$) Then
    If FastSymbol(a$, ",") Then
    If Not IsExp(bstack, a$, p) Then
        p = cLid
    End If
    On Error Resume Next
    r = CDbl(DateFromString(s$, p))
    If SG < 0 Then r = -r
     If Err.Number > 0 Then
    
    WrongArgument a$
    Err.clear
    Exit Function
    End If
    Else
    On Error Resume Next
    If s$ = "UTC" Then
    r = CDbl(Int(GetUTCDate))
    Else
    r = CDbl(DateValue(s$))
    End If
    If SG < 0 Then r = -r
     If Err.Number > 0 Then
    
    WrongArgument a$
    Err.clear
    Exit Function
    End If
    End If
    On Error GoTo 0
    
    IsDataVal = FastSymbol(a$, ")", True)
      Else
     
                MissParam a$
    End If

End Function
Function IsSymbolNoSpace(a$, c$, Optional l As Long = 1) As Boolean
' not for greek identifiers. see isStr1()
    Dim j As Long
    j = Len(a$)
    If j = 0 Then Exit Function
    If UCase(Mid$(a$, 1, l)) = c$ Then
        a$ = NLtrim$(Mid$(a$, l + 1))
        
        IsSymbolNoSpace = True
    End If
End Function
Function FindItem(bstackstr As basetask, v As Variant, a$, r$, w2 As Long, Optional ByVal wasarr As Boolean = False) As Boolean
Dim useHandler As mHandler, fastcol As FastCollection, pppp As mArray, w1 As Long, p As Variant, s$
'Dim prev As Variant
'Set prev = v
FindItem = True
againtype:
        r$ = Typename(v)
        If r$ = "mHandler" Then
            Set useHandler = v
            Select Case useHandler.t1
            Case 1
                Set fastcol = useHandler.objref
                If FastSymbol(a$, ")(", , 2) Or True Then
                    If IsExp(bstackstr, a$, p) Then
                        If Not fastcol.Find(p) Then GoTo keynotexist
                        
                            If fastcol.IsObj Then
                                w2 = fastcol.index
                               ' Set prev = v
                            
                                Set v = fastcol.ValueObj
                                GoTo againtype
                            Else
                                wasarr = True
                                GoTo checkit
                            End If
                        ElseIf IsStrExp(bstackstr, a$, s$) Then
                            If fastcol.IsObj Then
                                w2 = fastcol.index
                               ' Set prev = v
                                Set v = fastcol.ValueObj
                                GoTo againtype
                            Else
                            If fastcol.StructLen > 0 Then GoTo checkit
                                r$ = Typename(fastcol.Value)
                            End If
                        Else
                            'MissParam a$
                            FindItem = False
                            Exit Function
keynotexist:
                            indexout a$
                            FindItem = False
                            Exit Function
                    End If
                Else
                    ' new
checkit:
                    If fastcol.StructLen > 0 Then
                    FindItem = False
                    Exit Function
                    ElseIf wasarr Then
                    'r$ = Typename(fastcol.Value)
                    FindItem = False
                    Exit Function
                    Else
                    FindItem = False
                    Exit Function
                    End If
                End If
            Case 2
                'r$ = "Buffer"
                    FindItem = False
                    Exit Function
                
            Case 3
                w1 = useHandler.indirect
                If w1 > -1 And w1 <= var2used Then
                                r$ = Typename(var(w1))
                                If r$ = "mHandler" Then Set v = var(w1): GoTo againtype
                    Else
                            r$ = Typename(useHandler.objref)
                                       If FastSymbol(a$, ")(", , 2) Or True Then
                                        If r$ = "mArray" Then
                                            Set pppp = useHandler.objref
                                                If IsExp(bstackstr, a$, p) Then
                                                   pppp.index = p
                                                    If MyIsObject(pppp.Value) Then
                                                    w2 = p
                                                   ' Set prev = v
                                                         Set v = pppp.Value
                                                         wasarr = False
                                                         GoTo againtype
                                                    Else
                                                       ' r$ = Typename(pppp.Value)
                                                    End If
                                                Else

                                                FindItem = False
                                                Exit Function
                                            End If
                                        Else
                                                FindItem = False
                                                Exit Function
                                        End If
                                        
                                        End If
                                        
                                        End If
                                    
            Case 4
                    FindItem = False
                    Exit Function
            Case Else
                r$ = Typename(v.objref)
                
            End Select
        ElseIf Typename(v) = "PropReference" Then
                    FindItem = False
                    Exit Function
        End If
        Set bstackstr.lastobj = Nothing
        Set bstackstr.lastpointer = Nothing
        FindItem = FastSymbol(a$, ")")
        If useHandler Is Nothing Then FindItem = False: Exit Function
        If TypeOf useHandler.objref Is mArray Then
            Set v = useHandler.objref
        Else
         Set pppp = New mArray
            Set pppp.GroupRef = useHandler
            pppp.Arr = False
            w2 = -101
            Set v = pppp
        End If
End Function
Public Function exeSelect(ExecuteLong, once As Boolean, bstack As basetask, b$, v As Long, Lang As Long) As Boolean
Dim ok As Boolean, x1 As Long, y1 As Long, sp As Variant, st As Variant, sw$, slct As Long, ss$
Dim x2 As Long, y2 As Long, p As Variant, w$, DUM As Boolean, i As Long, nd&
        exeSelect = True
                x1 = 0 ' mode numbers using p, sp and st
                ' x1=2 using sw$ w$ ss$

            If IsLabelSymbolNew(b$, "ΜΕ", "CASE", Lang) Then
            
                        If IsExp(bstack, b$, sp, , True) Then
                        x1 = 1
                        ElseIf IsStrExp(bstack, b$, sw$) Then
                        x1 = 2
                        End If
                    If x1 > 0 Then ' SELECT CASE NUMBER or STRING
                        SetNextLine b$
                        While MaybeIsSymbol(b$, "'\")
                        SetNextLine b$
                        Wend
                    slct = 1
                       If NocharsInLine(b$) Then
                       ExpectedCaseorElseorEnd
                        ExecuteLong = 0
                            Exit Function
                            End If
                        Do
                        If NocharsInLine(b$) Then
               
                            Exit Do
                        End If
                                If IsLabelSymbolNew(b$, "ΜΕ", "CASE", Lang) Then  ' WE HAVE CASE
                                If ok Then
                                ExpectedEndSelect
                                ExecuteLong = 0
                                Exit Function
                  
                                    End If
                                If slct > 0 Then         ' WE ARE IN SEARCH
                                Do
                                ' εδώ κοιτάμε τα CASE
                                x2 = 0
                                If x1 = 1 Then
                                If IsExp(bstack, b$, p, , True) Then x2 = 1
                                Else
                                If IsStrExp(bstack, b$, w$) Then x2 = 2
                                End If
                                       If x2 > 0 Then 'WE HAVE NUMBER OR STRING
                                            If IsLabelSymbolNew(b$, "ΕΩΣ", "TO", Lang) Then   ' range ?
                                            y1 = 0
                                               If x1 = 1 Then
                                                    If IsExp(bstack, b$, st, , True) Then y1 = 1

                                                Else
                                                    If IsStrExp(bstack, b$, ss$) Then y1 = 2
                                                End If
                                                If y1 > 0 Then
                                                y2 = 0
                                                   If x1 = 1 Then
                                                    If (sp >= p And sp <= st) Then y2 = 1

                                                Else
                                                    If sw$ >= w$ And sw$ <= ss$ Then y2 = 2
                                                End If
                                                    If y2 > 0 Or slct = -1 Then 'slct=-1 from break
                                                   If slct = 1 Then slct = 0   ' slct=0 we found
                                                   ' start ExecuteLong command or block

                                            End If
                                                Else
                                                    MyEr "Wrong expression type in To clause in Case", "Λάθος τύπος έκφρασης στην Έως στην Με"
                                                    ExecuteLong = 0
                                                    Exit Function
                                                End If
                                            Else
                                            ' NO WE HAVE ONE VALUE...X1 MASTER, X2 ONE VALUE  Y2 FOR LAST CHECK
                                                y2 = 0
                                                If x1 = 1 Then
                                                    If sp = p Then y2 = 1
                                                Else
                                                    If w$ = sw$ Then y2 = 2
                                                End If
                                                If y2 > 0 Or slct = -1 Then ' ONE VALUE
                                                     If slct = 1 Then slct = 0
                                                End If
                                            End If
                                        Else
                                            If x1 = 1 Then
                                                b$ = Str$(sp) & " " & b$
                                            Else
                                                ' HERE............................IS A PROBLEM IF SW$ HAS <3 ASCII CODE
                                                b$ = Sput(sw$) + b$
                                            End If
                                        If IsExp(bstack, b$, p, , True) Then
                                            If p <> 0 Or slct = -1 Then
                                             If slct = 1 Then slct = 0
                                             ' start ExecuteLong command or block
                                             End If
                                        Else
                                        MyEr "Expected logic half expression in Case", "Περίμενα λογική μισή έκφραση στην Με"
                                        ExecuteLong = 0
                                        Exit Function
                                        End If
                                        End If
                                        If slct = 0 Then
                                            If Left$(b$, 4) = vbCrLf + vbCrLf Then
                                                                ExpectedCaseorElseorEnd
                                                                b$ = Mid$(b$, 3)
                                                                ExecuteLong = 0: Exit Function
                                                                End If
                                                    SetNextLine b$
                                                     '    v = Len(b$)
conthere:
                                                        If FastSymbol(b$, "{") Then  ' block
                                                          ss$ = block(b$)
                                                            DUM = False
                                                            i = 1
                                                            ' #3 call a block
                                                            TraceStore bstack, nd&, b$, 0
                                                            Call executeblock(i, bstack, ss$, False, DUM, , True)
                                                           TraceRestore bstack, nd&
                                                            
                                                            
                                                            If i = 1 Then
                                                            FastSymbol b$, "}"
                                                            If Not MaybeIsSymbol(b$, "'\") Then
                                                            If Not Left$(b$, 2) = vbCrLf Then
                                                                ExpectedCommentsOnly
                                                                ExecuteLong = 0: Exit Function
                                                            End If
                                                            End If
                                                            Else
                                                            
                                                            If i = 0 Then
                                                                b$ = ss$ & b$
                                                                ExecuteLong = 0: Exit Function
                                                            ElseIf i = 2 Then
                                                                If Len(ss$) > 0 Then b$ = ss$
                                                                If DUM = True And b$ <> "" Then
                                                                slct = -1
                                                                Else
                                                                GoTo ContGoto
                                                                End If
                                                            ElseIf i = 3 Then
                                                                If Len(ss$) > 0 Then b$ = ss$
                                                               If DUM = True And b$ <> "" Then slct = 0
                                                            End If
                                                            End If
                                                        Else   ' or line
                                                            DUM = True
                                                        
                                                            i = 1
                                                            
                                                            While IsLabelSymbolNew(b$, "ΜΕ", "CASE", Lang)
                                                           SetNextLine b$
                                                            Wend
                                                            ' #4 call one command
                                                            If MaybeIsSymbol(b$, "{") Then
                                                            GoTo conthere
                                                            End If
                                                            once = True
                                                            
                                                            ss$ = GetNextLine(b$) + vbCrLf + "'"
                                                           
                                                            TraceStore bstack, nd&, b$, 3
                                                            Call executeblock(i, bstack, ss$, once, DUM, True, True)
                                                            bstack.addlen = nd&

                                                            If i = 0 Then
                                                                ExecuteLong = 0: Exit Function
                                                            ElseIf i = 1 And ss$ = vbNullString And once Then 'this is an exit ΟΚ3
                                                                
                                                                b$ = vbNullString
                                                                ExecuteLong = 1
                                                                Exit Function
                                                            ElseIf i = 2 Then
                                                                If DUM = True And Len(ss$) > 0 Then
                                                                    slct = -1
                                                                ElseIf Len(ss$) > 0 Then
                                                                    b$ = ss$
                                                                    ExecuteLong = 2
                                                                    once = False
                                                                    Exit Function
                                                                Else
                                                                    ExecuteLong = i
                                                                    b$ = ss$
                                                                    exeSelect = DUM
                                                                End If
                                                            ElseIf i = 3 Then
                                                                If DUM = True And ss$ <> "" Then
                                                                    slct = 0
                                                                Else
                                                                    i = 2
                                                                    b$ = vbNullString
                                                                    exeSelect = True
                                                                End If
                                                            ElseIf i = 5 Then
                                                                ExecuteLong = 2
                                                                exeSelect = False
                                                            End If
                                                        End If
                                                        Exit Do
                                        End If
                                    Loop While FastSymbol(b$, ",")
                                    
                                     End If
                                SetNextLine b$
                                                                     If Left$(b$, 2) = vbCrLf Then
                                                                     ExpectedCaseorElseorEnd
                                                                ExecuteLong = 0
                                                                Exit Function
                                                                End If
                                ' drop case
                                
                                If IsLabelSymbolNew(b$, "ΜΕ", "CASE", Lang, , , True) Then
         
                                ElseIf IsLabelSymbolNew(b$, "ΑΛΛΙΩΣ", "ELSE", Lang, , , True) Then
   
                                ElseIf IsLabelSymbolNew(b$, "ΤΕΛΟΣ", "END", Lang, , , True) Then
                             
                                Else
                                    If FastSymbol(b$, "{") Then
                                           If slct >= 0 Then
                                                    ss$ = block(b$) + "}"
                                                    b$ = NLtrim$(Mid$(b$, 2))
                                            Else
                                                    ss$ = block(b$)
                                                    DUM = False
                                                    i = 1
                                                    ' #7 call block inside Case (Break) ok
                                                    TraceStore bstack, nd&, b$, 0
                                                            Call executeblock(i, bstack, ss$, False, DUM, , True)
                                                            TraceRestore bstack, nd&
                                                            If i = 1 Then
                                                            FastSymbol b$, "}"
                                                            If Not MaybeIsSymbol(b$, "'\") Then
                                                            If Not Left$(b$, 2) = vbCrLf Then
                                                                ExpectedCommentsOnly
                                                                ExecuteLong = 0: Exit Function
                                                            End If
                                                            End If
                                                                    Else
                                                            
                                                            If i = 0 Then
                                                               b$ = ss$ & b$
                                                                ExecuteLong = 0: Exit Function
                                                            ElseIf i = 2 Then
                                                                    If Len(ss$) > 0 Then b$ = ss$
                                                                    If DUM = True And b$ <> "" Then
                                                                    slct = -1
                                                                    Else
                                                                    GoTo ContGoto
                                                                    End If
                                                           
                                                             ElseIf i = 3 Then
                                                                    If Len(ss$) > 0 Then b$ = ss$
                                                                    If DUM = True And b$ <> "" Then slct = 0
        
                                                             End If
                                                            End If
                                            End If
                                     
                                        SetNextLine b$
                                      ElseIf slct < 0 Then
                                                        DUM = True
                                                   
                                                            i = 1
                                                            ' #8 call one command inside Case (Break) ok
                                                            once = True
                                                            
                                                            ss$ = GetNextLine(b$) + vbCrLf + "'"
                                                           
                                                            TraceStore bstack, nd&, b$, 3
                                                            Call executeblock(i, bstack, ss$, once, DUM, True, True)
                                                            bstack.addlen = nd&
                                                            If i = 0 Then
                                                                ExecuteLong = 0: Exit Function
                                                            ElseIf i = 1 And ss$ = vbNullString And once Then  'this is an exit ΟΚ3
                                                                b$ = vbNullString
                                                                ExecuteLong = 1
                                                                Exit Function
                                                            ElseIf i = 2 Then
                                                                If DUM = True And Len(ss$) > 0 Then
                                                                    slct = -1
                                                                ElseIf Len(ss$) > 0 Then
                                                                    b$ = ss$
                                                                    ExecuteLong = 2
                                                                    once = False
                                                                    Exit Function
                                                                Else
                                                                    ExecuteLong = i
                                                                    b$ = ss$
                                                                    exeSelect = DUM
                                                                End If
                                                            ElseIf i = 3 Then
                                                                If DUM = True And ss$ <> "" Then
                                                                    slct = 0
                                                                Else
                                                                    i = 2
                                                                    b$ = vbNullString
                                                                    exeSelect = True
                                                                End If
                                                            ElseIf i = 5 Then
                                                                ExecuteLong = 2
                                                                exeSelect = False
                                                            End If
  
        
                                    SetNextLine b$
                                    Else
                                    SetNextLine b$
                                
                                        End If
                                    
                                End If
                                
                                ElseIf IsLabelSymbolNew(b$, "ΑΛΛΙΩΣ", "ELSE", Lang) Then
                                        IsLabelSymbolNew b$, "ΜΕ", "CASE", Lang
                                           If ok Then
                                ExpectedEndSelect
                                ExecuteLong = 0
                                Exit Function
                                Else
                                    ok = True
                                    End If
                                    SetNextLine b$
                                    If FastSymbol(b$, "{") Then
                                        ss$ = block(b$)
                                    If slct > 0 Then
                                                    DUM = False
                                                    i = 1
                                                    ' #9 call block inside Else
                                                    TraceStore bstack, nd&, b$, 0
                                                       Call executeblock(i, bstack, ss$, False, DUM, , True)
                                                       TraceRestore bstack, nd&
                                                       If i = 1 Then
                                                            FastSymbol b$, "}"
                                                            If Not MaybeIsSymbol(b$, "'\") Then
                                                            If Not Left$(b$, 2) = vbCrLf Then
                                                                ExpectedCommentsOnly
                                                                ExecuteLong = 0: Exit Function
                                                            End If
                                                            End If
                                                            Else
                                                              
                                                            If i = 0 Then
                                                                b$ = ss$ & b$
                                                                ExecuteLong = 0: Exit Function
                                                            ElseIf i = 2 Then
                                                                        If Len(ss$) > 0 Then b$ = ss$
                                                                          If DUM = True And b$ <> "" Then
                                                                            slct = -1
                                                                          ElseIf b$ <> "" Then
                                                                        GoTo ContGoto
                                                                          Else
                                                                            once = True
                                                                            Exit Function
                                                                        End If
                                                            ElseIf i = 3 Then
                                                            If Len(ss$) > 0 Then b$ = ss$
                                                                If DUM = True And b$ <> "" Then slct = 0: b$ = Mid$(b$, 2): GetNextLine (ss$)
                                                            End If
                                                            End If
                                        Else
                                        b$ = NLtrim$(Mid$(b$, 2))
                                        End If
                                    Else
                                    If slct > 0 Then
                                                                       DUM = True
                                             
                                                            i = 1
                                                            ' #10 call one command inside ELSE
                                                            once = True
                                                            DUM = True
                                                            'TraceStore bstack, nd&, b$, 0
                                                            ss$ = GetNextLine(b$) + vbCrLf + "'"
                                                            TraceStore bstack, nd&, b$, 3
                                                            Call executeblock(i, bstack, ss$, once, DUM, True, True)
                                                        
                                                            TraceRestore bstack, nd&
                                                            If i = 0 Then
                                                                    ExecuteLong = 0: Exit Function
                                                            ElseIf i = 1 And ss$ = vbNullString And once Then  'this is an exit
                                                                b$ = vbNullString
                                                                ExecuteLong = 1
                                                                Exit Function
                                                            ElseIf i = 2 Then
                                                              
                                                                          If DUM = True And Len(ss$) > 0 Then
                                                                            slct = -1
                                                                          ElseIf Len(ss$) > 0 Then
                                                                            b$ = ss$
                                                                            ExecuteLong = 2
                                                                            once = False
                                                                             Exit Function
                                                                          Else
                                                                            ExecuteLong = i
                                                                            b$ = ss$
                                                                            exeSelect = DUM
                                                                            Exit Function
                                                                        End If
                                                            ElseIf i = 3 Then
                                                                If DUM = True And ss$ <> "" Then
                                                                    slct = 0
                                                                Else
                                                                    i = 2
                                                                    b$ = vbNullString
                                                                    exeSelect = True
                                                                End If
                                                            ElseIf i = 5 Then
                                                                ExecuteLong = 2
                                                                exeSelect = False
                                              End If
                                    End If
                                End If
                                SetNextLine b$
                                slct = 0
                        ElseIf IsLabelSymbolNew(b$, "ΤΕΛΟΣ", "END", Lang) Then
                            If IsLabelSymbolNew(b$, "ΕΠΙΛΟΓΗΣ", "SELECT", Lang) Then
                                slct = 0
                                Exit Do
                            Else
                                ExpectedEndSelect
                                ExecuteLong = 0
                                Exit Function
                            End If
                        Else
                             If ok Then
                             ExpectedEndSelect2
                             Else
                             ExpectedCaseorElseorEnd2
                             End If
                            ExecuteLong = 0
                            Exit Function
                        End If
  
                        Loop
                        If slct > 0 Then
                        ExecuteLong = 0: Exit Function
                        End If
                        
                    '-----------ENDIF ---------------
                       Else
                        ExecuteLong = 0
                        Exit Function
                    End If
        Else
           ExecuteLong = 0
           Exit Function
        End If
     exeSelect = False
     Exit Function
ContGoto:
        If myexit(bstack) Then ExecuteLong = 1: Exit Function
        If MyTrim$(b$) = vbNullString Or FastSymbol(b$, ":") Then
                ExecuteLong = 0
                MissingLabel
                Exit Function
        Else
        ' GET OUT FOR NEXT
    
        
        i = Abs(IsLabelOnly(b$, w$))

                If i = 1 Then
                
                once = False
                b$ = w$
                ExecuteLong = 2
                Exit Function
                ElseIf i = 0 Then
                If IsNumberLabel(b$, w$) Then
                      once = False
                b$ = w$
                ExecuteLong = 2
                Exit Function
                Else
                 b$ = w$ & b$
                End If
                Else
                b$ = w$ & b$
                
                End If
              End If
End Function
Sub MakeArray(basestack As basetask, frm$, o As Long, rest$, pppp As mArray, Optional lcl As Boolean = False, Optional globalonly As Boolean = False) 'global
Dim p As Variant, X As Variant, i As Long, F As Long, s$, ss$
    MaybeIsSymbolReplace rest$, ")", ChrW(8)
    Select Case o
    Case 5, 6, 7
    If lcl Then
    
      GlobalArr basestack, here$ & "." + basestack.GroupName & frm$, rest$, i, F, True
    Else
    GlobalArr basestack, basestack.GroupName & frm$, rest$, i, F, True, , globalonly
    End If
    p = i
    If i < 0 Then o = 0
    Case Else
    o = 0
    End Select

    
    Select Case o
    Case 5
        X = 0
      If FastSymbol(rest$, "=") Then
            If IsExp(basestack, rest$, X) Then
            
                        If neoGetArray(basestack, frm$, pppp, , globalonly, Not lcl) Then   '' basestack.GroupName & f
                             If Not basestack.lastobj Is Nothing Then
                                                    If Typename(basestack.lastobj) = "Group" Then
                                                        If basestack.lastobj.IamSuperClass Then
                                                    Set pppp.GroupRef = basestack.lastobj.SuperClassList
                                                Else
                                                    Set pppp.GroupRef = basestack.lastobj
                                                    End If
                                                     pppp.IHaveClass = True
                                                    Set basestack.lastobj = Nothing
                                                    pppp.SerialItem 0, 0, 3
                                                    End If
                                         Else
                                                pppp.SerialItem X, 0, 3
                                        End If
                        End If
            Else
                o = 0
            End If
    ElseIf FastSymbol(rest$, "<<", , 2) Then
    
   F = 1
         s$ = aheadstatus(rest$, True, F)
         If F > 0 Then
                s$ = Left$(rest$, F - 1)
                rest$ = Mid$(rest$, F)
                If neoGetArray(basestack, frm$, pppp) Then
                      For i = 0 To pppp.UpperMonoLimit
                        If IsExp(basestack, (s$), X) Then
                                        If Not basestack.lastobj Is Nothing Then
                                                 If Typename(basestack.lastobj) = "Group" Then
                                                     Set pppp.GroupRef = Nothing
                                                     pppp.IHaveClass = False
                                                     If basestack.lastobj.IamSuperClass Then
                                                    Dim myOBJ As Object
                                          pppp.CopyGroupObj basestack.lastobj.SuperClassList, myOBJ
                        
                                              Set myOBJ.SuperClassList = basestack.lastobj.SuperClassList
                                              Set pppp.item(i) = myOBJ
                                              Set myOBJ = Nothing
                                             
                                               Else
                                                  
                                                       Set pppp.item(i) = basestack.lastobj
                                                      End If
                                        ElseIf Typename(basestack.lastobj) = "mHandler" Then
                                        Set pppp.item(i) = basestack.lastobj
                                        ElseIf Typename(basestack.lastobj) = "mStiva" Then
                                        Set pppp.item(i) = basestack.lastobj
                                        ElseIf Typename(basestack.lastobj) = myArray Then
                                        Set pppp.item(i) = basestack.lastobj
                                                    Else
                                                        Set basestack.lastobj = Nothing
                                                        MyEr "object not supported", "Το αντικείμενο δεν υποστηρίζεται"
                                                        GoTo ex1
                                                        Exit For
                                                     End If
                                                     
                                        Else
                                                pppp.item(i) = X
                                        End If
                                        
                        Else
                          Set basestack.lastobj = Nothing

                            MissNumExpr
                            GoTo ex1
                            Exit For
                        End If
                        Next i
                          Set basestack.lastobj = Nothing
            Else
         '   it = 0
                End If
          End If
     
    End If

    Case 7
    X = 0
    If FastSymbol(rest$, "=") Then
    If IsExp(basestack, rest$, X) Then
   If neoGetArray(basestack, frm$, pppp, , Not lcl) Then  '' basestack.GroupName &
    pppp.SerialItem Int(X), 0, 3
    End If
    Else
    o = 0
    End If
    ElseIf FastSymbol(rest$, "<<", , 2) Then
         F = 1
         s$ = aheadstatus(rest$, True, F)
         If F > 0 Then
                s$ = Left$(rest$, F - 1)
                rest$ = Mid$(rest$, F)
                If neoGetArray(basestack, frm$, pppp) Then
                        For i = 0 To pppp.UpperMonoLimit
                        If IsExp(basestack, (s$), X) Then
                            If Typename(basestack.lastobj) = "lambda" Then
                                    Set pppp.item(i) = basestack.lastobj
                            ElseIf basestack.lastobj Is Nothing Then
                                    pppp.item(i) = Int(X)
                            Else
                                Set basestack.lastobj = Nothing
                                   MyEr "Only Lambda objects here", "Μόνο λάμδα αντικείμενα εδώ"
                                   
                                   GoTo ex1
                                   Exit For
                            End If
                        Else
                            Set basestack.lastobj = Nothing
                            
                                            MissNumExpr
                            GoTo ex1
                       
                        End If
                        Next i
                        Set basestack.lastobj = Nothing
            Else
            MissNumExpr
            GoTo ex1
           End If
          End If
    End If
    Case 6
        s$ = vbNullString
    If FastSymbol(rest$, "=") Then
    If IsStrExp(basestack, rest$, s$) Then
    If neoGetArray(basestack, frm$, pppp, , Not lcl) Then ''basestack.GroupName &
    pppp.SerialItem s$, 0, 3
    End If
    End If
    ElseIf FastSymbol(rest$, "<<", , 2) Then
    F = 1
         s$ = aheadstatus(rest$, True, F)
         If F > 0 Then
                s$ = Left$(rest$, F - 1)
                rest$ = Mid$(rest$, F)
                If neoGetArray(basestack, frm$, pppp) Then
               
                        For i = 0 To pppp.UpperMonoLimit
                        If IsStrExp(basestack, (s$), ss$) Then
                            If Typename(basestack.lastobj) = "Group" Then
                                    Set pppp.GroupRef = Nothing
                                    pppp.IHaveClass = False
                                    Set pppp.item(i) = basestack.lastobj
                            ElseIf Typename(basestack.lastobj) = "lambda" Then
                                    Set pppp.item(i) = basestack.lastobj
                            ElseIf Typename(basestack.lastobj) = "mHandler" Then
                                        Set pppp.item(i) = basestack.lastobj
                                        ElseIf Typename(basestack.lastobj) = "mStiva" Then
                                        Set pppp.item(i) = basestack.lastobj
                                        ElseIf Typename(basestack.lastobj) = myArray Then
                                        Set pppp.item(i) = basestack.lastobj
                                    
                            ElseIf basestack.lastobj Is Nothing Then
                                    pppp.item(i) = ss$
                            Else
                                                     Set basestack.lastobj = Nothing
                                                        MyEr "object not supported", "Το αντικείμενο δεν υποστηρίζεται"
                                                        GoTo ex1
                                                        Exit For
                            End If
                        Else
                            Set basestack.lastobj = Nothing
                            MissStringExpr
                            GoTo ex1
                            Exit For
                        End If
                        Next i
                        Set basestack.lastobj = Nothing
            Else

                MissStringExpr
                GoTo ex1
            End If
        End If
    End If
    End Select
    If o = 0 Then
      MyEr "Array dimensions missing ", "Ο πίνακας δεν έχει διαστάσεις "
    rest$ = basestack.GroupName & frm$ & rest$
    End If
ex1:
    Set basestack.lastpointer = Nothing
    
End Sub

Sub MarkIf(bstack As basetask, a As Long, b As Boolean)
Dim s As mStiva2
Set s = bstack.RetStack
s.PushVal b
s.PushVal a
s.PushVal -3  ' mark for IF
End Sub
Function HaveMark(bstack As basetask, a As Long, b As Boolean) As Boolean
Dim s As mStiva2
Set s = bstack.RetStack
If s.Total >= 3 Then
HaveMark = s.LookTopVal = -3
a = s.StackItem(2)
b = s.StackItem(3)
End If
End Function
Function HaveMark2(bstack As basetask) As Boolean
Dim s As mStiva2
Set s = bstack.RetStack
If s.Total >= 3 Then
If s.LookTopVal = -3 Then s.drop 3: HaveMark2 = True
End If
End Function
Sub DropMark(bstack As basetask)
Dim s As mStiva2
Set s = bstack.RetStack
If s.Total >= 3 Then
If s.LookTopVal = -3 Then s.drop 3
End If
End Sub
Function interpret(bstack As basetask, b$, Optional ByPass As Boolean) As Boolean
Dim di As Object, myobject As Object, i As Long, x1 As Long, ok As Boolean, sp As Variant
Set di = bstack.Owner
Dim prive As basket
'b$ = Trim$(b$)
Dim w$, ww#, LLL As Long, sss As Long, v As Long, p As Variant, ss$, sw$, ohere$
Dim pppp As mArray, i1 As Long, Lang As Long
Dim r1 As Long, r2 As Long
' uink$ = VbNullString
di.FontTransparent = True
ohere$ = here$
If Not ByPass Then here$ = vbNullString
bstack.LoadOnly = ByPass
sss = Len(b$)
Do While Len(b$) <> LLL
If LastErNum <> 0 Then Exit Do
LLL = Len(b$)


If FastSymbol(b$, "{") Then
If Not interpret(bstack, block(b$)) Then interpret = False: here$ = ohere$: GoTo there1
If FastSymbol(b$, "}") Then
sss = Len(b$)
GoTo loopcontinue1
'LLL = Len(b$)


Else
interpret = False: here$ = ohere$: GoTo there1
End If
End If
jumpforCR1:
If FastSymbol(b$, vbCrLf, , 2) Then
        While FastSymbol(b$, vbCrLf, , 2)
        Wend
     ''   UINK$ = VbNullString
        sss = LLL
        End If

While MaybeIsSymbol(b$, "\'")
 SetNextLine b$
    sss = Len(b$)
    LLL = sss
Wend
If FastSymbol(b$, ":") Then
sss = LLL
''UINK$ = VbNullString
End If
If NOEXECUTION Then interpret = False: here$ = ohere$: GoTo there1

If NocharsInLine(b$) Then interpret = True: here$ = ohere$: GoTo there1
If IsSymbol(b$, "@") Then
i1 = IsLabelAnew("", b$, w$, Lang)  '' NO FORM AA@BBB ALLOWED HERE
w$ = "@" + w$
GoTo PROCESSCOMMAND   'IS A COMMAND
Else
i1 = IsLabelAnew("", b$, w$, Lang) '' NO FORM AA@BBB ALLOWED HERE
End If
  If trace And (bstack.Process Is Nothing) And Not bypasstrace Then
  If bstack.IamLambda Then
  If pagio$ = "GREEK" Then
  Form2.label1(0) = "ΛΑΜΔΑ()"
  Else
  Form2.label1(0) = "LAMBDA()"
  End If
  Else
    Form2.label1(0) = here$
    End If
    Form2.label1(1) = w$
    Form2.label1(2) = GetStrUntil(vbCrLf, b$ & vbCrLf, False)
 TestShowSub = vbNullString
 TestShowStart = 0
    Set Form2.Process = bstack
    stackshow bstack
    If Not Form1.Visible Then
    Form1.Show , Form5   'OK
    End If

    If STbyST And bstack.IamChild Then
        STbyST = False
        If Not STEXIT Then
        If Not STq Then
        Form2.gList4.ListIndex = 0
        End If
        End If
        Do
        If di.Visible Then di.Refresh
        ProcTask2 bstack
        Loop Until STbyST Or STq Or STEXIT Or NOEXECUTION Or myexit(bstack)
            If Not STEXIT Then
        If Not STq Then
        Form2.gList4.ListIndex = 0
        End If
        End If
        STq = False
        If STEXIT Then
        NOEXECUTION = True
        trace = False
        STEXIT = False
        GoTo there1
        End If
    End If
'Sleep 5
   '' SleepWaitNO 5
    If STEXIT Then
    
    trace = False
    STEXIT = False
    GoTo there1
    Else
    
    End If
End If
Select Case i1
Case 1234
GoTo jumpforCR1
Case 2
NoRef2
interpret = False
GoTo there1
Case 1

    If sss = LLL Then
  If comhash.Find2(w$, i, v) Then
  If v <> 0 Then GoTo PROCESSCOMMAND
  End If
    ss$ = vbNullString
    If MaybeIsSymbol(b$, "/*-+=~^|<>") Then
        If FastOperator(b$, "<=", i, 2, False) Then
        ' LOOK GLOBAL
        If GetVar(bstack, w$, v, True) Then
        w$ = varhash.lastkey
            Mid$(b$, i, 2) = "  "
            
            GoTo assignvalue
        ElseIf GetlocalVar(w$, v) Then
            w$ = varhash.lastkey
            Mid$(b$, i, 2) = "  "
            GoTo assignvalue
        Else
            ' NO SUCH VARIABLE
            interpret = False
            GoTo there1
        End If
        ' do something here
        ElseIf varhash.Find(myUcase(w$), v) Then
        ' CHECK VAR
            If FastOperator(b$, "=", i) Then
assignvalue:
                If MyIsNumeric(var(v)) Then
assignvalue2:
                    If IsExp(bstack, b$, p) Then
assignvalue3:
                        If bstack.lastobj Is Nothing Then
                        If VarType(var(v)) = vbLong Then
                        On Error Resume Next
                            var(v) = CLng(Int(p))
                            If Err.Number > 0 Then OverflowLong: interpret = 0: GoTo there1
                            On Error GoTo 0
                        Else
                            var(v) = p
                        End If
                        Else
checkobject:
                        Set myobject = bstack.lastobj
                            If TypeOf bstack.lastobj Is Group Then ' oh is a group
                                Set bstack.lastobj = Nothing
                                UnFloatGroup bstack, w$, v, myobject, True ' global??
                                Set myobject = Nothing
                            ElseIf CheckIsmArray(myobject) Then
                                    Set var(v) = New mHandler
                                    var(v).t1 = 3
                                    Set var(v).objref = myobject
                                    If TypeOf bstack.lastobj Is mHandler Then
                                    With bstack.lastobj
                                        If .UseIterator Then
                                            var(v).UseIterator = True
                                            var(v).index_start = .index_start
                                            var(v).index_End = .index_End
                                            var(v).index_cursor = .index_cursor
                                        End If
                                        End With
                                    End If
                            ElseIf TypeOf myobject Is mHandler Then
                                
                                If myobject.indirect > -1 Then
                                    Set var(v) = var(myobject.indirect)
                                Else
                                    Set var(v) = myobject
                                End If
                                 With bstack.lastobj
                                        If .UseIterator Then
                                            var(v).UseIterator = True
                                            var(v).index_start = .index_start
                                            var(v).index_End = .index_End
                                            var(v).index_cursor = .index_cursor
                                        End If
                                    End With
                                
                                
                                Set bstack.lastobj = Nothing
                            ElseIf TypeOf myobject Is lambda Then
                                
                                        GlobalSub w$ + "()", "", , , v
                                   Set var(v) = myobject
                                Set bstack.lastobj = Nothing
                            ElseIf TypeOf myobject Is mEvent Then
                             Set var(v) = myobject
                            CopyEvent var(v), bstack
                            Set var(v) = bstack.lastobj
                            ElseIf TypeOf myobject Is VarItem Then
                                
                                var(v) = myobject.ItemVariant
                            Else
                                Set myobject = Nothing
                                Set bstack.lastobj = Nothing
                                If VarType(var(v)) = vbLong Then
                                    NoObjectpAssignTolong
                                Else
                                    NoObjectAssign
                                End If
                                interpret = False: GoTo there1
                            End If
                            Set bstack.lastobj = Nothing
                            Set myobject = Nothing
                        End If
                    ElseIf IsStrExp(bstack, b$, ss$) Then
                    If bstack.lastobj Is Nothing Then
                    If ss$ = vbNullString Then
                    var(v) = 0#
                    Else
                    If IsNumberCheck(ss$, p) Then
                    var(v) = p
                    End If
                    End If
                    Else
                    GoTo checkobject
                    End If
                    Else
                    ' if is string then what???
                    If Typename(bstack.lastobj) = "mHandler" Then
                    GoTo checkobject
                    End If
                        NoValueForVar w$
                        interpret = False
                        GoTo there1
                    End If
                    GoTo loopcontinue1
                    
                Else
                If Not MyIsObject(var(v)) Then
                
                If IsStrExp(bstack, b$, ss$) Then
                If ss$ = vbNullString Then
                    var(v) = 0#
                Else
                 If IsNumberCheck(ss$, p) Then
                    var(v) = p
                    End If
                End If
                GoTo loopcontinue1
                Else
                    MyEr "Expected String expression", "Περίμενα έκφραση Αλφαριθμητική"
                    
                    Exit Function
                End If
                    ElseIf var(v) Is Nothing Then
                        AssigntoNothing  ' Use Declare
                        interpret = False
                        GoTo there1
                    ElseIf TypeOf var(v) Is Group Then
                        If IsExp(bstack, b$, p) Then
                        
                            'If Not TypeOf var(v) Is Group Then
                            'GoTo assignvalue3
                            'Else
                            If var(v).HasSet Then
                                If bstack.lastobj Is Nothing Then
                                    bstack.soros.PushVal p
                                Else
                                    bstack.soros.PushObj bstack.lastobj
                                    Set bstack.lastobj = Nothing
                                End If
                                NeoCall2 ObjPtr(bstack), w$ + "." + ChrW(&H1FFF) + ":=()", ok
                            ElseIf bstack.lastobj Is Nothing Then
                                NeedAGroupInRightExpression
                                interpret = False
                                GoTo there1
                            ElseIf TypeOf bstack.lastobj Is Group Then
                                Set myobject = bstack.lastobj
                                Set bstack.lastobj = Nothing
                                ss$ = bstack.GroupName
                                If var(v).HasValue Or var(v).HasSet Then
                                    PropCantChange
                                    interpret = 0
                                    GoTo there1
                                Else
                                If Len(var(v).GroupName) > Len(w$) Then
                                ' here$ is ""
                                    UnFloatGroupReWriteVars bstack, w$, v, myobject
                                Else
                                    bstack.GroupName = Left$(w$, Len(w$) - Len(var(v).GroupName) + 1)
                                    If Len(var(v).GroupName) > 0 Then
                                        w$ = Left$(var(v).GroupName, Len(var(v).GroupName) - 1)
                                        UnFloatGroupReWriteVars bstack, w$, v, myobject
                                    Else
                                        GroupWrongUse
                                        interpret = 0
                                        GoTo there1
                                    End If
                                End If
                                End If
                                Set myobject = Nothing
                                bstack.GroupName = ss$
                            Else
                                WrongObject
                                interpret = False
                                GoTo there1
                            End If
                            GoTo loopcontinue1
                        Else
noexpression:
                            Set myobject = Nothing
                            Set bstack.lastobj = Nothing
                            MissNumExpr
                            interpret = False
                            GoTo there1
                        End If
                    ElseIf TypeOf var(v) Is PropReference Then
                    If IsExp(bstack, b$, p) Then
                                If FastSymbol(b$, "@") Then
                                    If IsExp(bstack, b$, sp) Then
                                        var(v).index = p: sp = 0
                                    ElseIf IsStrExp(bstack, b$, ss$) Then
                                        var(v).index = ss$: ss$ = vbNullString
                                    End If
                                    var(v).UseIndex = True
                                End If
                            var(v).Value = p
                    
                    Else
                    GoTo noexpression
                    End If
                    GoTo loopcontinue1
                    ElseIf TypeOf var(v) Is Constant Then
                    CantAssignValue
                    interpret = False
                    GoTo there1
                    ElseIf TypeOf var(v) Is lambda Then
                        ' exist and take something else
                        If IsExp(bstack, b$, p) Then
                            If bstack.lastobj Is Nothing Then
                                Expected "lambda", "λάμδα"
                            ElseIf TypeOf bstack.lastobj Is lambda Then
                                Set var(v) = bstack.lastobj
                                Set bstack.lastobj = Nothing
                                GoTo loopcontinue1
                            Else
                                Expected "lambda", "λάμδα"
                            End If
                            interpret = False
                            GoTo there1

                        Else
                            MissNumExpr
                            interpret = False
                            GoTo there1
                        End If
                    ElseIf TypeOf var(v) Is mHandler Then  ' CHECK IF IT IS A HANDLER
                        If IsExp(bstack, b$, p) Then
                            If var(v).ReadOnly Then
                                ReadOnly
                                interpret = False: GoTo there1
                            End If
                            If bstack.lastobj Is Nothing Then
                                MissingObjReturn
                                interpret = False: GoTo there1
                            ElseIf Typename(bstack.lastobj) = "mHandler" Then
                                Set myobject = New mHandler
                                bstack.lastobj.CopyTo myobject
                                If bstack.lastobj.indirect > -0 Then
                                CheckDeepAny myobject
                                bstack.lastobj.indirect = -1
                                Set bstack.lastobj.objref = myobject
                                Set var(v) = bstack.lastobj
                                Set myobject = New mHandler
                                bstack.lastobj.CopyTo myobject
                                
                                End If
                                Set var(v) = myobject
                            ElseIf Typename(bstack.lastobj) = myArray Then
                                Set myobject = New mHandler
                                myobject.t1 = 3
                                Set myobject.objref = bstack.lastobj
                                Set var(v) = myobject
                                
                            Else
                                
                                Set myobject = var(v)
                                myobject.t1 = 0
                                Set myobject.objref = bstack.lastobj
                                
                            End If
                            Set myobject = Nothing
                        Else
                            MissNumExpr
                            interpret = False
                            GoTo there1
                        End If
                        Set bstack.lastobj = Nothing
                        Set myobject = Nothing
                    ElseIf TypeOf var(v) Is mEvent Then
                      If IsExp(bstack, b$, p) Then
                            Set var(v) = bstack.lastobj
                            CopyEvent var(v), bstack
                            Set var(v) = bstack.lastobj
                            Set bstack.lastobj = Nothing
                        Else
                            MissNumExpr
                            interpret = 0
                            GoTo there1
                        End If
                    Else
                        i = 1
                        GoTo somethingelse
                    End If
                End If
            Else
                ' or do something else
                
somethingelse:
                If InStr("/*-+=~^&|<>", Mid$(b$, i, 1)) > 0 Then
                    If InStr("/*-+=~^&|<>!", Mid$(b$, i + 1, 1)) > 0 Then
                        ss$ = Mid$(b$, i, 2)
                        Mid$(b$, i, 2) = "  "
                    Else
                        ss$ = Mid$(b$, i, 1)
                        Mid$(b$, i, 1) = " "
                    End If
                Else
                    GoTo PROCESSCOMMAND
                End If
                On Error GoTo err123456
                If MyIsNumeric(var(v)) Then
                If VarType(var(v)) = vbLong Then
                On Error GoTo forlong
                Select Case ss$
                    Case "="
                        v = globalvar(w$, CLng(Int(p)), , True)
                        GoTo assignvalue2
                    Case "+="
                        If IsExp(bstack, b$, p) Then
                            var(v) = CLng(Int(p) + var(v))
                        Else
                            GoTo noexpression
                        End If
                    Case "-="
                        If IsExp(bstack, b$, p) Then
                            var(v) = CLng(-Int(p) + var(v))
                        Else
                            GoTo noexpression
                        End If
                    Case "*="
                        If IsExp(bstack, b$, p) Then
                            var(v) = CLng(Int(p) * var(v))
                        Else
                            GoTo noexpression
                        End If
                    Case "/="
                        If IsExp(bstack, b$, p) Then
                            If Int(p) = 0 Then
                                DevZero
                                interpret = False
                                GoTo there1
                            End If
                            var(v) = CLng(var(v) / Int(p))
                        Else
                            GoTo noexpression
                        End If
                    Case "-!"
                        var(v) = CLng(-var(v))
                    Case "++"
                        var(v) = CLng(1 + var(v))
                    Case "--"
                        var(v) = CLng(var(v) - 1)
                    Case "~"
                        var(v) = CLng(-1 - (var(v) <> 0))
                    Case Else
                    GoTo PROCESSCOMMAND
                    
                End Select
                On Error GoTo 0
                
                Else
                Select Case ss$
                    Case "="
                        v = globalvar(w$, p, , True)
                        GoTo assignvalue2
                    Case "+="
                        If IsExp(bstack, b$, p) Then
                            var(v) = p + var(v)
                            If RoundDouble Then If VarType(var(v)) = vbDouble Then var(v) = MyRound(var(v), 13)
                        Else
                            GoTo noexpression
                        End If
                    Case "-="
                        If IsExp(bstack, b$, p) Then
                            var(v) = -p + var(v)
                            If RoundDouble Then If VarType(var(v)) = vbDouble Then var(v) = MyRound(var(v), 13)
                        Else
                            GoTo noexpression
                        End If
                    Case "*="
                        If IsExp(bstack, b$, p) Then
                            var(v) = p * var(v)
                             If RoundDouble Then If VarType(var(v)) = vbDouble Then var(v) = MyRound(var(v), 13)
                        Else
                            GoTo noexpression
                        End If
                    Case "/="
                        If IsExp(bstack, b$, p) Then
                            If p = 0 Then
                                DevZero
                                interpret = False
                                GoTo there1
                            End If
                            var(v) = var(v) / p
                            If RoundDouble Then If VarType(var(v)) = vbDouble Then var(v) = MyRound(var(v), 13)
                        Else
                            GoTo noexpression
                        End If
                    Case "-!"
                        var(v) = -var(v)
                    Case "++"
                        var(v) = 1 + var(v)
                    Case "--"
                        var(v) = var(v) - 1
                    Case "~"
                     
                         Select Case VarType(var(v))
                        Case vbBoolean
                            var(v) = Not CBool(var(v))
                        Case vbCurrency
                            var(v) = CCur(Not CBool(var(v)))
                        Case vbDecimal
                            var(v) = CDec(Not CBool(var(v)))
                        Case Else
                            var(v) = CDbl(Not CBool(var(v)))
                        End Select
                        
                        
                    Case Else
                    GoTo PROCESSCOMMAND
                End Select
                On Error Resume Next
               
                End If
                ElseIf TypeOf var(v) Is Group Then
                    If IsExp(bstack, b$, p) Then
                        If bstack.lastobj Is Nothing Then
                            bstack.soros.PushVal p
                        Else
                            bstack.soros.PushObj bstack.lastobj
                            Set bstack.lastobj = Nothing
                        End If
                    End If
                    NeoCall2 ObjPtr(bstack), w$ + "." + ChrW(&H1FFF) + ss$ + "()", ok
                    If Not ok Then
                        If LastErNum = 0 Then
                            MisOperatror (ss$)
                        End If
                        interpret = False
                        GoTo there1
                    End If
                    
                Else
                    Set myobject = var(v)
                    
                    If CheckIsmArray(myobject) Then
                        If IsExp(bstack, b$, p) Then
                            If Not bstack.lastobj Is Nothing Then
                            If TypeOf bstack.lastobj Is mArray Then
                            Set var(v) = New mHandler
                            var(v).t1 = 3
                            Set var(v).objref = bstack.lastobj
                            Else
                                Set myobject = bstack.lastobj
                                If CheckIsmArray(myobject) Then
                                    Set var(v) = New mHandler
                                    var(v).t1 = 3
                                    Set var(v).objref = myobject
                                Else
                                NotArray
                                interpret = False
                                GoTo there1
                                End If
                                End If
                            Else
                                myobject.Compute2 p, ss$
                            End If
                            Set myobject = Nothing
                            Set bstack.lastobj = Nothing
                        Else
                            myobject.Compute3 ss$
                            Set myobject = Nothing
                            Set bstack.lastobj = Nothing
                        End If
                    ElseIf TypeOf myobject Is mHandler Then
                    If myobject.t1 = 4 Then
                        If myobject.ReadOnly Then
                                ReadOnly
                             interpret = False
                                GoTo there1
                        ElseIf ss$ = "++" Then
                        If myobject.index_start < myobject.objref.count - 1 Then
                            myobject.index_start = myobject.index_start + 1
                            myobject.objref.index = myobject.index_start
                            myobject.index_cursor = myobject.objref.Value
                        End If
                        ElseIf ss$ = "--" Then
                    If myobject.index_start > 0 Then
                            myobject.index_start = myobject.index_start - 1
                            myobject.objref.index = myobject.index_start
                            myobject.index_cursor = myobject.objref.Value
                        End If
                        ElseIf ss$ = "-!" Then
                        myobject.sign = -myobject.sign
                        Else
                        NoOperatorForThatObject ss$
                         interpret = False
                            GoTo there1
                        End If
                        End If
                    Else
                    MyEr "Object not support operator " + ss$, "Το αντικείμενο δεν υποστηρίζει το τελεστή " + ss$
                    interpret = False
                    GoTo there1
                    End If
                End If
            End If
            GoTo loopcontinue1
            

        ElseIf Not bstack.StaticCollection Is Nothing Then
            If bstack.ExistVar(w$) Then
                If FastOperator(b$, "=", i) Then
                
                    If IsExp(bstack, b$, p) Then
checkobject1:
                        Set myobject = bstack.lastobj
                        If CheckIsmArray(myobject) Then
                            Set bstack.lastobj = New mHandler
                            bstack.lastobj.t1 = 3
                            Set bstack.lastobj = myobject
                            bstack.SetVarobJ w$, bstack.lastobj
                        Else
                            bstack.SetVar w$, p
                        End If
                        Set myobject = Nothing
                        Set bstack.lastobj = Nothing
                        GoTo loopcontinue1
                    ElseIf IsStrExp(bstack, b$, ss$) Then
                    If bstack.lastobj Is Nothing Then
                    If ss$ = vbNullString Then
                    p = 0
                    Else
                    p = val(ss$)
                    End If
                    End If
                    GoTo checkobject1
                    
                    Else
                        GoTo aproblem1
                    End If
                Else
                    If InStr("/*-+~", Mid$(b$, i, 1)) > 0 Then
                        If InStr("=+-!", Mid$(b$, i + 1, 1)) > 0 Then
                            ss$ = Mid$(b$, i, 2)
                            Mid$(b$, i, 2) = "  "
                        Else
                            ss$ = Mid$(b$, i, 1)
                            Mid$(b$, i, 1) = " "
                        End If
                    End If
                 If Not bstack.AlterVar(w$, p, ss$, False) Then interpret = False: GoTo there1
                GoTo loopcontinue1
                End If
                
                
            End If
            If FastOperator(b$, "=", i) Then ' MAKE A NEW ONE IF FOUND =
                v = globalvar(w$, p, , True)
                GoTo assignvalue
            ElseIf GetVar(bstack, w$, v, True) Then
                    GoTo somethingelse
            End If
        ElseIf FastOperator(b$, "=", i) Then ' MAKE A NEW ONE IF FOUND =
jumpiflocal:
            v = globalvar(w$, p, , True)
            GoTo assignvalue
        ElseIf GetVar(bstack, w$, v, True) Then
        ' CHECK FOR GLOBAL
            GoTo somethingelse
            Else
        GoTo PROCESSCOMMAND
        End If
        
Else


            
          '**********************************************************
PROCESSCOMMAND:
      
            If Trim$(w$) <> "" Then
      
            Select Case w$
        Dim y1 As Long
        Dim x2 As Long, y2 As Long, SBR$, nd&
          Case "CALL", "ΚΑΛΕΣΕ"
        ' CHECK FOR NUMBER...
        If bstack.NoRun Then
            bstack.callx1 = 0
            bstack.callohere = vbNullString
            b$ = NLtrim(b$)
            SetNextLineNL b$
        Else
         If lckfrm > 0 Then lckfrm = sb2used + 1
        NeoCall ObjPtr(bstack), b$, Lang, ok
        If Not ok Then
            interpret = 0
            GoTo there1
        End If
        End If
            Case " ", ChrW(160)
            ' nothing
          '  SSS = Len(B$)
            Case "SLOW", "ΑΡΓΑ"
            extreme = False
            SLOW = True
            interpret = True
            here$ = ohere$
            GoTo there1
            Case "FAST", "ΓΡΗΓΟΡΑ"
            If FastSymbol(b$, "!") Then extreme = True Else extreme = False
            SLOW = False
            interpret = True
            here$ = ohere$
            GoTo there1
            Case "GLOBAL", "ΓΕΝΙΚΟ", "ΓΕΝΙΚΗ", "ΓΕΝΙΚΕΣ", "LOCAL", "ΤΟΠΙΚΑ", "ΤΟΠΙΚΗ", "ΤΟΠΙΚΕΣ"
           b$ = w$ + " " + b$
           interpret = Execute(bstack, b$, True) = 1
           GoTo there1
            
            Case "USER", "ΧΡΗΣΤΗΣ"
      
               ss$ = PurifyPath(GetStrUntil("\", Trim$(GetNextLine(b$) + "\")))
               
                 If ss$ <> "" Then
                    dset
                    
                    userfiles = GetSpecialfolder(CLng(26)) & "\M2000_USER\"
                    
                    If Not isdir(userfiles) Then MkDir userfiles
                
                    
                    ss$ = AddBackslash(userfiles + ss$)
                    
                    If PathMakeDirs(ss$) Or isdir(ss$) Then
                    userfiles = ss$
                    mcd = userfiles
                    original bstack, "CLS"
                    Else

                    PlainBaSket di, players(GetCode(di)), "Bad User Name"
                    End If
                    Else
                    ss$ = UCase(userfiles)
                    DropLeft "\M2000_USER\", ss$
                    If ss$ = vbNullString Then
                    
                    Else
                    
                    PlainBaSket di, players(GetCode(di)), GetStrUntil("\", Tcase(ss$))
                    End If
                    End If
                     interpret = True
            GoTo there1
            Case "TARGET", "ΣΤΟΧΟΣ"
           ' If di.name <> "DIS" And di.name <> "dSprite" Then interpret = False: here$ = OHERE$: goto there1
                If Abs(IsLabel(bstack, b$, w$)) = 1 Then
                    If Not GetVar(bstack, w$, v) Then 'getvar
                     v = globalvar(w$, 0#, , True)
                  ''  x1 = GetVar(bstack, W$, v)
                      End If
                Else
                    interpret = False
                   here$ = ohere$: GoTo there1
                End If
                If Not FastSymbol(b$, ",") Then
                  interpret = False
                  Exit Do
                ElseIf IsStrExp(bstack, b$, ss$) Then  ' COMMAND
                If ss$ = vbNullString Then interpret = False: here$ = ohere$: GoTo there1
                x1 = 1
                y1 = 1
                x2 = -1
                y2 = -1
                nd& = 0
                SBR$ = vbNullString
                On Error GoTo err123456
                  With players(GetCode(di))
                If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p) Then x1 = Abs(p) Mod (.mx + 1)
                If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p) Then y1 = Abs(p) Mod (.My + 1)
                If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p) Then x2 = CLng(p)
                If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p) Then y2 = CLng(p)
                If FastSymbol(b$, ",") Then If IsExp(bstack, b$, p) Then nd& = Abs(p)
                If FastSymbol(b$, ",") Then If Not IsStrExp(bstack, b$, SBR$) Then interpret = False: here$ = ohere$: GoTo there1
                
               
err123456:
                If Err.Number = 6 Then
                OverflowLong
                interpret = False: here$ = ohere$: GoTo there1
                
                End If
              Targets = False
                ReDim Preserve q(UBound(q()) + 1)
              
                q(UBound(q()) - 1) = BoxTarget(bstack, x1, y1, x2, y2, SBR$, nd&, ss$, .Xt, .Yt, .uMineLineSpace)
                End With
                var(v) = UBound(q()) - 1
                Targets = True
                ElseIf IsExp(bstack, b$, p) Then
                  q(var(v)).Enable = Not (p = 0)
                  RTarget bstack, q(var(v))
                Else
                interpret = False
                here$ = ohere$:             GoTo there1
                End If
                Case "ΔΙΑΚΟΠΤΕΣ", "SWITCHES"
                    If IsStrExp(bstack, b$, ss$) Then
                    Switches ss$, bstack.IamChild Or bstack.IamAnEvent  ' NON LOCAL FROM cli OR using SET SWITCHES
                End If
                Case "MONITOR", "ΕΛΕΓΧΟΣ"
                    If IsSupervisor Then
                    prive = players(GetCode(di))
                    
                    monitor bstack, prive, Lang
                    players(GetCode(di)) = prive
                    Else
                    BadCommand
                    End If
                Case "ΣΕΝΑΡΙΟ", "SCRIPT"
                If IsLabelOnly(b$, ss$) Then
                 If GetSub(myUcase(ss$, True), nd&) Then
                           b$ = vbCrLf + sbf(nd&).sb & b$
                   Else
                   b$ = ss$ + " " + b$
                   If IsStrExp(bstack, b$, w$) Then
                           b$ = vbCrLf + w$ + b$
                   Else
                   ' skip
                   End If
                   End If
                ElseIf IsStrExp(bstack, b$, w$) Then
                           b$ = vbCrLf + w$ + b$
                   End If
Case "RETURN", "ΕΠΙΣΤΡΟΦΗ"
    LastErNum = 0
       If IsExp(bstack, b$, p) Then
                If bstack.lastobj Is Nothing Then
                ElseIf Typename(bstack.lastobj) = "mHandler" Then
                        Select Case bstack.lastobj.t1
                           Case 1
                                  If ChangeValues(bstack, b$) Then GoTo loopcontinue1
                                  
                           Case 2
                                    If ChangeValuesMem(bstack, b$, Lang) Then GoTo loopcontinue1
                           Case 3
                                    If ChangeValuesArray(bstack, b$) Then GoTo loopcontinue1
                           End Select
                End If
            ElseIf IsStrExp(bstack, b$, ss$) Then
                    append_table bstack, ss$, b$, True, Lang
                GoTo loopcontinue1
                 End If
  BadUseofReturn
       interpret = False
       GoTo there1
            Case "CONTINUE", "ΣΥΝΕΧΙΣΕ"
            If HaltLevel > 0 Then
                     If NORUN1 Then NORUN1 = False: interpret = True: b$ = vbNullString: GoTo there1   ' send environment....to hell
                    If bstack.IamChild Or bstack.IamAnEvent Then NERR = True: NOEXECUTION = True
                    ExTarget = True: INK$ = Chr(27): UKEY$ = Chr$(27)  ': UINK$ = Chr(27)    ' send escape...for any good reason...
            Else
            GoTo contnoproper
            End If
            Case "CONST", "ΣΤΑΘΕΡΗ", "ΣΤΑΘΕΡΕΣ"
            ConstNew bstack, b$, w$, True, Lang
                    If LastErNum = -1 Then
                    interpret = False
                    GoTo there1
                    End If
                Case "ΤΕΛΟΣ", "END"

                    If NORUN1 Then NORUN1 = False: interpret = True: b$ = vbNullString: GoTo there1   ' send environment....to hell
                    If bstack.IamChild Or bstack.IamAnEvent Then NERR = True: NOEXECUTION = True
                    ExTarget = True: INK$ = Chr(27): UKEY$ = Chr$(27)  ': UINK$ = Chr(27)    ' send escape...for any good reason...
                Case Else
                    LastErNum = 0 ' LastErNum1 = 0
                    LastErName = vbNullString   ' every command from Query call identifier
                    LastErNameGR = vbNullString  ' interpret is like execute without if for repeat while select structures
                    If comhash.Find2(w$, i, v) Then
                        If v <> 0 Then
                            If v = 32 Then
                                If Not Identifier(bstack, w$, b$, True, Lang) Then
                                    If NOEXECUTION Then
                                            MyEr "", ""
                                            interpret = False
                                    End If
                                    here$ = ohere$: GoTo there1
                              Else
                              If bstack.callx1 > 0 Then
                              If bstack.NoRun Then
                              bstack.callx1 = 0
                              bstack.callohere = vbNullString
                              b$ = NLtrim(b$)
                              SetNextLineNL b$
                              ElseIf Not ProcModuleEntry(bstack, "", 0, b$, Lang) Then
                                    If MOUT And b$ = vbNullString Then
                                    Else
                                        MyErMacro b$, "unknown identifier " & w$, "’γνωστο αναγνωριστικό " & w$
                                    End If
                                End If
                                bstack.RemoveOptionals
                                End If
                              GoTo loopcontinue1
                              End If
                           
                         '' ElseIf v = 2000 Then
                          
                          Else
contnoproper:
                            MyEr "No proper command for command line interpreter", "Δεν είναι η κατάλληλη εντολή για τον διερμηνευτή γραμμής"
                            interpret = False
                          
                            here$ = ohere$: GoTo there1
                          End If
                     End If
                    
                    If i <> 0 Then
                     If IsBadCodePtr(i) = 0 Then
                        If Not CallByPtr(i, bstack, b$, Lang) Then
                               If NOEXECUTION Then
                                    MyEr "", ""
                                    interpret = False
                                    End If
                                    here$ = ohere$: GoTo there1
                        End If
                        End If
                    Else
                            If Not Identifier(bstack, w$, b$, Not comhash.Find(w$, i1), Lang) Then
                            
                                    If NOEXECUTION Then
                                    MyEr "", ""
                                    interpret = False
                                    End If
                                    here$ = ohere$: GoTo there1
                            End If
                    End If
                    
                    ElseIf Not Identifier(bstack, w$, b$, Not comhash.Find(w$, i1), Lang) Then
                    
                            If NOEXECUTION Then
                            MyEr "", ""
                            interpret = False
                            End If
                            here$ = ohere$: GoTo there1
                            ElseIf bstack.callx1 > 0 Then
                              If lckfrm > 0 Then lckfrm = sb2used + 1
                              If bstack.NoRun Then
                              bstack.callx1 = 0
                              bstack.callohere = vbNullString
                              b$ = NLtrim(b$)
                              SetNextLineNL b$
                              ElseIf Not ProcModuleEntry(bstack, "", 0, b$, Lang) Then
                                          If MOUT And b$ = vbNullString Then
                                Else
                                    MyErMacro b$, "unknown identifier " & w$, "’γνωστο αναγνωριστικό " & w$
                                End If
                             End If
                       '''     funcno = funcno - 1
                    End If
                    
                End Select
                End If
            End If
        Else
        If w$ <> "" Then
        b$ = w$ & " " & b$
        If Abs(IsLabel(bstack, b$, w$)) Then
        b$ = w$ & " " & b$
         If FindNameForGroup(bstack, w$) Then
 MyEr "Unknown Property " & w$, "’γνωστη ιδιότητα " & w$
 Else
MyEr "Unknown Variable " & w$, "’γνωστη μεταβλητή " & w$
End If

        
        Else

       SyntaxError
        End If
        b$ = vbNullString
        interpret = False
        GoTo there1
        End If
    End If
Case 3

ss$ = vbNullString
        i = 1
        If Len(b$) > 1 Then
        If InStr("/*-+=~^&|<>", Mid$(b$, i, 1)) > 0 Then
        
                    If InStr("/*-+=~^&|<>!", Mid$(b$, i + 1, 1)) > 0 Then
                        ss$ = Mid$(b$, i, 2)
                        Mid$(b$, i, 2) = "  "
                        If ss$ = "<=" Then ss$ = "g"
                    Else
                        ss$ = Mid$(b$, i, 1)
                        Mid$(b$, i, 1) = " "
                    End If
         End If
       End If

If ss$ <> "" Then
            If ss$ = "=" Then
                If GetVar(bstack, w$, v) Then
                sw$ = ss$
                    If IsStrExp(bstack, b$, ss$) Then
                    If Typename$(bstack.lastobj) = "lambda" Then
                                  GlobalSub w$ + "()", "", , , v
                                               Set var(v) = bstack.lastobj
                                                Set bstack.lastobj = Nothing
                    ElseIf Typename$(var(v)) = "Group" Then
                    
                    If sw$ = "g" Then
                           sw$ = ":="
                           If Not var(v).HasSet Then GroupCantSetValue: interpret = False: GoTo there1
                           End If
                           If bstack.lastobj Is Nothing Then
                                bstack.soros.PushStr ss$
                            Else
                                bstack.soros.PushObj bstack.lastobj
                                Set bstack.lastobj = Nothing
                            End If
                            NeoCall2 ObjPtr(bstack), Left$(w$, Len(w$) - 1) + "." + ChrW(&H1FFF) + sw$ + "()", ok
                    If Not ok Then
                        If LastErNum = 0 Then
                            MisOperatror (ss$)
                        End If
                        interpret = False
                        GoTo there1
                    End If
                    ElseIf TypeOf var(v) Is Constant Then
                    CantAssignValue
                    interpret = False
                    GoTo there1
                    Else
                         
                         If CheckVarOnlyNo(var(v), ss$) Then
                           ExpectedObj Typename(var(v))
                           GoTo there1
                         End If
                        End If
                    Else
aproblem1:
                       NoValueForVar w$
                    Exit Do  '???
                    End If
                ElseIf IsStrExp(bstack, b$, ss$) Then
                    
                                If bstack.lastobj Is Nothing Then
              globalvar w$, ss$, , True
            Else
            If Typename$(bstack.lastobj) = "lambda" Then
                       If Not GetVar(bstack, w$, x1, True) Then x1 = globalvar(w$, p, , True)
                             GlobalSub w$ + "()", "", , , x1
                                        Set myobject = bstack.lastobj
                                        Set bstack.lastobj = Nothing
                                        If x1 <> 0 Then
                                        
                                          Set var(x1) = myobject
                                                Set myobject = Nothing
                                           
                                            
                                        End If
            End If
            End If
                ElseIf LastErNum = 0 Then
                                    
                    SyntaxError
                    interpret = False
                    GoTo there1
                    Else
                   Exit Do  '???
                End If
          
            ElseIf ss$ = "+=" Then
                            If GetVar(bstack, w$, v) Then
                                If IsStrExp(bstack, b$, ss$) Then
                                    If MyIsObject(var(v)) Then

                                            NoOperatorForThatObject "+="
                                            
                                            interpret = False
                                            GoTo there1

                                    Else
                                var(v) = CStr(var(v)) + ss$
                                    End If
                                Else
                                    MissStringExpr
                                End If
                            Else
                                ExpectedVariable
                            End If
            Else
            ' one now option
                If GetVar(bstack, w$, v) Then
                        If IsStrExp(bstack, b$, ss$) Then
                             CheckVar var(v), ss$
                        Else
                            NoValueForVar w$
                        Exit Do
                        End If
                Else
                    Nosuchvariable w$
                End If
        End If
End If
          
Case 4
If FastSymbol(b$, "=") Then '................................
           
            If GetVar(bstack, w$, v) Then
                If IsExp(bstack, b$, p) Then
                
                
                If Not bstack.lastobj Is Nothing Then
                        If TypeOf bstack.lastobj Is lambda Then
                        If Typename(var(v)) = "lambda" Then
                                                Set var(v) = bstack.lastobj

                                                Else
                                    GlobalSub w$ + "()", "", , , v
                                               Set var(v) = bstack.lastobj
                                                
                                        End If
                           Set bstack.lastobj = Nothing
                         Else
                       SyntaxError
                        End If
                        ElseIf MyIsObject(var(v)) Then
                        If TypeOf var(v) Is Constant Then
                            CantAssignValue
                            interpret = False
                            GoTo there1
                        Else
                           ExpectedObj Typename(var(v))
                           GoTo there1
                           End If
                        Else
                        var(v) = MyRound(p)
                        End If
                Else
                  MissNumExpr
                Exit Do
                End If
            ElseIf IsExp(bstack, b$, p) Then
             If Not bstack.lastobj Is Nothing Then
                
                If Typename$(bstack.lastobj) = "lambda" Then
                    
                       If Not GetVar(bstack, w$, x1, True) Then x1 = globalvar(w$, p, , True)
                             GlobalSub w$ + "()", "", , , x1
                                        Set myobject = bstack.lastobj
                                        Set bstack.lastobj = Nothing
                                        If x1 <> 0 Then
                                        
                                          Set var(x1) = myobject
                                                Set myobject = Nothing
                                           
                                            
                                        End If
                                        Else
                                SyntaxError
            End If
            Else
            globalvar w$, p, , True
            End If
                ElseIf LastErNum = 0 Then
                                
                SyntaxError
                interpret = False
                GoTo there1
                Else
               Exit Do
            End If
 Else
    If FastSymbol(b$, "+=", , 2) Then
    ss$ = "+"
    ElseIf FastSymbol(b$, "/=", , 2) Then
    ss$ = "/"
    ElseIf FastSymbol(b$, "-=", , 2) Then
    ss$ = "-"
    ElseIf FastSymbol(b$, "*=", , 2) Then
    ss$ = "*"
    ElseIf IsOperator0(b$, "++", 2) Then
    ss$ = "++"
    ElseIf IsOperator0(b$, "--", 2) Then
    ss$ = "--"
    ElseIf IsOperator0(b$, "-!", 2) Then
    ss$ = "-!"
         ElseIf IsOperator0(b$, "~") Then
        ss$ = "!!"
    ElseIf FastSymbol(b$, "<=", , 2) Then
    ss$ = "="
    End If
        If ss$ = vbNullString Then
                    NoValueForVar w$
                    interpret = False
                     GoTo there1
    End If
    If GetVar(bstack, w$, v) Then
        If Len(ss$) = 1 Then
                    If IsExp(bstack, b$, p) Then
                            On Error Resume Next
                            Select Case ss$
                            Case "="
                            var(v) = MyRound(p)
                                Case "+"
                                var(v) = MyRound(p) + MyRound(var(v))
                                Case "*"
                                 var(v) = MyRound(MyRound(p) * MyRound(var(v)))
                                Case "-"
                                var(v) = MyRound(var(v)) - MyRound(p)
                                Case "/"
                                If MyRound(p) = 0 Then
                                   interpret = False
                                 GoTo there1
                                End If
                                 var(v) = MyRound(MyRound(var(v) / MyRound(p)))
                                 Case "!"
                                 var(v) = -1 - (var(v) <> 0)
                            End Select
                            If Err.Number = 6 Then
                            interpret = False
                            GoTo there1
                            End If
                
                    Else
                                   interpret = False
                                 GoTo there1
                    End If
        Else
        If ss$ = "++" Then
        var(v) = 1 + var(v)
        ElseIf ss$ = "--" Then
        var(v) = var(v) - 1
        ElseIf ss$ = "-!" Then
        var(v) = -var(v)
        Else

                      var(v) = -1 - (var(v) <> 0)
        End If
        End If
    Else
                   interpret = False
        GoTo there1
    End If
End If
Case 5

If neoGetArray(bstack, w$, pppp) Then
againarray22:
    If FastSymbol(b$, ")") Then
    'need to found an expression
        If FastSymbol(b$, "=") Then
            If IsExp(bstack, b$, p) Then
                If Not bstack.lastobj Is Nothing Then
                    bstack.lastobj.CopyArray pppp
                    pppp.final = False
                    Set bstack.lastobj = Nothing
                    GoTo loopcontinue1
                End If
            Else
                SyntaxError
            End If
            interpret = False
            GoTo there1
        End If
        End If
If Not NeoGetArrayItem(pppp, bstack, w$, v, b$) Then interpret = False: here$ = ohere$: GoTo there1
On Error Resume Next
If MaybeIsSymbol(b$, ":+-*/~") Then
With pppp
        If IsOperator0(b$, "++", 2) Then
            .item(v) = .itemnumeric(v) + 1
            GoTo loopcontinue1
        ElseIf IsOperator0(b$, "--", 2) Then
            .item(v) = .itemnumeric(v) - 1
            GoTo loopcontinue1
        ElseIf IsOperator(b$, "+=", 2) Then
            If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
            .item(v) = .itemnumeric(v) + p
        ElseIf IsOperator(b$, "-=", 2) Then
            If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
            .item(v) = .itemnumeric(v) - p
        ElseIf IsOperator(b$, "*=", 2) Then
            If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
            .item(v) = .itemnumeric(v) * p
        ElseIf IsOperator(b$, "/=", 2) Then
            If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
            If p = 0 Then
             DevZero
             Else
             .item(v) = pppp.itemnumeric(v) / p
            End If
        ElseIf IsOperator0(b$, "-!", 2) Then
            .item(v) = -.itemnumeric(v)
            GoTo loopcontinue1
        ElseIf IsOperator0(b$, "~") Then
            Select Case VarType(.itemnumeric(v))
            Case vbBoolean
                .item(v) = Not CBool(.itemnumeric(v))
            Case vbInteger
                .item(v) = CInt(Not CBool(.itemnumeric(v)))
            Case vbLong
                .item(v) = CLng(Not CBool(.itemnumeric(v)))
            Case vbCurrency
                .item(v) = CCur(Not CBool(.itemnumeric(v)))
            Case vbDecimal
                .item(v) = CDec(Not CBool(.itemnumeric(v)))
            Case Else
                .item(v) = CDbl(Not CBool(.itemnumeric(v)))
            End Select
            GoTo loopcontinue1
      ElseIf FastSymbol(b$, ":=", , 2) Then

    If IsExp(bstack, b$, p) Then
        .item(v) = p
    ElseIf IsStrExp(bstack, b$, ss$) Then
      If Not MyIsObject(.item(v)) Then
          .item(v) = ss$
          Else
        CheckVar .item(v), ss$
        
        End If

    Else
        Exit Do
    End If
    If FastSymbol(b$, ",") Then v = v + 1: GoTo contarr1
    GoTo loopcontinue1
        End If
.item(v) = MyRound(.itemnumeric(v), 13)
GoTo loopcontinue1
End With
End If


If IsOperator0(b$, ".") Then

If Typename(pppp.item(v)) = "Group" Then
interpret = SpeedGroup(bstack, pppp, "", w$, b$, v)
Set pppp = Nothing
GoTo loopcontinue1
End If
ElseIf IsOperator(b$, "(") Then
If Typename(pppp.item(v)) = myArray Then
Set pppp = pppp.item(v)
GoTo againarray22
End If
ElseIf Not FastSymbol(b$, "=") Then
here$ = ohere$: GoTo there1
End If

If Not IsExp(bstack, b$, p) Then here$ = ohere$: GoTo there1

 If Not bstack.lastobj Is Nothing Then
     Set myobject = pppp.GroupRef
     If pppp.IHaveClass Then

            Set pppp.item(v) = bstack.lastobj
            Set pppp.item(v).LinkRef = myobject
            With pppp.item(v)
                 .HasStrValue = myobject.HasStrValue
                .HasValue = myobject.HasValue
                .HasSet = myobject.HasSet
                .HasParameters = myobject.HasParameters
                .HasParametersSet = myobject.HasParametersSet
                
                Set .SuperClassList = myobject.SuperClassList
                Set .Events = myobject.Events
                .highpriorityoper = myobject.highpriorityoper
                .HasUnary = myobject.HasUnary
            End With
     Else
            If Typename(bstack.lastobj) = "mHandler" Then
                               Set pppp.item(v) = bstack.lastobj
     
            Else
                   If Not bstack.lastobj Is Nothing Then
                          If TypeOf bstack.lastobj Is mArray Then
                                 If bstack.lastobj.Arr Then
                                         Set pppp.item(v) = CopyArray(bstack.lastobj)

                                 Else
  
   
                                            Set pppp.item(v) = bstack.lastobj
                                            If TypeOf bstack.lastobj Is Group Then Set pppp.item(v).LinkRef = myobject
                                 End If
                          Else
                          
                                  Set pppp.item(v) = bstack.lastobj
                                  If TypeOf bstack.lastobj Is Group Then Set pppp.item(v).LinkRef = myobject
                          End If
                   Else
                  
                          Set pppp.item(v) = bstack.lastobj
                          If TypeOf bstack.lastobj Is Group Then Set pppp.item(v).LinkRef = myobject
                   End If
            End If
        End If
     
     Set bstack.lastobj = Nothing
     Else
     If pppp.Arr Then
     pppp.item(v) = p
     ElseIf Typename(pppp.GroupRef) = "PropReference" Then
    
     pppp.GroupRef.Value = p
     End If
    End If
Do While FastSymbol(b$, ",")
If pppp.UpperMonoLimit > v Then
v = v + 1
If Not IsExp(bstack, b$, p) Then here$ = ohere$: GoTo there1
If Not bstack.lastobj Is Nothing Then
     Set myobject = pppp.GroupRef
     If pppp.IHaveClass Then
         Set pppp.item(v) = bstack.lastobj
            
            With pppp.item(v)
                 .HasStrValue = myobject.HasStrValue
                .HasValue = myobject.HasValue
                .HasSet = myobject.HasSet
                .HasParameters = myobject.HasParameters
                .HasParametersSet = myobject.HasParametersSet
                 Set .SuperClassList = myobject.SuperClassList
                Set .Events = myobject.Events
                .highpriorityoper = myobject.highpriorityoper
                .HasUnary = myobject.HasUnary
            End With
        
        
     Else
        Set pppp.item(v) = bstack.lastobj
    End If
    Set pppp.item(v).LinkRef = myobject
    Set bstack.lastobj = Nothing
     Else
pppp.item(v) = p
End If
Else
Exit Do
End If
Loop
Else
interpret = False: here$ = ohere$: GoTo there1
End If
Case 6
If neoGetArray(bstack, w$, pppp) Then
    If FastSymbol(b$, ")") Then
    'need to found an expression
        If FastSymbol(b$, "=") Then
            If IsStrExp(bstack, b$, ss$) Then
                If Not bstack.lastobj Is Nothing Then
                If TypeOf bstack.lastobj Is mHandler Then
                If bstack.lastobj.t1 = 3 Then
                If pppp.Arr Then
         
                        bstack.lastobj.objref.CopyArray pppp
                        pppp.final = False
            
                        Else
                        NotArray
                        End If
                        Else
                        NotArray
                        End If
                Else
                    bstack.lastobj.CopyArray pppp
                    End If
                    Set bstack.lastobj = Nothing
                    GoTo loopcontinue1
                End If
            Else
                SyntaxError
            End If
               interpret = False
            GoTo there1
        End If
        End If
againstrarr22:
If Not NeoGetArrayItem(pppp, bstack, w$, v, b$) Then interpret = False: here$ = ohere$: GoTo there1
On Error Resume Next
If Typename(pppp.item(v)) = myArray And pppp.Arr Then
If FastSymbol(b$, "(") Then
Set pppp = pppp.item(v)
GoTo againstrarr22
End If
End If
If Not FastSymbol(b$, "=") Then
    If FastSymbol(b$, ":=", , 2) Then
contarr1:
    ss$ = Left$(aheadstatus(b$), 1)
        If ss$ = "S" Then
        If Not IsStrExp(bstack, b$, ss$) Then interpret = False: here$ = ohere$: GoTo there1
        Else
        If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
        ss$ = Trim$(Str$(p))
        End If
             If Not MyIsObject(pppp.item(v)) Then
          pppp.item(v) = ss$
          Else
        CheckVar pppp.item(v), ss$
        
        End If
        Do While FastSymbol(b$, ",")
        If pppp.UpperMonoLimit > v Then
        v = v + 1
          ss$ = Left$(aheadstatus(b$), 1)
                        If ss$ = "S" Then
        If Not IsStrExp(bstack, b$, ss$) Then interpret = False: here$ = ohere$: GoTo there1
        Else
        If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
        ss$ = Trim$(Str$(p))
        End If
        
                If Not MyIsObject(pppp.item(v)) Then
                  pppp.item(v) = ss$
                  Else
                CheckVar pppp.item(v), ss$
                
                End If
        Else
        Exit Do
        End If
        Loop
   ElseIf IsOperator(b$, "+=", 2) Then
    If pppp.IsStringItem(v) Then
    If Not IsStrExp(bstack, b$, ss$) Then GoTo st1222
    If bstack.lastobj Is Nothing Then
        pppp.ItemStr(v) = pppp.item(v) + ss$
    Else
st1222:
        MyEr "Need a string", "Χρειάζομαι ένα αλφαριθμητικό"
        interpret = False: here$ = ohere$: GoTo there1
    End If
    Else
    GoTo st1222
    End If
        
    Else
    interpret = False: here$ = ohere$: GoTo there1
    End If
Else
        If Not IsStrExp(bstack, b$, ss$) Then interpret = False: here$ = ohere$: GoTo there1
        
        
    If Not MyIsObject(pppp.item(v)) Then
    If pppp.Arr Then
    If bstack.lastobj Is Nothing Then
        pppp.item(v) = ss$
    
    Else
    If Typename(bstack.lastobj) = myArray Then
    If bstack.lastobj.Arr Then
        Set pppp.item(v) = CopyArray(bstack.lastobj)
    Else
         Set pppp.item(v) = bstack.lastobj.GroupRef
    End If
    Else
        Set pppp.item(v) = bstack.lastobj
        End If
        Set bstack.lastobj = Nothing
        End If
        Else
        pppp.GroupRef.Value = ss$
        End If
    Else
        CheckVar pppp.item(v), ss$
    End If
        Do While FastSymbol(b$, ",")
        If pppp.UpperMonoLimit > v Then
        v = v + 1
                If Not IsStrExp(bstack, b$, ss$) Then here$ = ohere$: GoTo there1
        
                If Not MyIsObject(pppp.item(v)) Then
                  pppp.item(v) = ss$
                  Else
                CheckVar pppp.item(v), ss$
                
                End If
        Else
        Exit Do
        End If
        Loop
End If
Else
interpret = 0: here$ = ohere$: GoTo there1
End If
Case 7
If neoGetArray(bstack, w$, pppp) Then
    If FastSymbol(b$, ")") Then
    'need to found an expression
        If FastSymbol(b$, "=") Then
            If IsStrExp(bstack, b$, ss$) Then
                If Not bstack.lastobj Is Nothing Then
                    bstack.lastobj.CopyArray pppp
                    Set bstack.lastobj = Nothing
                    GoTo loopcontinue1
                End If
            Else
                SyntaxError
            End If
              interpret = False
            GoTo there1
        End If
        End If
againintarr7:
If Not NeoGetArrayItem(pppp, bstack, w$, v, b$) Then interpret = False: here$ = ohere$: GoTo there1
On Error Resume Next
If Typename(pppp.item(v)) = myArray And pppp.Arr Then
If FastSymbol(b$, "(") Then
Set pppp = pppp.item(v)
GoTo againintarr7
End If
End If
If MaybeIsSymbol(b$, "+-*/~") Then
If IsOperator0(b$, "++", 2) Then
pppp.item(v) = pppp.itemnumeric(v) + 1
ElseIf IsOperator0(b$, "--", 2) Then
pppp.item(v) = pppp.itemnumeric(v) - 1
ElseIf IsOperator(b$, "+=", 2) Then
If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
pppp.item(v) = pppp.itemnumeric(v) + MyRound(p)
ElseIf IsOperator(b$, "-=", 2) Then
If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
pppp.item(v) = pppp.itemnumeric(v) - MyRound(p)
ElseIf IsOperator(b$, "*=", 2) Then
If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
pppp.item(v) = MyRound(pppp.itemnumeric(v) * MyRound(p))
ElseIf IsOperator(b$, "/=", 2) Then
If Not IsExp(bstack, b$, p) Then interpret = False: here$ = ohere$: GoTo there1
If MyRound(p) = 0 Then
 DevZero
 Else
 pppp.item(v) = MyRound(pppp.itemnumeric(v) / MyRound(p))
End If
ElseIf IsOperator0(b$, "-!", 2) Then
pppp.item(v) = -pppp.itemnumeric(v)
ElseIf IsOperator0(b$, "~") Then
        With pppp
        Select Case VarType(.itemnumeric(v))
            Case vbBoolean
                .item(v) = Not CBool(.itemnumeric(v))
            Case vbInteger
                .item(v) = CInt(Not CBool(.itemnumeric(v)))
            Case vbLong
                .item(v) = CLng(Not CBool(.itemnumeric(v)))
            Case vbCurrency
                .item(v) = CCur(Not CBool(.itemnumeric(v)))
            Case vbDecimal
                .item(v) = CDec(Not CBool(.itemnumeric(v)))
            Case Else
                .item(v) = CDbl(Not CBool(.itemnumeric(v)))
        End Select
        End With
End If

GoTo loopcontinue1
End If
If Not FastSymbol(b$, "=") Then here$ = ohere$: GoTo there1
If Not IsExp(bstack, b$, p) Then here$ = ohere$: GoTo there1
If Not bstack.lastobj Is Nothing Then
    If TypeOf bstack.lastobj Is mArray Then
                                 If bstack.lastobj.Arr Then
                                         Set pppp.item(v) = CopyArray(bstack.lastobj)

                                 Else
  
   
                                            Set pppp.item(v) = bstack.lastobj
                                            If TypeOf bstack.lastobj Is Group Then Set pppp.item(v).LinkRef = myobject
                                 End If
                          Else
                          
                                  Set pppp.item(v) = bstack.lastobj
                                  If TypeOf bstack.lastobj Is Group Then Set pppp.item(v).LinkRef = myobject
                          End If
Else
p = MyRound(p)

If Err.Number > 0 Then interpret = False: here$ = ohere$: GoTo there1
pppp.item(v) = p
End If
Do While FastSymbol(b$, ",")

If pppp.UpperMonoLimit > v Then
v = v + 1
If Not IsExp(bstack, b$, p) Then here$ = ohere$: GoTo there1
pppp.item(v) = MyRound(p)
Else
Exit Do
End If
Loop
Else
interpret = False: here$ = ohere$: GoTo there1
End If
Case Else
If MaybeIsSymbol(b$, ",-+*/_!@()[];<>|~`'\") Then
SyntaxError
End If
End Select
loopcontinue1:
Loop
here$ = ohere$
If LastErNum = -2 Then
sss = CLng(Execute(bstack, b$, True))
b$ = vbNullString
interpret = False

GoTo there1
forlong:
OverflowLong
interpret = False

GoTo there1


ElseIf LastErNum <> -0 Then
b$ = " "
End If
interpret = b$ = vbNullString
there1:
bstack.LoadOnly = False
End Function
Public Sub PushErrStage(basestack As basetask)
        With basestack.RetStack
                        .PushVal subHash.count
                        .PushVal varhash.count
                        .PushVal sb2used
                        .PushVal basestack.SubLevel
                        .PushVal var2used

                        .PushVal -4
                         basestack.ErrVars = var2used
        End With
       
End Sub
Public Sub PopStagePart(basestack As basetask, Parts As Long)
Dim nok As Boolean, target As Long
        target = basestack.RetStackTotal - Parts
        If target < 0 Then target = 0
        While basestack.RetStackTotal > target
        With basestack.RetStack
        If .LookTopVal = -4 Then
jumphere:
           .drop 1
           basestack.ErrVars = CLng(.PopVal)
        If nok Then
           basestack.SubLevel = CLng(.PopVal)
            var2used = basestack.ErrVars
            sb2used = CLng(.PopVal)
            varhash.ReduceHash CLng(.PopVal), var()
            subHash.ReduceHash CLng(.PopVal), sbf()
            Else
            .drop 4
            End If
        Else

        While basestack.RetStackTotal > target
        Select Case .LookTopVal
        Case -1
            nok = True
        .drop 7
        Case -2
            nok = True
            .drop 5  ' never happen???
        Case -3
            .drop 3
            basestack.UseofIf = basestack.UseofIf - 1
        Case -4
            GoTo jumphere
        Case Else
         .drop 2  ' string in topval (gosub to label)
             nok = True
        End Select
        Wend
        End If
        End With
        Wend
End Sub
Public Sub PopStagePartContinue2(basestack As basetask, Parts As Long)
' drop until find a If
Dim nok As Boolean, target As Long
        target = basestack.RetStackTotal - Parts
        If target < 0 Then target = 0
        While basestack.RetStackTotal > target
        With basestack.RetStack
        If .LookTopVal = -4 Then
jumphere:
           .drop 1
           basestack.ErrVars = CLng(.PopVal)
        If nok Then
           basestack.SubLevel = CLng(.PopVal)
            var2used = basestack.ErrVars
            sb2used = CLng(.PopVal)
            varhash.ReduceHash CLng(.PopVal), var()
            subHash.ReduceHash CLng(.PopVal), sbf()
            basestack.ResetSkip
            Else
            .drop 4
            End If
        Else

        While basestack.RetStackTotal > target
        Select Case .LookTopVal
        Case -1
           ' nok = True
            '.drop 7
            Exit Sub
        Case -2
            nok = True
            .drop 5  ' never happen???
        Case -3
            .drop 3
            basestack.UseofIf = basestack.UseofIf - 1
        Case -4
            GoTo jumphere
        Case Else
            Exit Sub
        End Select
        Wend
        End If
        End With
        Wend
End Sub

Public Sub PopStagePartContinue(basestack As basetask, Parts As Long)
' drop until find a If
Dim nok As Boolean, target As Long
        target = basestack.RetStackTotal - Parts
        If target < 0 Then target = 0
        While basestack.RetStackTotal > target
        With basestack.RetStack
        If .LookTopVal = -4 Then
jumphere:
           .drop 1
           basestack.ErrVars = CLng(.PopVal)
        If nok Then
           basestack.SubLevel = CLng(.PopVal)
            var2used = basestack.ErrVars
            sb2used = CLng(.PopVal)
            varhash.ReduceHash CLng(.PopVal), var()
            subHash.ReduceHash CLng(.PopVal), sbf()
            basestack.ResetSkip
            Else
            .drop 4
            End If
        Else

        While basestack.RetStackTotal > target
        Select Case .LookTopVal
        Case -1
            nok = True
        .drop 7
        basestack.ResetSkip
        Case -2
            nok = True
            .drop 5  ' never happen???
        Case -3
            .drop 3
            basestack.UseofIf = basestack.UseofIf - 1
        Case -4
            GoTo jumphere
        Case Else
            Exit Sub
        End Select
        Wend
        End If
        End With
        Wend
End Sub

Public Sub PopErrStage(basestack As basetask)
Dim nok As Boolean
        With basestack.RetStack
        If .LookTopVal = -4 Then
jumphere:
           .drop 1
           basestack.ErrVars = CLng(.PopVal)
        If nok Then
           basestack.SubLevel = CLng(.PopVal)
            var2used = basestack.ErrVars
            sb2used = CLng(.PopVal)
            varhash.ReduceHash CLng(.PopVal), var()
            subHash.ReduceHash CLng(.PopVal), sbf()
            basestack.ResetSkip
            Else
            .drop 4
            End If
        Else

        While basestack.RetStackTotal > 0
        Select Case .LookTopVal
        Case -1
            nok = True
        .drop 7
        basestack.ResetSkip
        Case -2
            nok = True
            .drop 5  ' never happen???
        Case -3
            .drop 3
            basestack.UseofIf = basestack.UseofIf - 1
        Case -4
            GoTo jumphere
        Case Else
         .drop 2  ' string in topval (gosub to label)
             nok = True
        End Select
        Wend
        End If
        End With
End Sub
Function expanddot(bstack As basetask, w$) As Boolean
Dim i As Integer, j As Long
For i = 1 To Len(w$)
If Mid$(w$, i, 1) = "." Then
    j = j + 1
Else
    Exit For
End If
Next i
w$ = Mid$(w$, j + 1)
If bstack.GetDotNew(w$, j) Then
If Len(here$) > 0 Then
If j = 1 Then
If Len(w$) > Len(here$) Then
    If Left$(w$, Len(here$) + 1) = here$ + "." Then w$ = Mid$(w$, Len(here$) + 2)
End If
End If
End If
expanddot = True
End If

End Function
Public Function GetNextLineNoTrim(c$) As String
Dim i, j$
i = InStr(c$, vbCrLf)
If i = 0 Then GetNextLineNoTrim = c$: c$ = vbNullString Else GetNextLineNoTrim = Left$(c$, i - 1): c$ = Mid$(c$, i)
End Function
Function PrepareLambda(basestask As basetask, myl As lambda, ByVal v As Long, frm$, c As Constant) As Boolean
On Error GoTo 1234
If Typename(var(v)) = "Constant" Then
    Set c = var(v)
    If Not c.flag Then
    InternalError
    PrepareLambda = False
    Exit Function
    End If
    Set myl = c.Value
Else
    Set myl = var(v)
End If
         myl.name = here$
            
            myl.CopyToVar basestask, here$ = vbNullString, var()
            'sbf(0).sb = var(i).code$
            basestask.OriginalCode = -v
            basestask.FuncRec = subHash.LastKnown

            frm$ = myl.code$
PrepareLambda = True
Exit Function
1234
InternalError
PrepareLambda = False

End Function

Sub BackPort(a$)
If Len(a$) = 0 Then a$ = Chr(8) Else Mid$(a$, 1, 1) = Chr(8)
End Sub
Function ExistNum(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
Dim p As Variant, dd As Long, dn As Long, X As Variant, anything As Object, s$
ExistNum = False
    If IsExp(bstack, a$, p) Then
    If Typename(bstack.lastobj) = "mHandler" Then
    Set anything = bstack.lastobj
    Set bstack.lastobj = Nothing
       If Not CheckLastHandler(anything) Then
        InternalError
        ExistNum = False
        Exit Function
        End If
    With anything
        If anything.indirect >= 0 Then
                dd = anything.indirect
                dn = 0
                Do While (TypeOf var(dd) Is mHandler) And dn < 20
                    If var(dd).indirect >= 0 Then dd = var(dd).indirect Else Exit Do
                    dn = dn + 1
                Loop
                If dn = 20 Then
                InternalError
                Exit Function
                End If
                Select Case dd
                Case 1 To var2used
                    r = SG * MyIsObject(var(dd))
                Case Else
                    r = 0
                End Select
                
                ExistNum = FastSymbol(a$, ")", True)
                
                Set anything = Nothing
                Exit Function
        ElseIf TypeOf .objref Is FastCollection Then
            If FastSymbol(a$, ",") Then
                If IsExp(bstack, a$, p, , True) Then
                    If FastSymbol(a$, ",") Then
                        If IsExp(bstack, a$, X, , True) Then
                        X = Int(X)
                        If X = 0 Then
                                r = .objref.FindOne(p, X)
                                r = SG * X
                        ElseIf X > 0 Then
                            dn = X
                            X = 0
                            r = .objref.FindOne(p, X)
                            X = -X + dn - 1
                            r = SG * .objref.FindOne(p, X)
                            Else
                                r = SG * .objref.FindOne(p, X)
                            End If
                        Else
                            MissParam a$
                            Set anything = Nothing
                            Exit Function
                        End If
                    Else
                    r = SG * .objref.Find(p)
                    End If
                ElseIf IsStrExp(bstack, a$, s$) Then
                    Set bstack.lastobj = Nothing
                    If FastSymbol(a$, ",") Then
                        If IsExp(bstack, a$, X) Then
                            X = Int(X)
                            If X = 0 Then
                                r = .objref.FindOne(s$, X)
                                r = SG * X
                        ElseIf X > 0 Then
                            dn = X
                            X = 0
                            r = .objref.FindOne(s$, X)
                            X = -X + dn - 1
                            r = SG * .objref.FindOne(s$, X)
                            Else
                                r = SG * .objref.FindOne(s$, X)
                            End If
                        Else
                            MissParam a$
                            Set anything = Nothing
                            Exit Function
                        End If
                    Else
                        r = SG * .objref.Find(s$)
                        
                    End If
                End If
                
                
                ExistNum = FastSymbol(a$, ")", True)
                
                Set anything = Nothing
                Exit Function
            End If
        End If
    End With
      
    End If
    Set anything = Nothing
    
    MissParam a$
    Set bstack.lastobj = Nothing
    ElseIf IsStrExp(bstack, a$, s$) Then
    s$ = CFname(s$)
    If s$ <> "" Then
      r = SG * (InStr(s$, "*") = 0 And InStr(s$, "?") = 0)
      Else
  
    r = 0
    End If
   
    
    
    ExistNum = FastSymbol(a$, ")", True)
    Else
        MissParam a$
    End If

    
End Function
Function MySwap(bstack As basetask, rest$, Lang As Long) As Boolean
Dim s$, ss$, F As Long, col As Long, x1 As Long, i As Long, pppp As mArray, pppp1 As mArray
    F = Abs(IsLabel(bstack, rest$, s$))
    MySwap = True
    If F = 1 Or F = 4 Then col = 1
    If F = 5 Or F = 7 Then col = 2
    If F = 0 Then MissingnumVar:  Exit Function
    If (F = 3 Or F = 6) And col > 0 Then SyntaxError: MySwap = False:    Exit Function
    If col = 1 Then
        If GetVar(bstack, s$, F) Then
                If Not FastSymbol(rest$, ",") Then MissingnumVar:  Exit Function
                i = Abs(IsLabel(bstack, rest$, ss$))
              If i = 1 Or i = 4 Then
                If GetVar(bstack, ss$, x1) Then
         If MyIsObject(var(F)) Then
            If TypeOf var(F) Is Constant Then
            CantAssignValue
            MySwap = False: Exit Function
            End If
        End If
        If MyIsObject(var(x1)) Then
            If TypeOf var(x1) Is Constant Then
            CantAssignValue
            MySwap = False: Exit Function
            End If
        End If
                    SwapVariant var(F), var(x1)
                    
                    
                Exit Function
                Else
                    Nosuchvariable ss$
                    MySwap = False
                    Exit Function
                End If
            ElseIf i = 5 Or i = 7 Then
                If neoGetArray(bstack, ss$, pppp) Then
                If Not pppp.Arr Then NotArray: Exit Function
                    If Not NeoGetArrayItem(pppp, bstack, ss$, x1, rest$, True) Then Exit Function
                        If MyIsObject(var(F)) Then
                            If TypeOf var(F) Is Constant Then
                            CantAssignValue
                            MySwap = False: Exit Function
                            End If
                        End If
                    SwapVariant2 var(F), pppp, x1
                    
                    
                Else
                
                    NoSwap ss$
                    MySwap = False
                    Exit Function
                End If
            Else
                MissingnumVar
                MySwap = False
                Exit Function
            End If
        Else
            Nosuchvariable s$
            
            Exit Function
        End If
    ElseIf col = 2 Then
        If neoGetArray(bstack, s$, pppp) Then
        If Not pppp.Arr Then NotArray: Exit Function
        If Not NeoGetArrayItem(pppp, bstack, s$, F, rest$) Then Exit Function
            If Not FastSymbol(rest$, ",") Then MissingnumVar:  Exit Function
                i = Abs(IsLabel(bstack, rest$, ss$))
                  
            If i = 1 Or i = 4 Then
                    If GetVar(bstack, ss$, x1) Then
                    If pppp.IHaveClass Then
                            NoSwap ""
                    Else
                             If MyIsObject(var(x1)) Then
                                If TypeOf var(x1) Is Constant Then
                                CantAssignValue
                                MySwap = False: Exit Function
                                End If
                            End If
                    
                    
                         SwapVariant2 var(x1), pppp, F
                     End If
                        
                    Else
                        MissingnumVar
                        MySwap = False
                        Exit Function
                    End If
            ElseIf i = 5 Or i = 7 Then
                    If neoGetArray(bstack, ss$, pppp1) Then
                    If Not pppp1.Arr Then NotArray: Exit Function
                        If Not NeoGetArrayItem(pppp1, bstack, ss$, x1, rest$) Then Exit Function
                   If pppp.IHaveClass Xor Not pppp1.IHaveClass Then
                            
                        SwapVariant3 pppp, F, pppp1, x1
                        If pppp.IHaveClass Then
                            Set pppp.item(F).LinkRef = pppp1.GroupRef
                            Set pppp1.item(x1).LinkRef = pppp.GroupRef
                            End If
                        Else
                        NoSwap ""
                        Exit Function
                        End If
                        
                    Else
                        MissingnumVar
                        
                        Exit Function
                    End If
            Else
                MissingnumVar
                
                Exit Function
            End If
        Else
            MissingnumVar
            
            Exit Function
        End If
    ElseIf F = 3 Then
            If GetVar(bstack, s$, F) Then
            If Not FastSymbol(rest$, ",") Then MissingnumVar:  Exit Function
                i = Abs(IsLabel(bstack, rest$, ss$))
                 If i = 6 Then
                    If Not neoGetArray(bstack, ss$, pppp) Then MissingStrVar:  Exit Function
                    If Not pppp.Arr Then NotArray: Exit Function
                    If Not NeoGetArrayItem(pppp, bstack, ss$, x1, rest$) Then Exit Function
                     If MyIsObject(var(F)) Then
                        If TypeOf var(F) Is Constant Then
                        CantAssignValue
                        MySwap = False: Exit Function
                        End If
                    End If
                    SwapVariant2 var(F), pppp, x1

                ElseIf i = 3 Then
                    If Not GetVar(bstack, ss$, x1) Then: Exit Function
                     If MyIsObject(var(F)) Then
                        If TypeOf var(F) Is Constant Then
                        CantAssignValue
                        MySwap = False: Exit Function
                        End If
                    End If
                   SwapVariant var(F), var(x1)
                Else
                MissFuncParameterStringVar
                MySwap = False
                End If
                
                
            Else
                    
                    MissFuncParameterStringVar
                    MySwap = False
            End If
    ElseIf F = 6 Then
            If neoGetArray(bstack, s$, pppp) Then
            If Not pppp.Arr Then NotArray: Exit Function
                If Not NeoGetArrayItem(pppp, bstack, s$, x1, rest$) Then Exit Function
                If Not FastSymbol(rest$, ",") Then MissingnumVar:  Exit Function
                i = Abs(IsLabel(bstack, rest$, ss$))
     
                If i = 6 Then
                    If Not neoGetArray(bstack, ss$, pppp1) Then MissingStrVar:  Exit Function
                    If Not pppp.Arr Then NotArray: Exit Function
                    If Not NeoGetArrayItem(pppp1, bstack, ss$, i, rest$) Then Exit Function

                   SwapVariant3 pppp, x1, pppp1, i
 
                ElseIf i = 3 Then
                    If Not GetVar(bstack, ss$, i) Then: Exit Function
                    If MyIsObject(var(i)) Then
                        If TypeOf var(i) Is Constant Then
                        CantAssignValue
                        MySwap = False: Exit Function
                        End If
                    End If


                  SwapVariant2 var(i), pppp, x1
                    Else
                MissFuncParameterStringVar
                MySwap = False
                End If
                
                
            Else
                
                MissPar
                MySwap = False
                
            End If
    Else
                 
                MissPar
                MySwap = False
    End If
    Exit Function

End Function
Public Function TraceThis(bstack As basetask, di As Object, b$, w$, SBB$) As Boolean
    TraceThis = True
    PrepareLabel bstack
    Form2.label1(1) = w$
    Form2.label1(2) = GetStrUntil(vbCrLf, b$ & vbCrLf, False)
    If Len(b$) = 0 Then
    WaitShow = 0
    bypassST = False
    Set Form2.Process = bstack
    Exit Function
    Else
        If WaitShow = 0 Or Len(b$) < WaitShow Then
            WaitShow = 0
            If bstack.OriginalCode < 0 Then
            lasttracecode = -bstack.OriginalCode
                SBB$ = GetNextLine((var(-bstack.OriginalCode).code$))
            Else
            lasttracecode = bstack.OriginalCode
                SBB$ = GetNextLine((sbf(Abs(bstack.OriginalCode)).sb))
            End If
            If Left$(SBB$, 10) = "'11001EDIT" Then
                TestShowSub = Mid$(sbf(Abs(bstack.OriginalCode)).sb, Len(SBB$) + 3)
                If TestShowSub = vbNullString Then
                    TestShowSub = Mid$(sbf(FindPrevOriginal(bstack)).sb, Len(SBB$) + 3)
                End If
                If InStr(TestShowSub, b$) = 0 Then
                    WaitShow = Len(b$)
                End If
            Else
                If bstack.OriginalCode <> 0 Then
                    If bstack.OriginalCode < 0 Then
                        TestShowSub = var(-bstack.OriginalCode).code$
                    Else
                        TestShowSub = sbf(Abs(bstack.OriginalCode)).sb
                    End If
                Else
                    If bstack.IamThread Then
                        If bstack.Process Is Nothing Then
                        Else
                            TestShowSub = bstack.Process.CodeData
                        End If
                    Else
                        TestShowSub = b$
                    End If
                End If
            End If
        End If
        If bstack.addlen Then
            If Len(TestShowSub) - bstack.addlen - Len(b$) > 0 Then
                TestShowStart = Len(TestShowSub) - bstack.addlen - Len(b$) + 1
            Else
                TestShowStart = 1
            End If
        Else
            TestShowStart = Len(TestShowSub) - Len(b$) + 1 ' rinstr(TestShowSub, b$)
        End If
        If TestShowStart <= 0 Then
            TestShowStart = rinstr(TestShowSub, Mid$(b$, 2)) - 1
        End If
     bypassST = False
          
    Set Form2.Process = bstack
    stackshow bstack
        
    End If
    
    If Not Form1.Visible Then
        Form1.Show , Form5   'OK
    End If

    If STbyST Then
        STbyST = False
        If Not STEXIT Then
            If Not STq Then
                Form2.gList4.ListIndex = 0
            End If
        End If
        If Not TaskMaster Is Nothing Then
            If TaskMaster.QueueCount > 0 And TaskMaster.Processing Then TaskMaster.StopProcess
        End If
      
        Do
            BLOCKkey = False
            If Not IsWine Then If di.Visible Then di.Refresh
            ProcTask2 bstack
        Loop Until STbyST Or STq Or STEXIT Or bypassST Or NOEXECUTION Or myexit(bstack) Or Not Form2.Visible

        If Not TaskMaster Is Nothing Then
           If TaskMaster.QueueCount > 0 And Not TaskMaster.Processing Then TaskMaster.StartProcess
        End If
        If Not STEXIT Then
            If Not STq Then
                Form2.gList4.ListIndex = 0
            End If
        End If
        STq = False
        If STEXIT Then
            NOEXECUTION = True
            trace = False
            STEXIT = False
            TraceThis = False
            Exit Function
        End If
    Else
If tracecounter > 0 Then If Not IsWine Then MyDoEvents1 Form2, True

    End If
    If STEXIT Then
        trace = False
        STEXIT = False
        TraceThis = False
        Exit Function
    End If
End Function


Function DriveSerial1(bstack As basetask, a$, r As Variant, SG As Variant) As Boolean
    Dim s$
    If IsStrExp(bstack, a$, s$) Then
    r = SG * DriveSerial(Left$(s$, 3))
  
    
    
    DriveSerial1 = FastSymbol(a$, ")", True)
    Else
         MissParam a$
    End If
End Function
Function MakeForm(basestack As basetask, rest$) As Boolean
On Error Resume Next
MakeForm = True
Dim Scr As Object, XX As Single, p As Variant, x1 As Long, y1 As Long, X As Double, Y As Double
Dim w3 As Long, w4 As Long, sX As Double, adjustlinespace As Boolean, SZ As Single, reduce As Single
Dim monitor As Long
reduce = 1
Set Scr = basestack.Owner
'monitor = FindFormSScreen(scr)

Dim basketcode As Long, mAddTwipsTop As Long


If Left$(Typename(Scr), 3) = "Gui" Then
If Typename(Scr) = "GuiM2000" Then FastSymbol rest$, "!": GoTo there1
ElseIf Scr.name = "Form1" Then

Else
If FastSymbol(rest$, "!") Then reduce = 0.9
there1:

basketcode = GetCode(Scr)
'If players(basketcode).double Then SetNormal scr
With players(basketcode)
SetNormal Scr
mAddTwipsTop = .uMineLineSpace  ' the basic

If IsExp(basestack, rest$, p) Then
    If p < 10 Then p = 10
    X = 4
    XX = 4
    If Scr.name = "DIS" Then
    Do
    Y = CDbl(XX)
    XX = CSng(X)
    nForm basestack, XX, w3, w4, mAddTwipsTop  'using line spacing
    If XX > CSng(X) Then X = CDbl(XX)
    
    If Form1.Width * reduce < w3 * p Then Exit Do
    X = X + 0.25
    Loop
 
    
    Else
    Do
    
    Y = CDbl(XX)
    XX = CSng(X)

    nForm basestack, XX, w3, w4, mAddTwipsTop  'using line spacing
    If XX > CSng(X) Then X = CDbl(XX)
    
    If Scr.Width * reduce < w3 * p Then Exit Do
    
    X = X + 0.4
    Loop
    End If
    X = Y
    sX = 0
   
    If FastSymbol(rest$, ",") Then
        If IsExp(basestack, rest$, sX) Then
        '' ok
        
       mAddTwipsTop = 0  ' find a new one
       players(basketcode).MineLineSpace = 0
       players(basketcode).uMineLineSpace = 0
        adjustlinespace = True
    ''    mmx = scr.Width
''mmy = scr.Height
        Else
        MakeForm = False
        MissNumExpr
        Set Scr = Nothing
        Exit Function
        End If
   
End If
If FastSymbol(rest$, ";") And Scr.name = "DIS" Then
adjustlinespace = False
If IsWine Then
    Form1.Move ScrInfo(Console).Left, ScrInfo(Console).top, ScrInfo(Console).Width - 1, ScrInfo(Console).Height - 1
Else
    Form1.Move ScrInfo(Console).Left, ScrInfo(Console).top, ScrInfo(Console).Width, ScrInfo(Console).Height
End If
    Form1.backcolor = players(-1).Paper
    
Sleep 1
End If
nForm basestack, CSng(X), w3, w4, 0
Dim mmx As Long, mmy As Long
If sX = 0 Then
SZ = CSng(X)
mmx = Scr.Width * reduce
 If Scr.name = "DIS" Then
 mmy = CLng(mmx * Form1.Height / Form1.Width) ' WHY 3/4 ??
 Else
 mmy = Scr.Width * reduce
 End If
 players(basketcode).MineLineSpace = mAddTwipsTop
 players(basketcode).uMineLineSpace = mAddTwipsTop
FrameText Scr, SZ, CLng(w3 * p), mmy, players(basketcode).Paper
Else
If Scr.name = "DIS" Then
If (sX * w4) > Form1.Height * reduce Then
Y = Form1.Height * reduce
While sX * w4 > Form1.Height * reduce

XX = Y / (dv20 * sX)

nForm basestack, XX, w3, w4, 0  'using no spacing so we put a lot of lines
X = CDbl(XX)
Y = Y * 0.9
Wend


End If
Else
If sX * w4 > Scr.Height * reduce Then
Y = Scr.Height * reduce
Do While sX * w4 > Scr.Height * reduce

XX = Y / (dv20 * sX)
nForm basestack, XX, w3, w4, 0  'using no spacing so we put a lot of lines
If X = CDbl(XX) Then Exit Do
X = CDbl(XX)
Y = Y * 0.9
Loop


End If

End If
If Scr.name = "DIS" Then
If Not adjustlinespace Then If Scr.Height * reduce >= Form1.Height * reduce - dv15 Then mAddTwipsTop = dv15 * (((Scr.Height * reduce - sX * w4) / sX / 2) \ dv15)
End If
nForm basestack, (X), w3, w4, mAddTwipsTop
SZ = CSng(X)
'If mmx < scr.Width Then
mmx = Scr.Width * reduce


'If mmx < scr.Width Then
mmy = Scr.Height * reduce
If adjustlinespace Then
If Scr.name = "DIS" Then
mAddTwipsTop = dv15 * (((Form1.Height * reduce - sX * w4) / sX / 2) \ dv15)

Else
mAddTwipsTop = dv15 * (((Scr.Height * reduce - sX * w4) / sX / 2) \ dv15)
End If
sX = CLng(sX * (w4 + mAddTwipsTop * 2))
Else
sX = CLng(sX * w4)
End If
players(basketcode).MineLineSpace = mAddTwipsTop
players(basketcode).uMineLineSpace = mAddTwipsTop
FrameText Scr, SZ, CLng(w3 * p), CLng(sX), players(basketcode).Paper, Not (Scr.name = "DIS")

End If


ElseIf FastSymbol(rest$, ";") And Scr.name = "DIS" Then



If Form1.top > VirtualScreenHeight() - 100 Then Form1.top = ScrInfo(Console).top
If IsWine Then
         Form1.Width = ScrInfo(Console).Width - 1
         Form1.Height = ScrInfo(Console).Height
         Form1.Move ScrInfo(Console).Left, ScrInfo(Console).top
Else
        Form1.Move ScrInfo(Console).Left, ScrInfo(Console).top, ScrInfo(Console).Width, ScrInfo(Console).Height
    
End If
Form1.backcolor = players(-1).Paper
Form1.Cls
With players(-1)
        .mysplit = 0
        .MAXXGRAPH = Form1.Width
        .MAXYGRAPH = Form2.Height
        SetText Form1
        End With
MyMode Scr
ElseIf Scr.name = "DIS" Then

w3 = Form1.Left + Scr.Left
w4 = Form1.top + Scr.top
If IsWine And Form1.Width = ScrInfo(Console).Width Then Form1.Width = ScrInfo(Console).Width - dv15
If Form1.top > VirtualScreenHeight() - 100 Then Form1.top = ScrInfo(Console).top: w4 = Form1.top + Scr.top
scrMove00 Scr
If IsWine Then
    If Scr.Width = ScrInfo(Console).Width Then
       Form1.Width = Scr.Width - 1
    Else
        Form1.Width = Scr.Width
    End If
    Form1.Height = Scr.Height
    Form1.Move w3, w4
Else
    Form1.Move w3, w4, Scr.Width, Scr.Height
End If
Form1.Cls
        With players(-1)
        .mysplit = 0
        .MAXXGRAPH = Form1.Width
        .MAXYGRAPH = Form2.Height
        SetText Form1
        End With
SetText Scr

Set Scr = Nothing
Exit Function
Else
'' CROP LAYER
If basketcode > 0 Then
With players(basketcode)
.MAXXGRAPH = .mx * .Xt
.MAXYGRAPH = .My * .Yt
End With
With Form1.dSprite(basestack.tolayer)
.Move .Left, .top, players(basketcode).MAXXGRAPH, players(basketcode).MAXYGRAPH
End With

End If
End If

players(basketcode).MineLineSpace = mAddTwipsTop
players(basketcode).uMineLineSpace = mAddTwipsTop
MakeForm = True
.curpos = 0
.currow = 0

End With
End If
SetText Scr


End Function

Sub ClearLoadedForms()
Dim i As Long, j As Long, start As Long
j = Forms.count
Debug.Print ""
While j > 0
For i = start To Forms.count - 1
If TypeOf Forms(i) Is GuiM2000 Then Unload Forms(i): start = i: Exit For
Next i
j = j - 1
Wend
End Sub
Function getSafeFormList() As LongHash
Dim i As Long, mycol As safeforms
If Not varhash.Find(ChrW(&HFFBF) + here$, i) Then
i = AllocVar()
varhash.ItemCreator ChrW(&HFFBF) + here$, i
Set mycol = New safeforms
Set var(i) = mycol
Else
Set mycol = var(i)
End If
Set getSafeFormList = mycol.mylist
End Function
Function ProcBrowser(bstack As basetask, rest$, Lang As Long) As Boolean
Dim s$, w$, X As Double
ProcBrowser = True
If Not IsStrExp(bstack, rest$, s$) Then

    If Not Abs(IsLabelFileName(bstack, rest$, s$, , w$)) = 1 Then
         If NOEDIT Then
                If Form1.view1.Visible Then
                    Form1.KeyPreview = True
                    ProcTask2 bstack
                    Form1.view1.SetFocus: Form1.KeyPreview = False
                Else
                    Form1.KeyPreview = True
                End If
        End If
            Exit Function
    Else
     s$ = w$ '' low case
    End If
End If
            If FastSymbol(rest$, ",") Then
                    If IsExp(bstack, rest$, X) Then IEX = CLng(X): IESizeX = Form1.ScaleWidth - IEX Else MissNumExpr: ProcBrowser = False: Exit Function
                If FastSymbol(rest$, ",") Then
                    If IsExp(bstack, rest$, X) Then IEY = CLng(X): IESizeY = Form1.ScaleHeight - IEY Else MissNumExpr: ProcBrowser = False: Exit Function
                                If FastSymbol(rest$, ",") Then
                    If IsExp(bstack, rest$, X) Then IESizeX = CLng(X) Else MissNumExpr: ProcBrowser = False: Exit Function
                                    If FastSymbol(rest$, ",") Then
                    If IsExp(bstack, rest$, X) Then IESizeY = CLng(X) Else MissNumExpr: ProcBrowser = False: Exit Function
                 End If
                End If
             End If
           End If
           If IESizeX = 0 Or IESizeY = 0 Then
           IEX = Form1.ScaleWidth / 8
           IEY = Form1.ScaleHeight / 8
           IESizeX = Form1.ScaleWidth * 6 / 8
           IESizeY = Form1.ScaleHeight * 6 / 8
           End If

If myLcase(Left$(s$, 8)) = "https://" Or myLcase(Left$(s$, 7)) = "http://" Or myLcase(Left$(s$, 4)) = "www." Or myLcase(Left$(s$, 6)) = "about:" Then
Form1.IEUP s$
ElseIf s$ <> "" Then
Form1.IEUP "file:" & strTemp + s$
Else
Form1.IEUP ""
Form1.KeyPreview = True
End If
ProcTask2 bstack

End Function

Function MyScore(bstack As basetask, rest$) As Boolean
Dim s$, sX As Double, p As Variant
MyScore = False
If IsExp(bstack, rest$, p) Then
If p >= 1 And p <= 16 Then
If FastSymbol(rest$, ",") Then
If IsExp(bstack, rest$, sX) Then
If FastSymbol(rest$, ",") Then
If IsStrExp(bstack, rest$, s$) Then
voices(p - 1) = s$
BEATS(p - 1) = sX
MyScore = True
End If
End If
End If
End If
End If
End If
End Function

Function MyPlayScore(bstack As basetask, rest$) As Boolean
Dim task As TaskInterface, sX As Double, p As Variant

MyPlayScore = True
If IsExp(bstack, rest$, p) Then
    If p = 0 Then
    TaskMaster.MusicTaskNum = 0
    TaskMaster.OnlyMusic = True
    Do
    TaskMaster.TimerTickNow
    Loop Until TaskMaster.PlayMusic = False
    TaskMaster.OnlyMusic = False   '' forget it in revision 130
   mute = True
    Else
    mute = False
    If FastSymbol(rest$, ",") Then
        If IsExp(bstack, rest$, sX) Then
          If sX < 1 Then
          sX = 0
          Do While TaskMaster.ThrowOne(CLng(p))
          sX = sX - 1
          If sX < -100 Then Exit Do
          Loop
          Else
          Set task = New MusicBox
          Set task.Owner = Form1.DIS
         
          task.Parameters CLng(p), CLng(sX)
          TaskMaster.MusicTaskNum = TaskMaster.MusicTaskNum + 1
          TaskMaster.AddTask task
          End If
          Do While FastSymbol(rest$, ",")
           MyPlayScore = False
        If IsExp(bstack, rest$, p) Then
             If FastSymbol(rest$, ",") Then
                If IsExp(bstack, rest$, sX) Then
                If sX < 1 Then
                        sX = 0
                        Do While TaskMaster.ThrowOne(CLng(p))
                        sX = sX - 1
                        If sX < -100 Then Exit Do
                        Loop
                  Else
                    Set task = New MusicBox
                    Set task.Owner = Form1.DIS
                    task.Parameters CLng(p), CLng(sX)
                    TaskMaster.MusicTaskNum = TaskMaster.MusicTaskNum + 1
                     TaskMaster.AddTask task
              End If
                MyPlayScore = True
                 End If
            End If
        End If
        If MyPlayScore = False Then
          mute = True
        Exit Do
        End If
          Loop
        End If
    End If
    End If
Else

MyPlayScore = False
End If
End Function

Function IdPara(basestack As basetask, rest$, Lang As Long) As Boolean
Dim x1 As Long, y1 As Long, i As Long, it As Long, vvl As Variant
Dim X As Double, Y As Double, s$, what$, w3 As Long, w4 As Long, z As Double
Dim xa As Long, ya As Long
Dim pppp As mArray


IdPara = True
If IsLabelSymbolNew(rest$, "ΣΤΟ", "TO", Lang) Then
        If Not IsExp(basestack, rest$, Y) Then
            MissNumExpr
            IdPara = False
            Exit Function
        Else
        
          Y = Y - 1
                     If Y < 0 Then Y = -1
         If FastSymbol(rest$, ",") Then
                    If IsExp(basestack, rest$, X) Then
                        X = Int(X)
                        If X < 1 Then
                        MyErMacro rest$, "the index base must be >=1", "η βάση δείκτη πρέπει να είναι >=1"
                        
                        Exit Function
                        End If
                   
                    End If
                    If FastSymbol(rest$, ",") Then
                     If IsExp(basestack, rest$, z) Then
                         z = Int(z)
                         If z < 1 Then
                         MyErMacro rest$, "the lenght base must be >=1", "το μήκος πρέπει να είναι >=1"
                         
                         Exit Function
                         End If
                    
                     End If
                    Else
                    z = 0
                    End If
            Else
                X = 0
            End If
        
            x1 = Abs(IsLabel(basestack, rest$, what$))
            If x1 = 3 Then
                    If GetVar(basestack, what$, i) Then
                        If Typename(var(i)) = doc Then
                                If Not FastSymbol(rest$, "=") Then
                                    MissSymbolMyEr "="
                                    IdPara = False
                                    Exit Function
                                Else
                                    If Not IsStrExp(basestack, rest$, s$) Then
                                        MissStringExpr
                                        IdPara = False
                                        Exit Function
                                    Else
                                    If Y = -1 Then
                                    Y = var(i).DocParagraphs
                                    End If
                                   If var(i).ParagraphFromOrder(Y + 1) = -1 Then
                                   CheckVar var(i), s$
                                    ElseIf Y < 1 Then
                                     w3 = var(i).ParagraphFromOrder(1)
                                     w4 = X
                                       If z > 0 Then
                                    var(i).BackSpaceNchars w3, w4 + CLng(z), CLng(z)
                                    End If
                                    If w3 < 1 Then w3 = 1
                                    If Len(s$) > 0 Then var(i).InsertDoc w3, w4, s$
                                    Else
                                    w3 = var(i).ParagraphFromOrder(Y + 1)
                                    w4 = X
                                       If z > 0 Then
                                    var(i).BackSpaceNchars w3, w4 + CLng(z), CLng(z)
                                    End If
                                    If w3 < 1 Then w3 = 1
                                    If Len(s$) > 0 Then var(i).InsertDoc w3, w4, s$
                                    End If
                                    End If
                                End If
                        Else
                             MissingDoc   ' only doc not string var
                             IdPara = False
                            Exit Function
                        End If
                    Else
                        Nosuchvariable what$
                        IdPara = False
                        Exit Function
                    End If
            ElseIf x1 = 6 Then
                    If neoGetArray(basestack, what$, pppp) Then
                        If Not NeoGetArrayItem(pppp, basestack, what$, it, rest$) Then IdPara = False: Exit Function
                        If Typename(pppp.item(it)) = doc Then
                                    If Not FastSymbol(rest$, "=") Then
                                            MissSymbolMyEr "="
                                            IdPara = False
                                            Exit Function
                                            Else
                                If IsStrExp(basestack, rest$, s$) Then
                                
                                
                                    If pppp.item(it).ParagraphFromOrder(Y + 1) = -1 Then
                                       CheckVar pppp.item(it), s$
                                        ElseIf Y < 1 Then
                                   w3 = pppp.item(it).ParagraphFromOrder(1)
                                     w4 = X
                                       If z > 0 Then
                                    pppp.item(it).BackSpaceNchars w3, w4 + CLng(z), CLng(z)
                                    End If
                                    If w3 < 1 Then w3 = 1
                                    If Len(s$) > 0 Then pppp.item(it).InsertDoc w3, w4, s$
                                   
                                        Else
                                        w3 = pppp.item(it).ParagraphFromOrder(Y + 1)
                                    w4 = X
                                       If z > 0 Then
                                    pppp.item(it).BackSpaceNchars w3, w4 + CLng(z), CLng(z)
                                    End If
                                    If w3 < 1 Then w3 = 1
                                    If Len(s$) > 0 Then pppp.item(it).InsertDoc w3, w4, s$
                                    
                                        End If
                                
                                
                                
                                Else
                                    MissStringExpr
                                    IdPara = False
                                    Exit Function
                                
                                End If

                            End If
                        Else
                             MissingDoc   ' only doc not string var
                             IdPara = False
                            Exit Function
                        End If
                    End If
            Else
                MissingDoc   ' only doc not string var
                IdPara = False
                Exit Function
            End If
        End If
 ElseIf IsExp(basestack, rest$, X) Then
    X = Int(X)
    If X < 1 Then
    MyErMacro rest$, "the index base must be >=1", "η βάση δείκτη πρέπει να είναι >=1"
    ' not needed to change idpara must be true because macro embed an ERROR command
    Exit Function
    End If
    If FastSymbol(rest$, ",") Then
        If Not IsExp(basestack, rest$, Y) Then
        MissNumExpr
        IdPara = False
        Exit Function
        End If
        Y = Int(Y)
        If Y < 0 Then
            MyErMacro rest$, "number to delete chars must positive or zero", "ο αριθμός για να διαγράψω πρέπει να είναι θετικός ή μηδέν"
            Exit Function
        End If
    Else
    Y = 0  ' only insert
    End If

     x1 = Abs(IsLabel(basestack, rest$, what$))
        If x1 = 3 Then
            If GetVar(basestack, what$, i) Then
        
                If Typename(var(i)) = doc Then
                    If Not FastSymbol(rest$, "=") Then
                    MissSymbolMyEr "="
                    IdPara = False
                    Exit Function
                    Else
                            If Not IsStrExp(basestack, rest$, s$) Then
                                MissStringExpr
                                IdPara = False
                                Exit Function
                            Else
                                    If Y = 0 Then
                                           var(i).FindPos 1, 0, CLng(X), x1, y1, w3, w4
                                           If w4 = 0 Then
                                          ' ' merge to previous
                                           End If
       
                                    Else
                                             var(i).FindPos 1, 0, X + Y, x1, y1, w3, w4
                                            ' so now we now the paragraph w3 and the position w4
                                            var(i).BackSpaceNchars w3, w4, Y
                                    End If
                                    If s$ <> "" Then var(i).InsertDoc w3, w4, s$
                            End If
                     End If
                ElseIf Typename(var(i)) = "Constant" Then
                CantAssignValue
                    IdPara = False
            Exit Function
                
                Else
                    If Not FastSymbol(rest$, "=") Then
                    MissSymbolMyEr "="
                    IdPara = False
                    Exit Function
                    Else
                    If Not IsStrExp(basestack, rest$, s$) Then
                                MissStringExpr
                                IdPara = False
                                Exit Function
                            Else
                                    If Y = 0 Then
                                        var(i) = Left$(var(i), X - 1) & s$ & Mid$(var(i), X)
                                    Else
                                        If s$ = vbNullString Then
                                        var(i) = Left$(var(i), X - 1) & Mid$(var(i), X + Y)
                                        Else
                                        If Len(s$) = Y Then
                                        Mid$(var(i), X, Y) = s$
                                        ElseIf Len(s$) < Y Then
                                        Mid$(var(i), X, Y) = s$ + space$(Y - Len(s$))
                                        Else
                                        var(i) = Left$(var(i), X - 1) & s$ & Mid$(var(i), X + Y)
                                        End If
                                        End If
                                    End If
                            End If
                    End If
                
                End If
            Else
            Nosuchvariable what$
            IdPara = False
            Exit Function
            
            End If
        ElseIf x1 = 6 Then
        
        
        If neoGetArray(basestack, what$, pppp) Then
                If Not NeoGetArrayItem(pppp, basestack, what$, it, rest$) Then IdPara = False: Exit Function
                If Typename(pppp.item(it)) = doc Then
                    If FastSymbol(rest$, "=") Then
                        If IsStrExp(basestack, rest$, s$) Then
                      If Y = 0 Then
                                     pppp.item(it).FindPos 1, 0, CLng(X), xa, ya, w3, w4
                                           If w4 = 0 Then
                                          ' ' merge to previous
                                           End If

                      Else
                                     pppp.item(it).FindPos 1, 0, X + Y, xa, ya, w3, w4
                                            ' so now we now the paragraph w3 and the position w4
                                            pppp.item(it).BackSpaceNchars w3, w4, Y
                      End If
                       If s$ <> "" Then pppp.item(it).InsertDoc w3, w4, s$
                        Else
                            MissStringExpr
                            IdPara = False
                        End If
                    End If
                Else
                If FastSymbol(rest$, "=") Then
                If IsStrExp(basestack, rest$, s$) Then
                If Y = 0 Then
                    pppp.item(it) = Left$(pppp.item(it), X - 1) & s$ & Mid$(var(i), X)
                Else
                                                        If s$ = vbNullString Then
                                        pppp.item(it) = Left$(pppp.item(it), X - 1) & Mid$(pppp.item(it), X + Y)
                                        Else
                                      vvl = pppp.item(it)
                                       If Len(vvl) = Y Then
                                      
                                        Mid$(vvl, X, Y) = s$
                                        ElseIf Len(s$) < Y Then
                                            Mid$(vvl, X, Y) = s$ + space$(Y - Len(s$))
                                        Else
                                        vvl = Left$(vvl, X - 1) & s$ & Mid$(vvl, X + Y)
                                        End If
                                        pppp.item(it) = vvl
                                        End If
                End If
                Else
                     MissStringExpr
                            IdPara = False
                End If
                End If
                End If
        Else
            IdPara = True
        End If
        
        
        
        Else
        MissingStrVar
        IdPara = False
        ' wrong parameter
        End If


 
 
End If

End Function
Sub stackshow(b As basetask)
Static OldPagio$
Dim p As Variant, r$, AL$, s$, dl$, dl2$
Static once As Boolean, ok As Boolean
If once Then Exit Sub
once = True

If TestShowCode Then
With Form2.testpad
.enabled = True
.SelectionColor = rgb(255, 64, 128)
.nowrap = True
.Text = TestShowSub
If Len(Form2.label1(1)) > 0 Then
If AscW(Form2.label1(1)) = 8191 Then
.SelStartSilent = TestShowStart - 1
.SelLength = Len(Mid$(Form2.label1(1), 7))
Else
.SelStartSilent = TestShowStart - Len(Form2.label1(1)) - 1
.SelLength = Len(Form2.label1(1))
End If


.enabled = False
If .SelLength > 1 And Not AscW(Form2.label1(1)) = 8191 Then
If Not myUcase(.SelText, True) = Form2.label1(1) Then
End If
End If
Else
.enabled = False
End If
''Debug.Print b.addlen
'MyDoEvents
End With

once = False
Exit Sub
Else
Form2.testpad.nowrap = False
End If

If pagio$ <> OldPagio$ Then
Form2.FillAgainLabels
OldPagio$ = pagio$
End If


Dim stack As mStiva
Set stack = b.soros

If Form2.Compute <> "" Then
If Form2.Compute.Prompt = "? " Then dl$ = Form2.Compute
With Form2.testpad
.enabled = True
.ResetSelColors
''
.nowrap = False
''
End With
Do
dl2 = dl$
ok = False
stackshowonly = True
If FastSymbol(dl$, ")") Then
ok = True
ElseIf IsExp(b, dl$, p) Then
    If AL$ = vbNullString Then
        If pagio$ = "GREEK" Then
        AL$ = "? " & Left$(dl2$, Len(dl2$) - Len(dl$)) & "=" & MyCStr(p)
        Else
        AL$ = "? " & Left$(dl2$, Len(dl2$) - Len(dl$)) & "=" & MyCStr(p)
        End If
            
    Else
        AL$ = AL$ & "," & Left$(dl2$, Len(dl2$) - Len(dl$)) & "=" & MyCStr(p)
    End If
    ok = True
    ElseIf IsStrExp(b, dl$, s$) Then
    If Len(dl2$) - Len(dl$) >= 0 Then
    
    
    If AL$ = vbNullString Then
        AL$ = Left$(dl2$, Len(dl2$) - Len(dl$)) & "=" & Chr(34) + s$ & Chr(34)
    Else
        AL$ = AL$ + "," + Left$(dl2$, Len(dl2$) - Len(dl$)) & "=" & Chr(34) + s$ & Chr(34)
    End If
    ok = True
    End If
    ElseIf InStr(dl$, ",") > 0 Then
       If InStr(dl$, Chr(2)) > 0 Then
     r$ = GetStrUntil(Chr(2), dl$, False)
     s$ = "<"
If ISSTRINGA(dl$, r$) Then If pagio$ <> "GREEK" Then s$ = s$ & r$
If ISSTRINGA(dl$, r$) Then If pagio$ = "GREEK" Then s$ = s$ & r$
AL$ = s$ & ">" & AL$
ok = True
Else
AL$ = AL$ & " " & GetStrUntil(",", dl$)
    
     dl$ = vbNullString
  
End If
    
    ok = True
    ElseIf dl$ <> "" Then
      If InStr(dl$, Chr(2)) > 0 Then
     r$ = GetStrUntil(Chr(2), dl$, False)
     s$ = "<"
If ISSTRINGA(dl$, r$) Then If pagio$ <> "GREEK" Then s$ = s$ & r$
If ISSTRINGA(dl$, r$) Then If pagio$ = "GREEK" Then s$ = s$ & r$
AL$ = s$ & ">" & AL$
ok = True
Else
     AL$ = AL$ & " " & dl$
     dl$ = vbNullString
  
End If

    End If
    
DropLeft ",", dl$

Loop Until Not ok
End If
stackshowonly = False
If AL$ <> "" Then AL$ = AL$ & vbCrLf
    If pagio$ = "GREEK" Then
    AL$ = AL$ & "Σωρός "
    Else
    AL$ = AL$ & "Stack "
    End If
If stack.Total = 0 Then
    If pagio$ = "GREEK" Then
    AL$ = AL$ & "Αδειος"
    Else
    AL$ = AL$ & "Empty"
    End If
Else
    If pagio$ = "GREEK" Then
    AL$ = AL$ & "Κορυφή "
    Else
    AL$ = AL$ & "Top "
    End If

End If
Dim i As Long

Do
i = i + 1
If stack.Total < i Or Len(AL$) > 400 Then Exit Do

If stack.StackItemType(i) = "N" Or stack.StackItemType(i) = "L" Then
AL$ = AL$ & MyCStr(stack.StackItem(i)) & " "
ElseIf stack.StackItemType(i) = "S" Then
r$ = stack.StackItem(i)
    If Len(r$) > 78 Then
    AL$ = AL$ & Chr(34) + Left$(r$, 75) & "..." & Chr(34)
    Else
    AL$ = AL$ & Chr(34) + r$ & Chr(34)
    End If
 ElseIf stack.StackItemType(i) = ">" Then
   If pagio$ = "LATIN" Then
    AL$ = AL$ & "[Optional] "
    Else
    AL$ = AL$ & "[Προαιρετικό] "
    End If
ElseIf stack.StackItemType(i) = "*" Then

AL$ = AL$ & stack.StackItemTypeObjectType(i) & " "
Else  '??
AL$ = AL$ & stack.StackItemTypeObjectType(i) & " "
End If

Loop
With Form2
    .gList1.backcolor = &H3B3B3B
        .label1(2) = .label1(2)
    
        .testpad.enabled = True
        .testpad.Text = AL$
        .testpad.SetRowColumn 1, 1
        .testpad.enabled = False
End With
once = False
End Sub

Sub makegroup(bstack As basetask, what$, i As Long)
Dim it As Long
it = globalvar(what$, it)
    MakeitObject2 var(it)
    If var(i).IamApointer Then
        If var(i).link.IamFloatGroup Then
           Set var(it).LinkRef = var(i).link
            var(it).IamApointer = True
            var(it).isref = True
        Else
            With var(i).link
            
                var(it).edittag = .edittag
                var(it).FuncList = .FuncList
                var(it).GroupName = myUcase(what$) + "."
                Set var(it).Sorosref = .soros.Copy
                var(it).HasValue = .HasValue
                var(it).HasSet = .HasSet
                var(it).HasStrValue = .HasStrValue
                var(it).HasParameters = .HasParameters
                var(it).HasParametersSet = .HasParametersSet
            
                        Set var(it).Events = .Events
            
                var(it).highpriorityoper = .highpriorityoper
                var(it).HasUnary = .HasUnary
            End With
        End If
    
    Else
        With var(i)
            var(it).edittag = .edittag
            var(it).FuncList = .FuncList
            var(it).GroupName = myUcase(what$) + "."
            Set var(it).Sorosref = .soros.Copy
            var(it).HasValue = .HasValue
            var(it).HasSet = .HasSet
            var(it).HasStrValue = .HasStrValue
            var(it).HasParameters = .HasParameters
            var(it).HasParametersSet = .HasParametersSet
            Set var(it).Events = .Events
            var(it).highpriorityoper = .highpriorityoper
            var(it).HasUnary = .HasUnary
        End With
        var(it).IamRef = Len(bstack.UseGroupname) > 0
    End If
    If var(i).HasStrValue Then
        globalvar what$ + "$", it, True
    End If
            
        
End Sub
Function ExecCode(basestack As basetask, rest$) As Boolean ' experimental
' ver .001
Dim p As Variant, mm As MemBlock, w2 As Long
    If IsExp(basestack, rest$, p) Then
        If Not basestack.lastobj Is Nothing Then
          If Not TypeOf basestack.lastobj Is mHandler Then
            Set basestack.lastobj = Nothing
            Exit Function
            End If
            With basestack.lastobj
                  If Not TypeOf .objref Is MemBlock Then
                      Set basestack.lastobj = Nothing
                      Exit Function
                  ElseIf .objref.NoRun Then
                       Set basestack.lastobj = Nothing
                       Exit Function
                  End If
            End With
            Set mm = basestack.lastobj.objref
            If mm.Status = 0 Then
            w2 = mm.GetPtr(0)
            If FastSymbol(rest$, ",") Then
            If Not IsExp(basestack, rest$, p) Then
                Set basestack.lastobj = Nothing
                Set mm = Nothing
                MissPar
                Exit Function
            End If
            If p < 0 Or p >= mm.SizeByte Then
                Set basestack.lastobj = Nothing
                Set mm = Nothing
                MyEr "Offset out of buffer", "Διεύθυνση εκτός διάρθρωσης"
                Exit Function
            End If

            SetUpForExecution w2, mm.SizeByte
            w2 = cUlng(uintnew(w2) + p)
            End If
            Set basestack.lastobj = Nothing
            Dim what As Long
            what = CallWindowProc(w2, 0&, 0&, 0&, 0&)
            If what <> 0 Then MyEr "Error " & what, "Λάθος " & what
            ReleaseExecution w2, mm.SizeByte
            ExecCode = what = 0
            Set mm = Nothing
            End If
            End If
        
    End If
    Set basestack.lastobj = Nothing
End Function
Sub MyMode(Scr As Object)
Dim x1 As Long, y1 As Long
On Error Resume Next
With players(GetCode(Scr))
    x1 = Scr.Width
    y1 = Scr.Height
    If Left$(Typename(Scr), 3) = "Gui" Then
    Else
    If Scr.name = "Form1" Then
    DisableTargets q(), -1
    
    ElseIf Scr.name = "DIS" Then
    DisableTargets q(), 0
    
    ElseIf Scr.name = "dSprite" Then
    DisableTargets q(), val(Scr.index)
    End If
    End If
    If .SZ < 4 Then .SZ = 4
        Err.clear
        Scr.Font.Size = .SZ
        If Err.Number > 0 Then
                MYFONT = "ARIAL"
                Scr.Font.name = MYFONT
                Scr.Font.charset = .charset
                Scr.Font.name = MYFONT
                Scr.Font.charset = .charset
        End If
        .uMineLineSpace = .MineLineSpace
        FrameText Scr, .SZ, x1, y1, .Paper
    .currow = 0
    .curpos = 0
    .XGRAPH = 0
    .YGRAPH = 0
End With
End Sub
Function ProcSave(basestack As basetask, rest$, Lang As Long) As Boolean
Dim pa$, w$, s$, col As Long, prg$, x1 As Long, par As Boolean, i As Long, noUse As Long, lcl As Boolean
On Error Resume Next
If lckfrm <> 0 Then MyEr "Save is locked", "Η αποθήκευση είναι κλειδωμένη": rest$ = vbNullString: Exit Function
lcl = IsLabelSymbolNew(rest$, "ΤΟΠΙΚΑ", "LOCAL", Lang) Or basestack.IamChild Or basestack.IamAnEvent
x1 = Abs(IsLabelFileName(basestack, rest, pa$, , s$))

If x1 <> 1 Then
rest$ = pa$ + rest$: x1 = IsStrExp(basestack, rest$, pa$)
Else
pa$ = s$: s$ = vbNullString
End If

If x1 <> 0 Then
        If subHash.count = 0 Or pa$ = vbNullString Then MyEr "Nothing to save", "Δεν υπάρχει κάτι να σώσω":              Exit Function
        If ExtractType(pa$) = "gsb" Then pa$ = ExtractPath(pa$) + ExtractNameOnly(pa$)
        If ExtractPath(pa$) <> "" Then
                If InStr(ExtractPath(pa$), mcd) <> 1 Then pa$ = pa$ & ".gsb" Else pa$ = pa$ & ".gsb"
        Else
                pa$ = mcd + pa$ & ".gsb"
        End If
        If Not WeCanWrite(pa$) Then Exit Function
        
      
           For i = subHash.count - 1 To 0 Step -1
       subHash.ReadVar i, s$, col
                If Right$(s$, 2) = "()" Then
                If Not InStr(s$, ChrW(&H1FFF)) > 0 Then
                s$ = Left$(s$, Len(s$) - 2)
                
                If Right$(sbf(col).sb, 2) <> vbCrLf Then sbf(col).sb = sbf(col).sb + vbCrLf
                If Lang Then
                
                        If Not blockCheck(sbf(col).sb, DialogLang, noUse, "Function " & s$ + "()" + vbCrLf) Then Exit Function
                                prg$ = s$ & " {" & sbf(col).sb & "}" & vbCrLf + prg$
                                If lcl Then
                                    prg$ = "FUNCTION " + prg$
                                Else
                                    prg$ = "FUNCTION GLOBAL " + prg$
                                End If
                        Else
                                If Not blockCheck(sbf(col).sb, DialogLang, noUse, "Συνάρτηση " & s$ + "()" + vbCrLf) Then Exit Function
                                prg$ = s$ & " {" & sbf(col).sb & "}" & vbCrLf + prg$
                                If lcl Then
                                    prg$ = "ΣΥΝΑΡΤΗΣΗ " + prg$
                                Else
                                    prg$ = "ΣΥΝΑΡΤΗΣΗ ΓΕΝΙΚΗ " + prg$
                                End If
                        End If
                End If
                Else
                        If Right$(sbf(col).sb, 2) <> vbCrLf Then sbf(col).sb = sbf(col).sb + vbCrLf
                        If Lang Then
                                If Not blockCheck(sbf(col).sb, DialogLang, noUse, "Module " & s$ + vbCrLf) Then Exit Function
                                prg$ = s$ & " {" & sbf(col).sb & "}" & vbCrLf + prg$
                                If lcl Then
                                    prg$ = "MODULE " + prg$
                                Else
                                    prg$ = "MODULE GLOBAL " + prg$
                                End If
                        Else
                                If Not blockCheck(sbf(col).sb, DialogLang, noUse, "Τμήμα " & s$ + vbCrLf) Then Exit Function
                                prg$ = s$ & " {" & sbf(col).sb & "}" & vbCrLf + prg$
                                If lcl Then
                                    prg$ = "ΤΜΗΜΑ " + prg$
                                Else
                                    prg$ = "ΤΜΗΜΑ ΓΕΝΙΚΟ " + prg$
                                End If
                        End If
                End If
        Next i
        w$ = vbNullString
        If FastSymbol(rest$, "@@", , 2) Then
            ' default password  - one space only - coder use default internal password
                If Not IsStrExp(basestack, rest$, w$) Then w$ = " "
        ElseIf FastSymbol(rest$, "@") Then
                ' One space only
                w$ = " "
        End If
        par = False
        If FastSymbol(rest$, ",") Then
                If Abs(IsLabel(basestack, rest$, s$)) = 1 Then
                        prg$ = prg$ & s$
                ElseIf FastSymbol(rest$, "{") Then
                        prg$ = prg$ & block(rest$)
                        If Not FastSymbol(rest$, "}") Then Exit Function
                End If
        End If
        ' reuse s$, col$
        If Len(w$) > 1 Then  'scrable col by George
                s$ = vbNullString: For col = 1 To Int((33 * Rnd) + 1): s$ = s$ & Chr(65 + Int((23 * Rnd) + 1)): Next col
                ' insert a variable length label......to make a variable length file
                prg$ = s$ & ":" & vbCrLf + prg$
                prg$ = mycoder.encryptline(prg$, w$, Len(prg$) Mod 33)
                par = True
        ElseIf Len(w$) = 1 Then   ' I have to check that...
                s$ = vbNullString:   For col = 1 To Int((33 * Rnd) + 1): s$ = s$ & Chr(65 + Int((23 * Rnd) + 1)): Next col
                prg$ = s$ & ":" & vbCrLf + prg$
                prg$ = mycoder.must1(prg$)
                par = True
        End If
        s$ = vbNullString
        If CFname(pa$) <> "" Then
                If Lang = 1 Then
                        If MsgBoxN("Replace " + ExtractNameOnly(pa$), vbOKCancel, MesTitle$) <> vbOK Then
                        MyEr "File not saved -1005", "Δεν σώθηκε το αρχείο -1005"
                        ProcSave = True
                        Exit Function
                        End If
                Else
                        If MsgBoxN("Αλλαγή " + ExtractNameOnly(pa$), vbOKCancel, MesTitle$) <> vbOK Then
                        MyEr "File not saved -1005", "Δεν σώθηκε το αρχείο -1005"
                        ProcSave = True
                        Exit Function
                        End If
                End If
                s$ = "*"
        End If
        If Not WeCanWrite(pa$) Then Exit Function
        If par Then
                If s$ = "*" Then
                       '' If CFname(ExtractPath(pa$) & ExtractNameOnly(pa$) & ".bck") <> "" Then killfile GetDosPath(ExtractPath(pa$) & ExtractNameOnly(pa$) & ".bck"): Sleep 30
                        MakeACopy pa$, ExtractPath(pa$) & ExtractNameOnly(pa$) & ".bck"
                End If
                If Not SaveUnicode(pa$, prg$, 0) Then BadFilename
                Else
                If s$ <> "" Then
                        ''If CFname(ExtractPath(pa$) & ExtractNameOnly(pa$) & ".bck") <> "" Then killfile GetDosPath(ExtractPath(pa$) & ExtractNameOnly(pa$) & ".bck"):  Sleep 30
                        MakeACopy pa$, ExtractPath(pa$) & ExtractNameOnly(pa$) & ".bck"
                End If
                ProcSave = SaveUnicode(pa$, prg$, 2)  ' 2 = utf-8 standard save mode for version 7
                If here$ = vbNullString Then LASTPROG$ = pa$
        End If
 ProcSave = True
Else
MyEr "A name please or use Ctrl+A to perform SAVE COMMAND$  (the last loading)", "Ένα όνομα παρακαλώ, ή πάτα το ctrl+Α για να αποθηκεύσεις με το όνομα του προγράμματος που φορτώθηκε τελευταία"
End If

End Function



Function Infinity() As Double
PutMem1 VarPtr(Infinity) + 7, &H7F
PutMem1 VarPtr(Infinity) + 6, &HF0
End Function
