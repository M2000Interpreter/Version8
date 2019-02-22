VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   9765
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TextP0.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6345
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   0
      MouseIcon       =   "TextP0.frx":0582
      Picture         =   "TextP0.frx":06D4
      ScaleHeight     =   405
      ScaleWidth      =   390
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox dSprite 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   675
      Index           =   0
      Left            =   6465
      ScaleHeight     =   675
      ScaleWidth      =   780
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1170
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox PrinterDocument1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   8010
      MouseIcon       =   "TextP0.frx":081E
      ScaleHeight     =   1140
      ScaleWidth      =   1185
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4965
      Visible         =   0   'False
      Width           =   1185
   End
   Begin M2000.gList gList1 
      Height          =   1575
      Left            =   7485
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2955
      Visible         =   0   'False
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   2778
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin M2000.gList List1 
      Height          =   1920
      Left            =   7755
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   510
      Visible         =   0   'False
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   3387
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser view1 
      Height          =   6000
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   8000
      ExtentX         =   14111
      ExtentY         =   10583
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.PictureBox DIS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   5640
      Left            =   780
      MouseIcon       =   "TextP0.frx":0B28
      MousePointer    =   1  'Arrow
      ScaleHeight     =   5640
      ScaleWidth      =   5640
      TabIndex        =   5
      Top             =   300
      Visible         =   0   'False
      Width           =   5640
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private onetime As Long
Public fState As Long
Public lockme As Boolean
Public TrueVisible As Boolean
Public WithEvents TEXT1 As TextViewer
Attribute TEXT1.VB_VarHelpID = -1
Public EditTextWord As Boolean
' by default EditTextWord is false, so we look for identifiers not words
Private Pad$, s$
Private LastDocTitle$, para1 As Long, PosPara1 As Long, Para2 As Long, PosPara2 As Long, Para3 As Long, PosPara3 As Long
Public ShadowMarks As Boolean
Private nochange As Boolean
Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal psString As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Public MY_BACK As New cDIBSection
Private mynum$
Dim OneOnly As Boolean
Public WithEvents HTML As HTMLDocument
Attribute HTML.VB_VarHelpID = -1
''Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
     ByVal nIndex As Integer, ByVal dwNewLong As Long) As Long
''Private Const GWL_STYLE = (-16)
Private DisStack As basetask
Private MeStack As basetask
Dim lookfirst As Boolean, look1 As Boolean
Private Declare Function GetLocaleInfo Lib "KERNEL32" Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function GetKeyboardLayout& Lib "user32" (ByVal dwLayout&) ' not NT?
Private Const DWL_ANYTHREAD& = 0
Const LOCALE_ILANGUAGE = 1
Private Declare Function PeekMessageW Lib "user32" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Const WM_KEYFIRST = &H100
 Const WM_KEYLAST = &H108
 Private Type POINTAPI
    X As Long
    Y As Long
End Type
 Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    Time As Long
    pt As POINTAPI
End Type
Public Point2Me As Object
Public TabControl As Long
Private Declare Function GetCommandLineW Lib "KERNEL32" () As Long

Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As Long, ByVal Value As Long)
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal Ptr As Long, ByVal Length As Long) As Long
Private Declare Function GetModuleHandleW Lib "KERNEL32" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "KERNEL32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long) As Long
    Private Const GWL_WNDPROC = -4
    Private m_Caption As String
Public lastitem As Long
Public Property Get CaptionW() As String
    If m_Caption = "M2000" Then
        CaptionW = vbNullString
    Else
        CaptionW = m_Caption
    End If
End Property


Public Property Let CaptionW(ByVal NewValue As String)
    Static WndProc As Long, VBWndProc As Long
    'If NewValue = vbnullstring Then NewValue = "M2000"
    m_Caption = NewValue

    If WndProc = 0 Then
        WndProc = GetProcAddress(GetModuleHandleW(StrPtr("user32")), "DefWindowProcW")
        VBWndProc = GetWindowLongA(hWnd, GWL_WNDPROC)
    End If


    If WndProc <> 0 Then
        SetWindowLongW hWnd, GWL_WNDPROC, WndProc
        SetWindowTextW hWnd, StrPtr(m_Caption)
        SetWindowLongA hWnd, GWL_WNDPROC, VBWndProc
    Else
        Caption = m_Caption
       
    End If
End Property
Public Function commandW() As String
Static mm$
If mm$ <> "" Then commandW = mm$: Exit Function
If m_bInIDE Then
mm$ = command
Else
Dim Ptr As Long: Ptr = GetCommandLineW
    If Ptr Then
        PutMem4 VarPtr(commandW), SysAllocStringLen(Ptr, lstrlenW(Ptr))
     If AscW(commandW) = 34 Then
       commandW = Mid$(commandW, InStr(commandW, """ ") + 2)
       Else
            commandW = Mid$(commandW, InStr(commandW, " ") + 1)
        End If
    End If
    End If
    If mm$ = vbNullString And command <> "" Then commandW = command Else commandW = mm$
End Function


Public Function GetLastKeyPressed() As Long
Dim Message As Msg
    If mynum$ <> "" Then
        GetLastKeyPressed = -1
    ElseIf PeekMessageW(Message, 0, WM_KEYFIRST, WM_KEYLAST, 0) Then
        GetLastKeyPressed = Message.wParam
    Else
        GetLastKeyPressed = -1
    
    End If
    Exit Function
End Function

Private Sub DIS_LostFocus()
If iamactive Then
iamactive = False
DestroyCaret
End If
End Sub

Private Sub DIS_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, shift As Integer, X As Single, Y As Single, state As Integer)
On Error Resume Next
If Not TaskMaster Is Nothing Then
  If TaskMaster.QueueCount > 0 Then
              TaskMaster.RestEnd1
   TaskMaster.TimerTick
End If
        End If
End Sub

Private Sub dSprite_GotFocus(index As Integer)
If lockme Then TEXT1.SetFocus: Exit Sub

End Sub

Private Sub dSprite_LostFocus(index As Integer)
If iamactive Then
iamactive = False
DestroyCaret
End If
End Sub

Private Sub dSprite_OLEDragOver(index As Integer, data As DataObject, Effect As Long, Button As Integer, shift As Integer, X As Single, Y As Single, state As Integer)
On Error Resume Next
If Not TaskMaster Is Nothing Then
  If TaskMaster.QueueCount > 0 Then
              TaskMaster.RestEnd1
   TaskMaster.TimerTick
End If
        End If
End Sub

Private Sub Form_Activate()
If ttl Then TrueVisible = True
If ASKINUSE Then
'Me.ZOrder 1
Else
If QRY Or GFQRY Then
If IsWine Then If DIS.Visible Then DIS.SetFocus
End If
End If
If Form1.Visible Then releasemouse = True: If lockme Then hookme TEXT1.glistN
    If Typename(ActiveControl) = "gList" Then
                
                Hook hWnd, ActiveControl
                
                End If
End Sub
Public Sub UNhookMe()
Set LastGlist = Nothing
UnHook hWnd
End Sub
Private Sub Form_Deactivate()
UNhookMe
End Sub

Private Sub Form_GotFocus()
UseEsc = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, shift As Integer)
Dim i As Long
 If List1.LeaveonChoose Then Exit Sub
clickMe = -1
i = -1
If KeyCode = vbKeyV Then
Exit Sub
End If
If shift <> 4 And mynum$ <> "" Then
On Error Resume Next
If Left$(mynum$, 1) = "0" Then
i = val(mynum$)
Else
i = val(mynum$)
End If
mynum$ = vbNullString
Else
i = GetLastKeyPressed
End If

 If i <> -1 And i <> 94 Then
UKEY$ = ChrW(i)
 Else
 If i <> -1 Then UKEY$ = vbNullString
 End If

End Sub

Private Sub Form_LostFocus()
If iamactive Then
iamactive = False
DestroyCaret
End If

UseEsc = False
End Sub

Private Sub Form_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, shift As Integer, X As Single, Y As Single, state As Integer)
On Error Resume Next
If Not TaskMaster Is Nothing Then
  If TaskMaster.QueueCount > 0 Then
              TaskMaster.RestEnd1
   TaskMaster.TimerTick
End If
        End If
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 0 Then WindowState = 0: Exit Sub
End Sub

Private Sub gList1_ChangeListItem(item As Long, content As String)
Dim i As Long

If nochange Then
nochange = True
With TEXT1
i = .SelLength
.Form1mn1Enabled = i > 0
.Form1mn2Enabled = i > 0
.Form1mn3Enabled = Clipboard.GetFormat(13) Or Clipboard.GetFormat(1)
.Form1sdnEnabled = i > 0 And (.Length - .SelStart) > i
.Form1supEnabled = i > 0 And .SelStart > i
.Form1mscatEnabled = .Form1sdnEnabled Or .Form1supEnabled
.Form1rthisEnabled = .Form1mscatEnabled
End With
nochange = False
End If
End Sub

Private Sub gList1_ChangeSelStart(thisselstart As Long)
Dim i As Long

If gList1.enabled Then
With TEXT1
i = .SelLength
.Form1mn1Enabled = i > 0
.Form1mn2Enabled = i > 0
.Form1mn3Enabled = Clipboard.GetFormat(13) Or Clipboard.GetFormat(1)
.Form1sdnEnabled = i > 0 And (.Length - .SelStart) > i
.Form1supEnabled = i > 0 And .SelStart > i
.Form1mscatEnabled = .Form1sdnEnabled Or .Form1supEnabled
.Form1rthisEnabled = .Form1mscatEnabled
End With
End If
End Sub


Private Sub gList1_GetBackPicture(pic As Object)
Set pic = Point2Me
End Sub

Private Sub gList1_HeaderSelected(Button As Integer)
Dim i As Long

If Not gList1.enabled Then Exit Sub
With TEXT1
If .UsedAsTextBox Then Exit Sub
i = .SelLength
.Form1mn1Enabled = i > 0
.Form1mn2Enabled = i > 0
.Form1mn3Enabled = Clipboard.GetFormat(13) Or Clipboard.GetFormat(1)
.Form1sdnEnabled = i > 0 And (.Length - .SelStart) > .SelLength
.Form1supEnabled = i > 0 And .SelStart > .SelLength
.Form1mscatEnabled = .Form1sdnEnabled Or .Form1supEnabled
.Form1rthisEnabled = .Form1mscatEnabled
End With
UNhookMe
MyPopUp.feedlabels TEXT1, EditTextWord
MyPopUp.Up

End Sub

Private Sub gList1_KeyDownAfter(KeyCode As Integer, shift As Integer)
If KeyCode = vbKeyTab Then
If shift = 2 Then
choosenext
KeyCode = 0
End If
End If
End Sub

Private Sub glist1_MarkOut()
Pack1
End Sub
Public Sub Pack1()
Dim i As Long
With TEXT1
i = .SelLength
.Form1mn1Enabled = i > 0
.Form1mn2Enabled = i > 0
.Form1mn3Enabled = Clipboard.GetFormat(13) Or Clipboard.GetFormat(1)
.Form1sdnEnabled = i > 0 And (.Length - .SelStart) > i
.Form1supEnabled = i > 0 And .SelStart > i
.Form1mscatEnabled = .Form1sdnEnabled Or .Form1supEnabled
.Form1rthisEnabled = .Form1mscatEnabled
End With
End Sub

Private Sub gList1_SyncKeyboard111(KeyAscii As Integer)
If KeyAscii = 9 Then KeyAscii = 0: Exit Sub
If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub gList1_OutPopUp(X As Single, Y As Single, myButton As Integer)
Dim i As Long

If Not gList1.enabled Then Exit Sub
With TEXT1
If .UsedAsTextBox Then Exit Sub
i = .SelLength
.Form1mn1Enabled = i > 0
.Form1mn2Enabled = i > 0
.Form1mn3Enabled = Clipboard.GetFormat(13) Or Clipboard.GetFormat(1)
.Form1sdnEnabled = i > 0 And (.Length - .SelStart) > .SelLength
.Form1supEnabled = i > 0 And .SelStart > .SelLength
.Form1mscatEnabled = .Form1sdnEnabled Or .Form1supEnabled
.Form1rthisEnabled = .Form1mscatEnabled
End With
UNhookMe
MyPopUp.feedlabels TEXT1, EditTextWord
MyPopUp.Up X + gList1.Left, Y + gList1.top
myButton = 0
End Sub

Public Sub helpmeSub()
If Not EditTextWord Then
If Trim$(TEXT1.SelText) <> "" Then
ffhelp myUcase(Trim(TEXT1.SelText), True)
Else
vHelp
End If
Else
If abt Then
feedback$ = Trim$(TEXT1.SelText)
feednow$ = FeedbackExec$
CallGlobal feednow$
Else
vHelp
End If
End If
End Sub

Private Sub ffhelp(a$)
If Left$(a$, 1) < "Α" Then
fHelp basestack1, a$, True
Else
fHelp basestack1, a$
End If
End Sub



Private Sub List1_ListError(code As Long)
Dim dummy As Long
List1.ListIndex = -1
List1.LeaveonChoose = False
List1.Visible = False
If List1.Tag <> "" Then
If QRY Or GFQRY Then
Else

dummy = interpret(basestack1, List1.Tag)
'Me.KeyPreview = True
End If
End If
MyEr "Menu Error " & CStr(code), "Λάθος στην ΕΠΙΛΟΓΗ αριθμός " & CStr(code)
End Sub


Private Sub glist1_RegisterGlist(this As gList)
On Error Resume Next
hookme this
If Err.Number > 0 Then this.NoWheel = True
End Sub

Private Sub gList1_UnregisterGlist()
On Error Resume Next
Set LastGlist = Nothing
If Err.Number > 0 Then gList1.NoWheel = True
End Sub

Private Function HTML_oncontextmenu() As Boolean
HTML_oncontextmenu = False
End Function



Private Sub HTML_onkeydown()
Select Case view1.Document.parentWindow.event.KeyCode
Case vbKeyF1
IEUP homepage$
Form1.KeyPreview = False
Case vbKeyEscape
If escok Then
IEUP ""
While KeyPressed(&H1B)
MyDoEvents
Refresh
Wend
    If QRY Or GFQRY Then KeyPreview = True
INK$ = vbNullString
''UINK$ = VbNullString
End If

End Select
End Sub

Private Sub List1_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal X As Long, ByVal Y As Long)
If item = -1 Then

Else
List1.mousepointer = 1
If lastitem = item Then Exit Sub
If List1.ListSep(item) Then Exit Sub
List1.ListindexPrivateUse = item
List1.ShowMe2
lastitem = item
'List1.ListindexPrivateUse = -1
End If
End Sub

Private Sub list1_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
If item = List1.ListIndex Then

List1.FillThere thisHDC, thisrect, &HFFFFFF, -List1.LeftMarginPixels  ' or black in reverse
List1.WriteThere thisrect, List1.list(item), List1.PanPos / dv15, List1.addpixels / 2, 0
skip = True
Else
skip = False
End If
End Sub

Private Sub List1_PanLeftRight(Direction As Boolean)
Dim dummy As Boolean
If List1.Tag <> "" Then
If QRY Or GFQRY Then
Else

dummy = interpret(basestack1, List1.Tag)
'Me.KeyPreview = True
End If
Else
List1.LeaveonChoose = False
List1.Visible = False

End If
End Sub

Private Sub List1_Selected2(item As Long)
Dim dummy As Boolean

If List1.Tag <> "" Then
If QRY Or GFQRY Then
Else

dummy = interpret(basestack1, List1.Tag)
'Me.KeyPreview = True
End If
Else
List1.LeaveonChoose = False
List1.Visible = False
End If
End Sub



Public Sub mn5sub()
 'If Not EditTextWord Then
 ' check if { } is ok...
 'If Not blockCheck(TEXT1.Text, DialogLang) Then Exit Sub
 'End If

CancelEDIT = True
MyDoEvents
NOEDIT = True
End Sub

Public Sub mscatsub()
''
Dim l As Long, w As Long, s$, TempLcid As Long, OldLcid As Long
Dim el As Long, eW As Long, safety As Long, TT$

w = TEXT1.mDoc.MarkParagraphID
eW = w
TEXT1.SelStartSilent = TEXT1.SelStart  'MOVE CHARPOS TO SELSTART

el = TEXT1.Charpos  ' charpos maybe is in the start or the end of block
s$ = TEXT1.SelText
OldLcid = TEXT1.mDoc.lcid
TempLcid = FoundLocaleId(s$)
If TempLcid <> 0 Then TEXT1.mDoc.lcid = TempLcid

l = el + 1
If EditTextWord Then
Do
If TEXT1.mDoc.FindWord(s$, True, w, l) Then
TT$ = TEXT1.mDoc.TextParagraph(w)
Mid$(TT$, l, Len(s$)) = s$
TEXT1.mDoc.ReWritePara w, TT$
TEXT1.mDoc.WrapAgainBlock w, w
TEXT1.mDoc.ColorThis (w)
Else
w = 1
l = 0
safety = safety + 1
End If
Loop Until (w = eW And l = el) Or safety = 2

Else
Do
If TEXT1.mDoc.FindIdentifier(s$, True, w, l) Then
TT$ = TEXT1.mDoc.TextParagraph(w)
Mid$(TT$, l, Len(s$)) = s$
TEXT1.mDoc.TextParagraph(w) = TT$
TEXT1.mDoc.WrapAgainBlock w, w
TEXT1.mDoc.ColorThis (w)
Else
w = 1
l = 0
safety = safety + 1
End If
Loop Until (w = eW And l = el) Or safety = 2

End If
TEXT1.mDoc.lcid = OldLcid
TEXT1.mDoc.WrapAgain

TEXT1.Render
End Sub

Public Sub rthissub(Optional anystr As Boolean = False)
If TEXT1.mDoc.busy Then Exit Sub
Dim l As Long, w As Long, s$, TempLcid As Long, OldLcid As Long, noinp As Double
Dim el As Long, eW As Long, safety As Long, TT$, w1 As Long, i1 As Long
Dim neo$, mDoc10 As Document, addthat As Long, w2 As Long
w = TEXT1.mDoc.MarkParagraphID
eW = w
TEXT1.SelStartSilent = TEXT1.SelStart  'MOVE CHARPOS TO SELSTART
el = TEXT1.Charpos  ' charpos maybe is in the start or the end of block
s$ = TEXT1.SelText
TEXT1.SelStartSilent = TEXT1.SelStart
el = TEXT1.Charpos  ' charpos maybe is in the start or the end of block

If pagio$ = "GREEK" Then
neo$ = InputBoxN("Αλλαγή Λέξης (Shift για σταμάτημα)", "Συγγραφή Κειμένου", s$, noinp)
Else
neo$ = InputBoxN("Replace Word (use Shift for Stop)", "Text Editor", s$, noinp)
End If
If noinp <> 1 Then Exit Sub
OldLcid = TEXT1.mDoc.lcid
TempLcid = FoundLocaleId(s$)
If TempLcid <> 0 Then TEXT1.mDoc.lcid = TempLcid
If Len(neo$) >= Len(s$) Then
    Set mDoc10 = New Document
    mDoc10 = neo$
    w1 = 0
    i1 = 0
    
    If EditTextWord Or anystr Then
        If anystr Then
        If mDoc10.FindStrDown(s$, w1, i1) Then addthat = i1 - 1: If Len(neo$) = Len(s$) And addthat = 0 Then Exit Sub
        Else
        If mDoc10.FindWord(s$, True, w1, i1) Then addthat = i1 - 1: If Len(neo$) = Len(s$) And addthat = 0 Then Exit Sub
        End If
    Else
        If mDoc10.FindIdentifier(s$, True, w1, i1) Then addthat = i1 - 1: If Len(neo$) = Len(s$) And addthat = 0 Then Exit Sub
    End If
    
End If

i1 = el
l = i1 + addthat
w1 = w
If EditTextWord Or anystr Then
TEXT1.glistN.dropkey = True
Dim ok1 As Boolean
Do
If anystr Then
ok1 = TEXT1.mDoc.FindStrDown(s$, w, l)
Else
ok1 = TEXT1.mDoc.FindWord(s$, True, w, l)
End If
If ok1 Then
If safety And w = w1 Then
If w2 > 0 Then If w2 <> w Then TEXT1.mDoc.WrapAgainBlock w2, w2:  TEXT1.mDoc.ColorThis w2
w2 = w
If l = i1 Then
 TEXT1.SelLengthSilent = 0
TEXT1.mDoc.MarkParagraphID = w
TEXT1.glistN.enabled = False
TEXT1.ParaSelStart = l
TEXT1.glistN.enabled = True
TEXT1.SelLength = Len(s$)
TEXT1.AddUndo ""
TEXT1.SelText = neo$
TEXT1.RemoveUndo neo$
Exit Do
ElseIf l - addthat < i1 Then
i1 = i1 + Len(neo$) - Len(s$)
Else

End If
End If
TEXT1.SelLengthSilent = 0
TEXT1.mDoc.MarkParagraphID = w
TEXT1.glistN.enabled = False
TEXT1.ParaSelStart = l
TEXT1.glistN.enabled = True
TEXT1.SelLength = Len(s$)
TEXT1.AddUndo ""
TEXT1.SelText = neo$
TEXT1.RemoveUndo neo$
l = l + Len(neo$)

Else
w = 1
l = 0
safety = safety + 1
End If
Loop Until safety = 2 Or KeyPressed(16)
TEXT1.glistN.dropkey = False

Else
''If l > 0 Then l = l - 1
TEXT1.glistN.dropkey = True
Do
If TEXT1.mDoc.FindIdentifier(s$, True, w, l) Then
'If w2 > 0 Then TEXT1.mDoc.ColorThis w2: TEXT1.WrapMarkedPara
If w2 > 0 Then If w2 <> w Then TEXT1.mDoc.WrapAgainBlock w2, w2:    TEXT1.mDoc.ColorThis w2
w2 = w
If safety And w = w1 Then

If l = i1 Then
 TEXT1.SelLengthSilent = 0
TEXT1.mDoc.MarkParagraphID = w
TEXT1.glistN.enabled = False
TEXT1.ParaSelStart = l
TEXT1.glistN.enabled = True
TEXT1.SelLength = Len(s$)
TEXT1.AddUndo ""
TEXT1.SelText = neo$
TEXT1.RemoveUndo neo$

Exit Do
ElseIf l - addthat < i1 Then
i1 = i1 + Len(neo$) - Len(s$)
Else

End If
End If
TEXT1.SelLengthSilent = 0
TEXT1.mDoc.MarkParagraphID = w
TEXT1.glistN.enabled = False
TEXT1.ParaSelStart = l
TEXT1.glistN.enabled = True
TEXT1.SelLength = Len(s$)
TEXT1.AddUndo ""

TEXT1.SelText = neo$
TEXT1.RemoveUndo neo
 
l = l + Len(neo$)

Else
w = 1
l = 0
safety = safety + 1
End If
Loop Until safety = 2 Or KeyPressed(16)
TEXT1.glistN.dropkey = False
End If
TEXT1.mDoc.lcid = OldLcid
If w2 > 0 Then TEXT1.mDoc.WrapAgainBlock w2, w2:  TEXT1.mDoc.ColorThis w2
TEXT1.Render

End Sub

Public Sub sdnSub()
Dim b$
b$ = s$
s$ = TEXT1.SelText
If s$ = vbNullString Or InStr(s$, Chr$(13)) > 0 Or InStr(s$, Chr$(10)) > 0 Then s$ = b$
SearchDown s$
End Sub
Sub SearchDown(s$, Optional anystr As Boolean = False)
Dim l As Long, w As Long, TempLcid As Long, OldLcid As Long
w = TEXT1.mDoc.MarkParagraphID   ' this is the not the order
TEXT1.SelStartSilent = TEXT1.SelStart
l = TEXT1.Charpos + 1

OldLcid = TEXT1.mDoc.lcid
TempLcid = FoundLocaleId(s$)
If TempLcid <> 0 Then TEXT1.mDoc.lcid = TempLcid
If EditTextWord Or anystr Then
    If anystr Then
  If Not TEXT1.mDoc.FindStrDown(s$, w, l) Then GoTo sdnOut
  Else
    If Not TEXT1.mDoc.FindWord(s$, True, w, l) Then GoTo sdnOut
    End If
Else
    If Not TEXT1.mDoc.FindIdentifier(s$, True, w, l) Then GoTo sdnOut
End If
TEXT1.SelLengthSilent = 0
TEXT1.mDoc.MarkParagraphID = w
TEXT1.glistN.enabled = False
TEXT1.ParaSelStart = l
TEXT1.glistN.enabled = True
TEXT1.SelLength = Len(s$)
sdnOut:
TEXT1.mDoc.lcid = OldLcid
End Sub

Public Sub supsub()
Dim b$
b$ = s$
s$ = TEXT1.SelText
If s$ = vbNullString Or InStr(s$, Chr$(13)) > 0 Or InStr(s$, Chr$(10)) > 0 Then s$ = b$
Searchup s$
End Sub
Sub Searchup(s$, Optional anystr As Boolean = False)
Dim l As Long, w As Long, TempLcid As Long, OldLcid As Long
w = TEXT1.mDoc.MarkParagraphID
TEXT1.SelStartSilent = TEXT1.SelStart - (TEXT1.SelLength > 1)
l = TEXT1.Charpos + 1
OldLcid = TEXT1.mDoc.lcid
TempLcid = FoundLocaleId(s$)
If TempLcid <> 0 Then TEXT1.mDoc.lcid = TempLcid
If EditTextWord Or anystr Then
   If anystr Then
   If Not TEXT1.mDoc.FindStrUp(s$, w, l) Then GoTo sdupOut
   Else
       If Not TEXT1.mDoc.FindWord(s$, False, w, l) Then GoTo sdupOut
    End If
Else
    If Not TEXT1.mDoc.FindIdentifier(s$, False, w, l) Then GoTo sdupOut
End If
TEXT1.SelLengthSilent = 0
TEXT1.mDoc.MarkParagraphID = w
TEXT1.glistN.enabled = False
TEXT1.ParaSelStart = l
TEXT1.glistN.enabled = True
TEXT1.SelLength = Len(s$)
sdupOut:
TEXT1.mDoc.lcid = OldLcid
End Sub
Public Function InIDECheck() As Boolean
    m_bInIDE = True
    InIDECheck = True
End Function


Private Sub DIS_GotFocus()
If lockme Then TEXT1.SetFocus: Exit Sub
clickMe2 = -1
End Sub

Private Sub DIS_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = vbKeyPause Then
Form_KeyDown KeyCode, shift
End If
End Sub
Public Sub GiveASoftBreak(Sorry As Boolean)
clickMe2 = -1
' Try first with escape
If Sorry Then
Form_KeyDown vbKeyPause, (0)
Else  'CTRL C
Form_KeyDown &HFFFE, (0)
End If

End Sub

Private Sub DIS_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
If Not NoAction Then
NoAction = True

If Button > 0 And Targets Then

If Button = 1 Then
Dim sel&
    sel& = ScanTarget(q(), CLng(X), CLng(Y), 0)
    If sel& >= 0 Then
        Select Case q(sel&).Id Mod 100
        Case Is < 10
        If TaskMaster Is Nothing Then
        
        If Not interpret(DisStack, (q(sel&).Comm)) Then Beep
        If LastErNum1 > 0 Then MyEr "", ""
        Else
        TaskMaster.StopProcess
        If Not interpret(DisStack, (q(sel&).Comm)) Then Beep
        If LastErNum1 > 0 Then MyEr "", ""
        TaskMaster.StartProcess
        End If
        Case Else
        INK$ = q(sel&).Comm
        End Select
    End If
End If

If Not nomore Then NoAction = False

End If




End If

End Sub




Private Sub dSprite_MouseDown(index As Integer, Button As Integer, shift As Integer, X As Single, Y As Single)
Dim p As Long, u2 As Long
If lockme Then Exit Sub
If Not NoAction Then
NoAction = True
Dim sel&
p = val("0" & dSprite(index).Tag)
With players(p)
    u2 = .uMineLineSpace * 2

        If Button > 0 And Targets Then

        sel& = ScanTarget(q(), CLng(X), CLng(Y), index)
            If sel& >= 0 Then
                If Button = 1 Then
                Select Case q(sel&).Id Mod 100
                Case Is < 10
                If Not interpret(DisStack, "LAYER " & dSprite(index).Tag + " {" + vbCrLf + q(sel&).Comm + vbCrLf & "}") Then Beep
                Case Else
                INK$ = q(sel&).Comm
                End Select


End If
End If


If Not nomore Then NoAction = False

End If
End With
End If

End Sub


Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
Dim i As Long
Form1.Font.charset = GetCharSet(GetCodePage(GetLCIDFromKeyboard))

Static ctrl As Boolean, noentrance As Boolean
If KeyCode = 13 And List1.Visible And (Not List1.LeaveonChoose) And Not QRY Then
KeyCode = 0
List1.PressSoft
Exit Sub
End If
If KeyCode = 13 And trace Then
If Not STq Then STbyST = True: KeyCode = 0

End If
clickMe = HighLow(CLng(shift), CLng(KeyCode))
If clickMe2 = -2 Then clickMe2 = clickMe
If clickMe = 27 And escok Then
NOEXECUTION = True
If Not TaskMaster Is Nothing Then
If TaskMaster.Processing Then
TaskMaster.StopProcess

End If
TaskMaster.Dispose
End If
If exWnd <> 0 Then
MyDoEvents
    nnn$ = "bye bye"
    exWnd = 0

    End If
If view1.Visible Then
MyDoEvents

view1.Navigate "about:blank"
Sleep 50
view1.Visible = False
If QRY Or GFQRY Then KeyPreview = True Else Form1.KeyPreview = False
 End If

End If

If clickMe2 <> -1 Then KeyCode = 0: Exit Sub

If BLOCKkey Then Exit Sub
If noentrance Then
KeyCode = 0
Exit Sub
End If
If shift = 4 Then
If KeyCode = 18 Then
If mynum$ = vbNullString Then mynum$ = "0"
KeyCode = 0
Exit Sub
End If
Select Case KeyCode
Case vbKeyAdd
mynum$ = "&h"
Case vbKey0 To vbKey9
mynum$ = mynum$ + Chr$(KeyCode - vbKey0 + 48)
Case vbKeyNumpad0 To vbKeyNumpad9
mynum$ = mynum$ + Chr$(KeyCode - vbKeyNumpad0 + 48)
Case vbKeyA To vbKeyF
If Left$(mynum$, 1) = "&" Then
mynum$ = mynum$ + Chr$(KeyCode - vbKeyNumpad0 + 65)
Else
mynum$ = vbNullString
End If
Case Else
mynum$ = vbNullString
End Select

Exit Sub
End If

mynum$ = vbNullString


Select Case KeyCode
Case vbKeyE, vbKeyD
If ctrl And (shift And &H2) = 2 Then
If QRY Then
If pagio$ = "GREEK" Then
INK$ = INK$ & "ΣΥΓΓΡΑΦΗ "
Else
INK$ = INK$ & "EDIT "
End If
End If
End If
Case vbKeyA
If ctrl And (shift And &H2) = 2 Then
If QRY Then
If LASTPROG$ <> "" Then
If pagio$ = "GREEK" Then
INK$ = "ΣΩΣΕ ΕΝΤΟΛΗ$" & vbCr
Else
INK$ = "SAVE COMMAND$" & vbCr
End If
End If
End If
End If
Case vbKeyS
If ctrl And (shift And &H2) = 2 Then
If QRY Then
If pagio$ = "GREEK" Then
INK$ = "ΣΩΣΕ "
Else
INK$ = "SAVE "
End If
End If
End If
Case vbKeyL
If ctrl And (shift And &H2) = 2 Then
If QRY Then
If pagio$ = "GREEK" Then
INK$ = "ΛΙΣΤΑ "
Else
INK$ = "LOAD "
End If
End If
End If
Case vbKeyF
If ctrl And (shift And &H2) = 2 Then
If QRY Then
If pagio$ = "GREEK" Then
INK$ = "ΦΟΡΤΩΣΕ "
Else
INK$ = "FILES "
End If
End If
End If
Case vbKeyP, vbKeyT
If ctrl And (shift And &H2) = 2 Then
If QRY Then
If pagio$ = "GREEK" Then
INK$ = "ΤΥΠΩΣΕ "
Else
INK$ = "PRINT "
End If
End If
End If
Case vbKeyM
If ctrl And (shift And &H2) = 2 Then
 If QRY Then
 If pagio$ = "GREEK" Then
INK$ = "ΤΜΗΜΑΤΑ "
Else
INK$ = "MODULES "
End If
 End If
End If
Case vbKeyU
If ctrl And (shift And &H2) = 2 Then
 If QRY Then
 If pagio$ = "GREEK" Then
INK$ = "ΡΥΘΜΙΣΕΙΣ " + vbCr
Else
INK$ = "SETTINGS " + vbCr
End If
 End If
End If
Case vbKeyN
If ctrl And (shift And &H2) = 2 Then
 If QRY Then
 If pagio$ = "GREEK" Then
INK$ = "ΤΜΗΜΑΤΑ ? " + vbCr
Else
INK$ = "MODULES ? " + vbCr
End If
 End If
End If
Case vbKeyTab
    If (shift And 1) = 1 Then
    INK$ = INK$ & Chr$(6)
    KeyCode = 0
    ElseIf ctrl Or shift = 2 Or Not (QRY Or GFQRY) Then
    ctrl = False
        choosenext
        KeyCode = 0
        
    End If
Case vbKeyV
    If ctrl And (shift And &H2) = 2 Then
        Pad$ = GetTextData(CF_UNICODETEXT)
        If Pad$ <> "" Then
          '  For i = 1 To Len(pad$)
              '  If Asc(Mid$(pad$, i, 1)) > 31 Then
                INK$ = Pad$
              '  Exit Sub
             '   End If
           ' Next i
        End If
         KeyCode = 0
        Exit Sub
    End If

Case vbKeyC, &HFFFE
If (ctrl And (shift And &H2) = 2) Or KeyCode = &HFFFE Then
If QRY Then
INK$ = INK$ & "CLS" & Chr$(13)
Else
KeyCode = 0
If Form4.Visible Then
Form4.Visible = False
    If TEXT1.Visible Then
        TEXT1.SetFocus
    Else
        Form1.SetFocus
    End If
End If
'EXECSTOP
End If
End If
Case vbKeyPause  '(this is the break key!!!!!'
If Forms.count > 5 Then KeyCode = 0: Exit Sub
If Not TaskMaster Is Nothing Then If TaskMaster.QueueCount > 0 Then KeyCode = 0: Exit Sub
If QRY Or GFQRY Then
If Form4.Visible Then Form4.Visible = False
i = MOUT
If ASKINUSE Then
If BreakMe Then Exit Sub
Unload NeoMsgBox: ASKINUSE = False: Exit Sub
End If
BreakMe = True
If MsgBoxN("Break Key - Hard Reset" + vbCrLf + "Μ2000 - Execution Stop / Τερματισμός Εκτέλεσης", vbYesNo, MesTitle$) <> vbNo Then

MOUT = i

extreme = False
If AVIRUN Then AVI.GETLOST
On Error Resume Next
noentrance = True
NoAction = True
If Not TaskMaster Is Nothing Then TaskMaster.Dispose: MyEr "", ""

If Me.Visible Then Me.SetFocus
closeAll ' we closed all files
QRY = False
GFQRY = False
escok = True
INK$ = Chr$(27) + Chr$(27)
If MOUT = False Then
NOEXECUTION = True
MOUT = True
Else
MOUT = False
End If
If List1.Visible Then
List1.Tag = vbNullString
List1.Visible = False
List1.LeaveonChoose = False
INK$ = vbNullString
End If
noentrance = False

End If
BreakMe = False
End If
If IsWine Then releasemouse = True
KeyCode = 0
Case vbKeyLeft
INK$ = INK$ & Chr(0) + Chr(75)   ' GWBASIC Codes
Case vbKeyRight
INK$ = INK$ & Chr(0) + Chr(77)
Case vbKeyUp
INK$ = INK$ & Chr(0) + Chr(72)
Case vbKeyDown
INK$ = INK$ & Chr(0) + Chr(80)
Case vbKeyInsert
INK$ = INK$ & Chr(0) + Chr(82)
Case vbKeyDelete
INK$ = INK$ & Chr(0) + Chr(83)
Case vbKeyPageUp
INK$ = INK$ & Chr(0) + Chr(73)
Case vbKeyPageDown
INK$ = INK$ & Chr(0) + Chr(81)
Case vbKeyHome
INK$ = INK$ & Chr(0) + Chr(71)
Case vbKeyEnd
INK$ = INK$ & Chr(0) + Chr(79)
Case vbKeyEscape
If List1.LeaveonChoose Then Exit Sub
INK$ = INK$ & Chr(27)
If escok Then
If AVIRUN Then
AVI.GETLOST
End If
NOEXECUTION = True
End If
Case vbKeyF1 To vbKeyF12
If FKey >= 0 Then FKey = KeyCode - vbKeyF1 + 1
If Abs(FKey) = 1 And ctrl And (shift And &H2) = 2 Then
If lastAboutHTitle <> "" Then abt = True: vH_title$ = vbNullString

FKey = 0: KeyCode = 0: vHelp
ElseIf FKey = 1 And (shift And 1) Then
FKey = 13
ElseIf FKey = 4 And ctrl And QRY Then
interpret DisStack, "END"
End If

Case vbKeyControl

ctrl = True
KeyCode = 0
Exit Sub
Case Else
If ctrl And (shift And &H2) = 2 And lckfrm = 0 And KeyCode <> 3 And KeyCode <> 16 Then
If escok Then
STq = False
STEXIT = False
STbyST = True
Form2.Show , Form1
 PrepareLabel basestack1
 
Form2.label1(1) = "..."
Form2.label1(2) = "..."
    Form2.gList3(2).backcolor = &H3B3B3B
    TestShowCode = False
     TestShowSub = vbNullString
 TestShowStart = 0
     Set Form2.Process = basestack1
   stackshow basestack1
Form1.Show , Form5
trace = True
End If
End If
End Select

ctrl = False
 If List1.LeaveonChoose Then Exit Sub
 If KeyCode = 91 Then Exit Sub
i = GetLastKeyPressed
 If i <> -1 And i <> 94 Then UKEY$ = ChrW(i) Else If i <> -1 Then UKEY$ = vbNullString
 If List1.Visible Then
 Else
KeyCode = 0
End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 And view1.Visible Then view1.SetFocus: KeyAscii = 0: Exit Sub
If clickMe2 = -2 And clickMe <> -1 Then clickMe2 = clickMe
If clickMe2 <> -1 And Not List1.LeaveonChoose Then KeyAscii = 0: Exit Sub

If Right$(INK$, 1) = Chr$(6) And KeyAscii = 9 Then

Else
If mynum$ <> "" Then
If mynum$ <> "" Then Exit Sub
End If
If UKEY$ <> "" Then
INK$ = INK$ & UKEY$
UKEY$ = vbNullString
Else
If KeyAscii = 22 Then
KeyAscii = 0
Else
INK$ = INK$ & GetKeY(KeyAscii)
End If
End If
End If
End Sub
Friend Sub EXECSTOP()
Dim iamhere As Boolean
If iamhere Then Exit Sub
iamhere = True
NOEXECUTION = False
If MsgBoxN("[Ctrl + C] Μ2000 - Execution Stop / Τερματισμός Εκτέλεσης", vbYesNo, MesTitle$) = vbYes Then
extreme = False
If AVIRUN Then
AVI.GETLOST
End If
 On Error Resume Next
'noentrance = True
If TaskMaster.QueueCount > 0 Then TaskMaster.Dispose
NoAction = True
Close ' we closed all files
escok = True
If QRY Then
INK$ = Chr$(27) + Chr$(27)
MyDoEvents
End If
QRY = False
RRCOUNTER = 0
REFRESHRATE = 25
ResetPrefresh
INK$ = Chr$(27) + Chr$(27)
NOEXECUTION = True
'trace = True
If lckfrm > 0 Then
If MOUT = False Then
NOEXECUTION = True
MOUT = True
Else
MOUT = False
End If
End If

'noentrance = False
End If
EmptyClipboard

iamhere = False
End Sub


Private Sub Form_Load()
TabControl = 6
Set DisStack = New basetask
Set MeStack = New basetask
Debug.Assert (InIDECheck = True)
onetime = 0
Set fonttest = Form1.dSprite(0)
Set TEXT1 = New TextViewer

Set TEXT1.Container = gList1
With TEXT1.glistN
.DragEnabled = False ' only drop - we can change this from popup menu
.enabled = False
TEXT1.FileName = vbNullString
.addpixels = 0
TEXT1.showparagraph = False

TEXT1.EditDoc = True
TEXT1.TabWidth = 4

.LeftMarginPixels = 10
.WordCharLeft = ConCat(":", "{", "}", "[", "]", ",", "(", ")", "!", ";", "=", ">", "<", "'", """", " ", "+", "-", "/", "*", "^", "@", Chr$(9))
.WordCharRight = ConCat(":", "{", "}", "[", "]", ",", ")", "!", ";", "=", ">", "<", "'", """", " ", "+", "-", "/", "*", "^", Chr$(9))
.WordCharRightButIncluded = "("

End With
List1.NoWheel = True
List1.FreeMouse = True
List1.LeftMarginPixels = 4
List1.NoPanRight = False
List1.SingleLineSlide = True
Dim s$
Set DisStack.Owner = DIS

List1.BypassLeaveonChoose = False

Set MeStack.Owner = Me


ThereIsAPrinter = IsPrinter
If ThereIsAPrinter Then

pname = Printer.DeviceName
port = Printer.port

End If

dset
'If LoadFont(GetCurDir(True) & "TT6492M_.TTF") Then
'defFontname = "monospace 821 greek bt"
'MYFONT = defFontname
'myBold = True
'Else
MYFONT = "Verdana"
defFontname = MYFONT
myBold = False
'End If
myCharSet = 0
With Form1
.Font.name = MYFONT
.Font.Strikethrough = False
.Font.Underline = False
.Font.bold = myBold
MYFONT = .Font.name
    .Font.charset = myCharSet
    .DIS.Font.charset = myCharSet
    .DIS.Font.name = MYFONT
    .DIS.Font.bold = myBold
    .TEXT1.Font.charset = myCharSet
    .TEXT1.Font.name = MYFONT
    .TEXT1.Font.bold = myBold
    
    .List1.charset = myCharSet
    .List1.Font.name = MYFONT
    .List1.FontBold = myBold
     
End With


s$ = commandW
If Not ISSTRINGA(s$, cLine) Then
cLine = mylcasefILE(Trim(s$))
Else
cLine = mylcasefILE(cLine)
End If
While Left$(cLine, 1) = Chr(34) And Right$(cLine, 1) = Chr(34) And Len(cLine) > 2
cLine = Mid$(cLine, 2, Len(cLine) - 2)
Wend
If ExtractType(cLine) <> "gsb" Then cLine = vbNullString
If cLine <> "" Then
para$ = ExtractPath(cLine) + ExtractName(cLine)
cLine = Trim$(Mid$(cLine, Len(para$) + 1))
s$ = cLine + " " + s$
cLine = para$
ElseIf s$ <> "" Then
para$ = Trim$(s$)
End If



Switches para$  ' ,TRUE CHECK THIS
    
    l_complete = True
  
 
111:

  On Error Resume Next
  Dim i As Long
  
      For i = 0 To Controls.count - 1
     If Typename(Controls(i)) <> "Menu" Then Controls(i).TabStop = False
      Next i
End Sub




Private Sub Form_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
If NoAction Then Exit Sub
NoAction = True
Dim sel&

If Button > 0 And Targets Then
sel& = ScanTarget(q(), CLng(X), CLng(Y), -1)

If sel& >= 0 Then

If Button = 1 Then


Select Case q(sel&).Id Mod 100
Case Is < 10

If Not interpret(MeStack, (q(sel&).Comm)) Then Beep
Case Else
INK$ = q(sel&).Comm
End Select


Else

End If

End If
If Not nomore Then NoAction = False

End If
If lockme Then Exit Sub
'MOUB = Button
clickMe2 = -1



End Sub

'Private Sub Form_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
'If lockme Then Exit Sub
'If Button > 0 Then MOUB = Button

'If NOEDIT = True And (exWnd = 0 Or Button) Then
'Me.KeyPreview = True
'End If
'End Sub

'Private Sub Form_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
'If lockme Then Exit Sub
' MOUB = 0

'End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If byPassCallback Then Exit Sub
Cancel = NoAction

End Sub

Private Sub iForm_Resize()
DIS.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
Public Sub Up()
UpdateWindow hWnd
End Sub

Sub something()
Set basestack1.Owner = DIS
Set DisStack.Owner = DIS
NOEXECUTION = False
basestack1.toprinter = False
MOUT = False
On Error Resume Next
Const HWND_BROADCAST = &HFFFF&
Const WM_FONTCHANGE = &H1D
Dim pn As Long, a As New cDIBSection
AutoRedraw = True
 If App.StartMode = vbSModeStandalone Then If OneOnly Then Exit Sub
OneOnly = True

 stacksize = 900000
If m_bInIDE Then funcdeep = 128 Else funcdeep = 3260
KeyPreview = True
If Not ttl Then Load Form3
ttl = True
Form3.Visible = False
escok = False
Sleep 10
If pagio$ = "GREEK" Then
FK$(13) = "ΣΥΓΓΡΑΦΕΑΣ"
Else
FK$(13) = "WRITER"
End If
Console = FindFormSScreen(Me)

Me.WindowState = 0
If Me.WindowState = 0 Then
With ScrInfo(Console)
Me.Move .Left, ScrY(), .Width, .Height
End With
End If

basestack1.myCharSet = 0
    Font.charset = basestack1.myCharSet
    basestack1.Owner.Font.charset = basestack1.myCharSet
    TEXT1.Font.charset = basestack1.myCharSet
    List1.Font.charset = basestack1.myCharSet

basestack1.Owner.Move 0, 0, ScaleWidth, ScaleHeight

If NoAction Then Exit Sub
Dim dummy As Boolean, i
NOEXECUTION = False
'
myBreak basestack1
MyNew basestack1, "", 0
MyClear basestack1, ""

Dim mybasket As basket
mybasket = players(DisForm)
PlaceBasket DIS, mybasket

If cLine <> "" Then LASTPROG$ = cLine

s_complete = True
players(DisForm) = mybasket

End Sub
Public Sub MyPrompt(LoadFileAndSwitches$, Prompt$)
Static forwine As Long

onetime = onetime + 1
On Error GoTo finale
elevatestatus = -1
s_complete = True
ExTarget = False
Dim helpcnt As Long, qq$
If Not UseMe Is Nothing Then
If Not UseMe.IamVisible Then
UseMe.Show
End If
End If
''MyDoEvents

Dim mybasket As basket
mybasket = players(DisForm)
Do
Do
escok = True
MOUT = True

MyDoEvents

 If cLine = vbNullString Then

   If trace Then
   PrepareLabel basestack1
    Form2.label1(1) = "..."
    Form2.label1(2) = "..."
    Form2.gList3(2).backcolor = &H3B3B3B
    TestShowCode = False
    TestShowSub = vbNullString
    TestShowStart = 0
    Set Form2.Process = basestack1
    stackshow basestack1
    Form2.ComputeNow

   
    End If
   
    If Not Form1.Visible Then
    Form1.Show , Form5: releasemouse = False
    If ttl Then
    If Form3.Visible Then
    Form3.skiptimer = True
    Form3.WindowState = 0
    End If
    End If
    End If
    If Not releasemouse Then
    If Screen.ActiveForm Is Nothing Then
    Else
    ' Form1.SetFocus
    releasemouse = True
    End If
    End If
  
    NORUN1 = False

 players(DisForm) = mybasket
 ''reset refresh system
  REFRESHRATE = 25
  ResetPrefresh
  k1 = 0
 Show
 
If onetime = 1 Then
    If IsWine Then
    If forwine = 0 Then
    forwine = 1
    newStart basestack1, ""
    MOUT = True
    ElseIf Form1.Visible = False Then
    Else
    Form1.SetFocus
    End If
    ElseIf Form1.Visible = False Then
    Else
    onetime = 2
    Form1.SetFocus
    End If
    
    
    
End If
breakit:
If NOEXECUTION Then
If MKEY$ = "@Start" + Chr$(13) Then
qq$ = "@start"
MKEY$ = vbNullString
NOEXECUTION = False
MOUT = True
GoTo conthere
End If
NOEXECUTION = False
MOUT = True
End If

    QUERY basestack1, Prompt$, qq$, (mybasket.mx * 4), True
    If NOEXECUTION And MOUT Then
    If MKEY$ = "@Start" + Chr$(13) Then qq$ = "@start"
        NOEXECUTION = False
    

    End If
conthere:
    
    If ExTarget = True Then Exit Sub
      mybasket = players(DisForm)

      If basestack1.Owner Is Nothing Then Exit Sub
If basestack1.Owner.Visible = True Then basestack1.Owner.Refresh Else basestack1.Owner.Visible = True
    FKey = 0
    If pagio$ = "GREEK" Then
    FK$(13) = "ΣΥΓΓΡΑΦΕΑΣ"
    Else
    FK$(13) = "WRITER"
    End If
    INK$ = vbNullString
    mybasket.pageframe = 0
    MYSCRnum2stop = holdcontrol(DIS, mybasket)
    HoldReset 1, mybasket
If LoadFileAndSwitches$ = vbNullString And qq$ = vbNullString Then helpcnt = helpcnt + 1

        If helpcnt > 4 Then
    If basestack1.Owner.Font.charset <> 161 Then
    qq$ = " HELP": helpcnt = -100000
    Else
    qq$ = " ΒΟΗΘΕΙΑ": helpcnt = -100000
    End If
    End If
 crNew basestack1, mybasket
If NOEXECUTION Then GoTo breakit
 
Else
   sHelp "", "", 0, 0
   qq$ = "LOAD" & """" + cLine + """"
   If Len(Left$(cLine, rinstr(cLine, "\"))) > 0 Then

    mcd = Left$(cLine, rinstr(cLine, "\"))
   End If
    PlaceCaption ExtractNameOnly(cLine)
   cLine = vbNullString
End If

If Not MOUT Then
NOEXECUTION = False
ResetBreak
MOUT = interpret(basestack1, "START")
qq$ = vbNullString

MOUT = interpret(basestack1, "cls")

mybasket = players(DisForm)
End If
NOEXECUTION = False
Loop Until qq$ <> ""

NoAction = True

ResetBreak
players(DisForm) = mybasket
NoAction = True
NOEXECUTION = False
basestack1.toprinter = False
MOUT = False
ClearLabels
If Not interpret(basestack1, qq$) Then
mybasket = players(DisForm)
'' ClearLoadedForms  not needed any more
If NERR Then Exit Do
    basestack1.toprinter = False
    If MOUT Then
            NOEXECUTION = False
            ResetBreak
'
            
            mybasket = players(DisForm)
            MOUT = False
        Else
        
        If NOEXECUTION Then

                closeAll
                mybasket = players(DisForm)
                If byPassCallback Then Exit Do
               If Not MKEY$ = "@Start" + Chr$(13) Then PlainBaSket DIS, mybasket, "ESC " & qq$
        Else
        ' look last error
                If Left$(LastErName & " ", 1) <> "?" Then
                        closeAll
                        crNew basestack1, mybasket
                        If basestack1.Owner.Font.charset <> 161 Then
                        
                        wwPlain basestack1, mybasket, " ? " & LastErName, basestack1.Owner.Width, 1000, True
                        If Left$(FK$(13), 4) = "EDIT" Then crNew basestack1, mybasket: wwPlain basestack1, mybasket, "Use SHIFT F1, edit, ESC to return", basestack1.Owner.Width, 1000, True
                        Else
                        wwPlain basestack1, mybasket, " ? " & LastErNameGR, basestack1.Owner.Width, 1000, True
                        If Left$(FK$(13), 4) = "EDIT" Then crNew basestack1, mybasket: wwPlain basestack1, mybasket, "Με το SHIFT F1 διορθώνεις, ESC επιστρέφεις", basestack1.Owner.Width, 1000, True
                        End If
                            
                            LastErName = "?" & LastErName
                            LastErNameGR = "?" & LastErNameGR
                Else
                        mybasket = players(DisForm)
                        wwPlain basestack1, mybasket, " ? " & qq$, basestack1.Owner.Width, 1000, True
                End If
        End If
        crNew basestack1, mybasket
        LastErNum = 0: LastErNum1 = 0
        LastErName = vbNullString
        LastErNameGR = vbNullString
        ExTarget = False
        End If
        players(DisForm) = mybasket
        End If
        mybasket = players(DisForm)
        
        LCTbasketCur DIS, mybasket
         If mybasket.curpos > 0 Then
          crNew basestack1, mybasket
        
        
          
         End If
 mybasket.curpos = 0
MOUT = True
NoAction = False
If ExTarget Then Exit Do
If ttl Then

If Form3.WindowState = 1 Then Form3.WindowState = 0
End If
para$ = vbNullString

  Loop
elevatestatus = 0
Exit Sub
finale:
ExTarget = True
elevatestatus = 0
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set DisStack.Owner = Nothing
     Set DisStack = Nothing
    Set MeStack.Owner = Nothing
    Set MeStack = Nothing
    TEXT1.Dereference
    Set Point2Me = Nothing
    Set fonttest = Nothing
    TrueVisible = False
    byPassCallback = True
End Sub
Public Sub helper1()
If DisStack Is Nothing Then
Else
    Set DisStack.Owner = Nothing
     Set DisStack = Nothing
End If
   If MeStack Is Nothing Then
Else
    Set MeStack.Owner = Nothing
    Set MeStack = Nothing
    End If
    TEXT1.Dereference
    Set Point2Me = Nothing
    Set fonttest = Nothing
End Sub
    
Private Sub List1_DblClick()
Dim dummy As Boolean
List1.Visible = False
If List1.Tag <> "" Then
If QRY Or GFQRY Then

Else

dummy = interpret(basestack1, List1.Tag)
'Me.KeyPreview = True
End If
End If
End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
Dim dummy As Boolean
If KeyAscii = 13 Then
List1.Visible = False
If List1.Tag <> "" Then
If QRY Or GFQRY Then
Else

dummy = interpret(basestack1, List1.Tag)
'Me.KeyPreview = True
End If
End If
End If
End Sub




Private Sub gList1_KeyDown(KeyCode As Integer, shift As Integer)
Static ctrl As Boolean, noentrance As Boolean, where As Long, noinp As Double

Dim aa$, a$, JJ As Long, ii As Long, gothere As Long
If KeyCode = vbKeyEscape Then
KeyCode = 0
 If Not EditTextWord Then
 ' check if { } is ok...
 If Not blockCheck(TEXT1.Text, DialogLang, gothere) Then
 On Error Resume Next
 gList1.ListindexPrivateUse = gothere
 Exit Sub
 End If
 End If
 If TEXT1.UsedAsTextBox Then Result = 99
NOEDIT = True: noentrance = False: Exit Sub
End If
If KeyCode = vbKeyPause Then
 KeyCode = 0: NOEDIT = True: noentrance = False
 If Form4.Visible Then Form4.Visible = False
            If Form1.Visible Then
             If TEXT1.Visible Then
                TEXT1.SetFocus
                Form1.SetFocus
            End If
            End If
            If Forms.count > 5 Then KeyCode = 0: Exit Sub
            If Not TaskMaster Is Nothing Then If TaskMaster.QueueCount > 0 Then KeyCode = 0: Exit Sub
            If BreakMe Then noentrance = False: Exit Sub
            If ASKINUSE Then
                
                
                Unload NeoMsgBox: ASKINUSE = False: noentrance = False: Exit Sub
                End If
            BreakMe = True
            If Form3.ask(basestack1, "Break Key - Hard Reset" & vbCrLf & "Μ2000 - Execution Stop / Τερματισμός Εκτέλεσης") = 1 Then
            
            If AVIRUN Then
            AVI.GETLOST
            End If


On Error Resume Next
noentrance = True
NoAction = True
Close ' we closed all files
QRY = False
escok = True
INK$ = Chr$(27) + Chr$(27)

If MOUT = False Then
NOEXECUTION = True
MOUT = True
Else
MOUT = False
End If
End If
BreakMe = False
noentrance = False
 Exit Sub
 
End If
'***************************************
'Exit Sub
If TEXT1.UsedAsTextBox Then
Select Case KeyCode
Case Is = vbKeyTab And (shift Mod 2 = 1), vbKeyUp
Result = -1
Case vbKeyReturn, vbKeyTab, vbKeyDown
Result = 1
Case Else
noentrance = False
Exit Sub
End Select
KeyCode = 0

NOEDIT = True: noentrance = False: Exit Sub

Exit Sub
End If

If noentrance Then
KeyCode = 0
noentrance = False
Exit Sub
End If
noentrance = True
With TEXT1
.Form1mn1Enabled = .SelLength > 1
.Form1mn2Enabled = .Form1mn1Enabled
.Form1mn3Enabled = Clipboard.GetFormat(13) Or Clipboard.GetFormat(1)
End With

If KeyCode = 13 And shift = 2 Then
KeyCode = 0
shift = 0
UKEY$ = vbNullString
TEXT1.insertbrackets
noentrance = False
Exit Sub
End If
Select Case KeyCode
Case vbKeyReturn
nochange = True



If TEXT1.AutoIntNewLine Then
    KeyCode = 0
    nochange = False
    Exit Sub
End If
nochange = False
Case vbKeyControl
ctrl = True
KeyCode = 0
Case vbKeyF1
If (shift And 2) = 2 Then
If TEXT1.SelText <> "" Then
helpmeSub
Else

vHelp
End If
Else
TEXT1.nowrap = Not TEXT1.nowrap
TEXT1.Render
TEXT1.ManualInform
End If

KeyCode = 0
Case vbKeyF2
If shift <> 0 Then
If s$ = vbNullString And TEXT1.SelText <> "" Then s$ = TEXT1.SelText
If pagio$ = "GREEK" Then
s$ = InputBoxN("Αναζήτησε προς τα πάνω:", "Συγγραφή Κειμένου", s$, noinp)
Else
s$ = InputBoxN("Search to top:", "Text Editor", s$, noinp)
End If
If MyTrim$(s$) <> "" And noinp = 1 Then Searchup s$, shift Mod 2 = 1
shift = 0
ElseIf TEXT1.SelText <> "" Or s$ <> "" Then
supsub
End If

KeyCode = 0
Case vbKeyF3
If shift <> 0 Then
If s$ = vbNullString And TEXT1.SelText <> "" Then s$ = TEXT1.SelText
If pagio$ = "GREEK" Then

s$ = InputBoxN("Αναζήτησε προς τα κάτω:", "Συγγραφή Κειμένου", s$, noinp)
Else
s$ = InputBoxN("Search to down:", "Text Editor", s$, noinp)
End If
If MyTrim$(s$) <> "" And noinp = 1 Then SearchDown s$, shift Mod 2 = 1
shift = 0
ElseIf TEXT1.SelText <> "" Or s$ <> "" Then

sdnSub
End If
KeyCode = 0
Case vbKeyF4
If TEXT1.SelText <> "" Then mscatsub
KeyCode = 0
Case vbKeyF5
If TEXT1.SelText <> "" Then rthissub shift Mod 2 = 1
KeyCode = 0
Case vbKeyF6  ' Set/Show/Reset Para1
MarkSoftButton para1, PosPara1
KeyCode = 0
Case vbKeyF7  'Set/Show/Reset Para2
MarkSoftButton Para2, PosPara2
KeyCode = 0
Case vbKeyF8  'Set/Show/Reset Para2
MarkSoftButton Para3, PosPara3
KeyCode = 0

Case vbKeyF9  ' Count Words/
If shift <> 0 Then
TEXT1.NoCenterLineEdit = Not TEXT1.NoCenterLineEdit
If UserCodePage = 1253 Then
If TEXT1.NoCenterLineEdit Then
TEXT1.ReplaceTitle = "Ελεύθερη διόρθωση στο κείμενο"
Else
TEXT1.ReplaceTitle = "Διόρθωση στη κεντρική γραμμή"
End If
Else
If TEXT1.NoCenterLineEdit Then
TEXT1.ReplaceTitle = "Free line edit mode"
Else
TEXT1.ReplaceTitle = "Center line edit mode"
End If
End If

Else
If TEXT1.glistN.lines > 1 Then
If UserCodePage = 1253 Then
TEXT1.ReplaceTitle = "Λέξεις στο κείμενο:" + CStr(TEXT1.mDoc.WordCount)
Else
TEXT1.ReplaceTitle = "Words in text:" + CStr(TEXT1.mDoc.WordCount)
End If
End If
End If
KeyCode = 0
Case vbKeyF10
If shift <> 0 Then
With TEXT1
If .UsedAsTextBox Then Exit Sub
ii = .SelLength
.Form1mn1Enabled = ii > 1
.Form1mn2Enabled = ii > 1
.Form1mn3Enabled = Clipboard.GetFormat(13) Or Clipboard.GetFormat(1)
.Form1sdnEnabled = ii > 0 And (.Length - .SelStart) > .SelLength
.Form1supEnabled = ii > 0 And .SelStart > .SelLength
.Form1mscatEnabled = .Form1sdnEnabled Or .Form1supEnabled
.Form1rthisEnabled = .Form1mscatEnabled
End With
UNhookMe
MyPopUp.feedlabels TEXT1, EditTextWord
MyPopUp.Up
Else
TEXT1.showparagraph = Not TEXT1.showparagraph
TEXT1.mDoc.WrapAgain
TEXT1.Render
End If
KeyCode = 0

Case vbKeyF11
fState = fState + 1
SetText1
TEXT1.WrapAll
TEXT1.ManualInform
KeyCode = 0
Case vbKeyF12
If shift <> 0 Then
mn5sub

Else
showmodules
End If
KeyCode = 0

Case vbKeyTab
nochange = True
If TEXT1.HaveMarkedText Then TEXT1.SelStartSilent = TEXT1.SelStart
    gList1.enabled = False
    JJ = TEXT1.SelStart
    where = JJ
    ii = 1 + TEXT1.SelStart - TEXT1.ParaSelStart
    
    If TEXT1.SelLength > 0 Then
    
    JJ = TEXT1.SelLength + JJ - ii
    TEXT1.SelStart = ii
    TEXT1.SelLength = JJ
    JJ = where
    Else
    TEXT1.SelStart = ii
    End If


    If TEXT1.SelText <> "" Then
        a$ = vbCrLf + TEXT1.SelText & "*"
        If gList1.UseTab Then
            If shift <> 0 Then
                If InStr(a$, vbCrLf + vbTab) = 0 Then
                a$ = Replace(a$, vbCrLf + String$(TabControl + (Len(TEXT1.CurrentParagraph) - Len(LTrim(TEXT1.CurrentParagraph))) Mod TabControl, ChrW(160)), vbCrLf)
                a$ = Replace(a$, vbCrLf + space$(TabControl + (Len(TEXT1.CurrentParagraph) - Len(LTrim(TEXT1.CurrentParagraph))) Mod TabControl), vbCrLf)
                Else
                a$ = Replace(a$, vbCrLf + vbTab, vbCrLf)
                
                End If
                TEXT1.InsertTextNoRender = Mid$(a$, 3, Len(a$) - 3)
                 TEXT1.SelStartSilent = ii
                 TEXT1.SelLengthSilent = Len(a$) - 3
                 
            Else
                If InStr(a$, vbCrLf + " ") > 0 Or InStr(a$, vbCrLf + ChrW(160)) > 0 Then
                a$ = Replace(a$, vbCrLf, vbCrLf + space$(TabControl))
                TEXT1.InsertTextNoRender = Mid$(a$, 3, Len(a$) - 3)
                TEXT1.SelStartSilent = where + TabControl
                TEXT1.SelLengthSilent = Len(a$) - 3 - (where + TabControl - ii)
                Else
                a$ = Replace(a$, vbCrLf, vbCrLf + vbTab)
                
                TEXT1.InsertTextNoRender = Mid$(a$, 3, Len(a$) - 3)
                TEXT1.SelStartSilent = where + 1
                TEXT1.SelLengthSilent = Len(a$) - 3 - (where + 1 - ii)
               End If
            End If
        Else
            If shift <> 0 Then
                a$ = Replace(a$, vbCrLf + String$(TabControl + (Len(TEXT1.CurrentParagraph) - Len(LTrim(TEXT1.CurrentParagraph))) Mod TabControl, ChrW(160)), vbCrLf)
                a$ = Replace(a$, vbCrLf + space$(TabControl + (Len(TEXT1.CurrentParagraph) - Len(LTrim(TEXT1.CurrentParagraph))) Mod TabControl), vbCrLf)
                TEXT1.InsertTextNoRender = Mid$(a$, 3, Len(a$) - 3)
                 TEXT1.SelStartSilent = ii
                 TEXT1.SelLengthSilent = Len(a$) - 3
                 
            Else
                a$ = Replace(a$, vbCrLf, vbCrLf + space$(TabControl))
                TEXT1.InsertTextNoRender = Mid$(a$, 3, Len(a$) - 3)
                TEXT1.SelStartSilent = where + TabControl
                TEXT1.SelLengthSilent = Len(a$) - 3 - (where + TabControl - ii)
               
            End If
        End If
    Else
        If shift And 1 <> 1 Then
    
            If Mid$(TEXT1.CurrentParagraph, 1, TabControl) = space$(TabControl) Or Mid$(TEXT1.CurrentParagraph, 1, TabControl) = String$(TabControl, ChrW(160)) Then
        
                    TEXT1.SelStartSilent = ii
                    TEXT1.SelLengthSilent = TabControl
                    TEXT1.InsertTextNoRender = vbNullString
                    TEXT1.SelStartSilent = ii
            ElseIf Left$(TEXT1.CurrentParagraph, 1) = vbTab Then
                    TEXT1.SelStartSilent = ii
                    TEXT1.SelLengthSilent = 1
                    TEXT1.InsertTextNoRender = vbNullString
                    TEXT1.SelStartSilent = ii
            Else
                    TEXT1.SelStartSilent = ii
                    TEXT1.SelLengthSilent = Len(TEXT1.CurrentParagraph) - Len(NLtrim(TEXT1.CurrentParagraph))
                    TEXT1.InsertTextNoRender = vbNullString
                    TEXT1.SelStartSilent = ii
            End If
        Else
        
            If gList1.UseTab And MyTrimL2(TEXT1.CurrentParagraph) < JJ - ii + 1 - (JJ = ii) Then
            If (shift And 1) = 1 And Left$(TEXT1.CurrentParagraph, 1) <> ChrW(9) Then
            TEXT1.SelStartSilent = JJ
            TEXT1.RemoveUndo space(TabControl)
            TEXT1.InsertText = space(TabControl)
            
            TEXT1.SelStartSilent = where + TabControl
            Else
            TEXT1.SelStartSilent = JJ
            TEXT1.RemoveUndo vbTab
            TEXT1.InsertText2 = vbTab
            TEXT1.SelStartSilent = JJ + 1
            End If
            
            Else
            TEXT1.SelStartSilent = JJ
            TEXT1.RemoveUndo space(TabControl)
            TEXT1.InsertText = space(TabControl)
            
            TEXT1.SelStartSilent = where + TabControl
        End If
    End If
End If
gList1.enabled = True
TEXT1.ReColorBlock
TEXT1.glistN.Noflashingcaret = False
TEXT1.Render

nochange = False
KeyCode = 0
shift = 0
'gList1_MarkOut
Case Else

ctrl = False
End Select
noentrance = False
End Sub

Private Sub List1_SyncKeyboardUnicode(a As String)
'refresh
MyDoEvents2
If QRY Or GFQRY Then
If a = Left$(INK$, 1) Then INK$ = Mid$(a, 2)
a = ""
Else
INK$ = INK$ & a
End If
End Sub

Private Sub TEXT1_CtrlPlusF1()
If lastAboutHTitle <> "" Then abt = True: vH_title$ = vbNullString
vHelp
End Sub

Private Sub TEXT1_Inform(tLine As Long, tPos As Long)
If TEXT1.UsedAsTextBox Then

Else
If UserCodePage = 1253 Then
textinformCaption = "Γραμμή(" + CStr(tLine) + ")-Θέση(" + CStr(TEXT1.Charpos) + ")"
TEXT1.ReplaceTitle = "[" + CStr(TEXT1.Charpos) + "-" + CStr(tLine) + "/" + CStr(TEXT1.mDoc.DocLines) + "]  §:" + CStr(TEXT1.mDoc.DocParagraphs) + Mark$ + " " + GetLCIDFromKeyboardLanguage

Else
textinformCaption = "Line(" + CStr(tLine) + ")-Pos(" + CStr(TEXT1.Charpos) + ")"
TEXT1.ReplaceTitle = "[" + CStr(TEXT1.Charpos) + "-" + CStr(tLine) + "/" + CStr(TEXT1.mDoc.DocLines) + "] §:" + CStr(TEXT1.mDoc.DocParagraphs) + Mark$ + " " + GetLCIDFromKeyboardLanguage
End If

End If
End Sub


Private Sub view1_BeforeNavigate2(ByVal pDisp As Object, Url As Variant, flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
If look1 Then
look1 = False:  lookfirst = False

Cancel = True  ' 2 times
End If

If lookfirst Then look1 = True: view1.Silent = True


End Sub


Private Sub view1_DocumentComplete(ByVal pDisp As Object, Url As Variant)
   Set HTML = view1.Document

End Sub

Private Sub view1_NavigateComplete2(ByVal pDisp As Object, Url As Variant)
'
On Error Resume Next
If look1 Then
''Set pDisp = view1.Object
view1.SetFocus

End If
End Sub

Private Sub view1_NewWindow2(ppDisp As Object, Cancel As Boolean)
Static prev$
Cancel = True
If HTML Is Nothing Then Exit Sub
If prev$ = HTML.activeElement.toString Then

Else
prev$ = HTML.activeElement.toString
view1.Navigate prev$
Sleep 50
End If
End Sub



Private Sub view1_TitleChange(ByVal Text As String)

If LCase$(Right$(Text, 5)) <> "done/" Then
If InStr(Text, "?") = 1 Then
nnn$ = TClear(Text): Sleep 5
'Beep
view1.Navigate "http://done/"
End If
Else

End If

End Sub
Function TClear(ByVal txt As String) As String
Dim Nb As String, ic As Long
txt = StrConv(txt, vbUnicode)
For ic = 1 To Len(txt) Step 2
Nb = Nb + Mid$(txt, ic, 1)
Next ic
TClear = Nb & "."
End Function
Private Sub view1_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub view1_LostFocus()
'Me.KeyPreview = True
End Sub




Public Sub view1_StatusTextChange11(bstack As basetask, ByVal t1 As String)
On Error Resume Next
exWnd = 0

view1.Visible = False
Sleep 1
PREPARE bstack, t1
Sleep 1
If Form1.Visible Then Form1.Refresh
End Sub
Private Sub PREPARE(basestack As basetask, ByVal Nb As String)
Dim g As Long, ic As Long
Dim VP As String, vv As String, CM$, b As String
If needset Then
Nb = StrConv(Nb, vbUnicode)
For ic = 1 To Len(Nb) Step 2
b = b + Mid$(Nb, ic, 1)
Next ic
Else
b = Nb
b = Replace(b, Chr(0), "")
End If
g = InStr(b, "?")
If g > 0 Then b = Mid$(b, g + 1) Else Exit Sub
If b <> "" Then
Do While Parameters(b, VP, vv)
CM$ = VP & "$=" & """" + vv + """"
Execute basestack, CM$, True
Loop
MyDoEvents
''Me.KeyPreview = True
End If
End Sub
Public Sub IEUP(ThisFile As String)
Static once As Boolean
If once Then Exit Sub
once = True
If ThisFile = vbNullString Then

If exWnd <> 0 Then
 Set HTML = Nothing
MyDoEvents
homepage$ = vbNullString
'view1.TabStop = False
    nnn$ = "bye bye"
    exWnd = 0
    
    End If
If view1.Visible Then
MyDoEvents

view1.Navigate "about:blank"
Sleep 50
view1.Visible = False

 End If
 once = False
    Exit Sub
End If
   


On Error Resume Next
'tf2$ = THISFILE
needset = False
Dim MSD As String
MSD = App.path
AddDirSep MSD
'View1.TabStop = True

If Form1.Visible = False Then Form1.Visible = True: Sleep 400
view1.Visible = True

lookfirst = True
look1 = False



If IESizeX = 0 Or IESizeY < 100 Then
IEX = 0
IEY = 0
IESizeX = Me.ScaleWidth
IESizeY = Me.ScaleHeight
End If
If (IESizeX - IEX) > Me.ScaleWidth Then IESizeX = Me.ScaleWidth - IEX
If (IESizeY - IEY) > Me.ScaleHeight Then IESizeY = Me.ScaleHeight - IEY
If IsWine Then
With view1
On Error Resume Next
    .Visible = True
    .top = IEY
    .Left = IEX
    .Width = IESizeX
    .Height = IESizeY
    .Refresh
    .Refresh2

End With
Else
view1.Move IEX, IEY, IESizeX, IESizeY
End If



view1.RegisterAsBrowser = True
   If homepage$ = vbNullString Then homepage$ = ThisFile$
   exWnd = 1
view1.Navigate ThisFile$

Do
view1.Visible = True
MyDoEvents2 Me
Sleep 5
Loop Until view1.Visible Or MOUT

If Not MOUT Then
view1.setfoucs
Me.KeyPreview = False
End If

'follow IEX, IEY
cnt = False
once = False
End Sub
Public Sub follow(ByVal nx As Long, ByVal ny As Long)
Exit Sub
IEX = nx
IEY = ny
If exWnd > 0 Then

End If

End Sub


Private Function Parameters(a As String, b As String, c As String) As Boolean
Dim i, ch As Boolean, vl As Boolean, chs$, all$, many As Long
b = vbNullString
c = vbNullString

'parameters = False
ch = False
vl = False
Do While i < Len(a)
i = i + 1
Select Case Mid$(a, i, 1)
Case "%"
If Mid$(a, i + 1, 1) = "u" Then
i = i + 1
'we have four bytes
many = 6
Else
many = 4
End If
chs$ = "&H"
ch = True
Case ";"
If Not vl Then
' throw it is &amp;
b = vbNullString
End If
Case "+"
If vl = True Then
c = c & " "
Else
b = b & " "
End If
Case "="
vl = True
Case "&", "#"
If b <> "" Then 'skip
vl = False
' here is the end
Exit Do
End If
Case Else
If ch = True Then
chs$ = chs$ & Mid$(a, i, 1)
If Len(chs$) = many Then
If many = 4 Then
chs$ = Chr(Int(chs$))
Else
chs$ = StrConv(Chr(CLng("&h" & Mid$(chs$, 5))) + Chr(CLng(Left$(chs$, 4))), vbFromUnicode)
End If
ch = False
If vl Then
c = c + chs$
Else
b = b + chs$
End If
End If
ElseIf vl = False Then
b = b + Mid$(a, i, 1)
Else
c = c + Mid$(a, i, 1)
End If
End Select
Loop
If c <> "" Then Parameters = True
a = Mid$(a, i + 1)
End Function


Public Sub myBreak(basestack As basetask)
   ClearLoadedForms
Dim cc As Object
Set cc = New cRegistry

cc.ClassKey = HKEY_CURRENT_USER
 cc.SectionKey = "Software\"
    cc.SectionKey = basickey

   cc.ValueKey = "FONT"
        cc.ValueType = REG_SZ
If Not cc.KeyExists Then
myBold = True
MYFONT = defFontname
            mNoUseDec = True
            If Not Form1.FontName = MYFONT Then
            MYFONT = "Arial"
            Form1.FontName = MYFONT
            Form1.Font.Italic = False
            Form1.FontName = MYFONT
            
            End If
            MYFONT = Form1.FontName
            
            FFONT = MYFONT
            Err.clear
            DIS.FontName = MYFONT
            DIS.Font.Italic = False
            DIS.FontName = MYFONT
            If Err.Number > 0 Then
            Err.clear
            MYFONT = defFontname
            End If
            If TEXT1.Font.charset <> 161 Then
            
    Font.charset = basestack.myCharSet
    Font.bold = basestack.myBold
    
    DIS.Font.charset = basestack.myCharSet
    DIS.Font.bold = basestack.myBold
    TEXT1.Font.charset = basestack.myCharSet
    TEXT1.Font.bold = basestack.myBold

    List1.Font.charset = basestack.myCharSet
    List1.Font.bold = basestack.myBold

                pagio$ = "LATIN"
                DialogSetupLang 1
                Else
                DIS.Font.charset = basestack.myCharSet
                DialogSetupLang 0
                pagio$ = "GREEK"
            End If
            
                    SzOne = 14
             PenOne = 15
             PaperOne = 1
             DIS.ForeColor = mycolor(PenOne)
             On Error Resume Next
             cc.Value = Form1.FontName
             cc.ValueKey = "NEWSECURENAMES"
                cc.ValueType = REG_DWORD
              cc.Value = -1
             cc.ValueKey = "DIV"
        cc.ValueType = REG_DWORD
        
            UseIntDiv = cc.Value
          
                 cc.ValueKey = "LINESPACE"
        cc.ValueType = REG_DWORD
        If cc.Value >= 0 And cc.Value <= 120 * dv15 Then
     FeedBasket Form1.DIS, players(0), CLng(cc.Value) \ 2
    Else
   FeedBasket Form1.DIS, players(0), CLng(cc.Value) \ 2
    End If
               cc.ValueKey = "SIZE"
        cc.ValueType = REG_DWORD
             cc.Value = 14
              cc.ValueKey = "BOLD"
        cc.ValueType = REG_DWORD
             cc.Value = 1
              cc.ValueKey = "PEN"
        cc.ValueType = REG_DWORD
        cc.Value = 15
          cc.ValueKey = "PAPER"
        cc.ValueType = REG_DWORD
          cc.Value = 1
                 cc.ValueKey = "COMMAND"
        cc.ValueType = REG_SZ
        cc.Value = pagio$
                cc.ValueKey = "HTML"
        cc.ValueType = REG_SZ
        cc.Value = pagiohtml$
       
               cc.ValueKey = "FUNCDEEP"  ' BY DEFAULT
             cc.ValueType = REG_DWORD
             If Not m_bInIDE Then cc.Value = funcdeep
             ' mTextCompare
            cc.ValueKey = "TEXTCOMPARE"
             cc.ValueType = REG_DWORD
                    cc.Value = CLng(mTextCompare)
            cc.ValueKey = "DEC"
             cc.ValueType = REG_DWORD
                    cc.Value = CLng(mNoUseDec)
               cc.ValueKey = "CASESENSITIVE"
        cc.ValueType = REG_SZ
        If cc.Value = vbNullString Then
        If casesensitive = True Then
         cc.Value = "YES"
        Else
    
        cc.Value = "NO"
        End If
        
        End If
        cc.ValueKey = "FOR-LIKE-BASIC"
            cc.ValueType = REG_DWORD
            cc.Value = CLng(0)  ' NO ΒΥ DEFAULT
                 cc.ValueKey = "PRIORITY-OR"
            cc.ValueType = REG_DWORD
            cc.Value = CLng(0)
            priorityOr = False
Else
' *****************************
        If cc.Value = vbNullString Then
        cc.Value = defFontname
        MYFONT = defFontname
        
        Else
        MYFONT = cc.Value
        On Error Resume Next
        
        Me.Font.name = MYFONT
        Me.Font.Italic = False
        Me.Font.name = MYFONT
        If Me.Font.name <> MYFONT Then
        MYFONT = defFontname
        End If
       
        End If
FFONT = MYFONT
Err.clear
DIS.FontName = MYFONT
DIS.Font.Italic = False
DIS.FontName = MYFONT
If Err.Number > 0 Then
Err.clear
MYFONT = defFontname
End If
    cc.ValueKey = "BOLD"
        cc.ValueType = REG_DWORD
        basestack.myBold = cc.Value <> 0
        Form1.Font.bold = basestack.myBold
            cc.ValueKey = "NEWSECURENAMES"
            cc.ValueType = REG_DWORD
            SecureNames = cc.Value
            cc.ValueKey = "DIV"
            cc.ValueType = REG_DWORD
        
            UseIntDiv = cc.Value
         cc.ValueKey = "LINESPACE"
        cc.ValueType = REG_DWORD
        If cc.Value >= 0 And cc.Value <= 120 * dv15 Then
  FeedBasket Form1.DIS, players(0), CLng(cc.Value) \ 2
    Else
  FeedBasket Form1.DIS, players(0), 0
    End If

    cc.ValueKey = "SIZE"
        cc.ValueType = REG_DWORD
        If cc.Value = 0 Then
        cc.Value = 14
        SzOne = 14
        Else
        If cc.Value >= 8 And cc.Value <= 48 Then
        SzOne = cc.Value
        Else
        cc.Value = 14
        SzOne = 14
        End If
        End If
    cc.ValueKey = "PEN"
        cc.ValueType = REG_DWORD
        PenOne = cc.Value
    If Not (PenOne >= 0 And PenOne <= 15) Then PenOne = 15
        
    cc.ValueKey = "PAPER"
        cc.ValueType = REG_DWORD
      
    If cc.Value = PenOne Then cc.Value = 16 - PenOne
   
        DIS.ForeColor = mycolor(PenOne)
    cc.ValueKey = "PAPER"
        cc.ValueType = REG_DWORD
        PaperOne = cc.Value
        cc.ValueKey = "COMMAND"
        cc.ValueType = REG_SZ
        If cc.Value = vbNullString Then
        cc.Value = "GREEK"
        End If
        pagio$ = cc.Value
        cc.ValueKey = "HTML"
        cc.ValueType = REG_SZ
        If cc.Value = vbNullString Then
        cc.Value = "DARK"
        End If
         pagiohtml$ = cc.Value
        cc.ValueKey = "FUNCDEEP"  ' RESET
        cc.ValueType = REG_DWORD
        If Not m_bInIDE Then
            If Not cc.Value = 0 Then funcdeep = cc.Value
            If funcdeep > 3260 Then
            ' fix it
            cc.Value = 3260
            funcdeep = 3260
            End If
            
        Else
        If m_bInIDE Then funcdeep = 128 Else funcdeep = 300
        End If
                    cc.ValueKey = "TEXTCOMPARE"
             cc.ValueType = REG_DWORD
               mTextCompare = CBool(cc.Value)
             cc.ValueKey = "DEC"
             cc.ValueType = REG_DWORD
                     mNoUseDec = CBool(cc.Value)
                     CheckDec
        cc.ValueKey = "CASESENSITIVE"
        cc.ValueType = REG_SZ
       
        If cc.Value = "YES" Then
         casesensitive = True
        Else
    
       casesensitive = False
        End If
            cc.ValueKey = "PRIORITY-OR"
            cc.ValueType = REG_DWORD
             priorityOr = CLng(cc.Value) = True
            ' cc.ValueKey = "FOR-LIKE-BASIC"
             ' ForLikeBasic = CLng(cc.Value) = True
        Set cc = Nothing
        End If
       DIS.ForeColor = mycolor(PenOne) ' NOW PEN IS RGB VALUE
            Font.charset = basestack.myCharSet
    Font.bold = basestack.myBold
    DIS.Font.charset = basestack.myCharSet
    DIS.Font.bold = basestack.myBold
    TEXT1.Font.charset = basestack.myCharSet
    TEXT1.Font.bold = basestack.myBold
    List1.Font.charset = basestack.myCharSet
    List1.Font.bold = basestack.myBold
  
        Select Case pagio$
        Case "GREEK"
         GREEK basestack1
        Case Else   '"LATIN"
            LATIN basestack1
        
         End Select
If OperatingSystem > System_Windows_7 Then
                MouseShow True
                ElseIf basestack.tolayer > 0 Or basestack.toback Then
                 DIS.mousepointer = 1
                          Set DIS.mouseicon = Nothing
                                           
               End If
End Sub

Public Sub mn1sub()
TEXT1.MarkCut
End Sub

Public Sub mn2sub()
TEXT1.MarkCopy
End Sub

Public Sub mn3sub()
On Error Resume Next
Dim aa$
aa$ = GetTextData(13)
If aa$ = vbNullString Then aa$ = Clipboard.GetText(1)
With TEXT1
If .ParaSelStart = 2 And .glistN.list(.glistN.ListIndex) = vbNullString Then
.SelStart = .SelStart - 1
End If
.AddUndo ""
.SelText = aa$
.RemoveUndo aa$
End With
End Sub
Public Sub mn4sub()
Dim gothere As Long
 If Not EditTextWord Then
 ' check if { } is ok...
 If Not blockCheck(TEXT1.Text, DialogLang, gothere) Then
 Exit Sub
 End If
 End If

MyDoEvents
NOEDIT = True

End Sub


Private Sub wdragSub()
TEXT1.glistN.DragEnabled = Not TEXT1.glistN.DragEnabled
End Sub

Public Sub wordwrapsub()
TEXT1.nowrap = Not TEXT1.nowrap
TEXT1.Render
TEXT1.ManualInform
End Sub
Function GetKeY(ascii As Integer) As String
    Dim Buffer As String, ret As Long
    Buffer = String$(514, 0)
    Dim r&, k&
      r = GetKeyboardLayout(DWL_ANYTHREAD) And &HFFFF
      r = CLng(val("&H" & Right(Hex(r), 4)))
    ret = GetLocaleInfo(r, LOCALE_ILANGUAGE, StrPtr(Buffer), Len(Buffer))
    If ret > 0 Then
        GetKeY = ChrW$(AscW(StrConv(ChrW$(ascii Mod 256), 64, CLng(val("&h" + Left$(Buffer, ret - 1))))))
    Else
        GetKeY = ChrW$(AscW(StrConv(ChrW$(ascii Mod 256), 64, 1033)))
    End If
End Function
Public Function GetLCIDFromKeyboard() As Long
    Dim Buffer As String, ret&, r&
    Buffer = String$(514, 0)
      r = GetKeyboardLayout(DWL_ANYTHREAD) And &HFFFF
      r = val("&H" & Right(Hex(r), 4))
        ret = GetLocaleInfo(r, LOCALE_ILANGUAGE, StrPtr(Buffer), Len(Buffer))
    GetLCIDFromKeyboard = CLng(val("&h" + Left$(Buffer, ret - 1)))
End Function
Sub MarkSoftButton(para As Long, pospara As Long)
If TEXT1.glistN.lines = 1 Then Exit Sub
If ShadowMarks Then Exit Sub
If para = 0 Then 'set
    para = TEXT1.mDoc.MarkParagraphID
    pospara = TEXT1.ParaSelStart
    
    If UserCodePage = 1253 Then
        TEXT1.ReplaceTitle = "Ο δείκτης τώρα θα δείχνει αυτή την παράγραφο"
    Else
        TEXT1.ReplaceTitle = "Mark now move to this Paragraph and Position"
    End If
ElseIf para = TEXT1.mDoc.MarkParagraphID And pospara = TEXT1.Charpos Then 'Reset
    para = 0
    
    If UserCodePage = 1253 Then
    TEXT1.ReplaceTitle = "Διαγραφή Δείκτη"
    Else
    TEXT1.ReplaceTitle = "Mark Deleted"
    End If
Else ' goto that paragraph
    If Not TEXT1.mDoc.InvalidPara(para) Then
        TEXT1.SelLengthSilent = 0
        TEXT1.mDoc.MarkParagraphID = para
        TEXT1.glistN.enabled = False
        TEXT1.ParaSelStart = pospara
        TEXT1.glistN.enabled = True
        TEXT1.ManualInform
    Else
        para = 0
        If UserCodePage = 1253 Then
            TEXT1.ReplaceTitle = "Δεν βρέθηκε παράγραφος - διαγράφτηκε ο δείκτης"
        Else
            TEXT1.ReplaceTitle = "Paragraph noτ found - mark deleted"
        End If
    End If
End If
End Sub

Function Mark$()
If ShadowMarks Then Mark$ = vbNullString: Exit Function
If TEXT1.Title = vbNullString Then  'reset all para
para1 = 0: Para2 = 0: Para3 = 0
ElseIf LastDocTitle$ <> TEXT1.Title Then
para1 = 0: Para2 = 0: Para3 = 0
LastDocTitle$ = TEXT1.Title
End If
Dim s$
If para1 <> 0 Then
If TEXT1.mDoc.InvalidPara(para1) Then para1 = 0
If para1 = TEXT1.mDoc.MarkParagraphID Then
s$ = " [F6] "
Else
s$ = " *F6 "
End If
Else
s$ = " -F6" + ChrW(&H25CA)
End If
If Para2 <> 0 Then
If TEXT1.mDoc.InvalidPara(Para2) Then Para2 = 0
If Para2 = TEXT1.mDoc.MarkParagraphID Then
s$ = s$ + " [F7] "
Else
s$ = s$ + " *F7"
End If
Else
s$ = s$ + " -F7" + ChrW(&H25CA)
End If
If Para3 <> 0 Then
If TEXT1.mDoc.InvalidPara(Para3) Then Para3 = 0
If Para3 = TEXT1.mDoc.MarkParagraphID Then
s$ = s$ + " [F8] "
Else
s$ = s$ + " *F8 "
End If
Else
s$ = s$ + " -F8" + ChrW(&H25CA)
End If
Mark$ = s$

End Function
Public Sub ResetMarks()
para1 = 0: Para2 = 0: Para3 = 0


End Sub
Public Sub hookme(this As gList)
Set LastGlist = this
End Sub
Public Function mybreak1() As Boolean
Dim i As Long
If Form4.Visible Then Form4.Visible = False
i = MOUT
If ASKINUSE Then
If BreakMe Then Exit Function
Unload NeoMsgBox: ASKINUSE = False: Exit Function
End If
BreakMe = True

INK$ = vbNullString
If MsgBoxN("Break Key - Hard Reset" + vbCrLf + "Μ2000 - Execution Stop / Τερματισμός Εκτέλεσης", vbYesNo, MesTitle$) <> vbNo Then
                
                
                If AVIRUN Then AVI.GETLOST
                On Error Resume Next
                LastErName = vbNullString
                LastErNum = 0
                LastErNum1 = 0
                If Me.Visible Then Me.SetFocus
                closeAll ' we closed all files
                QRY = False
                GFQRY = False
                escok = True
                INK$ = Chr$(27) + Chr$(27)
                If List1.Visible Then
                                List1.Tag = vbNullString
                                List1.Visible = False
                                List1.LeaveonChoose = False
                                INK$ = vbNullString
                End If
                
                mybreak1 = True
End If
 
 
BreakMe = False
End Function
Public Sub SetText1()
If (600 - hueconv(TEXT1.backcolor)) Mod 360 > 30 And lightconv(TEXT1.backcolor) >= 128 Then TEXT1.ColorSet = 1 Else TEXT1.ColorSet = 0
Select Case fState
Case 0
shortlang = False
TEXT1.NoColor = EditTextWord
Case 1
shortlang = False
TEXT1.NoColor = True
Case 2
shortlang = True
TEXT1.NoColor = EditTextWord
Case 3
shortlang = True
TEXT1.NoColor = True
fState = -1
End Select

TEXT1.mDoc.ColorEvent = Not TEXT1.NoColor

End Sub



