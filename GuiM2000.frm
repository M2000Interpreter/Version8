VERSION 5.00
Begin VB.Form GuiM2000 
   AutoRedraw      =   -1  'True
   BackColor       =   &H003B3B3B&
   BorderStyle     =   0  'None
   Caption         =   "aaa"
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9210
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "GuiM2000.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame ResizeMark 
      Appearance      =   0  'Flat
      BackColor       =   &H003B3B3B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   150
      Left            =   8475
      TabIndex        =   0
      Top             =   4080
      Visible         =   0   'False
      Width           =   135
   End
   Begin M2000.gList gList2 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   873
      Max             =   1
      Vertical        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Backcolor       =   3881787
      ForeColor       =   16777215
      CapColor        =   16777215
   End
End
Attribute VB_Name = "GuiM2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CopyFromLParamToRect Lib "User32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Private Declare Function DestroyCaret Lib "User32" () As Long
Dim setupxy As Single
Dim Lx As Single, ly As Single, dr As Boolean
Dim scrTwips As Long
Dim bordertop As Long, borderleft As Long
Dim allwidth As Long, itemWidth As Long
Private ExpandWidth As Boolean, lastfactor As Single
Private myEvent As mEvent
Private GuiControls As New Collection
'Dim gList1 As gList
Dim onetime As Boolean, PopupOn As Boolean
Dim alfa As New GuiButton
Public MyName$
Public modulename$
Public prive As Long
Private ByPassEvent As Boolean
Private mIndex As Long
Private mSizable As Boolean
Public Relax As Boolean
Private MarkSize As Long
Public MY_BACK As New cDIBSection
Dim CtrlFont As New StdFont
Dim novisible As Boolean
Private mModalId As Double, mModalIdPrev As Double
Private mPopUpMenu As Boolean
Public IamPopUp As Boolean
Private mEnabled As Boolean
Public WithEvents mDoc As Document
Attribute mDoc.VB_VarHelpID = -1
Dim mQuit
Public Sub AddGuiControl(widget As Object)
GuiControls.Add widget
End Sub
Public Sub TestModal(alfa As Double)
If mModalId = alfa Then
mModalId = mModalIdPrev
mModalIdPrev = 0
Enablecontrol = True
End If
End Sub
Property Get Modal() As Double
    Modal = mModalId
End Property
Property Let Modal(RHS As Double)
mModalIdPrev = mModalId
mModalId = RHS
End Property
Public Property Get PopUpMenuVal() As Boolean
PopUpMenuVal = mPopUpMenu
End Property
Public Property Let PopUpMenuVal(RHS As Boolean)
mPopUpMenu = RHS
End Property
Public Property Let Enablecontrol(RHS As Boolean)
If RHS = False Then UnHook hWND '  And Not Me Is Screen.ActiveForm Then UnHook hWnd
If Len(MyName$) = 0 Then Exit Property
'If rhs = Fals Then UnHook hWnd
If mEnabled = False And RHS = True Then Me.enabled = True
mEnabled = RHS

Dim w As Object
If Controls.Count > 0 Then
For Each w In Me.Controls
If w Is gList2 Then
gList2.enabled = RHS
gList2.mousepointer = 0
ElseIf w.Visible Then
w.enabled = RHS
If TypeOf w Is gList Then w.TabStop = RHS
End If
Next w
End If
Me.enabled = RHS
End Property
Public Property Get Enablecontrol() As Boolean
If Len(MyName$) = 0 Then Enablecontrol = False: Exit Property
Enablecontrol = mEnabled


End Property


Property Get NeverShow() As Boolean
NeverShow = Not novisible
End Property
Friend Property Set EventObj(aEvent As Object)
Set myEvent = aEvent
Set myEvent.excludeme = New FastCollection
End Property

Public Sub Callback(b$)
If Quit Then Exit Sub
If myEvent Is Nothing Then
Set EventObj = New mEvent
End If
If ByPassEvent Then
    If myEvent.excludeme.IamBusy Then Exit Sub
    Dim Mark$
    Mark$ = Split(b$, "(")(0)
    If myEvent.excludeme.ExistKey3(Mark$) Then Exit Sub
    If Not TaskMaster Is Nothing Then TaskMaster.tickdrop = 0
    
    If Visible Then
       myEvent.excludeme.AddKey2 Mark$
    If CallEventFromGuiOne(Me, myEvent, b$) Then
       If Not Quit Then myEvent.excludeme.Remove Mark$
    End If
    Else
        CallEventFromGuiOne Me, myEvent, b$
    End If
Else
    CallEventFromGui Me, myEvent, b$
End If
End Sub
Public Sub CallbackNow(b$, VR())
If Quit Then Exit Sub
If myEvent Is Nothing Then
Set EventObj = New mEvent
End If

If myEvent.excludeme.IamBusy Then Exit Sub
Dim Mark$
Mark$ = Split(b$, "(")(0)
If myEvent.excludeme.ExistKey3(Mark$) Then Exit Sub
If Visible Then myEvent.excludeme.AddKey2 Mark$
If CallEventFromGuiNow(Me, myEvent, b$, VR()) Then myEvent.excludeme.Remove Mark$

End Sub


Public Sub ShowmeALL()
Dim w As Object

If Controls.Count > 0 Then
For Each w In Controls
If w.enabled Then w.Visible = True
Next w
End If

gList2.PrepareToShow
End Sub
Public Sub RefreshALL()
Dim w As Object
If Controls.Count > 0 Then
For Each w In Controls
If w.Visible Then
If TypeOf w Is gList Then w.ShowMe2
End If
Next w
End If
Refresh
End Sub

Private Sub Form_Click()
If gList2.Visible Then gList2.SetFocus
If mIndex > -1 Then
    Callback MyName$ + ".Click(" + CStr(index) + ")"
Else
    Callback MyName$ + ".Click()"
End If
End Sub

Private Sub Form_Activate()
On Error Resume Next
If Not Quit Then
If myEvent Is Nothing Then
Set EventObj = New mEvent
End If
If Not myEvent.excludeme.IamBusy Then
Set myEvent.excludeme = New FastCollection
End If
End If
If PopupOn Then PopupOn = False
If novisible Then Hide: Unload Me
If gList2.HeadLine <> "" Then If ttl Then Form3.CaptionW = gList2.HeadLine: Form3.Refresh
MarkSize = 4
ResizeMark.width = MarkSize * dv15
ResizeMark.Height = MarkSize * dv15
ResizeMark.Left = width - MarkSize * dv15
ResizeMark.Top = Height - MarkSize * dv15

ResizeMark.BackColor = GetPixel(Me.hdc, 0, 0)
ResizeMark.Visible = Sizable
If Sizable Then ResizeMark.ZOrder 0

If Typename(ActiveControl) = "gList" Then
Hook hWND, ActiveControl
Else
Hook hWND, Nothing
End If

End Sub
Private Sub Form_Deactivate0()
If PopupOn Then
UnHook hWND

Exit Sub
End If
If IamPopUp Then
If mModalId = ModalId And ModalId <> 0 Then
        
        If Visible Then Hide
       
        ModalId = 0
            novisible = False
End If
Else
    If mModalId = ModalId And ModalId <> 0 Then
        If Visible Then
            On Error Resume Next
            Me.SetFocus
        Else
        UnHook hWND
            If mModalId <> 0 Then ModalId = 0
 
            
        End If
    
    Else
    UnHook hWND
    End If
   
    End If
End Sub


Private Sub Form_Deactivate()
            UnHook hWND
If PopupOn Then

Exit Sub
End If
If IamPopUp Then
If mModalId = ModalId And ModalId <> 0 Then
If Visible Then Hide
ModalId = 0
novisible = False
End If
Else
If mModalId = ModalId And ModalId <> 0 Then If Not Visible Then If mModalId <> 0 Then ModalId = 0
End If

End Sub


Private Sub Form_Initialize()
'myEvent.excludeme.ResetGui
mEnabled = True
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If Me.Visible Then
If ActiveControl Is Nothing Then
Dim w As Object
    If Controls.Count > 0 Then
    For Each w In Controls
    If w.Visible Then
    If TypeOf w Is gList Then
    w.SetFocus
    Exit For
    End If
    End If
    Next w
    Set w = Nothing
    End If
    Else
    
    If Typename(ActiveControl) = "gList" Then ActiveControl.SetFocus
End If
Else
'Debug.Print MyName$
choosenext
End If
End Sub

Private Sub Form_LostFocus()
If mIndex > -1 Then
    Callback MyName$ + ".LostFocus(" + CStr(index) + ")"
Else
    Callback MyName$ + ".LostFocus()"
End If
If HOOKTEST <> 0 Then
UnHook hWND
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
If Not Relax Then



Relax = True
If mIndex > -1 Then
    Callback MyName$ + ".MouseDown(" + CStr(index) + "," + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
Else
    Callback MyName$ + ".MouseDown(" + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
End If



Relax = False
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
If Not Relax Then
Relax = True

If mIndex > -1 Then
Callback MyName$ + ".MouseMove(" + CStr(index) + "," + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
Else
Callback MyName$ + ".MouseMove(" + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
End If
Relax = False
End If

End Sub

Private Sub Form_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
If Not Relax Then

Relax = True

If mIndex > -1 Then
Callback MyName$ + ".MouseUp(" + CStr(index) + "," + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
Else
Callback MyName$ + ".MouseUp(" + CStr(Button) + "," + CStr(shift) + "," + CStr(x) + "," + CStr(y) + ")"
End If
Relax = False
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If mModalId = ModalId And ModalId <> 0 Then
    If Visible Then Hide: Quit = True
    If mModalId <> 0 Then ModalId = 0
    Cancel = True
    novisible = False
ElseIf mModalId <> 0 And Visible Then
    mModalId = mModalIdPrev
    mModalIdPrev = 0
    If mModalId > 0 Then
        Cancel = True
    Else
      '    Set LastGlist = Nothing
        Quit = True
    End If
Else
mModalIdPrev = 0
'Set LastGlist = Nothing
Quit = True
End If
End Sub

Private Sub Form_Resize()
gList2.MoveTwips 0, 0, Me.width, gList2.HeightTwips
ResizeMark.Move width - ResizeMark.width, Height - ResizeMark.Height
End Sub


Private Sub gList2_CtrlPlusF1()
    If mIndex > -1 Then
        Callback MyName$ + ".About(" + CStr(index) + ")"
    Else
        Callback MyName$ + ".About()"
    End If
End Sub

Private Sub gList2_EnterOnly()
    If mIndex > -1 Then
        Callback MyName$ + ".Enter(" + CStr(index) + ")"
    Else
        Callback MyName$ + ".Enter()"
    End If
End Sub

Private Sub gList2_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
If item = -1 Then
FillThere thisHDC, thisrect, gList2.CapColor
FillThereMyVersion thisHDC, thisrect, &H999999
skip = True
End If
End Sub
Private Sub gList2_ExposeItemMouseMove(Button As Integer, ByVal item As Long, ByVal x As Long, ByVal y As Long)
If gList2.DoubleClickCheck(Button, item, x, y, 10 * lastfactor, 10 * lastfactor, 8 * lastfactor, -1) Then
    ByeBye
End If
End Sub
Sub ByeBye()
Dim var(1) As Variant
var(0) = CLng(0)
If mIndex > -1 Then
If Not Quit Then CallEventFromGuiNow Me, myEvent, MyName$ + ".Unload(" + CStr(mIndex) + ")", var()
Else
If Not Quit Then CallEventFromGuiNow Me, myEvent, MyName$ + ".Unload()", var()
End If
If var(0) = 0 Then
                     If ttl Then
                     Form3.CaptionW = vbNullString
                     If Form3.WindowState = 1 Then Form3.WindowState = 0
               
                    Unload Form3
             End If
                              Unload Me
                      End If
End Sub

Private Sub Form_Load()
If onetime Then
novisible = True
Exit Sub
End If
onetime = True
mQuit = False
' try0001
Set LastGlist = Nothing
scrTwips = Screen.TwipsPerPixelX
' clear data...
lastfactor = 1
setupxy = 20
gList2.Font.Size = 14.25 * dv15 / 15
gList2.enabled = True
gList2.CapColor = rgb(255, 160, 0)
gList2.FloatList = True
gList2.MoveParent = True
gList2.HeadLine = vbNullString
gList2.HeadLine = "Form"
gList2.HeadlineHeight = gList2.HeightPixels
gList2.SoftEnterFocus
gList2.TabStop = False
With gList2.Font
CtrlFont.name = .name
CtrlFont.Size = .Size
CtrlFont.bold = .bold
End With
gList2.FloatLimitTop = VirtualScreenHeight() - 600
gList2.FloatLimitLeft = VirtualScreenWidth() - 450
End Sub

Private Sub Form_Unload(Cancel As Integer)
UNhookMe
Quit = True
Set myEvent = Nothing

If prive <> 0 Then
players(prive).used = False
players(prive).MAXXGRAPH = 0  '' as a flag
prive = 0
End If
Dim w As Object
If GuiControls.Count > 0 Then
For Each w In GuiControls
    w.deconstruct
Next w
End If
End Sub
Private Sub FillBack(thathDC As Long, there As RECT, bgcolor As Long)
' create brush
Dim my_brush As Long
my_brush = CreateSolidBrush(bgcolor)
FillRect thathDC, there, my_brush
DeleteObject my_brush
End Sub
Private Sub FillThere(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim a As RECT
CopyFromLParamToRect a, thatRect

FillBack thathDC, a, thatbgcolor
End Sub

Private Sub FillThereMyVersion(thathDC As Long, thatRect As Long, thatbgcolor As Long)
Dim a As RECT, b As Long
b = 2 * lastfactor
If b < 2 Then b = 2
If setupxy - b < 0 Then b = setupxy \ 4 + 1
CopyFromLParamToRect a, thatRect
a.Left = b
a.Right = setupxy - b
a.Top = b
a.Bottom = setupxy - b
FillThere thathDC, VarPtr(a), 0
b = 5 * lastfactor
a.Left = b
a.Right = setupxy - b
a.Top = b
a.Bottom = setupxy - b
FillThere thathDC, VarPtr(a), rgb(255, 160, 0)
End Sub

Public Property Get Title() As Variant
Title = gList2.HeadLine
End Property

Public Property Let Title(ByVal vNewValue As Variant)
' A WORKAROUND TO CHANGE TITLE WHEN FORM IS DISABLED BY A MODAL FORM
On Error Resume Next
Dim oldenable As Boolean
oldenable = gList2.enabled
gList2.enabled = True
gList2.HeadLine = vbNullString
If Trim(vNewValue) = vbNullString Then vNewValue = " "
gList2.HeadLine = vNewValue
gList2.HeadlineHeight = gList2.HeightPixels
'If oldenable = False Then
gList2.ShowMe
gList2.enabled = oldenable
End Property
Public Property Get index() As Long
index = mIndex
End Property

Public Property Let index(ByVal RHS As Long)
mIndex = RHS
End Property
Public Sub CloseNow()
Dim w As Object
    If mModalId = ModalId And ModalId <> 0 Then
        ModalId = 0
      If Visible Then Hide
    Else
    mModalId = 0
    For Each w In GuiControls
    If Typename(w) Like "Gui*" Then
    w.deconstruct
    End If
Next w
Set w = Nothing
         If ttl Then
                     Form3.CaptionW = vbNullString
                     If Form3.WindowState = 1 Then Form3.WindowState = 0
               
                    Unload Form3
             End If

Unload Me
    End If
End Sub
Public Function Control(index) As Object
On Error Resume Next
Set Control = Controls(index)
If Err > 0 Then Set Control = Me
End Function
Public Sub Opacity(mAlpha, Optional mlColor = 0, Optional mTRMODE = 0)
SetTrans Me, CInt(Abs(mAlpha)) Mod 256, CLng(mycolor(mlColor)), CBool(mTRMODE)
End Sub
Public Sub Hold()
MY_BACK.ClearUp
If MY_BACK.Create(Form1.width / DXP, Form1.Height / DYP) Then
MY_BACK.LoadPictureBlt hdc
If MY_BACK.bitsPerPixel <> 24 Then Conv24 MY_BACK
End If
End Sub
Public Sub Release()
MY_BACK.PaintPicture hdc
End Sub


Public Property Get ByPass() As Variant
ByPass = ByPassEvent
End Property

Public Property Let ByPass(ByVal vNewValue As Variant)
ByPassEvent = CBool(vNewValue)
End Property
Property Get TitleHeight() As Variant
TitleHeight = gList2.Height
End Property
Public Sub FontAttr(ThisFontName, Optional ThisMode = -1, Optional ThisBold = True)
Dim aa As New StdFont
If ThisFontName <> "" Then

aa.name = ThisFontName

If ThisMode > 7 Then aa.Size = ThisMode Else aa = 7
aa.bold = ThisBold
Set gList2.Font = aa
gList2.Height = gList2.HeadlineHeightTwips
lastfactor = gList2.HeadlineHeight / 30
setupxy = 20 * lastfactor
 gList2.Dynamic

End If
End Sub
Public Sub CtrlFontAttr(ThisFontName, Optional ThisMode = -1, Optional ThisBold = True)

If ThisFontName <> "" Then

CtrlFont.name = ThisFontName

If ThisMode > 7 Then CtrlFont.Size = ThisMode Else CtrlFont = 7
CtrlFont.bold = ThisBold

End If
End Sub
Public Property Get CtrlFontName()
    CtrlFontName = CtrlFont.name
End Property
Public Property Get CtrlFontSize()
    CtrlFontSize = CtrlFont.Size
End Property
Public Property Get CtrlFontBold()
    CtrlFontBold = CtrlFont.bold
End Property

Private Sub gList2_KeyDown(KeyCode As Integer, shift As Integer)
'
Dim VR(2)
VR(0) = KeyCode
VR(1) = shift
If mIndex > -1 Then
    CallbackNow MyName$ + ".KeyDown(" + CStr(index) + ")", VR()
Else
    CallbackNow MyName$ + ".KeyDown()", VR()
End If
shift = VR(1)
KeyCode = VR(0)

End Sub

Private Sub gList2_RefreshDesktop()
If Form1.Visible Then Form1.Refresh: If Form1.DIS.Visible Then Form1.DIS.Refresh
End Sub
Public Sub PopUp(vv As Variant, ByVal x As Variant, ByVal y As Variant)
Dim var1() As Variant, retobject As Object, that As Object, hmonitor As Long
ReDim var1(0 To 1)
Dim var2() As String
ReDim var2(0 To 0)
hmonitor = FindFormSScreen(Me)
x = x + Left
y = y + Top
Set that = vv
If Me Is that Then Exit Sub
If that.Visible Then
If Not that.enabled Then Exit Sub
End If
If x + that.width > ScrInfo(hmonitor).width + ScrInfo(hmonitor).Left Then
If y + that.Height > ScrInfo(hmonitor).Height + ScrInfo(hmonitor).Top Then
that.Move ScrInfo(hmonitor).width - that.width + ScrInfo(hmonitor).Left, ScrInfo(hmonitor).Height - that.Height + ScrInfo(hmonitor).Top
Else
that.Move ScrInfo(hmonitor).width - that.width + ScrInfo(hmonitor).Left, y + ScrInfo(hmonitor).Top
End If
ElseIf y + that.Height > ScrInfo(hmonitor).Height + ScrInfo(hmonitor).Top Then
that.Move x, ScrInfo(hmonitor).Height - Height + ScrInfo(hmonitor).Top
Else
that.Move x, y
End If
var1(1) = 1
Set var1(0) = Me
that.IamPopUp = True
CallByNameFixParamArray that, "Show", VbMethod, var1(), var2(), 2
Set that = Nothing
Set var1(0) = Nothing
'Show
MyDoEvents

End Sub
Public Sub PopUpPos(vv As Variant, ByVal x As Variant, ByVal y As Variant, ByVal y1 As Variant)
Dim that As Object, hmonitor As Long
x = x + Left
y = y + Top + y1
hmonitor = FindFormSScreen(Me)
Set that = vv
If Me Is that Then Exit Sub
If that.Visible Then
If Not that.enabled Then Exit Sub
End If
If x + that.width > ScrInfo(hmonitor).width + ScrInfo(hmonitor).Left Then
If y + that.Height > ScrInfo(hmonitor).Height + ScrInfo(hmonitor).Top Then
that.Move ScrInfo(hmonitor).width + ScrInfo(hmonitor).Left - that.width, y - that.Height - y1 + ScrInfo(hmonitor).Top
Else
that.Move ScrInfo(hmonitor).width + ScrInfo(hmonitor).Left - that.width, y + ScrInfo(hmonitor).Top
End If
ElseIf y + that.Height > ScrInfo(hmonitor).Height + ScrInfo(hmonitor).Top Then
that.Move x, y - that.Height - y1 + ScrInfo(hmonitor).Top
Else
that.Move x, y
End If
that.ShowmeALL
'If ModalId <> 0 Then
PopupOn = True

that.Show , Me

End Sub
Public Sub hookme(this As gList)
Set LastGlist = this
End Sub

Private Sub mDoc_MayQuit(Yes As Variant)
If mQuit Or Not Visible Then Yes = True
MyDoEvents1 Me
'ProcTask2 basestack1
End Sub

Private Sub ResizeMark_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
If Sizable And Not dr Then
    x = x + ResizeMark.Left
    y = y + ResizeMark.Top
    If (y > Height - 150 And y < Height) And (x > width - 150 And x < width) Then
    
    dr = Button = 1
    ResizeMark.mousepointer = vbSizeNWSE
    Lx = x
    ly = y
    If dr Then Exit Sub
    
    End If
    
End If
End Sub

Private Sub ResizeMark_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Dim addy As Single, addX As Single
If Not Relax Then
    x = x + ResizeMark.Left
    y = y + ResizeMark.Top
    If Button = 0 Then If dr Then Me.mousepointer = 0: dr = False: Relax = False: Exit Sub
    Relax = True
    If dr Then
         If y < (Height - 150) Or y >= Height Then addy = (y - ly) Else addy = dv15 * 5
         If x < (width - 150) Or x >= width Then addX = (x - Lx) Else addX = dv15 * 5
         If width + addX >= 1800 And width + addX < VirtualScreenWidth() Then
             If Height + addy >= 1800 And Height + addy < VirtualScreenHeight() Then
                Lx = x
                ly = y
                Move Left, Top, width + addX, Height + addy
                If mIndex > -1 Then
                    Callback MyName$ + ".Resize(" + CStr(index) + ")"
                Else
                    Callback MyName$ + ".Resize()"
                End If
            End If
        End If
        Relax = False
        Exit Sub
    Else
        If Sizable Then
            If (y > Height - 150 And y < Height) And (x > width - 150 And x < width) Then
                    dr = Button = 1
                    ResizeMark.mousepointer = vbSizeNWSE
                    Lx = x
                    ly = y
                    If dr Then Relax = False: Exit Sub
                Else
                    ResizeMark.mousepointer = 0
                    dr = 0
                End If
            End If
    End If
Relax = False
End If
End Sub

Public Property Get Sizable() As Variant
Sizable = mSizable
End Property

Public Property Let Sizable(ByVal vNewValue As Variant)
mSizable = vNewValue
ResizeMark.enabled = vNewValue
If ResizeMark.enabled Then
ResizeMark.Visible = Me.Visible
Else
ResizeMark.Visible = False
End If
End Property
Public Property Let SizerWidth(ByVal vNewValue As Variant)
If vNewValue \ dv15 > 1 Then
    MarkSize = vNewValue \ dv15
    With ResizeMark
    .width = MarkSize * dv15
    .Height = MarkSize * dv15
    .Move width - .width, Height - .Height
    End With
End If
End Property

Public Property Get header() As Variant
header = gList2.Visible
End Property

Public Property Let header(ByVal vNewValue As Variant)
gList2.Visible = vNewValue
End Property


Sub GetFocus()
On Error Resume Next
Me.SetFocus
End Sub
Public Sub UNhookMe()
Set LastGlist = Nothing
UnHook hWND
End Sub

Public Property Get Quit() As Variant
Quit = mQuit
End Property

Public Property Let Quit(ByVal vNewValue As Variant)
mQuit = vNewValue
End Property
