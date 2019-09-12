VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuiTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' If you copy and paste code
' don't forget to set default property to vartext (changed from text)
Public WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Private mprompt As String
Private mkind As String, mPartColor As Long
Public SkipProcAtExitFocus As Boolean
Private Declare Function GlobalAlloc Lib "KERNEL32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "KERNEL32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_DISCARDABLE = &H100
Private Const GMEM_DISCARDED = &H4000
Private Const GMEM_FIXED = &H0
Private Const GMEM_INVALID_HANDLE = &H8000
Private Const GMEM_LOCKCOUNT = &HFF
Private Const GMEM_MODIFY = &H80
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_NOCOMPACT = &H10
Private Const GMEM_NODISCARD = &H20
Private Const GMEM_NOT_BANKED = &H1000
Private Const GMEM_NOTIFY = &H4000
Private Const GMEM_SHARE = &H2000
Private Const GMEM_VALID_FLAGS = &H7F72
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Const GMEM_LOWER = GMEM_NOT_BANKED
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetClipboardData Lib "user32" _
    (ByVal wFormat As Long) As Long
    Private Declare Function SetClipboardData Lib "user32" _
    (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function OpenClipboard Lib "user32" _
    (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" _
    () As Long
    Private Declare Function IsClipboardFormatAvailable Lib "user32" _
    (ByVal wFormat As Long) As Long
    Private Pad$
    ' for Gui purposes
Public ctrlName As String
Private mesCtrlName As String
Dim mIndex As Long, mStayOpen As Boolean
Dim Callback As GuiM2000
Dim white As Long, black As Long, dimm1 As Long, dimm2 As Long
Public Property Get GetCallBack() As Form
Set GetCallBack = Callback
End Property
Public Sub Construct(there As Form, name$)
 mIndex = -1
 Set Callback = there
 On Error Resume Next
 ctrlName = GetTag$
 Set glistN = there.Controls.Add("m2000.glist", ctrlName)
 there.AddGuiControl Me
 glistN.Dynamic
 there.Controls(ctrlName).TabStop = True
 mesCtrlName = name$
End Sub
Public Sub ConstructArray(there As Form, name$, i As Long)
 Set Callback = there
 On Error Resume Next
 ctrlName = GetTag$
 Set glistN = there.Controls.Add("m2000.glist", ctrlName)
 mIndex = i
 there.AddGuiControl Me
 glistN.Dynamic
 there.Controls(ctrlName).TabStop = True
 mesCtrlName = name$
End Sub
Public Property Get Text() As String
If glistN.listcount = 0 Then
glistN.additem mprompt
Else
Text = glistN.list(0)
End If
End Property
Public Property Get vartext() As String
If glistN.listcount = 0 Then
glistN.additem ""
Else
vartext = Mid$(glistN.list(0), 1 + Len(mprompt))
End If
End Property
Public Property Let vartext(ByVal RHS As String)
glistN.restrictLines = 1
glistN.HeadLine = vbNullString ' no info
glistN.HeadlineHeight = 0

If glistN.listcount = 0 Then
glistN.additem mprompt + RHS
Else
glistN.list(0) = mprompt + RHS
End If
glistN.SelStart = Len(mprompt + RHS) + 1
glistN.ShowMe
End Property
Public Property Let Text(ByVal RHS As String)
'Dim oldright As Long
'oldright = glistN.LeftMarginPixels
On Error Resume Next
If glistN Is Nothing Then Exit Property
glistN.restrictLines = 1
glistN.HeadLine = vbNullString ' no info
glistN.HeadlineHeight = 0
If glistN.listcount = 0 Then
glistN.additem RHS
Else
glistN.list(0) = RHS
End If
'glistN.LeftMarginPixels = oldright

glistN.ShowMe2
End Property
Public Property Get Info() As String

Info = glistN.HeadLine

End Property

Public Property Let Info(ByVal RHS As String)
locked = False
If RHS = vbNullString Then
glistN.InfoDropBarClick = False
Text = Text
Else
glistN.restrictLines = 0
glistN.HeadLineHeightMinimum = glistN.HeightPixels
glistN.HeadLine = RHS
glistN.InfoDropBarClick = True
glistN.ShowMe
locked = True
End If

End Property

Public Property Get top() As Variant
top = Callback.Controls(ctrlName).top
End Property
Public Property Let top(ByVal RHS As Variant)
Callback.Controls(ctrlName).top = RHS
End Property
Public Property Get Left() As Variant
Left = Callback.Controls(ctrlName).Left
End Property

Public Property Let Left(ByVal RHS As Variant)
Callback.Controls(ctrlName).Left = RHS
End Property


Public Sub SetUp()
black = 0
dimm1 = rgb(100, 200, 160)
dimm2 = rgb(200, 120, 60)
Dim ff As New StdFont
With Callback
white = .Controls(1).ForeColor
ff.name = .CtrlFontName
ff.charset = Form1.DIS.Font.charset
ff.Size = .CtrlFontSize
ff.bold = .CtrlFontBold
End With

With glistN
.enabled = True
Info = vbNullString
If .listcount = 0 Then Text = vbNullString
.NoCaretShow = True
.VerticalCenterText = True
If Not .CenterText Then .LeftMarginPixels = 8 Else .LeftMarginPixels = 0
.backcolor = rgb(200, 120, 60)
Set .Font = ff
.ForeColor = white
.dcolor = Callback.Controls(1).dcolor
.ShowMe
End With
locked = False
enabled = True
Retired

End Sub


Public Property Get Prompt() As String
Prompt = mprompt
End Property

Public Property Let Prompt(ByVal RHS As String)
mprompt = RHS
glistN.PromptLineIdent = Len(RHS)
Text = RHS
End Property


Private Sub Class_Terminate()
Set glistN = Nothing
End Sub
Private Sub glistN_OutPopUp(X As Single, Y As Single, myButton As Integer)
If myButton And &H2 = 2 Then
If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".PopUp(" + Str(mIndex) + "," + Str(X + Callback.Controls(ctrlName).Left) + "," + Str(Y + Callback.Controls(ctrlName).top) + "," + Str(myButton) + "," + Str(glistN.ListIndex + 1) + ")"
   Else
      Callback.Callback mesCtrlName + ".PopUp(" + Str(X + Callback.Controls(ctrlName).Left) + "," + Str(Y + Callback.Controls(ctrlName).top) + "," + Str(myButton) + "," + Str(glistN.ListIndex + 1) + ")"
      End If
      End If
End Sub

Private Sub glistN_RegisterGlist(this As gList)
If Not glistN.Spinner Then this.NoWheel = True
On Error Resume Next
Set LastGlist = this
If Err.Number > 0 Then this.NoWheel = True
End Sub



Private Sub glistN_ChangeListItem(item As Long, content As String)
'If Not enabled Then Exit Sub
Dim this As String, thispos As Long, VR(2)
thispos = glistN.SelStart - Len(mprompt)
this = Mid$(content, 1 + Len(mprompt))
'' RaiseEvent ValidString(this, thispos)
VR(0) = this
VR(1) = thispos
   If mIndex >= 0 Then
   Callback.CallbackNow mesCtrlName + ".ValidString(" + Str(mIndex) + ")", VR()
   Else
      Callback.CallbackNow mesCtrlName + ".ValidString()", VR()
      End If
      
thispos = VR(1) + Len(mprompt)
content = mprompt + VR(0)
If thispos <> glistN.SelStart And VR(0) <> "" Then glistN.SelStart = thispos: glistN.ShowMe
End Sub
Private Sub glistN_CtrlPlusF1()
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".About(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".About()"
      End If
End Sub
Public Sub ResetPan()
glistN.PanPos = 0
End Sub

Private Sub glistN_ChangeSelStart(thisselstart As Long)
If thisselstart <= Len(mprompt) Then thisselstart = Len(mprompt) + 1
'glistN.ShowPan
With glistN
.ShowPan
.ExternalCursor .SelStart, .list(.ListIndex), .ForeColor
End With
End Sub

Public Property Get ThisKind() As String
ThisKind = mkind
End Property

Public Property Let ThisKind(ByVal RHS As String)
mkind = RHS
glistN.LastLinePart = " " & RHS
End Property


Public Property Get enabled() As Boolean
enabled = glistN.EditFlag
End Property

Public Property Let enabled(ByVal RHS As Boolean)
glistN.EditFlag = RHS
glistN.NoCaretShow = Not RHS
glistN.NoBarClick = Not RHS
End Property
'


Private Sub glistN_CheckGotFocus()

If Not locked Then
If glistN.ListIndex < 0 And glistN.listcount > 0 Then glistN.ListindexPrivateUse = 0
' RaiseEvent GotFocus
If enabled Then
glistN.FadeLastLinePart = mPartColor
If glistN.BackStyle <> 1 Then
glistN.ForeColor = black
glistN.backcolor = white
End If
glistN.NoCaretShow = False:
glistN.PanPos = 0
Else
glistN.FadeLastLinePart = black
If glistN.BackStyle <> 1 Then
glistN.ForeColor = white
glistN.backcolor = dimm1
End If
glistN.NoCaretShow = True
End If
glistN.ShowMe
End If

End Sub
Public Sub Retired()
If mStayOpen Then
glistN.FadeLastLinePart = mPartColor
If glistN.BackStyle <> 1 Then
glistN.ForeColor = black
glistN.backcolor = white
End If
glistN.NoCaretShow = False
Else
If glistN.BackStyle <> 1 Then
glistN.FadeLastLinePart = 0
glistN.ForeColor = white
glistN.backcolor = dimm2
End If
glistN.NoCaretShow = True
End If
 glistN.PanPos = 0: glistN.ShowMe
End Sub

Private Sub glistN_CheckLostFocus()
If Not locked Then
Callback.LastActive = ctrlName
If mStayOpen Then
glistN.NoCaretShow = True

If glistN.Visible Then glistN.PanPos = 0: glistN.PrepareToShow
Exit Sub
End If
If Not SkipProcAtExitFocus Then
If glistN.BackStyle <> 1 Then
glistN.ForeColor = white
glistN.backcolor = dimm2
End If
glistN.NoCaretShow = True: glistN.PanPos = 0: glistN.ShowMe
Else
'
End If
End If
' RaiseEvent LostFocus
End Sub

Private Sub glistN_HeaderSelected(Button As Integer)
' INFO SELECTED
Text = Text
End Sub

Private Sub glistN_KeyDown(KeyCode As Integer, shift As Integer)
Dim VR(2)
VR(0) = KeyCode
VR(1) = shift
   If mIndex >= 0 Then
   Callback.CallbackNow mesCtrlName + ".KeyDown(" + Str(mIndex) + ")", VR()
   Else
      Callback.CallbackNow mesCtrlName + ".KeyDown()", VR()
      End If
KeyCode = VR(0)
shift = VR(1)
If KeyCode = vbKeyF10 And shift <> 0 Then
KeyCode = 0
If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".PopUp(" + Str(mIndex) + "," + Str(MOUSEX(Callback.Left)) + "," + Str(MOUSEY(Callback.top)) + "," + Str(2) + "," + Str(glistN.ListIndex + 1) + ")"
   Else
      Callback.Callback mesCtrlName + ".PopUp(" + Str(MOUSEX(Callback.Left)) + "," + Str(MOUSEY(Callback.top)) + "," + Str(2) + "," + Str(glistN.ListIndex + 1) + ")"
      End If
Exit Sub
End If
If KeyCode = vbKeyUp Then
ElseIf KeyCode = vbKeyDown Then
ElseIf KeyCode = vbKeyV And (shift And 3) = 2 Then
If Not enabled Then Exit Sub
Pad$ = GetTextData(13)
If Pad$ = vbNullString Then Pad$ = Clipboard.GetText(1)
If Pad$ <> "" Then
If (glistN.maxchar - Len(mprompt)) > 0 Then
    vartext = Mid$(Pad$, 1, (glistN.maxchar - Len(mprompt)))

Else
    vartext = Pad$
    End If
End If
End If

End Sub

Public Property Get SelStart() As Long
SelStart = glistN.SelStart
End Property

Public Property Let SelStart(ByVal RHS As Long)
glistN.SelStart = RHS
End Property

Public Property Get Visible() As Boolean
Visible = glistN.Visible
End Property

Public Property Let Visible(ByVal RHS As Boolean)
glistN.Visible = RHS
End Property
Public Sub SetFocus()
On Error Resume Next
glistN.SetFocus
End Sub

Private Sub glistN_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
' RaiseEvent MouseMove(Button, shift, x, y)
End Sub

Private Sub glistN_Selected(item As Long)
If item >= 0 Then
   '' If (Not mCallback Is Nothing) Then mCallback.Press Index
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Click(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".Click()"
      End If
    End If
End Sub

Private Sub glistN_Selected2(item As Long)
If item >= 0 Then
  If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".Enter(" + Str(mIndex) + ")"
   Else
      Callback.Callback mesCtrlName + ".Enter()"
      End If
End If
End Sub


Private Sub glistN_SyncKeyboard(item As Integer)
' RaiseEvent KeyPress(item)
End Sub
Private Sub glistN_SpinnerValue(ThatValue As Long)
''' RaiseEvent SpinnerValue(ThatValue)
   If mIndex >= 0 Then
   Callback.Callback mesCtrlName + ".SpinnerValue(" + Str(mIndex) + "," + LTrim$(Str(ThatValue)) + ")"
   Else
      Callback.Callback mesCtrlName + ".SpinnerValue(" + LTrim$(Str(ThatValue)) + ")"
      End If
     '' vartext = str(ThatValue)
End Sub
Public Property Get locked() As Boolean
locked = Not glistN.enabled
End Property

Public Property Let locked(ByVal RHS As Boolean)
glistN.enabled = Not RHS
End Property
Public Sub Spinner(ShowBar As Boolean, MinValue As Long, maxvalue As Long, LargeValue As Long)
If Not ShowBar Then
glistN.Spinner = False
glistN.ShowBar = False
glistN.EnabledBar = False
Else
glistN.enabled = False
glistN.restrictLines = 1
glistN.Spinner = True
glistN.SetSpin MinValue, maxvalue, LargeValue
glistN.enabled = True
glistN.EnabledBar = True
glistN.ShowBar = True

End If
glistN.Vertical = True
End Sub


Public Property Get Value() As Long
If glistN.Spinner Then
Value = glistN.ValueSilent
End If

End Property

Public Property Let Value(ByVal RHS As Long)
If glistN.Spinner Then
glistN.ValueSilent = RHS
End If

End Property



Public Property Get MaxCharLength() As Long
MaxCharLength = glistN.maxchar
End Property

Public Property Let MaxCharLength(ByVal RHS As Long)
glistN.maxchar = RHS
End Property

Private Function Myleft$(a As Variant, pos As Variant)
Myleft$ = Mid$(a, 1, pos)
End Function
Private Function GetTextData(ByVal lFormatId As Long) As String
'' Addition to myTextBox for ctrl+V
Dim bData() As Byte, sr As String, sr1 As String
sr1 = Clipboard.GetText(1)
If (OpenClipboard(0) <> 0) Then

        
        If (GetBinaryData(lFormatId, bData())) Then
        sr = bData
    
     If IsWine Then
            sr1 = Myleft$(sr, Len(sr1))
            GetTextData = Myleft$(sr1, Len(sr1))
    Else
                GetTextData = Myleft$(sr, Len(sr1))
    End If
    
          
        End If

End If
CloseClipboard
End Function
Private Function GetClipboardMemoryHandle( _
        ByVal lFormatId As Long _
    ) As Long

    

    If (IsClipboardFormatAvailable(lFormatId) <> 0) Then
         
        GetClipboardMemoryHandle = GetClipboardData(lFormatId)
        
    End If
End Function
Private Function GetBinaryData( _
        ByVal lFormatId As Long, _
        ByRef bData() As Byte _
    ) As Boolean

Dim hMem As Long, lSize As Long, lPtr As Long
    

    Erase bData
    
    hMem = GetClipboardMemoryHandle(lFormatId)

    If (hMem <> 0) Then

        lSize = GlobalSize(hMem)

        lPtr = GlobalLock(hMem)
        If (lSize > 0) Then

            ReDim bData(0 To lSize - 2) As Byte

            CopyMemory bData(0), ByVal lPtr, lSize - 1
        End If

        GlobalUnlock hMem

        GetBinaryData = (lSize > 0)

    End If
End Function

Private Sub glistN_UnregisterGlist()
If Not glistN.Spinner Then glistN.NoWheel = True
If glistN.TabStopSoft Then Callback.LastActive = ctrlName
On Error Resume Next
Set LastGlist = Nothing
If Err.Number > 0 Then glistN.NoWheel = True

End Sub

Public Sub Move(X, Y, Optional w, Optional h)
If IsMissing(w) Then
glistN.MoveTwips CLng(X), CLng(Y), 0, 0
ElseIf IsMissing(h) Then
glistN.MoveTwips CLng(X), CLng(Y), CLng(w), 0
Else
glistN.MoveTwips CLng(X), CLng(Y), CLng(w), CLng(h)
End If
End Sub
Public Sub deconstruct()

If Not glistN Is Nothing Then glistN.Shutdown
Set Callback = Nothing
Set glistN = Nothing
End Sub
Public Sub FontAttr(ThisFontName, Optional ThisMode = -1, Optional ThisBold = True)
Dim aa As New StdFont
If ThisFontName <> "" Then

aa.name = ThisFontName

If ThisMode > 7 Then aa.Size = ThisMode Else aa = 7
aa.bold = ThisBold
Set glistN.Font = aa
End If
End Sub
Public Property Get index() As Long
index = mIndex
End Property

Public Property Let index(ByVal RHS As Long)
mIndex = RHS
End Property
Public Property Let ShowAlways(ByVal vNewValue As Variant)
mStayOpen = CBool(vNewValue)
If mStayOpen Then
glistN.ForeColor = rgb(0, 0, 0)
glistN.backcolor = rgb(255, 255, 255)
End If
End Property
Sub GetFocus()
On Error Resume Next
Callback.Controls(ctrlName).SetFocus
End Sub
Public Property Let Default(RHS)
On Error Resume Next
If CBool(RHS) Then Callback.Default = ctrlName
End Property
Sub Transparent()
Dim X  As Long, Y As Long
glistN.FadeLastLinePart = mPartColor
glistN.ForeColor = black
glistN.BackStyle = 1
glistN.GetLeftTop X, Y
glistN.RepaintFromOut Callback.Image, X, Y
Callback.RefreshList = Callback.RefreshList + 1
End Sub
Sub TransparentOff()
On Error Resume Next
If Not glistN Is Nothing Then
glistN.BackStyle = 0
glistN.backcolor = mycolor(dimm2)
glistN.PanPos = 0
glistN.ShowMe
Callback.RefreshList = Callback.RefreshList - 1
End If
End Sub
Sub ColorsReset()
white = rgb(255, 255, 255)
black = 0
dimm1 = rgb(100, 200, 160)
dimm2 = rgb(200, 120, 60)
End Sub
Sub Colors(white1, Optional black1, Optional dimm11, Optional dimm21)

white = mycolor(white1)
If Not IsMissing(black1) Then
black = mycolor(black1)
If Not IsMissing(dimm11) Then
dimm1 = dimm11
If Not IsMissing(dimm21) Then dimm2 = dimm21
End If
End If
If enabled Then
glistN.ForeColor = white
Else
glistN.ForeColor = black
End If

'Retired
End Sub

Public Property Get FadePartColor() As Variant
FadePartColor = mPartColor
End Property

Public Property Let FadePartColor(ByVal vNewValue As Variant)
On Error Resume Next
mPartColor = CLng(vNewValue)
End Property

Public Property Let icon(RHS)
glistN.icon = RHS

End Property
Public Property Let Drag(RHS)
glistN.DragEnabled = RHS
End Property
Public Property Get Drag()
Drag = glistN.DragEnabled
End Property
Public Property Let drop(RHS)
glistN.DropEnabled = RHS
End Property
Public Property Get drop()
drop = glistN.DropEnabled
End Property
Private Sub glistN_AddSelStart(val As Long)
Dim w As Long, a$
If glistN.SelStart > 0 Then
a$ = Mid$(vartext, glistN.SelStart, 1)
If Len(a$) > 0 Then
w = AscW(a$)
If w > -10241 And w < -9984 Then
val = 2
End If
End If
End If
End Sub

Private Sub glistN_SubSelStart(val As Long)
Dim w As Long, a$
If glistN.SelStart < 2 Then Exit Sub
a$ = Mid$(vartext, glistN.SelStart - 1, 1)
If Len(a$) > 0 Then
w = AscW(a$)
If w > -9215 And w < -8960 Then

val = 2

End If
End If
End Sub
