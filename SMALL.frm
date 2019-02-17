VERSION 5.00
Begin VB.Form Form3 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H000080FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "M2000"
   ClientHeight    =   570
   ClientLeft      =   -47955
   ClientTop       =   48315
   ClientWidth     =   1365
   Icon            =   "SMALL.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   1365
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   510
      Top             =   165
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum BOOL
    FALSE 
    TRUE 
End Enum
#If False Then
    Dim FALSE , TRUE 
#End If
Public hideme As Boolean
Private foundform5 As Boolean


Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function GetModuleHandleW Lib "KERNEL32" (ByVal lpModuleName As Long) As Long


Private Declare Function GetProcAddress Lib "KERNEL32" (ByVal hModule As Long, ByVal lpProcName As String) As Long


Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long


Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long


Private Declare Function SetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long) As Long
    Private Const GWL_WNDPROC = -4
    Private m_Caption As String

'Public lastform As Form
Private affiliatehwnd As Long
Public skiptimer As Boolean
Private Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub PutMem4 Lib "msvbvm60" (Destination As Any, Value As Any)
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal OleStr As Long, ByVal BLen As Long) As Long

Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const WM_SETTEXT = &HC

Public Property Get CaptionW() As String
    If m_Caption = "M2000" Then
        CaptionW = vbNullString
    Else
        CaptionW = m_Caption
    End If
End Property


Public Property Let CaptionW(ByVal NewValue As String)

If LenB(NewValue) = 0 Then NewValue = "M2000"
    m_Caption = NewValue
DefWindowProcW Me.hWnd, WM_SETTEXT, 0, ByVal StrPtr(NewValue)
  
    If lastform Is Nothing Then
    Show
    MyDoEvents1 Me, True
   End If

     
End Property
Public Property Let CaptionWsilent(ByVal NewValue As String)

If LenB(NewValue) = 0 Then NewValue = "M2000"
    m_Caption = NewValue
DefWindowProcW Me.hWnd, WM_SETTEXT, 0, ByVal StrPtr(NewValue)
End Property
Public Function ask(bstack As basetask, a$) As Double
If ASKINUSE Then Exit Function
DialogSetupLang DialogLang
AskText$ = a$
ask = NeoASK(bstack)

End Function
Public Function NeoASK(bstack As basetask) As Double
If ASKINUSE Then Exit Function
Dim safety As Long
Dim oldesc As Boolean, zz As Form
    oldesc = escok
'using AskTitle$, AskText$, AskCancel$, AskOk$, AskDIB$
Static once As Boolean
If once Then Exit Function
once = True
ASKINUSE = True
If Not Screen.ActiveForm Is Nothing Then
If TypeOf Screen.ActiveForm Is GuiM2000 Then Screen.ActiveForm.UNhookMe
Set zz = Screen.ActiveForm
End If
Dim INFOONLY As Boolean
k1 = 0
If AskTitle$ = vbNullString Then AskTitle$ = MesTitle$
If AskCancel$ = vbNullString Then INFOONLY = True
If AskOk$ = vbNullString Then AskOk$ = "OK"


If Not Screen.ActiveForm Is Nothing Then
If Screen.ActiveForm Is MyPopUp Then
   If MyPopUp.LASTActiveForm Is Form1 Then
        NeoMsgBox.Show , Form1
        MoveFormToOtherMonitorOnly NeoMsgBox, False
   ElseIf MyPopUp.LASTActiveForm Is Nothing Then
     NeoMsgBox.Show , MyPopUp.LASTActiveForm
     MoveFormToOtherMonitorOnly NeoMsgBox, True
   End If
ElseIf Screen.ActiveForm Is Form1 Then
NeoMsgBox.Show , Screen.ActiveForm
MoveFormToOtherMonitorOnly NeoMsgBox, False
Else
NeoMsgBox.Show , Screen.ActiveForm
MoveFormToOtherMonitorOnly NeoMsgBox, True
End If
ElseIf form5iamloaded Then
MyDoEvents1 Form5
Sleep 1
NeoMsgBox.Show , Form5
MoveFormToOtherMonitorCenter NeoMsgBox
Else
NeoMsgBox.Show
MoveFormToOtherMonitorCenter NeoMsgBox
End If
'End If
On Error Resume Next
''SleepWait3 10
Sleep 1
If Form1.Visible Then
Form1.Refresh
ElseIf form5iamloaded Then
Form5.Refresh
Else
MyDoEvents
End If
Sleep 1
safety = uintnew(timeGetTime) + 30
While Not NeoMsgBox.Visible And safety < uintnew(timeGetTime)
    MyDoEvents
Wend
If NeoMsgBox.Visible = False Then
    MyEr "can't open msgbox", "äåí ìðïñþ íá áíïßîù ôïí äéÜëïãï"
    GoTo conthere
    Exit Function
End If

If AskInput Then
NeoMsgBox.gList3.SetFocus
End If
  If bstack.ThreadsNumber = 0 Then
    On Error Resume Next
    If Not (bstack.toback Or bstack.toprinter) Then If bstack.Owner.Visible Then bstack.Owner.Refresh
    End If
    If Not NeoMsgBox.Visible Then
    NeoMsgBox.Visible = True
    MyDoEvents
    End If
    Dim mycode As Double, oldcodeid As Double, x As Form
mycode = Rnd * 12312314
oldcodeid = Modalid

 For Each x In Forms
                            If x.Visible And x.name = "GuiM2000" Then
                     
                           If x.Enablecontrol Then
                               x.Modal = mycode
                                x.Enablecontrol = False
                            End If
                            End If
                    Next x
                     Set x = Nothing
If INFOONLY Then
NeoMsgBox.command1(0).SetFocus
End If
Modalid = mycode

Do
If TaskMaster Is Nothing Then
        mywaitOld bstack, 5
      Sleep 1
      Else
    
      If Not TaskMaster.Processing And TaskMaster.QueueCount = 0 Then
        DoEvents
      Else
       TaskMaster.TimerTickNow
       TaskMaster.StopProcess
       
       DoEvents
       TaskMaster.StartProcess
       End If
      End If
Loop Until NOEXECUTION Or Not ASKINUSE
 Modalid = mycode
k1 = 0
 BLOCKkey = True
While KeyPressed(&H1B) ''And UseEsc

ProcTask2 bstack
NOEXECUTION = False
Wend
BLOCKkey = False
AskTitle$ = vbNullString
Dim z As Form
 Set z = Nothing

           For Each x In Forms
            If x.Visible And x.name = "GuiM2000" Then
            If Not x.Enablecontrol Then x.TestModal mycode
          If x.Enablecontrol Then Set z = x
            End If
            Next x
             Set x = Nothing
          If Not zz Is Nothing Then Set z = zz
          
          If Typename(z) = "GuiM2000" Then
            z.ShowmeALL
            z.SetFocus
            Set z = Nothing
            ElseIf Not z Is Nothing Then
            If z.Visible Then z.SetFocus
          End If
          Modalid = oldcodeid
          
If INFOONLY Then
NeoASK = 1
Else
NeoASK = Abs(AskCancel$ = vbNullString) + 1
End If
If NeoASK = 1 Then
If AskInput Then
bstack.soros.PushStr AskStrInput$
End If
End If
conthere:
AskCancel$ = vbNullString
once = False
ASKINUSE = False
INK$ = vbNullString
escok = oldesc
End Function

Sub mywait(bstack As basetask, PP As Double, Optional SLEEPSHORT As Boolean = False)
Dim p As Boolean, e As Boolean
On Error Resume Next
If bstack.Process Is Nothing Then
''If extreme Then MyDoEvents1 Form1
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





        If Form1.DIS.Visible And Not bstack.toprinter Then
        MyDoEvents0 Form1.DIS
   
        Else
        MyDoEvents0 Me
        End If
If SLEEPSHORT Then Sleep 1
If e Then
p = bstack.Process.Done
If Err.Number = 0 Then
If p <> 0 Then
Exit Do
End If
End If
End If
Loop Until PP <= CDbl(timeGetTime) Or NOEXECUTION Or MOUT

                       If exWnd <> 0 Then
                MyTitle$ bstack
                End If
End Sub

Private Sub mywaitOld(bstack As basetask, PP As Double)
Dim p As Boolean, e As Boolean

On Error Resume Next
If bstack.ThreadsNumber = 0 Then GoTo cont1
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
cont1:
PP = PP + CDbl(timeGetTime)

Do

If Not TaskMaster Is Nothing Then
If TaskMaster.Processing And Not bstack.TaskMain Then
        If Not bstack.toprinter Then bstack.Owner.Refresh
        TaskMaster.TimerTick
       ' SleepWait 1
       TaskMaster.StopProcess
       DoEvents
       TaskMaster.StartProcess
       
Else
        ' SleepWait 1
        
        MyDoEvents
        End If
        Else
        DoEvents
        End If
If e Then
p = bstack.Process.Done
If Err.Number = 0 Then
If p <> 0 Then
Exit Do
End If
End If
End If
Loop Until PP <= CDbl(timeGetTime) Or NOEXECUTION Or MOUT

                       If exWnd <> 0 Then
                MyTitle$ bstack
                End If
            
End Sub




Private Sub Form_DblClick()
On Error Resume Next
If lastform Is Nothing Then Exit Sub
If WindowState = 0 Then lastform.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
If QRY Or GFQRY Then
If Form1.Visible Then Form1.SetFocus
ElseIf KeyCode = 27 And ASKINUSE Then

    NOEXECUTION = True
Else
choosenext
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If QRY Or GFQRY Then
If Form1.Visible Then Form1.SetFocus
Else

End If
If Not BLOCKkey Then INK$ = INK$ & Chr(KeyAscii)
End Sub

Private Sub Form_Load()
Debug.Assert (InIDECheck = True)
Timer1.Interval = 10000
Timer1.enabled = False
If Not byPassCallback Then Set Me.icon = Form1.icon
End Sub




Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim nook As Long, uselastform As Boolean
Timer1.enabled = False
If UnloadMode = (vbFormControlMenu Or byPassCallback) And lastform Is Nothing Then

Dim F As Form, i As Long, F1 As GuiM2000
For i = Forms.count - 1 To 0 Step -1
Set F = Forms(i)
If TypeOf F Is GuiM2000 Then
    Set F1 = F
    
    If Modalid <> 0 Then
        If F1.Modal = Modalid Then
            If F1.MyName$ <> "" Then
                If lastform Is F1 Then
                    If Not F1.Enablecontrol Then Cancel = True: Exit Sub
                    F1.ByeBye2 nook: uselastform = True
                Else
                    F1.ByeBye2 nook
                End If
            End If
        End If
    Else
        If F1.MyName$ <> "" Then
        If lastform Is F1 Then
            If Not F1.Enablecontrol Then Cancel = True: Exit Sub
                F1.ByeBye2 nook: uselastform = True
            Else
                F1.ByeBye2 nook
            End If
        End If
    End If

End If

If nook Then Cancel = True: Exit Sub
Next i
Set F = Nothing

If Not lastform Is Nothing Then
Exit Sub
End If

        If exWnd <> 0 Then
        Form1.IEUP ("")
        Cancel = True
        Exit Sub
        End If
   Timer1.enabled = False
    NOEXECUTION = True
    ExTarget = True
    INK$ = Chr(27)
    If Not TaskMaster Is Nothing Then
    TaskMaster.Dispose
    End If
    NOEDIT = True
    MOUT = True
    Cancel = Not byPassCallback
Else

If lastform Is Nothing Then ttl = False: Exit Sub
If UnloadMode = vbFormControlMenu Then
Set F1 = lastform
F1.ByeBye
Cancel = True
End If
End If


End Sub

Private Sub Form_Resize()
Static once As BOOL
If Timer1.enabled Then Exit Sub
If once Then Exit Sub
once = TRUE 
If Not lastform Is Nothing Then
If lastform.WindowState = 1 Then
lastform.WindowState = 0: Me.skiptimer = True: Me.WindowState = 0: once = False: Exit Sub
End If
End If

 hideme = (Me.WindowState = 1)

  If hideme Then
    reopen2 = False
    reopen4 = False
    If Form4.Visible Then Form4.Visible = False: reopen4 = True
    If Form2.Visible Then If trace Then Form2.Visible = False: reopen2 = True
     Timer1_Timer
     once = False
    Exit Sub
    
 ElseIf Forms.count > 4 Then
 If Not lastform Is Nothing Then
 Timer1.enabled = True
     ElseIf Not (CaptionW <> "" And WindowState = 0) Then
     If Not Form1.TrueVisible Then once = False: Exit Sub
    End If
     ElseIf WindowState = 0 Then
     
     Refresh
    
 End If
If Not lastform Is Nothing Then
Timer1.Interval = 30
Else
 Timer1.enabled = Timer1.Interval < 10000
 End If
 once = False
End Sub



Private Sub Form_Unload(Cancel As Integer)
Set lastform = Nothing
End Sub

Private Sub Timer1_Timer()
' On Error Resume Next
Static once As Boolean
If once Then Exit Sub
once = True
Dim F As Form, F1 As GuiM2000, i As Long, thismodal As Double, f2 As GuiM2000
If DIALOGSHOW Or ASKINUSE Or skiptimer Then
skiptimer = False
Timer1.enabled = False
once = False
Exit Sub
End If
Timer1.enabled = False
Timer1.Interval = 20
If Not lastform Is Nothing Then
If TypeOf lastform Is GuiM2000 Then
If Not hideme Then
    Set F1 = lastform
    If F1.NeverShow Then
starthere:
        Set f2 = F1
        If F1.Modal <> 0 Then
            thismodal = F1.Modal
            Set f2 = F1
            If F1.Enablecontrol Then
            'we have the top
                For i = 0 To Forms.count - 1
                    If TypeOf Forms(i) Is GuiM2000 Then
                        Set F1 = Forms(i)
                        If Not f2 Is F1 Then
                            F1.Visible = F1.VisibleOldState Or F1.Visible
                            F1.VisibleOldState = False
                            F1.MinimizeOff
                            If F1.Visible Then
                                
                                If Form1.Visible Then
                                    F1.Show , Form1
                                Else
                                    F1.Show
                                End If
                            End If
                        End If
                    End If
                Next i
                Set F1 = f2
                F1.Visible = F1.VisibleOldState Or F1.Visible
                F1.VisibleOldState = False
                If F1.Visible Then
                    F1.MinimizeOff
                    If Form1.Visible Then
                        F1.Show , Form1
                    Else
                        F1.Show
                    End If
                End If
                Set F1 = Nothing
                Set f2 = Nothing
            Else
            ' we have something else
                
                For i = 0 To Forms.count - 1
                    If TypeOf Forms(i) Is GuiM2000 Then
                        Set F1 = Forms(i)
                        If F1.Enablecontrol Then
                        ' we found the top
                            GoTo starthere
                        End If
                    End If
                Next i
                ' nothing found something wrong
                ' do nothing
                once = False
                Exit Sub
            End If
        Else
            F1.Visible = F1.VisibleOldState Or F1.Visible
            F1.VisibleOldState = False
            If F1.Visible Then
                If Form1.Visible Then
                    F1.Show , Form1
                Else
                    F1.Show
                End If
            End If
        End If
    End If
Else
        Set F1 = lastform
        F1.VisibleOldState = F1.Visible
        F1.Visible = False
        If F1.Modal <> 0 Then
        thismodal = F1.Modal
        For i = Forms.count - 1 To 0 Step -1
            If TypeOf Forms(i) Is GuiM2000 Then
                Set F1 = Forms(i)
                If Not F1.Minimized Then
                    If F1.Modal = thismodal Then
                        F1.VisibleOldState = True
                        F1.Visible = False
                        F1.MinimizeON
                    End If
                End If
            End If
        Next i
        Set F1 = Nothing
        End If
End If
End If
once = False
Exit Sub
ElseIf Not hideme Then
If Not (Form1.TrueVisible Or Form1.Visible) Then
If foundform5 Then
Form5.Visible = True
'DoEvents
End If
Form1.Visible = Form1.Visible And Form1.TrueVisible
If Not ttl Then
ttl = True ' for form1

'Form1.Top = ScrInfo(Console).Top
If Not Form1.Visible Then
If Not IsSelectorInUse Then Form1.Show , Form5
Else
If Not IsSelectorInUse Then Form1.Show , Form5
End If
End If
'DoEvents
Else
If foundform5 Then
Form5.Visible = True
'DoEvents
End If
End If

'Sleep 500
Form1.Visible = Form1.Visible Or Form1.TrueVisible
If Form1.Visible And Not IsSelectorInUse Then
If Not trace Then reopen2 = False
If vH_title$ = vbNullString Then reopen4 = False
If reopen4 Then Form4.Show , Form1: Form4.Visible = True
If reopen2 Then Form2.Show , Form1: Form2.Visible = True
   
   
   
   
 
       
Sleep 1

If Form1.Visible Then Form1.SetFocus:  Form1.KeyPreview = True

If Form1.Visible Then Sleep 2
 Set F = Nothing
 Else
  Form1.Visible = Form1.TrueVisible
 Form1.Visible = False

    For Each F In Forms
       If Typename$(F) = "GuiM2000" Then
    Set F1 = F
        If F1.NeverShow Then
    If F1 Is lastform Then
        F1.Visible = F1.VisibleOldState Or F1.Visible
        F1.VisibleOldState = False
    End If
    End If
      
        
       If F1.Visible Then
       If Form1.Visible Then
       F1.Show , Form1
       Else
       F1.Show
       End If
       End If
 
       End If
       Next

End If
Else

If Not ((exWnd <> 0) Or AVIRUN Or IsSelectorInUse) Then
'

Form1.TrueVisible = Form1.Visible
Form1.Visible = False
'Form1.Hide
If Form5.Visible Then Form5.Visible = False: foundform5 = True
End If
 For Each F In Forms
        If TypeOf F Is GuiM2000 Then
            Set F1 = F
            If F1.TrueVisible Then
                F1.VisibleOldState = True
                F1.Visible = False
                F1.MinimizeON
            End If
        End If
       Next
End If
End Sub

Public Function InIDECheck() As Boolean
    m_bInIDE = True
    InIDECheck = True
End Function

Public Property Get lastform() As Form
If affiliatehwnd = 0 Then
    Set lastform = Nothing
Else
Dim i As Long
    For i = 1 To Forms.count - 1
        If Forms(i).hWnd = affiliatehwnd Then Set lastform = Forms(i)
    Next i
End If
End Property

Public Property Set lastform(vNewValue As Form)
On Error Resume Next
Dim i
For i = 1 To Forms.count - 1
If Forms(i) Is vNewValue Then affiliatehwnd = Forms(i).hWnd: Exit Property
Next i
affiliatehwnd = 0
End Property
