VERSION 5.00
Begin VB.Form AVI 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1065
   ClientLeft      =   6285
   ClientTop       =   -8790
   ClientWidth     =   1380
   Icon            =   "Avi.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1065
   ScaleWidth      =   1380
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   285
      Top             =   375
   End
End
Attribute VB_Name = "AVI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim doubleclick As Long
Dim ERNUM As Long
Private mInterval As Long
Private intTimes As Long

Private Sub Form_Activate()

CLICK_COUNT = CLICK_COUNT + 1
If getout Then Exit Sub
If CLICK_COUNT < 2 Then


If Height > 30 Then
If Not UseAviXY Then
SetWindowPos Me.hWnd, HWND_TOPMOST, Me.Left / dv15, _
                        Me.top / dv15, Me.Width / dv15, _
                        Me.Height / dv15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
Else


SetWindowPos Me.hWnd, HWND_TOPMOST, aviX / dv15, _
                        aviY / dv15, Me.Width / dv15, _
                        Me.Height / dv15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
End If
End If
ElseIf AVIRUN Then
getout = True
Else
End If
If getout Then
GETLOST
End If
End Sub

Private Sub Form_Click()
If MediaPlayer1.isMoviePlaying Then GETLOST
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, shift As Integer)
KeyCode = 0
If Form1.Visible Then
If Form1.TEXT1.Visible Then
Form1.TEXT1.SetFocus
Else
Form1.SetFocus
End If
End If

End Sub

Private Sub Form_Load()
Me.backcolor = Form1.DIS.backcolor
If getout Then Exit Sub
CLICK_COUNT = 0
Dim CD As String
getout = False
On Error Resume Next
If avifile = vbNullString Then
GETLOST
getout = True
Else
MediaPlayer1.hideMovie
MediaPlayer1.FileName = avifile
Timer1.enabled = False
Interval = MediaPlayer1.Length


If MediaPlayer1.Error = 0 Then

If UseAviSize And MediaPlayer1.Height > 2 Then
If AviSizeX = 0 And AviSizeY = 0 Then
If aviX = 0 And aviY = 0 Then

AviSizeX = (ScrInfo(Console).Width - 1) * 0.99
AviSizeY = (ScrInfo(Console).Height - 1) * 0.99
Else
AviSizeX = Me.Scalewidth
AviSizeY = Me.Scaleheight
End If
Else
If AviSizeX = 0 And AviSizeY <> 0 Then
AviSizeX = CLng(AviSizeY * MediaPlayer1.Width / CDbl(MediaPlayer1.Height))
End If
If AviSizeY = 0 Then
AviSizeY = CLng(AviSizeX * MediaPlayer1.Height / CDbl(MediaPlayer1.Width))
End If
If aviX = 0 And aviY = 0 Then

aviX = ((ScrInfo(Console).Width - 1) * 0.99 - AviSizeX) / 2 + ScrInfo(Console).Left
aviY = ((ScrInfo(Console).Height - 1) * 0.99 - AviSizeY) / 2 + ScrInfo(Console).top
End If
End If

Me.move aviX, aviY, AviSizeX, AviSizeY
MyDoEvents
MediaPlayer1.openMovieWindow Me.hWnd, "child"

MediaPlayer1.sizeLocateMovie 0, 0, ScaleX(AviSizeX, vbTwips, vbPixels), ScaleY(AviSizeY, vbTwips, vbPixels) + 1
'Show
ElseIf MediaPlayer1.Height > 2 Then
Me.move Left, top, ScaleX(MediaPlayer1.Width, vbPixels, vbTwips), ScaleY(MediaPlayer1.Height, vbPixels, vbTwips) + 1

MediaPlayer1.openMovieWindow AVI.hWnd, "child"




Else
Me.move Left, top, ScaleX(c&, vbPixels, vbTwips), ScaleY(MediaPlayer1.Height, vbPixels, vbTwips)

MediaPlayer1.minimizeMovie
MediaPlayer1.openMovie

End If


If MediaPlayer1.Height <= 2 Then
Width = 0
Height = 0
Else

If Not UseAviXY Then
Me.move ((ScrInfo(Console).Width - 1) - Width) / 2, ((ScrInfo(Console).Height - 1) - Height) / 2
End If
End If


Timer1.enabled = False



Else
getout = True
Width = 0
Height = 0
'Show
End If
End If
AVIUP = True
End Sub

Public Sub Avi2Up()
Timer1.enabled = True
Me.ZOrder
MediaPlayer1.playMovie

Timer1.enabled = True
AVIRUN = True
End Sub



Private Sub Form_Unload(Cancel As Integer)
getout = False
AVIRUN = False
AVIUP = False
End Sub



Public Sub GETLOST()
getout = True
Timer1.enabled = False
Hide
MediaPlayer1.hideMovie
MediaPlayer1.stopMovie
MediaPlayer1.closeMovie
AVIRUN = False
MyDoEvents
If Form1.Visible Then Form1.Refresh: If Form1.DIS.Visible Then Form1.DIS.Refresh
Unload Me
End Sub


Private Sub Frame1_Click()
GETLOST
End Sub

Private Sub Timer1_Timer()
If intTimes > 0 Then
    intimes = intTimes - 1
    Timer1.Interval = 32767#
Else
    GETLOST
End If

End Sub
Public Property Get Interval() As Long
Interval = mInterval
End Property

Public Property Let Interval(aa As Long)
Dim Comp As Long
mInterval = aa
Comp = CLng(aa \ 32768#)
If Comp = 0 Then
    Timer1.Interval = aa
    intTimes = 0
Else
    Timer1.Interval = aa - CLng(Comp * 32768#)
    intTimes = Comp
End If
End Property
