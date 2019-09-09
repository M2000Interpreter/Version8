VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RecordMci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long 'Get the error message of the mcidevice if any
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long 'Send command strings to the mci device
Private data As String * 128 ' Used to store our return data
Public error As Long ' Used to store our error message
Public Height As Long
Public Width As Long
Public Length As Long
Private bt As Long
Private ch As Long
Private Terminate As String
 Private mFilename  As String ' Used to store our file
 Public Property Let FileName(n As String)
mFilename = n
End Property
Public Property Get FileName() As String
FileName = mFilename
End Property
Public Sub Rec_Initialize()
If Terminate <> "" Then
error = mciSendString(Terminate, 0, 0, 0)
End If
error = mciSendString("open new type waveaudio alias capture", 0&, 0, 0)
error = mciSendString("set capture time format ms", 0, 0, 0)
If error <= 0 Then Terminate = "close capture"
End Sub
Public Sub GeneralMCI(OpenString As String, CloseString As String)
error = mciSendString(OpenString, 0&, 0, 0)
If error = 0 Then Terminate = CloseString
End Sub
Public Sub oneMCI(NowString As String)
error = mciSendString(NowString, 0&, 0, 0)
End Sub
Private Function NeedByteFactor() As Long
If bt = 0 Then Bit8
If ch = 0 Then Mono
NeedByteFactor = bt * ch
If NeedByteFactor = 0 Then
NeedByteFactor = 1
End If
End Function

Public Sub Bit8()

 error = mciSendString("set capture bitspersample 8", 0&, 0, 0)
 If error <= 0 Then bt = 1
End Sub
Public Sub Bit16()
 error = mciSendString("set capture bitspersample 16", 0&, 0, 0)
  If error <= 0 Then bt = 2
End Sub
Public Sub QualityRadio()

  error = mciSendString("set capture samplespersec 11025 BYTESPERSEC " & CStr(11025 * NeedByteFactor), 0&, 0, 0)
  
End Sub
Public Sub QualityAny(qrec As Double)
 error = mciSendString("set capture samplespersec " & CStr(CLng(qrec)) & " BYTESPERSEC " & CStr(CLng(qrec) * NeedByteFactor), 0&, 0, 0)
End Sub
Public Sub Mono()
error = mciSendString("set capture channels 1", 0&, 0, 0)
 If error <= 0 Then ch = 1
End Sub
Public Sub Stereo()
error = mciSendString("set capture channels 2", 0&, 0, 0)
If error <= 0 Then ch = 2
End Sub
Public Sub RecFast()
Bit8
QualityRadio
Mono
Capture
' stop
' save and play
' or play
End Sub
Public Sub Capture(Optional ByVal ins As Boolean = False)
If ins Then
error = mciSendString("record capture insert", 0&, 0, 0)
Else
error = mciSendString("record capture", 0&, 0, 0)
End If
End Sub
Public Sub ReCapture()
error = mciSendString("record capture overwrite", 0&, 0, 0)
End Sub
Sub CutRecordMs(recStart As Double, recEnd As Double)
error = mciSendString("set capture time format ms", 0, 0, 0)
error = mciSendString("delete capture from " & CStr(recStart) & " to " & CStr(recEnd), 0&, 0, 0)

End Sub
Public Function getPositionInMS() As Long
    'Get the position of the movie in milliseconds
    error = mciSendString("set capture time format ms", 0, 0, 0)
    error = mciSendString("status capture position", data, 128, 0)
    getPositionInMS = val(data)
End Function
Public Function getLengthInMS() As Long
    'Get the position of the movie in milliseconds
    error = mciSendString("set capture time format ms", 0, 0, 0)
    error = mciSendString("status capture length", data, 128, 0)
    getLengthInMS = val(data)
End Function
Sub recStop()
error = mciSendString("stop capture", 0, 0, 0)
End Sub
Sub recPlay()
error = mciSendString("play capture", 0, 0, 0)
End Sub
Sub recPlayFromMs(recPos As Double)
recPos = CLng(recPos)
error = mciSendString("set capture time format ms", 0, 0, 0)
error = mciSendString("play capture from " & CStr(recPos), 0, 0, 0)
End Sub
Public Function setVolume(Value As Long)
 error = mciSendString("setaudio capture volume to " & CStr(Value), 0, 0, 0)
End Function
Public Function isRecPlaying() As Boolean
    Dim isPlaying As String
    error = mciSendString("status capture mode", data, 128, 0)
    MyDoEvents
    isPlaying = Left(data, 7)
    If isPlaying = "playing" And Not NOEXECUTION Then
        isRecPlaying = True
    Else
        isRecPlaying = False
    End If
End Function
Sub Save()
If mFilename <> "" Then
error = mciSendString("save capture " & """" + mFilename + """", 0&, 0, 0)
End If
End Sub
Sub AnyMci(CommandString As String)
error = mciSendString(CommandString, 0&, 0, 0)

End Sub
Function AnyMciValue(CommandString As String)
 error = mciSendString(CommandString, data, 128, 0)
    AnyMciValue = val(data)
End Function
Public Function checkError() As String
    'A very useful function for getting any errors
    'associated with the mci device
    checkError = space$(255)
    mciGetErrorString error, checkError, Len(checkError)
End Function
Sub SaveAs(fnm As String)

error = mciSendString("save capture " & """" + fnm + """", 0&, 0, 0)
If error <= 0 Then
mFilename = fnm
End If

End Sub

Private Sub Class_Terminate()
If Terminate <> "" Then
error = mciSendString(Terminate, 0, 0, 0)

End If
End Sub
