VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' M2000 ver 1.
Option Explicit
' chr(34) is """" maybe somewhere in a module as public
Const mway As Long = DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS Or DT_TABSTOP
Private col() As Variant, lastdata As String, thispara As Long
Private TextPixelOffset As Long
Public ColorSet As Long
Public NoCenterLineEdit As Boolean
Public NumberOnly As Boolean
Public NumberIntOnly As Boolean
Public WithEvents glistN As gList
Attribute glistN.VB_VarHelpID = -1
Private blockKeyboard As Boolean
Public WithEvents mDoc As Document    ' document
Attribute mDoc.VB_VarHelpID = -1
Public Form1mn1Enabled As Boolean
Public Form1mn2Enabled As Boolean
Public Form1mn3Enabled As Boolean
Public Form1sdnEnabled As Boolean
Public Form1supEnabled As Boolean
Public Form1mscatEnabled As Boolean
Public Form1rthisEnabled As Boolean
Dim WithEvents mDoc1 As Document  ' header
Attribute mDoc1.VB_VarHelpID = -1
Dim UndoDoc As New Document  ' one level of undo no events ' no break lines
Dim undopara As Long, undopos As Long, undotext As String, undomode As Long, undotexthtml As String
Dim Redopara As Long, Redopos As Long, Redotext As String, Redomode As Long
Dim undostate As Boolean
Dim Redostate As Boolean
Dim NoHeadwrap As Boolean
Private Type UndoChain
    undopara As Long
    undopos As Long
    undomode As Long
    undotext As String
End Type
Dim NeoUndo As Long, UndoMax As Long
Dim NeoRedo As Long, RedoMax As Long

Dim MultiUndo() As UndoChain
Dim MultiRedo() As UndoChain

Dim GetUndoNext As Boolean
Dim myfilename As String
Dim all As Long, onelineheight As Long
Event inform(tLine As Long, tPos As Long)
Event ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
Event CtrlPlusF1()
Public LastSelStart As Long
Public NoMark As Boolean
Public NoTitle As Boolean
Public NoColor As Boolean
Public Title As String
Public UsedAsTextBox As Boolean  '' only in m2000. Enter exit - use maxchar - No wrap
Private mEditDoc As Boolean
Public AutoNumber As Boolean
Public Charpos As Long
Public showparagraph As Boolean
Private showparagraphWork As Boolean
Public SelectionColor As Long
Public SelectionTextColor As Long
Private showparagraphonce As Boolean
Private mNoWrap As Boolean
Private BlockStartPara As Long, BlockStartPos As Long, BlockStartLine As Long, BlockTopline As Long
Private BlockStartCursor As Long, BlockEndCursor As Long
Private BlockEndPara As Long, BlockEndPos As Long, BlockEndline As Long, BlockBottomline As Long
Private Declare Function CopyFromLParamToRect Lib "user32" Alias "CopyRect" (lpDestRect As RECT, ByVal lpSourceRect As Long) As Long
Private Declare Function GlobalAlloc Lib "KERNEL32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "KERNEL32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Const GMEM_DDESHARE = &H2000&
Private Const GMEM_DISCARDABLE = &H100&
Private Const GMEM_DISCARDED = &H4000&
Private Const GMEM_FIXED = &H0&
Private Const GMEM_INVALID_HANDLE = &H8000&
Private Const GMEM_LOCKCOUNT = &HFF&
Private Const GMEM_MODIFY = &H80&
Private Const GMEM_MOVEABLE = &H2&
Private Const GMEM_NOCOMPACT = &H10&
Private Const GMEM_NODISCARD = &H20&
Private Const GMEM_NOT_BANKED = &H1000&
Private Const GMEM_NOTIFY = &H4000&
Private Const GMEM_SHARE = &H2000&
Private Const GMEM_VALID_FLAGS = &H7F72&
Private Const GMEM_ZEROINIT = &H40&
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
    Private Declare Function GetLocaleInfo Lib "KERNEL32" Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
Private Declare Function GetKeyboardLayout& Lib "user32" (ByVal dwLayout&) ' not NT?
Private Const DWL_ANYTHREAD& = 0&
Const LOCALE_ILANGUAGE = 1&
Private OldPar As Long, dragCharpos As Long
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal addr As Long, retval As Integer)
Private Type DRAWTEXTPARAMS
     cbSize As Long
     iTabLength As Long
     iLeftMargin As Long
     iRightMargin As Long
     uiLengthDrawn As Long
End Type
Dim tParam As DRAWTEXTPARAMS
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal addr As Long, ByVal NewVal As Integer)
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (ByRef lpRect As RECT) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExW" (ByVal hDC As Long, ByVal lpsz As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long, ByVal lpDrawTextParams As Long) As Long
Private LastListIndex As Long
Public Property Set Container(glistNN As gList)
Set glistN = glistNN
glistN.DropEnabled = True
glistN.DragEnabled = True
glistN.PreserveNpixelsHeaderRightTwips = 8
End Property
Sub Show()
glistN.PrepareToShow 5
End Sub
Public Sub PasteDoc(thisdoc As String)

 Const CF_UNICODETEXT = 13
 ClearRedo
If thisdoc = vbNullString Then
thisdoc = GetTextData(CF_UNICODETEXT)
Else
PushUndo
End If

undotext = thisdoc
undopara = mDoc.MarkParagraphID
If mDoc.TextParagraphLen(mDoc.MarkParagraphID) = 0 And glistN.SelStart > 0 Then
undopos = 1 + CharPosStart - glistN.SelStart
Else
undopos = CharPosStart
End If
undomode = 0
mDoc.InsertDoc mDoc.MarkParagraphID, Charpos, thisdoc
Render
End Sub
Public Sub PasteText(thisdoc As String)
PushUndo
undotext = thisdoc + vbCrLf
undopara = mDoc.MarkParagraphID
If mDoc.TextParagraphLen(mDoc.MarkParagraphID) = 0 And glistN.SelStart > 0 Then
undopos = 1 + CharPosStart - glistN.SelStart
Else
undopos = CharPosStart
End If
undomode = 0
InsertText = thisdoc + vbCrLf
End Sub

Private Function GetTextData(ByVal lFormatId As Long) As String
Dim bData() As Byte, sr As String, sr1 As String
sr1 = Clipboard.GetText(1)
If (OpenClipboard(0) <> 0) Then

        
        If (GetBinaryData(lFormatId, bData())) Then
        sr = bData
    
     If IsWine Then
            sr1 = Left$(sr, Len(sr1))
            GetTextData = Left$(sr1, Len(sr1))
    Else
                GetTextData = Left$(sr, Len(sr1))
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


Public Sub HorSrcoll(widthtwips As Long)
glistN.PanPos = widthtwips
End Sub
Public Property Get FileName() As String
FileName = myfilename
End Property
Sub NewTitle(RHS As String, addpixelstoo As Long, Optional factor As Single = 1)
Title = RHS


Set mDoc1 = New Document
If Title = vbNullString Then
Else
NoHeadwrap = False
mDoc1.textDoc = Title

glistN.VerticalCenterText = True
If mDoc1.DocLines >= glistN.lines + 1 Then
mDoc1.EmptyDoc
NoHeadwrap = True
mDoc1.textDoc = Title
End If
glistN.HeadLine = vbNullString ' reset
glistN.HeadLine = " " ' give a fake headline
onelineheight = glistN.HeadlineHeight + addpixelstoo
glistN.addpixels = 4 * factor ' 'line spacing
End If
all = mDoc.DocLines
glistN.MultiLineEditBox = True
glistN.ShowBar = True
glistN.ShowBar = False  ' now is in autohide mode
''glistN.ShowMe
glistN.StickBar = True
mDoc.WrapAgain
mDoc1.WrapAgain
glistN.PrepareToShow 10
all = mDoc.DocLines
glistN.CalcAndShowBar
End Sub
Public Property Let FileName(ByVal RHS As String)
myfilename = RHS
onelineheight = 0
Dim thisdoc As New Document, header As New Document
Set mDoc = thisdoc
Set mDoc1 = header
Dim i As Long
glistN.NoCaretShow = True
ClearRedo
ClearUndo
If RHS = vbNullString Then
mDoc.textDoc = vbNullString
Else
mDoc.textDoc = ReadUnicodeOrANSI(RHS)
End If

Dim what$, ThatPara As Long, pastchars
    pastchars = mDoc.TextParagraphLine(1, what$, ThatPara)
mDoc.MarkParagraphID = ThatPara
NoTitle = Title = vbNullString
nowrap = True


If Not NoTitle Then
mDoc1.textDoc = Title

glistN.HeadLine = vbNullString ' reset
glistN.HeadLine = " " ' give a fake headline
onelineheight = glistN.HeadlineHeight
End If
''glistN.AddPixels = 4  ' use this before  ---- changed for m2000 version 2
all = mDoc.DocLines
glistN.MultiLineEditBox = True
glistN.ShowBar = True
glistN.ShowBar = False  ' now is in autohide mode
glistN.ShowMe
glistN.StickBar = True
mDoc.WrapAgain
mDoc1.WrapAgain
glistN.PrepareToShow 10
'glistN.ShowMe2
all = mDoc.DocLines
glistN.CalcAndShowBar
'glistN.FloatList = True
End Property
Public Sub ResetSelColors()
SelectionTextColor = &HFFFFFF
SelectionColor = &H666666
End Sub
Property Let TabWidth(RHS As Long)
    tParam.iTabLength = Abs(RHS)
    glistN.TabWidthChar = Abs(RHS)
End Property


Private Sub Class_Initialize()
tParam.cbSize = LenB(tParam)
tParam.iTabLength = 4

Set mDoc = New Document
Set mDoc1 = New Document
ResetSelColors
mDoc.lcid = cLid
mDoc1.lcid = cLid
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
undopara = -1
Redopara = -1
End Sub

Private Sub glistN_addone(that As String)  '
' push undo addone
' from delete or backspace..If Not enabled Then Exit Sub
ClearRedo
If that = vbCrLf Then
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart
undomode = 2

ElseIf undomode = 1 And undotext <> "" Then
If undopos = SelStart Then
' we add after
undotext = undotext + that
Else
undotext = that + undotext
undopos = SelStart
End If
Else
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart ' always the left side of char - so 0 is inside
undomode = 1
End If

End Sub
Public Sub AddUndo(that As String)
If Not enabled Then Exit Sub
ClearRedo
If that = vbCrLf Then
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart
undomode = 2

Else
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart ' always the left side of char - so 0 is inside
undomode = 1
End If

End Sub
Public Sub RemoveUndo(that As String)
ClearRedo
If that = vbCrLf Then
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart
undomode = 3
Else
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
If mDoc.TextParagraphLen(mDoc.MarkParagraphID) = 0 And glistN.SelStart > 0 Then
undopos = 1 + SelStart - glistN.SelStart
Else
undopos = SelStart
End If
undomode = 0
End If
End Sub
Public Sub insertbrackets()
Dim i As Long, j As Long, ss$
If SelLength > 0 Then
ClearRedo
PushUndo
glistN_MarkDelete False
ss$ = undotext
End If
Dim a$
glistN.enabled = False
    a$ = Left$(CurrentParagraph, Charpos)
    If glistN.UseTab And (MyTrimL2NoTab(a$) = 1 Or Len(a$) = 0) Then
    i = ((Len(a$) - Len(NLtrim$(a$))))
    j = SelStart
    If ss$ <> "" Then
        ss$ = Replace(ss$, vbCrLf, vbCrLf + vbTab)
        RemoveUndo "{" + vbCrLf + String$(i + 1, vbTab) + ss$ + vbCrLf + String$(i, vbTab) + "}"
        InsertText = "{" + vbCrLf + String$(i + 1, vbTab) + ss$ + vbCrLf + String$(i, vbTab) + "}"
    Else
        RemoveUndo "{" + vbCrLf + String$(i + 1, vbTab) + vbCrLf + String$(i, vbTab) + "}"
        InsertText = "{" + vbCrLf + String$(i + 1, vbTab) + vbCrLf + String$(i, vbTab) + "}"
    End If
    SelStartSilent = j + i + 4
    Else
    i = ((Len(a$) - Len(LTrim$(a$))) \ 6) * 6
    j = SelStart
    If ss$ <> "" Then
        ss$ = Replace(ss$, vbCrLf, vbCrLf + space$(6))
        RemoveUndo "{" + vbCrLf + space$(i + 6) + ss$ + vbCrLf + space$(i) + "}"
        InsertText = "{" + vbCrLf + space$(i + 6) + ss$ + vbCrLf + space$(i) + "}"
    Else
        RemoveUndo "{" + vbCrLf + space$(i + 6) + vbCrLf + space$(i) + "}"
        InsertText = "{" + vbCrLf + space$(i + 6) + vbCrLf + space$(i) + "}"
    End If
    SelStartSilent = j + i + 3 + 6
    End If
glistN.enabled = True
Render
End Sub
Public Sub ManualInform()
If Not glistN.enabled Then Exit Sub
If glistN.SelStart = 0 Then
RaiseEvent inform(glistN.ListIndex + 1, 1)
Else
RaiseEvent inform(glistN.ListIndex + 1, (glistN.SelStart))
End If
glistN.ShowMe
End Sub



Private Sub glistN_CorrectCursorAfterDrag()
    If Not mDoc.InvalidPara(OldPar) Then
        SelLengthSilent = 0
        mDoc.MarkParagraphID = OldPar
        glistN.enabled = False
        ParaSelStart = dragCharpos
        glistN.enabled = True
        ManualInform
        
    End If
End Sub

Private Sub glistN_CtrlPlusF1()
RaiseEvent CtrlPlusF1
End Sub

Private Sub glistN_DragOverCursor(ok As Boolean)
ok = True
End Sub

Private Sub glistN_getpair(a As String, b As String)
If Not UsedAsTextBox Then
Select Case a
Case Chr$(34)
b = Chr$(34)
Case "("
b = ")"
Case "{"
b = "}"
Case "["
b = "]"
End Select
End If
End Sub

Private Sub glistN_HaveMark(Yes As Boolean)
Yes = HaveMarkedText
End Sub

Private Sub glistN_MarkCut(ThatData As String)
ThatData = undotext
End Sub

Private Sub glistN_Maybelanguage()
If Not glistN.enabled Then Exit Sub
If glistN.SelStart = 0 Then
RaiseEvent inform(glistN.ListIndex + 1, 1)
Else
RaiseEvent inform(glistN.ListIndex + 1, (glistN.SelStart))
End If
Render
End Sub

Private Sub glistN_PureListOff()
showparagraphWork = True

End Sub

Private Sub glistN_PureListOn()
showparagraphWork = False

End Sub

Private Sub glistN_PushMark2Undo(that As String)
ClearRedo
undopara = BlockEndPara
undotext = that
undomode = 1
undopos = SelStart
End Sub

Private Sub glistN_PushUndoIfMarked()
If HaveMarkedText Then PushUndo
End Sub

Private Sub glistN_RealCurReplace(a As String)
If showparagraph Then
a = Replace(a, vbTab, ChrW(&H21E5))
End If
End Sub



Private Sub glistN_RemoveOne(that As String)
' push undo removeone
If Not enabled Then Exit Sub

ClearRedo
If that = vbCrLf Then
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
undopos = SelStart
undomode = 3
ElseIf undomode = 0 And undotext <> "" Then
If undopos + Len(undotext) = SelStart Then
' we add after
undotext = undotext + that

Else
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
If mDoc.TextParagraphLen(mDoc.MarkParagraphID) = 0 And glistN.SelStart > 0 Then
undopos = 1 + SelStart - glistN.SelStart
Else
If glistN.ListIndex = all Then
undopos = SelStart - 1
Else
undopos = SelStart
End If
End If
 ' always the left side of char - so 0 is inside
undomode = 0
End If

Else
PushUndo
undopara = mDoc.MarkParagraphID
undotext = that
If mDoc.TextParagraphLen(mDoc.MarkParagraphID) = 0 And glistN.SelStart > 0 Then
undopos = 1 + SelStart - glistN.SelStart
Else
undopos = SelStart
End If
undomode = 0
End If
End Sub


Private Sub glistN_MayRefresh(ok As Boolean)
ok = Not HaveMarkedText
End Sub


Private Sub glistN_ChangeListItem(item As Long, content As String)
Dim selstart2 As Long, checklistindex As Long
'' only for M2000, use NumberOnly


If NumberOnly Then
If Not ValidNum(content, False, NumberIntOnly) Then


undotext = vbNullString
content = mDoc.TextLine(item + 1)
Exit Sub
End If
End If
If UsedAsTextBox Then
If Len(content) > glistN.maxchar Then
Charpos = glistN.maxchar
glistN.SelStartEventAlways = glistN.SelStart - 1
Exit Sub
End If
End If
mDoc.TextLine(item + 1) = content
WrapMarkedPara
If all <> mDoc.DocLines Then
    Render
Else
    checklistindex = mDoc.FirstParagraphLine(mDoc.MarkParagraphID) + mDoc.TextParagraph2Line(mDoc.MarkParagraphID, Charpos, selstart2)
    If (checklistindex - glistN.ScrollFrom) >= glistN.lines + 1 Then
     glistN.ListindexPrivateUse = checklistindex
     glistN.ScrollToTextEdit checklistindex - glistN.lines
     glistN.SelStartEventAlways = selstart2
     

    
    ElseIf checklistindex < glistN.ScrollFrom Then
     glistN.ListindexPrivateUse = checklistindex
     glistN.ScrollToTextEdit checklistindex
          glistN.SelStartEventAlways = selstart2
    Else
    glistN.ListindexPrivateUse = checklistindex
    glistN.SelStart = selstart2
        End If
End If

End Sub

Private Sub glistN_ChangeSelStart(thisselstart As Long)
If glistN.ListIndex < 0 Then Exit Sub
Dim what$, ThatPara As Long, pastchars As Long
pastchars = mDoc.TextParagraphLine2(glistN.ListIndex + 1, what$, ThatPara, False)
Charpos = pastchars + glistN.SelStart

If Not glistN.enabled Then Exit Sub
If glistN.SelStart = 0 Then
RaiseEvent inform(glistN.ListIndex + 1, 1)
Else
RaiseEvent inform(glistN.ListIndex + 1, (glistN.SelStart))
End If
If NoCenterLineEdit Then Exit Sub
If LastListIndex <> glistN.ListIndex Then LastListIndex = glistN.ListIndex: glistN.ShowMe

End Sub

Private Sub glistN_DragData(ThatData As String)
FillBlock
ThatData = UndoDoc.textDoc
 OldPar = mDoc.MarkParagraphID
 dragCharpos = Charpos
End Sub

Private Sub glistN_DragPasteData(ThatData As String)

Dim a() As String, k As Long
a() = Split(ThatData, ChrW(9))
PasteDoc Join(a(), "    ")
RelocateMarkedText

End Sub

Private Sub glistN_DropFront(ok As Boolean)
ok = False
' we need to know if we delete mark before paste data
' when we perform dragdrop to the same document
' by defult ok=false so we drop behind so we can delete marked text and then we paste
If HaveMarkedText Then
If glistN.ListIndex < BlockStartLine Then
' we are behind
ElseIf glistN.ListIndex = BlockStartLine Then
If glistN.SelStart < BlockStartCursor Then
Else
ok = True
End If
Else
ok = True
End If
End If
End Sub

Private Sub glistN_DropOk(ok As Boolean)
If HaveMarkedText Then
If ClickInsideMarkedText Then
Else
ok = True
End If
Else
ok = True
End If
End Sub
Private Sub glistN_DragOverDone(a As Boolean)

If Not mDoc.InvalidPara(OldPar) Then
        a = True
        SelLengthSilent = 0
        mDoc.MarkParagraphID = OldPar
        glistN.enabled = False
        ParaSelStart = dragCharpos
        glistN.enabled = True
        ManualInform
        OldPar = 0
    End If
End Sub
Private Sub glistN_DragOverNow(a As Boolean)
If OldPar = 0 Then
a = True
 SelLengthSilent = 0
 OldPar = mDoc.MarkParagraphID
 dragCharpos = Charpos
End If
End Sub
Private Sub glistN_ExposeListcount(cListCount As Long)
cListCount = all
End Sub
Private Sub gListN_ExposeRect(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
Dim a As RECT, b As RECT, i As Long, pastchars As Long, hRgn As Long
Dim final$, fbypass As Boolean
Dim OldColor As Long
Dim what$, ThatPara As Long
CopyFromLParamToRect a, thisrect
CopyFromLParamToRect b, thisrect
If UsedAsTextBox Then
b.Bottom = a.Bottom - 2
End If
Dim skipcolor As Boolean
OldColor = glistN.ForeColor
If item < 0 Then  ' this is a headline
FillBack thisHDC, b, glistN.CapColor
  mDoc1.TextParagraphLine2 1, what$, ThatPara, False
  showparagraphonce = False
   glistN.HeadlineHeight = mDoc1.DocLines * onelineheight

  For i = 1 To mDoc1.DocLines
  
    PrintItem2 thisHDC, mDoc1.TextLine2(i, False, True), a
    a.top = a.top + onelineheight
  Next i
  RaiseEvent ExposeRect(item, thisrect, thisHDC, skip)
  skip = True
   Exit Sub
   
Else
showparagraphWork = showparagraph
  ' a.Left = a.Left + 1
    a.top = a.top + glistN.addpixels \ 2
    pastchars = mDoc.TextParagraphLine2(item + 1, what$, ThatPara, showparagraph)
    
        If MarkedTextAny Then
        showparagraphWork = True
    If item > BlockTopline And item < BlockBottomline Then
        FillBack thisHDC, b, SelectionColor
         SetTextColor thisHDC, SelectionTextColor
         skipcolor = True
    ElseIf item = BlockTopline And item <> BlockBottomline Then

        
             If BlockEndline < BlockStartLine Then
                 If BlockEndCursor > 1 Then  ' if not leave it as is
                 b.Left = b.Left + glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockEndCursor - 1))
                 PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
                final$ = Mid$(glistN.list(item), BlockEndCursor)
                fbypass = True
                 End If
             Else
                  If BlockStartCursor > 1 Then  ' if not leave it as is
                 b.Left = b.Left + glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockStartCursor - 1))
                 PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
              
                 final$ = Mid$(glistN.list(item), BlockStartCursor)
                 fbypass = True
                 End If
             End If
     
             FillBack thisHDC, b, SelectionColor
             'a.Left = b.Left + 1
             
         SetTextColor thisHDC, SelectionTextColor
         skipcolor = True
     ElseIf item = BlockBottomline And item <> BlockTopline Then
      
             If BlockBottomline = BlockEndline Then
                 If BlockEndCursor <= Len(glistN.list(item)) Then
                 If BlockEndCursor > 1 Then ' if not leave it as is
                 b.Right = glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockEndCursor - 1)) + glistN.PanPosPixels + glistN.LeftMarginPixels
                 PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
                final$ = Mid$(glistN.list(item), 1, BlockEndCursor - 1)
             fbypass = False
                 Else
                   If ThatPara = mDoc.MarkParagraphID Then glistN.ExternalCursor glistN.SelStart, what$, OldColor
                  GoTo conthere
                 End If
                 End If
                 
             Else
                  If BlockStartCursor <= Len(glistN.list(item)) Then  ' if not leave it as is
               If BlockStartCursor > 1 Then
               b.Right = glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockStartCursor - 1)) + glistN.PanPosPixels + glistN.LeftMarginPixels
                 PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
                final$ = Mid$(glistN.list(item), 1, BlockStartCursor - 1)
               fbypass = False
                     Else
                       If ThatPara = mDoc.MarkParagraphID Then glistN.ExternalCursor glistN.SelStart, what$, OldColor
                      GoTo conthere
                 End If
                 End If
                 
             End If
     
             FillBack thisHDC, b, SelectionColor
             a.Right = b.Right + 1
             
         SetTextColor thisHDC, SelectionTextColor
         skipcolor = True
    ElseIf BlockBottomline = item Then ' in one line only
     skipcolor = False
        If BlockEndCursor > BlockStartCursor Then
         If BlockStartCursor > 1 Then b.Left = b.Left + glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockStartCursor - 1))
          If BlockEndCursor > 1 Then b.Right = glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockEndCursor - 1)) + glistN.PanPosPixels + glistN.LeftMarginPixels
          If BlockStartCursor = 0 Then
     final$ = Mid$(glistN.list(item), 1, BlockEndCursor - 1)
          Else
           skipcolor = False
 final$ = Mid$(glistN.list(item), BlockStartCursor, BlockEndCursor - BlockStartCursor)
 End If

 PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
          FillBack thisHDC, b, SelectionColor
             a.Right = b.Right + 1
             'a.Left = b.Left + 1
         SetTextColor thisHDC, SelectionTextColor
        ElseIf BlockEndCursor < BlockStartCursor Then
        If BlockEndCursor > 1 Then b.Left = b.Left + glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockEndCursor - 1))
        If BlockStartCursor > 1 Then b.Right = glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockStartCursor - 1)) + glistN.PanPosPixels + glistN.LeftMarginPixels
 If BlockEndCursor > 0 Then
 final$ = Mid$(glistN.list(item), BlockEndCursor, BlockStartCursor - BlockEndCursor)
 Else
 final$ = Mid$(glistN.list(item), 1, BlockStartCursor - 1)
 End If
 
 PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
                 FillBack thisHDC, b, SelectionColor
             a.Right = b.Right + 1
             'a.Left = b.Left + 1
         SetTextColor thisHDC, SelectionTextColor
        End If
       fbypass = False
    End If
 ''  skipcolor = True
End If
    If ThatPara = mDoc.MarkParagraphID Then
            If EditDoc Then
                        ' old patch -- removed here
                      If mNoWrap And (glistN.ListIndex = item) Then
                      glistN.ExternalCursor glistN.SelStart, what$, OldColor
                       ElseIf (Not skipcolor) And (glistN.ListIndex = item) Then
                      glistN.ExternalCursor glistN.SelStart, what$, OldColor
                       End If
            ElseIf Not NoMark Then
            
                    FillBack thisHDC, b, 0
                     SetTextColor thisHDC, SelectionTextColor
                 
                     
            End If
    
    Else
         
    If Not skipcolor Then SetTextColor thisHDC, glistN.ForeColor
 

    End If
    If final$ <> "" Or fbypass Then

   If Not fbypass Then
 
   'PrintItem thisHDC, final$, a
   hRgn = CreateRectRgnIndirect(b)
    SelectClipRgn thisHDC, hRgn
   PrintItem thisHDC, mDoc.TextLine2(item + 1, showparagraph, False), a
   SelectClipRgn thisHDC, &H0
   DeleteObject hRgn
   Else
'   PrintItem1 thisHDC, final$, a
hRgn = CreateRectRgnIndirect(b)
    SelectClipRgn thisHDC, hRgn
   PrintItem1 thisHDC, mDoc.TextLine2(item + 1, showparagraph, False), a
   SelectClipRgn thisHDC, &H0
   DeleteObject hRgn
   End If
   showparagraphWork = False
    Else
conthere:
If skipcolor Or NoColor Then
 PrintItem1 thisHDC, mDoc.TextLine2(item + 1, showparagraph, False), a
Else
  PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
End If
  End If
End If
SetTextColor thisHDC, OldColor
 showparagraphonce = False

skip = True
End Sub

Private Sub gListN_ExposeRectOLD(ByVal item As Long, ByVal thisrect As Long, ByVal thisHDC As Long, skip As Boolean)
Dim a As RECT, b As RECT, i As Long, pastchars As Long
Dim final$, fbypass As Boolean
Dim OldColor As Long
Dim what$, ThatPara As Long
CopyFromLParamToRect a, thisrect
CopyFromLParamToRect b, thisrect
If UsedAsTextBox Then
b.Bottom = a.Bottom - 2
End If
Dim skipcolor As Boolean
OldColor = glistN.ForeColor
If item < 0 Then  ' this is a headline
FillBack thisHDC, b, glistN.CapColor
  mDoc1.TextParagraphLine2 1, what$, ThatPara, False
  showparagraphonce = False
   glistN.HeadlineHeight = mDoc1.DocLines * onelineheight

  For i = 1 To mDoc1.DocLines
  
 '   PrintItem2 thisHDC, mDoc1.TextLine(i), A
    PrintItem2 thisHDC, mDoc1.TextLine2(i, False, True), a
    a.top = a.top + onelineheight
  Next i
  RaiseEvent ExposeRect(item, thisrect, thisHDC, skip)
  skip = True
   Exit Sub
   
Else
   a.Left = a.Left + 1
    a.top = a.top + glistN.addpixels \ 2
    pastchars = mDoc.TextParagraphLine2(item + 1, what$, ThatPara, showparagraph)
    
        If MarkedTextAny Then
    If item > BlockTopline And item < BlockBottomline Then
        FillBack thisHDC, b, SelectionColor
         SetTextColor thisHDC, SelectionTextColor
         skipcolor = True
    ElseIf item = BlockTopline And item <> BlockBottomline Then

        
             If BlockEndline < BlockStartLine Then
                 If BlockEndCursor > 1 Then  ' if not leave it as is
                 b.Left = b.Left + glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockEndCursor - 1))
                 PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
                final$ = Mid$(glistN.list(item), BlockEndCursor)
                fbypass = True
                 End If
             Else
                  If BlockStartCursor > 1 Then  ' if not leave it as is
                 b.Left = b.Left + glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockStartCursor - 1))
                 PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
              
                 final$ = Mid$(glistN.list(item), BlockStartCursor)
                 fbypass = True
                 End If
             End If
     
             FillBack thisHDC, b, SelectionColor
             a.Left = b.Left + 1
             
         SetTextColor thisHDC, SelectionTextColor
         skipcolor = True
     ElseIf item = BlockBottomline And item <> BlockTopline Then
      
             If BlockBottomline = BlockEndline Then
                 If BlockEndCursor <= Len(glistN.list(item)) Then
                 If BlockEndCursor > 1 Then ' if not leave it as is
                 b.Right = glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockEndCursor - 1)) + glistN.PanPosPixels + glistN.LeftMarginPixels
                 PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
                final$ = Mid$(glistN.list(item), 1, BlockEndCursor - 1)
             fbypass = False
                 Else
                   If ThatPara = mDoc.MarkParagraphID Then glistN.ExternalCursor glistN.SelStart, what$, OldColor
                  GoTo conthere
                 End If
                 End If
                 
             Else
                  If BlockStartCursor <= Len(glistN.list(item)) Then  ' if not leave it as is
               If BlockStartCursor > 1 Then
               b.Right = glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockStartCursor - 1)) + glistN.PanPosPixels + glistN.LeftMarginPixels
                 PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
                final$ = Mid$(glistN.list(item), 1, BlockStartCursor - 1)
               fbypass = False
                     Else
                       If ThatPara = mDoc.MarkParagraphID Then glistN.ExternalCursor glistN.SelStart, what$, OldColor
                      GoTo conthere
                 End If
                 End If
                 
             End If
     
             FillBack thisHDC, b, SelectionColor
             a.Right = b.Right + 1
             
         SetTextColor thisHDC, SelectionTextColor
         skipcolor = True
    ElseIf BlockBottomline = item Then ' in one line only
     skipcolor = False
        If BlockEndCursor > BlockStartCursor Then
         If BlockStartCursor > 1 Then b.Left = b.Left + glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockStartCursor - 1))
          If BlockEndCursor > 1 Then b.Right = glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockEndCursor - 1)) + glistN.PanPosPixels + glistN.LeftMarginPixels
          If BlockStartCursor = 0 Then
     final$ = Mid$(glistN.list(item), 1, BlockEndCursor - 1)
          Else
           skipcolor = False
 final$ = Mid$(glistN.list(item), BlockStartCursor, BlockEndCursor - BlockStartCursor)
 End If

 PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
          FillBack thisHDC, b, SelectionColor
             a.Right = b.Right + 1
             a.Left = b.Left + 1
         SetTextColor thisHDC, SelectionTextColor
        ElseIf BlockEndCursor < BlockStartCursor Then
        If BlockEndCursor > 1 Then b.Left = b.Left + glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockEndCursor - 1))
        If BlockStartCursor > 1 Then b.Right = glistN.UserControlTextWidthPixels(Mid$(glistN.list(item), 1, BlockStartCursor - 1)) + glistN.PanPosPixels + glistN.LeftMarginPixels
 If BlockEndCursor > 0 Then
 final$ = Mid$(glistN.list(item), BlockEndCursor, BlockStartCursor - BlockEndCursor)
 Else
 final$ = Mid$(glistN.list(item), 1, BlockStartCursor - 1)
 End If
 
 PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
                 FillBack thisHDC, b, SelectionColor
             a.Right = b.Right + 1
             a.Left = b.Left + 1
         SetTextColor thisHDC, SelectionTextColor
        End If
       fbypass = False
    End If
 ''  skipcolor = True
End If
    If ThatPara = mDoc.MarkParagraphID Then
            If EditDoc Then
                        ' old patch -- removed here
                      If mNoWrap And (glistN.ListIndex = item) Then
                      glistN.ExternalCursor glistN.SelStart, what$, OldColor
                       ElseIf (Not skipcolor) And (glistN.ListIndex = item) Then
                            glistN.ExternalCursor glistN.SelStart, what$, OldColor
                       End If
            ElseIf Not NoMark Then
            
                    FillBack thisHDC, b, 0
                     SetTextColor thisHDC, SelectionTextColor
            End If
    
    Else
         
    If Not skipcolor Then SetTextColor thisHDC, glistN.ForeColor
 

    End If
    If final$ <> "" Or fbypass Then

   If Not fbypass Then
 
   PrintItem thisHDC, final$, a
   Else
   PrintItem1 thisHDC, final$, a
   End If
    Else
conthere:
If skipcolor Or NoColor Then
    
  PrintItem1 thisHDC, mDoc.TextLine2(item + 1, showparagraph, False), a
Else
  PrintItem1X thisHDC, mDoc.TextLine2(item + 1, showparagraph, NoColor), a, mDoc.lastcolordata()
End If
  End If
End If
SetTextColor thisHDC, OldColor
 showparagraphonce = False

skip = True
End Sub

Public Sub Render()
If mDoc.busy Then Exit Sub
Dim selstart2 As Long
mDoc.busy = True
all = mDoc.DocLines
glistN.ListIndex = mDoc.FirstParagraphLine(mDoc.MarkParagraphID) + mDoc.TextParagraph2Line(mDoc.MarkParagraphID, Charpos, selstart2, NoColor)
glistN.SelStart = selstart2
If glistN.ListIndex < glistN.ScrollFrom Then
If glistN.ListIndex > glistN.lines Then
glistN.ScrollToSilent glistN.ListIndex - glistN.lines / 2
Else
glistN.ScrollToSilent 0
End If
ElseIf glistN.ListIndex > glistN.ScrollFrom + glistN.lines Then
If glistN.listcount - glistN.ListIndex < glistN.lines Then
glistN.ScrollToSilent glistN.listcount - glistN.lines
Else
glistN.ScrollToSilent glistN.ListIndex - glistN.lines / 2
End If
ElseIf glistN.ListIndex > glistN.lines And glistN.listcount - glistN.ListIndex < glistN.lines Then
glistN.ScrollToSilent glistN.listcount - glistN.lines
End If
If glistN.enabled Then glistN.ShowPan
mDoc.busy = False
End Sub

Public Sub WrapMarkedPara()
If glistN.ListIndex >= 0 Then
mDoc.ReWritePara mDoc.MarkParagraphID, mDoc.TextParagraph(mDoc.MarkParagraphID)
End If
End Sub




Private Sub glistN_KeyDown(KeyCode As Integer, shift As Integer)
If KeyCode = vbKeyTab Then
If shift = 2 Then choosenext
        
    KeyCode = 0
    Exit Sub
    

End If

Dim selstart2 As Long, b As Boolean, Pad$, cc As Long
Dim l1 As Long, c1 As Long
If KeyCode = 18 Then
Exit Sub
End If
If KeyCode = 16 And shift <> 0 Then
KeyCode = 0
If glistN.SelStart = 0 Then
RaiseEvent inform(glistN.ListIndex + 1, 1)
Else
RaiseEvent inform(glistN.ListIndex + 1, (glistN.SelStart))
End If
Exit Sub
End If
If NoMark Then KeyCode = 0: Exit Sub

If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyEnd Or KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
'
ElseIf KeyCode = vbKeyX And (shift And 3) = 2 Then
MarkCut
Show
KeyCode = 0
ElseIf KeyCode = vbKeyC And (shift And 3) = 2 Then
MarkCopy
KeyCode = 0
ElseIf KeyCode = vbKeyV And (shift And 3) = 2 Then
Pad$ = GetTextData(13)
If Pad$ = vbNullString Then Pad$ = Clipboard.GetText(1)
If ParaSelStart = 2 And glistN.list(glistN.ListIndex) = vbNullString Then
SelStart = SelStart - 1
End If
AddUndo ""
If UsedAsTextBox Then
Pad$ = GetNextLine((Pad$))
If Len(Text) + Len(Pad$) > glistN.maxchar Then
Pad$ = Left$(Pad$, glistN.maxchar - Len(Text))
End If
End If
SelText = Pad$
RemoveUndo Pad$
KeyCode = 0
'mDoc.WrapAgainColorBlock mDoc.MarkParagraphID, mDoc.MarkParagraphID
ElseIf KeyCode = vbKeyA And (shift And 3) = 2 Then
glistN.MarkNext = 0
glistN.MarkALL
KeyCode = 0
ElseIf KeyCode = vbKeyY And (shift And 3) = 2 Then
' Redo
If Redopara <> -1 Then

PushUndo

undopara = Redopara
undopos = Redopos
undotext = Redotext
undomode = Redomode

b = enabled
enabled = False
SelLength = 0
If Redomode = 3 Then
SelStart = Redopos - 2
Else
SelStart = Redopos
End If
If Redomode = 0 Or Redomode = 5 Then ' now we do the opposite

If Redomode = 5 Then


PopRedo

If glistN.SelStart = 0 Then
SelLength = Len(undotext) + 1
Else
SelLength = Len(undotext)
End If
    glistN_MarkDelete True
    glistN.MarkNext = 0
PushUndo
InsertText = Redotext
undopara = Redopara
undopos = Redopos
undotext = Redotext
undomode = Redomode

Else
InsertText = Redotext

End If
ElseIf Redomode = 1 Then
If glistN.SelStart = 0 Then
SelLength = Len(Redotext) + 1
Else
SelLength = Len(Redotext)
End If
SelText = vbNullString
ElseIf Redomode = 2 Then
   keyzLineUp
ElseIf Redomode = 3 Then
glistN_SplitLine

End If

enabled = b
'If Not NoCenterLineEdit Then
'glistN.ShowMe
'Else
glistN.ShowPan
'End If
If Not PopRedo Then
Redopara = -1
Redotext = vbNullString
Redomode = 0
End If
End If
KeyCode = 0
ElseIf KeyCode = vbKeyZ And (shift And 3) = 2 Then
' 2nd version of undo system
' undopos now is selstart
If undopara <> -1 Then
' seve to redo...
PushRedo
Redopara = undopara
Redopos = undopos
Redotext = undotext
Redomode = undomode


' MOVE TO PARAGRAPH...AND CHARPOS
b = enabled
enabled = False

   SelLength = 0
   SelStart = undopos

If undomode = 0 Or undomode = 4 Then

    If glistN.SelStart = 0 Then
        SelLength = Len(undotext) + 1
    Else
        SelLength = Len(undotext)
    End If
    
    If undomode = 4 Then
    glistN_MarkDelete True
    glistN.MarkNext = 0
    PopUndo
    PushRedo
        Redopara = undopara
        Redopos = undopos
        Redotext = undotext
        Redomode = undomode

    ' now we have undomode 5
    InsertText = undotext
    Else
    SelText = vbNullString
    End If
ElseIf undomode = 3 Then
    keyzLineUp
ElseIf undomode = 2 Then
    glistN_SplitLine
Else
        InsertText = undotext
End If
enabled = b
'If Not NoCenterLineEdit Then
'glistN.ShowMe
'Else
glistN.ShowPan
'End If

End If
 If Not PopUndo Then
 ClearUndo
undopara = -1
undotext = vbNullString
undomode = 0
End If
KeyCode = 0
ElseIf KeyCode = vbKeyA And (shift And 3) = 3 Then
glistN.MarkNext = 0
BlockEndPara = mDoc.MarkParagraphID
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.ShowPan
KeyCode = 0

ElseIf KeyCode = vbKeyDelete Then
If HaveMarkedText Then
ClearRedo
PushUndo
glistN_MarkDelete False

KeyCode = 0
End If
ElseIf KeyCode = vbKeyBack Then
If HaveMarkedText Then
ClearRedo
PushUndo
glistN_MarkDelete False

KeyCode = 0
End If
End If

End Sub

Private Sub glistN_KeyDownAfter(KeyCode As Integer, shift As Integer)
If NoMark Then Exit Sub
If HaveMarkedText Then
If shift = 0 Then
glistN.MarkNext = 0
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
ElseIf (shift And 1) = 1 Then

If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyEnd Or KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
glistN_MarkOut
End If
End If
ElseIf (shift And 1) = 1 And glistN.MarkNext = 0 Then
If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyEnd Or KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
glistN.MarkNext = 1
glistN_MarkIn
End If
ElseIf (shift And 1) = 1 And glistN.MarkNext > 0 Then
If KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Or KeyCode = vbKeyEnd Or KeyCode = vbKeyHome Or KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
glistN.MarkNext = 2
glistN_MarkOut
End If
End If
End Sub

Private Sub glistN_LineDown()
Dim b$, nline As Long
mDoc.TextParagraphLine glistN.ListIndex + 1, b$, nline
mDoc.MarkParagraphID = nline

If mDoc.IsLast(mDoc.MarkParagraphID) Then Exit Sub

  
 
 If mDoc.TextParagraphLen(nline) = 0 Then
 Charpos = 1
  mDoc.RemoveEmptyParagraph nline
 mDoc.MarkParagraphID = nline
 Else
 Charpos = ParaSelStart
  mDoc.ReWritePara nline, mDoc.TextParagraph(nline) + mDoc.DeleteNextParagraph(nline)

 End If
 
  all = mDoc.DocLines
 
  ParaSelStart = Charpos
End Sub

Private Sub glistN_LineUp()
Dim olID As Long, b$ ', c$
olID = mDoc.MarkParagraphID
If ParaSelStart < 2 Then
If mDoc.ParagraphOrder(mDoc.MarkParagraphID) > 1 Then
glistN.PressKey vbKeyLeft, 0
If olID <> mDoc.MarkParagraphID Then
b$ = mDoc.DeleteNextParagraph(mDoc.MarkParagraphID)
mDoc.ReWritePara mDoc.MarkParagraphID, mDoc.TextParagraph(mDoc.MarkParagraphID) + b$
 all = mDoc.DocLines
ClearRedo
PushUndo
undotext = vbCrLf
undopara = mDoc.MarkParagraphID
undopos = SelStart
undomode = 2
'End If
Render
End If
End If
Else
glistN.PressKey vbKeyLeft, 0
End If
all = mDoc.DocLines

End Sub

Private Sub keyzLineUp()
Dim olID As Long, b$ ', c$
olID = mDoc.MarkParagraphID
If ParaSelStart < 2 Then
If mDoc.ParagraphOrder(mDoc.MarkParagraphID) > 1 Then
glistN.PressKey vbKeyLeft, 0
If olID <> mDoc.MarkParagraphID Then
b$ = mDoc.DeleteNextParagraph(mDoc.MarkParagraphID)
mDoc.ReWritePara mDoc.MarkParagraphID, mDoc.TextParagraph(mDoc.MarkParagraphID) + b$
Render
End If
End If
Else
glistN.PressKey vbKeyLeft, 0
End If
all = mDoc.DocLines
End Sub

Public Sub MarkCut()
MarkCopy
PushUndo
glistN_MarkDelete False

End Sub
Private Sub FillBlock()
Dim order1 As Long, order2 As Long, a$, b$, i As Long, nn As Long
' TAKE FIRST VBCRLF IF FIRST LINE =""
On Error GoTo fbl

 Set UndoDoc = New Document
 UndoDoc.CrTail = False
    If BlockStartPara = BlockEndPara Then
        b$ = mDoc.TextParagraph(BlockStartPara)
        
        UndoDoc.CrTail = False
        If b$ = vbNullString Then
        Else
        If BlockStartPos < BlockEndPos Then
           '' UndoDoc.textDoc = Mid$(b$, BlockStartPos, BlockEndPos - BlockStartPos)
            UndoDoc.textParagraphPlusColor Mid$(b$, BlockStartPos, BlockEndPos - BlockStartPos), mDoc.Paracolordata(BlockStartPara), BlockStartPos - 1
            If BlockEndPos > Len(b$) + 1 Then UndoDoc.CrTail = True
        Else
          '' UndoDoc.textDoc = Mid$(b$, BlockEndPos, BlockStartPos - BlockEndPos)
            UndoDoc.textParagraphPlusColor Mid$(b$, BlockEndPos, BlockStartPos - BlockEndPos), mDoc.Paracolordata(BlockStartPara), BlockEndPos - 1
            If BlockStartPos > Len(b$) + 1 Then UndoDoc.CrTail = True
        End If
        End If
    Else
        order1 = mDoc.ParagraphOrder(BlockStartPara)
        order2 = mDoc.ParagraphOrder(BlockEndPara)
        If order2 > order1 Then '  normal
            a$ = mDoc.TextParagraph(BlockStartPara)
            UndoDoc.LfLead = False
         If BlockStartPos < 1 Then
               UndoDoc.AppendParagraphPlusColor a$, mDoc.Paracolordata(BlockStartPara)
               '' UndoDoc.AppendParagraph a$
            Else
            If Len(a$) + 1 < BlockStartPos Then
                UndoDoc.LfLead = True
                UndoDoc.AppendParagraph ""  '' no color
               Else
            
                ''UndoDoc.AppendParagraph Mid$(a$, BlockStartPos)
                UndoDoc.AppendParagraphPlusColor Mid$(a$, BlockStartPos), mDoc.Paracolordata(BlockStartPara), BlockStartPos - 1
                End If
            End If
            UndoDoc.MarkParagraphID = 1
            nn = BlockStartPara
            order2 = order2 - 1
            If order2 > order1 Then
                For i = order1 + 1 To order2
               ' UndoDoc.MarkParagraphID = UndoDoc.InsertAfterParagraph(UndoDoc.MarkParagraphID, mDoc.TextNextParagraph(nn))
                UndoDoc.MarkParagraphID = UndoDoc.InsertAfterParagraphPlusColor(UndoDoc.MarkParagraphID, mDoc.TextNextParagraph(nn), mDoc.Paracolordata(nn))
                Next i
            End If
                b$ = mDoc.TextParagraph(BlockEndPara)
                UndoDoc.CrTail = False
                
               If BlockEndPos > 0 Then
               ''UndoDoc.AppendParagraph Left$(b$, BlockEndPos - 1) 'maybe is empty
               UndoDoc.AppendParagraphPlusColor Left$(b$, BlockEndPos - 1), mDoc.Paracolordata(BlockEndPara)
               Else
               UndoDoc.CrTail = True
               End If
        Else
            a$ = mDoc.TextParagraph(BlockEndPara)
        '   If a$ = VbNullString Then UndoDoc.AppendParagraph ""
         UndoDoc.LfLead = False
            If BlockEndPos < 1 Then
        
              UndoDoc.AppendParagraphPlusColor a$, mDoc.Paracolordata(BlockEndPara)
               ' UndoDoc.AppendParagraph a$
                
            Else
                 If Len(a$) + 1 < BlockEndPos Then
                UndoDoc.LfLead = True
                UndoDoc.AppendParagraph ""
            Else
                ' UndoDoc.AppendParagraph Mid$(a$, BlockEndPos)
                If BlockEndPos >= 1 Then
                  UndoDoc.AppendParagraphPlusColor Mid$(a$, BlockEndPos), mDoc.Paracolordata(BlockEndPara), BlockEndPos - 1
                  Else
                  UndoDoc.AppendParagraphPlusColor Mid$(a$, BlockEndPos), mDoc.Paracolordata(BlockEndPara)
                  End If
                End If
            End If
            UndoDoc.MarkParagraphID = 1
             nn = BlockEndPara
            order1 = order1 - 1
            If order1 > order2 Then
                For i = order2 + 1 To order1
                ''UndoDoc.MarkParagraphID = UndoDoc.InsertAfterParagraph(UndoDoc.MarkParagraphID, mDoc.TextNextParagraph(nn))
                UndoDoc.MarkParagraphID = UndoDoc.InsertAfterParagraphPlusColor(UndoDoc.MarkParagraphID, mDoc.TextNextParagraph(nn), mDoc.Paracolordata(nn))
                Next i
            End If
                b$ = mDoc.TextParagraph(BlockStartPara)
                If BlockStartPos < 1 Then BlockStartPos = 1
                  'UndoDoc.AppendParagraph Left$(b$, BlockStartPos - 1) 'maybe is empty
                  UndoDoc.AppendParagraphPlusColor Left$(b$, BlockStartPos - 1), mDoc.Paracolordata(BlockStartPara)
        
        End If
End If
fbl:
End Sub

Public Sub MarkCopy()
If Not NoColor Then ColorBlockOnly
If HaveMarkedText Then
FillBlock
Dim pp1 As String, pp2 As String
pp1 = UndoDoc.textDoc

If Not NoColor Then pp2 = SelHtmlText
Clipboard.clear  ' always

blockKeyboard = True
MyDoEvents
blockKeyboard = False
 On Error Resume Next
 ProcTask2 basestack1
''mywait basestack1, 10

If Not NoColor Then SimpleHtmlData pp2
 SetTextData 13, pp1

End If
End Sub
Private Function SetTextDataLong( _
        ByVal lFormatId As Long, _
         dLong As Long _
    ) As Boolean
    ' use strptr and lenb
    Dim hMem As Long, lPtr As Long
 ''   Dim checkme As Long
    Dim lSize As Long
        lSize = 4
    hMem = GlobalAlloc(0, lSize)
If (hMem > 0) Then
        lPtr = GlobalLock(hMem)
        CopyMemory ByVal lPtr, dLong, lSize
       '' CopyMemory checkme, ByVal lPtr, lSize
     
        GlobalUnlock hMem
       If (OpenClipboard(0) <> 0) Then
     SetClipboardData lFormatId, hMem
      CloseClipboard
       End If
          
    End If
    

End Function
Private Function SetTextData( _
        ByVal lFormatId As Long, _
        ByVal sText As String _
    ) As Boolean
    If lFormatId = 0 Then Exit Function
    Dim hMem As Long, lPtr As Long
    Dim lSize As Long
        lSize = LenB(sText)
    hMem = GlobalAlloc(0, lSize + 2)
If (hMem > 0) Then
        lPtr = GlobalLock(hMem)
        CopyMemory ByVal lPtr, ByVal StrPtr(sText), lSize + 1
        GlobalUnlock hMem
       If (OpenClipboard(0) <> 0) Then
     SetClipboardData lFormatId, hMem
      CloseClipboard
      Else
      GlobalFree hMem
       End If
    End If
End Function

Private Sub glistN_MarkDelete(preservecursor As Boolean)
Dim order1 As Long, order2 As Long, a$, b$, i As Long
Dim oldid As Long, OLDPOS As Long
oldid = mDoc.MarkParagraphID
OLDPOS = Charpos
 If HaveMarkedText Then
 Set UndoDoc = New Document
    If BlockStartPara = BlockEndPara Then
        b$ = mDoc.TextParagraph(BlockStartPara)
        If BlockStartPos < BlockEndPos Then
            UndoDoc.textDoc = Mid$(b$, BlockStartPos, BlockEndPos - BlockStartPos)
            b$ = Mid$(b$, 1, BlockStartPos - 1) + Mid$(b$, BlockEndPos)
                        If OLDPOS > BlockStartPos And oldid = BlockStartPara Then
            OLDPOS = OLDPOS - (BlockEndPos - BlockStartPos)
            End If
            Charpos = BlockStartPos
        Else
           UndoDoc.textDoc = Mid$(b$, BlockEndPos, BlockStartPos - BlockEndPos)
            b$ = Left$(b$, BlockEndPos - 1) + Mid$(b$, BlockStartPos)
            If OLDPOS > BlockEndPos And oldid = BlockStartPara Then
            OLDPOS = OLDPOS - (BlockStartPos - BlockEndPos)
            End If
            Charpos = BlockEndPos
        End If
        
        mDoc.ReWritePara BlockStartPara, b$
        
    Else
        order1 = mDoc.ParagraphOrder(BlockStartPara)
        order2 = mDoc.ParagraphOrder(BlockEndPara)
  
        If order2 > order1 Then '  normal
            UndoDoc.AppendParagraph "" ' start paragraph no 1
            UndoDoc.MarkParagraphID = 1
        
            b$ = mDoc.TextParagraph(BlockEndPara)
            mDoc.RemoveDocParaIndex BlockEndPara
            If BlockEndPos < 1 Then
            UndoDoc.AppendParagraph ""
            b$ = vbNullString
            Else
            UndoDoc.AppendParagraph Left$(b$, BlockEndPos - 1) 'maybe is empty
            b$ = Mid$(b$, BlockEndPos)
            End If
            
            
            order2 = order2 - 1
            If order2 > order1 Then
                For i = order1 + 1 To order2
                UndoDoc.MarkParagraphID = UndoDoc.InsertAfterParagraph(UndoDoc.MarkParagraphID, mDoc.DeleteNextParagraph(BlockStartPara))
                    
                Next i
            End If
            a$ = mDoc.TextParagraph(BlockStartPara)
            UndoDoc.ReWritePara 1, a$, True
            If BlockStartPos = 1 Then
                mDoc.ReWritePara BlockStartPara, b$
                UndoDoc.ReWritePara 1, a$, True
            Else
            UndoDoc.ReWritePara 1, Mid$(a$, BlockStartPos), True
                mDoc.ReWritePara BlockStartPara, Left$(a$, BlockStartPos - 1) + b$
            End If
            mDoc.MarkParagraphID = BlockStartPara
            Charpos = BlockStartPos
        Else
            UndoDoc.AppendParagraph "" ' start paragraph no 1
            UndoDoc.MarkParagraphID = 1
             b$ = mDoc.TextParagraph(BlockStartPara)
            mDoc.RemoveDocParaIndex BlockStartPara
             UndoDoc.AppendParagraph Left$(b$, BlockStartPos - 1) 'maybe is empty
            b$ = Mid$(b$, BlockStartPos)
            order1 = order1 - 1
            If order1 > order2 Then
                For i = order2 + 1 To order1
                    
                    UndoDoc.MarkParagraphID = UndoDoc.InsertAfterParagraph(UndoDoc.MarkParagraphID, mDoc.DeleteNextParagraph(BlockEndPara))
               
                Next i
            End If
            a$ = mDoc.TextParagraph(BlockEndPara)
            
            If BlockEndPos = 1 Then
                mDoc.ReWritePara BlockEndPara, b$
                UndoDoc.ReWritePara 1, a$
            Else
            UndoDoc.ReWritePara 1, Mid$(a$, BlockEndPos)
                mDoc.ReWritePara BlockEndPara, Left$(a$, BlockEndPos - 1) + b$
            End If
            mDoc.MarkParagraphID = BlockEndPara
            
            Charpos = BlockEndPos
        
        
        End If

End If
all = mDoc.DocLines
If Not preservecursor Then
undopara = mDoc.MarkParagraphID
undopos = SelStart
undotext = UndoDoc.textDoc
undomode = 1 ' insert text at position changed to 1
End If
Set UndoDoc = New Document


glistN.MarkNext = 0
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
If preservecursor Then
If mDoc.ParagraphOrder(oldid) > 0 Then

mDoc.MarkParagraphID = oldid
Charpos = OLDPOS
End If
End If

Render

End If

End Sub

'
Private Sub glistN_MarkDestroy()
' destroy or open to change
If Not ClickInsideMarkedText Then
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
Else
glistN.MarkNext = 1
' we have to raise event now
End If
End Sub

Private Sub glistN_MarkDestroyAny()
' destroy always
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.MarkNext = 0
End Sub

Private Sub glistN_MarkIn()
If Not glistN.EditFlag Then glistN.MarkNext = 0: Exit Sub
BlockStartPara = mDoc.MarkParagraphID
BlockStartPos = ParaSelStart
'If glistN.SelStart > 1 Then
BlockStartCursor = glistN.SelStart '- 1
'Else
'BlockStartCursor = 1
'End If
BlockEndCursor = BlockStartCursor + 1

BlockStartLine = glistN.ListIndex
BlockEndline = BlockStartLine
BlockEndPara = BlockStartPara
BlockEndPos = BlockStartPos
End Sub
Function HaveMarkedText() As Boolean
If BlockEndPos < 1 Then BlockEndPos = 1
If glistN.MarkNext = 0 Then
Else
If BlockStartPara <> BlockEndPara Then
HaveMarkedText = True
ElseIf BlockStartPos <> BlockEndPos Then
HaveMarkedText = True
End If
End If
End Function
Function MarkedTextAny() As Boolean
If BlockStartPara <> BlockEndPara Then
MarkedTextAny = True
ElseIf BlockStartPos <> BlockEndPos Then
MarkedTextAny = True
End If
End Function


Function ClickInsideMarkedText() As Boolean
' we need that for drag and drop, copy or move
' it is simple
' we now the current line and selstart of glistN
If MarkedTextAny Then
If BlockTopline <> BlockBottomline Then
If glistN.ListIndex = BlockTopline Then
    If BlockEndline < BlockStartLine Then
    If glistN.SelStart >= BlockEndCursor Then ClickInsideMarkedText = True
    Else
    If glistN.SelStart >= BlockStartCursor Then ClickInsideMarkedText = True
    End If
ElseIf glistN.ListIndex = BlockBottomline Then
    If BlockEndline < BlockTopline Then
    If glistN.SelStart <= BlockStartCursor Then ClickInsideMarkedText = True
    Else
    If glistN.SelStart <= BlockEndCursor Then ClickInsideMarkedText = True
    End If
ElseIf glistN.ListIndex < BlockTopline Or glistN.ListIndex > BlockBottomline Then
' FALSE
Else
    ClickInsideMarkedText = True
End If
ElseIf BlockTopline = glistN.ListIndex Then
If BlockEndCursor < BlockStartCursor Then
If glistN.SelStart >= BlockEndCursor And glistN.SelStart <= BlockStartCursor Then ClickInsideMarkedText = True
Else
If glistN.SelStart >= BlockStartCursor And glistN.SelStart <= BlockEndCursor Then ClickInsideMarkedText = True
End If

End If
End If
End Function
Sub RelocateMarkedText()
' when we resize glist, need to calculate the position of marked text
' we use a second set of variables
'  BlockStartPara As Long, BlockStartPos As Long
' BlockEndPara As Long, BlockEndPos As Long
'
BlockStartLine = mDoc.FirstParagraphLine(BlockStartPara) + mDoc.TextParagraph2Line(BlockStartPara, BlockStartPos, BlockStartCursor)
BlockEndline = mDoc.FirstParagraphLine(BlockEndPara) + mDoc.TextParagraph2Line(BlockEndPara, BlockEndPos, BlockEndCursor)
If BlockEndline < BlockStartLine Then
    BlockTopline = BlockEndline
    BlockBottomline = BlockStartLine
Else
    BlockTopline = BlockStartLine
    BlockBottomline = BlockEndline
End If
End Sub
Private Sub glistN_MarkOut()
'
BlockEndPara = mDoc.MarkParagraphID
BlockEndPos = ParaSelStart
BlockEndline = glistN.ListIndex
BlockEndCursor = glistN.SelStart
If BlockEndline < BlockStartLine Then
    BlockTopline = BlockEndline
    BlockBottomline = BlockStartLine
Else
    BlockTopline = BlockStartLine
    BlockBottomline = BlockEndline
End If

End Sub

Private Sub glistN_ReadListItem(item As Long, content As String)
If item >= 0 Then

content = mDoc.TextLine2(item + 1, showparagraphWork And showparagraph)

End If
End Sub



'End Sub

Private Sub glistN_ScrollSelected(item As Long, Y As Long)
If NoMark Then Exit Sub
Dim what$, ThatPara As Long
   mDoc.TextParagraphLine item, what$, ThatPara
   mDoc.MarkParagraphID = ThatPara
End Sub

Private Sub glistN_Selected(item As Long)
       If undomode = 1 And undotext <> "" Then PushUndo
    
       
If NoMark Then Exit Sub

EditDoc = EditDoc
Dim what$, ThatPara As Long
   mDoc.TextParagraphLine2 item, what$, ThatPara, showparagraphWork And showparagraph
   mDoc.MarkParagraphID = ThatPara
End Sub


Private Sub glistN_Selected2(item As Long)
If undomode = 1 And undotext <> "" Then PushUndo
If ClickInsideMarkedText Then
' start drag
glistN.MarkNext = 4

glistN.DragNow
End If
End Sub

Private Sub glistN_ShowExternalCursor()
glistN.PrepareToShow 5
End Sub



Private Sub glistN_SplitLine()
'
Dim b$, dummy As Long, a As Long

If glistN.ListIndex >= 0 Then
    b$ = mDoc.TextParagraph(mDoc.MarkParagraphID)
If Charpos = 1 And b$ <> "" Then
    dummy = mDoc.InsertBeforeParagraph(mDoc.MarkParagraphID, "")
        mDoc.ColorThis (dummy)
     all = mDoc.DocLines
    mDoc.MarkParagraphID = dummy
    SetRowColumn glistN.ListIndex + 2, 0

    
  ''
   
ElseIf ParaSelStart = 1 Then
    mDoc.MarkParagraphID = mDoc.InsertAfterParagraph(mDoc.MarkParagraphID, "")
    all = mDoc.DocLines
    ParaSelStart = 1

ElseIf ParaSelStart > Len(b$) Then

    mDoc.MarkParagraphID = mDoc.InsertAfterParagraph(mDoc.MarkParagraphID, "")
    
      all = mDoc.DocLines
    ParaSelStart = Len(b$)

Else  ' ANYWHERE

dummy = ParaSelStart

    mDoc.ReWritePara mDoc.MarkParagraphID, Left$(b$, ParaSelStart - 1), True    ' write and format
    
    dummy = mDoc.InsertAfterParagraph(mDoc.MarkParagraphID, Mid$(b$, dummy))
'    mDoc.startfrom(mDoc.MarkParagraphID) = 0
If mDoc.startfrom(mDoc.MarkParagraphID) <> 0 Then
    mDoc.ColorThis (mDoc.MarkParagraphID)
    If mDoc.startfrom(mDoc.MarkParagraphID) <> mDoc.startfrom(dummy) Then
          mDoc.ColorThis (dummy)
    End If
Else
    mDoc.ColorThis (mDoc.MarkParagraphID), , False
    mDoc.ColorThis (dummy)

End If
 mDoc.MarkParagraphID = dummy
    all = mDoc.DocLines
    ParaSelStart = 1

End If
End If
End Sub


Private Sub glistN_SyncKeyboard(item As Integer)
If NoMark Then item = 0
If blockKeyboard Then item = 0
If item > 32 And item <> 127 Then GetUndoNext = True
End Sub


Private Sub mDoc_BreakLine(data As String, datanext As String)
 ' Nowrap = True mean no Wrap
 Dim data1 As String, datanext1 As String
 If nowrap Then
    datanext = data
    data = vbNullString
 Else
    If showparagraph Then
         data1 = Replace(Replace(data, " ", ChrW(&HB7)), ChrW(&H2007), ChrW(&H2423))
        
        glistN.BreakLine data1, datanext1, , ChrW(&HB7)
        If data1 <> "" Then
        
        datanext = Left$(data, Len(datanext1))
        data = Right$(data, Len(data1))
        Else
            datanext = data
             data = vbNullString
        End If
      
    Else
        glistN.BreakLine data, datanext
    End If
End If

End Sub
Private Sub makebase(basecolor As Long)
ReDim col(1)
col(0) = basecolor
col(1) = 1

End Sub
Private Sub addone(ByVal thispos As Long, ByVal thiscolor As Long)
Dim k As Long

k = UBound(col())
If VarType(col(0)) = 0 Then
makebase 14
If thispos <> 1 Then
k = UBound(col())
Else
col(0) = thiscolor
Exit Sub
End If
End If
' LOOK THIS SECOND TIME
If k > 1 Then
If col(k - 1) = thiscolor Then Exit Sub
If col(k) = thispos Then
col(k - 1) = thiscolor: Exit Sub
End If
End If

''''If col(k - 1) = thiscolor And k > 1 Then Exit Sub

ReDim Preserve col(k + 2)
col(k + 1) = thiscolor
col(k + 2) = thispos


End Sub
Private Sub ColorizeInnerStatements(a$, k As Long, pos2 As Long, feedback As Long, Optional jump As Boolean = False)
Dim drop$, pos1 As Long, onetime As Boolean, kind As Long, lastkind As Long, feedthis As Long, counter As Long
Dim lastcol As Long
a$ = Replace(a$, vbTab, " ")
If jump Then
kind = IsLabelOnly((a$), drop$)
If kind = 0 Then
addone pos2, 14
lastcol = 14
If FastSymbol(a$, "(") Then
 pos2 = k - Len(a$) + 1
End If
GoTo again2
End If
GoTo fromjump

End If
onetime = True

If k < 1 Then
makebase 14
mDoc.Paracolordata(thispara) = col()
If mDoc.startfrom(thispara) = thispara Then
feedback = thispara
ReDim col(1)
Do While mDoc.Advance(thispara)
If mDoc.startfrom(thispara) = feedback Then
mDoc.startfrom(thispara) = 0
mDoc.Paracolordata(thispara) = col()
End If
Loop
feedback = feedback <> thispara
Exit Sub
End If

End If
If feedback > 0 Then
'lastcol = 15
a$ = Mid$(a$, feedback + 1)
'pos2 = feedback + 1
 addone feedback + 1, lastcol
feedback = 0
GoTo checkthistoo
Else
lastcol = 14
End If
again:
If RTrim$(a$) = vbNullString Then
If pos2 > 0 Then
If col((UBound(col()) \ 2) * 2) <> lastcol Then
addone pos2 + 1, lastcol
ElseIf UBound(col()) < 2 Then
addone pos2, lastcol
End If
  Else
  makebase 14
  End If
mDoc.Paracolordata(thispara) = col()
Exit Sub
' FOR TESTS
  pos1 = 0
  For k = 1 To UBound(col()) Step 2
    If VarType(col(k)) = 0 Then mDoc.Paracolordata(thispara) = col(): Exit Sub
    If col(k) > pos1 Then pos1 = col(k)
    If col(k) < pos1 Then mDoc.Paracolordata(thispara) = col(): Exit Sub
  Next k
    For k = 0 To UBound(col()) Step 2
   If col(k) > 15 Then mDoc.Paracolordata(thispara) = col(): Exit Sub
  Next k
  
End If
If MaybeIsSymbol(a$, ":", "\") Then
a$ = Mid$(a$, 2)
If AscW(Left$(a$, 1) + " ") = 92 Then
If pos2 > 1 Then
    If col(UBound(col())) = 1 Or InStr(" #@.+-~^%$,!({'?|/&)}" + Chr$(34), Mid$(lastdata, pos2 - 1, 1)) > 0 Then

    If pos2 > 0 Then
        addone pos2, 14
    Else
        addone 1, 14
    End If
 pos2 = k - Len(a$) + 1
GoTo novar
End If
End If
a$ = Mid$(a$, 2)
End If

If Len(a$) > 1 Then
If AscW(a$) <> 32 Then
lastkind = 1
GoTo again3

End If
Else
'lastkind = 1
'GoTo again3
End If
Else
novar:
While FastSymbol(a$, ":")
If pos2 > 0 Then
addone pos2, 14
Else
addone 1, 14
End If
 pos2 = k - Len(a$) + 1
Wend
End If
drop$ = "@(": kind = Abs(FastSymbol(a$, drop$, , 2)) * 4
If kind = 0 Then
If IsSymbol(a$, "@") Then
        If pos2 > 0 Then
        addone pos2, 14
        Else
        addone 1, 14
        End If
     pos2 = k - Len(a$) + 1
End If
End If
again500:
kind = IsLabelDot("", a$, drop$)

If kind = 0 Then drop$ = "$(": kind = Abs(FastSymbol(a$, drop$, , 2)) * 4
If kind = 0 Then drop$ = "~(": kind = Abs(FastSymbol(a$, drop$, , 2)) * 4
If kind = 0 Then drop$ = "->": kind = Abs(FastSymbol(a$, drop$, , 2))
fromjump:
If kind = 0 Then drop$ = vbNullString
If kind > 0 Then
    If ismine(drop$) Then
                lastcol = 11
                If pos2 > 0 Then
                   addone pos2, 11
                Else
                   addone k - Len(a$) - Len(drop$) + 1, 11
                   End If
        '' for the end
                If ismine2(drop$) Then
              
                kind = -2
                If Trim$(a$) <> "" Then
                                        pos2 = k - Len(a$) + 1
                                        If ismine22(drop$) Then
                                        
                                        ElseIf MaybeIsSymbol(a$, "{") Then
                                        
                                            addone pos2, 14
                                            pos2 = k - Len(a$) + 1
                                            onetime = True
                                            GoTo again
                                        ElseIf ismine5(drop$) Then
                                        GoTo again500
                                        ElseIf ismine3(drop$) Then
                                            onetime = True
                                            GoTo again
                                       
                                        End If
                                End If
             
        
                ElseIf ismine1(drop$) Then
                  
                kind = -1
                                If Trim$(a$) <> "" Then
                                                  pos2 = k - Len(a$) + 1
                                        If FastSymbol(a$, "{") Then
                                    
                                                  addone pos2, 14
                                              
                                                  pos2 = k - Len(a$) + 1
                                                  
                                                  onetime = True
                                                  GoTo again
                                          Else
                                           mDoc.Paracolordata(thispara) = col()
                                                  Exit Sub
                                        End If
                                End If
                ElseIf ismine33(drop$) Then
                ' fix for case >100, =50...
                pos2 = k - Len(a$) + 1
                Do
                lastcol = 14
                
                
                addone pos2, 14
                If FastSymbol(a$, "<") Then
                    If Not FastSymbol(a$, "=") Then
                        FastSymbol a$, ">"
                    End If
                ElseIf FastSymbol(a$, ">") Then
                    FastSymbol a$, "="
                ElseIf FastSymbol(a$, "=") Then
                    FastSymbol a$, "="
                ElseIf FastSymbol(a$, "~") Then
                
                End If
                
                pos1 = 1
                drop$ = aheadstatus(a$, False, pos1)
                If pos1 > 0 Then
                drop$ = Mid$(a$, 1, pos1 - 1)
                    ColorizeExpressions drop$, k - Len(a$) + Len(drop$), pos2, feedthis
                a$ = Mid$(a$, pos1)
                End If
                Loop Until Not FastSymbol(a$, ",")
                onetime = True
                
                GoTo here1
                End If
                 Else
                lastcol = 10
             If Len(a$) > 0 Then
             
              If AscW(a$) = 92 Then
             lastkind = 1
         
              End If
              End If
                If pos2 > 0 Then
                        
         
                
                
                 addone pos2, 10
                Else
                lastkind = 0
               ' If drop$ = "" Then Stop
                 addone k - Len(a$) - Len(drop$) + 1, 10
                
                 End If
             
  End If
  pos2 = k - Len(a$) + 1
  If kind > 3 Then
CheckThis:
  pos1 = 1
 drop$ = aheadstatus(a$, False, pos1)
marvel:

  If pos1 > 1 Then
marvel2:
drop$ = Mid$(a$, 1, pos1 - 1)
 ColorizeExpressions drop$, k - Len(a$) + Len(drop$), pos2, feedthis
 If feedthis = 11 Or feedthis = 14 Or feedthis = 10 Then
 'If kind < 5 Then lastcol = feedthis
feedthis = 0
End If
newjump:
 feedback = feedthis
 If col((UBound(col()) \ 2) * 2) = 7 Then
 a$ = vbNullString
 Else
 
  a$ = Mid$(a$, pos1)
  End If
checkthistoo:
If RTrim$(a$) = vbNullString Then
If pos2 > 0 Then
If feedthis <> 15 Then

  If col((UBound(col()) \ 2) * 2) <> lastcol And pos2 < pos1 Then addone pos2, lastcol

  End If
  Else
  makebase 14
  End If
  
mDoc.Paracolordata(thispara) = col()
If feedthis = 15 Then GoTo use15

Exit Sub
End If
If kind > 4 And onetime Then
If FastSymbol(a$, ",") Then
  addone pos2, 14
    pos2 = k - Len(a$) + 1
  pos1 = 1
  GoTo CheckThis
  End If
End If
If FastSymbol(a$, "->", , 2) Then
addone pos2, 14
    pos2 = k - Len(a$) + 1
If FastSymbol(a$, "{") Then
    pos2 = k - Len(a$) + 1
    onetime = True
    GoTo again
End If
  ElseIf FastSymbol(a$, ",") Then
  addone pos2, 14
    pos2 = k - Len(a$) + 1
  pos1 = 1

  End If
ElseIf kind > 4 And onetime Then
If FastSymbol(a$, ",") Then
  addone pos2, 14
    pos2 = k - Len(a$) + 1
  pos1 = 1
  GoTo CheckThis
  End If

  End If

    If FastSymbol(a$, ")") Then
  addone pos2, lastcol
    pos2 = k - Len(a$) + 1
  End If
  End If


here3:
If kind > 0 Then
If MaybeIsSymbol(a$, ".") Then
If FastSymbol(a$, ".(", , 2) Then
addone pos2, 14
a$ = vbNullString
    pos2 = k + 1
    
ElseIf FastSymbol(a$, "..", , 2) Then
    lastcol = 10
    lastkind = 1
    addone pos2, lastcol
    pos2 = k - Len(a$) + 1
    GoTo here3
Else
    pos1 = pos1 + 1
    lastcol = 14
    GoTo CheckThis
  End If
End If
End If
If FastSymbol(a$, "<=", , 2) Then
  addone pos2, 15
    pos2 = k - Len(a$) + 1
    Else
  
again4:
  If MaybeIsSymbolNoSpace(a$, "~", True) Then
    addone pos2, 15
    pos2 = k - Len(a$) + 1
  ElseIf FastSymbol(a$, "=>", , 2) Then
    addone pos2, 14
  
    pos2 = k - Len(a$) + 1
  ElseIf FastSymbol(a$, "=") Then
  If onetime And lastcol = 11 Then
    
  Else
    addone pos2, 15
  End If
    pos2 = k - Len(a$) + 1
  ElseIf MaybeIsSymbol(a$, "+-*/", "=", True) Then
  addone pos2, 15
    pos2 = k - Len(a$) + 1
  ElseIf FastSymbol(a$, "/*-+=~^|<>") Then
  addone pos2, 14
    pos2 = k - Len(a$) + 1
  If FastSymbol(a$, "/*-+=~^&|<>") Then
  addone pos2, 14
    pos2 = k - Len(a$) + 1
  End If
End If
  End If
again2:
Do
again3:
  pos1 = 1
  
  drop$ = aheadstatus(a$, False, pos1)
  
  If pos1 > 1 Or drop$ = "S" Then
  If pos1 = 1 Then
  drop$ = a$
  pos1 = Len(a$) + 1
  counter = 1
  Else
  drop$ = Mid$(a$, 1, pos1 - 1)
  If drop$ = vbNullString Then drop$ = " "
  End If
feedback = 0
 ColorizeExpressions drop$, k - Len(a$) + Len(drop$), pos2, feedback
If feedback <> 15 Then
If feedback = 0 And lastcol = 0 Then feedback = 14
If drop$ = "(" Then
If feedback = 100 Then
    lastcol = 14
Else
lastcol = feedback
End If
Else 'If lastcol <> 11 And lastcol <> 10 Then
If feedback = 100 Then
    lastcol = 14
Else
If feedback > 0 And kind > 0 Then
lastcol = feedback
feedback = 0

a$ = Mid$(a$, pos1)

ColorizeInnerStatements a$, k, pos2, feedback, True

If MaybeIsSymbol(a$, ")") Then
addone pos2, lastcol
FastSymbol a$, ")"
Exit Sub
End If
End If
End If
End If
feedback = 0
End If
 If feedback = 100 Then
  a$ = Mid$(a$, pos1)
 GoTo again112
 ElseIf feedback = 15 Then
use15:
 counter = 1
nextline:
 feedback = 0
 
 drop$ = mDoc.NextTextParagraph(thispara, feedback)
  If feedback = 0 Then
  a$ = Mid$(a$, pos1)
 mDoc.Paracolordata(thispara) = col()
  Exit Sub
 End If
 pos1 = 1
 countbrakcet drop$, pos1, counter
 mDoc.Paracolordata(thispara) = col()
  a$ = drop$
 'If feedthis = 15 Then
 k = Len(a$)
 'Else
 'k = Len(a$) - 1
 'End If
 If mDoc.startfrom(thispara) = 0 Then
 mDoc.startfrom(thispara) = thispara
  End If
  mDoc.startfrom(feedback) = mDoc.startfrom(thispara)
 thispara = feedback
 
    
    lastdata = a$
    pos2 = 1
    If Len(a$) > 0 Then
    ReDim col(3)
             col(0) = 15
            col(1) = 1
             col(2) = 15
             col(3) = k + 1
             Else
    ReDim col(1)
             col(0) = 15
            col(1) = 1
             End If
If counter = 0 Then
If UBound(col()) > 1 Then
            
            col(3) = pos1
            If col(3) = col(1) Then
            pos1 = pos1 + 1
            col(3) = pos1
            col(2) = 14
            End If
            Else
            col(1) = pos1
            End If
If Left(a$, 1) = "}" Then
Else
a$ = Mid$(a$, pos1 + 1)
pos2 = k - Len(a$) + 1
End If
feedback = 0 'lastcol
feedthis = 0
 
Else
    GoTo nextline

End If
    'kind = 0
   ' feedback = 0

            pos1 = 1

GoTo CheckThis
ElseIf feedback = 0 And pos1 > 1 And kind > 0 Then

GoTo newjump

 End If
 If pos1 > 0 Then
    If Mid$(a$, pos1 - 1, 1) = " " Then lastkind = 0
 
  a$ = Mid$(a$, pos1)
End If
If MaybeIsSymbol(a$, ":{") Then
   
    onetime = True
    GoTo again
End If
If InStr(drop$, " ") = 0 Or kind >= 0 Then
GoTo CheckThis

End If
  End If
again112:
    If MaybeIsSymbol(a$, ",;", , True) Then
     addone pos2, 14
    pos2 = k - Len(a$) + 1
    lastkind = 0
    onetime = False
    ElseIf FastSymbol(a$, "}") Then
    lastkind = 0
         addone pos2, 14
    pos2 = k - Len(a$) + 1
    onetime = True
    If MaybeIsSymbol(a$, ".;+(") Then
GoTo CheckThis
End If
    GoTo again112
    ElseIf MaybeIsSymbol(a$, "'") Then
        lastkind = 0
         pos2 = k - Len(a$) + 1
        FastSymbol a$, "'"
        
        If pos2 > 0 Then
         addone pos2, 7
         
         Else
        addone 1, 7
        End If
        a$ = vbNullString
        pos2 = k
        If pos2 > 0 Then
         addone pos2 + 1, 14
        End If
   mDoc.Paracolordata(thispara) = col()
   Exit Sub
    Else
   onetime = True
    mDoc.Paracolordata(thispara) = col()
   Exit Do
   End If
   Loop
'If a$ = vbNullString Then
'
 'mDoc.Paracolordata(thispara) = col()
 'Exit Sub

'End If
GoTo again
ElseIf MaybeIsSymbol(a$, "\", , True) Then

If lastkind = 1 Then
 pos2 = k - Len(a$) + 1
 If AscW(a$ + ".") <> 32 Then
 If MaybeIsSymbol(a$, ".", ".", True) Then
 
 End If
 If MaybeIsSymbolNoSpace(a$, "\") Then GoTo again
    pos2 = k - Len(a$) + 1
    
    lastkind = 0
  If MaybeIsSymbol(a$, "\", , True) Then GoTo here1
  GoTo again
 
  End If
 Else
here1:
 
 If pos2 > 0 Then
     addone pos2, 7
    
    Else
    addone 1, 7
    End If
 pos2 = k - Len(a$) + 1
 mDoc.Paracolordata(thispara) = col()
 Exit Sub
 End If

ElseIf MaybeIsSymbol(a$, "'") Then

    If pos2 > 1 Then
         addone pos2, 7
         Else
    
    addone 1, 7
    End If
 pos2 = k - Len(a$) + 1
 mDoc.Paracolordata(thispara) = col()
 Exit Sub

ElseIf Not onetime Then
If FastSymbol(a$, ",") Then
mDoc.Paracolordata(thispara) = col()
Exit Sub  ' stop
ElseIf FastSymbol(a$, ":") Then
onetime = True
ElseIf FastSymbol(a$, ")") Then
addone pos2, lastcol
pos2 = k - Len(a$) + 1
    GoTo again4
ElseIf FastSymbol(a$, ";") Then
addone pos2, 14
    pos2 = k - Len(a$) + 1
    GoTo again2
ElseIf FastSymbol(a$, "}") Then
addone pos2, 14
addone k - Len(a$), 11
pos2 = k - Len(a$) + 1
    GoTo again3

ElseIf FastSymbol(a$, "{") Then
onetime = True
ElseIf FastSymbol(a$, "?") Then
addone pos2, 14
addone k - Len(a$), 11
pos2 = k - Len(a$) + 1
GoTo again2
Else
addone pos2, lastcol
mDoc.Paracolordata(thispara) = col()
Exit Sub
End If
addone pos2, 14
    pos2 = k - Len(a$) + 1
    GoTo again
ElseIf IsSymbol(a$, "?") Then

addone 1, 11
        pos2 = k - Len(a$) + 1
        GoTo again2
ElseIf IsSymbol(a$, "=") Then

If pos2 > 0 Then
addone pos2, 11
Else
addone 1, 11
End If
        pos2 = k - Len(a$) + 1
        onetime = False
        GoTo again2
ElseIf FastSymbol(a$, "{") Then
onetime = True
If pos2 > 0 Then
If col((UBound(col()) \ 2) * 2) <> 14 Then addone pos2, 14
'addone pos2, 14
 pos2 = k - Len(a$) + 1
    If FastSymbol(a$, "&") Then
          addone pos2, 10
      pos2 = k - Len(a$) + 1
End If
Else
addone 1, 14
    pos2 = k - Len(a$) + 1
End If
    
        GoTo again

ElseIf FastSymbol(a$, "}") Then
If pos2 > 0 Then
addone pos2, 14
Else
addone 1, 14
If MaybeIsSymbol(a$, "(") Then
GoTo CheckThis
ElseIf MaybeIsSymbol(a$, ",;") Then
a$ = Mid$(a$, 2)

 GoTo CheckThis
End If
End If
        pos2 = k - Len(a$) + 1
        GoTo again

ElseIf pos2 = 0 Then
If IsNumberLabel(a$, drop$) Then
        addone k - Len(a$) - Len(drop$) + 1, 15
        pos2 = k - Len(a$) + 1
        GoTo again
 End If
 End If
If pos2 > 0 Then addone pos2, 14

End Sub
Private Sub ColorizeExpressions(ByVal a$, k As Long, pos2 As Long, feedthis As Long)
Dim drop$, pos1 As Long, kind As Long, localfeedthis As Long, thispara1 As Long, lastpar As Long
Dim lastcol As Long
Dim n$
again:
If a$ = vbNullString Then
If feedthis = 0 Then
If lastpar > 0 Then
feedthis = lastpar
End If
End If
Exit Sub
End If
drop$ = "@(": kind = Abs(FastSymbol(a$, drop$, , 2)) * 4
If kind = 0 Then
    If IsSymbol(a$, "@") Then
        If pos2 > 0 Then
            addone pos2, 14
        Else
            addone 1, 14
        End If
        pos2 = k - Len(a$) + 1
    
    End If
getlabelonly:
    n$ = a$
    kind = IsLabelOnly(n$, drop$)
    pos2 = k - Len(a$) + 1
    If kind = 0 And n$ <> a$ Then
    feedthis = 100
    a$ = n$: Exit Sub
    End If
    a$ = n$
    If kind = 100 Then
    If pos2 > 0 Then
            addone pos2, 14
        Else
            addone 1, 14
        End If
         pos2 = k - Len(a$) + 1
         
    kind = 0
    GoTo again
    
    End If
End If
If kind = 0 Then drop$ = "$(": kind = Abs(FastSymbol(a$, drop$, , 2)) * 4
If kind = 0 Then drop$ = "~(": kind = Abs(FastSymbol(a$, drop$, , 2)) * 4
If kind = 0 Then
  If FastSymbol(a$, "->", , 2) Then
    pos2 = k - Len(a$) - 1
    addone pos2, 11
        a$ = NLtrim$(a$)
        pos2 = k - Len(a$) + 1
        If MaybeIsSymbolNoSpace(a$, "{") Then
        pos2 = k - Len(a$) + 1
        addone pos2, 14
        a$ = Mid$(a$, 2)
         localfeedthis = 0
           pos2 = k - Len(a$) + 1
           If pos2 <= k Then
                drop$ = block(a$)
                If Len(drop$) > 0 Then
                    ColorizeInnerStatements drop$, k - Len(a$), pos2, localfeedthis
           
                    If Len(a$) > 0 Then ColorizeInnerStatements a$, k, pos2, localfeedthis
                    drop$ = vbNullString
                Else
                ColorizeInnerStatements a$, k, pos2, localfeedthis
               End If

           End If
            kind = 0
            drop$ = vbNullString
            If pos2 > k Then Exit Sub
        Else
           GoTo checklambda
            
        End If
    
    End If
End If
If kind = 0 Then drop$ = vbNullString
If kind > 0 Then
    If ismine(drop$) Then
    lastcol = 11
    If kind > 4 Then lastpar = 11
    If pos2 > 0 Then
       addone pos2, 11
    Else
       addone k - Len(a$) - Len(drop$) + 1, 11
       End If
        '' for the end
   Else
   If kind > 4 And kind < 8 Then lastpar = 10
   If kind = 10 Then
        lastcol = 13
   ElseIf kind = 20 Then
        lastcol = 12
   Else
        lastcol = 10
   End If
   
   If pos2 > 0 Then
       addone pos2, lastcol
   Else
       addone k - Len(a$) - Len(drop$) + 1, lastcol
      
       End If
  End If
  pos2 = k - Len(a$) + 1
 If kind > 3 Then
checklambda:
  Do
  pos1 = 1
  drop$ = aheadstatus(a$, False, pos1)
  If pos1 > 1 Then
  ColorizeExpressions Mid$(a$, 1, pos1 - 1), k - Len(a$) + pos1 - 1, pos2, feedthis
  If feedthis = 14 Or feedthis = 11 Then
  If lastcol <> 11 And lastcol <> 10 Then lastcol = feedthis
  feedthis = 0
  End If
  If feedthis > 0 And pos1 > Len(a$) Then Exit Sub
  feedthis = 0
   a$ = Mid$(a$, pos1)
   End If
   If FastSymbol(a$, ",") Then
     addone pos2, 14
    pos2 = k - Len(a$) + 1
   Else
   Exit Do
   End If
   Loop
    If MaybeIsSymbol(a$, "(") Then
addone pos2, 14
lastpar = 14
FastSymbol a$, "("
pos2 = k - Len(a$) + 1
  pos1 = 1
  drop$ = aheadstatus(a$, False, pos1)
    thispara1 = thispara
  If pos1 > 1 Then

  ColorizeExpressions Mid$(a$, 1, pos1 - 1), k - Len(a$) + pos1 - 1, pos2, feedthis
    a$ = Mid$(a$, pos1)
    End If
    If thispara1 <> thispara Then
    Exit Sub
    End If
  addone pos2, 14
FastSymbol a$, ")"
pos2 = k - Len(a$) + 1
ElseIf IsSymbol(a$, ",") Then
addone pos2, 14
pos2 = k - Len(a$) + 1
    GoTo again
ElseIf MaybeIsSymbol(a$, ")") Then
    pos2 = k - Len(a$) + 1
    If kind >= 10 Then
    addone pos2, 14
    
    ElseIf kind > 0 Then
  addone pos2, lastcol
  Else
  addone pos2, 14
  End If
  pos2 = k - Len(a$) + 2
  addone pos2, 14
  FastSymbol a$, ")"
   
    pos2 = k - Len(a$) + 1
  End If
  End If

GoTo again
ElseIf MaybeIsSymbol(a$, "(", , True) Then
    addone pos2, 14
    pos2 = k - Len(a$) + 1
    'lastcol = 14
   ' ColorizeExpressions a$, k - Len(a$) + 1, pos2, feedthis
   ' a$ = Mid$(a$, pos2)
   GoTo again
ElseIf MaybeIsSymbol(a$, "{") Then
addone pos2, 15
pos1 = 1
blockStringAhead a$, pos1
If Mid$(a$, pos1, 1) <> "}" Then feedthis = 15 Else feedthis = 0
a$ = Mid$(a$, pos1 + 1)
drop$ = vbNullString
pos2 = k - Len(a$) + 1
If feedthis = 15 Then Exit Sub
    GoTo again

ElseIf MaybeIsSymbol(a$, Chr(34)) Then

If Not ISSTRINGA(a$, drop$) Then
a$ = vbNullString
addone pos2, 10
Else
addone pos2, 15
End If
drop$ = vbNullString
pos2 = k - Len(a$) + 1
    GoTo again
Else
If col((UBound(col()) \ 2) * 2) <> 14 Then addone pos2, 14
a$ = Mid$(a$, 2)
If Len(a$) > 0 Then GoTo again
End If

If pos2 > 0 Then addone pos2, 14


End Sub


Private Sub mDoc_ColorDataEvent(thispara1 As Long, data As String, feedback As Long)

thispara = thispara1
col() = mDoc.Paracolordata(thispara)
If NoColor Then
If UBound(col()) > 1 Then
ReDim col(1)
mDoc.Paracolordata(thispara) = col()
End If
Exit Sub
End If
If feedback = 0 Then
ReDim col(1)
End If
lastdata = data
ColorizeInnerStatements (data), Len(data), (0), feedback
If feedback = -1 Then
If thispara <> 0 Then
mDoc.Paracolordata(thispara) = col()
thispara1 = thispara
End If
Else
mDoc.Paracolordata(thispara) = col()
thispara1 = thispara
End If
End Sub




Function IsLabelDot(where$, a$, r$) As Long 'ok
' for left side...no &

Dim rr&, one As Boolean, c$, firstdot$, gr As Boolean
r$ = vbNullString
If a$ = vbNullString Then IsLabelDot = 0: Exit Function

a$ = NLtrim$(a$)
    Do While Len(a$) > 0
    c$ = Left$(a$, 1)
    If AscW(c$) < 256 Then
        Select Case AscW(c$)
        Case 64  '"@"
        
            If Mid$(a$, 2, 1) <> "(" And r$ <> "" Then

                r$ = r$ + "@"
                a$ = Mid$(a$, 2)
            Else
                IsLabelDot = 0: a$ = firstdot$ + a$: Exit Function
            End If
        Case 63 '"?"
        If r$ = vbNullString And firstdot$ = vbNullString Then
        r$ = "?"
        a$ = Mid$(a$, 2)
        IsLabelDot = 1
        Exit Function
    
        ElseIf firstdot$ = vbNullString Then
        IsLabelDot = 1
        Exit Function
        Else
        IsLabelDot = 0
        Exit Function
        End If
        Case 46 '"."
            If one Then
            Exit Do
            Exit Do
            ElseIf r$ <> "" And Len(a$) > 1 Then
            If Mid$(a$, 2, 2) = ". " Or Mid$(a$, 2, 1) = " " Then Exit Do
            r$ = r$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1
            Else
            firstdot$ = firstdot$ + "."
            a$ = Mid$(a$, 2)
            End If
       Case 92, 94, 123 To 126 '"\","^", "{" To "~"
        Exit Do

        Case 48 To 57, 95 '"0" To "9", "_"
           If one Then
            If firstdot$ <> "" Then a$ = firstdot$ + a$
            Exit Do
            ElseIf r$ <> "" Then
            r$ = r$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1 'is an identifier or floating point variable
            Else
            Exit Do
            End If
        Case Is < 0, Is > 64 ' >=A and negative
            If one Then
            Exit Do
            Else
            r$ = r$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1 'is an identifier or floating point variable
            End If
        Case 36 ' "$"
        
            If one Then
            If InStr(r$, "$.") Then
            rr& = 3          ' is string variable with string pointer
            r$ = r$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            End If
            Exit Do
            End If
            If r$ <> "" Then
            one = True
            If Mid$(a$, 2, 1) <> "." Then
            rr& = 3
            r$ = r$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            Else
            rr& = 1       ' is string variable
            r$ = r$ & Left$(a$, 2)
            a$ = Mid$(a$, 3)
            
            End If
            
            Else
            Exit Do
            End If
        Case 37 ' "%"
            If one Then Exit Do
            If r$ <> "" Then
            one = True
            rr& = 4 ' is long variable
            r$ = r$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            Else
            Exit Do
            End If
        Case 40 ' "("
            If r$ <> "" Then
                            If Mid$(a$, 2, 2) = ")@" Then
                                    r$ = r$ & "()."
                                  
                                 a$ = Mid$(a$, 4)
                               Else
                                       Select Case rr&
                                       Case 1
                                       rr& = 5 ' float array or function
                                       Case 3
                                       rr& = 6 'string array or function
                                       Case 4
                                       rr& = 7 ' long array
                                       Case Else
                                       Exit Do
                                       End Select
                                       r$ = r$ & Left$(a$, 1)
                                       a$ = Mid$(a$, 2)
                                   Exit Do
                            
                          End If
               Else
                        Exit Do
            
            End If
        Case Else
        Exit Do
        End Select
        Else
            If one Then
            Exit Do
            Else
            gr = True
            r$ = r$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1 'is an identifier or floating point variable
            End If
        End If

    Loop
       r$ = firstdot$ + myUcase(r$, gr)
    IsLabelDot = rr&
   'a$ = LTrim$(a$)

End Function
Function IsLabelOnly(a$, r$) As Long 'ok

Dim rr&, one As Boolean, c$, dot&, Extra&, num As Boolean
r$ = vbNullString
If a$ = vbNullString Then IsLabelOnly = 0: Exit Function
a$ = NLtrim$(a$)
    Do While Len(a$) > 0
     c$ = Left$(a$, 1) 'ANYCHAR HERE
    If AscW(c$) < 256 Then
        Select Case AscW(c$)
                Case 35 ' "#"
        If r$ <> "" Then
        Exit Do
        Else
        
        r$ = "#"
           a$ = Mid$(a$, 2)
          
          rr& = 10
           End If
        Case 64  '"@"
           If r$ = vbNullString Then
              a$ = Mid$(a$, 2)
              ElseIf Mid$(a$, 2, 1) <> "(" And r$ <> "" Then
            r$ = r$ & "."
              a$ = Mid$(a$, 2)
              Else
                 IsLabelOnly = 0: Exit Function
            End If
        Case 46 '"."
            If one Then
                Exit Do
            ElseIf r$ <> "" Then
                r$ = r$ & Left$(a$, 1)
                a$ = Mid$(a$, 2)
            ElseIf Not Mid$(a$, 2, 1) Like "[0-9]" Then
                       If r$ <> "" Then
                            r$ = r$ & Left$(a$, 1)
                            rr& = 1
                        Else
                            dot& = dot& + 1
                        End If
                        If Mid$(a$, 2, 1) Like "[eEÂ≈]" Then
                            r$ = r$ & Mid$(a$, 2, 1)
                            a$ = Mid$(a$, 3)
                        Else
                             a$ = Mid$(a$, 2)
                        End If
            Else
                If r$ = vbNullString And dot& > 0 Then
                    r$ = String$(dot& + 1, ".")
                    a$ = Mid$(a$, 2)
                    IsLabelOnly = 1
                Else
                If MaybeIsSymbolNoSpace(Mid$(a$, 2, 1), "1234567890") Then
                    FastSymbol a$, "."
                    While MaybeIsSymbolNoSpace(a$, "1234567890", True)
                    Wend
                    If MaybeIsSymbolNoSpace(Mid$(a$, 3, 1), "Ee≈Â") Then
                        a$ = Mid$(a$, 4)
                        MaybeIsSymbolNoSpace a$, "+-", True
                        MaybeIsSymbolNoSpace a$, "123456789", True
                        While MaybeIsSymbolNoSpace(a$, "1234567890", True)
                        Wend
                        If MaybeIsSymbolNoSpace(a$, ".") Then
                            a$ = vbNullString
                        End If
                    End If
                    If MaybeIsSymbolNoSpace(a$, ".") Then
                        IsLabelOnly = 1
                    Else
                        IsLabelOnly = 100
                    End If
                    Exit Function
                End If
                IsLabelOnly = 0
            End If
            Exit Function
        End If
     Case 92, 94, 123 To 126 '"\","^", "{" To "~"
Exit Do
        
        Case 48 To 57, 95 '"0" To "9", "_"
       
           If one Then
           
            Exit Do
            ElseIf r$ <> "" Then
            r$ = r$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            If Extra& > 0 Then
            If Left$(a$, 1) = "%" Then
            If Len(r$) > 6 Then rr& = 1: Exit Do
            a$ = Mid$(a$, 2)
            Extra& = Extra& + 10
            rr& = 1: Exit Do
            
             ElseIf Left$(a$, 1) = "&" Then
             Extra& = Extra& + 10
            a$ = Mid$(a$, 2)
            rr& = 1: Exit Do
            End If
            If Len(r$) = 10 Then rr& = 1: Exit Do
            End If
               rr& = 1 'is an identifier or floating point variable
            ElseIf Mid$(a$ + "__", 2, 2) Like "[◊˜Xx][0-9A-Fa-f]" Then
             r$ = r$ & Left$(a$, 3)
            a$ = Mid$(a$, 4)
               rr& = 1
            Extra& = 9
            If Left$(a$, 1) = "%" Then
                a$ = Mid$(a$, 2)
                Extra& = Extra& + 10
                rr& = 1: Exit Do
            
             ElseIf Left$(a$, 1) = "&" Then
                 Extra& = Extra& + 10
                a$ = Mid$(a$, 2)
                rr& = 1: Exit Do
            End If
            ElseIf AscW(Mid$(a$, 2, 1) + " ") = 38 Then
            a$ = Mid$(a$, 3)
            Extra& = 100
            Exit Do
            Else
           
          If dot& > 0 Then
            a$ = "." + a$: dot& = 0
            ElseIf Mid$(a$, 2, 1) = "." Then
            a$ = Mid$(a$, 3)
checkexp:
              While MaybeIsSymbolNoSpace(a$, "1234567890", True)
              Wend

              If MaybeIsSymbolNoSpace(a$, "Ee≈Â", True) Then
                  MaybeIsSymbolNoSpace a$, "+-", True
                  If MaybeIsSymbolNoSpace(a$, "+-", True) Then
                  IsLabelOnly = 1
                  Exit Function
                  End If
                  MaybeIsSymbolNoSpace a$, "123456789", True
                  While MaybeIsSymbolNoSpace(a$, "1234567890", True)
                  Wend
                  If MaybeIsSymbolNoSpace(a$, ".") Then
                  IsLabelOnly = 1
                  a$ = vbNullString
                  Exit Function
                  End If
              Else
              While MaybeIsSymbolNoSpace(a$, "1234567890", True)
              Wend
              End If
            If MaybeIsSymbolNoSpace(a$, ".") Then
             
              IsLabelOnly = 1
              Else
              IsLabelOnly = 100
              End If
              
              Exit Function
            ElseIf MaybeIsSymbolNoSpace(a$, "1234567890", True) Then
              While MaybeIsSymbolNoSpace(a$, "1234567890", True)
              Wend
                  If MaybeIsSymbolNoSpace(a$, ".", True) Then
                                
                    GoTo checkexp
                  End If
                  MaybeIsSymbolNoSpace a$, "&", True
                  
                  
                IsLabelOnly = 100
                Exit Function
            End If
            Exit Do
            End If
Case Is < 0, Is > 64 ' >=A and negative

            If one Then
            Exit Do
            Else
  If Extra& > 0 Then
        If Not c$ Like "[a-fA-F]" Then Exit Do
        r$ = r$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            
            If Left$(a$, 1) = "%" Then
                If Len(r$) > 6 Then rr& = 1: Exit Do
                Extra& = Extra& + 10
                a$ = Mid$(a$, 2)
                rr& = 1: Exit Do
            
             ElseIf Left$(a$, 1) = "&" Then
             Extra& = Extra& + 10
                a$ = Mid$(a$, 2)
                rr& = 1: Exit Do
            End If
            If Len(r$) = 10 Then rr& = 1: Exit Do
        
        
  Else
            r$ = r$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
       
            rr& = 1 'is an identifier or floating point variable
  
  
  
  End If
              
        
  
            End If
      Case 38 ' "&"
            If r$ = vbNullString Then rr& = 2:    a$ = Mid$(a$, 2)
            Exit Do
       Case 36 ' "$"
       If one Then Exit Do
            If r$ <> "" Then
            one = True
            rr& = 3 ' is string variable
            r$ = r$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            Else
            Exit Do
            End If
        Case 37 ' "%"
      If one Then
            If InStr(r$, "$.") Then
            rr& = 3          ' is string variable with string pointer
            r$ = r$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            End If
            Exit Do
            End If
            If r$ <> "" Then
            one = True
            If Mid$(a$, 2, 1) <> "." Then
            rr& = 3
            r$ = r$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            Else
            rr& = 1       ' is string variable
            r$ = r$ & Left$(a$, 2)
            a$ = Mid$(a$, 3)
            
            End If
            
            Else
            Exit Do
            End If
        Case 40 ' "("
            If r$ <> "" Then
                If Mid$(a$, 2, 2) = ")@" Then
                 r$ = r$ & "()."
              a$ = Mid$(a$, 4)
                Else
                Select Case rr&
                Case 1
                rr& = 5 ' float array or function
                Case 3
                rr& = 6 'string array or function
                Case 4
                rr& = 7 ' long array
                Case Else
                Exit Do
                End Select
                r$ = r$ & Left$(a$, 1)
                a$ = Mid$(a$, 2)
               Exit Do
            
            End If
           Else
           Exit Do
            End If
        Case Else
        Exit Do
        End Select
        Else
            If one Then
            Exit Do
            Else
             r$ = r$ & Left$(a$, 1)
            a$ = Mid$(a$, 2)
            rr& = 1 'is an identifier or floating point variable
            End If
        End If
    Loop
    If rr& > 0 Then
        If Left$(r$, 1) = "#" Then
          If Len(r$) > 1 Then
            If Mid$(r$, 2) Like "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]" Then
                rr& = 10
            Else
                a$ = r$ + a$
                rr& = 0
            End If
         Else
        a$ = r$ + a$
        rr& = 0
        End If
        End If
    End If

    IsLabelOnly = rr& + Extra&
   
End Function


Function MaybeIsSymbol(a$, c$, Optional tail$ = vbNullString, Optional Remove As Boolean = False) As Boolean
Dim i As Long, ok As Boolean
If a$ = vbNullString Then Exit Function
i = 1
While Mid$(a$, i, 1) = " "
    i = i + 1
Wend
If i > Len(a$) Then Exit Function
If i > 1 Then a$ = Mid$(a$, i): i = 1

ok = InStr(c$, Mid$(a$, i, 1)) > 0
If tail$ <> "" And ok Then
ok = Mid$(a$, i + 1, Len(tail$)) Like tail$
End If
If Remove Then If ok Then a$ = Mid$(a$, i + 1 + Len(tail$))
MaybeIsSymbol = ok
End Function
Function MaybeIsSymbolNoSpace(a$, c$, Optional Remove As Boolean = False) As Boolean
Dim ok As Boolean

If a$ = vbNullString Then Exit Function
ok = InStr(c$, Left$(a$, 1))
If Remove And ok Then a$ = Mid$(a$, 2)
MaybeIsSymbolNoSpace = ok
End Function
Private Sub mDoc_Lastline()
If showparagraph Then showparagraphonce = True
End Sub

Private Sub mDoc1_BreakLine(data As String, datanext As String)
If NoHeadwrap Or (NoTitle) Or UsedAsTextBox Then
    datanext = data
    data = vbNullString
Else
 glistN.BreakLine data, datanext
End If
End Sub
Private Sub FillBack(thathDC As Long, there As RECT, bgcolor As Long)
' create brush
Dim my_brush As Long
my_brush = CreateSolidBrush(bgcolor)
FillRect thathDC, there, my_brush
DeleteObject my_brush
End Sub
Private Sub PrintItem1(mHdc As Long, c As String, r As RECT, Optional way As Long = DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS Or DT_TABSTOP)
Dim cc As String
'r.Top = r.Top + TextPixelOffset
If showparagraphonce Then
    cc = c & "∂"
    DrawTextEx mHdc, StrPtr(cc), -1, r, way, VarPtr(tParam)
Else

    DrawTextEx mHdc, StrPtr(c), -1, r, way, VarPtr(tParam)
End If
'r.Top = r.Top - TextPixelOffset
End Sub
Private Sub PrintItem(mHdc As Long, c As String, r As RECT, Optional way As Long = DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS Or DT_TABSTOP)
    'r.Top = r.Top + TextPixelOffset
    DrawTextEx mHdc, StrPtr(c), -1, r, way, VarPtr(tParam)
    'r.Top = r.Top - TextPixelOffset
End Sub
Private Sub PrintItem2(mHdc As Long, c As String, r As RECT, Optional way As Long = DT_SINGLELINE Or DT_NOPREFIX Or DT_NOCLIP Or DT_EXPANDTABS Or DT_TABSTOP)
    Dim that As Long
    that = way
    r.Left = r.Left + 4
    If glistN.VerticalCenterText Then that = that Or DT_VCENTER Or DT_EXPANDTABS Or DT_TABSTOP
    DrawTextEx mHdc, StrPtr(c), -1, r, that, VarPtr(tParam)
End Sub
    
Public Property Get EditDoc() As Boolean
EditDoc = mEditDoc
End Property

Public Property Let EditDoc(ByVal RHS As Boolean)
mEditDoc = RHS
If Not RHS Then glistN.MarkNext = 0
glistN.EditFlag = RHS
If glistN.enabled Then
glistN.ShowMe2
End If
End Property

Public Property Get ParaSelStart() As Long
If glistN.ListIndex < 0 Then ParaSelStart = 1: Exit Property
Dim what$, ThatPara As Long, pastchars
    pastchars = mDoc.TextParagraphLine(glistN.ListIndex + 1, what$, ThatPara)
Charpos = pastchars + glistN.SelStart
If Charpos = 0 Then
ParaSelStart = 1
Else
ParaSelStart = Charpos
End If
End Property
Public Sub WrapAll()
mDoc.WrapAgain
End Sub
Public Property Let ParaSelStart(ByVal RHS As Long)
' we have a new ParaSelStart
' we want a new listindex...and a new selstart

Charpos = RHS
Render
End Property

Public Property Get nowrap() As Boolean
nowrap = mNoWrap
End Property

Public Property Let nowrap(ByVal RHS As Boolean)
If Not (mNoWrap = RHS) Then
mNoWrap = RHS
glistN.NoScroll = Not mNoWrap
    If Not mDoc.IsEmpty Then
    WrapAll
   If HaveMarkedText Then RelocateMarkedText
    
    HorSrcoll 0
    Render
    End If
End If
End Property

Public Property Get enabled() As Boolean
enabled = glistN.enabled
End Property

Public Property Let enabled(ByVal RHS As Boolean)
If Not RHS Then
glistN.Noflashingcaret = True
Show

Else
If IsWine Then
glistN.Noflashingcaret = Not glistN.NoCaretShow And Not glistN.backcolor = rgb(255, 255, 355)
Else
glistN.Noflashingcaret = Not glistN.NoCaretShow
End If
End If
glistN.enabled = RHS
End Property
Sub SetFocus()
glistN.SetFocus
End Sub

Public Property Get ScrollKeyUpDown() As Boolean
ScrollKeyUpDown = glistN.NoFreeMoveUpDown
End Property

Public Property Let ScrollKeyUpDown(ByVal RHS As Boolean)
glistN.NoFreeMoveUpDown = RHS
End Property
Public Sub SetRowColumn(Row As Long, ByVal Column As Long, Optional noShow As Boolean = False)
' base 0 for column as input
Column = Column + 1
' base 1
If Column < 0 Then Column = 0
If Row = -1 Then
glistN.ListindexPrivateUse = glistN.listcount - 1
Else
If Row > glistN.listcount + 1 Then
Row = glistN.listcount + 1
End If
glistN.ListindexPrivateUse = Row - 1  ' no events
End If
Dim a As String, b As Long
mDoc.TextParagraphLine glistN.ListIndex + 1, a, b
mDoc.MarkParagraphID = b
If Len(glistN.list(glistN.ListIndex)) < Column Then
If Row = glistN.listcount + 1 Then
glistN.SelStartEventAlways = Len(glistN.list(glistN.ListIndex))
Else
glistN.SelStartEventAlways = Column  'Len(glistN.List(glistN.listindex)) + 1
End If
Else
glistN.SelStartEventAlways = Column
End If
If glistN.enabled Then
If Not noShow Then glistN.ShowMe
End If
End Sub
Public Function TextFormatBreak(lineBreaker As String) As String
Text = mDoc.textFormat(lineBreaker)
End Function

Public Property Get Text() As String
Text = mDoc.textDoc
End Property
Public Property Let Text(RHS As String)
Dim oldenabled As Boolean
LastListIndex = 0
If Me.HaveMarkedText Then
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.MarkNext = 0
End If
oldenabled = glistN.enabled
glistN.enabled = False
Set mDoc = New Document
mDoc.ColorEvent = Not NoColor
mDoc.textDoc = RHS
all = mDoc.DocLines
BlockStartPara = 0
BlockStartPos = 1
BlockStartLine = 0
BlockTopline = 0
BlockStartCursor = 0
BlockEndCursor = 0
BlockEndPara = 0
BlockEndPos = 1
BlockEndline = 0
BlockBottomline = 0
ClearRedo
ClearUndo
glistN.enabled = oldenabled
Dim what$, ThatPara As Long, pastchars
   pastchars = mDoc.TextParagraphLine(1, what$, ThatPara)
mDoc.MarkParagraphID = ThatPara
End Property
Public Property Let Text2(RHS As String)
Dim oldenabled As Boolean
If Me.HaveMarkedText Then
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.MarkNext = 0
End If
oldenabled = glistN.enabled
glistN.enabled = False
Set mDoc = New Document
mDoc.ColorEvent = Not NoColor
mDoc.textDoc2 = RHS
all = mDoc.DocLines
BlockStartPara = 0
BlockStartPos = 1
BlockStartLine = 0
BlockTopline = 0
BlockStartCursor = 0
BlockEndCursor = 0
BlockEndPara = 0
BlockEndPos = 1
BlockEndline = 0
BlockBottomline = 0
ClearRedo
ClearUndo
glistN.enabled = oldenabled
Dim what$, ThatPara As Long, pastchars
   pastchars = mDoc.TextParagraphLine(1, what$, ThatPara)
mDoc.MarkParagraphID = ThatPara
End Property
Public Property Let TextColor(RHS As String)
Dim oldenabled As Boolean
If Me.HaveMarkedText Then
BlockStartPara = BlockEndPara
BlockStartPos = BlockEndPos
glistN.MarkNext = 0
End If
oldenabled = glistN.enabled
glistN.enabled = False
Set mDoc = New Document
mDoc.textDocColor = RHS
all = mDoc.DocLines
ClearRedo
ClearUndo
glistN.enabled = oldenabled
Dim what$, ThatPara As Long, pastchars
   pastchars = mDoc.TextParagraphLine(1, what$, ThatPara)
mDoc.MarkParagraphID = ThatPara
End Property
Public Property Get TextParagraphs() As Long
TextParagraphs = mDoc.DocParagraphs
End Property
Public Property Get TextLines() As Long
TextLines = mDoc.DocLines
End Property
Public Property Let LineSpacingSilent(ByVal RHS As Long)
glistN.addpixels = RHS
''Render
End Property
Public Property Let LineSpacing(ByVal RHS As Long)
glistN.addpixels = RHS
Render
End Property

Public Property Get SelLength() As Long
If HaveMarkedText Then
If BlockStartLine = BlockEndline Then
SelLength = Abs(BlockStartCursor - BlockEndCursor)
Else
Dim a$(), oldmydoc As Document, i As Long, sum As Long
' maybe we didn't need to preserve old UndoDoc...I need to clarify this..
Set oldmydoc = UndoDoc
Set UndoDoc = New Document
FillBlock
sum = UndoDoc.SizeCRLF + 2
If UndoDoc.LfLead Then sum = sum - 1
If UndoDoc.CrTail Then sum = sum + 1
SelLength = sum - 2 ' no the last vbcrlf
End If
Set UndoDoc = oldmydoc
End If
End Property
Public Property Let SelLengthSilent(ByVal RHS As Long)
If RHS <= 0 Then
    ' a copy of glistN_MarkDestroyAny
    BlockEndPara = mDoc.MarkParagraphID
    BlockStartPara = BlockEndPara
    BlockStartPos = BlockEndPos
    glistN.MarkNext = 0
'If glistN.Enabled Then glistN.ShowPan
Else
glistN.MarkNext = 1
   BlockStartCursor = glistN.SelStart
    BlockStartPara = mDoc.MarkParagraphID
    BlockStartPos = ParaSelStart  ' selstart isn't from start of the doc...
 
    BlockStartLine = glistN.ListIndex
        BlockEndPara = BlockStartPara
        BlockEndPos = BlockStartPos
        BlockEndline = BlockStartLine
    BlockEndCursor = BlockStartCursor + 1
    
mDoc.FindPos BlockStartLine + 1, BlockStartCursor, RHS, BlockEndline, BlockEndCursor, BlockEndPara, BlockEndPos

BlockEndline = BlockEndline - 1 ' base 0


    BlockTopline = BlockStartLine
    BlockBottomline = BlockEndline
If RHS <> 0 Then
glistN.MarkNext = 2
'If glistN.Enabled Then glistN.ShowPan
End If
End If
End Property
Public Property Let SelLength(ByVal RHS As Long)
If RHS <= 0 Then
    ' a copy of glistN_MarkDestroyAny
    BlockEndPara = mDoc.MarkParagraphID
    BlockStartPara = BlockEndPara
    BlockStartPos = BlockEndPos
    glistN.MarkNext = 0
If glistN.enabled Then glistN.ShowPan
Else
glistN.MarkNext = 1
   BlockStartCursor = glistN.SelStart
    BlockStartPara = mDoc.MarkParagraphID
    BlockStartPos = ParaSelStart  ' selstart isn't from start of the doc...
 
    BlockStartLine = glistN.ListIndex
        BlockEndPara = BlockStartPara
        BlockEndPos = BlockStartPos
        BlockEndline = BlockStartLine
    BlockEndCursor = BlockStartCursor + 1
    
mDoc.FindPos BlockStartLine + 1, BlockStartCursor, RHS, BlockEndline, BlockEndCursor, BlockEndPara, BlockEndPos

BlockEndline = BlockEndline - 1 ' base 0

    BlockTopline = BlockStartLine
    BlockBottomline = BlockEndline
If RHS <> 0 Then
glistN.MarkNext = 2
If glistN.enabled Then glistN.ShowPan
End If
End If
End Property

Public Property Get Visible() As Boolean
Visible = glistN.Visible
End Property

Public Property Let Visible(ByVal RHS As Boolean)
''If Not rhs Then LastSelStart = SelStart
glistN.Visible = RHS

End Property
Public Property Get FontSize() As Single

  FontSize = glistN.FontSize
 
End Property

Public Property Let FontSize(New_FontSize As Single)
glistN.FontSize = New_FontSize
End Property
Public Property Get Font() As Font
 Set Font = glistN.Font
End Property

Public Property Set Font(New_Font As Font)
    Set glistN.Font = New_Font
End Property
Public Property Let backcolor(ByVal RHS As OLE_COLOR)

    
glistN.backcolor = RHS
  
    
End Property
Public Property Get backcolor() As OLE_COLOR
    backcolor = glistN.backcolor
End Property
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = glistN.ForeColor
End Property

Public Property Let ForeColor(ByVal RHS As OLE_COLOR)

    glistN.ForeColor = RHS

End Property

Public Property Get top() As Long
top = glistN.TopTwips
End Property

Public Property Let top(ByVal RHS As Long)
glistN.TopTwips = RHS
End Property
Public Property Get Height() As Long
Height = glistN.HeightTwips
End Property

Public Property Let Height(ByVal RHS As Long)
glistN.HeightTwips = RHS
End Property
Public Sub Move(ByVal mleft As Long, ByVal mtop As Long, mWidth As Long, mHeight As Long)
glistN.MoveTwips mleft, mtop, mWidth, mHeight
End Sub
Public Sub Dereference()
Set glistN = Nothing
Set mDoc = Nothing
Set mDoc1 = Nothing

End Sub

Public Property Get SelText() As String
If HaveMarkedText Then
FillBlock

If UndoDoc.LfLead Then
SelText = Mid$(UndoDoc.textDoc, 2)
ElseIf UndoDoc.CrTail Then
SelText = UndoDoc.textDoc + vbCr
Else  ''If UndoDoc.LfLead = UndoDoc.CrTail Then
SelText = UndoDoc.textDoc
End If
End If
End Property
Public Property Get SelHtmlText() As String
If HaveMarkedText Then
''FillBlock
Dim curcolors(0 To 15) As Variant, i As Long
For i = 0 To 15
curcolors(i) = Right$("000000" + Hex$(htmlM2000(i)), 6)

Next i

If UndoDoc.LfLead Then
SelHtmlText = Replace(Mid$(UndoDoc.textHtmlDoc(curcolors()), 2), vbCrLf, "<br>")
Else
SelHtmlText = Replace(UndoDoc.textHtmlDoc(curcolors()), vbCrLf, "<br>")
End If


End If
End Property
Public Property Let SelText(ByVal RHS As String)
Dim kk As Long, Pad$


kk = SelStart
SelStart = kk
glistN_MarkDelete False
glistN.MarkNext = 1
BlockStartPara = mDoc.MarkParagraphID
BlockStartPos = ParaSelStart
BlockStartCursor = glistN.SelStart '- 1
BlockEndCursor = BlockStartCursor + 1
BlockStartLine = glistN.ListIndex
BlockEndline = BlockStartLine
BlockEndPara = BlockStartPara
BlockEndPos = BlockStartPos
Dim k As Long, l As Long
If Right$(RHS, 1) = vbCr Then RHS = RHS + vbLf
If RHS <> "" Then

If NumberOnly Then

Pad$ = space$(Len(RHS) + Len(glistN.list(0)))
Mid$(Pad$, 1, Charpos - 1) = Left$(glistN.list(0), Charpos - 1)
Mid$(Pad$, Charpos, Len(RHS)) = RHS
If Len(glistN.list(0)) > 0 Then
Mid$(Pad$, Charpos + Len(RHS) - 1) = Mid$(glistN.list(0), Charpos)
End If
RHS = glistN.list(0)
glistN.list(0) = Pad$
If Not glistN.list(0) = Pad$ Then glistN.list(0) = RHS
Else
mDoc.InsertDoc mDoc.MarkParagraphID, Charpos, RHS
End If
Render
End If
BlockEndPara = mDoc.MarkParagraphID
BlockEndPos = ParaSelStart
BlockEndline = glistN.ListIndex
BlockEndCursor = glistN.SelStart
If BlockEndline < BlockStartLine Then
    BlockTopline = BlockEndline
    BlockBottomline = BlockStartLine
Else
    BlockTopline = BlockStartLine
    BlockBottomline = BlockEndline
End If


k = kk
glistN.MarkNext = 2
SelStart = k

End Property
Public Property Let InsertTextNoRender(ByVal RHS As String)
Dim b As Boolean
ClearRedo
PushUndo
glistN_MarkDelete False
glistN.MarkNext = 0
undomode = 5 ' like 1
RemoveUndo (RHS)
undomode = 4 ' like 0
b = enabled
enabled = False
Dim k As Long, l As Long
If Right$(RHS, 1) = vbCr Then RHS = RHS + vbLf
If RHS <> "" Then


mDoc.InsertDoc mDoc.MarkParagraphID, Charpos, RHS
all = mDoc.DocLines

End If
enabled = b
End Property
Public Property Let InsertText2(ByVal RHS As String)
glistN_MarkDelete True
glistN.MarkNext = 0
If Right$(RHS, 1) = vbCr Then RHS = RHS + vbLf
mDoc.ResetColor mDoc.MarkParagraphID
mDoc.InsertDoc mDoc.MarkParagraphID, Charpos, RHS
End Property

Public Property Let InsertText(ByVal RHS As String)
glistN_MarkDelete False
glistN.MarkNext = 0
Dim k As Long, l As Long
If Right$(RHS, 1) = vbCr Then RHS = RHS + vbLf
If RHS <> "" Then
mDoc.ResetColor mDoc.MarkParagraphID
mDoc.InsertDoc mDoc.MarkParagraphID, Charpos, RHS
Render

End If
End Property
Public Property Get CharPosStart() As Long
Dim findit As Long
Dim a As Long, b As Long, c As Long, d As Long
findit = -1
mDoc.FindPos glistN.ListIndex + 1, glistN.SelStart, findit, a, b, c, d
CharPosStart = findit
End Property
Public Property Get SelStart() As Long
'
Dim findit As Long
Dim a As Long, b As Long, c As Long, d As Long
findit = -1
If glistN.MarkNext > 0 Then
If BlockTopline = BlockBottomline Then
If BlockStartCursor < BlockEndCursor Then
mDoc.FindPos BlockTopline + 1, BlockStartCursor, findit, a, b, c, d
Else
mDoc.FindPos BlockTopline + 1, BlockEndCursor, findit, a, b, c, d
End If
ElseIf BlockTopline = BlockStartLine Then
mDoc.FindPos BlockTopline + 1, BlockStartCursor, findit, a, b, c, d
Else
mDoc.FindPos BlockTopline + 1, BlockEndCursor, findit, a, b, c, d
End If
Else
mDoc.FindPos glistN.ListIndex + 1, glistN.SelStart, findit, a, b, c, d
End If
If findit <> -1 Then LastSelStart = findit
SelStart = findit
End Property

Public Property Let SelStart(ByVal RHS As Long)
Dim findit As Long, sel1 As Long
sel1 = SelLength
    BlockEndPara = mDoc.MarkParagraphID
    BlockStartPara = BlockEndPara
    BlockStartPos = BlockEndPos
    glistN.MarkNext = 0
Dim a As Long, b As Long, c As Long, d As Long
findit = RHS
mDoc.FindPos 1, 0, findit, a, b, c, d

SetRowColumn a, b, sel1 <> 0
If sel1 <> 0 Then SelLength = sel1: LastSelStart = sel1
End Property
Public Property Let SelStartSilent(ByVal RHS As Long)
Dim findit As Long
Dim a As Long, b As Long, c As Long, d As Long
LastSelStart = RHS
findit = RHS
mDoc.FindPos 1, 0, findit, a, b, c, d
SetRowColumn a, b, True
End Property

Sub ClearUndo()
undopara = -1
NeoUndo = 0
UndoMax = 20
ReDim MultiUndo(1 To UndoMax)
End Sub
Function FindNewUndo() As Long
If NeoUndo = 0 Then
UndoMax = 20
ReDim MultiUndo(1 To UndoMax)
End If
NeoUndo = NeoUndo + 1
If NeoUndo > UndoMax Then
UndoMax = UndoMax * 2
ReDim Preserve MultiUndo(1 To UndoMax)
ElseIf UndoMax > 20 And NeoUndo < UndoMax / 2 Then
UndoMax = UndoMax / 2
ReDim Preserve MultiUndo(1 To UndoMax)
End If
FindNewUndo = NeoUndo
End Function
Private Sub PushUndo()
If undopara = -1 Then Exit Sub
Dim b As Long
b = FindNewUndo
With MultiUndo(b)
.undotext = undotext  ' new  seltext
.undopara = undopara  ' not need for this 2nd version
.undomode = undomode  ' needed
.undopos = undopos  ' new not charpos but selstart
undopos = ParaSelStart
undotext = vbNullString
undopara = mDoc.MarkParagraphID  ' nod needed maybe for trace errors later
undomode = -1
undopara = -1
End With
End Sub
Private Function PopUndo() As Boolean
' false ...no undo saved
If NeoUndo < 1 Then Exit Function

With MultiUndo(NeoUndo)
undotext = .undotext
.undotext = vbNullString  ' flush
undopara = .undopara
undomode = .undomode
undopos = .undopos
End With
NeoUndo = NeoUndo - 1
PopUndo = True
If undopara < 0 Then Exit Function

End Function
Sub ClearRedo()
Redopara = -1
NeoRedo = 0
RedoMax = 20
ReDim MultiRedo(1 To RedoMax)
End Sub
Function FindNewRedo() As Long
If NeoRedo = 0 Then
RedoMax = 20
ReDim MultiRedo(1 To RedoMax)
End If
NeoRedo = NeoRedo + 1
If NeoRedo > RedoMax Then
RedoMax = RedoMax * 2
ReDim Preserve MultiRedo(1 To RedoMax)
ElseIf RedoMax > 20 And NeoRedo < RedoMax / 2 Then
RedoMax = RedoMax / 2
ReDim Preserve MultiRedo(1 To RedoMax)
End If
FindNewRedo = NeoRedo
End Function
Private Sub PushRedo()

If Redopara = -1 Then Exit Sub
Dim b As Long
b = FindNewRedo
With MultiRedo(b)
.undotext = Redotext
.undopara = Redopara
.undomode = Redomode
.undopos = Redopos
Redopos = ParaSelStart
Redotext = vbNullString
Redopara = mDoc.MarkParagraphID
Redomode = -1
End With
End Sub
Private Function PopRedo() As Boolean
' false ...no Redo saved
If NeoRedo < 1 Then Exit Function
With MultiRedo(NeoRedo)
Redotext = .undotext
.undotext = vbNullString  ' flush
Redopara = .undopara
Redomode = .undomode
Redopos = .undopos
End With
NeoRedo = NeoRedo - 1
PopRedo = True
End Function

Public Sub ResetUndoRedo()
ClearUndo
ClearRedo
Redopara = -1
undopara = -1
End Sub
Public Function CurrentParagraph() As String
'If showparagraphWork And showparagraph Then
'CurrentParagraph = Replace(Replace(mdoc.TextParagraph(mdoc.MarkParagraphID), " ", ChrW$(&HB7)), ChrW$(&H2007), ChrW$(&H2423))
'Else
CurrentParagraph = mDoc.TextParagraph(mDoc.MarkParagraphID)
'End If
End Function
Public Function NextParagraph() As String
Dim dummy As Long
If showparagraphWork And showparagraph Then
NextParagraph = Replace(Replace(mDoc.NextTextParagraph(mDoc.MarkParagraphID, dummy), " ", ChrW$(&HB7)), ChrW$(&H2007), ChrW$(&H2423))
Else
NextParagraph = mDoc.NextTextParagraph(mDoc.MarkParagraphID, dummy)
End If
End Function
Public Function AutoIntNewLine() As Boolean
' if autonumber
Dim a$, newid As Long 'mDoc.MarkParagraphID
Dim b$, c$, d$, e$, skip As Boolean, tr$, maxw As Long
showparagraphWork = False
b$ = CurrentParagraph
On Error Resume Next
If AutoNumber And Not mDoc.IsPart(mDoc.MarkParagraphID) Then
If IsNumberLabel(b$, c$) Then
d$ = NextParagraph: If d$ <> "" Then If IsNumberLabel(d$, e$) Then skip = True
If Not skip Then c$ = Trim$(c$)
If e$ <> "" Then If val(e$) > val(c$) + 10 Then skip = False
If (val(c$) Mod 10 = 0) And Not skip Then
tr$ = String$(Len(NLTrim2(c$)), "0")
a$ = Format((val(c$) + 10), tr$) + Mid$(CurrentParagraph, 1, MyTrimL(CurrentParagraph) - 1) + Mid$(CurrentParagraph, Len(c$) + MyTrimL(CurrentParagraph), MyTrimL(Mid$(CurrentParagraph, Len(c$) + MyTrimL(CurrentParagraph))) - 1)
Else
If e$ <> "" And Trim$(Str((val(c$) + 1))) = Trim$(e$) Then
maxw = glistN.UserControlTextWidth(c$)
a$ = space$(Len(c$))
While glistN.UserControlTextWidth(a$) < maxw
a$ = a$ + " "
Wend
c$ = vbNullString

Else
tr$ = String$(Len(NLTrim2(c$)), "0")
a$ = Format((val(c$) + 1), tr$) + Mid$(CurrentParagraph, 1, MyTrimL(CurrentParagraph) - 1) + Mid$(CurrentParagraph, Len(c$) + MyTrimL(CurrentParagraph), MyTrimL(Mid$(CurrentParagraph, Len(c$) + MyTrimL(CurrentParagraph))) - 1)

End If
End If
Else
If glistN.UseTab And MyTrimL2NoTab(CurrentParagraph) = 1 Then
a$ = String$(Len(CurrentParagraph) - Len(NLtrim$(CurrentParagraph)), vbTab)
Else
a$ = space$(Len(CurrentParagraph) - Len(NLtrim$(CurrentParagraph)))
End If
If Len(a$) > Charpos Then a$ = vbNullString
End If
Else
If glistN.UseTab And MyTrimL2NoTab(CurrentParagraph) = 1 Then
a$ = String$(Len(CurrentParagraph) - Len(NLtrim$(CurrentParagraph)), vbTab)
Else
a$ = space$(Len(CurrentParagraph) - Len(NLtrim$(CurrentParagraph)))
End If
If Len(a$) > Charpos Then a$ = vbNullString

End If

If a$ <> "" Then
RemoveUndo vbCrLf + a$
InsertText = vbCrLf + a$
AutoIntNewLine = True

End If
showparagraphWork = True

End Function

Public Property Get Length() As Long
Length = mDoc.SizeCRLF
End Property


Function GetLCIDFromKeyboard() As Long
    Dim Buffer As String, ret&, r&
    Buffer = String$(514, 0)
      r = GetKeyboardLayout(DWL_ANYTHREAD) And &HFFFF
      r = val("&H" & Right(Hex(r), 4))
        ret = GetLocaleInfo(r, LOCALE_ILANGUAGE, StrPtr(Buffer), Len(Buffer))
    GetLCIDFromKeyboard = CLng(val("&h" + Left$(Buffer, ret - 1)))
End Function
Public Property Let ReplaceTitle(a$)

Dim curlines As Long
NoTitle = False
NoHeadwrap = False
curlines = mDoc1.DocLines
mDoc1.EmptyDoc
mDoc1.textDoc = Title$ + a$
mDoc1.WrapAgain
If mDoc1.DocLines >= curlines + glistN.lines + 1 Then
mDoc1.EmptyDoc
NoHeadwrap = True
mDoc1.textDoc = Title$ + a$
End If
glistN.HeadLine = vbNullString ' reset
glistN.HeadLine = " " ' give a fake headline
onelineheight = glistN.HeadlineHeight
glistN.HeadlineHeight = mDoc1.DocLines * onelineheight

there12345:
glistN.ShowMe False, True

End Property
Private Sub PrintItem1X(mHdc As Long, ByVal c As String, r As RECT, col() As Variant)
On Error Resume Next
    Dim cc As String, all As Long, CM As String, offset As Long, addthis As Long, hRgn As Long
    Dim b As RECT
    Dim ww$(), items&, i, rleft, cuts()
    If UBound(col()) <= 1 Or NoColor Then
        PrintItem1 mHdc, c, r
        Exit Sub
    End If
    rleft = r.Left
    all = rleft + glistN.UserControlTextWidthPixels(c)
    r.Left = all
    If InStr(c, vbTab) > 0 Then
    ReDim cuts(UBound(col()))
    ww$() = SplitStringUsingArray(c, col(), items&, cuts())
    If showparagraphonce Then
        cc = "∂"
        DrawTextEx mHdc, StrPtr(cc), -1, r, mway, VarPtr(tParam)
    End If
    cc = vbNullString
    
    
    For i = items& To 0 Step -1
        CM = ww$(i)
        r.Left = offset + rleft + addthis
        r.top = r.top + TextPixelOffset
        b = r
        If i > 0 Then
        
            b.Left = rleft + glistN.UserControlTextWidthPixels(Left$(c, cuts(i - 1))) + 1
          If b.Left < 0 Then b.Left = 0
          
        
        End If

        If Not NoColor Then SetTextColor mHdc, M2000(col(i * 2))
         hRgn = CreateRectRgnIndirect(b)
    SelectClipRgn mHdc, hRgn
        DrawTextEx mHdc, StrPtr(CM), -1, r, mway, VarPtr(tParam)
          SelectClipRgn mHdc, &H0
   DeleteObject hRgn
           r.top = r.top - TextPixelOffset
   
    Next i
Else
    ww$() = SplitStringUsingArrayNoTab(c, col(), items&)
    If showparagraphonce Then
        cc = "∂"
        DrawTextEx mHdc, StrPtr(cc), -1, r, mway, VarPtr(tParam)
    End If
    cc = vbNullString
    CM = c
    For i = items& To 0 Step -1
        c = ww$(i)
        CM = Left$(CM, Len(CM) - Len(ww$(i)))
        offset = glistN.UserControlTextWidthPixels(CM)
        ' for Kenring
        addthis = glistN.UserControlTextWidthPixels(CM + c) - offset - glistN.UserControlTextWidthPixels(c)
        r.Left = offset + rleft + addthis
        r.top = r.top + TextPixelOffset
        
        If Not NoColor Then SetTextColor mHdc, M2000(col(i * 2))
    
        DrawTextEx mHdc, StrPtr(c), -1, r, mway, VarPtr(tParam)
        r.top = r.top - TextPixelOffset
        If r.Left < 0 Then Exit For
    Next i

End If
    r.Left = rleft
    End Sub

    


Private Sub PrintItem1XOLD(mHdc As Long, ByVal c As String, r As RECT, col() As Variant)
On Error Resume Next
    Dim cc As String, all As Long, CM As String, offset As Long, addthis As Long, hRgn As Long
    Dim b As RECT
    Dim ww$(), items&, i, rleft, cuts()
    If UBound(col()) <= 1 Or NoColor Then
        PrintItem1 mHdc, c, r
        Exit Sub
    End If
    rleft = r.Left
    all = rleft + glistN.UserControlTextWidthPixels(c)
    r.Left = all
    If InStr(c, vbTab) > 0 Then
    ReDim cuts(UBound(col()))
    ww$() = SplitStringUsingArray(c, col(), items&, cuts())
    If showparagraphonce Then
        cc = "∂"
        DrawTextEx mHdc, StrPtr(cc), -1, r, mway, VarPtr(tParam)
    End If
    cc = vbNullString
    
    
    For i = items& To 0 Step -1
        CM = ww$(i)
        r.Left = offset + rleft + addthis
        r.top = r.top + TextPixelOffset
        b = r
        If i > 0 Then
        
            b.Left = rleft + glistN.UserControlTextWidthPixels(Left$(c, cuts(i - 1))) + 1
          If b.Left < 0 Then b.Left = 0
          
        
        End If

        If Not NoColor Then SetTextColor mHdc, M2000(col(i * 2))
         hRgn = CreateRectRgnIndirect(b)
    SelectClipRgn mHdc, hRgn
        DrawTextEx mHdc, StrPtr(CM), -1, r, mway, VarPtr(tParam)
          SelectClipRgn mHdc, &H0
   DeleteObject hRgn
           r.top = r.top - TextPixelOffset
   
    Next i
Else
    ww$() = SplitStringUsingArrayNoTab(c, col(), items&)
    If showparagraphonce Then
        cc = "∂"
        DrawTextEx mHdc, StrPtr(cc), -1, r, mway, VarPtr(tParam)
    End If
    cc = vbNullString
    CM = c
    For i = items& To 0 Step -1
        c = ww$(i)
        CM = Left$(CM, Len(CM) - Len(ww$(i)))
        offset = glistN.UserControlTextWidthPixels(CM)
        ' for Kenring
        addthis = glistN.UserControlTextWidthPixels(CM + c) - offset - glistN.UserControlTextWidthPixels(c)
        r.Left = offset + rleft + addthis
        r.top = r.top + TextPixelOffset
        
        If Not NoColor Then SetTextColor mHdc, M2000(col(i * 2))
    
        DrawTextEx mHdc, StrPtr(c), -1, r, mway, VarPtr(tParam)
        r.top = r.top - TextPixelOffset
        If r.Left < 0 Then Exit For
    Next i

End If
    r.Left = rleft
    End Sub



Function SplitStringUsingArray(ByVal c$, bb() As Variant, wc&, cuts() As Variant) As String()
Dim cc$()
ReDim cc$((UBound(bb()) + 1) / 2)
Dim st&, ste&, w&
wc& = 0
ste& = UBound(bb())
For w& = 1 To ste& - 2 Step 2
On Error Resume Next
If bb(w& + 2) - bb(w&) < 1 Then
cc$(wc&) = Mid$(c$, 1, bb(w&))
cuts(wc&) = bb(w& + 2)
Else
If wc& > 0 Then
cc$(wc&) = cc$(wc& - 1) + Mid$(c$, bb(w&), bb(w& + 2) - bb(w&))
Else
cc$(wc&) = Mid$(c$, bb(w&), bb(w& + 2) - bb(w&))
End If
cuts(wc&) = Len(cc$(wc&))

End If
If bb(w& + 2) - bb(w&) = 0 Then
wc& = wc& - 1
SplitStringUsingArray = cc$()
Exit Function
End If
wc& = wc& + 1
Next w&
If bb(w&) > 0 Then
If Trim$(Mid$(c$, bb(w&))) <> vbNullString Then

cc$(wc&) = c$
cuts(wc&) = Len(c$)
Else
 wc& = wc& - 1
End If
End If

SplitStringUsingArray = cc$()
End Function

Function SplitStringUsingArrayNoTab(c$, bb() As Variant, wc&) As String()
Dim cc$()
ReDim cc$((UBound(bb()) + 1) / 2)
Dim st&, ste&, w&
wc& = 0
ste& = UBound(bb())
For w& = 1 To ste& - 2 Step 2
On Error Resume Next
If bb(w& + 2) - bb(w&) < 1 Then
cc$(wc&) = Mid$(c$, bb(w&))
Else
cc$(wc&) = Mid$(c$, bb(w&), bb(w& + 2) - bb(w&))
End If
If cc$(wc&) = vbNullString Then
wc& = wc& - 1
SplitStringUsingArrayNoTab = cc$()
Exit Function
End If
wc& = wc& + 1
Next w&
If bb(w&) > 0 Then
cc$(wc&) = Mid$(c$, bb(w&))
End If
If cc$(wc&) = vbNullString Then wc& = wc& - 1
SplitStringUsingArrayNoTab = cc$()
End Function

Function M2000(ch As Variant) As Long
Dim b As Integer

Select Case ColorSet
Case 0
If ch = 13 Then
M2000 = rgb(&HEE, &HEE, &H33)
ElseIf ch = 12 Then
M2000 = rgb(&HDD, &HAA, 0)
Else
M2000 = QBColor(ch Mod 16)
End If
Case 1
Select Case ch
Case 7
M2000 = rgb(&H68, &H68, &H13)
Case 10
M2000 = 0   '' black  '' basic
Case 11
M2000 = rgb(32, 32, 255)
Case 12
M2000 = rgb(&H80, &H40, &H10)
Case 13
M2000 = rgb(&HAA, &H90, &H20)
Case 14
M2000 = rgb(&HAA, &H60, &H30)
Case 15
M2000 = rgb(&H63, &H30, &H63) '' string

End Select
End Select

End Function
Function htmlM2000(ch As Variant) As Long
Dim b As Integer

Select Case Abs(pagiohtml$ = "DARK")
Case 0
If ch = 13 Then
htmlM2000 = rgb(&H33, &HEE, &HEE)
ElseIf ch = 12 Then
htmlM2000 = rgb(0, &HAA, &HDD)

Else
htmlM2000 = (QBColor(ch Mod 16) And &HFF00&) + (QBColor(ch Mod 16) \ &H10000) + ((QBColor(ch Mod 16) And &HFF) * &H10000)
End If
Case 1
Select Case ch
Case 7
htmlM2000 = rgb(&H13, &H68, &H68)
Case 10
htmlM2000 = 0   '' black  '' basic
Case 11
htmlM2000 = rgb(255, 32, 32)
Case 12
htmlM2000 = rgb(&H10, &H40, &H80)
Case 13
htmlM2000 = rgb(&H20, &H37, &H97)
Case 14
htmlM2000 = rgb(&H30, &H60, &HAA)
Case 15
htmlM2000 = rgb(&H63, &H30, &H63) '' string
Case Else
htmlM2000 = rgb(255, 255, 255)
End Select
End Select

End Function
Private Function aheadstatus(a$, Optional srink As Boolean = True, Optional pos As Long = 1) As String 'ok
Dim b$, part$, w$, pos2 As Long, word$

If a$ = vbNullString Then Exit Function
If pos = 0 Then pos = 1
Do While pos <= Len(a$)

    w$ = Mid$(a$, pos, 1)
    If part$ = vbNullString And w$ = "0" Then
        If pos + 2 <= Len(a$) Then
            If LCase(Mid$(a$, pos, 2)) Like "0[x˜]" Then
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

    If w$ = Chr(34) Then
    
        If part$ <> "" Then
        b$ = b$ & part$
        End If
         'If Right(b$, 1) <> "l" Then Exit Do
        part$ = "S"
        pos = pos + 1
        Do While pos <= Len(a$)
        If Mid$(a$, pos, 1) = Chr(34) And Mid$(a$, pos - 1, 1) <> "`" Then Exit Do
    If Asc(Mid$(a$, pos, 1)) < 32 Then Exit Do
   
        pos = pos + 1
        Loop
    ElseIf w$ = "(" Then
        If part$ <> "" Then
            ' after
            ElseIf Right$(b$, 1) = "a" Then
            b$ = Left$(b$, Len(b$) - 1)
            part$ = vbNullString
            Else
            part$ = "N"
              
        End If
      pos = pos + 1
        If Not BlockParam2(a$, pos) Then Exit Do
        
       If Mid$(a$, pos + 1, 1) <> "." Then b$ = b$ & part$
        part$ = vbNullString
        
    ElseIf w$ = "{" Then
If word$ <> "" Then
  Select Case myUcase(word$, True)
  Case "THEN", "ELSE", "‘œ‘≈"
  pos = pos - 4
  Exit Do
  Case "ELSE.IF"
  pos = pos - 7
  Exit Do
  Case "¡ÀÀ…Ÿ”"
  pos = pos - 6
  Exit Do
  Case "¡ÀÀ…Ÿ”.¡Õ"
  pos = pos - 9
  Exit Do
  End Select
  End If
    

    If part$ <> "" Then b$ = b$ & part$
        If LenB(b$) = 0 Then
        ElseIf Right(b$, 1) <> "l" Then
        Exit Do
        End If
 
        blockStringAhead a$, pos
        pos = pos + 1
        part$ = "S"
        

            If pos <= Len(a$) Then
        If Not blockStringAhead(a$, pos) Then Exit Do
        pos = pos - 1
        End If
''Exit Do

    Else
        Select Case w$
        Case ","  ' bye bye
        Exit Do
        Case "%"
            If part$ = vbNullString Then
            ' lathos
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
                    part$ = vbNullString
            End If

        Case "-", "+", "|"
                    b$ = b$ & part$
                    If b$ = vbNullString Then
                        part$ = "o"
                    Else
                    
                part$ = "o"
                End If
        Case "*", "/", "^"
            If part$ <> "o" Then
            b$ = b$ & part$
            End If
            part$ = "o"
        Case " ", ChrW$(160)
          If word$ <> "" Then
  Select Case myUcase(word$, True)
  Case "THEN", "ELSE", "‘œ‘≈"
  pos = pos - 4
  Exit Do
  Case "ELSE.IF"
  pos = pos - 7
  Exit Do
  Case "¡ÀÀ…Ÿ”"
  pos = pos - 6
  Exit Do
  Case "¡ÀÀ…Ÿ”.¡Õ"
  pos = pos - 9
  Exit Do
  End Select
  End If
jumphere:
        word$ = vbNullString
            If part$ <> "" Then
            b$ = b$ & part$
            part$ = vbNullString
            Else
            'skip
            End If
        Case "0" To "9", "."
            If part$ = "N" Then
            If Len(a$) < pos Then
            If Mid$(a$, pos + 1, 1) Like "[&@#%~]" Then
            pos = pos + 1
            End If
            End If
            ElseIf part$ = "S" Then
            
            Else
            b$ = b$ & part$
            part$ = "N"
            End If
        Case "&"
        If part$ = vbNullString Then
        part$ = "S"
        Else
        b$ = part$
        part$ = "S"
        End If
        Case "e", "E", "Â", "≈"
        If LenB(word$) = 0 Then part$ = "N"
        
        word$ = word$ + w$
           If part$ = "S" Then
            
            
            ElseIf part$ = "N" Then
            
            b$ = b$ & part$
            'part$ = "N"
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
        word$ = word$ + w$
        ElseIf part$ = "S" Then
       
        Else
        
     b$ = b$ & part$
     part$ = "N"
     word$ = word$ + w$
            End If
        End Select
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
Sub ReColor()
WrapAll
ManualInform
End Sub
Sub ReColorBlock()
Dim order1, order2
order1 = 0: order2 = 0
If BlockEndPara <> BlockStartPara Then
order1 = mDoc.ParagraphOrder(BlockStartPara)
order2 = mDoc.ParagraphOrder(BlockEndPara)
End If
If order1 <= order2 Then
mDoc.WrapAgainBlock BlockStartPara, BlockEndPara
'mDoc.WrapAgainColorBlock BlockStartPara, BlockEndPara
Else
'mDoc.WrapAgainBlock BlockEndPara, BlockStartPara
mDoc.WrapAgainColorBlock BlockEndPara, BlockStartPara
End If


ManualInform
End Sub
Sub ColorBlockOnly()
Dim order1, order2
order1 = 0: order2 = 0
If BlockEndPara <> BlockStartPara Then
order1 = mDoc.ParagraphOrder(BlockStartPara)
order2 = mDoc.ParagraphOrder(BlockEndPara)
End If
If order1 <= order2 Then
mDoc.WrapAgainColorBlock2 BlockStartPara, BlockEndPara
Else
mDoc.WrapAgainColorBlock2 BlockEndPara, BlockStartPara
End If
ManualInform
End Sub
Private Function ISSTRINGA(a$, r$) As Boolean
If Len(a$) < 2 Then ISSTRINGA = False: Exit Function
Dim q$, w As Long
r$ = vbNullString
q$ = Chr(34)
w = 1
Do
w = w + 1
w = InStr(w, a$, q$)
If w > 0 Then
If Mid$(a$, w - 1, 1) <> "`" Then Exit Do
ElseIf w = 0 Then
Exit Do
End If
Loop
If w = 0 Then ISSTRINGA = False: Exit Function
r$ = Mid$(a$, 2, w - 2)
a$ = NLtrim$(Mid$(a$, w + 1))
ISSTRINGA = True
End Function
Private Sub countbrakcet(a$, k As Long, counter As Long)
Dim i As Long, j As Integer, blind As Boolean
For i = Abs(k) To Len(a$)
j = AscW(Mid$(a$, i, 1))
Select Case j
Case 34
blind = Not blind
Case 123
If Not blind Then counter = counter + 1
Case 125
If Not blind Then counter = counter - 1: If counter = 0 Then k = i: Exit For
End Select
Next i

End Sub
Function blockStringAhead(s$, pos1 As Long) As Long
Dim i As Long, j As Long, c As Long
c = Len(s$)
i = pos1
If i > c Then blockStringAhead = False: Exit Function
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
j = j + 1
End Select
i = i + 1

Loop Until i > c Or j = 0
pos1 = i - 1
blockStringAhead = True


End Function

Function FastSymbol(a$, c$, Optional mis As Boolean = False, Optional cl As Long = 1) As Boolean
Dim i As Long, j As Long
'If Len(c$) <> cl Then Stop  ; only for check
j = Len(a$)
If j = 0 Then Exit Function
i = MyTrimL(a$)
If i > j Then Exit Function  ' this is not good
If j - i < cl - 1 Then
If mis Then MyEr "missing " & c$, "ÎÂﬂÂÈ " & c$
Exit Function
End If
'If c$ = Mid$(a$, i, cl) Then
If InStr(c$, Mid$(a$, i, cl)) > 0 Then
'a$ = Mid$(a$, MyTrimLi(a$, i + cl))
a$ = Mid$(a$, cl + 1)
'Mid$(a$, i, cl) = space$(cl)
FastSymbol = True
ElseIf mis Then
MyEr "missing " & c$, "ÎÂﬂÂÈ " & c$
End If
End Function

