VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mmname As String
Dim dnum As Long, ndnum As Long
Dim limit(10) As Long  ' max 9 dimensions
Dim dtable(10) As Long
Dim nlimit(10) As Long  ' max 9 dimensions
Dim limBase(10) As Long
Dim nLimbase(10) As Long

Private actualdata() As Variant
Dim maxitems As Long
Dim maxitems1 As Long
Dim nmaxitems As Long
Dim nmaxitems1 As Long
Public myarrbase As Long
Public common As Boolean
Public IHaveClass As Boolean
Public IHaveGui As Boolean
' get a floating Group, maybe with groups inside....as floating group
' used also for other mHandler with t1 = 1 (Inventories)
Public GroupRef As Object
' only for check
'Private mGroupRef As Object
Public CodeName As String
Private oldmaxitems As Long
Public Arr As Boolean
Public RevOrder As Boolean
Private Type SAFEArrayPart
   cDims                        As Integer
   fFeatures                    As Integer
   cbElements                   As Long
   cLocks                       As Long
   pvData                       As Long
   End Type
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal addr As Long, retval As Long)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal addr As Long, ByVal NewVal As Long)
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal addr As Long, retval As Integer)
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal addr As Long, retval As Byte)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
  Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" _
      (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
        Private Declare Sub CopyPart Lib "KERNEL32" Alias "RtlMoveMemory" _
      (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SafeArrayGetDim Lib "OleAut32.dll" (ByVal pSA As Long) As Long
Private Declare Function SafeArrayLock Lib "OleAut32.dll" (ByVal pSA As Long) As Long
Private Declare Function SafeArrayUnlock Lib "OleAut32.dll" (ByVal pSA As Long) As Long
  
Private Declare Function SafeArrayAccessData Lib "OleAut32.dll" (ByVal pSA As Long, ppVdata As Long) As Long
Private Declare Function SafeArrayUnaccessData Lib "OleAut32.dll" (ByVal pSA As Long) As Long
Private Declare Function SafeArrayGetElemsize Lib "OleAut32.dll" (ByVal pSA As Long) As Long
Private Declare Sub SafeArrayCopyData Lib "OleAut32.dll" (ByVal psaSource As Any, ByVal psaTarget As Any)
Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    cElements As Long
    lLBound As Long
End Type
Public index As Long, final As Boolean, refgroup As Object, grouperror As Boolean
Public Sub Compute0(p As Variant, ss$)
If final Then CantAssignValue: Exit Sub
Dim v As Long
On Error Resume Next
v = 0
Do While v <= Me.UpperMonoLimit
If MyIsNumeric(actualdata(v)) Then
                    Select Case ss$
                                Case "=", "g"
                                actualdata(v) = p
                                Case "+"
                                actualdata(v) = p + actualdata(v)
                                If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
                                Case "*"
                                 actualdata(v) = p * actualdata(v)
                                 If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
                                Case "-"
                                actualdata(v) = actualdata(v) - p
                                If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
                                Case "/"
                                If p = 0 Then Exit Do
                                 actualdata(v) = actualdata(v) / p
                                 If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
                                  Case "~", "!!"
                                 actualdata(v) = -1 - (actualdata(v) <> 0)
                                 Case "++"
                                actualdata(v) = 1 + actualdata(v)
                                Case "--"
                                actualdata(v) = actualdata(v) - 1
                                  Case "-!"
                                actualdata(v) = -actualdata(v)
                                Case "+="
                                actualdata(v) = p + actualdata(v)
                                If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
                                Case "*="
                                 actualdata(v) = p * actualdata(v)
                                 If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
                                Case "-="
                                actualdata(v) = actualdata(v) - p
                                If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
                                Case "/="
                                If p = 0 Then Exit Do
                                 actualdata(v) = actualdata(v) / p
                                 If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
                                  
                                 End Select
                             
                            
End If
                           v = v + 1
                           Err.clear
Loop

End Sub
Public Sub Compute(p As Double, ss$)
If final Then CantAssignValue: Exit Sub
Dim v As Long
On Error Resume Next
v = 0
Do While v <= Me.UpperMonoLimit
If MyIsNumeric(actualdata(v)) Then
                    Select Case ss$
                                Case "=", "g"
                                actualdata(v) = p
                                Case "+"
                                actualdata(v) = p + actualdata(v)
                                 If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
                                Case "*"
                                 actualdata(v) = p * actualdata(v)
                                  If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
                                Case "-"
                                actualdata(v) = actualdata(v) - p
                                 If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
                                Case "/"
                                If p = 0 Then Exit Do
                                 actualdata(v) = actualdata(v) / p
                                  If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
                                  Case "~", "!!"
                                 actualdata(v) = -1 - (actualdata(v) <> 0)
                                 Case "++"
                                actualdata(v) = 1 + actualdata(v)
                                Case "--"
                                actualdata(v) = actualdata(v) - 1
                                  Case "-!"
                                actualdata(v) = -actualdata(v)
                                  
                                 End Select
                          
                            
End If
                           v = v + 1
                           Err.clear
Loop

End Sub
Public Sub Compute2(p As Variant, ss$)
If final Then CantAssignValue: Exit Sub
Dim v As Long
On Error Resume Next
v = 0
Do While v <= Me.UpperMonoLimit
If MyIsNumeric(actualdata(v)) Then
                    Select Case ss$
                                Case "="
                                actualdata(v) = p
                                Case "+="
                                actualdata(v) = actualdata(v) + p
                                If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
                                Case "*="
                                 actualdata(v) = p * actualdata(v)
                                 If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
                                Case "-="
                                actualdata(v) = actualdata(v) - p
                                If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
                                Case "/="
                                If MyRound(p, 13) = 0 Then Exit Do
                                 actualdata(v) = actualdata(v) / p
                               If RoundDouble Then If VarType(actualdata(v)) = vbDouble Then actualdata(v) = MyRound(actualdata(v), 13)
                                 End Select
                               
                            
End If
                           v = v + 1
                           Err.clear
Loop

End Sub
Public Sub Compute3(ss$)
Dim v As Long
If final Then CantAssignValue: Exit Sub
On Error Resume Next
v = 0
Do While v <= Me.UpperMonoLimit
If MyIsNumeric(actualdata(v)) Then
                    Select Case ss$
                                  Case "~"
                                 actualdata(v) = -1 - (actualdata(v) <> 0)
                                 Case "++"
                                actualdata(v) = 1 + actualdata(v)
                                Case "--"
                                actualdata(v) = actualdata(v) - 1
                                  Case "-!"
                                actualdata(v) = -actualdata(v)
                                  
                                 End Select
End If
                           v = v + 1
                           Err.clear
Loop

End Sub
Property Get Value() As Variant
If index >= 0 And index <= maxitems1 Then

If IsObject(actualdata(index)) Then
    If Typename(actualdata(index)) = "PropReference" Then
        Value = actualdata(index)
    Else
        Set Value = actualdata(index)
    End If
Else
Value = actualdata(index)
End If
End If
End Property
Public Sub PushProp(a As Variant)
Set GroupRef = a
Arr = False
End Sub
Public Sub CopySerialize(mmm() As Variant)
' may we get an exception here
Dim m_base As Long, first12 As Long
Dim m_base_0 As Long, first12_0 As Long, r1 As Long, r2 As Long

GetMem4 VarPtrArray(mmm()), m_base
GetMem4 VarPtrArray(actualdata()), m_base_0
r1 = SafeArrayLock(m_base)
r2 = SafeArrayLock(m_base_0)
r1 = SafeArrayAccessData(ByVal m_base, first12)
r2 = SafeArrayAccessData(ByVal m_base_0, first12_0)

'GetMem4 m_base + 12, first1
'GetMem4 m_base_0 + 12, first1_0
PutMem4 m_base + 12, first12_0
PutMem4 m_base_0 + 12, first12
r2 = SafeArrayUnaccessData(ByVal m_base_0)
r1 = SafeArrayUnaccessData(ByVal m_base)
r2 = SafeArrayUnlock(m_base_0)
r1 = SafeArrayUnlock(m_base)
End Sub
Public Sub CopyBackSerialize(mmm() As Variant)
' may we get an exception here
Dim m_base As Long, first12 As Long
Dim m_base_0 As Long, first12_0 As Long, r1 As Long, r2 As Long
If Not RevOrder Then Err.Raise vbObjectError + 12345, "CopyBack", "Failed. is not vb6"
Dim backup()
backup() = actualdata()
GetMem4 VarPtrArray(mmm()), m_base
GetMem4 VarPtrArray(backup()), m_base_0
r1 = SafeArrayLock(m_base)
r2 = SafeArrayLock(m_base_0)
r1 = SafeArrayAccessData(ByVal m_base, first12)
r2 = SafeArrayAccessData(ByVal m_base_0, first12_0)

'GetMem4 m_base + 12, first1
'GetMem4 m_base_0 + 12, first1_0
PutMem4 m_base + 12, first12_0
PutMem4 m_base_0 + 12, first12
r2 = SafeArrayUnaccessData(ByVal m_base_0)
r1 = SafeArrayUnaccessData(ByVal m_base)
r2 = SafeArrayUnlock(m_base_0)
r1 = SafeArrayUnlock(m_base)
End Sub
Public Sub Car(m As Object)
Dim k As New mArray
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
If count > 0 Then
    k.PushDim 1
    k.PushEnd
    k.SerialItem actualdata(0), 0, 1
    If IsObject(Me.item(0)) Then
    Select Case Typename(Me.item(0))
    Case "Group"
    Set obj = Me.item(0)
    CopyGroupObj obj, obj2
    Set k.item(0) = obj2
    Set obj2 = Nothing
    Case "lambda"
    Dim actualdata1()
    With Me.item(0)
    .CopyTo spare, actualdata1()
    End With
    Set k.item(0) = spare
    Set spare = Nothing
    Case Else
    If TypeOf Me.item(0) Is mHandler Then
        If Me.item(0).t1 = 3 Then
        'Set m = New mHandler
        'm.t1 = 3
        Set m = Me.item(0)
       ' If Not CheckIsmArray(m, var()) Then
        '    Set m = New mHandler
         '   m.t1 = 3
          '  Set m.objref = Me.item(0)
        'End If
        Set k = Nothing
        Exit Sub
        End If
     ElseIf TypeOf Me.item(0) Is mArray Then
             Set m = New mHandler
        m.t1 = 3
        Set m.objref = Me.item(0)
        Set k = Nothing
        Exit Sub
    End If
    Set k.item(0) = Me.item(0)
    
    End Select
    Else
    k.item(0) = Me.item(0)
    End If
End If
k.Arr = True
Set m = k
Set k = Nothing
End Sub
Public Sub Cdr(m As Object)
Dim k As New mArray
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
If dnum > 0 Then
k.StartResize: k.PushDim Me.count - 1: k.PushEnd
Else
k.Arr = True
Set k.GroupRef = Me.GroupRef
k.IHaveGui = Me.IHaveGui
k.IHaveClass = Me.IHaveClass
Set m = k
Exit Sub
End If
k.myarrbase = myarrbase
k.RevOrder = False
For i = 1 To Me.UpperMonoLimit
If IsObject(Me.item(i)) Then
Select Case Typename(Me.item(i))
Case "Group"
Set obj = Me.item(i)
CopyGroupObj obj, obj2
Set k.item(i - 1) = obj2
Set obj2 = Nothing
Case "lambda"
Dim actualdata1()
With Me.item(i)
.CopyTo spare, actualdata1()
End With
Set k.item(i - 1) = spare
Set spare = Nothing
Case Else
Set k.item(i - 1) = Me.item(i)
End Select
Else
k.item(i - 1) = Me.item(i)
End If
Next i
'' new
Set k.GroupRef = Me.GroupRef
k.IHaveGui = Me.IHaveGui
k.IHaveClass = Me.IHaveClass
Set m = k

End Sub
Public Sub processAppend(m As Object)
Dim pppp As New mArray, pppp1 As mArray
Set pppp1 = m
If dnum > 0 Then
    Me.CopyArray pppp
    If Not pppp1.IsEmpty Then pppp1.AppendArray pppp
ElseIf Not pppp1.IsEmpty Then
    Set pppp = pppp1
End If
pppp.Arr = True
Set m = pppp
Set pppp = Nothing
Set pppp1 = Nothing
End Sub

Public Sub AppendArray(k As mArray)
'append this to k
If k.count = 0 Then
Me.CopyArray k
Exit Sub
End If
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
Dim fromhere As Long
fromhere = k.UpperMonoLimit + 1
k.SerialItem v, Me.count + k.count, 10
k.myarrbase = myarrbase
k.RevOrder = RevOrder
If Right$(k.arrname, 2) = "%(" Then
For i = 0 To Me.UpperMonoLimit
If IsObject(Me.item(i)) Then
Select Case Typename(Me.item(i))
Case "Group"
Set obj = Me.item(i)
CopyGroupObj obj, obj2
Set k.item(i + fromhere) = obj2
Set obj2 = Nothing
Case "lambda"
Dim actualdata()
With Me.item(i)
.CopyTo spare, actualdata()
End With
Set k.item(i + fromhere) = spare
Set spare = Nothing
Case Else
Set k.item(i + fromhere) = Me.item(i)
End Select
Else
k.item(i + fromhere) = Int(Me.item(i))
End If
Next i
Else
''''''''''****************************
For i = 0 To Me.UpperMonoLimit

If IsObject(Me.item(i)) Then
Select Case Typename(Me.item(i))
Case "Group"
Set obj = Me.item(i)
CopyGroupObj obj, obj2
Set k.item(i + fromhere) = obj2
Set obj2 = Nothing
Case Else
Set k.item(i + fromhere) = Me.item(i)
End Select
Else
k.item(i + fromhere) = Me.item(i)
End If
Next i
End If
'' new
Set k.GroupRef = Me.GroupRef
k.IHaveGui = Me.IHaveGui
k.IHaveClass = Me.IHaveClass
End Sub



Public Sub CopyArray(k As mArray)
'copy this to k

If k Is Me Then Exit Sub
If k Is Nothing Then Set k = New mArray
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
k.SerialItem v, Me.UpperMonoLimit, 8
k.myarrbase = myarrbase
k.RevOrder = False
k.StartResize
For i = 0 To dnum - 1
k.PushDim limit(i), limBase(i)
Next i

k.PushEnd
k.RevOrder = RevOrder
Set k.GroupRef = Me.GroupRef
If Right$(k.arrname, 2) = "%(" Then
For i = 0 To Me.UpperMonoLimit

If IsObject(Me.item(i)) Then
Select Case Typename(Me.item(i))
Case "Group"
Set obj = Me.item(i)
CopyGroupObj obj, obj2
Set k.item(i) = obj2
Set obj2 = Nothing
Case "lambda"
Dim actualdata()
With Me.item(i)
.CopyTo spare, actualdata()
End With
Set k.item(i) = spare
Set spare = Nothing

Case Else
Set k.item(i) = Me.item(i)
End Select
Else
k.item(i) = Int(Me.item(i))
End If
Next i
Else
''''''''''****************************
For i = 0 To Me.UpperMonoLimit

If IsObject(Me.item(i)) Then
Select Case Typename(Me.item(i))
Case "Group"
Set obj = Me.item(i)
CopyGroupObj obj, obj2
Set k.item(i) = obj2
Set obj2 = Nothing
Case "lambda"
Dim actualdata1()
With Me.item(i)
.CopyTo spare, actualdata1()
End With
Set k.item(i) = spare
Set spare = Nothing
Case Else
Set k.item(i) = Me.item(i)
End Select
Else
k.item(i) = Me.item(i)
End If
Next i
End If
'' new

k.Arr = Me.Arr
k.IHaveGui = Me.IHaveGui
k.IHaveClass = Me.IHaveClass
k.final = Me.final
End Sub
Public Sub CopyArrayRev(k As mArray)
'copy this to k
If k Is Me Then Exit Sub
Dim v As Variant, i As Long, obj As Object, obj2 As Object, spare As lambda
Dim mm As Long, M1 As Long
k.SerialItem v, Me.UpperMonoLimit, 8
k.myarrbase = myarrbase
k.RevOrder = False
k.StartResize
For i = 0 To dnum - 1
k.PushDim limit(i), limBase(i)
Next i

k.PushEnd
k.RevOrder = RevOrder
Set k.GroupRef = Me.GroupRef
mm = Me.UpperMonoLimit
If Right$(k.arrname, 2) = "%(" Then
For i = 0 To Me.UpperMonoLimit
M1 = mm - i
If IsObject(Me.item(M1)) Then
Select Case Typename(Me.item(M1))
Case "Group"
Set obj = Me.item(M1)
CopyGroupObj obj, obj2
Set k.item(i) = obj2
Set obj2 = Nothing
Case "lambda"
Dim actualdata()
With Me.item(M1)
.CopyTo spare, actualdata()
End With
Set k.item(i) = spare
Set spare = Nothing

Case Else
Set k.item(i) = Me.item(M1)
End Select
Else
k.item(i) = Int(Me.item(M1))
End If
Next i
Else
''''''''''****************************
For i = 0 To Me.UpperMonoLimit
M1 = mm - i
If IsObject(Me.item(M1)) Then
Select Case Typename(Me.item(M1))
Case "Group"
Set obj = Me.item(i)
CopyGroupObj obj, obj2
Set k.item(i) = obj2
Set obj2 = Nothing
Case "lambda"
Dim actualdata1()
With Me.item(M1)
.CopyTo spare, actualdata1()
End With
Set k.item(i) = spare
Set spare = Nothing
Case Else
Set k.item(i) = Me.item(M1)
End Select
Else
k.item(i) = Me.item(M1)
End If
Next i
End If
'' new

k.Arr = Me.Arr
k.IHaveGui = Me.IHaveGui
k.IHaveClass = Me.IHaveClass
k.final = Me.final
End Sub

Public Function UpperMonoLimit()
UpperMonoLimit = maxitems1
End Function
Public Property Get count()
If dnum = 0 Then
count = 0
Else
count = maxitems1 + 1 ' sa
End If
End Property
Public Property Let arrname(aName As String)
mmname = aName
End Property
Public Property Get arrname() As String
arrname = mmname
End Property
Public Sub PushDim(nDim As Long, Optional limBase As Variant)
If ndnum < 10 Then
nlimit(ndnum) = nDim
If IsMissing(limBase) Then
nLimbase(ndnum) = myarrbase
Else
nLimbase(ndnum) = CLng(limBase)
End If
nmaxitems = nmaxitems * nDim
ndnum = ndnum + 1
End If
nmaxitems1 = nmaxitems - 1
End Sub
Public Sub PushEnd()
On Error GoTo there11
If ndnum = 0 Then
dnum = 0
maxitems = 1
maxitems1 = 0
ReDim actualdata(maxitems1) As Variant
Else
Dim i As Long
For i = 0 To 9
limit(i) = nlimit(i): limBase(i) = nLimbase(i)
Next i

    dnum = ndnum
    maxitems = nmaxitems
    maxitems1 = nmaxitems1

If nmaxitems1 < 0 Then
    ReDim actualdata(1) As Variant
Else
    ReDim Preserve actualdata(maxitems1) As Variant
End If
Dim mx As Long
mx = maxitems
i = 0
If limit(i) > 0 Then
If RevOrder Then
Dim revlimit(), revlimbase()
ReDim revlimit(UBound(limit())), revlimbase(UBound(limit()))
    For i = 0 To dnum - 1
    revlimit(i) = limit(dnum - i - 1)
    revlimbase(i) = limBase(dnum - i - 1)
    Next i
    For i = 0 To dnum - 1
    limit(i) = revlimit(i)
    limBase(i) = revlimbase(i)
    Next i
    
End If
For i = 0 To dnum - 1
    dtable(i) = mx / limit(i)
    mx = dtable(i)
Next i
Else
GoTo there11
End If

End If
Exit Sub
there11:
dnum = 0
maxitems = 1
maxitems1 = 0
ReDim actualdata(maxitems)
End Sub
Public Sub StartResize()
oldmaxitems = maxitems1
ndnum = 0 ' no actions
nmaxitems = 1
nmaxitems1 = 0
End Sub
Friend Sub GetDnum(cursor As Long, item, lim)
If dnum = 0 Then item = 0: lim = 0: Exit Sub
If RevOrder Then
item = limit(dnum - cursor - 1)
lim = limBase(dnum - cursor - 1)
Else
item = limit(cursor)
lim = limBase(cursor)
End If
End Sub
Friend Function bDnum() As Variant
bDnum = dnum
End Function

Public Function SerialItem(item As Variant, cursor As Long, command As Long) As Boolean
If command = 1 Then
If dnum = 0 Then Exit Function
If cursor < 0 Or cursor >= maxitems Then Exit Function
If VarType(item) = vbObject Then
Set actualdata(cursor) = item
Else
actualdata(cursor) = item
End If
ElseIf command = 2 Then
If dnum = 0 Then Exit Function
If cursor < 0 Or cursor >= maxitems Then Exit Function
item = actualdata(cursor)
ElseIf command = 3 Then
    Fill item
ElseIf command = 4 Then
    cursor = maxitems
ElseIf command = 5 Then
cursor = dnum
item = dnum
ElseIf command = 6 Then
If cursor > 9 Or cursor < 0 Then
item = 0
Else
If RevOrder Then
item = limit(dnum - cursor - 1)
Else
item = limit(cursor)
End If
End If
ElseIf command = 7 Then  'erase all data
Dim aa As Variant
Fill aa
ElseIf command = 8 Then
If dnum = 1 Then
If cursor = -1 Then
dnum = 0 ' no actions
maxitems = 1
maxitems1 = 0
ndnum = 0 ' no actions
nmaxitems = 1
nmaxitems1 = 0
ReDim actualdata(0)
CodeName = GetAnyName$
Arr = True
limit(0) = 0
Else
    ReDim Preserve actualdata(cursor) As Variant
    
    maxitems = cursor
End If
    limit(0) = maxitems

End If
ElseIf command = 9 Then
If dnum = 1 Then
    ReDim Preserve actualdata(cursor) As Variant
    maxitems = cursor - 1
    limit(0) = maxitems
End If
maxitems1 = maxitems - 1
ElseIf command = 10 Then
   
    ReDim Preserve actualdata(cursor) As Variant
    maxitems = cursor
    limit(0) = maxitems
dnum = 1
maxitems1 = maxitems - 1
ElseIf command = 11 Then   ' new from 9.4  rev 4
If cursor > 9 Or cursor < 0 Then
item = 0
Else
If RevOrder Then
item = limBase(dnum - cursor - 1)
Else
item = limBase(cursor)
End If
End If
ElseIf command = 12 Then   ' new from 9.4  rev 4
    For command = 0 To 9
        limBase(command) = cursor
    Next command
    myarrbase = cursor
End If
SerialItem = True

End Function

Public Function PushOffset(curitem As Long, curdim As Long, nDim As Long) As Boolean
If curdim >= dnum Then
' error...
PushOffset = False
Else


    If RevOrder Then
         Dim pr As Long
        pr = dnum - curdim - 1
    If nDim >= limit(pr) - limBase(pr) Then Exit Function
    curitem = curitem + dtable(pr) * (nDim + limBase(pr))
    Else
    If nDim >= limit(curdim) - limBase(curdim) Then Exit Function
    curitem = curitem + dtable(curdim) * (nDim + limBase(curdim))
    End If

    
PushOffset = True
End If
End Function
Public Property Get itemPtr(curitem As Long) As Long
itemPtr = VarPtr(actualdata(curitem))
End Property
Friend Sub SwapItem(curitem As Long, thisvariant)
SwapVariant actualdata(curitem), thisvariant
End Sub
Friend Function IsStringItem(curitem As Long) As Variant
IsStringItem = Typename(actualdata(curitem)) = "String"
End Function
Public Property Get item(curitem As Long) As Variant
On Error GoTo there1
If Arr Then
If maxitems1 >= curitem Then
If IsObject(actualdata(curitem)) Then
If Typename(actualdata(curitem)) = "PropReference" Then
item = actualdata(curitem).Value
Else
Set item = actualdata(curitem)
End If
Else

item = actualdata(curitem)
End If
Exit Property
End If
Else
If curitem = -2 Then
Set item = GroupRef
ElseIf curitem < 0 Then
' used for test only - no mHandler here
'If TypeOf GroupRef Is mHandler Then
'If GroupRef.t1 = 3 Then
'If MyIsObject(GroupRef.objref.Value) Then
'Set item = GroupRef.objref.Value
'Else
'item = GroupRef.objref.Value
'End If
'Else
If GroupRef.t1 <> 1 Then Err.Raise 1010, , "Need an Inventory Object"
If GroupRef.objref.Done Then
    If GroupRef.objref.IsObj Then
        Set item = GroupRef.objref.ValueObj
    Else
        item = GroupRef.objref.Value
    End If
Else
    Set item = GroupRef
End If

End If
'End If
End If
there1:

End Property
Public Property Get itemnumeric(curitem As Long) As Variant
On Error GoTo there1
If Arr Then
If maxitems1 >= curitem Then
If IsObject(actualdata(curitem)) Then
If Typename(actualdata(curitem)) = "PropReference" Then
itemnumeric = actualdata(curitem).Value
ElseIf Typename(actualdata(curitem)) = "mHandler" Then
If actualdata(curitem).t1 = 4 Then
If actualdata(curitem).sign = -1 Then
itemnumeric = CDbl(-actualdata(curitem).index_cursor)
Else
itemnumeric = CDbl(actualdata(curitem).index_cursor)
End If

Else
itemnumeric = 0
End If
Else
itemnumeric = 0
End If
Else
If VarType(actualdata(curitem)) = 8 Then
itemnumeric = val(actualdata(curitem))
Else
itemnumeric = actualdata(curitem)
End If
End If
Exit Property
Else
''Stop
End If
Else
If curitem = -2 Then
ElseIf curitem < 0 Then
If GroupRef.objref.Done Then
If GroupRef.objref.IsObj Then
itemnumeric = 0
Else
itemnumeric = GroupRef.objref.Value
End If
End If
End If
End If
there1:

End Property

Public Property Let item(curitem As Long, item As Variant)

On Error Resume Next
If curitem < 0 Then

ElseIf final Then
If Not MyIsObject(actualdata(curitem)) Then
If VarType(actualdata(curitem)) = vbEmpty Then
actualdata(curitem) = item
Else
CantAssignValue
End If
Else
CantAssignValue
End If
ElseIf Not IsObject(actualdata(curitem)) Then
actualdata(curitem) = item
Else
If Not MyIsObject(item) Then
    If TypeOf actualdata(curitem) Is PropReference Then
        actualdata(curitem).Value = item
    ElseIf TypeOf actualdata(curitem) Is VarItem Then
        actualdata(curitem) = item
    ElseIf TypeOf actualdata(curitem) Is mHandler Then
        If actualdata(curitem).t1 = 4 Then
            actualdata(curitem) = item
        End If
    Else
        Set actualdata(curitem) = Nothing
        actualdata(curitem) = CLng(0)
    End If
Else
actualdata(curitem).Value = item
End If
End If

End Property
Public Property Let ItemStr(curitem As Long, item As String)

On Error Resume Next
If curitem < 0 Then

ElseIf final Then
If Not MyIsObject(actualdata(curitem)) Then
If VarType(actualdata(curitem)) = vbEmpty Then
MoveStringToVariant item, actualdata(curitem)
Else
CantAssignValue
End If
Else
CantAssignValue
End If
ElseIf Not IsObject(actualdata(curitem)) Then
MoveStringToVariant item, actualdata(curitem)
Else
If TypeOf actualdata(curitem) Is PropReference Then
    actualdata(curitem).ValueStr = item
Else
    Set actualdata(curitem) = Nothing
    actualdata(curitem) = CLng(0)
End If
End If

End Property
Public Property Set item(curitem As Long, item As Variant)
If final Then CantAssignValue: Exit Property
On Error Resume Next
If curitem < 0 Then
If Not GroupRef Is Nothing Then
    If TypeOf GroupRef Is mHandler Then
        If GroupRef.t1 = 1 Then
            GroupRef.objref.index = Abs(curitem + 100)
            GroupRef.objref.Done = True
            Set GroupRef.objref.ValueObj = item
            Exit Property
           
        End If
    ElseIf curitem = -2 Then
    If Typename$(item) = Typename$(GroupRef) Then
    Set GroupRef = item
    
    Exit Property
    End If
    End If
End If
MyEr "Wrong setting", "Λάθος όρισμα"
ElseIf curitem <= maxitems1 Then
Set actualdata(curitem) = item
Else

If grouperror Then
MyEr "Something is wrong with this group", "Κάτι πάει στραβά με αυτή την ομάδα"
Else
MyEr "Something is wrong with this array", "Κάτι πάει στραβά με αυτόν τον πίνακα"
End If
End If
End Property

Private Sub Fill(item As Variant)
Dim aa As Object
If dnum > 0 Then
Dim i As Long
If IHaveClass Then
For i = 0 To maxitems - 1
CopyGroupObj GroupRef, aa, True
Set actualdata(i) = aa
Next i
Else
If Typename(item) = "Empty" Then
ReDim actualdata(maxitems)
''maxitems = 0

Else
For i = 0 To maxitems - 1
actualdata(i) = item
Next i
End If
End If
End If
End Sub
Public Sub Fillobj(bstack As basetask)
Dim obj As Variant, i As Long
Set obj = bstack.lastobj
    Set bstack.lastobj = Nothing
For i = 0 To maxitems - 1
Set actualdata(i) = obj
Next i
Set obj = Nothing
End Sub
Public Sub FillLambda(bstack As basetask)
Dim obj As Variant, i As Long
Set obj = bstack.lastobj
    Set bstack.lastobj = Nothing
For i = 0 To maxitems - 1
 CopyLambda obj, bstack
Set actualdata(i) = bstack.lastobj
Set bstack.lastobj = Nothing
Next i
Set obj = Nothing
End Sub
Sub MakeitObject2(var As Variant)
Dim aa As Object
Set aa = New Group
Set var = aa
End Sub
Sub CopyGroupObj(myobject1 As Object, myobject2 As Object, Optional LinkToMe As Boolean = False)
 Dim v As Long, s$, frm$, vvl As Variant, x1 As Long, ss$
Dim subgroup As Object, sub2 As Object, myArray As mArray, mySecondArray As mArray

Dim mygroup2 As New Group
 Dim myobject As Group
 Set myobject = myobject1
 If myobject.IamApointer Then
 Set myobject2 = myobject
 Exit Sub
 End If
 
             With myobject
                If .IamFloatGroup Then
         
                 .PeekItem 1, vvl
                 mygroup2.BeginFloat vvl + 2
                 For x1 = 0 To vvl * 2 + 2
                    .PeekItem x1, vvl
                    If Typename(vvl) = "Group" Then
                    Set subgroup = vvl
                    CopyGroupObj subgroup, sub2
                    Set subgroup = Nothing
                    Set vvl = sub2
                    Set sub2 = Nothing
                    ElseIf Typename(vvl) = "mArray" Then
                    If Not vvl.common Then
                        Set myArray = vvl
                        Set mySecondArray = New mArray
                        myArray.CopyArray mySecondArray
                        Set myArray = Nothing
                        Set vvl = mySecondArray
                        Set mySecondArray = Nothing
                    End If
                    End If
                    mygroup2.PokeItem x1, vvl
     
                Next x1
                
                mygroup2.HasStrValue = myobject.HasStrValue
                mygroup2.HasValue = myobject.HasValue
                mygroup2.HasSet = myobject.HasSet
                mygroup2.HasParameters = myobject.HasParameters
                mygroup2.HasParametersSet = myobject.HasParametersSet
                Set mygroup2.SuperClassList = myobject.SuperClassList
                'If Not GroupRef Is Nothing Then
                If LinkToMe Then
                    Set mygroup2.LinkRef = myobject
                Else
                    .PeekItem x1, vvl
                    mygroup2.PokeItem x1, vvl
                End If
                Set mygroup2.Events = .Events
                mygroup2.highpriorityoper = .highpriorityoper
                mygroup2.HasUnary = .HasUnary
                End If
            End With
         Set myobject2 = mygroup2
End Sub
Property Get IsEmpty() As Variant
IsEmpty = dnum = 0
End Property
Function IsObj() As Boolean
If index >= 0 And index <= maxitems1 Then
If IsObject(actualdata(index)) Then
    IsObj = True
End If
End If
End Function
Function IsEnum(p As Variant) As Boolean
Dim obj As mHandler
If index >= 0 And index <= maxitems1 Then
If IsObject(actualdata(index)) Then
If Typename(actualdata(index)) = "mHandler" Then
    Set obj = actualdata(index)
    IsEnum = obj.t1 = 4
    If IsEnum Then p = obj.index_cursor * obj.sign
End If
End If
End If
End Function
Private Sub Class_Initialize()
dnum = 0 ' no actions
maxitems = 1
maxitems1 = 0
ndnum = 0 ' no actions
nmaxitems = 1
nmaxitems1 = 0
ReDim actualdata(0)
CodeName = GetAnyName$
Arr = True
myarrbase = -ArrBase
End Sub

Private Sub Class_Terminate()
Set GroupRef = Nothing
Erase actualdata()
End Sub
Private Function GetAnyName$()
Dim ss$, j As Long
For j = 1 To 4
ss$ = ss$ & Chr(65 + Int((23 * Rnd) + 1))
Next j
GetAnyName$ = ChrW(&H1FFF) + ss$
End Function
Function MyIsNumeric(v As Variant) As Boolean
Dim n As Byte
GetMem1 VarPtr(v), n
If n < 2 Then Exit Function
MyIsNumeric = n < 8 Or n = 14
End Function
Function MyIsObject(v As Variant) As Boolean
Dim n As Byte
GetMem1 VarPtr(v), n
MyIsObject = n = 9
End Function
Function GetCopy() As Variant()
GetCopy = actualdata
End Function
'' Only for test. Not mHandler with t1=3 can be in Groupref
'Public Property Set GroupRef(RHS)
'If Not RHS Is Nothing Then
'If TypeOf RHS Is mHandler Then
'If RHS.t1 = 3 Then Stop
'End If
'End If
'Set mGroupRef = RHS
'End Property
'Public Property Get GroupRef() As Object
'Set GroupRef = mGroupRef
'End Property

