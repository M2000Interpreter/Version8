VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MemBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mPtr As Long
Private mItems As Long
Private ItemLen As Long
Private mBytes As Long
Private ErrorMes As Long
Public structref As FastCollection
Public UseStruct As Boolean
Private mExecBloc As Boolean
Private Strings As LongHash
Private BasicItem As Integer

Private Type SAFEARRAYBOUND
    cElements As Long
    lLBound As Long
End Type
Private Type SAFEARRAY1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    onebound(0 To 0) As SAFEARRAYBOUND
End Type
Private Width As Long, Height As Long
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal addr As Long, retval As Byte)
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Sub Costruct(ByteLen As Long, Optional Times As Long = 1, Optional clearme As Long = 0, Optional executable As Boolean = False, Optional thatbasicitem As Integer)

If mItems <> 0 Then ErrorMes = 1: Exit Sub
If Times < 1 Then ErrorMes = 2: Exit Sub
If ByteLen < 0 Then
BasicItem = vbString
ElseIf Not IsMissing(thatbasicitem) Then
BasicItem = thatbasicitem
End If
ByteLen = Abs(ByteLen)

If ByteLen < 1 Then ErrorMes = ErrorMes + 4
If ErrorMes <> 0 Then Exit Sub
ItemLen = ByteLen
mItems = Times
mExecBloc = Not executable
mPtr = Blockmalloc(ByteLen * Times, clearme, mExecBloc)
'If mExecBloc Then
If mPtr <> 0 Then mBytes = ByteLen * Times
End Sub
Property Get ItemSize()
ItemSize = ItemLen
End Property
Public Sub ClearError()
    ErrorMes = 0
End Sub
Public Property Get Status()
    Status = ErrorMes
End Property
Public Sub ResizeItems(NewSize As Long, Optional clearme As Long = 0)

If mBytes = 0 Then Exit Sub
Dim oPtr As Long


If NewSize > 0 Then
    If mExecBloc Then
        oPtr = BlockResizemalloc(mPtr, ItemLen * NewSize, clearme)
        If oPtr <> 0 Then
            mItems = NewSize
            mPtr = oPtr
        Else
            ErrorMes = 3
        End If
    Else
    
    If ItemLen * NewSize <= mBytes Then
            mPtr = BlockResizeVirtual(mPtr, ItemLen * NewSize)
            mItems = NewSize
            mBytes = ItemLen * NewSize
    Else
       oPtr = Blockmalloc(ItemLen * NewSize, clearme, mExecBloc)
       If oPtr <> 0 Then
            VirtualBlockCopy oPtr, mPtr, ItemLen * NewSize, mBytes
           
            BlockFreeVirtual mPtr, mBytes
            mItems = NewSize
            mBytes = ItemLen * NewSize
            mPtr = oPtr
        Else
            ErrorMes = 3
        End If
    End If
    End If
Else
    ErrorMes = 2
End If
End Sub
Public Property Get ItemsAreStrings() As Boolean
ItemsAreStrings = BasicItem = vbString
End Property
Public Property Get WhatIsBasicItem() As Integer
WhatIsBasicItem = BasicItem
End Property

Public Property Get GetBytePtr(offset As Variant) As Long
If ErrorMes = 0 And mItems > 0 Then
    If offset < 0 Or offset >= mItems * ItemLen Then
        ErrorMes = 4
    Else
    GetBytePtr = signlong(uintnew(mPtr) + offset)
    End If
Else
    GetBytePtr = 0  ' null
End If
End Property
Public Property Get GetPtr(offset As Variant) As Long
If ErrorMes = 0 And mItems > 0 Then
If offset < 0 Or offset >= mItems Then
    ' zero
    ErrorMes = 4
Else
    GetPtr = cUlng(uintnew(mPtr) + ItemLen * offset)
    End If
Else
    GetPtr = 0  ' null
End If
End Property
Public Property Get SizeByte() As Long
If ErrorMes = 0 And mItems > 0 Then
If mExecBloc Then
    SizeByte = BlockSize(mPtr)
    Else
    SizeByte = mBytes
    End If
End If
End Property
Public Property Get items() As Long
If ErrorMes = 0 And mItems > 0 And ItemLen > 0 Then
If mExecBloc Then
    items = BlockSize(mPtr) \ ItemLen
    Else
    items = mBytes \ ItemLen
    End If
    
End If
End Property
Public Function ValidArea(ByVal ThisStart, ByVal ThisWidth) As Boolean
Dim a As Currency
' ThisStart is a sign long which hold same bits with Unsigned address
If ErrorMes = 0 And mItems > 0 Then
a = uintnew(CCur(mPtr))   ' make mptr an Unsigned
If ThisStart = 0 Then
    ThisStart = a
Else
    ThisStart = uintnew(ThisStart)
End If
ThisWidth = CCur(ThisWidth)
' we have to make both Unsigned and then check them
If ThisStart - a >= 0 Then
If mExecBloc Then
If ThisStart + ThisWidth <= a + BlockSize(mPtr) Then ValidArea = True
Else
If ThisStart + ThisWidth <= a + mBytes Then ValidArea = True
End If
End If
End If

End Function
Public Function ValidArea2(ByVal ThisOffsetU As Double, ByVal ThisWidth As Double) As Boolean
' ThisOffsetU is Unsigned
If ErrorMes = 0 And mItems > 0 Then
If ThisOffsetU >= 0 Then
If mExecBloc Then
If ThisOffsetU * ItemLen + ThisWidth <= uintnew2(CDbl(BlockSize(mPtr))) Then ValidArea2 = True
Else
If ThisOffsetU * ItemLen + ThisWidth <= mBytes Then ValidArea2 = True
End If
End If
End If

End Function

Private Sub Class_Initialize()
Width = -1&
Height = -1&
End Sub

Private Sub Class_Terminate()
Set structref = Nothing
If mPtr <> 0 Then
If NoRun Then
If BlockSize(mPtr) > 0 Then
    BlockFree mPtr
End If
Else
        BlockFreeVirtual mPtr, mBytes
        mBytes = 0
End If
End If
mPtr = 0
End Sub

Public Property Get NoRun() As Boolean
NoRun = mExecBloc
End Property

Public Function PutStringAtOffset(addr As Variant, AnyString As String) As Variant
Dim offset
offset = uintnew(addr) - uintnew(mPtr)
If ErrorMes = 0 And mItems > 0 Then
    If offset < 0 Or offset + 3 >= mItems * ItemLen Then
        ErrorMes = 4
    Else
        If Strings Is Nothing Then Set Strings = New LongHash
        With Strings
        If .ExistKey(cUlng(offset + 1)) Then
           .ValueStr = AnyString
           PutStringAtOffset = .StrPointer
        Else
        .ItemCreator2 cUlng(offset + 1), AnyString
        PutStringAtOffset = .StrPointer
        End If
        .Done = False
        End With
    End If
Else
    PutStringAtOffset = 0  ' null
End If
End Function
Public Function GetStringFromOffset(addr As Variant, AnyString As String) As Variant
Dim offset
offset = uintnew(addr) - uintnew(mPtr)
If ErrorMes = 0 And mItems > 0 Then
    If offset < 0 Or offset + 3 >= mItems * ItemLen Then
        ErrorMes = 4
        AnyString = "Crash"
    Else
        If Strings Is Nothing Then Set Strings = New LongHash
        With Strings
        If .ExistKey(cUlng(offset + 1)) Then
           AnyString = .Value
           GetStringFromOffset = .StrPointer
        Else
            AnyString = "Crash"
            GetStringFromOffset = addr
        End If
        .Done = False
        End With
    End If
    
Else
    AnyString = "Crash"
    GetStringFromOffset = 0  ' null
End If

End Function
Public Function getData(F As Long, addr, lengthInBytes As Long) As Boolean
Dim offset, place
offset = uintnew(addr) - uintnew(mPtr)
If ErrorMes = 0 And mItems > 0 Then
    If offset < 0 Or offset + lengthInBytes - 1 >= mItems * ItemLen Then
        ErrorMes = 4
        getData = False
    Else
    On Error GoTo there
    Dim a1 As SAFEARRAY1D, a() As Byte
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLBound = 0
        .onebound(0).cElements = lengthInBytes
        .pvData = addr
    End With
        CopyMemory ByVal VarPtrArray(a()), VarPtr(a1), 4

        place = 1
        Get #F, , a()
        CopyMemory ByVal VarPtrArray(a), 0&, 4
        getData = True
  End If
    
Else
there:
        If place = 1 Then
                CopyMemory ByVal VarPtrArray(a), 0&, 4
        End If
        getData = False
End If
        
End Function
Public Function ExportToByte() As Byte()
Dim place
If ErrorMes = 0 And mItems > 0 Then

    On Error GoTo there
    Dim a1 As SAFEARRAY1D, a() As Byte
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLBound = 0
        .onebound(0).cElements = SizeByte
        .pvData = GetPtr(0)
    End With
        CopyMemory ByVal VarPtrArray(a()), VarPtr(a1), 4

        place = 1
        On Error Resume Next
        ExportToByte = a()
        If Err.Number > 0 Then Exit Function
        CopyMemory ByVal VarPtrArray(a), 0&, 4
       
  End If
    
Exit Function
there:
        If place = 1 Then
                CopyMemory ByVal VarPtrArray(a), 0&, 4
        End If
        

End Function



Public Function putData(F As Long, addr, lengthInBytes As Long) As Boolean
Dim offset, place
offset = uintnew(addr) - uintnew(mPtr)
If ErrorMes = 0 And mItems > 0 Then
    If offset < 0 Or offset + lengthInBytes - 1 >= mItems * ItemLen Then
        ErrorMes = 4
        putData = False
    Else
    On Error GoTo there
    Dim a1 As SAFEARRAY1D, a() As Byte
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLBound = 0
        .onebound(0).cElements = lengthInBytes
        .pvData = addr
    End With
        CopyMemory ByVal VarPtrArray(a()), VarPtr(a1), 4

        place = 1
        On Error Resume Next
        Put #F, , a()
        If Err.Number > 0 Then Close F: Exit Function
        CopyMemory ByVal VarPtrArray(a), 0&, 4
        putData = True
  End If
    
Else
there:
        If place = 1 Then
                CopyMemory ByVal VarPtrArray(a), 0&, 4
        End If
        putData = False
End If
        
End Function

Public Function GetStdPicture(Optional Width As Long = -1, Optional Height As Long = -1, Optional ByVal backcolor As Long = vbWhite, Optional RetainRatio As Boolean = False) As IPicture
   Dim a1 As SAFEARRAY1D, a() As Byte, Token As Long
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLBound = 0
        .onebound(0).cElements = Me.SizeByte
        .pvData = Me.GetPtr(0)
    End With
    CopyMemory ByVal VarPtrArray(a()), VarPtr(a1), 4
    
    Set GetStdPicture = LoadImageFromBuffer2(a(), Width, Height, backcolor, RetainRatio)
    
    CopyMemory ByVal VarPtrArray(a), 0&, 4
 
End Function

Public Function DrawSpriteToHdc(bstack As basetask, sprt As Boolean, angle!, zoomfactor!, blend!, Optional backcolor As Long = -1) As Boolean
   Dim a1 As SAFEARRAY1D, a() As Byte, Token As Long
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLBound = 0
        .onebound(0).cElements = Me.SizeByte
        .pvData = Me.GetPtr(0)
    End With
    CopyMemory ByVal VarPtrArray(a()), VarPtr(a1), 4
    DrawSpriteToHdc = DrawSpriteFromBuffer(bstack, a(), Not sprt, angle!, zoomfactor!, blend!, backcolor)
    
    CopyMemory ByVal VarPtrArray(a), 0&, 4
 
End Function
Public Function DrawImageToHdc(where As Object, Optional X As Long = 0&, Optional Y As Long = 0&, Optional Width As Long = -1, Optional Height As Long = -1) As Boolean
   Dim a1 As SAFEARRAY1D, a() As Byte, Token As Long
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLBound = 0
        .onebound(0).cElements = Me.SizeByte
        .pvData = Me.GetPtr(0)
    End With
    CopyMemory ByVal VarPtrArray(a()), VarPtr(a1), 4
    DrawImageToHdc = DrawImageFromBuffer(a(), where.hDC, X, Y, Width, Height)
    
    CopyMemory ByVal VarPtrArray(a), 0&, 4
 
End Function
Public Function ReadImageSizeX(SizeX As Variant) As Boolean
If Width <> -1 Then SizeX = Width: ReadImageSizeX = True: Exit Function
   Dim a1 As SAFEARRAY1D, a() As Byte, Token As Long
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLBound = 0
        .onebound(0).cElements = Me.SizeByte
        .pvData = Me.GetPtr(0)
    End With
    CopyMemory ByVal VarPtrArray(a()), VarPtr(a1), 4
    If ReadSizeImageFromBuffer(a(), Width, Height) Then
    SizeX = Width
    ReadImageSizeX = True
        Else
        Width = -1&
        Height = -1&
    End If
    
    CopyMemory ByVal VarPtrArray(a), 0&, 4


End Function
Public Function ReadImageSizeY(SizeY As Variant) As Boolean
If Height <> -1 Then SizeY = Height: ReadImageSizeY = True: Exit Function
   Dim a1 As SAFEARRAY1D, a() As Byte, Token As Long
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLBound = 0
        .onebound(0).cElements = Me.SizeByte
        .pvData = Me.GetPtr(0)
    End With
    CopyMemory ByVal VarPtrArray(a()), VarPtr(a1), 4
    If ReadSizeImageFromBuffer(a(), Width, Height) Then
    SizeY = Height
    ReadImageSizeY = True
    Else
        Width = -1&
        Height = -1&
    End If
    
    CopyMemory ByVal VarPtrArray(a), 0&, 4
End Function
Public Sub ResetImageSize()
Dim a1 As SAFEARRAY1D, a() As Byte, Token As Long
     With a1
        .cbElements = 1
        .cDims = 1
        .onebound(0).lLBound = 0
        .onebound(0).cElements = Me.SizeByte
        .pvData = Me.GetPtr(0)
    End With
    CopyMemory ByVal VarPtrArray(a()), VarPtr(a1), 4
    If Not ReadSizeImageFromBuffer(a(), Width, Height) Then
     Width = -1&
        Height = -1&
    End If
    
    CopyMemory ByVal VarPtrArray(a), 0&, 4
End Sub
Function Copy(r, p) As Object

Dim aCopy As New MemBlock, BLen As Long, mItems As Long
If mItems > 0 Then
    If mExecBloc Then
        BLen = BlockSize(mPtr)
    Else
        BLen = mBytes
    End If
    mItems = BLen \ ItemLen
    If r <> 0 Then
        If r >= 0 And r < mItems Then
            If p >= 1 Then
                aCopy.Costruct ItemLen, CLng(p), , r = 0 And mExecBloc
                If mItems > p Then
                    p = mItems
                End If
                CopyMemory ByVal aCopy.GetPtr(0), ByVal Me.GetPtr(r), ItemLen * p
            Else
                aCopy.Costruct ItemLen, 1&, , mExecBloc
                CopyMemory ByVal aCopy.GetPtr(0), ByVal Me.GetPtr(r), ItemLen
            End If
        Else
            aCopy.Costruct ItemLen, 1&, , mExecBloc
            CopyMemory ByVal aCopy.GetPtr(0), ByVal Me.GetPtr(r), ItemLen
        End If
    Else
    aCopy.Costruct ItemLen, mItems, , mExecBloc
    CopyMemory ByVal aCopy.GetPtr(0), ByVal mPtr, BLen
    End If
    Set Copy = aCopy
End If
End Function
Public Function Search(from As Long, what) As Long
Dim limit As Long, b As Byte, i As Long, w As Byte
b = Abs(what) And &HFF
limit = SizeByte - 1

For i = from To limit
    GetMem1 GetBytePtr(i), w
    If b = w Then
    Search = i + 1: Exit Function
    End If
Next
Search = i + 1
End Function
