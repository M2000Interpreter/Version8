VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "mHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Generic Handler
Option Explicit
Public objref As Object
Public t1 As Variant, indirect As Long
Public ReadOnly As Boolean
Public index_start As Variant ' 1 for start
Public index_End As Variant  ' -1 for end
Public index_cursor As Variant
Public sign As Variant
Public UseIterator As Boolean
Public oEnum As Variant
Public IamEnum As Boolean
Private Sub Class_Initialize()
indirect = -1
index_start = 0
index_End = -1
UseIterator = False
sign = 1
End Sub

''Public Par1 As Variant
Private Sub Class_Terminate()
If Not objref Is Nothing Then
'If Typename(objref) = "GuiM2000" Then Stop
Set objref = Nothing
End If
End Sub
Public Sub CopyTo(a As mHandler)
If IamEnum Then Set a = Me: Exit Sub
Set a = New mHandler
a.t1 = t1
''A.Par1 = Par1   ' key or index ???
' no copy for reference
If t1 = 1 Then a.ReadOnly = ReadOnly
a.indirect = indirect
a.index_start = index_start
a.index_End = index_End
a.index_cursor = index_cursor
a.UseIterator = UseIterator
a.sign = sign
Set a.objref = objref
End Sub
Public Function Iterate()
        Iterate = GetNext(ObjPtr(oEnum), index_cursor)
        If IsObject(index_cursor) Then Set objref = index_cursor: index_cursor = vbEmpty
        index_End = 1
        If Not Iterate Then index_End = -1
End Function
Sub ConstructEnumerator(penum As Variant)
IamEnum = True
Set oEnum = penum
If GetNext(ObjPtr(penum), index_cursor) Then
If IsObject(index_cursor) Then Set objref = index_cursor: index_cursor = vbEmpty
    index_End = 0
End If
End Sub
Friend Function CreateFromPicture(b As Long, w1 As Long, w2 As Long, w3 As Long) As Object
If t1 = 2 Then
Dim mm As MemBlock
Set mm = objref
Set CreateFromPicture = mm.CreateFromPicture(b, w1, w2, w3)
End If
End Function
