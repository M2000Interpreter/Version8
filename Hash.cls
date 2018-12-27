VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Hash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

 Private Type item
 Key As String
  ndx As Long
  KeyType As Integer  ' 0 as is , 1, reference/don't delete
  originaltype As Byte
  globalvar As Byte
 firsthash As Long
 lastpos As Long
 Pleft As Long  ' a list
 End Type
'
Private PriveSpace() As item
Dim MaxSpace As Long
Dim hashlen As Long
Dim toplim As Long
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal addr As Long, retval As Integer)
Private index As Long
Private Hash() As Long
Private Sub Class_Initialize()
    MaxSpace = 10
    ReDim PriveSpace(MaxSpace) As item, Hash(MaxSpace * 2 + 3)
    hashlen = MaxSpace * 2 + 3
    toplim = -1

End Sub
Friend Sub ReadVar(where, name$, num As Long)
    name$ = PriveSpace(where).Key
    num = PriveSpace(where).ndx
End Sub
Public Property Get lastNDX() As String
lastNDX = PriveSpace(index).ndx
End Property
Public Property Get lastkey() As String
lastkey = PriveSpace(index).Key
End Property

Private Sub ExpandHash()

    Dim i As Long
    hashlen = hashlen * 2 + 3
    
    ReDim Hash(hashlen) As Long
    For i = 0 To MaxSpace \ 2 - 1

    place HashFunc2(i), i
    Next i
End Sub
Friend Sub ReduceHash(newTop As Long, v())
    Dim i As Long
    
    For i = toplim To newTop Step -1
       With PriveSpace(i)
        
    
        If .Pleft > 0 Then
            Hash(.lastpos) = .Pleft
            .Key = vbNullString
            .Pleft = 0
        Else
            Hash(.lastpos) = 0
            .Key = vbNullString
        End If
        If .KeyType = 0 And .ndx >= 0 Then
            If IsObject(v(.ndx)) Then
             Set v(.ndx) = Nothing
            End If
            v(.ndx) = Empty
            End If
        .KeyType = 0
        .globalvar = CByte(0)
        .originaltype = CByte(0)
        .ndx = -1
            
        End With

    Next i
    If MaxSpace - newTop + 1 > 2 * newTop + 2 Then
        MaxSpace = 2 * (newTop + 1) + 1
        ReDim Preserve PriveSpace(MaxSpace) As item
        
        ReDim Hash(MaxSpace * 2 + 3)
       hashlen = MaxSpace * 2 + 3
     '  reHash...
     toplim = newTop - 1
    For i = 0 To toplim
        place HashFunc2(i), i
    Next i

Else
 toplim = newTop - 1
    End If
   
End Sub

Private Function Malloc() As Long

    If toplim + 1 >= MaxSpace Then
           '' expand hash
           MaxSpace = MaxSpace * 2
           ReDim Preserve PriveSpace(MaxSpace) As item
           If MaxSpace > hashlen Then ExpandHash
    End If
        toplim = toplim + 1
        Malloc = toplim
End Function
Property Get count()
count = toplim + 1
End Property
Friend Function Find2(Key As String, num As Long, Optional checktype As Boolean, Optional isglobal As Boolean) As Boolean
Dim k As Long
num = 0

If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then
    num = PriveSpace(k).ndx: Find2 = True: index = k:  checktype = PriveSpace(k).originaltype: isglobal = PriveSpace(k).globalvar: Exit Function
    End If
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function
Friend Function Find3(Key As String, num As Long, isglobal As Boolean) As Boolean
Dim k As Long
num = 0

If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then
    num = PriveSpace(k).ndx: Find3 = True: index = k: isglobal = PriveSpace(k).globalvar: Exit Function
    End If
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function
Friend Function Find(Key As String, num As Long) As Boolean
Dim k As Long
num = 0

If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then
    num = PriveSpace(k).ndx: Find = True: index = k: Exit Function
    End If
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function
Friend Function ExistKey(Key As String) As Boolean
Dim k As Long

If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then ExistKey = True: Exit Function
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function

Friend Function findRebound(Key As String, newnum As Long, v()) As Boolean
Dim k As Long
If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
With PriveSpace(k)
    If .Key = Key Then
    If .KeyType = 0 Then
        If IsObject(v(.ndx)) Then Set v(.ndx) = Nothing
        v(.ndx) = Empty
        .KeyType = 1
    End If
    .ndx = newnum
     findRebound = True: Exit Function
    End If
    End With
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function
Friend Sub ItemCreator(Key As String, num As Long, Optional isref As Boolean = False, Optional isglobal As Boolean = False, Optional deftype As Boolean = False)
Dim a As Long
a = Malloc()

On Error GoTo exxx
place HashFunc(Key), a
With PriveSpace(a)
  '  .Key = Key
    .ndx = num
    .KeyType = CByte(Abs(isref))
    .originaltype = CByte(deftype)
    .globalvar = CByte(isglobal)
    .firsthash = HashFunc0(Key)
     SwapStrings .Key, Key
End With

exxx:
End Sub
Property Get Percent()
Percent = count / hashlen
End Property
Private Sub place(ByVal b, ByVal a)
    Dim k As Long
    k = Hash(b)
    If Not Hash(b) = a + 1 Then
         Hash(b) = a + 1
         PriveSpace(a).Pleft = k
    End If
    PriveSpace(a).lastpos = b
End Sub
Private Sub ab(ByVal a As Long, b As Long)
b = (b + a + 1024 * a) Xor &H83F3CF
End Sub
Public Function HashFunc0(a$) As Long
Dim i&, p2 As Long, p1 As Integer, p4 As Long, k As Long, p3 As Long
  p4 = Len(a$) * 2 '- 2
  p2 = StrPtr(a$): p4 = p2 + p4 - 2
   For i = p4 To p4 - 4 Step -2
  GetMem2 i, p1
  p3 = UINT(p1)
  HashFunc0 = (HashFunc0 + p3 + 1024& * p3) Xor &H83F3CF
  Next i
  k = 2
  For i = p4 - 6 To p2 Step -2
  GetMem2 i, p1
  p3 = UINT(p1)
  HashFunc0 = (HashFunc0 + p3 + 1024& * p3) Xor &H83F3CF
 i = i - 2 * k
  k = k + 1
  Next i
 
End Function
Private Function UINT(ByVal a As Integer) As Long 'δίνει έναν integer σαν Unsigned integer σε long
 Dim b As Integer
 b = a And &HFFFF
 If b < 0 Then
 UINT = CLng(&H10000 + b)
 Else
 UINT = CLng(b)
 End If
 
 End Function
Private Function HashFunc2(where As Long)
HashFunc2 = MyMod(uintnew(PriveSpace(where).firsthash Or &HAFFAFF7 + Len(PriveSpace(where).Key) + hashlen + 1), hashlen)
End Function

Private Function HashFunc(a$)
HashFunc = MyMod(uintnew(HashFunc0(a$) Or &HAFFAFF7 + Len(a$) + hashlen + 1), hashlen)
End Function
Private Function MyMod(r1, po) As Variant
MyMod = r1 - Fix(r1 / po) * po
End Function
Private Function uintnew(ByVal a As Currency) As Currency
If a > 2147483647@ Then a = 2147483647@
If a < -2147483648@ Then a = -2147483648@
If a < 0 Then
uintnew = 4294967296@ + a
Else
uintnew = a
End If
End Function

