VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sbHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 Private Type item
 Key As String
  ndx As Long
  KeyType As Long  ' 0 as is , 1, reference/don't delete
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
Private lastfind As String
Private Hash() As Long
Public index As Long
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


Private Sub ExpandHash()

    Dim i As Long
    hashlen = hashlen * 2 + 3
    
    ReDim Hash(hashlen) As Long
    For i = 0 To MaxSpace \ 2 - 1

    place HashFunc2(i), i
    Next i
End Sub
Friend Sub ReduceHash(newTop As Long, v() As modfun)
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
            v(.ndx).sb = vbNullString
            v(.ndx).sbc = 0
            v(.ndx).sbgroup = vbNullString
            v(.ndx).goodname = vbNullString
            v(.ndx).locked = False
            Set v(.ndx).subs = Nothing
            End If
        .KeyType = 0
        .ndx = -1
            
        End With

    Next i
    If newTop <= 50 Then
            If MaxSpace > 50 Then
                        MaxSpace = 50
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
            
    Else
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
   End If
End Sub
Function RenameKey(Key As String, NewKey As String)
Dim k As Long

If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then
    
    PriveSpace(k).Key = NewKey
    PriveSpace(k).firsthash = HashFunc0(NewKey)
                ReDim Hash(MaxSpace * 2 + 3)
               hashlen = MaxSpace * 2 + 3
             '  reHash...
            For k = 0 To toplim
                place HashFunc2(k), k
            Next k

    
            RenameKey = True
            Exit Function
               
 End If
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0

 End If
            
End Function
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
Property Get LastKnown() As String
LastKnown = lastfind
End Property

Friend Function Find(Key As String, num As Long) As Boolean
Dim k As Long
num = 0
If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then
    num = PriveSpace(k).ndx: Find = True: lastfind = Key: index = k: Exit Function
    End If
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0
End If
End Function
Friend Function Find2(Key As String, num As Long, internal As Long) As Boolean
'' used with KeyType as Long not a flag
Dim k As Long
internal = 0
num = 0
If Len(Key) = 0 Then Exit Function
k = Hash(HashFunc(Key)) - 1

If k >= 0 Then
Do
    With PriveSpace(k)
    If .Key = Key Then
    internal = .KeyType
    num = .ndx: Find2 = True: lastfind = Key: index = k: Exit Function
    End If
     k = .Pleft - 1
     End With
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


Friend Sub ItemCreator(Key As String, num As Long, Optional isref As Boolean = False)
Dim a As Long
a = Malloc()
place HashFunc(Key), a
With PriveSpace(a)
    .ndx = num
    If isref Then
        .KeyType = 1
    Else
        .KeyType = 0
    End If
    .firsthash = HashFunc0(Key)
    SwapStrings .Key, Key
End With
End Sub
Friend Sub ItemCreatorNoSwap(Key As String, num As Long, Optional isref As Boolean = False)
Dim a As Long
a = Malloc()
place HashFunc(Key), a
With PriveSpace(a)
    .ndx = num
    If isref Then
        .KeyType = 1
    Else
        .KeyType = 0
    End If
    .firsthash = HashFunc0(Key)
    .Key = Key
End With



End Sub
Friend Sub ItemCreator2(Key As String, num As Long, internal As Long)
'' used with KeyType as Long not a flag
Dim a As Long
a = Malloc()
place HashFunc(Key), a
With PriveSpace(a)
    .ndx = num
    .KeyType = internal
    .firsthash = HashFunc0(Key)
     SwapStrings .Key, Key
End With

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
Function Show() As String
Dim i As Long
For i = 0 To toplim
If InStr(PriveSpace(i).Key, ChrW(&H1FFF)) = 0 And InStr(PriveSpace(i).Key, ChrW(&HFFBF)) = 0 Then
Show = Chr(1) + PriveSpace(i).Key + Str$(PriveSpace(i).ndx) + Chr(1) + Show
End If
Next i
End Function
Function ShowRev() As String
Dim i As Long
If toplim > UBound(PriveSpace()) Then
toplim = UBound(PriveSpace()) - 1
InternalEror
Else
For i = 0 To toplim
' remove if to take hidden functions
If InStr(PriveSpace(i).Key, ChrW(&H1FFF)) = 0 And InStr(PriveSpace(i).Key, ChrW(&HFFBF)) = 0 Then
ShowRev = ShowRev + Chr(1) + PriveSpace(i).Key + Str$(PriveSpace(i).ndx) + Chr(1)
End If
Next i
End If
End Function

