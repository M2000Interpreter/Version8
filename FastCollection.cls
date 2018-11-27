VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FastCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' George's FastCollection
' Only Add, and Find.
' Key as string may be a hex number (used for enumarates)
Private Declare Function CompareString Lib "KERNEL32" Alias "CompareStringW" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As Long, ByVal cchCount1 As Long, ByVal lpString2 As Long, ByVal cchCount2 As Long) As Long
Public NowDec$
' No same key. --- this is changed
Private AllowSameKey As Boolean
' We can remove a key - if AllowSameKey=false

' We can add key (only unique keys)
' return variant (can be anything)
'**************
'* Propertiea:
'*      Count
'*'* from 1 to Count we can access all the members of FastCollection
'*     .Index=1  used for input/output
'* We can get Key or Value
'*      .KeyToString
'*      .Value
'* We can change Valus
'*    set .ValueObj = objectX
'*     .ValueVar
'**************************
'* Methods
'* AddKey "10" [, variant]
'* ToKey "10"
'* SwapValues "10", "30"
'*
'* .Done ..operation is ok
'* .Remove 10
'* .AddKey without parameter add 1 to lastkey
'* .LastKey
'* .LastIndex
'* .Frwd  (for index, exit if not .done)
'* .Back  (for index, exit if not .done)
Private mSortstyle As VbCompareMethod
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Type item
    Key As String
    KeyType As Integer
    iValue As Variant
    sValue As Long
    firsthash As Long
    lastpos As Long
    Pleft As Long  ' a list
End Type
Private PriveSpace() As item
Dim MaxSpace As Long
Dim hashlen As Long
Dim toplim As Long
Public Stable As Long
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal addr As Long, retval As Integer)
Private lastfind As String
Private lasttype As Integer
Private lastkey As String
Private Hash() As Long
Public Done As Boolean
Public index As Long
Public StructLen As Long
Private ParentIndex As Long
Public NumericSort As Boolean
Public IamBusy As Boolean
Public useclid As Long
Private UseUcaseKeys As Boolean
Private Sub Class_Initialize()
    MaxSpace = 5
    ReDim PriveSpace(MaxSpace) As item, Hash(MaxSpace * 2 + 3)
    hashlen = MaxSpace * 2 + 3
    toplim = -1
    mSortstyle = vbTextCompare
    NowDec$ = DefaultDec$
End Sub
Public Sub ResetGui()
    MaxSpace = 50
    ReDim PriveSpace(MaxSpace) As item, Hash(MaxSpace * 2 + 3)
    hashlen = MaxSpace * 2 + 3
    toplim = -1
    mSortstyle = vbTextCompare
End Sub
Public Sub SetBinaryCompare()
mSortstyle = vbBinaryCompare
End Sub
Public Sub SetTextCompare()
mSortstyle = vbTextCompare
End Sub
Public Sub SetDataBaseCompare()
mSortstyle = vbDatabaseCompare
End Sub


Public Sub AddKey2(RHS, Optional aValue As Variant)
IamBusy = True
    index = -1
    
    lastkey = Normalize(RHS)
    If Len(lastkey) = 0 Then RHS = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0): lastkey = RHS
    Done = False
    lasttype = VarType(RHS)
    If Not IsMissing(aValue) Then
    
   
        If IsObject(aValue) Then
        Value = -1
        Set ValueObj = aValue
        Else
        Value = aValue
        End If
    Else
        Value = aValue
    End If
    
IamBusy = False
End Sub
Public Sub AddKey(RHS, Optional aValue As Variant)
    index = -1
    
     lastkey = Normalize(RHS)
 
    If Len(lastkey) = 0 Then RHS = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0): lastkey = RHS
    Done = False
    lasttype = VarType(RHS)
    If Not IsMissing(aValue) Then
    
   
        If IsObject(aValue) Then
        Value = -1
        Set ValueObj = aValue
        Else
        Value = aValue
        End If
    Else
        Value = aValue
        If index >= 0 Then
        PriveSpace(index).KeyType = lasttype
        Else
        PriveSpace(index).KeyType = 0 ' invalid
        End If
    End If
End Sub
Public Sub Sort()
If toplim = -1 Then Exit Sub
If AllowSameKey Then
If Stable Then
InsertionSort PriveSpace(), LBound(PriveSpace()), toplim
Else
QuickSortItemsNumbers PriveSpace(), LBound(PriveSpace()), toplim
End If
Else
If NumericSort Then
QuickSortItemsNumbers PriveSpace(), LBound(PriveSpace()), toplim
Else
QuickSortItems PriveSpace(), LBound(PriveSpace()), toplim
End If
End If
ReHash
End Sub
Public Sub SortDes()

If toplim = -1 Then Exit Sub
If AllowSameKey Then
If Stable Then
InsertionSort PriveSpace(), LBound(PriveSpace()), toplim, 1
Else
QuickSortItemsRevNumbers PriveSpace(), LBound(PriveSpace()), toplim
End If
Else
If NumericSort Then
QuickSortItemsRevNumbers PriveSpace(), LBound(PriveSpace()), toplim
Else
QuickSortItemsRev PriveSpace(), LBound(PriveSpace()), toplim
End If
End If
ReHash
End Sub


Public Sub ToKey(RHS)
    Done = Find(RHS)
    If Done Then lastkey = Normalize(RHS)
End Sub
Public Sub RemoveWithNoFind()
If AllowSameKey Then Exit Sub
Dim k As Long, k1 As Long, vvv As Variant, top1 As Long
If toplim < 0 Then Exit Sub
top1 = toplim
If Done Then
       If ParentIndex > -1 Then
            PriveSpace(ParentIndex).Pleft = PriveSpace(index).Pleft

        Else
        
        Hash(PriveSpace(index).lastpos) = PriveSpace(index).Pleft

        End If
        SwapVariant vvv, PriveSpace(top1).iValue
        PriveSpace(index).iValue = -1
        PriveSpace(index) = PriveSpace(top1)
        Dim new_item As item
        PriveSpace(top1) = new_item
        SwapVariant vvv, PriveSpace(index).iValue
        k = Hash(PriveSpace(index).lastpos) - 1
        If k = top1 Then
            ' we have to give the new position to Hash()
            Hash(PriveSpace(index).lastpos) = index + 1
        Else
            Do While k >= 0 And k <> top1 And k <> k1
                   k1 = k
                   k = PriveSpace(k).Pleft - 1
            Loop
            If k = k1 Then
               ReHash
            ElseIf k = top1 Then
               PriveSpace(k1).Pleft = index + 1
            End If
        End If
  
here:
       toplim = toplim - 1
      
        
       
        
  ReduceHash toplim
End If
End Sub
Public Sub drop(RHS)
If Not AllowSameKey Then Exit Sub
RHS = RHS - 1
If (toplim - RHS) >= 0 And (RHS >= 0) Then
Dim i As Long
For i = toplim To RHS Step -1
       With PriveSpace(i)
        If .Pleft > 0 Then
            Hash(.lastpos) = .Pleft
            .Key = vbNullString
            .iValue = CDbl(0)
            .Pleft = 0
        Else
            Hash(.lastpos) = 0
            .Key = vbNullString
        End If
        End With
toplim = RHS - 1
  ReduceHash toplim
Next i
Else
   MaxSpace = 5
    ReDim PriveSpace(MaxSpace) As item, Hash(MaxSpace * 2 + 3)
    hashlen = MaxSpace * 2 + 3
    toplim = -1
    mSortstyle = vbTextCompare
End If

End Sub
Public Sub Remove(RHS)
If AllowSameKey Then Exit Sub
Dim k As Long, k1 As Long, vvv As Variant, top1 As Long, ndx As Long, Parndx As Long
If toplim < 0 Then Exit Sub
top1 = toplim
If Find2(RHS, Parndx, ndx) Then
        Done = True
        index = ndx
        ' here we make a clean to pointers to PriveSpace, so we exclude the removing item
        If Parndx > -1 Then
            ' if we have a parent then just push to parent the Pleft pointer
            ' Pleft is always +1 (so Pleft 0 give -1 null for this context)
            PriveSpace(Parndx).Pleft = PriveSpace(ndx).Pleft

        Else
        ' so we have no parent, just out Pleft to Hash
        Hash(PriveSpace(ndx).lastpos) = PriveSpace(ndx).Pleft

        End If
        ' we don't have to see Hash() any more for RHS
        ' but we have to see Hash() for top1
        Dim new_item As item
        ' We have to get change top1 to ndx, maybe we remove top1
        ' we can check this but it is faster to not check this (more often top1 isn't ndx)
        ' Using SwapVariant we can swap any kind of Variant, including objects (swaping pointers only)
        SwapVariant vvv, PriveSpace(top1).iValue
        PriveSpace(ndx).iValue = -1 ' is good to pass a number - so we break any pointer to object
        PriveSpace(ndx) = PriveSpace(top1) ' now we perform a copy (no objects included)
        PriveSpace(top1) = new_item
        ' so now we put back iValue (maybe is an object)
        SwapVariant vvv, PriveSpace(ndx).iValue
        ' here Lastpos is from previous PriveSpace(top1) which is PriveSpace(ndx)
        k = Hash(PriveSpace(ndx).lastpos) - 1
        If k = top1 Then
            ' we have to give the new position to Hash()
            Hash(PriveSpace(ndx).lastpos) = ndx + 1
        Else
         
         Do While k >= 0 And k <> top1 And k <> k1
                k1 = k
                k = PriveSpace(k).Pleft - 1
         Loop
         If k = k1 Then
            ReHash
         ElseIf k = top1 Then
            PriveSpace(k1).Pleft = ndx + 1
         End If
        End If
  
here:
       toplim = toplim - 1
       ReduceHash toplim
Else
Done = False
End If
End Sub

Property Let Value(RHS As Variant)
Done = False
If index = -1 Then
ItemCreator lastkey, RHS

Else
    PriveSpace(index).iValue = RHS
End If
Done = True
End Property
Property Set ValueObj(RHS As Variant)
Done = False
If index = -1 Then
    ItemCreator lastkey, RHS
Else
    Set PriveSpace(index).iValue = RHS
End If
Done = True
End Property
Property Get sValue() As Long
Done = False
If index = -1 Then
Else
    Done = True
    sValue = PriveSpace(index).sValue
End If
End Property
Property Let sValue(RHS As Long)
If index = -1 Then
Else
    Done = True
    PriveSpace(index).sValue = RHS
End If
End Property
Property Get KeyTypeValue() As Integer
Done = False
If index = -1 Then
Else
    Done = True
    KeyTypeValue = PriveSpace(index).KeyType
End If
End Property
Property Let KeyTypeValue(RHS As Integer)
If index = -1 Then
Else
    Done = True
    PriveSpace(index).KeyType = RHS
End If
End Property

Property Get Value() As Variant
Done = False
If index = -1 Then
ElseIf Not IsObject(PriveSpace(index).iValue) Then
    Done = True
    Value = PriveSpace(index).iValue
    
    If Typename(Value) = "Error" Then
    
    Value = MyVal()
    If VarType(Value) = vbString Then Value = PriveSpace(index).Key
    
    End If
End If
End Property
Private Function MyVal() As Variant
On Error Resume Next
MyVal = PriveSpace(index).Key
Select Case PriveSpace(index).KeyType
Case vbString
If MyVal Like "[-0-9.]*" Then
    MyVal = CDbl(Replace$(MyVal, ".", NowDec$))
Else
    MyVal = vbNullString
End If
Case vbDouble
MyVal = CDbl(Replace$(MyVal, ".", NowDec$))
Case vbCurrency
MyVal = CCur(Replace$(MyVal, ".", NowDec$))
Case vbDecimal
MyVal = CDec(Replace$(MyVal, ".", NowDec$))
Case Else
MyVal = CDbl(Replace$(MyVal, ".", NowDec$))
End Select

End Function
Function ValueType(ch As Long, p As Variant, s As String) As Boolean
Done = False
Dim val1 As Variant
If index = -1 Then
ValueType = False
ElseIf Not IsObject(PriveSpace(index).iValue) Then
ValueType = True
    Done = True
    val1 = PriveSpace(index).iValue
    If Typename(val1) = "Error" Then
        If ch = 2 Then
            s = KeyToString
        Else
        p = 0
            On Error Resume Next
            p = KeyToNumber
        End If
    Else
        If ch = 2 Then
            s = val1
        Else
            p = 0
            On Error Resume Next
            p = val1
        End If
    End If
End If
End Function
Property Get ValueObj() As Variant
Done = False
If index = -1 Then
ElseIf IsObject(PriveSpace(index).iValue) Then
    Set ValueObj = PriveSpace(index).iValue
    Done = True
End If

End Property
Private Sub ExpandHash()
    hashlen = MaxSpace * 2 + 3
    ReHash
End Sub
Private Sub ReHash()
    Dim i As Long
    ReDim Hash(hashlen) As Long
    For i = 0 To toplim
           place HashFunc2(i), i
        
    Next i
End Sub
Private Function CheckHash() As Boolean
    Dim i As Long
    For i = 0 To toplim
        If PriveSpace(i).Pleft > 0 Then
        If PriveSpace(PriveSpace(i).Pleft - 1).Key = vbNullString Then Exit Function
        If i = (PriveSpace(i).Pleft - 1) Then Exit Function
        End If
    Next i
    CheckHash = True
End Function


Private Sub ReduceHash(newTop As Long)
    If newTop <= 5 Then
            If newTop < 0 Then
                newTop = -1
                toplim = -1
                MaxSpace = 5
                ReDim PriveSpace(MaxSpace) As item
                
                ReDim Hash(MaxSpace * 2 + 3)
                hashlen = MaxSpace * 2 + 3
            ElseIf MaxSpace > 40 Then
                MaxSpace = 5
                ReDim Preserve PriveSpace(MaxSpace) As item
                
                ReDim Hash(MaxSpace * 2 + 3)
                hashlen = MaxSpace * 2 + 3
                
                ReHash
            End If
    Else
            If MaxSpace - newTop + 1 > 2 * newTop + 2 Then
            MaxSpace = 2 * (newTop + 1) + 1
            ReDim Preserve PriveSpace(MaxSpace) As item
            ReDim Hash(MaxSpace * 2 + 3)
            hashlen = MaxSpace * 2 + 3
            ReHash
    End If
   End If
End Sub
Private Function Normalize(Key) As String

        If VarType(Key) = vbString Then
        If UseUcaseKeys Then
            Normalize = myUcase(Key)
        Else
            Normalize = Key
        End If
        ElseIf VarType(Key) = vbBoolean Then
        Normalize = Str(CLng(Key))
        Else
        Normalize = LTrim$(Str$(Key))
        End If
End Function
Private Function Normalize2(Key) As String
        If VarType(Key) = vbString Then
        Normalize2 = myUcase(Key)
        ElseIf VarType(Key) = vbBoolean Then
        Normalize2 = Str(CLng(Key))
        Else
        Normalize2 = myUcase(LTrim$(Str$(Key)))
        End If
End Function
Function ChangeKey(key1, key2)
If AllowSameKey Then Exit Function
    Dim Key As String, NewKey As String
    Dim k As Long
    If key1 = key2 Then ChangeKey = True
    Key = Normalize(key1)
    If Len(Key) = 0 Then Exit Function
    If ExistKey(key2) Then Exit Function
    
    NewKey = Normalize(key2)
    If Len(NewKey) = 0 Then Exit Function
    
    k = Hash(HashFunc(Key)) - 1
    If k >= 0 Then
    Do
            If PriveSpace(k).Key = Key Then
                PriveSpace(k).Key = NewKey
                PriveSpace(k).firsthash = HashFunc0(NewKey)
                ReDim Hash(MaxSpace * 2 + 3)
                hashlen = MaxSpace * 2 + 3
                ReHash
                ChangeKey = True
                Exit Function
            End If
            k = PriveSpace(k).Pleft - 1
    Loop Until k < 0
    End If
End Function
Sub swap(key1, key2)
Dim k As Long, c As item
If Find(key1) Then
k = index
If Find(key2) Then

    c = PriveSpace(k): PriveSpace(k) = PriveSpace(index): PriveSpace(index) = c
    index = k
    ReHash
End If
End If
End Sub
Sub SwapValues(key1, key2)
Dim k As Long
If Find(key1) Then
k = index
If Find(key2) Then
    SwapVariant PriveSpace(k).iValue, PriveSpace(index).iValue
End If
End If
End Sub
Sub SwapNext(key1)
Dim k As Long
On Error Resume Next
If Done Then
    k = index
    If Find(key1) Then
        SwapVariant PriveSpace(k).iValue, PriveSpace(index).iValue
    End If
End If
End Sub
Sub SwapByIndex(RHS As Long)
On Error Resume Next
If Done Then
    SwapVariant PriveSpace(RHS).iValue, PriveSpace(index).iValue
    index = RHS
End If
End Sub


Private Function Malloc() As Long
    If toplim + 1 >= MaxSpace Then
        MaxSpace = MaxSpace * 2
        ReDim Preserve PriveSpace(MaxSpace) As item
        If MaxSpace > hashlen * 3 / 4 Then ExpandHash
    End If
    toplim = toplim + 1
    Malloc = toplim
End Function
Property Get count()
    count = toplim + 1
End Property
Public Property Let UcaseKeys(RHS)
If toplim = -1 Then
UseUcaseKeys = CBool(RHS)
End If
End Property

Property Get IsEmpty()
    IsEmpty = toplim = -1
End Property
Property Get LastKnown() As String
    LastKnown = lastfind
End Property
Function IsObj() As Boolean
If index = -1 Then
Else
IsObj = IsObject(PriveSpace(index).iValue)
End If
End Function

Function Find(RHS) As Boolean
Dim k As Long, Key As String, k1 As Long
Key = Normalize(RHS)
If Len(Key) = 0 Then Key = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
ParentIndex = -1
Done = False
k = Hash(HashFunc(Key)) - 1
k1 = -2
If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then Find = True: lastfind = Key: index = k: Done = True: Exit Function
     ParentIndex = k
     k1 = k
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0 Or k1 = k
If k1 = k Then
ReHash
k = Hash(HashFunc(Key)) - 1
k1 = -2
ParentIndex = -1
If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then Find = True: lastfind = Key: index = k: Done = True: Exit Function
     ParentIndex = k
     k1 = k
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0 Or k = k1
End If
End If
End If

End Function

Private Function Find2(RHS, Parndx As Long, ndx As Long) As Boolean
Dim k As Long, Key As String, k1 As Long
Key = Normalize(RHS)
If Len(Key) = 0 Then Key = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
Parndx = -1
k = Hash(HashFunc(Key)) - 1
k1 = -2
If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then Find2 = True: ndx = k: Exit Function
     Parndx = k
     k1 = k
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0 Or k1 = k
If k1 = k Then
Parndx = -1
k = Hash(HashFunc(Key)) - 1
k1 = -2
If k >= 0 Then
Do
    If PriveSpace(k).Key = Key Then Find2 = True: ndx = k: Exit Function
     Parndx = k
     k1 = k
     k = PriveSpace(k).Pleft - 1
Loop Until k < 0 Or k1 = k
End If
End If
End If
End Function

Function ExistKey0(RHS) As Boolean
If AllowSameKey Then Exit Function
Dim k As Long, Key As String, k1 As Long
Key = Normalize(RHS)
If Len(Key) = 0 Then Key = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
k = Hash(HashFunc(Key)) - 1
k1 = -2
If k >= 0 Then
    Do
        If PriveSpace(k).Key = Key Then ExistKey0 = True: index = k: Exit Function
        k1 = k
        k = PriveSpace(k).Pleft - 1
    Loop Until k < 0 Or k1 = k
    If k = k1 Then
        ReHash
        k = Hash(HashFunc(Key)) - 1
        k1 = -2
        If k >= 0 Then
        Do
            If PriveSpace(k).Key = Key Then ExistKey0 = True: index = k: Exit Function
            k1 = k
            k = PriveSpace(k).Pleft - 1
        Loop Until k < 0 Or k1 = k
        End If
    End If
End If
End Function
Function ExistKey(RHS) As Boolean
Dim k As Long, k1 As Long, Key As String
Key = Normalize(RHS)
If Len(Key) = 0 Then Key = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
k = Hash(HashFunc(Key)) - 1
k1 = -2
If k >= 0 Then
    Do
        If PriveSpace(k).Key = Key Then ExistKey = True: index = k: Exit Function
        k1 = k
        k = PriveSpace(k).Pleft - 1
        
    Loop Until k < 0 Or k = k1
    If k = k1 Then
        ReHash
        k = Hash(HashFunc(Key)) - 1
        If k >= 0 Then
        Do
            If PriveSpace(k).Key = Key Then ExistKey = True: index = k: Exit Function
            k1 = k
            k = PriveSpace(k).Pleft - 1
        Loop Until k < 0 Or k = k1
        End If
    End If
End If
End Function
Function ExistKey3(RHS) As Boolean
Dim k As Long, k1 As Long, Key As String
Key = Normalize(RHS)
If Len(Key) = 0 Then Key = Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0)
k1 = -2
k = Hash(HashFunc(Key)) - 1
If k >= 0 Then
    Do
        If PriveSpace(k).Key = Key Then ExistKey3 = True: Exit Function
        k1 = k
        k = PriveSpace(k).Pleft - 1
    Loop Until k < 0 Or k = k1
    If k = k1 Then
        ReHash
        k = Hash(HashFunc(Key)) - 1
        If k >= 0 Then
            Do
                If PriveSpace(k).Key = Key Then ExistKey3 = True: Exit Function
                k1 = k
                k = PriveSpace(k).Pleft - 1
            Loop Until k < 0 Or k = k1
        End If
    End If
End If
End Function

Private Function ExistKey2(Key As String, nkey As Long) As Boolean
Dim k As Long, k1 As Long
If Len(Key) = 0 Then ExistKey2 = True: Exit Function
k = Hash(HashFunc1(nkey, Key)) - 1
k1 = -2
If k >= 0 Then
    Do
        If PriveSpace(k).Key = Key Then ExistKey2 = True: Exit Function
        k1 = k
        k = PriveSpace(k).Pleft - 1
    Loop Until k < 0 Or k = k1
    If k = k1 Then
        ReHash
        k = Hash(HashFunc(Key)) - 1
        If k >= 0 Then
        Do
            If PriveSpace(k).Key = Key Then ExistKey2 = True: Exit Function
            k1 = k
            k = PriveSpace(k).Pleft - 1
            
        Loop Until k < 0 Or k = k1
        End If
    End If

End If
End Function

Private Sub ItemCreator(Key As String, nValue As Variant)
Dim a As Long, kk As Long
Done = False
kk = HashFunc0(Key)
If Not AllowSameKey Then If ExistKey2(Key, kk) Then Exit Sub  ' done is false
a = Malloc()
With PriveSpace(a)
    .Key = Key
    If IsObject(nValue) Then
    Set .iValue = nValue
    Else
    .iValue = nValue
    End If
    .firsthash = kk
End With

place HashFunc1(kk, Key), a

index = a
Done = True
End Sub

Property Get Percent()
Percent = 100 * count / hashlen
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
Public Function HashFunc0(a$) As Long
  Dim i&, p2 As Long, p1 As Integer, p4 As Long
  p4 = Len(a$) * 2
  p2 = StrPtr(a$): p4 = p2 + p4 - 2
  For i = p4 To p2 Step -1
    GetMem2 i, p1
    HashFunc0 = HashFunc0 Xor p1
  Next i
  If HashFunc0 = 0 Then HashFunc0 = 1
End Function
Private Function HashFunc2(where As Long)
HashFunc2 = MyMod(uintnew(PriveSpace(where).firsthash Or &HAFFAFF7) + Len(PriveSpace(where).Key) + hashlen + 1, hashlen)
End Function
Private Function HashFunc1(nkey As Long, a$)
HashFunc1 = MyMod(uintnew(nkey Or &HAFFAFF7) + Len(a$) + hashlen + 1, hashlen)
End Function
Private Function HashFunc(a$)
HashFunc = MyMod(uintnew(HashFunc0(a$) Or &HAFFAFF7) + Len(a$) + hashlen + 1, hashlen)
End Function
Private Function UINT(ByVal a As Integer) As Long
 Dim b As Integer
 b = a And &HFFFF
 If b < 0 Then
 UINT = CLng(&H10000 + b)
 Else
 UINT = CLng(b)
 End If
 
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

Private Sub SwapVariant(ByRef a As Variant, ByRef b As Variant)
   Dim t(0 To 3) As Long ' 4 Longs * 4 bytes each = 16 bytes
   CopyMemory t(0), ByVal VarPtr(a), 16
   CopyMemory ByVal VarPtr(a), ByVal VarPtr(b), 16
   CopyMemory ByVal VarPtr(b), t(0), 16
End Sub
' only for tests
Private Function GetRefCount(ByRef vvv) As Long
    If Not IsObject(vvv) Then Exit Function
    Dim obj As IUnknown
    Set obj = vvv
    If obj Is Nothing Then Exit Function
    CopyMemory GetRefCount, ByVal (ObjPtr(obj)) + 4, 4
    GetRefCount = GetRefCount - 2
    Set obj = Nothing
End Function
Public Property Get ReferCount() As Long
ReferCount = GetRefCount(Me)
End Property
Public Property Get KeyToString2() As Variant
If index > -1 Then

    KeyToString2 = PriveSpace(index).Key + Str$(PriveSpace(index).lastpos < toplim) + Str$(PriveSpace(index).Pleft) ' + Str$(Int(Percent * 1000) / 100) + Str$(hashlen) + Str$(count)
End If
End Property
Public Property Get KeyToString() As Variant
If index > -1 Then

    KeyToString = PriveSpace(index).Key
End If
End Property
Public Property Get HassKeyNumberPart() As Variant
If index > -1 Then
HassKeyNumberPart = Left$(PriveSpace(index).Key, 1) Like "[-0-9.]*"
End If
End Property
Public Property Get KeyToNumber() As Variant
If index > -1 Then
    KeyToNumber = MyVal()
    If VarType(KeyToNumber) = vbString Then KeyToNumber = CDbl(0)
End If
End Property
Public Property Get ReferCountValue() As Long
If index = -1 Then
ElseIf IsObject(PriveSpace(index).iValue) Then
    ReferCountValue = GetRefCount(PriveSpace(index).iValue)
End If
End Property
Public Sub ToStart()
index = 0
Done = count > 0
End Sub
Public Sub ToEnd()
index = count - 1
Done = count > 0
End Sub

Public Sub NextIndex()
Done = False
    If index = -1 Then
 
    Else
        If index < toplim Then Done = True: index = index + 1
    End If

End Sub
Public Function ToNextIndex()
Done = False
    If index = -1 Then
 
    Else
        If index < toplim Then Done = True: index = index + 1
    End If
    ToNextIndex = Done
End Function
Public Sub PrevIndex()
Done = False
    If index = -1 Then
    
    Else
        If index > 0 Then Done = True: index = index - 1
    End If

End Sub
Public Function ToPrevIndex()
Done = False
    If index = -1 Then
    
    Else
        If index > 0 Then Done = True: index = index - 1
    End If
    ToPrevIndex = Done
End Function

Private Sub Class_Terminate()
    Erase PriveSpace()
End Sub

Private Sub QuickSortItems(Arr() As item, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
Dim Piv As item, tmp As item '<- adjust types here, when switching to something different than Long
  If UB - LB = 1 Then
     M1 = LB

     If CompareStr2(Arr(M1).Key, Arr(UB).Key, mSortstyle) = 1 Then tmp = Arr(M1): Arr(M1) = Arr(UB): Arr(UB) = tmp
     Exit Sub
  Else
  ' we have unique names
    M1 = (LB + UB) \ 2
    Piv = Arr(M1)
    M1 = LB
   Do While CompareStr2(Arr(M1).Key, Piv.Key, mSortstyle) = -1: M1 = M1 + 1: Loop
    
  End If
  
  M2 = UB
  Do
    Do While CompareStr2(Arr(M2).Key, Piv.Key, mSortstyle) = 1: M2 = M2 - 1: Loop

    If M1 <= M2 Then
      tmp = Arr(M1): Arr(M1) = Arr(M2): Arr(M2) = tmp 'swap
      M1 = M1 + 1
      M2 = M2 - 1
    End If
    If M1 > M2 Then Exit Do
    Do While CompareStr2(Arr(M1).Key, Piv.Key, mSortstyle) = -1: M1 = M1 + 1: Loop

  Loop
  If LB < M2 Then QuickSortItems Arr, LB, M2
  If M1 < UB Then QuickSortItems Arr, M1, UB
End Sub
Private Sub QuickSortItemsNumbers(Arr() As item, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
On Error GoTo abc1
Dim Piv As item, tmp As item '<- adjust types here, when switching to something different than Long
     If UB - LB = 1 Then
     M1 = LB
       If compareStr4(Arr(M1).Key, Arr(UB).Key) = 1 Then tmp = Arr(M1): Arr(M1) = Arr(UB): Arr(UB) = tmp
      Exit Sub
     Else
       M1 = (LB + UB) \ 2 '+ 1
       Piv = Arr(M1)
             If compareStr4(Arr(M1).Key, Piv.Key) = 0 Then
                M2 = UB - 1
                M1 = LB
                Piv = Arr(LB)
                
                Do
                    M1 = M1 + 1
                    If M1 > M2 Then
                        If compareStr4(Arr(UB).Key, Piv.Key) = -1 Then tmp = Arr(LB): Arr(LB) = Arr(UB): Arr(UB) = tmp
                        Exit Sub
                    End If
                Loop Until compareStr4(Arr(M1).Key, Piv.Key) <> 0
                Piv = Arr(M1)
                If M1 > LB Then If compareStr4(Arr(LB).Key, Piv.Key) = 1 Then tmp = Arr(LB): Arr(LB) = Arr(M1): Arr(M1) = tmp: Piv = Arr(M1)
            Else
                Piv = Arr(M1)
                M1 = LB
                Do While compareStr4(Arr(M1).Key, Piv.Key) = -1: M1 = M1 + 1: Loop
            End If
    End If
    M2 = UB
    Do
      Do While compareStr4(Arr(M2).Key, Piv.Key) = 1: M2 = M2 - 1: Loop
      If M1 <= M2 Then
       If M1 <> M2 Then tmp = Arr(M1): Arr(M1) = Arr(M2): Arr(M2) = tmp
        M1 = M1 + 1
        M2 = M2 - 1
      End If
      If M1 > M2 Then Exit Do
      Do While compareStr4(Arr(M1).Key, Piv.Key) = -1: M1 = M1 + 1: Loop
    Loop
    If LB < M2 Then QuickSortItemsNumbers Arr, LB, M2
    If M1 < UB Then QuickSortItemsNumbers Arr, M1, UB
    Exit Sub
abc1:
    

End Sub

Private Sub QuickSortItemsRevNumbers(Arr() As item, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
On Error GoTo abc1
Dim Piv As item, tmp As item '<- adjust types here, when switching to something different than Long
     If UB - LB = 1 Then
     M1 = LB
       If compareStr4(Arr(M1).Key, Arr(UB).Key) = -1 Then tmp = Arr(M1): Arr(M1) = Arr(UB): Arr(UB) = tmp
      Exit Sub
     Else
       M1 = (LB + UB) \ 2 '+ 1
       Piv = Arr(M1)
             If compareStr4(Arr(M1).Key, Piv.Key) = 0 Then
                M2 = UB - 1
                M1 = LB
                Piv = Arr(LB)
                
                Do
                    M1 = M1 + 1
                    If M1 > M2 Then
                        If compareStr4(Arr(UB).Key, Piv.Key) = 1 Then tmp = Arr(LB): Arr(LB) = Arr(UB): Arr(UB) = tmp
                        Exit Sub
                    End If
                Loop Until compareStr4(Arr(M1).Key, Piv.Key) <> 0
                Piv = Arr(M1)
                If M1 > LB Then If compareStr4(Arr(LB).Key, Piv.Key) = -1 Then tmp = Arr(LB): Arr(LB) = Arr(M1): Arr(M1) = tmp: Piv = Arr(M1)
            Else
                Piv = Arr(M1)
                M1 = LB
                Do While compareStr4(Arr(M1).Key, Piv.Key) = 1: M1 = M1 + 1: Loop
            End If
    End If
    M2 = UB
    Do
      Do While compareStr4(Arr(M2).Key, Piv.Key) = -1: M2 = M2 - 1: Loop
      If M1 <= M2 Then
       If M1 <> M2 Then tmp = Arr(M1): Arr(M1) = Arr(M2): Arr(M2) = tmp
        M1 = M1 + 1
        M2 = M2 - 1
      End If
      If M1 > M2 Then Exit Do
      Do While compareStr4(Arr(M1).Key, Piv.Key) = 1: M1 = M1 + 1: Loop
    Loop
    If LB < M2 Then QuickSortItemsRevNumbers Arr, LB, M2
    If M1 < UB Then QuickSortItemsRevNumbers Arr, M1, UB
    Exit Sub
abc1:
    
End Sub
Private Sub QuickSortItemsRev(Arr() As item, ByVal LB As Long, ByVal UB As Long)
Dim M1 As Long, M2 As Long
Dim Piv As item, tmp As item '<- adjust types here, when switching to something different than Long
  If UB - LB = 1 Then
     M1 = LB

     If CompareStr2(Arr(M1).Key, Arr(UB).Key, mSortstyle) = -1 Then tmp = Arr(M1): Arr(M1) = Arr(UB): Arr(UB) = tmp
     Exit Sub
  Else
  ' we have unique names
    M1 = (LB + UB) \ 2
    Piv = Arr(M1)
    M1 = LB
   Do While CompareStr2(Arr(M1).Key, Piv.Key, mSortstyle) = 1: M1 = M1 + 1: Loop
    
  End If
  
  M2 = UB
  Do
    Do While CompareStr2(Arr(M2).Key, Piv.Key, mSortstyle) = -1: M2 = M2 - 1: Loop

    If M1 <= M2 Then
      tmp = Arr(M1): Arr(M1) = Arr(M2): Arr(M2) = tmp 'swap
      M1 = M1 + 1
      M2 = M2 - 1
    End If
    If M1 > M2 Then Exit Do
    Do While CompareStr2(Arr(M1).Key, Piv.Key, mSortstyle) = 1: M1 = M1 + 1: Loop

  Loop
  If LB < M2 Then QuickSortItemsRev Arr, LB, M2
  If M1 < UB Then QuickSortItemsRev Arr, M1, UB

End Sub
Public Function AllowAnyKey()
AllowSameKey = True
Stable = True
End Function

Function compareStr4(ByVal a$, ByVal b$) As Long
Dim i As Long, j As Long, a1$, b1$, p1 As Variant, p2 As Variant, n$, k As Long, k1 As Long
If CompareString(useclid, &H1000, StrPtr(a$), Len(a$), StrPtr(b$), Len(b$)) = 2 Then Exit Function
n$ = "[-0-9.]"
k = Sgn(Len(a$) - Len(b$))
k1 = k
again:
j = IIf(Len(a$) >= Len(b$), Len(b$), Len(a$))

For i = 1 To j

    If Mid$(a$, i, 1) Like n$ Then
    a1$ = Mid$(a$, i)
    b1$ = Mid$(b$, i)
        IsNumberD2 a1$, p1
        If IsNumberD2(b1$, p2) Then
            compareStr4 = Sgn(MyRound(p1, 8) - MyRound(p2, 8))
            If compareStr4 = 0 Then
                If Len(a1$) * Len(b1$) <> 0 Then
                    If Len(n$) > 6 Then n$ = "[-0-9]"
                    k = Sgn(Len(a1$) - Len(b1$))
                     
                    a$ = a1$
                    b$ = b1$
                   
                    GoTo again
                End If
                If Len(a1$) + Len(b1$) = 0 Then
                    compareStr4 = k
                Else
                    compareStr4 = Sgn(Len(a1$) - Len(b1$))
                End If
                Exit Function
            End If
            Exit Function
        Else
        a$ = Mid$(a$, i)
    b$ = Mid$(b$, i)
    IsNumberD2 a$, p1
       
          If Len(a$) > 0 Then
       
            Select Case CompareString(useclid, &H1000, StrPtr(a$), Len(a$), StrPtr(b$), Len(b$))
                Case 1, 3
                compareStr4 = -1
                Case 2
                If k > 0 Then
                compareStr4 = -1
                Else
                compareStr4 = 1
                End If
            End Select
        Else
        compareStr4 = -1
        End If
        End If
        Exit Function
    ElseIf Mid$(b$, i, 1) Like n$ Then
    a$ = Mid$(a$, i)
    b$ = Mid$(b$, i)
    IsNumberD2 b$, p1
    If Len(b$) > 0 Then
       
            Select Case CompareString(useclid, &H1000, StrPtr(a$), Len(a$), StrPtr(b$), Len(b$))
                Case 3
                compareStr4 = 1
                Case 1, 2
                If k < 0 Then
                compareStr4 = 1
                Else
                compareStr4 = -1
                End If
                
                
            End Select
        Else
        compareStr4 = 1
        End If
        Exit Function
    End If
    compareStr4 = -2 + CompareString(useclid, 0, StrPtr(a$) + i * 2 - 2, 1, StrPtr(b$) + i * 2 - 2, 1)
    If compareStr4 <> 0 Then
    Exit Function
    End If
Next i
compareStr4 = Sgn(Len(a$) - Len(b$))
If compareStr4 = 0 Then compareStr4 = k1
End Function


Function CompareStr2(a$, b$, what As Long) As Long
On Error GoTo comperr
Select Case what
Case 1 To 2
CompareStr2 = StrComp(a$, b$, what)
Case Else
If useclid = 0 Then
CompareStr2 = StrComp(a$, b$)
Else
' we cant use something that give "equal" keys
CompareStr2 = -2 + CompareString(useclid, &H1000, StrPtr(a$), Len(a$), StrPtr(b$), Len(b$))
End If
End Select
Exit Function
comperr:
CompareStr2 = -2 + CompareString(0, 0, StrPtr(a$), Len(a$), StrPtr(b$), Len(b$))

End Function

Public Property Get IsQueue() As Variant
IsQueue = AllowSameKey
End Property
Private Sub InsertionSort(Arr() As item, ByVal iMin As Long, ByVal iMax As Long, Optional what As Long = -1)
    Dim i As Long
    Dim j As Long
    Dim varSwap As item
    i = iMin + 1
    iMax = iMax + 1
    While i < iMax
        j = i
        Do While j > iMin
        If compareStr4(Arr(j).Key, Arr(j - 1).Key) = what Then
            varSwap = Arr(j): Arr(j) = Arr(j - 1): Arr(j - 1) = varSwap
            j = j - 1
        Else
            Exit Do
        End If
        Loop
        i = i + 1
    Wend
End Sub

