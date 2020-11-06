VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PropReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim ObjectRef As Long
Dim vtref As Long
Dim mIndex As Variant
Public UseIndex As Boolean
Dim Init As Boolean
Dim IamMethod  As Boolean
Dim hardlink As Object, useobj As Boolean
Private mLastObj As Object
Public Arr As Boolean
Dim oEnum As IUnknown

Function IsObj()
    IsObj = useobj
End Function

Sub Construct(ref As Long, vtrefnumber As Long, Optional indirect As Boolean = False)
If useobj Then Exit Sub
Init = True

ObjectRef = ref
If indirect Then
vtref = -vtrefnumber
Else
vtref = vtrefnumber
End If
End Sub
Sub ConstructObj(obj As Object, vtrefnumber As Long)
Init = True
Set hardlink = obj
useobj = True
'ObjectRef = ref
vtref = vtrefnumber
End Sub
Public Property Get lastobj() As Object
Set lastobj = mLastObj
Set mLastObj = Nothing
End Property

Public Property Get Value() As Variant

Dim RETVAR As Variant
Dim check As Boolean
Set mLastObj = Nothing
If Init Then
If UseIndex Then
If useobj Then
 Value = ReadPropIndexObj(hardlink, vtref, mIndex)
Else
  Value = ReadPropIndex(ObjectRef, vtref, mIndex, check, RETVAR)
 If check Then
 Set mLastObj = RETVAR
 End If
 End If

Else


If useobj Then
If ReadPropObj(hardlink, vtref, RETVAR) Then
If IsObject(RETVAR) Then
Set Value = RETVAR
Else
Value = RETVAR
End If
End If
Else
 If ReadProp(ObjectRef, vtref, RETVAR) Then
 If IsObject(RETVAR) Then
Set Value = RETVAR
Else
On Error Resume Next
If TypeOf RETVAR Is IUnknown Then Set Value = RETVAR: Exit Property
Err.Clear
Value = RETVAR
End If
 End If
 End If
 End If
Else
' ??????????

End If
End Property
Public Property Let Value(vNewValue As Variant)
If Init Then
If Not UseIndex Then
If useobj Then
WritePropObj hardlink, vtref, vNewValue
Else
WriteProp ObjectRef, vtref, vNewValue
End If
Else
If useobj Then
WritePropIndexObj hardlink, vtref, vNewValue, mIndex
Else
WritePropIndex ObjectRef, vtref, vNewValue, mIndex
End If
End If
End If

End Property
Public Property Let ValueStr(s As String)
Dim vNewValue
vNewValue = vbNullString
SwapString2Variant s, vNewValue
If Init Then
If Not UseIndex Then
If useobj Then
WritePropObj hardlink, vtref, vNewValue
Else
WriteProp ObjectRef, vtref, vNewValue
End If
Else
If useobj Then
WritePropIndexObj hardlink, vtref, vNewValue, mIndex
Else
WritePropIndex ObjectRef, vtref, vNewValue, mIndex
End If
End If
End If

End Property
Public Property Get isMethod() As Boolean
isMethod = IamMethod
End Property

Public Property Let isMethod(ByVal vNewValue As Boolean)
IamMethod = vNewValue
End Property

Private Sub Class_Terminate()
Set hardlink = Nothing
End Sub
Public Property Get Index() As Variant
Index = mIndex
End Property

Public Property Let Index(ByVal RHS As Variant)
mIndex = RHS
End Property

