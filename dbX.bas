Attribute VB_Name = "databaseX"
'This is the new version for ADO.
Option Explicit
Public ArrBase As Long
Dim AABB As Long
Dim conCollection As FastCollection
Dim Init As Boolean
'  to be changed User and UserPassword
Public JetPrefixUser As String
Public JetPostfixUser As String
Public JetPrefix As String
Public JetPostfix As String
'old Microsoft.Jet.OLEDB.4.0
' Microsoft.ACE.OLEDB.12.0
Public Const JetPrefixOld = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
Public Const JetPostfixOld = ";Jet OLEDB:Database Password=100101;"
Public Const JetPrefixHelp = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
Public Const JetPostfixHelp = ";Jet OLEDB:Database Password=100101;"
Public DBUser As String ' '= VbNullString ' "admin"  ' or ""
Public DBUserPassword   As String ''= VbNullString
Public extDBUser As String ' '= VbNullString ' "admin"  ' or ""
Public extDBUserPassword   As String ''= VbNullString
Public DBtype As String ' can be mdb or something else
Public Const DBtypeHelp = ".mdb" 'allways help has an mdb as type"
Const DBSecurityOFF = ";Persist Security Info=False"

Private Declare Function MoveFileW Lib "kernel32.dll" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long) As Long
Private Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Public Sub KillFile(sFilenName As String)
DeleteFileW StrPtr(sFilenName)
End Sub

Public Function MoveFile(pOldPath As String, pNewPath As String)

    MoveFileW StrPtr(pOldPath), StrPtr(pNewPath)
    
End Function
Public Function isdir(F$) As Boolean
On Error Resume Next
Dim mm As New recDir
Dim lookfirst As Boolean
Dim Pad$
If F$ = vbNullString Then Exit Function
If F$ = "." Then F$ = mcd
If InStr(F$, "\..") > 0 Or F$ = ".." Or Left$(F$, 3) = "..\" Then
If Right$(F$, 1) <> "\" Then
Pad$ = ExtractPath(F$ & "\", True, True)
Else
Pad$ = ExtractPath(F$, True, True)
End If
If Pad$ = vbNullString Then
If Right$(F$, 1) <> "\" Then
Pad$ = ExtractPath(mcd + F$ & "\", True)
Else
Pad$ = ExtractPath(mcd + F$, True)
End If
End If
lookfirst = mm.isdir(Pad$)
If lookfirst Then F$ = Pad$
Else
F$ = mylcasefILE(F$)
lookfirst = mm.isdir(F$)
If Not lookfirst Then

Pad$ = mcd + F$

lookfirst = mm.isdir(Pad$)
If lookfirst Then F$ = Pad$

End If
End If
isdir = lookfirst
End Function
Public Sub fHelp(bstack As basetask, d$, Optional Eng As Boolean = False)
Dim sql$, b$, p$, c$, gp$, r As Double, bb As Long, i As Long
Dim cd As String, doriginal$, monitor As Long
d$ = Replace(d$, " ", ChrW(160))
On Error GoTo E5
'ON ERROR GoTo 0
If Not Form4.Visible Then
monitor = FindFormSScreen(Form1)
Else
monitor = FindFormSScreen(Form4)
End If
If HelpLastWidth > ScrInfo(monitor).Width Then HelpLastWidth = -1
doriginal$ = d$
d$ = Replace(d$, "'", "")
If d$ <> "" Then If Right$(d$, 1) = "(" Then d$ = d$ + ")"
If d$ = vbNullString Or d$ = "F12" Then
d$ = vbNullString
If Right$(d$, 1) = "(" Then d$ = d$ + ")"
p$ = subHash.Show

While ISSTRINGA(p$, c$)
'IsLabelA "", c$, b$
b$ = GetName(GetStrUntil(" ", c$))

If Right$(b$, 1) = "(" Then b$ = b$ + ")"
If gp$ <> "" Then gp$ = b$ + ", " + gp$ Else gp$ = b$
Wend
If vH_title$ <> "" Then b$ = "<| " & vH_title$ & vbCrLf & vbCrLf Else b$ = vbNullString
If Eng Then
        sHelp "User Modules/Functions [F12]", b$ & gp$, (ScrInfo(monitor).Width - 1) * 3 / 5, (ScrInfo(monitor).Height - 1) * 4 / 7
Else
        sHelp "Τμήματα/Συναρτήσεις Χρήστη [F12]", b$ & gp$, (ScrInfo(monitor).Width - 1) * 3 / 5, (ScrInfo(monitor).Height - 1) * 4 / 7
End If
vHelp Not Form4.Visible
Exit Sub
ElseIf GetSub(d$, i) Then
GoTo conthere
ElseIf GetlocalSubExtra(d$, i) Or d$ = here$ Then
conthere:
If d$ = here$ Then i = bstack.OriginalCode
If vH_title$ <> "" Then
b$ = "<| " & vH_title$ & vbCrLf & vbCrLf
Else
If Eng Then
b$ = "<| " & "User Modules/Functions [F12]" & vbCrLf & vbCrLf
Else
b$ = "<| " & "Τμήματα/Συναρτήσεις Χρήστη [F12]" & vbCrLf & vbCrLf
End If
End If
If Right$(d$, 1) = ")" Then

If Eng Then c$ = "[Function]" Else c$ = "[Συνάρτηση]"
Else
If Eng Then c$ = "[Module]" Else c$ = "[Function]"
End If

Dim ss$
    ss$ = GetNextLine((SBcode(i)))
    If Left$(ss$, 10) = "'11001EDIT" Then
    
    ss$ = Mid$(SBcode(i), Len(ss$) + 3)
    Else
     ss$ = SBcode(i)
     End If
        sHelp d$, c$ + "  " & b$ & ss$, (ScrInfo(monitor).Width - 1) * 3 / 5, (ScrInfo(monitor).Height - 1) * 4 / 7
    
        vHelp Not Form4.Visible
Exit Sub
End If




JetPrefix = JetPrefixHelp
JetPostfix = JetPostfixHelp
DBUser = vbNullString
DBUserPassword = vbNullString

cd = App.path
AddDirSep cd

p$ = Chr(34)
c$ = ","
d$ = doriginal$
If Right$(d$, 2) = "()" Then d$ = Left$(d$, Len(d$) - 1)
If AscW(d$) < 128 Then
sql$ = "SELECT * FROM [COMMANDS] WHERE ENGLISH >= '" & UCase(d$) & "'"
Else
sql$ = "SELECT * FROM [COMMANDS] WHERE DESCRIPTION >= '" & myUcase(d$, True) & "'"
End If
b$ = mylcasefILE(cd & "help2000")
getrow bstack, p$ & b$ & p$ & c$ & p$ & sql$ & p$ & ",1," & p$ & p$ & c$ & p$ & p$, False, , , True
sql$ = p$ & b$ & p$ & c$ & p$ & "GROUP" & p$
If bstack.IsNumber(r) Then
If bstack.IsString(gp$) Then
If bstack.IsString(b$) Then
If bstack.IsString(p$) Then
If bstack.IsNumber(r) Then
getrow bstack, sql$ & "," & CStr(1) & "," & Chr(34) & "GROUPNUM" & Chr(34) & "," & Str$(r), False, , , True
If bstack.IsNumber(r) Then
If bstack.IsNumber(r) Then
If bstack.IsString(c$) Then
' nothing
Dim sec$
        If Right$(gp$, 1) = "(" Then gp$ = gp$ + ")": p$ = p$ + ")"
        
        If Eng Then
        sec$ = "Identifier: " + p$ + ", Gr: " + gp$ + vbCrLf
        gp$ = p$
        
        Else
        gp$ = gp$
        sec$ = "Αναγνωριστικό: " + gp$ + ", En: " + p$ + vbCrLf
        End If
        If vH_title$ <> "" Then
            If vH_title$ = gp$ And Form4.Visible = True Then GoTo E5
        End If
        bb = InStr(b$, "__<ENG>__")
        If bb > 0 Then
            If Eng Then
            c$ = "List [" & NLtrim$(Mid$(c$, InStr(c$, ",") + 1)) & "]"
                b$ = Mid$(b$, bb + 11)
            Else
            c$ = "Λίστα [" & Mid$(c$, 1, InStr(c$, ",") - 1) & "]"
                b$ = Left$(b$, bb - 1)
            End If
            Else
             c$ = "Λίστα [" & Mid$(c$, 1, InStr(c$, ",") - 1) & "], List [" & NLtrim$(Mid$(c$, InStr(c$, ",") + 1)) & "]"
        End If
        If vH_title$ <> "" Then b$ = "<| " & vH_title$ & vbCrLf & vbCrLf & b$ Else b$ = vbCrLf & b$
        
        sHelp gp$, sec$ + c$ & "  " & b$, (ScrInfo(monitor).Width - 1) * 3 / 5, (ScrInfo(monitor).Height - 1) * 4 / 7
    
        vHelp Not Form4.Visible
        End If
    
    End If
End If

End If
End If
End If
End If
End If
E5:
JetPrefix = JetPrefixUser
JetPostfix = JetPostfixUser
DBUser = extDBUser
DBUserPassword = extDBUserPassword
Err.clear
End Sub
Public Function inames(i As Long, Lang As Long) As String
If (i And &H3) <> 1 Then
Select Case Lang
Case 1

inames = "DESCENDING"
Case Else
inames = "ΦΘΙΝΟΥΣΑ"
End Select
Else
Select Case Lang
Case 1
inames = "ASCENDING"
Case Else
inames = "ΑΥΞΟΥΣΑ"
End Select

End If

End Function
Public Function fnames(i As Long, Lang As Long) As String
Select Case i
Case 1
    Select Case Lang
    Case 1
    fnames = "BOOLEAN"
    Case Else
     fnames = "ΛΟΓΙΚΟΣ"
    End Select
    Exit Function
Case 2
    Select Case Lang
    Case 1
    fnames = "BYTE"
    Case Else
     fnames = "ΨΗΦΙΟ"
    End Select
   Exit Function

Case 3
        Select Case Lang
    Case 1
    fnames = "INTEGER"
    Case Else
     fnames = "ΑΚΕΡΑΙΟΣ"
    End Select
   Exit Function
Case 4
        Select Case Lang
    Case 1
    fnames = "LONG"
    Case Else
     fnames = "ΜΑΚΡΥΣ"
    End Select
   Exit Function
 
Case 5
        Select Case Lang
    Case 1
    fnames = "CURRENCY"
    Case Else
     fnames = "ΛΟΓΙΣΤΙΚΟ"
    End Select
   Exit Function

Case 6
    Select Case Lang
    Case 1
    fnames = "SINGLE"
    Case Else
     fnames = "ΑΠΛΟΣ"
    End Select
   Exit Function

Case 7
    Select Case Lang
    Case 1
    fnames = "DOUBLE"
    Case Else
     fnames = "ΔΙΠΛΟΣ"
    End Select
   Exit Function
Case 8
    Select Case Lang
    Case 1
    fnames = "DATEFIELD"
    Case Else
     fnames = "ΗΜΕΡΟΜΗΝΙΑ"
    End Select
   Exit Function
Case 9 '.....................ole 205
    Select Case Lang
    Case 1
    fnames = "BINARY"
    Case Else
     fnames = "ΔΥΑΔΙΚΟ"
    End Select
   Exit Function
Case 10 '..........................................202
    Select Case Lang
    Case 1
    fnames = "TEXT"
    Case Else
     fnames = "ΚΕΙΜΕΝΟ"
    End Select
   Exit Function
Case 11 '...........205
    fnames = "OLE"
    Exit Function
Case 12 '...........................202
    Select Case Lang
    Case 1
    fnames = "MEMO"
    Case Else
     fnames = "ΥΠΟΜΝΗΜΑ"
    End Select
Case Else
fnames = "?"
End Select
End Function

Public Sub NewBase(bstackstr As basetask, r$)
Dim base As String, othersettings As String
If FastSymbol(r$, "1") Then
ArrBase = 1
Exit Sub
ElseIf FastSymbol(r$, "0") Then
ArrBase = 0
Exit Sub
End If
If Not IsStrExp(bstackstr, r$, base) Then Exit Sub ' make it to give error
If FastSymbol(r$, ",") Then
If Not IsStrExp(bstackstr, r$, othersettings) Then Exit Sub  ' make it to give error
End If
 On Error Resume Next
 If Left$(base, 1) = "(" Or JetPostfix = ";" Then Exit Sub ' we can't create in ODBC
If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
If ExtractType(base) = vbNullString Then base = base & ".mdb"

If CFname((base)) <> "" Then
 If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
' check to see if is our
RemoveOneConn base
If CheckMine(base) Then
KillFile base
Err.clear

Else
MyEr "Can 't delete the Base", "Δεν μπορώ να διαγράψω τη βάση"

Exit Sub
End If
End If

 CreateObject("ADOX.Catalog").Create (JetPrefix & base & JetPostfix & othersettings)  'create a new, empty *.mdb-File

End Sub

Public Sub TABLENAMES(base As String, bstackstr As basetask, r$, Lang As Long)
Dim tablename As String, scope As Long, cnt As Long, srl As Long, stac1 As New mStiva
Dim myBase  ' variant
scope = 1

If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, tablename) Then
scope = 2

End If
End If


    Dim vindx As Boolean

    On Error Resume Next
            If Left$(base, 1) = "(" Or JetPostfix = ";" Then
        'skip this
        Else
            If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
            If ExtractType(base) = vbNullString Then base = base & ".mdb"
            If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
        End If
    If True Then
        On Error Resume Next
        If Not getone(base, myBase) Then
            Set myBase = CreateObject("ADODB.Connection")
            If DriveType(Left$(base, 3)) = "Cd-Rom" Then
                srl = DriveSerial(Left$(base, 3))
                If srl = 0 And Not GetDosPath(base) = vbNullString Then
                    If Lang = 0 Then
                        If Not ask("Βάλε το CD/Δισκέτα με το αρχείο " & ExtractName(base)) = vbCancel Then Exit Sub
                    Else
                        If Not ask("Put CD/Disk with file " & ExtractName(base)) = vbCancel Then Exit Sub
                    End If
                End If
                If myBase = vbNullString Then
                    If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                        myBase.Open JetPrefix & JetPostfix
                        If Err.Number Then
                        MyEr Err.Description, Err.Description
                        Exit Sub
                        End If
                    Else
                        myBase.Open JetPrefix & GetDosPath(base) & ";Mode=Share Deny Write" & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF      'open the Connection
                    End If
                End If
                If Err.Number > 0 Then
                    Do While srl <> DriveSerial(Left$(base, 3))
                        If Lang = 0 Then
                            If ask("Βάλε το CD/Δισκέτα με αριθμό σειράς " & CStr(srl) & " στον οδηγό " & Left$(base, 1)) = vbCancel Then Exit Do
                        Else
                            If ask("Put CD/Disk with serial number " & CStr(srl) & " in drive " & Left$(base, 1)) = vbCancel Then Exit Do
                        End If
                    Loop
                    If srl = DriveSerial(Left$(base, 3)) Then
                        Err.clear
                        If myBase = vbNullString Then myBase.Open JetPrefix & GetDosPath(base) & ";Mode=Share Deny Write" & JetPostfix & "User Id=" & DBUser & ";Password=" & DBSecurityOFF       'open the Connection
                    End If
                End If
            Else
                If myBase = vbNullString Then
                ' check if we have ODBC
                    If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                        myBase.Open JetPrefix & JetPostfix
                        If Err.Number Then
                            MyEr Err.Description, Err.Description
                            Exit Sub
                        End If
                    Else
                        Err.clear
                        myBase.Open JetPrefix & GetDosPath(base) & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
                        If Err.Number = -2147467259 Then
                           Err.clear
                           myBase.Open JetPrefixOld & GetDosPath(base) & JetPostfixOld & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
                           If Err.Number = 0 Then
                               JetPrefix = JetPrefixOld
                               JetPostfix = JetPostfixOld
                           Else
                               MyEr "Maybe Need Jet 4.0 library", "Μαλλον χρειάζεται η Jet 4.0 βιβλιοθήκη ρουτινών"
                           End If
                        End If
                    End If
                End If
        End If
        If Err.Number > 0 Then GoTo g102
        PushOne base, myBase
    End If
  Dim cat, TBL, rs
     Dim i As Long, j As Long, k As Long, KB As Boolean
  
           Set rs = CreateObject("ADODB.Recordset")
        Set TBL = CreateObject("ADOX.TABLE")
           Set cat = CreateObject("ADOX.Catalog")
           Set cat.activeconnection = myBase
           If cat.activeconnection.errors.count > 0 Then
           MyEr "Can't connect to Base", "Δεν μπορώ να συνδεθώ με τη βάση"
           Exit Sub
           End If
        If cat.TABLES.count > 0 Then
        For Each TBL In cat.TABLES
        
        If TBL.Type = "TABLE" Then
        vindx = False
        KB = False
        If scope <> 2 Then
        
        cnt = cnt + 1
                            stac1.DataStr TBL.name
                       If TBL.indexes.count > 0 Then
                                         For j = 0 To TBL.indexes.count - 1
                                                   With TBL.indexes(j)
                                                   If (.unique = False) And (.indexnulls = 0) Then
                                                        KB = True
                                                  Exit For
             '
                                                       End If
                                                   End With
                                                Next j
                                              If KB Then
                    
                                                     stac1.DataVal CDbl(1)
                                                     
                                                Else
                                                    stac1.DataVal CDbl(0)
                                                End If
                                               
                                           
                                            Else
                                            stac1.DataVal CDbl(0)
                                        End If
         ElseIf tablename = TBL.name Then
         cnt = 1
                     rs.Open "Select * From [" & TBL.name & "] ;", myBase, 3, 4 'adOpenStatic, adLockBatchOptimistic
                                         stac1.Flush
                                        stac1.DataVal CDbl(rs.fields.count)
                                        If TBL.indexes.count > 0 Then
                                         For j = 0 To TBL.indexes.count - 1
                                                   With TBL.indexes(j)
                                                   If (.unique = False) And (.indexnulls = 0) Then
                                                   vindx = True
                                                   Exit For
                                                       End If
                                                   End With
                                                Next j
                                                If vindx Then
                                                
                                                     stac1.DataVal CDbl(1)
                                                Else
                                                    stac1.DataVal CDbl(0)
                                                End If
                                            Else
                                            stac1.DataVal CDbl(0)
                                        End If
                     For i = 0 To rs.fields.count - 1
                     With rs.fields(i)
                             stac1.DataStr .name
                             If .Type = 203 And .DEFINEDSIZE >= 536870910# Then
                             
                                         If Lang = 1 Then
                                        stac1.DataStr "MEMO"
                                        Else
                                        stac1.DataStr "ΥΠΟΜΝΗΜΑ"
                                        End If
                                        
                                        stac1.DataVal CDbl(0)
                            
                             ElseIf .Type = 205 Then
                                       
                                            stac1.DataStr "OLE"
                                       
                                       
                                            stac1.DataVal CDbl(0)
                                     ElseIf .Type = 202 And .DEFINEDSIZE <> 536870910# Then
                                            If Lang = 1 Then
                                            stac1.DataStr "TEXT"
                                            Else
                                            stac1.DataStr "ΚΕΙΜΕΝΟ"
                                            End If
                                            stac1.DataVal CDbl(.DEFINEDSIZE)
                                    
                             Else
                                        stac1.DataStr ftype(.Type, Lang)
                                        stac1.DataVal CDbl(.DEFINEDSIZE)
                             
                             End If
                     End With
                     Next i
                     rs.Close
                     If vindx Then
                    If TBL.indexes.count > 0 Then
                             For j = 0 To TBL.indexes.count - 1
                          With TBL.indexes(j)
                          If (.unique = False) And (.indexnulls = 0) Then
                          stac1.DataVal CDbl(.Columns.count)
                          For k = 0 To .Columns.count - 1
                            stac1.DataStr .Columns(k).name
                             stac1.DataStr inames(.Columns(k).sortorder, Lang)
                          Next k
                             Exit For
                             
                             End If
                          End With
                       Next j
                    End If
                     End If
             End If
             End If
            
                                     
                         
               Next TBL
               Set TBL = Nothing
    End If
    If scope = 1 Then
    stac1.PushVal CDbl(cnt)
    Else
    If cnt = 0 Then
     MyEr "No such TABLE in DATABASE", "Δεν υπάρχει τέτοιο αρχείο στη βάση δεδομένων"
    End If
    End If
     bstackstr.soros.MergeTop stac1
     Else
     RemoveOneConn myBase
     MyEr "No such DATABASE", "Δεν υπάρχει τέτοια βάση δεδομένων"
    End If
g102:
End Sub

Public Sub append_table(bstackstr As basetask, base As String, r$, ED As Boolean, Optional Lang As Long = -1)
Dim table$, i&, par$, ok As Boolean, t As Double, j&
Dim gindex As Long
ok = False

If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, table$) Then
ok = True
End If
End If
If Lang <> -1 Then If IsLabelSymbolNew(r$, "ΣΤΟ", "TO", Lang) Then If IsExp(bstackstr, r$, t) Then gindex = CLng(t) Else SyntaxError
Dim Id$
  If InStr(UCase(Trim$(table$)) + " ", "SELECT") = 1 Then
Id$ = table$
Else
Id$ = "SELECT * FROM [" + table$ + "]"
End If


If Not ok Then Exit Sub


If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this
Else
    If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = vbNullString Then base = base & ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
End If
          On Error Resume Next
          Dim myBase
          
               If Not getone(base, myBase) Then
           
              Set myBase = CreateObject("ADODB.Connection")
                If DriveType(Left$(base, 3)) = "Cd-Rom" Then
                ' we can do NOTHING...
                    MyEr "Can't update base to a CD-ROM", "Δεν μπορώ να γράψω στη βάση δεδομένων σε CD-ROM"
                    Exit Sub
                Else
                If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                    myBase.Open JetPrefix & JetPostfix
                    If Err.Number Then
                        MyEr Err.Description, Err.Description
                        Exit Sub
                    End If
                Else
                        Err.clear
                        myBase.Open JetPrefix & GetDosPath(base) & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
                        If Err.Number = -2147467259 Then
                           Err.clear
                           myBase.Open JetPrefixOld & GetDosPath(base) & JetPostfixOld & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
                           If Err.Number = 0 Then
                               JetPrefix = JetPrefixOld
                               JetPostfix = JetPostfixOld
                           Else
                               MyEr "Maybe Need Jet 4.0 library", "Μαλλον χρειάζεται η Jet 4.0 βιβλιοθήκη ρουτινών"
                           End If
                        End If
                    End If
                End If
                PushOne base, myBase
            End If
           Err.clear
         
         '  If Err.Number > 0 Then GoTo thh
           
           
         '  Set rec = myBase.OpenRecordset(table$, dbOpenDynaset)
          Dim rec, LL$
          
           Set rec = CreateObject("ADODB.Recordset")
            Err.clear
           rec.Open Id$, myBase, 3, 4 'adOpenStatic, adLockBatchOptimistic

 If Err.Number <> 0 Then
LL$ = myBase ' AS A STRING
Set myBase = Nothing
RemoveOneConn base
 Set myBase = CreateObject("ADODB.Connection")
 myBase.Open = LL$
 PushOne base, myBase
 Err.clear
rec.Open Id$, myBase, 3, 4
If Err.Number Then
MyEr Err.Description & " " & Id$, Err.Description & " " & Id$
Exit Sub
End If
End If
   
   
If ED Then
If gindex > 0 Then
Err.clear
    rec.MoveLast
    rec.MoveFirst
    rec.AbsolutePosition = gindex '  - 1
    If Err.Number <> 0 Then
    MyEr "Wrong index for table " & table$, "Λάθος δείκτης για αρχείο " & table$
    End If
Else
    rec.MoveLast
End If
' rec.Edit  no need for undo
Else
rec.AddNew
End If
i& = 0
While FastSymbol(r$, ",")
If ED Then
    While FastSymbol(r$, ",")
    i& = i& + 1
    Wend
End If
If IsStrExp(bstackstr, r$, par$) Then
    rec.fields(i&) = par$
ElseIf IsExp(bstackstr, r$, t) Then

    rec.fields(i&) = CStr(t)   '??? convert to a standard format
End If

i& = i& + 1
Wend
Err.clear
rec.UpdateBatch  ' update be an updatebatch
If Err.Number > 0 Then
MyEr "Can't append " & Err.Description, "Αδυναμία προσθήκης:" & Err.Description
End If

End Sub
Public Sub getrow(bstackstr As basetask, r$, Optional ERL As Boolean = True, Optional Search$ = " = ", Optional Lang As Long = 0, Optional IamHelpFile As Boolean = False)

Dim base As String, table$, from As Long, first$, Second$, ok As Boolean, fr As Double, stac1$, p As Double, i&
ok = False
If IsStrExp(bstackstr, r$, base) Then
If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, table$) Then
If FastSymbol(r$, ",") Then
If IsExp(bstackstr, r$, fr) Then
from = CLng(fr)
If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, first$) Then
If FastSymbol(r$, ",") Then
If Search$ = vbNullString Then
    If IsStrExp(bstackstr, r$, Search$) Then
    Search$ = " " & Search$ & " "
        If FastSymbol(r$, ",") Then
                If IsExp(bstackstr, r$, p) Then
                Second$ = Search$ & Str$(p)
                ok = True
            ElseIf IsStrExp(bstackstr, r$, Second$) Then
            If InStr(Second$, "'") > 0 Then
                Second$ = Search$ & Chr(34) & Second$ & Chr(34)
            Else
                Second$ = Search$ & "'" & Second$ & "'"
                End If
                ok = True
            End If
        End If
 
        End If
    Else
     If IsExp(bstackstr, r$, p) Then
            Second$ = Search$ & Str$(p)
            ok = True
            ElseIf IsStrExp(bstackstr, r$, Second$) Then
                      If InStr(Second$, "'") > 0 Then
                Second$ = Search$ & Chr(34) & Second$ & Chr(34)
            Else
                Second$ = Search$ & "'" & Second$ & "'"
                End If
            ok = True
        End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
'Dim wrkDefault As Workspace,
Dim ii As Long
Dim myBase  ' as variant


Dim rec   '  as variant  too  - As Recordset
Dim srl As Long
On Error Resume Next
' new addition to handle ODBC
' base=""
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this

Else
If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
If ExtractType(base) = vbNullString Then base = base & ".mdb"
If Not IamHelpFile Then If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
End If

g05:
Err.clear
   On Error Resume Next
Dim Id$
   
      If first$ = vbNullString Then
If InStr(UCase(Trim$(table$)) + " ", "SELECT") = 1 Then
Id$ = table$
Else
Id$ = "SELECT * FROM [" + table$ + "]"
  End If
   Else
Id$ = "SELECT * FROM [" & table$ & "] WHERE [" & first$ & "] " & Second$
 End If

   If Not getone(base, myBase) Then
   
      Set myBase = CreateObject("ADODB.Connection")
   
      
    If DriveType(Left$(base, 3)) = "Cd-Rom" Then
        srl = DriveSerial(Left$(base, 3))
        If srl = 0 And Not GetDosPath(base) = vbNullString Then
                If Lang = 0 Then
                    If Not ask("Βάλε το CD/Δισκέτα με το αρχείο " & ExtractName(base)) = vbCancel Then Exit Sub
                Else
                    If Not ask("Put CD/Disk with file " & ExtractName(base)) = vbCancel Then Exit Sub
                End If
         End If

 
 '  If mybase = VbNullString Then ' mybase.Mode = adShareDenyWrite
   If myBase = vbNullString Then myBase.Open JetPrefix & GetDosPath(base) & ";Mode=Share Deny Write" & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection

            If Err.Number > 0 Then
            
            Do While srl <> DriveSerial(Left$(base, 3))
                If Lang = 0 Then
                If ask("Βάλε το CD/Δισκέτα με αριθμό σειράς " & CStr(srl) & " στον οδηγό " & Left$(base, 1)) = vbCancel Then Exit Do
                Else
                If ask("Put CD/Disk with serial number " & CStr(srl) & " in drive " & Left$(base, 1)) = vbCancel Then Exit Do
                End If
            Loop
            If srl = DriveSerial(Left$(base, 3)) Then
            Err.clear
        If myBase = vbNullString Then myBase.Open JetPrefix & GetDosPath(base) & ";Mode=Share Deny Write" & JetPostfix & "User Id=" & DBUser & ";Password=" & DBSecurityOFF      'open the Connection
        
            End If
        
        End If
    Else
'     myBase.Open JetPrefix & """" & GetDosPath(BASE) & """" & ";Jet OLEDB:Database Password=100101;User Id=" & DBUser  & ";Password=" & DBUserPassword & ";" &  DBSecurityOFF  'open the Connection
 If myBase = vbNullString Then
 If Left$(base, 1) = "(" Or JetPostfix = ";" Then
 myBase.Open JetPrefix & JetPostfix
 If Err.Number Then
 MyEr Err.Description, Err.Description
 Exit Sub
 End If
 Else
        Err.clear
        myBase.Open JetPrefix & GetDosPath(base) & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
        If Err.Number = -2147467259 Then
           Err.clear
           myBase.Open JetPrefixOld & GetDosPath(base) & JetPostfixOld & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
           If Err.Number = 0 Then
               JetPrefix = JetPrefixOld
               JetPostfix = JetPostfixOld
           Else
               MyEr "Maybe Need Jet 4.0 library", "Μαλλον χρειάζεται η Jet 4.0 βιβλιοθήκη ρουτινών"
           End If
        End If
 End If
 End If


    End If

   If Err.Number > 0 Then GoTo g10
   
      PushOne base, myBase
      
      End If

Dim LL$
   Set rec = CreateObject("ADODB.Recordset")
 Err.clear
 
  rec.Open Id$, myBase, 3, 4
If Err.Number <> 0 Then
LL$ = myBase ' AS A STRING
Set myBase = Nothing
RemoveOneConn base
 Set myBase = CreateObject("ADODB.Connection")
 myBase.Open = LL$
 PushOne base, myBase
 Err.clear
rec.Open Id$, myBase, 3, 4
If Err.Number Then
MyEr Err.Description & " " & Id$, Err.Description & " " & Id$
Exit Sub
End If
End If

   

   
  If rec.EOF Then
   ' stack$(BASESTACK) = " 0" & stack$(BASESTACK)
   bstackstr.soros.PushVal CDbl(0)
   rec.Close
  myBase.Close
    
    Exit Sub
  End If
  rec.MoveLast
  ii = rec.RecordCount

If ii <> 0 Then
If from >= 0 Then
  rec.MoveFirst
    If ii >= from Then
  rec.Move from - 1
  End If
End If
    For i& = rec.fields.count - 1 To 0 Step -1

   Select Case rec.fields(i&).Type
Case 1, 2, 3, 4, 5, 6

 If IsNull(rec.fields(i&)) Then
        bstackstr.soros.PushUndefine          '.PushStr "0"
    Else
        bstackstr.soros.PushVal CDbl(rec.fields(i&))
    
End If
Case 7
If IsNull(rec.fields(i&)) Then
    
     bstackstr.soros.PushStr ""
 Else
  
   bstackstr.soros.PushStr CStr(CDate(rec.fields(i&)))
  End If


Case 130, 8, 203, 202
If IsNull(rec.fields(i&)) Then
    
     bstackstr.soros.PushStr ""
 Else
  
   bstackstr.soros.PushStr CStr(rec.fields(i&))
  End If
Case 11, 12 ' this is the binary field so we can save unicode there
   Case Else
'
   bstackstr.soros.PushStr "?"
 End Select
   Next i&
   End If
   
   'stack$(BaseSTACK) = " " & Trim$(Str$(II)) + stack$(BaseSTACK)
   bstackstr.soros.PushVal CDbl(ii)


Exit Sub
g10:
If ERL Then
If Lang = 0 Then
If ask("Το ερώτημα SQL δεν μπορεί να ολοκληρωθεί" & vbCrLf & table$, True) = vbRetry Then GoTo g05
Else
If ask("SQL can't complete" & vbCrLf & table$) = vbRetry Then GoTo g05
End If
Err.clear
MyErMacro r$, "Can't read a database table :" & table$, "Δεν μπορώ να διαβάσω πίνακα :" & table$
End If
On Error Resume Next


End Sub

Public Sub GetNames(bstackstr As basetask, r$, bv As Object, Lang)
Dim base As String, table$, from As Long, many As Long, ok As Boolean, fr As Double, stac1$, i&
ok = False
If IsStrExp(bstackstr, r$, base) Then
If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, table$) Then
If FastSymbol(r$, ",") Then
If IsExp(bstackstr, r$, fr) Then
from = CLng(fr)
If FastSymbol(r$, ",") Then
If IsExp(bstackstr, r$, fr) Then
many = CLng(fr)

ok = True
End If
End If
End If
End If
End If
End If
End If
Dim ii As Long
Dim myBase ' variant
Dim rec
Dim srl As Long
On Error Resume Next
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this
Else
    If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = vbNullString Then base = base & ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
End If
Dim Id$
  If InStr(UCase(Trim$(table$)) + " ", "SELECT") = 1 Then
Id$ = table$
Else
Id$ = "SELECT * FROM [" + table$ + "]"
End If

     If Not getone(base, myBase) Then
   
      Set myBase = CreateObject("ADODB.Connection")
   
   
   If DriveType(Left$(base, 3)) = "Cd-Rom" Then
       srl = DriveSerial(Left$(base, 3))
    If srl = 0 And Not GetDosPath(base) = vbNullString Then
    
       If Lang = 0 Then
    If Not ask("Βάλε το CD/Δισκέτα με το αρχείο " & ExtractName(base)) = vbCancel Then Exit Sub
    Else
      If Not ask("Put CD/Disk with file " & ExtractName(base)) = vbCancel Then Exit Sub
    End If
     End If

     myBase.Open JetPrefix & GetDosPath(base) & ";Mode=Share Deny Write" & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF    'open the Connection

               If Err.Number > 0 Then
        
            Do While srl <> DriveSerial(Left$(base, 3))
            If Lang = 0 Then
            If ask("Βάλε το CD/Δισκέτα με αριθμό σειράς " & CStr(srl) & " στον οδηγό " & Left$(base, 1)) = vbCancel Then Exit Do
            Else
            If ask("Put CD/Disk with serial number " & CStr(srl) & " in drive " & Left$(base, 1)) = vbCancel Then Exit Do
            End If
            Loop
            If srl = DriveSerial(Left$(base, 3)) Then
            Err.clear
   myBase.Open JetPrefix & GetDosPath(base) & ";Mode=Share Deny Write" & JetPostfix & "User Id=" & DBUser & ";Password=" & DBSecurityOFF   'open the Connection
                
            End If
        
        End If
   Else
    If Left$(base, 1) = "(" Or JetPostfix = ";" Then
 myBase.Open JetPrefix & JetPostfix
 If Err.Number Then
 MyEr Err.Description, Err.Descnullription
 Exit Sub
 End If
 Else
        Err.clear
        myBase.Open JetPrefix & GetDosPath(base) & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
        If Err.Number = -2147467259 Then
           Err.clear
           myBase.Open JetPrefixOld & GetDosPath(base) & JetPostfixOld & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
           If Err.Number = 0 Then
               JetPrefix = JetPrefixOld
               JetPostfix = JetPostfixOld
           Else
               MyEr "Maybe Need Jet 4.0 library", "Μαλλον χρειάζεται η Jet 4.0 βιβλιοθήκη ρουτινών"
           End If
        End If
End If
End If
On Error GoTo g101
      PushOne base, myBase
      
      End If
 Dim LL$
   Set rec = CreateObject("ADODB.Recordset")
    Err.clear
     rec.Open Id$, myBase, 3, 4
      If Err.Number <> 0 Then
LL$ = myBase ' AS A STRING
Set myBase = Nothing
RemoveOneConn base
 Set myBase = CreateObject("ADODB.Connection")
 myBase.Open = LL$
 PushOne base, myBase
 Err.clear
rec.Open Id$, myBase, 3, 4
If Err.Number Then
MyEr Err.Description & " " & Id$, Err.Description & " " & Id$
Exit Sub
End If
End If


 ' DBEngine.Idle dbRefreshCache

  If rec.EOF Then
   ''''''''''''''''' stack$(BASESTACK) = " 0" & stack$(BASESTACK)
bstackstr.soros.PushVal CDbl(0)
  Exit Sub
 
'    wrkDefault.Close
  End If
  rec.MoveLast
  ii = rec.RecordCount

If ii <> 0 Then
If from >= 0 Then
  rec.MoveFirst
    If ii >= from Then
  rec.Move from - 1
  End If
End If
If many + from - 1 > ii Then many = ii - from + 1
bstackstr.soros.PushVal CDbl(ii)
''''''''''''''''' stack$(BASESTACK) = " " & Trim$(Str$(II)) + stack$(BASESTACK)

    For i& = 1 To many
    bv.additemFast CStr(rec.fields(0))   ' USING gList
    
    If i& < many Then rec.MoveNext
    Next
  End If
rec.Close
'myBase.Close

Exit Sub
g101:
MyErMacro r$, "Can't read a table from database", "Δεν μπορώ να διαβάσω ένα πίνακα βάσης δεδομένων"

'myBase.Close
End Sub
Public Sub CommExecAndTimeOut(bstackstr As basetask, r$)
Dim base As String, com2execute As String, comTimeOut As Double
Dim ok As Boolean
comTimeOut = 30
If IsStrExp(bstackstr, r$, base) Then
    If FastSymbol(r$, ",") Then
        If IsStrExp(bstackstr, r$, com2execute) Then
        ok = True
            If FastSymbol(r$, ",") Then
                If Not IsExp(bstackstr, r$, comTimeOut) Then
                ok = False
                End If
            End If
        End If
    End If
End If
If Not ok Then Exit Sub
On Error Resume Next
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this
Else
    If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = vbNullString Then base = base & ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
End If

Dim myBase
    
    On Error Resume Next
       If Not getone(base, myBase) Then
           
              Set myBase = CreateObject("ADODB.Connection")
                If DriveType(Left$(base, 3)) = "Cd-Rom" Then
                ' we can do NOTHING...
                    MyEr "Can't execute command in a CD-ROM", "Δεν μπορώ εκτελέσω εντολή στη βάση δεδομένων σε CD-ROM"
                    Exit Sub
                Else
                    If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                        myBase.Open JetPrefix & JetPostfix
                        If Err.Number Then
                        MyEr Err.Description, Err.Description
                        Exit Sub
                        End If
                    Else
                        Err.clear
                        myBase.Open JetPrefix & GetDosPath(base) & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
                        If Err.Number = -2147467259 Then
                           Err.clear
                           myBase.Open JetPrefixOld & GetDosPath(base) & JetPostfixOld & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
                           If Err.Number = 0 Then
                               JetPrefix = JetPrefixOld
                               JetPostfix = JetPostfixOld
                           Else
                               MyEr "Maybe Need Jet 4.0 library", "Μαλλον χρειάζεται η Jet 4.0 βιβλιοθήκη ρουτινών"
                           End If
                        End If
                    End If
                End If
                PushOne base, myBase
    End If
        Dim erdesc$, rs As Object
           Err.clear
           If comTimeOut >= 10 Then myBase.CommandTimeout = CLng(comTimeOut)
           If Err.Number > 0 Then Err.clear: myBase.errors.clear
           com2execute = Replace(com2execute, Chr(9), " ")
           com2execute = Replace(com2execute, vbCrLf, "")
           com2execute = Replace(com2execute, ";", vbCrLf)
           If InStr(com2execute, vbCrLf) > 0 Then
            Dim commands() As String, i As Long
            commands() = Split(com2execute, vbCrLf)
            For i = LBound(commands()) To UBound(commands())
            If Len(MyTrim(commands(i))) > 0 Then
            Set rs = myBase.Execute(commands(i))
            If Typename(rs) = "Recordset" Then
            If rs.fields.count > 0 Then
            
            End If
            End If
            If myBase.errors.count <> 0 Then
            
            Exit For
            End If
            End If
            Next i
           Else
            Set rs = myBase.Execute(com2execute)
            If Typename(rs) = "Recordset" Then
            If rs.fields.count > 0 Then
            bstackstr.soros.PushObj rs
            Set rs = Nothing
            End If
            End If
            End If
If myBase.errors.count <> 0 Then
For i = 0 To myBase.errors.count - 1
erdesc$ = erdesc$ + myBase.errors(i)
Next i
MyEr "Can't execute command:" + erdesc$, "Δεν μπορώ να εκτελέσω την εντολή:" + erdesc$
 myBase.errors.clear
End If

' we have response


End Sub



Public Sub MyOrder(bstackstr As basetask, r$)
Dim base As String, tablename As String, fs As String, i&, o As Double, ok As Boolean
ok = False
If IsStrExp(bstackstr, r$, base) Then
If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, tablename) Then
ok = True
End If
End If
End If

If Not ok Then Exit Sub
On Error Resume Next
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this
Else
    If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = vbNullString Then base = base & ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
End If
    
    Dim myBase
    
    On Error Resume Next
       If Not getone(base, myBase) Then
           
              Set myBase = CreateObject("ADODB.Connection")
                If DriveType(Left$(base, 3)) = "Cd-Rom" Then
                ' we can do NOTHING...
                    MyEr "Can't update base to a CD-ROM", "Δεν μπορώ να γράψω στη βάση δεδομένων σε CD-ROM"
                    Exit Sub
                Else
                    If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                        myBase.Open JetPrefix & JetPostfix
                        If Err.Number Then
                        MyEr Err.Description, Err.Description
                        Exit Sub
                        End If
                    Else
                        Err.clear
                        myBase.Open JetPrefix & GetDosPath(base) & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
                        If Err.Number = -2147467259 Then
                           Err.clear
                           myBase.Open JetPrefixOld & GetDosPath(base) & JetPostfixOld & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
                           If Err.Number = 0 Then
                               JetPrefix = JetPrefixOld
                               JetPostfix = JetPostfixOld
                           Else
                               MyEr "Maybe Need Jet 4.0 library", "Μαλλον χρειάζεται η Jet 4.0 βιβλιοθήκη ρουτινών"
                           End If
                        End If
                    End If
                 
                End If
                PushOne base, myBase
            End If
           Err.clear
           Dim LL$, mcat, pIndex, mtable
           Dim okntable As Boolean
          
            Err.clear
            Set mcat = CreateObject("ADOX.Catalog")
            mcat.activeconnection = myBase

            

        If Err.Number <> 0 Then
LL$ = myBase ' AS A STRING
Set myBase = Nothing
RemoveOneConn base
 Set myBase = CreateObject("ADODB.Connection")
 myBase.Open = LL$
 PushOne base, myBase
 Err.clear
            Set mcat = CreateObject("ADOX.Catalog")
            mcat.activeconnection = myBase
            

If Err.Number Then
MyEr Err.Description & " " & tablename, Err.Description & " " & tablename
Exit Sub
End If
End If
Err.clear
mcat.TABLES(tablename).indexes("ndx").Remove
Err.clear
mcat.TABLES(tablename).indexes.Refresh

   If mcat.TABLES.count > 0 Then
   okntable = True
        For Each mtable In mcat.TABLES
        If mtable.Type = "TABLE" Then
        If mtable.name = tablename Then
        okntable = False
        Exit For
        End If
        End If
        Next mtable
'        Set mtable = Nothing
        If okntable Then GoTo t111
Else
t111:
MyEr "No tables in Database " + ExtractNameOnly(base), "Δεν υπάρχουν αρχεία στη βάση δεδομένων " + ExtractNameOnly(base)
Exit Sub
End If
' now we have mtable from mybase
If mtable Is Nothing Then
Else
 mtable.indexes("ndx").Remove  ' remove the old index/
 End If
 Err.clear
 If mcat.activeconnection.errors.count > 0 Then
 mcat.activeconnection.errors.clear
 End If
 Err.clear
   Set pIndex = CreateObject("ADOX.Index")
    pIndex.name = "ndx"  ' standard
    pIndex.indexnulls = 0 ' standrard
  
        While FastSymbol(r$, ",")
        If IsStrExp(bstackstr, r$, fs) Then
        If FastSymbol(r$, ",") Then
        If IsExp(bstackstr, r$, o) Then
        
        pIndex.Columns.Append fs
        If o = 0 Then
        pIndex.Columns(fs).sortorder = CLng(1)
        Else
        pIndex.Columns(fs).sortorder = CLng(2)
        End If
        End If
        End If
                 
        End If
        Wend
        If pIndex.Columns.count > 0 Then
        mtable.indexes.Append pIndex
             If Err.Number Then
          '   mtable.Append pIndex
         MyEr Err.Description, Err.Description
         Exit Sub
        End If
mcat.TABLES.Append mtable
Err.clear
mcat.TABLES.Refresh
End If
    
End Sub
Public Sub NewTable(bstackstr As basetask, r$)
'BASE As String, tablename As String, ParamArray flds()
Dim base As String, tablename As String, fs As String, i&, n As Double, l As Double, ok As Boolean
ok = False
If IsStrExp(bstackstr, r$, base) Then
If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, tablename) Then
ok = True
End If
End If
End If

If Not ok Then Exit Sub
On Error Resume Next
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this
Else
    If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = vbNullString Then base = base & ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
End If
    Dim okndx As Boolean, okntable As Boolean, one_ok As Boolean
    ' Dim wrkDefault As Workspace
    Dim myBase ' As Database
    Err.clear
    On Error Resume Next
                   If Not getone(base, myBase) Then
           
              Set myBase = CreateObject("ADODB.Connection")
                If DriveType(Left$(base, 3)) = "Cd-Rom" Then
                ' we can do NOTHING...
                    MyEr "Can't update base to a CD-ROM", "Δεν μπορώ να γράψω στη βάση δεδομένων σε CD-ROM"
                    Exit Sub
                Else
                If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                    myBase.Open JetPrefix & JetPostfix
                    If Err.Number Then
                    MyEr Err.Description, Err.Description
                    Exit Sub
                    End If
                Else
                    Err.clear
                    myBase.Open JetPrefix & GetDosPath(base) & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
                    If Err.Number = -2147467259 Then
                       Err.clear
                       myBase.Open JetPrefixOld & GetDosPath(base) & JetPostfixOld & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
                       If Err.Number = 0 Then
                           JetPrefix = JetPrefixOld
                           JetPostfix = JetPostfixOld
                       Else
                           MyEr "Maybe Need Jet 4.0 library", "Μαλλον χρειάζεται η Jet 4.0 βιβλιοθήκη ρουτινών"
                       End If
                    End If
                End If
                End If
                PushOne base, myBase
            End If
           Err.clear

    On Error Resume Next
   okntable = True
Dim cat, mtable, LL$
  Set cat = CreateObject("ADOX.Catalog")
           Set cat.activeconnection = myBase


If Err.Number <> 0 Then
LL$ = myBase ' AS A STRING
Set myBase = Nothing
RemoveOneConn base
 Set myBase = CreateObject("ADODB.Connection")
 myBase.Open = LL$
 PushOne base, myBase
 Err.clear
 Set cat.activeconnection = myBase
If Err.Number Then
MyEr Err.Description & " " & mtable, Err.Description & " " & mtable
Exit Sub
End If
End If

    Set mtable = CreateObject("ADOX.TABLE")
         
' check if table exist

           If cat.TABLES.count > 0 Then
        For Each mtable In cat.TABLES
          If mtable.Type = "TABLE" Then
        If mtable.name = tablename Then
        okntable = False
        Exit For
        End If
        End If
        Next mtable
       If okntable Then
       Set mtable = CreateObject("ADOX.TABLE")      ' get a fresh one
        mtable.name = tablename
       End If
    
    
 With mtable.Columns

                Do While FastSymbol(r$, ",")
                
                        If IsStrExp(bstackstr, r$, fs) Then
                        one_ok = True
                                If FastSymbol(r$, ",") Then
                                        If IsExp(bstackstr, r$, n) Then
                                
                                            If FastSymbol(r$, ",") Then
                                                If IsExp(bstackstr, r$, l) Then
                                                If n = 8 Then n = 7: l = 0
                                                If n = 10 Then n = 202
                                                If n = 12 Then n = 203: l = 0
                                                    If l <> 0 Then
                                                
                                                     .Append fs, n, l
                                                    Else
                                                     .Append fs, n
                                           
                                                    End If
                                        
                                                End If
                                            End If
                                        End If
                        
                                End If
                
                        End If
                
                Loop
               
End With
        If okntable Then
        
        cat.TABLES.Append mtable
        If Err.Number Then
        If Err.Number = -2147217859 Then
        Err.clear
        Else
         MyEr Err.Description, Err.Description
         Exit Sub
        End If
        
        End If
        cat.TABLES.Refresh
        ElseIf Not one_ok Then
        cat.TABLES.Delete tablename
        cat.TABLES.Refresh
        End If
        
' may the objects find the creator...


End If



End Sub


Sub BaseCompact(bstackstr As basetask, r$)

Dim base As String, conn, BASE2 As String, realtype$
If Not IsStrExp(bstackstr, r$, base) Then
MissParam r$
Else
If FastSymbol(r$, ",") Then
If Not IsStrExp(bstackstr, r$, BASE2) Then
MissParam r$
Exit Sub
End If
End If
'only for mdb
If Left$(base, 1) = "(" Or JetPostfix = ";" Then Exit Sub ' we can't compact in ODBC use control panel

''If JetPrefix <> JetPrefixHelp Then Exit Sub
  On Error Resume Next
  
If ExtractPath(base) = vbNullString Then
base = mylcasefILE(mcd + base)
Else
  If Not CanKillFile(base) Then FilePathNotForUser: Exit Sub
End If
realtype$ = mylcasefILE(Trim$(ExtractType(base)))
If realtype$ <> "" Then
    base = ExtractPath(base, True) + ExtractNameOnly(base)
    If BASE2 = vbNullString Then BASE2 = strTemp & LTrim$(Str(Timer)) & "_0." + realtype$ Else BASE2 = ExtractPath(BASE2) + LTrim$(Str(Timer)) + "_0." + realtype$
    Set conn = CreateObject("JRO.JetEngine")
    base = base & "." + realtype$

   conn.CompactDatabase JetPrefix & base & JetPostfixUser, _
                                GetStrUntil(";", "" + JetPrefix) & _
                                GetStrUntil(":", "" + JetPostfix) & ":Engine Type=5;" & _
                                "Data Source=" & BASE2 & JetPostfixUser
                                

    
    If Err.Number = 0 Then
    If ExtractPath(base) <> ExtractPath(BASE2) Then
       KillFile base
       Sleep 50
        If Err.Number = 0 Then
            MoveFile BASE2, base
            Sleep 50

        Else
            If GetDosPath(BASE2) <> "" Then KillFile BASE2
        End If
    
    Else
        KillFile base
        MoveFile BASE2, base
            Sleep 50
    
    End If
       
    
    
    
    Else
      
      
 
      MyErMacro r$, "Can't compact databese " & ExtractName(base) & "." & " use a back up", "Πρόβλημα με την βάση " & ExtractName(base) & ".mdb χρησιμοποίησε ένα σωσμένο αρχείο"
      End If
      Err.clear
    End If
End If
End Sub

Public Function DELfields(bstackstr As basetask, r$) As Boolean
Dim base$, table$, first$, Second$, ok As Boolean, p As Double
ok = False
If IsExp(bstackstr, r$, p) Then
If bstackstr.lastobj Is Nothing Then
MyEr "Expected Inventory", "Περίμενα Κατάσταση"
Exit Function
End If
If Not TypeOf bstackstr.lastobj Is mHandler Then
MyEr "Expected Inventory", "Περίμενα Κατάσταση"
Exit Function
ElseIf Not bstackstr.lastobj.t1 = 1 Then
MyEr "Expected Inventory", "Περίμενα Κατάσταση"
Exit Function
End If
Dim aa As FastCollection
Set aa = bstackstr.lastobj.objref
If aa.StructLen > 0 Then
MyEr "Structure members are ReadOnly", "Τα μέλη της δομής είναι μόνο για ανάγνωση"
Exit Function
End If
Set bstackstr.lastobj = Nothing
Do While FastSymbol(r$, ",")
ok = False
If IsExp(bstackstr, r$, p) Then
aa.Remove p
If Not aa.Done Then MyEr "Key not exist", "Δεν υπάρχει τέτοιο κλειδί": Exit Do
ok = True
ElseIf IsStrExp(bstackstr, r$, first$) Then
aa.Remove first$
If Not aa.Done Then MyEr "Key not exist", "Δεν υπάρχει τέτοιο κλειδί": Exit Do
ok = True
Else
    Exit Do
End If
Loop
DELfields = ok
Set aa = Nothing
Exit Function

ElseIf IsStrExp(bstackstr, r$, base$) Then
If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, table$) Then
If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, first$) Then
If FastSymbol(r$, ",") Then
If IsStrExp(bstackstr, r$, Second$) Then
ok = True

           If InStr(Second$, "'") > 0 Then
                Second$ = Chr(34) & Second$ & Chr(34)
            Else
                Second$ = "'" & Second$ & "'"
                End If
ElseIf IsExp(bstackstr, r$, p) Then
ok = True
Second$ = Trim$(Str$(p))
Else
MissParam r$
End If
Else
MissParam r$

End If
Else
MissParam r$

End If
Else
MissParam r$

End If
Else
MissParam r$
End If
Else
On Error Resume Next
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this we can 't killfile the base for odbc
Else
    If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = vbNullString Then base = base & ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: DELfields = False: Exit Function
    If CheckMine(base) Then KillFile base: DELfields = True: Exit Function
    
End If

End If
Else
MissParam r$
End If
If Not ok Then DELfields = False: Exit Function
On Error Resume Next
If Left$(base, 1) = "(" Or JetPostfix = ";" Then
'skip this
Else
    If ExtractPath(base) = vbNullString Then base = mylcasefILE(mcd + base)
    If ExtractType(base) = vbNullString Then base = base & ".mdb"
    If Not CanKillFile(base) Then FilePathNotForUser: DELfields = False: Exit Function
End If

Dim myBase
   On Error Resume Next
                   If Not getone(base, myBase) Then
           
              Set myBase = CreateObject("ADODB.Connection")
                If DriveType(Left$(base, 3)) = "Cd-Rom" Then
                ' we can do NOTHING...
                    MyEr "Can't update base to a CD-ROM", "Δεν μπορώ να γράψω στη βάση δεδομένων σε CD-ROM"
                    Exit Function
                Else
                    If Left$(base, 1) = "(" Or JetPostfix = ";" Then
                        myBase.Open JetPrefix & JetPostfix
                        If Err.Number Then
                        MyEr Err.Description, Err.Description
                        DELfields = False: Exit Function
                        End If
                    Else
                        Err.clear
                        myBase.Open JetPrefix & GetDosPath(base) & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
                        If Err.Number = -2147467259 Then
                           Err.clear
                           myBase.Open JetPrefixOld & GetDosPath(base) & JetPostfixOld & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF     'open the Connection
                           If Err.Number = 0 Then
                               JetPrefix = JetPrefixOld
                               JetPostfix = JetPostfixOld
                           Else
                               MyEr "Maybe Need Jet 4.0 library", "Μαλλον χρειάζεται η Jet 4.0 βιβλιοθήκη ρουτινών"
                           End If
                        End If
                    End If
                End If
                PushOne base, myBase
            End If
           Err.clear

    On Error Resume Next
Dim rec
   
   
   
   If first$ = vbNullString Then
   MyEr "Nothing to delete", "Τίποτα για να σβήσω"
   DELfields = False
   Exit Function
   Else
   myBase.errors.clear
   myBase.Execute "DELETE * FROM [" & table$ & "] WHERE " & first$ & " = " & Second$
   If myBase.errors.count > 0 Then
   MyEr "Can't delete " & table$, "Δεν μπορώ να διαγράψω"
   Else
    DELfields = True
   End If
   
   End If
   Set rec = Nothing

End Function

Function CheckMine(DBFileName) As Boolean
' M2000 changed to ADO...

Dim Cnn1
 Set Cnn1 = CreateObject("ADODB.Connection")

 On Error Resume Next
 Cnn1.Open JetPrefix & DBFileName & ";Jet OLEDB:Database Password=;User Id=" & DBUser & ";Password=" & DBUserPassword & ";"  ' &  DBSecurityOFF 'open the Connection
 If Err Then
 Err.clear
 Cnn1.Open JetPrefix & DBFileName & JetPostfix & "User Id=" & DBUser & ";Password=" & DBUserPassword & ";" & DBSecurityOFF    'open the Connection
 If Err Then
 Else
 CheckMine = True
 End If
 Cnn1.Close
 Else
 End If
End Function


Private Sub PushOne(conname As String, v As Variant)
On Error Resume Next
conCollection.AddKey conname, v
'Set v = conCollection(conname)
End Sub
Sub CloseAllConnections()
Dim v As Variant, bb As Boolean
On Error Resume Next
If Not Init Then Exit Sub
If conCollection.count > 0 Then
Dim i As Long
Err.clear
For i = conCollection.count - 1 To 0 Step -1
On Error Resume Next
conCollection.index = i
If conCollection.IsObj Then
With conCollection.ValueObj
bb = .connectionstring <> ""
If Err.Number = 0 Then
If .mode > 0 Then
If .state = 1 Then
   .Close
ElseIf .state = 2 Then
    .Close
ElseIf .state > 2 Then
Call .Cancel
.Close
End If
    
End If
End If
End With
End If
conCollection.Remove conCollection.KeyToString
Err.clear

Next i
Set conCollection = New FastCollection
End If
Err.clear
End Sub
Public Sub RemoveOneConn(conname)
On Error Resume Next
Dim vv
If conCollection Is Nothing Then Exit Sub
If Not conCollection.ExistKey(conname) Then Exit Sub

Set vv = conCollection.ValueObj
If Typename$(vv) = "Empty" Then
' old code here
Err.clear
    vv = conCollection(conname)
    If Not Err.Number <> 0 Then
        conCollection.Remove conname
        
Else
    Err.clear
    conname = mylcasefILE(conname)
    If ExtractPath(conname) = vbNullString Then conname = mylcasefILE(mcd + conname)
    If ExtractType(CVar(conname)) = vbNullString Then conname = mylcasefILE(conname + ".mdb")
    conCollection.Remove conname
    End If
    
    Exit Sub
End If
If Not Err.Number <> 0 Then
Err.clear
If vv.connectionstring <> "" Then
If Err.Number = 0 Then If vv.activeconnection <> "" Then vv.Close
Err.clear
End If
conCollection.Remove conname
Err.clear
End If
End Sub
Private Function getone(conname As String, this As Variant) As Boolean
On Error Resume Next
Dim v As Variant
InitMe

If conCollection.ExistKey(conname) Then
Set this = conCollection.ValueObj
End If
End Function

Private Sub InitMe()
If Init Then Exit Sub
Set conCollection = New FastCollection
Init = True
End Sub
Function ftype(ByVal a As Long, Lang As Long) As String
Select Case Lang
Case 0
Select Case a
    Case 0
ftype = "ΑΔΕΙΟ"
    Case 2
ftype = "ΨΗΦΙΟ"
    Case 3
ftype = "ΑΚΕΡΑΙΟΣ"
    Case 4
ftype = "ΑΠΛΟΣ"
    Case 5
ftype = "ΔΙΠΛΟΣ"
    Case 6
ftype = "ΛΟΓΙΣΤΙΚΟ"
    Case 7
ftype = "ΗΜΕΡΟΜΗΝΙΑ"
    Case 8
ftype = "BSTR"
    Case 9
ftype = "IDISPATCH"
    Case 10
ftype = "ERROR"
    Case 11
ftype = "ΛΟΓΙΚΟΣ"
    Case 12
ftype = "VARIANT"
    Case 13
ftype = "IUNKNOWN"
    Case 14
ftype = "DECIMAL"
    Case 16
ftype = "TINYINT"
    Case 17
ftype = "UNSIGNEDTINYINT"
    Case 18
ftype = "UNSIGNEDSMALLINT"
    Case 19
ftype = "UNSIGNEDINT"
    Case 20
ftype = "ΜΑΚΡΥΣ"   'LONG
    Case 21
ftype = "UNSIGNEDBIGINT"
    Case 64
ftype = "FILETIME"
    Case 72
ftype = "GUID"
    Case 128
ftype = "BINARY"
    Case 129
ftype = "CHAR"
    Case 130
ftype = "WCHAR"
    Case 131
ftype = "NUMERIC"
    Case 132
ftype = "USERDEFINED"
    Case 133
ftype = "DBDATE"
    Case 134
ftype = "DBTIME"
    Case 135
ftype = "ΗΜΕΡΟΜΗΝΙΑ" 'DBTIMESTAMP
    Case 136
ftype = "CHAPTER"
    Case 138
ftype = "PROPVARIANT"
    Case 139
ftype = "VARNUMERIC"
    Case 200
ftype = "VARCHAR"
    Case 201
ftype = "LONGVARCHAR"
    Case 202
ftype = "ΚΕΙΜΕΝΟ" '"VARWCHAR"
    Case 203
ftype = "LONGVARWCHAR"
    Case 204
ftype = "ΔΥΑΔΙΚΟ"  ' "VARBINARY"
    Case 205
ftype = "OLE" '"LONGVARBINARY"
    Case 8192
ftype = "ARRAY"
Case Else
ftype = "????"


End Select

Case Else  ' this is for 1
Select Case a
    Case 0
ftype = "EMPTY"
    Case 2
ftype = "BYTE"  'SMALLINT
    Case 3
ftype = "INTEGER"
    Case 4
ftype = "SINGLE"
    Case 5
ftype = "DOUBLE"
    Case 6
ftype = "CURRENCY"
    Case 7
ftype = "DATE"
    Case 8
ftype = "BSTR"
    Case 9
ftype = "IDISPATCH"
    Case 10
ftype = "ERROR"
    Case 11
ftype = "BOOLEAN"
    Case 12
ftype = "VARIANT"
    Case 13
ftype = "IUNKNOWN"
    Case 14
ftype = "DECIMAL"
    Case 16
ftype = "TINYINT"
    Case 17
ftype = "UNSIGNEDTINYINT"
    Case 18
ftype = "UNSIGNEDSMALLINT"
    Case 19
ftype = "UNSIGNEDINT"
    Case 20
ftype = "BIGINT"
    Case 21
ftype = "UNSIGNEDBIGINT"
    Case 64
ftype = "FILETIME"
    Case 72
ftype = "GUID"
    Case 128
ftype = "BINARY"
    Case 129
ftype = "CHAR"
    Case 130
ftype = "WCHAR"
    Case 131
ftype = "NUMERIC"
    Case 132
ftype = "USERDEFINED"
    Case 133
ftype = "DBDATE"
    Case 134
ftype = "DBTIME"
    Case 135
ftype = "DBTIMESTAMP"
    Case 136
ftype = "CHAPTER"
    Case 138
ftype = "PROPVARIANT"
    Case 139
ftype = "VARNUMERIC"
    Case 200
ftype = "VARCHAR"
    Case 201
ftype = "LONGVARCHAR"
    Case 202
ftype = "VARWCHAR"
    Case 203
ftype = "LONGVARWCHAR"
    Case 204
ftype = "VARBINARY"
    Case 205
ftype = "OLE"
    Case 8192
ftype = "ARRAY"


Case Else
ftype = "????"
End Select
End Select
End Function
Sub GeneralErrorReport(aBasBase As Variant)
Dim errorObject

 For Each errorObject In aBasBase.activeconnection.errors
 'Debug.Print "Description :"; errorObject.Description
 'Debug.Print "Number:"; Hex(errorObject.Number)
 Next
End Sub
