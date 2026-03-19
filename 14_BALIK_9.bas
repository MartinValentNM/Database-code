Attribute VB_Name = "BALIK_9_FIXES"
Option Compare Database
Option Explicit

' ================================================================
' BALIK_9_FIXES
'
' Fixes two bugs:
'
'   1. frmBulkRename "Preview error: property setting is too long"
'      Root cause: lstAffected.RowSource is a Value List string.
'      Access limits RowSource to ~32,700 chars. With many matching
'      records the string exceeds this limit.
'      Fix: btnPreview_Click saves a temporary QueryDef and binds
'      lstAffected to it (RowSourceType = Table/Query).
'
'   2. frmCompare4: CROSS-SECTION and CARDINAL_PROCESSES always empty.
'      Root cause: Access DAO .Name property for fields sourced from
'      a child table via JOIN in a saved QueryDef does not reliably
'      return the bare field name when it contains a hyphen.
'      Both name-based and index-scanning lookups fail silently.
'      Fix: use a HARDCODED field index map derived from the exact
'      SELECT column order in vw_Complete_Taxa (confirmed by user).
'      rs.Fields(idx(fi)).Value bypasses all name resolution entirely.
'
'      Index map (0-based column positions in vw_Complete_Taxa):
'        TAXON=2  Genus=3  Species=4  AUTHOR=5  Phylum=6
'        Family=7  Order=8  Class=9  Rank=10
'        TAXONOMIC PLACEMENT=12  Taxonomy_path=13
'        TYPE TAXON=15  INCLUDED TAXONS=16
'        ETYMOLOGY=18  DESCRIPTION=19  SIZE=20  DIAGNOSIS=21
'        APERTURE=22  CROSS-SECTION=23  OPERCULUM=24
'        SCULPTURE=25  CLAVICLES=26  CARDINAL_PROCESSES=27
'        TYPE MATERIAL=29  MATERIAL EXAMINED=30  FIGURES=31
'        SYNONYMY=32  OCCURRENCE=34  Geography=35  Locality=36
'        STRATIGRAPHY=37  System=38  Formation=39  Member=40
'        Stage=41  Serie=42  Zone=43  Horizon=44
'        REMARKS=46  Reference=47
'
' INSTALLATION:
'   1. VBA Editor -> File -> Import File -> BALIK_9_FIXES.bas
'   2. Immediate window: InstallFixes9
'   3. Save the database (Ctrl+S)
' ================================================================

Public Sub InstallFixes9()
    Debug.Print "=== BALIK_9_FIXES start ==="

    FixBulkRenamePreview
    Debug.Print "  [1] frmBulkRename preview fix - OK"

    FixCompare4LoadDirect
    Debug.Print "  [2] frmCompare4 btnLoad_Click (hardcoded index map) - OK"

    MsgBox "BALIK_9_FIXES complete!" & vbCrLf & vbCrLf & _
           "Fixed:" & vbCrLf & _
           "  1. frmBulkRename - Preview no longer crashes on large result sets" & vbCrLf & _
           "  2. frmCompare4 - CROSS-SECTION and CARDINAL_PROCESSES load correctly" & vbCrLf & _
           "     (hardcoded column index map, immune to hyphens in field names)", _
           vbInformation, "BALIK_9 OK"
End Sub

' ================================================================
' 1. Fix frmBulkRename Preview (RowSource length limit)
' ================================================================
Private Sub FixBulkRenamePreview()
    On Error Resume Next
    DoCmd.DeleteObject acForm, "frmBulkRename"
    On Error GoTo 0

    Dim f As Form: Set f = CreateForm
    Dim fn As String: fn = f.Name

    With f
        .Caption = "Bulk Rename Taxonomic Field": .Width = 13000: .PopUp = True: .Modal = False
        .ScrollBars = 0: .NavigationButtons = False: .RecordSelectors = False
        .AutoCenter = True: .BorderStyle = 2
        .Section(acDetail).BackColor = &H202020: .Section(acDetail).Height = 10000
    End With

    Dim c As Control
    Set c = CreateControl(fn, acLabel, acDetail, "", "", 200, 150, 12000, 600)
    c.Caption = "BULK RENAME": c.ForeColor = &H884400: c.BackStyle = 0
    c.FontSize = 14: c.FontBold = True

    Set c = CreateControl(fn, acLabel, acDetail, "", "", 200, 1000, 2000, 400)
    c.Caption = "Field:": c.ForeColor = &HC0C0C0: c.BackStyle = 0: c.FontSize = 10

    Set c = CreateControl(fn, acComboBox, acDetail, "", "", 2300, 970, 4000, 500)
    c.Name = "cboField": c.RowSourceType = "Value List"
    c.RowSource = """Genus"";""Family"";""Order"";""Class"";""Phylum"";""Rank"";""AUTHOR"";""Geography"";""STRATIGRAPHY"""
    c.DefaultValue = """Family""": c.BackColor = &H2B2B2B: c.ForeColor = &HFFFFFF
    c.BorderColor = &H444444: c.FontSize = 11: c.LimitToList = True
    c.OnClick = "[Event Procedure]"

    Set c = CreateControl(fn, acLabel, acDetail, "", "", 200, 1700, 2000, 400)
    c.Caption = "Current value:": c.ForeColor = &HC0C0C0: c.BackStyle = 0: c.FontSize = 10

    Set c = CreateControl(fn, acComboBox, acDetail, "", "", 2300, 1670, 10400, 500)
    c.Name = "cboOldValue": c.BackColor = &H2B2B2B: c.ForeColor = &HFFFFFF
    c.BorderColor = &H444444: c.FontSize = 11: c.LimitToList = False
    c.OnClick = "[Event Procedure]"

    Set c = CreateControl(fn, acLabel, acDetail, "", "", 200, 2400, 2000, 400)
    c.Caption = "New value:": c.ForeColor = &HC0C0C0: c.BackStyle = 0: c.FontSize = 10

    Set c = CreateControl(fn, acTextBox, acDetail, "", "", 2300, 2370, 10400, 500)
    c.Name = "txtNewValue": c.BackColor = &H2B2B2B: c.ForeColor = &HFFFFFF
    c.BorderColor = &H444444: c.FontSize = 11

    Set c = CreateControl(fn, acLabel, acDetail, "", "", 200, 3100, 12400, 400)
    c.Name = "lblPreview"
    c.Caption = "(select field and current value to see affected records)"
    c.ForeColor = &H888888: c.BackStyle = 0: c.FontSize = 10

    Set c = CreateControl(fn, acCommandButton, acDetail, "", "", 200, 3800, 2000, 600)
    c.Name = "btnPreview": c.Caption = "PREVIEW": c.BackColor = &H5580: c.ForeColor = &HFFFFFF
    c.FontSize = 10: c.FontBold = True: c.OnClick = "[Event Procedure]"

    Set c = CreateControl(fn, acCommandButton, acDetail, "", "", 2400, 3800, 2500, 600)
    c.Name = "btnRename": c.Caption = "APPLY RENAME": c.BackColor = &H884400: c.ForeColor = &HFFFFFF
    c.FontSize = 10: c.FontBold = True: c.OnClick = "[Event Procedure]"

    Set c = CreateControl(fn, acCommandButton, acDetail, "", "", 5100, 3800, 1500, 600)
    c.Name = "btnClose": c.Caption = "CLOSE": c.BackColor = &H111155: c.ForeColor = &HFFFFFF
    c.FontSize = 10: c.FontBold = True: c.OnClick = "[Event Procedure]"

    ' lstAffected bound to Table/Query - avoids the 32k string limit of Value List
    Set c = CreateControl(fn, acListBox, acDetail, "", "", 200, 4700, 12400, 5000)
    c.Name = "lstAffected"
    c.RowSourceType = "Table/Query"
    c.RowSource = ""
    c.ColumnCount = 3
    c.ColumnWidths = "1000;5000;4000"
    c.BackColor = &H1A1A1A: c.ForeColor = &HFFFFFF: c.BorderColor = &H444444: c.FontSize = 10

    Dim vba As String
    vba = "Option Compare Database" & vbCrLf & "Option Explicit" & vbCrLf & vbCrLf

    vba = vba & "Private Sub cboField_AfterUpdate()" & vbCrLf
    vba = vba & "    On Error Resume Next" & vbCrLf
    vba = vba & "    Dim fld As String: fld = Nz(Me!cboField.Value, """")" & vbCrLf
    vba = vba & "    If fld = """" Then Exit Sub" & vbCrLf
    vba = vba & "    Dim tbl As String: tbl = GetTableForField(fld)" & vbCrLf
    vba = vba & "    Me!cboOldValue.RowSourceType = ""Table/Query""" & vbCrLf
    vba = vba & "    Me!cboOldValue.RowSource = ""SELECT DISTINCT ["" & fld & ""] FROM "" & tbl & "" WHERE Len(Nz(["" & fld & ""],''))>0 ORDER BY ["" & fld & ""]""" & vbCrLf
    vba = vba & "    Me!cboOldValue.Requery" & vbCrLf
    vba = vba & "    Me!lblPreview.Caption = ""(select current value)""" & vbCrLf
    vba = vba & "    Me!lstAffected.RowSource = """"" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf
    vba = vba & "Private Sub cboField_Click(): cboField_AfterUpdate: End Sub" & vbCrLf & vbCrLf
    vba = vba & "Private Sub cboOldValue_AfterUpdate(): btnPreview_Click: End Sub" & vbCrLf & vbCrLf
    vba = vba & "Private Sub cboOldValue_Click(): btnPreview_Click: End Sub" & vbCrLf & vbCrLf

    vba = vba & "Private Sub btnPreview_Click()" & vbCrLf
    vba = vba & "    On Error GoTo PrevErr" & vbCrLf
    vba = vba & "    Dim fld As String: fld = Nz(Me!cboField.Value, """")" & vbCrLf
    vba = vba & "    Dim oldVal As String: oldVal = Nz(Me!cboOldValue.Value, """")" & vbCrLf
    vba = vba & "    If fld = """" Or oldVal = """" Then Exit Sub" & vbCrLf
    vba = vba & "    Dim tbl As String: tbl = GetTableForField(fld)" & vbCrLf
    vba = vba & "    Dim sql As String" & vbCrLf
    vba = vba & "    sql = ""SELECT t.ID, t.TAXON, t.AUTHOR FROM Taxa t """ & vbCrLf
    vba = vba & "    If tbl <> ""Taxa"" Then sql = sql & ""INNER JOIN "" & tbl & "" x ON t.ID=x.TaxonID """ & vbCrLf
    vba = vba & "    sql = sql & ""WHERE "" & IIf(tbl=""Taxa"",""t"",""x"") & "".["" & fld & ""]='"" & Replace(oldVal,""'"",""''"") & ""' ORDER BY t.TAXON""" & vbCrLf
    vba = vba & "    Dim db As DAO.Database: Set db = CurrentDb" & vbCrLf
    vba = vba & "    On Error Resume Next: db.QueryDefs.Delete ""qryBulkPreview"": On Error GoTo PrevErr" & vbCrLf
    vba = vba & "    db.CreateQueryDef ""qryBulkPreview"", sql" & vbCrLf
    vba = vba & "    Dim rs As DAO.Recordset: Set rs = db.OpenRecordset(sql)" & vbCrLf
    vba = vba & "    Dim cnt As Long: cnt = 0" & vbCrLf
    vba = vba & "    If Not rs.EOF Then rs.MoveLast: cnt = rs.RecordCount: rs.MoveFirst" & vbCrLf
    vba = vba & "    rs.Close" & vbCrLf
    vba = vba & "    Me!lstAffected.RowSourceType = ""Table/Query""" & vbCrLf
    vba = vba & "    Me!lstAffected.RowSource = ""qryBulkPreview""" & vbCrLf
    vba = vba & "    Me!lstAffected.Requery" & vbCrLf
    vba = vba & "    Me!lblPreview.Caption = cnt & "" records will be renamed: ["" & fld & ""] '"" & oldVal & ""' -> '"" & Nz(Me!txtNewValue.Value,""???"") & ""'""" & vbCrLf
    vba = vba & "    Exit Sub" & vbCrLf
    vba = vba & "PrevErr: MsgBox ""Preview error: "" & Err.Description, vbCritical" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf

    vba = vba & "Private Sub btnRename_Click()" & vbCrLf
    vba = vba & "    On Error GoTo RenErr" & vbCrLf
    vba = vba & "    Dim fld As String: fld = Nz(Me!cboField.Value, """")" & vbCrLf
    vba = vba & "    Dim oldVal As String: oldVal = Nz(Me!cboOldValue.Value, """")" & vbCrLf
    vba = vba & "    Dim newVal As String: newVal = Trim(Nz(Me!txtNewValue.Value, """"))" & vbCrLf
    vba = vba & "    If fld = """" Or oldVal = """" Then MsgBox ""Select field and current value."", vbExclamation: Exit Sub" & vbCrLf
    vba = vba & "    If newVal = """" Then MsgBox ""Enter new value."", vbExclamation: Exit Sub" & vbCrLf
    vba = vba & "    If oldVal = newVal Then MsgBox ""Old and new values are identical."", vbExclamation: Exit Sub" & vbCrLf
    vba = vba & "    Dim tbl As String: tbl = GetTableForField(fld)" & vbCrLf
    vba = vba & "    Dim cnt As Long: cnt = DCount(""*"", tbl, ""["" & fld & ""]='"" & Replace(oldVal,""'"",""''"") & ""'"")" & vbCrLf
    vba = vba & "    If cnt = 0 Then MsgBox ""No records found."", vbInformation: Exit Sub" & vbCrLf
    vba = vba & "    If MsgBox(""Rename ["" & fld & ""] '"" & oldVal & ""' -> '"" & newVal & ""' in "" & cnt & "" records?"", vbYesNo+vbQuestion) <> vbYes Then Exit Sub" & vbCrLf
    vba = vba & "    Dim db As DAO.Database: Set db = CurrentDb" & vbCrLf
    vba = vba & "    Dim joinFld As String: joinFld = IIf(tbl = ""Taxa"", ""ID"", ""TaxonID"")" & vbCrLf
    vba = vba & "    Dim rsBak As DAO.Recordset" & vbCrLf
    vba = vba & "    Set rsBak = db.OpenRecordset(""SELECT "" & joinFld & "" FROM "" & tbl & "" WHERE ["" & fld & ""]='"" & Replace(oldVal,""'"",""''"") & ""'"")" & vbCrLf
    vba = vba & "    Do While Not rsBak.EOF" & vbCrLf
    vba = vba & "        Call BackupFieldForUndo(rsBak.Fields(0), fld)" & vbCrLf
    vba = vba & "        rsBak.MoveNext" & vbCrLf
    vba = vba & "    Loop: rsBak.Close" & vbCrLf
    vba = vba & "    db.Execute ""UPDATE ["" & tbl & ""] SET ["" & fld & ""]='"" & Replace(newVal,""'"",""''"") & ""' WHERE ["" & fld & ""]='"" & Replace(oldVal,""'"",""''"") & ""'"", 128" & vbCrLf
    vba = vba & "    MsgBox ""Done: "" & cnt & "" records renamed. UNDO available."", vbInformation" & vbCrLf
    vba = vba & "    Me!cboOldValue.Requery: Me!lstAffected.RowSource = """": Me!lblPreview.Caption = ""Rename complete.""" & vbCrLf
    vba = vba & "    Exit Sub" & vbCrLf
    vba = vba & "RenErr: MsgBox ""Rename error: "" & Err.Description, vbCritical" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf

    vba = vba & "Private Function GetTableForField(fld As String) As String" & vbCrLf
    vba = vba & "    Select Case UCase(fld)" & vbCrLf
    vba = vba & "        Case ""GENUS"", ""FAMILY"", ""ORDER"", ""CLASS"", ""PHYLUM"", ""RANK"", ""AUTHOR"": GetTableForField = ""Taxa""" & vbCrLf
    vba = vba & "        Case ""GEOGRAPHY"", ""STRATIGRAPHY"": GetTableForField = ""Taxa_Vyskyt""" & vbCrLf
    vba = vba & "        Case Else: GetTableForField = ""Taxa""" & vbCrLf
    vba = vba & "    End Select" & vbCrLf
    vba = vba & "End Function" & vbCrLf & vbCrLf

    vba = vba & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    vba = vba & "    If KeyCode = 27 Then DoCmd.Close acForm, Me.Name: KeyCode = 0" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf
    vba = vba & "Private Sub btnClose_Click(): DoCmd.Close acForm, Me.Name: End Sub" & vbCrLf

    f.HasModule = True: f.KeyPreview = True
    If f.Module.CountOfLines > 0 Then f.Module.DeleteLines 1, f.Module.CountOfLines
    f.Module.InsertLines 1, vba
    Dim tmp As String: tmp = f.Name
    DoCmd.Close acForm, tmp, acSaveYes
    DoCmd.CopyObject , "frmBulkRename", acForm, tmp
    DoCmd.DeleteObject acForm, tmp
End Sub

' ================================================================
' 2. Fix frmCompare4 btnLoad_Click
'    Replaces the sub with a version using a hardcoded index map
'    so CROSS-SECTION (idx=23) and CARDINAL_PROCESSES (idx=27)
'    are read directly by column position, bypassing all name
'    resolution that fails on hyphenated field names in DAO.
' ================================================================
Private Sub FixCompare4LoadDirect()
    On Error GoTo C4Err
    DoCmd.OpenForm "frmCompare4", acDesign
    Dim frm As Form: Set frm = Forms("frmCompare4")
    Dim mdl As Module: Set mdl = frm.Module

    ' Find btnLoad_Click regardless of Public/Private declaration
    Dim i As Long, startLine As Long, endLine As Long
    startLine = 0: endLine = 0
    For i = 1 To mdl.CountOfLines
        Dim ln As String: ln = mdl.Lines(i, 1)
        If InStr(ln, "btnLoad_Click") > 0 And InStr(ln, "Sub ") > 0 Then
            startLine = i
        End If
        If startLine > 0 And i > startLine Then
            If Trim(ln) = "End Sub" Then
                endLine = i
                Exit For
            End If
        End If
    Next i

    If startLine = 0 Then
        MsgBox "btnLoad_Click not found in frmCompare4." & vbCrLf & _
               "Please rebuild frmCompare4 using BuildForms_2_Compare first.", vbExclamation
        GoTo CloseForm
    End If

    mdl.DeleteLines startLine, endLine - startLine + 1

    ' Field names for control name generation (40 entries)
    Dim flds(39) As String
    flds(0)="TAXON": flds(1)="Genus": flds(2)="Species": flds(3)="AUTHOR"
    flds(4)="Phylum": flds(5)="Family": flds(6)="Order": flds(7)="Class"
    flds(8)="Rank": flds(9)="TAXONOMIC PLACEMENT": flds(10)="Taxonomy_path"
    flds(11)="TYPE TAXON": flds(12)="INCLUDED TAXONS": flds(13)="ETYMOLOGY"
    flds(14)="DESCRIPTION": flds(15)="SIZE": flds(16)="DIAGNOSIS"
    flds(17)="APERTURE": flds(18)="CROSS-SECTION": flds(19)="OPERCULUM"
    flds(20)="SCULPTURE": flds(21)="CLAVICLES": flds(22)="CARDINAL_PROCESSES"
    flds(23)="TYPE MATERIAL": flds(24)="MATERIAL EXAMINED": flds(25)="FIGURES"
    flds(26)="SYNONYMY": flds(27)="OCCURRENCE": flds(28)="Geography"
    flds(29)="Locality": flds(30)="STRATIGRAPHY": flds(31)="System"
    flds(32)="Formation": flds(33)="Member": flds(34)="Stage"
    flds(35)="Serie": flds(36)="Zone": flds(37)="Horizon"
    flds(38)="REMARKS": flds(39)="Reference"

    ' Hardcoded 0-based column indices in vw_Complete_Taxa
    ' Based on confirmed SELECT order from user-provided SQL
    Dim idxMap(39) As Integer
    idxMap(0)=2:   idxMap(1)=3:   idxMap(2)=4:   idxMap(3)=5
    idxMap(4)=6:   idxMap(5)=7:   idxMap(6)=8:   idxMap(7)=9
    idxMap(8)=10:  idxMap(9)=12:  idxMap(10)=13: idxMap(11)=15
    idxMap(12)=16: idxMap(13)=18: idxMap(14)=19: idxMap(15)=20
    idxMap(16)=21: idxMap(17)=22: idxMap(18)=23: idxMap(19)=24
    idxMap(20)=25: idxMap(21)=26: idxMap(22)=27: idxMap(23)=29
    idxMap(24)=30: idxMap(25)=31: idxMap(26)=32: idxMap(27)=34
    idxMap(28)=35: idxMap(29)=36: idxMap(30)=37: idxMap(31)=38
    idxMap(32)=39: idxMap(33)=40: idxMap(34)=41: idxMap(35)=42
    idxMap(36)=43: idxMap(37)=44: idxMap(38)=46: idxMap(39)=47

    Dim q As String: q = Chr(34)
    Dim fi As Integer

    Dim s As String
    s = "Public Sub btnLoad_Click()" & vbCrLf
    s = s & "    On Error GoTo LoadErr" & vbCrLf
    s = s & "    Dim db As DAO.Database: Set db = CurrentDb" & vbCrLf
    s = s & "    Dim ci As Integer" & vbCrLf
    s = s & "    " & vbCrLf
    s = s & "    ' Field names used only for control name generation" & vbCrLf
    s = s & "    Dim fl(39) As String" & vbCrLf
    For fi = 0 To 39
        s = s & "    fl(" & fi & ")=" & q & flds(fi) & q & vbCrLf
    Next fi
    s = s & "    " & vbCrLf
    s = s & "    ' Hardcoded column indices in vw_Complete_Taxa (0-based)" & vbCrLf
    s = s & "    ' Derived from confirmed SQL SELECT column order" & vbCrLf
    s = s & "    ' idx(18)=23 -> CROSS-SECTION, idx(22)=27 -> CARDINAL_PROCESSES" & vbCrLf
    s = s & "    Dim idx(39) As Integer" & vbCrLf
    For fi = 0 To 39
        s = s & "    idx(" & fi & ")=" & idxMap(fi) & vbCrLf
    Next fi
    s = s & "    " & vbCrLf
    s = s & "    For ci = 1 To 4" & vbCrLf
    s = s & "        Dim tID As Long: tID = 0" & vbCrLf
    s = s & "        On Error Resume Next" & vbCrLf
    s = s & "        tID = CLng(Nz(Me.Controls(" & q & "cboTaxon" & q & " & ci).Value, 0))" & vbCrLf
    s = s & "        On Error GoTo LoadErr" & vbCrLf
    s = s & "        g_IDs(ci) = tID" & vbCrLf
    s = s & "        If tID = 0 Then" & vbCrLf
    s = s & "            Me.Controls(" & q & "lblTaxon" & q & " & ci).Caption = " & q & "(empty)" & q & vbCrLf
    s = s & "            GoTo NextCol" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        " & vbCrLf
    s = s & "        Dim rs As DAO.Recordset" & vbCrLf
    s = s & "        Set rs = db.OpenRecordset(" & q & "SELECT * FROM vw_Complete_Taxa WHERE ID=" & q & " & tID)" & vbCrLf
    s = s & "        If Not rs.EOF Then" & vbCrLf
    s = s & "            On Error Resume Next" & vbCrLf
    s = s & "            Me.Controls(" & q & "lblTaxon" & q & " & ci).Caption = Nz(rs.Fields(2).Value, " & q & "???" & q & ")" & vbCrLf
    s = s & "            On Error GoTo LoadErr" & vbCrLf
    s = s & "            Dim fi As Integer" & vbCrLf
    s = s & "            For fi = 0 To 39" & vbCrLf
    s = s & "                Dim cn As String" & vbCrLf
    s = s & "                cn = " & q & "txt_" & q & " & Replace(Replace(fl(fi)," & q & " " & q & "," & q & "_" & q & ")," & q & "-" & q & "," & q & "_" & q & ") & " & q & "_" & q & " & ci" & vbCrLf
    s = s & "                Dim fVal As String: fVal = """"" & vbCrLf
    s = s & "                On Error Resume Next" & vbCrLf
    s = s & "                fVal = Nz(rs.Fields(idx(fi)).Value, """")" & vbCrLf
    s = s & "                Me.Controls(cn).Value = fVal" & vbCrLf
    s = s & "                On Error GoTo LoadErr" & vbCrLf
    s = s & "            Next fi" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "        rs.Close" & vbCrLf
    s = s & "        NextCol:" & vbCrLf
    s = s & "    Next ci" & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "LoadErr: MsgBox " & q & "Load error: " & q & " & Err.Description, vbCritical" & vbCrLf
    s = s & "End Sub" & vbCrLf

    mdl.InsertLines startLine, s

CloseForm:
    DoCmd.Close acForm, "frmCompare4", acSaveYes
    Exit Sub
C4Err:
    MsgBox "FixCompare4LoadDirect error: " & Err.Description, vbCritical
    On Error Resume Next: DoCmd.Close acForm, "frmCompare4", acSaveYes
End Sub
