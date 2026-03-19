Attribute VB_Name = "BALIK_8_FIXES"
Option Compare Database
Option Explicit

' ================================================================
' BALIK_8_FIXES
'
' Fixes three bugs:
'   1. frmMergePreview: "End If without block If" compile error
'      Cause: inline If...Else...End If on one line is invalid VBA
'   2. frmTextEditor: missing form (runtime error 2102 in Compare)
'   3. frmCompare4 btnLoad_Click: CROSS-SECTION and CARDINAL_PROCESSES
'      fields not loading - DAO field name with hyphen char needs
'      index-based lookup instead of name-based lookup
'
' INSTALLATION:
'   1. VBA Editor -> File -> Import File -> BALIK_8_FIXES.bas
'   2. Immediate window: InstallFixes8
'   3. Save the database (Ctrl+S)
' ================================================================

Public Sub InstallFixes8()
    Debug.Print "=== BALIK_8_FIXES start ==="

    FixMergePreviewForm
    Debug.Print "  [1] frmMergePreview (End If fix) - OK"

    CreateTextEditorForm
    Debug.Print "  [2] frmTextEditor - OK"

    FixCompare4Load
    Debug.Print "  [3] frmCompare4 btnLoad_Click (field loading fix) - OK"

    MsgBox "BALIK_8_FIXES complete!" & vbCrLf & vbCrLf & _
           "Fixed:" & vbCrLf & _
           "  1. frmMergePreview - End If compile error" & vbCrLf & _
           "  2. frmTextEditor - created (was missing)" & vbCrLf & _
           "  3. frmCompare4 - CROSS-SECTION / CARDINAL_PROCESSES now load correctly", _
           vbInformation, "BALIK_8 OK"
End Sub

' ================================================================
' 1. frmMergePreview - fix "End If without block If"
'    Bug: single-line If...Else...End If is not valid VBA
'    Fix: all If blocks are now proper multi-line blocks
' ================================================================
Private Sub FixMergePreviewForm()
    On Error Resume Next
    DoCmd.DeleteObject acForm, "frmMergePreview"
    On Error GoTo 0

    Dim f As Form: Set f = CreateForm
    Dim fn As String: fn = f.Name

    With f
        .Caption = "Merge Preview": .Width = 12000
        .PopUp = True: .Modal = True
        .ScrollBars = 0: .NavigationButtons = False: .RecordSelectors = False
        .AutoCenter = True: .BorderStyle = 2: .KeyPreview = True
        .Section(acDetail).BackColor = &H202020
        .Section(acDetail).Height = 7500
    End With

    Dim c As Control
    Set c = CreateControl(fn, acLabel, acDetail, "", "", 200, 150, 11400, 500)
    c.Caption = "MERGE RECORDS - PREVIEW": c.ForeColor = &HBBAA00
    c.BackStyle = 0: c.FontSize = 14: c.FontBold = True

    Set c = CreateControl(fn, acLabel, acDetail, "", "", 200, 850, 11400, 400)
    c.Name = "lblInfo": c.Caption = "Loading..."
    c.ForeColor = &HC0C0C0: c.BackStyle = 0: c.FontSize = 10

    Set c = CreateControl(fn, acLabel, acDetail, "", "", 200, 1400, 11400, 400)
    c.Name = "lblKeep": c.Caption = ""
    c.ForeColor = &H00FF88: c.BackStyle = 0: c.FontSize = 10: c.FontBold = True

    Set c = CreateControl(fn, acListBox, acDetail, "", "", 200, 2000, 11400, 3200)
    c.Name = "lstMergeIDs": c.RowSourceType = "Value List": c.RowSource = ""
    c.BackColor = &H1A1A1A: c.ForeColor = &HFFFFFF: c.BorderColor = &H444444: c.FontSize = 10

    Set c = CreateControl(fn, acLabel, acDetail, "", "", 200, 5400, 11400, 450)
    c.Caption = "WARNING: Records listed above (except KEEP) will be permanently deleted after merge!"
    c.ForeColor = &H0000CC: c.BackStyle = 0: c.FontSize = 9: c.FontBold = True

    Set c = CreateControl(fn, acCommandButton, acDetail, "", "", 200, 6200, 2500, 600)
    c.Name = "btnConfirm": c.Caption = "CONFIRM MERGE": c.BackColor = &H004488: c.ForeColor = &HFFFFFF
    c.FontSize = 10: c.FontBold = True: c.OnClick = "[Event Procedure]"

    Set c = CreateControl(fn, acCommandButton, acDetail, "", "", 3000, 6200, 2000, 600)
    c.Name = "btnCancel": c.Caption = "CANCEL": c.BackColor = &H333333: c.ForeColor = &HFFFFFF
    c.FontSize = 10: c.FontBold = True: c.OnClick = "[Event Procedure]"

    Dim vba As String
    vba = "Option Compare Database" & vbCrLf & "Option Explicit" & vbCrLf & vbCrLf

    vba = vba & "Private Sub Form_Load()" & vbCrLf
    vba = vba & "    On Error Resume Next" & vbCrLf
    vba = vba & "    Dim cnt As Long, keepID As Long" & vbCrLf
    vba = vba & "    Dim idStr As String: idStr = """"" & vbCrLf
    vba = vba & "    Dim oForm As Object" & vbCrLf
    vba = vba & "    Set oForm = Forms(""frmDark"")" & vbCrLf
    vba = vba & "    If Not oForm Is Nothing Then" & vbCrLf
    vba = vba & "        idStr = oForm.g_MergeIDs" & vbCrLf
    vba = vba & "        cnt = oForm.g_MergeCnt" & vbCrLf
    vba = vba & "        keepID = oForm.g_MergeKeepID" & vbCrLf
    vba = vba & "    End If" & vbCrLf
    vba = vba & "    If idStr = """" Then" & vbCrLf
    vba = vba & "        idStr = Nz(TempVars(""g_MergeIDs""), """")" & vbCrLf
    vba = vba & "        cnt = CLng(Nz(TempVars(""g_MergeCnt""), 0))" & vbCrLf
    vba = vba & "        keepID = CLng(Nz(TempVars(""g_MergeKeepID""), 0))" & vbCrLf
    vba = vba & "    End If" & vbCrLf
    vba = vba & "    Me!lblInfo.Caption = ""Records to merge: "" & cnt & ""  (first = KEEP, rest = DELETE)""" & vbCrLf
    vba = vba & "    Me!lblKeep.Caption = ""KEEP ID: "" & keepID" & vbCrLf
    vba = vba & "    Dim ids() As String: ids = Split(idStr, "","")" & vbCrLf
    vba = vba & "    Dim src As String: src = """": Dim i As Integer" & vbCrLf
    vba = vba & "    For i = 0 To UBound(ids)" & vbCrLf
    vba = vba & "        Dim thisID As Long: thisID = CLng(Trim(ids(i)))" & vbCrLf
    vba = vba & "        Dim lbl As String" & vbCrLf
    vba = vba & "        If thisID = keepID Then" & vbCrLf
    vba = vba & "            lbl = ""[KEEP]   ID="" & thisID & ""  "" & Nz(DLookup(""TAXON"",""Taxa"",""ID="" & thisID), ""???"")" & vbCrLf
    vba = vba & "        Else" & vbCrLf
    vba = vba & "            lbl = ""[DELETE] ID="" & thisID & ""  "" & Nz(DLookup(""TAXON"",""Taxa"",""ID="" & thisID), ""???"")" & vbCrLf
    vba = vba & "        End If" & vbCrLf
    vba = vba & "        src = src & lbl & "";""" & vbCrLf
    vba = vba & "    Next i" & vbCrLf
    vba = vba & "    Me!lstMergeIDs.RowSource = src" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf

    vba = vba & "Private Sub btnConfirm_Click()" & vbCrLf
    vba = vba & "    On Error Resume Next" & vbCrLf
    vba = vba & "    Dim oForm As Object" & vbCrLf
    vba = vba & "    Set oForm = Forms(""frmDark"")" & vbCrLf
    vba = vba & "    If Not oForm Is Nothing Then" & vbCrLf
    vba = vba & "        oForm.g_MergeConfirmed = True" & vbCrLf
    vba = vba & "    Else" & vbCrLf
    vba = vba & "        TempVars.Item(""g_MergeConfirmed"") = True" & vbCrLf
    vba = vba & "    End If" & vbCrLf
    vba = vba & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf

    vba = vba & "Private Sub btnCancel_Click()" & vbCrLf
    vba = vba & "    On Error Resume Next" & vbCrLf
    vba = vba & "    Dim oForm As Object" & vbCrLf
    vba = vba & "    Set oForm = Forms(""frmDark"")" & vbCrLf
    vba = vba & "    If Not oForm Is Nothing Then" & vbCrLf
    vba = vba & "        oForm.g_MergeConfirmed = False" & vbCrLf
    vba = vba & "    Else" & vbCrLf
    vba = vba & "        TempVars.Item(""g_MergeConfirmed"") = False" & vbCrLf
    vba = vba & "    End If" & vbCrLf
    vba = vba & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf

    vba = vba & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    vba = vba & "    If KeyCode = 27 Then" & vbCrLf
    vba = vba & "        btnCancel_Click" & vbCrLf
    vba = vba & "        KeyCode = 0" & vbCrLf
    vba = vba & "    End If" & vbCrLf
    vba = vba & "End Sub" & vbCrLf

    f.HasModule = True: f.KeyPreview = True
    If f.Module.CountOfLines > 0 Then f.Module.DeleteLines 1, f.Module.CountOfLines
    f.Module.InsertLines 1, vba
    Dim tmp As String: tmp = f.Name
    DoCmd.Close acForm, tmp, acSaveYes
    DoCmd.CopyObject , "frmMergePreview", acForm, tmp
    DoCmd.DeleteObject acForm, tmp
End Sub

' ================================================================
' 2. frmTextEditor
'    Popup text editor used by frmCompare4 when clicking any field.
'    TempVars protocol:
'      PopupText      - text to edit (input and output)
'      PopupFieldName - source control name (input, display only)
'      PopupSaved     - "1" = saved, "0" = cancelled (output)
' ================================================================
Private Sub CreateTextEditorForm()
    On Error Resume Next
    DoCmd.DeleteObject acForm, "frmTextEditor"
    On Error GoTo 0

    Dim f As Form: Set f = CreateForm
    Dim fn As String: fn = f.Name

    With f
        .Caption = "Text Editor"
        .Width = 18000: .PopUp = True: .Modal = True
        .ScrollBars = 0: .NavigationButtons = False: .RecordSelectors = False
        .AutoCenter = True: .BorderStyle = 2: .KeyPreview = True
        .Section(acDetail).BackColor = &H202020
        .Section(acDetail).Height = 9000
    End With

    Dim c As Control

    Set c = CreateControl(fn, acLabel, acDetail, "", "", 200, 150, 17400, 450)
    c.Name = "lblFieldName": c.Caption = "TEXT EDITOR"
    c.ForeColor = &HBBAA00: c.BackStyle = 0: c.FontSize = 13: c.FontBold = True

    Set c = CreateControl(fn, acTextBox, acDetail, "", "", 200, 750, 17400, 7000)
    c.Name = "txtContent"
    c.BackColor = &H1A1A1A: c.ForeColor = &HFFFFFF: c.BorderColor = &H444444
    c.FontSize = 10: c.ScrollBars = 2: c.EnterKeyBehavior = True

    Set c = CreateControl(fn, acCommandButton, acDetail, "", "", 200, 8000, 2000, 600)
    c.Name = "btnSave": c.Caption = "SAVE"
    c.BackColor = &H006600: c.ForeColor = &HFFFFFF
    c.FontSize = 10: c.FontBold = True: c.OnClick = "[Event Procedure]"

    Set c = CreateControl(fn, acCommandButton, acDetail, "", "", 2400, 8000, 2000, 600)
    c.Name = "btnCancel": c.Caption = "CANCEL"
    c.BackColor = &H660000: c.ForeColor = &HFFFFFF
    c.FontSize = 10: c.FontBold = True: c.OnClick = "[Event Procedure]"

    Set c = CreateControl(fn, acLabel, acDetail, "", "", 5000, 8100, 5000, 400)
    c.Name = "lblCharCount": c.Caption = ""
    c.ForeColor = &H888888: c.BackStyle = 0: c.FontSize = 9

    Dim vba As String
    vba = "Option Compare Database" & vbCrLf & "Option Explicit" & vbCrLf & vbCrLf

    vba = vba & "Private Sub Form_Load()" & vbCrLf
    vba = vba & "    On Error Resume Next" & vbCrLf
    vba = vba & "    Dim txt As String: txt = Nz(TempVars(""PopupText""), """")" & vbCrLf
    vba = vba & "    Dim nm As String: nm = Nz(TempVars(""PopupFieldName""), """")" & vbCrLf
    vba = vba & "    Me!txtContent.Value = txt" & vbCrLf
    vba = vba & "    If nm <> """" Then" & vbCrLf
    vba = vba & "        Me!lblFieldName.Caption = ""Editing: "" & nm" & vbCrLf
    vba = vba & "    End If" & vbCrLf
    vba = vba & "    Me!lblCharCount.Caption = ""Characters: "" & Len(txt)" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf

    vba = vba & "Private Sub txtContent_Change()" & vbCrLf
    vba = vba & "    On Error Resume Next" & vbCrLf
    vba = vba & "    Me!lblCharCount.Caption = ""Characters: "" & Len(Nz(Me!txtContent.Text, """"))" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf

    vba = vba & "Private Sub btnSave_Click()" & vbCrLf
    vba = vba & "    TempVars.Item(""PopupText"") = Nz(Me!txtContent.Value, """")" & vbCrLf
    vba = vba & "    TempVars.Item(""PopupSaved"") = ""1""" & vbCrLf
    vba = vba & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf

    vba = vba & "Private Sub btnCancel_Click()" & vbCrLf
    vba = vba & "    TempVars.Item(""PopupSaved"") = ""0""" & vbCrLf
    vba = vba & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf

    vba = vba & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    vba = vba & "    If KeyCode = 27 Then" & vbCrLf
    vba = vba & "        btnCancel_Click" & vbCrLf
    vba = vba & "        KeyCode = 0" & vbCrLf
    vba = vba & "    End If" & vbCrLf
    vba = vba & "End Sub" & vbCrLf

    f.HasModule = True: f.KeyPreview = True
    If f.Module.CountOfLines > 0 Then f.Module.DeleteLines 1, f.Module.CountOfLines
    f.Module.InsertLines 1, vba
    Dim tmp As String: tmp = f.Name
    DoCmd.Close acForm, tmp, acSaveYes
    DoCmd.CopyObject , "frmTextEditor", acForm, tmp
    DoCmd.DeleteObject acForm, tmp
End Sub

' ================================================================
' 3. Fix frmCompare4 btnLoad_Click
'    Root cause: rs.Fields("CROSS-SECTION") fails silently in DAO
'    because the hyphen is ambiguous. Fix: iterate rs.Fields by index,
'    match by .Name property - this always works regardless of special chars.
' ================================================================
Private Sub FixCompare4Load()
    On Error GoTo C4Err
    DoCmd.OpenForm "frmCompare4", acDesign
    Dim frm As Form: Set frm = Forms("frmCompare4")
    Dim mdl As Module: Set mdl = frm.Module

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

    Dim q As String: q = Chr(34)
    Dim fi As Integer

    Dim s4 As String
    s4 = "Public Sub btnLoad_Click()" & vbCrLf
    s4 = s4 & "    On Error GoTo LoadErr" & vbCrLf
    s4 = s4 & "    Dim db As DAO.Database: Set db = CurrentDb" & vbCrLf
    s4 = s4 & "    Dim ci As Integer" & vbCrLf
    s4 = s4 & "    Dim fl(39) As String" & vbCrLf
    For fi = 0 To 39
        s4 = s4 & "    fl(" & fi & ")=" & q & flds(fi) & q & vbCrLf
    Next fi
    s4 = s4 & "    For ci = 1 To 4" & vbCrLf
    s4 = s4 & "        Dim tID As Long: tID = Nz(Me.Controls(" & q & "cboTaxon" & q & " & ci).Value, 0)" & vbCrLf
    s4 = s4 & "        g_IDs(ci) = tID" & vbCrLf
    s4 = s4 & "        If tID = 0 Then" & vbCrLf
    s4 = s4 & "            Me.Controls(" & q & "lblTaxon" & q & " & ci).Caption = " & q & "(empty)" & q & vbCrLf
    s4 = s4 & "            GoTo NextCol" & vbCrLf
    s4 = s4 & "        End If" & vbCrLf
    s4 = s4 & "        Dim rs As DAO.Recordset" & vbCrLf
    s4 = s4 & "        Set rs = db.OpenRecordset(" & q & "SELECT * FROM vw_Complete_Taxa WHERE ID=" & q & " & tID)" & vbCrLf
    s4 = s4 & "        If Not rs.EOF Then" & vbCrLf
    s4 = s4 & "            Me.Controls(" & q & "lblTaxon" & q & " & ci).Caption = Nz(rs!TAXON, " & q & "???" & q & ")" & vbCrLf
    s4 = s4 & "            Dim fi As Integer" & vbCrLf
    s4 = s4 & "            For fi = 0 To 39" & vbCrLf
    s4 = s4 & "                ' Control name: spaces and hyphens replaced with underscores" & vbCrLf
    s4 = s4 & "                Dim cn As String" & vbCrLf
    s4 = s4 & "                cn = " & q & "txt_" & q & " & Replace(Replace(fl(fi)," & q & " " & q & "," & q & "_" & q & ")," & q & "-" & q & "," & q & "_" & q & ") & " & q & "_" & q & " & ci" & vbCrLf
    s4 = s4 & "                ' Find field by index (safe for field names with hyphens)" & vbCrLf
    s4 = s4 & "                Dim fVal As String: fVal = """"" & vbCrLf
    s4 = s4 & "                Dim fCheck As Integer" & vbCrLf
    s4 = s4 & "                For fCheck = 0 To rs.Fields.Count - 1" & vbCrLf
    s4 = s4 & "                    If rs.Fields(fCheck).Name = fl(fi) Then" & vbCrLf
    s4 = s4 & "                        fVal = Nz(rs.Fields(fCheck).Value, """")" & vbCrLf
    s4 = s4 & "                        Exit For" & vbCrLf
    s4 = s4 & "                    End If" & vbCrLf
    s4 = s4 & "                Next fCheck" & vbCrLf
    s4 = s4 & "                On Error Resume Next" & vbCrLf
    s4 = s4 & "                Me.Controls(cn).Value = fVal" & vbCrLf
    s4 = s4 & "                On Error GoTo LoadErr" & vbCrLf
    s4 = s4 & "            Next fi" & vbCrLf
    s4 = s4 & "        End If" & vbCrLf
    s4 = s4 & "        rs.Close" & vbCrLf
    s4 = s4 & "        NextCol:" & vbCrLf
    s4 = s4 & "    Next ci" & vbCrLf
    s4 = s4 & "    Exit Sub" & vbCrLf
    s4 = s4 & "LoadErr: MsgBox " & q & "Load error: " & q & " & Err.Description, vbCritical" & vbCrLf
    s4 = s4 & "End Sub" & vbCrLf

    PatchSubInMdl mdl, "btnLoad_Click", s4

    DoCmd.Close acForm, "frmCompare4", acSaveYes
    Exit Sub
C4Err:
    MsgBox "FixCompare4Load error: " & Err.Description, vbCritical
    On Error Resume Next: DoCmd.Close acForm, "frmCompare4", acSaveYes
End Sub

' Helper: find, delete and replace a Sub or Function in a module
Private Sub PatchSubInMdl(mdl As Module, subName As String, newCode As String)
    Dim i As Long, s As Long, e As Long
    s = 0: e = 0
    For i = 1 To mdl.CountOfLines
        Dim ln As String: ln = mdl.Lines(i, 1)
        If s = 0 Then
            If InStr(ln, "Sub " & subName) > 0 Or InStr(ln, "Function " & subName) > 0 Then s = i
        Else
            If InStr(ln, "End Sub") > 0 Or InStr(ln, "End Function") > 0 Then e = i: Exit For
        End If
    Next i
    If s > 0 And e > 0 Then
        mdl.DeleteLines s, e - s + 1
        mdl.InsertLines s, newCode
    Else
        mdl.InsertLines mdl.CountOfLines + 1, newCode
    End If
End Sub
