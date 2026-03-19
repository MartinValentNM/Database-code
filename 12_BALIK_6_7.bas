Attribute VB_Name = "BALIK_6_7_FINAL"
Option Compare Database
Option Explicit

' ================================================================
' BALIK_6_7_FINAL  -  Unified installation module
' ================================================================
' What it installs:
'
'  [A] NEW FORMS (were missing, causing runtime errors)
'      A1. frmBulkFieldPicker  - field picker dialog for Bulk Edit
'      A2. frmMergePreview     - confirmation dialog before merge
'      A3. frmSavedFilters     - dialog for saving / loading filters
'
'  [B] REINSTALLED FORMS (bugs fixed)
'      B1. frmValidator  - fixed Inconsistent Family check:
'                          a) COUNT(DISTINCT) -> DAO cursor (Memo-safe)
'                          b) JOIN on Memo fields -> DCount technique
'
'  [C] TABLE
'      C1. FilterPresets  - storage for named filter presets
'                           (ID, FilterName, FilterData, SavedAt)
'
'  [D] PATCH frmDark (main search form)
'      D1. btnExport_Click   - scope: checked -> filtered -> all
'      D2. btnBulkEdit_Click - same scope logic
'      D3. btnSaveFilter_Click + btnLoadFilter_Click (new buttons)
'      D4. SerializeFilters / DeserializeFilters (helper functions)
'
'  [E] PATCH frmCompare4
'      E1. btnSaveAll_Click - detailed report of what was saved
'
' ================================================================
' INSTALLATION:
'   1. VBA Editor -> File -> Import File -> BALIK_6_7_FINAL.bas
'   2. Immediate window: InstallAll
'   3. Save the database (Ctrl+S)
' ================================================================

Public Sub InstallAll()
    Debug.Print "=== BALIK_6_7_FINAL start ==="

    ' A - new forms
    CreateBulkFieldPickerForm
    Debug.Print "  [A1] frmBulkFieldPicker - OK"
    CreateMergePreviewForm
    Debug.Print "  [A2] frmMergePreview - OK"

    ' B - reinstalled forms (fixed)
    CreateValidatorForm_Fixed
    Debug.Print "  [B1] frmValidator (Memo-safe) - OK"

    ' C - table
    CreateFilterPresetsTable
    Debug.Print "  [C1] FilterPresets table - OK"

    ' A3 - depends on table, must run after C1
    CreateSavedFiltersForm
    Debug.Print "  [A3] frmSavedFilters - OK"

    ' D - patch frmDark
    PatchFrmDark
    Debug.Print "  [D]  frmDark patch - OK"

    ' E - patch frmCompare4
    PatchCompare4SaveAll
    Debug.Print "  [E]  frmCompare4 SaveAll - OK"

    MsgBox "BALIK_6_7_FINAL complete!" & vbCrLf & vbCrLf & _
           "Installed:" & vbCrLf & _
           "  A1. frmBulkFieldPicker" & vbCrLf & _
           "  A2. frmMergePreview" & vbCrLf & _
           "  A3. frmSavedFilters" & vbCrLf & _
           "  B1. frmValidator (Memo/Inconsistent bug fixed)" & vbCrLf & _
           "  C1. Table FilterPresets" & vbCrLf & _
           "  D.  frmDark: Export/BulkEdit scope, SAVE/LOAD FILTER buttons" & vbCrLf & _
           "  E.  frmCompare4: Save All with detailed report", _
           vbInformation, "Installation OK"
End Sub

' ================================================================
' A1. frmBulkFieldPicker
'     Field picker dialog for Bulk Edit.
'     Communicates via TempVars:
'       BulkFieldList     - pipe-separated field list (input from frmDark)
'       BulkSelectedField - chosen field name (output)
'       BulkNewValue      - new value to apply (output)
'       BulkConfirmed     - True/False (output)
' ================================================================
Private Sub CreateBulkFieldPickerForm()
    On Error Resume Next
    DoCmd.DeleteObject acForm, "frmBulkFieldPicker"
    On Error GoTo 0

    Dim f As Form: Set f = CreateForm
    Dim fn As String: fn = f.Name

    With f
        .Caption = "Bulk Edit - Field Picker"
        .Width = 10000: .PopUp = True: .Modal = True
        .ScrollBars = 0: .NavigationButtons = False: .RecordSelectors = False
        .AutoCenter = True: .BorderStyle = 2: .KeyPreview = True
        .Section(acDetail).BackColor = &H202020
        .Section(acDetail).Height = 7200
    End With

    Dim c As Control
    Set c = CreateControl(fn, acLabel, acDetail, "", "", 200, 150, 9400, 500)
    c.Caption = "BULK EDIT - SELECT FIELD": c.ForeColor = &HBBAA00
    c.BackStyle = 0: c.FontSize = 14: c.FontBold = True

    Set c = CreateControl(fn, acLabel, acDetail, "", "", 200, 900, 4000, 350)
    c.Caption = "Field to edit:": c.ForeColor = &HC0C0C0: c.BackStyle = 0: c.FontSize = 10

    Set c = CreateControl(fn, acListBox, acDetail, "", "", 200, 1300, 9400, 2800)
    c.Name = "lstFields": c.RowSourceType = "Value List": c.RowSource = ""
    c.BackColor = &H1A1A1A: c.ForeColor = &HFFFFFF: c.BorderColor = &H444444: c.FontSize = 10

    Set c = CreateControl(fn, acLabel, acDetail, "", "", 200, 4300, 4000, 350)
    c.Caption = "New value:": c.ForeColor = &HC0C0C0: c.BackStyle = 0: c.FontSize = 10

    Set c = CreateControl(fn, acTextBox, acDetail, "", "", 200, 4700, 9400, 450)
    c.Name = "txtNewValue"
    c.BackColor = &H2A2A2A: c.ForeColor = &HFFFFFF: c.BorderColor = &H555555: c.FontSize = 10

    Set c = CreateControl(fn, acCommandButton, acDetail, "", "", 200, 5500, 2000, 600)
    c.Name = "btnOK": c.Caption = "OK": c.BackColor = &H884400: c.ForeColor = &HFFFFFF
    c.FontSize = 10: c.FontBold = True: c.OnClick = "[Event Procedure]"

    Set c = CreateControl(fn, acCommandButton, acDetail, "", "", 2400, 5500, 2000, 600)
    c.Name = "btnCancel": c.Caption = "CANCEL": c.BackColor = &H333333: c.ForeColor = &HFFFFFF
    c.FontSize = 10: c.FontBold = True: c.OnClick = "[Event Procedure]"

    Dim vba As String
    vba = "Option Compare Database" & vbCrLf & "Option Explicit" & vbCrLf & vbCrLf
    vba = vba & "Private Sub Form_Load()" & vbCrLf
    vba = vba & "    On Error Resume Next" & vbCrLf
    vba = vba & "    Dim raw As String: raw = Nz(TempVars(""BulkFieldList""), """")" & vbCrLf
    vba = vba & "    Dim parts() As String: parts = Split(raw, ""|"")" & vbCrLf
    vba = vba & "    Dim src As String: src = """": Dim i As Integer" & vbCrLf
    vba = vba & "    For i = 0 To UBound(parts)" & vbCrLf
    vba = vba & "        If Trim(parts(i)) <> """" Then src = src & Trim(parts(i)) & "";""" & vbCrLf
    vba = vba & "    Next i" & vbCrLf
    vba = vba & "    Me!lstFields.RowSource = src" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf
    vba = vba & "Private Sub btnOK_Click()" & vbCrLf
    vba = vba & "    If IsNull(Me!lstFields.Value) Or Me!lstFields.Value = """" Then" & vbCrLf
    vba = vba & "        MsgBox ""Please select a field."", vbExclamation: Exit Sub" & vbCrLf
    vba = vba & "    End If" & vbCrLf
    vba = vba & "    TempVars.Item(""BulkSelectedField"") = Me!lstFields.Value" & vbCrLf
    vba = vba & "    TempVars.Item(""BulkNewValue"") = Nz(Me!txtNewValue, """")" & vbCrLf
    vba = vba & "    TempVars.Item(""BulkConfirmed"") = True" & vbCrLf
    vba = vba & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf
    vba = vba & "Private Sub btnCancel_Click()" & vbCrLf
    vba = vba & "    TempVars.Item(""BulkConfirmed"") = False" & vbCrLf
    vba = vba & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf
    vba = vba & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    vba = vba & "    If KeyCode = 27 Then btnCancel_Click: KeyCode = 0" & vbCrLf
    vba = vba & "End Sub" & vbCrLf

    f.HasModule = True: f.KeyPreview = True
    If f.Module.CountOfLines > 0 Then f.Module.DeleteLines 1, f.Module.CountOfLines
    f.Module.InsertLines 1, vba
    Dim tmp As String: tmp = f.Name
    DoCmd.Close acForm, tmp, acSaveYes
    DoCmd.CopyObject , "frmBulkFieldPicker", acForm, tmp
    DoCmd.DeleteObject acForm, tmp
End Sub

' ================================================================
' A2. frmMergePreview
'     Confirmation dialog before merging records.
'     Reads global variables from frmDark (g_MergeIDs, g_MergeCnt,
'     g_MergeKeepID) and sets g_MergeConfirmed = True/False.
' ================================================================
Private Sub CreateMergePreviewForm()
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
    vba = vba & "    ' Try to read global variables from frmDark" & vbCrLf
    vba = vba & "    Dim oForm As Object: Set oForm = Forms(""frmDark"")" & vbCrLf
    vba = vba & "    If Not oForm Is Nothing Then" & vbCrLf
    vba = vba & "        idStr = oForm.g_MergeIDs: cnt = oForm.g_MergeCnt: keepID = oForm.g_MergeKeepID" & vbCrLf
    vba = vba & "    End If" & vbCrLf
    vba = vba & "    ' Fallback: try TempVars" & vbCrLf
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
    vba = vba & "    Dim oForm As Object: Set oForm = Forms(""frmDark"")" & vbCrLf
    vba = vba & "    If Not oForm Is Nothing Then: oForm.g_MergeConfirmed = True: Else: TempVars.Item(""g_MergeConfirmed"") = True: End If" & vbCrLf
    vba = vba & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf
    vba = vba & "Private Sub btnCancel_Click()" & vbCrLf
    vba = vba & "    On Error Resume Next" & vbCrLf
    vba = vba & "    Dim oForm As Object: Set oForm = Forms(""frmDark"")" & vbCrLf
    vba = vba & "    If Not oForm Is Nothing Then: oForm.g_MergeConfirmed = False: Else: TempVars.Item(""g_MergeConfirmed"") = False: End If" & vbCrLf
    vba = vba & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf
    vba = vba & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    vba = vba & "    If KeyCode = 27 Then btnCancel_Click: KeyCode = 0" & vbCrLf
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
' B1. frmValidator - fixed version
'     Bug 1: COUNT(DISTINCT ...) is not valid Access SQL syntax
'     Bug 2: JOIN on Memo fields (Genus, Family) fails in Access
'     Fix: DAO cursor + DCount for Inconsistent Family check
'          (slower than JOIN, but reliable on Memo fields)
' ================================================================
Private Sub CreateValidatorForm_Fixed()
    On Error Resume Next
    DoCmd.DeleteObject acForm, "frmValidator"
    On Error GoTo 0

    Dim f As Form: Set f = CreateForm
    Dim fn As String: fn = f.Name

    With f
        .Caption = "Data Validator": .Width = 18000: .PopUp = True: .Modal = False
        .ScrollBars = 0: .NavigationButtons = False: .RecordSelectors = False
        .AutoCenter = True: .BorderStyle = 2
        .Section(acDetail).BackColor = &H202020: .Section(acDetail).Height = 10000
    End With

    Dim c As Control
    Set c = CreateControl(fn, acLabel, acDetail, "", "", 200, 150, 14000, 600)
    c.Caption = "DATA VALIDATOR": c.ForeColor = &HBBAA00: c.BackStyle = 0
    c.FontSize = 16: c.FontBold = True

    Set c = CreateControl(fn, acCheckBox, acDetail, "", "", 200, 950, 300, 300)
    c.Name = "chkMissing": c.DefaultValue = "-1"
    Set c = CreateControl(fn, acLabel, acDetail, "", "", 600, 930, 6000, 400)
    c.Caption = "Missing required fields (TAXON, AUTHOR, Rank)": c.ForeColor = &HC0C0C0: c.BackStyle = 0: c.FontSize = 10

    Set c = CreateControl(fn, acCheckBox, acDetail, "", "", 200, 1450, 300, 300)
    c.Name = "chkDupTaxon": c.DefaultValue = "-1"
    Set c = CreateControl(fn, acLabel, acDetail, "", "", 600, 1430, 6000, 400)
    c.Caption = "Duplicate TAXON names": c.ForeColor = &HC0C0C0: c.BackStyle = 0: c.FontSize = 10

    Set c = CreateControl(fn, acCheckBox, acDetail, "", "", 200, 1950, 300, 300)
    c.Name = "chkInconsist": c.DefaultValue = "-1"
    Set c = CreateControl(fn, acLabel, acDetail, "", "", 600, 1930, 6000, 400)
    c.Caption = "Inconsistent Family/Order/Class (same genus, different values)": c.ForeColor = &HC0C0C0: c.BackStyle = 0: c.FontSize = 10

    Set c = CreateControl(fn, acCheckBox, acDetail, "", "", 200, 2450, 300, 300)
    c.Name = "chkEmptyGeog": c.DefaultValue = "-1"
    Set c = CreateControl(fn, acLabel, acDetail, "", "", 600, 2430, 6000, 400)
    c.Caption = "Records without Geography and Stratigraphy": c.ForeColor = &HC0C0C0: c.BackStyle = 0: c.FontSize = 10

    Set c = CreateControl(fn, acCheckBox, acDetail, "", "", 200, 2950, 300, 300)
    c.Name = "chkNoDesc": c.DefaultValue = "0"
    Set c = CreateControl(fn, acLabel, acDetail, "", "", 600, 2930, 6000, 400)
    c.Caption = "Records without Description": c.ForeColor = &HC0C0C0: c.BackStyle = 0: c.FontSize = 10

    Set c = CreateControl(fn, acCommandButton, acDetail, "", "", 200, 3600, 2000, 600)
    c.Name = "btnRun": c.Caption = "RUN VALIDATION": c.BackColor = &H884400: c.ForeColor = &HFFFFFF
    c.FontSize = 10: c.FontBold = True: c.OnClick = "[Event Procedure]"

    Set c = CreateControl(fn, acCommandButton, acDetail, "", "", 2400, 3600, 2000, 600)
    c.Name = "btnOpenDetail": c.Caption = "OPEN DETAIL": c.BackColor = &H228844: c.ForeColor = &HFFFFFF
    c.FontSize = 10: c.FontBold = True: c.OnClick = "[Event Procedure]"

    Set c = CreateControl(fn, acCommandButton, acDetail, "", "", 4600, 3600, 1500, 600)
    c.Name = "btnClose": c.Caption = "CLOSE": c.BackColor = &H111155: c.ForeColor = &HFFFFFF
    c.FontSize = 10: c.FontBold = True: c.OnClick = "[Event Procedure]"

    Set c = CreateControl(fn, acLabel, acDetail, "", "", 200, 4500, 17400, 400)
    c.Name = "lblSummary": c.Caption = "": c.ForeColor = &HFF8C00
    c.BackStyle = 0: c.FontSize = 10: c.FontBold = True

    Set c = CreateControl(fn, acListBox, acDetail, "", "", 200, 5100, 17400, 4500)
    c.Name = "lstResults": c.RowSourceType = "Value List": c.RowSource = ""
    c.BackColor = &H1A1A1A: c.ForeColor = &HFFFFFF: c.BorderColor = &H444444: c.FontSize = 10

    Dim q As String: q = Chr(34)
    Dim vba As String
    vba = "Option Compare Database" & vbCrLf & "Option Explicit" & vbCrLf & vbCrLf
    vba = vba & "Private g_IDs() As Long" & vbCrLf
    vba = vba & "Private g_IDCount As Long" & vbCrLf & vbCrLf

    ' btnRun_Click
    vba = vba & "Private Sub btnRun_Click()" & vbCrLf
    vba = vba & "    On Error GoTo ValErr" & vbCrLf
    vba = vba & "    Dim db As DAO.Database: Set db = CurrentDb" & vbCrLf
    vba = vba & "    Dim rs As DAO.Recordset" & vbCrLf
    vba = vba & "    Dim items As String: items = " & q & q & vbCrLf
    vba = vba & "    Dim cnt As Long: cnt = 0" & vbCrLf
    vba = vba & "    ReDim g_IDs(0 To 5000): g_IDCount = 0" & vbCrLf & vbCrLf

    ' Missing required fields
    vba = vba & "    If Me!chkMissing Then" & vbCrLf
    vba = vba & "        Set rs = db.OpenRecordset(" & q & "SELECT ID,TAXON,AUTHOR,Rank FROM Taxa WHERE Len(Nz(TAXON,''))=0 OR Len(Nz(AUTHOR,''))=0 OR Len(Nz(Rank,''))=0" & q & ")" & vbCrLf
    vba = vba & "        Do While Not rs.EOF" & vbCrLf
    vba = vba & "            Dim miss As String: miss = " & q & q & vbCrLf
    vba = vba & "            If Len(Nz(rs!TAXON," & q & q & "))=0 Then miss=miss & " & q & "TAXON " & q & vbCrLf
    vba = vba & "            If Len(Nz(rs!AUTHOR," & q & q & "))=0 Then miss=miss & " & q & "AUTHOR " & q & vbCrLf
    vba = vba & "            If Len(Nz(rs!Rank," & q & q & "))=0 Then miss=miss & " & q & "Rank " & q & vbCrLf
    vba = vba & "            items=items & " & q & "[MISSING:" & q & " & Trim(miss) & " & q & "] ID=" & q & " & rs!ID & " & q & "  " & q & " & Nz(rs!TAXON," & q & "???" & q & ") & " & q & ";" & q & vbCrLf
    vba = vba & "            g_IDs(g_IDCount)=rs!ID: g_IDCount=g_IDCount+1: cnt=cnt+1" & vbCrLf
    vba = vba & "            rs.MoveNext" & vbCrLf
    vba = vba & "        Loop: rs.Close" & vbCrLf
    vba = vba & "    End If" & vbCrLf & vbCrLf

    ' Duplicate TAXON names
    vba = vba & "    If Me!chkDupTaxon Then" & vbCrLf
    vba = vba & "        Set rs = db.OpenRecordset(" & q & "SELECT ID,TAXON FROM Taxa WHERE Len(Nz(TAXON,''))>0 ORDER BY ID" & q & ")" & vbCrLf
    vba = vba & "        Do While Not rs.EOF" & vbCrLf
    vba = vba & "            If DCount(" & q & "*" & q & "," & q & "Taxa" & q & "," & q & "TAXON='" & q & " & Replace(Nz(rs!TAXON," & q & q & ")," & q & "'" & q & "," & q & "''" & q & ") & " & q & "'" & q & ")>1 Then" & vbCrLf
    vba = vba & "                items=items & " & q & "[DUPLICATE] ID=" & q & " & rs!ID & " & q & "  " & q & " & Nz(rs!TAXON," & q & q & ") & " & q & ";" & q & vbCrLf
    vba = vba & "                g_IDs(g_IDCount)=rs!ID: g_IDCount=g_IDCount+1: cnt=cnt+1" & vbCrLf
    vba = vba & "            End If" & vbCrLf
    vba = vba & "            rs.MoveNext" & vbCrLf
    vba = vba & "        Loop: rs.Close" & vbCrLf
    vba = vba & "    End If" & vbCrLf & vbCrLf

    ' Inconsistent Family - DAO cursor + DCount (Memo-safe, no JOIN, no COUNT DISTINCT)
    vba = vba & "    If Me!chkInconsist Then" & vbCrLf
    vba = vba & "        Set rs = db.OpenRecordset(" & q & "SELECT ID,TAXON,Genus,Family FROM Taxa WHERE Len(Nz(Genus,''))>0 AND Len(Nz(Family,''))>0 ORDER BY Genus,ID" & q & ")" & vbCrLf
    vba = vba & "        Do While Not rs.EOF" & vbCrLf
    vba = vba & "            Dim gGenus As String: gGenus = Left(Nz(rs!Genus," & q & q & "),100)" & vbCrLf
    vba = vba & "            Dim gFamily As String: gFamily = Left(Nz(rs!Family," & q & q & "),100)" & vbCrLf
    vba = vba & "            Dim safeG As String: safeG = Replace(gGenus," & q & "'" & q & "," & q & "''" & q & ")" & vbCrLf
    vba = vba & "            Dim safeF As String: safeF = Replace(gFamily," & q & "'" & q & "," & q & "''" & q & ")" & vbCrLf
    vba = vba & "            Dim crit As String" & vbCrLf
    vba = vba & "            crit = " & q & "Left(Nz(Genus,''),100)='" & q & " & safeG & " & q & "' AND Left(Nz(Family,''),100)<>'" & q & " & safeF & " & q & "' AND ID<>" & q & " & rs!ID" & vbCrLf
    vba = vba & "            If DCount(" & q & "*" & q & "," & q & "Taxa" & q & ",crit)>0 Then" & vbCrLf
    vba = vba & "                items=items & " & q & "[INCONSIST.Fam] ID=" & q & " & rs!ID & " & q & "  " & q & " & Nz(rs!TAXON," & q & q & ") & " & q & " [" & q & " & gGenus & " & q & " / " & q & " & gFamily & " & q & "];" & q & vbCrLf
    vba = vba & "                g_IDs(g_IDCount)=rs!ID: g_IDCount=g_IDCount+1: cnt=cnt+1" & vbCrLf
    vba = vba & "            End If" & vbCrLf
    vba = vba & "            rs.MoveNext" & vbCrLf
    vba = vba & "        Loop: rs.Close" & vbCrLf
    vba = vba & "    End If" & vbCrLf & vbCrLf

    ' Records without Geography and Stratigraphy
    vba = vba & "    If Me!chkEmptyGeog Then" & vbCrLf
    vba = vba & "        Set rs = db.OpenRecordset(" & q & "SELECT t.ID,t.TAXON FROM Taxa t LEFT JOIN Taxa_Vyskyt v ON t.ID=v.TaxonID WHERE Len(Nz(v.Geography,''))=0 AND Len(Nz(v.STRATIGRAPHY,''))=0" & q & ")" & vbCrLf
    vba = vba & "        Do While Not rs.EOF" & vbCrLf
    vba = vba & "            items=items & " & q & "[NO LOCATION] ID=" & q & " & rs!ID & " & q & "  " & q & " & Nz(rs!TAXON," & q & q & ") & " & q & ";" & q & vbCrLf
    vba = vba & "            g_IDs(g_IDCount)=rs!ID: g_IDCount=g_IDCount+1: cnt=cnt+1" & vbCrLf
    vba = vba & "            rs.MoveNext" & vbCrLf
    vba = vba & "        Loop: rs.Close" & vbCrLf
    vba = vba & "    End If" & vbCrLf & vbCrLf

    ' Records without Description
    vba = vba & "    If Me!chkNoDesc Then" & vbCrLf
    vba = vba & "        Set rs = db.OpenRecordset(" & q & "SELECT t.ID,t.TAXON FROM Taxa t LEFT JOIN Taxa_Popis p ON t.ID=p.TaxonID WHERE Len(Nz(p.DESCRIPTION,''))=0" & q & ")" & vbCrLf
    vba = vba & "        Do While Not rs.EOF" & vbCrLf
    vba = vba & "            items=items & " & q & "[NO DESCRIPTION] ID=" & q & " & rs!ID & " & q & "  " & q & " & Nz(rs!TAXON," & q & q & ") & " & q & ";" & q & vbCrLf
    vba = vba & "            g_IDs(g_IDCount)=rs!ID: g_IDCount=g_IDCount+1: cnt=cnt+1" & vbCrLf
    vba = vba & "            rs.MoveNext" & vbCrLf
    vba = vba & "        Loop: rs.Close" & vbCrLf
    vba = vba & "    End If" & vbCrLf & vbCrLf

    vba = vba & "    Me!lstResults.RowSource = items" & vbCrLf
    vba = vba & "    Me!lblSummary.Caption = " & q & "Issues found: " & q & " & cnt" & vbCrLf
    vba = vba & "    Exit Sub" & vbCrLf
    vba = vba & "ValErr: MsgBox " & q & "Error: " & q & " & Err.Description, vbCritical" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf

    ' btnOpenDetail_Click
    vba = vba & "Private Sub btnOpenDetail_Click()" & vbCrLf
    vba = vba & "    On Error Resume Next" & vbCrLf
    vba = vba & "    If IsNull(Me!lstResults.Value) Then MsgBox " & q & "Select a record." & q & ",vbExclamation: Exit Sub" & vbCrLf
    vba = vba & "    Dim sel As String: sel = Me!lstResults.Value" & vbCrLf
    vba = vba & "    Dim p1 As Long: p1 = InStr(sel," & q & "ID=" & q & ")" & vbCrLf
    vba = vba & "    If p1=0 Then Exit Sub" & vbCrLf
    vba = vba & "    Dim idStr As String: idStr = Mid(sel,p1+3)" & vbCrLf
    vba = vba & "    Dim sp As Long: sp = InStr(idStr," & q & " " & q & "): If sp>0 Then idStr=Left(idStr,sp-1)" & vbCrLf
    vba = vba & "    Dim tID As Long: tID = CLng(idStr)" & vbCrLf
    vba = vba & "    DoCmd.OpenForm " & q & "frmDetailTaxa" & q & ",acNormal,," & q & "ID=" & q & " & tID" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf
    vba = vba & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    vba = vba & "    If KeyCode=27 Then DoCmd.Close acForm,Me.Name: KeyCode=0" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf
    vba = vba & "Private Sub btnClose_Click(): DoCmd.Close acForm,Me.Name: End Sub" & vbCrLf

    f.HasModule = True: f.KeyPreview = True
    If f.Module.CountOfLines > 0 Then f.Module.DeleteLines 1, f.Module.CountOfLines
    f.Module.InsertLines 1, vba
    Dim tmp As String: tmp = f.Name
    DoCmd.Close acForm, tmp, acSaveYes
    DoCmd.CopyObject , "frmValidator", acForm, tmp
    DoCmd.DeleteObject acForm, tmp
End Sub

' ================================================================
' C1. Table FilterPresets
'     Columns: ID (PK AutoNumber), FilterName (Text 100),
'              FilterData (Memo), SavedAt (Date)
'     Skipped silently if the table already exists.
' ================================================================
Private Sub CreateFilterPresetsTable()
    On Error Resume Next
    Dim chk As DAO.TableDef
    Set chk = CurrentDb.TableDefs("FilterPresets")
    If Not chk Is Nothing Then Exit Sub   ' already exists
    On Error GoTo 0

    Dim db As DAO.Database: Set db = CurrentDb
    Dim td As DAO.TableDef: Set td = db.CreateTableDef("FilterPresets")

    Dim f1 As DAO.Field: Set f1 = td.CreateField("ID", dbLong)
    f1.Attributes = dbAutoIncrField: td.Fields.Append f1

    Dim f2 As DAO.Field: Set f2 = td.CreateField("FilterName", dbText, 100)
    td.Fields.Append f2

    Dim f3 As DAO.Field: Set f3 = td.CreateField("FilterData", dbMemo)
    td.Fields.Append f3

    Dim f4 As DAO.Field: Set f4 = td.CreateField("SavedAt", dbDate)
    td.Fields.Append f4

    Dim idx As DAO.Index: Set idx = td.CreateIndex("PrimaryKey")
    idx.Primary = True
    idx.Fields.Append idx.CreateField("ID")
    td.Indexes.Append idx

    db.TableDefs.Append td
    db.TableDefs.Refresh
End Sub

' ================================================================
' A3. frmSavedFilters
'     Dialog for managing saved filter presets.
'     Communicates with frmDark via TempVars:
'       FilterDataToSave  - serialized filter data (input on Save)
'       FilterAction      - "load" / "cancel" (output)
'       FilterData        - data to load (output on Load)
'       FilterName        - name of selected preset (output)
' ================================================================
Private Sub CreateSavedFiltersForm()
    On Error Resume Next
    DoCmd.DeleteObject acForm, "frmSavedFilters"
    On Error GoTo 0

    Dim f As Form: Set f = CreateForm
    Dim fn As String: fn = f.Name

    With f
        .Caption = "Saved Filters": .Width = 12000
        .PopUp = True: .Modal = True
        .ScrollBars = 0: .NavigationButtons = False: .RecordSelectors = False
        .AutoCenter = True: .BorderStyle = 2: .KeyPreview = True
        .Section(acDetail).BackColor = &H202020
        .Section(acDetail).Height = 8500
    End With

    Dim c As Control
    Set c = CreateControl(fn, acLabel, acDetail, "", "", 200, 150, 11400, 500)
    c.Caption = "SAVED FILTERS": c.ForeColor = &HBBAA00
    c.BackStyle = 0: c.FontSize = 14: c.FontBold = True

    Set c = CreateControl(fn, acListBox, acDetail, "", "", 200, 850, 11400, 4000)
    c.Name = "lstFilters": c.RowSourceType = "Table/Query"
    c.RowSource = "SELECT ID,FilterName,SavedAt FROM FilterPresets ORDER BY SavedAt DESC"
    c.ColumnCount = 3: c.ColumnWidths = "0cm;6cm;3cm": c.BoundColumn = 1
    c.BackColor = &H1A1A1A: c.ForeColor = &HFFFFFF: c.BorderColor = &H444444: c.FontSize = 10

    Set c = CreateControl(fn, acLabel, acDetail, "", "", 200, 5050, 5000, 350)
    c.Caption = "Save current filter as:": c.ForeColor = &HC0C0C0: c.BackStyle = 0: c.FontSize = 10

    Set c = CreateControl(fn, acTextBox, acDetail, "", "", 200, 5450, 8000, 420)
    c.Name = "txtFilterName"
    c.BackColor = &H2A2A2A: c.ForeColor = &HFFFFFF: c.BorderColor = &H555555: c.FontSize = 10

    Set c = CreateControl(fn, acCommandButton, acDetail, "", "", 200, 6200, 2000, 600)
    c.Name = "btnSave": c.Caption = "SAVE": c.BackColor = &H006600: c.ForeColor = &HFFFFFF
    c.FontSize = 10: c.FontBold = True: c.OnClick = "[Event Procedure]"

    Set c = CreateControl(fn, acCommandButton, acDetail, "", "", 2400, 6200, 2400, 600)
    c.Name = "btnLoad": c.Caption = "LOAD SELECTED": c.BackColor = &H004488: c.ForeColor = &HFFFFFF
    c.FontSize = 10: c.FontBold = True: c.OnClick = "[Event Procedure]"

    Set c = CreateControl(fn, acCommandButton, acDetail, "", "", 5000, 6200, 2000, 600)
    c.Name = "btnDelete": c.Caption = "DELETE": c.BackColor = &H660000: c.ForeColor = &HFFFFFF
    c.FontSize = 10: c.FontBold = True: c.OnClick = "[Event Procedure]"

    Set c = CreateControl(fn, acCommandButton, acDetail, "", "", 9200, 6200, 2000, 600)
    c.Name = "btnClose": c.Caption = "CLOSE": c.BackColor = &H333333: c.ForeColor = &HFFFFFF
    c.FontSize = 10: c.FontBold = True: c.OnClick = "[Event Procedure]"

    Set c = CreateControl(fn, acLabel, acDetail, "", "", 200, 7100, 11000, 400)
    c.Name = "lblStatus": c.Caption = "": c.ForeColor = &H00BBFF: c.BackStyle = 0: c.FontSize = 9

    Dim vba As String
    vba = "Option Compare Database" & vbCrLf & "Option Explicit" & vbCrLf & vbCrLf

    vba = vba & "Private Sub btnSave_Click()" & vbCrLf
    vba = vba & "    Dim nm As String: nm = Trim(Nz(Me!txtFilterName, """"))" & vbCrLf
    vba = vba & "    If nm = """" Then MsgBox ""Enter a filter name."", vbExclamation: Exit Sub" & vbCrLf
    vba = vba & "    Dim data As String: data = Nz(TempVars(""FilterDataToSave""), """")" & vbCrLf
    vba = vba & "    If data = """" Then MsgBox ""No filter data. Apply a filter in the main form first."", vbExclamation: Exit Sub" & vbCrLf
    vba = vba & "    Dim existID As Long" & vbCrLf
    vba = vba & "    existID = Nz(DLookup(""ID"",""FilterPresets"",""FilterName='"" & Replace(nm,""'"",""''"") & ""'""), 0)" & vbCrLf
    vba = vba & "    If existID > 0 Then" & vbCrLf
    vba = vba & "        If MsgBox(""Filter '"" & nm & ""' exists. Overwrite?"", vbYesNo) = vbNo Then Exit Sub" & vbCrLf
    vba = vba & "        CurrentDb.Execute ""UPDATE FilterPresets SET FilterData='"" & Replace(data,""'"",""''"") & ""',SavedAt=Now() WHERE ID="" & existID, 128" & vbCrLf
    vba = vba & "    Else" & vbCrLf
    vba = vba & "        CurrentDb.Execute ""INSERT INTO FilterPresets(FilterName,FilterData,SavedAt) VALUES('"" & Replace(nm,""'"",""''"") & ""','"" & Replace(data,""'"",""''"") & ""',Now())"", 128" & vbCrLf
    vba = vba & "    End If" & vbCrLf
    vba = vba & "    Me!lstFilters.Requery" & vbCrLf
    vba = vba & "    Me!lblStatus.Caption = ""Saved: "" & nm" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf

    vba = vba & "Private Sub btnLoad_Click()" & vbCrLf
    vba = vba & "    If IsNull(Me!lstFilters.Value) Then MsgBox ""Select a filter."", vbExclamation: Exit Sub" & vbCrLf
    vba = vba & "    Dim selID As Long: selID = CLng(Me!lstFilters.Value)" & vbCrLf
    vba = vba & "    Dim data As String: data = Nz(DLookup(""FilterData"",""FilterPresets"",""ID="" & selID), """")" & vbCrLf
    vba = vba & "    If data = """" Then MsgBox ""Filter data is empty."", vbExclamation: Exit Sub" & vbCrLf
    vba = vba & "    TempVars.Item(""FilterAction"") = ""load""" & vbCrLf
    vba = vba & "    TempVars.Item(""FilterData"") = data" & vbCrLf
    vba = vba & "    TempVars.Item(""FilterName"") = Nz(DLookup(""FilterName"",""FilterPresets"",""ID="" & selID), """")" & vbCrLf
    vba = vba & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf

    vba = vba & "Private Sub btnDelete_Click()" & vbCrLf
    vba = vba & "    If IsNull(Me!lstFilters.Value) Then MsgBox ""Select a filter."", vbExclamation: Exit Sub" & vbCrLf
    vba = vba & "    Dim selID As Long: selID = CLng(Me!lstFilters.Value)" & vbCrLf
    vba = vba & "    Dim nm As String: nm = Nz(DLookup(""FilterName"",""FilterPresets"",""ID="" & selID), """")" & vbCrLf
    vba = vba & "    If MsgBox(""Delete filter '"" & nm & ""'?"", vbYesNo+vbQuestion) = vbNo Then Exit Sub" & vbCrLf
    vba = vba & "    CurrentDb.Execute ""DELETE FROM FilterPresets WHERE ID="" & selID, 128" & vbCrLf
    vba = vba & "    Me!lstFilters.Requery" & vbCrLf
    vba = vba & "    Me!lblStatus.Caption = ""Deleted: "" & nm" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf

    vba = vba & "Private Sub btnClose_Click()" & vbCrLf
    vba = vba & "    TempVars.Item(""FilterAction"") = ""cancel""" & vbCrLf
    vba = vba & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    vba = vba & "End Sub" & vbCrLf & vbCrLf
    vba = vba & "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    vba = vba & "    If KeyCode=27 Then btnClose_Click: KeyCode=0" & vbCrLf
    vba = vba & "End Sub" & vbCrLf

    f.HasModule = True: f.KeyPreview = True
    If f.Module.CountOfLines > 0 Then f.Module.DeleteLines 1, f.Module.CountOfLines
    f.Module.InsertLines 1, vba
    Dim tmp As String: tmp = f.Name
    DoCmd.Close acForm, tmp, acSaveYes
    DoCmd.CopyObject , "frmSavedFilters", acForm, tmp
    DoCmd.DeleteObject acForm, tmp
End Sub

' ================================================================
' D. Patch frmDark
'    Replaces / adds functions directly in the existing form module.
'    New functions: scope-aware Export, scope-aware BulkEdit,
'                   SaveFilter, LoadFilter, SerializeFilters,
'                   DeserializeFilters
'    New buttons:   btnSaveFilter, btnLoadFilter (Row C)
' ================================================================
Private Sub PatchFrmDark()
    On Error GoTo PatchDarkErr
    DoCmd.OpenForm "frmDark", acDesign
    Dim frm As Form: Set frm = Forms("frmDark")
    Dim mdl As Module: Set mdl = frm.Module

    ' Add buttons if they do not exist yet
    On Error Resume Next
    Dim dummy As Control: Set dummy = frm.Controls("btnSaveFilter")
    Dim btnsMissing As Boolean: btnsMissing = (Err.Number <> 0)
    Err.Clear: On Error GoTo PatchDarkErr

    If btnsMissing Then
        Dim histY As Long: histY = 0
        On Error Resume Next
        histY = frm.Controls("btnHistory").Top
        On Error GoTo PatchDarkErr

        Dim c As Control
        Set c = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 3800, histY, 2000, 400)
        c.Name = "btnSaveFilter": c.Caption = "SAVE FILTER"
        c.BackColor = &H336600: c.ForeColor = &HFFFFFF
        c.FontSize = 9: c.FontBold = True: c.OnClick = "[Event Procedure]"

        Set c = CreateControl(frm.Name, acCommandButton, acDetail, "", "", 5900, histY, 2000, 400)
        c.Name = "btnLoadFilter": c.Caption = "LOAD FILTER"
        c.BackColor = &H004466: c.ForeColor = &HFFFFFF
        c.FontSize = 9: c.FontBold = True: c.OnClick = "[Event Procedure]"
    End If

    ' Replace / append all patched functions
    PatchSub mdl, "btnExport_Click", CodeExport()
    PatchSub mdl, "btnBulkEdit_Click", CodeBulkEdit()
    PatchSub mdl, "btnSaveFilter_Click", CodeSaveFilter()
    PatchSub mdl, "btnLoadFilter_Click", CodeLoadFilter()
    PatchSub mdl, "SerializeFilters", CodeSerialize()
    PatchSub mdl, "DeserializeFilters", CodeDeserialize()

    DoCmd.Close acForm, "frmDark", acSaveYes
    Exit Sub
PatchDarkErr:
    MsgBox "PatchFrmDark error: " & Err.Description, vbCritical
    On Error Resume Next: DoCmd.Close acForm, "frmDark", acSaveYes
End Sub

' Helper: finds a Sub/Function in the module by name,
'         deletes it and inserts the new code at the same position.
'         If not found, appends to the end.
Private Sub PatchSub(mdl As Module, subName As String, newCode As String)
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

' ----------------------------------------------------------------
' D1. Export XLSX - scope: checked -> filtered -> all
' ----------------------------------------------------------------
Private Function CodeExport() As String
    Dim q As String: q = Chr(34)
    Dim s As String
    s = "Private Sub btnExport_Click()" & vbCrLf
    s = s & "    On Error GoTo ExportErr" & vbCrLf
    s = s & "    Dim db As DAO.Database: Set db = CurrentDb" & vbCrLf
    s = s & "    Dim sql As String, scopeLabel As String" & vbCrLf
    s = s & "    " & vbCrLf
    s = s & "    ' Build ID list from checked (Compare = True) records" & vbCrLf
    s = s & "    Dim rs As DAO.Recordset: Set rs = Me.subResults.Form.Recordset.Clone" & vbCrLf
    s = s & "    Dim idList As String: idList = """"" & vbCrLf
    s = s & "    On Error Resume Next: rs.MoveFirst: On Error GoTo ExportErr" & vbCrLf
    s = s & "    Do While Not rs.EOF" & vbCrLf
    s = s & "        If Nz(rs.Fields(" & q & "Compare" & q & "),0)=True Then idList=idList & rs.Fields(" & q & "ID" & q & ") & "","" " & vbCrLf
    s = s & "        rs.MoveNext" & vbCrLf
    s = s & "    Loop: rs.Close" & vbCrLf
    s = s & "    " & vbCrLf
    s = s & "    If idList <> """" Then" & vbCrLf
    s = s & "        ' Checked records take priority" & vbCrLf
    s = s & "        idList = Left(idList,Len(idList)-1)" & vbCrLf
    s = s & "        sql = " & q & "SELECT * FROM vw_Complete_Taxa WHERE ID IN (" & q & " & idList & " & q & ")" & q & vbCrLf
    s = s & "        scopeLabel = " & q & "checked taxa" & q & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        ' Fallback: use current subform RecordSource (= active filter)" & vbCrLf
    s = s & "        sql = Me.subResults.Form.RecordSource" & vbCrLf
    s = s & "        If sql = """" Then sql = " & q & "SELECT * FROM vw_Complete_Taxa" & q & vbCrLf
    s = s & "        scopeLabel = IIf(BuildWhereClause("""") <> """", " & q & "filtered taxa" & q & ", " & q & "ALL taxa" & q & ")" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    " & vbCrLf
    s = s & "    If MsgBox(" & q & "Export " & q & " & scopeLabel & " & q & " to XLSX?" & q & ",vbYesNo+vbQuestion," & q & "Export" & q & ")=vbNo Then Exit Sub" & vbCrLf
    s = s & "    On Error Resume Next: db.QueryDefs.Delete " & q & "qryExportTemp" & q & ": On Error GoTo ExportErr" & vbCrLf
    s = s & "    db.CreateQueryDef " & q & "qryExportTemp" & q & ", sql" & vbCrLf
    s = s & "    DoCmd.TransferSpreadsheet acExport,acSpreadsheetTypeExcel12Xml," & q & "qryExportTemp" & q & ",CurrentProject.Path & " & q & "\Export.xlsx" & q & ",True" & vbCrLf
    s = s & "    MsgBox " & q & "Exported " & q & " & scopeLabel & " & q & " to:" & q & " & Chr(10) & CurrentProject.Path & " & q & "\Export.xlsx" & q & ",vbInformation," & q & "Export done" & q & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "ExportErr: MsgBox " & q & "Export error: " & q & " & Err.Description,vbCritical" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeExport = s
End Function

' ----------------------------------------------------------------
' D2. Bulk Edit - scope: checked -> filtered -> all
' ----------------------------------------------------------------
Private Function CodeBulkEdit() As String
    Dim q As String: q = Chr(34)
    Dim s As String
    s = "Private Sub btnBulkEdit_Click()" & vbCrLf
    s = s & "    On Error GoTo BulkErr" & vbCrLf
    s = s & "    " & vbCrLf
    s = s & "    ' Determine scope: checked / filtered / all" & vbCrLf
    s = s & "    Dim rs As DAO.Recordset: Set rs = Me.subResults.Form.Recordset.Clone" & vbCrLf
    s = s & "    Dim idList As String: idList = """"" & vbCrLf
    s = s & "    On Error Resume Next: rs.MoveFirst: On Error GoTo BulkErr" & vbCrLf
    s = s & "    Do While Not rs.EOF" & vbCrLf
    s = s & "        If Nz(rs.Fields(" & q & "Compare" & q & "),0)=True Then idList=idList & rs.Fields(" & q & "ID" & q & ") & "","" " & vbCrLf
    s = s & "        rs.MoveNext" & vbCrLf
    s = s & "    Loop: rs.Close" & vbCrLf
    s = s & "    Dim useChecked As Boolean: useChecked = (idList <> """")" & vbCrLf
    s = s & "    If useChecked Then idList = Left(idList,Len(idList)-1)" & vbCrLf
    s = s & "    Dim scopeLabel As String" & vbCrLf
    s = s & "    scopeLabel = IIf(useChecked," & q & "CHECKED taxa" & q & ",IIf(BuildWhereClause("""") <> """"," & q & "FILTERED taxa" & q & "," & q & "ALL taxa (no filter active!)" & q & "))" & vbCrLf
    s = s & "    " & vbCrLf
    s = s & "    ' Build field list from vw_Complete_Taxa" & vbCrLf
    s = s & "    Dim db As DAO.Database: Set db = CurrentDb" & vbCrLf
    s = s & "    Dim qd As DAO.QueryDef: On Error Resume Next" & vbCrLf
    s = s & "    Set qd = db.QueryDefs(" & q & "vw_Complete_Taxa" & q & ")" & vbCrLf
    s = s & "    If qd Is Nothing Then MsgBox " & q & "View vw_Complete_Taxa not found." & q & ",vbCritical: Exit Sub" & vbCrLf
    s = s & "    On Error GoTo BulkErr" & vbCrLf
    s = s & "    Dim fieldList As String: fieldList = """": Dim fl As DAO.Field" & vbCrLf
    s = s & "    For Each fl In qd.Fields" & vbCrLf
    s = s & "        If fl.Name <> " & q & "ID" & q & " And fl.Name <> " & q & "Compare" & q & " Then fieldList=fieldList & fl.Name & ""|""" & vbCrLf
    s = s & "    Next fl" & vbCrLf
    s = s & "    " & vbCrLf
    s = s & "    TempVars.Add " & q & "BulkFieldList" & q & ",fieldList" & vbCrLf
    s = s & "    TempVars.Add " & q & "BulkSelectedField" & q & ",""""" & vbCrLf
    s = s & "    TempVars.Add " & q & "BulkNewValue" & q & ",""""" & vbCrLf
    s = s & "    TempVars.Add " & q & "BulkConfirmed" & q & ",False" & vbCrLf
    s = s & "    DoCmd.OpenForm " & q & "frmBulkFieldPicker" & q & ",acNormal,,,,acDialog" & vbCrLf
    s = s & "    " & vbCrLf
    s = s & "    If Not TempVars(" & q & "BulkConfirmed" & q & ") Then Exit Sub" & vbCrLf
    s = s & "    Dim fldName As String: fldName = TempVars(" & q & "BulkSelectedField" & q & ")" & vbCrLf
    s = s & "    Dim newVal As String: newVal = TempVars(" & q & "BulkNewValue" & q & ")" & vbCrLf
    s = s & "    If fldName = """" Then Exit Sub" & vbCrLf
    s = s & "    " & vbCrLf
    s = s & "    ' Confirm before executing" & vbCrLf
    s = s & "    If MsgBox(" & q & "BULK EDIT" & q & " & Chr(10) & " & q & "Field: " & q & " & fldName & Chr(10) & " & q & "New value: " & q & " & newVal & Chr(10) & " & q & "Scope: " & q & " & scopeLabel,vbYesNo+vbExclamation," & q & "Confirm" & q & ")=vbNo Then Exit Sub" & vbCrLf
    s = s & "    " & vbCrLf
    s = s & "    Dim sqlSel As String, sqlUpd As String, wc As String" & vbCrLf
    s = s & "    If useChecked Then" & vbCrLf
    s = s & "        sqlSel = " & q & "SELECT ID FROM vw_Complete_Taxa WHERE ID IN (" & q & " & idList & " & q & ")" & q & vbCrLf
    s = s & "        sqlUpd = " & q & "UPDATE vw_Complete_Taxa SET [" & q & " & fldName & " & q & "]='" & q & " & Replace(newVal," & q & "'" & q & "," & q & "''" & q & ") & " & q & "' WHERE ID IN (" & q & " & idList & " & q & ")" & q & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        wc = BuildWhereClause("""")" & vbCrLf
    s = s & "        sqlSel = " & q & "SELECT ID FROM vw_Complete_Taxa" & q & " & IIf(wc<>"""","" WHERE "" & wc,"""")" & vbCrLf
    s = s & "        sqlUpd = " & q & "UPDATE vw_Complete_Taxa SET [" & q & " & fldName & " & q & "]='" & q & " & Replace(newVal," & q & "'" & q & "," & q & "''" & q & ") & " & q & "'" & q & " & IIf(wc<>"""","" WHERE "" & wc,"""")" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    " & vbCrLf
    s = s & "    ' Backup each record to Taxa_History before updating" & vbCrLf
    s = s & "    Dim rsBulk As DAO.Recordset: Set rsBulk = db.OpenRecordset(sqlSel)" & vbCrLf
    s = s & "    Do While Not rsBulk.EOF: Call BackupFieldForUndo(rsBulk!ID,fldName): rsBulk.MoveNext: Loop: rsBulk.Close" & vbCrLf
    s = s & "    db.Execute sqlUpd,128" & vbCrLf
    s = s & "    Me.subResults.Form.Requery" & vbCrLf
    s = s & "    MsgBox " & q & "Bulk edit done: " & q & " & fldName & " & q & " -> " & q & " & newVal & Chr(10) & " & q & "Scope: " & q & " & scopeLabel,vbInformation," & q & "Bulk Edit" & q & vbCrLf
    s = s & "    Exit Sub" & vbCrLf
    s = s & "BulkErr: MsgBox " & q & "Bulk edit error: " & q & " & Err.Description,vbCritical" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeBulkEdit = s
End Function

' ----------------------------------------------------------------
' D3. SaveFilter / LoadFilter
' ----------------------------------------------------------------
Private Function CodeSaveFilter() As String
    Dim q As String: q = Chr(34)
    Dim s As String
    s = "Private Sub btnSaveFilter_Click()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Dim data As String: data = SerializeFilters()" & vbCrLf
    s = s & "    If data = """" Then MsgBox " & q & "No active filters to save." & q & ",vbExclamation: Exit Sub" & vbCrLf
    s = s & "    TempVars.Item(" & q & "FilterDataToSave" & q & ") = data" & vbCrLf
    s = s & "    TempVars.Item(" & q & "FilterAction" & q & ") = " & q & "cancel" & q & vbCrLf
    s = s & "    DoCmd.OpenForm " & q & "frmSavedFilters" & q & ",acNormal,,,,acDialog" & vbCrLf
    s = s & "    ' If user clicked Load inside the Save dialog, apply it immediately" & vbCrLf
    s = s & "    If Nz(TempVars(" & q & "FilterAction" & q & "),"""") = " & q & "load" & q & " Then btnLoadFilter_Click" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeSaveFilter = s
End Function

Private Function CodeLoadFilter() As String
    Dim q As String: q = Chr(34)
    Dim s As String
    s = "Private Sub btnLoadFilter_Click()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    ' Open dialog only when called directly (not chained from SaveFilter)" & vbCrLf
    s = s & "    If Nz(TempVars(" & q & "FilterAction" & q & "),"""") <> " & q & "load" & q & " Then" & vbCrLf
    s = s & "        TempVars.Item(" & q & "FilterAction" & q & ") = " & q & "cancel" & q & vbCrLf
    s = s & "        TempVars.Item(" & q & "FilterDataToSave" & q & ") = SerializeFilters()" & vbCrLf
    s = s & "        DoCmd.OpenForm " & q & "frmSavedFilters" & q & ",acNormal,,,,acDialog" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "    If Nz(TempVars(" & q & "FilterAction" & q & "),"""") <> " & q & "load" & q & " Then Exit Sub" & vbCrLf
    s = s & "    Dim data As String: data = Nz(TempVars(" & q & "FilterData" & q & "),"""")" & vbCrLf
    s = s & "    Dim nm As String: nm = Nz(TempVars(" & q & "FilterName" & q & "),"""")" & vbCrLf
    s = s & "    If data = """" Then Exit Sub" & vbCrLf
    s = s & "    DeserializeFilters data" & vbCrLf
    s = s & "    ApplyFilters" & vbCrLf
    s = s & "    Me!lblActiveFilter.Caption = " & q & "[Preset: " & q & " & nm & " & q & "] " & q & " & Left(BuildWhereClause(""""),80)" & vbCrLf
    s = s & "    Me!lblStatusDot.ForeColor = &H0044FF" & vbCrLf
    s = s & "    TempVars.Item(" & q & "FilterAction" & q & ") = " & q & "cancel" & q & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeLoadFilter = s
End Function

' ----------------------------------------------------------------
' D4. SerializeFilters / DeserializeFilters
'     Format: ctrlName=value|ctrlName=value|...
'     Only non-empty / non-zero values are stored.
' ----------------------------------------------------------------
Private Function CodeSerialize() As String
    Dim q As String: q = Chr(34)
    Dim s As String
    s = "Public Function SerializeFilters() As String" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Dim result As String: result = """": Dim c As Control" & vbCrLf
    s = s & "    For Each c In Me.Controls" & vbCrLf
    s = s & "        Dim nm As String: nm = c.Name" & vbCrLf
    s = s & "        ' Only process filter controls (txt, cbo, chk prefix)" & vbCrLf
    s = s & "        If Left(nm,3) <> " & q & "txt" & q & " And Left(nm,3) <> " & q & "cbo" & q & " And Left(nm,3) <> " & q & "chk" & q & " Then GoTo NextCtrl" & vbCrLf
    s = s & "        Dim val As String: val = """"" & vbCrLf
    s = s & "        If TypeName(c) = " & q & "TextBox" & q & " Then val = Nz(c.Value,"""")" & vbCrLf
    s = s & "        If TypeName(c) = " & q & "ComboBox" & q & " Then val = Nz(c.Value,"""")" & vbCrLf
    s = s & "        If TypeName(c) = " & q & "CheckBox" & q & " Then val = IIf(Nz(c.Value,False)," & q & "1" & q & "," & q & "0" & q & ")" & vbCrLf
    s = s & "        If val <> """" And val <> " & q & "0" & q & " Then result = result & nm & " & q & "=" & q & " & val & " & q & "|" & q & vbCrLf
    s = s & "        NextCtrl:" & vbCrLf
    s = s & "    Next c" & vbCrLf
    s = s & "    SerializeFilters = result" & vbCrLf
    s = s & "End Function" & vbCrLf
    CodeSerialize = s
End Function

Private Function CodeDeserialize() As String
    Dim q As String: q = Chr(34)
    Dim s As String
    s = "Public Sub DeserializeFilters(data As String)" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    ' Clear all filter controls first" & vbCrLf
    s = s & "    Dim c As Control" & vbCrLf
    s = s & "    For Each c In Me.Controls" & vbCrLf
    s = s & "        If TypeName(c) = " & q & "TextBox" & q & " Then c.Value = """"" & vbCrLf
    s = s & "        If TypeName(c) = " & q & "ComboBox" & q & " Then c.Value = Null" & vbCrLf
    s = s & "        If TypeName(c) = " & q & "CheckBox" & q & " Then c.Value = False" & vbCrLf
    s = s & "    Next c" & vbCrLf
    s = s & "    " & vbCrLf
    s = s & "    ' Apply saved values" & vbCrLf
    s = s & "    Dim parts() As String: parts = Split(data," & q & "|" & q & ")" & vbCrLf
    s = s & "    Dim p As Variant" & vbCrLf
    s = s & "    For Each p In parts" & vbCrLf
    s = s & "        If InStr(p," & q & "=" & q & ")>0 Then" & vbCrLf
    s = s & "            Dim nm As String: nm = Left(p,InStr(p," & q & "=" & q & ")-1)" & vbCrLf
    s = s & "            Dim vl As String: vl = Mid(p,InStr(p," & q & "=" & q & ")+1)" & vbCrLf
    s = s & "            On Error Resume Next" & vbCrLf
    s = s & "            Dim ctl As Control: Set ctl = Me.Controls(nm)" & vbCrLf
    s = s & "            If Err.Number=0 Then" & vbCrLf
    s = s & "                If TypeName(ctl)=" & q & "CheckBox" & q & " Then ctl.Value=(vl=" & q & "1" & q & ")" & vbCrLf
    s = s & "                If TypeName(ctl)=" & q & "ComboBox" & q & " Then ctl.Value=IIf(vl=" & q & q & ",Null,vl)" & vbCrLf
    s = s & "                If TypeName(ctl)=" & q & "TextBox" & q & " Then ctl.Value=vl" & vbCrLf
    s = s & "            End If" & vbCrLf
    s = s & "            On Error GoTo 0" & vbCrLf
    s = s & "        End If" & vbCrLf
    s = s & "    Next p" & vbCrLf
    s = s & "End Sub" & vbCrLf
    CodeDeserialize = s
End Function

' ================================================================
' E. Patch frmCompare4 - btnSaveAll with detailed feedback
'    Before: showed only "Saved." with no information
'    After:  lists taxon names, field count, History info, Undo tip
' ================================================================
Private Sub PatchCompare4SaveAll()
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
    Dim s As String
    s = "Private Sub btnSaveAll_Click()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Dim ci As Integer, fi As Integer" & vbCrLf
    s = s & "    Dim savedTaxa As String: savedTaxa = """"" & vbCrLf
    s = s & "    Dim totalFields As Long: totalFields = 0" & vbCrLf
    s = s & "    Dim fl(39) As String" & vbCrLf
    Dim i As Integer
    For i = 0 To 39
        s = s & "    fl(" & i & ")=" & q & flds(i) & q & vbCrLf
    Next i
    s = s & "    For ci = 1 To 4" & vbCrLf
    s = s & "        If g_IDs(ci) = 0 Then GoTo NextSave" & vbCrLf
    s = s & "        Dim taxNm As String" & vbCrLf
    s = s & "        taxNm = Nz(DLookup(" & q & "TAXON" & q & "," & q & "Taxa" & q & "," & q & "ID=" & q & " & g_IDs(ci))," & q & "ID=" & q & " & g_IDs(ci))" & vbCrLf
    s = s & "        savedTaxa = savedTaxa & Chr(10) & " & q & "  [" & q & " & ci & " & q & "] " & q & " & taxNm" & vbCrLf
    s = s & "        For fi = 0 To 39" & vbCrLf
    s = s & "            Dim cn As String" & vbCrLf
    s = s & "            cn = " & q & "txt_" & q & " & Replace(Replace(fl(fi)," & q & " " & q & "," & q & "_" & q & ")," & q & "-" & q & "," & q & "_" & q & ") & " & q & "_" & q & " & ci" & vbCrLf
    s = s & "            Dim ctl As Control: Set ctl = Me.Controls(cn)" & vbCrLf
    s = s & "            Call BackupFieldForUndo(g_IDs(ci),fl(fi))" & vbCrLf
    s = s & "            Call UpdateFieldInDB(g_IDs(ci),fl(fi),Nz(ctl.Value,""""))" & vbCrLf
    s = s & "            totalFields = totalFields + 1" & vbCrLf
    s = s & "        Next fi" & vbCrLf
    s = s & "        NextSave:" & vbCrLf
    s = s & "    Next ci" & vbCrLf
    s = s & "    If savedTaxa = """" Then" & vbCrLf
    s = s & "        MsgBox " & q & "No taxa loaded - nothing to save." & q & ",vbExclamation" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        MsgBox " & q & "SAVE ALL completed." & q & " & Chr(10) & Chr(10) & _" & vbCrLf
    s = s & "            " & q & "Fields saved: " & q & " & totalFields & Chr(10) & _" & vbCrLf
    s = s & "            " & q & "Taxa updated:" & q & " & savedTaxa & Chr(10) & Chr(10) & _" & vbCrLf
    s = s & "            " & q & "Data stored in Taxa + logged in Taxa_History." & q & " & Chr(10) & _" & vbCrLf
    s = s & "            " & q & "Use UNDO [Ctrl+Z] in the main search to revert." & q & ",vbInformation," & q & "Save All - Done" & q & vbCrLf
    s = s & "        On Error Resume Next: Forms(" & q & "frmDark" & q & ").subResults.Form.Requery" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf

    PatchSub mdl, "btnSaveAll_Click", s

    DoCmd.Close acForm, "frmCompare4", acSaveYes
    Exit Sub
C4Err:
    MsgBox "PatchCompare4SaveAll error: " & Err.Description, vbCritical
    On Error Resume Next: DoCmd.Close acForm, "frmCompare4", acSaveYes
End Sub
