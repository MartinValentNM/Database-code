Attribute VB_Name = "BALIK_11"
Option Compare Database
Option Explicit

' ================================================================
' BALIK_11 - instaluje se po BALIK_10
'
' Po instalaci VZDY spustit: CreateDetailFormTaxa_Ultimate
'
' Co dela:
'   [1] ALTER TABLE Taxa_Poznamky ADD COLUMN NOTES MEMO
'   [2] NOTES pole v TAXONOMY sekci (vyska 2400, unbound):
'         - OnGotFocus  -> LoadNotesField(Form)
'         - Tlacitko [S] vedle pole -> SaveNotesField(Form)
'         - NavPrev / NavNext / CloseForm patchnuty:
'           pred navigaci/zavrenim auto-ulozi NOTES
'         - bNotesOrigVal sleduje puvodni hodnotu
'         - SaveNotesField: UPSERT do Taxa_Poznamky
'         - LoadNotesField: SELECT z Taxa_Poznamky
'         - NOTES do SaveField + BackupFieldForUndo routing
'         - FormCurrent_Update: LoadNotesField pri kazdem zaznamu
'   [3] Bug CARDINAL_PROCESSES: RefreshDetailIfOpen v btnClear_Click
'   [4] Skryti hornich zalozek + levy sidebar
' ================================================================

Public Sub InstallBALIK11()
    Debug.Print "=== BALIK_11 start ==="
    Step1_AddNotesColumn
    Debug.Print "  [1] NOTES sloupec"
    Step2_PatchDetailCode
    Debug.Print "  [2] NOTES pole"
    Step3_FixCardinalBug
    Debug.Print "  [3] CARDINAL bug"
    Step4_PatchSidebarAndTabs
    Debug.Print "  [4] Sidebar"
    MsgBox "BALIK_11 hotovo!" & vbCrLf & vbCrLf & _
           "Nyni SPUST: CreateDetailFormTaxa_Ultimate", _
           vbInformation, "BALIK_11 OK"
End Sub

' ================================================================
' [1] NOTES sloupec
' ================================================================
Private Sub Step1_AddNotesColumn()
    On Error Resume Next
    CurrentDb.Execute "ALTER TABLE Taxa_Poznamky ADD COLUMN NOTES MEMO", 128
    Debug.Print "    " & IIf(Err.Number = 0, "NOTES pridan", "NOTES: " & Err.Description)
    On Error GoTo 0
End Sub

' ================================================================
' [2] Patch 11_code_detail
' ================================================================
Private Sub Step2_PatchDetailCode()
    On Error GoTo Err2
    Dim cm As Object: Set cm = FindModule("CreateDetailFormTaxa_Ultimate")
    If cm Is Nothing Then
        MsgBox "[2] Modul nenalezen - importuj 11_code_detail.bas", vbCritical
        Exit Sub
    End If
    Dim i As Long, ln As String

    ' 2a) Case "X": H = 2400
    For i = 1 To cm.CountOfLines
        If InStr(cm.Lines(i, 1), "Case ""R"": H = 1800") > 0 Then
            If InStr(IIf(i < cm.CountOfLines, cm.Lines(i + 1, 1), ""), "Case ""X""") = 0 Then
                cm.InsertLines i + 1, "            Case ""X"": H = 2400"
                Debug.Print "    Case X pridan"
            Else: Debug.Print "    Case X uz existuje"
            End If
            Exit For
        End If
    Next i

    ' 2b) NOTES\1\1\X do DEF_TAXONOMY
    For i = 1 To cm.CountOfLines
        If InStr(cm.Lines(i, 1), """INCLUDED TAXONS\1\1\N""") > 0 Then
            Dim nxtN As String
            nxtN = IIf(i < cm.CountOfLines, cm.Lines(i + 1, 1), "")
            If InStr(nxtN, "NOTES") = 0 Then
                cm.ReplaceLine i, Replace(cm.Lines(i, 1), _
                    """INCLUDED TAXONS\1\1\N""", _
                    """INCLUDED TAXONS\1\1\N"" & vbLf & _")
                cm.InsertLines i + 1, "    ""NOTES\1\1\X"""
                Debug.Print "    NOTES\1\1\X pridano"
            ElseIf InStr(nxtN, "\1\1\N") > 0 Then
                cm.ReplaceLine i + 1, Replace(nxtN, "\1\1\N", "\1\1\X")
                Debug.Print "    NOTES vyska N->X"
            Else: Debug.Print "    NOTES uz existuje"
            End If
            Exit For
        End If
    Next i

    ' 2c) MakePair: txt_NOTES unbound + tlacitko [S]
    For i = 1 To cm.CountOfLines
        If InStr(cm.Lines(i, 1), "' Hide auto-label created by Access for TextBox") > 0 Then
            Dim prevL As String
            prevL = IIf(i > 1, Trim(cm.Lines(i - 1, 1)), "")
            If InStr(prevL, "NOTES") = 0 And InStr(prevL, "End If") = 0 Then
                cm.InsertLines i, "    ' NOTES: unbound + tlacitko SAVE"
                i = i + 1
                cm.InsertLines i, "    If UCase(fName) = ""NOTES"" Then"
                i = i + 1
                cm.InsertLines i, "        txt.ControlSource = """" "
                i = i + 1
                cm.InsertLines i, "        txt.ScrollBars = 2"
                i = i + 1
                cm.InsertLines i, "        txt.EnterKeyBehavior = True"
                i = i + 1
                cm.InsertLines i, "        txt.OnGotFocus = ""=LoadNotesField(Form)"""
                i = i + 1
                cm.InsertLines i, "        Dim btnSN As Control"
                i = i + 1
                cm.InsertLines i, "        Set btnSN = CreateControl(frm.Name, acCommandButton, acDetail, pg.Name, """", X + lw + GRID_GAP + txtW + 60, Y, 700, H)"
                i = i + 1
                cm.InsertLines i, "        btnSN.Name = ""btnSaveNotes"": btnSN.Caption = ""S"""
                i = i + 1
                cm.InsertLines i, "        btnSN.OnClick = ""=SaveNotesField(Form)"""
                i = i + 1
                cm.InsertLines i, "        btnSN.BackColor = &H005580: btnSN.ForeColor = &HFFFFFF"
                i = i + 1
                cm.InsertLines i, "        btnSN.FontSize = 9: btnSN.FontBold = True"
                i = i + 1
                cm.InsertLines i, "    End If"
                Debug.Print "    MakePair NOTES patch pridan"
            Else: Debug.Print "    MakePair NOTES patch uz existuje"
            End If
            Exit For
        End If
    Next i

    ' 2d) NOTES do SaveField + BackupFieldForUndo routing
    For i = 1 To cm.CountOfLines
        ln = cm.Lines(i, 1)
        If InStr(ln, "Case ""REMARKS"", ""REFERENCE""") > 0 Then
            If InStr(ln, "NOTES") = 0 Then
                cm.ReplaceLine i, Replace(ln, _
                    "Case ""REMARKS"", ""REFERENCE""", _
                    "Case ""REMARKS"", ""REFERENCE"", ""NOTES""")
                Debug.Print "    Routing NOTES pridan (radek " & i & ")"
            End If
        End If
    Next i

    ' 2e) bNotesOrigVal do state promennych
    Dim hasOV As Boolean: hasOV = False
    For i = 1 To cm.CountOfLines
        If InStr(cm.Lines(i, 1), "bNotesOrigVal") > 0 Then hasOV = True: Exit For
    Next i
    If Not hasOV Then
        For i = 1 To cm.CountOfLines
            If InStr(cm.Lines(i, 1), "bInCurrentUpdate") > 0 And _
               InStr(cm.Lines(i, 1), "Private") > 0 Then
                cm.InsertLines i + 1, "Private bNotesOrigVal    As String"
                Debug.Print "    bNotesOrigVal pridan"
                Exit For
            End If
        Next i
    Else: Debug.Print "    bNotesOrigVal uz existuje"
    End If

    ' 2f) FormCurrent_Update: LoadNotesField pri kazdem zaznamu
    For i = 1 To cm.CountOfLines
        If InStr(cm.Lines(i, 1), "bInCurrentUpdate = False") > 0 Then
            Dim ctx As String: ctx = ""
            Dim ci As Long
            For ci = i - 1 To 1 Step -1
                If InStr(cm.Lines(ci, 1), "Function FormCurrent_Update") > 0 Then ctx = "ok": Exit For
                If InStr(cm.Lines(ci, 1), "End Function") > 0 Then Exit For
            Next ci
            If ctx = "ok" Then
                Dim hasLd As Boolean: hasLd = False
                Dim li As Long
                For li = i - 1 To i - 6 Step -1
                    If li < 1 Then Exit For
                    If InStr(cm.Lines(li, 1), "LoadNotesField") > 0 Then hasLd = True: Exit For
                Next li
                If Not hasLd Then
                    cm.InsertLines i, "    Call LoadNotesField(frm)"
                    Debug.Print "    FormCurrent_Update: LoadNotesField pridano"
                Else: Debug.Print "    FormCurrent_Update: uz existuje"
                End If
                Exit For
            End If
        End If
    Next i

    ' 2g) NavPrev: SaveNotesField pred navigaci
    Call PatchNavFn(cm, "Function NavPrev", "DoCmd.GoToRecord , , acPrevious")
    ' 2h) NavNext: SaveNotesField pred navigaci
    Call PatchNavFn(cm, "Function NavNext", "DoCmd.GoToRecord , , acNext")
    ' 2i) CloseForm: SaveNotesField pred zavrenim
    Call PatchNavFn(cm, "Function CloseForm", "DoCmd.Close acForm")

    ' 2j) LoadNotesField()
    Dim hasLNF As Boolean: hasLNF = False
    For i = 1 To cm.CountOfLines
        If InStr(cm.Lines(i, 1), "Function LoadNotesField") > 0 Then hasLNF = True: Exit For
    Next i
    If Not hasLNF Then
        Dim q As String: q = Chr(34)
        Dim lnf As String
        lnf = ""
        lnf = lnf & "' LoadNotesField - nacte NOTES, ulozi original" & vbCrLf
        lnf = lnf & "Public Function LoadNotesField(frm As Form) As Boolean" & vbCrLf
        lnf = lnf & "    On Error Resume Next" & vbCrLf
        lnf = lnf & "    Dim tID As Long: tID = Nz(frm!txtID, 0)" & vbCrLf
        lnf = lnf & "    If tID = 0 Then LoadNotesField = False: Exit Function" & vbCrLf
        lnf = lnf & "    Dim rs As DAO.Recordset" & vbCrLf
        lnf = lnf & "    Set rs = CurrentDb.OpenRecordset(" & q & "SELECT NOTES FROM Taxa_Poznamky WHERE TaxonID=" & q & " & tID)" & vbCrLf
        lnf = lnf & "    Dim val As String: val = " & q & q & vbCrLf
        lnf = lnf & "    If Not rs.EOF Then val = Nz(rs!NOTES, " & q & q & ")" & vbCrLf
        lnf = lnf & "    rs.Close" & vbCrLf
        lnf = lnf & "    frm.Controls(" & q & "txt_NOTES" & q & ").Value = val" & vbCrLf
        lnf = lnf & "    bNotesOrigVal = val" & vbCrLf
        lnf = lnf & "    LoadNotesField = True" & vbCrLf
        lnf = lnf & "End Function"
        cm.InsertLines cm.CountOfLines + 1, lnf
        Debug.Print "    LoadNotesField pridana"
    Else: Debug.Print "    LoadNotesField uz existuje"
    End If

    ' 2k) SaveNotesField() - UPSERT, ulozi vzdy kdyz se lisi od originu
    Dim hasSNF As Boolean: hasSNF = False
    For i = 1 To cm.CountOfLines
        If InStr(cm.Lines(i, 1), "Function SaveNotesField") > 0 Then hasSNF = True: Exit For
    Next i
    If Not hasSNF Then
        Dim q2 As String: q2 = Chr(34)
        Dim snf As String
        snf = ""
        snf = snf & "' SaveNotesField - UPSERT NOTES, vola se z [S] tlacitka + nav funkci" & vbCrLf
        snf = snf & "Public Function SaveNotesField(frm As Form) As Boolean" & vbCrLf
        snf = snf & "    On Error Resume Next" & vbCrLf
        snf = snf & "    Dim tID As Long: tID = Nz(frm!txtID, 0)" & vbCrLf
        snf = snf & "    If tID = 0 Then SaveNotesField = False: Exit Function" & vbCrLf
        snf = snf & "    Dim newVal As String" & vbCrLf
        snf = snf & "    newVal = Nz(frm.Controls(" & q2 & "txt_NOTES" & q2 & ").Value, " & q2 & q2 & ")" & vbCrLf
        snf = snf & "    If newVal = bNotesOrigVal Then SaveNotesField = True: Exit Function" & vbCrLf
        snf = snf & "    Dim db As DAO.Database: Set db = CurrentDb" & vbCrLf
        snf = snf & "    Dim sv As String: sv = Replace(newVal, " & q2 & "'" & q2 & ", " & q2 & "''" & q2 & ")" & vbCrLf
        snf = snf & "    db.Execute " & q2 & "UPDATE Taxa_Poznamky SET [NOTES]='" & q2 & " & sv & " & q2 & "' WHERE TaxonID=" & q2 & " & tID, 128" & vbCrLf
        snf = snf & "    If db.RecordsAffected = 0 Then" & vbCrLf
        snf = snf & "        db.Execute " & q2 & "INSERT INTO Taxa_Poznamky (TaxonID,NOTES) VALUES (" & q2 & " & tID & " & q2 & ",'" & q2 & " & sv & " & q2 & "')" & q2 & ", 128" & vbCrLf
        snf = snf & "    End If" & vbCrLf
        snf = snf & "    bNotesOrigVal = newVal" & vbCrLf
        snf = snf & "    SaveNotesField = True" & vbCrLf
        snf = snf & "End Function"
        cm.InsertLines cm.CountOfLines + 1, snf
        Debug.Print "    SaveNotesField pridana"
    Else: Debug.Print "    SaveNotesField uz existuje"
    End If

    Exit Sub
Err2: MsgBox "[2] Chyba: " & Err.Description, vbCritical
End Sub

' ================================================================
' Helper: vlozi SaveNotesField na zacatek NavPrev/NavNext/CloseForm
' ================================================================
Private Sub PatchNavFn(cm As Object, fnSig As String, firstLine As String)
    On Error Resume Next
    Dim i As Long, sL As Long: sL = 0
    For i = 1 To cm.CountOfLines
        If InStr(cm.Lines(i, 1), fnSig) > 0 Then sL = i: Exit For
    Next i
    If sL = 0 Then Debug.Print "    " & fnSig & " nenalezena": Exit Sub
    ' Zkontroluj jestli uz SaveNotesField neni uvnitr
    Dim eL As Long: eL = 0
    For i = sL + 1 To cm.CountOfLines
        If InStr(cm.Lines(i, 1), "End Function") > 0 Then eL = i: Exit For
    Next i
    If eL = 0 Then Exit Sub
    Dim k As Long, hasIt As Boolean: hasIt = False
    For k = sL To eL
        If InStr(cm.Lines(k, 1), "SaveNotesField") > 0 Then hasIt = True: Exit For
    Next k
    If hasIt Then
        Debug.Print "    " & fnSig & ": SaveNotesField uz existuje"
        Exit Sub
    End If
    ' Vloz za "On Error Resume Next" nebo za prvni radek funkce
    Dim insertAt As Long: insertAt = sL + 1
    For i = sL + 1 To eL
        If InStr(cm.Lines(i, 1), "On Error Resume Next") > 0 Then
            insertAt = i + 1: Exit For
        End If
    Next i
    cm.InsertLines insertAt, "    On Error Resume Next: Call SaveNotesField(Screen.ActiveForm): On Error Resume Next"
    Debug.Print "    " & fnSig & ": SaveNotesField pridano"
End Sub

' ================================================================
' [3] Bug CARDINAL_PROCESSES
' ================================================================
Private Sub Step3_FixCardinalBug()
    On Error GoTo Err3
    DoCmd.OpenForm "frmDark", acDesign
    Dim frm As Form: Set frm = Forms("frmDark")
    Dim mdl As Module: Set mdl = frm.Module
    Dim i As Long

    Dim hasRDIO As Boolean: hasRDIO = False
    For i = 1 To mdl.CountOfLines
        If InStr(mdl.Lines(i, 1), "Sub RefreshDetailIfOpen") > 0 Then hasRDIO = True: Exit For
    Next i
    If Not hasRDIO Then
        Dim hf As String
        hf = ""
        hf = hf & vbCrLf
        hf = hf & "Private Sub RefreshDetailIfOpen()" & vbCrLf
        hf = hf & "    On Error Resume Next" & vbCrLf
        hf = hf & "    Dim df As Form: Set df = Forms(""frmDetailTaxa"")" & vbCrLf
        hf = hf & "    If Not df Is Nothing Then df.Requery" & vbCrLf
        hf = hf & "    On Error GoTo 0" & vbCrLf
        hf = hf & "End Sub"
        mdl.InsertLines mdl.CountOfLines + 1, hf
        Debug.Print "    RefreshDetailIfOpen pridana"
    Else: Debug.Print "    RefreshDetailIfOpen uz existuje"
    End If

    Dim sL As Long, eL As Long: sL = 0: eL = 0
    For i = 1 To mdl.CountOfLines
        If InStr(mdl.Lines(i, 1), "Sub btnClear_Click") > 0 Then sL = i
        If sL > 0 And i > sL Then
            If Trim(mdl.Lines(i, 1)) = "End Sub" Then eL = i: Exit For
        End If
    Next i
    If sL > 0 And eL > 0 Then
        Dim hasR As Boolean: hasR = False
        Dim k As Long
        For k = sL To eL
            If InStr(mdl.Lines(k, 1), "RefreshDetailIfOpen") > 0 Then hasR = True: Exit For
        Next k
        If Not hasR Then
            mdl.InsertLines eL, "    Call RefreshDetailIfOpen"
            Debug.Print "    RefreshDetailIfOpen -> btnClear_Click"
        Else: Debug.Print "    btnClear_Click: uz volano"
        End If
    Else: Debug.Print "    btnClear_Click nenalezen"
    End If

    DoCmd.Close acForm, "frmDark", acSaveYes
    Exit Sub
Err3: MsgBox "[3] Chyba: " & Err.Description, vbCritical
    On Error Resume Next: DoCmd.Close acForm, "frmDark", acSaveYes
End Sub

' ================================================================
' [4] Sidebar + skryti zalozek
' ================================================================
Private Sub Step4_PatchSidebarAndTabs()
    On Error GoTo Err4
    Dim cm As Object: Set cm = FindModule("CreateDetailFormTaxa_Ultimate")
    If cm Is Nothing Then MsgBox "[4] Modul nenalezen.", vbCritical: Exit Sub
    Dim i As Long, ln As String

    ' a) tc.Style = 2
    Dim inMain As Boolean: inMain = False
    For i = 1 To cm.CountOfLines
        ln = cm.Lines(i, 1)
        If InStr(ln, "Sub CreateDetailFormTaxa_Ultimate") > 0 Then inMain = True
        If inMain And InStr(ln, "End Sub") > 0 And i > 5 Then inMain = False
        If inMain And InStr(ln, "tc.BackStyle = 0") > 0 Then
            If InStr(IIf(i < cm.CountOfLines, Trim(cm.Lines(i + 1, 1)), ""), "tc.Style = 2") = 0 Then
                cm.InsertLines i + 1, "    tc.Style = 2  ' Skryj horni zalozky"
                Debug.Print "    tc.Style=2 pridan"
            Else: Debug.Print "    tc.Style=2 uz existuje"
            End If
            Exit For
        End If
    Next i

    ' b) BuildSidebarNav volani
    For i = 1 To cm.CountOfLines
        If InStr(cm.Lines(i, 1), "BuildRelationshipsTab tmp, tc, 4") > 0 Then
            If InStr(IIf(i < cm.CountOfLines, Trim(cm.Lines(i + 1, 1)), ""), "BuildSidebarNav") = 0 Then
                cm.InsertLines i + 1, "    BuildSidebarNav tmp, tc"
                Debug.Print "    BuildSidebarNav volani pridano"
            Else: Debug.Print "    BuildSidebarNav uz volano"
            End If
            Exit For
        End If
    Next i

    ' c) BuildSidebarNav Sub
    Dim hasSB As Boolean: hasSB = False
    For i = 1 To cm.CountOfLines
        If InStr(cm.Lines(i, 1), "Sub BuildSidebarNav") > 0 Then hasSB = True: Exit For
    Next i
    If Not hasSB Then
        Dim ins As Long: ins = 0
        For i = 1 To cm.CountOfLines
            If InStr(cm.Lines(i, 1), "Private Sub BuildRelationshipsTab") > 0 Then ins = i: Exit For
        Next i
        If ins = 0 Then ins = cm.CountOfLines + 1
        Dim q As String: q = Chr(34)
        Dim sc As String
        sc = ""
        sc = sc & "' ===== SIDEBAR NAV =====" & vbCrLf
        sc = sc & "Private Sub BuildSidebarNav(frm As Form, tc As Control)" & vbCrLf
        sc = sc & "    Const SB_W  As Long = 1800" & vbCrLf
        sc = sc & "    Const BTN_H As Long = 800" & vbCrLf
        sc = sc & "    Const BTN_G As Long = 60" & vbCrLf
        sc = sc & "    Const BTN_W As Long = 1780" & vbCrLf
        sc = sc & "    Const START_Y As Long = 200" & vbCrLf
        sc = sc & "    Dim bg As Control" & vbCrLf
        sc = sc & "    Set bg = CreateControl(frm.Name, acRectangle, acDetail, " & q & q & ", " & q & q & ", 0, 0, SB_W, TAB_H)" & vbCrLf
        sc = sc & "    bg.BackColor = &H1A1A1A: bg.BackStyle = 1: bg.BorderStyle = 0: bg.Name = " & q & "sbBg" & q & vbCrLf
        sc = sc & "    Dim caps(4) As String" & vbCrLf
        sc = sc & "    caps(0) = " & q & "TAXONOMY" & q & vbCrLf
        sc = sc & "    caps(1) = " & q & "DESCRIPTION" & q & vbCrLf
        sc = sc & "    caps(2) = " & q & "OCCURRENCE" & q & vbCrLf
        sc = sc & "    caps(3) = " & q & "MATERIAL" & q & vbCrLf
        sc = sc & "    caps(4) = " & q & "REL-SHIPS" & q & vbCrLf
        sc = sc & "    Dim bi As Integer, bY As Long: bY = START_Y" & vbCrLf
        sc = sc & "    For bi = 0 To 4" & vbCrLf
        sc = sc & "        Dim btn As Control" & vbCrLf
        sc = sc & "        Set btn = CreateControl(frm.Name, acCommandButton, acDetail, " & q & q & ", " & q & q & ", 10, bY, BTN_W, BTN_H)" & vbCrLf
        sc = sc & "        btn.Name = " & q & "sbBtn" & q & " & bi" & vbCrLf
        sc = sc & "        btn.Caption = caps(bi)" & vbCrLf
        sc = sc & "        btn.OnClick = " & q & "=NavToSection(" & q & " & bi & " & q & ")" & q & vbCrLf
        sc = sc & "        btn.FontSize = 8: btn.FontBold = True: btn.ForeColor = &HFFFFFF" & vbCrLf
        sc = sc & "        If bi = 0 Then" & vbCrLf
        sc = sc & "            btn.BackColor = &H005580" & vbCrLf
        sc = sc & "        Else" & vbCrLf
        sc = sc & "            btn.BackColor = &H2B2B2B" & vbCrLf
        sc = sc & "        End If" & vbCrLf
        sc = sc & "        bY = bY + BTN_H + BTN_G" & vbCrLf
        sc = sc & "    Next bi" & vbCrLf
        sc = sc & "End Sub" & vbCrLf
        cm.InsertLines ins, sc
        Debug.Print "    BuildSidebarNav pridana"
    Else: Debug.Print "    BuildSidebarNav uz existuje"
    End If

    ' d) NavToSection()
    Dim hasNav As Boolean: hasNav = False
    For i = 1 To cm.CountOfLines
        If InStr(cm.Lines(i, 1), "Function NavToSection") > 0 Then hasNav = True: Exit For
    Next i
    If Not hasNav Then
        Dim q2 As String: q2 = Chr(34)
        Dim nc As String
        nc = ""
        nc = nc & "Public Function NavToSection(idx As Integer) As Boolean" & vbCrLf
        nc = nc & "    On Error Resume Next" & vbCrLf
        nc = nc & "    Dim frm As Form: Set frm = Screen.ActiveForm" & vbCrLf
        nc = nc & "    If frm Is Nothing Then NavToSection = False: Exit Function" & vbCrLf
        nc = nc & "    frm.Controls(" & q2 & "tabMain" & q2 & ").Value = idx" & vbCrLf
        nc = nc & "    Dim bi As Integer" & vbCrLf
        nc = nc & "    For bi = 0 To 4" & vbCrLf
        nc = nc & "        On Error Resume Next" & vbCrLf
        nc = nc & "        Dim nb As Control: Set nb = frm.Controls(" & q2 & "sbBtn" & q2 & " & bi)" & vbCrLf
        nc = nc & "        If Not nb Is Nothing Then" & vbCrLf
        nc = nc & "            If bi = idx Then" & vbCrLf
        nc = nc & "                nb.BackColor = &H005580" & vbCrLf
        nc = nc & "            Else" & vbCrLf
        nc = nc & "                nb.BackColor = &H2B2B2B" & vbCrLf
        nc = nc & "            End If" & vbCrLf
        nc = nc & "        End If" & vbCrLf
        nc = nc & "    Next bi" & vbCrLf
        nc = nc & "    NavToSection = True" & vbCrLf
        nc = nc & "End Function"
        cm.InsertLines cm.CountOfLines + 1, nc
        Debug.Print "    NavToSection pridana"
    Else: Debug.Print "    NavToSection uz existuje"
    End If

    ' e) GRID_START_X 400 -> 2000
    For i = 1 To cm.CountOfLines
        ln = cm.Lines(i, 1)
        If InStr(ln, "GRID_START_X") > 0 And InStr(ln, "Const") > 0 Then
            If InStr(ln, "= 400") > 0 Then
                cm.ReplaceLine i, Replace(ln, "= 400", "= 2000  ' Sidebar offset")
                Debug.Print "    GRID_START_X 400->2000"
            ElseIf InStr(ln, "= 2000") > 0 Then
                Debug.Print "    GRID_START_X uz 2000"
            Else: Debug.Print "    GRID_START_X: zkontroluj rucne"
            End If
            Exit For
        End If
    Next i

    Exit Sub
Err4: MsgBox "[4] Chyba: " & Err.Description, vbCritical
End Sub

' ================================================================
' Helper: najdi CodeModule podle nazvu funkce
' ================================================================
Private Function FindModule(fnName As String) As Object
    Dim comp As Object
    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        On Error Resume Next
        Dim j As Long
        For j = 1 To comp.CodeModule.CountOfLines
            If InStr(comp.CodeModule.Lines(j, 1), fnName) > 0 And _
               InStr(comp.CodeModule.Lines(j, 1), "Sub ") > 0 Then
                Set FindModule = comp.CodeModule
                Exit Function
            End If
        Next j
        On Error GoTo 0
    Next comp
    Set FindModule = Nothing
End Function
