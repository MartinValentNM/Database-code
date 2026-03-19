Attribute VB_Name = "BALIK_10"
Option Compare Database
Option Explicit

' ================================================================
' BALIK_10
' Instaluje se po BALIK_9. Dela presne toto:
'
'   [1] Smaze mrtvy btnCompare_Click z modulu frmDark
'   [2] Prida 4 tlacitka do frmDark:
'         Row B: IMPORT XLSX, BIBTEX
'         Row C: DIFF VIEWER (Left=8000), PRINT REPORT (Left=10000)
'         Pozice DIFF VIEWER a PRINT REPORT jsou ZA SAVE/LOAD FILTER
'   [3] Prida Ctrl+B a Ctrl+P do Form_KeyDown
'   [4] Opravi DefaultValue v frmImportXlsx (#NAZEV? problem)
'         Pokud formular neexistuje, krok se preskoci s upozornenim.
'   [5] Opravi Overflow v CreateTaxaReport_Ultimate
'         Section height 40000 -> 31000 (Access limit je 31680 twips)
'
' INSTALACE:
'   1. File -> Import File -> BALIK_10.bas
'   2. Immediate Window: InstallBALIK10
'   3. Ctrl+S
' ================================================================

Public Sub InstallBALIK10()
    Debug.Print "=== BALIK_10 start ==="

    Step1_RemoveDeadCompare
    Debug.Print "  [1] btnCompare_Click odstranen"

    Step2_AddButtons
    Debug.Print "  [2] Tlacitka pridana"

    Step3_UpdateShortcuts
    Debug.Print "  [3] Zkratky aktualizovany"

    Step4_FixImportXlsx
    Debug.Print "  [4] frmImportXlsx DefaultValue"

    Step5_FixReportOverflow
    Debug.Print "  [5] Report overflow"

    Step6_RemovePrintReport
    Debug.Print "  [6] Print Report odstranen"

    MsgBox "BALIK_10 hotovo!" & vbCrLf & vbCrLf & _
           "[1] btnCompare_Click odstranen" & vbCrLf & _
           "[2] IMPORT XLSX, BIBTEX, DIFF VIEWER, PRINT REPORT" & vbCrLf & _
           "[3] Ctrl+B = BibTeX, Ctrl+P = Print Report" & vbCrLf & _
           "[4] frmImportXlsx DefaultValue opraven" & vbCrLf & _
           "[5] Report Section height opraven (31000)" & vbCrLf & _
           "[6] Print Report odstranen z frmDark", _
           vbInformation, "BALIK_10 OK"
End Sub

' ================================================================
' [1] Smaz mrtvy btnCompare_Click
'     (stary dvou-sloupcovy compare, nahrazen btnCompare4_Click)
' ================================================================
Private Sub Step1_RemoveDeadCompare()
    On Error GoTo Err1
    DoCmd.OpenForm "frmDark", acDesign
    Dim frm As Form: Set frm = Forms("frmDark")
    Dim mdl As Module: Set mdl = frm.Module

    Dim i As Long, sL As Long, eL As Long
    sL = 0: eL = 0
    For i = 1 To mdl.CountOfLines
        Dim ln As String: ln = mdl.Lines(i, 1)
        ' Hledame presne "btnCompare_Click" - NE "btnCompare4_Click"
        If InStr(ln, "Sub btnCompare_Click") > 0 Then sL = i
        If sL > 0 And i > sL Then
            If Trim(ln) = "End Sub" Then eL = i: Exit For
        End If
    Next i

    If sL > 0 And eL > 0 Then
        mdl.DeleteLines sL, eL - sL + 1
        Debug.Print "    Smazan na radcich " & sL & "-" & eL
    Else
        Debug.Print "    Nenalezen (ok)"
    End If

    DoCmd.Close acForm, "frmDark", acSaveYes
    Exit Sub
Err1: MsgBox "[1] Chyba: " & Err.Description, vbCritical
    On Error Resume Next: DoCmd.Close acForm, "frmDark", acSaveYes
End Sub

' ================================================================
' [2] Pridej 4 tlacitka do frmDark
'
' Pouzita barevna schema Package 1:
'   Oranzova #884400 = import/nastroje (IMPORT XLSX, DIFF VIEWER, PRINT REPORT)
'   Zelena   #228844 = export/analyza (BIBTEX)
'
' Row B pozice (leftCol=300):
'   IMPORT CSV:  leftCol+8600=8900,  w=1700, ends 10600
'   BATCH DOCX:  leftCol+10400=10700, w=1800, ends 12500
'   IMPORT XLSX: 12600, w=1800, ends 14400   (novy)
'   BIBTEX:      14500, w=1500, ends 16000   (novy)
'
' Row C pozice (absolute, leftCol=300):
'   HISTORY:     300,   w=1700, ends 2000
'   VALIDATE:    2200,  w=1500, ends 3700
'   SAVE FILTER: 3800,  w=2000, ends 5800   (BALIK_6_7)
'   LOAD FILTER: 5900,  w=2000, ends 7900   (BALIK_6_7)
'   DIFF VIEWER: 8000,  w=1900, ends 9900   (novy - ZA LOAD FILTER)
'   PRINT REPORT:10000, w=1900, ends 11900  (novy - ZA DIFF VIEWER)
'   BACKUP DB:   12400, w=1700, ends 14100
' ================================================================
Private Sub Step2_AddButtons()
    On Error GoTo Err2
    DoCmd.OpenForm "frmDark", acDesign
    Dim frm As Form: Set frm = Forms("frmDark")
    Dim mdl As Module: Set mdl = frm.Module
    Dim fn As String: fn = frm.Name
    Dim q As String: q = Chr(34)

    ' Zjisti Y pozice radku B a C
    Dim rowBY As Long, rowCY As Long
    On Error Resume Next
    rowBY = frm.Controls("btnImportCSV").Top
    rowCY = frm.Controls("btnHistory").Top
    On Error GoTo Err2

    ' ── IMPORT XLSX (Row B, Left=12600) ───────────────────────────
    If Not Exists(frm, "btnImportXlsx") Then
        Dim b1 As Control
        Set b1 = CreateControl(fn, acCommandButton, acDetail, "", "", 12600, rowBY, 1800, 400)
        b1.Name = "btnImportXlsx": b1.Caption = "IMPORT XLSX"
        b1.BackColor = &H884400: b1.ForeColor = &HFFFFFF
        b1.BorderColor = &HBB6600: b1.BorderWidth = 2
        b1.FontSize = 9: b1.FontBold = True
        b1.OnClick = "[Event Procedure]"
        ' Handler uz existuje z Package 1 - nepridavame znovu
        Debug.Print "    IMPORT XLSX pridan (Left=12600)"
    Else
        Debug.Print "    IMPORT XLSX uz existuje"
    End If

    ' ── BIBTEX (Row B, Left=14500) ────────────────────────────────
    If Not Exists(frm, "btnBibTeX") Then
        Dim b2 As Control
        Set b2 = CreateControl(fn, acCommandButton, acDetail, "", "", 14500, rowBY, 1500, 400)
        b2.Name = "btnBibTeX": b2.Caption = "BIBTEX"
        b2.BackColor = &H228844: b2.ForeColor = &HFFFFFF
        b2.BorderColor = &H33AA55: b2.BorderWidth = 2
        b2.FontSize = 9: b2.FontBold = True
        b2.OnClick = "[Event Procedure]"
        ' Handler uz existuje z Package 1 - nepridavame znovu
        Debug.Print "    BIBTEX pridan (Left=14500)"
    Else
        Debug.Print "    BIBTEX uz existuje"
    End If

    ' ── DIFF VIEWER (Row C, Left=8000) ────────────────────────────
    ' Pokud jiz existuje, presun na spravnou pozici
    On Error Resume Next
    Dim bDiff As Control: Set bDiff = frm.Controls("btnDiffViewer")
    On Error GoTo Err2
    If Not bDiff Is Nothing Then
        bDiff.Left = 8000: bDiff.Top = rowCY
        Debug.Print "    DIFF VIEWER presunut na Left=8000"
    Else
        Dim b3 As Control
        Set b3 = CreateControl(fn, acCommandButton, acDetail, "", "", 8000, rowCY, 1900, 400)
        b3.Name = "btnDiffViewer": b3.Caption = "DIFF VIEWER"
        b3.BackColor = &H884400: b3.ForeColor = &HFFFFFF
        b3.BorderColor = &HBB6600: b3.BorderWidth = 2
        b3.FontSize = 9: b3.FontBold = True
        b3.OnClick = "[Event Procedure]"
        ' Handler uz existuje z Package 1 - nepridavame znovu
        Debug.Print "    DIFF VIEWER pridan (Left=8000)"
    End If

    ' ── PRINT REPORT (Row C, Left=10000) ──────────────────────────
    ' Pokud jiz existuje, presun na spravnou pozici
    On Error Resume Next
    Dim bPrint As Control: Set bPrint = frm.Controls("btnPrintReport")
    On Error GoTo Err2
    If Not bPrint Is Nothing Then
        bPrint.Left = 10000: bPrint.Top = rowCY
        Debug.Print "    PRINT REPORT presunut na Left=10000"
    Else
        Dim b4 As Control
        Set b4 = CreateControl(fn, acCommandButton, acDetail, "", "", 10000, rowCY, 1900, 400)
        b4.Name = "btnPrintReport": b4.Caption = "PRINT REPORT"
        b4.BackColor = &H884400: b4.ForeColor = &HFFFFFF
        b4.BorderColor = &HBB6600: b4.BorderWidth = 2
        b4.FontSize = 9: b4.FontBold = True
        b4.OnClick = "[Event Procedure]"
        ' Handler je novy - Package 1 ho nema
        Dim hP As String
        hP = "Private Sub btnPrintReport_Click()" & vbCrLf
        hP = hP & "    On Error GoTo PrintErr" & vbCrLf
        hP = hP & "    Dim ok As Boolean: ok = False" & vbCrLf
        hP = hP & "    Dim o As AccessObject" & vbCrLf
        hP = hP & "    For Each o In CurrentProject.AllReports" & vbCrLf
        hP = hP & "        If o.Name = " & q & "rptTaxaDetail" & q & " Then ok = True: Exit For" & vbCrLf
        hP = hP & "    Next o" & vbCrLf
        hP = hP & "    If Not ok Then" & vbCrLf
        hP = hP & "        MsgBox " & q & "Report rptTaxaDetail neexistuje." & q & " & vbCrLf & _" & vbCrLf
        hP = hP & "               " & q & "Spus nejprve Package 10: BuildForms_5_PreviewReport." & q & ", vbExclamation" & vbCrLf
        hP = hP & "        Exit Sub" & vbCrLf
        hP = hP & "    End If" & vbCrLf
        hP = hP & "    Dim tID As Long: tID = 0" & vbCrLf
        hP = hP & "    On Error Resume Next" & vbCrLf
        hP = hP & "    tID = CLng(Nz(Me.subResults.Form.Recordset.Fields(" & q & "ID" & q & ").Value, 0))" & vbCrLf
        hP = hP & "    On Error GoTo PrintErr" & vbCrLf
        hP = hP & "    If tID = 0 Then MsgBox " & q & "Vyberte zaznam." & q & ", vbExclamation: Exit Sub" & vbCrLf
        hP = hP & "    DoCmd.OpenReport " & q & "rptTaxaDetail" & q & ", acViewPreview, , " & q & "ID=" & q & " & tID" & vbCrLf
        hP = hP & "    Exit Sub" & vbCrLf
        hP = hP & "PrintErr: MsgBox " & q & "Chyba: " & q & " & Err.Description, vbCritical" & vbCrLf
        hP = hP & "End Sub" & vbCrLf
        mdl.InsertLines mdl.CountOfLines + 1, hP
        Debug.Print "    PRINT REPORT pridan (Left=10000, vcetne handleru)"
    End If

    DoCmd.Close acForm, "frmDark", acSaveYes
    Exit Sub
Err2: MsgBox "[2] Chyba: " & Err.Description, vbCritical
    On Error Resume Next: DoCmd.Close acForm, "frmDark", acSaveYes
End Sub

' ================================================================
' [3] Aktualizuj Form_KeyDown - pridej Ctrl+B a Ctrl+P
' ================================================================
Private Sub Step3_UpdateShortcuts()
    On Error GoTo Err3
    DoCmd.OpenForm "frmDark", acDesign
    Dim frm As Form: Set frm = Forms("frmDark")
    Dim mdl As Module: Set mdl = frm.Module

    Dim i As Long, sL As Long, eL As Long: sL = 0: eL = 0
    For i = 1 To mdl.CountOfLines
        Dim ln As String: ln = mdl.Lines(i, 1)
        If InStr(ln, "Form_KeyDown") > 0 And InStr(ln, "Sub ") > 0 Then sL = i
        If sL > 0 And i > sL Then
            If Trim(ln) = "End Sub" Then eL = i: Exit For
        End If
    Next i
    If sL = 0 Then
        Debug.Print "    Form_KeyDown nenalezen"
        GoTo Close3
    End If

    mdl.DeleteLines sL, eL - sL + 1

    Dim s As String
    s = "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Dim ctrl As Boolean: ctrl = (Shift And acCtrlMask) > 0" & vbCrLf
    s = s & "    ' F4 = Quick Preview" & vbCrLf
    s = s & "    If KeyCode = 115 And Shift = 0 Then Call OpenQuickPreview: KeyCode = 0: Exit Sub" & vbCrLf
    s = s & "    ' F5 = Filter/Search" & vbCrLf
    s = s & "    If KeyCode = 116 And Shift = 0 Then Call ApplyFilters: KeyCode = 0: Exit Sub" & vbCrLf
    s = s & "    ' Esc = Clear" & vbCrLf
    s = s & "    If KeyCode = 27 And Shift = 0 Then Call btnClear_Click: KeyCode = 0: Exit Sub" & vbCrLf
    s = s & "    ' Ctrl+N = New Taxon" & vbCrLf
    s = s & "    If KeyCode = 78 And ctrl Then Call btnNewTaxon_Click: KeyCode = 0: Exit Sub" & vbCrLf
    s = s & "    ' Ctrl+Z = Undo" & vbCrLf
    s = s & "    If KeyCode = 90 And ctrl Then Call btnUndoAction_Click: KeyCode = 0: Exit Sub" & vbCrLf
    s = s & "    ' Ctrl+D = Open Detail" & vbCrLf
    s = s & "    If KeyCode = 68 And ctrl Then Call btnDetail_Click: KeyCode = 0: Exit Sub" & vbCrLf
    s = s & "    ' Ctrl+E = Export XLSX" & vbCrLf
    s = s & "    If KeyCode = 69 And ctrl Then Call btnExport_Click: KeyCode = 0: Exit Sub" & vbCrLf
    s = s & "    ' Ctrl+B = BibTeX" & vbCrLf
    s = s & "    If KeyCode = 66 And ctrl Then Call btnBibTeX_Click: KeyCode = 0: Exit Sub" & vbCrLf
    s = s & "    ' Ctrl+P = Print Report" & vbCrLf
    s = s & "    If KeyCode = 80 And ctrl Then Call btnPrintReport_Click: KeyCode = 0: Exit Sub" & vbCrLf
    s = s & "End Sub" & vbCrLf
    mdl.InsertLines sL, s
    Debug.Print "    Form_KeyDown aktualizovan (+Ctrl+B, +Ctrl+P)"

    ' Odstran [F5] z popisu tlacitka FILTER
    On Error Resume Next
    frm.Controls("btnFilter").Caption = "FILTER [F5]"
    ' Zkus obe varianty
    Dim bfCaption As String
    bfCaption = Nz(frm.Controls("btnFilter").Caption, "")
    If InStr(bfCaption, "[F") > 0 Then
        frm.Controls("btnFilter").Caption = "FILTER"
        Debug.Print "    btnFilter caption: FILTER [F5] -> FILTER"
    End If
    On Error GoTo Err3

Close3:
    DoCmd.Close acForm, "frmDark", acSaveYes
    Exit Sub
Err3: MsgBox "[3] Chyba: " & Err.Description, vbCritical
    On Error Resume Next: DoCmd.Close acForm, "frmDark", acSaveYes
End Sub

' ================================================================
' [4] frmImportXlsx - oprav DefaultValue (#NAZEV? problem)
'
'     Package 3 nastavil: c.DefaultValue = "TAXON"
'     Access vyhodnocuje "TAXON" jako vyraz (nazev pole).
'     Pole neexistuje v unbound formulari -> #NAZEV?
'
'     Oprava: DefaultValue = Chr(34) & "TAXON" & Chr(34)
'     = hodnota "TAXON" vcetne uvozovek -> Access to ulozi jako
'     literalni retezec, ne jako vyraz.
'
'     Pokud frmImportXlsx neexistuje, krok se preskoci.
' ================================================================
Private Sub Step4_FixImportXlsx()
    On Error GoTo Err4

    ' Zkus otevrit formular - pokud neexistuje, preskoc
    On Error Resume Next
    DoCmd.OpenForm "frmImportXlsx", acDesign
    Dim frm As Form
    Set frm = Forms("frmImportXlsx")
    If Err.Number <> 0 Or frm Is Nothing Then
        Debug.Print "    frmImportXlsx nelze otevrit (" & Err.Description & ")"
        Debug.Print "    Spus nejprve Package 3: BuildForms_3_ImportExport"
        Err.Clear: On Error GoTo 0: Exit Sub
    End If
    On Error GoTo Err4

    Dim fields(19) As String, defaults(19) As String
    fields(0) = "txtMapTaxon":      defaults(0) = "TAXON"
    fields(1) = "txtMapAuthor":     defaults(1) = "AUTHOR"
    fields(2) = "txtMapRank":       defaults(2) = "Rank"
    fields(3) = "txtMapGenus":      defaults(3) = "Genus"
    fields(4) = "txtMapFamily":     defaults(4) = "Family"
    fields(5) = "txtMapOrder":      defaults(5) = "Order"
    fields(6) = "txtMapGeog":       defaults(6) = "Geography"
    fields(7) = "txtMapStrat":      defaults(7) = "STRATIGRAPHY"
    fields(8) = "txtMapDesc":       defaults(8) = "DESCRIPTION"
    fields(9) = "txtMapSyn":        defaults(9) = "SYNONYMY"
    fields(10) = "txtMapDiag":      defaults(10) = "DIAGNOSIS"
    fields(11) = "txtMapEtym":      defaults(11) = "ETYMOLOGY"
    fields(12) = "txtMapOccur":     defaults(12) = "OCCURRENCE"
    fields(13) = "txtMapSystem":    defaults(13) = "System"
    fields(14) = "txtMapFormation": defaults(14) = "Formation"
    fields(15) = "txtMapStage":     defaults(15) = "Stage"
    fields(16) = "txtMapFigures":   defaults(16) = "FIGURES"
    fields(17) = "txtMapTypeMat":   defaults(17) = "TYPE MATERIAL"
    fields(18) = "txtMapMatEx":     defaults(18) = "MATERIAL EXAMINED"
    fields(19) = "txtMapRemarks":   defaults(19) = "REMARKS"

    Dim mi As Integer, fixed As Integer: fixed = 0
    For mi = 0 To 19
        On Error Resume Next
        frm.Controls(fields(mi)).DefaultValue = Chr(34) & defaults(mi) & Chr(34)
        If Err.Number = 0 Then fixed = fixed + 1
        Err.Clear
        On Error GoTo Err4
    Next mi

    Debug.Print "    Opraveno " & fixed & " DefaultValue poli"

    ' Oprav layout: presun tlacitka a log pod mapping pole
    ' Mapping pole konci na y~9500, tlacitka musi byt nize
    On Error Resume Next
    frm.Controls("btnImportXlsx").Top = 9800
    frm.Controls("btnImportXlsx").Left = 200
    frm.Controls("btnCloseXlsx").Top = 9800
    frm.Controls("btnCloseXlsx").Left = 2900
    frm.Controls("lblXlsxStatus").Top = 10600
    frm.Controls("txtXlsxLog").Top = 11100
    frm.Section(acDetail).Height = 14200
    On Error GoTo Err4

    Debug.Print "    Layout opraven (tlacitka na y=9800)"
    DoCmd.Close acForm, "frmImportXlsx", acSaveYes
    Exit Sub
Err4: MsgBox "[4] Chyba: " & Err.Description, vbCritical
    On Error Resume Next: DoCmd.Close acForm, "frmImportXlsx", acSaveYes
End Sub

' ================================================================
' [5] Opravi Overflow v BuildForms_5_PreviewReport
'
'     CreateTaxaReport_Ultimate nastavuje Section(acDetail).Height = 40000
'     Access HARD LIMIT je 22 palcu = 31680 twips.
'     40000 > 31680 -> Overflow (Error 6) pri spusteni Package 10.
'
'     Opravi v modulu 5_preview pres VBE bez otevirani formulare.
' ================================================================
Private Sub Step5_FixReportOverflow()
    On Error GoTo Err5

    ' Hledame modul ktery obsahuje CreateTaxaReport_Ultimate
    ' (neznamy nazev - muze byt 5_preview, 10_code_preview, atd.)
    Dim comp As Object, cm As Object
    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        On Error Resume Next
        Dim testLines As Long: testLines = comp.CodeModule.CountOfLines
        Dim j As Long
        For j = 1 To testLines
            If InStr(comp.CodeModule.Lines(j, 1), "CreateTaxaReport_Ultimate") > 0 And _
               InStr(comp.CodeModule.Lines(j, 1), "Sub ") > 0 Then
                Set cm = comp.CodeModule
                Debug.Print "    Nalezen modul: " & comp.Name
                Exit For
            End If
        Next j
        On Error GoTo Err5
        If Not cm Is Nothing Then Exit For
    Next comp

    If cm Is Nothing Then
        Debug.Print "    Modul s CreateTaxaReport_Ultimate nenalezen"
        Debug.Print "    Spus nejprve Package 10: BuildForms_5_PreviewReport"
        Exit Sub
    End If

    Dim i As Long, fixed As Integer: fixed = 0
    For i = 1 To cm.CountOfLines
        Dim ln As String: ln = cm.Lines(i, 1)
        ' Opravi inicializaci vysky sekce (40000 -> 31000)
        If InStr(ln, ".Section(acDetail).Height = 40000") > 0 Then
            cm.ReplaceLine i, Replace(ln, "40000&", "31000&")
            fixed = fixed + 1
            Debug.Print "    Radek " & i & ": Height 40000 -> 31000"
        End If
        ' Opravi finalni nastaveni vysky na konci funkce
        If InStr(ln, "rpt.Section(acDetail).Height = Y + 600") > 0 Then
            cm.ReplaceLine i, _
                Replace(ln, "rpt.Section(acDetail).Height = Y + 600", _
                        "rpt.Section(acDetail).Height = IIf(Y + 600 > 31000, 31000, Y + 600)")
            fixed = fixed + 1
            Debug.Print "    Radek " & i & ": Y+600 omezen na max 31000"
        End If
    Next i

    If fixed = 0 Then
        Debug.Print "    Zadne radky k oprave nenalezeny (uz opraveno nebo jiny format)"
    Else
        Debug.Print "    Opraveno " & fixed & " radku"
    End If
    Exit Sub
Err5: MsgBox "[5] Chyba: " & Err.Description, vbCritical
End Sub


' ================================================================
' [6] Odstran tlacitko PRINT REPORT a jeho handler z frmDark
'     (report rptTaxaDetail neexistuje v zakladnim buildu)
' ================================================================
Private Sub Step6_RemovePrintReport()
    On Error GoTo Err6
    DoCmd.OpenForm "frmDark", acDesign
    Dim frm As Form: Set frm = Forms("frmDark")
    Dim mdl As Module: Set mdl = frm.Module
    Dim fn As String: fn = frm.Name

    ' Smaz tlacitko
    On Error Resume Next
    DeleteControl fn, "btnPrintReport"
    If Err.Number = 0 Then
        Debug.Print "    btnPrintReport tlacitko smazano"
    Else
        Debug.Print "    btnPrintReport nenalezeno (ok)"
    End If
    Err.Clear
    On Error GoTo Err6

    ' Smaz handler z modulu
    Dim i As Long, sL As Long, eL As Long: sL = 0: eL = 0
    For i = 1 To mdl.CountOfLines
        Dim ln As String: ln = mdl.Lines(i, 1)
        If InStr(ln, "Sub btnPrintReport_Click") > 0 Then sL = i
        If sL > 0 And i > sL Then
            If Trim(ln) = "End Sub" Then eL = i: Exit For
        End If
    Next i
    If sL > 0 And eL > 0 Then
        mdl.DeleteLines sL, eL - sL + 1
        Debug.Print "    btnPrintReport_Click handler smazan"
    Else
        Debug.Print "    btnPrintReport_Click handler nenalezen (ok)"
    End If

    ' Odstran Ctrl+P z Form_KeyDown
    Dim i2 As Long, sL2 As Long, eL2 As Long: sL2 = 0: eL2 = 0
    For i2 = 1 To mdl.CountOfLines
        Dim ln2 As String: ln2 = mdl.Lines(i2, 1)
        If InStr(ln2, "Form_KeyDown") > 0 And InStr(ln2, "Sub ") > 0 Then sL2 = i2
        If sL2 > 0 And i2 > sL2 Then
            If Trim(ln2) = "End Sub" Then eL2 = i2: Exit For
        End If
    Next i2
    If sL2 > 0 And eL2 > 0 Then
        ' Smaz a vloz znovu bez Ctrl+P
        mdl.DeleteLines sL2, eL2 - sL2 + 1
        Dim s As String
        s = "Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)" & vbCrLf
        s = s & "    On Error Resume Next" & vbCrLf
        s = s & "    Dim ctrl As Boolean: ctrl = (Shift And acCtrlMask) > 0" & vbCrLf
        s = s & "    ' F4 = Quick Preview" & vbCrLf
        s = s & "    If KeyCode = 115 And Shift = 0 Then Call OpenQuickPreview: KeyCode = 0: Exit Sub" & vbCrLf
        s = s & "    ' F5 = Filter/Search" & vbCrLf
        s = s & "    If KeyCode = 116 And Shift = 0 Then Call ApplyFilters: KeyCode = 0: Exit Sub" & vbCrLf
        s = s & "    ' Esc = Clear" & vbCrLf
        s = s & "    If KeyCode = 27 And Shift = 0 Then Call btnClear_Click: KeyCode = 0: Exit Sub" & vbCrLf
        s = s & "    ' Ctrl+N = New Taxon" & vbCrLf
        s = s & "    If KeyCode = 78 And ctrl Then Call btnNewTaxon_Click: KeyCode = 0: Exit Sub" & vbCrLf
        s = s & "    ' Ctrl+Z = Undo" & vbCrLf
        s = s & "    If KeyCode = 90 And ctrl Then Call btnUndoAction_Click: KeyCode = 0: Exit Sub" & vbCrLf
        s = s & "    ' Ctrl+D = Open Detail" & vbCrLf
        s = s & "    If KeyCode = 68 And ctrl Then Call btnDetail_Click: KeyCode = 0: Exit Sub" & vbCrLf
        s = s & "    ' Ctrl+E = Export XLSX" & vbCrLf
        s = s & "    If KeyCode = 69 And ctrl Then Call btnExport_Click: KeyCode = 0: Exit Sub" & vbCrLf
        s = s & "    ' Ctrl+B = BibTeX" & vbCrLf
        s = s & "    If KeyCode = 66 And ctrl Then Call btnBibTeX_Click: KeyCode = 0: Exit Sub" & vbCrLf
        s = s & "End Sub" & vbCrLf
        mdl.InsertLines sL2, s
        Debug.Print "    Ctrl+P odstranen z Form_KeyDown"
    End If

    DoCmd.Close acForm, "frmDark", acSaveYes
    Exit Sub
Err6: MsgBox "[6] Chyba: " & Err.Description, vbCritical
    On Error Resume Next: DoCmd.Close acForm, "frmDark", acSaveYes
End Sub

' ================================================================
' Helper
' ================================================================
Private Function Exists(frm As Form, nm As String) As Boolean
    On Error Resume Next
    Dim c As Control: Set c = frm.Controls(nm)
    Exists = Not (c Is Nothing)
    On Error GoTo 0
End Function
