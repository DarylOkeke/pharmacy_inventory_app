Option Explicit

' ===== SHEET/TABLE NAMES (used by current demo) =====
Private Const SHEET_SUP_PRESETS As String = "SupplierPresets"
Private Const TABLE_SUP_PRESETS As String = "tblPresets"

Private Const SHEET_SUPPLIER As String = "SupplierData"
Private Const SUP_TABLE As String = "Table3"
Private Const SUP_HEADER_ROW As Long = 8
Private Const SUP_START_ROW  As Long = 9

Private Const SHEET_PHYSICAL As String = "PhysicalCount"
Private Const PHYS_TABLE As String = "Table13"
Private Const HDR_DRUGNAME As String = "Drug Name"
Private Const HDR_DRUGID   As String = "Drug ID"
Private Const HDR_COUNT    As String = "Physical Count"
Private Const HDR_DATELOG  As String = "Date logged"

' Tracks which dataset the user imported last (used by the input form + report)
Public CurrentDataset As String

' ===========================
'            MACROS
' ===========================

' -------- ImportWeeklyData --------
' Choose a preset (SupplierPresets!tblPresets) or import a file (A:C from row 5),
' write into SupplierData!Table3 (A9:C...), then sync PhysicalCount Table13.
Public Sub ImportWeeklyData()
    Dim stage As String
    Dim prevCalc As XlCalculation
    Dim prevEvents As Boolean, prevUpdating As Boolean
    Dim f As Variant, wbSrc As Workbook, wsSrc As Worksheet

    On Error GoTo Fail
    stage = "Prep"
    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' --- picker ---
    stage = "Show picker"
    EnsureSupplierPresetsSeeded
    Load frmSupplierPicker
    frmSupplierPicker.InitPicker
    frmSupplierPicker.Show

    If frmSupplierPicker.UserChoice = "" Then GoTo CleanExit   ' cancel

    If frmSupplierPicker.UserChoice = "preset" Then
        ' -------- PRESET PATH --------
        stage = "Import from preset"
        If Not InlineImportFromPreset(frmSupplierPicker.SelectedDataset) Then GoTo CleanExit

        CurrentDataset = frmSupplierPicker.SelectedDataset
        ClearLoggedInventory CurrentDataset
        MsgBox "Loaded " & CurrentDataset & " dataset", vbInformation
        ShowAndActivate SHEET_SUPPLIER

    ElseIf frmSupplierPicker.UserChoice = "file" Then
        ' -------- FILE PATH --------
        stage = "File picker"
        f = Application.GetOpenFilename("Excel Files (*.xlsx;*.xls;*.xlsm),*.xlsx;*.xls;*.xlsm")
        If VarType(f) = vbBoolean And f = False Then GoTo CleanExit   ' user canceled

        stage = "Open source workbook"
        Set wbSrc = Application.Workbooks.Open(CStr(f), ReadOnly:=True)

        stage = "Resolve source sheet"
        Set wsSrc = ResolveSourceSheet(wbSrc)
        If wsSrc Is Nothing Then Err.Raise 5, , "No worksheet found in the selected workbook."

        stage = "Read A:C from row 5 down"
        Const SRC_FIRST_ROW As Long = 5
        Dim lastSrc As Long
        lastSrc = LastRowAny(wsSrc, Array(1, 2, 3))
        If lastSrc < SRC_FIRST_ROW Then
            MsgBox "Selected sheet has no data at or below row 5 (columns A–C).", vbExclamation
        Else
            Dim srcData As Variant
            srcData = wsSrc.Range(wsSrc.Cells(SRC_FIRST_ROW, 1), wsSrc.Cells(lastSrc, 3)).Value
            WriteSupplierDataFromArray srcData

            Dim dsName As String
            On Error Resume Next
            dsName = Dir$(CStr(f)) ' file name only
            On Error GoTo 0
            If Len(Trim$(dsName)) = 0 Then dsName = "Manual File"
            CurrentDataset = dsName
            ClearLoggedInventory CurrentDataset
            MsgBox "Loaded " & CurrentDataset & " dataset", vbInformation
            ShowAndActivate SHEET_SUPPLIER
        End If

        On Error Resume Next
        wbSrc.Close SaveChanges:=False
        On Error GoTo 0

    Else
        Err.Raise 5, , "Unknown picker choice: " & frmSupplierPicker.UserChoice
    End If

    ' --- sync PhysicalCount after SupplierData is updated ---
    stage = "Sync PhysicalCount"
    UpdatePhysicalCountTable

CleanExit:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevUpdating
    Application.Calculation = prevCalc
    Exit Sub

Fail:
    MsgBox "ImportWeeklyData failed at stage: " & stage & vbCrLf & _
           "Error " & Err.Number & " - " & Err.Description, vbExclamation
    On Error Resume Next
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevUpdating
    Application.Calculation = prevCalc
End Sub

' -------- ShowInventoryInputForm --------
' Launch the inventory input form and pre-load drugs from SupplierData!Table3
Public Sub ShowInventoryInputForm()
    Dim wsSup As Worksheet, lo As ListObject
    Set wsSup = SheetByName(SHEET_SUPPLIER)
    If wsSup Is Nothing Then
        MsgBox "SupplierData sheet not found.", vbExclamation
        Exit Sub
    End If

    Set lo = EnsureSupplierTable(wsSup)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        MsgBox "No supplier data loaded. Import weekly data first.", vbInformation
        Exit Sub
    End If

    Load frmInventoryInput
    frmInventoryInput.LoadDrugs lo.DataBodyRange, IIf(Len(Trim$(CurrentDataset)) > 0, CurrentDataset, "Active Dataset")
    frmInventoryInput.Show
End Sub

' -------- GenerateReport --------
Public Sub GenerateReport()
    Dim prevCalc As XlCalculation, prevEvents As Boolean, prevUpdating As Boolean
    On Error GoTo CleanFail
    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim wsSup As Worksheet, wsRep As Worksheet
    Set wsSup = SheetByName(SHEET_SUPPLIER)

    If Not SupplierHasAnyRows() Then
        MsgBox "Import a dataset, then log your inventory before generating a report", vbExclamation
        GoTo CleanExit
    End If

    ' Supplier dictionaries (by Drug Name)
    Dim dictNameToID As Object, dictNameToExp As Object, dictNameToPretty As Object
    Set dictNameToID = CreateObject("Scripting.Dictionary")
    Set dictNameToExp = CreateObject("Scripting.Dictionary")
    Set dictNameToPretty = CreateObject("Scripting.Dictionary")
    LoadSupplierDictionaries wsSup, dictNameToID, dictNameToExp, dictNameToPretty

    ' Fallback map: DrugID -> Expected
    Dim dictIDToExp As Object
    Set dictIDToExp = CreateObject("Scripting.Dictionary")
    Dim supLast As Long, rRow As Long
    Dim idVal As String, expVal As Variant
    supLast = LastRowAny(wsSup, Array(1, 2, 3))
    For rRow = SUP_START_ROW To supLast
        idVal = NzStr(wsSup.Cells(rRow, "B").Value)
        expVal = wsSup.Cells(rRow, "C").Value
        If Len(idVal) > 0 And IsNumeric(expVal) Then
            dictIDToExp(NormKey(idVal)) = CDbl(expVal)
        End If
    Next rRow

    ' Logged inventory only (enforce logging)
    Dim dictNameToPhys As Object, dictPhysPretty As Object
    Set dictNameToPhys = CreateObject("Scripting.Dictionary")
    Set dictPhysPretty = CreateObject("Scripting.Dictionary")

    If Len(Trim$(CurrentDataset)) = 0 Then
        If Not LoadLoggedInventoryDict("Active Dataset", dictNameToPhys, dictPhysPretty) Then
            MsgBox "Log inventory before generating a report", vbExclamation
            GoTo CleanExit
        End If
    Else
        If Not LoadLoggedInventoryDict(CurrentDataset, dictNameToPhys, dictPhysPretty) Then
            MsgBox "Log inventory before generating a report", vbExclamation
            GoTo CleanExit
        End If
    End If

    ' Build/Reset Report sheet
    Set wsRep = AddOrResetSheet("Report", wsSup)

    With wsRep.Range("A1:F2")
        .Merge
        .Value = "Inventory Report"
        .Font.Bold = True
        .Font.Size = 20
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With wsRep
        .Range("A4").Value = "Drug Name"
        .Range("B4").Value = "Drug ID"
        .Range("C4").Value = "Physical Count"
        .Range("D4").Value = "Expected Count"
        .Range("E4").Value = "Status"
        .Range("F4").Value = "Comments"
        .Range("A4:F4").Font.Bold = True
    End With

    Dim r As Long: r = 5
    Dim k As Variant, displayName As String
    Dim drugID As Variant, phys As Variant, expct As Variant
    Dim statusTxt As String, diff As Double

    For Each k In dictNameToPhys.Keys
        phys = dictNameToPhys(k)
        If Not IsNumeric(phys) Or CDbl(phys) <= 0 Then GoTo NextK

        displayName = IIf(dictPhysPretty.Exists(k), CStr(dictPhysPretty(k)), CStr(k))
        drugID = vbNullString
        If dictNameToID.Exists(k) Then drugID = dictNameToID(k)

        expct = vbNullString
        If dictNameToExp.Exists(k) Then expct = dictNameToExp(k)

        If (Not IsNumeric(expct)) And Len(NzStr(drugID)) > 0 Then
            Dim idKey As String: idKey = NormKey(CStr(drugID))
            If dictIDToExp.Exists(idKey) Then expct = dictIDToExp(idKey)
        End If

        If IsNumeric(expct) Then
            diff = CDbl(phys) - CDbl(expct)
            If diff > 0 Then
                statusTxt = "surplus of " & CStr(CLng(diff))
            ElseIf diff < 0 Then
                statusTxt = "shortage of " & CStr(CLng(Abs(diff)))
            Else
                statusTxt = "even"
            End If
        Else
            statusTxt = "no expected"
        End If

        With wsRep
            .Cells(r, "A").Value = displayName
            .Cells(r, "B").Value = drugID
            .Cells(r, "C").Value = phys
            .Cells(r, "D").Value = expct
            .Cells(r, "E").Value = statusTxt
            .Cells(r, "F").Value = vbNullString
        End With
        r = r + 1
NextK:
    Next k

    Dim lastDataRow As Long: lastDataRow = r - 1

    If lastDataRow >= 5 Then
        Dim tbl As ListObject
        Set tbl = wsRep.ListObjects.Add( _
            SourceType:=xlSrcRange, _
            Source:=wsRep.Range("A4:F" & lastDataRow), _
            XlListObjectHasHeaders:=xlYes)
        On Error Resume Next
        tbl.Name = "ReportTable"
        If Err.Number <> 0 Then
            Err.Clear
            tbl.Name = "ReportTable_" & Format(Now, "yyyymmdd_hhnnss")
        End If
        On Error GoTo 0

        wsRep.Columns("A:F").AutoFit

        Dim dataRange As Range: Set dataRange = wsRep.Range("A5:F" & lastDataRow)
        With dataRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISNUMBER(SEARCH(""surplus"",$E5))")
            .Interior.Color = RGB(198, 239, 206)
        End With
        With dataRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISNUMBER(SEARCH(""shortage"",$E5))")
            .Interior.Color = RGB(255, 199, 206)
        End With
    Else
        wsRep.Columns("A:F").AutoFit
    End If

    AddReportButtons wsRep
    ShowAndActivate "Report"
    Application.Goto wsRep.Range("A4"), False

CleanExit:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevUpdating
    Application.Calculation = prevCalc
    Exit Sub

CleanFail:
    MsgBox "GenerateReport failed: " & Err.Number & " - " & Err.Description, vbExclamation
    GoTo CleanExit
End Sub

' -------- UpdatePhysicalCountTable --------
Public Sub UpdatePhysicalCountTable()
    Dim prevCalc As XlCalculation, prevEvents As Boolean, prevUpdating As Boolean
    On Error GoTo CleanFail
    prevCalc = Application.Calculation
    prevEvents = Application.EnableEvents
    prevUpdating = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim lo As ListObject: Set lo = GetPhysicalTable()
    If lo Is Nothing Then Err.Raise 5, , "PhysicalCount table '" & PHYS_TABLE & "' not found."

    Dim cName As Long: cName = HeaderIndex(lo, HDR_DRUGNAME)
    Dim cID As Long:   cID = HeaderIndex(lo, HDR_DRUGID)
    If cName = 0 Then cName = 1
    If cID = 0 Then cID = 3

    Dim wsSup As Worksheet: Set wsSup = SheetByName(SHEET_SUPPLIER)
    Dim dictNameToID As Object, dictNameToPretty As Object, dictNameToExp As Object
    Set dictNameToID = CreateObject("Scripting.Dictionary")
    Set dictNameToExp = CreateObject("Scripting.Dictionary")
    Set dictNameToPretty = CreateObject("Scripting.Dictionary")
    LoadSupplierDictionaries wsSup, dictNameToID, dictNameToExp, dictNameToPretty

    Dim dictPhysRows As Object: Set dictPhysRows = CreateObject("Scripting.Dictionary")
    Dim i As Long, nm As String
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.DataBodyRange.Rows.Count
            nm = NzStr(lo.DataBodyRange.Cells(i, cName).Value)
            If Len(nm) > 0 Then dictPhysRows(NormKey(nm)) = i
        Next i
    End If

    Dim k As Variant, rowIx As Long
    For Each k In dictNameToID.Keys
        If dictPhysRows.Exists(k) Then
            rowIx = dictPhysRows(k)
            If lo.DataBodyRange.Cells(rowIx, cID).Value <> dictNameToID(k) Then
                lo.DataBodyRange.Cells(rowIx, cID).Value = dictNameToID(k)
            End If
        Else
            Dim lr As ListRow: Set lr = lo.ListRows.Add
            With lr.Range
                .Cells(1, cName).Value = dictNameToPretty(k)
                .Cells(1, cID).Value = dictNameToID(k)
                If lo.ListColumns.Count >= 2 Then .Cells(1, 2).Value = "Newly added drug"
                If lo.ListColumns.Count >= 4 Then .Cells(1, 4).Value = "Newly added drug"
                If lo.ListColumns.Count >= 5 Then .Cells(1, 5).Value = "Newly added drug"
            End With
        End If
    Next k

CleanExit:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevUpdating
    Application.Calculation = prevCalc
    Exit Sub
CleanFail:
    MsgBox "UpdatePhysicalCountTable failed: " & Err.Number & " - " & Err.Description, vbExclamation
    GoTo CleanExit
End Sub

' ===========================
'     PUBLIC HELPERS
' ===========================
Public Function GetPhysicalTable() As ListObject
    On Error Resume Next
    Set GetPhysicalTable = SheetByName(SHEET_PHYSICAL).ListObjects(PHYS_TABLE)
    On Error GoTo 0
End Function

' Public so forms can call it (kept minimal)
Public Function HeaderIndexLocal(lo As ListObject, headerText As String) As Long
    Dim i As Long: HeaderIndexLocal = 0
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(Trim$(CStr(lo.ListColumns(i).Name)), Trim$(headerText), vbTextCompare) = 0 Then
            HeaderIndexLocal = i
            Exit Function
        End If
    Next i
End Function

' Safe sheet getter (Public so forms can call)
Public Function SheetByName(sName As String) As Worksheet
    On Error Resume Next
    Set SheetByName = ThisWorkbook.Worksheets(sName)
    On Error GoTo 0
End Function

' Save logged counts into a hidden sheet keyed by dataset
Public Sub SaveLoggedInventory(ByVal datasetName As String, _
                               ByVal drugNames As Variant, _
                               ByVal drugIDs As Variant, _
                               ByVal counts As Variant)
    Dim ws As Worksheet
    Set ws = SheetByName("LoggedInventory")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "LoggedInventory"
        ws.Visible = xlSheetHidden
        ws.Range("A1:D1").Value = Array("Dataset", "Drug Name", "Drug ID", "Count")
    End If

    Dim last As Long, r As Long
    last = lastRow(ws, 1)
    Application.ScreenUpdating = False
    For r = last To 2 Step -1
        If StrComp(CStr(ws.Cells(r, 1).Value), datasetName, vbTextCompare) = 0 Then ws.Rows(r).Delete
    Next r
    Application.ScreenUpdating = True

    Dim i As Long, startRow As Long
    startRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    For i = LBound(drugNames) To UBound(drugNames)
        ws.Cells(startRow, 1).Value = datasetName
        ws.Cells(startRow, 2).Value = drugNames(i)
        ws.Cells(startRow, 3).Value = drugIDs(i)
        ws.Cells(startRow, 4).Value = counts(i)
        startRow = startRow + 1
    Next i

    MsgBox "Inventory logged for: " & datasetName, vbInformation
End Sub

' Clear all logged inventory, or only for a specific dataset if provided
Public Sub ClearLoggedInventory(Optional ByVal datasetName As String = "")
    Dim ws As Worksheet
    Dim last As Long, r As Long

    Set ws = SheetByName("LoggedInventory")
    If ws Is Nothing Then Exit Sub

    last = lastRow(ws, 1)
    If last < 2 Then Exit Sub

    Application.ScreenUpdating = False
    If Len(Trim$(datasetName)) = 0 Then
        ws.Rows("2:" & last).Delete
    Else
        For r = last To 2 Step -1
            If StrComp(CStr(ws.Cells(r, 1).Value), datasetName, vbTextCompare) = 0 Then
                ws.Rows(r).Delete
            End If
        Next r
    End If
    Application.ScreenUpdating = True
End Sub

' Navigate to a "HOME" sheet if present; otherwise SupplierData
Public Sub GoHOME()
    Dim ws As Worksheet
    Set ws = SheetByName("HOME")
    If ws Is Nothing Then
        Set ws = SheetByName(SHEET_SUPPLIER)
        If ws Is Nothing Then
            MsgBox "No 'HOME' or '" & SHEET_SUPPLIER & "' sheet found.", vbExclamation
            Exit Sub
        End If
    End If
    ws.Activate
End Sub

' Export the Report sheet to PDF (same folder as workbook)
Public Sub ExportReport()
    Dim ws As Worksheet
    Set ws = SheetByName("Report")
    If ws Is Nothing Then
        MsgBox "Report sheet not found. Generate the report first.", vbExclamation
        Exit Sub
    End If

    Dim savePath As String, base As String
    base = Left$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1)
    savePath = ThisWorkbook.Path & Application.PathSeparator & base & "_InventoryReport_" & Format(Now, "yyyymmdd_hhmmss") & ".pdf"

    On Error GoTo Fail
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=savePath, Quality:=xlQualityStandard, _
                           IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    MsgBox "Exported PDF:" & vbCrLf & savePath, vbInformation
    Exit Sub
Fail:
    MsgBox "Export failed: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' ===========================
'     PRIVATE HELPERS
' ===========================

' Build dictionaries from SupplierData A:C (rows 9 : last)
Private Sub LoadSupplierDictionaries(wsSup As Worksheet, _
    dictNameToID As Object, dictNameToExp As Object, dictNameToPretty As Object)

    Dim supLast As Long: supLast = LastRowAny(wsSup, Array(1, 2, 3))
    Dim r As Long

    For r = SUP_START_ROW To supLast
        Dim nm As String: nm = NzStr(wsSup.Cells(r, "A").Value)
        If Len(nm) = 0 Then GoTo NextR

        Dim key As String: key = NormKey(nm)
        If Not dictNameToID.Exists(key) Then
            dictNameToID(key) = NzStr(wsSup.Cells(r, "B").Value)
            dictNameToExp(key) = NzNum(wsSup.Cells(r, "C").Value)
            dictNameToPretty(key) = nm
        End If
NextR:
    Next r
End Sub

' Pick the source worksheet:
' 1) first visible sheet containing "supplier" in its name
' 2) first visible sheet
' 3) first sheet at all
Private Function ResolveSourceSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            If InStr(1, ws.Name, "supplier", vbTextCompare) > 0 Then
                Set ResolveSourceSheet = ws
                Exit Function
            End If
        End If
    Next ws
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            Set ResolveSourceSheet = ws
            Exit Function
        End If
    Next ws
    If wb.Worksheets.Count > 0 Then
        Set ResolveSourceSheet = wb.Worksheets(1)
    Else
        Set ResolveSourceSheet = Nothing
    End If
End Function

' Ensure SupplierData table exists as A8:C? with header row at 8, first body row at 9
Private Function EnsureSupplierTable(wsSup As Worksheet) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = wsSup.ListObjects(SUP_TABLE)
    If lo Is Nothing Then Set lo = wsSup.ListObjects("Table 3")
    On Error GoTo 0

    If lo Is Nothing Then
        If Len(NzStr(wsSup.Range("A" & SUP_HEADER_ROW).Value)) = 0 Then wsSup.Range("A" & SUP_HEADER_ROW).Value = "Drug Name"
        If Len(NzStr(wsSup.Range("B" & SUP_HEADER_ROW).Value)) = 0 Then wsSup.Range("B" & SUP_HEADER_ROW).Value = "Drug ID"
        If Len(NzStr(wsSup.Range("C" & SUP_HEADER_ROW).Value)) = 0 Then wsSup.Range("C" & SUP_HEADER_ROW).Value = "Expected Quantity"

        Dim rng As Range
        Set rng = wsSup.Range("A" & SUP_HEADER_ROW & ":C" & SUP_START_ROW) ' A8:C9
        Set lo = wsSup.ListObjects.Add(xlSrcRange, rng, , xlYes)
        On Error Resume Next
        lo.Name = SUP_TABLE
        On Error GoTo 0
    End If

    Set EnsureSupplierTable = lo
End Function

' Header index (case-insensitive) inside a ListObject
Private Function HeaderIndex(lo As ListObject, headerText As String) As Long
    Dim i As Long
    HeaderIndex = 0
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(Trim$(CStr(lo.ListColumns(i).Name)), Trim$(headerText), vbTextCompare) = 0 Then
            HeaderIndex = i: Exit Function
        End If
    Next i
End Function

' Write a 2D array [rows x 3] to SupplierData!Table3 (A9:C...), resizing the table cleanly
Private Sub WriteSupplierDataFromArray(ByVal srcData As Variant)
    On Error GoTo Trap
    Dim wsSup As Worksheet, lo As ListObject
    Dim nRows As Long, nCols As Long
    Dim hdr As Range, newRange As Range

    Set wsSup = SheetByName(SHEET_SUPPLIER)
    If wsSup Is Nothing Then Err.Raise 9, , "Sheet '" & SHEET_SUPPLIER & "' not found."

    Set lo = EnsureSupplierTable(wsSup)
    If lo Is Nothing Then Err.Raise 9, , "Could not locate SupplierData table '" & SUP_TABLE & "'."

    If IsEmpty(srcData) Then Err.Raise 5, , "No rows to import."

    On Error GoTo BadShape
    nRows = UBound(srcData, 1)
    nCols = UBound(srcData, 2)
    On Error GoTo Trap

    If nCols <> 3 Then Err.Raise 5, , "Import array must have 3 columns; got " & nCols & "."

    Set hdr = wsSup.Range("A" & SUP_HEADER_ROW) ' A8
    Set newRange = wsSup.Range(hdr, hdr.Offset(nRows, 3 - 1))
    lo.Resize newRange

    wsSup.Range("A" & SUP_START_ROW).Resize(nRows, 3).Value = srcData
    Exit Sub

BadShape:
    Err.Raise 5, , "Import array is not 2D."
Trap:
    MsgBox "WriteSupplierDataFromArray failed:" & vbCrLf & _
           "Rows: " & nRows & ", Cols: " & nCols & vbCrLf & _
           "Error " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

' ===== PRESET DATA (sheet + table) =====
Public Sub EnsureSupplierPresetsSeeded()
    Dim ws As Worksheet, lo As ListObject
    Set ws = SheetByName(SHEET_SUP_PRESETS)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SHEET_SUP_PRESETS
        ws.Visible = xlSheetHidden
    End If

    ws.Range("A1").Value = "Dataset"
    ws.Range("B1").Value = "Drug Name"
    ws.Range("C1").Value = "Drug ID"
    ws.Range("D1").Value = "Expected Quantity"

    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_SUP_PRESETS)
    On Error GoTo 0
    If lo Is Nothing Then
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:D2"), , xlYes)
        On Error Resume Next: lo.Name = TABLE_SUP_PRESETS: On Error GoTo 0
        lo.TableStyle = "TableStyleMedium2"
    Else
        lo.HeaderRowRange.Cells(1, 1).Value = "Dataset"
        If lo.ListColumns.Count >= 2 Then lo.HeaderRowRange.Cells(1, 2).Value = "Drug Name"
        If lo.ListColumns.Count >= 3 Then lo.HeaderRowRange.Cells(1, 3).Value = "Drug ID"
        If lo.ListColumns.Count >= 4 Then lo.HeaderRowRange.Cells(1, 4).Value = "Expected Quantity"
    End If

    Dim needSeed As Boolean
    If lo.DataBodyRange Is Nothing Then
        needSeed = True
    Else
        needSeed = (Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0)
    End If

    If needSeed Then
        Dim datasets As Variant: datasets = Array("Urbana Clinic", "Chicago Center", "Peoria Hospital")
        Dim n As Long: n = 25
        Dim i As Long, r As Long, d As Long
        r = 1
        Dim arr() As Variant: ReDim arr(1 To (UBound(datasets) + 1) * n, 1 To 4)

        Randomize
        For d = LBound(datasets) To UBound(datasets)
            For i = 1 To n
                arr(r, 1) = datasets(d)
                arr(r, 2) = FakeDrugName(i + d * n)
                arr(r, 3) = FakeID()
                arr(r, 4) = FakeQty()
                r = r + 1
            Next i
        Next d

        lo.Resize ws.Range("A1:D" & (UBound(arr, 1) + 1))
        lo.DataBodyRange.Value = arr
        ws.Columns("A:D").AutoFit
    End If
End Sub

' Pull a chosen preset into SupplierData!Table3 (Drug, ID, Expected)
Private Function InlineImportFromPreset(ByVal datasetName As String) As Boolean
    On Error GoTo Fail

    Dim want As String: want = NormText(datasetName)
    If Len(want) = 0 Then
        MsgBox "No preset selected.", vbExclamation
        Exit Function
    End If

    Dim wsP As Worksheet: Set wsP = SheetByName(SHEET_SUP_PRESETS)
    If wsP Is Nothing Then
        MsgBox "Sheet '" & SHEET_SUP_PRESETS & "' not found.", vbExclamation
        Exit Function
    End If

    Dim loP As ListObject
    On Error Resume Next
    Set loP = wsP.ListObjects(TABLE_SUP_PRESETS)
    On Error GoTo 0
    If loP Is Nothing Then
        MsgBox "Table '" & TABLE_SUP_PRESETS & "' not found on '" & SHEET_SUP_PRESETS & "'.", vbExclamation
        Exit Function
    End If
    If loP.DataBodyRange Is Nothing Then
        MsgBox "'" & TABLE_SUP_PRESETS & "' has no rows.", vbExclamation
        Exit Function
    End If

    Dim cDS As Long, cDrug As Long, cID As Long, cExp As Long
    cDS = ColIndexByHeader(loP, "Dataset")
    cDrug = ColIndexByHeader(loP, "Drug Name")
    cID = ColIndexByHeader(loP, "Drug ID")
    cExp = ColIndexByHeader(loP, "Expected Quantity")
    If cDS = 0 Or cDrug = 0 Or cID = 0 Or cExp = 0 Then
        MsgBox "Preset table missing required headers (Dataset, Drug Name, Drug ID, Expected Quantity).", vbExclamation
        Exit Function
    End If

    Dim arr As Variant
    arr = loP.DataBodyRange.Value

    Dim nRows As Long, nCols As Long
    On Error GoTo Fail
    nRows = UBound(arr, 1)
    nCols = UBound(arr, 2)

    If cDS > nCols Or cDrug > nCols Or cID > nCols Or cExp > nCols Then
        MsgBox "'" & TABLE_SUP_PRESETS & "' body has only " & nCols & _
               " columns; expected at least 4. Try resizing the table to include A:D.", vbExclamation
        Exit Function
    End If

    Dim results As New Collection
    Dim i As Long, ds As String
    For i = 1 To nRows
        ds = NormText(CStr(arr(i, cDS)))
        If ds = want Then
            results.Add Array(arr(i, cDrug), arr(i, cID), arr(i, cExp))
        End If
    Next i

    If results.Count = 0 Then
        MsgBox "Preset '" & datasetName & "' has no matching rows in '" & TABLE_SUP_PRESETS & "'.", vbInformation
        Exit Function
    End If

    Dim outArr() As Variant, rowv As Variant
    ReDim outArr(1 To results.Count, 1 To 3)
    For i = 1 To results.Count
        rowv = results(i)
        outArr(i, 1) = rowv(0)
        outArr(i, 2) = rowv(1)
        outArr(i, 3) = rowv(2)
    Next i

    WriteSupplierDataFromArray outArr
    InlineImportFromPreset = True
    Exit Function

Fail:
    MsgBox "InlineImportFromPreset failed for '" & datasetName & "':" & vbCrLf & _
           "Error " & Err.Number & " - " & Err.Description, vbExclamation
End Function

' Buttons on Report
Private Sub AddReportButtons(ws As Worksheet)
    On Error Resume Next
    ws.Shapes("btnHOME").Delete
    ws.Shapes("btnExport").Delete
    On Error GoTo 0

    Dim topCell As Range
    Set topCell = ws.Range("H1")

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(Type:=msoShapeRoundedRectangle, _
                                 Left:=topCell.Left, _
                                 Top:=topCell.Top + 4, _
                                 Width:=90, Height:=26)
    shp.Name = "btnHOME"
    shp.TextFrame2.TextRange.Characters.Text = "HOME"
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.Fill.ForeColor.RGB = RGB(0, 97, 0)
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp.OnAction = "GoHOME"

    Dim shp2 As Shape
    Set shp2 = ws.Shapes.AddShape(Type:=msoShapeRoundedRectangle, _
                                  Left:=ws.Range("J1").Left, _
                                  Top:=topCell.Top + 4, _
                                  Width:=120, Height:=26)
    shp2.Name = "btnExport"
    shp2.TextFrame2.TextRange.Characters.Text = "Export Report"
    shp2.TextFrame2.TextRange.Font.Bold = msoTrue
    shp2.Fill.ForeColor.RGB = RGB(0, 97, 0)
    shp2.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    shp2.OnAction = "ExportReport"
End Sub

' Show a sheet by name and activate it
Public Sub ShowAndActivate(ByVal sheetName As String)
    Dim ws As Worksheet
    Set ws = SheetByName(sheetName)
    If ws Is Nothing Then
        MsgBox "Sheet not found: " & sheetName, vbExclamation
        Exit Sub
    End If
    ws.Visible = xlSheetVisible
    ws.Activate
End Sub

' Create-or-reset a worksheet (keeps the name; avoids name-collision errors)
Private Function AddOrResetSheet(ByVal sheetName As String, Optional ByVal afterSheet As Worksheet) As Worksheet
    Dim Sh As Object

    On Error Resume Next
    Set Sh = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If Not Sh Is Nothing Then
        If TypeOf Sh Is Worksheet Then
            Dim w As Worksheet
            Set w = Sh
            On Error Resume Next
            w.Cells.Clear
            Do While w.ListObjects.Count > 0
                w.ListObjects(1).Unlist
            Loop
            Dim s As Shape
            For Each s In w.Shapes
                s.Delete
            Next s
            Dim co As ChartObject
            For Each co In w.ChartObjects
                co.Delete
            Next co
            On Error GoTo 0

            If Not afterSheet Is Nothing Then
                On Error Resume Next
                w.Move After:=afterSheet
                On Error GoTo 0
            End If

            Set AddOrResetSheet = w
            Exit Function
        Else
            On Error Resume Next
            Sh.Name = sheetName & "_old"
            If Err.Number <> 0 Then
                Err.Clear
                Sh.Name = sheetName & "_old_" & Format(Now, "yyyymmdd_hhnnss")
            End If
            On Error GoTo 0
        End If
    End If

    Dim newWS As Worksheet
    If afterSheet Is Nothing Then
        Set newWS = ThisWorkbook.Worksheets.Add
    Else
        Set newWS = ThisWorkbook.Worksheets.Add(After:=afterSheet)
    End If

    On Error Resume Next
    newWS.Name = sheetName
    If Err.Number <> 0 Then
        Err.Clear
        newWS.Name = sheetName & "_" & Format(Now, "yyyymmdd_hhnnss")
    End If
    On Error GoTo 0

    Set AddOrResetSheet = newWS
End Function

' -------- Utility: which rows have content across several columns?
Private Function LastRowAny(ws As Worksheet, ByVal cols As Variant) As Long
    Dim i As Long, c As Long, r As Long, maxR As Long
    maxR = 1
    For i = LBound(cols) To UBound(cols)
        c = CLng(cols(i))
        If Application.WorksheetFunction.CountA(ws.Columns(c)) > 0 Then
            r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
            If r > maxR Then maxR = r
        End If
    Next i
    LastRowAny = maxR
End Function

' Last used row in a single column
Private Function lastRow(ws As Worksheet, col As Long) As Long
    With ws
        If Application.WorksheetFunction.CountA(.Columns(col)) = 0 Then
            lastRow = 1
        Else
            lastRow = .Cells(.Rows.Count, col).End(xlUp).Row
        End If
    End With
End Function

' Return the 1-based column index inside a ListObject by header text (robust).
Private Function ColIndexByHeader(lo As ListObject, ByVal headerText As String) As Long
    Dim i As Long, need As String, have As String
    need = NormText(headerText)
    ColIndexByHeader = 0
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        have = NormText(CStr(lo.ListColumns(i).Name))
        If have = need Then ColIndexByHeader = i: Exit Function
    Next i
End Function

' Normalize text: replace non-breaking spaces/tabs, collapse spaces, trim, lowercase.
Private Function NormText(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace$(t, Chr$(160), " ")
    t = Replace$(t, vbTab, " ")
    t = Trim$(t)
    Do While InStr(t, "  ") > 0
        t = Replace$(t, "  ", " ")
    Loop
    NormText = LCase$(t)
End Function

' Normalize key for dictionary
Private Function NormKey(ByVal s As String) As String
    NormKey = LCase$(Trim$(s))
End Function

' Null/Empty? "" ; else trimmed string
Private Function NzStr(v As Variant) As String
    If IsError(v) Or IsNull(v) Then
        NzStr = vbNullString
    Else
        NzStr = Trim$(CStr(v))
    End If
End Function

' Numeric? Double ; else ""
Private Function NzNum(v As Variant) As Variant
    If IsNumeric(v) Then
        NzNum = CDbl(v)
    Else
        NzNum = vbNullString
    End If
End Function

' Supplier has any rows?
Private Function SupplierHasAnyRows() As Boolean
    Dim ws As Worksheet, lo As ListObject
    SupplierHasAnyRows = False
    Set ws = SheetByName(SHEET_SUPPLIER)
    If ws Is Nothing Then Exit Function
    Set lo = EnsureSupplierTable(ws)
    If lo Is Nothing Then Exit Function
    If Not lo.DataBodyRange Is Nothing Then
        SupplierHasAnyRows = (lo.DataBodyRange.Rows.Count > 0)
    End If
End Function

' --- Fake value generators for seeding presets ---
Private Function FakeDrugName(ByVal seed As Long) As String
    Dim pre, root, suf, form, dose
    pre = Array("Azo", "Carbo", "Dexo", "Lumi", "Zylo", "Metra", "Vita", "Nexo", "Helio", "Orbi")
    root = Array("vent", "cort", "pril", "mab", "zid", "dopa", "nex", "cillin", "vast", "zole")
    suf = Array("tab", "cap", "sol", "inj", "XR", "SR", "ODT", "gel", "susp", "elix")
    form = Array("Tab", "Cap", "Oral Sol.", "Inj.", "XR Tab", "SR Cap", "ODT", "Gel", "Susp.", "Elixir")
    dose = Array("5mg", "10mg", "20mg", "40mg", "80mg", "120mL", "250mg", "500mg")

    Randomize seed
    FakeDrugName = pre(Int(Rnd * 10)) & root(Int(Rnd * 10)) & " " & form(Int(Rnd * 10)) & " " & dose(Int(Rnd * 8))
End Function

Private Function FakeID() As Long
    FakeID = 10000 + Int(Rnd * 90000)
End Function

Private Function FakeQty() As Long
    FakeQty = 1 + Int(Rnd * 100)
End Function

' Load counts from LoggedInventory into dicts used by GenerateReport.
' Returns True if any rows were loaded.
Private Function LoadLoggedInventoryDict(ByVal datasetName As String, _
                                         ByRef dictNameToPhys As Object, _
                                         ByRef dictPretty As Object) As Boolean
    Dim ws As Worksheet: Set ws = SheetByName("LoggedInventory")
    LoadLoggedInventoryDict = False
    If ws Is Nothing Then Exit Function

    Dim lastR As Long: lastR = lastRow(ws, 1)
    If lastR < 2 Then Exit Function

    Dim r As Long, nm As String, qty As Variant, key As String
    For r = 2 To lastR
        If StrComp(CStr(ws.Cells(r, 1).Value), datasetName, vbTextCompare) = 0 Then
            nm = NzStr(ws.Cells(r, 2).Value)
            qty = ws.Cells(r, 4).Value
            If Len(nm) > 0 And IsNumeric(qty) Then
                key = NormKey(nm)
                dictNameToPhys(key) = CDbl(qty)
                dictPretty(key) = nm
                LoadLoggedInventoryDict = True
            End If
        End If
    Next r
End Function
