''— Find the row in SupplierData matching the drug name
Function FindDrugRowInWeeklyExpected(drugName As String, wsExpected As Worksheet) As Long
    Dim lastRow As Long, i As Long
    lastRow = wsExpected.Cells(wsExpected.Rows.Count, 1).End(xlUp).Row
    For i = 9 To lastRow
        If Trim(UCase(wsExpected.Cells(i, 1).Value)) = Trim(UCase(drugName)) Then
            FindDrugRowInWeeklyExpected = i
            Exit Function
        End If
    Next i
    FindDrugRowInWeeklyExpected = 0 ' Not found
End Function
'— COMPLETE DYNAMIC BULLETPROOF MODULE1 CODE
Option Explicit

' Constants for sheet names - protects against accidental renaming
Private Const SHEET_WEEKLYEXPECTED As String = "SupplierData"
Private Const SHEET_TECHCOUNTS As String = "PhysicalCount"
Private Const SHEET_EXTERNAL As String = "Current Product Inventory Overv"
Private Const TABLE_NAME As String = "Table13"

' Dynamic range constants - will be updated during import
Private startRow As Long
Private EndRow As Long

'— Validate critical sheets exist (prevents crashes from accidental renames)
Function ValidateSheets() As Boolean
    On Error Resume Next
    Dim wsE As Worksheet, wsT As Worksheet
    Dim lo As ListObject
    Set wsE = ThisWorkbook.Sheets(SHEET_WEEKLYEXPECTED)
    Set wsT = ThisWorkbook.Sheets(SHEET_TECHCOUNTS)
    Set lo = wsT.ListObjects(TABLE_NAME)
    On Error GoTo 0
    Dim debugMsg As String
    debugMsg = ""
    If wsE Is Nothing Then debugMsg = debugMsg & "SupplierData sheet not found." & vbCrLf
    If wsT Is Nothing Then debugMsg = debugMsg & "PhysicalCount sheet not found." & vbCrLf
    If lo Is Nothing Then
        debugMsg = debugMsg & "Table13 not found in PhysicalCount sheet." & vbCrLf
        ' List all tables found on PhysicalCount
        Dim tblList As String
        tblList = "Tables found on PhysicalCount:" & vbCrLf
        Dim tbl As ListObject
        For Each tbl In wsT.ListObjects
            tblList = tblList & "- " & tbl.Name & vbCrLf
        Next tbl
        debugMsg = debugMsg & tblList
    End If
    If debugMsg <> "" Then
        MsgBox "Critical Error: Required sheets or tables are missing or renamed." & vbCrLf & vbCrLf & debugMsg & vbCrLf & _
               "Please ensure these exist:" & vbCrLf & _
               "• " & SHEET_TECHCOUNTS & " sheet" & vbCrLf & _
               "• " & SHEET_WEEKLYEXPECTED & " sheet" & vbCrLf & _
               "• " & TABLE_NAME & " table in " & SHEET_TECHCOUNTS, vbCritical, "Workbook Structure Error"
        ValidateSheets = False
        Exit Function
    End If
    ValidateSheets = True
End Function

'— Get dynamic data range from WeeklyExpected sheet
Function GetWeeklyExpectedDataRange() As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Sheets(SHEET_WEEKLYEXPECTED)
    
    ' Find last row with data in column A starting from row 10
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Ensure we have at least row 10
    If lastRow < 10 Then lastRow = 10
    
    GetWeeklyExpectedDataRange = lastRow
End Function

...existing code...

'— Enhanced Main report generator with bulletproof error handling and drug name matching (DYNAMIC)
Sub GenerateReport()
    If Not ValidateSheets() Then Exit Sub

    ' Get custom report title from user
    Dim reportTitle As String
    reportTitle = InputBox("What would you like to title your report?", "Report Title", "Inventory Reconciliation Report")
    ' Check if user cancelled (InputBox returns empty string when cancelled)
    If reportTitle = "" Then
        MsgBox "Report generation cancelled.", vbInformation
        Exit Sub
    End If
    reportTitle = Trim(reportTitle) ' Clean up any extra spaces

    Dim wsE As Worksheet, wsT As Worksheet, wsR As Worksheet
    Dim r As Long, diff As Double
    Dim noMatchCount As Integer
    Dim noMatchList As String
    Set wsE = ThisWorkbook.Sheets(SHEET_WEEKLYEXPECTED)
    Set wsT = ThisWorkbook.Sheets(SHEET_TECHCOUNTS)
    ' Remove any previous report
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Report").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set wsR = ThisWorkbook.Sheets.Add(After:=wsT)
    wsR.Name = "Report"
    wsR.Range("A1").Value = reportTitle
    wsR.Range("A1").Font.Size = 16
    wsR.Range("A1").Font.Bold = True
    wsR.Range("A1").Font.Name = "Calibri"
    wsR.Range("A1").HorizontalAlignment = xlCenter
    wsR.Range("A1:F1").Merge
    wsR.Range("A4:F4").Value = Array("Drug Name", "Drug ID", "Physical Count", "Expected Count", "Status", "Comments")
    r = 5
    noMatchCount = 0
    noMatchList = ""
    Dim iT As Long, drugName As String, physCount As Variant, expectedCount As Variant, rowE As Long
    For iT = 7 To 106
        drugName = Trim(wsT.Cells(iT, 1).Value)
        physCount = wsT.Cells(iT, 6).Value
        If drugName <> "" And physCount <> "" Then
            rowE = 0
            For rowE = 9 To wsE.Cells(wsE.Rows.Count, 1).End(xlUp).Row
                If Trim(UCase(wsE.Cells(rowE, 1).Value)) = Trim(UCase(drugName)) Then Exit For
            Next rowE
            If rowE > 0 And Trim(UCase(wsE.Cells(rowE, 1).Value)) = Trim(UCase(drugName)) Then
                expectedCount = wsE.Cells(rowE, 3).Value
                wsR.Cells(r, "A").Value = drugName
                wsR.Cells(r, "B").Value = wsE.Cells(rowE, 2).Value
                wsR.Cells(r, "C").Value = physCount
                wsR.Cells(r, "D").Value = expectedCount
                diff = Val(physCount) - Val(expectedCount)
                If diff > 0 Then
                    wsR.Cells(r, "E").Value = "surplus of " & diff
                ElseIf diff < 0 Then
                    wsR.Cells(r, "E").Value = "shortage of " & Abs(diff)
                Else
                    wsR.Cells(r, "E").Value = ""
                End If
                wsR.Cells(r, "F").Value = ""
                r = r + 1
            Else
                noMatchCount = noMatchCount + 1
                If noMatchList <> "" Then noMatchList = noMatchList & ", "
                noMatchList = noMatchList & Chr(34) & drugName & Chr(34)
            End If
        End If
    Next iT
    If r = 5 Then
        MsgBox "No valid data found to create report." & vbCrLf & vbCrLf & _
               IIf(noMatchCount > 0, "Drugs with no match in WeeklyExpected: " & noMatchList, ""), vbExclamation
        Application.DisplayAlerts = False
        wsR.Delete
        Application.DisplayAlerts = True
        Exit Sub
    End If
    Dim lo As ListObject
    Set lo = wsR.ListObjects.Add(SourceType:=xlSrcRange, Source:=wsR.Range("A4").CurrentRegion, XlListObjectHasHeaders:=xlYes)
    lo.Name = "tblReport"
    lo.TableStyle = "TableStyleLight9"
    wsR.Columns.AutoFit
    On Error Resume Next
    With lo.ListColumns("Status").DataBodyRange
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlTextString, String:="surplus", TextOperator:=xlContains
        .FormatConditions(1).Interior.Color = RGB(0, 176, 80)
        .FormatConditions.Add Type:=xlTextString, String:="shortage", TextOperator:=xlContains
        .FormatConditions(2).Interior.Color = RGB(255, 170, 51)
    End With
    On Error GoTo 0
    Dim totalRows As Long, startRow As Long
    totalRows = lo.ListRows.Count
    startRow = lo.Range.Row + lo.Range.Rows.Count + 1
    wsR.Cells(startRow, "A").Value = "Total Items:"
    wsR.Cells(startRow, "B").Value = totalRows
    wsR.Cells(startRow + 1, "A").Value = "Generated on:"
    wsR.Cells(startRow + 1, "B").Value = Format(Now, "m/d/yyyy h:nn AM/PM")
    If noMatchCount > 0 Then
        wsR.Cells(startRow + 3, "A").Value = "Warning - Drugs not found in WeeklyExpected:"
        wsR.Cells(startRow + 3, "A").Font.Bold = True
        wsR.Cells(startRow + 3, "A").Font.Color = RGB(255, 0, 0)
        wsR.Cells(startRow + 4, "A").Value = noMatchList
        wsR.Cells(startRow + 4, "A").Font.Color = RGB(255, 0, 0)
    End If
    wsR.Columns("B").AutoFit
    wsR.Columns("A").AutoFit
    wsR.Activate
    Dim completionMsg As String
    completionMsg = "Report is ready on the 'Report' sheet!"
    If noMatchCount > 0 Then
        completionMsg = completionMsg & vbCrLf & vbCrLf & _
                       "Warning: " & noMatchCount & " drug(s) could not be matched with WeeklyExpected data." & vbCrLf & _
                       "These drugs were skipped from the report. Check spelling and data consistency."
    End If
    MsgBox completionMsg, IIf(noMatchCount > 0, vbExclamation, vbInformation)
    Exit Sub
End Sub

'— FULLY DYNAMIC Import weekly data with automatic table updating
Sub ImportWeeklyData()
    On Error GoTo ImportError
    If Not ValidateSheets() Then Exit Sub

    Dim fd As FileDialog, wbSrc As Workbook
    Dim wsSrc As Worksheet, wsDest As Worksheet
    Dim filePath As String
    Dim headerRow As Long, srcLastRow As Long
    Dim importedRows As Long
    Dim i As Long, j As Long

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        .Title = "Select supplier file"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls;*.xlsx;*.xlsm"
        If .Show <> -1 Then Exit Sub
        filePath = .SelectedItems(1)
    End With

    If Dir(filePath) = "" Then
        MsgBox "Selected file no longer exists.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ThisWorkbook.Sheets(SHEET_WEEKLYEXPECTED).ImportInProgress = True

    Set wbSrc = Workbooks.Open(Filename:=filePath, ReadOnly:=True)
    Set wsSrc = wbSrc.Sheets(1) ' Use the first sheet
    Set wsDest = ThisWorkbook.Sheets(SHEET_WEEKLYEXPECTED)

    ' Find header row
    headerRow = 0
    For i = 1 To 20
        If Trim(UCase(wsSrc.Cells(i, 1).Value)) = "DRUG NAME" And _
           Trim(UCase(wsSrc.Cells(i, 2).Value)) = "DRUG ID" And _
           Trim(UCase(wsSrc.Cells(i, 3).Value)) = "EXPECTED QUANTITY" Then
            headerRow = i
            Exit For
        End If
    Next i
    If headerRow = 0 Then
        MsgBox "Header row not found. Please check the supplier file format.", vbExclamation
        GoTo ImportCleanup
    End If

    ' Find last row with data in column 1
    srcLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    If srcLastRow <= headerRow Then
        MsgBox "No data found after header row.", vbExclamation
        GoTo ImportCleanup
    End If

    ' Clear only columns A and C from row 9 down
    Dim destLastRow As Long
    destLastRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row
    If destLastRow >= 9 Then
        wsDest.Range("A9:A" & destLastRow).ClearContents
        wsDest.Range("C9:C" & destLastRow).ClearContents
    End If

    ' Import data from supplier file
    importedRows = 0
    j = 9
    ' Clear only columns A and C in Table13
    Dim wsTech As Worksheet
    Dim loTech As ListObject
    Set wsTech = ThisWorkbook.Sheets(SHEET_TECHCOUNTS)
    Set loTech = wsTech.ListObjects(TABLE_NAME)
    If loTech.ListRows.Count > 0 Then
        loTech.ListColumns(1).DataBodyRange.ClearContents ' Column A: Drug Name
        loTech.ListColumns(3).DataBodyRange.ClearContents ' Column C: Drug ID
    End If
    Dim techRow As Long
    techRow = 1
    For i = headerRow + 1 To srcLastRow
        If Trim(wsSrc.Cells(i, 1).Value) <> "" Then
            wsDest.Cells(j, 1).Value = wsSrc.Cells(i, 1).Value ' Drug Name
            wsDest.Cells(j, 2).Value = wsSrc.Cells(i, 2).Value ' Drug ID
            wsDest.Cells(j, 3).Value = wsSrc.Cells(i, 3).Value ' Expected Quantity
            ' Add row to TechCounts Table13
            If techRow > loTech.ListRows.Count Then
                loTech.ListRows.Add
            End If
            loTech.DataBodyRange.Cells(techRow, 1).Value = wsSrc.Cells(i, 1).Value ' Table13 Col A: Drug Name
            loTech.DataBodyRange.Cells(techRow, 3).Value = wsSrc.Cells(i, 2).Value ' Table13 Col C: Drug ID
            importedRows = importedRows + 1
            j = j + 1
            techRow = techRow + 1
        End If
    Next i

ImportCleanup:
    On Error Resume Next
    ThisWorkbook.Sheets(SHEET_WEEKLYEXPECTED).ImportInProgress = False
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Import completed! " & importedRows & " records imported.", vbInformation, "Import Complete"
    Exit Sub

ImportError:
    MsgBox "Error importing data: " & Err.Description & vbCrLf & vbCrLf & _
           "Please verify the file format and try again.", vbExclamation
    GoTo ImportCleanup
End Sub

'— Update TechCounts Table13 to match WeeklyExpected drug list
Function UpdateTechCountsTable(wsExpected As Worksheet, wsTech As Worksheet, lastRow As Long) As Long
    On Error GoTo UpdateError
    
    Dim lo As ListObject
    Dim i As Long, newDrugsCount As Long
    Dim drugName As String, ndcCode As String
    Dim existingDrugs As Object
    Dim existingRow As Long
    
    Set lo = wsTech.ListObjects(TABLE_NAME)
    Set existingDrugs = CreateObject("Scripting.Dictionary")
    
    ' Build dictionary of existing drugs in Table13
    Dim cell As Range
    For Each cell In lo.ListColumns("NameAndStrength").DataBodyRange
        If cell.Value <> "" Then
            existingDrugs(Trim(UCase(cell.Value))) = cell.Row
        End If
    Next cell
    
    newDrugsCount = 0
    
    ' Loop through WeeklyExpected data and ensure all drugs exist in Table13
    For i = 9 To lastRow
        drugName = Trim(wsExpected.Cells(i, 1).Value) ' Column A
        ndcCode = Trim(wsExpected.Cells(i, 2).Value)  ' Column B (Drug ID)
        ' Expected Quantity is now in column C if needed elsewhere
        If drugName <> "" Then
            ' Check if drug already exists in Table13
            If Not existingDrugs.Exists(Trim(UCase(drugName))) Then
                ' Drug doesn't exist - add new row to Table13
                Dim newRow As ListRow
                Set newRow = lo.ListRows.Add
                ' Populate new row
                newRow.Range(1, 1).Value = drugName                    ' Column A: NameAndStrength
                newRow.Range(1, 2).Value = "Newly added drug"          ' Column B: Canister
                newRow.Range(1, 3).Value = ndcCode                     ' Column C: Drug ID
                newRow.Range(1, 4).Value = "Newly added drug"          ' Column D: Description
                newRow.Range(1, 5).Value = "Newly added drug"          ' Column E: Manufacturer
                ' Columns F and G (Count and Date Logged) remain empty for user input
                newDrugsCount = newDrugsCount + 1
                ' Add to dictionary to avoid duplicates
                existingDrugs(Trim(UCase(drugName))) = newRow.Range.Row
            Else
                ' Drug exists - update Drug ID in case it changed
                existingRow = existingDrugs(Trim(UCase(drugName)))
                wsTech.Cells(existingRow, 3).Value = ndcCode ' Update Column C with new Drug ID
            End If
        End If
    Next i
    
    ' Sort the table alphabetically by NameAndStrength column (A-Z) if new drugs were added
    If newDrugsCount > 0 And lo.ListRows.Count > 0 Then
        On Error Resume Next
        lo.Sort.SortFields.Clear
        lo.Sort.SortFields.Add Key:=lo.ListColumns("NameAndStrength").Range, _
                              SortOn:=xlSortOnValues, _
                              Order:=xlAscending, _
                              DataOption:=xlSortNormal
        With lo.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        On Error GoTo UpdateError
    End If
    
    UpdateTechCountsTable = newDrugsCount
    Exit Function
    
UpdateError:
    MsgBox "Error updating TechCounts table: " & Err.Description, vbExclamation
    UpdateTechCountsTable = 0
End Function

'— Enhanced Clear all inputs with error protection (DYNAMIC)
Sub ClearInputs()
    On Error GoTo ClearError
    
    If Not ValidateSheets() Then Exit Sub

    ' Set flag to bypass protection during programmatic clearing
    ThisWorkbook.Sheets(SHEET_TECHCOUNTS).ClearInProgress = True

    ' Clear the Count and Timestamp columns in Table13 (DYNAMIC)
    Dim lo As ListObject
    Dim physCountColumn As ListColumn
    Dim timestampColumn As ListColumn

    Set lo = ThisWorkbook.Sheets(SHEET_TECHCOUNTS).ListObjects(TABLE_NAME)

    ' Get Count column (dynamic detection)
    On Error Resume Next
    Set physCountColumn = lo.ListColumns("Count")
    If physCountColumn Is Nothing Then
        ' Fallback - try column 6 (F)
        If lo.ListColumns.Count >= 6 Then
            Set physCountColumn = lo.ListColumns(6)
        End If
    End If
    On Error GoTo ClearError

    ' Get Timestamp column (dynamic detection)
    On Error Resume Next
    If lo.ListColumns.Count >= 7 Then
        Set timestampColumn = lo.ListColumns(7) ' Timestamp column
    End If
    On Error GoTo ClearError

    ' Clear Count column if found
    If Not physCountColumn Is Nothing Then
        physCountColumn.DataBodyRange.ClearContents
    End If

    ' Clear Timestamp column if found
    If Not timestampColumn Is Nothing Then
        timestampColumn.DataBodyRange.ClearContents
    End If

    ' Clear the flag
    ThisWorkbook.Sheets(SHEET_TECHCOUNTS).ClearInProgress = False

    MsgBox "Inputs cleared.", vbInformation
    Exit Sub
    
ClearError:
    ' Ensure flag is cleared on error
    On Error Resume Next
    ThisWorkbook.Sheets(SHEET_TECHCOUNTS).ClearInProgress = False
    On Error GoTo 0
    MsgBox "Error clearing inputs: " & Err.Description, vbExclamation
End Sub

'— Enhanced Show the filter UserForm
Sub ShowFilterForm()
    On Error GoTo ShowFormError
    
    If Not ValidateSheets() Then Exit Sub
    UserForm1.Show
    Exit Sub
    
ShowFormError:
    MsgBox "Error opening filter form: " & Err.Description, vbExclamation
End Sub

'— Enhanced Clear filters with error protection
Sub ClearFilters()
    On Error GoTo ClearFilterError
    
    If Not ValidateSheets() Then Exit Sub
    
    With ThisWorkbook.Sheets(SHEET_TECHCOUNTS).ListObjects(TABLE_NAME)
        If .AutoFilter.FilterMode Then .AutoFilter.ShowAllData
    End With
    
    ' Also clear the UserForm inputs
    On Error Resume Next
    UserForm1.ClearUserFormInputs
    On Error GoTo 0
    Exit Sub
    
ClearFilterError:
    MsgBox "Error clearing filters: " & Err.Description, vbExclamation
End Sub

'— EMERGENCY RECOVERY: Fix disabled events and restore functionality
Sub RestoreWorkbookFunctionality()
    On Error Resume Next
    
    ' Re-enable critical Excel settings
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' Clear any import flags that might be stuck
    ThisWorkbook.Sheets(SHEET_TECHCOUNTS).ImportInProgress = False
    ThisWorkbook.Sheets(SHEET_WEEKLYEXPECTED).ImportInProgress = False
    ThisWorkbook.Sheets(SHEET_TECHCOUNTS).ClearInProgress = False
    
    ' Remove any problematic data validation (Worksheet_Change event handles validation)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PhysicalCount")
    
    ' Get dynamic range for validation clearing
    Dim lo As ListObject
    Set lo = ws.ListObjects(TABLE_NAME)
    lo.ListColumns("Count").DataBodyRange.Validation.Delete
    
    ' Fix formatting issues that cause date conversion problems
    lo.ListColumns("Count").DataBodyRange.NumberFormat = "General"
    lo.ListColumns("Count").DataBodyRange.ClearFormats
    
    On Error GoTo 0
    MsgBox "Workbook functionality restored!" & vbCrLf & vbCrLf & _
           "• Events re-enabled" & vbCrLf & _
           "• Conflicting data validation removed" & vbCrLf & _
           "• Import/Clear flags cleared" & vbCrLf & _
           "• Cell formatting reset to prevent date conversion" & vbCrLf & _
           "• Validation now handled by worksheet events only", vbInformation, "Recovery Complete"
End Sub


