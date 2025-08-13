Option Explicit

' Generate a report comparing physical and expected counts
Sub GenerateReport()
    Dim wsPhys As Worksheet, wsSupp As Worksheet, wsRep As Worksheet
    Dim lo As ListObject, dictID As Object, dictExp As Object
    Dim lastRow As Long, r As Range, dataRow As Long
    Dim drugName As String, physCount As Variant
    Dim expected As Variant, status As String, drugID As String

    Set wsPhys = ThisWorkbook.Sheets("PhysicalCount")
    Set wsSupp = ThisWorkbook.Sheets("SupplierData")
    Set lo = wsPhys.ListObjects("Table13")

    ' build lookup dictionaries from SupplierData
    Set dictID = CreateObject("Scripting.Dictionary")
    Set dictExp = CreateObject("Scripting.Dictionary")
    lastRow = wsSupp.Cells(wsSupp.Rows.Count, "A").End(xlUp).Row
    Dim i As Long
    For i = 9 To lastRow
        drugName = Trim(wsSupp.Cells(i, "A").Value)
        If drugName <> "" Then
            dictID(drugName) = wsSupp.Cells(i, "B").Value
            dictExp(drugName) = wsSupp.Cells(i, "C").Value
        End If
    Next i

    ' recreate Report sheet
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Report").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set wsRep = ThisWorkbook.Worksheets.Add
    wsRep.Name = "Report"

    wsRep.Range("A4:F4").Value = Array("Drug Name", "Drug ID", "Physical Count", "Expected Count", "Status", "Comments")
    dataRow = 5

    For Each r In lo.DataBodyRange.Rows
        drugName = Trim(r.Cells(1).Value)
        If drugName <> "" Then
            physCount = r.Cells(6).Value
            If dictID.Exists(drugName) Then
                drugID = dictID(drugName)
                expected = dictExp(drugName)
                If IsNumeric(physCount) And IsNumeric(expected) Then
                    If physCount > expected Then
                        status = "surplus"
                    ElseIf physCount < expected Then
                        status = "shortage"
                    Else
                        status = "match"
                    End If
                Else
                    status = "N/A"
                End If
            Else
                drugID = ""
                expected = ""
                status = "Not in SupplierData"
            End If

            wsRep.Cells(dataRow, 1).Value = drugName
            wsRep.Cells(dataRow, 2).Value = drugID
            wsRep.Cells(dataRow, 3).Value = physCount
            wsRep.Cells(dataRow, 4).Value = expected
            wsRep.Cells(dataRow, 5).Value = status
            wsRep.Cells(dataRow, 6).Value = ""
            dataRow = dataRow + 1
        End If
    Next r

    wsRep.Columns("A:F").AutoFit
End Sub

' Clear physical counts and timestamps in Table13
Sub ClearInputs()
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("PhysicalCount").ListObjects("Table13")
    lo.ListColumns(6).DataBodyRange.ClearContents
    If lo.ListColumns.Count >= 7 Then
        lo.ListColumns(7).DataBodyRange.ClearContents
    End If
End Sub

' Import supplier data from a workbook containing a SupData sheet
Sub ImportWeeklyData()
    Dim fd As FileDialog, fileName As Variant
    Dim wbSrc As Workbook, wsSrc As Worksheet
    Dim wsDest As Worksheet, wsPhys As Worksheet
    Dim lo As ListObject, lastRow As Long, data As Variant
    Dim i As Long

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    If fd.Show <> -1 Then Exit Sub
    fileName = fd.SelectedItems(1)

    Set wbSrc = Workbooks.Open(fileName)
    On Error GoTo ImportError
    Set wsSrc = wbSrc.Sheets("SupData")

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then GoTo CleanUp
    data = wsSrc.Range("A2:C" & lastRow).Value

    Set wsDest = ThisWorkbook.Sheets("SupplierData")
    wsDest.Range("A9:C" & wsDest.Rows.Count).ClearContents
    wsDest.Range("A9").Resize(UBound(data, 1), UBound(data, 2)).Value = data

    Set wsPhys = ThisWorkbook.Sheets("PhysicalCount")
    Set lo = wsPhys.ListObjects("Table13")
    lo.ListColumns(1).DataBodyRange.ClearContents
    lo.ListColumns(3).DataBodyRange.ClearContents
    For i = lo.ListRows.Count + 1 To UBound(data, 1)
        lo.ListRows.Add
    Next i
    For i = 1 To UBound(data, 1)
        lo.DataBodyRange.Cells(i, 1).Value = data(i, 1)
        lo.DataBodyRange.Cells(i, 3).Value = data(i, 2)
    Next i

CleanUp:
    wbSrc.Close SaveChanges:=False
    Exit Sub
ImportError:
    MsgBox "Selected workbook must contain a sheet named 'SupData'.", vbExclamation
    Resume CleanUp
End Sub

' Ensure PhysicalCount table contains all drugs from SupplierData
Sub UpdatePhysicalCountTable()
    Dim wsSupp As Worksheet, wsPhys As Worksheet
    Dim lo As ListObject, lastRow As Long
    Dim drugName As String, drugID As Variant
    Dim matchRow As Variant, i As Long

    Set wsSupp = ThisWorkbook.Sheets("SupplierData")
    Set wsPhys = ThisWorkbook.Sheets("PhysicalCount")
    Set lo = wsPhys.ListObjects("Table13")

    lastRow = wsSupp.Cells(wsSupp.Rows.Count, "A").End(xlUp).Row
    For i = 9 To lastRow
        drugName = wsSupp.Cells(i, "A").Value
        drugID = wsSupp.Cells(i, "B").Value
        If drugName <> "" Then
            On Error Resume Next
            matchRow = Application.Match(drugName, lo.ListColumns(1).DataBodyRange, 0)
            On Error GoTo 0
            If IsError(matchRow) Or matchRow = 0 Then
                lo.ListRows.Add
                With lo.ListRows(lo.ListRows.Count)
                    .Range(1, 1).Value = drugName
                    .Range(1, 3).Value = drugID
                    .Range(1, 2).Value = "Newly added drug"
                    .Range(1, 4).Value = "Newly added drug"
                    .Range(1, 5).Value = "Newly added drug"
                End With
            Else
                lo.ListColumns(3).DataBodyRange.Cells(matchRow).Value = drugID
            End If
        End If
    Next i
End Sub

' Show the simple filter form
Sub ShowFilterForm()
    UserForm1.Show
End Sub

' Remove all filters from the PhysicalCount table
Sub ClearFilters()
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("PhysicalCount").ListObjects("Table13")
    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    On Error Resume Next
    UserForm1.ClearUserFormInputs
    On Error GoTo 0
End Sub

' Restore events, screen updating and clear validation
Sub RestoreWorkbookFunctionality()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    On Error Resume Next
    With ThisWorkbook.Sheets("PhysicalCount").ListObjects("Table13").ListColumns(6).DataBodyRange
        .Validation.Delete
        .NumberFormat = "General"
    End With
    On Error GoTo 0
End Sub
