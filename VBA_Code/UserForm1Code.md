Option Explicit

Private Sub UserForm_Initialize()
    Dim lo As ListObject, cell As Range
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Set lo = ThisWorkbook.Sheets("PhysicalCount").ListObjects("Table13")
    Me.cboDrugName.Clear
    For Each cell In lo.ListColumns(1).DataBodyRange
        If Trim(cell.Value) <> "" Then
            If Not dict.Exists(cell.Value) Then
                dict.Add cell.Value, True
                Me.cboDrugName.AddItem cell.Value
            End If
        End If
    Next cell
End Sub

Private Sub btnApply_Click()
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("PhysicalCount").ListObjects("Table13")
    If Me.cboDrugName.Value <> "" Then
        lo.Range.AutoFilter Field:=1, Criteria1:=Me.cboDrugName.Value
    End If
    Me.Hide
End Sub

Private Sub btnClear_Click()
    ClearFilters
    Me.Hide
End Sub

Public Sub ClearUserFormInputs()
    Me.cboDrugName.Value = ""
End Sub
