Option Explicit

'— Setup TechCounts sheet without data validation (handled by Worksheet_Change event)
Private Sub Workbook_Open()
    ' Remove any existing data validation to prevent conflicts with Worksheet_Change event
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("PhysicalCount")
    On Error Resume Next
    ws.Range("F2:F726").Validation.Delete
    On Error GoTo 0
End Sub

'— Validation function for whole numbers or addition expressions
Function ValidateWholeNumberOrAddition(cellValue As Variant) As Boolean
    Dim testValue As String
    Dim result As Variant
    Dim i As Integer
    
    On Error GoTo ValidationFailed
    
    ' Convert to string and trim
    testValue = Trim(CStr(cellValue))
    
    ' Empty cells are valid (will be handled elsewhere)
    If testValue = "" Then
        ValidateWholeNumberOrAddition = True
        Exit Function
    End If
    
    ' Check if it's a simple whole number
    If IsNumeric(testValue) Then
        If CDbl(testValue) = CLng(CDbl(testValue)) And CDbl(testValue) >= 0 Then
            ValidateWholeNumberOrAddition = True
            Exit Function
        End If
    End If
    
    ' Check if it contains only numbers, plus signs, and spaces
    For i = 1 To Len(testValue)
        If Not (IsNumeric(Mid(testValue, i, 1)) Or Mid(testValue, i, 1) = "+" Or Mid(testValue, i, 1) = " ") Then
            ValidateWholeNumberOrAddition = False
            Exit Function
        End If
    Next i
    
    ' Check if it starts or ends with + (invalid)
    If Left(testValue, 1) = "+" Or Right(testValue, 1) = "+" Then
        ValidateWholeNumberOrAddition = False
        Exit Function
    End If
    
    ' Check for consecutive + signs
    If InStr(testValue, "++") > 0 Then
        ValidateWholeNumberOrAddition = False
        Exit Function
    End If
    
    ' Try to evaluate the expression
    result = Evaluate("=" & testValue)
    
    ' Check if result is a number and a whole number
    If IsNumeric(result) Then
        If CDbl(result) = CLng(CDbl(result)) And CDbl(result) >= 0 Then
            ValidateWholeNumberOrAddition = True
            Exit Function
        End If
    End If
    
ValidationFailed:
    ValidateWholeNumberOrAddition = False
End Function


