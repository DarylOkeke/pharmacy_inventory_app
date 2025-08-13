Option Explicit

' Flag to bypass protection during imports
Public ImportInProgress As Boolean
' Flag to bypass protection during programmatic clearing
Public ClearInProgress As Boolean

'— Enhanced entry handling with validation for expressions and timestamp on TechCounts sheet (DYNAMIC)
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rng As Range
    Dim originalValue As Variant
    Dim evaluatedValue As Variant
    Dim isValid As Boolean
    Dim lo As ListObject
    Dim physCountColumn As ListColumn
    Dim timestampColumn As ListColumn
    Dim timestampCell As Range
    
    ' Ensure events are enabled (recovery from errors)
    If Not Application.EnableEvents Then
        Application.EnableEvents = True
        Exit Sub
    End If
    
    ' Get Table13 reference
    Set lo = Me.ListObjects("Table13")
    
    ' Get Count column (dynamic detection)
    On Error Resume Next
    Set physCountColumn = lo.ListColumns("Count")
    If physCountColumn Is Nothing Then
        ' Fallback - try column F
        Set physCountColumn = lo.ListColumns(6)
    End If
    On Error GoTo 0
    
    ' Get Timestamp column (dynamic detection)
    On Error Resume Next
    Set timestampColumn = lo.ListColumns(7) ' Typically column G
    On Error GoTo 0
    
    ' Removed protection popup and logic for editing columns except Count. All edits now allowed without prompt.
    
    ' Handle count entries in Count column with enhanced error protection
    If Not physCountColumn Is Nothing And Not Intersect(Target, physCountColumn.DataBodyRange) Is Nothing Then
        Application.EnableEvents = False
        
        For Each rng In Intersect(Target, physCountColumn.DataBodyRange)
            If rng.Value <> "" Then
                originalValue = rng.Value
                isValid = False
                
                ' Enhanced validation - check for invalid data types first
                On Error GoTo ValidationError
                
                ' Convert to string for validation
                Dim strValue As String
                strValue = Trim(CStr(originalValue))
                
                ' SPECIAL FIX: If Excel converted input to date, extract the original number
                If VarType(originalValue) = vbDate Then
                    ' Excel converted a number to date - try to get back the original number
                    ' For dates like 1/22/1900, the day part is likely the intended number
                    Dim dayPart As Integer
                    dayPart = Day(originalValue)
                    strValue = CStr(dayPart)
                    
                    ' Clear any date formatting and set to General
                    rng.NumberFormat = "General"
                End If
                
                ' Reject if it's a range reference, formula, or contains invalid characters
                If Left(strValue, 1) = "=" Or InStr(strValue, ":") > 0 Or InStr(strValue, "-") > 0 Or InStr(strValue, "*") > 0 Or InStr(strValue, "/") > 0 Or InStr(strValue, "^") > 0 Then
                    GoTo ValidationError
                End If
                
                ' Check if it's a simple whole number first
                If IsNumeric(strValue) Then
                    evaluatedValue = CDbl(strValue)
                    If evaluatedValue = CLng(evaluatedValue) And evaluatedValue >= 0 Then
                        rng.Value = CLng(evaluatedValue)
                        isValid = True
                    End If
                ElseIf ValidateExpressionFormat(strValue) Then
                    ' Try to evaluate as addition expression
                    evaluatedValue = Evaluate("=" & strValue)
                    If IsError(evaluatedValue) Then
                        GoTo ValidationError
                    End If
                    If IsNumeric(evaluatedValue) Then
                        If evaluatedValue = CLng(evaluatedValue) And evaluatedValue >= 0 Then
                            rng.Value = CLng(evaluatedValue)
                            isValid = True
                        End If
                    End If
                End If
                
                On Error GoTo 0
                
                ' Only add timestamp if validation passed
                If isValid And Not timestampColumn Is Nothing Then
                    ' Temporarily enable import flag to prevent protection prompt for timestamp
                    ImportInProgress = True
                    ' Find corresponding timestamp cell in the same row
                    Set timestampCell = timestampColumn.DataBodyRange.Cells(rng.Row - timestampColumn.DataBodyRange.Row + 1, 1)
                    timestampCell.Value = Format(Now, "m/d h:nn AM/PM")
                    timestampColumn.DataBodyRange.EntireColumn.AutoFit
                    ImportInProgress = False
                Else
                    GoTo ValidationError
                End If
            Else
                ' If count value is empty/deleted, clear the corresponding timestamp
                If Not timestampColumn Is Nothing Then
                    ' Temporarily enable import flag to prevent protection prompt
                    ImportInProgress = True
                    Set timestampCell = timestampColumn.DataBodyRange.Cells(rng.Row - timestampColumn.DataBodyRange.Row + 1, 1)
                    timestampCell.Value = ""
                    ImportInProgress = False
                End If
                ' Also ensure the cell formatting is set to General for count column
                rng.NumberFormat = "General"
            End If
        Next rng
        
        Application.EnableEvents = True
        Exit Sub
        
ValidationError:
        ' Enhanced error handling with recovery
        Application.EnableEvents = True
        rng.Value = ""
        ' Clear timestamp without triggering protection
        If Not timestampColumn Is Nothing Then
            ImportInProgress = True
            Set timestampCell = timestampColumn.DataBodyRange.Cells(rng.Row - timestampColumn.DataBodyRange.Row + 1, 1)
            timestampCell.Value = ""
            ImportInProgress = False
        End If
        
        MsgBox "Invalid input detected!" & vbCrLf & vbCrLf & _
               "The following are NOT allowed:" & vbCrLf & _
               "• Formulas (starting with =)" & vbCrLf & _
               "• Range references (e.g., A1:B5)" & vbCrLf & _
               "• Decimals or negative numbers" & vbCrLf & _
               "• Subtraction or other operations" & vbCrLf & vbCrLf & _
               "Please enter only:" & vbCrLf & _
               "• Whole numbers (e.g., 150)" & vbCrLf & _
               "• Addition expressions (e.g., 50+100+25)", _
               vbExclamation, "Invalid Count Entry"
        
        rng.Select
        Exit Sub
    End If
End Sub

'— Validate expression format (only numbers, +, and spaces allowed)
Private Function ValidateExpressionFormat(expression As String) As Boolean
    Dim i As Integer
    Dim char As String
    
    expression = Trim(expression)
    
    ' Check for empty string
    If Len(expression) = 0 Then
        ValidateExpressionFormat = False
        Exit Function
    End If
    
    ' Check each character
    For i = 1 To Len(expression)
        char = Mid(expression, i, 1)
        If Not (IsNumeric(char) Or char = "+" Or char = " ") Then
            ValidateExpressionFormat = False
            Exit Function
        End If
    Next i
    
    ' Check if it starts or ends with +
    If Left(expression, 1) = "+" Or Right(expression, 1) = "+" Then
        ValidateExpressionFormat = False
        Exit Function
    End If
    
    ' Check for consecutive + signs
    If InStr(expression, "++") > 0 Then
        ValidateExpressionFormat = False
        Exit Function
    End If
    
    ValidateExpressionFormat = True
End Function


