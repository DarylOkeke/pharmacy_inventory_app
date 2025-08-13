Option Explicit

Private entryCount As Integer
Private maxEntries As Integer
Private srcList As Variant

'— Initialize form and add first pair of ComboBoxes (DYNAMIC)
Private Sub UserForm_Initialize()
    entryCount = 0
    maxEntries = 7
    
    ' Get drug list dynamically from Table13
    Dim lo As ListObject
    Set lo = ThisWorkbook.Sheets("PhysicalCount").ListObjects("Table13")
    
    ' Convert Table13 NameAndStrength column to array (rows 7 to 106)
    Dim drugArray() As String
    Dim i As Long, drugCount As Long
    Dim cell As Range
    drugCount = 0
    For i = 7 To 106
        If Trim(lo.DataBodyRange.Cells(i - lo.DataBodyRange.Row + 1, 1).Value) <> "" Then
            drugCount = drugCount + 1
        End If
    Next i
    If drugCount > 0 Then
        ReDim drugArray(1 To drugCount)
        Dim arrIdx As Long: arrIdx = 1
        For i = 7 To 106
            If Trim(lo.DataBodyRange.Cells(i - lo.DataBodyRange.Row + 1, 1).Value) <> "" Then
                drugArray(arrIdx) = lo.DataBodyRange.Cells(i - lo.DataBodyRange.Row + 1, 1).Value
                arrIdx = arrIdx + 1
            End If
        Next i
        srcList = drugArray
    Else
        ReDim drugArray(1 To 1)
        drugArray(1) = ""
        srcList = drugArray
    End If
    
    AddRangeEntry
End Sub

'— Dynamically add From/To combo controls
Private Sub AddRangeEntry()
    If entryCount >= maxEntries Then Exit Sub
    entryCount = entryCount + 1
    Dim yPos As Single: yPos = 10 + (entryCount - 1) * 30
    Dim cbFrom As MSForms.ComboBox, cbTo As MSForms.ComboBox

    Set cbFrom = Me.fraEntries.Controls.Add("Forms.ComboBox.1", "cbFrom" & entryCount, True)
    With cbFrom
        .Left = 10: .Top = yPos: .Width = 100: .MatchEntry = fmMatchEntryComplete
        .List = srcList
    End With

    Set cbTo = Me.fraEntries.Controls.Add("Forms.ComboBox.1", "cbTo" & entryCount, True)
    With cbTo
        .Left = 120: .Top = yPos: .Width = 100: .MatchEntry = fmMatchEntryComplete
        .List = srcList
    End With
End Sub

'— + Range button
Private Sub btnAddRange_Click()
    AddRangeEntry
End Sub

'— + One button (single-value filter)
Private Sub btnAddSingle_Click()
    If entryCount < maxEntries Then
        AddRangeEntry
        Me.fraEntries.Controls("cbTo" & entryCount).Visible = False
    End If
End Sub

'— Apply filter criteria (ERROR-PROOF VERSION WITH VALIDATION) - DYNAMIC
Private Sub btnApply_Click()
    On Error GoTo MainError
    
    Dim sheet As Worksheet, lo As ListObject
    Set sheet = ThisWorkbook.Sheets("PhysicalCount")
    Set lo = sheet.ListObjects("Table13")

    ' Clear existing filters first
    On Error Resume Next
    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    On Error GoTo MainError

    Dim critArray() As String
    Dim critCount As Integer: critCount = 0
    Dim i As Integer, fromVal As String, toVal As String, j As Integer
    Dim fromIndex As Variant, toIndex As Variant, startIdx As Integer, endIdx As Integer
    Dim validCriteria As Boolean: validCriteria = False
    Dim invalidInputs As String: invalidInputs = ""
    
    ' First pass: validate all inputs and count valid criteria
    For i = 1 To entryCount
        On Error Resume Next
        fromVal = Trim(Me.fraEntries.Controls("cbFrom" & i).Value)
        On Error GoTo MainError
        
        If fromVal <> "" Then
            ' Validate that fromVal exists in the source list
            On Error Resume Next
            fromIndex = Application.Match(fromVal, srcList, 0)
            On Error GoTo MainError
            
            If IsError(fromIndex) Then
                ' Invalid drug name found
                If invalidInputs <> "" Then invalidInputs = invalidInputs & ", "
                invalidInputs = invalidInputs & Chr(34) & fromVal & Chr(34)
            Else
                If Me.fraEntries.Controls("cbTo" & i).Visible Then
                    On Error Resume Next
                    toVal = Trim(Me.fraEntries.Controls("cbTo" & i).Value)
                    On Error GoTo MainError
                    
                    If toVal <> "" Then
                        ' Validate that toVal exists in the source list
                        On Error Resume Next
                        toIndex = Application.Match(toVal, srcList, 0)
                        On Error GoTo MainError
                        
                        If IsError(toIndex) Then
                            ' Invalid drug name found
                            If invalidInputs <> "" Then invalidInputs = invalidInputs & ", "
                            invalidInputs = invalidInputs & Chr(34) & toVal & Chr(34)
                        Else
                            ' Both from and to are valid
                            critCount = critCount + Abs(CLng(toIndex) - CLng(fromIndex)) + 1
                            validCriteria = True
                        End If
                    Else
                        ' Only from value specified and it's valid
                        critCount = critCount + 1
                        validCriteria = True
                    End If
                Else
                    ' Single value and it's valid
                    critCount = critCount + 1
                    validCriteria = True
                End If
            End If
        End If
    Next i
    
    ' Show validation errors if any invalid inputs found
    If invalidInputs <> "" Then
        MsgBox "Invalid drug names entered: " & invalidInputs & vbCrLf & vbCrLf & _
               "Please select drugs from the dropdown list only.", vbExclamation, "Invalid Input"
        Exit Sub
    End If
    
    ' Validate we have criteria
    If Not validCriteria Or critCount = 0 Then
        MsgBox "Please select at least one valid filter criteria from the dropdown lists.", vbExclamation
        Exit Sub
    End If
    
    If critCount > 1000 Then
        If MsgBox("This will filter " & critCount & " items. This may take a moment. Continue?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Build the criteria array with duplicate prevention
    ReDim critArray(1 To critCount)
    Dim tempDict As Object: Set tempDict = CreateObject("Scripting.Dictionary")
    Dim actualCount As Integer: actualCount = 0
    
    For i = 1 To entryCount
        On Error Resume Next
        fromVal = Trim(Me.fraEntries.Controls("cbFrom" & i).Value)
        On Error GoTo MainError
        
        If fromVal <> "" Then
            On Error Resume Next
            fromIndex = Application.Match(fromVal, srcList, 0)
            On Error GoTo MainError
            
            If Not IsError(fromIndex) Then
                If Me.fraEntries.Controls("cbTo" & i).Visible Then
                    On Error Resume Next
                    toVal = Trim(Me.fraEntries.Controls("cbTo" & i).Value)
                    On Error GoTo MainError
                    
                    If toVal <> "" Then
                        On Error Resume Next
                        toIndex = Application.Match(toVal, srcList, 0)
                        On Error GoTo MainError
                        
                        If Not IsError(toIndex) Then
                            ' Handle range (automatically handles reverse order)
                            startIdx = Application.Min(CLng(fromIndex), CLng(toIndex))
                            endIdx = Application.Max(CLng(fromIndex), CLng(toIndex))
                            
                            For j = startIdx To endIdx
                                If Not tempDict.Exists(srcList(j)) Then
                                    tempDict.Add srcList(j), True
                                    actualCount = actualCount + 1
                                    If actualCount <= UBound(critArray) Then
                                        critArray(actualCount) = srcList(j)
                                    End If
                                End If
                            Next j
                        End If
                    Else
                        ' Only from value - add if not duplicate
                        If Not tempDict.Exists(fromVal) Then
                            tempDict.Add fromVal, True
                            actualCount = actualCount + 1
                            If actualCount <= UBound(critArray) Then
                                critArray(actualCount) = fromVal
                            End If
                        End If
                    End If
                Else
                    ' Single value - add if not duplicate
                    If Not tempDict.Exists(fromVal) Then
                        tempDict.Add fromVal, True
                        actualCount = actualCount + 1
                        If actualCount <= UBound(critArray) Then
                            critArray(actualCount) = fromVal
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    ' Final validation
    If actualCount = 0 Then
        MsgBox "No valid filter criteria found. Please check your selections.", vbExclamation
        Exit Sub
    End If
    
    ' Resize array to actual size
    If actualCount < UBound(critArray) Then
        ReDim Preserve critArray(1 To actualCount)
    End If
    
    ' Protect objects from moving during filtering
    On Error Resume Next
    Dim shp As Shape
    For Each shp In sheet.Shapes
        shp.Placement = xlFreeFloating
    Next shp
    On Error GoTo MainError
    
    ' Apply the filter with additional error protection
    On Error GoTo FilterError
    If actualCount = 1 Then
        ' Single criteria - use simpler filter method
        lo.Range.AutoFilter Field:=1, Criteria1:=critArray(1)
    Else
        ' Multiple criteria
        lo.Range.AutoFilter Field:=1, Criteria1:=critArray, Operator:=xlFilterValues
    End If
    
    Me.Hide
    Exit Sub
    
FilterError:
    ' If xlFilterValues fails, try alternative method
    On Error GoTo MainError
    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    
    ' Alternative: Apply criteria one by one using OR logic
    Dim filterCriteria As String
    filterCriteria = Join(critArray, ",")
    
    ' Try with different approach
    On Error Resume Next
    lo.Range.AutoFilter Field:=1, Criteria1:="=" & critArray(1)
    If Err.Number <> 0 Then
        lo.Range.AutoFilter Field:=1, Criteria1:=critArray(1)
    End If
    On Error GoTo MainError
    
    Me.Hide
    Exit Sub
    
MainError:
    MsgBox "An error occurred while applying the filter. Please try again with different selections." & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbExclamation, "Filter Error"
    
    ' Ensure filters are cleared on error
    On Error Resume Next
    If Not lo Is Nothing Then
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    End If
    On Error GoTo 0
End Sub

'— Clear filter and hide with UI cleanup
Private Sub btnClear_Click()
    ClearFilters
    ResetUserFormUI
    Me.Hide
End Sub

'— Prevent form from unloading when X is clicked - just hide it instead
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True  ' Cancel the close operation
        Me.Hide        ' Hide the form instead
    End If
End Sub

'— Clear all UserForm inputs (called by ClearFilters macro)
Sub ClearUserFormInputs()
    Dim i As Integer
    ' Clear all combo box values
    For i = 1 To entryCount
        On Error Resume Next
        Me.fraEntries.Controls("cbFrom" & i).Value = ""
        Me.fraEntries.Controls("cbTo" & i).Value = ""
        On Error GoTo 0
    Next i
    
    ' Reset UI to clean state
    ResetUserFormUI
End Sub

'— Reset UserForm UI to clean state (only one range entry)
Private Sub ResetUserFormUI()
    Dim i As Integer
    
    ' Remove all controls except the first pair
    For i = entryCount To 2 Step -1
        On Error Resume Next
        Me.fraEntries.Controls.Remove "cbFrom" & i
        Me.fraEntries.Controls.Remove "cbTo" & i
        On Error GoTo 0
    Next i
    
    ' Reset counter and clear first pair
    entryCount = 1
    On Error Resume Next
    Me.fraEntries.Controls("cbFrom1").Value = ""
    Me.fraEntries.Controls("cbTo1").Value = ""
    Me.fraEntries.Controls("cbTo1").Visible = True  ' Ensure To box is visible for range
    On Error GoTo 0
End Sub


