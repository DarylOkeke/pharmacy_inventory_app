
Option Explicit

" Flag to bypass protection during imports
Public ImportInProgress As Boolean

'— Protection for WeeklyExpected sheet columns A-K
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Check if user is trying to edit protected columns A-K (unless import is in progress)
    If Not ImportInProgress And Not Intersect(Target, Me.Range("A:K")) Is Nothing Then
        Dim originalValue As Variant
        originalValue = Target.Value ' Store the value before undoing
        
        Application.EnableEvents = False
        Application.Undo
        Application.EnableEvents = True
        
        If MsgBox("Are you sure you want to edit this? These values shouldn't be changed unless you are sure you found an error.", _
                  vbYesNo + vbExclamation, "Edit Protected Data") = vbNo Then
            Exit Sub
        Else
            ' User confirmed, allow the edit by applying the stored value
            Application.EnableEvents = False
            Target.Value = originalValue ' Apply the original change
            Application.EnableEvents = True
        End If
    End If
End Sub



