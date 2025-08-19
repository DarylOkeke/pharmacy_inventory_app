Option Explicit

' ===== EXACTLY ONE RULE =====
' Keep only: Active sheet + "HOME" + "About" visible.
' Hide every other sheet (VeryHidden) on every activation.

Private Sub Workbook_Open()
    Application.EnableEvents = True
    EnsureUiSheetsExistAndVisible
    ApplyActiveVisibility
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' Fires whenever the user lands on any sheet (worksheet or chart sheet)
    ApplyActiveVisibility
End Sub

' ---- Core visibility enforcer ----
Public Sub ApplyActiveVisibility()
    Dim s As Object
    Dim activeNM As String
    On Error GoTo done

    activeNM = UCase$(ActiveSheet.Name)

    ' Loop **all** sheets (worksheets + chart sheets)
    Dim anyVisible As Boolean: anyVisible = False
    For Each s In ThisWorkbook.Sheets
        If IsUiName(s.Name) Or UCase$(s.Name) = activeNM Then
            If s.Visible <> xlSheetVisible Then s.Visible = xlSheetVisible
            anyVisible = True
        Else
            If s.Visible <> xlSheetVeryHidden Then s.Visible = xlSheetVeryHidden
        End If
    Next s

    ' Safety: if something went odd and nothing is visible, show HOME
    If Not anyVisible Then EnsureSheetVisible "HOME"
done:
End Sub

' ---- helpers ----
Private Function IsUiName(ByVal nm As String) As Boolean
    nm = UCase$(nm)
    IsUiName = (nm = "HOME" Or nm = "ABOUT")
End Function

Private Sub EnsureUiSheetsExistAndVisible()
    EnsureSheetVisible "HOME"
    EnsureSheetVisible "About"
End Sub

Private Sub EnsureSheetVisible(ByVal sheetName As String)
    Dim s As Object, found As Boolean
    For Each s In ThisWorkbook.Sheets
        If UCase$(s.Name) = UCase$(sheetName) Then
            If s.Visible <> xlSheetVisible Then s.Visible = xlSheetVisible
            found = True
            Exit For
        End If
    Next s
    If Not found Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Worksheets.add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = sheetName
        ws.Visible = xlSheetVisible
    End If
End Sub

