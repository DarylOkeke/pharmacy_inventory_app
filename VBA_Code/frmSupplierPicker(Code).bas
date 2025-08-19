Option Explicit

' Exposed results the caller reads after Me.Show
Public UserChoice As String          ' will always be "preset" or ""
Public SelectedDataset As String     ' dataset name when UserChoice="preset"

' -----------------------
' Initialize the picker UI
' -----------------------
Public Sub InitPicker()
    Dim ws As Worksheet, lo As ListObject
    Dim dict As Object
    Dim r As Range, v As Variant

    ' Reset outputs
    UserChoice = ""
    SelectedDataset = ""

    ' Clear and (re)configure listbox
    On Error Resume Next
    lstDatasets.Clear
    lstDatasets.ColumnCount = 1
    lstDatasets.BoundColumn = 1
    lstDatasets.ColumnHeads = False
    On Error GoTo 0

    ' 1) Ensure the presets are created/seeded
    On Error Resume Next
    EnsureSupplierPresetsSeeded
    On Error GoTo 0

    ' 2) Find the (hidden) presets sheet & table
    Set ws = SheetByName("SupplierPresets")
    If ws Is Nothing Then
        MsgBox "Preset sheet 'SupplierPresets' not found.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set lo = ws.ListObjects("tblPresets")
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "Preset table 'tblPresets' not found.", vbExclamation
        Exit Sub
    End If

    ' 3) If the table has no body yet, nothing to list
    If lo.DataBodyRange Is Nothing Then
        MsgBox "No preset rows exist yet. Seeded table is empty.", vbInformation
        Exit Sub
    End If

    ' 4) Build a unique list of dataset names from the "Dataset" column
    '    Prefer header match; fallback to column 1 if needed.
    Dim colDS As Long: colDS = 0
    Dim i As Long
    For i = 1 To lo.ListColumns.Count
        If StrComp(Trim$(lo.ListColumns(i).Name), "Dataset", vbTextCompare) = 0 Then
            colDS = i: Exit For
        End If
    Next i
    If colDS = 0 Then colDS = 1

    Set dict = CreateObject("Scripting.Dictionary")
    For Each r In lo.ListColumns(colDS).DataBodyRange.Cells
        v = Trim$(CStr(r.value))
        If Len(v) > 0 Then
            If Not dict.Exists(v) Then dict.add v, True
        End If
    Next r

    ' 5) Populate ListBox
    If dict.Count > 0 Then
        lstDatasets.List = dict.Keys
        ' Select first item by default for a quicker OK
        On Error Resume Next
        lstDatasets.ListIndex = 0
        On Error GoTo 0
    Else
        MsgBox "No dataset names found in presets.", vbInformation
    End If
End Sub

' -----------------------
' OK / Cancel
' -----------------------
Private Sub cmdOK_Click()
    If lstDatasets.ListIndex < 0 Then
        MsgBox "Pick a preset dataset.", vbInformation
        Exit Sub
    End If
    UserChoice = "preset"
    SelectedDataset = CStr(lstDatasets.List(lstDatasets.ListIndex))
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    UserChoice = ""
    SelectedDataset = ""
    Me.Hide
End Sub

' Double-click to accept quickly
Private Sub lstDatasets_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lstDatasets.ListIndex >= 0 Then
        cmdOK_Click
    End If
End Sub

' Enter = OK, Esc = Cancel (nice usability touch)
Private Sub lstDatasets_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdOK_Click
    ElseIf KeyCode = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

' Hide on [X] so caller can still read properties
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        cmdCancel_Click
    End If
End Sub

