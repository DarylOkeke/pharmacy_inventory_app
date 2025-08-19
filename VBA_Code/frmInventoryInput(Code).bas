Option Explicit

' We build one row per drug: static labels for Drug Name & ID + a textbox for count.
' We keep arrays of control references to collect values on Log.

Private mDataset As String
Private mNames() As String
Private mIDs() As String
Private mTxt() As MSForms.TextBox

Public Sub LoadDrugs(ByVal body As Range, ByVal datasetName As String)
    Dim r As Long, n As Long, i As Long
    Dim arr As Variant
    Dim nameCol As Long, idCol As Long

    mDataset = datasetName
    Me.lblTitle.Caption = "Enter Physical Counts"

    If body Is Nothing Then
        MsgBox "No supplier rows.", vbExclamation
        Unload Me
        Exit Sub
    End If

    ' Detect columns by header (works even if order changes)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = body.ListObject
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "Internal error: no table bound.", vbExclamation
        Unload Me
        Exit Sub
    End If

    ' === FIX: call the public helper in modInventoryOps ===
    nameCol = modInventoryOps.HeaderIndexLocal(lo, "Drug Name"): If nameCol = 0 Then nameCol = 1
    idCol = modInventoryOps.HeaderIndexLocal(lo, "Drug ID"):     If idCol = 0 Then idCol = 2

    arr = body.value
    n = UBound(arr, 1)

    ReDim mNames(1 To n)
    ReDim mIDs(1 To n)
    ReDim mTxt(1 To n)

    ' Layout constants
    Dim topY As Single, rowH As Single
    Dim xName As Single, xID As Single, xQty As Single
    Dim wName As Single, wID As Single, wQty As Single

    topY = 10
    rowH = 22
    xName = 8:   wName = 260
    xID = 275:   wID = 80
    xQty = 360:  wQty = 70

    ' Header row
    Dim lbl As MSForms.Label
    Set lbl = Me.fraList.Controls.add("Forms.Label.1")
    lbl.Caption = "Drug Name"
    lbl.Left = xName: lbl.Top = topY: lbl.width = wName: lbl.Font.Bold = True

    Set lbl = Me.fraList.Controls.add("Forms.Label.1")
    lbl.Caption = "Drug ID"
    lbl.Left = xID: lbl.Top = topY: lbl.width = wID: lbl.Font.Bold = True

    Set lbl = Me.fraList.Controls.add("Forms.Label.1")
    lbl.Caption = "Count"
    lbl.Left = xQty: lbl.Top = topY: lbl.width = wQty: lbl.Font.Bold = True

    topY = topY + rowH + 2

    ' Rows
    For i = 1 To n
        mNames(i) = CStr(arr(i, nameCol))
        mIDs(i) = CStr(arr(i, idCol))

        ' name label
        Set lbl = Me.fraList.Controls.add("Forms.Label.1")
        lbl.Caption = mNames(i)
        lbl.Left = xName: lbl.Top = topY: lbl.width = wName

        ' id label
        Set lbl = Me.fraList.Controls.add("Forms.Label.1")
        lbl.Caption = mIDs(i)
        lbl.Left = xID: lbl.Top = topY: lbl.width = wID

       ' qty textbox (shifted left, narrower, left-aligned)
Dim tb As MSForms.TextBox
Set tb = Me.fraList.Controls.add("Forms.TextBox.1")
tb.Left = xQty - 20          ' was xQty; move it ~20px left
tb.Top = topY - 2
tb.width = wQty - 10         ' was wQty; make a bit narrower
tb.height = rowH
tb.TextAlign = fmTextAlignLeft
tb.Tag = i
Set mTxt(i) = tb

topY = topY + rowH
Next i

Me.fraList.ScrollHeight = topY + 10
End Sub


Private Sub cmdLog_Click()
    Dim n As Long, i As Long
    n = UBound(mNames)

    Dim names() As String, ids() As String, qty() As Variant
    ReDim names(1 To n)
    ReDim ids(1 To n)
    ReDim qty(1 To n)

    For i = 1 To n
        names(i) = mNames(i)
        ids(i) = mIDs(i)
        If Len(Trim$(mTxt(i).Text)) = 0 Then
            qty(i) = 0
        ElseIf IsNumeric(mTxt(i).Text) Then
            qty(i) = CLng(mTxt(i).Text)
        Else
            MsgBox "Row " & i & " has a non-numeric count.", vbExclamation
            Exit Sub
        End If
    Next i

    modInventoryOps.SaveLoggedInventory mDataset, names, ids, qty
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' Optional: size/layout
    Me.width = 470
    Me.height = 520
    Me.fraList.Left = 10
    Me.fraList.Top = 35
    Me.fraList.width = 440
    Me.fraList.height = 420

    Me.cmdLog.Caption = "Log"
    Me.cmdLog.Left = 290
    Me.cmdLog.Top = 465
    Me.cmdLog.width = 75

    Me.cmdCancel.Caption = "Cancel"
    Me.cmdCancel.Left = 370
    Me.cmdCancel.Top = 465
    Me.cmdCancel.width = 75

    Me.lblTitle.Left = 10
    Me.lblTitle.Top = 10
    Me.lblTitle.width = 440
    Me.lblTitle.Caption = "Enter Physical Counts"
End Sub

