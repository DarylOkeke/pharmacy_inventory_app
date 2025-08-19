Option Explicit
' ===========================
'   modSampleAnalytics (clean)
' ===========================
' Public entry points:
' - GenerateSampleData30    -> builds "SampleData" with Inventory + Dispense tables (+ Expiry/LeadTime/SafetyStock)
' - BuildSampleAnalytics    -> builds "Sample Analytics" with KPIs, ops columns (ABC/ROP/ORDER/D2E) + 3 charts
' - AddSampleLauncherButton -> optional launcher button on HOME/active
' - GoHOMESample            -> nav helper

' ---------- Public entry points ----------

Public Sub GenerateSampleData30()
    Dim ws As Worksheet
    Dim drugs As Variant, n As Long, i As Long, d As Long
    Dim idBase As Long, hdrRow As Long, r As Long, lastRow As Long
    Dim startDate As Date, dayDate As Date
    Dim qty As Long, unitCost As Double
    Dim curStock As Long, expStock As Long

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo CleanFail

    ' Remove any prior sample artifacts
    PurgeSampleArtifacts

    ' Fresh sheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "SampleData"

    ' Clean look
    On Error Resume Next
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

    ' ---------- INVENTORY (A:G) ----------
    ws.Range("A1:G1").Value = Array("Drug Name", "Drug ID", "Current Stock", "Expected Stock", "Expiry Date", "LeadTimeDays", "SafetyStock")
    ws.Range("A1:G1").Font.Bold = True

    drugs = Array( _
        "Azovent Tab 10mg", "Carboprill XR 20mg", "Dexomab Inj.", "Lumizole Cap 500mg", "Zylozid Oral Sol.", _
        "MetraGel 5%", "Vitanex ODT 5mg", "Nexvast Tab 40mg", "Heliocillin Susp.", "Orbivast SR 10mg", _
        "Ventacort Tab 5mg", "Cortanex Cap 250mg", "Prilizole Tab 20mg", "Zidomab Inj.", "Dopazole Elixir")
    n = UBound(drugs) - LBound(drugs) + 1
    idBase = 10000

    Randomize
    For i = 0 To n - 1
        curStock = 50 + Int(Rnd * 150)   ' 50..199
        expStock = 40 + Int(Rnd * 120)   ' 40..159
        ws.Cells(i + 2, 1).Value = drugs(i)
        ws.Cells(i + 2, 2).Value = idBase + i
        ws.Cells(i + 2, 3).Value = curStock
        ws.Cells(i + 2, 4).Value = expStock
        ws.Cells(i + 2, 5).Value = Date + 7 + Int(Rnd * 180)   ' Expiry 1–6 months out
        ws.Cells(i + 2, 6).Value = 7 + Int(Rnd * 15)           ' Lead time 7–21 days
        ws.Cells(i + 2, 7).Value = 5 + Int(Rnd * 11)           ' Safety stock 5–15
    Next i
    ws.Columns("A:G").AutoFit
    ws.Range("E2:E" & n + 1).NumberFormat = "m/d/yyyy"

    ' Inventory table A1:G(n+1)
    Dim invRng As Range, loInv As ListObject
    Set invRng = ws.Range(ws.Cells(1, 1), ws.Cells(n + 1, 7))
    Set loInv = ws.ListObjects.Add(xlSrcRange, invRng, , xlYes)
    loInv.Name = "SampleDataTbl_Inventory"
    loInv.DisplayName = "SampleDataTbl_Inventory"
    ApplyGreenTableStyle loInv

    ' ---------- DISPENSE LOGS (A:F) ----------
    hdrRow = n + 3
    ws.Cells(hdrRow, 1).Resize(1, 6).Value = Array("Date", "Drug Name", "Drug ID", "QtyDispensed", "UnitCost", "TotalCost")
    ws.Cells(hdrRow, 1).Resize(1, 6).Font.Bold = True

    startDate = Date - 29
    r = hdrRow + 1

    For i = 0 To n - 1
        For d = 0 To 29
            dayDate = startDate + d
            qty = CLng(MaxD(0#, Round(NormalRand(5, 2), 0)))
            unitCost = Round(5 + Rnd * 95, 2)
            ws.Cells(r, 1).Value = dayDate
            ws.Cells(r, 2).Value = drugs(i)
            ws.Cells(r, 3).Value = idBase + i
            ws.Cells(r, 4).Value = qty
            ws.Cells(r, 5).Value = unitCost
            ws.Cells(r, 6).Value = qty * unitCost
            r = r + 1
        Next d
    Next i

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ws.Range(ws.Cells(hdrRow + 1, 1), ws.Cells(lastRow, 1)).NumberFormat = "m/d/yyyy"
    ws.Range(ws.Cells(hdrRow + 1, 5), ws.Cells(lastRow, 6)).NumberFormat = "$#,##0.00"
    ws.Columns("A:F").AutoFit

    ' Logs table A:F
    Dim logRng As Range, loLog As ListObject
    Set logRng = ws.Range(ws.Cells(hdrRow, 1), ws.Cells(lastRow, 6))
    Set loLog = ws.ListObjects.Add(xlSrcRange, logRng, , xlYes)
    loLog.Name = "SampleDataTbl_Dispense"
    loLog.DisplayName = "SampleDataTbl_Dispense"
    ApplyGreenTableStyle loLog

    ' Build analytics launcher
    On Error Resume Next
    AddButton ws, "btnBuildSampleAnalytics", "Build Sample Analytics", ws.Range("H1"), "BuildSampleAnalytics", 180, 28
    On Error GoTo 0

    ws.Activate
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Sample data generated (30 days).", vbInformation
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "GenerateSampleData30 failed: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub


Public Sub BuildSampleAnalytics()
    Dim wsData As Worksheet, wsA As Worksheet
    Dim loInv As ListObject, loLog As ListObject
    Dim arrInv As Variant, arrLog As Variant
    Dim dictQty As Object, dictSpend As Object, dictAvgCost As Object
    Dim dictLead As Object, dictSS As Object, dictExp As Object
    Dim i As Long, key As Variant

    Set wsData = SheetByName("SampleData")
    If wsData Is Nothing Then
        MsgBox "SampleData not found. Click 'Generate Sample Data (30d)' first.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo CleanFail

    ' Get tables by name
    On Error Resume Next
    Set loInv = wsData.ListObjects("SampleDataTbl_Inventory")
    Set loLog = wsData.ListObjects("SampleDataTbl_Dispense")
    On Error GoTo 0

    If loInv Is Nothing Or loLog Is Nothing Then
        MsgBox "Required tables not found (SampleDataTbl_Inventory / SampleDataTbl_Dispense). Re-run GenerateSampleData30.", vbExclamation
        GoTo CleanFail
    End If
    If loInv.DataBodyRange Is Nothing Or loLog.DataBodyRange Is Nothing Then
        MsgBox "Tables appear to be empty.", vbExclamation
        GoTo CleanFail
    End If

    arrInv = loInv.DataBodyRange.Value   ' 1=Name, 2=ID, 3=Current, 4=Expected, 5=Expiry, 6=Lead, 7=SS
    arrLog = loLog.DataBodyRange.Value   ' 1=Date, 2=Name, 3=ID, 4=Qty, 5=Unit, 6=Total

    Set dictQty = CreateObject("Scripting.Dictionary")
    Set dictSpend = CreateObject("Scripting.Dictionary")
    Set dictAvgCost = CreateObject("Scripting.Dictionary")

    ' Optional indices for extra inventory fields (safe if missing)
    Dim idxExp As Long, idxLead As Long, idxSS As Long
    On Error Resume Next
    idxExp = loInv.ListColumns("Expiry Date").Index
    idxLead = loInv.ListColumns("LeadTimeDays").Index
    idxSS = loInv.ListColumns("SafetyStock").Index
    On Error GoTo 0

    Set dictLead = CreateObject("Scripting.Dictionary")
    Set dictSS = CreateObject("Scripting.Dictionary")
    Set dictExp = CreateObject("Scripting.Dictionary")

    ' Map inventory extras by Drug Name (lowercased)
    For i = 1 To UBound(arrInv, 1)
        Dim nmInv As String
        nmInv = LCase$(Trim$(CStr(arrInv(i, 1))))
        If Len(nmInv) > 0 Then
            If idxLead > 0 Then dictLead(nmInv) = NzD(arrInv(i, idxLead)) Else dictLead(nmInv) = 7
            If idxSS > 0 Then dictSS(nmInv) = NzD(arrInv(i, idxSS)) Else dictSS(nmInv) = 0
            If idxExp > 0 Then dictExp(nmInv) = arrInv(i, idxExp) Else dictExp(nmInv) = Empty
        End If
    Next i

    ' Aggregate logs by Drug Name
    Dim nm As String, q As Double, uc As Double, spend As Double
    For i = 1 To UBound(arrLog, 1)
        If IsDate(arrLog(i, 1)) Then
            nm = LCase$(Trim$(CStr(arrLog(i, 2))))
            If Len(nm) > 0 Then
                q = NzD(arrLog(i, 4))
                uc = NzD(arrLog(i, 5))
                spend = q * uc
                If dictQty.Exists(nm) Then
                    dictQty(nm) = dictQty(nm) + q
                    dictSpend(nm) = dictSpend(nm) + spend
                Else
                    dictQty(nm) = q
                    dictSpend(nm) = spend
                End If
            End If
        End If
    Next i
    For Each key In dictQty.Keys
        If dictQty(key) > 0 Then
            dictAvgCost(key) = dictSpend(key) / dictQty(key)
        Else
            dictAvgCost(key) = vbNullString
        End If
    Next key

    ' ----- Build Analytics sheet -----
    DeleteSheetIfExists_Sample "Sample Analytics"
    Set wsA = ThisWorkbook.Worksheets.Add(After:=wsData)
    wsA.Name = "Sample Analytics"

    On Error Resume Next
    wsA.Activate
    ActiveWindow.DisplayGridlines = False
    On Error GoTo 0

    With wsA.Range("A1:R2")
        .Merge
        .Value = "Inventory Analytics (Sample) — last 30 days"
        .Font.Bold = True
        .Font.Size = 18
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(226, 239, 218)
    End With

    ' 11 columns (no Risk)
    wsA.Range("A4:K4").Value = Array( _
        "Drug Name", "Drug ID", "Current Stock", "Expected Stock", "Dispensed (30d)", _
        "Avg Daily Dispense", "Avg Unit Cost", "Spend (30d)", "Turnover (30d)", _
        "Days Until Stockout", "Note")
    wsA.Range("A4:K4").Font.Bold = True
    wsA.Range("A4:K4").Interior.Color = RGB(198, 239, 206)

    Dim rowOut As Long: rowOut = 5
    Dim pretty As String, drugID As Variant
    Dim phys As Double, expct As Double
    Dim usedQty As Double, avgDaily As Double
    Dim avgCost As Variant, totSpend As Double
    Dim avgStock As Double, turnover As Variant, daysLeft As Variant

    For i = 1 To UBound(arrInv, 1)
        pretty = CStr(arrInv(i, 1))
        drugID = arrInv(i, 2)
        phys = NzD(arrInv(i, 3))
        expct = NzD(arrInv(i, 4))

        usedQty = 0: totSpend = 0: avgCost = vbNullString
        nm = LCase$(Trim$(pretty))
        If dictQty.Exists(nm) Then usedQty = CDbl(dictQty(nm))
        If dictSpend.Exists(nm) Then totSpend = CDbl(dictSpend(nm))
        If dictAvgCost.Exists(nm) Then avgCost = dictAvgCost(nm)

        avgDaily = usedQty / 30#
        If expct > 0 Then avgStock = (phys + expct) / 2# Else avgStock = phys
        If avgStock > 0 Then turnover = usedQty / avgStock Else turnover = vbNullString
        If avgDaily > 0 Then daysLeft = phys / avgDaily Else daysLeft = vbNullString

        With wsA
            .Cells(rowOut, 1).Value = pretty
            .Cells(rowOut, 2).Value = drugID
            .Cells(rowOut, 3).Value = phys
            .Cells(rowOut, 4).Value = expct
            .Cells(rowOut, 5).Value = usedQty
            .Cells(rowOut, 6).Value = avgDaily
            .Cells(rowOut, 7).Value = avgCost
            .Cells(rowOut, 8).Value = totSpend
            .Cells(rowOut, 9).Value = IIf(avgStock > 0, turnover, vbNullString)
            .Cells(rowOut, 10).Value = daysLeft
            .Cells(rowOut, 11).Value = vbNullString   ' Note
        End With
        rowOut = rowOut + 1
    Next i

    Dim lastDataRow As Long: lastDataRow = rowOut - 1

    If lastDataRow >= 5 Then
        Dim tbl As ListObject
        Set tbl = wsA.ListObjects.Add(xlSrcRange, wsA.Range("A4:K" & lastDataRow), , xlYes)
        On Error Resume Next: tbl.Name = "SampleAnalyticsTable": On Error GoTo 0
        ApplyGreenTableStyle tbl

        ' Number formats
        wsA.Range("C5:D" & lastDataRow).NumberFormat = "0"
        wsA.Range("E5:F" & lastDataRow).NumberFormat = "0.00"
        wsA.Range("G5").Resize(lastDataRow - 4, 1).NumberFormat = "$#,##0.00"
        wsA.Range("H5").Resize(lastDataRow - 4, 1).NumberFormat = "$#,##0.00"
        wsA.Range("I5").Resize(lastDataRow - 4, 1).NumberFormat = "0.00"
        wsA.Range("J5").Resize(lastDataRow - 4, 1).NumberFormat = "0.0"

        ' Extra ops columns
        Dim colABC As ListColumn, colROP As ListColumn, colOrder As ListColumn, colExpDays As ListColumn
        Set colABC = tbl.ListColumns.Add: colABC.Name = "ABC Class"
        Set colROP = tbl.ListColumns.Add: colROP.Name = "Reorder Point"
        Set colOrder = tbl.ListColumns.Add: colOrder.Name = "Order Now"
        Set colExpDays = tbl.ListColumns.Add: colExpDays.Name = "Days to Expiry"

        Dim totalRows As Long: totalRows = tbl.DataBodyRange.Rows.Count
        Dim rIdx As Long, nmRow As String
        Dim lead As Double, ss As Double, rop As Double, orderNow As String
        Dim expDate As Variant, daysToExp As Variant

        For rIdx = 1 To totalRows
            nmRow = LCase$(Trim$(tbl.DataBodyRange.Cells(rIdx, 1).Value)) ' Drug Name
            lead = IIf(dictLead.Exists(nmRow), CDbl(dictLead(nmRow)), 7#)
            ss = IIf(dictSS.Exists(nmRow), CDbl(dictSS(nmRow)), 0#)
            rop = NzD(tbl.DataBodyRange.Cells(rIdx, 6).Value) * lead + ss ' AvgDaily * Lead + SS
            tbl.ListColumns("Reorder Point").DataBodyRange.Cells(rIdx, 1).Value = rop

            If NzD(tbl.DataBodyRange.Cells(rIdx, 3).Value) <= rop Then
                orderNow = "ORDER NOW"
            Else
                orderNow = ""
            End If
            tbl.ListColumns("Order Now").DataBodyRange.Cells(rIdx, 1).Value = orderNow

            If dictExp.Exists(nmRow) And IsDate(dictExp(nmRow)) Then
                expDate = CDate(dictExp(nmRow))
                daysToExp = DateDiff("d", Date, expDate)
            Else
                daysToExp = vbNullString
            End If
            tbl.ListColumns("Days to Expiry").DataBodyRange.Cells(rIdx, 1).Value = daysToExp
        Next rIdx

        ' ABC Class by Spend(30d) descending
        Dim spendRange As Range, rankVal As Long, nCount As Long, cutA As Long, cutB As Long
        Set spendRange = tbl.ListColumns("Spend (30d)").DataBodyRange
        nCount = spendRange.Rows.Count
        cutA = Application.WorksheetFunction.RoundUp(0.2 * nCount, 0) ' top 20%
        cutB = Application.WorksheetFunction.RoundUp(0.5 * nCount, 0) ' next 30%

        For rIdx = 1 To totalRows
            rankVal = SafeRank(NzD(tbl.DataBodyRange.Cells(rIdx, 8).Value), spendRange) ' col H
            Select Case True
                Case rankVal <= cutA: tbl.ListColumns("ABC Class").DataBodyRange.Cells(rIdx, 1).Value = "A"
                Case rankVal <= cutB: tbl.ListColumns("ABC Class").DataBodyRange.Cells(rIdx, 1).Value = "B"
                Case Else:            tbl.ListColumns("ABC Class").DataBodyRange.Cells(rIdx, 1).Value = "C"
            End Select
        Next rIdx

        ' Formats + conditional formats
        tbl.ListColumns("Reorder Point").DataBodyRange.NumberFormat = "0.0"
        tbl.ListColumns("Days to Expiry").DataBodyRange.NumberFormat = "0"

        With tbl.ListColumns("Order Now").DataBodyRange.FormatConditions.Add(Type:=xlTextString, String:="ORDER NOW", TextOperator:=xlContains)
            .Interior.Color = RGB(255, 235, 156)
        End With
        With tbl.ListColumns("Days to Expiry").DataBodyRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=30")
            .Interior.Color = RGB(255, 242, 204)
        End With
        With tbl.ListColumns("ABC Class").DataBodyRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""A""")
            .Interior.Color = RGB(198, 239, 206)
        End With

        ' AutoFit after adding extras
        tbl.Range.Columns.AutoFit
    End If

    ' ---- Charts ----
    Dim anchorRow As Long: anchorRow = lastDataRow + 2
    Dim ch1 As ChartObject, ch2 As ChartObject, ch3 As ChartObject

    Set ch1 = MakeColumnChartExplicit(wsA, _
        xVals:=wsA.Range(wsA.Cells(5, 1), wsA.Cells(lastDataRow, 1)), _
        yVals:=wsA.Range(wsA.Cells(5, 8), wsA.Cells(lastDataRow, 8)), _
        title:="Spend (30d) by Drug", _
        leftTop:=wsA.Range("A" & anchorRow), width:=340, height:=220)
    StyleChartGreen ch1

    Set ch2 = MakeColumnChartExplicit(wsA, _
        xVals:=wsA.Range(wsA.Cells(5, 1), wsA.Cells(lastDataRow, 1)), _
        yVals:=wsA.Range(wsA.Cells(5, 10), wsA.Cells(lastDataRow, 10)), _
        title:="Days Until Stockout", _
        leftTop:=wsA.Range("G" & anchorRow), width:=340, height:=220)
    StyleChartGreen ch2

    Set ch3 = MakeColumnChartExplicit(wsA, _
        xVals:=wsA.Range(wsA.Cells(5, 1), wsA.Cells(lastDataRow, 1)), _
        yVals:=wsA.Range(wsA.Cells(5, 9), wsA.Cells(lastDataRow, 9)), _
        title:="Turnover (30d)", _
        leftTop:=wsA.Range("M" & anchorRow), width:=340, height:=220)
    StyleChartGreen ch3

    AddButton wsA, "btnHOMESampleAnalytics", "HOME", wsA.Range("L1"), "GoHOMESample", 90, 26)

    wsA.Activate
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "BuildSampleAnalytics failed: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub


Public Sub AddSampleLauncherButton()
    Dim ws As Worksheet
    Set ws = SheetByName("HOME")
    If ws Is Nothing Then
        Set ws = ActiveSheet
        If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(1)
    End If
    AddButton ws, "btnGenSampleData", "Generate Sample Data (30d)", ws.Range("H1"), "GenerateSampleData30", 200, 28
    MsgBox "Launcher button added on sheet: " & ws.Name, vbInformation
End Sub

Public Sub GoHOMESample()
    Dim ws As Worksheet
    Set ws = SheetByName("HOME")
    If ws Is Nothing Then
        Set ws = SheetByName("SampleData")
        If ws Is Nothing Then Exit Sub
    End If
    ws.Activate
End Sub

' ---------- Helpers (only those actually used) ----------

Private Sub DeleteSheetIfExists_Sample(ByVal sName As String)
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(sName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Private Function SheetByName(ByVal sName As String) As Worksheet
    On Error Resume Next
    Set SheetByName = ThisWorkbook.Worksheets(sName)
    On Error GoTo 0
End Function

Private Function NormalRand(mu As Double, sigma As Double) As Double
    Const PI As Double = 3.14159265358979
    Dim u1 As Double, u2 As Double
    u1 = Rnd: If u1 = 0 Then u1 = 0.0001
    u2 = Rnd
    NormalRand = mu + sigma * Sqr(-2# * Log(u1)) * Cos(2# * PI * u2)
End Function

Private Function MaxD(a As Double, b As Double) As Double
    If a >= b Then MaxD = a Else MaxD = b
End Function

Private Function NzD(v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then
        NzD = 0#
    ElseIf VarType(v) = vbString Then
        If Len(Trim$(v & "")) = 0 Then NzD = 0# Else NzD = CDbl(Val(v))
    ElseIf IsNumeric(v) Then
        NzD = CDbl(v)
    Else
        NzD = 0#
    End If
End Function

Private Sub AddButton(ws As Worksheet, btnName As String, btnText As String, anchor As Range, onActionMacro As String, _
                      Optional width As Single = 120, Optional height As Single = 26)
    On Error Resume Next
    ws.Shapes(btnName).Delete
    On Error GoTo 0

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(Type:=msoShapeRoundedRectangle, _
                                 Left:=anchor.Left, _
                                 Top:=anchor.Top + 4, _
                                 Width:=width, Height:=height)
    With shp
        .Name = btnName
        .TextFrame2.TextRange.Characters.Text = btnText
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 97, 0)                  ' dark green
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255) ' white
        .OnAction = onActionMacro
        .Shadow.Type = msoShadow6
        .Shadow.ForeColor.RGB = RGB(200, 200, 200)
    End With
End Sub

' Apply a consistent green table look with manual striping
Private Sub ApplyGreenTableStyle(lo As ListObject)
    On Error Resume Next
    lo.TableStyle = ""
    lo.TableStyle = "TableStyleLight9"
    lo.ShowTableStyleRowStripes = False
    lo.ShowTableStyleColumnStripes = False

    If Not lo.HeaderRowRange Is Nothing Then
        lo.HeaderRowRange.Interior.Color = RGB(198, 239, 206)
        lo.HeaderRowRange.Font.Bold = True
    End If

    If Not lo.DataBodyRange Is Nothing Then
        Dim r As Long, lastR As Long
        lastR = lo.DataBodyRange.Rows.Count
        lo.DataBodyRange.Interior.ColorIndex = xlColorIndexNone
        For r = 1 To lastR
            If (r Mod 2) = 1 Then
                lo.DataBodyRange.Rows(r).Interior.Color = RGB(235, 248, 238)
            Else
                lo.DataBodyRange.Rows(r).Interior.Color = RGB(221, 235, 221)
            End If
        Next r
    End If
    On Error GoTo 0
End Sub

' Create a column chart with explicit X/Y ranges
Private Function MakeColumnChartExplicit(ws As Worksheet, xVals As Range, yVals As Range, _
                                         title As String, leftTop As Range, _
                                         Optional width As Single = 420, Optional height As Single = 240) As ChartObject
    Dim ch As ChartObject
    Set ch = ws.ChartObjects.Add(Left:=leftTop.Left, Top:=leftTop.Top, Width:=width, Height:=height)
    With ch.Chart
        .ChartType = xlColumnClustered
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .XValues = xVals
            .Values = yVals
            .Name = title
            .Format.Fill.ForeColor.RGB = RGB(0, 176, 80)
            .Format.Line.Visible = msoFalse
        End With
        On Error Resume Next
        .ChartGroups(1).GapWidth = 120
        .ChartGroups(1).Overlap = 0
        On Error GoTo 0
        .HasTitle = True
        .ChartTitle.Text = title
        .HasLegend = False
        On Error Resume Next
        .Axes(xlCategory).TickLabelSpacing = 1
        On Error GoTo 0
    End With
    Set MakeColumnChartExplicit = ch
End Function

' Light green chart styling
Private Sub StyleChartGreen(ByVal ch As ChartObject)
    On Error Resume Next
    With ch.Chart
        .ChartArea.RoundedCorners = True
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .PlotArea.Format.Line.Visible = msoFalse
        .Axes(xlCategory).TickLabels.Font.Size = 9
        .Axes(xlValue).TickLabels.Font.Size = 9
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(220, 220, 220)
        .Axes(xlValue).MajorGridlines.Format.Line.Transparency = 0.2
    End With
    On Error GoTo 0
End Sub

' Descending rank (RANK.EQ) — 1 = highest
Private Function SafeRank(ByVal v As Double, ByVal rng As Range) As Long
    Dim c As Range, countGreater As Long
    countGreater = 0
    For Each c In rng.Cells
        If IsNumeric(c.Value) Then
            If CDbl(c.Value) > v Then countGreater = countGreater + 1
        End If
    Next c
    SafeRank = countGreater + 1
End Function

' Unhide & delete old Sample sheets; clear any leftover table/names with same display names
Private Sub PurgeSampleArtifacts()
   Dim ws As Worksheet, lo As ListObject

   Application.DisplayAlerts = False
   On Error Resume Next

   ' Delete the specific sheets so their tables/names go away with them
   For Each ws In ThisWorkbook.Worksheets
       If ws.Name = "SampleData" Or ws.Name = "Sample Analytics" Then
           ws.Visible = xlSheetVisible
           ws.Delete
       End If
   Next ws

   ' If any tables with these names survived on other sheets, rename them (extreme edge)
   For Each ws In ThisWorkbook.Worksheets
       For Each lo In ws.ListObjects
           If StrComp(lo.Name, "SampleDataTbl_Inventory", vbTextCompare) = 0 _
              Or StrComp(lo.DisplayName, "SampleDataTbl_Inventory", vbTextCompare) = 0 Then
               lo.Name = "SampleDataTbl_Inventory_old_" & Format(Now, "yyyymmdd_hhnnss")
               lo.DisplayName = lo.Name
           End If
           If StrComp(lo.Name, "SampleDataTbl_Dispense", vbTextCompare) = 0 _
              Or StrComp(lo.DisplayName, "SampleDataTbl_Dispense", vbTextCompare) = 0 Then
               lo.Name = "SampleDataTbl_Dispense_old_" & Format(Now, "yyyymmdd_hhnnss")
               lo.DisplayName = lo.Name
           End If
       Next lo
   Next ws

   On Error GoTo 0
   Application.DisplayAlerts = True
End Sub
