' VBA Macro for Filtering, Applying XLOOKUP, and Updating Data
' This is a sample code for automating data processing in Excel VBA

Sub Filter_Blanks_And_Lookup()
    Dim ws As Worksheet
    Dim visibleRow As Range
    Dim firstVisibleRow As Range
    Dim formulaStr As String
    Dim lastRow As Long
    Dim cell As Range
    Dim lookupCell As Range
    Dim filterLastRow As Long
    
    ' Disable automatic calculation and screen updating to improve speed
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("SampleData")
    
    ' Ensure AutoFilter is enabled
    If ws.AutoFilterMode = False Then ws.Range("A1").AutoFilter
    
    ' Filter column CR (column 96) to show only blanks
    ws.Range("A1").AutoFilter Field:=96, Criteria1:="="

    ' Clear contents in column N for visible rows
    On Error Resume Next
    Set visibleRow = ws.Range("N2:N30000").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not visibleRow Is Nothing Then
        For Each cell In visibleRow
            cell.ClearContents
        Next cell
        
        ' Set XLOOKUP formula for column N
        formulaStr = "=XLOOKUP(@SampleData!B2:B30000,ReferenceSheet!A2:A30000,ReferenceSheet!D2:D30000,,,1)"
        
        ' Find the last row in column B
        lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
        
        ' Apply formula to the first visible row
        Set firstVisibleRow = visibleRow.Cells(1, 1)
        firstVisibleRow.Formula = formulaStr
        
        ' Fill formula down to last row
        For i = firstVisibleRow.Row + 1 To lastRow
            If ws.Rows(i).Hidden = False Then
                ws.Cells(i, "N").Formula = formulaStr
            End If
        Next i
    End If
    
    ' Repeat process for column BY
    On Error Resume Next
    Set visibleRow = ws.Range("BY2:BY30000").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not visibleRow Is Nothing Then
        formulaStr = "=XLOOKUP(@SampleData!B2:B30000,ReferenceSheet!H2:H30000,ReferenceSheet!I2:I30000,,,1)"
        
        ' Apply formula in column BY
        Set firstVisibleRow = visibleRow.Cells(1, 1)
        firstVisibleRow.Formula = formulaStr
        
        ' Fill formula down
        For i = firstVisibleRow.Row + 1 To lastRow
            If ws.Rows(i).Hidden = False Then
                ws.Cells(i, "BY").Formula = formulaStr
            End If
        Next i
    End If

    ' Convert formulas to values & replace #N/A with blanks
    For Each lookupCell In ws.Range("N2:N" & lastRow)
        If lookupCell.HasFormula Then lookupCell.Value = lookupCell.Value
        If IsError(lookupCell.Value) Then lookupCell.Value = ""
    Next lookupCell

    For Each lookupCell In ws.Range("BY2:BY" & lastRow)
        If lookupCell.HasFormula Then lookupCell.Value = lookupCell.Value
        If IsError(lookupCell.Value) Then lookupCell.Value = ""
    Next lookupCell

    ' Filter column N for values 1 or blank
    ws.Range("A1").AutoFilter Field:=14, Criteria1:="1", Operator:=xlOr, Criteria2:="="
    
    ' Find last visible row
    filterLastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Update columns P and S
    On Error Resume Next
    Set visibleRow = ws.Range("P2:P" & filterLastRow).SpecialCells(xlCellTypeVisible)
    If Not visibleRow Is Nothing Then visibleRow.Value = 1

    Set visibleRow = ws.Range("S2:S" & filterLastRow).SpecialCells(xlCellTypeVisible)
    If Not visibleRow Is Nothing Then visibleRow.Value = "Yes"
    
    ' Clear filter and apply new filter conditions
    ws.AutoFilter.ShowAllData
    ws.Range("A1").AutoFilter Field:=96, Criteria1:="="
    ws.Range("A1").AutoFilter Field:=16, Criteria1:="1"
    
    ' Re-enable automatic calculation
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
