Sub Filter_CR_Blanks_And_Lookup_Fill_S_Yes()
    Dim ws As Worksheet
    Dim visibleRow As Range
    Dim cell As Range
    Dim yesterday As Date
    Dim lastRow As Long
    Dim foundYesterday As Boolean
    Dim formulaStr As String
    Dim firstVisibleRow As Range
    Dim filterLastRow As Long
    
    ' Disable automatic calculation and screen updating to improve performance
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' Set the Data worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    ' Enable AutoFilter if not already enabled
    If ws.AutoFilterMode = False Then ws.Range("A1").AutoFilter
    
    ' Filter column CR (column 96) to show only Blanks
    ws.Range("A1").AutoFilter Field:=96, Criteria1:="="
    
    ' Filter column S (column 19) to show only 'Yes'
    ws.Range("A1").AutoFilter Field:=19, Criteria1:="Yes"
    
    ' Get yesterday's date
    yesterday = Date - 1
    
    ' Find the last row in column B
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ' Fill column CS (column 97) with yesterday's date for visible rows
    On Error Resume Next
    Set visibleRow = ws.Range("CS2:CS" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not visibleRow Is Nothing Then
        For Each cell In visibleRow
            cell.Value = yesterday
            cell.NumberFormat = "d/m/yyyy"
        Next cell
    End If
    
    ' Fill column CR (column 96) with "Stop" for visible rows
    On Error Resume Next
    Set visibleRow = ws.Range("CR2:CR" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not visibleRow Is Nothing Then
        For Each cell In visibleRow
            cell.Value = "Stop"
        Next cell
    End If
    
    ' Check if any visible row in column CS contains yesterday's date
    On Error Resume Next
    Set visibleRow = ws.Range("CS2:CS" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    foundYesterday = False
    If Not visibleRow Is Nothing Then
        For Each cell In visibleRow
            If cell.Value = yesterday Then
                foundYesterday = True
                Exit For
            End If
        Next cell
    End If
    
    ' If yesterday's date is found, filter column CS to show only yesterday
    If foundYesterday Then
        ws.Range("A1").AutoFilter Field:=97, Criteria1:=Format(yesterday, "d/m/yyyy")
        
        ' Fill column CO (column 93) with "Trf" for visible rows
        On Error Resume Next
        Set visibleRow = ws.Range("CO2:CO" & lastRow).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        
        If Not visibleRow Is Nothing Then
            For Each cell In visibleRow
                cell.Value = "Trf"
            Next cell
        End If
    Else
        ws.AutoFilterMode = False ' If yesterday's date is not found, clear all filters
    End If
    
    ' Clear all filters
    ws.AutoFilterMode = False
    
    ' ---------------------------------- Set formulas for XLOOKUP ----------------------------------
    
    ' Set the formula for column N
    formulaStr = "=XLOOKUP(@Data!B2:B30000,Report!A2:A30000,Report!D2:D30000,,,1)"
    On Error Resume Next
    Set visibleRow = ws.Range("N2:N30000").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not visibleRow Is Nothing Then
        visibleRow.Cells(1, 1).Formula = formulaStr
        For Each cell In visibleRow
            If ws.Rows(cell.Row).Hidden = False Then
                ws.Cells(cell.Row, "N").Formula = formulaStr
            End If
        Next cell
    End If
    
    ' Set the formula for column BY
    formulaStr = "=XLOOKUP(@Data!B2:B30000,Report!H2:H30000,Report!I2:I30000,,,1)"
    On Error Resume Next
    Set visibleRow = ws.Range("BY2:BY30000").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not visibleRow Is Nothing Then
        visibleRow.Cells(1, 1).Formula = formulaStr
        For Each cell In visibleRow
            If ws.Rows(cell.Row).Hidden = False Then
                ws.Cells(cell.Row, "BY").Formula = formulaStr
            End If
        Next cell
    End If
    
    ' Convert formulas in columns N and BY to their values and remove #N/A
    For Each cell In ws.Range("N2:N" & lastRow)
        If cell.HasFormula Then
            cell.Value = cell.Value ' Replace formula with result
        End If
        If IsError(cell.Value) Then
            cell.Value = "" ' Remove #N/A errors
        End If
    Next cell
    
    For Each cell In ws.Range("BY2:BY" & lastRow)
        If cell.HasFormula Then
            cell.Value = cell.Value ' Replace formula with result
        End If
        If IsError(cell.Value) Then
            cell.Value = "" ' Remove #N/A errors
        End If
    Next cell
    
    ' Filter column N for values that are either 1 or Blank
    ws.Range("A1").AutoFilter Field:=14, Criteria1:="1", Operator:=xlOr, Criteria2:="="
    
    ' Set values for columns P and S for the filtered rows
    filterLastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    On Error Resume Next
    Set visibleRow = ws.Range("P2:P" & filterLastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If Not visibleRow Is Nothing Then
        visibleRow.Value = 1 ' Set visible cells in column P to 1
    End If
    
    Set visibleRow = ws.Range("S2:S" & filterLastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If Not visibleRow Is Nothing Then
        visibleRow.Value = "Yes" ' Set visible cells in column S to 'Yes'
    End If
    
    ' Clear filter on column N
    ws.AutoFilter.ShowAllData
    
    ' Filter column CR for blanks and column P for 1
    ws.Range("A1").AutoFilter Field:=96, Criteria1:="="
    ws.Range("A1").AutoFilter Field:=16, Criteria1:="1"
    
    ' Re-enable automatic calculation and screen updating after the operation
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
