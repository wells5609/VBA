Option Explicit

' Returns the cell's address
' Alias of Range.Address
Function CellAddress( _
    cell As Excel.Range, _
    Optional row_absolute As Long = 1, _
    Optional column_absolute As Long = 1) As String
    
    CellAddress = cell.Address(row_absolute, column_absolute)
End Function

' Returns the cell's address, including the sheet name.
Function CellSheetAddress( _
    cell As Excel.Range, _
    Optional row_absolute As Long = 1, _
    Optional column_absolute As Long = 1) As String
    
    CellSheetAddress = cell.Worksheet.Name & "!" & cell.Address(row_absolute, column_absolute)
End Function

' Returns the cell's address, including the file path and sheet name.
Function CellAbsoluteAddress( _
    cell As Excel.Range, _
    Optional row_absolute As Long = 1, _
    Optional column_absolute As Long = 1) As String
    
    With cell
        CellAbsoluteAddress = "'" & .Worksheet.Parent.Path & "\[" & .Worksheet.Parent.Name & "]" _
            & .Worksheet.Name & "'!" & cell.Address(row_absolute, column_absolute)
    End With
End Function

' Whether a sheet with the given name exists in a workbook.
Function SheetExists( _
    ByVal Name As String, _
    Optional wb As Excel.Workbook = Nothing) As Boolean
    
    Dim Sheet As Excel.Worksheet
    If wb Is Nothing Then Set wb = Excel.ActiveWorkbook
    On Error Resume Next
    Set Sheet = wb.Sheets(Name)
    On Error GoTo 0
    If Sheet Is Nothing Then
        SheetExists = False
    Else
        SheetExists = True
    End If
End Function

' Returns the sheet name of a cell.
Function SheetName(Optional cell As Excel.Range = Nothing) As String
    If cell Is Nothing Then Set cell = Excel.ActiveCell
    SheetName = cell.Worksheet.Name
End Function

' Returns the first cell in a range as an Excel.Range object.
Function FirstCell(rng As Excel.Range) As Excel.Range
    Set FirstCell = rng.Cells(1, 1)
End Function

' Returns the last cell in a range as an Excel.Range object.
Function LastCell(rng As Excel.Range) As Excel.Range
    Set LastCell = rng.Cells(rng.rows.Count, rng.Columns.Count)
End Function

' Returns the first row in a range as an Excel.Range object.
Function FirstRow(rng As Excel.Range) As Excel.Range
    Set FirstRow = rng.rows(1)
End Function

' Returns the last column in a range as an Excel.Range object.
Function LastColumn(rng As Excel.Range) As Excel.Range
    Set LastColumn = rng.Columns(rng.Columns.Count)
End Function
