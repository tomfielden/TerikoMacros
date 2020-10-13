Attribute VB_Name = "Module2"
    ' Example:
    ' ActiveSheet.ListObjects("DataTable").Range.Select
    ' ActiveSheet.ListObjects("DataTable").ListColumns("Date").DataBodyRange.Select
    ' ActiveSheet.ListObjects("DataTable").ListColumns("BILLED_REBATE_AMT").Range.Select
    ' ActiveSheet.ListObjects("DataTable").ListColumns.Add(5).Name = "NewColumn"
    '
    ' Dim rg As Range
    ' Set rg = ActiveSheet.Range("A1").CurrentRegion  ' A better way to select the whole data table!

Function SheetExists(sheetName As String) As Boolean
    For Each Sheet In ActiveWorkbook.Worksheets
        If Sheet.Name = sheetName Then SheetExists = True
    Next Sheet
End Function

Function TableExists(tableName As String) As Boolean
    ' Loop over all known tables and match if found
    For Each table In ActiveSheet.ListObjects
        If table.Name = tableName Then TableExists = True
    Next table
End Function

Function ColumnExists(tableName As String, columnName As String) As Boolean
    ColumnExists = False
    If Not TableExists(tableName) Then
        Exit Function
    End If
    Set table = ActiveSheet.ListObjects(tableName)
    For Each column In table.ListColumns
        If column.Name = columnName Then ColumnExists = True
    Next column
End Function

Sub InsertColumnBefore(tableName As String, columnName As String, beforeName As String)
    Dim before_column As Long
    If ColumnExists(tableName, columnName) Then
        Exit Sub
    End If
    Set table = ActiveSheet.ListObjects(tableName)
    With table
        before_column = .ListColumns(beforeName).index
        .ListColumns.Add(before_column).Name = columnName
    End With
    Call FormatHeader(tableName, columnName, "Good")
End Sub

Sub FormatHeader(tableName As String, columnName As String, styleName As String)
    Set column = ActiveSheet.ListObjects(tableName).ListColumns(columnName)
    ActiveSheet.ListObjects(tableName).HeaderRowRange(column.index).Select
    If styleName = "Good" Then
        Selection.style = "Good"
    End If
End Sub

Sub FormatColumn(tableName As String, columnName As String, Optional styleName As String = "Default")
    Set column = ActiveSheet.ListObjects(tableName).ListColumns(columnName)
    If styleName = "Center" Then
        Call FormatColumnCenter(column)
    End If
End Sub
Sub FormatColumnCenter(column As Variant)
    column.DataBodyRange.Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub


Sub CreateDataTable(tableName As String)
    
    ' Leave the subroutine early if the DataTable already exists so we can call this many times
    If TableExists(tableName) Then
        Exit Sub
    End If
    
    ' Select the entire dataset
    ' Range("A1").Select
    ' Range(Selection, Selection.End(xlDown)).Select
    ' Range(Selection, Selection.End(xlToRight)).Select
    ActiveSheet.Range("A1").CurrentRegion.Select

    ' Create a table object with first row as headers
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = tableName

End Sub


Sub TRF_01()
'
' Mfg_Analysis_Start_1
'
    Call CreateDataTable("DataTable")
    
    Call InsertColumnBefore("DataTable", "Item Description", "PRODUCT_DESCRIPTION")
    Call InsertColumnBefore("DataTable", "Item Pack", "Pack Size")
    
    Call InsertColumnBefore("DataTable", "School Year", "Date")
    Call InsertColumnBefore("DataTable", "School Year 1H", "Date")
    Call InsertColumnBefore("DataTable", "Year", "Date")
    

End Sub

