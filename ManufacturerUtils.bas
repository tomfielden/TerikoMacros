Attribute VB_Name = "ManufacturerUtils"
Option Explicit
Option Base 1

Function SheetExists(sheetName As String) As Boolean
    Dim sheet As Worksheet

    For Each sheet In ActiveWorkbook.Worksheets
        If sheet.Name = sheetName Then SheetExists = True
    Next sheet
End Function


Private Function TableExists(tableName As String) As Boolean
    ' Loop over all known tables and match if found
    Dim table As ListObject

    For Each table In ActiveSheet.ListObjects
        If table.Name = tableName Then TableExists = True
    Next table
End Function


Private Function ColumnExists(tableName As String, columnName As String) As Boolean
    Dim table As ListObject
    Dim column As ListColumn

    ColumnExists = False
    If Not TableExists(tableName) Then
        Exit Function
    End If
    Set table = ActiveSheet.ListObjects(tableName)
    For Each column In table.ListColumns
        If column.Name = columnName Then ColumnExists = True
    Next column
End Function


Private Sub CreateDataTable(tableName As String)
    
    ' Leave the subroutine early if the DataTable already exists so we can call this many times
    If TableExists(tableName) Then
        Exit Sub
    End If
    
    ' Select the entire dataset
    ActiveSheet.Range("A1").CurrentRegion.Select

    ' Create a table object with first row as headers
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = tableName

End Sub


Private Sub InsertColumnBefore(tableName As String, columnName As String, beforeName As String)
    ' 
    Dim table As ListObject
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


Private Sub FormatHeader(tableName As String, columnName As String, styleName As String)
    Dim column As ListColumn

    Set column = ActiveSheet.ListObjects(tableName).ListColumns(columnName)
    ActiveSheet.ListObjects(tableName).HeaderRowRange(column.index).Select
    If styleName = "Good" Then
        Selection.style = "Good"
    End If
End Sub


Private Sub FormatColumn(tableName As String, columnName As String, Optional styleName As String = "Default")
    Dim column As ListColumn

    Set column = ActiveSheet.ListObjects(tableName).ListColumns(columnName)
    If styleName = "Center" Then
        Call FormatColumnCenter(column)
    End If
End Sub


Private Sub FormatColumnCenter(column As Variant)
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


Sub Man_01()
    '
    ' Mfg_Analysis_Start_1
    '
    CreateDataTable "DataTable"
    
    InsertColumnBefore "DataTable", "Item Description", "PRODUCT_DESCRIPTION"
    InsertColumnBefore "DataTable", "Item Pack", "Pack Size"
    
    InsertColumnBefore "DataTable", "School Year", "Date"
    InsertColumnBefore "DataTable", "School Year 1H", "Date"
    InsertColumnBefore "DataTable", "Year", "Date"

End Sub


