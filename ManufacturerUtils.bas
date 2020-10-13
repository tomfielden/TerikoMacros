Attribute VB_Name = "ManufacturerUtils"
Option Explicit
Option Base 1

Private Function SheetExists(sheetName As String) As Boolean
    Dim sheet As Worksheet

    For Each sheet In ActiveWorkbook.Worksheets
        If sheet.Name = sheetName Then
            SheetExists = True
            Exit Function
        End If
    Next sheet
    SheetExists = False
End Function


Private Function TableExists(tableName As String) As Boolean
    ' Loop over all known tables and match if found
    Dim table As ListObject

    For Each table In ActiveSheet.ListObjects
        If table.Name = tableName Then
            TableExists = True
            Exit Function
        End If
    Next table
    TableExists = False
End Function


Private Function ColumnExists(table As ListObject, columnName As String) As Boolean
    Dim column As ListColumn

    For Each column In table.ListColumns
        If column.Name = columnName Then
            ColumnExists = True
            Exit Function
        End If
    Next column
    ColumnExists = False
End Function


Private Function GetDataTable(tableName As String) As ListObject
    ' Leave the subroutine early if the DataTable already exists so we can call this many times

    If TableExists(tableName) Then
        Exit Function
    End If
    
    ' Select the entire dataset
    ActiveSheet.Range("A1").CurrentRegion.Select

    ' Create a table object with first row as headers
    ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = tableName
    Set GetDataTable = ActiveSheet.ListObjects(tableName)

End Function


Private Sub InsertColumnBefore(table As ListObject, columnName As String, beforeName As String)
    '
    Dim column As ListColumn
    Dim before_column As Long

    If ColumnExists(table, columnName) Then
        Exit Sub
    End If

    With table
        before_column = .ListColumns(beforeName).index
        .ListColumns.Add(before_column).Name = columnName
    End With

    Set column = table.ListColumns(columnName)

    FormatHeader table, column, "Good"
End Sub


Private Sub FormatHeader(table As ListObject, column As ListColumn, styleName As String)
    '
    table.HeaderRowRange(column.index).Select
    If styleName = "Good" Then
        Selection.style = "Good"
    End If
End Sub


Private Sub FormatColumn(column As ListColumn, Optional styleName As String = "Default")
    '
    If styleName = "Center" Then
        Call FormatColumnCenter(column)
    End If
End Sub


Private Sub FormatColumnCenter(column As ListColumn)
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


Sub PrepareTable()
    '
    ' Mfg_Analysis_Start_1
    ' Re-entrant code. We don't need to call this directly.
    '
    Dim table As ListObject

    Set table = GetDataTable("DataTable")
    
    UpdateItemDescription table, "Item Description", "PRODUCT_DESCRIPTION"
    InsertColumnBefore table, "Item Pack", "Pack Size"
    
    InsertColumnBefore table, "School Year", "Date"
    InsertColumnBefore table, "School Year 1H", "Date"
    InsertColumnBefore table, "Year", "Date"

End Sub


Sub UpdateItemDescription(table As ListObject, itemColumnName As String, beforeColumnName As String)
    '
    InsertColumnBefore table, itemColumnName, beforeColumnName


End Sub


