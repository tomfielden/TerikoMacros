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

    If Not TableExists(tableName) Then
        ' Select the entire dataset
        ActiveSheet.Range("A1").CurrentRegion.Select
    
        ' Create a table object with first row as headers
        ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes).Name = tableName
    End If
    
    Set GetDataTable = ActiveSheet.ListObjects(tableName)

End Function


Private Function InsertColumnBefore(table As ListObject, columnName As String, beforeName As String) As ListColumn
    '
    Dim before_column As Long

    If Not ColumnExists(table, columnName) Then
        before_column = table.ListColumns(beforeName).Index
        table.ListColumns.Add(before_column).Name = columnName
    End If

    Set InsertColumnBefore = table.ListColumns(columnName)
    FormatHeader table, InsertColumnBefore, "Good"
    FormatColumn InsertColumnBefore, "Center"
End Function


Private Sub FormatHeader(table As ListObject, column As ListColumn, styleName As String)
    '
    table.HeaderRowRange(column.Index).Select
    If styleName = "Good" Then
        Selection.style = "Good"
    End If
End Sub


Private Sub FormatColumn(column As ListColumn, Optional styleName As String = "Default")
    '
    If styleName = "Center" Then
        FormatColumnCenter column
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


Private Function GetArray(table As ListObject, ByRef columnNames As Variant) As Variant
    ' Given a table and a list of column names
    ' Return a 2D array of data values from table body

    Dim arr() As Variant
    Dim column As Variant
    Dim num_rows As Long
    Dim num_cols As Long
    Dim col As Long
    Dim row As Long
    Dim columnName As String
    
    num_rows = table.DataBodyRange.Rows.Count
    num_cols = UBound(columnNames, 1) - LBound(columnNames, 1) + 1
    ReDim arr(num_rows, num_cols)
    
    For col = 1 To num_cols
        columnName = columnNames(col)
        column = table.ListColumns(columnName).DataBodyRange
        ' Copy data into array
        For row = 1 To num_rows
            arr(row, col) = column(row, 1)
        Next row
    Next col
    GetArray = arr
End Function


Private Sub AddIndexToArray(ByRef arr As Variant)
    Dim num_rows As Long
    Dim num_cols As Long
    Dim row As Long

    num_rows = UBound(arr, 1)
    num_cols = UBound(arr, 2)

    ReDim Preserve arr(num_rows, num_cols + 1) As Variant
    ' Create index values
    For row = 1 To num_rows
        arr(row, num_cols + 1) = row
    Next row
End Sub


Private Function ArrayToSheet(ByRef arr As Variant, sheetName As String) As Worksheet
    Dim sheet As Worksheet

    If SheetExists(sheetName) Then
        Application.DisplayAlerts = False
        ActiveWorkbook.Worksheets(sheetName).Delete
        Application.DisplayAlerts = True
    End If
    ActiveWorkbook.Worksheets.Add.Name = sheetName
    Set sheet = ActiveWorkbook.Worksheets(sheetName)
    sheet.Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)) = arr
    Set ArrayToSheet = sheet
End Function


Private Sub SortArray(ByRef arr As Variant, ByRef sort_cols As Variant)
    Dim sheetName As String
    Dim sheet As Worksheet
    Dim sort_col As Variant
    Dim ws_index As Long

    ws_index = ActiveWorkbook.ActiveSheet.Index

    sheetName = "__SortArray__"
    Set sheet = ArrayToSheet(arr, sheetName)

    With sheet.Sort
        .SortFields.Clear
        For Each sort_col In sort_cols
            If sort_col < 0 Then
                .SortFields.Add2 _
                    Key:=Columns(-sort_col), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlDescending, _
                    DataOption:=xlSortNormal
            Else
                .SortFields.Add2 _
                    Key:=Columns(sort_col), _
                    SortOn:=xlSortOnValues, _
                    Order:=xlAscending, _
                    DataOption:=xlSortNormal
            End If
        Next sort_col
        .SetRange Range(sheet.Cells(1, 1), sheet.Cells(UBound(arr, 1), UBound(arr, 2)))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    arr = sheet.Range(sheet.Cells(1, 1), sheet.Cells(UBound(arr, 1), UBound(arr, 2))).Value

    ' Cleanup
    ActiveWorkbook.Worksheets(ws_index).Select
    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets(sheetName).Delete
    Application.DisplayAlerts = True

End Sub


Private Sub UpdateItemDescription(table As ListObject, itemColumnName As String, beforeColumnName As String)
    '
    Dim arr() As Variant
    
    InsertColumnBefore table, itemColumnName, beforeColumnName
    arr = GetArray(table, [{"Manufacturer", "#SKU", "PRODUCT_DESCRIPTION", "Cases (NVD)"}])
    AddIndexToArray arr
    SortArray arr, [{1, 2, -4}]
    
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



