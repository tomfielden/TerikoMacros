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


Private Function InsertColumnBefore(table As ListObject, columnName As String, beforeName As String, Optional styleName As String = "Left") As ListColumn
    '
    Dim before_column As Long
    
    If Not ColumnExists(table, beforeName) Then
        MsgBox "Missing required column: " & beforeName
        End
    End If

    If Not ColumnExists(table, columnName) Then
        before_column = table.ListColumns(beforeName).index
        table.ListColumns.Add(before_column).Name = columnName
    End If

    Set InsertColumnBefore = table.ListColumns(columnName)
    FormatHeader table, InsertColumnBefore, "Good"
    FormatColumn InsertColumnBefore, styleName
End Function


Private Sub FormatHeader(table As ListObject, column As ListColumn, styleName As String)
    '
    table.HeaderRowRange(column.index).Select
    If styleName = "Good" Then
        Selection.style = "Good"
    End If
End Sub


Private Sub FormatColumn(column As ListColumn, Optional styleName As String)
    '
    If styleName = "Center" Then
        FormatColumnCenter column
    ElseIf styleName = "Left" Then
        FormatColumnLeft column
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

Private Sub FormatColumnLeft(column As ListColumn)
    column.DataBodyRange.Select
    With Selection
        .HorizontalAlignment = xlLeft
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
        If Not ColumnExists(table, columnName) Then
            MsgBox "Missing required column: " & columnName
            End
        End If
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


Private Sub AddColumnToArray(ByRef arr As Variant)
    Dim num_rows As Long
    Dim num_cols As Long

    num_rows = UBound(arr, 1)
    num_cols = UBound(arr, 2)

    ReDim Preserve arr(num_rows, num_cols + 1) As Variant

End Sub


Private Function GetArrayColumn(ByRef arr As Variant, col_index As Long) As Variant
    Dim num_rows As Long
    Dim num_cols As Long
    Dim row As Long
    Dim column() As Variant

    num_rows = UBound(arr, 1)
    num_cols = UBound(arr, 2)
    ReDim column(num_rows, 1) As Variant
    
    For row = 1 To num_rows
        column(row, 1) = arr(row, col_index)
    Next row

    GetArrayColumn = column
End Function


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

    ws_index = ActiveWorkbook.ActiveSheet.index

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

    arr = sheet.Range(sheet.Cells(1, 1), sheet.Cells(UBound(arr, 1), UBound(arr, 2))).value

    ' Cleanup
    ActiveWorkbook.Worksheets(ws_index).Select
    Application.DisplayAlerts = False
    With ActiveWorkbook.Worksheets(sheetName)
        .Activate
        .Delete
    End With
    Application.DisplayAlerts = True

End Sub


Private Sub UpsertItemColumn(table As ListObject, itemColumnName As String, beforeColumnName As String)
    '
    ' Insert the Item column in the desired place, if not present
    ' Pull necessary values into and array
    ' Add an index field (to allow resort) and a blank field for results
    ' Sort the array (with the help of temporary sheet
    ' Apply the grouping algorithm
    ' Resort to original order
    ' Store results into target column, even if not empty
    '
    Dim arr() As Variant
    Dim column() As Variant
    Dim num_rows As Long
    Dim row As Long
    
    InsertColumnBefore table, itemColumnName, beforeColumnName
    
    arr = GetArray(table, Array("Manufacturer", "#SKU", "Cases (Product Detail)", beforeColumnName))
    num_rows = UBound(arr, 1)
    
    AddIndexToArray arr         ' index 5
    SortArray arr, [{1, 2, -3}]
    AddColumnToArray arr        ' index 6

    ' Initialize the first row
    arr(1, 6) = arr(1, 4)
    For row = 2 To num_rows
        ' manufacturer and SKU match previous row
        If arr(row, 1) = arr(row - 1, 1) And arr(row, 2) = arr(row - 1, 2) And Not IsEmpty(arr(row, 2)) Then
            ' copy previous description
            arr(row, 6) = arr(row - 1, 6)
        Else
            ' use current description
            arr(row, 6) = arr(row, 4)
        End If
    Next row
    
    'Re-Sort by the index column to get original order
    SortArray arr, [{5}]
    ' Write the result to the table
    table.ListColumns(itemColumnName).DataBodyRange.value = GetArrayColumn(arr, 6)
End Sub


Sub UpsertQtrColumn(table As ListObject, columnName As String, dateColumnName As String)
    '
    Dim column As Variant
    Dim row As Long
    Dim m As Variant
    Dim y As Variant
    Dim value As Variant
    
    InsertColumnBefore table, columnName, dateColumnName
    column = table.ListColumns(dateColumnName).DataBodyRange

    For row = 1 To UBound(column, 1)
        m = month(column(row, 1))
        y = Year(column(row, 1))
        value = Application.WorksheetFunction.RoundUp(m / 3, 0)
        column(row, 1) = "Q" & value & "-" & y
    Next row
    table.ListColumns(columnName).DataBodyRange.value = column
    
End Sub


Sub UpsertSYHalfColumn(table As ListObject, columnName As String, dateColumnName As String)
    '
    Dim column As Variant
    Dim row As Long
    Dim m As Variant
    Dim y As Variant
    Dim value As Variant
    
    InsertColumnBefore table, columnName, dateColumnName
    column = table.ListColumns(dateColumnName).DataBodyRange

    For row = 1 To UBound(column, 1)
        m = month(column(row, 1))
        y = Year(column(row, 1))
        value = Application.WorksheetFunction.RoundUp(m / 6, 0)
        If value = 1 Then
            column(row, 1) = "1H-" & y - 1 & "-" & y - 2000       ' 1H-<last>-<this>
        ElseIf value = 2 Then
            column(row, 1) = "2H-" & y & "-" & y + 1 - 2000       ' 2H-<this>-<next>
        Else
            column(row, 1) = value & "-" & y                    ' This shouldn't happen
        End If
    Next row
    table.ListColumns(columnName).DataBodyRange.value = column
End Sub


Sub UpsertSYColumn(table As ListObject, columnName As String, dateColumnName As String)
    '
    Dim column As Variant
    Dim row As Long
    Dim m As Variant
    Dim y As Variant
    Dim value As Variant
    
    InsertColumnBefore table, columnName, dateColumnName
    column = table.ListColumns(dateColumnName).DataBodyRange

    For row = 1 To UBound(column, 1)
        m = month(column(row, 1))
        y = Year(column(row, 1))
        value = Application.WorksheetFunction.RoundUp(m / 6, 0)
        If value = 1 Then
            column(row, 1) = y - 1 & "-" & y - 2000       ' <last>-<this>
        ElseIf value = 2 Then
            column(row, 1) = y & "-" & y + 1 - 2000       ' <this>-<next>
        Else
            column(row, 1) = value & "-" & y        ' This shouldn't happen
        End If
    Next row
    table.ListColumns(columnName).DataBodyRange.value = column
End Sub


Sub UpsertYearColumn(table As ListObject, columnName As String, dateColumnName As String)
    '
    Dim column As Variant
    Dim row As Long
    
    InsertColumnBefore table, columnName, dateColumnName
    column = table.ListColumns(dateColumnName).DataBodyRange

    For row = 1 To UBound(column, 1)
        column(row, 1) = Year(column(row, 1))
    Next row
    table.ListColumns(columnName).DataBodyRange.value = column
End Sub

Sub PrepareTable()
    '
    ' Mfg_Analysis_Start_1
    ' Re-entrant code. We don't need to call this directly.
    '
    Dim table As ListObject
    Dim cur_sheet As Worksheet
    Dim cur_range As Range
    
    ' Save so we can restore at end
    Set cur_sheet = ActiveSheet
    Set cur_range = Selection

    Set table = GetDataTable("DataTable")
    
    UpsertQtrColumn table, "Qtr", "Date"
    UpsertSYHalfColumn table, "SY-Half", "Date"
    UpsertSYColumn table, "SY", "Date"
    UpsertYearColumn table, "Year", "Date"

    UpsertItemColumn table, "Item Description", "PRODUCT_DESCRIPTION"
    UpsertItemColumn table, "Item Pack", "Pack Size"
    
    ' Restore original seletion
    cur_sheet.Select
    cur_range.Select
End Sub



