Attribute VB_Name = "Module3"
Option Base 1


Public Sub UseCollection()
    Dim coll As New Collection
    
    coll.Add "Apple"
    coll.Add "Orange"
    coll.Add "Pear"
    
    Total = coll.Count
    
    fruit = coll.Item(1)
    
    Debug.Print fruit
End Sub

Sub ReadingRange()

    Dim arr As Variant
    arr = ActiveSheet.Range("A1").CurrentRegion.Value
    
    Debug.Print LBound(arr, 1); UBound(arr, 1)

End Sub

Function CompareRows(ByRef arr() As Variant, key_cols As Variant, row_a As Long, row_b As Long) As Boolean
    ' Allow "negative" key_col meaning to sort descending
    Dim row_i As Long
    Dim row_j As Long
    Dim value_i As Variant
    Dim value_j As Variant
    
    For col = 1 To UBound(key_cols)
        key_col = key_cols(col)
        If key_col > 0 Then
            row_i = row_a
            row_j = row_b
        Else
            row_i = row_b
            row_j = row_a
            key_col = -key_col
        End If
        value_i = arr(row_i, key_col)
        value_j = arr(row_i, key_col)
        If value_i = Empty Then
            CompareRows = False
            Exit Function
        End If
        If value_j = Empty Then
            CompareRows = True
            Exit Function
        End If
        If value_i < value_j Then
            CompareRows = True
            Exit Function
        End If
        If value_i > value_j Then
            CompareRows = False
            Exit Function
        End If
        ' the first col values must be equal for each row at this point
    Next col
    CompareRows = False
End Function

Sub SwapRows(ByRef arr() As Variant, row_a As Long, row_b As Long, Optional num_cols As Long = -1)
    Dim temp As Variant
    If num_cols = -1 Then num_cols = SizeOfArray(arr, 2)
    
    For col = 1 To num_cols
        temp = arr(row_a, col)
        arr(row_a, col) = arr(row_b, col)
        arr(row_b, col) = temp
    Next col
End Sub

Sub Quicksort(ByRef arr() As Variant, key_cols As Variant, Optional low As Long = -1, Optional high As Long = -1)
    'Sorts a multi-dimensional VBA array. Negative key_col implies descending else ascending order.
    Dim pivot As Long
    Dim i As Long
    Dim j As Long
 
    If low = -1 Then low = LBound(arr, 1)
    If high = -1 Then high = UBound(arr, 1)
    If low >= high Then Exit Sub

    i = low
    j = high
    pivot = (low + high) \ 2
 
    While i <= j
        While CompareRows(arr, key_cols, i, pivot) And i < high
            i = i + 1
        Wend
  
        While CompareRows(arr, key_cols, pivot, j) And j > low
            j = j - 1
        Wend
 
        If i <= j Then
            Call SwapRows(arr, i, j)
            i = i + 1
            j = j - 1
        End If
    Wend
 
    If low < j Then Call Quicksort(arr, key_cols, low, j)
    If i < high Then Call Quicksort(arr, key_cols, i, high)
End Sub

Sub UseQuicksort()
    Dim arr()  As Variant
    arr = [{"one", 1, 5; "two", 2, 6; "three", 3, 7; "four", 4, 8}]
    Call Quicksort(arr, [{-1}])
    Call Quicksort(arr, [{2}])
End Sub

Function SizeOfArray(ByRef arr As Variant, Optional dimension = 1)
    SizeOfArray = 0
    On Error Resume Next
    SizeOfArray = UBound(arr, dimension) - LBound(arr, dimension) + 1
End Function

Function GetArrayFromTable(tableName As String, ByRef columnNames As Variant) As Variant
    Dim arr() As Variant
    Dim column As Variant
    Dim num_rows As Long
    Dim num_cols As Long
    
    If Not TableExists(tableName) Then
        GetArrayFromTable = arr
        Exit Function
    End If
    
    Set table = ActiveSheet.ListObjects(tableName)
    num_rows = table.DataBodyRange.Rows.Count
    num_cols = SizeOfArray(columnNames)
    ReDim arr(num_rows, num_cols)
    
    For col = 1 To num_cols
        columnName = columnNames(col - 1) ' zero-based
        column = table.ListColumns(columnName).DataBodyRange
        ' Copy data into array
        For row = 1 To num_rows
            arr(row, col) = column(row, 1)
        Next row
    Next col
    GetArrayFromTable = arr
End Function

Sub UseGetArray()
    Dim size As Long
    Dim arr() As Variant
    Dim columnNames As Variant
    Dim num_rows As Long
    Dim num_cols As Long
    
    ' Add columns with the first ones as sort keys
    columnNames = Split("Manufacturer,#SKU,PRODUCT_DESCRIPTION,Cases (NVD)", ",")  ' zero-based
    arr = GetArrayFromTable("DataTable", columnNames)
    num_rows = SizeOfArray(arr, 1)
    num_cols = SizeOfArray(arr, 2)
    ' add index column + output column to data array
    ReDim Preserve arr(num_rows, num_cols + 2) As Variant
    ' Create index values
    For row = 1 To num_rows
        arr(row, num_cols + 1) = row
    Next row
    Call Quicksort(arr, [{1,-2}])
    If SheetExists("Results") Then
        Application.DisplayAlerts = False
        ActiveWorkbook.Worksheets("Results").Delete
        Application.DisplayAlerts = True
    End If
    ActiveWorkbook.Worksheets.Add.Name = "Results"
    'ActiveWorkbook.Worksheets("Results").Range("A1").CurrentRegion.Select
    'Selection.ClearContents
    ActiveWorkbook.Worksheets("Results").Range("A1").Resize(UBound(arr, 1), UBound(arr, 2)) = arr
End Sub
