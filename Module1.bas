Attribute VB_Name = "Module1"

Sub Covid_Track_Start_1()
'
' Covid_Track_Start_1
'

'
'COMMENT: make data a table, then sum all the months for IPS Overall Growth
    Range("A1").Select
    Selection.EntireRow.Insert
    Selection.EntireRow.Insert
    Selection.EntireRow.Insert
    ActiveCell.FormulaR1C1 = "IPS CoVid Tracker"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=TODAY()"
    Range("A2").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("4:4").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("A5").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$5:$O$537"), , xlYes).Name = _
        "Table1"
    Range("A1").Select
    With Selection.Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Selection.Font.Bold = True
    Range("A3").Select
    Selection.EntireRow.Insert
    Selection.EntireRow.Insert
    Selection.EntireRow.Insert
    Range("Q2").Select
    ActiveCell.FormulaR1C1 = "2019"
    Range("Q3").Select
    ActiveCell.FormulaR1C1 = "2020"
    Range("A1").Select
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    Range("Table1[[#Headers],[January]]").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.End(xlUp).Select
    Range("D2").Select
    Selection.End(xlToRight).Select
    Range("R1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Font.Bold = True
    Selection.Font.Underline = xlUnderlineStyleSingle
    Range("R2").Select
    ActiveCell.FormulaR1C1 = "=SUMIFS(C[-14],C2,RC17,C3,""VOLUME"")"
    Range("R2").Select
    Selection.Copy
    Range("R2:AC3").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("R4").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C/R[-2]C-1"
    Range("R4").Select
    Selection.style = "Percent"
    Selection.Copy
    Range("S4:AC4").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("Q4").Select
    ActiveCell.FormulaR1C1 = "IPS OVERALL"
    Range("Q5").Select
    columns("Q:Q").EntireColumn.AutoFit
    Range("Q4").Select
    Range(Selection, Selection.End(xlUp)).Select
    Selection.Font.Bold = True
    Range("Q5").Select
'COMMENT: filter to 2020 Volume % so we're ready to search for particular manufacturers
        ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=2, Criteria1:= _
        "2020"
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=3, Criteria1:= _
        "% Difference in VOLUME from the Previous along Year of Period Date"
    Range("A7").Select
'COMMENT: Make a text copy of the table so as not to have issues with formulas above
    Range("Q6").Select
    Selection.EntireRow.Insert
    Selection.EntireRow.Insert
    Selection.EntireRow.Insert
    Range("Q1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("Q4").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Range("Q7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("R1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("R6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
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


Sub Mfg_Analysis_Start_1()
'
' Mfg_Analysis_Start_1
'

'
    Dim DataTable As Range
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Set DataTable = Selection
    ActiveSheet.ListObjects.Add(xlSrcRange, DataTable, , xlYes). _
        Name = "Table1"
    
    columns("K:K").Select
    Selection.Insert Shift:=xlToRight
    columns("M:M").Select
    Selection.Insert Shift:=xlToRight
    columns("P:P").Select
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Selection.Insert Shift:=xlToRight
    Range("Table1[[#Headers],[Column5]]").Select
    ActiveCell.FormulaR1C1 = "School Year"
    Range("Table1[[#Headers],[Column4]]").Select
    ActiveCell.FormulaR1C1 = "1H School Year"
    Range("Table1[[#Headers],[Column3]]").Select
    ActiveCell.FormulaR1C1 = "Year"
    Range("Table1[[#Headers],[School Year]:[Year]]").Select
    Selection.style = "Good"
    Range("Table1[[#Headers],[Column2]]").Select
    Selection.style = "Good"
    ActiveCell.FormulaR1C1 = "Item Pack"
    Range("Table1[[#Headers],[Column1]]").Select
    ActiveCell.FormulaR1C1 = "Item Description"
    Range("Table1[[#Headers],[Item Description]]").Select
    Selection.style = "Good"
    Range("R9").Select
    ActiveCell.FormulaR1C1 = "=YEAR(RC[1])"
    Range("R9").Select
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
    Range(Selection, Selection.End(xlDown)).Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
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
    Range("R9").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("Table1[[#Headers],[Date]]").Select
    Application.CutCopyMode = False

' School_Year_Columns
    columns("S:S").Select
    Selection.NumberFormat = "m/d/yyyy"
    Range("Q9").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=19, Operator:= _
        xlFilterValues, Criteria2:=Array(1, "1/1/2020", 1, "2/1/2020", 1, "3/1/2020", 1, _
        "4/1/2020", 1, "5/1/2020", 1, "6/1/2020", 1, "7/1/2019", 1, "8/1/2019", 1, "9/1/2019", 1, _
        "10/1/2019", 1, "11/1/2019", 1, "12/1/2019")
    Range("P9").Select
    ActiveCell.FormulaR1C1 = "2019-20 SY"
    Range("P9").Select
    Selection.Copy
    Range("P18").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=19, Operator:= _
        xlFilterValues, Criteria2:=Array(1, "1/1/2019", 1, "2/1/2019", 1, "3/1/2019", 1, _
        "4/1/2019", 1, "5/1/2019", 1, "6/1/2019", 1, "7/1/2018", 1, "8/1/2018", 1, "9/1/2018", 1, _
        "10/1/2018", 1, "11/1/2018", 1, "12/1/2018")
    Range("P10").Select
    ActiveCell.FormulaR1C1 = "2018-19 SY"
    Range("P10").Select
    Selection.Copy
    Range("P11").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=19
    
        ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=19, Operator:= _
        xlFilterValues, Criteria2:=Array(1, "7/1/2019", 1, "8/1/2019", 1, "9/1/2019", 1, _
        "10/1/2019", 1, "11/1/2019", 1, "12/1/2019")
    Range("Q9").Select
    ActiveCell.FormulaR1C1 = "Jul-Dec 2019"
    Range("Q9").Select
    Selection.Copy
    Range("Q18").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=19, Operator:= _
        xlFilterValues, Criteria2:=Array(1, "7/1/2018", 1, "8/1/2018", 1, "9/1/2018", 1, _
        "10/1/2018", 1, "11/1/2018", 1, "12/1/2018")
    Range("Q70002").Select
    ActiveCell.FormulaR1C1 = "Jul-Dec 2018"
    Range("Q70002").Select
    Selection.Copy
    Range("Q70005").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=19, Operator:= _
        xlFilterValues, Criteria2:=Array(1, "7/1/2020", 1, "8/1/2020", 1, "9/1/2020", 1, _
        "10/1/2020")
    Range("Q49044").Select
    ActiveCell.FormulaR1C1 = "Jul-Dec 2020 YTD"
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range("Q49044").Select
    ActiveCell.FormulaR1C1 = "Jul-Sept 2020 YTD"
    Range("Q49044").Select
    Selection.Copy
    Range("Q49045").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=19
    Range("Table1[[#Headers],[Year]]").Select
    Selection.EntireColumn.Insert
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=20, Operator:= _
        xlFilterValues, Criteria2:=Array(1, "1/1/2020", 1, "2/1/2020")
    Range("Table1[[#Headers],[Column1]]").Select
    ActiveCell.FormulaR1C1 = "Jan-Feb Year"
    Range("R49046").Select
    ActiveCell.FormulaR1C1 = "Jan-Feb 2020"
    Range("R49046").Select
    Selection.Copy
    Range("R49049").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=20, Operator:= _
        xlFilterValues, Criteria2:=Array(1, "1/1/2019", 1, "2/1/2019")
    Range("R12").Select
    ActiveCell.FormulaR1C1 = "Jan-Feb 2019"
    Range("R12").Select
    Selection.Copy
    Range("R13").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=20, Operator:= _
        xlFilterValues, Criteria2:=Array(1, "1/1/2018", 1, "2/1/2018")
    Range("R70045").Select
    Selection.End(xlUp).Select
    Range("R70003").Select
    ActiveCell.FormulaR1C1 = "Jan-Feb 2018"
    Range("R70003").Select
    Selection.Copy
    Range("R70004").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=20

End Sub

Sub Multi_Mfg_Item_Pack()
'
' Multi_Mfg_Item_Pack Macro
'

'
    Range("Table1[[#Headers],[Manufacturer]]").Select
    ActiveWorkbook.Worksheets("Manufacturer Analysis (7)").ListObjects("Table1"). _
        Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Manufacturer Analysis (7)").ListObjects("Table1"). _
        Sort.SortFields.Add2 Key:=Range("Table1[Manufacturer]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Manufacturer Analysis (7)").ListObjects("Table1"). _
        Sort.SortFields.Add2 Key:=Range("Table1['#SKU]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Manufacturer Analysis (7)").ListObjects("Table1"). _
        Sort.SortFields.Add2 Key:=Range("Table1[Cases (Product Detail)]"), SortOn:= _
        xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Manufacturer Analysis (7)").ListObjects( _
        "Table1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Single_Mfg_Item_Sort()
'
' Single_Mfg_Item_Sort Macro
'

'
    Range("Table1[[#Headers],[Manufacturer]]").Select
    ActiveWorkbook.Worksheets("Manufacturer Analysis (7)").ListObjects("Table1"). _
        Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Manufacturer Analysis (7)").ListObjects("Table1"). _
        Sort.SortFields.Add2 Key:=Range("Table1['#SKU]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Manufacturer Analysis (7)").ListObjects("Table1"). _
        Sort.SortFields.Add2 Key:=Range("Table1[Cases (Product Detail)]"), SortOn:= _
        xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Manufacturer Analysis (7)").ListObjects( _
        "Table1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub Item_Pack_Formulas()
'
' Item_Pack_Formulas
'

'
    Range("K9").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(ISBLANK(RC[-2]),RC[1],IF(R[-1]C[-2]=RC[-2],R[-1]C,RC[1]))"
    Range("K9").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("M9").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=IF(ISBLANK(RC[-4]),RC[1],IF(R[-1]C[-4]=RC[-4],R[-1]C,RC[1]))"
    Range("M9").Select
    Selection.End(xlDown).Select
    Selection.End(xlUp).Select
    Range("M9").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.Replace What:="999999", Replacement:="9", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="999999", Replacement:="9", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="999999", Replacement:="9", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="999", Replacement:="9", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="999", Replacement:="9", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="999", Replacement:="9", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="999", Replacement:="9", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="999", Replacement:="9", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="999", Replacement:="9", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="0000000", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="0000000", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="0000000", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="000", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="000", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="000", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="000", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="000", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Selection.Replace What:="000", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    Range("Table1[[#Headers],[Item Pack]]").Select
    Application.CutCopyMode = False
    
End Sub


Sub Mfg_Pivot_CY()
'
' Mfg_Pivot_CY Macro
'

'
    Range("Table1[[#Headers],[Manufacturer]]").Select
    Application.CutCopyMode = False
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Table1", Version:=6).CreatePivotTable TableDestination:="Sheet1!R3C1", _
        tableName:="PivotTable2", DefaultVersion:=6
    Sheets("Sheet1").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable2")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable2").RepeatAllLabels xlRepeatLabels
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Manufacturer").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Member Name").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("MEMBER#").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Capacity").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("State").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("SECTOR").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("MAJOR_CAT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("MINOR_CAT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("#SKU").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("BRAND").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Item Description"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("PRODUCT_DESCRIPTION"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Item Pack").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Pack Size").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Weight (In Pounds)"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("School Year").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("1H School Year").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Jan-Feb Year").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Year").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Date").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("REBATE_INVOICE_DATE"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DISTRIBUTOR").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DISTRIBUTOR_NAME"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("NVD_RATE_BASIS_TYPE"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Distributor2").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("MANUFACTURER_NAME"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("PRODUCT_DESCRIPTION3"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Red flag (look into)"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("BILLED_REBATE_AMT"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Cases (NVD)").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Cases (Product Detail)"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Purchase $'s").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Total Order Weight"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").RowGrand = False
    ActiveSheet.PivotTables("PivotTable2").RowAxisLayout xlTabularRow
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Date")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Date").CurrentPage = _
        "(All)"
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Date")
        .PivotItems("1/1/2018").Visible = False
        .PivotItems("2/1/2018").Visible = False
        .PivotItems("3/1/2018").Visible = False
        .PivotItems("4/1/2018").Visible = False
        .PivotItems("5/1/2018").Visible = False
        .PivotItems("6/1/2018").Visible = False
        .PivotItems("7/1/2018").Visible = False
        .PivotItems("8/1/2018").Visible = False
        .PivotItems("9/1/2018").Visible = False
        .PivotItems("10/1/2018").Visible = False
        .PivotItems("11/1/2018").Visible = False
        .PivotItems("12/1/2018").Visible = False
        .PivotItems("1/1/2020").Visible = False
        .PivotItems("2/1/2020").Visible = False
        .PivotItems("3/1/2020").Visible = False
        .PivotItems("4/1/2020").Visible = False
        .PivotItems("5/1/2020").Visible = False
        .PivotItems("6/1/2020").Visible = False
        .PivotItems("7/1/2020").Visible = False
        .PivotItems("8/1/2020").Visible = False
        .PivotItems("9/1/2020").Visible = False
        .PivotItems("10/1/2020").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Date"). _
        EnableMultiplePageItems = True
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Manufacturer")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Cases (Product Detail)"), _
        "Sum of Cases (Product Detail)", xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Purchase $'s"), "Sum of Purchase $'s", xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Total Order Weight"), "Sum of Total Order Weight", _
        xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Cases (NVD)"), "Sum of Cases (NVD)", xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("BILLED_REBATE_AMT"), "Sum of BILLED_REBATE_AMT", _
        xlSum
    Range("B4").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    columns("B:F").Select
    Selection.style = "Comma"
    Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    ActiveSheet.PivotTables("PivotTable2").PivotSelect _
        "'Sum of Cases (Product Detail)'", xlDataAndLabel, True
    Range("B3").Select
    ActiveSheet.PivotTables("PivotTable2").DataPivotField.PivotItems( _
        "Sum of Cases (Product Detail)").Caption = "Cases"
    Range("C3").Select
    ActiveSheet.PivotTables("PivotTable2").DataPivotField.PivotItems( _
        "Sum of Purchase $'s").Caption = "Purch$"
    Range("D3").Select
    ActiveSheet.PivotTables("PivotTable2").DataPivotField.PivotItems( _
        "Sum of Total Order Weight").Caption = "LBS"
    Range("E3").Select
    ActiveSheet.PivotTables("PivotTable2").DataPivotField.PivotItems( _
        "Sum of Cases (NVD)").Caption = "Cases NVD"
    Range("F3").Select
    ActiveSheet.PivotTables("PivotTable2").DataPivotField.PivotItems( _
        "Sum of BILLED_REBATE_AMT").Caption = "Billed NVD"
    Range("B3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ColumnWidth = 12
    Range("C1").Select
    Selection.EntireRow.Insert
    Selection.EntireRow.Insert
    Selection.EntireRow.Insert
    Selection.EntireRow.Insert
    Selection.EntireRow.Insert
    ActiveSheet.Next.Select
    Range("A1:C3").Select
    Selection.Copy
    ActiveSheet.Previous.Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("A3").Select
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Pivots CY"
    Range("A1").Select
End Sub


Sub CreatePivot1()
'
' Create the manufacturer CY pivot
'

'
    Range("Table1[[#Headers],[Manufacturer]]").Select
    Application.CutCopyMode = False
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Table1", Version:=6).CreatePivotTable TableDestination:="Sheet1!R3C1", _
        tableName:="PivotTable2", DefaultVersion:=6
    Sheets("Sheet1").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable2")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable2").RepeatAllLabels xlRepeatLabels
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Manufacturer").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Member Name").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("MEMBER#").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Capacity").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("State").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("SECTOR").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("MAJOR_CAT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("MINOR_CAT").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("#SKU").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("BRAND").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Item Description"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("PRODUCT_DESCRIPTION"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Item Pack").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Pack Size").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Weight (In Pounds)"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("School Year").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("1H School Year").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Jan-Feb Year").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Year").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Date").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("REBATE_INVOICE_DATE"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DISTRIBUTOR").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("DISTRIBUTOR_NAME"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("NVD_RATE_BASIS_TYPE"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Distributor2").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("MANUFACTURER_NAME"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("PRODUCT_DESCRIPTION3"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Red flag (look into)"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("BILLED_REBATE_AMT"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Cases (NVD)").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Cases (Product Detail)"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Purchase $'s").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Total Order Weight"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotTable2").RowGrand = False
    ActiveSheet.PivotTables("PivotTable2").RowAxisLayout xlTabularRow
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Date")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Date").CurrentPage = _
        "(All)"
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Date")
        .PivotItems("1/1/2018").Visible = False
        .PivotItems("2/1/2018").Visible = False
        .PivotItems("3/1/2018").Visible = False
        .PivotItems("4/1/2018").Visible = False
        .PivotItems("5/1/2018").Visible = False
        .PivotItems("6/1/2018").Visible = False
        .PivotItems("7/1/2018").Visible = False
        .PivotItems("8/1/2018").Visible = False
        .PivotItems("9/1/2018").Visible = False
        .PivotItems("10/1/2018").Visible = False
        .PivotItems("11/1/2018").Visible = False
        .PivotItems("12/1/2018").Visible = False
        .PivotItems("1/1/2020").Visible = False
        .PivotItems("2/1/2020").Visible = False
        .PivotItems("3/1/2020").Visible = False
        .PivotItems("4/1/2020").Visible = False
        .PivotItems("5/1/2020").Visible = False
        .PivotItems("6/1/2020").Visible = False
        .PivotItems("7/1/2020").Visible = False
        .PivotItems("8/1/2020").Visible = False
        .PivotItems("9/1/2020").Visible = False
        .PivotItems("10/1/2020").Visible = False
    End With
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Date"). _
        EnableMultiplePageItems = True
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Manufacturer")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Cases (Product Detail)"), _
        "Sum of Cases (Product Detail)", xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Purchase $'s"), "Sum of Purchase $'s", xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Total Order Weight"), "Sum of Total Order Weight", _
        xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Cases (NVD)"), "Sum of Cases (NVD)", xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("BILLED_REBATE_AMT"), "Sum of BILLED_REBATE_AMT", _
        xlSum
    Range("B4").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Range(Selection, Selection.End(xlUp)).Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    columns("B:F").Select
    Selection.style = "Comma"
    Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
    ActiveSheet.PivotTables("PivotTable2").PivotSelect _
        "'Sum of Cases (Product Detail)'", xlDataAndLabel, True
    Range("B3").Select
    ActiveSheet.PivotTables("PivotTable2").DataPivotField.PivotItems( _
        "Sum of Cases (Product Detail)").Caption = "Cases"
    Range("C3").Select
    ActiveSheet.PivotTables("PivotTable2").DataPivotField.PivotItems( _
        "Sum of Purchase $'s").Caption = "Purch$"
    Range("D3").Select
    ActiveSheet.PivotTables("PivotTable2").DataPivotField.PivotItems( _
        "Sum of Total Order Weight").Caption = "LBS"
    Range("E3").Select
    ActiveSheet.PivotTables("PivotTable2").DataPivotField.PivotItems( _
        "Sum of Cases (NVD)").Caption = "Cases NVD"
    Range("F3").Select
    ActiveSheet.PivotTables("PivotTable2").DataPivotField.PivotItems( _
        "Sum of BILLED_REBATE_AMT").Caption = "Billed NVD"
    Range("B3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.ColumnWidth = 12
    Range("C1").Select
    Selection.EntireRow.Insert
    Selection.EntireRow.Insert
    Selection.EntireRow.Insert
    Selection.EntireRow.Insert
    Selection.EntireRow.Insert
    ActiveSheet.Next.Select
    Range("A1:C3").Select
    Selection.Copy
    ActiveSheet.Previous.Select
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Range("A3").Select
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Pivots CY"
    Range("A1").Select
End Sub


Sub CreatePivot2()
'
' Create Mfg_Pivot_YOY
'

'
    Range("A10").Select
    Application.CutCopyMode = False
    ActiveSheet.PivotTables("PivotTable2").PivotCache.Refresh
    columns("A:F").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "Pivots YOY"
    Range("A10").Select
    With ActiveSheet.PivotTables("PivotTable2").PivotFields( _
        "Same Store 2018-19 SY to 2020-21 SYTD")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable3").PivotFields( _
        "Same Store 2018-19 SY to 2020-21 SYTD")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable3").PivotFields( _
        "Same Store 2018-19 SY to 2020-21 SYTD").ClearAllFilters
    ActiveSheet.PivotTables("PivotTable3").PivotFields( _
        "Same Store 2018-19 SY to 2020-21 SYTD").CurrentPage = "Y"
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "IPS YOY Volumes For:"
    Range("A2").Select
    Sheets("Pivots YOY").Select
    Sheets("Pivots YOY").Move Before:=Sheets(1)
End Sub

Sub TestMacroTables()
' this is a test of creating a table better

    Dim DataTable As Range
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Set DataTable = Selection
    ActiveSheet.ListObjects.Add(xlSrcRange, DataTable, , xlYes). _
        Name = "Table1"
    
    


End Sub
