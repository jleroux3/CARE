Attribute VB_Name = "Module3"
Sub MacroTest()
Attribute MacroTest.VB_ProcData.VB_Invoke_Func = "e\n14"
'
' MacroTest Macro
'
' Keyboard Shortcut: Ctrl+e
'
    Application.Run "PERSONAL.XLSB!GenerateReports"
    Sheets("Tier1_Actual").Select
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Rows("3:3").Select
    Selection.Cut
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "CONFIDENTIAL"
    Range("A59").Select
    ActiveWindow.SmallScroll Down:=-66
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "AB 2398 Monthly Rolling Forecast"
    Range("A59").Select
    ActiveWindow.SmallScroll Down:=-57
    Rows("4:4").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A4").Select
    ActiveCell.FormulaR1C1 = _
        "Number of Full Time Equivalent (FTE) Employees in State of California working on carpet recycling"
    Rows("4:4").Select
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
    Range("A4:I4").Select
    Range("I4").Activate
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
    Range("A4:F4").Select
    Range("F4").Activate
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
    Selection.Merge
    Rows("5:5").Select
    Selection.Cut
    Rows("4:4").Select
    Selection.Insert Shift:=xlDown
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "Actual_Num_FTE_CA_Emp_BeginQ"
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Range("A5").Select
    ActiveCell.FormulaR1C1 = _
        "Number of Ca FTE Employees at beginning of this quarter"
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "Number of FTE CA Jobs lost this quarter"
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "Number of FTE CA Jobs gained this quarter"
    Range("A8").Select
    ActiveCell.FormulaR1C1 = "Number of FTE CA Employees at end of this quarter"
    Range("A9").Select
    Columns("A:A").EntireColumn.AutoFit
    ActiveWindow.SmallScroll ToRight:=1
    ActiveWindow.LargeScroll ToRight:=1
    ActiveWindow.SmallScroll ToRight:=1
    ActiveWindow.LargeScroll ToRight:=-1
    ActiveWindow.SmallScroll ToRight:=-2
    Range("A5").Select
    ActiveCell.FormulaR1C1 = _
        "Number of CA FTE Employees at beginning of this quarter"
    Rows("8:8").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Cut
    Rows("10:10").Select
    Selection.Insert Shift:=xlDown
    Range("A9").Select
    ActiveCell.FormulaR1C1 = _
        "Post-consumer carpet pounds directly collected by you for this quarter (Do NOT report pounds you are purchasing from other collectors)"
    Range("A9:F9").Select
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
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("A10").Select
    ActiveCell.FormulaR1C1 = _
        "Post-consumer carpet pounds directly collected by you from California for this quarter"
    Range("A11").Select
    Columns("A:A").EntireColumn.AutoFit
    Range("A11").Select
    ActiveCell.FormulaR1C1 = _
        "Post-consumer carpet pounds collected by you from OUTSIDE California for this quarter"
    Range("A12").Select
    Columns("A:A").EntireColumn.AutoFit
    Range("A11").Select
    ActiveCell.FormulaR1C1 = _
        "Post-consumer carpet pounds directly collected by you from OUTSIDE California for this quarter"
    Range("A13").Select
    Columns("A:A").ColumnWidth = 74
    Range("A12").Select
    ActiveCell.FormulaR1C1 = "TOTAL Post-consumer carpet pounds"
    Range("A55").Select
    ActiveWindow.SmallScroll Down:=-141
    Range("A13").Select
    ActiveCell.FormulaR1C1 = _
        "Carpet directly collected by YOU from Califronia by FIBER type (Do NOT report pounds you are purchasing from other collectors)"
    Range("F14").Select
    Rows("12:12").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Cut
    Rows("14:14").Select
    Selection.Insert Shift:=xlDown
    Range("A13").Select
    ActiveCell.FormulaR1C1 = _
        "Carpet directly collected by YOU from california by FIBER type (Do NOT report pounds you are purchasing from other collectors)"
    Range("A13:F13").Select
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("A14").Select
    ActiveCell.FormulaR1C1 = "Nylon 6"
    Range("A15").Select
    ActiveCell.FormulaR1C1 = "Nylon 6,6"
    Range("A15").Select
    ActiveCell.FormulaR1C1 = "Nylon 6, 6"
    Range("A16").Select
    ActiveCell.FormulaR1C1 = "Polypropylene"
    Range("A17").Select
    ActiveCell.FormulaR1C1 = "PET"
    Range("A18").Select
    ActiveCell.FormulaR1C1 = "Wool"
    Range("A19").Select
    ActiveCell.FormulaR1C1 = "Other/Mixed Fibers"
    Range("A20").Select
    ActiveCell.FormulaR1C1 = "TOTAL Post-consumer carpet pounds"
    Rows("21:21").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A21").Select
    ActiveCell.FormulaR1C1 = "Line20mustequalLine10"
    Rows("22:22").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A22").Select
    ActiveCell.FormulaR1C1 = _
        "Accounting for total PC Carpet Inputs & Beginning Inventory this quarter"
    Range("A22:F22").Select
    Rows("12:12").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A12").Select
    ActiveCell.FormulaR1C1 = "TOTAL Post-consumer carpet pounds"
    Rows("13:13").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A13").Select
    ActiveCell.FormulaR1C1 = _
        "Carpet directly collected by YOU from California by FIBER type (Do NOT report pounds you are purchasing from other collectors)"
    Range("A13:F13").Select
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
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("A14").Select
    ActiveCell.FormulaR1C1 = "Nylon 6"
    Range("A15").Select
    Range("A15").Select
    ActiveCell.FormulaR1C1 = "Actual_Nylon6_6_CPT_Collected_CA"
    Range("A15").Select
    ActiveCell.FormulaR1C1 = "Nylon 6, 6"
    Range("A16").Select
    ActiveCell.FormulaR1C1 = "Polypropylene"
    Range("A17").Select
    ActiveCell.FormulaR1C1 = "PET"
    Range("A18").Select
    ActiveCell.FormulaR1C1 = "Wool"
    Range("A19").Select
    Range("A19").Select
    ActiveCell.FormulaR1C1 = "Actual_Other_MF_CPT_Collected_CA"
    Range("A19").Select
    ActiveCell.FormulaR1C1 = "Other/Mixed Fibers"
    Range("A20").Select
    Rows("20:20").Select
    Selection.Insert Shift:=xlDown
    Range("A20").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Range("A21").Select
    ActiveCell.FormulaR1C1 = "Actual_Beg_Inv_WCPT_CA_Qtr_Beg"
    Range("A21").Select
    Rows("21:21").Select
    Selection.Insert Shift:=xlDown
    Range("A21").Select
    Range("A21").Select
    ActiveCell.FormulaR1C1 = "Line 20 must equal Line 10"
    Rows("22:22").Select
    Selection.Insert Shift:=xlDown
    Range("A22").Select
    ActiveCell.FormulaR1C1 = _
        "Accounting for total PC Carpet Outputs & Beginning Inventory this quarter"
    Range("A22:F22").Select
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
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("A23").Select
    ActiveCell.FormulaR1C1 = "Actual_Beg_Inv_WCPT_CA_Qtr_Beg"
    Range("A23").Select
    ActiveCell.FormulaR1C1 = _
        "Beginning Inventory of Whole Carpet from CA at start of quarter (should equal prior quarter ending inventory"
    Range("A24").Select
    ActiveCell.FormulaR1C1 = "Actual_WCPT_Collected_CA"
    Range("A24").Select
    ActiveCell.FormulaR1C1 = "Whole Carpet Collected from California (Row 10)"
    Range("A25").Select
    ActiveCell.FormulaR1C1 = "Whole carpet from CA received from other collectors"
    Rows("26:26").Select
    Selection.Insert Shift:=xlDown
    Range("A26:F26").Select
    Range("A26").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Rows("27:27").Select
    Selection.Insert Shift:=xlDown
    Range("A27").Select
    ActiveCell.FormulaR1C1 = _
        "Accounting for total PC Carpet Outputs & Ending Inventory"
    Range("A28").Select
    ActiveCell.FormulaR1C1 = "Actual_CPT_Out_Reused"
    Range("A28").Select
    ActiveCell.FormulaR1C1 = "Re-Used"
    Range("A29").Select
    ActiveCell.FormulaR1C1 = "Internally Used Whole Carpet this quarter"
    Range("A30").Select
    ActiveCell.FormulaR1C1 = _
        "Whole carpet shipped to US customers OUTSIDE California"
    Range("A31").Select
    ActiveCell.FormulaR1C1 = _
        "Whole carpet shipped to customers outside the United States"
    Range("A32").Select
    ActiveCell.FormulaR1C1 = "Whole carpet shipped to customers INSIDE California"
    Range("A33").Select
    ActiveCell.FormulaR1C1 = _
        "Non-carpet materials with value (i.e. carpet cushion)"
    Range("A34").Select
    ActiveCell.FormulaR1C1 = "WTE"
    Range("A35").Select
    ActiveWindow.SmallScroll Down:=27
    Range("A35").Select
    ActiveCell.FormulaR1C1 = "Incinerated"
    Range("A36").Select
    ActiveCell.FormulaR1C1 = "Landfilled"
    Range("A37").Select
    ActiveCell.FormulaR1C1 = "Ending Inventory of Whole Carpet"
    Rows("38:38").Select
    Selection.Insert Shift:=xlDown
    Range("A38").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Rows("39:39").Select
    Selection.Insert Shift:=xlDown
    Range("A39").Select
    ActiveCell.FormulaR1C1 = "Line 38 must equal Line 26"
    Range("A39").Select
    ActiveCell.FormulaR1C1 = "Line 38 must equal Line 26"
    Rows("40:40").Select
    Selection.Insert Shift:=xlDown
    Range("A40").Select
    ActiveCell.FormulaR1C1 = "Production of Internally Used Whole Carpet"
    Range("A41").Select
    ActiveCell.FormulaR1C1 = "Internally Used Whole Carpet"
    Range("A42").Select
    ActiveCell.FormulaR1C1 = "Processed"
    Range("A43").Select
    ActiveCell.FormulaR1C1 = "Landfilled"
    Range("A44").Select
    ActiveCell.FormulaR1C1 = "WTE"
    Range("A45").Select
    ActiveCell.FormulaR1C1 = "Incinerated"
    Rows("46:46").Select
    Selection.Insert Shift:=xlDown
    Range("A46").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Rows("47:47").Select
    Selection.Insert Shift:=xlDown
    ActiveCell.FormulaR1C1 = "Line 46 must equal Line 41"
    Rows("48:48").Select
    Selection.Insert Shift:=xlDown
    Range("A48").Select
    ActiveCell.FormulaR1C1 = _
        "Output and other destinations of post-consumer carpet internally processed this uarter"
    Range("A49").Select
    Range("A48").Select
    ActiveCell.FormulaR1C1 = _
        "Output and other destinations of post-consumer carpet internally processed this quarter"
    Range("A49").Select
    ActiveCell.FormulaR1C1 = _
        "Beginning Inventory of Processed Goods from prior quarter"
    Range("A50").Select
    ActiveCell.FormulaR1C1 = "Processed"
    Range("A51").Select
    ActiveCell.FormulaR1C1 = "Actual_TypeI_Out_Fiber"
    Rows("51:51").Select
    Selection.Insert Shift:=xlDown
    Range("A51").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Rows("52:52").Select
    Selection.Insert Shift:=xlDown
    Range("A52").Select
    ActiveCell.FormulaR1C1 = "Type 1 Outputs"
    Range("A53").Select
    ActiveCell.FormulaR1C1 = "Actual_TypeI_Out_Fiber"
    Range("A53").Select
    ActiveCell.FormulaR1C1 = "Fiber"
    Range("A54").Select
    ActiveCell.FormulaR1C1 = "DePoly or Chemical Component"
    Range("A55").Select
    ActiveCell.FormulaR1C1 = "Shredded Carpet tile used for tile backing"
    Range("A56").Select
    ActiveCell.FormulaR1C1 = _
        "Number of Ash Tests run this quarter (min 1 per 1M pounds)"
    Range("A57").Select
    ActiveCell.FormulaR1C1 = _
        "Average Ash Test Results over quarter for Type 1pounds"
    Rows("58:58").Select
    Selection.Insert Shift:=xlDown
    Range("A58").Select
    ActiveCell.FormulaR1C1 = "Total Type 1 Output: SOLD & SHIPPED"
    Rows("59:59").Select
    Selection.Insert Shift:=xlDown
    Range("A59").Select
    ActiveCell.FormulaR1C1 = "Type 2 Outputs"
    Range("A60").Select
    ActiveCell.FormulaR1C1 = "Filler"
    Range("A61").Select
    ActiveCell.FormulaR1C1 = "CAAF"
    Range("A62").Select
    ActiveWindow.ScrollRow = 28
    ActiveWindow.ScrollRow = 29
    ActiveWindow.ScrollRow = 30
    ActiveWindow.ScrollRow = 31
    ActiveWindow.ScrollRow = 32
    ActiveWindow.ScrollRow = 33
    ActiveWindow.ScrollRow = 34
    ActiveWindow.ScrollRow = 35
    ActiveWindow.ScrollRow = 36
    ActiveWindow.ScrollRow = 37
    ActiveWindow.ScrollRow = 38
    ActiveWindow.ScrollRow = 39
    ActiveWindow.ScrollRow = 40
    ActiveWindow.ScrollRow = 41
    ActiveWindow.ScrollRow = 42
    Range("A62").Select
    ActiveCell.FormulaR1C1 = "Cement Kiln feedstock"
    Range("A63").Select
    ActiveCell.FormulaR1C1 = "Carcass Sold"
    Rows("64:64").Select
    Selection.Insert Shift:=xlDown
    Range("A64").Select
    ActiveCell.FormulaR1C1 = "Total Type 2 Output: SOLD & SHIPPED"
    Range("A65").Select
    ActiveCell.FormulaR1C1 = "Landfilled"
    Range("A66").Select
    ActiveCell.FormulaR1C1 = "WTE"
    Range("A67").Select
    ActiveCell.FormulaR1C1 = "Incinerated"
    Range("A68").Select
    ActiveCell.FormulaR1C1 = "Ending Inventory Processed Goods this quarter"
    Range("A69").Select
    ActiveCell.FormulaR1C1 = "A_Type1_Output_Payout_Per_Pound"
    Rows("69:69").Select
    Selection.Insert Shift:=xlDown
    Range("A69").Select
    ActiveCell.FormulaR1C1 = "TOTAL Recycled Pounds This Quarter"
    Rows("70:70").Select
    Selection.Insert Shift:=xlDown
    Range("A70").Select
    ActiveCell.FormulaR1C1 = "Line 69 must equal line 51"
    Rows("71:71").Select
    Selection.Insert Shift:=xlDown
    Range("A71").Select
    ActiveCell.FormulaR1C1 = "Calculations for funding"
    Range("A72").Select
    Range("A72").Select
    ActiveWindow.SmallScroll Down:=9
    Rows("73:73").Select
    Selection.Delete Shift:=xlUp
    Rows("74:74").Select
    Selection.Delete Shift:=xlUp
    Range("A72").Select
    ActiveCell.FormulaR1C1 = "A_Type1_Output_Payout_Per_Pound"
    Rows("75:75").Select
    Selection.Delete Shift:=xlUp
    Rows("76:76").Select
    Selection.Delete Shift:=xlUp
    Range("A72").Select
    ActiveCell.FormulaR1C1 = "type 1 Output, $0.06/lb."
    Range("A72").Select
    ActiveCell.FormulaR1C1 = "Type 1 Output, $0.06/lb."
    Range("A73").Select
    ActiveCell.FormulaR1C1 = "Type 2 Output, $0.03/lb."
    Range("A74").Select
    ActiveCell.FormulaR1C1 = "CAAF, $0.03/lb."
    Range("A75").Select
    Range("A75").Select
    ActiveCell.FormulaR1C1 = "Cement Kiln feedstock, $0.03/lb."
    Range("A76").Select
    ActiveCell.FormulaR1C1 = "Total Requested ($)"
    Range("A77").Select
    ActiveCell.FormulaR1C1 = "Date Filed"
    Range("A77").Select
End Sub
