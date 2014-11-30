Attribute VB_Name = "Module4"
Sub ForecastMacro()
Attribute ForecastMacro.VB_Description = "This macro is responsible for formatting the Tier 1 Forecast report."
Attribute ForecastMacro.VB_ProcData.VB_Invoke_Func = "f\n14"
'
' ForecastMacro Macro
' This macro is responsible for formatting the Tier 1 Forecast report.
'
' Keyboard Shortcut: Ctrl+f
'
    Application.Run "PERSONAL.XLSB!MacroTest"
    Sheets("Tier1_Forecast").Select
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Rows("5:5").Select
    Selection.Delete Shift:=xlUp
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Rows("3:3").Select
    Selection.Delete Shift:=xlUp
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp
    Rows("3:3").Select
    Selection.Copy
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown
    Application.CutCopyMode = False
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Rows("4:4").Select
    Selection.Delete Shift:=xlUp
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "CONFIDENTIAL"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "AB 2398 Monthly Rolling Forecast"
    Rows("4:4").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A5").Select
    ActiveCell.FormulaR1C1 = _
        "Number of CA FTE Employees at the beginning of this quarter"
    Range("A6").Select
    Columns("A:A").EntireColumn.AutoFit
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "Number of FTE CA Jobs lost this quarter"
    Range("A7").Select
    ActiveCell.FormulaR1C1 = "Number of FTE CA Jobs gained this quarter"
    Range("A8").Select
    ActiveCell.FormulaR1C1 = "Number of FTE CA Employees at end of this quarter"
    Rows("9:9").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A10").Select
    ActiveCell.FormulaR1C1 = _
        "Post-consumer carpet pounds directly collected by you from California for this quarter"
    Range("A11").Select
    ActiveCell.FormulaR1C1 = _
        "Post-consumer carpet pounds directly collected by you from California for this quarter"
    Range("A11").Select
    ActiveCell.FormulaR1C1 = _
        "Post-consumer carpet pounds directly collected by you from OUTSIDE California for this quarter"
    Rows("12:12").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A12").Select
    ActiveCell.FormulaR1C1 = "TOTAL Post-consumer carpet pounds"
    Rows("13:13").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Rows("14:14").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A14").Select
    ActiveCell.FormulaR1C1 = "Nylon 6"
    Rows("15:15").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A15").Select
    ActiveCell.FormulaR1C1 = "Nylon6,6"
    Range("A15").Select
    ActiveCell.FormulaR1C1 = "Nylon 6,6"
    Rows("16:16").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A16").Select
    ActiveCell.FormulaR1C1 = "Polypropylene"
    Rows("17:17").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A17").Select
    ActiveCell.FormulaR1C1 = "PET"
    Rows("18:18").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A18").Select
    ActiveCell.FormulaR1C1 = "Wool"
    Rows("19:19").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A19").Select
    ActiveCell.FormulaR1C1 = "Other/Mixed Fibers"
    Rows("20:20").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A20").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Rows("21:21").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A21").Select
    ActiveCell.FormulaR1C1 = "Line 20 must equal line 10"
    Rows("22:22").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A23").Select
    ActiveCell.FormulaR1C1 = _
        "Beginning Inventory of Whole Carpet from CA at start of quarter (should equal prior quarter ending inventory)."
    Range("A24").Select
    ActiveCell.FormulaR1C1 = "Whole Carpet Collected from California (Row 10)"
    Range("A25").Select
    ActiveCell.FormulaR1C1 = "Whole Carpet from CA received from other collectors"
    Rows("26:26").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.FormulaR1C1 = "T"
    Range("A26").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Rows("27:27").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A28").Select
    ActiveCell.FormulaR1C1 = "Re-Used"
    Range("A29").Select
    ActiveCell.FormulaR1C1 = "Internally Used Whole Carpet this quarter"
    Range("A30").Select
    ActiveCell.FormulaR1C1 = _
        "Whole carpet shipped to US customers OUTSIDE California"
    Range("A31").Select
    ActiveCell.FormulaR1C1 = _
        "Whole carpet shipped to US customers OUTSIDE the United States"
    Range("A32").Select
    ActiveCell.FormulaR1C1 = "Whole carpet shipped to customers INSIDE California"
    Range("A33").Select
    ActiveCell.FormulaR1C1 = _
        "Non-carpet materials with value (i.e. carpet cushion)"
    Range("A34").Select
    ActiveCell.FormulaR1C1 = "WTE"
    Range("A35").Select
    ActiveCell.FormulaR1C1 = "Incinerated"
    Range("A36").Select
    ActiveCell.FormulaR1C1 = "Landfilled"
    Range("A37").Select
    ActiveCell.FormulaR1C1 = "Ending Inventory of Whole Carpet"
    Rows("38:38").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A38").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Rows("39:39").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A39").Select
    ActiveCell.FormulaR1C1 = "Line 38 must equal line 26"
    Rows("40:40").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
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
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A46").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Rows("47:47").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A47").Select
    ActiveCell.FormulaR1C1 = "Line 46 must equal line 41"
    Rows("48:48").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A49").Select
    ActiveCell.FormulaR1C1 = _
        "Beginning Inventory of Processed Goods from prior quarter"
    Range("A50").Select
    ActiveCell.FormulaR1C1 = "Processed"
    Rows("51:51").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A51").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    Rows("52:52").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A52").Select
    ActiveCell.FormulaR1C1 = "Type 1 Outputs"
    Range("A53").Select
    ActiveCell.FormulaR1C1 = "Fiber"
    Range("A54").Select
    ActiveCell.FormulaR1C1 = "DePoly or Chemical Component"
    Range("A55").Select
    ActiveCell.FormulaR1C1 = "Shredded Carpet tile used for tile backing"
    Range("A56").Select
    ActiveCell.FormulaR1C1 = _
        "Number of Ash tests run this quarter (min 1 per 1M pounds)"
    Range("A57").Select
    ActiveCell.FormulaR1C1 = _
        "Average Ash Test Results over quarter for Type 1 pounds"
    Rows("58:58").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A58").Select
    ActiveCell.FormulaR1C1 = "Total Type 1 Ountput: SOLD & SHIPPED"
    Rows("59:59").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A59").Select
    ActiveCell.FormulaR1C1 = "Type 2 Outputs"
    Rows("60:60").Select
    Selection.Delete Shift:=xlUp
    Rows("60:60").Select
    Selection.Delete Shift:=xlUp
    Range("A60").Select
    ActiveCell.FormulaR1C1 = "Filler"
    Rows("61:61").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A61").Select
    ActiveCell.FormulaR1C1 = "Total Type 2 Output: SOLD & SHIPPED"
    Range("A62").Select
    ActiveCell.FormulaR1C1 = "CAAF"
    Range("A63").Select
    ActiveCell.FormulaR1C1 = "Cement Kiln feedstock"
    Range("A64").Select
    ActiveCell.FormulaR1C1 = "Carcass Sold"
    Range("A65").Select
    ActiveCell.FormulaR1C1 = "Landfilled"
    Range("A66").Select
    ActiveCell.FormulaR1C1 = "WTE"
    Range("A67").Select
    ActiveCell.FormulaR1C1 = "Incinerated"
    Range("A68").Select
    ActiveCell.FormulaR1C1 = "Ending Inventory Processed Goods this quarter"
    Rows("69:69").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A69").Select
    ActiveCell.FormulaR1C1 = "TOTAL Recycled Pounds This Quarter"
    Rows("70:70").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A70").Select
    ActiveCell.FormulaR1C1 = "Line 69 must equal line 51"
    Rows("71:71").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A72").Select
    ActiveCell.FormulaR1C1 = "Total_Payout_Adjustments"
    Range("A73").Select
    Rows("72:72").Select
    Selection.Delete Shift:=xlUp
    Rows("73:73").Select
    Selection.Delete Shift:=xlUp
    Rows("74:74").Select
    Selection.Delete Shift:=xlUp
    Rows("75:75").Select
    Selection.Delete Shift:=xlUp
    Range("A72").Select
    ActiveCell.FormulaR1C1 = "Type 1 Output, $0.06/lb."
    Range("A73").Select
    ActiveCell.FormulaR1C1 = "Type 2 Output, $0.03/lb."
    Range("A74").Select
    ActiveCell.FormulaR1C1 = "CAAF, $0.03/lb."
    Range("A75").Select
    ActiveCell.FormulaR1C1 = "Cement Kiln feedstock, $0.03/lb"
    Range("A76").Select
    ActiveCell.FormulaR1C1 = "Total Requested ($)"
    Range("A77").Select
End Sub
