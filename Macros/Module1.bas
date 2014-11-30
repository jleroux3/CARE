Attribute VB_Name = "Module1"
Sub TransposeData()
Attribute TransposeData.VB_ProcData.VB_Invoke_Func = "t\n14"
'
' TransposeData Macro
'
    'Range("A1:CY23").Select
    'Cells.Select
    ActiveSheet.UsedRange.Copy
    'Selection.Copy
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Tier1_Actual"
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Range("A1").Select
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "Tier1_Forecast"
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Range("A1").Select
    ActiveWindow.SmallScroll Down:=-15
    
    'Sub Loop_Example()
    Dim Firstrow As Long
    Dim Lastrow As Long
    Dim Lrow As Long
    Dim CalcMode As Long
    Dim ViewMode As Long

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
    
    'Delete procedure:
    
    'We use the ActiveSheet but you can replace this with
    'Sheets("MySheet")if you want
    'With ActiveSheet
    With Sheets("Tier1_Actual")

        'We select the sheet so we can change the window view
        .Select

        'If you are in Page Break Preview Or Page Layout view go
        'back to normal view, we do this for speed
        ViewMode = ActiveWindow.View
        ActiveWindow.View = xlNormalView

        'Turn off Page Breaks, we do this for speed
        .DisplayPageBreaks = False

        'Set the first and last row to loop through
        Firstrow = .UsedRange.Cells(1).Row
        Lastrow = .UsedRange.Rows(.UsedRange.Rows.Count).Row

        'We loop from Lastrow to Firstrow (bottom to top)
        For Lrow = Lastrow To Firstrow Step -1

            'We check the values in the A column in this example
            With .Cells(Lrow, "A")

                If Not IsError(.Value) Then

                    'If .Value = "ron" Then .EntireRow.Delete
                    If .Value Like "*Forecast*" Then .EntireRow.Delete
                    'This will delete each row with the Value "ron"
                    'in Column A, case sensitive.

                End If

            End With

        Next Lrow

    End With

    ActiveWindow.View = ViewMode
    With Application
        .ScreenUpdating = True
        .Calculation = CalcMode
    End With

'End Sub

    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With
    
    'Delete procedure:
    
    'We use the ActiveSheet but you can replace this with
    'Sheets("MySheet")if you want
    'With ActiveSheet
    With Sheets("Tier1_Forecast")

        'We select the sheet so we can change the window view
        .Select

        'If you are in Page Break Preview Or Page Layout view go
        'back to normal view, we do this for speed
        ViewMode = ActiveWindow.View
        ActiveWindow.View = xlNormalView

        'Turn off Page Breaks, we do this for speed
        .DisplayPageBreaks = False

        'Set the first and last row to loop through
        Firstrow = .UsedRange.Cells(1).Row
        Lastrow = .UsedRange.Rows(.UsedRange.Rows.Count).Row

        'We loop from Lastrow to Firstrow (bottom to top)
        For Lrow = Lastrow To Firstrow Step -1

            'We check the values in the A column in this example
            With .Cells(Lrow, "A")

                If Not IsError(.Value) Then

                    'If .Value = "ron" Then .EntireRow.Delete
                    If .Value Like "*Actual*" Then .EntireRow.Delete
                    'This will delete each row with the Value "ron"
                    'in Column A, case sensitive.

                End If

            End With

        Next Lrow

    End With

    ActiveWindow.View = ViewMode
    With Application
        .ScreenUpdating = True
        .Calculation = CalcMode
    End With

    
End Sub

