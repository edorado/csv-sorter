Attribute VB_Name = "Module4"
Sub BankofAmerica()
Attribute BankofAmerica.VB_ProcData.VB_Invoke_Func = "b\n14"
'
' BankofAmerica Macro
'
' Keyboard Shortcut: Option+Cmd+b

Dim x As Integer

    Rows("1:6").Select
    Selection.Delete Shift:=xlUp
    Columns("D:D").Select
    Selection.Delete Shift:=xlToLeft
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Payee"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Description"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "Bank of America"
    Range("E2:E" & Range("a65000").End(xlUp).Row).FillDown
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Category"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-1]<0,""Uncategorized"",""Other income"")"
    Range("D2").Select
    Range("D2:D" & Range("a65000").End(xlUp).Row).FillDown
    Range("D1").Select
    ActiveWorkbook.Worksheets("BankofAmerica").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BankofAmerica").Sort.SortFields.Add Key:=Range("D1") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BankofAmerica").Sort
        .SetRange Range("A:E")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E2").Select
    x = Application.WorksheetFunction.CountIf(Range("D:D"), "Uncategorized")
    Rows("2:" & x + 1).Select
    Selection.Cut
    Worksheets.Add(After:=Worksheets(1)).Name = "expense"
    Worksheets.Add(After:=Worksheets(1)).Name = "income"
    Sheets("expense").Select
    Range("A2").Select
    ActiveSheet.Paste
    Sheets("BankofAmerica").Select
    Selection.Delete Shift:=xlUp
    Range("A1:E1").Select
    Selection.Copy
    Sheets("expense").Select
    Range("A1").Select
    ActiveSheet.Paste
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Sheets("BankofAmerica").Select
    Cells.Select
    Selection.Cut
    Sheets("income").Select
    ActiveSheet.Paste
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Sheets("expense").Select
    Range("F2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-3]*-1"
    Range("F2").Select
    Range("F2:F" & Range("a65000").End(xlUp).Row).FillDown
    Selection.Copy
    Range("F2:F" & Range("a65000").End(xlUp).Row).Select
    Selection.Copy
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Columns("F:F").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Cells.Select
    Sheets("BankofAmerica").Select
    Range("A1").Select
    Sheets("expense").Select
    Sheets("expense").Move
    ActiveWorkbook.SaveAs Filename:="bofaexpense.csv", FileFormat:=xlCSV, _
        CreateBackup:=False
    ActiveWindow.Close
    Sheets("income").Select
    Sheets("income").Move
    ActiveWorkbook.SaveAs Filename:="bofaincome.csv", FileFormat:=xlCSV, _
        CreateBackup:=False
    ActiveWindow.Close
    Sheets("BankofAmerica").Select
End Sub
Sub WellsFargo()
Attribute WellsFargo.VB_ProcData.VB_Invoke_Func = "s\n14"
'
' WellsFargo Macro
'
' Keyboard Shortcut: Option+Cmd+s
Dim x As Integer

    Sheets("WellsFargo").Select
    Range("A1").Select
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Date"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Amount"
    Columns("C:C").Select
    ActiveWindow.ScrollRow = 1
    Selection.Delete Shift:=xlToLeft
    Columns("C:C").Select
    Selection.Delete Shift:=xlToLeft
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Payee"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Description"
    Range("D2").Select
    ActiveCell.FormulaR1C1 = "Wells Fargo"
    Range("D2").Select
    Range("D2:D" & Range("a65000").End(xlUp).Row).FillDown
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Category"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-3]<0,""Uncategorized"",""Other income"")"
    Range("E2").Select
    Selection.AutoFill Destination:=Range("E2:E295")
    Range("E2:E295").Select
    Range("E1").Select
    ActiveWorkbook.Worksheets("WellsFargo").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("WellsFargo").Sort.SortFields.Add Key:=Range("E1") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("WellsFargo").Sort
        .SetRange Range("A:E")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E2").Select
    x = Application.WorksheetFunction.CountIf(Range("E:E"), "Uncategorized")
    Rows("2:" & x + 1).Select
    Selection.Cut
    Worksheets.Add(After:=Worksheets(2)).Name = "expense"
    Worksheets.Add(After:=Worksheets(2)).Name = "income"
    Sheets("expense").Select
    Range("A2").Select
    ActiveSheet.Paste
    Sheets("WellsFargo").Select
    Selection.Delete Shift:=xlUp
    Range("A1:E1").Select
    Selection.Copy
    Sheets("expense").Select
    Range("A1").Select
    ActiveSheet.Paste
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Sheets("WellsFargo").Select
    Cells.Select
    Selection.Cut
    Sheets("income").Select
    ActiveSheet.Paste
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Sheets("expense").Select
    Range("F2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-4]*-1"
    Range("F2").Select
    Range("F2:F" & Range("a65000").End(xlUp).Row).FillDown
    Selection.Copy
    Range("F2:F" & Range("a65000").End(xlUp).Row).Select
    Selection.Copy
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Columns("F:F").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Cells.Select
    Sheets("WellsFargo").Select
    Range("A1").Select
    Sheets("expense").Select
    Sheets("expense").Move
    ActiveWorkbook.SaveAs Filename:="wellsexpense.csv", FileFormat:=xlCSV, _
        CreateBackup:=False
    ActiveWindow.Close
    Sheets("income").Select
    Sheets("income").Move
    ActiveWorkbook.SaveAs Filename:="wellsincome.csv", FileFormat:=xlCSV, _
        CreateBackup:=False
    ActiveWindow.Close
    Sheets("WellsFargo").Select
    
    
End Sub
Sub Chase()
Attribute Chase.VB_ProcData.VB_Invoke_Func = "v\n14"
'
' Chase Macro
'
' Keyboard Shortcut: Option+Cmd+v
'
Dim x As Integer

    Sheets("Chase").Select
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "Date"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Payee"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Description"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "Chase"
    Range("E2").Select
    Range("E2:E" & Range("a65000").End(xlUp).Row).FillDown
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "Category"
    Range("E2").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-2]<0,""Uncategorized"",""Other income"")"
    Range("E2").Select
    Range("E2:E" & Range("a65000").End(xlUp).Row).FillDown
    Range("E1").Select
    ActiveWorkbook.Worksheets("Chase").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Chase").Sort.SortFields.Add Key:=Range("E1"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Chase").Sort
        .SetRange Range("A:E")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E2").Select
    x = Application.WorksheetFunction.CountIf(Range("E:E"), "Uncategorized")
    Rows("2:" & x + 1).Select
    Selection.Cut
    Worksheets.Add(After:=Worksheets(3)).Name = "expense"
    Worksheets.Add(After:=Worksheets(3)).Name = "income"
    Sheets("expense").Select
    Range("A2").Select
    ActiveSheet.Paste
    Sheets("Chase").Select
    Selection.Delete Shift:=xlUp
    Range("A1:E1").Select
    Selection.Copy
    Sheets("expense").Select
    Range("A1").Select
    ActiveSheet.Paste
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Sheets("Chase").Select
    Cells.Select
    Selection.Cut
    Sheets("income").Select
    ActiveSheet.Paste
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Sheets("expense").Select
    Range("F2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-3]*-1"
    Range("F2").Select
    Range("F2:F" & Range("a65000").End(xlUp).Row).FillDown
    Selection.Copy
    Range("F2:F" & Range("a65000").End(xlUp).Row).Select
    Selection.Copy
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Columns("F:F").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    Cells.Select
    Sheets("Chase").Select
    Range("A1").Select
    Sheets("expense").Select
    Sheets("expense").Move
    ActiveWorkbook.SaveAs Filename:="chaseexpense.csv", FileFormat:=xlCSV, _
        CreateBackup:=False
    ActiveWindow.Close
    Sheets("income").Select
    Sheets("income").Move
    ActiveWorkbook.SaveAs Filename:="chaseincome.csv", FileFormat:=xlCSV, _
        CreateBackup:=False
    ActiveWindow.Close
    Sheets("Chase").Select
    
End Sub
