Attribute VB_Name = "Module1"
Option Explicit

Sub transform_data()

Application.ScreenUpdating = True

Dim endpoint As Range
Dim checkpoint As Range
Dim xColIndex As Long
Dim xRowIndex As Long
Dim levelnumber As Long

Dim xIndex As Long
Dim NumRows As Long
Dim CumColumns As Long
Dim NumColumns As Long

Dim ColNumb As Long
Dim iCtr As Long
Dim counter As Long
Dim first_level As Range
Dim business_rule As Range

Dim current As Worksheet
Dim NewData As Worksheet
Dim dictionary As Worksheet
Dim CurrentCopy As Worksheet

Dim Message As String

On Error Resume Next
Set current = ThisWorkbook.Worksheets("NewCurrent")
Set NewData = ThisWorkbook.Worksheets("NewData")
Set dictionary = ThisWorkbook.Worksheets("Dictionary")

'Dictionary columns required
Dim BusinessRuleName As String
Dim FirstLevel As String
Dim ID As String
Dim Code As String
Dim Name As String
Dim Level As String
Dim LevelNameMatch As String
Dim LevelID As String
Dim LevelName As String
Dim NumberOfLevels As String
Dim LevelsSkipped As String

Dim IDLetter As String
Dim CodeLetter  As String
Dim NameLetter  As String
Dim LevelLetter  As String
Dim LevelNameMatchLetter  As String
Dim LevelIDLetter  As String
Dim LevelNameLetter  As String
Dim NumberOfLevelsLetter  As String
Dim LevelsSkippedLetter  As String

BusinessRuleName = "Business Rule Name"
FirstLevel = "First Level"
ID = "ID"
Code = "Code"
Name = "Name"
Level = "Level"
LevelNameMatch = "Level Name Match"
LevelID = "Level ID"
LevelName = "Level Name"
NumberOfLevels = "Number of Levels"
LevelsSkipped = "Levels Skipped"


'Check if NewCurrent sheet contains data
If current.Range("A1") = "" Then
    MsgBox "No data!"
    Application.ScreenUpdating = True
    Exit Sub
End If

'Delete NewData sheet if exists
On Error Resume Next
If Err.Number = 0 Then
    Application.DisplayAlerts = False
    NewData.Delete
End If

'Delete CurrentCopy sheet if exists
On Error Resume Next
If Err.Number = 0 Then
    Application.DisplayAlerts = False
    CurrentCopy.Delete
End If

'create a copy of NewCurrent sheet
current.Activate
current.Copy After:=current
ActiveSheet.Name = "NewCurrentCopy"

Call FindCell(dictionary, FirstLevel, 40)
Set first_level = ActiveCell.Offset(1, 0)

Call FindCell(dictionary, NumberOfLevels, 40)
Set endpoint = ActiveCell.Offset(1, 0)

Call FindCell(dictionary, BusinessRuleName, 40)
Set business_rule = ActiveCell.Offset(1, 0)

IDLetter = FindCellLetter(dictionary, ID, 40)
CodeLetter = FindCellLetter(dictionary, Code, 40)
NameLetter = FindCellLetter(dictionary, Name, 40)
LevelLetter = FindCellLetter(dictionary, Level, 40)
LevelNameMatchLetter = FindCellLetter(dictionary, LevelNameMatch, 40)
LevelIDLetter = FindCellLetter(dictionary, LevelID, 40)
LevelNameLetter = FindCellLetter(dictionary, LevelName, 40)
LevelsSkippedLetter = FindCellLetter(dictionary, LevelsSkipped, 40)

Worksheets("NewCurrentCopy").Activate

ColNumb = ActiveSheet.UsedRange.Columns.Count

For iCtr = ColNumb To 1 Step -1
    If Not IsError(Application.Match(Cells(1, iCtr), dictionary.Range(LevelsSkippedLetter & ":" & LevelsSkippedLetter), 0)) Then
        Columns(iCtr + 1).Delete
        Columns(iCtr).Delete
    End If
Next

Sheets.Add(After:=Sheets(Sheets.Count)).Name = "NewData"

NewData.Activate
Range("A1").FormulaR1C1 = "a"
Range("B1").FormulaR1C1 = "b"
Range("C1").FormulaR1C1 = "c"
Range("D1").FormulaR1C1 = "d"
Range("E1").FormulaR1C1 = "e"
Range("F1").FormulaR1C1 = "f"
Range("G1").FormulaR1C1 = "Level ID"
Range("H1").FormulaR1C1 = "Level Name"
Range("I1").FormulaR1C1 = "ID"
Range("J1").FormulaR1C1 = "Code"
Range("K1").FormulaR1C1 = "Name"
Range("L1").FormulaR1C1 = "Manager Number"
Range("M1").FormulaR1C1 = "Manager Name"
Range("N1").FormulaR1C1 = "Email"
Range("O1").FormulaR1C1 = "Order"
Range("P1").FormulaR1C1 = "combine"
Range("Q1").FormulaR1C1 = "duplicates?"
Range("R1").FormulaR1C1 = "number length?"

Worksheets("NewCurrentCopy").Select
Range("A1").Select

If business_rule Then
    Do Until ActiveCell.Value = business_rule
        ActiveCell.Offset(, 1).Activate
        counter = counter + 1
        If counter = 40 Then
            Range("A1").Activate
            Exit Sub
        End If
    Loop
    ActiveCell.EntireColumn.Delete
End If


Do Until ActiveCell.Value = first_level
    ActiveCell.Offset(, 1).Activate
    counter = counter + 1
    If counter = 40 Then
        Range("A1").Activate
        Exit Sub
    End If
Loop

Dim first_level_cell_one As String
Dim first_level_cell_two As String
Dim column_two As String
Dim column_three As String
Dim column_three_cell_two As String
Dim column_four As String
Dim column_four_cell_two As String
Dim column_minus_two As String
Dim first_col As String
Dim second_col As String
Dim third_col As String
Dim fourth_col As String
Dim second_minus_col As String

Dim group_counter As Long
Dim group_id As Variant
Dim new_group_id As Variant

first_level_cell_one = ActiveCell.AddressLocal(False, False)
first_level_cell_two = ActiveCell.Offset(1, 0).AddressLocal(False, False)
column_two = ActiveCell.Offset(, 1).AddressLocal(False, False)
column_three = ActiveCell.Offset(, 2).AddressLocal(False, False)
column_three_cell_two = ActiveCell.Offset(1, 2).AddressLocal(False, False)
column_four = ActiveCell.Offset(, 3).AddressLocal(False, False)
column_four_cell_two = ActiveCell.Offset(1, 3).AddressLocal(False, False)
column_minus_two = ActiveCell.Offset(, -2).AddressLocal(False, False)
first_col = Left(first_level_cell_one, Len(first_level_cell_one) - 1)
second_col = Left(column_two, Len(column_two) - 1)
third_col = Left(column_three, Len(column_three) - 1)
fourth_col = Left(column_four, Len(column_four) - 1)
second_minus_col = Left(column_minus_two, Len(column_minus_two) - 1)

levelnumber = 0
'needed for sorting approvers
group_counter = 1

Do
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
    End If

    If levelnumber = endpoint Then Exit Do

    Set checkpoint = Range(first_level_cell_two, Range(first_level_cell_two).End(xlDown))

    'If level has no approvers, skip it
    Do Until WorksheetFunction.CountA(checkpoint) > 0
        group_id = Application.WorksheetFunction.Index(Sheets("Dictionary").Range(LevelLetter & ":" & LevelLetter), _
            Application.WorksheetFunction.Match(Range(first_level_cell_one), _
            dictionary.Range(LevelNameMatchLetter & ":" & LevelNameMatchLetter), 0))
            
            
        Columns(first_col & ":" & second_col).Select ''Activate
        Selection.Delete Shift:=xlToLeft
        Set checkpoint = Range(first_level_cell_two, Range(first_level_cell_two).End(xlDown))

        levelnumber = levelnumber + 1
        new_group_id = Application.WorksheetFunction.Index(Sheets("Dictionary").Range(LevelLetter & ":" & LevelLetter), _
            Application.WorksheetFunction.Match(Range(first_level_cell_one), _
            dictionary.Range(LevelNameMatchLetter & ":" & LevelNameMatchLetter), 0))
            

        If group_id <> new_group_id Then
            group_counter = 1
        End If

        If levelnumber = endpoint Then Exit Do
    Loop

    If levelnumber = endpoint Then Exit Do

    group_id = Application.WorksheetFunction.Index(Sheets("Dictionary").Range(LevelLetter & ":" & LevelLetter), _
        Application.WorksheetFunction.Match(Range(first_level_cell_one), _
        dictionary.Range(LevelNameMatchLetter & ":" & LevelNameMatchLetter), 0))
        

    Columns(third_col & ":" & third_col).Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range(first_level_cell_one).Copy
    Range(column_three_cell_two).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Selection.AutoFill Destination:=Range(column_three_cell_two & ":" & third_col & Range(second_minus_col & Rows.Count).End(xlUp).Row), Type:=xlFillCopy

    Range(column_four_cell_two).Select
    ActiveCell.Value = group_counter
    Selection.AutoFill Destination:=Range(column_four_cell_two & ":" & fourth_col & Range(second_minus_col & Rows.Count).End(xlUp).Row), Type:=xlFillCopy

    ActiveSheet.Range(first_level_cell_one).AutoFilter Field:=Range(first_level_cell_one).Column, Criteria1:="<>"

    Range(column_minus_two).Select
    xIndex = Application.ActiveCell.Column
    xRowIndex = Application.ActiveSheet.Cells(Rows.Count, xIndex).End(xlUp).Row
    Range(Cells(2, xIndex), Cells(xRowIndex, xIndex)).Select
    Selection.Resize(, 6).Select

    Selection.Copy
    ActiveSheet.Paste Destination:=Worksheets("NewData").Range("A" & Rows.Count).End(xlUp).Offset(1, 0)
    Columns("G:J").Select
    Selection.Delete Shift:=xlToLeft

    levelnumber = levelnumber + 1

    new_group_id = Application.WorksheetFunction.Index(Sheets("Dictionary").Range(LevelLetter & ":" & LevelLetter), _
        Application.WorksheetFunction.Match(Range(first_level_cell_one), _
        dictionary.Range(LevelNameMatchLetter & ":" & LevelNameMatchLetter), 0))


    If group_id <> new_group_id Then
        group_counter = 1
    Else
        group_counter = group_counter + 1
    End If
Loop

Application.DisplayAlerts = False
Worksheets("NewCurrentCopy").Delete '?

Worksheets("NewData").Select

Range("G2").Formula = "=INDEX(Dictionary!" & LevelLetter & ":" & LevelLetter & ", MATCH(E2, Dictionary!" & LevelNameMatchLetter & ":" & LevelNameMatchLetter & ", 0))"
Range("G2").AutoFill Destination:=Range("G2:G" & Range("A" & Rows.Count).End(xlUp).Row)

Range("H2").Formula = "=VLOOKUP(G2,Dictionary!" & LevelIDLetter & ":" & LevelNameLetter & ", 2,0)"
Range("H2").AutoFill Destination:=Range("H2:H" & Range("A" & Rows.Count).End(xlUp).Row)

Range("I2").Formula = "=INDEX(Dictionary!" & IDLetter & ":" & IDLetter & ", MATCH(A2, Dictionary!" & CodeLetter & ":" & CodeLetter & ", 0))"
Range("I2").AutoFill Destination:=Range("I2:I" & Range("A" & Rows.Count).End(xlUp).Row)

Range("J2").Formula = "=A2"
Range("J2").AutoFill Destination:=Range("J2:J" & Range("A" & Rows.Count).End(xlUp).Row)

Range("K2").Formula = "=VLOOKUP(J2,Dictionary!" & CodeLetter & ":" & NameLetter & ", 2,0)"
Range("K2").AutoFill Destination:=Range("K2:K" & Range("A" & Rows.Count).End(xlUp).Row)

Range("L2").FormulaR1C1 = _
        "=MID(RC[-9],SEARCH(""("",RC[-9])+1,SEARCH("")"",RC[-9])-SEARCH(""("",RC[-9])-1)"
Range("L2").AutoFill Destination:=Range("L2:L" & Range("A" & Rows.Count).End(xlUp).Row)

Range("M2").FormulaR1C1 = "=LEFT(RC[-10],LEN(RC[-10])-11)"
Range("M2").AutoFill Destination:=Range("M2:M" & Range("A" & Rows.Count).End(xlUp).Row)

Range("N2").FormulaR1C1 = "=IF(RC[-10]=""YES"",TRUE,FALSE)"
Range("N2").AutoFill Destination:=Range("N2:N" & Range("A" & Rows.Count).End(xlUp).Row)

Range("O2").Formula = "=F2"
Range("O2").AutoFill Destination:=Range("O2:O" & Range("A" & Rows.Count).End(xlUp).Row)

Columns("G:O").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Application.CutCopyMode = False

Columns("A:F").Select
Selection.Delete Shift:=xlToLeft

Cells.EntireColumn.AutoFit

Dim MyString As String
Dim MyMessage As String
Dim result As String

'quality check
Range("J2").Activate
ActiveCell.FormulaR1C1 = "=RC[-9]&RC[-7]&RC[-4]"
Range("J2").AutoFill Destination:=Range("J2:J" & Range("A" & Rows.Count).End(xlUp).Row)

ActiveSheet.Range("J1").AutoFilter Field:=Range("J1").Column, Criteria1:="#N/A"
MyMessage = "#NAs found! Please correct NewCurrent or update Dictionary and try again."

result = FindMyString("#N/A", "J1", "J:J", MyMessage)
If result = "" Then
    Exit Sub
End If


Range("K2").Activate
ActiveCell.FormulaR1C1 = "=IF(COUNTIF(C[-1],RC[-1])>1, ""Duplicate"",""ok"")"
Range("K2").AutoFill Destination:=Range("K2:K" & Range("A" & Rows.Count).End(xlUp).Row)

ActiveSheet.Range("K1").AutoFilter Field:=Range("K1").Column, Criteria1:="Duplicate"

MyMessage = "Duplicates found! Please correct NewCurrent and try again."

result = FindMyString("Duplicate", "K1", "K:K", MyMessage)
If result = "" Then
    Exit Sub
End If


Range("L2").Activate
ActiveCell.FormulaR1C1 = "=IF(LEN(RC[-6])=8,""ok"",""Incorrect Number"")"
Range("L2").AutoFill Destination:=Range("L2:L" & Range("A" & Rows.Count).End(xlUp).Row)

ActiveSheet.Range("L1").AutoFilter Field:=Range("L1").Column, Criteria1:="Incorrect Number"

MyMessage = "Incorrect Numbers found! Please correct NewCurrent and try again."

result = FindMyString("Incorrect Number", "L2", "L:L", MyMessage)
If result = "" Then
    Exit Sub
End If

Columns("J:L").Select
Selection.Delete Shift:=xlToLeft

Range("A1").Select

MsgBox "All set!"

Application.ScreenUpdating = True

End Sub

Function FindCell(SheetName As Worksheet, CellName As String, CounterLimit As Long) As Range

    Dim counter As Long

    SheetName.Activate
    Range("A1").Activate
    counter = 0
    Do Until ActiveCell.Value = CellName
        ActiveCell.Offset(, 1).Activate
        counter = counter + 1
        If counter = CounterLimit Then
            Range("A1").Activate
            MsgBox "Cell: " & CellName & "not found!"
            Exit Function
        End If
    Loop

End Function

Function FindCellLetter(SheetName As Worksheet, CellName As String, CounterLimit As Long) As String

    Dim counter As Long

    Dim FuncRange As String
    Dim FuncColLength As Long

    SheetName.Activate
    Range("A1").Activate
    counter = 0
    Do Until ActiveCell.Value = CellName
        ActiveCell.Offset(, 1).Activate
        counter = counter + 1
        If counter = CounterLimit Then
            Range("A1").Activate
            MsgBox "Cell: " & CellName & "not found!"
            Exit Function
        End If
    Loop
    
    FuncRange = ActiveCell.AddressLocal(False, False)
    FuncColLength = Len(FuncRange)
    FindCellLetter = Left(FuncRange, FuncColLength - 1)
    

End Function

Function FindMyString(MyString As String, MyCell As String, MyColumn As String, MyMessage As String) As String

    Dim Rng As Range

    If Trim(MyString) <> "" Then
        With Sheets("NewData").Range(MyColumn)
            Set Rng = .Find(What:=MyString, _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext = xlNext, _
                            MatchCase:=False)
            If Not Rng Is Nothing Then
                Range(MyCell).Select
                FindMyString = ""
                MsgBox MyMessage
                Exit Function
            End If
        End With
    End If
    
    FindMyString = "All is ok"
    
    If ActiveSheet.AutoFilterMode Then
        ActiveSheet.AutoFilterMode = False
    End If

End Function
