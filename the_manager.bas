Attribute VB_Name = "Module1"
Option Explicit

Sub Make_a_new_file()

Application.ScreenUpdating = False

Dim endpoint As Range
Dim checkpoint As Range
Dim xColIndex As Long
Dim xRowIndex As Long
Dim xIndex As Long
Dim NumRows As Long
Dim CumColumns As Long
Dim NumColumns As Long
Dim levelnumber As Long

If Worksheets("NewCurrent").Range("A1") = "" Then
    MsgBox "No data!"
    Application.ScreenUpdating = True
    Exit Sub
End If


Application.ScreenUpdating = True

End Sub
