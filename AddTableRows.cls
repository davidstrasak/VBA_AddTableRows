Private Sub Worksheet_SelectionChange(ByVal Target As Range)
''THIS ALWAYS ADDS A ROW AT THE END OF THE TABLE

''The file here is looking for the data inserted in the bottom left cell.
If Not IsEmpty(Range("Table1").Resize(1, 1).Offset(Range("Table1").Rows.Count - 1).Value) Then
    Application.ScreenUpdating = False 'this makes all of this faster
    Application.Calculation = xlManual
    Call AddRow
    Sheet1.Calculate
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
End If
End Sub
Private Sub AddRow()
'' Macro that adds lines to the projectplan table
    Dim CountedRows As Integer, usedTable As String
    usedTable = "Table1"
          
    CountedRows = Range(usedTable).Rows.Count

    If Not IsEmpty(Range(usedTable).Resize(1, 1).Offset(CountedRows - 1).Value) Then
    Range(usedTable).Resize(1).Offset(CountedRows).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range(usedTable).Resize(1).Offset(CountedRows).Value = Range(usedTable).Resize(1).Value
    Application.CutCopyMode = False
    
    ''Copies or deletes cells - change it based on the table
    Range(usedTable).Resize(1, 2).Offset(CountedRows).ClearContents
    Range(usedTable).Resize(1, 1).Offset(CountedRows, 2).FillDown
    End If
End Sub
