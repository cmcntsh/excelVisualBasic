# excelVisualBasic
Excel Visual Basic information and code snippets

```
Dim lastRow As Long, lastColumn As Long, ws As Worksheet
Set ws = ActiveSheet

' Get the number of the last row that has data
lastRow = ws.Cells.Find(What:="*", _
        After:=ws.Cells(1), _
        Lookat:=xlPart, _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False).Row
        
' Get the number of the last column that has data
lastColumn = ws.Cells.Find(What:="*", _
        After:=ws.Cells(1), _
        Lookat:=xlPart, _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False).Column
        
Dim activeRow As Long, activeColumn As Long
' Get the row and column numbers of the active cell
activeRow = ActiveCell.Row
activeColumn = ActiveCell.Column


' insert a row below the active cell
ActiveCell.Offset(1).EntireRow.Insert Shift:=xlShiftDown

```

Macro to loop through rows in a column and duplicate rows with data
```
Sub Macro1()
'
' Macro1 Macro
'

Dim lastRow As Long, lastColumn As Long, ws As Worksheet
Dim activeRow As Long, activeColumn As Long
Set ws = ActiveSheet

' Get the number of the last row that has data
lastRow = ws.Cells.Find(What:="*", _
        After:=ws.Cells(1), _
        Lookat:=xlPart, _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False).Row
        
' Get the number of the last column that has data
lastColumn = ws.Cells.Find(What:="*", _
        After:=ws.Cells(1), _
        Lookat:=xlPart, _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False).Column

' get the row and column numbers for the active cell
activeRow = ActiveCell.Row
activeColumn = ActiveCell.Column

' loop through each row until we hit the last one with data
While activeRow < lastRow
  If ActiveCell.Value = "" Then
    ' if no data in the cell, move to the next row
    ActiveCell.Offset(1).Activate
  Else
    ' if data is in the cell, duplicate the row
    ActiveCell.Rows("1:1").EntireRow.Copy
    ActiveCell.Offset(1).Rows("1:1").EntireRow.Insert
    ' add one to the last row since we inserted one
    lastRow = lastRow + 1
    ' copy the data block to the main column and clear the contents for current and next row
    ActiveCell.Range("A1:C1").Cut ActiveCell.Offset(1, -3)
    ActiveCell.Offset(1).Range("A1:C1").ClearContents
    ' move to the next row
    ActiveCell.Offset(1).Activate
  End If
  ' update the value for the active row
  activeRow = ActiveCell.Row
Wend

'ActiveCell.Offset(1).EntireRow.Insert Shift:=xlShiftDown

'ws.Cells(startRow + 1, startColumn + 1).Select
        
'MsgBox "Last row number is " & startRow & " " & startColumn

End Sub
```
