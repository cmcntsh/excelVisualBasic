# excelVisualBasic
Excel Visual Basic information and code snippets

```
Dim lastRow As Long, lastColumn As Long, ws As Worksheet

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
```
