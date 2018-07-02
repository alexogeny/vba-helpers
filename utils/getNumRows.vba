Public Function getNumRows(sheetName, Optional colIndex As String = "A")
    With ActiveWorkbook.Sheets(sheetName)
    getNumRows = .Cells(.Rows.Count, colIndex).End(xlUp).Row
    End With
End Function
