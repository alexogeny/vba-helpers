Public Function filterData(sheetName, colIndex, searchCriteria, Optional inverse As Boolean = False, Optional delete as Boolean = False)
    Application.ScreenUpdating = False
    If inverse = True Then
        searchCriteria = "<>" & searchCriteria
    End If
    With ActiveWorkbook.Sheets(sheetName)
    .AutoFilterMode = False
    .UsedRange.AutoFilter Field:=colIndex, Criteria1:=searchCriteria
    If delete = True Then
        .AutoFilter.Range.Offset(1, 0).EntireRow.Delete
    End If
    .AutoFilterMode = False
    End With
    Application.ScreenUpdating = True
End Function
