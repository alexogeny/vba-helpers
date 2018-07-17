Public Function sheetExists(ByVal sheetToFind As String) As Boolean
    sheetExists = False
    For Each sheet In Worksheets
        If sheetToFind = sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next sheet
End Function
