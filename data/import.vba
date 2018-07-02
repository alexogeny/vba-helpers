Public Function importData(fileName, sheetName, Optional customSelect As String = "*")
    Dim Conn As New ADODB.Connection
    Dim mRs As New ADODB.Recordset
    Dim conStr As String, sqlStr As String
    conStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Environ("USERPROFILE") & "\ExcelDataFiles\" & fileName & ".xlsx;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
    sqlStr = "SELECT " & customSelect & " FROM [" & fileName & "$]"

    Conn.ConnectionString = conStr
    Conn.Open
    mRs.Open sqlStr, Conn, adOpenDynamic, adLockOptimistic
    
    DeleteCells(sheetName)
    
    Dim c As Long
    For c = 1 To mRs.Fields.Count
        With ActiveWorkbook.Sheets(sheetName).Cells(1, c)
            .Value = mRs.Fields(c - 1).Name
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next c
    
    If Not mRs.EOF Then
        ActiveWorkbook.Sheets(sheetName).Range("A2").CopyFromRecordset mRs
    End If
    
End Function
