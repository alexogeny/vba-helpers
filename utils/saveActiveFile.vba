Public Sub saveActiveFile()
    Dim bFileSaveAs As Boolean
    bFileSaveAs = Application.Dialogs(xlDialogSaveAs).Show
    If Not bFileSaveAs Then MsgBox "User cancelled", vbCritical
End Sub
