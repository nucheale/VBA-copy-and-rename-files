Sub renameAndMoveFiles()

Dim objFSO As Object
Dim i As Integer
Dim drive As String

Set objFSO = CreateObject("Scripting.FileSystemObject")
'Set objFolder = objFSO.GetFolder(folderPath)

lastRow = Cells(Rows.Count, 2).End(xlUp).Row
drive = Left(ActiveWorkbook.Path, 3)

'Debug.Print drive

For i = 3 To lastRow
    oldFolderPath0 = Replace(Cells(i, 1), "/", "\")
    oldFolderPath = ActiveWorkbook.Path & "\" & oldFolderPath0
    oldFileName = oldFolderPath & Cells(i, 2)
    
    newFolderPath0 = Replace(Cells(i, 4), "/", "\")
    newFolderPath = ActiveWorkbook.Path & "\" & newFolderPath0
    newFileName = newFolderPath & Cells(i, 5)
    
    If Dir(oldFileName) <> "" Then
        objFSO.CopyFile oldFileName, newFileName
        Else
    End If
Next i

End Sub
