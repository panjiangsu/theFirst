'*************************************************************************
'*更新需更新文件夹列表中各文件夹中的文件
'*************************************************************************
Sub UpdateFiles()
    Count = sdicUpdateFoldersList.Count
    For i = 1 To Count
        Set EveryFolder = FSO.GetFolder(sdicUpdateFoldersList.Item(i))
        For Each EveryFile In EveryFolder.Files
            DesFilePath = ReplacePath(EveryFile.Path, sstrOriPos, sstrDesPos)
            If FileExist(DesFilePath) Then
                Set DesFile = FSO.GetFile(DesFilePath)
                Ori = EveryFile.DateLastModified
                Des = DesFile.DateLastModified
                If Ori > Des Then PutInToFilesList EveryFile.Path, sdicCopyFilesList
            Else
                PutInToFilesList EveryFile.Path, sdicCopyFilesList
            End If
        Next
    
    
    If sblnAbsAcc Then
        desfolderpath = ReplacePath(EveryFolder.Path, sstrOriPos, sstrDesPos)
        Set DesFolder = FSO.GetFolder(desfolderpath)
        EveryFolderPath = EveryFolder.Path
        For Each File In DesFolder.Files
            If Not FileExist(EveryFolderPath & "\" & File.Name) Then
                PutInToFilesList File.Path, sdicOldFilesList
            End If
        Next
    End If

    Next
    
    DeleteOldFiles sdicOldFilesList
    CopyUpdateFiles sdicCopyFilesList
    
End Sub
