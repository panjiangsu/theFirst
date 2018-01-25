'**************************************************************************
'目的：将FolderList中的所有文件夹删除
'**************************************************************************
Sub DeleteOldFolders(FoldersList)
    N=FoldersList.Count
    For i=1 To N
        FSO.DeleteFolder FoldersList.Item(i),True
    Next
End Sub

