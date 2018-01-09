'**************************************************************************
'目的：将FilesList中的所有文件夹删除
'**************************************************************************
Sub DeleteOldFiles(FilesList)
    N=FilesList.Count
    For i=1 To N
        FSO.DeleteFile FilesList.Item(i),True
    Next
End Sub

