'**************************************************************************
'Ŀ�ģ���FolderList�е������ļ���ɾ��
'**************************************************************************
Sub DeleteOldFolders(FoldersList)
    N=FoldersList.Count
    For i=1 To N
        FSO.DeleteFolder FoldersList.Item(i),True
    Next
End Sub

