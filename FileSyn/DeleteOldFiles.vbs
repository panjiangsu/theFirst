'**************************************************************************
'Ŀ�ģ���FilesList�е������ļ���ɾ��
'**************************************************************************
Sub DeleteOldFiles(FilesList)
    N=FilesList.Count
    For i=1 To N
        FSO.DeleteFile FilesList.Item(i),True
    Next
End Sub

