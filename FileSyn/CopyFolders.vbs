'**************************************************************************
'Ŀ�ģ���FoldesrList�е������ļ�����ԭʼλ�ø��Ƶ�Ŀ��λ��
'**************************************************************************
Sub CopyFolders(FoldersList)
    N=FoldersList.Count
    NeedSize=0
    For i=1 To N
        Set TempFolder=FSO.GetFolder(FoldersList.Item(i))
        NeedSize=NeedSize+TempFolder.Size
    Next
    Set D=FSO.GetDrive(FSO.GetDriveName(sstrDesPos))
    DesAvailableSize=D.AvailableSpace
    If NeedSize>DesAvailableSize Then Quit "Ŀ��λ���ļ��еĿռ���������,�޷���������"
    For i=1 To N
        FSO.CopyFolder FoldersList.Item(i),ReplacePath(FoldersList.Item(i),sstrOriPos,sstrDesPos),False
    Next
End Sub

