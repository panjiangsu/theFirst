'**************************************************************************
'Ŀ�ģ���FilesList�е������ļ�����ԭʼλ�ø��Ƶ�Ŀ��λ��
'**************************************************************************
Sub CopyUpdateFiles(FilesList)
    N=FilesList.Count
    NeedSize=0
    For i=1 To N
        Set TempFile=FSO.GetFile(FilesList.Item(i))
        NeedSize=NeedSize+TempFile.Size
    Next
    Set D=FSO.GetDrive(FSO.GetDriveName(sstrDesPos))
    DesAvailableSize=D.AvailableSpace
    If NeedSize>DesAvailableSize Then Quit "Ŀ��λ���ļ��еĿռ���������,�޷���������"
    For i=1 To N
        FSO.CopyFile FilesList.Item(i),ReplacePath(FilesList.Item(i),sstrOriPos,sstrDesPos),True
    Next
End Sub

