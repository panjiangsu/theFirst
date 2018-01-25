'**************************************************************************
'目的：将FilesList中的所有文件夹由原始位置复制到目标位置
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
    If NeedSize>DesAvailableSize Then Quit "目标位置文件夹的空间容量不够,无法继续操作"
    For i=1 To N
        FSO.CopyFile FilesList.Item(i),ReplacePath(FilesList.Item(i),sstrOriPos,sstrDesPos),True
    Next
End Sub

