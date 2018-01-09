'**************************************************************************
'目的：将FoldesrList中的所有文件夹由原始位置复制到目标位置
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
    If NeedSize>DesAvailableSize Then Quit "目标位置文件夹的空间容量不够,无法继续操作"
    For i=1 To N
        FSO.CopyFolder FoldersList.Item(i),ReplacePath(FoldersList.Item(i),sstrOriPos,sstrDesPos),False
    Next
End Sub

