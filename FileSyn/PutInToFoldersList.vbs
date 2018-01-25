'**********************************************************************
'目的：将特定的文件夹(路径)加入列表中
'**********************************************************************
Sub PutInToFoldersList(FolderPath,NeedList)
    n=NeedList.count
    NeedList.Add n+1,FolderPath
End Sub
