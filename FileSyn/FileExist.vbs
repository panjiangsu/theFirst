'**********************************************************************
'目的：检测给定的文件是否存在
'**********************************************************************
Function FileExist(File)
    FileExist=FSO.FileExists(File)
End Function
