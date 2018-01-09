'*************************************************************************
'目的：将路径名Path开始部分的FirstPath替换为SecondPath	
'*************************************************************************
Function ReplacePath(Path,FirstPath,SecondPath)
    ReplacePath=SecondPath & Right(Path,Len(Path)-Len(Firstpath))
End Function

