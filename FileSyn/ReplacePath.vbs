'*************************************************************************
'Ŀ�ģ���·����Path��ʼ���ֵ�FirstPath�滻ΪSecondPath	
'*************************************************************************
Function ReplacePath(Path,FirstPath,SecondPath)
    ReplacePath=SecondPath & Right(Path,Len(Path)-Len(Firstpath))
End Function

