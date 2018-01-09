'***************************************************************************
'Ŀ�ģ����������ַ����Ƿ���һ��·���ĸ�ʽ
'***************************************************************************
Function IsPath(Str)
    '����һ���µ�������ʽ
    Set re = New RegExp
     
    '�ֱ����ò�ͬ��Pattern,���ڶ�μ����ȷ�������ַ����Ƿ�Ϊ·����ʽ
    Pattern1 = "^[A-Za-z]:"    '��ģʽ������ȷ���ַ�������ʼ�����Ƿ������������ʾ
    Pattern2 = "\\[^\\/:\*\?<>\|\42]+"    '��ģʽ��ȷ���ַ����ĺ��沿���Ƿ����·����ʾ
    
    '����������ʽ�Ĳ���
    re.IgnoreCase = True
    re.Global = False

    '����һ���ɸ�����ƥ��õ���������װ�ַ���
    Dim MatchString

    '��һ��ģʽ����
    re.pattern = Pattern1
    Set matches = re.Execute(Str)
    If matches.Count = 0 Then IsPath = False: Exit Function
    MatchString = matches.Item(0).Value

    
    '�ڶ���ģʽ����
    re.pattern = MatchString & Pattern2
    Set matches =re.Execute(Str)
    If matches.Count = 0 Then IsPath = False: Exit Function
    MatchString = matches.Item(0).Value
    MatchString = Replace(MatchString, "\", "\\")
    Do
        re.pattern = MatchString & Pattern2
        Set matches = re.Execute(Str)
        If matches.Count = 0 Then Exit Do
        MatchString = matches.Item(0).Value
        MatchString = Replace(MatchString, "\", "\\")
    Loop
    MatchString = Replace(MatchString, "\\", "\")

    '���յõ���������װ�ַ����������String,��ô����·����ʽ������Ͳ���
    If MatchString = Str Then IsPath = True Else IsPath = False
End Function
