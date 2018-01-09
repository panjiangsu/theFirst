'***************************************************************************
'目的：检测输入的字符串是否是一个路径的格式
'***************************************************************************
Function IsPath(Str)
    '申明一个新的正则表达式
    Set re = New RegExp
     
    '分别设置不同的Pattern,用于多次检测以确定输入字符串是否为路径格式
    Pattern1 = "^[A-Za-z]:"    '该模式仅仅是确认字符串的起始部分是否符合驱动器表示
    Pattern2 = "\\[^\\/:\*\?<>\|\42]+"    '该模式是确认字符串的后面部分是否符合路径表示
    
    '设置正则表达式的参数
    re.IgnoreCase = True
    re.Global = False

    '申明一个由各部分匹配得到的重新组装字符串
    Dim MatchString

    '第一个模式检验
    re.pattern = Pattern1
    Set matches = re.Execute(Str)
    If matches.Count = 0 Then IsPath = False: Exit Function
    MatchString = matches.Item(0).Value

    
    '第二个模式检验
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

    '最终得到的重新组装字符串如果等于String,那么就是路径格式，否则就不是
    If MatchString = Str Then IsPath = True Else IsPath = False
End Function
