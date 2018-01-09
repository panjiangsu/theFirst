'***********************************************************************************
'目的：检查两个输入的两个信息是否为空，若为空则出错退出程序执行
'*********************************************************************************** 
Function AreNoEmpty(Start,Des)
    If Start="" Then StartSign=0 Else StartSign=1
    If Des="" Then EndSign=0 Else EndSign=2
    Select Case StartSign+EndSign
        Case 0
            Quit "起始和目标文件夹位置都为空"
        Case 1
            Quit "目标文件夹位置为空"
        Case 2 
            Quit "起始文件夹位置为空"
        Case Else
            AreNoEmpty=True
    End Select
End Function

'***********************************************************************************
'目的：检查两个输入的两个信息是否为正确的路径格式，直至输入了正
'            确的路径格式为止
'*********************************************************************************** 
Function ArePath(Start,Des)
    Do 
    If IsPath(Start) Then StartSign=1 Else StartSign=0
    If IsPath(Des) Then EndSign=2 Else EndSign=0
    Select Case StartSign+EndSign
        Case 0
            MsgBox "起始和目标文件夹位置都不是标准的路径格式，请重新输入"
            StartPos=InputBox("请输入起始位置文件夹:")
            EndPos=InputBox("请输入目标位置文件夹:")
            AreNoEmpty StartPos,EndPos
        Case 1
            MsgBox "目标文件夹位置不是标准的路径格式，请重新输入"
            EndPos=InputBox("请输入目标位置文件夹:")
            AreNoEmpty StartPos,EndPos
        Case 2 
            MsgBox "起始文件夹位置不是标准的路径格式，请重新输入"
            StartPos=InputBox("请输入起始位置文件夹:")
            AreNoEmpty StartPos,EndPos
        Case 3
            ArePath=True
            Exit Do
    End Select
    Loop
End Function

'***********************************************************************************
'目的：检查两个输入的两个路径的存在，并作相应的处理
'***********************************************************************************
Function AreExist(Start,Des)
    If FSO.FolderExists(Start) Then StartSign=1 Else StartSign=0
    If FSO.FolderExists(Des) Then EndSign=2 Else EndSign=0
    Select Case StartSign+EndSign
        Case 0
            Quit "起始和目标文件夹位置都不存在"
        Case 1
            sdicCopyFoldersList.Add 1,Start
            sblnDoneInStart=True
        Case 2
            Quit "起始文件夹位置不存在"
        Case 3
            AreExist=True
    End Select
End Function

'********************************************************************
'目的：检查输入的原始位置、目标位置是否符合路径规则
'            要求并将这两个值和是否需要一致性备份的值一起
'            赋值给相应的变量。在该过程中还应该考虑到一个
'            特殊情况，即如果目标位置无对应的原始位置文件
'            夹，那么就直接将原始位置文件夹的内容完全复制
'            到目标位置。
'*********************************************************************
Sub InputInformation()
    StartPos=CStr(InputBox("请输入起始位置文件夹:"))
    EndPos=CStr(InputBox("请输入目标位置文件夹:"))
    AreNoEmpty StartPos,EndPos
    ArePath StartPos,EndPos
    AreExist StartPos,EndPos
    sstrOriPos=StartPos
    sstrDesPos=EndPos
    If sblnDoneInStart Then CopyFolders(sdicCopyFoldersList):Exit Sub
    Button=MsgBox("您需要精确控制文件/文件夹的一致吗?",vbYesNo,"精确控制确认")
    If Button=6 Then 
        sblnAbsAcc=True
        tInf="您输入的起始位置文件夹为："&sstrOriPos&vbNewline
        tInf= tInf&"您输入的目标位置文件夹为："&sstrDesPos&vbNewline
        tInf= tInf&"您选择了需要精确控制文件/文件夹的一致，这样如果您将起始位置文件夹当成目标位置文件夹的话，那么原始位置文件夹中有而目标位置文件夹中没有的子文件夹或文件将被删除，导致您原始位置文件夹中的数据丢失。请再次确认输入的两个文件夹是否正确。"
        blnSign=MsgBox(tinf,vbOkCancel,"重要提示！")
        If blnSign=vbCancel Then Quit "您不能确定输入的起始和目标位置文件夹的正确性而选择退出。"
    Else
        sblnAbsAcc=False
    End If
    
End Sub
