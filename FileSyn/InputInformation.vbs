'***********************************************************************************
'Ŀ�ģ�������������������Ϣ�Ƿ�Ϊ�գ���Ϊ��������˳�����ִ��
'*********************************************************************************** 
Function AreNoEmpty(Start,Des)
    If Start="" Then StartSign=0 Else StartSign=1
    If Des="" Then EndSign=0 Else EndSign=2
    Select Case StartSign+EndSign
        Case 0
            Quit "��ʼ��Ŀ���ļ���λ�ö�Ϊ��"
        Case 1
            Quit "Ŀ���ļ���λ��Ϊ��"
        Case 2 
            Quit "��ʼ�ļ���λ��Ϊ��"
        Case Else
            AreNoEmpty=True
    End Select
End Function

'***********************************************************************************
'Ŀ�ģ�������������������Ϣ�Ƿ�Ϊ��ȷ��·����ʽ��ֱ����������
'            ȷ��·����ʽΪֹ
'*********************************************************************************** 
Function ArePath(Start,Des)
    Do 
    If IsPath(Start) Then StartSign=1 Else StartSign=0
    If IsPath(Des) Then EndSign=2 Else EndSign=0
    Select Case StartSign+EndSign
        Case 0
            MsgBox "��ʼ��Ŀ���ļ���λ�ö����Ǳ�׼��·����ʽ������������"
            StartPos=InputBox("��������ʼλ���ļ���:")
            EndPos=InputBox("������Ŀ��λ���ļ���:")
            AreNoEmpty StartPos,EndPos
        Case 1
            MsgBox "Ŀ���ļ���λ�ò��Ǳ�׼��·����ʽ������������"
            EndPos=InputBox("������Ŀ��λ���ļ���:")
            AreNoEmpty StartPos,EndPos
        Case 2 
            MsgBox "��ʼ�ļ���λ�ò��Ǳ�׼��·����ʽ������������"
            StartPos=InputBox("��������ʼλ���ļ���:")
            AreNoEmpty StartPos,EndPos
        Case 3
            ArePath=True
            Exit Do
    End Select
    Loop
End Function

'***********************************************************************************
'Ŀ�ģ�����������������·���Ĵ��ڣ�������Ӧ�Ĵ���
'***********************************************************************************
Function AreExist(Start,Des)
    If FSO.FolderExists(Start) Then StartSign=1 Else StartSign=0
    If FSO.FolderExists(Des) Then EndSign=2 Else EndSign=0
    Select Case StartSign+EndSign
        Case 0
            Quit "��ʼ��Ŀ���ļ���λ�ö�������"
        Case 1
            sdicCopyFoldersList.Add 1,Start
            sblnDoneInStart=True
        Case 2
            Quit "��ʼ�ļ���λ�ò�����"
        Case 3
            AreExist=True
    End Select
End Function

'********************************************************************
'Ŀ�ģ���������ԭʼλ�á�Ŀ��λ���Ƿ����·������
'            Ҫ�󲢽�������ֵ���Ƿ���Ҫһ���Ա��ݵ�ֵһ��
'            ��ֵ����Ӧ�ı������ڸù����л�Ӧ�ÿ��ǵ�һ��
'            ��������������Ŀ��λ���޶�Ӧ��ԭʼλ���ļ�
'            �У���ô��ֱ�ӽ�ԭʼλ���ļ��е�������ȫ����
'            ��Ŀ��λ�á�
'*********************************************************************
Sub InputInformation()
    StartPos=CStr(InputBox("��������ʼλ���ļ���:"))
    EndPos=CStr(InputBox("������Ŀ��λ���ļ���:"))
    AreNoEmpty StartPos,EndPos
    ArePath StartPos,EndPos
    AreExist StartPos,EndPos
    sstrOriPos=StartPos
    sstrDesPos=EndPos
    If sblnDoneInStart Then CopyFolders(sdicCopyFoldersList):Exit Sub
    Button=MsgBox("����Ҫ��ȷ�����ļ�/�ļ��е�һ����?",vbYesNo,"��ȷ����ȷ��")
    If Button=6 Then 
        sblnAbsAcc=True
        tInf="���������ʼλ���ļ���Ϊ��"&sstrOriPos&vbNewline
        tInf= tInf&"�������Ŀ��λ���ļ���Ϊ��"&sstrDesPos&vbNewline
        tInf= tInf&"��ѡ������Ҫ��ȷ�����ļ�/�ļ��е�һ�£��������������ʼλ���ļ��е���Ŀ��λ���ļ��еĻ�����ôԭʼλ���ļ������ж�Ŀ��λ���ļ�����û�е����ļ��л��ļ�����ɾ����������ԭʼλ���ļ����е����ݶ�ʧ�����ٴ�ȷ������������ļ����Ƿ���ȷ��"
        blnSign=MsgBox(tinf,vbOkCancel,"��Ҫ��ʾ��")
        If blnSign=vbCancel Then Quit "������ȷ���������ʼ��Ŀ��λ���ļ��е���ȷ�Զ�ѡ���˳���"
    Else
        sblnAbsAcc=False
    End If
    
End Sub
