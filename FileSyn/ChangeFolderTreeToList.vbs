'********************************************************************************
'Ŀ�ģ�������������ʼ�ļ����µ������ļ���(����)������Щ�ļ���
'            ����Ҫ��ת��Ϊ�б���(����)�ṹ��
'********************************************************************************
'********************************************************************************
'α���㷨��ʾ
'������
'�ù����еĺ����㷨�Ǳ�����״�ṹ������
'���õı����㷨˼�룺
'����һ�������Ϊ�Ƿ�Ҫ��������Ҷ�ڵ��ָʾ���ñ��intSign
'��ͨ�������õ����ֵ����Cycle(���ڹ�����ʹ�õĶ���)֮��Ĺ�ϵ
'�ǣ�intSign=Cycle.Count-Start������Start��ÿһ�α���֮ǰ��
'Cycle.Countֵ��ͨ������ʽ�����ж�����һ�α����Ĺ������Ƿ�����
'����Ҷ�ڵ�(intSign>0����)�����ɴ��������Ƿ���Ҫ����������
'�����㷨��
'|����Start=0�������ڵ�����Cycle�����в�����intSign=Cycle.Count-Start
'ʹ���������ܹ��ɸ��ڵ����
'
'|1*�������Cycle�е�ÿһ���ڵ㣬��������Ҫ�Ĵ������ݴ���Ľ��
'�����Ƿ���Ҫ���������ڵ��еĿ��ܵ�Ҷ�ڵ㣬����Ӧ����(���ӻ�����)
'
'|2*һ�α������������¼���intSign��ֵ
'
'|�ж�intSign��ֵ�����>0����ô���ظ�1*��2*���裬�����˳�ִ��
'***********************************************************************************

Sub ChangeFolderTreeToList(OpSign)
    '��������㷨������Ҫ�ı�������ʼֵ����
    '��¼���ڵ�·����Cycle�������
    Dim Cycle
    Set Cycle = CreateObject("scripting.Dictionary")
    If OpSign = "CU" Then
       Cycle.Add 1, sstrOriPos 'Cycle����ļ�ֵ����������Ŀ���Ⱥ����θ�ֵΪ1��2��3...����ĿΪ�ڵ�·��
    Else
       Cycle.Add 1, sstrDesPos
    End If
    
    '�Ƿ���Ҫ����ѭ���ı�ʶ��ʼֵ����
    intsign = 1
    
    '������Ҫ�������ļ������ͣ��ֱ������������
    Select Case OpSign
        Case "CU"
        
        '����intSignֵ�Ƿ����0��ȷ���Ƿ�����Ҫ�����Ľڵ�
        Do While intsign > 0
            Start = Cycle.Count 'ÿ��ѭ������ʼλ�ó�ʼֵ����
            For i = Cycle.Count - intsign + 1 To Cycle.Count 'ÿ�������ڵ�����������Ŀ��ʼ��������Ŀ����
            
               '��鲢������ʼλ���ļ����е��ļ����Ƿ�Ϊ������ļ���
               If FolderExist(ReplacePath(Cycle.Item(i), sstrOriPos, sstrDesPos)) Then '������ļ����Ƿ������Ŀ���ļ�����
                   PutInToFoldersList Cycle.Item(i), sdicUpdateFoldersList '��������ļ��з���������ļ����б���
            
            
                   Set f = fso.GetFolder(Cycle.Item(i)) '�õ��ڵ�-�ļ���
                   Set Fs = f.SubFolders '�õ��ڵ�����ļ���
            
                   '����ڵ������ļ��������Cycle-�������ڵ�
                   n = Fs.Count
                   If n > 0 Then
                      For Each sf In Fs
                      Cycle.Add (Cycle.Count + 1), sf.Path
                      Next
                   End If
            
            '�����Ŀ��λ���ļ�����û�е��ļ����������追���ļ����б���
            Else
               PutInToFoldersList Cycle.Item(i), sdicCopyFoldersList
            End If
            
            Next
                
        'ȷ���Ƿ������ӵĽڵ�
        intsign = Cycle.Count - Start
        
    Loop
    
        Case "D"
    
    
        Do While intsign > 0
            Start = Cycle.Count 'ÿ��ѭ������ʼλ�ó�ʼֵ����
            For i = Cycle.Count - intsign + 1 To Cycle.Count 'ÿ�������ڵ�����������Ŀ��ʼ��������Ŀ����
                
                '��鲢����Ŀ��λ���ļ����е��ļ����Ƿ�Ϊɾ���ļ���
                If Not FolderExist(ReplacePath(Cycle.Item(i), sstrDesPos, sstrOriPos)) Then
                
                   PutInToFoldersList Cycle.Item(i), sdicOldFoldersList
                
                Else
               
            
                   Set f = fso.GetFolder(Cycle.Item(i)) '�õ��ڵ�-�ļ���
                   Set Fs = f.SubFolders '�õ��ڵ�����ļ���
            
                   '����ڵ������ļ��������Cycle-�������ڵ�
                   n = Fs.Count
                   If n > 0 Then
                      For Each sf In Fs
                          Cycle.Add (Cycle.Count + 1), sf.Path
                      Next
                   End If
                   
                End If
            
        Next
                
        'ȷ���Ƿ������ӵĽڵ�
        intsign = Cycle.Count - Start
        
    Loop
    
    End Select
    
    
End Sub