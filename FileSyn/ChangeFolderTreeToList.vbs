'********************************************************************************
'目的：遍历给定的起始文件夹下的所有文件夹(树型)并将这些文件夹
'            按照要求转换为列表型(线型)结构。
'********************************************************************************
'********************************************************************************
'伪码算法表示
'描述：
'该过程中的核心算法是遍历树状结构的数据
'采用的遍历算法思想：
'设立一个标记作为是否还要进行搜索叶节点的指示，该标记intSign
'和通过遍历得到的字典对象Cycle(仅在过程中使用的对象)之间的关系
'是：intSign=Cycle.Count-Start，其中Start是每一次遍历之前的
'Cycle.Count值，通过该算式可以判断在上一次遍历的过程中是否有新
'增的叶节点(intSign>0就有)，并由此来决定是否需要继续遍历。
'遍历算法：
'|设置Start=0并将根节点纳入Cycle对象中并计算intSign=Cycle.Count-Start
'使遍历过程能够由根节点进入
'
'|1*逐个搜索Cycle中的每一个节点，进行所需要的处理并根据处理的结果
'决定是否需要增加搜索节点中的可能的叶节点，作相应处理(增加或不增加)
'
'|2*一次遍历结束后，重新计算intSign的值
'
'|判断intSign的值，如果>0，那么就重复1*和2*步骤，否则退出执行
'***********************************************************************************

Sub ChangeFolderTreeToList(OpSign)
    '定义遍历算法中所需要的变量及初始值设置
    '记录各节点路径的Cycle对象变量
    Dim Cycle
    Set Cycle = CreateObject("scripting.Dictionary")
    If OpSign = "CU" Then
       Cycle.Add 1, sstrOriPos 'Cycle对象的键值按照增加项目的先后依次附值为1、2、3...，项目为节点路径
    Else
       Cycle.Add 1, sstrDesPos
    End If
    
    '是否需要继续循环的标识初始值设置
    intsign = 1
    
    '根据所要搜索的文件夹类型，分别进行搜索处理
    Select Case OpSign
        Case "CU"
        
        '根据intSign值是否大于0来确定是否还有需要遍历的节点
        Do While intsign > 0
            Start = Cycle.Count '每次循环的起始位置初始值设置
            For i = Cycle.Count - intsign + 1 To Cycle.Count '每次搜索节点由新增的项目开始到新增项目结束
            
               '检查并设置起始位置文件夹中的文件夹是否为需更新文件夹
               If FolderExist(ReplacePath(Cycle.Item(i), sstrOriPos, sstrDesPos)) Then '检查子文件夹是否存在于目标文件夹中
                   PutInToFoldersList Cycle.Item(i), sdicUpdateFoldersList '将需更新文件夹放入需更新文件夹列表中
            
            
                   Set f = fso.GetFolder(Cycle.Item(i)) '得到节点-文件夹
                   Set Fs = f.SubFolders '得到节点的子文件夹
            
                   '如果节点有子文件夹则加入Cycle-遍历各节点
                   n = Fs.Count
                   If n > 0 Then
                      For Each sf In Fs
                      Cycle.Add (Cycle.Count + 1), sf.Path
                      Next
                   End If
            
            '如果是目标位置文件夹中没有的文件夹则纳入需拷贝文件夹列表中
            Else
               PutInToFoldersList Cycle.Item(i), sdicCopyFoldersList
            End If
            
            Next
                
        '确定是否有增加的节点
        intsign = Cycle.Count - Start
        
    Loop
    
        Case "D"
    
    
        Do While intsign > 0
            Start = Cycle.Count '每次循环的起始位置初始值设置
            For i = Cycle.Count - intsign + 1 To Cycle.Count '每次搜索节点由新增的项目开始到新增项目结束
                
                '检查并设置目的位置文件夹中的文件夹是否为删除文件夹
                If Not FolderExist(ReplacePath(Cycle.Item(i), sstrDesPos, sstrOriPos)) Then
                
                   PutInToFoldersList Cycle.Item(i), sdicOldFoldersList
                
                Else
               
            
                   Set f = fso.GetFolder(Cycle.Item(i)) '得到节点-文件夹
                   Set Fs = f.SubFolders '得到节点的子文件夹
            
                   '如果节点有子文件夹则加入Cycle-遍历各节点
                   n = Fs.Count
                   If n > 0 Then
                      For Each sf In Fs
                          Cycle.Add (Cycle.Count + 1), sf.Path
                      Next
                   End If
                   
                End If
            
        Next
                
        '确定是否有增加的节点
        intsign = Cycle.Count - Start
        
    Loop
    
    End Select
    
    
End Sub