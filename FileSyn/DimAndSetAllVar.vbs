Dim sstrOriPos
Dim sstrDesPos
Dim sblnAbsAcc
Dim sdicCopyFoldersList
Dim sdicUpdateFoldersList
Dim sdicCopyFilesList
Dim sdicOldFoldersList
Dim sdicOldFilesList
Dim sblnDoneInStart
Dim FSO

Sub SetVars()
    Set FSO=CreateObject("Scripting.FileSystemObject")
    
    Set sdicCopyFoldersList=CreateObject("Scripting.Dictionary")
    Set sdicUpdateFoldersList=CreateObject("Scripting.Dictionary")
    Set sdicCopyFoldersList=CreateObject("Scripting.Dictionary")
    Set sdicCopyFilesList=CreateObject("Scripting.Dictionary")
    Set sdicOldFoldersList=CreateObject("Scripting.Dictionary")
    Set sdicOldFilesList=CreateObject("Scripting.Dictionary")
End Sub

