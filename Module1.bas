Attribute VB_Name = "Module1"
Sub RenameAndSortFiles()
    Dim folderPath As String
    Dim fileName As String
    Dim fso As Object
    Dim file As Object
    Dim dateStr As String
    Dim i As Integer
    
    folderPath = "C:\Users\11185\2025"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For Each file In fso.GetFolder(folderPath).Files
        If InStr(file.Name, "����") = 0 And InStr(file.Name, "�}�N��") = 0 Then
            dateStr = Format(Now, "yyyy_mm_dd_hhmmss")
            fileName = "����f�[�^_" & dateStr & i & ".xlsx"
            file.Name = fileName
            i = i + 1
        End If
    Next
    MsgBox "�t�@�C����������"
    

End Sub

