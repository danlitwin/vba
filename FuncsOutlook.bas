
Private Function FindFolder(Name As String, Optional TheFolders As Outlook.Folders)
    Dim SubFolder As Outlook.MAPIFolder
    
    On Error Resume Next
    
    If IsMissing(TheFolders) Then Set TheFolders = Application.Session.Folders
    Set FindFolder = Nothing
    
    For Each SubFolder In TheFolders
        If (LCase(Left(SubFolder.Name, 8)) = "public f") Or (LCase(Left(SubFolder.Name, 8)) = "filesite") Then
            ' skip it
        ElseIf LCase(SubFolder.Name) Like LCase(Name) Then
            Set FindFolder = SubFolder
            Exit For
        Else
            Set FindFolder = FindFolder(Name, SubFolder.Folders)
            If Not FindFolder Is Nothing Then Exit For
        End If
    Next
End Function
