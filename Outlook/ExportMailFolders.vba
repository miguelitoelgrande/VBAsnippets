' Quickly export a bunch of mails from a folder (including subfolders) to the filesystem (as single .msg files)
' Also generates subfolders according to the structure of the folder in outlook
'
' Based on: https://www.datanumen.com/blogs/quickly-export-subfolders-items-outlook-folder-windows-folder/
' with some fixes
'
Private objFileSystem As Object
 
Private Sub ExportFolderWithAllItems()
    Dim objFolder As Outlook.Folder
    Dim strPath As String
 
    'Specify the root local folder
    'Change it as per your needs
    strPath = "D:\MailExport\"
 
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
 
    'Select a Outlook PST file or Outlook folder
    Set objFolder = Outlook.Application.Session.PickFolder
 
    Call ProcessFolders(objFolder, strPath)
 
    MsgBox "Complete", vbExclamation
End Sub
 
 
' MM Snippet from: https://stackoverflow.com/questions/37024107/remove-unicode-characters-in-a-string
Public Function StripNonAsciiChars(ByVal InputString As String) As String
    Dim i As Integer
    Dim RegEx As Object
    Set RegEx = CreateObject("VBScript.RegExp")
    With RegEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = "[^\u0000-\u007F]"
        StripNonAsciiChars = Excel.WorksheetFunction.Trim(RegEx.Replace(InputString, " "))
    End With
End Function
  
 
 
Private Sub ProcessFolders(objCurrentFolder As Outlook.Folder, strCurrentPath As String)
    Dim objItem As Object
    Dim strSubject, strFileName, strFilePath As String
    Dim objSubfolder As Outlook.Folder
 
    'Create the local folder based on the Outlook folder
    strCurrentPath = strCurrentPath & objCurrentFolder.Name
    objFileSystem.CreateFolder strCurrentPath
 
    For Each objItem In objCurrentFolder.Items
 
        strSubject = objItem.Subject
 
        'Remove unsupported characters in the subject
        strSubject = Replace(strSubject, "/", " ")
        strSubject = Replace(strSubject, "\", " ")
		strSubject = Replace(strSubject, "|", " ")  ' MM: more Windows Filenames do not like
		strSubject = Replace(strSubject, "<", " ")  ' MM: more Windows Filenames do not like
		strSubject = Replace(strSubject, ">", " ")  ' MM: more Windows Filenames do not like		
		strSubject = Replace(strSubject, "!", " ")  ' MM: more Windows Filenames do not like		
		strSubject = Replace(strSubject, "*", " ")  ' MM: more Windows Filenames do not like		
        strSubject = Replace(strSubject, ":", "")
        strSubject = Replace(strSubject, "?", " ")
        strSubject = Replace(strSubject, Chr(34), " ")
		strSubject = StripNonAsciiChars(strSubject) ' MM: more Windows Filenames do not like

        strFileName = strSubject & ".msg"
 
        i = 0
        Do Until False
           strFilePath = strCurrentPath & "\" & strFileName
           'Check if there exist a file in the same name
           If objFileSystem.FileExists(strFilePath) Then
              'Add a sequence order to the file name
              i = i + 1
              strFileName = strSubject & " (" & i & ").msg"
           Else
              Exit Do
          End If
        Loop
 
        'Save as MSG file
        objItem.SaveAs strFilePath, olMSG
    Next
 
    'Process subfolders recursively
    If objCurrentFolder.folders.Count > 0 Then
       For Each objSubfolder In objCurrentFolder.folders
           Call ProcessFolders(objSubfolder, strCurrentPath & "\")
       Next
    End If
End Sub
