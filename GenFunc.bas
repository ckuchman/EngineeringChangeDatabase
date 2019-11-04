Attribute VB_Name = "GenFunc"
Option Compare Database

Public Function CreateEmailWithOutlook( _
    MessageTo As String, _
    Subject As String, _
    MessageBody As String, _
    Optional Attachment As String)

    ' Define app variable and get Outlook using the "New" keyword
    Dim olApp As New Outlook.Application
    Dim mItem As Outlook.MailItem  ' An Outlook Mail item
    Dim mAttachments As Outlook.Attachments
 
    ' Create a new email object
    Set mItem = olApp.CreateItem(olMailItem)
    Set mAttachments = mItem.Attachments
    
    'Adds attachment to
    If Attachment <> "" Then
        mAttachments.Add Attachment
    End If
    
    ' Add the To/Subject/Body to the message and display the message
    With mItem
        .To = MessageTo
        .Subject = Subject
        .Body = MessageBody
        .Send       ' Send the message immediately
    End With

    ' Release all object variables
    Set mItem = Nothing
    Set olApp = Nothing

End Function

Public Function FileFolderExists(strFullPath As String) As Boolean
'Author       : Ken Puls (www.excelguru.ca)
'Macro Purpose: Check if a file or folder exists
'from http://www.excelguru.ca/content.php?157-Function-To-Check-If-File-Or-Directory-(Folder)-Exists
    On Error GoTo EarlyExit
    'myTest = Dir(strFullPath, vbDirectory)
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then FileFolderExists = True
EarlyExit:
    On Error GoTo 0
End Function

Public Function CheckFolderExistCreate(RootFolder As String, FolderName As String) As Boolean
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
'MsgBox "Inside CheckFolderExistCreate code" & ", RootFolder:" & RootFolder & " , FolderName:" & FolderName
'MsgBox "RootFolder:" & RootFolder
'MsgBox "FolderName:" & FolderName

If Right(RootFolder, 1) <> "\" Then RootFolder = RootFolder & "\"

'create directory if it does not exist
'If Dir(myTmpPath & "\D", vbDirectory) = "" Then Call MkDir(myTmpPath & "\D\")
'was not working...If Dir(RootFolder & CheckFolderName, vbDirectory) = "" Then
If fso.FolderExists(RootFolder & FolderName) = False Then
    'MsgBox "Didn't Exist, running make directory for RootFolder & FolderName:" & RootFolder & FolderName
    Call MkDir((RootFolder & FolderName))
'    If FSO.FolderExists(RootFolder & CheckFolderName) = True Then
'      MsgBox "Double check of make directory, it worked!"
'    End If
Else
'    MsgBox RootFolder & FolderName & "  Did Exist"
'    MsgBox "Did Exist RootFolder:" & RootFolder
'    MsgBox "Did Exist FolderName:" & FolderName
End If
Set fso = Nothing
End Function

Public Function CreateStandardFolder(FolderType As String, FormName As Form, Optional PR As Boolean)
    Dim RequiredPath As String
    Dim RequiredFolder As String
    
    RequiredPath = "\\wfs.local\Watson\Engineering\03_Engineering\ECNs\" & FolderType & "_Secondary_Documents\"
    
    If PR Then
        RequiredFolder = FormName.ODBCID
    Else
        RequiredFolder = FormName.ID
    End If
    
    If Not FileFolderExists(RequiredPath & RequiredFolder) Then
        Debug.Print "The RequiredPath did not exist but will be created"
        Call CheckFolderExistCreate(RequiredPath, RequiredFolder)
        If Not FileFolderExists(RequiredPath & RequiredFolder) Then
            Debug.Print "The RequiredPath is still missing"
        Else
            Debug.Print "The RequiredPath was successfully created"
        End If
    Else
        'Debug.Print "The RequiredPath exists!"
    End If
    
    If PR Then
        FormName.SecondaryDocumentationFolder = RequiredPath & FormName.ODBCID & "#" & RequiredPath & FormName.ODBCID
    Else
        FormName.SecondaryDocumentationFolder = RequiredPath & FormName.ID & "#" & RequiredPath & FormName.ID
    End If
        
End Function


