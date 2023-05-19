VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'Private WithEvents InboxItems As Outlook.Items

Private Sub Application_NewMail()

  ' Code executes when a new email arrives
   
    Dim objNamespace As Outlook.NameSpace
    Dim objFolder As Outlook.MAPIFolder
    Dim objItem As Object
    Dim objMail As Outlook.MailItem
    Dim objAttachment As Outlook.Attachment
    Dim objSubItem As Object
    Dim objSubItemMail As Outlook.MailItem
    Dim objSubAttachment As Outlook.Attachment
    Dim strPDFPath As String
    Dim errNumber As String
    Dim timeNow
    Dim timeNowStr As String
    Dim strAttachmentPath  As String
    
    'Test scenarios
    'two .msg attachments
        'with the same pdf   - result - only one pdf is saved - no error is raised. Overwrites any existing same named pdf.
        'with different pdfs - result - both pdfs are saved - no error is raised.
        'one with two pdf attachments and one with one.
        'one with two pdf attachments and one with none.
        'one with one pdf attachments and one with none.
    'one .msg attachment
        'with two pdfs
            'of same name - result - one single pdf is saved - no error is raised.
            'of different names - result - the multiple pdfs are saved - no error is raised. Tested only for 2 pdf attachments.
    'zero attachments
        'two pdfs only
            'routine does not run
        'one pdf only
            'routine does not run
        
On Error GoTo errorHandler
    strAttachmentPath = "C:\Users\Les\OneDrive - Analysisie\Documents\Outlook VBA\TempAttachedEmailItems\" '& objAttachment.FileName
    'Delete all .msg files in this folder
    fn_DeleteFilesInFolder (strAttachmentPath)
    
    
    ' Set the folder where the emails are located
    Set objNamespace = Application.GetNamespace("MAPI")
    Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)    '.Folders("Attachments")
    'MsgBox "The event was triggered."
'Set objOutlook = New Outlook.Application
'Set objNamespace = objOutlook.GetNamespace("MAPI")
'Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)


  ' Get the most recent email in the inbox folder
    Set objMail = objFolder.Items(objFolder.Items.Count)
    
        If TypeOf objMail Is Outlook.MailItem Then  'in the Inbox
            ' Check if the email has attachments
            If objMail.Attachments.Count > 0 Then
                ' Loop through each attachment
                For Each objAttachment In objMail.Attachments
                    If objAttachment.Type = olEmbeddeditem Then   ' Checks if the attachment is an email
                        ' Get the embedded item
                        'Set objSubItem = objAttachment.EmbeddedItem       'replaced line
                        Set objSubItem = objAttachment
                          ' Cast the embedded item to a MailItem object
                            'Set objSubItem = objAttachment.EmbeddedItem
                        Call fn_SaveAndOpenEmailAttachment(objSubItem)

                    End If
                Next objAttachment
            End If
        End If
  
    ' Clean up objects
    Set objSubAttachment = Nothing
    Set objSubItem = Nothing
    Set objAttachment = Nothing
    Set objMail = Nothing
    Set objItem = Nothing
    Set objFolder = Nothing
    Set objNamespace = Nothing

Exit Sub
    
''''''''''''''''''''
errorHandler:
'Call procedure to display error.
'DisplayError
timeNow = Now()
timeNowStr = CStr(timeNow)

errNumber = objMail.Subject & ":" & timeNowStr
fn_assignerror (errNumber)

MsgBox "There was an error : " & Err.Number & ": " & Err.Description & Chr(32) & " The recently received email with Subject: " & Chr(34) & objMail.Subject & Chr(34) & " gave rise to a VBA error handling its attachments." & Chr(32) & "Please report this to IT department quoting error number : " & errNumber


    ' Clean up objects
    Set objSubAttachment = Nothing
    Set objSubItem = Nothing
    Set objAttachment = Nothing
    Set objMail = Nothing
    Set objItem = Nothing
    Set objFolder = Nothing
    Set objNamespace = Nothing

Resume 'exitSub
'''''''''''''''''''''

End Sub

Function fn_SaveAndOpenEmailAttachment(objAttachment As Object)   'Outlook.MailItem

    'Dim objMail As Outlook.MailItem
    'Dim objAttachment As Outlook.Attachment
    Dim strAttachmentPath As String
    Dim objSavedItem As Outlook.MailItem
    Dim strSavePath As String
    Dim objExtractedAttachment As Outlook.Attachment
    Dim cleanedFileName  As String
    Dim i As Integer
    i = 1   'used for multiple attachments within one .msg mail item

    'Specify the folder path where you want to save the .msg file
    strAttachmentPath = "C:\Users\Les\OneDrive - Analysisie\Documents\Outlook VBA\TempAttachedEmailItems\" '& objAttachment.FileName

    'Delete all .msg files in this folder
    fn_DeleteFilesInFolder (strAttachmentPath)

    'Debug.Print objAttachment.FileName
    'strip out . and spaces from filename
    cleanedFileName = Replace(Replace(objAttachment.FileName, ".", ""), " ", "")
    cleanedFileName = Replace(cleanedFileName, "msg", ".msg")

    objAttachment.SaveAsFile strAttachmentPath & cleanedFileName    ' Save the attachment as a .msg file. Overwrites existing duplicate files.

    ' Open the saved .msg file
    Set objSavedItem = Application.Session.OpenSharedItem(strAttachmentPath & cleanedFileName)
    
    ' Perform any additional actions with the opened .msg file
    ' Extract the attachment from the opened .msg file
    If objSavedItem.Attachments.Count > 0 Then
'        Set objExtractedAttachment = objSavedItem.Attachments(1)       'replaced this line with the following Do While Loop because there is possibility of having
'        multiple attachments per one single .msg item.
            
            Do While i <= objSavedItem.Attachments.Count
                Set objExtractedAttachment = objSavedItem.Attachments(i)        'Starts at 1.       M.1AC.pdf
                
                ' Specify the folder path where you want to save the extracted attachment
                    strSavePath = "C:\Users\Les\OneDrive - Analysisie\Documents\Outlook VBA\TempAttachedEmailItems\ExtractedPDfs\" '& objAttachment.FileName
                
                ' Save the extracted attachment to the specified folder
                If Right(objExtractedAttachment.FileName, 3) = "pdf" Then
                    objExtractedAttachment.SaveAsFile strSavePath & objExtractedAttachment.FileName
                End If
                i = i + 1
            Loop

    End If
    
        
   ' Close the open email
    objSavedItem.Close olDiscard
    Set objSavedItem = Nothing
    'Kill strAttachmentPath & cleanedFileName
    
    ' Clean up objects
    Set objExtractedAttachment = Nothing
    Set objAttachment = Nothing
    'Set objMail = Nothing

End Function

Function fn_assignerror(errornumber As String)
    Dim filePath As String
    Dim fileNumber As Integer
    Dim appendString As String
    
    ' Set the file path
    filePath = "C:\Users\Les\OneDrive - Analysisie\Documents\Outlook VBA\TempAttachedEmailItems\Error Log\EmailAttachmentExtractorErrorLog.txt"
    
    ' Build the string to be appended
    appendString = errornumber
    
    ' Open the file in append mode
    fileNumber = FreeFile
    Open filePath For Append As #fileNumber
    
    ' Write the string to the file
    Print #fileNumber, appendString
    
    ' Close the file
    Close #fileNumber
End Function

Sub fn_DeleteFilesInFolder(folderPath As String)
    Dim file As String
    
    ' Check if the folder path ends with a backslash
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    
    ' Loop through each file in the folder
    file = Dir(folderPath & "*.*")
    Do While Len(file) > 0
        ' Delete the file using the Kill method
        Kill folderPath & file
        
        ' Move to the next file
        file = Dir
    Loop
End Sub

