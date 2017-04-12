Attribute VB_Name = "MainModule"
' This little VBA script will iterate through a set of selected messages
' saving each attachment to a folder corresponding to the conversation topic,
' subject, or "Uncategorized". It will delete the attachments from the mail
' message and insert links into the message indicating where the attachment
' has been saved.
'
' Attachments are stored in folder under a folder named "EmailAttachments"
' in your documents folder. If you want to store them someplace else change
' the line of code below.
'
' The code will ignore attachments starting with "image" which are typically
' embedded images such as signature logos or images pasted directly into the
' message body.
'
' To use it, select one or more messages in a mail or search folder then run
' "SaveAllAttachments". You will be prompted to allow or deny access to modify
' messages. You can choose to allow for a period of time or you will be
' prompted for each attachment.
'
Public Sub SaveAllAttachments()
    Dim objOLapp As Outlook.Application
    Dim objSelection As Outlook.Selection
    Dim Item As Object
    Dim attItem As Outlook.Attachment
    Dim strBasePath As String
    Dim strSavedFile As String
    Dim strMsg As String
    Dim i As Integer

    ' Initialize some objects
    Set objOLapp = CreateObject("Outlook.Application")
    Set objSelection = objOLapp.ActiveExplorer.Selection
    strBasePath = CreateObject("WScript.Shell").SpecialFolders(16) & "\EmailAttachments\"

    For Each Item In objSelection
        If TypeOf Item Is Outlook.MailItem Then
            Dim oMail As Outlook.MailItem: Set oMail = Item
            If oMail.Attachments.Count > 0 Then
                ' Can't use for each because the iterator gets
                ' screwed up when you delete an attachment
                For i = oMail.Attachments.Count To 1 Step -1
                    Set attItem = oMail.Attachments.Item(i)
                    If attItem.Type <> olOLE And Left(attItem.DisplayName, 5) <> "image" Then
                        strSavedFile = SaveAttachment(attItem, strBasePath)
                        AnnotateMessage oMail, strSavedFile
                        Debug.Print strSavedFile
                        attItem.Delete
                    End If
                Next i
            End If
        End If
    Next

End Sub


Private Sub DisplayAttachmentMetadata(objItem As Outlook.Attachment)
    Dim strType As String
    Dim strMsg As String
    
    On Error GoTo ErrRoutine
    
    If objItem Is Nothing Then
        MsgBox "A reference for the attachment object could not be retrieved."
    Else
        ' Display information about the item.
        Select Case objItem.Type
            Case OlAttachmentType.olByValue
                strType = "File"
            Case OlAttachmentType.olByReference
                strType = "Referenced File"
            Case OlAttachmentType.olEmbeddeditem
                strType = "Embedded Item"
            Case OlAttachmentType.olOLE
                strType = "OLE Object"
        End Select

        strMsg = "Conversation Topic: " & objItem.Parent.ConversationTopic & vbCrLf & _
                    "Subject: " & objItem.Parent.Subject & vbCrLf & _
                    "Attachment Type is " & strType & vbCrLf & _
                    "Size is " & Round(objItem.Size / 1024, 0) & "K" & vbCrLf
        If objItem.Type <> olOLE Then
                    strMsg = strMsg & "File name is " & objItem.FileName & vbCrLf
        End If
        strMsg = strMsg & vbCrLf

        Debug.Print strMsg
'        MsgBox strMsg
    End If

EndRoutine:
    Set objItem = Nothing
    Set objNamespace = Nothing
    Exit Sub

ErrRoutine:
    MsgBox Err.Number & " - " & Err.Description, _
        vbOKOnly Or vbCritical, _
        "DisplayItemMetadata"
    GoTo EndRoutine
End Sub

' Saves an attachment to a subdirectory of the base directory named for either the
' Conversation Topic or Subject in that order if available or "Uncategorized" otherwise.
'
' TODO: Make sure there are no double backslashes in the entire path string
'
' Returns a string with the file name (full path) used to save the attachment

Private Function SaveAttachment(ByVal objAttachment As Outlook.Attachment, ByVal strBaseDir As String) As String
    Dim strTargetDir As String
    Dim strSavedFile As String
    Dim strSubDir As String
    Dim bBool As Boolean

    If objAttachment.Type <> olOLE Then
        If objAttachment.Parent.ConversationTopic <> "" Then
            strSubDir = objAttachment.Parent.ConversationTopic
        ElseIf objAttachment.Parent.Subject <> "" Then
            strSubDir = objAttachment.Parent.Subject
        Else
            strSubDir = "Uncategorized"
        End If
    
        strSubDir = MakeLegalFileName(strSubDir, True) & "\"
        bBool = CheckPath(strBaseDir & strSubDir, True)
        strSavedFile = strBaseDir & strSubDir & MakeLegalFileName(objAttachment.FileName, True)
        objAttachment.SaveAsFile strSavedFile
    End If

    SaveAttachment = strSavedFile

End Function

' Replaces illegal filename characters with an underscore
' Optionally replaces spaces with underscores as well
Private Function MakeLegalFileName(ByVal strFileNameIn As String, Optional ByVal bReplaceSpace As Boolean = False) As String
    Dim i As Integer
    Dim strIllegals As String
    
    If bReplaceSpace Then
        strIllegals = "\/|?*<>"": "
    Else
        strIllegals = "\/|?*<>"":"
    End If
    
    ' Regardless we have to trim leading and trailing spaces!
    MakeLegalFileName = Trim(strFileNameIn)
    For i = 1 To Len(strIllegals)
        MakeLegalFileName = Replace(MakeLegalFileName, Mid$(strIllegals, i, 1), "_")
    Next i
End Function

' Checks the string expected to be in the form of a Windows directory path
' and returns true or false if the path exists.
'
' If the optional CreatePath flag is set to True (default = False) then
' the path is created if it doesn't already exist and the function returns True.
'
' TODO: Check the input string for correct formatting.
'
Private Function CheckPath(ByRef FolderPath As String, _
    Optional ByVal CreatePath As Boolean = False) As Boolean
    
    Dim nodes() As String
    Dim strDirectory As String
    Dim i As Integer
    
    ' First check if the requested directory exists
    If Len(Dir(FolderPath, vbDirectory)) > 0 Then
        CheckPath = True
    Else
        CheckPath = False
    End If
    
    If Not CheckPath And CreatePath Then
         'If they supplied an ending path seperator, cut it for now
        If Right$(FolderPath, 1) = Chr(92) Then
            FolderPath = Left$(FolderPath, Len(FolderPath) - 1)
        End If
        
        nodes = Split(FolderPath, Chr(92))
        strDirectory = nodes(0)
        For i = 1 To UBound(nodes)
            strDirectory = strDirectory & Chr(92) & nodes(i)
            If Len(Dir(strDirectory, vbDirectory)) = 0 Then MkDir strDirectory
        Next i
        If Len(Dir(strDirectory, vbDirectory)) > 0 Then
            ' Add back the trailing backslash
            FolderPath = FolderPath & Chr(92)
            CheckPath = True
        End If
    End If
    
End Function

' When an attachment is saved to disk then annotate the message to provide
' a link to the saved file.
Private Sub AnnotateMessage(ByVal olMsg As Outlook.MailItem, ByVal strSavedFile As String)
    Dim strSavedFileLink As String

    Select Case olMsg.BodyFormat
        Case olFormatHTML
            strSavedFileLink = "<a href='file://" & strSavedFile & "'>" & strSavedFile & "</a>"
            olMsg.HTMLBody = "<p>" & "Attachment saved to " & strSavedFileLink & "</p>" & olMsg.HTMLBody
        Case olFormatRichText
            strSavedFileLink = vbCrLf & "{\field{\*\fldinst HYPERLINK ""file://" & strSavedFile & _
                            """}{\fldrslt " & strSavedFile & "}}"
            olMsg.RTFBody = "Attachment saved to " & strSavedFileLink & vbCrLf & olMsg.RTFBody
        Case Else
            strSavedFileLink = vbCrLf & "<file://" & strSavedFile & ">"
            olMsg.Body = vbCrLf & "Attachment saved to " & strSavedFileLink & vbCrLf & olMsg.Body
    End Select

    olMsg.Save

End Sub
