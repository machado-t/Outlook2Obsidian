Attribute VB_Name = "SaveEmail"

Option Explicit
'======================================================================================='

' Declare ShellExecute API for opening Obsidian links
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Sub ExtractEmail()

    Dim vaultPathToSaveFileTo As String
    Dim emailFileNameStartChr As String
    Dim emailTypeLink As String
    Dim personNameStartChar As String
    config vaultPathToSaveFileTo, personNameStartChar, emailFileNameStartChr, emailTypeLink

    '================================================'
    ' Save as plain text
    Const OLTXT = 0
    ' Object holding variable
    Dim obj As Object
    ' Instantiate an Outlook Email Object
    Dim oMail As Outlook.MailItem
    ' If something breaks, skip to the end, tidy up and shut the door
    On Error GoTo EndClean:
    ' Establish the environment and selected items (emails)
    ' NOTE: selecting a conversation-view stack wont work
    '       you'll need to select one of the emails
    Dim fileName As String
    Dim temporarySubjectLineString As String
    Dim currentExplorer As Explorer
        Set currentExplorer = Application.ActiveExplorer
    Dim Selection As Selection
        Set Selection = currentExplorer.Selection
    ' For each email in the Selection
    ' Assigning email item to the `obj` holding variable
    For Each obj In Selection
        ' set the oMail object equal to that mail item
        Set oMail = obj
        ' Is it an Email?
        If oMail.Class <> 43 Then
          MsgBox "This code only works with Emails."
          GoTo EndClean: ' you broke it
        End If

        ' Yank the mail items subject line to `temporarySubjectLineString`
        temporarySubjectLineString = oMail.Subject
        ' function call the name cleaner to remove any
        '    illegal characters from the subject line
        ReplaceCharsForFileName temporarySubjectLineString, ""
        ' Yank the received date-time to a holding variable

        ' Build Recipient string based on receipient collection
        Dim recips As Outlook.Recipients
            Set recips = oMail.Recipients
        Dim recip As Outlook.Recipient
        Dim result As String
        Dim recipString As String
            recipString = ""

        For Each recip In recips
            recipString = recipString & vbTab
            recipString = recipString & "- "
            recipString = recipString & formatName(recip.Name, personNameStartChar)
            recipString = recipString & vbCrLf
        Next
        ' Build the result file content to be sent to the mail item body
        ' Then save that mail item same as the meeting extractor
        Dim sender As String
            sender = formatName(oMail.sender, personNameStartChar)
        Dim dtDate As Date
            dtDate = oMail.ReceivedTime
        Dim resultString As String

        ' resultString = ""
        ' resultString = resultString & "# [[" & emailFileNameStartChr & Format(oMail.ReceivedTime, "yyyy-mm-dd hhnn") & " " & temporarySubjectLineString & "|" & temporarySubjectLineString & "]]"
        ' resultString = resultString & vbCrLf & vbCrLf & vbCrLf
        
        ' Start YAML frontmatter
        resultString = "---" & vbCrLf
        
        ' Add classification and optional properties
        resultString = resultString & "class: email" & vbCrLf
        resultString = resultString & "area: " & vbCrLf
        resultString = resultString & "project: " & vbCrLf
        resultString = resultString & "title: """ & temporarySubjectLineString & """" & vbCrLf
        resultString = resultString & "date: " & Format(oMail.ReceivedTime, "yyyy-MM-dd HH:mm") & vbCrLf
        resultString = resultString & "from: """ & sender & """" & vbCrLf

        ' Convert recipients to YAML list
        resultString = resultString & "to:" & vbCrLf
        For Each recip In recips
            resultString = resultString & "  - """ & formatName(recip.Name, personNameStartChar) & """" & vbCrLf
        Next

        ' Add tags
        resultString = resultString & "tags:" & vbCrLf

        ' Add related
        resultString = resultString & "related:" & vbCrLf
        
        ' Placeholder for attachments
        resultString = resultString & "attachments: []" & vbCrLf

        ' End YAML block
        resultString = resultString & "---" & vbCrLf & vbCrLf
        
        ' Add a Markdown task before the email body
        resultString = resultString & "- [ ] " & temporarySubjectLineString & vbCrLf & vbCrLf

        ' Add a horizontal rule for separation
        resultString = resultString & "---" & vbCrLf & vbCrLf
        
        ' Add original email content
        resultString = resultString & oMail.Body


        ' Make a dummy email to hold the details we're saving
        ' This way we dont get junk in the message header when saving
        Dim outputItem As MailItem
            Set outputItem = Application.CreateItem(olMailItem)
        outputItem.Body = resultString

        ' Now we create the file name
        fileName = emailFileNameStartChr
        fileName = fileName & Format(dtDate, "yyyy-mm-dd", vbUseSystemDayOfWeek, vbUseSystem)
        fileName = fileName & Format(dtDate, " hhMM", vbUseSystemDayOfWeek, vbUseSystem)
        fileName = fileName & " " & temporarySubjectLineString & ".md"

        ' Save the result
        ' outputItem.SaveAs vaultPathToSaveFileTo & fileName, OLTXT
        SaveAsUTF8 vaultPathToSaveFileTo & fileName, resultString
        
        ' ' Construct the Obsidian URI (must replace backslashes with `%5C` for Windows paths)
        ' Dim obsidianURI As String
        ' obsidianURI = "obsidian://open?path=" & Replace(vaultPathToSaveFileTo & fileName, "\", "%5C")
        
        ' ' Open the file in Obsidian
        ' ShellExecute 0, "open", obsidianURI, vbNullString, vbNullString, 1

        ' Fully encode the file path for Obsidian URI
        Dim obsidianURI As String
        ' obsidianURI = "obsidian://open?path=" & URLEncode(vaultPathToSaveFileTo & fileName)
        ' obsidianURI = "obsidian://open?path=" & UrlEncodeUtf8(vaultPathToSaveFileTo & fileName)
        obsidianURI = "obsidian://open?path=" & UrlEncodeUtf8NoBom(vaultPathToSaveFileTo & fileName)



        ' Use ShellExecute to open the note in Obsidian
        ShellExecute 0, "open", obsidianURI, vbNullString, vbNullString, 1


    Next
EndClean:
    Set obj = Nothing
    Set oMail = Nothing
    Set outputItem = Nothing
End Sub
