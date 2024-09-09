Sub sendFromOutbox()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.NameSpace
    Dim olOutbox As Outlook.MAPIFolder
    Dim olInbox As Outlook.MAPIFolder
    Dim olMail As Outlook.MailItem
    Dim i As Integer
    Dim isOutboxInView As Boolean
    
    ' Initialize Outlook Application object
    Set olApp = Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
    
    ' Get the Outbox and Inbox folders
    Set olOutbox = olNamespace.GetDefaultFolder(olFolderOutbox)
    Set olInbox = olNamespace.GetDefaultFolder(olFolderInbox)
    
    ' Check if Outbox is currently in view
    isOutboxInView = (olApp.ActiveExplorer.CurrentFolder.EntryID = olOutbox.EntryID)
    
    ' If Outbox is in view, switch to Inbox
    If isOutboxInView Then
        Set olApp.ActiveExplorer.CurrentFolder = olInbox
    End If
    
    ' Loop through each item in the Outbox from last to first
    For i = olOutbox.Items.Count To 1 Step -1
        If TypeOf olOutbox.Items(i) Is MailItem Then
            Set olMail = olOutbox.Items(i)
            
            ' Save the mail and then send it
            olMail.Save
            olMail.Send
        End If
    Next i
    
    ' Cleanup
    Set olMail = Nothing
    Set olOutbox = Nothing
    Set olInbox = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
End Sub
