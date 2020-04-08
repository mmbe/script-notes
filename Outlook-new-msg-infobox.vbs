Option Explicit
Private WithEvents Items As Outlook.Items

Private Sub Application_Startup()
    Dim olNs As Outlook.NameSpace
    Dim Inbox  As Outlook.MAPIFolder
    Dim olRecip As Recipient
 
    Set olNs = Application.GetNamespace("MAPI")
    Set olRecip = olNs.CreateRecipient("emailacc@email.com")
    Set Inbox = olNs.GetSharedDefaultFolder(olRecip, olFolderInbox) 'InBOX,to add other folder inside Inbox follow with .Folders("")
    Set Items = Inbox.Items
End Sub

Private Sub Items_ItemAdd(ByVal Item As Object)
    Dim Msg As Outlook.MailItem
    Dim msgTxt
    Dim retVal
    Dim cmd
        
    If TypeOf Item Is Outlook.MailItem Then
        'popup CMD window:
        'msgTxt = "CMD /k echo " & Date & " " & Time() & "  SENDER: " & Item.Sender & "     SUBJECT: " & Item.Subject
        'retVal = Shell(msgTxt, vbNormalFocus)
        ' popup box above everything else:
        msgTxt = Date & " " & Time() & vbCrLf & _
        "SENDER:     " & Item.Sender & vbCrLf & _
        "SUBJECT:    " & Item.Subject
        retVal = MsgBox(msgTxt, 4160, "You have new message!")
    End If
End Sub
