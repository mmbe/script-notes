Sub EndUser()
    Dim Reply As Outlook.MailItem
    Dim objContent As Outlook.MailItem
    Dim objRegexp As VBScript_RegExp_55.RegExp
    Dim colRecip As VBScript_RegExp_55.MatchCollection
    Dim strAddress As String
    
    Set objRegexp = New VBScript_RegExp_55.RegExp
    
    'please set your path to mail template:
    Set objContent = Application.CreateItemFromTemplate("C:\Users\User\AppData\Roaming\Microsoft\Templates\file.oft")
    
    On Error Resume Next
    
    For Each M In Application.ActiveExplorer.Selection
    Set Reply = M.Reply
    If Reply.To = "noreply@email.com" Then
    With objRegexp
    .IgnoreCase = True
    .Global = True
    .Pattern = "(([\w-\.]*\@[\w-\.]*)\s*)"
        Set colRecip = objRegexp.Execute(Reply.Body)
    End With
    Reply.To = colRecip.Item(3)
    End If
    With Reply
        .CC = "ccemail@email.com"
        '.Subject = "Subject line"
        '.Save
        .Display
        .HTMLBody = objContent.HTMLBody & Reply.HTMLBody
        '.Send
    End With
    Next
    Set objContent = Nothing
    Set objRegexp = Nothing
    Set colRecip = Nothing
End Sub
