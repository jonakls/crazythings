Sub SendMails()
    
    Dim MailObject As Object
    ' Interator of mails in excel tabble
    For i = 2 To 4
        Set MailObject = CreateObject("Outlook.Application").CreateItem(0)
        
        MailObject.To = Cells(i, 1).Value ' Mail receptor
        MailObject.Subject = Cells(i, 2).Value ' Mail subject
        MailObject.Body = Cells(i, 3).Value ' Mail message
        MailObject.Archive = Cells(i, 4).Value ' Mail archive
        MailObject.Archive = Cells(i, 5).Value ' Mail archive2
        
        MailObject.Send ' Send mail
        MsgBox "Mail sent", vbOKOnly, "Successfully sent"
    Next
End Sub