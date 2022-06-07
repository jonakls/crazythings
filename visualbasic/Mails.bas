Sub SendMail()

    Dim MailObject As Object
    Dim Value As Integer
    Dim CellValue As Variant
    
    CellValue = Cells(6, 1).Value
    Value = 6
    Set MailObject = CreateObject("Outlook.Application").CreateItem(0)
    
    Do While CellValue <> Empty
        CellValue = Cells(Value, 1).Value
        Value = Value + 1
    Loop
    
    For I = 6 To Value
        
        
        If Cells(I, 1).Value = "" Then GoTo Cycle:
        
        MailObject.To = Cells(I, 1).Value
        MailObject.Subject = Cells(I, 2).Value
        MailObject.Body = Cells(I, 3).Value
        
        If Not Cells(I, 4).Value = "" Then
            MailObject.Attachments.Add Cells(I, 4).Value
        ElseIf Not Cells(I, 5) = "" Then
            MailObject.Attachments.Add Cells(I, 5).Value
        ElseIf Not Cells(I, 6).Value Then
            MailObject.Attachments.Add Cells(I, 6).Value
        End If
        
        MailObject.Send
Cycle:
    Next
End Sub