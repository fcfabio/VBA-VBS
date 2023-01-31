
Function ScheduleOutlookTask(recipient, subject, body, startDate, reminderSet)
    Set objOApp = CreateObject("Outlook.Application") 'New Outlook.Application
    Set objAppt = objOApp.CreateItem(olTaskItem)
    With objAppt
        .Assign
        .Recipients.Add recipient
        .Subject = subject
        .StartDate = CDate(startDate)
        .Body = body
        .ReminderSet = reminderSet
        .Save
        .send
    End With
End Function


Function SendOutlookEmail(recipient, subject, body, cc, bcc, priority, attachments)

  Set OutApp = CreateObject("Outlook.Application")
  Set OutMail = OutApp.CreateItem(0)

  With OutMail
    .To = recipient
    .cc = cc
    .bcc = bcc
    .subject = subject
    .HTMLBody = body
    .Importance = priority
    If Not IsNull(attachments) Or attachments = "" Then
        For Each attachmentPath In Split(attachments, ";")
            .attachments.Add Trim(attachmentPath)
        Next
    End If
    .send
  End With

  Set OutMail = Nothing
  Set OutApp = Nothing

End Function

