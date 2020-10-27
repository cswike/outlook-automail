Public WithEvents objReminders As Outlook.Reminders

Sub AutoMailDoc()
'this is configured to send to specific people, see lines 22-24
    Dim olApp As Outlook.Application
    Dim olMail As Outlook.MailItem
    Dim blRunning As Boolean   
     'get application
    blRunning = True
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then
        Set olApp = New Outlook.Application
        blRunning = False
    End If
    On Error GoTo 0
    
    Set olMail = olApp.CreateItem(olMailItem)
    With olMail
         'Specify the email subject
        .Subject = "EMAIL SUBJECT GOES HERE " & Date
         'Specify who it should be sent to
         'Repeat this line to add further recipients
         '***IMPORTANT: If email addresses change and/or people leave, below here is where you need to change the email addresses or delete line(s).***
        .Recipients.Add "RECIPIENT EMAIL HERE"
        .Recipients.Add "RECIPIENT 2 EMAIL HERE"
         'specify the file to attach
         'repeat this line to add further attachments
        .Attachments.Add "PATH TO FILE ATTACHMENT HERE"
         'specify the text to appear in the email
        .Body = "BODY OF EMAIL GOES HERE"
         'Choose which of the following 2 lines to have commented out
         '.Display 'This will display the message for you to check and send yourself
        .Send 'This will send the message straight away
    End With
    
    If Not blRunning Then olApp.Quit
    
    Set olApp = Nothing
    Set olMail = Nothing
    
End Sub

Private Sub Application_Startup()
    Set objReminders = Outlook.Application.Reminders
End Sub

'When a Reminder Pops up
Private Sub objReminders_ReminderFire(ByVal ReminderObject As Reminder)
    Dim objTask As Outlook.TaskItem
   
    'If It's a Task's Reminder
    If TypeOf ReminderObject.Item Is TaskItem Then
        If ReminderObject = "TASK NAME HERE" Then
            Set objTask = ReminderObject.Item
            'After 0 seconds
            Wait (0)
            'Mark Task Complete
            objTask.Complete = True
            objTask.Save
           
            'Call AutoMailDoc to send out email & attachment
            AutoMailDoc
        End If
    End If
End Sub
 
Function Wait(nSeconds As Integer) As Boolean
    Dim dCurrentTime As Date
    dCurrentTime = Now
    Do Until DateAdd("s", nSeconds, dCurrentTime) <= Now
       DoEvents
    Loop
End Function
