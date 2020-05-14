Attribute VB_Name = "OutlookMacrosSW"
Public Sub SetCustomFlag()

Dim myOlExp As Outlook.Explorer
Dim myOlSel As Outlook.Selection
Dim oMail As Outlook.MailItem
Dim x As Long
Dim FollowDays As Long
Dim DayoWeek As Long
Dim Input_Box_switch
Dim Message, Title, Default, MyValue

' Turn input box on/off (on = 1)
Input_Box_switch = 0

DayoWeek = Weekday(Date)

Message = "<Holiday Warning> How many days to the due date?"     ' Set prompt.
Title = "Multi-msg Custom Flag"    ' Set title.
If DayoWeek = 5 Or DayoWeek = 6 Then
    Default = 4    ' Set default.
Else
    Default = 2    ' Set default.
End If

If Input_Box_switch = 1 Then
    FollowDays_Text = InputBox(Message, Title, Default)
    If FollowDays_Text = "" Then Exit Sub
    
    If IsNumeric(FollowDays_Text) Then
        FollowDays = CLng(FollowDays_Text)
    Else
        FollowDays = 1
    End If
Else
    FollowDays = Default
End If

Set myOlExp = Application.ActiveExplorer
Set myOlSel = myOlExp.Selection
For x = 1 To myOlSel.Count
    With myOlSel.item(x)
        .MarkAsTask (1)
        ' sets a specific due date
        .TaskDueDate = Date + FollowDays
        .FlagRequest = "Follow up"
        .Save
    End With
Next x

Set objMsg = Nothing
End Sub

Function FindNewestFile(folder_dir As String, file_ext As String) As String
    file_find = Dir(folder_dir & file_ext)
    
    latest_file = ""
    latest_fileDate = 0
    Do While Len(file_find) > 0
        file_curr = folder_dir & file_find
        fileDate = DateValue(FileDateTime(file_curr))
        If fileDate > latest_fileDate Then
            latest_file = file_curr
            latest_fileDate = fileDate
        End If
        file_find = Dir
    Loop
    FindNewestFile = latest_file
End Function

Function OutlookExchangeGreeting(ExchanjUser)
    Select Case ExchanjUser.Firstname
        Case "jonathan"
            OutlookExchangeGreeting = "Hi John"
        Case "Yoshihiro"
            OutlookExchangeGreeting = "Yoshi-san"
        Case Else
            OutlookExchangeGreeting = "Hi " & ExchanjUser.Firstname
    End Select
End Function

Sub ReplyWithGreeting()
    'Inspired by the following URL, modified by Shota Watanabe. https://www.extendoffice.com/documents/outlook/2983-outlook-auto-greeting.html#vba
    Dim oMItem As Outlook.MailItem
    Dim oMItemReply As Outlook.MailItem
    Dim oMSender As Outlook.AddressEntry
    Dim oMContact As Outlook.ContactItem
    Dim oMContactR As Outlook.ContactItem
    Dim oMContacts As Outlook.Items
    Dim sGreetName As String
    Dim sGreetNames() As String
    Dim sGreetTime As String

    On Error Resume Next
    Select Case TypeName(Application.ActiveWindow)
        Case "Explorer"
            Set oMItem = ActiveExplorer.Selection.item(1)
        Case "Inspector"
            Set oMItem = ActiveInspector.CurrentItem
        Case Else
    End Select
    On Error GoTo 0
    If oMItem Is Nothing Then GoTo ExitProc
    On Error Resume Next
    
    Set oMSender = oMItem.Sender
    Set oMContact = oMSender.GetContact
    Set oMExContact = oMSender.GetExchangeUser
    
    If (Not oMContact Is Nothing And Not oMItem.SenderName = "Shota Watanabe") Then
        If (Not oMContact.NickName = "") Then
            sGreetName = oMContact.NickName
        Else
            sGreetName = "Hi " & oMContact.Firstname
        End If
    ElseIf (Not oMExContact Is Nothing) Then
        
        'If I am the sender, get the recipient name
        If (oMExContact.Firstname = "Shota") Then
           ' this code is from AddGreeting below
            oMItem.Recipients.ResolveAll
            
            Set oMRecipient = oMItem.Recipients(1)
            Set oMContactR = oMRecipient.AddressEntry.GetContact
            Set oMRecipientEx = oMRecipient.AddressEntry.GetExchangeUser
            
            If (Not oMContactR Is Nothing) Then
                If (Not oMContactR.NickName = "") Then
                    sGreetName = oMContactR.NickName
                Else
                    sGreetName = "Hi " & oMContactR.Firstname
                End If
            ElseIf (Not oMRecipientEx Is Nothing) Then
                sGreetName = OutlookExchangeGreeting(oMRecipientEx)
            Else
                sGreetNames() = Split(oMRecipient.Name, " ")
                If InStr(1, oMRecipient.Name, ",", 1) > 0 Then
                    sGreetName = "Hi " & sGreetNames(1)
                Else
                    sGreetName = "Hi " & sGreetNames(0)
                End If
            End If
        Else
            sGreetName = OutlookExchangeGreeting(oMExContact)
        End If
    Else
        sGreetNames() = Split(oMItem.SenderName, " ")
        If InStr(1, oMItem.SenderName, ",", 1) > 0 Then
            sGreetName = "Hi " & sGreetNames(1)
        Else
            sGreetName = "Hi " & sGreetNames(0)
        End If
    End If
    
    Select Case Time
        Case Is < 0.5
            sGreetTime = "Good Morning"
        Case 0.5 To 0.75
            sGreetTime = "Good Afternoon"
        Case Else
            sGreetTime = "Good Day"
    End Select
    Set oMItemReply = oMItem.ReplyAll
    With oMItemReply
        .Display
        .HTMLBody = "<body style=""font-family : Calibri; font-size : 11.0pt"">" & sGreetName & ",<br><br>" & sGreetTime & ",</body>" & .HTMLBody
        If (oMItem.BodyFormat = 1) Then .BodyFormat = olFormatPlain
    End With
    oMItem.UnRead = False
ExitProc:
    Set oMItem = Nothing
    Set oMItemReply = Nothing
End Sub

Sub AddGreeting()
    'Inspired by the following URL, modified by Shota Watanabe. https://www.extendoffice.com/documents/outlook/2983-outlook-auto-greeting.html#vba
    Dim oMItem As Outlook.MailItem
    Dim oMContact As Outlook.ContactItem
    Dim oMContacts As Outlook.Items
    Dim sGreetName As String
    Dim sGreetNames() As String
    Dim sGreetTime As String
    Dim result As String
    
    On Error Resume Next
    Set oMItem = Application.ActiveInspector.CurrentItem
    oMItem.Recipients.ResolveAll
    
    On Error GoTo 0
    If oMItem Is Nothing Then GoTo ExitProc
    On Error Resume Next
    
    Set oMRecipient = oMItem.Recipients(1)
    Set oMContact = oMRecipient.AddressEntry.GetContact
    Set oMRecipientEx = oMRecipient.AddressEntry.GetExchangeUser
    
    'Set oMContacts = Application.Session.GetDefaultFolder(olFolderContacts).Items
    'https://www.datanumen.com/blogs/quickly-insert-recipient-names-email-body-outlook/
    'For i = 1 To 3
     '   strFilter = "[Email" & i & "Address] = " & oMRecipient.Address
      '  Set oMContact = oMContacts.Find(strFilter)
       ' If Not (oMContact Is Nothing) Then
        '    If (Not oMContact.NickName = "") Then
         '   sGreetName = oMContact.NickName
          '  Else
           '     sGreetName = "Hi " & oMContact.Firstname
           ' End If
          '  Exit For
       ' End If
    'Next
    
    If (Not oMContact Is Nothing) Then
        If (Not oMContact.NickName = "") Then
        sGreetName = oMContact.NickName
        Else
            sGreetName = "Hi " & oMContact.Firstname
        End If
    ElseIf (Not oMRecipientEx Is Nothing) Then
        sGreetName = OutlookExchangeGreeting(oMRecipientEx)
    Else
        sGreetNames() = Split(oMRecipient.Name, " ")
        If InStr(1, oMRecipient.Name, ",", 1) > 0 Then
            sGreetName = "Hi " & sGreetNames(1)
        Else
            sGreetName = "Hi " & sGreetNames(0)
        End If
    End If
    
    Select Case Time
        Case Is < 0.5
            sGreetTime = "Good Morning"
        Case 0.5 To 0.75
            sGreetTime = "Good Afternoon"
        Case Else
            sGreetTime = "Good Day"
    End Select
    
    With oMItem
        .Display
        .HTMLBody = "<body style=""font-family : Calibri; font-size : 11.0pt"">" & sGreetName & ",<br><br>" & sGreetTime & ",</body>" & .HTMLBody
        If (oMItem.BodyFormat = 1) Then .BodyFormat = olFormatPlain
    End With
    
ExitProc:
    Set oMItem = Nothing

End Sub

Sub AddGreetingName()
    'Inspired by the following URL, modified by Shota Watanabe. https://www.extendoffice.com/documents/outlook/2983-outlook-auto-greeting.html#vba
    Dim oMItem As Outlook.MailItem
    Dim oMContact As Outlook.ContactItem
    Dim oMContacts As Outlook.Items
    Dim sGreetName As String
    Dim sGreetNames() As String
    Dim sGreetTime As String
    Dim result As String
    
    On Error Resume Next
    Set oMItem = Application.ActiveInspector.CurrentItem
    oMItem.Recipients.ResolveAll
    
    On Error GoTo 0
    If oMItem Is Nothing Then GoTo ExitProc
    On Error Resume Next
    
    Set oMRecipient = oMItem.Recipients(1)
    Set oMContact = oMRecipient.AddressEntry.GetContact
    Set oMRecipientEx = oMRecipient.AddressEntry.GetExchangeUser
    
    'Set oMContacts = Application.Session.GetDefaultFolder(olFolderContacts).Items
    'https://www.datanumen.com/blogs/quickly-insert-recipient-names-email-body-outlook/
    'For i = 1 To 3
     '   strFilter = "[Email" & i & "Address] = " & oMRecipient.Address
      '  Set oMContact = oMContacts.Find(strFilter)
       ' If Not (oMContact Is Nothing) Then
        '    If (Not oMContact.NickName = "") Then
         '   sGreetName = oMContact.NickName
          '  Else
           '     sGreetName = "Hi " & oMContact.Firstname
           ' End If
          '  Exit For
       ' End If
    'Next
    
    If (Not oMContact Is Nothing) Then
        If (Not oMContact.NickName = "") Then
        sGreetName = oMContact.NickName
        Else
            sGreetName = "Hi " & oMContact.Firstname
        End If
    ElseIf (Not oMRecipientEx Is Nothing) Then
        sGreetName = OutlookExchangeGreeting(oMRecipientEx)
    Else
        sGreetNames() = Split(oMRecipient.Name, " ")
        If InStr(1, oMRecipient.Name, ",", 1) > 0 Then
            sGreetName = "Hi " & sGreetNames(1)
        Else
            sGreetName = "Hi " & sGreetNames(0)
        End If
    End If
    
    With oMItem
        .Display
        .HTMLBody = "<body style=""font-family : Calibri; font-size : 11.0pt"">" & sGreetName & ",<br><br></body>" & .HTMLBody
        If (oMItem.BodyFormat = 1) Then .BodyFormat = olFormatPlain
    End With
    
ExitProc:
    Set oMItem = Nothing

End Sub

Sub NewRFQ()
    Dim MyItem As Outlook.MailItem
    Set MyItem = Application.CreateItemFromTemplate("C:\Users\shota\AppData\Roaming\Microsoft\Templates\RFQTemplate.oft")
    MyItem.Display
End Sub

Sub ChangeSelectionFont()
' https://www.reddit.com/r/vba/comments/4dmj3t/outlook_vba_macro_to_change_font_namesize_of_text/d1sbr74/
' By: pmo86

    Dim item As MailItem
    Set item = Application.ActiveInspector.CurrentItem
    
    Set objDoc = item.GetInspector.WordEditor
    
    With objDoc.Application.Selection.Font
        .Name = "Calibri"
        .Size = 11
    End With
End Sub
