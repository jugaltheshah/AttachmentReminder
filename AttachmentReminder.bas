' This code checks each outgoing message to see if the user meant to attach something but forgot '
' Allows user to cancel sending if they did really mean to attach something, or continue otherwise '
' Will stop checking if it hits a line with "From: " - this is to minimize false positives if you're just replying to somebody that sent an attachment
' Usage - add to the ThisOutlookSession in the code editor in Outlook. Should fire on all emails sent thereafter
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
 
    Dim Message() As String
    Dim Catchwords() As Variant
    Dim Catchsubjects() As Variant
    Dim possibleAttachment As Boolean
   
    Const SEARCHUNTIL = "From: " ' Search until this phrase if the message is a reply or forward
   
    Message = Split(Item.Body, vbLf)
   
    ' These are the words to catch - more can be added using the syntax below
    Catchwords = Array("attach", "attached", "attaching", "attachments", "enclosed")
   
    ' Add any subjects that need catching here (in case there's no words that indicate an attachment, but the subject implies one)
    ' Eg if an email goes out regularly that should always have an attachment, but it's not stated in the body, we can still catch it
    Catchsubjects = Array("Test", "Testing")
   
    On Error GoTo handleError
   
    possibleAttachment = False
           
    ' Only check email if there's no attachments - must be modified if email signature includes a picture, as that counts as an attachment!
    If Item.Attachments.Count = 0 Then
        ' If the item's subject is one of the flagged ones then catch the missing attachment
	' Code to be added
        For Each Line In Message
            If InStr(Line, SEARCHUNTIL) = 0 And possibleAttachment = False Then
                For Each word In Catchwords
                    ' This will catch the keywords but not anything with a keyword followed by a question mark
                    ' The question mark bit is meant to weed out false positives - eg "can you send me the attachment?"
                    If (InStr(LCase$(Line), word) <> 0) And (InStr(LCase$(Line), word & "?") = 0) Then
                        possibleAttachment = True
                        Exit For ' Optimization to speed things up - since we only need one possible attachment, exit once found
                    End If
                Next
            Else
                Exit For ' We've hit a boundary, most likely below this line is the message being replied to / forwarded
            End If
        Next
    End If
   
    ' This is what pops up the message box - customize text here
    If possibleAttachment Then
        SendWithoutAttachment = MsgBox("Send message without attachment?", vbQuestion + vbYesNo + vbMsgBoxSetForeground + vbDefaultButton2)
        If Not SendWithoutAttachment = vbYes Then
            Cancel = True
        End If
    End If
   
handleError:
    If Err.Number <> 0 Then
        MsgBox "Outlook Attachment Reminder Error: " & Err.Description, vbExclamation, "Outlook Attachment Reminder Error"
    End If
 
End Sub