Attribute VB_Name = "Email"

Option Explicit

Public Sub RegisterProduct(Address, UserName, ProductNo, SendStatus)

On Error GoTo Err_Handler

Dim Session As Control, Message As Control
Dim sbj As String 'PCN1916
Dim bdy As String

sbj = "Profiler Product Registration Request." 'PCN3423
bdy = "Please Send Me the Registration Code for My Profiler V6 Software, User Name:" & UserName & ", Product No :" & ProductNo & "." 'PCN3423

'MAPI Session & Message Control programming starts -------------
Set Session = Registration.MAPISession1
Set Message = Registration.MAPIMessages1

Session.SignOn

Message.SessionID = Session.SessionID
Message.Compose
Message.RecipAddress = Address
Message.MsgSubject = sbj
Message.MsgNoteText = bdy

Message.AddressResolveUI = True
Message.ResolveName
Message.send False

Session.SignOff
'MAPI Session & Message Control programming ends ------

SendStatus = True

Exit Sub
Err_Handler:
    Select Case Err
    Case 94
        'MsgBox DisplayMessage("Subject, Body text, or Attachment is not specified."), vbExclamation 'PCN2111
        ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Subject, Body text, or Attachment is not specified."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
        SendStatus = False
    Case 32050
        Session.SignOff
        Resume
    Case 32002
       'MsgBox DisplayMessage("Invalid Path @@Please verify the Path or EMail Address."), vbExclamation 'PCN2111
        ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Invalid Path @@Please verify the Path or EMail Address."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
        SendStatus = False
    Case 32014
        'MsgBox DisplayMessage("Unknown Recipient."), vbExclamation 'PCN2111
        ProfilerMessageBox.ProfilerMsgBoxLbl.Caption = DisplayMessage("Unknown Recipient."): ProfilerMessageBox.Show vbModal: ProfilerMessageBox.ZOrder 0
        SendStatus = False
    Case Else
        MsgBox Err & "-E1" & Error$, vbExclamation 'PCN2111
        SendStatus = False
    End Select
End Sub
