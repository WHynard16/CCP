VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim Department As String
Dim NxtDept As String
Dim NxtTask As String
Dim Recipient As String

Sub SendEmail()
Dim Mail As New Message
Dim config As Configuration
Set config = Mail.Configuration

config(cdoSendUsingMethod) = cdoSendUsingPort
config(cdoSMTPServer) = "smtp.gmail.com"
config(cdoSMTPServerPort) = 25
config(cdoSMTPAuthenticate) = cdoBasic
config(cdoSMTPUseSSL) = True
config(cdoSendUserName) = "william.hynard@gmail.com"
config(cdoSendPassword) = "Vulcan558"
config.Fields.Update

Mail.To = Recipient
Mail.From = config(cdoSendUserName)
Mail.Subject = "Customer Care Process Notification"
Mail.HTMLBody = "<b>" & Department & " has completed the previous step.<br><br>Next Task Owner: " & NxtDept & "<br>Next Task: " & NxtTask & "</b>"

On Error Resume Next

Mail.Send

If Err.Number <> 0 Then
MsgBox Err.Description, vbCritical, "There was an Error!"
Exit Sub
End If

MsgBox "Notification has been sent to the next task owner.", vbinfomation, "Sent"

End Sub
Private Sub CheckBox1_Click()
If CheckBox1.Value = True And Sheets("SALES").CheckBox3.Value = False Then
MsgBox ("Previous step has not been completed." & vbNewLine & "Sales step required")
CheckBox1.Value = False
End If

If CheckBox1.Value = True And Range("F2").Value <> True Then
Range("F2").Value = True
Label1.Caption = Date + Time
CheckBox1.Enabled = False

Department = "Finance & Accounts"
NxtDept = "Sales"
NxtTask = Sheets(NxtDept).Range("B5").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If

End Sub

Private Sub CheckBox2_Click()
If CheckBox2.Value = True And Sheets("SALES").CheckBox5.Value = False And CheckBox1.Value = False Then
MsgBox ("Previous step has not been completed." & vbNewLine & "Sales step required" & vbNewLine & "Previous Department Step is Incomplete.")
CheckBox2.Value = False
Else

    If CheckBox2.Value = True And Sheets("SALES").CheckBox5.Value = False Then
    MsgBox ("Previous step has not been completed." & vbNewLine & "Sales step required")
    CheckBox2.Value = False
    Else
    
        If CheckBox2.Value = True And CheckBox1.Value = False Then
        MsgBox ("Previous Department Step is Incomplete.")
        CheckBox2.Value = False
        End If
    End If
End If

If CheckBox2.Value = True And Range("F3").Value <> True Then
Range("F3").Value = True
Label2.Caption = Date + Time
CheckBox2.Enabled = False

Department = "Finance & Accounts"
NxtDept = "Sales"
NxtTask = Sheets(NxtDept).Range("B7").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If

End Sub

Private Sub CheckBox4_Click()
If CheckBox4.Value = True And Sheets("SALES").CheckBox9.Value = False And CheckBox2.Value = False Then
MsgBox ("Previous step has not been completed." & vbNewLine & "Sales step required" & vbNewLine & "Previous Department Step is Incomplete.")
CheckBox4.Value = False
Else

    If CheckBox4.Value = True And Sheets("SALES").CheckBox9.Value = False Then
    MsgBox ("Previous step has not been completed." & vbNewLine & "Sales step required")
    CheckBox4.Value = False
    Else
    
        If CheckBox4.Value = True And CheckBox2.Value = False Then
        MsgBox ("Previous Department Step is Incomplete.")
        CheckBox4.Value = False
        End If
    End If
End If

If CheckBox4.Value = True And Range("F4").Value <> True Then
Range("F4").Value = True
Label4.Caption = Date + Time
CheckBox4.Enabled = False

Department = "Finance & Accounts"
NxtDept = "Sales"
NxtTask = Sheets(NxtDept).Range("B11").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If

End Sub

Private Sub CheckBox5_Click()
If CheckBox5.Value = True And Sheets("SALES").CheckBox11.Value = False And CheckBox4.Value = False Then
MsgBox ("Previous step has not been completed." & vbNewLine & "Sales step required" & vbNewLine & "Previous Department Step is Incomplete.")
CheckBox5.Value = False
Else

    If CheckBox5.Value = True And Sheets("SALES").CheckBox11.Value = False Then
    MsgBox ("Previous step has not been completed." & vbNewLine & "Sales step required")
    CheckBox5.Value = False
    Else
    
        If CheckBox5.Value = True And CheckBox4.Value = False Then
        MsgBox ("Previous Department Step is Incomplete.")
        CheckBox5.Value = False
        End If
    End If
End If

If CheckBox5.Value = True And Range("F5").Value <> True Then
Range("F5").Value = True
Label5.Caption = Date + Time
CheckBox5.Enabled = False

Department = "Finance & Accounts"
NxtDept = "Sales"
NxtTask = Sheets(NxtDept).Range("B13").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If

End Sub

Private Sub CheckBox6_Click()
If CheckBox6.Value = True And Sheets("SERVICE").CheckBox5.Value = False And CheckBox1.Value = False Then
MsgBox ("Previous step has not been completed." & vbNewLine & "Service step required" & vbNewLine & "Previous Department Step is Incomplete.")
CheckBox6.Value = False
Else

    If CheckBox6.Value = True And Sheets("SERVICE").CheckBox5.Value = False Then
    MsgBox ("Previous step has not been completed." & vbNewLine & "Service step required")
    CheckBox6.Value = False
    Else
    
        If CheckBox6.Value = True And CheckBox1.Value = False Then
        MsgBox ("Previous Department Step is Incomplete.")
        CheckBox6.Value = False
        End If
    End If
End If

If CheckBox6.Value = True And Range("F6").Value <> True Then
Range("F6").Value = True
Label6.Caption = Date + Time
CheckBox6.Enabled = False

Department = "Finance & Accounts"
NxtDept = "Service"
NxtTask = Sheets(NxtDept).Range("B6").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If

End Sub
