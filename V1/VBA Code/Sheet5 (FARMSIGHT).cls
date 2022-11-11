VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
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
If Sheets("SERVICE").CheckBox6.Value = False And CheckBox1.Value = True Then
MsgBox ("Previous Department Step is Incomplete." & vbNewLine & "Service step required")
CheckBox1.Value = False
End If

If CheckBox1.Value = True And Range("F2").Value <> True Then
Range("F2").Value = True
Label1.Caption = Date + Time
CheckBox1.Enabled = False

Department = "Farmsight"
NxtDept = "Sales"
NxtTask = Sheets(NxtDept).Range("B14").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If
End Sub

Private Sub CheckBox10_Click()
If CheckBox9.Value = False And CheckBox10.Value = True Then
MsgBox ("Previous step has not been completed.")
CheckBox10.Value = False
End If

If CheckBox10.Value = True And Range("F11").Value <> True Then
Range("F11").Value = True
Label10.Caption = Date + Time
CheckBox10.Enabled = False

Department = "Farmsight"
NxtDept = "Farmsight"
NxtTask = Sheets(NxtDept).Range("B12").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If

End Sub

Private Sub CheckBox11_Click()
If CheckBox10.Value = False And CheckBox11.Value = True Then
MsgBox ("Previous step has not been completed.")
CheckBox11.Value = False
End If

If CheckBox11.Value = True And Range("F12").Value <> True Then
Range("F12").Value = True
Label11.Caption = Date + Time
CheckBox11.Enabled = False

Department = "Farmsight"
NxtDept = "Farmsight"
NxtTask = Sheets(NxtDept).Range("B13").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If
End Sub

Private Sub CheckBox12_Click()
If CheckBox11.Value = False And CheckBox12.Value = True Then
MsgBox ("Previous step has not been completed.")
CheckBox12.Value = False
End If

If CheckBox12.Value = True And Range("F13").Value <> True Then
Range("F13").Value = True
Label12.Caption = Date + Time
CheckBox12.Enabled = False

Department = "Farmsight"
NxtDept = "Farmsight"
NxtTask = Sheets(NxtDept).Range("B14").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If
End Sub

Private Sub CheckBox13_Click()
If CheckBox12.Value = False And CheckBox13.Value = True Then
MsgBox ("Previous step has not been completed.")
CheckBox13.Value = False
End If

If CheckBox13.Value = True And Range("F14").Value <> True Then
Range("F14").Value = True
Label1.Caption = Date + Time
CheckBox13.Enabled = False

Department = "Farmsight"
NxtDept = "Farmsight"
NxtTask = Sheets(NxtDept).Range("B15").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If
End Sub

Private Sub CheckBox14_Click()
If CheckBox13.Value = False And CheckBox14.Value = True Then
MsgBox ("Previous step has not been completed.")
CheckBox14.Value = False
End If

If CheckBox14.Value = True And Range("F15").Value <> True Then
Range("F15").Value = True
Label14.Caption = Date + Time
CheckBox14.Enabled = False

Department = "Farmsight"
NxtDept = "Farmsight"
NxtTask = Sheets(NxtDept).Range("B16").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If
End Sub

Private Sub CheckBox15_Click()
If CheckBox14.Value = False And CheckBox15.Value = True Then
MsgBox ("Previous step has not been completed.")
CheckBox15.Value = False
End If

If CheckBox15.Value = True And Range("F16").Value <> True Then
Range("F16").Value = True
Label15.Caption = Date + Time
CheckBox15.Enabled = False

Department = "Farmsight"
NxtDept = "Farmsight"
NxtTask = Sheets(NxtDept).Range("B17").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If
End Sub

Private Sub CheckBox16_Click()
If CheckBox15.Value = False And CheckBox16.Value = True Then
MsgBox ("Previous step has not been completed.")
CheckBox16.Value = False
End If

If CheckBox16.Value = True And Range("F17").Value <> True Then
Range("F17").Value = True
Label16.Caption = Date + Time
CheckBox16.Enabled = False

Department = "Farmsight"
NxtDept = "Farmsight"
NxtTask = Sheets(NxtDept).Range("B18").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If
End Sub

Private Sub CheckBox17_Click()
If CheckBox16.Value = False And CheckBox17.Value = True Then
MsgBox ("Previous step has not been completed.")
CheckBox17.Value = False
End If

If CheckBox17.Value = True And Range("F18").Value <> True Then
Range("F18").Value = True
Label17.Caption = Date + Time
CheckBox17.Enabled = False

Department = "Farmsight"
NxtDept = "Farmsight"
NxtTask = Sheets(NxtDept).Range("B19").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If
End Sub

Private Sub CheckBox18_Click()
If CheckBox17.Value = False And CheckBox18.Value = True Then
MsgBox ("Previous step has not been completed.")
CheckBox18.Value = False
End If

If CheckBox18.Value = True And Range("F19").Value <> True Then
Range("F19").Value = True
Label18.Caption = Date + Time
CheckBox18.Enabled = False

Department = "Farmsight"
NxtDept = "Farmsight"
NxtTask = Sheets(NxtDept).Range("B20").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If
End Sub

Private Sub CheckBox19_Click()
If CheckBox18.Value = False And CheckBox19.Value = True Then
MsgBox ("Previous step has not been completed.")
CheckBox19.Value = False
End If

If CheckBox19.Value = True And Range("F20").Value <> True Then
Range("F20").Value = True
Label19.Caption = Date + Time
CheckBox19.Enabled = False

Department = "Farmsight"
NxtDept = "Farmsight"
NxtTask = Sheets(NxtDept).Range("B21").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If
End Sub

Private Sub CheckBox2_Click()
If CheckBox2.Value = True And Sheets("SALES").CheckBox13.Value = False And CheckBox1.Value = False Then
MsgBox ("Previous step has not been completed." & vbNewLine & "Sales step required" & vbNewLine & "Previous Department Step is Incomplete.")
CheckBox2.Value = False
Else

    If CheckBox2.Value = True And Sheets("SALES").CheckBox13.Value = False Then
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

Department = "Farmsight"
NxtDept = "Farmsight"
NxtTask = Sheets(NxtDept).Range("B4").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If
End Sub

Private Sub CheckBox20_Click()
If CheckBox19.Value = False And CheckBox20.Value = True Then
MsgBox ("Previous step has not been completed.")
CheckBox20.Value = False
End If

If CheckBox20.Value = True And Range("F2").Value <> True Then
Range("F21").Value = True
Label20.Caption = Date + Time
CheckBox20.Enabled = False

End If
'lock document from editing when complete
End Sub

Private Sub CheckBox3_Click()
If CheckBox2.Value = False And CheckBox3.Value = True Then
MsgBox ("Previous step has not been completed.")
CheckBox3.Value = False
End If

If CheckBox3.Value = True And Range("F4").Value <> True Then
Range("F4").Value = True
Label3.Caption = Date + Time
CheckBox3.Enabled = False

Department = "Farmsight"
NxtDept = "Farmsight"
NxtTask = Sheets(NxtDept).Range("B5").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If
End Sub

Private Sub CheckBox4_Click()
If CheckBox3.Value = False And CheckBox4.Value = True Then
MsgBox ("Previous step has not been completed.")
CheckBox4.Value = False
End If

If CheckBox4.Value = True And Range("F5").Value <> True Then
Range("F5").Value = True
Label4.Caption = Date + Time
CheckBox4.Enabled = False

Department = "Farmsight"
NxtDept = "Service"
NxtTask = Sheets(NxtDept).Range("B7").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If
End Sub

Private Sub CheckBox5_Click()
If CheckBox5.Value = True And Sheets("SALES").CheckBox14.Value = False And CheckBox4.Value = False Then
MsgBox ("Previous step has not been completed." & vbNewLine & "Sales step required" & vbNewLine & "Previous Department Step is Incomplete.")
CheckBox5.Value = False
Else

    If CheckBox5.Value = True And Sheets("SALES").CheckBox14.Value = False Then
    MsgBox ("Previous step has not been completed." & vbNewLine & "Sales step required")
    CheckBox5.Value = False
    Else
    
        If CheckBox5.Value = True And CheckBox4.Value = False Then
        MsgBox ("Previous Department Step is Incomplete.")
        CheckBox5.Value = False
        End If
    End If
End If

If CheckBox5.Value = True And Range("F6").Value <> True Then
Range("F6").Value = True
Label5.Caption = Date + Time
CheckBox5.Enabled = False

Department = "Farmsight"
NxtDept = "Farmsight"
NxtTask = Sheets(NxtDept).Range("B7").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If
End Sub

Private Sub CheckBox6_Click()
If CheckBox5.Value = False And CheckBox6.Value = True Then
MsgBox ("Previous step has not been completed.")
CheckBox6.Value = False
End If

If CheckBox6.Value = True And Range("F7").Value <> True Then
Range("F7").Value = True
Label6.Caption = Date + Time
CheckBox6.Enabled = False

Department = "Farmsight"
NxtDept = "Farmsight"
NxtTask = Sheets(NxtDept).Range("B8").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If
End Sub

Private Sub CheckBox7_Click()
If CheckBox6.Value = False And CheckBox7.Value = True Then
MsgBox ("Previous step has not been completed.")
CheckBox7.Value = False
End If

If CheckBox7.Value = True And Range("F8").Value <> True Then
Range("F8").Value = True
Label7.Caption = Date + Time
CheckBox7.Enabled = False

Department = "Farmsight"
NxtDept = "Farmsight"
NxtTask = Sheets(NxtDept).Range("B9").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If
End Sub

Private Sub CheckBox8_Click()
If CheckBox7.Value = False And CheckBox8.Value = True Then
MsgBox ("Previous step has not been completed.")
CheckBox8.Value = False
End If

If CheckBox8.Value = True And Range("F9").Value <> True Then
Range("F9").Value = True
Label8.Caption = Date + Time
CheckBox8.Enabled = False

Department = "Farmsight"
NxtDept = "Farmsight"
NxtTask = Sheets(NxtDept).Range("B10").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If
End Sub

Private Sub CheckBox9_Click()
If CheckBox8.Value = False And CheckBox9.Value = True Then
MsgBox ("Previous step has not been completed.")
CheckBox9.Value = False
End If

If CheckBox9.Value = True And Range("F10").Value <> True Then
Range("F10").Value = True
Label9.Caption = Date + Time
CheckBox9.Enabled = False

Department = "Farmsight"
NxtDept = "Farmsight"
NxtTask = Sheets(NxtDept).Range("B11").Value

If Sheets(NxtDept).Range("I7").Value <> "" Then
Recipient = Sheets(NxtDept).Range("I7").Value
Else
Recipient = InputBox("No Email Was Entered" & vbNewLine & "Please Enter An Email", "Enter Email")
End If
Call SendEmail
End If
End Sub
