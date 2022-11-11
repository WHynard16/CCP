VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
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
Mail.HTMLBody = "<p>Wholegood Number: " & Sheets("DETAILS").Range("C2").Value & "<br>Customer: " & Sheets("DETAILS").Range("C3").Value & "<br><br>" & Department & " has completed the previous step.<br><br>Next Task Owner: " & NxtDept & "<br>Next Task: " & NxtTask & "</p>"

On Error Resume Next

Mail.Send

If Err.Number <> 0 Then
MsgBox Err.Description, vbCritical, "There was an Error!"
Exit Sub
End If

MsgBox "Notification has been sent to the next task owner.", vbinfomation, "Sent"

End Sub
Private Sub CheckBox1_Click()

If CheckBox1.Value = True And Sheets("DETAILS").Range("E5").Value = False Then
CheckBox1.Value = False
MsgBox ("Wholegood Details Missing" & vbNewLine & "Please Complete Before Continuing")
End If

If CheckBox1.Value = True And Range("F2").Value <> True Then
Range("F2").Value = True
Range("D2").Value = Date + Time
CheckBox1.Enabled = False

Department = "Sales"
NxtDept = "Sales"
NxtTask = Sheets("SALES").Range("B3").Value

If Sheets(NxtDept).TextBox1.Text <> "" Then
Recipient = Sheets(NxtDept).TextBox1.Text
Else
Recipient = InputBox("No Recipient was found" & vbNewLine & "Please Enter the Email for the " & NxtDept & " Recipient", "Enter Email")
End If

Call SendEmail
End If

End Sub
Private Sub CheckBox2_Click()

If CheckBox1.Value = False And CheckBox2.Value = True Then
CheckBox2.Value = False
MsgBox ("Previous Department Step is Incomplete.")
End If

If CheckBox2.Value = True And Range("F3").Value <> True Then
Range("F3").Value = True
Range("D3").Value = Date + Time
CheckBox2.Enabled = False

Department = "Sales"
NxtDept = "Sales"
NxtTask = Sheets("SALES").Range("B4").Value

If Sheets(NxtDept).TextBox1.Text <> "" Then
Recipient = Sheets(NxtDept).TextBox1.Text

Else
Recipient = InputBox("No Recipient was found" & vbNewLine & "Please Enter the Email for the " & NxtDept & " Recipient", "Enter Email")

End If

Call SendEmail
End If

End Sub

Private Sub CheckBox3_Click()

If CheckBox2.Value = False And CheckBox3.Value = True Then
CheckBox3.Value = False
MsgBox ("Previous Department Step is Incomplete.")

End If

If CheckBox3.Value = True And Range("F4").Value <> True Then
Range("F4").Value = True
Range("D4").Value = Date + Time
CheckBox3.Enabled = False

Department = "Sales"
NxtDept = "Finance & Accounts"
NxtTask = Sheets("FINANCE & ACCOUNTS").Range("B2").Value

If Sheets(NxtDept).TextBox1.Text <> "" Then
Recipient = Sheets("FINANCE & ACCOUNTS").Range("I7").Value

Else
Recipient = InputBox("No Recipient was found" & vbNewLine & "Please Enter the Email for the " & NxtDept & " Recipient", "Enter Email")

End If

Call SendEmail
End If
End Sub

Private Sub CheckBox4_Click()

If CheckBox4.Value = True And Sheets("FINANCE & ACCOUNTS").CheckBox1.Value = False And CheckBox3.Value = False Then
MsgBox ("Previous step has not been completed." & vbNewLine & "Finance & Accounts step required" & vbNewLine & "Previous Department Step is Incomplete.")
CheckBox4.Value = False
Else

    If CheckBox4.Value = True And Sheets("FINANCE & ACCOUNTS").CheckBox1.Value = False Then
    MsgBox ("Previous step has not been completed." & vbNewLine & "Finance & Accounts step required")
    CheckBox4.Value = False
    Else
    
        If CheckBox4.Value = True And CheckBox3.Value = False Then
        MsgBox ("Previous Department Step is Incomplete.")
        CheckBox4.Value = False
        End If
    End If
End If
    
   
If CheckBox4.Value = True And Range("F5").Value <> True Then
Range("F5").Value = True
Range("D5").Value = Date + Time
CheckBox4.Enabled = False

Department = "Sales"
NxtDept = "Sales"
NxtTask = Sheets("SALES").Range("B6").Value

If Sheets(NxtDept).TextBox1.Text <> "" Then
Recipient = Sheets(NxtDept).TextBox1.Text

Else
Recipient = InputBox("No Recipient was found" & vbNewLine & "Please Enter the Email for the " & NxtDept & " Recipient", "Enter Email")

End If
Call SendEmail
    End If
End Sub

Private Sub CheckBox5_Click()
If CheckBox4.Value = False And CheckBox5.Value = True Then
MsgBox ("Previous Department Step is Incomplete.")
CheckBox5.Value = False
End If

If CheckBox5.Value = True And Range("F6").Value <> True Then
Range("F6").Value = True
Range("D6").Value = Date + Time
CheckBox5.Enabled = False

Department = "Sales"
NxtDept = "Finance & Accounts"
NxtTask = Sheets("FINANCE & ACCOUNTS").Range("B3").Value

If Sheets(NxtDept).TextBox1.Text <> "" Then
Recipient = Sheets(NxtDept).TextBox1.Text

Else
Recipient = InputBox("No Recipient was found" & vbNewLine & "Please Enter the Email for the " & NxtDept & " Recipient", "Enter Email")

End If
Call SendEmail
End If
End Sub

Private Sub CheckBox6_Click()
If CheckBox6.Value = True And Sheets("PARTS").CheckBox2.Value = False And CheckBox5.Value = False Then
MsgBox ("Previous step has not been completed." & vbNewLine & "Parts step required" & vbNewLine & "Previous Department Step is Incomplete.")
CheckBox6.Value = False
Else

    If CheckBox6.Value = True And Sheets("PARTS").CheckBox2.Value = False Then
    MsgBox ("Previous step has not been completed." & vbNewLine & "Parts step required")
    CheckBox6.Value = False
    Else
    
        If CheckBox6.Value = True And CheckBox5.Value = False Then
        MsgBox ("Previous Department Step is Incomplete.")
        CheckBox6.Value = False
        End If
    End If
End If

If CheckBox6.Value = True And Range("F7").Value <> True Then
Range("F7").Value = True
Range("D7").Value = Date + Time
CheckBox6.Enabled = False

Department = "Sales"
NxtDept = "Sales"
NxtTask = Sheets("SALES").Range("B8").Value

If Sheets(NxtDept).TextBox1.Text <> "" Then
Recipient = Sheets(NxtDept).TextBox1.Text

Else
Recipient = InputBox("No Recipient was found" & vbNewLine & "Please Enter the Email for the " & NxtDept & " Recipient", "Enter Email")

End If
Call SendEmail
End If
End Sub

Private Sub CheckBox7_Click()
If CheckBox6.Value = False And CheckBox7.Value = True Then
MsgBox ("Previous Department Step is Incomplete.")
CheckBox7.Value = False
End If

If CheckBox7.Value = True And Range("F8").Value <> True Then
Range("F8").Value = True
Range("D8").Value = Date + Time
CheckBox7.Enabled = False

Department = "Sales"
NxtDept = "Sales"
NxtTask = Sheets("SALES").Range("B9").Value

If Sheets(NxtDept).TextBox1.Text <> "" Then
Recipient = Sheets(NxtDept).TextBox1.Text

Else
Recipient = InputBox("No Recipient was found" & vbNewLine & "Please Enter the Email for the " & NxtDept & " Recipient", "Enter Email")

End If
Call SendEmail
End If
End Sub


Private Sub CheckBox8_Click()
If CheckBox7.Value = False And CheckBox8.Value = True Then

MsgBox ("Previous Department Step is Incomplete.")
CheckBox8.Value = False
End If

If CheckBox8.Value = True And Range("F9").Value <> True Then
Range("F9").Value = True
Range("D9").Value = Date + Time
CheckBox8.Enabled = False

Department = "Sales"
NxtDept = "Sales"
NxtTask = Sheets("SALES").Range("B10").Value

If Sheets(NxtDept).TextBox1.Text <> "" Then
Recipient = Sheets(NxtDept).TextBox1.Text

Else
Recipient = InputBox("No Recipient was found" & vbNewLine & "Please Enter the Email for the " & NxtDept & " Recipient", "Enter Email")

End If
Call SendEmail
End If
End Sub

Private Sub CheckBox9_Click()
If CheckBox9.Value = True And CheckBox8.Value = False Then
        MsgBox ("Previous Department Step is Incomplete.")
        CheckBox9.Value = False
End If

If CheckBox9.Value = True And Range("F10").Value <> True Then
Range("F10").Value = True
Range("D10").Value = Date + Time
CheckBox9.Enabled = False

Department = "Sales"
NxtDept = "Service"
NxtTask = Sheets("SERVICE").Range("B2").Value

If Sheets(NxtDept).TextBox1.Text <> "" Then
Recipient = Sheets(NxtDept).TextBox1.Text

Else
Recipient = InputBox("No Recipient was found" & vbNewLine & "Please Enter the Email for the " & NxtDept & " Recipient", "Enter Email")

End If
Call SendEmail
End If

End Sub

Private Sub CheckBox10_Click()
If CheckBox10.Value = True And Sheets("Service").CheckBox2.Value = False And CheckBox9.Value = False Then
MsgBox ("Previous step has not been completed." & vbNewLine & "Service & Accounts step required" & vbNewLine & "Previous Department Step is Incomplete.")
CheckBox10.Value = False
Else

    If CheckBox10.Value = True And Sheets("Service").CheckBox2.Value = False Then
    MsgBox ("Previous step has not been completed." & vbNewLine & "Service step required")
    CheckBox10.Value = False
    Else
    
        If CheckBox10.Value = True And CheckBox9.Value = False Then
        MsgBox ("Previous Department Step is Incomplete.")
        CheckBox10.Value = False
        End If
    End If
End If

If CheckBox10.Value = True And Range("F11").Value <> True Then
Range("F11").Value = True
Range("D11").Value = Date + Time
CheckBox10.Enabled = False

Department = "Sales"
NxtDept = "Finance & Accounts"
NxtTask = Sheets("FINANCE & ACCOUNTS").Range("B6").Value

If Sheets(NxtDept).TextBox1.Text <> "" Then
Recipient = Sheets(NxtDept).TextBox1.Text

Else
Recipient = InputBox("No Recipient was found" & vbNewLine & "Please Enter the Email for the " & NxtDept & " Recipient", "Enter Email")

End If
Call SendEmail
End If
End Sub

Private Sub CheckBox11_Click()
If CheckBox11.Value = True And Sheets("FINANCE & ACCOUNTS").CheckBox5.Value = False And CheckBox10.Value = False Then
MsgBox ("Previous step has not been completed." & vbNewLine & "Finance & Accounts step required" & vbNewLine & "Previous Department Step is Incomplete.")
CheckBox11.Value = False
Else

    If CheckBox11.Value = True And Sheets("FINANCE & ACCOUNTS").CheckBox5.Value = False Then
    MsgBox ("Previous step has not been completed." & vbNewLine & "Finance & Accounts step required")
    CheckBox11.Value = False
    Else
    
        If CheckBox11.Value = True And CheckBox10.Value = False Then
        MsgBox ("Previous Department Step is Incomplete.")
        CheckBox11.Value = False
        End If
    End If
End If

If CheckBox11.Value = True And Range("F12").Value <> True Then
Range("F12").Value = True
Range("D12").Value = Date + Time
CheckBox11.Enabled = False

Department = "Sales"
NxtDept = "Parts"
NxtTask = Sheets("PARTS").Range("B4").Value

If Sheets(NxtDept).TextBox1.Text <> "" Then
Recipient = Sheets(NxtDept).TextBox1.Text

Else
Recipient = InputBox("No Recipient was found" & vbNewLine & "Please Enter the Email for the " & NxtDept & " Recipient", "Enter Email")

End If
Call SendEmail
End If

End Sub

Private Sub CheckBox12_Click()
If CheckBox12.Value = True And Sheets("FARMSIGHT").CheckBox1.Value = False And CheckBox11.Value = False Then
MsgBox ("Previous step has not been completed." & vbNewLine & "Farmsight step required" & vbNewLine & "Previous Department Step is Incomplete.")
CheckBox12.Value = False
Else

    If CheckBox12.Value = True And Sheets("FARMSIGHT").CheckBox1.Value = False Then
    MsgBox ("Previous step has not been completed." & vbNewLine & "Farmsight step required")
    CheckBox12.Value = False
    Else
    
        If CheckBox12.Value = True And CheckBox11.Value = False Then
        MsgBox ("Previous Department Step is Incomplete.")
        CheckBox12.Value = False
        End If
    End If
End If

If CheckBox12.Value = True And Range("F13").Value <> True Then
Range("F13").Value = True
Range("D13").Value = Date + Time
CheckBox12.Enabled = False

Department = "Sales"
NxtDept = "Sales"
NxtTask = Sheets("SALES").Range("B14").Value

If Sheets(NxtDept).TextBox1.Text <> "" Then
Recipient = Sheets(NxtDept).TextBox1.Text

Else
Recipient = InputBox("No Recipient was found" & vbNewLine & "Please Enter the Email for the " & NxtDept & " Recipient", "Enter Email")

End If
Call SendEmail
End If

End Sub

Private Sub CheckBox13_Click()
If CheckBox13.Value = True And CheckBox12.Value = False Then
        MsgBox ("Previous Department Step is Incomplete.")
        CheckBox13.Value = False
        End If

If CheckBox13.Value = True And Range("F14").Value <> True Then
Range("F14").Value = True
Range("D14").Value = Date + Time
CheckBox13.Enabled = False

Department = "Sales"
NxtDept = "Farmsight"
NxtTask = Sheets("Farmsight").Range("B3").Value

If Sheets(NxtDept).TextBox1.Text <> "" Then
Recipient = Sheets(NxtDept).TextBox1.Text

Else
Recipient = InputBox("No Recipient was found" & vbNewLine & "Please Enter the Email for the " & NxtDept & " Recipient", "Enter Email")

End If
Call SendEmail
End If
End Sub

Private Sub CheckBox14_Click()
If CheckBox14.Value = True And Sheets("SERVICE").CheckBox6.Value = False And CheckBox13.Value = False Then
MsgBox ("Previous step has not been completed." & vbNewLine & "Service step required" & vbNewLine & "Previous Department Step is Incomplete.")
CheckBox4.Value = False
Else

    If CheckBox14.Value = True And Sheets("SERVICE").CheckBox6.Value = False Then
    MsgBox ("Previous step has not been completed." & vbNewLine & "Service step required")
    CheckBox14.Value = False
    Else
    
        If CheckBox14.Value = True And CheckBox13.Value = False Then
        MsgBox ("Previous Department Step is Incomplete.")
        CheckBox14.Value = False
        End If
    End If
End If

If CheckBox14.Value = True And Range("F15").Value <> True Then
Range("F15").Value = True
Range("D15").Value = Date + Time
CheckBox14.Enabled = False

Department = "Sales"
NxtDept = "Farmsight"
NxtTask = Sheets("Farmsight").Range("B6").Value

If Sheets(NxtDept).TextBox1.Text <> "" Then
Recipient = Sheets(NxtDept).TextBox1.Text

Else
Recipient = InputBox("No Recipient was found" & vbNewLine & "Please Enter the Email for the " & NxtDept & " Recipient", "Enter Email")

End If
Call SendEmail
End If

End Sub

