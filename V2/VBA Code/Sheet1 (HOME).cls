VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Sales()
Sheets("SALES").Visible = True
Sheets("SALES").Activate
End Sub
Sub Service()
Sheets("SERVICE").Visible = True
Sheets("SERVICE").Activate
End Sub
Sub Parts()
Sheets("PARTS").Visible = True
Sheets("PARTS").Activate
End Sub
Sub Farmsight()
Sheets("FARMSIGHT").Visible = True
Sheets("FARMSIGHT").Activate
End Sub
Sub Accounts()
Sheets("FINANCE & ACCOUNTS").Visible = True
Sheets("FINANCE & ACCOUNTS").Activate
End Sub
Sub Details()
Sheets("DETAILS").Visible = True
Sheets("DETAILS").Activate
End Sub
Sub Home()
Sheets("HOME").Visible = True
Sheets("HOME").Activate
End Sub
Sub SalesLogin()
Dim Password As String
Dim Text As String

Password = "Sales"
Text = InputBox("Enter Password:", "Sales Login")

If Password = Text Then
Sheets("SALES").Visible = True
Sheets("SALES").Activate
Else
If Text = "" Then
Exit Sub
Else
MsgBox ("Password is Incorrect." & vbNewLine & "Please Try Again.")
End If
End If
End Sub
Sub ServiceLogin()
Dim Password As String
Dim Text As String

Password = "Service"
Text = InputBox("Enter Password:", "Service Login")

If Password = Text Then
Sheets("SERVICE").Visible = True
Sheets("SERVICE").Activate
Else
If Text = "" Then
Exit Sub
Else
MsgBox ("Password is Incorrect." & vbNewLine & "Please Try Again.")
End If
End If
End Sub
Sub PartsLogin()
Dim Password As String
Dim Text As String

Password = "Parts"
Text = InputBox("Enter Password:", "Parts Login")

If Password = Text Then
Sheets("PARTS").Activate
Else
If Text = "" Then
Exit Sub
Else
MsgBox ("Password is Incorrect." & vbNewLine & "Please Try Again.")
End If
End If
End Sub
Sub FarmsightLogin()
Dim Password As String
Dim Text As String

Password = "Farmsight"
Text = InputBox("Enter Password:", "Farmsight Login")

If Password = Text Then
Sheets("FARMSIGHT").Activate
Else
If Text = "" Then
Exit Sub
Else
MsgBox ("Password is Incorrect." & vbNewLine & "Please Try Again.")
End If
End If
End Sub
Sub FinanceAndAccountsLogin()
Dim Password As String
Dim Text As String

Password = "Accounts"
Text = InputBox("Enter Password:", "Finance & Accounts Login")

If Password = Text Then
Sheets("FINANCE & ACCOUNTS").Activate
Else
If Text = "" Then
Exit Sub
Else
MsgBox ("Password is Incorrect." & vbNewLine & "Please Try Again.")
End If
End If
End Sub

