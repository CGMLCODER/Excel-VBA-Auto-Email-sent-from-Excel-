Option Explicit

Sub Auto_email()
Dim OutlookApp As Object
Dim OutlookMail As Object
Dim ws As Worksheet
Dim ws_1 As Worksheet
Dim emailAddress As String
Dim subject As String
Dim body As String
Dim i As Long
Dim j As Long
Dim sin As String
Dim x As Long
Dim R_name As String
Dim final As String
Dim S_Name As String
Dim s_mail As String
Dim file_pth As String


'Set your worksheet
Set ws = ThisWorkbook.Sheets("Email Add") ' Change "Email Add" to your sheet name
Set ws_1 = ThisWorkbook.Sheets("Message") ' Change "Message" to your sheet name l"
x = ws_1.Cells(Rows.Count, 2).End(xlUp).Row
S_Name = ws.Cells(2, 8).Value
s_mail = ws.Cells(3, 8).Value
file_pth = ws.Cells(2, 11).Value
sin = "<br><br>Best regards,<br><b><br>" & S_Name & "</b><br>"
'Loop through each row in the worksheet
For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Assuming email addresses start from row 2
ws.Cells(i, 5).ClearContents
If ws.Cells(i, 4).Value = "Y" Then
emailAddress = ws.Cells(i, 3).Value ' Assuming email addresses are in column A
subject = ws.Cells(2, 6).Value ' Assuming subject is in cell F2
R_name = ws.Cells(i, 2).Value
body = "Hi " & R_name & "<br><br>" 'ws_1.Cells(2, 2).Value ""

For j = 2 To x
body = body & ws_1.Cells(j, 2).Value & "<br>"
Next j
'img adding
body = body
'sin adding
body = body & "<br>" & sin
'Create Outlook application and mail item
Set OutlookApp = CreateObject("Outlook.Application")
Set OutlookMail = OutlookApp.CreateItem(0)
'Set email properties
With OutlookMail
.To = emailAddress
.subject = subject
.htmlbody = body

.display

End With

'Release objects
Set OutlookMail = Nothing
Set OutlookApp = Nothing
ws.Cells(i, 5).Value = "Email Sent"
ws.Cells(i, 5).Interior.ColorIndex = 8

Else
ws.Cells(i, 5).Value = "Email Not sent "
ws.Cells(i, 5).Interior.ColorIndex = 3
End If



Next i
End Sub

