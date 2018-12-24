Dim ToAddress
Dim CCAddress
Dim MessageSubject
Dim MessageBody
Dim MessageAttachment
Const FilePath = "C:\Users\hero_hamada\Documents\DailyStatus"


Dim ol, ns, newMail

ToAddress = "abc.xyz@outlook.com"
'Brindha.Nagaraj@wright.com;
MessageSubject = "Daily Job Status : " & (FormatDateTime(now(), 1)) 



' Build HTML for message body.
MessageBody = "Hi abc/xyz" & vbCrLf & vbCrLf & _
"Please find attached Job Status sheet with this e-mail." & vbCrLf & _
"Apart from here and there, all other reports has been completed successfully." & vbCrLf & vbCrLf & _
"Regards," & vbCrLf & _
"Sanjeet"


' connect to Outlook
Set ol = WScript.CreateObject("Outlook.Application")
Set ns = ol.getNamespace("MAPI")

Set newMail = ol.CreateItem(olMailItem)
newMail.Subject = MessageSubject
newMail.Body = MessageBody & vbCrLf

' validate the recipient, just in case...
Set myRecipient = ns.CreateRecipient(ToAddress)
myRecipient.Resolve
If Not myRecipient.Resolved Then
  MsgBox "Unknown recipient"
Else
  newMail.Recipients.Add ToAddress
newMail.CC = "some.body@outlook.com"

  count=0
  set objFSO = CreateObject("Scripting.FileSystemObject")
  set objFolder= objFSO.GetFolder(FilePath)
  For Each f In objFolder.Files
	'msgBox f.name
	newMail.Attachments.Add(FilePath & "\" & f.name)
  count=count+1
  Next

  if count=0 then
  	msgbox "No files found"
  wscript.quit
  end if

  newMail.Send
End If

Set ol = Nothing

