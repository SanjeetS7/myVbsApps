'save password file
'Option Explicit  == All variables must be defined
'Vbscript set vs simple assignment
'Function vs Sub in vbscript

Option Explicit
Private Function Encrypt(ByVal string)
	Dim x, i, tmp
	For i = 1 To Len( string )
		x = Mid( string, i, 1 )
		tmp = tmp & Chr( Asc( x ) + 1 )
	Next
	tmp = StrReverse( tmp )
	Encrypt = tmp
End Function

Private Function Decrypt(ByVal encryptedstring)
	Dim x, i, tmp
	encryptedstring = StrReverse( encryptedstring )
	For i = 1 To Len( encryptedstring )
		x = Mid( encryptedstring, i, 1 )
		tmp = tmp & Chr( Asc( x ) - 1 )
	Next
	Decrypt = tmp
End Function

Function setPassword()

	Dim fso, fPath, pass, replaceS, rFile, rFileA, wFile, i, pArray(5), content, tempF
	Const write = 2
	Const read = 1
	Const append = 8
	Set fso = CreateObject("Scripting.FileSystemObject")
	set fPath = fso.GetFolder("cred")
	tempF = Replace(Wscript.ScriptFullName, Wscript.ScriptName, "\cred\temp.txt")
	fso.CreateTextFile tempF
	
	Set rFile = fso.OpenTextFile(fPath &"\credentials.txt", read)
		For i = 1 to 5
			pArray(i) = rFile.ReadLine
		Next
	rFile.Close
	replaceS = pArray(1)
	msgBox replaceS
	
	Set rFileA = fso.OpenTextFile(fPath &"\credentials.txt", read)
	content = rFileA.readAll
	rFileA.Close
	msgBox content
	
	Set wFile = fso.OpenTextFile(tempF, append)
	pass = Encrypt(InputBox("Input your password Here", "Password"))
	
	wFile.WriteLine Replace(content, replaceS, pass)
	'wFile.Write(Replace(content,replaceS, pass))
	wFile.Close
	fso.CopyFile tempF, fPath & "\credentials.txt", True
	fso.DeleteFile tempF
	
	
End Function

setPassword()