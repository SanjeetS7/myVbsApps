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
	Dim fso, fPath, oFile, pass, replaceS
	Const wr = 2
	Const r = 1
	Set fso = CreateObject("Scripting.FileSystemObject")
	set fPath = fso.GetFolder("cred")
	
	Set rFile = fso.OpenTextFile(fPath &"\credentials.txt", R)
		For i = 1 to 5
			pArray(i) = rFile.ReadLine
		Next
	replaceS = pArray(1)
	rfie.Close	
	
	Set wFile = fso.OpenTextFile(fPath &"\credentials.txt", wr, true)
	pass = Encrypt(InputBox("Input your password Here", "Password"))
	
	content = wFile.readAll
	'msgBox pass 
	oFile.Write(Replace(content,replaceS, pass))
End Function