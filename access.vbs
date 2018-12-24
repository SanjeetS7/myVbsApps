'===========================
strServer = "torbigdata01-qa.example.com"
strUsername = "username@int"	'InputBox("Enter the user name to access " & strServer, "Network Username", "username@domain")
strPassword = "Password#124"	'InputBox("Enter the password for the user " & strUsername, "Network Password")

strCommand = "cmd /c cmdkey /add:" & strServer & " /user:" & strUsername & " /pass:" & strPassword
Set objShell = CreateObject("WScript.Shell")
objShell.Run strCommand, 0, True

MsgBox strUsername & " has been added to the credentials list."

Set objFSO = CreateObject("Scripting.FileSystemObject")

strSourceFile = "C:\Users\ds\Desktop\loginfo.txt"
strDestDir = "\\torbigdata01-qa.example.com\e$\Data_Processor\"				'Double backslash used to access network shares with UNC naming scheme.	
If Right(strDestDir, 1) <> "\" Then
      strDestDir = strDestDir & "\"
End If
If objFSO.FileExists(strSourceFile) Then
      If objFSO.FolderExists(strDestDir) Then
            objFSO.CopyFile strSourceFile, strDestDir, True
      Else
            MsgBox "Please check that the destination folder exists of:" & VbCrLf & strDestDir
      End If
Else
      MsgBox "Please check that the source file exists of:" & VbCrLf & strSourceFile
End If

strCommand = "cmd /c cmdkey /delete:" & strServer
Set objShell = CreateObject("WScript.Shell")
'objShell.Run strCommand, 0, True
'=======================


