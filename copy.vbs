'===========================
strServer = "torbaetl01-qa.tornier.com"
strUsername = InputBox("Enter the user name to access " & strServer, "Network Username", "username@domain")
strPassword = InputBox("Enter the password for the user " & strUsername, "Network Password")

strCommand = "cmd /c cmdkey /add:" & strServer & " /user:" & strUsername & " /pass:" & strPassword
Set objShell = CreateObject("WScript.Shell")
objShell.Run strCommand, 0, True

MsgBox strUsername & " has been added to the credentials list."

Set objFSO = CreateObject("Scripting.FileSystemObject")

strSourceFile = "C:\Users\453562\Desktop\Cognos Reports.txt"
strDestDir = "d$\Data_Manager_Backup\"
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
objShell.Run strCommand, 0, True
'=======================
