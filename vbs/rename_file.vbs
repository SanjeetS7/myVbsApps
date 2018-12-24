set oFso = createobject("scripting.filesystemobject")
sDirectorypath = "C:\Users\hero_hamada\Documents\DailyStatus"
rename_files(sDirectorypath)

sub rename_files(folder)
  set oFolder = oFso.getfolder(folder)
  for each oFile in oFolder.files
	file_name = Left(oFso.GetFileName(oFile), len(oFso.GetFileName(oFile))-19)
    new_name = file_name & myDateFormat(date()) & "." & oFso.GetExtensionName(oFile)
    wscript.echo "renaming " & file_name & " => " & new_name
    errResult = oFso.MoveFile(oFile, "C:\Users\hero_hamada\Documents\DailyStatus\" & new_name)
  next
  for each oSubFolder in oFolder.subfolders
    rename_files(oSubFolder)
  next
end sub


Function myDateFormat(myDate)
    d = WhatEver(Day(myDate)) & Suffix(WhatEver(Day(myDate)))
    m = MonthName(WhatEver(Month(myDate)))
    y = Year(myDate)
    myDateFormat= d & "-" & m & "-" & y
End Function

Function WhatEver(num)
    If(Len(num)=1) Then
        WhatEver="0"&num
    Else
        WhatEver=num
    End If
End Function


'****************************************
'*     VBScript - by Gegniani          *
'****************************************

Function Suffix(num)  
    Dim retVal, digits, lastTwoDigits, lastDigit
 
    digits = CInt(Len(num))
    If (digits > 1) Then
        lastTwoDigits = CInt(Mid(num, Len(num) - 1, 2)) 'vbscript is 1-based not 0-based
        If (lastTwoDigits = 11) Or (lastTwoDigits = 12) Or (lastTwoDigits = 13) Then
            retVal = "th"
            Suffix = retVal : Exit Function
        End If
    End If
    lastDigit = CInt(Mid(num, Len(num), 1))
    Select Case lastDigit
        Case 1 retVal = "st"
        Case 2 retVal = "nd"
        Case 3 retVal = "rd"
        Case Else retVal = "th"
    End Select
    Suffix = retVal
End Function