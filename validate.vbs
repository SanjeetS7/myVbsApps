Dim objFSO, dataArray, clippedArray
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Create an array out of the CSV

'open the data file
Set oTextStream = objFSO.OpenTextFile("C:\Users\453562\Desktop\Quota\us_quota_file_dist.csv")
Set newFile = objFSO.CreateTextFile("C:\Users\453562\Desktop\loginfo.txt")
'make an array from the data file
dataArray = Split(oTextStream.ReadAll, vbNewLine)
'close the data file
oTextStream.Close
WScript.Echo "No. of Rows is " & UBound(dataArray)+1

For Each strLine In dataArray
    'Now make an array from each line
    clippedArray =  Split(strLine,"|")
	lngHowManyColumns = UBound(clippedArray)
	WScript.Echo UBound(clippedArray)	 

Next
WScript.Echo "No. of Columns is " & lngHowManyColumns

WScript.Echo "Done"
