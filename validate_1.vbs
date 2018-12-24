Option Explicit

Const csFSpec = "C:\Users\453562\Desktop\Quota\us_quota_file_dist.csv"
Const csFSep = "|"
Const cnCols = 15    ' one less to optimize for UBound()!
Const cnRows = 6    ' one more to optimize for .Line

Dim goFS : Set goFS = CreateObject("Scripting.FileSystemObject")
Dim oTS : Set oTS = goFS.OpenTextFile(csFSpec)
Dim sL, nC
Do Until oTS.AtEndOfStream
   sL = oTS.ReadLine()
   nC = UBound(Split(sL, csFSep))
   If cnCols <> nC Then
      WScript.Echo "Col Error in line", oTS.Line - 1 & ":", cnCols, "<>", nC, "(" & sL & ")"
   End If
Loop
'If cnRows <> oTS.Line Then
'   WScript.Echo "Row Error: ", cnRows, "<>", oTS.Line 
'End If
oTS.Close
WScript.Echo "Total No of columns is " & cnCols