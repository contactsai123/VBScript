Set WshShell = CreateObject("WScript.Shell")
StrCur = WshShell.CurrentDirectory
WScript.Echo StrCur
Set xlApp = CreateObject("Excel.Application")
Set objFso = CreateObject("Scripting.FileSystemObject")
if objFso.FileExists(StrCur&"\Testdata.xlsx") then
	WScript.echo "File exists"
	else
	WScript.ECHO "File not exists"
End if
Set xlbook = xlApp.Workbooks.Open(StrCur&"\Testdata.xlsx")
Wscript.echo "Workbook opened!"
Set xlbook = nothing
Set xlApp = nothing