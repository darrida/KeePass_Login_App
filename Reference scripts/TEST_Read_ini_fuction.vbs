Dim objFile
Set wshshell = wscript.createObject("wscript.shell")
Set objFSO = wscript.CreateObject("Scripting.FileSystemObject")

Private Function ReadINILoc(loc)
	Const ForReading = 1
	Const ForWriting = 2
	
	'Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile("keepass_login_config.ini", ForReading)
	
	Do Until objTextFile.AtEndOfStream
		strNextLine = objTextFile.Readline
		intLineFinder = InStr(strNextLine, loc)
		If intLineFinder <> 0 Then
			ReadINILoc = strNextLine
		End If
	Loop
	
	objTextFile.Close
End Function

Dim test_path, test_file, test_enc

test1 = ReadINILoc("KP1_PATH")
test_path = Mid(test1, (InStr(test1, "=")) + 1)
print2=MsgBox(test_path)

test2 = ReadINILoc("KP1_FILE")
test_file = Mid(test2, (InStr(test2, "=")) + 1)
print3=MsgBox(test_file)

test3 = ReadINILoc("KP1_ENCRYPT_HASH")
test_enc = Mid(test3, (InStr(test3, "=")) + 1)
print4=MsgBox(test_enc)

'print1=MsgBox(intLineFinder)