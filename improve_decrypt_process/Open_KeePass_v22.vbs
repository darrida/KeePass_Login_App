Dim objFile, DecryptFile, DecryptFileLoc
Set wshshell = wscript.createObject("wscript.shell")
Set objFSO = wscript.CreateObject("Scripting.FileSystemObject")
Set objFSOReader = wscript.CreateObject("Scripting.FileSystemObject")

KPLocation = ReadINI("KEEPASS_PATH")
KPFileName = ReadINI("KEEPASS_APP")
DecryptFileLoc = ReadINI("DECRYPT_FILE_PATH")
DecryptFile = ReadINI("DECRYPT_FILE")

KP1_PATH = ReadINI("KP1_PATH")
KP1_FILE = ReadINI("KP1_FILE")
KP1_ENCRYPT_HASH = ReadINI("KP1_ENCRYPT_HASH")

KP2_PATH = ReadINI("KP2_PATH")
KP2_FILE = ReadINI("KP2_FILE")
KP2_ENCRYPT_HASH = ReadINI("KP2_ENCRYPT_HASH")

KP3_PATH = ReadINI("KP3_PATH")
KP3_FILE = ReadINI("KP3_FILE")
KP3_ENCRYPT_HASH = ReadINI("KP3_ENCRYPT_HASH")
KP3_KEY_PATH = ReadINI("KP3_KEY_PATH")
KP3_KEY = ReadINI("KP3_KEY")


Call DecryptFileExistOpen()
KeyEnd = inputbox("try me")
Dim Test
Set Test = (New DecryptClass)(KP1_ENCRYPT_HASH, KeyEnd)
Test.show
wscript.Quit

call Include("Decrypt_pass_v4.vbs")

Sub Include (Scriptnaam)
  Dim oFile
  Set oFile = objFSO.OpenTextFile(Scriptnaam)
  ExecuteGlobal oFile.ReadAll()
  oFile.Close
End Sub

'Call main()

Private Function main()
	Call DecryptFileExistOpen()
	wshshell.run "%COMSPEC%" 
	
	Call KP_LOGIN(KP1_PATH, KP1_FILE, KP1_ENCRYPT_HASH)
	Call KP_LOGIN(KP2_PATH, KP2_FILE, KP2_ENCRYPT_HASH)
	Call KP_LOGIN_KEY(KP3_PATH, KP3_FILE, KP3_ENCRYPT_HASH, KP3_KEY_PATH, KP3_KEY)
	
	wshshell.AppActivate("Administrator:")
	wshshell.SendKeys "exit" & "{ENTER}"
End Function

Private Function KP_LOGIN_KEY(path, file, password, key_path, key_file)
	If (path <> "") Then
		wshshell.AppActivate("Administrator:")
		wscript.Sleep 1000
		wshshell.SendKeys chr(34) + KPLocation + KPFileName + chr(34) + " " + chr(34) + path + file + chr(34) + " -preselect:" + chr(34) + key_path + key_file + chr(34)
		wshshell.SendKeys "{ENTER}"
		Call WaitForWindow("Open Database - " + file)
		Call Paste(DecryptPass(password))
		wshshell.SendKeys "{ENTER}"
		Call WaitForWindow(file + " - KeePass")
	End If
End Function

Private Function KP_LOGIN(path, file, password)
	If (path <> "") Then
		wshshell.AppActivate("Administrator:")
		wscript.Sleep 1000
		wshshell.SendKeys chr(34) + KPLocation + KPFileName + chr(34) + " " + chr(34) + path + file + chr(34)
		wshshell.SendKeys "{ENTER}"
		Call WaitForWindow("Open Database - " + file)
		Call Paste(DecryptPass(password))
		wshshell.SendKeys "{ENTER}"
		Call WaitForWindow(file + " - KeePass")
	End If
End Function

Private Function Paste(copy)
	Set oExec = WshShell.Exec("clip")
	Set oIn = oExec.stdIn
	oIn.WriteLine copy
	oIn.Close
	Do While oExec.Status = 0
		WScript.Sleep 100
	Loop
	wshshell.SendKeys "^V"
	WScript.Sleep 100
	ClearCopy()
End Function

Private Function ClearCopy()
	Set oExec = WshShell.Exec("clip")
	Set oIn = oExec.stdIn
	oIn.WriteLine ""
	oIn.Close
	Do While oExec.Status = 0
		WScript.Sleep 100
	Loop
End Function

Private Function WaitForWindow(title)
	count1 = 0
		Do Until wshshell.AppActivate(title) 
			If (count1 < 50) Then
				wscript.Sleep 100
				count1 = count1 + 1
			Else
				MsgBox "TIMEOUT Error: Waited for a window titled " + title + ", but it never showed up."
				wscript.Quit
			End If
		Loop
	wscript.Sleep 1000
End Function

Private Function DecryptFileExistOpen()
	On Error Resume Next
	Const ForReading = 1
	
	Set objFile = objFSO.OpenTextFile(DecryptFileLoc + DecryptFile, ForReading)
	If IsObject(objFile) Then
		Continue
	Else 
		Set objFile = objFSO.OpenTextFile(DecryptFileLoc + "\" + DecryptFile, 1, ForReading)
	End If
	If IsObject(objFile) = False Then
		MsgBox "KeePass Open Script unable to run. Decrypt file is not available."
		wscript.Quit
	End If
	extFunctionsStr2 = objFile.ReadAll
	objFile.Close
	Set objFile = Nothing
	Set objFSO = Nothing
	ExecuteGlobal extFunctionsStr2
End Function

Private Function ReadINI(search)
	Const ForReading = 1
	Const ForWriting = 2
	
	'Set objFSOReader = CreateObject("Scripting.FileSystemObject")
	Set objINIFile = objFSOReader.OpenTextFile("keepass_login_config.ini", ForReading)
	
	Do Until objINIFile.AtEndOfStream
		strNextLine = objINIFile.Readline
		intLineFinder = InStr(strNextLine, search)
		If intLineFinder <> 0 Then
			ReadINILoc = strNextLine
			ReadINIIndex = InStr(ReadINILoc, "=")
			ReadINIResult = Mid(ReadINIloc, ReadINIIndex + 1)
		End If
	Loop
	ReadINI = ReadINIResult
	objINIFile.Close
End Function