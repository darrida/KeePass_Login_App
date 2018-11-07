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

Call main()

Private Function main()
	Call DecryptFileExistOpen()
	wshshell.run "%COMSPEC%" 
	'WaitForWindow("Select Administrator: C:\Windows\System32\cmd.exe")
	wscript.Sleep 1000
	wshshell.SendKeys "C:\KeePass-Modified\KeePass.exe C:\local-work\bh_keepass\NewDatabase.kdbx -preselect:C:\local-work\KeePass_Login_App_ini\NewDatabase.key"
	wshshell.SendKeys "{ENTER}"
	Call WaitForWindow("Open Database - NewDatabase.kdbx")
	wshshell.SendKeys "Peterpan11" 'DecryptPass(KP1_ENCRYPT_HASH)
	wshshell.SendKeys "{ENTER}"
	wscript.Sleep 1000
	'wshshell.run "cmd"
	
	wshshell.AppActivate("Administrator:")' C:\Windows\System32\cmd.exe")
	wscript.Sleep 100
	'Call WaitForWindow("Administrator:")
	wshshell.SendKeys "C:\KeePass-Modified\KeePass.exe " + KP1_PATH + KP1_FILE
	wshshell.SendKeys "{ENTER}"
	Call WaitForWindow("Open Database - " + KEY1_FILE)
	'C:\local-work\bh_keepass\NewDatabase.kdbx"
	wshshell.SendKeys DecryptPass(KP1_ENCRYPT_HASH)
	wshshell.SendKeys "{ENTER}"
End Function

Private Function WaitForWindow(title)
	count1 = 0
		Do Until wshshell.AppActivate(title) 
			If (count1 < 50) Then
				wscript.Sleep 100
				count1 = count1 + 1
			Else
				MsgBox "TIMEOUT Error 1: KeePass is not open. Please try again."
				wscript.Quit
			End If
		Loop
	wscript.Sleep 100
End Function

Private Function OpenKP()
	On Error Resume Next
	
	If (wshshell.AppActivate("KeePass")) Then
		wshshell.AppActivate "KeePass"
	Else
		KPexe = KPLocation + KPFileName
		statusCode = wshshell.Run(KPexe, 1, false)
		count1 = 0
		Do Until wshshell.AppActivate("KeePass") 
			If (count1 < 50) Then
				wscript.Sleep 100
				count1 = count1 + 1
			Else
				MsgBox "TIMEOUT Error 1: KeePass is not open. Please try again."
				wscript.Quit
			End If
		Loop
	End If
	
	If IsObject(statusCode) Then
		Continue
	Else
		KPexe = KPLocation + "\" + KPFileName
		statusCode = wshshell.Run(KPexe, 1, false)
	End If

End Function

Private Function KPLogin(location, filename, password)
	url = location + filename
	
	count2 = 0
	Do Until wshshell.AppActivate("KeePass") 
		If (count2 < 50) Then
			wscript.Sleep 100
			count2 = count2 + 1
		Else
			MsgBox "TIMEOUT Error 2: KeePass is not open. Please try again."
			wscript.Quit
		End If
	Loop
	
	If (wshshell.AppActivate("KeePass")) Then
		wshshell.AppActivate "KeePass"
		wshshell.SendKeys ("^o")
		wscript.Sleep 1000
		wshshell.SendKeys location + filename
		wshshell.SendKeys "{ENTER}"
		wscript.Sleep 1000
	Else
		MsgBox "KeePass not detected. Please try again."
		wscript.Quit
	End If
	' Target window requesting password
	' Input password
	
	count = 0
	Do Until wshshell.AppActivate("Open Database - " & filename) 
		If (count < 50) Then
			wscript.Sleep 100
			count = count + 1
		Else
			MsgBox "TIMEOUT Error 3: Database login window for " + filename + " is not open. Please try again."
			wscript.Quit
		End If
	Loop
	
	If (wshshell.AppActivate("Open Database - " & filename)) Then
		wshshell.AppActivate "Open Database - " & filename
		wshshell.SendKeys DecryptPass(password)
		wscript.Sleep 100
		wshshell.SendKeys "{ENTER}"
	Else
		MsgBox "Database login window for " + filename + " is not open. Please try again."
		wscript.Quit
	End If

End Function

Private Function KPLogin_Key(location, filename, password, keylocation, keyfile)
	count = 0
	Do Until wshshell.AppActivate("KeePass") 
		If (count < 50) Then
			wscript.Sleep 100
			count = count + 1
		Else
			MsgBox "TIMEOUT Error 2: KeePass is not open. Please try again."
			wscript.Quit
		End If
	Loop
	
	If (wshshell.AppActivate("KeePass")) Then
		wshshell.AppActivate "KeePass"
		wshshell.SendKeys ("^o")
		wshshell.SendKeys ("%n")
		wscript.Sleep 1000
		wshshell.SendKeys location + filename '& "{ENTER}"
		wshshell.SendKeys "{ENTER}"
		'wscript.Sleep 1000
		'wshshell.SendKeys filename & "{ENTER}"
		wscript.Sleep 1000
	Else
		MsgBox "KeePass not detected. Please try again."
	End If
	
	count = 0
	Do Until wshshell.AppActivate("Open Database - " & filename) 
		If (count < 50) Then
			wscript.Sleep 100
			count = count + 1
		Else
			MsgBox "TIMEOUT Error 3: Database login window for " + filename + " is not open. Please try again."
			wscript.Quit
		End If
	Loop
	
	If (wshshell.AppActivate("Open Database - " & filename)) Then
		wshshell.AppActivate "Open Database - " & filename
		wshshell.SendKeys DecryptPass(password)
		'wscript.Quit
		wscript.Sleep 1000
		wshshell.SendKeys "{TAB}"
		wshshell.SendKeys "{TAB}"
		wshshell.SendKeys "{TAB}"
		wshshell.SendKeys "{TAB}"
		wshshell.SendKeys "{ENTER}"
		wscript.Sleep 1000
		wshshell.SendKeys keylocation + keyfile
		wshshell.SendKeys "{ENTER}"
		'wscript.Sleep 1000
		'wshshell.SendKeys keyfile & "{ENTER}"
		wscript.Sleep 1000
		wshshell.SendKeys "{TAB}"
		'wscript.Sleep 1000
		wshshell.SendKeys "{TAB}"
		'wscript.Sleep 1000
		wshshell.SendKeys "{ENTER}"
	Else
		MsgBox "Database login window for " + filename + " is not open. Please try again."
		wscript.Quit
	End If
	
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
