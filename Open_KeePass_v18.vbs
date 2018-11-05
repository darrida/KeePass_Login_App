Dim objFile
Set wshshell = wscript.createObject("wscript.shell")
Set objFSO = wscript.CreateObject("Scripting.FileSystemObject")

'
'test1 = ReadINI("KP1_PATH")
'test2 = ReadINI("KP1_FILE")
'test3 = ReadINI("KP1_ENCRYPT_HASH")

'KeePass Files without a key
'FileLoc1 = 
'File1 = 
'Filepw1 = 

'KeePass Files with a key
'FileLoc3 = 
'File3 = 
'Filepw3 = 
'KeyLoc3 = 
'Key3 = 
'*********************************
' For additional KeyPass files (NO key):
'(1) Create 3 additional lines for each:
'	[LocationFile] = "C:\Location"
'	[KeePassDBName] = "Name.kdbx"
'	[EncryptedPasswordText] = "name.txt"
'(2) Add a new "Call" statement to the bottom of the main() function:
'	Call KPLogin([LocationFile], [KeePassDBName], [EncryptedPasswordText])
'
' For additional KeyPass files (WITH key):
'(1) Create 3 additional lines for each:
'	[LocationFile] = "C:\Location"
'	[KeePassDBName] = "Name.kdbx"
'	[EncryptedPasswordText] = "name.txt"
'	[KeyLocation] = "C:\Location"
'	[KeyFileName] = "name.key"
'(2) Add a new "Call" statement to the bottom of the main() function:
'	Call KPLogin_Key([LocationFile], [KeePassDBName], [EncryptedPasswordText], [KeyLocation], [KeyFileName])
'*********************************
DecryptFileLoc = ReadINI("DECRYPT_FILE_PATH")
print1=MsgBox(DecryptFileLoc)
DecryptFile = ReadINI("DECRYPT_FILE")
print1=MsgBox(DecryptFile)

KPLocation = ReadINI("KEEPASS_PATH")
print1=MsgBox(KPLocation)
KPFileName = ReadINI("KEEPASS_APP")
print1=MsgBox(KPFileName)

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
	'test_path = ReadINI("KP1_PATH")
	'test_file = ReadINI("KP1_FILE")
	'test_enc = ReadINI("KP1_ENCRYPT_HASH")
	'Call DecryptFileExistOpen()
	Call OpenKP()
	Call KPLogin(ReadINI("KP1_PATH"), ReadINI("KP1_FILE"), ReadINI("KP1_ENCRYPT_HASH"))
'	Call KPLogin(FileLoc2, File2, Filepw2)
'	Call KPLogin(test_path, test_file, test_enc)
'	Call KPLogin_Key(FileLoc3, File3, Filepw3, KeyLoc3, Key3)
End Function

Private Function OpenKP()

	On Error Resume Next
	
	If (wshshell.AppActivate("KeePass")) Then
		wshshell.AppActivate "KeePass"
	Else
		KPexe = ReadINI("KEEPASS_PATH") + ReadINI("KEEPASS_APP")
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
		KPexe = ReadINI("KEEPASS_PATH") + "\" + ReadINI("KEEPASS_APP")
		statusCode = wshshell.Run(KPexe, 1, false)
	End If

End Function

Private Function KPLogin(location, filename, password)
	' Target the open KeePass application
	' Open new file window
	' Input path into name field
	' Input file name into name field to select file
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
	' Target the open KeePass application
	' Open new file window
	' Input path into name field
	' Input file name into name field to select file
	
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
	
	Set objFile = objFSO.OpenTextFile(DecryptFileLoc + DecryptFile, 1, ForReading)
	If IsObject(objFile) Then
		Continue
	Else 
		Set objFile = objFSO.OpenTextFile(DecryptFileLoc + "\" + DecryptFile, 1, ForReading)
	End If
	If IsObject(objFile) = False Then
		MsgBox "KeePass Open Script unable to run. Decrypt2 file is not available."
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
	
	'Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile("keepass_login_config.ini", ForReading)
	
	Do Until objTextFile.AtEndOfStream
		strNextLine = objTextFile.Readline
		intLineFinder = InStr(strNextLine, search)
		If intLineFinder <> 0 Then
			ReadINILoc = strNextLine
			ReadINIIndex = InStr(ReadINILoc, "=")
			ReadINIResult = Mid(ReadINIloc, ReadINIIndex + 1)
		End If
	Loop
	ReadINI = ReadINIResult
	objTextFile.Close
End Function
