



Dim wshshell : Set wshshell = wscript.CreateObject("WScript.Shell")

wshshell.run "cmd.exe /C copy ""H:\Scripts\KeePass_Launcher_Login\2018.07.10_New Encryption\Keepass.exe"" "" ""X:\IT\IT Customer Service\TAC Specific\TAC_KeyPass_Database.kdbx"""

wscript.Quit
'Set objCMD = wshshell.Exec("cmd.exe")    'write the full path of application 
Set objCMD = wshshell.Exec("cmd.exe")    'write the full path of application 
WScript.Sleep 1000                               'stop script 1 sec waiting run the App
Set oCMD = objCMD.stdIn
wshshell.SendKeys "test"
CMD_processID = objCMD.ProcessID
MsgBox "my App PID  :   "  & CMD_processID
wshshell.AppActivate("PuTTY Connection Manager")
wscript.sleep 1000
wshshell.AppActivate(CMD_processID)
wscript.sleep 1000
MsgBox "is it still there?"
oCMD.Close
Do While objCMD.Status = 0
		WScript.Sleep 100
Loop
MsgBox "done"