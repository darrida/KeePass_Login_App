Const ForReading = 1

Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objTextFile = objFSO.OpenTextFile("sample.ini", ForReading)


Do Until objTextFile.AtEndOfStream

    strNextLine = objTextFile.Readline


    intLineFinder = InStr(strNextLine, "DisplayWelcomeDlg")

    If intLineFinder <> 0 Then

        strNextLine = "DisplayWelcomeDlg=NO"

    End If


    intLineFinder = InStr(strNextLine, "UserName")

    If intLineFinder <> 0 Then

        strNextLine = "UserName=" + path1

    End If


    strNewFile = strNewFile & strNextLine & vbCrLf

Loop


objTextFile.Close


Set objTextFile = objFSO.OpenTextFile("sample.ini", ForWriting)


objTextFile.WriteLine strNewFile

objTextFile.Close