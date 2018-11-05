'******************************************************************************
'** Script:         encrypt.vbs
'** Version:      1.0
'** Created:      20-1-2009 22:28
'** Author:         Adriaan Westra
'** E-mail:         
'**
'** Purpose / Comments:
'**               Demonstrate a simple encryption algorithm.
'**
'** Doel / Commentaar:
'**               Demoscript voor een simpel encryptie algoritme.
'**
'**
'** Changelog :
'** date / time            : 
'** 20-1-2009 22:28  :   Initial version
'**
'******************************************************************************
'******************************************************************************
'**   Constants for opening files
'**   Constanten voor het openen van files
Const ForWriting = 2 
Const ForAppending = 8
Dim strEncrypted
Dim strKey
Dim intSeed
'******************************************************************************
'** This is the key that will be use for en/decrypting the text
strKey = "[encrypt key goes here]" + inputbox("Finish Decrypt Key") ' End of key: *
'******************************************************************************
'** This is the seed that is used for randomizing the en/decryption
intSeed = 6
'******************************************************************************
'** This information is passed from the primary 
Function DecryptPass(EncryptedText)
	On Error Resume Next
	'Set objFSO = CreateObject("Scripting.FileSystemObject")
	'Set objFSO = CreateObject("Scripting.FileSystemObject")
	'Set objFile3 = objFSO.OpenTextFile(EncryptedTextFile, 1)
	'If IsObject(objFile3) Then
	'	strDecrypt = objFile3.ReadAll
	'	DecryptPass = Decrypt( strDecrypt, strKey, intSeed)
	'Else
	'	'Set objFile3 = objFSO.OpenTextFile(Encrypted_Files_Location + "\" + EncryptedTextFile, 1)
	'	Set objFile3 = objFSO.OpenTextFile(EncryptedTextFile, 1)
	'	strDecrypt = objFile3.ReadAll
		DecryptPass = Decrypt( EncryptedText, strKey, intSeed)
	'End If
	'If IsObject(objFile3) Then
	'	Continue
	'Else
	'	MsgBox "Enable to open decryption file. Make sure location and file names are correct."
	'	wscript.Quit
	'End If
End Function
'wscript.echo DecryptPass(EncryptedTextFile)
'******************************************************************************
'** Function:     String2Asc
'** Version:      1.0
'** Created:      20-1-2009 22:35
'** Author:       Adriaan Westra
'** E-mail:         
'**
'** Purpose / Comments:
'**
'**      Transform a string in to an array with ascii values
'**
'** Change Log :
'**
'** 20-1-2009 22:36 : Initial Version
'**
'** Arguments :  
'**
'**   strIn   :   string to be converted
'**
'** Returns   :
'**
'**   an array with ascii values
'**         
'******************************************************************************
Function String2Asc( strIn)
  arrResult = Array()
  ReDim arrResult( CInt( Len( strIn ) ) )
  For intI = 0 to Len(strIn) - 1
      arrResult( intI ) = Asc( Mid( strIn,intI + 1 ,1 ) )
  Next
  String2Asc = arrResult
End Function
'******************************************************************************
'** Function:     Decrypt
'** Version:      1.0
'** Created:      20-1-2009 22:35
'** Author:       Adriaan Westra
'** E-mail:         
'**
'** Purpose / Comments:
'**
'**      Decrypt an encrypted string
'**
'** Change Log :
'**
'** 20-1-2009 22:36 : Initial Version
'**
'** Arguments :  
'**
'**   strDecrypt   :   string to be Decrypted
'**   strKey       :   string used as encryption key
'**   intSeed      :   integer used to make the encryption random
'**
'** Returns   :
'**
'**   A Decrypted string
'**         
'******************************************************************************
Function Decrypt( strDecrypt, strKey, intSeed)
  Rnd(-1)
  Randomize intSeed
  intRnd =  Int( ( Len(strKey) - 1 + 1 ) * Rnd + 1 )
  
  arrDecrypt = String2Asc(strDecrypt)
  arrKey = String2Asc(strKey)
    
  For intI = 0 to UBound( arrDecrypt ) - 1
      
      intPointer = intI + intRnd
      If intPointer > UBound(arrKey) Then
         intPointer = intPointer -  ((UBound(arrKey) + 1 ) * Int(intPointer / (UBound(arrKey) + 1)))
      End If
      
      intCalc = arrDecrypt(intI) - arrKey(intPointer)
      
      If intCalc < 0 Then
      	intCalc = intCalc + 256 
      End If
      strDecrypted = strDecrypted & Chr(intCalc)
  Next
  Decrypt = strDecrypted
End Function

'wscript.echo "Decrypted text : " & Decrypt(strDecrypt, strKey, intSeed)