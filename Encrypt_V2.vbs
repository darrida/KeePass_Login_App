'******************************************************************************
'** Created:      20-1-2009 22:28
'** Original Author:         Adriaan Westra
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
strKey = "[key goes here]"
'******************************************************************************
'** Tis is the text that is going to be encrypted
strEncrypt = inputbox("Password to Encrypt")
'******************************************************************************
'** This is the seed that is used for randpmizing the en/decryption
intSeed = 6
'******************************************************************************
'** Call the Encrypt function to encrypt the text
strEncryptedText = Encrypt(strEncrypt, strKey, intSeed)
'******************************************************************************
'** Store the encrypted text in a file for later use
Set objFSO = CreateObject("Scripting.FileSystemObject")
strFile = "encrypted.txt"
Set objTxTFile = objFSO.OpenTextFile(strFile, ForWriting, True)
objTxTFile.Write strEncryptedText
objTxTFile.Close
'******************************************************************************
'** Get the encrypted text from file
Set objTxtFile = objFSO.OpenTextFile(strFile)
    If Not objTxtFile.AtEndOfStream Then
      strDecrypt = objTxtFile.ReadAll
   End If
objTxTFile.Close
'******************************************************************************
'** Output the encrypted text to screen
wscript.echo "Encrypted text : " & strDecrypt

'******************************************************************************
'** Output the decrypted text to screen
wscript.echo "Decrypted text : " & Decrypt(strDecrypt, strKey, intSeed)
wscript.echo "Decrypted text is located in same directory as this script."
wscript.quit
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
Function Encrypt( strEncrypt, strKey, intSeed)
  Rnd(-1)
  Randomize intSeed
  intRnd =  Int( ( Len(strKey) - 1 + 1 ) * Rnd + 1 )
  
  arrEncrypt = String2Asc(strEncrypt)
  arrKey = String2Asc(strKey)
  
  For intI = 0 to UBound( arrEncrypt ) - 1
      
      intPointer = intI + intRnd
      If intPointer > UBound(arrKey) Then
         intPointer = intPointer -  ((UBound(arrKey) + 1 ) * Int(intPointer / (UBound(arrKey) + 1)))
      End If
      
      intCalc = arrEncrypt(intI) + arrKey(intPointer)
      
      If intCalc > 256 Then
      	intCalc = intCalc - 256 
      End If
      strEncrypted = strEncrypted & Chr(intCalc)
  Next
  encrypt = strEncrypted
End Function
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