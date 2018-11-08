Class DecryptClass
	'Constants for opening files
	Const ForWriting = 2 
	Const ForAppending = 8
	Dim strEncrypted
	Dim strKey
	Dim intSeed

	'FinishStrKey = inputbox("Finish Decrypt Key")
	private strKey = "aslKERSi03483IOSER8#$("

	'This is the seed that is used for randomizing the en/decryption
	private intSeed = 6 '<=[input seed number - single number is ok]

	'This information is passed from the primary
	
	public Sub show()
		MsgBox decryptedHash
		show = decryptedHash
	End Sub
	
	Public default function init(encryptedText, decryptKeyEnd)
		hashedPW = encryptedText
		endStrKey = decryptKeyEnd
		DecryptProcess(encryptedText, decryptKeyEnd)
	End Function
	
	Private Function DecryptProcess(hash, keyEnd)
		fullStryKey = strKey + keyEnd
		decryptedHash = Decrypt(hash, fullStryKey, intSeed)
	End Function
	
	Function DecryptPass(EncryptedText)
		On Error Resume Next
		DecryptPass = Decrypt(EncryptedText, strKey, intSeed)
	End Function
	
	
	
	
	Private Function String2Asc( strIn)
	  arrResult = Array()
	  ReDim arrResult( CInt( Len( strIn ) ) )
	  For intI = 0 to Len(strIn) - 1
		  arrResult( intI ) = Asc( Mid( strIn,intI + 1 ,1 ) )
	  Next
	  String2Asc = arrResult
	End Function
	
	Private Function Decrypt( strDecrypt, strKey, intSeed)
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

End Class



'******************************************************************************
'** Created:      20-1-2009 22:28
'** Original Author of Decrypt Functions:         Adriaan Westra
'******************************************************************************