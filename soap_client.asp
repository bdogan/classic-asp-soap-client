<%
Class oSoapClient

	Private Property Get FileGetContents(strFile)
		Dim pFSO : Set pFSO = Server.CreateObject("Scripting.FileSystemObject")
		Dim objFile : Set objFile = pFSO.OpenTextFile(strFile, 1)
		FileGetContents = objFile.ReadAll()
		Set objFile = Nothing
		If (CheckError) Then FileGetContents = False : Err.Clear
		Set pFSO = Nothing
	End Property

	Private Property Get StrReplace(pKeyValues, pStr)
		Dim pTarget, pKey, Cursor : Cursor = 0
		For Each pValue In pKeyValues
			If (Cursor Mod 2 = 0) Then 
				pKey = pValue
			Else
				pStr = Replace(pStr, "%" & pKey & "%", pValue, 1, -1, 0)
			End If
			Cursor = Cursor + 1
		Next
		StrReplace = pStr
	End Property

	Private Property Get Utf8Decode(byval UTF82TR_Data)
		If Len(UTF82TR_Data) = 0 Then Exit Property
		UTF82TR_Data = Replace(UTF82TR_Data ,"Ü","�",1,-1,0)
		UTF82TR_Data = Replace(UTF82TR_Data ,"Ç","�",1,-1,0)
		UTF82TR_Data = Replace(UTF82TR_Data ,"İ","�",1,-1,0)
		UTF82TR_Data = Replace(UTF82TR_Data ,"Ö","�",1,-1,0)
		UTF82TR_Data = Replace(UTF82TR_Data ,"ü","�",1,-1,0)
		UTF82TR_Data = Replace(UTF82TR_Data ,"ş","�",1,-1,0)
		UTF82TR_Data = Replace(UTF82TR_Data ,"ğ","�",1,-1,0)
		UTF82TR_Data = Replace(UTF82TR_Data ,"ç","�",1,-1,0)
		UTF82TR_Data = Replace(UTF82TR_Data ,"ı","�",1,-1,0)
		UTF82TR_Data = Replace(UTF82TR_Data ,"ö","�",1,-1,0)
		Utf8Decode = UTF82TR_Data
	End Property

	Private Property Get BinaryToString(Binary)
	  'Antonin Foller, http://www.motobit.com
	  'Optimized version of a simple BinaryToString algorithm.
	  Dim cl1, cl2, cl3, pl1, pl2, pl3
	  Dim L
	  cl1 = 1
	  cl2 = 1
	  cl3 = 1
	  L = LenB(Binary)
	  Do While cl1<=L
		pl3 = pl3 & Chr(AscB(MidB(Binary,cl1,1)))
		cl1 = cl1 + 1
		cl3 = cl3 + 1
		If cl3>300 Then
		  pl2 = pl2 & pl3
		  pl3 = ""
		  cl3 = 1
		  cl2 = cl2 + 1
		  If cl2>200 Then
			pl1 = pl1 & pl2
			pl2 = ""
			cl2 = 1
		  End If
		End If
	  Loop
	  BinaryToString = pl1 & pl2 & pl3
	End Property

	Private Property Get IsArray(ByVal pArr)
		IsArray = (TypeName(pArr) = "Variant()")
	End Property
	
	Public BaseUrl
	Public LastRequestRaw
	Public Property Get Execute(pFile, pKeyValues)
		If (NOT IsArray(pKeyValues)) Then pKeyValues = Array()
		Dim oXmlHTTP : Set oXmlHTTP = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
		oXmlHTTP.setTimeouts 100000, 100000, 200000, 200000
		oXmlHTTP.Open "POST", BaseUrl, False 
		oXmlHTTP.setRequestHeader "Content-Type", "text/xml;charset=ISO-8859-9" 
		LastRequestRaw = StrReplace(pKeyValues, FileGetContents(pFile))
		oXmlHTTP.send LastRequestRaw
		Execute = Utf8Decode(BinaryToString(oXmlHTTP.responseBody))
		Set oXmlHTTP = Nothing
	End Property

End Class
%>