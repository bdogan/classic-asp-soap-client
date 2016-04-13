<!--#include file="soap_client.asp" -->
<%

Set SoapClient = New oSoapClient
SoapClient.BaseUrl = "http://data.kitapbilisim.com/data.asmx"

Function KitapBilisimError(pReason)
	Response.Write "<p>KitapBilisim xml hatasý: """ & pReason & """ <br/> Lütfen kitapbilisim.com ile iletiþime geçiniz</p>"
	Response.End
End Function

' Set Variables
Dim pFirmId : pFirmId = "117"
Dim pIpAddress : pIpAddress = "212.174.10.2"
Dim pDate : pDate = "12/03/2016"
Dim pOffset : pOffset = "0"
Dim pLimit : pLimit = "50"
Dim pBookListSoapFile : pBookListSoapFile = Server.Mappath("kitap_listele.xml")

' Get First Response
Dim pResponse : pResponse = SoapClient.Execute(pBookListSoapFile, Array("firmaId", pFirmId, "ip", pIpAddress, "tarih", pDate, "ilkIndeks", pOffset, "adet", pLimit))
Dim oBooksXml : Set oBooksXml = Server.CreateObject("Msxml2.DOMDocument")
oBooksXml.Async = False
oBooksXml.LoadXML pResponse
If (oBooksXml.parseError.errorCode <> 0) Then Set oBooksXml = Nothing : KitapBilisimError(oBooksXml.parseError.reason)

' Get Counts
Dim pTotalRecords : pTotalRecords = CLng(oBooksXml.getElementsByTagName("ToplamVeriAdet")(0).Text)
Dim pRecords : pRecords = CLng(oBooksXml.getElementsByTagName("GelenVeriAdet")(0).Text)
Dim pTotalFetchedRecords : pTotalFetchedRecords = 0 : pTotalFetchedRecords = pTotalFetchedRecords + pRecords

' Get other records
Dim oBookInnersXml : Set oBookInnersXml = Server.CreateObject("Msxml2.DOMDocument")
oBookInnersXml.Async = False
Do While (pTotalRecords > pTotalFetchedRecords) ' While all records to be fetched
	pOffset = pTotalFetchedRecords
	pResponse = SoapClient.Execute(pBookListSoapFile, Array("firmaId", pFirmId, "ip", pIpAddress, "tarih", pDate, "ilkIndeks", pOffset, "adet", pLimit))
	oBookInnersXml.LoadXML pResponse
	If (oBookInnersXml.parseError.errorCode <> 0) Then Set oBookInnersXml = Nothing : KitapBilisimError(oBookInnersXml.parseError.reason)
	pRecords = CLng(oBookInnersXml.getElementsByTagName("GelenVeriAdet")(0).Text)
	If (pRecords = 0 OR Err.Number <> 0 OR oBookInnersXml.getElementsByTagName("Kitap").length = 0) Then Exit Do
	' Append records
	Dim pBook
	For Each pBook in oBookInnersXml.getElementsByTagName("Kitap")
		oBooksXml.getElementsByTagName("Veriler")(0).appendChild(pBook)
	Next
	pTotalFetchedRecords = pTotalFetchedRecords + pRecords
Loop
Set oBookInnersXml = Nothing
Set SoapClient = Nothing
If (Err.Number <> 0) Then Set oBooksXml = Nothing : KitapBilisimError(Err.Description & " at line " & Err.Line)

oBooksXml.getElementsByTagName("GelenVeriAdet")(0).Text = Cstr(pTotalFetchedRecords)

' So Lets Begin
Dim kitaplar : Set kitaplar = oBooksXml.getElementsByTagName("Kitap")
Dim kitap
For Each kitap in kitaplar
	Dim KitapId : KitapId = CLng(kitap.getElementsByTagName("KitapId")(0).Text)
	Dim Ad : Ad = CStr(kitap.getElementsByTagName("Ad")(0).Text)
	Dim Aciklama : Aciklama = Cstr(kitap.getElementsByTagName("Aciklama")(0).Text)
	'Response.Write KitapId & " - " & Ad & "</br>"
	'Response.Write kitap.xml & "</br>"
Next
Response.ContentType = "text/xml;charset=ISO-8859-9"
Response.Write oBooksXml.xml

%>