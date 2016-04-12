<!--#include file="soap_client.asp" -->
<%
Set SoapClient = New oSoapClient
SoapClient.BaseUrl = "http://data.kitapbilisim.com/data.asmx"

' For View
Response.ContentType = "text/xml;charset=ISO-8859-9"

'Response.Write SoapClient.Execute(Server.Mappath("kategori_listele.xml"), Null)
Response.Write SoapClient.Execute(Server.Mappath("kitap_listele.xml"), Array("firmaId", "117", "ip", "212.174.10.2", "tarih", "01/01/2016", "ilkIndeks", 0, "adet", 30))

Set SoapClient = Nothing
%>