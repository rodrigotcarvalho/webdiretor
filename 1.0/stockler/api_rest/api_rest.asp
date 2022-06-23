<!--#include file="aspJSON1.18.asp"-->
<!--#include file="../inc/funcoes_comuns.asp"-->
<%

Function ChamaAPI(P_Api_Rest, P_METODO, P_JSON)

	if P_Api_Rest<>"" then
	
		Dim oXMLHTTP
		'Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
		Set oXMLHTTP = Server.CreateObject("Microsoft.XMLHTTP")			
		oXMLHTTP.Open P_METODO, P_Api_Rest, False				
		oXMLHTTP.setRequestHeader "Content-Type", "application/json"
		oXMLHTTP.setRequestHeader "Accept", "application/json"	
		oXMLHTTP.SetRequestHeader "User-Agent", "ASP/3.0"
		oXMLHTTP.send P_JSON

		Response.AddHeader "Content-Type", "application/json;charset=UTF-8"
		Response.Charset = "UTF-8"		
		If oXMLHTTP.Status = 201 Then
		
			statusReturn = oXMLHTTP.Status
			pageReturn = oXMLHTTP.responseText
			Set oXMLHTTP = Nothing
			'response.Write "Status: " & statusReturn & " | "			
			'response.write pageReturn			
			ChamaAPI = pageReturn		
		Else

			statusReturn = oXMLHTTP.Status
			pageReturn = oXMLHTTP.responseText
			Set oXMLHTTP = Nothing
			'response.Write "Status: " & statusReturn & " | "			
			'response.write pageReturn	
			'response.end
		End If
	end if
End Function

%>