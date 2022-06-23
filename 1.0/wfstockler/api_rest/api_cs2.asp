<!--#include file="api_rest.asp"-->
<%
nomeApi = request.Form("API")
nomeDoc = request.Form("Doc")
documentB64 = request.Form("B64")
nomeRespFin = request.Form("nomeRespFin")
cpfRespFin = request.Form("cpfRespFin")
emailRespFin = request.Form("emailRespFin")
tp_cont = request.Form("tipo_contrato")


url="../rematricula/rem/index.asp?opt=ok&tp_contrato="&tp_cont

Function EnviaContratoClickSign(nomeApi, nomeDoc, documentB64, nomeRespFin, cpfRespFin, emailRespFin, tp_cont)


	keyDoc = FazUploadClickSign(nomeApi, nomeDoc, documentB64)
	
	emailRespFin = "osmarpio@globo.com"
	'emailRespFin = "rodrigotcarvalho@gmail.com"

	keySignatario1 = CriaSignatario(""&emailRespFin&"","","email",""&nomeRespFin&"", ""&cpfRespFin&"", "", true, "email")
	
	keyAssociacao1 = AssociacaoSignatarioDocumento(keyDoc, keySignatario1, "contractor", 1, "")
	
	'webdiretor@gmail.com
	keySignatario2 = CriaSignatario("osmarpio@simplynet.com.br","","email","Escola Dinamis", "", "", false, "email")
	
	keyAssociacao2 = AssociacaoSignatarioDocumento(keyDoc, keySignatario2, "contractee", 2, "")
	
	Notificado = NotificarAssinatura(keyAssociacao1, "Olá!"&chr(13)&chr(10)&"Segue documento para ser assinado eletronicamente através da Clicksign. Para assinar basta clicar no link abaixo e seguir o passo-a-passo."&chr(13)&chr(10)&"É bem simples!")

EnviaContratoClickSign = "S"
end Function


Function FazUploadClickSign(P_API, P_DOC, P_B64)
'P_DOC corresponde ao caminho a ser criado na Click Sign para mandar os arquivos

	host= "https://sandbox.clicksign.com"
	key = "ad294264-36dc-451b-a551-b129b4e8dbc4"

	validade = 90 'meses
		
	data_expiracao = SomaData(now, validade)
	
	anoExpiracao = DatePart("yyyy", data_expiracao) 
	
	mesExpiracao = DatePart("m", data_expiracao) 
	mesExpiracao = pad_zeros(mesExpiracao, 2)
	
	diaExpiracao = DatePart("d", data_expiracao) 
	diaExpiracao = pad_zeros(diaExpiracao, 2)
	
	dataExpiracao = anoExpiracao&"-"&mesExpiracao&"-"&diaExpiracao
	
	hora = DatePart("h", now) 
	min = DatePart("n", now)
	seg= DatePart("s", now) 

	if P_API = "postPDF" then
		B64 = "data:application/pdf;base64,"&P_B64
		JSON = ParseJsonPDF(P_DOC, B64, dataExpiracao&"T"&hora&":"&min&":"&seg&"-03:00")
		api_rest = host&"/api/v1/documents?access_token={"&key&"}"
		
	elseif P_API = "postDoc" then		
		B64 = "data:application/msword;base64,"&P_B64	
		
	elseif P_API = "criaSignatario" then		
		B64 = "data:application/msword;base64,"&P_B64			
	
	else
		B64 = "data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,"&P_B64		
					
	end if

	uploadDoc = ChamaAPI(api_rest, "POST", JSON)
	FazUploadClickSign = leKeyJson("document", "key", uploadDoc)			
	
end function

Function CriaSignatario(P_email,P_tel,P_auth,P_nome, P_CPF, P_datNasc, P_ObrigaCPF, P_Delivery)
'Formato Telefone = 11987654321
'Formato Data de Nascimento = 1983-03-31

	host= "https://sandbox.clicksign.com"
	key = "ad294264-36dc-451b-a551-b129b4e8dbc4"

		api_rest = host&"/api/v1/signers?access_token={"&key&"}"
		JSON = ParseJsonSignatario(P_email,P_tel,P_auth,P_nome, P_CPF, P_datNasc, P_ObrigaCPF, P_Delivery)	


	uploadDoc = ChamaAPI(api_rest, "POST", JSON)
	CriaSignatario = leKeyJson("signer", "key", uploadDoc)			
	
end function

Function AssociacaoSignatarioDocumento(P_key_doc, P_key_signer, P_sign_as, P_group, P_message)
	host= "https://sandbox.clicksign.com"
	key = "ad294264-36dc-451b-a551-b129b4e8dbc4"

		api_rest = host&"/api/v1/lists?access_token={"&key&"}"
		JSON = ParseJsonAssociacao(P_key_doc, P_key_signer, P_sign_as, P_group, P_message)	


	uploadDoc = ChamaAPI(api_rest, "POST", JSON)
	AssociacaoSignatarioDocumento = leKeyJson("list", "request_signature_key", uploadDoc)			

end function

Function NotificarAssinatura(P_key_doc, P_message)
	host= "https://sandbox.clicksign.com"
	key = "ad294264-36dc-451b-a551-b129b4e8dbc4"

		api_rest = host&"/api/v1/notifications?access_token={"&key&"}"
		JSON = ParseJsonNotificacao(P_key_doc, P_message)


	uploadDoc = ChamaAPI(api_rest, "POST", JSON)
	NotificarAssinatura = "S"	

end function


Function ParseJsonPDF(P_DOC, P_B64, P_DeadLine)

	Set oJSON = New aspJSON
	
	With oJSON.data
	
    .Add "document", oJSON.Collection()	
		 With oJSON.data("document")

			.Add "path", P_DOC                      
			.Add "content_base64", P_B64      
			.Add "deadline_at", P_DeadLine  
			.Add "auto_close", true    
			.Add "locale", "pt-BR"  
			.Add "sequence_enabled", true  
			.Add "remind_interval", 3

		End With	 			    
	End With
	
	ParseJsonPDF = oJSON.JSONoutput()

end function

Function ParseJsonSignatario(P_email,P_tel,P_auth,P_nome, P_CPF, P_datNasc, P_ObrigaCPF, P_Delivery)	

	Set oJSON = New aspJSON
	
	With oJSON.data
	
    .Add "signer", oJSON.Collection()	
		 With oJSON.data("signer")

			.Add "email", P_email                      
			.Add "phone_number", P_tel 
			.Add "auths", oJSON.Collection()
				With .item("auths")
					.Add 0, P_auth   
				End With          
			.Add "name", P_nome    
			.Add "documentation", P_CPF 
			.Add "birthday", P_datNasc
			.Add "has_documentation", P_ObrigaCPF 
			.Add "delivery", P_Delivery 
		End With	 			    
	End With


	
	ParseJsonSignatario = oJSON.JSONoutput()
end function

Function ParseJsonAssociacao(P_key_doc, P_key_signer, P_sign_as, P_group, P_message)

	Set oJSON = New aspJSON
	
	With oJSON.data
	
    .Add "list", oJSON.Collection()	
		 With oJSON.data("list")

			.Add "document_key", P_key_doc                      
			.Add "signer_key", P_key_signer      
			.Add "sign_as", P_sign_as  
			.Add "group", P_group    
			.Add "message", P_message  

		End With	 			    
	End With
	
	ParseJsonAssociacao = oJSON.JSONoutput()

end function

Function ParseJsonNotificacao(P_key_doc, P_message)

	Set oJSON = New aspJSON
	
	With oJSON.data

		.Add "request_signature_key", P_key_doc                      
		.Add "message", P_message  
					    
	End With
	
	ParseJsonNotificacao = oJSON.JSONoutput()

end function


Function leKeyJson(P_Master, P_key_a_buscar, P_JSON)
	Set oJSON = New aspJSON

	oJSON.loadJSON(P_JSON)
	
	leKeyJson = oJSON.data(""&P_Master&"")(""&P_key_a_buscar&"")
end function


'VERSÕES ORIGINAIS====================================================================================


Function PostData(link, data)
'https://stackoverflow.com/questions/17933301/how-to-send-a-http-post-from-classic-asp-with-a-parameter-to-a-web-api/28094656
	on error resume next
	if link<>"" then
		data = "{'Name': '" & data & "'}"
		data = Replace(data, "'", """")
		
		Dim oXMLHTTP
		Set oXMLHTTP = CreateObject("Msxml2.XMLHTTP.3.0")
		if oXMLHTTP is nothing then 
			Set oXMLHTTP = CreateObject("Microsoft.XMLHTTP")
			oXMLHTTP.Open "POST", link, False
			oXMLHTTP.setRequestHeader "Content-Type", "application/json"
			oXMLHTTP.send data
			
			If oXMLHTTP.Status = 200 Then
				PostData = oXMLHTTP.responseText
			Else
				response.Write "Status: " & oXMLHTTP.Status & " | "
				response.Write oXMLHTTP.responseText
				response.end
			End If
		End If
	end if
End Function

Function ParseJsonOriginal()
'https://www.aspjson.com/
'https://github.com/gerritvankuipers/aspjson
	Set oJSON = New aspJSON
	
	With oJSON.data

		.Add "familyName", "Smith"                      'Create value
		.Add "familyMembers", oJSON.Collection()
	
		With oJSON.data("familyMembers")
	
			.Add 0, oJSON.Collection()                  'Create unnamed object
			With .item(0)
				.Add "firstName", "John"
				.Add "age", 41
	
				.Add "job", oJSON.Collection()          'Create named object
				With .item("job")
					.Add "function", "Webdeveloper"
					.Add "salary", 70000
				End With
			End With
	
	
			.Add 1, oJSON.Collection()
			With .item(1)
				.Add "firstName", "Suzan"
				.Add "age", 38
				.Add "interests", oJSON.Collection()    'Create array
				With .item("interests")
					.Add 0, "Reading"
					.Add 1, "Tennis"
					.Add 2, "Painting"
				End With
			End With
	
			.Add 2, oJSON.Collection()
			With .item(2)
				.Add "firstName", "John Jr."
				.Add "age", 2.5
			End With
	
		End With

	End With

end function


Function leJsonOriginal()
'https://www.aspjson.com/
'https://github.com/gerritvankuipers/aspjson
Set oJSON = New aspJSON

'Load JSON string
oJSON.loadJSON(jsonstring)

'Get single value
Response.Write oJSON.data("firstName") & "<br>"

'Loop through collection
For Each phonenr In oJSON.data("phoneNumbers")
    Set this = oJSON.data("phoneNumbers").item(phonenr)
    Response.Write _
    this.item("type") & ": " & _
    this.item("number") & "<br>"
Next

'Update/Add value
oJSON.data("firstName") = "James"

'Return json string
Response.Write oJSON.JSONoutput()
end function


%>