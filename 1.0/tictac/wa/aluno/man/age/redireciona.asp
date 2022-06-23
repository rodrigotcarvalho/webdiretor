<%
entrevista = request.Form("entrevista")
submit = request.Form("Submit")
cod  = request.querystring("opt")


if cod >0 then

	response.Redirect("incluir.asp?ori=I&opt="&cod)	
elseif submit = "Excluir" then 	
	response.Redirect("confirma.asp?opt="&entrevista)	
else
	' Caso o usuário selecione mais de uma entrevista, apenas a primeira é que poderá ser alterada ou ter o conteúdo incluído
	dados_entrevista = 	split(entrevista,",")
	entrevista_encaminhar = dados_entrevista(0)
	if submit = "Alterar" then 	
		response.Redirect("incluir.asp?ori=A&opt="&entrevista_encaminhar)
	elseif left(submit,5) = "Conte" then 	
		response.Redirect("incluir.asp?ori=C&opt="&entrevista_encaminhar)
	end if	
end if		
%>	