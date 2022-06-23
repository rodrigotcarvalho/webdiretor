<%
function enviaEmail(pNomeRemetente, pEmailRemetente, pNomeDestinatario, pEmailDestinatario, pEmailRetorno, pVetorCC, pVetorBcc, pAssunto, pMensagem, pMsgHtml, pAutentica, pEnderecoFisicoAnexo)
	'Declaramos as váriaveis a serem utilizadas no script
	Dim AspEmail, nomeRemetente, emailRemetente, nomeDestinatario, emailDestinatario, emailRetorno, assunto, mensagem, servidor
	
	nomeRemetente=pNomeRemetente
	emailRemetente=pEmailRemetente 'O endereço de e-mail deve ser preenchido com uma conta existente em seu próprio domínio.
	nomeDestinatario=pNomeDestinatario
	emailDestinatario=pEmailDestinatario
	
	emailRetorno=pEmailRetorno
	responderPara=pEmailRetorno
	assunto = pAssunto
	mensagem=pMensagem
	servidor="localhost"
	
	'Agora configuramos o componente utilizando os dados informados nas variáveis
	 
	'Instancia o objeto na memória
	SET AspEmail = Server.CreateObject("Persits.MailSender")
	 
	'Contfigura o servidor SMTP a ser utilizado
	AspEmail.Host = servidor
	 
	'Configura o Nome do remetente da mensagem
	AspEmail.FromName = nomeRemetente
	 
	'Configura o e-mail do remetente da mensagem que OBRIGATORIAMENTE deve ser um e-mail do seu próprio domínio
	AspEmail.From = emailRemetente
	 
	'Configura o E-mail de retorno para você ser avisado em caso de problemas no envio da mensagem
	AspEmail.MailFrom = emailRemetente  'Desabilitar essa linha caso o servidor esteja configurado para rodar em 64 bits
	 
	 
	'Configura o e-mail que receberá as respostas desta mensagem
	AspEmail.AddReplyTo responderPara
			'response.Write(responderPara&"---------------------------------responderPara<BR>")		 
	'Configura os destinatários da mensagem
	AspEmail.AddAddress emailDestinatario, nomeDestinatario
			'response.Write(emailDestinatario&"---------------------------------emailDestinatario<BR>")		 
	'Configura para enviar e-mail Com Cópia
	'AspEmail.AddCC "nome0@dominio.com.br", "Nome"
	'AspEmail.AddCC "nome1@dominio.com.br", "Nome"
	'AspEmail.AddCC "nome2@dominio.com.br", "Nome"
	if not isnull(pVetorCC) then
	
		vetorCC = split(pVetorCC,"#!#")
		for c = 0 to ubound(vetorCC)
			 response.write(vetorCC(c)&"<BR>")
			nomeRemetente = split(vetorCC(c),"@")
			AspEmail.AddCC vetorCC(c), nomeRemetente(0)
		next
	end if		
	'response.Write("---------------------------------<BR>")	
	if not isnull(pVetorBcc) then
	
		vetorBCC = split(pVetorBcc,"#!#")
		for b = 0 to ubound(vetorBCC)
			nomeRemetente = split(vetorBCC(b),"@")
			 response.write(vetorBCC(b)&"<BR>")			
			AspEmail.AddBCC vetorBCC(b), nomeRemetente(0)
			'response.Write(vetorBCC(b))
		next
	end if	
	 'response.End()
	'Configura o Assunto da mensagem enviada
	AspEmail.Subject = assunto
	 
	'Configura o formato da mensagem para HTML
	if pMsgHtml = "S" then
		AspEmail.IsHTML = True
	end if
	 
	'Configura o conteúdo da Mensagem
	AspEmail.Body = mensagem
	 
	'Definir porta no caso de envio autenticado
	if pAutentica = "S" then
		AspEmail.Username = "tictictac@simplynet.com.br"
		AspEmail.Password = "Tictictac0"		
		'AspEmail.TLS = True
		AspEmail.Port = 25'587
	end if	
	 
	'Utilize este código caso queira enviar arquivo anexo
	if not isnull(pEnderecoFisicoAnexo) then
	
		vetorAnexos = split(pEnderecoFisicoAnexo,"#!#")
		for a = 0 to ubound(vetorAnexos)
			AspEmail.AddAttachment(vetorAnexos(a))
		next
	'	AspEmail.AddAttachment("E:\home\SEU_LOGIN_FTP\Web\caminho_do_arquivo")
	end if	
	 
	'Para quem utiliza serviços da REVENDA conosco
	'AspEmail.AddAttachment("E:\vhosts\DOMINIO_COMPLETO\httpdocs\caminho_do_arquivo")
	 
	'#Ativa o tratamento de erros
	On Error Resume Next
	 
	'Envia a mensagem
	AspEmail.Send

	'Caso ocorra problemas no envio, descreve os detalhes do mesmo.
	If Err <> 0 Then
		erro = "<b><font color='red'> Erro ao enviar a mensagem para "&emailDestinatario&".</font></b><br>"
		erro = erro & "<b>Erro.Description:</b> " & Err.Description & "<br>"
		erro = erro & "<b>Erro.Number:</b> "      & Err.Number & "<br>"
		erro = erro & "<b>Erro.Source:</b> "      & Err.Source & "<br>"
		'response.write erro
		enviaEmail = erro
		'response.End()
	Else
		'response.write "<font color='blue'><b>Mensagem enviada com sucesso para</b></font> " & emailDestinatario&"<BR>"	
		enviaEmail = "S"
	End If
	' response.End()
	'## Remove a referência do componente da memória ##
	SET AspEmail = Nothing
'response.Write("==========================================================================<BR>")	
end function
%>