<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/enviarEmail.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes7.asp"-->
<%
ano_letivo_wf = session("ano_letivo_wf")
session("ano_letivo_wf")=ano_letivo_wf

chave=session("chave")
session("chave")=chave

if transicao = "S" then
   url = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Web\wd\"&ambiente_escola&"\anexos\"
   url = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynetcloud1e1\Web\"&ambiente_escola&"\anexos\"
else
	if left(ambiente_escola,5) = "teste" then
		url = "E:\home\simplynet\Web\wdteste\"&ambiente_escola&"\anexos\"
	else
		url = "E:\home\wd.simplynet\Web\"&ambiente_escola&"\anexos\"
	end if	
end if	

	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0

	Set CON_WF = Server.CreateObject("ADODB.Connection") 
	ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_WF.Open ABRIR_WF		
	
	Set CON_wr = Server.CreateObject("ADODB.Connection") 
	ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_wr.Open ABRIR_wr
	
nome_escola=nome_da_escola
co_assunto = request.form("assunto")
cc = request.form("cc")
tipo = request.form("tipo")
mensagem = request.form("msg")
destinatarios = request.form("dest")
unidade = request.form("unidade")
curso = request.form("curso")
etapa = request.form("etapa")
turma = request.form("turma")
qtdAnexos = request.form("qtdAnexos")
if qtdAnexos>0 then
	for qtd = 1 to qtdAnexos
	     anexo = request.form("anexo"&qtd)
		if qtd = 1 then
			pAnexos = url&anexo
		else
			pAnexos = pAnexos&"#!#"&url&anexo		
		end if
	next
end if

Set RSa = Server.CreateObject("ADODB.Recordset")
SQLa = "SELECT TX_Titulo_Assunto FROM TB_Email_Assunto where CO_Assunto = "&co_assunto
RSa.Open SQLa, CON0

assunto=RSa("TX_Titulo_Assunto")
mensagem = "<htmL><head></head><body>"&mensagem&"</body></html>"
if destinatarios="" or isnull(destinatarios) then
else
	destinatario=split(destinatarios,", ")	
	if not isnumeric(curso) then
		if curso = "nulo" then
			curso = ""
		end if
	end if
	
	if not isnumeric(etapa) then
		if etapa = "nulo" then
			etapa = ""
		end if
	end if
	
	if not isnumeric(turma) then
		if turma = "nulo" then
			turma = ""
		end if
	end if		

	alunos_vetor=alunos_turma(ano_letivo_wf,unidade,curso,etapa,turma,"nome")	

	alunos=split(alunos_vetor,"#$#")
	total_email=0
	conta_destinatarios=0
	vetor_destinatarios=""
	for a=0 to ubound(alunos)
		aluno=split(alunos(a),"#!#")	
		co_matric=aluno(0)
		for d=0 to ubound(destinatario)		
			if destinatario(d)="a" then
				Set RSA = Server.CreateObject("ADODB.Recordset")
				sqlA = "select TX_EMail_Usuario,IN_Aut_email from TB_Usuario where CO_Usuario="&co_matric
				set RSA = CON_wf.Execute (sqlA)					
				
				if RSA.EOF then
				
				else	
					email_aluno = RSA("TX_EMail_Usuario")
					if isnull(email_aluno) or email_aluno="" then		
					else
						email_aluno=Replace(email_aluno, " ", "")	
					end if					
					aut_aluno=RSA("IN_Aut_email")	
					
					if aut_aluno=TRUE then
						if conta_destinatarios=0 then
							publico=email_aluno
							bloco=email_aluno
							conta_destinatarios=conta_destinatarios+1					
						else
							publico=publico&","&email_aluno
							if bloco="" then
								bloco=email_aluno
							else	
								bloco=bloco&","&email_aluno	
							end if	
							if conta_destinatarios>35 then
								if conta_destinatarios=total_email then
									vetor_destinatarios=bloco
								else
									vetor_destinatarios=vetor_destinatarios&"#!#"&bloco
								end if
								conta_destinatarios=0
								bloco=""
							else
								conta_destinatarios=conta_destinatarios+1								
							end if
						end if
						total_email=total_email+1	
					end if	
				end if	
			elseif  destinatario(d)="r" then

				Set RSF = Server.CreateObject("ADODB.Recordset")
				sqlF = "select CO_Usuario from TB_RespxAluno where CO_Aluno="&co_matric&" and (TP_Resp='F' or TP_Resp='P')"
				set RSF = CON_wf.Execute (sqlF)		

				if RSF.EOF then
				else
					resp=RSF("CO_Usuario")
				
					Set RSFM = Server.CreateObject("ADODB.Recordset")
					sqlFM = "select TX_EMail_Usuario,IN_Aut_email from TB_Usuario where CO_Usuario="&resp
					set RSFM = CON_wf.Execute (sqlFM)					
	
					if RSFM.EOF then
					
					else		
						email_resp = RSFM("TX_EMail_Usuario")
						if isnull(email_resp) or email_resp="" then		
						else
							email_resp=Replace(email_resp, " ", "")	
						end if	
						aut_resp=RSFM("IN_Aut_email")
						
						if aut_resp=TRUE then
							if total_email=0 then
								publico=email_resp
								bloco=email_resp
								conta_destinatarios=conta_destinatarios+1						
							else
								publico=publico&","&email_resp
								if bloco="" then
									bloco=email_resp
								else	
									bloco=bloco&","&email_resp	
								end if					
								if conta_destinatarios>35 then
									if conta_destinatarios=total_email then
										vetor_destinatarios=bloco
									else
										vetor_destinatarios=vetor_destinatarios&"#!#"&bloco
									end if
									conta_destinatarios=0
									bloco=""
								else
									conta_destinatarios=conta_destinatarios+1								
								end if
							end if
							total_email=total_email+1	
						end if	
					end if					
				end if
			elseif  destinatario(d)="i" then
			
				Set RSP = Server.CreateObject("ADODB.Recordset")
				sqlP = "select CO_Usuario from TB_RespxAluno where CO_Aluno="&co_matric&" and TP_Resp='I'"
				set RSP = CON_wf.Execute (sqlP)	

				if RSP.EOF then
				else
				
					contato=RSP("CO_Usuario")
					
					Set RSFP = Server.CreateObject("ADODB.Recordset")
					sqlFP = "select TX_EMail_Usuario,IN_Aut_email from TB_Usuario where CO_Usuario="&contato
					set RSFP = CON_wf.Execute (sqlFP)					
							
					if RSFP.EOF then
					
					else
						email_cont = RSFP("TX_EMail_Usuario")
						if isnull(email_cont) or email_cont="" then		
						else
							email_cont=Replace(email_cont, " ", "")	
						end if	
						
						aut_cont=RSFP("IN_Aut_email")
						
						if aut_cont=TRUE then
							if total_email=0 then
								publico=email_cont
								bloco=email_cont
								conta_destinatarios=conta_destinatarios+1						
							else
								publico=publico&","&email_cont
								if bloco="" then
									bloco=email_cont
								else	
									bloco=bloco&","&email_cont	
								end if	
								if conta_destinatarios>35 then
									if conta_destinatarios=total_email then
										vetor_destinatarios=bloco
									else
										vetor_destinatarios=vetor_destinatarios&"#!#"&bloco
									end if
									conta_destinatarios=0
									bloco=""
								else
									conta_destinatarios=conta_destinatarios+1								
								end if
							end if
							total_email=total_email+1	
						end if
					end if												
				end if
			end if	
		next
	next
	
	response.write(vetor_destinatarios)
	response.end()
	if total_email<=35 then
		vetor_destinatarios=bloco
	else
		vetor_destinatarios=vetor_destinatarios&"#!#"&bloco

		vetor_destinatarios = "webdiretor@gmail.com,rodrigotcarvalho@gmail.com,osmarpio@openlink.com.br,osmarpio@simplynet.com.br,osmarpio@gmail.com,osmar@simplynet.com.br,osmarpio@globo.com#!#webdiretor@gmail.com,rodrigotcarvalho@gmail.com,osmarpio@openlink.com.br,osmarpio@simplynet.com.br,osmarpio@gmail.com,osmar@simplynet.com.br,osmarpio@globo.com"	

		publico=split(vetor_destinatarios,"#!#")
		for e= 0 to ubound(publico)
			if e=0 then
				publico_temp=publico(e)
			else
				publico_temp=publico_temp&Chr(13)&publico(e)		
			end if
			'response.Write(publico(e))
			'response.Write("ERRO no Bloco "&e&" da combinação U:"&unidade&" C:"&curso&" E:"&etapa&" T:"&turma&"<BR>")				
			emailMsg=split(publico(e),",")
			for rec=0 to ubound(emailMsg) 
				'response.Write(emailMsg(rec)&"<BR>")
				'emailEnviado = enviaEmail(nome_da_escola, email_suporte_escola, nome_da_escola, emailMsg(rec), email_suporte_escola, "", "", assunto, mensagem,"S","N", pAnexos)
				if emailEnviado <>"S" then
					response.Write(emailEnviado)
					response.End()
				end if
			next
					   'enviaEmail(pNomeRemetente, pEmailRemetente, pNomeDestinatario, pEmailDestinatario, pEmailRetorno, pVetorCC, pVetorBcc, pAssunto, pMensagem, pMsgHtml, pAutentica, pEnderecoFisicoAnexo)	

		next	
	end if	
'response.End()
' Mensagem para coordenador

mensagem = "Foram enviadas "&total_email&" mensagens, para os endereços: "&publico_temp&", com a seguinte mensagem: "&Chr(13)&Chr(13)&mensagem&Chr(13)&Chr(13)&" Sistema Web Diretor"	

'	emailEnviado = enviaEmail(nome_da_escola, email_suporte_escola, nome_da_escola, cc, email_suporte_escola, "", "", assunto, mensagem, "S", "N", pAnexos)
	
	if emailEnviado <>"S" then
		response.Write(emailEnviado)
		response.End()
	end if	

outro=""

call GravaLog (chave,outro)

response.Redirect("index.asp?nvg=WF-AS-CO-EEF&opt=ok")
%>
<%If Err.number<>0 then
errnumb = Err.number
errdesc = Err.Description
lsPath = Request.ServerVariables("SCRIPT_NAME")
arPath = Split(lsPath, "/")
GetFileName =arPath(UBound(arPath,1))
passos = 0
for way=0 to UBound(arPath,1)
passos=passos+1
next
seleciona1=passos-2
pasta=arPath(seleciona)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>