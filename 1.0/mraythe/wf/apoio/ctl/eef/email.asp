<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes7.asp"-->
<%
ano_letivo_wf = session("ano_letivo_wf")
session("ano_letivo_wf")=ano_letivo_wf

chave=session("chave")
session("chave")=chave

	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1	
	
	Set CON2 = Server.CreateObject("ADODB.Connection") 
	ABRIR2 = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON2.Open ABRIR2		

	Set CON_WF = Server.CreateObject("ADODB.Connection") 
	ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_WF.Open ABRIR_WF		
	
	Set CON_wr = Server.CreateObject("ADODB.Connection") 
	ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_wr.Open ABRIR_wr
	

co_assunto = request.form("assunto")
cc = request.form("cc")
tipo = request.form("tipo")
mensagem = request.form("msg")
destinatarios = request.form("dest")
unidade = request.form("unidade")
curso = request.form("curso")
etapa = request.form("etapa")
turma = request.form("turma")

Set RSa = Server.CreateObject("ADODB.Recordset")
SQLa = "SELECT TX_Titulo_Assunto FROM TB_Email_Assunto where CO_Assunto = "&co_assunto
RSa.Open SQLa, CON0

assunto=RSa("TX_Titulo_Assunto")
if destinatarios="" or isnull(destinatarios) then
else
	destinatario=split(destinatarios,", ")	
	alunos_vetor=alunos_turma(ano_letivo_wf,unidade,curso,etapa,turma,"nome_ativo")	

	alunos=split(alunos_vetor,"#$#")
	total_email=0
	conta_destinatarios=0
	vetor_destinatarios=""
	for a=0 to ubound(alunos)
		aluno=split(alunos(a),"#!#")	
		co_matric=aluno(0)
		sql_aluno = "NULO"
		sql_resp = "NULO"		
		sql_contatos = "NULO"	
	
		Set RST = Server.CreateObject("ADODB.Recordset")
		CONEXAOT = "Select * from TB_Matriculas WHERE CO_Situacao = 'C' And NU_Ano="& SESSION("ano_letivo") &" and CO_Matricula = "&co_matric
		Set RST = CON1.Execute(CONEXAOT)
		
		if NOT RST.EOF then
			for d=0 to ubound(destinatario)	
	
				if destinatario(d)="a" then
					sql_aluno = " TP_Contato = 'ALUNO' "
	'				Set RSA = Server.CreateObject("ADODB.Recordset")
	'				sqlA = "select TX_EMail_Usuario,IN_Aut_email, ST_Usuario from TB_Usuario where CO_Usuario="&co_matric
	'				set RSA = CON_wf.Execute (sqlA)					
	'				
	'				if RSA.EOF then
	'				
	'				else	
	'					email_aluno = RSA("TX_EMail_Usuario")
	'					if isnull(email_aluno) or email_aluno="" then		
	'					else
	'						email_aluno=Replace(email_aluno, " ", "")	
	'					end if					
	'					aut_aluno=RSA("IN_Aut_email")	
	'					situacao_aluno = RSA("ST_Usuario")	
	'					if aut_aluno=TRUE and situacao_aluno="L" then
	'						if conta_destinatarios=0 then
	'							publico=email_aluno
	'							bloco=email_aluno
	'							conta_destinatarios=conta_destinatarios+1					
	'						else
	'							publico=publico&","&email_aluno
	'							if bloco="" then
	'								bloco=email_aluno
	'							else	
	'								bloco=bloco&","&email_aluno	
	'							end if	
	'							if conta_destinatarios>35 then
	'								if conta_destinatarios=total_email then
	'									vetor_destinatarios=bloco
	'								else
	'									vetor_destinatarios=vetor_destinatarios&"#!#"&bloco
	'								end if
	'								conta_destinatarios=0
	'								bloco=""
	'							else
	'								conta_destinatarios=conta_destinatarios+1								
	'							end if
	'						end if
	'						total_email=total_email+1	
	'					end if	
	'				end if	
				elseif  destinatario(d)="r" then

					Set RSF = Server.CreateObject("ADODB.Recordset")
					sqlF = "select CO_Usuario from TB_RespxAluno where CO_Aluno="&co_matric&" and (TP_Resp='F' or TP_Resp='P')"
					set RSF = CON_wf.Execute (sqlF)		
				
					Set RS = Server.CreateObject("ADODB.Recordset")
					SQL = "SELECT TP_Resp_Fin, TP_Resp_Ped FROM TB_Alunos where CO_Matricula = "&co_matric
					RS.Open SQL, CON1
				
					resp_fin  = RS("TP_Resp_Fin")
					resp_ped  = RS("TP_Resp_Ped")					


					sql_resp = " TP_Contato IN ('"&resp_fin&"', '"&resp_ped&"')"
	'				while not RSF.EOF 
	'					resp=RSF("CO_Usuario")
	'				
	'					Set RSFM = Server.CreateObject("ADODB.Recordset")
	'					sqlFM = "select TX_EMail_Usuario,IN_Aut_email, ST_Usuario from TB_Usuario where CO_Usuario="&resp
	'					set RSFM = CON_wf.Execute (sqlFM)					
	'	
	'					if RSFM.EOF then
	'					
	'					else		
	'						email_resp = RSFM("TX_EMail_Usuario")
	'						if isnull(email_resp) or email_resp="" then		
	'						else
	'							email_resp=Replace(email_resp, " ", "")	
	'						end if	
	'						aut_resp=RSFM("IN_Aut_email")
	'						situacao_resp=RSFM("ST_Usuario")										
	'						if aut_resp=TRUE and situacao_resp="L" then
	'							if total_email=0 then
	'								publico=email_resp
	'								bloco=email_resp
	'								conta_destinatarios=conta_destinatarios+1						
	'							else
	'								publico=publico&","&email_resp
	'								if bloco="" then
	'									bloco=email_resp
	'								else	
	'									bloco=bloco&","&email_resp	
	'								end if					
	'								if conta_destinatarios>35 then
	'									if conta_destinatarios=total_email then
	'										vetor_destinatarios=bloco
	'									else
	'										vetor_destinatarios=vetor_destinatarios&"#!#"&bloco
	'									end if
	'									conta_destinatarios=0
	'									bloco=""
	'								else
	'									conta_destinatarios=conta_destinatarios+1								
	'								end if
	'							end if
	'							total_email=total_email+1	
	'						end if	
	'					end if					
	'				RSF.MOVENEXT
	'				wend
				elseif  destinatario(d)="i" then
					sql_contatos = "TP_Contato IN ('PAI', 'MAE')"	

	'				Set RSP = Server.CreateObject("ADODB.Recordset")
	'				sqlP = "select CO_Usuario from TB_RespxAluno where CO_Aluno="&co_matric&" and TP_Resp='I'"
	'				set RSP = CON_wf.Execute (sqlP)	
	'
	'				if RSP.EOF then
	'		
	'				else
	'				
	'					contato=RSP("CO_Usuario")
	'		
	'					Set RSFP = Server.CreateObject("ADODB.Recordset")
	'					sqlFP = "select TX_EMail_Usuario,IN_Aut_email, ST_Usuario from TB_Usuario where CO_Usuario="&contato
	'					set RSFP = CON_wf.Execute (sqlFP)					
	'							
	'					if RSFP.EOF then
	'				
	'					else
	'						email_cont = RSFP("TX_EMail_Usuario")
	'						
	'						if isnull(email_cont) or email_cont="" then		
	'						else
	'							email_cont=Replace(email_cont, " ", "")	
	'						end if	
	'						
	'						aut_cont=RSFP("IN_Aut_email")
	'						situacao_cont = RSFP("ST_Usuario")											
	'						if aut_cont=TRUE and situacao_cont="L" then		
	'							if total_email=0 then
	'								publico=email_cont
	'								bloco=email_cont
	'								conta_destinatarios=conta_destinatarios+1						
	'							else
	'								publico=publico&","&email_cont
	'								if bloco="" then
	'									bloco=email_cont
	'								else	
	'									bloco=bloco&","&email_cont	
	'								end if	
	'								if conta_destinatarios>35 then
	'									if conta_destinatarios=total_email then
	'										vetor_destinatarios=bloco
	'									else
	'										vetor_destinatarios=vetor_destinatarios&"#!#"&bloco
	'									end if
	'									conta_destinatarios=0
	'									bloco=""
	'								else
	'									conta_destinatarios=conta_destinatarios+1								
	'								end if
	'							end if
	'							total_email=total_email+1	
	'						end if
	'					end if												
	'				end if
				end if	
			next
		
			sql_query=""
	
		
			if 	sql_aluno <> "NULO"	then			
				sql_query= sql_query&" AND "&sql_aluno		
			end if	
			if 	sql_resp <> "NULO"	then			
				sql_query= sql_query&" AND "&sql_resp		
			end if	
			if 	sql_contatos <> "NULO"	then			
				sql_query= sql_query&" AND "&sql_contatos		
			end if																										
					
			Set RSc = Server.CreateObject("ADODB.Recordset")
			SQLc = "SELECT DISTINCT(TX_EMail) FROM TB_Contatos where CO_Matricula = "&co_matric&sql_query
			'response.Write(SQLc&"<BR>")
			RSc.Open SQLc, CON2		
		
			while not RSc.EOF 
				email_contato=RSc("TX_EMail")
	'		response.Write(email_contato&"<BR>")				
				if isnull(email_contato) or email_contato="" or InStr(email_contato,"@")=0 then		
				else
					email_contato=Replace(email_contato, " ", "")				
					if total_email=0 then
						publico=email_cont
						bloco=email_cont
						conta_destinatarios=conta_destinatarios+1			
					else
						publico=publico&","&email_contato
						if bloco="" then
							bloco=email_contato
						else	
							bloco=bloco&","&email_contato	
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
			RSc.MOVENEXT
			wend							
		next
	end if
'		response.Write(total_email&"<BR>")	
'response.End()
	if total_email<=35 then
		vetor_destinatarios=bloco
	else
		vetor_destinatarios=vetor_destinatarios&"#!#"&bloco
	end if
' Envia mensagem para os alunos/responsáveis/contatos	

	publico=split(vetor_destinatarios,"#!#")
				'email_teste = "webdiretor@gmail.com,rodrigotcarvalho@gmail.com,osmarpio@globo.com,osmarpio@openlink.com.br,osmar@simplynet.com.br,suporte@simplynet.com.br"
	for e= 0 to ubound(publico)
		if e=0 then
			publico_temp=publico(e)
		else
			publico_temp=publico_temp&Chr(13)&publico(e)		
		end if		

'	response.End()


		response.Write("ERRO Bloco "&e&" da combinação U:"&unidade&" C:"&curso&" E:"&etapa&" T:"&turma&" Destinatarios: "&destinatarios&"<BR>")	
				
		Set objCDOSYSMail = Server.CreateObject("CDO.Message")
		Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration") 'objeto de configuração do CDO
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 90
		objCDOSYSCon.Fields.update
		Set objCDOSYSMail.Configuration = objCDOSYSCon
		objCDOSYSMail.From = email_suporte_escola
		objCDOSYSMail.To = email_suporte_escola
		'objCDOSYSMail.Cc = ""
		'objCDOSYSMail.Bcc = email_teste		
		objCDOSYSMail.Bcc = "webdiretor@gmail.com"	
		'objCDOSYSMail.Bcc = "osmarpio@openlink.com.br"
		objCDOSYSMail.Bcc = publico(e)
'		if Session("arquivos_anexados")<>"nulo" then
'			anexos=split(Session("arquivos_anexados"),"#!#")
'			for atch=0 to ubound(anexos)
'				objCDOSYSMail.AddAttachment CAMINHO_upload&anexos(atch)
'			next
'		end if			
		objCDOSYSMail.Subject = assunto
		objCDOSYSMail.TextBody = mensagem
		'objCDOSYSMail.HtmlBody = mensagem
		objCDOSYSMail.Send 
		Set objCDOSYSMail = Nothing
		Set objCDOSYSCon = Nothing
	next	
end if	
'response.End()
' Mensagem para coordenador
'publico_temp=email_teste
mensagem = "Foram enviadas "&total_email&" mensagens, para os endereços: "&publico_temp&", com a seguinte mensagem: "&Chr(13)&Chr(13)&mensagem&Chr(13)&Chr(13)&" Sistema Web Diretor"	


Set objCDOSYSMail = Server.CreateObject("CDO.Message")
Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration") 'objeto de configuração do CDO
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 90
objCDOSYSCon.Fields.update
Set objCDOSYSMail.Configuration = objCDOSYSCon
objCDOSYSMail.From = email_suporte_escola
objCDOSYSMail.To = cc
'objCDOSYSMail.BCC = "osmarpio@sopenlink.com.br"
objCDOSYSMail.BCC = "webdiretor@gmail.com"	
'if Session("arquivos_anexados")<>"nulo" then
'	anexos=split(Session("arquivos_anexados"),"#!#")
'	for atch=0 to ubound(anexos)
'		objCDOSYSMail.AddAttachment CAMINHO_upload&anexos(atch)
'	next
'end if			
objCDOSYSMail.Subject = assunto
objCDOSYSMail.TextBody = mensagem
'objCDOSYSMail.HtmlBody = mensagem
objCDOSYSMail.Send 
Set objCDOSYSMail = Nothing
Set objCDOSYSCon = Nothing

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