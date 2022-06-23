<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes7.asp"-->
<%
ano_letivo_wf = session("ano_letivo_wf")
session("ano_letivo_wf")=ano_letivo_wf

chave=session("chave")
session("chave")=chave
'response.Write(Session("arquivos_anexados"))
'response.end()

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
	
nome_escola="Col&eacute;gio Stockler"
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
	alunos_vetor=alunos_turma(ano_letivo_wf,unidade,curso,etapa,turma,"nome")	
	
	alunos=split(alunos_vetor,"#$#")
	total_email=0
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
					email_aluno=RSA("TX_EMail_Usuario")
					aut_aluno=RSA("IN_Aut_email")	
					
					if aut_aluno=TRUE then
						if total_email=0 then
							publico=email_aluno
				
						else
							publico=publico&","&email_aluno
						end if
						total_email=total_email+1	
					end if	
				end if	
			elseif  destinatario(d)="r" then

				Set RSF = Server.CreateObject("ADODB.Recordset")
				sqlF = "select CO_Usuario from TB_RespxAluno where CO_Aluno="&co_matric&" and (TP_Resp='F' or TP_Resp='P')"
				set RSF = CON_wf.Execute (sqlF)		
				
				resp=RSF("CO_Usuario")
				
				Set RSFM = Server.CreateObject("ADODB.Recordset")
				sqlFM = "select TX_EMail_Usuario,IN_Aut_email from TB_Usuario where CO_Usuario="&resp
				set RSFM = CON_wf.Execute (sqlFM)					

				if RSFM.EOF then
				
				else						
					email_resp=RSFM("TX_EMail_Usuario")
					aut_resp=RSFM("IN_Aut_email")
					
					if aut_resp=TRUE then
						if total_email=0 then
							publico=email_resp
				
						else
							publico=publico&","&email_resp
						end if
						total_email=total_email+1	
					end if	
				end if					
			elseif  destinatario(d)="i" then
			
				Set RSP = Server.CreateObject("ADODB.Recordset")
				sqlP = "select CO_Usuario from TB_RespxAluno where CO_Aluno="&co_matric&" and TP_Resp='I'"
				set RSP = CON_wf.Execute (sqlP)	
				
				contato=RSP("CO_Usuario")
				
				Set RSFP = Server.CreateObject("ADODB.Recordset")
				sqlFP = "select TX_EMail_Usuario,IN_Aut_email from TB_Usuario where CO_Usuario="&contato
				set RSFP = CON_wf.Execute (sqlFP)					
						
				if RSFP.EOF then
				
				else
					email_cont=RSFP("TX_EMail_Usuario")
					aut_cont=RSFP("IN_Aut_email")
					
					if aut_cont=TRUE then
						if total_email=0 then
							publico=email_cont
				
						else
							publico=publico&","&email_cont
						end if
						total_email=total_email+1	
					end if
				end if												
			end if	
'			if destinatario(d)="a" then
'				usuario=co_matric
'			elseif  destinatario(d)="r" then
'
'				Set RSF = Server.CreateObject("ADODB.Recordset")
'				sqlF = "select CO_Usuario from TB_RespxAluno where CO_Aluno="&co_matric&" and (TP_Resp='F' or TP_Resp='P')"
'				set RSF = CON_wf.Execute (sqlF)		
'				
'				usuario=RSF("CO_Usuario")
'				
'			elseif  destinatario(d)="i" then
'			
'				Set RSP = Server.CreateObject("ADODB.Recordset")
'				sqlP = "select CO_Usuario from TB_RespxAluno where CO_Aluno="&co_matric&" and TP_Resp='I'"
'				set RSP = CON_wf.Execute (sqlP)	
'				
'				usuario=RSP("CO_Usuario")
'				
'			end if	
'			
'			Set RSd = Server.CreateObject("ADODB.Recordset")
'			sqld = "select NO_Usuario,TX_EMail_Usuario,IN_Aut_email from TB_Usuario where CO_Usuario="&usuario
'			set RSd = CON_wf.Execute (sqld)					
'			
'			if RSd.EOF then
'			
'			else
'				nome=RSd("NO_Usuario")					
'				email=RSd("TX_EMail_Usuario")
'				autorizacao=RSd("IN_Aut_email")	
'				
'
'				
'				if autorizacao=TRUE then
'					if total_email=0 then
'						publico=email
'			
'					else
'						publico=publico&", "&email
'					end if					
'				
'								
'					mensagem = "Prezado(a) "&nome&",<br><br>"&mensagem&"<br><br>"&nome_escola	
'		
'					Set objCDO = Server.CreateObject("CDONTS.NewMail")
'					objCDO.From = "suportewebdiretorstockler@webdiretor.com.br"
'					objCDO.To = email
'					' O 0 significa que o corpo da mensagem contém tags em HTML
'					' Para texto simples utiliza-se 1
'					objCDO.Bodyformat = 0
'					objCDO.MailFormat = 0
'					
'					objCDO.Subject = assunto
'					objCDO.Body = mensagem
'					objCDO.Send()
'					Set objCDO = Nothing					
'				
'				
'					total_email=total_email+1	
'				end if	
'			end if				
		next
	next
	publico_temp=publico
	'publico="webdiretor@gmail.com"	
	'publico=cc
' Envia mensagem para os alunos/responsáveis/contatos	
'	Set objCDO = Server.CreateObject("CDONTS.NewMail")
'	objCDO.From = "suportewebdiretorstockler@webdiretor.com.br"
'	'objCDO.To = cc
'	objCDO.Bcc=publico
'	
'	' O 0 significa que o corpo da mensagem contém tags em HTML
'	' Para texto simples utiliza-se 1
'	objCDO.Bodyformat = 0
'	objCDO.MailFormat = 0
'	
'	objCDO.Subject = assunto
'	objCDO.Body = mensagem
'	if Session("arquivos_anexados")<>"nulo" then
'		anexos=split(Session("arquivos_anexados"),"#!#")
'		for atch=0 to ubound(anexos)
'			objCDO.AttachFile CAMINHO_upload&"\"&anexos(atch)
'		next
'	end if	
'	objCDO.Send()
'	Set objCDO = Nothing

Set objCDOSYSMail = Server.CreateObject("CDO.Message")
Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration") 'objeto de configuração do CDO
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
objCDOSYSCon.Fields.update
Set objCDOSYSMail.Configuration = objCDOSYSCon
objCDOSYSMail.From = "suportewebdiretorstockler@webdiretor.com.br"
'objCDOSYSMail.Cc = "webdiretor@gmail.com"	
objCDOSYSMail.Bcc = publico
if Session("arquivos_anexados")<>"nulo" then
	anexos=split(Session("arquivos_anexados"),"#!#")
	for atch=0 to ubound(anexos)
		objCDOSYSMail.AddAttachment CAMINHO_upload&anexos(atch)
	next
end if			
objCDOSYSMail.Subject = assunto
objCDOSYSMail.HtmlBody = mensagem
objCDOSYSMail.Send 'envia o e-mail com o anexo
Set objCDOSYSMail = Nothing
Set objCDOSYSCon = Nothing

end if	
'cc="webdiretor@gmail.com"
publico=publico_temp
' Mensagem para coordenador
mensagem = "Foram enviadas "&total_email&" mensagens, para os endereços: "&publico&", com a seguinte mensagem:<br><br>"&mensagem&"<br><br>Sistema Web Diretor"	

'Set objCDO = Server.CreateObject("CDONTS.NewMail")
'objCDO.From = "suportewebdiretorstockler@webdiretor.com.br"
'objCDO.To = cc
'' O 0 significa que o corpo da mensagem contém tags em HTML
'' Para texto simples utiliza-se 1
'objCDO.Bodyformat = 0
'objCDO.MailFormat = 0
'
'objCDO.Subject = assunto
'objCDO.Body = mensagem
'
'if Session("arquivos_anexados")<>"nulo" then
'	anexos=split(Session("arquivos_anexados"),"#!#")
'	for atch=0 to ubound(anexos)
'		objCDO.AttachFile CAMINHO_upload&anexos(atch)
'	next
'end if	
'objCDO.Send()
'Set objCDO = Nothing

Set objCDOSYSMail = Server.CreateObject("CDO.Message")
Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration") 'objeto de configuração do CDO
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
objCDOSYSCon.Fields.update
Set objCDOSYSMail.Configuration = objCDOSYSCon
objCDOSYSMail.From = "suportewebdiretorstockler@webdiretor.com.br"
objCDOSYSMail.To = cc
if Session("arquivos_anexados")<>"nulo" then
	anexos=split(Session("arquivos_anexados"),"#!#")
	for atch=0 to ubound(anexos)
		objCDOSYSMail.AddAttachment CAMINHO_upload&anexos(atch)
	next
end if			
objCDOSYSMail.Subject = assunto
objCDOSYSMail.HtmlBody = mensagem
objCDOSYSMail.Send 'envia o e-mail com o anexo
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