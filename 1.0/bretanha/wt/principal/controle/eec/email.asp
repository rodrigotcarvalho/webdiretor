<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes7.asp"-->
<%
ano_letivo = session("ano_letivo")
session("ano_letivo") = ano_letivo
co_usr = session("co_user")
nivel=4

chave=session("chave")
session("chave")=chave


ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
data = dia &"/"& mes &"/"& ano

	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CON7 = Server.CreateObject("ADODB.Connection") 
	ABRIR7 = "DBQ="& CAMINHO_pf & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON7.Open ABRIR7		
	
	Set CON8 = Server.CreateObject("ADODB.Connection") 
	ABRIR8 = "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON8.Open ABRIR8		
	
	Set RSe = Server.CreateObject("ADODB.Recordset")
	SQLe = "SELECT Login FROM TB_Operador"
	RSe.Open SQLe, CON8	
	
	from=RSe("Login")
	

		
ck_email = request.form("ck_email")
co_assunto	= request.form("tipo_email")


		Set RSa = Server.CreateObject("ADODB.Recordset")
		SQLa = "SELECT TX_Titulo_Assunto FROM TB_Email_Assunto where CO_Assunto = "&co_assunto
		RSa.Open SQLa, CON0
		
		assunto_padrao=RSa("TX_Titulo_Assunto")

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT TX_Conteudo_Email FROM TB_Email_Mensagem where CO_Email = "&co_assunto
		RS0.Open SQL0, CON0																						
	
		if RS0.EOF then
			mensagem_padrao = ""
		else
			mensagem_padrao = RS0("TX_Conteudo_Email")	
		end if	


	dados_form = split(ck_email,", ")
	
	for e = 0 to ubound(dados_form)
		dados_mensagem = split(dados_form(e),"#!#")
		co_matricula = dados_mensagem(0)	
		end_email=dados_mensagem(1)		
		resp_fin=dados_mensagem(2)
		nome_aluno = dados_mensagem(3)		
		meses = dados_mensagem(4)
		desc_meses = replace(dados_mensagem(5),"-",", ")					
		in_sexo=dados_mensagem(6)							
		
			
		assunto=assunto_padrao&" - Ref: "&	co_matricula
		
		mensagem = "Prezado(a) Senhor(a) "&resp_fin&Chr(13)&" Respons&aacute;vel "
	    if in_sexo = "F" then
			mensagem=mensagem&"pela aluna "
		else
			mensagem=mensagem&"pelo aluno "	
		end if	
		
		mensagem=mensagem&nome_aluno&", matr&iacute;cula "&co_matricula&Chr(13)&Chr(13)	
		
		If meses = 1 then
			mensagem=mensagem&"Parcela em aberto do m&ecirc;s de "&desc_meses&" "	
		else
			mensagem=mensagem&"Parcelas em aberto dos meses: "&desc_meses&" "			
		end if		
			
		mensagem=mensagem&"do ano letivo de "&session("ano_letivo")&Chr(13)&Chr(13)
		mensagem=Replace(mensagem&mensagem_padrao,Chr(13),"<BR>")
							
		Set objCDOSYSMail = Server.CreateObject("CDO.Message")
		Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration") 'objeto de configuração do CDO
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 90
		objCDOSYSCon.Fields.update
		Set objCDOSYSMail.Configuration = objCDOSYSCon
		objCDOSYSMail.From = from
		objCDOSYSMail.To = end_email
		'objCDOSYSMail.Cc = ""
		objCDOSYSMail.Bcc = from
		'objCDOSYSMail.Bcc = "osmarpio@sopenlink.com.br"
		'objCDOSYSMail.Bcc = publico(e)	
		objCDOSYSMail.Subject = assunto
		'objCDOSYSMail.TextBody = mensagem
		objCDOSYSMail.HtmlBody = mensagem
		objCDOSYSMail.Send 
		Set objCDOSYSMail = Nothing
		Set objCDOSYSCon = Nothing
		
		Set RSe = Server.CreateObject("ADODB.Recordset")				
		SQLe="SELECT CO_Matricula_Escola from TB_Email_Enviado where CO_Matricula_Escola = "&co_matricula
		RSe.Open SQLe, CON7	
		
		IF RSe.EOF then
			Set RS = server.createobject("adodb.recordset")			
			RS.open "TB_Email_Enviado", CON7, 2, 2 
			RS.addnew
			
				RS("CO_Matricula_Escola") = co_matricula
				RS("CO_Email") = co_assunto
				RS("DT_Envio") = data		
			RS.update
			set RS=nothing		
		
		else
			Set RS1 = Server.CreateObject("ADODB.Recordset")
			CONEXAO1 = "DELETE * from TB_Email_Enviado WHERE CO_Matricula_Escola = "& co_matricula 
			Set RS1 = CON7.Execute(CONEXAO1)
	
			Set RS = server.createobject("adodb.recordset")			
			RS.open "TB_Email_Enviado", CON7, 2, 2 
			RS.addnew
			
				RS("CO_Matricula_Escola") = co_matricula
				RS("CO_Email") = co_assunto
				RS("DT_Envio") = data		
			RS.update
			set RS=nothing			
		end if
				
	next	



call GravaLog (chave,outro)

response.Redirect("msgs.asp?opt=ok")
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