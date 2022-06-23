<!--#include file="inc/caminhos.asp"-->
<!--#include file="inc/funcoes.asp"-->
<!--#include file="../global/funcoes_diversas.asp"-->
<%

opt=session("obr")
dados=split(opt,"$!$")
nova_senha=dados(0)
email=dados(1)
autorizo=dados(2)
data=replace(dados(3),"-","/")
acesso=dados(4)
horario=replace(dados(5),"-",":")
co_user=dados(6)


if left(ambiente_escola,5)= "teste" then
	area="wdteste"
	link="http://www.simplynet.com.br/"&area&"/"&ambiente_wf
else
	area="wd"
	link="http://www.simplynet.com.br/"&area&"/"&ambiente_wf
end if		


	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR

if session("tp")="R" then

	Set RS = Server.CreateObject("ADODB.Recordset")
'	SQL= "UPDATE TB_Usuario SET Senha= '"& nova_senha & "', TX_EMail_Usuario='"& email & "', IN_Aut_email="& autorizo & ", DA_Cadastro='"&data&"', NU_Acesso= "& acesso & ", HO_ult_Acesso = '"& horario & "', DA_Ult_Acesso = '"& data & "' WHERE CO_Usuario = "&co_user
	SQL= "UPDATE TB_Usuario SET Senha= '"& nova_senha & "', TX_EMail_Usuario='"& email & "', IN_Aut_email="& autorizo & ", DA_Cadastro='"&data&"', NU_Acesso= "& acesso & ", HO_ult_Acesso = '"& horario & "', DA_Ult_Acesso = '"& data & "', ST_Usuario='T' WHERE CO_Usuario = "&co_user
	
	RS.Open SQL, CON
	
						

	Set RSano = Server.CreateObject("ADODB.Recordset")
	SQLano = "SELECT * FROM TB_Ano_Letivo where ST_Ano_Letivo='L' order by NU_Ano_Letivo"
	RSano.Open SQLano, CON

	ano_letivo=RSano("NU_Ano_Letivo")

	
	call GravaLog ("ENT",ano_letivo)
	session("ano_letivo") = ano_letivo		
dados=co_user&"$!$"&email
		
dados=Base64Encode(dados)



	Set RS1 = Server.CreateObject("ADODB.Recordset")			
	SQL1 = "select * from TB_Usuario where CO_Usuario = " & co_user 
	RS1.Open SQL1, CON
	
	nome=RS1("NO_Usuario")	

	assunto=nome_escola&" - Confirmação de Acesso ao Web Família"

	mensagem="<font face=""Arial, Helvetica, sans-serif"" size=""2"">Prezado(a) usuário , "&nome&"<BR><BR>"
	mensagem=mensagem&"Seja bem vindo ao site Web Família do "&nome_escola&".<BR>"
	mensagem=mensagem&"Esperamos lhe fornecer durante o ano letivo várias informações importantes sobre o desenvolvimento escolar de seu filho.<br>Ao clicar no link abaixo o sistema procederá com a autorização de seu acesso, bastando realizar um novo login.<BR>"
	mensagem=mensagem&"Muito obrigado.<BR><BR>"
	mensagem=mensagem&"Para liberar o acesso ao sistema Web Família clique "
	 mensagem=mensagem&"<a href="""&link&"/check_acesso.asp?opt=l&dd="&dados&""">aqui</a>.</font>"
	
	Set objCDOSYSMail = Server.CreateObject("CDO.Message")
	Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration") 'objeto de configuração do CDO
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
	objCDOSYSCon.Fields.update
	Set objCDOSYSMail.Configuration = objCDOSYSCon
	objCDOSYSMail.From = mail_suporte
	objCDOSYSMail.to = 	email
	'objCDOSYSMail.Cc = ""	
	'objCDOSYSMail.Bcc = ""
'	if Session("arquivos_anexados")<>"nulo" then
'		anexos=split(Session("arquivos_anexados"),"#!#")
'		for atch=0 to ubound(anexos)
'			objCDOSYSMail.AddAttachment CAMINHO_upload&anexos(atch)
'		next
'	end if			
	objCDOSYSMail.Subject = assunto
	objCDOSYSMail.HtmlBody = mensagem
	objCDOSYSMail.Send 'envia o e-mail com o anexo
	Set objCDOSYSMail = Nothing
	Set objCDOSYSCon = Nothing

	response.redirect ("check_acesso.asp?opt=b&lg="&co_user)

else

	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL= "UPDATE TB_Usuario SET Senha= '"& nova_senha & "', TX_EMail_Usuario='"& email & "', IN_Aut_email="& autorizo & ", DA_Cadastro='"&data&"', NU_Acesso= "& acesso & ", HO_ult_Acesso = '"& horario & "', DA_Ult_Acesso = '"& data & "' WHERE CO_Usuario = "&co_user
	RS.Open SQL, CON
					
	Set RSano = Server.CreateObject("ADODB.Recordset")
	SQLano = "SELECT * FROM TB_Ano_Letivo where ST_Ano_Letivo='L' order by NU_Ano_Letivo"
	RSano.Open SQLano, CON

	ano_letivo=RSano("NU_Ano_Letivo")
	
	call GravaLog ("ENT",ano_letivo)
	session("ano_letivo") = ano_letivo	

	response.redirect ("inicio.asp?opt=ad")
end if						

%>