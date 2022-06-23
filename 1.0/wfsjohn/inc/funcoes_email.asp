<!--#include file="caminhos.asp"-->
<!--#include file="bd_bloqueto.asp"-->
<!--#include file="parametros.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="funcoes7.asp"-->
<%
Server.ScriptTimeout = 60 'valor em segundos

function email_anexo(p_vencimento, p_cod_aluno, dados)

e_vencimento = p_vencimento
e_cod = p_cod_aluno
ano_letivo = session("ano_letivo")

if dados = "S" or dados = "N" then
	eh_segunda_via = dados
else
	p_nosso_numero = dados
end if	

gerado =  GeraBloquetoNN(dados, e_cod, e_vencimento, "S", "T", P_NOSSO_NUMERO)

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
data = dia &"/"& mes &"/"& ano

chave=session("nvg")
session("nvg")=chave

	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CON2 = Server.CreateObject("ADODB.Connection") 
	ABRIR2 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON2.Open ABRIR2	
	
	Set CON6 = Server.CreateObject("ADODB.Connection") 
	ABRIR6 = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON6.Open ABRIR6		
	
	Set CON7 = Server.CreateObject("ADODB.Connection") 
	ABRIR7 = "DBQ="& CAMINHO_pf & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON7.Open ABRIR7		
	
	Set CON8 = Server.CreateObject("ADODB.Connection") 
	ABRIR8 = "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON8.Open ABRIR8	
	
	Set CON_wr = Server.CreateObject("ADODB.Connection") 
	ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_wr.Open ABRIR_wr			
	
	Set RSe = Server.CreateObject("ADODB.Recordset")
	SQLe = "SELECT Login FROM TB_Operador"
	RSe.Open SQLe, CON8	
	
	operador=RSe("Login")
	
'	Set RSem = Server.CreateObject("ADODB.Recordset")
'	SQLem = "select * from Email where CO_Escola="&session("escola")
'	set RSem = CON_wr.Execute (SQLem)
'
'	email_suporte=RSem("Suporte")	
	
from=email_financeiro



		Set RSa = Server.CreateObject("ADODB.Recordset")
		SQLa = "SELECT TB_Alunos.NO_Aluno, TB_Alunos.TP_Resp_Fin, TB_Alunos.IN_Sexo, TB_Matriculas.NU_Unidade, TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma FROM TB_Alunos, TB_Matriculas where TB_Matriculas.CO_Matricula = TB_Alunos.CO_Matricula AND TB_Matriculas.CO_Matricula = "& e_cod &" AND TB_Matriculas.NU_Ano = "& ano_letivo
		RSa.Open SQLa, CON2
		

		
		nome_aluno = RSa("NO_Aluno")
		tp_resp_fin = RSa("TP_Resp_Fin")
		in_sexo = RSa("IN_Sexo")		
		nu_unidade = RSa("NU_Unidade")
		co_curso = RSa("CO_Curso")
		co_etapa = RSa("CO_Etapa")
		co_turma = RSa("CO_Turma")
		
	
		
        Set RSc = Server.CreateObject("ADODB.Recordset")
		SQLc = "SELECT NO_Contato,CO_CPF_PFisica, TX_EMail FROM TB_Contatos where CO_Matricula = "& e_cod &" AND TP_Contato = '"& tp_resp_fin&"'"
		RSc.Open SQLc, CON6		
		
		If RSc.EOF then
			nome_resp ="Nome não cadastrado para o "&tp_resp_fin
		else
			nome_resp = RSc("NO_Contato")
			cpf_resp = RSc("CO_CPF_PFisica")
			email_resp =RSc("TX_EMail")
					
			if cpf_resp = "" or isnull(cpf_resp) then
			
			else
				cpf_resp = replace(cpf_resp,"-","")
				cpf_resp = replace(cpf_resp,".","")				
			end if
			
			if isnull(email_resp) or email_resp="" then
				email_resp ="Email não cadastrado"
			end if		
				
		end if		
	
		nome_meses=GeraNomesNovaVersao("MES",e_vencimento,variavel2,variavel3,variavel4,variavel5,CON0,outro)
		meses = 1
		

		ck_mail = e_cod&"#!#"&email_resp&"#!#"&nome_resp&"#!#"&nome_aluno&"#!#"&meses&"#!#"&nome_meses&"#!#"&in_sexo		

		Set RSa = Server.CreateObject("ADODB.Recordset")
		SQLa = "SELECT TX_Titulo_Assunto FROM TB_Email_Assunto where CO_Assunto = 20"
		RSa.Open SQLa, CON0
		
		assunto_padrao=RSa("TX_Titulo_Assunto")

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT TX_Conteudo_Email FROM TB_Email_Mensagem where CO_Email = 20"
		RS0.Open SQL0, CON0																						
	
		if RS0.EOF then
			mensagem_padrao = ""
		else
			mensagem_padrao = RS0("TX_Conteudo_Email")	
		end if	

	dados_form = split(ck_mail,", ")

	for e = 0 to ubound(dados_form)
		dados_mensagem = split(dados_form(e),"#!#")
		co_matricula = dados_mensagem(0)	
		end_email=dados_mensagem(1)		
		resp_fin=dados_mensagem(2)
		nome_aluno = dados_mensagem(3)		
		meses = dados_mensagem(4)
		desc_meses = replace(dados_mensagem(5),"-",", ")					
		in_sexo=dados_mensagem(6)							

		if InStr(end_email,"@")=0 then
		else 
		
		    assunto=assunto_padrao&" - Ref: "&	co_matricula
			
			
			if eh_segunda_via = "N" then
				assunto = Replace(assunto,"2ª Via de Bloqueto","Bloqueto")
			end if

			if Mid(ambiente_escola,1,5) = "teste"  then
				end_email = "osmarpio@openlink.com.br"	
				'end_email = "webdiretor@gmail.com"		
			    assunto="--- TESTE - favor desconsiderar ---"& assunto
				warning = "===================================================<br>TESTE - Favor desconsiderar<BR>=====================================================<BR>"
			end if
		
			

			
			mensagem = warning&"Prezado(a) Senhor(a) "&resp_fin&", respons&aacute;vel "
			
			
			if in_sexo = "F" then
				mensagem=mensagem&"pela aluna "
			else
				mensagem=mensagem&"pelo aluno "	
			end if	
			
			mensagem=mensagem&nome_aluno&", matr&iacute;cula "&co_matricula&Chr(13)&Chr(13)	
			
			if eh_segunda_via = "S" then
				mensagem=mensagem&"Segue em anexo a 2&ordf; Via do Bloqueto banc&aacute;rio da parcela referente ao m&ecirc;s de "&desc_meses&" do ano letivo de "&ano_letivo&Chr(13)&Chr(13)	
			else
				mensagem=mensagem&"Segue em anexo o Bloqueto banc&aacute;rio da parcela referente ao m&ecirc;s de "&desc_meses&" do ano letivo de "&ano_letivo&Chr(13)&Chr(13)	
			end if

			'mensagem=mensagem&"Qualquer d&uacute;vida por favor entre em contato com a secretaria da escola D&iacute;namis."&Chr(13)&Chr(13)	  
			mensagem=Replace(mensagem&mensagem_padrao,Chr(13),"<BR>")
			
			'response.write(from&"=>"&end_email&"<br>"&operador&"<br>")
			'end_email = "rodrigotcarvalho@gmail.com"
			'operador=""
			'response.write(co_matricula&" "&end_email&"=>"&operador&"<br>=====================================")			
								
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
			'objCDOSYSMail.Bcc = operador
			objCDOSYSMail.Subject = assunto
			'objCDOSYSMail.TextBody = mensagem
			objCDOSYSMail.HtmlBody = mensagem
			objCDOSYSMail.AddAttachment gerado
			objCDOSYSMail.Send 
			Set objCDOSYSMail = Nothing
			Set objCDOSYSCon = Nothing
			
			dim fs
			Set fs=Server.CreateObject("Scripting.FileSystemObject")
			if fs.FileExists(gerado) then
			  fs.DeleteFile(gerado)
			end if
			set fs=nothing

		end if		
	next	
	

email_anexo = "S"
end function
%>