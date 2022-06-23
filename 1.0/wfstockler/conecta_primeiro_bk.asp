<%'On Error Resume Next%>
<!--#include file="inc/caminhos.asp"-->
<%
session("senha") = ""
session("lg") = ""
session("ti") = ""
codigo_seguranca=session("codigo_seguranca")
login =request.form("login")
teste_login = isnumeric(login)
senha_form =request.form("senha_original")
senha_low=LCase(senha_form)
escola =request.form("escola")
texto_imagem =request.form("texto_imagem")
texto_imagem=LCase(texto_imagem)
pas1 =request.form("pas1")
pas1=LCase(pas1)
pas2 =request.form("pas2")
pas2=LCase(pas2)
autorizo =request.form("autorizo")
email =request.form("email")
session("senha") = senha_low
session("lg") = login
session("ti") = texto_imagem

erro=0

	Set conexao = Server.CreateObject("ADODB.Connection") 
	ABRIR= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
	conexao.Open ABRIR
	
	Set RS = Server.CreateObject("ADODB.Recordset")	
	consulta_ctl = "select * from TB_Controle"
	RS.Open consulta_ctl, conexao

	controle=RS("CO_controle")

if controle= "D" then
	response.Redirect("manutencao.asp")
end if

IF autorizo = "ok" then
	autorizo= TRUE
ELSE
	autorizo= FALSE
END IF


if login = "" then
	erro=1
elseif teste_login = false then
	erro=4
elseif senha_low = "" then
	erro=2
else

	Set RS1 = Server.CreateObject("ADODB.Recordset")			
	consulta_prim = "select * from TB_Usuario where CO_Usuario = " & login 
	RS1.Open consulta_prim, conexao

		
	if RS1.eof and RS1.bof then 
		erro=4
	else
		senha_bd=RS1("Senha")
		
		co_user= RS1("CO_Usuario")			
		nome_user= RS1("NO_Usuario")
		st_usuario = RS1("ST_Usuario")
		tp_usuario = RS1("TP_Usuario")
		acesso = RS1("NU_Acesso")		
		data_de = RS1("DA_Ult_Acesso")
		hora_de = RS1("HO_ult_Acesso")				

		
' alterado em 14/02/2011 para permitir que sejam comparadso os CFPs iniciados com zero com a senha iniciada com zero no primeiro acesso dos responsáveis			
		teste_senha_bd = isnumeric(senha_bd)
		
		if teste_senha_bd = TRUE then
			senha_bd=senha_bd*1
		end if	
		
		teste_senha = isnumeric(senha_low)
		
		if teste_senha = TRUE then
			senha_low=senha_low*1		
		end if			
	
'=======================================================================================================================================================					
		
		if st_usuario="B" then
			erro=5
		elseif senha_bd <> senha_low then
			erro=6				
		elseif codigo_seguranca ="" then
			session("senha") = senha_low
			session("pas1") = pas1
			session("pas2") = pas2
			session("mail_prim") = email		
			erro=7
		elseif codigo_seguranca <> texto_imagem then
			session("senha") = senha_low
			session("pas1") = pas1
			session("pas2") = pas2
			session("mail_prim") = email
			erro=8
		elseif isnull("pas1") or pas1="" THEN
			session("senha") = senha_low
			session("mail_prim") = email
			erro=10
		elseif isnull("pas2") or pas2="" THEN
			session("senha") = senha_low
			session("pas1") = pas1
			session("mail_prim") = email
			erro=11
		elseif pas=pas1 THEN
			session("senha") = senha_low
			session("mail_prim") = email
			erro=12
		elseif pas1<>pas2 THEN
			session("senha") = senha_low
			session("mail_prim") = email
			erro=13
		elseif texto_imagem = "" then
			session("senha") = senha_low
			session("pas1") = pas1
			session("pas2") = pas2
			session("mail_prim") = email
			erro=3
		elseif (email = "" or isnull(email)) and tp_usuario="R" then		
			session("senha") = senha_low
			session("pas1") = pas1
			session("pas2") = pas2
			erro=15			
		else
	
			if acesso>0 and senha_low<>login then
				erro=14
			else
				acesso = acesso + 1			
				session("tp") = tp_usuario
				session("nome") = nome_user
				session("login") = co_user
				session("acesso") = acesso		
				session("co_user") = co_user			
				session("sistema_local")="raiz"
				session("escola")=escola
				session("trava")="n"	
				session("dia_t") = data_inicio
				session("hora_t") = hora_de
				
			
				ano = DatePart("yyyy", now)
				mes = DatePart("m", now) 
				dia = DatePart("d", now) 
				hora = DatePart("h", now) 
				min = DatePart("n", now) 
	
				data = dia &"/"& mes &"/"& ano
				horario = hora & ":"& min
	
	
				data_replace = replace(data,"/","-")
				horario_replace = replace(data,":","-")
	
				session("obr")=pas1&"$!$"&email&"$!$"&autorizo&"$!$"&data_replace&"$!$"&acesso&"$!$"&horario&"$!$"&co_user		 		
		'response.End()
			END IF
		end if
	end if
end if	

if erro>0 then	
	response.Redirect("primeiro.asp?opt="&erro)
else
	response.Redirect ("atualiza_acesso.asp")
end if
%>