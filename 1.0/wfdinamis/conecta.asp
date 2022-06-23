<%'On Error Resume Next%>
<!--#include file="inc/caminhos.asp"-->
<!--#include file="inc/funcoes.asp"-->

<%
session("senha") = ""
session("lg") = ""
session("ti") = ""
'codigo_seguranca=session("codigo_seguranca")
codigo_seguranca=session("codigo_seguranca")
login =request.form("login")
senha =request.form("senha")
escola =request.form("escola")
texto_imagem =request.form("texto_imagem")




texto_imagem=LCase(texto_imagem)
logar =request.form("log")
senha=LCase(senha)
	
session("senha") = senha
session("lg") = login
session("ti") = texto_imagem

Set conexao_ctl = Server.CreateObject("ADODB.Connection") 
ABRIR_ctl = "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
conexao_ctl.Open ABRIR_ctl

consulta_ctl = "select * from TB_Controle"
set tabela_ctl = conexao_ctl.Execute (consulta_ctl)

controle=tabela_ctl("CO_controle")

if controle= "D" then
	response.Redirect("manutencao.asp")
end if

if logar= "on" then

	teste_login = isnumeric(login)

	if login = "" then
		response.Redirect("default.asp?opt=01")
	
	elseif teste_login = false then
		response.Redirect("default.asp?opt=04")
	elseif senha = "" then
		response.Redirect("default.asp?opt=02")
		elseif texto_imagem = "" then
		response.Redirect("default.asp?opt=03")
	else
		Set conexao = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		conexao.Open ABRIR
		
		
		consulta = "select * from TB_Usuario where CO_Usuario = " & login 
		set tabela = conexao.Execute (consulta)
		
		if tabela.eof and tabela.bof then 
			response.Redirect("default.asp?opt=04")
		else
			senha_bd=tabela("Senha")
			st_usuario = tabela("ST_Usuario")
			tp_usuario = tabela("TP_Usuario")

' alterado em 14/02/2011 para permitir que sejam comparadso os CFPs iniciados com zero com a senha iniciada com zero no primeiro acesso dos responsáveis			
			teste_senha_bd = isnumeric(senha_bd)
			
			if teste_senha_bd = TRUE then
				senha_bd=senha_bd*1
			end if	
			
			teste_senha = isnumeric(senha)
			
			if teste_senha = TRUE then
				senha=senha*1		
			end if		

			teste_login = isnumeric(login)
			
			if teste_login = TRUE then
				login=login*1		
			end if				
				
'=======================================================================================================================================================			
			if st_usuario="B" then
				response.Redirect("default.asp?opt=05")
			elseif senha_bd<> senha then
				response.Redirect("default.asp?opt=06")
			elseif codigo_seguranca ="" then
				response.Redirect("default.asp?opt=07")
			elseif codigo_seguranca <> texto_imagem then
				response.Redirect("default.asp?opt=08")
			else
				co_user= tabela("CO_Usuario")
				session("nome") = tabela("NO_Usuario")
				session("login") = tabela("CO_Usuario")
				acesso = tabela("NU_Acesso")
				session("tp") = tp_usuario		
				data_de = tabela("DA_Ult_Acesso")
				hora_de = tabela("HO_ult_Acesso")
				session("acesso") = acesso
				session("co_user") = co_user
				session("permissao") = permissao
				session("sistema_local")="raiz"
				session("escola")=escola
				session("trava")="n"	
				
				if session("acesso")= 0 or login=senha then
					response.redirect ("primeiro.asp?opt=09")
				end if
			
				acesso=acesso+1
					
				if data_de="" or isnull(data_de) then
				else			
					dados_dtd= split(data_de, "/" )
					dia_de= dados_dtd(0)
					mes_de= dados_dtd(1)
					ano_de= dados_dtd(2)
				end if
				
				if hora_de="" or isnull(hora_de) then
				else	
					dados_hrd= split(hora_de, ":" )
					h_de= dados_hrd(0)
					min_de= dados_hrd(1)
				end if
				if dia_de<10 then
					dia_de="0"&dia_de
				end if
				if mes_de<10 then
					mes_de="0"&mes_de
				end if
				if h_de<10 then
					h_de="0"&h_de
				end if
				hora_de=h_de&":"&min_de
				
				data_inicio=dia_de&"/"&mes_de&"/"&ano_de
							
				session("dia_t") = data_inicio
				session("hora_t") = hora_de			
				ano = DatePart("yyyy", now)
				mes = DatePart("m", now) 
				dia = DatePart("d", now) 
				hora = DatePart("h", now) 
				min = DatePart("n", now) 
				
	
				data = dia &"/"& mes &"/"& ano
				horario = hora & ":"& min


				Set RSano = Server.CreateObject("ADODB.Recordset")
				SQLano = "SELECT * FROM TB_Ano_Letivo where ST_Ano_Letivo='L' order by NU_Ano_Letivo"
				RSano.Open SQLano, conexao
		
				ano_letivo=RSano("NU_Ano_Letivo")
				session("ano_letivo") = ano_letivo


				Set conexao2 = Server.CreateObject("ADODB.Connection") 
				ABRIR2 = "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
				conexao2.Open ABRIR2
				
				strSQL3= "UPDATE TB_Usuario SET NU_Acesso= "& acesso & ", HO_ult_Acesso = '"& horario & "', DA_Ult_Acesso = '"& data & "' WHERE CO_Usuario = "& session("co_user")
				set tabela3 = conexao2.Execute (strSQL3)
	
				call GravaLog ("ENT",session("ano_letivo"))
				
			end if
		end if
	end if	
elseif logar="prim" then
	teste_login = isnumeric(login)
	pas1 =request.form("pas1")
	pas2 =request.form("pas2")
	autorizo =request.form("autorizo")
	email =request.form("email")
	
	pas1=LCase(pas1)
	pas2=LCase(pas2)
	
	IF autorizo = "ok" then
		autorizo= TRUE
	ELSE
		autorizo= FALSE
	END IF
	
	
	if login = "" then
		response.Redirect("primeiro.asp?opt=01")
	
	elseif teste_login = false then
		response.Redirect("primeiro.asp?opt=04")
	
	elseif senha = "" then
		response.Redirect("primeiro.asp?opt=02")
	else
		Set conexao = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		conexao.Open ABRIR
		
		consulta = "select * from TB_Usuario where CO_Usuario = " & login 
		set tabela = conexao.Execute (consulta)
		
		if tabela.eof and tabela.bof then 
			response.Redirect("primeiro.asp?opt=04")
		else
			senha_bd=tabela("Senha")
			st_usuario = tabela("ST_Usuario")
			tp_usuario = tabela("TP_Usuario")
			if st_usuario="B" then
				response.Redirect("primeiro.asp?opt=05")
			elseif senha_bd<> senha then
				response.Redirect("primeiro.asp?opt=06")
			elseif codigo_seguranca ="" then
				session("senha") = senha
				session("pas1") = pas1
				session("pas2") = pas2
				session("mail_prim") = email		
				response.Redirect("primeiro.asp?opt=07")
			elseif codigo_seguranca <> texto_imagem then
				session("senha") = senha
				session("pas1") = pas1
				session("pas2") = pas2
				session("mail_prim") = email
				response.Redirect("primeiro.asp?opt=08")
			elseif isnull("pas1") or pas1="" THEN
				session("senha") = senha
				session("mail_prim") = email
				response.Redirect("primeiro.asp?opt=10")
			elseif isnull("pas2") or pas2="" THEN
				session("senha") = senha
				session("pas1") = pas1
				session("mail_prim") = email
				response.Redirect("primeiro.asp?opt=11")
			elseif pas=pas1 THEN
				session("senha") = senha
				session("mail_prim") = email
				response.Redirect("primeiro.asp?opt=12")
			elseif pas1<>pas2 THEN
				session("senha") = senha
				session("mail_prim") = email
				response.Redirect("primeiro.asp?opt=13")
			elseif texto_imagem = "" then
				session("senha") = senha
				session("pas1") = pas1
				session("pas2") = pas2
				session("mail_prim") = email
				response.Redirect("primeiro.asp?opt=03")	
			else
				co_user= tabela("CO_Usuario")		
				session("nome") = tabela("NO_Usuario")
				session("login") = tabela("CO_Usuario")
				acesso = tabela("NU_Acesso")
				if acesso>0 and senha<>login then
					response.Redirect("primeiro.asp?opt=14")
				end if
						
				session("tp") = tp_usuario
				data_de = tabela("DA_Ult_Acesso")
				hora_de = tabela("HO_ult_Acesso")
				session("acesso") = acesso
				acesso = acesso + 1
				session("acesso") = acesso			
				session("co_user") = co_user			
				session("permissao") = permissao
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
	
		
				Set RSano = Server.CreateObject("ADODB.Recordset")
				SQLano = "SELECT * FROM TB_Ano_Letivo where ST_Ano_Letivo='L' order by NU_Ano_Letivo"
				RSano.Open SQLano, conexao
				ano_letivo=RSano("NU_Ano_Letivo")
				session("ano_letivo") = ano_letivo
	
				Set conexao2 = Server.CreateObject("ADODB.Connection") 
				ABRIR2 = "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
				conexao2.Open ABRIR2
				
				strSQL3= "UPDATE TB_Usuario SET Senha= '"& pas1 & "', TX_EMail_Usuario='"& email & "', IN_Aut_email="& autorizo & ", DA_Cadastro='"&data&"', NU_Acesso= "& acesso & ", HO_ult_Acesso = '"& horario & "', DA_Ult_Acesso = '"& data & "' WHERE CO_Usuario = "& session("co_user")
				set tabela3 = conexao2.Execute (strSQL3)
	
				call GravaLog ("ENT",session("ano_letivo"))
			end if
		end if
	end if	
end if

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
response.redirect("inc/erro.asp")
end if

if tp_usuario = "R" then
	response.redirect ("inicio.asp?opt=sa")
elseif tp_usuario="E" then
	response.Redirect("select_usr.asp?opt=off")
else
	response.redirect ("inicio.asp?opt=ad")
end if

%>
</html>