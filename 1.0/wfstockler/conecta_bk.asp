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
			

		Set RS1 = Server.CreateObject("ADODB.Recordset")			
		consulta_prim = "select * from TB_Usuario where CO_Usuario = " & login 
		RS1.Open consulta_prim, conexao
		
		if RS1.eof and RS1.bof then 
			response.Redirect("default.asp?opt=04")
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
			
			teste_senha = isnumeric(senha)
			
			if teste_senha = TRUE then
				senha=senha*1		
			end if			
'=======================================================================================================================================================			
			if st_usuario="B" then
				response.Redirect("default.asp?opt=05")
			elseif st_usuario="T" then
				response.Redirect("check_acesso.asp?opt=b&lg="&login)	
			elseif senha_bd<> senha then
				response.Redirect("default.asp?opt=06")
			elseif codigo_seguranca ="" then
				response.Redirect("default.asp?opt=07")
			elseif codigo_seguranca <> texto_imagem then
				response.Redirect("default.asp?opt=08")
			else
				
				if acesso= 0 or login=senha then
					response.redirect ("primeiro.asp?opt=09")
				else
			
					acesso=acesso+1
		
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
	
	
					Set RSupdt = Server.CreateObject("ADODB.Recordset")				
					strSQL3= "UPDATE TB_Usuario SET NU_Acesso= "& acesso & ", HO_ult_Acesso = '"& horario & "', DA_Ult_Acesso = '"& data & "' WHERE CO_Usuario = "& session("co_user")
					RSupdt.Open strSQL3, conexao
	
					call GravaLog ("ENT",session("ano_letivo"))
				end if	
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