<%'On Error Resume Next%>
<!--#include file="inc/caminhos.asp"-->
<!--#include file="inc/funcoes.asp"-->
<%
session("senha") = ""
session("lg") = ""
session("ti") = ""
codigo_seguranca=session("codigo_seguranca")
login =request.form("login")
senha =request.form("senha")
escola =request.form("escola")
texto_imagem =request.form("texto_imagem")
texto_imagem=LCase(texto_imagem)
session("senha") = senha
session("lg") = login
session("ti") = texto_imagem


Set conexao_ctl = Server.CreateObject("ADODB.Connection") 
	ABRIR_ctl = "DBQ="& CAMINHOctl & ";Driver={Microsoft Access Driver (*.mdb)}"
	conexao_ctl.Open ABRIR_ctl
	
	consulta_ctl = "select * from TB_Controle"
	set tabela_ctl = conexao_ctl.Execute (consulta_ctl)

controle=tabela_ctl("CO_controle")

if controle= "D" then
response.Redirect("manutencao.asp")
end if

if login = "" then
response.Redirect("default.asp?opt=01")
elseif senha = "" then
response.Redirect("default.asp?opt=02")
elseif texto_imagem = "" then
response.Redirect("default.asp?opt=03")
else
	Set conexao = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	conexao.Open ABRIR
	
	consulta = "select * from TB_Usuario where Login = '" & login & "'"
	set tabela = conexao.Execute (consulta)
	
	if tabela.eof and tabela.bof then 
		response.Redirect("default.asp?opt=04")
	else
		senha_bd=tabela("Senha")
		st_usuario = tabela("ST_Usuario")
		
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
		
	consulta1 = "select * from TB_Autoriz_Usuario_Grupo where CO_Usuario = " & co_user 
	set tabela1 = conexao.Execute (consulta1)
		
		grupo = tabela1("CO_Grupo")
		
	consulta2 = "select * from TB_Autoriz_Grupo where CO_Grupo = '" & grupo & "'"
	set tabela2 = conexao.Execute (consulta2)
	
	permissao= tabela2("Permissao")
		
		
			
			session("nome") = tabela("NO_Usuario")
			session("login") = tabela("Login")
			session("tp") = tp
			session("grupo") = grupo
			acesso = tabela("NU_Acesso")
			session("dia_t") = tabela("DA_Ult_Acesso")
			session("hora_t") = tabela("HO_ult_Acesso")
			acesso = acesso + 1
			session("acesso") = acesso
			session("co_user") = co_user
			session("permissao") = permissao
			session("sistema_local")="raiz"
			session("escola")=escola
			session("trava")="n"		
			ano = DatePart("yyyy", now)
			mes = DatePart("m", now) 
			dia = DatePart("d", now) 
			hora = DatePart("h", now) 
			min = DatePart("n", now) 

			data = dia &"/"& mes &"/"& ano
			horario = hora & ":"& min


		Set RSano = Server.CreateObject("ADODB.Recordset")
		SQLano = "SELECT * FROM TB_Ano_Letivo where ST_Ano_Letivo='L'"
		RSano.Open SQLano, conexao

if RSano.EOF then
		Set RSano = Server.CreateObject("ADODB.Recordset")
		SQLano = "SELECT MAX(NU_Ano_Letivo) AS ano_letivo FROM TB_Ano_Letivo"
		RSano.Open SQLano, conexao
		
ano_letivo=RSano("ano_letivo")
session("ano_letivo") = ano_letivo
else
ano_letivo=RSano("NU_Ano_Letivo")
session("ano_letivo") = ano_letivo
end if

	Set con_wf = Server.CreateObject("ADODB.Connection") 
	ABRIR_wf = "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
	con_wf.Open ABRIR_wf
	
		Set RSanowf = Server.CreateObject("ADODB.Recordset")
		SQLanowf = "SELECT * FROM TB_Ano_Letivo where ST_Ano_Letivo='L'"
		RSanowf.Open SQLanowf, con_wf

if RSano.EOF then
		Set RSanowf = Server.CreateObject("ADODB.Recordset")
		SQLanowf = "SELECT MAX(NU_Ano_Letivo) AS ano_letivo FROM TB_Ano_Letivo"
		RSanowf.Open SQLano, SQLanowf
		
ano_letivo_wf=RSanowf("ano_letivo")
session("ano_letivo_wf") = ano_letivo_wf
else
ano_letivo_wf=RSanowf("NU_Ano_Letivo")
session("ano_letivo_wf") = ano_letivo_wf
end if


	Set conexao2 = Server.CreateObject("ADODB.Connection") 
	ABRIR2 = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	conexao2.Open ABRIR2
	
			strSQL3= "UPDATE TB_Usuario SET NU_Acesso= "& acesso & ", HO_ult_Acesso = '"& horario & "', DA_Ult_Acesso = '"& data & "' WHERE CO_Usuario = "& session("co_user")
			set tabela3 = conexao2.Execute (strSQL3)
	
			call GravaLog ("WR-AUT-AUT-LOGWR",session("ano_letivo"))
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

response.redirect ("inicio.asp")
%>