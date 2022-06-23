<%'On Error Resume Next%>
<!--#include file="inc/caminhos.asp"-->
<!--#include file="inc/funcoes.asp"-->

<%
session("senha") = ""
session("lg") = ""
session("ti") = ""
codigo_seguranca=session("codigo_seguranca")
login =request.form("login")
teste_login = isnumeric(login)
senha =request.form("senha")
senha=LCase(senha)
escola =request.form("escola")
texto_imagem =request.form("texto_imagem")
texto_imagem=LCase(texto_imagem)
pas1 =request.form("pas1")
pas1=LCase(pas1)
pas2 =request.form("pas2")
pas2=LCase(pas2)
autorizo =request.form("autorizo")
email =request.form("email")
session("senha") = senha
session("lg") = login
session("ti") = texto_imagem


	Set conexao_ctl = Server.CreateObject("ADODB.Connection") 
	ABRIR_ctl = "DBQ="& CAMINHO_wf&";Driver={Microsoft Access Driver (*.mdb)}"
	conexao_ctl.Open ABRIR_ctl
	
	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1		
	
	consulta_ctl = "select * from TB_Controle"
	set tabela_ctl = conexao_ctl.Execute (consulta_ctl)

	controle=tabela_ctl("CO_controle")

if controle= "D" then
	response.Redirect("manutencao.asp")
end if


	
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
		ABRIR = "DBQ="& CAMINHO_wf&";Driver={Microsoft Access Driver (*.mdb)}"
		conexao.Open ABRIR
		
		consulta_prim = "select * from TB_Usuario where CO_Usuario = " & login 
		set tabela_prim = conexao.Execute (consulta_prim)
		
		if tabela_prim.eof and tabela_prim.bof then 
			response.Redirect("primeiro.asp?opt=04")
		else
			senha_bd=tabela_prim("Senha")
			st_usuario = tabela_prim("ST_Usuario")
			
			if st_usuario="B" then
				response.Redirect("primeiro.asp?opt=05")
			elseif senha_bd <> senha then
				'response.Write(senha_bd &"<>"& senha&"ERRO")
				'response.end()
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
				co_user= tabela_prim("CO_Usuario")
			
				session("nome") = tabela_prim("NO_Usuario")
				session("login") = tabela_prim("CO_Usuario")
				acesso = tabela_prim("NU_Acesso")
		
				if acesso>0 and senha<>login then
					response.Redirect("primeiro.asp?opt=14")
				end if
				
				Set RSano = Server.CreateObject("ADODB.Recordset")
				SQLano = "SELECT * FROM TB_Ano_Letivo where ST_Ano_Letivo='L' order by NU_Ano_Letivo"
				RSano.Open SQLano, conexao
		
				ano_letivo=RSano("NU_Ano_Letivo")
				session("ano_letivo") = ano_letivo
				min_ano_letivo = ano_letivo
				session("ano_vigente") = ano_letivo		
					
				if tp_usuario = "R" then
					Set RS2 = Server.CreateObject("ADODB.Recordset")						
					SQL2 = "select * from TB_RespxAluno where CO_Usuario = " & co_user &" ORDER BY CO_Aluno"
					RS2.Open SQL2, conexao						
						
					while not RS2.EOF
						co_aluno=RS2("CO_Aluno")
						
						Set RS3 = Server.CreateObject("ADODB.Recordset")					
						SQL3 = "select MIN(NU_Ano) as min_ano_letivo, MAX(NU_Ano) as max_ano_letivo from TB_Matriculas where CO_Matricula = " & co_aluno 
						RS3.Open SQL3, CON1
						
						temp_min_ano_letivo=RS3("min_ano_letivo")
						temp_min_ano_letivo = temp_min_ano_letivo*1
						min_ano_letivo = min_ano_letivo*1
						if temp_min_ano_letivo<min_ano_letivo  then
							min_ano_letivo = temp_min_ano_letivo
						end if
						
						temp_max_ano_letivo=RS3("max_ano_letivo")
						temp_max_ano_letivo = temp_max_ano_letivo*1
						max_ano_letivo = max_ano_letivo*1
						if temp_max_ano_letivo> max_ano_letivo then
							max_ano_letivo = temp_max_ano_letivo
						end if	
					RS2.MOVENEXT
					Wend						
				elseif tp_usuario<>"E" then
					Set RS3 = Server.CreateObject("ADODB.Recordset")					
					SQL3 = "select MIN(NU_Ano) as min_ano_letivo, MAX(NU_Ano) as max_ano_letivo  from TB_Matriculas where CO_Matricula = " & co_user 
					RS3.Open SQL3, CON1					
					
					min_ano_letivo=RS3("min_ano_letivo")
					max_ano_letivo=RS3("max_ano_letivo")	
				end if				
					ano_letivo=ano_letivo*1	
					menor_ano_exibido = menor_ano_exibido*1				
					menor_ano_exibido = ano_letivo-anos_anteriores_exibidos

					if min_ano_letivo< menor_ano_exibido then
						min_ano_letivo = menor_ano_exibido
					end if	
					max_ano_letivo=max_ano_letivo*1
					ano_letivo=ano_letivo*1
					if max_ano_letivo< ano_letivo then
						session("ano_letivo") = max_ano_letivo
					end if					
										
					session("min_ano_letivo") = min_ano_letivo
					session("max_ano_letivo") = max_ano_letivo		
			
				session("tp") = tabela_prim("TP_Usuario")
				data_de = tabela_prim("DA_Ult_Acesso")
				hora_de = tabela_prim("HO_ult_Acesso")
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
	
				Set conexao2 = Server.CreateObject("ADODB.Connection") 
				ABRIR2 = "DBQ="& CAMINHO_wf&";Driver={Microsoft Access Driver (*.mdb)}"
				conexao2.Open ABRIR2
				
				strSQL3= "UPDATE TB_Usuario SET Senha= '"& pas1 & "', TX_EMail_Usuario='"& email & "', IN_Aut_email="& autorizo & ", DA_Cadastro='"&data&"', NU_Acesso= "& acesso & ", HO_ult_Acesso = '"& horario & "', DA_Ult_Acesso = '"& data & "' WHERE CO_Usuario = "& session("co_user")
				set tabela3 = conexao2.Execute (strSQL3)
		
				call GravaLog ("ENT",session("ano_letivo"))
			end if
		end if
	end if	


if session("tp")="R" then
response.redirect ("inicio.asp?opt=sa")
else
response.redirect ("inicio.asp?opt=ad")
end if

%>