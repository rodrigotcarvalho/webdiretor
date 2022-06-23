<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<%
chave=session("nvg")
session("nvg")=chave

opt = request.QueryString("opt")










		Set CON_o = Server.CreateObject("ADODB.Connection") 
		ABRIR_o = "DBQ="& CAMINHO_o & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_o.Open ABRIR_o


ano_letivo = session("ano_letivo")
co_usr = session("co_user")



if opt="exc" then
	exclui_ocorrencia=request.form("exclui_ocorrencia")
				
	vertorExclui = split(exclui_ocorrencia,", ")
	for i =0 to ubound(vertorExclui)
	
		exclui = split(vertorExclui(i),"?")
		
		cod = exclui(0)
		da_ocorrencia= exclui(1)
		ho_ocorrencia= exclui(2)
		co_ocorrencia= exclui(3)
		assunto=exclui(4)
		
		data_log=da_ocorrencia				
		dados_data=split(da_ocorrencia,"/")
		dia=dados_data(0)
		mes=dados_data(1)
		ano=dados_data(2)
		
		da_ocorrencia_cons=mes&"/"&dia&"/"&ano
		
		response.Write(">>"&ho_ocorrencia)
	
			Set RS = Server.CreateObject("ADODB.Recordset")
			SQL = "DELETE * from TB_Ocorrencia_Aluno where CO_Matricula ="& cod&" AND CO_Assunto = '"&assunto &"' AND CO_Ocorrencia ="& co_ocorrencia &" AND DA_Ocorrencia= #"&da_ocorrencia_cons&"# AND mid(HO_Ocorrencia,1,16)=#12/30/1899 "&ho_ocorrencia&"#"
			RS.Open SQL, CON_o
	
	next
	obr=session("obr")
	session("obr")=obr
	
	
	outro= "Excluir,"&data_log&","&cod
	call GravaLog (chave,outro)
	response.redirect("resumo.asp?or=2&opt=ok3")

elseif opt="inc" then


	cod= request.form("cod")
	
	tp_ocor=request.form("tp_ocor")
	assunto=request.form("assunto")
	disciplina=request.form("disciplina")
	aula=request.form("aula")
	co_prof=request.form("co_prof")
	data_inclui= request.form("data")
	data_altera= request.form("data_altera")
	observacao= request.form("observacao")
	
	dia_de= request.form("dia_de")
	mes_de= request.form("mes_de")
	ano_de= request.form("ano_de")
	data_inclui=dia_de&"/"&mes_de&"/"&ano_de
	data_busca=mes_de&"/"&dia_de&"/"&ano_de
	hora_ate= request.form("hora_ate")
	min_ate= request.form("min_ate")
	hora=hora_ate&":"&min_ate
	
	IF co_prof="999999" THEN
	co_prof=NULL
	END IF
	
	IF disciplina="999999" THEN
	disciplina=""
	END IF
	
	'response.Write("<BR>'"&hora_ate&"'")
	'response.Write("<BR>'"&min_ate&"'")
	'
	'response.Write("<BR>'"&cod&"'")
	'response.Write("<BR>'"&data_inclui&"'")
	'response.Write("<BR>'"&hora_ate&"'")
	'response.Write("<BR>'"&assunto&"'")
	'response.Write("<BR>'"&tp_ocor&"'")
	'response.Write("<BR>'"&aula&"'")
	'response.Write("<BR>'"&co_prof&"'")
	'response.Write("<BR>'"&disciplina&"'")
	'response.Write("<BR>'"&observacao&"'")
	'response.Write("<BR>'"&co_usr&"'")
	'response.Write("<BR>'"&hora&"'")
	'response.end()
		Set RSo = server.createobject("adodb.recordset")
		SQLo = "SELECT * FROM TB_Ocorrencia_Aluno WHERE CO_Matricula ="&cod&" AND DA_Ocorrencia = #"&data_busca&"# AND Right(HO_Ocorrencia,8) = #"&hora&"# AND CO_Assunto='"&assunto&"' AND CO_Ocorrencia ="& tp_ocor
		RSo.Open SQLo, CON_o

	IF RSo.EOF then
		Set RS = server.createobject("adodb.recordset")
		RS.open "TB_Ocorrencia_Aluno", CON_o, 2, 2 'which table do you want open
		RS.addnew
		  RS("CO_Matricula") = cod
		  RS("DA_Ocorrencia") = data_inclui
		  RS("HO_Ocorrencia") = hora
		  RS("CO_Assunto") = assunto
		  RS("CO_Ocorrencia") = tp_ocor
		  RS("NU_Aula") = aula
		  RS("CO_Professor") = co_prof
		  RS("NO_Materia") = disciplina
		  RS("TX_Observa") = observacao
		  RS("CO_Usuario")= co_usr
		  RS.update
		  
		set RS=nothing
		outro= "Incluir,"&data_altera&","&cod
		call GravaLog (chave,outro)
		obr=cod&"?dt?999999?1/1/"&ano_letivo&"?0:0?1/1/"&ano_letivo&", 0:0?12/31/"&ano_letivo&"?23:59?31/12/"&ano_letivo&", 23:59"
		session("obr")=obr
		'response.redirect("resumo.asp?cod="&cod&"&or=2&opt=ok1")
		cod_url=Replace(cod, " ", "")		
		session("cod_url") = cod_url	
		data_inclui=Replace(data_inclui, "/", "$$$")		
		response.redirect("email_responsaveis.asp?opt=sim&ori=i&obr="&obr&"&opt_a=ok1&dt="&data_inclui)				
		'response.redirect("email_responsaveis.asp?opt=ask&ori=i&vt=s&obr="&obr&"&opt_a=ok1&dt="&data_inclui)			
	else
		response.redirect("inclui.asp?cod="&cod&"&opt=err1")
	end if

elseif opt="multi" then


	cod= request.form("cod")
	
	tp_ocor=request.form("tp_ocor")
	assunto=request.form("assunto")
	disciplina=request.form("disciplina")
	aula=request.form("aula")
	co_prof=request.form("co_prof")
	data_altera= request.form("data_altera")
	hora_altera= request.form("hora_altera")
	observacao= request.form("observacao")
	obr= request.form("obr")
	obr_split=split(obr,"$!$")
	unidade=obr_split(0)
	curso=obr_split(1)
	co_etapa=obr_split(2)
	turma=obr_split(3)
	
	session("un_compara_aoc") = unidade
	session("cs_compara_aoc") = curso
	session("et_compara_aoc") = co_etapa	
	Session("co_ocor_aoc") = tp_ocor
	Session("prof_aoc") = co_prof
	Session("co_materia_aoc") = disciplina 
	
	dia_de= request.form("dia_de")
	mes_de= request.form("mes_de")
	ano_de= request.form("ano_de")
	data_inclui=dia_de&"/"&mes_de&"/"&ano_de
	data_busca=mes_de&"/"&dia_de&"/"&ano_de
	hora_ate= request.form("hora_ate")
	min_ate= request.form("min_ate")
	hora_bd=hora_ate&":"&min_ate
	
	IF co_prof="999999" or co_prof="" THEN
	co_prof=NULL
	END IF
	
	IF disciplina="999999" or disciplina="" THEN
	disciplina=""
	END IF
	
	cod_split=split(cod,", ")
	qtd_erro=0
	qtd_gravado=0
	for i=0 to ubound(cod_split)
	
		codigo=cod_split(i)
	
	
			Set RSo = server.createobject("adodb.recordset")
			SQLo = "SELECT * FROM TB_Ocorrencia_Aluno WHERE CO_Matricula ="&codigo&" AND DA_Ocorrencia = #"&data_busca&"# AND Right(HO_Ocorrencia,8) = #"&hora_bd&"# AND CO_Assunto='"&assunto&"' AND CO_Ocorrencia ="& tp_ocor
			RSo.Open SQLo, CON_o
	
		IF RSo.EOF then
			Set RS = server.createobject("adodb.recordset")
			RS.open "TB_Ocorrencia_Aluno", CON_o, 2, 2 'which table do you want open
			RS.addnew
			  RS("CO_Matricula") = codigo
			  RS("DA_Ocorrencia") = data_inclui
			  RS("HO_Ocorrencia") = hora_bd
			  RS("CO_Assunto") = assunto
			  RS("CO_Ocorrencia") = tp_ocor
			  RS("NU_Aula") = aula
			  RS("CO_Professor") = co_prof
			  RS("NO_Materia") = disciplina
			  RS("TX_Observa") = observacao
			  RS("CO_Usuario")= co_usr
			  RS.update
			  
			set RS=nothing
			
			if qtd_gravado=0 then
				gravado=codigo
			else
				gravado=gravado&"$!$"&codigo	
			end if
			qtd_gravado=qtd_gravado+1		
		else
			if qtd_erro=0 then
				erro=codigo
			else
				erro=erro&"$!$"&codigo	
			end if
			qtd_erro=qtd_erro+1	
		end if
	Next	

	outro= "Incluir-Multimatriculas,"&data_altera&","&hora_altera
	call GravaLog (chave,outro)

	if qtd_erro=0 then
		cod_url=Replace(cod, " ", "")		
		session("cod_url") = cod_url		
		data_inclui=Replace(data_inclui, "/", "$$$")	
		response.redirect("email_responsaveis.asp?opt=sim&ori=m&vt=s&obr="&obr&"&opt_a=ok1&dt="&data_inclui)							
		'response.redirect("email_responsaveis.asp?opt=ask&ori=m&vt=s&obr="&obr&"&opt_a=ok1&dt="&data_inclui)	
		'response.redirect("select_alunos.asp?cod="&cod&"&vt=s&obr="&obr&"&opt=ok1")	
	elseif qtd_gravado=0 then	
		response.redirect("select_alunos.asp?cod="&erro&"&vt=s&obr="&obr&"&opt=err1")
	else	
		response.redirect("select_alunos.asp?cod="&erro&"&vt=s&obr="&obr&"&opt=err2")
	end if	

elseif opt="alt" then
	cod= request.form("cod")
	
	tp_ocor=request.form("tp_ocor")
	assunto=request.form("assunto")
	disciplina=request.form("disciplina")
	aula=request.form("aula")
	co_prof=request.form("co_prof")
	data_inclui= request.form("data")
	data_altera= request.form("data_altera")
	hora_ate= request.form("hora")
	observacao= request.form("observacao")
	
	'IF co_prof="999999" THEN
	'co_prof=NULL
	'END IF
	
	IF disciplina="999999" THEN
	disciplina=NULL
	END IF
	'response.Write("<BR>'"&cod&"'")
	'response.Write("<BR>'"&data_inclui&"'")
	'response.Write("<BR>'"&hora_ate&"'")
	'response.Write("<BR>'"&assunto&"'")
	'response.Write("<BR>'"&tp_ocor&"'")
	'response.Write("<BR>'"&co_prof&"'")
	'response.Write("<BR>'"&disciplina&"'")
	'response.Write("<BR>'"&observacao&"'")
	'response.Write("<BR>'"&co_usr&"'")
	IF co_prof="999999" THEN
	response.Write("UPDATE TB_Ocorrencia_Aluno SET CO_Assunto ='"& assunto &"', CO_Ocorrencia ="& tp_ocor &" , NU_Aula ='"& aula &"', CO_Professor ="& co_prof &", NO_Materia ='"& disciplina &"', TX_Observa ='"& observacao &"', CO_Usuario ="& co_usr &"  WHERE CO_Matricula = "&cod&" AND CO_Assunto = '"&assunto &"' AND CO_Ocorrencia ="& tp_ocor &" AND DA_Ocorrencia =#"& data_altera &"# AND mid(HO_Ocorrencia,1,16) =#12/30/1899 "& hora_ate &"#")
	
	sql_atualiza= "UPDATE TB_Ocorrencia_Aluno SET CO_Assunto ='"& assunto &"', CO_Ocorrencia ="& tp_ocor &" , NU_Aula ='"& aula &"', CO_Professor =NULL, NO_Materia ='"& disciplina &"', TX_Observa ='"& observacao &"', CO_Usuario ="& co_usr &"  WHERE CO_Matricula = "&cod&" AND CO_Assunto = '"&assunto &"' AND CO_Ocorrencia ="& tp_ocor &" AND DA_Ocorrencia =#"& data_altera &"# AND mid(HO_Ocorrencia,1,16) =#12/30/1899 "& hora_ate &"#"
	else
	sql_atualiza= "UPDATE TB_Ocorrencia_Aluno SET CO_Assunto ='"& assunto &"', CO_Ocorrencia ="& tp_ocor &" , NU_Aula ='"& aula &"', CO_Professor ="& co_prof &", NO_Materia ='"& disciplina &"', TX_Observa ='"& observacao &"', CO_Usuario ="& co_usr &"  WHERE CO_Matricula = "&cod&" AND CO_Assunto = '"&assunto &"' AND CO_Ocorrencia ="& tp_ocor &" AND DA_Ocorrencia =#"& data_altera &"# AND mid(HO_Ocorrencia,1,16) =#12/30/1899 "& hora_ate &"#"
	END IF

	'response.end()
	Set RS2 = CON_o.Execute(sql_atualiza)
	outro= "Alterar,"&data_inclui&","&cod
	call GravaLog (chave,outro)
	obr=cod&"?dt?999999?1/1/"&ano_letivo&"?0:0?1/1/"&ano_letivo&", 0:0?12/31/"&ano_letivo&"?23:59?31/12/"&ano_letivo&", 23:59"
	session("obr")=obr
	'response.redirect("resumo.asp?cod="&cod&"&or=2&opt=ok2")
	cod_url=Replace(cod, " ", "")	
	session("cod_url") = cod_url	
	data_altera=Replace(data_altera, "/", "$$$")
	response.redirect("email_responsaveis.asp?opt=sim&ori=a&vt=s&obr="&obr&"&opt_a=ok2&dt="&data_altera)							
	'response.redirect("email_responsaveis.asp?opt=ask&ori=a&vt=s&obr="&obr&"&opt_a=ok2&dt="&data_altera)	
end if



%>