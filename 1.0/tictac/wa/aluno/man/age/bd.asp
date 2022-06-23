<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<%
chave=session("nvg")
session("nvg")=chave

opt = request.QueryString("opt")

		Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4 = "DBQ="& CAMINHO_e & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4	


ano_letivo = session("ano_letivo")
co_usr = session("co_user")



if opt="exc" then
exclui_entrevista=request.form("exclui_entrevista")
			
vertorExclui = split(exclui_entrevista,", ")
for i =0 to ubound(vertorExclui)

exclui = split(vertorExclui(i),"?")

cod = exclui(0)
da_entrevista= exclui(1)
ho_entrevista= exclui(2)


data_log=da_entrevista				
dados_data=split(da_entrevista,"/")
dia=dados_data(0)
mes=dados_data(1)
ano=dados_data(2)

da_entrevista_cons=mes&"/"&dia&"/"&ano

'response.Write(">>"&ho_entrevista)

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "DELETE * from TB_Entrevistas where CO_Matricula ="& cod&" AND DA_entrevista= #"&da_entrevista_cons&"# AND mid(HO_entrevista,1,16)=#12/30/1899 "&ho_entrevista&"#"	
		RS.Open SQL, CON4

next
obr=session("obr")
session("obr")=obr


outro= "Excluir,"&data_log&","&cod
call GravaLog (chave,outro)
response.redirect("resumo.asp?or=2&opt=ok")

elseif opt="inc" then


	cod= request.form("cod")
	tipo=request.form("tipo")
	assunto=request.form("assunto")
	participantes=request.form("participantes")
	agendado=request.form("agendado")
	data_inclui= request.form("data")
	data_altera= request.form("data_altera")
	observacao= request.form("observacao")
	
	dia_de= request.form("dia_de")
	mes_de= request.form("mes_de")
	ano_de= request.form("ano_de")
	data_entrevista=dia_de&"/"&mes_de&"/"&ano_de
	
	hora_de= request.form("hora_de")
	min_de= request.form("min_de")
	hora_entrevista=hora_de&":"&min_de
	
	Set RS = server.createobject("adodb.recordset")
	
	RS.open "TB_Entrevistas", CON4, 2, 2 'which table do you want open
	RS.addnew
	  RS("CO_Matricula") = cod
	  RS("DA_Entrevista") = data_entrevista
	  RS("HO_Entrevista") = hora_entrevista
	  RS("TP_Entrevista") = tipo
	  RS("NO_Participantes") = participantes
	  RS("ST_Entrevista") = 3
	  RS("CO_Agendado_com") = agendado
	  RS("TX_Observa") = observacao
	  RS("CO_Usuario")= co_usr
	  RS.update
	  
	set RS=nothing
	outro= "Incluir,"&data_entrevista&","&cod
	call GravaLog (chave,outro)
	'obr=cod&"?dt?999999?1/1/"&ano_letivo&"?0:0?1/1/"&ano_letivo&", 0:0?12/31/"&ano_letivo&"?23:59?31/12/"&ano_letivo&", 23:59"
	'session("obr")=obr
	response.redirect("resumo.asp?cod="&cod&"&or=2&opt=ok1")
elseif opt="alt" then
	cod= request.form("cod")
	tipo=request.form("tipo")
	assunto=request.form("assunto")
	participantes=request.form("participantes")
	agendado=request.form("agendado")
	data_inclui= request.form("data")
	data_altera= request.form("data_altera")
	observacao= request.form("observacao")
	status_entrevista = request.form("status")
	
	dia_de= request.form("dia_de")
	mes_de= request.form("mes_de")
	ano_de= request.form("ano_de")
		
	hora_de= request.form("hora_de")
	min_de= request.form("min_de")
	if isnull(dia_de) or dia_de="" then
	    dia_de= request.form("dia_de_disable")
		mes_de= request.form("mes_de_disable")
		ano_de= request.form("ano_de_disable")	
		hora_de = request.form("hora_de_disable")
		min_de= request.form("min_de_disable")		
	end if
	
	dia_original= request.form("dia_original")
	mes_original= request.form("mes_original")
	ano_original= request.form("ano_original")
	data_entrevista=dia_de&"/"&mes_de&"/"&ano_de
	da_entrevista_cons=mes_original&"/"&dia_original&"/"&ano_original	
	
	hora_original= request.form("hora_original")
	min_original= request.form("min_original")
	hora_entrevista=hora_de&":"&min_de
	hora_entrevista_cons=hora_original&":"&min_original



	Set RS0 = Server.CreateObject("ADODB.Recordset")
	SQL = "Select * from TB_Entrevistas WHERE CO_Matricula = "& cod &" AND DA_Entrevista = #"& da_entrevista_cons &"# AND mid(HO_entrevista,1,16)=#12/30/1899 "&hora_entrevista_cons&"#"
	RS0.Open SQL, CON4
	
	If RS0.EOF THEN	
		Set RS = Server.CreateObject("ADODB.Recordset")			
		RS.open "TB_Entrevistas", CON4, 2, 2 'which table do you want open
		RS.addnew
		  RS("CO_Matricula") = cod
		  RS("DA_Entrevista") = data_entrevista
		  RS("HO_Entrevista") = hora_entrevista
		  RS("TP_Entrevista") = tipo
		  RS("NO_Participantes") = participantes
		  RS("ST_Entrevista") = status_entrevista
		  RS("CO_Agendado_com") = agendado
		  RS("TX_Observa") = observacao
		  RS("CO_Usuario")= co_usr
		  RS.update
		  
		set RS=nothing	
	else
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "DELETE * from TB_Entrevistas where CO_Matricula ="& cod&" AND DA_entrevista= #"&da_entrevista_cons&"# AND mid(HO_entrevista,1,16)=#12/30/1899 "&hora_entrevista_cons&"#"		
		RS.Open SQL, CON4
		
		Set RS = Server.CreateObject("ADODB.Recordset")			
		RS.open "TB_Entrevistas", CON4, 2, 2 'which table do you want open
		RS.addnew
		  RS("CO_Matricula") = cod
		  RS("DA_Entrevista") = data_entrevista
		  RS("HO_Entrevista") = hora_entrevista
		  RS("TP_Entrevista") = tipo
		  RS("NO_Participantes") = participantes
		  RS("ST_Entrevista") = status_entrevista
		  RS("CO_Agendado_com") = agendado
		  RS("TX_Observa") = observacao
		  RS("CO_Usuario")= co_usr
		  RS.update  
		set RS=nothing	
	end if	
	
outro= "Alterar,"&data_entrevista&","&cod
call GravaLog (chave,outro)
'obr=cod&"?dt?999999?1/1/"&ano_letivo&"?0:0?1/1/"&ano_letivo&", 0:0?12/31/"&ano_letivo&"?23:59?31/12/"&ano_letivo&", 23:59"
'session("obr")=obr
'response.redirect("resumo.asp?cod="&cod&"&or=2&opt=ok2")
response.redirect("incluir.asp?ori=A&opt="&cod&"?"&data_entrevista&"?"&hora_entrevista&"&res=ok1")

elseif opt="con" then
cod= request.form("cod")
tipo=request.form("tipo")
assunto=request.form("assunto")
participantes=request.form("participantes")
agendado=request.form("agendado")
data_inclui= request.form("data")
data_altera= request.form("data_altera")
observacao= request.form("observacao")
conteudo= request.form("conteudo")

dia_de= request.form("dia_de")
mes_de= request.form("mes_de")
ano_de= request.form("ano_de")
data_entrevista=dia_de&"/"&mes_de&"/"&ano_de
da_entrevista_cons=mes_de&"/"&dia_de&"/"&ano_de

hora_de= request.form("hora_de")
min_de= request.form("min_de")
hora_entrevista=hora_de&":"&min_de
if tipo = "" or isnull(tipo) then
	tipo=tipo*1
end if	
'if tipo = 1 then

'else

	Set RS0 = Server.CreateObject("ADODB.Recordset")
	CONEXAO0 = "Select * from TB_Entrevistas_Conteudo WHERE CO_Matricula = "& cod &" AND DA_Entrevista = #"& da_entrevista_cons &"#  AND mid(HO_entrevista,1,16)=#12/30/1899 "&hora_entrevista&"#"
	Set RS0 = CON4.Execute(CONEXAO0)
	
	If RS0.EOF THEN	
	
		Set RS = Server.CreateObject("ADODB.Recordset")		
		RS.open "TB_Entrevistas_Conteudo", CON4, 2, 2 'which table do you want open
		RS.addnew
		  RS("CO_Matricula") = cod
		  RS("DA_Entrevista") = data_entrevista
		  RS("HO_Entrevista") = hora_entrevista
		  RS("TX_Conteudo") = conteudo
		  Ra("CO_Usuario")= co_usr
		  RS.update
		  
		set RS=nothing		
	'response.Write("OK")					
	else			

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		CONEXAO0 = "DELETE * from TB_Entrevistas_Conteudo WHERE CO_Matricula = "& cod &" AND DA_Entrevista = #"& data_entrevista &"# AND mid(HO_entrevista,1,16)=#12/30/1899 "&hora_entrevista&"#"
		Set RS0 = CON4.Execute(CONEXAO0)

		Set RS = Server.CreateObject("ADODB.Recordset")					
		RS.open "TB_Entrevistas_Conteudo", CON4, 2, 2 'which table do you want open
		RS.addnew
		  RS("CO_Matricula") = cod
		  RS("DA_Entrevista") = data_entrevista
		  RS("HO_Entrevista") = hora_entrevista
		  RS("TX_Conteudo") = conteudo
		  RS("CO_Usuario")= co_usr
		  RS.update
		  
		set RS=nothing		

	end if
	
	 sql = "update TB_Entrevistas set ST_Entrevista = 1 WHERE CO_Matricula = "& cod &" AND DA_Entrevista = #"& data_entrevista &"# AND mid(HO_entrevista,1,16)=#12/30/1899 "&hora_entrevista&"#"	      
         
	CON4.execute(sql)     
	CON4.close   	
'end if		
outro= "Conteúdo,"&data_entrevista&","&cod
call GravaLog (chave,outro)
'obr=cod&"?dt?999999?1/1/"&ano_letivo&"?0:0?1/1/"&ano_letivo&", 0:0?12/31/"&ano_letivo&"?23:59?31/12/"&ano_letivo&", 23:59"
'session("obr")=obr
response.redirect("incluir.asp?ori=C&opt="&cod&"?"&data_entrevista&"?"&hora_entrevista&"&res=ok2")
end if



%>