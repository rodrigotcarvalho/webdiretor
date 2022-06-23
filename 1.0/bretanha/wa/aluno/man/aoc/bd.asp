<%'On Error Resume Next%>

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
co_prof=request.form("no_prof")
data_inclui= request.form("data")
data_altera= request.form("data_altera")
observacao= request.form("observacao")

dia_de= request.form("dia_de")
mes_de= request.form("mes_de")
ano_de= request.form("ano_de")
data_inclui=dia_de&"/"&mes_de&"/"&ano_de

hora_ate= request.form("hora_ate")
min_ate= request.form("min_ate")
hora=hora_ate&":"&min_ate

IF co_prof="999999" THEN
co_prof=NULL
END IF

IF disciplina="999999" THEN
disciplina=""
END IF

response.Write("<BR>'"&hora_ate&"'")
response.Write("<BR>'"&min_ate&"'")

response.Write("<BR>'"&cod&"'")
response.Write("<BR>'"&data_inclui&"'")
response.Write("<BR>'"&hora_ate&"'")
response.Write("<BR>'"&assunto&"'")
response.Write("<BR>'"&tp_ocor&"'")
response.Write("<BR>'"&aula&"'")
response.Write("<BR>'"&co_prof&"'")
response.Write("<BR>'"&disciplina&"'")
response.Write("<BR>'"&observacao&"'")
response.Write("<BR>'"&co_usr&"'")
response.Write("<BR>'"&hora&"'")
'response.end()
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
outro= "Incluir,"&data_inclui&","&cod
call GravaLog (chave,outro)
obr=cod&"?dt?999999?1/1/"&ano_letivo&"?0:0?1/1/"&ano_letivo&", 0:0?12/31/"&ano_letivo&"?23:59?31/12/"&ano_letivo&", 23:59"
session("obr")=obr
response.redirect("resumo.asp?cod="&cod&"&or=2&opt=ok1")

elseif opt="alt" then
cod= request.form("cod")
tp_ocor=request.form("tp_ocor")
assunto=request.form("assunto")
disciplina=request.form("disciplina")
aula=request.form("aula")
co_prof=request.form("no_prof")
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
response.redirect("resumo.asp?cod="&cod&"&or=2&opt=ok2")
end if



%>