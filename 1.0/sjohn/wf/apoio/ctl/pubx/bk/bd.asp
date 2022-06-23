<!--#include file="../../../../inc/caminhos.asp"-->
<%
opt=request.QueryString("opt")

if opt="i" or opt="a" then
tit = request.form("tit")
arquivo = request.form("arquivo")
tipo_arquivo = session("tipo_arquivo")
session("tipo_arquivo") =tipo_arquivo


unidade_grava = request.form("unidade")
curso_grava = request.form("curso")
etapa_grava = request.form("etapa")
turma_grava = request.form("turma")
'response.Write(turma_grava)

if tit="" or isnull(tit) then
total=len(arquivo)-4
tit=Left(arquivo,total)
end if

unidade_grava1 = request.form("unidade1")
curso_grava1 = request.form("curso1")
etapa_grava1 = request.form("etapa1")
turma_grava1 = request.form("turma1")

unidade_grava2 = request.form("unidade2")
curso_grava2 = request.form("curso2")
etapa_grava2 = request.form("etapa2")
turma_grava2 = request.form("turma2")

unidade_grava3 = request.form("unidade3")
curso_grava3 = request.form("curso3")
etapa_grava3 = request.form("etapa3")
turma_grava3 = request.form("turma3")

unidade_grava4 = request.form("unidade4")
curso_grava4 = request.form("curso4")
etapa_grava4 = request.form("etapa4")
turma_grava4 = request.form("turma4")

if unidade_grava="999990" or unidade_grava="" or isnull(unidade_grava) then
sql_un="Unidade = NULL, "
unidade_grava= NULL
else
sql_un="Unidade= '"&unidade_grava&"', "
end if

if curso_grava="999990" or curso_grava="" or isnull(curso_grava) then
sql_cu="Curso = NULL,"
curso_grava=NULL
else
sql_cu="Curso='"&curso_grava&"', "
end if

if etapa_grava="999990" or etapa_grava="" or isnull(etapa_grava) then
sql_et="Etapa = NULL, "
etapa_grava=NULL
else
sql_et="Etapa='"&etapa_grava&"', "

end if

if turma_grava="999990" or turma_grava="" or isnull(turma_grava) then
sql_tu="Turma = NULL , "
turma_grava=NULL
else
sql_tu="Turma='"&turma_grava&"', "
end if


if unidade_grava1="999990" and curso_grava1="999990" and etapa_grava1="999990" and turma_grava1="999990" then
assoc1="n"
else
if unidade_grava1="999990" or unidade_grava1="" or isnull(unidade_grava1) then
'sql_un="Unidade = NULL, "
'unidade_grava1= NULL
else
'sql_un="Unidade= '"&unidade_grava1&"', "
end if

if curso_grava1="999990" or curso_grava1="" or isnull(curso_grava1) then
'sql_cu="Curso = NULL,"
curso_grava1=NULL
'else
'sql_cu="Curso='"&curso_grava1&"', "
end if

if etapa_grava1="999990" or etapa_grava1="" or isnull(etapa_grava1) then
'sql_et="Etapa = NULL, "
etapa_grava1=NULL
'else
'sql_et="Etapa='"&etapa_grava1&"', "

end if

if turma_grava1="999990" or turma_grava1="" or isnull(turma_grava1) then
'sql_tu="Turma = NULL , "
turma_grava1=NULL
'else
'sql_tu="(Turma='"&turma_grava1&"', "
end if

assoc1="s" 
end if

if unidade_grava2="999990" and curso_grava2="999990" and etapa_grava2="999990" and turma_grava2="999990" then
assoc2="n"
else
if unidade_grava2="999990" or unidade_grava2="" or isnull(unidade_grava2) then
'sql_un="Unidade = NULL, "
unidade_grava2= NULL
'else
'sql_un="Unidade= '"&unidade_grava2&"', "
end if

if curso_grava2="999990" or curso_grava2="" or isnull(curso_grava2) then
'sql_cu="Curso = NULL,"
curso_grava2=NULL
'else
'sql_cu="Curso='"&curso_grava2&"', "
end if

if etapa_grava2="999990" or etapa_grava2="" or isnull(etapa_grava2) then
'sql_et="Etapa = NULL, "
etapa_grava2=NULL
'else
'sql_et="Etapa='"&etapa_grava2&"', "

end if

if turma_grava2="999990" or turma_grava2="" or isnull(turma_grava2) then
'sql_tu="Turma = NULL , "
turma_grava2=NULL
'else
'sql_tu="(Turma='"&turma_grava2&"', "
end if

assoc2="s" 
end if

if unidade_grava3="999990" and curso_grava3="999990" and etapa_grava3="999990" and turma_grava3="999990" then
assoc3="n"
else
if unidade_grava3="999990" or unidade_grava3="" or isnull(unidade_grava3) then
'sql_un="Unidade = NULL, "
unidade_grava3= NULL
'else
'sql_un="Unidade= '"&unidade_grava3&"', "
end if

if curso_grava3="999990" or curso_grava3="" or isnull(curso_grava3) then
'sql_cu="Curso = NULL,"
curso_grava3=NULL
'else
'sql_cu="Curso='"&curso_grava3&"', "
end if

if etapa_grava3="999990" or etapa_grava3="" or isnull(etapa_grava3) then
'sql_et="Etapa = NULL, "
etapa_grava3=NULL
'else
'sql_et="Etapa='"&etapa_grava3&"', "

end if

if turma_grava3="999990" or turma_grava3="" or isnull(turma_grava3) then
'sql_tu="Turma = NULL , "
turma_grava4=NULL
'else
'sql_tu="(Turma='"&turma_grava3&"', "
end if

assoc3="s" 
end if

if unidade_grava4="999990" and curso_grava4="999990" and etapa_grava4="999990" and turma_grava4="999990" then
assoc4="n"
else
if unidade_grava4="999990" or unidade_grava4="" or isnull(unidade_grava4) then
'sql_un="Unidade = NULL, "
unidade_grava4= NULL
else
'sql_un="Unidade= '"&unidade_grava4&"', "
end if

if curso_grava4="999990" or curso_grava4="" or isnull(curso_grava4) then
'sql_cu="Curso = NULL,"
curso_grava4=NULL
else
'sql_cu="Curso='"&curso_grava4&"', "
end if

if etapa_grava4="999990" or etapa_grava4="" or isnull(etapa_grava4) then
'sql_et="Etapa = NULL, "
etapa_grava4=NULL
else
'sql_et="Etapa='"&etapa_grava4&"', "

end if

if turma_grava4="999990" or turma_grava4="" or isnull(turma_grava4) then
'sql_tu="Turma = NULL , "
turma_grava4=NULL
else
'sql_tu="(Turma='"&turma_grava4&"', "
end if

assoc4="s" 
end if


dia_de= request.form("dia_de")
mes_de= request.form("mes_de")
ano_de= request.form("ano_de")
data_inclui=dia_de&"/"&mes_de&"/"&ano_de

tipo_arquivo=tipo_arquivo*1

end if

    	Set CON_WF = Server.CreateObject("ADODB.Connection") 
		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_WF.Open ABRIR_WF




if opt="i" then

Set RS_updt = server.createobject("adodb.recordset")

RS_updt.open "TB_Documentos", CON_WF, 2, 2 'which table do you want open
RS_updt.addnew

	RS_updt("NO_Doc")=arquivo
	RS_updt("Unidade") = unidade_grava
	RS_updt("Curso") = curso_grava
	RS_updt("Etapa") = etapa_grava
	RS_updt("Turma")=turma_grava
	RS_updt("TI1_Doc")=tit
	RS_updt("TP_Doc")=tipo_arquivo
	RS_updt("DA_Doc")=data_inclui	

RS_updt.update
set RS_updt=nothing

if assoc1="s" then

Set RS_updt = server.createobject("adodb.recordset")

RS_updt.open "TB_Documentos", CON_WF, 2, 2 'which table do you want open
RS_updt.addnew

	RS_updt("NO_Doc")=arquivo
	RS_updt("Unidade") = unidade_grava1
	RS_updt("Curso") = curso_grava1
	RS_updt("Etapa") = etapa_grava1
	RS_updt("Turma")=turma_grava1
	RS_updt("TI1_Doc")=tit
	RS_updt("TP_Doc")=tipo_arquivo
	RS_updt("DA_Doc")=data_inclui	

RS_updt.update
set RS_updt=nothing
end if

if assoc2="s" then

Set RS_updt = server.createobject("adodb.recordset")

RS_updt.open "TB_Documentos", CON_WF, 2, 2 'which table do you want open
RS_updt.addnew

	RS_updt("NO_Doc")=arquivo
	RS_updt("Unidade") = unidade_grava2
	RS_updt("Curso") = curso_grava2
	RS_updt("Etapa") = etapa_grava2
	RS_updt("Turma")=turma_grava2
	RS_updt("TI1_Doc")=tit
	RS_updt("TP_Doc")=tipo_arquivo
	RS_updt("DA_Doc")=data_inclui	

RS_updt.update
set RS_updt=nothing
end if

if assoc3="s" then

Set RS_updt = server.createobject("adodb.recordset")

RS_updt.open "TB_Documentos", CON_WF, 2, 2 'which table do you want open
RS_updt.addnew

	RS_updt("NO_Doc")=arquivo
	RS_updt("Unidade") = unidade_grava3
	RS_updt("Curso") = curso_grava3
	RS_updt("Etapa") = etapa_grava3
	RS_updt("Turma")=turma_grava3
	RS_updt("TI1_Doc")=tit
	RS_updt("TP_Doc")=tipo_arquivo
	RS_updt("DA_Doc")=data_inclui	

RS_updt.update
set RS_updt=nothing
end if

if assoc4="s" then

Set RS_updt = server.createobject("adodb.recordset")

RS_updt.open "TB_Documentos", CON_WF, 2, 2 'which table do you want open
RS_updt.addnew

	RS_updt("NO_Doc")=arquivo
	RS_updt("Unidade") = unidade_grava4
	RS_updt("Curso") = curso_grava4
	RS_updt("Etapa") = etapa_grava4
	RS_updt("Turma")=turma_grava4
	RS_updt("TI1_Doc")=tit
	RS_updt("TP_Doc")=tipo_arquivo
	RS_updt("DA_Doc")=data_inclui	

RS_updt.update
set RS_updt=nothing

end if

response.Redirect("incluir.asp?opt=ok")

elseif opt="a" then
co_doc = request.form("co_doc")

'response.Write "UPDATE TB_Documentos SET TP_Doc='"&tipo_arquivo&"',NO_Doc='"&arquivo&"', "&sql_un&sql_cu&sql_et&sql_tu&"TI1_Doc ='"&tit&"', DA_Doc ='"&data_inclui&"' WHERE CO_Doc = "& co_doc
'response.end()
sql_atualiza= "UPDATE TB_Documentos SET TP_Doc='"&tipo_arquivo&"',NO_Doc='"&arquivo&"', "&sql_un&sql_cu&sql_et&sql_tu&"TI1_Doc ='"&tit&"', DA_Doc ='"&data_inclui&"' WHERE CO_Doc = "& co_doc
Set RS_updt2 = CON_WF.Execute(sql_atualiza)

if assoc1="s" then

Set RS_updt = server.createobject("adodb.recordset")

RS_updt.open "TB_Documentos", CON_WF, 2, 2 'which table do you want open
RS_updt.addnew

	RS_updt("NO_Doc")=arquivo
	RS_updt("Unidade") = unidade_grava1
	RS_updt("Curso") = curso_grava1
	RS_updt("Etapa") = etapa_grava1
	RS_updt("Turma")=turma_grava1
	RS_updt("TI1_Doc")=tit
	RS_updt("TP_Doc")=tipo_arquivo
	RS_updt("DA_Doc")=data_inclui	

RS_updt.update
set RS_updt=nothing
end if

if assoc2="s" then

Set RS_updt = server.createobject("adodb.recordset")

RS_updt.open "TB_Documentos", CON_WF, 2, 2 'which table do you want open
RS_updt.addnew

	RS_updt("NO_Doc")=arquivo
	RS_updt("Unidade") = unidade_grava2
	RS_updt("Curso") = curso_grava2
	RS_updt("Etapa") = etapa_grava2
	RS_updt("Turma")=turma_grava2
	RS_updt("TI1_Doc")=tit
	RS_updt("TP_Doc")=tipo_arquivo
	RS_updt("DA_Doc")=data_inclui	

RS_updt.update
set RS_updt=nothing
end if

if assoc3="s" then

Set RS_updt = server.createobject("adodb.recordset")

RS_updt.open "TB_Documentos", CON_WF, 2, 2 'which table do you want open
RS_updt.addnew

	RS_updt("NO_Doc")=arquivo
	RS_updt("Unidade") = unidade_grava3
	RS_updt("Curso") = curso_grava3
	RS_updt("Etapa") = etapa_grava3
	RS_updt("Turma")=turma_grava3
	RS_updt("TI1_Doc")=tit
	RS_updt("TP_Doc")=tipo_arquivo
	RS_updt("DA_Doc")=data_inclui	

RS_updt.update
set RS_updt=nothing
end if

if assoc4="s" then

Set RS_updt = server.createobject("adodb.recordset")

RS_updt.open "TB_Documentos", CON_WF, 2, 2 'which table do you want open
RS_updt.addnew

	RS_updt("NO_Doc")=arquivo
	RS_updt("Unidade") = unidade_grava4
	RS_updt("Curso") = curso_grava4
	RS_updt("Etapa") = etapa_grava4
	RS_updt("Turma")=turma_grava4
	RS_updt("TI1_Doc")=tit
	RS_updt("TP_Doc")=tipo_arquivo
	RS_updt("DA_Doc")=data_inclui	

RS_updt.update
set RS_updt=nothing
end if


response.Redirect("alterar.asp?opt=ok&c="&co_doc)

elseif opt="e" then
dia_de= Session("dia_de")
mes_de= Session("dia_de")
dia_ate=Session("dia_ate")
mes_ate=Session("mes_ate")
unidade=Session("unidade")
curso=Session("curso")
etapa=Session("etapa")
turma=Session("turma")
tit=Session("tit")
check_status=Session("check_status")
tp_doc=session("tipo_arquivo")

Session("dia_de")=dia_de
Session("dia_de")=mes_de
Session("dia_ate")=dia_ate
Session("mes_ate")=mes_ate
Session("unidade")=unidade
Session("curso")=curso
Session("etapa")=etapa
Session("turma")=turma
Session("tit")=tit
Session("check_status")=check_status
session("tipo_arquivo") =tp_doc




exclui_doc=request.form("exclui_doc")

vertorExclui = split(exclui_doc,", ")
conta_ocorr=0
for i =0 to ubound(vertorExclui)

co_doc = vertorExclui(i)
		
		Set RS_doc = Server.CreateObject("ADODB.Recordset")
SQL_doc = "DELETE * FROM TB_Documentos where CO_Doc="&co_doc
		RS_doc.Open SQL_doc, CON_WF
		
next		
response.Redirect("docs.asp?opt=ok&pagina=1&v=s")
end if
%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
 window.history.forward(1);
</script>
</head>
<body>
</body>
</html>