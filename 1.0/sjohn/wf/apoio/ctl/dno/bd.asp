<!--#include file="../../../../inc/caminhos.asp"-->
<%
opt=request.QueryString("opt")


tit = request.form("tit")
conteudo = request.form("conteudo")
tipo_arquivo = request.form("tipo_doc")
unidade_grava = request.form("unidade")
curso_grava = request.form("curso")
etapa_grava = request.form("etapa")
turma_grava = request.form("turma")

if unidade_grava="999990" then
	unidade_grava=null
sql_un="Unidade = null, "
else
sql_un="Unidade ='"&unidade_grava&"', "
end if

if curso_grava="999990" then
curso_grava=null
sql_cu="Curso = null, "
else
sql_cu="Curso ='"&curso_grava&"', "
end if

if etapa_grava="999990" then
etapa_grava=null
sql_et="Etapa = null, "
else
sql_et="Etapa ='"&etapa_grava&"', "
end if

if turma_grava="999990" then
turma_grava=null
sql_tu="Turma = null, "
else
sql_tu="Turma ='"&turma_grava&"', "
end if

dia_de= request.form("dia_de")
mes_de= request.form("mes_de")
ano_de= request.form("ano_de")
data_inclui=dia_de&"/"&mes_de&"/"&ano_de

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

if unidade_grava1="999990" and curso_grava1="999990" and etapa_grava1="999990" and turma_grava1="999990" then
assoc1="n"
else
if unidade_grava1="999990" then
unidade_grava1=null
end if

if curso_grava1="999990" then
curso_grava1=null
end if

if etapa_grava1="999990" then
etapa_grava1=null
end if

if turma_grava1="999990" then
turma_grava1=null
end if

assoc1="s" 
end if

if unidade_grava2="999990" and curso_grava2="999990" and etapa_grava2="999990" and turma_grava2="999990" then
assoc2="n"
else
if unidade_grava2="999990" then
unidade_grava2=null
end if

if curso_grava2="999990" then
curso_grava2=null
end if

if etapa_grava2="999990" then
etapa_grava2=null
end if

if turma_grava2="999990" then
turma_grava2=null
end if

assoc2="s" 
end if

if unidade_grava3="999990" and curso_grava3="999990" and etapa_grava3="999990" and turma_grava3="999990" then
assoc3="n"
else
if unidade_grava3="999990" then
unidade_grava3=null
end if

if curso_grava3="999990" then
curso_grava3=null
end if

if etapa_grava3="999990" then
etapa_grava3=null
end if

if turma_grava3="999990" then
turma_grava3=null
end if

assoc3="s" 
end if

if unidade_grava4="999990" and curso_grava4="999990" and etapa_grava4="999990" and turma_grava4="999990" then
assoc4="n"
else
if unidade_grava4="999990" then
unidade_grava4=null
end if

if curso_grava4="999990" then
curso_grava4=null
end if

if etapa_grava4="999990" then
etapa_grava4=null
end if

if turma_grava4="999990" then
turma_grava4=null
end if

assoc4="s" 
end if

tipo_arquivo=tipo_arquivo*1

    	Set CON_WF = Server.CreateObject("ADODB.Connection") 
		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_WF.Open ABRIR_WF

dia_ate= request.form("dia_ate")
mes_ate= request.form("mes_ate")
ano_ate= request.form("ano_ate")

if dia_ate=0 or dia_ate="0" or mes_ate=0 or mes_ate="0" or ano_ate=0 or ano_ate="0" then
data_vig_inclui=NULL
sql_vg="NT_DT_Vg= NULL, "
else
data_vig_inclui=dia_ate&"/"&mes_ate&"/"&ano_ate
sql_vg="NT_DT_Vg='"&data_vig_inclui&"', "
end if


if opt="i" then

Set RS_updt = server.createobject("adodb.recordset")

RS_updt.open "TB_Noticias", CON_WF, 2, 2 'which table do you want open
RS_updt.addnew

	RS_updt("NT_Conteudo")=conteudo
	RS_updt("NT_Titulo")=tit
	RS_updt("TP_Noticia")=tipo_arquivo
	RS_updt("NT_DT_Pb")=data_inclui
	RS_updt("NT_DT_Vg")=data_vig_inclui
	RS_updt("Unidade") = unidade_grava
	RS_updt("Curso") = curso_grava
	RS_updt("Etapa") = etapa_grava
	RS_updt("Turma")=turma_grava	
if dia_ate=0 or dia_ate="0" or mes_ate=0 or mes_ate="0" or ano_ate=0 or ano_ate="0" then
else	
	RS_updt("NT_DT_Vg")=data_vig_inclui
end if			

RS_updt.update
set RS_updt=nothing

if assoc1="s" then

Set RS_updt = server.createobject("adodb.recordset")

RS_updt.open "TB_Noticias", CON_WF, 2, 2 'which table do you want open
RS_updt.addnew

	RS_updt("NT_Conteudo")=conteudo
	RS_updt("NT_Titulo")=tit
	RS_updt("TP_Noticia")=tipo_arquivo
	RS_updt("NT_DT_Pb")=data_inclui
	RS_updt("NT_DT_Vg")=data_vig_inclui
	RS_updt("Unidade") = unidade_grava1
	RS_updt("Curso") = curso_grava1
	RS_updt("Etapa") = etapa_grava1
	RS_updt("Turma")=turma_grava1	

RS_updt.update
set RS_updt=nothing
end if

if assoc2="s" then

Set RS_updt = server.createobject("adodb.recordset")

RS_updt.open "TB_Noticias", CON_WF, 2, 2 'which table do you want open
RS_updt.addnew

	RS_updt("NT_Conteudo")=conteudo
	RS_updt("NT_Titulo")=tit
	RS_updt("TP_Noticia")=tipo_arquivo
	RS_updt("NT_DT_Pb")=data_inclui
	RS_updt("NT_DT_Vg")=data_vig_inclui
	RS_updt("Unidade") = unidade_grava2
	RS_updt("Curso") = curso_grava2
	RS_updt("Etapa") = etapa_grava2
	RS_updt("Turma")=turma_grava2	

RS_updt.update
set RS_updt=nothing
end if

if assoc3="s" then

Set RS_updt = server.createobject("adodb.recordset")

RS_updt.open "TB_Noticias", CON_WF, 2, 2 'which table do you want open
RS_updt.addnew

	RS_updt("NT_Conteudo")=conteudo
	RS_updt("NT_Titulo")=tit
	RS_updt("TP_Noticia")=tipo_arquivo
	RS_updt("NT_DT_Pb")=data_inclui
	RS_updt("NT_DT_Vg")=data_vig_inclui
	RS_updt("Unidade") = unidade_grava3
	RS_updt("Curso") = curso_grava3
	RS_updt("Etapa") = etapa_grava3
	RS_updt("Turma")=turma_grava3	

RS_updt.update
set RS_updt=nothing
end if

if assoc4="s" then

Set RS_updt = server.createobject("adodb.recordset")

RS_updt.open "TB_Noticias", CON_WF, 2, 2 'which table do you want open
RS_updt.addnew

	RS_updt("NT_Conteudo")=conteudo
	RS_updt("NT_Titulo")=tit
	RS_updt("TP_Noticia")=tipo_arquivo
	RS_updt("NT_DT_Pb")=data_inclui
	RS_updt("NT_DT_Vg")=data_vig_inclui
	RS_updt("Unidade") = unidade_grava4
	RS_updt("Curso") = curso_grava4
	RS_updt("Etapa") = etapa_grava4
	RS_updt("Turma")=turma_grava4
RS_updt.update
set RS_updt=nothing

end if

response.Redirect("incluir.asp?opt=ok")

elseif opt="a" then
co_not = request.form("co_not")
sql_atualiza= "UPDATE TB_Noticias SET "&sql_un&sql_cu&sql_et&sql_tu&sql_vg&"TP_Noticia='"&tipo_arquivo&"', NT_Titulo ='"&tit&"', NT_DT_Pb ='"&data_inclui&"', NT_Conteudo='"&conteudo&"' WHERE NT_Codigo = "& co_not
'sql_atualiza= "UPDATE TB_Noticias SET TP_Noticia='"&tipo_arquivo&"', Unidade ='"&unidade_grava&"', Curso ='"&curso_grava&"', Etapa ='"&etapa_grava&"', Turma ='"&turma_grava&"', NT_Titulo ='"&tit&"', NT_DT_Pb ='"&data_inclui&"', NT_DT_Vg='"&data_vig_inclui&"', NT_Conteudo='"&conteudo&"' WHERE NT_Codigo = "& co_not

response.Write(sql_atualiza)
Set RS_updt2 = CON_WF.Execute(sql_atualiza)

if assoc1="s" then

Set RS_updt = server.createobject("adodb.recordset")

RS_updt.open "TB_Noticias", CON_WF, 2, 2 'which table do you want open
RS_updt.addnew

	RS_updt("NT_Conteudo")=conteudo
	RS_updt("NT_Titulo")=tit
	RS_updt("TP_Noticia")=tipo_arquivo
	RS_updt("NT_DT_Pb")=data_inclui
	RS_updt("NT_DT_Vg")=data_vig_inclui
	RS_updt("Unidade") = unidade_grava1
	RS_updt("Curso") = curso_grava1
	RS_updt("Etapa") = etapa_grava1
	RS_updt("Turma")=turma_grava1	

RS_updt.update
set RS_updt=nothing
end if

if assoc2="s" then

Set RS_updt = server.createobject("adodb.recordset")

RS_updt.open "TB_Noticias", CON_WF, 2, 2 'which table do you want open
RS_updt.addnew

	RS_updt("NT_Conteudo")=conteudo
	RS_updt("NT_Titulo")=tit
	RS_updt("TP_Noticia")=tipo_arquivo
	RS_updt("NT_DT_Pb")=data_inclui
	RS_updt("NT_DT_Vg")=data_vig_inclui
	RS_updt("Unidade") = unidade_grava2
	RS_updt("Curso") = curso_grava2
	RS_updt("Etapa") = etapa_grava2
	RS_updt("Turma")=turma_grava2	

RS_updt.update
set RS_updt=nothing
end if

if assoc3="s" then

Set RS_updt = server.createobject("adodb.recordset")

RS_updt.open "TB_Noticias", CON_WF, 2, 2 'which table do you want open
RS_updt.addnew

	RS_updt("NT_Conteudo")=conteudo
	RS_updt("NT_Titulo")=tit
	RS_updt("TP_Noticia")=tipo_arquivo
	RS_updt("NT_DT_Pb")=data_inclui
	RS_updt("NT_DT_Vg")=data_vig_inclui
	RS_updt("Unidade") = unidade_grava3
	RS_updt("Curso") = curso_grava3
	RS_updt("Etapa") = etapa_grava3
	RS_updt("Turma")=turma_grava3	

RS_updt.update
set RS_updt=nothing
end if

if assoc4="s" then

Set RS_updt = server.createobject("adodb.recordset")

RS_updt.open "TB_Noticias", CON_WF, 2, 2 'which table do you want open
RS_updt.addnew

	RS_updt("NT_Conteudo")=conteudo
	RS_updt("NT_Titulo")=tit
	RS_updt("TP_Noticia")=tipo_arquivo
	RS_updt("NT_DT_Pb")=data_inclui
	RS_updt("NT_DT_Vg")=data_vig_inclui
	RS_updt("Unidade") = unidade_grava4
	RS_updt("Curso") = curso_grava4
	RS_updt("Etapa") = etapa_grava4
	RS_updt("Turma")=turma_grava4
RS_updt.update
set RS_updt=nothing

end if


response.Redirect("alterar.asp?opt=ok&c="&co_not)

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


Session("dia_de")=dia_de
Session("dia_de")=mes_de
Session("dia_ate")=dia_ate
Session("mes_ate")=mes_ate
Session("unidade")=unidade
Session("curso")=curso
Session("etapa")=etapa
Session("turma")=turma
Session("tit")=tit





exclui_doc=request.form("exclui_doc")

vertorExclui = split(exclui_doc,", ")
conta_ocorr=0
for i =0 to ubound(vertorExclui)

co_doc = vertorExclui(i)
		
		Set RS_doc = Server.CreateObject("ADODB.Recordset")
SQL_doc = "DELETE * FROM TB_Noticias where NT_Codigo="&co_doc
		RS_doc.Open SQL_doc, CON_WF
		
next		
response.Redirect("docs.asp?opt=ok&pagina=1&v=s")
end if
%>