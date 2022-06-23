<!--#include file="../../../../inc/connect_ct.asp"-->
<!--#include file="../../../../inc/connect_pr.asp"-->

<%
'opt= request.querystring("opt")

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0

		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT

tp= request.form("tp_familiares")
nome= request.form("nome_pub")
if tp="p" then
tp_nome="Pai"
	if session("pai_ok")="s" then
	session("pai_ok")=session("pai_ok")
	else
	session("pai_ok")="s"
	end if
session("mae_ok")=session("mae_ok")
elseif tp="m" then
tp_nome="Mãe"
	if session("mae_ok")="s" then
	session("mae_ok")=session("mae_ok")
	else
	session("mae_ok")="s"
	end if
session("pai_ok")=session("pai_ok")
end if

%>
<font class="form_corpo"><%response.Write(Server.URLEncode(nome))%></font>