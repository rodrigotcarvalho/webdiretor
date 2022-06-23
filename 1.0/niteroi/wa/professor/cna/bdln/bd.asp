<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<%
'opt=request.form("opt")
'co_prof = request.form("prof")
'curso = request.Form("curso")
'unidade = request.Form("unidade")
'etapa= request.Form("etapa")
'turma= request.Form("turma")
'co_materia = request.form("mat")
'periodo = request.form("periodo")
'nota= request.form("nota")
opt=request.QueryString("opt")
curso = request.querystring("c")
unidade = request.querystring("u")
turma = request.querystring("t")
etapa = request.querystring("e")
co_prof= request.querystring("pr")
co_materia= request.querystring("d")
periodo= request.querystring("P")
nota=request.querystring("nt")

chave=session("chave")
session("chave")=chave


	Set CON2 = Server.CreateObject("ADODB.Connection") 
	ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON2.Open ABRIR2
	periodo = periodo*1
	

	
if opt="blq" then
	st= "x"	
	outro="Bloqueio da P:"&periodo&",D:"&co_materia&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma

elseif opt="dblq" then
	st= ""
	outro="Desbloqueio da P:"&periodo&",D:"&co_materia&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma
elseif opt="fblq" then
	st= "x"	
	outro="Bloqueio de todas as disciplinas do P:"&periodo&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma
elseif opt="fdblq" then
	st= ""
	outro="Desbloqueio todas as disciplinas do P:"&periodo&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma
end if
'response.Write("sql_atualiza= UPDATE TB_Da_Aula SET ST_Per_1 = '"&st&"' WHERE CO_Professor="& co_prof &" AND CO_Materia='"& co_materia &"' AND NU_Unidade="& unidade &" AND CO_Curso='"& curso &"' AND CO_Etapa='"& etapa &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& nota &"'")

if periodo=1 then
	situcao_periodo="ST_Per_1"
elseif periodo=2 then
	situcao_periodo="ST_Per_2"
elseif periodo =3 then
	situcao_periodo="ST_Per_3"
elseif periodo =4 then
	situcao_periodo="ST_Per_4"
elseif periodo =5 then
	situcao_periodo="ST_Per_5"
elseif periodo =6 then
	situcao_periodo="ST_Per_6"
end if

if opt="blq" or opt="dblq" then
	Set RS_b = Server.CreateObject("ADODB.Recordset")
	sql_atualiza= "UPDATE TB_Da_Aula SET "&situcao_periodo&" = '"&st&"' WHERE CO_Professor="& co_prof &" AND CO_Materia_principal='"& co_materia &"' AND NU_Unidade="& unidade &" AND CO_Curso='"& curso &"' AND CO_Etapa='"& etapa &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& nota &"'"
	Set RS2 = CON2.Execute(sql_atualiza)
else
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Da_Aula where NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& etapa &"'AND CO_Turma='"& turma &"'"
		RS2.Open SQL2, CON2
		
	if RS2.EOF then
	
	else		
		while not RS2.EOF					  
			co_materia = RS2("CO_Materia_Principal")
			co_prof = RS2("CO_Professor")	
			nota = RS2("TP_Nota")
			Set RS_b = Server.CreateObject("ADODB.Recordset")
			sql_atualiza= "UPDATE TB_Da_Aula SET "&situcao_periodo&" = '"&st&"' WHERE CO_Professor="& co_prof &" AND CO_Materia_principal='"& co_materia &"' AND NU_Unidade="& unidade &" AND CO_Curso='"& curso &"' AND CO_Etapa='"& etapa &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& nota &"'"
			response.Write(sql_atualiza&"<BR>")
			Set RS_b = CON2.Execute(sql_atualiza)
		RS2.MOVENEXT
		WEND
	end if	

end if
	call GravaLog (chave,outro)
if opt="blq" or opt="fblq" then

	response.Redirect("altera.asp?or=01&opt=ok&curso="&curso&"&unidade="&unidade&"&etapa="&etapa&"&turma="&turma&"&ano="&ano_letivo)

elseif opt="dblq" or opt="fdblq" then

	response.Redirect("altera.asp?or=02&opt=ok&curso="&curso&"&unidade="&unidade&"&etapa="&etapa&"&turma="&turma&"&ano="&ano_letivo)

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
pasta=arPath(seleciona1)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>