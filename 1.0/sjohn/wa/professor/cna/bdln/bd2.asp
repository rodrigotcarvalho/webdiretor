<%'On Error Resume Next%>

<!--#include file="../../../../inc/funcoes.asp"-->
<%

opt=request.QueryString("opt")
dados=split(opt,"_")
unidade = dados(0)
curso = dados(1)
etapa = dados(2)
turma = dados(3)
periodo = dados(4)
opt = dados(5)


chave=session("chave")
session("chave")=chave


Set CON2 = Server.CreateObject("ADODB.Connection") 
ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
CON2.Open ABRIR2
periodo = periodo*1

if unidade="999990" or unidade="" or isnull(unidade) then
	SQL_ALUNOS= "where CO_Curso <> '0'"
else	
	SQL_ALUNOS= "where NU_Unidade = "& unidade 	
	if curso="999990" or curso="" or isnull(curso) then
		SQL_CURSO=" AND CO_Curso <> '0'"	
	else
		if isnumeric(curso) then
			curso=curso*1
			if curso=0 then
				gera_dados="N"
				SQL_CURSO=" AND CO_Curso <> '0'"
			end if
		end if
		SQL_CURSO=" AND CO_Curso = '"& curso &"'"			
	end if

	if etapa="999990" or etapa="" or isnull(etapa) then
		SQL_ETAPA=""		
	else
		SQL_ETAPA=" AND CO_Etapa = '"& etapa &"'"				
	end if

	if turma="999990" or turma="" or isnull(turma) then
		SQL_TURMA=""		
	else
		SQL_TURMA=" AND CO_Turma = '"& turma &"' "			
	end if	

	SQL_ALUNOS= SQL_ALUNOS&SQL_CURSO&SQL_ETAPA&SQL_TURMA
		
'	Set RS = Server.CreateObject("ADODB.Recordset")
'	SQL= "Select distinct ST_Per_1 as P1, distinct ST_Per_2  as P2, distinct ST_Per_3  as P3, distinct ST_Per_4  as P4, distinct ST_Per_5  as P5, distinct ST_Per_6 as P6 from TB_Da_Aula "&SQL_ALUNOS
'	response.Write(SQL)
'	Set RS= CON2.Execute(SQL)
'		
'	while not RS.EOF 
'		p1=RS("P1")
'		p2=RS("P2")
'		p3=RS("P3")
'		p4=RS("P4")
'		p5=RS("P5")
'		p6=RS("P6")
'											
'	'	response.Write(p1&","&p1&","&p2&","&p3&","&p4&","&p5&","&p6&"<BR>")
'	
'	RS.MOVENEXT
'	WEND
		'response.End()
	
	if opt="blq" then
		st= "x"	
		outro="Bloqueio da P:"&periodo&",D:"&co_materia&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma
	
	elseif opt="dblq" then
		st= ""
		outro="Desbloqueio da P:"&periodo&",D:"&co_materia&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma
	end if

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


	Set RS_b = Server.CreateObject("ADODB.Recordset")
	sql_atualiza= "UPDATE TB_Da_Aula SET "&situcao_periodo&" = '"&st&"' "&SQL_ALUNOS
	response.Write(sql_atualiza)
	Set RS2 = CON2.Execute(sql_atualiza)

	call GravaLog (chave,outro)
	if opt="blq" or opt="fblq" then
	
		response.Redirect("altera.asp?or=01&opt=ok&curso="&curso&"&unidade="&unidade&"&etapa="&etapa&"&turma="&turma&"&periodo="&periodo&"&comando="&opt&"&ano="&ano_letivo)
	
	elseif opt="dblq" or opt="fdblq" then
	
		response.Redirect("altera.asp?or=02&opt=ok&curso="&curso&"&unidade="&unidade&"&etapa="&etapa&"&turma="&turma&"&periodo="&periodo&"&comando="&opt&"&ano="&ano_letivo)
	
	end if
end if	
 %>
 <%
If Err.number<>0 then
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