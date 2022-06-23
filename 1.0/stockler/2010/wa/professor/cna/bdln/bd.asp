<%On Error Resume Next%>

<!--#include file="../../../../inc/funcoes.asp"-->
<%
opt=request.form("opt")
co_prof = request.form("prof")
grau = request.Form("curso")
unidades = request.Form("unidade")
serie= request.Form("etapa")
turma= request.Form("turma")
co_materia = request.form("mat")
periodo = request.form("periodo")
nota= request.form("nota")

chave=session("chave")
session("chave")=chave


	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
periodo = periodo*1
if opt="blq" then
st= "x"


outro="Bloqueio da P:"&periodo&",D:"&co_materia&",U:"&unidades&",C:"&grau&",E:"&serie&",T:"&turma&""

elseif opt="dblq" then
st= ""

outro="Desbloqueio da P:"&periodo&",D:"&co_materia&",U:"&unidades&",C:"&grau&",E:"&serie&",T:"&turma&""

end if
'response.Write("sql_atualiza= UPDATE TB_Da_Aula SET ST_Per_1 = '"&st&"' WHERE CO_Professor="& co_prof &" AND CO_Materia='"& co_materia &"' AND NU_Unidade="& unidades &" AND CO_Curso='"& grau &"' AND CO_Etapa='"& serie &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& nota &"'")

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

sql_atualiza= "UPDATE TB_Da_Aula SET "&situcao_periodo&" = '"&st&"' WHERE CO_Professor="& co_prof &" AND CO_Materia_principal='"& co_materia &"' AND NU_Unidade="& unidades &" AND CO_Curso='"& grau &"' AND CO_Etapa='"& serie &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& nota &"'"
Set RS2 = CON.Execute(sql_atualiza)

call GravaLog (chave,outro)

if opt="blq" then

response.Redirect("altera.asp?or=01&opt=ok&curso="&grau&"&unidade="&unidades&"&etapa="&serie&"&turma="&turma&"&ano="&ano_letivo)

elseif opt="dblq" then

response.Redirect("altera.asp?or=02&opt=ok&curso="&grau&"&unidade="&unidades&"&etapa="&serie&"&turma="&turma&"&ano="&ano_letivo)

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