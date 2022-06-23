<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
</head>


<% 
opt = request.QueryString("opt")

chave=session("chave")
session("chave")=chave

ano = session("ano_letivo")
cod_prof = request.form("cod_prof")
curso = request.Form("curso")
unidade = request.Form("unidade")
co_etapa= request.Form("etapa")
turma= request.Form("turma")
mat_fil = request.form("mat_prin")
tabela = request.Form("tabela")
coordenador= request.Form("coordenador")
pendentes="Ajuste"


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		

if opt = "inc" then

		Set RS6 = Server.CreateObject("ADODB.Recordset")
		SQL6 = "SELECT * FROM TB_Materia where CO_Materia ='"& mat_fil &"'"
		RS6.Open SQL6, CON0
		
if RS6.EOF then	

mat_prin = " "
mat_fil = " "	

else
		
check = RS6("CO_Materia_Principal")		
		
if isnull(check) THEN		

mat_prin = mat_fil

else
mat_prin = mat_fil
mat_fil = RS6("CO_Materia_Principal")

end if
end if
' response.Write(check&"/"&mat_prin&"/"&mat_fil)
Set RS = server.createobject("adodb.recordset")
RS.open "TB_Da_Aula", CON, 2, 2 'which table do you want open

RS.addnew

RS("CO_Professor") = cod_prof
RS("CO_Curso") = curso
RS("NU_Unidade") = unidade
RS("CO_Etapa")= co_etapa
RS("CO_Turma")= turma
RS("CO_Materia_Principal") = mat_prin
RS("CO_Materia") = mat_fil
RS("TP_Nota") = tabela
RS("CO_Cord")= coordenador
RS.update
  
set RS=nothing

cod_prof = cod_prof

call GravaLog (chave,"Incluiu - PROF:"&cod_prof&"/U:"&unidade&"/C:"&curso&"/E:"&co_etapa&"/T:"&tabela&"/D:"&mat_fil)

response.Redirect("altera.asp?opt=ok&cod_cons="&cod_prof&"")

elseif opt="exc" then

ano_letivo = session("ano_letivo")
cod_prof =  request.form("cod_prof")
nome_prof = request.form("nome_prof")
co_usr_prof = request.form("co_usr_prof")
curso = request.Form("curso")
unidade = request.Form("unidade")
grade = request.Form("grade")

response.Write(grade)
session("cod_prof")=cod_prof
session("nome_prof")=nome_prof
session("co_usr_prof")=co_usr_prof

miss=0
vertorExclui = split(grade,", ")
for i =0 to ubound(vertorExclui)

exclui = split(vertorExclui(i),"-")

unidade = exclui(0)
curso= exclui(1)
co_etapa= exclui(2)
turma= exclui(3)
mat_prin= exclui(4)
mat_fil= exclui(5)
tabela= exclui(6)
coordenador= exclui(7)

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * from TB_Da_Aula where CO_Materia_Principal='"& mat_prin &"'AND CO_Materia='"& mat_fil &"'AND NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		RS.Open SQL, CON


	'	response.Write("SQL = SELECT * from TB_Da_Aula where AND CO_Materia_Principal='"& mat_prin &"'AND CO_Materia='"& mat_fil &"'AND NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'")
pr=0
While not RS.EOF
pr=pr+1
RS.MOVENEXT
wend

if pr=1 then

if tabela ="TB_NOTA_A" then
	CAMINHOn = CAMINHO_na

elseif tabela="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb
	
elseif tabela ="TB_NOTA_C" then
	CAMINHOn = CAMINHO_nc
	
elseif tabela ="TB_NOTA_D" then
	CAMINHOn = CAMINHO_nd

elseif tabela ="TB_NOTA_E" then
	CAMINHOn = CAMINHO_ne
		
elseif tabela ="TB_NOTA_F" then
	CAMINHOn = CAMINHO_nf	
		
elseif tabela ="TB_NOTA_K" then
	CAMINHOn = CAMINHO_nk							
end if	

		Set CON_A = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_A.Open ABRIR
		
		Set CON_N = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIR3
		
		Set RSa = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RSa = CON_A.Execute(SQL_A)

While Not RSa.EOF
nu_matricula = RSa("CO_Matricula")

	  	Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL_N = "Select * from "& tabela &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia_Principal='"& mat_fil &"'AND CO_Materia='"& mat_prin &"'"
		Set RS3 = CON_N.Execute(SQL_N)
if RS3.EOF then

'response.Write("SQL = DELETE * from TB_Da_Aula where CO_Professor="& cod_prof &" AND CO_Materia_Principal='"& mat_prin &"'AND CO_Materia='"& mat_fil &"'AND NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'")

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "DELETE * from TB_Da_Aula where CO_Professor="& cod_prof &" AND CO_Materia_Principal='"& mat_prin &"'AND CO_Materia='"& mat_fil &"'AND NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		RS.Open SQL, CON
		
cod_prof = cod_prof

call GravaLog (chave,"Excluiu - PROF:"&cod_prof&"/U:"&unidade&"/C:"&curso&"/E:"&co_etapa&"/T:"&tabela&"/D:"&mat_fil)
miss=miss
RSa. Movenext
else

miss=miss+1
RSa. Movenext
end if
wend
'caso a turma não possua alunos

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "DELETE * from TB_Da_Aula where CO_Professor="& cod_prof &" AND CO_Materia_Principal='"& mat_prin &"'AND CO_Materia='"& mat_fil &"'AND NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		RS.Open SQL, CON

call GravaLog (chave,"Excluiu - PROF:"&cod_prof&"/U:"&unidade&"/C:"&curso&"/E:"&co_etapa&"/T:"&tabela&"/D:"&mat_fil)
else

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "DELETE * from TB_Da_Aula where CO_Professor="& cod_prof &" AND CO_Materia_Principal='"& mat_prin &"'AND CO_Materia='"& mat_fil &"'AND NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		RS.Open SQL, CON
		
cod_prof = cod_prof
miss=miss
call GravaLog (chave,"Excluiu - PROF:"&cod_prof&"/U:"&unidade&"/C:"&curso&"/E:"&co_etapa&"/T:"&tabela&"/D:"&mat_fil)
end if

pendentes= pendentes&", "&unidade&"-"&curso&"-"&co_etapa&"-"&turma&"-"&mat_prin&"-"&mat_fil&"-"&tabela&"-"&coordenador
next



if miss=0 then
response.Redirect("altera.asp?opt=ok2&cod_cons="&cod_prof)

else
session("pendentes")=pendentes
response.Redirect("bd2.asp?opt=sw")
end if
end if



%>

</html>
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