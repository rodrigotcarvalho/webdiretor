<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<%obr=request.QueryString("obr")
nota=request.QueryString("nota")
nvg = session("chave")
chave=nvg
session("chave")=chave
dados= split(obr, "?" )
co_materia = dados(0)
unidades= dados(1)
grau= dados(2)
serie= dados(3)
turma= dados(4)
periodo = dados(5)
ano_letivo = dados(6)
co_prof = dados(7)
co_usr = session("co_usr")
escola= session("escola")


	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR

	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1
	
	Set CON2 = Server.CreateObject("ADODB.Connection") 
	ABRIR2 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON2.Open ABRIR2
	
	Set CON_wr = Server.CreateObject("ADODB.Connection") 
	ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_wr.Open ABRIR_wr
	
		Set RS1 = Server.CreateObject("ADODB.Recordset")
	consulta1 = "select * from Email where CO_Escola="&escola
	set RS1 = CON_wr.Execute (consulta1)
	
	mail_suporte=RS1("Suporte")
	mail_CC=RS1("Mail_Simplynet")
		
	consulta1 = "select * from TB_Usuario where CO_Usuario="&co_usr
	set RS1 = CON1.Execute (consulta1)
	
nome = RS1("NO_Usuario")

	consulta2a = "select * from TB_Unidade where NU_Unidade="&unidades
	set RS2a = CON2.Execute (consulta2a)
	
no_unidades = RS2a("NO_Unidade")

	consulta2b = "select * from TB_Etapa where CO_Curso='"&grau&"' AND CO_Etapa='"&serie&"'"
	set RS2b = CON2.Execute (consulta2b)
	
no_serie = RS2b("NO_Etapa")

	consulta2c = "select * from TB_Curso where CO_Curso='"&grau&"'"
	set RS2c = CON2.Execute (consulta2c)
	
no_grau = RS2c("NO_Curso")

	consulta2d = "select * from TB_Materia where CO_Materia='"&co_materia&"'"
	set RS2d = CON2.Execute (consulta2d)
	
NO_Materia = RS2d("NO_Materia")


	consulta = "select * from TB_Da_Aula where CO_Professor="&co_prof&" AND NU_Unidade="&unidades&" AND CO_Curso='"&grau&"' AND CO_Etapa='"&serie&"' AND CO_Materia='"&co_materia&"'"
	set RS = CON.Execute (consulta)
	
coord = RS("CO_Cord")

	consulta_mail = "select * from TB_Usuario where CO_Usuario="&coord
	set RS_mail = CON1.Execute (consulta_mail)
	
mail = RS_mail("Email_Usuario")
'mail ="webdiretor@gmail.com"

'response.Write(mail&"<<")

'response.end()
st= "x"
periodo = periodo*1

if periodo=1 then
sql_atualiza= "UPDATE TB_Da_Aula SET ST_Per_1 = '"&st&"' WHERE CO_Professor="& co_prof &" AND CO_Materia_Principal='"& co_materia &"' AND NU_Unidade="& unidades &" AND CO_Curso='"& grau &"' AND CO_Etapa='"& serie &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& nota &"'"
Set RS2 = CON.Execute(sql_atualiza)
elseif periodo=2 then
sql_atualiza= "UPDATE TB_Da_Aula SET ST_Per_2 = '"&st&"' WHERE CO_Professor="& co_prof &" AND CO_Materia_Principal='"& co_materia &"' AND NU_Unidade="& unidades &" AND CO_Curso='"& grau &"' AND CO_Etapa='"& serie &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& nota& "'"
Set RS2 = CON.Execute(sql_atualiza)
elseif periodo =3 then
sql_atualiza= "UPDATE TB_Da_Aula SET ST_Per_3 = '"&st&"' WHERE CO_Professor="& co_prof &" AND CO_Materia_Principal='"& co_materia &"' AND NU_Unidade="& unidades &" AND CO_Curso='"& grau &"' AND CO_Etapa='"& serie &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& nota &"'"
Set RS2 = CON.Execute(sql_atualiza)
end if

'response.Write("sql_atualiza= UPDATE TB_Da_Aula SET ST_Per_1 = '"&st&"' WHERE CO_Professor="& co_prof &" AND CO_Materia_Principal='"& co_materia &"' AND NU_Unidade="& unidades &" AND CO_Curso='"& grau &"' AND CO_Etapa='"& serie &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& nota &"'")


mensagem="O(A) Professor(a) "& nome &" lan�ou todas as notas de "& NO_Materia &" do "& no_serie &" do "& no_grau &", unidade: "& no_unidades &", turma "& turma&" do Periodo "& periodo&""
'Dim objCDO
Set objCDO = Server.CreateObject("CDONTS.NewMail")
objCDO.From = mail_suporte
objCDO.To = mail
objCDO.CC = ""
objCDO.BCC = mail_CC
objCDO.Subject = "Confirma��o do Lan�amento de Notas atrav�s do Sistema Web Diretor"
objCDO.Body = mensagem
objCDO.Send()
Set objCDO = Nothing

  'else
 'Se n�o for enviado mostra o erro que ocoreu
    ' Response.Write ("Ocorreu um erro.<BR>")
    ' Response.Write ("O Erro � " & Mailer.Response)
 ' end if

response.Redirect("index.asp?nvg="&nvg&"&ori=01&opt=cln")
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