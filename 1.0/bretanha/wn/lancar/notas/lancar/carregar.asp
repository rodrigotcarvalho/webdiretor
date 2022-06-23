<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes7.asp"-->


<%
nivel = 4
obr=request.form("obr")
nota=request.form("nota")



chave = session("chave")
session("chave")=chave
dados= split(obr, "$!$" )
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

	Set CONG = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONG.Open ABRIR

	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set CON2 = Server.CreateObject("ADODB.Connection") 
	ABRIR2 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON2.Open ABRIR2
	
	Set CON_wr = Server.CreateObject("ADODB.Connection") 
	ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_wr.Open ABRIR_wr
	

	consulta2d = "select * from TB_Materia where CO_Materia='"&co_materia&"'"
	set RS2d = CON2.Execute (consulta2d)
	
NO_Materia = RS2d("NO_Materia")
mat_princ=RS2d("CO_Materia_Principal")

if mat_princ="" or isnull(mat_princ) then
	mat_princ=co_materia
end if


	wrk_ind_carregado = CarregaTotalFaltas(co_prof, unidades, grau, serie, turma, mat_princ, co_materia, periodo, outro)  
	
	Set RS1 = Server.CreateObject("ADODB.Recordset")
	consulta1 = "select * from Email where CO_Escola="&escola
	set RS1 = CON_wr.Execute (consulta1)
	
	'response.Write(consulta1)
	
	mail_suporte=RS1("Suporte")
	mail_CC=RS1("Mail_Simplynet")
	mail_rodan=RS1("Mail_Rodan")
		
	consulta1 = "select * from TB_Usuario where CO_Usuario="&co_usr
	set RS1 = CON.Execute (consulta1)
	
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


	consulta = "select * from TB_Da_Aula where CO_Professor="&co_prof&" AND NU_Unidade="&unidades&" AND CO_Curso='"&grau&"' AND CO_Etapa='"&serie&"' AND CO_Materia_Principal='"&co_materia&"'"
'response.Write(consulta&"-")
	set RS = CONG.Execute (consulta)
	
coord = RS("CO_Cord")

	consulta_mail = "select * from TB_Usuario where CO_Usuario="&coord
	set RS_mail = CON.Execute (consulta_mail)
	
'response.Write(consulta_mail)
'response.end()
	if RS_mail.EOF then
		mail = "webdiretor@gmail.com"
		'response.write("<font class=form_corpo>Não é possível enviar a mensagem, pois o Coordenador não possui e-mail cadastrado.<br><a href=javascript:window.history.go(-1)>voltar</a></font>")
		'response.end()	
	else
		mail = RS_mail("Email_Usuario")
		'mail = mail_rodan
		if isnull(mail) or mail="" then
		response.write("<font class=form_corpo>Não é possível enviar a mensagem, pois o Coordenador codígo "&coord&" não possui e-mail cadastrado.<br><a href=javascript:window.history.go(-1)>voltar</a></font>")
		response.end()
	end if
end if


st= "x"
periodo = periodo*1

if periodo=1 then
	sql_atualiza= "UPDATE TB_Da_Aula SET ST_Per_1 = '"&st&"' WHERE CO_Professor="& co_prof &" AND CO_Materia_Principal='"& co_materia &"' AND NU_Unidade="& unidades &" AND CO_Curso='"& grau &"' AND CO_Etapa='"& serie &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& nota &"'"
elseif periodo=2 then
	sql_atualiza= "UPDATE TB_Da_Aula SET ST_Per_2 = '"&st&"' WHERE CO_Professor="& co_prof &" AND CO_Materia_Principal='"& co_materia &"' AND NU_Unidade="& unidades &" AND CO_Curso='"& grau &"' AND CO_Etapa='"& serie &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& nota& "'"
elseif periodo =3 then
	sql_atualiza= "UPDATE TB_Da_Aula SET ST_Per_3 = '"&st&"' WHERE CO_Professor="& co_prof &" AND CO_Materia_Principal='"& co_materia &"' AND NU_Unidade="& unidades &" AND CO_Curso='"& grau &"' AND CO_Etapa='"& serie &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& nota &"'"
elseif periodo =4 then
	sql_atualiza= "UPDATE TB_Da_Aula SET ST_Per_4 = '"&st&"' WHERE CO_Professor="& co_prof &" AND CO_Materia_Principal='"& co_materia &"' AND NU_Unidade="& unidades &" AND CO_Curso='"& grau &"' AND CO_Etapa='"& serie &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& nota &"'"
elseif periodo =5 then
	sql_atualiza= "UPDATE TB_Da_Aula SET ST_Per_5 = '"&st&"' WHERE CO_Professor="& co_prof &" AND CO_Materia_Principal='"& co_materia &"' AND NU_Unidade="& unidades &" AND CO_Curso='"& grau &"' AND CO_Etapa='"& serie &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& nota &"'"
elseif periodo =6 then
	sql_atualiza= "UPDATE TB_Da_Aula SET ST_Per_6 = '"&st&"' WHERE CO_Professor="& co_prof &" AND CO_Materia_Principal='"& co_materia &"' AND NU_Unidade="& unidades &" AND CO_Curso='"& grau &"' AND CO_Etapa='"& serie &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& nota &"'"
end if
'response.Write(sql_atualiza)
'response.end()
Set RS2 = CONG.Execute(sql_atualiza)


mensagem="O(A) Professor(a) "& nome &" lançou todas as notas de "& NO_Materia &" do "& no_serie &" do "& no_grau &", unidade: "& no_unidades &", turma "& turma&" do Periodo "& periodo&""
'Dim objCDO
Set objCDO = Server.CreateObject("CDONTS.NewMail")
objCDO.From = mail_suporte
objCDO.To = mail
objCDO.CC = ""
objCDO.BCC = mail_CC
objCDO.Subject = "Confirmação do Lançamento de Notas através do Sistema Web Diretor"
objCDO.Body = mensagem
objCDO.Send()
Set objCDO = Nothing

'  else
' Se não for enviado mostra o erro que ocoreu
'     Response.Write ("Ocorreu um erro.<BR>")
'     Response.Write ("O Erro é " & Mailer.Response)
'  end if

'response.Redirect("index.asp?nvg="&chave&"&ori=01&opt=cln")
'response.Write("OK")
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