<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<%
opt = REQUEST.QueryString("obr")
p = REQUEST.QueryString("p")

obr = split(opt,"?")

unidade = obr(0)
curso = obr(1)
co_etapa = obr(2)
turma = obr(3)
ano_letivo = obr(4)

obr=unidade&"?"&curso&"?"&co_etapa&"?"&turma&"?"&ano_letivo



		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2

		Set RSTB = Server.CreateObject("ADODB.Recordset")
		CONEXAOTB = "Select * from TB_Da_Aula WHERE NU_Unidade="& unidade &" AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		Set RSTB = CON2.Execute(CONEXAOTB)
		
nota= RSTB("TP_Nota")

if nota = "TB_NOTA_A" Then		
		response.Redirect("imprime_a.asp?p="&p&"&obr="&obr)
elseif nota = "TB_NOTA_B" Then
		response.Redirect("imprime_b.asp?p="&p&"&obr="&obr)
elseif nota = "TB_NOTA_C" Then
		response.Redirect("imprime_c.asp?p="&p&"&obr="&obr)
elseif nota = "TB_NOTA_D" Then
		response.Redirect("imprime_d.asp?p="&p&"&obr="&obr)	
elseif nota = "TB_NOTA_E" Then
		response.Redirect("imprime_e.asp?p="&p&"&obr="&obr)		
elseif nota = "TB_NOTA_F" Then
		response.Redirect("imprime_f.asp?p="&p&"&obr="&obr)			
elseif nota = "TB_NOTA_K" Then
		response.Redirect("imprime_k.asp?p="&p&"&obr="&obr)						
else
response.Write("ERRO! Não existe tabela de notas associada a esta turma.")
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