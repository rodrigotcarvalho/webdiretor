<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/bd_parametros.asp"-->
<!--#include file="../inc/bd_alunos.asp"-->
<!--#include file="../inc/bd_webfamilia.asp"-->

<%
	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1

    Set RSa = Server.CreateObject("ADODB.Recordset")
	SQLa = "SELECT * FROM TB_Alunos" 
	RSa.Open SQLa, CON1

while not RSa.EOF
resp = buscaResponsavelFinanceiro(RSa("CO_Matricula"))


ucet = buscaUCET(RSa("CO_Matricula"),2015)
vetorUCET = split(ucet,"#!#")
if ubound(vetorUCET)>=0 then
	nu_unidade =  vetorUCET(0)
	co_curso = vetorUCET(1)
	co_etapa = vetorUCET(2)
	co_turma = vetorUCET(3)
end if	

modelo = modeloContratoAdendo(nu_unidade,co_curso,co_etapa,co_turma,"C")
response.Write(resp&" "&RSa("CO_Matricula")&" "&modelo&"<BR>")
RSa.MOVENEXT
WEND
%>