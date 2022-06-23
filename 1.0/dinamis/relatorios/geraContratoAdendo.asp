<%
 ano_letivo = request.querystring("ano")
 co_aluno = request.querystring("aluno")
 tipo = request.querystring("tipo")
session("ano_letivo") = ano_letivo   
session("aluno_contrato") = co_aluno   
session("tipo_contrato") = tipo   
url = "contratos/"&tipo&".asp"
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/bd_parametros.asp"-->
<!--#include file="../inc/bd_alunos.asp"-->
<!--#include file="../inc/bd_contato.asp"-->
<!--#include file="../inc/bd_webfamilia.asp"-->
<!--#include file="../inc/funcoes_contratos.asp"-->
<%
if tipo = "CONTRATO_1A" then
%>
<!--#include file="contratos/CONTRATO_1A.asp"-->


<%
else
%>
<html>
<head>
</head>
<body> Erro em geraContratoAdendo.asp - Documento não localizado para <%response.write(ano_letivo&"-"&co_aluno&"-"&tipo)%>
</body>
</html>
<%end if%>



