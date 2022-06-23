<!--#include file="../../inc/caminhos.asp"-->
<!--#include file="../../inc/bd_webfamilia.asp"-->
<%
cod= request.form("cod_aluno")
gravaRematricula(cod)
%>
Boleto gerado com sucesso!