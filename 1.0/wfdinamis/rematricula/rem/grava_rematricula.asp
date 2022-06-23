<!--#include file="../../inc/caminhos.asp"-->
<!--#include file="../../inc/bd_webfamilia.asp"-->
<%
cod= request.form("cod_aluno")
tipo_contrato= session("versao_contrato_adendo")
gravado = gravaRematricula(cod, tipo_contrato)
%>
Aguarde alguns instantes e verifique sua caixa de email!