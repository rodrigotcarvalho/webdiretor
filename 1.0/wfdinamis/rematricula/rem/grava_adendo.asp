<!--#include file="../../inc/caminhos.asp"-->
<!--#include file="../../inc/bd_webfamilia.asp"-->
<%
cod= request.form("cod_aluno")
mod_adendo= request.form("mod_adendo")
opcao_adendo= request.form("opcao_adendo")
gravado = gravaOpcaoAdendo(cod, mod_adendo,opcao_adendo)
%>
Aguarde alguns instantes e verifique sua caixa de email!