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

<%elseif tipo = "CONTRATO_1B" then %>
<!--#include file="contratos/CONTRATO_1B.asp"-->

<%elseif tipo = "CONTRATO_2A" then %>
<!--#include file="contratos/CONTRATO_2A.asp"-->

<%elseif tipo = "CONTRATO_2B" then %>
<!--#include file="contratos/CONTRATO_2B.asp"-->

<%elseif tipo = "CONTRATO_3" then %>
<!--#include file="contratos/CONTRATO_3.asp"-->

<%elseif tipo = "CONTRATO_5" then %>
<!--#include file="contratos/CONTRATO_5.asp"-->

<%elseif tipo = "CONTRATO_8" then %>
<!--#include file="contratos/CONTRATO_8.asp"-->

<%elseif tipo = "CONTRATO_G1" then %>
<!--#include file="contratos/CONTRATO_G1.asp"-->

<%elseif tipo = "CONTRATO_G2" then %>
<!--#include file="contratos/CONTRATO_G2.asp"-->

<%elseif tipo = "CONTRATO_G3" then %>
<!--#include file="contratos/CONTRATO_G3.asp"-->

<%elseif tipo = "CONTRATO_G4" then %>
<!--#include file="contratos/CONTRATO_G4.asp"-->

<%
'ADENDOS------------------------------------------
elseif tipo = "ADENDO_1A" then %>
<!--#include file="contratos/ADENDO_1A.asp"-->

<%elseif tipo = "ADENDO_2A" then %>
<!--#include file="contratos/ADENDO_2A.asp"-->

<%elseif tipo = "ADENDO_2B" then %>
<!--#include file="contratos/ADENDO_2B.asp"-->

<%elseif tipo = "ADENDO_3" then %>
<!--#include file="contratos/ADENDO_3.asp"-->

<%elseif tipo = "ADENDO_5" then %>
<!--#include file="contratos/ADENDO_5.asp"-->

<%elseif tipo = "ADENDO_8" then %>
<!--#include file="contratos/ADENDO_8.asp"-->

<%elseif tipo = "ADENDO_G1" then %>
<!--#include file="contratos/ADENDO_G1.asp"-->

<%elseif tipo = "ADENDO_G2" then %>
<!--#include file="contratos/ADENDO_G2.asp"-->

<%elseif tipo = "ADENDO_G3" then %>
<!--#include file="contratos/ADENDO_G3.asp"-->

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



