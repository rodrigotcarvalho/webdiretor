<!--#include file="../../inc/graficos.asp"-->
<%
faixas=replace(session("faixas"),",",".")
categorias=replace(session("categorias"),",",".")
'response.Write("Pizza("&faixas&","&categorias)
call ColunaAgrupada(faixas,categorias)
%>
