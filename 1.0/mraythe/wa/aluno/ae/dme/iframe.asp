<!--#include file="../../../../inc/graficos.asp"-->
<%
faixas=session("faixas")
categorias=session("categorias")
'response.Write("Pizza("&faixas&","&categorias)
call Pizza(faixas,categorias)
%>
