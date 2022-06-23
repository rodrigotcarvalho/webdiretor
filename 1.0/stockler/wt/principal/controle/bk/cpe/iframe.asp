<!--#include file="../../../../inc/graficos.asp"-->
<%
faixas=session("faixas")
categorias=session("categorias")
legendas=session("legendas")
'response.Write("Pizza("&faixas&","&categorias)
call ColunaAgrupadaParam(faixas,categorias,"S",legendas,0,0,0)
%>
