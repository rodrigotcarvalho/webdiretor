<!--#include file="../../../../inc/graficos.asp"-->
<%
tipo_grafico=request.QueryString("opt")
faixas=session("faixas")
categorias=session("categorias")
'response.Write("Pizza("&faixas&","&categorias)
'call Pizza(faixas,categorias)

Response.Write("Gerar Grafico do tipo "&tipo_grafico)
%>
