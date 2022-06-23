<!--#include file="../../../../inc/graficos.asp"-->
<%
tipo_grafico=request.QueryString("opt")
faixas=session("faixas")
categorias=session("categorias")

if tipo_grafico="pizza" then
	call Pizza(faixas,categorias)
elseif tipo_grafico="barra" then
	call Barra(faixas,categorias)
end if

%>
