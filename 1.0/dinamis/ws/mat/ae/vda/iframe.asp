<!--#include file="../../../../inc/graficos.asp"-->
<%
tipo_grafico=request.QueryString("opt")
faixas=session("faixas")
categorias=session("categorias")
legenda=session("legenda")
series=session("series")
if tipo_grafico="pizza" then
	call Pizza(faixas,categorias)
elseif tipo_grafico="barra" then
	call Barra(faixas,categorias)
elseif tipo_grafico="coluna" then
	call Coluna(faixas,categorias,legenda,series,outro)	
elseif tipo_grafico="coluna_empilhada" then
	call ColunaEmpilhada(faixas,categorias,legenda,series,outro)
elseif tipo_grafico="coluna_agrupada" then
	call ColunaAgrupada_s3D(faixas,categorias)	
end if
%>
