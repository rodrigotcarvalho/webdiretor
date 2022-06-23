<!--#include file="../../../../inc/graficos.asp"-->
<%
faixas=session("faixas")
categorias=session("categorias")
tp_grafico=session("tp_grafico")
call StackedColuna_2D_ou_3D(faixas,categorias,tp_grafico)
%>
