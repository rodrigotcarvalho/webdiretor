<!--#include file="../inc/boletos.asp"-->
<% 

dados = request.form("vencimento")
cod_matric = request.querystring("c")
mes_solici=request.querystring("opt")
de = request.QueryString("de")
ate = request.QueryString("ate")
ucet = request.QueryString("ucet")	
tipo = request.querystring("tp")
restricao = request.querystring("r")

	boletoGerado = GeraBoletos(1, dados, cod_matric, mes_solici, ucet, de, ate, tipo, restricao)
%>

