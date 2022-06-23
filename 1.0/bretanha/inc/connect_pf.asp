<!--#include file="caminhos.asp"-->
<%
ano_letivo = session("ano_letivo") 

	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR

	Set RSAV = Server.CreateObject("ADODB.Recordset")
	SQLAV = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&ano_letivo&"'"
	RSAV.Open SQLAV, CON		

	situacao_ano=RSAV("ST_Ano_Letivo")
	
	if situacao_ano="L" then
		ano_vigente=ano_letivo
	else	
		ano_vigente=DatePart("yyyy", now)
	end if		

		CAMINHO_pf = "e:\home\simplynet\dados\bretanha\BD\"&ano_vigente&"\Posicao.mdb"
%>