<!--#include file="caminhos.asp"-->
<!--#include file="funcoes6.asp"-->
<!--#include file="../../global/tabelas_escolas.asp"-->
<%
opt=request.querystring("opt")
		
this_file = Request.ServerVariables("SCRIPT_NAME")
arPath = Split(this_Path, "/")

nivel=4


		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON_wr = Server.CreateObject("ADODB.Connection") 
		ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_wr.Open ABRIR_wr		

		Set CONg = Server.CreateObject("ADODB.Connection") 
		ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONg.Open ABRIRg
		
		Set CON_AL = Server.CreateObject("ADODB.Connection") 
		ABRIR_AL = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_AL.Open ABRIR_AL		
	
		Set CON_WF = Server.CreateObject("ADODB.Connection") 
		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_WF.Open ABRIR_WF	
	
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT	
%>
<div align="center">
<%if opt="al" then
ano_letivo_altera=Request.Form("ano_pub")

session("ano_letivo") = ano_letivo_altera
end if
%>
 
