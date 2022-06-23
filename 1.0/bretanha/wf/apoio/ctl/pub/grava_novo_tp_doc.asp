<!--#include file="../../../../inc/caminhos.asp" -->
<%
ano_letivo_wf = Session("ano_letivo_wf")

if transicao = "S" then
	area="wd"
	site_escola="simplynet2.tempsite.ws/wd/"&ambiente_escola&"/wf/apoio/ctl/pub/"
else
	if left(ambiente_escola,5)= "teste" then
		area="wdteste"
		site_escola="www.simplynet.com.br/"&area&"/"&ambiente_escola&"/wf/apoio/ctl/pub"
	else
		area="wd"
		site_escola="www.webdiretor.com.br/"&ambiente_escola&"/wf/apoio/ctl/pub"
	end if
end if

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
		
	maior_pasta=request.form("maior_pasta")
	nome_pasta=request.form("nome_pasta")
	espira=request.form("espira")
	dia_exp=request.form("dia_exp")
	mes_exp=request.form("mes_exp")
	data_exp=dia_exp&"/"&mes_exp&"/"&ano_letivo_wf
	
	maior_pasta=maior_pasta*1
	nova_pasta =maior_pasta+1 
	
	if espira="s" then
		espira=TRUE
	else	
		espira=FALSE
		data_exp=NULL
	end if	

	Set RS_updt = server.createobject("adodb.recordset")
	RS_updt.open "TB_Tipo_Pasta_Doc", CON0, 2, 2  'which table do you want open
	RS_updt.addnew
	
		RS_updt("CO_Pasta_Doc")=nova_pasta
		RS_updt("NO_Pasta") = nome_pasta
		RS_updt("IN_Expira") = espira
		RS_updt("DA_Expira") = data_exp

	RS_updt.update
	set RS_updt=nothing
	
CON0.Close
Set CON0 = Nothing
	
	
'RESPONSE.Write("http://"&site_escola&"/sndocs/criapasta.asp?al="&ano_letivo_wf&"&mp="&nova_pasta)
response.Redirect("http://"&site_escola&"/sndocs/criapasta.asp?al="&ano_letivo_wf&"&mp="&nova_pasta&"&env="&ambiente_escola)
%>