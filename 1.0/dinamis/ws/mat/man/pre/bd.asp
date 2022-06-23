<!--#include file="../../../../inc/caminhos.asp"-->
<%
	Set conw = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
	conw.Open ABRIR
	
	
	  sql="UPDATE TB_Ano_Letivo SET "
	  sql=sql & "DT_Inicio_Rematricula=#" & request.Form("dataLancamentoInicio") & "#,"
	  sql=sql & "DT_Final_Rematricula=#" & request.Form("dataLancamentoFim") & "#,"
	  sql=sql & "DT_Bloqueto_Rematricula=#" & request.Form("dataLancamentoBloqueto") & "#"
	  sql=sql & " WHERE NU_Ano_Letivo='" & Session("ano_letivo")&"'" 
	  on error resume next
	  conw.Execute sql
	  if Err.number<>0 then
		response.write("Erro "&Err.Description)
	  else
	    conw.close
		response.Redirect("index.asp?nvg=WS-MA-MA-PRE&opt=ok")
	  end if
%>
	


