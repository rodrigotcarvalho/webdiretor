<%session("ano_letivo")  = 2010%>
<!--#include file="inc/caminhos.asp"-->
<%
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
response.Write(ABRIR0)
		CON0.Open ABRIR0
		
    	Set CON_WF = Server.CreateObject("ADODB.Connection") 
		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		response.Write(ABRIR_WF)
		CON_WF.Open ABRIR_WF		
		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT DISTINCT NU_Unidade FROM TB_Unidade"
		RS0.Open SQL0, CON0
		
		
While not RS0.EOF
	unidade = RS0("NU_Unidade")



	Set RS1 = Server.CreateObject("ADODB.Recordset")
	SQL1 = "SELECT DISTINCT CO_Curso FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade
	RS1.Open SQL1, CON0
		
		
	While not RS1.EOF
		curso = RS1("CO_Curso")
		
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"'"
		RS2.Open SQL2, CON0
					
		While not RS2.EOF
			etapa = RS2("CO_Etapa")	
			
		
			for per= 1 to 14
				if per=1 then	
					wrk_variavel="CO_apr1"
				elseif per=2 then	
					wrk_variavel="CO_apr2"
				elseif per=3 then	
					wrk_variavel="CO_apr3"
				elseif per=4 then	
					wrk_variavel="CO_apr4"
				elseif per=5 then			
					wrk_variavel="CO_apr5"
				elseif per=6 then		
					wrk_variavel="CO_apr6"
				elseif per=7 then	
					wrk_variavel="CO_apr7"	
				elseif per=8 then
					wrk_variavel="CO_prova1"
				elseif per=9 then	
					wrk_variavel="CO_prova2"
				elseif per=10 then	
					wrk_variavel="CO_prova3"
				elseif per=11 then	
					wrk_variavel="CO_prova4"
				elseif per=12 then			
					wrk_variavel="CO_prova5"
				elseif per=13 then		
					wrk_variavel="CO_prova6"
				elseif per=14 then	
					wrk_variavel="CO_prova7"	
				end if
			
			
	
			
			
				Set RS = Server.CreateObject("ADODB.Recordset")
				sql= "UPDATE TB_Autoriza_WF SET "&wrk_variavel &"='D' WHERE NU_Unidade="&unidade&" AND CO_Curso='"&curso&"' AND CO_Etapa='"&etapa&"'"
				RESPONSE.Write(sql&"<br>")
				Set RS = CON_WF.Execute(sql)
			next
		RS2.movenext
		WEND	
	RS1.movenext
	WEND		
RS0.movenext
WEND
%>	

