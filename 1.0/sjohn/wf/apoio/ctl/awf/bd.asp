<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<%
chave = session("chave")

variavel= request.form("var_pub")
valor=request.form("valor_pub")

co=split(variavel,"_")

'for vp=0 to ubound(co)
'response.Write(co(vp)&"<BR>")
'next

unidade=co(1)
curso=co(2)

    	Set CON_WF = Server.CreateObject("ADODB.Connection") 
		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_WF.Open ABRIR_WF

if co(4)="t" then
	if co(6)=1 then
		wrk_variavel="CO_apr1"
	elseif co(6)=2 then	
		wrk_variavel="CO_apr2"
	elseif co(6)=3 then	
		wrk_variavel="CO_apr3"
	elseif co(6)=4 then	
		wrk_variavel="CO_apr4"
	elseif co(6)=5 then			
		wrk_variavel="CO_apr5"
	elseif co(6)=6 then		
		wrk_variavel="CO_apr6"
	elseif co(6)=7 then	
		wrk_variavel="CO_apr7"	
	end if
else
	if co(6)=1 then
		wrk_variavel="CO_prova1"
	elseif co(6)=2 then	
		wrk_variavel="CO_prova2"
	elseif co(6)=3 then	
		wrk_variavel="CO_prova3"
	elseif co(6)=4 then	
		wrk_variavel="CO_prova4"
	elseif co(6)=5 then			
		wrk_variavel="CO_prova5"
	elseif co(6)=6 then		
		wrk_variavel="CO_prova6"
	elseif co(6)=7 then	
		wrk_variavel="CO_prova7"	
	end if
end if


if co(5) = "b" then
	acao="Bloqueou "&wrk_variavel
	valor="D"
else
	acao="Desbloqueou "&wrk_variavel
	valor="L"	
end if	

Set RS = Server.CreateObject("ADODB.Recordset")
sql= "UPDATE TB_Autoriza_WF SET "&wrk_variavel &"='"&valor&"' WHERE NU_Unidade="&unidade&" AND CO_Curso='"&curso&"' AND CO_Etapa='"&co(3)&"'"
Set RS = CON_WF.Execute(sql)
response.write(sql)
response.Flush()
outro="P:"&co(6)&",U:"&unidade&",C:"&curso&",E:"&co(3)&",Acao:"&acao
call GravaLog (chave,outro)
%>