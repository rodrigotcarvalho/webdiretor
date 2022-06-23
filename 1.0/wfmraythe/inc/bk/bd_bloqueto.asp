<!--#include file="caminhos.asp"-->
<%
		Set CONbl = Server.CreateObject("ADODB.Connection") 
		ABRIRbl = "DBQ="& CAMINHO_bl & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONbl.Open ABRIRbl		

function MensagensBloqueto (P_CO_Matricula_Escola, P_Dat_venc, P_Num_Msg)

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT TX_Msg_0"&P_Num_Msg&" as Mensagem FROM TB_Bloqueto WHERE CO_Matricula_Escola ="& P_CO_Matricula_Escola&" AND DA_Vencimento =#"&P_Dat_venc&"#"
'response.write(SQL)
		RS.Open SQL, CONbl
		
	MensagensBloqueto = RS("Mensagem")	
		
end function

%>