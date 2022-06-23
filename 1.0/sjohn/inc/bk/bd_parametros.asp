<!--#include file="caminhos.asp"-->
<%
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
	
Function busca_materia_mae(codigo_materia)	

		Set RSBMM = Server.CreateObject("ADODB.Recordset")
		SQLBMM = "Select * from TB_Materia WHERE CO_Materia ='"& codigo_materia &"'"
		Set RSBMM = CONG.Execute(SQLBMM)
			
		if RSBMM.EOF then
			materia_mae = ""	
		else			
			materia_mae	 = RSBMM("CO_Materia_Principal")			
		end if	
		if materia_mae="" or isnull(materia_mae) then
			materia_mae = codigo_materia
		end if
		busca_materia_mae = materia_mae
end function
	
%>
