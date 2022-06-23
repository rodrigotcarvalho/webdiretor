<!--#include file="../../../../inc/caminhos.asp" -->
<!--#include file="../../../../inc/funcoes.asp" -->
<%
opt = request.QueryString("opt")

	Set CONG = Server.CreateObject("ADODB.Connection") 
	ABRIRG = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONG.Open ABRIRG
	
	if opt="exc" then

		vetor_cod_cons = session("vetor_cod_cons")
		Set RSDBA = Server.CreateObject("ADODB.Recordset")
		SQLDBA = "DELETE * from TB_Bonus_Media_Anual WHERE CO_Matricula IN ("& vetor_cod_cons &")"	
		Set RSDBA = CONG.Execute(SQLDBA)	
		
		outro=Left("Excluir,Matrics:"&vetor_cod_cons,255)

		call GravaLog (session("nvg"),outro)		
		
		response.Redirect("select_alunos.asp?opt=ok")			
	else
		vetor_cod_cons = request.Form("vetor_cod_cons")
		obr = request.Form("obr")	
		ori = request.Form("ori")	
		
		session("vetor_cod_cons") = vetor_cod_cons	
		session("obr") = obr	
		session("ori") = ori			
		
		Set RSDBA = Server.CreateObject("ADODB.Recordset")
		SQLDBA = "DELETE * from TB_Bonus_Media_Anual WHERE CO_Matricula IN ("& vetor_cod_cons &")"
	
		Set RSDBA = CONG.Execute(SQLDBA)
			
		cod_cons = split(vetor_cod_cons,", ")
		For bma=0 to ubound(cod_cons)		
		
			val_bonus = request.Form("bonus_"&cod_cons(bma))		
		
			Set RS = server.createobject("adodb.recordset")		
			RS.open "TB_Bonus_Media_Anual", CONG, 2, 2 'which table do you want open
			RS.addnew
			
				RS("CO_Matricula") = cod_cons(bma)
				RS("bonus") = val_bonus			
			
			RS.update
			set RS=nothing			
		Next
		
		outro=Left("Incluir,Matrics:"&vetor_cod_cons,255)

		call GravaLog (session("nvg"),outro)		
		
		response.Redirect("index.asp?opt=ok&nvg="&session("nvg"))
	end if	
%>