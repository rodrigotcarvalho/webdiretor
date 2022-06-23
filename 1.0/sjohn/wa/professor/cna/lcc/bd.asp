<!--#include file="../../../../inc/caminhos.asp" -->
<!--#include file="../../../../inc/funcoes.asp" -->
<%
opt = request.QueryString("opt")
obr = request.Form("obr")	
	Set CONG = Server.CreateObject("ADODB.Connection") 
	ABRIRG = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONG.Open ABRIRG
	
	vetor_cod_cons = request.Form("vetor_cod_cons")	
    codigo_mat_prin = request.Form("mat_prin")	
    periodo_coc = request.Form("periodo")	 	
	
	if opt="exc" then
	
		vetor_aluno_matricula= request.Form("vetor_aluno_matricula")	
		vetor_cod_cons = split(vetor_aluno_matricula,",")
		for j = 0 to ubound(vetor_cod_cons)
		
			vetor_exclui_aluno = split(vetor_cod_cons(j),"#!#")
			nu_matricula = vetor_exclui_aluno(0)
			codigo_mat_prin = vetor_exclui_aluno(1)		

			Set RSDBA = Server.CreateObject("ADODB.Recordset")
			SQLDBA = "DELETE * from TB_COC WHERE CO_Matricula = "&nu_matricula&" And CO_Materia= '"&codigo_mat_prin&"'"	
			Set RSDBA = CONG.Execute(SQLDBA)	
			outro=Left("Excluir,Matrics:"&vetor_cod_cons(j),255)
	
			call GravaLog (session("nvg"),outro)	
		next	
		
		response.Redirect("select_alunos.asp?opt=ok&vt=s&obr="&obr)			
	else

		obr = request.Form("obr")	
		ori = request.Form("ori")	
		
		session("vetor_cod_cons") = vetor_cod_cons	
		session("obr") = obr	
		session("ori") = ori			
		
		Set RSDBA = Server.CreateObject("ADODB.Recordset")
		SQLDBA = "DELETE * from TB_COC WHERE CO_Matricula IN ("& vetor_cod_cons &") And CO_Materia= '"&codigo_mat_prin&"'"
	
		Set RSDBA = CONG.Execute(SQLDBA)
			
		cod_cons = split(vetor_cod_cons,", ")
		For bma=0 to ubound(cod_cons)		
		
			val_bonus = request.Form("bonus_"&cod_cons(bma))		
		
			Set RS = server.createobject("adodb.recordset")		
			RS.open "TB_COC", CONG, 2, 2 'which table do you want open
			RS.addnew
			
				RS("CO_Matricula") = cod_cons(bma)
				RS("CO_Materia") = codigo_mat_prin			
				if periodo_coc="T" then
					RS("STatus1") = "APC" 
				elseif periodo_coc="F" then	
					RS("STatus2") = "APC" 
				elseif periodo_coc="R" then			
					RS("STatus3") = "APC" 	
				end if	
			
			RS.update
			set RS=nothing			
		Next
		
		outro=Left("Incluir,Matrics:"&vetor_cod_cons,255)

		call GravaLog (session("nvg"),outro)		
		
		response.Redirect("index.asp?opt=ok&nvg="&session("nvg"))
	end if	
%>