<%
vetor_cod_cons=request.form("alunos")
codigo_mat_prin = request.Form("mat_prin")	
periodo_escolhido = request.Form("periodo")	

obr=request.form("obr")
submit = request.Form("Submit")
session("vetor_cod_cons") = vetor_cod_cons
session("codigo_mat_prin") = codigo_mat_prin
session("periodo_escolhido") = periodo_escolhido
if Left(submit,7) = "Retirar" then 	
	response.Redirect("confirmar.asp?obr="&obr)	
elseif submit = "Incluir" then 	
	response.Redirect("altera.asp?ori=03&obr="&obr)
end if	
%>	