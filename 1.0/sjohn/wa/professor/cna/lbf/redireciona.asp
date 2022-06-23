<%
vetor_cod_cons=request.form("alunos")
obr=request.form("obr")
submit = request.Form("Submit")
session("vetor_cod_cons") = vetor_cod_cons

if submit = "Retirar Bonus" then 	
	response.Redirect("confirmar.asp?obr="&obr)	
elseif submit = "Incluir" then 	
	response.Redirect("altera.asp?ori=03&obr="&obr)
end if	
%>	