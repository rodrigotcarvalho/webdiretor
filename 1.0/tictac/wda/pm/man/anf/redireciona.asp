<%
nota_fiscal=request.form("nota_fiscal")
submit = request.Form("Submit")

nota_fiscal = replace(nota_fiscal,"/","$!$")


if submit = "Excluir" then 	
	response.Redirect("confirma.asp?opt=exc&cod="&nota_fiscal)	
else
	' Caso o usuário selecione mais de uma nota_fiscal, apenas a primeira é que poderá ser alterada ou ter o conteúdo incluído
	dados_nota_fiscal = 	split(nota_fiscal,",")
	nota_fiscal_encaminhar = dados_nota_fiscal(0)
	if submit = "Alterar" then 	
		response.Redirect("confirma.asp?opt=alt&cod="&nota_fiscal_encaminhar)
	else
		response.Redirect("../../../../relatorios/swd022.asp?opt=resumo&obr="&nota_fiscal_encaminhar)		
	end if	
end if		
%>	