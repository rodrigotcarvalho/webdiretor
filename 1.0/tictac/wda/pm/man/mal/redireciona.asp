<%
pedido=request.form("pedido")
submit = request.Form("Submit")
dados_msg = request.Form("dados_msg")
pedido = replace(pedido,"/","$!$")


if submit = "Excluir" then 	
	response.Redirect("confirma.asp?opt=exc&cod="&pedido)	
elseif submit = "Cancelar" then 	
	response.Redirect("confirma.asp?opt=exc&cod="&pedido)		
else
	' Caso o usuário selecione mais de uma nota_fiscal, apenas a primeira é que poderá ser alterada ou ter o conteúdo incluído
	dados_pedido = 	split(pedido,",")
	pedido_encaminhar = dados_pedido(0)
	if submit = "Alterar" then 	
		response.Redirect("confirma.asp?opt=alt&cod="&pedido_encaminhar)
	else	
		response.Redirect("../../../../relatorios/swd024.asp?obr="&pedido_encaminhar&"'")		
	end if	
end if		
%>	