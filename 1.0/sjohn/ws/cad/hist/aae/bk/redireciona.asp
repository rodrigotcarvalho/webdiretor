<!--#include file="../../../../inc/caminhos.asp"-->
<%
historico=request.form("historico")
submit = request.Form("Submit")
historico = replace(historico,", ", ",")
if submit = "Excluir" then 	
	response.Redirect("confirma.asp?opt=exc&dad="&historico)	
else
	' Caso o usuário selecione mais de um pedido, apenas o primeiro é que poderá ser alterado

	dados_historico = 	split(historico,",")
	historico_encaminhar = dados_historico(0)
	if submit = "Alterar" then 	
		response.Redirect("incluir.asp?opt=alt&cod="&historico_encaminhar)
	elseif submit = "Incluir" then 	
		response.Redirect("incluir.asp?opt=inc&cod="&historico_encaminhar)
	elseif submit = "Procurar" then 			
		session("busca1")=request.form("busca1") 
		session("busca2")=request.form("busca2")	
		response.Redirect("index.asp?opt=search&nvg="&chave)
	else		
	
		Set CON7 = Server.CreateObject("ADODB.Connection") 
		ABRIR7 = "DBQ="& CAMINHO_h & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON7.Open ABRIR7	
	
		vetor_historico = split(historico_encaminhar, "$!$")
	 	
		Set RS = Server.CreateObject("ADODB.Recordset")				
		SQL = "SELECT * FROM TB_Historico_Ano Where DA_Ano = "&vetor_historico(0)&" AND NU_Seq ="&vetor_historico(1)&" AND CO_Matricula = "&vetor_historico(2)	
		RS.Open SQL, CON7		
		
		curso_hist = RS("TP_Curso")

		if curso_hist = "EFA" or curso_hist = "EFS" then
			response.Redirect("../../../../relatorios/swd036ef.asp?obr="&historico_encaminhar&"&tipo="&curso_hist)	
		elseif curso_hist = "EM" then	
			response.Redirect("../../../../relatorios/swd036em.asp?obr="&historico_encaminhar)		
		else
			response.Write("Não existe relatório para o tipo de curso "&curso_hist)							
		end if	
	end if	
end if		
%>	