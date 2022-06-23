<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/parametros.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<%
chave=session("nvg")
session("nvg")=chave

opt = request.QueryString("opt")


		Set CON9 = Server.CreateObject("ADODB.Connection") 
		ABRIR9 = "DBQ="& CAMINHO_ax & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON9.Open ABRIR9	


ano_letivo = session("ano_letivo")
co_usr = session("co_user")

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
data_atualiza = dia &"/"& mes &"/"& ano
data_log = data_atualiza
if opt="exc" then
exclui_pedido=request.form("exclui_pedido")
'response.Write(exclui_nf&"<BR>")
exclui_pedido = replace(exclui_pedido,"$!$","/")	
'response.Write(exclui_nf&"<BR>")			
vertorExclui = split(exclui_nf,", ")
for i =0 to ubound(vertorExclui)
'response.Write(vertorExclui(i))
exclui = split(vertorExclui(i),"?")
'response.Write(exclui(0)&"   "&exclui(1)&"<BR>")

nu_pedido = exclui(0)
data_pedido= exclui(1)

				
dados_data=split(data_pedido,"/")
dia=dados_data(0)
mes=dados_data(1)
ano=dados_data(2)

data_pedido_cons=mes&"/"&dia&"/"&ano

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Mov_Estoque NU_Pedido ="& nu_pedido
		RS.Open SQL, CON9
		
		situacao = RS("ST_Pedido")
		
		if situacao="A" then
	
			Set RSD = Server.CreateObject("ADODB.Recordset")
			SQLD = "Select CO_Item, QT_Item From TB_Mov_Estoque_Item where NU_Pedido ="& nu_pedido
			RSD.Open SQLD, CON9	
			
			while not RSD.EOF
				  co_item = RSD("CO_Item")
				  quantidade_item = RSD("QT_Item")
		  
					Set RSC = Server.CreateObject("ADODB.Recordset")
					SQLC = "Select * From TB_Item where CO_Item = "&co_item			
					RSC.Open SQLC, CON9	
					
					if RSC.EOF then
					qtd_atual = 0					
					else
					   qtd_atual = RSC("QT_Atual")	   
					   if isnull(qtd_atual) or qtd_atual="" then
						  qtd_atual = 0
					   end if
					end if	   
				   qtd_atual = qtd_atual*1			   
				   quantidade_item = quantidade_item*1
				   qtd_atual = qtd_atual+quantidade_item
		
				   sql_atualiza="UPDATE TB_Item SET "
				   sql_atualiza=sql_atualiza & "QT_Atual=" & qtd_atual & ","
				   sql_atualiza=sql_atualiza & "DA_Ult_Atua=#" & data_atualiza & "#"
				   sql_atualiza=sql_atualiza & " WHERE CO_Item=" & co_item
				   Set RS3 = CON9.Execute(sql_atualiza)
						   
					
			RSD.MOVENEXT
			WEND			
		
		end if					  

'		Set RSI = Server.CreateObject("ADODB.Recordset")
'		SQLI = "DELETE * from TB_Mov_Estoque_Item WHERE NU_Pedido ="& cod_nf	
'		RSI.Open SQLI, CON9		
		if situaca<>"C" then
		   sql_atualiza="UPDATE TB_Mov_Estoque SET "
		   sql_atualiza=sql_atualiza & "ST_Pedido='C',"
		   sql_atualiza=sql_atualiza & " WHERE NU_Pedido=" & nu_pedido
		   Set RS3 = CON9.Execute(sql_atualiza)	

'		   Set RSD = Server.CreateObject("ADODB.Recordset")
'		   SQLD = "DELETE * from TB_Mov_Estoque WHERE NU_Pedido ="& nu_pedido			
'		   RSD.Open SQLD, CON9
		
			outro= "Cancelar,"&data_log&","&nu_pedido
			call GravaLog (chave,outro)
		end if
next
obr=session("obr")
session("obr")=obr

'response.End()

response.redirect("resumo.asp?or=2&opt=ok1")

elseif opt="inc" then

	qtd_itens = request.form("qtd_itens")
	itens_criados  = request.form("itens_criados")
	cadastra="S"	
	for q=1 to itens_criados
		item_fornecedor_n = request.form("item_fornecedor_"&q)
		quantidade_n = request.form("quantidade_"&q)
			
		if isnull(item_fornecedor_n) or item_fornecedor_n="" then
				
		else
		    if item_fornecedor_n <> "nulo" then
								
				Set RSC = Server.CreateObject("ADODB.Recordset")
				SQLC = "Select * From TB_Item where CO_Item = "&item_fornecedor_n
				RSC.Open SQLC, CON9	
				
				if RSC.EOF then
				
				else
				   qtd_atual = RSC("QT_Atual")	   
				   if isnull(qtd_atual) or qtd_atual="" then
					  qtd_atual = 0
				   end if
				   
				   qtd_atual = qtd_atual*1			   
				   quantidade_n = quantidade_n*1
				   qtd_atual = qtd_atual-quantidade_n
				   
				   if qtd_atual<0 then
				   	  cadastra="N"	
				   end if	   				   
				end if	
			end if						
		end if	
	next		

	if cadastra="S" then
		'Apura novamente o número do pedido, pois alguém pode ter incluído um outro pedido durante o processamento
		Set RS = Server.CreateObject("ADODB.Recordset")
		sql = "Select MAX(NU_Pedido) as Max_Seq From TB_Mov_Estoque"
		RS.Open sql, CON9  
		
		if RS.EOF then
			nu_pedido = 0
		else
			nu_pedido = RS("Max_Seq")	
			if nu_pedido = "" or isnull(nu_pedido) then
				nu_pedido = 0			
			end if			
		end if
		nu_pedido=nu_pedido*1
		nu_pedido = nu_pedido+1
			
		projeto = request.form("projeto")
		'valor = request.form("valor")
		dia_nf = request.form("dia_nf")
		mes_nf = request.form("mes_nf")
		ano_nf = request.form("ano_nf")
		unidade = request.form("unidade")
		curso = request.form("curso")
		etapa = request.form("etapa")	
		turma = request.form("turma")
		obs = request.form("obs")	
		solicitado = request.form("solicitado")		
		if 	etapa = "nulo" then
			etapa = null
		end if	
		
		if 	turma = "nulo" then
			turma = null
		end if			
					
		data_inclui = dia_nf&"/"&mes_nf&"/"&ano_nf
		
	
		
		Set RS = server.createobject("adodb.recordset")
		RS.open "TB_Mov_Estoque", CON9, 2, 2 'which table do you want open
		RS.addnew
		  RS("NU_Pedido") = nu_pedido
		  RS("DA_Pedido") = data_inclui
		  RS("CO_Projeto") = projeto
		  RS("NU_Unidade") = unidade
		  RS("CO_Curso") = curso
		  RS("CO_Etapa") = etapa
		  RS("CO_Turma") = turma	  	  	  
		  RS("TX_Observa") = obs
		  RS("ST_Pedido") = "P"	 	  
		  RS("DA_Atendido") = NULL
		  RS("CO_Usuario") = solicitado
		  RS.update
		  
		set RS=nothing
		
		nu_seq_item = 0
		for n=1 to itens_criados
			item_fornecedor_n = request.form("item_fornecedor_"&n)
			quantidade_n = request.form("quantidade_"&n)
				
			if isnull(item_fornecedor_n) or item_fornecedor_n="" then
					
			else
				if item_fornecedor_n <> "nulo" then
					nu_seq_item = nu_seq_item+1
					
					Set RSI = server.createobject("adodb.recordset")
					RSI.open "TB_Mov_Estoque_Item", CON9, 2, 2 'which table do you want open
					RSI.addnew
					  RSI("NU_Pedido") = nu_pedido
					  RSI("CO_Item") = item_fornecedor_n
					  RSI("QT_Solicitado") = quantidade_n
					  RSI.update
					  
					set RSI=nothing
					
					Set RSC = Server.CreateObject("ADODB.Recordset")
					SQLC = "Select * From TB_Item where CO_Item = "&item_fornecedor_n
					RSC.Open SQLC, CON9	
					
					if RSC.EOF then
					
					else
					   qtd_atual = RSC("QT_Atual")	   
					   if isnull(qtd_atual) or qtd_atual="" then
						  qtd_atual = 0
					   end if
					   
					   qtd_atual = qtd_atual*1			   
					   quantidade_n = quantidade_n*1
					   qtd_atual = qtd_atual-quantidade_n
			
					   sql_atualiza="UPDATE TB_Item SET "
					   sql_atualiza=sql_atualiza & "QT_Atual=" & qtd_atual & ","
					   sql_atualiza=sql_atualiza & "DA_Ult_Atua=#" & data_atualiza & "#"
					   sql_atualiza=sql_atualiza & " WHERE CO_Item=" & item_fornecedor_n
					   Set RS2 = CON9.Execute(sql_atualiza)

	'response.Write(sql_atualiza&"<BR>")					   
					end if	
					   
		'				if Err.number<>0 then
		'				 response.write(Err.Description)
		'				end if
				end if
				
				
			end if	
		next	
		'response.end()
		
		outro= "Incluir nota_fiscal :"&nota_fiscal&","&data_inclui
		call GravaLog (chave,outro)
	
		response.redirect("resumo.asp?opt=ok3")	
	else	
		response.redirect("erro.asp")	
	end if	
elseif opt="alt" then
	nota_fiscal = request.form("nota_fiscal")
	fornecedor = request.form("fornecedor")
	valor = request.form("valor")
	dia_nf = request.form("dia_nf")
	mes_nf = request.form("mes_nf")
	ano_nf = request.form("ano_nf")
	data_nf = mes_nf&"/"&dia_nf&"/"&ano_nf
	data_nfa  = mes_nf&"/"&dia_nf+1&"/"&ano_nf 
	qtd_itens = request.form("qtd_itens")
	itens_criados  = request.form("itens_criados")
	
   sql_atualiza="UPDATE TB_NFiscais_Compra SET "
   sql_atualiza=sql_atualiza & "CO_Fornecedor=" & fornecedor & ","
   sql_atualiza=sql_atualiza & "VA_NotaF='" & valor & "',"   
   sql_atualiza=sql_atualiza & "CO_Usuario_Reg=" & co_usr      
   sql_atualiza=sql_atualiza & " WHERE NU_NotaF='" & nota_fiscal& "' AND DA_NotaF =#" & data_nf & "#"
   Set RS2 = CON9.Execute(sql_atualiza)	
'response.Write(sql_atualiza&"<BR>")	  
'response.End() 
	Set RSN = Server.CreateObject("ADODB.Recordset")
	SQLN = "Select * From TB_NFiscais_Compra_Item where NU_NotaF='" & nota_fiscal& "'"
	
	RSN.Open SQLN, CON9	
	
	While not RSN.EOF
		item_fornecedor_nf = RSN("CO_Item")
		quantidade_nf = RSN("QT_Item")
		
'	response.Write(quantidade_nf&"<BR>")			
	
		Set RSC = Server.CreateObject("ADODB.Recordset")
		SQLC = "Select * From TB_Item where CO_Item = "&item_fornecedor_nf
'	response.Write(SQLC&"<BR>")		
		RSC.Open SQLC, CON9	

		if RSC.EOF then
		
		else
		   qtd_atual = RSC("QT_Atual")	   
		   if isnull(qtd_atual) or qtd_atual="" then
			  qtd_atual = 0
		   end if
		   
		   qtd_atual = qtd_atual*1			   
		   quantidade_nf = quantidade_nf*1
		   qtd_atual = qtd_atual-quantidade_nf

		   sql_atualiza="UPDATE TB_Item SET "
		   sql_atualiza=sql_atualiza & "QT_Atual=" & qtd_atual & ","
		   sql_atualiza=sql_atualiza & "DA_Ult_Atua=#" & data_atualiza & "#"
		   sql_atualiza=sql_atualiza & " WHERE CO_Item=" & item_fornecedor_nf
		   Set RS2 = CON9.Execute(sql_atualiza)
		   
'	response.Write(sql_atualiza&"<BR>")
		
'				if Err.number<>0 then
'				 response.write(Err.Description)
'				end if
		end if					
	
	
	RSN.MOVENEXT
	WEND   
	
	Set RSD = Server.CreateObject("ADODB.Recordset")
	SQLD = "DELETE * from TB_NFiscais_Compra_Item where NU_NotaF='" & nota_fiscal& "'"
	response.Write(SQLD&"<BR>")		
	RSD.Open SQLD, CON9	
	
    nu_seq_item = 0
'	response.Write(itens_criados&"<BR>")
	for n=1 to itens_criados
		item_fornecedor_n = request.form("item_fornecedor_"&n)
		quantidade_n = request.form("quantidade_"&n)
		valor_n = request.form("valor_"&n)
'	response.Write(item_fornecedor_n&"<BR>")	
'	response.Write(quantidade_n&"<BR>")	
		if isnull(item_fornecedor_n) or item_fornecedor_n="" then
				
		else
		    if item_fornecedor_n <> "nulo" then
			
			    nu_seq_item = nu_seq_item+1
				
				Set RSI = server.createobject("adodb.recordset")
				RSI.open "TB_NFiscais_Compra_Item", CON9, 2, 2 'which table do you want open
				RSI.addnew
				  RSI("NU_NotaF") = nota_fiscal
				  RSI("NU_Seq_Item") = nu_seq_item
				  RSI("CO_Item") = item_fornecedor_n
				  RSI("QT_Item") = quantidade_n
				  RSI("VA_Unitario") = valor_n
				  RSI.update
				  
				set RSI=nothing
				
				Set RSC = Server.CreateObject("ADODB.Recordset")
				SQLC = "Select * From TB_Item where CO_Item = "&item_fornecedor_n
				RSC.Open SQLC, CON9	
				
				if RSC.EOF then
				
				else
				   qtd_atual = RSC("QT_Atual")	   
				   if isnull(qtd_atual) or qtd_atual="" then
					  qtd_atual = 0
				   end if
				   
				   qtd_atual = qtd_atual*1			   
				   quantidade_n = quantidade_n*1
				   qtd_atual = qtd_atual+quantidade_n
		
				   sql_atualiza="UPDATE TB_Item SET "
				   sql_atualiza=sql_atualiza & "QT_Atual=" & qtd_atual & ","
				   sql_atualiza=sql_atualiza & "DA_Ult_Atua=#" & data_atualiza & "#"
				   sql_atualiza=sql_atualiza & " WHERE CO_Item=" & item_fornecedor_n
				   Set RS2 = CON9.Execute(sql_atualiza)
				   
				'response.Write("<BR>"&sql_atualiza)
				
	'				if Err.number<>0 then
	'				 response.write(Err.Description)
	'				end if
				end if						
			end if
		end if
	next	
'	response.end()
	outro= "Alterar nota_fiscal :"&nota_fiscal&","&data_nf
		
	call GravaLog (chave,outro)
	session("obr")=obr
	response.redirect("resumo.asp?cod="&cod&"&or=2&opt=ok2")
end if



%>