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

		Set CON99 = Server.CreateObject("ADODB.Connection") 
		ABRIR99 = "DBQ="& ACAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON99.Open ABRIR99	
		
ano_letivo = session("ano_letivo")
co_usr = session("co_user")

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
data_atualiza = dia &"/"& mes &"/"& ano
data_log = data_atualiza
if opt="exc" then
exclui_pedido=request.form("exclui_pedido")

exclui_pedido = replace(exclui_pedido,"$!$","/")	

vertorExclui = split(exclui_pedido,", ")
for i =0 to ubound(vertorExclui)
'		response.Write("=====================================================<BR>")
exclui = split(vertorExclui(i),"?")
nu_pedido = exclui(0)
data_pedido= exclui(1)

				
dados_data=split(data_pedido,"/")
dia=dados_data(0)
mes=dados_data(1)
ano=dados_data(2)

data_pedido_cons=mes&"/"&dia&"/"&ano

		Set RSI = Server.CreateObject("ADODB.Recordset")
		SQLI = "Select * From TB_Mov_Estoque_Item where NU_Pedido ="& nu_pedido	
'		response.Write(SQLI&"<BR>")
		RSI.Open SQLI, CON9	
		
		while not RSI.EOF

		
		  cod_item = RSI("CO_Item")
		  quantidade_item = RSI("QT_Solicitado")
'			 response.Write(cod_item&":"&quantidade_item&"<BR>")
			Set RSC = Server.CreateObject("ADODB.Recordset")
			SQLC = "Select * From TB_Item where CO_Item = "&cod_item		
'			  response.Write(SQLC&"<BR>")			
			RSC.Open SQLC, CON9	
			
			if RSC.EOF then
			
			else
			   qtd_atual = RSC("QT_Atual")	   
			   if isnull(qtd_atual) or qtd_atual="" then
				  qtd_atual = 0
			   end if
'			 response.Write(cod_item&":"&qtd_atual&"<BR>")				   
			   qtd_atual = qtd_atual*1			   
			   quantidade_item = quantidade_item*1
			   qtd_atual = qtd_atual+quantidade_item

			   sql_atualiza="UPDATE TB_Item SET "
			   sql_atualiza=sql_atualiza & "QT_Atual=" & qtd_atual & ", "
			   sql_atualiza=sql_atualiza & "DA_Ult_Atua=#" & data_atualiza & "# "
			   sql_atualiza=sql_atualiza & "WHERE CO_Item=" & cod_item		   
			   Set RS3 = CON9.Execute(sql_atualiza)
'			  response.Write(sql_atualiza&"<BR>")
			  If Err.number<>0 then
				Response.Write "ERRO: " & Err.Description & "<BR>"
				response.End()
			  end if					   
			end if	
		RSI.MOVENEXT
		WEND		
		'response.End()		
	   sql_update="UPDATE TB_Mov_Estoque SET ST_Pedido='C' WHERE NU_Pedido=" & nu_pedido   
	   Set RS9 = CON9.Execute(sql_update)		

		
	outro= "Excluir,"&data_log&","&nu_pedido
	call GravaLog (chave,outro)
'response.Flush()
next
obr=session("obr")
session("obr")=obr

'response.End()

response.redirect("resumo.asp?or=2&opt=ok1")

elseif opt="inc" then
	nu_pedido_form = request.form("nu_pedido")
	qtd_itens = request.form("qtd_itens")
	itens_criados  = request.form("itens_criados")
	cadastra="S"	
	for q=1 to itens_criados
		qtd_atualizada = 0
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
				   qtd_atualizada = qtd_atual-quantidade_n
				   'response.Write(qtd_atualizada&" "&cadastra&" "&item_fornecedor_n&" "&qtd_atual&" - "&quantidade_n&"<BR>")				   
				   if qtd_atualizada<0 then
				   	  cadastra="N"	
					  item_invalido = item_fornecedor_n
					  qtd_digit_invalido = quantidade_n
					  qtd_atual_invalido = qtd_atual					  					  
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
		if 	etapa = "nulo"or etapa = "999990"  then
			etapa = null
		end if	
		
		if 	turma = "nulo" or turma = "999990" then
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
		  If Err.number<>0 then
		    response.Write(nu_pedido)
			Response.Write "ERRO: " & Err.Description & "<BR>"
			response.End()
		  end if		  
		  
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
	'No caso de erro grava na tabela auxiliar
	
'		Set RS = Server.CreateObject("ADODB.Recordset")
'		sql = "Select MAX(NU_Pedido) as Max_Seq From TBA_Mov_Estoque"
'		RS.Open sql, CON9  
'		
'		if RS.EOF then
'			nu_pedido = 0
'		else
'			nu_pedido = RS("Max_Seq")	
'			if nu_pedido = "" or isnull(nu_pedido) then
'				nu_pedido = 0			
'			end if			
'		end if
'		nu_pedido=nu_pedido*1
'		nu_pedido = nu_pedido+1	
			
		chavearray=split(chave,"-")
		A_sistema=chavearray(0)
		A_modulo=chavearray(1)
		A_setor=chavearray(2)
		A_funcao=chavearray(3)
		A_ano = DatePart("yyyy", now)
		A_mes = DatePart("m", now) 
		A_dia = DatePart("d", now) 	
		
		A_dia=A_dia*1
		
		A_mes=A_mes*1
		
		if A_dia<10 then
		A_dia="0"&A_dia
		end if
		if A_mes<10 then
		A_mes="0"&A_mes
		end if			
		
		Set RSC = Server.CreateObject("ADODB.Recordset")
		SQLC = "Select * From TB_Item where CO_Item = "&item_invalido
		RSC.Open SQLC, CON9				
		
		nome_item_invalido=RSC("NO_Item")
	
		Set RS = server.createobject("adodb.recordset")
		RS.open "TBA_Msg_Erro", CON99, 2, 2 'which table do you want open
		RS.addnew
		  RS("CO_Sistema") = A_sistema
		  RS("CO_Modulo") = A_modulo
		  RS("CO_Setor") = A_setor
		  RS("CO_Funcao") = A_funcao
		  RS("CO_Usuario") = session("co_user")
		  RS("NO_Tabela") = "TBA_Mov_Estoque"	
		  RS("CO_Chave_Erro") = nu_pedido_form			  		    
		  RS("DSC_Msg_Erro") = "N&atilde;o h&aacute; em estoque a quantidade solicitada ("&qtd_digit_invalido&") do item "&nome_item_invalido&". A quantidade dispon&iacute;vel &eacute; "&qtd_atual_invalido
		  RS("DA_Ult_Atua") = A_ano&A_mes&A_dia	  	  	  
		  RS.update
		  
		set RS=nothing		
			
		

			
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
		if 	etapa = "nulo"or etapa = "999990"  then
			etapa = null
		end if	
		
		if 	turma = "nulo" or turma = "999990" then
			turma = null
		end if		
					
		data_inclui = dia_nf&"/"&mes_nf&"/"&ano_nf
		
	
		
		Set RS = server.createobject("adodb.recordset")
		RS.open "TBA_Mov_Estoque", CON99, 2, 2 'which table do you want open
		RS.addnew
		  RS("NU_Pedido") = nu_pedido_form
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
		  RS("DA_Ult_Atua") = A_ano&A_mes&A_dia	  
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
					RSI.open "TBA_Mov_Estoque_Item", CON99, 2, 2 'which table do you want open
					RSI.addnew
					  RSI("NU_Pedido") = nu_pedido_form
					  RSI("CO_Item") = item_fornecedor_n
					  RSI("QT_Solicitado") = quantidade_n
					  RSI("DA_Ult_Atua") = A_ano&A_mes&A_dia	  					  
					  RSI.update
					  If Err.number<>0 then
						Response.Write "A descrição fornecida é: " & Err.Description & "<BR>"
					  end if
					set RSI=nothing					
				end if
				
				
			end if	
		next		
		data_inclui=replace(data_inclui,"/","$!$")
		'response.end()
		response.redirect("confirma.asp?opt=erri&cod="&nu_pedido_form&"?"&data_inclui)	
	end if	
elseif opt="alt" then

nu_pedido = request.form("nu_pedido")
situacao_pedido = request.form("situacao_pedido")
'RESPONSE.Write(">"&situacao_pedido)


    IF situacao_pedido<>"C" THEN
		Set RSC = Server.CreateObject("ADODB.Recordset")
		SQLC = "Select * From TB_Mov_Estoque_Item where NU_Pedido = "&nu_pedido
		RSC.Open SQLC, CON9	
		
		while not RSC.EOF
			item_restabelecido = RSC("CO_Item")
			qtd_restabelecido = RSC("QT_Solicitado")	
		'response.Write(item_restabelecido&"'"&qtd_restabelecido&"'<BR>")				
			Set RSI = Server.CreateObject("ADODB.Recordset")
			SQLI = "Select * From TB_Item where CO_Item = "&item_restabelecido
			RSI.Open SQLC, CON9	
			
			if RSI.EOF then
			
			else
			   qtd_estoque = RSI("QT_Atual")	
		'response.Write("'"&qtd_estoque&"'<BR>")			      
			   if isnull(qtd_estoque) or qtd_estoque="" then
				  qtd_estoque = 0
			   end if
			   
			   qtd_estoque = qtd_estoque*1			   
			   qtd_restabelecido = qtd_restabelecido*1
			   qtd_estoque = qtd_estoque+qtd_restabelecido
	
			   sql_atualiza="UPDATE TB_Item SET "
			   sql_atualiza=sql_atualiza & "QT_Atual=" & qtd_estoque & ","
			   sql_atualiza=sql_atualiza & "DA_Ult_Atua=#" & data_atualiza & "#"
			   sql_atualiza=sql_atualiza & " WHERE CO_Item=" & item_restabelecido
			   Set RS2 = CON9.Execute(sql_atualiza)
	
			end if			
		
		RSC.MOVENEXT
		WEND
    END IF

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
				   qtd_atualizada = qtd_atual-quantidade_n
				   
				   if qtd_atualizada<0 then
				   	  cadastra="N"	
					  item_invalido = item_fornecedor_n
					  qtd_digit_invalido = quantidade_n
					  qtd_atual_invalido = qtd_atual					  					  
				   end if	   				   
				end if	
			end if						
		end if	
	next		
qtd_atual = 0
	if cadastra="S" then
	
		Set RSD = Server.CreateObject("ADODB.Recordset")
		CONEXAOD = "DELETE * from TB_Mov_Estoque_Item WHERE NU_Pedido = "&nu_pedido
		Set RSD = CON9.Execute(CONEXAOD)	
	
	
	
	
			
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
		
		Set RSD = Server.CreateObject("ADODB.Recordset")
		CONEXAOD = "DELETE * from TB_Mov_Estoque WHERE NU_Pedido = "&nu_pedido
		Set RSD = CON9.Execute(CONEXAOD)	
		
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
'response.Write(item_fornecedor_n&" "&quantidade_n&"<BR>")				
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
'	response.Write(SQLC&"<BR>")						
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

'	response.Write(sql_atualiza&"<BR>")					   
					end if	
					   
						if Err.number<>0 then
						 response.write(Err.Description)
						end if
				end if
				
				
			end if	
		next	
'		response.end()
		
		outro= "Alterar nota_fiscal :"&nota_fiscal&","&data_inclui
		call GravaLog (chave,outro)
	
		response.redirect("resumo.asp?opt=ok2")	
	else
	'No caso de erro grava na tabela auxiliar
	
	
		'Retorna o estoque para a condição anterior
		Set RSC = Server.CreateObject("ADODB.Recordset")
		SQLC = "Select * From TB_Mov_Estoque_Item where NU_Pedido = "&nu_pedido
		RSC.Open SQLC, CON9	
		
		while not RSC.EOF
			item_restabelecido = RSC("CO_Item")
			qtd_restabelecido = RSC("QT_Solicitado")	
			
			Set RSC = Server.CreateObject("ADODB.Recordset")
			SQLC = "Select * From TB_Item where CO_Item = "&item_restabelecido
			RSC.Open SQLC, CON9	
			
			if RSC.EOF then
			
			else
			   qtd_estoque = RSC("QT_Atual")	   
			   if isnull(qtd_estoque) or qtd_estoque="" then
				  qtd_estoque = 0
			   end if
			   
			   qtd_estoque = qtd_estoque*1			   
			   qtd_restabelecido = qtd_restabelecido*1
			   qtd_estoque = qtd_estoque-qtd_restabelecido
	
			   sql_atualiza="UPDATE TB_Item SET "
			   sql_atualiza=sql_atualiza & "QT_Atual=" & qtd_estoque & ","
			   sql_atualiza=sql_atualiza & "DA_Ult_Atua=#" & data_atualiza & "#"
			   sql_atualiza=sql_atualiza & " WHERE CO_Item=" & item_restabelecido
			   Set RS2 = CON9.Execute(sql_atualiza)
	
			end if			
		
		RSC.MOVENEXT
		WEND
		
		Set RSC = Server.CreateObject("ADODB.Recordset")
		SQLC = "Select * From TB_Item where CO_Item = "&item_invalido
		RSC.Open SQLC, CON9				
		
		nome_item_invalido=RSC("NO_Item")	
	
	
		chavearray=split(chave,"-")
		A_sistema=chavearray(0)
		A_modulo=chavearray(1)
		A_setor=chavearray(2)
		A_funcao=chavearray(3)
		A_ano = DatePart("yyyy", now)
		A_mes = DatePart("m", now) 
		A_dia = DatePart("d", now) 	
		
		A_dia=A_dia*1
		
		A_mes=A_mes*1
		
		if A_dia<10 then
		A_dia="0"&A_dia
		end if
		if A_mes<10 then
		A_mes="0"&A_mes
		end if			
		Set RS = server.createobject("adodb.recordset")
		RS.open "TBA_Msg_Erro", CON99, 2, 2 'which table do you want open
		RS.addnew
		  RS("CO_Sistema") = A_sistema
		  RS("CO_Modulo") = A_modulo
		  RS("CO_Setor") = A_setor
		  RS("CO_Funcao") = A_funcao
		  RS("CO_Usuario") = session("co_user")
		  RS("NO_Tabela") = "TBA_Mov_Estoque"	
		  RS("CO_Chave_Erro") = nu_pedido			  
		  RS("DSC_Msg_Erro") = "N&atilde;o h&aacute; em estoque a quantidade solicitada ("&qtd_digit_invalido&") do item "&nome_item_invalido&". A quantidade dispon&iacute;vel &eacute; "&qtd_atual_invalido
		  RS("DA_Ult_Atua") = A_ano&A_mes&A_dia	  	  	  
		  RS.update
		  
		set RS=nothing		

	  		    

				
			
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
		RS.open "TBA_Mov_Estoque", CON99, 2, 2 'which table do you want open
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
					RSI.open "TBA_Mov_Estoque_Item", CON99, 2, 2 'which table do you want open
					RSI.addnew
					  RSI("NU_Pedido") = nu_pedido
					  RSI("CO_Item") = item_fornecedor_n
					  RSI("QT_Solicitado") = quantidade_n
					  RSI.update
					  
					set RSI=nothing					
				end if
				
				
			end if	
		next		
		data_inclui=replace(data_inclui,"/","$!$")
		
		response.redirect("confirma.asp?opt=erra&cod="&nu_pedido&"?"&data_inclui)	
	end if	




	outro= "Alterar nota_fiscal :"&nota_fiscal&","&data_nf
		
	call GravaLog (chave,outro)
	session("obr")=obr

end if



%>