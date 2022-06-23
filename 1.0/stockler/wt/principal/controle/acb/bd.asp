<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<%
opt=request.QueryString("opt")

cod_form=Session("cod_form")
nome_form=Session("nome_form")
contrato=Session("contrato")
ativos = Session("ativos")
cancelados = Session("cancelados")		
sem_parcelas = Session("sem_parcelas")
so_bolsistas = Session("so_bolsistas")
dia_de= Session("dia_de")
mes_de= Session("mes_de")
ano_de=Session("ano_de")
dia_ate=Session("dia_ate")
mes_ate=Session("mes_ate")
ano_ate=Session("ano_ate")
bolsa=Session("bolsa")
desconto_de=Session("desconto_de")
desconto_ate=session("desconto_ate")
unidade=Session("unidade")
curso=Session("curso")
etapa=Session("etapa")
turma=session("turma")

nvg = session("chave")
session("chave")=nvg

Session("cod_form")=cod_form
Session("nome_form")=nome_form
Session("contrato")=contrato
Session("ativos")=ativos
Session("cancelados")=cancelados	
Session("sem_parcelas")=sem_parcelas
Session("so_bolsistas")=so_bolsistas
Session("dia_de")=dia_de
Session("mes_de")=mes_de
Session("ano_de")=ano_de
Session("dia_ate")=dia_ate
Session("mes_ate")=mes_ate
Session("ano_ate")=ano_ate
Session("bolsa")=bolsa
Session("desconto_de")=desconto_de
session("desconto_ate") =desconto_ate
Session("unidade")=unidade
Session("curso")=curso
Session("etapa")=etapa
session("turma") =turma

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0	

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1

		Set CON5 = Server.CreateObject("ADODB.Connection") 
		ABRIR5 = "DBQ="& CAMINHO_cr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON5.Open ABRIR5

if opt="c" then

	matricula_contrato = request.form("matricula")
	ano_contrato = request.form("ano_contrato")
	nu_contrato = request.form("contrato") 
	plano_pagto = request.form("plano_pagto") 
	
	plano_pagto=replace(plano_pagto,"%20%"," ")	
	
	Set RSp = Server.CreateObject("ADODB.Recordset")
	SQLp = "SELECT * FROM TB_Plano_Pagamento WHERE NU_Ano_Letivo = "&ano_contrato&" AND CO_PlanoPG = '"&plano_pagto&"'"
	response.Write(SQLp)
	RSp.Open SQLp, CON0
	
	reserva_vaga_prmtro=RSp("VA_Reserva_Vaga")		
	anuidade_prmtro=RSp("VA_Anuidade")	

	dia_contrato_frm = request.form("dia_contrato_frm") 
	mes_contrato_frm = request.form("mes_contrato_frm") 
	ano_contrato_frm = request.form("ano_contrato_frm") 
	dt_contrato = dia_contrato_frm&"/"&mes_contrato_frm&"/"&ano_contrato_frm
	situac_contrato = request.form("situacao") 
	nu_parcelas = request.form("parcelas") 		
	vencimento = request.form("dia_vencimento") 
	mes_inicio_parcelas = request.form("mes_inicio") 
	inicio_parcelas="01/"&mes_inicio_parcelas&"/"&ano_contrato_frm
	resp_fin_prc = request.form("rfp") 
	resp_fin_alt = request.form("rfa") 
	dia_util= request.form("dia_util") 
	proporcional = request.form("proporcional") 

	reserva_vaga = request.form("reserva_vaga") 
	anuidade = request.form("anuidade") 
	
	if situac_contrato = "C" then
		dia_cancelamento = request.form("dia_cancelamento") 
		mes_cancelamento = request.form("mes_cancelamento") 
		ano_cancelamento = request.form("ano_cancelamento") 
		
		dia_cancelamento=dia_cancelamento*1
		if dia_cancelamento = 0 then	
			dia_cancelamento = DatePart("d", now)
		end if	

		mes_cancelamento=mes_cancelamento*1
		if mes_cancelamento = 0 then	
			mes_cancelamento = DatePart("m", now)
		end if	
		
		ano_cancelamento=ano_cancelamento*1
		if ano_cancelamento = 0 then	
			ano_cancelamento = DatePart("yyyy", now)
		end if
		sql_cancela = ", DT_Cancela  = '"&dia_cancelamento&"/"&mes_cancelamento&"/"&ano_cancelamento&"'"
	ELSE
		sql_cancela = ", DT_Cancela  = NULL"
	end if
	
	if left (reserva_vaga,2) = "R$" then
		reserva_vaga=replace(reserva_vaga,"R$","")	
		'reserva_vaga=replace(reserva_vaga,",",".")				
	else
	'	reserva_vaga=replace(reserva_vaga,",",".")							
	end if		
	
	if isnumeric(reserva_vaga) then
		reserva_vaga=reserva_vaga*1	
		reserva_vaga_prmtro=reserva_vaga_prmtro*1
		if reserva_vaga_prmtro <> reserva_vaga then
			reserva_vaga=formatcurrency(reserva_vaga)		
			sql_valor = ", VA_Desconto_Reserva_Vaga = '"&reserva_vaga&"'"
		else
			sql_valor = ", VA_Desconto_Reserva_Vaga = NULL"		
		end if	
	end if	
	
	if left (anuidade,2) = "R$" then
		anuidade=replace(anuidade,"R$","")
		'anuidade=replace(anuidade,",",".")			
	else
		'anuidade=replace(anuidade,",",".")									
	end if	
	
	if isnumeric(anuidade) then
		anuidade_prmtro=anuidade_prmtro*1
		anuidade=anuidade*1
		if anuidade_prmtro <> anuidade then
			anuidade=formatcurrency(anuidade)
			sql_valor = sql_valor&", VA_Desconto_Anuidade = '"&anuidade&"'"
		else
			sql_valor =  sql_valor&", VA_Desconto_Anuidade = NULL"					
		end if		
	end if	

	Set RS = Server.CreateObject("ADODB.Recordset")
	sql_atualiza= "UPDATE TB_Contrato SET ST_Contrato ='"&situac_contrato&"', DT_Contrato='"&dt_contrato&"', CO_Plano_Pagamento='"&plano_pagto&"', NU_Parcelas="&nu_parcelas&",  DI_Prefencia='"&vencimento&"',  IN_Parcela='"&inicio_parcelas&"',  RP_Fina_Principal='"&resp_fin_prc&"', RP_Fina_Alter='"&resp_fin_alt&"',  TP_Vencimento='"&dia_util&"',   TP_Calculo='"&proporcional&"'"&sql_cancela&sql_valor&" where  CO_Matricula ="& matricula_contrato&" AND NU_Ano_Letivo ="& ano_contrato&" AND NU_Contrato = "&nu_contrato
	RS.Open sql_atualiza, CON5
		
	Set RSCONTF = Server.CreateObject("ADODB.Recordset")
	SQLF="UPDATE TB_Alunos SET TP_Resp_Fin = '"&resp_fin_prc&"' WHERE CO_Matricula ="& matricula_contrato
	RSCONTF.Open SQLF, CON1

	if nu_contrato<100000 then
		if nu_contrato<10000 then
			if nu_contrato<1000 then
				if nu_contrato<100 then
					if nu_contrato<10 then
						nu_contrato="00000"&nu_contrato							
					else
						nu_contrato="0000"&nu_contrato					
					end if						
				else
					nu_contrato="000"&nu_contrato					
				end if	
			else
				nu_contrato="00"&nu_contrato					
			end if
		else
			nu_contrato="0"&nu_contrato					
		end if
	end if			
	
	concatena_contrato = ano_contrato&"/"&nu_contrato	
	call GravaLog (nvg,concatena_contrato)		


response.Redirect("alterar_contrato.asp?opt=ok&mc="&matricula_contrato&"&ac="&ano_contrato&"&nc="&nu_contrato)	

elseif opt="b" then
	matricula_contrato = request.form("matricula")
	ano_contrato = request.form("ano_contrato")
	nu_contrato = request.form("contrato") 
	
	Set RSB= Server.CreateObject("ADODB.Recordset")
	SQLB = "SELECT * FROM TB_Contrato_Bolsas where CO_Matricula ="&matricula_contrato&" AND NU_Ano_Letivo="&ano_contrato&" AND NU_Contrato = "&nu_contrato
	RSB.Open SQLB, CON5	
	
	if RSB.EOF then	
		bolsista="Não"	
	else
		bolsista="Sim"			
	end if		

	aplicacao=request.form("aplicacao_bolsa")
	b1_bolsa=request.form("b1_tipo_bolsa")
	b1_tipo_desconto=request.form("b1_tipo_desconto")	
	b1_desconto=request.form("b1_desconto")
	b1_prazo=request.form("b1_prazo")
	b1_ap_bolsa=request.form("b1_aplica_bolsa")
	b1_ob_bolsa=request.form("b1_observacao")
	log1=b1_bolsa
		
	if b1_bolsa<>"nulo" then
		if b1_prazo = "s" then
			dia_inicio1 = request.form("b1_dia_de") 
			mes_inicio1 = request.form("b1_mes_de") 
			ano_inicio1 = request.form("b1_ano_de") 
			
			dia_inicio1=dia_inicio1*1			
			if dia_inicio1 = 0 then	
				dia_inicio1 = DatePart("d", now)
			end if	
	
			mes_inicio1=mes_inicio1*1
			if mes_inicio1 = 0 then	
				mes_inicio1 = DatePart("m", now)
			end if	
			
			ano_inicio1=ano_inicio1*1
			if ano_inicio1 = 0 then	
				ano_inicio1 = DatePart("yyyy", now)
			end if
			
			b1_vl_inic=dia_inicio1&"/"&mes_inicio1&"/"&ano_inicio1
			
			dia_fim1 = request.form("b1_dia_ate") 
			mes_fim1 = request.form("b1_mes_ate") 
			ano_fim1 = request.form("b1_ano_ate") 		
			
			dia_fim1=dia_fim1*1			
			if dia_fim1 = 0 then	
				dia_fim1 = 31
			end if	
	
			mes_fim1=mes_fim1*1
			if mes_fim1 = 0 then	
				mes_fim1 = 12
			end if	
			
			ano_fim1=ano_fim1*1
			if ano_fim1 = 0 then	
				ano_fim1 = ano_contrato
			end if							
			
			b1_vl_fim=dia_fim1&"/"&mes_fim1&"/"&ano_fim1
			b1_pc_inic=NULL
			b1_pc_fim=NULL
			sql_prazo1 = ", VL_Inicio1  = '"&b1_vl_inic&"', VL_Fim1  = '"&b1_vl_fim&"', PC_Inicio1=NULL, PC_Fim1=NULL"				
		else
			b1_vl_inic = NULL
			b1_vl_fim= NULL
			b1_pc_inic=request.form("b1_pi")
			b1_pc_fim=request.form("b1_pf")	
			sql_prazo1 = ", VL_Inicio1  = NULL, VL_Fim1  = NULL, PC_Inicio1="&b1_pc_inic&", PC_Fim1="&b1_pc_fim						
		end if	
		
		if b1_tipo_desconto="nulo" then
			b1_tipo_desconto = NULL	
			sql_tipo_desconto1 = "NULL"
		else
			sql_tipo_desconto1 = ""
		end if	
		
		if isnull(b1_desconto) or b1_desconto="" then
			b1_desconto = NULL	
			sql_desconto1=", VA_Desconto1=NULL"
			log1=log1&"-nulo"			
		else
			sql_desconto1=", VA_Desconto1="&b1_desconto	
			if b1_tipo_desconto="P" then
				nom_desconto=b1_desconto&"%"
			else
				nom_desconto=formatcurrency(b1_desconto)
			end if					
			log1=log1&"-"&nom_desconto
		end if			
		
		if b1_dt_conce="" OR isnull(b1_dt_conce) then					
			dia_concessao1 = DatePart("d", now)
			mes_concessao1 = DatePart("m", now)
			ano_concessao1 = DatePart("yyyy", now)
			dt_concessao1=dia_concessao1&"/"&mes_concessao1&"/"&ano_concessao1
			sql_concessao1 = ", DT_Concessao1  = '"&dt_concessao1&"'"	
		else
			sql_concessao1 = ""		
		end if	

	
		if b1_ap_bolsa="nulo" then
			b1_ap_bolsa = "A"
		end if	
		sql_bolsa1 =", AP_Bolsa1='"&b1_ap_bolsa&"'"
		sql1="CO_Bolsa1='"&b1_bolsa&"'"&sql_desconto1&sql_prazo1&sql_concessao1&sql_bolsa1&", OB_Bolsa1='"&b1_ob_bolsa&"',"			
	else
		b1_bolsa = NULL
		b1_vl_inic = NULL
		b1_vl_fim = NULL
		b1_pc_inic = NULL
		b1_pc_fim = NULL
		dt_concessao1=NULL	
		b1_ap_bolsa = "A"	
		sql_concessao1 = ", DT_Concessao1  = NULL"		
		sql_bolsa1 =", AP_Bolsa1='"&b1_ap_bolsa&"'"
		sql1="CO_Bolsa1=NULL, VA_Desconto1=NULL"&sql_prazo1&sql_concessao1&sql_bolsa1&", OB_Bolsa1='"&b1_ob_bolsa&"',"						
	end if			
				
				

'============			
	b2_bolsa=request.form("b2_tipo_bolsa")
	b2_tipo_desconto=request.form("b2_tipo_desconto")
	b2_desconto=request.form("b2_desconto")
	b2_prazo=request.form("b2_prazo")	
	b2_ap_bolsa=request.form("b2_aplica_bolsa")			
	b2_ob_bolsa=request.form("b2_observacao")	
	log2=b2_bolsa
	
	if b2_bolsa<>"nulo" then
		if b2_prazo = "s" then
			dia_inicio2 = request.form("b2_dia_de") 
			mes_inicio2 = request.form("b2_mes_de") 
			ano_inicio2 = request.form("b2_ano_de") 
			
			dia_inicio2=dia_inicio2*1			
			if dia_inicio2 = 0 then	
				dia_inicio2 = DatePart("d", now)
			end if	
	
			mes_inicio2=mes_inicio2*1
			if mes_inicio2 = 0 then	
				mes_inicio2 = DatePart("m", now)
			end if	
			
			ano_inicio2=ano_inicio2*1
			if ano_inicio2 = 0 then	
				ano_inicio2 = DatePart("yyyy", now)
			end if
			
			b2_vl_inic=dia_inicio2&"/"&mes_inicio2&"/"&ano_inicio2
			
			dia_fim2 = request.form("b2_dia_ate") 
			mes_fim2 = request.form("b2_mes_ate") 
			ano_fim2 = request.form("b2_ano_ate") 		
			
			dia_fim2=dia_fim2*1			
			if dia_fim2 = 0 then	
				dia_fim2 = 31
			end if	
	
			mes_fim2=mes_fim2*1
			if mes_fim2 = 0 then	
				mes_fim2 = 12
			end if	
			
			ano_fim2=ano_fim2*1
			if ano_fim2 = 0 then	
				ano_fim2 = ano_contrato
			end if							
			
			b2_vl_fim=dia_fim2&"/"&mes_fim2&"/"&ano_fim2
			b2_pc_inic=NULL
			b2_pc_fim=NULL
			sql_prazo2 = ", VL_Inicio2  = '"&b2_vl_inic&"', VL_Fim2  = '"&b2_vl_fim&"', PC_Inicio2=NULL, PC_Fim2=NULL"		
		else
			b2_vl_inic = NULL
			b2_vl_fim= NULL
			b2_pc_inic=request.form("b2_pi")
			b2_pc_fim=request.form("b2_pf")			
			sql_prazo2 = ", VL_Inicio2  =NULL, VL_Fim2  = NULL, PC_Inicio2="&b2_pc_inic&", PC_Fim2="&b2_pc_fim				
		end if	
		
		if b2_tipo_desconto="nulo" then
			b2_tipo_desconto = NULL	
			sql_tipo_desconto2 = "NULL"
		else
			sql_tipo_desconto2 = ""
		end if	
		
		if isnull(b2_desconto) or b2_desconto="" then
			b2_desconto = NULL	
			sql_desconto2=", VA_Desconto2=NULL"
			log2=log2&"-nulo"
		else
			sql_desconto2=", VA_Desconto2="&b2_desconto	
			if b2_tipo_desconto="P" then
				nom_desconto=b2_desconto&"%"
			else
				nom_desconto=formatcurrency(b2_desconto)
			end if					
			log2=log2&"-"&nom_desconto			
		end if			
		
		if b2_dt_conce="" OR isnull(b2_dt_conce) then			
			dia_concessao2 = DatePart("d", now)
			mes_concessao2 = DatePart("m", now)
			ano_concessao2 = DatePart("yyyy", now)				
			dt_concessao2=dia_concessao2&"/"&mes_concessao2&"/"&ano_concessao2
			sql_concessao2 = ", DT_Concessao2  = '"&dt_concessao2&"'"		
		else
			sql_concessao2 = ""		
		end if
	
		if b2_ap_bolsa="nulo" then
			b2_ap_bolsa = "A"
		end if	
		sql_bolsa2 =", AP_Bolsa2='"&b2_ap_bolsa&"'"
		sql2="CO_Bolsa2='"&b2_bolsa&"'"&sql_desconto2&sql_prazo2&sql_concessao2&sql_bolsa2&", OB_Bolsa2='"&b2_ob_bolsa&"',"			
		
	else
		b2_bolsa = NULL
		b2_vl_inic = NULL
		b2_vl_fim = NULL
		b2_pc_inic = NULL
		b2_pc_fim = NULL
		dt_concessao2=NULL	
		b2_ap_bolsa = "A"	
		sql_concessao2 = ", DT_Concessao2  = NULL"
		sql_bolsa2 =", AP_Bolsa2='"&b2_ap_bolsa&"'"
		sql2="CO_Bolsa2=NULL, VA_Desconto2=NULL"&sql_prazo2&sql_concessao2&sql_bolsa2&", OB_Bolsa2='"&b2_ob_bolsa&"',"									
	end if		
	
			

'============						
	b3_bolsa=request.form("b3_tipo_bolsa")
	b3_tipo_desconto=request.form("b3_tipo_desconto")		
	b3_desconto=request.form("b3_desconto")
	b3_prazo=request.form("b3_prazo")	
	b3_ap_bolsa=request.form("b3_aplica_bolsa")					
	b3_ob_bolsa=request.form("b3_observacao")		
	log3=b3_bolsa
			
	if b3_bolsa<>"nulo" then
		if b3_prazo = "s" then
			dia_inicio3 = request.form("b3_dia_de") 
			mes_inicio3 = request.form("b3_mes_de") 
			ano_inicio3 = request.form("b3_ano_de") 

			dia_inicio3=dia_inicio3*1			
			if dia_inicio3 = 0 then	
				dia_inicio3 = DatePart("d", now)
			end if	
	
			mes_inicio3=mes_inicio3*1
			if mes_inicio3 = 0 then	
				mes_inicio3 = DatePart("m", now)
			end if	
			
			ano_inicio3=ano_inicio3*1
			if ano_inicio3 = 0 then	
				ano_inicio3 = DatePart("yyyy", now)
			end if
			
			b3_vl_inic=dia_inicio3&"/"&mes_inicio3&"/"&ano_inicio3
			
			dia_fim3 = request.form("b3_dia_ate") 
			mes_fim3 = request.form("b3_mes_ate") 
			ano_fim3 = request.form("b3_ano_ate") 		
			
			dia_fim3=dia_fim3*1			
			if dia_fim3 = 0 then	
				dia_fim3 = 31
			end if	
	
			mes_fim3=mes_fim3*1
			if mes_fim3 = 0 then	
				mes_fim3 = 12
			end if	
			
			ano_fim3=ano_fim3*1
			if ano_fim3 = 0 then	
				ano_fim3 = ano_contrato
			end if							
			
			b3_vl_fim=dia_fim3&"/"&mes_fim3&"/"&ano_fim3
			b3_pc_inic=NULL
			b3_pc_fim=NULL
			sql_prazo3 = ", VL_Inicio3  = '"&b3_vl_inic&"', VL_Fim3  = '"&b3_vl_fim&"', PC_Inicio3=NULL, PC_Fim3=NULL"
		else
			b3_vl_inic = NULL
			b3_vl_fim= NULL
			b3_pc_inic=request.form("b3_pi")
			b3_pc_fim=request.form("b3_pf")	
			sql_prazo3 = ", VL_Inicio3  = NULL, VL_Fim3  = NULL, PC_Inicio3="&b3_pc_inic&", PC_Fim3="&b3_pc_fim			
		end if	

		if b3_tipo_desconto="nulo" then
			b3_tipo_desconto = NULL	
			sql_tipo_desconto3 = "NULL"
		else
			sql_tipo_desconto3 = ""
		end if	
		
		if isnull(b3_desconto) or b3_desconto="" then
			b3_desconto = NULL	
			sql_desconto3=", VA_Desconto3=NULL"
			log3=log3&"-nulo"
		else
			sql_desconto3=", VA_Desconto3="&b3_desconto	
			if b3_tipo_desconto="P" then
				nom_desconto=b3_desconto&"%"
			else
				nom_desconto=formatcurrency(b3_desconto)
			end if					
			log3=log3&"-"&nom_desconto				
		end if			
		
		if b3_dt_conce="" OR isnull(b3_dt_conce) then				
			dia_concessao3 = DatePart("d", now)
			mes_concessao3 = DatePart("m", now)
			ano_concessao3 = DatePart("yyyy", now)
			dt_concessao3=dia_concessao3&"/"&mes_concessao3&"/"&ano_concessao3
			sql_concessao3 = ", DT_Concessao3 = '"&dt_concessao3&"'"	
		else
			sql_concessao3 = ""		
		end if
	
		
		if b3_ap_bolsa="nulo" then
			b3_ap_bolsa = "A"
		end if	
		sql_bolsa3 =", AP_Bolsa3='"&b3_ap_bolsa&"'"
		sql3="CO_Bolsa3='"&b3_bolsa&"'"&sql_desconto3&sql_prazo3&sql_concessao3&sql_bolsa3&", OB_Bolsa3='"&b3_ob_bolsa&"',"			
	else
		b3_bolsa = NULL
		b3_vl_inic = NULL
		b3_vl_fim = NULL
		b3_pc_inic = NULL
		b3_pc_fim = NULL
		dt_concessao3=NULL	
		b3_ap_bolsa = "A"		
		sql_concessao3 = ", DT_Concessao3 = NULL"	
		sql_bolsa3 =", AP_Bolsa3='"&b3_ap_bolsa&"'"
		sql3="CO_Bolsa3=NULL, VA_Desconto3=NULL"&sql_prazo3&sql_concessao3&sql_bolsa3&", OB_Bolsa3='"&b3_ob_bolsa&"',"						
	end if		
	
	usuario_bolsa=session("co_user")
			
	if nu_contrato<100000 then
		if nu_contrato<10000 then
			if nu_contrato<1000 then
				if nu_contrato<100 then
					if nu_contrato<10 then
						nu_contrato="00000"&nu_contrato							
					else
						nu_contrato="0000"&nu_contrato					
					end if						
				else
					nu_contrato="000"&nu_contrato					
				end if	
			else
				nu_contrato="00"&nu_contrato					
			end if
		else
			nu_contrato="0"&nu_contrato					
		end if
	end if			
	
	concatena_contrato = ano_contrato&"/"&nu_contrato	
	call GravaLog (nvg,concatena_contrato&" Bolsas: "&log1&", "&log2&", "&log3)
			
			
	if b1_bolsa<>"nulo" or b2_bolsa<>"nulo" or b3_bolsa<>"nulo" then	
		if aplicacao="B" then
			aplicacao="S"
		end if		

		Set RS = Server.CreateObject("ADODB.Recordset")
		sql_atualiza_bolsa= "UPDATE TB_Contrato SET AP_Bolsa ='"&aplicacao&"' where  CO_Matricula ="& matricula_contrato&" AND NU_Ano_Letivo ="& ano_contrato&" AND NU_Contrato = "&nu_contrato
		RS.Open sql_atualiza_bolsa, CON5
		
		if bolsista="Sim" then				
			Set RS = Server.CreateObject("ADODB.Recordset")
			sql_atualiza_bolsa= "UPDATE TB_Contrato_Bolsas SET "&sql1&sql2&sql3&" CO_Usuario = "&usuario_bolsa&" where  CO_Matricula ="& matricula_contrato&" AND NU_Ano_Letivo ="& ano_contrato&" AND NU_Contrato = "&nu_contrato
			RS.Open sql_atualiza_bolsa, CON5
		else	
			Set RS = server.createobject("adodb.recordset")
			RS.open "TB_Contrato_Bolsas", CON5, 2, 2 'which table do you want open
			RS.addnew	
			
				RS("CO_Matricula") = matricula_contrato
				RS("NU_Ano_Letivo") = ano_contrato
				RS("NU_Contrato") = nu_contrato
				RS("CO_Bolsa1")= b1_bolsa	
				RS("VA_Desconto1")=b1_desconto
				RS("VL_Inicio1")= b1_vl_inic
				RS("VL_Fim1")= b1_vl_fim				
				RS("PC_Inicio1")= b1_pc_inic
				RS("PC_Fim1")= b1_pc_fim									
				RS("AP_Bolsa1")=b1_ap_bolsa
				RS("DT_Concessao1")=dt_concessao1
				RS("OB_Bolsa1")=b1_ob_bolsa
				RS("CO_Bolsa2")= b2_bolsa	
				RS("VA_Desconto2")=b2_desconto
				RS("VL_Inicio2")= b2_vl_inic
				RS("VL_Fim2")= b2_vl_fim				
				RS("PC_Inicio2")= b2_pc_inic
				RS("PC_Fim2")= b2_pc_fim									
				RS("AP_Bolsa2")=b2_ap_bolsa
				RS("DT_Concessao2")=dt_concessao2
				RS("OB_Bolsa2")=b2_ob_bolsa
				RS("CO_Bolsa3")= b3_bolsa	
				RS("VA_Desconto3")=b3_desconto
				RS("VL_Inicio3")= b3_vl_inic
				RS("VL_Fim3")= b3_vl_fim				
				RS("PC_Inicio3")= b3_pc_inic
				RS("PC_Fim3")= b3_pc_fim									
				RS("AP_Bolsa3")=b3_ap_bolsa
				RS("DT_Concessao3")=dt_concessao3
				RS("OB_Bolsa3")=b3_ob_bolsa								
				RS("CO_Usuario")=usuario_bolsa
			RS.update
			RS.close
			set RS=nothing			
		end if
	else
		if bolsista="Sim" then			
			Set RS = Server.CreateObject("ADODB.Recordset")
			sql_atualiza_bolsa= "UPDATE TB_Contrato SET AP_Bolsa = NULL where CO_Matricula ="& matricula_contrato&" AND NU_Ano_Letivo ="& ano_contrato&" AND NU_Contrato = "&nu_contrato
			RS.Open sql_atualiza_bolsa, CON5
			
			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "DELETE * from TB_Contrato_Bolsas where CO_Matricula ="& matricula_contrato&" AND NU_Ano_Letivo ="& ano_contrato&" AND NU_Contrato = "&nu_contrato
			Set RS1 = CON5.Execute(SQL1)			
		end if		
	end if	
response.Redirect("alterar_bolsa.asp?opt=ok&mc="&matricula_contrato&"&ac="&ano_contrato&"&nc="&nu_contrato)	


elseif opt="e" then
	excluir_contrato = request.form("excluir_contratos")
	vetor_contratos = split(excluir_contrato,", ")

	FOR vec = 0 to ubound(vetor_contratos)
	
		vetor_temp = split(vetor_contratos(vec),"-")
		matricula_contrato = vetor_temp(0)
	
		vetor_temp2 = split(vetor_temp(1),"$")
		ano_contrato = vetor_temp2(0)
		nu_contrato = vetor_temp2(1)	 		
	
		if nu_contrato<100000 then
			if nu_contrato<10000 then
				if nu_contrato<1000 then
					if nu_contrato<100 then
						if nu_contrato<10 then
							nu_contrato="00000"&nu_contrato							
						else
							nu_contrato="0000"&nu_contrato					
						end if						
					else
						nu_contrato="000"&nu_contrato					
					end if	
				else
					nu_contrato="00"&nu_contrato					
				end if
			else
				nu_contrato="0"&nu_contrato					
			end if
		end if			
		
		concatena_contrato = ano_contrato&"/"&nu_contrato	

		dia_cancelamento = DatePart("d", now)
		mes_cancelamento = DatePart("m", now)
		ano_cancelamento = DatePart("yyyy", now)

	Set RS = Server.CreateObject("ADODB.Recordset")
	sql_atualiza= "UPDATE TB_Contrato SET ST_Contrato ='C',  DT_Cancela  = '"&dia_cancelamento&"/"&mes_cancelamento&"/"&ano_cancelamento&"' where  CO_Matricula ="& matricula_contrato&" AND NU_Ano_Letivo ="& ano_contrato&" AND NU_Contrato = "&nu_contrato
	RS.Open sql_atualiza, CON5
	
	call GravaLog (nvg,"Cancelou: "&concatena_contrato)
		
	NEXT		
response.Redirect("contratos.asp?opt=ok")

end if
%>