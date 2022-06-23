<%
Function parametros_gerais(tipo_dado)

	parametros_gerais = parametros_gerais_novo(unidade,curso,etapa,turma,disciplina,tipo_dado,outro)
end function

Function parametros_gerais_novo(unidade,curso,etapa,turma,disciplina,tipo_dado,outro)
	if tipo_dado="arred_nota" then
		parametros_gerais_novo="quarto"
	elseif tipo_dado="arred_media" then 
		parametros_gerais_novo="quarto"
	elseif tipo_dado="decimais_nota" then 
		parametros_gerais_novo=1	
	elseif tipo_dado="decimais_media" then 
		parametros_gerais_novo=1		
	elseif tipo_dado="arred_nt_sub_disc" then
		parametros_gerais_novo="mat"
	elseif tipo_dado="arred_nt" then
		parametros_gerais_novo="mat"
	elseif tipo_dado="arred_md_sub_disc" then 
		parametros_gerais_novo="mat"			
	elseif tipo_dado="arred_md" then 
		parametros_gerais_novo="mat_dez"	
	elseif tipo_dado="decimais_nt" then 
		parametros_gerais_novo=0	
	elseif tipo_dado="decimais_md_sub_disc" then 
		parametros_gerais_novo=0			
	elseif tipo_dado="decimais_md" then 
		parametros_gerais_novo=0	
	elseif tipo_dado="exibe_md_decimais" then 
		parametros_gerais_novo="S"	
	elseif tipo_dado="compara_m2" then 
		parametros_gerais_novo="N"		
	elseif tipo_dado="compara_m3" then 
		parametros_gerais_novo="N"			
	elseif tipo_dado="aproxima_m1" then 
		parametros_gerais_novo="N"	
	elseif tipo_dado="aproxima_m2" then 
		parametros_gerais_novo="N"		
	elseif tipo_dado="aproxima_m3" then 
		parametros_gerais_novo="N"		
	elseif tipo_dado="default_lbg" then 
		parametros_gerais_novo=10		
		'parametros_gerais_novo=5	
	elseif tipo_dado="default_lbs" then 
		parametros_gerais_novo=10															
	end if
end function

Function Periodo_Media(tp_modelo,tipo_media,outro)
	if tp_modelo="B" then
		if tipo_media="MA" then
			Periodo_Media=4
		elseif tipo_media="REC" then
			Periodo_Media=5
		elseif tipo_media="MF" then
			Periodo_Media=6		
		end if
	else
		if tipo_media="MA" then
			Periodo_Media=3
		elseif tipo_media="REC" then
			Periodo_Media=4
		elseif tipo_media="MF" then
			Periodo_Media=5					
		end if
	end if
End Function	

Function dados_boletim(tp_modelo,tp_freq,ln_busca,tp_dados,tb_nota)
if tp_modelo="B" then
	if ln_busca=1 then
		if tp_dados="tit" then
			dados_boletim="Disciplinas#!#Aproveitamento#!#M&eacute;dia&nbsp;da&nbsp;Turma#!#Frequencia"
		elseif tp_dados="rowspan" then
			dados_boletim="2#!#1#!#1#!#1"
		elseif tp_dados="colspan" then
			dados_boletim="1#!#13#!#4#!#4"
		elseif tp_dados="pdf_rowspan" then
			dados_boletim="3#!#1#!#1#!#1"	
		elseif tp_dados="mrd_tit" then
			dados_boletim="N&ordm;#!#Nome#!#Aproveitamento#!#M&eacute;dia&nbsp;da&nbsp;Turma#!#Frequencia"
		elseif tp_dados="mrd_rowspan" then
			dados_boletim="2#!#2#!#1#!#1#!#1"
		elseif tp_dados="mrd_colspan" then
			dados_boletim="1#!#1#!#13#!#4#!#4"			
		end if
	elseif ln_busca=2 then
		if tp_dados="tit" then
   dados_boletim="BIM 1#!#BIM 2#!#REC<br>PARAL#!#BIM 1 *#!#BIM 2 *#!#BIM 3#!#BIM 4#!#SOMA<br>BIM#!#M&Eacute;DIA<br>ANUAL#!#PROVA RECUP<br>FINAL#!#M&Eacute;DIA RECUP<br>FINAL#!#PROVA<br>FINAL#!#Result#!#B1#!#B2#!#B3#!#B4#!#B1#!#B2#!#B3#!#B4" 
		elseif tp_dados="pdf_tit" then
   dados_boletim="Bim1#!#Bim2#!#Rec<br>Par#!#Bim1*#!#Bim2*#!#Bim3#!#Bim4#!#Soma<br>Bim#!#Md<br>Anual#!#Pr.R<br>Final#!#Md.R<br>Final#!#Prova<br>Final#!#Res#!#B1#!#B2#!#B3#!#B4#!#B1#!#B2#!#B3#!#B4" 
		elseif tp_dados="periodo_ref" then
			dados_boletim="1#!#2#!#2#!#1#!#2#!#3#!#4#!#0#!#0#!#5#!#0#!#6#!#0#!#1#!#2#!#3#!#4#!#1#!#2#!#3#!#4"
		elseif tp_dados="rowspan" then
			dados_boletim="1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1"
		elseif tp_dados="colspan" then
			dados_boletim="1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1"
		elseif tp_dados="pdf_rowspan" then
			dados_boletim="2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2"					
   		elseif tp_dados="tipo_calc" then
			'BDM - Média buscada diretamente do BD
			'BDR - Recuperação buscada diretamente do BD
			'BDF - Falta buscada diretamente do BD						
	
			dados_boletim="BDM#!#BDM#!#BDR#!#ASTER#!#ASTER#!#BDM#!#BDM#!#SOMA#!#MA#!#RF#!#MF#!#PF#!#RES#!#CMT#!#CMT#!#CMT#!#CMT#!#BDF#!#BDF#!#BDF#!#BDF"
		end if
	end if
else
	if ln_busca=1 then
		if tp_dados="tit" then
			dados_boletim="Disciplinas#!#Aproveitamento#!#M&eacute;dia&nbsp;da&nbsp;Turma#!#Frequencia"
		elseif tp_dados="rowspan" then
			dados_boletim="2#!#1#!#1#!#1"
		elseif tp_dados="colspan" then
			dados_boletim="1#!#13#!#3#!#3"
		elseif tp_dados="pdf_rowspan" then
			dados_boletim="3#!#1#!#1#!#1"				
		elseif tp_dados="mrd_tit" then
			dados_boletim="N&ordm;#!#Nome#!#Aproveitamento#!#M&eacute;dia&nbsp;da&nbsp;Turma#!#Frequencia"
		elseif tp_dados="mrd_rowspan" then
			dados_boletim="2#!#2#!#1#!#1#!#1"
		elseif tp_dados="mrd_colspan" then
			dados_boletim="1#!#1#!#13#!#3#!#3"			
		end if
  	elseif ln_busca=2 then
		if tp_dados="tit" then
			dados_boletim="TRI 1#!#REC<br>PARAL#!#TRI 1*#!#TRI 2#!#REC<br>PARAL#!#TRI 2*#!#TRI 3#!#SOMA<br>TRI#!#M&Eacute;DIA<br>ANUAL#!#PROVA RECUP<br>FINAL#!#M&Eacute;DIA RECUP<br>FINAL#!#PROVA<br>FINAL#!#Result#!#T1#!#T2#!#T3#!#T1#!#T2#!#T3"
		elseif tp_dados="pdf_tit" then
		   dados_boletim="Tri1#!#Rec<br>Par#!#Tri1*#!#Tri2#!#Rec<br>Par#!#Tri2*#!#Tri3#!#Soma<br>Tri#!#Md<br>Anual#!#Pr.R<br>Final#!#Md.R<br>Final#!#Prova<br>Final#!#Res#!#T1#!#T2#!#T3#!#T1#!#T2#!#T3" 
		elseif tp_dados="periodo_ref" then
  			dados_boletim="1#!#1#!#1#!#2#!#2#!#2#!#3#!#0#!#0#!#4#!#0#!#5#!#0#!#1#!#2#!#3#!#1#!#2#!#3"
		elseif tp_dados="rowspan" then
			dados_boletim="1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1"
		elseif tp_dados="colspan" then
			dados_boletim="1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1#!#1"
		elseif tp_dados="pdf_rowspan" then
			dados_boletim="2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2#!#2"			
		elseif tp_dados="tipo_calc" then
			'BDM - Média buscada diretamente do BD
			'BDR - Recuperação buscada diretamente do BD
			'BDF - Falta buscada diretamente do BD						
			
			dados_boletim="BDM#!#BDR#!#ASTER#!#BDM#!#BDR#!#ASTER#!#BDM#!#SOMA#!#MA#!#RF#!#MF#!#PF#!#RES#!#CMT#!#CMT#!#CMT#!#BDF#!#BDF#!#BDF"
		end if
	end if
end if
end function

Function tipo_divisao_ano(curso,co_etapa,tipo_dado)

ano_letivo=session("ano_letivo")

	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"'"
	RS.Open SQL, CON0
	
	if ano_letivo<2011 and tipo_dado="tp_modelo" then
		modelo="B"
	else
		if RS.EOF then
			modelo="B"
			freq="D"
		else
'			curso=curso*1
'			if curso<2 then
'				modelo="B"
'				freq="M"			
'			else
'				modelo="T"
'				freq="M"			
'			end if
			modelo=RS("TP_Modelo")
			freq=RS("IN_Frequencia")
		end if
	end if
	
	if tipo_dado="tp_modelo" then
		tipo_divisao_ano=modelo
	elseif tipo_dado="in_frequencia" then
		tipo_divisao_ano=freq
	end if
end function

Function var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo,tp_dados)
'if tp_modelo="B" then
	if tb_nota="TB_NOTA_A" then
		if tp_dados="BDM" or tp_dados="CMT" then
			var_bd_periodo="VA_Media3"
		elseif tp_dados="BDM1" then
			var_bd_periodo="VA_Media1"		
		elseif tp_dados="BDM2" then
			var_bd_periodo="VA_Media2"	
		elseif tp_dados="BDB" then
			var_bd_periodo="VA_Extra"								
		elseif tp_dados="BDR" then
			var_bd_periodo="VA_Rec"
		elseif tp_dados="RF" then
			var_bd_periodo="VA_RF"
		elseif tp_dados="PF" then
			var_bd_periodo="VA_PF"							
		elseif tp_dados="BDF" then
			IF tp_freq="M" then
				var_bd_periodo="NU_Faltas"
			else
				periodo=periodo*1
				if periodo = 1 then
					var_bd_periodo="NU_Faltas_P1"
				elseif periodo = 2 then
					var_bd_periodo="NU_Faltas_P2"
				elseif periodo = 3 then
					var_bd_periodo="NU_Faltas_P3"
				elseif periodo = 4 then
					var_bd_periodo="NU_Faltas_P4"
				end if
			end if	
		end if
	elseif tb_nota="TB_NOTA_B" then
		if tp_dados="BDM" or tp_dados="CMT" then
			var_bd_periodo="VA_Media3"
		elseif tp_dados="BDM1" then
			var_bd_periodo="VA_Media1"		
		elseif tp_dados="BDM2" then
			var_bd_periodo="VA_Media2"	
		elseif tp_dados="BDB" then
			var_bd_periodo="VA_Extra"				
		elseif tp_dados="BDR" then
			var_bd_periodo="VA_Rec"
		elseif tp_dados="RF" then
			var_bd_periodo="VA_RF"
		elseif tp_dados="PF" then
			var_bd_periodo="VA_PF"				
		elseif tp_dados="BDF" then
			IF tp_freq="M" then
				var_bd_periodo="NU_Faltas"
			else
				periodo=periodo*1
				if periodo = 1 then
					var_bd_periodo="NU_Faltas_P1"
				elseif periodo = 2 then
					var_bd_periodo="NU_Faltas_P2"
				elseif periodo = 3 then
					var_bd_periodo="NU_Faltas_P3"
				elseif periodo = 4 then
					var_bd_periodo="NU_Faltas_P4"
				end if
			end if	
		end if
	elseif tb_nota="TB_NOTA_C" then
		if tp_dados="BDM" or tp_dados="CMT" then
			var_bd_periodo="VA_Media3"
		elseif tp_dados="BDM1" then
			var_bd_periodo="VA_Media1"		
		elseif tp_dados="BDM2" then
			var_bd_periodo="VA_Media2"	
		elseif tp_dados="BDB" then
			var_bd_periodo="VA_Extra"				
		elseif tp_dados="BDR" then
			var_bd_periodo="VA_Rec"
		elseif tp_dados="RF" then
			var_bd_periodo="VA_RF"
		elseif tp_dados="PF" then
			var_bd_periodo="VA_PF"				
		elseif tp_dados="BDF" then
			IF tp_freq="M" then
				var_bd_periodo="NU_Faltas"
			else
				periodo=periodo*1
				if periodo = 1 then
					var_bd_periodo="NU_Faltas_P1"
				elseif periodo = 2 then
					var_bd_periodo="NU_Faltas_P2"
				elseif periodo = 3 then
					var_bd_periodo="NU_Faltas_P3"
				elseif periodo = 4 then
					var_bd_periodo="NU_Faltas_P4"
				end if
			end if		
		end if
	elseif tb_nota="TB_NOTA_D" then
		if tp_dados="BDM" or tp_dados="CMT" then
			var_bd_periodo="VA_Media3"
		elseif tp_dados="BDM1" then
			var_bd_periodo="VA_Media1"		
		elseif tp_dados="BDM2" then
			var_bd_periodo="VA_Media2"	
		elseif tp_dados="BDB" then
			var_bd_periodo="VA_Extra"				
		elseif tp_dados="BDR" then
			var_bd_periodo="VA_Rec"
		elseif tp_dados="RF" then
			var_bd_periodo="VA_RF"
		elseif tp_dados="PF" then
			var_bd_periodo="VA_PF"				
		elseif tp_dados="BDF" then
			IF tp_freq="M" then
				var_bd_periodo="NU_Faltas"
			else
				periodo=periodo*1
				if periodo = 1 then
					var_bd_periodo="NU_Faltas_P1"
				elseif periodo = 2 then
					var_bd_periodo="NU_Faltas_P2"
				elseif periodo = 3 then
					var_bd_periodo="NU_Faltas_P3"
				elseif periodo = 4 then
					var_bd_periodo="NU_Faltas_P4"
				end if
			end if	
		end if
	elseif tb_nota="TB_NOTA_E" then
		if tp_dados="BDM" or tp_dados="CMT" then
			var_bd_periodo="VA_Media3"
		elseif tp_dados="BDM1" then
			var_bd_periodo="VA_Media1"		
		elseif tp_dados="BDM2" then
			var_bd_periodo="VA_Media2"	
		elseif tp_dados="BDB" then
			var_bd_periodo="VA_Extra"				
		elseif tp_dados="BDR" then
			var_bd_periodo="VA_Rec"
		elseif tp_dados="RF" then
			var_bd_periodo="VA_RF"
		elseif tp_dados="PF" then
			var_bd_periodo="VA_PF"				
		elseif tp_dados="BDF" then
			IF tp_freq="M" then
				var_bd_periodo="NU_Faltas"
			else
				periodo=periodo*1
				if periodo = 1 then
					var_bd_periodo="NU_Faltas_P1"
				elseif periodo = 2 then
					var_bd_periodo="NU_Faltas_P2"
				elseif periodo = 3 then
					var_bd_periodo="NU_Faltas_P3"
				elseif periodo = 4 then
					var_bd_periodo="NU_Faltas_P4"
				end if
			end if	
		end if
	end if
'else
'end if
end function

Function verifica_dados_tabela(opcao,subopcao,outro)
	if opcao="A" then
		nom_cols="STATUS#!#DATA#!#HORA"
		wrk_cols="tx_avalia#!#da_ult_acesso#!#ho_ult_acesso"
		nom_bd_cols="TX_Avalia#!#DA_Ult_Acesso#!#HO_ult_Acesso"
		ind_calc="0#!#0#!#0"			
		gera_pdf="sim"		
		if subopcao="action" then
			verifica_dados_tabela="../../../../inc/bda.asp"
		elseif subopcao="notas_a_lancar" then
			verifica_dados_tabela=1
		elseif subopcao="peso_col" then
			'indica em qual coluna incluir o peso no modelo &nbsp;#!#Pesos#!#
			verifica_dados_tabela=""
		elseif subopcao="peso_bd_var" then		
		'nome da variavel na base de dados	
			verifica_dados_tabela=""
		elseif subopcao="peso_wrk_var" then								
		'nome que será usado pelo programa
			verifica_dados_tabela=""
		elseif subopcao="nome_cols" then					
			verifica_dados_tabela="N&ordm;#!#Nome#!#"&nom_cols
		elseif subopcao="bd_var" then					
			verifica_dados_tabela=nom_bd_cols
		elseif subopcao="wrk_var" then				
			verifica_dados_tabela=wrk_cols
		elseif subopcao="calc" then				
			verifica_dados_tabela=ind_calc
		elseif subopcao="bol_av_col" then				
			verifica_dados_tabela="Disciplina#!#"&nom_cols&"#!#Alterado por#!#Data/Hora"
		elseif subopcao="bol_av_span" then			
			verifica_dados_tabela="0#!#0#!#0#!#0#!#0#!#0"
		elseif subopcao="bol_av_wrk_var" then					
			verifica_dados_tabela=wrk_cols
		elseif subopcao="bol_av_bd_var" then					
			verifica_dados_tabela=nom_bd_cols
		elseif subopcao="bol_av_calc" then	
			verifica_dados_tabela="0#!#0#!#0#!#0#!#0#!#0"
		elseif subopcao="bol_av_autoriza_wf" then				
			verifica_dados_tabela="T#!#T#!#T#!#T#!#T#!#T"
			'verifica_dados_tabela="0#!#T#!#T#!#T#!#T#!#P#!#M#!#M#!#M#!#M#!#M"			
		elseif subopcao="bol_av_legenda" then				
			'verifica_dados_tabela="Av1-Avalia&ccedil;&atilde;o 1, Av2-Avalia&ccedil;&atilde;o 2, For-Formativo, Rf-Recupera&ccedil;&atilde;o Final, Pf-Prova Final, M1-M&eacute;dia 1, Ext-Nota Extra, M2-M&eacute;dia 2, Rec - Recupera&ccedil;&atilde;o Semestral e M3-M&eacute;dia 3"	
			verifica_dados_tabela=""							
		end if
	elseif opcao="B" then
		nom_cols="STATUS#!#DATA#!#HORA"
		wrk_cols="tx_avalia#!#da_ult_acesso#!#ho_ult_acesso"
		nom_bd_cols="TX_Avalia#!#DA_Ult_Acesso#!#HO_ult_Acesso"
		ind_calc="0#!#0#!#0"			
		gera_pdf="sim"		
		if subopcao="action" then
			verifica_dados_tabela="../../../../inc/bdb.asp"
		elseif subopcao="notas_a_lancar" then
			verifica_dados_tabela=1
		elseif subopcao="peso_col" then
			'indica em qual coluna incluir o peso no modelo &nbsp;#!#Pesos#!#
			verifica_dados_tabela=""
		elseif subopcao="peso_bd_var" then		
		'nome da variavel na base de dados	
			verifica_dados_tabela=""
		elseif subopcao="peso_wrk_var" then								
		'nome que será usado pelo programa
			verifica_dados_tabela=""
		elseif subopcao="nome_cols" then					
			verifica_dados_tabela="N&ordm;#!#Nome#!#"&nom_cols
		elseif subopcao="bd_var" then					
			verifica_dados_tabela=nom_bd_cols
		elseif subopcao="wrk_var" then				
			verifica_dados_tabela=wrk_cols
		elseif subopcao="calc" then				
			verifica_dados_tabela=ind_calc
		elseif subopcao="bol_av_col" then				
			verifica_dados_tabela="Disciplina#!#"&nom_cols&"#!#Alterado por#!#Data/Hora"
		elseif subopcao="bol_av_span" then			
			verifica_dados_tabela="0#!#0#!#0#!#0#!#0#!#0"
		elseif subopcao="bol_av_wrk_var" then					
			verifica_dados_tabela=wrk_cols
		elseif subopcao="bol_av_bd_var" then					
			verifica_dados_tabela=nom_bd_cols
		elseif subopcao="bol_av_calc" then	
			verifica_dados_tabela="0#!#0#!#0#!#0#!#0#!#0"
		elseif subopcao="bol_av_autoriza_wf" then				
			verifica_dados_tabela="T#!#T#!#T#!#T#!#T#!#T"
			'verifica_dados_tabela="0#!#T#!#T#!#T#!#T#!#P#!#M#!#M#!#M#!#M#!#M"			
		elseif subopcao="bol_av_legenda" then				
			'verifica_dados_tabela="Av1-Avalia&ccedil;&atilde;o 1, Av2-Avalia&ccedil;&atilde;o 2, For-Formativo, Rf-Recupera&ccedil;&atilde;o Final, Pf-Prova Final, M1-M&eacute;dia 1, Ext-Nota Extra, M2-M&eacute;dia 2, Rec - Recupera&ccedil;&atilde;o Semestral e M3-M&eacute;dia 3"	
			verifica_dados_tabela=""							
		end if	
	elseif opcao="C" then	
		nom_cols="STATUS#!#DATA#!#HORA"
		wrk_cols="tx_avalia#!#da_ult_acesso#!#ho_ult_acesso"
		nom_bd_cols="TX_Avalia#!#DA_Ult_Acesso#!#HO_ult_Acesso"
		ind_calc="0#!#0#!#0"			
		gera_pdf="sim"		
		if subopcao="action" then
			verifica_dados_tabela="../../../../inc/bdc.asp"
		elseif subopcao="notas_a_lancar" then
			verifica_dados_tabela=1
		elseif subopcao="peso_col" then
			'indica em qual coluna incluir o peso no modelo &nbsp;#!#Pesos#!#
			verifica_dados_tabela=""
		elseif subopcao="peso_bd_var" then		
		'nome da variavel na base de dados	
			verifica_dados_tabela=""
		elseif subopcao="peso_wrk_var" then								
		'nome que será usado pelo programa
			verifica_dados_tabela=""
		elseif subopcao="nome_cols" then					
			verifica_dados_tabela="N&ordm;#!#Nome#!#"&nom_cols
		elseif subopcao="bd_var" then					
			verifica_dados_tabela=nom_bd_cols
		elseif subopcao="wrk_var" then				
			verifica_dados_tabela=wrk_cols
		elseif subopcao="calc" then				
			verifica_dados_tabela=ind_calc
		elseif subopcao="bol_av_col" then				
			verifica_dados_tabela="Disciplina#!#"&nom_cols&"#!#Alterado por#!#Data/Hora"
		elseif subopcao="bol_av_span" then			
			verifica_dados_tabela="0#!#0#!#0#!#0#!#0#!#0"
		elseif subopcao="bol_av_wrk_var" then					
			verifica_dados_tabela=wrk_cols
		elseif subopcao="bol_av_bd_var" then					
			verifica_dados_tabela=nom_bd_cols
		elseif subopcao="bol_av_calc" then	
			verifica_dados_tabela="0#!#0#!#0#!#0#!#0#!#0"
		elseif subopcao="bol_av_autoriza_wf" then				
			verifica_dados_tabela="T#!#T#!#T#!#T#!#T#!#T"
			'verifica_dados_tabela="0#!#T#!#T#!#T#!#T#!#P#!#M#!#M#!#M#!#M#!#M"			
		elseif subopcao="bol_av_legenda" then				
			'verifica_dados_tabela="Av1-Avalia&ccedil;&atilde;o 1, Av2-Avalia&ccedil;&atilde;o 2, For-Formativo, Rf-Recupera&ccedil;&atilde;o Final, Pf-Prova Final, M1-M&eacute;dia 1, Ext-Nota Extra, M2-M&eacute;dia 2, Rec - Recupera&ccedil;&atilde;o Semestral e M3-M&eacute;dia 3"	
			verifica_dados_tabela=""					
		end if			
	elseif opcao="D" then	
		nom_cols="STATUS#!#DATA#!#HORA"
		wrk_cols="tx_avalia#!#da_ult_acesso#!#ho_ult_acesso"
		nom_bd_cols="TX_Avalia#!#DA_Ult_Acesso#!#HO_ult_Acesso"
		ind_calc="0#!#0#!#0"				
		gera_pdf="sim"		
		if subopcao="action" then
			verifica_dados_tabela="../../../../inc/bdd.asp"
		elseif subopcao="notas_a_lancar" then
			verifica_dados_tabela=1
		elseif subopcao="peso_col" then
			'indica em qual coluna incluir o peso no modelo &nbsp;#!#Pesos#!#
			verifica_dados_tabela=""
		elseif subopcao="peso_bd_var" then		
		'nome da variavel na base de dados	
			verifica_dados_tabela=""
		elseif subopcao="peso_wrk_var" then								
		'nome que será usado pelo programa
			verifica_dados_tabela=""
		elseif subopcao="nome_cols" then					
			verifica_dados_tabela="N&ordm;#!#Nome#!#"&nom_cols
		elseif subopcao="bd_var" then					
			verifica_dados_tabela=nom_bd_cols
		elseif subopcao="wrk_var" then				
			verifica_dados_tabela=wrk_cols
		elseif subopcao="calc" then				
			verifica_dados_tabela=ind_calc
		elseif subopcao="bol_av_col" then				
			verifica_dados_tabela="Disciplina#!#"&nom_cols&"#!#Alterado por#!#Data/Hora"
		elseif subopcao="bol_av_span" then			
			verifica_dados_tabela="0#!#0#!#0#!#0#!#0#!#0"
		elseif subopcao="bol_av_wrk_var" then					
			verifica_dados_tabela=wrk_cols
		elseif subopcao="bol_av_bd_var" then					
			verifica_dados_tabela=nom_bd_cols
		elseif subopcao="bol_av_calc" then	
			verifica_dados_tabela="0#!#0#!#0#!#0#!#0#!#0"
		elseif subopcao="bol_av_autoriza_wf" then				
			verifica_dados_tabela="T#!#T#!#T#!#T#!#T#!#T"
			'verifica_dados_tabela="0#!#T#!#T#!#T#!#T#!#P#!#M#!#M#!#M#!#M#!#M"			
		elseif subopcao="bol_av_legenda" then				
			'verifica_dados_tabela="Av1-Avalia&ccedil;&atilde;o 1, Av2-Avalia&ccedil;&atilde;o 2, For-Formativo, Rf-Recupera&ccedil;&atilde;o Final, Pf-Prova Final, M1-M&eacute;dia 1, Ext-Nota Extra, M2-M&eacute;dia 2, Rec - Recupera&ccedil;&atilde;o Semestral e M3-M&eacute;dia 3"	
			verifica_dados_tabela=""						
		end if	
	elseif opcao="E" then	
		nom_cols="STATUS#!#DATA#!#HORA"
		wrk_cols="tx_avalia#!#da_ult_acesso#!#ho_ult_acesso"
		nom_bd_cols="TX_Avalia#!#DA_Ult_Acesso#!#HO_ult_Acesso"
		ind_calc="0#!#0#!#0"				
		gera_pdf="sim"		
		if subopcao="action" then
			verifica_dados_tabela="../../../../inc/bde.asp"
		elseif subopcao="notas_a_lancar" then
			verifica_dados_tabela=1
		elseif subopcao="peso_col" then
			'indica em qual coluna incluir o peso no modelo &nbsp;#!#Pesos#!#
			verifica_dados_tabela=""
		elseif subopcao="peso_bd_var" then		
		'nome da variavel na base de dados	
			verifica_dados_tabela=""
		elseif subopcao="peso_wrk_var" then								
		'nome que será usado pelo programa
			verifica_dados_tabela=""
		elseif subopcao="nome_cols" then					
			verifica_dados_tabela="N&ordm;#!#Nome#!#"&nom_cols
		elseif subopcao="bd_var" then					
			verifica_dados_tabela=nom_bd_cols
		elseif subopcao="wrk_var" then				
			verifica_dados_tabela=wrk_cols
		elseif subopcao="calc" then				
			verifica_dados_tabela=ind_calc
		elseif subopcao="bol_av_col" then				
			verifica_dados_tabela="Disciplina#!#"&nom_cols&"#!#Alterado por#!#Data/Hora"
		elseif subopcao="bol_av_span" then			
			verifica_dados_tabela="0#!#0#!#0#!#0#!#0#!#0"
		elseif subopcao="bol_av_wrk_var" then					
			verifica_dados_tabela=wrk_cols
		elseif subopcao="bol_av_bd_var" then					
			verifica_dados_tabela=nom_bd_cols
		elseif subopcao="bol_av_calc" then	
			verifica_dados_tabela="0#!#0#!#0#!#0#!#0#!#0"
		elseif subopcao="bol_av_autoriza_wf" then				
			verifica_dados_tabela="T#!#T#!#T#!#T#!#T#!#T"
			'verifica_dados_tabela="0#!#T#!#T#!#T#!#T#!#P#!#M#!#M#!#M#!#M#!#M"			
		elseif subopcao="bol_av_legenda" then				
			'verifica_dados_tabela="Av1-Avalia&ccedil;&atilde;o 1, Av2-Avalia&ccedil;&atilde;o 2, For-Formativo, Rf-Recupera&ccedil;&atilde;o Final, Pf-Prova Final, M1-M&eacute;dia 1, Ext-Nota Extra, M2-M&eacute;dia 2, Rec - Recupera&ccedil;&atilde;o Semestral e M3-M&eacute;dia 3"	
			verifica_dados_tabela=""						
		end if	

	elseif opcao="F" then	
		nom_cols="STATUS#!#DATA#!#HORA"
		wrk_cols="tx_avalia#!#da_ult_acesso#!#ho_ult_acesso"
		nom_bd_cols="TX_Avalia#!#DA_Ult_Acesso#!#HO_ult_Acesso"
		ind_calc="0#!#0#!#0"				
		gera_pdf="sim"		
		if subopcao="action" then
			verifica_dados_tabela="../../../../inc/bdf.asp"
		elseif subopcao="notas_a_lancar" then
			verifica_dados_tabela=1
		elseif subopcao="peso_col" then
			'indica em qual coluna incluir o peso no modelo &nbsp;#!#Pesos#!#
			verifica_dados_tabela=""
		elseif subopcao="peso_bd_var" then		
		'nome da variavel na base de dados	
			verifica_dados_tabela=""
		elseif subopcao="peso_wrk_var" then								
		'nome que será usado pelo programa
			verifica_dados_tabela=""
		elseif subopcao="nome_cols" then					
			verifica_dados_tabela="N&ordm;#!#Nome#!#"&nom_cols
		elseif subopcao="bd_var" then					
			verifica_dados_tabela=nom_bd_cols
		elseif subopcao="wrk_var" then				
			verifica_dados_tabela=wrk_cols
		elseif subopcao="calc" then				
			verifica_dados_tabela=ind_calc
		elseif subopcao="bol_av_col" then				
			verifica_dados_tabela="Disciplina#!#"&nom_cols&"#!#Alterado por#!#Data/Hora"
		elseif subopcao="bol_av_span" then			
			verifica_dados_tabela="0#!#0#!#0#!#0#!#0#!#0"
		elseif subopcao="bol_av_wrk_var" then					
			verifica_dados_tabela=wrk_cols
		elseif subopcao="bol_av_bd_var" then					
			verifica_dados_tabela=nom_bd_cols
		elseif subopcao="bol_av_calc" then	
			verifica_dados_tabela="0#!#0#!#0#!#0#!#0#!#0"
		elseif subopcao="bol_av_autoriza_wf" then				
			verifica_dados_tabela="T#!#T#!#T#!#T#!#T#!#T"
			'verifica_dados_tabela="0#!#T#!#T#!#T#!#T#!#P#!#M#!#M#!#M#!#M#!#M"			
		elseif subopcao="bol_av_legenda" then				
			'verifica_dados_tabela="Av1-Avalia&ccedil;&atilde;o 1, Av2-Avalia&ccedil;&atilde;o 2, For-Formativo, Rf-Recupera&ccedil;&atilde;o Final, Pf-Prova Final, M1-M&eacute;dia 1, Ext-Nota Extra, M2-M&eacute;dia 2, Rec - Recupera&ccedil;&atilde;o Semestral e M3-M&eacute;dia 3"	
			verifica_dados_tabela=""					
		end if			
	End if
End Function	

Function escolheBancoPauta(p_tb_notas,p_subopcao,p_outro)
	if p_subopcao="M" then
		if p_tb_notas="TB_NOTA_A" then
			escolheBancoPauta = "Materia_Lecionada_A"
		elseif p_tb_notas="TB_NOTA_B" then
			escolheBancoPauta = "Materia_Lecionada_B"
		elseif p_tb_notas="TB_NOTA_C" then
			escolheBancoPauta = "Materia_Lecionada_C"
		end if	
	else
		if p_tb_notas="TB_NOTA_A" then
			escolheBancoPauta = "Pauta_A"
		elseif p_tb_notas="TB_NOTA_B" then
			escolheBancoPauta = "Pauta_B"
		elseif p_tb_notas="TB_NOTA_C" then
			escolheBancoPauta = "Pauta_C"
		end if	
	end if
End Function	
	
Function verificaCaminhoBancoPauta(p_tb_notas,p_subopcao,p_outro)
	if p_subopcao="M" then
		if p_tb_notas="Materia_Lecionada_A" then
			verificaCaminhoBancoPauta = CAMINHO_mca
		elseif p_tb_notas="Materia_Lecionada_B" then
			verificaCaminhoBancoPauta = CAMINHO_mcb
		elseif p_tb_notas="Materia_Lecionada_C" then
			verificaCaminhoBancoPauta = CAMINHO_mcc
		end if	
	else
		if p_tb_notas="Pauta_A" then
			verificaCaminhoBancoPauta = CAMINHO_pta
		elseif p_tb_notas="Pauta_B" then
			verificaCaminhoBancoPauta = CAMINHO_ptb
		elseif p_tb_notas="Pauta_C" then
			verificaCaminhoBancoPauta = CAMINHO_ptc
		end if	
	end if
End Function

Function funcao_vetor_medias(tipo_retorno, sigla, outro)
	if tipo_retorno="S" then
		funcao_vetor_medias = "MA#!#RF#!#MF"
	elseif tipo_retorno="N" then
		funcao_vetor_medias = "M&eacute;dia Anual#!#Recupera&ccedil;&atilde;o Final#!#M&eacute;dia Final"	
	elseif tipo_retorno="I" then		
		if sigla = "MA" then
			funcao_vetor_medias = "M&eacute;dia Anual"		
		elseif sigla = "RF" then
			funcao_vetor_medias = "Recupera&ccedil;&atilde;o Final"			
		elseif sigla = "MF" then
			funcao_vetor_medias = "M&eacute;dia Final"			
		end if								
	end if
End Function	

	%>
