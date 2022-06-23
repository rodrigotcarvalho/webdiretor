<%Function verifica_dados_tabela(CAMINHOn,opcao,outro)

pasta=split(CAMINHOn, "\")
escola=pasta(4)

'============================================================================================================================================


if escola="boechat" or escola="testeboechat" then

	if opcao="A" then		
			tb="TB_NOTA_A"
			ln_pesos_cols=""
			ln_pesos_vars=""		
			nm_pesos_vars=""				
			ln_nom_cols="#!#"
			nm_vars="#!#"
			nm_bd="#!#"	
			vars_calc="#!#"	
			action=""
			notas_a_lancar=0
			gera_pdf="nao"		
			
'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols=""
			ln_bol_av_span=""
			nm_bol_av_vars=""
			ln_bol_av_vars=""	

			vars_bol_av_calc=""
			legenda_bol_av=""
			exibe_apr_pr_bol_av=""				

	elseif opcao="B" then	
			tb="TB_NOTA_B"
			ln_pesos_cols=""
			ln_pesos_vars=""		
			nm_pesos_vars=""				
			ln_nom_cols="#!#"
			nm_vars="#!#"
			nm_bd="#!#"		
			vars_calc="#!#"	
			action=""
			notas_a_lancar=0
			gera_pdf="nao"	
			
'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols=""
			ln_bol_av_span=""
			nm_bol_av_vars=""
			ln_bol_av_vars=""	

			vars_bol_av_calc=""
			legenda_bol_av=""
			exibe_apr_pr_bol_av=""									

	elseif opcao="C" then
			tb="TB_NOTA_C"
'Planilha de Notas			
			ln_pesos_cols="&nbsp;#!#Pesos#!#&nbsp;#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#VA_MAX_AV1#!#VA_MAX_AV2#!#VA_MAX_AV3#!#VA_MAX_AV4#!#VA_MAX_AV5#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"			
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#p_av1#!#p_av2#!#p_av3#!#p_av4#!#p_av5#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"	
			ln_nom_cols="N&ordm;#!#Nome#!#FAL#!#AV1#!#AV2#!#AV3#!#AV4#!#AV5#!#NPA#!#NAV#!#PR#!#M1#!#BON#!#M2#!#REC#!#M3"
			nm_vars="faltas#!#av1#!#av2#!#av3#!#av4#!#av5#!#npa#!#media_teste#!#pr#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd="NU_Faltas#!#VA_AV1#!#VA_AV2#!#VA_AV3#!#VA_AV4#!#VA_AV5#!#CALCULADO#!#VA_Avaliacao#!#VA_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#CALC1#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			action="../../../../inc/bdc.asp"	
			notas_a_lancar=9	
			gera_pdf="sim"



'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols="Disciplina#!#FAL#!#AV1#!#AV2#!#AV3#!#AV4#!#AV5#!#NPA#!#NAV#!#PR#!#M1#!#Bon#!#M2#!#REC#!#M3#!#Alterado por#!#Data/Hora#!!#N#!#P#!#N#!#P#!#N#!#P#!#N#!#P#!#N#!#P"
			ln_bol_av_span="ROWSPAN#!#ROWSPAN#!#COLSPAN#!#COLSPAN#!#COLSPAN#!#COLSPAN#!#COLSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN"
			nm_bol_av_vars="faltas#!#av1#!#p_av1#!#av2#!#p_av2#!#av3#!#p_av3#!#av4#!#p_av4#!#av5#!#p_av5#!#npa#!#nav#!#pr#!#media1#!#bon#!#media2#!#rec#!#media3"
			ln_bol_av_vars="NU_Faltas#!#VA_AV1#!#VA_MAX_AV1#!#VA_AV2#!#VA_MAX_AV2#!#VA_AV3#!#VA_MAX_AV3#!#VA_AV4#!#VA_MAX_AV4#!#VA_AV5#!#VA_MAX_AV5#!#CALCULADO#!#VA_Avaliacao#!#VA_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"	

			vars_bol_av_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#CALC1#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#P#!#M#!#M#!#M#!#M#!#M"
			legenda_bol_av="FAL - Faltas, AV - Avalia&ccedil;&atilde;o, N - Nota, P - Peso, NPA - Nota Parcial de Avalia��es, NAV - Nota das Avalia��es, Pr - Prova, M - M�dias, Bon - B�nus, Rec - Recupera��o"
			
	else
			tb=""
			ln_pesos_cols=""
			ln_pesos_vars=""		
			nm_pesos_vars=""				
			ln_nom_cols="#!#"
			nm_vars="#!#"
			nm_bd="#!#"		
			vars_calc="#!#"	
			action=""
			notas_a_lancar=0
			gera_pdf="nao"
			
'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols=""
			ln_bol_av_span=""
			nm_bol_av_vars=""
			ln_bol_av_vars=""	

			vars_bol_av_calc=""
			legenda_bol_av=""
			exibe_apr_pr_bol_av=""				
	end if	
	
'=========================================================================================================================================	
	
elseif escola="bretanha" or escola="testebretanha" then

	if opcao="A" then		
			tb="TB_NOTA_A"
			ln_pesos_cols="&nbsp;#!#Pesos#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PE_Teste#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PE_Prova#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"			
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#pt#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#pp#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
			ln_nom_cols="N&ordm;#!#Nome#!#F#!#T1#!#T2#!#T3#!#MT#!#P1#!#P2#!#P3#!#MP#!#M1#!#Bon#!#M2#!#Rec#!#M3"
			nm_vars="faltas#!#t1#!#t2#!#t3#!#media_teste#!#p1#!#p2#!#p3#!#media_prova#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd="NU_Faltas#!#VA_Teste1#!#VA_Teste2#!#VA_Teste3#!#MD_Teste#!#VA_Prova1#!#VA_Prova2#!#VA_Prova3#!#MD_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			action="../../../../inc/bda.asp"
			notas_a_lancar=9
			gera_pdf="sim"			

	elseif opcao="B" then	
			tb="TB_NOTA_B"
			ln_pesos_cols="&nbsp;#!#Pesos#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PESO#!#PESO#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PE_Teste#!#PE_Prova1#!#PE_Prova2#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#pt#!#pp1#!#pp2#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
			ln_nom_cols="N&ordm;#!#Nome#!#F#!#T1#!#T2#!#MT#!#P1#!#P2#!#M1#!#Bon#!#M2#!#Rec#!#M3"
			nm_vars="faltas#!#t1#!#t2#!#media_teste#!#p1#!#p2#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd="NU_Faltas#!#VA_Teste1#!#VA_Teste2#!#MD_Teste#!#VA_Prova1#!#VA_Prova2#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			action="../../../../inc/bdb.asp"
			notas_a_lancar=7
			gera_pdf="sim"							

	elseif opcao="C" then
			tb="TB_NOTA_C"
			ln_pesos_cols="&nbsp;#!#Pesos#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PESO1#!#&nbsp;#!#&nbsp;#!#PESO2#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PE_Teste#!#&nbsp;#!#&nbsp;#!#PE_Prova#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"			
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#pt#!#&nbsp;#!#&nbsp;#!#pp#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"			
			ln_nom_cols="N&ordm;#!#Nome#!#F#!#T1#!#T2#!#MT#!#P1#!#P2#!#MP#!#M1#!#Bon#!#M2#!#Rec#!#M3"
			nm_vars="faltas#!#t1#!#t2#!#media_teste#!#p1#!#p2#!#media_prova#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd="NU_Faltas#!#VA_Teste1#!#VA_Teste2#!#MD_Teste#!#VA_Prova1#!#VA_Prova2#!#MD_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			action="../../../../inc/bdc.asp"
			notas_a_lancar=7
			gera_pdf="sim"	

	else
			tb=""
			ln_pesos_cols=""
			ln_pesos_vars=""		
			nm_pesos_vars=""				
			ln_nom_cols="#!#"
			nm_vars="#!#"
			nm_bd="#!#"		
			vars_calc="#!#"	
			action=""
			notas_a_lancar=0
			gera_pdf="nao"
	end if								

'============================================================================================================================================	

elseif escola="jbarro" or escola="testejbarro" then

	if opcao="A" then		
			tb="TB_NOTA_A"
			ln_pesos_cols=""
			ln_pesos_vars=""			
			nm_pesos_vars=""
			ln_nom_cols="N&ordm;#!#Nome#!#F#!#Tr1#!#Tr2#!#Tr3#!#MTR#!#Te1#!#Te2#!#Te3#!#MTE#!#Pr#!#Sim#!#M1#!#Bon#!#M2#!#Rec#!#M3"
			nm_vars="faltas#!#tr1#!#tr2#!#tr3#!#media_teste#!#te1#!#te2#!#te3#!#media_prova#!#pr#!#sim#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd="NU_Faltas#!#VA_TR1#!#VA_TR2#!#VA_TR3#!#VA_MTR#!#VA_TE1#!#VA_TE2#!#VA_TE3#!#VA_MTE#!#VA_Prova#!#VA_Simul#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			action="../../../../inc/bda.asp"
			notas_a_lancar=11
			gera_pdf="sim"		
			
'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols="Disciplina#!#F#!#Tr1#!#Tr2#!#Tr3#!#MTR#!#Te1#!#Te2#!#Te3#!#MTE#!#Pr#!#Sim#!#M1#!#EXT#!#M2#!#REC#!#M3#!#Alterado por#!#Data/Hora"
			ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			nm_bol_av_vars=nm_vars
			ln_bol_av_vars=nm_bd	

			vars_bol_av_calc=vars_calc
			exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#P#!#P#!#M#!#M#!#M#!#M#!#M"
			legenda_bol_av="F - Faltas, Tr-Trabalho, MTR-M&eacute;dia dos Trabalhos,  Te-Testes, MTE-M&eacute;dia dos Testes, Pr� Prova, Sim�Simulado, M-M&eacute;dia, Bon-B&ocirc;nus e Rec-Recupera&ccedil;&atilde;o"				

	elseif opcao="B" then	
			tb="TB_NOTA_B"
			ln_pesos_cols=""
			ln_pesos_vars=""			
			nm_pesos_vars=""
			ln_nom_cols="N&ordm;#!#Nome#!#F#!#Tr1#!#Tr2#!#Tr3#!#MTR#!#Te1#!#Te2#!#Te3#!#MTE#!#Pr#!#M1#!#Bon#!#M2#!#Rec#!#M3"
			nm_vars="faltas#!#tr1#!#tr2#!#tr3#!#media_teste#!#te1#!#te2#!#te3#!#media_prova#!#pr#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd="NU_Faltas#!#VA_TR1#!#VA_TR2#!#VA_TR3#!#VA_MTR#!#VA_TE1#!#VA_TE2#!#VA_TE3#!#VA_MTE#!#VA_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			action="../../../../inc/bdb.asp"
			notas_a_lancar=10
			gera_pdf="sim"		
			
'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols="Disciplina#!#F#!#Tr1#!#Tr2#!#Tr3#!#MTR#!#Te1#!#Te2#!#Te3#!#MTE#!#Pr#!#M1#!#EXT#!#M2#!#REC#!#M3#!#Alterado por#!#Data/Hora"
			ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			nm_bol_av_vars=nm_vars
			ln_bol_av_vars=nm_bd	

			vars_bol_av_calc=vars_calc
			exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#P#!#M#!#M#!#M#!#M#!#M"
			legenda_bol_av="F - Faltas, Tr-Trabalho, MTR-M&eacute;dia dos Trabalhos,  Te-Testes, MTE-M&eacute;dia dos Testes, Pr� Prova, M-M&eacute;dia, Bon-B&ocirc;nus e Rec-Recupera&ccedil;&atilde;o"							

	elseif opcao="C" then
			tb=""
			ln_pesos_cols=""
			ln_pesos_vars=""		
			nm_pesos_vars=""				
			ln_nom_cols="#!#"
			nm_vars="#!#"
			nm_bd="#!#"		
			vars_calc="#!#"	
			action=""
			notas_a_lancar=0
			gera_pdf="nao"

'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols=""
			ln_bol_av_span=""
			nm_bol_av_vars=""
			ln_bol_av_vars=""	

			vars_bol_av_calc=""
			legenda_bol_av=""
			exibe_apr_pr_bol_av=""

	else
			tb=""
			ln_pesos_cols=""
			ln_pesos_vars=""		
			nm_pesos_vars=""				
			ln_nom_cols="#!#"
			nm_vars="#!#"
			nm_bd="#!#"		
			vars_calc="#!#"	
			action=""
			notas_a_lancar=0
			gera_pdf="nao"
			
'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols=""
			ln_bol_av_span=""
			nm_bol_av_vars=""
			ln_bol_av_vars=""	

			vars_bol_av_calc=""
			legenda_bol_av=""
			exibe_apr_pr_bol_av=""				
	end if								

'============================================================================================================================================	
	
elseif escola="mraythe" or escola="testemraythe" then

	if opcao="A" then	

		tb="TB_NOTA_A"
		gera_pdf="sim"	
		action="../../../../inc/bda.asp"
						
'No 4� per�odo somente PR � lan�ado======================================================================
		if isnumeric(outro) then
			outro=outro*1
		elseif outro="<2011-4" then	
			outro=4
		else
			outro=1			
		end if	

		if outro=4 then
			ln_pesos_cols=""
			ln_pesos_vars=""			
			nm_pesos_vars=""			
			ln_nom_cols="N&ordm;#!#Nome#!#PR#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#M1#!#EXT#!#M2#!#&nbsp;#!#M3"
			nm_vars="p1#!#nav#!#nav#!#nav#!#nav#!#nav#!#nav#!#media1#!#bon#!#media2#!#nav#!#media3"
			nm_bd="VA_PR#!#VA_TR#!#VA_S1#!#VA_S2#!#VA_SS#!#NU_Faltas#!#VA_ATV#!#VA_Media1#!#VA_Extra#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			notas_a_lancar=2
			gera_pdf="sim"			

'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols="Disciplina#!#PR#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#M1#!#EXT#!#M2#!#&nbsp;#!#M3#!#Alterado por#!#Data/Hora"
			ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			nm_bol_av_vars=nm_vars
			ln_bol_av_vars=nm_bd	

			vars_bol_av_calc=vars_calc
			exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#P#!#P#!#M#!#M#!#M#!#M#!#M"
			legenda_bol_av="F - Faltas, Tr-Trabalho, S�Simulado, Ss-Soma dos Simulados, Pr� Prova, Atv-Atividade, M-M&eacute;dia, Ext-Nota Extra e Rec-Recupera&ccedil;&atilde;o"
		else
			ln_pesos_cols=""
			ln_pesos_vars=""			
			nm_pesos_vars=""			
			ln_nom_cols="N&ordm;#!#Nome#!#F#!#TR#!#S1#!#S2#!#SS#!#PR#!#ATV#!#M1#!#EXT#!#M2#!#REC#!#M3"
			nm_vars="faltas#!#t1#!#t2#!#t3#!#media_teste#!#p1#!#p2#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd="NU_Faltas#!#VA_TR#!#VA_S1#!#VA_S2#!#VA_SS#!#VA_PR#!#VA_ATV#!#VA_Media1#!#VA_Extra#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			notas_a_lancar=8

'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols="Disciplina#!#F#!#TR#!#S1#!#S2#!#SS#!#PR#!#ATV#!#M1#!#EXT#!#M2#!#REC#!#M3#!#Alterado por#!#Data/Hora"
			ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			nm_bol_av_vars=nm_vars
			ln_bol_av_vars=nm_bd	

			vars_bol_av_calc=vars_calc
			exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#P#!#P#!#M#!#M#!#M#!#M#!#M"
			legenda_bol_av="F - Faltas, Tr-Trabalho, S�Simulado, Ss-Soma dos Simulados, Pr� Prova, Atv-Atividade, M-M&eacute;dia, Ext-Nota Extra e Rec-Recupera&ccedil;&atilde;o"
		end if	
		
		
	elseif opcao="B" then	
			tb="TB_NOTA_B"
			gera_pdf="sim"					
			action="../../../../inc/bdb.asp"	
'Antes de 2011 n�o existia os pesos======================================================================	
		if outro="<2011" or  outro="<2011-4" then		
		
			'No 4� per�odo somente PR � lan�ado======================================================================
			if outro="<2011-4" then			
				ln_pesos_cols=""
				ln_pesos_vars=""			
				nm_pesos_vars=""			
				ln_nom_cols="N&ordm;#!#Nome#!#PR#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#M1#!#EXT#!#M2#!#&nbsp;#!#M3"
				nm_vars="p1#!#nav#!#nav#!#nav#!#nav#!#nav#!#nav#!#nav#!#nav#!#nav#!#nav#!#nav#!#media1#!#bon#!#media2#!#nav#!#media3"
				nm_bd="VA_Pr#!#VA_Tr1#!#VA_Tr2#!#VA_Tr3#!#VA_Tr4#!#VA_Str#!#VA_Te1#!#VA_Te2#!#VA_Te3#!#VA_Te4#!#VA_Mte#!#NU_Faltas#!#VA_Media1#!#VA_Extra#!#VA_Media2#!#VA_Rec#!#VA_Media3"
				vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
	
				notas_a_lancar=2
						
	
	
	'Boletim de Avalia&ccedil;&atilde;o				
	'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
				ln_bol_av_cols="Disciplina#!#PR#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#M1#!#EXT#!#M2#!#&nbsp;#!#M3#!#Alterado por#!#Data/Hora"
				ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
				nm_bol_av_vars=nm_vars
				ln_bol_av_vars=nm_bd	
	
				vars_bol_av_calc=vars_calc
				exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#P#!#M#!#M#!#M#!#M#!#M"
				legenda_bol_av="F - Faltas, Tr-Trabalho, Str-Soma dos trabalhos, Te-Teste, Pr�Prova, M-M&eacute;dia, Ext-Nota Extra e Rec-Recupera&ccedil;&atilde;o"
			else
				ln_pesos_cols=""
				ln_pesos_vars=""			
				nm_pesos_vars=""		
				ln_nom_cols="N&ordm;#!#Nome#!#F#!#TR1#!#TR2#!#TR3#!#TR4#!#STR#!#TE1#!#TE2#!#TE3#!#TE4#!#MTE#!#PR#!#M1#!#EXT#!#M2#!#REC#!#M3"
				nm_vars="faltas#!#tr1#!#tr2#!#tr3#!#tr4#!#media_teste#!#t1#!#t2#!#t3#!#t4#!#media_prova#!#p1#!#media1#!#bon#!#media2#!#rec#!#media3"
				nm_bd="NU_Faltas#!#VA_Tr1#!#VA_Tr2#!#VA_Tr3#!#VA_Tr4#!#VA_Str#!#VA_Te1#!#VA_Te2#!#VA_Te3#!#VA_Te4#!#VA_Mte#!#VA_Pr#!#VA_Media1#!#VA_Extra#!#VA_Media2#!#VA_Rec#!#VA_Media3"
				vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
	
				notas_a_lancar=12	
						
	
	
				'Boletim de Avalia&ccedil;&atilde;o				
				'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
				ln_bol_av_cols="Disciplina#!#F#!#TR1#!#TR2#!#TR3#!#TR4#!#STR#!#TE1#!#TE2#!#TE3#!#TE4#!#MTE#!#PR#!#M1#!#EXT#!#M2#!#REC#!#M3#!#Alterado por#!#Data/Hora"
				ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
				nm_bol_av_vars=nm_vars
				ln_bol_av_vars=nm_bd	
	
				vars_bol_av_calc=vars_calc
				exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#P#!#M#!#M#!#M#!#M#!#M"
				legenda_bol_av="F - Faltas, Tr-Trabalho, Str-Soma dos trabalhos, Te-Teste, Pr�Prova, M-M&eacute;dia, Ext-Nota Extra e Rec-Recupera&ccedil;&atilde;o"
			end if							
		else
			'No 4� per�odo somente PR � lan�ado======================================================================

			if outro=4 then			
				ln_pesos_cols=""
				ln_pesos_vars=""			
				nm_pesos_vars=""			
				ln_nom_cols="N&ordm;#!#Nome#!#PR#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#M1#!#EXT#!#M2#!#&nbsp;#!#M3"
				nm_vars="p1#!#nav#!#nav#!#nav#!#nav#!#nav#!#nav#!#nav#!#nav#!#nav#!#nav#!#nav#!#media1#!#bon#!#media2#!#nav#!#media3"
				nm_bd="VA_Pr#!#VA_Tr1#!#VA_Tr2#!#VA_Tr3#!#VA_Tr4#!#VA_Str#!#VA_Te1#!#VA_Te2#!#VA_Te3#!#VA_Te4#!#VA_Mte#!#NU_Faltas#!#VA_Media1#!#VA_Extra#!#VA_Media2#!#VA_Rec#!#VA_Media3"
				vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
	
				notas_a_lancar=2
						
	
	
	'Boletim de Avalia&ccedil;&atilde;o				
	'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
				ln_bol_av_cols="Disciplina#!#PR#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#M1#!#EXT#!#M2#!#&nbsp;#!#M3#!#Alterado por#!#Data/Hora"
				ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
				nm_bol_av_vars=nm_vars
				ln_bol_av_vars=nm_bd	
	
				vars_bol_av_calc=vars_calc
				exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#P#!#M#!#M#!#M#!#M#!#M"
				legenda_bol_av="F - Faltas, Tr-Trabalho, Str-Soma dos trabalhos, Te-Teste, Pr�Prova, M-M&eacute;dia, Ext-Nota Extra e Rec-Recupera&ccedil;&atilde;o"
			else
				ln_pesos_cols="&nbsp;#!#Pesos#!#&nbsp;#!#PESO#!#PESO#!#PESO#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
				ln_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#VA_Pt1#!#VA_Pt2#!#VA_Pt3#!#VA_Pt4#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
				nm_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#va_pt1#!#va_pt2#!#va_pt3#!#va_pt4#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"		
				ln_nom_cols="N&ordm;#!#Nome#!#F#!#TR1#!#TR2#!#TR3#!#TR4#!#STR#!#TE1#!#TE2#!#TE3#!#TE4#!#MTE#!#PR#!#M1#!#EXT#!#M2#!#REC#!#M3"
				nm_vars="faltas#!#tr1#!#tr2#!#tr3#!#tr4#!#media_teste#!#t1#!#t2#!#t3#!#t4#!#media_prova#!#p1#!#media1#!#bon#!#media2#!#rec#!#media3"
				nm_bd="NU_Faltas#!#VA_Tr1#!#VA_Tr2#!#VA_Tr3#!#VA_Tr4#!#VA_Str#!#VA_Te1#!#VA_Te2#!#VA_Te3#!#VA_Te4#!#VA_Mte#!#VA_Pr#!#VA_Media1#!#VA_Extra#!#VA_Media2#!#VA_Rec#!#VA_Media3"
				vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
	
				notas_a_lancar=12	
						
	
	
	'Boletim de Avalia&ccedil;&atilde;o				
	'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
				ln_bol_av_cols="Disciplina#!#F#!#TR1#!#TR2#!#TR3#!#TR4#!#STR#!#TE1#!#TE2#!#TE3#!#TE4#!#MTE#!#PR#!#M1#!#EXT#!#M2#!#REC#!#M3#!#Alterado por#!#Data/Hora"
				ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
				nm_bol_av_vars=nm_vars
				ln_bol_av_vars=nm_bd	
	
				vars_bol_av_calc=vars_calc
				exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#P#!#M#!#M#!#M#!#M#!#M"
				legenda_bol_av="F - Faltas, Tr-Trabalho, Str-Soma dos trabalhos, Te-Teste, Pr�Prova, M-M&eacute;dia, Ext-Nota Extra e Rec-Recupera&ccedil;&atilde;o"
			end if				
		end if

	elseif opcao="C" then
			tb="TB_NOTA_C"
			gera_pdf="sim"				
			action="../../../../inc/bdc.asp"
'Antes de 2011 n�o existia os pesos======================================================================	
		if outro="<2011" or  outro="<2011-4" then		
		
			'No 4� per�odo somente PR � lan�ado======================================================================
			if outro="<2011-4" then		
				ln_pesos_cols=""
				ln_pesos_vars=""			
				nm_pesos_vars=""			
				ln_nom_cols="N&ordm;#!#Nome#!#PR#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#M1#!#EXT#!#M2#!#&nbsp;#!#M3"
				nm_vars="p2#!#nav#!#nav#!#nav#!#nav#!#nav#!#nav#!#nav#!#media1#!#bon#!#media2#!#nav#!#media3"
				nm_bd="VA_PF#!#VA_Tr1#!#VA_Tr2#!#VA_Tr3#!#VA_Tr4#!#VA_Str#!#VA_TE#!#NU_Faltas#!#VA_Media1#!#VA_Extra#!#VA_Media2#!#VA_Rec#!#VA_Media3"
				vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
		
				notas_a_lancar=2	
							
	
	'Boletim de Avalia&ccedil;&atilde;o				
	'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
				ln_bol_av_cols="Disciplina#!#PR#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#M1#!#EXT#!#M2#!#&nbsp;#!#M3#!#Alterado por#!#Data/Hora"
				ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
				nm_bol_av_vars=nm_vars
				ln_bol_av_vars=nm_bd	
	
				vars_bol_av_calc=vars_calc
				exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#A#!#P#!#P#!#M#!#M#!#M#!#M#!#M"
				legenda_bol_av="F - Faltas, Tr-Trabalho, Str-Soma dos trabalhos, Te-Teste, Pr�Prova, M-M&eacute;dia, Ext-Nota Extra e Rec-Recupera&ccedil;&atilde;o"
			else
				ln_pesos_cols=""
				ln_pesos_vars=""			
				nm_pesos_vars=""		
				ln_nom_cols="N&ordm;#!#Nome#!#F#!#TR1#!#TR2#!#TR3#!#TR4#!#STR#!#TE#!#PR#!#M1#!#EXT#!#M2#!#REC#!#M3"
				nm_vars="faltas#!#t1#!#t2#!#t3#!#t4#!#media_teste#!#p1#!#p2#!#media1#!#bon#!#media2#!#rec#!#media3"
				nm_bd="NU_Faltas#!#VA_Tr1#!#VA_Tr2#!#VA_Tr3#!#VA_Tr4#!#VA_Str#!#VA_TE#!#VA_PF#!#VA_Media1#!#VA_Extra#!#VA_Media2#!#VA_Rec#!#VA_Media3"
				vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
		
				notas_a_lancar=9	
							
	
	'Boletim de Avalia&ccedil;&atilde;o				
	'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
				ln_bol_av_cols="Disciplina#!#F#!#TR1#!#TR2#!#TR3#!#TR4#!#STR#!#TE#!#PR#!#M1#!#EXT#!#M2#!#REC#!#M3#!#Alterado por#!#Data/Hora"
				ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
				nm_bol_av_vars=nm_vars
				ln_bol_av_vars=nm_bd	
	
				vars_bol_av_calc=vars_calc
				exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#A#!#A#!#P#!#M#!#M#!#M#!#M#!#M"
				legenda_bol_av="F - Faltas, Tr-Trabalho, Str-Soma dos trabalhos, Te-Teste, Pr�Prova, M-M&eacute;dia, Ext-Nota Extra e Rec-Recupera&ccedil;&atilde;o"	
			end if						
			else
		'No 4� per�odo somente PR � lan�ado======================================================================
			outro=outro*1		
			if outro=4 then					
				ln_pesos_cols=""
				ln_pesos_vars=""			
				nm_pesos_vars=""			
				ln_nom_cols="N&ordm;#!#Nome#!#PR#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#M1#!#EXT#!#M2#!#&nbsp;#!#M3"
				nm_vars="p2#!#nav#!#nav#!#nav#!#nav#!#nav#!#nav#!#nav#!#media1#!#bon#!#media2#!#nav#!#media3"
				nm_bd="VA_PF#!#VA_Tr1#!#VA_Tr2#!#VA_Tr3#!#VA_Tr4#!#VA_Str#!#VA_TE#!#NU_Faltas#!#VA_Media1#!#VA_Extra#!#VA_Media2#!#VA_Rec#!#VA_Media3"
				vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
		
				notas_a_lancar=2	
							
	
	'Boletim de Avalia&ccedil;&atilde;o				
	'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
				ln_bol_av_cols="Disciplina#!#PR#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#M1#!#EXT#!#M2#!#&nbsp;#!#M3#!#Alterado por#!#Data/Hora"
				ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
				nm_bol_av_vars=nm_vars
				ln_bol_av_vars=nm_bd	
	
				vars_bol_av_calc=vars_calc
				exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#A#!#A#!#P#!#M#!#M#!#M#!#M#!#M"
				legenda_bol_av="F - Faltas, Tr-Trabalho, Str-Soma dos trabalhos, Te-Teste, Pr�Prova, M-M&eacute;dia, Ext-Nota Extra e Rec-Recupera&ccedil;&atilde;o"
			else
				ln_pesos_cols="&nbsp;#!#Pesos#!#&nbsp;#!#PESO#!#PESO#!#PESO#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
				ln_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#VA_Pt1#!#VA_Pt2#!#VA_Pt3#!#VA_Pt4#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
				nm_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#va_pt1#!#va_pt2#!#va_pt3#!#va_pt4#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"		
				ln_nom_cols="N&ordm;#!#Nome#!#F#!#TR1#!#TR2#!#TR3#!#TR4#!#STR#!#TE#!#PR#!#M1#!#EXT#!#M2#!#REC#!#M3"
				nm_vars="faltas#!#t1#!#t2#!#t3#!#t4#!#media_teste#!#p1#!#p2#!#media1#!#bon#!#media2#!#rec#!#media3"
				nm_bd="NU_Faltas#!#VA_Tr1#!#VA_Tr2#!#VA_Tr3#!#VA_Tr4#!#VA_Str#!#VA_TE#!#VA_PF#!#VA_Media1#!#VA_Extra#!#VA_Media2#!#VA_Rec#!#VA_Media3"
				vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
		
				notas_a_lancar=9	
							
	
	'Boletim de Avalia&ccedil;&atilde;o				
	'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
				ln_bol_av_cols="Disciplina#!#F#!#TR1#!#TR2#!#TR3#!#TR4#!#STR#!#TE#!#PR#!#M1#!#EXT#!#M2#!#REC#!#M3#!#Alterado por#!#Data/Hora"
				ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
				nm_bol_av_vars=nm_vars
				ln_bol_av_vars=nm_bd	
	
				vars_bol_av_calc=vars_calc
				exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#A#!#A#!#P#!#M#!#M#!#M#!#M#!#M"
				legenda_bol_av="F - Faltas, Tr-Trabalho, Str-Soma dos trabalhos, Te-Teste, Pr�Prova, M-M&eacute;dia, Ext-Nota Extra e Rec-Recupera&ccedil;&atilde;o"	
			end if
		end if
	else
			tb=""
			ln_pesos_cols=""
			ln_pesos_vars=""		
			nm_pesos_vars=""				
			ln_nom_cols="#!#"
			nm_vars="#!#"
			nm_bd="#!#"		
			vars_calc="#!#"	
			action=""
			notas_a_lancar=0
			gera_pdf="nao"
			
'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols=""
			ln_bol_av_span=""
			nm_bol_av_vars=""
			ln_bol_av_vars=""	

			vars_bol_av_calc=""
			legenda_bol_av=""
			exibe_apr_pr_bol_av=""				
	end if	
'============================================================================================================================================


elseif escola="sjohn" or escola="testesjohn" then

	if opcao="A" then		
			tb="TB_NOTA_A"
			ln_pesos_cols="&nbsp;#!#Pesos#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PE_Teste#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PE_Prova#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"			
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#pt#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#pp#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"			
			ln_nom_cols="N&ordm;#!#Nome#!#T1#!#T2#!#T3#!#T4#!#MT#!#P1#!#P2#!#P3#!#MP#!#M1#!#Bon#!#M2#!#REC#!#M3"
			nm_vars="t1#!#t2#!#t3#!#t4#!#media_teste#!#p1#!#p2#!#p3#!#media_prova#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd="VA_Teste1#!#VA_Teste2#!#VA_Teste3#!#VA_Teste4#!#MD_Teste#!#VA_Prova1#!#VA_Prova2#!#VA_Prova3#!#MD_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			ln_nom_cols_lna="N&ordm;#!#Nome#!#T1#!#T2#!#T3#!#T4#!#PesoT#!#MT#!#P1#!#P2#!#P3#!#PesoP#!#MP#!#M1#!#Bon#!#M2#!#REC#!#M3"
			nm_vars_lna="t1#!#t2#!#t3#!#t4#!#pt#!#media_teste#!#p1#!#p2#!#p3#!#pp#!#media_prova#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd_lna="VA_Teste1#!#VA_Teste2#!#VA_Teste3#!#VA_Teste4#!#PE_Teste#!#MD_Teste#!#VA_Prova1#!#VA_Prova2#!#VA_Prova3#!#PE_Prova#!#MD_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc_lna="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"			

			action="../../../../inc/bda.asp"
			notas_a_lancar=9
			notas_a_lancar_lna=notas_a_lancar+2			
			gera_pdf="sim"			

	elseif opcao="B" then	
			tb="TB_NOTA_B"
			ln_pesos_cols="&nbsp;#!#Pesos#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PE_Teste#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PE_Prova#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"			
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#pt#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#pp#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"	
			ln_nom_cols="N&ordm;#!#Nome#!#T1#!#T2#!#T3#!#T4#!#MT#!#P1#!#S#!#P2#!#MP#!#M1#!#Bon#!#M2#!#REC#!#M3"
			nm_vars="t1#!#t2#!#t3#!#t4#!#media_teste#!#p1#!#simul#!#p2#!#media_prova#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd="VA_Teste1#!#VA_Teste2#!#VA_Teste3#!#VA_Teste4#!#MD_Teste#!#VA_Prova1#!#VA_Simul#!#VA_Prova2#!#MD_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			ln_nom_cols_lna="N&ordm;#!#Nome#!#T1#!#T2#!#T3#!#T4#!#PesoT#!#MT#!#P1#!#S#!#P2#!#PesoP#!#MP#!#M1#!#Bon#!#M2#!#REC#!#M3"
			nm_vars_lna="t1#!#t2#!#t3#!#t4#!#pt#!#media_teste#!#p1#!#simul#!#p2#!#pp#!#media_prova#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd_lna="VA_Teste1#!#VA_Teste2#!#VA_Teste3#!#VA_Teste4#!#PE_Teste#!#MD_Teste#!#VA_Prova1#!#VA_Simul#!#VA_Prova2#!#PE_Prova#!#MD_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc_lna="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"		
						
			action="../../../../inc/bdb.asp"	
			notas_a_lancar=9	
			notas_a_lancar_lna=notas_a_lancar+2				
			gera_pdf="sim"							

	elseif opcao="C" then
			tb="TB_NOTA_C"
			ln_pesos_cols="&nbsp;#!#Pesos#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PESO#!#&nbsp;#!#&nbsp;#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PE_Teste#!#&nbsp;#!#&nbsp;#!#PE_Prova#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"			
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#pt#!#&nbsp;#!#&nbsp;#!#pp#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"	
			ln_nom_cols="N&ordm;#!#Nome#!#T1#!#T2#!#T3#!#T4#!#MT#!#P1#!#P2#!#MP#!#M1#!#Bon#!#M2#!#REC#!#M3"
			nm_vars="t1#!#t2#!#t3#!#t4#!#media_teste#!#p1#!#p2#!#media_prova#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd="VA_Teste1#!#VA_Teste2#!#VA_Teste3#!#VA_Teste4#!#MD_Teste#!#VA_Prova1#!#VA_Prova2#!#MD_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			ln_nom_cols_lna="N&ordm;#!#Nome#!#T1#!#T2#!#T3#!#T4#!#PesoT#!#MT#!#P1#!#P2#!#PesoP#!#MP#!#M1#!#Bon#!#M2#!#REC#!#M3"
			nm_vars_lna="t1#!#t2#!#t3#!#t4#!#pt#!#media_teste#!#p1#!#p2#!#pp#!#media_prova#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd_lna="VA_Teste1#!#VA_Teste2#!#VA_Teste3#!#VA_Teste4#!#PE_Teste#!#MD_Teste#!#VA_Prova1#!#VA_Prova2#!#PE_Prova#!#MD_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc_lna="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"			
			action="../../../../inc/bdc.asp"	
			notas_a_lancar=8
			notas_a_lancar_lna=notas_a_lancar+2					
			gera_pdf="sim"
			
	elseif opcao="E" then	
			tb="TB_NOTA_E"
			ln_pesos_cols="&nbsp;#!#Pesos#!#&nbsp;#!#&nbsp;#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PE_Teste#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PE_Prova#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"			
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#pt#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#pp#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"	
			ln_nom_cols="N&ordm;#!#Nome#!#T1#!#T2#!#MT#!#P1#!#S#!#P2#!#MP#!#M1#!#Bon#!#M2#!#REC#!#M3"
			nm_vars="t1#!#t2#!#media_teste#!#p1#!#simul#!#p2#!#media_prova#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd="VA_Teste1#!#VA_Teste2#!#MD_Teste#!#VA_Prova1#!#VA_Simul#!#VA_Prova2#!#MD_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			ln_nom_cols_lna="N&ordm;#!#Nome#!#T1#!#T2#!#PesoT#!#MT#!#P1#!#S#!#P2#!#PesoP#!#MP#!#M1#!#Bon#!#M2#!#REC#!#M3"
			nm_vars_lna="t1#!#t2#!#pt#!#media_teste#!#p1#!#simul#!#p2#!#pp#!#media_prova#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd_lna="VA_Teste1#!#VA_Teste2#!#PE_Teste#!#MD_Teste#!#VA_Prova1#!#VA_Simul#!#VA_Prova2#!#PE_Prova#!#MD_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc_lna="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"				
			
			action="../../../../inc/bde.asp"	
			notas_a_lancar=7
			notas_a_lancar_lna=notas_a_lancar+2					
			gera_pdf="sim"			
			
	elseif opcao="F" then	
			tb="TB_NOTA_F"
			ln_pesos_cols="&nbsp;#!#Pesos#!#&nbsp;#!#&nbsp;#!#PESO#!#PESO#!#PESO#!#&nbsp;#!#&nbsp#!#&nbsp;#!#&nbsp;#!#&nbsp;"
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#PE_Teste#!#PE_Prova1#!#PE_Prova2#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"			
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#pt#!#pp1#!#pp2#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"	
			ln_nom_cols="N&ordm;#!#Nome#!#TD1#!#TD2#!#MTD#!#TS1#!#TS2#!#M1#!#Bon#!#M2#!#REC#!#M3"
			nm_vars="t1#!#t2#!#media_teste#!#p1#!#p2#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd="VA_Teste1#!#VA_Teste2#!#MD_Teste#!#VA_Prova1#!#VA_Prova2#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			ln_nom_cols_lna="N&ordm;#!#Nome#!#T1#!#T2#!#PesoT#!#MT#!#PesoP1#!#P1#!#PesoP2#!#P2#!#MP#!#M1#!#Bon#!#M2#!#REC#!#M3"
			nm_vars_lna="t1#!#t2#!#pt#!#media_teste#!#pp1#!#p1#!#pp2#!#p2#!#media_prova#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd_lna="VA_Teste1#!#VA_Teste2#!#PE_Teste#!#MD_Teste#!#PE_Prova1#!#VA_Prova1#!#PE_Prova2#!#VA_Prova2#!#MD_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc_lna="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"				
			
			action="../../../../inc/bdf.asp"	
			notas_a_lancar=6	
			notas_a_lancar_lna=notas_a_lancar+2				
			gera_pdf="sim"		
	elseif opcao="K" then		
			tb="TB_NOTA_K"
			ln_pesos_cols="&nbsp;#!#Pesos#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"			
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"			
			ln_nom_cols="N&ordm;#!#Nome#!#AV1#!#AV2#!#AV3#!#AV4#!#AV5#!#MAV#!#SIM#!#RS#!#BAT#!#RB#!#M1#!#Bon#!#M2#!#REC#!#M3"
			nm_vars="av1#!#av2#!#av3#!#av4#!#av5#!#media_teste#!#sim#!#rs#!#bat#!#rb#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd="VA_Av1#!#VA_Av2#!#VA_Av3#!#VA_Av4#!#VA_Av5#!#VA_Mav#!#VA_Sim#!#rs#!#VA_Bat#!#rb#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			ln_nom_cols_lna=ln_nom_cols
			nm_vars_lna=nm_vars
			nm_bd_lna=nm_bd
			vars_calc_lna=vars_calc			

			action="../../../../inc/bdk.asp"
			notas_a_lancar=11
			notas_a_lancar_lna=notas_a_lancar+2			
			gera_pdf="sim"				
			
	elseif opcao="L" then		
			tb="TB_NOTA_L"
			ln_pesos_cols="&nbsp;#!#Pesos#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"			
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"			
			ln_nom_cols="N&ordm;#!#Nome#!#T1#!#T2#!#T3#!#T4#!#MT#!#P1#!#P2#!#MP#!#S#!#M1#!#BAT#!#BON#!#M2#!#REC#!#M3"
			nm_vars="t1#!#t2#!#t3#!#t4#!#media_teste#!#p1#!#p2#!#media_prova#!#simul_coord#!#media1#!#bat_coord#!#bon#!#media2#!#rec#!#media3"
			nm_bd="VA_Teste1#!#VA_Teste2#!#VA_Teste3#!#VA_Teste4#!#MD_Teste#!#VA_Prova1#!#VA_Prova2#!#MD_Prova#!#VA_Sim#!#VA_Media1#!#VA_Bat#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			ln_nom_cols_lna=ln_nom_cols
			nm_vars_lna=nm_vars
			nm_bd_lna=nm_bd
			vars_calc_lna=vars_calc			

			action="../../../../inc/bdl.asp"
			notas_a_lancar=10
			notas_a_lancar_lna=notas_a_lancar+2			
			gera_pdf="sim"					

	elseif opcao="M" then		
			tb="TB_NOTA_M"
			ln_pesos_cols="&nbsp;#!#Pesos#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"			
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"			
			ln_nom_cols="N&ordm;#!#Nome#!#AV1#!#AV2#!#AV3#!#AV4#!#AV5#!#SIM#!#MAV#!#BAT#!#BSI#!#M1#!#Bon#!#M2#!#REC#!#M3"
			nm_vars="av1#!#av2#!#av3#!#av4#!#av5#!#simul_coord#!#media_teste#!#bat_coord#!#bsi_coord#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd="VA_Av1#!#VA_Av2#!#VA_Av3#!#VA_Av4#!#VA_Av5#!#VA_Sim#!#VA_Mav#!#VA_Bat#!#VA_Bsi#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			ln_nom_cols_lna=ln_nom_cols
			nm_vars_lna=nm_vars
			nm_bd_lna=nm_bd
			vars_calc_lna=vars_calc			

			action="../../../../inc/bdm.asp"
			notas_a_lancar=10
			notas_a_lancar_lna=notas_a_lancar+2			
			gera_pdf="sim"			

									
	elseif opcao="V" then	
			tb="TB_NOTA_V"
			ln_pesos_cols="&nbsp;#!#Pesos#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"			
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"	
		  'ln_nom_cols="N&ordm;#!#Nome#!#PT1#!#PT2#!#PT3#!#PT4#!#PT5#!#PT6#!#PT7#!#PT8#!#PT9#!#PT10#!#M1#!#Bon#!#M2#!#REC#!#M3"
			ln_nom_cols="N&ordm;#!#Nome#!#AV1#!#AV2#!#PT1#!#PT2#!#PT3#!#PT4#!#PT5#!#PT6#!#PT7#!#PT8#!#M1#!#Bon#!#M2#!#REC#!#M3"
			
			nm_vars="p1#!#p2#!#p3#!#p4#!#p5#!#p6#!#p7#!#p8#!#p9#!#p10#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd="VA_PT1#!#VA_PT2#!#VA_PT3#!#VA_PT4#!#VA_PT5#!#VA_PT6#!#VA_PT7#!#VA_PT8#!#VA_PT9#!#VA_PT10#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			ln_nom_cols_lna=ln_nom_cols
			nm_vars_lna=nm_vars
			nm_bd_lna=nm_bd
			vars_calc_lna=vars_calc			
			action="../../../../inc/bdv.asp"	
			notas_a_lancar=12	
			notas_a_lancar_lna=notas_a_lancar				
			gera_pdf="sim"	
	elseif opcao="LSM" then	
			tb="TB_Area_ConhecimentoxTB_Simulado_Geral"
			ln_pesos_cols=""
			ln_pesos_vars=""			
			nm_pesos_vars=""	
			ln_nom_cols="N&ordm;#!#Nome#!#SIM"
			
			nm_vars="val_simulado"
			nm_bd="VA_SIM"
			vars_calc="0"
			ln_nom_cols_lna=ln_nom_cols
			nm_vars_lna=nm_vars
			nm_bd_lna=nm_bd
			vars_calc_lna=vars_calc			
			action="bd.asp?opt="&opcao	
			notas_a_lancar=1	
			notas_a_lancar_lna=notas_a_lancar				
			gera_pdf="sim"																	
elseif opcao="LBA" then	
			tb="TB_Bonus_Atualidade"
			ln_pesos_cols=""
			ln_pesos_vars=""			
			nm_pesos_vars=""	
			ln_nom_cols="N&ordm;#!#Nome#!#DEFAULT#!#BONUS#!#"
			
			nm_vars="DEFAULT_LBG#!#val_bat#!#CK_LBG"
			nm_bd="DEFAULT_LBG#!#VA_BAT#!#CK_LBG"
			vars_calc="0"
			ln_nom_cols_lna=ln_nom_cols
			nm_vars_lna=nm_vars
			nm_bd_lna=nm_bd
			vars_calc_lna=vars_calc			
			action="bd.asp?opt="&opcao	
			notas_a_lancar=2	
			notas_a_lancar_lna=notas_a_lancar				
			gera_pdf="sim"	
elseif opcao="LBG" then	
			tb="TB_Bonus_Atualidade_Geral"
			ln_pesos_cols=""
			ln_pesos_vars=""			
			nm_pesos_vars=""	
			ln_nom_cols="N&ordm;#!#Nome#!#DEFAULT#!#BONUS#!#"
			
			nm_vars="DEFAULT_LBG#!#val_bat#!#CK_LBG"
			nm_bd="DEFAULT_LBG#!#VA_BAT#!#CK_LBG"
			vars_calc="0"
			ln_nom_cols_lna=ln_nom_cols
			nm_vars_lna=nm_vars
			nm_bd_lna=nm_bd
			vars_calc_lna=vars_calc			
			action="bd.asp?opt="&opcao	
			notas_a_lancar=2	
			notas_a_lancar_lna=notas_a_lancar				
			gera_pdf="sim"	
elseif opcao="LBS" then	
			tb="TB_Bonus_Simulado"
			ln_pesos_cols=""
			ln_pesos_vars=""			
			nm_pesos_vars=""	
			ln_nom_cols="N&ordm;#!#Nome#!#DEFAULT#!#BONUS SIM#!#"
			
			nm_vars="DEFAULT_LBS#!#val_bsi#!#CK_LBS"
			nm_bd="DEFAULT_LBS#!#VA_BSI#!#CK_LBS"
			vars_calc="0"
			ln_nom_cols_lna=ln_nom_cols
			nm_vars_lna=nm_vars
			nm_bd_lna=nm_bd
			vars_calc_lna=vars_calc			
			action="bd.asp?opt="&opcao	
			notas_a_lancar=2	
			notas_a_lancar_lna=notas_a_lancar				
			gera_pdf="sim"	
	else
			tb=""
			ln_pesos_cols=""
			ln_pesos_vars=""		
			nm_pesos_vars=""				
			ln_nom_cols="#!#"
			nm_vars="#!#"
			nm_bd="#!#"		
			vars_calc="#!#"	
			ln_nom_cols_lna=ln_nom_cols
			nm_vars_lna=nm_vars
			nm_bd_lna=nm_bd
			vars_calc_lna=vars_calc				
			action=""
			notas_a_lancar=0
			notas_a_lancar_lna=notas_a_lancar					
			gera_pdf="nao"
	end if	
	
'============================================================================================================================================	
	
elseif escola="stockler" or escola="testestockler" then

	if opcao="A" then		
			tb="TB_NOTA_A"
			ln_pesos_cols=""
			ln_pesos_vars=""			
			nm_pesos_vars=""			
			ln_nom_cols="N&ordm;#!#Nome#!#F#!#AV1#!#AV2#!#AV3#!#MAV#!#PR#!#M1#!#AT#!#M2#!#REC#!#M3"
			nm_vars="faltas#!#av1#!#av2#!#av3#!#media_teste#!#pr#!#media1#!#at#!#media2#!#rec#!#media3"
			nm_bd="NU_Faltas#!#VA_AV1#!#VA_AV2#!#VA_AV3#!#MAV_Avaliacao#!#VA_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			action="../../../../inc/bda.asp"
			notas_a_lancar=7
			gera_pdf="sim"			



'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols="Disciplina#!#F#!#AV1#!#AV2#!#AV3#!#MAV#!#PR#!#M1#!#AT#!#M2#!#REC#!#M3#!#Alterado por#!#Data/Hora"
			ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			nm_bol_av_vars=nm_vars
			ln_bol_av_vars=nm_bd	

			vars_bol_av_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#P#!#M#!#M#!#M#!#M#!#M"
			legenda_bol_av="F - Faltas, AV - Avalia&ccedil;&atilde;o, MAV - M�dia das Avalia��es, PR - Prova, M - M�dias, AT - Atualidades, REC - Recupera��o"


	elseif opcao="B" then	
			tb="TB_NOTA_B"
			ln_pesos_cols=""
			ln_pesos_vars=""			
			nm_pesos_vars=""			
			ln_nom_cols="N&ordm;#!#Nome#!#F#!#AV1#!#AV2#!#AV3#!#MAV#!#PR#!#M1#!#AT#!#M2#!#REC#!#M3"
			nm_vars="faltas#!#av1#!#av2#!#av3#!#media_teste#!#pr#!#media1#!#at#!#media2#!#rec#!#media3"
			nm_bd="NU_Faltas#!#VA_AV1#!#VA_AV2#!#VA_AV3#!#MAV_Avaliacao#!#VA_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			action="../../../../inc/bdb.asp"	
			notas_a_lancar=7	
			gera_pdf="sim"							

'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols="Disciplina#!#F#!#AV1#!#AV2#!#AV3#!#MAV#!#PR#!#M1#!#AT#!#M2#!#REC#!#M3#!#Alterado por#!#Data/Hora"
			ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			nm_bol_av_vars=nm_vars
			ln_bol_av_vars=nm_bd	

			vars_bol_av_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#P#!#M#!#M#!#M#!#M#!#M"
			legenda_bol_av="F - Faltas, AV - Avalia&ccedil;&atilde;o, MAV - M�dia das Avalia��es, PR - Prova, M - M�dias, AT - Atualidades, REC - Recupera��o"


	elseif opcao="C" then
			tb=""
			ln_pesos_cols=""
			ln_pesos_vars=""		
			nm_pesos_vars=""				
			ln_nom_cols="#!#"
			nm_vars="#!#"
			nm_bd="#!#"		
			vars_calc="#!#"	
			action=""
			notas_a_lancar=0
			gera_pdf="nao"
			
'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols=""
			ln_bol_av_span=""
			nm_bol_av_vars=""
			ln_bol_av_vars=""	

			vars_bol_av_calc=""
			legenda_bol_av=""
			exibe_apr_pr_bol_av=""		
	else
			tb=""
			ln_pesos_cols=""
			ln_pesos_vars=""		
			nm_pesos_vars=""				
			ln_nom_cols="#!#"
			nm_vars="#!#"
			nm_bd="#!#"		
			vars_calc="#!#"	
			action=""
			notas_a_lancar=0
			gera_pdf="nao"
			
'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols=""
			ln_bol_av_span=""
			nm_bol_av_vars=""
			ln_bol_av_vars=""	

			vars_bol_av_calc=""
			legenda_bol_av=""
			exibe_apr_pr_bol_av=""				
	end if		


'============================================================================================================================================	
	
elseif escola="vitoria" or escola="testevitoria" then

	if opcao="A" then		
			tb=""
			ln_pesos_cols=""
			ln_pesos_vars=""		
			nm_pesos_vars=""				
			ln_nom_cols="#!#"
			nm_vars="#!#"
			nm_bd="#!#"		
			vars_calc="#!#"	
			action=""
			notas_a_lancar=0
			gera_pdf="nao"
			
'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols=""
			ln_bol_av_span=""
			nm_bol_av_vars=""
			ln_bol_av_vars=""	

			vars_bol_av_calc=""
			legenda_bol_av=""
			exibe_apr_pr_bol_av=""	
			
'	elseif opcao="B" then	
'			tb="TB_NOTA_B"
'			ln_pesos_cols=""
'			ln_pesos_vars=""			
'			nm_pesos_vars=""
'			ln_nom_cols="N&ordm;#!#Nome#!#F#!#Te1#!#Te2#!#MTE#!#Tr1#!#Tr2#!#MTR#!#Sim#!#Cf#!#Pr#!#M1#!#Bon#!#M2#!#Rec#!#Cfr#!#M3"
'			nm_vars="faltas#!#te1#!#te2#!#media_teste#!#tr1#!#tr2#!#media_prova#!#sim#!#cf#!#pr#!#media1#!#bon#!#media2#!#rec#!#cfr#!#media3"
'			nm_bd="NU_Faltas#!#VA_TE1#!#VA_TE2#!#VA_MTE#!#VA_TR1#!#VA_TR2#!#VA_MTR#!#VA_SIM#!#VA_CF#!#VA_PR#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_CFR#!#VA_Media3"
'			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
'			action="../../../../inc/bdb.asp"
'			notas_a_lancar=11
'			gera_pdf="sim"		
'			
''Boletim de Avalia&ccedil;&atilde;o				
''Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
'			ln_bol_av_cols="Disciplina#!#F#!#Te1#!#Te2#!#MTE#!#Tr1#!#Tr2#!#MTR#!#Sim#!#Cf#!#Pr#!#M1#!#Bon#!#M2#!#Rec#!#Cfr#!#M3#!#Alterado por#!#Data/Hora"
'			ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
'			nm_bol_av_vars=nm_vars
'			ln_bol_av_vars=nm_bd	
'
'			vars_bol_av_calc=vars_calc
'			exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#A#!#A#!#P#!#P#!#P#!#M#!#M#!#M#!#M#!#M#!#M"
'			legenda_bol_av="F - Faltas, Te-Testes, MTE-M&eacute;dia dos Testes, Tr-Trabalho, MTR-M&eacute;dia dos Trabalhos, Sim - Simulado, Cf - Conceito Formativo, Pr� Prova, M-M&eacute;dia, Bon-B&ocirc;nus e Rec-Recupera&ccedil;&atilde;o"							
	elseif opcao="B" then	
			tb="TB_NOTA_B"
			ln_pesos_cols=""
			ln_pesos_vars=""			
			nm_pesos_vars=""
			ln_nom_cols="N&ordm;#!#Nome#!#F#!#Te1#!#Te2#!#MTE#!#Tr1#!#Tr2#!#MTR#!#Cf#!#Pr#!#M1#!#Bon#!#M2#!#Rec#!#Cfr#!#M3"
			nm_vars="faltas#!#te1#!#te2#!#media_teste#!#tr1#!#tr2#!#media_prova#!#cf#!#pr#!#media1#!#bon#!#media2#!#rec#!#cfr#!#media3"
			nm_bd="NU_Faltas#!#VA_TE1#!#VA_TE2#!#VA_MTE#!#VA_TR1#!#VA_TR2#!#VA_MTR#!#VA_CF#!#VA_PR#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_CFR#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			action="../../../../inc/bdb.asp"
			notas_a_lancar=10
			gera_pdf="sim"		
			
'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols="Disciplina#!#F#!#Te1#!#Te2#!#MTE#!#Tr1#!#Tr2#!#MTR#!#Sim#!#Cf#!#Pr#!#M1#!#Bon#!#M2#!#Rec#!#Cfr#!#M3#!#Alterado por#!#Data/Hora"
			ln_bol_av_wf_cols="Disciplina#!#F#!#Te1#!#Te2#!#MTE#!#Tr1#!#Tr2#!#MTR#!#Cf#!#Pr#!#M1#!#Bon#!#M2#!#Rec#!#Cfr#!#M3#!#Alterado por#!#Data/Hora"			

			ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"			
			ln_bol_av_wf_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			
			nm_bol_av_vars="faltas#!#te1#!#te2#!#media_teste#!#tr1#!#tr2#!#media_prova#!#sim#!#cf#!#pr#!#media1#!#bon#!#media2#!#rec#!#cfr#!#media3"
			nm_bol_av_wf_vars="faltas#!#te1#!#te2#!#media_teste#!#tr1#!#tr2#!#media_prova#!#cf#!#pr#!#media1#!#bon#!#media2#!#rec#!#cfr#!#media3"
						
			ln_bol_av_vars="NU_Faltas#!#VA_TE1#!#VA_TE2#!#VA_MTE#!#VA_TR1#!#VA_TR2#!#VA_MTR#!#VA_SIM#!#VA_CF#!#VA_PR#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_CFR#!#VA_Media3"	
			ln_bol_av_wf_vars="NU_Faltas#!#VA_TE1#!#VA_TE2#!#VA_MTE#!#VA_TR1#!#VA_TR2#!#VA_MTR#!#VA_CF#!#VA_PR#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_CFR#!#VA_Media3"				

			vars_bol_av_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			vars_bol_av_wf_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
						
			exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#A#!#A#!#P#!#P#!#P#!#M#!#M#!#M#!#M#!#M#!#M"
			exibe_apr_pr_bol_av_wf="0#!#A#!#A#!#A#!#A#!#A#!#A#!#P#!#P#!#M#!#M#!#M#!#M#!#M#!#M"
						
			legenda_bol_av="F - Faltas, Te-Testes, MTE-M&eacute;dia dos Testes, Tr-Trabalho, MTR-M&eacute;dia dos Trabalhos, Sim - Simulado, Cf - Conceito Formativo, Pr� Prova, M-M&eacute;dia, Bon-B&ocirc;nus e Rec-Recupera&ccedil;&atilde;o"
	elseif opcao="BWF" then	
	'Usado apenas para exibir o boletim de avalia��es do web fam�lia
			tb="TB_NOTA_B"
			ln_pesos_cols=""
			ln_pesos_vars=""			
			nm_pesos_vars=""
			ln_nom_cols="N&ordm;#!#Nome#!#F#!#Te1#!#Te2#!#MTE#!#Tr1#!#Tr2#!#MTR#!#Cf#!#Pr#!#M1#!#Bon#!#M2#!#Rec#!#Cfr#!#M3"
			nm_vars="faltas#!#te1#!#te2#!#media_teste#!#tr1#!#tr2#!#media_prova#!#cf#!#pr#!#media1#!#bon#!#media2#!#rec#!#cfr#!#media3"
			nm_bd="NU_Faltas#!#VA_TE1#!#VA_TE2#!#VA_MTE#!#VA_TR1#!#VA_TR2#!#VA_MTR#!#VA_CF#!#VA_PR#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_CFR#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			action="../../../../inc/bdb.asp"
			notas_a_lancar=10
			gera_pdf="sim"		
			
'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols="Disciplina#!#F#!#Te1#!#Te2#!#MTE#!#Tr1#!#Tr2#!#MTR#!#Cf#!#Pr#!#M1#!#Bon#!#M2#!#Rec#!#Cfr#!#M3#!#Alterado por#!#Data/Hora"			

			ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			
			nm_bol_av_vars="faltas#!#te1#!#te2#!#media_teste#!#tr1#!#tr2#!#media_prova#!#cf#!#pr#!#media1#!#bon#!#media2#!#rec#!#cfr#!#media3"
						
			ln_bol_av_vars="NU_Faltas#!#VA_TE1#!#VA_TE2#!#VA_MTE#!#VA_TR1#!#VA_TR2#!#VA_MTR#!#VA_CF#!#VA_PR#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_CFR#!#VA_Media3"				

			vars_bol_av_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
						
			exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#A#!#A#!#P#!#P#!#M#!#M#!#M#!#M#!#M#!#M"
						
			legenda_bol_av="F - Faltas, Te-Testes, MTE-M&eacute;dia dos Testes, Tr-Trabalho, MTR-M&eacute;dia dos Trabalhos, Sim - Simulado, Cf - Conceito Formativo, Pr� Prova, M-M&eacute;dia, Bon-B&ocirc;nus e Rec-Recupera&ccedil;&atilde;o"					
	elseif opcao="C" then
			tb="TB_NOTA_C"
			ln_pesos_cols=""
			ln_pesos_vars=""			
			nm_pesos_vars=""
			ln_nom_cols="N&ordm;#!#Nome#!#F#!#Am#!#Ab#!#Cf#!#M1#!#Bon#!#M2#!#Rec#!#M3"
			nm_vars="faltas#!#am#!#ab#!#cf#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd="NU_Faltas#!#VA_AM#!#VA_AB#!#VA_CF#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			action="../../../../inc/bdc.asp"
			notas_a_lancar=6
			gera_pdf="sim"		
			
'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols="Disciplina#!#F#!#Am#!#Ab#!#Cf#!#M1#!#Bon#!#M2#!#Rec#!#M3#!#Alterado por#!#Data/Hora"
			ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			nm_bol_av_vars=nm_vars
			ln_bol_av_vars=nm_bd	

			vars_bol_av_calc=vars_calc
			exibe_apr_pr_bol_av="0#!#A#!#A#!#P#!#M#!#M#!#M#!#M#!#M"
			legenda_bol_av="F - Faltas, Am-Avalia&ccedil;&atilde;o Mensal, Ab-Avalia&ccedil;&atilde;o Bimestral, Cf - Conceito Formativo, M-M&eacute;dia, Bon-B&ocirc;nus e Rec-Recupera&ccedil;&atilde;o"				
	

	elseif opcao="E" then
			tb="TB_NOTA_E"
			ln_pesos_cols=""
			ln_pesos_vars=""			
			nm_pesos_vars=""
			ln_nom_cols="N&ordm;#!#Nome#!#F#!#At#!#Pt#!#T#!#Cf#!#M1#!#Bon#!#M2#!#Rec#!#M3"
			nm_vars="faltas#!#at#!#pt#!#t#!#cf#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd="NU_Faltas#!#VA_AT#!#VA_PT#!#VA_T#!#VA_CF#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			action="../../../../inc/bde.asp"
			notas_a_lancar=7
			gera_pdf="sim"		
			
'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols="Disciplina#!#F#!#Am#!#Ab#!#T#!#Cf#!#M1#!#Bon#!#M2#!#Rec#!#M3#!#Alterado por#!#Data/Hora"
			ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			nm_bol_av_vars=nm_vars
			ln_bol_av_vars=nm_bd	

			vars_bol_av_calc=vars_calc
			exibe_apr_pr_bol_av="0#!#A#!#A#!#P#!#P#!#M#!#M#!#M#!#M#!#M"
			legenda_bol_av="F - Faltas, Am-Avalia&ccedil;&atilde;o Mensal, Ab-Avalia&ccedil;&atilde;o Bimestral, Cf - Conceito Formativo, M-M&eacute;dia, Bon-B&ocirc;nus e Rec-Recupera&ccedil;&atilde;o"				
	

	elseif opcao="F" then
			tb="TB_NOTA_F"
			ln_pesos_cols=""
			ln_pesos_vars=""			
			nm_pesos_vars=""
			ln_nom_cols="N&ordm;#!#Nome#!#F#!#Am#!#At#!#Cf#!#M1#!#Bon#!#M2#!#Rec#!#M3"
			nm_vars="faltas#!#am#!#at#!#cf#!#media1#!#bon#!#media2#!#rec#!#media3"
			nm_bd="NU_Faltas#!#VA_AM#!#VA_AT#!#VA_CF#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			action="../../../../inc/bdf.asp"
			notas_a_lancar=6
			gera_pdf="sim"		
			
'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols="Disciplina#!#F#!#Am#!#Ab#!#Cf#!#M1#!#Bon#!#M2#!#Rec#!#M3#!#Alterado por#!#Data/Hora"
			ln_bol_av_span="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			nm_bol_av_vars=nm_vars
			ln_bol_av_vars=nm_bd	

			vars_bol_av_calc=vars_calc
			exibe_apr_pr_bol_av="0#!#A#!#A#!#P#!#M#!#M#!#M#!#M#!#M"
			legenda_bol_av="F - Faltas, Am-Avalia&ccedil;&atilde;o Mensal, Ab-Avalia&ccedil;&atilde;o Bimestral, Cf - Conceito Formativo, M-M&eacute;dia, Bon-B&ocirc;nus e Rec-Recupera&ccedil;&atilde;o"				
	

	else
			tb=""
			ln_pesos_cols=""
			ln_pesos_vars=""		
			nm_pesos_vars=""				
			ln_nom_cols="#!#"
			nm_vars="#!#"
			nm_bd="#!#"		
			vars_calc="#!#"	
			action=""
			notas_a_lancar=0
			gera_pdf="nao"
			
'Boletim de Avalia&ccedil;&atilde;o				
'Aten��o na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a fun��o Boletim de Avalia&ccedil;&atilde;o	
			ln_bol_av_cols=""
			ln_bol_av_span=""
			nm_bol_av_vars=""
			ln_bol_av_vars=""	

			vars_bol_av_calc=""
			legenda_bol_av=""
			exibe_apr_pr_bol_av=""				
	end if				
	
								
end if
'itens 0 a 9 Planilha de notas
'itens 0 a 10 Planilha de notas PDF
'itens 0 e 11 a 17 Boletim de Avalia&ccedil;&atilde;o
'itens 18 a 22 Lan�amento de notas por aluno
verifica_dados_tabela=tb&"#$#"&ln_pesos_cols&"#$#"&ln_pesos_vars&"#$#"&nm_pesos_vars&"#$#"&ln_nom_cols&"#$#"&nm_vars&"#$#"&nm_bd&"#$#"&vars_calc&"#$#"&action&"#$#"&notas_a_lancar&"#$#"&gera_pdf&"#$#"&ln_bol_av_cols&"#$#"&ln_bol_av_span&"#$#"&nm_bol_av_vars&"#$#"&ln_bol_av_vars&"#$#"&vars_bol_av_calc&"#$#"&legenda_bol_av&"#$#"&exibe_apr_pr_bol_av&"#$#"&			ln_nom_cols_lna&"#$#"&nm_vars_lna&"#$#"&nm_bd_lna&"#$#"&vars_calc_lna&"#$#"&notas_a_lancar_lna	
end function
%>