<%
tp_arredondamento="mat_dez"
email_suporte="suportewebdiretordinamis@webdiretor.com.br"

Function dados_planilha_notas(ano_letivo,unidade,curso,etapa,turma,disciplina_mae,disciplina_filha,periodo,opcao,outro)

'response.Write(opcao&"-"&periodo)

	if opcao="A" then		
		tb="TB_NOTA_A"
		action="../../../../inc/bda.asp"
		gera_pdf="sim"		

		if periodo<4 then
			ln_pesos_cols="&nbsp;#!#Pesos das Aprs#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"		
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#va_v_apr1#!#va_v_apr2#!#va_v_apr3#!#va_v_apr4#!#va_v_apr5#!#va_v_apr6#!#va_v_apr7#!#va_v_apr8#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;&nbsp;"
		else
			ln_pesos_cols="&nbsp;#!#Pesos das Aprs#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"		
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#va_v_apr1#!#va_v_apr2#!#va_v_apr3#!#va_v_apr4#!#va_v_apr5#!#va_v_apr6#!#va_v_apr7#!#va_v_apr8#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
		end if			
			
		if periodo=1 then
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#V_Apr1_P1#!#V_Apr2_P1#!#V_Apr3_P1#!#V_Apr4_P1#!#V_Apr5_P1#!#V_Apr6_P1#!#V_Apr7_P1#!#V_Apr8_P1#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"	
			ln_nom_cols="N&ordm;#!#Nome#!#Apr1#!#Apr2#!#Apr3#!#Apr4#!#Apr5#!#Apr6#!#Tec1#!#Tec2#!#Sapr#!#Pr1#!#Pr2#!#Bon#!#Me#!#Mc#!#Faltas"
			nm_vars="Apr1_P1#!#Apr2_P1#!#Apr3_P1#!#Apr4_P1#!#Apr5_P1#!#Apr6_P1#!#Apr7_P1#!#Apr8_P1#!#VA_Sapr1#!#VA_Pr1#!#VA_Te1#!#VA_Bon1#!#VA_Me1#!#VA_Mc1#!#NU_Faltas_P1"
			nm_bd=nm_vars
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"	
			notas_a_lancar=12		
		elseif periodo=2 then
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#V_Apr1_P2#!#V_Apr2_P2#!#V_Apr3_P2#!#V_Apr4_P2#!#V_Apr5_P2#!#V_Apr6_P2#!#V_Apr7_P2#!#V_Apr8_P2#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"	
			ln_nom_cols="N&ordm;#!#Nome#!#Apr1#!#Apr2#!#Apr3#!#Apr4#!#Apr5#!#Apr6#!#Tec1#!#Tec2#!#Sapr#!#Pr1#!#Pr2#!#ECE1#!#Me#!#Mc*#!#Faltas"
			nm_vars="Apr1_P2#!#Apr2_P2#!#Apr3_P2#!#Apr4_P2#!#Apr5_P2#!#Apr6_P2#!#Apr7_P2#!#Apr8_P2#!#VA_Sapr2#!#VA_Pr2#!#VA_Te2#!#VA_Bon2#!#VA_Me2#!#VA_Mc2#!#NU_Faltas_P2"
			nm_bd=nm_vars
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"			
			notas_a_lancar=12						
		elseif periodo=3 then
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#V_Apr1_P3#!#V_Apr2_P3#!#V_Apr3_P3#!#V_Apr4_P3#!#V_Apr5_P3#!#V_Apr6_P3#!#V_Apr7_P3#!#V_Apr8_P3#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"	
			ln_nom_cols="N&ordm;#!#Nome#!#Apr1#!#Apr2#!#Apr3#!#Apr4#!#Apr5#!#Apr6#!#Tec1#!#Tec2#!#Sapr#!#Pr1#!#Pr2#!#Bon#!#Me#!#Mc#!#Faltas"
			nm_vars="Apr1_P3#!#Apr2_P3#!#Apr3_P3#!#Apr4_P3#!#Apr5_P3#!#Apr6_P3#!#Apr7_P3#!#Apr8_P3#!#VA_Sapr3#!#VA_Pr3#!#VA_Te3#!#VA_Bon3#!#VA_Me3#!#VA_Mc3#!#NU_Faltas_P3"
			nm_bd=nm_vars
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"	
			notas_a_lancar=12					
		elseif periodo=4 then
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#V_Apr1_EC#!#V_Apr2_EC#!#V_Apr3_EC#!#V_Apr4_EC#!#V_Apr5_EC#!#V_Apr6_EC#!#V_Apr7_EC#!#V_Apr8_EC#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"	
			ln_nom_cols="N&ordm;#!#Nome#!#Apr1#!#Apr2#!#Apr3#!#Apr4#!#Apr5#!#Apr6#!#Tec1#!#Tec2#!#Sapr#!#Pr1#!#Me#!#Mc"
			nm_vars="Apr1_EC#!#Apr2_EC#!#Apr3_EC#!#Apr4_EC#!#Apr5_EC#!#Apr6_EC#!#Apr7_EC#!#Apr8_EC#!#VA_Sapr_EC#!#VA_Pr4#!#VA_Me_EC#!#VA_Mfinal"
			nm_bd=nm_vars
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			notas_a_lancar=9					
		end if
	elseif opcao="B" then	
		tb="TB_NOTA_B"
		action="../../../../inc/bdb.asp"
		gera_pdf="sim"

		if periodo<4 then
			ln_pesos_cols="&nbsp;#!#Pesos das Aprs#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"		
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#va_v_apr1#!#va_v_apr2#!#va_v_apr3#!#va_v_apr4#!#va_v_apr5#!#va_v_apr6#!#va_v_apr7#!#va_v_apr8#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;&nbsp;"
		else
			ln_pesos_cols="&nbsp;#!#Pesos das Aprs#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"		
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#va_v_apr1#!#va_v_apr2#!#va_v_apr3#!#va_v_apr4#!#va_v_apr5#!#va_v_apr6#!#va_v_apr7#!#va_v_apr8#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
		end if			
	
		if periodo=1 then
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#V_Apr1_P1#!#V_Apr2_P1#!#V_Apr3_P1#!#V_Apr4_P1#!#V_Apr5_P1#!#V_Apr6_P1#!#V_Apr7_P1#!#V_Apr8_P1#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"	
			ln_nom_cols="N&ordm;#!#Nome#!#Apr1#!#Apr2#!#Apr3#!#Apr4#!#Apr5#!#Apr6#!#Tec1#!#Tec2#!#Sapr#!#Pr1#!#Pr2#!#Bon#!#Me#!#Mc#!#Faltas"
			nm_vars="Apr1_P1#!#Apr2_P1#!#Apr3_P1#!#Apr4_P1#!#Apr5_P1#!#Apr6_P1#!#Apr7_P1#!#Apr8_P1#!#VA_Sapr1#!#VA_Pr1#!#VA_Te1#!#VA_Bon1#!#VA_Me1#!#VA_Mc1#!#NU_Faltas_P1"
			nm_bd=nm_vars
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"	
			notas_a_lancar=12		
		elseif periodo=2 then
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#V_Apr1_P2#!#V_Apr2_P2#!#V_Apr3_P2#!#V_Apr4_P2#!#V_Apr5_P2#!#V_Apr6_P2#!#V_Apr7_P2#!#V_Apr8_P2#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"	
			ln_nom_cols="N&ordm;#!#Nome#!#Apr1#!#Apr2#!#Apr3#!#Apr4#!#Apr5#!#Apr6#!#Tec1#!#Tec2#!#Sapr#!#Pr1#!#Pr2#!#ECE1#!#Me#!#Mc*#!#Faltas"
			nm_vars="Apr1_P2#!#Apr2_P2#!#Apr3_P2#!#Apr4_P2#!#Apr5_P2#!#Apr6_P2#!#Apr7_P2#!#Apr8_P2#!#VA_Sapr2#!#VA_Pr2#!#VA_Te2#!#VA_Bon2#!#VA_Me2#!#VA_Mc2#!#NU_Faltas_P2"
			nm_bd=nm_vars
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"		
			notas_a_lancar=12						
		elseif periodo=3 then
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#V_Apr1_P3#!#V_Apr2_P3#!#V_Apr3_P3#!#V_Apr4_P3#!#V_Apr5_P3#!#V_Apr6_P3#!#V_Apr7_P3#!#V_Apr8_P3#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"	
			ln_nom_cols="N&ordm;#!#Nome#!#Apr1#!#Apr2#!#Apr3#!#Apr4#!#Apr5#!#Apr6#!#Tec1#!#Tec2#!#Sapr#!#Pr1#!#Pr2#!#Bon#!#Me#!#Mc#!#Faltas"
			nm_vars="Apr1_P3#!#Apr2_P3#!#Apr3_P3#!#Apr4_P3#!#Apr5_P3#!#Apr6_P3#!#Apr7_P3#!#Apr8_P3#!#VA_Sapr3#!#VA_Pr3#!#VA_Te3#!#VA_Bon3#!#VA_Me3#!#VA_Mc3#!#NU_Faltas_P3"
			nm_bd=nm_vars
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"	
			notas_a_lancar=12					
		elseif periodo=4 then
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#V_Apr1_EC#!#V_Apr2_EC#!#V_Apr3_EC#!#V_Apr4_EC#!#V_Apr5_EC#!#V_Apr6_EC#!#V_Apr7_EC#!#V_Apr8_EC#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"	
			ln_nom_cols="N&ordm;#!#Nome#!#Apr1#!#Apr2#!#Apr3#!#Apr4#!#Apr5#!#Apr6#!#Tec1#!#Tec2#!#Sapr#!#Pr1#!#Me#!#Mc"
			nm_vars="Apr1_EC#!#Apr2_EC#!#Apr3_EC#!#Apr4_EC#!#Apr5_EC#!#Apr6_EC#!#Apr7_EC#!#Apr8_EC#!#VA_Sapr_EC#!#VA_Pr4#!#VA_Me_EC#!#VA_Mfinal"
			nm_bd=nm_vars
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			notas_a_lancar=9					
		end if
	
	elseif opcao="C" then
		tb="TB_NOTA_C"
		action="../../../../inc/bdc.asp"
		gera_pdf="sim"
	
		if periodo<4 then
			ln_pesos_cols="&nbsp;#!#Pesos das Aprs#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"		
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#va_v_apr1#!#va_v_apr2#!#va_v_apr3#!#va_v_apr4#!#va_v_apr5#!#va_v_apr6#!#va_v_apr7#!#va_v_apr8#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;&nbsp;"
		else
			ln_pesos_cols="&nbsp;#!#Pesos das Aprs#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#PESO#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"		
			nm_pesos_vars="&nbsp;#!#&nbsp;#!#va_v_apr1#!#va_v_apr2#!#va_v_apr3#!#va_v_apr4#!#va_v_apr5#!#va_v_apr6#!#va_v_apr7#!#va_v_apr8#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
		end if			
		
		if periodo=1 then
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#V_Apr1_P1#!#V_Apr2_P1#!#V_Apr3_P1#!#V_Apr4_P1#!#V_Apr5_P1#!#V_Apr6_P1#!#V_Apr7_P1#!#V_Apr8_P1#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"	
			ln_nom_cols="N&ordm;#!#Nome#!#Apr1#!#Apr2#!#Apr3#!#Apr4#!#Apr5#!#Apr6#!#Tec1#!#Tec2#!#Sapr#!#Pr1#!#Pr2#!#Bon#!#Me#!#Mc#!#Faltas"
			nm_vars="Apr1_P1#!#Apr2_P1#!#Apr3_P1#!#Apr4_P1#!#Apr5_P1#!#Apr6_P1#!#Apr7_P1#!#Apr8_P1#!#VA_Sapr1#!#VA_Pr1#!#VA_Te1#!#VA_Bon1#!#VA_Me1#!#VA_Mc1#!#NU_Faltas_P1"
			nm_bd=nm_vars
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"	
			notas_a_lancar=12		
		elseif periodo=2 then
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#V_Apr1_P2#!#V_Apr2_P2#!#V_Apr3_P2#!#V_Apr4_P2#!#V_Apr5_P2#!#V_Apr6_P2#!#V_Apr7_P2#!#V_Apr8_P2#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"	
			ln_nom_cols="N&ordm;#!#Nome#!#Apr1#!#Apr2#!#Apr3#!#Apr4#!#Apr5#!#Apr6#!#Tec1#!#Tec2#!#Sapr#!#Pr1#!#Pr2#!#ECE1#!#Me#!#Mc*#!#Faltas"
			nm_vars="Apr1_P2#!#Apr2_P2#!#Apr3_P2#!#Apr4_P2#!#Apr5_P2#!#Apr6_P2#!#Apr7_P2#!#Apr8_P2#!#VA_Sapr2#!#VA_Pr2#!#VA_Te2#!#VA_Bon2#!#VA_Me2#!#VA_Mc2#!#NU_Faltas_P2"
			nm_bd=nm_vars
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"		
			notas_a_lancar=12						
		elseif periodo=3 then
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#V_Apr1_P3#!#V_Apr2_P3#!#V_Apr3_P3#!#V_Apr4_P3#!#V_Apr5_P3#!#V_Apr6_P3#!#V_Apr7_P3#!#V_Apr8_P3#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"	
			ln_nom_cols="N&ordm;#!#Nome#!#Apr1#!#Apr2#!#Apr3#!#Apr4#!#Apr5#!#Apr6#!#Tec1#!#Tec2#!#Sapr#!#Pr1#!#Pr2#!#Bon#!#Me#!#Mc#!#Faltas"
			nm_vars="Apr1_P3#!#Apr2_P3#!#Apr3_P3#!#Apr4_P3#!#Apr5_P3#!#Apr6_P3#!#Apr7_P3#!#Apr8_P3#!#VA_Sapr3#!#VA_Pr3#!#VA_Te3#!#VA_Bon3#!#VA_Me3#!#VA_Mc3#!#NU_Faltas_P3"
			nm_bd=nm_vars
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"	
			notas_a_lancar=12					
		elseif periodo=4 then
			ln_pesos_vars="&nbsp;#!#&nbsp;#!#V_Apr1_EC#!#V_Apr2_EC#!#V_Apr3_EC#!#V_Apr4_EC#!#V_Apr5_EC#!#V_Apr6_EC#!#V_Apr7_EC#!#V_Apr8_EC#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"	
			ln_nom_cols="N&ordm;#!#Nome#!#Apr1#!#Apr2#!#Apr3#!#Apr4#!#Apr5#!#Apr6#!#Tec1#!#Tec2#!#Sapr#!#Pr1#!#Me#!#Mc"
			nm_vars="Apr1_EC#!#Apr2_EC#!#Apr3_EC#!#Apr4_EC#!#Apr5_EC#!#Apr6_EC#!#Apr7_EC#!#Apr8_EC#!#VA_Sapr_EC#!#VA_Pr4#!#VA_Me_EC#!#VA_Mfinal"
			nm_bd=nm_vars
			vars_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			notas_a_lancar=9					
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
	end if	
	
	
	
dados_planilha_notas=tb&"#$#"&ln_pesos_cols&"#$#"&ln_pesos_vars&"#$#"&nm_pesos_vars&"#$#"&ln_nom_cols&"#$#"&nm_vars&"#$#"&nm_bd&"#$#"&vars_calc&"#$#"&action&"#$#"&notas_a_lancar&"#$#"&gera_pdf	
end function	
	
	
Function dados_boletim_avaliacao(ano_letivo,unidade,curso,etapa,turma,disciplina_mae,disciplina_filha,periodo,outro)




	if opcao="A" then		
			tb="TB_NOTA_A"
			ln_bol_av_cols="Disciplina#!#FAL#!#AV1#!#AV2#!#AV3#!#AV4#!#AV5#!#NPA#!#NAV#!#PR#!#M1#!#Bon#!#M2#!#REC#!#M3#!#Alterado por#!#Data/Hora#!!#N#!#P#!#N#!#P#!#N#!#P#!#N#!#P#!#N#!#P"
			ln_bol_av_span="ROWSPAN#!#ROWSPAN#!#COLSPAN#!#COLSPAN#!#COLSPAN#!#COLSPAN#!#COLSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN"
			nm_bol_av_vars="faltas#!#av1#!#p_av1#!#av2#!#p_av2#!#av3#!#p_av3#!#av4#!#p_av4#!#av5#!#p_av5#!#npa#!#nav#!#pr#!#media1#!#bon#!#media2#!#rec#!#media3"
			ln_bol_av_vars="NU_Faltas#!#VA_AV1#!#VA_MAX_AV1#!#VA_AV2#!#VA_MAX_AV2#!#VA_AV3#!#VA_MAX_AV3#!#VA_AV4#!#VA_MAX_AV4#!#VA_AV5#!#VA_MAX_AV5#!#CALCULADO#!#VA_Avaliacao#!#VA_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"	

			vars_bol_av_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#CALC1#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#P#!#M#!#M#!#M#!#M#!#M"
			legenda_bol_av="FAL - Faltas, AV - Avaliação, N - Nota, P - Peso, NPA - Nota Parcial de Avaliações, NAV - Nota das Avaliações, Pr - Prova, M - Médias, Bon - Bônus, Rec - Recuperação"	

	elseif opcao="B" then	
			tb="TB_NOTA_B"
			ln_bol_av_cols="Disciplina#!#FAL#!#AV1#!#AV2#!#AV3#!#AV4#!#AV5#!#NPA#!#NAV#!#PR#!#M1#!#Bon#!#M2#!#REC#!#M3#!#Alterado por#!#Data/Hora#!!#N#!#P#!#N#!#P#!#N#!#P#!#N#!#P#!#N#!#P"
			ln_bol_av_span="ROWSPAN#!#ROWSPAN#!#COLSPAN#!#COLSPAN#!#COLSPAN#!#COLSPAN#!#COLSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN"
			nm_bol_av_vars="faltas#!#av1#!#p_av1#!#av2#!#p_av2#!#av3#!#p_av3#!#av4#!#p_av4#!#av5#!#p_av5#!#npa#!#nav#!#pr#!#media1#!#bon#!#media2#!#rec#!#media3"
			ln_bol_av_vars="NU_Faltas#!#VA_AV1#!#VA_MAX_AV1#!#VA_AV2#!#VA_MAX_AV2#!#VA_AV3#!#VA_MAX_AV3#!#VA_AV4#!#VA_MAX_AV4#!#VA_AV5#!#VA_MAX_AV5#!#CALCULADO#!#VA_Avaliacao#!#VA_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"	

			vars_bol_av_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#CALC1#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#P#!#M#!#M#!#M#!#M#!#M"
			legenda_bol_av="FAL - Faltas, AV - Avaliação, N - Nota, P - Peso, NPA - Nota Parcial de Avaliações, NAV - Nota das Avaliações, Pr - Prova, M - Médias, Bon - Bônus, Rec - Recuperação"					

	elseif opcao="C" then
			tb="TB_NOTA_C"
			ln_bol_av_cols="Disciplina#!#FAL#!#AV1#!#AV2#!#AV3#!#AV4#!#AV5#!#NPA#!#NAV#!#PR#!#M1#!#Bon#!#M2#!#REC#!#M3#!#Alterado por#!#Data/Hora#!!#N#!#P#!#N#!#P#!#N#!#P#!#N#!#P#!#N#!#P"
			ln_bol_av_span="ROWSPAN#!#ROWSPAN#!#COLSPAN#!#COLSPAN#!#COLSPAN#!#COLSPAN#!#COLSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN"
			nm_bol_av_vars="faltas#!#av1#!#p_av1#!#av2#!#p_av2#!#av3#!#p_av3#!#av4#!#p_av4#!#av5#!#p_av5#!#npa#!#nav#!#pr#!#media1#!#bon#!#media2#!#rec#!#media3"
			ln_bol_av_vars="NU_Faltas#!#VA_AV1#!#VA_MAX_AV1#!#VA_AV2#!#VA_MAX_AV2#!#VA_AV3#!#VA_MAX_AV3#!#VA_AV4#!#VA_MAX_AV4#!#VA_AV5#!#VA_MAX_AV5#!#CALCULADO#!#VA_Avaliacao#!#VA_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"	

			vars_bol_av_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#CALC1#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
			exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#P#!#M#!#M#!#M#!#M#!#M"
			legenda_bol_av="FAL - Faltas, AV - Avaliação, N - Nota, P - Peso, NPA - Nota Parcial de Avaliações, NAV - Nota das Avaliações, Pr - Prova, M - Médias, Bon - Bônus, Rec - Recuperação"

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
end function
	



'Boletim de avaliação				
'Atenção na escrita de Alterado por e Data/Hora, pois escritas diferentes impactam a função Boletim de Avaliação	
'			ln_bol_av_cols="Disciplina#!#FAL#!#AV1#!#AV2#!#AV3#!#AV4#!#AV5#!#NPA#!#NAV#!#PR#!#M1#!#Bon#!#M2#!#REC#!#M3#!#Alterado por#!#Data/Hora#!!#N#!#P#!#N#!#P#!#N#!#P#!#N#!#P#!#N#!#P"
'			ln_bol_av_span="ROWSPAN#!#ROWSPAN#!#COLSPAN#!#COLSPAN#!#COLSPAN#!#COLSPAN#!#COLSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN#!#ROWSPAN"
'			nm_bol_av_vars="faltas#!#av1#!#p_av1#!#av2#!#p_av2#!#av3#!#p_av3#!#av4#!#p_av4#!#av5#!#p_av5#!#npa#!#nav#!#pr#!#media1#!#bon#!#media2#!#rec#!#media3"
'			ln_bol_av_vars="NU_Faltas#!#VA_AV1#!#VA_MAX_AV1#!#VA_AV2#!#VA_MAX_AV2#!#VA_AV3#!#VA_MAX_AV3#!#VA_AV4#!#VA_MAX_AV4#!#VA_AV5#!#VA_MAX_AV5#!#CALCULADO#!#VA_Avaliacao#!#VA_Prova#!#VA_Media1#!#VA_Bonus#!#VA_Media2#!#VA_Rec#!#VA_Media3"	
'
'			vars_bol_av_calc="0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#0#!#CALC1#!#0#!#0#!#0#!#0#!#0#!#0#!#0"
'			exibe_apr_pr_bol_av="0#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#A#!#P#!#M#!#M#!#M#!#M#!#M"
'			legenda_bol_av="FAL - Faltas, AV - Avaliação, N - Nota, P - Peso, NPA - Nota Parcial de Avaliações, NAV - Nota das Avaliações, Pr - Prova, M - Médias, Bon - Bônus, Rec - Recuperação"
%>