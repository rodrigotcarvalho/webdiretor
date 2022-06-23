<!--#include file="bd_grade.asp"-->
<%
Function Calc_Prim_Media (unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, conexao, tb_nota, prd_prim_media, tipo_calculo, aproxima, outro)	

	media_acumulada=0
	peso_periodo_acumulado=0
	calcula_media="s"	
	
	arred_md = parametros_gerais(unidade, curso, co_etapa, turma, codigo_materia_pr,"arred_md",0)
	decimais_md = parametros_gerais(unidade, curso, co_etapa, turma, codigo_materia_pr,"decimais_md",0)	

	tp_materia=tipo_materia(codigo_materia, curso, co_etapa)
	nome_nota=var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo,tp_dados)	
	
	vetor_materia_filhas=busca_materias_filhas(codigo_materia)		
		
	disc_obrigat=Disciplina_Obrigatoria(codigo_materia,CON0,outro)	
				
	prd_prim_media=prd_prim_media*1
	
	Set RSp = Server.CreateObject("ADODB.Recordset")
	SQLp = "SELECT SUM(NU_Peso) as SUM_PESO FROM TB_Periodo where NU_Periodo <="& prd_prim_media
	RSp.Open SQLp, CON0	
	
	if RSp.EOF then
		peso_per=1
	else
		peso_per=RSp("SUM_PESO")			
	end if		
	
	soma_media = Calcula_Soma(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, tp_materia, conexao, tb_nota, prd_prim_media, "N", "CPM", outro)

	if soma_media = "&nbsp;" then
		calcula_media="n"
	else
		calcula_media="s"		
	end if	
	
	if calcula_media="n" then
		primeira_media="&nbsp;"
		resultado= "&nbsp;"		
	else	
		primeira_media=arredonda(soma_media/peso_per,arred_md,decimais_md,0)
		primeira_media = AcrescentaBonusMediaAnual(cod_aluno, codigo_materia, primeira_media)
		if aproxima="S" then
			if primeira_media >67 and primeira_media <70 then
				primeira_media =70
			end if			
		end if			
		resultado= Apura_Resultado(unidade, curso, co_etapa, turma, codigo_materia, primeira_media, "R1", 0)		
	end if	
	

Calc_Prim_Media=primeira_media&"#!#"&resultado
end Function





Function Calc_Seg_Media (unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, conexao, tb_nota, prd_seg_media, tipo_calculo, compara, aproxima, outro)
	
	media_acumulada=0	

	nome_nota=var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo,tp_dados)	
	tp_materia=tipo_materia(codigo_materia, curso, co_etapa)	
	vetor_materia_filhas=busca_materias_filhas(codigo_materia)	
	arred_md = parametros_gerais(unidade, curso, co_etapa, turma, codigo_materia_pr,"arred_md",0)
	decimais_md = parametros_gerais(unidade, curso, co_etapa, turma, codigo_materia_pr,"decimais_md",0)		
	aproxima_m1 = parametros_gerais(unidade, curso, co_etapa, turma, codigo_materia_pr,"aproxima_m1",0)		
	
	prd_prim_media=Periodo_Media(tp_modelo,"MA",outro)
	
	primeira_media = Calc_Prim_Media (unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, conexao, tb_nota,prd_prim_media, tipo_calculo, aproxima_m1, outro)	
	
	inf_primeira_media=split(primeira_media,"#!#")
	prim_resultado=inf_primeira_media(1)

	if prim_resultado = "APR" or prim_resultado="&nbsp;" then
		segunda_media=inf_primeira_media(0)
		resultado=prim_resultado
	else
		pesos_media=Peso_Calc_Media(unidade, curso, co_etapa, turma, disciplina, "M2", outro)
		
		vetor_peso_medias=split(pesos_media,"#!#")
		peso_m2_m1=vetor_peso_medias(0)*1
		peso_m2_m2=vetor_peso_medias(1)*1
		peso_periodo_acumulado=peso_m2_m1+peso_m2_m2	
				
		var_bd=var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo,"BDM")		
		vetor_materia=busca_materias_filhas(codigo_materia)		
		
		if tp_materia="T_F_F_N" or tp_materia="F_T_F_N"	 or tp_materia="F_F_T_N" then
			media_periodo=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, vetor_materia, conexao, tb_nota, prd_seg_media, var_bd, outro)
			
		elseif tp_materia="T_T_F_N" then
			media_periodo=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, vetor_materia, conexao, tb_nota, prd_seg_media, var_bd, outro)		
				
		elseif tp_materia="T_F_T_N" then	
			media_periodo=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, vetor_materia, conexao, tb_nota, prd_seg_media, var_bd, outro)		
						
		end if		

		if  media_periodo="&nbsp;" or isnull(media_periodo) then
			teste_ano=verifica_ano_letivo(session("ano_letivo"),variavel_2,variavel_3,variavel_4,variavel_5,CON,"con", detalhe_busca)			
			if teste_ano="B" then
				media_acumulada=inf_primeira_media(0)*peso_periodo_acumulado
			else
				sem_media="s"	
			end if			
		else
			if inf_primeira_media(0)="&nbsp;" then
				sem_media="s"
			else
				inf_primeira_media(0)=inf_primeira_media(0)*1
				media_periodo=media_periodo*1
				if compara_m2 = "S" then				
					if media_periodo> inf_primeira_media(0) then
						media_acumulada=(inf_primeira_media(0)*peso_m2_m1)+(media_periodo*peso_m2_m2)					
					else
						media_acumulada=inf_primeira_media(0)
						peso_periodo_acumulado=1
					end if	
				else
					media_acumulada=(inf_primeira_media(0)*peso_m2_m1)+(media_periodo*peso_m2_m2)				
				end if	
			end if	
		end if	
	
		disc_obrigat=disciplina_obrigatoria(codigo_materia,CON0,outro)	
					
		if 	(media_acumulada=0 and disc_obrigat="N") or sem_media="s" then
			segunda_media="&nbsp;"
			resultado="&nbsp;"
		else
			segunda_media=arredonda(media_acumulada/peso_periodo_acumulado,arred_md,decimais_md,0)
			if aproxima="S" then	
				if segunda_media >67 and segunda_media <70 then
					segunda_media =70
				end if			
			end if				
			resultado= Apura_Resultado(unidade, curso, co_etapa, turma, codigo_materia, segunda_media, "R2", outro)				
		end if	
	end if

Calc_Seg_Media=segunda_media&"#!#"&resultado
end Function






Function Calc_Ter_Media (unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, conexao, tb_nota, prd_ter_media, tipo_calculo, compara, aproxima, outro)

	media_acumulada=0
	
	nome_nota=var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo,tp_dados)	
	tp_materia=tipo_materia(codigo_materia, curso, co_etapa)	
	vetor_materia_filhas=busca_materias_filhas(codigo_materia)	
						
	arred_md = parametros_gerais(unidade, curso, co_etapa, turma, codigo_materia_pr,"arred_md",0)
	decimais_md = parametros_gerais(unidade, curso, co_etapa, turma, codigo_materia_pr,"decimais_md",0)		
	aproxima_m1 = parametros_gerais(unidade, curso, co_etapa, turma, codigo_materia_pr,"aproxima_m1",0)	
	aproxima_m2 = parametros_gerais(unidade, curso, co_etapa, turma, codigo_materia_pr,"aproxima_m2",0)	
	compara_m2 = parametros_gerais(unidade, curso, co_etapa, turma, codigo_materia_pr,"compara_m2",0)		
		
	prd_prim_media=Periodo_Media(tp_modelo,"MA",outro)
	prd_seg_media=Periodo_Media(tp_modelo,"RF",outro)	
	primeira_media = Calc_Prim_Media (unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, conexao, tb_nota, prd_prim_media, tipo_calculo, aproxima_m1, outro)

	inf_primeira_media=split(primeira_media,"#!#")
	prim_resultado=inf_primeira_media(1)

	if tipo_calculo = "ATA" then
		if prim_resultado<>"&nbsp;" then
			terceira_media=inf_primeira_media(0)
			resultado=prim_resultado

			segunda_media = Calc_Seg_Media (unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, conexao, tb_nota, prd_seg_media, tipo_calculo, compara_m2, aproxima_m2, outro)	
		
			inf_segunda_media=split(segunda_media,"#!#")
			seg_resultado=inf_segunda_media(1)

'if cod_aluno=31408 then
'response.Write(prim_resultado&"-"&codigo_materia&"-tm "&terceira_media&"<br />")		
'end if	

			if seg_resultado<>"&nbsp;" then
				terceira_media=inf_segunda_media(0)
				resultado=seg_resultado
			end if		

'if cod_aluno=31408 then
'response.Write(seg_resultado&"-"&codigo_materia&"-tm "&terceira_media&"<br />")		
'end if			
			var_bd=var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo,"BDM")		
			vetor_materia=busca_materias_filhas(codigo_materia)		
		
			if tp_materia="T_F_F_N" or tp_materia="F_T_F_N"	 or tp_materia="F_F_T_N" then
				media_periodo=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, vetor_materia, conexao, tb_nota, prd_ter_media, var_bd, outro)
				
			elseif tp_materia="T_T_F_N" then
				media_periodo=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, vetor_materia, conexao, tb_nota, prd_ter_media, var_bd, outro)		
					
			elseif tp_materia="T_F_T_N" then	
				media_periodo=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, vetor_materia, conexao, tb_nota, prd_ter_media, var_bd, outro)		
							
			end if		
			

			pesos_media=Peso_Calc_Media(unidade, curso, co_etapa, turma, disciplina, "M3", outro)

			vetor_peso_medias=split(pesos_media,"#!#")
			peso_m3_m1=vetor_peso_medias(0)*1
			peso_m3_m2=vetor_peso_medias(1)*1
			peso_m3_m3=vetor_peso_medias(2)*1		
			peso_periodo_acumulado=peso_m3_m1+peso_m3_m2+peso_m3_m3	

			if media_periodo <> "&nbsp;" then
				if inf_primeira_media(0)="&nbsp;" or inf_segunda_media(0)="&nbsp;" then
					sem_media="s"
				else
					calcula_m3 = "s"	
					inf_segunda_media(0)=inf_segunda_media(0)*1
					media_periodo=media_periodo*1
					if compara = "S" then
						if media_periodo> inf_segunda_media(0) then
							media_acumulada=(inf_primeira_media(0)*peso_m3_m1)+(inf_segunda_media(0)*peso_m3_m2)+(media_periodo*peso_m3_m3)
						else
							media_acumulada=inf_segunda_media(0)
							peso_periodo_acumulado=1
						end if
					else
						media_acumulada=(inf_primeira_media(0)*peso_m3_m1)+(inf_segunda_media(0)*peso_m3_m2)+(media_periodo*peso_m3_m3)					
					end if						
				end if	

				disc_obrigat=disciplina_obrigatoria(codigo_materia,CON0,outro)						
	
				if 	(media_acumulada=0 and disc_obrigat="N") or sem_media="s" then
					terceira_media="&nbsp;"
				else
					if calcula_m3 = "s"	then
						terceira_media=arredonda(media_acumulada/peso_periodo_acumulado,arred_md,0,0)
						if isnumeric(inf_segunda_media(0)) then
							inf_segunda_media(0)=inf_segunda_media(0)*1
							terceira_media=terceira_media*1	
							if compara="S" then			
								if inf_segunda_media(0)>terceira_media then
									terceira_media=inf_segunda_media(0)
								else
									terceira_media=formatnumber(terceira_media,0)
								end if
							else
								terceira_media=formatnumber(terceira_media,0)						
							end if	
						else
							terceira_media=formatnumber(terceira_media,0)				
						end if	
						if aproxima="S" then	
							if terceira_media >67 and terceira_media <70 then
								terceira_media =70
							end if			
						end if					
						resultado= Apura_Resultado(unidade, curso, co_etapa, turma, codigo_materia, terceira_media, "R3", outro)
					end if
				end if
			end if	
		end if								
	else
	
	
		if prim_resultado = "APR" or prim_resultado = "REP" or prim_resultado="&nbsp;" then
			terceira_media=inf_primeira_media(0)
			resultado=prim_resultado
		else
		
			segunda_media = Calc_Seg_Media (unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, conexao, tb_nota, prd_seg_media, tipo_calculo, compara_m2, aproxima_m2, outro)	
		
			inf_segunda_media=split(segunda_media,"#!#")
			seg_resultado=inf_segunda_media(1)
	'	response.Write(cod_aluno&"-"&codigo_materia&"-"&segunda_media&"<br />")	
			if seg_resultado = "APR" or seg_resultado = "REP" or seg_resultado="&nbsp;" then
				terceira_media=inf_segunda_media(0)
				resultado=seg_resultado
			else			
	
				var_bd=var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo,"BDM")		
				vetor_materia=busca_materias_filhas(codigo_materia)		

				if tp_materia="T_F_F_N" or tp_materia="F_T_F_N"	 or tp_materia="F_F_T_N" then
					media_periodo=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, vetor_materia, conexao, tb_nota, prd_ter_media, var_bd, outro)
					
				elseif tp_materia="T_T_F_N" then
					media_periodo=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, vetor_materia, conexao, tb_nota, prd_ter_media, var_bd, outro)		
						
				elseif tp_materia="T_F_T_N" then	
					media_periodo=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, vetor_materia, conexao, tb_nota, prd_ter_media, var_bd, outro)		
								
				end if		
				
				pesos_media=Peso_Calc_Media(unidade, curso, co_etapa, turma, disciplina, "M3", outro)
	
				vetor_peso_medias=split(pesos_media,"#!#")
				peso_m3_m1=vetor_peso_medias(0)*1
				peso_m3_m2=vetor_peso_medias(1)*1
				peso_m3_m3=vetor_peso_medias(2)*1		
				peso_periodo_acumulado=peso_m3_m1+peso_m3_m2+peso_m3_m3	
	'	response.Write(cod_aluno&"-"&codigo_materia&"-"&media_periodo&"<br />")		
				if media_periodo="&nbsp;" or isnull(media_periodo) then
					teste_ano=verifica_ano_letivo(session("ano_letivo"),variavel_2,variavel_3,variavel_4,variavel_5,CON,"con", detalhe_busca)				
					if teste_ano="B" then
						media_acumulada=inf_segunda_media(0)
						calcula_m3 = "s"
					else
						if tipo_calculo = "boletim" then
							terceira_media=inf_segunda_media(0)
							resultado=seg_resultado
							calcula_m3 = "n"						
						else	
							sem_media="s"					
						end if	
					end if	
				else
					if inf_primeira_media(0)="&nbsp;" or inf_segunda_media(0)="&nbsp;" then
						sem_media="s"
					else
						calcula_m3 = "s"	
						inf_segunda_media(0)=inf_segunda_media(0)*1
						media_periodo=media_periodo*1
						if compara = "S" then
							if media_periodo> inf_segunda_media(0) then
								media_acumulada=(inf_primeira_media(0)*peso_m3_m1)+(inf_segunda_media(0)*peso_m3_m2)+(media_periodo*peso_m3_m3)
							else
								media_acumulada=inf_segunda_media(0)
								peso_periodo_acumulado=1
							end if
						else
							media_acumulada=(inf_primeira_media(0)*peso_m3_m1)+(inf_segunda_media(0)*peso_m3_m2)+(media_periodo*peso_m3_m3)					
						end if						
					end if	
				end if	
		'	response.Write(cod_aluno&"-"&codigo_materia&"-"&inf_primeira_media(0)&"*"&peso_m3_m1&")+("&inf_segunda_media(0)&"*"&peso_m3_m2&")+("&media_periodo&"*"&peso_m3_m3&"<br />")			
				disc_obrigat=disciplina_obrigatoria(codigo_materia,CON0,outro)						
	'	response.Write(cod_aluno&"-"&codigo_materia&"-"&media_acumulada&"<br />")						
				if 	(media_acumulada=0 and disc_obrigat="N") or sem_media="s" then
					terceira_media="&nbsp;"
				else
					if calcula_m3 = "s"	then
						terceira_media=arredonda(media_acumulada/peso_periodo_acumulado,arred_md,0,0)
						if isnumeric(inf_segunda_media(0)) then
							inf_segunda_media(0)=inf_segunda_media(0)*1
							terceira_media=terceira_media*1	
							if compara="S" then			
								if inf_segunda_media(0)>terceira_media then
									terceira_media=inf_segunda_media(0)
								else
									terceira_media=formatnumber(terceira_media,0)
								end if
							else
								terceira_media=formatnumber(terceira_media,0)						
							end if	
						else
							terceira_media=formatnumber(terceira_media,0)				
						end if	
						if aproxima="S" then	
							if terceira_media >67 and terceira_media <70 then
								terceira_media =70
							end if			
						end if					
						resultado= Apura_Resultado(unidade, curso, co_etapa, turma, codigo_materia, terceira_media, "R3", outro)
					end if						
				end if
			end if		
		end if
	end if	
'	response.Write(cod_aluno&"-"&codigo_materia&"-tm "&terceira_media&"<br />")		

		
'if cod_aluno=31408 then
'	response.end()
'end if		
Calc_Ter_Media=terceira_media&"#!#"&resultado

end Function

Function Peso_Calc_Media(unidade, curso, etapa, turma, disciplina, media_calc, outro)

if media_calc="M2" then
	sql_media="NU_Peso_Media_M2_M1, NU_Peso_Media_M2_M2"
elseif media_calc="M3" then	
	sql_media="NU_Peso_Media_M3_M1, NU_Peso_Media_M3_M2, NU_Peso_Media_M3_M3"
end if	

	Set RSra = Server.CreateObject("ADODB.Recordset")
	SQLra = "SELECT "&sql_media&" FROM TB_Regras_Aprovacao where CO_Curso = '"&curso&"' and CO_Etapa = '"&etapa&"'"
	RSra.Open SQLra, CON0	
			
	if RSra.EOF then
		peso_m1=1
		peso_m2=1
	else
		if media_calc="M2" then
			peso_m1=RSra("NU_Peso_Media_M2_M1")
			peso_m2=RSra("NU_Peso_Media_M2_M2")
			peso_m3="&nbsp;"
			if isnull(peso_m1) then
				peso_m1=1		
			end if
			if isnull(peso_m2) then
				peso_m2=1		
			end if	
		elseif media_calc="M3" then
			peso_m1=RSra("NU_Peso_Media_M3_M1")
			peso_m2=RSra("NU_Peso_Media_M3_M2")
			peso_m3=RSra("NU_Peso_Media_M3_M3")	
			if isnull(peso_m1) then
				peso_m1=1		
			end if
			if isnull(peso_m2) then
				peso_m2=1		
			end if		
			if isnull(peso_m3) then
				peso_m3=1		
			end if										
		end if		
	end if	
Peso_Calc_Media=peso_m1&"#!#"&peso_m2&"#!#"&peso_m3
end function	

Function Apura_Resultado(unidade, curso, co_etapa, turma, codigo_materia, valor, tipo_resultado, outro)	
	
	Set RSra = Server.CreateObject("ADODB.Recordset")
	SQLra = "SELECT * FROM TB_Regras_Aprovacao where CO_Curso = '"&curso&"' and CO_Etapa = '"&co_etapa&"'"
	RSra.Open SQLra, CON0	
			
	valor_m1=RSra("NU_Valor_M1")
	m1_menor=RSra("NU_Int_Me_Ma_Igual_M1")
	m1_maior_igual=RSra("NU_Int_Me_Me_M1")
	res1_3=RSra("NO_Expr_Ma_Igual_M1")
	res1_2=RSra("NO_Expr_Int_M1_V")
	res1_1=RSra("NO_Expr_Int_M1_F")
	peso_m2_m1=RSra("NU_Peso_Media_M2_M1")
	peso_m2_m2=RSra("NU_Peso_Media_M2_M2")
	
	valor_m2=RSra("NU_Valor_M2")
	m2_menor=RSra("NU_Int_Me_Ma_Igual_M2")
	m2_maior_igual=RSra("NU_Int_Me_Me_M2")	
	res2_3=RSra("NO_Expr_Ma_Igual_M2")
	res2_2=RSra("NO_Expr_Int_M2_V")
	res2_1=RSra("NO_Expr_Int_M2_F")
	peso_m3_m1=RSra("NU_Peso_Media_M3_M1")
	peso_m3_m2=RSra("NU_Peso_Media_M3_M2")
	peso_m3_m3=RSra("NU_Peso_Media_M3_M3")
	
	valor_m3=RSra("NU_Valor_M3")
	'm3_menor=RSra("NU_Int_Me_Ma_Igual_M3")
	m3_maior_igual=RSra("NU_Int_Me_Me_M3")	
	res3_1=RSra("NO_Expr_Int_M3_V")
	res3_2=RSra("NO_Expr_Ma_Igual_M3")

		
	valor=valor*1	
	m1_menor=m1_menor*1
	m1_maior_igual=m1_maior_igual*1
	m1_menor=m1_menor*1
	m2_maior_igual=m2_maior_igual*1
	m2_menor=m2_menor*1
	m3_maior_igual=m3_maior_igual*1
	
	teste_ano=verifica_ano_letivo(session("ano_letivo"),variavel_2,variavel_3,variavel_4,variavel_5,CON,"con", detalhe_busca)	

	if tipo_resultado="R1" then
		if valor < m1_menor then
			resultado=res1_1		
		elseif valor < m1_maior_igual then
				resultado=res1_2
		elseif valor >= m1_maior_igual then
			resultado=res1_3
		elseif teste_ano="B" then
			resultado=res1_1			
		else
			if valor >= m1_menor then
				resultado=res1_2
			else
				resultado=res1_1	
			end if
		end if	
	elseif tipo_resultado="R2" then	
		if valor >= m2_maior_igual then
			resultado=res2_3
		elseif teste_ano="B" then
			resultado=res2_1			
		else
			if valor >= m2_menor then
				resultado=res2_2
			else
				resultado=res2_1	
			end if
		end if	
	elseif tipo_resultado="R3" then	
		if valor >= m3_maior_igual then
			resultado=res3_2
		else
			resultado=res3_1	
		end if	
	end if
Apura_Resultado=resultado	
end Function

Function novo2_apura_resultado_aluno(curso, etapa, cod_cons, vetor_materia, vetor_medias, frequencia, periodo_m1,  periodo_m2, periodo_m3, tipo_apuracao, mostra_res_prim_media, mostra_res_seg_media, outro)
'mostra_res_prim_media = parametro que indica que deve-se mostrar ou não disciplinas com resultado diferente de APR ou REP na primeira média
'mostra_res_seg_media = parametro que indica que deve-se mostrar ou não disciplinas com resultado diferente de APR ou REP na segunda média

'Valores de tipo_apuracao
'ata - Serve para a ata de resultados e para as declarações


	teste_ano=verifica_ano_letivo(session("ano_letivo"),variavel_2,variavel_3,variavel_4,variavel_5,CON,"con", detalhe_busca)
	
	Set RSra = Server.CreateObject("ADODB.Recordset")
	SQLra = "SELECT * FROM TB_Regras_Aprovacao where CO_Curso = '"&curso&"' and CO_Etapa = '"&etapa&"'"
	RSra.Open SQLra, CON0	
			

	qtd_max_rec=RSra("NU_Disc_Ult_Periodo")
	qtd_max_dep=RSra("NU_Qt_Dis_Dep")	
	res_pfi=RSra("NO_Expr_Int_M1_V")
	res_rec=RSra("NO_Expr_Int_M2_V")	
	res_apr=RSra("NO_Expr_Maior_Igual_VL_Abr")
	res_dep=RSra("NO_Expr_Cond_Verdade_Abr")
	res_rep=RSra("NO_Expr_Cond_Falso_Abr")
	maximo_faltas=RSra("NU_Per_Aprovacao")
	qtd_rec=0
	qtd_dep=0
	if isnumeric(frequencia) then
		frequencia=frequencia*1
	end if
	if isnumeric(maximo_faltas) then	
		maximo_faltas=maximo_faltas*1
	end if	
	if frequencia<maximo_faltas then
		libera_resultado="s"
		result_temp="REP"	
	else
		
		if isnull(qtd_max_rec) or qtd_max_rec="" then
			qtd_max_rec=0
		end if
'		response.Write(vetor_medias&"<BR>")
'		response.Write(vetor_materia)
		resultados_materia = split(vetor_medias, "#$#" )
		vetor_disc_cntrle = split(vetor_materia, "#!#" )	
		libera_resultado="s"
		result_temp	= "APR"
		for rm=0 to ubound(resultados_materia)	
			nota_materia = split(resultados_materia(rm), "#!#" )
			res_aluno=nota_materia(1)
				
			if result_temp="REP" then
				libera_resultado="s"	
			else
				if res_aluno="" or isnull(res_aluno) or res_aluno="&nbsp;" or res_aluno=" " then
					disc_obrigat=disciplina_obrigatoria(vetor_disc_cntrle(rm),CON0,outro)	
					if disc_obrigat="S" then
						libera_resultado="n"							
					end if	
				else

					if res_aluno = "APR" then
					else
						result_temp=res_aluno
						if res_aluno = "REP" then
							tipo_media = "MF"							
							modifica_result = Verifica_Conselho_Classe(cod_cons, vetor_disc_cntrle(rm), tipo_media, outro)
							if modifica_result <> "N" then
								result_temp = modifica_result
							else
								qtd_dep=qtd_dep+1								
							end if								
						elseif res_aluno = "REC" then	
							tipo_media = "RF"
							modifica_result = Verifica_Conselho_Classe(cod_cons, vetor_disc_cntrle(rm), tipo_media, outro)
							if modifica_result <> "N" then
								result_temp = modifica_result
							else
								qtd_rec=qtd_rec+1																
							end if								
						elseif res_aluno = "PFI" then
							tipo_media = "MA"
							modifica_result = Verifica_Conselho_Classe(cod_cons, vetor_disc_cntrle(rm), tipo_media, outro)
							if modifica_result <> "N" then
								result_temp = modifica_result
							else
								qtd_pfi=qtd_pfi+1									
							end if																														
						end if

					
					end if	
				end if
			end if	
		Next
	end if			
	if 	libera_resultado="s" then
			if result_temp="REP" then
				resultado_aluno=res_rep	
			else			
				if qtd_dep>qtd_max_dep then
					resultado_aluno=res_rep	
				elseif qtd_dep>0  then
					resultado_aluno=res_dep					
				elseif qtd_rec>qtd_max_rec  then	
					resultado_aluno=res_rep	
				elseif qtd_rec>0  then
					if mostra_res_seg_media = "S" then
						resultado_aluno=res_rec	
					else
						resultado_aluno="&nbsp;"	
					end if						
				elseif qtd_pfi>0  then
					if mostra_res_prim_media = "S" then
						resultado_aluno=res_pfi
					else
						resultado_aluno="&nbsp;"	
					end if					
				else
					resultado_aluno=res_apr										
				end if
			end if	
		novo2_apura_resultado_aluno=resultado_aluno
	else
		novo2_apura_resultado_aluno="&nbsp;"		
	end if	

'response.End()
end function
%>