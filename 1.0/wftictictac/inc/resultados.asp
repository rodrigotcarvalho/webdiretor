<%
Function Calc_Prim_Media (unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, prd_prim_media, tipo_calculo, outro)
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CONn = Server.CreateObject("ADODB.Connection") 
	ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONn.Open ABRIRn	

	media_acumulada=0
	peso_periodo_acumulado=0
	calcula_media="s"
	
	tp_model=tipo_divisao_ano(curso,co_etapa,"tp_modelo")
	tp_freq=tipo_divisao_ano(curso,co_etapa,"in_frequencia")		
	nome_nota=var_bd_periodo(tp_model,tp_freq,tb_nota,prd_prim_media,"BDM")
	tp_materia=tipo_materia(codigo_materia, curso, co_etapa)	

		
	Set RS3a = Server.CreateObject("ADODB.Recordset")
	SQL3a = "SELECT * FROM TB_Materia where CO_Materia ='"& codigo_materia &"' order by NU_Ordem_Boletim"
	RS3a.Open SQL3a, CON0	
	if RS3a.EOF then
		disc_obrigat="s"
	else
		ind_obr= RS3a("IN_Obrigatorio")	
		
		if ind_obr=TRUE then
			disc_obrigat="s"
		else
			disc_obrigat="n"
		end if			
	end if				

'	for cons_per=1 to prd_prim_media
	
		Set RSp = Server.CreateObject("ADODB.Recordset")
		SQLp = "SELECT sum(NU_Peso) as peso_acumulado FROM TB_Periodo where TP_Modelo='"&tp_model&"' AND NU_Periodo <="& prd_prim_media
		RSp.Open SQLp, CON0	
	
'		if RSp.EOF then
'			peso_per=1
'		else
'			peso_per=RSp("NU_Peso")			
'		end if		
		peso_periodo_acumulado=peso_periodo_acumulado+peso_per
		peso_periodo_acumulado=RSp("peso_acumulado")		
'		if tp_materia="T_F_F_N" then		
'			codigo_materia_pr=busca_materia_mae(codigo_materia)		
'			media_periodo=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia_pr, codigo_materia, CONn , tb_nota, cons_per, nome_nota, outro)
'		elseif tp_materia="T_T_F_N" then	
'			vetor_filhas_T_T_F_N=busca_materias_filhas(codigo_materia)
'			 media_periodo=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, vetor_filhas_T_T_F_N, CONn, tb_nota, cons_per, nome_nota, outro)					
'		elseif tp_materia="T_F_T_N" then
'			vetor_filhas_T_F_T_N=busca_materias_filhas(codigo_materia)		
'			media_periodo=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, vetor_filhas_T_F_T_N, CONn, tb_nota, cons_per, nome_nota, outro)	
'		end if	
'		if  media_periodo="&nbsp;" or isnull(media_periodo) then
'			calcula_media="n"
'		else	
'			media_acumulada=media_acumulada+(media_periodo*peso_per)
'		end if		
'	Next
	media_acumulada=Calcula_Soma(tp_model, tp_freq, unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, CONn, tp_materia, tb_nota, prd_prim_media, outro)	
	if calcula_media<>"s" then
'		if 	media_acumulada=0 and disc_obrigat="n" then
			primeira_media="&nbsp;"
			resultado= "&nbsp;"		
	else
		if  media_acumulada="&nbsp;" or isnull(media_acumulada) then
			primeira_media="&nbsp;"
			resultado= "&nbsp;"	
		else
			primeira_media=arredonda(media_acumulada/peso_periodo_acumulado,"mat_dez",1,0)
			resultado= Apura_Resultado(unidade, curso, co_etapa, turma, codigo_materia, primeira_media, "R1", outro)	
		end if		
	end if	

Calc_Prim_Media=primeira_media&"#!#"&resultado
end Function





Function Calc_Seg_Media (unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, prd_seg_media, tipo_calculo, outro)
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CONn = Server.CreateObject("ADODB.Connection") 
	ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONn.Open ABRIRn	
	
	media_acumulada=0	
	
	tp_model=tipo_divisao_ano(curso,co_etapa,"tp_modelo")
	tp_freq=tipo_divisao_ano(curso,co_etapa,"in_frequencia")		
	nome_nota=var_bd_periodo(tp_model,tp_freq,tb_nota,prd_prim_media,"BDM")
	
	tp_materia=tipo_materia(codigo_materia, curso, co_etapa)	
	
	Set RS3a = Server.CreateObject("ADODB.Recordset")
	SQL3a = "SELECT * FROM TB_Materia where CO_Materia ='"& codigo_materia &"' order by NU_Ordem_Boletim"
	RS3a.Open SQL3a, CON0	
	if RS3a.EOF then
		disc_obrigat="s"
	else
		ind_obr= RS3a("IN_Obrigatorio")	
		
		if ind_obr=TRUE then
			disc_obrigat="s"
		else
			disc_obrigat="n"
		end if			
	end if			
	
	Set RSano = Server.CreateObject("ADODB.Recordset")
	SQLano = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&ano_letivo&"'"
	RSano.Open SQLano, CON

	teste_ano=RSano("ST_Ano_Letivo")		
	
	Set RSra = Server.CreateObject("ADODB.Recordset")
	SQLra = "SELECT NU_Peso_Media_M2_M1,NU_Peso_Media_M2_M2 FROM TB_Regras_Aprovacao where CO_Curso = '"&curso&"' and CO_Etapa = '"&etapa&"'"
	RSra.Open SQLra, CON0	
			
	if RSra.EOF then
		peso_m2_m1=1
		peso_m2_m2=1
	else
		peso_m2_m1=RSra("NU_Peso_Media_M2_M1")
		peso_m2_m2=RSra("NU_Peso_Media_M2_M2")
		
		if isnull(peso_m2_m1) then
			peso_m2_m1=1		
		end if
		
		if isnull(peso_m2_m2) then
			peso_m2_m2=1		
		end if		
	end if	
	peso_m2_m1=peso_m2_m1*1
	peso_m2_m2=peso_m2_m2*1	

	peso_periodo_acumulado=peso_m2_m1+peso_m2_m2	
	
	prd_prim_media=prd_seg_media-1
	
	primeira_media = Calc_Prim_Media (unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, prd_prim_media, tipo_calculo, outro)	
	inf_primeira_media=split(primeira_media,"#!#")
	prim_resultado=inf_primeira_media(1)
		
	if prim_resultado = "APR" or prim_resultado="&nbsp;" then
		'segunda_media=inf_primeira_media(0)
		segunda_media="&nbsp;"
		resultado=prim_resultado
	else	
		if tp_materia="T_F_F_N" then
			codigo_materia_pr=busca_materia_mae(codigo_materia)			
			media_periodo=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia_pr, codigo_materia, CONn , tb_nota, prd_seg_media, nome_nota, outro)
		elseif tp_materia="T_T_F_N" then	
			vetor_filhas_T_T_F_N=busca_materias_filhas(codigo_materia)
			 media_periodo=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, vetor_filhas_T_T_F_N, CONn, tb_nota, prd_seg_media, nome_nota, outro)					
		elseif tp_materia="T_F_T_N" then
			vetor_filhas_T_F_T_N=busca_materias_filhas(codigo_materia)		
			media_periodo=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, vetor_filhas_T_F_T_N, CONn, tb_nota, prd_seg_media, nome_nota, outro)								
		end if	


		if  media_periodo="&nbsp;" or isnull(media_periodo) then
			if teste_ano="B" then
				media_acumulada=inf_primeira_media(0)
			end if			
		else
			media_acumulada=(inf_primeira_media(0)*peso_m2_m1)+(media_periodo*peso_m2_m2)
		end if	
		
		if 	media_acumulada=0 and disc_obrigat="n" then
			segunda_media="&nbsp;"
			resultado="&nbsp;"		
		else
			if media_acumulada=0 then
				segunda_media="&nbsp;"			
				resultado=prim_resultado					
			else
				segunda_media=arredonda(media_acumulada/peso_periodo_acumulado,"mat_dez",1,0)
				resultado= Apura_Resultado(unidade, curso, co_etapa, turma, codigo_materia, segunda_media, "R2", outro)				
			end if	
		end if	
	end if	
Calc_Seg_Media=segunda_media&"#!#"&resultado
end Function






Function Calc_Ter_Media (unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, prd_ter_media, tipo_calculo, outro)

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CONn = Server.CreateObject("ADODB.Connection") 
	ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONn.Open ABRIRn	

	media_acumulada=0
	tp_model=tipo_divisao_ano(curso,co_etapa,"tp_modelo")
	tp_freq=tipo_divisao_ano(curso,co_etapa,"in_frequencia")		
	nome_nota=var_bd_periodo(tp_model,tp_freq,tb_nota,prd_prim_media,"BDM")
	tp_materia=tipo_materia(codigo_materia, curso, co_etapa)		
	
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& codigo_materia &"'"
	RS.Open SQL, CON0
	
	mae= RS("IN_MAE")
	fil= RS("IN_FIL")
	in_co= RS("IN_CO")
	peso= RS("NU_Peso")		
	
	Set RS3a = Server.CreateObject("ADODB.Recordset")
	SQL3a = "SELECT * FROM TB_Materia where CO_Materia ='"& codigo_materia &"' order by NU_Ordem_Boletim"
	RS3a.Open SQL3a, CON0	
	if RS3a.EOF then
		disc_obrigat="s"
	else
		ind_obr= RS3a("IN_Obrigatorio")	
		
		if ind_obr=TRUE then
			disc_obrigat="s"
		else
			disc_obrigat="n"
		end if			
	end if		
	
	Set RSano = Server.CreateObject("ADODB.Recordset")
	SQLano = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&ano_letivo&"'"
	RSano.Open SQLano, CON

	teste_ano=RSano("ST_Ano_Letivo")	
	
	Set RSra = Server.CreateObject("ADODB.Recordset")
	SQLra = "SELECT NU_Peso_Media_M3_M1,NU_Peso_Media_M3_M2,NU_Peso_Media_M3_M3 FROM TB_Regras_Aprovacao where CO_Curso = '"&curso&"' and CO_Etapa = '"&co_etapa&"'"
	RSra.Open SQLra, CON0	


	if RSra.EOF then
		peso_m3_m1=1
		peso_m3_m2=1
		peso_m3_m3=1	
	else
		peso_m3_m1=RSra("NU_Peso_Media_M3_M1")
		peso_m3_m2=RSra("NU_Peso_Media_M3_M2")
		peso_m3_m3=RSra("NU_Peso_Media_M3_M3")
		
		if isnull(peso_m3_m1) then
			peso_m3_m1=1		
		end if
		
		if isnull(peso_m3_m2) then
			peso_m3_m2=1		
		end if	
		
		if isnull(peso_m3_m3) then
			peso_m3_m3=1		
		end if				
	end if	
	peso_m3_m1=peso_m3_m1*1
	peso_m3_m2=peso_m3_m2*1	
	peso_m3_m3=peso_m3_m3*1
	peso_periodo_acumulado=peso_m3_m1+peso_m3_m2+peso_m3_m3
	
	prd_prim_media=prd_ter_media-2
	prd_seg_media=prd_ter_media-1
		
	primeira_media = Calc_Prim_Media (unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, prd_prim_media, tipo_calculo, outro)	
	inf_primeira_media=split(primeira_media,"#!#")
	prim_resultado=inf_primeira_media(1)
	
	if prim_resultado = "APR" or prim_resultado = "REP" or prim_resultado="&nbsp;" then
		terceira_media=inf_primeira_media(0)
		terceira_media_ficha="&nbsp;"			
		resultado=prim_resultado
		resultado_ficha=resultado	
	else
	
		segunda_media = Calc_Seg_Media (unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, prd_seg_media, tipo_calculo, outro)	
		inf_segunda_media=split(segunda_media,"#!#")
		seg_resultado=inf_segunda_media(1)
'if cod_aluno=340 then
'response.Write(codigo_materia&"="&seg_resultado&"<BR>")
'end if

		if seg_resultado = "APR" or seg_resultado = "REP" or seg_resultado="&nbsp;" then
			terceira_media=inf_segunda_media(0)
			terceira_media_ficha="&nbsp;"		
			resultado=seg_resultado
			resultado_ficha=resultado 
		else		
			if tipo_calculo="sem_calculo" then		
				if tp_materia="T_F_F_N" then
					codigo_materia_pr=busca_materia_mae(codigo_materia)					
					media_periodo=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia_pr, codigo_materia, CONn , tb_nota, prd_ter_media, nome_nota, outro)	
				elseif tp_materia="T_T_F_N" then	
					vetor_filhas_T_T_F_N=busca_materias_filhas(codigo_materia)
					 media_periodo=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, vetor_filhas_T_T_F_N, CONn, tb_nota, prd_ter_media, nome_nota, outro)					
				elseif tp_materia="T_F_T_N" then
					vetor_filhas_T_F_T_N=busca_materias_filhas(codigo_materia)		
					media_periodo=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, vetor_filhas_T_F_T_N, CONn, tb_nota, prd_ter_media, nome_nota, outro)										
				end if	
	
				if media_periodo="&nbsp;" or isnull(media_periodo) then
					if teste_ano="B" then
						terceira_media=inf_segunda_media(0)
						terceira_media_ficha="&nbsp;"
						resultado= Apura_Resultado(unidade, curso, co_etapa, turma, codigo_materia, terceira_media, "R3", outro)	
						resultado_ficha= Apura_Resultado(unidade, curso, co_etapa, turma, codigo_materia, terceira_media_ficha, "R3", outro)
					else	
						terceira_media=inf_segunda_media(0)
						terceira_media_ficha="&nbsp;"		
						resultado= inf_segunda_media(1)
						resultado_ficha= resultado																		
					end if						
				else
					terceira_media=media_periodo
					terceira_media_ficha=terceira_media
					resultado= Apura_Resultado(unidade, curso, co_etapa, turma, codigo_materia, terceira_media, "R3", outro)
					resultado_ficha=resultado								
				end if		
						
			else		
				codigo_materia_pr=busca_materia_mae(codigo_materia)	
				if tp_materia="T_F_F_N" then
					media_periodo=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia_pr, codigo_materia, CONn , tb_nota, prd_ter_media, nome_nota, outro)	
				elseif tp_materia="T_T_F_N" then	
					 media_periodo=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia_pr, codigo_materia, CONn, tb_nota, periodo, nome_nota, outro)					
				elseif tp_materia="T_F_T_N" then
					media_periodo=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia_pr, codigo_materia, CONn, tb_nota, prd_ter_media, nome_nota, outro)						
				end if	
		
				if media_periodo="&nbsp;" or isnull(media_periodo) then
					if teste_ano="B" then
						media_acumulada=inf_segunda_media(0)
					end if	
				else
					media_acumulada=(inf_primeira_media(0)*peso_m3_m1)+(inf_segunda_media(0)*peso_m3_m2)+(media_periodo*peso_m3_m3)
				end if	
	
				if 	media_acumulada=0 and disc_obrigat="n" then
					terceira_media="&nbsp;"
					terceira_media_ficha=terceira_media
					resultado_ficha=resultado			
				else
					if 	media_acumulada=0 then
						terceira_media="&nbsp;"
						terceira_media_ficha=terceira_media
						resultado=seg_resultado
						resultado_ficha=resultado		
					else				
						terceira_media=arredonda(media_acumulada/peso_periodo_acumulado,"mat_dez",1,0)
						terceira_media_ficha=terceira_media
						resultado= Apura_Resultado(unidade, curso, co_etapa, turma, codigo_materia, terceira_media, "R3", outro)
						resultado_ficha=resultado		
					end if					
				end if	
			end if
		end if	
	end if		
	if outro="ficha" then
		Calc_Ter_Media=terceira_media_ficha&"#!#"&resultado_ficha
	else
		Calc_Ter_Media=terceira_media&"#!#"&resultado
	end if	
end Function

Function Apura_Resultado(unidade, curso, co_etapa, turma, codigo_materia, valor, tipo_resultado, outro)	

	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set RSano = Server.CreateObject("ADODB.Recordset")
	SQLano = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&ano_letivo&"'"
	RSano.Open SQLano, CON

	teste_ano=RSano("ST_Ano_Letivo")		
	
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
	res3_2=RSra("NO_Expr_Int_M3_V")
	res3_1=RSra("NO_Expr_Int_M3_F")

	if isnumeric(valor) then
		valor=valor*1	
	end if
	m1_maior_igual=m1_maior_igual*1
	m1_menor=m1_menor*1
	m2_maior_igual=m2_maior_igual*1
	m2_menor=m2_menor*1
	m3_maior_igual=m3_maior_igual*1
	
	Set RSano = Server.CreateObject("ADODB.Recordset")
	SQLano = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&ano_letivo&"'"
	RSano.Open SQLano, CON

	teste_ano=RSano("ST_Ano_Letivo")	

	if tipo_resultado="R1" then
		if valor >= m1_maior_igual then
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
		if isnumeric(valor) then
			if valor >= m3_maior_igual then
				resultado=res3_2
			else
				resultado=res3_1	
			end if	
		elseif teste_ano="B" then
			resultado=res3_1	
		end if	
	end if
Apura_Resultado=resultado	
end Function
%>