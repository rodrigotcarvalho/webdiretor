<%
'No Saint John esse arquivo para média anual e final no arquivos resultados também
'===========================================================================================================================================
Function Calcula_Frequencia(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia_pr, codigo_materia, conexao , tb_nota, periodo, carga_aula, tipo_calculo, tipo_retorno, outro)

	if tp_freq = "D" then
		Set RSF = Server.CreateObject("ADODB.Recordset")
		SQL_N = "Select * from TB_Frequencia_Periodo WHERE CO_Matricula = "& cod_aluno
		Set RSF = conexao.Execute(SQL_N)
		soma_faltas=0			
		
		if RSF.eof THEN
			f1="&nbsp;"
			f2="&nbsp;"
			f3="&nbsp;"
			f4="&nbsp;"	
		else	
			f1=RSF("NU_Faltas_P1")
			f2=RSF("NU_Faltas_P2")
			f3=RSF("NU_Faltas_P3")
			f4=RSF("NU_Faltas_P4")				
			
			if isnull(f1) or f1= "" then
			else
				f1=f1*1
				soma_faltas=soma_faltas*1
				soma_faltas=soma_faltas+f1		
			end if
			
			if isnull(f2) or f2= "" then
			else
				f2=f2*1
				soma_faltas=soma_faltas*1
				soma_faltas=soma_faltas+f2		
			end if
			
			if isnull(f3) or f3= "" then
			else
				f3=f3*1
				soma_faltas=soma_faltas*1
				soma_faltas=soma_faltas+f3		
			end if
			
			if isnull(f4) or f4= "" then
			else
				f4=f4*1
				soma_faltas=soma_faltas*1
				soma_faltas=soma_faltas+f4		
			end if	
			if soma_faltas=0 then
				soma_faltas="&nbsp;"
			end if	
		end if	
	else
		if tipo_calculo="aluno" then
			Set RSF = Server.CreateObject("ADODB.Recordset")
			SQL_N = "Select SUM(NU_Faltas) as Sum_Faltas from "&tb_nota&" WHERE CO_Matricula = "& cod_aluno
			Set RSF = conexao.Execute(SQL_N)	
			
			if RSF.EOF then
				soma_faltas="&nbsp;"
			else
				soma_faltas=RSF("Sum_Faltas")				
			end if
		elseif tipo_calculo="disciplina" then
			Set RSF = Server.CreateObject("ADODB.Recordset")
			SQL_N = "Select SUM(NU_Faltas) as Sum_Faltas from "&tb_nota&" WHERE CO_Matricula = "& cod_aluno&" AND CO_Materia_Principal = '"&codigo_materia_pr&"' AND CO_Materia='"&codigo_materia&"'"
			Set RSF = conexao.Execute(SQL_N)	
			
			if RSF.EOF then
				soma_faltas="&nbsp;"
			else
				soma_faltas=RSF("Sum_Faltas")				
			end if	
		elseif tipo_calculo="disciplina_periodo" then
			Set RSF = Server.CreateObject("ADODB.Recordset")
			SQL_N = "Select SUM(NU_Faltas) as Sum_Faltas from "&tb_nota&" WHERE CO_Matricula = "& dados_alunos(0)&" AND CO_Materia_Principal = '"&codigo_materia_pr&"' AND CO_Materia='"&codigo_materia&"' AND NU_Periodo="&periodo
			Set RSF = conexao.Execute(SQL_N)	
			
			if RSF.EOF then
				soma_faltas="&nbsp;"
			else
				soma_faltas=RSF("Sum_Faltas")				
			end if												
		end if									
	END IF	
	if tipo_retorno="percent" then
		if isnumeric(soma_faltas) then	
			soma_faltas=soma_faltas*1

			frequencia=((carga_aula-soma_faltas)/carga_aula)*100
			if frequencia<100 then
				frequencia=arredonda(frequencia,"mat_dez",1,0)	
			end if	
		else
			frequencia=100
		end if		
	elseif tipo_retorno="soma" then	
		frequencia=soma_faltas
	end if		

Calcula_Frequencia=	frequencia
end function

'===========================================================================================================================================
'serve também para todas mae=FALSE
Function Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia_pr, codigo_materia, conexao , tb_nota, periodo, nome_nota, outro)


		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT "&nome_nota&" FROM "&tb_nota&" where CO_Matricula ="& cod_aluno &" AND CO_Materia_Principal ='"& codigo_materia_pr &"' AND CO_Materia ='"& codigo_materia &"' And NU_Periodo="&periodo
		RS1.Open SQL1, conexao
		
			if RS1.EOF then
				va_m3="&nbsp;"
			else
				va_m3=RS1(nome_nota)				
			end if		
	Calcula_Media_T_F_F_N=va_m3
'Response.write(va_m3)
'response.end()
end function

'===========================================================================================================================================
Function Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia_pr, vetor_materia, conexao, tb_nota, periodo, nome_nota, outro)	

anulou="n"
acumula=0
divisor=0

	co_materia_mae_fil= split(vetor_materia,"#!#")
	media_mae_acumula=0	
			
	for j=0 to ubound(co_materia_mae_fil)			
		disciplina_filha=co_materia_mae_fil(j)	
		'Exclui a mãe		
		if disciplina_filha <> codigo_materia_pr then
			Set RS = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT * FROM TB_Programa_Aula where CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Materia ='"&disciplina_filha &"'"
			RS.Open SQL, CON0	

			peso=RS("NU_Peso")
			divisor=divisor*1
			if peso="" or isnull(peso) then
				divisor=divisor+1

				peso_multiplica=1
			else	
				peso_multiplica=peso

				peso=peso*1
				divisor=divisor+peso
			end if			
				
			media_aluno=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia_pr, disciplina_filha, conexao, tb_nota, periodo, nome_nota,outro)	
			if media_aluno="" or isnull(media_aluno) or media_aluno="&nbsp;" then
				anulou="s"
			else
				acumula=acumula*1	
				media_aluno=media_aluno*peso_multiplica
				acumula=acumula+media_aluno
			end if	
		end if					
	next

	if divisor =0 then
		anulou="s"
	end if	
	arred_media=parametros_gerais(unidade, curso, co_etapa, turma, codigo_materia_pr,"arred_md_sub_disc",0)
	decimais_media=parametros_gerais(unidade, curso, co_etapa, turma, codigo_materia_pr,"decimais_md_sub_disc",0)	
	if anulou="s" then
		va_m3="&nbsp;"
	else
		va_m3=acumula/divisor
		if va_m3>100 then
			va-m3=100
		end if

		va_m3=arredonda(va_m3,arred_media,decimais_media,0)
	end if

Calcula_Media_T_T_F_N=va_m3		
end function


'===========================================================================================================================================
Function Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia_pr, vetor_materia, conexao, tb_nota, periodo, nome_nota, outro)	

anulou="n"
acumula=0
divisor=0		
	co_materia_mae_fil= split(vetor_materia,"#!#")
	media_mae_acumula=0			
				
	campo_falta=var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo,"BDF")
					
	for j=0 to ubound(co_materia_mae_fil)			
		disciplina_filha=co_materia_mae_fil(j)	
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Programa_Aula where CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Materia ='"&disciplina_filha &"'"
		RS.Open SQL, CON0	

		peso=RS("NU_Peso")
		divisor=divisor*1
		if peso="" or isnull(peso) then
			divisor=divisor+1
			peso_multiplica=1
		else	
			divisor=divisor+peso		
			peso_multiplica=peso
		end if			
			
		media_aluno=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia_pr, disciplina_filha, conexao, tb_nota, periodo, nome_nota, 0)	
										 
		
		if (media_aluno="" or isnull(media_aluno) or media_aluno="&nbsp;") AND nome_nota<>campo_falta then
			anulou="s"
		else
			IF media_aluno="&nbsp;" THEN
			ELSE
				acumula=acumula*1	
				media_aluno=media_aluno*peso_multiplica			
				acumula=acumula+media_aluno
			END IF
		end if					
	next

	if divisor =0 then
		anulou="s"
	end if	
	arred_media=parametros_gerais(unidade, curso, co_etapa, turma, codigo_materia_pr,"arred_md_sub_disc",0)
	decimais_media=parametros_gerais(unidade, curso, co_etapa, turma, codigo_materia_pr,"decimais_md_sub_disc",0)	
	if anulou="s" then
		va_m3="&nbsp;"
	else
		if acumula>100 then
			acumula=100
		end if
		if nome_nota=campo_falta then
			va_m3=acumula
		else
			va_m3=acumula
			va_m3=arredonda(va_m3,arred_media,decimais_media,0)
		end if	
	end if
	'response.Write(va_m3&"-"&acumula&"/"&divisor&" "&anulou&"-<BR>")		

Calcula_Media_T_F_T_N=va_m3		
end function

'===========================================================================================================================================
Function Calcula_Asterisco(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_cons, co_materia, conexao, tp_materia, tb_nota, periodo)	

	var_bd=var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo,"BDM")

	vetor_materia=busca_materias_filhas(co_materia)		
	if tp_materia="T_F_F_N" or tp_materia="F_T_F_N"	 or tp_materia="F_F_T_N" then
		va_media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_cons, co_materia, vetor_materia, conexao, tb_nota, periodo, var_bd, outro)
		
	elseif tp_materia="T_T_F_N" then
		va_media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_cons, co_materia, vetor_materia, conexao, tb_nota, periodo, var_bd, outro)		
			
	elseif tp_materia="T_F_T_N" then	
		va_media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_cons, co_materia, vetor_materia, conexao, tb_nota, periodo, var_bd, outro)		
					
	end if

	if va_media="&nbsp;" then
		Calcula_Asterisco="&nbsp;"
	else
		If tp_modelo="B" then
			periodo_rec= Periodo_Media(tp_modelo,"REC",outro)
		else
			periodo_rec=periodo
		end if
		
		var_bd_rec=var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo_rec,"BDR")	
		
		if tp_materia="T_F_F_N" or tp_materia="F_T_F_N"	 or tp_materia="F_F_T_N" then
		
			codigo_materia_pr=busca_materia_mae(co_materia)
			va_rec=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_cons, codigo_materia_pr, co_materia, conexao, tb_nota, periodo_rec, var_bd_rec, outro)
			
		elseif tp_materia="T_T_F_N" then
		
			vetor_materia=busca_materias_filhas(co_materia)
			va_rec=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_cons, co_materia, vetor_materia, conexao, tb_nota, periodo_rec, var_bd_rec, outro)		
				
		elseif tp_materia="T_F_T_N" then
		
			vetor_materia=busca_materias_filhas(co_materia)			
			va_rec=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_cons, co_materia, vetor_materia, conexao, tb_nota, periodo_rec, var_bd_rec, outro)	
					
		end if
		
		if isnull(va_rec) or va_rec="" then
			va_rec="&nbsp;"
		end if
	arred_media=parametros_gerais(unidade, curso, co_etapa, turma, codigo_materia_pr,"arred_md",0)
	decimais_media=parametros_gerais(unidade, curso, co_etapa, turma, codigo_materia_pr,"decimais_md",0)			
		if va_rec="&nbsp;" then
			Calcula_Asterisco="&nbsp;"
		else
		
			if tp_materia="F_T_F_N"	 or tp_materia="F_F_T_N" then		
				va_media=va_media*1
				va_rec=va_rec*1
				if va_media<7 and va_media<va_rec then
					Calcula_Asterisco=va_rec
				else
					Calcula_Asterisco=va_media
				end if			
			else	
				if va_media<7 and va_media<va_rec then
					va_ast=(va_media+va_rec)/2
					if va_ast< va_media then
						va_ast=va_media
					end if
					Calcula_Asterisco=arredonda(va_ast,arred_media,decimais_media,0)
				else
					Calcula_Asterisco=va_media
				end if
			end if	
		end if
	end if
end function




'===========================================================================================================================================
Function Calcula_Soma(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_cons, co_materia, tp_materia, conexao, tb_nota, maximo_periodo,usa_asteristico,  origem, outro)	

	valor="naoNulo"	
	acumula=0
	


	for prd_soma=1 to maximo_periodo
	
		'If tp_modelo="B" then
			periodo_rec= Periodo_Media(tp_modelo,"REC",outro)
		'else
		'	periodo_rec=prd_soma
		'end if	
		var_bd=var_bd_periodo(tp_modelo,tp_freq,tb_nota,prd_soma,"BDM")
		var_bd_rec=var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo_rec,"BDR")		
			
		if tp_materia="T_F_F_N" or tp_materia="F_T_F_N"	 or tp_materia="F_F_T_N" then
			'response.Write(unidade&"-"&curso&"-"&co_etapa&"-"&turma&"-"&cod_cons&"-"&codigo_materia_pr&"-"&co_materia&"-"&conexao&"-"&tb_nota&"-"&prd_soma&"-"&var_bd&"-"&outro&"<BR>")		
			codigo_materia_pr=busca_materia_mae(co_materia)
			va_media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_cons, codigo_materia_pr, co_materia, conexao, tb_nota, prd_soma, var_bd, outro)
			'response.Write(va_media&"<BR>")
			if usa_asteristico = "S" then
				if periodo_rec>=prd_soma then
					va_ast=Calcula_Asterisco(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_cons, co_materia, conexao, tp_materia, tb_nota, periodo_rec)
				else
					va_ast="&nbsp;"
				end if
			else
				va_ast="&nbsp;"			
			end if	

			if isnumeric(va_media) then
				if isnumeric(va_ast) then
					va_ast=va_ast*1
					va_media=va_media*1					
					if va_ast>va_media then
						somar=va_ast
					else
						somar=va_media			
					end if	
				else
					somar=va_media					
				end if	
			else
				valor="nulo"
			end if					
'response.Write(somar&"<BR>")
		elseif tp_materia="T_T_F_N" then
		
			vetor_filhas_T_T_F_N=busca_materias_filhas(co_materia)
			va_media=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_cons, co_materia, vetor_filhas_T_T_F_N, conexao, tb_nota, prd_soma, var_bd, outro)		
			va_ast=Calcula_Asterisco(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_cons, co_materia, conexao, tp_materia, tb_nota, periodo_rec)	
			
			if isnumeric(va_media) then
				if isnumeric(va_ast) then
					va_ast=va_ast*1
					va_media=va_media*1					
					if va_ast>va_media then
						somar=va_ast
					else
						somar=va_media			
					end if	
				else
					somar=va_media					
				end if	
			else
				valor="nulo"
			end if				
	
		elseif tp_materia="T_F_T_N" then
		
			vetor_filhas_T_F_T_N=busca_materias_filhas(co_materia)			
			va_media=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_cons, co_materia, vetor_filhas_T_F_T_N, conexao, tb_nota, prd_soma, var_bd, outro)	
			va_ast=Calcula_Asterisco(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_cons, co_materia, conexao, tp_materia, tb_nota, periodo_rec)	
			
			if isnumeric(va_media) then
				if isnumeric(va_ast) then
					va_ast=va_ast*1
					va_media=va_media*1					
					if va_ast>va_media then
						somar=va_ast
					else
						somar=va_media			
					end if	
				else
					somar=va_media					
				end if	
			else
				valor="nulo"
			end if								
		end if
'response.Write(somar&" "&valor&"<BR>")		
		if valor<>"nulo" then	
			acumula=acumula+somar
		end if			
	next	
	arred_media=parametros_gerais(unidade, curso, co_etapa, turma, codigo_materia_pr,"arred_md",0)
	decimais_media=parametros_gerais(unidade, curso, co_etapa, turma, codigo_materia_pr,"decimais_md",0)	

	if valor="nulo" then
		Calcula_Soma="&nbsp;"
	else
		Calcula_Soma=arredonda(acumula,arred_media,decimais_media,0)
	end if
end function

'=================================================
FUNCTION CALCULA_LINHA_BOLETIM(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, co_materia_mae, disciplina, CONn , tb_nota, periodo_cons, nome_nota, PERIODO_ANUAL, PERIODO_RECUPERACAO, PERIODO_FINAL, TIPO_MATERIA, outro)

	vetor_linha_boletim = -1
	'response.write(TIPO_MATERIA&"("&unidade&","& curso&","&co_etapa&","&turma&","&tp_modelo&","&tp_freq&","&cod_aluno&","&co_materia_mae&","&disciplina&","&CONn&","&tb_nota&","&periodo_cons&","&nome_nota&","&outro&")<br>")
	'response.end()

    IF TIPO_MATERIA="T_F_F_N" THEN		
		media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, co_materia_mae, disciplina, CONn , tb_nota, periodo_cons, nome_nota, outro)	
	ELSEIF TIPO_MATERIA="T_T_F_N" THEN
		media=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, co_materia_mae, disciplina, CONn, tb_nota, periodo_cons, nome_nota, outro)											      	
	ELSE
		media=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, co_materia_mae, disciplina, CONn, tb_nota, periodo_cons, nome_nota, outro)			
	END IF

	if not isnumeric(media)then
		conceito_media=""
		calcula_media_anual="nao"				
	else
		conceito_media=converte_conceito(unidade, curso, co_etapa, turma, periodo, disciplina, media, outro)			
		if isnumeric(conceito_media) then	
			conceito_media = formatNumber(conceito_media/10,1)	
		end if					
		calcula_media_anual="sim"
	end if
	'response.write(vetor_linha_boletim&"------------------"&conceito_media&"<BR>")
	if vetor_linha_boletim = -1 then
		vetor_linha_boletim = conceito_media
	else
		vetor_linha_boletim = vetor_linha_boletim&"#!#"&conceito_media
	end if		
'response.write(vetor_linha_boletim&"ok<br>")
	
	periodo_cons=periodo_cons*1	
	PERIODO_ANUAL=PERIODO_ANUAL*1	
	PERIODO_RECUPERACAO=PERIODO_RECUPERACAO*1	
	PERIODO_FINAL=PERIODO_FINAL*1		
	if periodo_cons=PERIODO_ANUAL then
		conceito=converte_conceito(unidade, curso, co_etapa, turma, periodo, disciplina, media, outro)							

		if calcula_media_anual="sim" then
	
			media_calc1=Calc_Prim_Media (unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, disciplina, CONn, tb_nota, PERIODO_ANUAL, tipo_calculo, aproxima_m1, outro)	
			response.write(vetor_linha_boletim&"ok3<br>")	
			resultados=split(media_calc1,"#!#")
			media_anual=resultados(0)
			resultado_anual=resultados(1)
			vetor_aluno=vetor_aluno*1
	
			if resultado_anual<>"&nbsp;" then
				tipo_media = "MA"
				modifica_result = Verifica_Conselho_Classe(cod_aluno, disciplina, tipo_media, outro)
				if modifica_result <> "N" then
					resultado_anual = modifica_result
				end if		
			end if							
			media_anual = arredonda(media_anual,"mat",1,outro)	
			conceito_anual=converte_conceito(unidade, curso, co_etapa, turma, periodo, disciplina, media_anual, outro)													
			if isnumeric(conceito_anual) then				
				conceito_anual=conceito_anual/10									
				conceito_anual = formatNumber(conceito_anual,1)	
			end if						
		else
			conceito_anual=""
			resultado_anual=""
		end if														
		
		vetor_linha_boletim = vetor_linha_boletim&"#!#"&conceito_anual&"#!#"&resultado_anual
	
	elseif periodo_cons=PERIODO_RECUPERACAO then									
		if media="" or isnull(media) or media_anual="" or isnull(media_anual)then
			conceito_recup=""
			resultado_recup=""
		else
			media_calc2=Calc_Seg_Media (unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, disciplina, CONn, tb_nota, PERIODO_RECUPERACAO, tipo_calculo, compara_m2, aproxima_m2, outro)
			resultados=split(media_calc2,"#!#")
			media_recup=resultados(0)
			resultado_recup=resultados(1)	
			if resultado_recup<>"&nbsp;" then
				tipo_media = "RF"
				modifica_result = Verifica_Conselho_Classe(cod_aluno, disciplina, tipo_media, outro)
				if modifica_result <> "N" then
					resultado_recup = modifica_result
				end if																										
			end if										
			media_recup = arredonda(media_recup,"mat",1,outro)
			conceito_recup = converte_conceito(unidade, curso, co_etapa, turma, periodo, disciplina, media_recup, outro)			
			if isnumeric(conceito_recup) then				
				conceito_recup=conceito_recup/10									
				conceito_recup = formatNumber(conceito_recup,1)	
			end if																												
		end if					
			
		vetor_linha_boletim = vetor_linha_boletim&"#!#"&conceito_media&"#!#"&conceito_recup&"#!#"&resultado_recup														
	elseif periodo_cons=PERIODO_FINAL then									
		if media="" or isnull(media) or media_anual="" or isnull(media_anual)  or media_recup="" or isnull(media_recup)then
			conceito_recup=""
			resultado_final=""							
		else
			media_calc3=Calc_Ter_Media (unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, disciplina, CONn, tb_nota, PERIODO_FINAL, tipo_calculo, compara_m3, aproxima_m3, outro)
			resultados=split(media_calc3,"#!#")
			media_final=resultados(0)
			resultado_final=resultados(1)	
			if resultado_final<>"&nbsp;" then
				tipo_media = "MF"
				modifica_result = Verifica_Conselho_Classe(cod_aluno, disciplina, tipo_media, outro)
				if modifica_result <> "N" then
					resultado_final = modifica_result
				end if	
			end if	
			media_final = arredonda(media_final,"mat",1,outro)
			conceito_final = converte_conceito(unidade, curso, co_etapa, turma, periodo, disciplina, media_final, outro)			
			if isnumeric(conceito_final) then				
				conceito_final=conceito_final/10									
				conceito_final = formatNumber(conceito_final,1)	
			end if																
		end if		
			
		vetor_linha_boletim = vetor_linha_boletim&"#!#"&conceito_final&"#!#"&resultado_final									
	end if
'response.write(vetor_linha_boletim)
CALCULA_LINHA_BOLETIM = vetor_linha_boletim	
END FUNCTION				



%>
