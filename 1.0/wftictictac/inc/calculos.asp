
<%'===========================================================================================================================================
'serve também para todas mae=FALSE
Function Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia_pr, codigo_materia, conexao , tb_nota, periodo, nome_nota, outro)

Server.ScriptTimeout = 900

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT "&nome_nota&" FROM "&tb_nota&" where CO_Matricula ="& cod_aluno &" AND CO_Materia_Principal ='"& codigo_materia_pr &"' AND CO_Materia ='"& codigo_materia &"' And NU_Periodo="&periodo
		RS1.Open SQL1, conexao
		
			if RS1.EOF then
				va_m3="&nbsp;"
			else
				va_m3=RS1(nome_nota)				
			end if		
	Calcula_Media_T_F_F_N=va_m3

end function

'===========================================================================================================================================
Function Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia_pr, vetor_materia, conexao, tb_nota, periodo, nome_nota, outro)	

anulou="n"
acumula=0
divisor=0
			
	co_materia_mae_fil= split(vetor_materia,"#!#")
	media_mae_acumula=0						
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
			peso_multiplica=peso

			peso=peso*1
			divisor=divisor+peso
		end if			
			
		media_aluno=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia_pr, disciplina_filha, conexao, tb_nota, periodo, nome_nota,outro)	
			if media_aluno="" or isnull(media_aluno) or media_aluno="&nbsp;" then
				anulou="s"
			else
				acumula=acumula*1	
				media_aluno=media_aluno*peso_multiplica
				acumula=acumula+media_aluno
			end if					
	next

	if divisor =0 then
		anulou="s"
	end if	

	if anulou="s" then
		va_m3="&nbsp;"
	else
		va_m3=acumula/divisor
		va_m3=arredonda(va_m3,parametros_gerais("arred_media"),parametros_gerais("decimais_media"),0)
	end if

Calcula_Media_T_T_F_N=va_m3		
end function


'===========================================================================================================================================
Function Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia_pr, vetor_materia, conexao, tb_nota, periodo, nome_nota, outro)	

anulou="n"
acumula=0
divisor=0		
	co_materia_mae_fil= split(vetor_materia,"#!#")
	media_mae_acumula=0			
	
	tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")
	tp_freq=tipo_divisao_ano(curso,co_etapa,"in_frequencia") 				
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
		
		media_aluno=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia_pr, disciplina_filha, conexao, tb_nota, periodo, nome_nota, 0)	

		if (media_aluno="" or isnull(media_aluno) or media_aluno="&nbsp;") AND nome_nota<>campo_falta then
			anulou="s"
		else
			IF media_aluno="&nbsp;" or media_aluno="" or isnull(media_aluno)THEN
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
			
	if anulou="s" then
		va_m3="&nbsp;"
	else
		if nome_nota=campo_falta then
			va_m3=acumula
		else
			va_m3=acumula/divisor
			va_m3=arredonda(va_m3,parametros_gerais("arred_media"),parametros_gerais("decimais_media"),0)
		end if	
	end if

Calcula_Media_T_F_T_N=va_m3		
end function

'===========================================================================================================================================
Function Calcula_Asterisco(tp_modelo, tp_freq, unidade, curso, co_etapa, turma, cod_cons, co_materia, conexao, tp_materia, tb_nota, periodo)	

	var_bd=var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo,"BDM")
		
	if tp_materia="T_F_F_N" or tp_materia="F_T_F_N"	 or tp_materia="F_F_T_N" then
	
		codigo_materia_pr=busca_materia_mae(co_materia)
		va_media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_cons, codigo_materia_pr, co_materia, conexao, tb_nota, periodo, var_bd, outro)
		
	elseif tp_materia="T_T_F_N" then
	
		vetor_materia=busca_materias_filhas(co_materia)
		va_media=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, cod_cons, co_materia, vetor_materia, conexao, tb_nota, periodo, var_bd, outro)		
			
	elseif tp_materia="T_F_T_N" then
		vetor_materia=busca_materias_filhas(co_materia)			
		va_media=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, cod_cons, co_materia, vetor_materia, conexao, tb_nota, periodo, var_bd, outro)	
				
	end if

	if va_media="&nbsp;" then
		Calcula_Asterisco="&nbsp;"
	else
		if isnumeric(va_media) then
			va_media = formatNumber(va_media,1)	
			va_media=va_media*1	
		end if	
	
		If tp_modelo="B" then
			periodo_rec= 2
		else
			periodo_rec=periodo
		end if
		
		var_bd_rec=var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo_rec,"BDR")	
		
		if tp_materia="T_F_F_N" or tp_materia="F_T_F_N"	 or tp_materia="F_F_T_N" then
		
			codigo_materia_pr=busca_materia_mae(co_materia)
			va_rec=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_cons, codigo_materia_pr, co_materia, conexao, tb_nota, periodo_rec, var_bd_rec, outro)
			
		elseif tp_materia="T_T_F_N" then
		
			vetor_materia=busca_materias_filhas(co_materia)
			va_rec=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, cod_cons, co_materia, vetor_materia, conexao, tb_nota, periodo_rec, var_bd_rec, outro)		
				
		elseif tp_materia="T_F_T_N" then
		
			vetor_materia=busca_materias_filhas(co_materia)			
			va_rec=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, cod_cons, co_materia, vetor_materia, conexao, tb_nota, periodo_rec, var_bd_rec, outro)	
					
		end if
		
		if isnull(va_rec) or va_rec="" then
			va_rec="&nbsp;"
		else
			if isnumeric(va_rec) then
				va_rec = formatNumber(va_rec,1)	
				va_rec=va_rec*1	
			end if	
		end if
		
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
					va_media=va_media*1
					va_rec=va_rec*1
					va_ast=(va_media+va_rec)/2

					if va_ast< va_media then
						va_ast=va_media
					end if
					Calcula_Asterisco=arredonda(va_ast,parametros_gerais("arred_media"),parametros_gerais("decimais_media"),0)
				else
					Calcula_Asterisco=va_media
				end if			
			end if	
		end if
	end if
end function




'===========================================================================================================================================
Function Calcula_Soma(tp_modelo, tp_freq, unidade, curso, co_etapa, turma, cod_cons, co_materia, conexao, tp_materia, tb_nota, maximo_periodo, outro)	

	valor="ok"	
	acumula=0
	
'response.Write(cod_cons&"_"&co_materia&"_"&tp_materia&"<BR>")

	for prd_soma=1 to maximo_periodo
'			response.Write("P="&prd_soma&"<BR>")	
		somar=0
		If tp_modelo="B" then
			periodo_rec= 2
		else
			periodo_rec=prd_soma
		end if	
		var_bd=var_bd_periodo(tp_modelo,tp_freq,tb_nota,prd_soma,"BDM")
		var_bd_rec=var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo_rec,"BDR")		
			
		if tp_materia="T_F_F_N" or tp_materia="F_T_F_N"	 or tp_materia="F_F_T_N" then
			
			codigo_materia_pr=busca_materia_mae(co_materia)
			va_media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_cons, codigo_materia_pr, co_materia, conexao, tb_nota, prd_soma, var_bd, outro)
			if periodo_rec>=prd_soma then
				va_ast=Calcula_Asterisco(tp_modelo, tp_freq, unidade, curso, co_etapa, turma, cod_cons, co_materia, conexao, tp_materia, tb_nota, prd_soma)
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

		elseif tp_materia="T_T_F_N" then
		
			vetor_filhas_T_T_F_N=busca_materias_filhas(co_materia)
			va_media=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, cod_cons, co_materia, vetor_filhas_T_T_F_N, conexao, tb_nota, prd_soma, var_bd, outro)		
			va_ast=Calcula_Asterisco(tp_modelo, tp_freq, unidade, curso, co_etapa, turma, cod_cons, co_materia, conexao, tp_materia, tb_nota, periodo_rec)	
			
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
			va_media=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, cod_cons, co_materia, vetor_filhas_T_F_T_N, conexao, tb_nota, prd_soma, var_bd, outro)	
			va_ast=Calcula_Asterisco(tp_modelo, tp_freq, unidade, curso, co_etapa, turma, cod_cons, co_materia, conexao, tp_materia, tb_nota, periodo_rec)	
			
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
'		response.Write("S+"&somar&"<BR>")			
		if valor<>"nulo" then	
			acumula=acumula+somar
		end if	
'		response.Write("A+"&acumula&"<BR>")				
	next	
'response.End()
	if valor="nulo" then
		Calcula_Soma="&nbsp;"
	else
		Calcula_Soma=arredonda(acumula,parametros_gerais("arred_media"),parametros_gerais("decimais_media"),0)
	end if
end function
%>
