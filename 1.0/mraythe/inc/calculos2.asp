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
		SQL1 = "SELECT VA_Media3 FROM "&tb_nota&" where CO_Matricula ="& cod_aluno &" AND CO_Materia_Principal ='"& codigo_materia &"' AND CO_Materia ='"& codigo_materia &"' And NU_Periodo="&periodo
		RS1.Open SQL1, conexao
		
		if RS1.EOF then
			va_m3=""
		else
			va_m3=RS1("VA_Media3")				
		end if		
	Calcula_Media_T_F_F_N=va_m3

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
		va_m3=arredonda(va_m3,arred_media,decimais_media,0)
	end if

Calcula_Media_T_T_F_N=va_m3		
end function


'===========================================================================================================================================
Function Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia_pr, vetor_materia, conexao, tb_nota, periodo, nome_nota, outro)	

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	

	Set RS1a = Server.CreateObject("ADODB.Recordset")
	SQL1a = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& codigo_materia &"' order by NU_Ordem_Boletim"
	RS1a.Open SQL1a, CON0
		
	if RS1a.EOF then
	else
		co_materia_fil_check=1 
		peso_acumula=0
		va_m3_acumula=0

		while not RS1a.EOF
			co_mat_fil= RS1a("CO_Materia")	
			
			Set RSp2 = Server.CreateObject("ADODB.Recordset")
			SQLp2 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia = '"& co_mat_fil &"' order by NU_Ordem_Boletim"

			RSp2.Open SQLp2, CON0	
									
			nu_peso_fil=RSp2("NU_Peso")	
						
			peso_acumula=peso_acumula+nu_peso_fil
										
			Set RS3 = Server.CreateObject("ADODB.Recordset")
			SQL_N = "Select VA_Media3 from "& tb_nota &" WHERE CO_Matricula = "& cod_aluno &" AND CO_Materia = '"& co_mat_fil &"' AND CO_Materia_Principal = '"& codigo_materia &"' AND NU_Periodo="&periodo
			Set RS3 = conexao.Execute(SQL_N)						
	

			if RS3.EOF then
				va_m3_temp=""
			else					
				va_m3_temp=RS3("VA_Media3")
			end if
	
			if isnull(va_m3_temp) or va_m3_temp="&nbsp;"  or va_m3_temp="" then
				sem_nota="s"
			else
				va_m3_acumula=va_m3_acumula+va_m3_temp								
			end if	
																							
		RS1a.MOVENEXT
		wend		
		
		if sem_nota="s" then
			va_m3=""
		else	
			va_m3=va_m3_acumula/peso_acumula
			va_m3 = arredonda(va_m3,"quarto_dez",1,0)								
		end if	
	end if
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

	valor="ok"	
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
'			response.Write(unidade&"-"&curso&"-"&co_etapa&"-"&turma&"-"&cod_cons&"-"&codigo_materia_pr&"-"&co_materia&"-"&conexao&"-"&tb_nota&"-"&prd_soma&"-"&var_bd&"-"&outro&"<BR>")		
			codigo_materia_pr=busca_materia_mae(co_materia)
			va_media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_cons, codigo_materia_pr, co_materia, conexao, tb_nota, prd_soma, var_bd, outro)

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
%>
