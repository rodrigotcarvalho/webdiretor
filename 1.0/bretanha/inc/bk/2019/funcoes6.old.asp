<!--#include file="../../global/funcoes_diversas.asp" -->
<%
Function tabela_notas(CON, unidade, curso, co_etapa, turma, periodo, disciplina, outro)

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidade &" And CO_Curso= '"& curso &"' And CO_Etapa = '"& co_etapa &"'"
		Set RS = CON.Execute(SQL)
			
		if RS.EOF then
			tabela_notas = ""	
		else
			tabela_notas = RS("TP_Nota")
		end if	

end function

'===========================================================================================================================================

Function caminho_notas(CON, tb_nota, outro)

		if tb_nota="TB_NOTA_A" then
			CAMINHO_n=CAMINHO_na
		elseif tb_nota="TB_NOTA_B" then
			CAMINHO_n=CAMINHO_nb
		elseif tb_nota="TB_NOTA_C" then
			CAMINHO_n=CAMINHO_nc
		elseif tb_nota="TB_NOTA_D" then
			CAMINHO_n=CAMINHO_nd	
		elseif tb_nota="TB_NOTA_E" then
			CAMINHO_n=CAMINHO_ne
		else
			CAMINHO_n = "ERRO"					
		end if
	
	caminho_notas=CAMINHO_n

end function

'===========================================================================================================================================
Function alunos_esta_turma(CON, ano_letivo, nome_campo, unidade, curso, co_etapa, turma, cursando, ordena, outro)

	if ano_letivo=0 then
		sql_ano_letivo = ""
	else
		sql_ano_letivo = " and NU_Ano = "&ano_letivo
	end if
	
	if nome_campo="*" then
		sql_nome_campo = "*"
	else
		sql_nome_campo = "CO_Matricula"
	end if	

	if cursando = "C" then
		sql_cursando= " AND CO_Situacao = 'C'"
	else
		sql_cursando = ""
	end if

	if ordena = "*" then
		sql_ordena = ""
	else
		sql_ordena = " order by "&ordena		
	end if
	
	if turma="*" then
		sql_turma = ""
	else
		sql_turma=" AND CO_Turma = '"&turma&"'"
	end if	


	Set RSA= Server.CreateObject("ADODB.Recordset")
	SQLA = "Select "&sql_nome_campo&" from TB_Matriculas WHERE NU_Unidade = "& unidade &" And CO_Curso = '"& curso &"' And CO_Etapa = '"& co_etapa&"'"&sql_turma&sql_ano_letivo&sql_cursando&sql_ordena
	Set RSA = CON.Execute(SQLA)
	
	total_alunos=0
	while not RSA.EOF
		campo = RSA(sql_nome_campo)	
		if total_alunos=0 then
			informacoes_alunos = campo
		else
			informacoes_alunos = informacoes_alunos&"#!#"&campo		
		end if
		total_alunos=total_alunos+1
	RSA.MOVENEXT
	WEND
	
	alunos_esta_turma=informacoes_alunos

end function

'==========================================================================================================================================
Function tipo_divisao_ano(curso,co_etapa,tipo_dado)

ano_letivo=session("ano_letivo")

'	Set RS = Server.CreateObject("ADODB.Recordset")
'	SQL = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"'"
'	RS.Open SQL, CON0
'	
'	if ano_letivo<2011 and tipo_dado="tp_modelo" then
'		modelo="B"
'	else
'		if RS.EOF then
'			modelo="B"
'			freq="D"
'		else
''			curso=curso*1
''			if curso<2 then
''				modelo="B"
''				freq="M"			
''			else
''				modelo="T"
''				freq="M"			
''			end if
'			modelo=RS("TP_Modelo")
'			freq=RS("IN_Frequencia")
'		end if
'	end if

	modelo="B"
	freq="D"	
	
	if tipo_dado="tp_modelo" then
		tipo_divisao_ano=modelo
	elseif tipo_dado="in_frequencia" then
		tipo_divisao_ano=freq
	end if
end function






'===========================================================================================================================================
Function tipo_materia(co_materia, curso, co_etapa)

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& co_materia &"'"
		RS.Open SQL, CON0
	
		mae= RS("IN_MAE")
		fil= RS("IN_FIL")
		in_co= RS("IN_CO")
		peso= RS("NU_Peso")
		
		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) then
			tipo_materia="T_F_F_N"
		elseif (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) then	
			tipo_materia="T_T_F_N"
		elseif (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then
			tipo_materia="T_F_T_N"
		elseif (mae=FALSE and fil=TRUE and in_co=FALSE and isnull(peso)) then	
			tipo_materia="F_T_F_N"			
		elseif (mae=FALSE and fil=FALSE and in_co=TRUE and isnull(peso)) then
			tipo_materia="F_F_T_N"
		end if	
end function


'===========================================================================================================================================
Function busca_materia_mae(co_materia)

	Set RS1 = Server.CreateObject("ADODB.Recordset")
	SQL1 = "SELECT * FROM TB_Materia WHERE CO_Materia='"&co_materia&"'"
	RS1.Open SQL1, CON0
	
	if RS1.EOF then
		busca_materia_mae=co_materia
	else
		materia_mae=RS1("CO_Materia_Principal")
		
		if isnull(materia_mae) or materia_mae="" then
			busca_materia_mae=co_materia		
		else
			busca_materia_mae=materia_mae		
		end if	
	end if

end function				
'===========================================================================================================================================

Function busca_materias_filhas(co_materia)

co_materia_check=1

	Set RS1 = Server.CreateObject("ADODB.Recordset")
	SQL1 = "SELECT * FROM TB_Materia WHERE CO_Materia_Principal='"&co_materia&"'"
	RS1.Open SQL1, CON0
	
	if RS1.EOF then
		busca_materias_filhas=co_materia
	else
		while not RS1.EOF
		co_mat_fil=RS1("CO_Materia")
			if co_materia_check=1 then
				vetor_materia_filha=co_mat_fil
			else
				vetor_materia_filha=vetor_materia_filha&"#!#"&co_mat_fil
			end if
			co_materia_check=co_materia_check+1		
		RS1.MOVENEXT
		WEND
		busca_materias_filhas=vetor_materia_filha
	end if
	
end function	

'===========================================================================================================================================
Function periodos(periodo, opcao)

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0

if opcao="num" then
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Periodo order by NU_Periodo"
	RS.Open SQL, CON0
	conta=0	
	while not RS.EOF
		nu_periodo =  RS("NU_Periodo")

		if conta=0 then
			vetor_periodo=nu_periodo			
		else
			vetor_periodo=vetor_periodo&"#!#"&nu_periodo		
		end if
		conta=conta+1
	RS.Movenext
	Wend		

elseif opcao="sigla" then

	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Periodo where NU_Periodo="&periodo
	RS.Open SQL, CON0
	
	vetor_periodo = RS("SG_Periodo")	

elseif opcao="nome" then

	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Periodo where NU_Periodo="&periodo
	RS.Open SQL, CON0
	
	vetor_periodo = RS("NO_Periodo")	

elseif opcao="todas_siglas" then
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Periodo order by NU_Periodo"
	RS.Open SQL, CON0
	
	while not RS.EOF
		nu_periodo =  RS("NU_Periodo")
		sg_periodo = RS("SG_Periodo")

		
	RS.Movenext
	Wend
elseif opcao="todos_nomes" then
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Periodo order by NU_Periodo"
	RS.Open SQL, CON0
	
	while not RS.EOF
		nu_periodo =  RS("NU_Periodo")
		no_periodo = RS("NO_Periodo")		

		
	RS.Movenext
	Wend	
end if
periodos=vetor_periodo
end function






'===========================================================================================================================================
Function programa_aula(vetor_materia, unidade, curso, co_etapa, turma)

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
	
if vetor_materias<>"nulo" then		
	co_materia= split(vetor_materia,"#!#")
	co_materia_check=1	
	For i=0 to ubound(co_materia)
	
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& co_materia(i) &"'"
		RS.Open SQL, CON0
	
		mae= RS("IN_MAE")
		fil= RS("IN_FIL")
		in_co= RS("IN_CO")
		peso= RS("NU_Peso")
		
		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) or (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) then
	
			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_materia(i) &"' order by NU_Ordem_Boletim"
			RS1.Open SQL1, CON0
				
			if RS1.EOF then
				if co_materia_check=1 then
					vetor_materia_exibe=co_materia(i)
				else
					vetor_materia_exibe=vetor_materia_exibe&"#!#"&co_materia(i)
				end if
				co_materia_check=co_materia_check+1		
			else
			co_materia_fil_check=1 
				while not RS1.EOF
					co_mat_fil= RS1("CO_Materia")				
					if co_materia_check=1 and co_materia_fil_check=1 then
						vetor_materia_exibe=co_materia(i)&"#!#"&co_mat_fil
					elseif co_materia_fil_check=1 then
						vetor_materia_exibe=vetor_materia_exibe&"#!#"&co_materia(i)&"#!#"&co_mat_fil
					else
						vetor_materia_exibe=vetor_materia_exibe&"#!#"&co_mat_fil			
					end if
					co_materia_check=co_materia_check+1
					co_materia_fil_check=co_materia_fil_check+1 									
				RS1.MOVENEXT
				wend
				vetor_materia_exibe=vetor_materia_exibe&"#!#MED"	
			end if
		end if	
	NEXT
programa_aula=vetor_materia_exibe
else
end if
end function


Function posicao_materia_tabela(co_materia, unidade, curso, co_etapa, turma)

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0

if co_materia="nulo" then
	posicao=0	
elseif co_materia="MED" then
		posicao=3		
else	
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& co_materia &"'"
'response.Write(SQL&"<BR>")
    RS.Open SQL, CON0

	mae= RS("IN_MAE")
	fil= RS("IN_FIL")
	in_co= RS("IN_CO")
	peso= RS("NU_Peso")
	
	if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) or (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) or (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then
		posicao=1
	elseif mae=False and fil=TRUE and in_co=FALSE THEN
		posicao=2
	elseif mae=False and fil=False and in_co=TRUE and isnull(peso) THEN
		posicao=0					
	end if	
end if		
posicao_materia_tabela=posicao
end function


Function calcula_medias(unidade, curso, co_etapa, turma, periodo, vetor_aluno, vetor_materia, caminho_nota, tb_nota, nome_nota, tipo_calculo)

'response.Write(vetor_materia&"<BR>")
'response.Write(vetor_aluno&"<BR>")
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CONn = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONn.Open ABRIRn		

if tipo_calculo="media_turma" then	
	
	co_materia= split(vetor_materia,"#!#")	
	co_materia_check=1	
	
	For i=0 to ubound(co_materia)
	
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& co_materia(i) &"'"
		RS.Open SQL, CON0
	
		mae= RS("IN_MAE")
		fil= RS("IN_FIL")
		in_co= RS("IN_CO")
		peso= RS("NU_Peso")
		
		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) then

			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& vetor_aluno &") AND CO_Materia ='"& co_materia(i)&"' And NU_Periodo="&periodo
			RS1.Open SQL1, CONn
			
			media_turma=RS1("MediaDeVA_Media3")
			if media_turma="" or isnull(media_turma) then
			else
			media_turma=formatnumber(media_turma,0)
			end if 
				if co_materia_check=1 then
					vetor_quadro=media_turma
				else
					vetor_quadro=vetor_quadro&"#!#"&media_turma
				end if
				
		elseif (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) then
		
			vetor_mae_filhas=""
	
			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL2 = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_materia(i) &"'"
			RS2.Open SQL2, CON0
				
			co_materia_fil_check=0 
			while not RS2.EOF
				co_mat_fil= RS2("CO_Materia")				
				if co_materia_fil_check=0 then
					vetor_mae_filhas=co_materia(i)&"#!#"&co_mat_fil
				else
					vetor_mae_filhas=vetor_mae_filhas&"#!#"&co_mat_fil			
				end if
				co_materia_fil_check=co_materia_fil_check+1 									
			RS2.MOVENEXT
			wend				
			
			co_materia_mae_fil= split(vetor_mae_filhas,"#!#")
			media_mae_acumula=0			
			for j=0 to ubound(co_materia_mae_fil)			
		
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL3 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& vetor_aluno &") AND CO_Materia ='"& co_materia_mae_fil(j)&"' And NU_Periodo="&periodo
				RS3.Open SQL3, CONn

'response.Write(media_mae_acumula)					
				media_turma=RS3("MediaDeVA_Media3")
				if media_turma="" or isnull(media_turma) then
				media_filha_acumula=0	
				else
				media_turma=formatnumber(media_turma,0)
				media_filha_acumula=media_turma
				end if 

				if co_materia_check=1 then
					vetor_quadro=media_turma
					media_mae_acumula=media_mae_acumula+media_filha_acumula	
				else
					vetor_quadro=vetor_quadro&"#!#"&media_turma
					media_mae_acumula=media_mae_acumula*1
					media_turma=media_turma*1
					media_mae_acumula=media_mae_acumula+media_filha_acumula		
				end if		
				'response.Write(co_materia_mae_fil(j)&"-"&media_turma&"-"&media_mae_acumula&"-"&co_materia_fil_check&"<BR>")		
			next
			media_mae=media_mae_acumula/co_materia_fil_check
			media_mae=formatnumber(media_mae,0)
			vetor_quadro=vetor_quadro&"#!#"&media_mae	
		end if		
	co_materia_check=co_materia_check+1			
	NEXT
	
elseif tipo_calculo="media_geral" then	
	
				Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& vetor_aluno &") And NU_Periodo="&periodo
			RS1.Open SQL1, CONn
			
			media_turma=RS1("MediaDeVA_Media3")
			if media_turma="" or isnull(media_turma) then
			media_turma=0
			else
			media_turma=formatnumber(media_turma,0)
			end if 

			vetor_quadro=media_turma	
			
						
elseif tipo_calculo="boletim" then	

	co_materia= split(vetor_materia,"#!#")	
	co_materia_check=0	

	vetor_periodo= split(periodo,"#!#")	
	maior_periodo=vetor_periodo(ubound(vetor_periodo))
	maior_periodo=maior_periodo*1
	
	if maior_periodo=5 then
		total_completa=0
	elseif maior_periodo=4 then
		total_completa=2		
	elseif maior_periodo=3 then
		total_completa=5	
	elseif maior_periodo=2 then
		total_completa=6
	elseif maior_periodo=1 then
		total_completa=10		
	end if		
	
	for cp=0 to total_completa
		if cp=0 then
		else
			completa_periodos=completa_periodos&"#!#"
		end if	
	next	

	For f2=0 to ubound(co_materia)
	
	soma=0	
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& co_materia(f2) &"'"
'RESPONSE.Write(SQL)
		RS.Open SQL, CON0
	
		mae= RS("IN_MAE")
		fil= RS("IN_FIL")
		in_co= RS("IN_CO")
		peso= RS("NU_Peso")
		
		calcula_media_anual="sim"	
		calcula_ms1="S"				
		calcula_ms3="S"				
		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) then
			for f2a=0 to ubound(vetor_periodo)
			periodo_cons=vetor_periodo(f2a)
			conceito=""	
				if periodo_cons= 1 or periodo_cons= 3 then
					acumula_ms=0		
					conta_per=0		
				end if				
			

				per=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, vetor_aluno, co_materia(f2), caminho_nota, tb_nota, periodo_cons)
				
				if per="&nbsp;" or per="" or isnull(per) then
					acumula_ms=acumula_ms	
					calcula_media_anual="nao"
					if periodo_cons= 2 THEN
						calcula_ms1="N"
					END IF					
					IF periodo_cons= 4 then
						calcula_ms3="N"					
					END IF
				else
					acumula_ms=acumula_ms+per
					conta_per=conta_per+1	
					
					per=arredonda(per,"mat",0,0)
					per=per/10	
					per=formatnumber(per,1)											
				end if					
							
				if co_materia_check=0 AND periodo_cons=1 then
					vetor_quadro=per
				elseif periodo_cons=1 then		
					vetor_quadro=vetor_quadro&per
				else
					vetor_quadro=vetor_quadro&"#!#"&per
				end if	
					
					
				if periodo_cons= 2 or periodo_cons= 4 then
					if conta_per=0 then
						conta_per=1
					end if
					if acumula_ms=0 OR (periodo_cons= 2  AND calcula_ms1="N") OR (periodo_cons= 4 AND calcula_ms3="N") then
						media="&nbsp;"
					else
						ms=acumula_ms/conta_per
						media=arredonda(ms,"mat",0,0)
						media=media/10	
						media=formatnumber(media,1)		
					end if				
					vetor_quadro=vetor_quadro&"#!#"&media

					if periodo_cons= 2 then
						recs=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, vetor_aluno, co_materia(f2), caminho_nota, tb_nota, "recs")
		
						if recs="&nbsp;" or isnull(recs) or recs="" then
							if isnumeric(media) then	
								media_rec=media
								media_rec=formatnumber(media_rec,1)	
							else
								media_rec=media							
							end if									
						else
							if isnumeric(media) then		
								media=media*10	
								recs=recs*1			
								if media>recs then
									media_rec=ms/10
									media_rec=formatnumber(media_rec,1)	
									
									recs=recs/10
									recs=formatnumber(recs,1)								
								else										
									media_rec=(ms+recs)/2

									media_rec=arredonda(media_rec,"mat",1,0)	
									media_rec=media_rec/10									
									media_rec=formatnumber(media_rec,1)			
									
									recs=arredonda(recs,"mat",0,0)
									recs=recs/10
									recs=formatnumber(recs,1)
								end if	
							else
								media_rec="&nbsp;"
							end if											
						end if	
						vetor_quadro=vetor_quadro&"#!#"&recs&"#!#"&media_rec	

					elseif periodo_cons= 4 then 	
						if calcula_media_anual="sim" then
							media_calc1=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, vetor_aluno, co_materia(f2), caminho_nota, tb_nota, 5, 4, 5, "anual", 0)	
							if media_calc1="&nbsp;#!#&nbsp;" then'
								media_anual=""
								resultado_anual=""	
							else	
								resultados=split(media_calc1,"#!#")
								media_anual=resultados(0)
								resultado_anual=resultados(1)
								if media_anual="&nbsp;" then
									media_anual=""
									resultado_anual=""
								else								
									media_anual=formatnumber(media_anual,1)		
								end if	
							end if						
						else
							media_anual=""
							resultado_anual=""
						end if		
						if co_materia_check=0 AND periodo_cons=1 then
							vetor_quadro=media_anual
						elseif periodo_cons=1 then		
							vetor_quadro=vetor_quadro&media_anual
						else
							vetor_quadro=vetor_quadro&"#!#"&media_anual
						end if	
					end if																						
				end if						
											
				if periodo_cons=5 then								
					if  media_anual="" or isnull(media_anual) then
						media_final=""
						resultado_final=""							
					else
						media_calc3=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, vetor_aluno, co_materia(f2), caminho_nota, tb_nota, 5, 4, 5, "final", 0)	
						if media_calc3="&nbsp;#!#&nbsp;" then'
							media_final=""
							resultado_final=""	
						else	
							resultados=split(media_calc3,"#!#")
							media_final=resultados(0)
							resultado_final=resultados(1)					
							media_final=formatnumber(media_final,1)	
						end if	
					end if					
					vetor_quadro=vetor_quadro&"#!#"&media_final								
				end if
			Next	
				
			for f2b=0 to ubound(vetor_periodo)
				periodo_cons=vetor_periodo(f2b)		
				if periodo_cons=1 then
					per_faltas="f1"			
				elseif periodo_cons=2 then
					per_faltas="f2"				
				elseif periodo_cons=3 then
					per_faltas="f3"				
				elseif periodo_cons=4 then
					per_faltas="f4"			
				end if
				if periodo_cons<>5 then
					faltas=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, vetor_aluno, co_materia(f2), caminho_nota, tb_nota, per_faltas)				
					if per_faltas="f1" then				
						vetor_faltas=faltas	
					else	
						vetor_faltas=vetor_faltas&"#!#"&faltas	
					end if	
				end if
					
			next
		vetor_quadro=vetor_quadro&"#!#"&completa_periodos&vetor_faltas&"#$#"	

		elseif (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) then

		elseif (mae=TRUE and fil=TRUE and in_co=FALSE) then
		
		elseif (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then
		vetor_quadro=vetor_quadro&"#$#"						
		end if		
	co_materia_check=co_materia_check+1			
	
'RESPONSE.Write(media_anual_mae&"-"&soma_mae&"-"&divisor_anual&"-"&co_materia_fil_check)					
'RESPONSE.END()	
	NEXT		
else	




end if
calcula_medias=vetor_quadro&"#$#"
'response.Write(calcula_medias)

'if vetor_aluno=20090022 then
'	RESPONSE.END()	
'end if
end function	

Function conta_medias(unidade, curso, co_etapa, turma, periodo, vetor_aluno, vetor_materia, caminho_nota, tb_nota, nome_nota, valor, operacao, outro, tipo_calculo)

'response.Write(vetor_materia&"<BR>")
'response.Write(vetor_aluno&"<BR>")
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CONn = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONn.Open ABRIRn		

	if operacoes(o)="menor" then
		operador=nome_nota&"<"&valor
	elseif operacoes(o)="maior" then
		operador=nome_nota&">="&valor
	elseif operacoes(o)="nulo" then
		operador="ISNULL("&nome_nota&")"
	end if	



if tipo_calculo="media_turma" then	
	
	co_materia= split(vetor_materia,"#!#")	
	co_materia_check=1	
	
	For i=0 to ubound(co_materia)
	
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& co_materia(i) &"'"
		RS.Open SQL, CON0
	
		mae= RS("IN_MAE")
		fil= RS("IN_FIL")
		in_co= RS("IN_CO")
		peso= RS("NU_Peso")
		
		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) then

			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT Count("&tb_nota&"."&nome_nota&")AS QtdDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& vetor_aluno &") AND CO_Materia ='"& co_materia(i)&"' And "&operador&" And NU_Periodo="&periodo
			RS1.Open SQL1, CONn
			
			media_turma=RS1("QtdDeVA_Media3")
			if media_turma="" or isnull(media_turma) then
			else
			media_turma=formatnumber(media_turma,0)
			end if 
				if co_materia_check=1 then
					vetor_quadro=media_turma
				else
					vetor_quadro=vetor_quadro&"#!#"&media_turma
				end if
				
		elseif (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) then
		
			vetor_mae_filhas=""
	
			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL2 = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_materia(i) &"' order by NU_Ordem_Boletim"
			RS2.Open SQL2, CON0
				
			co_materia_fil_check=0 
			while not RS2.EOF
				co_mat_fil= RS2("CO_Materia")				
				if co_materia_fil_check=0 then
					vetor_mae_filhas=co_materia(i)&"#!#"&co_mat_fil
				else
					vetor_mae_filhas=vetor_mae_filhas&"#!#"&co_mat_fil			
				end if
				co_materia_fil_check=co_materia_fil_check+1 									
			RS2.MOVENEXT
			wend				
			
			co_materia_mae_fil= split(vetor_mae_filhas,"#!#")
			media_mae_acumula=0			
			for j=0 to ubound(co_materia_mae_fil)			
		
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL3 = "SELECT Count("&tb_nota&"."&nome_nota&")AS QtdDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& vetor_aluno &") AND CO_Materia ='"& co_materia_mae_fil(j)&"' And "&operador&"  And NU_Periodo="&periodo
				RS3.Open SQL3, CONn

'response.Write(media_mae_acumula)					
				media_turma=RS3("QtdDeVA_Media3")
				if media_turma="" or isnull(media_turma) then
				media_filha_acumula=0	
				else
				media_turma=formatnumber(media_turma,0)
				media_filha_acumula=media_turma
				end if 

				if co_materia_check=1 then
					vetor_quadro=media_turma
					media_mae_acumula=media_mae_acumula+media_filha_acumula	
				else
					vetor_quadro=vetor_quadro&"#!#"&media_turma
					media_mae_acumula=media_mae_acumula*1
					media_turma=media_turma*1
					media_mae_acumula=media_mae_acumula+media_filha_acumula		
				end if		
				'response.Write(co_materia_mae_fil(j)&"-"&media_turma&"-"&media_mae_acumula&"-"&co_materia_fil_check&"<BR>")		
			next
			media_mae=media_mae_acumula/co_materia_fil_check
			media_mae=formatnumber(media_mae,0)
			vetor_quadro=vetor_quadro&"#!#"&media_mae	
		end if		
	co_materia_check=co_materia_check+1			
	NEXT
				
elseif tipo_calculo="nulo" then		

co_materia= split(vetor_materia,"#!#")	
	co_materia_check=1	
	co_aluno_check=1
	For i=0 to ubound(co_materia)

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& co_materia(i) &"'"
		RS.Open SQL, CON0
	
		mae= RS("IN_MAE")
		fil= RS("IN_FIL")
		in_co= RS("IN_CO")
		peso= RS("NU_Peso")
		
		aluno= split(vetor_aluno,",")	
		disciplina=co_materia(i)
		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) then
			conta_aluno=0	
			for n=0 to ubound(aluno)

				Set RS1 = Server.CreateObject("ADODB.Recordset")
				SQL1 = "SELECT * FROM "&tb_nota&" where CO_Matricula ="& aluno(n) &" AND CO_Materia ='"& disciplina &"' And NU_Periodo="&periodo
				RS1.Open SQL1, CONn
				
				if RS1.EOF then
				conta_aluno=conta_aluno+1
					if outro=disciplina then
						if conta_aluno=1 then
						aluno_nulo=aluno(n)
						else
						aluno_nulo=aluno_nulo&"#!#"&aluno(n)
						end if
					end if	
				else
					media_aluno=RS1("VA_Media3")
					if media_aluno="" or isnull(media_aluno) then
					conta_aluno=conta_aluno+1
						if outro=disciplina then
							if conta_aluno=1 then
							aluno_nulo=aluno(n)
							else
							aluno_nulo=aluno_nulo&"#!#"&aluno(n)
							end if
						end if	
					end if
				end if	
			Next	
			if co_materia_check=1 then
				vetor_quadro=conta_aluno
			else
				vetor_quadro=vetor_quadro&"#!#"&conta_aluno
			end if
		elseif (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) then
		
			vetor_mae_filhas=""
	
			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL2 = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_materia(i) &"' order by NU_Ordem_Boletim"
			RS2.Open SQL2, CON0
				
			co_materia_fil_check=0 
			while not RS2.EOF
				co_mat_fil= RS2("CO_Materia")				
				if co_materia_fil_check=0 then
					vetor_mae_filhas=co_materia(i)&"#!#"&co_mat_fil
				else
					vetor_mae_filhas=vetor_mae_filhas&"#!#"&co_mat_fil			
				end if
				co_materia_fil_check=co_materia_fil_check+1 									
			RS2.MOVENEXT
			wend				
			
			co_materia_mae_fil= split(vetor_mae_filhas,"#!#")
			media_mae_acumula=0						
			for j=0 to ubound(co_materia_mae_fil)			
				conta_aluno=0	
				disciplina_filha=co_materia_mae_fil(j)
				for n=0 to ubound(aluno)		
					Set RS3 = Server.CreateObject("ADODB.Recordset")
					SQL3 = "SELECT * FROM "&tb_nota&" where CO_Matricula ="& aluno(n) &" AND CO_Materia ='"&disciplina_filha &"' And NU_Periodo="&periodo
					RS3.Open SQL3, CONn
	
					if RS3.EOF then
					conta_aluno=conta_aluno+1
						if outro=disciplina_filha then
							if conta_aluno=1 then
							aluno_nulo=aluno(n)
							else
							aluno_nulo=aluno_nulo&"#!#"&aluno(n)
							end if
						end if	
					else
						media_aluno=RS3("VA_Media3")
						if media_aluno="" or isnull(media_aluno) then
						conta_aluno=conta_aluno+1	
							if outro=disciplina_filha then
								if conta_aluno=1 then
								aluno_nulo=aluno(n)
								else
								aluno_nulo=aluno_nulo&"#!#"&aluno(n)
								end if
							end if	
						end if
					end if					
				next
				if j=0 then
					vetor_quadro=vetor_quadro&"#!#"&conta_aluno
				else
					vetor_quadro=vetor_quadro&"#!#"&conta_aluno
					media_mae_acumula=media_mae_acumula*1
					conta_aluno=conta_aluno*1
					media_mae_acumula=media_mae_acumula+conta_aluno		
				end if		
				'response.Write(co_materia_mae_fil(j)&"-"&media_turma&"-"&media_mae_acumula&"-"&co_materia_fil_check&"<BR>")		
			next
			media_mae=media_mae_acumula/co_materia_fil_check
			media_mae=formatnumber(media_mae,0)
			vetor_quadro=vetor_quadro&"#!#"&media_mae	
		end if		
	co_materia_check=co_materia_check+1			
	NEXT
else
end if
Session("aluno_nulo")=aluno_nulo
conta_medias=vetor_quadro&"#$#"
'response.Write(calcula_medias)
end function	


'calcula as médias anuais e finais destes respectivos mapas
Function Calc_Med_An_Fin(unidade, curso, co_etapa, turma, vetor_aluno, vetor_materia, caminho_nota, tb_nota, qtd_periodos, periodo_m2, periodo_m3,tipo_calculo, outro)

	if  periodo_m2>0 then
		retira_periodo_m2=1
	else
		retira_periodo_m2=0			
	end if
	
	if periodo_m3>0 then
		retira_periodo_m3=1
	else
		retira_periodo_m3=0			
	end if
					
	'medias_necessarias=qtd_periodos-retira_periodo_m2-retira_periodo_m3
	medias_necessarias=qtd_periodos-1
'response.Write(vetor_materia&"<BR>")
'response.Write(vetor_aluno&"<BR>")
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CONn = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONn.Open ABRIRn	
		
	alunos= split(vetor_aluno,"#$#")			
	co_materia= split(vetor_materia,"#!#")	
	co_materia_check=1	
	co_matricula= vetor_aluno
	quantidade_alunos=0
	For a=0 to ubound(alunos)
'	response.Write(alunos(a))
		dados_aluno= split(alunos(a),"#!#")	
		quantidade_materias=0
		For c=0 to ubound(co_materia)
			Set RS = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& co_materia(c) &"'"
			RS.Open SQL, CON0
		
			mae= RS("IN_MAE")
			fil= RS("IN_FIL")
			in_co= RS("IN_CO")
			peso= RS("NU_Peso")		
			
			media_acumulada=0
			peso_periodo_acumulado=0
			qtd_medias=0

				'response.Write(mae&"-"&fil&"-"&in_co&"-"&peso&"<BR>")		
			if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) then
'response.Write("<BR>"&dados_aluno(1)&" CO_Matricula ="& dados_aluno(0)&"&"&co_materia(c) &"<BR>")	
				for periodo=1 to qtd_periodos
					Set RSn = Server.CreateObject("ADODB.Recordset")
					SQLn = "SELECT * FROM "&tb_nota&" where CO_Matricula ="& dados_aluno(0) &" AND CO_Materia ='"& co_materia(c) &"' AND CO_Materia_Principal ='"& co_materia(c) &"' AND NU_Periodo="&periodo				
					RSn.Open SQLn, CONn
	
						qtd_periodos=qtd_periodos*1
						periodo=periodo*1
						periodo_m2=periodo_m2*1				
						periodo_m3=periodo_m3*1		
					if RSn.EOF then					
						if periodo=1 then
							va_m31="&nbsp;"
							dividendo1=0
							divisor1=0
						elseif periodo=2 then
							va_m32="&nbsp;"
							va_rec_sem="&nbsp;"
							dividendo2=0
							dividendorec=0
							divisor2=0
							divisorrec=0
						elseif periodo=3 then
							va_m33="&nbsp;"
							dividendo3=0
							divisor3=0
						elseif periodo=4 then
							va_m34="&nbsp;"
							dividendo4=0
							divisor4=0
						elseif periodo=5 then
							va_m35="&nbsp;"
							media_rec=0
						end if						
						if periodo=periodo_m2 then
							rec_lancado="nao"
						end if
						if periodo=periodo_m3 then
							rec_lancado="nao"						
'							media_final=md
'							final_lancado="nao"							
						end if							
					else						
						Set RSPESO = Server.CreateObject("ADODB.Recordset")
						SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&periodo
						RSPESO.Open SQLPESO, CON0
						
						peso_periodo=RSPESO("NU_Peso")
						
						if peso_periodo=0 then
							peso_periodo=1
						end if

						if periodo=periodo_m2 then
							va_m34=RSn("VA_Media3")
							if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
								dividendo4=0
								divisor4=0
								qtd_medias=qtd_medias
							else
								dividendo4=va_m34*peso_periodo
								divisor4=peso_periodo
								qtd_medias=qtd_medias+1								
							end if								
'							va_m35=RSn("VA_Media3")
'					
'							rec_lancado="sim"
'							if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
'								media_rec="&nbsp;"
'							else
'								media_rec=va_m35*peso_periodo
'							end if															
						elseif periodo=periodo_m3 then
'							media_final=md
'							final_lancado="sim"		
							va_m35=RSn("VA_Media3")
										
							if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
								media_rec="&nbsp;"
								rec_lancado="nao"								
							else
								media_rec=(va_m35*peso_periodo)/10
								rec_lancado="sim"								
							end if			
'							response.Write(co_materia(c) &"_"&rec_lancado&"-"&	dados_aluno(0) &" mr "&media_rec&"<BR>")										
						else
							if periodo=1 then
								va_m31=RSn("VA_Media3")
								if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
									dividendo1=0
									divisor1=0
									qtd_medias=qtd_medias
								else
									dividendo1=va_m31*peso_periodo
									divisor1=peso_periodo
									qtd_medias=qtd_medias+1
								end if									
							elseif periodo=2 then
								va_m32=RSn("VA_Media3")
								va_rec_sem=RSn("VA_Rec")
								if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
									dividendo2=0
									divisor2=0
									qtd_medias=qtd_medias
								else
									dividendo2=va_m32*peso_periodo
									divisor2=peso_periodo
									qtd_medias=qtd_medias+1								
								end if
								
								if isnull(va_rec_sem) or va_rec_sem="&nbsp;"  or va_rec_sem="" then
									dividendorec=0
									divisorrec=0
								else
									dividendorec=va_rec_sem*peso_periodo
									divisorrec=peso_periodo
								end if								
							elseif periodo=3 then
								va_m33=RSn("VA_Media3")

								if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
									dividendo3=0
									divisor3=0
									qtd_medias=qtd_medias
								else
									dividendo3=va_m33*peso_periodo
									divisor3=peso_periodo
									qtd_medias=qtd_medias+1								
								end if								
'							elseif periodo=4 then
'								va_m34=RSn("VA_Media3")
'								if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
'									dividendo4=0
'									divisor4=0
'									qtd_medias=qtd_medias
'								else
'									dividendo4=va_m34*peso_periodo
'									divisor4=peso_periodo
'									qtd_medias=qtd_medias+1								
'								end if									
							end if													
						end if						
					end if
				Next

'response.Write("P1 "&dividendo1&"<BR>")
'response.Write("P2 "&dividendo2&"<BR>")


				dividendo_ms1=dividendo1+dividendo2
				divisor_ms1=divisor1+divisor2
										
				if divisor_ms1<2 then
					ms1="&nbsp;"
					dividendoms1=0
					divisorms1=0
				else
					ms1=dividendo_ms1/divisor_ms1
					ms1=arredonda(ms1,"mat",0,0)
'					media_final=media_final/10	
'					media_final=formatnumber(media_final,1)	
					dividendoms1=ms1
					divisorms1=1
					
				end if
'response.Write("MS1 "&ms1&"<BR>")				

			
				if divisorrec=0 then
					ms2=ms1
					if ms2="&nbsp;" then
						dividendoms2=0
						divisorms2=0
						dividendo_anual_ms2=0
						divisor_anual_ms2=0
					else
						dividendoms2=ms2
						divisorms2=1						
						dividendo_anual_ms2=ms2
						divisor_anual_ms2=1
					end if
				elseif ms1<>"&nbsp;" then
					dividendo_ms2=dividendoms1+dividendorec
					divisor_ms2=divisorms1+divisorrec																			
					ms2=dividendo_ms2/divisor_ms2			
ms2=ms2*1	
ms1=ms1*1							
					if ms2<ms1 then
						ms2=ms1								
					end if
					ms2=arredonda(ms2,"mat",0,0)															
					dividendo_anual_ms2=ms2
					divisor_anual_ms2=1
				end if
'response.Write("RS "&dividendorec&"<BR>")					
'response.Write("MS2 "&ms2&"<BR>")					
'response.Write("P3 "&dividendo3&"<BR>")
'response.Write("P4 "&dividendo4&"<BR>")				
				dividendo_ms3=dividendo3+dividendo4
				divisor_ms3=divisor3+divisor4
'response.Write(divisor_ms3&"="&divisor3&"+"&divisor4)			
				if divisor_ms3<2 then
					ms3="&nbsp;"
					dividendo_anual_ms3=0
					divisor_anual_ms3=0					
				else
					ms3=dividendo_ms3/divisor_ms3
					ms3=arredonda(ms3,"mat",0,0)							
				dividendo_anual_ms3=ms3
				divisor_anual_ms3=1						
				end if					
'response.Write("MS3 "&ms3&"<BR>")

	
				dividendo_anual_ms2=dividendo_anual_ms2*1
				dividendo_anual_ms3=dividendo_anual_ms3*1
				divisor_anual_ms2=divisor_anual_ms2*1
				divisor_anual_ms3=divisor_anual_ms3*1		
				dividendo_ma=dividendo_anual_ms2+dividendo_anual_ms3
				divisor_ma=divisor_anual_ms2+divisor_anual_ms3				
'response.Write(dividendo_ma&"="&dividendo_anual_ms2&"+"&dividendo_anual_ms3)						
				if peso_periodo_acumulado=0 then
					peso_periodo_acumulado=1
				end if	
'response.Write(	qtd_medias&">="&medias_necessarias&"<BR>")	
				if qtd_medias>=medias_necessarias then
					media_anual=dividendo_ma/divisor_ma								
					media_anual=media_anual/10	
					media_anual=formatnumber(media_anual,1)		
'response.Write(media_anual&"="&dividendo_ma&"/"&divisor_ma)			
					media_anual=media_anual*1						
					'if media_anual>67 and media_anual<70 then
					'	media_anual=70
					'end if	
'response.Write("MREC "&media_rec&"<BR>")						
					if tipo_calculo="anual" then
						resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","anual")
						media_anual = formatNumber(media_anual,1)
						resultado_materia=resultado					
'response.Write("resultado "&resultado&"<BR>")	
					elseif tipo_calculo="recuperacao" then
						if rec_lancado="nao" or media_rec="" or isnull(media_rec) then
						'verifica=1
							resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","anual")
							teste_resultado = split(resultado,"#!#")
							if teste_resultado(1) = "APR" or teste_resultado(1) = "Apr" or teste_resultado(1) = "REP" or teste_resultado(1) = "Rep" then 
								resultado_materia="#!#"
							else
								resultado_materia=resultado						
							end if
						else
						'verifica=2
							resultado=regra_aprovacao(curso,co_etapa,media_anual,media_rec,"&nbsp;","&nbsp;","&nbsp;","recuperacao")					
							resultado_materia=resultado
						end if	
				
					elseif tipo_calculo="final" then
						if rec_lancado="nao" or media_rec="" or isnull(media_rec) then
							resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","final")
							
'	response.Write(resultado&" res <BR>")		
							resultado_recuperacao= split(resultado,"#!#")
							if resultado_recuperacao(1)="APR" or resultado_recuperacao(1)="REP" then
								resultado_materia=resultado
							else
								resultado_materia="&nbsp;#!#&nbsp;"								
							end if
						
'						elseif final_lancado="nao" or media_final="" or isnull(media_final) then
'							resultado=regra_aprovacao(curso,co_etapa,media_anual,media_rec,"&nbsp;","&nbsp;","&nbsp;","final")
'							resultado_recuperacao= split(resultado,"#!#")
'							if resultado_recuperacao(1)="APR" or resultado_recuperacao(1)="REP" then
'								resultado_materia=resultado
'							else	
'								resultado_materia="&nbsp;#!#&nbsp;"								
'							end if							
						else
'if 	dados_aluno(0) =20040135 and co_materia(c) = "MAT1" then 
'response.Write(rec_lancado&"-"&curso&"-"&co_etapa&" ma "&media_anual&" rec "&media_rec&"<BR>")		
'end if						
							resultado=regra_aprovacao(curso,co_etapa,media_anual,media_rec,"&nbsp;",media_final,"&nbsp;","final")					
							resultado_materia=resultado
'if 	dados_aluno(0) =20040135 and co_materia(c) = "MAT1" then 
'response.Write("RES "&resultado_materia&"<BR>")		
'	response.end()
'end if								
						end if						
					end if	
				else
						resultado_materia="&nbsp;#!#&nbsp;"
				end if	

			end if	
			if quantidade_materias=0 then
				resultado_aluno=resultado_materia
				quantidade_materias=quantidade_materias+1
			else
				resultado_aluno=resultado_aluno&"#$#"&resultado_materia
				quantidade_materias=quantidade_materias+1
			end if
		next
		if quantidade_alunos=0 then
			resultado_turma=resultado_aluno
			quantidade_alunos=quantidade_alunos+1
		else
			resultado_turma=resultado_turma&"#%#"&resultado_aluno
			quantidade_alunos=quantidade_alunos+1
		end if

	next
'response.end()	
Calc_Med_An_Fin=resultado_turma		
END FUNCTION

Function regra_aprovacao (curso,etapa,m1_aluno,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,tipo_calculo)
'response.Write(m1_aluno&"-"&nota_aux_m2_1&"|")
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set RSra = Server.CreateObject("ADODB.Recordset")
	SQLra = "SELECT * FROM TB_Regras_Aprovacao where CO_Curso = '"&curso&"' and CO_Etapa = '"&etapa&"'"
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
	m3_menor=RSra("NU_Int_Me_Ma_Igual_M3")
	m3_maior_igual=RSra("NU_Int_Me_Me_M3")	
	res3_1=RSra("NO_Expr_Int_M3_F")
	res3_2=RSra("NO_Expr_Ma_Igual_M3")

		
	m1_aluno=m1_aluno*1	
	m1_maior_igual=m1_maior_igual*1
	m1_menor=m1_menor*1

'response.Write(m1_aluno&">="&m1_maior_igual&" "&m1_aluno&">="&m1_menor&"<BR>")
	if m1_aluno >= m1_maior_igual then
		resultado=res1_3
		resultado1="apr"
	elseif m1_aluno >= m1_menor then
		resultado=res1_2
	else
		resultado=res1_1
		resultado1="rep"			
	end if
	
	if tipo_calculo="waboletim" then
		m1_waboletim=m1_aluno
		resultado1_waboletim=resultado
	end if	
'response.Write("if "&m1_aluno &">="& m1_maior_igual &"then<BR>")
'response.Write("elseif "&m1_aluno &">"& m1_menor &"then<BR>")	
'response.Write(resultado&"<BR>")	
	if resultado1="apr" or resultado1="rep" then
		m2_aluno=m1_aluno	
		m3_aluno=m1_aluno
		if tipo_calculo="waboletim" then
			m2_waboletim="&nbsp;"
			m3_waboletim="&nbsp;"			
			resultado2_waboletim="&nbsp;"
			resultado3_waboletim="&nbsp;"		
		end if		
		
	else			
		if tipo_calculo="recuperacao" or tipo_calculo="final" or tipo_calculo="waboletim" then
			if nota_aux_m2_1="&nbsp;" then
				m2_aluno="&nbsp;"
				resultado="&nbsp;"	
				if tipo_calculo="waboletim" then
					m2_waboletim=m2_aluno	
					resultado2_waboletim=resultado	
				end if	
			else								
				m1_aluno_peso=m1_aluno*peso_m2_m1
				nota_aux_m2_1_peso=nota_aux_m2_1*peso_m2_m2
				m2_aluno=(m1_aluno_peso+nota_aux_m2_1_peso)/(peso_m2_m1+peso_m2_m2)
'response.Write(m1_aluno_peso&"+"&nota_aux_m2_1_peso&"/"&peso_m2_m1&"+"&peso_m2_m2&"<BR>")

				m2_aluno=arredonda(m2_aluno,"mat_dez",1,0)									
				m2_aluno=formatnumber(m2_aluno,1)	
				
				m2_aluno=m2_aluno*1
				m2_maior_igual=m2_maior_igual*1	
				m2_menor=m2_menor*1		
				if m2_aluno >= m2_maior_igual then
					resultado=res2_3
					resultado2="apr"
				elseif m2_aluno >= m2_menor then
					resultado=res2_2
					resultado2="rep"						
				else
					resultado=res2_1	
					resultado2="rep"					
				end if

				if tipo_calculo="waboletim" then
					m2_waboletim=m2_aluno		
					resultado2_waboletim=resultado	
				end if	

'response.Write("if "&m2_aluno &">="& m2_maior_igual &"then<BR>")
'response.Write("elseif "&m2_aluno &">"& m2_menor &"then<BR>")	
'response.Write(resultado&"<BR>")

			end if
' 			if	tipo_calculo="final" or tipo_calculo="waboletim" then
'				if resultado2="apr" or resultado2="rep" then
'					m3_aluno=m2_aluno				
'				else
'					if m2_aluno="&nbsp;" or nota_aux_m2_1="&nbsp;" or nota_aux_m3_1="&nbsp;" then		
'						m3_aluno="&nbsp;"
'						resultado="&nbsp;"			
'					else								
'						m1_aluno_peso=m1_aluno*peso_m3_m1
'						m2_aluno_peso=m2_aluno*peso_m3_m2
'						nota_aux_m3_1_peso=nota_aux_m3_1*peso_m3_m3
''response.Write(m3_aluno&"=("&m1_aluno_peso&"+"&m2_aluno_peso&"+"&nota_aux_m3_1_peso&")/("&peso_m3_m1&"+"&peso_m3_m2&"+"&peso_m3_m3)
'
'						m3_aluno=(m1_aluno_peso+m2_aluno_peso+nota_aux_m3_1_peso)/(peso_m3_m1+peso_m3_m2+peso_m3_m3)
'						m3_aluno=arredonda(m3_aluno,"mat_dez",1,0)									
'						m3_aluno=formatnumber(m3_aluno,1)	
'						m3_aluno=m3_aluno*1
'						valor_m3=valor_m3*1		
'						m3_maior_igual=m3_maior_igual*1		
'						if m3_aluno >= m3_maior_igual then
'							resultado=res3_2
'						else
'							resultado=res3_1	
'						end if
'						if tipo_calculo="waboletim" then
'							m3_waboletim=m3_aluno		
'							resultado3_waboletim=resultado	
'						end if							
'					end if		
'				end if
'			end if	
		end if	
	end if

	if tipo_calculo="anual" then
		m1_aluno = formatNumber(m1_aluno,1)	
		regra_aprovacao=m1_aluno&"#!#"&resultado
	elseif tipo_calculo="recuperacao" then
		if resultado1="apr" or resultado1="rep" then
			m1_aluno = formatNumber(m1_aluno,1)	
			regra_aprovacao=m1_aluno&"#!#"&resultado		
		else
			if m2_aluno<>"&nbsp;" then
				m2_aluno = formatNumber(m2_aluno,1)
			end if
			regra_aprovacao=m2_aluno&"#!#"&resultado
		end if
	elseif tipo_calculo="waboletim" then
			if m2_aluno<>"&nbsp;" then
				m2_aluno = formatNumber(m2_aluno,1)
			end if	

			if m3_aluno<>"&nbsp;" then
				m3_aluno = formatNumber(m3_aluno,1)
			end if
		regra_aprovacao=m1_waboletim&"#!#"&resultado1_waboletim&"#!#"&m2_waboletim&"#!#"&resultado2_waboletim&"#!#"&m3_waboletim&"#!#"&resultado3_waboletim
	else
		if resultado2="apr" or resultado2="rep" then
			if m2_aluno<>"&nbsp;" then
				m2_aluno = formatNumber(m2_aluno,1)
			end if
			regra_aprovacao=m2_aluno&"#!#"&resultado		
		else
			if m3_aluno<>"&nbsp;" then
				m3_aluno = formatNumber(m3_aluno,1)
			end if
			regra_aprovacao=m3_aluno&"#!#"&resultado			
		end if
	end if
	
	'Session("M2")=m2_aluno
	'Session("M3")=m3_aluno
end function

Function apura_resultado_aluno (curso,etapa,vetor_medias)

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set RSra = Server.CreateObject("ADODB.Recordset")
	SQLra = "SELECT * FROM TB_Regras_Aprovacao where CO_Curso = '"&curso&"' and CO_Etapa = '"&etapa&"'"
	RSra.Open SQLra, CON0	
			
	valor_apr=RSra("NU_Valor_Apr")
	valor_dep=RSra("NU_Valor_Dep")
	qtd_max_dep=RSra("NU_Qt_Dis_Dep")
	res_apr=RSra("NO_Expr_Maior_Igual_VL_Abr")
	res_dep=RSra("NO_Expr_Cond_Verdade_Abr")
	res_rep=RSra("NO_Expr_Cond_Falso_Abr")
	qtd_dep=0
	
'	valor_apr=70
'	valor_dep=50
'	qtd_max_dep=5
'	res_apr="AP"
'	res_dep="DP"
'	res_rep="RP"	

	resultados_materia = split(vetor_medias, "#$#" )
	libera_resultado="s"
for r=0 to ubound(resultados_materia)	
	nota_materia = split(resultados_materia(r), "#!#" )

	md_aluno=nota_materia(0)
	valor_apr=valor_apr*1
	valor_dep=valor_dep*1
	if md_aluno="" or isnull(md_aluno) or md_aluno="&nbsp;" or md_aluno=" "then
		libera_resultado="n"
	else
'response.Write(md_aluno&">="&valor_apr&";"&valor_dep&"<BR>")
		md_aluno=md_aluno*1
		if md_aluno >= valor_apr then
			res_aluno="apr"
		elseif md_aluno >= valor_dep then
			res_aluno="dep"
		else
			res_aluno="rep"			
		end if	
'response.Write(res_aluno&"<BR>")				
		if result_temp="rep" then
		else
			if res_aluno="" or isnull(res_aluno) or res_aluno="&nbsp;" or res_aluno=" "then
				libera_resultado="n"
			else
				result_temp=res_aluno
				if res_aluno = "dep" then
					qtd_dep=qtd_dep+1		
				end if
			end if
		end if	
'response.Write(result_temp&"<BR>")					
	end if
Next
if 	libera_resultado="s" then
'response.Write(result_temp&"<BR>")	
'response.End()
	qtd_dep=qtd_dep*1
	if qtd_dep>0 then
		qtd_max_dep=qtd_max_dep*1
		if qtd_dep>qtd_max_dep then
			apura_resultado_aluno=res_rep	
		else	
			apura_resultado_aluno=res_dep	
		end if	
	elseif result_temp="apr" then
		apura_resultado_aluno=res_apr
	elseif result_temp="rep" then
		apura_resultado_aluno=res_rep
	end if	
else
	apura_resultado_aluno="&nbsp;"		
end if	
end function	

Function replace_latin_char(variavel,tipo_replace)

	if tipo_replace="html" then
		strReplacement = variavel	
		strReplacement = replace(strReplacement,"À,","&Agrave;")
		strReplacement = replace(strReplacement,"Á","&Aacute;")
		strReplacement = replace(strReplacement,"Â","&Acirc;")
		strReplacement = replace(strReplacement,"Ã","&Atilde;")
		strReplacement = replace(strReplacement,"É","&Eacute;")
		strReplacement = replace(strReplacement,"Ê","&Ecirc;")
		strReplacement = replace(strReplacement,"Í","&Iacute;")
		strReplacement = replace(strReplacement,"Ó","&Oacute;")
		strReplacement = replace(strReplacement,"Ô","&Ocirc;")
		strReplacement = replace(strReplacement,"Õ","&Otilde;")
		strReplacement = replace(strReplacement,"Ú","&Uacute;")
		strReplacement = replace(strReplacement,"Ü","&Uuml;")	
		strReplacement = replace(strReplacement,"à","&agrave;")
		strReplacement = replace(strReplacement,"á","&aacute;")
		strReplacement = replace(strReplacement,"â","&acirc;")
		strReplacement = replace(strReplacement,"ã","&atilde;")
		strReplacement = replace(strReplacement,"ç","&ccedil;")
		strReplacement = replace(strReplacement,"é","&eacute;")
		strReplacement = replace(strReplacement,"ê","&ecirc;")
		strReplacement = replace(strReplacement,"í","&iacute;")
		strReplacement = replace(strReplacement,"ó","&oacute;")
		strReplacement = replace(strReplacement,"ô","&ocirc;")
		strReplacement = replace(strReplacement,"õ","&otilde;")
		strReplacement = replace(strReplacement,"ú","&uacute;")
		strReplacement = replace(strReplacement,"ü","&uuml;")			
	elseif tipo_replace="url" then
		strReplacement = Server.URLEncode(variavel)
		strReplacement = replace(strReplacement,"+"," ")
		strReplacement = replace(strReplacement,"%27","´")
		strReplacement = replace(strReplacement,"%27","'")
		strReplacement = replace(strReplacement,"%C0,","À")
		strReplacement = replace(strReplacement,"%C1","Á")
		strReplacement = replace(strReplacement,"%C2","Â")
		strReplacement = replace(strReplacement,"%C3","Ã")
		strReplacement = replace(strReplacement,"%C9","É")
		strReplacement = replace(strReplacement,"%CA","Ê")
		strReplacement = replace(strReplacement,"%CD","Í")
		strReplacement = replace(strReplacement,"%D3","Ó")
		strReplacement = replace(strReplacement,"%D4","Ô")
		strReplacement = replace(strReplacement,"%D5","Õ")
		strReplacement = replace(strReplacement,"%DA","Ú")
		strReplacement = replace(strReplacement,"%DC","Ü")	
		strReplacement = replace(strReplacement,"%E1","à")
		strReplacement = replace(strReplacement,"%E1","á")
		strReplacement = replace(strReplacement,"%E2","â")
		strReplacement = replace(strReplacement,"%E3","ã")
		strReplacement = replace(strReplacement,"%E7","ç")
		strReplacement = replace(strReplacement,"%E9","é")
		strReplacement = replace(strReplacement,"%EA","ê")
		strReplacement = replace(strReplacement,"%ED","í")
		strReplacement = replace(strReplacement,"%F3","ó")
		strReplacement = replace(strReplacement,"F4","ô")
		strReplacement = replace(strReplacement,"F5","õ")
		strReplacement = replace(strReplacement,"%FA","ú")
		strReplacement = replace(strReplacement,"%FC","ü")
	end if
replace_latin_char=strReplacement
end function	
%>