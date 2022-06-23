<!--#include file="../../global/funcoes_diversas.asp" -->
<!--#include file="parametros.asp" -->
<!--#include file="calculos.asp"-->
<!--#include file="resultados.asp"-->
<%

Function tabela_notas(CON, unidade, curso, co_etapa, turma, periodo, disciplina, outro)

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidade &" And CO_Curso= '"& curso &"' And CO_Etapa = '"& co_etapa &"'"
		Set RS = CON.Execute(SQL)
			
		if RS.EOF then
			tabela_notas = "ERRO"	
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
		elseif tb_nota="TB_NOTA_F" then
			CAMINHO_n=CAMINHO_nf
		elseif tb_nota="TB_NOTA_K" then
			CAMINHO_n=CAMINHO_nk		
		elseif tb_nota="TB_NOTA_L" then
			CAMINHO_n=CAMINHO_nl		
 		elseif tb_nota="TB_NOTA_M" then
			CAMINHO_n=CAMINHO_nm					
		elseif tb_nota="TB_NOTA_V" then
			CAMINHO_n=CAMINHO_nv						
		else
			CAMINHO_n = "ERRO"					
		end if
	
	caminho_notas=CAMINHO_n

end function

Function carga_aula(curso,co_etapa,vetor_materia,CONEXAO,tipo_resultado,outro)

	Set RS3 = Server.CreateObject("ADODB.Recordset")
	SQL3 = "SELECT SUM(NU_Aulas) as Carga FROM TB_Programa_Aula where CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Materia in ("&vetor_materia&")"
	RS3.Open SQL3, CONEXAO	

	if tipo_resultado = "frequencia" then
		carga_aula=200
	else	
		if RS3.EOF then
			carga_aula=200
		else
			carga_aula= RS3("Carga")			
		end if	
	end if	
End Function

Function converte_conceito(unidade, curso, co_etapa, turma, periodo, disciplina, media, outro)

	if curso=1 and co_etapa<6 and (disciplina="ARTC" or disciplina="EART" or disciplina="EFIS" or disciplina="INGL") then									
		teste_media = isnumeric(media)							
		if teste_media=TRUE then							
			if media > 90 then
			conceito="E"
			elseif (media > 70) and (media <= 90) then
			conceito="MB"
			elseif (media > 60) and (media <= 70) then							
			conceito="B"
			elseif (media > 49) and (media <= 60) then
			conceito="R"
			else							
			conceito="I"
			end if	
		end if	
	else
		conceito=media				
	end if	
converte_conceito=conceito
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

'===========================================================================================================================================
Function tipo_materia(co_materia, curso, co_etapa)

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& co_materia &"'"
		'response.write(SQL&"<BR>")
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
		elseif (mae=FALSE and fil=FALSE and in_co=TRUE) then
			tipo_materia="F_F_T_P"			
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

Function disciplina_obrigatoria(codigo_materia,CONEXAO,outro)
	Set RS3a = Server.CreateObject("ADODB.Recordset")
	SQL3a = "SELECT * FROM TB_Materia where CO_Materia ='"& codigo_materia &"' order by NU_Ordem_Boletim"
	RS3a.Open SQL3a, CONEXAO	
	if RS3a.EOF then
		disc_obrigat="s"
	else
		'ind_obr= RS3a("IN_Obrigatorio")	
		ind_obr=TRUE
		if ind_obr=TRUE then
			Disciplina_Obrigatoria="S"
		else
			Disciplina_Obrigatoria="N"
		end if			
	end if	
End Function

Function disciplina_regular(codigo_materia, curso, co_etapa, CONEXAO)

	Set RS3a = Server.CreateObject("ADODB.Recordset")
	SQL3a = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& codigo_materia &"'"
	RS3a.Open SQL3a, CONEXAO

	if RS3a.EOF then
		disciplina_regular="s"
	else
		ind_reg= RS3a("TP_Disciplina")	
		if ind_reg="R" then
			disciplina_regular="S"
		else
			disciplina_regular="N"
		end if			
	end if	
End Function


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
		
		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) or (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso))or (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then
	
			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_materia(i) &"' order by NU_Ordem_Boletim"
			RS1.Open SQL1, CON0
				
			if RS1.EOF then
				if co_materia_check=1 then
					vetor_materia_exibe=co_materia(i)
				else
					vetor_materia_exibe=vetor_materia_exibe&"#!#"&co_materia(i)
				end if
			else
				if mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso) then
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
						co_materia_fil_check=co_materia_fil_check+1 									
					RS1.MOVENEXT
					wend
					vetor_materia_exibe=vetor_materia_exibe&"#!#MED"	
				else
					if co_materia_check=1 then
						vetor_materia_exibe=co_materia(i)
					else
						vetor_materia_exibe=vetor_materia_exibe&"#!#"&co_materia(i)
					end if
				end if
			end if
		end if	
		co_materia_check=co_materia_check+1				
	NEXT
programa_aula=vetor_materia_exibe
else
end if
end function

Function programa_aula_boletim_ficha(vetor_materia, unidade, curso, co_etapa, turma)

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
		
		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) or (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso))or (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then
	
			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_materia(i) &"' order by NU_Ordem_Boletim"
			RS1.Open SQL1, CON0
				
			if RS1.EOF then
				if co_materia_check=1 then
					vetor_materia_exibe=co_materia(i)
				else
					vetor_materia_exibe=vetor_materia_exibe&"#!#"&co_materia(i)
				end if
			else
				ano_letivo=ano_letivo*1
				ano_letivo_prog_aula = ano_letivo_prog_aula*1

				if mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso) and ano_letivo<ano_letivo_prog_aula then
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
						co_materia_fil_check=co_materia_fil_check+1 									
					RS1.MOVENEXT
					wend
					vetor_materia_exibe=vetor_materia_exibe&"#!#MED"	
				else
					if co_materia_check=1 then
						vetor_materia_exibe=co_materia(i)
					else
						vetor_materia_exibe=vetor_materia_exibe&"#!#"&co_materia(i)
					end if
				end if
			end if
		end if	
		co_materia_check=co_materia_check+1				
	NEXT
programa_aula_boletim_ficha=vetor_materia_exibe
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
	
		if co_materia(i)<>"MED" Then
	
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
				calcula_media="s"						
				for j=0 to ubound(co_materia_mae_fil)			
			
					Set RS3 = Server.CreateObject("ADODB.Recordset")
					SQL3 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& vetor_aluno &") AND CO_Materia ='"& co_materia_mae_fil(j)&"' And NU_Periodo="&periodo
					RS3.Open SQL3, CONn
	
	'response.Write(media_mae_acumula)					
					media_turma=RS3("MediaDeVA_Media3")
					
					if media_turma="" or isnull(media_turma) then
						media_filha_acumula=0
						if j>0 then
							calcula_media="n"
						end if	
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
					co_materia_check=co_materia_check+1							
				next
				if calcula_media="s" then			
					media_mae=media_mae_acumula/co_materia_fil_check
					media_mae=formatnumber(media_mae,0)
				else
					media_mae=""			
				end if	
				vetor_quadro=vetor_quadro&"#!#"&media_mae	
				
			elseif (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso))then
			
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
				calcula_media="s"						
				for j=0 to ubound(co_materia_mae_fil)			
			
					Set RS3 = Server.CreateObject("ADODB.Recordset")
					SQL3 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& vetor_aluno &") AND CO_Materia ='"& co_materia_mae_fil(j)&"' And NU_Periodo="&periodo
					RS3.Open SQL3, CONn
	
	'response.Write(media_mae_acumula)					
					media_turma=RS3("MediaDeVA_Media3")
					
					if media_turma="" or isnull(media_turma) then
						media_filha_acumula=0
						if j>0 then
							calcula_media="n"
						end if	
					else
						media_turma=formatnumber(media_turma,0)
						media_filha_acumula=media_turma
					end if 		
					if co_materia_check=1 then
						'vetor_quadro=media_turma
						media_mae_acumula=media_mae_acumula+media_filha_acumula	
					else						
						'vetor_quadro=vetor_quadro&"#!#"&media_turma
						media_mae_acumula=media_mae_acumula*1
						media_turma=media_turma*1
						media_mae_acumula=media_mae_acumula+media_filha_acumula		
					end if						
					'response.Write(co_materia_mae_fil(j)&"-"&media_turma&"-"&media_mae_acumula&"-"&co_materia_fil_check&"<BR>")
					co_materia_check=co_materia_check+1							
				next
				if calcula_media="s" then	
					if media_mae_acumula>100 then
						media_mae_acumula=100
					end if			
					media_mae=media_mae_acumula
					media_mae=formatnumber(media_mae,0)
				else
					media_mae=""			
				end if	
				vetor_quadro=vetor_quadro&"#!#"&media_mae				
			end if		
		co_materia_check=co_materia_check+1	
		end if	
	NEXT
	
elseif tipo_calculo="media_geral" then	

	
'				Set RS1 = Server.CreateObject("ADODB.Recordset")
'			SQL1 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& vetor_aluno &") And NU_Periodo="&periodo
'			RS1.Open SQL1, CONn
'			
'			media_turma=RS1("MediaDeVA_Media3")
'			if media_turma="" or isnull(media_turma) then
'			media_turma=0
'			else
'			media_turma=formatnumber(media_turma,0)
'			end if 

	co_materia= split(vetor_materia,"#!#")	
	co_materia_check=1	
	acumulaNotas=0
	contaNotas=0
	For i=0 to ubound(co_materia)
	
		if co_materia(i)<>"MED" Then
	
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
					acumulaNotas=acumulaNotas*1
					media_turma=media_turma*1
					acumulaNotas = acumulaNotas+media_turma
					contaNotas = contaNotas+1
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
				calcula_media="s"						
				for j=0 to ubound(co_materia_mae_fil)			
			
					Set RS3 = Server.CreateObject("ADODB.Recordset")
					SQL3 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& vetor_aluno &") AND CO_Materia ='"& co_materia_mae_fil(j)&"' And NU_Periodo="&periodo			
					RS3.Open SQL3, CONn
	
	'response.Write(media_mae_acumula)					
					media_turma=RS3("MediaDeVA_Media3")
					
					if media_turma="" or isnull(media_turma) then
						media_filha_acumula=0
						if j>0 then
							calcula_media="n"
						end if	
					else
						media_turma=formatnumber(media_turma,0)
						media_filha_acumula=media_turma
					end if 		
					if co_materia_check=1 then
						media_mae_acumula=media_mae_acumula+media_filha_acumula	
					else						
						media_mae_acumula=media_mae_acumula*1
						media_turma=media_turma*1
						media_mae_acumula=media_mae_acumula+media_filha_acumula		
					end if						

					co_materia_check=co_materia_check+1							
				next
				if calcula_media="s" then	
			
					media_mae=media_mae_acumula/co_materia_fil_check
					media_mae=formatnumber(media_mae,0)
					acumulaNotas=acumulaNotas*1
					media_mae=media_mae*1							
					acumulaNotas = acumulaNotas+media_mae
					contaNotas = contaNotas+1	
				end if	
				
			elseif (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso))then
			
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
				calcula_media="s"						
				for j=0 to ubound(co_materia_mae_fil)			
			
					Set RS3 = Server.CreateObject("ADODB.Recordset")
					SQL3 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& vetor_aluno &") AND CO_Materia ='"& co_materia_mae_fil(j)&"' And NU_Periodo="&periodo				
					RS3.Open SQL3, CONn
	
	'response.Write(media_mae_acumula)					
					media_turma=RS3("MediaDeVA_Media3")
					
					if media_turma="" or isnull(media_turma) then
						media_filha_acumula=0
						if j>0 then
							calcula_media="n"
						end if	
					else
						media_turma=formatnumber(media_turma,0)
						media_filha_acumula=media_turma
					end if 		
					if co_materia_check=1 then
						media_mae_acumula=media_mae_acumula+media_filha_acumula	
					else						
						media_mae_acumula=media_mae_acumula*1
						media_turma=media_turma*1
						media_mae_acumula=media_mae_acumula+media_filha_acumula		
					end if						

					co_materia_check=co_materia_check+1							
				next
				if calcula_media="s" then			
					if media_mae_acumula>100 then
						media_mae_acumula=100
					end if					
					media_mae=media_mae_acumula
					media_mae=formatnumber(media_mae,0)
					acumulaNotas=acumulaNotas*1
					media_mae=media_mae*1							
					acumulaNotas = acumulaNotas+media_mae
					contaNotas = contaNotas+1			
				end if	
					
			end if		
		co_materia_check=co_materia_check+1	
		end if	
	NEXT

	if contaNotas=0 then
		vetor_quadro=0
	else
		vetor_quadro=formatnumber(acumulaNotas/contaNotas,0)	
	end if



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


elseif tipo_calculo="boletim" then	
	
	co_materia= split(vetor_materia,"#!#")	
	co_materia_check=0	

	vetor_periodo= split(periodo,"#!#")	
	PERIODO_ANUAL = Periodo_Media("T","MA",outro)
	PERIODO_RECUPERACAO=Periodo_Media("T","RF",outro)
	PERIODO_FINAL=Periodo_Media("T","MF",outro)
	'response.write(vetor_materia&"TFTN<BR>")			

	For f2=0 to ubound(co_materia)
		aproxima_m1	 = parametros_gerais(unidade, curso, etapa, turma, co_materia(f2),"aproxima_m1",0)	
		compara_m2 = parametros_gerais(unidade, curso, etapa, turma, co_materia(f2),"compara_m2",0)					
		aproxima_m2 = parametros_gerais(unidade, curso, etapa, turma, co_materia(f2),"aproxima_m2",0)		
		compara_m3 = parametros_gerais(unidade, curso, etapa, turma, co_materia(f2),"compara_m3",0)					
		aproxima_m3 = parametros_gerais(unidade, curso, etapa, turma, co_materia(f2),"aproxima_m3",0)	
		
		soma=0	
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& co_materia(f2) &"'"
'RESPONSE.Write(">>>>>>>>>>>>>>>>>>>>"&SQL&"<BR>")
		RS.Open SQL, CON0
	
		mae= RS("IN_MAE")
		fil= RS("IN_FIL")
		in_co= RS("IN_CO")
		peso= RS("NU_Peso")
		
		calcula_media_anual="sim"	
		media_calc1=""
		media_calc2=""
		media_calc3=""
		media_anual=""
		resultado_anual=""	
		media_recup=""
		resultado_recup=""	
		media_final=""
		resultado_final=""		
        media_anual_mae = ""			
		conceito_anual	=""
		conceito_recup	=""	
		conceito_final	=""			
		res1=""
		res2=""
		res3=""
									
		'response.write(co_materia(f2)&" mae= "&mae&" and fil= "&fil&" and in_co= "&in_co&" and peso = "&peso&"<BR>")
		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) then
				
				linha_materia = CALCULA_LINHA_BOLETIM(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, vetor_aluno, co_materia(f2), co_materia(f2), CONn , tb_nota, periodo, nome_nota, PERIODO_ANUAL, PERIODO_RECUPERACAO, PERIODO_FINAL, "T_F_F_N", outro)				
						
				if co_materia_check=0 then
					vetor_quadro=linha_materia
				else
					vetor_quadro=vetor_quadro&linha_materia
				end if				
				
			vetor_quadro=vetor_quadro&"#$#"
			'response.write(vetor_quadro)
'response.end()
		elseif (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) then

			vetor_mae_filhas=""
			disciplina_mae = co_materia(f2)
			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL2 = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_materia(f2) &"' order by NU_Ordem_Boletim"
			RS2.Open SQL2, CON0
				
			co_materia_fil_check=0 
			while not RS2.EOF
				co_mat_fil= RS2("CO_Materia")				
				if co_materia_fil_check=0 then
					vetor_mae_filhas=co_materia(f2)&"#!#"&co_mat_fil
				else
					vetor_mae_filhas=vetor_mae_filhas&"#!#"&co_mat_fil			
				end if
				co_materia_fil_check=co_materia_fil_check+1 									
			RS2.MOVENEXT
			wend	
			'response.write(vetor_mae_filhas)
			'response.end

			co_materia_fil_check=0 			
			co_materia_mae_fil= split(vetor_mae_filhas,"#!#")
			soma_mae=0			
			peso_acumula=0
			
			maior_periodo=vetor_periodo(ubound(vetor_periodo))
	
			Set RSpa = Server.CreateObject("ADODB.Recordset")
			SQLpa= "SELECT * FROM TB_Programa_Aula where CO_Materia='"&co_materia_mae_fil(f3)&"' AND CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"'"
			'response.Write(SQLpa&"<BR>")
			RSpa.Open SQLpa, CON0
										
			nu_peso_fil=RSpa("NU_Peso")	
			'response.Write(nu_peso_fil&"<BR>")					
			if isnull(nu_peso_fil) or nu_peso_fil="" then
				nu_peso_fil=1
			end if	
			peso_acumula=peso_acumula*1
			nu_peso_fil=nu_peso_fil*1
			peso_acumula=peso_acumula+nu_peso_fil						
			
			linha_materia = CALCULA_LINHA_BOLETIM(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, vetor_aluno, co_materia(f2), vetor_mae_filhas, CONn , tb_nota, periodo, nome_nota, PERIODO_ANUAL, PERIODO_RECUPERACAO, PERIODO_FINAL, "T_T_F_N", outro)		
		
			if co_materia_check=0 then
				vetor_quadro=linha_materia
			else
				vetor_quadro=vetor_quadro&linha_materia
			end if	
				

			vetor_quadro=vetor_quadro&"#$#"			
		elseif (mae=TRUE and fil=TRUE and in_co=FALSE) then
		
		elseif (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then	

			vetor_mat_filhas_tftn=busca_materias_filhas(co_materia(f2))
		
			linha_materia = CALCULA_LINHA_BOLETIM(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, vetor_aluno, co_materia(f2), vetor_mat_filhas_tftn, CONn , tb_nota, periodo, nome_nota, PERIODO_ANUAL, PERIODO_RECUPERACAO, PERIODO_FINAL, "T_F_T_N", outro)				
								
			if co_materia_check=0 then
				vetor_quadro=linha_materia
			else
				vetor_quadro=vetor_quadro&linha_materia
			end if				
				
			vetor_quadro=vetor_quadro&"#$#"										
		end if		
	co_materia_check=co_materia_check+1			
	NEXT	
else	
end if
'RESPONSE.END()	
calcula_medias=vetor_quadro&"#$#"
'response.Write(calcula_medias)
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

valor=replace(valor,",",".")
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
		elseif (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then
		
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
					'vetor_quadro=media_turma
					media_mae_acumula=media_mae_acumula+media_filha_acumula	
				else
					'vetor_quadro=vetor_quadro&"#!#"&media_turma
					media_mae_acumula=media_mae_acumula*1
					media_turma=media_turma*1
					media_mae_acumula=media_mae_acumula+media_filha_acumula		
				end if		
				'response.Write(co_materia_mae_fil(j)&"-"&media_turma&"-"&media_mae_acumula&"-"&co_materia_fil_check&"<BR>")		
			next
			if media_mae_acumula>100 then
				media_mae_acumula=100
			end if				
			media_mae=media_mae_acumula
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
		elseif (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then
		
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
			if media_mae_acumula>100 then
				media_mae_acumula=100
			end if				
			media_mae=media_mae_acumula
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


'calcula as m�dias anuais e finais destes respectivos mapas
Function Calc_Med_An_Fin(unidade, curso, co_etapa, turma, vetor_aluno, vetor_materia, caminho_nota, tb_nota, qtd_periodos, periodo_m2, periodo_m3,tipo_calculo, outro)

	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR

	Set RSano = Server.CreateObject("ADODB.Recordset")
	SQLano = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&ano_letivo&"'"
	RSano.Open SQLano, CON

teste_ano=RSano("ST_Ano_Letivo")

'IF ano_letivo>=2017 THEN
'qtd_periodos = qtd_periodos-1
'periodo_m2=periodo_m2-1
'periodo_m3=periodo_m3-1
'END IF



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
					
	medias_necessarias=qtd_periodos-retira_periodo_m2-retira_periodo_m3

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
	
	reclassifica = "N"	
	if tipo_calculo="anual_ftf" then
		tipo_calculo = "anual"
		reclassifica = "S"
	elseif tipo_calculo="recuperacao_ftf" then
		tipo_calculo = "recuperacao"	
		reclassifica = "S"				
	end if 			
	
	For a=0 to ubound(alunos)
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
			if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) OR ((mae=FALSE and fil=TRUE and in_co=FALSE) AND reclassifica = "S"	) then
				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL2 = "SELECT * FROM TB_Materia where CO_Materia='"& co_materia(c) &"' order by NU_Ordem_Boletim"
				RS2.Open SQL2, CON0
					

				if not RS2.EOF then
					co_mat_princ= RS2("CO_Materia_Principal")
				end if
				
				if isnull(co_mat_princ)	then
					co_mat_princ = co_materia(c)					
				end if				


				for periodo=1 to qtd_periodos
					Set RSn = Server.CreateObject("ADODB.Recordset")
					SQLn = "SELECT * FROM "&tb_nota&" where CO_Matricula ="& dados_aluno(0) &" AND CO_Materia ='"& co_materia(c) &"' AND CO_Materia_Principal ='"& co_mat_princ &"' AND NU_Periodo="&periodo				
					RSn.Open SQLn, CONn
	
						qtd_periodos=qtd_periodos*1
						periodo=periodo*1
						periodo_m2=periodo_m2*1				
						periodo_m3=periodo_m3*1		
					if RSn.EOF then
						media_acumulada=media_acumulada				
						peso_periodo_acumulado=peso_periodo_acumulado

						if periodo=periodo_m2 then
							rec_lancado="nao"
						end if
						if periodo=periodo_m3 then
							media_final=md
							final_lancado="nao"							
						end if							
					else
						Set RSPESO = Server.CreateObject("ADODB.Recordset")
						SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&periodo
						RSPESO.Open SQLPESO, CON0
						
						md=RSn("VA_Media3")
						peso_periodo=RSPESO("NU_Peso")

						if periodo=periodo_m2 then
							media_rec=md
							rec_lancado="sim"
						elseif periodo=periodo_m3 then
							media_final=md
							final_lancado="sim"							
						else		
							if md="" or isnull(md) then
								media_acumulada=media_acumulada				
								peso_periodo_acumulado=peso_periodo_acumulado						
							else
								media_acumulada=media_acumulada+(md*peso_periodo)
								peso_periodo_acumulado=peso_periodo_acumulado+peso_periodo
								qtd_medias=qtd_medias+1						
							end if
						end if						
					end if
'					response.write(periodo&" p "&media_acumulada&"-"&qtd_medias&"<BR>")
				Next

				if peso_periodo_acumulado=0 then
					peso_periodo_acumulado=1
				end if	
'response.Write(	qtd_medias&">="&medias_necessarias&"<BR>")			
				if qtd_medias>=medias_necessarias then
					media_anual=media_acumulada/peso_periodo_acumulado					
					decimo = media_anual - Int(media_anual)
	
					If decimo >= 0.5 Then
						nota_arredondada = Int(media_anual) + 1
						media_anual=nota_arredondada
					else
						nota_arredondada = Int(media_anual)
						media_anual=nota_arredondada											
					End If		



					media_anual = formatNumber(media_anual,0)			
					media_anual=media_anual*1		
					
					media_anual = AcrescentaBonusMediaAnual(dados_aluno(0), co_materia(c), media_anual)

'if dados_aluno(0) = 31931 then									
'response.Write("<BR>"&dados_aluno(1)&" CO_Matricula ="& dados_aluno(0)&"<BR>")
'end if
					if tipo_calculo="anual" then
						resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c),curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","anual")
						media_anual = formatNumber(media_anual,0)
						resultado_materia=resultado					
					elseif tipo_calculo="recuperacao" then
'if 	dados_aluno(0)=	31931 and reclassifica = "S"then	
'response.Write(rec_lancado&"-------------------------------------------------"&media_anual&" "&media_rec&"<BR>")				
'end if	
calculo = "recuperacao"
if 	reclassifica = "S"then
calculo = "recuperacao_ftf"
end if				
						if rec_lancado="nao" or media_rec="" or isnull(media_rec) then
						verifica=1
							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c),curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;",calculo)
							resultado_materia=resultado
						else
						verifica=2
							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c),curso,co_etapa,media_anual,media_rec,"&nbsp;","&nbsp;","&nbsp;",calculo)					
							resultado_materia=resultado
						end if	
'if 	dados_aluno(0)=	31931 then	
'	response.Write(verifica&"====================================================="&media_rec&"-"&resultado_materia&"<BR>")	
'	if reclassifica = "S" then
'	response.end()
'	end if				
'end if

					elseif tipo_calculo="final" then
						if rec_lancado="nao" or media_rec="" or isnull(media_rec) then
							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c),curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","final")
							resultado_recuperacao= split(resultado,"#!#")
							if resultado_recuperacao(1)="APR" or resultado_recuperacao(1)="REP"or resultado_recuperacao(1)="APC" then
								resultado_materia=resultado
							else
								if teste_ano="B" then
									m2_aluno=m1_aluno
									resultado_materia=resultado_recuperacao(0)&"#!#REP"						
								else
									resultado_materia="&nbsp;#!#&nbsp;"	
								end if							
							end if
						elseif final_lancado="nao" or media_final="" or isnull(media_final) then
							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c),curso,co_etapa,media_anual,media_rec,"&nbsp;","&nbsp;","&nbsp;","final")
							resultado_recuperacao= split(resultado,"#!#")
							if resultado_recuperacao(1)="APR" or resultado_recuperacao(1)="REP" then
								resultado_materia=resultado
							else	
								if teste_ano="B" then
									m2_aluno=m1_aluno
									resultado_materia=resultado_recuperacao(0)&"#!#REP"						
								else
									resultado_materia="&nbsp;#!#&nbsp;"	
								end if								
							end if							
						else
						'verifica=4
							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c),curso,co_etapa,media_anual,media_rec,"&nbsp;",media_final,"&nbsp;","final")					
							resultado_materia=resultado
						end if						
					end if	
				else
						resultado_materia="&nbsp;#!#&nbsp;"
				end if	
										
			elseif (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) then
			
				vetor_mae_filhas=""

				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL2 = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_materia(c) &"' order by NU_Ordem_Boletim"
				RS2.Open SQL2, CON0
					
				co_materia_fil_conta=0 
				while not RS2.EOF
					co_mat_fil= RS2("CO_Materia")				
					if co_materia_fil_conta=0 then
						vetor_mae_filhas=co_mat_fil
					else
						vetor_mae_filhas=vetor_mae_filhas&"#!#"&co_mat_fil			
					end if
					co_materia_fil_conta=co_materia_fil_conta+1 									
				RS2.MOVENEXT
				wend				
	
				co_materia_mae_fil= split(vetor_mae_filhas,"#!#")
					conta_media=0
					media_rec_acumula=0
					media_final_acumula=0	
		
										
				for periodo=1 to qtd_periodos					
					co_materia_fil_check=1		
					media_mae_acumula=0																				
					peso_fil_acumula = 0
										
					Set RSPESO = Server.CreateObject("ADODB.Recordset")
					SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&periodo
					RSPESO.Open SQLPESO, CON0
					
					peso_periodo=RSPESO("NU_Peso")
					co_materia_fil_check=0
					
					for j=0 to ubound(co_materia_mae_fil)			
					co_materia_fil_check=co_materia_fil_check+1

						Set RSpa = Server.CreateObject("ADODB.Recordset")
						SQLpa= "SELECT * FROM TB_Programa_Aula where CO_Materia='"&co_materia_mae_fil(j)&"' AND CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"'"
						RSpa.Open SQLpa, CON0
													
						nu_peso_fil=RSpa("NU_Peso")						
					
						if isnull(nu_peso_fil) or nu_peso_fil="" then
							nu_peso_fil=1
						end if		
						
						peso_fil_acumula = peso_fil_acumula+nu_peso_fil			
						
						Set RS3 = Server.CreateObject("ADODB.Recordset")
						SQL3 = "SELECT * FROM "&tb_nota&" where CO_Matricula ="& dados_aluno(0) &" AND CO_Materia_Principal ='"& co_materia(c)   &"' AND CO_Materia ='"& co_materia_mae_fil(j)&"' And NU_Periodo="&periodo
						RS3.Open SQL3, CONn

						if RS3.EOF then
							media_mae_acumula=media_mae_acumula	
							media_rec_acumula=media_rec_acumula
							media_final_acumula=media_final_acumula											
						else
							media_aluno=RS3("VA_Media3")
							periodo=periodo*1
							periodo_m2=periodo_m2*1					
							periodo_m3=periodo_m3*1																				

						
							if media_aluno="" or isnull(media_aluno) then
							else
								'media_aluno = formatNumber(media_aluno/10,1)							
								media_aluno = media_aluno*nu_peso_fil
								'media_aluno = media_aluno*10								
							end if
					
							if periodo=periodo_m2 then
								if media_aluno="" or isnull(media_aluno) then
									media_rec_acumula=media_rec_acumula
								else
									media_rec_acumula=media_rec_acumula*1
									media_aluno=media_aluno*1
									media_rec_acumula=media_rec_acumula+media_aluno
								end if 														
							elseif periodo=periodo_m3 then
								if media_aluno="" or isnull(media_aluno) then
									media_final_acumula=media_final_acumula
								else
									media_final_acumula=media_final_acumula*1
									media_aluno=media_aluno*1								
									media_final_acumula=media_final_acumula+media_aluno
								end if 							
							else						
								if media_aluno="" or isnull(media_aluno) then
									media_mae_acumula=media_mae_acumula
								else
									media_mae_acumula=media_mae_acumula*1
									media_aluno=media_aluno*1																
									media_mae_acumula=media_mae_acumula+media_aluno
								end if 
							end if
						end if				
						
						co_materia_fil_check=co_materia_fil_check*1
						co_materia_fil_conta=co_materia_fil_conta*1
						
						if co_materia_fil_check=co_materia_fil_conta then
							if media_mae_acumula=0 then
								md_mae_periodo=""
								media_acumulada=media_acumulada				
								peso_periodo_acumulado=peso_periodo_acumulado						
							else
								md_mae_periodo=media_mae_acumula/peso_fil_acumula
								md_mae_periodo=arredonda(md_mae_periodo,"mat",0,0)	
								media_acumulada=media_acumulada*1															
								media_acumulada=media_acumulada+(md_mae_periodo*peso_periodo)
								peso_periodo_acumulado=peso_periodo_acumulado*1
								peso_periodo=peso_periodo*1
								peso_periodo_acumulado=peso_periodo_acumulado+peso_periodo
								qtd_medias=qtd_medias+1	
							end if							
						end if


					NEXT					
'if dados_aluno(0)=31274 then
'	response.Write(" MED_MAE_PRD ='"& media_mae_acumula&"/"&peso_fil_acumula&"<BR>")
'end if	

'						response.write("MA "&media_mae_acumula&"<BR>")
	

					if media_rec_acumula=0 then
						media_rec=""
						rec_lancado="nao"						
					else
						media_rec=(media_rec_acumula)/(peso_fil_acumula)
						media_rec=formatnumber(media_rec,0)	
						rec_lancado="sim"								
					end if	

					if media_final_acumula=0 then
						media_final=""
						final_lancado="nao"						
					else
						media_final=(media_final_acumula)/peso_fil_acumula			
						media_final=formatnumber(media_final,0)	
						final_lancado="sim"								
					end if	
				next

				if peso_periodo_acumulado=0 then
					peso_periodo_acumulado=1
				end if	
'response.Write("--------------"&media_acumulada&"-"&peso_periodo_acumulado&"--------------")					
				if qtd_medias>=medias_necessarias then
					media_anual=media_acumulada/peso_periodo_acumulado					
					decimo = media_anual - Int(media_anual)
'response.Write("--------------"&media_anual&"--------------")
					media_anual=arredonda(media_anual,"mat_dez",0,0)
						
					media_anual = formatNumber(media_anual,0)			
					media_anual=media_anual*1		
					
					media_anual = AcrescentaBonusMediaAnual(dados_aluno(0), co_materia(c), media_anual)									
'					if media_anual>67 and media_anual<70 then
'						media_anual=70
'					end if		
'response.Write(dados_aluno(0)&"------------>-"&media_anual&"------------+-"&co_materia(c)) 					
'response.End()
					if tipo_calculo="anual" then
						resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c), curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","anual")
						media_anual = formatNumber(media_anual,0)
						resultado_materia=resultado					
					elseif tipo_calculo="recuperacao" then
						if rec_lancado="nao" or media_rec="" or isnull(media_rec) then

							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c), curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","recuperacao")
							resultado_materia=resultado
						else

							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c), curso,co_etapa,media_anual,media_rec,"&nbsp;","&nbsp;","&nbsp;","recuperacao")					
							resultado_materia=resultado
						end if	
	
					elseif tipo_calculo="final" then						
'response.Write("if "& rec_lancado&"=rec_lancado or media_rec="& media_rec&" or isnull(media_rec) or final_lancado="&final_lancado&" or media_final="&media_final&" or<BR>")		
						if rec_lancado="nao" or media_rec="" or isnull(media_rec) then
							'verifica=3
							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c), curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","final")
							resultado_recuperacao= split(resultado,"#!#")
							if resultado_recuperacao(1)="APR" or resultado_recuperacao(1)="REP" then
								resultado_materia=resultado
							else
								if teste_ano="B" then
									m2_aluno=m1_aluno
									resultado_materia=resultado_recuperacao(0)&"#!#REP"						
								else
									resultado_materia="&nbsp;#!#&nbsp;"	
								end if							
							end if
						elseif final_lancado="nao" or media_final="" or isnull(media_final) then
							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c), curso,co_etapa,media_anual,media_rec,"&nbsp;","&nbsp;","&nbsp;","final")
							resultado_recuperacao= split(resultado,"#!#")
							if resultado_recuperacao(1)="APR" or resultado_recuperacao(1)="REP" then
								resultado_materia=resultado
							else
								if teste_ano="B" then
									m2_aluno=m1_aluno
									resultado_materia=resultado_recuperacao(0)&"#!#REP"						
								else
									resultado_materia="&nbsp;#!#&nbsp;"	
								end if											
							end if							
						else
						'verifica=4
							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c), curso,co_etapa,media_anual,media_rec,"&nbsp;",media_final,"&nbsp;","final")					
							resultado_materia=resultado
						end if	
					end if						
'				response.Write(verifica&"-"&resultado_materia&"<BR>")	
				else
						resultado_materia="&nbsp;#!#&nbsp;"
				end if				

			elseif (mae=TRUE and fil=TRUE and in_co=FALSE) then			
			
			elseif (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then	
				vetor_mae_filhas=""

				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL2 = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_materia(c) &"' order by NU_Ordem_Boletim"
				RS2.Open SQL2, CON0
					
				co_materia_fil_conta=0 
				while not RS2.EOF
					co_mat_fil= RS2("CO_Materia")				
					if co_materia_fil_conta=0 then
						vetor_mae_filhas=co_mat_fil
					else
						vetor_mae_filhas=vetor_mae_filhas&"#!#"&co_mat_fil			
					end if
					co_materia_fil_conta=co_materia_fil_conta+1 									
				RS2.MOVENEXT
				wend				
	
				co_materia_mae_fil= split(vetor_mae_filhas,"#!#")
					conta_media=0
					media_rec_acumula=0
					media_final_acumula=0	
		
										
				for periodo=1 to qtd_periodos					
					co_materia_fil_check=1		
					media_mae_acumula=0																				
					peso_fil_acumula = 0
										
					Set RSPESO = Server.CreateObject("ADODB.Recordset")
					SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&periodo
					RSPESO.Open SQLPESO, CON0
					
					peso_periodo=RSPESO("NU_Peso")
					co_materia_fil_check=0
					
'						nome_nota = var_bd_periodo(tp_modelo,tp_freq, tb_nota,periodo,"BDM")	
'						media_aluno=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, dados_aluno(0), co_materia(c), vetor_mae_filhas, caminho_nota, tb_nota, periodo, nome_nota, outro)
'						
'							media_aluno=RS3("VA_Media3")
'							periodo=periodo*1
'							periodo_m2=periodo_m2*1					
'							periodo_m3=periodo_m3*1																				
'
''						
'							if media_aluno="" or isnull(media_aluno) then
'							else
'								'media_aluno = formatNumber(media_aluno/10,1)							
'								media_aluno = media_aluno*nu_peso_fil
'								'media_aluno = media_aluno*10								
'							end if
'					
'							if periodo=periodo_m2 then
'								if media_aluno="" or isnull(media_aluno) then
'									media_rec_acumula=media_rec_acumula
'								else
'									media_rec_acumula=media_rec_acumula*1
'									media_aluno=media_aluno*1
'									media_rec_acumula=media_rec_acumula+media_aluno
'								end if 														
'							elseif periodo=periodo_m3 then
'								if media_aluno="" or isnull(media_aluno) then
'									media_final_acumula=media_final_acumula
'								else
'									media_final_acumula=media_final_acumula*1
'									media_aluno=media_aluno*1								
'									media_final_acumula=media_final_acumula+media_aluno
'								end if 							
'							else						
'								if media_aluno="" or isnull(media_aluno) then
'									media_mae_acumula=media_mae_acumula
'								else
'									media_mae_acumula=media_mae_acumula*1
'									media_aluno=media_aluno*1																
'									media_mae_acumula=media_mae_acumula+media_aluno
'								end if 
'							end if
'						end if										
'					
					
					
					for j=0 to ubound(co_materia_mae_fil)			
					co_materia_fil_check=co_materia_fil_check+1

						Set RSpa = Server.CreateObject("ADODB.Recordset")
						SQLpa= "SELECT * FROM TB_Programa_Aula where CO_Materia='"&co_materia_mae_fil(j)&"' AND CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"'"
						RSpa.Open SQLpa, CON0
													
						nu_peso_fil=RSpa("NU_Peso")						
					
						if isnull(nu_peso_fil) or nu_peso_fil="" then
							nu_peso_fil=1
						end if		
						
						peso_fil_acumula = peso_fil_acumula+nu_peso_fil			
						
						Set RS3 = Server.CreateObject("ADODB.Recordset")
						SQL3 = "SELECT * FROM "&tb_nota&" where CO_Matricula ="& dados_aluno(0) &" AND CO_Materia_Principal ='"& co_materia(c)   &"' AND CO_Materia ='"& co_materia_mae_fil(j)&"' And NU_Periodo="&periodo
						'response.Write("CO_Materia_Principal ='"& co_materia(c)   &"' AND CO_Materia ='"& co_materia_mae_fil(j)&"<BR>")
						RS3.Open SQL3, CONn

						if RS3.EOF then
							media_mae_acumula=media_mae_acumula	
							media_rec_acumula=media_rec_acumula
							media_final_acumula=media_final_acumula											
						else
							media_aluno=RS3("VA_Media3")
							periodo=periodo*1
							periodo_m2=periodo_m2*1					
							periodo_m3=periodo_m3*1																				

						
							if media_aluno="" or isnull(media_aluno) then
							else
								'media_aluno = formatNumber(media_aluno/10,1)							
								media_aluno = media_aluno*nu_peso_fil
								'media_aluno = media_aluno*10								
							end if
					
							if periodo=periodo_m2 then
								if media_aluno="" or isnull(media_aluno) then
									media_rec_acumula=media_rec_acumula
								else
									media_rec_acumula=media_rec_acumula*1
									media_aluno=media_aluno*1
									media_rec_acumula=media_rec_acumula+media_aluno
								end if 														
							elseif periodo=periodo_m3 then
								if media_aluno="" or isnull(media_aluno) then
									media_final_acumula=media_final_acumula
								else
									media_final_acumula=media_final_acumula*1
									media_aluno=media_aluno*1								
									media_final_acumula=media_final_acumula+media_aluno
								end if 							
							else						
								if media_aluno="" or isnull(media_aluno) then
									media_mae_acumula=media_mae_acumula
								else
									media_mae_acumula=media_mae_acumula*1
									media_aluno=media_aluno*1																
									media_mae_acumula=media_mae_acumula+media_aluno
								end if 
							end if
						end if	
						
'	response.Write(periodo&" "& media_aluno&" "&media_mae_acumula&"<BR>")									
						
						co_materia_fil_check=co_materia_fil_check*1
						co_materia_fil_conta=co_materia_fil_conta*1
						
						if co_materia_fil_check=co_materia_fil_conta then
							if media_mae_acumula=0 then
								md_mae_periodo=""
								media_acumulada=media_acumulada				
								peso_periodo_acumulado=peso_periodo_acumulado						
							else
								if media_mae_acumula>100 then
									media_mae_acumula=100
								end if								
								md_mae_periodo=media_mae_acumula
								md_mae_periodo=arredonda(md_mae_periodo,"mat",0,0)	
								media_acumulada=media_acumulada*1															
								media_acumulada=media_acumulada+(md_mae_periodo*peso_periodo)
								peso_periodo_acumulado=peso_periodo_acumulado*1
								peso_periodo=peso_periodo*1
								peso_periodo_acumulado=peso_periodo_acumulado+peso_periodo
								qtd_medias=qtd_medias+1	
							end if							
						end if


					NEXT					
'if dados_aluno(0)=70076 then '31274
'	response.Write(" MED_MAE_PRD ='"& media_mae_acumula&" "&peso_periodo_acumulado&"<BR>")
'end if	

'						response.write("MA "&media_mae_acumula&"<BR>")
	

					if media_rec_acumula=0 then
						media_rec=""
						rec_lancado="nao"						
					else
						if media_rec_acumula>100 then
							media_rec_acumula=100
						end if													
						media_rec=media_rec_acumula
						media_rec=formatnumber(media_rec,0)	
						rec_lancado="sim"								
					end if	

					if media_final_acumula=0 then
						media_final=""
						final_lancado="nao"						
					else
						if media_final_acumula>100 then
							media_final_acumula=100
						end if						
						media_final=media_final_acumula	
						media_final=formatnumber(media_final,0)	
						final_lancado="sim"								
					end if	
				next
				
'if dados_aluno(0)=70076 then '31274
'	response.Write(media_rec&" media_rec  "& media_final&"<BR>")
'
'end if					

				if peso_periodo_acumulado=0 then
					peso_periodo_acumulado=1
				end if	
'response.Write("--------------"&media_acumulada&"-"&peso_periodo_acumulado&"--------------")					
				if qtd_medias>=medias_necessarias then
					media_anual=media_acumulada/peso_periodo_acumulado			
					decimo = media_anual - Int(media_anual)
'response.Write("--------------"&media_anual&"--------------")
					media_anual=arredonda(media_anual,"mat_dez",0,0)
						
					media_anual = formatNumber(media_anual,0)			
					media_anual=media_anual*1		
					
					media_anual = AcrescentaBonusMediaAnual(dados_aluno(0), co_materia(c), media_anual)									
'					if media_anual>67 and media_anual<70 then
'						media_anual=70
'					end if		
'response.Write("------------>-"&media_anual&"------------+-")	co_materia(c) 					
					if tipo_calculo="anual" then
						resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c), curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","anual")
						media_anual = formatNumber(media_anual,0)
						resultado_materia=resultado					
					elseif tipo_calculo="recuperacao" then
						if rec_lancado="nao" or media_rec="" or isnull(media_rec) then

							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c), curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","recuperacao")
							resultado_materia=resultado
						else

							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c), curso,co_etapa,media_anual,media_rec,"&nbsp;","&nbsp;","&nbsp;","recuperacao")					
							resultado_materia=resultado
						end if	
	
					elseif tipo_calculo="final" then						
'response.Write("if "& rec_lancado&"=rec_lancado or media_rec="& media_rec&" or isnull(media_rec) or final_lancado="&final_lancado&" or media_final="&media_final&" or<BR>")		
						if rec_lancado="nao" or media_rec="" or isnull(media_rec) then
							'verifica=3
							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c), curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","final")
							resultado_recuperacao= split(resultado,"#!#")
							if resultado_recuperacao(1)="APR" or resultado_recuperacao(1)="REP" then
								resultado_materia=resultado
							else
								if teste_ano="B" then
									m2_aluno=m1_aluno
									resultado_materia=resultado_recuperacao(0)&"#!#REP"						
								else
									resultado_materia="&nbsp;#!#&nbsp;"	
								end if							
							end if
						elseif final_lancado="nao" or media_final="" or isnull(media_final) then
							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c), curso,co_etapa,media_anual,media_rec,"&nbsp;","&nbsp;","&nbsp;","final")
							resultado_recuperacao= split(resultado,"#!#")
							if resultado_recuperacao(1)="APR" or resultado_recuperacao(1)="REP" then
								resultado_materia=resultado
							else
								if teste_ano="B" then
									m2_aluno=m1_aluno
									resultado_materia=resultado_recuperacao(0)&"#!#REP"						
								else
									resultado_materia="&nbsp;#!#&nbsp;"	
								end if											
							end if							
						else
						'verifica=4
							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c), curso,co_etapa,media_anual,media_rec,"&nbsp;",media_final,"&nbsp;","final")					
							resultado_materia=resultado
						end if	
					end if						
'				response.Write(verifica&"-"&resultado_materia&"<BR>")	
				else
						resultado_materia="&nbsp;#!#&nbsp;"
				end if								
				'vetor_mae_filhas=""
'		
'				Set RS2 = Server.CreateObject("ADODB.Recordset")
'				SQL2 = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_materia(c) &"' order by NU_Ordem_Boletim"
'				RS2.Open SQL2, CON0
'					
'				co_materia_fil_conta=0 
'				while not RS2.EOF
'					co_mat_fil= RS2("CO_Materia")				
'					if co_materia_fil_conta=0 then
'						vetor_mae_filhas=co_mat_fil
'					else
'						vetor_mae_filhas=vetor_mae_filhas&"#!#"&co_mat_fil			
'					end if
'					co_materia_fil_conta=co_materia_fil_conta+1 									
'				RS2.MOVENEXT
'				wend				
'	
'				co_materia_mae_fil= split(vetor_mae_filhas,"#!#")
'					conta_media=0
'					media_rec_acumula=0
'					media_final_acumula=0								
'				for periodo=1 to qtd_periodos					
'					co_materia_fil_check=1		
'					media_mae_acumula=0															
'					
'					Set RSPESO = Server.CreateObject("ADODB.Recordset")
'					SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&periodo
'					RSPESO.Open SQLPESO, CON0
'					
'					peso_periodo=RSPESO("NU_Peso")
'
'					co_materia_fil_check=0
'					for j=0 to ubound(co_materia_mae_fil)			
'					co_materia_fil_check=co_materia_fil_check+1
'						Set RS3 = Server.CreateObject("ADODB.Recordset")
'						SQL3 = "SELECT * FROM "&tb_nota&" where CO_Matricula ="& dados_aluno(0) &" AND CO_Materia_Principal ='"& co_materia(c)  &"' AND CO_Materia ='"& co_materia_mae_fil(j)&"' And NU_Periodo="&periodo
'
'						RS3.Open SQL3, CONn
'						if RS3.EOF then
'							media_mae_acumula=media_mae_acumula	
'							media_rec_acumula=media_rec_acumula
'							media_final_acumula=media_final_acumula											
'						else
'							media_aluno=RS3("VA_Media3")
'							periodo=periodo*1
'							periodo_m2=periodo_m2*1					
'							periodo_m3=periodo_m3*1																				
'
'							if periodo=periodo_m2 then
'								if media_aluno="" or isnull(media_aluno) then
'									media_rec_acumula=media_rec_acumula
'								else
'									media_rec_acumula=media_rec_acumula+media_aluno
'								end if 														
'							elseif periodo=periodo_m3 then
'								if media_aluno="" or isnull(media_aluno) then
'									media_final_acumula=media_final_acumula
'								else
'									media_final_acumula=media_final_acumula+media_aluno
'								end if 							
'							else						
'								if media_aluno="" or isnull(media_aluno) then
'									media_mae_acumula=media_mae_acumula
'								else
'									media_mae_acumula=media_mae_acumula+media_aluno
'								end if 
'							end if
'						end if	
'							co_materia_fil_check=co_materia_fil_check*1
'							co_materia_fil_conta=co_materia_fil_conta*1
'							if co_materia_fil_check=co_materia_fil_conta then
'								if media_mae_acumula=0 then
'									md=""
'									media_acumulada=media_acumulada				
'									peso_periodo_acumulado=peso_periodo_acumulado						
'								else
'									md=media_mae_acumula/co_materia_fil_conta				
'									media_acumulada=media_acumulada+(md*peso_periodo)
'									peso_periodo_acumulado=peso_periodo_acumulado+peso_periodo
'									qtd_medias=qtd_medias+1	
'								end if							
'						end if
'
'					NEXT					
'
'					if media_rec_acumula=0 then
'						media_rec=""
'						rec_lancado="nao"						
'					else
'						media_rec=(media_rec_acumula*peso_periodo)/co_materia_fil_conta			
'						media_rec=formatnumber(media_rec,0)	
'						rec_lancado="sim"								
'					end if	
'
'					if media_final_acumula=0 then
'						media_final=""
'						final_lancado="nao"						
'					else
'						media_final=(media_final_acumula*peso_periodo)/co_materia_fil_conta			
'						media_final=formatnumber(media_final,0)	
'						final_lancado="sim"								
'					end if	
'				next
'
'				if peso_periodo_acumulado=0 then
'					peso_periodo_acumulado=1
'				end if	
'							
'				if qtd_medias>=medias_necessarias then
'					media_anual=media_acumulada/peso_periodo_acumulado					
'					decimo = media_anual - Int(media_anual)
''					If decimo >= 0.75 Then
''						nota_arredondada = Int(media_anual) + 1
''						media_anual=nota_arredondada
''					elseIf decimo >= 0.25 Then
''						nota_arredondada = Int(media_anual) + 0.5
''						media_anual=nota_arredondada
''					else
''						nota_arredondada = Int(media_anual)
''						media_anual=nota_arredondada											
''					End If		
'					If decimo >= 0.5 Then
'						nota_arredondada = Int(media_anual) + 1
'						media_anual=nota_arredondada
'					else
'						nota_arredondada = Int(media_anual)
'						media_anual=nota_arredondada											
'					End If	
'					media_anual = formatNumber(media_anual,0)			
'					media_anual=media_anual*1	
'					
'					media_anual = AcrescentaBonusMediaAnual(dados_aluno(0), co_materia(c), media_anual)											
''					if media_anual>67 and media_anual<70 then
''						media_anual=70
''					end if						
'					if tipo_calculo="anual" then
'						resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c), curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","anual")
'						media_anual = formatNumber(media_anual,0)
'						resultado_materia=resultado					
'					elseif tipo_calculo="recuperacao" then
'						if rec_lancado="nao" or media_rec="" or isnull(media_rec) then
'
'							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c), curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","recuperacao")
'							resultado_materia=resultado
'						else
'
'							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c), curso,co_etapa,media_anual,media_rec,"&nbsp;","&nbsp;","&nbsp;","recuperacao")					
'							resultado_materia=resultado
'						end if	
'	
'					elseif tipo_calculo="final" then						
''response.Write("if "& rec_lancado&"=rec_lancado or media_rec="& media_rec&" or isnull(media_rec) or final_lancado="&final_lancado&" or media_final="&media_final&" or<BR>")		
'						if rec_lancado="nao" or media_rec="" or isnull(media_rec) then
'							'verifica=3
'							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c), curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","final")
'							resultado_recuperacao= split(resultado,"#!#")
'							if resultado_recuperacao(1)="APR" or resultado_recuperacao(1)="REP" then
'								resultado_materia=resultado
'							else
'								if teste_ano="B" then
'									m2_aluno=m1_aluno
'									resultado_materia=resultado_recuperacao(0)&"#!#REP"						
'								else
'									if teste_ano="B" then
'										m2_aluno=m1_aluno
'										resultado_materia=resultado_recuperacao(0)&"#!#REP"						
'									else
'										resultado_materia="&nbsp;#!#&nbsp;"	
'									end if		
'								end if							
'							end if
'						elseif final_lancado="nao" or media_final="" or isnull(media_final) then
'							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c), curso,co_etapa,media_anual,media_rec,"&nbsp;","&nbsp;","&nbsp;","final")
'							resultado_recuperacao= split(resultado,"#!#")
'							if resultado_recuperacao(1)="APR" or resultado_recuperacao(1)="REP" then
'								resultado_materia=resultado
'							else
'								resultado_materia="&nbsp;#!#&nbsp;"										
'							end if							
'						else
'						'verifica=4
'							resultado=novo_regra_aprovacao (dados_aluno(0), co_materia(c), curso,co_etapa,media_anual,media_rec,"&nbsp;",media_final,"&nbsp;","final")					
'							resultado_materia=resultado
'						end if	
'					end if						
''				response.Write(verifica&"-"&resultado_materia&"<BR>")	
'				else
'						resultado_materia="&nbsp;#!#&nbsp;"
'				end if						

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
'response.Write(m1_aluno&"_"&nota_aux_m2_1&"_"&nota_aux_m2_2&"_"&nota_aux_m3_1&"_"&nota_aux_m3_2)
	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR

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

	if m1_aluno >= m1_maior_igual then
		resultado=res1_3
		resultado1="apr"
	elseif m1_aluno >= m1_menor then
		resultado=res1_2
	else
		resultado=res1_1
		resultado1="rep"			
	end if
	
	
		Set RSano = Server.CreateObject("ADODB.Recordset")
	SQLano = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&ano_letivo&"'"
	RSano.Open SQLano, CON

teste_ano=RSano("ST_Ano_Letivo")


	
	
	if tipo_calculo="waboletim" then
		m1_waboletim=m1_aluno
		resultado1_waboletim=resultado
	end if	
'response.Write("if "&m1_aluno &">="& m1_maior_igual &"then<BR>")
'response.Write("elseif "&m1_aluno &">"& m1_menor &"then<BR>")	
'response.Write(resultado&"<BR>")	


	if (resultado1="apr" or resultado1="rep") then
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
'				m2_aluno="&nbsp;"
'				resultado="&nbsp;"	
				
				if teste_ano="B" then
					m2_aluno=m1_aluno
					resultado="REP"					
				else
					m2_aluno="&nbsp;"
					resultado="&nbsp;"	
				end if	
				
				if tipo_calculo="waboletim" then
					m2_waboletim=m2_aluno	
					resultado2_waboletim=resultado	
				end if	
			else								
				m1_aluno_peso=m1_aluno*peso_m2_m1
				nota_aux_m2_1_peso=nota_aux_m2_1*peso_m2_m2
'				response.Write(m1_aluno_peso&"="&m1_aluno&"*"&peso_m2_m1&"----")
'				response.Write(nota_aux_m2_1_peso&"="&nota_aux_m2_1&"*"&peso_m2_m2)
				m2_aluno=(m1_aluno_peso+nota_aux_m2_1_peso)/(peso_m2_m1+peso_m2_m2)
				decimo = m2_aluno - Int(m2_aluno)
'				If decimo >= 0.75 Then
'					nota_arredondada = Int(m2_aluno) + 1
'					m2_aluno=nota_arredondada
'				elseIf decimo >= 0.25 Then
'					nota_arredondada = Int(m2_aluno) + 0.5
'					m2_aluno=nota_arredondada
'				else
'					nota_arredondada = Int(m2_aluno)
'					m2_aluno=nota_arredondada											
'				End If	
				If decimo >= 0.5 Then
					nota_arredondada = Int(m2_aluno) + 1
					m2_aluno=nota_arredondada
				else
					nota_arredondada = Int(m2_aluno)
					m2_aluno=nota_arredondada											
				End If	
				m2_aluno = formatNumber(m2_aluno,0)
				m2_aluno=m2_aluno*1
				m2_maior_igual=m2_maior_igual*1	
				m2_menor=m2_menor*1		
				if m2_aluno >= m2_maior_igual then
					resultado=res2_3
					resultado2="apr"
				elseif m2_aluno >= m2_menor then
					resultado=res2_2
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
 			if	tipo_calculo="final" or tipo_calculo="waboletim" then
				if (resultado2="apr" or resultado2="rep") then
					m3_aluno=m2_aluno				
				else
					if m2_aluno="&nbsp;" or nota_aux_m2_1="&nbsp;" or nota_aux_m3_1="&nbsp;" then		
'						m3_aluno="&nbsp;"
'						resultado="&nbsp;"	
						if teste_ano="B" then
							m3_aluno=m2_aluno
							resultado="REP"					
						else
							m3_aluno="&nbsp;"
							resultado="&nbsp;"	
						end if									
					else								
						m1_aluno_peso=m1_aluno*peso_m3_m1
						m2_aluno_peso=m2_aluno*peso_m3_m2
						nota_aux_m3_1_peso=nota_aux_m3_1*peso_m3_m3
'response.Write(m3_aluno&"=("&m1_aluno_peso&"+"&m2_aluno_peso&"+"&nota_aux_m3_1_peso&")/("&peso_m3_m1&"+"&peso_m3_m2&"+"&peso_m3_m3)

						m3_aluno=(m1_aluno_peso+m2_aluno_peso+nota_aux_m3_1_peso)/(peso_m3_m1+peso_m3_m2+peso_m3_m3)
						decimo = m3_aluno - Int(m3_aluno)
'						If decimo >= 0.75 Then
'							nota_arredondada = Int(m3_aluno) + 1
'							m3_aluno=nota_arredondada
'						elseIf decimo >= 0.25 Then
'							nota_arredondada = Int(m3_aluno) + 0.5
'							m3_aluno=nota_arredondada
'						else
'							nota_arredondada = Int(m3_aluno)
'							m3_aluno=nota_arredondada											
'						End If	
						If decimo >= 0.5 Then
							nota_arredondada = Int(m3_aluno) + 1
							m3_aluno=nota_arredondada
						else
							nota_arredondada = Int(m3_aluno)
							m3_aluno=nota_arredondada											
						End If	
						m3_aluno = formatNumber(m3_aluno,0)
						m3_aluno=m3_aluno*1
						valor_m3=valor_m3*1		
						m3_maior_igual=m3_maior_igual*1		
						if m3_aluno >= m3_maior_igual then
							resultado=res3_2
						else
							resultado=res3_1	
						end if
						if tipo_calculo="waboletim" then
							m3_waboletim=m3_aluno		
							resultado3_waboletim=resultado	
						end if							
					end if		
				end if
			end if	
		end if	
	end if

	if tipo_calculo="anual" then
		m1_aluno = formatNumber(m1_aluno,0)		
		regra_aprovacao=m1_aluno&"#!#"&resultado
	elseif tipo_calculo="recuperacao" then
		if resultado1="apr" or resultado1="rep" then
			m1_aluno = formatNumber(m1_aluno,0)	
			regra_aprovacao=m1_aluno&"#!#"&resultado		
		else
			if m2_aluno<>"&nbsp;" then
				m2_aluno = formatNumber(m2_aluno,0)
			end if
			regra_aprovacao=m2_aluno&"#!#"&resultado
		end if
	elseif tipo_calculo="waboletim" then
			if m2_aluno<>"&nbsp;" then
				m2_aluno = formatNumber(m2_aluno,0)
			end if	

			if m3_aluno<>"&nbsp;" then
				m3_aluno = formatNumber(m3_aluno,0)
			end if
		regra_aprovacao=m1_waboletim&"#!#"&resultado1_waboletim&"#!#"&m2_waboletim&"#!#"&resultado2_waboletim&"#!#"&m3_waboletim&"#!#"&resultado3_waboletim
	else
		if resultado2="apr" or resultado2="rep" then
			if m2_aluno<>"&nbsp;" then
				m2_aluno = formatNumber(m2_aluno,0)
			end if
			regra_aprovacao=m2_aluno&"#!#"&resultado		
		else
			if m3_aluno<>"&nbsp;" then
				m3_aluno = formatNumber(m3_aluno,0)
			end if
			regra_aprovacao=m3_aluno&"#!#"&resultado			
		end if
	end if
	
	'Session("M2")=m2_aluno
	'Session("M3")=m3_aluno
end function

Function novo_regra_aprovacao (cod_aluno, cod_materia, curso,etapa, m1_aluno, nota_aux_m2_1, nota_aux_m2_2, nota_aux_m3_1,nota_aux_m3_2, tipo_calculo)
'response.Write(m1_aluno&"_"&nota_aux_m2_1&"_"&nota_aux_m2_2&"_"&nota_aux_m3_1&"_"&nota_aux_m3_2)
	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR

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

	if m1_aluno >= m1_maior_igual then
		resultado=res1_3
		resultado1="apr"
	elseif m1_aluno >= m1_menor then
		resultado=res1_2
	else
		resultado=res1_1
		resultado1="rep"			
	end if
	
	tipo_media = "MA"
	modifica_result = Verifica_Conselho_Classe(cod_aluno, cod_materia, tipo_media, outro)							
	if modifica_result <> "N" then
		resultado = modifica_result
		resultado1 = resultado
	end if		
	
	
		Set RSano = Server.CreateObject("ADODB.Recordset")
	SQLano = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&ano_letivo&"'"
	RSano.Open SQLano, CON

teste_ano=RSano("ST_Ano_Letivo")

	
	
	if tipo_calculo="waboletim" then
		m1_waboletim=m1_aluno
		resultado1_waboletim=resultado
	end if	
'response.Write("if "&m1_aluno &">="& m1_maior_igual &"then<BR>")
'response.Write("elseif "&m1_aluno &">"& m1_menor &"then<BR>")	
'response.Write(resultado&"<BR>")	
	if (resultado1="apr" or resultado1="rep" or resultado1="APC") and (tipo_calculo<>"waboletim_FTF" and tipo_calculo <> "recuperacao_ftf") then
		m2_aluno=m1_aluno	
		m3_aluno=m1_aluno
		if tipo_calculo="waboletim" then
			m2_waboletim="&nbsp;"
			m3_waboletim="&nbsp;"			
			resultado2_waboletim="&nbsp;"
			resultado3_waboletim="&nbsp;"
		elseif tipo_calculo="waboletim_FTF"	then
			m2_waboletim = nota_aux_m2_1
		end if		
		
	else			
		if tipo_calculo="recuperacao" or tipo_calculo="final" or tipo_calculo="waboletim" or tipo_calculo="waboletim_FTF" or tipo_calculo = "recuperacao_ftf" then
			if nota_aux_m2_1="&nbsp;" then
'				m2_aluno="&nbsp;"
'				resultado="&nbsp;"	
				
				if teste_ano="B" then
					m2_aluno=m1_aluno
					resultado="REP"					
				else
					m2_aluno="&nbsp;"
					resultado="&nbsp;"	
				end if	
				
				if tipo_calculo="waboletim" then
					m2_waboletim=m2_aluno	
					resultado2_waboletim=resultado	
				end if	
			else								
				m1_aluno_peso=m1_aluno*peso_m2_m1
				nota_aux_m2_1_peso=nota_aux_m2_1*peso_m2_m2
'				response.Write(cod_aluno&" "&tipo_calculo&" "&m1_aluno_peso&"="&m1_aluno&"*"&peso_m2_m1&"----")
'				response.Write(nota_aux_m2_1_peso&"="&nota_aux_m2_1&"*"&peso_m2_m2&"<BR>")
				m2_aluno=(m1_aluno_peso+nota_aux_m2_1_peso)/(peso_m2_m1+peso_m2_m2)
				decimo = m2_aluno - Int(m2_aluno)
'				If decimo >= 0.75 Then
'					nota_arredondada = Int(m2_aluno) + 1
'					m2_aluno=nota_arredondada
'				elseIf decimo >= 0.25 Then
'					nota_arredondada = Int(m2_aluno) + 0.5
'					m2_aluno=nota_arredondada
'				else
'					nota_arredondada = Int(m2_aluno)
'					m2_aluno=nota_arredondada											
'				End If	
				If decimo >= 0.5 Then
					nota_arredondada = Int(m2_aluno) + 1
					m2_aluno=nota_arredondada
				else
					nota_arredondada = Int(m2_aluno)
					m2_aluno=nota_arredondada											
				End If	
				m2_aluno = formatNumber(m2_aluno,0)
				m2_aluno=m2_aluno*1
				m2_maior_igual=m2_maior_igual*1	
				m2_menor=m2_menor*1		
				if m2_aluno >= m2_maior_igual then
					resultado=res2_3
					resultado2="apr"
				elseif m2_aluno >= m2_menor then
					resultado=res2_2
				else
					resultado=res2_1	
					resultado2="rep"					
				end if
				
				tipo_media = "RF"
				modifica_result = Verifica_Conselho_Classe(cod_aluno, cod_materia, tipo_media, outro)
				if modifica_result <> "N" then
					resultado = modifica_result
					resultado2 = resultado
				end if					

				if tipo_calculo="waboletim" or tipo_calculo="waboletim_FTF" then
					m2_waboletim=m2_aluno		
					resultado2_waboletim=resultado	
				end if	

'response.Write("if "&m2_aluno &">="& m2_maior_igual &"then<BR>")
'response.Write("elseif "&m2_aluno &">"& m2_menor &"then<BR>")	
'response.Write(resultado&"<BR>")

			end if
 			if	tipo_calculo="final" or tipo_calculo="waboletim" then
				if resultado2="apr" or resultado2="rep" or resultado2="APC" then
					m3_aluno=m2_aluno				
				else
					if m2_aluno="&nbsp;" or nota_aux_m2_1="&nbsp;" or nota_aux_m3_1="&nbsp;" then		
'						m3_aluno="&nbsp;"
'						resultado="&nbsp;"	
						if teste_ano="B" then
							m3_aluno=m2_aluno
							resultado="REP"					
						else
							m3_aluno="&nbsp;"
							resultado="&nbsp;"	
						end if									
					else								
						m1_aluno_peso=m1_aluno*peso_m3_m1
						m2_aluno_peso=m2_aluno*peso_m3_m2
						nota_aux_m3_1_peso=nota_aux_m3_1*peso_m3_m3
'response.Write(m3_aluno&"=("&m1_aluno_peso&"+"&m2_aluno_peso&"+"&nota_aux_m3_1_peso&")/("&peso_m3_m1&"+"&peso_m3_m2&"+"&peso_m3_m3)

						m3_aluno=(m1_aluno_peso+m2_aluno_peso+nota_aux_m3_1_peso)/(peso_m3_m1+peso_m3_m2+peso_m3_m3)
						decimo = m3_aluno - Int(m3_aluno)
'						If decimo >= 0.75 Then
'							nota_arredondada = Int(m3_aluno) + 1
'							m3_aluno=nota_arredondada
'						elseIf decimo >= 0.25 Then
'							nota_arredondada = Int(m3_aluno) + 0.5
'							m3_aluno=nota_arredondada
'						else
'							nota_arredondada = Int(m3_aluno)
'							m3_aluno=nota_arredondada											
'						End If	
						If decimo >= 0.5 Then
							nota_arredondada = Int(m3_aluno) + 1
							m3_aluno=nota_arredondada
						else
							nota_arredondada = Int(m3_aluno)
							m3_aluno=nota_arredondada											
						End If	
						m3_aluno = formatNumber(m3_aluno,0)
						m3_aluno=m3_aluno*1
						valor_m3=valor_m3*1		
						m3_maior_igual=m3_maior_igual*1		
						if m3_aluno >= m3_maior_igual then
							resultado=res3_2
						else
							resultado=res3_1	
						end if
						if tipo_calculo="waboletim" then
							m3_waboletim=m3_aluno		
							resultado3_waboletim=resultado	
						end if	
						
						tipo_media = "MF"
						modifica_result = Verifica_Conselho_Classe(cod_aluno, cod_materia, tipo_media, outro)
						if modifica_result <> "N" then
							resultado = modifica_result
							resultado3_waboletim=resultado								
						end if													
					end if		
				end if
			end if	
		end if	
	end if

	if tipo_calculo="anual" then
		m1_aluno = formatNumber(m1_aluno,0)		
		novo_regra_aprovacao=m1_aluno&"#!#"&resultado
	elseif tipo_calculo="recuperacao" or tipo_calculo = "recuperacao_ftf" then
		if (resultado1="apr" or resultado1="rep" or resultado1="APC") and tipo_calculo="recuperacao" then
			m1_aluno = formatNumber(m1_aluno,0)	
			novo_regra_aprovacao=m1_aluno&"#!#"&resultado		
		else
			if m2_aluno<>"&nbsp;" then
				m2_aluno = formatNumber(m2_aluno,0)
			end if
			novo_regra_aprovacao=m2_aluno&"#!#"&resultado
		end if
	elseif tipo_calculo="waboletim" then
			if m2_aluno<>"&nbsp;" then
				m2_aluno = formatNumber(m2_aluno,0)
			end if	

			if m3_aluno<>"&nbsp;" then
				m3_aluno = formatNumber(m3_aluno,0)
			end if
		novo_regra_aprovacao=m1_waboletim&"#!#"&resultado1_waboletim&"#!#"&m2_waboletim&"#!#"&resultado2_waboletim&"#!#"&m3_waboletim&"#!#"&resultado3_waboletim
'		response.write(novo_regra_aprovacao&"<BR>")
	elseif tipo_calculo="waboletim_FTF" then
				if m2_waboletim<>"&nbsp;" then
					m2_waboletim = formatNumber(m2_waboletim,0)
				end if	
	
			novo_regra_aprovacao=m1_waboletim&"#!##!#"&m2_waboletim&"#!##!#"&nota_aux_m3_1&"#!#"
'			response.write(novo_regra_aprovacao&"<BR>")		
	else
		if resultado2="apr" or resultado2="rep" or resultado2="APC" then
			if m2_aluno<>"&nbsp;" then
				m2_aluno = formatNumber(m2_aluno,0)
			end if
			novo_regra_aprovacao=m2_aluno&"#!#"&resultado		
		else
			if m3_aluno<>"&nbsp;" then
				m3_aluno = formatNumber(m3_aluno,0)
			end if
			novo_regra_aprovacao=m3_aluno&"#!#"&resultado			
		end if
	end if
	
	'Session("M2")=m2_aluno
	'Session("M3")=m3_aluno
end function

Function novo_apura_resultado_aluno(curso, etapa, cod_cons, vetor_materia, vetor_medias ,caminho_nota, tb_nota, total_periodo, periodo_m2, periodo_m3, outro)

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set RSra = Server.CreateObject("ADODB.Recordset")
	SQLra = "SELECT * FROM TB_Regras_Aprovacao where CO_Curso = '"&curso&"' and CO_Etapa = '"&etapa&"'"
	RSra.Open SQLra, CON0	
			
'	valor_apr=RSra("NU_Valor_Apr")
'	valor_dep=RSra("NU_Valor_Dep")
	qtd_max_rec=RSra("NU_Disc_Ult_Periodo")
	qtd_max_dep=RSra("NU_Qt_Dis_Dep")	
'	res_apr=RSra("NO_Expr_Maior_Igual_VL_Abr")
	res_dep=RSra("NO_Expr_Cond_Verdade_Abr")
'	res_rep=RSra("NO_Expr_Cond_Falso_Abr")
	qtd_rec=0
	qtd_dep=0
	
	
	if isnull(qtd_max_rec) or qtd_max_rec="" then
		qtd_max_rec=0
	end if

	resultados_materia = split(vetor_medias, "#$#" )
	libera_resultado="s"
for rm=0 to ubound(resultados_materia)	
		nota_materia = split(resultados_materia(rm), "#!#" )
		res_aluno=nota_materia(1)
		
	if result_temp="REP" then
		libera_resultado="s"	
	else
		if res_aluno="" or isnull(res_aluno) or res_aluno="&nbsp;" or res_aluno=" "then
			resultados_rec=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_cons, vetor_materia, caminho_nota, tb_nota, 6, 5, 6, "recuperacao", 0)
			valores_result_rec=split(resultados_rec,"#!#")
			if valores_result_rec(1) = "REC" then
				result_temp=res_aluno
				qtd_rec=qtd_rec+1
			elseif res_aluno="" or isnull(res_aluno) or res_aluno="&nbsp;" or res_aluno=" "then	
				resultados_pf=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_cons, vetor_materia, caminho_nota, tb_nota, 6, 5, 6, "anual", 0)
				valores_result_pf=split(resultados_pf,"#!#")
				if valores_result_pf(1) = "REP" then	
					result_temp=res_aluno				
				else	
					libera_resultado="n"
				end if
			else
				libera_resultado="n"				
			end if	
		else
			result_temp=res_aluno
			if res_aluno = "DEP" then
				qtd_dep=qtd_dep+1	
			end if
'			
'			if res_aluno >= valor_apr then
'				result_temp="apr"
'			elseif md_aluno >= valor_dep then
'				resultado="dep"
'				qtd_dep=qtd_dep+1
'			else
'				result_temp="rep"			
'			end if
		end if
	end if	
Next
if 	libera_resultado="s" then
		novo_apura_resultado_aluno=result_temp
		if res_aluno = "DEP" then
			if qtd_dep>qtd_max_dep then
				novo_apura_resultado_aluno=res_rep	
			else	
				novo_apura_resultado_aluno=res_dep	
			end if
		end if	
		
'	if result_temp="apr" then
'		apura_resultado_aluno=res_apr
'	elseif result_temp="rep" then
'		apura_resultado_aluno=res_rep
'	elseif result_temp="dep" then	
'		qtd_dep=qtd_dep*1
'		qtd_max_dep=qtd_max_dep*1
'		if qtd_dep>qtd_max_dep then
'			apura_resultado_aluno=res_rep	
'		else	
'			apura_resultado_aluno=res_dep	
'		end if
'	end if	
else
	novo_apura_resultado_aluno="&nbsp;"		
end if	
		

end function


Function apura_resultado_aluno (curso,etapa,vetor_medias)

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set RSra = Server.CreateObject("ADODB.Recordset")
	SQLra = "SELECT * FROM TB_Regras_Aprovacao where CO_Curso = '"&curso&"' and CO_Etapa = '"&etapa&"'"
	RSra.Open SQLra, CON0	
			
'	valor_apr=RSra("NU_Valor_Apr")
'	valor_dep=RSra("NU_Valor_Dep")
	qtd_max_rec=RSra("NU_Disc_Ult_Periodo")
	qtd_max_dep=RSra("NU_Qt_Dis_Dep")	
'	res_apr=RSra("NO_Expr_Maior_Igual_VL_Abr")
	res_dep=RSra("NO_Expr_Cond_Verdade_Abr")
'	res_rep=RSra("NO_Expr_Cond_Falso_Abr")
	qtd_dep=0
	
	if isnull(qtd_max_rec) or qtd_max_rec="" then
		qtd_max_rec=0
	end if
	
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
		res_aluno=nota_materia(1)
'		response.Write(res_aluno&"<BR>")
	if result_temp="REP" then
		libera_resultado="s"	
	else
		if res_aluno="" or isnull(res_aluno) or res_aluno="&nbsp;" or res_aluno=" "then
			libera_resultado="n"
		else
			result_temp=res_aluno
			if res_aluno = "DEP" then
				qtd_dep=qtd_dep+1	
			end if
'			
'			if res_aluno >= valor_apr then
'				result_temp="apr"
'			elseif md_aluno >= valor_dep then
'				resultado="dep"
'				qtd_dep=qtd_dep+1
'			else
'				result_temp="rep"			
'			end if
		end if
	end if	
Next
if 	libera_resultado="s" then
		apura_resultado_aluno=result_temp
		if res_aluno = "DEP" then
			if qtd_dep>qtd_max_dep then
				apura_resultado_aluno=res_rep	
			else	
				apura_resultado_aluno=res_dep	
			end if
		end if	
		
'	if result_temp="apr" then
'		apura_resultado_aluno=res_apr
'	elseif result_temp="rep" then
'		apura_resultado_aluno=res_rep
'	elseif result_temp="dep" then	
'		qtd_dep=qtd_dep*1
'		qtd_max_dep=qtd_max_dep*1
'		if qtd_dep>qtd_max_dep then
'			apura_resultado_aluno=res_rep	
'		else	
'			apura_resultado_aluno=res_dep	
'		end if
'	end if	
else
	apura_resultado_aluno="&nbsp;"		
end if	
end function	

Function replace_latin_char(variavel,tipo_replace)

	if tipo_replace="html" then
		strReplacement = variavel	
		strReplacement = replace(strReplacement,"�,","&Agrave;")
		strReplacement = replace(strReplacement,"�","&Aacute;")
		strReplacement = replace(strReplacement,"�","&Acirc;")
		strReplacement = replace(strReplacement,"�","&Atilde;")
		strReplacement = replace(strReplacement,"�","&Eacute;")
		strReplacement = replace(strReplacement,"�","&Ecirc;")
		strReplacement = replace(strReplacement,"�","&Iacute;")
		strReplacement = replace(strReplacement,"�","&Oacute;")
		strReplacement = replace(strReplacement,"�","&Ocirc;")
		strReplacement = replace(strReplacement,"�","&Otilde;")
		strReplacement = replace(strReplacement,"�","&Uacute;")
		strReplacement = replace(strReplacement,"�","&Uuml;")	
		strReplacement = replace(strReplacement,"�","&agrave;")
		strReplacement = replace(strReplacement,"�","&aacute;")
		strReplacement = replace(strReplacement,"�","&acirc;")
		strReplacement = replace(strReplacement,"�","&atilde;")
		strReplacement = replace(strReplacement,"�","&ccedil;")
		strReplacement = replace(strReplacement,"�","&eacute;")
		strReplacement = replace(strReplacement,"�","&ecirc;")
		strReplacement = replace(strReplacement,"�","&iacute;")
		strReplacement = replace(strReplacement,"�","&oacute;")
		strReplacement = replace(strReplacement,"�","&ocirc;")
		strReplacement = replace(strReplacement,"�","&otilde;")
		strReplacement = replace(strReplacement,"�","&uacute;")
		strReplacement = replace(strReplacement,"�","&uuml;")			
	elseif tipo_replace="url" then
		strReplacement = Server.URLEncode(variavel)
		strReplacement = replace(strReplacement,"+"," ")
		strReplacement = replace(strReplacement,"%27","�")
		strReplacement = replace(strReplacement,"%27","'")
		strReplacement = replace(strReplacement,"%C0,","�")
		strReplacement = replace(strReplacement,"%C1","�")
		strReplacement = replace(strReplacement,"%C2","�")
		strReplacement = replace(strReplacement,"%C3","�")
		strReplacement = replace(strReplacement,"%C9","�")
		strReplacement = replace(strReplacement,"%CA","�")
		strReplacement = replace(strReplacement,"%CD","�")
		strReplacement = replace(strReplacement,"%D3","�")
		strReplacement = replace(strReplacement,"%D4","�")
		strReplacement = replace(strReplacement,"%D5","�")
		strReplacement = replace(strReplacement,"%DA","�")
		strReplacement = replace(strReplacement,"%DC","�")	
		strReplacement = replace(strReplacement,"%E1","�")
		strReplacement = replace(strReplacement,"%E1","�")
		strReplacement = replace(strReplacement,"%E2","�")
		strReplacement = replace(strReplacement,"%E3","�")
		strReplacement = replace(strReplacement,"%E7","�")
		strReplacement = replace(strReplacement,"%E9","�")
		strReplacement = replace(strReplacement,"%EA","�")
		strReplacement = replace(strReplacement,"%ED","�")
		strReplacement = replace(strReplacement,"%F3","�")
		strReplacement = replace(strReplacement,"F4","�")
		strReplacement = replace(strReplacement,"F5","�")
		strReplacement = replace(strReplacement,"%FA","�")
		strReplacement = replace(strReplacement,"%FC","�")
	end if
replace_latin_char=strReplacement
end function		
%>