<!--#include file="../../global/funcoes_diversas.asp" -->
<%
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
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

	if curso=1 and co_etapa<6 and (co_materia_verifica(n)="ARTC" or co_materia_verifica(n)="EART" or co_materia_verifica(n)="EFIS" or co_materia_verifica(n)="INGL") then									
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
		RS.Open SQL, CON0
		
	
		mae= RS("IN_MAE")
		fil= RS("IN_FIL")
		in_co= RS("IN_CO")
		peso= RS("NU_Peso")
		
		if peso = 0 then
			peso = NULL
		end if
		
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


'===========================================================================================================================================

Function periodos(periodo, tp_modelo, opcao)

if opcao="num" then
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Periodo where TP_Modelo='"&tp_modelo&"' order by NU_Periodo"
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
	SQL = "SELECT * FROM TB_Periodo where TP_Modelo='"&tp_modelo&"' AND NU_Periodo="&periodo
	RS.Open SQL, CON0
	
	vetor_periodo = RS("SG_Periodo")	

elseif opcao="nome" then

	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Periodo where TP_Modelo='"&tp_modelo&"' AND NU_Periodo="&periodo
	RS.Open SQL, CON0
	
	vetor_periodo = RS("NO_Periodo")	

elseif opcao="todas_siglas" then
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Periodo where TP_Modelo='"&tp_modelo&"' order by NU_Periodo"
	RS.Open SQL, CON0
	
	while not RS.EOF
		nu_periodo =  RS("NU_Periodo")
		sg_periodo = RS("SG_Periodo")

		
	RS.Movenext
	Wend
elseif opcao="todos_nomes" then
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Periodo where TP_Modelo='"&tp_modelo&"' order by NU_Periodo"
	RS.Open SQL, CON0
	
	while not RS.EOF
		nu_periodo =  RS("NU_Periodo")
		no_periodo = RS("NO_Periodo")		

		
	RS.Movenext
	Wend	
end if
periodos=vetor_periodo
end function

'==========================================================================================================================================
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
			freq="M"
		else
'			curso=curso*1
'			if curso<2 then
'				modelo="B"
'				freq="M"			
'			else
'				modelo="T"
'				freq="M"			
'			end if
			modelo="B"
			freq="M"
'			modelo=RS("TP_Modelo")
'			freq=RS("IN_Frequencia")
		end if
	end if
	
	if isnull(modelo) or modelo="" then
		modelo = "B"
	end if
	
	if isnull(freq) or freq="" then
		freq = "D"
	end if	
	
	if tipo_dado="tp_modelo" then
		tipo_divisao_ano=modelo
	elseif tipo_dado="in_frequencia" then
		tipo_divisao_ano=freq
	end if
end function






'===========================================================================================================================================
'Function tipo_materia(co_materia, curso, co_etapa)
'
'		Set RS = Server.CreateObject("ADODB.Recordset")
'		SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& co_materia &"'"
'		RS.Open SQL, CON0
'	
'		mae= RS("IN_MAE")
'		fil= RS("IN_FIL")
'		in_co= RS("IN_CO")
'		peso= RS("NU_Peso")
'		
'		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) then
'			tipo_materia="T_F_F_N"
'		elseif (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) then	
'			tipo_materia="T_T_F_N"
'		elseif (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then
'			tipo_materia="T_F_T_N"
'		elseif (mae=FALSE and fil=TRUE and in_co=FALSE and isnull(peso)) then	
'			tipo_materia="F_T_F_N"			
'		elseif (mae=FALSE and fil=FALSE and in_co=TRUE and isnull(peso)) then
'			tipo_materia="F_F_T_N"
'		end if	
'end function


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
Function programa_aula(vetor_materia, unidade, curso, co_etapa, turma)

	
if vetor_materia<>"nulo" then		
	co_materia= split(vetor_materia,"#!#")
	co_materia_check=1	
	For f=0 to ubound(co_materia)
	
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& co_materia(f) &"'"
		RS.Open SQL, CON0
	
		mae= RS("IN_MAE")
		fil= RS("IN_FIL")
		in_co= RS("IN_CO")
		peso= RS("NU_Peso")
		
		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) then
			if co_materia_check=1 then
				vetor_materia_exibe=co_materia(f)
			else
				vetor_materia_exibe=vetor_materia_exibe&"#!#"&co_materia(f)
			end if
			co_materia_check=co_materia_check+1			
		elseif(mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) or (mae=TRUE and fil=TRUE and in_co=FALSE) then
	
			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_materia(f) &"' order by NU_Ordem_Boletim"
			RS1.Open SQL1, CON0
				
			if RS1.EOF then
				if co_materia_check=1 then
					vetor_materia_exibe=co_materia(f)
				else
					vetor_materia_exibe=vetor_materia_exibe&"#!#"&co_materia(f)
				end if
				co_materia_check=co_materia_check+1		
			else
			co_materia_fil_check=1 
				while not RS1.EOF
					co_mat_fil= RS1("CO_Materia")				
					if co_materia_check=1 and co_materia_fil_check=1 then
						vetor_materia_exibe=co_materia(f)&"#!#"&co_mat_fil
					elseif co_materia_fil_check=1 then
						vetor_materia_exibe=vetor_materia_exibe&"#!#"&co_materia(f)&"#!#"&co_mat_fil
					else
						vetor_materia_exibe=vetor_materia_exibe&"#!#"&co_mat_fil			
					end if
					co_materia_check=co_materia_check+1
					co_materia_fil_check=co_materia_fil_check+1 									
				RS1.MOVENEXT
				wend
				vetor_materia_exibe=vetor_materia_exibe&"#!#MED"	
			end if
		elseif(mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then
			Set RS1a = Server.CreateObject("ADODB.Recordset")
			SQL1a = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_materia(f) &"' order by NU_Ordem_Boletim"
			RS1a.Open SQL1a, CON0
				
			if RS1a.EOF then
				if co_materia_check=1 then
					vetor_materia_exibe=co_materia(f)
				else
					vetor_materia_exibe=vetor_materia_exibe&"#!#"&co_materia(f)
				end if
				co_materia_check=co_materia_check+1		
			else
			co_materia_fil_check=1 
				while not RS1a.EOF
					co_mat_fil= RS1a("CO_Materia")				
					if co_materia_check=1 and co_materia_fil_check=1 then
						vetor_materia_exibe=co_materia(f)&"#!#"&co_mat_fil
					elseif co_materia_fil_check=1 then
						vetor_materia_exibe=vetor_materia_exibe&"#!#"&co_materia(f)&"#!#"&co_mat_fil
					else
						vetor_materia_exibe=vetor_materia_exibe&"#!#"&co_mat_fil			
					end if
					co_materia_check=co_materia_check+1
					co_materia_fil_check=co_materia_fil_check+1 									
				RS1a.MOVENEXT
				wend
			end if				
		end if	
	NEXT

else
end if
programa_aula=vetor_materia_exibe
end function

'===========================================================================================================================================
Function autoriza_wf(unidade, curso, co_etapa, periodo, tipo_dado, conexao, outro)

	Set RS4 = Server.CreateObject("ADODB.Recordset")
	SQL4 = "SELECT * FROM TB_Autoriza_WF where NU_Unidade="&unidade&" AND CO_Curso='"&curso&"' AND CO_Etapa='"&co_etapa&"'"	
	RS4.Open SQL4, conexao
	
	if periodo = 1 then
		variavel_T= "CO_apr1"
		variavel_P= "CO_prova1"	
	elseif periodo = 2 then
		variavel_T= "CO_apr2"
		variavel_P= "CO_prova2"	
	elseif periodo = 3 then
		variavel_T= "CO_apr3"
		variavel_P= "CO_prova3"	
	elseif periodo = 4 then
		variavel_T= "CO_apr4"
		variavel_P= "CO_prova4"	
	elseif periodo = 5 then
		variavel_T= "CO_apr5"
		variavel_P= "CO_prova5"	
	elseif periodo = 6 then
		variavel_T= "CO_apr6"
		variavel_P= "CO_prova6"	
	elseif periodo = 7 then
		variavel_T= "CO_apr7"
		variavel_P= "CO_prova7"	
	end if	

	teste=RS4(variavel_T)
	prova=RS4(variavel_P)	

	if tipo_dado="T" then
		if teste="D" then
			autoriza_wf="n"
		else
			autoriza_wf="s"
		end if		
	elseif tipo_dado="P" then
		if prova="D" then
			autoriza_wf="n"
		else
			autoriza_wf="s"
		end if		
	elseif tipo_dado="M" then	
		if teste="D"AND prova="D" then
			autoriza_wf="n"
		else
			autoriza_wf="s"
		end if		
	end if	

end Function


'===========================================================================================================================================
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
	RS.Open SQL, CON0

	mae= RS("IN_MAE")
	fil= RS("IN_FIL")
	in_co= RS("IN_CO")
	peso= RS("NU_Peso")
	
	if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) or (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) or (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then
		posicao=1
	elseif mae=False and fil=TRUE and in_co=FALSE THEN
		posicao=2
	elseif mae=False and fil=False and in_co=TRUE THEN
		posicao=0					
	end if	
end if		
posicao_materia_tabela=posicao
end function

			
'===========================================================================================================================================



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
	if operacao="menor" then
		operador=nome_nota&"<"&valor
	elseif operacao="maior" then
		operador=nome_nota&">="&valor
	elseif operacao="nulo" then
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
				media_turma=formatnumber(media_turma,1)
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
			media_mae=formatnumber(media_mae,1)
			vetor_quadro=vetor_quadro&"#!#"&media_mae	
		elseif (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then
			dados_aluno=Split(vetor_aluno,",")
			soma_medias=0
			media_somada=0
			qtd_alunos=ubound(dados_aluno)+1
		
			if qtd_alunos=0 then
				qtd_alunos=1
			end if
			conta_aluno=0	
			for al=0 to ubound(dados_aluno)
				vetor_materia_filhas=busca_materias_filhas(co_materia(fb))	
				dividendo=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, dados_aluno(al),  co_materia(fb), vetor_materia_filhas, caminho_nota, tb_nota, periodo, nome_nota, outro)	

				dividendo_asterisco=Calcula_Asterisco(tp_modelo, tp_freq, unidade, curso, co_etapa, turma, dados_aluno(al), co_materia(fb), caminho_nota, "T_F_T_N", tb_nota, periodo)
		
'				medias=Split(medias_bimestres,"#!#")			
				
'				periodo=periodo*1
'							
'				if periodo = 1 then
'					dividendo=medias(0)
'					dividendo_asterisco=medias(4)
'				elseif periodo = 2 then	
'					dividendo=medias(1)
'					dividendo_asterisco=medias(5)
'				elseif periodo = 3 then	
'					dividendo=medias(6)
'					dividendo_asterisco=medias(10)
'				elseif periodo = 4 then	
'					dividendo=medias(7)
'					dividendo_asterisco=medias(11)
'				elseif periodo = 5 then	
'					dividendo=medias(12)
'					dividendo_asterisco=0
'				end if					
								
				if dividendo<>"&nbsp;" then	
					if dividendo_asterisco<>"&nbsp;" then							
						if dividendo>dividendo_asterisco then
							verifica_medias=dividendo
						else
							verifica_medias=dividendo_asterisco	
						end if	
					else
						verifica_medias=dividendo						
					end if
					media_somada=media_somada+1
				else
					media_somada=media_somada					
				end if		
			next	

			verifica_medias=verifica_medias*1
			valor=valor*1			

			if operacao="menor" then
				if verifica_medias<valor then
					conta_aluno=conta_aluno+1
				else
					conta_aluno=conta_aluno				
				end if
			elseif operacao="maior" then
				if verifica_medias>=valor then
					conta_aluno=conta_aluno+1
				else
					conta_aluno=conta_aluno				
				end if
			end if				
			
			if co_materia_check=1 then
				vetor_quadro=conta_aluno
			else
				vetor_quadro=vetor_quadro&"#!#"&conta_aluno	
			end if	
			
		elseif (mae=FALSE and fil=FALSE and in_co=TRUE) then

			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT Count("&tb_nota&"."&nome_nota&")AS QtdDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& vetor_aluno &") AND CO_Materia ='"& co_materia(i)&"' And "&operador&" And NU_Periodo="&periodo
			RS1.Open SQL1, CONn
			
			media_turma=RS1("QtdDeVA_Media3")
			if media_turma="" or isnull(media_turma) then

			end if 
			
			if co_materia_check=1 then
				vetor_quadro=media_turma
			else
				vetor_quadro=vetor_quadro&"#!#"&media_turma
			end if						

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
			media_mae=formatnumber(media_mae,1)
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
						if conta_aluno=1 then
						aluno_nulo=aluno(n)
						else
						aluno_nulo=aluno_nulo&"#!#"&aluno(n)
						end if
					else
						media_aluno=RS3("VA_Media3")
						if media_aluno="" or isnull(media_aluno) then
							conta_aluno=conta_aluno+1	
							if conta_aluno=1 then
							aluno_nulo=aluno(n)
							else
							aluno_nulo=aluno_nulo&"#!#"&aluno(n)
							end if
						end if
					end if					
				next
			Next
			if co_materia_check=1 then
				vetor_quadro=conta_aluno
			else
				vetor_quadro=vetor_quadro&"#!#"&conta_aluno
			end if	
		elseif (mae=FALSE and fil=FALSE and in_co=TRUE) then
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
		end if		

	co_materia_check=co_materia_check+1			
	NEXT
else
end if
Session("aluno_nulo")=aluno_nulo
conta_medias=vetor_quadro&"#$#"
'response.Write(calcula_medias)
end function



















'===========================================================================================================================================

Function calcula_medias(unidade, curso, co_etapa, turma, periodo, fn_vetor_aluno, fn_vetor_materia, caminho_nota, tb_nota, nome_nota, tipo_calculo)


	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CONn = Server.CreateObject("ADODB.Connection") 
	ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONn.Open ABRIRn	
	
	Set RSra = Server.CreateObject("ADODB.Recordset")
	SQLra = "SELECT * FROM TB_Regras_Aprovacao where CO_Curso = '"&curso&"' and CO_Etapa = '"&co_etapa&"'"
	RSra.Open SQLra, CON0	
		
	res_apr=RSra("NU_Valor_M1")
	res_rec=RSra("NU_Valor_M2")
	res_rep=RSra("NU_Valor_M3")

if tipo_calculo="media_turma" then	
	co_materia= split(fn_vetor_materia,"#!#")	
	co_materia_check=1	

	For fb=0 to ubound(co_materia)
	
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& co_materia(fb) &"'"
		RS.Open SQL, CON0
	
		mae= RS("IN_MAE")
		fil= RS("IN_FIL")
		in_co= RS("IN_CO")
		peso= RS("NU_Peso")
		
		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) then

			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& fn_vetor_aluno &") AND CO_Materia ='"& co_materia(fb)&"' And NU_Periodo="&periodo
			RS1.Open SQL1, CONn
			
			if RS1.EOF then
				media_turma=""
			else
				media_turma=RS1("MediaDeVA_Media3")
	
				if media_turma="" or isnull(media_turma) then
				else
				media_turma=formatnumber(media_turma,1)
				end if 
			end if
			if co_materia_check=1 then
				vetor_quadro=media_turma
			else
				vetor_quadro=vetor_quadro&"#!#"&media_turma
			end if
				
		elseif (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) then
		
			vetor_mae_filhas=""
	
			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL2 = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_materia(fb) &"'"
			RS2.Open SQL2, CON0
				
			co_materia_fil_check=0 
			while not RS2.EOF
				co_mat_fil= RS2("CO_Materia")				
				if co_materia_fil_check=0 then
					vetor_mae_filhas=co_materia(fb)&"#!#"&co_mat_fil
				else
					vetor_mae_filhas=vetor_mae_filhas&"#!#"&co_mat_fil			
				end if
				co_materia_fil_check=co_materia_fil_check+1 									
			RS2.MOVENEXT
			wend				
			
			co_materia_mae_fil= split(vetor_mae_filhas,"#!#")
			media_mae_acumula=0			
			for f3=0 to ubound(co_materia_mae_fil)			
		
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL3 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& fn_vetor_aluno &") AND CO_Materia ='"& co_materia_mae_fil(f3)&"' And NU_Periodo="&periodo
				RS3.Open SQL3, CONn

'response.Write(media_mae_acumula)					
				media_turma=RS3("MediaDeVA_Media3")
				if media_turma="" or isnull(media_turma) then
				media_filha_acumula=0	
				else
				media_turma=formatnumber(media_turma,1)
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
				'response.Write(co_materia_mae_fil(f3)&"-"&media_turma&"-"&media_mae_acumula&"-"&co_materia_fil_check&"<BR>")		
			next
			media_mae=media_mae_acumula/co_materia_fil_check
			media_mae=arredonda(media_mae,"mat_dez",1,0)
			vetor_quadro=vetor_quadro&"#!#"&media_mae	
			
		elseif (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then

			dados_aluno=Split(fn_vetor_aluno,",")
			soma_medias=0
			media_somada=0
			qtd_alunos=ubound(dados_aluno)+1
		
			if qtd_alunos=0 then
				qtd_alunos=1
			end if
			
			tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")		
			nome_nota=var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo,"BDM")
			
			for al=0 to ubound(dados_aluno)
				vetor_materia_filhas=busca_materias_filhas(co_materia(fb))	
				dividendo=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, dados_aluno(al), co_materia(fb), vetor_materia_filhas, CONn, tb_nota, periodo, nome_nota, outro)	
				dividendo_asterisco=Calcula_Asterisco(tp_modelo, tp_freq, unidade, curso, co_etapa, turma, dados_aluno(al), co_materia(fb), CONn, "T_F_T_N", tb_nota, periodo)	
		'	response.Write(medias_bimestres&", "&co_materia(fb))				
				
								
				if dividendo<>"&nbsp;" then	
					if dividendo_asterisco<>"&nbsp;" then							
						if dividendo>dividendo_asterisco then
							soma_medias=soma_medias+dividendo
						else
							soma_medias=soma_medias+dividendo_asterisco	
						end if	
					else
						soma_medias=soma_medias+dividendo						
					end if
					media_somada=media_somada+1
				else
					soma_medias=soma_medias
					media_somada=media_somada					
				end if		
			next	
			if media_somada=0 then
				media_mae="&nbsp;"
			else
				media_mae=soma_medias/media_somada
				media_mae=arredonda(media_mae,parametros_gerais("arred_media"),parametros_gerais("decimais_media"),0)			
			end if
			if co_materia_check=1 then
				vetor_quadro=media_mae
			else
				vetor_quadro=vetor_quadro&"#!#"&media_mae	
			end if	
			
		elseif (mae=FALSE and fil=FALSE and in_co=TRUE) then

			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& fn_vetor_aluno &") AND CO_Materia ='"& co_materia(fb)&"' And NU_Periodo="&periodo
			RS1.Open SQL1, CONn
			
			media_turma=RS1("MediaDeVA_Media3")
			if media_turma="" or isnull(media_turma) then
			else
			media_turma=formatnumber(media_turma,1)
			end if 
				if co_materia_check=1 then
					vetor_quadro=media_turma
				else
					vetor_quadro=vetor_quadro&"#!#"&media_turma
				end if							
		end if		
	co_materia_check=co_materia_check+1			
	NEXT
calcula_medias=vetor_quadro&"#$#"	








elseif tipo_calculo="media_geral" then	
	
			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& fn_vetor_aluno &") And NU_Periodo="&periodo
			RS1.Open SQL1, CONn
			
			media_turma=RS1("MediaDeVA_Media3")
			if media_turma="" or isnull(media_turma) then
				media_turma=0
			else
				media_turma=formatnumber(media_turma,1)
			end if 

			vetor_quadro=media_turma
			
calcula_medias=vetor_quadro&"#$#"		














		
elseif tipo_calculo="boletim" then	

cod_aluno=fn_vetor_aluno


	co_materia= split(fn_vetor_materia,"#!#")	
	co_materia_check=0	

	vetor_periodo= split(periodo,"#!#")	

	For fb=0 to ubound(co_materia)
	soma=0	
	
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& co_materia(fb) &"'"
		RS.Open SQL, CON0
	
		mae= RS("IN_MAE")
		fil= RS("IN_FIL")
		in_co= RS("IN_CO")
		peso= RS("NU_Peso")
		
		calcula_media_anual="sim"	
			
'or (mae=FALSE and fil=FALSE and in_co=TRUE) só serve para o Mapa de Resultados por Disciplinas		
		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) or (mae=FALSE and fil=FALSE and in_co=TRUE)  then
	
			medias_bimestres=Calc_Med_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, co_materia(fb), caminho_nota, tb_nota, 5, 5, 0,"nulo", "nulo")		
			medias__split=medias_bimestres
			medias=Split(medias__split,"#!#")
					
			dividendo1=medias(0)
			dividendo2=medias(1)
			dividendo3=medias(6)
			dividendo4=medias(7)
			dividendo5=medias(12)	

			
			if dividendo1<>"&nbsp;" and dividendo2<>"&nbsp;" and dividendo3<>"&nbsp;" and dividendo4<>"&nbsp;" then
				media_res=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_aluno, co_materia(fb), caminho_nota, tb_nota, 5, 5, 0, "boletim", 0)			
				resultados=medias_bimestres&"#!#"&media_res							
			else
				resultados=medias_bimestres&"#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"							
			end if															
				
			if co_materia_check=0 then
				vetor_quadro=resultados
			else	
				vetor_quadro=vetor_quadro&"#$#"&resultados
			end if				

		elseif (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) then
			
		elseif (mae=TRUE and fil=TRUE and in_co=FALSE) then
			
		elseif (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then
							
			medias_bimestres=Calc_Med_T_F_T_N(unidade, curso, co_etapa, turma, cod_aluno, co_materia(fb), caminho_nota, tb_nota, 5, 5, 0,"nulo", "nulo")			
			
			medias=Split(medias_bimestres,"#!#")
				
			dividendo1=medias(0)
			dividendo2=medias(1)
			dividendo3=medias(6)
			dividendo4=medias(7)

			if dividendo1<>"&nbsp;" and dividendo2<>"&nbsp;" and dividendo3<>"&nbsp;" and dividendo4<>"&nbsp;" then
				media_res=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_aluno, co_materia(fb), caminho_nota, tb_nota, 5, 5, 0, "boletim", 0)			
				resultados=medias_bimestres&"#!#"&media_res			
									
			else
				resultados=medias_bimestres&"#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"			
			end if						
																					
					
			if co_materia_check=0 then
				vetor_quadro=resultados
			else	
				vetor_quadro=vetor_quadro&"#$#"&resultados
			end if				
	
		end if		
		co_materia_check=co_materia_check+1			
	NEXT	
calcula_medias=vetor_quadro	
end if

end function







'===========================================================================================================================================

'calcula as médias anuais e finais destes respectivos mapas
Function Calc_Med_An_Fin(unidade, curso, co_etapa, turma, fn_vetor_aluno, vetor_materia, caminho_nota, tb_nota, qtd_periodos, periodo_m2, periodo_m3,tipo_calculo, outro)

Server.ScriptTimeout = 900

if periodo_m2=0 then
	retira_periodo_m2=0
else
	retira_periodo_m2=1
end if

if periodo_m3=0 then
	retira_periodo_m3=0
else
	retira_periodo_m3=1
end if

total_periodos_m1=qtd_periodos-retira_periodo_m2-retira_periodo_m3

'response.Write(vetor_materia&"<BR>")
'response.Write(vetor_aluno&"<BR>")
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CONn = Server.CreateObject("ADODB.Connection") 
	ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONn.Open ABRIRn	

	
	Set RSra = Server.CreateObject("ADODB.Recordset")
	SQLra = "SELECT * FROM TB_Regras_Aprovacao where CO_Curso = '"&curso&"' and CO_Etapa = '"&co_etapa&"'"
	RSra.Open SQLra, CON0	
		
	res_apr=RSra("NU_Valor_M1")
	res_rec=RSra("NU_Valor_M2")
	res_rep=RSra("NU_Valor_M3")
	
	alunos= split(vetor_aluno,"#$#")			
	cod_materia= split(vetor_materia,"#!#")	
	co_materia_check=1	
	co_matricula= vetor_aluno
	quantidade_alunos=0
	peso_periodo_acumulado=0
	
	for periodo=1 to total_periodos_m1
		Set RSPESO = Server.CreateObject("ADODB.Recordset")
		SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&periodo
		RSPESO.Open SQLPESO, CON0

		if RSPESO.EOF then	
			peso_periodo_acumulado=peso_periodo_acumulado						
		else
			peso_periodo=RSPESO("NU_Peso")
			if isnull(peso_periodo) or 	peso_periodo="" then
				peso_periodo_acumulado=peso_periodo_acumulado+1	
			else		
				peso_periodo_acumulado=peso_periodo_acumulado+peso_periodo				
			end if
		end if
	Next		
		
	For a=0 to ubound(alunos)
		dados_aluno= split(alunos(a),"#!#")	
		quantidade_materias=0
		For c=0 to ubound(cod_materia)

			Set RS = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& cod_materia(c) &"'"
			RS.Open SQL, CON0
		
			mae= RS("IN_MAE")
			fil= RS("IN_FIL")
			in_co= RS("IN_CO")
			peso= RS("NU_Peso")		
			
			media_acumulada=0

			if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) then
				medias_bimestres=Calc_Med_T_F_F_N(unidade, curso, co_etapa, turma, dados_aluno(0), cod_materia(c), caminho_nota, tb_nota, 5, 5, 0,"nulo", "nulo")	
				
				'response.Write(medias_bimestres&"<BR>")
				medias=Split(medias_bimestres,"#!#")			
					
				dividendo1=medias(0)
				dividendo2=medias(1)
				dividendo3=medias(6)
				dividendo4=medias(7)
				dividendo1a=medias(4)
				dividendo2a=medias(5)
				dividendo3a=medias(10)
				dividendo4a=medias(11)		
				dividendo5=medias(12)							
				media_acumulada=0				
			
				if dividendo1a="&nbsp;" then
					if dividendo1<>"&nbsp;" then
						dividendo1soma=dividendo1*1	
					else
						media_acumulada="&nbsp;"
					end if								
				else
'					if dividendo1>dividendo1a then
'						dividendo1soma=dividendo1*1
'					else
						dividendo1soma=dividendo1a*1
'					end if
				end if	
					
				if dividendo2a="&nbsp;" then
					if dividendo2<>"&nbsp;" then
						dividendo2soma=dividendo2*1
					else
						media_acumulada="&nbsp;"						
					end if	
				else				
'					if dividendo2>dividendo2a then
'						dividendo2soma=dividendo2*1
'					else
						dividendo2soma=dividendo2a*1
'					end if
				end if	

				if dividendo3a="&nbsp;" or dividendo3a="" or isnull(dividendo3a) then
					if dividendo3<>"&nbsp;" then
						dividendo3soma=dividendo3*1	
					else
						media_acumulada="&nbsp;"						
					end if					
				else										
'					if dividendo3>dividendo3a then
'						dividendo3soma=dividendo3*1
'					else
'response.Write("'"&dividendo3a&"'")
						dividendo3soma=dividendo3a*1
'					end if

				end if	

				if dividendo4a="&nbsp;" or dividendo4a="" or isnull(dividendo4a) then
					if dividendo4<>"&nbsp;" then
						dividendo4soma=dividendo4*1		
					else
						media_acumulada="&nbsp;"						
					end if						
				else									
'					if dividendo4>dividendo4a then
'						dividendo4soma=dividendo4*1
'					else
						dividendo4soma=dividendo4a*1
'					end if																								
				end if	
							

				if media_acumulada="&nbsp;" then
					if tipo_calculo="boletim" then
						resultado_materia="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
					else				
						resultado_materia="&nbsp;#!#&nbsp;"
					end if									
				else							
					media_acumulada=dividendo1soma+dividendo2soma+dividendo3soma+dividendo4soma
				'	response.Write(media_acumulada&"="&dividendo1soma&"+"&dividendo2soma&"+"&dividendo3soma&"+"&dividendo4soma&"/"&peso_periodo_acumulado)
					media_anual=media_acumulada/peso_periodo_acumulado		
					media_anual = arredonda(media_anual,"mat_dez",1,0)
				
					if tipo_calculo="anual" then
						resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","anual")
						media_anual = formatNumber(media_anual,1)
						resultado_materia=media_anual&"#!#"&resultado					
					elseif tipo_calculo="boletim" then
						if dividendo5="&nbsp;" then
							resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","boletim")	
						else
				
							resultado=regra_aprovacao(curso,co_etapa,media_anual,dividendo5,"&nbsp;","&nbsp;","&nbsp;","boletim")	
													
						end if					
						resultado_materia=resultado												
					else
						if dividendo5="&nbsp;" then
							resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","final")
						else	
							resultado=regra_aprovacao(curso,co_etapa,media_anual,dividendo5,"&nbsp;","&nbsp;","&nbsp;","final")					
						end if
						resultado_materia=resultado
					end if	
				end if	
										
			elseif (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) then
			
			elseif (mae=TRUE and fil=TRUE and in_co=FALSE) then			
			
			elseif (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then	
			
				medias_bimestres=Calc_Med_T_F_T_N(unidade, curso, co_etapa, turma, dados_aluno(0) , cod_materia(c), caminho_nota, tb_nota, 5, 5, 0,"nulo", "nulo")							
					
						
				medias=Split(medias_bimestres,"#!#")
					
				dividendo1=medias(0)
				dividendo2=medias(1)
				dividendo3=medias(6)
				dividendo4=medias(7)
				dividendo1a=medias(4)
				dividendo2a=medias(5)
				dividendo3a=medias(10)
				dividendo4a=medias(11)		
				dividendo5=medias(12)							
				media_acumulada=0				
			
				if dividendo1a="&nbsp;" then
					if dividendo1<>"&nbsp;" then
						dividendo1soma=dividendo1*1	
					else
						media_acumulada="&nbsp;"
					end if								
				else
					if dividendo1>dividendo1a then
						dividendo1soma=dividendo1*1
					else
						dividendo1soma=dividendo1a*1
					end if
				end if	
					
				if dividendo2a="&nbsp;" then
					if dividendo2<>"&nbsp;" then
						dividendo2soma=dividendo2*1
					else
						media_acumulada="&nbsp;"						
					end if	
				else				
					if dividendo2>dividendo2a then
						dividendo2soma=dividendo2*1
					else
						dividendo2soma=dividendo2a*1
					end if
				end if	

				if dividendo3a="&nbsp;" then
					if dividendo3<>"&nbsp;" then
						dividendo3soma=dividendo3*1	
					else
						media_acumulada="&nbsp;"						
					end if					
				else										
					if dividendo3>dividendo3a then
						dividendo3soma=dividendo3*1
					else
						dividendo3soma=dividendo3a*1
					end if
				end if	

				if dividendo4a="&nbsp;" then
					if dividendo4<>"&nbsp;" then
						dividendo4soma=dividendo4*1		
					else
						media_acumulada="&nbsp;"						
					end if						
				else									
					if dividendo4>dividendo4a then
						dividendo4soma=dividendo4*1
					else
						dividendo4soma=dividendo4a*1
					end if																								
				end if	
																	
				if media_acumulada="&nbsp;" then
					if tipo_calculo="boletim" then
						resultado_materia="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
					else				
						resultado_materia="&nbsp;#!#&nbsp;"
					end if									
				else
					media_acumulada=dividendo1soma+dividendo2soma+dividendo3soma+dividendo4soma
					media_anual=media_acumulada/4		
					media_anual = arredonda(media_anual,"mat_dez",1,0)	
		
					if tipo_calculo="anual" then
						resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","anual")
						media_anual = formatNumber(media_anual,1)
						resultado_materia=media_anual&"#!#"&resultado					
	
					elseif tipo_calculo="boletim" then
						if dividendo5="&nbsp;" then
							resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","boletim")	
						else
							resultado=regra_aprovacao(curso,co_etapa,media_anual,dividendo5,"&nbsp;","&nbsp;","&nbsp;","boletim")							
						end if				
	
						resultado_materia=resultado												
					else
						if dividendo5="&nbsp;" then
							resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","final")
						else	
							resultado=regra_aprovacao(curso,co_etapa,media_anual,dividendo5,"&nbsp;","&nbsp;","&nbsp;","final")					
						end if
						resultado_materia=resultado
					end if	
				end if													
					
			elseif (mae=FALSE and fil=FALSE and in_co=TRUE ) then	
				if tipo_calculo="boletim" then
					resultado_materia="&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"
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
'response.Write(dados_aluno(1)&"-"&tipo_calculo&"-"&media_anual&"-"&rec_lancado&"-"&md&"-"&media_rec&"<BR>")

	next
'response.Write(resultado_turma)
'response.End()
Calc_Med_An_Fin=resultado_turma		
END FUNCTION

Function regra_aprovacao (curso,etapa,m1_aluno,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,tipo_calculo)

Server.ScriptTimeout = 900

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
	res3_1=RSra("NO_Expr_Int_M3_V")
	res3_2=RSra("NO_Expr_Ma_Igual_M3")

		
	m1_aluno=m1_aluno*1	
	m1_maior_igual=m1_maior_igual*1
	m1_menor=m1_menor*1

	if m1_aluno >= m1_maior_igual then
		resultado=res1_3
	elseif m1_aluno >= m1_menor then
		resultado=res1_2
	else
		resultado=res1_1	
	end if
	
	if resultado=res1_3 then
		if tipo_calculo="boletim" then	
			m2_aluno="&nbsp;"
		else
			m2_aluno=m1_aluno	
		end if	
	else
		if tipo_calculo="boletim" then	
			if nota_aux_m2_1="&nbsp;" then
				m2_aluno="&nbsp;"	
			else		
				m1_aluno_peso=m1_aluno*peso_m2_m1
				nota_aux_m2_1_peso=nota_aux_m2_1*peso_m2_m2
				m2_aluno=(m1_aluno_peso+nota_aux_m2_1_peso)/(peso_m2_m1+peso_m2_m2)		
				m2_aluno=m2_aluno*1
				m2_maior_igual=m2_maior_igual*1
				if m2_aluno >= m2_maior_igual then
					resultado=res2_3
				elseif m2_aluno >= m2_menor then
					resultado=res2_2
				else
					resultado=res2_1	
				end if
				m2_aluno = arredonda(m2_aluno,"mat_dez",1,0)
			end if		
		elseif tipo_calculo="final" then
			if nota_aux_m2_1="&nbsp;" then
				m2_aluno="&nbsp;"
				resultado="&nbsp;"	
			else								
				m1_aluno_peso=m1_aluno*peso_m2_m1
				nota_aux_m2_1_peso=nota_aux_m2_1*peso_m2_m2
				m2_aluno=(m1_aluno_peso+nota_aux_m2_1_peso)/(peso_m2_m1+peso_m2_m2)
							
				m2_aluno=m2_aluno*1
				m2_maior_igual=m2_maior_igual*1	
	
				if m2_aluno >= m2_maior_igual then
					resultado=res2_3
				elseif m2_aluno >= m2_menor then
					resultado=res2_2
				else
					resultado=res2_1	
				end if
				m2_aluno = arredonda(m2_aluno,"mat_dez",1,0)
			end if
		end if
	end if
	if tipo_calculo="anual" then
		regra_aprovacao=resultado

	elseif tipo_calculo="boletim" then	

		regra_aprovacao=m1_aluno&"#!#"&nota_aux_m2_1&"#!#"&m2_aluno&"#!#"&resultado
		
	else
		'if m2_aluno<>"&nbsp;" then
		'	m2_aluno = formatNumber(m2_aluno,1)
		'end if
		regra_aprovacao=m2_aluno&"#!#"&resultado	
	end if
	
	'Session("M2")=m2_aluno
	'Session("M3")=m3_aluno
end function			
Function novo_apura_resultado_aluno (p_curso,p_etapa,cod_aluno,vetor_disciplinas,caminho_nota,tb_nota,prd_ter_media,tipo_calculo,outro)

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1	
	
	Set RSra = Server.CreateObject("ADODB.Recordset")
	SQLra = "SELECT * FROM TB_Regras_Aprovacao where CO_Curso = '"&p_curso&"' and CO_Etapa = '"&p_etapa&"'"
	RSra.Open SQLra, CON0	
			
'	valor_apr=RSra("NU_Valor_Apr")
'	valor_dep=RSra("NU_Valor_Dep")
	qtd_max_dep=RSra("NU_Qt_Dis_Dep")
	res_apr=RSra("NO_Expr_Maior_Igual_VL_Abr")
	res_dep=RSra("NO_Expr_Cond_Verdade_Abr")	
	res_rep=RSra("NO_Expr_Cond_Falso_Abr")
	qtd_rec=0	
	qtd_dep=0
'	valor_apr=70
'	valor_dep=50
'	qtd_max_dep=5
'	res_apr="AP"
'	res_dep="DP"
'	res_rep="RP"	
'response.Write(vetor_medias)
'response.End()
'		if cod_aluno=803 then
'		response.Write(qtd_dep&"-")
'		end if
	disciplinas = split(vetor_disciplinas, "#!#" )
	libera_resultado="s"
	prd_seg_media=prd_ter_media-1
	
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& cod_aluno &" and TB_Matriculas.NU_Ano="&ano_letivo
		RS.Open SQL, CON1	
			
		a_unidade= RS("NU_Unidade")
		a_curso= RS("CO_Curso")
		a_etapa= RS("CO_Etapa")
		a_turma= RS("CO_Turma")		
	
for r=0 to ubound(disciplinas)	
		resultado_2= Calc_Seg_Media (a_unidade, a_curso, a_etapa, a_turma, cod_aluno, disciplinas(r), caminho_nota, tb_nota, prd_seg_media, tipo_calculo, outro)
		teste_res=split(resultado_2, "#!#")
		if teste_res(1)="REC" then
			qtd_rec=qtd_rec+1
		end if
		
		resultados = Calc_Ter_Media (a_unidade, a_curso, a_etapa, a_turma, cod_aluno, disciplinas(r), caminho_nota, tb_nota, prd_ter_media, tipo_calculo, outro)
		nota_materia = split(resultados, "#!#" )
		res_aluno=nota_materia(1)
'		if cod_aluno=803 then
'		response.Write(res_aluno&"-")
'		end if

	if res_aluno="REP"  then
		qtd_dep=qtd_dep+1		
		result_rep="s"	
	else
		if res_aluno="" or isnull(res_aluno) or res_aluno="&nbsp;" or res_aluno=" "then
			Set RSm = Server.CreateObject("ADODB.Recordset")
			SQLm = "SELECT * FROM TB_Materia where CO_Materia ='"& disciplinas(r) &"'"
			RSm.Open SQLm, CON0				
			
			disc_obrigatoria=RSm("IN_Obrigatorio")
			
			if disc_obrigatoria=TRUE then
				libera_resultado="n"
			end if	
		else
'			libera_resultado="s"
			'result_temp=res_aluno
'			if res_aluno = "DEP" then
'				qtd_dep=qtd_dep+1		
'			end if
'			result_temp=res_aluno			
		end if
	end if	
'IF cod_aluno=827 THEN
'RESPONSE.Write(disciplinas(r)&"-"&res_aluno&"-"&libera_resultado&"<br>")
'END IF
Next
'IF cod_aluno=827 THEN
'RESPONSE.END()
'END IF

if libera_resultado="s" or result_rep="s" then
'	if result_temp = "DEP" then
'		if qtd_dep>qtd_max_dep then
'			apura_resultado_aluno=res_rep	
'		else	
'			apura_resultado_aluno=res_dep	
'		end if
'	else
'		apura_resultado_aluno=result_temp	
'	end if	
		if qtd_rec>2 then
			novo_apura_resultado_aluno=res_rep	
		elseif result_rep="s" then
			if qtd_dep>qtd_max_dep then
				novo_apura_resultado_aluno=res_rep	
			else	
				novo_apura_resultado_aluno=res_dep	
			end if
		else
			novo_apura_resultado_aluno=res_apr					
		end if	
else
	if qtd_rec>2 then
		novo_apura_resultado_aluno=res_rep	
	else
		novo_apura_resultado_aluno="&nbsp;"	
	end if		
end if	
'		if cod_aluno=803 then
'		response.Write(apura_resultado_aluno)
'		response.end()
'		end if
end function		
Function apura_resultado_aluno (curso,etapa,vetor_medias)

	Set conexao = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	conexao.Open ABRIR
	
	Set RSano = Server.CreateObject("ADODB.Recordset")
    SQLano = "SELECT ST_Ano_Letivo FROM TB_Ano_Letivo WHERE NU_Ano_Letivo = '"&session("ano_letivo")&"'"
	RSano.Open SQLano, conexao

	 bloqueia_resultado="N"
	
	IF RSano.EOF THEN
	 bloqueia_resultado="S"
	ELSEIF RSano("ST_Ano_Letivo")<>"L" THEN
	 bloqueia_resultado="S"	
	END IF

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set RSra = Server.CreateObject("ADODB.Recordset")
	SQLra = "SELECT * FROM TB_Regras_Aprovacao where CO_Curso = '"&curso&"' and CO_Etapa = '"&etapa&"'"
	RSra.Open SQLra, CON0	
			
	qtd_max_dep=RSra("NU_Qt_Dis_Dep")
	res_apr=RSra("NO_Expr_Ma_Igual_M1")
	res_rec=RSra("NO_Expr_Int_M1_F")
	res_pfi=RSra("NO_Expr_Int_M2_F")
	res_pfi_falso=RSra("NO_Expr_Int_M2_F")
	res_rep=RSra("NO_Expr_Int_M3_F")
	res_dep=RSra("NO_Expr_Cond_Verdade_Abr")
	res_dep_falso=RSra("NO_Expr_Cond_Falso_Abr")
	qtd_pfi=0
	qtd_dep=0
	
	curso=curso*1
	etapa=etapa*1
	if curso=1 and etapa>5 then
		qtd_max_pfi=3
	elseif curso=2 then
		qtd_max_pfi=4	
	end if	

	resultados_materia = split(vetor_medias, "#!#" )
	libera_resultado="s"
	
for r=0 to ubound(resultados_materia)	
		res_disciplina=resultados_materia(r)
'		response.Write(r&"-"&res_disciplina&"-"&result_temp&"-"&libera_resultado&"<BR>")
	if res_disciplina="" or isnull(res_disciplina) or res_disciplina="&nbsp;" or res_disciplina=" "then
		libera_resultado="n"
	else
		if result_temp="REP" then
		else
			if result_temp="REC" then
				if res_disciplina="REP" then	
					result_temp=res_disciplina
				end if			
			else
				if result_temp="PFI" then	
					if res_disciplina="REP" or res_disciplina="REC" then	
						result_temp=res_disciplina
					end if					
				else	
					result_temp=res_disciplina
				end if
			end if	
		end if					
		if res_disciplina = "REP" then
			qtd_dep=qtd_dep+1		
		end if
		if res_disciplina = "PFI" then
			qtd_pfi=qtd_pfi+1		
		end if		
	End if	
Next
if 	libera_resultado="s" then
'		resultado_aluno=result_temp
'		if res_aluno = "DEP" then
'			if qtd_dep>qtd_max_dep then
'				resultado_aluno=res_rep	
'			else	
'				resultado_aluno=res_dep	
'			end if
'		end if	
		
	if result_temp="APR" then
		resultado_aluno=res_apr
	elseif result_temp="REC" then
		resultado_aluno=res_rec
	elseif result_temp="PFI" then
		'resultado_aluno=res_pfi
		qtd_pfi=qtd_pfi*1
		qtd_max_pfi=qtd_max_pfi*1		
		if qtd_pfi>qtd_max_pfi then
			resultado_aluno=res_pfi_falso	
		else	
			resultado_aluno=res_pfi	
		end if					
	elseif result_temp="REP" then

		qtd_dep=qtd_dep*1
		qtd_max_dep=qtd_max_dep*1
		if qtd_dep>qtd_max_dep then
			resultado_aluno=res_dep_falso	
		else	
			resultado_aluno=res_dep	
		end if
	end if	
else
	resultado_aluno="&nbsp;"		
end if	
'response.Write(bloqueia_resultado&" "&result_temp&" "&resultado_aluno&" "&qtd_pfi&" "&qtd_max_pfi&" "&res_pfi_falso&" "&res_pfi&" "&qtd_dep&" "&qtd_max_dep&" "&res_dep_falso&" "&res_dep)
if bloqueia_resultado="N" then
	apura_resultado_aluno = resultado_aluno
else
	if resultado_aluno="&nbsp;" then
		apura_resultado_aluno = "REP"	
	elseif resultado_aluno="PFI" then
		'resultado_aluno=res_pfi
		qtd_pfi=qtd_pfi*1
		qtd_max_pfi=qtd_max_pfi*1		
		if qtd_pfi>qtd_max_pfi then
			apura_resultado_aluno=res_pfi_falso	
		else	
			apura_resultado_aluno=res_pfi	
		end if					
	elseif resultado_aluno="REP" then

		qtd_dep=qtd_dep*1
		qtd_max_dep=qtd_max_dep*1
		if qtd_dep>qtd_max_dep then
			apura_resultado_aluno=res_dep_falso	
		else	
			apura_resultado_aluno=res_dep	
		end if	
	else
		apura_resultado_aluno = resultado_aluno	
	end if
end if	

'response.Write("AR:"&apura_resultado_aluno)
'response.End()
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
		strReplacement = replace(strReplacement,"%2E",".")		
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
		strReplacement = replace(strReplacement,"%F4","ô")
		strReplacement = replace(strReplacement,"%F5","õ")
		strReplacement = replace(strReplacement,"%FA","ú")
		strReplacement = replace(strReplacement,"%FC","ü")
	end if
replace_latin_char=strReplacement
end function
%>