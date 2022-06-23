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
		
		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) or (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) or (mae=TRUE and fil=TRUE and in_co=FALSE) then
	
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
'				while not RS1.EOF
'					co_mat_fil= RS1("CO_Materia")				
'					if co_materia_check=1 and co_materia_fil_check=1 then
'						vetor_materia_exibe=co_materia(f)&"#!#"&co_mat_fil
'					elseif co_materia_fil_check=1 then
'						vetor_materia_exibe=vetor_materia_exibe&"#!#"&co_materia(f)&"#!#"&co_mat_fil
'					else
'						vetor_materia_exibe=vetor_materia_exibe&"#!#"&co_mat_fil			
'					end if
'					co_materia_check=co_materia_check+1
'					co_materia_fil_check=co_materia_fil_check+1 									
'				RS1.MOVENEXT
'				wend
				if co_materia_check=1 then
					vetor_materia_exibe=co_materia(f)
				else
					vetor_materia_exibe=vetor_materia_exibe&"#!#"&co_materia(f)
				end if
				co_materia_check=co_materia_check+1		
			end if				
		end if	
	NEXT

else
end if
programa_aula=vetor_materia_exibe
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
			verificar_medias=0	
			for al=0 to ubound(dados_aluno)
			dividendo=0
			
		
				medias_bimestres=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, dados_aluno(al), co_materia(i), caminho_nota, tb_nota, periodo)	
		
				'response.Write("="&medias_bimestres&"<BR>")			
				if medias_bimestres<>"&nbsp;" and medias_bimestres<>"" and not isnull(medias_bimestres) then	
					verificar_medias=medias_bimestres
					verificar_medias=verificar_medias*1
					valor_compara=replace(valor,".",",")
					valor_compara=valor_compara*1	
					if operacao="menor" then
						if verificar_medias<valor_compara then
							conta_aluno=conta_aluno+1
						else
							conta_aluno=conta_aluno				
						end if
					elseif operacao="maior" then
						
						if verificar_medias>=valor_compara then
							conta_aluno=conta_aluno+1
						else
							conta_aluno=conta_aluno				
						end if
					end if	
				else				
				end if		
			next	

				
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
	
	For f2=0 to ubound(co_materia)
	
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& co_materia(f2) &"'"
		RS.Open SQL, CON0
	
		mae= RS("IN_MAE")
		fil= RS("IN_FIL")
		in_co= RS("IN_CO")
		peso= RS("NU_Peso")
		
		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) then
		

			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("&vetor_aluno&") AND CO_Materia ='"& co_materia(f2)&"' And NU_Periodo="&periodo
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
				
		elseif (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) then
		
			vetor_mae_filhas=""
	
			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL2 = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_materia(f2) &"'"
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
			
			co_materia_mae_fil= split(vetor_mae_filhas,"#!#")
			media_mae_acumula=0			
			for f3=0 to ubound(co_materia_mae_fil)			
		
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL3 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& vetor_aluno &") AND CO_Materia ='"& co_materia_mae_fil(f3)&"' And NU_Periodo="&periodo
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
			media_mae=formatnumber(media_mae,1)
			vetor_quadro=vetor_quadro&"#!#"&media_mae	

		elseif (mae=TRUE and fil=FALSE and in_co=TRUE) then
		
			vetor_filhas=""
	
			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL2 = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_materia(f2) &"'"
			RS2.Open SQL2, CON0
				
			co_materia_fil_check=0 
			while not RS2.EOF
				co_mat_fil= RS2("CO_Materia")				
				if co_materia_fil_check=0 then
					vetor_filhas=co_mat_fil
				else
					vetor_filhas=vetor_filhas&"#!#"&co_mat_fil			
				end if
				co_materia_fil_check=co_materia_fil_check+1 									
			RS2.MOVENEXT
			wend				
			
			co_materia_fil= split(vetor_filhas,"#!#")
			media_mae_acumula=0	
			conta_nulo=0		
			for f3=0 to ubound(co_materia_fil)			
		
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL3 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& vetor_aluno &") AND CO_Materia ='"& co_materia_fil(f3)&"' And NU_Periodo="&periodo
				RS3.Open SQL3, CONn
				
				if RS3.EOF then
					sem_nota="s"
				else
'response.Write(media_mae_acumula)					
					media_turma=RS3("MediaDeVA_Media3")			
					if media_turma="" or isnull(media_turma) then
						media_filha_acumula=0
						conta_nulo=conta_nulo+1
					else
						media_turma=formatnumber(media_turma,1)
						media_filha_acumula=media_turma
					end if 
	
					media_mae_acumula=media_mae_acumula*1
					media_filha_acumula=media_filha_acumula*1
					media_mae_acumula=media_mae_acumula+media_filha_acumula		
				end if							
			next
			conta_nulo=conta_nulo*1
			co_materia_fil_check=co_materia_fil_check*1			
			if conta_nulo=co_materia_fil_check or sem_nota="s" or media_mae_acumula=0 then
				media_mae=""
			else
				media_mae=media_mae_acumula/co_materia_fil_check
				media_mae=media_mae*10
					decimo = media_mae - Int(media_mae)
					If decimo >= 0.5 Then
						nota_arredondada = Int(media_mae) + 1
						media_mae=nota_arredondada
					else
						nota_arredondada = Int(media_mae)
						media_mae=nota_arredondada											
					End If
				media_mae=media_mae/10			
				media_mae=formatnumber(media_mae,1)
			end if		
				if co_materia_check=1 then
					vetor_quadro=media_mae
				else
					vetor_quadro=vetor_quadro&"#!#"&media_mae
				end if
		end if	
	co_materia_check=co_materia_check+1			
	NEXT

calcula_medias=vetor_quadro&"#$#"	
elseif tipo_calculo="media_geral" then	
	
				Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& vetor_aluno &") And NU_Periodo="&periodo
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

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR		
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&session("ano_letivo")&"'"
		RS.Open SQL, CON		
		
		if RS.EOF then
			st_ano_letivo="L"	
		else		
			st_ano_letivo=RS("ST_Ano_Letivo")
		end if
	
	co_materia= split(vetor_materia,"#!#")	
	co_materia_check=0	

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT MAX(NU_Periodo) AS max_per FROM TB_Periodo"
		RS0.Open SQL0, CON0
		
	maior_periodo_tabela=RS0("max_per")
	vetor_periodo= split(periodo,"#!#")	
	maior_periodo_solicitado=vetor_periodo(ubound(vetor_periodo))

	
	vetor_alunos_turma = alunos_turma(SESSION("ano_letivo"),unidade,curso,co_etapa,turma,outro)	
	vetor_alunos= split(vetor_alunos_turma,"#$#")		
	for nm=0 to ubound(vetor_alunos)
		vetor_matriculas= split(vetor_alunos(nm),"#!#")		
		if nm=0 then
			matr_alunos_turma=vetor_matriculas(0)
		else
			matr_alunos_turma=matr_alunos_turma&","&vetor_matriculas(0)		
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
		
		if maior_periodo_solicitado>3 then
			calcula_media_anual="sim"		
			calcula_media_final="sim"	
		else
			calcula_media_anual="nao"		
			calcula_media_final="nao"			
		end if		
		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) then
			total_falta=0
			for f2a=1 to maior_periodo_tabela
				periodo_cons=f2a
				f2a=f2a*1
				maior_periodo_solicitado=maior_periodo_solicitado*1
				if f2a>maior_periodo_solicitado then	
					media=""
					media_soma=0
					'calcula_media_anual="nao"	
					falta=""
					media_turma=""
				else					
					media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, vetor_aluno, co_materia(f2), caminho_nota, tb_nota, periodo_cons)	
					if media="" or isnull(media) then
						media=""
						media_soma=0
						if periodo_cons<4 then
							calcula_media_anual="nao"	
							calcula_media_final="nao"								
						end if				
					else
						if media=0 then
							'media=""
							media_soma=0
							if periodo_cons<4 then
								calcula_media_anual="nao"	
								calcula_media_final="nao"								
							end if						
						else
							media_soma=media	
							media=formatnumber(media,1)										
						end if	
					end if
					
					if periodo_cons=1 then
						periodo_fal="f1"
					elseif periodo_cons=2 then	
						periodo_fal="f2"
					elseif periodo_cons=3 then
						periodo_fal="f3"
					end if
					
					if periodo_cons<4 then	
						falta=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, vetor_aluno, co_materia(f2), caminho_nota, tb_nota, periodo_fal)				
'						response.Write(falta&"-"&co_materia(f2)&"-"&vetor_aluno&"-"&periodo_fal&"<BR>")
					else
						falta=0
					end if
					if falta=0 then
						falta=""
					else
						falta=falta*1
						total_falta=total_falta*1
						total_falta=total_falta+falta
					end if				
	
					Set RS1 = Server.CreateObject("ADODB.Recordset")
					SQL1 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("&matr_alunos_turma&") AND CO_Materia ='"& co_materia(f2)&"' And NU_Periodo="&periodo_cons
					RS1.Open SQL1, CONn
				
					if RS1.EOF then
						media_turma=""
					else
						media_turma=RS1("MediaDeVA_Media3")
						if isnull(media_turma) or media_turma="" then
						
						else
							media_turma = arredonda(media_turma,"quarto_dez",1,0)
							media_turma=formatnumber(media_turma,1)				
						end if			
					end if	
				end if	
                 
				if periodo_cons=4 then
				
					qtd_periodos=maior_periodo_tabela
					divisor_anual=qtd_periodos-1
					
					if calcula_media_anual="sim" then
						media_anual=soma/divisor_anual	
						media_anual=media_anual*10			
						decimo = media_anual - Int(media_anual)
						If decimo >= 0.5 Then
							nota_arredondada = Int(media_anual) + 1
							media_anual=nota_arredondada
						else
							nota_arredondada = Int(media_anual)
							media_anual=nota_arredondada											
						End If		
						media_anual=media_anual/10						
						media_exibe = formatNumber(media_anual,1)												
					else
						media_exibe=""
					end if		
					
					if ((media="" or isnull(media)) and st_ano_letivo<>"B")	or media_anual="" or isnull(media_anual) then
						media_final=""
						total_final=""
					else
						total_final=(soma*1)+(media*2)
						if calcula_media_final="nao" then
							media_final=""
						else
							media_final=(total_final)/5
							media_final=media_final*10
							decimo = media_final - Int(media_final)
							If decimo >= 0.5 Then
								nota_arredondada = Int(media_final) + 1
								media_final=nota_arredondada
							else
								nota_arredondada = Int(media_final)
								media_final=nota_arredondada											
							End If		
							media_final=media_final/10												
							media_exibe = formatNumber(media_final,1)							
						end if													
						total_final=formatnumber(total_final,1)							
					end if
					soma=formatnumber(soma,1)	
							
					if total_falta=0 then
						total_falta=""	
					elseif calcula_media_anual="nao" and calcula_media_final="nao" then	
						total_falta=""						
					end if	
																						
					vetor_quadro=vetor_quadro&"#!#"&soma&"#!#"&media&"#!#"&total_final&"#!#"&media_exibe&"#!#"&vetor_med_turma&"#!#"&vetor_falta&"#!#"&total_falta
				else
					soma=soma+media_soma

					if co_materia_check=0 AND periodo_cons=1 then
						vetor_falta=falta
						vetor_med_turma=media_turma
						vetor_quadro=media
					elseif periodo_cons=1 then
						vetor_falta=falta
						vetor_med_turma=media_turma
						vetor_quadro=vetor_quadro&media					
					else
						vetor_falta=vetor_falta&"#!#"&falta
						vetor_med_turma=vetor_med_turma&"#!#"&media_turma
						vetor_quadro=vetor_quadro&"#!#"&media
					end if
				end if
			Next	
		vetor_quadro=vetor_quadro&"#$#"	
			
		elseif (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) then
		elseif (mae=TRUE and fil=TRUE and in_co=FALSE) then
	
		elseif (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then
			total_falta=0

			for f2a=1 to maior_periodo_tabela
				periodo_cons=f2a
				f2a=f2a*1
				maior_periodo_solicitado=maior_periodo_solicitado*1
				if f2a>maior_periodo_solicitado then	
					media=""
					media_soma=0
					'calcula_media_anual="nao"	
					falta=""
					media_turma=""
				else					
					media=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, vetor_aluno, co_materia(f2), caminho_nota, tb_nota, periodo_cons)	
					if media="" or isnull(media) then
						media=""
						media_soma=0
						if periodo_cons<4 then
							calcula_media_anual="nao"	
							calcula_media_final="nao"								
						end if			
					else
						if media=0 then
							media=""
							media_soma=0
							if periodo_cons<4 then
								calcula_media_anual="nao"	
								calcula_media_final="nao"								
							end if						
						else
							media_soma=media	
							media=formatnumber(media,1)										
						end if	
					end if
					
					if periodo_cons=1 then
						periodo_fal="f1"
					elseif periodo_cons=2 then	
						periodo_fal="f2"
					elseif periodo_cons=3 then
						periodo_fal="f3"
					end if
					
					if periodo_cons<4 then	
						falta=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, vetor_aluno, co_materia(f2), caminho_nota, tb_nota, periodo_fal)				
						'response.Write(falta&"-"&co_materia(f2)&"-"&vetor_aluno&"-"&periodo_fal&"<BR>")
					else
						falta=0
					end if
					if falta=0 then
						falta=""
					else
						falta=falta*1
						total_falta=total_falta*1
						total_falta=total_falta+falta
					end if				
	
					Set RS1 = Server.CreateObject("ADODB.Recordset")
					SQL1 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("&matr_alunos_turma&") AND CO_Materia_Principal ='"& co_materia(f2)&"' And NU_Periodo="&periodo_cons
					RS1.Open SQL1, CONn
					
						if RS1.EOF then
							media_turma=""
						else
							media_turma=RS1("MediaDeVA_Media3")
							if isnull(media_turma) or media_turma="" then
							
							else
								media_turma = arredonda(media_turma,"quarto_dez",1,0)
								media_turma=formatnumber(media_turma,1)				
							end if			
						end if	

				end if	

	
				if periodo_cons=4 then
				
					qtd_periodos=maior_periodo_tabela
					divisor_anual=qtd_periodos-1
'response.Write(soma&"/"&divisor_anual	&"<BR>")					
					if calcula_media_anual="sim" then
						media_anual=soma/divisor_anual	
						media_anual=media_anual*10			
						decimo = media_anual - Int(media_anual)
						If decimo >= 0.5 Then
							nota_arredondada = Int(media_anual) + 1
							media_anual=nota_arredondada
						else
							nota_arredondada = Int(media_anual)
							media_anual=nota_arredondada											
						End If		
						media_anual=media_anual/10						
						media_exibe = formatNumber(media_anual,1)													
					else
						media_exibe=""
					end if		
					
					if ((media="" or isnull(media)) and st_ano_letivo<>"B")	or media_anual="" or isnull(media_anual) then
						media_final=""
						total_final=""
					else
						total_final=(soma*1)+(media*2)
						if calcula_media_final="nao" then
							media_final=""
						else
							media_final=(total_final)/5
							media_final=media_final*10
							decimo = media_final - Int(media_final)
							If decimo >= 0.5 Then
								nota_arredondada = Int(media_final) + 1
								media_final=nota_arredondada
							else
								nota_arredondada = Int(media_final)
								media_final=nota_arredondada											
							End If		
							media_final=media_final/10												
							
							media_exibe = formatNumber(media_final,1)												
						end if
						total_final=formatnumber(total_final,1)	
					end if
					soma=formatnumber(soma,1)	
					
					if total_falta=0 then
						total_falta=""	
					elseif calcula_media_anual="nao" and calcula_media_final="nao" then	
						total_falta=""								
					end if						
					'response.Write(	media_anual&"-"&media_final)	
					'response.End()							
					vetor_quadro=vetor_quadro&"#!#"&soma&"#!#"&media&"#!#"&total_final&"#!#"&media_exibe&"#!#"&vetor_med_turma&"#!#"&vetor_falta&"#!#"&total_falta
				else
					soma=soma+media_soma

					if co_materia_check=0 AND periodo_cons=1 then
						vetor_falta=falta
						vetor_med_turma=media_turma
						vetor_quadro=media
					elseif periodo_cons=1 then
						vetor_falta=falta
						vetor_med_turma=media_turma
						vetor_quadro=vetor_quadro&media					
					else
						vetor_falta=vetor_falta&"#!#"&falta
						vetor_med_turma=vetor_med_turma&"#!#"&media_turma
						vetor_quadro=vetor_quadro&"#!#"&media
					end if
				end if
			Next	
		vetor_quadro=vetor_quadro&"#$#"	
				

		elseif (mae=TRUE and fil=FALSE and in_co=TRUE) then
	
		end if		
	co_materia_check=co_materia_check+1			
	
	NEXT	
calcula_medias=vetor_quadro	

'response.End()
end if
end function








'calcula as médias anuais e finais destes respectivos mapas
Function Calc_Med_An_Fin(unidade, curso, co_etapa, turma, vetor_aluno, vetor_materia, caminho_nota, tb_nota, qtd_periodos, periodo_m2, periodo_m3,tipo_calculo, outro)

		if periodo_m2>0 then
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
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CONn = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONn.Open ABRIRn	
		
		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR		
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&session("ano_letivo")&"'"
		RS.Open SQL, CON		
		
	if RS.EOF then
		st_ano_letivo="L"	
	else		
		st_ano_letivo=RS("ST_Ano_Letivo")
	end if		
		
		
	alunos= split(vetor_aluno,"#$#")			
	co_materia= split(vetor_materia,"#!#")	
	co_materia_check=1	
	co_matricula= vetor_aluno
	quantidade_alunos=0
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
			contando_medias=0
				'response.Write(mae&"-"&fil&"-"&in_co&"-"&peso&"<BR>")		
			if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) then
				for periodo=1 to qtd_periodos
					Set RSn = Server.CreateObject("ADODB.Recordset")
					SQLn = "SELECT * FROM "&tb_nota&" where CO_Matricula ="& dados_aluno(0) &" AND CO_Materia ='"& co_materia(c) &"' AND CO_Materia_Principal ='"& co_materia(c) &"' AND NU_Periodo="&periodo				
					RSn.Open SQLn, CONn

						qtd_periodos=qtd_periodos*1
						periodo=periodo*1
						periodo_m2=periodo_m2*1
'resultado 3 não é usado nessa escola						
						periodo_m3=periodo_m3*1		
					if RSn.EOF then
						media_acumulada=media_acumulada				
						peso_periodo_acumulado=peso_periodo_acumulado

						if periodo=periodo_m2 then
							rec_lancado="nao"
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
						else		
							if md="" or isnull(md) then
								media_acumulada=media_acumulada				
								peso_periodo_acumulado=peso_periodo_acumulado						
							else
								media_acumulada=media_acumulada+(md*peso_periodo)
								peso_periodo_acumulado=peso_periodo_acumulado+peso_periodo
								contando_medias=contando_medias+1						
							end if
						end if						
					end if
				Next

				if peso_periodo_acumulado=0 then
					peso_periodo_acumulado=1
				end if	

				if contando_medias>=medias_necessarias or (st_ano_letivo="B" and tipo_calculo="ata") then
					media_anual=media_acumulada/peso_periodo_acumulado	
					ma_para_calc_final = media_anual		
					media_anual=media_anual*10			
					decimo = media_anual - Int(media_anual)
					If decimo >= 0.5 Then
						nota_arredondada = Int(media_anual) + 1
						media_anual=nota_arredondada
					else
						nota_arredondada = Int(media_anual)
						media_anual=nota_arredondada											
					End If		
					media_anual=media_anual/10						
					media_anual = formatNumber(media_anual,1)			

					
					if tipo_calculo="anual" then
						resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","anual")
						media_anual = formatNumber(media_anual,1)
						resultado_materia=media_anual&"#!#"&resultado		
					elseif tipo_calculo="ata" then	
						if rec_lancado="nao" or media_rec="" or isnull(media_rec) then
							resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","ata")
							resultado_materia=resultado
						else
							resultado=regra_aprovacao(curso,co_etapa,ma_para_calc_final,media_rec,"&nbsp;","&nbsp;","&nbsp;","ata")					
							resultado_materia=resultado
						end if												
					elseif rec_lancado="nao" or media_rec="" or isnull(media_rec) then
						resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","final")
						resultado_materia=resultado
					else
						resultado=regra_aprovacao(curso,co_etapa,ma_para_calc_final,media_rec,"&nbsp;","&nbsp;","&nbsp;","final")					
						resultado_materia=resultado
						
'if dados_aluno(0)  = 20090036 then
'RESPONSE.Write(media_rec&" + "&media_acumulada&"/"&peso_periodo_acumulado&"<BR>")
'if co_materia(c) = "CIENC" then
'RESPONSE.Write(media_anual)
'	RESPONSE.End()
'end if	
'end if						
					end if	
				else
						resultado_materia="&nbsp;#!#&nbsp;"
				end if	
										
			elseif (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) then
			
			elseif (mae=TRUE and fil=TRUE and in_co=FALSE) then			
			
			elseif (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then			
				for periodo=1 to qtd_periodos
					md=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, dados_aluno(0), co_materia(c), caminho_nota, tb_nota, periodo)	
					
						if md="&nbsp;" or md="" or isnull(md) then
							media_acumulada=media_acumulada				
							peso_periodo_acumulado=peso_periodo_acumulado

							if periodo=periodo_m2 then
								rec_lancado="nao"
							end if
						else
							Set RSPESO = Server.CreateObject("ADODB.Recordset")
							SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&periodo
							RSPESO.Open SQLPESO, CON0					
							peso_periodo=RSPESO("NU_Peso")

						if periodo=periodo_m2 then
							media_rec=md
							rec_lancado="sim"
						else		
							if md="" or isnull(md) then
								media_acumulada=media_acumulada				
								peso_periodo_acumulado=peso_periodo_acumulado						
							else
								media_acumulada=media_acumulada+(md*peso_periodo)
								peso_periodo_acumulado=peso_periodo_acumulado+peso_periodo
								contando_medias=contando_medias+1						
							end if
						end if						
					end if									
				Next
				if peso_periodo_acumulado=0 then
					peso_periodo_acumulado=1
				end if	

				if contando_medias>=medias_necessarias or (st_ano_letivo="B" and tipo_calculo="ata") then
					media_anual=media_acumulada/peso_periodo_acumulado	
					ma_para_calc_final = media_anual						
					media_anual=media_anual*10			
					decimo = media_anual - Int(media_anual)
					If decimo >= 0.5 Then
						nota_arredondada = Int(media_anual) + 1
						media_anual=nota_arredondada
					else
						nota_arredondada = Int(media_anual)
						media_anual=nota_arredondada											
					End If		
					media_anual=media_anual/10						
					media_anual = formatNumber(media_anual,1)			


					
					if tipo_calculo="anual" then
						resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","anual")
						media_anual = formatNumber(media_anual,1)
						resultado_materia=media_anual&"#!#"&resultado	
					elseif tipo_calculo="ata" then	
						if rec_lancado="nao" or media_rec="" or isnull(media_rec) then
							resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","ata")
							resultado_materia=resultado
						else
							resultado=regra_aprovacao(curso,co_etapa,ma_para_calc_final,media_rec,"&nbsp;","&nbsp;","&nbsp;","ata")					
							resultado_materia=resultado
						end if											
					elseif rec_lancado="nao" or media_rec="" or isnull(media_rec) then
						resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","final")
						resultado_materia=resultado
					else
						resultado=regra_aprovacao(curso,co_etapa,ma_para_calc_final,media_rec,"&nbsp;","&nbsp;","&nbsp;","final")					
						resultado_materia=resultado
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
'response.Write(dados_aluno(1)&"-"&tipo_calculo&"-"&media_anual&"-"&rec_lancado&"-"&md&"-"&media_rec&"<BR>")

	next
'response.Write(resultado_turma)
'response.End()
Calc_Med_An_Fin=resultado_turma		
END FUNCTION

Function regra_aprovacao (curso,etapa,m1_aluno,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,tipo_calculo)

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&session("ano_letivo")&"'"
	RS.Open SQL, CON		
		
	if RS.EOF then
		st_ano_letivo="L"	
	else		
		st_ano_letivo=RS("ST_Ano_Letivo")
	end if		
	
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
		m2_aluno=m1_aluno	
	else	
		if tipo_calculo="ata" then
			if st_ano_letivo="B" then
				if nota_aux_m2_1="&nbsp;" then
					m2_aluno=m1_aluno	
				else
					tipo_calculo="final"
				end if
			else
				tipo_calculo="final"			
			end if	
		end if		
		if tipo_calculo="final" then
			if nota_aux_m2_1="&nbsp;" then
				m2_aluno="&nbsp;"
				resultado="&nbsp;"	
			else								
				m1_aluno_peso=m1_aluno*peso_m2_m1
				nota_aux_m2_1_peso=nota_aux_m2_1*peso_m2_m2
				m2_aluno=(m1_aluno_peso+nota_aux_m2_1_peso)/(peso_m2_m1+peso_m2_m2)
					m2_aluno=m2_aluno*10					
				decimo = m2_aluno - Int(m2_aluno)
				If decimo >= 0.5 Then
					nota_arredondada = Int(m2_aluno) + 1
					m2_aluno=nota_arredondada
				else
					nota_arredondada = Int(m2_aluno)
					m2_aluno=nota_arredondada											
				End If	
					m2_aluno=m2_aluno/10					
				m2_aluno = formatNumber(m2_aluno,1)
				m2_aluno=m2_aluno*1
				valor_m2=valor_m2*1	
	
				if m2_aluno >= m2_maior_igual then
					resultado=res2_3
				elseif m2_aluno >= m2_menor then
					resultado=res2_2
				else
					resultado=res2_1	
				end if
			end if
		end if

	end if
	if tipo_calculo="anual" then
		regra_aprovacao=resultado
	else
		if m2_aluno<>"&nbsp;" then
			m2_aluno = formatNumber(m2_aluno,1)
		end if
		regra_aprovacao=m2_aluno&"#!#"&resultado	
	end if
	
	'Session("M2")=m2_aluno
	'Session("M3")=m3_aluno
end function	

Function apura_resultado_aluno (curso,etapa,vetor_medias)

	apura_resultado_aluno=apura_resultado_geral_aluno (curso,etapa,vetor_medias, "F")

end function	

Function apura_resultado_geral_aluno (curso,etapa,vetor_medias, tipo)

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set RSra = Server.CreateObject("ADODB.Recordset")
	SQLra = "SELECT * FROM TB_Regras_Aprovacao where CO_Curso = '"&curso&"' and CO_Etapa = '"&etapa&"'"
	RSra.Open SQLra, CON0	
			

	if tipo="F" then
		valor_apr=RSra("NU_Valor_Apr")
		valor_dep=RSra("NU_Valor_Dep")
		qtd_max_dep=RSra("NU_Qt_Dis_Dep")	
		res_apr=RSra("NO_Expr_Maior_Igual_VL_Abr")
		res_dep=RSra("NO_Expr_Int_M1_F")
		res_rep=RSra("NO_Expr_Cond_Falso_Abr")
	elseif tipo="A" then
		valor_apr=RSra("NU_Valor_M1")
		valor_dep=0
		qtd_max_dep=RSra("NU_Qt_Dis_Dep")
		res_apr=RSra("NO_Expr_Maior_Igual_VL_Abr")
		res_dep=RSra("NO_Expr_Int_M1_F")
		res_rep=""	
	end if	
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

	if md_aluno="" or isnull(md_aluno) or md_aluno="&nbsp;" or md_aluno=" "then
		libera_resultado="n"
	else
		md_aluno=md_aluno*1
		valor_apr=valor_apr*1
		valor_dep=valor_dep*1	
		if result_temp<>"rep" then
			if md_aluno >= valor_apr and result_temp<>"dep" then
				result_temp="apr"
			elseif md_aluno >= valor_dep then
				result_temp="dep"
				qtd_dep=qtd_dep+1
			else
				result_temp="rep"			
			end if
		end if	
	end if
'response.Write(valor_apr&"<BR>")	
'response.Write(md_aluno&"<BR>")
'response.Write(result_temp)
Next
'response.Write(valor_apr&"<BR>")
'response.Write(valor_dep&"<BR>")
'response.Write(result_temp)
'response.End()
if 	libera_resultado="s" then
	qtd_dep=qtd_dep*1
	if result_temp="apr" and qtd_dep=0 then
		apura_resultado_geral_aluno=res_apr
	elseif result_temp="rep" then
		apura_resultado_geral_aluno=res_rep
	elseif result_temp="dep" then	
		apura_resultado_geral_aluno=res_dep
		qtd_max_dep=qtd_max_dep*1
'response.Write(qtd_dep&"-"&qtd_max_dep)
'response.End()		
'		if qtd_dep>qtd_max_dep then
'			apura_resultado_geral_aluno=res_rep	
'		else	
			apura_resultado_geral_aluno=res_dep	
'		end if
	end if	
else
	apura_resultado_geral_aluno="&nbsp;"		
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
		strReplacement = replace(strReplacement,"Ö","&Ouml;")		
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
		strReplacement = replace(strReplacement,"ö","&ouml;")		
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
		strReplacement = replace(strReplacement,"%D6","Ö")		
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
		strReplacement = replace(strReplacement,"%F6","ö")		
		strReplacement = replace(strReplacement,"%FA","ú")
		strReplacement = replace(strReplacement,"%FC","ü")
	end if
replace_latin_char=strReplacement
end function

'===========================================================================================================================================
'serve também para (mae=FALSE and fil=FALSE and in_co=TRUE) para o Mapa de Resultados por Disciplinas		
Function Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, periodo)
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CONn = Server.CreateObject("ADODB.Connection") 
	ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONn.Open ABRIRn	

if periodo="f1" then
	periodo_consulta=1
elseif periodo="f2" then	
	periodo_consulta=2
elseif periodo="f3" then
	periodo_consulta=3
else
	periodo_consulta=periodo
end if

if periodo="f1" or periodo="f2" or periodo="f3" then

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT NU_Faltas FROM "&tb_nota&" where CO_Matricula ="& cod_aluno &" AND CO_Materia_Principal ='"& codigo_materia &"' AND CO_Materia ='"& codigo_materia &"' And NU_Periodo="&periodo_consulta
		RS1.Open SQL1, CONn
		
		if RS1.EOF then
			va_m3=0
		else
			va_m3=RS1("NU_Faltas")
			if isnull(va_m3) or va_m3="" then				
				va_m3=0			
			end if	
		end if		
else
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT VA_Media3 FROM "&tb_nota&" where CO_Matricula ="& cod_aluno &" AND CO_Materia_Principal ='"& codigo_materia &"' AND CO_Materia ='"& codigo_materia &"' And NU_Periodo="&periodo_consulta
		RS1.Open SQL1, CONn
		
		if RS1.EOF then
			va_m3=""
		else
			va_m3=RS1("VA_Media3")				
		end if		
end if

Calcula_Media_T_F_F_N=va_m3

end function














'===========================================================================================================================================
Function Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, periodo)

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CONn = Server.CreateObject("ADODB.Connection") 
	ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONn.Open ABRIRn	

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
			
			if periodo="f1" then
				periodo_consulta=1
			elseif periodo="f2" then	
				periodo_consulta=2
			elseif periodo="f3" then
				periodo_consulta=3
			else
				periodo_consulta=periodo
			end if
			
			if periodo="f1" or periodo="f2" or periodo="f3" then
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL_N = "Select NU_Faltas from "& tb_nota &" WHERE CO_Matricula = "& cod_aluno &" AND CO_Materia = '"& co_mat_fil &"' AND CO_Materia_Principal = '"& codigo_materia &"' AND NU_Periodo="&periodo_consulta
'				response.Write(SQL_N&"<BR>")
				Set RS3 = CONn.Execute(SQL_N)						
		
	
				if RS3.EOF then
					va_m3_temp=0
				else					
					va_m3_temp=RS3("NU_Faltas")
				end if
	
				if isnull(va_m3_temp) or va_m3_temp="&nbsp;"  or va_m3_temp="" then
					va_m3_temp=0
				end if	
'				response.Write("F"&periodo_consulta&"="&va_m3_temp&"<BR>")					
				va_m3_acumula=va_m3_acumula*1
				va_m3_temp=va_m3_temp*1
				va_m3_acumula=va_m3_acumula+va_m3_temp								
'				response.Write("va_m3_acumula="&va_m3_acumula&"+"&va_m3_temp&"<BR>")				

			else						
				Set RSp2 = Server.CreateObject("ADODB.Recordset")
				SQLp2 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia = '"& co_mat_fil &"' order by NU_Ordem_Boletim"
	
				RSp2.Open SQLp2, CON0	
										
				nu_peso_fil=RSp2("NU_Peso")	
							
				peso_acumula=peso_acumula+nu_peso_fil
											
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL_N = "Select VA_Media3 from "& tb_nota &" WHERE CO_Matricula = "& cod_aluno &" AND CO_Materia = '"& co_mat_fil &"' AND CO_Materia_Principal = '"& codigo_materia &"' AND NU_Periodo="&periodo
				Set RS3 = CONn.Execute(SQL_N)						
		
	
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
			end if																									
		RS1a.MOVENEXT
		wend

		if periodo="f1" or periodo="f2" or periodo="f3" then			
			va_m3=va_m3_acumula
		else	
			if sem_nota="s" then
				va_m3=""
			else	
				va_m3=va_m3_acumula/peso_acumula
				va_m3 = arredonda(va_m3,"quarto_dez",1,0)								
			end if
		end if
	end if	
Calcula_Media_T_F_T_N=va_m3


end function				
%>