<!--#include file="../../../global/funcoes_diversas.asp" -->
<%
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
		
		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) then
			
			if co_materia_check=1 then
				vetor_materia_exibe=co_materia(f)
			else
				vetor_materia_exibe=vetor_materia_exibe&"#!#"&co_materia(f)
			end if
			co_materia_check=co_materia_check+1			
		
		elseif (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) or (mae=TRUE and fil=TRUE and in_co=FALSE) then
	
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
'response.Write(SQL)
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

				medias_bimestres=Calc_Med_T_F_T_N(unidade, curso, co_etapa, turma, dados_aluno(al), co_materia(fb), caminho_nota, tb_nota, 5, 5, 0,"nulo", "nulo")	
		
				medias=Split(medias_bimestres,"#!#")			
				
				periodo=periodo*1
							
				if periodo = 1 then
					dividendo=medias(0)
					dividendo_asterisco=medias(4)
				elseif periodo = 2 then	
					dividendo=medias(1)
					dividendo_asterisco=medias(5)
				elseif periodo = 3 then	
					dividendo=medias(6)
					dividendo_asterisco=medias(10)
				elseif periodo = 4 then	
					dividendo=medias(7)
					dividendo_asterisco=medias(11)
				elseif periodo = 5 then	
					dividendo=medias(12)
					dividendo_asterisco=0
				end if					
								
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

Function calcula_medias(unidade, curso, co_etapa, turma, periodo, vetor_aluno, vetor_materia, caminho_nota, tb_nota, nome_nota, tipo_calculo)


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
	
	co_materia= split(vetor_materia,"#!#")	
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
			SQL1 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& vetor_aluno &") AND CO_Materia ='"& co_materia(fb)&"' And NU_Periodo="&periodo
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
			media_mae=arredonda(media_mae,"mat_dez",1,0)
			vetor_quadro=vetor_quadro&"#!#"&media_mae	
			
		elseif (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then

			dados_aluno=Split(vetor_aluno,",")
			soma_medias=0
			media_somada=0
			qtd_alunos=ubound(dados_aluno)+1
		
			if qtd_alunos=0 then
				qtd_alunos=1
			end if
			
			for al=0 to ubound(dados_aluno)

				medias_bimestres=Calc_Med_T_F_T_N(unidade, curso, co_etapa, turma, dados_aluno(al), co_materia(fb), caminho_nota, tb_nota, 5, 5, 0,"nulo", "nulo")	
		'	response.Write(medias_bimestres&", "&co_materia(fb))				
				medias=Split(medias_bimestres,"#!#")			
				
				periodo=periodo*1
							
				if periodo = 1 then
					dividendo=medias(0)
					dividendo_asterisco=medias(4)
				elseif periodo = 2 then	
					dividendo=medias(1)
					dividendo_asterisco=medias(5)
				elseif periodo = 3 then	
					dividendo=medias(6)
					dividendo_asterisco=medias(10)
				elseif periodo = 4 then	
					dividendo=medias(7)
					dividendo_asterisco=medias(11)
				elseif periodo = 5 then	
					dividendo=medias(12)
					dividendo_asterisco=0
				end if					
								
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
				media_mae=arredonda(media_mae,"mat_dez",1,0)			
			end if
			if co_materia_check=1 then
				vetor_quadro=media_mae
			else
				vetor_quadro=vetor_quadro&"#!#"&media_mae_acumula	
			end if	
			
		elseif (mae=FALSE and fil=FALSE and in_co=TRUE) then

			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT Avg("&tb_nota&"."&nome_nota&")AS MediaDeVA_Media3 FROM "&tb_nota&" where CO_Matricula in("& vetor_aluno &") AND CO_Materia ='"& co_materia(fb)&"' And NU_Periodo="&periodo
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

cod_aluno=vetor_aluno

'RESPONSE.Write(conta_materia_check&"<br>")
	co_materia= split(vetor_materia,"#!#")	
	conta_materia_check=0	

	vetor_periodo= split(periodo,"#!#")	

	For fb=0 to ubound(co_materia)
	soma=0	
'RESPONSE.Write("<br>"&co_materia(fb)&"<br>")	
'RESPONSE.Write(conta_materia_check&"<br>")
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& co_materia(fb) &"'"

		RS.Open SQL, CON0
	
		mae= RS("IN_MAE")
		fil= RS("IN_FIL")
		in_co= RS("IN_CO")
		peso= RS("NU_Peso")
		
		calcula_media_anual="sim"	
			
'or (mae=FALSE and fil=FALSE and in_co=TRUE) s� serve para o Mapa de Resultados por Disciplinas		
		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) or (mae=FALSE and fil=FALSE and in_co=TRUE)  then
	
			medias_bimestres=Calc_Med_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, co_materia(fb), caminho_nota, tb_nota, 5, 5, 0,"nulo", "nulo")		
'rESPONSE.Write("=="&medias_bimestres&"==<br>")
			medias_split=medias_bimestres
			medias=Split(medias_split,"#!#")
			
			if ubound(vetor_periodo)=0 then		
				dividendo1=medias(0)			
				dividendo2="&nbsp;"
				dividendo3="&nbsp;"
				dividendo4="&nbsp;"
				dividendo5="&nbsp;"
				f1=medias(13)			
				f2="&nbsp;"
				f3="&nbsp;"
				f4="&nbsp;"			
			elseif ubound(vetor_periodo)=1 then			
				dividendo1=medias(0)			
				dividendo2=medias(1)
				dividendo3="&nbsp;"
				dividendo4="&nbsp;"
				dividendo5="&nbsp;"
				f1=medias(13)			
				f2=medias(14)
				f3="&nbsp;"
				f4="&nbsp;"					
			elseif ubound(vetor_periodo)=2 then		
				dividendo1=medias(0)			
				dividendo2=medias(1)
				dividendo3=medias(6)
				dividendo4="&nbsp;"
				dividendo5="&nbsp;"		
				f1=medias(13)			
				f2=medias(14)
				f3=medias(15)
				f4="&nbsp;"						
			elseif ubound(vetor_periodo)=3 then	
				dividendo1=medias(0)			
				dividendo2=medias(1)
				dividendo3=medias(6)
				dividendo4=medias(7)
				dividendo5="&nbsp;"		
				f1=medias(13)			
				f2=medias(14)
				f3=medias(15)
				f4=medias(16)								
			elseif ubound(vetor_periodo)=4 then		
				dividendo1=medias(0)			
				dividendo2=medias(1)
				dividendo3=medias(6)
				dividendo4=medias(7)
				dividendo5=medias(12)
				f1=medias(13)			
				f2=medias(14)
				f3=medias(15)
				f4=medias(16)															
			end if
			
			
			if isnull(dividendo1) or dividendo1="" then
				dividendo1="&nbsp;"
			end if

			if isnull(dividendo2) or dividendo2="" then
				dividendo2="&nbsp;"
			end if

			if isnull(dividendo3) or dividendo3="" then
				dividendo3="&nbsp;"
			end if		

			if isnull(dividendo4) or dividendo4="" then
				dividendo4="&nbsp;"
			end if
			
			
			if isnull(dividendo5) or dividendo5="" then
				dividendo5="&nbsp;"
			end if	
			
			if isnull(f1) or f1="" then
				f1="&nbsp;"
			end if	
			
			if isnull(f2) or f2="" then
				f2="&nbsp;"
			end if	
				
			if isnull(f3) or f3="" then
				f3="&nbsp;"
			end if	
		
			if isnull(f4) or f4="" then
				f4="&nbsp;"
			end if													
					
			
			if session("veio_de")="EFI" then
				medias_lancadas=dividendo1&"#!#"&dividendo2&"#!#"&dividendo3&"#!#"&dividendo4
				if dividendo1<>"&nbsp;" and dividendo2<>"&nbsp;" and dividendo3<>"&nbsp;" and dividendo4<>"&nbsp;" then
					media_res=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_aluno, co_materia(fb), caminho_nota, tb_nota, 5, 5, 0, "boletim", 0)			
					resultados=medias_lancadas&"#!#"&media_res&"#!#"&f1&"#!#"&f2&"#!#"&f3&"#!#"&f4					
				else
					resultados=medias_lancadas&"#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#"&f1&"#!#"&f2&"#!#"&f3&"#!#"&f4									
				end if						
			else
				if dividendo1<>"&nbsp;" and dividendo2<>"&nbsp;" and dividendo3<>"&nbsp;" and dividendo4<>"&nbsp;" then
					media_res=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_aluno, co_materia(fb), caminho_nota, tb_nota, 5, 5, 0, "boletim", 0)			
					resultados=medias_bimestres&"#!#"&media_res							
				else
					resultados=medias_bimestres&"#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"							
				end if					
				
			end if	
			
														
				
			if conta_materia_check=0 then
				vetor_quadro=resultados
			else	
				vetor_quadro=vetor_quadro&"#$#"&resultados
			end if				

		elseif (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) then
			
		elseif (mae=TRUE and fil=TRUE and in_co=FALSE) then
			
		elseif (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then

			medias_bimestres=Calc_Med_T_F_T_N(unidade, curso, co_etapa, turma, cod_aluno, co_materia(fb), caminho_nota, tb_nota, 5, 5, 0,"nulo", "nulo")			
			
			medias=Split(medias_bimestres,"#!#")
			
			if ubound(vetor_periodo)=0 then		
				dividendo1=medias(0)			
				dividendo2="&nbsp;"
				dividendo3="&nbsp;"
				dividendo4="&nbsp;"
				dividendo5="&nbsp;"
				f1=medias(13)			
				f2="&nbsp;"
				f3="&nbsp;"
				f4="&nbsp;"			
			elseif ubound(vetor_periodo)=1 then			
				dividendo1=medias(0)			
				dividendo2=medias(1)
				dividendo3="&nbsp;"
				dividendo4="&nbsp;"
				dividendo5="&nbsp;"
				f1=medias(13)			
				f2=medias(14)
				f3="&nbsp;"
				f4="&nbsp;"					
			elseif ubound(vetor_periodo)=2 then		
				dividendo1=medias(0)			
				dividendo2=medias(1)
				dividendo3=medias(6)
				dividendo4="&nbsp;"
				dividendo5="&nbsp;"		
				f1=medias(13)			
				f2=medias(14)
				f3=medias(15)
				f4="&nbsp;"						
			elseif ubound(vetor_periodo)=3 then	
				dividendo1=medias(0)			
				dividendo2=medias(1)
				dividendo3=medias(6)
				dividendo4=medias(7)
				dividendo5="&nbsp;"		
				f1=medias(13)			
				f2=medias(14)
				f3=medias(15)
				f4=medias(16)								
			elseif ubound(vetor_periodo)=4 then		
				dividendo1=medias(0)			
				dividendo2=medias(1)
				dividendo3=medias(6)
				dividendo4=medias(7)
				dividendo5=medias(12)
				f1=medias(13)			
				f2=medias(14)
				f3=medias(15)
				f4=medias(16)															
			end if
			
			
			if isnull(dividendo1) or dividendo1="" then
				dividendo1="&nbsp;"
			end if

			if isnull(dividendo2) or dividendo2="" then
				dividendo2="&nbsp;"
			end if

			if isnull(dividendo3) or dividendo3="" then
				dividendo3="&nbsp;"
			end if		

			if isnull(dividendo4) or dividendo4="" then
				dividendo4="&nbsp;"
			end if
			
			
			if isnull(dividendo5) or dividendo5="" then
				dividendo5="&nbsp;"
			end if	
			
			if isnull(f1) or f1="" then
				f1="&nbsp;"
			end if	
			
			if isnull(f2) or f2="" then
				f2="&nbsp;"
			end if	
				
			if isnull(f3) or f3="" then
				f3="&nbsp;"
			end if	
		
			if isnull(f4) or f4="" then
				f4="&nbsp;"
			end if													
					
			
			if session("veio_de")="EFI" then
				medias_lancadas=dividendo1&"#!#"&dividendo2&"#!#"&dividendo3&"#!#"&dividendo4
				if dividendo1<>"&nbsp;" and dividendo2<>"&nbsp;" and dividendo3<>"&nbsp;" and dividendo4<>"&nbsp;" then
					media_res=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_aluno, co_materia(fb), caminho_nota, tb_nota, 5, 5, 0, "boletim", 0)			
					resultados=medias_lancadas&"#!#"&media_res&"#!#"&f1&"#!#"&f2&"#!#"&f3&"#!#"&f4					
				else
					resultados=medias_lancadas&"#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#"&f1&"#!#"&f2&"#!#"&f3&"#!#"&f4									
				end if						
			else
				if dividendo1<>"&nbsp;" and dividendo2<>"&nbsp;" and dividendo3<>"&nbsp;" and dividendo4<>"&nbsp;" then
					media_res=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_aluno, co_materia(fb), caminho_nota, tb_nota, 5, 5, 0, "boletim", 0)			
					resultados=medias_bimestres&"#!#"&media_res							
				else
					resultados=medias_bimestres&"#!#&nbsp;#!#&nbsp;#!#&nbsp;#!#&nbsp;"							
				end if					
				
			end if	
																					
					
			if conta_materia_check=0 then
				vetor_quadro=resultados
			else	
				vetor_quadro=vetor_quadro&"#$#"&resultados
			end if				
	
		end if		
		conta_materia_check=conta_materia_check+1			
	NEXT	
calcula_medias=vetor_quadro	
end if
end function








'===========================================================================================================================================
'serve tamb�m para (mae=FALSE and fil=FALSE and in_co=TRUE) s� serve para o Mapa de Resultados por Disciplinas		
Function Calc_Med_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, co_materia, caminho_nota, tb_nota, qtd_periodos, periodo_m2, periodo_m3,tipo_calculo, outro)
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
'	res_rec=RSra("NU_Valor_M2")
'	res_rep=RSra("NU_Valor_M3")	
'
'	if periodo_m2>0 then
'		retira_periodo_m2=1
'	else
'		retira_periodo_m2=0			
'	end if
'	
'	if periodo_m3>0 then
'		retira_periodo_m3=1
'	else
'		retira_periodo_m3=0			
'	end if
'					
'	medias_necessarias=qtd_periodos-retira_periodo_m2-retira_periodo_m3
		
	dividendo1=0
	divisor1=0
	dividendo2=0
	divisor2=0			
	dividendorec12=0	
	divisorrec12=0			
	dividendo3=0	
	divisor3=0											
	dividendo4=0
	divisor4=0			
	dividendorec34=0
	divisorrec34=0					
	dividendo5=0
	divisor5=0			


	for fba=1 to qtd_periodos
		periodo_cons=fba
		nota="&nbsp;"
		falta="&nbsp;"
		rec="&nbsp;"					
		media="&nbsp;"				

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM "&tb_nota&" where CO_Matricula ="& cod_aluno &" AND CO_Materia ='"& co_materia &"' And NU_Periodo="&periodo_cons

		RS1.Open SQL1, CONn

		if RS1.EOF then
			if periodo_cons=1 then
				f1="&nbsp;"
				va_m31="&nbsp;"
			elseif periodo_cons=2 then
				f2="&nbsp;"
				va_m32="&nbsp;"
				va_rec_sem12="&nbsp;"
			elseif periodo_cons=3 then
				f3="&nbsp;"
				va_m33="&nbsp;"
			elseif periodo_cons=4 then
				f4="&nbsp;"
				va_m34="&nbsp;"
				va_rec_sem34="&nbsp;"							
			elseif periodo_cons=5 then
				va_m35="&nbsp;"						
			end if									
			if periodo_cons<>5 then
				calcula_media_anual="nao"							
			end if
		else
			if periodo_cons=1 then
				f1=RS1("NU_Faltas")
				va_m31=RS1("VA_Media3")
										
			elseif periodo_cons=2 then
				f2=RS1("NU_Faltas")
				va_m32=RS1("VA_Media3")
				va_rec_sem12=RS1("VA_Rec")

			elseif periodo_cons=3 then
				f3=RS1("NU_Faltas")
				va_m33=RS1("VA_Media3")
								
			elseif periodo_cons=4 then
				f4=RS1("NU_Faltas")
				va_m34=RS1("VA_Media3")
				va_rec_sem34=RS1("VA_Rec")
				
			elseif periodo_cons=5 then
				va_m35=RS1("VA_Media3")
				
			end if	
		end if	
	NEXT

	if isnull(f1) or f1="&nbsp;" or f1="" then
		f1="&nbsp;"
	end if	

	if isnull(f2) or f2="&nbsp;" or f2="" then
		f2="&nbsp;"
	end if	
	
	if isnull(f3) or f3="&nbsp;" or f3="" then
		f3="&nbsp;"
	end if	
	
	if isnull(f4) or f4="&nbsp;" or f4="" then
		f4="&nbsp;"
	end if	
	
		
	if isnull(va_m31) or va_m31="&nbsp;" or va_m31="" then
		dividendo1=0
		divisor1=0
		va_m31="&nbsp;" 
		calcula_media_anual="nao"
	else
		dividendo1=va_m31
		divisor1=1
	end if	
		
	if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
		dividendo2=0
		divisor2=0
		va_m32="&nbsp;"
		calcula_media_anual="nao"
	else
		dividendo2=va_m32
		divisor2=1
	end if
	
	if isnull(va_rec_sem12) or va_rec_sem12="&nbsp;"  or va_rec_sem12="" then
		dividendorec12=0
		divisorrec12=0
		va_rec_sem12="&nbsp;" 
	else
		dividendorec12=va_rec_sem12
		divisorrec12=1
	end if
		
	if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
		dividendo3=0
		divisor3=0
		va_m33="&nbsp;"
		calcula_media_anual="nao"
	else
		dividendo3=va_m33
		divisor3=1
	end if		


	if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
		dividendo4=0
		divisor4=0
		va_m34="&nbsp;"
		calcula_media_anual="nao"
	else
		dividendo4=va_m34
		divisor4=1
	end if	
	
	if isnull(va_rec_sem34) or va_rec_sem34="&nbsp;"  or va_rec_sem34="" then
		dividendorec34=0
		divisorrec34=0
		va_rec_sem34="&nbsp;" 
	else
		dividendorec34=va_rec_sem34
		divisorrec34=1				
	end if		

	if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
		dividendo5=0
		divisor5=0
		va_m35="&nbsp;"
	else
		dividendo5=va_m35
		divisor5=1
	end if		
	
	dividendo1=dividendo1*1
	dividendo2=dividendo2*1
	dividendo3=dividendo3*1
	dividendo4=dividendo4*1
	
	divisor_ms1=divisor1+divisor2
	
	if divisor_ms1=0 then
		va_m31a="&nbsp;"
		va_m32a="&nbsp;"					

	else
		
		media_sem1=(dividendo1+dividendo2)/divisor_ms1	
		media_sem1 = arredonda(media_sem1,"mat_dez",1,0)
		media_sem1=media_sem1*1
		res_apr=res_apr*1

		if divisorrec12=0 then
'alterado por �ltimo
'===============================
'antes		media_sem1 ="&nbsp;" 		
			va_m31a="&nbsp;"
			va_m32a="&nbsp;"
'================================			
		else
			if media_sem1<res_apr then
			
				divisor_ms1a=divisor1+divisorrec12						
				if divisor_ms1a=0 then
					dividendo1a=dividendo1	
					va_m31a="&nbsp;"							
				else		
					dividendo1=dividendo1*1	
					if dividendo1<res_apr then							
						dividendo1a=(dividendo1+dividendorec12)/divisor_ms1a
						dividendo1a = arredonda(dividendo1a,"mat_dez",1,0)
						dividendo1a=dividendo1a*1
						if dividendo1a>dividendo1 then
							va_m31a=dividendo1a
						else
							va_m31a=dividendo1
						end if		
					else
						dividendo1a=dividendo1
						va_m31a=dividendo1a							
					end if	
				end if	
			
				divisor_ms1b=divisor1+divisorrec12												
				if divisor_ms1b=0 then
					dividendo2a=dividendo2	
					va_m32a="&nbsp;"							
				else	
					dividendo2=dividendo2*1	
					if dividendo2<res_apr then									
						dividendo2a=(dividendo2+dividendorec12)/divisor_ms1b
						dividendo2a = arredonda(dividendo2a,"mat_dez",1,0)	
						dividendo2a=dividendo2a*1
						if dividendo2a>dividendo2 then
							va_m32a=dividendo2a		
						else
							va_m32a=dividendo2	
						end if							
					else
						dividendo2a=dividendo2
						va_m32a=dividendo2a						
					end if								
				end if							
				dividendo1a=dividendo1*1
				dividendo2a=dividendo2*1						
				media_sem1=(dividendo1a+dividendo2a)/divisor_ms1	
				media_sem1 = arredonda(media_sem1,"mat_dez",1,0)							
			else
				va_m31a=va_m31
				va_m32a=va_m32											
			end if		
		end if		
	end if
	
	divisor_ms2=divisor3+divisor4

	if divisor_ms2=0 then
		va_m33a="&nbsp;"
		va_m34a="&nbsp;"						
	else

		media_sem2=(dividendo3+dividendo4)/divisor_ms2		
		media_sem2 = arredonda(media_sem2,"mat_dez",1,0)							
		media_sem2=media_sem2*1
		res_apr=res_apr*1

		if divisorrec34=0 then
'alterado por �ltimo
'===============================
'antes		media_sem2 ="&nbsp;" 
			va_m33a="&nbsp;"
			va_m34a="&nbsp;"
'===============================				
		else
			if media_sem2<res_apr then
			
				divisor_ms2a=divisor3+divisorrec34						
				if divisor_ms2a=0 then
					dividendo3a=dividendo3	
					va_m33a="&nbsp;"							
				else
					dividendo3=dividendo3*1	
					if dividendo3<res_apr then										
						dividendo3a=(dividendo3+dividendorec34)/divisor_ms2a
						dividendo3a = arredonda(dividendo3a,"mat_dez",1,0)
						dividendo3a=dividendo3a*1
						if dividendo3a>dividendo3 then
							va_m33a=dividendo3a	
						else
							va_m33a=dividendo3	
						end if						
					else
						dividendo3a=dividendo3
						va_m33a=dividendo3a						
					end if
				end if	
			
				divisor_ms2b=divisor4+divisorrec34												
				if divisor_ms2b=0 then
					dividendo4a=dividendo4
					va_m34a="&nbsp;"								
				else		
					dividendo4=dividendo4*1	
					if dividendo4<res_apr then								
						dividendo4a=(dividendo4+dividendorec34)/divisor_ms2b
						dividendo4a = arredonda(dividendo4a,"mat_dez",1,0)	
						dividendo4a=dividendo4a*1
						if dividendo4a>dividendo4 then
							va_m34a=dividendo4a
						else
							va_m34a=dividendo4
						end if							
					else
						va_m34a=dividendo4
					end if								
				end if		
				dividendo3a=dividendo_mae3*1
				dividendo4a=dividendo_mae4*1						
				media_sem2=(dividendo3a+dividendo4a)/divisor_ms2		
				media_sem2 = arredonda(media_sem2,"mat_dez",1,0)												
			else
				va_m33a=va_m33
				va_m34a=va_m34										
			end if		
		end if												
	end if			
							
	if divisor1=0 or divisor2=0 then
		media_sem1="&nbsp;"				
	end if	
		
	if divisor3=0 or divisor4=0  then
		media_sem2="&nbsp;"										
	end if	
																	

	Calc_Med_T_F_F_N=va_m31&"#!#"&va_m32&"#!#"&media_sem1&"#!#"&va_rec_sem12&"#!#"&va_m31a&"#!#"&va_m32a&"#!#"&va_m33&"#!#"&va_m34&"#!#"&media_sem2&"#!#"&va_rec_sem34&"#!#"&va_m33a&"#!#"&va_m34a&"#!#"&va_m35&"#!#"&f1&"#!#"&f2&"#!#"&f3&"#!#"&f4


end function














'===========================================================================================================================================
Function Calc_Med_T_F_T_N(unidade, curso, co_etapa, turma, cod_aluno, co_materia, caminho_nota, tb_nota, qtd_periodos, periodo_m2, periodo_m3,tipo_calculo, outro)


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
'	res_rec=RSra("NU_Valor_M2")
'	res_rep=RSra("NU_Valor_M3")	
'
'	if periodo_m2>0 then
'		retira_periodo_m2=1
'	else
'		retira_periodo_m2=0			
'	end if
'	
'	if periodo_m3>0 then
'		retira_periodo_m3=1
'	else
'		retira_periodo_m3=0			
'	end if
'					
'	medias_necessarias=qtd_periodos-retira_periodo_m2-retira_periodo_m3
		
	co_mat_mae=co_materia

	
	conta_m31=0
	conta_m32=0
	conta_r12=0	
	conta_m33=0
	conta_m34=0
	conta_r34=0	
	conta_m35=0
				
	dividendo1=0
	divisor1=0
	dividendo2=0
	divisor2=0			
	dividendorec12=0	
	divisorrec12=0			
	dividendo3=0	
	divisor3=0											
	dividendo4=0
	divisor4=0			
	dividendorec34=0
	divisorrec34=0					
	dividendo5=0
	divisor5=0			

	dividendo_mae1=0
	dividendo_mae2=0
	dividendo_mae3=0
	dividendo_mae4=0	
	dividendo_mae5=0
	divisor_mae=0		
			
	divisor_ms1=0
	divisor_ms1a=0
	divisor_ms1b=0		
	
	divisor_ms2=0
	divisor_ms2a=0
	divisor_ms2b=0		
	
	Set RS1a = Server.CreateObject("ADODB.Recordset")
	SQL1a = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_mat_mae&"' order by NU_Ordem_Boletim"
	RS1a.Open SQL1a, CON0					
	if RS1a.EOF then
		response.Write("Cadastramento Incorreto para a Mat�ria "&co_mat_mae&" em TB_Programa_Aula")
	else
		co_materia_fil_check=0
		while not RS1a.EOF
			co_mat_fil= RS1a("CO_Materia")	
			Set RSa = Server.CreateObject("ADODB.Recordset")
			SQLa = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& co_mat_fil &"'"
			RSa.Open SQLa, CON0

			peso_fil= RSa("NU_Peso")					

			if isnull(peso_fil) or peso_fil="" then
				peso_fil=1
			end if
			
			divisor_mae=divisor_mae+peso_fil					
			for periodo_cons=1 to qtd_periodos
				nota="&nbsp;"
				falta="&nbsp;"
				rec="&nbsp;"					
				media="&nbsp;"				
			
				Set RS1 = Server.CreateObject("ADODB.Recordset")
				SQL1 = "SELECT * FROM "&tb_nota&" where CO_Matricula ="& cod_aluno &" AND CO_Materia ='"& co_mat_fil&"' AND CO_Materia_Principal ='"& co_mat_mae&"' And NU_Periodo="&periodo_cons

				RS1.Open SQL1, CONn
				if RS1.EOF then
					if periodo_cons=1 then
						f1="&nbsp;"
						va_m31="&nbsp;"
						conta_m31=conta_m31
					elseif periodo_cons=2 then
						f2="&nbsp;"
						va_m32="&nbsp;"
						va_rec_sem12="&nbsp;"
						conta_m32=conta_m32		
						conta_r12=conta_r12											
					elseif periodo_cons=3 then
						f3="&nbsp;"
						va_m33="&nbsp;"
						conta_m33=conta_m33							
					elseif periodo_cons=4 then
						f4="&nbsp;"
						va_m34="&nbsp;"
						conta_m34=conta_m34		
						conta_r34=conta_r34							
					elseif periodo_cons=5 then
						va_m35="&nbsp;"						
					end if									
					if periodo_cons<>5 then
						calcula_media_anual="nao"
						conta_m35=conta_m35														
					end if
				else
					if periodo_cons=1 then
						f1=RS1("NU_Faltas")
						va_m31=RS1("VA_Media3")
						if isnull(va_m31) or va_m31="&nbsp;" or va_m31="" then
							dividendo1=dividendo1
							divisor1=0
							conta_m31=conta_m31
							va_m31="&nbsp;" 
							calcula_media_anual="nao"
						else
							dividendo1=dividendo1+(va_m31*peso_fil)
							divisor1=1
							conta_m31=conta_m31+1							
						end if	
					elseif periodo_cons=2 then
						f2=RS1("NU_Faltas")
						va_m32=RS1("VA_Media3")
						va_rec_sem12=RS1("VA_Rec")
						if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
							dividendo2=dividendo2
							divisor2=0
							va_m32="&nbsp;"
							calcula_media_anual="nao"
							conta_m32=conta_m32									
						else
							dividendo2=dividendo2+(va_m32*peso_fil)
							divisor2=1
							conta_m32=conta_m32+1								
						end if
													
						if isnull(va_rec_sem12) or va_rec_sem12="&nbsp;"  or va_rec_sem12="" then
							dividendorec12=dividendorec12
							divisorrec12=0
							va_rec_sem12="&nbsp;" 
							conta_r12=conta_r12								
						else
							dividendorec12=dividendorec12+(va_rec_sem12*peso_fil)
							divisorrec12=1
							conta_r12=conta_r12+1							
						end if								
					elseif periodo_cons=3 then
						f3=RS1("NU_Faltas")
						va_m33=RS1("VA_Media3")
						if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
							dividendo3=dividendo3
							divisor3=0
							va_m33="&nbsp;"
							calcula_media_anual="nao"
							conta_m33=conta_m33								
						else
							dividendo3=dividendo3+(va_m33*peso_fil)
							divisor3=1
							conta_m33=conta_m33+1							
						end if
					elseif periodo_cons=4 then
						f4=RS1("NU_Faltas")
						va_m34=RS1("VA_Media3")
						va_rec_sem34=RS1("VA_Rec")	
						if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
							dividendo4=dividendo4
							divisor4=0
							va_m34="&nbsp;"
							calcula_media_anual="nao"
							conta_m34=conta_m34							
						else
							dividendo4=dividendo4+(va_m34*peso_fil)
							divisor4=1
							conta_m34=conta_m34+1							
						end if	
						
						if isnull(va_rec_sem34) or va_rec_sem34="&nbsp;"  or va_rec_sem34="" then
							dividendorec34=dividendorec34
							divisorrec34=0
							va_rec_sem34="&nbsp;" 
							conta_r34=conta_r34							
						else
							dividendorec34=dividendorec34+(va_rec_sem34*peso_fil)
							divisorrec34=1	
							conta_r34=conta_r34+1											
						end if														
					elseif periodo_cons=5 then
						va_m35=RS1("VA_Media3")		
						if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
							dividendo5=dividendo5
							divisor5=0
							va_m35="&nbsp;"
							conta_m35=conta_m35							
						else
							dividendo5=dividendo5+(va_m35*peso_fil)
							divisor5=1
							conta_m35=conta_m35+1							
						end if												
					end if	

				end if	
			NEXT	
			co_materia_fil_check=co_materia_fil_check+1	
		RS1a.MOVENEXT
		WEND
		
		if isnull(f1) or f1="&nbsp;" or f1="" then
			f1="&nbsp;"
		end if	
	
		if isnull(f2) or f2="&nbsp;" or f2="" then
			f2="&nbsp;"
		end if	
		
		if isnull(f3) or f3="&nbsp;" or f3="" then
			f3="&nbsp;"
		end if	
		
		if isnull(f4) or f4="&nbsp;" or f4="" then
			f4="&nbsp;"
		end if	
								
		if divisor_mae=0 then
			va_m1=dividendo1
			va_m2=dividendo2
			va_m3=dividendo3
			va_m4=dividendo4	
			va_m5=dividendo5	
			va_m1a="&nbsp;"	
			va_m2a="&nbsp;"	
			va_m3a="&nbsp;"	
			va_m4a="&nbsp;"																					
			va_mrec12=dividendorec12	
			va_mrec34=dividendorec34	
			dividendo_mae1=dividendo1
			dividendo_mae2=dividendo2
			dividendo_mae3=dividendo3
			dividendo_mae4=dividendo4	
			dividendo_mae5=dividendo5	
			dividendo_maerec12=dividendorec12	
			dividendor_maeec34=dividendorec34							
		else

			conta_m31=conta_m31*1
			conta_m32=conta_m32*1
			conta_r12=conta_r12*1
			conta_m33=conta_m33*1
			conta_m34=conta_m34*1
			conta_r34=conta_r34*1
			conta_m35=conta_m35*1	
			co_materia_fil_check=co_materia_fil_check*1	
				
			if conta_m31=co_materia_fil_check then		
				va_m1=dividendo1/divisor_mae
				dividendo_mae1 = arredonda(va_m1,"mat_dez",1,0)									
				va_mae_m1 = arredonda(va_m1,"mat_dez",1,0)							
			else
				va_m1="&nbsp;"
				media_sem1="&nbsp;"	
				va_m1a="&nbsp;"	
				va_mae_m1="&nbsp;"														
				calcula_media_anual="nao" 				
			end if
			
			if conta_m32=co_materia_fil_check then				
				va_m2=dividendo2/divisor_mae
				dividendo_mae2 = arredonda(va_m2,"mat_dez",1,0)								
				va_mae_m2 = arredonda(va_m2,"mat_dez",1,0)				
			else
				va_m2="&nbsp;"
				media_sem1="&nbsp;"	
				va_m2a="&nbsp;"	
				va_mae_m2="&nbsp;"														
				calcula_media_anual="nao" 					
			end if			
			
			if conta_m33=co_materia_fil_check then											
				va_m3=dividendo3/divisor_mae
				dividendo_mae3 = arredonda(va_m3,"mat_dez",1,0)
				va_mae_m3 = arredonda(va_m3,"mat_dez",1,0)									
			else
				va_m3="&nbsp;"
				media_sem2="&nbsp;"
				va_m3a="&nbsp;"	
				va_mae_m3="&nbsp;"														
				calcula_media_anual="nao" 					
			end if			
			
			if conta_m34=co_materia_fil_check then						
				va_m4=dividendo4/divisor_mae
				va_mae_m4 = arredonda(va_m4,"mat_dez",1,0)	
				dividendo_mae4 = arredonda(va_m4,"mat_dez",1,0)					
			else
				va_m4="&nbsp;"
				media_sem2="&nbsp;"
				va_m4a="&nbsp;"	
				va_mae_m4="&nbsp;"														
				calcula_media_anual="nao" 					
			end if	
			
			if conta_m35=co_materia_fil_check then																		
				va_m5=dividendo5/divisor_mae	
				va_m5 = arredonda(va_m5,"mat_dez",1,0)	
				dividendo_mae5 = arredonda(va_m5,"mat_dez",1,0)														
			else
				va_m5="&nbsp;"
			end if	
						
			if conta_r12 =co_materia_fil_check then
				va_mrec12=dividendorec12/divisor_mae						
				va_mrec12 = arredonda(va_mrec12,"mat_dez",1,0)
				dividendo_maerec12 = arredonda(va_mrec12,"mat_dez",1,0)									
			else
				va_mrec12="&nbsp;"
				dividendo_maerec12=0	
				va_m1a="&nbsp;"
				va_m2a="&nbsp;"	
				media_sem1="&nbsp;"								
			end if	
	
			if conta_r34 =co_materia_fil_check then
				va_mrec34=dividendorec34/divisor_mae
				va_mrec34 = arredonda(va_mrec34,"mat_dez",1,0)				
				dividendo_maerec34 = arredonda(va_mrec34,"mat_dez",1,0)				
			else			
				va_mrec34="&nbsp;"
				dividendo_maerec34=0
				va_m3a="&nbsp;"
				va_m4a="&nbsp;"				
			end if							
		end if

														
		divisor_ms1=divisor1+divisor2
		divisor_ms1_teste=divisor_ms1*co_materia_fil_check	

		if divisor_ms1_teste<(co_materia_fil_check*2) then
			va_m1a="&nbsp;"
			va_m2a="&nbsp;"							
		else
			dividendo_mae1=dividendo_mae1*1
			dividendo_mae2=dividendo_mae2*1
			media_sem1=(dividendo_mae1+dividendo_mae2)/divisor_ms1		

			media_sem1 = arredonda(media_sem1,"mat_dez",1,0)
			media_sem1=media_sem1*1
			res_apr=res_apr*1

			if media_sem1<res_apr then
			
				divisor_ms1a=divisor1+divisorrec12	
			
				if divisor_ms1a=0 then
					dividendo1a=dividendo_mae1
					va_m1a="&nbsp;"								
				else
					if divisorrec12=0 then		
						dividendo1a=dividendo_mae1
						va_m1a="&nbsp;"								
					else
						va_mae_m1=va_mae_m1*1
						res_apr=res_apr*1

						if va_mae_m1<res_apr then
							dividendo1a=(dividendo_mae1+dividendo_maerec12)/divisor_ms1a
							dividendo1a = arredonda(dividendo1a,"mat_dez",1,0)
							testedividendo1a=dividendo1a*1
							if testedividendo1a > va_mae_m1 then
								va_m1a=dividendo1a	
							else
								va_m1a=	va_mae_m1
							end if	
						else
							va_m1a=va_mae_m1														
						end if				
					end if												
				end if	
			
				divisor_ms1b=divisor1+divisorrec12												
				if divisor_ms1b=0 then
					dividendo2a=dividendo_mae2	
					va_m2a="&nbsp;"							
				else
					if divisorrec12=0 then		
						dividendo2a=dividendo_mae2	
						va_m2a="&nbsp;"								
					else
						va_mae_m2=va_mae_m2*1
						if va_mae_m2<res_apr then												
							dividendo2a=(dividendo_mae2+dividendo_maerec12)/divisor_ms1b
							dividendo2a = arredonda(dividendo2a,"mat_dez",1,0)								
							testedividendo2a=dividendo2a*1
							if testedividendo2a > va_mae_m2 then
								va_m2a=dividendo2a	
							else
								va_m2a=	va_mae_m2
							end if								
						else
							va_m2a=va_mae_m2
						end if									
					end if												
				end if														
			else
				va_m1a=dividendo_mae1
				va_m2a=dividendo_mae2										
			end if			
		end if
	
		divisor_ms2=divisor3+divisor4
		divisor_ms2_teste=divisor_ms2*co_materia_fil_check

		if divisor_ms2_teste<(co_materia_fil_check*2) then
			va_m3a="&nbsp;"
			va_m4a="&nbsp;"							
		else
			dividendo_mae3=dividendo_mae3*1
			dividendo_mae4=dividendo_mae4*1	
			media_sem2=(dividendo_mae3+dividendo_mae4)/divisor_ms2	
			media_sem2 = arredonda(media_sem2,"mat_dez",1,0)							
			media_sem2=media_sem2*1
			res_apr=res_apr*1

			if media_sem2<res_apr then
			
				divisor_ms2a=divisor3+divisorrec34		
							
				if divisor_ms2a=0 then
					dividendo3a=dividendo_mae3	
					va_m3a="&nbsp;"								
				else		
					if divisorrec34=0 then		
						dividendo3a=dividendo_mae3	
						va_m3a="&nbsp;"							
					else	
						va_mae_m3=va_mae_m3*1
						if va_mae_m3<res_apr then												
							dividendo3a=(dividendo_mae3+dividendo_maerec34)/divisor_ms2a
							dividendo3a = arredonda(dividendo3a,"mat_dez",1,0)
							va_m3a=dividendo3a	
						else
							va_m3a=va_mae_m3
						end if									
					end if	
				end if	
			
				divisor_ms2b=divisor4+divisorrec34												
				if divisor_ms2b=0 then
					dividendo4a=dividendo_mae4
					va_m4a="&nbsp;"								
				else
					if divisorrec34=0 then		
						dividendo3a=dividendo_mae3	
						va_m3a="&nbsp;"							
					else
						va_mae_m4=va_mae_m4*1
						if va_mae_m3<res_apr then													
							dividendo4a=(dividendo_mae4+dividendo_maerec34)/divisor_ms2b
							dividendo4a = arredonda(dividendo4a,"mat_dez",1,0)	
							va_m4a=dividendo4a	
						else
							va_m4a=va_mae_m4
						end if										
					end if																										
				end if																		
			else
				va_m3a=dividendo_mae3
				va_m4a=dividendo_mae4									
			end if													
		end if			
	END IF
																	

	Calc_Med_T_F_T_N=va_mae_m1&"#!#"&va_mae_m2&"#!#"&media_sem1&"#!#"&va_mrec12&"#!#"&va_m1a&"#!#"&va_m2a&"#!#"&va_mae_m3&"#!#"&va_mae_m4&"#!#"&media_sem2&"#!#"&va_mrec34&"#!#"&va_m3a&"#!#"&va_m4a&"#!#"&va_m5&"#!#"&f1&"#!#"&f2&"#!#"&f3&"#!#"&f4


end function














'===========================================================================================================================================

'calcula as m�dias anuais e finais destes respectivos mapas
Function Calc_Med_An_Fin(unidade, curso, co_etapa, turma, vetor_aluno, vetor_materia, caminho_nota, tb_nota, qtd_periodos, periodo_m2, periodo_m3,tipo_calculo, outro)

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
	
	for periodo=1 to qtd_periodos
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

Function apura_resultado_aluno (curso,etapa,vetor_medias)

'	Set CON0 = Server.CreateObject("ADODB.Connection") 
'	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
'	CON0.Open ABRIR0
'	
'	Set RSra = Server.CreateObject("ADODB.Recordset")
'	SQLra = "SELECT * FROM TB_Regras_Aprovacao where CO_Curso = '"&curso&"' and CO_Etapa = '"&etapa&"'"
'	RSra.Open SQLra, CON0	
'			
'	valor_apr=RSra("NU_Valor_Apr")
'	valor_dep=RSra("NU_Valor_Dep")
'	qtd_max_dep=RSra("NU_Qt_Dis_Dep")
'	res_apr=RSra("NO_Expr_Maior_Igual_VL_Abr")
'	res_dep=RSra("NO_Expr_Cond_Verdade_Abr")
'	res_rep=RSra("NO_Expr_Cond_Falso_Abr")
'	qtd_dep=0
'	
''	valor_apr=70
''	valor_dep=50
''	qtd_max_dep=5
''	res_apr="AP"
''	res_dep="DP"
''	res_rep="RP"	
'
'	resultados_materia = split(vetor_medias, "#$#" )
'	libera_resultado="s"
'for r=0 to ubound(resultados_materia)	
'	nota_materia = split(resultados_materia(r), "#!#" )
'
'	md_aluno=nota_materia(0)
'	valor_apr=valor_apr*1
'	valor_dep=valor_dep*1
'	if md_aluno="" or isnull(md_aluno) or md_aluno="&nbsp;" or md_aluno=" "then
'		libera_resultado="n"
'	else
'		if md_aluno >= valor_apr then
'			result_temp="apr"
'		elseif md_aluno >= valor_dep then
'			resultado="dep"
'			qtd_dep=qtd_dep+1
'		else
'			result_temp="rep"			
'		end if
'	end if
'Next
'if 	libera_resultado="s" then
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
'else
'	apura_resultado_aluno="&nbsp;"		
'end if	
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set RSra = Server.CreateObject("ADODB.Recordset")
	SQLra = "SELECT * FROM TB_Regras_Aprovacao where CO_Curso = '"&curso&"' and CO_Etapa = '"&etapa&"'"
	RSra.Open SQLra, CON0	
			
'	valor_apr=RSra("NU_Valor_Apr")
'	valor_dep=RSra("NU_Valor_Dep")
	qtd_max_dep=RSra("NU_Qt_Dis_Dep")
'	res_apr=RSra("NO_Expr_Maior_Igual_VL_Abr")
	res_dep=RSra("NO_Expr_Cond_Verdade_Abr")
'	res_rep=RSra("NO_Expr_Cond_Falso_Abr")
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
		res_aluno=nota_materia(1)
'		response.Write(res_aluno&"<BR>")
	if result_temp="REP" then
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