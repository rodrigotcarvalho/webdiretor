<!--#include file="../../global/funcoes_diversas.asp" -->
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
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
elseif tipo_calculo="boletim" then	
	
	co_materia= split(vetor_materia,"#!#")	
	co_materia_check=0	

	vetor_periodo= split(periodo,"#!#")	

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
		
		if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) then
		
			for f2a=0 to ubound(vetor_periodo)
			periodo_cons=vetor_periodo(f2a)
			
				Set RS1 = Server.CreateObject("ADODB.Recordset")
				SQL1 = "SELECT * FROM "&tb_nota&" where CO_Matricula ="& vetor_aluno &" AND CO_Materia ='"& co_materia(f2)&"' And NU_Periodo="&periodo_cons
'				response.Write(SQL1)
				RS1.Open SQL1, CONn

				if RS1.EOF then
					nota=""
					rec=""
					media=""
					falta=""
					media_soma=0
					calcula_media_anual="nao"	
					conceito=""		
					conceito_anual=""		
					conceito_recup=""	
					conceito_final=""	
				else
					'nota=RS1("VA_Media2")
					'rec=RS1("VA_Rec")
					media=RS1("VA_Media3")
'					if nota="" or isnull(nota) or nota=0 then
'						nota=""
'					else
'						nota = formatNumber(nota/10,1)	
'					end if
'					if rec="" or isnull(rec) then
'						rec=""
'					else
'						rec = formatNumber(rec/10,1)	
'					end if
					if media="" or isnull(media) or media=0 then
						media=""
						media_soma=0
						calcula_media_anual="nao"				
					else
						media_soma=media
						media = formatNumber(media/10,1)						
						if calcula_media_anual="nao" then
						else
							calcula_media_anual="sim"
						end if
					end if
					
'					falta=RS1("NU_Faltas")					
'					if falta=0 then
'						falta=""
'					end if
				end if						
				if periodo_cons=4 then				
						if curso=1 and co_etapa<6 and (co_materia(f2)="ARTC" or co_materia(f2)="EART" or co_materia(f2)="EFIS" or co_materia(f2)="INGL") then									
							teste_media = isnumeric(media)							
							if teste_media=TRUE then							
								if media > 9 then
								conceito="E"
								elseif (media > 7) and (media <= 9) then
								conceito="MB"
								elseif (media > 6) and (media <= 7) then							
								conceito="B"
								elseif (media > 4.9) and (media <= 6) then
								conceito="R"
								else							
								conceito="I"
								end if	
							end if	
						else
							conceito=media				
						end if						
					if co_materia_check=0 AND periodo_cons=1 then
						'vetor_quadro=nota&"#!#"&rec&"#!#"&media&"#!#"&falta
						vetor_quadro=conceito
					elseif periodo_cons=1 then
						'vetor_quadro=vetor_quadro&nota&"#!#"&rec&"#!#"&media&"#!#"&falta					
						vetor_quadro=vetor_quadro&conceito
					else
						'vetor_quadro=vetor_quadro&"#!#"&nota&"#!#"&rec&"#!#"&media&"#!#"&falta
						vetor_quadro=vetor_quadro&"#!#"&conceito
					end if
					qtd_periodos=ubound(vetor_periodo)+1
				'response.Write(qtd_periodos)
				'Response.End()
					divisor_anual=qtd_periodos*1	
					
					if calcula_media_anual="sim" then
						media_calc1=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, vetor_aluno, co_materia(f2), caminho_nota, tb_nota, 6, 5, 6, "anual", 0)	
						resultados=split(media_calc1,"#!#")
						media_anual=resultados(0)
						resultado_anual=resultados(1)
						decimo = media_anual - Int(media_anual)
						If decimo >= 0.5 Then
							nota_arredondada = Int(media_anual) + 1
							media_anual=nota_arredondada
						Else
							nota_arredondada = Int(media_anual)
							media_anual=nota_arredondada					
						End If
						if media_anual >67 and media_anual <70 then
							media_anual =70
						end if							
						media_anual=media_anual/10									
						media_anual = formatNumber(media_anual,1)						
					else
						media_anual=""
						resultado_anual=""
					end if			
					if curso=1 and co_etapa<6 and (co_materia(f2)="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then									
							teste_media_anual = isnumeric(media_anual)								
							if teste_media_anual=TRUE then								
							if media_anual > 9 then
							conceito_anual="E"
							elseif (media_anual > 7) and (media_anual <= 9) then
							conceito_anual="MB"
							elseif (media_anual > 6) and (media_anual <= 7) then							
							conceito_anual="B"
							elseif (media_anual > 4.9) and (media_anual <= 6) then
							conceito_anual="R"
							else							
							conceito_anual="I"
							end if	
						end if	
					else
						conceito_anual=media_anual				
					end if						
							
					vetor_quadro=vetor_quadro&"#!#"&conceito_anual&"#!#"&resultado_anual

				elseif periodo_cons=5 then									
					if media="" or isnull(media) or media_anual="" or isnull(media_anual)then
						media_recup=""
						resultado_recup=""
					else
						media_calc2=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, vetor_aluno, co_materia(f2), caminho_nota, tb_nota, 6, 5, 6, "recuperacao", 0)	
						resultados=split(media_calc2,"#!#")
						media_recup=resultados(0)
						resultado_recup=resultados(1)					
							decimo = media_recup - Int(media_recup)
							If decimo >= 0.5 Then
								nota_arredondada = Int(media_recup) + 1
								media_recup=nota_arredondada
							Else
								nota_arredondada = Int(media_recup)
								media_recup=nota_arredondada					
							End If
							media_recup=media_recup/10									
							media_recup = formatNumber(media_recup,1)																													
					end if					
						if curso=1 and co_etapa<6 and (co_materia(f2)="ARTC" or co_materia(f2)="EART" or co_materia(f2)="EFIS" or co_materia(f2)="INGL") then									
							teste_media_recup = isnumeric(media_recup)								
							if teste_media_recup=TRUE then									
								if media_recup > 9 then
								conceito_recup="E"
								elseif (media_recup > 7) and (media_recup <= 9) then
								conceito_recup="MB"
								elseif (media_recup > 6) and (media_recup <= 7) then							
								conceito_recup="B"
								elseif (media_recup > 469) and (media_recup <= 6) then
								conceito_recup="R"
								else							
								conceito_recup="I"
								end if	
							end if	
						else
							conceito_recup=media_recup				
						end if		
					vetor_quadro=vetor_quadro&"#!#"&media&"#!#"&conceito_recup&"#!#"&resultado_recup														
				elseif periodo_cons=6 then									
					if media="" or isnull(media) or media_anual="" or isnull(media_anual)  or media_recup="" or isnull(media_recup)then
						media_final=""
						resultado_final=""							
					else
						media_calc3=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, vetor_aluno, co_materia(f2), caminho_nota, tb_nota, 6, 5, 6, "final", 0)	
						resultados=split(media_calc3,"#!#")
						media_final=resultados(0)
						resultado_final=resultados(1)					
							decimo = media_final - Int(media_final)
							If decimo >= 0.5 Then
								nota_arredondada = Int(media_final) + 1
								media_final=nota_arredondada
							Else
								nota_arredondada = Int(media_final)
								media_final=nota_arredondada					
							End If
							media_final=media_final/10	
							media_final = formatNumber(media_final,1)
					end if
						if curso=1 and co_etapa<6 and (co_materia(f2)="ARTC" or co_materia(f2)="EART" or co_materia(f2)="EFIS" or co_materia(f2)="INGL") then	
							teste_media_final = isnumeric(media_final)								
							if teste_media_final=TRUE then									
								if media_final > 9 then
								conceito_final="E"
								elseif (media_final > 7) and (media_final <= 9) then
								conceito_final="MB"
								elseif (media_final > 6) and (media_final <= 7) then							
								conceito_final="B"
								elseif (media_final > 4.9) and (media_final <= 6) then
								conceito_final="R"
								else							
								conceito_final="I"
								end if	
							end if	
						else
							conceito_final=media_final				
						end if						
						
					vetor_quadro=vetor_quadro&"#!#"&conceito_final&"#!#"&resultado_final									
				else

				
					if curso=1 and co_etapa<6 and (co_materia(f2)="ARTC" or co_materia(f2)="EART" or co_materia(f2)="EFIS" or co_materia(f2)="INGL") then	
						teste_media = isnumeric(media)							
						if teste_media=TRUE then									
							if media > 9 then
							conceito="E"
							elseif (media > 7) and (media <= 9) then
							conceito="MB"
							elseif (media > 6) and (media <= 7) then							
							conceito="B"
							elseif (media > 4.9) and (media <= 6) then
							conceito="R"
							else							
							conceito="I"
							end if	
						end if						
					else
						conceito=media				
					end if		
					
									
					if co_materia_check=0 AND periodo_cons=1 then
						vetor_quadro=conceito
					elseif periodo_cons=1 then
						vetor_quadro=vetor_quadro&conceito
					else
						vetor_quadro=vetor_quadro&"#!#"&conceito
					end if
				end if
			Next	
		vetor_quadro=vetor_quadro&"#$#"	

		
'response.Write(vetor_quadro	&" A <br>")		
		elseif (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) then
'			nota_filha_acumula_1=0						
'			rec_filha_acumula_1=0								
			media_filha_acumula_1=0												
'			nota_filha_acumula_2=0							
'			rec_filha_acumula_2=0								
			media_filha_acumula_2=0
'			nota_filha_acumula_3=0							
'			rec_filha_acumula_3=0								
			media_filha_acumula_3=0
			media_filha_acumula_4=0		
			vetor_mae_filhas=""
	
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
			co_materia_fil_check=0 			
			co_materia_mae_fil= split(vetor_mae_filhas,"#!#")
			soma_mae=0			
			for f3=0 to ubound(co_materia_mae_fil)	
'PARA INCLUIR A LINHA DA MATÉRIA MÃE SEM APARECER NOTAS==================================================
				if f3=0 then
					for f3a=0 to ubound(vetor_periodo)
						periodo_cons=vetor_periodo(f3a)	
'						response.Write(" i "&periodo_cons&"<BR>")
'						nota=""
'						rec=""
						media=""
'						falta=""
						media_soma=0
							
						if periodo_cons=4 then
							media_anual=""
							resultado_anual=""																
							vetor_quadro=vetor_quadro&"#!#"&media_anual&"#!#"&resultado_anual
						elseif periodo_cons=5 then
							media_recup=""
							resultado_recup=""																
							vetor_quadro=vetor_quadro&"#!#"&media&"#!#"&media_recup&"#!#"&resultado_recup
						elseif periodo_cons=6 then
							media_final=""
							resultado_final=""																
							vetor_quadro=vetor_quadro&"#!#"&media_final&"#!#"&resultado_final
						else	
							if co_materia_check=0 AND periodo_cons=1 then
								vetor_quadro=media
							elseif periodo_cons=1 then
								vetor_quadro=vetor_quadro&media
							else
								vetor_quadro=vetor_quadro&"#!#"&media
							end if					
							'	vetor_quadro=vetor_quadro&nota&"#!#"&rec&"#!#"&media&"#!#"&falta
							'end if
						end if	
					next
'========================================================================================================
				else
					maior_periodo=vetor_periodo(ubound(vetor_periodo))
				
					for f3a=0 to ubound(vetor_periodo)
						periodo_cons=vetor_periodo(f3a)		
			
						Set RS3 = Server.CreateObject("ADODB.Recordset")
						SQL3 = "SELECT * FROM "&tb_nota&" where CO_Matricula ="& vetor_aluno &" AND CO_Materia ='"& co_materia_mae_fil(f3)&"' And NU_Periodo="&periodo_cons
	'response.Write(SQL3&"<BR>")
						RS3.Open SQL3, CONn
						
		
						if RS3.EOF then
							media=""
							media_soma=0
							if periodo_cons<5 then
								calcula_media_anual="nao"	
							end if	
						else
							media=RS3("VA_Media3")	

							if media="" or isnull(media) then
								media_soma=0
								if periodo_cons<5 then
									calcula_media_anual="nao"	
								end if									
							else
								media = formatNumber(media/10,1)
								media_soma=media	
								calcula_media_anual="sim"					
							end if
								
						end if	
						'response.Write(ubound(vetor_periodo)&"<BR>")

						if ubound(vetor_periodo)<3 then
							calcula_media_anual="nao"	
						end if		
						'response.Write(calcula_media_anual&"<BR>")
										
'PREPARANDO AS NOTAS PARA SEREM INCLUÍDAS NAS MÉDIAS======================================						
						if media="" or isnull(media) then
							media_filha_acumula="NULO"	
						else
							media_filha_acumula=media_soma
						end if 
'===================================================================================================						
'CALCULA MÉDIA DAS FILHAS===========================================================================	
					
						if periodo_cons=4 then
							soma_media_fil_4=media
							if curso=1 and co_etapa<6 and (co_materia_mae_fil(f3)="ARTC" or co_materia_mae_fil(f3)="EART" or co_materia_mae_fil(f3)="EFIS" or co_materia_mae_fil(f3)="INGL") then									
								if media<>"&nbsp;" then									
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

							if soma_media_fil_1<>"&nbsp;" and soma_media_fil_2<>"&nbsp;" and soma_media_fil_3<>"&nbsp;" and soma_media_fil_4<>"&nbsp;" then
								soma_filhas=(soma_media_fil_1*1)+(soma_media_fil_2*1)+(soma_media_fil_3*1)+(soma_media_fil_4*1)
								media_anual=soma_filhas/4
								media_anual=media_anual*10						
								decimo = formatNumber(media_anual - Int(media_anual),1)
								If decimo >= 0.5 Then
									nota_arredondada = Int(media_anual) + 1
									media_anual=nota_arredondada
								Else
									nota_arredondada = Int(media_anual)
									media_anual=nota_arredondada					
								End If
								media_anual = formatNumber(media_anual,1)								
								if media_anual>67 and media_anual<70then
									media_anual=70
								end if
								media_anual=media_anual/10		
								media_anual = formatNumber(media_anual,1)																
							else
								media_anual=""
							end if	
							resultado_anual=""																
							vetor_quadro=vetor_quadro&"#!#"&conceito&"#!#"&media_anual&"#!#"&resultado_anual
						elseif periodo_cons=5 then
							media_recup=""
							resultado_recup=""																
							vetor_quadro=vetor_quadro&"#!#"&media&"#!#"&media_recup&"#!#"&resultado_recup
						elseif periodo_cons=6 then
							media_final=""
							resultado_final=""																
							vetor_quadro=vetor_quadro&"#!#"&media_final&"#!#"&resultado_final
						else	
							if periodo_cons=1 then
								soma_media_fil_1=media
							elseif periodo_cons=2 then
								soma_media_fil_2=media
							elseif periodo_cons=3 then
								soma_media_fil_3=media							
							end if

						
							if curso=1 and co_etapa<6 and (co_materia_mae_fil(f3)="ARTC" or co_materia_mae_fil(f3)="EART" or co_materia_mae_fil(f3)="EFIS" or co_materia_mae_fil(f3)="INGL") then									
								if media<>"&nbsp;" then									
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
							
							if co_materia_check=0 AND periodo_cons=1 then
								vetor_quadro=conceito
							elseif periodo_cons=1 then
								vetor_quadro=vetor_quadro&conceito
							else
								vetor_quadro=vetor_quadro&"#!#"&conceito
							end if					
							'	vetor_quadro=vetor_quadro&nota&"#!#"&rec&"#!#"&media&"#!#"&falta
							'end if
						end if	
						
						
'=====================================================================================
'ARMAZENA PARA CALCULAR A MÉDIA DA MÃE================================================
						if media_filha_acumula="NULO" then
							if periodo_cons=1 then	
								media_filha_acumula_1="NULO"
							elseif periodo_cons=2 then	
								media_filha_acumula_2="NULO"
							elseif periodo_cons=3 then				
								media_filha_acumula_3="NULO"
							elseif periodo_cons=4 then					
								media_filha_acumula_4="NULO"
							elseif periodo_cons=5 then					
								media_filha_acumula_5="NULO"
							elseif periodo_cons=6 then					
								media_filha_acumula_6="NULO"																															
							end if						
						else						
							media_filha_acumula=media_filha_acumula*1						
							if periodo_cons=1 and media_filha_acumula_1<>"NULO" then	
								media_filha_acumula_1=media_filha_acumula_1*1
								media_filha_acumula_1=media_filha_acumula_1+media_filha_acumula
							elseif periodo_cons=2 and media_filha_acumula_2<>"NULO" then	
								media_filha_acumula_2=media_filha_acumula_2*1
								media_filha_acumula_2=media_filha_acumula_2+media_filha_acumula
							elseif periodo_cons=3 and media_filha_acumula_3<>"NULO" then	
								media_filha_acumula_3=media_filha_acumula_3*1						
								media_filha_acumula_3=media_filha_acumula_3+media_filha_acumula
							elseif periodo_cons=4 and media_filha_acumula_4<>"NULO" then
								media_filha_acumula_4=media_filha_acumula_4*1						
								media_filha_acumula_4=media_filha_acumula_4+media_filha_acumula
							elseif periodo_cons=5 and media_filha_acumula_5<>"NULO" then
								media_filha_acumula_5=media_filha_acumula_5*1						
								media_filha_acumula_5=media_filha_acumula_5+media_filha_acumula	
							elseif periodo_cons=6 and media_filha_acumula_6<>"NULO" then
								media_filha_acumula_6=media_filha_acumula_6*1						
								media_filha_acumula_6=media_filha_acumula_6+media_filha_acumula																																	
							end if
						end if	
						
						maior_periodo=maior_periodo*1						
						if maior_periodo=1 then	
							media_filha_acumula_2="NULO"
							media_filha_acumula_3="NULO"
							media_filha_acumula_4="NULO"
							media_filha_acumula_5="NULO"
							media_filha_acumula_6="NULO"								
						elseif maior_periodo=2 then				
							media_filha_acumula_3="NULO"
							media_filha_acumula_4="NULO"
							media_filha_acumula_5="NULO"
							media_filha_acumula_6="NULO"							
						elseif maior_periodo=3 then					
							media_filha_acumula_4="NULO"
							media_filha_acumula_5="NULO"
							media_filha_acumula_6="NULO"								
						elseif maior_periodo=4 then					
							media_filha_acumula_5="NULO"
							media_filha_acumula_6="NULO"							
						elseif maior_periodo=5 then					
							media_filha_acumula_6="NULO"																															
						end if								
'========================================================================================	
					next
				end if
				co_materia_fil_check=co_materia_fil_check+1	
			vetor_quadro=vetor_quadro&"#$#"							
			next

'CALCULA A MÉDIA==========================================================================		
					
							divisor_medias=co_materia_fil_check-1
							if media_filha_acumula_1<>"NULO" then							
								nota_media_1=(media_filha_acumula_1*10)/divisor_medias	
								decimo = nota_media_1 - Int(nota_media_1)
								If decimo >= 0.5 Then
									nota_arredondada = Int(nota_media_1) + 1
									nota_media_1=nota_arredondada
								Else
									nota_arredondada = Int(nota_media_1)
									nota_media_1=nota_arredondada					
								End If							
								nota_media_1=formatNumber(nota_media_1/10,1)	
							else
								nota_media_1="&nbsp;"
								calcula_media_anual="nao" 
							end if						

							if media_filha_acumula_2<>"NULO" then	
								nota_media_2=(media_filha_acumula_2*10)/divisor_medias							

								decimo = nota_media_2 - Int(nota_media_2)
								If decimo >= 0.5 Then
									nota_arredondada = Int(nota_media_2) + 1
									nota_media_2=nota_arredondada
								Else
									nota_arredondada = Int(nota_media_2)
									nota_media_2=nota_arredondada					
								End If									
								nota_media_2=formatNumber(nota_media_2/10,1)				
							else
								nota_media_2="&nbsp;"
								calcula_media_anual="nao" 
							end if	
							
							if media_filha_acumula_3<>"NULO" then	
								nota_media_3=(media_filha_acumula_3*10)/divisor_medias														
								decimo = nota_media_3 - Int(nota_media_3)
								If decimo >= 0.5 Then
									nota_arredondada = Int(nota_media_3) + 1
									nota_media_3=nota_arredondada
								Else
									nota_arredondada = Int(nota_media_3)
									nota_media_3=nota_arredondada					
								End If								
								nota_media_3=formatNumber(nota_media_3/10,1)		
							else
								nota_media_3="&nbsp;"
								calcula_media_anual="nao" 								
							end if	

							if media_filha_acumula_4<>"NULO" then	
								nota_media_4=(media_filha_acumula_4*10)/divisor_medias															
								decimo = nota_media_4 - Int(nota_media_4)
								If decimo >= 0.5 Then
									nota_arredondada = Int(nota_media_4) + 1
									nota_media_4=nota_arredondada
								Else
									nota_arredondada = Int(nota_media_4)
									nota_media_4=nota_arredondada					
								End If									
								nota_media_4=formatNumber(nota_media_4/10,1)																	
							else
								nota_media_4="&nbsp;"
								calcula_media_anual="nao" 								
							end if	
							
							if media_filha_acumula_5<>"NULO" then	
								nota_media_5=(media_filha_acumula_5*10)/divisor_medias															
								decimo = nota_media_5 - Int(nota_media_5)
								If decimo >= 0.5 Then
									nota_arredondada = Int(nota_media_5) + 1
									nota_media_5=nota_arredondada
								Else
									nota_arredondada = Int(nota_media_5)
									nota_media_5=nota_arredondada					
								End If									
								nota_media_5=formatNumber(nota_media_5/10,1)	
							else
								nota_media_5="&nbsp;"
							end if									
							
							if media_filha_acumula_6<>"NULO" then									
								nota_media_6=(media_filha_acumula_6*10)/divisor_medias															
								decimo = nota_media_6 - Int(nota_media_6)
								If decimo >= 0.5 Then
									nota_arredondada = Int(nota_media_6) + 1
									nota_media_6=nota_arredondada
								Else
									nota_arredondada = Int(nota_media_6)
									nota_media_6=nota_arredondada					
								End If										
								nota_media_6=formatNumber(nota_media_6/10,1)
							else
								nota_media_6="&nbsp;"
							end if		
							
							
							'response.End()				
							'média anual da mãe	====================================
							IF calcula_media_anual="sim" THEN
								'response.Write(nota_media_1&"+("&nota_media_2&"*1)+("&nota_media_3&"*1)+("&nota_media_4&"<BR>")
								soma_mae=(nota_media_1*1)+(nota_media_2*1)+(nota_media_3*1)+(nota_media_4*1)
								media_anual_mae=soma_mae/4
								'response.Write(media_anual_mae&"<BR>")		
								media_anual_mae=media_anual_mae*10						
								decimo = formatNumber(media_anual_mae - Int(media_anual_mae),1)
								If decimo >= 0.5 Then
									nota_arredondada = Int(media_anual_mae) + 1
									media_anual_mae=nota_arredondada
								Else
									nota_arredondada = Int(media_anual_mae)
									media_anual_mae=nota_arredondada					
								End If
								media_anual_mae = formatNumber(media_anual_mae,1)								
								if media_anual_mae>67 and media_anual_mae<70then
									media_anual_mae=70
								end if
								if nota_media_5<>"&nbsp;" then
									nota_media_5=nota_media_5*10
								end if
								if nota_media_6<>"&nbsp;" then
									nota_media_6=nota_media_6*10								
								end if
								resultados=regra_aprovacao (curso,co_etapa,media_anual_mae,nota_media_5,nota_aux_m2_2,nota_media_6,nota_aux_m3_2,"waboletim")
'if vetor_aluno=31323 then
'								response.Write(resultados&"<BR>")	
'	'response.Write(media_anual_mae)			
'
'end if	

								medias_resultados=split(resultados,"#!#")
								

								res1=medias_resultados(1)
								res2=medias_resultados(3)
								res3=medias_resultados(5)
								m2=medias_resultados(2)
								m3=medias_resultados(4)	
								
								if nota_media_5<>"&nbsp;" then
									nota_media_5=formatNumber(nota_media_5/10,1)
								end if
								if nota_media_6<>"&nbsp;" then
									nota_media_6=formatNumber(nota_media_6/10,1)								
								end if								
								
								if m2="&nbsp;" or m2="" or isnull(m2) then
								else
									m2=formatNumber(m2/10,1)
								end if

								if m3="&nbsp;" or m3="" or isnull(m3) then
								else
									m3=formatNumber(m3/10,1)	
																
								end if								
								media_anual_mae=media_anual_mae/10								
								media_anual_mae = formatNumber(media_anual_mae,1)						
							end if																								
								
						conceito_1=nota_media_1
						conceito_2=nota_media_2
						conceito_3=nota_media_3
						conceito_4=nota_media_4
						conceito_5=nota_media_5
						conceito_6=nota_media_6																														
						conceito_anual=media_anual_mae	
						conceito_recup=m2	
						conceito_final=m3																							

							
		vetor_quadro=vetor_quadro&conceito_1&"#!#"&conceito_2&"#!#"&conceito_3&"#!#"&conceito_4&"#!#"&conceito_anual&"#!#"&res1&"#!#"&conceito_5&"#!#"&conceito_recup&"#!#"&res2&"#!#"&conceito_final&"#!#"&res3									
		vetor_quadro=vetor_quadro&"#$#"	
'vetor_aluno=vetor_aluno*1
'9º ano turma 91 em 2010
'if vetor_aluno=27341 then
'				response.Write(vetor_quadro&"<BR>")							
'				response.end()			
'end if
'if vetor_aluno=30770 then
'	response.Write(vetor_quadro&"<BR><BR>")			
'	'response.Write(media_anual_mae)			
'	
'	'response.End()
'end if	
'if vetor_aluno=31323 then
'	response.Write(vetor_quadro&"<BR>")			
'	'response.Write(media_anual_mae)			
'	
'response.End()
'end if	
		elseif (mae=TRUE and fil=TRUE and in_co=FALSE) then
		
		elseif (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then
			nota_filha_acumula_1=0						
			rec_filha_acumula_1=0								
			media_filha_acumula_1=0
			falta_filha_acumula_1=0												
			nota_filha_acumula_2=0							
			rec_filha_acumula_2=0								
			media_filha_acumula_2=0
			falta_filha_acumula_2=0
			nota_filha_acumula_3=0							
			rec_filha_acumula_3=0								
			media_filha_acumula_3=0
			falta_filha_acumula_3=0
			media_filha_acumula_4=0
			falta_filha_acumula_4=0		
			vetor_mae_filhas=""
	
			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL2 = "SELECT * FROM TB_Programa_Aula where CO_Curso ='"& curso&"' AND CO_Etapa='"& co_etapa&"' AND CO_Materia ='"& co_materia(f2) &"'"
'response.Write(SQL2)
			RS2.Open SQL2, CON0
			
			nu_ordem= RS2("NU_Ordem_Boletim")
			
			proximo_nu_ordem=nu_ordem+1				
			co_disc="nao_localizado"	
			
			WHILE co_disc="nao_localizado"			
				Set RS2a = Server.CreateObject("ADODB.Recordset")
				SQL2a = "SELECT * FROM TB_Programa_Aula where CO_Curso ='"& curso&"' AND CO_Etapa='"& co_etapa&"' AND NU_Ordem_Boletim="& proximo_nu_ordem &" AND IN_MAE=FALSE AND IN_FIL=FALSE AND IN_CO=TRUE AND ISNULL(NU_Peso)"
	'response.Write(SQL2a&" A <BR>")	
				RS2a.Open SQL2a, CON0
	
				if RS2a.EOF then
					proximo_nu_ordem=proximo_nu_ordem+1
				ELSE
					co_disc="localizado"
					co_materia_fil_check=0 
					WHILE co_disc="localizado"	
						Set RS2c = Server.CreateObject("ADODB.Recordset")
						SQL2c = "SELECT * FROM TB_Programa_Aula where CO_Curso ='"& curso&"' AND CO_Etapa='"& co_etapa&"' AND NU_Ordem_Boletim="& proximo_nu_ordem &" AND IN_MAE=FALSE AND IN_FIL=FALSE AND IN_CO=TRUE AND ISNULL(NU_Peso)"
	'response.Write(SQL2c&" B <BR>")						
						RS2c.Open SQL2c, CON0					
						if RS2c.EOF then
							co_disc="finalizado"
						ELSE
							proximo_nu_ordem=proximo_nu_ordem+1						

							co_mat_fil= RS2c("CO_Materia")		
							if co_materia_fil_check=0 then
								vetor_mae_filhas=co_materia(f2)&"#!#"&co_mat_fil
							else
								vetor_mae_filhas=vetor_mae_filhas&"#!#"&co_mat_fil			
							end if
'							RESPONSE.Write(	vetor_mae_filhas&"-"&co_mat_fil&"-"&co_materia_fil_check&"<br>")	
							co_materia_fil_check=co_materia_fil_check+1 									
						end if
					wend	
				end if	
			wend				
			co_materia_fil_check=0 			
			co_materia_mae_fil= split(vetor_mae_filhas,"#!#")
			soma_mae=0			
			for f3=0 to ubound(co_materia_mae_fil)	
					for f3a=0 to ubound(vetor_periodo)
					periodo_cons=vetor_periodo(f3a)		
			
						Set RS3 = Server.CreateObject("ADODB.Recordset")
						SQL3 = "SELECT * FROM "&tb_nota&" where CO_Matricula ="& vetor_aluno &" AND CO_Materia ='"& co_materia_mae_fil(f3)&"' And NU_Periodo="&periodo_cons
	'response.Write(SQL3&" C <BR>")
						RS3.Open SQL3, CONn
		
						if RS3.EOF then
							nota=""
							rec=""
							media=""
							falta=""
							media_soma=0
							falta_filha_acumula=0
						else
							nota=RS3("VA_Media2")
							rec=RS3("VA_Rec")
							media=RS3("VA_Media3")						
							if media="" or isnull(media) or media=0 then
								media_soma=0			
							else
								media_soma=media																	
							end if
							falta=RS3("NU_Faltas")
							if falta=0 then
								falta=""
								falta_filha_acumula=0
							else
								falta_filha_acumula=falta
							end if	
						end if	
'PREPARANDO ASNOTAS PARA SEREM INCLUÍDAS NAS MÉDIAS======================================						
						if nota="" or isnull(nota) then
							nota_filha_acumula=0	
						else
							nota_filha_acumula=nota
						end if 
						
						if rec="" or isnull(rec) then
							rec_filha_acumula=0	
						else
							rec_filha_acumula=rec
						end if 
														
						if media="" or isnull(media) then
							media_filha_acumula=0	
						else
							media_filha_acumula=media
						end if 
'===================================================================================================						
'ARMAZENA PARA CALCULAR A MÉDIA DA MÃE================================================
						if periodo_cons=1 then	
							nota_filha_acumula_1=nota_filha_acumula_1*1	
							nota_filha_acumula=nota_filha_acumula*1									
							nota_filha_acumula_1=nota_filha_acumula_1+nota_filha_acumula
							rec_filha_acumula_1=rec_filha_acumula_1*1	
							rec_filha_acumula=rec_filha_acumula*1									
							rec_filha_acumula_1=rec_filha_acumula_1+rec_filha_acumula	
							media_filha_acumula_1=media_filha_acumula_1*1	
							media_filha_acumula=media_filha_acumula*1									
							media_filha_acumula_1=media_filha_acumula_1+media_filha_acumula
							falta_filha_acumula_1=falta_filha_acumula_1*1
							falta_filha_acumula=falta_filha_acumula*1									
							falta_filha_acumula_1=falta_filha_acumula_1+falta_filha_acumula																					
						elseif periodo_cons=2 then	
							nota_filha_acumula_2=nota_filha_acumula_2*1	
							nota_filha_acumula=nota_filha_acumula*1									
							nota_filha_acumula_2=nota_filha_acumula_2+nota_filha_acumula
							rec_filha_acumula_2=rec_filha_acumula_2*1	
							rec_filha_acumula=rec_filha_acumula*1									
							rec_filha_acumula_2=rec_filha_acumula_2+rec_filha_acumula	
							media_filha_acumula_2=media_filha_acumula_2*1	
							media_filha_acumula=media_filha_acumula*1									
							media_filha_acumula_2=media_filha_acumula_2+media_filha_acumula	
							falta_filha_acumula_2=falta_filha_acumula_2*1
							falta_filha_acumula=falta_filha_acumula*1									
							falta_filha_acumula_2=falta_filha_acumula_2+falta_filha_acumula								
						elseif periodo_cons=3 then	
							nota_filha_acumula_3=nota_filha_acumula_3*1	
							nota_filha_acumula=nota_filha_acumula*1									
							nota_filha_acumula_3=nota_filha_acumula_3+nota_filha_acumula
							rec_filha_acumula_3=rec_filha_acumula_3*1	
							rec_filha_acumula=rec_filha_acumula*1									
							rec_filha_acumula_3=rec_filha_acumula_3+rec_filha_acumula	
							media_filha_acumula_3=media_filha_acumula_3*1	
							'NO TERCEIRO PERÍODO AS MÉDIAS TEM PESO 2
							media_filha_acumula=media_filha_acumula*2		
							'========================================							
							media_filha_acumula_3=media_filha_acumula_3+media_filha_acumula	
							falta_filha_acumula_3=falta_filha_acumula_3*1
							falta_filha_acumula=falta_filha_acumula*1									
							falta_filha_acumula_3=falta_filha_acumula_3+falta_filha_acumula									
							soma_mae=media_filha_acumula_1+media_filha_acumula_2+media_filha_acumula_3
						elseif periodo_cons=4 then
							media_filha_acumula_4=media_filha_acumula_4*1	
							media_filha_acumula=media_filha_acumula*1									
							media_filha_acumula_4=media_filha_acumula_4+media_filha_acumula																				
							falta_filha_acumula_4=falta_filha_acumula_4*1
							falta_filha_acumula=falta_filha_acumula*1									
							falta_filha_acumula_4=falta_filha_acumula_4+falta_filha_acumula	
						end if
'RESPONSE.Write("CO_Matricula ="& vetor_aluno &" AND CO_Materia ='"& co_materia_mae_fil(f3)&"' And NU_Periodo="&periodo_cons&"<BR>")	
'RESPONSE.Write(media_filha_acumula_1&"-"&media_filha_acumula_2&"-"&media_filha_acumula_3&"-"&soma_mae&"<BR>")		

'========================================================================================	
					next
				co_materia_fil_check=co_materia_fil_check+1	
'			vetor_quadro=vetor_quadro&"#$#"							
			next
	
'CALCULA A MÉDIA==========================================================================	
'RESPONSE.end()	
							divisor_medias=10
							nota_media_1=formatNumber(nota_filha_acumula_1/divisor_medias,1)	
							rec_media_1=formatNumber(rec_filha_acumula_1/divisor_medias,1)									
							media_media_1=formatNumber(media_filha_acumula_1/divisor_medias,1)													
							nota_media_2=formatNumber(nota_filha_acumula_2/divisor_medias,1)								
							rec_media_2=formatNumber(rec_filha_acumula_2/divisor_medias,1)									
							media_media_2=formatNumber(media_filha_acumula_2/divisor_medias,1)	
							nota_media_3=formatNumber(nota_filha_acumula_3/divisor_medias,1)								
							rec_media_3=formatNumber(rec_filha_acumula_3/divisor_medias,1)									
							'NO TERCEIRO PERÍODO AS MÉDIAS TEM PESO 2====================================
							media_media_3=formatNumber(media_filha_acumula_3/(divisor_medias*2),1)	
							'============================================================================
							media_media_4=formatNumber(media_filha_acumula_4/divisor_medias,1)								
					
							qtd_periodos=ubound(vetor_periodo)+1
							divisor_anual=qtd_periodos*1

								media_anual_mae=soma_mae/divisor_anual
	'RESPONSE.Write(media_anual_mae&"-"&soma_mae&"-"&divisor_anual&"-"&ubound(vetor_periodo))								
								decimo = media_anual_mae - Int(media_anual_mae)
								If decimo >= 0.5 Then
									nota_arredondada = Int(media_anual_mae) + 1
									media_anual_mae=nota_arredondada
								Else
									nota_arredondada = Int(media_anual_mae)
									media_anual_mae=nota_arredondada					
								End If
								media_anual_mae=media_anual_mae/10
								media_anual_mae = formatNumber(media_anual_mae,1)	

							if media_filha_acumula_4=0 then
								media_final_mae=""
							else
								media_final_mae=((media_anual_mae*6)+(media_filha_acumula_4*4))/100
								decimo = media_final_mae - Int(media_final_mae)
								If decimo >= 0.5 Then
									nota_arredondada = Int(media_final_mae) + 1
									media_final_mae=nota_arredondada
								Else
									nota_arredondada = Int(media_final_mae)
									media_final_mae=nota_arredondada					
								End If
								media_final_mae = formatNumber(media_final_mae,1)																					
							end if	
							
							if nota_media_1=0 then
								nota_media_1=""
							end if
							if rec_media_1=0 then
								rec_media_1=""
							end if
							if media_media_1=0 then
								media_media_1=""
							end if
							if nota_media_2=0 then
								nota_media_2=""
							end if	
							if rec_media_2=0 then
								rec_media_2=""
							end if
							if media_media_2=0 then
								media_media_2=""
							end if																											
							if nota_media_3=0 then
								nota_media_3=""
							end if	
							if rec_media_3=0 then
								rec_media_3=""
							end if
							if media_media_3=0 then
								media_media_3=""
							end if	
							if media_media_4=0 then
								media_media_4=""
							end if	
							if media_anual_mae=0 or media_media_1="" or media_media_2="" or media_media_3="" then
								media_anual_mae=""
							end if																								

							if falta_filha_acumula_1=0 then
							falta_mae_1=""
							end if
							
							if falta_filha_acumula_2=0 then
							falta_mae_2=""
							end if

							if falta_filha_acumula_3=0 then
							falta_mae_3=""
							end if
																												
							if falta_filha_acumula_4=0 then
							falta_mae_4=""
							end if
							
		vetor_quadro=vetor_quadro&nota_media_1&"#!#"&rec_media_1&"#!#"&media_media_1&"#!#"&falta_mae_1&"#!#"&nota_media_2&"#!#"&rec_media_2&"#!#"&media_media_2&"#!#"&falta_mae_2&"#!#"&nota_media_3&"#!#"&rec_media_3&"#!#"&media_media_3&"#!#"&falta_mae_3&"#!#"&media_anual_mae&"#!#"&media_media_4&"#!#"&falta_mae_4&"#!#"&media_final_mae										
		vetor_quadro=vetor_quadro&"#$#"	
'response.Write(vetor_quadro	&"D <br>")						
		end if		
	co_materia_check=co_materia_check+1			
	
'RESPONSE.Write(media_anual_mae&"-"&soma_mae&"-"&divisor_anual&"-"&co_materia_fil_check)					
'RESPONSE.END()	
	NEXT	
else	
end if
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
'					If decimo >= 0.75 Then
'						nota_arredondada = Int(media_anual) + 1
'						media_anual=nota_arredondada
'					elseIf decimo >= 0.25 Then
'						nota_arredondada = Int(media_anual) + 0.5
'						media_anual=nota_arredondada
'					else
'						nota_arredondada = Int(media_anual)
'						media_anual=nota_arredondada											
'					End If		
					If decimo >= 0.5 Then
						nota_arredondada = Int(media_anual) + 1
						media_anual=nota_arredondada
					else
						nota_arredondada = Int(media_anual)
						media_anual=nota_arredondada											
					End If		



					media_anual = formatNumber(media_anual,0)			
					media_anual=media_anual*1						
					if media_anual>67 and media_anual<70 then
						media_anual=70
					end if	
'response.Write("<BR>"&dados_aluno(1)&" CO_Matricula ="& dados_aluno(0)&"<BR>")
		
					if tipo_calculo="anual" then
						resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","anual")
						media_anual = formatNumber(media_anual,0)
						resultado_materia=resultado					
					elseif tipo_calculo="recuperacao" then
						if rec_lancado="nao" or media_rec="" or isnull(media_rec) then
						verifica=1
							resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","recuperacao")
							resultado_materia=resultado
						else
						verifica=2
							resultado=regra_aprovacao(curso,co_etapa,media_anual,media_rec,"&nbsp;","&nbsp;","&nbsp;","recuperacao")					
							resultado_materia=resultado
						end if	
'if 	dados_aluno(0)=	31323 then	
'response.Write(verifica&"-"&media_rec&"-"&resultado_materia&"<BR>")					
'end if
					elseif tipo_calculo="final" then
						if rec_lancado="nao" or media_rec="" or isnull(media_rec) then
							resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","final")
							resultado_recuperacao= split(resultado,"#!#")
							if resultado_recuperacao(1)="APR" or resultado_recuperacao(1)="REP" then
								resultado_materia=resultado
							else
							resultado_materia="&nbsp;#!#&nbsp;"								
							end if
						elseif final_lancado="nao" or media_final="" or isnull(media_final) then
							resultado=regra_aprovacao(curso,co_etapa,media_anual,media_rec,"&nbsp;","&nbsp;","&nbsp;","final")
							resultado_recuperacao= split(resultado,"#!#")
							if resultado_recuperacao(1)="APR" or resultado_recuperacao(1)="REP" then
								resultado_materia=resultado
							else	
								resultado_materia="&nbsp;#!#&nbsp;"								
							end if							
						else
						'verifica=4
							resultado=regra_aprovacao(curso,co_etapa,media_anual,media_rec,"&nbsp;",media_final,"&nbsp;","final")					
							resultado_materia=resultado
						end if						
					end if	
				else
						resultado_materia="&nbsp;#!#&nbsp;"
				end if	
										
			elseif (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) then
				vetor_mae_filhas=""
'response.Write("<BR>"&dados_aluno(1)&" CO_Matricula ="& dados_aluno(0)&"&"&co_materia(c) &"<BR>")		
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
					
					Set RSPESO = Server.CreateObject("ADODB.Recordset")
					SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&periodo
					RSPESO.Open SQLPESO, CON0
					
					peso_periodo=RSPESO("NU_Peso")
'response.write(periodo&"-"&qtd_periodos&"-"&periodo_m2&"<BR>")
					co_materia_fil_check=0
					for j=0 to ubound(co_materia_mae_fil)			
					co_materia_fil_check=co_materia_fil_check+1
						Set RS3 = Server.CreateObject("ADODB.Recordset")
						SQL3 = "SELECT * FROM "&tb_nota&" where CO_Matricula ="& dados_aluno(0) &" AND CO_Materia_Principal ='"& co_materia(c)  &"' AND CO_Materia ='"& co_materia_mae_fil(j)&"' And NU_Periodo="&periodo
						'response.write(SQL3&"<BR>")
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
						'response.write("MD "&media_aluno&"<BR>")
							if periodo=periodo_m2 then
								if media_aluno="" or isnull(media_aluno) then
									media_rec_acumula=media_rec_acumula
								else
									media_rec_acumula=media_rec_acumula+media_aluno
								end if 														
							elseif periodo=periodo_m3 then
								if media_aluno="" or isnull(media_aluno) then
									media_final_acumula=media_final_acumula
								else
									media_final_acumula=media_final_acumula+media_aluno
								end if 							
							else						
								if media_aluno="" or isnull(media_aluno) then
									media_mae_acumula=media_mae_acumula
								else
									media_mae_acumula=media_mae_acumula+media_aluno
								end if 
							end if
						end if	
							co_materia_fil_check=co_materia_fil_check*1
							co_materia_fil_conta=co_materia_fil_conta*1
							if co_materia_fil_check=co_materia_fil_conta then
								if media_mae_acumula=0 then
									md=""
									media_acumulada=media_acumulada				
									peso_periodo_acumulado=peso_periodo_acumulado						
								else
									md=media_mae_acumula/co_materia_fil_conta
									md=arredonda(md,"mat",0,0)		
									media_acumulada=media_acumulada+(md*peso_periodo)
									peso_periodo_acumulado=peso_periodo_acumulado+peso_periodo
									qtd_medias=qtd_medias+1	
								end if							
						end if

						'response.write(media_aluno&"-"&media_mae_acumula&"<BR>")
					NEXT					
						'response.write("MA "&media_rec_acumula&"<BR>")
	

					if media_rec_acumula=0 then
						media_rec=""
						rec_lancado="nao"						
					else
						media_rec=(media_rec_acumula*peso_periodo)/co_materia_fil_conta			
						media_rec=formatnumber(media_rec,0)	
						rec_lancado="sim"								
					end if	

					if media_final_acumula=0 then
						media_final=""
						final_lancado="nao"						
					else
						media_final=(media_final_acumula*peso_periodo)/co_materia_fil_conta			
						media_final=formatnumber(media_final,0)	
						final_lancado="sim"								
					end if	
				next

				if peso_periodo_acumulado=0 then
					peso_periodo_acumulado=1
				end if	
'						response.write(qtd_medias&" - "&medias_necessarias&"<BR>")							
				if qtd_medias>=medias_necessarias then
					media_anual=media_acumulada/peso_periodo_acumulado					
					decimo = media_anual - Int(media_anual)
'					If decimo >= 0.75 Then
'						nota_arredondada = Int(media_anual) + 1
'						media_anual=nota_arredondada
'					elseIf decimo >= 0.25 Then
'						nota_arredondada = Int(media_anual) + 0.5
'						media_anual=nota_arredondada
'					else
'						nota_arredondada = Int(media_anual)
'						media_anual=nota_arredondada											
'					End If	
					If decimo >= 0.5 Then
						nota_arredondada = Int(media_anual) + 1
						media_anual=nota_arredondada
					else
						nota_arredondada = Int(media_anual)
						media_anual=nota_arredondada											
					End If		
					media_anual = formatNumber(media_anual,0)			
					media_anual=media_anual*1						
					if media_anual>67 and media_anual<70 then
						media_anual=70
					end if		
					
					if tipo_calculo="anual" then
						resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","anual")
						media_anual = formatNumber(media_anual,0)
						resultado_materia=resultado					
					elseif tipo_calculo="recuperacao" then
						if rec_lancado="nao" or media_rec="" or isnull(media_rec) then

							resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","recuperacao")
							resultado_materia=resultado
						else

							resultado=regra_aprovacao(curso,co_etapa,media_anual,media_rec,"&nbsp;","&nbsp;","&nbsp;","recuperacao")					
							resultado_materia=resultado
						end if	
	
					elseif tipo_calculo="final" then						
'response.Write("if "& rec_lancado&"=rec_lancado or media_rec="& media_rec&" or isnull(media_rec) or final_lancado="&final_lancado&" or media_final="&media_final&" or<BR>")		
						if rec_lancado="nao" or media_rec="" or isnull(media_rec) then
							'verifica=3
							resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","final")
							resultado_recuperacao= split(resultado,"#!#")
							if resultado_recuperacao(1)="APR" or resultado_recuperacao(1)="REP" then
								resultado_materia=resultado
							else
								resultado_materia="&nbsp;#!#&nbsp;"								
							end if
						elseif final_lancado="nao" or media_final="" or isnull(media_final) then
							resultado=regra_aprovacao(curso,co_etapa,media_anual,media_rec,"&nbsp;","&nbsp;","&nbsp;","final")
							resultado_recuperacao= split(resultado,"#!#")
							if resultado_recuperacao(1)="APR" or resultado_recuperacao(1)="REP" then
								resultado_materia=resultado
							else
								resultado_materia="&nbsp;#!#&nbsp;"										
							end if							
						else
						'verifica=4
							resultado=regra_aprovacao(curso,co_etapa,media_anual,media_rec,"&nbsp;",media_final,"&nbsp;","final")					
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
					
					Set RSPESO = Server.CreateObject("ADODB.Recordset")
					SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&periodo
					RSPESO.Open SQLPESO, CON0
					
					peso_periodo=RSPESO("NU_Peso")

					co_materia_fil_check=0
					for j=0 to ubound(co_materia_mae_fil)			
					co_materia_fil_check=co_materia_fil_check+1
						Set RS3 = Server.CreateObject("ADODB.Recordset")
						SQL3 = "SELECT * FROM "&tb_nota&" where CO_Matricula ="& dados_aluno(0) &" AND CO_Materia_Principal ='"& co_materia(c)  &"' AND CO_Materia ='"& co_materia_mae_fil(j)&"' And NU_Periodo="&periodo

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

							if periodo=periodo_m2 then
								if media_aluno="" or isnull(media_aluno) then
									media_rec_acumula=media_rec_acumula
								else
									media_rec_acumula=media_rec_acumula+media_aluno
								end if 														
							elseif periodo=periodo_m3 then
								if media_aluno="" or isnull(media_aluno) then
									media_final_acumula=media_final_acumula
								else
									media_final_acumula=media_final_acumula+media_aluno
								end if 							
							else						
								if media_aluno="" or isnull(media_aluno) then
									media_mae_acumula=media_mae_acumula
								else
									media_mae_acumula=media_mae_acumula+media_aluno
								end if 
							end if
						end if	
							co_materia_fil_check=co_materia_fil_check*1
							co_materia_fil_conta=co_materia_fil_conta*1
							if co_materia_fil_check=co_materia_fil_conta then
								if media_mae_acumula=0 then
									md=""
									media_acumulada=media_acumulada				
									peso_periodo_acumulado=peso_periodo_acumulado						
								else
									md=media_mae_acumula/co_materia_fil_conta				
									media_acumulada=media_acumulada+(md*peso_periodo)
									peso_periodo_acumulado=peso_periodo_acumulado+peso_periodo
									qtd_medias=qtd_medias+1	
								end if							
						end if

					NEXT					

					if media_rec_acumula=0 then
						media_rec=""
						rec_lancado="nao"						
					else
						media_rec=(media_rec_acumula*peso_periodo)/co_materia_fil_conta			
						media_rec=formatnumber(media_rec,0)	
						rec_lancado="sim"								
					end if	

					if media_final_acumula=0 then
						media_final=""
						final_lancado="nao"						
					else
						media_final=(media_final_acumula*peso_periodo)/co_materia_fil_conta			
						media_final=formatnumber(media_final,0)	
						final_lancado="sim"								
					end if	
				next

				if peso_periodo_acumulado=0 then
					peso_periodo_acumulado=1
				end if	
							
				if qtd_medias>=medias_necessarias then
					media_anual=media_acumulada/peso_periodo_acumulado					
					decimo = media_anual - Int(media_anual)
'					If decimo >= 0.75 Then
'						nota_arredondada = Int(media_anual) + 1
'						media_anual=nota_arredondada
'					elseIf decimo >= 0.25 Then
'						nota_arredondada = Int(media_anual) + 0.5
'						media_anual=nota_arredondada
'					else
'						nota_arredondada = Int(media_anual)
'						media_anual=nota_arredondada											
'					End If		
					If decimo >= 0.5 Then
						nota_arredondada = Int(media_anual) + 1
						media_anual=nota_arredondada
					else
						nota_arredondada = Int(media_anual)
						media_anual=nota_arredondada											
					End If	
					media_anual = formatNumber(media_anual,0)			
					media_anual=media_anual*1						
					if media_anual>67 and media_anual<70 then
						media_anual=70
					end if						
					if tipo_calculo="anual" then
						resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","anual")
						media_anual = formatNumber(media_anual,0)
						resultado_materia=resultado					
					elseif tipo_calculo="recuperacao" then
						if rec_lancado="nao" or media_rec="" or isnull(media_rec) then

							resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","recuperacao")
							resultado_materia=resultado
						else

							resultado=regra_aprovacao(curso,co_etapa,media_anual,media_rec,"&nbsp;","&nbsp;","&nbsp;","recuperacao")					
							resultado_materia=resultado
						end if	
	
					elseif tipo_calculo="final" then						
'response.Write("if "& rec_lancado&"=rec_lancado or media_rec="& media_rec&" or isnull(media_rec) or final_lancado="&final_lancado&" or media_final="&media_final&" or<BR>")		
						if rec_lancado="nao" or media_rec="" or isnull(media_rec) then
							'verifica=3
							resultado=regra_aprovacao(curso,co_etapa,media_anual,"&nbsp;","&nbsp;","&nbsp;","&nbsp;","final")
							resultado_recuperacao= split(resultado,"#!#")
							if resultado_recuperacao(1)="APR" or resultado_recuperacao(1)="REP" then
								resultado_materia=resultado
							else
								resultado_materia="&nbsp;#!#&nbsp;"								
							end if
						elseif final_lancado="nao" or media_final="" or isnull(media_final) then
							resultado=regra_aprovacao(curso,co_etapa,media_anual,media_rec,"&nbsp;","&nbsp;","&nbsp;","final")
							resultado_recuperacao= split(resultado,"#!#")
							if resultado_recuperacao(1)="APR" or resultado_recuperacao(1)="REP" then
								resultado_materia=resultado
							else
								resultado_materia="&nbsp;#!#&nbsp;"										
							end if							
						else
						'verifica=4
							resultado=regra_aprovacao(curso,co_etapa,media_anual,media_rec,"&nbsp;",media_final,"&nbsp;","final")					
							resultado_materia=resultado
						end if	
					end if						
'				response.Write(verifica&"-"&resultado_materia&"<BR>")	
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
'response.Write(m1_aluno&"_"&nota_aux_m2_1&"_"&nota_aux_m2_2&"_"&nota_aux_m3_1&"_"&nota_aux_m3_2)
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
'				response.Write(nota_aux_m2_1_peso&"="&nota_aux_m2_1&"*"&peso_m2_m2)
				nota_aux_m2_1_peso=nota_aux_m2_1*peso_m2_m2
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
				if resultado2="apr" or resultado2="rep" then
					m3_aluno=m2_aluno				
				else
					if m2_aluno="&nbsp;" or nota_aux_m2_1="&nbsp;" or nota_aux_m3_1="&nbsp;" then		
						m3_aluno="&nbsp;"
						resultado="&nbsp;"			
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
		if md_aluno >= valor_apr then
			result_temp="apr"
		elseif md_aluno >= valor_dep then
			resultado="dep"
			qtd_dep=qtd_dep+1
		else
			result_temp="rep"			
		end if
	end if
Next
if 	libera_resultado="s" then
	if result_temp="apr" then
		apura_resultado_aluno=res_apr
	elseif result_temp="rep" then
		apura_resultado_aluno=res_rep
	elseif result_temp="dep" then	
		qtd_dep=qtd_dep*1
		qtd_max_dep=qtd_max_dep*1
		if qtd_dep>qtd_max_dep then
			apura_resultado_aluno=res_rep	
		else	
			apura_resultado_aluno=res_dep	
		end if
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