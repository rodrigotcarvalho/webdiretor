<%'On Error Resume Next




ano_form = request.Form("ano_importado")
ano_sistema = SESSION("ano_letivo")
SESSION("ano_letivo") = ano_form

%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/bd_grade.asp"-->
<!--#include file="../../../../inc/parametros.asp"-->
<!--#include file="../../../../inc/bd_historico.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/calculos.asp"-->
<!--#include file="../../../../inc/resultados.asp"-->
<%
Server.ScriptTimeout = 300 'valor em segundos

ano_atual = DatePart("yyyy", now)

chave=session("chave")
session("chave")=chave
nome = session("nome") 
unidade = request.Form("unidade")
curso = request.form("curso")
co_etapa = request.Form("etapa")

session("unidade_trabalho")=unidade
session("curso_trabalho")=curso
session("etapa_trabalho")=co_etapa

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2
Const TristateTrue = -1
Const TristateFalse = 0

ano = DatePart("yyyy", now)
mes = DatePart("m", now)
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 
seg = DatePart("s", now)
if mes<10 then
mes=0&mes
end if
if dia<10 then
dia=0&dia
end if
if hora<10 then
hora=0&hora
end if
if min<10 then
min=0&min
end if
if seg<10 then
seg=0&seg
end if
data = dia&mes&ano&hora&min&seg
			
'arquivo="Historico"&data&".txt"
'Set fs = CreateObject("Scripting.FileSystemObject") 'cria  
'Set d = fs.CreateTextFile(caminho_gera_mov&arquivo, False) 	

'call caminhos(ano_form, null,CAMINHO, CAMINHOa,CAMINHO_al,CAMINHO_b,CAMINHO_bl,CAMINHO_ct,CAMINHOctl,CAMINHO_g,CAMINHO_log,CAMINHO_h,CAMINHO_msg,CAMINHO_na,CAMINHO_nb,CAMINHO_nc,CAMINHO_nd,CAMINHO_ne,CAMINHO_nf,CAMINHO_nk,CAMINHO_nv,CAMINHO_nw,CAMINHO_o,CAMINHO_p,CAMINHO_pf,CAMINHO_pr,CAMINHO_wf,caminho_arquivo,CAMINHO_ctrle,caminho_gera_mov,caminho_bd,CAMINHO_wr,CAMINHO_upload,CAMINHO_t)

	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR


	Set CONG = Server.CreateObject("ADODB.Connection") 
	ABRIRG = "DBQ="& CAMINHO_g  & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONG.Open ABRIRG

	Set CONA = Server.CreateObject("ADODB.Connection") 
	ABRIRA = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONA.Open ABRIRA
	
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0		
	
'response.Write(CAMINHO_g&"<BR>"&CAMINHO_pr&"<BR>")		
	
	if turma="999990" or turma="" or isnull(turma) then
		if co_etapa="999990" or co_etapa="" or isnull(co_etapa) then
			if curso="999990" or curso="" or isnull(curso) then		
				if unidade="999990" or unidade="" or isnull(unidade) then
					response.redirect(origem&"index.asp?nvg="&nvg&"&opt=err2")
				else	
					Set RS0 = Server.CreateObject("ADODB.Recordset")
					SQL0 = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&" and CO_Curso<>'0' ORDER BY CO_Curso,CO_Etapa"
					RS0.Open SQL0, CON0
					check_motriz=1
					WHILE NOT RS0.EOF
						curso=RS0("CO_Curso")
						co_etapa=RS0("CO_Etapa")
						
						Set RS0t = Server.CreateObject("ADODB.Recordset")
						SQL0t = "SELECT * FROM TB_Turma where NU_Unidade="&unidade&" AND CO_Curso ='"&curso&"' AND CO_Etapa ='"&co_etapa&"' ORDER BY CO_Turma"
						RS0t.Open SQL0t, CON0							
						WHILE NOT RS0t.EOF								
							turma=RS0t("CO_Turma")	

							if check_motriz=1 then
								vetor_motriz=unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
							else
								vetor_motriz=vetor_motriz&"#$#"&unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
							end if
							check_motriz=check_motriz+1 
						RS0t.MOVENEXT
						WEND	
					RS0.MOVENEXT
					WEND					
					RS0.Close
					Set RS0 = Nothing	
				end if		
			else	
				Set RS0 = Server.CreateObject("ADODB.Recordset")
				SQL0 = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&" AND CO_Curso ='"&curso&"' ORDER BY CO_Etapa"
				RS0.Open SQL0, CON0
				check_motriz=1
				WHILE NOT RS0.EOF
					co_etapa=RS0("CO_Etapa")					
					Set RS0t = Server.CreateObject("ADODB.Recordset")
					SQL0t = "SELECT * FROM TB_Turma where NU_Unidade="&unidade&" AND CO_Curso ='"&curso&"' AND CO_Etapa ='"&co_etapa&"' ORDER BY CO_Turma"
					RS0t.Open SQL0t, CON0							
					WHILE NOT RS0t.EOF								
						turma=RS0t("CO_Turma")	

						if check_motriz=1 then
							vetor_motriz=unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
						else
							vetor_motriz=vetor_motriz&"#$#"&unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
						end if
						check_motriz=check_motriz+1 
					RS0t.MOVENEXT
					WEND	
				RS0.MOVENEXT
				WEND
				
				RS0.Close
				Set RS0 = Nothing					
			end if						
		else				
			Set RS0t = Server.CreateObject("ADODB.Recordset")
			SQL0t = "SELECT * FROM TB_Turma where NU_Unidade="&unidade&" AND CO_Curso ='"&curso&"' AND CO_Etapa ='"&co_etapa&"' ORDER BY CO_Turma"
			RS0t.Open SQL0t, CON0					
					
			check_motriz=1			
			
			WHILE NOT RS0t.EOF								
				turma=RS0t("CO_Turma")	

				if check_motriz=1 then
					vetor_motriz=unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
				else
					vetor_motriz=vetor_motriz&"#$#"&unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
				end if
				check_motriz=check_motriz+1 
			RS0t.MOVENEXT
			WEND	
		end if	
		RS0t.Close
		Set RS0t = Nothing	
	ELSE
		vetor_motriz=unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma				
	end if		

conjunto_dados=split(vetor_motriz,"#$#")



for i=0 to ubound(conjunto_dados)	
	dados_select=split(conjunto_dados(i),"#!#")
	unidade_t=dados_select(0)
	curso_t=dados_select(1)
	co_etapa_t=dados_select(2)
	turma_t=dados_select(3)			

	tb_nota=tabela_notas(CONG, unidade_t, curso_t, co_etapa_t, turma_t, 0, 0, 0)
	caminho_nota=tabela_nota(ano_form,unidade_t,curso_t,co_etapa_t,turma_t,"cam",0)    
'	if tb_nota = "TB_NOTA_A" then
'		caminho_nota = "e:\home\simplynetcloud\dados\"&ambiente_escola&"\BD\"&ano_form&"\Modelo_A.mdb"
'	elseif tb_nota = "TB_NOTA_B" then		
'		caminho_nota = "e:\home\simplynetcloud\dados\"&ambiente_escola&"\BD\"&ano_form&"\Modelo_B.mdb"
'	elseif tb_nota = "TB_NOTA_C" then			
'		caminho_nota = "e:\home\simplynetcloud\dados\"&ambiente_escola&"\BD\"&ano_form&"\Modelo_C.mdb"
'	elseif tb_nota = "TB_NOTA_D" then			
'		caminho_nota = "e:\home\simplynetcloud\dados\"&ambiente_escola&"\BD\"&ano_form&"\Modelo_D.mdb"
'	elseif tb_nota = "TB_NOTA_E" then			
'		caminho_nota = "e:\home\simplynetcloud\dados\"&ambiente_escola&"\BD\"&ano_form&"\Modelo_E.mdb"
'	elseif tb_nota = "TB_NOTA_F" then			
'		caminho_nota = "e:\home\simplynetcloud\dados\"&ambiente_escola&"\BD\"&ano_form&"\Modelo_F.mdb"
'	elseif tb_nota = "TB_NOTA_K" then			
'		caminho_nota = "e:\home\simplynetcloud\dados\"&ambiente_escola&"\BD\"&ano_form&"\Modelo_K.mdb"
'	elseif tb_nota = "TB_NOTA_V" then					
'		caminho_nota = "e:\home\simplynetcloud\dados\"&ambiente_escola&"\BD\"&ano_form&"\Modelo_V.mdb"
'	elseif tb_nota = "TB_NOTA_W" then							
'		caminho_nota = "e:\home\simplynetcloud\dados\"&ambiente_escola&"\BD\"&ano_form&"\Modelo_W.mdb"
'	end if	
	
	matriculas=alunos_esta_turma(CONA, ano_form, "CO_Matricula", unidade_t, curso_t, co_etapa_t, turma_t, "*", "NU_Chamada", outro)	
	
	Set RS5 = Server.CreateObject("ADODB.Recordset")
	SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa_t &"' AND CO_Curso ='"& curso_t &"' AND IN_MAE = TRUE order by NU_Ordem_Boletim "
	RS5.Open SQL5, CON0
	co_materia_check=1
	
	IF RS5.EOF Then
		vetor_materia_exibe="nulo"
	else
		while not RS5.EOF
			co_mat_fil= RS5("CO_Materia")	
			carga = RS5("NU_Aulas")						
			
			if co_materia_check=1 then
				vetor_materia=co_mat_fil&"#!#"&carga
				vetor_materia_media=co_mat_fil					
			else
				vetor_materia=vetor_materia&"#$#"&co_mat_fil&"#!#"&carga
				vetor_materia_media=vetor_materia_media&"#!#"&co_mat_fil						
			end if
			co_materia_check=co_materia_check+1			
					
		RS5.MOVENEXT
		wend	
		vetor_materia_exibe = vetor_materia
		'vetor_materia_exibe=programa_aula(vetor_materia, unidade, curso, co_etapa, "nulo")			
		'response.Write(vetor_materia_media&"<BR>")
	end if		

	co_materia_exibe=Split(vetor_materia,"#$#")		
'response.Write(tb_nota&"<BR>")	
'response.Write(ano_form&","&unidade&","&curso&","&co_etapa&","&turma&"<BR>")	
	Set CON_N = Server.CreateObject("ADODB.Connection") 
	ABRIR_N = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_N.Open ABRIR_N

	Set RSano = Server.CreateObject("ADODB.Recordset")
	SQLano = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&ano_form&"'"
	RSano.Open SQLano, CON
	
	teste_ano_form=RSano("ST_Ano_Letivo")	
	

	
	if matriculas="" or isnull(matriculas) then
	else
		vetor_matriculas = split(matriculas, "#!#")
		for vm=0 to ubound(vetor_matriculas)
			
			co_matricula = vetor_matriculas(vm)		
			
			nu_seq_hist = maiorSequencialHistorico(co_matricula, ano_form)
			
			'response.Write(ehLancamentoManual(co_matricula, ano_form, nu_seq_hist)&"<BR>")
			
			if ehLancamentoManual(co_matricula, ano_form, nu_seq_hist) = "N" then

				historico_excluido = excluiHistorico(co_matricula, ano_form, nu_seq_hist)	
		'response.Write(unidade_t&" "&curso_t&" "&co_etapa_t&" "&turma_t&" "&co_matricula&" "&caminho_nota&" "&tb_nota&" "&vetor_materia_media&"<BR>")		
'				resultados=Calc_Med_An_Fin(unidade_t, curso_t, co_etapa_t, turma_t, co_matricula, vetor_materia_media, caminho_nota, tb_nota, 6, 5, 6, "final", 0)	
'	if 	co_matricula = 31841 then		
'	response.Write(resultados&"<BR>")
'	response.End()
'	end if
'				resultado_final_aluno=apura_resultado_aluno(curso_t, co_etapa_t, resultados)	

					
				
				
				soma_faltas=0												
				carga_total = 0
				soma_total_faltas = 0
				frequencia_total = 0
				
'				Set RSF = Server.CreateObject("ADODB.Recordset")
'				SQL_N = "Select * from TB_Frequencia_Periodo WHERE CO_Matricula = "& co_matricula
'				Set RSF = CON_N.Execute(SQL_N)
'						
'				
'				if RSF.eof THEN
'					f1="&nbsp;"
'					f2="&nbsp;"
'					f3="&nbsp;"
'					f4="&nbsp;"	
'				else	
'					f1=RSF("NU_Faltas_P1")
'					f2=RSF("NU_Faltas_P2")
'					f3=RSF("NU_Faltas_P3")
'					f4=RSF("NU_Faltas_P4")		
'					
'					if isnull(f1) or f1= "" then
'					else
'						f1=f1*1
'						soma_faltas=soma_faltas*1
'						soma_faltas=soma_faltas+f1		
'					end if
'					
'					if isnull(f2) or f2= "" then
'					else
'						f2=f2*1
'						soma_faltas=soma_faltas*1
'						soma_faltas=soma_faltas+f2		
'					end if
'					
'					if isnull(f3) or f3= "" then
'					else
'						f3=f3*1
'						soma_faltas=soma_faltas*1
'						soma_faltas=soma_faltas+f3		
'					end if
'					
'					if isnull(f4) or f4= "" then
'					else
'						f4=f4*1
'						soma_faltas=soma_faltas*1
'						soma_faltas=soma_faltas+f4		
'					end if									
'				END IF				
	
					
''				soma_faltas=soma_faltas*1
''				dias_de_aula_no_ano=200
''				
''				frequencia=((dias_de_aula_no_ano-soma_faltas)/dias_de_aula_no_ano)*100
''				if frequencia<100 then
''					frequencia=arredonda(frequencia,"mat_dez",1,0)	
''				end if	
					Set RS3 = Server.CreateObject("ADODB.Recordset")
					SQL3 = "SELECT SUM(NU_Aulas) as carga FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' and IN_MAE = TRUE"
					RS3.Open SQL3, CON0
					'response.Write(SQL3)
					if RS3.EOF then
						horas_aula=0
					else
						horas_aula= RS3("carga")			
					end if	



					Set RSF = Server.CreateObject("ADODB.Recordset")
					SQL_N = "Select * from "&tb_nota&" WHERE CO_Matricula = "& co_matricula
					Set RSF = CON_N.Execute(SQL_N)
					soma_faltas=0			
					
					if RSF.eof THEN
						f1="&nbsp;"
						f2="&nbsp;"
						f3="&nbsp;"
						f4="&nbsp;"	
					else
						while not RSF.eof 
							'MAT=RSF("CO_Materia")
							f1=RSF("NU_Faltas_P1")
							f2=RSF("NU_Faltas_P2")
							f3=RSF("NU_Faltas_P3")
							'f4=RSF("NU_Faltas_P4")		
							'response.Write(MAT&": "&f1&"= f1<BR>")
							'response.Write(MAT&": "&f2&"= f2<BR>")
							'response.Write(MAT&": "&f3&"= f3<BR>")																		
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
							
	'						if isnull(f4) or f4= "" then
	'						else
	'							f4=f4*1
	'							soma_faltas=soma_faltas*1
	'							soma_faltas=soma_faltas+f4		
	'						end if		
						RSF.MOVENEXT
						WEND							
					END IF				

	'		response.Write(cod_cons&"-"&soma_faltas&"<BR>")
			soma_faltas=soma_faltas*1
			horas_aula=horas_aula*1
			if horas_aula="" or isnull(horas_aula) or horas_aula=0 then
				horas_aula=1
			end if
			
			frequencia=(horas_aula-soma_faltas)/horas_aula
			frequencia=frequencia*100
			if frequencia<100 then
				frequencia=arredonda(frequencia,"mat_dez",1,0)	
			end if					
					tp_modelo=tipo_divisao_ano(curso_t,co_etapa_t,"tp_modelo")
					tp_freq=tipo_divisao_ano(curso_t,co_etapa_t,"in_frequencia")
					periodo_m1 = Periodo_Media(tp_modelo,"MA",outro)
					periodo_m2 = Periodo_Media(tp_modelo,"RF",outro)
					periodo_m3 = Periodo_Media(tp_modelo,"MF",outro)					
					co_materia_verifica= split(vetor_materia_exibe,"#$#")
'					ano_sistema = SESSION("ano_letivo")
'					SESSION("ano_letivo") = ano_form
					for cmv=0 to ubound(co_materia_verifica)
						co_materia_res= split(co_materia_verifica(cmv),"#!#")
						compara_m3 = parametros_gerais(unidade_t, curso_t, co_etapa_t, turma, co_materia_res(0),"compara_m3",0)					
						aproxima_m3 = parametros_gerais(unidade_t, curso_t, co_etapa_t, turma, co_materia_res(0),"aproxima_m3",0)											
						resultados=Calc_Ter_Media (unidade_t, curso_t, co_etapa_t, turma, tp_modelo, tp_freq, co_matricula, co_materia_res(0), CON_N, tb_nota, periodo_m3, "ATA", compara_m3, aproxima_m3, outro)
	
		'RESPONSE.Write(co_matricula&" "&co_materia_res(0)&" "&resultados&"<BR>")	
		'response.end()
						resultados_apurados = split(resultados, "#!#" )	
						if cmv=0 then
							vetor_materia_resultado = co_materia_res(0)
							vetor_resultados_apurados=resultados_apurados(1)
							vetor_resultados_completos=resultados						
						else
							vetor_materia_resultado = vetor_materia_resultado&"#!#"&co_materia_res(0)							
							vetor_resultados_apurados=vetor_resultados_apurados&"#!#"&resultados_apurados(1)	
							vetor_resultados_completos=vetor_resultados_completos&"#$#"&resultados											
						end if	
						
						if resultados_apurados(0)="&nbsp;" or resultados_apurados(0)="" or isnull(resultados_apurados(0)) then
							apura_frequencia="n"				
						end if				
					next	
'					SESSION("ano_letivo") = ano_sistema 
						
						'resultado_final_aluno=novo2_apura_resultado_aluno(curso_t, co_etapa_t, co_matricula, vetor_materia_resultado, vetor_resultados_completos, frequencia, periodo_m1,  periodo_m2, periodo_m3, "ata", "S", "S", outro)
						
						
					resultados_apurados = split(resultados, "#%#" )	
					
					if ubound(resultados_apurados)=-1 then
						resultado_final_aluno="REP"
					else				
						resultado_final_aluno=apura_resultado_aluno(curso_t, co_etapa_t, vetor_resultados_completos)
					end if						
								
				'response.Write(vetor_materia_resultado&"--"&vetor_resultados_completos)
				Set RSs = Server.CreateObject("ADODB.Recordset")
				SQL_s ="SELECT CO_Situacao FROM TB_Matriculas where CO_Matricula ="& co_matricula&" AND NU_Ano="&ano_form
				Set RSs = CONA.Execute(SQL_s)
		
				situac=RSs("CO_Situacao")
				
				if situac="C" THEN
	            	nu_aprovado = 4	
					
				elseif situac="T" THEN 	
	            	nu_aprovado = 7		
					
				elseif situac="L" THEN 	
	            	nu_aprovado = 6														
				else			
	            	nu_aprovado = 8	
				end if
				
			

				if resultado_final_aluno="&nbsp;" or resultado_final_aluno="" or isnull(resultado_final_aluno) then	
					resultado_final_aluno = ""		
'			if 	co_matricula = 32748 then
'				response.Write("resultado_final_aluno - " &teste_ano_form)						
'				end if							
					
					if teste_ano_form="B" then
						nu_aprovado = 5
					end if		
				elseif resultado_final_aluno = "APR" or resultado_final_aluno = "Apr" or resultado_final_aluno = "APC"or resultado_final_aluno = "Apc" then
					resultado_final_aluno = "A"	
					nu_aprovado = 1
				elseif resultado_final_aluno = "PFI" or resultado_final_aluno = "Pfi"  then
					resultado_final_aluno = "P"	
				elseif resultado_final_aluno = "REC" or resultado_final_aluno = "Rec"  then
					resultado_final_aluno = "E"		
				elseif resultado_final_aluno = "ECE" or resultado_final_aluno = "Ece"  then
					resultado_final_aluno = "E"	
				elseif resultado_final_aluno = "REP" or resultado_final_aluno = "Rep"  then
					resultado_final_aluno = "R"	
					nu_aprovado = 5								
				elseif resultado_final_aluno = "AP.D" or resultado_final_aluno = "AP.D"  then
					resultado_final_aluno = "D"			
					nu_aprovado = 2																																								
				end if		
'			if 	co_matricula = 32748 then
'				response.Write(nu_aprovado)						
'				response.End()
'				end if												

				Set RS0 = Server.CreateObject("ADODB.Recordset")
				SQL0 = "SELECT TB_Unidade.NO_Sede, TB_Unidade.SG_UF, TB_Municipios.NO_Municipio FROM TB_Unidade inner join TB_Municipios on TB_Municipios.CO_Municipio=TB_Unidade.CO_Municipio and TB_Municipios.SG_UF=TB_Unidade.SG_UF where NU_Unidade="&unidade_t
				RS0.Open SQL0, CON0		
				
				no_sede = RS0("NO_Sede")	
				co_uf = RS0("SG_UF")					
				no_munic = RS0("NO_Municipio")	
				
				seq_historico = maiorSequencialHistorico(co_matricula, ano_form)	
				
				dados_gravados = insereTbAnoHistorico(co_matricula, ano_form, seq_historico, curso_t, co_etapa_t, no_sede, "Brasil", no_munic, co_uf, nu_aprovado, NULL, ano_form,  "A", NULL, frequencia)								
				resultado_apurado= split(vetor_resultados_completos, "#$#" )
				for co=0 to ubound(co_materia_exibe)
				'response.Write("<BR>"&co_materia_exibe(co))				
					dados_materia_exibe=Split(co_materia_exibe(co),"#!#")		
					materia=dados_materia_exibe(0)
					carga_materia=dados_materia_exibe(1)	
				'response.Write("<BR>"&co_materia_exibe(co))		
					resultado_disciplina= split(resultado_apurado(co), "#!#" )					
	
					
					if resultado_disciplina(0)="&nbsp;" or resultado_disciplina(0)="" or isnull(resultado_disciplina(0)) then
						calcular_frequencia="n"		
						media = ""
					else
						media = converte_conceito(unidade_t, curso_t, co_etapa_t, turma_t, 6, materia, resultado_disciplina(0)	, outro)
						'if isnumeric(media) then
						'	media=media/10
						'end if							
					end if	
					nu_aprovado_disciplina = 0
					
					if resultado_disciplina(1)="&nbsp;" or resultado_disciplina(1)="" or isnull(resultado_disciplina(1)) then
						calcular_frequencia="n"		
						resultado_materia = ""	
					elseif resultado_disciplina(1) = "APR" or resultado_disciplina(1) = "Apr" or resultado_disciplina(1) = "APC"or resultado_disciplina(1) = "Apc" then
						resultado_materia = "A"	
						nu_aprovado_disciplina = 1						
					elseif resultado_disciplina(1) = "PFI" or resultado_disciplina(1) = "Pfi"  then
						resultado_materia = "P"	
					elseif resultado_disciplina(1) = "REC" or resultado_disciplina(1) = "Rec"  then
						resultado_materia = "E"		
					elseif resultado_disciplina(1) = "ECE" or resultado_disciplina(1) = "ECE"  then
						resultado_materia = "E"	
					elseif resultado_disciplina(1) = "REP" or resultado_disciplina(1) = "Rep"  then
						resultado_materia = "R"				
					elseif resultado_disciplina(1) = "AP.D" or resultado_disciplina(1) = "AP.D"  then
						resultado_materia = "D"																																					
					end if			
							
					conceito=media				
					nome_materia=GeraNomesNovaVersao("D",materia,variavel2,variavel3,variavel4,variavel5,CON0,outro)	
					'd.writeLine ano_form&";"&co_matricula&";"&nome_materia&";"&carga_materia&";"&frequencia&";"&conceito&";"&resultado_materia&";"&resultado_final_aluno
					
					nota_inserida = insereTbHistoricoNota(co_matricula, ano_form, seq_historico, nome_materia, carga_materia, frequencia, conceito, nu_aprovado_disciplina)	
					
				NEXT		
										
					
				frequencia_final = frequencia	
					
				'd.writeLine ano_form&";"&co_matricula&";FREQUENCIA;"&frequencia_final&";;;;"									
			End if
		NEXT
	End if	
NEXT
'			response.Write("<BR>"&dados_arquivo&"<BR>")
'			response.End()

'd.close

SESSION("ano_letivo") = ano_sistema

response.Redirect("index.asp?nvg=WS-CA-HE-INR&opt=ok")
'response.Redirect("download.asp?opt="&arquivo)
		
		%>
<%If Err.number<>0 then
errnumb = Err.number
errdesc = Err.Description
lsPath = Request.ServerVariables("SCRIPT_NAME")
arPath = Split(lsPath, "/")
GetFileName =arPath(UBound(arPath,1))
passos = 0
for way=0 to UBound(arPath,1)
passos=passos+1
next
seleciona1=passos-2
pasta=arPath(seleciona1)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>