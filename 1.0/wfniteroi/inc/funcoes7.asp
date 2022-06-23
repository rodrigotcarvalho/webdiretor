<!--#include file="../../global/funcoes_diversas.asp" -->
<!--#include file="funcoes6.asp"-->
<!--#include file="calculos.asp"-->
<!--#include file="resultados.asp"-->
<!--#include file="parametros.asp"-->
<%

Function grava_ficha(unidade, curso, co_etapa, turma, vetor_periodo_ctrl)
Server.ScriptTimeout = 900 'valor em segundos

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0

	Set CONt = Server.CreateObject("ADODB.Connection") 
	ABRIRt = "DBQ="& CAMINHO_t & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONt.Open ABRIRt
	
	Set CONa = Server.CreateObject("ADODB.Connection") 
	ABRIRa = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONa.Open ABRIRa		
	
	ano = DatePart("yyyy", now)
	mes = DatePart("m", now) 
	dia = DatePart("d", now) 
	hora = DatePart("h", now) 
	min = DatePart("n", now) 
	data = dia &"/"& mes &"/"& ano
	horario = hora & ":"& min
	
	tb_nota=tabela_nota(ano_letivo,unidade,curso,co_etapa,turma,"tb",0)
	caminho_nota=tabela_nota(ano_letivo,unidade,curso,co_etapa,turma,"cam",0)

	Set CONn = Server.CreateObject("ADODB.Connection") 
	ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONn.Open ABRIRn			

	vetor_materia=vetor_disciplinas(ano_letivo,unidade,curso,co_etapa,turma,"n",0)
'	vetor_materia_exibe=vetor_disciplinas(ano_letivo,unidade,curso,co_etapa,turma,"s",0)
	vetor_materia_exibe=vetor_materia
	
	tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")
	tp_freq=tipo_divisao_ano(curso,co_etapa,"in_frequencia")
	temp_nomes_periodos=dados_boletim(tp_modelo,tp_freq,2,"tit",tb_nota)
	vetor_num_periodo=dados_boletim(tp_modelo,tp_freq,2,"periodo_ref",tb_nota)	
	cols_notas_calc_vetor=dados_boletim(tp_modelo,tp_freq,2,"tipo_calc",tb_nota)	
		
	prd_prim_media=Periodo_Media(tp_modelo,"MA",outro)
	prd_seg_media=Periodo_Media(tp_modelo,"REC",outro)


	co_materia_exibe=Split(vetor_materia_exibe,"#!#")	
	colunas_notas_calc=split(cols_notas_calc_vetor,"#!#")
	
	
	Set RSt1 = Server.CreateObject("ADODB.Recordset")
	SQLt1 = "SELECT * FROM TB_Aluno_Esta_Turma where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'"
	RSt1.Open SQLt1, CONa
	
	co_matric_alunos_turma_check=1
	while not RSt1.EOF
	co_matricula= RSt1("CO_Matricula")
	
		if co_matric_alunos_turma_check=1 then
			co_matric_alunos_turma=co_matricula
		else
			co_matric_alunos_turma=co_matric_alunos_turma&","&co_matricula
		end if
	co_matric_alunos_turma_check=co_matric_alunos_turma_check+1
	RSt1.MOVENEXT
	wend		

		
	vetor_nomes_periodos=temp_nomes_periodos&"#!#Carga"
'response.Write(vetor_num_periodo&"-"&vetor_periodo_ctrl)
'response.End()	
	ajusta_periodos=split(vetor_nomes_periodos,"#!#")
	colunas_notas_periodo=split(vetor_num_periodo,"#!#")

	ultimo_campo_periodo=ubound(ajusta_periodos)+2

	if ubound(ajusta_periodos)<29 then
		nm=ubound(ajusta_periodos)
		while nm<30
			ReDim preserve ajusta_periodos(UBound(ajusta_periodos)+1)
			ajusta_periodos(Ubound(ajusta_periodos )) = NULL
			nm=nm+1
		wend	
	end if
	

	
	m1=ajusta_periodos(0)
	m2=ajusta_periodos(1)
	m3=ajusta_periodos(2)
	m4=ajusta_periodos(3)
	m5=ajusta_periodos(4)
	m6=ajusta_periodos(5)
	m7=ajusta_periodos(6)
	m8=ajusta_periodos(7)
	m9=ajusta_periodos(8)
	m10=ajusta_periodos(9)
	m11=ajusta_periodos(10)
	m12=ajusta_periodos(11)
	m13=ajusta_periodos(12)
	m14=ajusta_periodos(13)
	m15=ajusta_periodos(14)
	m16=ajusta_periodos(15)
	m17=ajusta_periodos(16)
	m18=ajusta_periodos(17)
	m19=ajusta_periodos(18)
	m20=ajusta_periodos(19)
	m21=ajusta_periodos(20)
	m22=ajusta_periodos(21)
	m23=ajusta_periodos(22)
	m24=ajusta_periodos(23)
	m25=ajusta_periodos(24)
	m26=ajusta_periodos(25)
	m27=ajusta_periodos(26)
	m28=ajusta_periodos(27)
	m29=ajusta_periodos(28)
	m30=ajusta_periodos(29)	
	

'	nome_periodo=split(vetor_nom_periodos,"#!#")

	Set RS0 = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Boletim_Cabecalho where NU_Unidade="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'"
	Set RS0 = CONt.Execute(SQL)
	
	If RS0.EOF THEN	
	
		Set RS = server.createobject("adodb.recordset")		
		RS.open "TB_Boletim_Cabecalho", CONt, 2, 2 'which table do you want open
		RS.addnew

			RS("NU_Unidade") = unidade
			RS("CO_Curso") = curso
			RS("CO_Etapa") = co_etapa
			RS("CO_Turma") = turma
			RS("DA_Grav")=data				
			RS("HO_Grav")=horario
			RS("CO_01")=m1
			RS("CO_02")=m2
			RS("CO_03")=m3									
			RS("CO_04")=m4
			RS("CO_05")=m5
			RS("CO_06")=m6
			RS("CO_07")=m7
			RS("CO_08")=m8
			RS("CO_09")=m9					
			RS("CO_10")=m10
			RS("CO_11")=m11
			RS("CO_12")=m12
			RS("CO_13")=m13								
			RS("CO_14")=m14
			RS("CO_15")=m15
			RS("CO_16")=m16
			RS("CO_17")=m17
			RS("CO_18")=m18
			RS("CO_19")=m19				
			RS("CO_20")=m20	
			RS("CO_21")=m21
			RS("CO_22")=m22
			RS("CO_23")=m23						
			RS("CO_24")=m24
			RS("CO_25")=m25
			RS("CO_26")=m26
			RS("CO_27")=m27
			RS("CO_28")=m28
			RS("CO_29")=m29
			RS("CO_30")=m30			
		RS.update
		set RS=nothing
		
	else

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "DELETE * from TB_Boletim_Cabecalho WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'"
		Set RS1 = CONt.Execute(SQL1)

		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Boletim_Cabecalho", CONt, 2, 2 'which table do you want open
		RS.addnew	
			RS("NU_Unidade") = unidade
			RS("CO_Curso") = curso
			RS("CO_Etapa") = co_etapa
			RS("CO_Turma") = turma
			RS("DA_Grav")=data				
			RS("HO_Grav")=horario
			RS("CO_01")=m1
			RS("CO_02")=m2
			RS("CO_03")=m3									
			RS("CO_04")=m4
			RS("CO_05")=m5
			RS("CO_06")=m6
			RS("CO_07")=m7
			RS("CO_08")=m8
			RS("CO_09")=m9					
			RS("CO_10")=m10
			RS("CO_11")=m11
			RS("CO_12")=m12
			RS("CO_13")=m13								
			RS("CO_14")=m14
			RS("CO_15")=m15
			RS("CO_16")=m16
			RS("CO_17")=m17
			RS("CO_18")=m18
			RS("CO_19")=m19				
			RS("CO_20")=m20	
			RS("CO_21")=m21
			RS("CO_22")=m22
			RS("CO_23")=m23						
			RS("CO_24")=m24
			RS("CO_25")=m25
			RS("CO_26")=m26
			RS("CO_27")=m27
			RS("CO_28")=m28
			RS("CO_29")=m29
			RS("CO_30")=m30
		RS.update
		RS.close
		set RS=nothing		

		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "DELETE * from TB_Boletim_Notas WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'"
		Set RS2 = CONt.Execute(SQL2)		
	end if
	

	alunos_vetor=alunos_turma(ano_letivo,unidade,curso,co_etapa,turma,"nome")
	
	if alunos_vetor="nulo" then
	
	else
		n_alunos= split(alunos_vetor,"#$#")			
		for al=0 to ubound(n_alunos)
			aluno= split(n_alunos(al),"#!#")
			cod_cons=aluno(0)
			
			ordem_exibe=1
			response.Write(cod_cons&"<BR>")	
			for cme=0 to ubound(co_materia_exibe)		
				co_materia_consulta=co_materia_exibe(cme)
			'response.Write(co_materia_consulta&"  - "&cme&"<BR>")			
				if 	co_materia_consulta<>"MED" then
					no_materia_exibe=GeraNomes("D",co_materia_consulta,variavel2,variavel3,variavel4,variavel5,CON0,outro)	
					
					posicao_materia=posicao_materia_tabela(co_materia_consulta, unidade, curso, co_etapa, turma)	
					posicao_materia=posicao_materia*1	
					tp_materia=tipo_materia(co_materia_consulta, curso, co_etapa)				
					if posicao_materia=2 then
						no_materia_exibe="&nbsp;&nbsp;&nbsp;&nbsp;"&no_materia_exibe	
					end if							
					
					Set RS3 = Server.CreateObject("ADODB.Recordset")
					SQL3 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia='"&co_materia_consulta&"'"
					RS3.Open SQL3, CON0
					'response.Write(SQL3)
					if RS3.EOF then
						carga_materia=NULL
					else
						'carga_materia= RS3("NU_Aulas")		
						in_mae= RS3("IN_MAE")	
						
						if in_mae=TRUE then
							carga_materia= RS3("NU_Aulas")	
						else
							carga_materia=NULL
						end if			
					end if	
					'response.Write(carga_materia)
				else
					no_materia_exibe="&nbsp;&nbsp;&nbsp;&nbsp;-->&nbsp;M&eacute;dia"
				end if			
				
				
				Set RS3a = Server.CreateObject("ADODB.Recordset")
				SQL3a = "SELECT * FROM TB_Materia where CO_Materia ='"& co_materia_consulta &"' order by NU_Ordem_Boletim"
				RS3a.Open SQL3a, CON0	
					if RS3a.EOF then
						disc_obrigat="s"
					else
						ind_obr= RS3a("IN_Obrigatorio")	
						
						if ind_obr=TRUE then
							disc_obrigat="s"
						else
							disc_obrigat="n"
						end if			
					end if	
					tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")
					tp_freq=tipo_divisao_ano(curso,co_etapa,"in_frequencia") 					
'					response.Write(co_materia_consulta&"<BR>")		
'					response.Write(disc_obrigat&"<BR>")												
					'response.Write(medias_materia(cme)&"<BR>")	
					for cln_notas=0 to ubound(colunas_notas_periodo)
					'response.Write(colunas_notas_periodo(cln_notas)&"<BR>")
						var_bd=var_bd_periodo(tp_modelo,tp_freq,tb_nota,colunas_notas_periodo(cln_notas),colunas_notas_calc(cln_notas))
'					response.Write(">>>"&colunas_notas_periodo(cln_notas)&"<BR>")
						if colunas_notas_calc(cln_notas)= "BDM" or colunas_notas_calc(cln_notas)= "BDR"or colunas_notas_calc(cln_notas)= "RF" or colunas_notas_calc(cln_notas)= "BDF" then
							if tp_materia="T_F_F_N" or tp_materia="F_T_F_N"	 or tp_materia="F_F_T_N" then				
								codigo_materia_pr=busca_materia_mae(co_materia_consulta)
								media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_cons, codigo_materia_pr, co_materia_consulta, CONn, tb_nota, colunas_notas_periodo(cln_notas), var_bd, outro)

							elseif tp_materia="T_T_F_N" then
							
								vetor_materia=busca_materias_filhas(co_materia_consulta)
								media=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_cons,co_materia_consulta, vetor_materia, CONn, tb_nota, colunas_notas_periodo(cln_notas), var_bd, outro)		

							elseif tp_materia="T_F_T_N" then
								vetor_materia=busca_materias_filhas(co_materia_consulta)						
								media=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_cons, co_materia_consulta, vetor_materia, CONn, tb_nota, colunas_notas_periodo(cln_notas), var_bd, outro)	
		
							end if
							
							if colunas_notas_calc(cln_notas)= "BDF" and isnumeric(media) then
								if media=0 then
									media="&nbsp;"
								else
									media=formatnumber(media,0)									
								end if	
							end if	
								
						else
							if colunas_notas_calc(cln_notas)= "ASTER" then
								media=Calcula_Asterisco(tp_modelo, tp_freq, unidade, curso, co_etapa, turma, cod_cons, co_materia_consulta, CONn, tp_materia, tb_nota, colunas_notas_periodo(cln_notas))		
							
							elseif colunas_notas_calc(cln_notas)= "SOMA" then
								maximo_periodo=Periodo_Media(tp_modelo,"MA",outro)
								media=Calcula_Soma(tp_modelo, tp_freq, unidade, curso, co_etapa, turma, cod_cons, co_materia_consulta, CONn, tp_materia, tb_nota,maximo_periodo, outro)		
				
							elseif colunas_notas_calc(cln_notas)= "MA" then
								prd_prim_media=Periodo_Media(tp_modelo,"MA",outro)
								primeira_media=Calc_Prim_Media (unidade, curso, co_etapa, turma, cod_cons, co_materia_consulta, caminho_nota, tb_nota, prd_prim_media, tipo_calculo, outro)
				
								inf_primeira_media=split(primeira_media,"#!#")
								media=inf_primeira_media(0)
								resultado=inf_primeira_media(1)				
							elseif colunas_notas_calc(cln_notas)= "MF" then
								prd_seg_media=Periodo_Media(tp_modelo,"REC",outro)
								segunda_media=Calc_Seg_Media (unidade, curso, co_etapa, turma, cod_cons, co_materia_consulta, caminho_nota, tb_nota, prd_seg_media, tipo_calculo, outro)

								inf_segunda_media=split(segunda_media,"#!#")
								media=inf_segunda_media(0)
								resultado=inf_segunda_media(1)
								
							elseif colunas_notas_calc(cln_notas)= "PF"	then
								prd_ter_media=Periodo_Media(tp_modelo,"MF",outro)
								terceira_media=Calc_Ter_Media (unidade, curso, co_etapa, turma, cod_cons,  co_materia_consulta, caminho_nota, tb_nota, prd_ter_media, "sem_calculo", "ficha")
				
								inf_terceira_media=split(terceira_media,"#!#")
								media=inf_terceira_media(0)
								resultado=inf_terceira_media(1)
								
								periodo_autoriza=prd_ter_media		
								periodo_res	=prd_ter_media									
								
							elseif colunas_notas_calc(cln_notas)= "CMT" then
				
								media=calcula_medias(unidade, curso, co_etapa, turma, colunas_notas_periodo(cln_notas), co_matric_alunos_turma, co_materia_consulta, caminho_nota, tb_nota, var_bd, "media_turma")	
								
								inf_cmt_media=split(media,"#$#")
								media=inf_cmt_media(0)				
												
							elseif colunas_notas_calc(cln_notas)= "RES" then
								'media=resultado
								media="&nbsp;"
							else
								media="&nbsp;"			
							end if			
						end if
						if cln_notas=0 then
							medias_materia=media
						else
							medias_materia=medias_materia&"#!#"&media
						end if	
					next		

				grava_notas = split(medias_materia,"#!#")	
				teste_grava_materia="n"
				if disc_obrigat="n" then
					for tstnts=0 to	ubound(grava_notas)
						IF ISNULL(grava_notas(tstnts)) and teste_grava_materia<>"s" THEN
							teste_grava_materia="n"							
						ELSEif grava_notas(tstnts)="&nbsp;" and teste_grava_materia<>"s" THEN
							teste_grava_materia="n"		
						else
							teste_grava_materia="s"												
						END IF				
					next			
				end if
'					response.Write(teste_grava_materia&"<BR>")							
				if disc_obrigat="s" or teste_grava_materia="s" then
					
					Set RS4 = server.createobject("adodb.recordset")			
					RS4.open "TB_Boletim_Notas", CONt, 2, 2 'which table do you want open
					RS4.addnew	
						RS4("NU_Unidade") = unidade
						RS4("CO_Curso") = curso
						RS4("CO_Etapa") = co_etapa
						RS4("CO_Turma") = turma
						RS4("CO_Matricula")= cod_cons					
						RS4("NU_Seq")=ordem_exibe
						'response.Write(no_materia_exibe&"<BR>")
						RS4("CO_01")=no_materia_exibe
						conta_disciplina="n"
					
						for tn=0 to ubound(grava_notas)			
							n_campo=tn+2
							if n_campo<10 then
								campo_gravacao="CO_0"&n_campo
							else
								campo_gravacao="CO_"&n_campo						
							end if				
							if grava_notas(tn) ="&nbsp;" or isnull(grava_notas(tn)) or grava_notas(tn)="" then
								grava=NULL
							else
								conta_disciplina="s"
								grava=grava_notas(tn)
							end if
							'response.Write(campo_gravacao&"='"&grava_notas(tn)&"'<BR>")	
							RS4(campo_gravacao)=grava
						next

						if ultimo_campo_periodo<10 then
							campo_gravacao="CO_0"&ultimo_campo_periodo
						else
							campo_gravacao="CO_"&ultimo_campo_periodo						
						end if		
							
						if no_materia_exibe="&nbsp;&nbsp;&nbsp;&nbsp;-->&nbsp;M&eacute;dia" or conta_disciplina="n" then
							RS4(campo_gravacao)=NULL
						else		
							RS4(campo_gravacao)=carga_materia	
						end if	
					RS4.update
					RS4.Close
					Set RS4 = Nothing
					ordem_exibe=ordem_exibe*1		
					ordem_exibe=ordem_exibe+1	
				end if		
			next
		next	
	end if			
'response.end()	
grava_ficha="ok"

end function

Function alunos_turma(ano_letivo,unidade,curso,co_etapa,turma,outro)

Server.ScriptTimeout = 900

	Set CON_AL = Server.CreateObject("ADODB.Connection") 
	ABRIR_AL = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_AL.Open ABRIR_AL

	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL_A = "Select * from TB_Matriculas WHERE NU_Ano="&ano_letivo&" AND NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
	Set RS = CON_AL.Execute(SQL_A)

	IF RS.EOF Then
		alunos_vetor="nulo"
	else		
		co_aluno_check=0
		While Not RS.EOF
		nu_matricula = RS("CO_Matricula")
		nu_chamada = RS("NU_Chamada")		
		
			Set RSs = Server.CreateObject("ADODB.Recordset")
			SQL_s ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& nu_matricula&" and TB_Matriculas.NU_Ano="&ano_letivo
			Set RSs = CON_AL.Execute(SQL_s)
	
			situac=RSs("CO_Situacao")
			nome_aluno=RSs("NO_Aluno")		
	
			if situac<>"C" then
				nome_aluno=nome_aluno&" - Aluno Inativo"
			end if

			if co_aluno_check=0 then
				alunos_vetor=nu_matricula&"#!#"&nu_chamada&"#!#"&nome_aluno
			else
				alunos_vetor=alunos_vetor&"#$#"&nu_matricula&"#!#"&nu_chamada&"#!#"&nome_aluno
			end if
			co_aluno_check=co_aluno_check+1	
		RS.MOVENEXT
		WEND
	END IF	
	
alunos_turma=alunos_vetor

end function

function vetor_disciplinas(ano_letivo,unidade,curso,co_etapa,turma,exibe,outro)

Server.ScriptTimeout = 900

	Set CON0= Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0

	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND IN_MAE=TRUE order by NU_Ordem_Boletim "
	RS.Open SQL, CON0
	co_materia_check=1
	IF RS.EOF Then
		vetor_materia_exibe="nulo"
	else
		while not RS.EOF
			co_mat_fil= RS("CO_Materia")		
			if co_materia_check=1 then
				vetor_materia=co_mat_fil
			else
				vetor_materia=vetor_materia&"#!#"&co_mat_fil
			end if
			co_materia_check=co_materia_check+1			
					
		RS.MOVENEXT
		wend						
	end if

	if vetor_materia_exibe="nulo" then
		Response.Write("Erro 1 - Não foram encontradas matérias para Etapa ='"& co_etapa &"' e Curso ="& curso)
	else
		vetor_materia_exibe=programa_aula(vetor_materia, unidade, curso, co_etapa, turma)
	end if
	
	if exibe="s" then
		vetor_disciplinas=vetor_materia_exibe
	else
		vetor_disciplinas=vetor_materia
	end if			
end function	

function tabela_nota(ano_letivo,unidade,curso,co_etapa,turma,tipo,outro)

Server.ScriptTimeout = 900
	
	Set CONg = Server.CreateObject("ADODB.Connection") 
	ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONg.Open ABRIRg	

	Set RS_nota = Server.CreateObject("ADODB.Recordset")
	CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"'"
	Set RS_nota = CONg.Execute(CONEXAO)


	if RS_nota.EOF then
		tipo="erro"
	else
		tb_nota = RS_nota("TP_Nota")
		if tb_nota ="TB_NOTA_A" then
			caminho_nota = CAMINHO_na
			opcao="A"
		elseif tb_nota="TB_NOTA_B" then
			caminho_nota = CAMINHO_nb
			opcao="B"		
		elseif tb_nota ="TB_NOTA_C" then
			caminho_nota = CAMINHO_nc
			opcao="C"
		elseif tb_nota ="TB_NOTA_D" then
			caminho_nota = CAMINHO_nd
			opcao="D"			
		elseif tb_nota ="TB_NOTA_E" then
			caminho_nota = CAMINHO_ne	
			opcao="E"					
		else
			tipo="erro"
		end if	
	end if	
 	
	if tipo="tb" then
		tabela_nota=tb_nota
	elseif tipo="cam" then	
		tabela_nota=caminho_nota
	elseif tipo="opt" then	
		tabela_nota=opcao
	elseif tipo="erro" then
		tabela_nota=tipo	
	end if	
end function


Function periodos_ACC(periodo,acumulado,qto_falta,tp_modelo,id,outro)

Server.ScriptTimeout = 900

	if acumulado="s" then
		for p=1 to periodo
			if p=1 then
				temp_num_periodo=p
				sigla_periodo=periodos(p,tp_modelo,"sigla")
				temp_nomes_periodos=sigla_periodo
			else
				temp_num_periodo=temp_num_periodo&"#!#"&p
				sigla_periodo=periodos(p,tp_modelo,"sigla")
				temp_nomes_periodos=temp_nomes_periodos&"#!#"&sigla_periodo
			end if
		next
		if tp_modelo="B" then
			if qto_falta="s" then
				vetor_periodo=split(temp_nomes_periodos,"#!#")
				num_periodo=split(temp_num_periodo,"#!#")		
				for v=0 to ubound(vetor_periodo)
					if vetor_periodo(v)="BIM1" then	
						temp_num_periodo=1
						periodo_exibe=vetor_periodo(v)
					elseif vetor_periodo(v)="BIM2" then	
						temp_num_periodo=temp_num_periodo&"#!#2"
						periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)
					elseif vetor_periodo(v)="BIM3" then	
						temp_num_periodo=temp_num_periodo&"#!#3#!#0"
						periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)&"#!#QF1"
					elseif vetor_periodo(v)="BIM4" then	
						temp_num_periodo=temp_num_periodo&"#!#4#!#0#!#0"
						periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)&"#!#MA#!#QF2"					
					elseif vetor_periodo(v)="REC" then	
						temp_num_periodo=temp_num_periodo&"#!#5#!#0"
						periodo_exibe=periodo_exibe&"#!#Rec#!#MF"	
					elseif vetor_periodo(v)="FINAL" then	
						temp_num_periodo=temp_num_periodo&"#!#5"
						periodo_exibe=periodo_exibe&"#!#Pr.f"											
					end if	
				next										
			else
				vetor_periodo=split(temp_nomes_periodos,"#!#")
				num_periodo=split(temp_num_periodo,"#!#")		
				for v=0 to ubound(vetor_periodo)
					if vetor_periodo(v)="BIM1" then	
						temp_num_periodo=1
						periodo_exibe=vetor_periodo(v)
					elseif vetor_periodo(v)="BIM2" then	
						temp_num_periodo=temp_num_periodo&"#!#2"
						periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)
					elseif vetor_periodo(v)="BIM3" then	
						temp_num_periodo=temp_num_periodo&"#!#3"
						periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)
					elseif vetor_periodo(v)="BIM4" then	
						temp_num_periodo=temp_num_periodo&"#!#4#!#0"
						periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)&"#!#MA"					
					elseif vetor_periodo(v)="REC" then	
						temp_num_periodo=temp_num_periodo&"#!#5#!#0"
						periodo_exibe=periodo_exibe&"#!#Rec#!#MF"	
					elseif vetor_periodo(v)="FINAL" then	
						temp_num_periodo=temp_num_periodo&"#!#6"
						periodo_exibe=periodo_exibe&"#!#Pr.f"						
					end if					
				next					
			end if	
		else
			if qto_falta="s" then
				vetor_periodo=split(temp_nomes_periodos,"#!#")
				num_periodo=split(temp_num_periodo,"#!#")		
				for v=0 to ubound(vetor_periodo)
					if vetor_periodo(v)="TRI1" then	
						temp_num_periodo="1#!#1#!#1"
						periodo_exibe=vetor_periodo(v)&"#!#Rec.P#!#TRI1*"
					elseif vetor_periodo(v)="TRI2" then	
						temp_num_periodo=temp_num_periodo&"#!#2#!#2#!#2#!#0"
						periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)&"#!#Rec.P#!#TRI2*#!#QF1"
					elseif vetor_periodo(v)="TRI3" then	
						temp_num_periodo=temp_num_periodo&"#!#3#!#0#!#0"
						periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)&"#!#MA#!#QF2"				
					elseif vetor_periodo(v)="REC" then	
						temp_num_periodo=temp_num_periodo&"#!#4#!#0#!#0"
						periodo_exibe=periodo_exibe&"#!#Pr.Rec.F#!#MF#!#QF3"	
					elseif vetor_periodo(v)="FINAL" then	
						temp_num_periodo=temp_num_periodo&"#!#5"
						periodo_exibe=periodo_exibe&"#!#Pr.f"											
					end if	
				next									
			else
				vetor_periodo=split(temp_nomes_periodos,"#!#")
				num_periodo=split(temp_num_periodo,"#!#")		
				for v=0 to ubound(vetor_periodo)
					if vetor_periodo(v)="TRI1" then	
						temp_num_periodo="1#!#1#!#1"
						periodo_exibe=vetor_periodo(v)&"#!#Rec.P#!#TRI1*"
					elseif vetor_periodo(v)="TRI2" then	
						temp_num_periodo=temp_num_periodo&"#!#2#!#2#!#2"
						periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)&"#!#Rec.P#!#TRI2*"
					elseif vetor_periodo(v)="TRI3" then	
						temp_num_periodo=temp_num_periodo&"#!#3#!#0"
						periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)&"#!#MA"				
					elseif vetor_periodo(v)="REC" then	
						temp_num_periodo=temp_num_periodo&"#!#4#!#0"
						periodo_exibe=periodo_exibe&"#!#Pr.Rec.F#!#MF"	
					elseif vetor_periodo(v)="Final" then	
						temp_num_periodo=temp_num_periodo&"#!#5"
						periodo_exibe=periodo_exibe&"#!#Pr.f"											
					end if				
				next					
			end if	
		end if			
	else	
		temp_num_periodo=periodo
		sigla_periodo=periodos(periodo,tp_modelo,"sigla")
		periodo_exibe=sigla_periodo
	end if

	if id="num" then	
		periodos_ACC=temp_num_periodo
	elseif id="nom" then	
		periodos_ACC=periodo_exibe
	end if
end function	

Function grava_ACC(unidade, curso, co_etapa, turma, periodo, acumulado, qto_falta, nota_m1, nota_m2, nota_m3, peso_m2_m1, peso_m2_m2, peso_m3_m1, peso_m3_m2, peso_m3_m3)
Server.ScriptTimeout = 900 'valor em segundos


	ano = DatePart("yyyy", now)
	mes = DatePart("m", now) 
	dia = DatePart("d", now) 
	hora = DatePart("h", now) 
	min = DatePart("n", now) 
	data = dia &"/"& mes &"/"& ano
	horario = hora & ":"& min

	vetor_materias=vetor_disciplinas(ano_letivo,unidade,curso,co_etapa,turma,"n",0)
	
	ajusta_materias=split(vetor_materias,"#!#")
	
	if ubound(ajusta_materias)<29 then
		nm=ubound(ajusta_materias)
		while nm<30
			ReDim preserve ajusta_materias(UBound(ajusta_materias)+1)
			ajusta_materias(Ubound(ajusta_materias )) = NULL
			nm=nm+1
		wend	
	end if
	
	m1=ajusta_materias(0)
	m2=ajusta_materias(1)
	m3=ajusta_materias(2)
	m4=ajusta_materias(3)
	m5=ajusta_materias(4)
	m6=ajusta_materias(5)
	m7=ajusta_materias(6)
	m8=ajusta_materias(7)
	m9=ajusta_materias(8)
	m10=ajusta_materias(9)
	m11=ajusta_materias(10)
	m12=ajusta_materias(11)
	m13=ajusta_materias(12)
	m14=ajusta_materias(13)
	m15=ajusta_materias(14)
	m16=ajusta_materias(15)
	m17=ajusta_materias(16)
	m18=ajusta_materias(17)
	m19=ajusta_materias(18)
	m20=ajusta_materias(19)
	m21=ajusta_materias(20)
	m22=ajusta_materias(21)
	m23=ajusta_materias(22)
	m24=ajusta_materias(23)
	m25=ajusta_materias(24)
	m26=ajusta_materias(25)
	m27=ajusta_materias(26)
	m28=ajusta_materias(27)
	m29=ajusta_materias(28)
	m30=ajusta_materias(29)	
	

'	nome_periodo=split(vetor_nom_periodos,"#!#")
	
	Set CONt = Server.CreateObject("ADODB.Connection") 
	ABRIRt = "DBQ="& CAMINHO_t & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONt.Open ABRIRt
	
	tb_nota=tabela_nota(ano_letivo,unidade,curso,co_etapa,turma,"tb",0)
	caminho_nota=tabela_nota(ano_letivo,unidade,curso,co_etapa,turma,"cam",0)
	tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")		

	periodo_m1=Periodo_Media(tp_modelo,"MA",outro)
	periodo_m2=Periodo_Media(tp_modelo,"REC",outro)
	periodo_m3=Periodo_Media(tp_modelo,"MF",outro)

	Set RS0 = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Mapao_Disciplinas where NU_Unidade="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'"
	Set RS0 = CONt.Execute(SQL)
	
	If RS0.EOF THEN	
	
		Set RS = server.createobject("adodb.recordset")		
		RS.open "TB_Mapao_Disciplinas", CONt, 2, 2 'which table do you want open
		RS.addnew

			RS("NU_Unidade") = unidade
			RS("CO_Curso") = curso
			RS("CO_Etapa") = co_etapa
			RS("CO_Turma") = turma
			RS("DA_Grav")=data				
			RS("HO_Grav")=horario
			RS("CO_01")=m1
			RS("CO_02")=m2
			RS("CO_03")=m3									
			RS("CO_04")=m4
			RS("CO_05")=m5
			RS("CO_06")=m6
			RS("CO_07")=m7
			RS("CO_08")=m8
			RS("CO_09")=m9					
			RS("CO_10")=m10
			RS("CO_11")=m11
			RS("CO_12")=m12
			RS("CO_13")=m13								
			RS("CO_14")=m14
			RS("CO_15")=m15
			RS("CO_16")=m16
			RS("CO_17")=m17
			RS("CO_18")=m18
			RS("CO_19")=m19				
			RS("CO_20")=m20	
			RS("CO_21")=m21
			RS("CO_22")=m22
			RS("CO_23")=m23						
			RS("CO_24")=m24
			RS("CO_25")=m25
			RS("CO_26")=m26
			RS("CO_27")=m27
			RS("CO_28")=m28
			RS("CO_29")=m29
			RS("CO_30")=m30			
		RS.update
		set RS=nothing
		
	else

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL = "DELETE * from TB_Mapao_Disciplinas WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'"
		Set RS0 = CONt.Execute(SQL)

		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Mapao_Disciplinas", CONt, 2, 2 'which table do you want open
		RS.addnew	
			RS("NU_Unidade") = unidade
			RS("CO_Curso") = curso
			RS("CO_Etapa") = co_etapa
			RS("CO_Turma") = turma
			RS("DA_Grav")=data				
			RS("HO_Grav")=horario
			RS("CO_01")=m1
			RS("CO_02")=m2
			RS("CO_03")=m3									
			RS("CO_04")=m4
			RS("CO_05")=m5
			RS("CO_06")=m6
			RS("CO_07")=m7
			RS("CO_08")=m8
			RS("CO_09")=m9					
			RS("CO_10")=m10
			RS("CO_11")=m11
			RS("CO_12")=m12
			RS("CO_13")=m13								
			RS("CO_14")=m14
			RS("CO_15")=m15
			RS("CO_16")=m16
			RS("CO_17")=m17
			RS("CO_18")=m18
			RS("CO_19")=m19				
			RS("CO_20")=m20	
			RS("CO_21")=m21
			RS("CO_22")=m22
			RS("CO_23")=m23						
			RS("CO_24")=m24
			RS("CO_25")=m25
			RS("CO_26")=m26
			RS("CO_27")=m27
			RS("CO_28")=m28
			RS("CO_29")=m29
			RS("CO_30")=m30
		RS.update
		RS.close
		set RS=nothing		

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "DELETE * from TB_Mapao_Notas WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'"
		Set RS1 = CONt.Execute(SQL1)		
	end if
	
	if acumulado="s" then
		vetor_num_periodos=periodos_ACC(periodo,"s",qto_falta,tp_modelo,"num",0)
		vetor_nom_periodos=periodos_ACC(periodo,"s",qto_falta,tp_modelo,"nom",0)
	else
		vetor_num_periodos=periodo
		vetor_nom_periodos=periodos_ACC(periodo,"n","n",tp_modelo,"nom",0)
	end if
	num_periodo=split(vetor_num_periodos,"#!#")	
	nom_periodo=split(vetor_nom_periodos,"#!#")
	
	alunos_vetor=alunos_turma(ano_letivo,unidade,curso,co_etapa,turma,0)
	n_alunos= split(alunos_vetor,"#$#")			

		for al=0 to ubound(n_alunos)
			aluno= split(n_alunos(al),"#!#")
			cod_cons=aluno(0)
			for per=0 to ubound(nom_periodo)
				ordem_periodo=per+1
				For mat=0 to ubound(ajusta_materias)		
					if ajusta_materias(mat)="" or isnull(ajusta_materias(mat)) then
						media=""
					else	
						media=ACC(unidade, curso, co_etapa, turma, cod_cons, ajusta_materias(mat), caminho_nota, tb_nota, nom_periodo(per), num_periodo(per), periodo_m1, periodo_m2, periodo_m3, nota_m1, nota_m2, nota_m3, peso_m2_m1, peso_m2_m2, peso_m3_m1, peso_m3_m2, peso_m3_m3)	
					end if			
					if mat=0 then
						vetor_grava_notas=media
					else	
						vetor_grava_notas=vetor_grava_notas&"#!#"&media
					end if						
				next
				vetor_grava_notas=replace(vetor_grava_notas,"&nbsp;","")
				grava_notas=split(vetor_grava_notas,"#!#")				
					
				Set RS2 = server.createobject("adodb.recordset")			
				RS2.open "TB_Mapao_Notas", CONt, 2, 2 'which table do you want open
				RS2.addnew	
					RS2("NU_Unidade") = unidade
					RS2("CO_Curso") = curso
					RS2("CO_Etapa") = co_etapa
					RS2("CO_Turma") = turma
					RS2("CO_Matricula")= cod_cons				
					RS2("NU_Seq_Per")=ordem_periodo
					RS2("NU_Seq_Per_Real")=num_periodo(per)
					RS2("CO_Per")=nom_periodo(per)
					RS2("CO_01")=grava_notas(0)
					RS2("CO_02")=grava_notas(1)
					RS2("CO_03")=grava_notas(2)									
					RS2("CO_04")=grava_notas(3)
					RS2("CO_05")=grava_notas(4)
					RS2("CO_06")=grava_notas(5)
					RS2("CO_07")=grava_notas(6)
					RS2("CO_08")=grava_notas(7)
					RS2("CO_09")=grava_notas(8)					
					RS2("CO_10")=grava_notas(9)
					RS2("CO_11")=grava_notas(10)
					RS2("CO_12")=grava_notas(11)
					RS2("CO_13")=grava_notas(12)								
					RS2("CO_14")=grava_notas(13)
					RS2("CO_15")=grava_notas(14)
					RS2("CO_16")=grava_notas(15)
					RS2("CO_17")=grava_notas(16)
					RS2("CO_18")=grava_notas(17)
					RS2("CO_19")=grava_notas(18)				
					RS2("CO_20")=grava_notas(19)	
					RS2("CO_21")=grava_notas(20)
					RS2("CO_22")=grava_notas(21)
					RS2("CO_23")=grava_notas(22)						
					RS2("CO_24")=grava_notas(23)
					RS2("CO_25")=grava_notas(24)
					RS2("CO_26")=grava_notas(25)
					RS2("CO_27")=grava_notas(26)
					RS2("CO_28")=grava_notas(27)
					RS2("CO_29")=grava_notas(28)
					RS2("CO_30")=grava_notas(29)
				RS2.update
				RS2.Close
				Set RS2 = Nothing

			next
			
		next
grava_ACC="ok"

end function

Function ACC(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, no_periodo, periodo, periodo_m1, periodo_m2, periodo_m3, nota_m1, nota_m2, nota_m3, peso_m2_m1, peso_m2_m2, peso_m3_m1, peso_m3_m2, peso_m3_m3)

Server.ScriptTimeout = 900

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CONn = Server.CreateObject("ADODB.Connection") 
	ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONn.Open ABRIRn	
		
if codigo_materia="MED" then
'	if no_periodo<>"QF1" and no_periodo<>"QF2" and no_periodo<>"MB" and no_periodo<>"MF" then		
'		periodo=periodo*1
'		if periodo=1 then
'			media=Session("md_p1")
'			Session("md_p1")=""
'		elseif periodo=2 then
'			media=Session("md_p2")
'			Session("md_p2")=""
'		elseif periodo=3 then
'			media=Session("md_p3")		
'			Session("md_p3")=""
'		elseif periodo=4 then
'			media=Session("md_p4")
'			Session("md_p4")=""
'		elseif periodo=5 then
'			media=Session("md_p5")	
'			Session("md_p5")=""
'		elseif periodo=6 then
'			media=Session("md_p6")	
'			Session("md_p6")=""	
'		end if																						
'	else
'		if no_periodo="QF1" then
'			media=Session("md_qf1")
'			Session("md_qf1")=""				
'		elseif no_periodo="QF2" then	
'			media=Session("md_qf2")
'			Session("md_qf2")=""
'		elseif no_periodo="MB" then		
'			media=Session("md_mb")
'			Session("md_mb")=""
'		elseif no_periodo="MF" then
'			media=Session("md_mf")
'			Session("md_mf")=""	
'		else
'			media=""
'		end if	
'	end if	
'
else

	tp_materia=tipo_materia(codigo_materia, curso, co_etapa)
	tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")
	tp_freq=tipo_divisao_ano(curso,co_etapa,"in_frequencia") 
	

	if no_periodo="BIM1" or no_periodo="BIM2" or no_periodo="BIM3" or no_periodo="BIM4" or no_periodo="TRI1" or no_periodo="TRI2" or no_periodo="TRI3" then
		var_bd=var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo,"BDM")	
		if tp_materia="T_F_F_N" or tp_materia="F_T_F_N"	 or tp_materia="F_F_T_N" then	
			codigo_materia_pr=busca_materia_mae(codigo_materia)	
			media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia_pr, codigo_materia, CONn, tb_nota, periodo, var_bd, outro)	 
	'response.Write(media&"-"&no_periodo&","&unidade&","&curso&","&co_etapa&","&turma&","&cod_aluno&","&codigo_materia_pr&","&codigo_materia&","&CONn&","&tb_nota&","&periodo&","&var_bd&","&outro&"<BR>")
		 
		elseif tp_materia="T_T_F_N" then
		
			vetor_materia=busca_materias_filhas(codigo_materia)
			media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, vetor_materia, CONn, tb_nota, periodo, var_bd, outro)		
				
		elseif tp_materia="T_F_T_N" then
			vetor_materia=busca_materias_filhas(codigo_materia)						
			media=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, ctp_modelo, tp_freq, od_aluno, codigo_materia, vetor_materia, CONn, tb_nota, periodo, var_bd, outro)	
		
		end if
	else
		acumula_media=0
		if no_periodo="QF1" then
			periodo_qf=periodo_m1-1
'			for periodo=1 to periodo_qf
'				var_bd=var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo,"BDM")								
'				if tp_materia="T_F_F_N" or tp_materia="F_T_F_N"	 or tp_materia="F_F_T_N" then	
'					codigo_materia_pr=busca_materia_mae(codigo_materia)	
'					qf_per=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia_pr, codigo_materia, CONn, tb_nota, periodo, var_bd, outro)	  
'				elseif tp_materia="T_T_F_N" then
'				
'					vetor_materia=busca_materias_filhas(codigo_materia)
'					media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, vetor_materia, CONn, tb_nota, periodo, var_bd, outro)		
'						
'				elseif tp_materia="T_F_T_N" then
'					vetor_materia=busca_materias_filhas(codigo_materia)						
'					media=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, vetor_materia, CONn, tb_nota, periodo, var_bd, outro)		
'				end if
'				qf_ast=Calcula_Asterisco(tp_modelo, tp_freq, unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, CONn, tp_materia, tb_nota, periodo)						
'				if qf_ast="&nbsp;" then
'					qf=qf_per
'				else
'					qf=qf_ast
'				end if						
'				if qf="&nbsp;" or qf="" or isnull(qf) then
'					acumula_media=acumula_media
'				else
'					acumula_media=acumula_media+qf
'				end if	
'			next			
			acumula_media=Calcula_Soma(tp_modelo, tp_freq, unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, CONn, tp_materia, tb_nota, periodo_qf, outro)	
			if acumula_media="&nbsp;" then	
				media=""				
			else	
				if acumula_media=0 then
					media=""
				else
					media_qf=(nota_m1*periodo_m1)-acumula_media			
					if media_qf<=0 then
						media_qf="&nbsp;"
					else	
						media=arredonda(media_qf,parametros_gerais("arred_media"),parametros_gerais("decimais_media"),0)
					end if					
				end if
			end if	
		elseif no_periodo="QF2" then
			periodo_qf=periodo_m2-1
'			for periodo=1 to periodo_qf
'				var_bd=var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo,"BDM")								
'				if tp_materia="T_F_F_N" or tp_materia="F_T_F_N"	 or tp_materia="F_F_T_N" then	
'					codigo_materia_pr=busca_materia_mae(codigo_materia)	
'					qf_per=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia_pr, codigo_materia, CONn, tb_nota, periodo, var_bd, outro)	  
'				elseif tp_materia="T_T_F_N" then
'				
'					vetor_materia=busca_materias_filhas(codigo_materia)
'					media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, vetor_materia, CONn, tb_nota, periodo, var_bd, outro)		
'						
'				elseif tp_materia="T_F_T_N" then
'					vetor_materia=busca_materias_filhas(codigo_materia)						
'					media=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, vetor_materia, CONn, tb_nota, periodo, var_bd, outro)		
'				end if
'				qf_ast=Calcula_Asterisco(tp_modelo, tp_freq, unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, CONn, tp_materia, tb_nota, periodo)						
'				if qf_ast="&nbsp;" then
'					qf=qf_per
'				else
'					qf=qf_ast
'				end if						
'				if qf="&nbsp;" or qf="" or isnull(qf) then
'					acumula_media=acumula_media
'				else
'					acumula_media=acumula_media+qf
'				end if	
'			next			
			acumula_media=Calcula_Soma(tp_modelo, tp_freq, unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, CONn, tp_materia, tb_nota, periodo_qf, outro)	
			if acumula_media="&nbsp;" then	
				media=""				
			else
				if acumula_media=0 then
					media="&nbsp;"
				else
					media_qf=(nota_m2*periodo_m2)-acumula_media	
					if media_qf<=0 then
						media="&nbsp;"
					else	
						media=arredonda(media_qf,parametros_gerais("arred_media"),parametros_gerais("decimais_media"),0)
					end if					
				end if
			end if
			
		elseif no_periodo="QF3" then
			periodo_qf=periodo_m3-1
			segunda_media=Calc_Seg_Media (unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, periodo_m2, tipo_calculo, outro)
			inf_segunda_media=split(segunda_media,"#!#")
			resultado=inf_segunda_media(1)	
			
			if resultado<>"APR" and resultado<>"&nbsp;"then
				'acumula_media=Calcula_Soma(tp_modelo, tp_freq, unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, CONn, tp_materia, tb_nota, periodo_qf, outro)	
				media_qf=nota_m3*peso_m3_m3
				media=arredonda(media_qf,parametros_gerais("arred_media"),parametros_gerais("decimais_media"),0)
			else
				media="&nbsp;"	
			end if	



		elseif no_periodo="MA" then	
			primeira_media=Calc_Prim_Media (unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, periodo_m1, tipo_calculo, outro)	

			inf_primeira_media=split(primeira_media,"#!#")
			media=inf_primeira_media(0)
	
		elseif no_periodo="Rec.P" then						
			var_bd=var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo,"BDR")					
			if tp_materia="T_F_F_N" or tp_materia="F_T_F_N"	 or tp_materia="F_F_T_N" then		
				codigo_materia_pr=busca_materia_mae(codigo_materia)	
				media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia_pr, codigo_materia, CONn, tb_nota, periodo, var_bd, outro)	
			elseif tp_materia="T_T_F_N" then
			
				vetor_materia=busca_materias_filhas(codigo_materia)
				media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, vetor_materia, CONn, tb_nota, periodo, var_bd, outro)		
					
			elseif tp_materia="T_F_T_N" then
				vetor_materia=busca_materias_filhas(codigo_materia)						
				media=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, cod_aluno, codigo_materia, vetor_materia, CONn, tb_nota, periodo, var_bd, outro)					
			end if
		elseif no_periodo="BIM1*" or no_periodo="TRI1*" then		
			media=Calcula_Asterisco(tp_modelo, tp_freq, unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, CONn, tp_materia, tb_nota, periodo)	
		elseif no_periodo="BIM2*" or no_periodo="TRI2*"  then		
			media=Calcula_Asterisco(tp_modelo, tp_freq, unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, CONn, tp_materia, tb_nota, periodo)																	
		elseif no_periodo="MF" then
			segunda_media=Calc_Seg_Media (unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, periodo_m2, tipo_calculo, outro)
			inf_segunda_media=split(segunda_media,"#!#")
			media=inf_segunda_media(0)
		elseif no_periodo = "Pr.f"	then
			prd_ter_media=Periodo_Media(tp_modelo,"MF",outro)
			terceira_media=Calc_Ter_Media (unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, periodo_m3, "sem_calculo", "ficha")

			inf_terceira_media=split(terceira_media,"#!#")
			media=inf_terceira_media(0)					
		else
			media=""
		end if	
	end if
end if			
ACC=media	
end function



%>