<!--#include file="parametros.asp"-->
<!--#include file="funcoes6.asp"-->
<!--#include file="funcoes7.asp"-->
<!--#include file="calculos.asp"-->
<!--#include file="resultados.asp"-->

<%
Function boletim_escolar (unidade,curso,co_etapa,turma,caminho_nota,tb_nota,cod_cons,tipo_boletim)
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CON_AL = Server.CreateObject("ADODB.Connection") 
	ABRIR_AL = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_AL.Open ABRIR_AL
	
	Set CON_WF = Server.CreateObject("ADODB.Connection") 
	ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_WF.Open ABRIR_WF	
	
	Set CONn = Server.CreateObject("ADODB.Connection") 
	ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONn.Open ABRIRn	
	
	Set CONa = Server.CreateObject("ADODB.Connection") 
	ABRIRa = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONa.Open ABRIRa		


	Set RSapr = Server.CreateObject("ADODB.Recordset")
	SQLapr = "Select * from TB_Regras_Aprovacao WHERE CO_Curso = '"& curso &"' AND CO_Etapa='"&co_etapa&"'"
	Set RSapr = CON0.Execute(SQLapr)
	
	if RSapr.EOF then
		ntvml=0
	else
		ntvml= RSapr("NU_Valor_M1")
	end if

	if tipo_boletim="MRD" or tipo_boletim="MRDI" then
		nome_foco="<div align=""left"">&nbsp;&nbsp;Nome</div>"
		vetor_materia=cod_cons

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
	else			
		nome_foco="Disciplinas"
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND IN_MAE=TRUE order by NU_Ordem_Boletim "
		RS.Open SQL, CON0

		co_materia_check=1
		IF RS.EOF Then
			vetor_materia_exibe="nulo"
		else
			while not RS.EOF
				co_mat_fil= RS("CO_Materia")	
				Set RSt = Server.CreateObject("ADODB.Recordset")
				SQLt = "SELECT * FROM TB_Materia where CO_Materia ='"& co_mat_fil &"' order by NU_Ordem_Boletim"
				RSt.Open SQLt, CON0
				
				if RSt.EOF then
					response.Write("Matéria não existe em TB_Materia")
					response.end()
				else
				
				nao_dependencia=RSt("IN_Obrigatorio")
					if nao_dependencia=True then
						if co_materia_check=1 then
							vetor_materia=co_mat_fil
						else
							vetor_materia=vetor_materia&"#!#"&co_mat_fil
						end if
						co_materia_check=co_materia_check+1	
					else
						alguma_media=9999
						testa_periodo=1
						
						while alguma_media=9999 and testa_periodo<7

		
							Set RS1 = Server.CreateObject("ADODB.Recordset")
							SQL1 = "SELECT * FROM "&tb_nota&" where CO_Matricula ="& cod_cons &" AND CO_Materia ='"& co_mat_fil&"' And NU_Periodo="&testa_periodo
							RS1.Open SQL1, CONn
			
							if RS1.EOF then		
								alguma_media=9999
							else
								valor=RS1("VA_Media3")
								if valor="" or isnull(valor) then
									alguma_media=9999
								else
									alguma_media=valor
								end if	
							end if	
						testa_periodo=testa_periodo+1
						wend	
	
						alguma_media=alguma_media*1
						if alguma_media=9999 then
						
						else
							if co_materia_check=1 then
								vetor_materia=co_mat_fil
							else
								vetor_materia=vetor_materia&"#!#"&co_mat_fil
							end if
							co_materia_check=co_materia_check+1		
						end if								
					end if
				end if
						
			RS.MOVENEXT
			wend						
		end if
		'---------------------------------------------
		'O trecho abaixo estava fora desse if
		'---------------------------------------------		
		if vetor_materia_exibe="nulo" then
			Response.Write("Erro 1 - Não foram encontradas matérias para Etapa ='"& co_etapa &"' e Curso ="& curso)
		else
			vetor_materia_exibe=programa_aula(vetor_materia, unidade, curso, co_etapa, turma)
		end if
	
	
	
	
		co_materia= split(vetor_materia,"#!#")	
		co_materia_check=1		
		'---------------------------------------------			
	end if		

	

if tipo_boletim="WA" then
	width_tabela=1000
	width_disciplina=226
	class_tit="tb_tit"
	class_subtit="tb_subtit"
	showt1="s"
	showt2="s"	
	showt3="s"	
	showt4="s"	
	showt5="s"	
	showt6="s"	
	cor_nota="#FF0000"
	bloqueia_notas="n"
elseif tipo_boletim="WAI" then
	width_tabela=1000
	width_disciplina=226
	class_tit="tabelaTit"
	class_subtit="tabelaTit"
	showt1="s"
	showt2="s"	
	showt3="s"	
	showt4="s"	
	showt5="s"	
	showt6="s"		
	cor_nota="#000000"	
	bloqueia_notas="n"	
elseif tipo_boletim="MRD" then
	width_tabela=990
	width_disciplina=226
	width_nu_chamada=30
	class_tit="tb_tit"
	class_subtit="tb_subtit"
	showt1="s"
	showt2="s"	
	showt3="s"	
	showt4="s"	
	showt5="s"	
	showt6="s"	
	cor_nota="#FF0000"	
	bloqueia_notas="n"	
elseif tipo_boletim="MRDI" then
	width_tabela=990
	width_disciplina=226
	width_nu_chamada=30
	class_tit="tabelaTit"
	class_subtit="tabelaTit"
	showt1="s"
	showt2="s"	
	showt3="s"	
	showt4="s"	
	showt5="s"	
	showt6="s"		
	cor_nota="#000000"	
	bloqueia_notas="n"	
elseif tipo_boletim="WF" then
	width_tabela=820
	width_disciplina=180
	class_tit="tb_tit"
	class_subtit="tb_subtit"
	cor_nota="#FF0000"	
	bloqueia_notas="s"
elseif tipo_boletim="WFI" then	
	width_tabela=990
	width_disciplina=210
	class_tit="tabelaTit"
	class_subtit="tabelaTit"
	cor_nota="#000000"	
	bloqueia_notas="s"	
end if

	tipo_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")
	tp_freq=tipo_divisao_ano(curso,co_etapa,"in_frequencia")
	
	if tipo_boletim="MRD" or tipo_boletim="MRDI" then
		cols_tit_vetor=dados_boletim(tipo_modelo,tp_freq,1,"mrd_tit",tb_nota)	
		cols_tit_rowspan_vetor=dados_boletim(tipo_modelo,tp_freq,1,"mrd_rowspan",tb_nota)
		cols_tit_colspan_vetor=dados_boletim(tipo_modelo,tp_freq,1,"mrd_colspan",tb_nota)
		cols_nome_notas_vetor=dados_boletim(tipo_modelo,tp_freq,2,"tit",tb_nota)	
		cols_notas_periodo_vetor=dados_boletim(tipo_modelo,tp_freq,2,"periodo_ref",tb_nota)
		cols_notas_rowspan_vetor=dados_boletim(tipo_modelo,tp_freq,2,"rowspan",tb_nota)
		cols_notas_colspan_vetor=dados_boletim(tipo_modelo,tp_freq,2,"colspan",tb_nota)
		cols_notas_calc_vetor=dados_boletim(tipo_modelo,tp_freq,2,"tipo_calc",tb_nota)	
	else
		cols_tit_vetor=dados_boletim(tipo_modelo,tp_freq,1,"tit",tb_nota)	
		cols_tit_rowspan_vetor=dados_boletim(tipo_modelo,tp_freq,1,"rowspan",tb_nota)
		cols_tit_colspan_vetor=dados_boletim(tipo_modelo,tp_freq,1,"colspan",tb_nota)
		cols_nome_notas_vetor=dados_boletim(tipo_modelo,tp_freq,2,"tit",tb_nota)	
		cols_notas_periodo_vetor=dados_boletim(tipo_modelo,tp_freq,2,"periodo_ref",tb_nota)
		cols_notas_rowspan_vetor=dados_boletim(tipo_modelo,tp_freq,2,"rowspan",tb_nota)
		cols_notas_colspan_vetor=dados_boletim(tipo_modelo,tp_freq,2,"colspan",tb_nota)
		cols_notas_calc_vetor=dados_boletim(tipo_modelo,tp_freq,2,"tipo_calc",tb_nota)
	end if

	colunas_tit=split(cols_tit_vetor,"#!#")
	colunas_tit_rowspan=split(cols_tit_rowspan_vetor,"#!#")
	colunas_tit_colspan=split(cols_tit_colspan_vetor,"#!#")
	colunas_nome_notas=split(cols_nome_notas_vetor,"#!#")
	colunas_notas_periodo=split(cols_notas_periodo_vetor,"#!#")
	cols_notas_rowspan=split(cols_notas_rowspan_vetor,"#!#")
	cols_notas_colspan=split(cols_notas_colspan_vetor,"#!#")
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

	qtd_colunas=0
	for qtdc=0 to ubound(colunas_tit_colspan)
		qtd_colunas=qtd_colunas+colunas_tit_colspan(qtdc)
	Next

	if tipo_boletim="MRD" or tipo_boletim="MRDI" then
	  	curso=curso*1
		teste_co_etapa=ISNUMERIC(co_etapa)
		if teste_co_etapa=TRUE then
			co_etapa=co_etapa*1
		end if
		IF curso=1 AND co_etapa>5 THEN	
			width_else=(width_tabela-width_disciplina-width_nu_chamada)/qtd_colunas
		else
	  		width_else=(width_tabela-width_disciplina)/(qtd_colunas-4)
		end if		
	else
	  	curso=curso*1
		teste_co_etapa=ISNUMERIC(co_etapa)
		if teste_co_etapa=TRUE then
			co_etapa=co_etapa*1
		end if
		IF curso=1 AND co_etapa>5 THEN
	  		width_else=(width_tabela-width_disciplina)/qtd_colunas
		else
	  		width_else=(width_tabela-width_disciplina)/(qtd_colunas-4)
		end if
	end if%>	

<table width="<%response.Write(width_tabela)%>" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
<%for col=0 to ubound(colunas_tit)
	IF tipo_boletim="WF" OR tipo_boletim="WFI" THEN
		IF col=0 then
			width_ln1="width="""&width_disciplina&""""
		else
			width_ln1=""
		end if
	else			
		IF col=0 then
			width_ln1="width="""&width_nu_chamada&""""
		elseIF col=1 then
			width_ln1="width="""&width_disciplina&""""
		else
			width_ln1=""
		end if
	end if
%>
  		  <td colspan="<%response.Write(colunas_tit_colspan(col))%>" rowspan="<%response.Write(colunas_tit_rowspan(col))%>" align="center"  class="<%response.Write(class_tit)%>" <%response.Write(width_ln1)%>><%response.Write(colunas_tit(col))%></td>
<%next%>
</tr>
<tr>
<%for colnn=0 to ubound(colunas_nome_notas)
	IF tipo_boletim="WA" or tipo_boletim="WAI" or tipo_boletim="WF" OR tipo_boletim="WFI" THEN
		IF col=0 then
			width_ln2="width="""&width_disciplina&""""
		else
			width_ln2="width="""&width_else&""""
		end if
	else			
		IF col=0 then
			width_ln2="width="""&width_nu_chamada&""""
		elseIF col=1 then
			width_ln2="width="""&width_disciplina&""""
		else
			width_ln2="width="""&width_else&""""
		end if
	end if

%>
  		  <td colspan="<%response.Write(cols_notas_colspan(col))%>" rowspan="<%response.Write(cols_notas_colspan(col))%>" align="center"  class="<%response.Write(class_tit)%>" <%response.Write(width_ln2)%>><%response.Write(colunas_nome_notas(colnn))%></td>
<%next%>
</tr>

<%		
	if tipo_boletim="MRD" or tipo_boletim="MRDI" then
		dados_alunos = split(alunos_vetor, "#$#")
		wrk_co_materia=co_materia
		tp_materia=tipo_materia(wrk_co_materia, curso, co_etapa)	
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Materia WHERE CO_Materia='"&wrk_co_materia&"'"
		RS1.Open SQL1, CON0

		no_materia=RS1("NO_Materia")			
		For a=0 to ubound(dados_alunos)
			inf_aluno = split(dados_alunos(a), "#!#")	
			wrk_cod_cons= inf_aluno(0)				
			nu_chamada_aluno = inf_aluno(1)
			nome_aluno = inf_aluno(2)

	
			if tipo_boletim="MRD" then
				if check mod 2 =0 then
					cor = "tb_fundo_linha_par" 
					cor2 = "tb_fundo_linha_impar" 				
				else 
					cor ="tb_fundo_linha_impar"
					cor2 = "tb_fundo_linha_par" 
				end if
			elseif tipo_boletim="MRDI" then
					cor ="tabela"
					cor2 = "tabela" 
			end if
												
	%>
	  <tr>
		<td width="<%response.Write(width_nu_chamada)%>" class="<%response.Write(cor)%>"><%response.Write(nu_chamada_aluno)%></td>
		<td width="<%response.Write(width_disciplina)%>" class="<%response.Write(cor)%>"><%response.Write(nome_aluno)%></td>
	<%for notas=0 to ubound(colunas_notas_periodo)
		var_bd=var_bd_periodo(tipo_modelo,tp_freq,tb_nota,colunas_notas_periodo(notas),colunas_notas_calc(notas))
		if colunas_notas_calc(notas)= "BDM" or colunas_notas_calc(notas)= "BDR"or colunas_notas_calc(notas)= "RF" or colunas_notas_calc(notas)= "BDF" then
			if tp_materia="T_F_F_N" or tp_materia="F_T_F_N"	 or tp_materia="F_F_T_N" then
				codigo_materia_pr=busca_materia_mae(wrk_co_materia)
				media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, wrk_cod_cons, codigo_materia_pr, wrk_co_materia, CONn, tb_nota, colunas_notas_periodo(notas), var_bd, outro)
				
			elseif tp_materia="T_T_F_N" then
			
				vetor_materia=busca_materias_filhas(wrk_co_materia)
				media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, wrk_cod_cons, wrk_co_materia, vetor_materia, CONn, tb_nota, colunas_notas_periodo(notas), var_bd, outro)		
					
			elseif tp_materia="T_F_T_N" then
				vetor_materia=busca_materias_filhas(wrk_co_materia)						
				media=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, wrk_cod_cons, wrk_co_materia, vetor_materia, CONn, tb_nota, colunas_notas_periodo(notas), var_bd, outro)	
						
			end if
			
			if colunas_notas_calc(notas)= "BDF" and isnumeric(media) then
				if media=0 then
					media="&nbsp;"
				end if	
			end if				
			
			periodo_autoriza=colunas_notas_periodo(notas)
		else
			if colunas_notas_calc(notas)= "ASTER" then
				media=Calcula_Asterisco(tipo_modelo, tp_freq, unidade, curso, co_etapa, turma, wrk_cod_cons, wrk_co_materia, CONn, tp_materia, tb_nota, colunas_notas_periodo(notas))		
				periodo_autoriza=colunas_notas_periodo(notas)
			elseif colunas_notas_calc(notas)= "SOMA" then
				maximo_periodo=Periodo_Media(tipo_modelo,"MA",outro)
				media=Calcula_Soma(tipo_modelo, tp_freq, unidade, curso, co_etapa, turma, wrk_cod_cons, wrk_co_materia, CONn, tp_materia, tb_nota,maximo_periodo, outro)	
				periodo_autoriza=maximo_periodo
			elseif colunas_notas_calc(notas)= "MA" then
				prd_prim_media=Periodo_Media(tipo_modelo,"MA",outro)
				primeira_media=Calc_Prim_Media (unidade, curso, co_etapa, turma, wrk_cod_cons, wrk_co_materia, CONn, tb_nota, prd_prim_media, tipo_calculo, outro)	

				inf_primeira_media=split(primeira_media,"#!#")
				media=inf_primeira_media(0)
				
				periodo_autoriza=prd_prim_media
			elseif colunas_notas_calc(notas)= "MF" then
				prd_seg_media=Periodo_Media(tipo_modelo,"MF",outro)
				segunda_media=Calc_Seg_Media (unidade, curso, co_etapa, turma, wrk_cod_cons, wrk_co_materia, CONn, tb_nota, prd_seg_media, tipo_calculo, outro)
				inf_segunda_media=split(segunda_media,"#!#")
				media=inf_segunda_media(0)
				resultado=inf_segunda_media(1)
				
				periodo_autoriza=prd_seg_media		
				periodo_res	=prd_seg_media	
			elseif colunas_notas_calc(notas)= "PF"	then
				prd_ter_media=Periodo_Media(tipo_modelo,"PF",outro)
				terceira_media=Calc_Ter_Media (unidade, curso, co_etapa, turma, wrk_cod_cons,  wrk_co_materia, CONn, tb_nota, prd_ter_media, "sem_calculo", "ficha")

				inf_terceira_media=split(terceira_media,"#!#")
				media=inf_terceira_media(0)
				resultado=inf_terceira_media(1)
				
				periodo_autoriza=prd_ter_media		
				periodo_res	=prd_ter_media							
			elseif colunas_notas_calc(notas)= "CMT" then
				media=calcula_medias(unidade, curso, co_etapa, turma, colunas_notas_periodo(notas), co_matric_alunos_turma, wrk_co_materia, caminho_nota, tb_nota, var_bd, "media_turma")	
				
				inf_cmt_media=split(media,"#$#")
				media=inf_cmt_media(0)											

			elseif colunas_notas_calc(notas)= "RES" then
				media=resultado
				periodo_autoriza=periodo_res
			else
				media="&nbsp;"			
			end if			
		end if
		teste = isnumeric(media)
			
		if bloqueia_notas="s" then
			mostra_nota=autoriza_wf(unidade, curso, co_etapa, periodo_autoriza, "M", CON_WF, outro)	
		else
			mostra_nota="s"
		end if		
	%>	
        <td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
			if teste=false then
				if mostra_nota="s" then
					if isnull(media) or media="" then
						response.Write("&nbsp;")					
					else
						response.Write(media)
					end if	
				else
					response.Write("&nbsp;")
				end if		
			else	
				if mostra_nota="s" then
					media=media*1	
					ntvml=ntvml*1
					
					if colunas_notas_calc(notas)<> "BDF" then
						media=formatnumber(media,parametros_gerais("decimais_media"))
					end if
					if media>=ntvml then	
						response.Write(media)
					else	
						response.Write("<font color="&cor_nota&">"&media&"</font>")	
					end if	
				end if	
			end if	
			
			%></div></td>
		<%next
		%>
	  </tr>
	<%	
		check=check+1
		next	
	else
		For a=0 to ubound(co_materia)
			tp_materia=tipo_materia(co_materia(a), curso, co_etapa)
	
			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT * FROM TB_Materia WHERE CO_Materia='"&co_materia(a)&"'"
			RS1.Open SQL1, CON0
				
			no_materia=RS1("NO_Materia")
	
			if tipo_boletim="WA" or tipo_boletim="WF"  then
				if check mod 2 =0 then
					cor = "tb_fundo_linha_par" 
					cor2 = "tb_fundo_linha_impar" 				
				else 
					cor ="tb_fundo_linha_impar"
					cor2 = "tb_fundo_linha_par" 
				end if
			elseif tipo_boletim="WAI" or tipo_boletim="WFI" then
					cor ="tabela"
					cor2 = "tabela" 
			end if
												
	%>
	  <tr>
		<td width="<%response.Write(width_disciplina)%>" class="<%response.Write(cor)%>"><%response.Write(no_materia)%></td>
	<%for notas=0 to ubound(colunas_notas_periodo)
		var_bd=var_bd_periodo(tipo_modelo,tp_freq,tb_nota,colunas_notas_periodo(notas),colunas_notas_calc(notas))
		
		if colunas_notas_calc(notas)= "BDM" or colunas_notas_calc(notas)= "BDR"or colunas_notas_calc(notas)= "RF" or colunas_notas_calc(notas)= "BDF" then
			if tp_materia="T_F_F_N" or tp_materia="F_T_F_N"	 or tp_materia="F_F_T_N" then
			
				codigo_materia_pr=busca_materia_mae(co_materia(a))
				media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_cons, codigo_materia_pr, co_materia(a), CONn, tb_nota, colunas_notas_periodo(notas), var_bd, outro)
				
			elseif tp_materia="T_T_F_N" then
			
				vetor_materia=busca_materias_filhas(co_materia(a))
				media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_cons, co_materia(a), vetor_materia, CONn, tb_nota, colunas_notas_periodo(notas), var_bd, outro)		
					
			elseif tp_materia="T_F_T_N" then
				vetor_materia=busca_materias_filhas(co_materia(a))						
				media=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, cod_cons, co_materia(a), vetor_materia, CONn, tb_nota, colunas_notas_periodo(notas), var_bd, outro)	
						
			end if
			
			if colunas_notas_calc(notas)= "BDF" and isnumeric(media) then
				if media=0 then
					media="&nbsp;"
				end if	
			end if				
			
			periodo_autoriza=colunas_notas_periodo(notas)
		else
			if colunas_notas_calc(notas)= "ASTER" then
				media=Calcula_Asterisco(tipo_modelo, tp_freq, unidade, curso, co_etapa, turma, cod_cons, co_materia(a), CONn, tp_materia, tb_nota, colunas_notas_periodo(notas))		
				periodo_autoriza=colunas_notas_periodo(notas)
			elseif colunas_notas_calc(notas)= "SOMA" then
				maximo_periodo=Periodo_Media(tipo_modelo,"MA",outro)
				media=Calcula_Soma(tipo_modelo, tp_freq, unidade, curso, co_etapa, turma, cod_cons, co_materia(a), CONn, tp_materia, tb_nota,maximo_periodo, outro)	
				periodo_autoriza=maximo_periodo
			elseif colunas_notas_calc(notas)= "MA" then
				prd_prim_media=Periodo_Media(tipo_modelo,"MA",outro)
				primeira_media=Calc_Prim_Media (unidade, curso, co_etapa, turma, cod_cons, co_materia(a), CONn, tb_nota, prd_prim_media, tipo_calculo, outro)	

				inf_primeira_media=split(primeira_media,"#!#")
				media=inf_primeira_media(0)
				
				periodo_autoriza=prd_prim_media
			elseif colunas_notas_calc(notas)= "MF" then
				prd_seg_media=Periodo_Media(tipo_modelo,"MF",outro)
				segunda_media=Calc_Seg_Media (unidade, curso, co_etapa, turma, cod_cons, co_materia(a), CONn, tb_nota, prd_seg_media, tipo_calculo, outro)
				inf_segunda_media=split(segunda_media,"#!#")
				media=inf_segunda_media(0)
				resultado=inf_segunda_media(1)
				
				periodo_autoriza=prd_seg_media		
				periodo_res	=prd_seg_media	
			elseif colunas_notas_calc(notas)= "PF"	then
				prd_ter_media=Periodo_Media(tipo_modelo,"PF",outro)
				terceira_media=Calc_Ter_Media (unidade, curso, co_etapa, turma, cod_cons,  co_materia(a), CONn, tb_nota, prd_ter_media, "sem_calculo", "ficha")

				inf_terceira_media=split(terceira_media,"#!#")
				media=inf_terceira_media(0)
				resultado=inf_terceira_media(1)
				
				periodo_autoriza=prd_ter_media		
				periodo_res	=prd_ter_media							
			elseif colunas_notas_calc(notas)= "CMT" then

				media=calcula_medias(unidade, curso, co_etapa, turma, colunas_notas_periodo(notas), co_matric_alunos_turma, co_materia(a), caminho_nota, tb_nota, var_bd, "media_turma")	
				
				inf_cmt_media=split(media,"#$#")
				media=inf_cmt_media(0)											

			elseif colunas_notas_calc(notas)= "RES" then
				media=resultado
				periodo_autoriza=periodo_res
			else
				media="&nbsp;"			
			end if			
		end if
		teste = isnumeric(media)
			
		if bloqueia_notas="s" then
			mostra_nota=autoriza_wf(unidade, curso, co_etapa, periodo_autoriza, "M", CON_WF, outro)	
		else
			mostra_nota="s"
		end if		
	%>	
        <td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
			if teste=false then
				if mostra_nota="s" then
					if isnull(media) or media="" then
						response.Write("&nbsp;")					
					else
						response.Write(media)
					end if	
				else
					response.Write("&nbsp;")
				end if		
			else	
				if mostra_nota="s" then
					media=media*1	
					ntvml=ntvml*1
					
					if colunas_notas_calc(notas)<> "BDF" then
						media=formatnumber(media,parametros_gerais("decimais_media"))
					end if
					if media>=ntvml then	
						response.Write(media)
					else	
						response.Write("<font color="&cor_nota&">"&media&"</font>")	
					end if	
				end if	
			end if	
			
			%></div></td>
		<%next
		%>
	  </tr>
	<%	
		check=check+1
		next
	end if	
%>		
		</table>
<%

end function	
%>
