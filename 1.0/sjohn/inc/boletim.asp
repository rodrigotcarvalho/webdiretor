<!--#include file="funcoes6.asp"-->
<%
Function boletim_escolar (unidade,curso,co_etapa,turma,caminho_nota,tb_nota,cod_cons,tipo_boletim)
ano_letivo=session("ano_letivo")

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
	
		if vetor_materia_exibe="nulo" then
			Response.Write("Erro 1 - Não foram encontradas matérias para Etapa ='"& co_etapa &"' e Curso ="& curso)
		else
			vetor_materia_exibe=programa_aula(vetor_materia, unidade, curso, co_etapa, turma)
		end if
	end if		
		co_materia= split(vetor_materia,"#!#")	
		co_materia_check=1	
		


		

if tipo_boletim="WA" then
	width_tabela=990
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
elseif tipo_boletim="WAI" then
	width_tabela=990
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
elseif tipo_boletim="WF" then
	width_tabela=820
	width_disciplina=180
	class_tit="tb_tit"
	class_subtit="tb_subtit"
	cor_nota="#FF0000"	
	
	Set RS4 = Server.CreateObject("ADODB.Recordset")
	SQL4 = "SELECT * FROM TB_Controle"
	RS4.Open SQL4, CON_WF
	
	co_apr1=RS4("CO_apr1")
	co_apr2=RS4("CO_apr2")
	co_apr3=RS4("CO_apr3")
	co_apr4=RS4("CO_apr4")
	co_apr5=RS4("CO_apr5")
	co_apr6=RS4("CO_apr6")
	co_prova1=RS4("CO_prova1")
	co_prova2=RS4("CO_prova2")
	co_prova3=RS4("CO_prova3")
	co_prova4=RS4("CO_prova4")	
	co_prova5=RS4("CO_prova5")
	co_prova6=RS4("CO_prova6")		

				
	if co_apr1="D"AND co_prova1="D" then
		showt1="n"
	else
		showt1="s"
	end if

	if co_apr2="D" AND co_prova2="D" then
		showt2="n"
	else
		showt2="s"	
	end if
			
	if co_apr3="D" AND co_prova3="D" then
		showt3="n"
	else
		showt3="s"
	end if
					
	if co_apr4="D" AND co_prova4="D" then
		showt4="n"
	else
		showt4="s"
	end if
		
	if co_apr5="D" AND co_prova5="D" then
		showt5="n"
	ELSE
		showt5="s"	
	end if
			
	if co_apr6="D" AND co_prova6="D" then
		showt6="n"
	ELSE
		showt6="s"	
	end if							
elseif tipo_boletim="WFI" then	
	width_tabela=990
	width_disciplina=210
	class_tit="tabelaTit"
	class_subtit="tabelaTit"
	cor_nota="#000000"
		
	Set RS4 = Server.CreateObject("ADODB.Recordset")
	SQL4 = "SELECT * FROM TB_Controle"
	RS4.Open SQL4, CON_WF
	
	co_apr1=RS4("CO_apr1")
	co_apr2=RS4("CO_apr2")
	co_apr3=RS4("CO_apr3")
	co_apr4=RS4("CO_apr4")
	co_apr5=RS4("CO_apr5")
	co_apr6=RS4("CO_apr6")
	co_prova1=RS4("CO_prova1")
	co_prova2=RS4("CO_prova2")
	co_prova3=RS4("CO_prova3")
	co_prova4=RS4("CO_prova4")	
	co_prova5=RS4("CO_prova5")
	co_prova6=RS4("CO_prova6")
			
	if co_apr1="D"AND co_prova1="D" then
		showt1="n"
	else
		showt1="s"
	end if

	if co_apr2="D" AND co_prova2="D" then
		showt2="n"
	else
		showt2="s"	
	end if
			
	if co_apr3="D" AND co_prova3="D" then
		showt3="n"
	else
		showt3="s"
	end if
					
	if co_apr4="D" AND co_prova4="D" then
		showt4="n"
	else
		showt4="s"
	end if
		
	if co_apr5="D" AND co_prova5="D" then
		showt5="n"
	ELSE
		showt5="s"	
	end if
			
	if co_apr6="D" AND co_prova6="D" then
		showt6="n"
	ELSE
		showt6="s"	
	end if							
	
end if

qtd_colunas=20
if tipo_boletim="MRD" or tipo_boletim="MRDI" then
	width_else=(width_tabela-width_disciplina-width_nu_chamada)/qtd_colunas
else
	width_else=(width_tabela-width_disciplina)/qtd_colunas
end if
%>	
<table width="<%response.Write(width_tabela)%>" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
	<%if tipo_boletim="MRD" or tipo_boletim="MRDI" then%>
	  <td width="<%response.Write(width_nu_chamada)%>" rowspan="3" class="<%response.Write(class_tit)%>">N&ordm;</td>
	 <%end if%>
	<td width="<%response.Write(width_disciplina)%>" rowspan="3" class="<%response.Write(class_tit)%>"><%response.Write(nome_foco)%></td>
	<td colspan="8" class="<%response.Write(class_tit)%>"><div align="center">Aproveitamento</div></td>
	<td colspan="4" class="<%response.Write(class_tit)%>"><div align="center">Resultado</div></td>
	<td colspan="4" class="<%response.Write(class_tit)%>"><div align="center">Freq&uuml;&ecirc;ncia (faltas)</div></td>
  </tr>
  <tr>
	<td width="<%response.Write(width_else)%>" rowspan="2" class="<%response.Write(class_subtit)%>"><div align="center">B1</div></td>
	<td width="<%response.Write(width_else)%>" rowspan="2" class="<%response.Write(class_subtit)%>"><div align="center">B2</div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(class_subtit)%>"><div align="center">MD</div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(class_subtit)%>"><div align="center">Rec</div></td>
	<td width="<%response.Write(width_else)%>" rowspan="2" class="<%response.Write(class_subtit)%>"><div align="center">B1*</div></td>
	<td width="<%response.Write(width_else)%>" rowspan="2" class="<%response.Write(class_subtit)%>"><div align="center">B2*</div></td>
	<td width="<%response.Write(width_else)%>" rowspan="2" class="<%response.Write(class_subtit)%>"><div align="center">B3</div></td>
	<td width="<%response.Write(width_else)%>" rowspan="2" class="<%response.Write(class_subtit)%>"><div align="center">B4</div>      <div align="center"></div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(class_subtit)%>"><div align="center">MD</div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(class_subtit)%>"><div align="center">Prova</div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(class_subtit)%>"><div align="center">MD</div></td>
	<td width="<%response.Write(width_else)%>" rowspan="2" class="<%response.Write(class_subtit)%>"><div align="center">RES</div></td>
	<td width="<%response.Write(width_else)%>" rowspan="2" class="<%response.Write(class_subtit)%>"><div align="center">B1</div></td>
	<td width="<%response.Write(width_else)%>" rowspan="2" class="<%response.Write(class_subtit)%>"><div align="center">B2</div></td>
	<td width="<%response.Write(width_else)%>" rowspan="2" class="<%response.Write(class_subtit)%>"><div align="center">B3</div></td>
	<td width="<%response.Write(width_else)%>" rowspan="2" class="<%response.Write(class_subtit)%>"><div align="center">B4</div></td>
  </tr>
  <tr>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(class_subtit)%>"><div align="center">SEM1</div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(class_subtit)%>"><div align="center">Par.</div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(class_subtit)%>"><div align="center">ANUAL</div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(class_subtit)%>"><div align="center">RECUP</div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(class_subtit)%>"><div align="center">FINAL</div></td>
  </tr>
<%
	check=0
	vetor_periodo=periodos(0, "num")

	if tipo_boletim="MRD" or tipo_boletim="MRDI" then
	
		n_alunos= split(alunos_vetor,"#$#")			
	
		for na=0 to ubound(n_alunos)
			aluno= split(n_alunos(na),"#!#")
			cod_cons=aluno(0)
			num_cham=aluno(1)
			nome_aluno=aluno(2)			

			medias=calcula_medias(unidade, curso, co_etapa, turma, vetor_periodo, cod_cons, vetor_materia, caminho_nota, tb_nota, nome_nota, "boletim")
			linha_medias=Split(medias,"#$#")
			
			For a=0 to ubound(co_materia)		
		
				if tipo_boletim="MRD" then
					if right(nome_aluno,16)=" - Aluno Inativo" then
						cor = "tb_fundo_linha_falta" 
						cor2 = "tb_fundo_linha_falta" 
					else
						if check mod 2 =0 then
							cor = "tb_fundo_linha_par" 
							cor2 = "tb_fundo_linha_impar" 				
						else 
							cor ="tb_fundo_linha_impar"
							cor2 = "tb_fundo_linha_par" 
						end if
					end if
				elseif tipo_boletim="MRDI" then
						cor ="tabela"
						cor2 = "tabela" 
				end if
				notas_disciplinas=Split(linha_medias(a),"#!#")
				teste0 = isnumeric(notas_disciplinas(0))
				teste1 = isnumeric(notas_disciplinas(1))
				teste2 = isnumeric(notas_disciplinas(2))
				teste3 = isnumeric(notas_disciplinas(3))
				teste4 = isnumeric(notas_disciplinas(4))
				teste5 = isnumeric(notas_disciplinas(5))
				teste6 = isnumeric(notas_disciplinas(6))
				teste7 = isnumeric(notas_disciplinas(7))
				teste8 = isnumeric(notas_disciplinas(8))
				teste9 = isnumeric(notas_disciplinas(9))
				teste10 = isnumeric(notas_disciplinas(10))
				teste11 = isnumeric(notas_disciplinas(11))
				teste12 = isnumeric(notas_disciplinas(12))
				teste13 = isnumeric(notas_disciplinas(13))
				teste14 = isnumeric(notas_disciplinas(14))
				teste15 = isnumeric(notas_disciplinas(15))
				teste16 = isnumeric(notas_disciplinas(16))													
				teste17 = isnumeric(notas_disciplinas(17))	
				teste18 = isnumeric(notas_disciplinas(18))	
				teste19 = isnumeric(notas_disciplinas(19))							
		%>
		  <tr>
			<td width="<%response.Write(width_nu_chamada)%>" class="<%response.Write(cor)%>"><div align="center"><%response.Write(num_cham)%></div></td>
			<td width="<%response.Write(width_disciplina)%>" class="<%response.Write(cor)%>"><%response.Write(nome_aluno)%></td>
			<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
		
		if teste0=false then
			if showt1="s" then
				response.Write(notas_disciplinas(0))
			else
				response.Write("&nbsp;")
			end if		
		else	
			if showt1="s" then
				notas_disciplinas(0)=notas_disciplinas(0)*1	
				ntvml=ntvml*1
				if notas_disciplinas(0)>=ntvml then	
					response.Write(formatnumber(notas_disciplinas(0),1))
				else	
					response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(0),1)&"</font>")	
				end if	
			end if	
		end if	
		
		%></div></td>
		<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
	
		if teste1=false then
			if showt2="s" then
				response.Write(notas_disciplinas(1))
			else
				response.Write("&nbsp;")
			end if		
		else	
			if showt2="s" then
				notas_disciplinas(1)=notas_disciplinas(1)*1	
				ntvml=ntvml*1
				if notas_disciplinas(1)>=ntvml then	
					response.Write(formatnumber(notas_disciplinas(1),1))
				else	
					response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(1),1)&"</font>")	
				end if	
			end if	
		end if		
		%></div></td>
		<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
		
		if teste2=false then
			if showt1="s" and showt2="s" then
				response.Write(notas_disciplinas(2))
			else
				response.Write("&nbsp;")
			end if		
		else	
			if showt1="s" and showt2="s" then
				notas_disciplinas(2)=notas_disciplinas(2)*1	
				ntvml=ntvml*1
				if notas_disciplinas(2)>=ntvml then	
					response.Write(formatnumber(notas_disciplinas(2),1))
				else	
					response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(2),1)&"</font>")	
				end if	
			end if	
		end if	
		
		%></div></td>
		<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
		
		if teste3=false then
			if showt1="s" and showt2="s" then
				response.Write(notas_disciplinas(3))
			else
				response.Write("&nbsp;")
			end if		
		else	
			if showt1="s" and showt2="s" then
				notas_disciplinas(3)=notas_disciplinas(3)*1	
				ntvml=ntvml*1
				if notas_disciplinas(3)>=ntvml then	
					response.Write(formatnumber(notas_disciplinas(3),1))
				else	
					response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(3),1)&"</font>")	
				end if	
			end if	
		end if		
		
		%></div></td>
		<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
		
		if teste4=false then
			if showt1="s" then
				response.Write(notas_disciplinas(4))
			else
				response.Write("&nbsp;")
			end if		
		else	
			if showt1="s" then
				notas_disciplinas(4)=notas_disciplinas(4)*1	
				ntvml=ntvml*1
				if notas_disciplinas(4)>=ntvml then	
					response.Write(formatnumber(notas_disciplinas(4),1))
				else	
					response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(4),1)&"</font>")	
				end if	
			end if	
		end if		
		
		%></div></td>
		<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center">
		  <%
		
		if teste5=false then
			if showt2="s" then
				response.Write(notas_disciplinas(5))
			else
				response.Write("&nbsp;")
			end if		
		else	
			if showt2="s" then
				notas_disciplinas(5)=notas_disciplinas(5)*1	
				ntvml=ntvml*1
				if notas_disciplinas(5)>=ntvml then	
					response.Write(formatnumber(notas_disciplinas(5),1))
				else	
					response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(5),1)&"</font>")	
				end if	
			end if	
		end if		
		
		%>
		  </div></td>            
		<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
		
		if teste6=false then
			if showt1="s" and showt2="s" then
				response.Write(notas_disciplinas(6))
			else
				response.Write("&nbsp;")
			end if		
		else	
			if showt1="s" and showt2="s" then
				notas_disciplinas(6)=notas_disciplinas(6)*1	
				ntvml=ntvml*1
				if notas_disciplinas(6)>=ntvml then	
					response.Write(formatnumber(notas_disciplinas(6),1))
				else	
					response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(6),1)&"</font>")	
				end if	
			end if	
		end if		
		
		%></div></td>
		 <td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
	
		if teste7=false then
			if showt3="s" then
				response.Write(notas_disciplinas(7))
			else
				response.Write("&nbsp;")
			end if		
		else	
			if showt3="s" then
				notas_disciplinas(7)=notas_disciplinas(7)*1	
				ntvml=ntvml*1
				if notas_disciplinas(7)>=ntvml then	
					response.Write(formatnumber(notas_disciplinas(7),1))
				else	
					response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(7),1)&"</font>")	
				end if	
			end if	
		end if		
		
		%></div></td>
		 <td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
		
		if teste17=false then
			if showt1="s" and showt2="s" and showt3="s" and showt4="s" then			
				response.Write(notas_disciplinas(17))
			else
				response.Write("&nbsp;")
			end if				
		else	
			if showt1="s" and showt2="s" and showt3="s" and showt4="s" then
				notas_disciplinas(17)=notas_disciplinas(17)*1	
				ntvml=ntvml*1
				if notas_disciplinas(17)>=ntvml then	
					response.Write(formatnumber(notas_disciplinas(17),1))
				else	
					response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(17),1)&"</font>")	
				end if	
			end if	
		end if		
		
		%></div></td>
		<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
		
		if teste18=false then
			if showt5="s" then
				response.Write(notas_disciplinas(18))
			else
				response.Write("&nbsp;")
			end if					
		else	
			if showt5="s" then
				notas_disciplinas(18)=notas_disciplinas(18)*1	
				ntvml=ntvml*1
				if notas_disciplinas(18)>=ntvml then	
					response.Write(formatnumber(notas_disciplinas(18),1))
				else	
					response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(18),1)&"</font>")	
				end if	
			end if	
		end if		
		
		%></div></td>  
		<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
		
		if teste19=false then
			if showt1="s" and showt2="s" and showt3="s" and showt4="s" and showt5="s" then
				response.Write(notas_disciplinas(19))
			else
				response.Write("&nbsp;")
			end if				
		else	
			if showt1="s" and showt2="s" and showt3="s" and showt4="s" and showt5="s" then
				notas_disciplinas(19)=notas_disciplinas(19)*1	
				ntvml=ntvml*1
				if notas_disciplinas(19)>=ntvml then	
					response.Write(formatnumber(notas_disciplinas(19),1))
				else	
					response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(19),1)&"</font>")	
				end if	
			end if	
		end if		
		
		%></div></td>  
		<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
		
		if teste20=false then
			if showt1="s" and showt2="s" and showt3="s" and showt4="s" and showt5="s" then
				response.Write(notas_disciplinas(20))
			else
				response.Write("&nbsp;")
			end if		
		else	
			if showt1="s" and showt2="s" and showt3="s" and showt4="s" and showt5="s" then
				notas_disciplinas(20)=notas_disciplinas(20)*1	
				ntvml=ntvml*1
				if notas_disciplinas(20)>=ntvml then	
					response.Write(formatnumber(notas_disciplinas(20),1))
				else	
					response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(20),1)&"</font>")	
				end if	
			end if	
		end if		
		
		%></div></td>                        
		<td width="<%response.Write(width_else)%>" class="<%response.Write(cor2)%>"><div align="center"><%response.Write(notas_disciplinas(13))%></div></td>
		<td width="<%response.Write(width_else)%>" class="<%response.Write(cor2)%>"><div align="center"><%response.Write(notas_disciplinas(14))%></div></td>
		<td width="<%response.Write(width_else)%>" class="<%response.Write(cor2)%>"><div align="center"><%response.Write(notas_disciplinas(15))%></div></td>
		<td width="<%response.Write(width_else)%>" class="<%response.Write(cor2)%>"><div align="center"><%response.Write(notas_disciplinas(16))%></div></td>
	  </tr>
	<%	
		check=check+1
		next 
	next	
else
	medias=calcula_medias(unidade, curso, etapa, turma, vetor_periodo, cod_cons, vetor_materia, caminho_nota, tb_nota, nome_nota, "boletim")
	linha_medias=Split(medias,"#$#")
							response.End()
	For a=0 to ubound(co_materia)
		posicao_materia=posicao_materia_tabela(co_materia(a), unidade, curso, etapa, turma)	
		
		if 	co_materia(a)<>"MED" then
			call GeraNomes(co_materia(a),unidade,curso,etapa,CON0)
			no_materia=session("no_materia")	
		else
			no_materia="M&eacute;dia"
		end if		

		if posicao_materia=0 then
			indentacao=""		
		elseif posicao_materia=1 then
			indentacao="&nbsp;&nbsp;&nbsp;"
		elseif posicao_materia=2 then
			indentacao="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		elseif posicao_materia=3 then
			indentacao="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;-->"	
		end if

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
		notas_disciplinas=Split(linha_medias(a),"#!#")
		teste0 = isnumeric(notas_disciplinas(0))
		teste1 = isnumeric(notas_disciplinas(1))
		teste2 = isnumeric(notas_disciplinas(2))
		teste3 = isnumeric(notas_disciplinas(3))
		teste4 = isnumeric(notas_disciplinas(4))
		teste5 = isnumeric(notas_disciplinas(5))
		teste6 = isnumeric(notas_disciplinas(6))
		teste7 = isnumeric(notas_disciplinas(7))
		teste8 = isnumeric(notas_disciplinas(8))
		teste9 = isnumeric(notas_disciplinas(9))
		teste10 = isnumeric(notas_disciplinas(10))
		teste11 = isnumeric(notas_disciplinas(11))
		teste12 = isnumeric(notas_disciplinas(12))
		teste13 = isnumeric(notas_disciplinas(13))
		teste14 = isnumeric(notas_disciplinas(14))
		teste15 = isnumeric(notas_disciplinas(15))
		teste16 = isnumeric(notas_disciplinas(16))													
		teste17 = isnumeric(notas_disciplinas(17))	
		teste18 = isnumeric(notas_disciplinas(18))	
		teste19 = isnumeric(notas_disciplinas(19))							
%>
  <tr>
	<td width="<%response.Write(width_disciplina)%>" class="<%response.Write(cor)%>"><%response.Write(indentacao&no_materia)%></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
	
	if teste0=false then
		if showt1="s" then
			response.Write(notas_disciplinas(0))
		else
			response.Write("&nbsp;")
		end if		
	else	
		if showt1="s" then
			notas_disciplinas(0)=notas_disciplinas(0)*1	
			ntvml=ntvml*1
			if notas_disciplinas(0)>=ntvml then	
				response.Write(formatnumber(notas_disciplinas(0),1))
			else	
				response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(0),1)&"</font>")	
			end if	
		end if	
	end if	
	
	%></div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%

	if teste1=false then
		if showt2="s" then
			response.Write(notas_disciplinas(1))
		else
			response.Write("&nbsp;")
		end if		
	else	
		if showt2="s" then
			notas_disciplinas(1)=notas_disciplinas(1)*1	
			ntvml=ntvml*1
			if notas_disciplinas(1)>=ntvml then	
				response.Write(formatnumber(notas_disciplinas(1),1))
			else	
				response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(1),1)&"</font>")	
			end if	
		end if	
	end if		
	%></div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
	
	if teste2=false then
		if showt1="s" and showt2="s" then
			response.Write(notas_disciplinas(2))
		else
			response.Write("&nbsp;")
		end if		
	else	
		if showt1="s" and showt2="s" then
			notas_disciplinas(2)=notas_disciplinas(2)*1	
			ntvml=ntvml*1
			if notas_disciplinas(2)>=ntvml then	
				response.Write(formatnumber(notas_disciplinas(2),1))
			else	
				response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(2),1)&"</font>")	
			end if	
		end if	
	end if	
	
	%></div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
	
	if teste3=false then
		if showt1="s" and showt2="s" then
			response.Write(notas_disciplinas(3))
		else
			response.Write("&nbsp;")
		end if		
	else	
		if showt1="s" and showt2="s" then
			notas_disciplinas(3)=notas_disciplinas(3)*1	
			ntvml=ntvml*1
			if notas_disciplinas(3)>=ntvml then	
				response.Write(formatnumber(notas_disciplinas(3),1))
			else	
				response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(3),1)&"</font>")	
			end if	
		end if	
	end if		
	
	%></div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
	
	if teste4=false then
		if showt1="s" then
			response.Write(notas_disciplinas(4))
		else
			response.Write("&nbsp;")
		end if		
	else	
		if showt1="s" then
			notas_disciplinas(4)=notas_disciplinas(4)*1	
			ntvml=ntvml*1
			if notas_disciplinas(4)>=ntvml then	
				response.Write(formatnumber(notas_disciplinas(4),1))
			else	
				response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(4),1)&"</font>")	
			end if	
		end if	
	end if		
	
	%></div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
	
	if teste5=false then
		if showt2="s" then
			response.Write(notas_disciplinas(5))
		else
			response.Write("&nbsp;")
		end if		
	else	
		if showt2="s" then
			notas_disciplinas(5)=notas_disciplinas(5)*1	
			ntvml=ntvml*1
			if notas_disciplinas(5)>=ntvml then	
				response.Write(formatnumber(notas_disciplinas(5),1))
			else	
				response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(5),1)&"</font>")	
			end if	
		end if	
	end if		
	
	%></div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center">
	  <%
	
	if teste6=false then
		if showt3="s" then
			response.Write(notas_disciplinas(6))
		else
			response.Write("&nbsp;")
		end if		
	else	
		if showt3="s" then
			notas_disciplinas(6)=notas_disciplinas(6)*1	
			ntvml=ntvml*1
			if notas_disciplinas(6)>=ntvml then	
				response.Write(formatnumber(notas_disciplinas(6),1))
			else	
				response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(6),1)&"</font>")	
			end if	
		end if	
	end if		
	
	%>
	  </div></td>        
	<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%

	if teste7=false then
		if showt4="s" then
			response.Write(notas_disciplinas(7))
		else
			response.Write("&nbsp;")
		end if		
	else	
		if showt4="s" then
			notas_disciplinas(7)=notas_disciplinas(7)*1	
			ntvml=ntvml*1
			if notas_disciplinas(7)>=ntvml then	
				response.Write(formatnumber(notas_disciplinas(7),1))
			else	
				response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(7),1)&"</font>")	
			end if	
		end if	
	end if		
	
	%></div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center">
	  <%
	
	if teste17=false then
		if showt1="s" and showt2="s" and showt3="s" and showt4="s" then		
			response.Write(notas_disciplinas(17))
		else
			response.Write("&nbsp;")
		end if				
	else	
		if showt1="s" and showt2="s" and showt3="s" and showt4="s" then
			notas_disciplinas(17)=notas_disciplinas(17)*1	
			ntvml=ntvml*1
			if notas_disciplinas(17)>=ntvml then	
				response.Write(formatnumber(notas_disciplinas(17),1))
			else	
				response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(17),1)&"</font>")	
			end if	
		end if	
	end if		
	
	%>
	</div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center">
    <%
	
	if teste18=false then
		if showt5="s" then
			response.Write(notas_disciplinas(18))
		else
			response.Write("&nbsp;")
		end if					
	else	
		if showt5="s" then
			notas_disciplinas(18)=notas_disciplinas(18)*1	
			ntvml=ntvml*1
			if notas_disciplinas(18)>=ntvml then	
				response.Write(formatnumber(notas_disciplinas(18),1))
			else	
				response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(18),1)&"</font>")	
			end if	
		end if	
	end if		
	
	%></div></td>  
	<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center"><%
	
	if teste19=false then
		if showt1="s" and showt2="s" and showt3="s" and showt4="s" and showt5="s" then
			response.Write(notas_disciplinas(19))
		else
			response.Write("&nbsp;")
		end if				
	else	
		if showt1="s" and showt2="s" and showt3="s" and showt4="s" and showt5="s" then
			notas_disciplinas(19)=notas_disciplinas(19)*1	
			ntvml=ntvml*1
			if notas_disciplinas(19)>=ntvml then	
				response.Write(formatnumber(notas_disciplinas(19),1))
			else	
				response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(19),1)&"</font>")	
			end if	
		end if	
	end if		
	
	%></div></td>   
	<td width="<%response.Write(width_else)%>" class="<%response.Write(cor)%>"><div align="center">
	  <%
	
	if teste20=false then
		if showt1="s" and showt2="s" and showt3="s" and showt4="s" and showt5="s" then
			response.Write(notas_disciplinas(20))
		else
			response.Write("&nbsp;")
		end if		
	else	
		if showt1="s" and showt2="s" and showt3="s" and showt4="s" and showt5="s" then
			notas_disciplinas(20)=notas_disciplinas(20)*1	
			ntvml=ntvml*1
			if notas_disciplinas(20)>=ntvml then	
				response.Write(formatnumber(notas_disciplinas(20),1))
			else	
				response.Write("<font color="&cor_nota&">"&formatnumber(notas_disciplinas(20),1)&"</font>")	
			end if	
		end if	
	end if		
	
	%>
	</div></td>                   
	<td width="<%response.Write(width_else)%>" class="<%response.Write(cor2)%>"><div align="center"><%response.Write(notas_disciplinas(13))%></div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(cor2)%>"><div align="center"><%response.Write(notas_disciplinas(14))%></div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(cor2)%>"><div align="center"><%response.Write(notas_disciplinas(15))%></div></td>
	<td width="<%response.Write(width_else)%>" class="<%response.Write(cor2)%>"><div align="center"><%response.Write(notas_disciplinas(16))%></div></td>
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
