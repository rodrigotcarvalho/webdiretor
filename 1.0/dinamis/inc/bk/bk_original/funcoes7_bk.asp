<!--#include file="../../global/funcoes_diversas.asp" -->
<!--#include file="funcoes6.asp"-->
<%

Function grava_ficha(unidade, curso, co_etapa, turma, vetor_periodo_ctrl)
Server.ScriptTimeout = 900 'valor em segundos



	Set CONt = Server.CreateObject("ADODB.Connection") 
	ABRIRt = "DBQ="& CAMINHO_t & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONt.Open ABRIRt
	
	ano = DatePart("yyyy", now)
	mes = DatePart("m", now) 
	dia = DatePart("d", now) 
	hora = DatePart("h", now) 
	min = DatePart("n", now) 
	data = dia &"/"& mes &"/"& ano
	horario = hora & ":"& min
	
	Set RS5 = Server.CreateObject("ADODB.Recordset")
	SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
	RS5.Open SQL5, CON0
	co_materia_check=1
	IF RS5.EOF Then
		vetor_materia_exibe="nulo"
	else
		while not RS5.EOF
			co_mat_fil= RS5("CO_Materia")				
			if co_materia_check=1 then
				vetor_materia=co_mat_fil
			else
				vetor_materia=vetor_materia&"#!#"&co_mat_fil
			end if
			co_materia_check=co_materia_check+1			
					
		RS5.MOVENEXT
		wend	
	'	
		vetor_materia_exibe=programa_aula(vetor_materia, unidade, curso, co_etapa, "nulo")			
'		response.Write(vetor_materia_exibe)
'		response.end()
	end if	
'	response.Write(vetor_periodo_ctrl&"<BR>")
'	response.Write(vetor_materia_exibe)
'response.end()	
	vetor_num_periodo=vetor_periodo_ctrl
	co_materia_exibe=Split(vetor_materia_exibe,"#!#")	
	periodo_max=Split(vetor_periodo_ctrl,"#!#")

'	Set RS0 = Server.CreateObject("ADODB.Recordset")
'	SQL0 = "SELECT * FROM TB_Periodo ORDER BY NU_Periodo"
'	RS0.Open SQL0, CON0
'	check_periodo=1
'	
'	WHILE NOT RS0.EOF
'		p=RS0("NU_Periodo")
	'	response.Write(periodo_max(Ubound(periodo_max ))&"<BR>")		
		'if p=1 then
			'temp_num_periodo=p
			'sigla_periodo=RS0("SG_Periodo")
			temp_nomes_periodos="Disciplina#!#M&eacute;dia<BR>1&ordm Tri#!#M&eacute;dia<BR>Acum."
		'else
'		periodo_max(Ubound(periodo_max))=periodo_max(Ubound(periodo_max))*1
'		p=p*1
'			if p>periodo_max(Ubound(periodo_max)) then
'				temp_num_periodo=temp_num_periodo
'			else
'				temp_num_periodo=temp_num_periodo&"#!#"&p
'			end if
'			sigla_periodo=RS0("SG_Periodo")
'			if p=5 then
'				temp_nomes_periodos=temp_nomes_periodos&"#!#M&eacute;dia<br>Anual#!#Result.#!#Prova<br>Final"
'			elseif p=6 then
'				temp_nomes_periodos=temp_nomes_periodos&"#!#M&eacute;dia<br>Final#!#Result.#!#Recup.#!#Result."
'			else
'				temp_nomes_periodos=temp_nomes_periodos&"#!#"&sigla_periodo				
'			end if
			
			
			temp_nomes_periodos=temp_nomes_periodos&"#!#M&eacute;dia<BR>2&ordm Tri#!#M&eacute;dia<BR>Acum.#!#M&eacute;dia<BR>3&ordm Tri#!#M&eacute;dia<BR>Acum."
			temp_nomes_periodos=temp_nomes_periodos&"#!#Result.#!#ECE#!#M&eacute;dia<br>Final#!#Result.<BR>Final#!#1&ordm; Tri#!#2&ordm; Tri#!#3&ordm; Tri#!#Total"
			
'		end if
'	RS0.MOVENEXT
'	WEND		
	
	
	
'response.Write("2TADA "&temp_num_periodo)
'Response.end()
'	vetor_num_periodo=temp_num_periodo
		
	vetor_nomes_periodos=temp_nomes_periodos&"#!#Carga"
	ajusta_periodos=split(vetor_nomes_periodos,"#!#")
	ultimo_campo_periodo=ubound(ajusta_periodos)+1

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
	
	tb_nota=tabela_nota(ano_letivo,unidade,curso,co_etapa,turma,"tb",0)
	caminho_nota=tabela_nota(ano_letivo,unidade,curso,co_etapa,turma,"cam",0)
	
	'num_periodo=split(vetor_num_periodo,"#!#")	
	
	alunos_vetor=alunos_turma(ano_letivo,unidade,curso,co_etapa,turma,"nome")
	n_alunos= split(alunos_vetor,"#$#")			

		for al=0 to ubound(n_alunos)
			aluno= split(n_alunos(al),"#!#")
			cod_cons=aluno(0)
'			response.Write(vetor_num_periodo)
'			response.end()

			medias=calcula_medias(unidade, curso, co_etapa, turma, vetor_num_periodo, cod_cons, vetor_materia, caminho_nota, tb_nota,"VA_Media3", "boletim")
'			medias="#$#"
'			response.Write(medias&"<BR><BR>")


			medias_materia = split(medias,"#$#")
				

			qtd_medias_materia = ubound(medias_materia)		
			
			ordem_exibe=1
			for k=0 to qtd_medias_materia		
				co_materia_consulta=co_materia_exibe(k)
			'response.Write(co_materia_consulta&"   "&medias_materia(k)&" - "&k&"<BR>")			
				if 	co_materia_consulta<>"MED" then
					call GeraNomes(co_materia_consulta,unidade,curso,etapa,CON0)
					no_materia_exibe=session("no_materia")	
					
					posicao_materia=posicao_materia_tabela(co_materia_consulta, unidade, curso, co_etapa, turma)	
					posicao_materia=posicao_materia*1					
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
									
					grava_notas = split(medias_materia(k),"#!#")
					
					for tn=0 to ubound(grava_notas)			
						n_campo=tn+2
						if n_campo<10 then
							campo_gravacao="CO_0"&n_campo
						else
							campo_gravacao="CO_"&n_campo						
						end if				
						
						
						periodo_max(Ubound(periodo_max))=periodo_max(Ubound(periodo_max))*1
						periodo_max(Ubound(periodo_max))=periodo_max(Ubound(periodo_max))+1
						if grava_notas(tn) ="&nbsp;" then
							grava=NULL
						elseif periodo_max(Ubound(periodo_max))=1 and (campo_gravacao<>"CO_02" and campo_gravacao<>"CO_03" AND campo_gravacao<>"CO_12") then
							grava=NULL
						elseif periodo_max(Ubound(periodo_max))=2 and (campo_gravacao<>"CO_02" and campo_gravacao<>"CO_03" AND campo_gravacao<>"CO_04" and campo_gravacao<>"CO_05" AND campo_gravacao<>"CO_12" AND campo_gravacao<>"CO_13") then
							grava=NULL
						elseif periodo_max(Ubound(periodo_max))=3 and (campo_gravacao="CO_09" or campo_gravacao="CO_10" or campo_gravacao="CO_11") then
							grava=NULL
						else						 
							grava=grava_notas(tn)
						end if
						
						if n_campo<12 and ((curso=1 and co_etapa<6 and (co_materia_consulta="ART" or co_materia_consulta="ESP")) or co_materia_consulta="EFI") then									
							teste_media = isnumeric(grava)							
							if teste_media=TRUE then							
								if grava >= 9 then
								grava="A"
								elseif (grava >= 7) and (grava < 9) then
								grava="B"
								elseif (grava >= 5) and (grava < 7) then							
								grava="C"
								elseif (grava >= 3) and (grava < 5) then
								grava="D"
								else							
								grava="E"
								end if	
							end if
						end if		
	
						RS4(campo_gravacao)=grava
					next
					
					if ultimo_campo_periodo<10 then
						campo_gravacao="CO_0"&ultimo_campo_periodo
					else
						campo_gravacao="CO_"&ultimo_campo_periodo						
					end if				
					if no_materia_exibe="&nbsp;&nbsp;&nbsp;&nbsp;-->&nbsp;M&eacute;dia" then
						RS4(campo_gravacao)=NULL
					else
						RS4(campo_gravacao)=carga_materia	
					end if	
				RS4.update
				RS4.Close
				Set RS4 = Nothing
				ordem_exibe=ordem_exibe*1		
				ordem_exibe=ordem_exibe+1		
			next
		next			
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


Function periodos_ACC(periodo,acumulado,qto_falta,id,outro)

Server.ScriptTimeout = 900

	if acumulado="s" then
		for p=1 to periodo
			if p=1 then
				temp_num_periodo=p
				sigla_periodo=periodos(p,"sigla")
				temp_nomes_periodos=sigla_periodo
			else
				temp_num_periodo=temp_num_periodo&"#!#"&p
				sigla_periodo=periodos(p,"sigla")
				temp_nomes_periodos=temp_nomes_periodos&"#!#"&sigla_periodo
			end if
		next
		if qto_falta="s" then
			vetor_periodo=split(temp_nomes_periodos,"#!#")
			num_periodo=split(temp_num_periodo,"#!#")		
			for v=0 to ubound(vetor_periodo)
				if vetor_periodo(v)="TRI1" then	
					temp_num_periodo=1
					periodo_exibe=vetor_periodo(v)
				elseif vetor_periodo(v)="TRI2" then	
					temp_num_periodo=temp_num_periodo&"#!#2#!#0"
					periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)&"#!#QF1"
				elseif vetor_periodo(v)="TRI3" then	
					temp_num_periodo=temp_num_periodo&"#!#3#!#0#!#0"
					periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)&"#!#MA#!#QF2"
				elseif vetor_periodo(v)="ECE" then	
					temp_num_periodo=temp_num_periodo&"#!#4#!#0"
					periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)&"#!#MF"										
				end if	
			next										
		else
			vetor_periodo=split(temp_nomes_periodos,"#!#")
			num_periodo=split(temp_num_periodo,"#!#")		
			for v=0 to ubound(vetor_periodo)
				if vetor_periodo(v)="TRI1" then	
					temp_num_periodo=1
					periodo_exibe=vetor_periodo(v)
				elseif vetor_periodo(v)="TRI2" then	
					temp_num_periodo=temp_num_periodo&"#!#2"
					periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)
				elseif vetor_periodo(v)="TRI3" then	
					temp_num_periodo=temp_num_periodo&"#!#3#!#0"
					periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)&"#!#MA"
				elseif vetor_periodo(v)="ECE" then	
					temp_num_periodo=temp_num_periodo&"#!#4#!#0"
					periodo_exibe=periodo_exibe&"#!#"&vetor_periodo(v)&"#!#MF"										
				end if	
			next					
		end if	
	else	
		temp_num_periodo=periodo
		sigla_periodo=periodos(periodo,"sigla")
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

	vetor_materias=vetor_disciplinas(ano_letivo,unidade,curso,co_etapa,turma,"s",0)
	
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

	periodo_m1=3
	periodo_m2=4
	periodo_m3=0

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
	
	tb_nota=tabela_nota(ano_letivo,unidade,curso,co_etapa,turma,"tb",0)
	caminho_nota=tabela_nota(ano_letivo,unidade,curso,co_etapa,turma,"cam",0)
	
	if acumulado="s" then
		vetor_num_periodos=periodos_ACC(periodo,"s",qto_falta,"num",0)
		vetor_nom_periodos=periodos_ACC(periodo,"s",qto_falta,"nom",0)
	else
		vetor_num_periodos=periodo
		vetor_nom_periodos=periodos_ACC(periodo,"n","n","nom",0)
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
						media=ACC(unidade, curso, co_etapa, turma, cod_cons, ajusta_materias(mat), caminho_nota, tb_nota, nom_periodo(per), num_periodo(per), periodo_m1, periodo_m2, periodo_m3, nota_m1, nota_m2, 999, peso_m2_m1, peso_m2_m2, peso_m3_m1, peso_m3_m2, peso_m3_m3)	
						if nom_periodo(per)<>"QF1" and nom_periodo(per)<>"QF2" and ((curso=1 and co_etapa<6 and (ajusta_materias(mat)="ART" or ajusta_materias(mat)="ESP" or ajusta_materias(mat)="ING")) or ajusta_materias(mat)="EFI") then																	
							if media="&nbsp;" or isnull(media) or media="" then
							else
								media=media*1				
								if media >= 9 then
									media="A"
								elseif (media >= 7) and (media < 9) then
									media="B"
								elseif (media >= 5) and (media < 7) then							
									media="C"
								elseif (media >= 3) and (media < 5) then
									media="D"
								else							
									media="E"
								end if		
							end if			
						end if	
					end if			
					if mat=0 then
						vetor_grava_notas=media
					else	
						vetor_grava_notas=vetor_grava_notas&"#!#"&media
					end if						
				next
				vetor_grava_notas=replace(vetor_grava_notas,"&nbsp;","")
'				response.Write(vetor_grava_notas)
'				response.end()
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
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& codigo_materia &"'"
	RS.Open SQL, CON0
	
	mae= RS("IN_MAE")
	fil= RS("IN_FIL")
	in_co= RS("IN_CO")
	peso= RS("NU_Peso")
'	response.Write(SQL&" - no_periodo="&no_periodo&"<BR>")	
'	response.Write(no_periodo&" - mae="&mae&" and fil="&fil&" and in_co="&in_co&" and peso="&peso&"<BR>")
'	if codigo_materia="INFO" and no_periodo="RECF" then
'	response.end()
'	end if
	
	if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(peso)) then
		if no_periodo<>"QF1" and no_periodo<>"QF2" and no_periodo<>"MA" and no_periodo<>"MS"and no_periodo<>"RECP" and no_periodo<>"BIM1*"and no_periodo<>"BIM2*" and no_periodo<>"MF" then	
			media=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, periodo,"nota_periodo")	
'	response.Write(media&" - codigo_materia="&codigo_materia&"<BR>")	
			if media="&nbsp;" or media="" or isnull(media) then
			else
				if media=0 then
					media="&nbsp;"
				else
					media=media/10
					media=arredonda(media,"mat_dez",1,0)	
				end if	
			end if					
		else
			acumula_media=0
			if no_periodo="QF1" then
			periodo_qf=periodo_m1-1
'				for periodo=1 to periodo_qf
'					qf=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, periodo)	
'
'					if qf="&nbsp;" or qf="" or isnull(qf) then
'						acumula_media=acumula_media
'
'					else
'						acumula_media=acumula_media+qf
'				
'					end if	
'				next
				resultado=Calc_Med_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"nulo", "nulo")								
				medias=split(resultado,"#!#")
				
				b1=medias(0)
				b2=medias(1)
				ms=medias(2)
				rp=medias(3)
				b1a=medias(4)
				b2a=medias(5)
				b3=medias(6)
				
				if b1="&nbsp;" or b1="" or isnull(b1) then
					media_qf=""
				else
					if b2="&nbsp;" or b2="" or isnull(b2) then								
						acumula_media=b1
					
					else	
						b1=b1/10
						b1=arredonda(b1,"mat_dez",1,0)	
						b2=b2/10
						b2=arredonda(b2,"mat_dez",1,0)	
						b1=b1*1	
						b2=b2*1													
						acumula_media=b1+b2	
				
					end if						
									
'					if b3="&nbsp;" or b3="" or isnull(b3) then								
'
'					else	
'						acumula_media=acumula_media*1						
'						b3=b3/10
'						b3=arredonda(b3,"mat_dez",1,0)								
'						acumula_media=acumula_media+b3	
'						
'					end if	
					
					nota_m1=nota_m1*1
					periodo_m1=periodo_m1*1
					acumula_media=acumula_media*1				
					
					media_qf=(nota_m1*periodo_m1)-acumula_media					
				end if	
				
				if media_qf="" or isnull(media_qf)then
					media=""
				else	
					media=arredonda(media_qf,"mat_dez",1,0)				
				end if
			elseif no_periodo="QF2" then
				verifica=Calc_Med_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"nulo", "nulo")								
				medias=split(verifica,"#!#")
				b1=medias(0)
				b2=medias(1)
				ms=medias(2)
				rp=medias(3)
				b1a=medias(4)
				b2a=medias(5)
				b3=medias(6)			
				if b1<>"&nbsp;" and b2<>"&nbsp;" and b3<>"&nbsp;" then 		
					resultado=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"final", 0)
					if resultado="&nbsp;#!#&nbsp;" then
						media=""	
					else
						media_qf=split(resultado,"#!#")
						media_qf(0)=media_qf(0)*1
						nota_m1=nota_m1*1	
						if media_qf(0)>=nota_m1 then
							media=""
						else
							peso_m2_m1=peso_m2_m1*1
							peso_m2_m2=peso_m2_m2*1	
							media=((nota_m2*(peso_m2_m1+peso_m2_m2))-(media_qf(0)*peso_m2_m1))/peso_m2_m2					
							if media<0 then
								media=""	
							else
								media=media/10					
								media=arredonda(media,"mat_dez",1,0)
							end if
						end if	
					end if
				ELSE	
					media=""
				END IF
			elseif no_periodo="MA" then	
				verifica=Calc_Med_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"nulo", "nulo")								
				medias=split(verifica,"#!#")
				b1=medias(0)
				b2=medias(1)
				ms=medias(2)
				rp=medias(3)
				b1a=medias(4)
				b2a=medias(5)
				b3=medias(6)			
				if b1<>"&nbsp;" and b2<>"&nbsp;" and b3<>"&nbsp;" then 									
					resultado=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"anual", outro)						
					media_qf=split(resultado,"#!#")
					media=media_qf(0)
				ELSE	
					media=""
				END IF				
			elseif no_periodo="MS" then		
				resultado=Calc_Med_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"nulo", "nulo")								
				media_sem=split(resultado,"#!#")
				media=media_sem(2)
				'media="MS"
			elseif no_periodo="RECP" then						
				resultado=Calc_Med_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"nulo", "nulo")				
				media_sem=split(resultado,"#!#")
				media=media_sem(3)
			elseif no_periodo="BIM1*" then		
				resultado=Calc_Med_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"nulo", "nulo")			
				media_sem=split(resultado,"#!#")
				media=media_sem(4)		
			elseif no_periodo="BIM2*" then		
				resultado=Calc_Med_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"nulo", "nulo")						
				media_sem=split(resultado,"#!#")
				media=media_sem(5)																	
			elseif no_periodo="MF" then
				verifica=Calc_Med_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"nulo", "nulo")								
				medias=split(verifica,"#!#")
				b1=medias(0)
				b2=medias(1)
				ms=medias(2)
				rp=medias(3)
				b1a=medias(4)
				b2a=medias(5)
				b3=medias(6)
				b4=medias(7)				
				if b1<>"&nbsp;" and b2<>"&nbsp;" and b3<>"&nbsp;" and b4<>"&nbsp;" then 				
					m5=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4,"nota_acumul")	
					if m5="&nbsp;" or m5="" or isnull(m5) then
						acumula_media=acumula_media
						media="&nbsp;"
					else
						resultado=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"final", outro)						
						media_qf=split(resultado,"#!#")
						media=media_qf(0)			
					end if				
				else
					media=""
				end if	
			else
				media=""
			end if	
		end if
	elseif mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso) then
'		co_materia_fil_check=1 
'		
'		Set RS1a = Server.CreateObject("ADODB.Recordset")
'		SQL1a = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& codigo_materia &"' order by NU_Ordem_Boletim"
'		RS1a.Open SQL1a, CON0	
'		
'		if RS1a.EOF then
'			response.Write("ERRO TB_Materia - ACC1")
'			response.end()
'		else
'			while not RS1a.EOF
'				co_mat_fil= RS1a("CO_Materia")				
'				if co_materia_fil_check=1 then
'					vetor_materia=co_mat_fil
'				else
'					vetor_materia=vetor_materia&"#!#"&co_mat_fil			
'				end if
'				co_materia_fil_check=co_materia_fil_check+1 									
'			RS1a.MOVENEXT
'			wend	
'		end if		
'		if no_periodo<>"QF1" and no_periodo<>"QF2" and no_periodo<>"MB" and no_periodo<>"MS"and no_periodo<>"BIM1*"and no_periodo<>"BIM2*" and no_periodo<>"MF" then	
'			media=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, vetor_materia, caminho_nota, tb_nota, periodo)	
'			periodo=periodo*1
'
'			if periodo=1 then
'				Session("md_p1")=media
'			elseif periodo=2 then
'				Session("md_p2")=media
'			elseif periodo=3 then
'				Session("md_p3")=media		
'			elseif periodo=4 then
'				Session("md_p4")=media		
'			elseif periodo=5 then
'				Session("md_p5")=media		
'			elseif periodo=6 then
'				Session("md_p6")=media		
'			end if																					
'			media=""
'		else
'			acumula_media=0
'			if no_periodo="QF1" then
'			periodo_qf=periodo_m1-1
'				for periodo=1 to periodo_qf
'					qf=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, vetor_materia, caminho_nota, tb_nota, periodo)	
'					response.Write(acumula_media&"p"&periodo&"<br>")
'					if qf="&nbsp;" or qf="" or isnull(qf) then
'						acumula_media=acumula_media
'					else
'						acumula_media=acumula_media+qf
'					end if	
'				next
'				nota_m1=nota_m1*1
'				periodo_m1=periodo_m1*1
'				acumula_media=acumula_media*1
'				response.Write("#"&nota_m1&"#'"&periodo_m1&"'$"&acumula_media&"$")
'				media_qf=(nota_m1*periodo_m1)-acumula_media
'				
'				response.Write("'"&media_qf&"'<BR>")
'				
'				media_qf=acumula_media/periodo_m1
'				if media_qf<0 or media_qf="" or isnull(media_qf)then
'					media=""
'				else	
'					media=nota_m1-media_qf
'					media=arredonda(media_qf,"mat_dez",1,0)
'				end if
'				Session("md_qf1")=media
'				media=""	
'			elseif no_periodo="QF2" then	
'				resultado=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"anual", 0)
'				if resultado="&nbsp;#!#&nbsp;" then
'					media=""	
'				else
'					media_qf=split(resultado,"#!#")
'					media_qf(0)=media_qf(0)*1
'					nota_m1=nota_m1*1	
'					if media_qf(0)>=nota_m1 then
'						media=""
'					else
'						peso_m2_m1=peso_m2_m1*1
'						peso_m2_m2=peso_m2_m2*1		
'	
'						media=((nota_m2*(peso_m2_m1+peso_m2_m2))-(media_qf(0)*peso_m2_m1))/peso_m2_m2
'						if media<0 then
'							media=""	
'						else					
'							media=arredonda(media,"mat_dez",1,0)
'						end if
'					end if	
'				end if
'				Session("md_qf2")=media
'				media=""
'			elseif no_periodo="MB" then		
'				resultado=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"anual", outro)						
'				media_qf=split(resultado,"#!#")
'				media=media_qf(0)
'				Session("md_mb")=media
'				media=""
'			elseif no_periodo="MS" then		
'				
'				resultado=Calc_Med_T_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"nulo", "nulo")				
'				media_sem=split(resultado,"#!#")
'				media=media_sem(2)
'			elseif no_periodo="RECP" then		
'				
'				resultado=Calc_Med_T_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"nulo", "nulo")				
'				media_sem=split(resultado,"#!#")
'				media=media_sem(3)
'			elseif no_periodo="BIM1*" then		
'				
'				resultado=Calc_Med_T_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"nulo", "nulo")						
'				media_sem=split(resultado,"#!#")
'				media=media_sem(3)		
'			elseif no_periodo="BIM2*" then		
'				
'				resultado=Calc_Med_T_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"nulo", "nulo")					
'				media_sem=split(resultado,"#!#")
'				media=media_sem(4)					
'			elseif no_periodo="MF" then
'				m5=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, vetor_materia, caminho_nota, tb_nota, 4)	
'				if m5="&nbsp;" or m5="" or isnull(m5) then
'					acumula_media=acumula_media
'					media="&nbsp;"
'				else
'					resultado=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"final", outro)						
'					media_qf=split(resultado,"#!#")
'					media=media_qf(0)			
'				end if	
'				Session("md_mf")=media
'				media=""
'			else
'				media=""
'			end if	
'		end if		

	elseif(mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then		
		
'		co_materia_fil_check=1
'		
'		Set RS1a = Server.CreateObject("ADODB.Recordset")
'		SQL1a = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& codigo_materia &"' order by NU_Ordem_Boletim"
'		RS1a.Open SQL1a, CON0
'			
'		if RS1.EOF then
'			response.Write("ERRO TB_Materia - ACC2")
'			response.end()	
'		else
'			while not RS1.EOF
'				co_mat_fil= RS1("CO_Materia")				
'				if co_materia_fil_check=1 then
'					vetor_materia=vetor_materia&"#!#"&codigo_materia&"#!#"&co_mat_fil
'				else
'					vetor_materia=vetor_materia&"#!#"&co_mat_fil			
'				end if
'				co_materia_fil_check=co_materia_fil_check+1 									
'			RS1.MOVENEXT
'			wend
'		end if	
'		if no_periodo<>"QF1" and no_periodo<>"QF2" and no_periodo<>"MB" and no_periodo<>"MS"and no_periodo<>"BIM1*"and no_periodo<>"BIM2*" and no_periodo<>"MF" then		
'			media=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, cod_aluno, vetor_materia, caminho_nota, tb_nota, periodo)				
'		else
'			acumula_media=0
'			if no_periodo="QF1" then
'			periodo_qf=periodo_m1-1
'				for periodo=1 to periodo_qf
'					qf=Calcula_Media_T_F_T_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, periodo)	
'					response.Write(acumula_media&"p"&periodo&"<br>")
'					if qf="&nbsp;" or qf="" or isnull(qf) then
'						acumula_media=acumula_media
'					else
'						acumula_media=acumula_media+qf
'					end if	
'				next
'				nota_m1=nota_m1*1
'				periodo_m1=periodo_m1*1
'				acumula_media=acumula_media*1
'				response.Write("#"&nota_m1&"#'"&periodo_m1&"'$"&acumula_media&"$")
'				media_qf=(nota_m1*periodo_m1)-acumula_media
'				
'				response.Write("'"&media_qf&"'<BR>")
'				
'				media_qf=acumula_media/periodo_m1
'				if media_qf<0 or media_qf="" or isnull(media_qf)then
'					media=""
'				else	
'					media=nota_m1-media_qf
'					media=arredonda(media_qf,"mat_dez",1,0)
'				end if
'			elseif no_periodo="QF2" then	
'				resultado=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"anual", 0)
'				if resultado="&nbsp;#!#&nbsp;" then
'					media=""	
'				else
'					media_qf=split(resultado,"#!#")
'					media_qf(0)=media_qf(0)*1
'					nota_m1=nota_m1*1	
'					if media_qf(0)>=nota_m1 then
'						media=""
'					else
'						peso_m2_m1=peso_m2_m1*1
'						peso_m2_m2=peso_m2_m2*1		
'	
'						media=((nota_m2*(peso_m2_m1+peso_m2_m2))-(media_qf(0)*peso_m2_m1))/peso_m2_m2
'						if media<0 then
'							media=""	
'						else					
'							media=arredonda(media,"mat_dez",1,0)
'						end if
'					end if	
'				end if
'			media=""
'			elseif no_periodo="MB" then		
'				resultado=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"anual", outro)						
'				media_qf=split(resultado,"#!#")
'				media=media_qf(0)
'			elseif no_periodo="MS" then		
'				resultado=Calc_Med_T_F_T_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"nulo", "nulo")						
'				media_sem=split(resultado,"#!#")
'				media=media_sem(2)
'				media="MS"
'			elseif no_periodo="BIM1*" then		
'				resultado=Calc_Med_T_F_T_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"nulo", "nulo")					
'				media_sem=split(resultado,"#!#")
'				media=media_sem(3)		
'			elseif no_periodo="BIM2*" then		
'				resultado=Calc_Med_T_F_T_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"nulo", "nulo")						
'				media_sem=split(resultado,"#!#")
'				media=media_sem(4)				
'			elseif no_periodo="MF" then
'				m5=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4)	
'				if m5="&nbsp;" or m5="" or isnull(m5) then
'					acumula_media=acumula_media
'					media="&nbsp;"
'				else
'					resultado=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"final", outro)						
'					media_qf=split(resultado,"#!#")
'					media=media_qf(0)			
'				end if	
'			else
'				media=""
'			end if	
'		end if
	elseif (mae=FALSE and fil=TRUE and in_co=FALSE) then
'		Set RS2 = Server.CreateObject("ADODB.Recordset")
'		SQL2 = "SELECT * FROM TB_Materia where CO_Materia ='"& codigo_materia &"'"
'		RS2.Open SQL2, CON0
'			
'		co_materia_fil_check=0 
'			codigo_materia_pr= RS2("CO_Materia_Principal")	
'
'
'		if no_periodo<>"QF1" and no_periodo<>"QF2" and no_periodo<>"MB" and no_periodo<>"MS"and no_periodo<>"BIM1*"and no_periodo<>"BIM2*" and no_periodo<>"MF" then
'			media=Calcula_Media_F_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia_pr, codigo_materia, caminho_nota, tb_nota, periodo)	
'			
'		else
'			acumula_media=0
'			if no_periodo="QF1" then
'			periodo_qf=periodo_m1-1
'				for periodo=1 to periodo_qf
'					qf=Calcula_Media_F_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia_pr, codigo_materia, caminho_nota, tb_nota, periodo)	
'
'					if qf="&nbsp;" or qf="" or isnull(qf) then
'						acumula_media=acumula_media
'
'					else
'						acumula_media=acumula_media+qf
'				
'					end if	
'				next
'				nota_m1=nota_m1*1
'				periodo_m1=periodo_m1*1
'				acumula_media=acumula_media*1
'				media_qf=(nota_m1*periodo_m1)-acumula_media
'
'				if media_qf<0 or media_qf="" or isnull(media_qf)then
'					media=""
'				else	
'					media=arredonda(media_qf,"mat_dez",1,0)				
'				end if
''			media=""
'			elseif no_periodo="QF2" then	
''				resultado=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"anual", 0)
''				if resultado="&nbsp;#!#&nbsp;" then
''					media=""	
''				else
''					media_qf=split(resultado,"#!#")
''					media_qf(0)=media_qf(0)*1
''					nota_m1=nota_m1*1	
''					if media_qf(0)>=nota_m1 then
''						media=""
''					else
''						peso_m2_m1=peso_m2_m1*1
''						peso_m2_m2=peso_m2_m2*1	
''						media=((nota_m2*(peso_m2_m1+peso_m2_m2))-(media_qf(0)*peso_m2_m1))/peso_m2_m2					
''						if media<0 then
''							media=""	
''						else					
''							media=arredonda(media,"mat_dez",1,0)
''						end if
''					end if	
''				end if
'				media=""
'			elseif no_periodo="MB" then		
''				resultado=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"anual", outro)						
''				media_qf=split(resultado,"#!#")
''				media=media_qf(0)
'				media=""
'			elseif no_periodo="MS" then		
'				resultado=Calc_Med_F_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"nulo", "nulo")						
'				media_sem=split(resultado,"#!#")
'				media=media_sem(2)
'				'media="MS"
'			elseif no_periodo="BIM1*" then		
'				resultado=Calc_Med_F_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"nulo", "nulo")					
'				media_sem=split(resultado,"#!#")
'				media=media_sem(3)		
'			elseif no_periodo="BIM2*" then		
'				resultado=Calc_Med_F_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"nulo", "nulo")						
'				media_sem=split(resultado,"#!#")
'				media=media_sem(4)					
'			elseif no_periodo="MF" then
''				m5=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4)	
''				if m5="&nbsp;" or m5="" or isnull(m5) then
''					acumula_media=acumula_media
''					media="&nbsp;"
''				else
''					resultado=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, 4, 4, 0,"final", outro)						
''					media_qf=split(resultado,"#!#")
''					media=media_qf(0)			
''				end if				
'				media=""
'			else
'				media=""
'			end if	
'		end if
'	
'	
	end if
end if	
ACC=media	
end function

'===========================================================================================================================================
'serve também para (mae=FALSE and fil=FALSE and in_co=TRUE) para o Mapa de Resultados por Disciplinas		
Function Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, periodo, nome_nota)

Server.ScriptTimeout = 900


	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CONn = Server.CreateObject("ADODB.Connection") 
	ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONn.Open ABRIRn	
	
	
	if nome_nota="nota_periodo" then
		if periodo=1 then
			m_cons="VA_Me1"
		elseif periodo=2 then
			m_cons="VA_Me2"
		elseif periodo=3 then
			m_cons="VA_Me3"
		elseif periodo=4 then
			m_cons="VA_Me_EC"
		elseif periodo=5 then
			m_cons="VA_Media3"
		elseif periodo=6 then
			m_cons="VA_Media3"
		end if		
	else
		if periodo=1 then
			m_cons="VA_Mc1"
		elseif periodo=2 then
			m_cons="VA_Mc2"
		elseif periodo=3 then
			m_cons="VA_Mc3"
		elseif periodo=4 then
			m_cons="VA_Mfinal"
		elseif periodo=5 then
			m_cons="VA_Media3"
		elseif periodo=6 then
			m_cons="VA_Media3"
		end if	
	end if		
		
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM "&tb_nota&" where CO_Matricula ="& cod_aluno &" AND CO_Materia_Principal ='"& codigo_materia &"' AND CO_Materia ='"& codigo_materia &"'"
		RS1.Open SQL1, CONn
		
			if RS1.EOF then
				va_m3="&nbsp;"
			else
				va_m3=RS1(m_cons)				
			end if		
	Calcula_Media_T_F_F_N=va_m3

end function

Function Calcula_Media_F_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia_pr, codigo_materia, caminho_nota, tb_nota, periodo)
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CONn = Server.CreateObject("ADODB.Connection") 
	ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONn.Open ABRIRn
	
		if periodo=1 then
			m_cons="VA_Mc1"
		elseif periodo=2 then
			m_cons="VA_Mc2"
		elseif periodo=3 then
			m_cons="VA_Mc3"
		elseif periodo=4 then
			m_cons="VA_Mfinal"
		elseif periodo=5 then
			m_cons="VA_Media3"
		elseif periodo=6 then
			m_cons="VA_Media3"
		end if			
		
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM "&tb_nota&" where CO_Matricula ="& cod_aluno &" AND CO_Materia_Principal ='"& codigo_materia_pr &"' AND CO_Materia ='"& codigo_materia &"'"
		RS1.Open SQL1, CONn
		
			if RS1.EOF then
				va_m3="&nbsp;"
			else
				va_m3=RS1(m_cons)				
			end if		
	Calcula_Media_F_T_F_N=va_m3

end function


'===========================================================================================================================================
Function Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia_pr, vetor_materia, caminho_nota, tb_nota, periodo)	
'anulou="n"
'acumula=0
'divisor=0
'	Set CON0 = Server.CreateObject("ADODB.Connection") 
'	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
'	CON0.Open ABRIR0
'	
'	Set CONn = Server.CreateObject("ADODB.Connection") 
'	ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
'	CONn.Open ABRIRn			
'			
'	co_materia_mae_fil= split(vetor_materia,"#!#")
'	media_mae_acumula=0						
'	for j=0 to ubound(co_materia_mae_fil)			
'		disciplina_filha=co_materia_mae_fil(j)	
'		
'		Set RS = Server.CreateObject("ADODB.Recordset")
'		SQL = "SELECT * FROM TB_Programa_Aula where CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Materia ='"&disciplina_filha &"'"
'		RS.Open SQL, CON0	
'
'		peso=RS("NU_Peso")
'		divisor=divisor*1
'		if peso="" or isnull(peso) then
'			divisor=divisor+1
'		else	
'			peso=peso*1
'			divisor=divisor+peso
'		end if			
'			
'		media_aluno=Calcula_Media_F_T_F_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia_pr, disciplina_filha, caminho_nota, tb_nota, periodo)	
'			if media_aluno="" or isnull(media_aluno) or media_aluno="&nbsp;" then
'				anulou="s"
'			else
'				acumula=acumula*1	
'				media_aluno=media_aluno*1
'				acumula=acumula+media_aluno
'			end if					
'	next
'
'	if divisor =0 then
'		anulou="s"
'	end if	
'
'	if anulou="s" then
'		va_m3="&nbsp;"
'	else
'		va_m3=acumula/divisor
'		va_m3=arredonda(va_m3,"mat_dez",1,0)
'	end if
'
'Calcula_Media_T_T_F_N=va_m3		
end function











'===========================================================================================================================================
Function Calc_Med_T_F_T_N(unidade, curso, co_etapa, turma, cod_aluno, codigo_materia, caminho_nota, tb_nota, qtd_periodos, periodo_m2, periodo_m3,tipo_calculo, outro)


end function
%>