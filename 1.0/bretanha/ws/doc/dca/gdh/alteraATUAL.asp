<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%
Server.ScriptTimeout = 300 'valor em segundos
ano_letivo = session("ano_letivo") 
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
			
arquivo="Historico"&data&".txt"
Set fs = CreateObject("Scripting.FileSystemObject") 'cria  
Set d = fs.CreateTextFile(caminho_gera_mov&arquivo, False) 	


	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO_g  & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR

	Set CONA = Server.CreateObject("ADODB.Connection") 
	ABRIRA = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONA.Open ABRIRA
	
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0		
	
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

	tb_nota=tabela_notas(CON, unidade_t, curso_t, co_etapa_t, turma_t, 0, 0, 0)

	caminho_nota=caminho_notas(CON, tb_nota, outro)	
	
	matriculas=alunos_esta_turma(CONA, ano_letivo, "CO_Matricula", unidade_t, curso_t, co_etapa_t, turma_t, "*", "NU_Chamada", outro)	
	
	Set RS5 = Server.CreateObject("ADODB.Recordset")
	SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa_t &"' AND CO_Curso ='"& curso_t &"' order by NU_Ordem_Boletim "
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
		
		'vetor_materia_exibe=programa_aula(vetor_materia, unidade, curso, co_etapa, "nulo")			
		'response.Write(vetor_materia_media&"<BR>")
	end if		

	co_materia_exibe=Split(vetor_materia,"#$#")		
	
	Set CON_N = Server.CreateObject("ADODB.Connection") 
	ABRIR_N = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_N.Open ABRIR_N
	

	
	if matriculas="" or isnull(matriculas) then
	else
		vetor_matriculas = split(matriculas, "#!#")
		for vm=0 to ubound(vetor_matriculas)
			
			co_matricula = vetor_matriculas(vm)					
		
			resultados=Calc_Med_An_Fin(unidade_t, curso_t, co_etapa_t, turma_t, co_matricula, vetor_materia_media, caminho_nota, tb_nota, 5, 4, 5, "final", 0)	

			'resultado_final_aluno=apura_resultado_aluno(curso_t, co_etapa_t, resultados)	
			resultado_final_aluno=apura_resultado_aluno2(co_matricula, curso_t, co_etapa_t, resultados,vetor_materia_media)			

			if resultado_final_aluno="&nbsp;" or resultado_final_aluno="" or isnull(resultado_final_aluno) then	
				resultado_final_aluno = ""	
			elseif resultado_final_aluno = "APR" or resultado_final_aluno = "Apr"  then
				resultado_final_aluno = "A"	
			elseif resultado_final_aluno = "APC" or resultado_final_aluno = "Apc"  then
				resultado_final_aluno = "C"					
			elseif resultado_final_aluno = "PFI" or resultado_final_aluno = "Pfi"  then
				resultado_final_aluno = "P"	
			elseif resultado_final_aluno = "REC" or resultado_final_aluno = "Rec"  then
				resultado_final_aluno = "E"		
			elseif resultado_final_aluno = "ECE" or resultado_final_aluno = "ECE"  then
				resultado_final_aluno = "E"	
			elseif resultado_final_aluno = "REP" or resultado_final_aluno = "Rep"  then
				resultado_final_aluno = "R"				
			elseif resultado_final_aluno = "AP.D" or resultado_final_aluno = "AP.D"  then
				resultado_final_aluno = "D"																																					
			end if					
			
			
			resultado_apurado= split(resultados, "#$#" )							
			carga_total = 0
			soma_total_faltas = 0
			frequencia_total = 0
			for co=0 to ubound(co_materia_exibe)
				dados_materia_exibe=Split(co_materia_exibe(co),"#!#")		
				materia=dados_materia_exibe(0)
				carga_materia=dados_materia_exibe(1)	

				resultado_disciplina= split(resultado_apurado(co), "#!#" )					
				
				if resultado_disciplina(0)="&nbsp;" or resultado_disciplina(0)="" or isnull(resultado_disciplina(0)) then
					calcula_frequencia="n"		
					media = ""
				else
					media = resultado_disciplina(0)	
				end if	
				
				if resultado_disciplina(1)="&nbsp;" or resultado_disciplina(1)="" or isnull(resultado_disciplina(1)) then
					calcula_frequencia="n"		
					resultado_disciplina = ""	
				elseif resultado_disciplina(1) = "APR" or resultado_disciplina(1) = "Apr"  then
					resultado_disciplina = "A"	
				elseif resultado_disciplina(1) = "PFI" or resultado_disciplina(1) = "Pfi"  then
					resultado_disciplina = "P"	
				elseif resultado_disciplina(1) = "REC" or resultado_disciplina(1) = "Rec"  then
					resultado_disciplina = "E"		
				elseif resultado_disciplina(1) = "ECE" or resultado_disciplina(1) = "ECE"  then
					resultado_disciplina = "E"	
				elseif resultado_disciplina(1) = "REP" or resultado_disciplina(1) = "Rep"  then
					resultado_disciplina = "R"				
				elseif resultado_disciplina(1) = "AP.D" or resultado_disciplina(1) = "AP.D"  then
					resultado_disciplina = "D"																																					
				end if			
				
					Set RS_N = Server.CreateObject("ADODB.Recordset")
					SQL_N = "SELECT SUM("&tb_nota&".NU_Faltas) AS faltas_disciplina FROM "&tb_nota&" where CO_Matricula ="& co_matricula&" and CO_Materia ='"& materia&"'"
					RS_N.Open SQL_N, CON_N
					
					if RS_N.EOF then
						faltas_disciplina=0
					else
						faltas_disciplina=RS_N("faltas_disciplina")
						
						if isnull(RS_N("faltas_disciplina")) or faltas_disciplina="" then
							faltas_disciplina=0						
						end if
					end if	
					faltas_disciplina=faltas_disciplina*1
					carga_materia=carga_materia*1
					frequencia=((carga_materia-faltas_disciplina)/carga_materia)*100
					if frequencia<100 then
						frequencia=arredonda(frequencia,"mat_dez",1,0)	
					end if	
					soma_total_faltas=soma_total_faltas+faltas_disciplina						
					carga_total = carga_total+carga_materia
'				if curso=1 and co_etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then									
'					teste_media = isnumeric(media)							
'					if teste_media=TRUE then							
'						if media > 90 then
'						conceito="E"
'						elseif (media > 70) and (media <= 90) then
'						conceito="MB"
'						elseif (media > 60) and (media <= 70) then							
'						conceito="B"
'						elseif (media > 49) and (media <= 60) then
'						conceito="R"
'						else							
'						conceito="I"
'						end if	
'					end if	
'				else
					conceito=media				
'				end if	
								
'				if calcula_frequencia="s" then
'					Set RSF = Server.CreateObject("ADODB.Recordset")
'					SQL_N = "Select * from TB_Frequencia_Periodo WHERE CO_Matricula = "& co_matricula
'					Set RSF = CON_N.Execute(SQL_N)
'					soma_faltas=0			
'					
'					if RSF.eof THEN
'						f1="&nbsp;"
'						f2="&nbsp;"
'						f3="&nbsp;"
'						f4="&nbsp;"	
'					else	
'						f1=RSF("NU_Faltas_P1")
'						f2=RSF("NU_Faltas_P2")
'						f3=RSF("NU_Faltas_P3")
'						f4=RSF("NU_Faltas_P4")		
'						
'						if isnull(f1) or f1= "" then
'						else
'							f1=f1*1
'							soma_faltas=soma_faltas*1
'							soma_faltas=soma_faltas+f1		
'						end if
'						
'						if isnull(f2) or f2= "" then
'						else
'							f2=f2*1
'							soma_faltas=soma_faltas*1
'							soma_faltas=soma_faltas+f2		
'						end if
'						
'						if isnull(f3) or f3= "" then
'						else
'							f3=f3*1
'							soma_faltas=soma_faltas*1
'							soma_faltas=soma_faltas+f3		
'						end if
'						
'						if isnull(f4) or f4= "" then
'						else
'							f4=f4*1
'							soma_faltas=soma_faltas*1
'							soma_faltas=soma_faltas+f4		
'						end if									
'					END IF				
'											
'					soma_faltas=soma_faltas*1
'					dias_de_aula_no_ano=200
'					
'					frequencia=((dias_de_aula_no_ano-soma_faltas)/dias_de_aula_no_ano)*100
'					if frequencia<100 then
'						frequencia=arredonda(frequencia,"mat_dez",1,0)	
'
'					end if	
'				else
'					frequencia=""
'				end if			
				nome_materia=GeraNomesNovaVersao("D",materia,variavel2,variavel3,variavel4,variavel5,CON0,outro)			
		
				d.writeLine ano_letivo&";"&co_matricula&";"&nome_materia&";"&carga_materia&";"&frequencia&";"&conceito&";"&resultado_disciplina&";"&resultado_final_aluno
			NEXT		
									
'				Set RSd = Server.CreateObject("ADODB.Recordset")
'				SQLd = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa_t &"' AND CO_Curso ='"& curso_t &"'"
'				RSd.Open SQLd, CON0
'		
'				tipo_freq=RSd("IN_Frequencia")						
'
'				if tipo_freq="D" then									
'								
'					Set RSF = Server.CreateObject("ADODB.Recordset")
'					SQL_N = "Select * from TB_Frequencia_Periodo WHERE CO_Matricula = "& co_matricula
'					Set RSF = CON_N.Execute(SQL_N)
'					
'
'					
'					if RSF.eof THEN
'						f1="&nbsp;"
'						f2="&nbsp;"
'						f3="&nbsp;"
'						f4="&nbsp;"	
'					else	
'						f1=RSF("NU_Faltas_P1")
'						f2=RSF("NU_Faltas_P2")
'						f3=RSF("NU_Faltas_P3")
'						f4=RSF("NU_Faltas_P4")		
'						
'						if isnull(f1) or f1= "" then
'						else
'							f1=f1*1
'							soma_total_faltas=soma_total_faltas*1
'							soma_total_faltas=soma_total_faltas+f1		
'						end if
'						
'						if isnull(f2) or f2= "" then
'						else
'							f2=f2*1
'							soma_total_faltas=soma_total_faltas*1
'							soma_total_faltas=soma_total_faltas+f2		
'						end if
'						
'						if isnull(f3) or f3= "" then
'						else
'							f3=f3*1
'							soma_total_faltas=soma_total_faltas*1
'							soma_total_faltas=soma_total_faltas+f3		
'						end if
'						
'						if isnull(f4) or f4= "" then
'						else
'							f4=f4*1
'							soma_total_faltas=soma_total_faltas*1
'							soma_total_faltas=soma_total_faltas+f4		
'						end if									
'					END IF				
'				else
'					Set RS1 = Server.CreateObject("ADODB.Recordset")
'					SQL1 = "SELECT SUM(NU_Faltas) as Total_Faltas FROM "&tb_nota&" where CO_Matricula ="& co_matricula
'					RS1.Open SQL1, CON_N	
'
'					if RS1.EOF then
'					else
'						soma_total_faltas=RS1("Total_Faltas")	
'						if isnull(soma_total_faltas) or soma_total_faltas="" then
'							soma_total_faltas=0
'						end if
'					end if									
'				end if	
				soma_total_faltas=soma_total_faltas*1			
				
'				if curso=1 then
'					co_etapa=co_etapa*1
'					if co_etapa<6 then
'						horas_aula=920
'					else
'						horas_aula=1080						
'					end if
'				end if
				horas_aula=carga_total
				horas_aula=horas_aula*1
				frequencia_final=((horas_aula-soma_total_faltas)/horas_aula)*100	

				if frequencia_final<100 then
					frequencia_final=arredonda(frequencia_final,"mat_dez",1,0)	
				end if			
				
				d.writeLine ano_letivo&";"&co_matricula&";FREQUENCIA;"&frequencia_final&";;;;"									
		NEXT
	End if	
NEXT
'			response.Write("<BR>"&dados_arquivo&"<BR>")
'			response.End()

d.close

response.Redirect("download.asp?opt="&arquivo)
		
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