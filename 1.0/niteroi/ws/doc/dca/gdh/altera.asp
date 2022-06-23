<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/parametros.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->
<!--#include file="../../../../inc/funcoes7.asp"-->
<!--#include file="../../../../inc/calculos.asp"-->
<!--#include file="../../../../inc/resultados.asp"-->
<%
Server.ScriptTimeout = 1200 'valor em segundos
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


'Apagar os arquivos antigos---------------------------------------------------------------
set FSD = CreateObject("Scripting.FileSystemObject")
set folder = FSD.getFolder (caminho_gera_mov)   
for each file in folder.files
	if (dateDiff("n", file.datecreated, now) >30) then
		File.delete
	end if
next
'-----------------------------------------------------------------------------------------

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
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR	

	Set CONA = Server.CreateObject("ADODB.Connection") 
	ABRIRA = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONA.Open ABRIRA
	
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CON2 = Server.CreateObject("ADODB.Connection") 
	ABRIR2 = "DBQ="& CAMINHO_g  & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON2.Open ABRIR2			
	
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
'response.Flush()
	dados_select=split(conjunto_dados(i),"#!#")
	unidade_t=dados_select(0)
	curso_t=dados_select(1)
	co_etapa_t=dados_select(2)
	turma_t=dados_select(3)			
	'response.Write(unidade_t&", "&curso_t&", "&co_etapa_t&", "&turma_t&"<BR>")
	tb_nota=tabela_notas(CON2, unidade_t, curso_t, co_etapa_t, turma_t, 0, 0, 0)
	'response.Write(tb_nota&"<BR>")	
	if tb_nota = "" or isnull(tb_nota) then
		inclui_etapa = "N"
	else
		caminho_nota=caminho_notas(CON2, tb_nota, outro)
	'response.Write(caminho_nota&"<BR>")					
		if caminho_nota = "ERRO" then
			inclui_etapa = "N"
		else
			inclui_etapa = "S"		
		end if		
	end if
	'response.Write(inclui_etapa&"<BR>")		
	if inclui_etapa = "S" then
		matriculas=alunos_esta_turma(CONA, ano_letivo, "CO_Matricula", unidade_t, curso_t, co_etapa_t, turma_t, "*", "NU_Chamada", outro)	
		
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa_t &"' AND CO_Curso ='"& curso_t &"' and IN_MAE = TRUE order by NU_Ordem_Boletim "
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
				horas_aula=0
				co_matricula = vetor_matriculas(vm)		
				
				tp_modelo=tipo_divisao_ano(curso_t,co_etapa_t,"tp_modelo")			
				prd_ter_media=Periodo_Media(tp_modelo,"MF",outro)
				
				resultado_final_aluno=novo_apura_resultado_aluno(curso_t,co_etapa_t,co_matricula,vetor_materia_media,caminho_nota,tb_nota,prd_ter_media,"final",outro)					
	
				if resultado_final_aluno="&nbsp;" or resultado_final_aluno="" or isnull(resultado_final_aluno) then	
					resultado_final_aluno = ""	
				elseif resultado_final_aluno = "APR" or resultado_final_aluno = "Apr"  then
					resultado_final_aluno = "A"	
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
				
				
				resultado_apurado= split(resultados_calculados, "#$#" )							
	
				for com=0 to ubound(co_materia_exibe)
					dados_materia_exibe=Split(co_materia_exibe(com),"#!#")		
					materia=dados_materia_exibe(0)
					carga_materia=dados_materia_exibe(1)	
					
					resultado_disciplina= Calc_Ter_Media (unidade_t, curso_t, co_etapa_t, turma_t, co_matricula, materia, caminho_nota, tb_nota, prd_ter_media, "sem_calculo", outro)
			
	
					resultados_materia = SPLIT(resultado_disciplina,"#!#")	
						
					Set RSm = Server.CreateObject("ADODB.Recordset")
					SQLm = "SELECT * FROM TB_Materia where CO_Materia ='"& materia &"'"
					RSm.Open SQLm, CON0				
					
					disc_obrigatoria=RSm("IN_Obrigatorio")
					
					if disc_obrigatoria=FALSE then
						if resultados_materia(0)="&nbsp;" or resultados_materia(0)="" or isnull(resultados_materia(0)) then
							inclui_disciplina="N"
							media = ""
						else
							inclui_disciplina="S"							
							media = resultados_materia(0)	
						end if						
					else
						inclui_disciplina="S"						
						if resultados_materia(0)="&nbsp;" or resultados_materia(0)="" or isnull(resultados_materia(0)) then	
							media = ""
						else
							media = resultados_materia(0)	
						end if						
					end if			
				
					if inclui_disciplina="S" then
	
						
						if resultados_materia(1)="&nbsp;" or resultados_materia(1)="" or isnull(resultados_materia(1)) then	
							resultados_materia = ""	
						elseif resultados_materia(1) = "APR" or resultados_materia(1) = "Apr"  then
							resultados_materia = "A"	
						elseif resultados_materia(1) = "PFI" or resultados_materia(1) = "Pfi"  then
							resultados_materia = "P"	
						elseif resultados_materia(1) = "REC" or resultados_materia(1) = "Rec"  then
							resultados_materia = "E"		
						elseif resultados_materia(1) = "ECE" or resultados_materia(1) = "ECE"  then
							resultados_materia = "E"	
						elseif resultados_materia(1) = "REP" or resultados_materia(1) = "Rep"  then
							resultados_materia = "R"				
						elseif resultados_materia(1) = "AP.D" or resultados_materia(1) = "AP.D"  then
							resultados_materia = "D"																																					
						end if								
					
						Set RS_N = Server.CreateObject("ADODB.Recordset")
						SQL_N = "SELECT SUM("&tb_nota&".NU_Faltas) AS total_faltas FROM "&tb_nota&" where CO_Matricula ="& co_matricula&" and CO_Materia ='"& materia&"'"
						RS_N.Open SQL_N, CON_N
						
						if RS_N.EOF then
							total_faltas=0
						else
							total_faltas=RS_N("total_faltas")
							
							if isnull(RS_N("total_faltas")) or total_faltas="" then
								total_faltas=0						
							end if
						end if	
						total_faltas=total_faltas*1
						carga_materia=carga_materia*1
						frequencia=((carga_materia-total_faltas)/carga_materia)*100
						if frequencia<100 then
							frequencia=arredonda(frequencia,"mat_dez",1,0)	
						end if	
'						


											
					
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
						nome_materia=GeraNomesNovaVersao("D",materia,variavel2,variavel3,variavel4,variavel5,CON0,outro)											

						d.writeLine ano_letivo&";"&co_matricula&";"&nome_materia&";"&carga_materia&";"&frequencia&";"&conceito&";"&resultados_materia&";"&resultado_final_aluno
						horas_aula = horas_aula+carga_materia
					end if
				NEXT	
				
				soma_total_faltas=0		
									
				Set RSd = Server.CreateObject("ADODB.Recordset")
				SQLd = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa_t &"' AND CO_Curso ='"& curso_t &"'"
				RSd.Open SQLd, CON0
		
				tipo_freq=RSd("IN_Frequencia")						

				if tipo_freq="D" then									
								
					Set RSF = Server.CreateObject("ADODB.Recordset")
					SQL_N = "Select * from TB_Frequencia_Periodo WHERE CO_Matricula = "& co_matricula
					Set RSF = CON_N.Execute(SQL_N)
					

					
					if RSF.eof THEN
						f1="&nbsp;"
						f2="&nbsp;"
						f3="&nbsp;"
						f4="&nbsp;"	
					else	
						f1=RSF("NU_Faltas_P1")
						f2=RSF("NU_Faltas_P2")
						f3=RSF("NU_Faltas_P3")
						f4=RSF("NU_Faltas_P4")		
						
						if isnull(f1) or f1= "" then
						else
							f1=f1*1
							soma_total_faltas=soma_total_faltas*1
							soma_total_faltas=soma_total_faltas+f1		
						end if
						
						if isnull(f2) or f2= "" then
						else
							f2=f2*1
							soma_total_faltas=soma_total_faltas*1
							soma_total_faltas=soma_total_faltas+f2		
						end if
						
						if isnull(f3) or f3= "" then
						else
							f3=f3*1
							soma_total_faltas=soma_total_faltas*1
							soma_total_faltas=soma_total_faltas+f3		
						end if
						
						if isnull(f4) or f4= "" then
						else
							f4=f4*1
							soma_total_faltas=soma_total_faltas*1
							soma_total_faltas=soma_total_faltas+f4		
						end if									
					END IF				
				else
					Set RS1 = Server.CreateObject("ADODB.Recordset")
					SQL1 = "SELECT SUM(NU_Faltas) as Total_Faltas FROM "&tb_nota&" where CO_Matricula ="& co_matricula
					RS1.Open SQL1, CON_N	

					if RS1.EOF then
					else
						soma_total_faltas=RS1("Total_Faltas")	
						if isnull(soma_total_faltas) or soma_total_faltas="" then
							soma_total_faltas=0
						end if
					end if									
				end if	
				soma_total_faltas=soma_total_faltas*1			
				
				if curso=1 then
					co_etapa=co_etapa*1
					if co_etapa<6 then
						horas_aula=920
					else
						horas_aula=1080						
					end if
				end if

				frequencia_final=((horas_aula-soma_total_faltas)/horas_aula)*100	

				if frequencia_final<100 then
					frequencia_final=arredonda(frequencia_final,"mat_dez",1,0)	
				end if			
				
				d.writeLine ano_letivo&";"&co_matricula&";FREQUENCIA:"&frequencia_final&";;;;;"										
			NEXT
		end if	
	End if	
NEXT
'			response.Write("<BR>"&dados_arquivo&"<BR>")
'			response.End()

d.close
set d=nothing 
set fs=nothing 

Dim oFSO
	Dim oFile
	Dim sSourceFile

	Set oFSO = CreateObject("Scripting.FileSystemObject")

	sSourceFile = caminho_gera_mov&arquivo

	Set oFile = oFSO.GetFile(sSourceFile)


	'response.Flush()
	
	'if oFile.size>5000 then 
	'response.Redirect("index.asp?nvg="&chave&"&fl="&arquivo&"&opt=ok2")
	'else
	response.Redirect("download.asp?opt="&arquivo)
	'end if		

	' Clean Up
	Set oFile = Nothing
	Set oFSO = Nothing
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