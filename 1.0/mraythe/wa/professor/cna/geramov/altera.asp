<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<%
ano_letivo = session("ano_letivo") 
ano_atual = DatePart("yyyy", now)

chave=session("chave")
session("chave")=chave
nome = session("nome") 
unidade = request.Form("unidade")
curso = request.form("curso")
etapa = request.Form("etapa")
periodo = request.Form("periodo")
dia_de= request.Form("dia_de")
mes_de = request.Form("mes_de")
dia_ate = request.Form("dia_ate")
mes_ate = request.Form("mes_ate")
data_de = dia_de&"/"&mes_de&"/"&ano_letivo
data_ate = dia_ate&"/"&mes_ate&"/"&ano_atual

session("unidade_trabalho")=unidade
session("curso_trabalho")=curso
session("etapa_trabalho")=etapa
session("periodo_trabalho")=periodo
session("dia_de_trabalho")=dia_de
session("mes_de_trabalho")=mes_de
session("dia_ate_trabalho")=dia_ate
session("mes_ate_trabalho")=mes_ate

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
			caminho_gera_mov = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\mraythe\BD\"
arquivo="MW"&data&".txt"

Set fs = CreateObject("Scripting.FileSystemObject") 'cria  
Set d = fs.CreateTextFile(caminho_gera_mov&arquivo, False) 
'd.write("teste")  
'd.writeblanklines(5)  
'd.writeline("deixei 5 linhas em branco")  
'd.close()  
'd.writeline("Início do Arquivo de Notas")  		

	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO_g  & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR

	Set CONA = Server.CreateObject("ADODB.Connection") 
	ABRIRA = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONA.Open ABRIRA
	
	Set CONPR = Server.CreateObject("ADODB.Connection") 
	ABRIRPR = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONPR.Open ABRIRPR		



	Set RS = Server.CreateObject("ADODB.Recordset")
	CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidade &" And CO_Curso= '"& curso &"' And CO_Etapa = '"& etapa &"'"
	Set RS = CON.Execute(CONEXAO)

'response.Write(CONEXAO)		
	
	nota_i = RS("TP_Nota")

	if nota_i="TB_NOTA_A" then
	CAMINHO_n=CAMINHO_na
	elseif nota_i="TB_NOTA_B" then
			CAMINHO_n=CAMINHO_nb
	elseif nota_i="TB_NOTA_C" then
			CAMINHO_n=CAMINHO_nc
	end if
			
	Set CON2 = Server.CreateObject("ADODB.Connection") 
	ABRIR2 = "DBQ="& CAMINHO_n & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON2.Open ABRIR2

	Set RSA= Server.CreateObject("ADODB.Recordset")
	CONEXAOA = "Select * from TB_Aluno_Esta_Turma WHERE NU_Unidade = "& unidade &" And CO_Curso = '"& curso &"' And CO_Etapa = '"& etapa &"' order by NU_Chamada"
	Set RSA = CONA.Execute(CONEXAOA)

	while not RSA.EOF
		ult_matric=""
		ult_materia=""
		mat = RSA("CO_Matricula")
		turma = RSA("CO_Turma")
		
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		CONEXAO3 = "Select * from "&nota_i&" WHERE CO_Matricula = "& mat&" AND NU_Periodo="&periodo&" AND (DA_Ult_Acesso BETWEEN #"&data_de&"# AND #"&data_ate&"#) order by CO_Matricula, CO_Materia_Principal, CO_Materia " 
		Set RS3 = CON2.Execute(CONEXAO3)
		

		
		while not RS3.eof
			materia=RS3("CO_Materia_principal")
			pula_linha="N"
			sem_nota1 = "n"
			sem_nota2 = "n"
			sem_nota3 = "n"
			if ult_matric = mat and ult_materia = materia then
				pula_linha="S"		
			else	
				ult_matric = mat 
				ult_materia = materia	
						
				Set RS_mat = Server.CreateObject("ADODB.Recordset")
				SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& materia &"'  order by NU_Ordem_Boletim"			
				'RESPONSE.Write(SQL&"<br>")			
				RS_mat.Open SQL, CONPR
			
				mae= RS_mat("IN_MAE")
				fil= RS_mat("IN_FIL")
				in_co= RS_mat("IN_CO")
				peso= RS_mat("NU_Peso")
				
				'RESPONSE.Write(pula_linha&"<br>")
										
				if (mae=TRUE and fil=FALSE and in_co=TRUE and isnull(peso)) then
	
					Set RS1a = Server.CreateObject("ADODB.Recordset")
					SQL1a = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& materia &"' order by NU_Ordem_Boletim"
				'RESPONSE.Write(SQL1a&"<br>")				
					RS1a.Open SQL1a, CONPR
						
					if RS1a.EOF then
					else
						co_materia_fil_check=1 
						peso_acumula=0
						fal_acumula=0
						va_m1_acumula=0
						va_m2_acumula=0
						va_m3_acumula=0
						rec_acumula=0																				
						bon_acumula=0
				
						while not RS1a.EOF
							co_mat_fil= RS1a("CO_Materia")
							
							Set RSp2 = Server.CreateObject("ADODB.Recordset")
							SQLp2 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia = '"& co_mat_fil &"' order by NU_Ordem_Boletim"
							'response.Write(SQLp2&"<BR>")						
							RSp2.Open SQLp2, CONPR	
													
							nu_peso_fil=RSp2("NU_Peso")	
										
							peso_acumula=peso_acumula+nu_peso_fil
														
							Set RS4 = Server.CreateObject("ADODB.Recordset")
							SQL_4 = "Select * from "& nota_i &" WHERE CO_Matricula = "& mat &" AND CO_Materia = '"& co_mat_fil &"' AND CO_Materia_Principal = '"& materia &"' AND NU_Periodo="&periodo
							'response.Write(SQL_4&"<BR>")
							Set RS4 = CON2.Execute(SQL_4)						
					

							if RS4.EOF then
								fal_temp=""
								va_m1_temp=""
								va_bon_temp=""
								va_m2_temp=""
								va_rec_temp=""
								va_m3_temp=""
							else					
								fal_temp=RS4("NU_Faltas")
								va_m1_temp=RS4("VA_Media1")
								va_bon_temp=RS4("VA_Extra")
								va_m2_temp=RS4("VA_Media2")
								va_rec_temp=RS4("VA_Rec")
								va_m3_temp=RS4("VA_Media3")
							end if
							'response.Write(va_m1_temp&"-"&va_m2_temp&"-"&va_m3_temp&"<BR>")							
	
							if isnull(fal_temp) or fal_temp="" then
								sem_notaf="s"
							else
								fal_acumula=fal_acumula+fal_temp								
							end if	
							if isnull(va_m1_temp) or va_m1_temp="" then
								sem_nota1="s"
							else
								va_m1_acumula=va_m1_acumula+va_m1_temp								
							end if	
							if isnull(va_m2_temp) or va_m2_temp="" then
								sem_nota2="s"
							else
								va_m2_acumula=va_m2_acumula+va_m2_temp								
							end if	
							if isnull(va_m3_temp) or va_m3_temp="" then
								sem_nota3="s"
							else
								va_m3_acumula=va_m3_acumula+va_m3_temp								
							end if													
							if isnull(va_rec_temp) or va_rec_temp="" then
								sem_notar="s"
							else
								rec_acumula=rec_acumula+va_rec_temp								
							end if	
							if isnull(va_bon_temp) or va_bon_temp="" then
								sem_notab="s"
							else
								bon_acumula=bon_acumula+va_bon_temp								
							end if	
																													
						RS1a.MOVENEXT
						wend
						
						'response.Write(sem_nota1&"-"&sem_nota2&"-"&sem_nota3&"<BR>")						

						if sem_notaf="s" then
							fal=""
						else	
							fal=fal_acumula								
						end if					
						if sem_nota1="s" then
							va_m1=""
						else	
							va_m1=va_m1_acumula/peso_acumula
							va_m1=va_m1*10
								decimo = va_m1 - Int(va_m1)
								If decimo >= 0.5 Then
									nota_arredondada = Int(va_m1) + 1
									va_m1=nota_arredondada
								else
									nota_arredondada = Int(va_m1)
									va_m1=nota_arredondada											
								End If
							va_m1=va_m1/10	
							va_m1 = formatNumber(va_m1,1)									
						end if
						
						if sem_nota2="s" then
							va_m2=""
						else	
							va_m2=va_m2_acumula/peso_acumula
							va_m2=va_m2*10
								decimo = va_m2 - Int(va_m2)
								If decimo >= 0.5 Then
									nota_arredondada = Int(va_m2) + 1
									va_m2=nota_arredondada
								else
									nota_arredondada = Int(va_m2)
									va_m2=nota_arredondada											
								End If
							va_m2=va_m2/10	
							va_m2 = formatNumber(va_m2,1)									
						end if
	
						if sem_nota3="s" then
							va_m3=""
						else	
							va_m3=va_m3_acumula/peso_acumula
							va_m3=va_m3*10
								decimo = va_m3 - Int(va_m3)
								If decimo >= 0.5 Then
									nota_arredondada = Int(va_m3) + 1
									va_m3=nota_arredondada
								else
									nota_arredondada = Int(va_m3)
									va_m3=nota_arredondada											
								End If
							va_m3=va_m3/10	
							va_m3 = formatNumber(va_m3,1)									
						end if
					end if	
					
					if sem_notar="s" then
						va_rec=""
					else	
						va_rec=rec_acumula/peso_acumula
						va_rec=va_rec*10
							decimo = va_rec - Int(va_rec)
							If decimo >= 0.5 Then
								nota_arredondada = Int(va_rec) + 1
								va_rec=nota_arredondada
							else
								nota_arredondada = Int(va_rec)
								va_rec=nota_arredondada											
							End If
						va_rec=va_rec/10	
						va_rec = formatNumber(va_rec,1)									
					end if				
				
				
					if sem_notab="s" then
						va_bon=""
					else	
						va_bon=bon_acumula/peso_acumula
						va_bon=va_bon*10
							decimo = va_bon - Int(va_bon)
							If decimo >= 0.5 Then
								nota_arredondada = Int(va_bon) + 1
								va_bon=nota_arredondada
							else
								nota_arredondada = Int(va_bon)
								va_bon=nota_arredondada											
							End If
						va_bon=va_bon/10	
						va_bon = formatNumber(va_bon,1)									
					end if				
				elseif (mae=FALSE ) then
					pula_linha="S"
				else		
					fal=RS3("NU_Faltas")
					va_m1=RS3("VA_Media1")
					va_bon=RS3("VA_Extra")
					va_m2=RS3("VA_Media2")
					va_rec=RS3("VA_Rec")
					va_m3=RS3("VA_Media3")
				end if	
			END IF	
			IF pula_linha="N" THEN
				d.writeLine mat&";"&materia&";"&periodo&";"&fal&";"&va_m1&";"&va_bon&";"&va_m2&";"&va_rec&";"&va_m3		
			END IF		
			'response.Write(nota_i&" - "&wrt&"<BR>")
		
		RS3.MOVENEXT						
		wend
	RSA.MOVENEXT						
	wend
'd.writeline("Fim do Arquivo")  
'd.writeblanklines(5)  
'd.writeline("Início do Arquivo de Faltas")
  
	Set RSA2= Server.CreateObject("ADODB.Recordset")
	CONEXAOA = "Select * from TB_Aluno_Esta_Turma WHERE NU_Unidade = "& unidade &" And CO_Curso = '"& curso &"' And CO_Etapa = '"& etapa &"' order by NU_Chamada"
	Set RSA2 = CONA.Execute(CONEXAOA)

	while not RSA2.EOF
		mat = RSA2("CO_Matricula")			
		Set RS32 = Server.CreateObject("ADODB.Recordset")
		CONEXAO32 = "Select * from TB_Frequencia_Periodo WHERE CO_Matricula = "& mat 
		Set RS32 = CON2.Execute(CONEXAO32)

		while not RS3.eof
			va_f1=RS32("NU_Faltas_P1")
			va_f2=RS32("NU_Faltas_P2")
			va_f3=RS32("NU_Faltas_P3")
			va_f4=RS32("NU_Faltas_P4")
	d.writeLine mat&";FREQ;"&va_f1&";"&va_f2&";"&va_f3&";"&va_f4			
	'response.Write(nota_i&" - "&wrt&"<BR>")
	
		RS32.MOVENEXT						
		wend
	RSA2.MOVENEXT						
	wend
'response.End()
'd.writeline("Fim do arquivo")  				
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