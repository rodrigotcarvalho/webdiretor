<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->
<!--#include file="../../../../inc/grava_notas.asp"-->
<%
excel=CAMINHO_upload&"resultados.xls"

ano_letivo = session("ano_letivo") 
chave=session("nvg")

Set CON0 = Server.CreateObject("ADODB.Connection") 
ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
CON0.Open ABRIR0	

Set CONa= Server.CreateObject("ADODB.Connection") 
ABRIRa = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
CONa.Open ABRIRa

Set CONg = Server.CreateObject("ADODB.Connection") 
ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
CONg.Open ABRIRg


Const adOpenStatic = 3
Const adLockPessimistic = 2

Dim cnnExcel
Dim rstExcel
Dim I
Dim iCols

Set cnnExcel = Server.CreateObject("ADODB.Connection")
cnnExcel.Open "DRIVER={Microsoft Excel Driver (*.xls)}; DBQ="&excel

Set rstExcel = Server.CreateObject("ADODB.Recordset")

'A referência [Plan1$] significa que a tabela é toda a planilha. 
'http://support.microsoft.com/kb/278973/pt-br
rstExcel.Open "SELECT * FROM [Plan1$]", cnnExcel,adOpenStatic

'  Response.Write "Colunas: <BR>"
iCols = rstExcel.Fields.Count
'For I = 0 To iCols - 1
'	Response.Write rstExcel.Fields.Item(I).Name & " - "
'Next
'
'  Response.Write "<BR>Linhas: <BR>"
'  rstExcel.MoveFirst
'  While Not rstExcel.EOF
'     For I = 0 To iCols - 1
'        Response.Write rstExcel.Fields.Item(I).Value & ","
'     Next
'     Response.Write "<BR>"
'     rstExcel.MoveNext
'  Wend
rstExcel.MoveFirst
While Not rstExcel.EOF
	m=rstExcel.Fields.Item(0).Value
	ma=rstExcel.Fields.Item(1).Value
	matric=rstExcel.Fields.Item(2).Value
	periodo=rstExcel.Fields.Item(3).Value
'	response.Write("<BR>"&m&" - "&ma&" - "&matric&" p "&periodo&"<BR>")
	
	if matric=9999 or matric = "" or isnull(matric) then
		ignora_aluno="s"	
	else		
		Set RSa  = Server.CreateObject("ADODB.Recordset")
		SQLa  = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& matric
	
		Set RSa  = CONa.Execute(SQLa )
	
		if RSa.EOF then
			ignora_aluno="s"
			'response.Redirect("index.asp?opt=err3&nvg="&chave&"&res="&matric&"$!$"&ano_letivo)	
		else
			ignora_aluno="n"
			unidade= RSa("NU_Unidade")
			curso= RSa("CO_Curso")
			etapa= RSa("CO_Etapa")
			turma= RSa("CO_Turma")
		end if
	end if
	
	if ignora_aluno="n" then	
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' AND IN_MAE=TRUE order by NU_Ordem_Boletim "
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
		end if
		
		vetor_materia_exibe=programa_aula(vetor_materia, unidade, curso, etapa, turma)	
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"'"
		Set RS = CONg.Execute(CONEXAO)
	
		nota = RS("TP_Nota")
		
		verifica_materia_programa=SPLIT(vetor_materia_exibe,"#!#")
		
		For i = 3 To iCols - 1
			achou=0
			variavel=rstExcel.Fields.Item(i).Name
		'response.Write(i&"D "&variavel&"<BR>")		
			for j=0 to ubound(verifica_materia_programa)
				
				if variavel=verifica_materia_programa(j) and achou=0 then
					disciplina=variavel
					achou=1
				end if	
			next
			
			if achou=1 then
				
				simulado=rstExcel.Fields.Item(i).Value
			'response.Write(i&" S "&simulado&"-"&matric&"<BR>")					
				if simulado="" or simulado=" " or isnull(simulado) then
				
				elseif simulado<2 then
					simulado=0
				elseif simulado<4 then	
					simulado=5
				elseif simulado<6 then	
					simulado=10
				elseif simulado<8 then	
					simulado=15	
				else
					simulado=20
				end if	
				
				'if disciplina="LP" then
				'	disciplina_pr="LP"	
				'	disciplina="POR3"
				'else							
			
					Set RSMT  = Server.CreateObject("ADODB.Recordset")
					SQL_MT  = "Select CO_Materia_Principal from TB_Materia WHERE CO_Materia = '"& disciplina&"'"
					Set RSMT  = CON0.Execute(SQL_MT)
							
					disciplina_pr = RSMT("CO_Materia_Principal")
							
					if Isnull(disciplina_pr) then
						disciplina_pr= disciplina
					end if		
				'end if
	'tb_a
	'"Faltas#!#PE_Teste#!#PE_Prova#!#VA_Teste1#!#VA_Teste2#!#VA_Teste3#!#VA_Teste4#!#VA_Prova1#!#VA_Prova2#!#VA_Prova3#!#VA_Bonus#!#VA_Rec"	
	'tb_b
	'"Faltas#!#PE_Teste#!#PE_Prova#!#VA_Teste1#!#VA_Teste2#!#VA_Prova1#!#VA_Simul#!#VA_Prova2#!#VA_Bonus#!#VA_Rec"
	'tb_c
	'"Faltas#!#PE_Teste#!#PE_Prova#!#VA_Teste1#!#VA_Teste2#!#VA_Teste3#!#VA_Teste4#!#VA_Prova1#!#VA_Prova2#!#VA_Bonus#!#VA_Rec"
				if nota ="TB_NOTA_A" then	
					caminho_n = CAMINHO_na
					submete=grava_nota(caminho_n,nota,matric,periodo,disciplina_pr,disciplina,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,0)
				elseif nota ="TB_NOTA_B" then	
					caminho_n = CAMINHO_nb	

					Set CON_N = Server.CreateObject("ADODB.Connection")
					ABRIR3 = "DBQ="& caminho_n & ";Driver={Microsoft Access Driver (*.mdb)}"
					CON_N.Open ABRIR3

					Set RS3 = Server.CreateObject("ADODB.Recordset")
					SQL_N = "Select * from "& nota &" WHERE CO_Matricula = "& matric & " AND CO_Materia_Principal = '"& disciplina_pr &"' AND CO_Materia = '"& disciplina &"' AND NU_Periodo="&periodo
					Set RS3 = CON_N.Execute(SQL_N)			 
				
					if RS3.EOF then 
						va_pt=NULL
						va_pp=NULL
						va_t1=NULL
						va_t2=NULL
						va_p1=NULL
						va_simul=n6
						va_p2=NULL
						va_bon=NULL
						va_rec=NULL				
					else
						va_pt=RS3("PE_Teste")
						va_pp=RS3("PE_Prova")
						va_t1=RS3("VA_Teste1")
						va_t2=RS3("VA_Teste2")
						va_p1=RS3("VA_Prova1")
						va_p2=RS3("VA_Prova2")
						va_bon=RS3("VA_Bonus")
						va_rec=RS3("VA_Rec")	
					end if					
					
							
					'response.Write(caminho_n&","&nota&","&matric&","&periodo&","&disciplina_pr&","&disciplina&",9999,"&va_pt&","&va_pp&","&va_t1&","&va_t2&","&va_p1&","&simulado&"<BR>")
					submete=grava_nota(caminho_n,nota,matric,periodo,disciplina_pr,disciplina,9999,va_pt,va_pp,va_t1,va_t2,va_p1,simulado,va_p2,va_bon,va_rec,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,0)
					
				elseif nota ="TB_NOTA_C" then	
					caminho_n = CAMINHO_nc	
					submete=grava_nota(caminho_n,nota,matric,periodo,disciplina_pr,disciplina,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,9999,0)
				end if
			
				if submete<>"ok" then
					response.Redirect("index.asp?opt=err4&nvg="&chave&"&res="&submete)	
				end if		
			end if
		Next	
	end if
rstExcel.MoveNext
Wend

rstExcel.Close
Set rstExcel = Nothing

cnnExcel.Close
Set cnnExcel = Nothing

SET FSO = Server.CreateObject("Scripting.FileSystemObject")
FSO.deletefile(excel) 


response.Redirect("index.asp?opt=ok&nvg="&chave)	
%>
