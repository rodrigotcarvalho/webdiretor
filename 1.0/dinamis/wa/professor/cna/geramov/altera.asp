<%On Error Resume Next%>
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
			
arquivo="MW"&data&".txt"
response.Write(caminho_gera_mov&arquivo&",,,")
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
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidade &"And CO_Curso= '"& curso &"' And CO_Etapa = '"& etapa &"'"
		Set RS = CON.Execute(CONEXAO)
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
			response.Write(CONEXAOA)

			Set RSA = CONA.Execute(CONEXAOA)

while not RSA.EOF

			mat = RSA("CO_Matricula")
			turma = RSA("CO_Turma")
			
			Set RS2 = Server.CreateObject("ADODB.Recordset")
			CONEXAO2 = "Select * from "&nota_i&" WHERE CO_Matricula = "& mat 
			Set RS2 = CON2.Execute(CONEXAO2)

while not RS2.eof
			materia=RS2("CO_Materia")

if periodo=1 then
va_apr1=RS2("Apr1_P1")
va_apr2=RS2("Apr2_P1")
va_apr3=RS2("Apr3_P1")
va_apr4=RS2("Apr4_P1")
va_apr5=RS2("Apr5_P1")
va_apr6=RS2("Apr6_P1")
va_apr7=RS2("Apr7_P1")
va_apr8=RS2("Apr8_P1")
va_v_apr1=RS2("V_Apr1_P1")
va_v_apr2=RS2("V_Apr2_P1")
va_v_apr3=RS2("V_Apr3_P1")
va_v_apr4=RS2("V_Apr4_P1")
va_v_apr5=RS2("V_Apr5_P1")
va_v_apr6=RS2("V_Apr6_P1")
va_v_apr7=RS2("V_Apr7_P1")
va_v_apr8=RS2("V_Apr8_P1")
va_sapr=RS2("VA_Sapr1")
va_pr=RS2("VA_Pr1")
va_te=RS2("VA_Te1")
va_bon=RS2("VA_Bon1")
va_me=RS2("VA_Me1")
va_mc=RS2("VA_Mc1")
va_faltas=RS2("NU_Faltas_P1")
elseif periodo=2 then
va_apr1=RS2("Apr1_P2")
va_apr2=RS2("Apr2_P2")
va_apr3=RS2("Apr3_P2")
va_apr4=RS2("Apr4_P2")
va_apr5=RS2("Apr5_P2")
va_apr6=RS2("Apr6_P2")
va_apr7=RS2("Apr7_P2")
va_apr8=RS2("Apr8_P2")
va_v_apr1=RS2("V_Apr1_P2")
va_v_apr2=RS2("V_Apr2_P2")
va_v_apr3=RS2("V_Apr3_P2")
va_v_apr4=RS2("V_Apr4_P2")
va_v_apr5=RS2("V_Apr5_P2")
va_v_apr6=RS2("V_Apr6_P2")
va_v_apr7=RS2("V_Apr7_P2")
va_v_apr8=RS2("V_Apr8_P2")
va_sapr=RS2("VA_Sapr2")
va_pr=RS2("VA_Pr2")
va_te=RS2("VA_Te2")
va_bon=RS2("VA_Bon2")
va_me=RS2("VA_Me2")
va_mc=RS2("VA_Mc2")
va_faltas=RS2("NU_Faltas_P2")
elseif periodo=3 then
va_apr1=RS2("Apr1_P3")
va_apr2=RS2("Apr2_P3")
va_apr3=RS2("Apr3_P3")
va_apr4=RS2("Apr4_P3")
va_apr5=RS2("Apr5_P3")
va_apr6=RS2("Apr6_P3")
va_apr7=RS2("Apr7_P3")
va_apr8=RS2("Apr8_P3")
va_v_apr1=RS2("V_Apr1_P3")
va_v_apr2=RS2("V_Apr2_P3")
va_v_apr3=RS2("V_Apr3_P3")
va_v_apr4=RS2("V_Apr4_P3")
va_v_apr5=RS2("V_Apr5_P3")
va_v_apr6=RS2("V_Apr6_P3")
va_v_apr7=RS2("V_Apr7_P3")
va_v_apr8=RS2("V_Apr8_P3")
va_sapr=RS2("VA_Sapr3")
va_pr=RS2("VA_Pr3")
va_te=RS2("VA_Te3")
va_bon=RS2("VA_Bon3")
va_me=RS2("VA_Me3")
va_mc=RS2("VA_Mc3")
va_faltas=RS2("NU_Faltas_P3")
elseif periodo=4 then
va_apr1=RS2("Apr1_EC")
va_apr2=RS2("Apr2_EC")
va_apr3=RS2("Apr3_EC")
va_apr4=RS2("Apr4_EC")
va_apr5=RS2("Apr5_EC")
va_apr6=RS2("Apr6_EC")
va_apr7=RS2("Apr7_EC")
va_apr8=RS2("Apr8_EC")
va_v_apr1=RS2("V_Apr1_EC")
va_v_apr2=RS2("V_Apr2_EC")
va_v_apr3=RS2("V_Apr3_EC")
va_v_apr4=RS2("V_Apr4_EC")
va_v_apr5=RS2("V_Apr5_EC")
va_v_apr6=RS2("V_Apr6_EC")
va_v_apr7=RS2("V_Apr7_EC")
va_v_apr8=RS2("V_Apr8_EC")
va_sapr=RS2("VA_Sapr_EC")
va_pr=RS2("VA_Pr4")
va_bon=RS2("VA_Bon1")
va_me=RS2("VA_Me_EC")
va_mc=RS2("VA_Mfinal")
end if
'va_faltas&";"&va_apr1&";"&va_apr2&";"&va_apr3&";"&va_apr4&";"&va_apr5&";"&va_apr6&";"&va_apr7&";"&va_apr8&";"&va_v_apr1&";"&va_v_apr2&";"&va_v_apr3&";"&va_v_apr4&";"&va_v_apr5&";"&va_v_apr6&";"&va_v_apr7&";"&va_v_apr8&";"&va_pr&";"&va_te&";"&va_me&";"&va_mc&";"&			


d.writeLine mat&";"&materia&";"&periodo&";"&va_faltas&";"&va_apr1&";"&va_apr2&";"&va_apr3&";"&va_apr4&";"&va_apr5&";"&va_apr6&";"&va_apr7&";"&va_apr8&";"&va_v_apr1&";"&va_v_apr2&";"&va_v_apr3&";"&va_v_apr4&";"&va_v_apr5&";"&va_v_apr6&";"&va_v_apr7&";"&va_v_apr8&";"&va_pr&";"&va_te&";"&va_me&";"&va_bon&";"&va_mc	

	RS2.MOVENEXT						
wend
						RSA.MOVENEXT						
					wend

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