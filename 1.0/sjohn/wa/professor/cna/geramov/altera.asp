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
'response.Write(caminho_gera_mov&",,,")
Set fs = CreateObject("Scripting.FileSystemObject") 'cria  
Set d = fs.CreateTextFile(caminho_gera_mov&arquivo, False)  
'd.write("teste")  
'd.writeblanklines(5)  
'd.writeline("deixei 5 linhas em branco")  
'd.close()  
'd.writeline("In�cio do Arquivo de Notas")  		

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
		elseif nota_i ="TB_NOTA_D" then
				CAMINHO_n = CAMINHO_nd
		elseif nota_i ="TB_NOTA_E" then
				CAMINHO_n = CAMINHO_ne					
		end if
				
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_n & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2

			Set RSA= Server.CreateObject("ADODB.Recordset")
			CONEXAOA = "Select * from TB_Aluno_Esta_Turma WHERE NU_Unidade = "& unidade &" And CO_Curso = '"& curso &"' And CO_Etapa = '"& etapa &"' order by NU_Chamada"
			Set RSA = CONA.Execute(CONEXAOA)

while not RSA.EOF

			mat = RSA("CO_Matricula")
			turma = RSA("CO_Turma")
			
			Set RS3 = Server.CreateObject("ADODB.Recordset")
			CONEXAO3 = "Select * from "&nota_i&" WHERE CO_Matricula = "& mat&" AND NU_Periodo="&periodo&" AND (DA_Ult_Acesso BETWEEN #"&data_de&"# AND #"&data_ate&"#) order by CO_Matricula" 
			Set RS3 = CON2.Execute(CONEXAO3)

while not RS3.eof
		materia=RS3("CO_Materia")
		va_m1=RS3("VA_Media1")
		va_bon=RS3("VA_Bonus")
		va_m2=RS3("VA_Media2")
		va_rec=RS3("VA_Rec")
		va_m3=RS3("VA_Media3")
d.writeLine mat&";"&materia&";"&periodo&";"&va_m1&";"&va_bon&";"&va_m2&";"&va_rec&";"&va_m3			
'response.Write(nota_i&" - "&wrt&"<BR>")

	RS3.MOVENEXT						
wend
						RSA.MOVENEXT						
					wend
'd.writeline("Fim do Arquivo")  
'd.writeblanklines(5)  
'd.writeline("In�cio do Arquivo de Faltas")
  
			Set RSA= Server.CreateObject("ADODB.Recordset")
			CONEXAOA = "Select * from TB_Aluno_Esta_Turma WHERE NU_Unidade = "& unidade &" And CO_Curso = '"& curso &"' And CO_Etapa = '"& etapa &"' order by NU_Chamada"
			Set RSA = CONA.Execute(CONEXAOA)

while not RSA.EOF
			mat = RSA("CO_Matricula")			
			Set RS3 = Server.CreateObject("ADODB.Recordset")
			CONEXAO3 = "Select * from TB_Frequencia_Periodo WHERE CO_Matricula = "& mat 
			Set RS3 = CON2.Execute(CONEXAO3)

while not RS3.eof
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
d.writeLine mat&";FREQ;"&va_f1&";"&va_f2&";"&va_f3&";"&va_f4			
'response.Write(nota_i&" - "&wrt&"<BR>")

	RS3.MOVENEXT						
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