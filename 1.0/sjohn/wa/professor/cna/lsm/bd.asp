<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/parametros.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->
<!--#include file="../../../../../global/tabelas_escolas.asp"-->
<!--#include file="../../../../inc/atualiza_planilha.asp"-->
<%
Server.ScriptTimeout = 600 'valor em segundos
chave = session("chave")
session("chave")=chave

opt=request.QueryString("opt")
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 
data = dia &"/"& mes &"/"& ano
horario = hora & ":"& min

opcao = request.QueryString("opt")
unidade= request.form("unidade")
curso= request.form("curso")
etapa= request.form("etapa")
turma= request.form("turma")
periodo= request.form("periodo")

areadoconhecimento=request.form("areaconhecimento")

ano_letivo = request.form("ano_letivo")
co_usr = request.form("co_usr")
max = request.form("max")
submit = request.Form("Submit")	

obr=unidade&"$!$"&curso&"$!$"&etapa&"$!$"&turma&"$!$"&periodo&"$!$"&areadoconhecimento

		Set CONg = Server.CreateObject("ADODB.Connection") 
		ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONg.Open ABRIRg	
		
		Set CON_AL = Server.CreateObject("ADODB.Connection") 
		ABRIR_AL = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_AL.Open ABRIR_AL

tb_nota = tabela_notas(CONg, unidade, curso, etapa, turma, periodo, co_materia, outro)

CAMINHOn = caminho_notas(CONg, tb_nota, outro)	

val_default_lbs = parametros_gerais(unidade,curso,etapa,turma,co_materia,"default_lbs",outro)

dados_tabela=verifica_dados_tabela(CAMINHOn,opcao,outro)
	dados_separados=split(dados_tabela,"#$#")
	tb=dados_separados(0)

i=1

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CONMT = Server.CreateObject("ADODB.Connection") 
		ABRIRMT = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONMT.Open ABRIRMT
		
		Set RSAC = Server.CreateObject("ADODB.Recordset")		
		CONEXAOAC = "SELECT TB_Area_ConhecimentoxTB_Programa_Aula.[CO_Materia] AS MATERIA FROM TB_Area_ConhecimentoxTB_Programa_Aula WHERE (((TB_Area_ConhecimentoxTB_Programa_Aula.[CO_Curso])='"&curso&"') AND ((TB_Area_ConhecimentoxTB_Programa_Aula.[CO_Etapa])='"&etapa&"') AND ((TB_Area_ConhecimentoxTB_Programa_Aula.[TP_Area])="&areadoconhecimento&")) ORDER BY TB_Area_ConhecimentoxTB_Programa_Aula.[CO_Materia]"
		Set RSAC = CONMT.Execute(CONEXAOAC)		
		
		conta_mat_ac=0
		while not RSAC.EOF
			co_materia = RSAC("MATERIA")
			if conta_mat_ac=0 then
				materia_ac = co_materia
			else
				materia_ac = materia_ac&"#!#"&co_materia
			end if	
			conta_mat_ac=conta_mat_ac+1
		RSAC.MOVENEXT
		WEND		
		
		vetor_materia_ac = split(materia_ac, "#!#")
fail = 0
for i=1 to max
	grava="ok"
		'response.Write(i&"<<BR>")	
	
	nu_matricula = request.form("nu_matricula_"&i)
	
	if nu_matricula = "falta" then
		while nu_matricula = "falta" 
			i=i+1
			nu_matricula = request.form("nu_matricula_"&i)
		wend		
	else
		nu_matricula = request.form("nu_matricula_"&i)
	end if
	va_sim=request.form("val_simulado_"&i)
	
	if va_sim="" or isnull(va_sim) then
		va_sim=NULL		
	else
		if isnumeric(va_sim) then					
			va_sim=va_sim*1
			val_default_lbs = val_default_lbs*1
				if int(va_sim) <> va_sim then
					fail = 1 
					erro = "int"
					url = nu_matricula&"_"&va_sim&"_"&erro
					grava = "no"				
				end if				
		else
			fail = 1 
			erro = "num"
			url = nu_matricula&"_"&va_sim&"_"&erro
			grava = "no"
		end if
	end if

	
	if grava="no" then
	else	
	
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		CONEXAO0 = "delete from "&tb&" WHERE CO_Matricula = "& nu_matricula &" AND TP_Area = "&areadoconhecimento&" AND NU_Periodo="&periodo	
		Set RS0 = CON.Execute(CONEXAO0)
	
			
		'response.Write("4"&turma &"/"&co_materia_pr)
		Set RS = server.createobject("adodb.recordset")
		
		RS.open tb, CON, 2, 2 'which table do you want open
		RS.addnew
			RS("TP_Area") = areadoconhecimento	
			RS("CO_Matricula") = nu_matricula	
			RS("NU_Periodo")=periodo			
			RS("VA_SIM")=va_sim
		
		RS.update
		set RS=nothing
			
		
	end if
	'response.Write(i&"-grava-"&grava&"hp=err_"&url&"&obr="&obr&"<br>")
if submit<>"Salvar" then
'	Set RSA = Server.CreateObject("ADODB.Recordset")
'	SQL_A = "Select * from TB_Matriculas WHERE NU_Ano="&ano_letivo&" AND NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
'	Set RSA = CON_AL.Execute(SQL_A)
'	
'	
'	While Not RSA.EOF
'		nu_matricula = RSA("CO_Matricula")
		
			Set RSF = Server.CreateObject("ADODB.Recordset")
			CONEXAO0 = "Select * from "&tb&" WHERE CO_Matricula = "& nu_matricula &" AND NU_Periodo="&periodo&" AND TP_Area	="&areadoconhecimento	
			Set RSF = CON.Execute(CONEXAO0)
			
			While Not RSF.EOF
					
			    atualizou = atualiza_planilha_simulado(opt,unidade, curso, etapa, turma, periodo, areadoconhecimento, nu_matricula, va_sim, outro)
				
			RSF.MOVENEXT
			WEND	
'	RSA.MOVENEXT
'	WEND		
end if		
	
next



'response.Write(fail&" - "&int(va_f1)&" - "&va_f1&" <BR>")
fail=fail*1
if fail = 1 then

	response.Redirect("notas.asp?opt=err_"&erro&"&obr="&obr) 

END IF



outro="U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma&""

			call GravaLog (chave,outro)

response.Redirect("notas.asp?opt=ok&obr="&obr)


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