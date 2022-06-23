<%'On Error Resume Next%>
<!--#include file="caminhos.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="funcoes_migra_notas.asp"-->
<!--#include file="atualiza_bds.asp"-->
<%
call cabecalho(1)
chave = session("chave")
session("chave")=chave

split_chave=split(chave,"-")
sistema_origem=split_chave(0)
if sistema_origem="WN" then
	endereco_origem="../wn/lancar/notas/lancar/"
elseif sistema_origem="WA" then	
	endereco_origem="../wa/professor/cna/notas/"
end if	

opt=request.QueryString("opt")
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 
data = dia &"/"& mes &"/"& ano
horario = hora & ":"& min

co_materia = request.form("co_materia")
unidade= request.form("unidade")
curso= request.form("curso")
etapa= request.form("etapa")
turma= request.form("turma")
periodo = request.form("periodo")
ano_letivo = request.form("ano_letivo")
co_prof = request.form("co_prof")
co_usr = request.form("co_usr")
max = request.form("max")

obr=co_materia&"$!$"&unidade&"$!$"&curso&"$!$"&etapa&"$!$"&turma&"$!$"&periodo&"$!$"&ano_letivo&"$!$"&co_prof

i=1

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_ns & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CONMT = Server.CreateObject("ADODB.Connection") 
		ABRIRMT = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONMT.Open ABRIRMT	
		
				
		Set RSMT  = Server.CreateObject("ADODB.Recordset")
		SQL_MT  = "Select CO_Materia_Principal from TB_Materia WHERE CO_Materia = '"& co_materia&"'"
		Set RSMT  = CONMT.Execute(SQL_MT)
		
co_materia_pr = RSMT("CO_Materia_Principal")
		
if Isnull(co_materia_pr) then
co_materia_pr= co_materia
else
co_materia_pr = co_materia_pr
end if
			va_pt=request.form("pt")
			va_pp=request.form("pp")


	Set RS5a = Server.CreateObject("ADODB.Recordset")
	SQL5a = "SELECT * FROM TB_Programa_Subs where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia_Filha ='"& co_materia &"'"
	RS5a.Open SQL5a, CONMT
	
	co_materia_pr = RS5a("CO_Materia_Principal")		
	in_faltas = RS5a("IN_Faltas")
	in_bonus = RS5a("IN_Bonus")
	rec_semestral = RS5a("IN_Rec_Semestral")
	

fail = 0
for i=1 to max
	grava="ok"
	
	
	nu_matricula = request.form("nu_matricula_"&i)

	
	if nu_matricula = "falta" then
			i=i*1
			max=max*1
		if i=max then
		grava = "no"
		else
		   'Loop para descobrir o primeiro aluno após os faltantes
			while nu_matricula = "falta"
				i=i+1
				nu_matricula = request.form("nu_matricula_"&i)						
			wend	
				
			va_faltas=request.form("faltas_"&i)
			va_t1=request.form("t1_"&i)
			va_t2=request.form("t2_"&i)
			va_t3=request.form("t3_"&i)
			va_p1=request.form("p1_"&i)
			va_p2=request.form("p2_"&i)
			va_p3=request.form("p3_"&i)
			va_bon=request.form("bon_"&i)
			va_rec=request.form("rec_"&i)
		end if	
	else
			nu_matricula = request.form("nu_matricula_"&i)			
			va_faltas=request.form("faltas_"&i)
			va_t1=request.form("t1_"&i)
			va_t2=request.form("t2_"&i)
			va_t3=request.form("t3_"&i)
			va_p1=request.form("p1_"&i)
			va_p2=request.form("p2_"&i)
			va_p3=request.form("p3_"&i)
			va_bon=request.form("bon_"&i)
			va_rec=request.form("rec_"&i)
	end if
	
	if i=1 then
		vetor_matricula = nu_matricula	
	else
		vetor_matricula = vetor_matricula&","&nu_matricula				
	end if	
	
if fail = 0 then 		
Session("va_faltas")=va_faltas
Session("va_pt")=va_pt
Session("va_pp")=va_pp
Session("va_t1")=va_t1
Session("va_t2")=va_t2
Session("va_t3")=va_t3
Session("va_p1")=va_p1
Session("va_p2")=va_p2
Session("va_p3")=va_p3
Session("va_bon")=va_bon
Session("va_rec")=va_rec	
end if

	gravou = Grava_BDS(nu_matricula, co_materia_pr, co_materia, periodo, in_faltas, in_bonus, rec_semestral, va_faltas, va_pt, va_pp, va_t1, va_t2, va_t3, va_p1, va_p2, va_p3, va_bon, va_rec, data, horario, co_usr, "S", "S")

	if gravou <>"S"then
		fail = 1
		vetor_atualizou_mae = split(gravou, "$!$")
		materia_mae = vetor_atualizou_mae(0)		
		url = vetor_atualizou_mae(1)			
		Exit For
	END IF	

next




if fail = 1 then
	response.Redirect(endereco_origem&"notas.asp?or=01&opt=err6&hp=err_"&url&"&obr="&obr) 
else

	atualizou_mae = atualiza_disciplina_mae(vetor_matricula, curso, etapa, co_materia_pr, periodo, data, horario, co_usr)
	
	if atualizou_mae <> "S" then
		vetor_atualizou_mae = split(atualizou_mae, "$!$")
		materia_mae = vetor_atualizou_mae(0)		
		url = vetor_atualizou_mae(1)
		obr = obr&"&complemento="&materia_mae	
		response.Redirect(endereco_origem&"notas.asp?or=01&opt=err6&hp=err_"&url&"&obr="&obr) 				
	end if
END IF

outro="P:"&periodo&",D:"&co_materia&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma&""

			call GravaLog (chave,outro)

if opt="cln" then
outro="P:"&periodo&",D:"&co_materia&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma&", Comunicou"
			call GravaLog (chave,outro)
response.Redirect(endereco_origem&"comunicar.asp?or=01&nota=TB_Nota_S&opt=ok&obr="&obr)
comunicou = comunica_disc_mae(unidade, curso, etapa, co_prof, co_materia_pr, periodo, "TB_Nota_A")
else
response.Redirect(endereco_origem&"notas.asp?or=01&opt=ok&obr="&obr)
end if

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
response.redirect("erro.asp")
end if
%>