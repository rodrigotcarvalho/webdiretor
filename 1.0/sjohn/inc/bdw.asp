<%'On Error Resume Next%>
<!--#include file="caminhos.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="funcoes7.asp"-->
<%
call cabecalho(1)
chave = session("chave")
session("chave")=chave

split_chave=split(chave,"-")
sistema_origem=split_chave(0)

if sistema_origem="WN" then
	endereco_origem="../wn/lancar/notas/laq/"
elseif sistema_origem="WA" then	
	endereco_origem="../wa/professor/cna/maq/"
end if	

opt=request.QueryString("opt")
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 
data = dia &"/"& mes &"/"& ano
horario = hora & ":"& min
co_materia = request.form("co_mat")
co_materia_pr = request.form("co_mat_prin")
unidade= request.form("unidade")
curso= request.form("curso")
etapa= request.form("etapa")
turma= request.form("turma")
ano_letivo = request.form("ano_letivo")
co_prof = request.form("co_prof")
co_usr = request.form("co_usr")
max = request.form("max")

obr=co_materia&"?"&unidade&"?"&curso&"?"&etapa&"?"&turma&"?"&periodo&"?"&ano_letivo&"?"&co_prof

i=1

		Set CON_TBW = Server.CreateObject("ADODB.Connection") 
		ABRIR_TBW = "DBQ="& CAMINHO_nw & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_TBW.Open ABRIR_TBW
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
'		Set RS8 = Server.CreateObject("ADODB.Recordset")
'		SQL8 = "SELECT * FROM TB_Materia where CO_Materia='"& co_materia &"'"
'		RS8.Open SQL8, CON0
'
'		no_mat= RS8("NO_Materia")	
'		co_materia_pr= RS8("CO_Materia_Principal")
'		
'if Isnull(co_materia_pr) then
'	co_materia_pr= co_materia
'end if

dados_periodo =  periodos(periodo, "num")
total_periodo = split(dados_periodo,"#!#") 
notas_a_lancar = ubound(total_periodo)-2

vetor_matrics = alunos_turma(session("ano_letivo"),unidade,curso,etapa,turma,"num")

vetor_alunos = split(vetor_matrics,"#$#") 

check = 2
nu_chamada_ckq = 0
linha_tabela=1


for a = 0 to ubound(vetor_alunos) 
dados_alunos = split(vetor_alunos(a),"#!#") 

	fail = 0
	for b=0 to notas_a_lancar		
		identificacao = "av_n"&dados_alunos(1)&"_p"&total_periodo(b)
		response.Write(identificacao&"<BR>")		
		if b=0 then
			wrk_val_nota_per_1 = request.form(identificacao)
		elseif b=1 then
			wrk_val_nota_per_2 = request.form(identificacao)
		elseif b=2 then 			
			wrk_val_nota_per_3 = request.form(identificacao)
		elseif b=3 then 			 	
			wrk_val_nota_per_4 = request.form(identificacao)
		end if	
	Next

	IF dados_alunos(3) = "C" THEN
		Set RSC = Server.CreateObject("ADODB.Recordset")
		SQLC = "Select * from TB_Nota_W WHERE CO_Matricula = "& dados_alunos(0) &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"'"
		Set RSC = CON_TBW.Execute(SQLC)
	
'		response.Write(CONEXAO0&"<BR>")
			
		If RSC.EOF THEN	
			Set RS = server.createobject("adodb.recordset")		
			RS.open "TB_Nota_W", CON_TBW, 2, 2 'which table do you want open
			RS.addnew
			
				RS("CO_Matricula") = dados_alunos(0)
				RS("CO_Materia_Principal") = co_materia_pr
				RS("CO_Materia") = co_materia
				RS("VA_Ava1") = wrk_val_nota_per_1	
				RS("VA_Ava2")=wrk_val_nota_per_2
				RS("VA_Ava3")=wrk_val_nota_per_3
				RS("VA_Ava4")=wrk_val_nota_per_4
				RS("DA_Ult_Acesso") = data
				RS("HO_ult_Acesso") = horario
				RS("CO_Usuario")= co_usr
			
			RS.update
			set RS=nothing
			
		else
			Set RSD = Server.CreateObject("ADODB.Recordset")
			SQLD = "DELETE * from TB_Nota_W WHERE CO_Matricula = "& dados_alunos(0) &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"'"
			Set RSD = CON_TBW.Execute(SQLD)
					'response.Write(CONEXAO0&"<BR>")
	
			Set RS = server.createobject("adodb.recordset")		
			RS.open "TB_Nota_W", CON_TBW, 2, 2 'which table do you want open
			RS.addnew
			
				RS("CO_Matricula") = dados_alunos(0)
				RS("CO_Materia_Principal") = co_materia_pr
				RS("CO_Materia") = co_materia
				RS("VA_Ava1") = wrk_val_nota_per_1	
				RS("VA_Ava2")=wrk_val_nota_per_2
				RS("VA_Ava3")=wrk_val_nota_per_3
				RS("VA_Ava4")=wrk_val_nota_per_4
				RS("DA_Ult_Acesso") = data
				RS("HO_ult_Acesso") = horario
				RS("CO_Usuario")= co_usr
			
			RS.update
			set RS=nothing	
			
		end if
	END IF
NEXT


outro="D:"&co_materia&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma&""
call GravaLog (chave,outro)

if opt="cln" then
	outro="P:"&periodo&",D:"&co_materia&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma&", Comunicou"
	call GravaLog (chave,outro)
	response.Redirect(endereco_origem&"comunicar.asp?or=01&nota=TB_Nota_F&opt=ok&obr="&obr)
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