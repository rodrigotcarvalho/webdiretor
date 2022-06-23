<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/parametros.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->
<%
opt=request.QueryString("opt")

grava_nota = session("grava_nota")
session("grava_nota")=grava_nota
voltaDireto = session("voltaDireto")
session("voltaDireto") = voltaDireto

		Set CON0= Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 
data = dia &"/"& mes &"/"& ano
horario = hora & ":"& min		

if opt="i" then
	wrkDataLancamento = request.form("dataLancamento")

	
	if wrkDataLancamento="" or isnull(wrkDataLancamento) then
		response.Redirect("alterar.asp?opt=err1")
	end if
		
	tx_aula = request.form("tx_aula")
	tx_observacao = request.form("tx_observacao") 
	obr = request.form("obr")

	vetor_alunos = split(lista_alunos,", ")
	
	vetor_obr = split(obr,"$!$")
	
	co_materia = vetor_obr(0) 
	co_materia_pr = busca_materia_mae(co_materia)	
	
	banco_pauta = escolheBancoPauta(vetor_obr(8),"M",p_outro)
	caminho_pauta = verificaCaminhoBancoPauta(banco_pauta,"M",p_outro)

	Set CONP = Server.CreateObject("ADODB.Connection") 
	ABRIRP = "DBQ="& caminho_pauta & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONP.Open ABRIRP	
		vetorDataExibe = split(wrkDataLancamento,"/")
		dataExclui = vetorDataExibe(1)&"/"&vetorDataExibe(0)&"/"&vetorDataExibe(2)	
	conteudo_excluido = ExcluiConteudo(CONP, dataExclui, vetor_obr(7) , co_materia_pr, co_materia, vetor_obr(1), vetor_obr(2), vetor_obr(3), vetor_obr(4), vetor_obr(5))	
	
	
	Set RS = server.createobject("adodb.recordset")		
	RS.open "TB_Materia_Lecionada", CONP, 2, 2 'which table do you want open
	RS.addnew
	
		RS("CO_Professor") = vetor_obr(7) 
		RS("CO_Materia_Principal") = co_materia_pr
		RS("CO_Materia") = co_materia
		RS("NU_Unidade")=vetor_obr(1) 				
		RS("CO_Curso")=vetor_obr(2) 
		RS("CO_Etapa")=vetor_obr(3) 
		RS("CO_Turma")=vetor_obr(4) 
		RS("NU_Periodo")=vetor_obr(5) 
		RS("DT_Aula")=wrkDataLancamento			
		RS("TX_Aula")=tx_aula
		RS("TX_Obs") = tx_observacao
		RS("DA_Ult_Acesso")=data 
		RS("HO_ult_Acesso")=horario
		RS("CO_Usuario")=session("co_user") 			
		
		
		
	RS.update
	set RS=nothing	
	

	wrkDataLancamento=replace(wrkDataLancamento,"/",".")
	response.Redirect("alterar.asp?opt=ok1&P_DATA_AULA="&wrkDataLancamento)	
elseif opt="e" then	
	unidade= session("unidades")
	curso= session("grau")
	etapa= session("serie")
	turma= session("turma")
	co_materia = session("co_materia")
	co_mat_prin = session("co_mat_prin")
	periodo = session("periodo")
	co_prof = session("co_prof")
	co_usr = session("co_usr")
	tb = session("nota")
	session("co_materia")=co_materia
	session("co_mat_prin")=co_mat_prin			
	session("unidades")=unidade
	session("grau")=curso
	session("serie")=etapa
	session("turma")=turma
	session("periodo")=periodo
	session("co_prof") = co_prof 
	session("nota") = tb
	
	
	banco_pauta = escolheBancoPauta(tb,"M",p_outro)
	caminho_pauta = verificaCaminhoBancoPauta(banco_pauta,"M",p_outro)

	Set CONP = Server.CreateObject("ADODB.Connection") 
	ABRIRP = "DBQ="& caminho_pauta & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONP.Open ABRIRP	
	
	vetorDatas=session("obr")
	session("obr")=vetorDatas	

	vetorExclui = split(vetorDatas,", ")
	conta_ocorr=0
	for e =0 to ubound(vetorExclui)
		exclui = replace(vetorExclui(e),".","/")
		vetorDataExibe = split(exclui,"/")
		dataExclui = vetorDataExibe(1)&"/"&vetorDataExibe(0)&"/"&vetorDataExibe(2)		
		conteudo_excluido = ExcluiConteudo(CONP, dataExclui, co_prof , co_mat_prin, co_materia, unidade, curso, etapa, turma, periodo)	
	next
	response.Redirect("notas.asp?opt=ok")	
end if

function ExcluiConteudo(CONPauta, P_DT_AULA, co_prof, co_mat_prin, co_materia, unidade, curso, etapa, turma, periodo)


	Set RSP = Server.CreateObject("ADODB.Recordset")
	SQLP = "Select * from TB_Materia_Lecionada WHERE DT_Aula = #"&P_DT_AULA&"# AND CO_Professor  = "& co_prof &" AND CO_Materia_Principal = '"& co_mat_prin &"' AND CO_Materia = '"& co_materia &"' AND NU_Unidade  = "& unidade &" AND CO_Curso  = '"& curso &"' AND CO_Etapa  = '"& etapa &"' AND CO_Turma  = '"& turma &"' AND NU_Periodo = "& periodo
	Set RSP = CONPauta.Execute(SQLP)	
	
	if RSP.EOF  then
		ExcluiPauta = "N"
	else
		Set RSD = Server.CreateObject("ADODB.Recordset")
		SQLD = "DELETE * from TB_Materia_Lecionada WHERE DT_Aula = #"&P_DT_AULA&"# AND CO_Professor  = "& co_prof &" AND CO_Materia_Principal = '"& co_mat_prin &"' AND CO_Materia = '"& co_materia &"' AND NU_Unidade  = "& unidade &" AND CO_Curso  = '"& curso &"' AND CO_Etapa  = '"& etapa &"' AND CO_Turma  = '"& turma &"' AND NU_Periodo = "& periodo
		Set RSD = CONPauta.Execute(SQLD)						
		ExcluiConteudo = "S"	
	end if		
end function

%>