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

if opt="i" then
	wrkDataLancamento = request.form("dataLancamentoForm")

	
	if wrkDataLancamento="" or isnull(wrkDataLancamento) then
		response.Redirect("alterar.asp?opt=err1")
	end if
		
	lista_alunos = request.form("nu_linha#!#matricula")
	qtd_aulas = request.form("qtdAulasForm")
	qtdPrevistas = request.form("previstasForm")	
	obr = request.form("obr")


	vetor_alunos = split(lista_alunos,", ")
	
	vetor_obr = split(obr,"$!$")
	
	co_materia = vetor_obr(0) 
	co_materia_pr = busca_materia_mae(co_materia)	
	
	banco_pauta = escolheBancoPauta(vetor_obr(8),p_subopcao,p_outro)
	caminho_pauta = verificaCaminhoBancoPauta(banco_pauta,p_subopcao,p_outro)

	Set CONP = Server.CreateObject("ADODB.Connection") 
	ABRIRP = "DBQ="& caminho_pauta & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONP.Open ABRIRP	

	pauta_excluida = ExcluiPauta(CONP, wrkDataLancamento, vetor_obr(7) , co_materia_pr, co_materia, vetor_obr(1), vetor_obr(2), vetor_obr(3), vetor_obr(4), vetor_obr(5))	
	

	Set RSTP = Server.CreateObject("ADODB.Recordset")	
	SQLTP = "Select NU_Pauta from TB_Pauta WHERE CO_Professor  = "& vetor_obr(7) &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Unidade  = "& vetor_obr(1) &" AND CO_Curso  = '"& vetor_obr(2) &"' AND CO_Etapa  = '"& vetor_obr(3) &"' AND CO_Turma  = '"& vetor_obr(4) &"' AND NU_Periodo = "& vetor_obr(5)	
	Set RSTP = CONP.Execute(SQLTP)	
	
	if NOT RSTP.EOF then	
		nu_pauta = RSTP("NU_Pauta") 
		
		if isnull(qtdPrevistas) or qtdPrevistas="" then
			strSQL3= "UPDATE TB_Pauta SET NU_Dia_Previsto= NULL WHERE CO_Professor  = "& vetor_obr(7) &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Unidade  = "& vetor_obr(1) &" AND CO_Curso  = '"& vetor_obr(2) &"' AND CO_Etapa  = '"& vetor_obr(3) &"' AND CO_Turma  = '"& vetor_obr(4) &"' AND NU_Periodo = "& vetor_obr(5)		
		else
			strSQL3= "UPDATE TB_Pauta SET NU_Dia_Previsto= "& qtdPrevistas & " WHERE CO_Professor  = "& vetor_obr(7) &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Unidade  = "& vetor_obr(1) &" AND CO_Curso  = '"& vetor_obr(2) &"' AND CO_Etapa  = '"& vetor_obr(3) &"' AND CO_Turma  = '"& vetor_obr(4) &"' AND NU_Periodo = "& vetor_obr(5)		
		end if				
		
	
		set tabela3 = CONP.Execute (strSQL3)		
		
	else
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		CONEXAO0 = "Select Max(NU_Pauta) as Max_NU_Pauta from TB_Pauta"
		Set RS0 = CONP.Execute(CONEXAO0)	
		
		if RS0.EOF then
			nu_pauta=0
		else
			nu_pauta =RS0("Max_NU_Pauta") 
			if isnull(nu_pauta) or nu_pauta="" then
				nu_pauta=0
			end if
		end if 	
		nu_pauta = nu_pauta+1
	
		Set RS = server.createobject("adodb.recordset")		
		RS.open "TB_Pauta", CONP, 2, 2 'which table do you want open
		RS.addnew
		
			RS("NU_Pauta") = nu_pauta
			RS("CO_Professor") = vetor_obr(7) 
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Unidade")=vetor_obr(1) 				
			RS("CO_Curso")=vetor_obr(2) 
			RS("CO_Etapa")=vetor_obr(3) 
			RS("CO_Turma")=vetor_obr(4) 
			RS("NU_Periodo")=vetor_obr(5) 
			RS("NU_Dia_Previsto")=qtdPrevistas
			
		
		RS.update
		set RS=nothing	
	end if	
	
	Set RS1 = Server.CreateObject("ADODB.Recordset")
	CONEXAO1 = "Select Max(NU_Seq) as Max_NU_Seq from TB_Pauta_Aula Where NU_Pauta = "&nu_pauta
	Set RS1 = CONP.Execute(CONEXAO1)
				
	if RS1.EOF then
		nu_seq=0
	else
		nu_seq =RS1("Max_NU_Seq") 
		if isnull(nu_seq) or nu_seq="" then
			nu_seq=0
		end if		
	end if 		
	
	for a=1 to qtd_aulas
	
		Set RS = server.createobject("adodb.recordset")		
		RS.open "TB_Pauta_Aula", CONP, 2, 2 'which table do you want open
		RS.addnew
		
			RS("NU_Pauta") = nu_pauta
			RS("NU_Seq") = nu_seq+a
			RS("DT_Aula") = wrkDataLancamento
			RS("NU_Tempo") = a
		RS.update
		set RS=nothing		
	Next	

	
	for ck=0 to ubound(vetor_alunos)
		vetor_dados = split(vetor_alunos(ck),"#!#")	
		linha = vetor_dados(0)
		matricula = vetor_dados(1)
		response.Write(matricula&"<BR>")		
		for q=1 to qtd_aulas
			avaliacao = request.form("check_"&q&"_"&linha)
			response.Write(a&">>>"&avaliacao&"<BR>")
			
			if avaliacao="S" then						
					
				
				Set RS = server.createobject("adodb.recordset")		
				RS.open "TB_Pauta_Faltas", CONP, 2, 2 'which table do you want open
				RS.addnew
				
					RS("NU_Pauta") = nu_pauta
					RS("NU_Seq") = nu_seq+q 
					RS("CO_Matricula") = matricula
				RS.update
				set RS=nothing				
			end if					
		next	

	next
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
	
	
	banco_pauta = escolheBancoPauta(tb,p_subopcao,p_outro)
	caminho_pauta = verificaCaminhoBancoPauta(banco_pauta,p_subopcao,p_outro)

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
		pauta_excluida = ExcluiPauta(CONP, dataExclui, co_prof , co_mat_prin, co_materia, unidade, curso, etapa, turma, periodo)	
	next
	response.Redirect("notas.asp?opt=ok")	
end if

function ExcluiPauta(CONPauta, P_DT_AULA, co_prof, co_mat_prin, co_materia, unidade, curso, etapa, turma, periodo)
	
	Set RSP = Server.CreateObject("ADODB.Recordset")	
	
    if P_DT_AULA="" or isnull(P_DT_AULA) then
		SQL = "Select TB_Pauta_Aula.NU_Pauta, TB_Pauta_Aula.NU_Seq, TB_Pauta_Aula.DT_Aula, TB_Pauta_Aula.NU_Tempo from TB_Pauta INNER JOIN TB_Pauta_Aula on TB_Pauta.NU_Pauta=TB_Pauta_Aula.NU_Pauta WHERE CO_Professor  = "& co_prof &" AND CO_Materia_Principal = '"& co_mat_prin &"' AND CO_Materia = '"& co_materia &"' AND NU_Unidade  = "& unidade &" AND CO_Curso  = '"& curso &"' AND CO_Etapa  = '"& etapa &"' AND CO_Turma  = '"& turma &"' AND NU_Periodo = "& periodo	

	else
		V_DT_AULA=split(P_DT_AULA,"/")
		WRK_DT_AULA = V_DT_AULA(1)&"/"&V_DT_AULA(0)&"/"&V_DT_AULA(2)		
		SQL = "Select TB_Pauta_Aula.NU_Pauta, TB_Pauta_Aula.NU_Seq, TB_Pauta_Aula.DT_Aula, TB_Pauta_Aula.NU_Tempo from TB_Pauta INNER JOIN TB_Pauta_Aula on TB_Pauta.NU_Pauta=TB_Pauta_Aula.NU_Pauta WHERE DT_Aula = #"&WRK_DT_AULA&"# AND CO_Professor  = "& co_prof &" AND CO_Materia_Principal = '"& co_mat_prin &"' AND CO_Materia = '"& co_materia &"' AND NU_Unidade  = "& unidade &" AND CO_Curso  = '"& curso &"' AND CO_Etapa  = '"& etapa &"' AND CO_Turma  = '"& turma &"' AND NU_Periodo = "& periodo		
	
	end if	
	
	Set RSP = CONPauta.Execute(SQL)	
	
	if RSP.EOF  then
		ExcluiPauta = "N"
	else
		while not RSP.EOF 
			NU_Pauta = RSP("NU_Pauta")
			NU_Seq = RSP("NU_Seq") 
			data_Pauta = RSP("DT_Aula")
			wrkQtdAulasLancadas = RSP("NU_Tempo")	
				
			TBPautaFaltasExcluida = ExcluiTBPautaFaltas(CONPauta, NU_Pauta, NU_Seq)
			
			TBPautaAulaExcluida = ExcluiTBPautaAula(CONPauta, NU_Pauta, NU_Seq)
				
			TBPautaExcluida = ExcluiTBPauta(CONPauta, NU_Pauta)
			
		RSP.movenext
		wend						
		ExcluiPauta = "S"	
	end if		
end function

function ExcluiTBPautaFaltas(CONPauta, NU_Pauta, NU_Seq)

	Set RSC = Server.CreateObject("ADODB.Recordset")
	SQLC = "Select * FROM TB_Pauta_Faltas WHERE NU_Pauta="&NU_Pauta&" AND NU_Seq = "& NU_Seq 
	Set RSC = CONPauta.Execute(SQLC)
	
	ExcluiTBPautaFaltas	  = "N"		
		
	if NOT RSC.EOF then
		Set RSD = Server.CreateObject("ADODB.Recordset")
		SQLD = "DELETE * FROM TB_Pauta_Faltas WHERE NU_Pauta="&NU_Pauta&" AND NU_Seq = "& NU_Seq 
		Set RSD = CONPauta.Execute(SQLD)
		
		ExcluiTBPautaFaltas	  = "S"		
	end if
			
		
				
end function
function ExcluiTBPautaAula(CONPauta, NU_Pauta, NU_Seq)

	Set RSC = Server.CreateObject("ADODB.Recordset")
	SQLC = "Select * FROM TB_Pauta_Aula WHERE NU_Pauta="&NU_Pauta&" AND NU_Seq = "& NU_Seq 
	Set RSC = CONPauta.Execute(SQLC)
	
	ExcluiTBPautaAula = "N"		
		
	if NOT RSC.EOF then

		Set RSD = Server.CreateObject("ADODB.Recordset")
		SQLD = "DELETE * FROM TB_Pauta_Aula WHERE NU_Pauta="&NU_Pauta&" AND NU_Seq = "& NU_Seq 
		Set RSD = CONPauta.Execute(SQLD)
			
		ExcluiTBPautaAula = "S"	
	end if				
end function

function ExcluiTBPauta(CONPauta, NU_Pauta)

	Set RSC = Server.CreateObject("ADODB.Recordset")
	SQLC = "Select * FROM TB_Pauta WHERE NU_Pauta="&NU_Pauta
	Set RSC = CONPauta.Execute(SQLC)
	
	ExcluiTBPauta = "N"		
		
	if NOT RSC.EOF then
	
		Set RSPA = Server.CreateObject("ADODB.Recordset")
		SQLPA = "Select 'S' FROM TB_Pauta_Aula WHERE NU_Pauta="&NU_Pauta
		Set RSPA = CONPauta.Execute(SQLPA)	
		
		if RSPA.EOF then	
		
			Set RSPF = Server.CreateObject("ADODB.Recordset")
			SQLPF = "Select 'S' FROM TB_Pauta_Faltas WHERE NU_Pauta="&NU_Pauta
			Set RSPF = CONPauta.Execute(SQLPF)	
			
			if RSPF.EOF then				
		
				Set RSD = Server.CreateObject("ADODB.Recordset")
				SQLD = "DELETE * FROM TB_Pauta WHERE NU_Pauta="&NU_Pauta
				Set RSD = CONPauta.Execute(SQLD)
								
				ExcluiTBPauta = "S"		
			end if
		end if			
	end if					
end function
%>