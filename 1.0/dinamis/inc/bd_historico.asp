<!--#include file="caminhos.asp"-->
<%

	Set CON7 = Server.CreateObject("ADODB.Connection") 
	ABRIR7 = "DBQ="& CAMINHO_h & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON7.Open ABRIR7	
	
Function maiorSequencialHistorico(num_matric_aluno, num_ano_historico)	
		if isnull(num_ano_historico) then
			num_ano_historico= session("ano_letivo")
		end if
		
			Set RSM = Server.CreateObject("ADODB.Recordset")
			SQLM = "SELECT MAX(NU_Seq) as sequencial FROM TB_Historico_Ano where CO_Matricula = "& num_matric_aluno &" AND DA_Ano = "&num_ano_historico&""	
			RSM.Open SQLM, CON7	
			
			IF RSM.EOF then			
				maiorSequencialHistorico = 1
			else
				maiorSequencialHistorico = RSM("sequencial")
				if isnull(maiorSequencialHistorico) or maiorSequencialHistorico="" then
					maiorSequencialHistorico = 1				
				end if
									
			end if

end function	

Function ehLancamentoManual(num_matric_aluno, num_ano_historico, nu_seq)	
		if isnull(num_ano_historico) then
			num_ano_historico= session("ano_letivo")
		end if
		
			Set RSM = Server.CreateObject("ADODB.Recordset")
			SQLM = "SELECT TP_Registro FROM TB_Historico_Ano where CO_Matricula = "& num_matric_aluno &" AND DA_Ano = "&num_ano_historico&" AND NU_Seq = "&nu_seq	
			RSM.Open SQLM, CON7	
			
			IF RSM.EOF then			
				ehLancamentoManual = "N"
			else
				if RSM("TP_Registro") = "M" then
					ehLancamentoManual = "S"				
				else
					ehLancamentoManual = "N"				
				end if			
			end if
end function	

Function excluiHistorico(num_matric_aluno, num_ano_historico, nu_seq)	

	nota_historico_excluida = excluiTbHistoricoNota(num_matric_aluno, num_ano_historico, nu_seq)	
	
	ano_historico_excluido = excluiTbHistoricoAno(num_matric_aluno, num_ano_historico, nu_seq)	

end function	

Function excluiTbHistoricoAno(num_matric_aluno, num_ano_historico, nu_seq)	

			Set RSM = Server.CreateObject("ADODB.Recordset")
			SQLM = "DELETE * FROM TB_Historico_Ano where CO_Matricula = "& num_matric_aluno &" AND DA_Ano = "&num_ano_historico&" AND NU_Seq = "&nu_seq	
			RSM.Open SQLM, CON7	
			
			'IF RSM.EOF then
			'	excluiTbHistoricoAno = "N"
			'else
				excluiTbHistoricoAno = "S"				
			'end if

end function	

Function excluiTbHistoricoNota(num_matric_aluno, num_ano_historico, nu_seq)	

			Set RSM = Server.CreateObject("ADODB.Recordset")
			SQLM = "DELETE * FROM TB_Historico_Nota where CO_Matricula = "& num_matric_aluno &" AND DA_Ano = "&num_ano_historico&" AND NU_Seq = "&nu_seq	
			RSM.Open SQLM, CON7	
			
			'IF RSM.EOF then
			'	excluiTbHistoricoNota = "N"
			'else
				excluiTbHistoricoNota = "S"				
			'end if			
end function	

Function converteTBTipoCurso(ano_letivo, curso)
	curso=curso*1
	if curso=1 then
      converteTBTipoCurso = "EFA"	  
	elseif curso=2 then
      converteTBTipoCurso = "EM"
	end if
end function	

Function insereTbAnoHistorico(p_num_matric_aluno, p_num_ano_historico, p_num_seq_historico, p_tipo_curso, p_co_seg, p_no_escola, p_no_pais, p_no_municipio, p_sg_uf, p_in_aprovado, p_tx_observacoes, p_da_ano_historico, p_tp_registro, p_carga_horaria, p_tx_frequencia)	


        tipo_curso = converteTBTipoCurso(p_num_ano_historico, p_tipo_curso)
			
		Set RSM = server.createobject("adodb.recordset")		
		RSM.open "TB_Historico_Ano", CON7, 2, 2 'which table do you want open
		RSM.addnew
		
			RSM("CO_Matricula") = p_num_matric_aluno
			RSM("DA_Ano") = p_num_ano_historico
			RSM("NU_Seq") = p_num_seq_historico
			RSM("TP_Curso") = tipo_curso
			RSM("CO_Seg") = p_co_seg
			RSM("NO_Escola") = p_no_escola
			RSM("NO_Pais") = p_no_pais				
			RSM("NO_Municipio") = p_no_municipio
			RSM("SG_UF") = p_sg_uf		
			RSM("IN_Aprovado") = p_in_aprovado								
			RSM("TX_Observacoes") = p_tx_observacoes	
			RSM("NU_ANO_Letivo") = p_da_ano_historico
			RSM("TP_Registro") = p_tp_registro	
			RSM("DT_Registro") = FormatDateTime(now,2)
			RSM("NU_Carga_Horaria_Total") = p_carga_horaria	
			RSM("TX_Frequencia_Total") = p_tx_frequencia																				
			
		RSM.update
		set RSM=nothing
		
		insereTbAnoHistorico = "S"
end function	

Function insereTbHistoricoNota(num_matric_aluno, num_ano_historico, num_seq_historico, no_materia, num_carga_horaria, tx_frequencia, va_nota, in_aprovado)	

			
		Set RSM = server.createobject("adodb.recordset")		
		RSM.open "TB_Historico_Nota", CON7, 2, 2 'which table do you want open
		RSM.addnew
		
			RSM("CO_Matricula") = num_matric_aluno
			RSM("DA_Ano") = num_ano_historico
			RSM("NU_Seq") = num_seq_historico
			RSM("NO_Materia") = no_materia
			RSM("NU_Carga_Horaria") = num_carga_horaria
			RSM("TX_Frequencia") = tx_frequencia
			RSM("VA_Nota") = va_nota				
			RSM("IN_Aprovado") = in_aprovado																									
			
		RSM.update
		set RSM=nothing
		
		insereTbHistoricoNota = "S"
end function	
	
Function tbResultadoFinal(tipoResultado)	
		if isnull(tipoResultado) then
			wrkAprov= ""
		else
		
			Set RSR = Server.CreateObject("ADODB.Recordset")
			SQLR = "SELECT * FROM TB_Resultado_Final where TP_Resultado = "& tipoResultado	
			RSR.Open SQLR, CON7	
			
			wrkAprov = RSR("NO_Resultado")

		end if
	tbResultadoFinal = wrkAprov

end function
Function tbHistoricoAnoEscola (codCons, codCurso, CodEtapa)
		
	SQLF = "SELECT * FROM TB_Historico_Ano where CO_Matricula = "& codCons &" and TP_Curso ='"& codCurso&"'"
		
	if not isnull(CodEtapa) then
		SQLF = SQLF &" and CO_Seg='"&CodEtapa&"'"	
	end if
		
	Set RSF = Server.CreateObject("ADODB.Recordset")
	SQLF = SQLF&" ORDER BY DA_Ano, NU_Seq"
	RSF.Open SQLF, CON7	
	
	wrkCount=0
	while not RSF.EOF
		anoHist = RSF("DA_Ano")
		codSeg = RSF("CO_Seg")
		nomEscola = RSF("NO_Escola")
		nomCidade = RSF("NO_Municipio")
		sgUf = RSF("SG_UF")
		inAprov = RSF("IN_Aprovado")
		
		wrkAprov = tbResultadoFinal(inAprov)
			
		if wrkCount=0 then
			wrkRetorno = anoHist&"#!#"&codSeg&"#!#"&nomEscola&"#!#"&nomCidade&"#!#"&sgUf&"#!#"&wrkAprov
		else
			wrkRetorno = wrkRetorno&"$!$"&anoHist&"#!#"&codSeg&"#!#"&nomEscola&"#!#"&nomCidade&"#!#"&sgUf&"#!#"&wrkAprov		
		end if
		wrkCount=wrkCount+1
	RSF.MOVENEXT
	WEND

	tbHistoricoAnoEscola = wrkRetorno
end function

Function etapaAnoHistorico (tipoHistorico, tipoRetorno)



	SQLE = "SELECT CO_Seg FROM TB_Segmento where TP_Curso ='"& tipoHistorico&"'"
	
	Set RSE = Server.CreateObject("ADODB.Recordset")
	SQLE = SQLE&" ORDER BY NU_Ordem"
	RSE.Open SQLE, CON7	
	
	wrkCount=0
	while not RSE.EOF

		codSeg = RSE("CO_Seg")
		
		if tipoRetorno="CO_SEG" then
		
            nom_etapa = codSeg
		else
			nom_etapa = etapaHistorico (tipoHistorico,codSeg,tipoRetorno)		
		end if	

		if wrkCount=0 then
			wrkRetorno = nom_etapa
		else
			wrkRetorno = wrkRetorno&"#!#"&nom_etapa	
		end if		
    	wrkCount=wrkCount+1	
	RSE.MOVENEXT
	WEND	

	etapaAnoHistorico = wrkRetorno
end function

Function HistoricoDisciplinas (codCons,tipoHistorico)

	wrkAnos = anoHistorico (codCons,tipoHistorico)
	vetorAnos= split(wrkAnos,"#!#")
	minAno=0
	for a=0 to ubound(vetorAnos)
		ano = vetorAnos(a)
		if not isnull(ano) and ano<>"" then
			if minAno=0 then
				minAno = ano
			end if				
			maxAno = ano
		end if		
	next

	Set RSD = Server.CreateObject("ADODB.Recordset")
	SQLD = "SELECT distinct NO_Materia FROM TB_Historico_Nota WHERE CO_Matricula ="&codCons&" and DA_Ano>= "&minAno&" and DA_Ano<= "&maxAno&" ORDER BY NO_Materia"
	RSD.Open SQLD, CON7	
	
	wrkCount=0
	while not RSD.EOF
		nome_materia = RSD("NO_Materia")
		if wrkCount=0 then
			wrkRetorno = nome_materia
		else
			wrkRetorno = wrkRetorno&"$!$"&nome_materia
		end if
		wrkCount=wrkCount+1
	RSD.MOVENEXT
	WEND

	HistoricoDisciplinas = wrkRetorno	

end function

Function anoHistorico (codCons,tipoHistorico)

	vetorEtapaAno = etapaAnoHistorico (tipoHistorico, "CO_SEG")
	
	vetorEtapas = split(vetorEtapaAno,"#!#")
	wrkCount=0	
	for e=0 to ubound(vetorEtapas)

		Set RSA = Server.CreateObject("ADODB.Recordset")
		SQLA = "SELECT DA_Ano FROM TB_Historico_Ano WHERE CO_Matricula ="&codCons&" AND TP_Curso = '"&tipoHistorico&"' and CO_Seg = '"&vetorEtapas(e)&"' order by DA_Ano"	
		RSA.Open SQLA, CON7	
	
		if RSA.EOF then
			wrkAnoHistorico = ""		
		else
	'	WHILE NOT RSA.EOF
			wrkAnoHistorico = RSA("DA_Ano")		
	
	
	'
	'
	'	RSA.MOVENEXT
	'	WEND	
		end if
		if wrkCount=0 then
			wrkRetorno = wrkAnoHistorico
		else
			wrkRetorno = wrkRetorno&"#!#"&wrkAnoHistorico
		end if	
		wrkCount=wrkCount+1				
    next

	anoHistorico = wrkRetorno	
	

end function

Function etapaHistorico (tipoHistorico,codEtapa,tipoResultado)

	Set RSE = Server.CreateObject("ADODB.Recordset")
	SQLE = "SELECT NO_Segmento, NO_Abreviado_Curso, CO_Conc FROM TB_Segmento WHERE TP_Curso = '"&tipoHistorico&"' and CO_Seg = '"&codEtapa&"'"
	RSE.Open SQLE, CON7	
	
	if tipoResultado = "S" then
		etapaHistorico = RSE("NO_Segmento")		
	elseif tipoResultado = "A" then
		etapaHistorico = RSE("NO_Abreviado_Curso")			
	elseif tipoResultado = "C" then
		etapaHistorico = RSE("CO_Conc")			
	end if

end function


Function tbHistoricoAnoNota (codCons, numAno, numSeq, nomDisciplina)

	if not isnull(numAno) and numAno<>"" then	

		Set RSN = Server.CreateObject("ADODB.Recordset")
		SQLN = "SELECT NU_Carga_Horaria, TX_Frequencia, VA_Nota, IN_Aprovado FROM TB_Historico_Nota WHERE CO_Matricula="&codCons&" AND NU_Seq=1 AND DA_Ano="&numAno&" AND NO_Materia ='"&nomDisciplina&"'"
		RSN.Open SQLN, CON7	

	
		if NOT RSN.EOF then
			carga = RSN("NU_Carga_Horaria")
			nota = RSN("VA_Nota")
			if (not isnull(nota) or nota<>"") and (isnull(carga) or carga="") then
				carga = "X"
			end if
			wrkRetorno = carga&"#!#"&RSN("TX_Frequencia")&"#!#"&nota&"#!#"&RSN("IN_Aprovado")
			
		end if	
	end if	
		
	tbHistoricoAnoNota = wrkRetorno
end function

Function tbHistoricoCarga(codCons, numAno, numSeq, vetDisciplina)

	disciplinasQuery = replace(vetDisciplina,"$!$","', '")

	Set RSC = Server.CreateObject("ADODB.Recordset")
	SQLC = "SELECT SUM(NU_Carga_Horaria) as Carga_Total FROM TB_Historico_Nota WHERE CO_Matricula="&codCons&" AND NU_Seq=1 AND DA_Ano="&numAno&" AND NO_Materia IN ('"&disciplinasQuery&"')"
	RSC.Open SQLC, CON7	
	
	if not RSC.EOF then
		wrkRetorno = RSC("Carga_Total")
	end if	
	
	tbHistoricoCarga = wrkRetorno
end function	

Function observacaoHistorico (codCons,tipoHistorico)

	Set RSA = Server.CreateObject("ADODB.Recordset")
	SQLA = "SELECT TX_Observacoes FROM TB_Historico_Ano WHERE CO_Matricula ="&codCons&" AND TP_Curso = '"&tipoHistorico&"' ORDER BY DA_Ano"
	RSA.Open SQLA, CON7	
	wrkCount=0
	WHILE NOT RSA.EOF
		wrkObsHistorico = RSA("TX_Observacoes")
		
		if not isnull(wrkObsHistorico) or wrkObsHistorico<>"" then
			if wrkCount=0 then
				wrkRetorno = wrkObsHistorico
			else
				wrkRetorno = wrkRetorno&"<br>"&wrkObsHistorico
			end if
			wrkCount=wrkCount+1	
		end if	

	RSA.MOVENEXT
	WEND

	observacaoHistorico = wrkRetorno	

end function

Function totaisHistorico (codCons,tipoHistorico,tipoInformacao)
	
	vetorEtapaAno = etapaAnoHistorico (tipoHistorico, "CO_SEG")	
	vetorEtapas = split(vetorEtapaAno,"#!#")
	wrkCount=0	
	for e=0 to ubound(vetorEtapas)

		Set RSA = Server.CreateObject("ADODB.Recordset")
		SQLA = "SELECT NU_Carga_Horaria_Total, TX_Frequencia_Total FROM TB_Historico_Ano WHERE CO_Matricula ="&codCons&" AND TP_Curso = '"&tipoHistorico&"' and CO_Seg = '"&vetorEtapas(e)&"' order by DA_Ano"	
		RSA.Open SQLA, CON7	
		
		if RSA.EOF then
			wrkInfHistorico = ""
		else

			if tipoInformacao = "F" then
				wrkInfHistorico = RSA("TX_Frequencia_Total")
			else
				wrkInfHistorico = RSA("NU_Carga_Horaria_Total")		
			end if	
		end if
		
		if wrkCount=0 then
			wrkRetorno = wrkInfHistorico
		else
			wrkRetorno = wrkRetorno&"#!#"&wrkInfHistorico
		end if
		wrkCount=wrkCount+1	

	NEXT

	totaisHistorico = wrkRetorno	

end function
%>