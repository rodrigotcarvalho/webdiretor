<!--#include file="utils.asp"-->
<%Function buscaDataPauta(caminhoBancoPauta, p_Co_prof, p_Unidade, p_Curso, p_Etapa, p_Turma, p_CO_Materia_Principal, p_CO_Materia, p_NU_Periodo, p_Vetor_Datas_Consulta, outro)


	wrkContaDatas = 0
	vetorDatas=""		
    vetorDatasConsulta=""
	Set CONPauta = Server.CreateObject("ADODB.Connection") 
	ABRIRPauta = "DBQ="& caminhoBancoPauta & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONPauta.Open ABRIRPauta
				
	Set RSP = Server.CreateObject("ADODB.Recordset")
	SQL = "Select TB_Pauta_Aula.DT_Aula from TB_Pauta INNER JOIN TB_Pauta_Aula on TB_Pauta.NU_Pauta=TB_Pauta_Aula.NU_Pauta WHERE CO_Professor  = "& p_Co_prof &" AND CO_Materia_Principal = '"& p_CO_Materia_Principal &"' AND CO_Materia = '"& p_CO_Materia &"' AND NU_Unidade  = "& p_Unidade &" AND CO_Curso  = '"& p_Curso &"' AND CO_Etapa  = '"& p_Etapa &"' AND CO_Turma  = '"& p_Turma &"' AND NU_Periodo = "& p_NU_Periodo&" GROUP BY TB_Pauta_Aula.DT_Aula ORDER BY TB_Pauta_Aula.DT_Aula "		
	Set RSP = CONPauta.Execute(SQL)
	
	while not RSP.EOF
		dataFormatadaVetor= split(RSP("DT_Aula"),"/")
		diaFormatado=dataFormatadaVetor(0)
		mesFormatado=dataFormatadaVetor(1)
		anoFormatado=dataFormatadaVetor(2)		
		if wrkContaDatas = 0 then
			vetorDatas = formataData(diaFormatado)&"/"&formataData(mesFormatado)
			vetorDatasConsulta = formataData(mesFormatado)&"/"&formataData(diaFormatado)&"/"& anoFormatado
		else
			vetorDatas = vetorDatas&"#!#"&formataData(diaFormatado)&"/"&formataData(mesFormatado)	
			vetorDatasConsulta = vetorDatasConsulta&"#!#"&formataData(mesFormatado)&"/"&formataData(diaFormatado)&"/"& anoFormatado			
		end if
		wrkContaDatas = wrkContaDatas +1
	RSP.MOVENEXT
	wend
	p_Vetor_Datas_Consulta = vetorDatasConsulta
	buscaDataPauta = vetorDatas	
End Function

Function buscaSeqDataPauta(caminhoBancoPauta, P_DATA_AULA, p_Co_prof, p_Unidade, p_Curso, p_Etapa, p_Turma, p_CO_Materia_Principal, p_CO_Materia, p_NU_Periodo, outro)


	wrkContaSeqs = 0
	vetorSeqs=""		
	Set CONPauta = Server.CreateObject("ADODB.Connection") 
	ABRIRPauta = "DBQ="& caminhoBancoPauta & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONPauta.Open ABRIRPauta
				
	Set RSP = Server.CreateObject("ADODB.Recordset")
	SQL = "Select TB_Pauta_Aula.NU_Pauta from TB_Pauta INNER JOIN TB_Pauta_Aula on TB_Pauta.NU_Pauta=TB_Pauta_Aula.NU_Pauta WHERE DT_Aula = #"&P_DATA_AULA&"# AND CO_Professor  = "& p_Co_prof &" AND CO_Materia_Principal = '"& p_CO_Materia_Principal &"' AND CO_Materia = '"& p_CO_Materia &"' AND NU_Unidade  = "& p_Unidade &" AND CO_Curso  = '"& p_Curso &"' AND CO_Etapa  = '"& p_Etapa &"' AND CO_Turma  = '"& p_Turma &"' AND NU_Periodo = "& p_NU_Periodo&" GROUP BY TB_Pauta_Aula.NU_Pauta "	
	Set RSP = CONPauta.Execute(SQL)
	
'	while not RSP.EOF	
'		if wrkContaSeqs = 0 then
'			vetorSeqs = RSP("NU_Pauta")
'		else
'			vetorSeqs = vetorSeqs&"#!#"&RSP("NU_Pauta")
'		end if
'		wrkContaSeqs = wrkContaSeqs +1
'	RSP.MOVENEXT
'	wend
'	buscaSeqDataPauta = vetorSeqs	
	if not RSP.EOF then
		Seq = RSP("NU_Pauta")
	end if
	buscaSeqDataPauta = Seq	
End Function

Function buscaSeqAula(caminhoBancoPauta, p_Nu_Data_pauta, P_DATA_AULA, outro)


	wrkContaSeq = 0
	vetorSeqs=""	
	Set CONPauta = Server.CreateObject("ADODB.Connection") 
	ABRIRPauta = "DBQ="& caminhoBancoPauta & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONPauta.Open ABRIRPauta
				
	Set RSP = Server.CreateObject("ADODB.Recordset")
	SQL = "Select TB_Pauta_Aula.NU_Seq from TB_Pauta_Aula WHERE NU_Pauta  = "& p_Nu_Data_pauta &" AND DT_Aula = #"&P_DATA_AULA&"#  ORDER BY TB_Pauta_Aula.NU_Seq "	
	Set RSP = CONPauta.Execute(SQL)
	
	while not RSP.EOF	
		if wrkContaSeq = 0 then
			vetorSeqs = RSP("NU_Seq")
		else
			vetorSeqs = vetorSeqs&"#!#"&RSP("NU_Seq")
		end if
		wrkContaSeq = wrkContaSeq +1
	RSP.MOVENEXT
	wend
	buscaSeqAula = vetorSeqs	
End Function

Function buscaTempoAula(caminhoBancoPauta, p_Nu_Data_pauta, p_Nu_Seq_pauta, outro)

	tempo=""	
	Set CONPauta = Server.CreateObject("ADODB.Connection") 
	ABRIRPauta = "DBQ="& caminhoBancoPauta & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONPauta.Open ABRIRPauta
				
	Set RSP = Server.CreateObject("ADODB.Recordset")
	SQL = "Select NU_Tempo from TB_Pauta_Aula WHERE NU_Pauta  = "& p_Nu_Data_pauta &" AND NU_Seq = "&p_Nu_Seq_pauta
	Set RSP = CONPauta.Execute(SQL)
	
	if not RSP.EOF then
		tempo = RSP("NU_Tempo")
	end if
	buscaTempoAula = tempo	
End Function

Function TotalFaltas(P_CAMINHO, p_data_pauta, p_nu_matricula, p_co_prof, p_unidade, p_curso, p_etapa, p_turma, p_mat_princ, p_co_materia, p_periodo, outro)
    seq_pauta = buscaSeqDataPauta(P_CAMINHO, p_data_pauta, p_co_prof, p_unidade, p_curso, p_etapa, p_turma, p_mat_princ, p_co_materia, p_periodo, outro)
    vetorSeqAula = buscaSeqAula(P_CAMINHO, seq_pauta, p_data_pauta, outro)
    seq_aula = Split(vetorSeqAula, "#!#")

	Set CONPauta = Server.CreateObject("ADODB.Connection") 
	ABRIRPauta = "DBQ="& P_CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONPauta.Open ABRIRPauta

    wrkTotalFaltas = 0

    For sa = 0 To UBound(seq_aula)
        RSP = Server.CreateObject("ADODB.Recordset")
        SQL = "Select * from TB_Pauta_Faltas WHERE CO_Matricula = " & p_nu_matricula & " AND NU_Pauta = " & seq_pauta & " AND NU_Seq = " & seq_aula(sa)
        set RSP = CONPauta.Execute(SQL)

        If Not RSP.EOF Then
            wrkTotalFaltas = wrkTotalFaltas + 1
        End If	
    Next

    TotalFaltas = wrkTotalFaltas
End Function
%>