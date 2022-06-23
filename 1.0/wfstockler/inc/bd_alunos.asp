<%	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1

function buscaAluno(p_co_matricula)
    Set RSa = Server.CreateObject("ADODB.Recordset")
	SQLa = "SELECT * FROM TB_Alunos where CO_Matricula = "& p_co_matricula 
	RSa.Open SQLa, CON1

retorno = RSa("CO_Matricula")&"#!#"&RSa("RA_Aluno")&"#!#"&RSa("NO_Aluno")&"#!#"&RSa("NO_Apelido")&"#!#"&RSa("IN_Sexo")&"#!#"&RSa("IN_Desteridade")
retorno =retorno&"#!#"&RSa("CO_Nacionalidade")&"#!#"&RSa("CO_Pais_Natural")&"#!#"&RSa("SG_UF_Natural")&"#!#"&RSa("CO_Municipio_Natural")
retorno =retorno&"#!#"&RSa("TX_MSN")&"#!#"&RSa("TX_ORKUT")&"#!#"&RSa("CO_Raca")&"#!#"&RSa("CO_Religiao")&"#!#"&RSa("NO_Pai")&"#!#"&RSa("NO_Mae")
retorno =retorno&"#!#"&RSa("IN_Pai_Falecido")&"#!#"&RSa("IN_Mae_Falecida")&"#!#"&RSa("CO_Estado_Civil")&"#!#"&RSa("TP_Resp_Fin")
retorno =retorno&"#!#"&RSa("TP_Resp_Ped")&"#!#"&RSa("DA_Entrada_Escola")&"#!#"&RSa("DA_Cadastro")&"#!#"&("NO_Colegio_Origem")
retorno =retorno&"#!#"&RSa("NO_Serie_Cursada")&"#!#"&RSa("SG_UF_Cursada")&"#!#"&RSa("CO_Municipio_Cursada")

buscaAluno = retorno
end function

function buscaUCET(p_co_matricula,p_ano_letivo)

    Set RS3 = Server.CreateObject("ADODB.Recordset")
	SQL3 = "select * from TB_Matriculas where NU_Ano="& p_ano_letivo &" AND CO_Matricula = " & p_co_matricula
    RS3.Open SQL3, CON1

    if not RS3.eof then
        nu_unidade= RS3("NU_Unidade")
        co_curso= RS3("CO_Curso")
        co_etapa= RS3("CO_Etapa")
        co_turma= RS3("CO_Turma")

        buscaUCET = nu_unidade&"#!#"&co_curso&"#!#"&co_etapa&"#!#"&co_turma
    end if
end function

function buscaTipoResponsavelFinanceiro(p_co_matricula)
    	Set RSa = Server.CreateObject("ADODB.Recordset")
		SQLa = "SELECT TP_Resp_Fin FROM TB_Alunos where CO_Matricula = "& p_co_matricula 
		RSa.Open SQLa, CON1

IF not RSa.EOF THEN
    buscaTipoResponsavelFinanceiro = RSa("TP_Resp_Fin")
ELSE
	buscaTipoResponsavelFinanceiro = ""
END IF
end function

function buscaTipoResponsavelFinanceiro(p_co_matricula)
    	Set RSa = Server.CreateObject("ADODB.Recordset")
		SQLa = "SELECT TP_Resp_Fin FROM TB_Alunos where CO_Matricula = "& p_co_matricula 
		RSa.Open SQLa, CON1

IF not RSa.EOF THEN
    buscaTipoResponsavelFinanceiro = RSa("TP_Resp_Fin")
ELSE
	buscaTipoResponsavelFinanceiro = ""
END IF

end function

function buscaCodEstadoCivil(p_co_matricula)
    	Set RSc = Server.CreateObject("ADODB.Recordset")
		SQLc = "SELECT CO_Estado_Civil FROM TB_Alunos where CO_Matricula = "& p_co_matricula 
		RSc.Open SQLc, CON1

IF not RSc.EOF THEN
    buscaCodEstadoCivil = RSc("CO_Estado_Civil")
ELSE
	buscaCodEstadoCivil = ""
END IF
end function

function listaMatriculas(p_ano_letivo, p_unidade, p_curso, p_etapa, p_turma, p_situacao, p_ordem)

	if isnull(p_ano_letivo) then
		p_ano_letivo = session("ano_letivo")
	end if
	
	if not isnull(p_unidade) then
		queryU = " AND TB_Matriculas.NU_Unidade = "&p_unidade
	end if	 
	
	if not isnull(p_curso) then
		queryC = " AND TB_Matriculas.CO_Curso = '"&p_curso&"'"
	end if	
	
	if not isnull(p_etapa) then
		queryE = " AND TB_Matriculas.CO_Etapa = '"&p_etapa&"'"
	end if	
	
	if not isnull(p_turma) then
		queryT = " AND TB_Matriculas.CO_Turma = '"&p_turma&"'"
	end if	
	
	if not isnull(p_situacao) then	
		queryS = " AND TB_Matriculas.CO_Situacao = '"&p_situacao&"'"
	end if	
	
	if isnull(p_ordem) then
		p_ordem = "M"
	end if			
	 
	 orderBy = "TB_Matriculas.NU_Unidade, TB_Matriculas.CO_Curso,TB_Matriculas.CO_Etapa,TB_Matriculas.CO_Turma,"
	if p_ordem = "M" then
		orderBy = orderBy&"TB_Matriculas.CO_Matricula"
	else
		orderBy = orderBy&"TB_Alunos.NO_Aluno"	
	end if			

	Set RS3 = Server.CreateObject("ADODB.Recordset")
	SQL3 = "select * from TB_Matriculas INNER JOIN TB_Alunos ON TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula where TB_Matriculas.NU_Ano="& p_ano_letivo&queryU&queryC&queryE&queryT&queryS&" order by "&orderBy
    RS3.Open SQL3, CON1
	contaAluno = 0 
	while not RS3.eof
		if contaAluno = 0 then
			lista = RS3("CO_Matricula")
		else
			lista = lista&"#!#"&RS3("CO_Matricula")		
		end if
		contaAluno = contaAluno+1	
	RS3.movenext
	wend 
listaMatriculas	 = lista
end function
%>
