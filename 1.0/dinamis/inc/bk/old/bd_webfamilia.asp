<%
    
Set conw = Server.CreateObject("ADODB.Connection") 
ABRIR = "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
conw.Open ABRIR

function vencimentoRematricula(p_ano_letivo)
    Set RSano = Server.CreateObject("ADODB.Recordset")
	SQLano = "SELECT DT_Bloqueto_Rematricula FROM TB_Ano_Letivo where NU_Ano_Letivo='"&p_ano_letivo&"'"
	RSano.Open SQLano, conw
    
    vencimentoRematricula=RSano("DT_Bloqueto_Rematricula")
end function

function gravaRematricula(P_CO_Aluno)

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 
data = dia &"/"& mes &"/"& ano
hora = hora & ":"& min

    Set RSrem = Server.CreateObject("ADODB.Recordset")
	SQLrem = "SELECT * FROM TB_Aunos_Rematriculados where CO_Matricula_Escola="&P_CO_Aluno
	RSrem.Open SQLrem, conw
	
	if not RSrem.EOF then
		Set RS0 = Server.CreateObject("ADODB.Recordset")	
		SQL0 = "delete * FROM TB_Aunos_Rematriculados where CO_Matricula_Escola="&P_CO_Aluno
		Set RS0 = conw.Execute(SQL0)		
	end if
	
	Set RSrem = server.createobject("adodb.recordset")
	RSrem.open "TB_Aunos_Rematriculados", conw, 2, 2 'which table do you want open
	RSrem.addnew	
		RSrem("CO_Matricula_Escola") = P_CO_Aluno
		RSrem("DA_Ult_Acesso") = data
		RSrem("HO_ult_Acesso") = hora
	RSrem.update
	set RSrem=nothing	
end function

Function ehPeriodoRematricula()
    dia = DatePart("d", now) 
	mes = DatePart("m", now) 
	ano = DatePart("yyyy", now)

hoje = mes&"/"&dia&"/"&ano

		Set RSRA = Server.CreateObject("ADODB.Recordset")
		SQLRA = "SELECT * FROM TB_Ano_Letivo where DT_Inicio_Rematricula <= #"&hoje&"# AND  (DT_Final_Rematricula >= #"&hoje&"# or DT_Final_Rematricula is null) AND NU_Ano_Letivo = '"&session("ano_letivo")&"'"
        RSRA.Open SQLRA, conw

if not RSRA.EOF THEN		
	ehPeriodoRematricula = "S"
ELSE
	ehPeriodoRematricula = "N"
END IF	
end function

Function tipoResponsavel(CO_USUARIO)

		Set RSRA = Server.CreateObject("ADODB.Recordset")
		SQLRA = "SELECT * FROM TB_RespxAluno where CO_USUARIO = "&CO_USUARIO
		RSRA.Open SQLRA, conw

if not RSRA.EOF THEN		
	tipoResponsavel = RSRA("TP_Resp")
ELSE
	tipoResponsavel = ""
END IF	
end function


Function modeloContratoAdendo(P_Unidade,P_Curso,P_Etapa,P_Turma,P_Contrato_Adendo)

		Set RSRB = Server.CreateObject("ADODB.Recordset")
		'SQLRB = "SELECT * FROM TB_Modelos_Matricula where Unidade = '"&P_Unidade&"' AND Curso ='"&P_Curso&"' and Etapa ='"&P_Etapa&"' and Turma = '"&P_Turma&"'"
		SQLRB = "SELECT * FROM TB_Modelos_Matricula where Unidade = '"&P_Unidade&"' AND Curso ='"&P_Curso&"' and Etapa ='"&P_Etapa&"'"
        RSRB.Open SQLRB, conw

if not RSRB.EOF THEN

    if P_Contrato_Adendo = "C" then		
	    modeloContratoAdendo = RSRB("Contrato")
    else
		if isnull(RSRB("Adendo")) or RSRB("Adendo") = "" then
			modeloContratoAdendo = "X"
		else
			modeloContratoAdendo = RSRB("Adendo")
		end if	

    end if
ELSE
	modeloContratoAdendo = ""
END IF	

end function

Function turno(P_Unidade,P_Curso,P_Etapa,P_Turma,P_Contrato_Adendo)

		Set RSRT = Server.CreateObject("ADODB.Recordset")
		SQLRT = "SELECT * FROM TB_Modelos_Matricula where Unidade = "&P_Unidade&" AND Curso ='"&P_Curso&"' and Etapa ='"&P_Etapa&"' and Turma = '"&P_Turma&"'"
		RSRT.Open SQLRT, conw

if not RSRT.EOF THEN
    turno = RSRT("Turno")
ELSE
	turno = ""
END IF	
end function

Function proximaUCET(P_Unidade,P_Curso,P_Etapa,P_Turma)

		Set RSRP = Server.CreateObject("ADODB.Recordset")
		SQLRP = "SELECT * FROM TB_Modelos_Matricula where Unidade ='"&P_Unidade&"' AND Curso ='"&P_Curso&"' and Etapa ='"&P_Etapa&"' and Turma = '"&P_Turma&"'"
        RSRP.Open SQLRP, conw

if not RSRP.EOF THEN
    proximaUCET = RSRP("P_Unidade")&"#!#"&RSRP("P_Curso")&"#!#"&RSRP("P_Etapa")&"#!#"&RSRP("P_Turma")&"#!#"&RSRP("P_Turno")
ELSE
    proximaUCET = "#!#"&"#!#"&"#!#"&"#!#"
END IF	
end function

function buscaResponsavelFinanceiro(P_CO_Aluno)

		Set RSRF = Server.CreateObject("ADODB.Recordset")
		SQLRF = "SELECT * FROM TB_RespxAluno where CO_Aluno = "&P_CO_Aluno&" AND TP_Resp = 'F'"
		RSRF.Open SQLRF, conw

if not RSRF.EOF THEN
    buscaResponsavelFinanceiro = RSRF("CO_Usuario")
ELSE
	buscaResponsavelFinanceiro = ""
END IF	
end function
%>
