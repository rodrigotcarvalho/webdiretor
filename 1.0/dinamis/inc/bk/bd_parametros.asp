<%

Set CON0 = Server.CreateObject("ADODB.Connection") 
ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
CON0.Open ABRIR0

function buscaMulta

	Set RSP = Server.CreateObject("ADODB.Recordset")
	SQLP ="SELECT VA_Multa FROM TB_Correcao"
	RSP.Open SQLP, CON0	
	
	IF EOF THEN
		multa=0
	ELSE
		multa=RSP("VA_Multa")	
	END IF
	buscaMulta = multa	
end function	

function buscaMora

    Set RSP = Server.CreateObject("ADODB.Recordset")
	SQLP ="SELECT VA_Mora FROM TB_Correcao"
	RSP.Open SQLP, CON0	
	
	IF EOF THEN
		mora=0
	ELSE
		mora=RSP("VA_Mora")
	END IF	
	buscaMora = mora
end	function

function CalculaMulta(vencimento, data_calc, val_original)
	qtd_dias = DateDiff("d",vencimento,data_calc)
	if qtd_dias>0 then
		fatorMulta = buscaMulta
		CalculaMulta = round(val_original*(fatorMulta/100),2)
	end if
end function


function CalculaMora(vencimento, data_calc, val_original)
	qtd_dias = DateDiff("d",vencimento,data_calc)
	if qtd_dias>0 then
		fatorMora = buscaMora
		val_mora = val_original*(fatorMora/100)*qtd_dias
		val_mora = val_mora*100
		val_mora = INT(val_mora)
		val_mora = val_mora/100
		CalculaMora = val_mora
	end if	
end function

function buscaOcupacao(p_cod_ocupa)

	Set RSR0 = Server.CreateObject("ADODB.Recordset")
	SQLR0 = "SELECT * FROM TB_Ocupacoes where CO_Ocupacao = "&p_cod_ocupa
    RSR0.Open SQLR0, CON0


buscaOcupacao = RSR0("NO_Ocupacao")

end function


function buscaCidade(p_sg_uf, p_cod_cidade)

	Set RS0 = Server.CreateObject("ADODB.Recordset")
	SQL0 = "SELECT * FROM TB_Municipios where SG_UF = '"&p_sg_uf&"' AND CO_Municipio = "&p_cod_cidade
    RS0.Open SQL0, CON0


buscaCidade = RS0("NO_Municipio")

end function

function buscaBairro(p_sg_uf, p_cod_cidade, p_cod_bairro)

	Set RS0 = Server.CreateObject("ADODB.Recordset")
	SQL0 = "SELECT * FROM TB_Bairros where SG_UF = '"&p_sg_uf&"' AND CO_Municipio = "&p_cod_cidade&" AND CO_Bairro = "&p_cod_bairro
    RS0.Open SQL0, CON0


buscaBairro = RS0("NO_Bairro")

end function

function buscaEstadoCivil(p_cod_estado_civil)

	if NOT (p_cod_estado_civil="" or isnull(p_cod_estado_civil)) then
		Set RSE = Server.CreateObject("ADODB.Recordset")
		SQLE = "SELECT TX_Estado_Civil FROM TB_Estado_Civil where CO_Estado_Civil = '"&p_cod_estado_civil&"'"
		RSE.Open SQLE, CON0
	
	
		buscaEstadoCivil = RSE("TX_Estado_Civil")
	end if
	
	
end function

function buscaTurno(p_nu_unidade, p_co_curso, p_co_etapa, p_co_turma)

    Set RS3 = Server.CreateObject("ADODB.Recordset")
	SQL3 = "SELECT * FROM TB_Turma WHERE NU_Unidade="& p_nu_unidade&" and CO_Curso='"& p_co_curso &"' AND CO_Etapa = '"&p_co_etapa&"' and CO_Turma ='"&p_co_turma&"'"		
	RS3.Open SQL3, CON0	
	
	if not RS3.EOF then

		co_turno = RS3("CO_Turno")
	
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Turno WHERE CO_Turno='"& co_turno&"'"
		RS4.Open SQL4, CON0	
	
		if not RS4.EOF then
		buscaTurno = RS4("NO_Turno")
		
		end if
	end if
end function
%>
