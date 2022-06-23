<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/funcoes_comuns.asp"-->
<!--#include file="../inc/bd_parametros.asp"-->
<!--#include file="../inc/bd_alunos.asp"-->
<!--#include file="../inc/bd_contato.asp"-->
<!--#include file="../inc/bd_webfamilia.asp"-->
<%function dadosCabecalho(p_tp_contrato_adendo, p_cod_aluno)

tipo_resp_fin = buscaTipoResponsavelFinanceiro(p_cod_aluno)

ucet = buscaUCET(p_cod_aluno,session("ano_letivo"))
vetorUCET = split(ucet,"#!#")
nu_unidade =  vetorUCET(0)
co_curso = vetorUCET(1)
co_etapa = vetorUCET(2)
co_turma = vetorUCET(3)


no_unidade = GeraNomesNovaVersao("U",nu_unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)
no_curso = GeraNomesNovaVersao("C",co_curso,variavel2,variavel3,variavel4,variavel5,CON0,outro)
no_etapa = GeraNomesNovaVersao("E",co_curso,co_etapa,variavel3,variavel4,variavel5,CON0,outro)
prep_curso=GeraNomesNovaVersao("PC",co_curso,variavel2,variavel3,variavel4,variavel5,CON0,outro)
no_turno = buscaTurno(nu_unidade,co_curso,co_etapa,co_turma)

vetorProximaUcet = proximaUCET(nu_unidade,co_curso,co_etapa,co_turma)
dadosProximaUcet = split(vetorProximaUcet,"#!#")

prox_unidade = dadosProximaUcet(0)
prox_curso = dadosProximaUcet(1)
prox_etapa = dadosProximaUcet(2)
prox_turma = dadosProximaUcet(3)
prox_turno = dadosProximaUcet(4)

no_prox_unidade = GeraNomesNovaVersao("U",prox_unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)
no_prox_curso = GeraNomesNovaVersao("C",prox_curso,variavel2,variavel3,variavel4,variavel5,CON0,outro)
no_prox_etapa = GeraNomesNovaVersao("E",prox_curso,prox_etapa,variavel3,variavel4,variavel5,CON0,outro)
prep_prox_curso=GeraNomesNovaVersao("PC",prox_curso,variavel2,variavel3,variavel4,variavel5,CON0,outro)
no_prox_turno = buscaTurno(prox_unidade,prox_curso,prox_etapa,prox_turma)

ucet_prox_ano = no_prox_etapa&" "&prep_prox_curso&" "&no_prox_curso


vetorAluno = buscaAluno(p_cod_aluno)
dadosAluno = split(vetorAluno, "#!#")
no_aluno = Server.HTMLEncode(dadosAluno(2))

vetorContato = buscaContato (p_cod_aluno, tipo_resp_fin)
dadosContato = split(vetorContato, "#!#")
contratante = Server.HTMLEncode(dadosContato(2))
endereco = dadosContato(13)&","&dadosContato(14)&" "&dadosContato(15)
endereco = Server.HTMLEncode(endereco)

co_bairro = dadosContato(16)
co_cidade  = dadosContato(17)
cep  = dadosContato(19)
cep = LEFT(cep,5)&"-"&RIGHT(cep,3)
uf = dadosContato(18)
email  = dadosContato(8)
cpf  = dadosContato(4)
identidade = dadosContato(5)
orgao = dadosContato(6)
tel_res  = dadosContato(20)
tel_com = dadosContato(28)
co_ocupacao = dadosContato(9)

if not isnull(co_ocupacao) and co_ocupacao<>"" then
    no_ocupacao = buscaOcupacao(co_ocupacao)
    no_ocupacao = Server.HTMLEncode(no_ocupacao)
end if

if not isnull(uf) and uf<>"" and not isnull(co_cidade) and co_cidade<>""  then
        no_cidade = buscaCidade(uf, co_cidade)
        no_cidade = Server.HTMLEncode(no_cidade)
    if not isnull(co_bairro) and co_bairro<>"" then
        no_bairro = buscaBairro(uf, co_cidade, co_bairro)
        no_bairro = Server.HTMLEncode(no_bairro)
    end if
end if

if not isnull(orgao) and orgao<>"" then
    orgao = Server.HTMLEncode(orgao)
end if

if not isnull(no_unidade) and no_unidade<>"" then
    no_unidade = Server.HTMLEncode(no_unidade)
end if

if not isnull(no_curso) and no_curso<>"" then
    no_curso = Server.HTMLEncode(no_curso)
end if

if not isnull(no_etapa) and no_etapa<>"" then
    no_etapa = Server.HTMLEncode(no_etapa)
end if

if not isnull(no_turno) and no_turno<>"" then
    no_turno = Server.HTMLEncode(no_turno)
end if

    estadoCivil = "___________________________"
if not isnull(tipo_resp_fin) and tipo_resp_fin<>"" then
	if tipo_resp_fin = "PAI" or tipo_resp_fin = "MAE" then
		codEstadoCivil = buscaCodEstadoCivil(p_cod_aluno)
		estadoCivil = buscaEstadoCivil(codEstadoCivil)
	end if
end if

escolaLn= geraIdentificacaoEscola(p_tp_contrato_adendo,"Linha")

if left(p_tp_contrato_adendo,8) = "CONTRATO" then
    resultado = "<center><b>CONTRATO DE PRESTA&Ccedil;&Atilde;O DE SERVI&Ccedil;OS EDUCACIONAIS PARA "&session("ano_letivo")+1&"<br>"
	if p_tp_contrato_adendo = "CONTRATO_1A" then
		resultado = resultado&"CONTRATO 1A  -  ADENDO ESC2</b></center><br>&nbsp;<br>&nbsp;<br>"
	elseif p_tp_contrato_adendo = "CONTRATO_1B" then
		resultado = resultado&"CONTRATO 1B - ADENDO ----</b></center><br>&nbsp;<br>&nbsp;<br>"
	elseif p_tp_contrato_adendo = "CONTRATO_2A" then
		resultado = resultado&"CONTRATO 2A  -  ADENDO BL</b></center><br>&nbsp;<br>"
	elseif p_tp_contrato_adendo = "CONTRATO_2B" then
		resultado = resultado&"CONTRATO 2B  -  ADENDO BL</b></center><br>&nbsp;<br>&nbsp;<br>"
	elseif p_tp_contrato_adendo = "CONTRATO_3" then
		resultado = resultado&"CONTRATO 3  -  ADENDO CMP</b></center><br>&nbsp;<br>&nbsp;<br>"
	elseif p_tp_contrato_adendo = "CONTRATO_5" then
		resultado = resultado&"CONTRATO 5  -  ADENDO CMP</b></center><br>&nbsp;<br>&nbsp;<br>"
	elseif p_tp_contrato_adendo = "CONTRATO_8" then
		resultado = resultado&"CONTRATO 8  -  ADENDO ESC1</b></center><br>&nbsp;<br>&nbsp;<br>"
	elseif p_tp_contrato_adendo = "CONTRATO_G1" then
		resultado = resultado&"CONTRATO G1  -  ADENDO GAV A</b></center><br>&nbsp;<br>"
	elseif p_tp_contrato_adendo = "CONTRATO_G2" then
		resultado = resultado&"CONTRATO G2  -  ADENDO GAV B</b></center><br>&nbsp;<br>"
	elseif p_tp_contrato_adendo = "CONTRATO_G3" then
		resultado = resultado&"CONTRATO G3  -  ADENDO GAV C</b></center><br>&nbsp;<br>"
	elseif p_tp_contrato_adendo = "CONTRATO_G4" then
		resultado = resultado&"CONTRATO G4  -  ADENDO ----</b></center><br>&nbsp;<br>"								
	end if
	
    resultado = resultado&"<b>CONTRATANTE:</b> "&contratante&"<br><table width=100% border=0 cellspacing=0 cellpadding=0><tr><td width=60% >Endere&ccedil;o:"&endereco&"</td><td width=40% >Bairro:"&no_bairro&"</td></tr></table>"
    resultado = resultado&"<table width=100% border=0 cellspacing=0 cellpadding=0><tr><td width=30% >Cidade:"&no_cidade&"</td><td width=20% >CEP:"&cep&"</td><td width=10% >Estado:"&uf&"</td><td width=40% >E-mail:"&email&"</td></tr></table>"
    resultado = resultado&"<table width=100% border=0 cellspacing=0 cellpadding=0><tr><td width=20% >CPF:"&cpf&"</td><td width=40% >Carteira de Identidade n&ordm;:"&identidade&"</td><td width=40% >&Oacute;rg&atilde;o Expedidor:"&orgao&"</td></tr></table>"
    resultado = resultado&"<table width=100% border=0 cellspacing=0 cellpadding=0><tr><td width=80% > Telefones de Contato:"&tel_res&"</td><td width=10% >&nbsp;</td><td width=10% >&nbsp;</td></tr></table>"
    resultado = resultado&"<table width=100% border=0 cellspacing=0 cellpadding=0><tr><td width=50% >Estado Civil:"&estadoCivil&"</td><td width=50% >Profiss&atilde;o:"&no_ocupacao&"</td></tr></table>"
    resultado = resultado&"<table width=100% border=0 cellspacing=0 cellpadding=0><tr><td width=100% >Aluno(a) Benefici&aacute;rio(a):"&no_aluno&"</td></tr></table>"
    resultado = resultado&"<table width=100% border=0 cellspacing=0 cellpadding=0><tr><td width=100% >"&Server.HTMLEncode(ucet_prox_ano)&"</td></tr></table>"	
    'resultado = resultado&"<table width=100% border=0 cellspacing=0 cellpadding=0><tr><td width=50% >Turma:"&co_turma&"</td><td width=50% >Turno:"&no_turno&"</td></tr></table>"
	
	
elseif left(p_tp_contrato_adendo,6) = "ADENDO" then
    if p_tp_contrato_adendo = "ADENDO_2B" or p_tp_contrato_adendo = "ADENDO_3" then
		quebraLinha1 = "<br>&nbsp;<br>"
		quebraLinha2 = "<br>&nbsp;"		
	else
		quebraLinha1 = "<br>&nbsp;<br>"	
		quebraLinha2 = quebraLinha1
	end if
    resultado = "<center>"&geraLogo(p_tp_contrato_adendo)&"<b>"&quebraLinha1&"ADITIVO AO CONTRATO DE PRESTA&Ccedil;&Atilde;O DE SERVI&Ccedil;OS EDUCACIONAIS PARA "&session("ano_letivo")+1&"<BR> ENTRE AS PARTES J&Aacute; QUALIFICADAS</b></center>"&quebraLinha2
    resultado = resultado&escolaLn&" e  "&contratante&" "
    resultado = resultado&"respons&aacute;vel pelo aluno "&no_aluno&" matriculado no "&lcase(no_etapa)&" "&geraPreposicaoEtapa(p_tp_contrato_adendo)&" "&Server.HTMLEncode(geraNomeEtapa(p_tp_contrato_adendo))&", " 
    resultado = resultado&"aceitam as condi&ccedil;&otilde;es gerais abaixo discriminadas e acordam com elas."
end if
resultado = resultado&"<BR><BR>"

dadosCabecalho = resultado
end function

function dadosRodape(p_tp_contrato_adendo, p_cod_aluno)
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 

if dia<10 then
    dia = "0"&dia
end if

nomeDaEscola = geraNomeEscola(p_tp_contrato_adendo)	

no_mes = Nome_Mes(mes,"completo","S",outro)
assinatura = "_______________________________"
resultado = "<table width=70% align=center  border=0 cellspacing=0 cellpadding=0><tr><td>Rio de Janeiro. "&dia&" de "&no_mes&" de "&ano&"<br><br><br></td></tr></table>"
resultado = resultado&"<table width=70% align=center  border=0 cellspacing=0 cellpadding=0>"
resultado = resultado&"<tr><td width=50% align=center >"&assinatura&"</td><td width=50% align=center >"&assinatura&"</td></tr>"
if left(p_tp_contrato_adendo,8) = "CONTRATO" then
    resultado = resultado&"<tr><td width=50% align=center >Contratante</td><td width=50% align=center >"&Server.HTMLEncode(nomeDaEscola)&"</td></tr>"
    resultado = resultado&"</table>" 
    resultado = resultado&"<BR><BR><center>TESTEMUNHAS</center><br><br>"
    resultado = resultado&"<table width=70% align=center border=0 cellspacing=0 cellpadding=0>"
    resultado = resultado&"<tr><td width=50% align=center >1&ordf;)"&assinatura&"</td><td width=50% align=center >2&ordf;)"&assinatura&"</td></tr>"    
	if p_tp_contrato_adendo = "CONTRATO_1A" or p_tp_contrato_adendo = "CONTRATO_1B" or p_tp_contrato_adendo = "CONTRATO_2A" or p_tp_contrato_adendo = "CONTRATO_2B" or p_tp_contrato_adendo = "CONTRATO_3" or p_tp_contrato_adendo = "CONTRATO_5" or p_tp_contrato_adendo = "CONTRATO_8" or p_tp_contrato_adendo = "CONTRATO_G1" or p_tp_contrato_adendo = "CONTRATO_G2" or p_tp_contrato_adendo = "CONTRATO_G3" or p_tp_contrato_adendo = "CONTRATO_G4" then
		resultado = resultado&"<tr><td width=50% align=center >"&buscaTestemunha(1)&"</td><td width=50% align=center >"&buscaTestemunha(3)&"</td></tr>"	
	else	
		resultado = resultado&"<tr><td width=50% align=center >"&buscaTestemunha(1)&"</td><td width=50% align=center >"&buscaTestemunha(2)&"</td></tr>"
	end if
    resultado = resultado&"</table>"
elseif left(p_tp_contrato_adendo,6) = "ADENDO" then
    resultado = resultado&"<tr><td width=50% align=center >"&Server.HTMLEncode(nomeDaEscola)&"</td><td width=50% align=center >Contratante</td></tr>" 
    resultado = resultado&"</table>" 
end if

dadosRodape = resultado
end function

function buscaTestemunha(p_num_testemunha)

	if p_num_testemunha = 1 then
		resultado = "Nome: IZABEL CHRISTINA BORGES<BR>" 
		resultado = resultado&"CIC: 185.367.863-53<BR>"			
		resultado = resultado&"IDENTIDADE: 07926163-2 - IFP<BR>"
	elseif p_num_testemunha = 2 then
		resultado = "Nome: SOR&Aacute;IA REZENDE NUNES<BR>"
		resultado = resultado&"CIC: 754.514.097-49<BR>"
		resultado = resultado&"IDENTIDADE: 05730026-1 - IFP<BR>"	
	elseif p_num_testemunha = 3 then
		resultado = "Nome: ZILDA COTIAS PLOMBON<BR>"
		resultado = resultado&"CIC: 389.446.777-00<BR>"
		resultado = resultado&"IDENTIDADE: 1933310-3 - DETRAN<BR>"			
	end if

buscaTestemunha = resultado
end function

function geraIdentificacaoEscola(p_tp_contrato_adendo, p_tp_retorno)

	if p_tp_contrato_adendo = "ADENDO_1A" or p_tp_contrato_adendo = "ADENDO_8" then
		endEscolaBr = "<br />Rua Visconde de Ouro Preto, 51<br />Botafogo - RJ"
		endEscola	 = ", Rua Visconde de Ouro Preto, 51"		
	elseif p_tp_contrato_adendo = "ADENDO_2A" or p_tp_contrato_adendo = "ADENDO_2B" then
		endEscolaBr = "<br />Rua Bar&atilde;o de Lucena, 31 e 37<br />Botafogo - RJ"
		endEscola	 = ", Rua Bar&atilde;o de Lucena, 31 e 37"		
	elseif p_tp_contrato_adendo = "ADENDO_3" then
		endEscolaBr = "<br />Rua Visconde de Ouro Preto, 54/56<br />Botafogo - RJ"
		endEscola	 = ", Rua Visconde de Ouro Preto, 54"			
	elseif p_tp_contrato_adendo = "ADENDO_5" then
		endEscolaBr = "<br />Rua Marqu&ecirc;s de Pinedo, 26<br />Laranjeiras - RJ"	
		endEscola	 = ", Rua Marqu&ecirc;s de Pinedo, 26"		
	elseif p_tp_contrato_adendo = "ADENDO_G1" then
		endEscolaBr = "<br />Rua General Rabelo, 51<br />G&aacute;vea - RJ"	
		endEscola	 = ", Rua General Rabelo, 51"	
	elseif p_tp_contrato_adendo = "ADENDO_G2" or p_tp_contrato_adendo = "ADENDO_G3" then
		endEscolaBr = "<br />Rua General Rabelo, 56<br />G&aacute;vea - RJ"
		endEscola	 = ", Rua General Rabelo, 56"					
	end if

	nomeEscola = geraNomeEscola(p_tp_contrato_adendo)	

	if p_tp_retorno="Linha" then
		geraIdentificacaoEscola = Server.HTMLEncode(nomeEscola)&endEscola
	else
		geraIdentificacaoEscola = Server.HTMLEncode(UCASE(nomeEscola))&endEscolaBr
	end if

end function


function geraLogo(p_tp_contrato_adendo)
	escolaTb = geraIdentificacaoEscola(p_tp_contrato_adendo,"Tabela")
	
	resultado = "<table width=100% border=0 cellspacing=0 cellpadding=0>"
	resultado = resultado&"<tr>"
	resultado = resultado&"<td width=50% align=right><img src=../img/logo_pdf.gif height=50 /></td>"
	resultado = resultado&"<td width=50% ><strong>"&escolaTb&"</strong></td>"
	resultado = resultado&"</tr>"
	resultado = resultado&"</table>"
	geraLogo = resultado
end function

function geraNomeEscola(p_tp_contrato_adendo)

	if p_tp_contrato_adendo = "CONTRATO_1A" or p_tp_contrato_adendo = "CONTRATO_1B" or p_tp_contrato_adendo = "CONTRATO_8" or p_tp_contrato_adendo = "ADENDO_1A" or p_tp_contrato_adendo = "ADENDO_8" then
		escola	 = "Escola Dínamis Ltda."		
	elseif p_tp_contrato_adendo = "CONTRATO_2A" or p_tp_contrato_adendo = "CONTRATO_2B" OR p_tp_contrato_adendo = "ADENDO_2A" or p_tp_contrato_adendo = "ADENDO_2B" then
		escola	 = "Jardim Escola B.L. Ltda."		
	elseif p_tp_contrato_adendo = "CONTRATO_3" or  p_tp_contrato_adendo = "ADENDO_3" then
		escola	 = "Creche Experimental Dínamis Ltda."			
	elseif p_tp_contrato_adendo = "CONTRATO_5" or p_tp_contrato_adendo = "ADENDO_5" then
		escola	 = "Jardim Escola M.P. Ltda."		
	elseif p_tp_contrato_adendo = "CONTRATO_G1" or  p_tp_contrato_adendo = "ADENDO_G1" then
		escola	 = "Jardim Escola Stockler Ltda."	
	elseif p_tp_contrato_adendo = "CONTRATO_G2" or p_tp_contrato_adendo = "CONTRATO_G3" or p_tp_contrato_adendo = "CONTRATO_G4" or p_tp_contrato_adendo = "ADENDO_G2" or p_tp_contrato_adendo = "ADENDO_G3" then
		escola	 = "Colégio L. Stockler Ltda."					
	end if
	geraNomeEscola = escola
end function

function geraNomeEtapa(p_tp_contrato_adendo)

	if p_tp_contrato_adendo = "ADENDO_1A" or p_tp_contrato_adendo = "ADENDO_8" or p_tp_contrato_adendo = "ADENDO_G2" or p_tp_contrato_adendo = "ADENDO_G3" then
		nomeEtapa	 = "Ensino Fundamental"		
	elseif p_tp_contrato_adendo = "ADENDO_2A" or p_tp_contrato_adendo = "ADENDO_2B" or p_tp_contrato_adendo = "ADENDO_3"  or p_tp_contrato_adendo = "ADENDO_5" or p_tp_contrato_adendo = "ADENDO_G1" then
		nomeEtapa	 = "Educação Infantil"								
	end if
	geraNomeEtapa = nomeEtapa
end function

function geraPreposicaoEtapa(p_tp_contrato_adendo)

	if p_tp_contrato_adendo = "ADENDO_1A" or p_tp_contrato_adendo = "ADENDO_8" or p_tp_contrato_adendo = "ADENDO_G2" or p_tp_contrato_adendo = "ADENDO_G3" then
		prepEtapa	 = "do"		
	elseif p_tp_contrato_adendo = "ADENDO_2A" or p_tp_contrato_adendo = "ADENDO_2B" or p_tp_contrato_adendo = "ADENDO_3" or  p_tp_contrato_adendo = "ADENDO_5" or p_tp_contrato_adendo = "ADENDO_G1" then
		prepEtapa	 = "da"											
	end if
	geraPreposicaoEtapa = prepEtapa
end function

%>