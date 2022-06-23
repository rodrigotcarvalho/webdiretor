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

if isnull(no_prox_etapa) or no_prox_etapa="" then
    no_prox_etapa = prox_etapa
end if

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
	nome_adendo = modeloContratoAdendo(nu_unidade,co_curso,co_etapa,co_turma,"A")

	nome_contrato = replace(p_tp_contrato_adendo, "_", " ")
	
	if isnull(nome_adendo) then
		nome_adendo = "----"
	else
		nome_adendo = replace(nome_adendo, "_", " ")
	end if
	resultado = resultado&nome_contrato&"  -  "&nome_adendo&"</b></center><br>&nbsp;<br>&nbsp;<br>"	
	'if p_tp_contrato_adendo = "CONTRATO_1A" then
	'	resultado = resultado&"CONTRATO 1A  -  ADENDO ESC2</b></center><br>&nbsp;<br>&nbsp;<br>"
	'elseif p_tp_contrato_adendo = "CONTRATO_1B" then
	'	resultado = resultado&"CONTRATO 1B - ADENDO ----</b></center><br>&nbsp;<br>&nbsp;<br>"
	'elseif p_tp_contrato_adendo = "CONTRATO_2A" then
	'	resultado = resultado&"CONTRATO 2A  -  ADENDO BL</b></center><br>&nbsp;<br>"
	'elseif p_tp_contrato_adendo = "CONTRATO_2B" then
	'	resultado = resultado&"CONTRATO 2B  -  ADENDO BL</b></center><br>&nbsp;<br>&nbsp;<br>"
	'elseif p_tp_contrato_adendo = "CONTRATO_3" then
	'	resultado = resultado&"CONTRATO 3  -  ADENDO CMP</b></center><br>&nbsp;<br>&nbsp;<br>"
	'elseif p_tp_contrato_adendo = "CONTRATO_5" then
	'	resultado = resultado&"CONTRATO 5  -  ADENDO CMP</b></center><br>&nbsp;<br>&nbsp;<br>"
	'elseif p_tp_contrato_adendo = "CONTRATO_8" then
	'	resultado = resultado&"CONTRATO 8  -  ADENDO ESC1</b></center><br>&nbsp;<br>&nbsp;<br>"
	'elseif p_tp_contrato_adendo = "CONTRATO_G1" then
	'	resultado = resultado&"CONTRATO G1  -  ADENDO GAV A</b></center><br>&nbsp;<br>"
	'elseif p_tp_contrato_adendo = "CONTRATO_G2" then
	'	resultado = resultado&"CONTRATO G2  -  ADENDO GAV B</b></center><br>&nbsp;<br>"
	'elseif p_tp_contrato_adendo = "CONTRATO_G3" then
	'	resultado = resultado&"CONTRATO G3  -  ADENDO GAV C</b></center><br>&nbsp;<br>"
	'elseif p_tp_contrato_adendo = "CONTRATO_G4" then
	'	resultado = resultado&"CONTRATO G4  -  ADENDO ----</b></center><br>&nbsp;<br>"								
	'end if
	
    resultado = resultado&"<b>CONTRATANTE:</b> "&contratante&"<br>Endere&ccedil;o: "&SpacePad(endereco,"&nbsp;",100,"R")&" Bairro:"&no_bairro
    resultado = resultado&"<br>Cidade: "&SpacePad(no_cidade,"&nbsp;",60, "R")&" CEP: "&SpacePad(cep,"&nbsp;",14, "R")&" Estado: "&SpacePad(uf,"&nbsp;",4, "R")&" E-mail:"&email
    resultado = resultado&"<br>CPF: "&SpacePad(cpf,"&nbsp;",30, "R")&" Carteira de Identidade n&ordm;: "&SpacePad(identidade,"&nbsp;",34, "R")&" &Oacute;rg&atilde;o Expedidor: "&orgao
    resultado = resultado&"<br>Telefones de Contato: "&tel_res
    resultado = resultado&"<br>Estado Civil: "&SpacePad(estadoCivil,"&nbsp;",49, "R")&" Profiss&atilde;o:"&no_ocupacao
    resultado = resultado&"<br>Aluno(a) Benefici&aacute;rio(a): "&no_aluno
    resultado = resultado&"<br>"&Server.HTMLEncode(ucet_prox_ano)
    'resultado = resultado&"<table width=100% border=0 cellspacing=0 cellpadding=0><tr><td width=50% >Turma:"&co_turma&"</td><td width=50% >Turno:"&no_turno&"</td></tr></table>"
	if prox_curso=0 then
		resultado = resultado&", Turma:____________________ Turno:____________________ "	
	elseif p_tp_contrato_adendo = "CONTRATO1A" or p_tp_contrato_adendo = "CONTRATO1A1" or p_tp_contrato_adendo = "CONTRATO1B" or p_tp_contrato_adendo = "CONTRATO1B1" or p_tp_contrato_adendo = "CONTRATO8" or p_tp_contrato_adendo = "CONTRATOG1" then
		resultado = resultado&", Turno:____________________ "
	end if	
	
	
elseif left(p_tp_contrato_adendo,6) = "ADENDO" then
    if p_tp_contrato_adendo = "ADENDO_3" then
		quebraLinha1 = "<br>&nbsp;<br>&nbsp;<br>"
		quebraLinha2 = quebraLinha1	
	else
		quebraLinha1 = "<br>&nbsp;<br>&nbsp;<br>"	
		quebraLinha2 = quebraLinha1
	end if
    resultado = "<center><b>"&quebraLinha1&"ADITIVO AO CONTRATO DE PRESTA&Ccedil;&Atilde;O DE SERVI&Ccedil;OS EDUCACIONAIS PARA "&session("ano_letivo")+1&"<BR> ENTRE AS PARTES J&Aacute; QUALIFICADAS</b></center>"&quebraLinha2
    resultado = resultado&escolaLn&" e  "&contratante&" "
    resultado = resultado&"respons&aacute;vel pelo aluno "&no_aluno&" matriculado no "&Server.HTMLEncode(no_prox_etapa)&" "&geraPreposicaoEtapa(p_tp_contrato_adendo, prox_curso)&" "&Server.HTMLEncode(geraNomeEtapa(p_tp_contrato_adendo, prox_curso))&", " 
    resultado = resultado&"aceitam as condi&ccedil;&otilde;es gerais abaixo discriminadas e acordam com elas."
end if
resultado = resultado&"<BR><BR>"

dadosCabecalho = resultado
end function

function converte_nome_adendo(p_nu_unidade,p_co_curso,p_co_etapa,p_co_turma,tp_contrato_adendo)

vetorProximaUcet = proximaUCET(p_nu_unidade,p_co_curso,p_co_etapa,p_co_turma)
dadosProximaUcet = split(vetorProximaUcet,"#!#")

prox_unidade = dadosProximaUcet(0)
prox_curso = dadosProximaUcet(1)
prox_etapa = dadosProximaUcet(2)
prox_turma = dadosProximaUcet(3)
prox_turno = dadosProximaUcet(4)

if tp_contrato_adendo = "ADENDOESC2" AND session("ano_letivo") <=2017 then
  	prox_unidade = prox_unidade*1
	prox_curso = prox_curso*1
	prox_etapa = prox_etapa*1
	
	if (prox_unidade=1 and prox_curso=1 and prox_etapa >= 7) or (prox_unidade=8 and prox_curso=1 and prox_etapa =6) then
		converte_nome_adendo = "ADENDO_1A"
	elseif prox_unidade=1 and prox_curso=2 and prox_etapa =1 then
		converte_nome_adendo = "ADENDO_1B"	
	end if
elseif tp_contrato_adendo = "ADENDOESC2" then
	converte_nome_adendo = "ADENDO_12"		
elseif tp_contrato_adendo = "ADENDOESC3" then
	converte_nome_adendo = "ADENDO_1B"	
elseif tp_contrato_adendo = "ADENDOBL1" then
	converte_nome_adendo = "ADENDO_2A"	
elseif tp_contrato_adendo = "ADENDOBL2" then
	converte_nome_adendo = "ADENDO_2C"		
elseif tp_contrato_adendo = "ADENDOBL3" then
	converte_nome_adendo = "ADENDO_2B"	
elseif tp_contrato_adendo = "ADENDOCMP1" then
	converte_nome_adendo = "ADENDO_3"		
elseif tp_contrato_adendo = "ADENDOCMP2" then
	converte_nome_adendo = "ADENDO_3A"		
elseif tp_contrato_adendo = "ADENDOCMP3" then
	converte_nome_adendo = "ADENDO_3B"			
elseif tp_contrato_adendo = "ADENDOESC1" then
	converte_nome_adendo = "ADENDO_8"	
elseif tp_contrato_adendo = "ADENDOGAVA" then
	converte_nome_adendo = "ADENDO_1A"	
elseif tp_contrato_adendo = "ADENDOGAVB" then
	converte_nome_adendo = "ADENDO_1B"	
elseif tp_contrato_adendo = "ADENDOGAVC" then
	converte_nome_adendo = "ADENDO_1C"	
elseif tp_contrato_adendo = "ADENDOCR1" then
	converte_nome_adendo = "ADENDO_31"		
elseif tp_contrato_adendo = "ADENDOCR2" then
	converte_nome_adendo = "ADENDO_32"	
elseif tp_contrato_adendo = "ADENDOBM1" then
	converte_nome_adendo = "ADENDO_2A"	
elseif tp_contrato_adendo = "ADENDOBM2" then
	converte_nome_adendo = "ADENDO_22"			
elseif tp_contrato_adendo = "ADENDOBM3" then
	converte_nome_adendo = "ADENDO_2B"	
end if
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
endereco = ""
resultado = SpacePad(endereco,"&nbsp;",10,"L")&"Rio de Janeiro. "&dia&" de "&no_mes&" de "&ano&"<BR>&nbsp;<BR>&nbsp;<BR>&nbsp;"
resultado = resultado&SpacePad(endereco,"&nbsp;",30,"L")&assinatura&SpacePad(endereco,"&nbsp;",20,"L")&assinatura&"<BR>"
if left(p_tp_contrato_adendo,8) = "CONTRATO" then
    resultado = resultado&SpacePad(endereco,"&nbsp;",50,"L")&"Contratante"&SpacePad(endereco,"&nbsp;",60,"L")&Server.HTMLEncode(nomeDaEscola)

	'If para retirar as testemunhas do contrato a partir de 2021
	if session("ano_letivo") <2020 then
		resultado = resultado&"<BR>&nbsp;<BR>&nbsp;<center>TESTEMUNHAS</center><BR>&nbsp;<BR>&nbsp;<BR>&nbsp;"
		resultado = resultado&SpacePad(endereco,"&nbsp;",30,"L")&assinatura&SpacePad(endereco,"&nbsp;",20,"L")&assinatura &"<BR>"  
		if p_tp_contrato_adendo = "CONTRATO_1A" or p_tp_contrato_adendo = "CONTRATO_1B" or p_tp_contrato_adendo = "CONTRATO_2A" or p_tp_contrato_adendo = "CONTRATO_2B" or p_tp_contrato_adendo = "CONTRATO_3" or p_tp_contrato_adendo = "CONTRATO_5" or p_tp_contrato_adendo = "CONTRATO5A" or p_tp_contrato_adendo = "CONTRATO5B" or p_tp_contrato_adendo = "CONTRATO_8" or p_tp_contrato_adendo = "CONTRATO_G1" or p_tp_contrato_adendo = "CONTRATO_G2" or p_tp_contrato_adendo = "CONTRATO_G3" or p_tp_contrato_adendo = "CONTRATO_G4" or p_tp_contrato_adendo = "CONTRATO1A" or p_tp_contrato_adendo = "CONTRATO1B" or p_tp_contrato_adendo = "CONTRATO2A" or p_tp_contrato_adendo = "CONTRATO2B" or p_tp_contrato_adendo = "CONTRATO3" or p_tp_contrato_adendo = "CONTRATO8" or p_tp_contrato_adendo = "CONTRATOG1" then
			resultado = resultado&SpacePad(buscaTestemunha(1,"N"),"&nbsp;",60,"L")&SpacePad(buscaTestemunha(3,"N"),"&nbsp;",47,"L")&"<BR>"	
			resultado = resultado&SpacePad(buscaTestemunha(1,"C"),"&nbsp;",50,"L")&SpacePad(buscaTestemunha(3,"C"),"&nbsp;",67,"L")&"<BR>"		
			resultado = resultado&SpacePad(buscaTestemunha(1,"I"),"&nbsp;",59,"L")&SpacePad(buscaTestemunha(3,"I"),"&nbsp;",57,"L")&"<BR>"					
		else	
			resultado = resultado&SpacePad(buscaTestemunha(1,"N"),"&nbsp;",60,"L")&SpacePad(buscaTestemunha(2,"N"),"&nbsp;",47,"L")&"<BR>"	
			resultado = resultado&SpacePad(buscaTestemunha(1,"C"),"&nbsp;",50,"L")&SpacePad(buscaTestemunha(2,"C"),"&nbsp;",67,"L")&"<BR>"		
			resultado = resultado&SpacePad(buscaTestemunha(1,"I"),"&nbsp;",59,"L")&SpacePad(buscaTestemunha(2,"I"),"&nbsp;",60,"L")&"<BR>"	
		end if
	end if	
elseif left(p_tp_contrato_adendo,6) = "ADENDO" then
'	if p_tp_contrato_adendo = "ADENDO_3" then
'		resultado = resultado&SpacePad(endereco,"&nbsp;",35,"L")&Server.HTMLEncode(nomeDaEscola)&SpacePad(endereco,"&nbsp;",45,"L")&"Contratante" 	
'	else
		resultado = resultado&SpacePad(endereco,"&nbsp;",45,"L")&Server.HTMLEncode(nomeDaEscola)&SpacePad(endereco,"&nbsp;",55,"L")&"Contratante" 
'	end if
end if

dadosRodape = resultado
end function

function buscaTestemunha(p_num_testemunha, p_tipo)

	if p_num_testemunha = 1 then
		if p_tipo = "N" then
			resultado = "Nome: IZABEL CHRISTINA BORGES" 
		elseif p_tipo = "C" then
			resultado = resultado&"CIC: 185.367.863-53"	
		elseif p_tipo = "I" then					
			resultado = resultado&"IDENTIDADE: 07926163-2 - IFP"
		end if	
	elseif p_num_testemunha = 2 then
		if p_tipo = "N" then	
			resultado = "Nome: SOR&Aacute;IA REZENDE NUNES"
		elseif p_tipo = "C" then			
			resultado = "CIC: 754.514.097-49<BR>"
		elseif p_tipo = "I" then				
			resultado = "IDENTIDADE: 05730026-1 - IFP"	
		end if			
	elseif p_num_testemunha = 3 then
		if p_tipo = "N" then	
			resultado = "Nome: ZILDA COTIAS PLOMBON"	
		elseif p_tipo = "C" then			
			resultado = "CIC: 389.446.777-00"
		elseif p_tipo = "I" then				
			resultado = "IDENTIDADE: 1933310-3 - IFP"	
		end if							
	end if

buscaTestemunha = resultado
end function

function geraIdentificacaoEscola(p_tp_contrato_adendo, p_tp_retorno)

	if p_tp_contrato_adendo = "ADENDO_1A" or p_tp_contrato_adendo = "ADENDO_8" or p_tp_contrato_adendo = "ADENDOESC2" then
		endEscolaBr = "<br />Rua Visconde de Ouro Preto, 51<br />Botafogo - RJ"
		endEscola	 = ", Rua Visconde de Ouro Preto, 51"		
	elseif p_tp_contrato_adendo = "ADENDO_2A" or p_tp_contrato_adendo = "ADENDO_2B" or p_tp_contrato_adendo = "ADENDO_2A2B" or p_tp_contrato_adendo = "ADENDOBL3" then
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

	if session("ano_letivo")>=2017 then
		if p_tp_contrato_adendo = "CONTRATO_G1" or  p_tp_contrato_adendo = "ADENDO_G1" or p_tp_contrato_adendo = "CONTRATOG1" or p_tp_contrato_adendo = "CONTRATOG1" or p_tp_contrato_adendo = "ADENDOGAVA" or p_tp_contrato_adendo = "ADENDOGAVB" or p_tp_contrato_adendo = "ADENDOGAVC" then
			escola	 = "Jardim Escola Stockler Ltda."	
		elseif p_tp_contrato_adendo = "CONTRATO_G2" or p_tp_contrato_adendo = "CONTRATO_G3" or p_tp_contrato_adendo = "CONTRATO_G4" or p_tp_contrato_adendo = "ADENDO_G2" or p_tp_contrato_adendo = "ADENDO_G3" then
			escola	 = "Colégio L. Stockler Ltda."	
		else
			escola	 = "Escola Dínamis Ltda."			
		end if	
	
	else
		if p_tp_contrato_adendo = "CONTRATO_1A" or p_tp_contrato_adendo = "CONTRATO_1A1" or p_tp_contrato_adendo = "CONTRATO_1B" or p_tp_contrato_adendo = "CONTRATO_1B1" or p_tp_contrato_adendo = "CONTRATO_8" or p_tp_contrato_adendo = "ADENDO_1A" or p_tp_contrato_adendo = "ADENDO_8" or p_tp_contrato_adendo = "CONTRATO_2A" or p_tp_contrato_adendo = "CONTRATO_2A1" or p_tp_contrato_adendo = "CONTRATO_2B" or p_tp_contrato_adendo = "CONTRATO_2B1" OR p_tp_contrato_adendo = "ADENDO_2A" or p_tp_contrato_adendo = "ADENDO_2B" or p_tp_contrato_adendo = "ADENDO_2A2B" or p_tp_contrato_adendo = "CONTRATO_3" or  p_tp_contrato_adendo = "ADENDO_3" or p_tp_contrato_adendo = "CONTRATO_5" or p_tp_contrato_adendo = "CONTRATO_5A" or p_tp_contrato_adendo = "CONTRATO_5A1" or p_tp_contrato_adendo = "CONTRATO_5B" or p_tp_contrato_adendo = "CONTRATO_5B1" or p_tp_contrato_adendo = "ADENDO_5" or p_tp_contrato_adendo = "ADENDOESC1"  or p_tp_contrato_adendo = "ADENDOESC2" or p_tp_contrato_adendo = "ADENDOESC3" or p_tp_contrato_adendo = "ADENDOBL1" or p_tp_contrato_adendo = "ADENDOBL2" or p_tp_contrato_adendo = "ADENDOBL3" or p_tp_contrato_adendo = "ADENDOCMP1" or p_tp_contrato_adendo = "ADENDOCMP2" or p_tp_contrato_adendo = "ADENDOCMP3" or p_tp_contrato_adendo = "CONTRATO_31" or p_tp_contrato_adendo = "CONTRATO_81" then
			escola	 = "Escola Dínamis Ltda."		
	'	elseif p_tp_contrato_adendo = "CONTRATO_2A" or p_tp_contrato_adendo = "CONTRATO_2B" OR p_tp_contrato_adendo = "ADENDO_2A" or p_tp_contrato_adendo = "ADENDO_2B" or p_tp_contrato_adendo = "ADENDO_2A2B" then
	'		escola	 = "Jardim Escola B.L. Ltda."		
	'	elseif p_tp_contrato_adendo = "CONTRATO_3" or  p_tp_contrato_adendo = "ADENDO_3" then
	'		escola	 = "Creche Experimental Dínamis Ltda."			
	'	elseif p_tp_contrato_adendo = "CONTRATO_5" or p_tp_contrato_adendo = "ADENDO_5" then
	'		escola	 = "Jardim Escola M.P. Ltda."		
		elseif p_tp_contrato_adendo = "CONTRATO_G1" or  p_tp_contrato_adendo = "ADENDO_G1" then
			escola	 = "Jardim Escola Stockler Ltda."	
		elseif p_tp_contrato_adendo = "CONTRATO_G2" or p_tp_contrato_adendo = "CONTRATO_G3" or p_tp_contrato_adendo = "CONTRATO_G4" or p_tp_contrato_adendo = "ADENDO_G2" or p_tp_contrato_adendo = "ADENDO_G3" then
			escola	 = "Colégio L. Stockler Ltda."					
		end if
	end if	
	geraNomeEscola = escola
end function

function geraNomeEtapa(p_tp_contrato_adendo, p_curso)

	if p_tp_contrato_adendo = "ADENDO_1A" or p_tp_contrato_adendo = "ADENDO_8" or p_tp_contrato_adendo = "ADENDO_G2" or p_tp_contrato_adendo = "ADENDO_G3"  or (p_tp_contrato_adendo = "ADENDOESC1" AND p_curso = 1) or (p_tp_contrato_adendo = "ADENDOESC2" AND p_curso = 1)   then
		nomeEtapa	 = "Ensino Fundamental"		
	elseif p_tp_contrato_adendo = "ADENDO_2A" or p_tp_contrato_adendo = "ADENDO_2B" or p_tp_contrato_adendo = "ADENDO_3"  or p_tp_contrato_adendo = "ADENDO_5" or p_tp_contrato_adendo = "ADENDO_G1" or p_tp_contrato_adendo = "ADENDOESC1" or p_tp_contrato_adendo = "ADENDOBL1" or p_tp_contrato_adendo = "ADENDOBL2" or p_tp_contrato_adendo = "ADENDOBL3" or p_tp_contrato_adendo = "ADENDOCMP1" or p_tp_contrato_adendo = "ADENDOCMP2" or p_tp_contrato_adendo = "ADENDOCMP3" or p_tp_contrato_adendo = "ADENDOCR1" or p_tp_contrato_adendo = "ADENDOCR2" or p_tp_contrato_adendo = "ADENDOBM1" or p_tp_contrato_adendo = "ADENDOBM2" or p_tp_contrato_adendo = "ADENDOBM3" then
		nomeEtapa	 = "Educação Infantil"
	elseif p_tp_contrato_adendo = "ADENDOESC2" or p_tp_contrato_adendo = "ADENDOESC3" then
		nomeEtapa	 = "Ensino Médio"		
	end if
	geraNomeEtapa = nomeEtapa
end function

function geraPreposicaoEtapa(p_tp_contrato_adendo, p_curso)

if p_tp_contrato_adendo = "ADENDO_1A" or p_tp_contrato_adendo = "ADENDO_8" or p_tp_contrato_adendo = "ADENDO_G2" or p_tp_contrato_adendo = "ADENDO_G3"  or (p_tp_contrato_adendo = "ADENDOESC1" AND p_curso = 1) or p_tp_contrato_adendo = "ADENDOESC2" or p_tp_contrato_adendo = "ADENDOESC3" then
		prepEtapa	 = "do"		
	elseif p_tp_contrato_adendo = "ADENDO_2A" or p_tp_contrato_adendo = "ADENDO_2B" or p_tp_contrato_adendo = "ADENDO_3"  or p_tp_contrato_adendo = "ADENDO_5" or p_tp_contrato_adendo = "ADENDO_G1" or p_tp_contrato_adendo = "ADENDOESC1" or p_tp_contrato_adendo = "ADENDOBL1" or p_tp_contrato_adendo = "ADENDOBL2" or p_tp_contrato_adendo = "ADENDOBL3" or p_tp_contrato_adendo = "ADENDOCMP1" or p_tp_contrato_adendo = "ADENDOCMP2" or p_tp_contrato_adendo = "ADENDOCMP3" or p_tp_contrato_adendo = "ADENDOCR1" or p_tp_contrato_adendo = "ADENDOCR2" or p_tp_contrato_adendo = "ADENDOBM1" or p_tp_contrato_adendo = "ADENDOBM2" or p_tp_contrato_adendo = "ADENDOBM3" then
		prepEtapa	 = "da"											
	end if
	geraPreposicaoEtapa = prepEtapa
end function

%>