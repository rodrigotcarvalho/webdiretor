<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/banner.asp"-->
<%
Function cabecalho (nivel)
Session.LCID = 1046 
nome = session("nome") 
acesso = session("acesso")
co_usr = session("co_user")
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
grupo=session("grupo")
escola=session("escola")
chave=session("chave")
		
this_file = Request.ServerVariables("SCRIPT_NAME")
arPath = Split(this_Path, "/")




if nome = "" or acesso = "" or co_usr = "" or permissao = "" or ano_letivo = "" or chave = "" or isnull(chave) then
if nivel=0 then
response.Redirect("default.asp?opt=00")
elseif nivel=1 then
response.Redirect("../default.asp?opt=00")
elseif nivel=2 then
response.Redirect("../../default.asp?opt=00")
elseif nivel=3 then
response.Redirect("../../../default.asp?opt=00")
elseif nivel=4 then
response.Redirect("../../../../default.asp?opt=00")
end if
else
session("escola")=escola
session("nome") = nome
session("acesso") = acesso
session("co_usr") = co_usr
session("tp") = tp
session("ano_letivo") = ano_letivo
session("permissao") = permissao
session("sistema_local")=sistema_local
session("chave")=chave
session("grupo") = grupo
end if
call banner(nivel,this_file,sistema_local,nome,permissao,ano_letivo)
end function




Function navegacao (Conexao,chave, nivel)
session("chave")=chave

Select case nivel

case 0
origem ="Voc&ecirc; est&aacute; em Web Diretor"
case 1
chavearray=split(chave,"-")
sistema=chavearray(0)
		Set RSc1 = Server.CreateObject("ADODB.Recordset")
		SQLc1 = "SELECT * FROM TB_Sistema where CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc1.Open SQLc1, Conexao

sistema_nome=RSc1("TX_Descricao")
link_sistema=RSc1("CO_Pasta")

origem = "Voc&ecirc; est&aacute; em <a href='../inicio.asp' class='caminho' target='_self'>Web Diretor</a> > "&sistema_nome
case 2

chavearray=split(chave,"-")
sistema=chavearray(0)
modulo=chavearray(1)



		Set RSc1 = Server.CreateObject("ADODB.Recordset")
		SQLc1 = "SELECT * FROM TB_Sistema where CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc1.Open SQLc1, Conexao
		
		sistema_nome=RSc1("TX_Descricao")
		link_sistema=RSc1("CO_Pasta")



		Set RSc2 = Server.CreateObject("ADODB.Recordset")
		SQLc2 = "SELECT * FROM TB_Modulo where CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc2.Open SQLc2, Conexao

		modulo_nome=RSc2("TX_Descricao")
		link_modulo=RSc2("CO_Pasta")
	
	
origem = "Voc&ecirc; est&aacute; em <a href='../../inicio.asp' class='caminho' target='_self'>Web Diretor</a> > <a href='../../"&link_sistema&"/index.asp?nvg="&sistema&"' class='caminho' target='_self'>"&sistema_nome&"</a> > <a href='../"&link_modulo&"/index.asp?nvg="&chave&"' class='caminho' target='_self'>"&modulo_nome&"</a>"
		
case 3
chavearray=split(chave,"-")
sistema=chavearray(0)
modulo=chavearray(1)
setor=chavearray(2)
		Set RSc1 = Server.CreateObject("ADODB.Recordset")
		SQLc1 = "SELECT * FROM TB_Sistema where CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc1.Open SQLc1, Conexao
		
		sistema_nome=RSc1("TX_Descricao")
		link_sistema=RSc1("CO_Pasta")

		Set RSc2 = Server.CreateObject("ADODB.Recordset")
		SQLc2 = "SELECT * FROM TB_Modulo where CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc2.Open SQLc2, Conexao

		modulo_nome=RSc2("TX_Descricao")
		link_modulo=RSc2("CO_Pasta")
		
		Set RSc3 = Server.CreateObject("ADODB.Recordset")
		SQLc3 = "SELECT * FROM TB_Setor where CO_Setor = '"&setor&"' and CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc3.Open SQLc3, Conexao

		setor_nome=RSc3("TX_Descricao")
		link_setor=RSc3("CO_Pasta")

origem = "Voc&ecirc; est&aacute; em <a href='../../../inicio.asp' class='caminho' target='_self'>Web Diretor</a> > <a href='../../../"&link_sistema&"/index.asp?nvg="&sistema&"' class='caminho' target='_self'>"&sistema_nome&"</a> > <a href='../../"&link_modulo&"/index.asp?nvg="&sistema&"-"&modulo&"' class='caminho' target='_self'>"&modulo_nome&"</a> > <a href='../"&link_setor&"/index.asp?nvg="&chave&"' class='caminho' target='_self'>"&setor_nome&"</a>"

case 4
chavearray=split(chave,"-")
sistema=chavearray(0)
modulo=chavearray(1)
setor=chavearray(2)
funcao=chavearray(3)

grupo=session("grupo")
negado=request.querystring("neg")
ano_letivo = session("ano_letivo") 


		Set RSal = Server.CreateObject("ADODB.Recordset")
		SQLal = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&ano_letivo&"'"
		RSal.Open SQLal, Conexao
		
		sit_an=RSal("ST_Ano_Letivo")
		
		Set RSac = Server.CreateObject("ADODB.Recordset")
		SQLac = "SELECT * FROM TB_Autoriz_Grupo_Funcao where CO_Funcao = '"&funcao&"' and CO_Setor = '"&setor&"' and CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' and CO_Grupo= '"&grupo&"'"
		RSac.Open SQLac, Conexao
		

		if RSac.eof then
			autoriza=1
		else
			autoriza=RSac("TP_Acesso")
		end if	
		'response.Write(autoriza)
		autoriza=autoriza*1
		if autoriza=0 and negado<>1 then
		nvg=sistema&"-"&modulo&"-"&setor&"-"&funcao
		response.Redirect("../../../../inc/negado.asp?nvg="&nvg&"&neg=1")
		elseif autoriza=1 then
		session("trava")="s"
		elseif autoriza=5  AND sit_an="L"then
		session("trava")="n"
		else
		session("trava")="s"
		end if
		
		session("autoriza")=autoriza

		Set RSc1 = Server.CreateObject("ADODB.Recordset")
		SQLc1 = "SELECT * FROM TB_Sistema where CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc1.Open SQLc1, Conexao
		
		sistema_nome=RSc1("TX_Descricao")
		link_sistema=RSc1("CO_Pasta")

		Set RSc2 = Server.CreateObject("ADODB.Recordset")
		SQLc2 = "SELECT * FROM TB_Modulo where CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc2.Open SQLc2, Conexao

		modulo_nome=RSc2("TX_Descricao")
		link_modulo=RSc2("CO_Pasta")
		
		Set RSc3 = Server.CreateObject("ADODB.Recordset")
		SQLc3 = "SELECT * FROM TB_Setor where CO_Setor = '"&setor&"' and CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc3.Open SQLc3, Conexao

		setor_nome=RSc3("TX_Descricao")
		link_setor=RSc3("CO_Pasta")
		
		Set RSc4 = Server.CreateObject("ADODB.Recordset")
		SQLc4 = "SELECT * FROM TB_Funcao where CO_Funcao = '"&funcao&"' and CO_Setor = '"&setor&"' and CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc4.Open SQLc4, Conexao

		funcao_nome=RSc4("TX_Descricao")
		link_funcao=RSc4("CO_Pasta")

if negado="1" then
origem = "Voc&ecirc; est&aacute; em <a href='../inicio.asp' class='caminho' target='_self'>Web Diretor</a> > <a href='../"&link_sistema&"/index.asp?nvg="&sistema&"' class='caminho' target='_self'>"&sistema_nome&"</a> > <a href='../"&link_sistema&"/"&link_modulo&"/index.asp?nvg="&sistema&"-"&modulo&"'class='caminho' target='_self'>"&modulo_nome&"</a> > <a href='../"&link_sistema&"/"&link_modulo&"/"&link_setor&"/index.asp?nvg="&sistema&"-"&modulo&"-"&setor&"' class='caminho' target='_self'>"&setor_nome&"</a> > <a href='../"&link_sistema&"/"&link_modulo&"/"&link_setor&"/"&link_funcao&"/index.asp?nvg="&chave&"' class='caminho' target='_self'>"&funcao_nome&"</a>"

else
origem = "Voc&ecirc; est&aacute; em <a href='../../../../inicio.asp' class='caminho' target='_self'>Web Diretor</a> > <a href='../../../../"&link_sistema&"/index.asp?nvg="&sistema&"' class='caminho' target='_self'>"&sistema_nome&"</a> > <a href='../../../"&link_modulo&"/index.asp?nvg="&sistema&"-"&modulo&"'class='caminho' target='_self'>"&modulo_nome&"</a> > <a href='../../"&link_setor&"/index.asp?nvg="&sistema&"-"&modulo&"-"&setor&"' class='caminho' target='_self'>"&setor_nome&"</a> > <a href='../"&link_funcao&"/index.asp?nvg="&chave&"' class='caminho' target='_self'>"&funcao_nome&"</a>"
end if
		
end select

Session("caminho")=origem
'session("chave")=chave
chave=session("chave")
end function

FUNCTION linkFuncao(Conexao,sistema,modulo,setor,funcao,nivel)


		Set RSc1 = Server.CreateObject("ADODB.Recordset")
		SQLc1 = "SELECT * FROM TB_Sistema where CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc1.Open SQLc1, Conexao
		
		sistema_nome=RSc1("TX_Descricao")
		link_sistema=RSc1("CO_Pasta")

		Set RSc2 = Server.CreateObject("ADODB.Recordset")
		SQLc2 = "SELECT * FROM TB_Modulo where CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc2.Open SQLc2, Conexao

		modulo_nome=RSc2("TX_Descricao")
		link_modulo=RSc2("CO_Pasta")
		
		Set RSc3 = Server.CreateObject("ADODB.Recordset")
		SQLc3 = "SELECT * FROM TB_Setor where CO_Setor = '"&setor&"' and CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc3.Open SQLc3, Conexao

		setor_nome=RSc3("TX_Descricao")
		link_setor=RSc3("CO_Pasta")
		
		Set RSc4 = Server.CreateObject("ADODB.Recordset")
		SQLc4 = "SELECT * FROM TB_Funcao where CO_Funcao = '"&funcao&"' and CO_Setor = '"&setor&"' and CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc4.Open SQLc4, Conexao

		funcao_nome=RSc4("TX_Descricao")
		link_funcao=RSc4("CO_Pasta")


link_funcao="../../../../"&link_sistema&"/"&link_modulo&"/"&link_setor&"/"&link_funcao
session("link_funcao")=link_funcao
end function
'///////////////////////////////////////////////    MENSAGENS     //////////////////////////////////////////////////////////////////////////////


FUNCTION mensagens(nivel,msg,tab,dados)
escola=Session("escola")

SELECT CASE msg
'Mensagens Gerais de 0 a 49
case 0
wrt = "Escolha uma das op��es abaixo"

case 1
wrt = "Selecione uma unidade, um curso, uma etapa e uma turma. "

case 2
wrt = "Selecione uma etapa e uma turma."

case 3
wrt = "Selecione uma etapa, uma turma, um per�odo e uma avalia��o."

case 4
wrt = "Para consultar � necess�rio selecionar uma etapa!"

case 5
wrt = "Esta fun��o permite voc� fazer contato com a equipe t�cnica que realiza a manuten��o do sistema Web Diretor. Utilize sempre que poss�vel este canal para nos transmitir alguma informa��o relevante sobre o funcionamento desse produto. Obrigado pela sua aten��o!"

case 6
wrt = "Mensagem enviada."

case 7
wrt = "Escolha um novo usu�rio."

case 8
wrt = "Escolha uma nova senha."

case 9
wrt = "Usu�rio alterado com sucesso."

case 10
wrt = "Senha alterada com sucesso."

case 11
wrt = "Selecione uma disciplina e um per�odo."

case 12
wrt = "E-mail alterado com sucesso."

case 13
wrt = "Usu�rio j� existe!"

case 14
wrt = "Digite seu novo endere�o de correio eletr�nico"

case 15
wrt = "Endere�o de correio eletr�nico j� existe!"

case 16
wrt = "Selecione uma etapa, uma turma e um per�odo."

case 17
wrt = "Selecione uma etapa e um per�odo."

case 18
wrt = "Gr�fico comparativo."

case 19
wrt = "Selecione uma etapa, uma disciplina e um per�odo."

case 20
wrt = "Selecione uma etapa"

case 21
wrt = "Clique na op��o desejada"

case 22
wrt = "Confirma a reinicializa��o da senha do usu�rio abaixo?"

case 23
wrt = "Senha reinicializada com sucesso"

case 24
wrt = "Usu�rio "&situacao&" com sucesso"

case 25
wrt = "Selecione uma unidade, um curso, uma etapa e uma m�dia."

'Web Fam�lia de 50 a 99



case 51
wrt = "Selecione o tipo de documento e os arquivos que deseja disponibilizar para upload"

case 52
wrt = "Arquivo(s) "&Session("arquivos") &" enviado(s) com sucesso! Total de Bytes enviados:"&Session("upl_total")

case 53
wrt = "Selecione pelo menos um arquivo"

case 54
wrt = "Preencha os dados abaixo para associar um documento"

case 55
wrt = "Associa��o realizada com Sucesso"

case 56
wrt = "Preencha os dados abaixo para incluir uma not�cia"

case 57
wrt = "Not�cia inclu�da com sucesso"

case 58
wrt = "Confirma a exclus�o do(s) documento(s) abaixo?"

case 59
wrt = "Documento(s) exclu�do(s) com sucesso"

case 60
wrt = "Confirma a exclus�o do(s) arquivo(s) abaixo?"

case 61
wrt = "Arquivo(s) exclu�do(s) com sucesso"

case 62
wrt = "Selecione o tipo de documento"

case 63
wrt = "Confirma a exclus�o da(s) not�cia(s) abaixo?"

case 64
wrt = "Not�cia(s) exclu�da(s) com sucesso"

case 65
wrt = "Confirma a exclus�o do(s) evento(s) abaixo?"

case 66
wrt = "Evento(s) exclu�do(s) com sucesso"

case 67
wrt = "Preencha os dados abaixo para incluir um evento"

case 68
wrt = "Evento inclu�do com sucesso"

case 69
wrt = "Para consultar os dados do usu�rio digite o c�digo ou Nome e clique no bot&atilde;o Procurar."

case 70
wrt = "Escolha um usu�rio para consultar o cadastro."

case 71
wrt = "Verifique os dados do usu�rio."

case 72
wrt = "N�o foi encontrado nenhum usu�rio com este c�digo."

' erro na busca por nome
case 73
wrt = "N�o foi encontrado nenhum usu�rio com este nome."

case 74
wrt = "Pasta criada com sucesso!"

case 75
wrt = "Pasta modificada com sucesso!"

case 76
wrt = "Preencha os dados abaixo para incluir uma mensagem"

case 77
wrt = "Mensagem inclu�da com sucesso!"

case 78
wrt = "Mensagem exclu�da com sucesso!"

case 79
wrt = "Confirma a exclus�o da(s) mensagem(ns) abaixo?"


'alunos de 300 a 399
case 300
wrt = "Para consultar os dados do Aluno digite a Matr�cula ou Nome e clique no bot&atilde;o Procurar."

' listagem de alunos

case 301
wrt = "Escolha um Aluno para consultar o cadastro."

case 302
wrt = "Verifique os dados do Aluno."

case 303
wrt = "N�o foi encontrado nenhum Aluno com este c�digo."

' erro na busca por nome
case 304
wrt = "N�o foi encontrado nenhum Aluno com este nome."

case 305
wrt = "Lista de alunos associados a turma abaixo."

case 306
wrt = "Verifique os dados dos familiares."

case 307
wrt = "Selecione uma unidade e um m�s."

case 308
wrt = "Comparar Turma por M�dia Geral."

case 309
wrt = "Verifique os dados do Aluno e escolha uma disciplina e um per�odo."

case 310
wrt = "Escolha os crit�rios para pesquisar as ocorr�ncias do aluno e clique no bot�o prosseguir."

case 311
wrt = "Confirma a exclus�o dessa(s) disciplina(s)."


case 312
wrt = "Ocorr�ncia inclu�da com sucesso!"

case 313
wrt = "Ocorr�ncia alterada com sucesso!"

case 314
wrt = "Ocorr�ncia exclu�da com sucesso!"

case 315
wrt = "Preencha os dados abaixo e clique no bot�o Confirmar para Incluir uma nova ocorr�ncia."

case 316
wrt = "Preencha os dados abaixo e clique no bot�o Confirmar para atualizar esta ocorr�ncia."

case 317
wrt = "Selecione uma situa��o para o aluno e escreva o motivo da inativa��o."

case 318
wrt = "J� existe o mesmo tipo de ocorr�ncia registrada para a mesma hora."

case 319
wrt = "N&atilde;o foi poss&iacute;vel gravar a ocorr&ecirc;ncia para nenhuma matr&iacute;cula, pois j&aacute; existe o mesmo tipo de ocorr&ecirc;ncia registrada para a mesma hora."

case 320
wrt = "Para as matr&iacute;culas "&cod&" n&atilde;o foi poss&iacute;vel gravar a ocorr&ecirc;ncia, pois j&aacute; existe o mesmo tipo de ocorr&ecirc;ncia registrada para a mesma hora. Para as demais a ocorr�ncia foi gravada com sucesso."

case 321
wrt = "Escolha na lista abaixo as matr&iacute;culas as quais deseja gravar a ocorr&ecirc;ncia."


'web secretaria 400 a 450
case 400
wrt = "Para consultar os dados do Aluno digite a Matr�cula ou Nome e clique no bot&atilde;o Procurar. Caso o aluno n�o esteja cadastrado no sistema clique <a href='../../../cad/man/aal/cadastra.asp?nvg=WS-CA-MA-AAL' class='avisos'>aqui</a>."

case 401
wrt = "Matr�cula efetuada com sucesso!"

case 402
wrt = "Preencha os campos abaixo."

case 403
wrt = "Aluno j� matriculado para este ano letivo. Matr�culas para o pr�ximo Ano Letivo est�o fechadas!"

case 404
wrt = "Para alterar os dados do Aluno digite a Matr�cula ou Nome e clique no bot&atilde;o Procurar. Caso o aluno n�o esteja cadastrado no sistema clique <a href='../../../cad/man/aal/cadastra.asp?nvg=WS-CA-MA-AAL' class='avisos'>aqui</a>."

case 405
dados=dados

separa=split(dados,"#sep#")
ordem_familiares=separa(0)
qtd_tipo_familiares=separa(1)
cod_familiar=separa(2)
cod_vinculado=separa(3)
cod_aluno=separa(4)
wrt1 ="<input name='ordem' type='hidden' value='"&ordem_familiares&"'>"
'wrt2 ="<input name='cod_prim' type='hidden' value='"&cod_familiar_prim&"'>"
wrt2 ="<input name='qtd' type='hidden' value='"&qtd_tipo_familiares&"'>"
wrt3 ="<input name='foco' type='hidden' value='"&cod_familiar&"'>"
wrt4 ="<input name='cod_vinculado' type='hidden' value='"&cod_vinculado&"'>"
wrt5 ="<input name='cod_al' type='hidden' value='"&cod_aluno&"'>"
wrt6 =Server.URLEncode("Confirma a exclus�o desse familiar?")

wrt = wrt1&wrt2&wrt3&wrt4&wrt5&wrt6&"<br><input type='button' name='Submit2' value='Sim' onClick='ExcluiFamiliares(ordem.value,qtd.value,foco.value,cod_al.value)' class='botao_prosseguir_sim' >&nbsp;&nbsp;&nbsp;<input type='button' name='Submit2' value='"&Server.URLEncode("N�o")&"' onClick='recuperarFamiliares(ordem.value,qtd.value,foco.value,cod_vinculado.value,cod_al.value)' class='botao_prosseguir_nao' >"

case 406
dados=dados

separa=split(dados,"#sep#")
ordem_familiares=separa(0)
qtd_tipo_familiares=separa(1)
cod_familiar=separa(2)
cod_vinculado=separa(3)
cod_aluno=separa(4)
'cod_nome = Split(ordem_familiares, "!!")
'cod_familiar_prim=cod_nome(0)
wrt1 ="<input name='ordem' type='hidden' value='"&ordem_familiares&"'>"
'wrt2 ="<input name='cod_prim' type='hidden' value='"&cod_familiar_prim&"'>"
wrt2 ="<input name='qtd' type='hidden' value='"&qtd_tipo_familiares&"'>"
wrt3 ="<input name='foco' type='hidden' value='"&cod_familiar&"'>"
wrt4 ="<input name='cod_vinculado' type='hidden' value='"&cod_vinculado&"'>"
wrt5 ="<input name='cod_al' type='hidden' value='"&cod_aluno&"'>"
wrt6 =Server.URLEncode("O CPF Digitado possui dados cadastrados. Deseja aproveitar esses dados?")

wrt = wrt1&wrt2&wrt3&wrt4&wrt5&wrt6&"<br><input type='button' name='Submit2' value='Sim' onClick='recuperarFamiliares(ordem.value,qtd.value,foco.value,cod_vinculado.value,cod_al.value)' class='botao_prosseguir_sim' >&nbsp;&nbsp;&nbsp;<input type='button' name='Submit2' value='"&Server.URLEncode("N�o")&"' onClick='ExcluiFamiliares(ordem.value,qtd.value,foco.value,cod_al.value)' class='botao_prosseguir_nao' >"

case 407
wrt = "Deve ser selecionado um respons�vel financeiro para o aluno!"

case 408
wrt = "Deve ser selecionado um respons�vel pedag�gico para o aluno!"

case 409
wrt = "� obrigat�rio o preenchimento dos campos: Nome, Telefones de Contato e Endere�o residencial para o respons�vel financeiro!"

case 410
wrt = "� obrigat�rio o preenchimento dos campos: Nome, Telefones de Contato e Endere�o residencial para o respons�vel pedag�gico!"

case 411
wrt = "Ao se confirmar o cadastro desse aluno, esse n�mero de matr�cula n�o poder� mais ser utilizado!"

case 412
wrt = "Cadastro efetuado com sucesso! Inclua todos os dados necess�rios."

case 413
wrt = "Selecione uma nova combina��o de Unidade, Curso, Etapa, Turma e N�mero de chamada para o aluno."

case 414
wrt = "Selecione um m�todo para enturmar os alunos em situa��o de pr�-matr�cula."

case 415
wrt = "N�o existem alunos em situa��o de pr�-matr�cula."

case 416
wrt = "Somente � poss�vel remanejar alunos com situa��o igual a 'Cursando'."

case 417
wrt = "Confirma a exclus�o desse(s) hist�rico(s)."

case 418
wrt = "Hist�rico inclu�do com sucesso!"

case 419
wrt = "Hist�rico alterado com sucesso!"

case 420
wrt = "Hist�rico(s) exclu�do(s) com sucesso!"


'professores de 600 a 899

case 600
wrt =  "Os Professores em vermelho est�o inativos. A mensagem 'n�o cadastrado' indica que n�o existe professor associado �quela disciplina naquela turma"
wrt = wrt &"<br>A mensagem 'nome em branco' indica que o nome do professor n�o est� registrado no cadastro. Para bloquear a planilha clique na letra 'N' do per�odo escolhido"

case 601
wrt = "Confirma o " 
if opt="blq" then
wrt= wrt &"BLOQUEIO"
else
wrt= wrt &"DESBLOQUEIO"
end if
wrt= wrt &" das notas do trimestre "&periodo&" de "&no_materia&", Unidade:"&no_unidade&" - "&no_etapa&" do "&no_curso&" Turma "&turma&""

case 602
if orig=01 then
act= "Planilha bloqueada"
elseif orig=02 then
act= "Planilha desbloqueada"
elseif orig=03 then
act= "Planilhas bloqueadas"
elseif orig=04 then
act= "Planilhas desbloqueadas"
end if

wrt = act&" com sucesso!"

case 603
wrt = "Avalia��es n�o lan�adas"

case 604
wrt = "Para consultar a Grade de aulas digite o C&oacute;digo ou Nome de um Professor e clique no bot&atilde;o Procurar."
wrt = wrt &"<br>Se preferir obter uma lista completa de TODOS os professores clique <a href='index.asp?opt=listall&nvg="&nvg&"' class='avisos'>aqui</a>"

case 605
wrt = "N�o foi encontrado nenhum professor com este c�digo."

case 606
wrt = "Escolha um professor para consultar a Grade de Aulas. Os Professores em vermelho est�o inativos."

case 607
wrt = "Para atualizar os dados do Professor digite o C&oacute;digo ou Nome e clique no bot&atilde;o Procurar."
wrt = wrt &"Se preferir adicionar um NOVO professor clique <a href='altera.asp?ori=02&nvg="&nvg&"' class='avisos'>aqui</a>."
wrt = wrt &"<BR>Se preferir obter uma lista completa de TODOS os professores clique <a href='index.asp?opt=listall&nvg="&nvg&"' class='avisos'>aqui</a>"

case 608
wrt = "Confirme o professor para consultar a Grade de Aulas."

case 609
wrt = "O per�odo relacionado pela letra 'S' indica que a planilha est� Bloqueada e 'N' que est� Desbloqueada."

case 610
wrt = "N�o foi encontrado nenhum professor com este c�digo."

case 611
wrt = "N�o foi encontrado nenhum professor com este nome."

case 612
wrt = "Escolha um professor para atualizar o cadastro. Os Professores em vermelho est�o inativos."

case 613
wrt = "Confirme se � o professor correto para atualizar o cadastro."

case 614
wrt = "Preencha cuidadosamente os dados do Professor e click no bot&atilde;o CONFIRMAR para atualizar o cadastro"

case 615
wrt = "Professor c�digo "&cod_cons&" e usu�rio "&tx_login&" inclu�do com sucesso!"

case 616
wrt = "Dados do Professor c�digo "&cod_cons&" alterados com sucesso!"

case 617
wrt = "Selecione a Data e a Hora as quais voc� deseja iniciar o monitoramento de notas e clique em iniciar."

case 618
mes_mnl=mes_mnl*1
min_mnl=min_mnl*1
			  if mes_mnl< 10 then
			  mes_wrt="0"&mes_mnl
			  else
			  mes_wrt=mes_mnl					  
			  end if 
					  
			  if min_mnl< 10 then
			  min_wrt="0"&min_mnl
			  else
			  min_wrt=min_mnl					  
			  end if 
wrt = "Inicio da monitora��o a partir do dia "&dia_mnl&"/"&mes_wrt&"/"&ano_mnl&" as "&hora_mnl&":"&min_wrt&" Dados atualizados a cada minuto."

case 619
wrt = "N�o foram encontradas turmas cadastradas para voc�. Entre em contato com o seu coordenador."


case 620
if errou="pv1" or errou="pv2" or errou="pv3" or errou="pv4" or errou="pv5" or errou="pv6" Then
wrt = "Valor inv�lido para o campo  "&errado
elseif errou="sp" Then
wrt = "Soma dos Pesos maior que 10"
elseif errou="pt" Then
wrt = "Um dos pesos tem valor inv�lido"
elseif errou="pr1pr2" Then
wrt = "Soma das Pr's maior que 10"
else
wrt = "Valor inv�lido para o campo  "&errado&"  do n�mero de chamada <b>"&errante&"</b>"
end if

' erro na busca por c�digo
case 621
wrt = "Voc� est� " 
if opt="cln" then
wrt= wrt &"comunicando"
else
wrt= wrt &"lan�ando"
end if


		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set RSpr = Server.CreateObject("ADODB.Recordset")
		SQLpr = "SELECT * FROM TB_Periodo where NU_Periodo = "&periodo
		RSpr.Open SQLpr, CON0

no_periodo=RSpr("NO_Periodo")

wrt= wrt &" notas de "&no_periodo&" de "&no_materia&", Unidade:"&no_unidades&" - "&no_serie&" do "&no_grau&" Turma "&turma&""

case 622
wrt = "Notas lan�adas com sucesso."

case 623
wrt = "Comunicado efetuado!"

case 624
wrt = "Estas notas j&aacute; foram lan&ccedil;adas.Para alter&aacute;-las pe&ccedil;a autoriza&ccedil;&atilde;o ao coordenador"

case 625
wrt = "Escolha um Coordenador para consultar os Professores sob sua coordena��o."

case 626
wrt = "Os Professores em vermelho est�o inativos. A mensagem 'n�o cadastrado'indica que n�o existe professor associado �quela disciplina naquela turma"
wrt = wrt &"<br>A mensagem 'nome em branco' indica que o nome do professor n�o est� registrado no cadastro"

case 627
wrt = "Para excluir, selecione uma ou mais disciplinas e clique em excluir.<br>Para incluir uma nova disciplina na Grade de Aulas, selecione uma unidade e um curso."

case 628
wrt = "Disciplina inclu�da com sucesso"

case 629
wrt = "Disciplina exclu�da com sucesso"

case 630
wrt = "N�o � poss�vel marcar uma disciplina na Grade de Aulas e selecionar uma unidade e um curso ao mesmo tempo.<br>Por favor selecione somente disciplina(s) para excluir ou selecione uma unidade para incluir uma nova disciplina na Grade de Aulas"

case 631
wrt = "Selecione uma disciplina, um modelo e um coordenador."

case 632
wrt = "Para atualizar � necess�rio selecionar uma disciplina,um modelo e um coordenador"

case 633
wrt = "Verifique os dados preenchidos e clique no bot�o Confirmar para continuar a inclus�o ou no bot�o Alterar para voltar e modificar algum dado."


case 634
wrt = "Verifique as disciplinas selecionadas e clique no bot�o confirmar para Excluir ou no bot�o Cancelar para voltar e modificar algum dado."

case 635
wrt = "Professores que n�o comunicaram."

case 636
wrt = "Para imprimir clique <a class='avisos' href='#' onClick=MM_openBrWindow('imprime.asp?or=01&obr="&obr&"&p=p','','status=yes,menubar=yes,scrollbars=yes,resizable=yes,width=1030,height=500,top=50,left=50')>aqui</a>."

case 637
wrt = "Escolha um professor e um per�odo."

case 638
wrt =  "Os Professores em vermelho est�o inativos. A mensagem 'n�o cadastrado' indica que n�o existe professor associado �quela disciplina naquela turma"
wrt = wrt &"<br>A mensagem 'nome em branco' indica que o nome do professor n�o est� registrado no cadastro. Clique no nome da disciplina para ver o mapa de resultado."

case 639
wrt = "Arquivo "& fl &" enviado com sucesso."

case 640
wrt = "Aten��o! Estas notas j&aacute; foram lan&ccedil;adas pelo professor."

case 641
wrt = "Inclua as faltas no per�odo desejado"

case 642
wrt = "Faltas lan�adas com sucesso"

case 643
wrt = "Para atualizar os dados do Professor digite o C&oacute;digo ou Nome e clique no bot&atilde;o Procurar."
wrt = wrt &"<BR>Se preferir obter uma lista completa de TODOS os professores clique <a href='index.asp?opt=listall&nvg="&nvg&"' class='avisos'>aqui</a>"

case 644
wrt = "� necess�rio escolher pelo menos uma unidade"

case 645
wrt = "Imprimir <a class='avisos' href='#' onClick=MM_openBrWindow('imprime.asp?obr="&obr&"&p=p','','status=yes,menubar=yes,scrollbars=yes,resizable=yes,width=1030,height=500,top=50,left=50')>html</a> / <a class='avisos' href='../../../../relatorios/swd015.asp?obr="&obr&"'>pdf</a>."

case 646
wrt = "Para carregar as notas do simulado &eacute; necess&aacute;rio que o arquivo seja o modelo padr&atilde;o que pode ser baixado <a href=resultados.xls>aqui</a>."

case 647
wrt = "Arquivo "&Session("arquivo") &" enviado com sucesso! Total de Bytes enviados:"&Session("upl_total")

case 648
wrt = "Nenhum arquivo selecionado!"

case 649
wrt = "O arquivo "&Session("arquivo") &" n&atilde;o possui o nome correto!"

case 650
dados_erro=split(dados,"$!$")
wrt = "Matr&iacute;cula "&dados_erro(0)&" n&atilde;o cadastrada em TB_Matriculas para o Ano Letivo "&dados_erro(1)&"!. Favor verificar e reenviar a planilha!"

case 651
dados_erro=split(dados,"$!$")
wrt = "Erro na grava��o da matr&iacute;cula: "&dados_erro(0)&", disciplina: "&dados_erro(1)&", nota: "&dados_erro(2)&". Favor corrigir e reenviar a planilha!"


case 652
separa_dados=split(dados,"#$#")
separa=split(separa_dados(0),"#!#")
no_unidade=separa(0)
no_curso=separa(1)
no_etapa=separa(2)
data_grav=separa(3)
hora_grav=separa(4)
obr_mapa=separa_dados(1)
wrt = "Existem informa��es geradas em "&data_grav&" �s "&hora_grav&" para a Unidade: "&no_unidade&", Curso: "&no_curso&", Etapa: "&no_etapa&" e Turma: "&turma&".<BR>Deseja reprocessar essas informa��es? <a class='avisos' href='gera_base.asp?opt=rgnrt&obr="&obr_mapa&"' onclick=redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show')>sim</a> / <a class='avisos' href='#' onClick=redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show');MM_openBrWindow('mapa.asp?obr="&obr_mapa&"','')>n�o</a>"

case 653
wrt = "O arquivo est� sendo gerado!"

case 654
separa_dados=split(dados,"$$$")
tipo_busca=split(separa_dados(0),"$!$")
tipo=tipo_busca(0)
	if tipo="a" then
		separa=split(separa_dados(1),"#!#")
		no_aluno=separa(0)
		co_aluno=separa(1)
		data_grav=separa(2)
		hora_grav=separa(3)
		wrt =  "Existem informa��es geradas em "&data_grav&" �s "&hora_grav&" para o aluno "&no_aluno&", matr�cula "&co_aluno&"."
		javascript="MM_callJS('document.busca.busca1.focus()');"
	else
		separa=split(separa_dados(1),"#!#")
		no_unidade=separa(0)
		no_curso=separa(1)
		no_etapa=separa(2)
		data_grav=separa(3)
		hora_grav=separa(4)
		wrt =  "Existem informa��es geradas em "&data_grav&" �s "&hora_grav&" para a Unidade: "&no_unidade&", Curso: "&no_curso&", Etapa: "&no_etapa&" e Turma: "&turma&"."
		javascript=""
	end if
	obr_mapa=separa_dados(2)
	wrt = wrt&"<BR>Deseja reprocessar essas informa��es? <a class='avisos' href='gera_base.asp?opt=rgnrt&obr="&separa_dados(0)&"$$$"&obr_mapa&"' onclick=redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show')>sim</a> / <a class='avisos' href='../../../../relatorios/swd025.asp?obr="&separa_dados(0)&"$$$"&obr_mapa&"&ori=ebe' onclick=""redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show','divTabela','','show');AlternarMensagem('divMensagem2');mclosetime();"&javascript&""">n�o</a> "
	
case 655
separa_dados=split(dados,"$$$")
tipo_busca=separa_dados(0)
obr_mapa=separa_dados(2)
wrt = "O arquivo gerado com sucesso! Clique <a class='avisos' href='../../../../relatorios/swd025.asp?obr="&tipo_busca&"$$$"&obr_mapa&"&ori=ebe' onclick=""redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show','divTabela','','show');AlternarMensagem('divMensagem2');mclosetime();"">aqui</a> para fazer o download."

case 656
separa_dados=split(dados,"$$$")
tipo_busca=split(separa_dados(0),"$!$")
tipo=tipo_busca(0)
	if tipo="a" then
		separa=split(separa_dados(1),"#!#")
		no_aluno=separa(0)
		co_aluno=separa(1)
		data_grav=separa(2)
		hora_grav=separa(3)
		wrt =  "Existem informa��es geradas em "&data_grav&" �s "&hora_grav&" para o aluno "&no_aluno&", matr�cula "&co_aluno&"."
		javascript="MM_callJS('document.busca.busca1.focus()');"
	else
		separa=split(separa_dados(1),"#!#")
		no_unidade=separa(0)
		no_curso=separa(1)
		no_etapa=separa(2)
		data_grav=separa(3)
		hora_grav=separa(4)
		wrt =  "Existem informa��es geradas em "&data_grav&" �s "&hora_grav&" para a Unidade: "&no_unidade&", Curso: "&no_curso&", Etapa: "&no_etapa&" e Turma: "&turma&"."
		javascript=""
	end if
	obr_mapa=separa_dados(2)
	wrt = wrt&"<BR>Deseja reprocessar essas informa��es?  <a class='avisos' href='gera_base.asp?opt=rgnrt&obr="&separa_dados(0)&"$$$"&obr_mapa&"' onclick=redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show')>sim</a> / <a class='avisos' href='../../../../relatorios/swd048.asp?obr="&separa_dados(0)&"$$$"&obr_mapa&"&ori=efi' onclick=""redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show','divTabela','','show');AlternarMensagem('divMensagem2');mclosetime();"&javascript&""">n�o</a>"
	
case 657
separa_dados=split(dados,"$$$")
tipo_busca=separa_dados(0)
obr_mapa=separa_dados(2)
wrt = "O arquivo gerado com sucesso! Clique <a class='avisos' href='../../../../relatorios/swd048.asp?obr="&tipo_busca&"$$$"&obr_mapa&"&ori=efi' onclick=""redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show','divTabela','','show');AlternarMensagem('divMensagem2');mclosetime();"">aqui</a> para fazer o download."


case 658
separa_dados=split(dados,"#$#")
separa=split(separa_dados(0),"#!#")
no_unidade=separa(0)
no_curso=separa(1)
no_etapa=separa(2)
data_grav=separa(3)
hora_grav=separa(4)
obr_mapa=separa_dados(1)
wrt = "Existem informa��es geradas em "&data_grav&" �s "&hora_grav&" para a Unidade: "&no_unidade&", Curso: "&no_curso&", Etapa: "&no_etapa&" e Turma: "&turma&".<BR>Deseja reprocessar essas informa��es? <a class='avisos' href='gera_base.asp?opt=rgnrt&obr="&obr_mapa&"' onclick=redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show')>sim</a> / <a class='avisos' href='#' onClick=redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show');MM_openBrWindow('mapa.asp?obr="&obr_mapa&"','')>n�o</a>"

case 659
wrt = "O arquivo est� sendo gerado!"


case 660
separa_dados=split(dados,"#$#")
separa=split(separa_dados(0),"#!#")
no_unidade=separa(0)
no_curso=separa(1)
no_etapa=separa(2)
data_grav=separa(3)
hora_grav=separa(4)
obr_mapa=separa_dados(1)
wrt =  "Existem informa��es geradas em "&data_grav&" �s "&hora_grav&" para a Unidade: "&no_unidade&", Curso: "&no_curso&", Etapa: "&no_etapa&" e Turma: "&turma&".<BR>Deseja reprocessar essas informa��es? <a class='avisos' href='gera_base.asp?opt=rgnrt&obr="&obr_mapa&"' onclick=redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show')>sim</a> / <a class='avisos' href=gera_pdf.asp?obr="&obr_mapa&" onclick=""redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show','divTabela','','show');AlternarMensagem('divMensagem2');mclosetime();"">n�o</a>"

case 661
wrt = "O arquivo gerado com sucesso! Clique <a class='avisos' href=gera_pdf.asp?obr="&obr_mapa&" onclick=""redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show','divTabela','','show');AlternarMensagem('divMensagem2');mclosetime();"">aqui</a> para fazer o download."

case 662
separa_dados=split(dados,"$$$")
tipo_busca=separa_dados(0)
dados_funcao=split(separa_dados(2),"$!$")
unidade_fn = dados_funcao(0)
curso_fn = dados_funcao(1)
co_etapa_fn = dados_funcao(2)
turma_fn = dados_funcao(3)
periodo_fn = dados_funcao(4)
origem_fn = dados_funcao(6)

if origem_fn = "boletim" then
	origem_fn="Os Boletins foram gerados"
elseif origem_fn = "ficha" then
	origem_fn="As Fichas foram geradas"
end if	

if isnull(unidade_fn) or unidade_fn="" then
	unidade_msg="Todas as Unidades, "
else
	'unidade_msg_t= GeraNomes("U",unidade_fn,variavel2,variavel3,variavel4,variavel5,CON0,outro)
	call GeraNomes(materia,unidade_fn,curso_fn,co_etapa_fn,CON0)
	unidade_msg_t	=session("no_unidades")
	unidade_msg	= "Unidade "&unidade_msg_t&", "
end if
if isnull(curso_fn) or curso_fn="" then
else
	if curso_fn=999990 then
		curso_msg="todos os Cursos, "
	else
		'curso_msg_t= GeraNomes("C",curso_fn,variavel2,variavel3,variavel4,variavel5,CON0,outro)
		call GeraNomes(materia,unidade_fn,curso_fn,co_etapa_fn,CON0)
		curso_msg_t= session("no_grau")	
		curso_msg="Curso "&curso_msg_t&", "			
	end if	
end if	
if isnull(co_etapa_fn) or co_etapa_fn="" then
else
	if co_etapa_fn=999990 then
		co_etapa_msg="todas as Etapas, "
	else
		co_etapa_msg="Etapa "&co_etapa_fn&", "		
	end if	
end if	
if isnull(turma_fn) or turma_fn="" then
else
	if isnumeric(turma_fn) then
		if turma_fn=999990 then
			turma_msg="Todas as Turmas"
		else
			turma_msg="Turma "&turma_fn
		end if	
	else
		if turma_fn="999990" then
			turma_msg="Todas as Turmas"
		else
			turma_msg="Turma "&turma_fn
		end if	
	end if	
end if	
if isnull(periodo_fn) or periodo_fn="" then
else
	if periodo_fn=0 or periodo_fn=999990 then
		periodo_msg=" e todos os Per�odos!"
	else
'		divisao=tipo_divisao_ano(curso,co_etapa,tipo_dado)
'		periodo_msg_t = periodos(periodo_fn, divisao, "nome")
		periodo_msg_t = periodos(periodo_fn, "nome")		
		periodo_msg=" do "&periodo_msg_t&"!"
	end if	
end if	
wrt = origem_fn&" com sucesso para "&unidade_msg&curso_msg&co_etapa_msg&turma_msg&periodo_msg

case 663
wrt = "Avalia��es qualitativas lan�adas com sucesso!"

case 664
wrt = "Imprimir <a class='avisos' href='../../../../relatorios/swd015_maq.asp?obr="&obr&"'>pdf</a>."

case 665
wrt = "Imprimir <a class='avisos' href='../../../../relatorios/swd025_aq.asp?obr="&obr&"'>pdf</a>."

'Web Secretaria de 700 a 799
case 700
wrt = "O aluno solicitado n�o est� ATIVO no ano letivo de "&session("ano_letivo")&"!"

case 701
wrt = "N�o � poss�vel gerar boletim para alunos da Educa��o Infantil!"

case 702
wrt = "A op��o em destaque � a que foi selecionada. � poss�vel alterar o gr�fico selecionando uma das outras op��es."

case 703
wrt = "N�o existem alunos APROVADOS que atendam as condi��es de busca solicitadas."

case 704
wrt = "N�o existem alunos REPROVADOS que atendam as condi��es de busca solicitadas."

case 705
wrt = "Somente � poss�vel gerar esse relat�rio para alunos do 3� ano do Ensino M�dio."

case 706
wrt = "N�o existem alunos que atendam as condi��es de busca solicitadas."

case 707
separa_dados=split(dados,"$$$")
tipo_busca=split(separa_dados(0),"$!$")
	tipo=tipo_busca(0)
separa=split(separa_dados(1),"$!$")
	no_unidade=separa(0)
	no_curso=separa(1)
	no_etapa=separa(2)
	data_grav=separa(3)
	hora_grav=separa(4)
obr_mapa=separa_dados(2)
wrt = "O Ano Letivo est� encerrado e n�o existem dados gerados para a Unidade: "&no_unidade&", Curso: "&no_curso&", Etapa: "&no_etapa&" e Turma: "&turma&". Para que sejam gerados os dados � necess�ria a abertura do Ano Letivo."

case 708
wrt = "� poss�vel alterar o gr�fico selecionando uma das outras op��es."

case 709
wrt = "Email enviado com sucesso!"

case 710
wrt = "Selecione na lista os alunos os quais deseja retirar bonus"

case 711
wrt = "Bonus lan�ados com sucesso!"

case 712
wrt = "Bonus retirados com sucesso!"

case 713
wrt = "Confirma a exclus�o do bonus do(s) aluno(s) abaixo?"


'Mensagens de sistema de 9700 a 9999
case 9700
wrt = "Acesso n�o permitido a esta fun��o!"

case 9701
wrt = "Acesso permitido somente para consulta!"

case 9702
wrt = "Para imprimir clique <a class='avisos' href='#' onClick=MM_openBrWindow('imprime.asp?or=01&obr="&obr&"&p=p','','status=yes,menubar=yes,scrollbars=yes,resizable=yes,width=1030,height=500,top=50,left=50')>aqui</a>."

case 9703
wrt = "Aten&ccedil;&atilde;o! Ano Letivo est&aacute; Finalizado. As fun&ccedil;&otilde;es s&oacute; poder&atilde;o ser consultadas!<a href=../inicio.asp><img src=../img/ok.gif align=absbottom></a>"

case 9704
wrt = "Selecione as op��es desejadas."

case 9705
wrt = "Dados alterados com sucesso!"

case 9706
wrt = "Selecione os par�metros desejados"

case 9707
wrt = "Resultado encontrado de acordo com par�metros informados"

case 9708
wrt = "Altere os dados necess�rios"

case 9709
wrt = "Dados alterados com sucesso"

case 9710
wrt = "ERRO!"

case 9711
wrt = "Digite a matr�cula ou o nome do aluno"

case 9712
wrt = "Carregando. Aguarde a abertura da nova janela."

case 9713
wrt = "Ano Letivo Inv&aacute;lido."

end select




SELECT CASE tab


' primeira tela
case 0

%>
<table width="1000" height="52" border="3" align="center" cellpadding="0" cellspacing="0" bordercolor="#EEEEEE" class="aviso1">
  <tr> 
            
    <td height="46"> <div align="center"> 
      <%SELECT CASE nivel
				case 0%>
      <img src="img/atencao.gif" width="23" height="25" align="absmiddle"> 
      <%case 1%>
      <img src="../img/atencao.gif" width="23" height="25" align="absmiddle"> 
      <%		case 2%>
      <img src="../../img/atencao.gif" width="23" height="25" align="absmiddle"> 
      <%		case 3%>
      <img src="../../../img/atencao.gif" width="23" height="25" align="absmiddle"> 
      <%		case 4%>
      <img src="../../../../img/atencao.gif" width="23" height="25" align="absmiddle"> 
    <%end select%>
		  <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>
          <%response.Write(wrt)%>
          </strong></font></div>
                </div></td>
          </tr>
        </table>
		
<%
' erro
case 1
%>
<table width="1000" height="30" border="3" align="center" cellpadding="0" cellspacing="0" bordercolor="#EEEEEE" bgcolor="#FFE8E8" class="aviso2">
  <tr> 
            <td> <div align="center"> 
                <p>
		<%SELECT CASE nivel
				case 0%>
				<img src="img/pare.gif" width="28" height="25" align="absmiddle">
		<%case 1%>
				<img src="../img/pare.gif" width="28" height="25" align="absmiddle">
		<%case 2%>
				<img src="../../img/pare.gif" width="28" height="25" align="absmiddle">
		<%case 3%>
				<img src="../../../img/pare.gif" width="28" height="25" align="absmiddle">
		<%case 4%>
				<img src="../../../../img/pare.gif" width="28" height="25" align="absmiddle">												
		<%end select%>
                <font color="#CC0000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%response.Write(wrt)%></strong></font> 
                </p>
              </div></td>
  </tr>
</table>
<%
' inclus�o / altera��o de dados
case 2
%>
<table width="1000" height="30" border="3" align="center" cellpadding="0" cellspacing="0" bordercolor="#EEEEEE" bgcolor="#F2F9EE">
  <tr> 
            <td> <div align="center"> 
        <p>
		<%SELECT CASE nivel
						case 0%>
				<img src="img/atencao2.gif" width="23" height="25" align="absmiddle">
		<%case 1%>
		<img src="../img/atencao2.gif" width="23" height="25" align="absmiddle"> 
		<%case 2%>
				<img src="../../img/atencao2.gif" width="23" height="25" align="absmiddle">
		<%case 3%>
				<img src="../../../img/atencao2.gif" width="23" height="25" align="absmiddle">
  <%case 4%>
				<img src="../../../../img/atencao2.gif" width="23" height="25" align="absmiddle">		
		<%end select%>		
          <font color="#CC0000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
          <%response.Write(wrt)%>
          </strong></font> </p>
              </div></td>
          </tr>
        </table>
<%
end select

End Function

' Verifica Acesso
Function VerificaAcesso (CON,chave,nivel)
'0 - Sem Acesso, 1 - S� Consulta , 2 - S� Inclui,  3 - S� Altera, 4 - S� Exclui e  5 - Acesso Completo
chavearray=split(chave,"-")
sistema=chavearray(0)
modulo=chavearray(1)
setor=chavearray(2)
funcao=chavearray(3)
grupo=session("grupo")
	
		Set RSac = Server.CreateObject("ADODB.Recordset")
		SQLac = "SELECT * FROM TB_Autoriz_Grupo_Funcao where CO_Grupo = '"&grupo&"' and CO_Funcao = '"&funcao&"' and CO_Setor = '"&setor&"' and CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"'"
		RSac.Open SQLac, CON

		funcao_acesso=RSac("TP_Acesso")
funcao_acesso=funcao_acesso*1
Select case funcao_acesso
case 0
autoriza="no"

case 1
autoriza="con"

case 2
autoriza="in"

case 3
autoriza="al"

case 4
autoriza="ex"

case 5
autoriza="full"
end select

Session("autoriza")=autoriza
End Function

'///////////////////////////////////////////////    Grava LOG  //////////////////////////////////////////////////////////////
Function GravaLog (nvg,outro)

onde = Split(nvg, "-")
stm=onde(0)
mdl=onde(1)
str=onde(2)
fc=onde(3)

	co_usr = session("co_user")
	
	
	hora = DatePart("h", now) 
	min = DatePart("n", now) 
	dia = DatePart("d", now) 
	mes = DatePart("m", now) 
	ano = DatePart("yyyy", now)

if dia<10 then 
dia = "0"&dia
end if

if mes<10 then
mes = "0"&mes
end if

if hora<10 then 
hora = "0"&hora
end if

if min<10 then
min = "0"&min
end if	
	 
	gravahora= hora&":"&min
	gravadata= dia&"/"&mes&"/"&ano

tx_desc = outro

		Set CONL = Server.CreateObject("ADODB.Connection") 
		ABRIRL = "DBQ="& CAMINHO_log & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONL.Open ABRIRL	

Set RSL = server.createobject("adodb.recordset")

RSL.open "TB_Log_Ocorrencias", CONL, 2, 2 'which table do you want open

RSL.addnew
RSL("CO_Sistema") = stm
RSL("CO_Modulo") = mdl
RSL("CO_Setor") = str
RSL("CO_Funcao") = fc
RSL("TX_Descricao") = tx_desc
RSL("HO_ult_Acesso") = gravahora
RSL("DA_Ult_Acesso") = gravadata
RSL("CO_Usuario") = co_usr
RSL.update
  
set RSL=nothing

end function


%>
