<!--#include file="caminhos.asp"-->
<!--#include file="banner.asp"-->
<%'response.Charset="ISO-8859-1"%> 
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

sistema_nome=server.HTMLEncode(RSc1("TX_Descricao"))
link_sistema=RSc1("CO_Pasta")

origem = "Voc&ecirc; est&aacute; em <a href='../inicio.asp' class='caminho' target='_self'>Web Diretor</a> > "&sistema_nome
case 2

chavearray=split(chave,"-")
sistema=chavearray(0)
modulo=chavearray(1)



		Set RSc1 = Server.CreateObject("ADODB.Recordset")
		SQLc1 = "SELECT * FROM TB_Sistema where CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc1.Open SQLc1, Conexao
		
		sistema_nome=server.HTMLEncode(RSc1("TX_Descricao"))
		link_sistema=RSc1("CO_Pasta")



		Set RSc2 = Server.CreateObject("ADODB.Recordset")
		SQLc2 = "SELECT * FROM TB_Modulo where CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc2.Open SQLc2, Conexao

		modulo_nome=server.HTMLEncode(RSc2("TX_Descricao"))
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
		
		sistema_nome=server.HTMLEncode(RSc1("TX_Descricao"))
		link_sistema=RSc1("CO_Pasta")

		Set RSc2 = Server.CreateObject("ADODB.Recordset")
		SQLc2 = "SELECT * FROM TB_Modulo where CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc2.Open SQLc2, Conexao

		modulo_nome=server.HTMLEncode(RSc2("TX_Descricao"))
		link_modulo=RSc2("CO_Pasta")
		
		Set RSc3 = Server.CreateObject("ADODB.Recordset")
		SQLc3 = "SELECT * FROM TB_Setor where CO_Setor = '"&setor&"' and CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc3.Open SQLc3, Conexao

		setor_nome=server.HTMLEncode(RSc3("TX_Descricao"))
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


		Set RSac = Server.CreateObject("ADODB.Recordset")
		SQLac = "SELECT * FROM TB_Autoriz_Grupo_Funcao where CO_Funcao = '"&funcao&"' and CO_Setor = '"&setor&"' and CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' and CO_Grupo= '"&grupo&"'"
		RSac.Open SQLac, Conexao
		
		Set RSal = Server.CreateObject("ADODB.Recordset")
		SQLal = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&ano_letivo&"'"
		RSal.Open SQLal, Conexao
		
		sit_an=RSal("ST_Ano_Letivo")
		
		autoriza=RSac("TP_Acesso")
		if autoriza="0" and negado<>"1" then
		nvg=sistema&"-"&modulo&"-"&setor&"-"&funcao
		response.Redirect("../../../../inc/negado.asp?nvg="&nvg&"&neg=1")
		elseif autoriza="1" then
		session("trava")="s"
		elseif autoriza="5"  AND sit_an="L"then
		session("trava")="n"
		else
		session("trava")="s"
		end if
		
		session("autoriza")=autoriza

		Set RSc1 = Server.CreateObject("ADODB.Recordset")
		SQLc1 = "SELECT * FROM TB_Sistema where CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc1.Open SQLc1, Conexao
		
		sistema_nome=server.HTMLEncode(RSc1("TX_Descricao"))
		link_sistema=RSc1("CO_Pasta")

		Set RSc2 = Server.CreateObject("ADODB.Recordset")
		SQLc2 = "SELECT * FROM TB_Modulo where CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc2.Open SQLc2, Conexao

		modulo_nome=server.HTMLEncode(RSc2("TX_Descricao"))
		link_modulo=RSc2("CO_Pasta")
		
		Set RSc3 = Server.CreateObject("ADODB.Recordset")
		SQLc3 = "SELECT * FROM TB_Setor where CO_Setor = '"&setor&"' and CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc3.Open SQLc3, Conexao

		setor_nome=server.HTMLEncode(RSc3("TX_Descricao"))
		link_setor=RSc3("CO_Pasta")
		
		Set RSc4 = Server.CreateObject("ADODB.Recordset")
		SQLc4 = "SELECT * FROM TB_Funcao where CO_Funcao = '"&funcao&"' and CO_Setor = '"&setor&"' and CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc4.Open SQLc4, Conexao

		funcao_nome=server.HTMLEncode(RSc4("TX_Descricao"))
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
wrt = server.HTMLEncode("Escolha uma das opções abaixo")

case 1
wrt = server.HTMLEncode("Selecione uma unidade, um curso, uma etapa e uma turma. ")

case 2
wrt = server.HTMLEncode("Selecione uma etapa e uma turma.")

case 3
wrt = server.HTMLEncode("Selecione uma etapa, uma turma, um período e uma avalia&ccedil;&atilde;o.")

case 4
wrt = server.HTMLEncode("Para consultar é necessário selecionar uma etapa!")

case 5
wrt = server.HTMLEncode("Esta fun&ccedil;&atilde;o permite você fazer contato com a equipe técnica que realiza a manuten&ccedil;&atilde;o do sistema Web Diretor. Utilize sempre que possível este canal para nos transmitir alguma informa&ccedil;&atilde;o relevante sobre o funcionamento desse produto. Obrigado pela sua aten&ccedil;&atilde;o!")

case 6
wrt =server.HTMLEncode("Mensagem enviada.")

case 7
wrt = server.HTMLEncode("Escolha um novo usuário.")

case 8
wrt = server.HTMLEncode("Escolha uma nova senha.")

case 9
wrt = server.HTMLEncode("Usuário alterado com sucesso.")

case 10
wrt = server.HTMLEncode("Senha alterada com sucesso.")

case 11
wrt = server.HTMLEncode("Selecione uma disciplina e um período.")

case 12
wrt = server.HTMLEncode("E-mail alterado com sucesso.")

case 13
wrt = server.HTMLEncode("Usuário já existe!")

case 14
wrt = server.HTMLEncode("Digite seu novo endere&ccedil;o de correio eletrônico")

case 15
wrt = server.HTMLEncode("Endere&ccedil;o de correio eletrônico já existe!")

case 16
wrt = server.HTMLEncode("Selecione uma etapa, uma turma e um período.")

case 17
wrt = server.HTMLEncode("Selecione uma etapa e um período.")

case 18
wrt = server.HTMLEncode("Gráfico comparativo.")

case 19
wrt = server.HTMLEncode("Selecione uma etapa, uma disciplina e um período.")

case 20
wrt = server.HTMLEncode("Selecione uma etapa")

case 21
wrt = server.HTMLEncode("Clique na op&ccedil;&atilde;o desejada")

case 22
wrt = server.HTMLEncode("Confirma a reinicializa&ccedil;&atilde;o da senha do usuário abaixo?")

case 23
wrt = server.HTMLEncode("Senha reinicializada com sucesso")

case 24
wrt = server.HTMLEncode("Usuário "&situacao&" com sucesso")

case 25
wrt = server.HTMLEncode("Selecione uma unidade, um curso, uma etapa e uma média.")

'Web Família de 50 a 99



case 51
wrt = server.HTMLEncode("Selecione o tipo de documento e os arquivos que deseja disponibilizar para upload")

case 52
wrt = server.HTMLEncode("Arquivo(s) "&Session("arquivos") &" enviado(s) com sucesso! Total de Bytes enviados:"&Session("upl_total"))

case 53
wrt = server.HTMLEncode("Selecione pelo menos um arquivo")

case 54
wrt = server.HTMLEncode("Preencha os dados abaixo para associar um documento")

case 55
wrt = server.HTMLEncode("Associa&ccedil;&atilde;o realizada com Sucesso")

case 56
wrt = server.HTMLEncode("Preencha os dados abaixo para incluir uma notícia")

case 57
wrt = server.HTMLEncode("Notícia incluída com sucesso")

case 58
wrt = server.HTMLEncode("Confirma a exclus&atilde;o do(s) documento(s) abaixo?")

case 59
wrt = server.HTMLEncode("Documento(s) excluído(s) com sucesso")

case 60
wrt = server.HTMLEncode("Confirma a exclus&atilde;o do(s) arquivo(s) abaixo?")

case 61
wrt = server.HTMLEncode("Arquivo(s) excluído(s) com sucesso")

case 62
wrt = server.HTMLEncode("Selecione o tipo de documento")

case 63
wrt = server.HTMLEncode("Confirma a exclus&atilde;o da(s) notícia(s) abaixo?")

case 64
wrt = server.HTMLEncode("Notícia(s) excluída(s) com sucesso")

case 65
wrt = server.HTMLEncode("Confirma a exclus&atilde;o do(s) evento(s) abaixo?")

case 66
wrt = server.HTMLEncode("Evento(s) excluído(s) com sucesso")

case 67
wrt = server.HTMLEncode("Preencha os dados abaixo para incluir um evento")

case 68
wrt = server.HTMLEncode("Evento incluído com sucesso")

case 69
wrt = server.HTMLEncode("Para consultar os dados do usuário digite o código ou Nome e clique no bot&atilde;o Procurar.")

case 70
wrt = server.HTMLEncode("Escolha um usuário para consultar o cadastro.")

case 71
wrt = server.HTMLEncode("Verifique os dados do usuário.")

case 72
wrt = server.HTMLEncode("N&atilde;o foi encontrado nenhum usuário com este código.")

' erro na busca por nome
case 73
wrt = server.HTMLEncode("N&atilde;o foi encontrado nenhum usuário com este nome.")

case 74
wrt = server.HTMLEncode("Pasta criada com sucesso!")

case 75
wrt = server.HTMLEncode("Pasta modificada com sucesso!")

case 76
wrt = server.HTMLEncode("Preencha os dados abaixo para incluir uma mensagem")

case 77
wrt = server.HTMLEncode("Mensagem incluída com sucesso!")

case 78
wrt =server.HTMLEncode( "Mensagem excluída com sucesso!")

case 79
wrt = server.HTMLEncode("Confirma a exclus&atilde;o da(s) mensagem(ns) abaixo?")



'alunos de 300 a 399
case 300
wrt = "Para consultar os dados do Aluno digite a Matrícula ou Nome e clique no bot&atilde;o Procurar."

' listagem de alunos

case 301
wrt = "Escolha um Aluno para consultar o cadastro."

case 302
wrt = "Verifique os dados do Aluno."

case 303
wrt = "N&atilde;o foi encontrado nenhum Aluno com este código."

' erro na busca por nome
case 304
wrt = "N&atilde;o foi encontrado nenhum Aluno com este nome."

case 305
wrt = "Lista de alunos associados a turma abaixo."

case 306
wrt = "Verifique os dados dos familiares."

case 307
wrt = "Selecione uma unidade e um mês."

case 308
wrt = "Comparar Turma por Média Geral."

case 309
wrt = "Verifique os dados do Aluno e escolha uma disciplina e um período."

case 310
wrt = "Escolha os critérios para pesquisar as ocorrências do aluno e clique no bot&atilde;o prosseguir."

case 311
wrt = "Confirma a exclus&atilde;o dessa(s) disciplina(s)."


case 312
wrt = "Ocorrência incluída com sucesso!"

case 313
wrt = "Ocorrência alterada com sucesso!"

case 314
wrt = "Ocorrência excluída com sucesso!"

case 315
wrt = "Preencha os dados abaixo e clique no bot&atilde;o Confirmar para Incluir uma nova ocorrência."

case 316
wrt = "Preencha os dados abaixo e clique no bot&atilde;o Confirmar para atualizar esta ocorrência."

case 317
wrt = "Selecione uma situa&ccedil;&atilde;o para o aluno e escreva o motivo da inativa&ccedil;&atilde;o."

'web secretaria 400 a 450
case 400
wrt = "Para consultar os dados do Aluno digite a Matrícula ou Nome e clique no bot&atilde;o Procurar. Caso o aluno n&atilde;o esteja cadastrado no sistema clique <a href='../../../cad/man/aal/cadastra.asp?nvg=WS-CA-MA-AAL' class='avisos'>aqui</a>."

case 401
wrt = "Matrícula efetuada com sucesso!"

case 402
wrt = "Preencha os campos abaixo."

case 403
wrt = "Aluno já matriculado para este ano letivo. Matrículas para o próximo Ano Letivo est&atilde;o fechadas!"

case 404
wrt = "Para alterar os dados do Aluno digite a Matrícula ou Nome e clique no bot&atilde;o Procurar. Caso o aluno n&atilde;o esteja cadastrado no sistema clique <a href='../../../cad/man/aal/cadastra.asp?nvg=WS-CA-MA-AAL' class='avisos'>aqui</a>."

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
wrt6 =Server.URLEncode("Confirma a exclus&atilde;o desse familiar?")

wrt = wrt1&wrt2&wrt3&wrt4&wrt5&wrt6&"<br><input type='button' name='Submit2' value='Sim' onClick='ExcluiFamiliares(ordem.value,qtd.value,foco.value,cod_al.value)' class='botao_prosseguir_sim' >&nbsp;&nbsp;&nbsp;<input type='button' name='Submit2' value='"&Server.URLEncode("N&atilde;o")&"' onClick='recuperarFamiliares(ordem.value,qtd.value,foco.value,cod_vinculado.value,cod_al.value)' class='botao_prosseguir_nao' >"

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

wrt = wrt1&wrt2&wrt3&wrt4&wrt5&wrt6&"<br><input type='button' name='Submit2' value='Sim' onClick='recuperarFamiliares(ordem.value,qtd.value,foco.value,cod_vinculado.value,cod_al.value)' class='botao_prosseguir_sim' >&nbsp;&nbsp;&nbsp;<input type='button' name='Submit2' value='"&Server.URLEncode("N&atilde;o")&"' onClick='ExcluiFamiliares(ordem.value,qtd.value,foco.value,cod_al.value)' class='botao_prosseguir_nao' >"

case 407
wrt = "Deve ser selecionado um responsável financeiro para o aluno!"

case 408
wrt = "Deve ser selecionado um responsável pedagógico para o aluno!"

case 409
wrt = "É obrigatório o preenchimento dos campos: Nome, Telefones de Contato e Endere&ccedil;o residencial para o responsável financeiro!"

case 410
wrt = "É obrigatório o preenchimento dos campos: Nome, Telefones de Contato e Endere&ccedil;o residencial para o responsável pedagógico!"

case 411
wrt = "Ao se confirmar o cadastro desse aluno, esse número de matrícula n&atilde;o poderá mais ser utilizado!"

case 412
wrt = "Cadastro efetuado com sucesso! Inclua todos os dados necessários."

case 413
wrt = "Selecione uma nova combina&ccedil;&atilde;o de Unidade, Curso, Etapa, Turma e Número de chamada para o aluno."

case 414
wrt = "Selecione um método para enturmar os alunos em situa&ccedil;&atilde;o de pré-matrícula."

case 415
wrt = "N&atilde;o existem alunos em situa&ccedil;&atilde;o de pré-matrícula."

case 416
wrt = "Somente é possível remanejar alunos com situa&ccedil;&atilde;o igual a 'Cursando'."


'professores de 600 a 699

case 600
wrt =  "Os Professores em vermelho est&atilde;o inativos. A mensagem 'n&atilde;o cadastrado' indica que n&atilde;o existe professor associado àquela disciplina naquela turma"
wrt = wrt &"<br>A mensagem 'nome em branco' indica que o nome do professor n&atilde;o está registrado no cadastro. Para bloquear a planilha clique na letra 'N' do período escolhido"

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
act= "bloqueada"
elseif orig=02 then
act= "desbloqueada"
end if

wrt = "Planilha "&act&" com sucesso!"

case 603
wrt = "Avalia&ccedil;ões n&atilde;o lan&ccedil;adas"

case 604
wrt = "Para consultar a Grade de aulas digite o C&oacute;digo ou Nome de um Professor e clique no bot&atilde;o Procurar."
wrt = wrt &"<br>Se preferir obter uma lista completa de TODOS os professores clique <a href='index.asp?opt=listall&nvg="&nvg&"' class='avisos'>aqui</a>"

case 605
wrt = "N&atilde;o foi encontrado nenhum professor com este código."

case 606
wrt = "Escolha um professor para consultar a Grade de Aulas. Os Professores em vermelho est&atilde;o inativos."

case 607
wrt = "Para atualizar os dados do Professor digite o C&oacute;digo ou Nome e clique no bot&atilde;o Procurar."
wrt = wrt &"Se preferir adicionar um NOVO professor clique <a href='altera.asp?ori=02&nvg="&nvg&"' class='avisos'>aqui</a>."
wrt = wrt &"<BR>Se preferir obter uma lista completa de TODOS os professores clique <a href='index.asp?opt=listall&nvg="&nvg&"' class='avisos'>aqui</a>"

case 608
wrt = "Confirme o professor para consultar a Grade de Aulas."

case 609
wrt = "O período relacionado pela letra 'S' indica que a planilha está Bloqueada e 'N' que está Desbloqueada."

case 610
wrt = "N&atilde;o foi encontrado nenhum professor com este código."

case 611
wrt = "N&atilde;o foi encontrado nenhum professor com este nome."

case 612
wrt = "Escolha um professor para atualizar o cadastro. Os Professores em vermelho est&atilde;o inativos."

case 613
wrt = "Confirme se é o professor correto para atualizar o cadastro."

case 614
wrt = "Preencha cuidadosamente os dados do Professor e click no bot&atilde;o CONFIRMAR para atualizar o cadastro"

case 615
wrt = "Professor código "&cod_cons&" e usuário "&escola&co_usr_prof&" incluído com sucesso!"

case 616
wrt = "Dados do Professor código "&cod_cons&" alterados com sucesso!"

case 617
wrt = "Selecione a Data e a Hora as quais você deseja iniciar o monitoramento de notas e clique em iniciar."

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
wrt = "Inicio da monitora&ccedil;&atilde;o a partir do dia "&dia_mnl&"/"&mes_wrt&"/"&ano_mnl&" as "&hora_mnl&":"&min_wrt&" Dados atualizados a cada minuto."

case 619
wrt = "N&atilde;o foram encontradas turmas cadastradas para você. Entre em contato com o seu coordenador."


case 620
if errou="pv1" or errou="pv2" or errou="pv3" or errou="pv4" or errou="pv5" or errou="pv6" Then
wrt = "Valor inválido para o campo  "&errado
elseif errou="sp" Then
wrt = "Soma dos Pesos maior que 10"
elseif errou="pt" Then
wrt = "Um dos pesos tem valor inválido"
elseif errou="pr1pr2" Then
wrt = "Soma das Pr's maior que 10"
else
	wrt = "Valor inválido para o campo  "&errado&"  do número de chamada <b>"&errante&"</b>."
	if erro_mae = "S" then
		wrt = "Erro ao agregar as notas para gerar as médias da disciplina "&no_materia_mae&".<br><br>"&wrt&"<br><br>Devem ser verificadas as notas lan&ccedil;adas de <u>todas as disciplinas que compõem essas médias</u>" 
	end if	
end if

' erro na busca por código
case 621
wrt = "Você está " 
if opt="cln" then
wrt= wrt &"comunicando"
else
wrt= wrt &"lan&ccedil;ando"
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
wrt = "Notas lan&ccedil;adas com sucesso."

case 623
if dados = "I" then
	wrt = "Importando dados de Freq&uuml;&ecirc;ncia Aguarde!"
else
	wrt = "Comunicado efetuado!"
end if	

case 624
wrt = "Estas notas j&aacute; foram lan&ccedil;adas.Para alter&aacute;-las pe&ccedil;a autoriza&ccedil;&atilde;o ao coordenador"

case 625
wrt = "Escolha um Coordenador para consultar os Professores sob sua coordena&ccedil;&atilde;o."

case 626
wrt = "Os Professores em vermelho est&atilde;o inativos. A mensagem 'n&atilde;o cadastrado'indica que n&atilde;o existe professor associado àquela disciplina naquela turma"
wrt = wrt &"<br>A mensagem 'nome em branco' indica que o nome do professor n&atilde;o está registrado no cadastro"

case 627
wrt = "Para excluir, selecione uma ou mais disciplinas e clique em excluir.<br>Para incluir uma nova disciplina na Grade de Aulas, selecione uma unidade e um curso."

case 628
wrt = "Disciplina incluída com sucesso"

case 629
wrt = "Disciplina excluída com sucesso"

case 630
wrt = "N&atilde;o é possível marcar uma disciplina na Grade de Aulas e selecionar uma unidade e um curso ao mesmo tempo.<br>Por favor selecione somente disciplina(s) para excluir ou selecione uma unidade para incluir uma nova disciplina na Grade de Aulas"

case 631
wrt = "Selecione uma disciplina, um modelo e um coordenador."

case 632
wrt = "Para atualizar é necessário selecionar uma disciplina,um modelo e um coordenador"

case 633
wrt = "Verifique os dados preenchidos e clique no bot&atilde;o Confirmar para continuar a inclus&atilde;o ou no bot&atilde;o Alterar para voltar e modificar algum dado."


case 634
wrt = "Verifique as disciplinas selecionadas e clique no bot&atilde;o confirmar para Excluir ou no bot&atilde;o Cancelar para voltar e modificar algum dado."

case 635
wrt = "Professores que n&atilde;o comunicaram."

case 636
wrt = "Para imprimir clique <a class='avisos' href='#' onClick=MM_openBrWindow('imprime.asp?or=01&obr="&obr&"&p=p','','status=yes,menubar=yes,scrollbars=yes,resizable=yes,width=1030,height=500,top=50,left=50')>aqui</a>."

case 637
wrt = "Escolha um professor e um período."

case 638
wrt =  "Os Professores em vermelho est&atilde;o inativos. A mensagem 'n&atilde;o cadastrado' indica que n&atilde;o existe professor associado àquela disciplina naquela turma"
wrt = wrt &"<br>A mensagem 'nome em branco' indica que o nome do professor n&atilde;o está registrado no cadastro. Clique no nome da disciplina para ver o mapa de resultado."

case 639
wrt = "Arquivo "& fl &" enviado com sucesso."

case 640
wrt = "Aten&ccedil;&atilde;o! Estas notas j&aacute; foram lan&ccedil;adas pelo professor."

case 641
wrt = "Inclua as faltas no período desejado"

case 642
wrt = "Faltas lan&ccedil;adas com sucesso"

case 643
wrt = "Para atualizar os dados do Professor digite o C&oacute;digo ou Nome e clique no bot&atilde;o Procurar."
wrt = wrt &"<BR>Se preferir obter uma lista completa de TODOS os professores clique <a href='index.asp?opt=listall&nvg="&nvg&"' class='avisos'>aqui</a>"

case 644
wrt = "É necessário escolher pelo menos uma unidade"

case 645
wrt = "Imprimir <a class='avisos' href='#' onClick=MM_openBrWindow('imprime.asp?obr="&obr&"&p=p','','status=yes,menubar=yes,scrollbars=yes,resizable=yes,width=1030,height=500,top=50,left=50')>html</a> / <a class='avisos' href='../../../../relatorios/swd015.asp?obr="&obr&"'>pdf</a>."

case 646
wrt = "Para carregar as notas do simulado &eacute; necess&aacute;rio que o arquivo seja o modelo padr&atilde;o que pode ser baixado <a href=resultado.xls>aqui</a>."

case 647
wrt = "Arquivo "&Session("arquivo") &" enviado com sucesso! Total de Bytes enviados:"&Session("upl_total")

case 648
wrt = "Nenhum arquivo selecionado!"

case 649
wrt = "O arquivo "&Session("arquivo") &" n&atilde;o possui o nome correto!"

case 650
dados_erro=split(dados,"$!$")
wrt = "Matrícula "&dados_erro(0)&" n&atilde;o cadastrada em TB_Matriculas para o Ano Letivo "&dados_erro(1)&"!. Favor verificar e reenviar a planilha!"

case 651
dados_erro=split(dados,"$!$")
wrt = "Erro na grava&ccedil;&atilde;o da matrícula: "&dados_erro(0)&", disciplina: "&dados_erro(1)&", nota: "&dados_erro(2)&". Favor corrigir e reenviar a planilha!"


case 652
separa_dados=split(dados,"#$#")
separa=split(separa_dados(0),"#!#")
no_unidade=separa(0)
no_curso=separa(1)
no_etapa=separa(2)
data_grav=separa(3)
hora_grav=separa(4)
obr_mapa=separa_dados(1)
wrt = "Existem informa&ccedil;&otilde;es geradas em "&data_grav&" às "&hora_grav&" para a Unidade: "&no_unidade&", Curso: "&no_curso&", Etapa: "&no_etapa&" e Turma: "&turma&".<BR>Deseja reprocessar essas informa&ccedil;&otilde;es? <a class='avisos' href='gera_base.asp?opt=rgnrt&obr="&obr_mapa&"' onclick=redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show')>sim</a> / <a class='avisos' href='#' onClick=redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show');MM_openBrWindow('mapa.asp?obr="&obr_mapa&"','')>n&atilde;o</a>"

case 653
wrt = "O arquivo está sendo gerado!"

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
		wrt = server.HTMLEncode( "Existem informa&ccedil;ões geradas em "&data_grav&" às "&hora_grav&" para o aluno "&no_aluno&", matrícula "&co_aluno&".")
		javascript="MM_callJS('document.busca.busca1.focus()');"
	else
		separa=split(separa_dados(1),"#!#")
		no_unidade=separa(0)
		no_curso=separa(1)
		no_etapa=separa(2)
		data_grav=separa(3)
		hora_grav=separa(4)
		wrt =  server.HTMLEncode("Existem informa&ccedil;ões geradas em "&data_grav&" às "&hora_grav&" para a Unidade: "&no_unidade&", Curso: "&no_curso&", Etapa: "&no_etapa&" e Turma: "&turma&".")
		javascript=""
	end if
	obr_mapa=separa_dados(2)
	wrt = wrt&"<BR>Deseja reprocessar essas informa&ccedil;&otilde;es? <a class='avisos' href='gera_base.asp?opt=rgnrt&obr="&separa_dados(0)&"$$$"&obr_mapa&"' onclick=redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show')>sim</a> / <a class='avisos' href='../../../../relatorios/swd025.asp?obr="&separa_dados(0)&"$$$"&obr_mapa&"&ori=ebe' onclick=""redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show','divTabela','','show');AlternarMensagem('divMensagem2');mclosetime();"&javascript&""">n&atilde;o</a> "
	
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
		wrt =  server.HTMLEncode("Existem informa&ccedil;ões geradas em "&data_grav&" às "&hora_grav&" para o aluno "&no_aluno&", matrícula "&co_aluno&".")
		javascript="MM_callJS('document.busca.busca1.focus()');"
	else
		separa=split(separa_dados(1),"#!#")
		no_unidade=separa(0)
		no_curso=separa(1)
		no_etapa=separa(2)
		data_grav=separa(3)
		hora_grav=separa(4)
		wrt = server.HTMLEncode("Existem informa&ccedil;ões geradas em "&data_grav&" às "&hora_grav&" para a Unidade: "&no_unidade&", Curso: "&no_curso&", Etapa: "&no_etapa&" e Turma: "&turma&".")
		javascript=""
	end if
	obr_mapa=separa_dados(2)
	wrt = wrt&"<BR>Deseja reprocessar essas informa&ccedil;&otilde;es?  <a class='avisos' href='gera_base.asp?opt=rgnrt&obr="&separa_dados(0)&"$$$"&obr_mapa&"' onclick=redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show')>sim</a> / <a class='avisos' href='../../../../relatorios/swd048.asp?obr="&separa_dados(0)&"$$$"&obr_mapa&"&ori=efi' onclick=""redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show','divTabela','','show');AlternarMensagem('divMensagem2');mclosetime();"&javascript&""">n&atilde;o</a>"
	
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
wrt = "Existem informa&ccedil;ões geradas em "&data_grav&" às "&hora_grav&" para a Unidade: "&no_unidade&", Curso: "&no_curso&", Etapa: "&no_etapa&" e Turma: "&turma&".<BR>Deseja reprocessar essas informa&ccedil;ões? <a class='avisos' href='gera_base.asp?opt=rgnrt&obr="&obr_mapa&"' onclick=redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show')>sim</a> / <a class='avisos' href='#' onClick=redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show');MM_openBrWindow('mapa.asp?obr="&obr_mapa&"','')>n&atilde;o</a>"

case 659
wrt = "O arquivo está sendo gerado!"


case 660
separa_dados=split(dados,"#$#")
separa=split(separa_dados(0),"#!#")
no_unidade=separa(0)
no_curso=separa(1)
no_etapa=separa(2)
data_grav=separa(3)
hora_grav=separa(4)
obr_mapa=separa_dados(1)
wrt = server.HTMLEncode("Existem informa&ccedil;ões geradas em "&data_grav&" às "&hora_grav&" para a Unidade: "&no_unidade&", Curso: "&no_curso&", Etapa: "&no_etapa&" e Turma: "&turma&".")&"<BR>Deseja reprocessar essas informa&ccedil;&otilde;es? <a class='avisos' href='gera_base.asp?opt=rgnrt&obr="&obr_mapa&"' onclick=redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show')>sim</a> / <a class='avisos' href=gera_pdf.asp?obr="&obr_mapa&" onclick=""redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show','divTabela','','show');AlternarMensagem('divMensagem2');mclosetime();"">n&atilde;o</a>"

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
		periodo_msg=" e todos os Períodos!"
	else
'		divisao=tipo_divisao_ano(curso,co_etapa,tipo_dado)
'		periodo_msg_t = periodos(periodo_fn, divisao, "nome")
		periodo_msg_t = periodos(periodo_fn, "nome")		
		periodo_msg=" do "&periodo_msg_t&"!"
	end if	
end if	
wrt = origem_fn&" com sucesso para "&unidade_msg&curso_msg&co_etapa_msg&turma_msg&periodo_msg

case 663
separa_dados=split(dados,"$$$")
tipo_busca=split(separa_dados(0),"$!$")
tipo=tipo_busca(0)
	if tipo="a" then
		separa=split(separa_dados(1),"#!#")
		no_aluno=separa(0)
		co_aluno=separa(1)
		data_grav=separa(2)
		hora_grav=separa(3)
		wrt =  "Existem informa&ccedil;ões geradas em "&data_grav&" às "&hora_grav&" para o aluno "&no_aluno&", matrícula "&co_aluno&"."
		javascript="MM_callJS('document.busca.busca1.focus()');"
	else
		separa=split(separa_dados(1),"#!#")
		no_unidade=separa(0)
		no_curso=separa(1)
		no_etapa=separa(2)
		data_grav=separa(3)
		hora_grav=separa(4)
		wrt =  "Existem informa&ccedil;ões geradas em "&data_grav&" às "&hora_grav&" para a Unidade: "&no_unidade&", Curso: "&no_curso&", Etapa: "&no_etapa&" e Turma: "&turma&"."
		javascript=""
	end if
	obr_mapa=separa_dados(2)
	wrt = wrt&"<BR>O ano letivo está bloqueado e n&atilde;o é possível regerar as informa&ccedil;ões; Para fazer o download do relatório previamente gerado clique <a class='avisos' href='../../../../relatorios/swd048.asp?obr="&separa_dados(0)&"$$$"&obr_mapa&"&ori=efi' onclick=""redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show','divTabela','','show');AlternarMensagem('divMensagem2');mclosetime();"&javascript&""">aqui</a>"

case 664
separa_dados=split(dados,"$$$")
tipo_busca=split(separa_dados(0),"$!$")
tipo=tipo_busca(0)
	if tipo="a" then
		separa=split(separa_dados(1),"#!#")
		no_aluno=separa(0)
		co_aluno=separa(1)
		data_grav=separa(2)
		hora_grav=separa(3)
		wrt =  server.HTMLEncode("Existem informa&ccedil;ões geradas em "&data_grav&" às "&hora_grav&" para o aluno "&no_aluno&", matrícula "&co_aluno&".")
		javascript="MM_callJS('document.busca.busca1.focus()');"
	else
		separa=split(separa_dados(1),"#!#")
		no_unidade=separa(0)
		no_curso=separa(1)
		no_etapa=separa(2)
		data_grav=separa(3)
		hora_grav=separa(4)
		wrt =  "Existem informa&ccedil;&otilde;es geradas em "&data_grav&" às "&hora_grav&" para a Unidade: "&no_unidade&", Curso: "&no_curso&", Etapa: "&no_etapa&" e Turma: "&turma&"."
		javascript=""
	end if
	obr_mapa=separa_dados(2)
	wrt = wrt&server.HTMLEncode("<BR>O ano letivo está bloqueado e n&atilde;o é possível regerar as informa&ccedil;ões; Para fazer o download do relatório previamente gerado clique")&" <a class='avisos' href='../../../../relatorios/swd025.asp?obr="&separa_dados(0)&"$$$"&obr_mapa&"&ori=ebe' onclick=""redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show','divTabela','','show');AlternarMensagem('divMensagem2');mclosetime();"">aqui</a>"


case 665
wrt = "Conte&uacute;do lecionado lan&ccedil;ado com sucesso"

case 666
wrt = "Conte&uacute;do lecionado excluído com sucesso."

case 667
wrt = "Imprimir <a class='avisos' href='../../../../relatorios/swd300.asp?obr="&obr&"'>pdf</a>."

case 668
wrt = "Imprimir <a class='avisos' href='../../../../relatorios/swd301.asp?obr="&obr&"'>pdf</a>."

case 669
wrt = "Exibindo apenas as "&dados&" datas mais recentes."

case 670
if isnumeric(dados) then
	wrt = "Preparando para exibir apenas as "&dados&" datas mais recentes."
else
	wrt = "Preparando a Pauta Completa e apurando o Total de Faltas dos Alunos."
end if

case 671
wrt = "Selecione pelo menos uma data"

case 672
wrt = "Aten&ccedil;&atilde;o! Esta Pauta j&aacute;  est&aacute;  fechada pelo professor."

case 673
wrt = "Pauta lan&ccedil;ada com sucesso."

case 674
wrt = "Pauta excluída com sucesso."




case 700
wrt = server.HTMLEncode("O aluno solicitado n&atilde;o está ATIVO no ano letivo de "&session("ano_letivo")&"!")

case 701
wrt = server.HTMLEncode("N&atilde;o é possível gerar boletim para alunos da Educa&ccedil;&atilde;o Infantil!")

case 702
wrt = "A op&ccedil;&atilde;o em destaque é a que foi selecionada. É possível alterar o gráfico selecionando uma das outras op&ccedil;ões."

case 703
wrt = "N&atilde;o existem alunos APROVADOS que atendam as condi&ccedil;ões de busca solicitadas."

case 704
wrt = "N&atilde;o existem alunos REPROVADOS que atendam as condi&ccedil;ões de busca solicitadas."

case 705
wrt = "Somente é possível gerar esse relatório para alunos do 3º ano do Ensino Médio."

case 706
wrt = server.HTMLEncode("N&atilde;o existem alunos que atendam as condi&ccedil;ões de busca solicitadas.")

case 707
wrt = "Para gerar o relatório clique em prosseguir."

case 708
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
wrt = server.HTMLEncode("O Ano Letivo está encerrado e n&atilde;o existem dados gerados para a Unidade: "&no_unidade&", Curso: "&no_curso&", Etapa: "&no_etapa&" e Turma: "&turma&". Para que sejam gerados os dados é necessária a abertura do Ano Letivo.")

case 709
separa_dados=split(dados,"$$$")
wrt = server.HTMLEncode("O arquivo n&atilde;o foi gerado! Existiam pendência de lan&ccedil;amento de notas.")

case 710
separa_dados=split(dados,"$$$")
tipo_busca=separa_dados(0)
obr_mapa=separa_dados(2)
dados_funcao=split(obr_mapa,"$!$")
unidade_fn = dados_funcao(0)
curso_fn = dados_funcao(1)
co_etapa_fn = dados_funcao(2)
turma_fn = dados_funcao(3)
periodo_fn = dados_funcao(4)
origem_fn = dados_funcao(6)
if origem_fn = "boletim" then
	arquivo_fn="swd025.asp"
	origem_fn="ebe"
elseif origem_fn = "ficha" then
	arquivo_fn="swd048.asp"
	origem_fn="efi"
end if	
wrt = "O arquivo gerado com sucesso! No entanto, existem alunos sem notas lan&ccedil;adas e por isso esses alunos n&atilde;o foram gerados. Clique <a class='avisos' href='../../../../relatorios/"&arquivo_fn&"?obr="&tipo_busca&"$$$"&obr_mapa&"&ori="&origem_fn&"' onclick=""redimensiona();MM_showHideLayers('carregando','','show','carregando_fundo','','show','divTabela','','show');AlternarMensagem('divMensagem2');mclosetime();"">aqui</a> para fazer o download."

case 711
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
		periodo_msg=" e todos os Períodos!"
	else
'		divisao=tipo_divisao_ano(curso,co_etapa,tipo_dado)
'		periodo_msg_t = periodos(periodo_fn, divisao, "nome")
		periodo_msg_t = periodos(periodo_fn, "nome")		
		periodo_msg=" do "&periodo_msg_t&"!"
	end if	
end if	
wrt = origem_fn&" com sucesso para "&unidade_msg&curso_msg&co_etapa_msg&turma_msg&periodo_msg&" No entanto, existem alunos sem notas lan&ccedil;adas e por isso esses alunos n&atilde;o foram gerados."

case 712
wrt = server.HTMLEncode("É possível alterar o gráfico selecionando uma das outras op&ccedil;ões.")

case 713
wrt = "Confirma a exclus&atilde;o do bonus do(s) aluno(s) abaixo?"

case 714
wrt = "Bonus retirados com sucesso!"

case 715
wrt = "Selecione na lista os alunos os quais deseja retirar a aprova&ccedil;&atilde;o pelo Conselho de Classe"

case 716
wrt = "Aprova&ccedil;&atilde;o pelo Conselho de Classe lan&ccedil;ada com sucesso!"

case 717
wrt = "Aprova&ccedil;&atilde;o pelo Conselho de Classe retirada com sucesso!"

case 718
wrt = "Confirma a exclus&atilde;o da Aprova&ccedil;&atilde;o pelo Conselho de Classe do(s) aluno(s) abaixo?"

case 719
wrt = "Selecione na lista os alunos os quais deseja retirar bonus"

case 720
wrt = "Bonus lan&ccedil;ados com sucesso!"


'Mensagens de web Tesouraria de 800 a 999
case 900
wrt = "Selecione o local de um dos arquivos POSICAOWEB.txt ou BOLETOWEB.txt"

case 901
wrt = "É necessário selecionar o local de um dos arquivos POSICAOWEB.txt ou BOLETOWEB.txt"

case 902
wrt = "O nome do arquivo deve obrigatoriamente ser POSICAOWEB.txt ou BOLETOWEB.txt"

case 903
wrt = "Arquivo importado com sucesso! A base de dados Posi&ccedil;&atilde;o Financeira está sendo atualizada. Por favor aguarde."
'wrt = "Arquivos importados com sucesso! A base de dados Posi&ccedil;&atilde;o Financeira está sendo atualizada. Por favor aguarde."

case 904
wrt = "Arquivo importado com sucesso! A base de dados Boletos Bancários está sendo atualizada. Por favor aguarde."
'wrt = "A base de dados Posi&ccedil;&atilde;o Financeira atualizada com sucesso! A base de dados Boletos Bancários está sendo atualizada. Por favor aguarde."

case 905
wrt = "Base de dados Posi&ccedil;&atilde;o Financeira foi atualizada com sucesso!"
'wrt = "Apenas a base de dados Posi&ccedil;&atilde;o Financeira foi atualizada com sucesso! O arquivo BOLETOWEB.txt para atualiza&ccedil;&atilde;o da base Boletos Bancários n&atilde;o foi encontrado."

case 906
wrt = "Base de dados Boletos Bancários foi atualizada com sucesso!"
'wrt = "Apenas a base de dados Boletos Bancários foi atualizada com sucesso! O arquivo POSICAOWEB.txt para atualiza&ccedil;&atilde;o da base Posi&ccedil;&atilde;o Financeira n&atilde;o foi encontrado."

case 907
wrt = "As Bases de dados financeiros foram atualizadas com sucesso!"

case 908
wrt = "Nenhuma Bases de dados financeiros foi atualizada! Os arquivos POSICAOWEB.txt e BOLETOWEB.txt n&atilde;o foram encontrados!"

case 909
if dados = "POSICAOWEB.txt" then
	complemento = "8"
else
	complemento = "28"
end if	
wrt = "O Arquivo "&dados&" está formatado errado! Ele deveria ter "&complemento&" campos separados por ;"

case 910
wrt = "Nenhum responsável foi selecionado!"

case 911
wrt = "Para confirmar o envio da mensagem abaixo para os endere&ccedil;os de e-mail indicados clique em prosseguir."

case 912
wrt = "Mensagens enviadas com sucesso!"

case 913
'wrt = "O relatório será ordenado por Unidade seguido por Nome do Aluno, Data de Vencimento e Tipo de Lan&ccedil;amento.<br>Se desejar ordenar de maneira diferente selecione um ou mais itens abaixo."
wrt = "O relatório será ordenado por Unidade seguido por Nome do Aluno, Data de Vencimento e Tipo de Lan&ccedil;amento. Clique em prosseguir para gerar o relatório."

case 914
wrt = "Enviar 2ª via de bloqueto para o email "&dados&"?"


'Mensagens de sistema de 9700 a 9999
case 9700
wrt = "Acesso n&atilde;o permitido a esta fun&ccedil;&atilde;o!"

case 9701
wrt = "Acesso permitido somente para consulta!"

case 9702
wrt = "Para imprimir clique <a class='avisos' href='#' onClick=MM_openBrWindow('imprime.asp?or=01&obr="&obr&"&p=p','','status=yes,menubar=yes,scrollbars=yes,resizable=yes,width=1030,height=500,top=50,left=50')>aqui</a>."

case 9703
wrt = "Aten&ccedil;&atilde;o! Ano Letivo est&aacute; Finalizado. As fun&ccedil;&otilde;es s&oacute; poder&atilde;o ser consultadas!<a href=../inicio.asp><img src=../img/ok.gif align=absbottom></a>"

case 9704
wrt = "Selecione as op&ccedil;ões desejadas."

case 9705
wrt = "Dados alterados com sucesso!"

case 9706
wrt = "Selecione os par&acirc;metros desejados"

case 9707
wrt = "Resultado encontrado de acordo com par&acirc;metros informados"

case 9708
wrt = "Altere os dados necessários"

case 9709
wrt = "Dados alterados com sucesso"

case 9710
wrt = "ERRO!"

case 9711
wrt = "Digite a matrícula ou o nome do aluno"

case 9712
wrt = "Carregando. Aguarde a abertura da nova janela."

case 9713
wrt = "Ano Letivo Inv&aacute;lido."

case 9714
wrt = "Email enviado com sucesso!"

case 9715
wrt = "Preencha com a combina&ccedil;&atilde;o desejada"

case 9716
wrt = "Aluno n&atilde;o matrículado no ano letivo"

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
' inclus&atilde;o / altera&ccedil;&atilde;o de dados
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
'0 - Sem Acesso, 1 - Só Consulta , 2 - Só Inclui,  3 - Só Altera, 4 - Só Exclui e  5 - Acesso Completo
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

function gera_grid_linha(p_co_mat_prin, p_co_mat_sub, p_unidade, p_curso, p_co_etapa, p_turma)
	conta_linha = 0
	
	if p_co_mat_sub ="" then
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Da_Aula where UCASE(CO_Materia_Principal)='"& p_co_mat_prin &"' AND UCASE(CO_Materia)='"& p_co_mat_prin &"' AND NU_Unidade="& p_unidade &" AND CO_Curso='"& p_curso &"' AND CO_Etapa='"& p_co_etapa &"' AND CO_Turma='"& p_turma &"'"
		'	response.Write(SQL2&"<BR>")
		RS2.Open SQL2, CON2		
		
		wrk_co_materia = p_co_mat_prin
	else
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Da_Aula_Subs where UCASE(CO_Materia_Principal)='"& p_co_mat_prin &"' AND UCASE(CO_Materia)='"& p_co_mat_sub &"' AND NU_Unidade="& p_unidade &" AND CO_Curso='"& p_curso &"' AND CO_Etapa='"& p_co_etapa &"' AND CO_Turma='"& p_turma &"'"
		'	response.Write(SQL2&"<BR>")
		RS2.Open SQL2, CON2		
		wrk_co_materia = p_co_mat_sub		
	end if	
	
	while not RS2.EOF			

		co_professor = RS2("CO_Professor")
		co_cordenador= RS2("CO_Cord")
		p1 = RS2("ST_Per_1")
		p2 = RS2("ST_Per_2")
		p3 = RS2("ST_Per_3")
		p4 = RS2("ST_Per_4")				
		p5 = RS2("ST_Per_5")		
		
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Professor where CO_Professor="& co_professor
		RS1.Open SQL1, CON1
			
		if RS1.EOF then		
			no_prof = "nome em branco<br>"	
			ativo = "FALSO"			
		ELSE							
			no_prof= RS1("NO_Professor")
			if RS1("IN_Ativo_Escola") then
				ativo = "VERDADEIRO"				
			else
				ativo = "FALSO"				
			end if	
		END IF	
	'response.Write(no_prof&"<BR>")	
	Set RS8 = Server.CreateObject("ADODB.Recordset")
	SQL8 = "SELECT * FROM TB_Materia where CO_Materia='"& wrk_co_materia &"'"
	RS8.Open SQL8, CON0

	no_mat= RS8("NO_Materia")		
		
				
		if conta_linha=0 then
			vetor_linha = wrk_co_materia&"#$#"&no_mat&"#!#"&co_professor&"#$#"&no_prof&"#$#"&ativo&"#!#"&p1&"#!#"&p2&"#!#"&p3&"#!#"&p4&"#!#"&p5		
		else		
			vetor_linha = vetor_grid_linha&"$!$"&wrk_co_materia&"#$#"&no_mat&"#!#"&co_professor&"#$#"&no_prof&"#$#"&ativo&"#!#"&p1&"#!#"&p2&"#!#"&p3&"#!#"&p4&"#!#"&p5
		end if
		
		conta_linha=conta_linha+1	
						
	RS2.MOVENEXT
	WEND	
		
	gera_grid_linha = vetor_linha
end function	
%>
