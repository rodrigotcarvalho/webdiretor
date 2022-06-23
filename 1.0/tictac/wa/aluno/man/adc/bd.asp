<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<%
		Set CON5 = Server.CreateObject("ADODB.Connection") 
		ABRIR5 = "DBQ="& CAMINHO_ei & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON5.Open ABRIR5
		

	cod_cons  = request.form("cod_cons") 
	dia_de = request.form("dia_de") 
	mes_de = request.form("mes_de") 
	ano_de = request.form("ano_de") 		
	dat_adapta = dia_de&"/"&mes_de&"/"&ano_de
	irmao1 = request.form("irmao1") 
	idade1 = request.form("idade1") 
	irmao2 = request.form("irmao2") 
	idade2 = request.form("idade2") 
	irmao3 = request.form("irmao3") 
	idade3 = request.form("idade3") 
	outros = request.form("outros") 
	desejada = request.form("desejada") 
	esperada = request.form("esperada") 
	como_passou = request.form("como_passou") 
	normal = request.form("normal") 
	termo = request.form("termo") 
	prematuro = request.form("prematuro") 	
	cesariana = request.form("cesariana") 
	dia_nascimento = request.form("dia_nascimento") 
	materna = request.form("materna") 
	pegou_bem = request.form("pegou_bem") 
	artificial = request.form("artificial") 
	adaptacao_mudanca = request.form("adaptacao_mudanca") 
	chupava_dedo = request.form("chupava_dedo") 
	chupeta = request.form("chupeta") 
	alimentacao = request.form("alimentacao") 	
	dificuldade_alimentacao = request.form("dificuldade_alimentacao") 
	sentou = request.form("sentou") 
	arrastou = request.form("arrastou") 
	engatinhou = request.form("engatinhou") 
	andou = request.form("andou") 
	linguagem = request.form("linguagem") 
	dificuldade_fala = request.form("dificuldade_fala") 
	pedalar = request.form("pedalar") 
	infeccoes = request.form("infeccoes") 
	alergias = request.form("alergias") 
	outras_infeccoes = request.form("outras_infeccoes") 
	antitermico = request.form("antitermico") 
	antecedentes = request.form("antecedentes") 
	divertimentos = request.form("divertimentos") 
	higiene = request.form("higiene") 
	controle = request.form("controle") 
	sono = request.form("sono") 
	gosta_fazer = request.form("gosta_fazer") 
	caracteristicas = request.form("caracteristicas") 
	co_user = session("co_user")


	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Entrevistas_Inicial WHERE CO_Matricula = "&cod_cons

	RS.Open SQL, CON5

check = 2
ordem_original=1

	if RS.EOF	then

		Set RS = server.createobject("adodb.recordset")
		RS.open "TB_Entrevistas_Inicial", CON5, 2, 2 'which table do you want open
		RS.addnew
			RS("CO_Matricula") = cod_cons
			RS("DA_Adapta")  = dat_adapta
			RS("NO_Irmao1") = irmao1
			RS("ID_Irmao1") = idade1
			RS("NO_Irmao2") = irmao2
			RS("ID_Irmao2") = idade2
			RS("NO_Irmao3") = irmao3
			RS("ID_Irmao3") = idade3
			RS("TX_Outras_Pessoas") = outros
			RS("TX_ISC_Desejada") = desejada
			RS("TX_ISC_Esperada") = esperada
			RS("TX_ISC_Como_grav") = como_passou
			RS("TX_ISC_Normal") = normal
			RS("TX_ISC_Termo") = termo
			RS("TX_ISC_Prema") = prematuro
			RS("TX_ISC_Cesariana") = cesariana
			RS("TX_ISC_Como_Parto") = dia_nascimento
			RS("TX_ISC_Materna") = materna
			RS("TX_ISC_Pegou") = pegou_bem
			RS("TX_ISC_Artificial") = artificial
			RS("TX_ISC_Como_mud") = adaptacao_mudanca
			RS("TX_ISC_chupava") = chupava_dedo
			RS("TX_ISC_chupeta") = chupeta
			RS("TX_ISC_alim") = alimentacao
			RS("TX_ISC_Como_alim") = dificuldade_alimentacao	
			RS("TX_DP_Sentou") = sentou
			RS("TX_DP_Arrastou") = arrastou
			RS("TX_DP_Enga") = engatinhou
			RS("TX_DP_Andou") = andou
			RS("TX_DP_Ling") = linguagem
			RS("TX_DP_Obs") = dificuldade_fala
			RS("TX_DP_Anda_bem") = pedalar
			RS("TX_AP_Infec") = infeccoes
			RS("TX_AP_alergia") = alergias
			RS("TX_AP_outros") = outras_infeccoes
			RS("TX_AP_Antit") = antitermico
			RS("TX_AP_Antece") = antecedentes
			RS("TX_DF") = divertimentos
			RS("TX_AH_Hig") = higiene
			RS("TX_AH_Como") = controle
			RS("TX_AH_Sono") = sono 
			RS("TX_IN_Sob") = gosta_fazer
			RS("TX_IN_Carac") = caracteristicas
			RS("CO_Usuario") = co_user
		RS.update
		outro= "Incluir "&cod_cons	
		opt = "i"					
	else
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		CONEXAO0 = "DELETE * from TB_Entrevistas_Inicial WHERE CO_Matricula = "& cod_cons
		Set RS0 = CON5.Execute(CONEXAO0)
		Set RS = server.createobject("adodb.recordset")
		
		Set RS = server.createobject("adodb.recordset")
		RS.open "TB_Entrevistas_Inicial", CON5, 2, 2 'which table do you want open
		RS.addnew
			RS("CO_Matricula") = cod_cons
			RS("DA_Adapta")  = dat_adapta
			RS("NO_Irmao1") = irmao1
			RS("ID_Irmao1") = idade1
			RS("NO_Irmao2") = irmao2
			RS("ID_Irmao2") = idade2
			RS("NO_Irmao3") = irmao3
			RS("ID_Irmao3") = idade3
			RS("TX_Outras_Pessoas") = outros
			RS("TX_ISC_Desejada") = desejada
			RS("TX_ISC_Esperada") = esperada
			RS("TX_ISC_Como_grav") = como_passou
			RS("TX_ISC_Normal") = normal
			RS("TX_ISC_Termo") = termo
			RS("TX_ISC_Prema") = prematuro
			RS("TX_ISC_Cesariana") = cesariana
			RS("TX_ISC_Como_Parto") = dia_nascimento
			RS("TX_ISC_Materna") = materna
			RS("TX_ISC_Pegou") = pegou_bem
			RS("TX_ISC_Artificial") = artificial
			RS("TX_ISC_Como_mud") = adaptacao_mudanca
			RS("TX_ISC_chupava") = chupava_dedo
			RS("TX_ISC_chupeta") = chupeta
			RS("TX_ISC_alim") = alimentacao
			RS("TX_ISC_Como_alim") = dificuldade_alimentacao	
			RS("TX_DP_Sentou") = sentou
			RS("TX_DP_Arrastou") = arrastou
			RS("TX_DP_Enga") = engatinhou
			RS("TX_DP_Andou") = andou
			RS("TX_DP_Ling") = linguagem
			RS("TX_DP_Obs") = dificuldade_fala
			RS("TX_DP_Anda_bem") = pedalar
			RS("TX_AP_Infec") = infeccoes
			RS("TX_AP_alergia") = alergias
			RS("TX_AP_outros") = outras_infeccoes
			RS("TX_AP_Antit") = antitermico
			RS("TX_AP_Antece") = antecedentes
			RS("TX_DF") = divertimentos
			RS("TX_AH_Hig") = higiene
			RS("TX_AH_Como") = controle
			RS("TX_AH_Sono") = sono 
			RS("TX_IN_Sob") = gosta_fazer
			RS("TX_IN_Carac") = caracteristicas
			RS("CO_Usuario") = co_user
		RS.update
		outro= "Alterar "&cod_cons	
		opt = "a"	
	end if			
call GravaLog (chave,outro)
response.redirect("altera.asp?ori="&opt&"&cod_cons="&cod_cons&"&res=ok")
%>