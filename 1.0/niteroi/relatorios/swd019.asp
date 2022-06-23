<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<%
obr=request.QueryString("obr")
dados_informados = split(obr, "$!$")
ano_letivo = session("ano_letivo")
'Anamnese
arquivo = "SWD019"

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

if mes<10 then
meswrt="0"&mes
else
meswrt=mes
end if
if min<10 then
minwrt="0"&min
else
minwrt=min
end if

data = dia &"/"& meswrt &"/"& ano
horario = hora & ":"& minwrt	
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON5 = Server.CreateObject("ADODB.Connection") 
		ABRIR5 = "DBQ="& CAMINHO_ei & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON5.Open ABRIR5			
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0

	SET Pdf = Server.CreateObject("Persits.Pdf")
	SET Doc = Pdf.CreateDocument
	Set Logo1 = Doc.OpenImage( Server.MapPath( "../img/logo_niteroi_preto.gif") )
	Set Logo2 = Doc.OpenImage( Server.MapPath( "../img/logo_arariboia_preto.gif") )	
	Set Font = Doc.Fonts.LoadFromFile(Server.MapPath("../fonts/arial.ttf"))	
	Set Font_Tesoura = Doc.Fonts.LoadFromFile(Server.MapPath("../fonts/ZapfDingbats.ttf"))
	If Font.Embedding = 2 Then
	   Response.Write "Embedding of this font is prohibited."
	   Set Font = Nothing
	End If


if ubound(dados_informados) = 0 then
	origem="../wa/aluno/man/adc/"
	
	cod_cons = dados_informados(0)

	Set RS1 = Server.CreateObject("ADODB.Recordset")
	SQL1 = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod_cons
	RS1.Open SQL1, CON1
	
	if RS1.EOF then
		response.redirect(origem&"index.asp?nvg="&nvg&"&opt=err4")
	end if

	If Not IsArray(alunos_encontrados) Then alunos_encontrados = Array() End if	
	ReDim preserve alunos_encontrados(UBound(alunos_encontrados)+1)	
	alunos_encontrados(Ubound(alunos_encontrados)) = cod_cons	



else
	origem=""
	unidade=dados_informados(0)
	curso=dados_informados(1)
	co_etapa=dados_informados(2)
	turma=dados_informados(3)
	
	if unidade="999990" or unidade="" or isnull(unidade) then
		SQL_BUSCA_ALUNOS="NULO"
	else	
		SQL_ALUNOS= "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.CO_Situacao='C' AND TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.NU_Unidade = "& unidade		
		if curso="999990" or curso="" or isnull(curso) then
			SQL_CURSO=""
		else
			SQL_CURSO=" AND TB_Matriculas.CO_Curso = '"& curso &"'"			
		end if
	
		if co_etapa="999990" or co_etapa="" or isnull(co_etapa) then
			SQL_ETAPA=""		
		else
			SQL_ETAPA=" AND TB_Matriculas.CO_Etapa = '"& co_etapa &"'"				
		end if
	
		if turma="999990" or turma="" or isnull(turma) then
			SQL_TURMA=""		
		else
			SQL_TURMA=" AND TB_Matriculas.CO_Turma = '"& turma &"' "			
		end if
	
		SQL_BUSCA_ALUNOS= SQL_ALUNOS&SQL_CURSO&SQL_ETAPA&SQL_TURMA&" order by TB_Matriculas.NU_Unidade,TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma, TB_Alunos.NO_Aluno"
	end if

	if SQL_BUSCA_ALUNOS="NULO" then
	else
	
	nu_chamada_check = 1
		Set RSA = Server.CreateObject("ADODB.Recordset")
		CONEXAOA = SQL_BUSCA_ALUNOS
		Set RSA = CON1.Execute(CONEXAOA)
		vetor_matriculas="" 
		While Not RSA.EOF
			nu_matricula = RSA("CO_Matricula")
			nu_chamada = RSA("NU_Chamada")
			if nu_chamada_check = 1 and nu_chamada=nu_chamada_check then
				vetor_matriculas=nu_matricula
			elseif nu_chamada_check = 1 then
				while nu_chamada_check < nu_chamada
					nu_chamada_check=nu_chamada_check+1
				wend 
				vetor_matriculas=nu_matricula
			else
				vetor_matriculas=vetor_matriculas&"#!#"&nu_matricula
			end if	
		nu_chamada_check=nu_chamada_check+1	
		RSA.MoveNext
		Wend 
		RSA.Close
		Set RSA = Nothing	
	end if	

	if vetor_matriculas="" then
		alunos_encontrados = Array() 
	else
		alunos_encontrados = split(vetor_matriculas, "#!#" )	
	end if	
end if	

if ubound(alunos_encontrados)=-1 then
	response.redirect(origem&"index.asp?nvg="&nvg&"&opt=err1")	
else

	relatorios_gerados=-1			
	For i=0 to ubound(alunos_encontrados)	
		SET Page = Doc.Pages.Add(595,842)
		y_posicao=842
		total_linhas=0		
		max_linhas=34
		qtd_turma=0
		paginacao=1
						
		num_matric=alunos_encontrados(i)		

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos,TB_Matriculas WHERE TB_Alunos.CO_Matricula = TB_Matriculas.CO_Matricula AND NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Matricula ="& num_matric
		RS.Open SQL, CON1
	
		nome_aluno = RS("NO_Aluno")
		sexo_aluno = RS("IN_Sexo")		
		ano_aluno = RS("NU_Ano")
		rematricula = RS("DA_Rematricula")
		situacao = RS("CO_Situacao")
		encerramento= RS("DA_Encerramento")
		unidade= RS("NU_Unidade")
		curso= RS("CO_Curso")
		etapa= RS("CO_Etapa")
		turma= RS("CO_Turma")
		cham= RS("NU_Chamada")
		
		if sexo_aluno="F" then
			desinencia="a"
		else
			desinencia="o"
		end if
		
		no_unidade= GeraNomes("U",unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)
		no_curso=GeraNomes("C",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro) 	
		no_etapa=GeraNomes("E",curso,etapa,variavel3,variavel4,variavel5,CON0,outro) 	
		no_situacao=GeraNomes("SA",situacao,variavel2,variavel3,variavel4,variavel5,CON0,outro) 		
		co_concordancia_curso=GeraNomes("PC",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro) 	
		no_abrv_curso	= GeraNomes("CA",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro) 				
		
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Unidade WHERE NU_Unidade="& unidade
		RS2.Open SQL2, CON0
							
		rua_unidade = RS2("NO_Logradouro")		
		numero_unidade = RS2("NU_Logradouro")	
		complemento_unidade = RS2("TX_Complemento_Logradouro")	
		cep_unidade = RS2("CO_CEP")	
		bairro_unidade = RS2("CO_Bairro")	
		municipio_unidade = RS2("CO_Municipio")			
		uf_unidade = RS2("SG_UF")			


		if numero_unidade="" or isnull(numero_unidade)then
		else
			numero_unidade=" N&ordm; "&numero_unidade
		end if
			
		if complemento_unidade="" or isnull(complemento_unidade)then
		else
			complemento_unidade=" - "&complemento_unidade
		end if
		
		if cep_unidade="" or isnull(cep_unidade)then
		else
			cep_unidade=" - "&LEFT(cep_unidade,5)&"-"&RIGHT(cep_unidade,3)
		end if
		
		if uf_unidade="" or isnull(uf_unidade)then
		else
			uf_unidade_municipio=uf_unidade
			uf_unidade=" - "&uf_unidade
		end if
		
		if municipio_unidade="" or isnull(municipio_unidade) or uf_unidade_municipio="" or isnull(uf_unidade_municipio)then
		else		
			no_municipio = GeraNomes("Mun",uf_unidade_municipio,municipio_unidade,variavel3,variavel4,variavel5,CON0,outro) 
			
			if bairro_unidade="" or isnull(bairro_unidade)then
			else
				no_bairro = GeraNomes("Bai",uf_unidade_municipio,municipio_unidade,bairro_unidade,variavel4,variavel5,CON0,outro)
				bairro_unidade=" - "&no_bairro
			end if							
		end if
		endereco_unidade=rua_unidade&numero_unidade&complemento_unidade&bairro_unidade&cep_unidade&" - "&no_municipio&uf_unidade					
		
		
		no_curso= no_etapa&"&nbsp;"&co_concordancia_curso&"&nbsp;"&no_curso
		texto_turma = "&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;<b>Turma</b> "&turma
		mensagem_cabecalho=ano_letivo&"&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;"&no_curso&texto_turma
		
		
		
		
'CABEÇALHO==========================================================================================		
		Set Param_Logo_Gde = Pdf.CreateParam
		margem=25			
		linha=10		
		unidade = unidade*1	
		if unidade = 1 then
			largura_logo_gde=formatnumber(Logo1.Width*0.6,0)
			altura_logo_gde=formatnumber(Logo1.Height*0.6,0)
			area_utilizavel=Page.Width-(margem*2)
			Param_Logo_Gde("x") = margem
			Param_Logo_Gde("y") = Page.Height - altura_logo_gde -margem
			Param_Logo_Gde("ScaleX") = 0.6
			Param_Logo_Gde("ScaleY") = 0.6
			Page.Canvas.DrawImage Logo1, Param_Logo_Gde
		else
			largura_logo_gde=formatnumber(Logo2.Width*0.5,0)
			altura_logo_gde=formatnumber(Logo2.Height*0.5,0)
			area_utilizavel=Page.Width-(margem*2)
			Param_Logo_Gde("x") = margem
			Param_Logo_Gde("y") = Page.Height - altura_logo_gde -margem
			Param_Logo_Gde("ScaleX") = 0.5
			Param_Logo_Gde("ScaleY") = 0.5
			Page.Canvas.DrawImage Logo2, Param_Logo_Gde		
		end if

		x_texto=largura_logo_gde+ margem+10
		y_texto=formatnumber(Page.Height - margem,0)
		width_texto=Page.Width -largura_logo_gde - 80

	
		SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
		Text = "<p><i><b>"&UCASE(no_unidade)&"</b></i></p><br><font style=""font-size:10pt;"">"&endereco_unidade&"</FONT>" 			
		
		Do While Len(Text) > 0
			CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
		 
			If CharsPrinted = Len(Text) Then Exit Do
				SET Page = Page.NextPage
			Text = Right( Text, Len(Text) - CharsPrinted)
		Loop 

		y_texto=y_texto-altura_logo_gde+10
		SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height=30; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
		Text = "<center><i><b><font style=""font-size:18pt;"">ANAMNESE</font></b></i></center>"
		
		
		Do While Len(Text) > 0
			CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
		 
			If CharsPrinted = Len(Text) Then Exit Do
				SET Page = Page.NextPage
			Text = Right( Text, Len(Text) - CharsPrinted)
		Loop 

		
		Page.Canvas.SetParams "LineWidth=1" 
		Page.Canvas.SetParams "LineCap=0" 
		inicio_primeiro_separador=largura_logo_gde+margem+10
		altura_primeiro_separador= Page.Height - margem - 17
		With Page.Canvas
		   .MoveTo inicio_primeiro_separador, altura_primeiro_separador
		   .LineTo area_utilizavel+margem, altura_primeiro_separador
		   .Stroke
		End With 	


		Page.Canvas.SetParams "LineWidth=1" 
		Page.Canvas.SetParams "LineCap=0" 
		altura_segundo_separador= Page.Height - altura_logo_gde-margem - 20
		With Page.Canvas
		   .MoveTo margem, altura_segundo_separador
		   .LineTo area_utilizavel+margem, altura_segundo_separador
		   .Stroke
		End With 	

'		Set param_table1 = Pdf.CreateParam("width="&area_utilizavel&"; height=20; rows=1; cols=3; border=0; cellborder=0; cellspacing=0;")
'		Set Table = Doc.CreateTable(param_table1)
'		Table.Font = Font
'		y_primeira_tabela=altura_segundo_separador-10
'		x_primeira_tabela=margem+5
'		With Table.Rows(1)
'		   .Cells(1).Width = 50			   		   		   
'		   .Cells(2).Width = area_utilizavel-100
'		   .Cells(3).Width = 50	
'		End With
'		
'		Table(1, 1).AddText "<b>Ano Letivo:</b>", "size=8;html=true", Font 
'		Table(1, 2).AddText mensagem_cabecalho, "size=8;html=true", Font	
'		'Table(1, 3).AddText "<div align=""right""><b>Legenda:</b> Md=M&eacute;dia - Res=Resultado&nbsp;&nbsp;&nbsp;</div>", "size=8;html=true", Font	
'		Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 		

		
		y_nome_aluno=Page.Height - altura_logo_gde-46
		width_nome_aluno=Page.Width - margem
		
		SET Param_Nome_Aluno = Pdf.CreateParam("x="&margem&";y="&y_nome_aluno&"; height=50; width="&width_nome_aluno&"; alignment=left; html=True")
		Nome = "<font style=""font-size:11pt;""><b>Alun"&desinencia&": "&nome_aluno&"</b></font>"
		

		Do While Len(Nome) > 0
			CharsPrinted = Page.Canvas.DrawText(Nome, Param_Nome_Aluno, Font )
		 
			If CharsPrinted = Len(Nome) Then Exit Do
				SET Page = Page.NextPage
			Nome = Right( Nome, Len(Nome) - CharsPrinted)
		Loop 
		
		Page.Canvas.SetParams "LineWidth=2" 
		Page.Canvas.SetParams "LineCap=0" 
		With Page.Canvas
		   .MoveTo margem, Page.Height - altura_logo_gde-65
		   .LineTo Page.Width - margem, Page.Height - altura_logo_gde-65
		   .Stroke
		End With 	


		Set param_table1 = Pdf.CreateParam("width=533; height=25; rows=2; cols=8; border=0; cellborder=0; cellspacing=0;")
		Set Table = Doc.CreateTable(param_table1)
		Table.Font = Font
		y_table=Page.Height - altura_logo_gde-70
		
		With Table.Rows(1)
		   .Cells(1).Width = 40
		   .Cells(2).Width = 200
		   .Cells(3).Width = 25
		   .Cells(4).Width = 70
		   .Cells(5).Width = 60
		   .Cells(6).Width = 38
		   .Cells(7).Width = 50
		   .Cells(8).Width = 50      
		End With
		Table(1, 2).ColSpan = 5
		Table(1, 1).AddText "Unidade:", "size=9;", Font 
		Table(2, 1).AddText "Curso:", "size=9;", Font 
		Table(1, 2).AddText no_unidade, "size=9;", Font 
		Table(2, 2).ColSpan = 2
		Table(2, 2).AddText no_curso, "size=9; html=true", Font 
		'Table(2, 3).AddText no_etapa, "size=9;", Font 
		Table(2, 4).AddText "Turma: "&turma, "size=9;", Font 
		Table(2, 5).AddText "N&ordm;. Chamada: "&cham, "size=9; html=true", Font 
		Table(2, 6).AddText cham, "size=9;", Font 
		Table(1, 7).AddText "<div align=""right"">Matr&iacute;cula: </div>", "size=9; html=true", Font 
		Table(1, 8).AddText cod_cons, "size=9;alignment=right", Font 
		Table(2, 7).AddText "Ano Letivo: ", "size=9; alignment=right", Font 
		Table(2, 8).AddText ano_letivo, "size=9;alignment=right", Font 
		Page.Canvas.DrawTable Table, "x="&margem&", y="&y_table&"" 
	
		
		With Page.Canvas
		   .MoveTo margem, Page.Height - altura_logo_gde-100
		   .LineTo Page.Width - margem, Page.Height - altura_logo_gde-100
		   .Stroke
		End With 		
'================================================================================================================					
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Entrevistas_Inicial WHERE CO_Matricula = "&num_matric
		RS.Open SQL, CON5
	
		check = 2
		ordem_original=1
		
		if RS.EOF	then		
			dat_adapta = dd&"/"&mm&"/"&aa
			irmao1 = "" 
			idade1 = "" 
			irmao2 = "" 
			idade2 = ""
			irmao3 = ""
			idade3 = "" 
			outros = ""
			desejada = ""
			esperada = ""
			como_passou = ""
			normal = ""
			termo = ""
			prematuro = ""
			cesariana = ""
			dia_nascimento = ""
			materna = ""
			pegou_bem = ""
			artificial = "" 
			adaptacao_mudanca = "" 
			chupava_dedo = ""
			chupeta = ""
			alimentacao = ""	
			dificuldade_alimentacao = ""
			sentou = ""
			arrastou = "" 
			engatinhou = ""
			andou = ""
			linguagem = ""
			dificuldade_fala = "" 
			pedalar = ""
			infeccoes = "" 
			alergias = ""
			outras_infeccoes = ""
			antitermico = ""
			antecedentes = ""
			divertimentos = ""
			higiene = ""
			controle = "" 
			sono = ""
			gosta_fazer = ""
			caracteristicas = ""
		
		else			
			dat_adapta = RS("DA_Adapta") 
			irmao1 = RS("NO_Irmao1") 
			idade1 = RS("ID_Irmao1") 
			irmao2 = RS("NO_Irmao2") 
			idade2 = RS("ID_Irmao2") 
			irmao3 = RS("NO_Irmao3") 
			idade3 = RS("ID_Irmao3") 
			outros = RS("TX_Outras_Pessoas") 
			desejada = RS("TX_ISC_Desejada") 
			esperada = RS("TX_ISC_Esperada") 
			como_passou = RS("TX_ISC_Como_grav") 
			normal = RS("TX_ISC_Normal") 
			termo = RS("TX_ISC_Termo") 
			prematuro = RS("TX_ISC_Prema") 	
			cesariana = RS("TX_ISC_Cesariana") 
			dia_nascimento = RS("TX_ISC_Como_Parto") 
			materna = RS("TX_ISC_Materna") 
			pegou_bem = RS("TX_ISC_Pegou") 
			artificial = RS("TX_ISC_Artificial") 
			adaptacao_mudanca = RS("TX_ISC_Como_mud") 
			chupava_dedo = RS("TX_ISC_chupava") 
			chupeta = RS("TX_ISC_chupeta") 
			alimentacao = RS("TX_ISC_alim") 
			dificuldade_alimentacao = RS("TX_ISC_Como_alim") 	
			sentou = RS("TX_DP_Sentou") 
			arrastou = RS("TX_DP_Arrastou") 
			engatinhou = RS("TX_DP_Enga") 
			andou = RS("TX_DP_Andou") 
			linguagem = RS("TX_DP_Ling") 
			dificuldade_fala = RS("TX_DP_Obs") 
			pedalar = RS("TX_DP_Anda_bem") 
			infeccoes = RS("TX_AP_Infec") 
			alergias = RS("TX_AP_alergia") 
			outras_infeccoes = RS("TX_AP_outros") 
			antitermico = RS("TX_AP_Antit") 
			antecedentes = RS("TX_AP_Antece") 
			divertimentos = RS("TX_DF") 
			higiene = RS("TX_AH_Hig") 
			controle = RS("TX_AH_Como") 
			sono = RS("TX_AH_Sono") 
			gosta_fazer = RS("TX_IN_Sob") 
			caracteristicas = RS("TX_IN_Carac") 
			'co_user_bd = RS("CO_Usuario")		
		end if

		IF dat_adapta="" or isnull(dat_adapta) or dat_adapta="//" then
		else		
			dt_adpt = split(dat_adapta, "/")
			dd = dt_adpt(0)
			mm = dt_adpt(1)
			aa = dt_adpt(2)		
	
			if dd<10 then
				dia_txt="0"&dd
			else	
				dia_txt=dd		
			end if		
			
			Select case mm
			
				case 1
				mes_txt="janeiro"
				
				case 2
				mes_txt="fevereiro"
				
				case 3
				mes_txt="mar&ccedil;o"
				
				case 4
				mes_txt="abril"
				
				case 5
				mes_txt="maio"
		
				case 6
				mes_txt="junho"
				
				case 7
				mes_txt="julho"
				
				case 8
				mes_txt="agosto"		
				
				case 9
				mes_txt="setembro"
		
				case 10
				mes_txt="outubro"
				
				case 11
				mes_txt="novembro"
				
				case 12
				mes_txt="dezembro"				
			end select		
			
			dat_adapta_txt = dia_txt&" de "&mes_txt&" de "&aa		
		end if	
	
		altura_medias=1040
		largura_tabela=680
		y_table=y_table-50		
		Set param_table = Pdf.CreateParam("width="&largura_tabela&"; height="&altura_medias&"; rows=26; cols=4; border=1; cellborder=0.1; cellspacing=0;  x="&margem&"; y="&y_table&";")
		Set Entrevista = Doc.CreateTable(param_table)
		'param_table.Set "" 			
		Entrevista.Font = Font
		With Entrevista.Rows(1)
		   .Cells(1).Width = 100			
		   .Cells(3).Width = 100			     		         			   			         
		End With
		Entrevista.Rows(1).Cells(1).Height = 13
		Entrevista.Rows(2).Cells(1).Height = 17
		Entrevista.Rows(3).Cells(1).Height = 13
		Entrevista.Rows(4).Cells(1).Height = 13
		Entrevista.Rows(5).Cells(1).Height = 13
		Entrevista.Rows(6).Cells(1).Height = 17
		Entrevista.Rows(7).Cells(1).Height = 30
		Entrevista.Rows(8).Cells(1).Height = 17																
		Entrevista.Rows(9).Cells(1).Height = 39
		Entrevista.Rows(10).Cells(1).Height = 82		
		'Entrevista.Rows(11).Cells(1).Height = 120	
		Entrevista.Rows(11).Cells(1).Height = 151	
		Entrevista.Rows(12).Cells(1).Height = 17
		Entrevista.Rows(13).Cells(1).Height = 13
		Entrevista.Rows(14).Cells(1).Height = 13
		Entrevista.Rows(15).Cells(1).Height = 13
		Entrevista.Rows(16).Cells(1).Height = 13
		Entrevista.Rows(17).Cells(1).Height = 13
		Entrevista.Rows(18).Cells(1).Height = 13
		Entrevista.Rows(19).Cells(1).Height = 40		
		Entrevista.Rows(20).Cells(1).Height = 13	
		Entrevista.Rows(21).Cells(1).Height = 13	
		Entrevista.Rows(22).Cells(1).Height = 17	
		Entrevista.Rows(23).Cells(1).Height = 13
		Entrevista.Rows(24).Cells(1).Height = 13
		Entrevista.Rows(25).Cells(1).Height = 13
		Entrevista.Rows(26).Cells(1).Height = 13
'		Entrevista.Rows(27).Cells(1).Height = 13	
'		Entrevista.Rows(28).Cells(1).Height = 30
'		Entrevista.Rows(29).Cells(1).Height = 15	
'		Entrevista.Rows(30).Cells(1).Height = 17
'		Entrevista.Rows(31).Cells(1).Height = 13
'		Entrevista.Rows(32).Cells(1).Height = 17
'		Entrevista.Rows(33).Cells(1).Height = 13			
'		Entrevista.Rows(34).Cells(1).Height = 17
'		Entrevista.Rows(35).Cells(1).Height = 13	
'		Entrevista.Rows(36).Cells(1).Height = 13	
'		Entrevista.Rows(37).Cells(1).Height = 13	
'		Entrevista.Rows(38).Cells(1).Height = 13								
'		Entrevista.Rows(39).Cells(1).Height = 17
'		Entrevista.Rows(40).Cells(1).Height = 13		
'		Entrevista.Rows(41).Cells(1).Height = 13			
'		Entrevista.Rows(42).Cells(1).Height = 17	
'		Entrevista.Rows(43).Cells(1).Height = 30
'		Entrevista.Rows(44).Cells(1).Height = 13																																																					
		Entrevista(1, 2).ColSpan = 3
		Entrevista(6, 1).ColSpan = 4		
		Entrevista(7, 2).ColSpan = 3		
		Entrevista(8, 1).ColSpan = 4	
		Entrevista(9, 2).ColSpan = 3			
		Entrevista(10, 2).ColSpan = 3			
		Entrevista(11, 2).ColSpan = 3		
		Entrevista(12, 1).ColSpan = 4	
		Entrevista(13, 2).ColSpan = 3
		Entrevista(14, 2).ColSpan = 3	
		Entrevista(15, 2).ColSpan = 3	
		Entrevista(16, 2).ColSpan = 3	
		Entrevista(17, 2).ColSpan = 3	
		Entrevista(18, 1).ColSpan = 4	
		Entrevista(19, 1).ColSpan = 4	
		Entrevista(20, 1).ColSpan = 4	
		Entrevista(21, 1).ColSpan = 4	
		Entrevista(22, 1).ColSpan = 4	
		Entrevista(23, 2).ColSpan = 3	
		Entrevista(24, 2).ColSpan = 3
		Entrevista(25, 2).ColSpan = 3
'		Entrevista(26, 2).ColSpan = 3
'		Entrevista(27, 1).ColSpan = 4		
'		Entrevista(28, 1).ColSpan = 4	
'		Entrevista(28, 1).RowSpan = 2	
'		Entrevista(30, 1).ColSpan = 4	
'		Entrevista(31, 1).ColSpan = 4	
'		Entrevista(32, 1).ColSpan = 4	
'		Entrevista(33, 1).ColSpan = 4	
'		Entrevista(34, 1).ColSpan = 4	
'		Entrevista(35, 2).ColSpan = 3	
'		Entrevista(36, 1).ColSpan = 4	
'		Entrevista(37, 1).ColSpan = 4
'		Entrevista(38, 2).ColSpan = 3
'		Entrevista(39, 1).ColSpan = 4
'		Entrevista(40, 1).ColSpan = 4
'		Entrevista(41, 1).ColSpan = 4
'		Entrevista(42, 1).ColSpan = 4
'		Entrevista(43, 1).ColSpan = 4	
'		Entrevista(44, 1).ColSpan = 4																														
'		Entrevista.At(30, 1).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False"
'		Entrevista.At(31, 1).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False"																																																
		Entrevista(1, 1).AddText "<div align=""right""><b>Data de Adapta&ccedil;&atilde;o:&nbsp;</b></div>", "size=8; indentx=5;  html=true", Font 
		Entrevista(1, 2).AddText dat_adapta_txt, "size=8; indentx=5; html=true", Font 	
		Entrevista(2, 1).ColSpan = 4		
		Entrevista(2, 1).AddText "<div align=""center""><b>Irm&atilde;os</b></div>", "size=9; indenty=2; html=true", Font 			
		Entrevista(3, 1).AddText "<div align=""right""><b>Nome:&nbsp;</b></div>", "size=8; html=true", Font 
		Entrevista(3, 2).AddText irmao1, "size=8; html=true; indentx=2; expand=true", Font 		
		Entrevista(3, 3).AddText "<div align=""right""><b>Idade:&nbsp;</b></div>", "size=8; html=true", Font 
		Entrevista(3, 4).AddText idade1, "size=8; html=true; indentx=2; expand=true", Font 				
		Entrevista(4, 1).AddText "<div align=""right""><b>Nome:&nbsp;</b></div>", "size=8; html=true", Font 
		Entrevista(4, 2).AddText irmao2, "size=8; html=true; indentx=2; expand=true", Font 		
		Entrevista(4, 3).AddText "<div align=""right""><b>Idade:&nbsp;</b></div>", "size=8; html=true", Font 
		Entrevista(4, 4).AddText idade2, "size=8; html=true; indentx=2; expand=true", Font 	
		Entrevista(5, 1).AddText "<div align=""right""><b>Nome:&nbsp;</b></div>", "size=8; html=true", Font 
		Entrevista(5, 2).AddText irmao3, "size=8; html=true; indentx=2; expand=true", Font 		
		Entrevista(5, 3).AddText "<div align=""right""><b>Idade:&nbsp;</b></div>", "size=8; html=true", Font 
		Entrevista(5, 4).AddText idade3, "size=8; html=true; indentx=2; expand=true", Font 			
		Entrevista(6, 1).AddText "<div align=""center""><b>Outras pessoas residindo com a fam&iacute;lia</b></div>", "size=9; indenty=2; html=true", Font 			
		Entrevista(7, 1).AddText "<div align=""right""><b>Nomes:&nbsp;</b></div>", "size=8; html=true", Font 
		Entrevista(7, 2).AddText outros, "size=8; html=true;expand=true", Font 			
		Entrevista(8, 1).AddText "<div align=""center""><b>Informa&ccedil;&otilde;es sobre a crian&ccedil;a</b></div>", "size=9; indenty=2; html=true", Font 			
		Entrevista(9, 1).AddText "<div align=""right""><b>Gravidez:&nbsp;</b></div>", "size=8; html=true", Font 	
				
			Set SmallTable92 = Doc.CreateTable("Height=39; Width=580; cols=2; rows=3; border=0; cellborder=0.1; cellspacing=0;")
			With SmallTable92.Rows(1)
			   .Cells(1).Width = 120	
			   .Cells(2).Width = 460				   				     		         			   			         
			End With		
														
			SmallTable92.At(1, 1).Canvas.DrawText "<b>Desejada</b>", "x=1; y=10, size=8; html=true", Font		
			SmallTable92.At(1, 2).Canvas.DrawText desejada, "x=1; y=10, size=8; html=true", Font					
			SmallTable92.At(2, 1).Canvas.DrawText "<b>Esperada</b>", "x=1; y=10, size=8; html=true", Font		
			SmallTable92.At(2, 2).Canvas.DrawText esperada, "x=1; y=10, size=8; html=true", Font					
			SmallTable92(3, 1).AddText "<div align=""right""><b>Como Passou:&nbsp;</b></div>", "size=8; html=true", Font														
			SmallTable92.At(3, 2).Canvas.DrawText como_passou, "x=1; y=10, size=8; html=true", Font		
															

		Entrevista(10, 1).AddText "<div align=""right""><b>Parto:&nbsp;</b></div>", "size=8; html=true", Font 		
		
			Set SmallTable102 = Doc.CreateTable("Height=82; Width=580; cols=3; rows=5; border=0; cellborder=0.1; cellspacing=0;")	
			With SmallTable102.Rows(1)
			   .Cells(1).Width = 120		
			   .Cells(2).Width = 30		
			   .Cells(3).Width = 430				   		   			     		         			   			         
			End With		
			SmallTable102(1, 2).ColSpan = 2	
			SmallTable102(2, 2).ColSpan = 2	
			SmallTable102(3, 2).ColSpan = 2	
			SmallTable102(4, 2).ColSpan = 2			
			SmallTable102(5, 1).ColSpan = 2	
			SmallTable102.Rows(1).Cells(1).Height = 13		
			SmallTable102.Rows(2).Cells(1).Height = 13		
			SmallTable102.Rows(3).Cells(1).Height = 13		
			SmallTable102.Rows(4).Cells(1).Height = 13																	
			SmallTable102.Rows(5).Cells(1).Height = 30											
			SmallTable102.At(1, 1).Canvas.DrawText "<b>A Termo</b>", "x=1; y=10, size=8; html=true", Font		
			SmallTable102.At(1, 2).Canvas.DrawText termo, "x=1; y=10, size=8; html=true", Font					
			SmallTable102.At(2, 1).Canvas.DrawText "<b>Prematuro</b>", "x=1; y=10, size=8; html=true", Font		
			SmallTable102.At(2, 2).Canvas.DrawText prematuro, "x=1; y=10, size=8; html=true", Font		
			SmallTable102.At(3, 1).Canvas.DrawText "<b>Normal</b>", "x=1; y=10, size=8; html=true", Font		
			SmallTable102.At(3, 2).Canvas.DrawText normal, "x=1; y=10, size=8; html=true", Font					
			SmallTable102.At(4, 1).Canvas.DrawText "<b>Cesariana</b>", "x=1; y=10, size=8; html=true", Font		
			SmallTable102.At(4, 2).Canvas.DrawText cesariana, "x=1; y=10, size=8; html=true", Font		
			SmallTable102(5, 1).AddText "<div align=""right""><b>Como foi o dia do nascimento?&nbsp;</b></div>", "indenty=2; size=8; html=true", Font														
			SmallTable102.At(5, 3).Canvas.DrawText dia_nascimento, "x=1; y=30, size=8; html=true", Font																					

		Entrevista(11, 1).AddText "<div align=""right""><b>Alimenta&ccedil;&atilde;o:&nbsp;</b></div>", "size=8; html=true", Font 

			
			Set SmallTable112 = Doc.CreateTable("Height=151; Width=440; cols=3; rows=9; border=0; cellborder=0.1; cellspacing=0;")	
			With SmallTable112.Rows(1)
			   .Cells(1).Width = 120	
			   .Cells(2).Width = 30		
			   .Cells(3).Width = 290			   				     		         			   			         
			End With			
			SmallTable112(1, 2).ColSpan = 2	
			SmallTable112(2, 2).ColSpan = 2	
			SmallTable112(3, 2).ColSpan = 2	
			SmallTable112(4, 1).ColSpan = 2		
			SmallTable112(5, 1).ColSpan = 2		
			SmallTable112(6, 1).ColSpan = 2		
			SmallTable112(7, 1).ColSpan = 2		
			SmallTable112(8, 1).ColSpan = 3		
			SmallTable112(9, 1).ColSpan = 3																	
			SmallTable112.Rows(1).Cells(1).Height = 13		
			SmallTable112.Rows(2).Cells(1).Height = 13		
			SmallTable112.Rows(3).Cells(1).Height = 13		
			SmallTable112.Rows(4).Cells(1).Height = 30																	
			SmallTable112.Rows(5).Cells(1).Height = 13	
			SmallTable112.Rows(6).Cells(1).Height = 13		
			SmallTable112.Rows(7).Cells(1).Height = 13																	
			SmallTable112.Rows(8).Cells(1).Height = 13				
			SmallTable112.Rows(9).Cells(1).Height = 30									
			SmallTable112.At(1, 1).Canvas.DrawText "<b>Materna</b>", "x=1; y=10, size=8; html=true", Font		
			SmallTable112.At(1, 2).Canvas.DrawText materna, "x=1; y=10, size=8; html=true", Font					
			SmallTable112(2, 1).AddText "<div align=""right""><b>Pegou Bem o Seio?&nbsp;</b></div>", "indenty=2, size=8; html=true", Font		
			SmallTable112.At(2, 2).Canvas.DrawText pegou_bem, "x=1; y=10, size=8; html=true", Font		
			SmallTable112.At(3, 1).Canvas.DrawText "<b>Artificial</b>", "x=1; y=10, size=8; html=true", Font		
			SmallTable112.At(3, 2).Canvas.DrawText artificial, "x=1; y=10, size=8; html=true", Font																																	
			SmallTable112(4, 1).AddText "<div align=""right""><b>Como a aceita&ccedil;&atilde;o da Mudan&ccedil;a?&nbsp;</b></div>", "indenty=2, size=8; html=true", Font										
			SmallTable112.At(4, 3).Canvas.DrawText adaptacao_mudanca, "x=1; y=30, size=8; html=true", Font						
			SmallTable112.At(5, 1).Canvas.DrawText "<b>Chupava dedo:</b>", "x=1; y=10, size=8; html=true", Font		
			SmallTable112.At(5, 3).Canvas.DrawText chupava_dedo, "x=1; y=10, size=8; html=true", Font					
			SmallTable112.At(6, 1).Canvas.DrawText "<b>Chupeta:</b>", "x=1; y=10, size=8; html=true", Font		
			SmallTable112.At(6, 3).Canvas.DrawText chupeta, "x=1; y=10, size=8; html=true", Font		
			SmallTable112.At(7, 1).Canvas.DrawText "<b>Alimenta&ccedil;&atilde;o:</b>", "x=1; y=10, size=8; html=true", Font		
			SmallTable112.At(7, 3).Canvas.DrawText alimentacao, "x=1; y=10, size=8; html=true", Font							
			SmallTable112.At(8, 1).Canvas.DrawText "<b>Como a fam&iacute;lia reage quando h&aacute; dificuldade na alimenta&ccedil;&atilde;o?&nbsp;</b>", "x=1; y=10, size=8; html=true", Font		
			SmallTable112(9, 1).AddText dificuldade_alimentacao, "x=1; y=30, size=8; html=true", Font		
			
		Entrevista(12, 1).AddText "<div align=""center""><b>Desenvolvimento Psicomotor</b></div>", "size=9; indenty=2; html=true", Font 				
								
			Entrevista(13, 1).AddText "<div align=""right""><b>Sentou:&nbsp;</b></div>", "x=1; y=5, size=8; html=true", Font		
			Entrevista(13, 2).AddText sentou, "x=1; y=1, size=8; html=true", Font					
			Entrevista(14, 1).AddText "<div align=""right""><b>Arrastou:&nbsp;</b></div>", "x=1; y=1, size=8; html=true", Font		
			Entrevista(14, 2).AddText arrastou, "x=1; y=1, size=8; html=true", Font		
			Entrevista(15, 1).AddText "<div align=""right""><b>Engatinhou:&nbsp;</b></div>", "x=1; y=1, size=8; html=true", Font		
			Entrevista(15, 2).AddText engatinhou, "x=1; y=1, size=8; html=true", Font					
			Entrevista(16, 1).AddText "<div align=""right""><b>Andou:&nbsp;</b></div>", "x=1; y=1, size=8; html=true", Font		
			Entrevista(16, 2).AddText andou, "x=1; y=1, size=8; html=true", Font						
			Entrevista(17, 1).AddText "<div align=""right""><b>Linguagem:&nbsp;</b></div>", "x=1; y=1, size=8; html=true", Font		
			Entrevista(17, 2).AddText linguagem, "x=1; y=1, size=8; html=true", Font				
			Entrevista(18, 1).AddText "<b>Voc&ecirc;s observam alguma dificuldade na fala? Qual?&nbsp;</b>", "x=1; y=1, size=8; html=true", Font		
			Entrevista(19, 1).AddText dificuldade_fala, "x=1; y=1, size=8; html=true", Font	
			Entrevista(20, 1).AddText "<b>Anda bem em brinquedos que precise pedalar?&nbsp;</b>", "x=1; y=1, size=8; html=true", Font		
			Entrevista(21, 1).AddText pedalar, "x=1; y=1, size=8; html=true", Font							
			
		Entrevista(22, 1).AddText "<div align=""center""><b>Antecedentes Patol&oacute;gicos</b></div>", "size=9; indenty=2; html=true", Font 		
		Entrevista(23, 1).AddText "<div align=""right""><b>Infec&ccedil;&otilde;es:&nbsp;</b></div>", "x=1; y=1, size=8; html=true", Font		
		Entrevista(23, 2).AddText infeccoes, "x=1; y=1, size=8; html=true", Font	
		Entrevista(24, 1).AddText "<div align=""right""><b>Alergias:&nbsp;</b></div>", "x=1; y=1, size=8; html=true", Font		
		Entrevista(24, 2).AddText alergias, "x=1; y=1, size=8; html=true", Font		
		Entrevista(25, 1).AddText "<div align=""right""><b>Outros:&nbsp;</b></div>", "x=1; y=1, size=8; html=true", Font		
		Entrevista(25, 2).AddText outras_infeccoes, "x=1; y=1, size=8; html=true", Font	
		Entrevista(26, 1).AddText "<div align=""right""><b>Antit&eacute;rmico:&nbsp;</b></div>", "x=1; y=1, size=8; html=true", Font		
		Entrevista(26, 2).AddText antitermico, "x=1; y=1, size=8; html=true", Font	
		
		Page.Canvas.DrawTable Entrevista, "x="&margem&", y="&y_table&"" 
		
		
		
		
		
		
		
		
		
'		Entrevista(27, 1).AddText "<b>Antecedentes familiares:&nbsp;</b>", "x=1; y=1, size=8; html=true", Font		
'		Entrevista(28, 1).AddText antecedentes, "x=1; y=1, size=8; html=true", Font	
'											
'								
'			
'		Entrevista(32, 1).AddText "<div align=""center""><b>Divertimentos da Fam&iacute;lia</b></div>", "size=9; indenty=2; html=true", Font 				
'		Entrevista(33, 1).AddText divertimentos, "size=8; html=true; indentx=2; expand=true", Font 						
'																											
'		Entrevista(34, 1).AddText "<div align=""center""><b>Aquisi&ccedil;&atilde;o de h&aacute;bitos</b></div>", "size=9; indenty=2; html=true", Font 	
'		Entrevista(35, 1).AddText "<div align=""right""><b>Higiene (banho etc):&nbsp;</b></div>", "x=1; y=1, size=8; html=true", Font		
'		Entrevista(35, 2).AddText higiene, "x=1; y=1, size=8; html=true", Font															
'		Entrevista(36, 1).AddText "<b>Como e quando foi treinado (controle dos esfincteres)?</b>", "x=1; y=1, size=8; html=true", Font		
'		Entrevista(37, 1).AddText controle, "x=1; y=1, indentx=2; size=8; html=true", Font		
'		Entrevista(38, 1).AddText "<div align=""right""><b>Sono:&nbsp;</b></div>", "x=1; y=1, size=8; html=true", Font		
'		Entrevista(38, 2).AddText sono, "x=1; y=1, size=8; html=true", Font							
'		Entrevista(39, 1).AddText "<div align=""center""><b>Interesses</b></div>", "size=9; indenty=2; html=true", Font 	
'		Entrevista(40, 1).AddText "<b>Sob o ponto de vista de voc&ecirc;s, o que a crian&ccedil;a gosta mais de fazer? Tem alguma coisa que n&atilde;o goste?</b>", "size=8; indenty=2; html=true", Font 
'		Entrevista(41, 1).AddText gosta_fazer, "size=8; indenty=2; html=true", Font 
'		Entrevista(42, 1).AddText "<div align=""center""><b>Caracter&iacute;sticas</b></div>", "size=9; indenty=2; html=true", Font 	
'		Entrevista(43, 1).AddText "<b>Voc&ecirc;s consideram seu(ua) filho(a) uma crian&ccedil;a f&aacute;cil de lidar? Como reage quando &eacute; contrariada? Quem em casa d&aacute; mais aten&ccedil;&atilde;o &agrave; crian&ccedil;a? Escolaridade anterior?</b>", "size=8; indenty=2; html=true", Font 
'		Entrevista(44, 1).AddText caracteristicas, "size=8; indenty=2; html=true", Font 												

			'SmallTable102(5, 1).Canvas.DrawTable SmallTable10251, "x=0; y=40"	
			'SmallTable112(4, 1).Canvas.DrawTable SmallTable11241, "x=0; y=40"									
			Entrevista(9, 2).Canvas.DrawTable SmallTable92, "x=0; y=39"		
			Entrevista(10, 2).Canvas.DrawTable SmallTable102, "x=0; y=82"	
			Entrevista(11, 2).Canvas.DrawTable SmallTable112, "x=0; y=151"	
			'Entrevista(14, 2).Canvas.DrawTable SmallTable142, "x=0; y=220"	
			'Entrevista(15, 2).Canvas.DrawTable SmallTable152, "x=0; y=220"		
			'Entrevista(17, 2).Canvas.DrawTable SmallTable172, "x=0; y=140"					
			'Entrevista(19, 1).Canvas.DrawTable SmallTable191, "x=0; y=80"			
			'Entrevista(20, 2).Canvas.DrawTable SmallTable202, "x=0; y=80"																
							
		Page.Canvas.DrawTable Entrevista, "x="&margem&", y="&y_table&"" 	
		
					
	Next
	
		limite=0
'		Do While True
'		limite=limite+1
'		   LastRow = Page.Canvas.DrawTable( Entrevista, param_table )
'
'			if LastRow >= Entrevista.Rows.Count Then 
'				Exit Do ' entire table displayed
'			else
				 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
				
				Relatorio = arquivo&" - Sistema Web Diretor"
				Do While Len(Relatorio) > 0
					CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
				 
					If CharsPrinted = Len(Relatorio) Then Exit Do
					   SET Page = Page.NextPage
					Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
				Loop 
				
				Param_Relatorio.Add "alignment=right" 
				
				Paginacao = "1"
				Do While Len(Relatorio) > 0
					CharsPrinted = Page.Canvas.DrawText(Paginacao, Param_Relatorio, Font )
				 
					If CharsPrinted = Len(Paginacao) Then Exit Do
					   SET Page = Page.NextPage
					Paginacao = Right( Paginacao, Len(Relatorio) - CharsPrinted)
				Loop 
				
				
				Param_Relatorio.Add "html=true" 
				
				data_hora = "<center>Impresso em "&data &" &agrave;s "&horario&"</center>"
				Do While Len(Relatorio) > 0
					CharsPrinted = Page.Canvas.DrawText(data_hora, Param_Relatorio, Font )
				 
					If CharsPrinted = Len(data_hora) Then Exit Do
					   SET Page = Page.NextPage
					data_hora = Right( data_hora, Len(data_hora) - CharsPrinted)
				Loop 				   ' Display remaining part of table on the next page
				Set Page = Page.NextPage	
				'param_table.Add( "RowTo=1; RowFrom=1" ) ' Row 1 is header.
				'param_table.Add( "RowTo="&LastRow + 1&"; RowFrom="&LastRow + 1 ) '				
				'param_table("RowFrom1") = LastRow + 1 ' RowTo1 is omitted and presumed infinite
'NOVO CABEÇALHO==========================================================================================		
		Set Param_Logo_Gde = Pdf.CreateParam
		margem=25			
		linha=10		
		unidade = unidade*1	
		if unidade = 1 then
			largura_logo_gde=formatnumber(Logo1.Width*0.6,0)
			altura_logo_gde=formatnumber(Logo1.Height*0.6,0)
			area_utilizavel=Page.Width-(margem*2)
			Param_Logo_Gde("x") = margem
			Param_Logo_Gde("y") = Page.Height - altura_logo_gde -margem
			Param_Logo_Gde("ScaleX") = 0.6
			Param_Logo_Gde("ScaleY") = 0.6
			Page.Canvas.DrawImage Logo1, Param_Logo_Gde
		else
			largura_logo_gde=formatnumber(Logo2.Width*0.5,0)
			altura_logo_gde=formatnumber(Logo2.Height*0.5,0)
			area_utilizavel=Page.Width-(margem*2)
			Param_Logo_Gde("x") = margem
			Param_Logo_Gde("y") = Page.Height - altura_logo_gde -margem
			Param_Logo_Gde("ScaleX") = 0.5
			Param_Logo_Gde("ScaleY") = 0.5
			Page.Canvas.DrawImage Logo2, Param_Logo_Gde		
		end if

		x_texto=largura_logo_gde+ margem+10
		y_texto=formatnumber(Page.Height - margem,0)
		width_texto=Page.Width -largura_logo_gde - 80

	
		SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
		Text = "<p><i><b>"&UCASE(no_unidade)&"</b></i></p><br><font style=""font-size:10pt;"">"&endereco_unidade&"</FONT>" 			
		
		Do While Len(Text) > 0
			CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
		 
			If CharsPrinted = Len(Text) Then Exit Do
				SET Page = Page.NextPage
			Text = Right( Text, Len(Text) - CharsPrinted)
		Loop 

		y_texto=y_texto-altura_logo_gde+10
		SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height=30; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
		Text = "<center><i><b><font style=""font-size:18pt;"">ANAMNESE</font></b></i></center>"
		
		
		Do While Len(Text) > 0
			CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
		 
			If CharsPrinted = Len(Text) Then Exit Do
				SET Page = Page.NextPage
			Text = Right( Text, Len(Text) - CharsPrinted)
		Loop 

		
		Page.Canvas.SetParams "LineWidth=1" 
		Page.Canvas.SetParams "LineCap=0" 
		inicio_primeiro_separador=largura_logo_gde+margem+10
		altura_primeiro_separador= Page.Height - margem - 17
		With Page.Canvas
		   .MoveTo inicio_primeiro_separador, altura_primeiro_separador
		   .LineTo area_utilizavel+margem, altura_primeiro_separador
		   .Stroke
		End With 	


		Page.Canvas.SetParams "LineWidth=1" 
		Page.Canvas.SetParams "LineCap=0" 
		altura_segundo_separador= Page.Height - altura_logo_gde-margem - 20
		With Page.Canvas
		   .MoveTo margem, altura_segundo_separador
		   .LineTo area_utilizavel+margem, altura_segundo_separador
		   .Stroke
		End With 	

		y_nome_aluno=Page.Height - altura_logo_gde-46
		width_nome_aluno=Page.Width - margem
		
		SET Param_Nome_Aluno = Pdf.CreateParam("x="&margem&";y="&y_nome_aluno&"; height=50; width="&width_nome_aluno&"; alignment=left; html=True")
		Nome = "<font style=""font-size:11pt;""><b>Alun"&desinencia&": "&nome_aluno&"</b></font>"
		

		Do While Len(Nome) > 0
			CharsPrinted = Page.Canvas.DrawText(Nome, Param_Nome_Aluno, Font )
		 
			If CharsPrinted = Len(Nome) Then Exit Do
				SET Page = Page.NextPage
			Nome = Right( Nome, Len(Nome) - CharsPrinted)
		Loop 
		
		Page.Canvas.SetParams "LineWidth=2" 
		Page.Canvas.SetParams "LineCap=0" 
		With Page.Canvas
		   .MoveTo margem, Page.Height - altura_logo_gde-65
		   .LineTo Page.Width - margem, Page.Height - altura_logo_gde-65
		   .Stroke
		End With 	


		Set param_table1 = Pdf.CreateParam("width=533; height=25; rows=2; cols=8; border=0; cellborder=0; cellspacing=0;")
		Set Table = Doc.CreateTable(param_table1)
		Table.Font = Font
		y_table=Page.Height - altura_logo_gde-70
		
		With Table.Rows(1)
		   .Cells(1).Width = 40
		   .Cells(2).Width = 200
		   .Cells(3).Width = 25
		   .Cells(4).Width = 70
		   .Cells(5).Width = 60
		   .Cells(6).Width = 38
		   .Cells(7).Width = 50
		   .Cells(8).Width = 50      
		End With
		Table(1, 2).ColSpan = 5
		Table(1, 1).AddText "Unidade:", "size=9;", Font 
		Table(2, 1).AddText "Curso:", "size=9;", Font 
		Table(1, 2).AddText no_unidade, "size=9;", Font 
		Table(2, 2).ColSpan = 2
		Table(2, 2).AddText no_curso, "size=9; html=true", Font 
		'Table(2, 3).AddText no_etapa, "size=9;", Font 
		Table(2, 4).AddText "Turma: "&turma, "size=9;", Font 
		Table(2, 5).AddText "N&ordm;. Chamada: "&cham, "size=9; html=true", Font 
		Table(2, 6).AddText cham, "size=9;", Font 
		Table(1, 7).AddText "<div align=""right"">Matr&iacute;cula: </div>", "size=9; html=true", Font 
		Table(1, 8).AddText cod_cons, "size=9;alignment=right", Font 
		Table(2, 7).AddText "Ano Letivo: ", "size=9; alignment=right", Font 
		Table(2, 8).AddText ano_letivo, "size=9;alignment=right", Font 
		Page.Canvas.DrawTable Table, "x="&margem&", y="&y_table&"" 
	
		
		With Page.Canvas
		   .MoveTo margem, Page.Height - altura_logo_gde-100
		   .LineTo Page.Width - margem, Page.Height - altura_logo_gde-100
		   .Stroke
		End With 	
'================================================================================================================			
		y_nome_aluno=Page.Height - altura_logo_gde-46
		
altura_medias=1040
		largura_tabela=680
		y_table=y_table-50		
		Set param_table = Pdf.CreateParam("width="&largura_tabela&"; height="&altura_medias&"; rows=16; cols=4; border=1; cellborder=0.1; cellspacing=0;  x="&margem&"; y="&y_table&";")
		Set Entrevista = Doc.CreateTable(param_table)
		'param_table.Set "" 			
		Entrevista.Font = Font
		With Entrevista.Rows(1)
		   .Cells(1).Width = 100			
		   .Cells(3).Width = 100			     		         			   			         
		End With

		Entrevista.Rows(1).Cells(1).Height = 13	
		Entrevista.Rows(2).Cells(1).Height = 30
		Entrevista.Rows(3).Cells(1).Height = 15	
		Entrevista.Rows(4).Cells(1).Height = 17
		Entrevista.Rows(5).Cells(1).Height = 13
		Entrevista.Rows(6).Cells(1).Height = 17
		Entrevista.Rows(7).Cells(1).Height = 13			
		Entrevista.Rows(8).Cells(1).Height = 17
		Entrevista.Rows(9).Cells(1).Height = 13	
		Entrevista.Rows(10).Cells(1).Height = 13	
		Entrevista.Rows(11).Cells(1).Height = 17	
		Entrevista.Rows(12).Cells(1).Height = 13								
		Entrevista.Rows(13).Cells(1).Height = 39
		Entrevista.Rows(14).Cells(1).Height = 17		
		Entrevista.Rows(15).Cells(1).Height = 26			
		Entrevista.Rows(16).Cells(1).Height = 39																																																				
		
		Entrevista(1, 1).ColSpan = 4		
		Entrevista(2, 1).ColSpan = 4	
		Entrevista(2, 1).RowSpan = 2	
		Entrevista(4, 1).ColSpan = 4	
		Entrevista(5, 1).ColSpan = 4	
		Entrevista(6, 1).ColSpan = 4	
		Entrevista(7, 2).ColSpan = 3	
		Entrevista(8, 1).ColSpan = 4	
		Entrevista(9, 1).ColSpan = 4
		Entrevista(10, 2).ColSpan = 3
		Entrevista(11, 1).ColSpan = 4
		Entrevista(12, 1).ColSpan = 4
		Entrevista(13, 1).ColSpan = 4
		Entrevista(14, 1).ColSpan = 4
		Entrevista(15, 1).ColSpan = 4	
		Entrevista(16, 1).ColSpan = 4																																																																															
		
		Entrevista(1, 1).AddText "<b>Antecedentes familiares:&nbsp;</b>", "x=1; y=1, size=8; html=true", Font		
		Entrevista(2, 1).AddText antecedentes, "x=1; y=1, size=8; html=true", Font	
											
								
			
		Entrevista(4, 1).AddText "<div align=""center""><b>Divertimentos da Fam&iacute;lia</b></div>", "size=9; indenty=2; html=true", Font 				
		Entrevista(5, 1).AddText divertimentos, "size=8; html=true; indentx=2; expand=true", Font 						
																											
		Entrevista(6, 1).AddText "<div align=""center""><b>Aquisi&ccedil;&atilde;o de h&aacute;bitos</b></div>", "size=9; indenty=2; html=true", Font 	
		Entrevista(7, 1).AddText "<div align=""right""><b>Higiene (banho etc):&nbsp;</b></div>", "x=1; y=1, size=8; html=true", Font		
		Entrevista(7, 2).AddText higiene, "x=1; y=1, size=8; html=true", Font															
		Entrevista(8, 1).AddText "<b>Como e quando foi treinado (controle dos esfincteres)?</b>", "x=1; y=1, size=8; html=true", Font		
		Entrevista(9, 1).AddText controle, "x=1; y=1, indentx=2; size=8; html=true", Font		
		Entrevista(10, 1).AddText "<div align=""right""><b>Sono:&nbsp;</b></div>", "x=1; y=1, size=8; html=true", Font		
		Entrevista(10, 2).AddText sono, "x=1; y=1, size=8; html=true", Font							
		Entrevista(11, 1).AddText "<div align=""center""><b>Interesses</b></div>", "size=9; indenty=2; html=true", Font 	
		Entrevista(12, 1).AddText "<b>Sob o ponto de vista de voc&ecirc;s, o que a crian&ccedil;a gosta mais de fazer? Tem alguma coisa que n&atilde;o goste?</b>", "size=8; indenty=2; html=true", Font 
		Entrevista(13, 1).AddText gosta_fazer, "size=8; indenty=2; html=true", Font 
		Entrevista(14, 1).AddText "<div align=""center""><b>Caracter&iacute;sticas</b></div>", "size=9; indenty=2; html=true", Font 	
		Entrevista(15, 1).AddText "<b>Voc&ecirc;s consideram seu(ua) filho(a) uma crian&ccedil;a f&aacute;cil de lidar? Como reage quando &eacute; contrariada? Quem em casa d&aacute; mais aten&ccedil;&atilde;o &agrave; crian&ccedil;a? Escolaridade anterior?</b>", "size=8; indenty=2; html=true", Font 
		Entrevista(16, 1).AddText caracteristicas, "size=8; indenty=2; html=true", Font 												

															
							
		Page.Canvas.DrawTable Entrevista, "x="&margem&", y="&y_table&"" 		
		 
			end if
			if limite>100 then
			response.Write("ERRO!")
			response.end()
			end if 
		'Loop	
		
		 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
		
		Relatorio = arquivo&" - Sistema Web Diretor"
		Do While Len(Relatorio) > 0
			CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
		 
			If CharsPrinted = Len(Relatorio) Then Exit Do
			   SET Page = Page.NextPage
			Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
		Loop 
		
		Param_Relatorio.Add "alignment=right" 
		
		Paginacao = "2"
		Do While Len(Relatorio) > 0
			CharsPrinted = Page.Canvas.DrawText(Paginacao, Param_Relatorio, Font )
		 
			If CharsPrinted = Len(Paginacao) Then Exit Do
			   SET Page = Page.NextPage
			Paginacao = Right( Paginacao, Len(Relatorio) - CharsPrinted)
		Loop 
		
		
		Param_Relatorio.Add "html=true" 
		
		data_hora = "<center>Impresso em "&data &" &agrave;s "&horario&"</center>"
		Do While Len(Relatorio) > 0
			CharsPrinted = Page.Canvas.DrawText(data_hora, Param_Relatorio, Font )
		 
			If CharsPrinted = Len(data_hora) Then Exit Do
			   SET Page = Page.NextPage
			data_hora = Right( data_hora, Len(data_hora) - CharsPrinted)
		Loop 							
'END IF
Doc.SaveHttp("attachment; filename="&arquivo&".pdf")
 %>