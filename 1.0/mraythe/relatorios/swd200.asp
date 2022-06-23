<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 30 'valor em segundos
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<% 
arquivo="SWD200"
response.Charset="ISO-8859-1"
opt= request.QueryString("opt")
ori= request.QueryString("ori")

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
nvg=session("nvg")
session("nvg")=nvg
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

if ori="edf" then
origem="../ws/mat/man/eco/"
end if




if mes<10 then
mes="0"&mes
end if

data = dia &"/"& mes &"/"& ano

if mes=1 then
	mes_extenso="Janeiro"
elseif mes=2 then
	mes_extenso="Fevereiro"
elseif mes=3 then
	mes_extenso="Mar&ccedil;o"
elseif mes=4 then
	mes_extenso="Abril"
elseif mes=5 then
	mes_extenso="Maio"
elseif mes=6 then
	mes_extenso="Junho"
elseif mes=7 then
	mes_extenso="Julho"
elseif mes=8 then
	mes_extenso="Agosto"
elseif mes=9 then
	mes_extenso="Setembro"
elseif mes=10 then
	mes_extenso="Outubro"
elseif mes=11 then
	mes_extenso="Novembro"
elseif mes=12 then
	mes_extenso="Dezembro"
end if	
data_extenso="Rio de Janeiro, "&dia &" de "& mes_extenso &" de "& ano
if min<10 then
min="0"&min
end if

horario = hora & ":"& min

	'Dim AspPdf, Doc, Page, Font, Text, Param, Image, CharsPrinted
	'Instancia o objeto na mem&oacute;ria
	SET Pdf = Server.CreateObject("Persits.Pdf")
	SET Doc = Pdf.CreateDocument
	Set Logo = Doc.OpenImage( Server.MapPath( "../img/logo_pdf.gif") )
	Set Font = Doc.Fonts.LoadFromFile(Server.MapPath("../fonts/arial.ttf"))	
	Set Font_Tesoura = Doc.Fonts.LoadFromFile(Server.MapPath("../fonts/ZapfDingbats.ttf"))
	If Font.Embedding = 2 Then
	   Response.Write "Embedding of this font is prohibited."
	   Set Font = Nothing
	End If
	If Font_Tesoura.Embedding = 2 Then
	   Response.Write "Embedding of this font is prohibited."
	   Set Font = Nothing
	End If 			 		 

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0	
		
	
if opt="01" then
	cod_cons=request.QueryString("cod_cons")
	documento="D99999"

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Situacao='C' AND CO_Matricula ="& cod_cons
		RS1.Open SQL1, CON1
		
		if RS1.EOF then
			response.redirect(origem&"index.asp?nvg="&nvg&"&opt=err4")
		end if

		If Not IsArray(alunos_encontrados) Then alunos_encontrados = Array() End if	
		ReDim preserve alunos_encontrados(UBound(alunos_encontrados)+1)	
		alunos_encontrados(Ubound(alunos_encontrados)) = cod_cons	

	
elseif opt="02" then
	obr=request.QueryString("obr")
	dados_informados = split(obr, "$!$" )
	gera_declaracao_terceiro_ano="nao"
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


'RESPONSE.END()
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

		unidade_base= "nulo"
		curso_base= "nulo"
		co_etapa_base= "nulo"
		turma_base= "nulo"	

		cod_cons=alunos_encontrados(i)
					
				
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod_cons
		RS1.Open SQL1, CON1
		
		if RS1.EOF then
			response.redirect(origem&"index.asp?nvg="&nvg&"&opt=err4")
		else
		
			ano_aluno = RS1("NU_Ano")
			rematricula = RS1("DA_Rematricula")
			situacao = RS1("CO_Situacao")
			encerramento= RS1("DA_Encerramento")
			unidade= RS1("NU_Unidade")
			curso= RS1("CO_Curso")
			co_etapa= RS1("CO_Etapa")
			turma= RS1("CO_Turma")
			cham= RS1("NU_Chamada")
			
			teste_u = isnumeric(unidade)
			if teste_u= true then
				unidade=unidade*1
			end if	
			
			teste_c = isnumeric(curso)
			if teste_c= true then
				curso=curso*1
			end if
			
			teste_e = isnumeric(co_etapa)
			if teste_e= true then
				co_etapa=co_etapa*1
			end if								
	
			teste_t = isnumeric(turma)
			if teste_t= true then
				turma=turma*1
			end if	
			
			teste_u2 = isnumeric(unidade_base)
			if teste_u2= true then
				unidade_base=unidade_base*1
			end if	
			
			teste_c2 = isnumeric(curso_base)
			if teste_c2= true then
				curso_base=curso_base*1
			end if
			
			teste_e2 = isnumeric(co_etapa_base)
			if teste_e2= true then
				co_etapa_base=co_etapa_base*1
			end if								
	
			teste_t2 = isnumeric(turma_base)
			if teste_t2= true then
				turma_base=turma_base*1
			end if				
		end if
		
		if unidade_base<>unidade or curso_base<>curso or co_etapa_base<>co_etapa or	turma_base<>turma then
			regera_cabecalho_ucet="s"
			unidade_base= unidade
			curso_base= curso
			co_etapa_base= co_etapa
			turma_base= turma	
			qtd_turma=qtd_turma+1			
		else	
			regera_cabecalho_ucet="n"
		End if				

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod_cons
		RS.Open SQL, CON1
		
		nome_aluno = RS("NO_Aluno")
		resp_fin = RS("TP_Resp_Fin")

		
			Set RSdn = Server.CreateObject("ADODB.Recordset")
			SQLdn = "SELECT * FROM TB_Contatos WHERE CO_Matricula ="& cod_cons&" AND TP_Contato='"&resp_fin&"'"
			RSdn.Open SQLdn, CONCONT	

			if 	RSdn.EOF then
				responsavel="<i>N&atilde;o informado no cadastro</i>"				
			else
				responsavel=RSdn("NO_Contato")
			
				if responsavel="" or isnull(responsavel)  then
					responsavel="<i>N&atilde;o informado no cadastro</i>"
				else
					responsavel =replace_latin_char(responsavel,"html")
				end if	
			end if

		
		if nome_aluno="" or isnull(nome_aluno)  then
			nome_aluno="<i>N&atilde;o informado no cadastro</i>"
		else
			nome_aluno =replace_latin_char(nome_aluno,"html")
		end if	
		
			if sexo_aluno="F" then
				desinencia="a"
			else
				desinencia="o"
			end if
						
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Unidade WHERE NU_Unidade="& unidade
		RS2.Open SQL2, CON0
		no_unidade_abr = RS2("NO_Abr")									
		no_unidade = RS2("NO_Sede")		
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

		if bairro_unidade="" or isnull(bairro_unidade)then
		else
			bairro_unidade=" - "&bairro_unidade
		end if
		
		if uf_unidade="" or isnull(uf_unidade)then
		else
			uf_unidade_municipio=uf_unidade
			uf_unidade=" - "&uf_unidade
		end if
		
		if municipio_unidade="" or isnull(municipio_unidade) or uf_unidade_municipio="" or isnull(uf_unidade_municipio)then
		else
			Set RS3m = Server.CreateObject("ADODB.Recordset")
			SQL3m = "SELECT * FROM TB_Municipios WHERE SG_UF='"& uf_unidade_municipio &"' AND CO_Municipio="&municipio_unidade
			RS3m.Open SQL3m, CON0
			
			municipio_unidade=RS3m("NO_Municipio")						
		end if
		endereco_unidade=rua_unidade&numero_unidade&complemento_unidade&bairro_unidade&cep_unidade&"<br>"&municipio_unidade&uf_unidade					
					
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Curso WHERE CO_Curso='"& curso &"'"
		RS3.Open SQL3, CON0
		
		no_curso= RS3("NO_Curso")
		no_abrv_curso = RS3("NO_Abreviado_Curso")
		co_concordancia_curso = RS3("CO_Conc")	
		
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Etapa WHERE CO_Curso='"& curso &"' AND CO_Etapa ='"& co_etapa &"'"
		RS4.Open SQL4, CON0
		
		no_etapa = RS4("NO_Etapa")
		art_conc = RS4("CO_Conc")
						
		txt_curso= "d"&art_conc&" "&no_etapa&"&nbsp;"&co_concordancia_curso&"&nbsp;"&no_curso
		'texto_turma = "&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;Turma "&turma
		'mensagem_turma="<b>Unidade: "&no_unidade_abr&"&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;"&no_curso&texto_turma&"</b>"
		'mensagem_cabecalho=ano_letivo
		

'CABE&Ccedil;ALHO==========================================================================================		
		Set Param_Logo_Gde = Pdf.CreateParam

		largura_logo_gde=formatnumber(Logo.Width*0.4,0)
		altura_logo_gde=formatnumber(Logo.Height*0.4,0)
		margem=30	
		linha=10
		area_utilizavel=Page.Width-(margem*2)
		Param_Logo_Gde("x") = formatnumber(area_utilizavel/2,0)
		Param_Logo_Gde("y") = Page.Height - altura_logo_gde -margem
		Param_Logo_Gde("ScaleX") = 0.4
		Param_Logo_Gde("ScaleY") = 0.4
		Page.Canvas.DrawImage Logo, Param_Logo_Gde

		x_texto=largura_logo_gde+ margem+10
		y_posicao=y_posicao - margem
		y_posicao=y_posicao - altura_logo_gde-linha
		
		SET Param = Pdf.CreateParam("x="&margem&";y="&y_posicao&"; height=20; width="&area_utilizavel&"; alignment=center; size=14; color=#000000; html=true")
		
		Text = "<center><b><font style=""font-size:10pt;"">CONTRATO DE PRESTA&Ccedil;&Atilde;O DE SERVI&Ccedil;OS DE EDUCA&Ccedil;&Atilde;O ESCOLAR</FONT></b></center>" 			
		
		Do While Len(Text) > 0
			CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
		 
			If CharsPrinted = Len(Text) Then Exit Do
				SET Page = Page.NextPage
			Text = Right( Text, Len(Text) - CharsPrinted)
		Loop 
		
		y_posicao=y_posicao- 30
		
		y_diponivel=y_posicao							

'FIM DO CABE&Ccedil;ALHO==========================================================================================	
		
		SET Param = Pdf.CreateParam("x="&margem&";y="&y_posicao&"; height="&y_diponivel&"; width="&area_utilizavel&"; alignment=center; size=10; color=#000000; html=true")
		
		Text = "<div align=""justify""><p>Por meio do presente instrumento particular o CONTRATANTE, "&responsavel&" na qualidade de representante legal do aluno(a):<b>"&nome_aluno&"</b>, "&txt_curso&" , qualificado na ficha de matr&iacute;cula, que passa a fazer parte do presente contrato, de um lado, e de outro lado, como CONTRATADO: COL&Eacute;GIO SAINT JOHN LTDA., pessoa jur&iacute;dica de direito privado, com sede na, Av. Gal. Felic&iacute;ssimo Cardoso, 841, Barra da Tijuca &shy; Rio de Janeiro, CEP 22631-360, CNPJ n&ordm; 30.474.902/0001-20, mantenedor do Col&eacute;gio Saint John, firmam o presente CONTRATO DE PRESTA&Ccedil;&Atilde;O DE SERVI&Ccedil;OS EDUCACIONAIS conforme cl&aacute;usulas e condi&ccedil;&otilde;es abaixo:</p><p><b>CL&Aacute;USULA PRIMEIRA:</b> O objeto do presente contrato &eacute; regular a presta&ccedil;&atilde;o de servi&ccedil;os de educa&ccedil;&atilde;o escolar, pelo CONTRATADO, correspondente ao segmento e ano, com in&iacute;cio no dia 1&ordm; de janeiro de 2010, ao aluno indicado pelo CONTRATANTE; definir a contrapresta&ccedil;&atilde;o pecuni&aacute;ria e a forma de pagamento por parte do CONTRATANTE, bem como estabelecer os dispositivos ou eventuais restitui&ccedil;&otilde;es, por for&ccedil;a do cancelamento deste contrato, celebrado sob o disposto na Constitui&ccedil;&atilde;o Federal, em especial nos seus artigos 1&ordm;, inciso IV; 5&ordm;, inciso II; 206, incisos II e III, e 209; arts. 389, 476 e 597 do C&oacute;digo Civil Brasileiro; da Lei 8.069/90 (Estatuto da Crian&ccedil;a e do Adolescente); da Lei 8078/90, (C&oacute;digo de Defesa do Consumidor); Lei  9.870/99, Lei 11.114/05 e Lei 11.274/06, bem como das cl&aacute;usulas e condi&ccedil;&otilde;es adiante descritas, comprometendo-se reciprocamente a cumpri-las.</p><p><b>Par&aacute;grafo &Uacute;nico: </b>Caso a responsabilidade financeira seja transferida para outra pessoa no decorrer do per&iacute;odo letivo, dever&aacute; ser entregue na secretaria da unidade documento comprovando essa altera&ccedil;&atilde;o, bem como a qualifica&ccedil;&atilde;o do novo respons&aacute;vel.</p><p><b>CL&Aacute;USULA SEGUNDA: </b>N&atilde;o est&atilde;o inclu&iacute;dos neste contrato as atividades extracurriculares, os servi&ccedil;os especiais de recupera&ccedil;&atilde;o, refor&ccedil;o, material did&aacute;tico de uso individual, uniformes, alimenta&ccedil;&atilde;o, transporte escolar, segunda via de documentos escolares.</p><p><b>CL&Aacute;USULA TERCEIRA:</b> Somente estar&aacute; efetivada a matr&iacute;cula quando do preenchimento do &quot;Requerimento de Matr&iacute;cula&quot;, que &eacute; parte integrante deste Contrato, da entrega da documenta&ccedil;&atilde;o exigida, no prazo m&aacute;ximo de 30 dias, e ap&oacute;s o deferimento pelo CONTRATADO, bem como da assinatura do presente contrato, n&atilde;o podendo este sofrer qualquer tipo de rasura ou mudan&ccedil;a unilateral.</p><p><b>Par&aacute;grafo 1&ordm; - </b>Firmado o presente contrato, o CONTRATANTE submete-se ao Regimento Escolar e &agrave;s demais obriga&ccedil;&otilde;es constantes na legisla&ccedil;&atilde;o aplic&aacute;vel na &aacute;rea de educa&ccedil;&atilde;o, obrigando-se a fazer com que o aluno cumpra o calend&aacute;rio escolar, os hor&aacute;rios e as normas estabelecidas pelo CONTRATADO, assumindo total responsabilidade pelos problemas advindos da n&atilde;o observ&acirc;ncia das normas escolares citadas.</p><p><b>Par&aacute;grafo 2&ordm; - </b>O n&atilde;o comparecimento do aluno aos servi&ccedil;os educacionais, ora contratados, n&atilde;o tem o poder de desobrig&aacute;-lo das demais cl&aacute;usulas do presente contrato, tendo em vista que os servi&ccedil;os foram postos &agrave; disposi&ccedil;&atilde;o do CONTRATANTE.</p><p><b>Par&aacute;grafo 3&ordm; - </b>O CONTRATANTE assume total responsabilidade quanto &agrave;s declara&ccedil;&otilde;es prestadas neste contrato e no ato da matr&iacute;cula, relativas &agrave; aptid&atilde;o legal do aluno para frequ&ecirc;ncia na s&eacute;rie e graus indicados, quando for o caso, concordando, desde j&aacute;, que a n&atilde;o entrega dos documentos legais comprobat&oacute;rios das declara&ccedil;&otilde;es prestadas, at&eacute; 60 (sessenta) dias contados do in&iacute;cio das aulas, acarretar&aacute; o autom&aacute;tico cancelamento da vaga aberta ao aluno, rescindindo-se o presente contrato, encerrando-se a presta&ccedil;&atilde;o de servi&ccedil;os e isentando o CONTRATADO de qualquer responsabilidade pelos eventuais danos resultantes.</p><p><b>CL&Aacute;USULA QUARTA:</b> Compete ao CONTRATADO as orienta&ccedil;&otilde;es t&eacute;cnicas e pedag&oacute;gicas que se fizerem necess&aacute;rias para execu&ccedil;&atilde;o dos servi&ccedil;os contratados. Inclui-se a marca&ccedil;&atilde;o de datas das provas de aproveitamento, fixa&ccedil;&atilde;o de carga hor&aacute;ria, determina&ccedil;&atilde;o de aulas extras, indica&ccedil;&atilde;o de professores, al&eacute;m de outras provid&ecirc;ncias que a atividade docente exige.</p><p><b>CL&Aacute;USULA QUINTA: </b>O CONTRATANTE se responsabiliza por qualquer dano material ocasionado pelo aluno nas depend&ecirc;ncias da escola ou das empresas contratadas por esta para as atividades fora da sala de aula.</p><p><b>Par&aacute;grafo &Uacute;nico - </b>As aulas ser&atilde;o ministradas em salas de aula ou em locais que o CONTRATADO indicar, tendo em vista a natureza dos conte&uacute;dos a serem estudados e das t&eacute;cnicas pedag&oacute;gicas que ser&atilde;o aplicadas.</p></div>" 			
		
		Text =replace_latin_char(Text,"html")
		
		Do While Len(Text) > 0
			CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
		 
			If CharsPrinted = Len(Text) Then Exit Do
												
			Text = Right( Text, Len(Text) - CharsPrinted)
		Loop 
				
				

	SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000;html=true")
	Relatorio = arquivo
	Do While Len(Relatorio) > 0
		CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
	 
		If CharsPrinted = Len(Relatorio) Then Exit Do
		   SET Page = Page.NextPage
		Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
	Loop 
				
			SET Param_data_hora = Pdf.CreateParam("x="&margem&";y="&y_posicao&"; height="&y_diponivel&"; width="&area_utilizavel&"; alignment=center; size=8; color=#000000; html=true")

			Param_data_hora.Add "y="&margem+12&";html=true" 
			
			data_hora = "<div align=""Right"">"&paginacao&"<br>Impresso em "&data &" &agrave;s "&horario&"</div>"
			Do While Len(data_hora) > 0
				CharsPrinted = Page.Canvas.DrawText(data_hora, Param_data_hora, Font )			
				If CharsPrinted = Len(data_hora) Then Exit Do
				SET Page = Page.NextPage
				data_hora = Right( data_hora, Len(data_hora) - CharsPrinted)
			Loop 				
'FIM DA PRIMEIRA P&Aacute;GINA=========================================================================================================================================				
			
				paginacao = paginacao+1			
				SET Page = Page.NextPage			
				y_posicao=842
	'CABE&Ccedil;ALHO==========================================================================================		
				Set Param_Logo_Gde = Pdf.CreateParam
	
				largura_logo_gde=formatnumber(Logo.Width*0.4,0)
				altura_logo_gde=formatnumber(Logo.Height*0.4,0)
				margem=30	
				linha=10
				area_utilizavel=Page.Width-(margem*2)
				Param_Logo_Gde("x") = formatnumber(area_utilizavel/2,0)
				Param_Logo_Gde("y") = Page.Height - altura_logo_gde -margem
				Param_Logo_Gde("ScaleX") = 0.4
				Param_Logo_Gde("ScaleY") = 0.4
				Page.Canvas.DrawImage Logo, Param_Logo_Gde
		
				x_texto=largura_logo_gde+ margem+10
				y_posicao=y_posicao - margem
				y_posicao=y_posicao - altura_logo_gde-linha
				
				SET Param = Pdf.CreateParam("x="&margem&";y="&y_posicao&"; height=20; width="&area_utilizavel&"; alignment=center; size=14; color=#000000; html=true")
				
				Text = "<center><b><font style=""font-size:10pt;"">CONTRATO DE PRESTA&Ccedil;&Atilde;O DE SERVI&Ccedil;OS DE EDUCA&Ccedil;&Atilde;O ESCOLAR</FONT></b></center>" 			
				
				Do While Len(Text) > 0
					CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
				 
					If CharsPrinted = Len(Text) Then Exit Do
						SET Page = Page.NextPage
					Text = Right( Text, Len(Text) - CharsPrinted)
				Loop 
				
				y_posicao=y_posicao- 30
				
				y_diponivel=y_posicao							
'	
	'FIM DO CABE&Ccedil;ALHO==========================================================================================	

			SET Param = Pdf.CreateParam("x="&margem&";y="&y_posicao&"; height=40; width="&area_utilizavel&"; alignment=center; size=10; color=#000000; html=true")
			
			Text = "<div align=""justify""><p><b>CL&Aacute;USULA SEXTA:</b> Como contrapresta&ccedil;&atilde;o dos servi&ccedil;os de educa&ccedil;&atilde;o escolar, o CONTRATANTE pagar&aacute; a anuidade referente ao ano letivo de 2010 &agrave; vista ou em 12 (doze) parcelas mensais e sucessivas, com vencimento no dia 05 de cada m&ecirc;s, conforme tabela abaixo:</p></div>" 			
			
			Text =replace_latin_char(Text,"html")
			
			Do While Len(Text) > 0
				CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
			 
				If CharsPrinted = Len(Text) Then Exit Do
													
				Text = Right( Text, Len(Text) - CharsPrinted)
			Loop 
				
			y_posicao=y_posicao- 60
				
			Set param_table1 = Pdf.CreateParam("width=400; height=100; rows=6; cols=5; border=0.5; cellborder=0.5; cellspacing=0;")
			Set Table = Doc.CreateTable(param_table1)
			Table.Font = Font
			y_primeira_tabela=y_posicao
			With Table.Rows(1)
			   .Cells(1).Width = 160		   		   		   
			   .Cells(2).Width = 60
			   .Cells(3).Width = 60
			   .Cells(4).Width = 60
			   .Cells(5).Width = 60	
			   .Cells(1).Height = 20		   		   		   
			   .Cells(2).Height = 20	
			   .Cells(3).Height = 20	
			   .Cells(4).Height = 20	
			   .Cells(5).Height = 20		   				   	
			End With
			
			Table(1, 1).AddText "<center><b>CURSO</b></center>", "size=8;html=true", Font 
			Table(1, 2).AddText "<center><b>ANUIDADE</b></center>", "size=8;html=true", Font 
			Table(1, 3).AddText "<center><b>PARCELAS</b></center>", "size=8;html=true", Font 
			Table(1, 4).AddText "<center><b>VALOR</b></center>", "size=8;html=true", Font 
			Table(1, 5).AddText "<center><b>VENCIMENTO</b></center>", "size=8;html=true", Font 												
			Table(2, 1).AddText "Educa&ccedil;&atilde;o Infantil", "size=8;html=true", Font 
			Table(2, 2).AddText "<center>R$ 8.892,00</center>", "size=8;html=true", Font 
			Table(2, 3).AddText "<center>12</center>", "size=8;html=true", Font 
			Table(2, 4).AddText "<center>R$   741,00</center>", "size=8;html=true", Font 
			Table(2, 5).AddText "<center>05</center>", "size=8;html=true", Font 
			Table(3, 1).AddText "1&ordm; ao 5&ordm; Ano -  Ensino Fundamental", "size=8;html=true", Font 
			Table(3, 2).AddText "<center>R$10.080,00</center>", "size=8;html=true", Font 
			Table(3, 3).AddText "<center>12</center>", "size=8;html=true", Font 
			Table(3, 4).AddText "<center>R$   840,00</center>", "size=8;html=true", Font 
			Table(3, 5).AddText "<center>05</center>", "size=8;html=true", Font 
			Table(4, 1).AddText "6&ordm; ao 9&ordm; Ano -  Ensino Fundamental", "size=8;html=true", Font 
			Table(4, 2).AddText "<center>R$10.476,00</center>", "size=8;html=true", Font 
			Table(4, 3).AddText "<center>12</center>", "size=8;html=true", Font 
			Table(4, 4).AddText "<center>R$   873,00</center>", "size=8;html=true", Font 
			Table(4, 5).AddText "<center>05</center>", "size=8;html=true", Font 
			Table(5, 1).AddText "1&ordf; e 2&ordf; S&eacute;rie -  Ensino M&eacute;dio", "size=8;html=true", Font 
			Table(5, 2).AddText "<center>R$12.360,00</center>", "size=8;html=true", Font 
			Table(5, 3).AddText "<center>12</center>", "size=8;html=true", Font 
			Table(5, 4).AddText "<center>R$1.030,00</center>", "size=8;html=true", Font 
			Table(5, 5).AddText "<center>05</center>", "size=8;html=true", Font 	
			Table(6, 1).AddText "3&ordf; S&eacute;rie -  Ensino M&eacute;dio", "size=8;html=true", Font 
			Table(6, 2).AddText "<center>R$14.436,00</center>", "size=8;html=true", Font 
			Table(6, 3).AddText "<center>12</center>", "size=8;html=true", Font 
			Table(6, 4).AddText "<center>R$1.203,00</center>", "size=8;html=true", Font 
			Table(6, 5).AddText "<center>05</center>", "size=8;html=true", Font 													
				
			Page.Canvas.DrawTable Table, "x=100, y="&y_primeira_tabela&"" 	

			y_posicao=y_posicao-120	
		
			y_diponivel=y_posicao				
			
			SET Param = Pdf.CreateParam("x="&margem&";y="&y_posicao&"; height="&y_diponivel&"; width="&area_utilizavel&"; alignment=center; size=10; color=#000000; html=true")
			
			Text = "<div align=""justify""><p><b>Par&aacute;grafo 1&ordm; - </b>O adiamento ocasional da data de vencimento de qualquer parcela da anuidade, como eventuais descontos ou benef&iacute;cios por parte do CONTRATADO, n&atilde;o constituir&aacute; nova&ccedil;&atilde;o, ou seja, a substitui&ccedil;&atilde;o da obriga&ccedil;&atilde;o anterior pela da nova concess&atilde;o.</p><p><b>Par&aacute;grafo 2&ordm; - </b>A concess&atilde;o de bolsa de estudo integral somente isenta do pagamento da anuidade, devendo cumprir-se as demais obriga&ccedil;&otilde;es descritas neste contrato.</p><p><b>Par&aacute;grafo 3&ordm; - No pagamento efetuado ap&oacute;s a data de vencimento ser&aacute; acrescido multa de 2% (dois por cento). Quando o pagamento for efetuado ap&oacute;s 30 dias do vencimento, mesmo que por meios eletr&ocirc;nicos, al&eacute;m da multa de 2% (dois por cento), ser&atilde;o acrescidos juros de 1% (um por cento) ao m&ecirc;s.</b></p><p><b>Par&aacute;grafo 4&ordm; - </b>O CONTRATANTE declara que teve conhecimento pr&eacute;vio das condi&ccedil;&otilde;es financeiras deste contrato e que este foi exposto em local de f&aacute;cil acesso e visualiza&ccedil;&atilde;o (art. 2&ordm; da Lei n&ordm; 9.870/99), conhecendo-as e aceitando-as livremente.</p><p><b>CL&Aacute;USULA S&Eacute;TIMA: </b>Qualquer altera&ccedil;&atilde;o de valor da anuidade, que por for&ccedil;a de lei venha a ocorrer no decurso deste ano letivo, ser&aacute; repassada no m&ecirc;s da ocorr&ecirc;ncia.</p><p><b>CL&Aacute;USULA OITAVA: </b>No caso de inadimpl&ecirc;ncia, o CONTRATADO poder&aacute;:</p><p>I - Proceder ao competente protesto das parcelas vencidas em prazo superior a 30 (trinta) dias, na forma da lei, bem como a inscri&ccedil;&atilde;o da d&iacute;vida no pertinente cadastro.</p><p>II - Valer-se de firma especializada ou de profissionais de advocacia para efetuar a cobran&ccedil;a de seu cr&eacute;dito, seja extrajudicial ou judicial, sendo que o CONTRATANTE responder&aacute; pelos custos e honor&aacute;rios advocat&iacute;cios devidos.</p><p>III - N&atilde;o renovar a matr&iacute;cula do benefici&aacute;rio do CONTRATANTE, para o per&iacute;odo letivo posterior, caso este n&atilde;o tenha cumprido rigorosamente as cl&aacute;usulas do presente contrato.</p><p><b>CL&Aacute;USULA NONA: </b>Poder&atilde;o existir, a crit&eacute;rio do CONTRATADO, a extin&ccedil;&atilde;o de turmas, a mudan&ccedil;a de sede ou unidade de ensino, o agrupamento de classes, altera&ccedil;&otilde;es dos hor&aacute;rios de aulas e turmas, do calend&aacute;rio escolar e outras medidas que sejam necess&aacute;rias por raz&otilde;es de ordem administrativa, econ&ocirc;mico-financeira ou pedag&oacute;gicas.</p><p><b>Par&aacute;grafo &Uacute;nico - </b>Na hip&oacute;tese de inexist&ecirc;ncia ou de n&atilde;o forma&ccedil;&atilde;o de Turma, o CONTRATADO restituir&aacute; ao CONTRATANTE o valor pecuni&aacute;rio da fra&ccedil;&atilde;o da anuidade paga.</p><p><b>CL&Aacute;USULA D&Eacute;CIMA: O CONTRATANTE que pretenda rescindir o presente contrato dever&aacute; avisar previamente o CONTRATADO com 30 dias de anteced&ecirc;ncia e estar em dia com suas obriga&ccedil;&otilde;es financeiras at&eacute; o momento efetivo do t&eacute;rmino do aviso.</b></p></div>" 		
			
			Text =replace_latin_char(Text,"html")
		
			Text =replace_latin_char(Text,"html")
			
			Do While Len(Text) > 0
				CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
			 
				If CharsPrinted = Len(Text) Then Exit Do
													
				Text = Right( Text, Len(Text) - CharsPrinted)
			Loop 
				

	SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000;html=true")
'	Relatorio = "Sistema Web Diretor - SWD025"
	Relatorio = arquivo
	Do While Len(Relatorio) > 0
		CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
	 
		If CharsPrinted = Len(Relatorio) Then Exit Do
		   SET Page = Page.NextPage
		Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
	Loop 				
				
			SET Param_data_hora = Pdf.CreateParam("x="&margem&";y="&y_posicao&"; height="&y_diponivel&"; width="&area_utilizavel&"; alignment=center; size=8; color=#000000; html=true")

			Param_data_hora.Add "y="&margem+12&";html=true" 
			
			data_hora = "<div align=""Right"">"&paginacao&"<br>Impresso em "&data &" &agrave;s "&horario&"</div>"
			Do While Len(data_hora) > 0
				CharsPrinted = Page.Canvas.DrawText(data_hora, Param_data_hora, Font )			
				If CharsPrinted = Len(data_hora) Then Exit Do
				SET Page = Page.NextPage
				data_hora = Right( data_hora, Len(data_hora) - CharsPrinted)
			Loop 				
'FIM DA SEGUNDA P&Aacute;GINA=========================================================================================================================================				
			
				paginacao = paginacao+1			
				SET Page = Page.NextPage			
				y_posicao=842
	'CABE&Ccedil;ALHO==========================================================================================		
				Set Param_Logo_Gde = Pdf.CreateParam
	
				largura_logo_gde=formatnumber(Logo.Width*0.4,0)
				altura_logo_gde=formatnumber(Logo.Height*0.4,0)
				margem=30	
				linha=10
				area_utilizavel=Page.Width-(margem*2)
				Param_Logo_Gde("x") = formatnumber(area_utilizavel/2,0)
				Param_Logo_Gde("y") = Page.Height - altura_logo_gde -margem
				Param_Logo_Gde("ScaleX") = 0.4
				Param_Logo_Gde("ScaleY") = 0.4
				Page.Canvas.DrawImage Logo, Param_Logo_Gde
		
				x_texto=largura_logo_gde+ margem+10
				y_posicao=y_posicao - margem
				y_posicao=y_posicao - altura_logo_gde-linha
				
				SET Param = Pdf.CreateParam("x="&margem&";y="&y_posicao&"; height=20; width="&area_utilizavel&"; alignment=center; size=14; color=#000000; html=true")
				
				Text = "<center><b><font style=""font-size:10pt;"">CONTRATO DE PRESTA&Ccedil;&Atilde;O DE SERVI&Ccedil;OS DE EDUCA&Ccedil;&Atilde;O ESCOLAR</FONT></b></center>" 			
				
				Do While Len(Text) > 0
					CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
				 
					If CharsPrinted = Len(Text) Then Exit Do
						SET Page = Page.NextPage
					Text = Right( Text, Len(Text) - CharsPrinted)
				Loop 
				
				y_posicao=y_posicao- 30
				
				y_diponivel=y_posicao							
'	
	'FIM DO CABE&Ccedil;ALHO==========================================================================================	
																								
			y_diponivel=y_posicao				
			
			SET Param = Pdf.CreateParam("x="&margem&";y="&y_posicao&"; height="&y_diponivel&"; width="&area_utilizavel&"; alignment=center; size=10; color=#000000; html=true")
			
			Text = "<div align=""justify""><p><b>Par&aacute;grafo 1&ordm; - </b>A rescis&atilde;o do contrato se formaliza, t&atilde;o-somente, por meio de requerimento pr&oacute;prio, dispon&iacute;vel na Secretaria, e com a liquida&ccedil;&atilde;o das parcelas devidas, incluindo a parcela referente ao m&ecirc;s da rescis&atilde;o.</p><p><b>Par&aacute;grafo 2&ordm; - </b>Qualquer das partes s&oacute; poder&aacute; rescindir o presente contrato at&eacute; 30 (trinta) de setembro de 2010, em conformidade com a legisla&ccedil;&atilde;o pertinente.</p><p><b>CL&Aacute;USULA D&Eacute;CIMA PRIMEIRA: </b>O aluno que comprometer o nome ou reputa&ccedil;&atilde;o do Estabelecimento de Ensino, praticando atos de indisciplina, dar&aacute; causa &agrave; sua exclus&atilde;o, nos termos previstos no Regimento Interno.</p><p><b>CL&Aacute;USULA D&Eacute;CIMA SEGUNDA: </b>A rescis&atilde;o unilateral do presente contrato pelo CONTRATADO at&eacute; 30 (trinta) dias antes no in&iacute;cio do ano letivo ou indeferimento do &quot;requerimento de matr&iacute;cula&quot; implicar&aacute; na restitui&ccedil;&atilde;o integral das parcelas antecipadamente pagas pelo CONTRATANTE.</p><p><b>CL&Aacute;USULA D&Eacute;CIMA TERCEIRA: </b>As partes comprometem-se a comunicar, reciprocamente, por escrito e mediante recibo, qualquer mudan&ccedil;a de endere&ccedil;o, sob pena de serem consideradas v&aacute;lidas as correspond&ecirc;ncias enviadas aos endere&ccedil;os constantes do presente instrumento, inclusive para os efeitos de cita&ccedil;&atilde;o judicial.</p><p><b>Par&aacute;grafo &Uacute;nico - </b>A rescis&atilde;o unilateral do presente contrato pelo CONTRATANTE at&eacute; 30 (trinta) dias antes do in&iacute;cio do ano letivo faculta-lhe o recebimento de 50% (cinq&uuml;enta por cento) de tudo o que tiver sido pago, em at&eacute; 15 (quinze) dias ap&oacute;s a data do requerimento de cancelamento por escrito.</p><p><b>CL&Aacute;USULA D&Eacute;CIMA QUARTA: Fica assegurada ao aluno, em caso de &oacute;bito do CONTRATANTE, a continuidade da presta&ccedil;&atilde;o de servi&ccedil;o, at&eacute; o final do ano letivo em curso, sem &ocirc;nus no que diz respeito &agrave;s parcelas da anuidade, a partir da apresenta&ccedil;&atilde;o do documento e desde que n&atilde;o constem d&eacute;bitos anteriores.</b></p><p><b>CL&Aacute;USULA D&Eacute;CIMA QUINTA: </b>O CONTRATADO, livre de quaisquer &ocirc;nus para com o CONTRATANTE/aluno, poder&aacute; utilizar-se da imagem deste, para fins exclusivos de divulga&ccedil;&atilde;o da Escola e suas atividades, podendo, para tanto, reproduzi-la ou divulg&aacute;-la junto &agrave; internet, jornais e todos os demais meios de comunica&ccedil;&atilde;o, p&uacute;blicos ou privados, n&atilde;o podendo, em nenhuma hip&oacute;tese, a imagem ser utilizada de maneira contr&aacute;ria &agrave; moral, aos bons costumes ou &agrave; ordem p&uacute;blica.</p><p><b>CL&Aacute;USULA D&Eacute;CIMA SEXTA: </b>O presente contrato tem validade at&eacute; o t&eacute;rmino do ano letivo de 2010, sendo que a presente matr&iacute;cula n&atilde;o ser&aacute; renovada automaticamente. O CONTRATANTE dever&aacute; solicitar nova matr&iacute;cula para o ano seguinte, desde que tenha cumprido as cl&aacute;usulas deste contrato.</p><p><b>CL&Aacute;USULA D&Eacute;CIMA S&Eacute;TIMA: </b>Fica eleito o foro Regional da Barra da Tijuca Comarca da Capital, no Estado do Rio de Janeiro, como &uacute;nico competente para dirimir quest&otilde;es oriundas do presente instrumento.</p><p></p><p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;E por estarem as partes em acordo com todos os termos e condi&ccedil;&otilde;es do presente contrato, assinam o mesmo em 02 (duas) vias de igual teor e forma, para um s&oacute; efeito legal, juntamente com 02 (duas) testemunhas.</p></div>" 			
			
			Text =replace_latin_char(Text,"html")
		
			Text =replace_latin_char(Text,"html")
			
			Do While Len(Text) > 0
				CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
			 
				If CharsPrinted = Len(Text) Then Exit Do
													
				Text = Right( Text, Len(Text) - CharsPrinted)
			Loop 
			
			y_posicao=230
			
			SET Param_data_local = Pdf.CreateParam("x=180;y="&y_posicao&"; height="&y_diponivel&"; width="&area_utilizavel&"; alignment=center; size=8; color=#000000; html=true")
			data_local="Rio de Janeiro, _______ de ______________________de ___________"
			
			Do While Len(data_local) > 0
				CharsPrinted = Page.Canvas.DrawText(data_local, Param_data_local, Font )			
				If CharsPrinted = Len(data_local) Then Exit Do
				data_hora = Right( data_local, Len(data_local) - CharsPrinted)
			Loop 	
			
			y_posicao=200			
			
			SET Param_assinatura = Pdf.CreateParam("x=205;y="&y_posicao&"; height=10; width="&area_utilizavel&"; alignment=center; size=8; color=#000000; html=true")
			assinatura="______________________________________________"
			
			Do While Len(assinatura) > 0
				CharsPrinted = Page.Canvas.DrawText(assinatura, Param_assinatura, Font )			
				If CharsPrinted = Len(assinatura) Then Exit Do
				assinatura = Right( assinatura, Len(assinatura) - CharsPrinted)
			Loop 
			y_posicao=190				
			SET Param_assinatura = Pdf.CreateParam("x=275;y="&y_posicao&"; height=20; width="&area_utilizavel&"; alignment=center; size=8; color=#000000; html=true")
			assinatura="CONTRATANTE"
			
			Do While Len(assinatura) > 0
				CharsPrinted = Page.Canvas.DrawText(assinatura, Param_assinatura, Font )			
				If CharsPrinted = Len(assinatura) Then Exit Do
				assinatura = Right( assinatura, Len(assinatura) - CharsPrinted)
			Loop 								
			
			y_posicao=160			
			
			SET Param_assinatura = Pdf.CreateParam("x=205;y="&y_posicao&"; height=10; width="&area_utilizavel&"; alignment=center; size=8; color=#000000; html=true")
			assinatura="______________________________________________"
			
			Do While Len(assinatura) > 0
				CharsPrinted = Page.Canvas.DrawText(assinatura, Param_assinatura, Font )			
				If CharsPrinted = Len(assinatura) Then Exit Do
				assinatura = Right( assinatura, Len(assinatura) - CharsPrinted)
			Loop 
			y_posicao=150				
			SET Param_assinatura = Pdf.CreateParam("x=205;y="&y_posicao&"; height=40; width=200; alignment=center; size=8; color=#000000; html=true")
			assinatura="<center>Col&eacute;gio Saint John Ltda<br> CONTRATADO</center>"
			
			Do While Len(assinatura) > 0
				CharsPrinted = Page.Canvas.DrawText(assinatura, Param_assinatura, Font )			
				If CharsPrinted = Len(assinatura) Then Exit Do
				assinatura = Right( assinatura, Len(assinatura) - CharsPrinted)
			Loop 			
			
			
			y_posicao=120			
			
			SET Param_assinatura = Pdf.CreateParam("x="&margem&";y="&y_posicao&"; height=100; width="&area_utilizavel&"; alignment=center; size=8; color=#000000; html=true")
			assinatura="<p>TESTEMUNHAS:</p><p>1&ordf; Testemunha:<br>CPF:</p><p>2&ordf; Testemunha:<br>CPF:</p>"
			
			Do While Len(assinatura) > 0
				CharsPrinted = Page.Canvas.DrawText(assinatura, Param_assinatura, Font )			
				If CharsPrinted = Len(assinatura) Then Exit Do
				assinatura = Right( assinatura, Len(assinatura) - CharsPrinted)
			Loop 			
			
			y_posicao=100	

			Set param_table1 = Pdf.CreateParam("width=200; height=50; rows=1; cols=1; border=0.5; cellborder=0.5; cellspacing=0;")
			Set Table = Doc.CreateTable(param_table1)
			Table.Font = Font

			Table(1, 1).AddText "<center><p>Nesta data, declaro que recebi uma via do<br>presente contrato</p><p>Ass: ______________________________<br>CONTRATANTE</p></b></center>", "size=8;html=true", Font 				
			Page.Canvas.DrawTable Table, "x=365, y="&y_posicao&"" 



			
			
			SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000;html=true")
		'	Relatorio = "Sistema Web Diretor - SWD025"
			Relatorio = arquivo
			Do While Len(Relatorio) > 0
				CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
			 
				If CharsPrinted = Len(Relatorio) Then Exit Do
				   SET Page = Page.NextPage
				Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
			Loop 				
				
				
			SET Param_data_hora = Pdf.CreateParam("x="&margem&";y="&y_posicao&"; height="&y_diponivel&"; width="&area_utilizavel&"; alignment=center; size=8; color=#000000; html=true")
			
			Param_data_hora.Add "y="&margem+12&";html=true" 
			
			data_hora = "<div align=""Right"">"&paginacao&"<br>Impresso em "&data &" &agrave;s "&horario&"</div>"
			Do While Len(data_hora) > 0
				CharsPrinted = Page.Canvas.DrawText(data_hora, Param_data_hora, Font )			
				If CharsPrinted = Len(data_hora) Then Exit Do
				SET Page = Page.NextPage
				data_hora = Right( data_hora, Len(data_hora) - CharsPrinted)
			Loop 			

	NEXT			
			
			
	Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
end if
%>

