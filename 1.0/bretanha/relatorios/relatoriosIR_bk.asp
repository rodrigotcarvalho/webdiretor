<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 30 'valor em segundos

ano_letivo_ir = session("ano_letivo")
ano_IR=ano_letivo_ir
if instr(1,request.ServerVariables("URL"),"/wf")>0 then
	ano_IR=ano_IR-1
	session("ano_letivo")=ano_IR
end if
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/funcoes2.asp"-->

<% 
response.Charset="ISO-8859-1"
opt= request.QueryString("opt")


nivel=4
permissao = session("permissao") 
session("ano_letivo")=ano_letivo_ir
sistema_local=session("sistema_local")
nvg=session("nvg")
session("nvg")=nvg
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 


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
	'Instancia o objeto na memória
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
		response.Write(CAMINHO_bl)
		Set CONBL = Server.CreateObject("ADODB.Connection") 
		ABRIRBL = "DBQ="& CAMINHO_bl & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONBL.Open ABRIRBL

		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0	
		
	
if opt="01" then
	cod_cons=request.QueryString("cod_cons")
	co_declaracao= request.QueryString("dcl")	

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_IR &" AND CO_Matricula ="& cod_cons
		RS1.Open SQL1, CON1
		
		if RS1.EOF then
			tx_declaracao = "Aluno "&cod_cons&" n&atilde;o encontrado no ano letivo "&ano_IR
		
		end if
		If Not IsArray(alunos_encontrados) Then 
			alunos_encontrados = Array() 		
		End if	
		ReDim preserve alunos_encontrados(UBound(alunos_encontrados)+1)	
		alunos_encontrados(Ubound(alunos_encontrados)) = cod_cons	

	
elseif opt="02" then
	obr=request.QueryString("obr")
	dados_informados = split(obr, "_" )
	declaracao_terceiro_ano=request.QueryString("dcl")
	if isnull(declaracao_terceiro_ano) or declaracao_terceiro_ano="" then
		gera_declaracao_terceiro_ano="nao"
		unidade=dados_informados(0)
		curso=dados_informados(1)
		co_etapa=dados_informados(2)
		turma=dados_informados(3)
		co_declaracao=dados_informados(4)
	else
		gera_declaracao_terceiro_ano="sim"
		co_declaracao=declaracao_terceiro_ano
	end if
	IF ((isnull(unidade) or unidade="") and (isnull(curso) or curso="") and (isnull(co_etapa) or co_etapa="") and (isnull(turma) or turma="")) or gera_declaracao_terceiro_ano="sim" THEN
		if co_declaracao="swd204" or co_declaracao="swd209"  then
			nu_chamada_check = 1
			Set RSA = Server.CreateObject("ADODB.Recordset")
			SQL_BUSCA_ALUNOS= "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_IR &" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.CO_Curso = '2' AND TB_Matriculas.CO_Etapa = '3' order by TB_Matriculas.NU_Unidade,TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma, TB_Alunos.NO_Aluno"		
			Set RSA = CON1.Execute(SQL_BUSCA_ALUNOS)
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
		else
			response.redirect(origem&"index.asp?nvg="&nvg&"&opt=err3")				
		END IF		
	ELSE	
	
		if unidade="999990" or unidade="" or isnull(unidade) then
			SQL_BUSCA_ALUNOS="NULO"
		else	
			SQL_ALUNOS= "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_IR &" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.NU_Unidade = "& unidade		
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
	END IF	
	if vetor_matriculas="" then
		alunos_encontrados = Array() 
	else
		alunos_encontrados = split(vetor_matriculas, "#!#" )	
	end if	
end if	

if ubound(alunos_encontrados)>-1 then					

	relatorios_gerados=-1			
	For i=0 to ubound(alunos_encontrados)	
	
		SET Page = Doc.Pages.Add( 595, 842 )	
		cod_cons=alunos_encontrados(i)
		gera_relatorio_aluno="s"	
		
				
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_IR &" AND CO_Matricula ="& cod_cons
		RS1.Open SQL1, CON1
		
		if NOT RS1.EOF then
		
			ano_aluno = RS1("NU_Ano")
			rematricula = RS1("DA_Rematricula")
			situacao = RS1("CO_Situacao")
			encerramento= RS1("DA_Encerramento")
			unidade= RS1("NU_Unidade")
			curso= RS1("CO_Curso")
			co_etapa= RS1("CO_Etapa")
			turma= RS1("CO_Turma")
			cham= RS1("NU_Chamada")
			curso=curso*1
			if curso=2 then
				etapa=etapa*1
			end if					
			
			
			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL2 = "SELECT * FROM TB_Unidade WHERE NU_Unidade="& unidade
			RS2.Open SQL2, CON0
							
			nome_unidade = RS2("TX_Imp_Cabecalho")	
				
			rua_unidade = RS2("NO_Logradouro")		
			numero_unidade = RS2("NU_Logradouro")	
			complemento_unidade = RS2("TX_Complemento_Logradouro")	
			cep_unidade = RS2("CO_CEP")	
			bairro_unidade = RS2("CO_Bairro")	
			co_municipio_unidade = RS2("CO_Municipio")		
			tel_unidade = RS2("NUS_Telefones")	
			uf_unidade = RS2("SG_UF")	
			un_cnpj = RS2("CO_CGC")					
	
	
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
			
			if co_municipio_unidade="" or isnull(co_municipio_unidade) or uf_unidade_municipio="" or isnull(uf_unidade_municipio)then
			else
				Set RS3m = Server.CreateObject("ADODB.Recordset")
				SQL3m = "SELECT * FROM TB_Municipios WHERE SG_UF='"& uf_unidade_municipio &"' AND CO_Municipio="&co_municipio_unidade
				RS3m.Open SQL3m, CON0
				
				municipio_unidade=RS3m("NO_Municipio")	
				
				
					if bairro_unidade="" or isnull(bairro_unidade)then
					else
									
						Set RS3b = Server.CreateObject("ADODB.Recordset")
						SQL3b = "SELECT * FROM TB_Bairros WHERE SG_UF='"& uf_unidade_municipio &"' AND CO_Municipio="&co_municipio_unidade&" AND CO_Bairro = "&bairro_unidade
						RS3b.Open SQL3b, CON0				
					
						bairro_unidade= RS3b("NO_Bairro")
					end if				
				
									
			end if
			endereco_unidade=rua_unidade&numero_unidade&complemento_unidade&"<br>"&bairro_unidade&cep_unidade&" "&municipio_unidade&uf_unidade&"<br>Tel(s): "&tel_unidade	&"<br>CNPJ: "&un_cnpj		
			
						
			
			 gera_relatorio_aluno="s"

			
			if gera_relatorio_aluno="s" then
				Set RS = Server.CreateObject("ADODB.Recordset")
				SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod_cons
				RS.Open SQL, CON1
				
				nome_aluno = RS("NO_Aluno")
				sexo_aluno = RS("IN_Sexo")
				nome_pai = RS("NO_Pai")
				nome_mae = RS("NO_Mae")
				responsavel_financ = RS("TP_Resp_Fin")				
				
				Set RSdn = Server.CreateObject("ADODB.Recordset")
				SQLdn = "SELECT * FROM TB_Contatos WHERE CO_Matricula ="& cod_cons&" AND TP_Contato='ALUNO'"
				RSdn.Open SQLdn, CONCONT	
				dt_nascimento=RSdn("DA_Nascimento_Contato")
				
				if nome_aluno="" or isnull(nome_aluno)  then
					nome_aluno="<i>N&atilde;o informado no cadastro</i>"
				else
					nome_aluno =replace_latin_char(nome_aluno,"html")
				end if	
				if nome_pai="" or isnull(nome_pai)  then
					nome_pai="<i>N&atilde;o informado no cadastro</i>"
				else
					nome_pai =replace_latin_char(nome_pai,"html")
				end if	
				if nome_mae="" or isnull(nome_mae)  then
					nome_mae="<i>N&atilde;o informado no cadastro</i>"
				else
					nome_mae =replace_latin_char(nome_mae,"html")
				end if	
				
				if sexo_aluno="F" then
					desinencia="a"
				else
					desinencia="o"
				end if
				
				Set RSdn = Server.CreateObject("ADODB.Recordset")
				SQLdn = "SELECT * FROM TB_Contatos WHERE CO_Matricula ="& cod_cons&" AND TP_Contato='"&responsavel_financ&"'"
				RSdn.Open SQLdn, CONCONT	
				nome_resp_financ=RSdn("NO_Contato")	
				
				if nome_resp_financ="" or isnull(nome_resp_financ)  then
					nome_resp_financ="<i>N&atilde;o informado no cadastro</i>"
				else
					nome_resp_financ =replace_latin_char(nome_resp_financ,"html")
				end if								
				
				

						
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL3 = "SELECT * FROM TB_Curso WHERE CO_Curso='"& curso &"'"
				RS3.Open SQL3, CON0
				
				no_curso = RS3("NO_Curso")
				no_abrv_curso = RS3("NO_Abreviado_Curso")
				co_concordancia_curso = RS3("CO_Conc")	
				
				Set rssr = Server.CreateObject("ADODB.Recordset")
				qlsr= "SELECT * FROM TB_Etapa where CO_Curso = '"& curso &"' and CO_Etapa = '"& co_etapa &"'"
				rssr.Open qlsr, CON0

				no_etapa= rssr("NO_Etapa")
		

				
				nome_curso= no_etapa&" "&co_concordancia_curso&" "&no_curso
				'no_etapa = no_etapa&" "&co_concordancia_curso&" "&no_abrv_curso	
				
											
	'CABEÇALHO==========================================================================================		
				Set Param_Logo_Gde = Pdf.CreateParam

				largura_logo_gde=formatnumber(Logo.Width*0.8,0)
altura_logo_gde=formatnumber(Logo.Height*0.8,0)
				margem=30	
				area_utilizavel=Page.Width-(margem*2)
				Param_Logo_Gde("x") = margem
				Param_Logo_Gde("y") = Page.Height - altura_logo_gde -22
				Param_Logo_Gde("ScaleX") = 0.8
Param_Logo_Gde("ScaleY") = 0.8
				Page.Canvas.DrawImage Logo, Param_Logo_Gde
		
				x_texto=largura_logo_gde+ margem+10
				y_texto=formatnumber(Page.Height - margem,0)
				width_texto=Page.Width -largura_logo_gde - 80
		
			
				SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
				Text = "<p><i><b>"&UCASE(nome_unidade)&"</b></i></p><br><font style=""font-size:10pt;"">"&endereco_unidade&"</FONT>" 			
				
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
		
		Page.Canvas.SetParams "LineWidth=2" 
		Page.Canvas.SetParams "LineCap=0" 
		altura_segundo_separador= Page.Height - altura_logo_gde-margem
		With Page.Canvas
		   .MoveTo margem, altura_segundo_separador
		   .LineTo area_utilizavel+margem, altura_segundo_separador
		   .Stroke
		End With 						
				
				y_titulo=y_texto-altura_logo_gde-margem		
				
	
				
				x_declaracao= margem*3
				y_declaracao=y_titulo - (margem*5)
				width_declaracao=Page.Width - (margem*6)		
				

				tx_padrao = "N&atilde;o foi encontrada quita&ccedil;&atilde;o para "&desinencia&" alun"&desinencia&" <b>"&nome_aluno&"</b>, nascid"&desinencia&" em "&dt_nascimento&", filh"&desinencia&" de "&nome_pai&" e de "&nome_mae&", cursou neste Estabelecimento de Ensino o "&nome_curso&" referente ao ano letivo de "&ano_IR	

				
				size_assinatura = 15
				assinatura = "<center>Escola Bretanha</center>"
				if co_declaracao="EIR"  then
				
					arquivo = "SWD115"
					nome_documento = "Informe para Imposto de Renda"	

					
					x_data_extenso= margem*5
					y_data_extenso=y_declaracao - (margem*7)
					width_data_extenso=Page.Width - (margem*6)						
									
					SET Param_Data_Extenso = Pdf.CreateParam("x="&x_data_extenso&";y="&y_data_extenso&"; height=50; width="&width_data_extenso&"; size=13; alignment=Left; html=True")					
					
					Set RSQ = Server.CreateObject("ADODB.Recordset")
					SQLQ = "SELECT * FROM TB_NF_Ano_Anterior WHERE NU_Ano = "&ano_IR&" AND CO_Matricula  ="& cod_cons
					RSQ.Open SQLQ, CONBL					
					
					if not RSQ.EOF then
					
					tx_declaracao="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Declaramos, que "&desinencia&" alun"&desinencia&" <b>"&nome_aluno&"</b>, nascid"&desinencia&" em "&dt_nascimento&", filh"&desinencia&" de "&nome_pai&" e de "&nome_mae&", cursou neste Estabelecimento de Ensino o "&nome_curso&" tendo o seu respons&aacute;vel pago a import&acirc;ncia de "&formatcurrency(RSQ("VA_Total_NF"))&", referente a anuidade escolar e servi&ccedil;os prestados no ano letivo de "&ano_IR&"."
					else
						tx_declaracao=tx_padrao				
					end if
				else
					arquivo = "SWD116"				
					nome_documento = "Declara&ccedil;&atilde;o de Quita&ccedil;&atilde;o Anual de D&eacute;bitos"
					
					x_data_extenso= margem
					y_data_extenso=y_titulo - (margem*1.5)
					width_data_extenso=Page.Width - (margem*2)	
					
					SET Param_Data_Extenso = Pdf.CreateParam("x="&x_data_extenso&";y="&y_data_extenso&"; height=50; width="&width_data_extenso&"; size=13; alignment=Left; html=True")		
					
				y_prezado=y_data_extenso-margem	
				SET Param_prezado = Pdf.CreateParam("x="&margem&";y="&y_prezado&"; height="&altura_logo_gde&"; width="&area_utilizavel&"; alignment=center; size=13; color=#000000; html=true")					
					
					tx_prezado="Prezado(a) Sr(a). <B>"&nome_resp_financ&"</B>,<BR>Respons&aacute;vel pel"&desinencia&" noss"&desinencia&" alun"&desinencia&" <B>"&nome_aluno&"</B><BR>Matr&iacute;cula: <B>"&cod_cons&"</B>."
					
				Do While Len(tx_prezado) > 0
					CharsPrinted = Page.Canvas.DrawText(tx_prezado, Param_prezado, Font )
				 
					If CharsPrinted = Len(tx_prezado) Then Exit Do
						SET Page = Page.NextPage
					tx_prezado = Right( tx_prezado, Len(tx_prezado) - CharsPrinted)
				Loop 						
					
					
					Set RSQ = Server.CreateObject("ADODB.Recordset")
					SQLQ = "SELECT * FROM TB_Quitacao WHERE CO_Matricula  ="& cod_cons
					RSQ.Open SQLQ, CONBL		
	
					
					if not RSQ.EOF then					
																							
						tx_declaracao="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Em aten&ccedil;&atilde;o &agrave; Lei Federal n&ordm; 12007/09, a Escola Bretanha declara que todas as mensalidades de presta&ccedil;&atilde;o de servi&ccedil;os do ano letivo de "&ano_IR&", relativas &agrave; matricula n&ordm; "&cod_cons&", encontram-se quitadas.<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Esta declara&ccedil;&atilde;o substitui as quita&ccedil;&otilde;es dos boletos mensais de d&eacute;bitos para efeito de comprova&ccedil;&atilde;o de pagamento. Ficam ressalvados nesta declara&ccedil;&atilde;o eventuais d&eacute;bitos decorrentes de processos judiciais ou administrativos cuja decis&atilde;o venha a ser favor&aacute;vel &agrave; Escola Bretanha e/ou os valores posteriormente apurados como devidos ou nos quais tenha havido estorno no processamento do pagamento.<BR>&nbsp;<BR>&nbsp;Atenciosamente,"
					
					else
						tx_declaracao=tx_padrao			
					end if
				end if						
				
				
				SET Param = Pdf.CreateParam("x="&margem&";y="&y_titulo&"; height="&altura_logo_gde&"; width="&area_utilizavel&"; alignment=center; size=17; color=#000000; html=true")
		Text = "<center><b><U>"&nome_documento&"</U></b></center>"
				
				
				Do While Len(Text) > 0
					CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
				 
					If CharsPrinted = Len(Text) Then Exit Do
						SET Page = Page.NextPage
					Text = Right( Text, Len(Text) - CharsPrinted)
				Loop 				
				
				

				
				SET Param_Declaracao = Pdf.CreateParam("x="&x_declaracao&";y="&y_declaracao&"; height=300; width="&width_declaracao&"; alignment=left; html=True")
				declaracao = "<div align=""justify""><font style=""font-size:15pt;"">"&tx_declaracao&"</font></div>"
				
				Do While Len(declaracao) > 0
					CharsPrinted = Page.Canvas.DrawText(declaracao, Param_Declaracao, Font )
				 
					If CharsPrinted = Len(declaracao) Then Exit Do
						SET Page = Page.NextPage
					declaracao = Right( declaracao, Len(declaracao) - CharsPrinted)
				Loop 
				
				Do While Len(data_extenso) > 0
					CharsPrinted = Page.Canvas.DrawText(data_extenso, Param_Data_Extenso, Font )
				 
					If CharsPrinted = Len(data_extenso) Then Exit Do
						SET Page = Page.NextPage
					data_extenso = Right( data_extenso, Len(data_extenso) - CharsPrinted)
				Loop 	
									
				
					
			    y_assinatura=Y_declaracao-300	
	

				 SET Param_Relatorio = Pdf.CreateParam("x="&margem*5&";y="&y_assinatura&"; height=50; width="&Page.Width - (margem*10)&"; alignment=left; size="&size_assinatura&"; color=#000000;html= true")


				Do While Len(assinatura) > 0
					CharsPrinted = Page.Canvas.DrawText(assinatura, Param_Relatorio, Font )
				 
					If CharsPrinted = Len(assinatura) Then Exit Do
					   SET Page = Page.NextPage
					assinatura = Right( assinatura, Len(assinatura) - CharsPrinted)
				Loop 
									
					
				Page.Canvas.SetParams "LineWidth=1" 
				Page.Canvas.SetParams "LineCap=0" 
				inicio_primeiro_separador=largura_logo_gde+margem+10
				altura_primeiro_separador= Page.Height - margem - 17
				With Page.Canvas
				   .MoveTo margem, margem
				   .LineTo area_utilizavel+margem, margem
				   .Stroke
				End With 					
				
		
				 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width=50; alignment=left; size=8; color=#000000")

				Relatorio = arquivo
				Do While Len(Relatorio) > 0
					CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
				 
					If CharsPrinted = Len(Relatorio) Then Exit Do
					   SET Page = Page.NextPage
					Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
				Loop 
				
				
		SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
						
				Param_Relatorio.Add "alignment=right" 		
		
				data_hora = "Impresso em "&data &", "&horario&""
				Do While Len(Relatorio) > 0
					CharsPrinted = Page.Canvas.DrawText(data_hora, Param_Relatorio, Font )			
					If CharsPrinted = Len(data_hora) Then Exit Do
					SET Page = Page.NextPage
					data_hora = Right( data_hora, Len(data_hora) - CharsPrinted)
				Loop 				
				
			END IF 	
		else
			SET Param_Declaracao = Pdf.CreateParam("x=50;y=800; height=300; width=500; alignment=left; html=True")
				declaracao = "<div align=""justify""><font style=""font-size:15pt;"">"&tx_declaracao&"</font></div>"
				
				Do While Len(declaracao) > 0
					CharsPrinted = Page.Canvas.DrawText(declaracao, Param_Declaracao, Font )
				 
					If CharsPrinted = Len(declaracao) Then Exit Do
						SET Page = Page.NextPage
					declaracao = Right( declaracao, Len(declaracao) - CharsPrinted)
				Loop 	
				arquivo = "NaoLocalizado"
		End IF
		if gera_relatorio_aluno="s" then
			relatorios_gerados=relatorios_gerados+1
		else
			relatorios_gerados=relatorios_gerados
		end if
	Next	
	
	Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   

end if
%>

