<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 30 'valor em segundos
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<% 
arquivo="SWD125"
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
origem="../ws/mat/man/edf/"
end if


paginacao=1

if mes<10 then
mes="0"&mes
end if

data = dia &"/"& mes &"/"& ano

if mes=1 then
	mes_extenso="Janeiro"
elseif mes=2 then
	mes_extenso="Fevereiro"
elseif mes=3 then
	mes_extenso="Março"
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
	documento=dados_informados(4)
	
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

	SET Page = Doc.Pages.Add(595,842)
	y_posicao=842
	total_linhas=0		
	max_linhas=34
	qtd_turma=0
	primeiro_cabecalho="s"
	unidade_base= "nulo"
	curso_base= "nulo"
	co_etapa_base= "nulo"
	turma_base= "nulo"

	relatorios_gerados=-1			
	For i=0 to ubound(alunos_encontrados)	
		total_doc_aluno=0
		cod_cons=alunos_encontrados(i)
		gera_relatorio_aluno="s"			
				
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
	'			sexo_aluno = RS("IN_Sexo")
			
	'			Set RSdn = Server.CreateObject("ADODB.Recordset")
	'			SQLdn = "SELECT * FROM TB_Contatos WHERE CO_Matricula ="& cod_cons&" AND TP_Contato='ALUNO'"
	'			RSdn.Open SQLdn, CONCONT	
	'			dt_nascimento=RSdn("DA_Nascimento_Contato")
			
			if nome_aluno="" or isnull(nome_aluno)  then
				nome_aluno="<i>N&atilde;o informado no cadastro</i>"
			else
				nome_aluno =replace_latin_char(nome_aluno,"html")
			end if	
			
	'			if sexo_aluno="F" then
	'				desinencia="a"
	'			else
	'				desinencia="o"
	'			end if
							
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
			SQL4 = "SELECT * FROM TB_Etapa WHERE CO_Etapa ='"& co_etapa &"'"
			RS4.Open SQL4, CON0
			
			no_etapa = RS4("NO_Etapa")
							
			no_curso= no_etapa&"&nbsp;"&co_concordancia_curso&"&nbsp;"&no_curso
			texto_turma = "&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;Turma "&turma
			mensagem_turma="<b>Unidade: "&no_unidade_abr&"&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;"&no_curso&texto_turma&"</b>"
			mensagem_cabecalho=ano_letivo
			
	
			
			if primeiro_cabecalho="s" then
				primeiro_cabecalho="n"			
				
				'response.Write(paginacao&"<BR>")
	
	'CABEÇALHO==========================================================================================		
				Set Param_Logo_Gde = Pdf.CreateParam
	
				largura_logo_gde=formatnumber(Logo.Width*0.5,0)
altura_logo_gde=formatnumber(Logo.Height*0.5,0)
				margem=30	
				area_utilizavel=Page.Width-(margem*2)
				Param_Logo_Gde("x") = margem
				Param_Logo_Gde("y") = Page.Height - altura_logo_gde -margem
				Param_Logo_Gde("ScaleX") = 0.5
Param_Logo_Gde("ScaleY") = 0.5
				Page.Canvas.DrawImage Logo, Param_Logo_Gde
		
				x_texto=largura_logo_gde+ margem+10
				y_texto=y_posicao - margem
				
				y_posicao=y_texto
	
				
				width_texto=Page.Width -largura_logo_gde - 80
				
				SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
				Text = "<p><i><b>"&UCASE(no_unidade)&"</b></i></p><br><font style=""font-size:10pt;"">"&endereco_unidade&"</FONT>" 			
				
				Do While Len(Text) > 0
					CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
				 
					If CharsPrinted = Len(Text) Then Exit Do
						SET Page = Page.NextPage
					Text = Right( Text, Len(Text) - CharsPrinted)
				Loop 
				
				y_posicao=y_posicao- 17
				Page.Canvas.SetParams "LineWidth=1" 
				Page.Canvas.SetParams "LineCap=0" 
				inicio_primeiro_separador=largura_logo_gde+margem+10
				altura_primeiro_separador= y_posicao
				With Page.Canvas
				   .MoveTo inicio_primeiro_separador, altura_primeiro_separador
				   .LineTo area_utilizavel+margem, altura_primeiro_separador
				   .Stroke
				End With 	
	
				y_posicao=y_posicao-margem					
		
	
				
				SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_posicao&"; height=30; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
				Text = "<center><i><b><font style=""font-size:18pt;"">DOCUMENTOS FALTANTES</font></b></i></center>"
		
		
				Do While Len(Text) > 0
					CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
				 
					If CharsPrinted = Len(Text) Then Exit Do
						SET Page = Page.NextPage
					Text = Right( Text, Len(Text) - CharsPrinted)
				Loop 
	
				y_posicao=y_posicao-altura_logo_gde+30
									
				Page.Canvas.SetParams "LineWidth=2" 
				Page.Canvas.SetParams "LineCap=0" 
				altura_segundo_separador= y_posicao-5
				With Page.Canvas
				   .MoveTo margem, altura_segundo_separador
				   .LineTo area_utilizavel+margem, altura_segundo_separador
				   .Stroke
				End With 	
				y_posicao=y_posicao-10
				
				Set param_table1 = Pdf.CreateParam("width="&area_utilizavel&"; height=20; rows=1; cols=3; border=0; cellborder=0; cellspacing=0;")
				Set Table = Doc.CreateTable(param_table1)
				Table.Font = Font
				y_primeira_tabela=y_posicao
				With Table.Rows(1)
				   .Cells(1).Width = 50			   		   		   
				   .Cells(2).Width = area_utilizavel-200
				   .Cells(3).Width = 150	
				End With
				
				Table(1, 1).AddText "<b>Ano Letivo:</b>", "size=8;html=true", Font 
				Table(1, 2).AddText mensagem_cabecalho, "size=8;html=true", Font	
				Page.Canvas.DrawTable Table, "x="&margem&", y="&y_primeira_tabela&"" 	
				y_posicao=y_posicao-20
	
	'FIM DO CABEÇALHO==========================================================================================	
			end if
		
				
			
			if regera_cabecalho_ucet="s" then
				if qtd_turma>1 then
					Page.Canvas.SetParams "LineWidth=2" 
					Page.Canvas.SetParams "LineCap=0" 
					inicio_primeiro_separador=largura_logo_gde+margem+10
					altura_primeiro_separador= y_posicao
					With Page.Canvas
					   .MoveTo margem, y_posicao
					   .LineTo area_utilizavel+margem, y_posicao
					   .Stroke
					End With 
					y_posicao=y_posicao-20	
				end if
			

				if documento="D99999" then
					SET param_cabecalho_ucet = Pdf.CreateParam("x="&margem&";y="&y_posicao&"; height=20; width="&area_utilizavel&"; alignment=Left; size=10; color=#000000; html=true")	
				
					cabecalho_ucet = "<strong>"&mensagem_turma&"</strong>"
					Do While Len(cabecalho_ucet) > 0
						CharsPrinted = Page.Canvas.DrawText(cabecalho_ucet, param_cabecalho_ucet, Font )			
						If CharsPrinted = Len(cabecalho_ucet) Then Exit Do
						SET Page = Page.NextPage
						cabecalho_ucet = Right( cabecalho_ucet, Len(cabecalho_ucet) - CharsPrinted)
					Loop 						
					
					show_documento="s"
					redutor_y=20	
				else
					SET param_cabecalho_ucet = Pdf.CreateParam("x="&margem&";y="&y_posicao&"; height=20; width="&area_utilizavel&"; alignment=Left; size=10; color=#000000; html=true")					
					Set RSdt0 = Server.CreateObject("ADODB.Recordset")
					SQLdt0 = "SELECT * FROM TB_Documentos_Matricula where CO_Documento='"&documento&"' order by NO_Documento"			
					RSdt0.Open SQLdt0, CON0			
					no_doc_mat=RSdt0("NO_Documento")	
					cabecalho_ucet = "<strong>"&mensagem_turma&"</strong>"
					Do While Len(cabecalho_ucet) > 0
						CharsPrinted = Page.Canvas.DrawText(cabecalho_ucet, param_cabecalho_ucet, Font )			
						If CharsPrinted = Len(cabecalho_ucet) Then Exit Do
						SET Page = Page.NextPage
						cabecalho_ucet = Right( cabecalho_ucet, Len(cabecalho_ucet) - CharsPrinted)
					Loop 
					
					y_posicao=y_posicao-15
					
					param_cabecalho_ucet.Add "y="&y_posicao&";html=true" 
										
					cabecalho_ucet = "<strong>Alunos sem o documento: "&no_doc_mat&"</strong>"
					Do While Len(cabecalho_ucet) > 0
						CharsPrinted = Page.Canvas.DrawText(cabecalho_ucet, param_cabecalho_ucet, Font )			
						If CharsPrinted = Len(cabecalho_ucet) Then Exit Do
						SET Page = Page.NextPage
						cabecalho_ucet = Right( cabecalho_ucet, Len(cabecalho_ucet) - CharsPrinted)
					Loop 															
					show_documento="n"	
					redutor_y=15	
				end if
	

					
				y_posicao=y_posicao-redutor_y	
				Page.Canvas.SetParams "LineWidth=1" 
				Page.Canvas.SetParams "LineCap=0" 
				inicio_primeiro_separador=largura_logo_gde+margem+10
				altura_primeiro_separador= y_posicao
				With Page.Canvas
				   .MoveTo margem, y_posicao
				   .LineTo area_utilizavel+margem, y_posicao
				   .Stroke
				End With 						
				
				y_posicao=y_posicao-5
				
				total_linhas=total_linhas*1	
				total_linhas=total_linhas+1							
			end if	
	
			Set RSdt = Server.CreateObject("ADODB.Recordset")
			if show_documento="s" then
				SQLdt = "SELECT * FROM TB_Documentos_Matricula order by NO_Documento"
			else
				SQLdt = "SELECT * FROM TB_Documentos_Matricula where CO_Documento='"&documento&"' order by NO_Documento"			
			end if
			RSdt.Open SQLdt, CON0
			
			x_doc=margem+20		
			while not RSdt.EOF
				co_doc_mat=RSdt("CO_Documento")
				no_doc_mat=RSdt("NO_Documento")
					
				Set RSde = Server.CreateObject("ADODB.Recordset")
				SQLde = "SELECT * FROM TB_Documentos_Entregues where CO_Documento='"&co_doc_mat&"' And CO_Matricula="&cod_cons
				RSde.Open SQLde, CON0
				
								
				IF RSde.EOF then					
					if total_doc_aluno=0 then
						SET param_aluno = Pdf.CreateParam("x="&margem&";y="&y_posicao&"; height=17; width="&area_utilizavel&"; alignment=Left; size=10; color=#000000; html=true")	
						aluno = cod_cons&" - "&nome_aluno
						Do While Len(aluno) > 0
							CharsPrinted = Page.Canvas.DrawText(aluno, param_aluno, Font )			
							If CharsPrinted = Len(aluno) Then Exit Do
							SET Page = Page.NextPage
							aluno = Right( aluno, Len(aluno) - CharsPrinted)
						Loop 				
											
						y_posicao=y_posicao-17	
						total_linhas=total_linhas*1				
						total_linhas=total_linhas+1									
					end if
					
					if 	show_documento="s" then
						SET param_doc = Pdf.CreateParam("x="&x_doc&";y="&y_posicao&"; height=17; width="&area_utilizavel&"; alignment=Left; size=8; color=#000000; html=true")	
						
						Do While Len(no_doc_mat) > 0
							CharsPrinted = Page.Canvas.DrawText(no_doc_mat, param_doc, Font )			
							If CharsPrinted = Len(no_doc_mat) Then Exit Do
							SET Page = Page.NextPage
							no_doc_mat = Right( no_doc_mat, Len(no_doc_mat) - CharsPrinted)
						Loop 	
											
						y_posicao=y_posicao-17					
						total_doc_aluno=total_doc_aluno*1
						total_doc_aluno=total_doc_aluno+1
						total_linhas=total_linhas*1	
						total_linhas=total_linhas+1		
					end if										
				END IF
				
				if total_linhas>max_linhas then
									
					SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
				'	Relatorio = "Sistema Web Diretor - SWD025"
					Relatorio = arquivo
					Do While Len(Relatorio) > 0
						CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
					 
						If CharsPrinted = Len(Relatorio) Then Exit Do
						   SET Page = Page.NextPage
						Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
					Loop 
					
					Param_Relatorio.Add "y="&margem+12&";html=true" 
					
					data_hora = "<div align=""Right"">"&paginacao&"<br>Impresso em "&data &" &agrave;s "&horario&"</div>"
					Do While Len(Relatorio) > 0
						CharsPrinted = Page.Canvas.DrawText(data_hora, Param_Relatorio, Font )			
						If CharsPrinted = Len(data_hora) Then Exit Do
						SET Page = Page.NextPage
						data_hora = Right( data_hora, Len(data_hora) - CharsPrinted)
					Loop 						
					
					
					SET Page = Page.NextPage
					paginacao = paginacao+1
					total_linhas=0	
					y_posicao=842
					regera_cabecalho="s"
		'NOVO CABEÇALHO==========================================================================================		
					Set Param_Logo_Gde = Pdf.CreateParam
		
					largura_logo_gde=formatnumber(Logo.Width*0.5,0)
altura_logo_gde=formatnumber(Logo.Height*0.5,0)
					margem=30	
					area_utilizavel=Page.Width-(margem*2)
					Param_Logo_Gde("x") = margem
					Param_Logo_Gde("y") = Page.Height - altura_logo_gde -margem
					Param_Logo_Gde("ScaleX") = 0.5
Param_Logo_Gde("ScaleY") = 0.5
					Page.Canvas.DrawImage Logo, Param_Logo_Gde
			
					x_texto=largura_logo_gde+ margem+10
					y_texto=y_posicao - margem
					
					y_posicao=y_texto
		
					
					width_texto=Page.Width -largura_logo_gde - 80
					
					SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
					Text = "<p><i><b>"&UCASE(no_unidade)&"</b></i></p><br><font style=""font-size:10pt;"">"&endereco_unidade&"</FONT>" 			
					
					Do While Len(Text) > 0
						CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
					 
						If CharsPrinted = Len(Text) Then Exit Do
							SET Page = Page.NextPage
						Text = Right( Text, Len(Text) - CharsPrinted)
					Loop 
					
					y_posicao=y_posicao- 17
					Page.Canvas.SetParams "LineWidth=1" 
					Page.Canvas.SetParams "LineCap=0" 
					inicio_primeiro_separador=largura_logo_gde+margem+10
					altura_primeiro_separador= y_posicao
					With Page.Canvas
					   .MoveTo inicio_primeiro_separador, altura_primeiro_separador
					   .LineTo area_utilizavel+margem, altura_primeiro_separador
					   .Stroke
					End With 	
		
					y_posicao=y_posicao-margem					
			
		
					
					SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_posicao&"; height=30; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
					Text = "<center><i><b><font style=""font-size:18pt;"">DOCUMENTOS FALTANTES</font></b></i></center>"
			
			
					Do While Len(Text) > 0
						CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
					 
						If CharsPrinted = Len(Text) Then Exit Do
							SET Page = Page.NextPage
						Text = Right( Text, Len(Text) - CharsPrinted)
					Loop 
		
					y_posicao=y_posicao-altura_logo_gde+30
										
					Page.Canvas.SetParams "LineWidth=2" 
					Page.Canvas.SetParams "LineCap=0" 
					altura_segundo_separador= y_posicao
					With Page.Canvas
					   .MoveTo margem, altura_segundo_separador
					   .LineTo area_utilizavel+margem, altura_segundo_separador
					   .Stroke
					End With 	
					y_posicao=y_posicao-10
					
					Set param_table1 = Pdf.CreateParam("width="&area_utilizavel&"; height=20; rows=1; cols=3; border=0; cellborder=0; cellspacing=0;")
					Set Table = Doc.CreateTable(param_table1)
					Table.Font = Font
					y_primeira_tabela=y_posicao
					With Table.Rows(1)
					   .Cells(1).Width = 50			   		   		   
					   .Cells(2).Width = area_utilizavel-200
					   .Cells(3).Width = 150	
					End With
					
					Table(1, 1).AddText "<b>Ano Letivo:</b>", "size=8;html=true", Font 
					Table(1, 2).AddText mensagem_cabecalho, "size=8;html=true", Font	
					Page.Canvas.DrawTable Table, "x="&margem&", y="&y_primeira_tabela&"" 	
					y_posicao=y_posicao-20
		
		'FIM DO CABEÇALHO==========================================================================================	
					
					
				end if			
			RSdt.MOVENEXT
			WEND				
		End if
	Next
	
	SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000;html=true")
'	Relatorio = "Sistema Web Diretor - SWD025"
	Relatorio = arquivo
	Do While Len(Relatorio) > 0
		CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
	 
		If CharsPrinted = Len(Relatorio) Then Exit Do
		   SET Page = Page.NextPage
		Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
	Loop 
	
	Param_Relatorio.Add "y="&margem+12&";html=true" 
	
	data_hora = "<div align=""Right"">"&paginacao&"<br>Impresso em "&data &" &agrave;s "&horario&"</div>"
	Do While Len(Relatorio) > 0
		CharsPrinted = Page.Canvas.DrawText(data_hora, Param_Relatorio, Font )			
		If CharsPrinted = Len(data_hora) Then Exit Do
		SET Page = Page.NextPage
		data_hora = Right( data_hora, Len(data_hora) - CharsPrinted)
	Loop 		 	
			
	Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
end if
%>

