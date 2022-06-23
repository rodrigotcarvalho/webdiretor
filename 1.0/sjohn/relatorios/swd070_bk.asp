<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 60 'valor em segundos
'Emitir Classificação por média
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/calculos.asp"-->
<!--#include file="../inc/parametros.asp"-->
<!--#include file="../inc/bd_grade.asp"-->
<!--#include file="../../global/funcoes_diversas.asp"-->
<% 
arquivo="SWD070"

response.Charset="ISO-8859-1"
dados= request.QueryString("obr")
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

informacoes=request.QueryString("inf")
dados_query=split(informacoes,"$!$")
unidade = dados_query(0)
curso = dados_query(1)
co_etapa = dados_query(2)
periodo = dados_query(4)
tp_ord = dados_query(5)



'unidade=request.form("unidade")
'curso=request.form("curso")
'co_etapa=request.form("etapa")
'periodo=request.form("media")
'tp_ord=request.form("tp_ord")
'response.Write(tp_ord)
if tp_ord="C" then
	ordenacao = "CRESCENTE"
	sql_ordena = "ASC"
ELSE
	ordenacao = "DECRESCENTE"
	sql_ordena = "DESC"	
END IF	
if mes<10 then
mes="0"&mes
end if

data = dia &"/"& mes &"/"& ano

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

		Set CONt = Server.CreateObject("ADODB.Connection") 
		ABRIRt = "DBQ="& CAMINHO_t & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONt.Open ABRIRt

		if isnumeric(periodo) then
				Set RSp = Server.CreateObject("ADODB.Recordset")
				SQLp = "SELECT * FROM TB_Periodo where NU_Periodo="&periodo
				RSp.Open SQLp, CON0
					
			nome_periodo = RSp("NO_Periodo")
			nome_periodo=replace_latin_char(nome_periodo,"html")
			tipo_calculo = "MB"								
		else			
			nome_periodo = funcao_vetor_medias("I", periodo, outro)
			tipo_calculo = "MA"			
		end if

	if unidade="nulo" or unidade="" or isnull(unidade) then
		SQL_ALUNOS="Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada, TB_Matriculas.NU_Unidade, TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula"
	else	
		SQL_ALUNOS= "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada, TB_Matriculas.NU_Unidade, TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.NU_Unidade = "& unidade		
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
	SQL_ALUNOS= SQL_ALUNOS&SQL_CURSO&SQL_ETAPA&" AND CO_Situacao = 'C' order by TB_Matriculas.NU_Unidade,TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma, TB_Alunos.NO_Aluno"
	end if

	if SQL_ALUNOS="NULO" then
	else
	
		nu_chamada_check = 1
		mudou_turma = "N"		
		Set RSA = Server.CreateObject("ADODB.Recordset")
		CONEXAOA = SQL_ALUNOS
'		response.Write(CONEXAOA)
		Set RSA = CON1.Execute(CONEXAOA)
		vetor_matriculas="" 
			ctrl_unidade = 999999
			ctrl_curso = "ctrl"		
			ctrl_etapa = "ctrl"					
		While Not RSA.EOF
			nu_matricula = RSA("CO_Matricula")
			nu_chamada = RSA("NU_Chamada")
			db_unidade = RSA("NU_Unidade")	
			db_curso = RSA("CO_Curso")		
			db_etapa = RSA("CO_Etapa")		
			db_turma = RSA("CO_Turma")	
			
'response.Write(	nu_matricula&"-"&ctrl_unidade&"-"&ctrl_curso&"-"&ctrl_etapa&"-"&dbctrl_turmaturma&"|"&db_unidade&"-"&db_curso&"-"&db_etapa&"-"&db_turma&"<BR>")	
			if isnumeric(ctrl_unidade) then
				ctrl_unidade = ctrl_unidade*1
			end if			
			if isnumeric(ctrl_curso) then
				ctrl_curso = ctrl_curso*1
			end if
			if isnumeric(ctrl_etapa) then
				ctrl_etapa = ctrl_etapa*1
			end if
			if isnumeric(ctrl_turma) then
				ctrl_turma = ctrl_turma*1
			end if
			if isnumeric(db_unidade) then
				db_unidade = db_unidade*1
			end if
			if isnumeric(db_curso) then
				db_curso = db_curso*1
			end if
			if isnumeric(db_etapa) then
				db_etapa = db_etapa*1
			end if
			
			if nu_chamada_check <> 1 and (ctrl_unidade<>db_unidade or ctrl_curso<>db_curso or ctrl_etapa<>db_etapa) then
				ctrl_unidade = db_unidade
				ctrl_curso = db_curso	
				ctrl_etapa = db_etapa		
				
				vetor_matriculas = 	vetor_matriculas&"#$#"	
				vetor_turmas = vetor_turmas&"#$#"&db_unidade&"#!#"&db_curso&"#!#"&db_etapa&"#!#"&db_turma	
				mudou_turma = "S" 
			end if
																						
															
			if nu_chamada_check = 1 and ctrl_curso = "ctrl"	 then
				vetor_matriculas=nu_matricula
				ctrl_unidade = db_unidade
				ctrl_curso = db_curso	
				ctrl_etapa = db_etapa		
				ctrl_turma = db_turma					
				vetor_turmas = db_unidade&"#!#"&db_curso&"#!#"&db_etapa&"#!#"&db_turma				
			elseif mudou_turma = "S"  then			
				vetor_matriculas=vetor_matriculas&nu_matricula
				mudou_turma = "N" 	
			else
				vetor_matriculas=vetor_matriculas&"#!#"&nu_matricula
			end if
		nu_chamada_check=nu_chamada_check+1			
		RSA.MoveNext
		Wend 
	
	end if	

'	matriculas_encontradas = split(vetor_matriculas, "#!#" )	
'response.Write(vetor_turmas&"<br>")
'response.Write(vetor_matriculas&"<br>")
'response.Write("-------------------------------<br>")
'response.End()
	alunos_turmas_encontradas = split(vetor_matriculas, "#$#" )	
	turmas_encontradas = split(vetor_turmas, "#$#" )
	RSA.Close
	Set RSA = Nothing
'end if


	SET Page = Doc.Pages.Add(595,842)		
For trms=0 to ubound(turmas_encontradas)
	vetor_materias=""
	vetor_tipo_materia=""
	co_mat_cons=""
	tp_mat_cons=""			
	vetor_medias=""
	vetor_aluno_media=""
	media_ordenada = ""
	aluno_ordenada = ""	
	alunos_pesquisados=""
	medias_pesquisadas=""		
	if trms>0 then
		SET Page = Page.NextPage
	end if		
		

	ucet = split(turmas_encontradas(trms), "#!#" )		
	
	nome_unidade = verifica_nome(ucet(0),variavel_2,variavel_3,variavel_4,variavel_5,CON0,"u", "f")
	nome_curso = verifica_nome(variavel_1,ucet(1),variavel_3,variavel_4,variavel_5,CON0,"c", "f")
	no_abrv_curso = verifica_nome(variavel_1,ucet(1),ucet(2),variavel_4,variavel_5,CON0,"c", "a")
	co_concordancia_curso = verifica_nome(variavel_1,ucet(1),ucet(2),variavel_4,variavel_5,CON0,"cc", "f")
	nome_etapa = verifica_nome(variavel_1,ucet(1),ucet(2),variavel_4,variavel_5,CON0,"e", "f")	

	Set RS2 = Server.CreateObject("ADODB.Recordset")
	SQL2 = "SELECT * FROM TB_Unidade WHERE NU_Unidade="& ucet(0)
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
		Set RS3m = Server.CreateObject("ADODB.Recordset")
		SQL3m = "SELECT * FROM TB_Municipios WHERE SG_UF='"& uf_unidade_municipio &"' AND CO_Municipio="&municipio_unidade
		RS3m.Open SQL3m, CON0
		cod_cidade= municipio_unidade
		municipio_unidade=RS3m("NO_Municipio")	
		
		if bairro_unidade="" or isnull(bairro_unidade)then
		else
		
			Set RS3m = Server.CreateObject("ADODB.Recordset")
			SQL3m = "SELECT * FROM TB_Bairros WHERE SG_UF='"& uf_unidade_municipio &"' AND CO_Municipio="&cod_cidade&" and CO_Bairro = "&bairro_unidade
			RS3m.Open SQL3m, CON0	
			bairro_unidade=RS3m("NO_Bairro")				
			bairro_unidade=" - "&bairro_unidade
		end if									
	end if
	endereco_unidade=rua_unidade&numero_unidade&complemento_unidade&bairro_unidade&cep_unidade&"<br>"&municipio_unidade&uf_unidade		
	
	
'CABEÇALHO==========================================================================================		
	Set Param_Logo_Gde = Pdf.CreateParam
	margem=30				
	largura_logo_gde=formatnumber(Logo.Width*0.5,0)
	altura_logo_gde=formatnumber(Logo.Height*0.5,0)
	area_utilizavel=Page.Width - (margem*2)	
	y_logo_grande = Page.Height - altura_logo_gde -22
	Param_Logo_Gde("x") = margem
	Param_Logo_Gde("y") = y_logo_grande
	Param_Logo_Gde("ScaleX") = 0.5
	Param_Logo_Gde("ScaleY") = 0.5
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
			   .LineTo area_utilizavel, altura_primeiro_separador
			   .Stroke
			End With 					


	y_texto=altura_primeiro_separador-30
	width_texto=580-x_texto


	SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
	Text = "<p><center><i><b><font style=""font-size:12pt;"">Classifica&ccedil;&atilde;o dos Alunos por M&eacute;dias - Em ordem "&ordenacao&"</font></b></i></center></p>"
	

	Do While Len(Text) > 0
		CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
	 
		If CharsPrinted = Len(Text) Then Exit Do
			SET Page = Page.NextPage
		Text = Right( Text, Len(Text) - CharsPrinted)
	Loop 
	
	no_unidade = ucet(0)&" - "&nome_unidade
	no_curso= nome_etapa&" "&co_concordancia_curso&" "&nome_curso
	
	Page.Canvas.SetParams "LineWidth=2" 
	Page.Canvas.SetParams "LineCap=0"
	y_primeira_linha = y_texto-30 
	With Page.Canvas
		.MoveTo margem, y_primeira_linha
		.LineTo 570, y_primeira_linha
		.Stroke
	End With 	
	
	
	Set param_table1 = Pdf.CreateParam("width=540; height=25; rows=2; cols=4; border=0; cellborder=0; cellspacing=0;")
	Set Table = Doc.CreateTable(param_table1)
	Table.Font = Font
	y_table=y_primeira_linha-5
	
	With Table.Rows(1)
	   .Cells(1).Width = 40
   	   .Cells(2).Width = 340   
	   .Cells(3).Width = 60
	   .Cells(4).Width = 100      
	End With
	Table(1, 1).AddText "Unidade:", "size=9;", Font 
	Table(2, 1).AddText "Curso:", "size=9;", Font 
	Table(1, 2).AddText no_unidade, "size=9;", Font 
	Table(2, 2).AddText no_curso, "size=9;", Font 
	Table(1, 3).AddText "<div align=""right"">M&eacute;dia:</div>", "size=9; alignment=right;html=true", Font 
	Table(1, 4).AddText "<div align=""right"">"&nome_periodo&"</div>", "size=9;alignment=right;html=true", Font 				
	Table(2, 3).AddText "Ano Letivo: ", "size=9; alignment=right", Font 
	Table(2, 4).AddText ano_letivo, "size=9;alignment=right", Font 
	Page.Canvas.DrawTable Table, "x="&margem&", y="&y_table&"" 

	y_segunda_linha = y_table-30
	With Page.Canvas
	   .MoveTo margem, y_segunda_linha
	   .LineTo 570, y_segunda_linha
	   .Stroke
	End With 	
	
'================================================================================================================			

	
	tb_notas = tabela_notas(CON2, ucet(0), ucet(1), ucet(2), ucet(3), nulo, nulo, nulo)	
	
	caminho_n= caminho_notas(CON2, tb_notas, outro)
	
	Set CON_N = Server.CreateObject("ADODB.Connection") 
	ABRIRn = "DBQ="& caminho_n & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_N.Open ABRIRn		
	
	
	tp_modelo=tipo_divisao_ano(ucet(1),ucet(2),"tp_modelo")
	tp_freq=tipo_divisao_ano(ucet(1),ucet(2),"in_frequencia")

	
	Set RSNN = Server.CreateObject("ADODB.Recordset")
	CONEXAONN = "Select * from TB_Programa_Aula WHERE CO_Curso = '"& ucet(1) &"' AND CO_Etapa = '"& ucet(2)&"' and CO_Materia not in ('EF', 'EFIS') order by NU_Ordem_Boletim"
	Set RSNN = CON0.Execute(CONEXAONN)	
				
	nu_materia_check = 0
	while not RSNN.EOF
		curso=curso*1
		if isnumeric(co_etapa) then
			co_etapa=co_etapa*1
		end if	
		co_materia= RSNN("CO_Materia")	
		in_mae= RSNN("IN_MAE")
		in_fil= RSNN("IN_FIL")	
		if (co_materia="ARTC" and curso=1 and co_etapa<6) then
		else
		
			if in_mae=TRUE AND in_fil=TRUE then
				wrk_tipo_materia="mcf"
			elseif in_mae=TRUE AND in_fil=FALSE then	
				wrk_tipo_materia="msf"	
			elseif in_mae=FALSE AND in_fil=TRUE then
				wrk_tipo_materia="f"
			else				
				response.Write(co_materia&" - ERRO na classifica&ccedil;&atilde;o do tipo de mat&eacute;ria")
				response.End()
			end if		
			
			if wrk_tipo_materia<>"f" then	
				if nu_materia_check = 0 then
					vetor_materias=co_materia
					vetor_tipo_materia=wrk_tipo_materia
				else
					vetor_materias=vetor_materias&"#!#"&co_materia
					vetor_tipo_materia=vetor_tipo_materia&"#!#"&wrk_tipo_materia
				end if
			end if	
			nu_materia_check=nu_materia_check+1			
		end if	
	RSNN.MOVENEXT
	wend
	
	co_mat_cons=split(vetor_materias,"#!#")
	tp_mat_cons=split(vetor_tipo_materia,"#!#")	
	
	aluno=split(alunos_turmas_encontradas(trms),"#!#")
	
	nu_media_check = 1
	for aln=1 to ubound(aluno)
		nu_matricula=aluno(aln)
		acumula_aluno=0		
		media_numerador=0
		media_denominador=0
		pula_aluno="N"		
				
		for q=0 to ubound(co_mat_cons)
			disciplina=co_mat_cons(q)
			tp_disc=tp_mat_cons(q)	

			if tipo_calculo = "MA" then		
				resultado=Calc_Med_An_Fin(ucet(0), ucet(1), ucet(2), ucet(3), nu_matricula, disciplina, caminho_n, tb_notas, 6, 5, 6,"anual", 0)	
				'response.Write(periodo&" "&pula_aluno&" "&caminho_n&" "&nu_matricula&" "&disciplina&" "&resultado&"<BR>")
				if resultado="&nbsp;#!#&nbsp;" then
					pula_aluno="S"		
				else
					dados_resultado=split(resultado,"#!#")	
					media_aluno = dados_resultado(0)
					media_denominador=media_denominador+1
				end if	
				'response.Write(periodo&" "&pula_aluno&" "&caminho_n&" "&nu_matricula&" "&disciplina&" "&resultado&"<BR>")						
			else	
				if tp_disc="msf" then
	
					media_aluno=Calcula_Media_T_F_F_N(ucet(0), ucet(1), ucet(2), ucet(3), tp_modelo, tp_freq, nu_matricula, disciplina, disciplina, CON_N , tb_notas, periodo, "VA_Media2", outro)				
					if media_aluno="&nbsp;" then
						pula_aluno="S"				
					else
						media_denominador=media_denominador+1
					end if		
					'response.Write(nu_matricula&"|"&disciplina&"|"&media_aluno&"|"&pula_aluno&"<BR>")			
				elseif	tp_disc="mcf" then
					conta_filhas = 0
					
					Set RSNNa = Server.CreateObject("ADODB.Recordset")
					CONEXAONNa = "Select * from TB_Materia WHERE CO_Materia_Principal = '"& disciplina &"' order by NU_Ordem_Boletim"
					Set RSNNa = CON0.Execute(CONEXAONNa)	
					
					if RSNNa.EOF then
						Response.Write("Erro ao localizar mat&eacute;rias filhas para "&disciplina&" em TB_Materia")
						Response.end()
					else	
						while not RSNNa.EOF
						
							filha=RSNNa("CO_Materia")
			
							if conta_filhas = 0 then
								vetor_filhas=filha
							else
								vetor_filhas=vetor_filhas&"#!#"&filha
							end if
							
							conta_filhas=conta_filhas+1	
						RSNNa.MOVENEXT
						wend	
					end if	
				
					'media_aluno=Calcula_Media_T_T_F_N(unidade, curso, co_etapa, turma, nu_matricula, disciplina, vetor_filhas, CAMINHOn, notaFIL, periodo)	
					media_aluno=Calcula_Media_T_T_F_N(ucet(0), ucet(1), ucet(2), ucet(3), tp_modelo, tp_freq, nu_matricula, disciplina, vetor_filhas, CON_N, tb_notas, periodo, "VA_Media2", outro)				
					if media_aluno="&nbsp;" then
						pula_aluno="S"				
					else
						media_denominador=media_denominador+1
					end if		
					'response.Write(nu_matricula&":"&disciplina&":"&media_aluno&":"&pula_aluno&"<BR>")	
				else
					media_aluno=0
					media_denominador=media_denominador	
				end if	
			end if
			
			if media_aluno="&nbsp;" or media_aluno="" or isnull(media_aluno) then
				media_aluno=0	
				pula_aluno="S"						
			end if	
			media_numerador=media_numerador*1
			media_aluno=media_aluno*1
			acumula_aluno=acumula_aluno+media_aluno
		NEXT
		if pula_aluno="N" then
			media_numerador=acumula_aluno			

			if media_denominador=0 then
				media=media_numerador
			else
				media=media_numerador/media_denominador		
			end if
	
			media=formatnumber(media,1)
			
			if nu_media_check = 1 then
				vetor_medias=media
				vetor_aluno_media=nu_matricula
			else
			'response.Write(vetor_medias&"-"&media&"<BR>")			
				vetor_medias=vetor_medias&";"&media
				vetor_aluno_media=vetor_aluno_media&";"&nu_matricula
			end if
			nu_media_check=nu_media_check+1			
		end if			
	next
'faixa1=0
'faixa2=0
'faixa3=0
'faixa4=0
'faixa5=0
'response.Write(vetor_aluno_media&"<BR>")		
vetor_medias=split(vetor_medias,";")
vetor_aluno_media=split(vetor_aluno_media,";")



'for n=0 to ubound(vetor_medias)
'	analisa_media=vetor_medias(n)
'	analisa_media=analisa_media*1
'
'	if analisa_media>80 then
'		faixa5=faixa5+1
'		alunos_faixa5=alunos_faixa5&"#!#"&vetor_aluno_media(n)
'		medias_faixa5=medias_faixa5&"#!#"&analisa_media
'		
'	elseif analisa_media>60 then
'		faixa4=faixa4+1
'		alunos_faixa4=alunos_faixa4&"#!#"&vetor_aluno_media(n)
'		medias_faixa4=medias_faixa4&"#!#"&analisa_media
'
'	elseif analisa_media>40 then
'		faixa3=faixa3+1
'		alunos_faixa3=alunos_faixa3&"#!#"&vetor_aluno_media(n)
'		medias_faixa3=medias_faixa3&"#!#"&analisa_media
'
'	elseif analisa_media>20 then
'		faixa2=faixa2+1
'		alunos_faixa2=alunos_faixa2&"#!#"&vetor_aluno_media(n)
'		medias_faixa2=medias_faixa2&"#!#"&analisa_media
'
'	else
'		faixa1=faixa1+1
'		alunos_faixa1=alunos_faixa1&"#!#"&vetor_aluno_media(n)
'		medias_faixa1=medias_faixa1&"#!#"&analisa_media
'
'	end if					
'
'next
'session("faixas")=faixa1&"#!#"&faixa2&"#!#"&faixa3&"#!#"&faixa4&"#!#"&faixa5
'session("categorias")="0-20#!#21-40#!#41-60#!#61-80#!#81-100"
'faixas=session("faixas")
'categorias=session("categorias")
'
'classes=split(categorias,"#!#")
'response.Write(ubound(classes))

'for y=0 to ubound(classes)
'
'	nomes = classes(y)
'	faixa=y
'	faixa=faixa*1
'	faixa_origem=faixa_origem*1
'	
'	
'	if faixa_origem=0 then
'		alunos_pesquisados=alunos_faixa1
'		medias_pesquisadas=medias_faixa1		
'		
'	elseif faixa_origem=1 then	
'		alunos_pesquisados=alunos_faixa2
'		medias_pesquisadas=medias_faixa2	
'
'	elseif faixa_origem=2 then	
'		alunos_pesquisados=alunos_faixa3	
'		medias_pesquisadas=medias_faixa3
'			
'	elseif faixa_origem=3 then	
'		alunos_pesquisados=alunos_faixa4
'		medias_pesquisadas=medias_faixa4
'		
'	else	
'		alunos_pesquisados=alunos_faixa5
'		medias_pesquisadas=medias_faixa5
'		
'	end if			
'next

for n=0 to ubound(vetor_medias)
	if n=0 then
		alunos_pesquisados=vetor_aluno_media(n)
		medias_pesquisadas=vetor_medias(n)		
	else
		alunos_pesquisados=alunos_pesquisados&"#!#"&vetor_aluno_media(n)
		medias_pesquisadas=medias_pesquisadas&"#!#"&vetor_medias(n)	
	end if	
next	

altura_medias=20
y_medias=y_segunda_linha-10	
Set param_table2 = Pdf.CreateParam("width=540; height="&altura_medias&"; rows=1; cols=4; border=1; cellborder=0.1; cellspacing=0; x="&margem&"; y="&y_medias&"; MaxHeight=640")
Set Notas_Tit = Doc.CreateTable(param_table2)
Notas_Tit.Font = Font


With Notas_Tit.Rows(1)
   .Cells(1).Width = 60
   .Cells(2).Width = 340
   .Cells(3).Width = 60   
   .Cells(4).Width = 80	   
End With
Notas_Tit(1, 1).AddText "<div align=""center""><b>Matr&iacute;cula</b></div>", "size=10;indenty=3; html=true", Font 
Notas_Tit(1, 2).AddText "<div align=""center""><b>Nome do Aluno</b></div>", "size=10;indenty=3; html=true", Font 
Notas_Tit(1, 3).AddText "<div align=""center""><b>Turma</b></div>", "size=10;indenty=3; html=true", Font 
Notas_Tit(1, 4).AddText "<div align=""center""><b>M&eacute;dia</b></div>", "size=8;indenty=3;html=true", Font  
param_table2.Add "indenty=3;alignment=right;html=true"


aluno_ordena=split(alunos_pesquisados,"#!#")
medias_ordena=split(medias_pesquisadas,"#!#")

if ubound(aluno_ordena)=-1 then
	Set Row = Notas_Tit.Rows.Add(15) ' row height
	altura_medias=altura_medias+20


	param_table2.Add "indentx=2"	
	Row.Cells(1).ColSpan = 4		
	Row.Cells(1).AddText "<div align=""Center""><font style=""font-size:8pt;"">N&atilde;o existem alunos com notas dispon&iacute;veis para c&aacute;lculo da m&eacute;dia solicitada</font></div>", param_table2	
	param_table2.Add "indentx=0"	
else
' ordena vetor médias
	
	AuxVetor = medias_ordena
	AuxVetor_aluno = aluno_ordena
	for x=0 to ubound(medias_ordena)
		Valor=-1 
		menor_media=101
		maior_media=0		
		if tp_ord="C" then	
			for y=0 to ubound(AuxVetor)  
				media_teste=AuxVetor(y)
				media_teste=media_teste*1
				menor_media=menor_media*1
				maior_media=maior_media*1
				Valor=Valor*1
				y=y*1

					if media_teste<=menor_media and y<>Valor then
						menor_media=media_teste
						aluno_menor_media=AuxVetor_aluno(y)
						Valor=y
					end if
			next
			media_ordenada=media_ordenada&"#!#"&menor_media
			aluno_ordenada=aluno_ordenada&"#!#"&aluno_menor_media
		else
			for y=0 to ubound(AuxVetor)  
				media_teste=AuxVetor(y)
				media_teste=media_teste*1
				maior_media=maior_media*1
				Valor=Valor*1
				y=y*1
response.Write(AuxVetor_aluno(y)&"|"&media_teste&">"&maior_media&"&"&y&"<>"&Valor&"<BR>")
					if media_teste>=maior_media and y<>Valor then
						maior_media=media_teste
						aluno_maior_media=AuxVetor_aluno(y)
						Valor=y
					end if			
			next
			media_ordenada=media_ordenada&"#!#"&maior_media
			aluno_ordenada=aluno_ordenada&"#!#"&aluno_maior_media			
		end if
response.Write(aluno_ordenada&"<BR>")		

' Retirando o menor ou maior elemento do vetor das médias
			Dim tmpvetor
			tmpvetor = array()
			for z = LBound( AuxVetor ) to UBound ( AuxVetor )
				z=z*1
				Valor=Valor*1
				if z <> Valor then
					Redim preserve tmpvetor ( UBound(tmpvetor)+1 ) 
					tmpvetor ( UBound ( tmpvetor ) ) = AuxVetor( z )
				end if
			next
			AuxVetor = tmpvetor 'salvando novamente a Array
			tmpvetor = array() 'liberando a var tmp 

' Retirando o dono da menor ou maior média do vetor dos alunos
			Dim tmpvetor_aluno
			tmpvetor_aluno = array()

			for z = LBound( AuxVetor_aluno) to UBound ( AuxVetor_aluno )
				z=z*1
				Valor=Valor*1
				if z <> Valor then
					Redim preserve tmpvetor_aluno ( UBound(tmpvetor_aluno)+1 ) ' adicionei um elemento
					tmpvetor_aluno ( UBound ( tmpvetor_aluno ) ) = AuxVetor_aluno( z )
				end if
			next
			AuxVetor_aluno = tmpvetor_aluno 'salvando novamente a Array
			tmpvetor_aluno = array() 'liberando a var tmp 


	next

medias_exibe=split(media_ordenada,"#!#")
aluno_exibe=split(aluno_ordenada,"#!#")
response.Write(aluno_ordenada)
'response.End()
	check = 1 
	for k=1 to ubound(aluno_exibe)  
	
		Set RSnome = Server.CreateObject("ADODB.Recordset")
		SQLnome = "SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.NU_Ano ="& ano_letivo &" AND TB_Matriculas.CO_Matricula ="& aluno_exibe(k) 
		RSnome.Open SQLnome, CON1

		nu_chamada = RSnome("NU_Chamada")
		nome_aluno = RSnome("NO_Aluno")
		turma_aluno = RSnome("CO_Turma")		
		nome_aluno=replace_latin_char(nome_aluno,"html")			
		media_exibe=formatnumber(medias_exibe(k),1)
		
		Set Row = Notas_Tit.Rows.Add(15) ' row height
		altura_medias=altura_medias+20


		param_table2.Add "indentx=0"						
		Row.Cells(1).AddText "<div align=""Center""><font style=""font-size:8pt;"">"&aluno_exibe(k)&"</font></div>", param_table2	
		param_table2.Add "indentx=2"					
		Row.Cells(2).AddText "<div align=""Left""><font style=""font-size:8pt;"">"&nome_aluno&"</font></div>", param_table2	
		param_table2.Add "indentx=0"									
		Row.Cells(3).AddText "<div align=""Center""><font style=""font-size:8pt;"">"&turma_aluno&"</font></div>", param_table2					
		Row.Cells(4).AddText "<div align=""Center""><font style=""font-size:8pt;"">"&media_exibe&"</font></div>", param_table2	
		check=check+1
	Next
 end if   



'Page.Canvas.DrawTable Notas_Tit, "x="&margem&", y="&y_medias&"" 				
	Paginacao = 0	
	limite=0
	Do While True
		limite=limite+1
		Paginacao=Paginacao+1
	   LastRow = Page.Canvas.DrawTable( Notas_Tit, param_table2 )
	
'				if LastRow >= Notas_Tit.Rows.Count Then 
'			    	Exit Do ' entire table displayed
'				else
		 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width=530; alignment=left; size=8; color=#000000")
		
		Relatorio = arquivo&" - Sistema Web Diretor"
		Do While Len(Relatorio) > 0
			CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
		 
			If CharsPrinted = Len(Relatorio) Then Exit Do
			   SET Page = Page.NextPage
			Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
		Loop 
		
		Param_Relatorio.Add "alignment=right" 
		

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
		
		if LastRow >= Notas_Tit.Rows.Count Then 
			Exit Do ' entire table displayed
		end if			
		Set Page = Page.NextPage	
		param_table2.Add( "RowTo=1; RowFrom=1" ) ' Row 1 is header.
		param_table2("RowFrom1") = LastRow + 1 ' RowTo1 is omitted and presumed infinite
'NOVO CABEÇALHO==========================================================================================			
	Set Param_Logo_Gde = Pdf.CreateParam
	margem=30				
	largura_logo_gde=formatnumber(Logo.Width*0.5,0)
	altura_logo_gde=formatnumber(Logo.Height*0.5,0)
	area_utilizavel=Page.Width - (margem*2)	
	y_logo_grande = Page.Height - altura_logo_gde -22
	Param_Logo_Gde("x") = margem
	Param_Logo_Gde("y") = y_logo_grande
	Param_Logo_Gde("ScaleX") = 0.5
	Param_Logo_Gde("ScaleY") = 0.5
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
			   .LineTo area_utilizavel, altura_primeiro_separador
			   .Stroke
			End With 					


	y_texto=altura_primeiro_separador-30
	width_texto=580-x_texto


	SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
	Text = "<p><center><i><b><font style=""font-size:12pt;"">Classifica&ccedil;&atilde;o dos Alunos por M&eacute;dias - Em ordem "&ordenacao&"</font></b></i></center></p>"
	

	Do While Len(Text) > 0
		CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
	 
		If CharsPrinted = Len(Text) Then Exit Do
			SET Page = Page.NextPage
		Text = Right( Text, Len(Text) - CharsPrinted)
	Loop 
	
	no_unidade = ucet(0)&" - "&nome_unidade
	no_curso= nome_etapa&" "&co_concordancia_curso&" "&nome_curso
	
	Page.Canvas.SetParams "LineWidth=2" 
	Page.Canvas.SetParams "LineCap=0"
	y_primeira_linha = y_texto-30 
	With Page.Canvas
		.MoveTo margem, y_primeira_linha
		.LineTo 570, y_primeira_linha
		.Stroke
	End With 	
	
	
	Set param_table1 = Pdf.CreateParam("width=540; height=25; rows=2; cols=4; border=0; cellborder=0; cellspacing=0;")
	Set Table = Doc.CreateTable(param_table1)
	Table.Font = Font
	y_table=y_primeira_linha-5
	
	With Table.Rows(1)
	   .Cells(1).Width = 40
   	   .Cells(2).Width = 340   
	   .Cells(3).Width = 60
	   .Cells(4).Width = 100      
	End With
	Table(1, 1).AddText "Unidade:", "size=9;", Font 
	Table(2, 1).AddText "Curso:", "size=9;", Font 
	Table(1, 2).AddText no_unidade, "size=9;", Font 
	Table(2, 2).AddText no_curso, "size=9;", Font 
	Table(1, 3).AddText "<div align=""right"">M&eacute;dia:</div>", "size=9; alignment=right;html=true", Font 
	Table(1, 4).AddText "<div align=""right"">"&nome_periodo&"</div>", "size=9;alignment=right;html=true", Font 				
	Table(2, 3).AddText "Ano Letivo: ", "size=9; alignment=right", Font 
	Table(2, 4).AddText ano_letivo, "size=9;alignment=right", Font 
	Page.Canvas.DrawTable Table, "x="&margem&", y="&y_table&"" 

	y_segunda_linha = y_table-30
	With Page.Canvas
	   .MoveTo margem, y_segunda_linha
	   .LineTo 570, y_segunda_linha
	   .Stroke
	End With 	
	
'================================================================================================================	
		 				
		'end if			
		if limite>100 then
		response.Write("ERRO!")
		response.end()
		end if 
	Loop			
	Set param_table3 = Pdf.CreateParam("width=533; height=20; rows=2; cols=10; border=0; cellborder=0; cellspacing=0;")
	Set Legenda = Doc.CreateTable(param_table3)
	Legenda.Font = Font
	y_legenda=y_medias-altura_medias
	'response.Write(altura_medias)
	'response.end()
	With Legenda.Rows(1)
	   .Cells(1).Width = 40
	   .Cells(2).Width = 20
	   .Cells(3).Width = 40
	   .Cells(4).Width = 20
	   .Cells(5).Width = 40
	   .Cells(6).Width = 20
	   .Cells(7).Width = 40
	   .Cells(8).Width = 20 
	   .Cells(9).Width = 43 
	   .Cells(10).Width = 250             
	End With
	data_exibe = data&" &agrave;s "& horario

	Legenda(1, 1).Colspan= 8	
	'Legenda(1, 10).RowSpan = 2					
	Legenda(1, 10).AddText "<b><Div align=""right"">Documento impresso em: "&data_exibe&"</div></b>", "size=8; html=true", Font 				
	Page.Canvas.DrawTable Legenda, "x="&margem&", y="&y_legenda&"" 
			
			'				
	 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width=50; alignment=left; size=8; color=#000000")
	Relatorio = arquivo
	Do While Len(Relatorio) > 0
		CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
	 
		If CharsPrinted = Len(Relatorio) Then Exit Do
		   SET Page = Page.NextPage
		Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
	Loop 	
	CON_N.Close	
	Set CON_N = Nothing									
Next		


Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

