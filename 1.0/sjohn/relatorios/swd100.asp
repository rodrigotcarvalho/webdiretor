<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 60 'valor em segundos
'Carômetro
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/funcoes3.asp"-->
<!--#include file="../../global/funcoes_diversas.asp"-->
<% 
arquivo="SWD100"

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
turma = dados_query(3)

obr=unidade&"_"&curso&"_"&co_etapa&"_"&turma
if co_etapa = "f0"then
co_etapa=0
elseif co_etapa = "f1" or co_etapa = "m1"then
co_etapa=1
elseif co_etapa = "f2" or co_etapa = "m2"then
co_etapa = 2
elseif co_etapa = "f3" or co_etapa = "m3"then
co_etapa = 3
elseif co_etapa = "f4" then
co_etapa = 4
elseif co_etapa = "f5" then
co_etapa = 5
elseif co_etapa = "f6" then
co_etapa = 6
elseif co_etapa = "f7" then
co_etapa = 7
elseif co_etapa = "f8" then
co_etapa = 8
elseif co_etapa = "f55" then
co_etapa = 55
elseif co_etapa = "f66" then
co_etapa = 66
elseif co_etapa = "f77" then
co_etapa = 77
end if

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

	if unidade="nulo" or unidade="" or isnull(unidade) then
		SQL_ALUNOS="Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada, TB_Matriculas.NU_Unidade, TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma, TB_Alunos.NO_Aluno from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula"
	else	
		SQL_ALUNOS= "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada, TB_Matriculas.NU_Unidade, TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma, TB_Alunos.NO_Aluno from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.NU_Unidade = "& unidade		
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
			SQL_TURMA=" AND TB_Matriculas.CO_Turma = '"& turma &"'"				
		end if			
	'SQL_ALUNOS= SQL_ALUNOS&SQL_CURSO&SQL_ETAPA&SQL_TURMA&" AND CO_Situacao = 'C' order by TB_Matriculas.NU_Unidade,TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma, TB_Alunos.NO_Aluno"
	SQL_ALUNOS= SQL_ALUNOS&SQL_CURSO&SQL_ETAPA&SQL_TURMA&" order by TB_Matriculas.NU_Unidade,TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma, TB_Matriculas.NU_Chamada"	
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
	width_texto=area_utilizavel-largura_logo_gde-margem-50


	SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
	Text = "<p><center><i><b><font style=""font-size:16pt;"">Car&ocirc;metro</font></b></i></center></p>"
	

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
	Table(1, 3).AddText "<div align=""right"">Turma</div>", "size=9; alignment=right;html=true", Font 
	Table(1, 4).AddText "<div align=""right"">"&turma&"</div>", "size=9;alignment=right;html=true", Font 				
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

	
	aluno=split(alunos_turmas_encontradas(trms),"#!#")
	
	nu_media_check = 1
	x_foto = margem
	y_foto = y_segunda_linha
	

Dim objFSO
'Create an instance of the FileSystemObject object
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	qtd_fotos=0
	linhas=1
	pagina=1
	for aln=0 to ubound(aluno)
		nu_matricula=aluno(aln)
		qtd_fotos = qtd_fotos+1
		

		Set RSN = Server.CreateObject("ADODB.Recordset")
		CONEXAON = "Select NO_Aluno from TB_Alunos WHERE CO_Matricula ="&nu_matricula
		Set RSN = CON1.Execute(CONEXAON)		
		
		
		pula_aluno="N"	

		nome_aluno = RSN("NO_Aluno")
					
		'Verifica se o arquivo existe
		if objFSO.FileExists(Server.MapPath( "../img/fotos/aluno/"&nu_matricula&".jpg")) then		
			Set Foto = Doc.OpenImage( Server.MapPath( "../img/fotos/aluno/"&nu_matricula&".jpg") )
		else
			Set Foto = Doc.OpenImage( Server.MapPath( "../img/fotos/aluno/sem_foto.jpg") )			
		end if
'response.Write(aln&" 3 "&nome_aluno&"<BR>")				
		Set Param_Foto = Pdf.CreateParam
		largura_Foto=formatnumber(Foto.Width*0.24,0)
		altura_Foto=formatnumber(Foto.Height*0.24,0)
		if qtd_fotos=1 then
			y_foto = y_foto-altura_Foto-10	
		end if
		Param_Foto("x") = x_foto
		Param_Foto("y") = y_foto
		Param_Foto("ScaleX") = 0.24
		Param_Foto("ScaleY") = 0.24
		Page.Canvas.DrawImage Foto, Param_Foto	
		SET Param_Nome = Pdf.CreateParam("x="&x_foto&";y="&y_foto&"; height=40; width="&largura_Foto&"; alignment=center; size=7; color=#000000")
		Do While Len(nome_aluno) > 0
			CharsPrinted = Page.Canvas.DrawText(nome_aluno, Param_Nome, Font )
		 
			If CharsPrinted = Len(nome_aluno) Then Exit Do
			SET Page = Page.NextPage
			nome_aluno = Right(nome_aluno, Len(nome_aluno) - CharsPrinted)
		Loop 			

		if qtd_fotos mod 6 = 0 then
			y_foto = y_foto-altura_Foto-margem	
			x_foto = margem		
			linhas = linhas+1					
		else						
			x_foto = x_foto+largura_Foto+margem
		end if		
		if linhas>5 then
			Set param_table3 = Pdf.CreateParam("width=533; height=20; rows=2; cols=10; border=0; cellborder=0; cellspacing=0;")
			Set Legenda = Doc.CreateTable(param_table3)
			Legenda.Font = Font
			y_legenda=y_medias-altura_medias
			'response.Write(altura_medias)
			'response.end()
			With Legenda.Rows(1)
			   .Cells(1).Width = 40
			   .Cells(2).Width = 20
			   .Cells(3).Width = 250
			   .Cells(4).Width = 20
			   .Cells(5).Width = 40
			   .Cells(6).Width = 20
			   .Cells(7).Width = 40
			   .Cells(8).Width = 20 
			   .Cells(9).Width = 43 
			   .Cells(10).Width = 40             
			End With
			data_exibe = data&" &agrave;s "& horario
		
			Legenda(1, 3).Colspan= 7	
			'Legenda(1, 10).RowSpan = 2					
			Legenda(1, 3).AddText "<Div align=""center"">Documento impresso em: "&data_exibe&"</div>", "size=8; html=true", Font 		
			Legenda(1, 10).AddText "<Div align=""right"">"&pagina&"</div>", "size=8; html=true", Font 				
			Page.Canvas.DrawTable Legenda, "x="&margem&", y="&margem&"" 
					
					'				
			 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width=50; alignment=left; size=8; color=#000000")
			Relatorio = arquivo
			Do While Len(Relatorio) > 0
				CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
			 
				If CharsPrinted = Len(Relatorio) Then Exit Do
				   SET Page = Page.NextPage
				Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
			Loop 	
		SET Page = Page.NextPage
' NOVO CABEÇALHO==========================================================================================		
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
	width_texto=area_utilizavel-largura_logo_gde-margem-50


	SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
	Text = "<p><center><i><b><font style=""font-size:16pt;"">Car&ocirc;metro</font></b></i></center></p>"
	

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
	Table(1, 3).AddText "<div align=""right"">Turma</div>", "size=9; alignment=right;html=true", Font 
	Table(1, 4).AddText "<div align=""right"">"&turma&"</div>", "size=9;alignment=right;html=true", Font 				
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
			x_foto = margem
			y_foto = y_segunda_linha-altura_Foto-10	
			pagina=pagina+1	
			linhas=1		
		end if				
	next
Set objFSO = nothing
'response.End()
	Set param_table3 = Pdf.CreateParam("width=533; height=20; rows=2; cols=10; border=0; cellborder=0; cellspacing=0;")
	Set Legenda = Doc.CreateTable(param_table3)
	Legenda.Font = Font
	y_legenda=y_medias-altura_medias
	'response.Write(altura_medias)
	'response.end()
	With Legenda.Rows(1)
	   .Cells(1).Width = 40
	   .Cells(2).Width = 20
	   .Cells(3).Width = 250
	   .Cells(4).Width = 20
	   .Cells(5).Width = 40
	   .Cells(6).Width = 20
	   .Cells(7).Width = 40
	   .Cells(8).Width = 20 
	   .Cells(9).Width = 43 
	   .Cells(10).Width = 40             
	End With
	data_exibe = data&" &agrave;s "& horario

	Legenda(1, 3).Colspan= 7	
	'Legenda(1, 10).RowSpan = 2					
	Legenda(1, 3).AddText "<Div align=""center"">Documento impresso em: "&data_exibe&"</div>", "size=8; html=true", Font 		
	Legenda(1, 10).AddText "<Div align=""right"">"&pagina&"</div>", "size=8; html=true", Font 				
	Page.Canvas.DrawTable Legenda, "x="&margem&", y="&margem&"" 
			
			'				
	 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width=50; alignment=left; size=8; color=#000000")
	Relatorio = arquivo
	Do While Len(Relatorio) > 0
		CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
	 
		If CharsPrinted = Len(Relatorio) Then Exit Do
		   SET Page = Page.NextPage
		Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
	Loop 								
Next		


Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

