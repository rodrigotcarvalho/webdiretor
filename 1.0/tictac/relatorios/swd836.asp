<%'On Error Resume Next%>
<%
'Server.ScriptTimeout = 60
Server.ScriptTimeout = 600 'valor em segundos
'Lista de fornecedores
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/funcoes2.asp"-->

<% 
arquivo="SWD836"
response.Charset="ISO-8859-1"
'opt= request.QueryString("opt")
'ori= request.QueryString("ori")

'unidade_form = request.Form("unidade")
'curso_form  = request.Form("curso")
'etapa_form  = request.Form("etapa")
'turma_form  = request.Form("turma")
ma = request.form("ma")
ma = ma*1
select case ma
 case 0 
 mes_a = "Todos"
 case 1 
 mes_a = "Janeiro"
 case 2 
 mes_a = "Fevereiro"
 case 3 
 mes_a = "Mar&ccedil;o"
 case 4
 mes_a = "Abril"
 case 5
 mes_a = "Maio"
 case 6 
 mes_a = "Junho"
 case 7
 mes_a = "Julho"
 case 8 
 mes_a = "Agosto"
 case 9 
 mes_a = "Setembro"
 case 10 
 mes_a = "Outubro"
 case 11 
 mes_a = "Novembro"
 case 12 
 mes_a = "Dezembro"
end select					  


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

data_exibe = dia &"/"& meswrt &"/"& ano
horario = hora & ":"& minwrt	

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

		Set CON9 = Server.CreateObject("ADODB.Connection") 
		ABRIR9 = "DBQ="& CAMINHO_ax & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON9.Open ABRIR9


	Set RS2 = Server.CreateObject("ADODB.Recordset")
	SQL2 = "SELECT * FROM TB_Unidade WHERE NU_Unidade=1"
	RS2.Open SQL2, CON0
					
	no_unidade = RS2("TX_Imp_Cabecalho")		
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
	
	if municipio_unidade="" or isnull(municipio_unidade) or uf_unidade_municipio="" or isnull(uf_unidade_municipio) then
	else
		Set RS3m = Server.CreateObject("ADODB.Recordset")
		SQL3m = "SELECT * FROM TB_Municipios WHERE SG_UF='"& uf_unidade_municipio &"' AND CO_Municipio="&municipio_unidade
		RS3m.Open SQL3m, CON0
		cod_municipio = municipio_unidade
		municipio_unidade=RS3m("NO_Municipio")						
	end if
	
	if municipio_unidade="" or isnull(municipio_unidade) or uf_unidade_municipio="" or isnull(uf_unidade_municipio) or bairro_unidade="" or isnull(bairro_unidade)then
	else
		Set RSb = Server.CreateObject("ADODB.Recordset")
		SQLb = "SELECT * FROM TB_Bairros WHERE SG_UF='"& uf_unidade_municipio&"' AND CO_Municipio ="&cod_municipio&" AND CO_Bairro="&bairro_unidade		
		RSb.Open SQLb, CON0			
		
		bairro_unidade=" - "&RSb("NO_Bairro")			
	end if			
	endereco_unidade=rua_unidade&numero_unidade&complemento_unidade&bairro_unidade&cep_unidade&"<br>"&municipio_unidade&uf_unidade					
						

'
'			no_curso= no_etapa&"&nbsp;"&co_concordancia_curso&"&nbsp;"&no_curso
'			texto_turma = "&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;<b>Turma</b> "&turma
'
'			mensagem_cabecalho=ano_letivo&"&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;"&no_curso&texto_turma
'	
	SET Page = Doc.Pages.Add(595,842)
	
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT  * FROM TB_Fornecedor order by NO_Fornecedor"
	RS.Open SQL, CON9
	
num_tabela = 1
	limite=1
	Paginacao = 1
	
	while not RS.EOF			
'CABEÇALHO==========================================================================================		
	Set Param_Logo_Gde = Pdf.CreateParam
	margem=25			
	area_utilizavel=Page.Width - (margem*2)
	
	largura_logo_gde=formatnumber(Logo.Width*0.4,0)
	altura_logo_gde=formatnumber(Logo.Height*0.4,0)

   Param_Logo_Gde("x") = margem
   Param_Logo_Gde("y") = Page.Height - altura_logo_gde - margem
   Param_Logo_Gde("ScaleX") = 0.4
   Param_Logo_Gde("ScaleY") = 0.4
   Page.Canvas.DrawImage Logo, Param_Logo_Gde

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

	y_texto=y_texto-altura_logo_gde+30
	SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height=30; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
	Text = "<center><i><b><font style=""font-size:18pt;"">Rela&ccedil;&atilde;o de Fornecedores</font></b></i></center>"
	
	
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

	Set param_table1 = Pdf.CreateParam("width=547; height=20; rows=1; cols=4; border=0; cellborder=0; cellspacing=0;")
	Set Table = Doc.CreateTable(param_table1)
	Table.Font = Font
	y_primeira_tabela=altura_segundo_separador-10
	x_primeira_tabela=margem+5
	With Table.Rows(1)
	   .Cells(1).Width = 50
	   .Cells(2).Width = 150  
	   .Cells(3).Width = 50 		   		   		   
	End With
	
	
	Table(1, 1).AddText "<b>Unidade:</b>", "size=9;html=true", Font 
	Table(1, 2).AddText no_unidade, "size=9;html=true", Font 
	'Table(1, 3).AddText "<b>M&ecirc;s:</b>", "size=9;html=true", Font 
	'Table(1, 4).AddText "<div align=LEFT>"&mes_a&"</div>", "size=9;html=true", Font 
	Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 	
	
	Page.Canvas.SetParams "LineWidth=2" 
	Page.Canvas.SetParams "LineCap=0" 
	altura_terceiro_separador= y_primeira_tabela-20
	With Page.Canvas
	   .MoveTo margem, altura_terceiro_separador
	   .LineTo area_utilizavel+margem, altura_terceiro_separador
	   .Stroke
	End With 			

'================================================================================================================	

		cod_cons = RS("CO_Fornecedor")
		nome_projeto  = RS("NO_Fornecedor")
			apelido = RS("NO_Apelido_Fornecedor")
	rua = RS("NO_Logradouro")
	numero = RS("NU_Logradouro")
	complemento = RS("TX_Complemento_Logradouro")
	co_bairro= RS("CO_Bairro")
	co_municipio= RS("CO_Municipio")
	uf= RS("SG_UF")
	cep = RS("CO_CEP")
	telefone = RS("NUS_Telefones")
	cnpj = RS("CO_CNPJ")
	email = RS("TX_EMail")
	contatos = RS("NO_Contatos")	
	ativo = RS("IN_Ativo")
	
	
if isnull(uf) then 
uf = "RJ"
end if

if isnull(co_municipio) then 
	co_municipio = 6001
end if

Set RS2m = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Municipios WHERE CO_Municipio="&co_municipio&" AND SG_UF='"&uf&"'"
		RS2m.Open SQL2m, CON0
		
if not RS2m.EOF	then					

cidade = RS2m("NO_Municipio")
end if

		if not isnull(co_bairro) then
			Set RS2b = Server.CreateObject("ADODB.Recordset")
			SQL2b = "SELECT * FROM TB_Bairros WHERE CO_Bairro = "&co_bairro&" and CO_Municipio="&co_municipio&" AND SG_UF='"&uf&"'"
			RS2b.Open SQL2b, CON0
			
			if not RS2b.eof then
				bairro = RS2b("NO_Bairro")
			end if
	
		end if
		
		ind_ativo = RS("IN_Ativo")	
		if ind_ativo = TRUE then
			nom_ativo = "Ativo"
		else
			nom_ativo = "Inativo"			
		end if		

	colunas_de_notas=8
	total_de_colunas=8				
	altura_medias=180
	'resto_tabelas = num_tabela mod 3
	'if (resto_tabelas=0 and num_tabela >3) or num_tabela = 1 then
	if num_tabela = 1 then
		'y_segunda_tabela=altura_terceiro_separador-10
		y_segunda_tabela= 700
	else
		y_segunda_tabela = y_segunda_tabela-200
	end if	
response.write(num_tabela&" "&y_segunda_tabela&"<br>")
	
	Set param_table2 = Pdf.CreateParam("width=537; height="&altura_medias&"; rows=9; cols=8; border=0; cellborder=0.5; cellspacing=0; x="&margem&"; y="&y_segunda_tabela&"; MaxHeight=670")

	Set Notas_Tit = Doc.CreateTable(param_table2)
	Notas_Tit.Font = Font				
	largura_colunas=(537-100)/colunas_de_notas		
	
	Set param_materias = PDF.CreateParam	
	param_materias.Set "size=8;expand=false" 	
	param_materias.Add "indenty=3;alignment=right;html=true"
	param_materias.Add "indentx=2"	
	'param_materias.Add "expand=true" 
	'param_materias.Add "expand=false" 								
	
	With Notas_Tit.Rows(1)
	   .Cells(1).Width = 100
	   .Cells(2).Width = largura_colunas	
	   .Cells(3).Width = largura_colunas
	End With
	Notas_Tit(1, 3).ColSpan = 2	
	Notas_Tit(1, 5).ColSpan = 4	
	Notas_Tit(1, 1).AddText "<div align=""right""><b>C&oacute;digo:&nbsp;</b></div>", "size=8;indenty=3; html=true", Font 
	Notas_Tit(1, 2).AddText "<div align=""left"">"&cod_cons&"</div>", param_materias		
	Notas_Tit(1, 3).AddText "<div align=""right""><b>Nome do Fornecedor:&nbsp;</b></div>", "size=8;alignment=center; indenty=3;html=true", Font 
	Notas_Tit(1, 5).AddText "<div align=""left"">"&nome_projeto&"</div>", param_materias	
	
	Notas_Tit(2, 2).ColSpan = 3	
	Notas_Tit(2, 6).ColSpan = 3	
	Notas_Tit(2, 1).AddText "<div align=""right""><b>Apelido:&nbsp;</b></div>", "size=8;indenty=3; html=true", Font 
	Notas_Tit(2, 2).AddText "<div align=""left"">"&apelido&"</div>", param_materias		
	Notas_Tit(2, 5).AddText "<div align=""right""><b>CNPJ:&nbsp;</b></div>", "size=8;alignment=center; indenty=3;html=true", Font 
	Notas_Tit(2, 6).AddText "<div align=""left"">"&cnpj&"</div>", param_materias
	
	
	Notas_Tit(3, 2).ColSpan = 5	
	'Notas_Tit(3, 5).ColSpan = 4	
	Notas_Tit(3, 1).AddText "<div align=""right""><b>Logradouro:&nbsp;</b></div>", "size=8;indenty=3; html=true", Font 
	Notas_Tit(3, 2).AddText "<div align=""left"">"&rua&"</div>", param_materias		
	Notas_Tit(3, 7).AddText "<div align=""right""><b>N&uacute;mero:&nbsp;</b></div>", "size=8;alignment=center; indenty=3;html=true", Font 
	Notas_Tit(3, 8).AddText "<div align=""left"">"&numero&"</div>", param_materias	
	
	Notas_Tit(4, 2).ColSpan = 3	
	'Notas_Tit(4, 5).ColSpan = 4	
	Notas_Tit(4, 1).AddText "<div align=""right""><b>Complemento:&nbsp;</b></div>", "size=8;indenty=3; html=true", Font 
	Notas_Tit(4, 2).AddText "<div align=""left"">"&complemento&"</div>", param_materias	
	Notas_Tit(4, 5).AddText "<div align=""right""><b>CEP:&nbsp;</b></div>", "size=8;alignment=center; indenty=3;html=true", Font 
	Notas_Tit(4, 6).AddText "<div align=""left"">"&cep&"</div>", param_materias			
	Notas_Tit(4, 7).AddText "<div align=""right""><b>Estado:&nbsp;</b></div>", "size=8;alignment=center; indenty=3;html=true", Font 
	Notas_Tit(4, 8).AddText "<div align=""left"">"&UF&"</div>", param_materias	
	
	Notas_Tit(5, 2).ColSpan = 3	
	Notas_Tit(5, 6).ColSpan = 3	
	Notas_Tit(5, 1).AddText "<div align=""right""><b>Cidade:&nbsp;</b></div>", "size=8;indenty=3; html=true", Font 
	Notas_Tit(5, 2).AddText "<div align=""left"">"&cidade&"</div>", param_materias	
	Notas_Tit(5, 5).AddText "<div align=""right""><b>Bairro:&nbsp;</b></div>", "size=8;alignment=center; indenty=3;html=true", Font 
	Notas_Tit(5, 6).AddText "<div align=""left"">"&bairro&"</div>", param_materias						
	
	Notas_Tit(6, 2).ColSpan = 7	
	'Notas_Tit(6, 6).ColSpan = 3	
	Notas_Tit(6, 1).AddText "<div align=""right""><b>Telefones de Contato:&nbsp;</b></div>", "size=8;indenty=3; html=true", Font 
	Notas_Tit(6, 2).AddText "<div align=""left"">"&telefone&"</div>", param_materias	
		
	Notas_Tit(7, 2).ColSpan = 7
	'Notas_Tit(6, 6).ColSpan = 3	
	Notas_Tit(7, 1).AddText "<div align=""right""><b>Nomes de Contato:&nbsp;</b></div>", "size=8;indenty=3; html=true", Font 
	Notas_Tit(7, 2).AddText "<div align=""left"">"&contatos&"</div>", param_materias

	Notas_Tit(8, 2).ColSpan = 7	
	Notas_Tit(8, 1).AddText "<div align=""right""><b>E-mail:&nbsp;</b></div>", "size=8;alignment=center; indenty=3;html=true", Font 
	Notas_Tit(8, 2).AddText "<div align=""left"">"&email&"</div>", param_materias	

	Notas_Tit(9, 2).ColSpan = 7	
	Notas_Tit(9, 1).AddText "<div align=""right""><b>Status:&nbsp;</b></div>", "size=8;alignment=center; indenty=3;html=true", Font 
	Notas_Tit(9, 2).AddText "<div align=""left"">"&nom_ativo&"</div>", param_materias	
	
	
				
	Do While True
		limite=limite+1
		'Paginacao = Paginacao+1
	   LastRow = Page.Canvas.DrawTable( Notas_Tit, param_table2 )

		if LastRow >= Notas_Tit.Rows.Count Then 
			Exit Do ' entire table displayed
		else
		
			 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
			
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
			
			data_hora = "<center>Impresso em "&data_exibe &" &agrave;s "&horario&"</center>"
			Do While Len(Relatorio) > 0
				CharsPrinted = Page.Canvas.DrawText(data_hora, Param_Relatorio, Font )
			 
				If CharsPrinted = Len(data_hora) Then Exit Do
				   SET Page = Page.NextPage
				data_hora = Right( data_hora, Len(data_hora) - CharsPrinted)
			Loop 				   ' Display remaining part of table on the next page
			Set Page = Page.NextPage	
			
			param_table2.Add( "RowTo=1; RowFrom=1" ) ' Row 1 is header.
			param_table2("RowFrom1") = LastRow + 1 ' RowTo1 is omitted and presumed infinite
'NOVO CABEÇALHO==========================================================================================		
	Set Param_Logo_Gde = Pdf.CreateParam
	margem=25			
	area_utilizavel=Page.Width - (margem*2)
	
	largura_logo_gde=formatnumber(Logo.Width*0.4,0)
	altura_logo_gde=formatnumber(Logo.Height*0.4,0)

   Param_Logo_Gde("x") = margem
   Param_Logo_Gde("y") = Page.Height - altura_logo_gde - margem
   Param_Logo_Gde("ScaleX") = 0.4
   Param_Logo_Gde("ScaleY") = 0.4
   Page.Canvas.DrawImage Logo, Param_Logo_Gde

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

	y_texto=y_texto-altura_logo_gde+30
	SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height=30; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
	Text = "<center><i><b><font style=""font-size:18pt;"">Rela&ccedil;&atilde;o de Fornecedores</font></b></i></center>"
	
	
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

	Set param_table1 = Pdf.CreateParam("width=547; height=20; rows=1; cols=4; border=0; cellborder=0; cellspacing=0;")
	Set Table = Doc.CreateTable(param_table1)
	Table.Font = Font
	y_primeira_tabela=altura_segundo_separador-10
	x_primeira_tabela=margem+5
	With Table.Rows(1)
	   .Cells(1).Width = 50
	   .Cells(2).Width = 150  
	   .Cells(3).Width = 50 		   		   		   
	End With
	
	
	Table(1, 1).AddText "<b>Unidade:</b>", "size=9;html=true", Font 
	Table(1, 2).AddText no_unidade, "size=9;html=true", Font 
	'Table(1, 3).AddText "<b>M&ecirc;s:</b>", "size=9;html=true", Font 
	'Table(1, 4).AddText "<div align=LEFT>"&mes_a&"</div>", "size=9;html=true", Font 
	Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 	
	
	Page.Canvas.SetParams "LineWidth=2" 
	Page.Canvas.SetParams "LineCap=0" 
	altura_terceiro_separador= y_primeira_tabela-20
	With Page.Canvas
	   .MoveTo margem, altura_terceiro_separador
	   .LineTo area_utilizavel+margem, altura_terceiro_separador
	   .Stroke
	End With 			


	'================================================================================================================			
			 	end if
'				if limite>300 then
'					response.Write("ERRO!")
'					response.end()
'				end if 
			Loop
			
			 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
			
			Relatorio = arquivo&" -  Sistema Web Diretor"
			Do While Len(Relatorio) > 0
				CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
			 
				If CharsPrinted = Len(Relatorio) Then Exit Do
				   SET Page = Page.NextPage
				Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
			Loop 				

			SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
			Param_Relatorio.Add "alignment=right" 
			
		
'			
			
			Do While Len(Paginacao) > 0
				CharsPrinted = Page.Canvas.DrawText(Paginacao, Param_Relatorio, Font )			
				If CharsPrinted = Len(Paginacao) Then Exit Do
				SET Page = Page.NextPage
				Paginacao = Right( Paginacao, Len(Relatorio) - CharsPrinted)
			Loop 
						
			Param_Relatorio.Add "html=true" 
			
			data_hora = "<center>Impresso em "&data_exibe &" &agrave;s "&horario&"</center>"
			Do While Len(Relatorio) > 0
				CharsPrinted = Page.Canvas.DrawText(data_hora, Param_Relatorio, Font )			
				If CharsPrinted = Len(data_hora) Then Exit Do
				SET Page = Page.NextPage
				data_hora = Right( data_hora, Len(data_hora) - CharsPrinted)
			Loop
			

			'if resto_tabelas=0 then	
  if num_tabela =3 then 
	num_tabela = 1
	Set Page = Page.NextPage
	Paginacao = Paginacao+1
  ELSE
  num_tabela = num_tabela+1	  
  end if
																									
  RS.MOVENEXT
  WEND									

	

Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

