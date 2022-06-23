<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 60 'valor em segundos
'Apurar Despesas no Período
arquivo="SWD011"
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/parametros.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/funcoes7.asp"-->
<% 
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

if mes<10 then
mes="0"&mes
end if

data = dia &"/"& mes &"/"& ano

if min<10 then
min="0"&min
end if

horario = hora & ":"& min


tipo = request.Form("tipo")
modalidade = request.Form("modalidade")
etapa = request.Form("etapa")
turma = request.Form("turma")

dia_de= request.Form("dia_de")
mes_de = request.Form("mes_de")
dia_ate = request.Form("dia_ate")
mes_ate = request.Form("mes_ate")

data_de = mes_de&"/"&dia_de&"/"&ano_letivo
data_ate = mes_ate&"/"&dia_ate&"/"&ano_letivo
dia_de=dia_de*1
if dia_de<10 then
dia_de="0"&dia_de
end if
dia_ate=dia_ate*1
if dia_ate<10 then
dia_ate="0"&dia_ate
end if
mes_de=mes_de*1
if mes_de<10 then
mes_de="0"&mes_de
end if
mes_ate=mes_ate*1
if mes_ate<10 then
mes_ate="0"&mes_ate
end if
data_de_exibe = dia_de&"/"&mes_de&"/"&ano_letivo
data_ate_exibe = dia_ate&"/"&mes_ate&"/"&ano_letivo




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

sub Cabecalho(margem, page_w, page_h, logo_w, logo_h, area_utilizavel, altura_logo_gde, altura_primeiro_separador, tit, altura_separador)
'NOVO CABEÇALHO==========================================================================================		
	Set Param_Logo_Gde = Pdf.CreateParam	
	area_utilizavel=page_w - (margem*2)					
	largura_logo_gde=formatnumber(logo_w*0.3,0)
	altura_logo_gde=formatnumber(logo_h*0.3,0)

	Param_Logo_Gde("x") = margem
	Param_Logo_Gde("y") = page_h - altura_logo_gde -22
	Param_Logo_Gde("ScaleX") = 0.3
	Param_Logo_Gde("ScaleY") = 0.3
	Page.Canvas.DrawImage Logo, Param_Logo_Gde

	x_texto=largura_logo_gde+ margem+10
	y_texto=formatnumber(page_h - altura_logo_gde/2,0)
	width_texto=page_w -largura_logo_gde - 80			
	SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
	Text = "<p><i><b>"&UCASE(no_unidade)&"</b></i></p><br><font style=""font-size:10pt;"">"&endereco_unidade&"</FONT>" 			
	
	Do While Len(Text) > 0
		CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
	 
		If CharsPrinted = Len(Text) Then Exit Do
			SET Page = Page.NextPage
		Text = Right( Text, Len(Text) - CharsPrinted)
	Loop 	
	x_texto= margem			
	y_texto=y_texto-34			
	width_texto=page_w - (margem*2)


	SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
	Text = "<p><center><i><b><font style=""font-size:18pt;"">Apurar Despesas no Per&iacute;odo</font></b></i><br><font style=""font-size:10pt;"">"&tit&" De "&data_de_exibe&" at&eacute; "&data_ate_exibe&"</font></center></p>"
	

	Do While Len(Text) > 0
		CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
	 
		If CharsPrinted = Len(Text) Then Exit Do
			SET Page = Page.NextPage
		Text = Right( Text, Len(Text) - CharsPrinted)
	Loop 
	
	
	Page.Canvas.SetParams "LineWidth=1" 
	Page.Canvas.SetParams "LineCap=0" 
	inicio_primeiro_separador=largura_logo_gde+margem+10
	altura_primeiro_separador= y_texto+18
	With Page.Canvas
	   .MoveTo inicio_primeiro_separador, altura_primeiro_separador
	   .LineTo area_utilizavel+margem, altura_primeiro_separador
	   .Stroke
	End With 		

	altura_primeiro_separador = altura_primeiro_separador-55
	
	
	Page.Canvas.SetParams "LineWidth=2" 
	Page.Canvas.SetParams "LineCap=0" 
	With Page.Canvas
	   .MoveTo margem, altura_primeiro_separador
	   .LineTo Page.Width - margem, altura_primeiro_separador
	   .Stroke
	End With 					
'================================================================================================================

	altura_separador = altura_primeiro_separador-5	
end sub

sub Rodape(margem, area_utilizavel, Paginacao)
		SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
					
		Relatorio = "SWD011 - Sistema Web Diretor"
		Do While Len(Relatorio) > 0
			CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
		 
			If CharsPrinted = Len(Relatorio) Then Exit Do
			   SET Page = Page.NextPage
			Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
		Loop 
		
		Param_Relatorio.Add "alignment=right" 
		
		
		Do While Len(Paginacao) > 0
			CharsPrinted = Page.Canvas.DrawText(Paginacao, Param_Relatorio, Font )
		 
			If CharsPrinted = Len(Paginacao) Then Exit Do
			   SET Page = Page.NextPage
			Paginacao = Right( Paginacao, Len(Paginacao) - CharsPrinted)
		Loop 
		Paginacao = Paginacao+1
		
		Param_Relatorio.Add "html=true" 
		
		data_hora = "<center>Impresso em "&data &" &agrave;s "&horario&"</center>"
		Do While Len(data_hora) > 0
			CharsPrinted = Page.Canvas.DrawText(data_hora, Param_Relatorio, Font )
		 
			If CharsPrinted = Len(data_hora) Then Exit Do
			   SET Page = Page.NextPage
			data_hora = Right( data_hora, Len(data_hora) - CharsPrinted)
		Loop

end sub





















cols = 0
if tipo = "pr" then
'descontinuado
'		if isnull(turma) or turma = "" or turma = "999990" then
'			if isnull(etapa) or etapa = "" or etapa = "999990" then			
'				query = "SELECT TB_Mov_Estoque.NU_Unidade, TB_Mov_Estoque.CO_Curso, TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma, TB_Mov_Estoque.CO_Projeto, Sum(TB_Mov_Estoque.VA_Total_Pedido) AS SomaDeVA_Total_Pedido FROM TB_Mov_Estoque Where (DA_Pedido BETWEEN #"&data_de&"# AND #"&data_ate&"#) GROUP BY TB_Mov_Estoque.NU_Unidade, TB_Mov_Estoque.CO_Curso, TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma, TB_Mov_Estoque.CO_Projeto"	
'				cols = 6
'				cols_tits = "UNIDADE#!#CURSO#!#ETAPA#!#TURMA#!#PROJETO#!#TOTAL"
'				dados_buscados ="NU_Unidade#!#CO_Curso#!#CO_Etapa#!#CO_Turma#!#CO_Projeto#!#SomaDeVA_Total_Pedido" 
'			else	
'				query = "SELECT TB_Mov_Estoque.NU_Unidade, TB_Mov_Estoque.CO_Curso, TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Projeto, Sum(TB_Mov_Estoque.VA_Total_Pedido) AS SomaDeVA_Total_Pedido FROM TB_Mov_Estoque Where (DA_Pedido BETWEEN #"&data_de&"# AND #"&data_ate&"#)  AND TB_Mov_Estoque.NU_Unidade = 1 AND TB_Mov_Estoque.CO_Curso = '0' AND TB_Mov_Estoque.CO_Etapa = '"&etapa&"' GROUP BY TB_Mov_Estoque.NU_Unidade, TB_Mov_Estoque.CO_Curso, TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Projeto"	
'				cols = 5	
'				cols_tits = "UNIDADE#!#CURSO#!#ETAPA#!#PROJETO#!#TOTAL"	
'				dados_buscados ="NU_Unidade#!#CO_Curso#!#CO_Etapa#!#CO_Projeto#!#SomaDeVA_Total_Pedido" 										
'			end if
'		else
'			query = "SELECT TB_Mov_Estoque.CO_Projeto, Sum(TB_Mov_Estoque.VA_Total_Pedido) AS SomaDeVA_Total_Pedido FROM TB_Mov_Estoque Where (DA_Pedido BETWEEN #"&data_de&"# AND #"&data_ate&"#) AND TB_Mov_Estoque.NU_Unidade = 1 AND TB_Mov_Estoque.CO_Curso = '0' AND TB_Mov_Estoque.CO_Etapa = '"&etapa&"' AND TB_Mov_Estoque.CO_Turma = '"&turma&"' GROUP BY TB_Mov_Estoque.CO_Projeto"			
'			cols = 2
'				cols_tits = "PROJETO#!#TOTAL"	
'				dados_buscados ="CO_Projeto#!#SomaDeVA_Total_Pedido" 		
'										
'		end if	
elseif tipo = "it" then
''descontinuado
'		if modalidade = "eg" then
'			if isnull(turma) or turma = "" or turma = "999990" then
'				if isnull(etapa) or etapa = "" or etapa = "999990" then
'					query = "SELECT TB_Mov_Estoque.NU_Unidade, TB_Mov_Estoque.CO_Curso, TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma, TB_Mov_Estoque_Item.CO_Item, Sum(([QT_Solicitado]*[VA_Medio_Refer])) AS Expr1 FROM (TB_Mov_Estoque INNER JOIN TB_Mov_Estoque_Item ON TB_Mov_Estoque.NU_Pedido = TB_Mov_Estoque_Item.NU_Pedido) LEFT JOIN TB_Item ON TB_Mov_Estoque_Item.CO_Item = TB_Item.CO_Item GROUP BY TB_Mov_Estoque.NU_Unidade, TB_Mov_Estoque.CO_Curso, TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma, TB_Mov_Estoque_Item.CO_Item ORDER BY Sum(([QT_Solicitado]*[VA_Medio_Refer])) DESC"					
'					cols = 6					
'				    cols_tits = "UNIDADE#!#CURSO#!#ETAPA#!#TURMA#!#ITEM#!#TOTAL"
'					dados_buscados ="NU_Unidade#!#CO_Curso#!#CO_Etapa#!#CO_Turma#!#CO_Item#!#Expr1" 											
'				else
'					query = "SELECT TB_Mov_Estoque.NU_Unidade, TB_Mov_Estoque.CO_Curso, TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque_Item.CO_Item, Sum(([QT_Solicitado]*[VA_Medio_Refer])) AS Expr1 FROM (TB_Mov_Estoque INNER JOIN TB_Mov_Estoque_Item ON TB_Mov_Estoque.NU_Pedido = TB_Mov_Estoque_Item.NU_Pedido) LEFT JOIN TB_Item ON TB_Mov_Estoque_Item.CO_Item = TB_Item.CO_Item Where (DA_Pedido BETWEEN #"&data_de&"# AND #"&data_ate&"#) AND TB_Mov_Estoque.NU_Unidade = 1 AND TB_Mov_Estoque.CO_Curso = '0' AND TB_Mov_Estoque.CO_Etapa = '"&etapa&"'GROUP BY TB_Mov_Estoque.NU_Unidade, TB_Mov_Estoque.CO_Curso, TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque_Item.CO_Item ORDER BY Sum(([QT_Solicitado]*[VA_Medio_Refer])) DESC"						
'					cols = 5
'					cols_tits = "UNIDADE#!#CURSO#!#ETAPA#!#ITEM#!#TOTAL"	
'					dados_buscados ="NU_Unidade#!#CO_Curso#!#CO_Etapa#!#CO_Item#!#Expr1" 												
'				end if
'			else
'				query = "SELECT TB_Mov_Estoque_Item.CO_Item, Sum(([QT_Solicitado]*[VA_Medio_Refer])) AS Expr1 				FROM (TB_Mov_Estoque INNER JOIN TB_Mov_Estoque_Item ON TB_Mov_Estoque.NU_Pedido = TB_Mov_Estoque_Item.NU_Pedido) LEFT JOIN TB_Item ON TB_Mov_Estoque_Item.CO_Item = TB_Item.CO_Item Where (DA_Pedido BETWEEN #"&data_de&"# AND #"&data_ate&"#) AND TB_Mov_Estoque.NU_Unidade = 1 AND TB_Mov_Estoque.CO_Curso = '0' AND TB_Mov_Estoque.CO_Etapa = '"&etapa&"' AND TB_Mov_Estoque.CO_Turma = '"&turma&"' GROUP BY TB_Mov_Estoque_Item.CO_Item ORDER BY Sum(([QT_Solicitado]*[VA_Medio_Refer])) DESC"	
'				cols = 2	
'				cols_tits = "ITEM#!#TOTAL"	
'				dados_buscados ="CO_Item#!#Expr1" 															
'			end if		
'		elseif modalidade = "pp" then
'				query = "SELECT TB_Mov_Estoque.CO_Projeto, TB_Mov_Estoque_Item.CO_Item, Sum(([QT_Solicitado]*[VA_Medio_Refer])) AS Expr1 FROM (TB_Mov_Estoque INNER JOIN TB_Mov_Estoque_Item ON TB_Mov_Estoque.NU_Pedido = TB_Mov_Estoque_Item.NU_Pedido) LEFT JOIN TB_Item ON TB_Mov_Estoque_Item.CO_Item = TB_Item.CO_Item Where (DA_Pedido BETWEEN #"&data_de&"# AND #"&data_ate&"#)  GROUP BY TB_Mov_Estoque.CO_Projeto, TB_Mov_Estoque_Item.CO_Item ORDER BY Sum(([QT_Solicitado]*[VA_Medio_Refer])) DESC"
'			cols = 3	
'			cols_tits = "PROJETO#!#ITEM#!#TOTAL"
'			dados_buscados ="CO_Projeto#!#CO_Item#!#Expr1" 												
'		end if	
else
	if modalidade="cc" then
		tit = "Centro de Custo"
		cols = 4	
		cols_tits = "ITEM#!#NOME#!#QTDE REQUERIDA#!#VALOR DESPESA"
		dados_buscados ="CO_Item#!#NO_Item#!#QT_Solicitado#!#Expr1" 	
	
		if isnull(turma) or turma = "" or turma = "999990" then
			if isnull(etapa) or etapa = "" or etapa = "999990" then
				query = "SELECT TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma, TB_Mov_Estoque_Item.CO_Item,TB_Mov_Estoque_Item.QT_Solicitado, Sum(([QT_Solicitado]*[VA_Medio_Refer])) AS Expr1 FROM (TB_Mov_Estoque INNER JOIN TB_Mov_Estoque_Item ON TB_Mov_Estoque.NU_Pedido = TB_Mov_Estoque_Item.NU_Pedido) INNER JOIN TB_Item ON TB_Mov_Estoque_Item.CO_Item = TB_Item.CO_Item Where (DA_Pedido BETWEEN #"&data_de&"# AND #"&data_ate&"#) AND TB_Mov_Estoque.NU_Unidade = 1 AND TB_Mov_Estoque.CO_Curso = '0' GROUP BY TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma, TB_Mov_Estoque_Item.CO_Item,TB_Mov_Estoque_Item.QT_Solicitado ORDER BY TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma, Sum(([QT_Solicitado]*[VA_Medio_Refer])) DESC"															
			else
				query = "SELECT TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma, TB_Mov_Estoque_Item.CO_Item,TB_Mov_Estoque_Item.QT_Solicitado, Sum(([QT_Solicitado]*[VA_Medio_Refer])) AS Expr1 FROM (TB_Mov_Estoque INNER JOIN TB_Mov_Estoque_Item ON TB_Mov_Estoque.NU_Pedido = TB_Mov_Estoque_Item.NU_Pedido) INNER JOIN TB_Item ON TB_Mov_Estoque_Item.CO_Item = TB_Item.CO_Item Where (DA_Pedido BETWEEN #"&data_de&"# AND #"&data_ate&"#) AND TB_Mov_Estoque.NU_Unidade = 1 AND TB_Mov_Estoque.CO_Curso = '0' AND TB_Mov_Estoque.CO_Etapa = '"&etapa&"'GROUP BY TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma, TB_Mov_Estoque_Item.CO_Item,TB_Mov_Estoque_Item.QT_Solicitado ORDER BY TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma, Sum(([QT_Solicitado]*[VA_Medio_Refer])) DESC"														
			end if
		else
			query = "SELECT TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma, TB_Mov_Estoque_Item.CO_Item,TB_Mov_Estoque_Item.QT_Solicitado, Sum(([QT_Solicitado]*[VA_Medio_Refer])) AS Expr1 FROM (TB_Mov_Estoque INNER JOIN TB_Mov_Estoque_Item ON TB_Mov_Estoque.NU_Pedido = TB_Mov_Estoque_Item.NU_Pedido) INNER JOIN TB_Item ON TB_Mov_Estoque_Item.CO_Item = TB_Item.CO_Item Where (DA_Pedido BETWEEN #"&data_de&"# AND #"&data_ate&"#) AND TB_Mov_Estoque.NU_Unidade = 1 AND TB_Mov_Estoque.CO_Curso = '0' AND TB_Mov_Estoque.CO_Etapa = '"&etapa&"' AND TB_Mov_Estoque.CO_Turma = '"&turma&"' GROUP BY TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma, TB_Mov_Estoque_Item.CO_Item,TB_Mov_Estoque_Item.QT_Solicitado ORDER BY TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma, Sum(([QT_Solicitado]*[VA_Medio_Refer])) DESC"								
		end if	
	elseif modalidade="it" then
		tit = "Itens"
		cols = 6	
		cols_tits = "ITEM#!#NOME#!#QTDE REQUERIDA#!#PEDIDO#!#ETAPA#!#TURMA"
		dados_buscados ="CO_Item#!#NO_Item#!#QT_Solicitado#!#NU_Pedido#!#CO_Etapa#!#CO_Turma" 
		query = "SELECT TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma, TB_Mov_Estoque_Item.CO_Item,TB_Mov_Estoque_Item.QT_Solicitado, TB_Mov_Estoque.NU_Pedido FROM (TB_Mov_Estoque INNER JOIN TB_Mov_Estoque_Item ON TB_Mov_Estoque.NU_Pedido = TB_Mov_Estoque_Item.NU_Pedido) LEFT JOIN TB_Item ON TB_Mov_Estoque_Item.CO_Item = TB_Item.CO_Item ORDER BY TB_Item.CO_Item,  TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma DESC"															
	end if		
end if



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

		Set CON9 = Server.CreateObject("ADODB.Connection") 
		ABRIR9 = "DBQ="& CAMINHO_ax & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON9.Open ABRIR9
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0	

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
		
		if municipio_unidade="" or isnull(municipio_unidade) or uf_unidade_municipio="" or isnull(uf_unidade_municipio)then
		else		
		
			if bairro_unidade="" or isnull(bairro_unidade)then
			else
			
				Set RS3b = Server.CreateObject("ADODB.Recordset")
				SQL3b = "SELECT * FROM TB_Bairros WHERE SG_UF='"& uf_unidade_municipio &"' AND CO_Municipio="&municipio_unidade&" AND CO_Bairro = "&bairro_unidade
				RS3b.Open SQL3b, CON0
				
				bairro_unidade=RS3b("NO_Bairro")				
				bairro_unidade=" - "&bairro_unidade
			end if		
			Set RS3m = Server.CreateObject("ADODB.Recordset")
			SQL3m = "SELECT * FROM TB_Municipios WHERE SG_UF='"& uf_unidade_municipio &"' AND CO_Municipio="&municipio_unidade
			RS3m.Open SQL3m, CON0
			
			municipio_unidade=RS3m("NO_Municipio")						
		end if
		endereco_unidade=rua_unidade&numero_unidade&complemento_unidade&bairro_unidade&cep_unidade&"<br>"&municipio_unidade&uf_unidade	
				

			SET Page = Doc.Pages.Add( 595, 842 )
			margem=30	
			page_w = Page.Width
			page_h = Page.Height
			logo_w = Logo.Width
			logo_h = Logo.Height
'Cabeçalho====================================================================================================			
			call Cabecalho(margem, page_w, page_h, logo_w, logo_h, area_utilizavel, altura_logo_gde, altura_primeiro_separador, tit, altura_separador)
		

			y_primeira_tabela=Page.Height - altura_logo_gde
			x_primeira_tabela = margem

			
Set RSI = Server.CreateObject("ADODB.Recordset")
SQLI = query
'			response.Write(SQLI)
'			response.End()	
RSI.Open SQLI, CON9				
			
	
if tipo = "pr" or tipo = "it" then	
'descontinuado
'			Set param_table1 = Pdf.CreateParam("width=534; height=20; rows=1; cols="&cols&"; border=0; cellborder=0.1; cellspacing=0;")
'			Set Table = Doc.CreateTable(param_table1)
'			Table.Font = Font
'			
'			y_primeira_tabela=altura_primeiro_separador-10
'			x_primeira_tabela=margem
'			cel_width = formatnumber(534/cols,0)
'			With Table.Rows(1)
'			  ' .Cells(1).Height = 20
'			  cols=cols*1
'			  if cols=6 then
'				 .Cells(1).Width = cel_width	
'				 .Cells(2).Width = cel_width-20	
'				 .Cells(3).Width = cel_width	
'				 .Cells(4).Width = cel_width-50	
'				 .Cells(5).Width = cel_width+70	
'				 .Cells(6).Width = cel_width					 				 				 				 				 			  
'			  else
'				  for i = 1 to cols
'				   .Cells(i).Width = cel_width	
'				  next 		
'			  end if	     		   		   		   		   
'			End With
'
'			titulos = split(cols_tits, "#!#")
'			for i = 1 to cols
'				c=i-1
'				Table(1, i).AddText "<center><b>"&titulos(c)&"</b></center>", "size=8;indenty=2;html=true", Font 				
'			Next													
'			linha = 1		
'			IF RSI.EOF THEN
'			  For i = 1 to 20
'				Set Row = Table.Rows.Add(15)				  
'			  next
'			else
'				dados_cols = split(dados_buscados,"#!#")
'				while not RSI.EOF 
'					coluna = 0	
'				    linha = linha+1
'					Set Row = Table.Rows.Add(15)									
'					For d = 0 to ubound(dados_cols)
'						response.Write(dados_cols(d))
'						dado = RSI(dados_cols(d))
'						
'						if dados_cols(d) = "TB_Mov_Estoque.NU_Unidade" or dados_cols(d) = "NU_Unidade" then	
'						    unidade = dado
'							nome_item = GeraNomes("U",unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)							
'						elseif dados_cols(d) = "TB_Mov_Estoque.CO_Curso" or dados_cols(d) = "CO_Curso" then
'						    curso = dado	
'							nome_item = GeraNomes("C",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro)							
'						elseif dados_cols(d) = "TB_Mov_Estoque.CO_Etapa" or dados_cols(d) = "CO_Etapa" then	
'						    co_etapa = dado
'							nome_item = GeraNomes("E",curso,co_etapa,variavel3,variavel4,variavel5,CON0,outro)							
'						elseif dados_cols(d) = "TB_Mov_Estoque.CO_Turma" or dados_cols(d) = "CO_Turma" then	
'							nome_item = dado																														
'						elseif dados_cols(d) = "TB_Mov_Estoque.CO_Projeto" or dados_cols(d) = "CO_Projeto" then
'							Set RSC = Server.CreateObject("ADODB.Recordset")
'							SQLC = "Select NO_Projeto From TB_Projeto where CO_Projeto = "&dado
'							RSC.Open SQLC, CON9	
'							
'							nome_item = RSC("NO_Projeto")								
'							nome_item = replace_latin_char(nome_item,"html")						
'						elseif dados_cols(d) = "TB_Mov_Estoque_Item.CO_Item" or dados_cols(d) = "CO_Item"  then
'							Set RSC = Server.CreateObject("ADODB.Recordset")
'							SQLC = "Select NO_Item From TB_Item where CO_Item = "&dado
'							RSC.Open SQLC, CON9	
'							
'							nome_item = RSC("NO_Item")						
'							nome_item = replace_latin_char(nome_item,"html")
'						else
'							if isnumeric(dado) then
'								nome_item = formatcurrency(dado)
'							else
'								nome_item = replace_latin_char(nome_item,"html")							
'							end if								
'						end if
'						coluna = coluna+1	
'						
'						Table(linha, coluna).AddText "<center>"&nome_item&"</center>", "size=7;indentx=2;indenty=2;html=true", Font 		
'				   next				   
'				RSI.MOVENEXT
'				WEND 	
'			 END IF	 	
else
	titulos = split(cols_tits, "#!#")
	etapa_parametro="X1X2X3"
	turma_parametro="X1X2X3"	
	item_parametro="X1X2X3"	
	nome_item_parametro = "X1X2X3"		
	altura_tabela =	20
	conta_tabelas = 0
	MaxHeight = 700											
	linha = 2	
	Paginacao = 1	
	y_primeira_tabela=altura_primeiro_separador
	x_primeira_tabela=margem	
	
	Set param_table1 = Pdf.CreateParam("width=534; height=40; rows=2; cols="&cols&"; border=0; cellborder=0.1; cellspacing=0; x="&x_primeira_tabela&"; y="&y_primeira_tabela&"; MaxHeight="&MaxHeight)
	Set Table = Doc.CreateTable(param_table1)
	
	Table.At(1, 1).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False" 
	Table(1, 1).ColSpan = cols		
	cel_width = formatnumber(534/cols,0)
	With Table.Rows(1)
		if cols=6 then
		 .Cells(1).Width = cel_width-40		
		 .Cells(2).Width = cel_width+40	
		 .Cells(3).Width = cel_width-25		
		 .Cells(4).Width = cel_width-25	
		 .Cells(5).Width = cel_width+70	
		 .Cells(6).Width = cel_width-20					 				 				 				 				 			  
		else
		 .Cells(1).Width = cel_width-20		
		 .Cells(2).Width = cel_width+40	
		 .Cells(3).Width = cel_width-10		
		 .Cells(4).Width = cel_width-10	
		end if 	     		   		   		   		   
	End With		
	for i = 1 to cols
		c=i-1
		Table(2, i).AddText "<center><b>"&titulos(c)&"</b></center>", "size=8;indenty=2;html=true", Font 				
	Next			
	Table.Font = Font	
					
	IF RSI.EOF THEN		
		For i = 1 to 20
			Set Row = Table.Rows.Add(15)				  
		next
	else	
		while not RSI.EOF
			curso= 0
			co_etapa = RSI("CO_Etapa")
			co_turma = RSI("CO_Turma")
			codigo_item = RSI("CO_Item")					
			
			Set Row = Table.Rows.Add(15)			
			linha = linha+1		
			if (modalidade = "cc" and (etapa_parametro<>co_etapa or turma_parametro<> co_turma))  or (modalidade = "it" and item_parametro <>codigo_item) then
				etapa_parametro=co_etapa
				turma_parametro=co_turma
				item_parametro = codigo_item
				nome_etapa = GeraNomes("E",curso,co_etapa,variavel3,variavel4,variavel5,CON0,outro)	
				etapa_turma = nome_etapa&" - "&co_turma
				
				if modalidade = "cc" and conta_tabelas = 0 then	

					SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&altura_separador&"; height=50; width="&area_utilizavel&"; alignment=left; html=true; size=8; color=#000000")					
					
					Do While Len(etapa_turma) > 0
						CharsPrinted = Page.Canvas.DrawText(etapa_turma, Param_Relatorio, Font )
					 
						If CharsPrinted = Len(etapa_turma) Then Exit Do
						   SET Page = Page.NextPage
						etapa_turma = Right( etapa_turma, Len(etapa_turma) - CharsPrinted)
					Loop 										
				elseif conta_tabelas>0 then			
					if modalidade = "cc" then
						Table(linha, ubound(dados_cols)).AddText "<center>Total</center>", "size=7;indentx=2;indenty=2;html=true", Font 										
						Table(linha, ubound(dados_cols)+1).AddText "<center>"&formatcurrency(totalizador)&"</center>", "size=7;indentx=2;indenty=2;html=true", Font 						
					else
						Table(linha, 2).AddText "<center>Total</center>", "size=7;indentx=2;indenty=2;html=true", Font 										
						Table(linha, 3).AddText "<center>"&totalizador&"</center>", "size=7;indentx=2;indenty=2;html=true", Font 		
					end if
					Set Row = Table.Rows.Add(20)								
					linha=linha+1	
					
					if modalidade = "cc" then											
						Table(linha, 1).AddText  etapa_turma, "size=8;indentx=2;indenty=5;html=true", Font
					end if
											
					Table.At(linha, 1).SetBorderParams "Left=False, Right=False" 
					Table(linha, 1).ColSpan = ubound(dados_cols)+1	
					Set Row = Table.Rows.Add(20)								
					linha=linha+1	
					for i = 1 to cols
						c=i-1
						Table(linha, i).AddText "<center><b>"&titulos(c)&"</b></center>", "size=8;indenty=2;html=true", Font 				
					Next											
											
					Set Row = Table.Rows.Add(15)
					linha = linha+1					
				end if					
				totalizador = 0						
				conta_tabelas = conta_tabelas+1																		
			end if
							
			dados_cols = split(dados_buscados,"#!#")
			coluna = 0	
							
			For d = 0 to ubound(dados_cols)
				'response.Write(dados_cols(d))

				if dados_cols(d) = "TB_Mov_Estoque_Item.NO_Item" or dados_cols(d) = "NO_Item"  then
					dado = RSI("CO_Item")						
					Set RSC = Server.CreateObject("ADODB.Recordset")
					SQLC = "Select NO_Item From TB_Item where CO_Item = "&dado
					RSC.Open SQLC, CON9	
					
					
					nome_item = RSC("NO_Item")
					if nome_item_parametro <> nome_item then
						nome_item_parametro = 	nome_item															
						nome_item = replace_latin_char(nome_item,"html")
					else
						nome_item = ""
					end if	
				elseif dados_cols(d) = "TB_Mov_Estoque.CO_Etapa" or dados_cols(d) = "CO_Etapa" then	
					nome_item = GeraNomes("E",curso,RSI("CO_Etapa"),variavel3,variavel4,variavel5,CON0,outro)							
				elseif dados_cols(d) = "TB_Mov_Estoque.CO_Turma" or dados_cols(d) = "CO_Turma" then	
					nome_item = RSI("CO_Turma")
				else
					dado = RSI(dados_cols(d))					
					if isnumeric(dado) then
						if dados_cols(d) = "Expr1" then
							nome_item = formatcurrency(dado)
							totalizador = totalizador+nome_item
						elseif dados_cols(d) = "QT_Solicitado" and modalidade ="it" then
							nome_item = dado
							totalizador = totalizador+nome_item								
						elseif dados_cols(d) = "CO_Item" or dados_cols(d) = "NU_Pedido" then
							nome_item = dado
						else
							nome_item = formatnumber(dado)								
						end if
					else
						nome_item = dado							
					end if								
				end if
				if not isnumeric(dado) then
					nome_item = replace_latin_char(nome_item,"html")	
				end if
				coluna = coluna+1	
				
			Table(linha, coluna).AddText "<center>"&nome_item&"</center>", "size=7;indentx=2;indenty=2;html=true", Font 		
		   next				   
		RSI.MOVENEXT
		WEND 					
	 END IF	 		 
			 
end if	
'response.End()
Set Row = Table.Rows.Add(15)
linha=linha+1
if modalidade = "cc" then
	Table(linha, ubound(dados_cols)).AddText "<center>Total</center>", "size=7;indentx=2;indenty=2;html=true", Font 										
	Table(linha, ubound(dados_cols)+1).AddText "<center>"&formatcurrency(totalizador)&"</center>", "size=7;indentx=2;indenty=2;html=true", Font 
else
	Table(linha, 2).AddText "<center>Total</center>", "size=7;indentx=2;indenty=2;html=true", Font 										
	Table(linha, 3).AddText "<center>"&totalizador&"</center>", "size=7;indentx=2;indenty=2;html=true", Font 		
end if				 	


Do While True
	limite=limite+1
	   LastRow = Page.Canvas.DrawTable( Table, param_table1 )

		if LastRow >= Table.Rows.Count Then 
			Exit Do ' entire table displayed
		else
			call Rodape(margem, area_utilizavel, Paginacao)						
			
			' Display remaining part of table on the next page
			Set Page = Page.NextPage	
			param_table1.Add( "RowTo=1; RowFrom=1" ) ' Row 1 is header.
			param_table1("RowFrom1") = LastRow + 1 ' RowTo1 is omitted and presumed infinite
					
'Cabeçalho====================================================================================================			
			call Cabecalho(margem, page_w, page_h, logo_w, logo_h,area_utilizavel, altura_logo_gde, altura_primeiro_separador, tit, altura_separador)				 
		end if
		if limite>10000 then
		response.Write("ERRO!")
		response.end()
		end if 
'		if modalidade = "cc" then	
'			SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&altura_separador&"; height=50; width="&area_utilizavel&"; alignment=left; html=true; size=8; color=#000000")					
'										response.Write("3;"&altura_separador&"<BR>")
'			Do While Len(etapa_turma) > 0
'				CharsPrinted = Page.Canvas.DrawText(etapa_turma, Param_Relatorio, Font )
'			 
'				If CharsPrinted = Len(etapa_turma) Then Exit Do
'				   SET Page = Page.NextPage
'				etapa_turma = Right( etapa_turma, Len(etapa_turma) - CharsPrinted)
'			Loop 					
'		end if					
	Loop							

call Rodape(margem, area_utilizavel, Paginacao)				


Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

