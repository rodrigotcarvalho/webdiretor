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

if dia_de<10 then
dia_de="0"&dia_de
end if
if dia_ate<10 then
dia_ate="0"&dia_ate
end if

if mes_de<10 then
mes_de="0"&mes_de
end if
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
cols = 0
if tipo = "pr" then
		if isnull(turma) or turma = "" or turma = "999990" then
			if isnull(etapa) or etapa = "" or etapa = "999990" then			
				query = "SELECT TB_Mov_Estoque.NU_Unidade, TB_Mov_Estoque.CO_Curso, TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma, TB_Mov_Estoque.CO_Projeto, Sum(TB_Mov_Estoque.VA_Total_Pedido) AS SomaDeVA_Total_Pedido FROM TB_Mov_Estoque Where (DA_Pedido BETWEEN #"&data_de&"# AND #"&data_ate&"#) GROUP BY TB_Mov_Estoque.NU_Unidade, TB_Mov_Estoque.CO_Curso, TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma, TB_Mov_Estoque.CO_Projeto"	
				cols = 6
				cols_tits = "UNIDADE#!#CURSO#!#ETAPA#!#TURMA#!#PROJETO#!#TOTAL"
				dados_buscados ="NU_Unidade#!#CO_Curso#!#CO_Etapa#!#CO_Turma#!#CO_Projeto#!#SomaDeVA_Total_Pedido" 
			else	
				query = "SELECT TB_Mov_Estoque.NU_Unidade, TB_Mov_Estoque.CO_Curso, TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Projeto, Sum(TB_Mov_Estoque.VA_Total_Pedido) AS SomaDeVA_Total_Pedido FROM TB_Mov_Estoque Where (DA_Pedido BETWEEN #"&data_de&"# AND #"&data_ate&"#)  AND TB_Mov_Estoque.NU_Unidade = 1 AND TB_Mov_Estoque.CO_Curso = '0' AND TB_Mov_Estoque.CO_Etapa = '"&etapa&"' GROUP BY TB_Mov_Estoque.NU_Unidade, TB_Mov_Estoque.CO_Curso, TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Projeto"	
				cols = 5	
				cols_tits = "UNIDADE#!#CURSO#!#ETAPA#!#PROJETO#!#TOTAL"	
				dados_buscados ="NU_Unidade#!#CO_Curso#!#CO_Etapa#!#CO_Projeto#!#SomaDeVA_Total_Pedido" 										
			end if
		else
			query = "SELECT TB_Mov_Estoque.CO_Projeto, Sum(TB_Mov_Estoque.VA_Total_Pedido) AS SomaDeVA_Total_Pedido FROM TB_Mov_Estoque Where (DA_Pedido BETWEEN #"&data_de&"# AND #"&data_ate&"#) AND TB_Mov_Estoque.NU_Unidade = 1 AND TB_Mov_Estoque.CO_Curso = '0' AND TB_Mov_Estoque.CO_Etapa = '"&etapa&"' AND TB_Mov_Estoque.CO_Turma = '"&turma&"' GROUP BY TB_Mov_Estoque.CO_Projeto"			
			cols = 2
				cols_tits = "PROJETO#!#TOTAL"	
				dados_buscados ="CO_Projeto#!#SomaDeVA_Total_Pedido" 		
										
		end if	
elseif tipo = "it" then
		if modalidade = "eg" then
			if isnull(turma) or turma = "" or turma = "999990" then
				if isnull(etapa) or etapa = "" or etapa = "999990" then
					query = "SELECT TB_Mov_Estoque.NU_Unidade, TB_Mov_Estoque.CO_Curso, TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma, TB_Mov_Estoque_Item.CO_Item, Sum(([QT_Solicitado]*[VA_Medio_Refer])) AS Expr1 FROM (TB_Mov_Estoque INNER JOIN TB_Mov_Estoque_Item ON TB_Mov_Estoque.NU_Pedido = TB_Mov_Estoque_Item.NU_Pedido) LEFT JOIN TB_Item ON TB_Mov_Estoque_Item.CO_Item = TB_Item.CO_Item GROUP BY TB_Mov_Estoque.NU_Unidade, TB_Mov_Estoque.CO_Curso, TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque.CO_Turma, TB_Mov_Estoque_Item.CO_Item ORDER BY Sum(([QT_Solicitado]*[VA_Medio_Refer])) DESC"					
					cols = 6					
				    cols_tits = "UNIDADE#!#CURSO#!#ETAPA#!#TURMA#!#ITEM#!#TOTAL"
					dados_buscados ="NU_Unidade#!#CO_Curso#!#CO_Etapa#!#CO_Turma#!#CO_Item#!#Expr1" 											
				else
					query = "SELECT TB_Mov_Estoque.NU_Unidade, TB_Mov_Estoque.CO_Curso, TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque_Item.CO_Item, Sum(([QT_Solicitado]*[VA_Medio_Refer])) AS Expr1 FROM (TB_Mov_Estoque INNER JOIN TB_Mov_Estoque_Item ON TB_Mov_Estoque.NU_Pedido = TB_Mov_Estoque_Item.NU_Pedido) LEFT JOIN TB_Item ON TB_Mov_Estoque_Item.CO_Item = TB_Item.CO_Item Where (DA_Pedido BETWEEN #"&data_de&"# AND #"&data_ate&"#) AND TB_Mov_Estoque.NU_Unidade = 1 AND TB_Mov_Estoque.CO_Curso = '0' AND TB_Mov_Estoque.CO_Etapa = '"&etapa&"'GROUP BY TB_Mov_Estoque.NU_Unidade, TB_Mov_Estoque.CO_Curso, TB_Mov_Estoque.CO_Etapa, TB_Mov_Estoque_Item.CO_Item ORDER BY Sum(([QT_Solicitado]*[VA_Medio_Refer])) DESC"						
					cols = 5
					cols_tits = "UNIDADE#!#CURSO#!#ETAPA#!#ITEM#!#TOTAL"	
					dados_buscados ="NU_Unidade#!#CO_Curso#!#CO_Etapa#!#CO_Item#!#Expr1" 												
				end if
			else
				query = "SELECT TB_Mov_Estoque_Item.CO_Item, Sum(([QT_Solicitado]*[VA_Medio_Refer])) AS Expr1 				FROM (TB_Mov_Estoque INNER JOIN TB_Mov_Estoque_Item ON TB_Mov_Estoque.NU_Pedido = TB_Mov_Estoque_Item.NU_Pedido) LEFT JOIN TB_Item ON TB_Mov_Estoque_Item.CO_Item = TB_Item.CO_Item Where (DA_Pedido BETWEEN #"&data_de&"# AND #"&data_ate&"#) AND TB_Mov_Estoque.NU_Unidade = 1 AND TB_Mov_Estoque.CO_Curso = '0' AND TB_Mov_Estoque.CO_Etapa = '"&etapa&"' AND TB_Mov_Estoque.CO_Turma = '"&turma&"' GROUP BY TB_Mov_Estoque_Item.CO_Item ORDER BY Sum(([QT_Solicitado]*[VA_Medio_Refer])) DESC"	
				cols = 2	
				cols_tits = "ITEM#!#TOTAL"	
				dados_buscados ="CO_Item#!#Expr1" 															
			end if		
		elseif modalidade = "pp" then
				query = "SELECT TB_Mov_Estoque.CO_Projeto, TB_Mov_Estoque_Item.CO_Item, Sum(([QT_Solicitado]*[VA_Medio_Refer])) AS Expr1 FROM (TB_Mov_Estoque INNER JOIN TB_Mov_Estoque_Item ON TB_Mov_Estoque.NU_Pedido = TB_Mov_Estoque_Item.NU_Pedido) LEFT JOIN TB_Item ON TB_Mov_Estoque_Item.CO_Item = TB_Item.CO_Item Where (DA_Pedido BETWEEN #"&data_de&"# AND #"&data_ate&"#)  GROUP BY TB_Mov_Estoque.CO_Projeto, TB_Mov_Estoque_Item.CO_Item ORDER BY Sum(([QT_Solicitado]*[VA_Medio_Refer])) DESC"
			cols = 3	
			cols_tits = "PROJETO#!#ITEM#!#TOTAL"
			dados_buscados ="CO_Projeto#!#CO_Item#!#Expr1" 												
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
					
'CABEÇALHO==========================================================================================		
			Set Param_Logo_Gde = Pdf.CreateParam
			margem=30		
			area_utilizavel=Page.Width - (margem*2)					
			largura_logo_gde=formatnumber(Logo.Width*0.3,0)
			altura_logo_gde=formatnumber(Logo.Height*0.3,0)
	
			Param_Logo_Gde("x") = margem
			Param_Logo_Gde("y") = Page.Height - altura_logo_gde -22
			Param_Logo_Gde("ScaleX") = 0.3
			Param_Logo_Gde("ScaleY") = 0.3
			Page.Canvas.DrawImage Logo, Param_Logo_Gde
	
		    x_texto=largura_logo_gde+ margem+10
			y_texto=formatnumber(Page.Height - altura_logo_gde/2,0)
		    width_texto=Page.Width -largura_logo_gde - 80			
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
			width_texto=Page.Width - (margem*2)

		
			SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
			Text = "<p><center><i><b><font style=""font-size:18pt;"">Apurar Despesas no Per&iacute;odo</font></b></i><br><font style=""font-size:10pt;"">De "&data_de_exibe&" at&eacute; "&data_ate_exibe&"</font></center></p>"
			

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

			y_primeira_tabela=Page.Height - altura_logo_gde
			x_primeira_tabela = margem

			Set param_table1 = Pdf.CreateParam("width=534; height=20; rows=1; cols="&cols&"; border=0; cellborder=0.1; cellspacing=0;")
			Set Table = Doc.CreateTable(param_table1)
			Table.Font = Font
			
			y_primeira_tabela=altura_primeiro_separador-10
			x_primeira_tabela=margem
			cel_width = formatnumber(534/cols,0)
			With Table.Rows(1)
			  ' .Cells(1).Height = 20
			  cols=cols*1
			  if cols=6 then
				 .Cells(1).Width = cel_width	
				 .Cells(2).Width = cel_width-20	
				 .Cells(3).Width = cel_width	
				 .Cells(4).Width = cel_width-50	
				 .Cells(5).Width = cel_width+70	
				 .Cells(6).Width = cel_width					 				 				 				 				 			  
			  else
				  for i = 1 to cols
				   .Cells(i).Width = cel_width	
				  next 		
			  end if	     		   		   		   		   
			End With

			titulos = split(cols_tits, "#!#")
			for i = 1 to cols
				c=i-1
				Table(1, i).AddText "<center><b>"&titulos(c)&"</b></center>", "size=8;indenty=2;html=true", Font 				
			Next	
			Set RSI = Server.CreateObject("ADODB.Recordset")
			SQLI = query
			response.Write(SQLI)
'			response.End()	
			RSI.Open SQLI, CON9											
			linha = 1		
			IF RSI.EOF THEN
			  For i = 1 to 20
				Set Row = Table.Rows.Add(15)				  
			  next
			else
				dados_cols = split(dados_buscados,"#!#")
				while not RSI.EOF 
					coluna = 0	
				    linha = linha+1
					Set Row = Table.Rows.Add(15)									
					For d = 0 to ubound(dados_cols)
						response.Write(dados_cols(d))
						dado = RSI(dados_cols(d))
					
						
						if dados_cols(d) = "TB_Mov_Estoque.NU_Unidade" or dados_cols(d) = "NU_Unidade" then	
						    unidade = dado
							nome_item = GeraNomes("U",unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)							
						elseif dados_cols(d) = "TB_Mov_Estoque.CO_Curso" or dados_cols(d) = "CO_Curso" then
						    curso = dado	
							nome_item = GeraNomes("C",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro)							
						elseif dados_cols(d) = "TB_Mov_Estoque.CO_Etapa" or dados_cols(d) = "CO_Etapa" then	
						    co_etapa = dado
							nome_item = GeraNomes("E",curso,co_etapa,variavel3,variavel4,variavel5,CON0,outro)							
						elseif dados_cols(d) = "TB_Mov_Estoque.CO_Turma" or dados_cols(d) = "CO_Turma" then	
							nome_item = dado																														
						elseif dados_cols(d) = "TB_Mov_Estoque.CO_Projeto" or dados_cols(d) = "CO_Projeto" then
							Set RSC = Server.CreateObject("ADODB.Recordset")
							SQLC = "Select NO_Projeto From TB_Projeto where CO_Projeto = "&dado
							RSC.Open SQLC, CON9	
							
							nome_item = RSC("NO_Projeto")								
							nome_item = replace_latin_char(nome_item,"html")						
						elseif dados_cols(d) = "TB_Mov_Estoque_Item.CO_Item" or dados_cols(d) = "CO_Item"  then
							Set RSC = Server.CreateObject("ADODB.Recordset")
							SQLC = "Select NO_Item From TB_Item where CO_Item = "&dado
							RSC.Open SQLC, CON9	
							
							nome_item = RSC("NO_Item")						
							nome_item = replace_latin_char(nome_item,"html")
						else
							if isnumeric(dado) then
								nome_item = formatcurrency(dado)
							else
								nome_item = replace_latin_char(nome_item,"html")							
							end if								
						end if
						coluna = coluna+1	
						
						Table(linha, coluna).AddText "<center>"&nome_item&"</center>", "size=7;indentx=2;indenty=2;html=true", Font 		
				   next				   
				RSI.MOVENEXT
				WEND 	
			 END IF	 		
'			'Table(36, 1).AddText "<b>Assinatura do Requerente:</b> ____________________________________________________________________________________________", "size=8;indenty=2;html=true", Font 										
			Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 	
'		y_assinatura = margem*2
'		 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&y_assinatura&"; height=50; width="&area_utilizavel&"; alignment=left; html=true; size=8; color=#000000")				
'			
'		assinatura = "<b>Assinatura do Requerente:</b> ____________________________________________________________________________________________"
'		Do While Len(assinatura) > 0
'			CharsPrinted = Page.Canvas.DrawText(assinatura, Param_Relatorio, Font )
'		 
'			If CharsPrinted = Len(assinatura) Then Exit Do
'			   SET Page = Page.NextPage
'			assinatura = Right( assinatura, Len(assinatura) - CharsPrinted)
'		Loop 						

		 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
		Relatorio = arquivo&" - Sistema Web Diretor"
		Do While Len(Relatorio) > 0
			CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
		 
			If CharsPrinted = Len(Relatorio) Then Exit Do
			   SET Page = Page.NextPage
			Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
		Loop 
		
		
		 SET Param_Relatorio = Pdf.CreateParam("x=450;y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")		
		Param_Relatorio.Add "html=true" 
		
		data_hora = "Impresso em "&data &" &agrave;s "&horario&""
		Do While Len(Relatorio) > 0
			CharsPrinted = Page.Canvas.DrawText(data_hora, Param_Relatorio, Font )
		 
			If CharsPrinted = Len(data_hora) Then Exit Do
			   SET Page = Page.NextPage
			data_hora = Right( data_hora, Len(data_hora) - CharsPrinted)
		Loop 							
Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

