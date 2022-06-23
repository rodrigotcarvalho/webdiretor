<%'On Error Resume Next%>
<%
'Server.ScriptTimeout = 60
Server.ScriptTimeout = 600 'valor em segundos
'Notas Fiscais de Compra
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/funcoes2.asp"-->

<% 
arquivo="SWD021"
response.Charset="ISO-8859-1"
obr= request.QueryString("obr")

dados = split(obr, "$!$")
cod_nf = dados(0)
dia_de = dados(1)
mes_de = dados(2)
ano_de = dados(3)
dia_ate = dados(4)
mes_ate = dados(5)
ano_ate = dados(6)
ordem = dados(7)

if cod_nf = "" or isnull(cod_nf) then
	nota_exibe = "Todas"
	sql_cod = ""	
else
	nota_exibe = cod_nf
	sql_cod = "NU_NotaF ='"& cod_nf&"' AND "	
end if


Select case ordem

case "dt"
ordena="DA_NotaF"
ordena_exibe="Data da Compra"

case "nf"
ordena="NU_NotaF"
ordena_exibe="Nota Fiscal"

case "fr"
ordena="NO_Fornecedor"
ordena_exibe="Fornecedor"

case "vn"
ordena="VA_NotaF"
ordena_exibe="Valor da Nota"

end select	


select case mes_de
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

select case mes_ate
 case 0 
 mes_b = "Todos"
 case 1 
 mes_b = "Janeiro"
 case 2 
 mes_b = "Fevereiro"
 case 3 
 mes_b = "Mar&ccedil;o"
 case 4
 mes_b = "Abril"
 case 5
 mes_b = "Maio"
 case 6 
 mes_b = "Junho"
 case 7
 mes_b = "Julho"
 case 8 
 mes_b = "Agosto"
 case 9 
 mes_b = "Setembro"
 case 10 
 mes_b = "Outubro"
 case 11 
 mes_b = "Novembro"
 case 12 
 mes_b = "Dezembro"
end select	


dia_de=dia_de*1

if dia_de<10 then
	dia_de_exibe="0"&dia_de
else
	dia_de_exibe=dia_de
end if

dia_ate=dia_ate*1

if dia_ate<10 then
	dia_ate_exibe="0"&dia_ate
else
	dia_ate_exibe=dia_ate
end if


data_de_exibe =dia_de_exibe&"/"&mes_a&"/"&ano_de

data_ate_exibe =dia_ate&"/"&mes_b&"/"&ano_ate

periodo = " de "&data_de_exibe&" at&eacute; "&data_ate_exibe


data_de=mes_de&"/"&dia_de&"/"&ano_de


data_ate=mes_ate&"/"&dia_ate&"/"&ano_ate


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON9 = Server.CreateObject("ADODB.Connection") 
		ABRIR9 = "DBQ="& CAMINHO_ax & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON9.Open ABRIR9



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
	Set Logo1 = Doc.OpenImage( Server.MapPath( "../img/logo_niteroi_preto.gif") )
	Set Logo2 = Doc.OpenImage( Server.MapPath( "../img/logo_arariboia_preto.gif") )	
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
						

	
	SET Page = Doc.Pages.Add(595,842)
			
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

	y_texto=y_texto-altura_logo_gde+30
	SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height=30; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
	Text = "<center><i><b><font style=""font-size:18pt;"">Notas Fiscais de Compra</font></b></i></center>"
	
	
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

	Set param_table1 = Pdf.CreateParam("width=547; height=20; rows=1; cols=6; border=0; cellborder=0; cellspacing=0;")
	Set Table = Doc.CreateTable(param_table1)
	Table.Font = Font
	y_primeira_tabela=altura_segundo_separador-10
	x_primeira_tabela=margem+5
	With Table.Rows(1)
	   .Cells(1).Width = 60
	   .Cells(2).Width = 100  
	   .Cells(3).Width = 50 		 
	   .Cells(4).Width = 200 
	   .Cells(5).Width = 50 
	   .Cells(6).Width = 100 	   	   	    		   		   
	End With
	
	
	Table(1, 1).AddText "<b>Nota Fiscal:</b>", "size=9;html=true", Font 
	Table(1, 2).AddText "<div align=LEFT>"&nota_exibe&"</div>", "size=9;html=true", Font 
	Table(1, 3).AddText "<b>Per&iacute;odo:</b>", "size=9;html=true", Font 
	Table(1, 4).AddText "<div align=LEFT>"&periodo&"</div>", "size=9;html=true", Font 
	Table(1, 5).AddText "<b>Ordem:</b>", "size=9;html=true", Font 
	Table(1, 6).AddText "<div align=LEFT>"&ordena_exibe&"</div>", "size=9;html=true", Font 	
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

	colunas_de_notas=2
	total_de_colunas=6				
	altura_medias=20
	y_segunda_tabela=altura_terceiro_separador-10	
	Set param_table2 = Pdf.CreateParam("width=547; height="&altura_medias&"; rows=1; cols=6; border=0; cellborder=0.5; cellspacing=0; x="&margem&"; y="&y_segunda_tabela&"; MaxHeight=670")

	Set Notas_Tit = Doc.CreateTable(param_table2)
	Notas_Tit.Font = Font				
	largura_colunas=(547-80-60-50-120)/colunas_de_notas		
	
	With Notas_Tit.Rows(1)
	   .Cells(1).Width = 80
	   .Cells(2).Width = 50	
	   .Cells(3).Width = 120
	   .Cells(4).Width = 60			             
	   .Cells(5).Width = largura_colunas
	   .Cells(6).Width = largura_colunas
	End With
	Notas_Tit(1, 1).AddText "<div align=""center""><b>Nota Fiscal</b></div>", "size=9;indenty=2; html=true", Font 
	Notas_Tit(1, 2).AddText "<div align=""center""><b>Data Compra</b></div>", "size=9;alignment=center; indenty=2;html=true", Font 
	Notas_Tit(1, 3).AddText "<div align=""center""><b>Fornecedor</b></div>", "size=9;alignment=center; indenty=2;html=true", Font 
	Notas_Tit(1, 4).AddText "<div align=""center""><b>Valor Nota</b></div>", "size=9;alignment=center; indenty=2;html=true", Font 
	Notas_Tit(1, 5).AddText "<div align=""center""><b>Conferido Por</b></div>", "size=9;alignment=center; indenty=2;html=true", Font 
	Notas_Tit(1, 6).AddText "<div align=""center""><b>Registrado Por</b></div>", "size=9;alignment=center; indenty=2;html=true", Font 
	Set param_materias = PDF.CreateParam	
	param_materias.Set "size=7;expand=false" 			

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_NFiscais_Compra, TB_Fornecedor WHERE TB_Fornecedor.CO_Fornecedor = TB_NFiscais_Compra.CO_Fornecedor AND "&sql_cod&"(DA_NotaF BETWEEN #"&data_de&"# AND #"&data_ate&"#) order BY "&ordena
		RS.Open SQL, CON9	
	
		param_materias.Add "expand=true;html=true" 		
		linha = 2
if RS.EOF	then	
		Notas_Tit(linha, 1).ColSpan = 6								
		Notas_Tit(linha, 1).AddText "<div align=""center"">Nenhuma nota fiscal encontrada para os crit&eacute;rios informados</div>", param_materias			
else
	while not RS.EOF
		co_nf=RS("NU_NotaF")
		da_nf=RS("DA_NotaF")
		co_fornecedor=RS("CO_Fornecedor")
		valor_nf=RS("VA_NotaF")
		observacao=RS("TX_Observa")
		co_usu_conf=RS("CO_Usuario_Conf")
		co_usu_reg=RS("CO_Usuario_Reg")
		
		
		if co_fornecedor="" or isnull(co_fornecedor) then
			no_fornecedor=""
		else
		
			Set RSnom = Server.CreateObject("ADODB.Recordset")
			SQLnom = "SELECT NO_Fornecedor FROM TB_Fornecedor Where CO_Fornecedor="&co_fornecedor
			RSnom.Open SQLnom, CON9
			
			if RSnom.EOF then
				no_fornecedor=""	
			else
				no_fornecedor=RSnom("NO_Fornecedor")
			end if	
		end if
		
		
		if co_usu_conf="" or isnull(co_usu_conf) then
			no_conferidor=""
		else
		
				Set RSu = Server.CreateObject("ADODB.Recordset")
				SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& co_usu_conf
				RSu.Open SQLu, CON
		
			IF RSu.EOF then
				no_conferidor=""
			else
				no_conferidor=RSu("NO_Usuario")
			end if		
		end if
				
		if co_usu_reg="" or isnull(co_usu_reg) then
			no_registrador=""
		else
		
				Set RSu = Server.CreateObject("ADODB.Recordset")
				SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& co_usu_reg
				RSu.Open SQLu, CON
		
			IF RSu.EOF then
				no_registrador=""
			else
				no_registrador=RSu("NO_Usuario")
			end if		
		end if		
		
		data_split= Split(da_nf,"/")
		dia=data_split(0)
		mes=data_split(1)
		ano=data_split(2)
		
		
		dia=dia*1
		
		mes=mes*1
		hora=hora*1
		min=min*1
		
		if dia<10 then
		dia="0"&dia
		end if
		if mes<10 then
		mes="0"&mes
		end if
		if hora<10 then
		hora="0"&hora
		end if
		if min<10 then
		min="0"&min
		end if
		da_show=dia&"/"&mes&"/"&ano
						
		Set Row = Notas_Tit.Rows.Add(17) ' row height		
		Notas_Tit(linha, 1).AddText "<div align=""center"">"&co_nf&"</div>", param_materias			
		Notas_Tit(linha, 2).AddText "<div align=""center"">"&da_show&"</div>", param_materias			
		Notas_Tit(linha, 3).AddText "<div align=""center"">"&no_fornecedor&"</div>", param_materias
		param_materias.Add "expand=false" 	
		Notas_Tit(linha, 4).AddText "<div align=""center"">"&formatnumber(valor_nf,2)&"</div>", param_materias	
		Notas_Tit(linha, 5).AddText "<div align=""center"">"&no_conferidor&"</div>", param_materias																						
		Notas_Tit(linha, 6).AddText "<div align=""center"">"&no_registrador&"</div>", param_materias																						

		linha = linha+1		
	RS.MOVENEXT
	WEND		
END IF

	limite=0
	Paginacao = 0				
	Do While True
		limite=limite+1
		Paginacao = Paginacao+1
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

	y_texto=y_texto-altura_logo_gde+30
	SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height=30; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
	Text = "<center><i><b><font style=""font-size:18pt;"">Notas Fiscais de Compra</font></b></i></center>"
	
	
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
	Table(1, 3).AddText "<b>M&ecirc;s:</b>", "size=9;html=true", Font 
	Table(1, 4).AddText "<div align=LEFT>"&mes_a&"</div>", "size=9;html=true", Font 
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
			Loop 	
	
								

	

Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

