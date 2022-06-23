﻿<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/bd_parametros.asp"-->
<!--#include file="../inc/bd_alunos.asp"-->
<!--#include file="../inc/bd_webfamilia.asp"-->
<!--#include file="../inc/funcoes_contratos.asp"-->
<%
Server.ScriptTimeout = 1800 'valor em segundos

SET Pdf = Server.CreateObject("Persits.Pdf")
SET Doc = Pdf.CreateDocument
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
	


ano_letivo = session("ano_letivo")
session("ano_letivo") = ano_letivo

cod_aluno= request.QueryString("c")

tipo = request.queryString("t")
session("tipo_pdf") = tipo

ucet = buscaUCET(cod_aluno,session("ano_letivo"))
vetorUCET = split(ucet,"#!#")

tp_contrato_adendo = modeloContratoAdendo(vetorUCET(0),vetorUCET(1),vetorUCET(2),vetorUCET(3),tipo)
'response.write(tp_contrato_adendo)
'response.end()

tp_contrato_adendo = replace(tp_contrato_adendo," ","_")

arquivo = "SWD300"&tp_contrato_adendo



session("aluno_contrato") = cod_aluno   
session("tipo_contrato") = tp_contrato_adendo   
if cod_aluno= "" or isnull(cod_aluno) then
'cod_aluno=2344
end if
if tp_contrato_adendo= "" or isnull(tp_contrato_adendo) then
	tipo="tp_contrato_adendo"
end if

SET Page = Doc.Pages.Add(595,842)
margem=25	
area_utilizavel=Page.Width - (margem*2)	

if tipo = "A" then
	Set Logo = Doc.OpenImage( Server.MapPath( "../img/logo_pdf.gif") )
		Set Param_Logo_Gde = Pdf.CreateParam			
		largura_logo_gde=formatnumber(Logo.Width*0.5,0)
		altura_logo_gde=formatnumber(Logo.Height*0.5,0)

		Param_Logo_Gde("x") = margem
		Param_Logo_Gde("y") = Page.Height - altura_logo_gde - margem
		Param_Logo_Gde("ScaleX") = 0.5
		Param_Logo_Gde("ScaleY") = 0.5
		Page.Canvas.DrawImage Logo, Param_Logo_Gde

		x_texto=largura_logo_gde+ margem+10
		y_texto=formatnumber(Page.Height - margem,0)
		width_texto=Page.Width -largura_logo_gde - 80


		'SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
'		Text = "<p><font style=""font-size:10pt;"">"&geraIdentificacaoEscola(tp_contrato_adendo,"Tabela")&"</FONT>" 			
'		
'		Do While Len(Text) > 0
'			CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
'		 
'			If CharsPrinted = Len(Text) Then Exit Do
'				SET Page = Page.NextPage
'			Text = Right( Text, Len(Text) - CharsPrinted)
'		Loop 	

end if
		
	if tp_contrato_adendo = "ADENDO_G1" or tp_contrato_adendo = "ADENDO_G2" then
		y_cabecalho = Page.Height - margem -20		
	else
		y_cabecalho = Page.Height - margem -20
	end if				
SET Param = Pdf.CreateParam("x="&margem&";y="&y_cabecalho&"; height=180; width="&area_utilizavel&"; alignment=center; size=10; color=#000000; html=true")

Text = dadosCabecalho(tp_contrato_adendo, cod_aluno) 			

Do While Len(Text) > 0
	CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
 
	If CharsPrinted = Len(Text) Then Exit Do
		SET Page = Page.NextPage
	Text = Right( Text, Len(Text) - CharsPrinted)
Loop 
			
	 
	Set conc = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO_contrato & ";Driver={Microsoft Access Driver (*.mdb)}"
	conc.Open ABRIR
	
	contrato_adendo = tp_contrato_adendo
	if tipo = "C" then
		if session("ano_letivo")>=2017 then
			posicao = 9
		else
			posicao = 10	
		end if	
		
		if tp_contrato_adendo = "CONTRATO_3" or tp_contrato_adendo = "CONTRATO1A" or tp_contrato_adendo = "CONTRATO2A" or tp_contrato_adendo = "CONTRATO2B" or tp_contrato_adendo = "CONTRATO3" or tp_contrato_adendo = "CONTRATO8" then
			y_paragrafo = Page.Height - 190
		else
			y_paragrafo = Page.Height - 200
		end if	
		y_rodape = 200			
	else
		posicao = 8	
		if session("ano_letivo")>=2017 then
			tp_contrato_adendo = converte_nome_adendo(vetorUCET(0),vetorUCET(1),vetorUCET(2),vetorUCET(3),tp_contrato_adendo)
		end if	

		if tp_contrato_adendo = "ADENDO_G1" or tp_contrato_adendo = "ADENDO_G2" then
			y_paragrafo = Page.Height - 160			
		else
			y_paragrafo = Page.Height - 140		
		end if		


		y_rodape = 70			
	end if
'response.Write(tp_contrato_adendo)
'response.End()
	seq = Mid(tp_contrato_adendo,posicao, 1)
	if isnumeric(seq) then
		posicao = posicao+1
		complemento = Mid(tp_contrato_adendo,posicao, 1)			
	else
		grupo=seq
		posicao = posicao+1		
		seq = Mid(tp_contrato_adendo,posicao, 1)
		posicao = posicao+1
		complemento = Mid(tp_contrato_adendo,posicao, 1)
		sql_grupo = "Ind_Grupo = '"&grupo&"' and "
	end if
		
		Set RSC = Server.CreateObject("ADODB.Recordset")
		SQLC = "SELECT Num_Seq_Contrato_Adendo FROM TB_Tipo_Contrato_Adendo where Num_Ano_Letivo="&ano_letivo+1&" And TP_Relatorio = '"&tipo&"' and "&sql_grupo&"Num_Sequencial="&seq&" and Ind_Complemento = '"&complemento&"'"
'				response.Write(SQLC)
'				response.End()
		RSC.Open SQLC, conc

'response.Write(RSC.EOF)
'response.End()	
	if RSC.EOF then
		response.redirect("../ws/mat/man/eco/altera.asp?opt=err2&cod="&cod_aluno)	
	end if	
		
	Set RSP = Server.CreateObject("ADODB.Recordset")
	SQLP = "SELECT Num_Paragrafo, Dsc_Paragrafo FROM TB_Paragrafos where Num_Seq_Contrato_Adendo="&RSC("Num_Seq_Contrato_Adendo")
'response.write(SQLP)
	RSP.Open SQLP, conc
	y_tabela = 	y_paragrafo - 300
	IF SEQ=8 Then
		y_tabela = 	y_paragrafo - 180
	end if	
'response.End()	
	while not RSP.EOF
		SET Param = Pdf.CreateParam("x="&margem&";y="&y_paragrafo&"; height=640; width="&area_utilizavel&"; alignment=center; size=10; color=#000000; html=true")	
		wrk_paragrafo = replace(replace(RSP("Dsc_Paragrafo"),"<p>",""),"</p>","<br>")
		Do While Len(wrk_paragrafo) > 0
				CharsPrinted = Page.Canvas.DrawText(wrk_paragrafo, Param, Font )
			 
				If CharsPrinted = Len(wrk_paragrafo) Then Exit Do
					SET Page = Page.NextPage
				SET Param = Pdf.CreateParam("x="&margem&";y="&Page.Height-margem&"; height=640; width="&area_utilizavel&"; alignment=center; size=10; color=#000000; html=true; expand=true")						
				wrk_paragrafo = Right( wrk_paragrafo, Len(wrk_paragrafo) - CharsPrinted)
			Loop 

		'Verifica se existe tabela a ser gerada	
		Set RST = Server.CreateObject("ADODB.Recordset")
		SQLT = "SELECT * FROM TB_Tabelas where Num_Seq_Contrato_Adendo="&RSC("Num_Seq_Contrato_Adendo")&" AND Num_Tabela =" &RSP("Num_Paragrafo")
		RST.Open SQLT, conc
			'response.Write(SQLT)
			'response.end()
		if not RST.EOF then
        	table_height = RST("Qtd_Linhas")*15		
			Set param_table = Pdf.CreateParam("width="&area_utilizavel&"; height="&table_height&"; rows="&RST("Qtd_Linhas")&"; cols="&RST("Qtd_Colunas")&"; border=0; cellborder=0.1; cellspacing=0; x="&margem&"; y="&y_tabela&"; MaxHeight=650;")	
			
			Set Table = Doc.CreateTable(param_table)
			Table.Font = Font		
			'response.Write("SELECT * FROM TB_Celulas where Num_Seq_Contrato_Adendo="&RSC("Num_Seq_Contrato_Adendo")&" AND Num_Tabela ="&RST("Num_Tabela")&" "&y_tabela&"<BR>")
'response.End()
			Set RSCL = Server.CreateObject("ADODB.Recordset")		
			SQLCL = "SELECT * FROM TB_Celulas where Num_Seq_Contrato_Adendo="&RSC("Num_Seq_Contrato_Adendo")&" AND Num_Tabela ="&RST("Num_Tabela")
'			response.Write(SQLCL)
'			response.end()
			RSCL.Open SQLCL, conc	
			contar=0
			while not RSCL.EOF 
		contar=contar+1
				if RSCL("Row_Span")>0 and RSCL("Col_Span")>0 then
					linha = RSCL("Num_Linha")			
					coluna = RSCL("Num_Coluna")
					rowspan = RSCL("Row_Span")
					colspan = RSCL("Col_Span")

					if rowspan>1 then					
						Table(linha, coluna).RowSpan = rowspan
					end if		
					if colspan>1 then									
						Table(linha, coluna).ColSpan = colspan
					end if		
					
					texto = RSCL("Txt_Celula")
					response.Write(contar&"                "&linha&","&coluna&"("&rowspan&","&colspan&")<BR>"&texto )
				
					Table(linha, coluna).AddText texto, Param 

				end if
			RSCL.MOVENEXT
			wend	
		
			Do While True
				limite=limite+1
				Paginacao = Paginacao+1
			   LastRow = Page.Canvas.DrawTable( Table, param_table )
				if LastRow >= Table.Rows.Count Then 
					Exit Do ' entire table displayed
				end if
			loop					   	
		end if
		if tp_contrato_adendo = "ADENDO_3" then
			if RSP("Num_Paragrafo") = 1 then
				y_paragrafo = y_tabela - 150	
			elseif RSP("Num_Paragrafo") = 2 then
				y_paragrafo = y_paragrafo - 70					
			else
				y_paragrafo = y_paragrafo - 2*margem
			end if	
		elseif tp_contrato_adendo = "ADENDO_5" then
			if RSP("Num_Paragrafo") = 1 then
				y_paragrafo = y_tabela - 150				
			else
				y_paragrafo = y_paragrafo - 2*margem
			end if					
		elseif tp_contrato_adendo = "ADENDO_8" then
			if RSP("Num_Paragrafo") = 1 then
				y_paragrafo = y_tabela - 110				
			else
				y_paragrafo = y_paragrafo - 2*margem
			end if		
		else
			if RSP("Num_Paragrafo") = 1 then
				y_paragrafo = y_tabela - 90		
			else
				y_paragrafo = y_paragrafo - 2*margem
			end if	
		end if	
		y_tabela = y_paragrafo
		response.Write(y_tabela &"="&y_paragrafo&"<BR>")
		
	'Fim da geração da tabela
	Param.Add "y="&y_tabela
	RSP.MOVENEXT
	WEND
	
SET Param = Pdf.CreateParam("x="&margem&";y="&y_rodape&"; height=175; width="&area_utilizavel&"; alignment=center; size=9; color=#000000; html=true")	

Text = dadosRodape(contrato_adendo, cod_aluno)

			Do While Len(Text) > 0
				CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
				If CharsPrinted = Len(Text) Then Exit Do
					SET Page = Page.NextPage
				Text = Right( Text, Len(Text) - CharsPrinted)
				
			Loop 
'Set Pdf = Server.CreateObject("Persits.Pdf")
'Set Doc = Pdf.CreateDocument
'if ambiente_wf ="" or isnull(ambiente_wf) then
'	ambiente = ambiente_escola
'else
'	ambiente = ambiente_wf
'end if	
'if Left(ambiente,5) <> "teste" then
'urlPDF = "http://www.simplynet.com.br/wd/"&ambiente&"/relatorios/geraContratoAdendo.asp?ano="&ano_letivo&"&aluno="&cod_aluno&"&tipo="&contrato_adendo
'else
'urlPDF = "http://www.simplynet.com.br/wdteste/"&ambiente&"/relatorios/geraContratoAdendo.asp?ano="&ano_letivo&"&aluno="&cod_aluno&"&tipo="&contrato_adendo
'end if
'response.Write(urlPDF)
'Doc.ImportFromUrl urlPDF, "scale=0.9; hyperlinks=true; drawbackground=true;PageWidth=595;PageHeight=860;Debug=true;LeftMargin=28; RightMargin=28"
''Filename = Doc.Save("e:\home\simplynet\dados\testedinamis\BD\contrato.pdf", False )
'			response.End()
	Doc.SaveHttp("attachment; filename="&arquivo&".pdf")   
%>