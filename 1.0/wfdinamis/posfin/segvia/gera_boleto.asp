<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 60 'valor em segundos

%>
<!--#include file="../../inc/caminhos.asp"-->
<!--#include file="../../inc/funcoes2.asp"-->
<!--#include file="../../inc/funcoes6.asp"-->
<% 
response.Charset="ISO-8859-1"
dados = request.form("vencimento")
cod_cons = request.form("cod")
vetor_meses = split(dados,", ")

ano_letivo = session("ano_letivo")

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

	'Dim AspPdf, Doc, Page, Font, Text, Param, Image, CharsPrinted
	'Instancia o objeto na memória
	SET Pdf = Server.CreateObject("Persits.Pdf")
	SET Doc = Pdf.CreateDocument
	Set Logo = Doc.OpenImage( Server.MapPath( "../../img/logo_boleto.gif") )
	Set Font = Doc.Fonts.LoadFromFile(Server.MapPath("../../fonts/arial.ttf"))	
	Set Font_Tesoura = Doc.Fonts.LoadFromFile(Server.MapPath("../../fonts/ZapfDingbats.ttf"))
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

		Set CONBL = Server.CreateObject("ADODB.Connection") 
		ABRIRBL = "DBQ="& CAMINHO_bl & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONBL.Open ABRIRBL
		
		Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4 = "DBQ="& CAMINHO_pf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4		

	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod_cons
	RS.Open SQL, CON1
	
	nome_aluno = RS("NO_Aluno")
	sexo_aluno = RS("IN_Sexo")
	
	nome_aluno=replace_latin_char(nome_aluno,"html")	
	
	if sexo_aluno="F" then
		desinencia="a"
	else
		desinencia="o"
	end if


	Set RS1 = Server.CreateObject("ADODB.Recordset")
	SQL1 = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod_cons
	RS1.Open SQL1, CON1
	
	if RS1.EOF then
		response.redirect("index.asp?nvg="&nvg&"&opt=err1")
	else
	
		ano_aluno = RS1("NU_Ano")
		rematricula = RS1("DA_Rematricula")
		situacao = RS1("CO_Situacao")
		encerramento= RS1("DA_Encerramento")
		unidade= RS1("NU_Unidade")
		curso= RS1("CO_Curso")
		etapa= RS1("CO_Etapa")
		turma= RS1("CO_Turma")
		cham= RS1("NU_Chamada")
			
		call GeraNomes("PORT",unidade,curso,etapa,CON0)
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Unidade WHERE NU_Unidade="& unidade
		RS2.Open SQL2, CON0
		
		if RS2.EOF then
			no_unidade = ""
			co_cnpj = ""
		else				
			no_unidade = RS2("TX_Imp_Cabecalho")	
			co_cnpj = RS2("CO_CGC")			
		end if
		no_curso= session("no_grau")
		no_etapa = session("no_serie")
				
			
						
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Curso WHERE CO_Curso='"& curso &"'"
		RS3.Open SQL3, CON0
		
		no_abrv_curso = RS3("NO_Abreviado_Curso")
		co_concordancia_curso = RS3("CO_Conc")	
		
		no_unidade = unidade&" - "&no_unidade
		'no_curso= no_etapa&" "&co_concordancia_curso&" "&no_curso
		no_curso= no_curso&" - "&no_etapa
		'no_etapa = no_etapa&" "&co_concordancia_curso&" "&no_abrv_curso				
		for n=0 to ubound(vetor_meses)
			margem_x=20	
			margem_y=20		
			row_padrao=margem_y
			'if ubound(vetor_meses)mod 2 = 0 then
				'SET Page = Doc.Pages.Add( 595, 842 )
				'if ubound(vetor_meses) = n then	
				'	altura_inicial=421		
				'else
				'	altura_inicial=margem_y											
				'end if	
			'else
			'	altura_inicial=421			
			'end if	
'CABEÇALHO==========================================================================================		
			Set Param_Logo_Gde = Pdf.CreateParam					
			largura_logo_gde=Logo.Width 'formatnumber(Logo.Width*1,0)
			altura_logo_gde=Logo.Height 'formatnumber(Logo.Height*1,0)
			
			SET Page = Doc.Pages.Add( 595, 842 )
			altura_inicial=formatnumber((Page.Height/2)+(altura_logo_gde/2),0)		

			area_disponivel =  Page.Width - (2*margem_x)				
				
			Param_Logo_Gde("x") = margem_x
			Param_Logo_Gde("y") = Page.Height - altura_inicial
			Param_Logo_Gde("ScaleX") = 1
			Param_Logo_Gde("ScaleY") = 1
			Page.Canvas.DrawImage Logo, Param_Logo_Gde

			
			Set RS4 = Server.CreateObject("ADODB.Recordset")
			SQL4= "SELECT * FROM TB_Posicao WHERE VA_Realizado=0 AND CO_Matricula_Escola ="& cod_cons &" AND Mes = "&vetor_meses(n)
			RS4.Open SQL4, CON4	
		
			if RS4.EOF then
				response.redirect("index.asp?nvg="&nvg&"&opt=err2")
			else
				vencimento=RS4("DA_Vencimento")
				nu_cota=RS4("NU_Cota")
				
				vetor_vencimento = split(vencimento, "/")
				if ((((vetor_vencimento(1) = 1 or vetor_vencimento(1) = 3 or vetor_vencimento(1) = 5 or vetor_vencimento(1) = 7 or vetor_vencimento(1) = 8 or vetor_vencimento(1) = 10 or vetor_vencimento(1) = 12) and vetor_vencimento(0) = 31)   or   (vetor_vencimento(1) = 4 or vetor_vencimento(1) = 6 or vetor_vencimento(1) = 9 or vetor_vencimento(1) = 11) and vetor_vencimento(0) = 30)) then
					dia_vencimento = 1
					mes_vencimento = vetor_vencimento(1)+1
				elseif ((vetor_vencimento(1) = 2 and (vetor_vencimento(2) MOD 4 = 0) and  vetor_vencimento(0) = 29) or (vetor_vencimento(1) = 2 and  vetor_vencimento(0) = 28)) then
					dia_vencimento = 1
					mes_vencimento = vetor_vencimento(1)+1				
				else
					dia_vencimento = vetor_vencimento(0)+1
					mes_vencimento = vetor_vencimento(1)
				end if	
				
				if ((vetor_vencimento(1) = 12) and vetor_vencimento(0) = 31) then
					ano_vencimento = vetor_vencimento(2)+1			
				else
					ano_vencimento = vetor_vencimento(2)				
				end if 
				vencimento_fim = dia_vencimento&"/"&mes_vencimento&"/"&ano_vencimento
			end if
			
			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT * FROM TB_Bloqueto WHERE DA_Vencimento>=#"& vencimento &"# and DA_Vencimento<#"& vencimento_fim &"#   AND CO_Matricula_Escola ="& cod_cons
			RS1.Open SQL1, CONBL
			
			if RS1.EOF then		
				nu_carne = ""	
				nosso_numero = ""
				va_inicial = ""
			else	
				nu_carne=RS1("NU_Bloqueto")
				nosso_numero = RS1("CO_Nosso_Numero")
				va_inicial = RS1("VA_Inicial")
				cod_superior=RS1("CO_Superior")				
				cod_barras =RS1("CO_Barras")
			end if	

'			SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
'			Text = "<p><center><i><b><font style=""font-size:18pt;"">Boletim Escolar</font></b></i></center></p>"
'			
'
'			Do While Len(Text) > 0
'				CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
'			 
'				If CharsPrinted = Len(Text) Then Exit Do
'					SET Page = Page.NextPage
'				Text = Right( Text, Len(Text) - CharsPrinted)
'			Loop 


'================================================================================================================			
'			y_nome_aluno=Page.Height - altura_logo_gde-46
'			width_nome_aluno=Page.Width - margem_x
'			
'			SET Param_Nome_Aluno = Pdf.CreateParam("x="&margem_x&";y="&y_nome_aluno&"; height=50; width="&width_nome_aluno&"; alignment=left; html=True")
'			Nome = "<font style=""font-size:11pt;""><b>Alun"&desinencia&": "&nome_aluno&"</b></font>"
'			
'
'			Do While Len(Nome) > 0
'				CharsPrinted = Page.Canvas.DrawText(Nome, Param_Nome_Aluno, Font )
'			 
'				If CharsPrinted = Len(Nome) Then Exit Do
'					SET Page = Page.NextPage
'				Nome = Right( Nome, Len(Nome) - CharsPrinted)
'			Loop 
'			
'			Page.Canvas.SetParams "LineWidth=2" 
'			Page.Canvas.SetParams "LineCap=0" 
'			With Page.Canvas
'			   .MoveTo margem_x, Page.Height - altura_logo_gde-65
'			   .LineTo Page.Width - margem_x, Page.Height - altura_logo_gde-65
'			   .Stroke
'			End With 	

			x_tabela_1=(2*margem_x)+largura_logo_gde
			y_tabela_1=formatnumber(Page.Height - altura_inicial + altura_logo_gde,0)
			width_tabela_1=Page.Width - largura_logo_gde - (margem_x*3)

			Set param_table1 = Pdf.CreateParam("width="&width_tabela_1&"; height="&altura_logo_gde&"; rows=3; cols=4; border=1; cellborder=0.5; cellspacing=0;")
			Set Table = Doc.CreateTable(param_table1)
			Table.Font = Font
			
			With Table.Rows(3)
			   .Cells(1).Width = formatnumber(width_tabela_1/4,0)-10
			   .Cells(2).Width = formatnumber(width_tabela_1/4,0)-10
			   .Cells(3).Width = formatnumber(width_tabela_1/4,0)+30
			   .Cells(4).Width = formatnumber(width_tabela_1/4,0)-10  
			End With
			Table(1, 1).ColSpan = 4
			Table(2, 1).ColSpan = 2		
			Table(2, 3).ColSpan = 2				
			Table(1, 1).AddText "<sup><font style=""font-size:7pt;"">Alun"&desinencia&":</font></sup>&nbsp;&nbsp;&nbsp;<b>"&nome_aluno&"</b>", "size=9; html=true", Font 
			Table(2, 1).AddText "<sup><font style=""font-size:7pt;"">Matr&iacute;cula:</font></sup>&nbsp;&nbsp;&nbsp;"&cod_cons&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Turma:&nbsp;&nbsp;"&turma, "size=8; html=true", Font 
			Table(2, 3).AddText "<sup><font style=""font-size:7pt;"">Vencimento:</font></sup>&nbsp;&nbsp;&nbsp;"&vencimento&"", "size=8; html=true", Font 
			Table(3, 1).AddText "<sup><font style=""font-size:7pt;"">N&ordm; Cota:</font></sup>&nbsp;&nbsp;&nbsp;"&nu_cota&"", "size=8; html=true", Font 

			Table(3, 2).AddText "<sup><font style=""font-size:7pt;"">N&ordm; Carn&ecirc;:</font></sup>&nbsp;&nbsp;&nbsp;"&nu_carne&"", "size=8; html=true", Font 
			Table(3, 3).AddText "<sup><font style=""font-size:7pt;"">Nosso N&uacute;mero:</font></sup>&nbsp;&nbsp;&nbsp;"&nosso_numero&"", "size=8; html=true", Font 		
			Table(3, 4).AddText "<sup><font style=""font-size:7pt;"">Valor Cobrado:</font></sup>&nbsp;&nbsp;&nbsp;"&formatcurrency(va_inicial)&"", "size=8; html=true", Font 					
			'Table(2, 3).AddText no_etapa, "size=9;", Font 
'			Table(2, 4).AddText "Turma: "&turma, "size=9;", Font 
'			Table(2, 5).AddText "N&ordm;. Chamada: "&cham, "size=9; html=true", Font 
'			Table(2, 6).AddText cham, "size=9;", Font 
'			Table(1, 7).AddText "<div align=""right"">Matr&iacute;cula: </div>", "size=9; html=true", Font 
'			Table(1, 8).AddText cod_cons, "size=9;alignment=right", Font 
'			Table(2, 7).AddText "Ano Letivo: ", "size=9; alignment=right", Font 
'			Table(2, 8).AddText ano_letivo, "size=9;alignment=right", Font 
			Page.Canvas.DrawTable Table, "x="&x_tabela_1&", y="&y_tabela_1&"" 
			
			
			x_texto_1 = page.width-(margem_x*4)
			y_texto_1 = y_tabela_1+(margem_y/2)
			width_texto_1 = width_tabela_1
			SET Param = Pdf.CreateParam("x="&x_texto_1&";y="&y_texto_1&"; height="&row_padrao&"; width="&width_texto_1&"; alignment=RIGHT; size=7; color=#000000; html=true")
			Text = "<p><i><b><font style=""font-size:7pt;"">Recibo do Aluno</font></b></i></p>"
			

			Do While Len(Text) > 0
				CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
			 
				If CharsPrinted = Len(Text) Then Exit Do
					SET Page = Page.NextPage
				Text = Right( Text, Len(Text) - CharsPrinted)
			Loop 
			
			vetor_cnpj=SPLIT(co_cnpj,"/")
			if ubound(vetor_cnpj)>0 then
				if vetor_cnpj(1)<0 then
					vetor_cnpj(1)=vetor_cnpj(1)*10
				end if
				exibe_cnpj="CNPJ: "&vetor_cnpj(0)&"/"&vetor_cnpj(1)
			end if		
			
			x_texto_2 = margem_x
			y_texto_2 = altura_inicial-altura_logo_gde
			width_texto_2 = largura_logo_gde+margem_x
			SET Param = Pdf.CreateParam("x="&x_texto_2&";y="&y_texto_2&"; height="&row_padrao&"; width="&width_texto_2&"; alignment=RIGHT; size=5.5; color=#000000; html=true")		
			
			Do While Len(exibe_cnpj) > 0
				CharsPrinted = Page.Canvas.DrawText(exibe_cnpj, Param, Font )
			 
				If CharsPrinted = Len(exibe_cnpj) Then Exit Do
					SET Page = Page.NextPage
				exibe_cnpj = Right( exibe_cnpj, Len(exibe_cnpj) - CharsPrinted)
			Loop 			
			
			texto_3 = "Autentica&ccedil;&atilde;o Mec&acirc;nica"
			
			x_texto_3 = margem_x
			y_texto_3 = y_texto_2-margem_y
			width_texto_3 = width_texto_2
			SET Param = Pdf.CreateParam("x="&x_texto_3&";y="&y_texto_3&"; height="&row_padrao&"; width="&width_texto_3&"; alignment=RIGHT; size=5.5; color=#000000; html=true")							
				
			Do While Len(texto_3) > 0
				CharsPrinted = Page.Canvas.DrawText(texto_3, Param, Font )
			 
				If CharsPrinted = Len(texto_3) Then Exit Do
					SET Page = Page.NextPage
				texto_3 = Right( texto_3, Len(texto_3) - CharsPrinted)
			Loop 
			
			Page.Canvas.SetParams "LineWidth=0.5; LineCap=0; Dash1=3; DashPhase=0" 
			x_primeira_linha= page.width-margem_x
			y_primeira_linha = y_texto_3-(margem_y*2)
				With Page.Canvas
				   .MoveTo margem_x, y_primeira_linha
				   .LineTo x_primeira_linha, y_primeira_linha
				   .Stroke
				End With 

			Page.Canvas.SetParams "Dash1=0; DashPhase=0"


			texto_4 = "<font style=""font-size:13pt;"">BANCO ITA&Uacute;</FONT><font style=""font-size:15pt;""><b> |341-7|</b></FONT> <font style=""font-size:13pt;"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&cod_superior&"</FONT>"
			
			x_texto_4 = margem_x
			y_texto_4 = y_primeira_linha-margem_y
			width_texto_4 = Page.Width - (margem_x*2)
			SET Param = Pdf.CreateParam("x="&x_texto_4&";y="&y_texto_4&"; height="&row_padrao&"; width="&width_texto_4&"; alignment=RIGHT; size=5.5; color=#000000; html=true")							
				
			Do While Len(texto_4) > 0
				CharsPrinted = Page.Canvas.DrawText(texto_4, Param, Font )
			 
				If CharsPrinted = Len(texto_4) Then Exit Do
					SET Page = Page.NextPage
				texto_4 = Right( texto_4, Len(texto_4) - CharsPrinted)
			Loop 


			x_tabela_2=margem_x
			y_tabela_2=formatnumber(y_texto_4 - margem_y,0)
			rows_tabela_2 = 11
			
			width_tabela_2=Page.Width - (margem_x*2)
			height_tabela_2 = rows_tabela_2*row_padrao

			Set param_table2 = Pdf.CreateParam("width="&width_tabela_2&"; height=220; rows="&rows_tabela_2&"; cols=7; border=1; cellborder=0.5; cellspacing=0;")
			Set Table = Doc.CreateTable(param_table2)
			Table.Font = Font

			
'			With Table.Rows(3)
'			   .Cells(1).Width = formatnumber(width_tabela_1/4,0)-10
'			   .Cells(2).Width = formatnumber(width_tabela_1/4,0)-10
'			   .Cells(3).Width = formatnumber(width_tabela_1/4,0)+30
'			   .Cells(4).Width = formatnumber(width_tabela_1/4,0)-10  
'			End With
			Table(1, 1).ColSpan = 6
'			Table(2, 1).ColSpan = 2		
'			Table(2, 3).ColSpan = 2				
'			Table(1, 1).AddText "<sup><font style=""font-size:7pt;"">Alun"&desinencia&":</font></sup>&nbsp;&nbsp;&nbsp;<b>"&nome_aluno&"</b>", "size=9; html=true", Font 
'			Table(2, 1).AddText "<sup><font style=""font-size:7pt;"">Matr&iacute;cula:</font></sup>&nbsp;&nbsp;&nbsp;"&cod_cons&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Turma:&nbsp;&nbsp;"&turma, "size=8; html=true", Font 
'			Table(2, 3).AddText "<sup><font style=""font-size:7pt;"">Vencimento:</font></sup>&nbsp;&nbsp;&nbsp;"&vencimento&"", "size=8; html=true", Font 
'			Table(3, 1).AddText "<sup><font style=""font-size:7pt;"">N&ordm; Cota:</font></sup>&nbsp;&nbsp;&nbsp;"&nu_cota&"", "size=8; html=true", Font 
'
'			Table(3, 2).AddText "<sup><font style=""font-size:7pt;"">N&ordm; Carn&ecirc;:</font></sup>&nbsp;&nbsp;&nbsp;"&nu_carne&"", "size=8; html=true", Font 
'			Table(3, 3).AddText "<sup><font style=""font-size:7pt;"">Nosso N&uacute;mero:</font></sup>&nbsp;&nbsp;&nbsp;"&nosso_numero&"", "size=8; html=true", Font 		
'			Table(3, 4).AddText "<sup><font style=""font-size:7pt;"">Valor Cobrado:</font></sup>&nbsp;&nbsp;&nbsp;"&formatcurrency(va_inicial)&"", "size=8; html=true", Font 					

			Page.Canvas.DrawTable Table, "x="&x_tabela_2&", y="&y_tabela_2&""
			
			x_barcode=margem_x + 19' A distância mínima da margem da ficha é de 5 mm
			y_barcode=formatnumber(y_tabela_2 - height_tabela_2 - 98,0) ' A distância mínima da ficha é de 12 mm (49 px de espaço +49 px de distância)
			width_barcode=389 ' o tamanho deve ser 103mm	
							  ' A altura deverá ser 13mm 
			strParam = "x="&x_barcode&"; y="&y_barcode&"; height=49; width="&width_barcode&"; type=12" 'Barcode type 1 is UPC-A
			strData = cod_barras
			Page.Canvas.DrawBarcode strData, strParam 			 			
'			With Page.Canvas
'			   .MoveTo margem_x, Page.Height - altura_logo_gde-100
'			   .LineTo Page.Width - margem_x, Page.Height - altura_logo_gde-100
'			   .Stroke
'			End With 
	
'			altura_medias=40
'			Set param_table2 = Pdf.CreateParam("width=533; height="&altura_medias&"; rows=3; cols=15; border=1; cellborder=0.1; cellspacing=0;")
'			Set Notas_Tit = Doc.CreateTable(param_table2)
'			Notas_Tit.Font = Font
'			y_medias=Page.Height - altura_logo_gde-110
'			
'			With Notas_Tit.Rows(1)
'			   .Cells(1).Width = 183
'			   .Cells(2).Width = 25
'			   .Cells(3).Width = 25
'			   .Cells(4).Width = 25
'			   .Cells(5).Width = 25
'			   .Cells(6).Width = 25
'			   .Cells(7).Width = 25
'			   .Cells(8).Width = 25   
'			   .Cells(9).Width = 25
'			   .Cells(10).Width = 25
'			   .Cells(11).Width = 25   
'			   .Cells(12).Width = 25 
'			   .Cells(13).Width = 25			         
'			   .Cells(14).Width = 25 			         
'			   .Cells(15).Width = 25 			         		         
'			   			         
'			End With
'			Notas_Tit(1, 1).RowSpan = 3	
'			Notas_Tit(2, 2).RowSpan = 2	
'			Notas_Tit(2, 3).RowSpan = 2	
'			Notas_Tit(2, 4).RowSpan = 2	
'			Notas_Tit(2, 5).RowSpan = 2	
'			Notas_Tit(2, 6).RowSpan = 2	
'			Notas_Tit(2, 7).RowSpan = 2	
'			Notas_Tit(2, 8).RowSpan = 2	
'			Notas_Tit(2, 9).RowSpan = 2	
'			Notas_Tit(2, 10).RowSpan = 2	
'			Notas_Tit(2, 11).RowSpan = 2	
'			Notas_Tit(2, 12).RowSpan = 2	
'			Notas_Tit(2, 13).RowSpan = 2
'			Notas_Tit(2, 14).RowSpan = 2
'			Notas_Tit(2, 15).RowSpan = 2
'			Notas_Tit(1, 2).ColSpan = 7	
'			Notas_Tit(1, 9).ColSpan = 3				
'			Notas_Tit(1, 12).ColSpan = 4																																																																						
'			Notas_Tit(1, 1).AddText "<div align=""center""><b>Disciplinas</b></div>", "size=10;indenty=15; html=true", Font 
'			Notas_Tit(1, 2).AddText "<div align=""center""><b>Aproveitamento</b></div>", "size=8;alignment=center; indenty=1;html=true", Font  
'			Notas_Tit(1, 9).AddText "<div align=""center""><b>M&eacute;dia (da Turma)</b></div>", "size=8;alignment=center; indenty=1;html=true", Font
'			Notas_Tit(1, 12).AddText "<div align=""center""><b>Frequ&ecirc;ncia</b></div>", "size=8;alignment=center; indenty=1;html=true", Font
'			Notas_Tit(2, 2).AddText "PA1", "size=7;alignment=center;indenty=6;", Font 
'			Notas_Tit(2, 3).AddText "PA2", "size=7;alignment=center;indenty=6;", Font 
'			Notas_Tit(2, 4).AddText "PA3", "size=7;alignment=center;indenty=6;", Font 
'			Notas_Tit(2, 5).AddText "Total", "size=7 ;alignment=center;indenty=6;", Font 
'			Notas_Tit(2, 6).AddText "<div align=""center"">4&ordf; Aval.<br>p.2</div>", "size=7;alignment=center;indenty=2; html=true", Font 
'			Notas_Tit(2, 7).AddText "Total", "size=7;alignment=center;indenty=6;", Font 
'			Notas_Tit(2, 8).AddText "<div align=""center"">M&eacute;dia<br>Final</div>", "size=7;alignment=center;indenty=2;html=true", Font 
'			Notas_Tit(2, 9).AddText "PA1", "size=7;alignment=center;indenty=6;", Font  
'			Notas_Tit(2, 10).AddText "PA2", "size=7;alignment=center;indenty=6;", Font 
'			Notas_Tit(2, 11).AddText "PA3", "size=7;alignment=center;indenty=6;", Font 
'			Notas_Tit(2, 12).AddText "PA1", "size=7;alignment=center;indenty=6;", Font 
'			Notas_Tit(2, 13).AddText "PA2", "size=7;alignment=center;indenty=6;", Font 
'			Notas_Tit(2, 14).AddText "PA3", "size=7;alignment=center;indenty=6;", Font 
'			Notas_Tit(2, 15).AddText "Total", "size=7;alignment=center;indenty=6;", Font 
'			
'			Set param_materias = PDF.CreateParam	
'			param_materias.Set "expand=true" 
'			
'			acumula_peso_periodo=0			
'		
'					
''response.Write(resultados)
''response.End()
'			Page.Canvas.DrawTable Notas_Tit, "x="&margem_x&", y="&y_medias&"" 		
'
'				if 	wrk_exibe_resultado="s" THEN
''					resultados=Calc_Med_An_Fin(unidade, curso, etapa, turma, cod_cons, vetor_materia, caminho_nota, tb_nota, 4, 4, 0, "final", 0)		
''					response.Write(unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&cod_cons&"-"&vetor_materia&"-"&caminho_nota&"-"&tb_nota&" -r")	
'
''response.Write(cod_cons&" <br>")
''response.Write(conta_notas1&"<"&conta_notas2&" <br>")				
'					response.Write(resultados1&" -1<br>")
'					response.Write(resultados2&" -2<br>")
'
'					if conta_notas1<conta_notas2 then
'						resultados=resultados1
'						tipo_resultado="A"
'					else
'						resultados=resultados2		
'						tipo_resultado="F"									
'					end if	
''response.Write(resultados&" <br>")					
'					resultados_apurados = split(resultados, "#%#" )	
'					
'					if ubound(resultados_apurados)=-1 then
'						resultado_final_aluno="nulo"
'					else				
'						resultado_final_aluno=apura_resultado_geral_aluno(curso, etapa, resultados_apurados(0),tipo_resultado)
'					end if
'	
''response.Write(resultado_final_aluno&" <br>")	
''					response.end()					
'					if (resultado_final_aluno="Apr" or resultado_final_aluno="APR") then
'						legenda_resultado="<b>Resultado Final:</b>"
'						resultado_exibe="Aprovado"
'					elseif (resultado_final_aluno="Rep" or resultado_final_aluno="REP") then	
'						legenda_resultado="<b>Resultado Final:</b>"				
'						resultado_exibe="Reprovado"	
'					'	legenda_resultado=""				
'					'	resultado_exibe=""					
'					elseif (resultado_final_aluno="Rec" or resultado_final_aluno="REC") then	
'						legenda_resultado="<b>Resultado Final:</b>"					
'						resultado_exibe="Recupera&ccedil;&atilde;o"	
'					'	legenda_resultado=""				
'					'	resultado_exibe=""						
'					else
'					'	legenda_resultado="<b>Resultado Final:</b>"					
'						legenda_resultado=""				
'						resultado_exibe=""										
'					end if	
'	
'					Set param_table4 = Pdf.CreateParam("width=540; height=125; rows=3; cols=1; border=0; cellborder=0; cellspacing=0;")
'					Set quadro = Doc.CreateTable(param_table4)
'					quadro.Font = Font
'					quadro.Rows(1).Height = 20
'					'quadro.Rows(2).Height = 60
'					'quadro.Rows(3).Height = 45
'			 
'	
'					quadro(1, 1).AddText legenda_resultado&"&nbsp;&nbsp;"&resultado_exibe, "size=9;indentx=2; html=true", Font 
'					'quadro(2, 1).AddText "<b>Observa&ccedil;&otilde;es:</b>", "size=9;indentx=5;indenty=5;html=true;", Font 
'	Page.Canvas.DrawTable quadro, "x="&margem_x&", y="&y_medias-20-altura_medias&"" 								
'				end if	

'LINHAS QUE DIVIDEM OS PERÍODOS DA TABELA========================================================================
'			rows_notas=Ubound(co_materia_exibe)+3
'			rows_notas=rows_notas*1
'			altura_linha_divisora_notas=(rows_notas*20)
'			y_fim_linha=y_medias-altura_linha_divisora_notas
'			
'			Page.Canvas.SetParams "LineWidth=1" 
'			x_primeira_linha=140+30
'				With Page.Canvas
'				   .MoveTo x_primeira_linha, y_medias
'				   .LineTo x_primeira_linha, y_fim_linha
'				   .Stroke
'				End With 
'				
'			x_segunda_linha=x_primeira_linha+(19*4)+1
'				With Page.Canvas
'				   .MoveTo x_segunda_linha, y_medias
'				   .LineTo x_segunda_linha, y_fim_linha
'				   .Stroke
'				End With 
'				
'			x_terceira_linha=x_segunda_linha+(19*4)
'				With Page.Canvas
'				   .MoveTo x_terceira_linha, y_medias
'				   .LineTo x_terceira_linha, y_fim_linha
'				   .Stroke
'				End With 	
'				
'			x_quarta_linha=x_terceira_linha+(19*4)
'				With Page.Canvas
'				   .MoveTo x_quarta_linha, y_medias
'				   .LineTo x_quarta_linha, y_fim_linha
'				   .Stroke
'				End With 	
'				
'			x_quinta_linha=x_quarta_linha+75
'				With Page.Canvas
'				   .MoveTo x_quinta_linha, y_medias
'				   .LineTo x_quinta_linha, y_fim_linha
'				   .Stroke
'				End With 	
'				
'			x_sexta_linha=x_quinta_linha+(25*2)
'				With Page.Canvas
'				   .MoveTo x_sexta_linha, y_medias
'				   .LineTo x_sexta_linha, y_fim_linha
'				   .Stroke
'				End With 				
			'=============================================================================================================

	
'			Set CON_N = Server.CreateObject("ADODB.Connection") 
'			ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
'			CON_N.Open ABRIRn
'						
'			Set RSF = Server.CreateObject("ADODB.Recordset")
'			SQL_N = "Select * from TB_Frequencia_Periodo WHERE CO_Matricula = "& cod_cons
'			Set RSF = CON_N.Execute(SQL_N)
'			
'			if RSF.eof THEN
'				f1="&nbsp;"
'				f2="&nbsp;"
'				f3="&nbsp;"
'				f4="&nbsp;"			
'			else	
'				f1=RSF("NU_Faltas_P1")
'				f2=RSF("NU_Faltas_P2")
'				f3=RSF("NU_Faltas_P3")
'				f4=RSF("NU_Faltas_P4")		
'			END IF				
'		
'				Set param_table3 = Pdf.CreateParam("width=533; height=20; rows=2; cols=10; border=0; cellborder=0; cellspacing=0;")
'				Set Legenda = Doc.CreateTable(param_table3)
'				Legenda.Font = Font
'				y_legenda=y_medias-altura_medias
'				'response.Write(altura_medias)
'				'response.end()
'				With Legenda.Rows(1)
'				   .Cells(1).Width = 40
'				   .Cells(2).Width = 20
'				   .Cells(3).Width = 40
'				   .Cells(4).Width = 20
'				   .Cells(5).Width = 40
'				   .Cells(6).Width = 20
'				   .Cells(7).Width = 40
'				   .Cells(8).Width = 20 
'				   .Cells(9).Width = 43 
'				   .Cells(10).Width = 250             
'				End With
'				data_exibe = data&" &agrave;s "& horario
'
'				Legenda(1, 1).Colspan= 8	
'				'Legenda(1, 10).RowSpan = 2					
'				Legenda(1, 1).AddText "<b>Freq&uuml;&ecirc;ncia (Faltas):</b>", "size=7;html=true", Font 
'				Legenda(2, 1).AddText "Bimestre 1:", "size=7;html=true;", Font 
'				Legenda(2, 2).AddText ""&f1&"", "size=7;html=true;", Font 
'				Legenda(2, 3).AddText "Bimestre 2:", "size=7;html=true;", Font 
'				Legenda(2, 4).AddText ""&f2&"", "size=7;html=true;", Font 				
'				Legenda(2, 5).AddText "Bimestre 3:", "size=7;html=true;", Font 
'				Legenda(2, 6).AddText ""&f3&"", "size=7;html=true;", Font 
'				Legenda(2, 7).AddText "Bimestre 4:", "size=7;html=true;", Font
'				Legenda(2, 8).AddText ""&f4&"", "size=7;html=true;", Font 				 								
'				Legenda(1, 10).AddText "<b><Div align=""right"">Documento impresso em: "&data_exibe&"</div></b>", "size=8; html=true", Font 				
'				Page.Canvas.DrawTable Legenda, "x="&margem&", y="&y_legenda&"" 
				
				
				'Set param_table4 = Pdf.CreateParam("width=160; height=15; rows=1; cols=1; border=0.5; cellborder=0.5; cellspacing=0;indenty=2;")
				'Set Impresso = Doc.CreateTable(param_table4)
				'Impresso.Font = Font
				'y_impresso=y_medias-altura_medias-15
				'
				'Impresso(1, 1).AddText "<b><Div align=""center"">Documento impresso em:</div></b>", "size=8;html=true", Font 
				'Page.Canvas.DrawTable Impresso, "x=405, y="&y_impresso&""
				'
				'
				'data_exibe = data&" &agrave;s "& horario
				'Set param_table5 = Pdf.CreateParam("width=160; height=15; rows=1; cols=1; border=0.5; cellborder=0; cellspacing=0;")
				'Set Data_Impresso = Doc.CreateTable(param_table5)
				'y_data_impresso=y_impresso-15
				'Data_Impresso(1, 1).AddText "<Div align=""center"">"&data_exibe&"</div>", "size=8;indenty=2;html=true", Font 
				'Page.Canvas.DrawTable Data_Impresso, "x=405, y="&y_data_impresso&"" 
				
				
'				
'					 SET Param_Relatorio = Pdf.CreateParam("x="&margem_x&";y="&margem_y&"; height=50; width=50; alignment=left; size=8; color=#000000")
'				'	Relatorio = "Sistema Web Diretor - SWD025"
'					Relatorio = "SWD025"
'					Do While Len(Relatorio) > 0
'						CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
'					 
'						If CharsPrinted = Len(Relatorio) Then Exit Do
'						   SET Page = Page.NextPage
'						Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
'					Loop 
'				
				'	width_simplynet=Page.Width - 30
				'	 SET Param_simplynet= Pdf.CreateParam("x=30;y=145; height=50; width="&width_simplynet&"; alignment=center; size=8; color=#000000;html=true")
				'	simplynet = "<Div align=""center"">Simply Net Informa&ccedil;&atilde;o e Tecnologia</div>"
				'
				'	Do While Len(simplynet) > 0
				'		CharsPrinted = Page.Canvas.DrawText(simplynet, Param_simplynet, Font )
				'	 
				'		If CharsPrinted = Len(simplynet) Then Exit Do
				'		   SET Page = Page.NextPage
				'		simplynet = Right( simplynet, Len(simplynet) - CharsPrinted)
				'	Loop 

'INÍCIO DO RECIBO==================================================================================
		
		'Assinatura do responsável
'			Page.Canvas.SetParams "LineWidth=0.5" 
'			With Page.Canvas
'			   .MoveTo 330, 30
'			   .LineTo Page.Width - 30, 30
'			   .Stroke
'			End With 
'		
'		'Data===========================================================================
'			 SET Param_data = Pdf.CreateParam("x=30;y=40; height=50; width=100; alignment=Left; size=8; color=#000000")
'			data_preenche = "Data:_____/_____/_____"
'			CharsPrinted = Page.Canvas.DrawText(data_preenche, Param_data, Font )
'		'===========================================================================
'		
'			Page.Canvas.SetParams "Dash1=2; DashPhase=1"
'			Page.Canvas.SetParams "LineWidth=0.7" 
'			With Page.Canvas
'			   .MoveTo 30, 125
'			   .LineTo Page.Width - 30, 125
'			   .Stroke
'			End With 
'			
'
'			width_tesoura=Page.Width - 30
'			 SET Param_Tesoura = Pdf.CreateParam("x=30;y=136; height=50; width="&width_tesoura&"; alignment=center; size=16; color=#000000, html=true")
'			Tesoura = "<div align=""center"">&quot;</div>"
'			
'			Do While Len(Tesoura) > 0
'				CharsPrinted = Page.Canvas.DrawText(Tesoura, Param_Tesoura, Font_Tesoura )
'			 
'				If CharsPrinted = Len(Tesoura) Then Exit Do
'				   SET Page = Page.NextPage
'				Tesoura = Right( Tesoura, Len(Tesoura) - CharsPrinted)
'			Loop 
'			
'			largura_logo_pqno=Logo.Width 'formatnumber(Logo.Width*0.8,0)
'			altura_logo_pqno=Logo.Height 'formatnumber(Logo.Height*0.8,0)
'			
'			Set Param_Logo_Pqno = Pdf.CreateParam
'			   Param_Logo_Pqno("x") = 39 
'			   Param_Logo_Pqno("y") = altura_logo_pqno+45
'			   Param_Logo_Pqno("ScaleX") = 0.8
'			   Param_Logo_Pqno("ScaleY") = 0.8
'			   Page.Canvas.DrawImage Logo, Param_Logo_Pqno
'			
'			x_texto_recibo=largura_logo_pqno+ 43
'			y_texto_recibo=formatnumber(altura_logo_pqno*3,0)
'			width_texto=Page.Width -largura_logo_pqno - 100
'			SET Param_recibo = Pdf.CreateParam("x="&x_texto_recibo&";y="&y_texto_recibo&"; height="&altura_logo_pqno&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
'			Text_recibo = "<p><center><i><b>Col&eacute;gio Maria Raythe</b></i></center></p>"
'			
'			Do While Len(Text_recibo) > 0
'				CharsPrinted = Page.Canvas.DrawText(Text_recibo, Param_recibo, Font )
'			 
'				If CharsPrinted = Len(Text_recibo) Then Exit Do
'					SET Page = Page.NextPage
'				Text_recibo = Right( Text_recibo, Len(Text_recibo) - CharsPrinted)
'			Loop 
'			
'			Set param_recibo2 = Pdf.CreateParam("width=100; height=40; rows=2; cols=2; border=0; cellborder=0; cellspacing=0;")
'			Set recibo_2 = Doc.CreateTable(param_recibo2)
'			y_recibo_2=y_texto_recibo
'			With recibo_2.Rows(1)
'			   .Cells(1).Width = 50
'			   .Cells(2).Width = 50
'			End With
'			recibo_2(1, 1).AddText "<Div align=""center"">Ano Letivo:</div>", "size=8;indenty=2;html=true", Font 
'			recibo_2(1, 2).AddText "<Div align=""Right""><b>"&ano_letivo&"</b></div>", "size=9;indenty=2;html=true", Font 
'			recibo_2(2, 1).AddText "<Div align=""center"">Matr&iacute;cula:</div>", "size=8;indenty=2;html=true", Font 
'			recibo_2(2, 2).AddText "<Div align=""Right""><b>"&cod_cons&"</b></div>", "size=8;indenty=2;html=true", Font 
'			Page.Canvas.DrawTable recibo_2, "x=455, y="&y_recibo_2&"" 
'		
'		
'			
'			SET Param_Nome_Aluno_Recibo = Pdf.CreateParam("x=30;y=75; height=50; width="&width_nome_aluno&"; alignment=left; html=True")
'			Nome_Recibo = "<font style=""font-size:10pt;""><b>"&nome_aluno&"</b></font>"
'			Do While Len(Nome_Recibo) > 0
'				CharsPrinted = Page.Canvas.DrawText(Nome_Recibo, Param_Nome_Aluno_Recibo, Font )
'			 
'				If CharsPrinted = Len(Nome_Recibo) Then Exit Do
'					SET Page = Page.NextPage
'				Nome_Recibo = Right( Nome_Recibo, Len(Nome_Recibo) - CharsPrinted)
'			Loop 
'		
'			SET Param_Dados_Aluno_Recibo = Pdf.CreateParam("x=30;y=60; height=50; width="&width_nome_aluno&"; alignment=left; html=True")
'			'Dados_Recibo = "<font style=""font-size:8pt;"">Curso: <b>"&no_curso &"</b> S&eacute;rie: <b>"&no_etapa &"</b> Turma: <b>"& turma &"</b></font>"
'			Dados_Recibo = "<font style=""font-size:8pt;"">Curso: <b>"&no_curso &"</b> Turma: <b>"& turma &"</b></font>"				
'			Do While Len(Dados_Recibo) > 0
'				CharsPrinted = Page.Canvas.DrawText(Dados_Recibo, Param_Dados_Aluno_Recibo, Font )
'			 
'				If CharsPrinted = Len(Dados_Recibo) Then Exit Do
'					SET Page = Page.NextPage
'				Dados_Recibo = Right( Dados_Recibo, Len(Dados_Recibo) - CharsPrinted)
'			Loop 
'			y_texto_recibo2=y_texto_recibo-20
'			SET Param_recibo3 = Pdf.CreateParam("x="&x_texto_recibo&";y="&y_texto_recibo2&"; height=60; width="&width_texto&"; alignment=center; size=12; color=#000000; html=true")
'			Text_recibo3 = "<p><center><b><font style=""font-size:7pt;"">Devolver essa parte assinada ao col&eacute;gio</font><b></center></p>"
'			
'			Do While Len(Text_recibo3) > 0
'				CharsPrinted = Page.Canvas.DrawText(Text_recibo3, Param_recibo3, Font )
'			 
'				If CharsPrinted = Len(Text_recibo3) Then Exit Do
'					SET Page = Page.NextPage
'				Text_recibo3 = Right( Text_recibo3, Len(Text_recibo3) - CharsPrinted)
'			Loop
'			
'			SET Param_recibo4 = Pdf.CreateParam("x=230;y=40; height=60; width=300; alignment=left; size=8; color=#000000; html=true")
'			Responsavel_Recibo = "<font style=""font-size:8pt;"">Assinatura do Respons&aacute;vel:</font>"
'			Do While Len(Responsavel_Recibo) > 0
'				CharsPrinted = Page.Canvas.DrawText(Responsavel_Recibo, Param_recibo4, Font )
'			 
'				If CharsPrinted = Len(Responsavel_Recibo) Then Exit Do
'					SET Page = Page.NextPage
'				Responsavel_Recibo = Right( Responsavel_Recibo, Len(Responsavel_Recibo) - CharsPrinted)
'			Loop
NEXT	
		End IF	
			

	
'	RS0.Close
'	Set RS0 = Nothing	
'		
'	RS.Close
'	Set RS = Nothing
'	
'	RS1.Close
'	Set RS1 = Nothing
'	
'	RS3.Close
'	Set RS3 = Nothing
'	
'	RS5.Close
'	Set RS5 = Nothing	
'			
'	RStabela.Close
'	Set RStabela = Nothing	

arquivo="boleto"
Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

