<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 60 'valor em segundos
'Pedido de Movimentação do Almoxarifado
arquivo="SWD010"
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


grupo = request.Form("grupo")




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

		Set CON9 = Server.CreateObject("ADODB.Connection") 
		ABRIR9 = "DBQ="& CAMINHO_ax & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON9.Open ABRIR9
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0	


if grupo<>"nulo" then
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT NO_Grupo FROM TB_Grupo where CO_Grupo="&grupo
	RS.Open SQL, CON9
	
	if RS.eof then
		response.Write("Erro em TB_Grupo")
		response.end()
	else
		no_grupo = RS("NO_Grupo")
		no_grupo = replace_latin_char(no_grupo,"html")	
	end if	
else
		no_grupo = "___________________________"
end if	

		no_etapa = "___________________________"				

			SET Page = Doc.Pages.Add( 595, 842 )
					
'CABEÇALHO==========================================================================================		
			Set Param_Logo_Gde = Pdf.CreateParam
			margem=30				
			largura_logo_gde=formatnumber(Logo.Width*0.3,0)
			altura_logo_gde=formatnumber(Logo.Height*0.3,0)
	
			Param_Logo_Gde("x") = margem
			Param_Logo_Gde("y") = Page.Height - altura_logo_gde -22
			Param_Logo_Gde("ScaleX") = 0.3
			Param_Logo_Gde("ScaleY") = 0.3
			Page.Canvas.DrawImage Logo, Param_Logo_Gde
	
			'x_texto=largura_logo_gde+ 30
			x_texto= margem
			y_texto=formatnumber(Page.Height - altura_logo_gde/2,0)-10
			width_texto=Page.Width - (margem*2)

		
			SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
			Text = "<p><center><i><b><font style=""font-size:18pt;"">Pedido de Movimenta&ccedil;&atilde;o do Almoxarifado</font></b></i></center></p>"
			

			Do While Len(Text) > 0
				CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
			 
				If CharsPrinted = Len(Text) Then Exit Do
					SET Page = Page.NextPage
				Text = Right( Text, Len(Text) - CharsPrinted)
			Loop 
			
'================================================================================================================			
			y_primeira_tabela=Page.Height - altura_logo_gde
			x_primeira_tabela = margem
			area_utilizavel=Page.Width - (margem*2)
			altura_primeiro_separador = Page.Height - altura_logo_gde-30
			
			
			Page.Canvas.SetParams "LineWidth=2" 
			Page.Canvas.SetParams "LineCap=0" 
			With Page.Canvas
			   .MoveTo margem, altura_primeiro_separador
			   .LineTo Page.Width - margem, altura_primeiro_separador
			   .Stroke
			End With 	
	

			Set param_table1 = Pdf.CreateParam("width=534; height=120; rows=6; cols=6; border=0; cellborder=0.1; cellspacing=0;")
			Set Table = Doc.CreateTable(param_table1)
			Table.Font = Font
			
			y_primeira_tabela=altura_primeiro_separador-10
			x_primeira_tabela=margem
			
			With Table.Rows(1)
'			   .Cells(1).Height = 20
			   .Cells(1).Width = 23			   		   		   
			   .Cells(2).Width = 155
			   .Cells(3).Width = 23	
			   .Cells(4).Width = 155	
			   .Cells(5).Width = 23			   		   		   
			   .Cells(6).Width = 155			   		   
			End With
			Table(1, 1).ColSpan = 6
			Table(2, 1).ColSpan = 6		
			Table(3, 1).ColSpan = 6	
			Table(4, 1).ColSpan = 6	
			Table(5, 1).ColSpan = 6	
			'Table(35, 1).ColSpan = 4					
			'Table(36, 1).ColSpan = 4					
			Table.At(1, 1).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False"
			Table.At(2, 1).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False"	
			Table.At(3, 1).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False"	
			Table.At(4, 1).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False"	
			Table.At(5, 1).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False"
			'Table.At(35, 1).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False"						
			'Table.At(36, 1).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False"																												
			Table(1, 1).AddText "<b>Nome do Solicitante:</b> _____________________________________________________________________________________________________", "size=8;indenty=2;html=true", Font 
			Table(2, 1).AddText "<b>Etapa:</b> "&no_etapa&"     <b>Turma:</b> _________     <b>Projeto:</b>  ______________________________________________________________","size=8;indenty=2;html=true", Font	
 			Table(3, 1).AddText "<b>Data do Pedido:</b> ________________________________     <b>Atendido e Conferido por:</b> __________________________________________________", "size=8;indenty=2;html=true", Font	
			Table(4, 1).AddText "<b>Lan&ccedil;ado no Web Diretor em:</b>__________________________________________", "size=8;indenty=2;html=true", Font

			Table(6, 1).AddText "<center><b>QTD</b></center>", "size=8;indenty=2;html=true", Font 
			Table(6, 2).AddText "<center><b>ITEM</b></center>", "size=8;indenty=2;html=true", Font 				
			Table(6, 3).AddText "<center><b>QTD</b></center>", "size=8;indenty=2;html=true", Font 	
			Table(6, 4).AddText "<center><b>ITEM</b></center>", "size=8;indenty=2;html=true", Font
			Table(6, 5).AddText "<center><b>QTD</b></center>", "size=8;indenty=2;html=true", Font 	
			Table(6, 6).AddText "<center><b>ITEM</b></center>", "size=8;indenty=2;html=true", Font
			Set RSI = Server.CreateObject("ADODB.Recordset")
			if grupo<>"nulo" then	
				SQLI = "SELECT NO_Item FROM TB_Item where CO_Grupo="&grupo
			else
				SQLI = "SELECT NO_Item FROM TB_Item"	
			end if	
			RSI.Open SQLI, CON9	
			Set Row = Table.Rows.Add(15)							
			linha = 7
			coluna = 0
			IF RSI.EOF THEN
			  For i = 1 to 20
				Set Row = Table.Rows.Add(15)				  
			  next
			else
				while not RSI.EOF 
					coluna = coluna+2			
					nome_item = RSI("NO_Item")
					nome_item = replace_latin_char(nome_item,"html")
						
					Table(linha, coluna).AddText nome_item, "size=8;indentx=2;indenty=2;html=true", Font 
					IF coluna MOD 6 = 0 then
						Set Row = Table.Rows.Add(15)							
						linha = linha+1
						coluna = 0
					end if			
				
				RSI.MOVENEXT
				WEND 	
			 END IF	 		
			'Table(36, 1).AddText "<b>Assinatura do Requerente:</b> ____________________________________________________________________________________________", "size=8;indenty=2;html=true", Font 										
			Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 	
		y_assinatura = margem*2
		 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&y_assinatura&"; height=50; width="&area_utilizavel&"; alignment=left; html=true; size=8; color=#000000")				
			
		assinatura = "<b>Assinatura do Requerente:</b> ____________________________________________________________________________________________"
		Do While Len(assinatura) > 0
			CharsPrinted = Page.Canvas.DrawText(assinatura, Param_Relatorio, Font )
		 
			If CharsPrinted = Len(assinatura) Then Exit Do
			   SET Page = Page.NextPage
			assinatura = Right( assinatura, Len(assinatura) - CharsPrinted)
		Loop 						

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

