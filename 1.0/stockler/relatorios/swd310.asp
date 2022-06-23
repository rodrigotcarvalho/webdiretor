<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/bd_alunos.asp"-->
<!--#include file="../inc/bd_contato.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes2.asp"-->


<%
ordenacao = request.QueryString("opt")
arquivo = "SWD310"
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

selected_M = ""
selected_N = ""
selected_UCET = ""
if ordenacao = "M" then
	selected = "Matr&iacute;cula"
	orderBy = "co_matric"	
elseif ordenacao = "N" then 
	selected = "Nome do Aluno"
	orderBy = "nome"	
elseif ordenacao = "UCET" then 
	selected = "Unidade, Curso, Etapa e Turma"
	orderBy = "unidade,curso,etapa,turma"
elseif ordenacao = "R" then 	
	selected = "Nome do Respons&aacute;vel"
	orderBy = "nome"		
elseif ordenacao = "D" then 	
	selected = "Data da Rematr&iacute;cula"
	orderBy = "nome"		
end if


ano_letivo = session("ano_letivo")
co_usr = session("co_user")
nivel=4




		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CONwf = Server.CreateObject("ADODB.Connection") 
		ABRIRwf = "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONwf.Open ABRIRwf





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

	
	SET Page = Doc.Pages.Add(595,842)
			
'CABEÇALHO==========================================================================================		
	Set Param_Logo_Gde = Pdf.CreateParam
	margem=25			
	area_utilizavel=Page.Width - (margem*2)
	
	largura_logo_gde=formatnumber(Logo.Width*0.7,0)
	altura_logo_gde=formatnumber(Logo.Height*0.7,0)

   Param_Logo_Gde("x") = margem
   Param_Logo_Gde("y") = Page.Height - altura_logo_gde - margem
   Param_Logo_Gde("ScaleX") = 0.7
   Param_Logo_Gde("ScaleY") = 0.7
   Page.Canvas.DrawImage Logo, Param_Logo_Gde

	x_texto=largura_logo_gde+ margem+10
	y_texto=formatnumber(Page.Height - margem,0)
	width_texto=Page.Width -largura_logo_gde - 80


	SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height=30; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
	Text = "<center><i><b><font style=""font-size:18pt;"">Alunos que acessaram a Matr&iacute;cula On-line</font></b></i></center>"
	
	
	Do While Len(Text) > 0
		CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
	 
		If CharsPrinted = Len(Text) Then Exit Do
			SET Page = Page.NextPage
		Text = Right( Text, Len(Text) - CharsPrinted)
	Loop 

	
	Page.Canvas.SetParams "LineWidth=1" 
	Page.Canvas.SetParams "LineCap=0" 
	inicio_primeiro_separador=largura_logo_gde+margem+10
	altura_primeiro_separador= Page.Height - margem - 25


	Page.Canvas.SetParams "LineWidth=2" 
	Page.Canvas.SetParams "LineCap=0" 
	altura_segundo_separador= Page.Height - altura_logo_gde-40
	With Page.Canvas
	   .MoveTo margem, altura_segundo_separador
	   .LineTo area_utilizavel+margem, altura_segundo_separador
	   .Stroke
	End With 	

	Set param_table1 = Pdf.CreateParam("width=547; height=40; rows=2; cols=3; border=0; cellborder=0; cellspacing=0;")
	Set Table = Doc.CreateTable(param_table1)
	Table.Font = Font
	y_primeira_tabela=altura_segundo_separador-10
	x_primeira_tabela=margem+5
	With Table.Rows(1)
	   .Cells(1).Width = 100
	   .Cells(2).Width = 347  
	   .Cells(3).Width = 100 		   		   		   
	End With
	
	
	Table(1, 2).AddText "<center><b>Ordenação</b></center>", "size=9;html=true", Font 
	Table(2, 2).AddText "<center>"&selected&"</center>", "size=9;html=true", Font 

	Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 	
	
	Page.Canvas.SetParams "LineWidth=2" 
	Page.Canvas.SetParams "LineCap=0" 
	altura_terceiro_separador= y_primeira_tabela-40
	With Page.Canvas
	   .MoveTo margem, altura_terceiro_separador
	   .LineTo area_utilizavel+margem, altura_terceiro_separador
	   .Stroke
	End With 			

'================================================================================================================			

	colunas_de_notas=3
	total_de_colunas=7				
	altura_medias=20
	y_segunda_tabela=altura_terceiro_separador-10	
	Set param_table2 = Pdf.CreateParam("width=547; height="&altura_medias&"; rows=1; cols=5; border=0; cellborder=0.5; cellspacing=0; x="&margem&"; y="&y_segunda_tabela&"; MaxHeight=650")

	Set Notas_Tit = Doc.CreateTable(param_table2)
	Notas_Tit.Font = Font				
	largura_colunas=(547-80-45-80-150)/colunas_de_notas		
	
	With Notas_Tit.Rows(1)
	   .Cells(1).Width = 57
	   .Cells(2).Width = 155	
	   .Cells(3).Width = 155
	   .Cells(4).Width = 80			             
	   .Cells(5).Width = 100
	End With
	Notas_Tit(1, 1).AddText "<div align=""center""><b>Matr&iacute;cula</b></div>", "size=9;indenty=2; html=true", Font 
	Notas_Tit(1, 2).AddText "<div align=""center""><b>Nome do Aluno</b></div>", "size=9;alignment=center; indenty=2;html=true", Font 
	Notas_Tit(1, 3).AddText "<div align=""center""><b>Nome do Respons&aacute;vel Financeiro</b></div>", "size=9;alignment=center; indenty=2;html=true", Font 
	Notas_Tit(1, 4).AddText "<div align=""center""><b>CPF</b></div>", "size=9;alignment=center; indenty=2;html=true", Font 
	Notas_Tit(1, 5).AddText "<div align=""center""><b>Data da Rematr&iacute;cula.</b></div>", "size=9;alignment=center; indenty=2;html=true", Font 

	Set param_materias = PDF.CreateParam	
	param_materias.Set "size=7;expand=false" 			

Set Rs_ordena = Server.CreateObject ( "ADODB.RecordSet" )
'Vamos adicionar 2 campos nesse recordset!
'O método Append recebe 3 parâmetros:
'Nome do campo, Tipo, Tamanho (opcional)
'O tipo pertence à um DataTypeEnum, e você pode conferir os tipos em
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/ado270/htm/mdcstdatatypeenum.asp
'200 -> VarChar (String), 7 -> Data, 139 -> Numeric
Rs_ordena.Fields.Append "co_matric", 139, 10
Rs_ordena.Fields.Append "nome", 200, 255
Rs_ordena.Fields.Append "responsavel", 200, 255
Rs_ordena.Fields.Append "cpf", 200, 255
Rs_ordena.Fields.Append "data", 7
Rs_ordena.Fields.Append "hora", 7
Rs_ordena.Fields.Append "unidade", 200, 255
Rs_ordena.Fields.Append "curso", 200, 255
Rs_ordena.Fields.Append "etapa", 200, 255
Rs_ordena.Fields.Append "turma", 200, 255

'Vamos abrir o Recordset!
Rs_ordena.Open					
					Set RS = Server.CreateObject("ADODB.Recordset")
					SQL = "SELECT * FROM TB_Aunos_Rematriculados"
					RS.Open SQL, CONwf, 3, 3
linha=1
	if RS.EOF then								
		linha=linha+1
		Set Row = Notas_Tit.Rows.Add(15) ' row height	
		Notas_Tit(linha, 1).ColSpan = 5		
		Notas_Tit(linha, 1).AddText "<div align=""center"">Nenhum aluno rematriculado</div>", param_materias			
	else				
		While Not RS.EOF 		
			matric = RS("CO_Matricula_Escola")
			data = RS("DA_Ult_Acesso") 
			hora = RS("HO_ult_Acesso") 			
			aluno = buscaAluno(matric)
		    vetorAluno = split(aluno,"#!#")
			nome = Server.HTMLEncode(vetorAluno(2))
			tipo_resp_fin = buscaTipoResponsavelFinanceiro(matric)
			
			vetorContato = buscaContato (matric, tipo_resp_fin)
			dadosContato = split(vetorContato, "#!#")
			contratante = Server.HTMLEncode(dadosContato(2))
			cpfContratante = dadosContato(4)	
			
			ucet = buscaUCET(matric,session("ano_letivo"))
			vetorUCET = split(ucet,"#!#")
			nu_unidade =  vetorUCET(0)
			co_curso = vetorUCET(1)
			co_etapa = vetorUCET(2)
			co_turma = vetorUCET(3)
			
			Rs_ordena.AddNew
			Rs_ordena.Fields("co_matric").Value = matric
			Rs_ordena.Fields("nome").Value = nome
			Rs_ordena.Fields("responsavel").Value = contratante	
			Rs_ordena.Fields("cpf").Value = cpfContratante						
			Rs_ordena.Fields("data").Value = data
			Rs_ordena.Fields("hora").Value = hora
			Rs_ordena.Fields("unidade").Value = nu_unidade
			Rs_ordena.Fields("curso").Value = co_curso
			Rs_ordena.Fields("etapa").Value = co_etapa
			Rs_ordena.Fields("turma").Value = co_turma					
			'RS.AddNew
			
		RS.movenext
		wend	
		
		Rs_ordena.Sort = orderBy	
		
		
		
		Dim RowCount 
		RowCount = 0
		check=2	
		While Not Rs_ordena.EOF		
			linha=linha+1
			Set Row = Notas_Tit.Rows.Add(15) ' row height			
			param_materias.Add "expand=true;html=true;indenty=2;" 											
			Notas_Tit(linha, 1).AddText "<div align=""center"">"&Rs_ordena.Fields("co_matric").Value&"</div>", param_materias			
			Notas_Tit(linha, 2).AddText "<div align=""center"">"&Rs_ordena.Fields("nome").Value&"</div>", param_materias			
			Notas_Tit(linha, 3).AddText "<div align=""center"">"&Rs_ordena.Fields("responsavel").Value&"</div>", param_materias
			param_materias.Add "expand=false" 	
			Notas_Tit(linha, 4).AddText "<div align=""center"">"&Rs_ordena.Fields("cpf").Value&"</div>", param_materias	
			Notas_Tit(linha, 5).AddText "<div align=""center"">"&Rs_ordena.Fields("data").Value&" "&Rs_ordena.Fields("hora").Value&"</div>", param_materias																																															
		Rs_ordena.MoveNext
		Wend
	end if		
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
'CABEÇALHO==========================================================================================		
	Set Param_Logo_Gde = Pdf.CreateParam
	margem=25			
	area_utilizavel=Page.Width - (margem*2)
	
	largura_logo_gde=formatnumber(Logo.Width*0.7,0)
	altura_logo_gde=formatnumber(Logo.Height*0.7,0)

   Param_Logo_Gde("x") = margem
   Param_Logo_Gde("y") = Page.Height - altura_logo_gde - margem
   Param_Logo_Gde("ScaleX") = 0.7
   Param_Logo_Gde("ScaleY") = 0.7
   Page.Canvas.DrawImage Logo, Param_Logo_Gde

	x_texto=largura_logo_gde+ margem+10
	y_texto=formatnumber(Page.Height - margem,0)
	width_texto=Page.Width -largura_logo_gde - 80


	SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height=30; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
	Text = "<center><i><b><font style=""font-size:18pt;"">Alunos que acessaram a Matr&iacute;cula On-line</font></b></i></center>"
	
	
	Do While Len(Text) > 0
		CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
	 
		If CharsPrinted = Len(Text) Then Exit Do
			SET Page = Page.NextPage
		Text = Right( Text, Len(Text) - CharsPrinted)
	Loop 

	
	Page.Canvas.SetParams "LineWidth=1" 
	Page.Canvas.SetParams "LineCap=0" 
	inicio_primeiro_separador=largura_logo_gde+margem+10
	altura_primeiro_separador= Page.Height - margem - 25


	Page.Canvas.SetParams "LineWidth=2" 
	Page.Canvas.SetParams "LineCap=0" 
	altura_segundo_separador= Page.Height - altura_logo_gde-40
	With Page.Canvas
	   .MoveTo margem, altura_segundo_separador
	   .LineTo area_utilizavel+margem, altura_segundo_separador
	   .Stroke
	End With 	

	Set param_table1 = Pdf.CreateParam("width=547; height=40; rows=2; cols=3; border=0; cellborder=0; cellspacing=0;")
	Set Table = Doc.CreateTable(param_table1)
	Table.Font = Font
	y_primeira_tabela=altura_segundo_separador-10
	x_primeira_tabela=margem+5
	With Table.Rows(1)
	   .Cells(1).Width = 100
	   .Cells(2).Width = 347  
	   .Cells(3).Width = 100 		   		   		   
	End With
	
	
	Table(1, 2).AddText "<center><b>Ordenação</b></center>", "size=9;html=true", Font 
	Table(2, 2).AddText "<center>"&selected&"</center>", "size=9;html=true", Font 

	Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 	
	
	Page.Canvas.SetParams "LineWidth=2" 
	Page.Canvas.SetParams "LineCap=0" 
	altura_terceiro_separador= y_primeira_tabela-40
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

