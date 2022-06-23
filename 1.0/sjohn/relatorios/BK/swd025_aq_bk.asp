<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 30 'valor em segundos
'Boletim de Avaliações Qualitativas
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/funcoes7.asp"-->
<!--#include file="../../global/conta_alunos.asp"-->

<% 
response.Charset="ISO-8859-1"
opt= request.QueryString("opt")
ori= request.QueryString("ori")
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
nvg=session("nvg")
session("nvg")=nvg
arquivo="SWD025"

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

data = dia &"/"& meswrt &"/"& ano
horario = hora & ":"& minwrt	

	'Dim AspPdf, Doc, Page, Font, Text, Param, Image, CharsPrinted
	'Instancia o objeto na memória
	SET Pdf = Server.CreateObject("Persits.Pdf")
	SET Doc = Pdf.CreateDocument
	Set Logo = Doc.OpenImage( Server.MapPath( "../img/logo_pdf.gif") )
	Set Font = Doc.Fonts.LoadFromFile(Server.MapPath("../fonts/arial.ttf"))	
	If Font.Embedding = 2 Then
	   Response.Write "Embedding of this font is prohibited."
	   Set Font = Nothing
	End If
		 		 

	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR

	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1
		
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0	
	
	Set CON_N = Server.CreateObject("ADODB.Connection")
	ABRIR3 = "DBQ="& CAMINHO_nw & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_N.Open ABRIR3
	

if opt="ws" then

else	
	obr=request.QueryString("obr")
	dados_informados = split(obr, "$!$" )
	co_aluno = dados_informados(0)
	unidade = dados_informados(1)
	curso = dados_informados(2)
	co_etapa = dados_informados(3)
	turma = dados_informados(4)
end if


' 		Set RS0 = Server.CreateObject("ADODB.Recordset")
'		SQL_0 = "Select * from TB_Materia WHERE CO_Materia = '"& co_materia &"'"
'		Set RS0 = CON0.Execute(SQL_0)
'
'		no_materia= RS0("NO_Materia")
'		co_materia_pr= RS0("CO_Materia_Principal")
'		
'if Isnull(co_materia_pr) then
'	co_materia_pr= co_materia
'end if

dados_periodo =  periodos(periodo, "num")
total_periodo = split(dados_periodo,"#!#") 
notas_a_lancar = ubound(total_periodo)-2




nu_chamada_check = 1	

  vetor_matrics = co_aluno
  
alunos_encontrados = split(vetor_matrics,"#$#")   
For alne=0 to ubound(alunos_encontrados)	  
  	cod_cons=alunos_encontrados(alne)
	
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
  
  		
gera_pdf="sim"


	if gera_pdf="sim" then	
			
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Unidade WHERE NU_Unidade="& unidade
		RS2.Open SQL2, CON0
						
		no_unidade = RS2("NO_Sede")		
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

			co_municipio=municipio_unidade			
			municipio_unidade=RS3m("NO_Municipio")	
			
			if bairro_unidade="" or isnull(bairro_unidade) then
			else
			
				Set RS3m = Server.CreateObject("ADODB.Recordset")
				SQL3m = "SELECT * FROM TB_Bairros WHERE SG_UF='"& uf_unidade_municipio &"' AND CO_Municipio="&co_municipio&" AND CO_Bairro="&bairro_unidade
				RS3m.Open SQL3m, CON0
				
				bairro_unidade=RS3m("NO_Bairro")					
				bairro_unidade=" - "&bairro_unidade
			end if													
		end if
		endereco_unidade=rua_unidade&numero_unidade&complemento_unidade&bairro_unidade&cep_unidade&"<br>"&municipio_unidade&uf_unidade					
					
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Curso WHERE CO_Curso='"& curso &"'"
		RS3.Open SQL3, CON0
		
		no_curso= RS3("NO_Curso")
		no_abrv_curso = RS3("NO_Abreviado_Curso")
		co_concordancia_curso = RS3("CO_Conc")	
		
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Etapa WHERE CO_Etapa ='"& co_etapa &"'"
		RS4.Open SQL4, CON0
		
		no_etapa = RS4("NO_Etapa")
		
		

		no_curso= no_etapa&"&nbsp;"&co_concordancia_curso&"&nbsp;"&no_curso
		texto_turma = "&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;<b>Turma</b> "&turma
		texto_disciplina = "&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;<b>Disciplina:</b> "&no_materia
		'texto_periodo = "&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;<b>Per&iacute;odo:</b> "&no_periodo
		texto_periodo = ""
		mensagem_cabecalho=ano_letivo&"&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;"&no_curso&texto_turma&texto_disciplina&texto_periodo


		SET Page = Doc.Pages.Add(842, 595)
				
'CABEÇALHO==========================================================================================		
		Set Param_Logo_Gde = Pdf.CreateParam
		margem=25			
		area_utilizavel=Page.Width - (margem*2)
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
		Text = "<center><i><b><font style=""font-size:18pt;"">BOLETIM DE AVALIA&Ccedil;&Otilde;ES QUALITATIVAS</font></b></i></center>"
		
		
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

		'Set param_table1 = Pdf.CreateParam("width="&area_utilizavel&"; height=20; rows=1; cols=3; border=0; cellborder=0; cellspacing=0;")
'		Set Table = Doc.CreateTable(param_table1)
'		Table.Font = Font
		y_primeira_tabela=altura_segundo_separador-5
		x_primeira_tabela=margem+5
'		With Table.Rows(1)
'		   .Cells(1).Width = 50			   		   		   
'		   .Cells(2).Width = area_utilizavel-200
'		   .Cells(3).Width = 150	
'		End With
'		
'		Table(1, 1).AddText "<b>Ano Letivo:</b>", "size=8;html=true", Font 
'		Table(1, 2).AddText mensagem_cabecalho, "size=8;html=true", Font	
'		'Table(1, 3).AddText "<div align=""right""><b>Legenda:</b> Md=M&eacute;dia - Res=Resultado&nbsp;&nbsp;&nbsp;</div>", "size=8;html=true", Font	
'		Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 		
'
''================================================================================================================		
'			Page.Canvas.SetParams "LineWidth=2" 
'			Page.Canvas.SetParams "LineCap=0" 
'			With Page.Canvas
'			   .MoveTo margem, y_primeira_tabela-10
'			   .LineTo Page.Width - margem, y_primeira_tabela-10
'			   .Stroke
'			End With 	
			
			width_nome_aluno=Page.Width - margem
			
			SET Param_Nome_Aluno = Pdf.CreateParam("x="&margem&";y="&y_primeira_tabela&"; height=50; width="&width_nome_aluno&"; alignment=left; html=True")
			Nome = "<font style=""font-size:11pt;""><b>Alun"&desinencia&": "&nome_aluno&"</b></font>"
			

			Do While Len(Nome) > 0
				CharsPrinted = Page.Canvas.DrawText(Nome, Param_Nome_Aluno, Font )
			 
				If CharsPrinted = Len(Nome) Then Exit Do
					SET Page = Page.NextPage
				Nome = Right( Nome, Len(Nome) - CharsPrinted)
			Loop 			

			y_segunda_tabela=y_primeira_tabela-20	

			Set param_table1 = Pdf.CreateParam("width=533; height=25; rows=2; cols=8; border=0; cellborder=0; cellspacing=0;")
			Set Table = Doc.CreateTable(param_table1)
			Table.Font = Font
			y_table=Page.Height - altura_logo_gde-70
			
			With Table.Rows(1)
			   .Cells(1).Width = 40
			   .Cells(2).Width = 105
			   .Cells(3).Width = 25
			   .Cells(4).Width = 70
			   .Cells(5).Width = 60
			   .Cells(6).Width = 133
			   .Cells(7).Width = 50
			   .Cells(8).Width = 50      
			End With
			Table(1, 2).ColSpan = 5
			Table(1, 1).AddText "Unidade:", "size=9;", Font 
			Table(2, 1).AddText "Curso:", "size=9;", Font 
			Table(1, 2).AddText no_unidade, "size=9;", Font 
			Table(2, 2).ColSpan = 2
			Table(2, 2).AddText no_curso, "size=9;html=true", Font 
			'Table(2, 3).AddText no_etapa, "size=9;", Font 
			Table(2, 4).AddText "Turma: "&turma, "size=9;", Font 
			Table(2, 5).AddText "N&ordm;. Chamada: "&cham, "size=9; html=true", Font 
			Table(2, 6).AddText cham, "size=9;", Font 
			Table(1, 7).AddText "<div align=""right"">Matr&iacute;cula: </div>", "size=9; html=true", Font 
			Table(1, 8).AddText cod_cons, "size=9;alignment=right", Font 
			Table(2, 7).AddText "Ano Letivo: ", "size=9; alignment=right", Font 
			Table(2, 8).AddText ano_letivo, "size=9;alignment=right", Font 
			Page.Canvas.DrawTable Table, "x="&margem&", y="&y_segunda_tabela&"" 
		
			y_segundo_separador = y_segunda_tabela-35
			With Page.Canvas
			   .MoveTo margem, y_segundo_separador 
			   .LineTo Page.Width - margem, y_segundo_separador
			   .Stroke
			End With 

	
		colunas_de_notas=notas_a_lancar
		total_de_colunas=colunas_de_notas+2		
				
		altura_medias=30
		y_terceira_tabela=y_segundo_separador-10	
		Set param_table2 = Pdf.CreateParam("width="&area_utilizavel&"; height="&altura_medias&"; rows=2; cols="&total_de_colunas&"; border=1; cellborder=0.1; cellspacing=0; x="&margem&"; y="&y_terceira_tabela&"; MaxHeight=420")

		Set Notas_Tit = Doc.CreateTable(param_table2)
		Notas_Tit.Font = Font				
		largura_colunas=(area_utilizavel-220)/(colunas_de_notas+1)
		With Notas_Tit.Rows(1)
		   .Cells(1).Width = 220		             
			for d=2 to total_de_colunas
			 .Cells(d).Width = largura_colunas					
			next
		End With
		

		
		tabela_col=1
		linha=1
		fim_do_cabecalho=1	

		tabela_col=1
		Notas_Tit(linha, tabela_col).AddText "<div align=""center""><b>Disciplinas</b></div>", "size=8; indenty=2; alignment=center; html=true", Font	
		tabela_col=2						
		for e=0 to notas_a_lancar
			sigla_periodo =  periodos(total_periodo(e), "sigla")
		
			Notas_Tit(linha, tabela_col).AddText "<div align=""center""><b>"&sigla_periodo&"</b></div>", "size=8; indenty=2; alignment=center; html=true", Font				
			tabela_col=tabela_col+1
		next			
		Set param_materias = PDF.CreateParam	
		param_materias.Set "size=8;expand=true;html=true" 			
												
		
		conta_notas = 1 
		
		nu_chamada_ckq = 0
		
			Set RSprog = Server.CreateObject("ADODB.Recordset")
			SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' and IN_MAE= TRUE order by NU_Ordem_Boletim "
			RSprog.Open SQLprog, CON0
		
			check=2
			
			Set CON_N = Server.CreateObject("ADODB.Connection") 
			ABRIRn = "DBQ="& CAMINHO_nw & ";Driver={Microsoft Access Driver (*.mdb)}"
			CON_N.Open ABRIRn			
				
			while not RSprog.EOF		
		
		
                co_materia=RSprog("CO_Materia")
				mae=RSprog("IN_MAE")
				fil=RSprog("IN_FIL")
				in_co=RSprog("IN_CO")
				nu_peso=RSprog("NU_Peso")
				ordem=RSprog("NU_Ordem_Boletim")
			
				Set RS1a = Server.CreateObject("ADODB.Recordset")
				SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia='"&co_materia&"'"
				RS1a.Open SQL1a, CON0
					
				no_materia=RS1a("NO_Materia")
				co_materia_pr= RS1a("CO_Materia_Principal")
				
				if Isnull(co_materia_pr) then
					co_materia_pr= co_materia
				end if			

			no_materia = Server.HTMLencode(no_materia)					 
			
				linha=linha+1
				param_materias.Add "indentx=2"				
				Set Row = Notas_Tit.Rows.Add(15) ' row height						
				Notas_Tit(linha, 1).AddText no_materia, param_materias		
				coluna=1
				param_materias.Add "indentx=0"
				
				for c=0 to notas_a_lancar	
					qtd_filhas=0
					acumula_valor = 0 	
					
					if mae=true then
							
						Set RS1a = Server.CreateObject("ADODB.Recordset")
						SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia_Principal='"&co_materia_pr&"'"
						RS1a.Open SQL1a, CON0
							
						if RS1a.EOF then
							vetor_filhas = co_materia
						else
							while not RS1a.eof
								co_materia= RS1a("CO_Materia")
								if qtd_filhas=0 then
									vetor_filhas = co_materia
								else
									vetor_filhas = vetor_filhas&"#!#"&co_materia	
								end if
								qtd_filhas=qtd_filhas+1
							RS1a.MOVENEXT
							WEND				
						end if
					else
						vetor_filhas = co_materia						
					end if								
				

					
					filhas = split(vetor_filhas,"#!#")
					wrk_calcula_medias = "S"
					for f=0 to ubound(filhas)
						co_materia = filhas(f)
						
						if c=0 then
							wrk_bd_nota_per = "VA_Ava1"
						elseif c=1 then
							wrk_bd_nota_per = "VA_Ava2"
						elseif c=2 then 			
							wrk_bd_nota_per = "VA_Ava3"
						elseif c=3 then 			 	
							wrk_bd_nota_per = "VA_Ava4"
						end if									
						
						Set RSC = Server.CreateObject("ADODB.Recordset")
						SQLC = "Select * from TB_Nota_W WHERE CO_Matricula = "& cod_cons &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"'"
						Set RSC = CON_N.Execute(SQLC)	
							
						if RSC.EOF then 
							valor="&nbsp;"
							wrk_calcula_medias = "N"
						else																	
							valor=RSC(wrk_bd_nota_per)	
							
							if valor = "M" then
								valor = "MB"
							end if			   															
						end if	
						
						if qtd_filhas>0 then
							if valor = "I" then
								valor_convertido = 25
							elseif valor = "R" then
								valor_convertido = 50
							elseif valor = "B" then
								valor_convertido = 75
							else
								valor_convertido = 100
							end if		
							
							acumula_valor = acumula_valor+valor_convertido
						end if							
					next			
					coluna=coluna+1	
					if acumula_valor>0 and qtd_filhas>0 and wrk_calcula_medias = "S" then
						valor = acumula_valor/qtd_filhas
						
						if valor <=25 then
							valor="I"
						elseif valor <=50 then
							valor="R"
						elseif valor <=75 then
							valor="B"
						else
							valor="MB"
						end if	
					else
						if wrk_calcula_medias = "N" then
							valor="&nbsp;"
						end if	
					end if								
										
					Notas_Tit(linha, coluna).AddText "<div align=""center"">"&valor&"</DIV>", param_materias							
				next				
			RSprog.MOVENEXT
			WEND
		limite=0
		Do While True
		limite=limite+1
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
				
				Paginacao = "1"
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
				Set Page = Page.NextPage	
				param_table2.Add( "RowTo="&fim_do_cabecalho&"; RowFrom=1" ) ' Row 1 is header.
				param_table2("RowFrom1") = LastRow + 1 ' RowTo1 is omitted and presumed infinite
'NOVO CABEÇALHO==========================================================================================		
				Set Param_Logo_Gde = Pdf.CreateParam
				margem=25			
				area_utilizavel=Page.Width - (margem*2)
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
				Text = "<center><i><b><font style=""font-size:18pt;"">BOLETIM DE AVALIA&Ccedil;&Otilde;ES QUALITATIVAS</font></b></i></center>"
				
				
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
	
				Set param_table1 = Pdf.CreateParam("width="&area_utilizavel&"; height=20; rows=1; cols=3; border=0; cellborder=0; cellspacing=0;")
				Set Table = Doc.CreateTable(param_table1)
				Table.Font = Font
				y_primeira_tabela=altura_segundo_separador-10
				x_primeira_tabela=margem+5
				With Table.Rows(1)
				   .Cells(1).Width = 50			   		   		   
				   .Cells(2).Width = area_utilizavel-200
				   .Cells(3).Width = 150	
				End With
				
				Table(1, 1).AddText "<b>Ano Letivo:</b>", "size=8;html=true", Font 
				Table(1, 2).AddText mensagem_cabecalho, "size=8;html=true", Font	
				'Table(1, 3).AddText "<div align=""right""><b>Legenda:</b> Md=M&eacute;dia - Res=Resultado&nbsp;&nbsp;&nbsp;</div>", "size=8;html=true", Font	
				Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 			
'================================================================================================================				 
			end if
			if limite>100 then
			response.Write("ERRO!")
			response.end()
			end if 
		Loop
		
		SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
		
		Relatorio = arquivo&" - Sistema Web Diretor"
		Do While Len(Relatorio) > 0
			CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )			
			If CharsPrinted = Len(Relatorio) Then Exit Do
			SET Page = Page.NextPage
			Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
		Loop 
		
		Param_Relatorio.Add "alignment=right" 
		
		Paginacao = Paginacao+1
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
		Loop 				
	End IF					
'End IF							
next
	

Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

