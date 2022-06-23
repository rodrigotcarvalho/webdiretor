<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 60 'valor em segundos
'RELATORIO DE AVALIAÇAO
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/parametros.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/funcoes7.asp"-->
<% 
arquivo="SWD016"
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

inf_por_periodo=4


dados_informados = split(dados, "$!$" )

co_materia = dados_informados(0)
unidade = dados_informados(1)
curso = dados_informados(2)
co_etapa = dados_informados(3)
turma = dados_informados(4)
periodo_form = dados_informados(5)
co_prof = dados_informados(7)
aluno_form = dados_informados(8)


'if ori="ebe" then
'origem="../ws/doc/ofc/ebe/"
'end if

if mes<10 then
mes="0"&mes
end if

data = dia &"/"& mes &"/"& ano

if min<10 then
min="0"&min
end if

horario = hora & ":"& min

data_exibe = data&" &agrave;s "& horario

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
		
		Set CON3 = Server.CreateObject("ADODB.Connection") 
		ABRIR3 = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3			
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0	

		Set CONt = Server.CreateObject("ADODB.Connection") 
		ABRIRt = "DBQ="& CAMINHO_t & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONt.Open ABRIRt
'if opt="a" then
			
	If Not IsArray(alunos_encontrados) Then alunos_encontrados = Array() End if	
	ReDim preserve alunos_encontrados(UBound(alunos_encontrados)+1)	
	alunos_encontrados(Ubound(alunos_encontrados)) = aluno_form
	
'elseif opt="t" then
'
'	if unidade="999990" or unidade="" or isnull(unidade) then
'		SQL_ALUNOS="NULO"
'	else	
'		SQL_ALUNOS= "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.NU_Unidade = "& unidade		
'		if curso="999990" or curso="" or isnull(curso) then
'			SQL_CURSO=""
'		else
'			SQL_CURSO=" AND TB_Matriculas.CO_Curso = '"& curso &"'"			
'		end if
'	
'		if co_etapa="999990" or co_etapa="" or isnull(co_etapa) then
'			SQL_ETAPA=""		
'		else
'			SQL_ETAPA=" AND TB_Matriculas.CO_Etapa = '"& co_etapa &"'"				
'		end if
'	
'		if turma="999990" or turma="" or isnull(turma) then
'			SQL_TURMA=""		
'		else
'			SQL_TURMA=" AND TB_Matriculas.CO_Turma = '"& turma &"' "			
'		end if
'	
'	SQL_ALUNOS= SQL_ALUNOS&SQL_CURSO&SQL_ETAPA&SQL_TURMA&" order by TB_Matriculas.NU_Unidade,TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma, TB_Alunos.NO_Aluno"
'	end if
'
'	if SQL_ALUNOS="NULO" then
'	else
'	
'	nu_chamada_check = 1
'		Set RSA = Server.CreateObject("ADODB.Recordset")
'		CONEXAOA = SQL_ALUNOS
''response.Write(CONEXAOA)
'		Set RSA = CON1.Execute(CONEXAOA)
'		vetor_matriculas="" 
'		While Not RSA.EOF
'			nu_matricula = RSA("CO_Matricula")
'			nu_chamada = RSA("NU_Chamada")
'			if nu_chamada_check = 1 and nu_chamada=nu_chamada_check then
'				vetor_matriculas=nu_matricula
'			elseif nu_chamada_check = 1 then
'				while nu_chamada_check < nu_chamada
'					nu_chamada_check=nu_chamada_check+1
'				wend 
'				vetor_matriculas=nu_matricula
'			else
'				vetor_matriculas=vetor_matriculas&"#!#"&nu_matricula
'			end if
'		nu_chamada_check=nu_chamada_check+1			
'		RSA.MoveNext
'		Wend 
'	
'	end if	
''RESPONSE.Write(vetor_matriculas)
''RESPONSE.END()
''	matriculas_encontradas = split(vetor_matriculas, "#!#" )	
'
'	alunos_encontrados = split(vetor_matriculas, "#!#" )	
'
'	RSA.Close
'	Set RSA = Nothing
'end if


		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Professor where CO_Professor="& co_prof
		RS1.Open SQL1, CON3
			
		if RS1.EOF then	
			sexo_prof = "M"						
			nome_prof = "nome em branco"
		else			
			sexo_prof = RS1("IN_Sexo")			
			nome_prof = RS1("NO_Professor")
		end if
		
		if sexo_prof = "M" then
			profoa = "Professor"
		else		
			profoa = "Professora"
		end if			
		nome_prof = replace_latin_char(nome_prof,"html")	

 		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL_0 = "Select * from TB_Materia WHERE CO_Materia = '"& co_materia &"'"
		Set RS0 = CON0.Execute(SQL_0)

mat_princ=RS0("CO_Materia_Principal")

if mat_princ="" or isnull(mat_princ) then
	mat_princ=co_materia
end if
	
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
	

	
		tb_nota=tabela_nota(ano_letivo,unidade,curso,co_etapa,turma,"tb",0)
		caminho_nota=tabela_nota(ano_letivo,unidade,curso,co_etapa,turma,"cam",0)
					
		no_unidade= GeraNomes("U",unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)
		no_curso=GeraNomes("C",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro) 	
		no_etapa=GeraNomes("E",curso,co_etapa,variavel3,variavel4,variavel5,CON0,outro) 	
		no_disc=GeraNomes("D",co_materia,variavel2,variavel3,variavel4,variavel5,CON0,outro) 			
		tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")
		tp_freq=tipo_divisao_ano(curso,co_etapa,"in_frequencia")
		
		if co_materia = "GER" then
			titulo = "RELAT&Oacute;RIO"
		else
			titulo = "RELATO DE GRUPO"		
		end if
		
		Set CON_N = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIRn		

				
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Curso WHERE CO_Curso='"& curso &"'"
		RS3.Open SQL3, CON0
		
		no_abrv_curso = RS3("NO_Abreviado_Curso")
		co_concordancia_curso = RS3("CO_Conc")	
		
		no_unidade = unidade&" - "&no_unidade
		no_curso= no_etapa&" "&co_concordancia_curso&" "&no_curso
		'no_etapa = no_etapa&" "&co_concordancia_curso&" "&no_abrv_curso			
		
		Set RST = Server.CreateObject("ADODB.Recordset")
		SQLT = "SELECT * FROM TB_Turma WHERE NU_Unidade = "&unidade&" AND CO_Curso='"& curso &"' AND CO_Etapa='"& co_etapa &"' AND CO_Turma='"& turma &"'"
		RST.Open SQLT, CON0	
		
		no_auxiliares = RST("NO_Auxiliares")
		
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Periodo where TP_Modelo='"&tp_modelo&"' AND NU_Periodo ="& periodo_form 
		RS5.Open SQL5, CON0
		
		no_periodo = RS5("NO_Periodo")				

		SET Page = Doc.Pages.Add( 595, 842 )
		Paginacao = 1				
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
		y_texto=formatnumber(Page.Height - altura_logo_gde/2,0)
		width_texto=Page.Width - (margem*2)

	
		SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
		Text = "<p><center><i><b><font style=""font-size:18pt;"">"&titulo&"</font></b></i></center></p>"
		

		Do While Len(Text) > 0
			CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
		 
			If CharsPrinted = Len(Text) Then Exit Do
				SET Page = Page.NextPage
			Text = Right( Text, Len(Text) - CharsPrinted)
		Loop 
		
	
		y_nome_aluno=Page.Height - altura_logo_gde-46
		width_nome_aluno=Page.Width - margem
		
		SET Param_Nome_Aluno = Pdf.CreateParam("x="&margem&";y="&y_nome_aluno&"; height=50; width="&width_nome_aluno&"; alignment=left; html=True")
		Nome = "<font style=""font-size:11pt;""><b>Alun"&desinencia&": "&nome_aluno&"</b></font>"
		

		Do While Len(Nome) > 0
			CharsPrinted = Page.Canvas.DrawText(Nome, Param_Nome_Aluno, Font )
		 
			If CharsPrinted = Len(Nome) Then Exit Do
				SET Page = Page.NextPage
			Nome = Right( Nome, Len(Nome) - CharsPrinted)
		Loop 
		
		Page.Canvas.SetParams "LineWidth=2" 
		Page.Canvas.SetParams "LineCap=0" 
		With Page.Canvas
		   .MoveTo margem, Page.Height - altura_logo_gde-65
		   .LineTo Page.Width - margem, Page.Height - altura_logo_gde-65
		   .Stroke
		End With 	


		Set param_table1 = Pdf.CreateParam("width=533; height=25; rows=2; cols=8; border=0; cellborder=0; cellspacing=0;")
		Set Table = Doc.CreateTable(param_table1)
		Table.Font = Font
		y_table=Page.Height - altura_logo_gde-70
		
		With Table.Rows(1)
		   .Cells(1).Width = 40
		   .Cells(2).Width = 165
		   .Cells(3).Width = 25
		   .Cells(4).Width = 70
		   .Cells(5).Width = 40
		   .Cells(6).Width = 93
		   .Cells(7).Width = 50
		   .Cells(8).Width = 50      
		End With
		Table(1, 4).ColSpan = 3
		Table(1, 1).AddText "Unidade:", "size=9;", Font 
		Table(2, 1).AddText "Curso:", "size=9;", Font 
		Table(1, 2).AddText no_unidade, "size=9;", Font 
		Table(1, 4).AddText "Disciplina: "&no_disc, "size=9;", Font 		
		Table(2, 2).ColSpan = 2
		Table(2, 2).AddText no_curso, "size=9;", Font 
		'Table(2, 3).AddText no_etapa, "size=9;", Font 
		Table(2, 4).AddText "Turma: "&turma, "size=9;", Font 
'		Table(2, 5).AddText "N&ordm;. Chamada: "&cham, "size=9; html=true", Font 
'		Table(2, 6).AddText cham, "size=9;", Font 
		Table(2, 5).AddText "Per&iacute;odo: ", "size=9; html=true", Font 
		Table(2, 6).AddText no_periodo, "size=9;", Font 
		Table(1, 7).AddText "<div align=""right"">Matr&iacute;cula: </div>", "size=9; html=true", Font 
		Table(1, 8).AddText cod_cons, "size=9;alignment=right", Font 
		Table(2, 7).AddText "Ano Letivo: ", "size=9; alignment=right", Font 
		Table(2, 8).AddText ano_letivo, "size=9;alignment=right", Font 
		Page.Canvas.DrawTable Table, "x="&margem&", y="&y_table&"" 
	
		y_separador = Page.Height - altura_logo_gde-100
		With Page.Canvas
		   .MoveTo margem, y_separador
		   .LineTo Page.Width - margem, y_separador
		   .Stroke
		End With 
		y_nome_prof = y_separador-10
		SET Param_Nome_Prof = Pdf.CreateParam("x="&margem&";y="&y_nome_prof&"; height=50; width="&width_nome_aluno&"; alignment=left; html=True")
		Nome_Professor = "<font style=""font-size:9pt;""><b>"&profoa&": </b>"&nome_prof&"</font>"

		if co_materia="GER" then
			Nome_Professor = Nome_Professor&"<font style=""font-size:9pt;"">&nbsp;&nbsp;-&nbsp;&nbsp;<b>Auxiliar(es): </b>"&no_auxiliares&"</font>"
		end if	

		Do While Len(Nome_Professor) > 0
			CharsPrinted = Page.Canvas.DrawText(Nome_Professor, Param_Nome_Prof, Font )
		 
			If CharsPrinted = Len(Nome_Professor) Then Exit Do
			SET Page = Page.NextPage
			Nome_Professor = Right( Nome_Professor, Len(Nome_Professor) - CharsPrinted)
		Loop 		


'================================================================================================================							
		Set RSF = Server.CreateObject("ADODB.Recordset")
		SQL_N = "Select * from "& tb_nota &" WHERE CO_Matricula = "& cod_cons & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo_form		
		Set RSF = CON_N.Execute(SQL_N)
		
		if RSF.eof THEN
			texto=""
		else
			texto=RSF("TX_Avalia")
			if texto="&nbsp;" then
				texto=""			
			end if
		END IF		
		
		y_texto = y_nome_prof-30
		
		SET Param_Texto = Pdf.CreateParam("x="&margem&";y="&y_texto&"; height=600; width=533; alignment=left")			
		
		Do While Len(texto) > 0
			CharsPrinted = Page.Canvas.DrawText(texto, Param_Texto, Font )
		 
			If CharsPrinted = Len(texto) Then Exit Do
				texto = Right( texto, Len(texto) - CharsPrinted)				
				SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width=533; alignment=left; size=8; color=#000000")
		
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
				
				data_hora = "<center>Impresso em "&data &" &agrave;s "&horario&"</center>"
				Do While Len(Relatorio) > 0
					CharsPrinted = Page.Canvas.DrawText(data_hora, Param_Relatorio, Font )			
					If CharsPrinted = Len(data_hora) Then Exit Do
					SET Page = Page.NextPage
					data_hora = Right( data_hora, Len(data_hora) - CharsPrinted)
				Loop 	
				
				SET Page = Page.NextPage
				Paginacao = Paginacao+1						
			
				'NOVO CABEÇALHO====================================================================================		
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
						y_texto=formatnumber(Page.Height - altura_logo_gde/2,0)
						width_texto=Page.Width - (margem*2)
				
					
						SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
						Text = "<p><center><i><b><font style=""font-size:18pt;"">RELAT&Oacute;RIO</font></b></i></center></p>"
						
				
						Do While Len(Text) > 0
							CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
						 
							If CharsPrinted = Len(Text) Then Exit Do
								SET Page = Page.NextPage
							Text = Right( Text, Len(Text) - CharsPrinted)
						Loop 
						
					
						y_nome_aluno=Page.Height - altura_logo_gde-46
						width_nome_aluno=Page.Width - margem
						
						SET Param_Nome_Aluno = Pdf.CreateParam("x="&margem&";y="&y_nome_aluno&"; height=50; width="&width_nome_aluno&"; alignment=left; html=True")
						Nome = "<font style=""font-size:11pt;""><b>Alun"&desinencia&": "&nome_aluno&"</b></font>"
						
				
						Do While Len(Nome) > 0
							CharsPrinted = Page.Canvas.DrawText(Nome, Param_Nome_Aluno, Font )
						 
							If CharsPrinted = Len(Nome) Then Exit Do
								SET Page = Page.NextPage
							Nome = Right( Nome, Len(Nome) - CharsPrinted)
						Loop 
						
						Page.Canvas.SetParams "LineWidth=2" 
						Page.Canvas.SetParams "LineCap=0" 
						With Page.Canvas
						   .MoveTo margem, Page.Height - altura_logo_gde-65
						   .LineTo Page.Width - margem, Page.Height - altura_logo_gde-65
						   .Stroke
						End With 	
				
				
						Set param_table1 = Pdf.CreateParam("width=533; height=25; rows=2; cols=8; border=0; cellborder=0; cellspacing=0;")
						Set Table = Doc.CreateTable(param_table1)
						Table.Font = Font
						y_table=Page.Height - altura_logo_gde-70
						
						With Table.Rows(1)
						   .Cells(1).Width = 40
						   .Cells(2).Width = 165
						   .Cells(3).Width = 25
						   .Cells(4).Width = 70
						   .Cells(5).Width = 60
						   .Cells(6).Width = 73
						   .Cells(7).Width = 50
						   .Cells(8).Width = 50      
						End With
						Table(1, 4).ColSpan = 3
						Table(1, 1).AddText "Unidade:", "size=9;", Font 
						Table(2, 1).AddText "Curso:", "size=9;", Font 
						Table(1, 2).AddText no_unidade, "size=9;", Font 
						Table(1, 4).AddText "Disciplina: "&no_disc, "size=9;", Font 	
						Table(2, 2).ColSpan = 2
						Table(2, 2).AddText no_curso, "size=9;", Font 
						'Table(2, 3).AddText no_etapa, "size=9;", Font 
						Table(2, 4).AddText "Turma: "&turma, "size=9;", Font 
				'		Table(2, 5).AddText "N&ordm;. Chamada: "&cham, "size=9; html=true", Font 
				'		Table(2, 6).AddText cham, "size=9;", Font 
						Table(1, 7).AddText "<div align=""right"">Matr&iacute;cula: </div>", "size=9; html=true", Font 
						Table(1, 8).AddText cod_cons, "size=9;alignment=right", Font 
						Table(2, 7).AddText "Ano Letivo: ", "size=9; alignment=right", Font 
						Table(2, 8).AddText ano_letivo, "size=9;alignment=right", Font 
						Page.Canvas.DrawTable Table, "x="&margem&", y="&y_table&"" 
					
						y_separador = Page.Height - altura_logo_gde-100
						With Page.Canvas
						   .MoveTo margem, y_separador
						   .LineTo Page.Width - margem, y_separador
						   .Stroke
						End With 
						y_nome_prof = y_separador-10
						SET Param_Nome_Prof = Pdf.CreateParam("x="&margem&";y="&y_nome_prof&"; height=50; width="&width_nome_aluno&"; alignment=left; html=True")
						Nome_Professor = "<font style=""font-size:9pt;""><b>"&profoa&": </b>"&nome_prof&"</font>"
				
						if co_materia="INT" then
							Nome_Professor = Nome_Professor&"<font style=""font-size:9pt;"">&nbsp;&nbsp;-&nbsp;&nbsp;<b>Auxiliar(es): </b>"&no_auxiliares&"</font>"
						end if	
				
						Do While Len(Nome_Professor) > 0
							CharsPrinted = Page.Canvas.DrawText(Nome_Professor, Param_Nome_Prof, Font )
						 
							If CharsPrinted = Len(Nome_Professor) Then Exit Do
								SET Page = Page.NextPage
							Nome_Professor = Right( Nome_Professor, Len(Nome_Professor) - CharsPrinted)
						Loop 		
				
				
				'================================================================================================================							
				

		Loop 				

'		Set param_table3 = Pdf.CreateParam("width=250; height=20; rows=2; cols=10; border=0; cellborder=0; cellspacing=0;")
'		Set Legenda = Doc.CreateTable(param_table3)
'		Legenda.Font = Font
'		y_legenda=y_medias-altura_medias
'		'response.Write(altura_medias)
'		'response.end()
'		With Legenda.Rows(1)
'		   .Cells(1).Width = 40
'		   .Cells(2).Width = 20
'		   .Cells(3).Width = 40
'		   .Cells(4).Width = 20
'		   .Cells(5).Width = 40
'		   .Cells(6).Width = 20
'		   .Cells(7).Width = 20
'		   .Cells(8).Width = 20 
'		   .Cells(9).Width = 63 
'		   .Cells(10).Width = 250             
'		End With
'
'		Legenda(1, 1).Colspan= 8
'		Legenda(1, 10).AddText "<b><Div align=""right"">Documento impresso em: "&data_exibe&"</div></b>", "size=8; html=true", Font 							
'		Page.Canvas.DrawTable Legenda, "x="&margem&", y="&margem&"" 
'			
		SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width=533; alignment=left; size=8; color=#000000")
		
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
		
		data_hora = "<center>Impresso em "&data &" &agrave;s "&horario&"</center>"
		Do While Len(Relatorio) > 0
			CharsPrinted = Page.Canvas.DrawText(data_hora, Param_Relatorio, Font )			
			If CharsPrinted = Len(data_hora) Then Exit Do
			SET Page = Page.NextPage
			data_hora = Right( data_hora, Len(data_hora) - CharsPrinted)
		Loop 	

'		End IF	
'	END IF 		
Next						

	
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


Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

