<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 60 'valor em segundos
'Conteúdo da Entrevista
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

inf_por_periodo=4

obr=request.QueryString("obr")
dados_informados = split(obr, "?" )


	cod= dados_informados(0)
	ordem= dados_informados(1)
	status_entrevista= dados_informados(2)
	data_de= dados_informados(3)
	hora_de= dados_informados(4)
	data_inicio= dados_informados(5)
	data_ate= dados_informados(6)
	hora_ate= dados_informados(7)
	data_fim= dados_informados(8)

	dados_dtd= split(data_de, "/" )
	dia_de= dados_dtd(0)
	mes_de= dados_dtd(1)
	ano_de= dados_dtd(2)

	
	dados_dta= split(data_ate, "/" )
	dia_ate= dados_dta(0)
	mes_ate= dados_dta(1)
	ano_ate= dados_dta(2)
	
	if mes<10 then
	mes="0"&mes
	end if
	
	data = dia &"/"& mes &"/"& ano
	
	if min<10 then
	min="0"&min
	end if

	horario = hora & ":"& min
	
	ho_entrevista=hora&":"&min
	
'	data_split= Split(da_entrevista,"/")
'	dia=data_split(0)
'	mes=data_split(1)
'	ano=data_split(2)
'	
'	da_entrevista_cons=mes&"/"&dia&"/"&ano	
'	ho_entrevista_cons=hora&":"&min
	
	if status_entrevista="" or isnull(status_entrevista) then
		status_entrevista_form = "Todos"
	else
		entrevistas = split(status_entrevista,",")
		for s = 0 to ubound(entrevistas)
			Select case entrevistas(s)		
				case 1
				nome_status="Atendida"
				
				case 2
				nome_status="Cancelada"
				
				case 3
				nome_status="Pendente"	
			end select	
			if s = 0 then
				status_entrevista_form = nome_status	
			else
				status_entrevista_form = status_entrevista_form&", "&nome_status
			end if		
		next		
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

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2
		Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4 = "DBQ="& CAMINHO_e & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0	



max_notas_exibe=1*periodo_form
max_notas_exibe=max_notas_exibe
if cod=0 then
	desinencia="o"
	nome_aluno = "Todos"
	no_unidade= "Todas"
	no_curso="Todos"	
	no_etapa="Todas" 
	turma ="Todas"		
	
else	
	Set RSA = Server.CreateObject("ADODB.Recordset")
	CONEXAOA = "Select * from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.CO_Matricula = "& cod
	Set RSA = CON1.Execute(CONEXAOA)
	
		cod_cons = RSA("CO_Matricula")
		no_aluno= RSA("NO_Aluno")			
		cham = RSA("NU_Chamada")
		situac=	RSA("CO_Situacao")
		unidade = RSA("NU_Unidade")
		curso = RSA("CO_Curso")
		etapa = RSA("CO_Etapa")
		turma = RSA("CO_Turma")
		sexo_aluno = RSA("IN_Sexo")	
		ano_aluno = RSA("NU_Ano")
		rematricula = RSA("DA_Rematricula")
		situacao = RSA("CO_Situacao")
		encerramento= RSA("DA_Encerramento")
		
		if situac<>"C" then
			no_aluno=no_aluno&" - Aluno Inativo"
		end if			
		nome_aluno = replace_latin_char(no_aluno,"html")
	
		
		
		if sexo_aluno="F" then
			desinencia="a"
		else
			desinencia="o"
		end if
		
		
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod
		RS1.Open SQL1, CON1
		
		if RS1.EOF then
			'response.redirect("index.asp?nvg="&nvg&"&opt=err1")
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
			no_unidade= GeraNomes("U",unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)
			no_curso=GeraNomes("C",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro) 	
			no_etapa=GeraNomes("E",curso,etapa,variavel3,variavel4,variavel5,CON0,outro) 			
					
			Set RS3 = Server.CreateObject("ADODB.Recordset")
			SQL3 = "SELECT * FROM TB_Curso WHERE CO_Curso='"& curso &"'"
			RS3.Open SQL3, CON0
			
			no_abrv_curso = RS3("NO_Abreviado_Curso")
			co_concordancia_curso = RS3("CO_Conc")	
			
			no_unidade = unidade&" - "&no_unidade
			no_curso= no_etapa&" "&co_concordancia_curso&" "&no_curso
		end if	
end if		
		'no_etapa = no_etapa&" "&co_concordancia_curso&" "&no_abrv_curso				

		SET Page = Doc.Pages.Add( 842, 595)
				
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
		Text = "<p><center><i><b><font style=""font-size:18pt;"">Entrevistas</font></b></i></center></p>"
		

		Do While Len(Text) > 0
			CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
		 
			If CharsPrinted = Len(Text) Then Exit Do
				SET Page = Page.NextPage
			Text = Right( Text, Len(Text) - CharsPrinted)
		Loop 
		
'================================================================================================================			
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
		y_separador = Page.Height - altura_logo_gde-65
		With Page.Canvas
		   .MoveTo margem, y_separador
		   .LineTo Page.Width - margem, y_separador
		   .Stroke
		End With 	


		Set param_table1 = Pdf.CreateParam("width=780; height=25; rows=2; cols=8; border=0; cellborder=0; cellspacing=0;")
		Set Table = Doc.CreateTable(param_table1)
		Table.Font = Font
		y_table=Page.Height - altura_logo_gde-70
		
		With Table.Rows(1)
		   .Cells(1).Width = 40
		   .Cells(2).Width = 455
		   .Cells(3).Width = 25
		   .Cells(4).Width = 70
		   .Cells(5).Width = 60
		   .Cells(6).Width = 33
		   .Cells(7).Width = 50
		   .Cells(8).Width = 50      
		End With
		Table(1, 2).ColSpan = 5
		Table(1, 1).AddText "Unidade:", "size=9;", Font 
		Table(2, 1).AddText "Curso:", "size=9;", Font 
		Table(1, 2).AddText no_unidade, "size=9;", Font 
		Table(2, 2).ColSpan = 2
		Table(2, 2).AddText no_curso, "size=9;", Font 
		'Table(2, 3).AddText no_etapa, "size=9;", Font 
		Table(2, 4).AddText "Turma: "&turma, "size=9;", Font 
		Table(2, 5).AddText "N&ordm;. Chamada: "&cham, "size=9; html=true", Font 
		Table(2, 6).AddText cham, "size=9;", Font 
		Table(1, 7).AddText "<div align=""right"">Matr&iacute;cula: </div>", "size=9; html=true", Font 
		Table(1, 8).AddText cod_cons, "size=9;alignment=right", Font 
		Table(2, 7).AddText "Ano Letivo: ", "size=9; alignment=right", Font 
		Table(2, 8).AddText ano_letivo, "size=9;alignment=right", Font 
		Page.Canvas.DrawTable Table, "x="&margem&", y="&y_table&"" 
	
		
		With Page.Canvas
		   .MoveTo margem, Page.Height - altura_logo_gde-100
		   .LineTo Page.Width - margem, Page.Height - altura_logo_gde-100
		   .Stroke
		End With 
		
cod=cod*1
if cod=0 then
	sql_matricula = ""
else
	sql_matricula = "CO_Matricula ="&cod&" AND "
end if


if status_entrevista="" or isnull(status_entrevista) then
	sql_status_entrevista = ""
else
	sql_status_entrevista = "ST_Entrevista IN("&status_entrevista&") AND "	
end if
		
		y_table=y_table-40		

		Set param_table1 = Pdf.CreateParam("width=780; height=25; rows=1; cols=7; border=0; cellborder=0.5; cellspacing=0 ;x="&margem&"; y="&y_table&"; MaxHeight=450")
		Set Table = Doc.CreateTable(param_table1)
		Table.Font = Font

		
		Set param_materias = PDF.CreateParam	
		param_materias.Set "html = true; expand=true" 		
		
		With Table.Rows(1)
		   .Cells(1).Width = 80
		   .Cells(2).Width = 40
		   .Cells(3).Width = 180
		   .Cells(4).Width = 70
		   .Cells(5).Width = 180   
		   .Cells(6).Width = 180
		   .Cells(7).Width = 50 			     		   
		End With
		'Table(1, 2).ColSpan = 5
		Table(1, 1).AddText "Data / Hora", "size=9; alignment=center", Font 
		Table(1, 2).AddText "<div align=""center"">Matr&iacute;cula</div>", "size=9; alignment=center; html = true", Font 
		Table(1, 3).AddText "Nome do Aluno", "size=9; alignment=center", Font 
		Table(1, 4).AddText "Tipo", "size=9; alignment=center", Font 				
		Table(1, 5).AddText "Participantes", "size=9; alignment=center", Font 
		Table(1, 6).AddText "Atendido por", "size=9; alignment=center", Font 		
		Table(1, 7).AddText "Status", "size=9; alignment=center", Font 	
		
	Set RSe = Server.CreateObject("ADODB.Recordset")
	SQLe = "SELECT * FROM TB_Entrevistas WHERE "&sql_matricula&sql_status_entrevista&"(DA_Entrevista BETWEEN #"&data_de&"# AND #"&data_ate&"#) order BY DA_Entrevista, HO_Entrevista"
	RSe.Open SQLe, CON4, 3, 3



linha = 1
WHILE not RSe.EOF
	if check mod 2 =0 then
		cor = "tb_fundo_linha_par" 
	else 
		cor ="tb_fundo_linha_impar"
	end if 
  
	co_matric=RSe("CO_Matricula")
	da_entrevista=RSe("DA_Entrevista")
	ho_entrevista=RSe("HO_Entrevista")
	tp_entrevista=RSe("TP_Entrevista")
	no_participantes=RSe("NO_Participantes")
	st_entrevista=RSe("ST_Entrevista")
	co_agendado_com=RSe("CO_Agendado_com")
	tx_observaa=RSe("TX_Observa")
	co_usu_entrevista=RSe("CO_Usuario")
	
	if tp_entrevista="" or isnull(tp_entrevista) then
		tipo_entrevista=""
	else
	
		Set RST = Server.CreateObject("ADODB.Recordset")
		SQLT = "SELECT * FROM TB_Tipo_Entrevista Where tp_entrevista="&tp_entrevista
		RST.Open SQLT, CON0
	
		IF RST.EOF then
			tipo_entrevista=""	
		else
			tipo_entrevista=RST("TX_Descricao")
		end if	
	end if
				
	if co_agendado_com="" or isnull(co_agendado_com) then
		no_atendido=""
	else
		Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& co_agendado_com
		RSu.Open SQLu, CON
	
		IF RSu.EOF then
			no_atendido =""	
		else
			no_atendido =RSu("NO_Usuario")
		end if
			
	end if
	
	'if co_usu_entrevista="" or isnull(co_usu_entrevista) then
	'	no_atendido=""
	'else
	'		Set RSu = Server.CreateObject("ADODB.Recordset")
	'		SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& co_usu_entrevista
	'		RSu.Open SQLu, CON
	'
	'	IF RSu.EOF then
	'		no_agendado=""	
	'	else
	'		no_agendado=RSu("NO_Usuario")
	'	end if
	'		
	'end if
	
	hora_split= Split(ho_entrevista,":")
	hora=hora_split(0)
	min=hora_split(1)
	
	ho_entrevista = hora&":"&min
	hora_entrevista = hora&":"&min

	
	
	data_split= Split(da_entrevista,"/")
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
	da_show=dia&"/"&mes&"/"&ano&", "&hora&":"&min
	data_entrevista = dia&"/"&mes&"/"&ano
	
		Select case st_entrevista		
			case 1
			nome_status="Atendida"
			
			case 2
			nome_status="Cancelada"
			
			case 3
			nome_status="Pendente"	
		end select	
		
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& co_matric
	RS.Open SQL, CON1

	nome_aluno = RS("NO_Aluno")	
	Set Row = Table.Rows.Add(13) ' row height	
	linha = linha+1	
	param_materias.Add "size=8;expand=true" 												
	Table(linha, 1).AddText "<div align=""center"">"&da_show&"</div>", param_materias
	Table(linha, 2).AddText "<div align=""center"">"&co_matric&"</div>", param_materias	
	Table(linha, 3).AddText "<div align=""center"">"&nome_aluno&"</div>", param_materias
	Table(linha, 4).AddText "<div align=""center"">"&tipo_entrevista&"</div>", param_materias	
	Table(linha, 5).AddText "<div align=""center"">"&no_participantes&"</div>", param_materias																						
	Table(linha, 6).AddText "<div align=""center"">"&no_atendido&"</div>", param_materias		
	Table(linha, 7).AddText "<div align=""center"">"&nome_status&"</div>", param_materias	 
RSe.Movenext
'end if
WEND
			Do While True
				limite=limite+1
				Paginacao = Paginacao+1
			   LastRow = Page.Canvas.DrawTable( Table, param_table1 )
	
				if LastRow >= Table.Rows.Count Then 
			    	Exit Do ' entire table displayed
				else
				
					 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width=780; alignment=left; size=8; color=#000000")
					
					Relatorio = "SWD017 - Sistema Web Diretor"
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
					 ' Display remaining part of table on the next page
					Set Page = Page.NextPage	
					
					param_table2.Add( "RowTo=1; RowFrom=1" ) ' Row 1 is header.
					param_table2("RowFrom1") = LastRow + 1 ' RowTo1 is omitted and presumed infinite
'NOVO CABEÇALHO==========================================================================================		
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
		Text = "<p><center><i><b><font style=""font-size:18pt;"">Entrevistas</font></b></i></center></p>"
		

		Do While Len(Text) > 0
			CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
		 
			If CharsPrinted = Len(Text) Then Exit Do
				SET Page = Page.NextPage
			Text = Right( Text, Len(Text) - CharsPrinted)
		Loop 

	'================================================================================================================			
			 	end if
'				if limite>300 then
'					response.Write("ERRO!")
'					response.end()
'				end if 
			Loop

		 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width=780; alignment=left; size=8; color=#000000")
		
		Relatorio = "SWD017 - Sistema Web Diretor"
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
	


		



arquivo="SWD017"
Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

