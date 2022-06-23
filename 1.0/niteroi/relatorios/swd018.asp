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
dados_informados = split(obr, "$!$" )
co_matric = dados_informados(0)
da_entrevista = dados_informados(1)
ho_entrevista = dados_informados(2)

if ori="ebe" then
origem="../ws/doc/ofc/ebe/"
end if

if mes<10 then
mes="0"&mes
end if

data = dia &"/"& mes &"/"& ano

if min<10 then
min="0"&min
end if

horario = hora & ":"& min

	hora_split= Split(ho_entrevista,":")
	hora=hora_split(0)
	min=hora_split(1)
	
	ho_entrevista=hora&":"&min
	
	data_split= Split(da_entrevista,"/")
	dia=data_split(0)
	mes=data_split(1)
	ano=data_split(2)
	
	da_entrevista_cons=mes&"/"&dia&"/"&ano	
	ho_entrevista_cons=hora&":"&min


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

Set RSA = Server.CreateObject("ADODB.Recordset")
CONEXAOA = "Select * from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.CO_Matricula = "& co_matric
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
	SQL1 = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod_cons
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
		'no_etapa = no_etapa&" "&co_concordancia_curso&" "&no_abrv_curso				

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
		y_texto=formatnumber(Page.Height - altura_logo_gde/2,0)
		width_texto=Page.Width - (margem*2)

	
		SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
		Text = "<p><center><i><b><font style=""font-size:18pt;"">Conte&uacute;do da Entrevista</font></b></i></center></p>"
		

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


		Set param_table1 = Pdf.CreateParam("width=533; height=25; rows=2; cols=8; border=0; cellborder=0; cellspacing=0;")
		Set Table = Doc.CreateTable(param_table1)
		Table.Font = Font
		y_table=Page.Height - altura_logo_gde-70
		
		With Table.Rows(1)
		   .Cells(1).Width = 40
		   .Cells(2).Width = 205
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

			Set RSo = Server.CreateObject("ADODB.Recordset")
			SQLo = "SELECT * FROM TB_Entrevistas WHERE CO_Matricula ="&co_matric&" AND (DA_Entrevista=#"&da_entrevista_cons&"# AND mid(HO_Entrevista,1,16)=#12/30/1899 "&ho_entrevista_cons&"#)" 
			RSo.Open SQLo, CON4
			
		tp_entrevista=RSo("TP_Entrevista")
		partic_entrevista=RSo("NO_Participantes")
		st_entrevista=RSo("ST_Entrevista")
		ag_entrevista=RSo("CO_Agendado_com")
		ob_entrevista=RSo("TX_Observa")
		cu_entrevista=RSo("CO_Usuario")		
		
		data_exibe = split(da_entrevista,"/")
		data_exibe(0) = data_exibe(0)*1
		if data_exibe(0)<10 then
			dia_txt="0"&data_exibe(0)
		else	
			dia_txt=data_exibe(0)	
		end if	
		Select case data_exibe(1)
		
			case 1
			mes_txt="janeiro"
			
			case 2
			mes_txt="fevereiro"
			
			case 3
			mes_txt="mar&ccedil;o"
			
			case 4
			mes_txt="abril"
			
			case 5
			mes_txt="maio"
	
			case 6
			mes_txt="junho"
			
			case 7
			mes_txt="julho"
			
			case 8
			mes_txt="agosto"		
			
			case 9
			mes_txt="setembro"
	
			case 10
			mes_txt="outubro"
			
			case 11
			mes_txt="novembro"
			
			case 12
			mes_txt="dezembro"				
		end select		
		
		data_txt = 	dia_txt&" de "&mes_txt&" de "&data_exibe(2)	
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Tipo_Entrevista where TP_Entrevista = "&tp_entrevista
		RS.Open SQL, CON0
					
		tipo_txt= RS("TX_Descricao")
		
		
		if ag_entrevista="" or isnull(ag_entrevista) then
			no_agendado=""
		else
			Set RSu = Server.CreateObject("ADODB.Recordset")
			SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& ag_entrevista
			RSu.Open SQLu, CON
		
			IF RSu.EOF then
				no_atendido =""	
			else
				no_atendido =RSu("NO_Usuario")
			end if
				
		end if
			
		Select case st_entrevista		
			case 1
			nome_status="Atendida"
			
			case 2
			nome_status="Cancelada"
			
			case 3
			nome_status="Pendente"	
		end select			

		Set param_table1 = Pdf.CreateParam("width=533; height=50; rows=4; cols=5; border=0; cellborder=0.5; cellspacing=0;")
		Set Table = Doc.CreateTable(param_table1)
		Table.Font = Font
		y_table=y_table-35

		data_txt = replace_latin_char(data_txt,"html")	
		tipo_txt = replace_latin_char(tipo_txt,"html")	
		no_atendido = replace_latin_char(no_atendido,"html")
		nome_status = replace_latin_char(nome_status,"html")													
		partic_entrevista = replace_latin_char(partic_entrevista,"html")		
		
		With Table.Rows(1)
		   .Cells(1).Width = 116
		   .Cells(2).Width = 96
		   .Cells(3).Width = 97
		   .Cells(4).Width = 137
		   .Cells(5).Width = 87   
		End With
		Table(3, 1).ColSpan = 5
		Table(4, 1).ColSpan = 5		
		Table(1, 1).AddText "<div align=""center""><b>Data</b></div>", "size=9; alignment=center; html=true", Font 
		Table(1, 2).AddText "<div align=""center""><b>Hora</b></div>", "size=9; alignment=center; html=true", Font 
		Table(1, 3).AddText "<div align=""center""><b>Tipo</b></div>", "size=9; alignment=center; html=true", Font 
		Table(1, 4).AddText "<div align=""center""><b>Atendimento</b></div>", "size=9; alignment=center; html=true", Font 
		Table(1, 5).AddText "<div align=""center""><b>Status</b></div>", "size=9; alignment=center; html=true", Font 
		Table(2, 1).AddText "<div align=""center"">"&data_txt&"</div>", "size=9; alignment=center; html=true", Font 
		Table(2, 2).AddText "<div align=""center"">"&ho_entrevista&"</div>", "size=9; alignment=center; html=true", Font 
		Table(2, 3).AddText "<div align=""center"">"&tipo_txt&"</div>", "size=9; alignment=center; html=true", Font 
		Table(2, 4).AddText "<div align=""center"">"&no_atendido&"</div>", "size=9; alignment=center; html=true", Font 
		Table(2, 5).AddText "<div align=""center"">"&nome_status&"</div>", "size=9; alignment=center; html=true", Font 	
		Table(3, 1).AddText "<b>Participantes</b>", "size=9; indentx=15; alignment=left; html=true", Font 
		Table(4, 1).AddText partic_entrevista, "size=9; indentx=15; alignment=left; html=true", Font 						
		Page.Canvas.DrawTable Table, "x="&margem&", y="&y_table&"" 
		
		y_sub_titulo = y_separador-100
		SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_sub_titulo&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
		Text = "<b><font style=""font-size:9pt;"">Conte&uacute;do da Entrevista</font></b>"
		Do While Len(Text) > 0
			CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )		 
			If CharsPrinted = Len(Text) Then Exit Do
				SET Page = Page.NextPage
			Text = Right( Text, Len(Text) - CharsPrinted)
		Loop 		
		
		y_separador = y_sub_titulo-15
		
		Page.Canvas.SetParams "LineWidth=2" 
		Page.Canvas.SetParams "LineCap=0" 
		With Page.Canvas
		   .MoveTo margem, y_separador
		   .LineTo Page.Width - margem, y_separador
		   .Stroke
		End With 

		'=============================================================================================================	

		Set RSC = Server.CreateObject("ADODB.Recordset")
		SQLC = "SELECT * FROM TB_Entrevistas_Conteudo WHERE CO_Matricula ="& co_matric&" AND (DA_Entrevista=#"&da_entrevista_cons&"# AND mid(HO_Entrevista,1,16)=#12/30/1899 "&ho_entrevista_cons&"#)" 
		RSC.Open SQLC, CON4
		
		if RSC.EOF then
			texto = ""
		else 
			texto = RSC("TX_Conteudo")
		end if
			
		
		y_texto = y_separador-5
		
		SET Param_Texto = Pdf.CreateParam("x="&margem&";y="&y_texto&"; size=9; height=600; width=533; alignment=left")			
		
		Do While Len(texto) > 0
			CharsPrinted = Page.Canvas.DrawText(texto, Param_Texto, Font )
		 
			If CharsPrinted = Len(texto) Then Exit Do
				texto = Right( texto, Len(texto) - CharsPrinted)				
				SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width=533; alignment=left; size=8; color=#000000")
		
				Relatorio = "SWD018 - Sistema Web Diretor"
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
		Text = "<p><center><i><b><font style=""font-size:18pt;"">Conte&uacute;do da Entrevista</font></b></i></center></p>"
		

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
		With Page.Canvas
		   .MoveTo margem, Page.Height - altura_logo_gde-65
		   .LineTo Page.Width - margem, Page.Height - altura_logo_gde-65
		   .Stroke
		End With 	


		Set param_table1 = Pdf.CreateParam("width=533; height=50; rows=4; cols=5; border=0; cellborder=0.5; cellspacing=0;")
		Set Table = Doc.CreateTable(param_table1)
		Table.Font = Font
		y_table=y_table-35

		data_txt = replace_latin_char(data_txt,"html")	
		tipo_txt = replace_latin_char(tipo_txt,"html")	
		no_atendido = replace_latin_char(no_atendido,"html")
		nome_status = replace_latin_char(nome_status,"html")													
		partic_entrevista = replace_latin_char(partic_entrevista,"html")		
		
		With Table.Rows(1)
		   .Cells(1).Width = 116
		   .Cells(2).Width = 96
		   .Cells(3).Width = 97
		   .Cells(4).Width = 137
		   .Cells(5).Width = 87   
		End With
		Table(3, 1).ColSpan = 5
		Table(4, 1).ColSpan = 5		
		Table(1, 1).AddText "<div align=""center""><b>Data</b></div>", "size=9; alignment=center; html=true", Font 
		Table(1, 2).AddText "<div align=""center""><b>Hora</b></div>", "size=9; alignment=center; html=true", Font 
		Table(1, 3).AddText "<div align=""center""><b>Tipo</b></div>", "size=9; alignment=center; html=true", Font 
		Table(1, 4).AddText "<div align=""center""><b>Atendimento</b></div>", "size=9; alignment=center; html=true", Font 
		Table(1, 5).AddText "<div align=""center""><b>Status</b></div>", "size=9; alignment=center; html=true", Font 
		Table(2, 1).AddText "<div align=""center"">"&data_txt&"</div>", "size=9; alignment=center; html=true", Font 
		Table(2, 2).AddText "<div align=""center"">"&ho_entrevista&"</div>", "size=9; alignment=center; html=true", Font 
		Table(2, 3).AddText "<div align=""center"">"&tipo_txt&"</div>", "size=9; alignment=center; html=true", Font 
		Table(2, 4).AddText "<div align=""center"">"&no_atendido&"</div>", "size=9; alignment=center; html=true", Font 
		Table(2, 5).AddText "<div align=""center"">"&nome_status&"</div>", "size=9; alignment=center; html=true", Font 	
		Table(3, 1).AddText "<b>Participantes</b>", "size=9; indentx=15; alignment=left; html=true", Font 
		Table(4, 1).AddText partic_entrevista, "size=9; indentx=15; alignment=left; html=true", Font 						
		Page.Canvas.DrawTable Table, "x="&margem&", y="&y_table&"" 
		
		y_sub_titulo = y_separador-100
		SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_sub_titulo&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
		Text = "<b><font style=""font-size:9pt;"">Conte&uacute;do da Entrevista</font></b>"
		Do While Len(Text) > 0
			CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )		 
			If CharsPrinted = Len(Text) Then Exit Do
				SET Page = Page.NextPage
			Text = Right( Text, Len(Text) - CharsPrinted)
		Loop 		
		
		y_separador = y_sub_titulo-15
		
		Page.Canvas.SetParams "LineWidth=2" 
		Page.Canvas.SetParams "LineCap=0" 
		With Page.Canvas
		   .MoveTo margem, y_separador
		   .LineTo Page.Width - margem, y_separador
		   .Stroke
		End With 

		'==============================================================================================================							
				

		Loop 				
'			
		SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width=533; alignment=left; size=8; color=#000000")
		
		Relatorio = "SWD018 - Sistema Web Diretor"
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
	
end if			



arquivo="SWD018"
Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

