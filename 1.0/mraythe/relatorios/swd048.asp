<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 60 'valor em segundos
'FICHA INDIVIDUAL
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../../global/funcoes_diversas.asp"-->
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

separa_dados=split(dados,"$$$")
tipo_busca=split(separa_dados(0),"$!$")
opt=tipo_busca(0)

dados_turma=separa_dados(1)
dados_informados = split(dados_turma, "$!$" )

unidade=dados_informados(0)
curso=dados_informados(1)
co_etapa=dados_informados(2)
turma=dados_informados(3)
periodo_form=dados_informados(4)


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
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0	

		Set CONt = Server.CreateObject("ADODB.Connection") 
		ABRIRt = "DBQ="& CAMINHO_t & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONt.Open ABRIRt
if opt="a" then
	dados_aluno=tipo_busca(1)
			
	If Not IsArray(alunos_encontrados) Then alunos_encontrados = Array() End if	
	ReDim preserve alunos_encontrados(UBound(alunos_encontrados)+1)	
	alunos_encontrados(Ubound(alunos_encontrados)) = dados_aluno
	
elseif opt="t" then

	if unidade="999990" or unidade="" or isnull(unidade) then
		SQL_ALUNOS="NULO"
	else	
		SQL_ALUNOS= "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.NU_Unidade = "& unidade		
		if curso="999990" or curso="" or isnull(curso) then
			SQL_CURSO=""
		else
			SQL_CURSO=" AND TB_Matriculas.CO_Curso = '"& curso &"'"			
		end if
	
		if co_etapa="999990" or co_etapa="" or isnull(co_etapa) then
			SQL_ETAPA=""		
		else
			SQL_ETAPA=" AND TB_Matriculas.CO_Etapa = '"& co_etapa &"'"				
		end if
	
		if turma="999990" or turma="" or isnull(turma) then
			SQL_TURMA=""		
		else
			SQL_TURMA=" AND TB_Matriculas.CO_Turma = '"& turma &"' "			
		end if
	
	SQL_ALUNOS= SQL_ALUNOS&SQL_CURSO&SQL_ETAPA&SQL_TURMA&" order by TB_Matriculas.NU_Unidade,TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma, TB_Alunos.NO_Aluno"
	end if

	if SQL_ALUNOS="NULO" then
	else
	
	nu_chamada_check = 1
		Set RSA = Server.CreateObject("ADODB.Recordset")
		CONEXAOA = SQL_ALUNOS
'response.Write(CONEXAOA)
		Set RSA = CON1.Execute(CONEXAOA)
		vetor_matriculas="" 
		While Not RSA.EOF
			nu_matricula = RSA("CO_Matricula")
			nu_chamada = RSA("NU_Chamada")
			if nu_chamada_check = 1 and nu_chamada=nu_chamada_check then
				vetor_matriculas=nu_matricula
			elseif nu_chamada_check = 1 then
				while nu_chamada_check < nu_chamada
					nu_chamada_check=nu_chamada_check+1
				wend 
				vetor_matriculas=nu_matricula
			else
				vetor_matriculas=vetor_matriculas&"#!#"&nu_matricula
			end if
		nu_chamada_check=nu_chamada_check+1			
		RSA.MoveNext
		Wend 
	
	end if	
'RESPONSE.Write(vetor_matriculas)
'RESPONSE.END()
'	matriculas_encontradas = split(vetor_matriculas, "#!#" )	

	alunos_encontrados = split(vetor_matriculas, "#!#" )	

	RSA.Close
	Set RSA = Nothing
end if

Set RSano = Server.CreateObject("ADODB.Recordset")
SQLano = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&ano_letivo&"'"
RSano.Open SQLano, CON

if RSano.eof then
	response.Write("Erro no Ano Letivo")
	response.end()
else
	if RSano("ST_Ano_Letivo")="L" then
		verifica_periodos="s"
	else
		verifica_periodos="n"
	end if
end if	

if periodo_form=0 then
	verifica_periodos="n"
end if	

max_notas_exibe=1*periodo_form
max_notas_exibe=max_notas_exibe

	
For alne=0 to ubound(alunos_encontrados)	
	cod_cons=alunos_encontrados(alne)

		
	Set RS0 = Server.CreateObject("ADODB.Recordset")
	SQL0 = "SELECT * FROM TB_Periodo ORDER BY NU_Periodo"
	RS0.Open SQL0, CON0
	check_periodo=1
	WHILE NOT RS0.EOF
		periodo=RS0("NU_Periodo")
		if check_periodo=1 then
			vetor_periodo=periodo
		else
			vetor_periodo=vetor_periodo&"#!#"&periodo
		end if
		check_periodo=check_periodo+1 
	RS0.MOVENEXT
	WEND				
		
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod_cons
	RS.Open SQL, CON1
	
	nome_aluno = RS("NO_Aluno")
	sexo_aluno = RS("IN_Sexo")
	nome_pai = RS("NO_Pai")
	nome_mae = RS("NO_Mae")
	natural= RS("SG_UF_Natural")	
	nome_aluno=replace_latin_char(nome_aluno,"html")	

	if nome_pai="" or isnull(nome_pai) then
		nome_pai=" "
	else
		nome_pai=replace_latin_char(nome_pai,"html")	
	end if
	
	if nome_mae="" or isnull(nome_mae) then
		nome_mae=" "
	else	
		nome_mae=replace_latin_char(nome_mae,"html")
	end if	
	
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
		id_unico= RS("RA_Aluno")	
		
		curso=curso*1		
		if curso=0 then
			if opt="01" then
			response.redirect(origem&"index.asp?nvg="&nvg&"&opt=err2")
			end if
		else
			Set RStabela = Server.CreateObject("ADODB.Recordset")
			SQLtabela = "SELECT * FROM TB_Da_Aula where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& etapa &"' AND CO_Turma ='"& turma &"'" 
			RStabela.Open SQLtabela, CON2
	
			if 	RStabela.EOF then
					response.Write("ERRO1 - N&atilde;o cadastrado TP_Nota em TB_Da_Aula para NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& etapa &"' AND CO_Turma ='"& turma &"'" )
					response.end()
			else				
				tb_nota=RStabela("TP_Nota")		
					if tb_nota ="TB_NOTA_A" then
						caminho_nota = CAMINHO_na
					elseif tb_nota="TB_NOTA_B" then
						caminho_nota = CAMINHO_nb
					elseif tb_nota ="TB_NOTA_C" then
						caminho_nota = CAMINHO_nc
					elseif tb_nota ="TB_NOTA_D" then
						caminho_nota = CAMINHO_nd
					elseif tb_nota ="TB_NOTA_E" then
						caminho_nota = CAMINHO_ne						
					else
						response.Write("ERRO2 - N&atilde;o cadastrado TP_Nota em TB_Da_Aula para NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& etapa &"' AND CO_Turma ='"& turma &"'" )
						response.end()
					end if	
				end if
				

			
				call GeraNomes("PORT",unidade,curso,etapa,CON0)
				no_unidade = session("no_unidades")
				no_curso= session("no_grau")
				no_etapa = session("no_serie")
				
				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL2 = "SELECT * FROM TB_Unidade WHERE NU_Unidade="& unidade
				RS2.Open SQL2, CON0	
						
				sede=RS2("NO_Sede")
						
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL3 = "SELECT * FROM TB_Curso WHERE CO_Curso='"& curso &"'"
				RS3.Open SQL3, CON0
				
				no_abrv_curso = RS3("NO_Abreviado_Curso")
				co_concordancia_curso = RS3("CO_Conc")	
				
				no_unidade = unidade&" - "&no_unidade
				'no_curso= no_etapa&" "&co_concordancia_curso&" "&no_curso				
				no_curso= no_curso&" - "&no_etapa
				texto_turma = "<div align='Left'> - Turma:"&turma&"</div>"
				texto_chamada=" N&ordm;. Chamada: "&cham	
				
				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL2 = "SELECT * FROM TB_Contatos WHERE CO_Matricula="& cod_cons &" AND TP_Contato='ALUNO'"
				RS2.Open SQL2, CONCONT
	
				nascimento=RS2("DA_Nascimento_Contato")
	
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL3 = "SELECT * FROM TB_UF WHERE SG_UF='"& natural &"'"
				RS3.Open SQL3, CON0
				
			no_uf = RS3("NO_UF")						

			SET Page = Doc.Pages.Add( 595, 842 )
					
'CABEÇALHO==========================================================================================		
			Set Param_Logo_Gde = Pdf.CreateParam
			margem=30				
			largura_logo_gde=formatnumber(Logo.Width*0.4,0)
			altura_logo_gde=formatnumber(Logo.Height*0.4,0)
			area_utilizavel=Page.Width - (margem*2)
			Param_Logo_Gde("x") = margem
			Param_Logo_Gde("y") = Page.Height - altura_logo_gde -22
			Param_Logo_Gde("ScaleX") = 0.4
			Param_Logo_Gde("ScaleY") = 0.4
			Page.Canvas.DrawImage Logo, Param_Logo_Gde
	
			x_texto=largura_logo_gde+ margem+10
			'x_texto= margem
			y_texto=formatnumber(Page.Height - margem,0)
			'width_texto=Page.Width - 60
			width_texto=Page.Width -7*margem
		
			Set RSC = Server.CreateObject("ADODB.Recordset")
			SQLC = "SELECT * FROM TB_Cabecalhos WHERE CO_Documento=3"
			RSC.Open SQLC, CON0			

			documento = RSC("NO_Documento")
			txt1 = RSC("NO_Cabec_1")
			txt2 = RSC("NO_Cabec_2")
			txt3 = RSC("NO_Cabec_3")
			txt4 = RSC("NO_Cabec_4")		
			txt5 = RSC("NO_Cabec_5")		
			txt6 = RSC("NO_Cabec_6")													

			SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde+30&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
			Text = "<p><i><b>"&UCASE(txt1)&"</b></i></p><br><font style=""font-size:10pt;"">"&txt2&"<BR>"&txt3&"<BR>"&txt4&"<BR>"&txt5&"<BR>"&txt6&"</FONT>" 			

			
			Do While Len(Text) > 0
				CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
			 
				If CharsPrinted = Len(Text) Then Exit Do
					SET Page = Page.NextPage
				Text = Right( Text, Len(Text) - CharsPrinted)
			Loop 

			y_texto=y_texto-altura_logo_gde-10
			SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height=30; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
			Text = "<center><i><b><font style=""font-size:18pt;"">"&documento&"</font></b></i></center>"
			
			
			Do While Len(Text) > 0
				CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
			 
				If CharsPrinted = Len(Text) Then Exit Do
					SET Page = Page.NextPage
				Text = Right( Text, Len(Text) - CharsPrinted)
			Loop 	
			
'================================================================================================================			
			Page.Canvas.SetParams "LineWidth=2" 
			Page.Canvas.SetParams "LineCap=0" 
			altura_primeiro_separador= Page.Height - altura_logo_gde-margem
			With Page.Canvas
			   .MoveTo margem, altura_primeiro_separador
			   .LineTo Page.Width - margem, altura_primeiro_separador
			   .Stroke
			End With 
							
			Page.Canvas.SetParams "LineWidth=2" 
			Page.Canvas.SetParams "LineCap=0" 
			altura_segundo_separador= Page.Height - altura_logo_gde -margem - 45
			With Page.Canvas
			   .MoveTo margem, altura_segundo_separador
			   .LineTo area_utilizavel+margem, altura_segundo_separador
			   .Stroke
			End With 	
							
			Set param_table1 = Pdf.CreateParam("width="&area_utilizavel&"; height=72; rows=6; cols=8; border=0; cellborder=0; cellspacing=0;")
			Set Table = Doc.CreateTable(param_table1)
			Table.Font = Font
			y_primeira_tabela=altura_segundo_separador-10
			x_primeira_tabela=margem+5
			With Table.Rows(1)
			   .Cells(1).Width = 50
			   .Cells(2).Width = 15  
			   .Cells(3).Width = 20	
			   .Cells(4).Width = 70	
			   .Cells(5).Width = 20	
			   .Cells(6).Width = 50	
   			   .Cells(7).Width = 90				   		   		   
			End With
			
			Table(1, 2).ColSpan = 3
			Table(1, 5).ColSpan = 2	
			Table(2, 2).ColSpan = 7
			Table(3, 2).ColSpan = 5
			Table(4, 2).ColSpan = 5
			Table(5, 3).ColSpan = 2	
			Table(5, 5).ColSpan = 3													
			Table(6, 2).ColSpan = 2
			Table(6, 4).ColSpan = 3	
			
			Table(1, 1).AddText "Matr&iacute;cula:", "size=8;html=true", Font 
			Table(1, 2).AddText "<b>"&cod_cons&"</b>", "size=8;html=true", Font 
			Table(1, 5).AddText "ID-INEP:", "size=8;html=true", Font 
			Table(1, 7).AddText "<b>"&id_unico&"</b>", "size=8;html=true", Font 				
			Table(2, 1).AddText "Alun"&desinencia&":", "size=8;", Font 
			Table(2, 2).AddText "<b>"&nome_aluno&"</b>", "size=8;html=true", Font 
			Table(3, 1).AddText "Pai: ", "size=8;", Font 
			Table(3, 2).AddText nome_pai, "size=8;html=true", Font 
			Table(4, 1).AddText "M&atilde;e: ", "size=8;html=true", Font 
			Table(4, 2).AddText nome_mae, "size=8;html=true", Font	
			Table(5, 1).AddText "Sexo: ", "size=8;", Font 
			Table(5, 2).AddText sexo_aluno, "size=8;", Font 
			Table(5, 3).AddText " Nascimento: "&nascimento, "size=8;", Font 							
			Table(5, 5).AddText " Natural de: "&no_uf, "size=8;", Font		
			Table(6, 1).AddText "Ano Letivo: ", "size=8;html=true", Font 
			Table(6, 2).AddText ano_letivo, "size=8;", Font	
			Table(6, 4).AddText "Curso: "& no_curso, "size=8;", Font
			Table(6, 7).AddText texto_turma, "size=8;html=true", Font												
			Table(6, 8).AddText texto_chamada, "size=8;html=true", Font	
			Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 		

			if curso=1 and co_etapa<6 then
				sem_carga="s"
				cols=12
				width=280			
			else
				sem_carga="n"
				cols=13		
				width=250						
			end if		
		
			altura_medias=40
			Set param_table2 = Pdf.CreateParam("width=533; height="&altura_medias&"; rows=3; cols="&cols&"; border=1; cellborder=0.1; cellspacing=0;")
			Set Notas_Tit = Doc.CreateTable(param_table2)
			Notas_Tit.Font = Font
			y_medias=y_primeira_tabela-72-10	

			
			With Notas_Tit.Rows(1)
			   .Cells(1).Width = width
			   .Cells(2).Width = 23
			   .Cells(3).Width = 23
			   .Cells(4).Width = 23
			   .Cells(5).Width = 23
			   .Cells(6).Width = 23
			   .Cells(7).Width = 23
			   .Cells(8).Width = 23   
			   .Cells(9).Width = 23
			   .Cells(10).Width = 23
			   .Cells(11).Width = 23   
			   .Cells(12).Width = 23 
'			   .Cells(13).Width = 23			         
'			   .Cells(14).Width = 23 			         
'			   .Cells(15).Width = 23 	
			  if sem_carga="n" then 		         
			   .Cells(13).Width = 30 			         
			  end if		         
			End With
			Notas_Tit(1, 1).RowSpan = 3	
			if sem_carga="n" then 			
				Notas_Tit(1, 13).RowSpan = 3	
			end if			
			Notas_Tit(2, 2).RowSpan = 2	
			Notas_Tit(2, 3).RowSpan = 2	
			Notas_Tit(2, 4).RowSpan = 2	
			Notas_Tit(2, 5).RowSpan = 2	
			Notas_Tit(2, 6).RowSpan = 2	
			Notas_Tit(2, 7).RowSpan = 2	
			Notas_Tit(2, 8).RowSpan = 2	
			Notas_Tit(2, 9).RowSpan = 2	
			Notas_Tit(2, 10).RowSpan = 2	
			Notas_Tit(2, 11).RowSpan = 2	
			Notas_Tit(2, 12).RowSpan = 2	
			'Notas_Tit(2, 13).RowSpan = 2
			'Notas_Tit(2, 14).RowSpan = 2
			'Notas_Tit(2, 15).RowSpan = 2																																																	
			Notas_Tit(1, 2).ColSpan = 7	
			'Notas_Tit(1, 9).ColSpan = 3					 			
			Notas_Tit(1, 9).ColSpan = 4																																																																									
			Notas_Tit(1, 1).AddText "<div align=""center""><b>Disciplinas</b></div>", "size=10;indenty=15; html=true", Font 
			Notas_Tit(1, 2).AddText "<div align=""center""><b>Aproveitamento</b></div>", "size=8;alignment=center; indenty=1;html=true", Font  
	'		Notas_Tit(1, 9).AddText "<div align=""center""><b>M&eacute;dia (da Turma)</b></div>", "size=8;alignment=center; indenty=1;html=true", Font
			Notas_Tit(1, 9).AddText "<div align=""center""><b>Frequ&ecirc;ncia</b></div>", "size=8;alignment=center; indenty=1;html=true", Font
			Notas_Tit(2, 2).AddText "PA1", "size=7;alignment=center;indenty=6;", Font 
			Notas_Tit(2, 3).AddText "PA2", "size=7;alignment=center;indenty=6;", Font 
			Notas_Tit(2, 4).AddText "PA3", "size=7;alignment=center;indenty=6;", Font 
			Notas_Tit(2, 5).AddText "Total", "size=7 ;alignment=center;indenty=6;", Font 
			Notas_Tit(2, 6).AddText "<div align=""center"">4&ordf; Aval.<br>p.2</div>", "size=7;alignment=center;indenty=2; html=true", Font 
			Notas_Tit(2, 7).AddText "Total", "size=7;alignment=center;indenty=6;", Font 
			Notas_Tit(2, 8).AddText "<div align=""center"">M&eacute;dia<br>Final</div>", "size=7;alignment=center;indenty=2;html=true", Font 
'			Notas_Tit(2, 9).AddText "PA1", "size=7;alignment=center;indenty=6;", Font  
'			Notas_Tit(2, 10).AddText "PA2", "size=7;alignment=center;indenty=6;", Font 
'			Notas_Tit(2, 11).AddText "PA3", "size=7;alignment=center;indenty=6;", Font 
			Notas_Tit(2, 9).AddText "PA1", "size=7;alignment=center;indenty=6;", Font 
			Notas_Tit(2, 10).AddText "PA2", "size=7;alignment=center;indenty=6;", Font 
			Notas_Tit(2, 11).AddText "PA3", "size=7;alignment=center;indenty=6;", Font 
			Notas_Tit(2, 12).AddText "Total", "size=7;alignment=center;indenty=6;", Font 
			if sem_carga="n" then 				
				Notas_Tit(1, 13).AddText "Carga", "size=9;alignment=center;indenty=15;", Font 	
			end if				
			Set param_materias = PDF.CreateParam	
			param_materias.Set "expand=true" 
			
			acumula_peso_periodo=0			
			for wrk_periodo = 1 to 3
				
				Set RSPESO = Server.CreateObject("ADODB.Recordset")
				SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&wrk_periodo
				RSPESO.Open SQLPESO, CON0
				
				peso_periodo=RSPESO("NU_Peso")
				acumula_peso_periodo=acumula_peso_periodo+peso_periodo
			next				

				Set RS0 = Server.CreateObject("ADODB.Recordset")
				SQL = "SELECT * FROM TB_Boletim_Cabecalho where NU_Unidade="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'"
				Set RS0 = CONt.Execute(SQL)
				n_campo=0
				campo="inicio"
				resultados=""
				resultados1=""
				resultados2=""					
				while not campo="fim"
					wrk_exibe_resultado="s" 				
					n_campo=n_campo*1	
					n_campo=n_campo+1	
					if n_campo<10 then
						campo_gravacao="CO_0"&n_campo
					else
						campo_gravacao="CO_"&n_campo						
					end if					
					campo_exibe=RS0(campo_gravacao)
							
					if campo_exibe="" or isnull(campo_exibe) then
						campo="fim"
					else
						campo="continua"				
					end if
					
					if  n_campo>50 then
						campo="fim"					
					end if
				WEND		
				'é menos 2 por que o último campo é o que o programa achou como vazio
				total_campos=n_campo-1	
				
				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL2 = "Select * from TB_Boletim_Notas WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"' AND CO_Matricula="&cod_cons& " ORDER BY NU_Seq"
				Set RS2 = CONt.Execute(SQL2)	
				
				conta_linhas=0	
				resultados=""
				horas_aula=0
				total_faltas=0
				tipo_resultado="A"	
			
				while not RS2.EOF		
					
					Set Row = Notas_Tit.Rows.Add(20) ' row height
					param_materias.Add "indenty=3;alignment=right;html=true"
'					Row.Cells(1).AddText no_materia_exibe, param_materias
					altura_medias=altura_medias+20

					
					param_materias.Add "indentx=0"				
					for nn=1 to total_campos	
						if nn = 9 or nn=10 or nn=11 then
						else	
							if nn<10 then
								campo_gravacao="CO_0"&nn
							else
								campo_gravacao="CO_"&nn						
							end if					
							notas_exibe=RS2(campo_gravacao)
							'response.Write(campo_gravacao)
							if notas_exibe="" or isnull(notas_exibe) then
								nota=" "
								if nn<5 then
									wrk_exibe_resultado="n"
								end if									
							else
								nota=notas_exibe	
								nn=nn*1
								total_campos=total_campos*1
								campo_faltas=total_campos-1
								if nn=5 then
									if nota="" or isnull(nota) then
									
									else
	
										if wrk_exibe_resultado="n" then
											resultados1=""
										else	
											media_anual= nota/acumula_peso_periodo
											media_anual=media_anual*10			
											decimo = media_anual - Int(media_anual)
											If decimo >= 0.5 Then
												nota_arredondada = Int(media_anual) + 1
												media_anual=nota_arredondada
											else
												nota_arredondada = Int(media_anual)
												media_anual=nota_arredondada											
											End If	
											media_anual=media_anual/10						
											media_anual = formatNumber(media_anual,1)											
																			
											IF resultados1="" and conta_linhas=0 then
												resultados1=media_anual	
											else									
												resultados1=resultados1&"#!##$#"&media_anual	
											end if	
										end if
									end if	
								end if									
								if nn=6 then
									if nota="" or isnull(nota) then	
									else
										tipo_resultado="F"							
									end if	
								end if									
								if nn=8 then
'									if nota="" or isnull(nota) then
'									else
'										IF resultados="" and conta_linhas=0 then
'											resultados=nota	
'										else									
'											resultados=resultados&"#!##$#"&nota	
'										end if	
'									end if	
									if nota="" or isnull(nota) then
									else
										IF resultados2="" and conta_linhas=0 then
											resultados2=nota	
										else									
											resultados2=resultados2&"#!##$#"&nota	
										end if								
									end if	
								end if								
								if nn=campo_faltas then
									total_faltas=total_faltas*1
									nota=nota*1
									if nota="" or isnull(nota) then
									else
										total_faltas=total_faltas+nota	
									end if	
								end if						
								if nn=total_campos then
									horas_aula=horas_aula*1
									nota=nota*1
									horas_aula=horas_aula+nota
								end if	
							end if	
							nn=nn*1			
							if nn=1 then
								alinha="left"
								espaco="&nbsp;"
							else
								alinha="center"	
								espaco=""											
							end if
							if sem_carga="s" and nn=total_campos then 	
							else
							    if nn<9 then	
									indice = nn	
								else
									indice = nn-3
								end if					
								Row.Cells(indice).AddText "<div align="""&alinha&"""><font style=""font-size:8pt;"">"&espaco&nota&"</font></div>", param_materias			
							end if	
						end if	
					next	
					conta_linhas=conta_linhas+1					
				RS2.MOVENEXT
				WEND							

			Page.Canvas.DrawTable Notas_Tit, "x=30, y="&y_medias&"" 		



'	
'			Set CON_N = Server.CreateObject("ADODB.Connection") 
'			ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
'			CON_N.Open ABRIRn
'						
'			Set RSF = Server.CreateObject("ADODB.Recordset")
'			SQL_N = "Select * from TB_Frequencia_Periodo WHERE CO_Matricula = "& cod_cons
'			Set RSF = CON_N.Execute(SQL_N)
'			
'			soma_faltas=0			
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
'				
'				if isnull(f1) or f1= "" then
'				else
'					f1=f1*1
'					soma_faltas=soma_faltas*1
'					soma_faltas=soma_faltas+f1		
'				end if
'				
'				if isnull(f2) or f2= "" then
'				else
'					f2=f2*1
'					soma_faltas=soma_faltas*1
'					soma_faltas=soma_faltas+f2		
'				end if
'				
'				if isnull(f3) or f3= "" then
'				else
'					f3=f3*1
'					soma_faltas=soma_faltas*1
'					soma_faltas=soma_faltas+f3		
'				end if
'				
'				if isnull(f4) or f4= "" then
'				else
'					f4=f4*1
'					soma_faltas=soma_faltas*1
'					soma_faltas=soma_faltas+f4		
'				end if									
'			END IF				

			'soma_faltas=soma_faltas*1
			'dias_de_aula_no_ano=200
			
			'frequencia=((dias_de_aula_no_ano-soma_faltas)/dias_de_aula_no_ano)*100
			total_faltas=total_faltas*1
			if horas_aula="" or isnull(horas_aula) then
				frequencia=100
			elseif horas_aula=0 then
				frequencia=100			
			else	
				curso = curso*1
				co_etapa = co_etapa*1
				if curso=1 and co_etapa<6 then
					horas_aula= 1000
				end if

                                if ano_letivo>2019 and curso=1 and co_etapa=1 then
					horas_aula= 800
				end if

				horas_aula=horas_aula*1
				frequencia=((horas_aula-total_faltas)/horas_aula)*100
			end if	
			if frequencia<100 then
				frequencia=arredonda(frequencia,"mat_dez",1,0)	
			end if	
'		response.Write(frequencia&"+"&horas_aula&"-"&total_faltas)
'		response.End()
				Set param_table3 = Pdf.CreateParam("width=533; height=20; rows=2; cols=10; border=0; cellborder=0; cellspacing=0;")
				Set Legenda = Doc.CreateTable(param_table3)
				Legenda.Font = Font
				y_legenda=y_medias-altura_medias
				'response.Write(altura_medias)
				'response.end()
				With Legenda.Rows(1)
				   .Cells(1).Width = 40
				   .Cells(2).Width = 20
				   .Cells(3).Width = 40
				   .Cells(4).Width = 20
				   .Cells(5).Width = 40
				   .Cells(6).Width = 20
				   .Cells(7).Width = 40
				   .Cells(8).Width = 20 
				   .Cells(9).Width = 43 
				   .Cells(10).Width = 250             
				End With
				data_exibe = data&" &agrave;s "& horario

				Legenda(1, 1).Colspan= 8	


 		Set RSdt = Server.CreateObject("ADODB.Recordset")
		SQLdt = "SELECT * FROM TB_OBS_Ficha where CO_Matricula="&cod_cons
		RSdt.Open SQLdt, CONt

qtd_maxima = 794

		IF RSdt.EOF then
			ln1=""
			ln2=""
		else
			ln=RSdt("TX_OBS_LIN")
			txt = Len(ln)
			espaco=InStrRev(left(ln,79)," ")
			ln1=left(ln,espaco)
			ln2=mid(ln,espaco+1,718)
			
		end if  
		if ano_letivo = 2020 then
			ln = "Ano letivo regido pela Delibera&ccedil;&atilde;o CEE/RJ 376/2020, Lei Federal 14.040/2020, Lei Estadual 8.991/2020, Delibera&ccedil;&atilde;o CEE/RJ 384/2020, Parecer CEE/RJ 36/2020 e Resolu&ccedil;&atilde;o SEEDUC 5841/2020 (autoriza pr&aacute;tica das aulas online no per&iacute;odo da pandemia e dispensa a obrigatoriedade de observ&acirc;ncia do m&iacute;nimo de dias letivos)."
			espaco=InStrRev(left(ln,79)," ")
			ln1=left(ln,espaco)
			ln2=mid(ln,espaco+1,718)
		end if
				'response.write(txt)			
				'response.end()
				Legenda(1, 10).AddText "<b><Div align=""right"">CARGA HOR&Aacute;RIA TOTAL DE "&horas_aula&" HORAS</div></b>", "size=8; html=true", Font 
				Legenda(2, 10).AddText "<b><Div align=""right"">Documento impresso em: "&data_exibe&"</div></b>", "size=8; html=true", Font 				
				Page.Canvas.DrawTable Legenda, "x=30, y="&y_legenda&"" 
				
				
					 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width=50; alignment=left; size=8; color=#000000")

					Relatorio = "SWD048"
					Do While Len(Relatorio) > 0
						CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
					 
						If CharsPrinted = Len(Relatorio) Then Exit Do
						   SET Page = Page.NextPage
						Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
					Loop 
					
				y_obs_linha1=129
					With Page.Canvas
					   .MoveTo 100, y_obs_linha1
					   .LineTo 570, y_obs_linha1
					   .Stroke
					End With 	
				
				y_obs_linha2=y_obs_linha1-10
					With Page.Canvas
					   .MoveTo margem, y_obs_linha2
					   .LineTo 570, y_obs_linha2
					   .Stroke
					End With 
					
				y_obs_linha3=y_obs_linha2-10
					With Page.Canvas
					   .MoveTo margem, y_obs_linha3
					   .LineTo 570, y_obs_linha3
					   .Stroke
					End With 
					
				y_obs_linha4=y_obs_linha3-10
					With Page.Canvas
					   .MoveTo margem, y_obs_linha4
					   .LineTo 570, y_obs_linha4
					   .Stroke
					End With 
					
				y_obs_linha5=y_obs_linha4-10
					With Page.Canvas
					   .MoveTo margem, y_obs_linha5
					   .LineTo 570, y_obs_linha5
					   .Stroke
					End With 	
			 if txt<=qtd_maxima then				
					Observ = "O aluno obteve "&frequencia&"% de frequ&ecirc;ncia no ano letivo. "&ln1
					margem_obs=margem*3.3
					SET Param_Obs = Pdf.CreateParam("x="&margem_obs&";y=140; height=50; width=500; alignment=left; size=8; color=#000000;html=true")

					Do While Len(Observ) > 0
						CharsPrinted = Page.Canvas.DrawText(Observ, Param_Obs, Font )
					 
						If CharsPrinted = Len(Observ) Then Exit Do
						   SET Page = Page.NextPage
						Observ = Right(Observ, Len(Observ) - CharsPrinted)
					Loop 
					
					Observ = ln2
					margem_obs=margem*1.2
					SET Param_Obs = Pdf.CreateParam("x="&margem_obs&";y=128; height=100; width=535; alignment=left; size=8; color=#000000;html=true")

					Do While Len(Observ) > 0
						CharsPrinted = Page.Canvas.DrawText(Observ, Param_Obs, Font )
					 
						If CharsPrinted = Len(Observ) Then Exit Do
						   SET Page = Page.NextPage
						Observ = Right(Observ, Len(Observ) - CharsPrinted)
					Loop 	
	
			else			
					
					Observ = "Favor verificar folha anexa"
					margem_obs=margem*3.3
					SET Param_Obs = Pdf.CreateParam("x="&margem_obs&";y=140; height=50; width=500; alignment=left; size=8; color=#000000;html=true")

					Do While Len(Observ) > 0
						CharsPrinted = Page.Canvas.DrawText(Observ, Param_Obs, Font )
					 
						If CharsPrinted = Len(Observ) Then Exit Do
						   SET Page = Page.NextPage
						Observ = Right(Observ, Len(Observ) - CharsPrinted)
					Loop 				
					
			end if									

'INÍCIO DO RECIBO==================================================================================



				if 	wrk_exibe_resultado="s" THEN
'					resultados=Calc_Med_An_Fin(unidade, curso, etapa, turma, cod_cons, vetor_materia, caminho_nota, tb_nota, 4, 4, 0, "final", 0)		

					if tipo_resultado="A" then
						resultados=resultados1						
'response.Write("----A <br>")							
					else'
						resultados=resultados2		
'response.Write("----B <br>")														
					end if	
					resultados_apurados = split(resultados, "#%#" )	
					
					'response.Write(tipo_resultado&" -r")
					'response.end()					
					if ubound(resultados_apurados)=-1 then
						resultado_final_aluno="nulo"
					else				
						'resultado_final_aluno=apura_resultado_aluno(curso, etapa, resultados_apurados(0))
						 resultado_final_aluno=apura_resultado_geral_aluno(curso, etapa, resultados_apurados(0),tipo_resultado)		
					'response.Write(curso&", "&etapa&", "&resultados_apurados(0)&", "&tipo_resultado&" -r")
					'response.Write("<BR>"&resultado_final_aluno)
					'response.end()								 				
					end if
	
				
					
					if (resultado_final_aluno="Apr" or resultado_final_aluno="APR") then
					'	legenda_resultado="<b>Resultado:</b>"
						resultado_exibe="Aprovado"
					elseif (resultado_final_aluno="Rep" or resultado_final_aluno="REP") then	
					'	legenda_resultado="<b>Resultado:</b>"				
						resultado_exibe="Reprovado"	
					elseif (resultado_final_aluno="Rec" or resultado_final_aluno="REC") then	
						'legenda_resultado="<b>Resultado:</b>"					
						resultado_exibe="Recupera&ccedil;&atilde;o"	
					else
						'legenda_resultado="<b>Resultado:</b>"					
						resultado_exibe=""											
					end if	
				else	
						'legenda_resultado=""					
						resultado_exibe=""				
				end if	


				Set param_table4 = Pdf.CreateParam("width=540; height=125; rows=3; cols=1; border=1; cellborder=1; cellspacing=0;")
				Set quadro = Doc.CreateTable(param_table4)
				quadro.Font = Font
				quadro.Rows(1).Height = 20
				quadro.Rows(2).Height = 65
				quadro.Rows(3).Height = 40
         

				quadro(1, 1).AddText "<b>Resultado Final:</b>&nbsp;&nbsp;"&resultado_exibe, "size=9;indentx=2; html=true", Font 
				quadro(2, 1).AddText "<b>Observa&ccedil;&otilde;es:</b>", "size=9;indentx=5;indenty=5;html=true;", Font 
'				Legenda(1, 5).AddText "<b>R</b> = Recupera&ccedil;&atilde;o", "size=9;html=true", Font 
'				Legenda(1, 7).AddText "<b>Md</b> = M&eacute;dia", "size=9; html=true", Font 
'				Legenda(1, 10).AddText "<b><Div align=""right"">Documento impresso em: "&data_exibe&"</div></b>", "size=8; html=true", Font 
				Page.Canvas.DrawTable quadro, "x="&margem&", y=165" 

			Page.Canvas.SetParams "LineWidth=1" 
			
			
															
		'Data===========================================================================
			 SET Param_data = Pdf.CreateParam("x=473;y=70; height=40; width=100; alignment=Left; size=8; color=#000000")
			data_preenche = "Data:_____/_____/_____"
			CharsPrinted = Page.Canvas.DrawText(data_preenche, Param_data, Font )
		
			
		'===========================================================================
			 SET Param_secretario = Pdf.CreateParam("x=40;y=50; height=40; width=100; alignment=Left; size=9; color=#000000;html=true")
			secretario = "<b>Secretario(a)</b>"
			CharsPrinted = Page.Canvas.DrawText(secretario, Param_secretario, Font )

			 SET Param_diretor = Pdf.CreateParam("x=390;y=50; height=40; width=100; alignment=Left; size=9; color=#000000;html=true")
			diretor = "<b>Diretor(a</b>)"
			CharsPrinted = Page.Canvas.DrawText(diretor, Param_diretor, Font )
 			
			
'Página de observações--------------------------------------------------------------			
			if txt>qtd_maxima then			
				SET Page = Page.NextPage
						
	'Novo CABEÇALHO==========================================================================================		
				Set Param_Logo_Gde = Pdf.CreateParam
				margem=30				
				largura_logo_gde=formatnumber(Logo.Width*0.4,0)
				altura_logo_gde=formatnumber(Logo.Height*0.4,0)
				area_utilizavel=Page.Width - (margem*2)
				Param_Logo_Gde("x") = margem
				Param_Logo_Gde("y") = Page.Height - altura_logo_gde -22
				Param_Logo_Gde("ScaleX") = 0.4
				Param_Logo_Gde("ScaleY") = 0.4
				Page.Canvas.DrawImage Logo, Param_Logo_Gde
		
				x_texto=largura_logo_gde+ margem+10
				'x_texto= margem
				y_texto=formatnumber(Page.Height - margem,0)
				'width_texto=Page.Width - 60
				width_texto=Page.Width -7*margem
			
																
	
				SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde+30&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
				Text = "<p><i><b>"&UCASE(txt1)&"</b></i></p><br><font style=""font-size:10pt;"">"&txt2&"<BR>"&txt3&"<BR>"&txt4&"<BR>"&txt5&"<BR>"&txt6&"</FONT>" 			
	
				
				Do While Len(Text) > 0
					CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
				 
					If CharsPrinted = Len(Text) Then Exit Do
						SET Page = Page.NextPage
					Text = Right( Text, Len(Text) - CharsPrinted)
				Loop 
	
				y_texto=y_texto-altura_logo_gde-10
				SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height=30; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
				Text = "<center><i><b><font style=""font-size:18pt;"">"&documento&" - Observa&ccedil;&otilde;es</font></b></i></center>"
				
				
				Do While Len(Text) > 0
					CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
				 
					If CharsPrinted = Len(Text) Then Exit Do
						SET Page = Page.NextPage
					Text = Right( Text, Len(Text) - CharsPrinted)
				Loop 	
				
	'================================================================================================================			
				Page.Canvas.SetParams "LineWidth=2" 
				Page.Canvas.SetParams "LineCap=0" 
				altura_primeiro_separador= Page.Height - altura_logo_gde-margem
				With Page.Canvas
				   .MoveTo margem, altura_primeiro_separador
				   .LineTo Page.Width - margem, altura_primeiro_separador
				   .Stroke
				End With 
								
				Page.Canvas.SetParams "LineWidth=2" 
				Page.Canvas.SetParams "LineCap=0" 
				altura_segundo_separador= Page.Height - altura_logo_gde -margem - 45
				With Page.Canvas
				   .MoveTo margem, altura_segundo_separador
				   .LineTo area_utilizavel+margem, altura_segundo_separador
				   .Stroke
				End With 	
								
				Set param_table1 = Pdf.CreateParam("width="&area_utilizavel&"; height=72; rows=6; cols=8; border=0; cellborder=0; cellspacing=0;")
				Set Table = Doc.CreateTable(param_table1)
				Table.Font = Font
				y_primeira_tabela=altura_segundo_separador-10
				x_primeira_tabela=margem+5
				With Table.Rows(1)
				   .Cells(1).Width = 50
				   .Cells(2).Width = 15  
				   .Cells(3).Width = 20	
				   .Cells(4).Width = 70	
				   .Cells(5).Width = 20	
				   .Cells(6).Width = 50	
				   .Cells(7).Width = 90				   		   		   
				End With
				
				Table(1, 2).ColSpan = 3
				Table(1, 5).ColSpan = 2	
				Table(2, 2).ColSpan = 7
				Table(3, 2).ColSpan = 5
				Table(4, 2).ColSpan = 5
				Table(5, 3).ColSpan = 2	
				Table(5, 5).ColSpan = 3													
				Table(6, 2).ColSpan = 2
				Table(6, 4).ColSpan = 3	
				
				Table(1, 1).AddText "Matr&iacute;cula:", "size=8;html=true", Font 
				Table(1, 2).AddText "<b>"&cod_cons&"</b>", "size=8;html=true", Font 
				Table(1, 5).AddText "ID-INEP:", "size=8;html=true", Font 
				Table(1, 7).AddText "<b>"&id_unico&"</b>", "size=8;html=true", Font 				
				Table(2, 1).AddText "Alun"&desinencia&":", "size=8;", Font 
				Table(2, 2).AddText "<b>"&nome_aluno&"</b>", "size=8;html=true", Font 
				Table(3, 1).AddText "Pai: ", "size=8;", Font 
				Table(3, 2).AddText nome_pai, "size=8;html=true", Font 
				Table(4, 1).AddText "M&atilde;e: ", "size=8;html=true", Font 
				Table(4, 2).AddText nome_mae, "size=8;html=true", Font	
				Table(5, 1).AddText "Sexo: ", "size=8;", Font 
				Table(5, 2).AddText sexo_aluno, "size=8;", Font 
				Table(5, 3).AddText " Nascimento: "&nascimento, "size=8;", Font 							
				Table(5, 5).AddText " Natural de: "&no_uf, "size=8;", Font		
				Table(6, 1).AddText "Ano Letivo: ", "size=8;html=true", Font 
				Table(6, 2).AddText ano_letivo, "size=8;", Font	
				Table(6, 4).AddText "Curso: "& no_curso, "size=8;", Font
				Table(6, 7).AddText texto_turma, "size=8;html=true", Font												
				Table(6, 8).AddText texto_chamada, "size=8;html=true", Font	
				Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 		
	
				if curso=1 and co_etapa<6 then
					sem_carga="s"
					cols=12
					width=280			
				else
					sem_carga="n"
					cols=13		
					width=250						
				end if	
				
				Set param_table4 = Pdf.CreateParam("width=540; height=562; rows=1; cols=1; border=1; cellborder=1; cellspacing=0;")
				Set observacoes = Doc.CreateTable(param_table4)		
				Page.Canvas.DrawTable observacoes, "x=30, y="&y_medias&"" 											
			

				y_obs_linha=y_medias - 10
				Page.Canvas.SetParams "LineWidth=0.1"			
				while y_obs_linha> 40
					With Page.Canvas
					   .MoveTo margem, y_obs_linha
					   .LineTo 570, y_obs_linha
					   .Stroke
					End With 	
					y_obs_linha=y_obs_linha-10				
				wend
				
				altura_obs = y_medias

		if ano_letivo = 2020 then
			ln = "Ano letivo regido pela Delibera&ccedil;&atilde;o CEE/RJ 376/2020, Lei Federal 14.040/2020, Lei Estadual 8.991/2020, Delibera&ccedil;&atilde;o CEE/RJ 384/2020, Parecer CEE/RJ 36/2020 e Resolu&ccedil;&atilde;o SEEDUC 5841/2020 (autoriza pr&aacute;tica das aulas online no per&iacute;odo da pandemia e dispensa a obrigatoriedade de observ&acirc;ncia do m&iacute;nimo de dias letivos)."
			espaco=InStrRev(left(ln,79)," ")
			ln1=left(ln,espaco)
			ln2=mid(ln,espaco+1,718)
		end if
		
					Observ = "O aluno obteve "&frequencia&"% de frequ&ecirc;ncia no ano letivo. "&ln1
					margem_obs=margem*3.3
					SET Param_Obs = Pdf.CreateParam("x="&margem_obs&";y=140; height=50; width=500; alignment=left; size=8; color=#000000;html=true")

					Do While Len(Observ) > 0
						CharsPrinted = Page.Canvas.DrawText(Observ, Param_Obs, Font )
					 
						If CharsPrinted = Len(Observ) Then Exit Do
						   SET Page = Page.NextPage
						Observ = Right(Observ, Len(Observ) - CharsPrinted)
					Loop 
					
					Observ = ln2
					margem_obs=margem*1.2
					SET Param_Obs = Pdf.CreateParam("x="&margem_obs&";y=128; height=100; width=535; alignment=left; size=8; color=#000000;html=true")

					Do While Len(Observ) > 0
						CharsPrinted = Page.Canvas.DrawText(Observ, Param_Obs, Font )
					 
						If CharsPrinted = Len(Observ) Then Exit Do
						   SET Page = Page.NextPage
						Observ = Right(Observ, Len(Observ) - CharsPrinted)
					Loop 		
				
					'Observ = "O aluno obteve "&frequencia&"% de frequ&ecirc;ncia no ano letivo."
					'margem_obs=margem*1.2
					'SET Param_Obs = Pdf.CreateParam("x="&margem_obs&";y="&altura_obs&"; height=562; width=500; alignment=left; size=10; color=#000000;html=true")

					'Do While Len(Relatorio) > 0
					'	CharsPrinted = Page.Canvas.DrawText(Observ, Param_Obs, Font )
					'
					'	If CharsPrinted = Len(Observ) Then Exit Do
					'	   SET Page = Page.NextPage
					'	Observ = Right(Observ, Len(Observ) - CharsPrinted)
					'Loop 
					
					'Observ = ln
					'altura_obs = altura_obs - 20					
					'SET Param_Obs = Pdf.CreateParam("x="&margem_obs&";y="&altura_obs&"; height=562; width=535; alignment=left; size=10; color=#000000;")

					Do While Len(Observ) > 0
						CharsPrinted = Page.Canvas.DrawText(Observ, Param_Obs, Font )
					 
						If CharsPrinted = Len(Observ) Then Exit Do
						   SET Page = Page.NextPage
						Observ = Right(Observ, Len(Observ) - CharsPrinted)
					Loop 	
					
 					SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width=50; alignment=left; size=8; color=#000000")

					Relatorio = "SWD048"
					Do While Len(Relatorio) > 0
						CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
					 
						If CharsPrinted = Len(Relatorio) Then Exit Do
						   SET Page = Page.NextPage
						Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
					Loop 					
			end if					
			
		End IF	
	END IF 		
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

arquivo="SWD048"
Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

