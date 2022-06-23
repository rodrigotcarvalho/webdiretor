<%'On Error Resume Next%>
<%
Server.ScriptTimeout =60 'valor em segundos
'BOLETIM ESCOLAR
%>
<!--#include file="../../../global/funcoes_diversas.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
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
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now)

'quantidade de informações apresentadas por período
inf_por_periodo=1 

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

	
if opt="01" then
	cod_cons=request.QueryString("cod_cons")
	periodo_form=request.QueryString("prd")	
	
	If Not IsArray(alunos_encontrados) Then alunos_encontrados = Array() End if	
	ReDim preserve alunos_encontrados(UBound(alunos_encontrados)+1)	
	alunos_encontrados(Ubound(alunos_encontrados)) = cod_cons
	
elseif opt="02" then

	obr=request.QueryString("obr")
	dados_informados = split(obr, "_" )

	unidade=dados_informados(0)
	curso=dados_informados(1)
	co_etapa=dados_informados(2)
	turma=dados_informados(3)
	periodo_form=dados_informados(4)	
	
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

max_notas_exibe=inf_por_periodo*periodo_form
' o menos 1 é por que o vetor de notas começa no zero
'inf_por_periodo=inf_por_periodo*1
'max_notas_exibe=max_notas_exibe*1
'max_notas_exibe=max_notas_exibe
'response.Write(max_notas_exibe)
'response.end()
	
For i=0 to ubound(alunos_encontrados)	
	cod_cons=alunos_encontrados(i)
		
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
					response.Write("ERRO - N&atilde;o cadastrado TP_Nota em TB_Da_Aula para NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& etapa &"' AND CO_Turma ='"& turma &"'" )
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
					else
						response.Write("ERRO - N&atilde;o cadastrado TP_Nota em TB_Da_Aula para NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& etapa &"' AND CO_Turma ='"& turma &"'" )
						response.end()
					end if	
				end if
			
				call GeraNomes("PORT",unidade,curso,etapa,CON0)
				no_unidade = session("no_unidades")
				no_curso= session("no_grau")
				no_etapa = session("no_serie")
						
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
			margem=25			
			area_utilizavel=Page.Width - (margem*2)
			largura_logo_gde=formatnumber(Logo.Width*0.8,0)
			altura_logo_gde=formatnumber(Logo.Height*0.8,0)
	
		   Param_Logo_Gde("x") = margem
		   Param_Logo_Gde("y") = Page.Height - altura_logo_gde - margem
		   Param_Logo_Gde("ScaleX") = 0.8
		   Param_Logo_Gde("ScaleY") = 0.8
		   Page.Canvas.DrawImage Logo, Param_Logo_Gde
	
			x_texto=margem
			y_texto=formatnumber(Page.Height - margem - 20,0)
			width_texto=Page.Width - (2*margem)

		
			SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
			Text = "<p><center><i><b>Col&eacute;gio Stockler</b></i></center></p><br><p><center><i><b><font style=""font-size:18pt;"">Boletim Escolar</font></b></i></center></p>"

			Do While Len(Text) > 0
				CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )	 
				If CharsPrinted = Len(Text) Then Exit Do
					SET Page = Page.NextPage
				Text = Right( Text, Len(Text) - CharsPrinted)
			Loop 
		
'================================================================================================================			
			y_nome_aluno=Page.Height - altura_logo_gde-46
			width_nome_aluno=Page.Width - (2*margem)
			
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
			   .Cells(2).Width = 105
			   .Cells(3).Width = 25
			   .Cells(4).Width = 100
			   .Cells(5).Width = 60
			   .Cells(6).Width = 120
			   .Cells(7).Width = 50
			   .Cells(8).Width = 45      
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
		
		
			Set RS5 = Server.CreateObject("ADODB.Recordset")
			SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' AND IN_MAE=TRUE order by NU_Ordem_Boletim "
			RS5.Open SQL5, CON0
			co_materia_check=1
			IF RS5.EOF Then
				vetor_materia_exibe="nulo"
			else
				while not RS5.EOF
					co_mat_fil= RS5("CO_Materia")				
					if co_materia_check=1 then
						vetor_materia=co_mat_fil
					else
						vetor_materia=vetor_materia&"#!#"&co_mat_fil
					end if
					co_materia_check=co_materia_check+1			
							
				RS5.MOVENEXT
				wend	
			'response.Write(vetor_materia&"OK1")
				'vetor_materia_exibe=programa_aula(vetor_materia, unidade, curso, etapa, turma)	
				'response.Write(vetor_materia_exibe&"OK")						
			end if
			'response.end()	
			co_materia_exibe=Split(vetor_materia,"#!#")		
			
			width_notas=595-(2*margem)
			
			if tb_nota="TB_NOTA_A" then			
				altura_medias=50
				Set param_table2 = Pdf.CreateParam("width="&width_notas&"; height="&altura_medias&"; rows=3; cols=13; border=1; cellborder=0.1; cellspacing=0;")
				Set Notas_Tit = Doc.CreateTable(param_table2)
				Notas_Tit.Font = Font
				y_medias=Page.Height - altura_logo_gde-110
				
				width_celulas=formatnumber((width_notas-140-(3*28)-25)/8,0)
	
				With Notas_Tit.Rows(1)
				   .Cells(1).Width = 140
				   .Cells(2).Width = width_celulas
				   .Cells(3).Width = width_celulas
'				   .Cells(4).Width = 28
'				   .Cells(5).Width = 28
'				   .Cells(6).Width = width_celulas
'				   .Cells(7).Width = width_celulas
				   .Cells(4).Width = width_celulas   
				   .Cells(5).Width = width_celulas
'				   .Cells(10).Width = 28
'				   .Cells(11).Width = 28
'				   .Cells(12).Width = width_celulas
'				   .Cells(13).Width = width_celulas
				   .Cells(6).Width = 28
				   .Cells(7).Width = 28  
				   .Cells(8).Width = 28
				   .Cells(9).Width = 25     
				   .Cells(10).Width = width_celulas
				   .Cells(11).Width = width_celulas
				   .Cells(12).Width = width_celulas  
				   .Cells(13).Width = width_celulas			             
				End With
				Notas_Tit(1, 1).RowSpan = 3
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
				Notas_Tit(2, 13).RowSpan = 2
'				Notas_Tit(2, 14).RowSpan = 2
'				Notas_Tit(2, 15).RowSpan = 2
'				Notas_Tit(2, 16).RowSpan = 2
'				Notas_Tit(2, 17).RowSpan = 2
'				Notas_Tit(2, 18).RowSpan = 2
'				Notas_Tit(2, 19).RowSpan = 2	
'				Notas_Tit(2, 20).RowSpan = 2
'				Notas_Tit(2, 21).RowSpan = 2																																													
				'Notas_Tit(2, 14).RowSpan = 2
				'Notas_Tit(1, 17).RowSpan = 2
				Notas_Tit(1, 2).ColSpan = 4
				Notas_Tit(1, 6).ColSpan = 4
				Notas_Tit(1, 10).ColSpan = 4
				'Notas_Tit(1, 15).ColSpan = 2
				Notas_Tit(1, 1).AddText "<div align=""center""><b>Disciplinas</b></div>", "size=10;indenty=15; html=true", Font 
				Notas_Tit(1, 2).AddText "<div align=""center""><b>Aproveitamento</b></div>", "size=8;indenty=2; html=true", Font 
				Notas_Tit(1, 6).AddText "<div align=""center""><b>Resultado</b></div>", "size=8;indenty=2; html=true", Font 				
				Notas_Tit(1, 10).AddText "<div align=""center""><b>Freq&uuml;&ecirc;ncia (faltas)</b></div>", "size=8;indenty=2; html=true", Font
				 	
				Notas_Tit(2, 2).AddText "B1", "size=9;alignment=center;indenty=8;", Font 
				Notas_Tit(2, 3).AddText "B2", "size=9;alignment=center;indenty=8;", Font 
'				Notas_Tit(2, 4).AddText "<div align=""center"">Md<BR>Sem1</div>", "size=9;alignment=center;indenty=4; html=true", Font 
'				Notas_Tit(2, 5).AddText "<div align=""center"">Rec<BR>Par</div>", "size=9;alignment=center;indenty=4; html=true", Font 
'				Notas_Tit(2, 6).AddText "B1*", "size=9;alignment=center;indenty=8;", Font 
'				Notas_Tit(2, 7).AddText "B2*", "size=9;alignment=center;indenty=8;", Font 
				Notas_Tit(2, 4).AddText "B3", "size=9;alignment=center;indenty=8;", Font 
				Notas_Tit(2, 5).AddText "B4", "size=9;alignment=center;indenty=8;", Font 
'				Notas_Tit(2, 10).AddText "<div align=""center"">Md<BR>Sem2</div>", "size=9;alignment=center;indenty=4; html=true", Font 
'				Notas_Tit(2, 11).AddText "<div align=""center"">Rec<BR>Par</div>", "size=9;alignment=center;indenty=4; html=true", Font 
'				Notas_Tit(2, 12).AddText "B3*", "size=9;alignment=center;indenty=8;", Font 
'				Notas_Tit(2, 13).AddText "B4*", "size=9;alignment=center;indenty=8;", Font 
				Notas_Tit(2, 6).AddText "<div align=""center"">Md<BR>Anual</div>", "size=9;alignment=center;indenty=4; html=true", Font 
				Notas_Tit(2, 7).AddText "<div align=""center"">Pr<BR>Final</div>", "size=9;alignment=center;indenty=4; html=true", Font 
				Notas_Tit(2, 8).AddText "<div align=""center"">Md<BR>Final</div>", "size=9;alignment=center;indenty=4; html=true", Font 		
				Notas_Tit(2, 9).AddText "RES", "size=9;alignment=center;indenty=8;", Font 	
				Notas_Tit(2, 10).AddText "B1", "size=9;alignment=center;indenty=8;", Font 	
				Notas_Tit(2, 11).AddText "B2", "size=9;alignment=center;indenty=8;", Font 	
				Notas_Tit(2, 12).AddText "B3", "size=9;alignment=center;indenty=8;", Font 																	
				Notas_Tit(2, 13).AddText "B4", "size=9;alignment=center;indenty=8;", Font 
												
				Set param_materias = PDF.CreateParam	
				param_materias.Set "expand=true" 
'	response.Write (unidade&"% "&curso&"% "&etapa&"% "&turma&"% "&vetor_periodo&"% "&cod_cons&"% "&vetor_materia&"% "& caminho_nota&"% "&tb_nota&"% "&nome_nota&"% boletim<br>")	
	'response.end()	
				medias=calcula_medias(unidade, curso, etapa, turma, vetor_periodo, cod_cons, vetor_materia, caminho_nota, tb_nota, nome_nota, "boletim")

'		response.Write (medias)

				medias_materia = split(medias,"#$#")
						
				qtd_medias_materia = ubound(medias_materia)	
				
				for k=0 to qtd_medias_materia
					co_materia_consulta=co_materia_exibe(k)
'				response.Write (k&"="&co_materia_consulta&"<br>")
					if 	co_materia_consulta<>"MED" then
						call GeraNomes(co_materia_consulta,unidade,curso,etapa,CON0)
						no_materia_exibe=session("no_materia")	
					else
						no_materia_exibe="M&eacute;dia"
					end if
						
					'For j=0 to ubound(co_materia_exibe)	
						posicao_materia=posicao_materia_tabela(co_materia_consulta, unidade, curso, etapa, turma)	
						
						if posicao_materia=0 then
						
						elseif posicao_materia=1 then
							param_materias.Add "indentx=2"
						elseif posicao_materia=2 then
							param_materias.Add "indentx=15"
						elseif posicao_materia=3 then
							param_materias.Add "indentx=15"
							no_materia_exibe="--> "&no_materia_exibe		
						end if
					'Next			
						
						
					Set Row = Notas_Tit.Rows.Add(20) ' row height
					param_materias.Add "indenty=3;alignment=right;html=true"
					Row.Cells(1).AddText "<font style=""font-size:9pt;"">"&no_materia_exibe&"</font>", param_materias
					altura_medias=altura_medias+20
		'response.Write ("<BR>'"&medias_materia(k)&"'")					
					notas_exibe = split(medias_materia(k),"#!#")	
		'response.end()			
					param_materias.Add "indentx=0"	
	
	'					for m=0 to ubound(notas_exibe)
	'		'response.Write ("<BR>v"&notas_exibe(m)&"v")	
	'						if notas_exibe(m)="" or isnull(notas_exibe(m)) then
	'							nota=" "
	'						end if
	'						nota=notas_exibe(m)
	'						indice=m+1
	'			
	'						Row.Cells(indice).AddText "<div align=""center""><font style=""font-size:8pt;"">"&nota&"</font></div>", param_materias
	'					next
						'response.Write("<BR>")
					if notas_exibe(0)="" or notas_exibe(0)="&nbsp;" or isnull(notas_exibe(0)) then
						notas_exibe(0)=" "
					else
						notas_exibe(0)=formatnumber(notas_exibe(0),1)							
					end if
					
					if notas_exibe(1)="" or notas_exibe(1)="&nbsp;" or isnull(notas_exibe(1)) or (verifica_periodos="s" and max_notas_exibe<2) then
						notas_exibe(1)=" "
					else
						notas_exibe(1)=formatnumber(notas_exibe(1),1)							
					end if
					
					if notas_exibe(2)="" or notas_exibe(2)="&nbsp;" or isnull(notas_exibe(2)) then
						notas_exibe(2)=" "
					else
						notas_exibe(2)=formatnumber(notas_exibe(2),1)							
					end if
					
					if notas_exibe(3)="" or notas_exibe(3)="&nbsp;" or isnull(notas_exibe(3)) then
						notas_exibe(3)=" "
					else
						notas_exibe(3)=formatnumber(notas_exibe(3),1)							
					end if
					
					if notas_exibe(4)="" or notas_exibe(4)="&nbsp;" or isnull(notas_exibe(4)) then
						notas_exibe(4)=" "
					else
						notas_exibe(4)=formatnumber(notas_exibe(4),1)							
					end if
					
					if notas_exibe(5)="" or notas_exibe(5)="&nbsp;" or isnull(notas_exibe(5)) then
						notas_exibe(5)=" "
					else
						notas_exibe(5)=formatnumber(notas_exibe(5),1)							
					end if
					
					if notas_exibe(6)="" or notas_exibe(6)="&nbsp;" or isnull(notas_exibe(6)) or (verifica_periodos="s" and max_notas_exibe<3) then
						notas_exibe(6)=" "
					else
						notas_exibe(6)=formatnumber(notas_exibe(6),1)							
					end if																																				
	
					if notas_exibe(7)="" or notas_exibe(7)="&nbsp;" or isnull(notas_exibe(7)) or (verifica_periodos="s" and max_notas_exibe<4) then
						notas_exibe(7)=" "
					else
						notas_exibe(7)=formatnumber(notas_exibe(7),1)							
					end if
					
					if notas_exibe(8)="" or notas_exibe(8)="&nbsp;" or isnull(notas_exibe(8)) then
						notas_exibe(8)=" "
					else
						notas_exibe(8)=formatnumber(notas_exibe(8),1)							
					end if
					
					if notas_exibe(9)="" or notas_exibe(9)="&nbsp;" or isnull(notas_exibe(9)) then
						notas_exibe(9)=" "
					else
						notas_exibe(9)=formatnumber(notas_exibe(9),1)							
					end if
					
					if notas_exibe(10)="" or notas_exibe(10)="&nbsp;" or isnull(notas_exibe(10)) then
						notas_exibe(10)=" "
					else
						notas_exibe(10)=formatnumber(notas_exibe(10),1)							
					end if	
					
					if notas_exibe(11)="" or notas_exibe(11)="&nbsp;" or isnull(notas_exibe(11)) then
						notas_exibe(11)=" "
					else
						notas_exibe(11)=formatnumber(notas_exibe(11),1)							
					end if
					
					if notas_exibe(13)="" or notas_exibe(13)="&nbsp;" or isnull(notas_exibe(13)) then
						notas_exibe(13)=" "
					else
						notas_exibe(13)=notas_exibe(13)							
					end if
					
					if notas_exibe(14)="" or notas_exibe(14)="&nbsp;" or isnull(notas_exibe(14)) or (verifica_periodos="s" and max_notas_exibe<2) then
						notas_exibe(14)=" "
					else
						notas_exibe(14)=notas_exibe(14)					
					end if
					
					if notas_exibe(15)="" or notas_exibe(15)="&nbsp;" or isnull(notas_exibe(15)) or (verifica_periodos="s" and max_notas_exibe<3) then
						notas_exibe(15)=" "
					else
						notas_exibe(15)=notas_exibe(15)						
					end if
										
					if notas_exibe(16)="" or notas_exibe(16)="&nbsp;" or isnull(notas_exibe(16)) or (verifica_periodos="s" and max_notas_exibe<4) then
						notas_exibe(16)=" "
					else
						notas_exibe(16)=notas_exibe(16)						
					end if												
					
					if notas_exibe(17)="" or notas_exibe(17)="&nbsp;" or isnull(notas_exibe(17))  or (verifica_periodos="s" and max_notas_exibe<4) then
						notas_exibe(17)=" "
					else
						notas_exibe(17)=formatnumber(notas_exibe(17),1)							
					end if
					
					if notas_exibe(18)="" or notas_exibe(18)="&nbsp;" or isnull(notas_exibe(18))  or (verifica_periodos="s" and max_notas_exibe<5) then
						notas_exibe(18)=" "
					else
						notas_exibe(18)=formatnumber(notas_exibe(18),1)							
					end if
					
					if notas_exibe(19)="" or notas_exibe(19)="&nbsp;" or isnull(notas_exibe(19))  or (verifica_periodos="s" and max_notas_exibe<5) then
						notas_exibe(19)=" "
					else
						notas_exibe(19)=formatnumber(notas_exibe(19),1)							
					end if												
					
					if notas_exibe(20)="" or notas_exibe(20)="&nbsp;" or isnull(notas_exibe(20)) or (verifica_periodos="s" and max_notas_exibe<6) then
						notas_exibe(20)=" "
					else
						notas_exibe(20)=notas_exibe(20)
					end if
					
																
					Row.Cells(2).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(0)&"</font></div>", param_materias
					Row.Cells(3).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(1)&"</font></div>", param_materias
					Row.Cells(4).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(6)&"</font></div>", param_materias
					Row.Cells(5).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(7)&"</font></div>", param_materias
					Row.Cells(6).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(17)&"</font></div>", param_materias	
					Row.Cells(7).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(18)&"</font></div>", param_materias
					Row.Cells(8).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(19)&"</font></div>", param_materias
					Row.Cells(9).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(20)&"</font></div>", param_materias
					Row.Cells(10).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(13)&"</font></div>", param_materias
					Row.Cells(11).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(14)&"</font></div>", param_materias		
					Row.Cells(12).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(15)&"</font></div>", param_materias
					Row.Cells(13).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(16)&"</font></div>", param_materias																																												
				next
				Page.Canvas.DrawTable Notas_Tit, "x="&margem&", y="&y_medias&"" 		
	
	'LINHAS QUE DIVIDEM OS PERÍODOS DA TABELA========================================================================
	'				rows_notas=Ubound(co_materia_exibe)+3
	'				rows_notas=rows_notas*1
	'				altura_linha_divisora_notas=(rows_notas*20)
	'				y_fim_linha=y_medias-altura_linha_divisora_notas
	'				
	'				Page.Canvas.SetParams "LineWidth=1" 
	'				x_primeira_linha=140+30
	'					With Page.Canvas
	'					   .MoveTo x_primeira_linha, y_medias
	'					   .LineTo x_primeira_linha, y_fim_linha
	'					   .Stroke
	'					End With 
	'					
	'				x_segunda_linha=x_primeira_linha+(19*4)+1
	'					With Page.Canvas
	'					   .MoveTo x_segunda_linha, y_medias
	'					   .LineTo x_segunda_linha, y_fim_linha
	'					   .Stroke
	'					End With 
	'					
	'				x_terceira_linha=x_segunda_linha+(19*4)
	'					With Page.Canvas
	'					   .MoveTo x_terceira_linha, y_medias
	'					   .LineTo x_terceira_linha, y_fim_linha
	'					   .Stroke
	'					End With 	
	'					
	'				x_quarta_linha=x_terceira_linha+(19*4)
	'					With Page.Canvas
	'					   .MoveTo x_quarta_linha, y_medias
	'					   .LineTo x_quarta_linha, y_fim_linha
	'					   .Stroke
	'					End With 	
	'					
	'				x_quinta_linha=x_quarta_linha+75
	'					With Page.Canvas
	'					   .MoveTo x_quinta_linha, y_medias
	'					   .LineTo x_quinta_linha, y_fim_linha
	'					   .Stroke
	'					End With 	
	'					
	'				x_sexta_linha=x_quinta_linha+(25*2)
	'					With Page.Canvas
	'					   .MoveTo x_sexta_linha, y_medias
	'					   .LineTo x_sexta_linha, y_fim_linha
	'					   .Stroke
	'					End With 				
				'=============================================================================================================
			
				Set param_table3 = Pdf.CreateParam("width="&width_notas&"; height=20; rows=1; cols=13; border=0; cellborder=0; cellspacing=0;")
				Set Legenda = Doc.CreateTable(param_table3)
				Legenda.Font = Font
				y_legenda=y_medias-altura_medias
				'response.Write(altura_medias)
				'response.end()
				With Legenda.Rows(1)
				   .Cells(1).Width = 32
				   .Cells(2).Width = 3
				   .Cells(3).Width = 75
				   .Cells(4).Width = 3
				   .Cells(5).Width = 40
				   .Cells(6).Width = 3
				   .Cells(7).Width = 40 
'				   .Cells(8).Width = 3 
'				   .Cells(9).Width = 40
'				   .Cells(10).Width = 3 
'				   .Cells(11).Width = 40
'				   .Cells(12).Width = 3 					    					   					    
				   .Cells(13).Width = 140             
				End With
				data_exibe = data&" &agrave;s "& horario
				Legenda(1, 1).AddText "<b>Legenda:</b> ", "size=7;html=true", Font 
				Legenda(1, 3).AddText "<b>B</b>= M&eacute;dias Bimestrais", "size=7;html=true", Font 				
				Legenda(1, 5).AddText "<b>Md</b>= M&eacute;dia", "size=7;html=true", Font 
				Legenda(1, 7).AddText "<b>Pr</b>= Prova", "size=7;html=true", Font 					
				Legenda(1, 13).AddText "<b><Div align=""right"">Documento impresso em: "&data_exibe&"</div></b>", "size=6; html=true", Font 
				Page.Canvas.DrawTable Legenda, "x="&margem&", y="&y_legenda&"" 


			
			
			else			
				altura_medias=50
				Set param_table2 = Pdf.CreateParam("width="&width_notas&"; height="&altura_medias&"; rows=3; cols=21; border=1; cellborder=0.1; cellspacing=0;")
				Set Notas_Tit = Doc.CreateTable(param_table2)
				Notas_Tit.Font = Font
				y_medias=Page.Height - altura_logo_gde-110
				
				width_celulas=formatnumber((width_notas-100-(7*28)-25)/12,0)
	
				With Notas_Tit.Rows(1)
				   .Cells(1).Width = 100
				   .Cells(2).Width = width_celulas
				   .Cells(3).Width = width_celulas
				   .Cells(4).Width = 28
				   .Cells(5).Width = 28
				   .Cells(6).Width = width_celulas
				   .Cells(7).Width = width_celulas
				   .Cells(8).Width = width_celulas   
				   .Cells(9).Width = width_celulas
				   .Cells(10).Width = 28
				   .Cells(11).Width = 28
				   .Cells(12).Width = width_celulas
				   .Cells(13).Width = width_celulas
				   .Cells(14).Width = 28
				   .Cells(15).Width = 28  
				   .Cells(16).Width = 28
				   .Cells(17).Width = 25     
				   .Cells(18).Width = width_celulas
				   .Cells(19).Width = width_celulas
				   .Cells(20).Width = width_celulas  
				   .Cells(21).Width = width_celulas			             
				End With
				Notas_Tit(1, 1).RowSpan = 3
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
				Notas_Tit(2, 13).RowSpan = 2
				Notas_Tit(2, 14).RowSpan = 2
				Notas_Tit(2, 15).RowSpan = 2
				Notas_Tit(2, 16).RowSpan = 2
				Notas_Tit(2, 17).RowSpan = 2
				Notas_Tit(2, 18).RowSpan = 2
				Notas_Tit(2, 19).RowSpan = 2	
				Notas_Tit(2, 20).RowSpan = 2
				Notas_Tit(2, 21).RowSpan = 2																																													
				'Notas_Tit(2, 14).RowSpan = 2
				'Notas_Tit(1, 17).RowSpan = 2
				Notas_Tit(1, 2).ColSpan = 13
				Notas_Tit(1, 14).ColSpan = 4
				Notas_Tit(1, 18).ColSpan = 4
				'Notas_Tit(1, 15).ColSpan = 2
				Notas_Tit(1, 1).AddText "<div align=""center""><b>Disciplinas</b></div>", "size=10;indenty=15; html=true", Font 
				Notas_Tit(1, 2).AddText "<div align=""center""><b>Aproveitamento</b></div>", "size=8;indenty=2; html=true", Font 
				Notas_Tit(1, 14).AddText "<div align=""center""><b>Resultado</b></div>", "size=8;indenty=2; html=true", Font 				
				Notas_Tit(1, 18).AddText "<div align=""center""><b>Freq&uuml;&ecirc;ncia (faltas)</b></div>", "size=8;indenty=2; html=true", Font 	
				'Notas_Tit(1, 2).AddText "<div align=""center""><b>1&ordm; Per&iacute;odo</b></div>", "size=10;alignment=center; indenty=5;html=true", Font 
				'Notas_Tit(1, 6).AddText "<div align=""center""><b>2&ordm; Per&iacute;odo</b></div>", "size=10;alignment=center; indenty=5;html=true", Font 
				'Notas_Tit(1, 10).AddText "<div align=""center""><b>3&ordm; Per&iacute;odo</b></div>", "size=10;alignment=center; indenty=5;html=true", Font 
				'Notas_Tit(1, 14).AddText "<div align=""center""><b>M&eacute;dia<br>Anual</b></div>", "size=10;alignment=center; indenty=5;html=true", Font 
				'Notas_Tit(1, 15).AddText "<div align=""center""><b>Rec.</b></div>", "size=10;alignment=center; indenty=5;html=true", Font 
				'Notas_Tit(1, 17).AddText "<div align=""center""><b>Final</b></div>","size=10;alignment=center;indenty=15;html=true", Font 
				Notas_Tit(2, 2).AddText "B1", "size=9;alignment=center;indenty=8;", Font 
				Notas_Tit(2, 3).AddText "B2", "size=9;alignment=center;indenty=8;", Font 
				Notas_Tit(2, 4).AddText "<div align=""center"">Md<BR>Sem1</div>", "size=9;alignment=center;indenty=4; html=true", Font 
				Notas_Tit(2, 5).AddText "<div align=""center"">Rec<BR>Par</div>", "size=9;alignment=center;indenty=4; html=true", Font 
				Notas_Tit(2, 6).AddText "B1*", "size=9;alignment=center;indenty=8;", Font 
				Notas_Tit(2, 7).AddText "B2*", "size=9;alignment=center;indenty=8;", Font 
				Notas_Tit(2, 8).AddText "B3", "size=9;alignment=center;indenty=8;", Font 
				Notas_Tit(2, 9).AddText "B4", "size=9;alignment=center;indenty=8;", Font 
				Notas_Tit(2, 10).AddText "<div align=""center"">Md<BR>Sem2</div>", "size=9;alignment=center;indenty=4; html=true", Font 
				Notas_Tit(2, 11).AddText "<div align=""center"">Rec<BR>Par</div>", "size=9;alignment=center;indenty=4; html=true", Font 
				Notas_Tit(2, 12).AddText "B3*", "size=9;alignment=center;indenty=8;", Font 
				Notas_Tit(2, 13).AddText "B4*", "size=9;alignment=center;indenty=8;", Font 
				Notas_Tit(2, 14).AddText "<div align=""center"">Md<BR>Anual</div>", "size=9;alignment=center;indenty=4; html=true", Font 
				Notas_Tit(2, 15).AddText "<div align=""center"">Pr<BR>Final</div>", "size=9;alignment=center;indenty=4; html=true", Font 
				Notas_Tit(2, 16).AddText "<div align=""center"">Md<BR>Final</div>", "size=9;alignment=center;indenty=4; html=true", Font 		
				Notas_Tit(2, 17).AddText "RES", "size=9;alignment=center;indenty=8;", Font 	
				Notas_Tit(2, 18).AddText "B1", "size=9;alignment=center;indenty=8;", Font 	
				Notas_Tit(2, 19).AddText "B2", "size=9;alignment=center;indenty=8;", Font 	
				Notas_Tit(2, 20).AddText "B3", "size=9;alignment=center;indenty=8;", Font 																	
				Notas_Tit(2, 21).AddText "B4", "size=9;alignment=center;indenty=8;", Font 
												
				Set param_materias = PDF.CreateParam	
				param_materias.Set "expand=true" 
	'response.Write (unidade&"% "&curso&"% "&etapa&"% "&turma&"% "&vetor_periodo&"% "&cod_cons&"% "&vetor_materia&"% "& caminho_nota&"% "&tb_nota&"% "&nome_nota&"% boletim<br>")	
	'response.end()	
				medias=calcula_medias(unidade, curso, etapa, turma, vetor_periodo, cod_cons, vetor_materia, caminho_nota, tb_nota, nome_nota, "boletim")
'		response.Write (medias&" - "&verifica_periodos&"=s and "&max_notas_exibe)
'	response.end()
		
				medias_materia = split(medias,"#$#")
						
				qtd_medias_materia = ubound(medias_materia)	
				
				for k=0 to qtd_medias_materia
					co_materia_consulta=co_materia_exibe(k)
				
					if 	co_materia_consulta<>"MED" then
						call GeraNomes(co_materia_consulta,unidade,curso,etapa,CON0)
						no_materia_exibe=session("no_materia")	
					else
						no_materia_exibe="M&eacute;dia"
					end if
						
					'For j=0 to ubound(co_materia_exibe)	
						posicao_materia=posicao_materia_tabela(co_materia_consulta, unidade, curso, etapa, turma)	
						
						if posicao_materia=0 then
						
						elseif posicao_materia=1 then
							param_materias.Add "indentx=2"
						elseif posicao_materia=2 then
							param_materias.Add "indentx=15"
						elseif posicao_materia=3 then
							param_materias.Add "indentx=15"
							no_materia_exibe="--> "&no_materia_exibe		
						end if
					'Next			
						
						
					Set Row = Notas_Tit.Rows.Add(20) ' row height
					param_materias.Add "indenty=3;alignment=right;html=true"
					Row.Cells(1).AddText "<font style=""font-size:9pt;"">"&no_materia_exibe&"</font>", param_materias
					altura_medias=altura_medias+20
		'response.Write ("<BR>'"&medias_materia(k)&"'")					
					notas_exibe = split(medias_materia(k),"#!#")	
		'response.end()			
					param_materias.Add "indentx=0"	
	
	'					for m=0 to ubound(notas_exibe)
	'		'response.Write ("<BR>v"&notas_exibe(m)&"v")	
	'						if notas_exibe(m)="" or isnull(notas_exibe(m)) then
	'							nota=" "
	'						end if
	'						nota=notas_exibe(m)
	'						indice=m+1
	'			
	'						Row.Cells(indice).AddText "<div align=""center""><font style=""font-size:8pt;"">"&nota&"</font></div>", param_materias
	'					next
						'response.Write(">"&notas_exibe(5))
						'response.end()
					if notas_exibe(0)="" or notas_exibe(0)="&nbsp;" or isnull(notas_exibe(0)) then
						notas_exibe(0)=" "
					else
						notas_exibe(0)=formatnumber(notas_exibe(0),1)							
					end if
					
					if notas_exibe(1)="" or notas_exibe(1)="&nbsp;" or isnull(notas_exibe(1)) or (verifica_periodos="s" and max_notas_exibe<2) then
						notas_exibe(1)=" "
					else
						notas_exibe(1)=formatnumber(notas_exibe(1),1)							
					end if
					
					if notas_exibe(2)="" or notas_exibe(2)="&nbsp;" or isnull(notas_exibe(2)) or (verifica_periodos="s" and max_notas_exibe<2) then
						notas_exibe(2)=" "
					else
						notas_exibe(2)=formatnumber(notas_exibe(2),1)							
					end if
					
					if notas_exibe(3)="" or notas_exibe(3)="&nbsp;" or isnull(notas_exibe(3)) or (verifica_periodos="s" and max_notas_exibe<2) then
						notas_exibe(3)=" "
					else
						notas_exibe(3)=formatnumber(notas_exibe(3),1)							
					end if
					
					if notas_exibe(4)="" or notas_exibe(4)="&nbsp;" or isnull(notas_exibe(4)) or (verifica_periodos="s" and max_notas_exibe<2) then
						notas_exibe(4)=" "
					else
						notas_exibe(4)=formatnumber(notas_exibe(4),1)							
					end if
					
					if notas_exibe(5)="" or notas_exibe(5)="&nbsp;" or isnull(notas_exibe(5)) or (verifica_periodos="s" and max_notas_exibe<2) then
						notas_exibe(5)=" "
					else
						notas_exibe(5)=formatnumber(notas_exibe(5),1)							
					end if
					
					if notas_exibe(6)="" or notas_exibe(6)="&nbsp;" or isnull(notas_exibe(6)) or (verifica_periodos="s" and max_notas_exibe<3) then
						notas_exibe(6)=" "
					else
						notas_exibe(6)=formatnumber(notas_exibe(6),1)							
					end if																																				
	
					if notas_exibe(7)="" or notas_exibe(7)="&nbsp;" or isnull(notas_exibe(7)) or (verifica_periodos="s" and max_notas_exibe<4) then
						notas_exibe(7)=" "
					else
						notas_exibe(7)=formatnumber(notas_exibe(7),1)							
					end if
					
					if notas_exibe(8)="" or notas_exibe(8)="&nbsp;" or isnull(notas_exibe(8)) or (verifica_periodos="s" and max_notas_exibe<4) then
						notas_exibe(8)=" "
					else
						notas_exibe(8)=formatnumber(notas_exibe(8),1)							
					end if
					
					if notas_exibe(9)="" or notas_exibe(9)="&nbsp;" or isnull(notas_exibe(9)) or (verifica_periodos="s" and max_notas_exibe<4) then
						notas_exibe(9)=" "
					else
						notas_exibe(9)=formatnumber(notas_exibe(9),1)							
					end if
					
					if notas_exibe(10)="" or notas_exibe(10)="&nbsp;" or isnull(notas_exibe(10)) or (verifica_periodos="s" and max_notas_exibe<4) then
						notas_exibe(10)=" "
					else
						notas_exibe(10)=formatnumber(notas_exibe(10),1)							
					end if	
					
					if notas_exibe(11)="" or notas_exibe(11)="&nbsp;" or isnull(notas_exibe(11)) or (verifica_periodos="s" and max_notas_exibe<4)then
						notas_exibe(11)=" "
					else
						notas_exibe(11)=formatnumber(notas_exibe(11),1)							
					end if
					
					if notas_exibe(13)="" or notas_exibe(13)="&nbsp;" or isnull(notas_exibe(13)) then
						notas_exibe(13)=" "
					else
						notas_exibe(13)=notas_exibe(13)							
					end if
					
					if notas_exibe(14)="" or notas_exibe(14)="&nbsp;" or isnull(notas_exibe(14)) or (verifica_periodos="s" and max_notas_exibe<2) then
						notas_exibe(14)=" "
					else
						notas_exibe(14)=notas_exibe(14)					
					end if
					
					if notas_exibe(15)="" or notas_exibe(15)="&nbsp;" or isnull(notas_exibe(15)) or (verifica_periodos="s" and max_notas_exibe<3) then
						notas_exibe(15)=" "
					else
						notas_exibe(15)=notas_exibe(15)						
					end if
										
					if notas_exibe(16)="" or notas_exibe(16)="&nbsp;" or isnull(notas_exibe(16)) or (verifica_periodos="s" and max_notas_exibe<4) then
						notas_exibe(16)=" "
					else
						notas_exibe(16)=notas_exibe(16)						
					end if												
					
					if notas_exibe(17)="" or notas_exibe(17)="&nbsp;" or isnull(notas_exibe(17)) or (verifica_periodos="s" and max_notas_exibe<4) then
						notas_exibe(17)=" "
					else
						notas_exibe(17)=formatnumber(notas_exibe(17),1)							
					end if
					
					if notas_exibe(18)="" or notas_exibe(18)="&nbsp;" or isnull(notas_exibe(18)) or (verifica_periodos="s" and max_notas_exibe<5) then
						notas_exibe(18)=" "
					else
						notas_exibe(18)=formatnumber(notas_exibe(18),1)							
					end if
					
					if notas_exibe(19)="" or notas_exibe(19)="&nbsp;" or isnull(notas_exibe(19)) or (verifica_periodos="s" and max_notas_exibe<5) then
						notas_exibe(19)=" "
					else
						notas_exibe(19)=formatnumber(notas_exibe(19),1)							
					end if												
					
					if notas_exibe(20)="" or notas_exibe(20)="&nbsp;" or isnull(notas_exibe(20)) or (verifica_periodos="s" and max_notas_exibe<4) then
						notas_exibe(20)=" "
					else
						notas_exibe(20)=notas_exibe(20)
					end if
					
																
					Row.Cells(2).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(0)&"</font></div>", param_materias
					Row.Cells(3).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(1)&"</font></div>", param_materias
					Row.Cells(4).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(2)&"</font></div>", param_materias
					Row.Cells(5).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(3)&"</font></div>", param_materias
					Row.Cells(6).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(4)&"</font></div>", param_materias	
					Row.Cells(7).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(5)&"</font></div>", param_materias
					Row.Cells(8).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(6)&"</font></div>", param_materias
					Row.Cells(9).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(7)&"</font></div>", param_materias
					Row.Cells(10).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(8)&"</font></div>", param_materias
					Row.Cells(11).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(9)&"</font></div>", param_materias		
					Row.Cells(12).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(10)&"</font></div>", param_materias
					Row.Cells(13).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(11)&"</font></div>", param_materias
					Row.Cells(14).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(17)&"</font></div>", param_materias
					Row.Cells(15).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(18)&"</font></div>", param_materias
					Row.Cells(16).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(19)&"</font></div>", param_materias
					Row.Cells(17).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(20)&"</font></div>", param_materias
					Row.Cells(18).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(13)&"</font></div>", param_materias
					Row.Cells(19).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(14)&"</font></div>", param_materias
					Row.Cells(20).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(15)&"</font></div>", param_materias
					Row.Cells(21).AddText "<div align=""center""><font style=""font-size:8pt;"">"&notas_exibe(16)&"</font></div>", param_materias																																													
						
				next
				Page.Canvas.DrawTable Notas_Tit, "x="&margem&", y="&y_medias&"" 		
	
	'LINHAS QUE DIVIDEM OS PERÍODOS DA TABELA========================================================================
	'				rows_notas=Ubound(co_materia_exibe)+3
	'				rows_notas=rows_notas*1
	'				altura_linha_divisora_notas=(rows_notas*20)
	'				y_fim_linha=y_medias-altura_linha_divisora_notas
	'				
	'				Page.Canvas.SetParams "LineWidth=1" 
	'				x_primeira_linha=140+30
	'					With Page.Canvas
	'					   .MoveTo x_primeira_linha, y_medias
	'					   .LineTo x_primeira_linha, y_fim_linha
	'					   .Stroke
	'					End With 
	'					
	'				x_segunda_linha=x_primeira_linha+(19*4)+1
	'					With Page.Canvas
	'					   .MoveTo x_segunda_linha, y_medias
	'					   .LineTo x_segunda_linha, y_fim_linha
	'					   .Stroke
	'					End With 
	'					
	'				x_terceira_linha=x_segunda_linha+(19*4)
	'					With Page.Canvas
	'					   .MoveTo x_terceira_linha, y_medias
	'					   .LineTo x_terceira_linha, y_fim_linha
	'					   .Stroke
	'					End With 	
	'					
	'				x_quarta_linha=x_terceira_linha+(19*4)
	'					With Page.Canvas
	'					   .MoveTo x_quarta_linha, y_medias
	'					   .LineTo x_quarta_linha, y_fim_linha
	'					   .Stroke
	'					End With 	
	'					
	'				x_quinta_linha=x_quarta_linha+75
	'					With Page.Canvas
	'					   .MoveTo x_quinta_linha, y_medias
	'					   .LineTo x_quinta_linha, y_fim_linha
	'					   .Stroke
	'					End With 	
	'					
	'				x_sexta_linha=x_quinta_linha+(25*2)
	'					With Page.Canvas
	'					   .MoveTo x_sexta_linha, y_medias
	'					   .LineTo x_sexta_linha, y_fim_linha
	'					   .Stroke
	'					End With 				
				'=============================================================================================================
			
				Set param_table3 = Pdf.CreateParam("width="&width_notas&"; height=20; rows=1; cols=13; border=0; cellborder=0; cellspacing=0;")
				Set Legenda = Doc.CreateTable(param_table3)
				Legenda.Font = Font
				y_legenda=y_medias-altura_medias
				'response.Write(altura_medias)
				'response.end()
				With Legenda.Rows(1)
				   .Cells(1).Width = 32
				   .Cells(2).Width = 3
				   .Cells(3).Width = 75
				   .Cells(4).Width = 3
				   .Cells(5).Width = 95
				   .Cells(6).Width = 3
				   .Cells(7).Width = 110 
				   .Cells(8).Width = 3 
				   .Cells(9).Width = 40
				   .Cells(10).Width = 3 
				   .Cells(11).Width = 40
				   .Cells(12).Width = 3 					    					   					    
				   .Cells(13).Width = 140             
				End With
				data_exibe = data&" &agrave;s "& horario
				Legenda(1, 1).AddText "<b>Legenda:</b> ", "size=7;html=true", Font 
				Legenda(1, 3).AddText "<b>B</b>= M&eacute;dias Bimestrais", "size=7;html=true", Font 
				Legenda(1, 5).AddText "<b>B*</b>= B corrigidas por Rec Par", "size=7;html=true", Font 					
				Legenda(1, 7).AddText "<b>Rec Par</b> = Recupera&ccedil;&atilde;o Paralela", "size=7;html=true", Font 
				Legenda(1, 9).AddText "<b>Md</b> = M&eacute;dia", "size=7; html=true", Font 
				Legenda(1, 11).AddText "<b>Pr</b> = Prova", "size=7; html=true", Font 					
				Legenda(1, 13).AddText "<b><Div align=""right"">Documento impresso em: "&data_exibe&"</div></b>", "size=6; html=true", Font 
				Page.Canvas.DrawTable Legenda, "x="&margem&", y="&y_legenda&"" 
			end if									

				
	
			SET Param_Relatorio = Pdf.CreateParam("x=30;y=140; height=50; width=50; alignment=left; size=8; color=#000000")
		'	Relatorio = "Sistema Web Diretor - SWD025"
			Relatorio = "SWD025"
			Do While Len(Relatorio) > 0
				CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
			 
				If CharsPrinted = Len(Relatorio) Then Exit Do
				   SET Page = Page.NextPage
				Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
			Loop 
				
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
			Page.Canvas.SetParams "LineWidth=0.5" 
			With Page.Canvas
			   .MoveTo 330, 30
			   .LineTo Page.Width - 30, 30
			   .Stroke
			End With 
		
		'Data===========================================================================
			 SET Param_data = Pdf.CreateParam("x="&margem&";y=40; height=50; width=100; alignment=Left; size=8; color=#000000")
			data_preenche = "Data:_____/_____/_____"
			CharsPrinted = Page.Canvas.DrawText(data_preenche, Param_data, Font )
		'===========================================================================
		
			Page.Canvas.SetParams "Dash1=2; DashPhase=1"
			Page.Canvas.SetParams "LineWidth=0.7" 
			With Page.Canvas
			   .MoveTo 30, 125
			   .LineTo Page.Width - 30, 125
			   .Stroke
			End With 
			

			 SET Param_Tesoura = Pdf.CreateParam("x="&margem&";y=136; height=50; width="&area_utilizavel&"; alignment=center; size=16; color=#000000, html=true")
			Tesoura = "<div align=""center"">&quot;</div>"
			
			Do While Len(Tesoura) > 0
				CharsPrinted = Page.Canvas.DrawText(Tesoura, Param_Tesoura, Font_Tesoura )
			 
				If CharsPrinted = Len(Tesoura) Then Exit Do
				   SET Page = Page.NextPage
				Tesoura = Right( Tesoura, Len(Tesoura) - CharsPrinted)
			Loop 
		
			largura_logo_pqno=formatnumber(Logo.Width,0)
			altura_logo_pqno=formatnumber(Logo.Height,0)
			
			Set Param_Logo_Pqno = Pdf.CreateParam
			   Param_Logo_Pqno("x") = margem 
			   Param_Logo_Pqno("y") = altura_logo_pqno
			   Param_Logo_Pqno("ScaleX") = 0.3
			   Param_Logo_Pqno("ScaleY") = 0.3
			   Page.Canvas.DrawImage Logo, Param_Logo_Pqno
			
			x_texto_recibo=margem
			y_texto_recibo=altura_logo_pqno+margem
			width_texto=Page.Width -largura_logo_pqno - 100
			SET Param_recibo = Pdf.CreateParam("x="&x_texto_recibo&";y="&y_texto_recibo&"; height="&altura_logo_pqno&"; width="&area_utilizavel&"; alignment=center; size=14; color=#000000; html=true")
			Text_recibo = "<p><center><i><b>Col&eacute;gio Stockler</b></i></center></p>"
	
			Do While Len(Text_recibo) > 0
				CharsPrinted = Page.Canvas.DrawText(Text_recibo, Param_recibo, Font )
			 
				If CharsPrinted = Len(Text_recibo) Then Exit Do
					SET Page = Page.NextPage
				Text_recibo = Right( Text_recibo, Len(Text_recibo) - CharsPrinted)
			Loop 

			Set param_recibo2 = Pdf.CreateParam("width=80; height=40; rows=2; cols=2; border=0; cellborder=0; cellspacing=0;")

			Set recibo_2 = Doc.CreateTable(param_recibo2)

			y_recibo_2=y_texto_recibo-(formatnumber(altura_logo_pqno*0.3,0))
			y_recibo_tb=y_texto_recibo
			
			With recibo_2.Rows(1)
			   .Cells(1).Width = 50
			   .Cells(2).Width = 30
			End With
			recibo_2(1, 1).AddText "<Div align=""center"">Ano Letivo:</div>", "size=8;indenty=2;html=true", Font 
			recibo_2(1, 2).AddText "<Div align=""Right""><b>"&ano_letivo&"</b></div>", "size=9;indenty=2;html=true", Font 
			recibo_2(2, 1).AddText "<Div align=""center"">Matr&iacute;cula:</div>", "size=8;indenty=2;html=true", Font 
			recibo_2(2, 2).AddText "<Div align=""Right""><b>"&cod_cons&"</b></div>", "size=8;indenty=2;html=true", Font 
			Page.Canvas.DrawTable recibo_2, "x=485, y="&y_recibo_tb&"" 
		
			
			
			SET Param_Nome_Aluno_Recibo = Pdf.CreateParam("x="&margem&";y="&y_recibo_2&"; height=50; width="&width_nome_aluno&"; alignment=left; html=True")
			Nome_Recibo = "<font style=""font-size:10pt;""><b>"&nome_aluno&"</b></font>"
			Do While Len(Nome_Recibo) > 0
				CharsPrinted = Page.Canvas.DrawText(Nome_Recibo, Param_Nome_Aluno_Recibo, Font )
			 
				If CharsPrinted = Len(Nome_Recibo) Then Exit Do
					SET Page = Page.NextPage
				Nome_Recibo = Right( Nome_Recibo, Len(Nome_Recibo) - CharsPrinted)
			Loop 

			y_recibo_3=y_recibo_2-20
			
			SET Param_Dados_Aluno_Recibo = Pdf.CreateParam("x="&margem&";y="&y_recibo_3&"; height=50; width="&width_nome_aluno&"; alignment=left; html=True")
			'Dados_Recibo = "<font style=""font-size:8pt;"">Curso: <b>"&no_curso &"</b> S&eacute;rie: <b>"&no_etapa &"</b> Turma: <b>"& turma &"</b></font>"
			Dados_Recibo = "<font style=""font-size:8pt;"">Curso: <b>"&no_curso &"</b> Turma: <b>"& turma &"</b></font>"				
			Do While Len(Dados_Recibo) > 0
				CharsPrinted = Page.Canvas.DrawText(Dados_Recibo, Param_Dados_Aluno_Recibo, Font )
			 
				If CharsPrinted = Len(Dados_Recibo) Then Exit Do
					SET Page = Page.NextPage
				Dados_Recibo = Right( Dados_Recibo, Len(Dados_Recibo) - CharsPrinted)
			Loop 
			
			y_recibo_4=y_recibo_tb-20

			width_texto=Page.Width - 120
			SET Param_recibo3 = Pdf.CreateParam("x="&margem&";y="&y_recibo_4&"; height=30; width="&area_utilizavel&"; alignment=center; size=12; color=#000000; html=true")
			Text_recibo3 = "<BR><BR><BR><BR><BR><p><center><b><font style=""font-size:7pt;"">Devolver essa parte assinada ao col&eacute;gio</font><b></center></p>"
			
			Do While Len(Text_recibo3) > 0
				CharsPrinted = Page.Canvas.DrawText(Text_recibo3, Param_recibo3, Font )
			 
				If CharsPrinted = Len(Text_recibo3) Then Exit Do
					SET Page = Page.NextPage
				Text_recibo3 = Right( Text_recibo3, Len(Text_recibo3) - CharsPrinted)
			Loop
			
			SET Param_recibo4 = Pdf.CreateParam("x=230;y=40; height=60; width=300; alignment=left; size=8; color=#000000; html=true")
			Responsavel_Recibo = "<font style=""font-size:8pt;"">Assinatura do Respons&aacute;vel:</font>"
			Do While Len(Responsavel_Recibo) > 0
				CharsPrinted = Page.Canvas.DrawText(Responsavel_Recibo, Param_recibo4, Font )
			 
				If CharsPrinted = Len(Responsavel_Recibo) Then Exit Do
					SET Page = Page.NextPage
				Responsavel_Recibo = Right( Responsavel_Recibo, Len(Responsavel_Recibo) - CharsPrinted)
			Loop
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

arquivo="SWD025"
Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

