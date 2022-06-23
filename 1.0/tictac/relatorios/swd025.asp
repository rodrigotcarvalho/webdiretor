<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 60 'valor em segundos
'BOLETIM ESCOLAR
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
	in_lingua=RS("IN_Lingua")	
	nome_aluno=replace_latin_char(nome_aluno,"html")	
	
	if sexo_aluno="F" then
		desinencia="a"
	else
		desinencia="o"
	end if
	
	select case in_lingua
		case "D"
			lingua_est = "Ingl&ecirc;s"
		case "I"
			lingua_est = "Espanhol"
		case "V"
			lingua_est = "Franc&ecirc;s"
		case else
			lingua_est="ERRO"
	end select	
	
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
			tb_nota=tabela_nota(ano_letivo,unidade,curso,co_etapa,turma,"tb",0)
			caminho_nota=tabela_nota(ano_letivo,unidade,curso,co_etapa,turma,"cam",0)
						
			no_unidade= GeraNomes("U",unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)
			no_curso=GeraNomes("C",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro) 	
			no_etapa=GeraNomes("E",curso,co_etapa,variavel3,variavel4,variavel5,CON0,outro) 		
			tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")
			tp_freq=tipo_divisao_ano(curso,co_etapa,"in_frequencia")
			cols_tit_vetor=dados_boletim(tp_modelo,tp_freq,1,"tit",tb_nota)	
			cols_tit_rowspan_vetor=dados_boletim(tp_modelo,tp_freq,1,"pdf_rowspan",tb_nota)
			cols_tit_colspan_vetor=dados_boletim(tp_modelo,tp_freq,1,"colspan",tb_nota)
			cols_nome_notas_vetor=dados_boletim(tp_modelo,tp_freq,2,"pdf_tit",tb_nota)	
			cols_notas_periodo_vetor=dados_boletim(tp_modelo,tp_freq,2,"periodo_ref",tb_nota)
			cols_notas_rowspan_vetor=dados_boletim(tp_modelo,tp_freq,2,"pdf_rowspan",tb_nota)
			cols_notas_colspan_vetor=dados_boletim(tp_modelo,tp_freq,2,"colspan",tb_nota)
			cols_notas_calc_vetor=dados_boletim(tp_modelo,tp_freq,2,"tipo_calc",tb_nota)
		
			colunas_tit=split(cols_tit_vetor,"#!#")
			colunas_tit_rowspan=split(cols_tit_rowspan_vetor,"#!#")
			colunas_tit_colspan=split(cols_tit_colspan_vetor,"#!#")
			colunas_nome_notas=split(cols_nome_notas_vetor,"#!#")
			colunas_notas_periodo=split(cols_notas_periodo_vetor,"#!#")
			cols_notas_rowspan=split(cols_notas_rowspan_vetor,"#!#")
			cols_notas_colspan=split(cols_notas_colspan_vetor,"#!#")
			colunas_notas_calc=split(cols_notas_calc_vetor,"#!#")
					
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
			Text = "<p><center><i><b><font style=""font-size:18pt;"">Boletim Escolar</font></b></i></center></p>"
			

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

total_colunas=0			
for tc=0 to ubound(colunas_tit)
	coluna_afetada=coluna_afetada*1
	colunas_tit_colspan(tc)=colunas_tit_colspan(tc)*1
	total_colunas=total_colunas+colunas_tit_colspan(tc)							
next


			altura_medias=40
			largura_tabela=553
			Set param_table2 = Pdf.CreateParam("width="&largura_tabela&"; height="&altura_medias&"; rows=3; cols="&total_colunas&"; border=1; cellborder=0.1; cellspacing=0;")
			Set Notas_Tit = Doc.CreateTable(param_table2)
			Notas_Tit.Font = Font
			y_medias=Page.Height - altura_logo_gde-110
			larg_col=formatnumber((largura_tabela-130)/total_colunas,2)
			
			With Notas_Tit.Rows(1)
			   .Cells(1).Width = 130
				for wd=2 to total_colunas
				   .Cells(wd).Width = larg_col
				next				  		         			   			         
			End With

coluna_afetada=1
for ctr=0 to ubound(colunas_tit)
	coluna_afetada=coluna_afetada*1
	colunas_tit_colspan(ctr)=colunas_tit_colspan(ctr)*1
	
	colunas_tit_rowspan(ctr)=colunas_tit_rowspan(ctr)*1
	if colunas_tit_rowspan(ctr)	> 1 then
		Notas_Tit(1, coluna_afetada).RowSpan = colunas_tit_rowspan(ctr)		
	end if	
	coluna_afetada=coluna_afetada+colunas_tit_colspan(ctr)						
next

coluna_afetada=1
for ctc=0 to ubound(colunas_tit)
	coluna_afetada=coluna_afetada*1
	colunas_tit_colspan(ctc)=colunas_tit_colspan(ctc)*1
	if colunas_tit_colspan(ctc) > 1 then	
		Notas_Tit(1, coluna_afetada).ColSpan = colunas_tit_colspan(ctc)	
	end if			
	coluna_afetada=coluna_afetada+colunas_tit_colspan(ctc)
						
next

coluna_afetada=2
for cnnr=0 to ubound(colunas_nome_notas)
	coluna_afetada=coluna_afetada*1
	cols_notas_colspan(cnnr)=cols_notas_colspan(cnnr)*1

	cols_notas_rowspan(cnnr)=cols_notas_rowspan(cnnr)*1
	if cols_notas_rowspan(cnnr)	> 1 then	
		Notas_Tit(2, coluna_afetada).RowSpan = cols_notas_rowspan(cnnr)	
	end if			
	coluna_afetada=coluna_afetada+cols_notas_colspan(cnnr)				
next

coluna_afetada=2
for cnnc=0 to ubound(colunas_nome_notas)
	coluna_afetada=coluna_afetada*1
	cols_notas_colspan(cnnc)=cols_notas_colspan(cnnc)*1
	
	if cols_notas_colspan(cnnc) > 1 then		
		Notas_Tit(2, coluna_afetada).ColSpan = cols_notas_colspan(cnnc)		
	end if	
	coluna_afetada=coluna_afetada+cols_notas_colspan(cnnc)							
next

coluna_afetada=1
for ct=0 to ubound(colunas_tit)
	coluna_afetada=coluna_afetada*1
	cols_notas_colspan(ct)=cols_notas_colspan(ct)*1

	if ct=0 then
		indenty=14
		fontsize=10		
	else
		indenty=1
		fontsize=6			
	end if	
	Notas_Tit(1, coluna_afetada).AddText "<div align=""center""><b>"&colunas_tit(ct)&"</b></div>", "size="&fontsize&";indenty="&indenty&"; html=true", Font 

	coluna_afetada=coluna_afetada+colunas_tit_colspan(ct)
next

coluna_afetada=2
for cnn=0 to ubound(colunas_nome_notas)
	coluna_afetada=coluna_afetada*1
	cols_notas_colspan(cnn)=cols_notas_colspan(cnn)*1
	
	if Len(colunas_nome_notas(cnn)) > 6 then
		indenty=4
	else
		indenty=8	
	end if

	Notas_Tit(2, coluna_afetada).AddText "<div align=""center"">"&colunas_nome_notas(cnn)&"</div>", "size="&fontsize&";alignment=center;indenty="&indenty&"; html=true", Font 
	coluna_afetada=coluna_afetada+cols_notas_colspan(cnn)
next
'response.End()

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
'			Notas_Tit(2, 16).RowSpan = 2	
'			Notas_Tit(2, 17).RowSpan = 2
'			Notas_Tit(1, 2).ColSpan = 8	
'			Notas_Tit(1, 10).ColSpan = 4				
'			Notas_Tit(1, 14).ColSpan = 4																																																																						
'			Notas_Tit(1, 1).AddText "<div align=""center""><b>Disciplinas</b></div>", "size=10;indenty=15; html=true", Font 
'			Notas_Tit(1, 2).AddText "<div align=""center""><b>Aproveitamento</b></div>", "size=8;alignment=center; indenty=1;html=true", Font  
'			Notas_Tit(1, 9).AddText "<div align=""center""><b>Resultado</b></div>", "size=8;alignment=center; indenty=1;html=true", Font  
'			Notas_Tit(1, 14).AddText "<div align=""center""><b>Frequ&ecirc;ncia</b></div>", "size=8;alignment=center; indenty=1;html=true", Font
'			Notas_Tit(2, 2).AddText "BIM1", "size=7;alignment=center;indenty=6;", Font 
'			Notas_Tit(2, 3).AddText "BIM2", "size=7 ;alignment=center;indenty=6;", Font 
'			Notas_Tit(2, 4).AddText "<div align=""center"">Md<br>Sem</div>", "size=7;alignment=center;indenty=2;html=true", Font 
'			Notas_Tit(2, 5).AddText "<div align=""center"">Rec<br>Par</div>", "size=7;alignment=center;indenty=2;html=true", Font 
'			Notas_Tit(2, 6).AddText "BIM1*", "size=7;alignment=center;indenty=6", Font 
'			Notas_Tit(2, 7).AddText "BIM2*", "size=7 ;alignment=center;indenty=6;", Font 
'			Notas_Tit(2, 8).AddText "BIM3", "size=7;alignment=center;indenty=6", Font 
'			Notas_Tit(2, 9).AddText "BIM4", "size=7;alignment=center;indenty=6", Font 
'			Notas_Tit(2, 10).AddText "<div align=""center"">M&eacute;dia<br>Anual</div>", "size=7;alignment=center;indenty=2; html=true", Font 
'			Notas_Tit(2, 11).AddText "<div align=""center"">Prova<br>Recup</div>", "size=7;alignment=center;indenty=2;html=true", Font 
'			Notas_Tit(2, 12).AddText "<div align=""center"">M&eacute;dia<br>Final</div>", "size=7;alignment=center;indenty=2;html=true", Font 
'			Notas_Tit(2, 13).AddText "Res.", "size=7;alignment=center;indenty=6;", Font 
'			Notas_Tit(2, 14).AddText "BIM1", "size=7;alignment=center;indenty=6;", Font 
'			Notas_Tit(2, 15).AddText "BIM2", "size=7;alignment=center;indenty=6;", Font 
'			Notas_Tit(2, 16).AddText "BIM3", "size=7;alignment=center;indenty=6;", Font 
'			Notas_Tit(2, 17).AddText "BIM4", "size=7;alignment=center;indenty=6;", Font 				
			Set param_materias = PDF.CreateParam	
			param_materias.Set "expand=true" 
			
				Set RS0 = Server.CreateObject("ADODB.Recordset")
				SQL = "SELECT * FROM TB_Boletim_Cabecalho where NU_Unidade="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'"

				Set RS0 = CONt.Execute(SQL)
				n_campo=0
				campo="inicio"
				while not campo="fim"
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
					
					if  n_campo>30 then
						campo="fim"					
					end if
				WEND		
				'é menos 2 por que o último campo é o que o programa achou como vazio e retiro mais um por que o campo preenchido como "CARGA"
				'só será usado na função emitir ficha individual
				total_campos=n_campo-1	
			
				
				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL2 = "Select * from TB_Boletim_Notas WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"' AND CO_Matricula="&cod_cons& " ORDER BY NU_Seq"
				Set RS2 = CONt.Execute(SQL2)		
				
				while not RS2.EOF		
					
					Set Row = Notas_Tit.Rows.Add(20) ' row height
					param_materias.Add "indenty=3;alignment=right;html=true"
'					Row.Cells(1).AddText no_materia_exibe, param_materias
					altura_medias=altura_medias+20

					
					param_materias.Add "indentx=0"				
					for nn=1 to total_campos		
						if nn<10 then
							campo_gravacao="CO_0"&nn
						else
							campo_gravacao="CO_"&nn						
						end if					
						notas_exibe=RS2(campo_gravacao)
								
						if notas_exibe="" or isnull(notas_exibe) then
							nota=" "
						else
							if isnumeric(notas_exibe) and nn<14 then
								nota=formatnumber(notas_exibe,1)						
							else
								nota=notas_exibe
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
						Row.Cells(nn).AddText "<div align="""&alinha&"""><font style=""font-size:8pt;"">"&espaco&nota&"</font></div>", param_materias				
					next	
				RS2.MOVENEXT
				WEND							

			Page.Canvas.DrawTable Notas_Tit, "x="&margem&", y="&y_medias&"" 		


			'=============================================================================================================
			tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")
			tipo_freq=tipo_divisao_ano(curso,co_etapa,"in_frequencia")
	
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
			   .Cells(7).Width = 20
			   .Cells(8).Width = 20 
			   .Cells(9).Width = 63 
			   .Cells(10).Width = 250             
			End With
			data_exibe = data&" &agrave;s "& horario

			Legenda(1, 1).Colspan= 8

			if tipo_freq="D" then	
				Set CON_N = Server.CreateObject("ADODB.Connection") 
				ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
				CON_N.Open ABRIRn
							
				Set RSF = Server.CreateObject("ADODB.Recordset")
				SQL_N = "Select * from TB_Frequencia_Periodo WHERE CO_Matricula = "& cod_cons
				Set RSF = CON_N.Execute(SQL_N)
				
				if RSF.eof THEN
					f1=0
					f2=0
					f3=0
					f4=0			
				else	
					f1=RSF("NU_Faltas_P1")
					f2=RSF("NU_Faltas_P2")
					f3=RSF("NU_Faltas_P3")
					f4=RSF("NU_Faltas_P4")		
				END IF				
				soma_falta=0		
		
				if isnull(f1) or f1="" then
					f1=0
				else
					soma_falta=soma_falta*1
					f1=f1*1
					soma_falta=soma_falta+f1					
				end if	
				if isnull(f2) or f2="" then
					f2=0
				else
					soma_falta=soma_falta*1
					f2=f2*1
					soma_falta=soma_falta+f2					
				end if	
				if isnull(f3) or f3="" then
					f3=0
				else
					soma_falta=soma_falta*1
					f3=f3*1
					soma_falta=soma_falta+f3					
				end if											
				
	
				'Legenda(1, 10).RowSpan = 2					
				Legenda(1, 1).AddText "<b>Freq&uuml;&ecirc;ncia (Faltas):</b>", "size=7;html=true", Font 
				Legenda(2, 1).AddText "Trimestre 1:", "size=7;html=true;", Font 
				Legenda(2, 2).AddText ""&f1&"", "size=7;html=true;", Font 
				Legenda(2, 3).AddText "Trimestre 2:", "size=7;html=true;", Font 
				Legenda(2, 4).AddText ""&f2&"", "size=7;html=true;", Font 				
				Legenda(2, 5).AddText "Trimestre 3:", "size=7;html=true;", Font 
				Legenda(2, 6).AddText ""&f3&"", "size=7;html=true;", Font 
				Legenda(2, 7).AddText "Total:", "size=7;html=true;", Font
				Legenda(2, 8).AddText ""&soma_falta&"", "size=7;html=true;", Font 	
			End if							 								

			Legenda(1, 10).AddText "<b><Div align=""right"">Documento impresso em: "&data_exibe&"</div></b>", "size=8; html=true", Font 							
			Page.Canvas.DrawTable Legenda, "x="&margem&", y="&y_legenda&"" 				
				
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
				
				if tp_modelo="B" Then
					if curso=2 and co_etapa	= 3 then
						rows=2
						if lingua_est="ERRO" then
							legenda1=""						
						else
							legenda1="Observa&ccedil;&atilde;o : A L&iacute;ngua Estrangeira cursada &eacute; "&lingua_est
						end if
						legenda2="Bim* - Nota Retificada ap&oacute;s Recupera&ccedil;&atilde;o Paralela"
						h_table=30					
					else
						rows=1
						legenda1="Bim* - Nota Retificada ap&oacute;s Recupera&ccedil;&atilde;o Paralela"
						h_table=15
					end if					
				else
					if curso=2 and co_etapa	= 3 then
						rows=2
						if lingua_est="ERRO" then
							legenda1=""						
						else
							legenda1="Observa&ccedil;&atilde;o : A L&iacute;ngua Estrangeira cursada &eacute; "&lingua_est
						end if
						legenda2="Tri* - Nota Retificada ap&oacute;s Recupera&ccedil;&atilde;o Paralela"
						h_table=30
					else
						rows=1
						legenda1="Tri* - Nota Retificada ap&oacute;s Recupera&ccedil;&atilde;o Paralela"
						h_table=15					
					end if
				end if						
				
				Set param_table4 = Pdf.CreateParam("width="&largura_tabela&"; height="&h_table&"; rows="&rows&"; cols=1; border=0; cellborder=0; cellspacing=0;indenty=2;")
				Set Impresso = Doc.CreateTable(param_table4)
				Impresso.Font = Font
				y_impresso=y_medias-altura_medias-15
				curso=curso*1
				co_etapa=co_etapa*1
				if curso=2 and co_etapa	= 3 then
					Impresso(1, 1).AddText "<b><Div align=""Left"">"&legenda1&"</div></b>", "size=7;html=true", Font 					
					Impresso(2, 1).AddText "<Div align=""Left"">"&legenda2&"</div>", "size=7;html=true", Font 
				else
					Impresso(1, 1).AddText "<Div align=""Left"">"&legenda1&"</div>", "size=7;html=true", Font 									
				end if			
				Page.Canvas.DrawTable Impresso, "x="&margem&", y="&y_impresso&""
						
				 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width=50; alignment=left; size=8; color=#000000")
				'	Relatorio = "Sistema Web Diretor - SWD025"
					Relatorio = "SWD025"
					Do While Len(Relatorio) > 0
						CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
					 
						If CharsPrinted = Len(Relatorio) Then Exit Do
						   SET Page = Page.NextPage
						Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
					Loop 
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
			Page.Canvas.SetParams "LineWidth=0.5" 
			With Page.Canvas
			   .MoveTo 330, 30
			   .LineTo Page.Width - 30, 30
			   .Stroke
			End With 
		
		'Data===========================================================================
			 SET Param_data = Pdf.CreateParam("x=30;y=40; height=50; width=100; alignment=Left; size=8; color=#000000")
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
			

			width_tesoura=Page.Width - 30
			 SET Param_Tesoura = Pdf.CreateParam("x=30;y=136; height=50; width="&width_tesoura&"; alignment=center; size=16; color=#000000, html=true")
			Tesoura = "<div align=""center"">&quot;</div>"
			
			Do While Len(Tesoura) > 0
				CharsPrinted = Page.Canvas.DrawText(Tesoura, Param_Tesoura, Font_Tesoura )
			 
				If CharsPrinted = Len(Tesoura) Then Exit Do
				   SET Page = Page.NextPage
				Tesoura = Right( Tesoura, Len(Tesoura) - CharsPrinted)
			Loop 
			
			largura_logo_pqno=formatnumber(Logo.Width*0.15,0)
			altura_logo_pqno=formatnumber(Logo.Height*0.15,0)
			
			Set Param_Logo_Pqno = Pdf.CreateParam
			   Param_Logo_Pqno("x") = 39 
			   Param_Logo_Pqno("y") = altura_logo_pqno+41
			   Param_Logo_Pqno("ScaleX") = 0.15
			   Param_Logo_Pqno("ScaleY") = 0.15
			   Page.Canvas.DrawImage Logo, Param_Logo_Pqno
			
			x_texto_recibo=largura_logo_pqno+ 43
			y_texto_recibo=formatnumber(altura_logo_pqno*2.8,0)
			width_texto=Page.Width -largura_logo_pqno - 100
			SET Param_recibo = Pdf.CreateParam("x="&x_texto_recibo&";y="&y_texto_recibo&"; height="&altura_logo_pqno&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
			Text_recibo = "<p><center><i><b>"&nome_da_escola&"</b></i></center></p>"
			
			Do While Len(Text_recibo) > 0
				CharsPrinted = Page.Canvas.DrawText(Text_recibo, Param_recibo, Font )
			 
				If CharsPrinted = Len(Text_recibo) Then Exit Do
					SET Page = Page.NextPage
				Text_recibo = Right( Text_recibo, Len(Text_recibo) - CharsPrinted)
			Loop 
			
			Set param_recibo2 = Pdf.CreateParam("width=100; height=40; rows=2; cols=2; border=0; cellborder=0; cellspacing=0;")
			Set recibo_2 = Doc.CreateTable(param_recibo2)
			y_recibo_2=y_texto_recibo
			With recibo_2.Rows(1)
			   .Cells(1).Width = 50
			   .Cells(2).Width = 50
			End With
			recibo_2(1, 1).AddText "<Div align=""center"">Ano Letivo:</div>", "size=8;indenty=2;html=true", Font 
			recibo_2(1, 2).AddText "<Div align=""Right""><b>"&ano_letivo&"</b></div>", "size=9;indenty=2;html=true", Font 
			recibo_2(2, 1).AddText "<Div align=""center"">Matr&iacute;cula:</div>", "size=8;indenty=2;html=true", Font 
			recibo_2(2, 2).AddText "<Div align=""Right""><b>"&cod_cons&"</b></div>", "size=8;indenty=2;html=true", Font 
			Page.Canvas.DrawTable recibo_2, "x=455, y="&y_recibo_2&"" 
		
		
			
			SET Param_Nome_Aluno_Recibo = Pdf.CreateParam("x=30;y=75; height=50; width="&width_nome_aluno&"; alignment=left; html=True")
			Nome_Recibo = "<font style=""font-size:10pt;""><b>"&nome_aluno&"</b></font>"
			Do While Len(Nome_Recibo) > 0
				CharsPrinted = Page.Canvas.DrawText(Nome_Recibo, Param_Nome_Aluno_Recibo, Font )
			 
				If CharsPrinted = Len(Nome_Recibo) Then Exit Do
					SET Page = Page.NextPage
				Nome_Recibo = Right( Nome_Recibo, Len(Nome_Recibo) - CharsPrinted)
			Loop 
		
			SET Param_Dados_Aluno_Recibo = Pdf.CreateParam("x=30;y=60; height=50; width="&width_nome_aluno&"; alignment=left; html=True")
			'Dados_Recibo = "<font style=""font-size:8pt;"">Curso: <b>"&no_curso &"</b> S&eacute;rie: <b>"&no_etapa &"</b> Turma: <b>"& turma &"</b></font>"
			Dados_Recibo = "<font style=""font-size:8pt;"">Curso: <b>"&no_curso &"</b> Turma: <b>"& turma &"</b></font>"				
			Do While Len(Dados_Recibo) > 0
				CharsPrinted = Page.Canvas.DrawText(Dados_Recibo, Param_Dados_Aluno_Recibo, Font )
			 
				If CharsPrinted = Len(Dados_Recibo) Then Exit Do
					SET Page = Page.NextPage
				Dados_Recibo = Right( Dados_Recibo, Len(Dados_Recibo) - CharsPrinted)
			Loop 
			
			y_devolver=y_texto_recibo-20
			SET Param_recibo3 = Pdf.CreateParam("x="&x_texto_recibo&";y="&y_devolver&"; height=60; width="&width_texto&"; alignment=center; size=12; color=#000000; html=true")
			Text_recibo3 = "<p><center><b><font style=""font-size:7pt;"">Devolver essa parte assinada ao col&eacute;gio</font><b></center></p>"
			
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

