<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 60 'valor em segundos
'FICHA INDIVIDUAL
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/parametros.asp"-->
<!--#include file="../inc/calculos.asp"-->
<!--#include file="../inc/resultados.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/funcoes7.asp"-->
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
	Set Logo1 = Doc.OpenImage( Server.MapPath( "../img/logo_niteroi_preto.gif") )
	Set Logo2 = Doc.OpenImage( Server.MapPath( "../img/logo_arariboia_preto.gif") )	
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

		
'	Set RS0 = Server.CreateObject("ADODB.Recordset")
'	SQL0 = "SELECT * FROM TB_Periodo ORDER BY NU_Periodo"
'	RS0.Open SQL0, CON0
'	check_periodo=1
'	WHILE NOT RS0.EOF
'		periodo=RS0("NU_Periodo")
'		if check_periodo=1 then
'			vetor_periodo=periodo
'		else
'			vetor_periodo=vetor_periodo&"#!#"&periodo
'		end if
'		check_periodo=check_periodo+1 
'	RS0.MOVENEXT
'	WEND				
		
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod_cons
	RS.Open SQL, CON1
	
	nome_aluno = RS("NO_Aluno")
	sexo_aluno = RS("IN_Sexo")
	nome_pai = RS("NO_Pai")
	nome_mae = RS("NO_Mae")
	natural= RS("SG_UF_Natural")	
	nome_aluno=replace_latin_char(nome_aluno,"html")	
'	in_lingua=RS("IN_Lingua")

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

'	select case in_lingua
'		case "D"
			lingua_est = "Ingl&ecirc;s"
'		case "I"
'			lingua_est = "Espanhol"
'		case "V"
'			lingua_est = "Franc&ecirc;s"
'		case else
'			lingua_est="ERRO"	
'	end select								

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
			no_abrv_curso= RS3("NO_Abreviado_Curso")
			co_concordancia_curso = RS3("CO_Conc")	
			
			no_unidade = unidade&" - "&no_unidade
			no_curso= no_etapa&" "&co_concordancia_curso&" "&no_curso
			texto_turma = "<div align='Right'> - Turma:"&turma&"</div>"
			texto_chamada=" N&ordm;. Chamada: "&cham	
			
			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL2 = "SELECT * FROM TB_Contatos WHERE CO_Matricula="& cod_cons &" AND TP_Contato='ALUNO'"
			RS2.Open SQL2, CONCONT

			if RS2.EOF then
			else 	
				nascimento=RS2("DA_Nascimento_Contato")
			end if	

			Set RS3 = Server.CreateObject("ADODB.Recordset")
			SQL3 = "SELECT * FROM TB_UF WHERE SG_UF='"& natural &"'"
			RS3.Open SQL3, CON0
				
			no_uf = RS3("NO_UF")						

			SET Page = Doc.Pages.Add( 595, 842 )
					
'CABEÇALHO==========================================================================================		
			Set Param_Logo_Gde = Pdf.CreateParam
			margem=30				
			linha=10		
			unidade = unidade*1	
			if unidade = 1 then
				largura_logo_gde=formatnumber(Logo1.Width*0.5,0)
				altura_logo_gde=formatnumber(Logo1.Height*0.5,0)
				area_utilizavel=Page.Width-(margem*2)
				Param_Logo_Gde("x") = margem
				Param_Logo_Gde("y") = Page.Height - altura_logo_gde -margem
				Param_Logo_Gde("ScaleX") = 0.5
				Param_Logo_Gde("ScaleY") = 0.5
				Page.Canvas.DrawImage Logo1, Param_Logo_Gde
			else
				largura_logo_gde=formatnumber(Logo2.Width*0.5,0)
				altura_logo_gde=formatnumber(Logo2.Height*0.5,0)
				area_utilizavel=Page.Width-(margem*2)
				Param_Logo_Gde("x") = margem
				Param_Logo_Gde("y") = Page.Height - altura_logo_gde -margem
				Param_Logo_Gde("ScaleX") = 0.5
				Param_Logo_Gde("ScaleY") = 0.5
				Page.Canvas.DrawImage Logo2, Param_Logo_Gde		
			end if
	
			'x_texto=largura_logo_gde+ 30
			x_texto= margem
			y_texto=formatnumber(Page.Height - altura_logo_gde/2,0)
			width_texto=Page.Width - 60

		
			SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
			Text = "<p><center><i><b><font style=""font-size:18pt;"">Ficha Individual</font></b></i></center></p>"
			

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
	
			Set param_table1 = Pdf.CreateParam("width="&area_utilizavel&"; height=72; rows=6; cols=8; border=0; cellborder=0; cellspacing=0;")
			Set Table = Doc.CreateTable(param_table1)
			Table.Font = Font
			y_primeira_tabela=altura_primeiro_separador-10
			x_primeira_tabela=margem+5
			With Table.Rows(1)
			   .Cells(1).Width = 50
			   .Cells(2).Width = 15  
			   .Cells(3).Width = 20	
			   .Cells(4).Width = 70	
			   .Cells(5).Width = 20	
			   .Cells(6).Width = 50	
   			   .Cells(7).Width = 70				   		   		   
			End With
			Table(1, 2).ColSpan = 5
			Table(2, 2).ColSpan = 5
			Table(3, 2).ColSpan = 5
			Table(4, 2).ColSpan = 5
			Table(5, 3).ColSpan = 2	
			Table(5, 5).ColSpan = 3													
			Table(6, 2).ColSpan = 2
			Table(6, 4).ColSpan = 3	
			
			Table(1, 1).AddText "Matr&iacute;cula:", "size=8;html=true", Font 
			Table(1, 2).AddText "<b>"&cod_cons&"</b>", "size=8;html=true", Font 
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
total_colunas=0			
for tc=0 to ubound(colunas_tit)
	coluna_afetada=coluna_afetada*1
	colunas_tit_colspan(tc)=colunas_tit_colspan(tc)*1
	total_colunas=total_colunas+colunas_tit_colspan(tc)							
next
'incluir a coluna de carga		
total_colunas=total_colunas+1	
	
			altura_medias=40
			largura_tabela=570
			Set param_table2 = Pdf.CreateParam("width="&largura_tabela&"; height="&altura_medias&"; rows=3; cols="&total_colunas&"; border=1; cellborder=0.1; cellspacing=0;")
			Set Notas_Tit = Doc.CreateTable(param_table2)
			Notas_Tit.Font = Font
			y_medias=y_primeira_tabela-72-10	
			larg_col=formatnumber((largura_tabela-130-30)/total_colunas,2)
			
			With Notas_Tit.Rows(1)
			   .Cells(1).Width = 130
				for wd=2 to total_colunas
				   .Cells(wd).Width = larg_col
				next
				.Cells(total_colunas).Width = 30								  		         			   			         
			End With
coluna_afetada=1
for ctr=0 to ubound(colunas_tit)
	coluna_afetada=coluna_afetada*1
	colunas_tit_colspan(ctr)=colunas_tit_colspan(ctr)*1
	'pegar o span da primeira coluna pra depois incluir como span da coluna de carga	
	if ctr=0 then 
		span_prim_col=colunas_tit_rowspan(ctr)
	end if	
	
	colunas_tit_rowspan(ctr)=colunas_tit_rowspan(ctr)*1
	if colunas_tit_rowspan(ctr)	> 1 then
		Notas_Tit(1, coluna_afetada).RowSpan = colunas_tit_rowspan(ctr)		
	end if	
	coluna_afetada=coluna_afetada+colunas_tit_colspan(ctr)						
next
'incluir o span da coluna de carga
if span_prim_col > 1 then
	Notas_Tit(1, total_colunas).RowSpan = span_prim_col
end if	

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
		
			Notas_Tit(1, total_colunas).AddText "Carga", "size=9;alignment=center;indenty=15;", Font 				
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
					
					if  n_campo>50 then
						campo="fim"					
					end if
				WEND		
				'é menos 2 por que o último campo é o que o programa achou como vazio
				total_campos=n_campo
				
				tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")
				tipo_freq=tipo_divisao_ano(curso,co_etapa,"in_frequencia")	
				
			if tipo_freq="D" then
	
				Set CON_N = Server.CreateObject("ADODB.Connection") 
				ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
				CON_N.Open ABRIRn
							
				Set RSF = Server.CreateObject("ADODB.Recordset")
				SQL_N = "Select * from TB_Frequencia_Periodo WHERE CO_Matricula = "& cod_cons
				Set RSF = CON_N.Execute(SQL_N)
				
				soma_faltas=0			
				
				if RSF.eof THEN
					f1="&nbsp;"
					f2="&nbsp;"
					f3="&nbsp;"
					f4="&nbsp;"	
				else	
					f1=RSF("NU_Faltas_P1")
					f2=RSF("NU_Faltas_P2")
					f3=RSF("NU_Faltas_P3")
					f4=RSF("NU_Faltas_P4")		
					
					if isnull(f1) or f1= "" then
					else
						f1=f1*1
						soma_faltas=soma_faltas*1
						soma_faltas=soma_faltas+f1		
					end if
					
					if isnull(f2) or f2= "" then
					else
						f2=f2*1
						soma_faltas=soma_faltas*1
						soma_faltas=soma_faltas+f2		
					end if
					
					if isnull(f3) or f3= "" then
					else
						f3=f3*1
						soma_faltas=soma_faltas*1
						soma_faltas=soma_faltas+f3		
					end if
					
					if isnull(f4) or f4= "" then
					else
						f4=f4*1
						soma_faltas=soma_faltas*1
						soma_faltas=soma_faltas+f4		
					end if									
				END IF				
				total_faltas=soma_faltas
			else
				campo_faltas_inicio=0
				campo_faltas_fim=0
				for ccc= 0 to UBOUND(colunas_notas_calc)
				
					IF colunas_notas_calc(ccc) = "BDF" then
						if campo_faltas_inicio=0 then
							campo_faltas_inicio=ccc
						end if
					else	
						if campo_faltas_inicio<>0 then
							campo_faltas_fim=ccc
						end if
					end if
				next
				if campo_faltas_fim=0 then
					campo_faltas_fim=UBOUND(colunas_notas_calc)
				end if	
			end if
			'Adiciona 2, para compensar o 0, pois não existe o campo CO_00 e o CO_01 é o nome da disciplina
			campo_faltas_inicio=campo_faltas_inicio+2	
			campo_faltas_fim=campo_faltas_fim+2
				

				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL2 = "Select * from TB_Boletim_Notas WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"' AND CO_Matricula="&cod_cons& " ORDER BY NU_Seq"
				Set RS2 = CONt.Execute(SQL2)		
				total_faltas=0
				horas_aula=0
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
						'response.Write(campo_gravacao)
						if notas_exibe="" or isnull(notas_exibe) then
							nota=" "
						else
		
							if isnumeric(notas_exibe) and nn<14 then
								nota=formatnumber(notas_exibe,0)						
							else
								nota=notas_exibe
							end if	
							nn=nn*1
							campo_faltas_inicio=campo_faltas_inicio*1
							campo_faltas_fim=campo_faltas_fim*1
							if nn>=campo_faltas_inicio and nn<=campo_faltas_fim then
								total_faltas=total_faltas*1
								nota=nota*1
								if nota="" or isnull(nota) then
								else
									total_faltas=total_faltas+nota	
								end if	
							end if														
							if nn=total_campos then
								curso=curso*1
								co_etapa=co_etapa*1								
								if curso=1 and co_etapa<6  then
									nota="---"
									horas_aula=960
								else
									horas_aula=horas_aula*1
									nota=nota*1
									horas_aula=horas_aula+nota
								end if	
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

			Page.Canvas.DrawTable Notas_Tit, "x=30, y="&y_medias&"" 		

'			soma_faltas=soma_faltas*1
'			dias_de_aula_no_ano=200
'			
'			frequencia=((dias_de_aula_no_ano-soma_faltas)/dias_de_aula_no_ano)*100
			total_faltas=total_faltas*1
			horas_aula=horas_aula*1


			if horas_aula=0 or isnull(horas_aula) then
				frequencia=100
			else
				frequencia=((horas_aula-total_faltas)/horas_aula)*100				
			end if
		
			if frequencia<100 then
				frequencia=arredonda(frequencia,"mat_dez",1,0)	
			end if
			
		
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
					Legenda(1, 1).AddText "<b>Freq&uuml;&ecirc;ncia (Faltas):</b>", "size=7;html=true", Font 
					Legenda(2, 1).AddText "Trimestre 1:", "size=7;html=true;", Font 
					Legenda(2, 2).AddText ""&f1&"", "size=7;html=true;", Font 
					Legenda(2, 3).AddText "Trimestre 2:", "size=7;html=true;", Font 
					Legenda(2, 4).AddText ""&f2&"", "size=7;html=true;", Font 				
					Legenda(2, 5).AddText "Trimestre 3:", "size=7;html=true;", Font 
					Legenda(2, 6).AddText ""&f3&"", "size=7;html=true;", Font 
					Legenda(2, 7).AddText "Total:", "size=7;html=true;", Font
					Legenda(2, 8).AddText ""&total_faltas&"", "size=7;html=true;", Font 	
				End if
				Legenda(1, 10).AddText "<b><Div align=""right"">CARGA HOR&Aacute;RIA TOTAL DE "&horas_aula&" HORAS</div></b>", "size=8; html=true", Font 
				Legenda(2, 10).AddText "<b><Div align=""right"">Documento impresso em: "&data_exibe&"</div></b>", "size=8; html=true", Font 				
				Page.Canvas.DrawTable Legenda, "x=30, y="&y_legenda&"" 
				
				
				SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width=50; alignment=left; size=8; color=#000000")
				
				if tp_modelo="B" Then
					if curso=2 and co_etapa	= 3 then
						rows=2
'						if lingua_est="ERRO" then
'							legenda1=""						
'						else
'							legenda1="Observa&ccedil;&atilde;o : A L&iacute;ngua Estrangeira cursada &eacute; "&lingua_est
'						end if
'						legenda2="Bim* - Nota Retificada ap&oacute;s Recupera&ccedil;&atilde;o Paralela"
						h_table=30
						legenda1=""	
						legenda2=""												
					else
						rows=1
						'legenda1="Bim* - Nota Retificada ap&oacute;s Recupera&ccedil;&atilde;o Paralela"
						h_table=15
						legenda1=""	
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

				Relatorio = "SWD048"
				Do While Len(Relatorio) > 0
					CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
				 
					If CharsPrinted = Len(Relatorio) Then Exit Do
					   SET Page = Page.NextPage
					Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
				Loop 
			
				Observ = "O aluno obteve "&frequencia&"% de frequ&ecirc;ncia no ano letivo."
				margem_obs=margem*3.3
				SET Param_Obs = Pdf.CreateParam("x="&margem_obs&";y=140; height=50; width=500; alignment=left; size=8; color=#000000;html=true")

					Do While Len(Relatorio) > 0
						CharsPrinted = Page.Canvas.DrawText(Observ, Param_Obs, Font )
					 
						If CharsPrinted = Len(Observ) Then Exit Do
						   SET Page = Page.NextPage
						Observ = Right(Observ, Len(Observ) - CharsPrinted)
					Loop 

'INÍCIO DO RECIBO==================================================================================

				Set param_table4 = Pdf.CreateParam("width=540; height=125; rows=3; cols=1; border=1; cellborder=1; cellspacing=0;")
				Set quadro = Doc.CreateTable(param_table4)
				quadro.Font = Font
				quadro.Rows(1).Height = 20
				quadro.Rows(2).Height = 60
				quadro.Rows(3).Height = 45
         
exibe_resultado="s"


				if exibe_resultado="s" THEN
					vetor_materia=vetor_disciplinas(ano_letivo,unidade,curso,co_etapa,turma,"n",0)
					vetor_materia_resultado=split(vetor_materia,"#!#")
					prd_ter_media=Periodo_Media(tp_modelo,"MF",outro)
					resultados_apurados=""	
					for r=0 to ubound(vetor_materia_resultado)
						m3=Calc_Ter_Media (unidade, curso, co_etapa, turma, cod_cons, vetor_materia_resultado(r), caminho_nota, tb_nota, prd_ter_media, "sem_calculo", "ficha")		
						res=split(m3,"#!#")
						if r=0 then
							resultados_apurados=res(1)
						else
							resultados_apurados=resultados_apurados&"#!#"&res(1)						
						end if				
					next

					if resultados_apurados="" then
						resultado_final_aluno="nulo"
					else				
						resultado_final_aluno=apura_resultado_aluno(curso, co_etapa, resultados_apurados)
					end if
	
				Set RSs = Server.CreateObject("ADODB.Recordset")
				SQL_s ="SELECT CO_Situacao FROM TB_Matriculas where CO_Matricula ="& cod_cons&" AND NU_Ano="&ano_letivo
				Set RSs = CON1.Execute(SQL_s)
		
				situac=RSs("CO_Situacao")
				
				if situac<>"C" THEN
				
					Set RSt = Server.CreateObject("ADODB.Recordset")
					SQLt ="SELECT TX_Descricao_Situacao FROM TB_Situacao_Aluno where CO_Situacao = '"&situac&"'"
					Set RSt = CON0.Execute(SQLt)	
					
					if RSt.eof then			
						resultado_exibe="Cancelado"
					else
						resultado_exibe=replace_latin_char(RSt("TX_Descricao_Situacao"),"html")
					end if	
				elseif (resultado_final_aluno="Apr" or resultado_final_aluno="APR") then
					'	legenda_resultado="<b>Resultado:</b>"
						resultado_exibe="Aprovado"
					elseif (resultado_final_aluno="Rep" or resultado_final_aluno="REP") then	
					'	legenda_resultado="<b>Resultado:</b>"				
						resultado_exibe="Reprovado"	
					elseif (resultado_final_aluno="Pfi" or resultado_final_aluno="PFI") then	
						'legenda_resultado="<b>Resultado:</b>"					
						resultado_exibe="Prova Final"							
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
				
				quadro(1, 1).AddText "<b>Resultado Final:</b>&nbsp;&nbsp;"&resultado_exibe, "size=9;indentx=2; html=true", Font 
				quadro(2, 1).AddText "<b>Observa&ccedil;&otilde;es:</b>", "size=9;indentx=5;indenty=5;html=true;", Font 
'				Legenda(1, 5).AddText "<b>R</b> = Recupera&ccedil;&atilde;o", "size=9;html=true", Font 
'				Legenda(1, 7).AddText "<b>Md</b> = M&eacute;dia", "size=9; html=true", Font 
'				Legenda(1, 10).AddText "<b><Div align=""right"">Documento impresso em: "&data_exibe&"</div></b>", "size=8; html=true", Font 
				Page.Canvas.DrawTable quadro, "x="&margem&", y=165" 

			Page.Canvas.SetParams "LineWidth=1" 
			
			
			y_obs_linha1=129
				With Page.Canvas
				   .MoveTo 100, y_obs_linha1
				   .LineTo 570, y_obs_linha1
				   .Stroke
				End With 	
			
			y_obs_linha2=y_obs_linha1-16
				With Page.Canvas
				   .MoveTo margem, y_obs_linha2
				   .LineTo 570, y_obs_linha2
				   .Stroke
				End With 
				
			y_obs_linha3=y_obs_linha2-16
				With Page.Canvas
				   .MoveTo margem, y_obs_linha3
				   .LineTo 570, y_obs_linha3
				   .Stroke
				End With 				
		'Data===========================================================================
			 SET Param_data = Pdf.CreateParam("x=473;y=80; height=40; width=100; alignment=Left; size=8; color=#000000")
			data_preenche = "Data:_____/_____/_____"
			CharsPrinted = Page.Canvas.DrawText(data_preenche, Param_data, Font )
		
			
		'===========================================================================
			 SET Param_secretario = Pdf.CreateParam("x=40;y=50; height=40; width=100; alignment=Left; size=9; color=#000000;html=true")
			secretario = "<b>Secretario(a)</b>"
			CharsPrinted = Page.Canvas.DrawText(secretario, Param_secretario, Font )

			 SET Param_diretor = Pdf.CreateParam("x=390;y=50; height=40; width=100; alignment=Left; size=9; color=#000000;html=true")
			diretor = "<b>Diretor(a</b>)"
			CharsPrinted = Page.Canvas.DrawText(diretor, Param_diretor, Font )
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

