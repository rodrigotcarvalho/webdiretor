<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 180 'valor em segundos
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/funcoes7.asp"-->
<!--#include file="../../global/conta_alunos.asp"-->
<!--#include file="../../global/tabelas_escolas.asp"-->
<!--#include file="../../global/notas_calculos_diversos.asp"-->
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
Paginacao = "0"

	obr = request.QueryString("obr")
	dados=obr
	dados_funcao=split(obr,"$!$")

	unidade = dados_funcao(0)
	curso = dados_funcao(1)
	co_etapa = dados_funcao(2)
	turma = dados_funcao(3)
	periodo = dados_funcao(4)
	acumulado = dados_funcao(5)
	qto_falta = dados_funcao(6)	
	ano_letivo = dados_funcao(7)
	'Não são utilizadas nessa função
	'larg_tabela = dados_funcao(8)
	'alt_tabela = dados_funcao(9)
	
	
	obr_mapa=unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma&"$!$"&periodo&"$!$"&acumulado&"$!$"&qto_falta&"$!$"&ano_letivo	

arquivo="SWD102"

if ori="acc" then
origem="../wa/professor/cna/acc/"
end if

if mes<10 then
mes="0"&mes
end if

ano = DatePart("yyyy", now)
ano_vigente=ano


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

periodo_m1=3 
periodo_m2=4
periodo_m3= 99


	tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")
	tp_freq=tipo_divisao_ano(curso,co_etapa,"in_frequencia")	
	num_periodo_detalhe=periodos_ACC(periodo,acumulado,qto_falta,tp_modelo,"num",0)
	vetor_num_periodo=num_periodo_detalhe
	num_periodo=split(num_periodo_detalhe,"#!#")	

	vetor_nom_periodo=periodos_ACC(periodo,acumulado,qto_falta,tp_modelo,"nom",0)
	javascript_periodo=vetor_nom_periodo
	nom_periodo=split(vetor_nom_periodo,"#!#")
	
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CON_AL = Server.CreateObject("ADODB.Connection") 
	ABRIR_AL = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_AL.Open ABRIR_AL
	
	Set CONg = Server.CreateObject("ADODB.Connection") 
	ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONg.Open ABRIRg		
	
	Set CON3 = Server.CreateObject("ADODB.Connection") 
	ABRIR3 = "DBQ="& CAMINHO_o & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON3.Open ABRIR3	

	Set CONt = Server.CreateObject("ADODB.Connection") 
	ABRIRt = "DBQ="& CAMINHO_t & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONt.Open ABRIRt		
	
	cor_nota_vml="#FF0000"	
	cor_nota_azl="#0000FF"	
	cor_nota_prt="#000000"	
	cor_nota_vrd="#006600"	
	
	tb_nota=tabela_nota(ano_letivo,unidade,curso,co_etapa,turma,"tb",0)
	caminho_nota=tabela_nota(ano_letivo,unidade,curso,co_etapa,turma,"cam",0)
	avaliacao = var_bd_periodo(tp_modelo,tp_freq,tb_nota,periodo,"BDM")
	no_unidade= GeraNomes("U",unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)
	no_curso=GeraNomes("C",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro) 	
	no_etapa=GeraNomes("E",curso,co_etapa,variavel3,variavel4,variavel5,CON0,outro) 


	if tb_nota="erro" or caminho_nota="erro" then
		gera_pdf="nao"
	else
		gera_pdf="sim"	
	end if

if gera_pdf="sim" then	

	Set CON_N = Server.CreateObject("ADODB.Connection")
	ABRIR3 = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_N.Open ABRIR3

	Set RS2 = Server.CreateObject("ADODB.Recordset")
	SQL2 = "SELECT * FROM TB_Unidade WHERE NU_Unidade="& unidade
	RS2.Open SQL2, CON0
					
	no_unidade = RS2("TX_Imp_Cabecalho")		
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
		
		municipio_unidade=RS3m("NO_Municipio")						
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

	mensagem_cabecalho=ano_letivo&"&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;"&no_curso&texto_turma

	SET Page = Doc.Pages.Add(842, 595)
			
'CABEÇALHO==========================================================================================		
	Set Param_Logo_Gde = Pdf.CreateParam
	margem=25			
	linha=10		
	unidade = unidade*1	
	if unidade = 1 then
		largura_logo_gde=formatnumber(Logo1.Width*0.7,0)
		altura_logo_gde=formatnumber(Logo1.Height*0.7,0)
		area_utilizavel=Page.Width-(margem*2)
		Param_Logo_Gde("x") = margem
		Param_Logo_Gde("y") = Page.Height - altura_logo_gde -margem
		Param_Logo_Gde("ScaleX") = 0.7
		Param_Logo_Gde("ScaleY") = 0.7
		Page.Canvas.DrawImage Logo1, Param_Logo_Gde
	else
		largura_logo_gde=formatnumber(Logo2.Width*0.7,0)
		altura_logo_gde=formatnumber(Logo2.Height*0.7,0)
		area_utilizavel=Page.Width-(margem*2)
		Param_Logo_Gde("x") = margem
		Param_Logo_Gde("y") = Page.Height - altura_logo_gde -margem
		Param_Logo_Gde("ScaleX") = 0.7
		Param_Logo_Gde("ScaleY") = 0.7
		Page.Canvas.DrawImage Logo2, Param_Logo_Gde		
	end if

	x_texto=largura_logo_gde+ margem+10
	y_texto=formatnumber(Page.Height - margem,0)
	width_texto=Page.Width -largura_logo_gde - 80


	SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
	Text = "<p><i><b>"&UCASE(no_unidade)&"</b></i></p><br><font style=""font-size:10pt;"">"&endereco_unidade&"</FONT><br><center><i><b><font style=""font-size:18pt;"">MAP&Atilde;O DE M&Eacute;DIAS POR PER&Iacute;ODO</font></b></i></center>"
	
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

		
	Set RSapr = Server.CreateObject("ADODB.Recordset")
	SQLapr = "Select * from TB_Regras_Aprovacao WHERE CO_Curso = '"& curso &"' AND CO_Etapa='"&co_etapa&"'"
	Set RSapr = CON0.Execute(SQLapr)
	
	if RSapr.EOF then
		ntvml=0
	else
		ntazl= RSapr("NU_Valor_M1")		
		ntvml= RSapr("NU_Valor_M2")
		peso_m2_m1=RSapr("NU_Peso_Media_M2_M1")
		peso_m2_m2=RSapr("NU_Peso_Media_M2_M2")
		peso_m3_m1=RSapr("NU_Peso_Media_M3_M1")
		peso_m3_m2=RSapr("NU_Peso_Media_M3_M2")
		peso_m3_m3=RSapr("NU_Peso_Media_M3_M3")		
	end if

	vetor_materia=cod_cons

'	Set RS = Server.CreateObject("ADODB.Recordset")
'	SQL_A = "Select * from TB_Matriculas WHERE NU_Ano="&ano_letivo&" AND NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
'
'	Set RS = CON_AL.Execute(SQL_A)
'	IF RS.EOF Then
'		alunos_vetor="nulo"
'	else		
'		co_aluno_check=0
'		While Not RS.EOF
'		nu_matricula = RS("CO_Matricula")
'		nu_chamada = RS("NU_Chamada")		
'		
'			Set RSs = Server.CreateObject("ADODB.Recordset")
'			SQL_s ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& nu_matricula&" and TB_Matriculas.NU_Ano="&ano_letivo
'			Set RSs = CON_AL.Execute(SQL_s)
'	
'			situac=RSs("CO_Situacao")
'			nome_aluno=RSs("NO_Aluno")		
'			'strReplacement = Server.URLEncode(nome_aluno)	
'			strReplacement = nome_aluno
'			strReplacement = replace(strReplacement,"+"," ")
'			strReplacement = replace(strReplacement,"%27","´")
'			strReplacement = replace(strReplacement,"%27","'")
'			strReplacement = replace(strReplacement,"À,","&Agrave;")
'			strReplacement = replace(strReplacement,"Á","&Aacute;")
'			strReplacement = replace(strReplacement,"Â","&Acirc;")
'			strReplacement = replace(strReplacement,"Ã","&Atilde;")
'			strReplacement = replace(strReplacement,"É","&Eacute;")
'			strReplacement = replace(strReplacement,"Ê","&Ecirc;")
'			strReplacement = replace(strReplacement,"Í","&Iacute;")
'			strReplacement = replace(strReplacement,"Ó","&Oacute;")
'			strReplacement = replace(strReplacement,"Ô","&Ocirc;")
'			strReplacement = replace(strReplacement,"Õ","&Otilde;")
'			strReplacement = replace(strReplacement,"Ú","&Uacute;")
'			strReplacement = replace(strReplacement,"Ü","&Uuml;")	
'			strReplacement = replace(strReplacement,"à","&agrave;")
'			strReplacement = replace(strReplacement,"á","&aacute;")
'			strReplacement = replace(strReplacement,"â","&acirc;")
'			strReplacement = replace(strReplacement,"ã","&atilde;")
'			strReplacement = replace(strReplacement,"ç","&ccedil;")
'			strReplacement = replace(strReplacement,"é","&eacute;")
'			strReplacement = replace(strReplacement,"ê","&ecirc;")
'			strReplacement = replace(strReplacement,"í","&iacute;")
'			strReplacement = replace(strReplacement,"ó","&oacute;")
'			strReplacement = replace(strReplacement,"ô","&ocirc;")
'			strReplacement = replace(strReplacement,"õ","&otilde;")
'			strReplacement = replace(strReplacement,"ú","&uacute;")
'			strReplacement = replace(strReplacement,"ü","&uuml;")
'			nome_aluno =strReplacement	
'			
'			if situac<>"C" then
'				nome_aluno=nome_aluno&" - Aluno Inativo"
'			end if
'
'			if co_aluno_check=0 then
'				alunos_vetor=nu_matricula&"#!#"&nu_chamada&"#!#"&nome_aluno
'			else
'				alunos_vetor=alunos_vetor&"#$#"&nu_matricula&"#!#"&nu_chamada&"#!#"&nome_aluno
'			end if
'			co_aluno_check=co_aluno_check+1	
'		RS.MOVENEXT
'		WEND
'	END IF			

	alunos_vetor=alunos_turma(ano_letivo,unidade,curso,co_etapa,turma,0)


	n_alunos= split(alunos_vetor,"#$#")	

'	Set RS5 = Server.CreateObject("ADODB.Recordset")
'	SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND IN_MAE=TRUE order by NU_Ordem_Boletim "
'	RS5.Open SQL5, CON0
'	co_materia_check=1
'	IF RS5.EOF Then
'		vetor_materia_exibe="nulo"
'	else
'		while not RS5.EOF
'			co_mat_fil= RS5("CO_Materia")
'			'carga_materia= RS5("NU_Aulas")				
'			if co_materia_check=1 then
'				vetor_materia=co_mat_fil
'			else
'				vetor_materia=vetor_materia&"#!#"&co_mat_fil
'			end if
'			co_materia_check=co_materia_check+1			
'					
'		RS5.MOVENEXT
'		wend						
'	end if
'
'	if vetor_materia_exibe="nulo" then
'		Response.Write("Erro 1 - Não foram encontradas matérias para Etapa ='"& co_etapa &"' e Curso ="& curso)
'	else
'		vetor_materia_exibe=programa_aula(vetor_materia, unidade, curso, co_etapa, turma)
'	end if

	Set RSd = Server.CreateObject("ADODB.Recordset")
	SQLd = "SELECT * FROM TB_Mapao_Disciplinas where NU_Unidade="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'"
	Set RSd = CONt.Execute(SQLd)

	
	If RSd.EOF THEN	
		response.Write("mapa.asp?ln167 - ERRO no acesso a tabela TB_Mapao_Disciplinas")
		response.end()	
	else
		for conta_materias=1 to 30
			if conta_materias<10 then
				campo="CO_0"&conta_materias
			else
				campo="CO_"&conta_materias			
			end if
			no_mat=RSd(campo)
			if no_mat="" or isnull(no_mat) then

			else
				if conta_materias=1 then
					vetor_materia_exibe=no_mat
				else	
					vetor_materia_exibe=vetor_materia_exibe&"#!#"&no_mat
				end if
			end if		
		next
	end if	
	vet_co_materia= split(vetor_materia_exibe,"#!#")	
	co_materia_check=1	

	colunas_de_notas=(ubound(vet_co_materia)+1)
	total_de_colunas=colunas_de_notas+3						
	altura_medias=18
	y_segunda_tabela=y_primeira_tabela-20	
	Set param_table2 = Pdf.CreateParam("width="&area_utilizavel&"; height="&altura_medias&"; rows=1; cols="&total_de_colunas&"; border=1; cellborder=0.1; cellspacing=0; x="&margem&"; y="&y_segunda_tabela&"; MaxHeight=420")

	Set Notas_Tit = Doc.CreateTable(param_table2)
	Notas_Tit.Font = Font				
	largura_colunas=(area_utilizavel-40-20-220)/colunas_de_notas		
	With Notas_Tit.Rows(1)
	   .Cells(1).Width = 20
	   .Cells(2).Width = 220	
	   .Cells(3).Width = 40	   		             
		for d=4 to total_de_colunas
		 .Cells(d).Width = largura_colunas					
		next
	End With
	Notas_Tit(1, 1).AddText "<div align=""center""><b>N&ordm;</b></div>", "size=10;indenty=2; html=true", Font 
	Notas_Tit(1, 2).AddText "<div align=""center""><b>Nome</b></div>", "size=10;alignment=center; indenty=2;html=true", Font 
	Notas_Tit(1, 3).AddText "<div align=""center""><b>Per</b></div>", "size=10;alignment=center; indenty=2;html=true", Font 
		tabela_col=3
		for d=0 to ubound(vet_co_materia)
			tabela_col=tabela_col+1
			Notas_Tit(1, tabela_col).AddText "<div align=""center""><b>"&vet_co_materia(d)&"</b></div>", "size=10; indenty=2; alignment=center; html=true", Font
		next			
	Set param_materias = PDF.CreateParam	
	param_materias.Set "size=8;expand=true" 			
										
	
	linha=1
	
	for aln=0 to ubound(n_alunos)	

		param_materias.Add "indenty=1;alignment=right;html=true"

		aluno = split(n_alunos(aln), "#!#" )
		cod_cons=aluno(0)
		num_cham=aluno(1)
		no_aluno=aluno(2)				

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "Select * from TB_Mapao_Notas WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"' and CO_Matricula="&cod_cons&" ORDER BY NU_Seq_Per"
		Set RS1 = CONt.Execute(SQL1)		
		
		while not RS1.EOF
		
			param_materias.Add "indentx=5"			
			if no_aluno=no_aluno_anterior then
				mudou_aluno="n"
			else	
				mudou_aluno="s"
				no_aluno_anterior=no_aluno
			end if			
'		for prd=0 to ubound(nom_periodo)
			conta_notas=1
			vetor_nota_exibe=""
			seq_per=RS1("NU_Seq_Per")
			no_exibe_per=RS1("CO_Per")
			periodo_real=RS1("NU_Seq_Per_Real")
			for conta_notas=1 to ubound(vet_co_materia)+1
				if conta_notas<10 then
					campo="CO_0"&conta_notas
				else
					campo="CO_"&conta_notas			
				end if
				
				val_nota=RS1(campo)
				if conta_notas=1 then
					vetor_nota_exibe=val_nota
				else	
					vetor_nota_exibe=vetor_nota_exibe&"#!#"&val_nota
				end if
				vetor_nota_separa=vetor_nota_exibe	
			next
			vetor_nota=split(vetor_nota_exibe,"#!#")
			seq_per=seq_per*1
			
			Set Row = Notas_Tit.Rows.Add(12) ' row height			
			linha=linha+1
			
			if ubound(nom_periodo)=0 then
			
			else
'				if prd=0 then
				IF seq_per=1 then
					Notas_Tit(linha, 1).RowSpan = ubound(nom_periodo)+1	
					Notas_Tit(linha, 2).RowSpan = ubound(nom_periodo)+1				
				end if
			end if	
			'if mudou_aluno="s"then
				Notas_Tit(linha, 1).AddText num_cham, param_materias	
				Notas_Tit(linha, 2).AddText no_aluno, param_materias		
			'else
			'	Notas_Tit(linha, 1).AddText "&nbsp;", param_materias	
			'	Notas_Tit(linha, 2).AddText "&nbsp;", param_materias				
			'end if
			Notas_Tit(linha, 3).AddText "<div align=""center"">"&no_exibe_per&"</div>", "size=8;indenty=1; html=true", Font 			

			coluna=3
			param_materias.Add "indentx=0"
	
				
			for dsc=0 to ubound(vet_co_materia)	
				coluna=coluna+1		
				
				media=vetor_nota(dsc)


				teste = isnumeric(media)			

				if teste=false then
					Notas_Tit(linha, coluna).AddText "<div align=""center"">&nbsp;</div>", param_materias					
				else	
		'			media=media*1	
		'			ntazl=ntazl*1
		'			ntvml=ntvml*1
		
		'			if media>=ntazl then	
		'				response.Write("<font color="&cor_nota_prt&">"&formatnumber(media,1)&"</font>")				
		'			elseif media>=ntvml then	
		'				response.Write("<font color="&cor_nota_azl&">"&formatnumber(media,1)&"</font>")
		'			else	
		'				response.Write("<font color="&cor_nota_vml&">"&formatnumber(media,1)&"</font>")	
		'			end if	
					Notas_Tit(linha, coluna).AddText "<div align=""center"">"&formatnumber(media,0)&"</div>", "size=8;indenty=1; html=true", Font 		
				end if	
			next
'		next
		RS1.MOVENEXT
		WEND
	next
	limite=0
	Do While True
	limite=limite+1
	   LastRow = Page.Canvas.DrawTable( Notas_Tit, param_table2 )

		if LastRow >= Notas_Tit.Rows.Count Then 
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
							   ' Display remaining part of table on the next page
			Set Page = Page.NextPage	
			param_table2.Add( "RowTo=1; RowFrom=1" ) ' Row 1 is header.
			param_table2("RowFrom1") = LastRow + 1 ' RowTo1 is omitted and presumed infinite
'NOVO CABEÇALHO==========================================================================================		
			Set Param_Logo_Gde = Pdf.CreateParam
			margem=25			
			linha=10		
			unidade = unidade*1	
			if unidade = 1 then
				largura_logo_gde=formatnumber(Logo1.Width*0.7,0)
				altura_logo_gde=formatnumber(Logo1.Height*0.7,0)
				area_utilizavel=Page.Width-(margem*2)
				Param_Logo_Gde("x") = margem
				Param_Logo_Gde("y") = Page.Height - altura_logo_gde -margem
				Param_Logo_Gde("ScaleX") = 0.7
				Param_Logo_Gde("ScaleY") = 0.7
				Page.Canvas.DrawImage Logo1, Param_Logo_Gde
			else
				largura_logo_gde=formatnumber(Logo2.Width*0.7,0)
				altura_logo_gde=formatnumber(Logo2.Height*0.7,0)
				area_utilizavel=Page.Width-(margem*2)
				Param_Logo_Gde("x") = margem
				Param_Logo_Gde("y") = Page.Height - altura_logo_gde -margem
				Param_Logo_Gde("ScaleX") = 0.7
				Param_Logo_Gde("ScaleY") = 0.7
				Page.Canvas.DrawImage Logo2, Param_Logo_Gde		
			end if
	
			x_texto=largura_logo_gde+ margem+10
			y_texto=formatnumber(Page.Height - margem,0)
			width_texto=Page.Width -largura_logo_gde - 80
	
		
			SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
				Text = "<p><i><b>"&UCASE(no_unidade)&"</b></i></p><br><font style=""font-size:10pt;"">"&endereco_unidade&"</FONT><br><center><i><b><font style=""font-size:18pt;"">MAP&Atilde;O DE M&Eacute;DIAS POR PER&Iacute;ODO</font></b></i></center>"
			
			
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
			Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 			
'================================================================================================================				 
		end if
		if limite>100 then
		response.Write("ERRO!")
		response.end()
		end if 

		SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
			
	Loop
	

'	RS.Close
'	Set RS = Nothing
	
'	RS_nota.Close
'	Set RS_nota = Nothing
'	
'	RS2.Close
'	Set RS2 = Nothing
'		
'	RS3.Close
'	Set RS3 = Nothing
'
'	RS3m.Close
'	Set RS3m = Nothing
'
'	RS4.Close
'	Set RS4 = Nothing
'
'	RS5.Close
'	Set RS5 = Nothing	
						
Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
else
response.Write("ERRO")

End IF
		response.Redirect("index.asp?nvg=WA-PF-CN-ACC&opt=acc&obr="&obr_mapa)
%>

