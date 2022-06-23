<%'On Error Resume Next%>
<%
'Server.ScriptTimeout = 60
Server.ScriptTimeout = 600 'valor em segundos
'Lista de Reunião
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/funcoes2.asp"-->

<% 

response.Charset="ISO-8859-1"
opt= request.QueryString("opt")
ori= request.QueryString("ori")

unidade_form = request.Form("unidade")
curso_form  = request.Form("curso")
etapa_form  = request.Form("etapa")
turma_form  = request.Form("turma")

dr= request.Form("dr")
mr= request.Form("mr")
ar= request.Form("ar")

dr=dr*1

if dr<10 then
	dia_reuniao="0"&dr
else
	dia_reuniao=dr
end if

data_reuniao = dia_reuniao &"/"& mr &"/"& ar

motivo= request.Form("motivo")

unidade_form=unidade_form*1
if unidade_form=999990 then
	sql_ucet=""	
	unidade_form=1
else
	sql_ucet="AND TB_Matriculas.NU_Unidade = "& unidade_form&" "
	if isnumeric(curso_form) then
		if curso_form<>999990 then
			sql_ucet=sql_ucet&"AND TB_Matriculas.CO_Curso ='"& curso_form&"' "
			if isnumeric(etapa_form) then
				if etapa_form<>999990 then
					sql_ucet=sql_ucet&"AND TB_Matriculas.CO_Etapa = '"& etapa_form&"' "
				end if	
				if isnumeric(turma_form) then
					if turma_form<>999990 then
						sql_ucet=sql_ucet&"AND TB_Matriculas.CO_Turma = '"& turma_form&"' "
					end if								
				end if	
			else
				if etapa_form<>"999990" then
					sql_ucet=sql_ucet&"AND TB_Matriculas.CO_Etapa = '"& etapa_form&"' "
				end if	
				if isnumeric(turma_form) then
					if turma_form<>"999990" then
						sql_ucet=sql_ucet&"AND TB_Matriculas.CO_Turma = '"& turma_form&"' "
					end if						
				else
					if turma_form<>"999990" then
						sql_ucet=sql_ucet&"AND TB_Matriculas.CO_Turma = '"& turma_form&"' "
					end if								
				end if					
			end if	
		end if			
	end if		
end if	

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

	
	Set RSA = Server.CreateObject("ADODB.Recordset")
	CONEXAOA = "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada, TB_Matriculas.CO_Situacao, TB_Matriculas.NU_Unidade, TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma, TB_Matriculas.DA_Rematricula, TB_Alunos.NO_Aluno, TB_Alunos.SG_UF_Natural,TB_Alunos.CO_Municipio_Natural, TB_Alunos.NO_Pai, TB_Alunos.NO_Mae from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" "&sql_ucet&"AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula order by TB_Matriculas.NU_Unidade ASC, TB_Matriculas.CO_Curso ASC, TB_Matriculas.CO_Etapa ASC, TB_Matriculas.CO_Turma ASC, TB_Alunos.NO_Aluno ASC"
	Set RSA = CON1.Execute(CONEXAOA)


	vetor_matriculas="" 
	nu_seq_aluno=0
	nu_chamada_conta=1	
	While Not RSA.EOF
		nu_seq_aluno=nu_seq_aluno+1
		nu_matricula = RSA("CO_Matricula")
		nome_aluno= RSA("NO_Aluno")			
		nu_chamada = RSA("NU_Chamada")
		co_situacao = RSA("CO_Situacao")
		unidade_aluno =	RSA("NU_Unidade")	
		curso_aluno =	RSA("CO_Curso")
		etapa_aluno =	RSA("CO_Etapa")
		turma_aluno =	RSA("CO_Turma")
		dt_matricula= RSA("DA_Rematricula")
		uf_natural= RSA("SG_UF_Natural")
		cidade_natural= RSA("CO_Municipio_Natural")
		no_pai = RSA("NO_Pai")	
		no_mae = RSA("NO_Mae")		
				
		data_m=split(dt_matricula,"/")
		if data_m(0)<10 then
			dia_m="0"&data_m(0)
		else	
			dia_m=data_m(0)
		end if		

		if data_m(1)<10 then
			mes_m="0"&data_m(1)
		else	
			mes_m=data_m(1)
		end if				
		
		dt_matricula=dia_m&"/"&mes_m&"/"&data_m(2)
		
		if co_situacao="C" then
			no_situacao="Efetivado"
		else
			no_situacao=co_situacao	
		end if	
		
		Set RS3n= Server.CreateObject("ADODB.Recordset")
		SQL3n = "SELECT * FROM TB_Municipios WHERE SG_UF='"& uf_natural &"' AND CO_Municipio="&cidade_natural
		RS3n.Open SQL3n, CON0
		
		municipio_natural=RS3n("NO_Municipio")						
		natural=municipio_natural&" - "&uf_natural
		
		nome_aluno=replace_latin_char(nome_aluno,"html")	
		if isnull(no_pai) or no_pai="" then
		else
			no_pai=replace_latin_char(no_pai,"html")	
		end if	
		if isnull(no_mae) or no_mae="" then		
		else		
			no_mae=replace_latin_char(no_mae,"html")	
		end if	
		
		nu_chamada_conta=nu_chamada_conta*1
		if nu_chamada_conta = 1 then
			vetor_matriculas=nu_seq_aluno&"#!#"&nu_matricula&"#!#"&nu_chamada&"#!#"&nome_aluno&"#!#"&no_situacao&"#!#"&unidade_aluno&"#!#"&curso_aluno&"#!#"&etapa_aluno&"#!#"&turma_aluno&"#!#"&dt_matricula&"#!#"&natural&"#!#"&no_pai&"#!#"&no_mae
		else
			vetor_matriculas=vetor_matriculas&"#$#"&nu_seq_aluno&"#!#"&nu_matricula&"#!#"&nu_chamada&"#!#"&nome_aluno&"#!#"&no_situacao&"#!#"&unidade_aluno&"#!#"&curso_aluno&"#!#"&etapa_aluno&"#!#"&turma_aluno&"#!#"&dt_matricula&"#!#"&natural&"#!#"&no_pai&"#!#"&no_mae
		end if
	nu_chamada_conta=nu_chamada_conta+1		
	RSA.MoveNext
	Wend 

	Set RS2 = Server.CreateObject("ADODB.Recordset")
	SQL2 = "SELECT * FROM TB_Unidade WHERE NU_Unidade="&unidade_form
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
	
	if municipio_unidade="" or isnull(municipio_unidade) or uf_unidade_municipio="" or isnull(uf_unidade_municipio) then
	else
		Set RS3m = Server.CreateObject("ADODB.Recordset")
		SQL3m = "SELECT * FROM TB_Municipios WHERE SG_UF='"& uf_unidade_municipio &"' AND CO_Municipio="&municipio_unidade
		RS3m.Open SQL3m, CON0
		cod_municipio = municipio_unidade
		municipio_unidade=RS3m("NO_Municipio")						
	end if
	
	if municipio_unidade="" or isnull(municipio_unidade) or uf_unidade_municipio="" or isnull(uf_unidade_municipio) or bairro_unidade="" or isnull(bairro_unidade)then
	else
		Set RSb = Server.CreateObject("ADODB.Recordset")
		SQLb = "SELECT * FROM TB_Bairros WHERE SG_UF='"& uf_unidade_municipio&"' AND CO_Municipio ="&cod_municipio&" AND CO_Bairro="&bairro_unidade		
		RSb.Open SQLb, CON0			
		
		bairro_unidade=" - "&RSb("NO_Bairro")			
	end if			
	endereco_unidade=rua_unidade&numero_unidade&complemento_unidade&bairro_unidade&cep_unidade&"<br>"&municipio_unidade&uf_unidade					
						

'
'			no_curso= no_etapa&"&nbsp;"&co_concordancia_curso&"&nbsp;"&no_curso
'			texto_turma = "&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;<b>Turma</b> "&turma
'
'			mensagem_cabecalho=ano_letivo&"&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;"&no_curso&texto_turma
'	
	SET Page = Doc.Pages.Add(595,842)
			
'CABEÇALHO==========================================================================================		
	Set Param_Logo_Gde = Pdf.CreateParam
	margem=25			
	area_utilizavel=Page.Width - (margem*2)
	
	largura_logo_gde=formatnumber(Logo.Width*0.4,0)
	altura_logo_gde=formatnumber(Logo.Height*0.4,0)

   Param_Logo_Gde("x") = margem
   Param_Logo_Gde("y") = Page.Height - altura_logo_gde - margem
   Param_Logo_Gde("ScaleX") = 0.4
   Param_Logo_Gde("ScaleY") = 0.4
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
	Text = "<center><i><b><font style=""font-size:18pt;"">Lista de Reuni&atilde;o</font></b></i></center>"
	
	
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

	Set param_table1 = Pdf.CreateParam("width=547; height=20; rows=1; cols=4; border=0; cellborder=0; cellspacing=0;")
	Set Table = Doc.CreateTable(param_table1)
	Table.Font = Font
	y_primeira_tabela=altura_segundo_separador-10
	x_primeira_tabela=margem+5
	With Table.Rows(1)
	   .Cells(1).Width = 50
	   .Cells(2).Width = 150  
	   .Cells(3).Width = 50 		   		   		   
	End With
	
	
	Table(1, 1).AddText "<b>Data:</b>", "size=9;html=true", Font 
	Table(1, 2).AddText data_reuniao, "size=9;html=true", Font 
	Table(1, 3).AddText "<b>Motivo:</b>", "size=9;html=true", Font 
	Table(1, 4).AddText "<div align=LEFT>"&motivo&"</div>", "size=9;html=true", Font 
	Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 	
	
	Page.Canvas.SetParams "LineWidth=2" 
	Page.Canvas.SetParams "LineCap=0" 
	altura_terceiro_separador= y_primeira_tabela-20
	With Page.Canvas
	   .MoveTo margem, altura_terceiro_separador
	   .LineTo area_utilizavel+margem, altura_terceiro_separador
	   .Stroke
	End With 			

'================================================================================================================			

	colunas_de_notas=2
	total_de_colunas=5					
	altura_medias=20
	y_segunda_tabela=altura_terceiro_separador-10	
	Set param_table2 = Pdf.CreateParam("width=547; height="&altura_medias&"; rows=1; cols=5; border=0; cellborder=0.5; cellspacing=0; x="&margem&"; y="&y_segunda_tabela&"; MaxHeight=670")

	Set Notas_Tit = Doc.CreateTable(param_table2)
	Notas_Tit.Font = Font				
	largura_colunas=(547-20-45-120)/colunas_de_notas		
	
	With Notas_Tit.Rows(1)
	   .Cells(1).Width = 20
	   .Cells(2).Width = 45	
	   .Cells(3).Width = 120
	   .Cells(4).Width = largura_colunas			             
	   .Cells(5).Width = largura_colunas
	End With
	Notas_Tit(1, 1).AddText "<div align=""center""><b>N&ordm;</b></div>", "size=9;indenty=2; html=true", Font 
	Notas_Tit(1, 2).AddText "<div align=""center""><b>Matr&iacute;cula</b></div>", "size=9;alignment=center; indenty=2;html=true", Font 
	Notas_Tit(1, 3).AddText "<div align=""center""><b>Nome do Aluno</b></div>", "size=9;alignment=center; indenty=2;html=true", Font 
	Notas_Tit(1, 4).AddText "<div align=""center""><b>Nome do Respons&aacute;vel</b></div>", "size=9;alignment=center; indenty=2;html=true", Font 
	Notas_Tit(1, 5).AddText "<div align=""center""><b>Assinatura do Respons&aacute;vel</b></div>", "size=9;alignment=center; indenty=2;html=true", Font 

	Set param_materias = PDF.CreateParam	
	param_materias.Set "size=7;expand=false" 			
''response.Flush()										
	alunos_encontrados = split(vetor_matriculas, "#$#" )		
	linha=1
	for a=0 to ubound(alunos_encontrados)	
		param_materias.Add "indenty=2;alignment=right;html=true"
		param_materias.Add "indentx=0"	
		dados_alunos = split(alunos_encontrados(a), "#!#" )
		
		
	Set RS3 = Server.CreateObject("ADODB.Recordset")
	SQL3 = "SELECT * FROM TB_Curso WHERE CO_Curso='"& dados_alunos(6) &"'"
	RS3.Open SQL3, CON0
	
	if RS3.EOF then
		no_curso_aluno=dados_alunos(6) 			
	else	
		no_curso_aluno= RS3("NO_Curso")
	end if	
'			no_abrv_curso = RS3("NO_Abreviado_Curso")
'			co_concordancia_curso = RS3("CO_Conc")	
	
	Set RS4 = Server.CreateObject("ADODB.Recordset")
	SQL4 = "SELECT * FROM TB_Etapa WHERE CO_Curso='"& dados_alunos(6) &"' AND CO_Etapa ='"& dados_alunos(7) &"'"
	RS4.Open SQL4, CON0			

	if RS4.EOF then
		no_etapa_aluno = dados_alunos(7) 			
	else	
		no_etapa_aluno = RS4("NO_Etapa")
	end if	
	
	Set RS5 = Server.CreateObject("ADODB.Recordset")
	SQL5 = "SELECT * FROM TB_Turma WHERE CO_Curso='"& dados_alunos(6) &"' AND CO_Etapa ='"& dados_alunos(7) &"' AND CO_Turma ='"& dados_alunos(8) &"'"
	RS5.Open SQL5, CON0			

	if RS5.EOF then
		co_turno_aluno = ""			
	else	
		co_turno_aluno = RS5("CO_Turno")
		
		Set RS6 = Server.CreateObject("ADODB.Recordset")
		SQL6 = "SELECT * FROM TB_Turno WHERE CO_Turno='"& co_turno_aluno &"'"
		RS6.Open SQL6, CON0				
		if RS6.EOF then
			no_turno_aluno = ""			
		else	
			no_turno_aluno = RS6("NO_Turno")
			'no_turno_aluno = LEFT(no_turno_aluno,1)
		end if	
	end if			
	
'	Set RS7 = Server.CreateObject("ADODB.Recordset")
'	SQL7 = "SELECT * FROM TB_Contatos WHERE CO_Matricula="& dados_alunos(1) &" AND TP_Contato='ALUNO'"
'	RS7.Open SQL7, CONCONT	
'				
'	if RS7.EOF then		
'	else
'		nasc_aluno = RS7("DA_Nascimento_Contato")	
'		rua_aluno = RS7("NO_Logradouro_Res")
'		rua_num_aluno = RS7("NU_Logradouro_Res")				
'		compl_aluno = RS7("TX_Complemento_Logradouro_Res")	
'		bairro_aluno = RS7("CO_Bairro_Res")				
'		cep_aluno = RS7("CO_CEP_Res")		
'		
'		data_n=split(nasc_aluno,"/")
'		if data_n(0)<10 then
'			dia_n="0"&data_n(0)
'		else	
'			dia_n=data_n(0)
'		end if		
'
'		if data_n(1)<10 then
'			mes_n="0"&data_n(1)
'		else	
'			mes_n=data_n(1)
'		end if				
'		
'		nasc_aluno=dia_n&"/"&mes_n&"/"&data_n(2)						
'	
'		Set RS7a = Server.CreateObject("ADODB.Recordset")
'		SQL7a = "SELECT * FROM TB_Bairros WHERE CO_Bairro="& bairro_aluno
'		RS7a.Open SQL7a, CON0					
'		cidade_aluno = RS7a("CO_Municipio")
'		no_bairro_aluno = RS7a("NO_Bairro")
'		uf_aluno = RS7a("SG_UF")
'		rua_cep=left(cep_aluno,5)&"-"&right(cep_aluno,3)
'
'		Set RS7b = Server.CreateObject("ADODB.Recordset")
'		SQL7b = "SELECT * FROM TB_Municipios WHERE CO_Municipio="& cidade_aluno
'		RS7b.Open SQL7b, CON0					
'		no_cidade_aluno = RS7b("NO_Municipio")
'			
'		endereco_aluno=	rua_aluno&", "&rua_num_aluno&", "&compl_aluno&". "&no_bairro_aluno&", "&no_cidade_aluno&" - "&uf_aluno&". CEP: "&rua_cep
'	end if	
'	
'	pai_aluno = dados_alunos(11) 
'	mae_aluno = dados_alunos(12) 				
'	
'	Set RS8 = Server.CreateObject("ADODB.Recordset")
'	SQL8 = "SELECT * FROM TB_Contatos WHERE CO_Matricula="& dados_alunos(1) &" AND TP_Contato='PAI'"
'	RS8.Open SQL8, CONCONT	
'				
'	if RS8.EOF then		
'	else
'		'pai_aluno = RS8("NO_Contato")	
'		ocupacao_p = RS8("CO_Ocupacao")	
'
'		if isnull(ocupacao_p) or ocupacao_p="" then
'			ocupacao_pai = ""
'		else
'			Set RS8a = Server.CreateObject("ADODB.Recordset")
'			SQL8a = "SELECT * FROM TB_Ocupacoes WHERE CO_Ocupacao="& ocupacao_p
'			RS8a.Open SQL8a, CON0	
'			ocupacao_p_n = RS8a("NO_Ocupacao")
'			ocupacao_pai = " ("&ocupacao_p_n&")"
'		end if	
'							
'	end if						
'	
'	Set RS9 = Server.CreateObject("ADODB.Recordset")
'	SQL9 = "SELECT * FROM TB_Contatos WHERE CO_Matricula="& dados_alunos(1) &" AND TP_Contato='MAE'"
'	RS9.Open SQL9, CONCONT	
'				
'	if RS9.EOF then		
'	else
'		'mae_aluno = RS9("NO_Contato")	
'		ocupacao_m = RS9("CO_Ocupacao")
'		if isnull(ocupacao_m) or ocupacao_m="" then
'			ocupacao_mae = ""
'		else
'			Set RS8a = Server.CreateObject("ADODB.Recordset")
'			SQL8a = "SELECT * FROM TB_Ocupacoes WHERE CO_Ocupacao="& ocupacao_m
'			RS8a.Open SQL8a, CON0	
'			ocupacao_m_n = RS8a("NO_Ocupacao")
'			ocupacao_mae = " ("&ocupacao_m_n&")"
'		end if	
'								
'	end if				
	
									
		linha=linha+1
		Set Row = Notas_Tit.Rows.Add(17) ' row height	
		
		param_materias.Add "expand=true" 											
		Notas_Tit(linha, 1).AddText "<div align=""center"">"&dados_alunos(2)&"</div>", param_materias			
		Notas_Tit(linha, 2).AddText "<div align=""center"">"&dados_alunos(1)&"</div>", param_materias			
		Notas_Tit(linha, 3).AddText "<div align=""center"">"&dados_alunos(3)&"</div>", param_materias
		param_materias.Add "expand=false" 	
		Notas_Tit(linha, 4).AddText "<div align=""center""></div>", param_materias	
		Notas_Tit(linha, 5).AddText "<div align=""center""></div>", param_materias																						
	Next		
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
			
			Relatorio = "SWD051 - Sistema Web Diretor"
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
			Loop 				   ' Display remaining part of table on the next page
			Set Page = Page.NextPage	
			
			param_table2.Add( "RowTo=1; RowFrom=1" ) ' Row 1 is header.
			param_table2("RowFrom1") = LastRow + 1 ' RowTo1 is omitted and presumed infinite
'NOVO CABEÇALHO==========================================================================================		
	Set Param_Logo_Gde = Pdf.CreateParam
	margem=25			
	area_utilizavel=Page.Width - (margem*2)
	
	largura_logo_gde=formatnumber(Logo.Width*0.4,0)
	altura_logo_gde=formatnumber(Logo.Height*0.4,0)

   Param_Logo_Gde("x") = margem
   Param_Logo_Gde("y") = Page.Height - altura_logo_gde - margem
   Param_Logo_Gde("ScaleX") = 0.4
   Param_Logo_Gde("ScaleY") = 0.4
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
	Text = "<center><i><b><font style=""font-size:18pt;"">Lista de Reuni&atilde;o</font></b></i></center>"
	
	
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

	Set param_table1 = Pdf.CreateParam("width=547; height=20; rows=1; cols=4; border=0; cellborder=0; cellspacing=0;")
	Set Table = Doc.CreateTable(param_table1)
	Table.Font = Font
	y_primeira_tabela=altura_segundo_separador-10
	x_primeira_tabela=margem+5
	With Table.Rows(1)
	   .Cells(1).Width = 50
	   .Cells(2).Width = 150  
	   .Cells(3).Width = 50 		   		   		   
	End With
	
	
	Table(1, 1).AddText "<b>Data:</b>", "size=9;html=true", Font 
	Table(1, 2).AddText data_reuniao, "size=9;html=true", Font 
	Table(1, 3).AddText "<b>Motivo:</b>", "size=9;html=true", Font 
	Table(1, 4).AddText "<div align=LEFT>"&motivo&"</div>", "size=9;html=true", Font 
	Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 	
	
	Page.Canvas.SetParams "LineWidth=2" 
	Page.Canvas.SetParams "LineCap=0" 
	altura_terceiro_separador= y_primeira_tabela-20
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
			
			Relatorio = "SWD051 - Sistema Web Diretor"
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
			
			data_hora = "<center>Impresso em "&data &" &agrave;s "&horario&"</center>"
			Do While Len(Relatorio) > 0
				CharsPrinted = Page.Canvas.DrawText(data_hora, Param_Relatorio, Font )			
				If CharsPrinted = Len(data_hora) Then Exit Do
				SET Page = Page.NextPage
				data_hora = Right( data_hora, Len(data_hora) - CharsPrinted)
			Loop 	
	
								

	

arquivo="SWD051"
Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

