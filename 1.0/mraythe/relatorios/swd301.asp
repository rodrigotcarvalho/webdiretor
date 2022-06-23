<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 30 'valor em segundos
'CONTEÚDO LECIONADO
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/parametros.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/funcoes7.asp"-->
<!--#include file="../inc/utils.asp"-->
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

arquivo="SWD301"
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

if ori="ws" then
	unidade = request.Form("unidade")
	curso = request.Form("curso")
	co_etapa = request.Form("etapa")
	turma = request.Form("turma")
	periodo = request.Form("periodo")

	Set RSG = Server.CreateObject("ADODB.Recordset")
	SQLG = "SELECT CO_Materia_Principal, CO_Professor FROM TB_Da_Aula where CO_Professor is not null AND CO_Turma  = '"& turma &"' and CO_Etapa = '"&co_etapa &"' AND NU_Unidade = "&unidade&" and CO_Curso = '"&curso&"'"
	RSG.Open SQLG, CON2

	total_mat=0
	while not RSG.EOF
		co_prof = RSG("CO_Professor")
		co_materia = RSG("CO_Materia_Principal")
		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL_0 = "Select * from TB_Materia WHERE CO_Materia = '"& co_materia &"'"
		Set RS0 = CON0.Execute(SQL_0)
	
		mat_princ=RS0("CO_Materia_Principal")
		
		if mat_princ="" or isnull(mat_princ) then
			mat_princ=co_materia
		end if			
		
		if total_mat=0 then
			vetor_materia = co_materia
			vetor_mat_princ = mat_princ
			vetor_professor = co_prof						
		else
			vetor_materia = vetor_materia&"#!#"&co_materia
			vetor_mat_princ = vetor_mat_princ&"#!#"&mat_princ
			vetor_professor = vetor_professor&"#!#"&co_prof			
		
		end if
	total_mat=total_mat+1
	RSG.MOVENEXT
	wend	
else	
	obr=request.QueryString("obr")
	dados_informados = split(obr, "$!$" )
	co_materia = dados_informados(0)
	unidade = dados_informados(1)
	curso = dados_informados(2)
	co_etapa = dados_informados(3)
	turma = dados_informados(4)
	periodo = dados_informados(5)
	co_prof = dados_informados(7)
	
	Set RS0 = Server.CreateObject("ADODB.Recordset")
	SQL_0 = "Select * from TB_Materia WHERE CO_Materia = '"& co_materia &"'"
	Set RS0 = CON0.Execute(SQL_0)

	mat_princ=RS0("CO_Materia_Principal")
	
	if mat_princ="" or isnull(mat_princ) then
		mat_princ=co_materia
	end if		
	
	vetor_materia = co_materia
	vetor_mat_princ = mat_princ
	vetor_professor = co_prof		
end if	

materias = split(vetor_materia,"#!#")
materias_principais = split(vetor_mat_princ,"#!#")
professores = split(vetor_professor,"#!#")
Paginacao = 1
for tm=0 to ubound(materias)

	co_materia = materias(tm)
	mat_princ=materias_principais(tm)	
	co_prof = professores(tm)	
	
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

		nome_prof = replace_latin_char(nome_prof,"html")		

		tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")		

nome_prof = replace_latin_char(nome_prof,"html")	

nu_chamada_check = 1	

tb_nota=tabela_nota(ano_letivo,unidade,curso,co_etapa,turma,"tb",0)	
bancoPauta = escolheBancoPauta(tb_nota,"M",p_outro)
caminhoBancoPauta = verificaCaminhoBancoPauta(bancoPauta,"M",p_outro)

Set CONPauta = Server.CreateObject("ADODB.Connection") 
ABRIRPauta = "DBQ="& caminhoBancoPauta & ";Driver={Microsoft Access Driver (*.mdb)}"
CONPauta.Open ABRIRPauta

Set RSA = Server.CreateObject("ADODB.Recordset")
CONEXAOA = "Select DT_Aula, TX_Aula, TX_Obs from TB_Materia_Lecionada WHERE CO_Professor  = "& co_prof &" AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Unidade  = "& unidade &" AND CO_Curso  = '"& curso &"' AND CO_Etapa  = '"& co_etapa &"' AND CO_Turma  = '"& turma &"' AND NU_Periodo = "& periodo &" order by DT_Aula"

Set RSA = CONPauta.Execute(CONEXAOA)

vetor_matriculas="" 
obs_check = 0
if RSA.EOF then
	gera_pdf="nao"
else
	While Not RSA.EOF
		DT_Aula = RSA("DT_Aula")
		TX_Aula= RSA("TX_Aula")			
		TX_Obs = RSA("TX_Obs")
	'response.Write(CONEXAOA&"<BR>")	
		obs_check=obs_check*1
		if obs_check = 0 then
			vetor_matriculas=formata(DT_Aula,"DD/MM/YYYY")&"#!#"&TX_Aula&"#!#"&TX_Obs
		else
			vetor_matriculas=vetor_matriculas&"#$#"&formata(DT_Aula,"DD/MM/YYYY")&"#!#"&TX_Aula&"#!#"&TX_Obs
		end if
	obs_check=obs_check+1		
	RSA.MoveNext
	Wend 
end if



if vetor_matriculas="" and gera_pdf<>"sim" then
	gera_pdf="nao"
else

gera_pdf="sim" 

'		ln_pesos_cols=verifica_dados_tabela(opcao,"peso_col",outro)
'		ln_pesos_vars=verifica_dados_tabela(opcao,"peso_bd_var",outro)
'		nm_pesos_vars=verifica_dados_tabela(opcao,"peso_wrk_var",outro)
		ln_nom_cols="Data#!#Conte&uacute;do Lecionado#!#Observa&ccedil;&otilde;es"
'		nm_vars=verifica_dados_tabela(opcao,"wrk_var",outro)
'		nm_bd=verifica_dados_tabela(opcao,"bd_var",outro)
'		vars_calc=verifica_dados_tabela(opcao,"calc",outro)
'		action=verifica_dados_tabela(opcao,"action",outro)
'		notas_a_lancar=verifica_dados_tabela(opcao,"notas_a_lancar",outro)

'		linha_pesos=split(ln_pesos_cols,"#!#")
'		linha_pesos_variaveis=split(ln_pesos_vars,"#!#")
'		nome_pesos_variaveis=split(nm_pesos_vars,"#!#")
		linha_nome_colunas=split(ln_nom_cols,"#!#")
'		nome_variaveis=split(nm_vars,"#!#")
'		variaveis_bd=split(nm_bd,"#!#")	
'		calcula_variavel=split(vars_calc,"#!#")

		
	'end if

	if gera_pdf="sim" then	
	
			
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
			if bairro_unidade="" or isnull(bairro_unidade)then
			else
			
				Set RS3b = Server.CreateObject("ADODB.Recordset")
				SQL3b = "SELECT * FROM TB_Bairros WHERE SG_UF='"& uf_unidade_municipio &"' AND CO_Municipio="&municipio_unidade&" AND CO_Bairro = "&bairro_unidade
				RS3b.Open SQL3b, CON0
				
				bairro_unidade=RS3b("NO_Bairro")				
				bairro_unidade=" - "&bairro_unidade
			end if				
		
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
		
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Periodo where NU_Periodo ="& periodo 
		RS5.Open SQL5, CON0
		
		no_periodo = RS5("NO_Periodo")		
		
		Set RS6 = Server.CreateObject("ADODB.Recordset")
		SQL6 = "SELECT * FROM TB_Materia where CO_Materia='"& co_materia &"'"
		RS6.Open SQL6, CON0
		
		no_materia= RS6("NO_Materia")	
		
'		Set RST = Server.CreateObject("ADODB.Recordset")
'		SQLT = "SELECT * FROM TB_Turma WHERE CO_Turma='"& turma &"'"
'		RST.Open SQLT, CON0	
'		
'		no_auxiliares = RST("NO_Auxiliares")		
		
		if sexo_prof = "M" then
			profoa = "Professor"
		else		
			profoa = "Professora"
		end if	
		no_curso= no_etapa&"&nbsp;"&co_concordancia_curso&"&nbsp;"&no_curso
		texto_turma = "&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;<b>Turma</b> "&turma
		texto_disciplina = "&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;<b>Disciplina:</b> "&no_materia
		texto_periodo = "&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;<b>Per&iacute;odo:</b> "&no_periodo
		mensagem_cabecalho=ano_letivo&"&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;"&no_curso&texto_turma&texto_disciplina&texto_periodo&"&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;<b>"&profoa&":</b> "&nome_prof
		if co_materia="INT" then
			mensagem_cabecalho = mensagem_cabecalho&"&nbsp;&nbsp;-&nbsp;&nbsp;<b>Auxiliar(es): </b>"&no_auxiliares
		end if			


		SET Page = Doc.Pages.Add(842, 595)
				
'CABEÇALHO==========================================================================================		
		Set Param_Logo_Gde = Pdf.CreateParam
		margem=25			
		area_utilizavel=Page.Width - (margem*2)
		largura_logo_gde=formatnumber(Logo.Width*0.3,0)
		altura_logo_gde=formatnumber(Logo.Height*0.3,0)
		
		Param_Logo_Gde("x") = margem
		Param_Logo_Gde("y") = Page.Height - altura_logo_gde - margem
		Param_Logo_Gde("ScaleX") = 0.3
		Param_Logo_Gde("ScaleY") = 0.3
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
				Paginacao = Paginacao+1					
			Text = Right( Text, Len(Text) - CharsPrinted)
		Loop 

		y_texto=y_texto-altura_logo_gde+10
		SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height=30; width=500; alignment=center; size=14; color=#000000; html=true")
		Text = "<center><i><b><font style=""font-size:18pt;"">Di&aacute;rio de Classe - Conte&uacute;do Lecionado</font></b></i></center>"
				
		Do While Len(Text) > 0
			CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
		 
			If CharsPrinted = Len(Text) Then Exit Do
				SET Page = Page.NextPage
				Paginacao = Paginacao+1					
			Text = Right( Text, Len(Text) - CharsPrinted)
		Loop 
		
		Page.Canvas.SetParams "LineWidth=1" 
		Page.Canvas.SetParams "LineCap=0" 

		altura_assinatura= Page.Height - margem-50		

		SET Param = Pdf.CreateParam("x=550;y="&altura_assinatura&"; height=30; width=230; alignment=center; size=8; color=#000000; html=true")
		Text = "<center>Assinatura do Professor</center>"
		
		
		Do While Len(Text) > 0
			CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
		 
			If CharsPrinted = Len(Text) Then Exit Do
				SET Page = Page.NextPage
				Paginacao = Paginacao+1					
			Text = Right( Text, Len(Text) - CharsPrinted)
		Loop 
		
		

		With Page.Canvas
		   .MoveTo 550, altura_assinatura
		   .LineTo 780, altura_assinatura
		   .Stroke
		End With 		

		
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
		altura_segundo_separador= Page.Height - altura_logo_gde-margem - 20
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
		   .Cells(2).Width = area_utilizavel-100
		   .Cells(3).Width = 50	
		End With
		
		Table(1, 1).AddText "<b>Ano Letivo:</b>", "size=8;html=true", Font 
		Table(1, 2).AddText mensagem_cabecalho, "size=8;html=true", Font	
		'Table(1, 3).AddText "<div align=""right""><b>Legenda:</b> Md=M&eacute;dia - Res=Resultado&nbsp;&nbsp;&nbsp;</div>", "size=8;html=true", Font	
		Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 		

'================================================================================================================			
		colunas_de_notas=2
		total_de_colunas=colunas_de_notas+1			
		altura_medias=20
		y_segunda_tabela=y_primeira_tabela-20	
		Set param_table2 = Pdf.CreateParam("width="&area_utilizavel&"; height="&altura_medias&"; rows=1; cols="&total_de_colunas&"; border=1; cellborder=0.1; cellspacing=0; x="&margem&"; y="&y_segunda_tabela&"; MaxHeight=420")

		Set Notas_Tit = Doc.CreateTable(param_table2)
		Notas_Tit.Font = Font				
		largura_colunas=(area_utilizavel-50)/colunas_de_notas		
		With Notas_Tit.Rows(1)
		   .Cells(1).Width = 50		             
			for d=2 to total_de_colunas
			 .Cells(d).Width = largura_colunas					
			next
		End With
				
		alunos_encontrados = split(vetor_matriculas, "#$#" )


		linha=1
		fim_do_cabecalho=1	
		tabela_col=1
		for e=0 to ubound(linha_nome_colunas)
			Notas_Tit(linha, tabela_col).AddText "<div align=""center""><b>"&linha_nome_colunas(e)&"</b></div>", "size=9; indenty=3; alignment=center; html=true", Font				
			tabela_col=tabela_col+1
		next			
		Set param_materias = PDF.CreateParam	
		param_materias.Set "size=8;expand=true" 			
												
		
		conta_notas = 1 
		
		nu_chamada_ckq = 0
		
		for b=0 to ubound(alunos_encontrados)	
			param_materias.Add "indenty=2;alignment=right;html=true"
			param_materias.Add "indentx=5"	
			dados_alunos = split(alunos_encontrados(b), "#!#" )		 
			linha=linha+1
			Set Row = Notas_Tit.Rows.Add(15) ' row height						
			Notas_Tit(linha, 1).AddText dados_alunos(0), param_materias	
			Notas_Tit(linha, 2).AddText dados_alunos(1), param_materias	
			Notas_Tit(linha, 3).AddText dados_alunos(2), param_materias						
		next

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
						Paginacao = Paginacao+1						   
					Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
				Loop 
				
				Param_Relatorio.Add "alignment=right" 
				

 								
				Do While Len(Paginacao) > 0
					CharsPrinted = Page.Canvas.DrawText(Paginacao, Param_Relatorio, Font )
				 
					If CharsPrinted = Len(Paginacao) Then Exit Do
					   SET Page = Page.NextPage
					Paginacao = Right( Paginacao, Len(Relatorio) - CharsPrinted)
				Loop 
				
				
				Param_Relatorio.Add "html=true" 
				
				data_hora = "<center>Impresso em "&data &" &agrave;s "&horario&"</center>"
				Do While Len(data_hora) > 0
					CharsPrinted = Page.Canvas.DrawText(data_hora, Param_Relatorio, Font )
				 
					If CharsPrinted = Len(data_hora) Then Exit Do
					   SET Page = Page.NextPage
						Paginacao = Paginacao+1						   
					data_hora = Right( data_hora, Len(data_hora) - CharsPrinted)
				Loop 				   ' Display remaining part of table on the next page
				Set Page = Page.NextPage	
				Paginacao = Paginacao+1						
				param_table2.Add( "RowTo="&fim_do_cabecalho&"; RowFrom=1" ) ' Row 1 is header.
				param_table2("RowFrom1") = LastRow + 1 ' RowTo1 is omitted and presumed infinite
	'NOVO CABEÇALHO==========================================================================================		
			Set Param_Logo_Gde = Pdf.CreateParam
			margem=25			
			area_utilizavel=Page.Width - (margem*2)
			largura_logo_gde=formatnumber(Logo.Width*0.3,0)
			altura_logo_gde=formatnumber(Logo.Height*0.3,0)
			
			Param_Logo_Gde("x") = margem
			Param_Logo_Gde("y") = Page.Height - altura_logo_gde - margem
			Param_Logo_Gde("ScaleX") = 0.3
			Param_Logo_Gde("ScaleY") = 0.3
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
					Paginacao = Paginacao+1						
				Text = Right( Text, Len(Text) - CharsPrinted)
			Loop 
	
			y_texto=y_texto-altura_logo_gde+10
			SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height=30; width=500; alignment=center; size=14; color=#000000; html=true")
			Text = "<center><i><b><font style=""font-size:18pt;"">Di&aacute;rio de Classe - Conte&uacute;do Lecionado</font></b></i></center>"
					
			Do While Len(Text) > 0
				CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
			 
				If CharsPrinted = Len(Text) Then Exit Do
					SET Page = Page.NextPage
					Paginacao = Paginacao+1						
				Text = Right( Text, Len(Text) - CharsPrinted)
			Loop 
			
			Page.Canvas.SetParams "LineWidth=1" 
			Page.Canvas.SetParams "LineCap=0" 
	
			altura_assinatura= Page.Height - margem-50		
	
			SET Param = Pdf.CreateParam("x=550;y="&altura_assinatura&"; height=30; width=230; alignment=center; size=8; color=#000000; html=true")
			Text = "<center>Assinatura do Professor</center>"
			
			
			Do While Len(Text) > 0
				CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
			 
				If CharsPrinted = Len(Text) Then Exit Do
					SET Page = Page.NextPage
					Paginacao = Paginacao+1						
				Text = Right( Text, Len(Text) - CharsPrinted)
			Loop 
			
			
	
			With Page.Canvas
			   .MoveTo 550, altura_assinatura
			   .LineTo 780, altura_assinatura
			   .Stroke
			End With 		
	
			
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
			altura_segundo_separador= Page.Height - altura_logo_gde-margem - 20
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
			   .Cells(2).Width = area_utilizavel-100
			   .Cells(3).Width = 50	
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
			
			SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem*2&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
			
			'Relatorio = "Total de Aulas Previstas: ___________"
			Relatorio = ""
			Do While Len(Relatorio) > 0
				CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )			
				If CharsPrinted = Len(Relatorio) Then Exit Do
				SET Page = Page.NextPage
				Paginacao = Paginacao+1					
				Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
			Loop 
			
			Param_Relatorio.Add "alignment=right;" 
			
	Set RSPeriodo = Server.CreateObject("ADODB.Recordset")
	SQLPeriodo = "Select * from TB_Periodo WHERE NU_Periodo= "&periodo
	Set RSPeriodo = CON0.Execute(SQLPeriodo)
	
	dataFim = RSPeriodo("DA_Fim_Periodo")					
			
			Relatorio = "Encerrado em: "&formata(dataFim, "DD/MM/YYYY")
			Do While Len(Relatorio) > 0
				CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )			
				If CharsPrinted = Len(Relatorio) Then Exit Do
				SET Page = Page.NextPage
				Paginacao = Paginacao+1					
				Paginacao = Right( Relatorio, Len(Relatorio) - CharsPrinted)
			Loop 
								
			'Param_Relatorio.Add "html=true;size=12;" 
			
			'Relatorio = "<center>Total de Aulas Dadas: ___________</center>"
			'Relatorio = "<center>Encerrado em: "&formata(dataFim, "DD/MM/YYYY")&"</center>"
			Do While Len(Relatorio) > 0
				CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )			
				If CharsPrinted = Len(Relatorio) Then Exit Do
				SET Page = Page.NextPage
				data_hora = Right( Relatorio, Len(Relatorio) - CharsPrinted)
			Loop 		
			Param_Relatorio.Add "size=8;" 				
			
			SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")		
			
			Relatorio = arquivo&" - Sistema Web Diretor"
			Do While Len(Relatorio) > 0
				CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )			
				If CharsPrinted = Len(Relatorio) Then Exit Do
				SET Page = Page.NextPage
				Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
			Loop 
			
			Param_Relatorio.Add "alignment=right" 
			

			Do While Len(Paginacao) > 0
				CharsPrinted = Page.Canvas.DrawText(Paginacao, Param_Relatorio, Font )			
				If CharsPrinted = Len(Paginacao) Then Exit Do
				SET Page = Page.NextPage
				Paginacao = Right( Paginacao, Len(Paginacao) - CharsPrinted)
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
	End IF		
					
Next
	

Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

