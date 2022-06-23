<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 30 'valor em segundos
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/funcoes7.asp"-->
<!--#include file="../../global/notas_calculos_diversos.asp"-->
<% 
response.Charset="ISO-8859-1"
opt= request.QueryString("opt")
ori= request.QueryString("ori")
obr = request.QueryString("obr")

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

	dados_obr=split(obr,"$!$")
	unidade = dados_obr(0)
	curso = dados_obr(1)
	co_etapa = dados_obr(2)
	turma = dados_obr(3)
	periodo = dados_obr(4)
	
		tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")		

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

	Set RS5 = Server.CreateObject("ADODB.Recordset")
	SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
	RS5.Open SQL5, CON0
	co_materia_check=1
	IF RS5.EOF Then
		vetor_materia_exibe="nulo"
	else
		while not RS5.EOF
			co_mat_fil= RS5("CO_Materia")
			mae= RS5("IN_MAE")
			fil= RS5("IN_FIL")
			in_co= RS5("IN_CO")	
			if mae= TRUE then
				cmae="T"
			else
				cmae="F"
			end if	
			
			if fil= TRUE then
				cfil="T"
			else
				cfil="F"
			end if		
			
			if in_co= TRUE then
				cin_co="T"
			else
				cin_co="F"
			end if						
			
			wrk_tipo_materia=cmae&cfil&cin_co	
				
			if co_materia_check=1 then
				vetor_materia=co_mat_fil
				vetor_tipo_materia=wrk_tipo_materia
			else
				vetor_materia=vetor_materia&"#!#"&co_mat_fil
				vetor_tipo_materia=vetor_tipo_materia&"#!#"&wrk_tipo_materia
			end if
			co_materia_check=co_materia_check+1			
					
		RS5.MOVENEXT
		wend	
	'	response.Write(vetor_materia&"<BR>")
	'	vetor_materia_exibe=programa_aula(vetor_materia, unidade, curso, co_etapa, "nulo")			
		'response.Write(vetor_materia_exibe)
	end if		
	
	co_materia_exibe=Split(vetor_materia,"#!#")	
	tipo_materia_exibe=Split(vetor_tipo_materia,"#!#")				
	
'obr=request.QueryString("obr")
'dados_informados = split(obr, "?" )
'co_materia = dados_informados(0)
'unidade = dados_informados(1)
'curso = dados_informados(2)
'co_etapa = dados_informados(3)
'turma = dados_informados(4)
'periodo = dados_informados(5)
'co_prof = dados_informados(7)

for mt=0 to ubound(co_materia_exibe)
		co_materia=co_materia_exibe(mt)
		
		Set RSp = Server.CreateObject("ADODB.Recordset")
		SQLp = "SELECT * FROM TB_Da_Aula where CO_Materia_Principal='"& co_materia &"'AND NU_Unidade="& unidade &" AND CO_Curso='"& curso &"' AND CO_Etapa='"& co_etapa &"' AND CO_Turma='"& turma &"'"
		RSp.Open SQLp, CON2		
		
		if RSp.EOF then
			if tipo_materia_exibe(mt)= "TTF" or tipo_materia_exibe(mt)= "TFT" then
				gera_pdf="nao"	
			else	
				no_prof="&nbsp;"
			end if	
		else	
			  			
			co_prof = RSp("CO_Professor")
			
			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT * FROM TB_Professor where CO_Professor="& co_prof
			RS1.Open SQL1, CON3
						
			if RS1.EOF then				
				no_prof="&nbsp;"	
			else
				gera_pdf="sim"		
				no_prof= RS1("NO_Professor")
				no_prof= replace_latin_char(no_prof,"html")				
			end if				
		end if	
		
		if gera_pdf="sim" then		
	

 		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL_0 = "Select * from TB_Materia WHERE CO_Materia = '"& co_materia &"'"
		Set RS0 = CON0.Execute(SQL_0)

mat_princ=RS0("CO_Materia_Principal")

if mat_princ="" or isnull(mat_princ) then
	mat_princ=co_materia
end if

nu_chamada_check = 1	

Set RSA = Server.CreateObject("ADODB.Recordset")
'CONEXAOA = "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada, TB_Alunos.NO_Aluno from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Situacao='C' AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.NU_Unidade = "& unidade&" AND TB_Matriculas.CO_Curso = '"& curso &"' AND TB_Matriculas.CO_Etapa = '"& co_etapa &"' AND TB_Matriculas.CO_Turma = '"& turma &"' order by TB_Matriculas.NU_Chamada"
CONEXAOA = "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada, TB_Alunos.NO_Aluno, TB_Matriculas.CO_Situacao from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.NU_Unidade = "& unidade&" AND TB_Matriculas.CO_Curso = '"& curso &"' AND TB_Matriculas.CO_Etapa = '"& co_etapa &"' AND TB_Matriculas.CO_Turma = '"& turma &"' order by TB_Matriculas.NU_Chamada"
Set RSA = CON1.Execute(CONEXAOA)

vetor_matriculas="" 
While Not RSA.EOF
	nu_matricula = RSA("CO_Matricula")
	no_aluno= RSA("NO_Aluno")			
	nu_chamada = RSA("NU_Chamada")
	situac=RSA("CO_Situacao")
	
	if situac<>"C" then
		no_aluno=no_aluno&" - Aluno Inativo"
	end if			
		
	
	strReplacement = Server.URLEncode(nome_aluno)	
	'strReplacement = nome_aluno
	strReplacement = replace(strReplacement,"+"," ")
	strReplacement = replace(strReplacement,"%27","´")
	strReplacement = replace(strReplacement,"%27","'")
	strReplacement = replace(strReplacement,"À,","&Agrave;")
	strReplacement = replace(strReplacement,"Á","&Aacute;")
	strReplacement = replace(strReplacement,"Â","&Acirc;")
	strReplacement = replace(strReplacement,"Ã","&Atilde;")
	strReplacement = replace(strReplacement,"É","&Eacute;")
	strReplacement = replace(strReplacement,"Ê","&Ecirc;")
	strReplacement = replace(strReplacement,"Í","&Iacute;")
	strReplacement = replace(strReplacement,"Ó","&Oacute;")
	strReplacement = replace(strReplacement,"Ô","&Ocirc;")
	strReplacement = replace(strReplacement,"Õ","&Otilde;")
	strReplacement = replace(strReplacement,"Ú","&Uacute;")
	strReplacement = replace(strReplacement,"Ü","&Uuml;")	
	strReplacement = replace(strReplacement,"à","&agrave;")
	strReplacement = replace(strReplacement,"á","&aacute;")
	strReplacement = replace(strReplacement,"â","&acirc;")
	strReplacement = replace(strReplacement,"ã","&atilde;")
	strReplacement = replace(strReplacement,"ç","&ccedil;")
	strReplacement = replace(strReplacement,"é","&eacute;")
	strReplacement = replace(strReplacement,"ê","&ecirc;")
	strReplacement = replace(strReplacement,"í","&iacute;")
	strReplacement = replace(strReplacement,"ó","&oacute;")
	strReplacement = replace(strReplacement,"ô","&ocirc;")
	strReplacement = replace(strReplacement,"õ","&otilde;")
	strReplacement = replace(strReplacement,"ú","&uacute;")
	strReplacement = replace(strReplacement,"ü","&uuml;")
	nome_aluno =strReplacement
	nu_chamada_check=nu_chamada_check*1
	nu_chamada=nu_chamada*1
	if nu_chamada_check = 1 and nu_chamada=nu_chamada_check then
		vetor_matriculas=nu_matricula&"#!#"&nu_chamada&"#!#"&no_aluno
	elseif nu_chamada_check = 1 then
		while nu_chamada_check < nu_chamada
			nu_chamada_check=nu_chamada_check+1
		wend 
		vetor_matriculas=nu_matricula&"#!#"&nu_chamada&"#!#"&no_aluno
	else
		vetor_matriculas=vetor_matriculas&"#$#"&nu_matricula&"#!#"&nu_chamada&"#!#"&no_aluno
	end if
nu_chamada_check=nu_chamada_check+1		
RSA.MoveNext
Wend 

if curso=0 then
	gera_pdf="nao"
else

	Set RStabela = Server.CreateObject("ADODB.Recordset")
	SQLtabela = "SELECT * FROM TB_Da_Aula where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'" 
	RStabela.Open SQLtabela, CON2

	if 	RStabela.EOF then
		gera_pdf="nao"
	else				
		tb_nota=RStabela("TP_Nota")		
		if tb_nota ="TB_NOTA_A" then
			CAMINHOn = CAMINHO_na	
			opcao="A"	
		elseif tb_nota="TB_NOTA_B" then
			CAMINHOn = CAMINHO_nb			
			opcao="B"	
		elseif tb_nota ="TB_NOTA_C" then
			CAMINHOn = CAMINHO_nc	
			opcao="C"	
		end if		


		ln_pesos_cols=verifica_dados_tabela(opcao,"peso_col",outro)
		ln_pesos_vars=verifica_dados_tabela(opcao,"peso_bd_var",outro)
		nm_pesos_vars=verifica_dados_tabela(opcao,"peso_wrk_var",outro)
		ln_nom_cols=verifica_dados_tabela(opcao,"nome_cols",outro)
		nm_vars=verifica_dados_tabela(opcao,"wrk_var",outro)
		nm_bd=verifica_dados_tabela(opcao,"bd_var",outro)
		vars_calc=verifica_dados_tabela(opcao,"calc",outro)
		action=verifica_dados_tabela(opcao,"action",outro)
		notas_a_lancar=verifica_dados_tabela(opcao,"notas_a_lancar",outro)
			
		linha_pesos=split(ln_pesos_cols,"#!#")
		linha_pesos_variaveis=split(ln_pesos_vars,"#!#")
		nome_pesos_variaveis=split(nm_pesos_vars,"#!#")
		linha_nome_colunas=split(ln_nom_cols,"#!#")
		nome_variaveis=split(nm_vars,"#!#")
		variaveis_bd=split(nm_bd,"#!#")	
		calcula_variavel=split(vars_calc,"#!#")
		
		Set CON_N = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIR3		
		
		gera_pdf="sim"
	end if

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

		if bairro_unidade="" or isnull(bairro_unidade)then
		else
			bairro_unidade=" - "&bairro_unidade
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
		
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Periodo where TP_Modelo='"&tp_modelo&"'  AND NU_Periodo ="& periodo 
		RS5.Open SQL5, CON0
		
		no_periodo = RS5("NO_Periodo")		
		
		Set RS6 = Server.CreateObject("ADODB.Recordset")
		SQL6 = "SELECT * FROM TB_Materia where CO_Materia='"& co_materia &"'"
		RS6.Open SQL6, CON0
		
		no_materia= RS6("NO_Materia")		

		no_curso= no_etapa&"&nbsp;"&co_concordancia_curso&"&nbsp;"&no_curso
		texto_turma = "&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;<b>Turma</b> "&turma
		texto_disciplina = "&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;<b>Disciplina:</b> "&no_materia
		texto_periodo = "&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;<b>Per&iacute;odo:</b> "&no_periodo
		mensagem_cabecalho=ano_letivo&"&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;"&no_curso&texto_turma&texto_disciplina&texto_periodo
		mensagem_cabecalho=mensagem_cabecalho&"&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;Prof:&nbsp;"&no_prof

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
			Text = Right( Text, Len(Text) - CharsPrinted)
		Loop 

		y_texto=y_texto-altura_logo_gde+30
		SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height=30; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
		Text = "<center><i><b><font style=""font-size:18pt;"">PLANILHA DE NOTAS</font></b></i></center>"
		
		
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
		   .Cells(2).Width = area_utilizavel-51
		   .Cells(3).Width = 1	
		End With
		
		Table(1, 1).AddText "<b>Ano Letivo:</b>", "size=8;html=true", Font 
		Table(1, 2).AddText mensagem_cabecalho, "size=8;html=true", Font	
		'Table(1, 3).AddText "<div align=""right""><b>Legenda:</b> Md=M&eacute;dia - Res=Resultado&nbsp;&nbsp;&nbsp;</div>", "size=8;html=true", Font	
		Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 		

'================================================================================================================			
		colunas_de_notas=ubound(nome_variaveis)+1
		total_de_colunas=colunas_de_notas+2						
		altura_medias=30
		y_segunda_tabela=y_primeira_tabela-20	
		Set param_table2 = Pdf.CreateParam("width="&area_utilizavel&"; height="&altura_medias&"; rows=2; cols="&total_de_colunas&"; border=1; cellborder=0.1; cellspacing=0; x="&margem&"; y="&y_segunda_tabela&"; MaxHeight=420")

		Set Notas_Tit = Doc.CreateTable(param_table2)
		Notas_Tit.Font = Font				
		largura_colunas=(area_utilizavel-20-220)/colunas_de_notas		
		With Notas_Tit.Rows(1)
		   .Cells(1).Width = 20
		   .Cells(2).Width = 220			             
			for d=3 to total_de_colunas
			 .Cells(d).Width = largura_colunas					
			next
		End With
		
		alunos_encontrados = split(vetor_matriculas, "#$#" )							

		tabela_col=1
		for d=0 to ubound(linha_pesos)
			
			if linha_pesos(d)="PESO" then			
				linha_pesos_variaveis=split(ln_pesos_vars,"#!#")
				nome_pesos_variaveis=split(nm_pesos_vars,"#!#")						
		
				dados_alunos = split(alunos_encontrados(0), "#!#" )
				Set RSpeso = Server.CreateObject("ADODB.Recordset")
				SQL_peso = "Select * from "& tb_nota &" WHERE CO_Matricula = "& dados_alunos(0) & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				Set RSpeso = CON_N.Execute(SQL_peso)			 
				coluna=0	 
					
				if RSpeso.EOF then
					valor_peso=linha_pesos(d)
				else	
					valor_peso=RSpeso(""&linha_pesos_variaveis(d)&"")
				end if			

			Notas_Tit(1, tabela_col).AddText "<div align=""center""><b>"&valor_peso&"</b></div>", "size=8; indenty=2; alignment=center; html=true", Font
			else
			Notas_Tit(1, tabela_col).AddText "<div align=""center""><b>"&linha_pesos(d)&"</b></div>", "size=8; indenty=2; alignment=center; html=true", Font			
			end if
			tabela_col=tabela_col+1
		next			
		
		tabela_col=1
		for e=0 to ubound(linha_nome_colunas)
			Notas_Tit(2, tabela_col).AddText "<div align=""center""><b>"&linha_nome_colunas(e)&"</b></div>", "size=8; indenty=2; alignment=center; html=true", Font				
			tabela_col=tabela_col+1
		next			
		Set param_materias = PDF.CreateParam	
		param_materias.Set "size=8;expand=true" 			
												
		linha=2
		conta_notas = 1 
		
nu_chamada_ckq = 0

		for b=0 to ubound(alunos_encontrados)	
			param_materias.Add "indenty=2;alignment=right;html=true"
			param_materias.Add "indentx=5"	
			dados_alunos = split(alunos_encontrados(b), "#!#" )
			
'Verificando se algum aluno mudou de turma e inserindo uma linha em branco para o lugar do aluno

			teste_nu_cham=dados_alunos(1) - 1
		
			if (nu_chamada_ckq <>teste_nu_cham) then
				teste_nu_chamada = dados_alunos(1)-nu_chamada_ckq
				teste_nu_chamada=teste_nu_chamada-1

				for k=1 to teste_nu_chamada 				
					nu_chamada_ckq=nu_chamada_ckq+1
					coluna=1
				
					linha=linha+1
					Set Row = Notas_Tit.Rows.Add(15) 	
					for c=0 to ubound(linha_nome_colunas)
						Notas_Tit(linha, coluna).AddText "<div align=""center"">&nbsp;</DIV>", param_materias	
						coluna=coluna+1						
					next
				next		
'Inserindo o aluno seguinte aos que mudaram de turma	
				nu_chamada_ckq=nu_chamada_ckq+1		
				linha=linha+1				
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL_N = "Select * from "& tb_nota &" WHERE CO_Matricula = "& dados_alunos(0) & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				Set RS3 = CON_N.Execute(SQL_N)			  
				
				Set Row = Notas_Tit.Rows.Add(15) ' row height						
				Notas_Tit(linha, 1).AddText dados_alunos(1), param_materias	
				Notas_Tit(linha, 2).AddText dados_alunos(2), param_materias			
				coluna=3
				param_materias.Add "indentx=0"
				'for n=0 to ubound(co_materia_exibe)
				for c=0 to ubound(nome_variaveis)	
					if RS3.EOF then 
						valor="&nbsp;"
					else
						if variaveis_bd(c)="CALCULADO" then
							valor="&nbsp;"
							'Nesse caso o valor é calculado pela função calcular_nota chamada mais abaixo
						else
							valor=RS3(""&variaveis_bd(c)&"")
						end if											
					end if			
					if calcula_variavel(c)="CALC1" and valor="&nbsp;" then
						valor=calcular_nota(calcula_variavel(c),CAMINHOn,tb_nota,dados_alunos(0),mat_princ,co_materia,periodo)
						Notas_Tit(linha, coluna).AddText "<div align=""center"">"&valor&"</DIV>", param_materias
						coluna=coluna+1	
					else		
						Notas_Tit(linha, coluna).AddText "<div align=""center"">"&valor&"</DIV>", param_materias													
						coluna=coluna+1	
					end if				
				next
				
			else
				nu_chamada_ckq=nu_chamada_ckq+1					
'Se os números de chamada estiverem completos. Se não faltar aluno na turma.
			
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL_N = "Select * from "& tb_nota &" WHERE CO_Matricula = "& dados_alunos(0) & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				Set RS3 = CON_N.Execute(SQL_N)			 
			
				linha=linha+1
				Set Row = Notas_Tit.Rows.Add(15) ' row height						
				Notas_Tit(linha, 1).AddText dados_alunos(1), param_materias	
				Notas_Tit(linha, 2).AddText dados_alunos(2), param_materias			
				coluna=3
				param_materias.Add "indentx=0"
				
				for c=0 to ubound(nome_variaveis)	
					if RS3.EOF then 
						valor="&nbsp;"
					else
						if variaveis_bd(c)="CALCULADO" then
							valor="&nbsp;"
							'Nesse caso o valor é calculado pela função calcular_nota chamada mais abaixo
						else
							valor=RS3(""&variaveis_bd(c)&"")
						end if						
					end if			
					if calcula_variavel(c)="CALC1" and valor="&nbsp;" then
						valor=calcular_nota(calcula_variavel(c),CAMINHOn,tb_nota,dados_alunos(0),mat_princ,co_materia,periodo)
						Notas_Tit(linha, coluna).AddText "<div align=""center"">"&valor&"</DIV>", param_materias
						coluna=coluna+1
					else						
						Notas_Tit(linha, coluna).AddText "<div align=""center"">"&valor&"</DIV>", param_materias							
						coluna=coluna+1		
					end if					
				next			
			end if	
		next

		limite=0
		Do While True
		limite=limite+1
		   LastRow = Page.Canvas.DrawTable( Notas_Tit, param_table2 )

			if LastRow >= Notas_Tit.Rows.Count Then 
				Exit Do ' entire table displayed
			else
				 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
				
				Relatorio = "SWD015 - Sistema Web Diretor"
				Do While Len(Relatorio) > 0
					CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
				 
					If CharsPrinted = Len(Relatorio) Then Exit Do
					   SET Page = Page.NextPage
					Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
				Loop 
				
				Param_Relatorio.Add "alignment=right" 
				
				Paginacao = "1"
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
				param_table2.Add( "RowTo=2; RowFrom=1" ) ' Row 1 is header.
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
					Text = Right( Text, Len(Text) - CharsPrinted)
				Loop 
	
				y_texto=y_texto-altura_logo_gde+30
				SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height=30; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
				Text = "<center><i><b><font style=""font-size:18pt;"">PLANILHA DE NOTAS</font></b></i></center>"
				
				
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
			end if
			if limite>100 then
			response.Write("ERRO!")
			response.end()
			end if 
		Loop
		
		SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
		
		Relatorio = "SWD015 - Sistema Web Diretor"
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
		End IF	
		End IF			
		RSA.Close
		Set RSA = Nothing	
	
		RS2.Close
		Set RS2 = Nothing
			
		RS3.Close
		Set RS3 = Nothing
	
		RS3m.Close
		Set RS3m = Nothing
	
		RS4.Close
		Set RS4 = Nothing
		
		RS5.Close
		Set RS5 = Nothing				
		
		RS6.Close
		Set RS6 = Nothing		
				
		RStabela.Close
		Set RStabela = Nothing							
	End IF					
						
NEXT
	
arquivo="SWD015"
Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

