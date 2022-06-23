<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 600 'valor em segundos
'MAPA DE RESULTADO POR DISCIPLINA
%>
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/funcoes3.asp"-->
<!--#include file="../inc/funcoes5.asp"-->
<% 
response.Charset="ISO-8859-1"
opt = REQUEST.QueryString("obr")
nta= REQUEST.QueryString("n")

obr=split(opt,"_")
if not isnull(opt) then
co_materia= obr(0)
unidade = obr(1)
curso = obr(2)
co_etapa = obr(3)
turma = obr(4)
co_prof = obr(6)
end if
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
nvg=session("nvg")
session("nvg")=nvg

arquivo="SWD110"
'ano = DatePart("yyyy", now)
'mes = DatePart("m", now) 
'dia = DatePart("d", now) 
'hora = DatePart("h", now) 
'min = DatePart("n", now) 

if mes<10 then
mes="0"&mes
end if

'data = dia &"/"& mes &"/"& ano
'
'if min<10 then
'min="0"&min
'end if
'
'horario = hora & ":"& min

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
		
		Set CON5 = Server.CreateObject("ADODB.Connection") 
		ABRIR5 = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON5.Open ABRIR5		
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0	
				
		Set CON_A = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_A.Open ABRIR
		
call GeraNomes(co_materia,unidade,curso,co_etapa,CON0)

no_materia= session("no_materia")
no_unidades= session("no_unidades")
no_grau= session("no_grau")
no_serie= session("no_serie")


if nta="a" then
if avaliacao ="TES" then
campo_check ="VA_Teste" 
elseif avaliacao ="PRO" then
campo_check ="VA_Prova"
elseif avaliacao ="N3" then
campo_check ="VA_3Nota"
elseif avaliacao ="BON" then
campo_check ="VA_Bonus"
elseif avaliacao ="REC" then
campo_check ="VA_Rec"
end if

elseif nta="b" then
if avaliacao ="A1" then
campo_check ="VA_Nota_A1"
elseif avaliacao ="A2" then
campo_check ="VA_Nota_A2"
elseif avaliacao ="B1" then
campo_check ="VA_Nota_B1"
elseif avaliacao ="B2" then
campo_check ="VA_Nota_B2"
elseif avaliacao ="AV1" then
campo_check ="VA_Nota1"
elseif avaliacao ="AV2" then
campo_check ="VA_Nota2"
elseif avaliacao ="N3" then
campo_check ="VA_Nota3"
elseif avaliacao ="N4" then
campo_check ="VA_Nota4"
elseif avaliacao ="BON" then
campo_check ="VA_Bonus"
elseif avaliacao ="REC" then
campo_check ="VA_Rec"
end if

elseif nta="c" then
if avaliacao ="N1" then
campo_check ="VA_Nota1"
elseif avaliacao ="N2" then
campo_check ="VA_Nota2"
elseif avaliacao ="N3" then
campo_check ="VA_Nota3"
elseif avaliacao ="BON" then
campo_check ="VA_Bonus"
elseif avaliacao ="REC" then
campo_check ="VA_Rec"
end if
end if

obr=unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&periodo&"_"&ano_letivo

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

		Set RSTB = Server.CreateObject("ADODB.Recordset")
		CONEXAOTB = "Select * from TB_Da_Aula WHERE NU_Unidade="& unidade &" AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		Set RSTB = CON2.Execute(CONEXAOTB)
		
nota= RSTB("TP_Nota")

		
if nota = "TB_NOTA_A" Then		
		CAMINHOn = CAMINHO_na
elseif nota = "TB_NOTA_B" Then
		CAMINHOn = CAMINHO_nb
elseif nota = "TB_NOTA_C" Then
		CAMINHOn = CAMINHO_nc
end if

		Set CON3 = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")
un_endereco = RS0("NO_Logradouro")
un_complemento = RS0("TX_Complemento_Logradouro")
un_numero = RS0("NU_Logradouro")
un_bairro = RS0("CO_Bairro")
un_cidade = RS0("CO_Municipio")
un_uf = RS0("SG_UF")
un_tel = RS0("NUS_Telefones")
un_email = RS0("TX_EMail")
un_cep = RS0("CO_CEP")
un_ato = RS0("TX_Ato_Autorizativo")
un_cnpj = RS0("CO_CGC")


if un_ato="" or isnull(un_ato) then
separador1=0
else
separador1=1
end if

if un_complemento="" or isnull(un_complemento) then
separador2=0
else
separador2=1
end if
if un_email="" or isnull(un_email) then
separador3=0
else
separador3=1
end if

cep3=left(un_cep,5)
cep4 = right(un_cep,3)

un_cep=cep3&"-"&cep4


		Set RS11 = Server.CreateObject("ADODB.Recordset")
		SQL11 = "SELECT * FROM TB_Municipios WHERE SG_UF ='"& un_uf &"' AND CO_Municipio = "&un_cidade
		RS11.Open SQL11, CON0

cidade= RS11("NO_Municipio")

		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Bairros WHERE CO_Bairro ="& un_bairro &"AND SG_UF ='"& un_uf&"' AND CO_Municipio = "&un_cidade
		RS4.Open SQL4, CON0
if RS4.EOF then
bairro = ""
else
bairro= RS4("NO_Bairro")
end if

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Curso")



		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' and CO_Curso ='"& curso &"'"  
		RS3.Open SQL3, CON0
		
if RS3.EOF THEN
no_etapa="sem etapa"
else
no_etapa=RS3("NO_Etapa")
end if


P1=0
P2=0
P3=0
rec_ckeck="no"
res1=""
res2=""

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

select case mes
 case 1 
 mes = "janeiro"
 case 2 
 mes = "fevereiro"
 case 3 
 mes = "março"
 case 4
 mes = "abril"
 case 5
 mes = "maio"
 case 6 
 mes = "junho"
 case 7
 mes = "julho"
 case 8 
 mes = "agosto"
 case 9 
 mes = "setembro"
 case 10 
 mes = "outubro"
 case 11 
 mes = "novembro"
 case 12 
 mes = "dezembro"
end select

if min<10 then
min = "0"&min
end if

data = dia &" de "& mes &" de "& ano
horario = hora & ":"& min

		ntvml= 4.5
		ntazl= 6.99
		cor_nota_vml="#FF0000"	
		cor_nota_azl="#0000FF"	
		cor_nota_prt="#000000"	
		
		
	rec_lancado="sim"
		

				
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)
		
			check=2
		
		tb=nota
		
		Set RSA = Server.CreateObject("ADODB.Recordset")
		if ano_letivo=ano_vigente then	
		CONEXAOA = "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada, TB_Alunos.NO_Aluno from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Situacao='C' AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.NU_Unidade = "& unidade&" AND TB_Matriculas.CO_Curso = '"& curso &"' AND TB_Matriculas.CO_Etapa = '"& co_etapa &"' AND TB_Matriculas.CO_Turma = '"& turma &"' order by TB_Matriculas.NU_Chamada"
		else
			CONEXAOA = "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada, TB_Alunos.NO_Aluno from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.NU_Unidade = "& unidade&" AND TB_Matriculas.CO_Curso = '"& curso &"' AND TB_Matriculas.CO_Etapa = '"& co_etapa &"' AND TB_Matriculas.CO_Turma = '"& turma &"' order by TB_Matriculas.NU_Chamada"
		end if
	Set RSA = CON1.Execute(CONEXAOA)

	vetor_matriculas="" 
	gera_pdf="nao"
	nu_chamada_check = 1
	While Not RSA.EOF
		gera_pdf="sim"
		nu_matricula = RSA("CO_Matricula")
		no_aluno= RSA("NO_Aluno")			
		nu_chamada = RSA("NU_Chamada")
		
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
	if curso<>0 then	
		if gera_pdf="sim" then	
		
			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL2 = "SELECT * FROM TB_Unidade WHERE NU_Unidade="& unidade
			RS2.Open SQL2, CON0
							
			no_unidade = RS2("NO_Sede")		
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

			no_curso= no_etapa&"&nbsp;"&co_concordancia_curso&"&nbsp;"&no_curso
			texto_turma = "&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;<b>Turma</b> "&turma

			mensagem_cabecalho=ano_letivo&"&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;"&no_curso&texto_turma
	
						

			Set RS5 = Server.CreateObject("ADODB.Recordset")
			SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND IN_MAE=TRUE order by NU_Ordem_Boletim "
			RS5.Open SQL5, CON0
			co_materia_check=1
			IF RS5.EOF Then
				vetor_materia_exibe="nulo"
			else
				while not RS5.EOF
					co_mat_fil= RS5("CO_Materia")
					'carga_materia= RS5("NU_Aulas")				
					if co_materia_check=1 then
						vetor_materia=co_mat_fil
					else
						vetor_materia=vetor_materia&"#!#"&co_mat_fil
					end if
					co_materia_check=co_materia_check+1			
							
				RS5.MOVENEXT
				wend						
			end if	
		end if
	end if
	

	SET Page = Doc.Pages.Add(842, 595)
					
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
			Text = "<p><i><b>"&UCASE(no_unidade)&"</b></i></p><br><font style=""font-size:10pt;"">"&endereco_unidade&"</FONT><br><center><i><b><font style=""font-size:18pt;"">MAP&Atilde;O" 
			Text = Text&" DE RESULTADOS POR DISCIPLINA</font></b></i></center>"
			
			
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
			Table(1, 3).AddText "<div align=""right""><b>Disciplina:&nbsp;</b>"&Server.HTMLEncode(no_materia)&"&nbsp;&nbsp;&nbsp;</div>", "size=8;html=true", Font	
			Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 		
	'================================================================================================================	

			colunas_de_notas=12
			total_de_colunas=colunas_de_notas+2						
			altura_medias=45
			y_segunda_tabela=y_primeira_tabela-20	
			linha=2
			Set param_table2 = Pdf.CreateParam("width="&area_utilizavel&"; height="&altura_medias&"; rows="&			linha&"; cols="&total_de_colunas&"; border=1; cellborder=0.1; cellspacing=0; x="&margem&"; y="&y_segunda_tabela&"; MaxHeight=420")

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
			With Notas_Tit.Rows(2)
			   .Cells(1).Height = 25
			End With			
			Notas_Tit(1, 1).ColSpan = 10
			Notas_Tit(1, 11).ColSpan = 4
			Notas_Tit(1, 1).AddText "<div align=""center""><b>Aproveitamento</b></div>", "size=10;indenty=2; html=true", Font 
			Notas_Tit(1, 11).AddText "<div align=""center""><b>Frequ&ecirc;ncia</b></div>", "size=10;alignment=center; indenty=2;html=true", Font 					
			Notas_Tit(2, 1).AddText "<div align=""center""><b>N&ordm;</b></div>", "size=10;indenty=5; html=true", Font 
			Notas_Tit(2, 2).AddText "<div align=""center""><b>Nome</b></div>", "size=10;alignment=center; indenty=5;html=true", Font 
			Notas_Tit(2, 3).AddText "<div align=""center""><b>PA1</b></div>", "size=10;indenty=5; html=true", Font 
			Notas_Tit(2, 4).AddText "<div align=""center""><b>PA2</b></div>", "size=10;alignment=center; indenty=5;html=true", Font 
			Notas_Tit(2, 5).AddText "<div align=""center""><b>PA3</b></div>", "size=10;indenty=5; html=true", Font 
			Notas_Tit(2, 6).AddText "<div align=""center""><b>TOTAL</b></div>", "size=10;alignment=center; indenty=5;html=true", Font 
			Notas_Tit(2, 7).AddText "<div align=""center""><b>4&ordf; Aval.<br>p.2</b></div>", "size=8;indenty=1; html=true", Font 
			Notas_Tit(2, 8).AddText "<div align=""center""><b>TOTAL</b></div>", "size=10;alignment=center; indenty=5;html=true", Font 
			Notas_Tit(2, 9).AddText "<div align=""center""><b>FALTA</b></div>", "size=10;indenty=5; html=true", Font 
			Notas_Tit(2, 10).AddText "<div align=""center""><b>M&eacute;dia<br>Final</b></div>", "size=8;alignment=center; indenty=1;html=true", Font 
			Notas_Tit(2, 11).AddText "<div align=""center""><b>PA1</b></div>", "size=10;indenty=5; html=true", Font 
			Notas_Tit(2, 12).AddText "<div align=""center""><b>PA2</b></div>", "size=10;alignment=center; indenty=5;html=true", Font 	
			Notas_Tit(2, 13).AddText "<div align=""center""><b>PA3</b></div>", "size=10;alignment=center; indenty=5;html=true", Font 
			Notas_Tit(2, 14).AddText "<div align=""center""><b>TOTAL</b></div>", "size=10;indenty=5; html=true", Font 
																	
			tabela_col=2
				'for d=0 to ubound(co_materia_exibe)
'					tabela_col=tabela_col+1
'					Notas_Tit(1, tabela_col).ColSpan = 2
'					Notas_Tit(1, tabela_col).AddText "<div align=""center""><b>"&co_materia_exibe(d)&"</b></div>", "size=8; indenty=2; alignment=center; html=true", Font
'					Notas_Tit(2, tabela_col).AddText "<div align=""center""><b>Md</b></div>", "size=8;alignment=center; indenty=2;html=true", Font
'					proxima_coluna=tabela_col+1
'					Notas_Tit(2, proxima_coluna).AddText "<div align=""center""><b>Res</b></div>", "size=8;alignment=center; indenty=2;html=true", Font
'					'Para não dar o colspan na coluna seguinte.
'					tabela_col=tabela_col+1
'				next			
			Set param_materias = PDF.CreateParam	
			param_materias.Set "size=8;expand=true" 			
												
			alunos_encontrados = split(vetor_matriculas, "#$#" )
			
			for a=0 to ubound(alunos_encontrados)	
				param_materias.Add "indenty=2;alignment=right;html=true"
				param_materias.Add "indentx=5"	
				dados_alunos = split(alunos_encontrados(a), "#!#" )		
				Set Row = Notas_Tit.Rows.Add(15) ' row height		
				linha=linha+1		
				nu_matricula = dados_alunos(0)		
				Notas_Tit(linha, 1).AddText dados_alunos(1), param_materias	
				Notas_Tit(linha, 2).AddText dados_alunos(2), param_materias			
				coluna=2
				param_materias.Add "indentx=0"
				
				dividendo1=0
					divisor1=0
					dividendo2=0
					divisor2=0
					dividendo3=0
					divisor3=0
					dividendo4=0
					divisor4=0
					dividendo_ma=0
					divisor_ma=0
					dividendo5=0
					divisor5=0
					dividendo_mf=0
					divisor_mf=0
					dividendo6=0
					divisor6=0
					dividendo_rec=0
					divisor_rec=0
					res1="&nbsp;"
					res2="&nbsp;"
					res3="&nbsp;"
					m2="&nbsp;"
					m3="&nbsp;"										
			
			
'				Set RS1a = Server.CreateObject("ADODB.Recordset")
'				SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia='"&co_materia&"'"
'				RS1a.Open SQL1a, CON0
'					
'				no_materia=RS1a("NO_Materia")
			
				if check mod 2 =0 then
				cor = "tb_fundo_linha_par" 
				cor2 = "tb_fundo_linha_impar" 				
				else 
				cor ="tb_fundo_linha_impar"
				cor2 = "tb_fundo_linha_par" 
				end if
			
				va_m31="&nbsp;"
				va_m32="&nbsp;"
				va_m33="&nbsp;"
				va_m34="&nbsp;"
				
				med1="&nbsp;"
				med2="&nbsp;"
				med3="&nbsp;"
				med4="&nbsp;"	
				
				m2="&nbsp;"
				pendente="&nbsp;"
				mf3="&nbsp;"
				f1="&nbsp;"
				f2="&nbsp;"
				f3="&nbsp;"
				soma_f="&nbsp;"
				
					
				Set CON_N = Server.CreateObject("ADODB.Connection") 
				ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
				CON_N.Open ABRIRn
			
				for periodofil=1 to 4
										
						
					Set RSnFIL = Server.CreateObject("ADODB.Recordset")
					Set RS3 = Server.CreateObject("ADODB.Recordset")
					SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodofil
					Set RS3 = CON_N.Execute(SQL_N)
				

				
					if RS3.EOF then
						if periodofil=1 then
						va_m31="&nbsp;"
						elseif periodofil=2 then
						va_m32="&nbsp;"
						elseif periodofil=3 then
						va_m33="&nbsp;"
						elseif periodofil=4 then
						va_m34="&nbsp;"
						end if	
					else
						if periodofil=1 then
						va_m31=RS3("VA_Media3")
						elseif periodofil=2 then
						va_m32=RS3("VA_Media3")
						elseif periodofil=3 then
						va_m33=RS3("VA_Media3")
						elseif periodofil=4 then
						va_m34=RS3("VA_Media3")
						end if
					end if
				NEXT		
		
					if isnull(f1) or f1="&nbsp;"  or f1="" then
					soma_f1=0
					else
					soma_f1=f1
					end if

					if isnull(f2) or f2="&nbsp;"  or f2="" then
					soma_f2=0
					else
					soma_f2=f2
					end if
					
					if isnull(f3) or f3="&nbsp;"  or f3="" then
					soma_f3=0
					else
					soma_f3=f3
					end if					
					
					soma_f=soma_f1+soma_f2+soma_f3

					if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
					dividendo1=0
					divisor1=0
					else
					dividendo1=va_m31
					divisor1=1
					end if	
					
					if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
					dividendo2=0
					divisor2=0
					else
					dividendo2=va_m32
					divisor2=1
					end if
					
					if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
					dividendo3=0
					divisor3=0
					else
					dividendo3=va_m33
					divisor3=1
					end if
					
					if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
					nota_aux_m2_1="&nbsp;"
					dividendo4=0
					divisor4=0
					else
					nota_aux_m2_1=va_m34
					dividendo4=va_m34
					divisor4=1
					end if
								
					dividendo_ma=dividendo1+dividendo2+dividendo3
					divisor_ma=divisor1+divisor2+divisor3
					divisor_m3=divisor1+divisor2+divisor3+(divisor4*2)		
					'response.Write(dividendo_ma&"<<")
					
					if divisor_ma<3 then
					ma="&nbsp;"
					else
					ma=dividendo_ma
					end if
					
					if ma="&nbsp;" then
					else
					ma=ma*10
							decimo = ma - Int(ma)
							If decimo >= 0.5 Then
								nota_arredondada = Int(ma) + 1
								ma=nota_arredondada
							else
								nota_arredondada = Int(ma)
								ma=nota_arredondada											
							End If
					ma=ma/10						
						ma = formatNumber(ma,1)	
					end if

					
					if ma="&nbsp;" then
					dividendo_m2=0
					divisor_m2=0
					else
					dividendo_m2=ma+(dividendo4*2)
					divisor_m2=1
					end if
					
					if divisor_m2=0 then
						m2="&nbsp;"
						if divisor1=1 and divisor2=1 then
							pendente = 21-dividendo1-dividendo2
						else
							pendente="&nbsp;"
						end if	
					else
						m2=dividendo_m2							
						if m2>=21 then
							pendente=0
						else							
							pendente=(25-m2)*10
							decimo = pendente - Int(pendente)
							If decimo >= 0.5 Then
								nota_arredondada = Int(pendente) + 1
								pendente=nota_arredondada
							else
								nota_arredondada = Int(pendente)
								pendente=nota_arredondada					
							End If	
							pendente = (pendente/10)/2
							if pendente<0 then
								pendente=0
							end if
						end if
					end if
					
					if m2="&nbsp;" then
					else
					m3=m2/divisor_m3
					m3=m3*10					
						decimo = m3 - Int(m3)
							If decimo >= 0.5 Then
								nota_arredondada = Int(m3) + 1
								m3=nota_arredondada
							else
								nota_arredondada = Int(m3)
								m3=nota_arredondada					
							End If
					m3=m3/10						
						m3 = formatNumber(m3,1)					
					end if
				if divisor1=1 then
					med1=formatnumber(va_m31,1)
				end if	
				if divisor2=1 then
					med2=formatnumber(va_m32,1)
				end if	
				if divisor3=1 then
					med3=formatnumber(va_m33,1)
				end if				
				if divisor4=1 then
					med4=formatnumber(va_m34,1)
				end if	
				
				if m2<>"&nbsp;" then
					mf3=formatnumber(m3,1)
				end if					
										
									
					
				Notas_Tit(linha, 3).AddText "<div align=""center"">"&med1&"</DIV>", param_materias	
				Notas_Tit(linha, 4).AddText "<div align=""center"">"&med2&"</DIV>", param_materias		
				Notas_Tit(linha, 5).AddText "<div align=""center"">"&med3&"</DIV>", param_materias		
				Notas_Tit(linha, 6).AddText "<div align=""center"">"&ma&"</DIV>", param_materias																		
				Notas_Tit(linha, 7).AddText "<div align=""center"">"&med4&"</DIV>", param_materias																		
				Notas_Tit(linha, 8).AddText "<div align=""center"">"&m2&"</DIV>", param_materias																		
				Notas_Tit(linha, 9).AddText "<div align=""center"">"&pendente&"</DIV>", param_materias																		
				Notas_Tit(linha, 10).AddText "<div align=""center"">"&mf3&"</DIV>", param_materias																		
				Notas_Tit(linha, 11).AddText "<div align=""center"">"&f1&"</DIV>", param_materias																		
				Notas_Tit(linha, 12).AddText "<div align=""center"">"&f2&"</DIV>", param_materias																		
				Notas_Tit(linha, 13).AddText "<div align=""center"">"&f3&"</DIV>", param_materias																		
				Notas_Tit(linha, 14).AddText "<div align=""center"">"&soma_f&"</DIV>", param_materias																		

			next 
			
			limite=0
			Do While True
			limite=limite+1
			   LastRow = Page.Canvas.DrawTable( Notas_Tit, param_table2 )
	
				if LastRow >= Notas_Tit.Rows.Count Then 
			    	Exit Do ' entire table displayed
				else
					 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
					
					Relatorio = "SWD104 - Sistema Web Diretor"
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
					Loop 				   ' Display remaining part of table on the next page
					Set Page = Page.NextPage	
					param_table2.Add( "RowTo=2; RowFrom=1" ) ' Row 1 is header.
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
					Text = "<p><i><b>"&UCASE(no_unidade)&"</b></i></p><br><font style=""font-size:10pt;"">"&endereco_unidade&"</FONT><br><center><i><b><font style=""font-size:18pt;"">MAP&Atilde;O" 
			Text = Text&" DE RESULTADOS POR DISCIPLINA</font></b></i></center>"
					
					
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
			Table(1, 3).AddText "<div align=""right""><b>Disciplina:&nbsp;</b>"&Server.HTMLEncode(no_materia)&"&nbsp;&nbsp;&nbsp;</div>", "size=8;html=true", Font	
					Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 			
'================================================================================================================				 
			 	end if
				if limite>100 then
				response.Write("ERRO!")
				response.end()
				end if 
			Loop
			
			SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
			
			Relatorio = "SWD104 - Sistema Web Diretor"
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

Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

