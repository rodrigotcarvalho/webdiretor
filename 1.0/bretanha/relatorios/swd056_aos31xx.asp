<%'On Error Resume Next%>
<%
'Server.ScriptTimeout = 60
Server.ScriptTimeout = 600 'valor em segundos
'ATA DE RESULTADOS
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/funcoes7.asp"-->
<!--#include file="../../global/funcoes_diversas.asp"-->
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
'ano = DatePart("yyyy", now)
'mes = DatePart("m", now) 
'dia = DatePart("d", now) 
'hora = DatePart("h", now) 
'min = DatePart("n", now) 

if ori="ebe" then
origem="../ws/doc/ofc/ebe/"
end if

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
		
co_usr = session("co_user")

	SQL_NOM = "select * from TB_Usuario where CO_Usuario = " & co_usr
	set USR_NOM = CON.Execute (SQL_NOM)

nome_usuario=USR_NOM("NO_Usuario")		
	
if opt="01" then
	cod_cons=request.QueryString("cod_cons")
	
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
	
	if turma="999990" or turma="" or isnull(turma) then
		if co_etapa="999990" or co_etapa="" or isnull(co_etapa) then
			if curso="999990" or curso="" or isnull(curso) then		
				if unidade="999990" or unidade="" or isnull(unidade) then
					response.redirect(origem&"index.asp?nvg="&nvg&"&opt=err2")
				else	
					Set RS0 = Server.CreateObject("ADODB.Recordset")
					SQL0 = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&" ORDER BY CO_Curso,CO_Etapa"
					RS0.Open SQL0, CON0
					check_motriz=1
					WHILE NOT RS0.EOF
						curso=RS0("CO_Curso")
						co_etapa=RS0("CO_Etapa")
						
						Set RS0t = Server.CreateObject("ADODB.Recordset")
						SQL0t = "SELECT * FROM TB_Turma where NU_Unidade="&unidade&" AND CO_Curso ='"&curso&"' AND CO_Etapa ='"&co_etapa&"' ORDER BY CO_Turma"
						RS0t.Open SQL0t, CON0							
						WHILE NOT RS0t.EOF								
							turma=RS0t("CO_Turma")	

							if check_motriz=1 then
								vetor_motriz=unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
							else
								vetor_motriz=vetor_motriz&"#$#"&unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
							end if
							check_motriz=check_motriz+1 
						RS0t.MOVENEXT
						WEND	
					RS0.MOVENEXT
					WEND					
					RS0.Close
					Set RS0 = Nothing	
				end if		
			else	
				Set RS0 = Server.CreateObject("ADODB.Recordset")
				SQL0 = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&" AND CO_Curso ='"&curso&"' ORDER BY CO_Etapa"
				RS0.Open SQL0, CON0
				check_motriz=1
				WHILE NOT RS0.EOF
					co_etapa=RS0("CO_Etapa")					
					Set RS0t = Server.CreateObject("ADODB.Recordset")
					SQL0t = "SELECT * FROM TB_Turma where NU_Unidade="&unidade&" AND CO_Curso ='"&curso&"' AND CO_Etapa ='"&co_etapa&"' ORDER BY CO_Turma"
					RS0t.Open SQL0t, CON0							
					WHILE NOT RS0t.EOF								
						turma=RS0t("CO_Turma")	

						if check_motriz=1 then
							vetor_motriz=unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
						else
							vetor_motriz=vetor_motriz&"#$#"&unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
						end if
						check_motriz=check_motriz+1 
					RS0t.MOVENEXT
					WEND	
				RS0.MOVENEXT
				WEND
				
				RS0.Close
				Set RS0 = Nothing					
			end if						
		else				
			Set RS0t = Server.CreateObject("ADODB.Recordset")
			SQL0t = "SELECT * FROM TB_Turma where NU_Unidade="&unidade&" AND CO_Curso ='"&curso&"' AND CO_Etapa ='"&co_etapa&"' ORDER BY CO_Turma"
			RS0t.Open SQL0t, CON0					
					
			check_motriz=1			
			
			WHILE NOT RS0t.EOF								
				turma=RS0t("CO_Turma")	

				if check_motriz=1 then
					vetor_motriz=unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
				else
					vetor_motriz=vetor_motriz&"#$#"&unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
				end if
				check_motriz=check_motriz+1 
			RS0t.MOVENEXT
			WEND	
		end if	
		RS0t.Close
		Set RS0t = Nothing	
	ELSE
		vetor_motriz=unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma				
	end if					
'response.Write(vetor_motriz)
'response.end()
end if

conjunto_dados=split(vetor_motriz,"#$#")

for i=0 to ubound(conjunto_dados)	
	dados_select=split(conjunto_dados(i),"#!#")
	unidade=dados_select(0)
	curso=dados_select(1)
	co_etapa=dados_select(2)
	turma=dados_select(3)		
	
	nu_chamada_check = 1	
	
	Set RSA = Server.CreateObject("ADODB.Recordset")
	'CONEXAOA = "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada, TB_Matriculas.CO_Situacao, TB_Alunos.NO_Aluno from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Situacao='C' AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.NU_Unidade = "& unidade&" AND TB_Matriculas.CO_Curso = '"& curso &"' AND TB_Matriculas.CO_Etapa = '"& co_etapa &"' AND TB_Matriculas.CO_Turma = '"& turma &"' order by TB_Alunos.NO_Aluno"
	CONEXAOA = "Select TB_Matriculas.CO_Matricula, TB_Alunos.RA_Aluno, TB_Matriculas.NU_Chamada, TB_Matriculas.CO_Situacao, TB_Alunos.NO_Aluno from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.NU_Unidade = "& unidade&" AND TB_Matriculas.CO_Curso = '"& curso &"' AND TB_Matriculas.CO_Etapa = '"& co_etapa &"' AND TB_Matriculas.CO_Turma = '"& turma &"' order by TB_Alunos.NO_Aluno"
	
	Set RSA = CON1.Execute(CONEXAOA)

	vetor_matriculas="" 
	While Not RSA.EOF
		nu_matricula = RSA("CO_Matricula")
		registro = 	RSA("RA_Aluno")
		nome_aluno= RSA("NO_Aluno")			
		nu_chamada = RSA("NU_Chamada")
		situacao = RSA("CO_Situacao")	
		
		if registro="" or isnull(registro) then
			registro = "000000000000"
		end if	
		
		nome_aluno=replace_latin_char(nome_aluno,"html")	
		nu_chamada_check=nu_chamada_check*1
		nu_chamada=nu_chamada*1
		if nu_chamada_check = 1 and nu_chamada=nu_chamada_check then
			vetor_matriculas=nu_matricula&"#!#"&nu_chamada&"#!#"&nome_aluno&"#!#"&situacao&"#!#"&registro
		elseif nu_chamada_check = 1 then
			while nu_chamada_check < nu_chamada
				nu_chamada_check=nu_chamada_check+1
			wend 
			vetor_matriculas=nu_matricula&"#!#"&nu_chamada&"#!#"&nome_aluno&"#!#"&situacao&"#!#"&registro
		else
			vetor_matriculas=vetor_matriculas&"#$#"&nu_matricula&"#!#"&nu_chamada&"#!#"&nome_aluno&"#!#"&situacao&"#!#"&registro
		end if
	nu_chamada_check=nu_chamada_check+1		
	RSA.MoveNext
	Wend 

	if curso=0 then

	else
		Set RStabela = Server.CreateObject("ADODB.Recordset")
		SQLtabela = "SELECT * FROM TB_Da_Aula where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'" 
		RStabela.Open SQLtabela, CON2

		if 	RStabela.EOF then
			gera_pdf="nao"
		else				
			tb_nota=RStabela("TP_Nota")		
			if tb_nota ="TB_NOTA_A" then
				caminho_nota = CAMINHO_na
				gera_pdf="sim"
			elseif tb_nota="TB_NOTA_B" then
				caminho_nota = CAMINHO_nb
				gera_pdf="sim"
			elseif tb_nota ="TB_NOTA_C" then
				caminho_nota = CAMINHO_nc
				gera_pdf="sim"
			elseif tb_nota ="TB_NOTA_D" then
				caminho_nota = CAMINHO_nd
				gera_pdf="sim"
			elseif tb_nota ="TB_NOTA_E" then
				caminho_nota = CAMINHO_ne
				gera_pdf="sim"				
			else
				gera_pdf="nao"
			end if	
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
			
			Set RS2b = Server.CreateObject("ADODB.Recordset")
			SQL2b = "SELECT * FROM TB_Bairros WHERE SG_UF='"& uf_unidade&"' AND CO_Municipio="&municipio_unidade&" AND CO_Bairro="&bairro_unidade
			RS2b.Open SQL2b, CON0
			
			if RS2b.EOF then
				nome_bairro_unidade=""			
			else
				nome_bairro_unidade=RS2b("NO_Bairro")
			end if					


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
				bairro_unidade=" - "&nome_bairro_unidade
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
			area_utilizavel=Page.Width - (margem*2)
			largura_logo_gde=formatnumber(Logo.Width*0.6,0)
			altura_logo_gde=formatnumber(Logo.Height*0.6,0)
	
		   Param_Logo_Gde("x") = margem
		   Param_Logo_Gde("y") = Page.Height - altura_logo_gde - margem
		   Param_Logo_Gde("ScaleX") = 0.6
Param_Logo_Gde("ScaleY") = 0.6
		   Page.Canvas.DrawImage Logo, Param_Logo_Gde
	
			x_texto=largura_logo_gde+20
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
			Text = "<center><i><b><font style=""font-size:18pt;"">Ata de Resultados Finais</font></b></i></center>"
			
			
			Do While Len(Text) > 0
				CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
			 
				If CharsPrinted = Len(Text) Then Exit Do
					SET Page = Page.NextPage
				Text = Right( Text, Len(Text) - CharsPrinted)
			Loop 
	
			
			Page.Canvas.SetParams "LineWidth=1" 
			Page.Canvas.SetParams "LineCap=0" 
			inicio_primeiro_separador=x_texto
			altura_primeiro_separador= Page.Height - margem - 17
			With Page.Canvas
			   .MoveTo inicio_primeiro_separador, altura_primeiro_separador
			   .LineTo area_utilizavel+margem, altura_primeiro_separador
			   .Stroke
			End With 	
	
	
			Page.Canvas.SetParams "LineWidth=2" 
			Page.Canvas.SetParams "LineCap=0" 
			altura_segundo_separador= Page.Height - altura_logo_gde -margem - 10
			With Page.Canvas
			   .MoveTo margem, altura_segundo_separador
			   .LineTo area_utilizavel+margem, altura_segundo_separador
			   .Stroke
			End With 	

			Set param_table1 = Pdf.CreateParam("width="&area_utilizavel&"; height=40; rows=2; cols=3; border=0; cellborder=0; cellspacing=0;")
			Set Table = Doc.CreateTable(param_table1)
			Table.Font = Font
			y_primeira_tabela=altura_segundo_separador-10
			x_primeira_tabela=margem+5
			With Table.Rows(1)
			   .Cells(1).Width = 50			   		   		   
			   .Cells(2).Width = area_utilizavel-250
			   .Cells(3).Width = 200	
			End With
			Table(1, 1).ColSpan = 3			
			Table(1, 1).AddText "Aos 31 de dezembro de "&ano_letivo&" terminou-se o processo de apura&ccedil;&atilde;o de notas finais e nota global dos alunos deste estabelecimento com os seguintes resultados:", "size=8;html=true", Font	

			Table(2, 1).AddText "<b>Ano Letivo:</b>", "size=8;html=true", Font 
			Table(2, 2).AddText mensagem_cabecalho, "size=8;html=true", Font	
			Table(2, 3).AddText "<div align=""right""><b>Legenda:</b> FRQ = Frequencia / RF = Resultado Final&nbsp;&nbsp;&nbsp;</div>", "size=8;html=true", Font	
			Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 		

	'================================================================================================================			

			Set RS5 = Server.CreateObject("ADODB.Recordset")
			SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND IN_MAE=TRUE order by NU_Ordem_Boletim "
			RS5.Open SQL5, CON0
			
			co_materia_check=1
			dias_de_aula_no_ano=0
			
			IF RS5.EOF Then
				vetor_materia_exibe="nulo"
			else
				while not RS5.EOF
					co_mat_fil= RS5("CO_Materia")
					co_materia= RS5("CO_Disciplina_Rel")	
					carga_materia= RS5("NU_Aulas")				
					if co_materia_check=1 then
						vetor_materia=co_mat_fil
						vetor_materia_rel=co_materia						
					else
						vetor_materia=vetor_materia&"#!#"&co_mat_fil
						vetor_materia_rel=vetor_materia_rel&"#!#"&co_materia							
					end if
					co_materia_check=co_materia_check+1			
					dias_de_aula_no_ano=dias_de_aula_no_ano+carga_materia
				RS5.MOVENEXT
				wend						
			end if

'response.Write(vetor_materia)
'response.end()
			
			co_materia_exibe=Split(vetor_materia_rel,"#!#")		
			co_materia_verifica=Split(vetor_materia,"#!#")							
			colunas_de_notas=18
			total_de_colunas=20					
			altura_medias=25
			y_segunda_tabela=y_primeira_tabela-40	
			Set param_table2 = Pdf.CreateParam("width="&area_utilizavel&"; height="&altura_medias&"; rows=1; cols="&total_de_colunas&"; border=1; cellborder=0.1; cellspacing=0; x="&margem&"; y="&y_segunda_tabela&"; MaxHeight=320")

			Set Notas_Tit = Doc.CreateTable(param_table2)
			Notas_Tit.Font = Font				
			largura_colunas=(area_utilizavel-70-185)/colunas_de_notas		
			
			With Notas_Tit.Rows(1)
			   .Cells(1).Width = 70
			   .Cells(2).Width = 185			             
				for d=3 to total_de_colunas
				 .Cells(d).Width = largura_colunas					
				next
			End With
			'Notas_Tit(1, 1).RowSpan = 2
			'Notas_Tit(1, 2).RowSpan = 2		
			'Notas_Tit(1, 1).AddText "<div align=""center""><b>Matr&iacute;cula</b></div>", "size=8;indenty=6; html=true", Font 
			Notas_Tit(1, 1).AddText "<div align=""center""><b>RA</b></div>", "size=8;indenty=6; html=true", Font 			
			Notas_Tit(1, 2).AddText "<div align=""center""><b>Nome</b></div>", "size=8;alignment=center; indenty=6;html=true", Font 
			Notas_Tit(1, 19).AddText "<div align=""center""><b>FRQ.<br>(%)</b></div>", "size=8;alignment=center; indenty=0;html=true", Font
			Notas_Tit(1, 20).AddText "<div align=""center""><b>RF.</b></div>", "size=8;alignment=center; indenty=6;html=true", Font 

			tabela_col=2
			for d=0 to ubound(co_materia_exibe)
				tabela_col=tabela_col+1
				Notas_Tit(1, tabela_col).AddText "<div align=""center""><b>"&co_materia_exibe(d)&"</b></div>", "size=8; indenty=6; alignment=center; html=true", Font
			next	
					
			Set param_materias = PDF.CreateParam	
			param_materias.Set "size=8;expand=true"								
			alunos_encontrados = split(vetor_matriculas, "#$#" )		
			linha=1
			resultados=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, vetor_matriculas, vetor_materia, caminho_nota, tb_nota, 5, 4, 5, "final", 0)		
			resultados_apurados = split(resultados, "#%#" )	
			
			
			Set CON_N = Server.CreateObject("ADODB.Connection") 
			ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
			CON_N.Open ABRIRn				
		
			for a=0 to ubound(alunos_encontrados)	
				param_materias.Add "indenty=2;alignment=right;html=true"
				param_materias.Add "indentx=0"	
				dados_alunos = split(alunos_encontrados(a), "#!#" )
'response.Write(resultados_apurados(a)&"<<BR>")
				resultado_final_aluno=apura_resultado_aluno(curso, co_etapa, resultados_apurados(a))							
'response.Write(dados_alunos(0)&"-"&resultado_final_aluno&"=<BR><BR>")
				resultados_alunos = split(resultados_apurados(a), "#$#" )	
				linha=linha+1
				Set Row = Notas_Tit.Rows.Add(15) ' row height						
				'Notas_Tit(linha, 1).AddText "<div align=""center"">"&dados_alunos(0)&"</div>", param_materias
				 Notas_Tit(linha, 1).AddText "<div align=""center"">"&dados_alunos(4)&"</div>", param_materias				
									
				param_materias.Add "indentx=5"	
				Notas_Tit(linha, 2).AddText dados_alunos(2), param_materias
				coluna=2
				param_materias.Add "indentx=0"
				calcula_frequencia="s"
				for n=0 to ubound(co_materia_exibe)	
					coluna=coluna+1	
					resultados_materia = split(resultados_alunos(n), "#!#" )
					
					if resultados_materia(0)="&nbsp;" or resultados_materia(0)="" or isnull(resultados_materia(0)) then
						calcula_frequencia="n"				
					end if	
					
					media=resultados_materia(0)
'					if curso=1 and co_etapa<6 and (co_materia_verifica(n)="ARTC" or co_materia_verifica(n)="EART" or co_materia_verifica(n)="EFIS" or co_materia_verifica(n)="INGL") then									
'						teste_media = isnumeric(media)							
'						if teste_media=TRUE then							
'							if media > 90 then
'							conceito="E"
'							elseif (media > 70) and (media <= 90) then
'							conceito="MB"
'							elseif (media > 60) and (media <= 70) then							
'							conceito="B"
'							elseif (media > 49) and (media <= 60) then
'							conceito="R"
'							else							
'							conceito="I"
'							end if	
'						end if	
'					else
						conceito=media				
'					end if						
					
					Notas_Tit(linha, coluna).AddText "<div align=""center"">"&conceito&"</DIV>", param_materias	
					'coluna=coluna+1	
					'Notas_Tit(linha, coluna).AddText "<div align=""center"">"&resultados_materia(1)&"</DIV>", param_materias	
				next	
				soma_faltas=0
				if calcula_frequencia="s" then
					co_materia= split(vetor_materia,"#!#")	
					For cm=0 to ubound(co_materia)
						Set RS = Server.CreateObject("ADODB.Recordset")
						SQL = "SELECT * FROM TB_Periodo order by NU_Periodo"
						RS.Open SQL, CON0
						conta=0	
						while not RS.EOF
							periodo_cons=RS("NU_Periodo")	
							if periodo_cons=1 then
								per_faltas="f1"			
							elseif periodo_cons=2 then
								per_faltas="f2"				
							elseif periodo_cons=3 then
								per_faltas="f3"				
							elseif periodo_cons=4 then
								per_faltas="f4"			
							end if
							if periodo_cons<5 then
								faltas=Calcula_Media_T_F_F_N(unidade, curso, co_etapa, turma, dados_alunos(0), co_materia(cm), caminho_nota, tb_nota, per_faltas)				
								response.Write("faltas "&faltas&"<BR>")
								if faltas="" or isnull(faltas) or faltas="&nbsp;" then				
									soma_faltas=soma_faltas	
								else	
									soma_faltas=soma_faltas+faltas	
								end if	
							end if
						RS.Movenext
						Wend											
					next			
		
						
					soma_faltas=soma_faltas*1
					'dias_de_aula_no_ano=200			
					frequencia=((dias_de_aula_no_ano-soma_faltas)/dias_de_aula_no_ano)*100
					if frequencia<100 then
						frequencia=arredonda(frequencia,"mat_dez",1,0)	
						if frequencia=100 then 
							frequencia=100
						end if
					end if	
'					response.Write(frequencia&"="&dias_de_aula_no_ano&"-"&soma_faltas&"<BR>")
'					response.End()
				else
					frequencia=""
				end if				
								
				Notas_Tit(linha, 19).AddText "<div align=""center"">"&frequencia&"</DIV>", param_materias	
				if dados_alunos(3)="E" then
					resultado_final_aluno="Can"
				end if
				Notas_Tit(linha, 20).AddText "<div align=""center"">"&resultado_final_aluno&"</DIV>", param_materias															
			next

			limite=0
			Do While True
				limite=limite+1
				LastRow = Page.Canvas.DrawTable( Notas_Tit, param_table2 )
	
				if LastRow >= Notas_Tit.Rows.Count Then 
			    	Exit Do ' entire table displayed
				else
				
				y_declaracao=margem*4
				SET Param = Pdf.CreateParam("x="&margem&";y="&y_declaracao&"; height=10; width="&area_utilizavel&"; alignment=center; size=8; color=#000000; html=true")
'				Text = "E, para constar, eu ____________________________________________________________________________ Secret&aacute;rio(a), lavrei a presente ata que vai assinada pelo(a) Diretor(a) do Estabelecimento" 			
'				
'				Do While Len(Text) > 0
'					CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
'				 
'					If CharsPrinted = Len(Text) Then Exit Do
'						SET Page = Page.NextPage
'					Text = Right( Text, Len(Text) - CharsPrinted)
'				Loop 
'				y_diretor=margem*2
'								
'				Page.Canvas.SetParams "LineWidth=0.5" 
'				Page.Canvas.SetParams "LineCap=0" 
'
'				With Page.Canvas
'				   .MoveTo 300, y_diretor
'				   .LineTo 542, y_diretor
'				   .Stroke
'				End With 					
'				
'
'				SET Param = Pdf.CreateParam("x="&margem&";y="&y_diretor&"; height=10; width="&area_utilizavel&"; alignment=center; size=8; color=#000000;")
'				Text = "Diretor(a)" 			
'				
'				Do While Len(Text) > 0
'					CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
'				 
'					If CharsPrinted = Len(Text) Then Exit Do
'						SET Page = Page.NextPage
'					Text = Right( Text, Len(Text) - CharsPrinted)
'				Loop 

				SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
				
				Relatorio = "SWD056 - Sistema Web Diretor"
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
				param_table2.Add( "RowTo=1; RowFrom=1" ) ' Row 1 is header.
				param_table2("RowFrom1") = LastRow + 1 ' RowTo1 is omitted and presumed infinite
				
				'NOVO CABEÇALHO==========================================================================================		
				Set Param_Logo_Gde = Pdf.CreateParam
				margem=25			
				area_utilizavel=Page.Width - (margem*2)
				largura_logo_gde=formatnumber(Logo.Width*0.6,0)
				altura_logo_gde=formatnumber(Logo.Height*0.6,0)
		
			   Param_Logo_Gde("x") = margem
			   Param_Logo_Gde("y") = Page.Height - altura_logo_gde - margem
			   Param_Logo_Gde("ScaleX") = 0.6
			   Param_Logo_Gde("ScaleY") = 0.6
			   Page.Canvas.DrawImage Logo, Param_Logo_Gde
		
				x_texto=largura_logo_gde+20
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
				Text = "<center><i><b><font style=""font-size:18pt;"">Ata de Resultados Finais</font></b></i></center>"
				
				
				Do While Len(Text) > 0
					CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
				 
					If CharsPrinted = Len(Text) Then Exit Do
						SET Page = Page.NextPage
					Text = Right( Text, Len(Text) - CharsPrinted)
				Loop 
				
				
				Page.Canvas.SetParams "LineWidth=1" 
				Page.Canvas.SetParams "LineCap=0" 
				inicio_primeiro_separador=x_texto
				altura_primeiro_separador= Page.Height - margem - 17
				With Page.Canvas
				   .MoveTo inicio_primeiro_separador, altura_primeiro_separador
				   .LineTo area_utilizavel+margem, altura_primeiro_separador
				   .Stroke
				End With 	
				
				
				Page.Canvas.SetParams "LineWidth=2" 
				Page.Canvas.SetParams "LineCap=0" 
				altura_segundo_separador= Page.Height - altura_logo_gde -margem - 10
				With Page.Canvas
				   .MoveTo margem, altura_segundo_separador
				   .LineTo area_utilizavel+margem, altura_segundo_separador
				   .Stroke
				End With 	
				
				Set param_table1 = Pdf.CreateParam("width="&area_utilizavel&"; height=40; rows=2; cols=3; border=0; cellborder=0; cellspacing=0;")
				Set Table = Doc.CreateTable(param_table1)
				Table.Font = Font
				y_primeira_tabela=altura_segundo_separador-10
				x_primeira_tabela=margem+5
				With Table.Rows(1)
				   .Cells(1).Width = 50			   		   		   
				   .Cells(2).Width = area_utilizavel-200
				   .Cells(3).Width = 150	
				End With
				Table(1, 1).ColSpan = 3			
				Table(1, 1).AddText "Aos 31 de dezembro de "&ano_letivo&" terminou-se o processo de apura&ccedil;&atilde;o de notas finais e nota global dos alunos deste estabelecimento com os seguintes resultados:", "size=8;html=true", Font	
				
				Table(2, 1).AddText "<b>Ano Letivo:</b>", "size=8;html=true", Font 
				Table(2, 2).AddText mensagem_cabecalho, "size=8;html=true", Font	
				'			Table(2, 3).AddText "<div align=""right""><b>Legenda:</b> Md=M&eacute;dia - Res=Resultado&nbsp;&nbsp;&nbsp;</div>", "size=8;html=true", Font	
				Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 		
				
				'================================================================================================================			
			 	end if
				if limite>100 then
				response.Write("ERRO!")
				response.end()
				end if 
			Loop

			y_declaracao=margem*4
				SET Param = Pdf.CreateParam("x="&margem&";y="&y_declaracao&"; height=10; width="&area_utilizavel&"; alignment=center; size=8; color=#000000; html=true")
				Text = "E, para constar, eu ____________________________________________________________________________ Secret&aacute;rio(a), lavrei a presente ata que vai assinada pelo(a) Diretor(a) do Estabelecimento" 			
				
				Do While Len(Text) > 0
					CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
				 
					If CharsPrinted = Len(Text) Then Exit Do
						SET Page = Page.NextPage
					Text = Right( Text, Len(Text) - CharsPrinted)
				Loop 
				y_diretor=margem*2
								
				Page.Canvas.SetParams "LineWidth=0.5" 
				Page.Canvas.SetParams "LineCap=0" 

				With Page.Canvas
				   .MoveTo 300, y_diretor
				   .LineTo 542, y_diretor
				   .Stroke
				End With 					
				

				SET Param = Pdf.CreateParam("x="&margem&";y="&y_diretor&"; height=10; width="&area_utilizavel&"; alignment=center; size=8; color=#000000;")
				Text = "Diretor(a)" 			
				
				Do While Len(Text) > 0
					CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
				 
					If CharsPrinted = Len(Text) Then Exit Do
						SET Page = Page.NextPage
					Text = Right( Text, Len(Text) - CharsPrinted)
				Loop 

			
			SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
			
			Relatorio = "SWD056 - Sistema Web Diretor"
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
					
			RStabela.Close
			Set RStabela = Nothing							
		End IF					
	End IF		
Next						

	

arquivo="SWD056"
Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

