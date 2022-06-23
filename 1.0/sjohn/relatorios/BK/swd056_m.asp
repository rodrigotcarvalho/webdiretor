<%'On Error Resume Next%>
<%
'Server.ScriptTimeout = 60
Server.ScriptTimeout = 600 'valor em segundos
'ATA DE RESULTADOS
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/parametros.asp"-->
<!--#include file="../inc/funcoes_comuns.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/calculos.asp"-->
<!--#include file="../inc/resultados.asp"-->
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
	'Apenas alunos que estão cursando
	CONEXAOA = "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada, TB_Matriculas.CO_Situacao, TB_Alunos.NO_Aluno from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Situacao='C' AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.NU_Unidade = "& unidade&" AND TB_Matriculas.CO_Curso = '"& curso &"' AND TB_Matriculas.CO_Etapa = '"& co_etapa &"' AND TB_Matriculas.CO_Turma = '"& turma &"' order by TB_Alunos.NO_Aluno"
	'--------------------------------------------------------
	'Todos os alunos
	'CONEXAOA = "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada, TB_Matriculas.CO_Situacao, TB_Alunos.NO_Aluno from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.NU_Unidade = "& unidade&" AND TB_Matriculas.CO_Curso = '"& curso &"' AND TB_Matriculas.CO_Etapa = '"& co_etapa &"' AND TB_Matriculas.CO_Turma = '"& turma &"' order by TB_Alunos.NO_Aluno"
	
	Set RSA = CON1.Execute(CONEXAOA)

	vetor_matriculas="" 
	While Not RSA.EOF
		nu_matricula = RSA("CO_Matricula")
		nome_aluno= RSA("NO_Aluno")			
		nu_chamada = RSA("NU_Chamada")
		situacao = RSA("CO_Situacao")	
		
		nome_aluno=replace_latin_char(nome_aluno,"html")	
		nu_chamada_check=nu_chamada_check*1
		nu_chamada=nu_chamada*1
		if nu_chamada_check = 1 and nu_chamada=nu_chamada_check then
			vetor_matriculas=nu_matricula&"#!#"&nu_chamada&"#!#"&nome_aluno&"#!#"&situacao
		elseif nu_chamada_check = 1 then
			while nu_chamada_check < nu_chamada
				nu_chamada_check=nu_chamada_check+1
			wend 
			vetor_matriculas=nu_matricula&"#!#"&nu_chamada&"#!#"&nome_aluno&"#!#"&situacao
		else
			vetor_matriculas=vetor_matriculas&"#$#"&nu_matricula&"#!#"&nu_chamada&"#!#"&nome_aluno&"#!#"&situacao
		end if
	nu_chamada_check=nu_chamada_check+1		
	RSA.MoveNext
	Wend 

	if curso=0 then

	else
	
		tb_nota=tabela_notas(CON2, unidade, curso, co_etapa, turma, periodo, disciplina, outro)
		caminho_nota=caminho_notas(CON2, tb_nota, outro)	
		
		Set CON_N = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& caminho_nota & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIRn			
			
		tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")
		tp_freq=tipo_divisao_ano(curso,co_etapa,"in_frequencia")
		periodo_m1 = Periodo_Media(tp_modelo,"MA",outro)
		periodo_m2 = Periodo_Media(tp_modelo,"RF",outro)
		periodo_m3 = Periodo_Media(tp_modelo,"MF",outro)			
		
		if tp_modelo="ERRO" or tp_freq="ERRO" then	
			gera_pdf="nao"
		else
			gera_pdf="sim"				
		end if						
		if gera_pdf="sim" then	
		
			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL2 = "SELECT * FROM TB_Unidade WHERE NU_Unidade="& unidade
			RS2.Open SQL2, CON0
							
			no_unidade = RS2("NO_SEDE")		
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
				co_municipio=municipio_unidade
				municipio_unidade=RS3m("NO_Municipio")						

				if bairro_unidade="" or isnull(bairro_unidade) then
				else
				
					Set RS3m = Server.CreateObject("ADODB.Recordset")
					SQL3m = "SELECT * FROM TB_Bairros WHERE SG_UF='"& uf_unidade_municipio &"' AND CO_Municipio="&co_municipio&" AND CO_Bairro="&bairro_unidade
					RS3m.Open SQL3m, CON0
					
					bairro_unidade=RS3m("NO_Bairro")					
					bairro_unidade=" - "&bairro_unidade
				end if
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
			margem_vertical=60			
			area_utilizavel=Page.Width - (margem*2)
			largura_logo_gde=formatnumber(Logo.Width*0.5,0)
			altura_logo_gde=formatnumber(Logo.Height*0.5,0)
	
			Param_Logo_Gde("x") = margem
			Param_Logo_Gde("y") = formatnumber(Page.Height - altura_logo_gde - margem_vertical,0)
			Param_Logo_Gde("ScaleX") = 0.5
			Param_Logo_Gde("ScaleY") = 0.5
			Page.Canvas.DrawImage Logo, Param_Logo_Gde
	
			x_texto=largura_logo_gde+10
			y_texto=formatnumber(Page.Height - margem_vertical,0)'formatnumber(Page.Height - margem,0)
			width_texto=Page.Width -largura_logo_gde - 80


			SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
			Text = "<p><i><b>"&UCASE(no_unidade)&"</b></i></p><br><font style=""font-size:10pt;"">"&endereco_unidade&"</FONT>" 			

			
			Do While Len(Text) > 0
				CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
			 
				If CharsPrinted = Len(Text) Then Exit Do
					SET Page = Page.NextPage
				Text = Right( Text, Len(Text) - CharsPrinted)
			Loop 

			y_texto=y_texto-margem_vertical
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
			inicio_primeiro_separador=largura_logo_gde+10
			altura_primeiro_separador= Page.Height - 60 - 17
			With Page.Canvas
			   .MoveTo inicio_primeiro_separador, altura_primeiro_separador
			   .LineTo area_utilizavel+margem, altura_primeiro_separador
			   .Stroke
			End With 	
	
	
			Page.Canvas.SetParams "LineWidth=2" 
			Page.Canvas.SetParams "LineCap=0" 
			altura_segundo_separador= Page.Height - altura_logo_gde - 60 - 10
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
			Table(1, 1).AddText "Aos 31 de dezembro de "&ano_letivo&" foi encerrado o processo de apura&ccedil;&atilde;o de notas finais dos alunos deste estabelecimento com os seguintes resultados:", "size=8;html=true", Font	

			Table(2, 1).AddText "<b>Ano Letivo:</b>", "size=8;html=true", Font 
			Table(2, 2).AddText mensagem_cabecalho, "size=8;html=true", Font	
			'Table(2, 3).AddText "<div align=""right""><b>Legenda:</b> FRQ = Frequencia / RF = Resultado Final&nbsp;&nbsp;&nbsp;</div>", "size=8;html=true", Font	
			Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 		

	'================================================================================================================			

			Set RS5 = Server.CreateObject("ADODB.Recordset")
			SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND IN_MAE=TRUE order by NU_Ordem_Boletim "
			RS5.Open SQL5, CON0
			co_materia_check=1
			IF RS5.EOF Then
				vetor_materia_exibe="nulo"
			else
				while not RS5.EOF
					co_mat_fil= RS5("CO_Materia")
					co_materia= RS5("CO_Disciplina_Rel")				
					if co_materia_check=1 then
						vetor_materia=co_mat_fil
						vetor_materia_rel=co_materia	
						vetor_materia_carga= "'"&co_mat_fil&"'"
						vetor_materia_filhas=busca_materias_filhas(co_mat_fil)	
						
						if vetor_materia_filhas<>co_mat_fil then
							materia_filhas= split(vetor_materia_filhas,"#!#")
							for mf= 0 to ubound(materia_filhas)
								vetor_materia_carga=vetor_materia_carga&", '"&materia_filhas(mf)&"'"
							next	
						end if									
					else
						vetor_materia=vetor_materia&"#!#"&co_mat_fil
						vetor_materia_rel=vetor_materia_rel&"#!#"&co_materia
						vetor_materia_carga=vetor_materia_carga&", '"&co_mat_fil&"'"
						vetor_materia_filhas=busca_materias_filhas(co_mat_fil)	
						if vetor_materia_filhas<>co_mat_fil then	
							materia_filhas= split(vetor_materia_filhas,"#!#")
							for mf= 0 to ubound(materia_filhas)
								vetor_materia_carga=vetor_materia_carga&", '"&materia_filhas(mf)&"'"
							next									
						end if																			
					end if
					co_materia_check=co_materia_check+1		
					
					
							
				RS5.MOVENEXT
				wend						
			end if

			p_vetor_materia=vetor_materia
			co_materia_exibe=Split(vetor_materia_rel,"#!#")		
			co_materia_verifica=Split(vetor_materia,"#!#")							
			colunas_de_notas=18
			total_de_colunas=20					
			altura_medias=25
			y_segunda_tabela=y_primeira_tabela-40	
			Set param_table2 = Pdf.CreateParam("width="&area_utilizavel&"; height="&altura_medias&"; rows=1; cols="&total_de_colunas&"; border=1; cellborder=0.1; cellspacing=0; x="&margem&"; y="&y_segunda_tabela&"; MaxHeight=285")

			Set Notas_Tit = Doc.CreateTable(param_table2)
			Notas_Tit.Font = Font				
			largura_colunas=(area_utilizavel-40-200)/colunas_de_notas		
			
			With Notas_Tit.Rows(1)
			   .Cells(1).Width = 40
			   .Cells(2).Width = 200			             
				for d=3 to total_de_colunas
				 .Cells(d).Width = largura_colunas					
				next
			End With
			'Notas_Tit(1, 1).RowSpan = 2
			'Notas_Tit(1, 2).RowSpan = 2		
			Notas_Tit(1, 1).AddText "<div align=""center""><b>Matr&iacute;cula</b></div>", "size=8;indenty=6; html=true", Font 
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
			
			for ae=0 to ubound(alunos_encontrados)	
				param_materias.Add "indenty=2;alignment=right;html=true"
				param_materias.Add "indentx=0"				
				dados_alunos = split(alunos_encontrados(ae), "#!#" )
				
				apura_frequencia="s"					
				
				linha=linha+1
				Set Row = Notas_Tit.Rows.Add(15) ' row height						
				Notas_Tit(linha, 1).AddText "<div align=""center"">"&dados_alunos(0)&"</div>", param_materias	
				param_materias.Add "indentx=5"	
				Notas_Tit(linha, 2).AddText dados_alunos(2), param_materias
				coluna=2				
								
				for cmv=0 to ubound(co_materia_verifica)
					compara_m3 = parametros_gerais(unidade, curso, co_etapa, turma, co_materia_verifica(cmv),"compara_m3",0)					
					aproxima_m3 = parametros_gerais(unidade, curso, co_etapa, turma, co_materia_verifica(cmv),"aproxima_m3",0)											
					resultados=Calc_Ter_Media (unidade, curso, co_etapa, turma, tp_modelo, tp_freq, dados_alunos(0), co_materia_verifica(cmv), CON_N, tb_nota, periodo_m3, "ATA", compara_m3, aproxima_m3, outro)
					

					resultados_apurados = split(resultados, "#!#" )	
					if cmv=0 then
						vetor_resultados_apurados=resultados_apurados(1)
						vetor_resultados_completos=resultados						
					else
						vetor_resultados_apurados=vetor_resultados_apurados&"#!#"&resultados_apurados(1)	
						vetor_resultados_completos=vetor_resultados_completos&"#$#"&resultados											
					end if	
					
					if resultados_apurados(0)="&nbsp;" or resultados_apurados(0)="" or isnull(resultados_apurados(0)) then
						apura_frequencia="n"				
					end if				
				next	

	
				carga=carga_aula(curso,co_etapa,vetor_materia_carga,CON0,"frequencia",outro)				

				frequencia=Calcula_Frequencia(unidade, curso, co_etapa, turma, tp_modelo, tp_freq, dados_alunos(0), "&nbsp;", "&nbsp;", CON_N , tb_nota, 0, carga, "aluno", "percent", 0)					
					
				
				resultado_final_aluno=novo2_apura_resultado_aluno(curso, co_etapa, dados_alunos(0), p_vetor_materia, vetor_resultados_completos, frequencia, periodo_m1,  periodo_m2, periodo_m3, "ata", "S", "S", outro)

			
				param_materias.Add "indentx=0"	
				resultados_materia = split(vetor_resultados_completos, "#$#" )	
							
				for cmvr=0 to ubound(co_materia_verifica)
					coluna=coluna+1				
					vetor_exibe_resultados = split(resultados_materia(cmvr), "#!#" )		
					media  = vetor_exibe_resultados(0)
					result = vetor_exibe_resultados(1)
					
					IF co_materia_verifica(cmvr) = "RESUL" then
						if isnumeric(media)	then
							if media=0 then
								conceito = "REP"
							elseif media=100 then
								conceito = "APR"							
							end if
						else	
							conceito = "&nbsp;"							
						end if			
						Notas_Tit(linha, coluna).AddText "<div align=""center"">"&conceito&"</DIV>", param_materias	
					ELSE	
						if resultado_final_aluno="REP" or resultado_final_aluno="Rep" or resultado_final_aluno="APR" or resultado_final_aluno="Apr" then
							if isnumeric(media)	then
								conceito=converte_conceito(unidade, curso, co_etapa, turma, periodo, co_materia_verifica(cmvr), media, outro)	
							else
								conceito="&nbsp;"											
							end if											
									
							Notas_Tit(linha, coluna).AddText "<div align=""center"">"&conceito&"</DIV>", param_materias		
						else
							if result="REP" or result="APR" then
								if isnumeric(media)	then
									conceito=converte_conceito(unidade, curso, co_etapa, turma, periodo, co_materia_verifica(cmvr), media, outro)	
								else
									conceito="&nbsp;"		
								end if											
										
								Notas_Tit(linha, coluna).AddText "<div align=""center"">"&conceito&"</DIV>", param_materias	
							else
								Notas_Tit(linha, coluna).AddText "<div align=""center"">&nbsp;</DIV>", param_materias							
							end if							
						end if	
					END IF	
				next		
				if isnumeric(frequencia) then
					frequencia=formatnumber(frequencia,1)
				end if	

				if dados_alunos(3)<>"C" then	
					Notas_Tit(linha, 19).AddText "<div align=""center"">&nbsp;</DIV>", param_materias				
					Notas_Tit(linha, 20).AddText "<div align=""center"">CAN</DIV>", param_materias																					
				elseif apura_frequencia="s" then
					Notas_Tit(linha, 19).AddText "<div align=""center"">"&frequencia&"</DIV>", param_materias
					Notas_Tit(linha, 20).AddText "<div align=""center"">"&resultado_final_aluno&"</DIV>", param_materias								
				else
					Notas_Tit(linha, 19).AddText "<div align=""center"">&nbsp;</DIV>", param_materias				
					Notas_Tit(linha, 20).AddText "<div align=""center"">&nbsp;</DIV>", param_materias	
				end if																	


			

				
'				If Not IsArray(vetor_resultados) Then 
'					vetor_resultados = Array()
'				End if
'				'Verifica se o Valor que esta sendo inserido já esta no Vetor se estiver entao nao inseri para nao haver duplicidades do vetor
'				If InStr(Join(vetor_resultados), resultado_final_aluno) = 0 Then
'					'Este comando ira preservar o vetor e adciona + 1 valor
'					ReDim preserve vetor_resultados(UBound(vetor_resultados)+1)
'					'Este é o valor que estamos adicionando no vetor
'					vetor_resultados(Ubound(vetor_resultados )) = resultado_final_aluno				
'				end if														
			next
				
			resultados_simulados = "APR#!#REP#!#CAN"
			vetor_resultados = split (resultados_simulados, "#!#")

			limite=0		
			Do While True
				limite=limite+1
				LastRow = Page.Canvas.DrawTable( Notas_Tit, param_table2 )
	
				if LastRow >= Notas_Tit.Rows.Count Then 
			    	Exit Do ' entire table displayed
				else
					y_declaracao1=margem*4
					SET Param = Pdf.CreateParam("x="&margem&";y="&y_declaracao1&"; height=10; width="&area_utilizavel&"; alignment=center; size=8; color=#000000; html=true")


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
			margem_vertical=60		
			area_utilizavel=Page.Width - (margem*2)
			largura_logo_gde=formatnumber(Logo.Width*0.5,0)
			altura_logo_gde=formatnumber(Logo.Height*0.5,0)
	
			Param_Logo_Gde("x") = margem
			Param_Logo_Gde("y") = formatnumber(Page.Height - altura_logo_gde - margem_vertical,0)
			Param_Logo_Gde("ScaleX") = 0.5
			Param_Logo_Gde("ScaleY") = 0.5
			Page.Canvas.DrawImage Logo, Param_Logo_Gde
	
			x_texto=largura_logo_gde+10
			y_texto=formatnumber(Page.Height - margem_vertical,0)'formatnumber(Page.Height - margem,0)
			width_texto=Page.Width -largura_logo_gde - 80
					
					
					SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
					Text = "<p><i><b>"&UCASE(no_unidade)&"</b></i></p><br><font style=""font-size:10pt;"">"&endereco_unidade&"</FONT>" 			
					
					Do While Len(Text) > 0
						CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
					 
						If CharsPrinted = Len(Text) Then Exit Do
							SET Page = Page.NextPage
						Text = Right( Text, Len(Text) - CharsPrinted)
					Loop 
					
					y_texto=y_texto-margem_vertical
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
					inicio_primeiro_separador=largura_logo_gde+10
					altura_primeiro_separador= Page.Height - 60 - 17
					With Page.Canvas
					   .MoveTo inicio_primeiro_separador, altura_primeiro_separador
					   .LineTo area_utilizavel+margem, altura_primeiro_separador
					   .Stroke
					End With 	
					
					
					Page.Canvas.SetParams "LineWidth=2" 
					Page.Canvas.SetParams "LineCap=0" 
					altura_segundo_separador= Page.Height - altura_logo_gde -60 - 10
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
					Table(1, 1).AddText "Aos 31 de dezembro de "&ano_letivo&" foi encerrado o processo de apura&ccedil;&atilde;o de notas finais dos alunos deste estabelecimento com os seguintes resultados:", "size=8;html=true", Font	
		
					Table(2, 1).AddText "<b>Ano Letivo:</b>", "size=8;html=true", Font 
					Table(2, 2).AddText mensagem_cabecalho, "size=8;html=true", Font	
					'Table(2, 3).AddText "<div align=""right""><b>Legenda:</b> FRQ = Frequencia / RF = Resultado Final&nbsp;&nbsp;&nbsp;</div>", "size=8;html=true", Font	
					Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 		
					
					'================================================================================================================
							
			 	end if
				if limite>100 then
					response.Write("ERRO!")
					response.end()
				end if 
			Loop
			
			y_mensagem=y_declaracao1+35
			co_materia_exibe_legenda=Split(vetor_materia_rel,"#!#")	
			co_materia_verifica=Split(p_vetor_materia,"#!#")				
			colunas_legenda=ubound(co_materia_exibe_legenda)+1		
				Set param_table_legenda2 = Pdf.CreateParam("width="&area_utilizavel&"; height=40; rows=4; cols="&colunas_legenda&"; border=0; cellborder=0; cellspacing=0;")
				Set Table = Doc.CreateTable(param_table_legenda2)
				Table.Font = Font
				y_primeira_tabela=altura_segundo_separador-10
				x_primeira_tabela=margem+5
				With Table.Rows(1)
'					valor_celula=area_utilizavel/colunas_legenda
'					for t= 0 to ubound(co_materia_exibe_legenda)
'						t_coluna=c+1									
'					   .Cells(t_coluna).Width = valor_celula			   		   		   
'					next
					 .Cells(1).Width = area_utilizavel
				End With
				disciplinas = ""

				for c = 0 to ubound(co_materia_exibe_legenda)
					c_coluna=c+1
				
					Set RSm = Server.CreateObject("ADODB.Recordset")
					SQLm = "SELECT * FROM TB_Materia where CO_Materia ='"& co_materia_verifica(c) &"'"
					RSm.Open SQLm, CON0

					no_materia= RSm("NO_Materia")					
					if c=ubound(co_materia_exibe_legenda) then
						disciplinas= disciplinas&co_materia_exibe_legenda(c)&"-"&no_materia
					else
						disciplinas= disciplinas&co_materia_exibe_legenda(c)&"-"&no_materia&", "				
					end if					
				next
				Table(1, 1).AddText "<b>Legendas:</b>" , "size=7;html=true", Font			
				Table(2, 1).AddText "FRQ = Frequ&ecirc;ncia, RF = Resultado Final" , "size=7;html=true", Font							
				Table(3, 1).AddText "Disciplinas: "&disciplinas , "size=7;html=true; expand=true", Font	
				resultados_apurados = split(resultados, "#$#" )	
				Vetor = Array() 
				Vetor = Empty
				legenda_resultados=""
				for r = 0 to ubound (vetor_resultados)
					If Not IsArray(Vetor) Then 
						Vetor = Array() 
					End if				
					if vetor_resultados(r) ="APR" or vetor_resultados(r) = "REP" or vetor_resultados(r) = "CAN" then	
						If InStr(Join(Vetor), vetor_resultados(r)) = 0 Then
							ReDim preserve Vetor(UBound(Vetor)+1)
							Vetor(Ubound(Vetor)) = vetor_resultados(r)
						end if	
					end if	
				Next
				for v = 0 to ubound (Vetor)	
					if Vetor(v) = "APR" then
						nome_resultado="Aprovado"
					elseif Vetor(v) = "REP" then
						nome_resultado="Reprovado"
					elseif Vetor(v) = "CAN" then
						nome_resultado="Matr&iacute;cula Cancelada"	
					else
																
					end if		
					if v=ubound(Vetor) then
						legenda_resultados= legenda_resultados&Vetor(v)&" - "&nome_resultado
					else
						legenda_resultados= legenda_resultados&Vetor(v)&" - "&nome_resultado&", "				
					end if		
				next	
				Table(4, 1).AddText "Resultados: "& legenda_resultados, "size=7;html=true", Font 
'				Table(2, 2).AddText mensagem_cabecalho, "size=8;html=true", Font	
'				Table(2, 3).AddText "<div align=""right""><b>Legenda:</b> FRQ = Frequencia / RF = Resultado Final&nbsp;&nbsp;&nbsp;</div>", "size=8;html=true", Font	
				Page.Canvas.DrawTable Table, "x="&margem&", y="&y_mensagem&"" 		
	

			y_declaracao=margem*3
				SET Param = Pdf.CreateParam("x="&margem&";y="&y_declaracao&"; height=10; width="&area_utilizavel&"; alignment=center; size=8; color=#000000; html=true")
				Text = " E, para constar, eu ____________________________________________________________________________ Secret&aacute;rio(a), lavrei a presente ata que vai assinada pelo(a) Diretor(a) do Estabelecimento" 			
				
				Do While Len(Text) > 0
					CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
				 
					If CharsPrinted = Len(Text) Then Exit Do
						SET Page = Page.NextPage
					Text = Right( Text, Len(Text) - CharsPrinted)
				Loop 
				y_diretor=(margem*2)-10
								
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
		
'			RS3m.Close
'			Set RS3m = Nothing
		
			RS4.Close
			Set RS4 = Nothing
		
			RS5.Close
			Set RS5 = Nothing	
										
		End IF					
	End IF		
Next						

	

arquivo="SWD056"
Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

