<%'On Error Resume Next%>
<%
'Server.ScriptTimeout = 60
Server.ScriptTimeout = 600 'valor em segundos
'Emitir Carteirinhas dos Alunos
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../../global/funcoes_diversas.asp"-->
<% 

arquivo="SWD060"

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
	Set Logo = Doc.OpenImage( Server.MapPath( "../img/logo_carteirinha.png") )
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
	CONEXAOA = "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada, TB_Matriculas.CO_Situacao, TB_Alunos.NO_Aluno from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.NU_Unidade = "& unidade&" AND TB_Matriculas.CO_Curso = '"& curso &"' AND TB_Matriculas.CO_Etapa = '"& co_etapa &"' AND TB_Matriculas.CO_Turma = '"& turma &"' order by TB_Alunos.NO_Aluno"
	
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
			gera_pdf="sim"
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
		
			
		if gera_pdf="sim" then	
			SET Page = Doc.Pages.Add(595, 842)
			margem = 25
			espacamento = 155
			qtd_fotos=0
			linha=1
			pagina=1
			impressos=0
			area_utilizavel=Page.Width - (margem*2)
			
			Dim objFSO
				'Create an instance of the FileSystemObject object
			Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

			alunos_encontrados = split(vetor_matriculas, "#$#" )	
		
			Set param_table1 = Pdf.CreateParam("width="&area_utilizavel&"; height="&espacamento&"; rows=1; cols=2; border=1; cellborder=1; cellspacing=0;")
			Set Table = Doc.CreateTable(param_table1)
			Table.Font = Font
			y_primeira_tabela=Page.Height-margem
			x_primeira_tabela=margem
			largura_celula = CInt(area_utilizavel/2)	
			With Table.Rows(1)
				.Cells(1).Width = largura_celula		   		   		   
				.Cells(2).Width = largura_celula
			End With		
				
			x_logo = margem + 10
			x_foto = largura_celula-35
			y_foto_logo = y_primeira_tabela+75

			Set Param_logo = Pdf.CreateParam
			Param_logo("x") = x_logo

			Param_logo("ScaleX") = 0.20
			Param_logo("ScaleY") = 0.20		
			altura_logo = CINT(Logo.Height * 0.20)

			Set Param_foto = Pdf.CreateParam
			Param_foto("x") = x_foto

			Param_foto("ScaleX") = 0.24
			Param_foto("ScaleY") = 0.24

			for a=0 to ubound(alunos_encontrados)	
				dados_alunos = split(alunos_encontrados(a), "#!#" )	
				if dados_alunos(3) = "C" then
					y_foto_logo = y_foto_logo-espacamento
					Param_logo("y") = y_foto_logo
					Param_foto("y") = y_foto_logo

					Page.Canvas.DrawImage Logo, Param_logo	

					if objFSO.FileExists(Server.MapPath( "../img/fotos/aluno/"&dados_alunos(0)&".jpg")) then		
						Set foto = Doc.OpenImage( Server.MapPath( "../img/fotos/aluno/"&dados_alunos(0)&".jpg") )
					else
						Set foto = Doc.OpenImage( Server.MapPath( "../img/fotos/aluno/sem_foto.jpg") )			
					end if
					largura_Foto=formatnumber(foto.Width*0.24,0)
					altura_Foto=formatnumber(foto.Height*0.24,0)


					Page.Canvas.DrawImage foto, Param_foto	

					qtd_fotos = qtd_fotos+1
					impressos = impressos+1
					
					if qtd_fotos>1 then
						y_escola = y_escola-espacamento
					ELSE
						y_escola = y_primeira_tabela-15
					END IF

					SET Param_Escola = Pdf.CreateParam("x="&margem&";y="&y_escola&"; height="&altura_logo&"; width="&largura_celula&"; alignment=left; size=5.5; html=true; color=#000000")

					Escola = "<CENTER><font style=""font-size:9pt;""><b>Col&eacute;gio<br>Maria Raythe</b></font><br><i>Associa&ccedil;&atilde;o Franciscana<br>Nossa Senhora do Amparo</i></CENTER>"
					Do While Len(Escola) > 0
						CharsPrinted = Page.Canvas.DrawText(Escola, Param_Escola, Font )
				 
						If CharsPrinted = Len(Escola) Then Exit Do
							SET Page = Page.NextPage
						Escola = Right( Escola, Len(Escola) - CharsPrinted)
					Loop 

					Set param_table2 = Pdf.CreateParam("width=50; height=15; rows=1; cols=1; border=0.5; cellborder=0.5; cellspacing=0;")
					Set Table2 = Doc.CreateTable(param_table2)
					Table2.Font = Font
					y_segunda_tabela=y_escola-altura_logo+25
					x_segunda_tabela=CInt(area_utilizavel/4)
					
					Table2(1, 1).AddText "<CENTER>"&Session("ano_letivo")&"</CENTER>" , "size=9; html=true; indentx=0; indenty=0", Font
					Page.Canvas.DrawTable Table2, "x="&x_segunda_tabela&", y="&y_segunda_tabela&"" 

					Set param_table3 = Pdf.CreateParam("width=200; height=15; rows=1; cols=1; border=0.5; cellborder=0.5; cellspacing=0;")
					Set Table3 = Doc.CreateTable(param_table3)
					Table3.Font = Font
					y_terceira_tabela=y_segunda_tabela-20
					x_terceira_tabela=CInt(area_utilizavel/8)
					
					Table3(1, 1).AddText "<CENTER>AUTORIZA&Ccedil;&Atilde;O DE SA&Iacute;DA SOZINHO</CENTER>" , "size=9; html=true; indentx=0; indenty=0", Font
					Page.Canvas.DrawTable Table3, "x="&x_terceira_tabela&", y="&y_terceira_tabela&"" 

					y_aluno = y_terceira_tabela-30
					SET Param_Aluno = Pdf.CreateParam("x="&margem&";y="&y_aluno&"; height="&altura_logo&"; width="&largura_celula&"; alignment=left; size=9; html=true; color=#000000")
					Aluno = "<b><CENTER>Aluno:"&dados_alunos(2)&"<BR>&nbsp;<BR>Turma:"&turma&"</CENTER></b>"
					Do While Len(Aluno) > 0
						CharsPrinted = Page.Canvas.DrawText(Aluno, Param_Aluno, Font )
				 
						If CharsPrinted = Len(Aluno) Then Exit Do
							SET Page = Page.NextPage
						Aluno = Right( Aluno, Len(Aluno) - CharsPrinted)
					Loop 
					Table(linha, 2).AddText "<div align=""center"">Associação Franciscana Nossa Senhora do Amparo<BR>&nbsp;<br>Rua Haddock Lobo, 233 – Tijuca – RJ<BR>&nbsp;<br>CEP.: 20.260-141<BR>&nbsp;<br>Telefone: 2264-5474</div>", "size=9;html=true;indenty=40", Font
			
					SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
				
					Relatorio = arquivo&" - Sistema Web Diretor"
					Do While Len(Relatorio) > 0
						CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
				 
						If CharsPrinted = Len(Relatorio) Then Exit Do
							SET Page = Page.NextPage
						Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
					Loop 
				
					Param_Relatorio.Add "alignment=right" 
				
					Do While Len(pagina) > 0
						CharsPrinted = Page.Canvas.DrawText(pagina, Param_Relatorio, Font )
				 
						If CharsPrinted = Len(pagina) Then Exit Do
							SET Page = Page.NextPage
						pagina = Right( pagina, Len(pagina) - CharsPrinted)
					Loop 
				
				
					Param_Relatorio.Add "html=true" 
				
					data_hora = "<center>Impresso em "&data &" &agrave;s "&horario&"</center>"
					Do While Len(Relatorio) > 0
						CharsPrinted = Page.Canvas.DrawText(data_hora, Param_Relatorio, Font )
				 
						If CharsPrinted = Len(data_hora) Then Exit Do
							SET Page = Page.NextPage
						data_hora = Right( data_hora, Len(data_hora) - CharsPrinted)
					Loop 				   ' Display remaining part of table on the next page
				

					if qtd_fotos mod 5 = 0 then
						Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 
						Set Page = Page.NextPage	
						qtd_fotos=0
						linha=1
						pagina=pagina+1
						y_foto_logo = y_primeira_tabela+75
						Set param_table1 = Pdf.CreateParam("width="&area_utilizavel&"; height="&espacamento&"; rows=1; cols=2; border=1; cellborder=1; cellspacing=0;")						
					else
						Table.Rows.Add(espacamento)
						linha = linha+1
					end if

				
					if limite>100 then
						response.Write("ERRO!")
						response.end()
					end if 
				end if
			next	
			Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 
			
			y_declaracao=margem*4
						
			Relatorio = "SWD056 - Sistema Web Diretor"
			Do While Len(Relatorio) > 0
				CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )			
				If CharsPrinted = Len(Relatorio) Then Exit Do
				SET Page = Page.NextPage
				Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
			Loop 
			
			Param_Relatorio.Add "alignment=right" 
			
			Do While Len(pagina) > 0
				CharsPrinted = Page.Canvas.DrawText(pagina, Param_Relatorio, Font )			
				If CharsPrinted = Len(pagina) Then Exit Do
				SET Page = Page.NextPage
				Paginacao = Right( pagina, Len(pagina) - CharsPrinted)
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
		
			'RS2.Close
			'Set RS2 = Nothing
				
			'RS3.Close
			'Set RS3 = Nothing
		
			'RS3m.Close
			'Set RS3m = Nothing
		
			'RS4.Close
			'Set RS4 = Nothing
		
			'RS5.Close
			'Set RS5 = Nothing	
					
			'RStabela.Close
			'Set RStabela = Nothing							
		End IF					
	End IF		
Next						

	


Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

