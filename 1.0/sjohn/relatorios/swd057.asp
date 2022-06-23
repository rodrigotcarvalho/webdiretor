<%
'Server.ScriptTimeout = 60
Server.ScriptTimeout = 1200 'valor em segundos
'Livro de Registro de Matrículas
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/funcoes2.asp"-->

<% 

response.Charset="ISO-8859-1"
opt= request.QueryString("opt")
ori= request.QueryString("ori")
unidade_form=request.QueryString("un")

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
mes=mes*1
meswrt="0"&mes
else
meswrt=mes
end if
if min<10 then
min=min*1
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
		
		if unidade_form<2 then
			sql_unidade = " IN (0,1)" 
		else
			sql_unidade = " = "&unidade_form
		end if			
	
	Set RSA = Server.CreateObject("ADODB.Recordset")
	'CONEXAOA = "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada, TB_Matriculas.CO_Situacao, TB_Matriculas.NU_Unidade, TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma, TB_Matriculas.DA_Rematricula, TB_Alunos.NO_Aluno, TB_Alunos.SG_UF_Natural,TB_Alunos.CO_Municipio_Natural from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.NU_Unidade = "& unidade_form&" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula order by TB_Matriculas.NU_Unidade ASC, TB_Matriculas.CO_Curso ASC, TB_Matriculas.CO_Etapa ASC, TB_Matriculas.CO_Turma ASC, TB_Alunos.NO_Aluno ASC"
	CONEXAOA = "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada, TB_Matriculas.CO_Situacao, TB_Matriculas.NU_Unidade, TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma, TB_Matriculas.DA_Rematricula, TB_Alunos.NO_Aluno, TB_Alunos.SG_UF_Natural,TB_Alunos.CO_Municipio_Natural, TB_Alunos.NO_Pai, TB_Alunos.NO_Mae from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.NU_Unidade "& sql_unidade&" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula order by TB_Matriculas.DA_Rematricula ASC, TB_Alunos.CO_Matricula ASC"	
	response.Write(CONEXAOA)
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
			data_m(0)=data_m(0)*1		
			dia_m="0"&data_m(0)
		else	
			dia_m=data_m(0)
		end if		

		if data_m(1)<10 then
			data_m(1)=data_m(1)*1
			mes_m="0"&data_m(1)
		else	
			mes_m=data_m(1)
		end if				
		
		dt_matricula=dia_m&"/"&mes_m&"/"&data_m(2)
		
		if co_situacao="C" then
			no_situacao="Efetivado"
		elseif co_situacao="E" then
			no_situacao="Cancelado"
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

		
		if unidade_form<2 then
			unidade_cabecalho = 0 
		else
			unidade_cabecalho=unidade_form
		end if	
			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL2 = "SELECT * FROM TB_Unidade WHERE NU_Unidade="&unidade_cabecalho
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
			
			if bairro_unidade="" or isnull(bairro_unidade)then
			else
				Set RS3m = Server.CreateObject("ADODB.Recordset")
				SQL3m = "SELECT * FROM TB_Bairros WHERE SG_UF='"& uf_unidade_municipio &"' AND CO_Municipio="&municipio_unidade&" AND CO_Bairro="&bairro_unidade
				RS3m.Open SQL3m, CON0
				
				no_bairro_unidade=RS3m("NO_Bairro")				
				bairro_unidade=" - "&no_bairro_unidade
			end if
			

			
			if municipio_unidade="" or isnull(municipio_unidade) or uf_unidade_municipio="" or isnull(uf_unidade_municipio)then
			else
				Set RS3m = Server.CreateObject("ADODB.Recordset")
				SQL3m = "SELECT * FROM TB_Municipios WHERE SG_UF='"& uf_unidade_municipio &"' AND CO_Municipio="&municipio_unidade
				RS3m.Open SQL3m, CON0
				
				municipio_unidade=RS3m("NO_Municipio")						
			end if
			endereco_unidade=rua_unidade&numero_unidade&complemento_unidade&bairro_unidade&cep_unidade&"<br>"&municipio_unidade&uf_unidade					
						

'
'			no_curso= no_etapa&"&nbsp;"&co_concordancia_curso&"&nbsp;"&no_curso
'			texto_turma = "&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;<b>Turma</b> "&turma
'
'			mensagem_cabecalho=ano_letivo&"&nbsp;&nbsp;&nbsp;&nbsp;-&nbsp;&nbsp;&nbsp;&nbsp;"&no_curso&texto_turma
'	
			SET Page = Doc.Pages.Add(842, 595)
					
	'CABEÇALHO==========================================================================================		
			Set Param_Logo_Gde = Pdf.CreateParam
			margem=25			
			area_utilizavel=Page.Width - (margem*2)
			
			largura_logo_gde=formatnumber(Logo.Width*0.5,0)
			altura_logo_gde=formatnumber(Logo.Height*0.5,0)
	
		   Param_Logo_Gde("x") = margem
		   Param_Logo_Gde("y") = Page.Height - altura_logo_gde - margem
		   Param_Logo_Gde("ScaleX") = 0.5
		   Param_Logo_Gde("ScaleY") = 0.5
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
			Text = "<center><i><b><font style=""font-size:18pt;"">Livro de Registro de Matr&iacute;cula - Ano Letivo "&ano_letivo&"</font></b></i></center>"
			
			
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


	'================================================================================================================			

			colunas_de_notas=21
			total_de_colunas=23					
			altura_medias=20
			y_segunda_tabela=altura_segundo_separador-10	
			Set param_table2 = Pdf.CreateParam("width="&area_utilizavel&"; height="&altura_medias&"; rows=1; cols=13; border=0; cellborder=0.5; cellspacing=0; x="&margem&"; y="&y_segunda_tabela&"; MaxHeight=450")

			Set Notas_Tit = Doc.CreateTable(param_table2)
			Notas_Tit.Font = Font				
			largura_colunas=(area_utilizavel-50-210)/colunas_de_notas		
			
			With Notas_Tit.Rows(1)
			   .Cells(1).Width = 20
			   .Cells(2).Width = 45	
			   .Cells(3).Width = 60
			   .Cells(4).Width = 45			             
			   .Cells(5).Width = 30
			   .Cells(6).Width = 60			             
			   .Cells(7).Width = 50
			   .Cells(8).Width = 35
			   .Cells(9).Width = 137
			   .Cells(10).Width = 55			   			             
			   .Cells(11).Width = 50
			   .Cells(12).Width = 70			             
			   .Cells(13).Width = 145	
			End With
			Notas_Tit(1, 8).colspan = 2	
			Notas_Tit(1, 1).AddText "<div align=""center"">N&ordm;</div>", "size=9;indenty=2; html=true", Font 
			Notas_Tit(1, 2).AddText "<div align=""center"">Data</div>", "size=9;alignment=center; indenty=2;html=true", Font 
			Notas_Tit(1, 3).AddText "<div align=""center"">Curso</div>", "size=9;alignment=center; indenty=2;html=true", Font 
			Notas_Tit(1, 4).AddText "<div align=""center"">Etapa</div>", "size=9;alignment=center; indenty=2;html=true", Font 
			Notas_Tit(1, 5).AddText "<div align=""center"">Turno</div>", "size=9;alignment=center; indenty=2;html=true", Font 
			Notas_Tit(1, 6).AddText "<div align=""center"">Turma</div>", "size=9;alignment=center; indenty=2;html=true", Font 
			Notas_Tit(1, 7).AddText "<div align=""center"">Matr&iacute;cula</div>", "size=9;alignment=center; indenty=2;html=true", Font 
			Notas_Tit(1, 8).AddText "<div align=""left"">Nome</div>", "size=9; indenty=2;html=true", Font 
			Notas_Tit(1, 10).AddText "<div align=""center"">Situa&ccedil;&atilde;o</div>", "size=9;alignment=center; indenty=2;html=true", Font 
			Notas_Tit(1, 11).AddText "<div align=""center"">Nascimento</div>", "size=9;alignment=center; indenty=2;html=true", Font 
			Notas_Tit(1, 12).AddText "<div align=""center"">Naturalidade</div>", "size=9;alignment=center; indenty=2;html=true", Font 
			Notas_Tit(1, 13).AddText "<div align=""left"">Endere&ccedil;o</div>", "size=9; indenty=2;html=true", Font 
			Notas_Tit.At(1, 1).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False, BottomColor=Black" 
			Notas_Tit.At(1, 2).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False, BottomColor=Black" 
			Notas_Tit.At(1, 3).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False, BottomColor=Black" 
			Notas_Tit.At(1, 4).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False, BottomColor=Black" 
			Notas_Tit.At(1, 5).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False, BottomColor=Black" 
			Notas_Tit.At(1, 6).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False, BottomColor=Black" 
			Notas_Tit.At(1, 7).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False, BottomColor=Black" 
			Notas_Tit.At(1, 8).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False, BottomColor=Black" 
			Notas_Tit.At(1, 9).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False, BottomColor=Black" 
			Notas_Tit.At(1, 10).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False, BottomColor=Black" 
			Notas_Tit.At(1, 11).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False, BottomColor=Black" 
			Notas_Tit.At(1, 12).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False, BottomColor=Black" 
			Notas_Tit.At(1, 13).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False, BottomColor=Black" 

'			tabela_col=2
'				for d=0 to ubound(co_materia_exibe)
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
			param_materias.Set "size=8;expand=false" 			
''response.Flush()										
			alunos_encontrados = split(vetor_matriculas, "#$#" )		
			linha=1
			compara_curso_al = ""
			compara_etapa_al = ""
			compara_turma_al = ""
			compara_turno_al = ""			
			for a=0 to ubound(alunos_encontrados)	
				param_materias.Add "indenty=2;alignment=right;html=true"
				param_materias.Add "indentx=0"	
				dados_alunos = split(alunos_encontrados(a), "#!#" )
				
				if compara_curso_al <> dados_alunos(6) then
					compara_curso_al = dados_alunos(6) 
					Set RS3 = Server.CreateObject("ADODB.Recordset")
					SQL3 = "SELECT * FROM TB_Curso WHERE CO_Curso='"& dados_alunos(6) &"'"
					RS3.Open SQL3, CON0
					
					if RS3.EOF then
						no_curso_aluno=dados_alunos(6) 			
					else	
						no_curso_aluno= RS3("NO_Curso")
					end if	
				end if						
					
				if compara_etapa_al <> dados_alunos(7) then
					compara_etapa_al = dados_alunos(7)			
					Set RS4 = Server.CreateObject("ADODB.Recordset")
					SQL4 = "SELECT * FROM TB_Etapa WHERE CO_Curso='"& dados_alunos(6) &"' AND CO_Etapa ='"& dados_alunos(7) &"'"
					RS4.Open SQL4, CON0			
		
					if RS4.EOF then
						no_etapa_aluno = dados_alunos(7) 			
					else	
						no_etapa_aluno = RS4("NO_Etapa")
					end if	
				end if							
					
				if compara_turma_al<>dados_alunos(8) then		
					compara_turma_al = dados_alunos(8) 
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
				end if									
				
				
				
				Set RS7 = Server.CreateObject("ADODB.Recordset")
				SQL7 = "SELECT * FROM TB_Contatos WHERE CO_Matricula="& dados_alunos(1) &" AND TP_Contato='ALUNO'"
				RS7.Open SQL7, CONCONT	
							
				if RS7.EOF then		
				else
					nasc_aluno = RS7("DA_Nascimento_Contato")	
					rua_aluno = RS7("NO_Logradouro_Res")
					rua_num_aluno = RS7("NU_Logradouro_Res")				
					compl_aluno = RS7("TX_Complemento_Logradouro_Res")	
					bairro_aluno = RS7("CO_Bairro_Res")				
					cep_aluno = RS7("CO_CEP_Res")	
					
					data_n=split(nasc_aluno,"/")
					if data_n(0)<10 then
						data_n(0)=data_n(0)*1						
						dia_n="0"&data_n(0)
					else	
						dia_n=data_n(0)
					end if		
			
					if data_n(1)<10 then
						data_n(1)=data_n(1)*1					
						mes_n="0"&data_n(1)
					else	
						mes_n=data_n(1)
					end if				
					
					nasc_aluno=dia_n&"/"&mes_n&"/"&data_n(2)							
				
					Set RS7a = Server.CreateObject("ADODB.Recordset")
					SQL7a = "SELECT * FROM TB_Bairros WHERE CO_Bairro="& bairro_aluno
					RS7a.Open SQL7a, CON0	
									
					cidade_aluno = RS7a("CO_Municipio")
					no_bairro_aluno = RS7a("NO_Bairro")
					uf_aluno = RS7a("SG_UF")
					rua_cep=left(cep_aluno,5)&"-"&right(cep_aluno,3)
	
					Set RS7b = Server.CreateObject("ADODB.Recordset")
					SQL7b = "SELECT * FROM TB_Municipios WHERE CO_Municipio="& cidade_aluno
					RS7b.Open SQL7b, CON0					
					no_cidade_aluno = RS7b("NO_Municipio")
						
					endereco_aluno=	rua_aluno&", "&rua_num_aluno&", "&compl_aluno&". "&no_bairro_aluno&", "&no_cidade_aluno&" - "&uf_aluno&". CEP: "&rua_cep
				end if	
	
				compara_ocupacao = ""

				Set RS8 = Server.CreateObject("ADODB.Recordset")
				SQL8 = "SELECT CO_Ocupacao FROM TB_Contatos WHERE CO_Matricula="& dados_alunos(1) &" AND TP_Contato='PAI'"
				RS8.Open SQL8, CONCONT	
							
				if RS8.EOF then		
				else
					'pai_aluno = RS8("NO_Contato")	
					ocupacao_p = RS8("CO_Ocupacao")	
					if compara_ocupacao<>ocupacao_p then	
						compara_ocupacao=ocupacao_p
						if isnull(ocupacao_p) or ocupacao_p="" then
							ocupacao_pai = ""
						else
							Set RS8a = Server.CreateObject("ADODB.Recordset")
							SQL8a = "SELECT * FROM TB_Ocupacoes WHERE CO_Ocupacao="& ocupacao_p
							RS8a.Open SQL8a, CON0	
							ocupacao_p_n = RS8a("NO_Ocupacao")
							ocupacao_pai = " ("&ocupacao_p_n&")"
						end if	
					end if				
				end if						
				
				Set RS9 = Server.CreateObject("ADODB.Recordset")
				SQL9 = "SELECT CO_Ocupacao FROM TB_Contatos WHERE CO_Matricula="& dados_alunos(1) &" AND TP_Contato='MAE'"
				RS9.Open SQL9, CONCONT	
							
				if RS9.EOF then		
				else
					'mae_aluno = RS9("NO_Contato")	
					ocupacao_m = RS9("CO_Ocupacao")
					if compara_ocupacao<>ocupacao_m then	
						compara_ocupacao=ocupacao_m					
						if isnull(ocupacao_m) or ocupacao_m="" then
							ocupacao_mae = ""
						else
							Set RS8a = Server.CreateObject("ADODB.Recordset")
							SQL8a = "SELECT * FROM TB_Ocupacoes WHERE CO_Ocupacao="& ocupacao_m
							RS8a.Open SQL8a, CON0	
							ocupacao_m_n = RS8a("NO_Ocupacao")
							ocupacao_mae = " ("&ocupacao_m_n&")"
						end if	
					end if	
											
				end if				
			
											
				linha=linha+1
				Set Row = Notas_Tit.Rows.Add(13) ' row height	
				Notas_Tit(linha, 1).RowSpan = 3
				Notas_Tit(linha, 2).RowSpan = 3
				Notas_Tit(linha, 3).RowSpan = 3	
				Notas_Tit(linha, 4).RowSpan = 3	
				Notas_Tit(linha, 5).RowSpan = 3
				Notas_Tit(linha, 6).RowSpan = 3		
				Notas_Tit(linha, 7).RowSpan = 3	
				Notas_Tit(linha, 10).RowSpan = 3	
				Notas_Tit(linha, 11).RowSpan = 3
				Notas_Tit(linha, 12).RowSpan = 3		
				Notas_Tit(linha, 13).RowSpan = 3
				Notas_Tit(linha, 8).ColSpan = 2	
				
				param_materias.Add "expand=false" 												
				Notas_Tit(linha, 1).AddText "<div align=""center"">"&dados_alunos(0)&"</div>", param_materias
				Notas_Tit(linha, 2).AddText "<div align=""center"">"&dados_alunos(9)&"</div>", param_materias	
				Notas_Tit(linha, 3).AddText "<div align=""center"">"&no_curso_aluno&"</div>", param_materias
				Notas_Tit(linha, 4).AddText "<div align=""center"">"&no_etapa_aluno&"</div>", param_materias	
				Notas_Tit(linha, 5).AddText "<div align=""center"">"&no_turno_aluno&"</div>", param_materias																						
				Notas_Tit(linha, 6).AddText "<div align=""center"">"&dados_alunos(8)&"</div>", param_materias		
				Notas_Tit(linha, 7).AddText "<div align=""center"">"&dados_alunos(1)&"</div>", param_materias	
				param_materias.Add "indentx=2"	
				param_materias.Add "expand=true" 	
				Notas_Tit(linha, 8).AddText "<div align=""left"">"&dados_alunos(3)&"</div>", param_materias	
				param_materias.Add "expand=false" 	
				Notas_Tit(linha, 10).AddText "<div align=""center"">"&dados_alunos(4)&"</div>", param_materias	
				Notas_Tit(linha, 11).AddText "<div align=""center"">"&nasc_aluno&"</div>", param_materias	
				Notas_Tit(linha, 12).AddText "<div align=""center"">"&dados_alunos(10)&"</div>", param_materias	
				Notas_Tit(linha, 13).AddText "<div align=""left"">"&endereco_aluno&"</div>", param_materias	
				Notas_Tit.At(linha, 1).SetBorderParams "Left=False, Right=False, Top=False, Bottom=True, BottomColor=Black" 
				Notas_Tit.At(linha, 2).SetBorderParams "Left=False, Right=False, Top=False, Bottom=True, BottomColor=Black" 
				Notas_Tit.At(linha, 3).SetBorderParams "Left=False, Right=False, Top=False, Bottom=True, BottomColor=Black" 
				Notas_Tit.At(linha, 4).SetBorderParams "Left=False, Right=False, Top=False, Bottom=True, BottomColor=Black" 
				Notas_Tit.At(linha, 5).SetBorderParams "Left=False, Right=False, Top=False, Bottom=True, BottomColor=Black" 
				Notas_Tit.At(linha, 6).SetBorderParams "Left=False, Right=False, Top=False, Bottom=True, BottomColor=Black" 
				Notas_Tit.At(linha, 8).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False, BottomColor=Black" 
				Notas_Tit.At(linha, 7).SetBorderParams "Left=False, Right=False, Top=False, Bottom=True, BottomColor=Black" 
				Notas_Tit.At(linha, 9).SetBorderParams "Left=False, Right=False, Top=False, Bottom=True, BottomColor=Black" 
				Notas_Tit.At(linha, 10).SetBorderParams "Left=False, Right=False, Top=False, Bottom=True, BottomColor=Black" 
				Notas_Tit.At(linha, 11).SetBorderParams "Left=False, Right=False, Top=False, Bottom=True, BottomColor=Black" 
				Notas_Tit.At(linha, 12).SetBorderParams "Left=False, Right=False, Top=False, Bottom=True, BottomColor=Black" 
				Notas_Tit.At(linha, 13).SetBorderParams "Left=False, Right=False, Top=False, Bottom=True, BottomColor=Black" 
				
				pai_aluno = dados_alunos(11) 
				mae_aluno = dados_alunos(12) 

				param_materias.Add "expand=true" 
				Set Row = Notas_Tit.Rows.Add(13) ' row height	
				proxima_linha=linha+1
				Notas_Tit(proxima_linha, 8).AddText "<div align=""left"">Filia&ccedil;&atilde;o:</div>", param_materias		
				Notas_Tit.At(proxima_linha, 8).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False, BottomColor=Black" 
				'Notas_Tit(proxima_linha, 9).AddText "<div align=""left"">"&pai_aluno&ocupacao_pai&"</div>", param_materias
				Notas_Tit(proxima_linha, 9).AddText "<div align=""left"">"&pai_aluno&"</div>", param_materias								
				Notas_Tit.At(proxima_linha, 9).SetBorderParams "Left=False, Right=False, Top=False, Bottom=False, BottomColor=Black" 


				Set Row = Notas_Tit.Rows.Add(20) ' row height	
				proxima_linha=proxima_linha+1
				Notas_Tit.At(proxima_linha, 8).SetBorderParams "Left=False, Right=False, Top=False, Bottom=True, BottomColor=Black" 
				'Notas_Tit(proxima_linha, 9).AddText "<div align=""left"">"&mae_aluno&ocupacao_mae&"</div>", param_materias	
				Notas_Tit(proxima_linha, 9).AddText "<div align=""left"">"&mae_aluno&"</div>", param_materias										
				Notas_Tit.At(proxima_linha, 9).SetBorderParams "Left=False, Right=False, Top=False, Bottom=True, BottomColor=Black" 
				linha=proxima_linha
				param_materias.Add "expand=false" 
'				coluna=2
'				param_materias.Add "indentx=0"
'				calcula_frequencia="s"
'					coluna=coluna+1	
'					Notas_Tit(linha, coluna).AddText "<div align=""center""></DIV>", param_materias	
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
					
					Relatorio = "SWD057 - Sistema Web Diretor"
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
			largura_logo_gde=formatnumber(Logo.Width*0.5,0)
			altura_logo_gde=formatnumber(Logo.Height*0.5,0)
	
		   Param_Logo_Gde("x") = margem
		   Param_Logo_Gde("y") = Page.Height - altura_logo_gde - margem
		   Param_Logo_Gde("ScaleX") = 0.5
		   Param_Logo_Gde("ScaleY") = 0.5
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
			Text = "<center><i><b><font style=""font-size:18pt;"">Livro de Registro de Matr&iacute;cula - Ano Letivo "&ano_letivo&"</font></b></i></center>"
			
			
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

	'================================================================================================================			
			 	end if
'				if limite>300 then
'					response.Write("ERRO!")
'					response.end()
'				end if 
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
	
								

Server.ScriptTimeout = 60	

arquivo="SWD057"
Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

