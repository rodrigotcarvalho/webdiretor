<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 180 'valor em segundos
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/funcoes2.asp"-->

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
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

if ori="ede" then
origem="../ws/doc/ofc/ede/"
end if

if mes<10 then
mes="0"&mes
end if

data = dia &"/"& mes &"/"& ano

if mes=1 then
	mes_extenso="Janeiro"
elseif mes=2 then
	mes_extenso="Fevereiro"
elseif mes=3 then
	mes_extenso="Mar&ccedil;o"
elseif mes=4 then
	mes_extenso="Abril"
elseif mes=5 then
	mes_extenso="Maio"
elseif mes=6 then
	mes_extenso="Junho"
elseif mes=7 then
	mes_extenso="Julho"
elseif mes=8 then
	mes_extenso="Agosto"
elseif mes=9 then
	mes_extenso="Setembro"
elseif mes=10 then
	mes_extenso="Outubro"
elseif mes=11 then
	mes_extenso="Novembro"
elseif mes=12 then
	mes_extenso="Dezembro"
end if	
data_extenso="Rio de Janeiro, "&dia &" de "& mes_extenso &" de "& ano
if min<10 then
min="0"&min
end if

horario = hora & ":"& min

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
		
	
if opt="01" then
	cod_cons=request.QueryString("cod_cons")
	co_declaracao= request.QueryString("dcl")	

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod_cons
		RS1.Open SQL1, CON1
		
		if RS1.EOF then
			response.redirect(origem&"index.asp?nvg="&nvg&"&opt=err4")
		end if

		If Not IsArray(alunos_encontrados) Then alunos_encontrados = Array() End if	
		ReDim preserve alunos_encontrados(UBound(alunos_encontrados)+1)	
		alunos_encontrados(Ubound(alunos_encontrados)) = cod_cons	

	
elseif opt="02" then
	obr=request.QueryString("obr")
	dados_informados = split(obr, "_" )
	declaracao_terceiro_ano=request.QueryString("dcl")
	if isnull(declaracao_terceiro_ano) or declaracao_terceiro_ano="" then
		gera_declaracao_terceiro_ano="nao"
		unidade=dados_informados(0)
		curso=dados_informados(1)
		co_etapa=dados_informados(2)
		turma=dados_informados(3)
		co_declaracao=dados_informados(4)
	else
		gera_declaracao_terceiro_ano="sim"
		co_declaracao=declaracao_terceiro_ano
	end if
	IF ((isnull(unidade) or unidade="") and (isnull(curso) or curso="") and (isnull(co_etapa) or co_etapa="") and (isnull(turma) or turma="")) or gera_declaracao_terceiro_ano="sim" THEN
		if co_declaracao="swd204" or co_declaracao="swd209"  then
			nu_chamada_check = 1
			Set RSA = Server.CreateObject("ADODB.Recordset")
			SQL_BUSCA_ALUNOS= "Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula AND TB_Matriculas.CO_Curso = '2' AND TB_Matriculas.CO_Etapa = '3' order by TB_Matriculas.NU_Unidade,TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma, TB_Alunos.NO_Aluno"		
			Set RSA = CON1.Execute(SQL_BUSCA_ALUNOS)
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
		else
			response.redirect(origem&"index.asp?nvg="&nvg&"&opt=err3")				
		END IF		
	ELSE	
	
		if unidade="999990" or unidade="" or isnull(unidade) then
			SQL_BUSCA_ALUNOS="NULO"
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
		
		SQL_BUSCA_ALUNOS= SQL_ALUNOS&SQL_CURSO&SQL_ETAPA&SQL_TURMA&" order by TB_Matriculas.NU_Unidade,TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma, TB_Alunos.NO_Aluno"
		end if
	
		if SQL_BUSCA_ALUNOS="NULO" then
		else
		
		nu_chamada_check = 1
			Set RSA = Server.CreateObject("ADODB.Recordset")
			CONEXAOA = SQL_BUSCA_ALUNOS
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
			RSA.Close
			Set RSA = Nothing	
		end if	
	END IF	
	if vetor_matriculas="" then
		alunos_encontrados = Array() 
	else
		alunos_encontrados = split(vetor_matriculas, "#!#" )	
	end if	
end if	


'RESPONSE.END()
if ubound(alunos_encontrados)=-1 then
	if co_declaracao="swd200" or co_declaracao="swd201"  then
		response.redirect(origem&"index.asp?nvg="&nvg&"&opt=err1")
	elseif co_declaracao="swd202" or co_declaracao="swd203"  then
		response.redirect(origem&"index.asp?nvg="&nvg&"&opt=err2")
	elseif co_declaracao="swd204" or co_declaracao="swd209"  then
		response.redirect(origem&"index.asp?nvg="&nvg&"&opt=err3")	
	end if
else

	Set RSc = Server.CreateObject("ADODB.Recordset")
	SQLc = "SELECT * FROM TB_Cabecalhos WHERE CO_Documento=1"
	RSc.Open SQLc, CON0	
	
	nome_documento=RSc("NO_Documento")
	cabec_1=RSc("NO_Cabec_1")
	cabec_2=RSc("NO_Cabec_2")	
	cabec_3=RSc("NO_Cabec_3")	
	cabec_4=RSc("NO_Cabec_4")	
	cabec_5=RSc("NO_Cabec_5")	
	
	if nome_documento="" or isnull(nome_documento) then
	else
		nome_documento=replace_latin_char(UCASE(nome_documento),"html")
	end if
	if cabec_1="" or isnull(cabec_1)  then
	else
		cabec_1=replace_latin_char(cabec_1,"html")
	end if
	if cabec_2="" or isnull(cabec_2)  then
	else
		cabec_2=replace_latin_char(cabec_2,"html")
	end if
	if cabec_3="" or isnull(cabec_3)  then
	else
		cabec_3=replace_latin_char(cabec_3,"html")
	end if
	if cabec_4="" or isnull(cabec_4) then
	else
		cabec_4=replace_latin_char(cabec_4,"html")
	end if
	if cabec_5="" or isnull(cabec_5) then
	else
		cabec_5=replace_latin_char(cabec_5,"html")	
	end if

	relatorios_gerados=-1			
	For i=0 to ubound(alunos_encontrados)	
		cod_cons=alunos_encontrados(i)
		gera_relatorio_aluno="s"	
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Periodo ORDER BY NU_Periodo"
		RS0.Open SQL0, CON0
		check_periodo=1
		WHILE NOT RS0.EOF
			periodo=RS0("NU_Periodo")
			if check_periodo=1 then
				vetor_periodo=periodo
			else
				vetor_periodo=vetor_periodo&"#!#"&periodo
			end if
			check_periodo=check_periodo+1 
		RS0.MOVENEXT
		WEND				
				
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod_cons
		RS1.Open SQL1, CON1
		
		if RS1.EOF then
			response.redirect(origem&"index.asp?nvg="&nvg&"&opt=err4")
		else
		
			ano_aluno = RS1("NU_Ano")
			rematricula = RS1("DA_Rematricula")
			situacao = RS1("CO_Situacao")
			encerramento= RS1("DA_Encerramento")
			unidade= RS1("NU_Unidade")
			curso= RS1("CO_Curso")
			co_etapa= RS1("CO_Etapa")
			turma= RS1("CO_Turma")
			cham= RS1("NU_Chamada")
			curso=curso*1
			if curso=2 then
				etapa=etapa*1
			end if
			
			if (co_declaracao="swd204" or co_declaracao="swd209") AND (curso<>2 and etapa<>3) then
				gera_relatorio_aluno="n"	
			elseif (co_declaracao="swd200" or co_declaracao="swd201" or co_declaracao="swd202" or co_declaracao="swd203" or co_declaracao="swd204" or co_declaracao="swd208" or co_declaracao="swd209") AND (curso=0) then			
				gera_relatorio_aluno="n"
			else	
				if co_declaracao="swd200" or co_declaracao="swd201" or co_declaracao="swd202" or co_declaracao="swd203"  or co_declaracao="swd204"then	
					Set RStabela = Server.CreateObject("ADODB.Recordset")
					SQLtabela = "SELECT * FROM TB_Da_Aula where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'" 
					RStabela.Open SQLtabela, CON2
			
					if 	RStabela.EOF then
							response.Write("ERRO 1 - N&atilde;o cadastrado TP_Nota em TB_Da_Aula para NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'" )
							response.end()
					else				
						tb_nota=RStabela("TP_Nota")		
							if tb_nota ="TB_NOTA_A" then
								caminho_nota = CAMINHO_na
							elseif tb_nota="TB_NOTA_B" then
								caminho_nota = CAMINHO_nb
							elseif tb_nota ="TB_NOTA_C" then
								caminho_nota = CAMINHO_nc
							elseif tb_nota ="TB_NOTA_D" then
								caminho_nota = CAMINHO_nd
							else
								response.Write("ERRO 2 - N&atilde;o cadastrado TP_Nota em TB_Da_Aula para NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"'" )
								response.end()
							end if	
					end if
				
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
'response.Write(unidade&","& curso&","& co_etapa&","& turma&","& cod_cons&","& vetor_materia&","& caminho_nota&","& tb_nota&","&"<BR>")
					resultados=Calc_Med_An_Fin(unidade, curso, co_etapa, turma, cod_cons, vetor_materia, caminho_nota, tb_nota, 4, 4, 0, "final", 0)		

'response.Write(resultados&"<BR>")
					resultados_apurados = split(resultados, "#%#" )	
					if ubound(resultados_apurados)=-1 then
						resultado_final_aluno="nulo"
					else				
						resultado_final_aluno=apura_resultado_aluno(curso, co_etapa, resultados_apurados(0))
'response.Write(resultado_final_aluno&"<BR>")
					end if
					if (resultado_final_aluno="Apr" or resultado_final_aluno="APR") and (co_declaracao="swd200" or co_declaracao="swd201" or co_declaracao="swd204")then
					elseif (resultado_final_aluno="Rep" or resultado_final_aluno="REP") and (co_declaracao="swd202" or co_declaracao="swd203") then
					
				'	elseif (resultado_final_aluno="Rep" or resultado_final_aluno="REP") and (co_declaracao="swd204") and (curso=2 and etapa=3) then
					elseif resultado_final_aluno="&nbsp;" and (co_declaracao="swd200" or co_declaracao="swd201" or co_declaracao="swd202" or co_declaracao="swd203" or co_declaracao="swd204") then
						gera_relatorio_aluno="n"	
					
						gera_relatorio_aluno="n"										
					else
						gera_relatorio_aluno="n"
					end if
				end if	
'response.Write(gera_relatorio_aluno&"<BR>")

				
				if gera_relatorio_aluno="n" then
				else
					Set RS = Server.CreateObject("ADODB.Recordset")
					SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod_cons
					RS.Open SQL, CON1
					
					nome_aluno = RS("NO_Aluno")
					sexo_aluno = RS("IN_Sexo")
					nome_pai = RS("NO_Pai")
					nome_mae = RS("NO_Mae")
					
					Set RSdn = Server.CreateObject("ADODB.Recordset")
					SQLdn = "SELECT * FROM TB_Contatos WHERE CO_Matricula ="& cod_cons&" AND TP_Contato='ALUNO'"
					RSdn.Open SQLdn, CONCONT	
					dt_nascimento=RSdn("DA_Nascimento_Contato")
					
					if nome_aluno="" or isnull(nome_aluno)  then
						nome_aluno="<i>N&atilde;o informado no cadastro</i>"
					else
						nome_aluno =replace_latin_char(nome_aluno,"html")
					end if	
					if nome_pai="" or isnull(nome_pai)  then
						nome_pai="<i>N&atilde;o informado no cadastro</i>"
					else
						nome_pai =replace_latin_char(nome_pai,"html")
					end if	
					if nome_mae="" or isnull(nome_mae)  then
						nome_mae="<i>N&atilde;o informado no cadastro</i>"
					else
						nome_mae =replace_latin_char(nome_mae,"html")
					end if	
					
					if sexo_aluno="F" then
						desinencia="a"
					else
						desinencia="o"
					end if

					call GeraNomes("PORT",unidade,curso,co_etapa,CON0)
					no_unidade = session("no_unidades")
					no_curso= session("no_grau")
					no_etapa = session("no_serie")
							
					Set RS3 = Server.CreateObject("ADODB.Recordset")
					SQL3 = "SELECT * FROM TB_Curso WHERE CO_Curso='"& curso &"'"
					RS3.Open SQL3, CON0
					
					no_abrv_curso = RS3("NO_Abreviado_Curso")
					co_concordancia_curso = RS3("CO_Conc")	
					
					Set RS4 = Server.CreateObject("ADODB.Recordset")
					SQL4 = "SELECT * FROM TB_Etapa WHERE CO_Curso='"& curso &"' AND CO_Etapa ='"& co_etapa &"'"
					RS4.Open SQL4, CON0
					
					no_etapa = RS4("NO_Etapa")
					art_conc = RS4("CO_Conc")
					
					no_unidade = unidade&" - "&no_unidade
					no_curso= art_conc&" "&no_etapa&" "&co_concordancia_curso&" "&no_curso
					'no_etapa = no_etapa&" "&co_concordancia_curso&" "&no_abrv_curso				
		
					SET Page = Doc.Pages.Add( 595, 842 )
							
		'CABEÇALHO==========================================================================================		
					Set Param_Logo_Gde = Pdf.CreateParam

					largura_logo_gde=formatnumber(Logo.Width*0.4,0)
altura_logo_gde=formatnumber(Logo.Height*0.4,0)
					margem=30	
					area_utilizavel=Page.Width-(margem*2)
					Param_Logo_Gde("x") = margem
					Param_Logo_Gde("y") = Page.Height - altura_logo_gde -22
					Param_Logo_Gde("ScaleX") = 0.4
Param_Logo_Gde("ScaleY") = 0.4
					Page.Canvas.DrawImage Logo, Param_Logo_Gde
			
					x_texto=largura_logo_gde+ margem
					y_texto=formatnumber(Page.Height - altura_logo_gde/2,0)
					width_texto=area_utilizavel -largura_logo_gde
					
					if cabec_1="" or isnull(cabec_1)  then
						cabec_1_tx=""
					else
						cabec_1_tx="<font style=""font-size:14pt;"">"&cabec_1&"</font><br>"
					end if
					if cabec_2="" or isnull(cabec_2)  then
						cabec_2_tx=""
					else
						cabec_2_tx="<font style=""font-size:12pt;""><b>"&cabec_2&"</b></font><br>"
					end if
					if cabec_3="" or isnull(cabec_3)  then
						cabec_3_tx=""
					else
						cabec_3_tx="<font style=""font-size:10pt;"">"&cabec_3&"</font><br>"
					end if
					if cabec_4="" or isnull(cabec_4) then
						cabec_4_tx=""
					else
						cabec_4_tx="<font style=""font-size:10pt;"">"&cabec_4&"</font><br>"
					end if
					if cabec_5="" or isnull(cabec_5) then
						cabec_5_tx=""
					else
						cabec_5_tx="<font style=""font-size:10pt;"">"&cabec_5&"</font><br>"
					end if
				
					SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; color=#000000; html=true")
					Text = "<center>"&cabec_1_tx&cabec_2_tx&cabec_3_tx&cabec_4_tx&cabec_5_tx&"</center>"
					
					
					Do While Len(Text) > 0
						CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
					 
						If CharsPrinted = Len(Text) Then Exit Do
							SET Page = Page.NextPage
						Text = Right( Text, Len(Text) - CharsPrinted)
					Loop 
					
					y_titulo=y_texto-altura_logo_gde-margem		
					
					SET Param = Pdf.CreateParam("x="&margem&";y="&y_titulo&"; height="&altura_logo_gde&"; width="&area_utilizavel&"; alignment=center; size=17; color=#000000; html=true")
					Text = "<center><b><U>"&nome_documento&"</U></b></center>"
					
					
					Do While Len(Text) > 0
						CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
					 
						If CharsPrinted = Len(Text) Then Exit Do
							SET Page = Page.NextPage
						Text = Right( Text, Len(Text) - CharsPrinted)
					Loop 			
		
					declaramos="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Declaramos para os devidos fins que, "&desinencia&" alun"&desinencia&" <b>"&nome_aluno&"</b>, filh"&desinencia&" de "&nome_pai&" e de "&nome_mae&", nascid"&desinencia&" no dia "&dt_nascimento&", "
					declaramos_ainda="<Br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Declaramos, ainda, que os documentos de transfer&ecirc;ncia ser&atilde;o expedidos no prazo de 30 dias."
					
					if co_declaracao="swd200" then
						arquivo="SWD200"
						tx_declaracao=declaramos&"cursou regularmente "&no_curso&", neste Estabelecimento de Ensino, tendo sido aprovad"&desinencia&"."
					elseif co_declaracao="swd201" then
						arquivo="SWD201"
						tx_declaracao=declaramos&"cursou regularmente "&no_curso&", neste Estabelecimento de Ensino, tendo sido aprovad"&desinencia&"."&declaramos_ainda
					elseif co_declaracao="swd202" then
						arquivo="SWD202"
						tx_declaracao=declaramos&"cursou regularmente "&no_curso&", neste Estabelecimento de Ensino, tendo sido reprovad"&desinencia&"."
					elseif co_declaracao="swd203" then
						arquivo="SWD203"
						tx_declaracao=declaramos&"cursou regularmente "&no_curso&", neste Estabelecimento de Ensino, tendo sido reprovad"&desinencia&"."&declaramos_ainda												
					elseif co_declaracao="swd204" then
						arquivo="SWD204"
						tx_declaracao="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Certificamos que <b>"&nome_aluno&"</b> nascid"&desinencia&" em "&dt_nascimento&", filh"&desinencia&" de "&nome_pai&" e de "&nome_mae&", concluiu os estudos relativos ao Ensino M&eacute;dio, estando apto a prosseguir em n&iacute;vel superior de acordo com as prerrogativas legais.<BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Informamos, ainda, que o Certificado de Conclus&atilde;o do Ensino M&eacute;dio est&aacute; em fase de expedi&ccedil;&atilde;o e ser&aacute; entregue oportunamente ao estudante."
					elseif co_declaracao="swd205" then
						arquivo="SWD205"
						tx_declaracao=declaramos&"frequentou  regularmente "&no_curso&", neste Estabelecimento de Ensino, para o ano letivo de "&ano_letivo&"."				
					elseif co_declaracao="swd206" then
						arquivo="SWD206"
						tx_declaracao=declaramos&"est&aacute; matriculad"&desinencia&" n"&no_curso&" para o ano letivo de "&ano_letivo&"."
					elseif co_declaracao="swd207" then
						arquivo="SWD207"
						tx_declaracao=declaramos&"est&aacute; matriculad"&desinencia&" e frequentando regularmente "&no_curso&" no ano letivo de "&ano_letivo&"."
		
					elseif co_declaracao="swd208" then
						arquivo="SWD208"
						tx_declaracao="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Declaramos,  para fazer prova junto ao Servi&ccedil;o Militar, que "&desinencia&" alun"&desinencia&" <b>"&nome_aluno&"</b>, filh"&desinencia&" de "&nome_pai&" e de "&nome_mae&", nascid"&desinencia&" no dia "&dt_nascimento&", cursa regularmente "&no_curso&", neste Estabelecimento de Ensino, no ano letivo de "&ano_letivo&"."
		
					elseif co_declaracao="swd209" then
						arquivo="SWD209"
						tx_declaracao=declaramos&"est&aacute; cursando regularmente "&no_curso&", neste Estabelecimento de Ensino, podendo concluir em dezembro de "&ano_letivo&"."			
					end if	
					
					x_declaracao= margem*3
					y_declaracao=y_titulo - (margem*5)
					width_declaracao=Page.Width - (margem*6)	
					
					SET Param_Declaracao = Pdf.CreateParam("x="&x_declaracao&";y="&y_declaracao&"; height=300; width="&width_declaracao&"; alignment=left; html=True")
					declaracao = "<div align=""justify""><font style=""font-size:15pt;"">"&tx_declaracao&"</font></div>"
					
					Do While Len(declaracao) > 0
						CharsPrinted = Page.Canvas.DrawText(declaracao, Param_Declaracao, Font )
					 
						If CharsPrinted = Len(declaracao) Then Exit Do
							SET Page = Page.NextPage
						declaracao = Right( declaracao, Len(declaracao) - CharsPrinted)
					Loop 
		
		
					x_data_extenso= margem*5
					y_data_extenso=y_declaracao - (margem*7)
					width_data_extenso=Page.Width - (margem*6)	
					
					SET Param_Data_Extenso = Pdf.CreateParam("x="&x_data_extenso&";y="&y_data_extenso&"; height=50; width="&width_data_extenso&"; size=13; alignment=Left; html=true")
					
					
					Do While Len(data_extenso) > 0
						CharsPrinted = Page.Canvas.DrawText(data_extenso, Param_Data_Extenso, Font )
					 
						If CharsPrinted = Len(data_extenso) Then Exit Do
							SET Page = Page.NextPage
						data_extenso = Right( data_extenso, Len(data_extenso) - CharsPrinted)
					Loop 			
					
					y_assinatura=Y_declaracao-300
					Page.Canvas.SetParams "LineWidth=0.5" 
					Page.Canvas.SetParams "LineCap=0" 
					With Page.Canvas
					   .MoveTo margem*5, y_assinatura
					   .LineTo Page.Width - (margem*5), y_assinatura
					   .Stroke
					End With 			
			
					 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width=50; alignment=left; size=8; color=#000000")
				'	Relatorio = "Sistema Web Diretor - SWD025"
					Relatorio = arquivo
					Do While Len(Relatorio) > 0
						CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )
					 
						If CharsPrinted = Len(Relatorio) Then Exit Do
						   SET Page = Page.NextPage
						Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
					Loop 
				END IF 		
			End IF
		End if
		if gera_relatorio_aluno="s" then
			relatorios_gerados=relatorios_gerados+1
		else
			relatorios_gerados=relatorios_gerados
		end if
	Next	
response.Write(relatorios_gerados)	
'response.End()
	if relatorios_gerados>-1 then					
		Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
	else
		response.redirect(origem&"index.asp?nvg="&nvg&"&opt=err4")	
	end if
end if
%>

