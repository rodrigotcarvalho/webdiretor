<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 30 'valor em segundos
'PAUTA
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/parametros.asp"-->
<!--#include file="../inc/bd_pauta.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/funcoes7.asp"-->
<!--#include file="../inc/utils.asp"-->
<% 
colunas_de_notas=48
response.Charset="ISO-8859-1"
opt= request.QueryString("opt")
ori= request.QueryString("ori")
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
nvg=session("nvg")
session("nvg")=nvg

arquivo="SWD300"
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

Paginacao = 0

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
	SQLG = "SELECT CO_Materia_Principal, CO_Professor  FROM TB_Da_Aula where CO_Professor is not null AND CO_Turma  = '"& turma &"' and  CO_Etapa = '"&co_etapa &"' AND NU_Unidade = "&unidade&" and CO_Curso = '"&curso&"'"
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

	tb_nota=tabela_nota(ano_letivo,unidade,curso,co_etapa,turma,"tb",0)	
	bancoPauta = escolheBancoPauta(tb_nota,"P",p_outro)
	caminhoBancoPauta = verificaCaminhoBancoPauta(bancoPauta,"P",p_outro)
	
	Set CONPauta = Server.CreateObject("ADODB.Connection") 
	ABRIRPauta = "DBQ="& caminhoBancoPauta & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONPauta.Open ABRIRPauta
	
	Set RSP2 = Server.CreateObject("ADODB.Recordset")
	SQL2 = "Select NU_Dia_Previsto from TB_Pauta WHERE CO_Professor  = "& co_prof &" AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Unidade  = "& unidade &" AND CO_Curso  = '"& curso &"' AND CO_Etapa  = '"& co_etapa &"' AND CO_Turma  = '"& turma &"' AND NU_Periodo = "& periodo	
	Set RSP2 = CONPauta.Execute(SQL2)	
	
	if RSP2.eof then			
		qtdPrevistas = ""
	else
		qtdPrevistas = RSP2("NU_Dia_Previsto")	
	end if				
				
	Set RSA = Server.CreateObject("ADODB.Recordset")
	SQL = "Select TB_Pauta_Aula.DT_Aula from TB_Pauta INNER JOIN TB_Pauta_Aula on TB_Pauta.NU_Pauta=TB_Pauta_Aula.NU_Pauta WHERE CO_Professor  = "& co_prof &" AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Unidade  = "& unidade &" AND CO_Curso  = '"& curso &"' AND CO_Etapa  = '"& co_etapa &"' AND CO_Turma  = '"& turma &"' AND NU_Periodo = "& periodo&" GROUP BY TB_Pauta_Aula.DT_Aula ORDER BY TB_Pauta_Aula.DT_Aula "		
'	response.write(SQL )
'response.end()
	Set RSA = CONPauta.Execute(SQL)
	

	
	vetor_aulas="" 
	bancoPauta = escolheBancoPauta(tb_nota,"P",p_outro)
	caminhoBancoPauta = verificaCaminhoBancoPauta(bancoPauta,"P",p_outro)
	
	Set CON_N = Server.CreateObject("ADODB.Connection")
	ABRIR3 = "DBQ="& caminhoBancoPauta & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_N.Open ABRIR3
	
	aulas_check = 0
	
	While Not RSA.EOF
		DT_Aula = RSA("DT_Aula")	

		vetor_data = split(DT_Aula,"/")
		data_consulta = vetor_data(1)&"/"&vetor_data(0)&"/"&vetor_data(2)
		seq_pauta = buscaSeqDataPauta(caminhoBancoPauta, data_consulta, co_prof, unidade,curso,co_etapa,turma, mat_princ, co_materia, periodo, outro)		
		vetorSeqAula=buscaSeqAula(caminhoBancoPauta,seq_pauta, data_consulta, outro)	
		
		vetor_data = split(DT_Aula,"/")
		mes = vetor_data(1)
		mes = mes*1
		if mes_check = "" then
			mes_check =mes	
		end if	
		if mes<>mes_check then
			if total_meses = 0 then
				vetor_mes_colspan=colspan_mes			
			else
				vetor_mes_colspan=vetor_mes_colspan&"#$#"&colspan_mes			
			end if			
			mes_check = mes
			colspan_mes = 0
			total_meses = total_meses+1
		end if
		vetor_seq_data = split(vetorSeqAula,"#!#")	
		data_colspan = UBOUND(vetor_seq_data)+1
			
	
		colspan_mes = colspan_mes+(1*data_colspan)
		data_vetor=formata(DT_Aula,"DD/MM/YYYY")
		vetor_data = split(data_vetor,"/")	
		nu_chamada_check=nu_chamada_check*1
		if aulas_check = 0 then
			vetor_aulas = vetor_data(0)
			vetor_mes = mes	
			vetor_seq_pauta=seq_pauta
			vetor_nu_aulas = vetorSeqAula
			vetor_data_colspan = data_colspan
		else
			vetor_aulas=vetor_aulas&"#$#"&vetor_data(0)
			vetor_mes = vetor_mes&"#$#"&mes		
			vetor_seq_pauta=vetor_seq_pauta&"#$#"&seq_pauta
			vetor_nu_aulas = vetor_nu_aulas&"#$#"&vetorSeqAula
			vetor_data_colspan = vetor_data_colspan&"#$#"&data_colspan		
		end if
		aulas_check=aulas_check+1		
	RSA.MoveNext
	Wend 
	if total_meses = 0 then
		vetor_mes_colspan=colspan_mes			
	else
		vetor_mes_colspan=vetor_mes_colspan&"#$#"&colspan_mes			
	end if		
	
'	response.Write(aulas_check&"---<BR>")
'	response.Write(vetor_mes_colspan&"<BR>")
'	response.Write(vetor_mes&"<BR>")
'	response.Write(vetor_aulas&"<BR>")
'	response.Write(vetor_nu_aulas&"<BR>")
'	response.Write(vetor_data_colspan&"<BR>")
'	response.Write(vetor_seq_pauta&"<BR>")	
	if aulas_check>0 then
		mes = split(vetor_mes,"#$#")
		mes_colspan = split(vetor_mes_colspan,"#$#")
		aulas_encontrados = split(vetor_aulas, "#$#" )	
		nu_aulas_encontradas = split(vetor_nu_aulas, "#$#" )		
		aula_colspan_encontradas = split(vetor_data_colspan, "#$#" )
		seq_pauta = split(vetor_seq_pauta, "#$#" )	
		total_aulas=0
		conta_tempos=0
		colunas_preenchidas=0
		vetor_mes = ""
		vetor_mes_colspan=""
		vetor_nu_aulas = ""
		vetor_data_colspan=""
		vetor_seq_pauta = ""
		qtd_pag = 1
		mes_check=""
		mudou_pag="N"	
		mudou_mes="N"	

		
		aulas_dadas=0
		for t =0 to ubound(aula_colspan_encontradas)
		
			total_aulas=total_aulas+aula_colspan_encontradas(t)
			colunas_preenchidas=colunas_preenchidas*1
			aula_colspan_encontradas(t)	=aula_colspan_encontradas(t)*1
			conta_tempos_temp = conta_tempos+aula_colspan_encontradas(t)	
			conta_tempos_temp=conta_tempos_temp*1
			colunas_de_notas=colunas_de_notas*1
			aulas_dadas = aulas_dadas+aula_colspan_encontradas(t)
			if t=0 then
					mes_check=mes(t)
					colunas_preenchidas	= colunas_preenchidas+aula_colspan_encontradas(t)	
					vetor_mes = mes(t)
					vetor_aulas = aulas_encontrados(t)			
					vetor_nu_aulas = nu_aulas_encontradas(t)		
					vetor_data_colspan =  aula_colspan_encontradas(t)	
					vetor_seq_pauta = seq_pauta(t)	
					conta_tempos = aula_colspan_encontradas(t)	
			end if		
			mes_check=mes_check*1
			mes(t)=mes(t)*1	
		
			if mes(t)<>mes_check then
				mes_check=mes(t)
				if vetor_mes_colspan="" then
					vetor_mes_colspan = colunas_preenchidas			
				elseif mudou_pag="S" then
					vetor_mes_colspan = vetor_mes_colspan&colunas_preenchidas	
					mudou_pag="N"		
				else
					vetor_mes_colspan = vetor_mes_colspan&"#$#"&colunas_preenchidas	
				end if	
				colunas_preenchidas=0
				mudou_mes="S"
		'	response.Write(	mes(t)&"<>"&mes_check &"<BR>")		
		'	response.Write(vetor_mes_colspan&" ok <BR>")		
			end if	
		
			if conta_tempos_temp>colunas_de_notas then
				conta_tempos = aula_colspan_encontradas(t)		
				vetor_mes_colspan = vetor_mes_colspan&"#$#"&colunas_preenchidas&"$$$"	
				vetor_mes = vetor_mes&"$$$"&mes(t)				
				vetor_aulas = vetor_aulas&"$$$"&aulas_encontrados(t)	
				vetor_nu_aulas = vetor_nu_aulas&"$$$"&nu_aulas_encontradas(t)			
				vetor_data_colspan =  vetor_data_colspan&"$$$"&aula_colspan_encontradas(t)
				vetor_seq_pauta = vetor_seq_pauta&"$$$"&seq_pauta(t)
		
				qtd_pag=qtd_pag+1	
				colunas_preenchidas	= aula_colspan_encontradas(t)	
				mudou_pag="S"		
			elseif t>0 then
				colunas_preenchidas	= colunas_preenchidas+aula_colspan_encontradas(t)
				vetor_mes = vetor_mes&"#$#"&mes(t)			
				vetor_aulas = vetor_aulas&"#$#"&aulas_encontrados(t)	
				vetor_nu_aulas = vetor_nu_aulas&"#$#"&nu_aulas_encontradas(t)					
				vetor_data_colspan =  vetor_data_colspan&"#$#"&aula_colspan_encontradas(t)	
				vetor_seq_pauta = vetor_seq_pauta&"#$#"&seq_pauta(t)	
				conta_tempos = conta_tempos+aula_colspan_encontradas(t)					
			end if		
		
		
		
		Next

		if mudou_mes = "N" then
			vetor_mes_colspan = colunas_preenchidas	
		elseif qtd_pag = 1 or mudou_pag="N" then
			vetor_mes_colspan = vetor_mes_colspan&"#$#"&colunas_preenchidas	
		else
			vetor_mes_colspan = vetor_mes_colspan&"$$$"&colunas_preenchidas	
		end if
		
		total_meses = 0 
		vetor_pags_data = split(vetor_mes,"$$$")
		for pg =0 to ubound(vetor_pags_data)
			meses_pag = 0
			mes_check = "" 
			vetor_mes_pag = split(vetor_pags_data(pg),"#$#")
			for ms =0 to ubound(vetor_mes_pag)
				mes = vetor_mes_pag(ms)		
				nome_mes = GeraNomesNovaVersao("MES_ABR",mes,variavel2,variavel3,variavel4,variavel5,conexao,outro)	
				if mes_check <> "" then
					mes_check=mes_check*1
				end if	
				mes = mes*1
		
						
				if mes<>mes_check then
					if total_meses = 0 then
						mes_check =mes				
						vetor_meses=nome_mes
					elseif meses_pag = 0 then
						vetor_meses=vetor_meses&nome_mes
					else		
						vetor_meses=vetor_meses&"#$#"&nome_mes
					end if			
					mes_check = mes
					meses_pag=meses_pag+1
					total_meses = total_meses+1
				end if	
				
			Next	
		
			if total_meses = 0 then
				vetor_meses=nome_mes
			else
				if mes<>mes_check then
					vetor_meses=vetor_meses&"#$#"&nome_mes
				end if		
			end if	
			
			mes_check = mes		
			meses_pag=meses_pag+1		
			total_meses = total_meses+1		
			if pg<ubound(vetor_pags_data) then
				vetor_meses=vetor_meses&"$$$"
			end if
		Next
'		response.Write("<BR>")
'		response.Write(vetor_meses&"<BR>")
'		response.Write(vetor_mes_colspan&"<BR>")
'		response.Write(vetor_mes&"<BR>")
'		response.Write(vetor_aulas&"<BR>")
'		response.Write(vetor_nu_aulas&"<BR>")
'		response.Write(vetor_data_colspan&"<BR>")
'		response.Write(vetor_seq_pauta&"<BR>")
'		
'		
'		response.end()
		
		'if total_aulas<colunas_de_notas-3 then
		'	
		'else
		'	qtd_pag = 2
		'end if	
		
		
		
		
		
		'response.Write(qtd_aulas&" "&colunas_de_notas)
		'response.End()
		nu_chamada_check = 1	
		
		Set RSA = Server.CreateObject("ADODB.Recordset")
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
		
		'if curso=0 then
		'	gera_pdf="nao"
		'else
		
		
		
		'		ln_pesos_cols=verifica_dados_tabela(opcao,"peso_col",outro)
		'		ln_pesos_vars=verifica_dados_tabela(opcao,"peso_bd_var",outro)
		'		nm_pesos_vars=verifica_dados_tabela(opcao,"peso_wrk_var",outro)
				ln_nom_cols="Matr&iacute;cula#!#Nome#!#"
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
		
				
				gera_pdf="sim"
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
		 nu_pag=1
		pg_mes = split(vetor_meses,"$$$")
		pg_mes_colspan = split(vetor_mes_colspan,"$$$")
		pg_aulas_encontrados = split(vetor_aulas, "$$$" )	
		pg_nu_aulas_encontradas = split(vetor_nu_aulas, "$$$" )		
		pg_aula_colspan_encontradas = split(vetor_data_colspan, "$$$" )
		pg_seq_pauta = split(vetor_seq_pauta, "$$$" )	
		
	
		
		for pg =0 to ubound(pg_aula_colspan_encontradas)
		
			meses_encontrados = split(pg_mes(pg),"#$#")
			mes_colspan_encontrados = split(pg_mes_colspan(pg),"#$#")
			aulas_encontrados = split(pg_aulas_encontrados(pg), "#$#" )	
			nu_aulas_encontradas = split(pg_nu_aulas_encontradas(pg), "#$#" )		
			aula_colspan_encontradas = split(pg_aula_colspan_encontradas(pg), "#$#" )
			seq_pauta = split(pg_seq_pauta(pg), "#$#" )	
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
		
				y_texto=y_texto-altura_logo_gde+10
				SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height=30; width=500; alignment=center; size=14; color=#000000; html=true")
				Text = "<center><i><b><font style=""font-size:18pt;"">Di&aacute;rio de Classe - Pauta</font></b></i></center>"
						
				Do While Len(Text) > 0
					CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
				 
					If CharsPrinted = Len(Text) Then Exit Do
						SET Page = Page.NextPage
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
		
				total_de_colunas=colunas_de_notas+2			
				altura_medias=20
				y_segunda_tabela=y_primeira_tabela-20	
				Set param_table2 = Pdf.CreateParam("width="&area_utilizavel&"; height="&altura_medias&"; rows=3; cols="&total_de_colunas&"; border=1; cellborder=0.1; cellspacing=0; x="&margem&"; y="&y_segunda_tabela&"; MaxHeight=420")
		
				Set Notas_Tit = Doc.CreateTable(param_table2)
				Notas_Tit.Font = Font				
				largura_colunas=(area_utilizavel-40-150)/colunas_de_notas		
				With Notas_Tit.Rows(1)	
				   .Cells(1).Height = 10						
				   .Cells(1).Width = 40		 
				   .Cells(2).Width = 150				               
					for d=3 to total_de_colunas
					 .Cells(d).Width = largura_colunas					
					next
				End With
				
				With Notas_Tit.Rows(2)	
				   .Cells(1).Height = 10								
				End With		
				
				With Notas_Tit.Rows(3)	
				   .Cells(1).Height = 10								
				End With		
						
				alunos_encontrados = split(vetor_matriculas, "#$#" )	
				vetor_total_faltas=""
				if pg =0 then
					for b=0 to ubound(alunos_encontrados)	
						qtd_faltas_aluno=0
						dados_alunos = split(alunos_encontrados(b), "#!#" )					
						for a=0 to ubound(aulas_encontrados)
							nu_aula = split(nu_aulas_encontradas(a), "#!#" )			
							for n=0 to ubound(nu_aula)				
								Set RS3 = Server.CreateObject("ADODB.Recordset")
								SQL_N = "Select * from TB_Pauta_Faltas WHERE CO_Matricula = "& dados_alunos(0) & " AND NU_Pauta = "& seq_pauta(a) &" AND NU_Seq = "& nu_aula(n)
								Set RS3 = CON_N.Execute(SQL_N)			  
										
								if NOT RS3.EOF then 
									qtd_faltas_aluno=qtd_faltas_aluno+1				
								end if											
							next	
						next
						if b=0 then
							vetor_total_faltas=qtd_faltas_aluno												
						else
							vetor_total_faltas = vetor_total_faltas&"#!#"&qtd_faltas_aluno
						end if	
						'response.Write(vetor_total_faltas&"<BR>")						
					next		
					total_faltas = split(vetor_total_faltas,"#!#")
				end if
		
				linha=1
				fim_do_cabecalho=3	
				tabela_col=1
				Notas_Tit(1, 1).RowSpan = 3		
				Notas_Tit(1, 2).RowSpan = 3		
				if nu_pag=qtd_pag then
					Notas_Tit(1, colunas_de_notas).RowSpan = 3		
					Notas_Tit(1, colunas_de_notas).ColSpan = 3	
					Notas_Tit(1, colunas_de_notas).AddText "<div align=""center""><b>Total de<br>Faltas</b></div>", "size=7; indenty=5; alignment=center; html=true", Font							
				end if	
				for e=0 to ubound(linha_nome_colunas)
					colunas_de_notas=colunas_de_notas*1
					tabela_col=tabela_col*1
					nu_pag=nu_pag*1
					qtd_pag=qtd_pag*1
					linha=linha*1
					'response.Write(colunas_de_notas&"="&tabela_col&"<BR>")
		
						Notas_Tit(linha, tabela_col).AddText "<div align=""center""><b>"&linha_nome_colunas(e)&"</b></div>", "size=7; indenty=10; alignment=center; html=true", Font				
					tabela_col=tabela_col+1
		
				next		
		
					
				tabela_col=3	
				
				for m=0 to ubound(meses_encontrados)		
					if mes_colspan_encontrados(m) >1 then
						Notas_Tit(1, tabela_col).ColSpan = mes_colspan_encontrados(m)			
					end if	
					Notas_Tit(1, tabela_col).AddText "<div align=""center""><b>"&meses_encontrados(m)&"</b></div>", "size=7; indenty=0; alignment=center; html=true", Font				
					tabela_col=tabela_col+mes_colspan_encontrados(m)	
				next	
				
		'		Notas_Tit(1, 48).RowSpan = 3					
		'		Notas_Tit(1, 48).ColSpan = 3	
		'		Notas_Tit(1, 48).AddText "<div align=""center""><b>Total<br>de<br>Aulas</b></div>", "size=7; indenty=0; alignment=center; html=true", Font			
						
							
				tabela_col=3		
				for a=0 to ubound(aulas_encontrados)
					if aula_colspan_encontradas(a)	>1 then
						Notas_Tit(2, tabela_col).ColSpan = aula_colspan_encontradas(a)			
					end if
					Notas_Tit(2, tabela_col).AddText "<div align=""center""><b>"&aulas_encontrados(a)&"</b></div>", "size=7; indenty=0; alignment=center; html=true", Font
					nu_aula = split(nu_aulas_encontradas(a), "#!#" )			
					for n=0 to ubound(nu_aula)	
						tempo_aula = buscaTempoAula(caminhoBancoPauta, seq_pauta(a), nu_aula(n), outro)				
						Notas_Tit(3, tabela_col).AddText "<div align=""center""><b>"&tempo_aula&"</b></div>", "size=7; indenty=0; alignment=center; html=true", Font										
						tabela_col=tabela_col+1
					next	
				next			
				
			
				Set param_materias = PDF.CreateParam	
				param_materias.Set "size=6;expand=true" 			
														
				
				conta_notas = 1 
				linha=3			
				nu_chamada_ckq = 0
				
				for b=0 to ubound(alunos_encontrados)	
					param_materias.Add "indenty=2;alignment=right;html=true"
					param_materias.Add "indentx=5"	
					dados_alunos = split(alunos_encontrados(b), "#!#" )		 
					linha=linha+1
					Set Row = Notas_Tit.Rows.Add(15) ' row height							
					Notas_Tit(linha, 1).AddText dados_alunos(0), param_materias	
					Notas_Tit(linha, 2).AddText dados_alunos(2), param_materias	
					coluna=2
		
					for a=0 to ubound(aulas_encontrados)
						nu_aula = split(nu_aulas_encontradas(a), "#!#" )			
						for n=0 to ubound(nu_aula)				
							Set RS3 = Server.CreateObject("ADODB.Recordset")
							SQL_N = "Select * from TB_Pauta_Faltas WHERE CO_Matricula = "& dados_alunos(0) & " AND NU_Pauta = "& seq_pauta(a) &" AND NU_Seq = "& nu_aula(n)
							Set RS3 = CON_N.Execute(SQL_N)			  
							
								align="center"
								
								if RS3.EOF then 
									conteudo="<center>&bull;</center>"
									classe = "texto_azul"
								else															
									conteudo="<center>F</center>"																		
									classe = "texto_vermelho"
								end if						
							coluna=coluna+1
							param_materias.Add "indenty=2;alignment=center;html=true"
							param_materias.Add "indentx=0"	
							Notas_Tit(linha, coluna).AddText conteudo, param_materias												
						next	
					next
					param_materias.Add "indenty=2;alignment=center;html=false"						
					if nu_pag=qtd_pag then	
						Notas_Tit(linha, total_de_colunas-2).ColSpan = 3	
						Notas_Tit(linha, total_de_colunas-2).AddText total_faltas(b), param_materias														
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
						Loop 				   ' Display remaining part of table on the next page
						Set Page = Page.NextPage	
						param_table2.Add( "RowTo="&fim_do_cabecalho&"; RowFrom=1" ) ' Row 1 is header.
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
		
				y_texto=y_texto-altura_logo_gde+10
				SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height=30; width=500; alignment=center; size=14; color=#000000; html=true")
				Text = "<center><i><b><font style=""font-size:18pt;"">Di&aacute;rio de Classe - Pauta</font></b></i></center>"
						
				Do While Len(Text) > 0
					CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
				 
					If CharsPrinted = Len(Text) Then Exit Do
						SET Page = Page.NextPage
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
				
				if nu_pag=qtd_pag then	
				
					SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem*2&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
					
					Relatorio = "Total de Aulas Previstas: "&qtdPrevistas
					Do While Len(Relatorio) > 0
						CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )			
						If CharsPrinted = Len(Relatorio) Then Exit Do
						SET Page = Page.NextPage
						Relatorio = Right( Relatorio, Len(Relatorio) - CharsPrinted)
					Loop 
					
					Param_Relatorio.Add "alignment=right" 
					
	Set RSPeriodo = Server.CreateObject("ADODB.Recordset")
	SQLPeriodo = "Select * from TB_Periodo WHERE NU_Periodo= "&periodo
	Set RSPeriodo = CON0.Execute(SQLPeriodo)
	
	dataFim = RSPeriodo("DA_Fim_Periodo")					
					
					Relatorio = "Encerrado em: "&formata(dataFim, "DD/MM/YYYY")
					Do While Len(Relatorio) > 0
						CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )			
						If CharsPrinted = Len(Relatorio) Then Exit Do
						SET Page = Page.NextPage
						Paginacao = Right( Relatorio, Len(Relatorio) - CharsPrinted)
					Loop 
										
					Param_Relatorio.Add "html=true" 
					
					Relatorio = "<center>Total de Aulas Dadas: "&aulas_dadas&"</center>"
					Do While Len(Relatorio) > 0
						CharsPrinted = Page.Canvas.DrawText(Relatorio, Param_Relatorio, Font )			
						If CharsPrinted = Len(Relatorio) Then Exit Do
						SET Page = Page.NextPage
						data_hora = Right( Relatorio, Len(Relatorio) - CharsPrinted)
					Loop 			
				END IF
				
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
		nu_pag=nu_pag+1
		Next
	end if			

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
			
	End IF					
'End IF							
Next
	

Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>

