<!--#include file="caminhos.asp"-->
<!--#include file="funcoes2.asp"-->
<!--#include file="funcoes6.asp"-->
<%

Set CONBL = Server.CreateObject("ADODB.Connection") 
ABRIRBL = "DBQ="& CAMINHO_bl & ";Driver={Microsoft Access Driver (*.mdb)}"
CONBL.Open ABRIRBL


function ehIsentoRematricula(p_cod_cons, p_ano_letivo)
  
  vencimento=vencimentoRematricula(p_ano_letivo)
  
vetor_vencimento = split(vencimento, "/")
	vetor_vencimento(0)=vetor_vencimento(0)*1
	if vetor_vencimento(0)<10 then
		dia_vencimento="0"&vetor_vencimento(0)
	else
		dia_vencimento=vetor_vencimento(0)					
	end if
	
	vetor_vencimento(1)=vetor_vencimento(1)*1	
	if vetor_vencimento(1)<10 then
		mes_vencimento="0"&vetor_vencimento(1)
	else
		mes_vencimento=vetor_vencimento(1)						
	end if
		
	vencimento = dia_vencimento&"/"&mes_vencimento&"/"&vetor_vencimento(2)				
	vencimento_inicial = vetor_vencimento(1)&"/"&vetor_vencimento(0) &"/"&vetor_vencimento(2) 				
	if ((((vetor_vencimento(1) = 1 or vetor_vencimento(1) = 3 or vetor_vencimento(1) = 5 or vetor_vencimento(1) = 7 or vetor_vencimento(1) = 8 or vetor_vencimento(1) = 10 or vetor_vencimento(1) = 12) and vetor_vencimento(0) = 31)   or   (vetor_vencimento(1) = 4 or vetor_vencimento(1) = 6 or vetor_vencimento(1) = 9 or vetor_vencimento(1) = 11) and vetor_vencimento(0) = 30)) then
		dia_vencimento = 1
		mes_vencimento = vetor_vencimento(1)+1
	elseif ((vetor_vencimento(1) = 2 and (vetor_vencimento(2) MOD 4 = 0) and  vetor_vencimento(0) = 29) or (vetor_vencimento(1) = 2 and  vetor_vencimento(0) = 28)) then
		dia_vencimento = 1
		mes_vencimento = vetor_vencimento(1)+1				
	else
		dia_vencimento = vetor_vencimento(0)+1
		mes_vencimento = vetor_vencimento(1)
	end if	
	if ((vetor_vencimento(1) = 12) and vetor_vencimento(0) = 31) then
		ano_vencimento = vetor_vencimento(2)+1			
	else
		ano_vencimento = vetor_vencimento(2)				
	end if 
	vencimento_fim = mes_vencimento&"/"&dia_vencimento&"/"&ano_vencimento   
  
	Set RS1 = Server.CreateObject("ADODB.Recordset")
	SQL1 = "SELECT * FROM TB_Bloqueto WHERE DA_Vencimento>=#"& vencimento_inicial &"# and DA_Vencimento<#"& vencimento_fim &"#   AND CO_Matricula_Escola ="& p_cod_cons
	RS1.Open SQL1, CONBL
	
	if RS1.EOF then
		ehIsentoRematricula = "S"
	else
		ehIsentoRematricula = "N"	
	end if	  

end function






function GeraBloqueto(P_DADOS, P_MATRICULA, P_MES, P_SALVAR)
response.Charset="ISO-8859-1"
dados = P_DADOS
cod_cons = P_MATRICULA
mes=P_MES
salvar =P_SALVAR
complemento ="_"&cod_cons&mes&datepart("y",now)&datepart("h",now)&datepart("n",now)&datepart("s",now)
If Not IsArray(vetor_meses) Then 
	vetor_meses = Array()
End if
If InStr(Join(vetor_meses), mes) = 0 Then
	ReDim preserve vetor_meses(UBound(vetor_meses)+1)
	vetor_meses(Ubound(vetor_meses )) = mes
End if
'vetor_meses = split(dados,", ")


ano_letivo = session("ano_letivo")

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

	'Dim AspPdf, Doc, Page, Font, Text, Param, Image, CharsPrinted
	'Instancia o objeto na memória
	if salvar = "S" then
		addPastas = "../../../"
	end if	
	
    codigo_banco = "237-2"	
	nome_banco_msg = "Banco Bradesco"
	nome_banco_tit = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	
	
	SET Pdf = Server.CreateObject("Persits.Pdf")
	SET Doc = Pdf.CreateDocument
	Set Logo = Doc.OpenImage( Server.MapPath( addPastas&"../img/logo_boleto.gif") )
	Set Itau = Doc.OpenImage( Server.MapPath( addPastas&"../img/boleto_bradesco/logobanco.gif") )	
	Set Font = Doc.Fonts.LoadFromFile(Server.MapPath(addPastas&"../fonts/arial.ttf"))	
	Set Font_Tesoura = Doc.Fonts.LoadFromFile(Server.MapPath(addPastas&"../fonts/ZapfDingbats.ttf"))
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
		
		Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4 = "DBQ="& CAMINHO_pf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4		

	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod_cons
	RS.Open SQL, CON1
	
	nome_aluno = RS("NO_Aluno")
	sexo_aluno = RS("IN_Sexo")
	
	nome_aluno=replace_latin_char(nome_aluno,"html")	
	
	if sexo_aluno="F" then
		desinencia="a"
	else
		desinencia="o"
	end if


	Set RS1 = Server.CreateObject("ADODB.Recordset")
	SQL1 = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod_cons
	RS1.Open SQL1, CON1
	
	if RS1.EOF then
		response.redirect("index.asp?nvg="&nvg&"&opt=err1")
	else
	
		ano_aluno = RS1("NU_Ano")
		rematricula = RS1("DA_Rematricula")
		situacao = RS1("CO_Situacao")
		encerramento= RS1("DA_Encerramento")
		unidade= RS1("NU_Unidade")
		curso_aluno= RS1("CO_Curso")
		etapa_aluno= RS1("CO_Etapa")
		turma_aluno= RS1("CO_Turma")
		cham= RS1("NU_Chamada")
			
		'call GeraNomes("PORT",unidade,curso,etapa,CON0)
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Unidade WHERE NU_Unidade="& unidade
		RS2.Open SQL2, CON0
		
		if RS2.EOF then
			no_unidade = ""
			co_cnpj = ""
			no_sede = ""			
		else				
			no_unidade = RS2("TX_Imp_Cabecalho")	
			co_cnpj = RS2("CO_CGC")		
			no_sede = RS2("NO_Sede")	
			rua_sede = RS2("NO_Logradouro")	
			num_sede = RS2("NU_Logradouro")	
			co_bairro_sede = RS2("CO_Bairro")	
			co_cidade_sede = RS2("CO_Municipio")					
			complemento_sede = RS2("TX_Complemento_Logradouro")		
			cep_sede = RS2("CO_CEP")	
			
			if InStr(cep_sede,"-")= 0 then
				cep_sede = "CEP: "&left(cep_sede,5)&"-"&right(cep_sede,3)
			end if
			
			if co_bairro_sede<>"" then
				Set RS2b = Server.CreateObject("ADODB.Recordset")
				SQL2b = "SELECT * FROM TB_Bairros WHERE CO_Municipio="& co_cidade_sede&" AND CO_Bairro = "&co_bairro_sede
				RS2b.Open SQL2b, CON0
				
				if not RS2b.eof then
					no_bairro  = RS2b("NO_Bairro")	
				end if
			end if
			if complemento_sede<>"" and complemento_sede<>" " and not isnull(complemento_sede) then
				complemento_sede = " - "&complemento_sede
			end if
			
			endereco_sede =rua_sede&", "&num_sede&complemento_sede&" - "& no_bairro	&" - "&cep_sede							
		end if
		'no_curso= session("no_grau")
		'no_etapa = session("no_serie")
		
		
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Turma_Mensagem_Bloqueto WHERE NU_Unidade="& unidade&" and CO_Curso='"& curso_aluno &"' AND CO_Etapa = '"&etapa_aluno&"' and CO_Turma ='"&turma_aluno&"'"		
		'SQL3 = "SELECT * FROM TB_Turma_Mensagem_Bloqueto WHERE NU_Unidade="& unidade&" and CO_Grau='"& curso_aluno &"' AND CO_Serie = '"&etapa_aluno&"' and CO_Turma ='"&turma_aluno&"'"
		RS3.Open SQL3, CON0	
		
		if RS3.EOF then
			no_escola = ""
		else				
			no_escola = RS3("NO_Remetente")	
		end if					
			
						
'		Set RS3 = Server.CreateObject("ADODB.Recordset")
'		SQL3 = "SELECT * FROM TB_Curso WHERE CO_Curso='"& curso &"'"
'		RS3.Open SQL3, CON0

						
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Turma, TB_Turno WHERE TB_Turma.CO_Turno=TB_Turno.CO_Turno and NU_Unidade="& unidade&" and CO_Curso='"& curso_aluno &"' AND CO_Etapa = '"&etapa_aluno&"' and CO_Turma ='"&turma_aluno&"'"
		RS4.Open SQL4, CON0
		if RS4.EOF then
			no_turno = ""
		else			
			no_turno = RS4("NO_Turno")
		end if	
		
		'no_abrv_curso = RS3("NO_Abreviado_Curso")
		'co_concordancia_curso = RS3("CO_Conc")	
		
		no_unidade = unidade&" - "&no_unidade
		'no_curso= no_etapa&" "&co_concordancia_curso&" "&no_curso
		'no_curso= no_curso&" - "&no_etapa
		'no_etapa = no_etapa&" "&co_concordancia_curso&" "&no_abrv_curso				
		for n=0 to ubound(vetor_meses)
			margem_x=20	
			margem_y=20		
			row_padrao=margem_y
			'if ubound(vetor_meses)mod 2 = 0 then
				'SET Page = Doc.Pages.Add( 595, 842 )
				'if ubound(vetor_meses) = n then	
				'	altura_inicial=421		
				'else
				'	altura_inicial=margem_y											
				'end if	
			'else
			'	altura_inicial=421			
			'end if	

			Set Param_Logo_Gde = Pdf.CreateParam					
			largura_logo_gde=Logo.Width 'formatnumber(Logo.Width*1,0)
			altura_logo_gde=Logo.Height 'formatnumber(Logo.Height*1,0)
			
			SET Page = Doc.Pages.Add( 595, 852 )
'CABEÇALHO==========================================================================================		
			Set Param_Logo_Gde = Pdf.CreateParam				
			largura_logo_gde=formatnumber(Logo.Width*0.7,0)
			altura_logo_gde=formatnumber(Logo.Height*0.7,0)
	
			Param_Logo_Gde("x") = margem_x
			Param_Logo_Gde("y") = Page.Height - altura_logo_gde -22
			Param_Logo_Gde("ScaleX") = 0.7
			Param_Logo_Gde("ScaleY") = 0.7
			Page.Canvas.DrawImage Logo, Param_Logo_Gde

			'x_texto=largura_logo_gde+ 30
			x_texto= margem_x
			y_texto=formatnumber(Page.Height - altura_logo_gde/2,0)
			width_texto=Page.Width - (margem*2)

		
			SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura_logo_gde&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
			Text = "<p><center><i><b><font style=""font-size:18pt;"">"&no_escola&"</font></b></i></center></p>"
			

			Do While Len(Text) > 0
				CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
			 
				If CharsPrinted = Len(Text) Then Exit Do
					SET Page = Page.NextPage
				Text = Right( Text, Len(Text) - CharsPrinted)
			Loop 
			
			vetor_cnpj=SPLIT(co_cnpj,"/")
			if ubound(vetor_cnpj)>0 then
			'response.Write(">>>>>>>>>>"&vetor_cnpj(1))			
'				if vetor_cnpj(1)<0 then
'					vetor_cnpj(1)=vetor_cnpj(1)*10
'				end if
			cnpj_formatado = vetor_cnpj(0)&"/"&vetor_cnpj(1)
			exibe_cnpj="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CNPJ: "&vetor_cnpj(0)&"/"&vetor_cnpj(1)
			end if				
			
			
			x_cnpj = margem_x
			y_cnpj = formatnumber(Page.Height - altura_logo_gde-margem_y-5,0)
			width_cnpj = largura_logo_gde+margem_x
			SET Param = Pdf.CreateParam("x="&x_cnpj&";y="&y_cnpj&"; height=20; width="&width_cnpj&"; alignment=RIGHT; size=5.5; color=#000000; html=true")		
			
			Do While Len(exibe_cnpj) > 0
				CharsPrinted = Page.Canvas.DrawText(exibe_cnpj, Param, Font )
			 
				If CharsPrinted = Len(exibe_cnpj) Then Exit Do
					SET Page = Page.NextPage
				exibe_cnpj = Right( exibe_cnpj, Len(exibe_cnpj) - CharsPrinted)
			Loop 				
			
'================================================================================================================		

			wrk_mes=vetor_meses(n)*1
			Set RS4 = Server.CreateObject("ADODB.Recordset")
			SQL4= "SELECT * FROM TB_Posicao WHERE VA_Realizado=0 AND NO_Lancamento='Mensalidade' AND CO_Matricula_Escola ="& cod_cons &" AND Mes = "&wrk_mes
			RS4.Open SQL4, CON4	
		
			if RS4.EOF then
				response.redirect("index.asp?nvg="&nvg&"&opt=err2")
			else
				vencimento=RS4("DA_Vencimento")
				nu_cota=RS4("NU_Cota")

				vetor_vencimento = split(vencimento, "/")
				vetor_vencimento(0)=vetor_vencimento(0)*1
				if vetor_vencimento(0)<10 then
					dia_vencimento="0"&vetor_vencimento(0)
				else
					dia_vencimento=vetor_vencimento(0)					
				end if
				
				vetor_vencimento(1)=vetor_vencimento(1)*1
				if vetor_vencimento(1)<10 then
					mes_vencimento="0"&vetor_vencimento(1)
				else
					mes_vencimento=vetor_vencimento(1)						
				end if
		
				vencimento = dia_vencimento&"/"&mes_vencimento&"/"&vetor_vencimento(2)				
				vencimento_inicial = vetor_vencimento(1)&"/"&vetor_vencimento(0) &"/"&vetor_vencimento(2) 				
				if ((((vetor_vencimento(1) = 1 or vetor_vencimento(1) = 3 or vetor_vencimento(1) = 5 or vetor_vencimento(1) = 7 or vetor_vencimento(1) = 8 or vetor_vencimento(1) = 10 or vetor_vencimento(1) = 12) and vetor_vencimento(0) = 31)   or   (vetor_vencimento(1) = 4 or vetor_vencimento(1) = 6 or vetor_vencimento(1) = 9 or vetor_vencimento(1) = 11) and vetor_vencimento(0) = 30)) then
					dia_vencimento = 1
					mes_vencimento = vetor_vencimento(1)+1
				elseif ((vetor_vencimento(1) = 2 and (vetor_vencimento(2) MOD 4 = 0) and  vetor_vencimento(0) = 29) or (vetor_vencimento(1) = 2 and  vetor_vencimento(0) = 28)) then
					dia_vencimento = 1
					mes_vencimento = vetor_vencimento(1)+1				
				else
					dia_vencimento = vetor_vencimento(0)+1
					mes_vencimento = vetor_vencimento(1)
				end if	
				if ((vetor_vencimento(1) = 12) and vetor_vencimento(0) = 31) then
					ano_vencimento = vetor_vencimento(2)+1			
				else
					ano_vencimento = vetor_vencimento(2)				
				end if 
				vencimento_fim = mes_vencimento&"/"&dia_vencimento&"/"&ano_vencimento 
			end if
			
			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT * FROM TB_Bloqueto WHERE DA_Vencimento>=#"& vencimento_inicial &"# and DA_Vencimento<#"& vencimento_fim &"#   AND CO_Matricula_Escola ="& cod_cons
			'response.Write(SQL1)			
			RS1.Open SQL1, CONBL
			
			if RS1.EOF then		
				vetor_venc = split(vencimento_inicial,"/")
				vencimento_inicial_alt = vetor_venc(1)&"/"&vetor_venc(0)&"/"&vetor_venc(2)
				vencimento_fim_alt = dia_vencimento&"/"&mes_vencimento&"/"&ano_vencimento 
				Set RS1 = Server.CreateObject("ADODB.Recordset")
				SQL1 = "SELECT * FROM TB_Bloqueto WHERE DA_Vencimento>=#"& vencimento_inicial_alt &"# and DA_Vencimento<#"& vencimento_fim_alt &"#   AND CO_Matricula_Escola ="& cod_cons
				response.Write(SQL1)			
				RS1.Open SQL1, CONBL
				
				if RS1.EOF then		
					nu_carne = ""	
					nosso_numero = ""
					va_inicial = ""
				else	
					nu_carne=RS1("NU_Bloqueto")
					nosso_numero = "09/"&RS1("CO_Nosso_Numero")
					va_inicial = RS1("VA_Inicial")
					cod_superior=RS1("CO_Superior")				
					cod_barras =RS1("CO_Barras")
					turma =RS1("CO_Turma")
					no_cedente =RS1("NO_Cedente")
					co_agencia =RS1("CO_Agencia")
					co_conta =RS1("CO_Conta")
					da_process =RS1("DA_Processamento")
					msg01 =RS1("TX_Msg_01")
					msg02 =RS1("TX_Msg_02")
					msg03 = "" 'RS1("TX_Msg_03")
					msg04 = "" 'RS1("TX_Msg_04")				
					end_rua =RS1("NO_Logradouro_Empresa")
					end_num =RS1("NU_Logradouro_Empresa")
					end_comp =RS1("TX_Complemento_Logradouro_Empresa")
					end_bairro =RS1("NO_Bairro_Empresa")
					end_cid =RS1("NO_Cidade_Empresa")
					end_uf =RS1("SG_UF_Empresa")
					end_cep =RS1("CO_CEP_Empresa")
					no_curso=RS1("NO_Grau")
					no_etapa=RS1("NO_Serie")
					cpf_responsavel=RS1("CO_CPF")
					no_responsavel=RS1("NO_Responsavel")		
				end if		
			else	
				nu_carne=RS1("NU_Bloqueto")
				nosso_numero = "09/"&RS1("CO_Nosso_Numero")
				va_inicial = RS1("VA_Inicial")
				cod_superior=RS1("CO_Superior")				
				cod_barras =RS1("CO_Barras")
				turma =RS1("CO_Turma")
				no_cedente =RS1("NO_Cedente")
				co_agencia =RS1("CO_Agencia")
				co_conta =RS1("CO_Conta")
				da_process =RS1("DA_Processamento")
				msg01 =RS1("TX_Msg_01")
				msg02 =RS1("TX_Msg_02")
				msg03 ="" 'RS1("TX_Msg_03")
				msg04 ="" 'RS1("TX_Msg_04")				
				end_rua =RS1("NO_Logradouro_Empresa")
				end_num =RS1("NU_Logradouro_Empresa")
				end_comp =RS1("TX_Complemento_Logradouro_Empresa")
				end_bairro =RS1("NO_Bairro_Empresa")
				end_cid =RS1("NO_Cidade_Empresa")
				end_uf =RS1("SG_UF_Empresa")
				end_cep =RS1("CO_CEP_Empresa")
				no_curso=RS1("NO_Grau")
				no_etapa=RS1("NO_Serie")
				cpf_responsavel=RS1("CO_CPF")
				no_responsavel=RS1("NO_Responsavel")		
			end if		
			y_primeiro_separador = Page.Height - altura_logo_gde-46
			
			Page.Canvas.SetParams "LineWidth=0.5" 
			Page.Canvas.SetParams "LineCap=0" 
			Page.Canvas.SetParams "Dash1=2; DashPhase=1"
			With Page.Canvas
			   .MoveTo margem_x, y_primeiro_separador
			   .LineTo Page.Width - margem_x, y_primeiro_separador
			   .Stroke
			End With 				
			
			y_nome_aluno=y_primeiro_separador-5
			width_nome_aluno=Page.Width - margem_x
			
			SET Param_Nome_Aluno = Pdf.CreateParam("x="&margem_x&";y="&y_nome_aluno&"; height=50; width="&width_nome_aluno&"; alignment=left; html=True")
			Nome = "<font style=""font-size:11pt;""><b>Alun"&desinencia&": "&nome_aluno&"</b></font>"
			

			Do While Len(Nome) > 0
				CharsPrinted = Page.Canvas.DrawText(Nome, Param_Nome_Aluno, Font )
			 
				If CharsPrinted = Len(Nome) Then Exit Do
					SET Page = Page.NextPage
				Nome = Right( Nome, Len(Nome) - CharsPrinted)
			Loop 
			x_matricula = 350
			y_matricula=y_nome_aluno
			SET Param_cod_cons = Pdf.CreateParam("x="&x_matricula&";y="&y_matricula&"; height=50; width=225; alignment=right; html=False")			

			Do While Len(cod_cons) > 0
				CharsPrinted = Page.Canvas.DrawText(cod_cons, Param_cod_cons, Font )
			 
				If CharsPrinted = Len(cod_cons) Then Exit Do
					SET Page = Page.NextPage
				Nome = Right(cod_cons, Len(cod_cons) - CharsPrinted)
			Loop 	

			y_wd = y_matricula-20
			SET Param_WD = Pdf.CreateParam("x="&x_matricula&";y="&y_wd&"; height=50; width=225; alignment=right; html=False")			
			CharsPrinted = Page.Canvas.DrawText("WD", Param_WD, Font )						

	
			Set param_table1 = Pdf.CreateParam("width=533; height=40; rows=3; cols=8; border=0; cellborder=0; cellspacing=0;")
			Set Table = Doc.CreateTable(param_table1)
			Table.Font = Font
			y_table=y_nome_aluno-20
			
			With Table.Rows(1)
			   .Cells(1).Width = 40
			   .Cells(2).Width = 105
			   .Cells(3).Width = 25
			   .Cells(4).Width = 70
			   .Cells(5).Width = 60
			   .Cells(6).Width = 133
			   .Cells(7).Width = 50
			   .Cells(8).Width = 50      
			End With
			Table(1, 2).ColSpan = 7
			Table(1, 1).AddText "Sede:", "size=9;", Font 
			Table(2, 1).AddText "Curso:", "size=9;", Font 
			Table(1, 2).AddText no_sede, "size=9;", Font 
			Table(2, 2).ColSpan = 7
			Table(2, 2).AddText no_curso, "size=9;", Font 
			'Table(2, 3).AddText no_etapa, "size=9;", Font 
			Table(3, 1).ColSpan = 7		
			Table(3, 1).AddText no_etapa&"/Turma: "&turma, "size=9;", Font 
			'Table(2, 5).AddText "N&ordm;. Chamada: "&cham, "size=9; html=true", Font 
'			Table(2, 6).AddText cham, "size=9;", Font 
'			Table(1, 7).AddText "<div align=""right"">Matr&iacute;cula: </div>", "size=9; html=true", Font 
'			Table(1, 8).AddText cod_cons, "size=9;alignment=right", Font 
'			Table(2, 7).AddText "Ano Letivo: ", "size=9; alignment=right", Font 
'			Table(2, 8).AddText ano_letivo, "size=9;alignment=right", Font 
			Page.Canvas.DrawTable Table, "x="&margem_x&", y="&y_table&"" 
		
			y_segundo_separador = y_nome_aluno-65
			With Page.Canvas
			   .MoveTo margem_x, y_segundo_separador
			   .LineTo Page.Width - margem_x, y_segundo_separador
			   .Stroke
			End With 			
			
			
			'Mensagens na parte abaixo do cabeçalho com nome do aluno, sede e curso		
			Set param_table1 = Pdf.CreateParam("width=533; height=60; rows=4; cols=1; border=0; cellborder=0; cellspacing=0;")
			Set Table = Doc.CreateTable(param_table1)
			Table.Font = Font
			y_table=y_segundo_separador-10
			
			Table(1, 1).AddText msg01, "size=9; expand=true", Font 
			Table(2, 1).AddText msg02, "size=9; expand=true", Font 
			Table(3, 1).AddText msg03, "size=9; expand=true", Font 
			Table(4, 1).AddText msg04, "size=9; expand=true", Font 	
			Page.Canvas.DrawTable Table, "x="&margem_x&", y="&y_table&"" 				
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			
			multa = formatnumber(va_inicial*0.02,2)
			juros = formatnumber(va_inicial/3000,2)
			
			if isnumeric(mes_vencimento) then
				
				'mes_refcia = mes_vencimento-1
				mes_refcia = mes_vencimento
				mes_refcia = mes_refcia*1
				if mes_refcia=0 then
					mes_refcia = 12
				end if
			end if			
			nome_mes = GeraNomesNovaVersao("MES",mes_refcia,variavel2,variavel3,variavel4,variavel5,conexao,outro)
			
			TEXTO_UCASE = "" '"PARCELA REFERENTE AO M&Ecirc;S DE "&ucase(nome_mes)	
			
			Page.Canvas.SetParams "Dash1=0; DashPhase=0"

'Cria tabelas com informações de cobrança---------------------------------------------------------------------	
			wrk_y_osmar = 22
			constante_de_deslocamento = 280	+ wrk_y_osmar
			y_tabela_2=y_table-120
			for i = 1 to 2 
			
			
			var_instrucoes="<font style=""font-size:6pt;"">Instru&ccedil;&otilde;es (Todas as informa&ccedil;&otilde;es deste bloqueto s&atilde;o de exclusiva responsabilidade do benefici&aacute;rio)</font><font style=""font-size:3pt;""><BR>&nbsp;<br></FONT><font style=""font-size:7pt;"">"&msg01&"<BR>&nbsp;<BR>"&msg02
			
			var_instrucoes=	var_instrucoes&"</FONT><font style=""font-size:6pt;""><BR>&nbsp;"				
			
			var_instrucoes=	var_instrucoes&msg03&"<BR>Ap&oacute;s o vencimento pagar somente nas ag&ecirc;ncias do "&nome_banco_msg &" com MULTA de R$"&multa&" mais mora di&aacute;ria de R$"&juros&"<BR>"
			var_instrucoes=	var_instrucoes&"O imposto sobre Servi&ccedil;os j&aacute; inclu&iacute;do no pre&ccedil;o, foi calculado pela al&iacute;quota de 5% de acordo com a lei.<BR>"
			var_instrucoes=	var_instrucoes&"Servi&ccedil;os Adicionais poder&atilde;o ser cobrados a parte conforme prevista cláusula V par&aacute;grafo primeiro do contrato.<BR>"&TEXTO_UCASE&"</font>" 
				
			
				x_tabela_2=margem_x				
				rows_tabela_2 = 11
				
				width_tabela_2=Page.Width - (margem_x*2)
				height_tabela_2 = rows_tabela_2*row_padrao
	
				Set param_table2 = Pdf.CreateParam("width="&width_tabela_2&"; height=200; rows="&rows_tabela_2&"; cols=7; border=1; cellborder=0.5; cellspacing=0;")
				Set Table = Doc.CreateTable(param_table2)
				Table.Font = Font
	
				
				With Table.Rows(11)	
				   .Cells(1).Width = formatnumber(width_tabela_2/7,0) -5
				   .Cells(2).Width = formatnumber(width_tabela_2/7,0)-30
				   .Cells(3).Width = formatnumber(width_tabela_2/7,0)-30
				   .Cells(4).Width = formatnumber(width_tabela_2/7,0)-5 
				   .Cells(5).Width = formatnumber(width_tabela_2/7,0)-20
				   .Cells(6).Width = formatnumber(width_tabela_2/7,0)+10  		
				   .Cells(7).Width = formatnumber(width_tabela_2/7,0)+80 					   	   
				End With
				width_stb_1 = width_tabela_2+75'formatnumber(width_tabela_2/7,0) -5 + ((formatnumber(width_tabela_2/7,0)-30)*2)+formatnumber(width_tabela_2/7,0)-5 +formatnumber(width_tabela_2/7,0)-20+formatnumber(width_tabela_2/7,0)+10 
				width_stb_2	= formatnumber(width_tabela_2/7,0)+70 	
				width_stb_3	= width_stb_2	
				width_stb_4	= width_stb_2	
				width_stb_5	= width_stb_2	
				width_stb_6 = width_stb_2	
				width_stb_7 = width_stb_2	
				width_stb_8 = width_stb_2	
				width_stb_9 = width_stb_2		
				width_stb_10 = width_stb_2
				width_stb_11 = formatnumber(width_tabela_2/7,0) -5 + ((formatnumber(width_tabela_2/7,0)-30)*2)+formatnumber(width_tabela_2/7,0)-5 +formatnumber(width_tabela_2/7,0)-20+formatnumber(width_tabela_2/7,0)+10 
				width_stb_12 = formatnumber(width_tabela_2/7,0) -5	
				width_stb_13 = width_stb_12	
				width_stb_14 = ((formatnumber(width_tabela_2/7,0)-30)*2)
				width_stb_15 = formatnumber(width_tabela_2/7,0)-30	
				width_stb_16 = width_stb_15		
				width_stb_17 = formatnumber(width_tabela_2/7,0)-5 	
				width_stb_18 = width_stb_17		
				width_stb_19 = formatnumber(width_tabela_2/7,0)-20	
				width_stb_20 = formatnumber(width_tabela_2/7,0)-20+	formatnumber(width_tabela_2/7,0)+10  	
				width_stb_21 = formatnumber(width_tabela_2/7,0)+10  	
				width_stb_22 = width_tabela_2															
				Table.Rows(1).Cells(1).Height = 17
				Table.Rows(2).Cells(1).Height = 17
				Table.Rows(3).Cells(1).Height = 17
				Table.Rows(4).Cells(1).Height = 17
				Table.Rows(5).Cells(1).Height = 17
				Table.Rows(6).Cells(1).Height = 17
				Table.Rows(7).Cells(1).Height = 17
				Table.Rows(8).Cells(1).Height = 17	
				Table.Rows(9).Cells(1).Height = 17	
				Table.Rows(10).Cells(1).Height = 23
				Table.Rows(11).Cells(1).Height = 24																									
				Table(1, 1).ColSpan = 6
				Table(2, 1).ColSpan = 6	
				Table(3, 2).ColSpan = 2			
				Table(4, 5).ColSpan = 2		
				Table(5, 1).ColSpan = 6		
				Table(5, 1).RowSpan = 5		
				Table(10, 1).ColSpan = 7	
				Table(10, 1).RowSpan = 2				
				
				
				local_de_pagamento = "Pag&aacute;vel preferencialmente nas agências do Bradesco"				
				Set SmallTable = Doc.CreateTable("Height=17; Width="&width_stb_1&"; cols=2; rows=1; border=0; cellborder=0; cellspacing=0;")
				SmallTable.Rows(1).Cells(1).Width = 70
				SmallTable.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">Local de Pagamento:</font>", "x=1; y=16, size=5; html=true;", Font
				SmallTable.At(1, 2).AddText "<CENTER>"&local_de_pagamento&"</CENTER>", "size=7; html=true; indenty=1", Font
		
				Set SmallTable2 = Doc.CreateTable("Height=17; Width="&width_stb_2&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable2.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">Vencimento:</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable2.At(2, 1).AddText "<b><div align=""right"">"&vencimento&"</div></b> ", " alignment=right; size=7; html=true", Font		
				
				Set SmallTable3 = Doc.CreateTable("Height=17; Width="&width_stb_3&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable3.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">Ag&ecirc;ncia / C&oacute;digo Benefici&aacute;rio:</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable3.At(2, 1).AddText co_agencia&" / "&co_conta&" ", " alignment=right; size=7;", Font	
				
				Set SmallTable4 = Doc.CreateTable("Height=17; Width="&width_stb_4&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable4.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">Nosso N&uacute;mero:</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable4.At(2, 1).AddText nosso_numero&" ", " alignment=right; size=7;", Font	
				
				Set SmallTable5 = Doc.CreateTable("Height=17; Width="&width_stb_5&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable5.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">Valor do Documento:</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable5.At(2, 1).AddText "<b><div align=""right"">"&formatcurrency(va_inicial)&"</div></b> ", " alignment=right; size=7; html=true", Font				
	
				Set SmallTable6 = Doc.CreateTable("Height=17; Width="&width_stb_6&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable6.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">(-) Desconto / Abatimento</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable6.At(2, 1).AddText " ", " alignment=right; size=7;", Font				
	
				Set SmallTable7 = Doc.CreateTable("Height=17; Width="&width_stb_7&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable7.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">(-) Outras Dedu&ccedil;&otilde;es</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable7.At(2, 1).AddText " ", " alignment=right; size=7;", Font		
				
				Set SmallTable8 = Doc.CreateTable("Height=17; Width="&width_stb_8&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable8.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">(+) Mora / Multa</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable8.At(2, 1).AddText " ", " alignment=right; size=7;", Font			
	
				Set SmallTable9 = Doc.CreateTable("Height=17; Width="&width_stb_9&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable9.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">(+) Outros Acr&eacute;scimos</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable9.At(2, 1).AddText " ", " alignment=right; size=7;", Font			
	
				Set SmallTable10 = Doc.CreateTable("Height=17; Width="&width_stb_10&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable10.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">(=) Valor Cobrado</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable10.At(2, 1).AddText " ", " alignment=right; size=7;", Font	
				
				'Set SmallTable11 = Doc.CreateTable("Height=17; Width="&width_stb_11&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				'SmallTable11.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">Benefici&aacute;rio</font>", "x=1; y=8, size=5; html=true", Font
				'SmallTable11.At(2, 1).AddText "<CENTER>"&no_cedente&"</CENTER>", " html=true; size=7;", Font	
				Set SmallTable11 = Doc.CreateTable("Height=17; Width="&width_stb_11&"; cols=4; rows=1; border=0; cellborder=0; cellspacing=0;")
				SmallTable11.Rows(1).Cells(1).Width = 70
				SmallTable11.Rows(1).Cells(2).Width = 200								
				SmallTable11.Rows(1).Cells(3).Width = 30												
				SmallTable11.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">Benefici&aacute;rio</font>", "x=1; y=16, size=5; html=true", Font
				SmallTable11.At(1, 2).AddText no_cedente&"<BR>"&endereco_sede, " html=true; size=6;", Font	
				SmallTable11.At(1, 3).Canvas.DrawText "<font style=""font-size:5pt;"">CNPJ:</font>", "x=1; y=16, size=5; html=true", Font						
				SmallTable11.At(1, 4).AddText "<CENTER>"&cnpj_formatado&"</CENTER>", " html=true; size=7;", Font														
	
				dia=dia*1
				if dia<10 then
					dia="0"&dia
				end if
				
				mes=mes*1
				if mes<10 then
					mes="0"&mes
				end if
				
				data_documento = dia&"/"&mes&"/"&ano
											
				Set SmallTable12 = Doc.CreateTable("Height=17; Width="&width_stb_12&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable12.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">Data Documento</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable12.At(2, 1).AddText "<CENTER>"&data_documento&"</CENTER>", " html=true; size=7;", Font	
				
				Set SmallTable13 = Doc.CreateTable("Height=17; Width="&width_stb_13&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable13.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">Uso do Banco</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable13.At(2, 1).AddText "&nbsp;", " html=true; size=7;", Font			
	
				Set SmallTable14 = Doc.CreateTable("Height=17; Width="&width_stb_14&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable14.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">N&ordm; do Documento</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable14.At(2, 1).AddText "<CENTER>"&nu_cota&"</CENTER>", " html=true; size=7;", Font	
	
				Set SmallTable15 = Doc.CreateTable("Height=17; Width="&width_stb_15&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable15.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">Carteira</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable15.At(2, 1).AddText "<CENTER>09</CENTER>", " html=true; size=7;", Font			
	
				Set SmallTable16 = Doc.CreateTable("Height=17; Width="&width_stb_16&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable16.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">Moeda</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable16.At(2, 1).AddText "<CENTER>R$</CENTER>", " html=true; size=7;", Font	
				
				Set SmallTable17 = Doc.CreateTable("Height=17; Width="&width_stb_17&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable17.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">Esp&eacute;cie Doc.</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable17.At(2, 1).AddText "<CENTER>OU</CENTER>", " html=true; size=7;", Font		
				
				Set SmallTable18 = Doc.CreateTable("Height=17; Width="&width_stb_18&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable18.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">Quantidade</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable18.At(2, 1).AddText "<CENTER>&nbsp;</CENTER>", " html=true; size=7;", Font		
				
				Set SmallTable19 = Doc.CreateTable("Height=17; Width="&width_stb_19&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable19.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">Aceite</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable19.At(2, 1).AddText "<CENTER>&nbsp;</CENTER>", " html=true; size=7;", Font		
				
				Set SmallTable20 = Doc.CreateTable("Height=17; Width="&width_stb_20&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable20.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">Valor</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable20.At(2, 1).AddText "<CENTER>&nbsp;</CENTER>", " html=true; size=7;", Font		
				
				Set SmallTable21 = Doc.CreateTable("Height=17; Width="&width_stb_21&"; cols=1; rows=2; border=0; cellborder=0; cellspacing=0;")
				SmallTable21.At(1, 1).Canvas.DrawText "<font style=""font-size:5pt;"">Data Processamento</font>", "x=1; y=8, size=5; html=true", Font
				SmallTable21.At(2, 1).AddText "<CENTER>"&da_process&"</CENTER>", " html=true; size=7;", Font																									
	
	
				Set SmallTable22 = Doc.CreateTable("Height=45; Width="&width_stb_22&"; cols=4; rows=4; border=0; cellborder=0; cellspacing=0;")
				SmallTable22.At(1, 1).Canvas.DrawText "<font style=""font-size:7pt;"">Pagador:</font>", "x=1; y=11, size=7; html=true", Font
				SmallTable22.At(1, 2).AddText "<B>"&no_responsavel&"</B>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CPF:&nbsp;"&cpf_responsavel, " html=true; size=7;", Font	
				
				end_cep = Left(end_cep,5)&"-"&Right(end_cep,3)
				SmallTable22.At(2, 2).AddText end_rua &", "& end_num&"/ "&end_comp, " html=true; size=7;", Font	
				SmallTable22.At(3, 2).AddText end_bairro&"&nbsp;-&nbsp;"&end_cid&"&nbsp;-&nbsp;"&end_uf, " html=true; size=7;", Font		
				SmallTable22.At(4, 2).AddText end_cep, " html=true; size=7;", Font											
				
				SmallTable22.At(1, 3).Canvas.DrawText "<font style=""font-size:7pt;""><b>"&cod_cons&" - "&nome_aluno&"</b></font>", "x=1; y=11, size=7; html=true", Font	
				SmallTable22.At(2, 3).Canvas.DrawText "Curso: "&no_curso, "x=1; y=11, size=7; html=true", Font						
				SmallTable22.At(3, 3).Canvas.DrawText no_etapa&"/Turma:"&turma, "x=1; y=11, size=7; html=true", Font	
				'SmallTable22.At(4, 3).Canvas.DrawText turma, "x=1; y=11, size=7; html=true", Font	
				
				'vetor_nome = SPLIT(nome_aluno, " ")		
				'SmallTable22.At(1, 4).AddText "<CENTER>"&vetor_nome(0)&"</CENTER>", "size=9; indenty=12; html=true", Font
										
				With SmallTable22.Rows(4)	
				   .Cells(1).Width = 35
				   .Cells(2).Width = 296
				   .Cells(3).Width = 120
				   .Cells(4).Width = 100			   	   
				End With
				SmallTable22(1, 4).RowSpan = 4
		
														
				Table(1, 1).Canvas.DrawTable SmallTable, "x=0; y=17"
				Table(1, 7).Canvas.DrawTable SmallTable2, "x=0; y=17"	
				Table(2, 1).Canvas.DrawTable SmallTable11, "x=0; y=17"			
				Table(2, 7).Canvas.DrawTable SmallTable3, "x=0; y=17"	
				Table(3, 1).Canvas.DrawTable SmallTable12, "x=0; y=17"		
				Table(3, 2).Canvas.DrawTable SmallTable14, "x=0; y=17"	
				Table(3, 4).Canvas.DrawTable SmallTable17, "x=0; y=17"		
				Table(3, 5).Canvas.DrawTable SmallTable19, "x=0; y=17"		
				Table(3, 6).Canvas.DrawTable SmallTable21, "x=0; y=17"													
				Table(3, 7).Canvas.DrawTable SmallTable4, "x=0; y=17"	
				Table(4, 1).Canvas.DrawTable SmallTable13, "x=0; y=17"	
				Table(4, 2).Canvas.DrawTable SmallTable15, "x=0; y=17"		
				Table(4, 3).Canvas.DrawTable SmallTable16, "x=0; y=17"		
				Table(4, 4).Canvas.DrawTable SmallTable18, "x=0; y=17"		
				Table(4, 5).Canvas.DrawTable SmallTable20, "x=0; y=17"														
				Table(4, 7).Canvas.DrawTable SmallTable5, "x=0; y=17"
				Table(5, 1).AddText var_instrucoes, " size=7; html=true; indentx=1", Font	
				Table(5, 7).Canvas.DrawTable SmallTable6, "x=0; y=17"		
				Table(6, 7).Canvas.DrawTable SmallTable7, "x=0; y=17"	
				Table(7, 7).Canvas.DrawTable SmallTable8, "x=0; y=17"	
				Table(8, 7).Canvas.DrawTable SmallTable9, "x=0; y=17"	
				Table(9, 7).Canvas.DrawTable SmallTable10, "x=0; y=17"		
				Table(10, 1).Canvas.DrawTable SmallTable22, "x=0; y=45"					
					
	
				Page.Canvas.DrawTable Table, "x="&x_tabela_2&", y="&y_tabela_2&""
			
			
			

				y_tabela_2=formatnumber(y_tabela_2 - constante_de_deslocamento,0)			
			
			next
'Fim da tabela de informações de cobrança--------------------------------------------------------------------		
			Page.Canvas.SetParams "Dash1=2; DashPhase=1"
			texto_3 = "<div align=""right""><font style=""font-size:6pt;"">Autentica&ccedil;&atilde;o Mec&acirc;nica - Ficha de Compensa&ccedil;&atilde;o</fint></div>"
			
			x_texto_3 = 350
			y_texto_3 = formatnumber(y_table - 325,0)
			for j = 1 to 2 
				SET Param_texto_3 = Pdf.CreateParam("x="&x_texto_3&";y="&y_texto_3&"; height=50; width=225; html=True")			

				Do While Len(texto_3) > 0
					CharsPrinted = Page.Canvas.DrawText(texto_3, Param_texto_3, Font )
				 
					If CharsPrinted = Len(texto_3) Then Exit Do
						SET Page = Page.NextPage
					Nome = Right(texto_3, Len(texto_3) - CharsPrinted)
				Loop 
				y_texto_3 = y_texto_3-constante_de_deslocamento
			next


			
			Page.Canvas.SetParams "LineWidth=0.5; LineCap=0; Dash1=3; DashPhase=0" 
			x_primeira_linha= page.width-margem_x
			y_primeira_linha = formatnumber(y_table - 350 - wrk_y_osmar -margem_y,0)			
				With Page.Canvas
				   .MoveTo margem_x, y_primeira_linha
				   .LineTo x_primeira_linha, y_primeira_linha
				   .Stroke
				End With 

			Page.Canvas.SetParams "Dash1=0; DashPhase=0"

			x_texto_4 = margem_x+40
			y_texto_4 = y_primeira_linha-(margem_y/2)

			Set Param_Itau = Pdf.CreateParam	
			largura_Itau=formatnumber(Itau.Width*0.6,0)
			altura_Itau=formatnumber(Itau.Height*0.6,0)						
			Param_Itau("x") = margem_x
			Param_Itau("y") = y_texto_4-altura_Itau-2
			Param_Itau("ScaleX") = 0.6
			Param_Itau("ScaleY") = 0.6
			Page.Canvas.DrawImage Itau, Param_Itau
	

			texto_4 = "<font style=""font-size:10pt;"">"&nome_banco_tit&"</FONT><font style=""font-size:16pt;""><b> |"&codigo_banco&"|</b></FONT> <font style=""font-size:13pt;"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&cod_superior&"</FONT>"
			

			width_texto_4 = Page.Width - (margem_x*2)
			SET Param = Pdf.CreateParam("x="&x_texto_4&";y="&y_texto_4&"; height="&row_padrao&"; width="&width_texto_4&"; alignment=RIGHT; size=5.5; color=#000000; html=true")							
				
			Do While Len(texto_4) > 0
				CharsPrinted = Page.Canvas.DrawText(texto_4, Param, Font )
			 
				If CharsPrinted = Len(texto_4) Then Exit Do
					SET Page = Page.NextPage
				texto_4 = Right( texto_4, Len(texto_4) - CharsPrinted)
			Loop 


			
			x_barcode=margem_x + 19' A distância mínima da margem da ficha é de 5 mm
			y_barcode=formatnumber(y_texto_4 - 270,0) ' A distância mínima da ficha é de 12 mm (49 px de espaço)
			width_barcode=389-20 ' o tamanho deve ser 103mm	
							  ' A altura deverá ser 13mm 
			strParam = "x="&x_barcode&"; y="&y_barcode&"; height=44; width="&width_barcode&"; type=12" 'Barcode type 1 is UPC-A
			strData = cod_barras
			Page.Canvas.DrawBarcode strData, strParam 			 			

NEXT	
		End IF	
			

	


arquivo="boleto"

if salvar = "S" then
    nomeArquivo = caminho_gera_mov&arquivo&complemento & ".pdf"
	Filename = Doc.Save(nomeArquivo, False )
	'response.Write("Arquivo salvo em: "&nomeArquivo&"<br>")	
	'response.Write("enviar e-mail")	
	GeraBloqueto = nomeArquivo
else
	Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
end if	
End Function
%>