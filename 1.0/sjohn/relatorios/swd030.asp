 	<%'On Error Resume Next%>
<%
'Server.ScriptTimeout = 60
Server.ScriptTimeout = 600 'valor em segundos
'Registro de inadimplencia
%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes.asp"--> 
<!--#include file="../inc/funcoes6.asp"--> 
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/bd_parametros.asp"-->
      <% 

response.Charset="ISO-8859-1"
obr = request.querystring("obr")
opt = REQUEST.QueryString("opt")
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
nvg=session("nvg")
session("nvg")=nvg

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
		
		Set CON6 = Server.CreateObject("ADODB.Connection") 
		ABRIR6 = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON6.Open ABRIR6			
		
		Set CON7 = Server.CreateObject("ADODB.Connection") 
		ABRIR7 = "DBQ="& CAMINHO_pf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON7.Open ABRIR7			

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

dados= split(obr, "_" )
unidade_form = dados(0)
curso_form = dados(1)
etapa_form = dados(2)
tipo_parcela = dados(3)
ano_form = dados(4)




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

if opt="p" then

	data_inicio = "01/01/"& ano_form
	if dia<10 then
		diawrt="0"&dia
	else
		diawrt=dia
	end if	
	data_fim = diawrt &"/"& meswrt &"/"& ano_form
	data_de=data_inicio
			
	data_ate=meswrt&"/"&diawrt&"/"&ano_form		

'	if dia = 1 then
'		if mes =1 then
'         ano_form=ano_form*1
'		 ano_ate= ano_form-1	
'		 dia_data_ate = 31
'		 mes_data_ate = 12		 	 
'		 data_ate=mes_data_ate&"/"&dia_data_ate&"/"&ano_ate	
'		 ano_data_ate = ano_ate			 
'		else
'			if mes =2 then
'			 dia_data_ate = 31
'			 mes_data_ate = 12						
'			elseif mes =4 then
'			 dia_data_ate = 31
'			 mes_data_ate = 3		
'			elseif mes =6 then
'			 dia_data_ate = 31
'			 mes_data_ate = 5	
'			elseif mes =8 then
'			 dia_data_ate = 31
'			 mes_data_ate = 7	
'			elseif mes =9 then
'			 dia_data_ate = 31
'			 mes_data_ate = 8	
'			elseif mes =11 then
'			 dia_data_ate = 31
'			 mes_data_ate = 10			 		 		 		 			 	
'			elseif mes =5 then
'			 dia_data_ate = 30
'			 mes_data_ate = 4	
'			elseif mes =7 then
'			 dia_data_ate = 30
'			 mes_data_ate = 6		
'			elseif mes =10 then
'			 dia_data_ate = 30
'			 mes_data_ate = 9	
'			elseif mes =12 then
'			 dia_data_ate = 30
'			 mes_data_ate = 11	
'			elseif mes =3 then
'				if ano_form MOD 4 = 0 then
'					dia_data_ate = "29"
'				else
'					dia_data_ate = "28"			
'				end if	
'			 mes_data_ate = 2						
'			end if
'			if  mes_data_ate<10 then
'				mes_data_ate_wrk = "0"&mes_data_ate
'			else
'				mes_data_ate_wrk = mes_data_ate			
'			end if
'   		  data_ate=mes_data_ate_wrk&"/"&dia_data_ate&"/"&ano_form
'		  ano_data_ate = ano_form			  	
'		end if						 
'	else	
'		dia=dia*1
'		dia_data_ate= dia-1
'		if dia_data_ate<10 then
'			dia_ate_wrt="0"&dia_data_ate
'		else
'			dia_ate_wrt=dia_data_ate
'		end if		
'		mes_data_ate = mes
'		if  mes_data_ate<10 then
'			mes_data_ate_wrk = "0"&mes_data_ate
'		else
'			mes_data_ate_wrk = mes_data_ate			
'		end if			
'   		 data_ate=mes_data_ate_wrk&"/"&dia_data_ate&"/"&ano_form		
'		ano_data_ate = ano_form			 		
'	End if	

else
		IF opt="Jan" THEN
			mes_teste = 1
			mes_cobr = "01"
			ult_dia = "31"
			dia_data_ate=31
		ELSEIF opt="Fev" THEN
			mes_teste = 2		
			mes_cobr = "02"
			if ano_letivo MOD 4 = 0 then
				ult_dia = "29"
				dia_data_ate=29				
			else
				ult_dia = "28"		
				dia_data_ate=28					
			end if	
		ELSEIF opt="Mar" THEN
			mes_teste = 3		
			mes_cobr = "03"
			ult_dia = "31"
			dia_data_ate=31					
		ELSEIF opt="Abr" THEN
			mes_teste = 4		
			mes_cobr = "04"
			ult_dia = "30"
			dia_data_ate=30						
		ELSEIF opt="Mai" THEN
			mes_teste = 5		
			mes_cobr = "05"
			ult_dia = "31"
			dia_data_ate=31						
		ELSEIF opt="Jun" THEN
			mes_teste = 6		
			mes_cobr = "06"
			ult_dia = "30"
			dia_data_ate=30			
		ELSEIF opt="Jul" THEN
			mes_teste = 7		
			mes_cobr = "07"
			ult_dia = "31"
			dia_data_ate=31						
		ELSEIF opt="Ago" THEN
			mes_teste = 8		
			mes_cobr = "08"
			ult_dia = "31"
			dia_data_ate=31						
		ELSEIF opt="Set" THEN
			mes_teste = 9		
			mes_cobr = "09"
			ult_dia = "30"
			dia_data_ate=30			
		ELSEIF opt="Out" THEN
			mes_teste = 10		
			mes_cobr = "10"
			ult_dia = "31"
			dia_data_ate=31						
		ELSEIF opt="Nov" THEN
			mes_teste = 11		
			mes_cobr = "11"
			ult_dia = "30"
			dia_data_ate=30			
		ELSEIF opt="Dez" THEN
			mes_teste = 12		
			mes_cobr = "12"
			ult_dia = "31"	
			dia_data_ate=31																																		
		END IF	
		mes_data_ate = mes_teste
		if mes=mes_teste then
			ult_dia = dia
			if ult_dia<10 then
				ult_dia="0"&ult_dia
			end if				
		end if
	data_inicio = "01/"&mes_cobr&"/"& ano_form
	data_fim = ult_dia&"/"&mes_cobr&"/"& ano_form
	
	data_de=mes_cobr&"/01/"&ano_form
	data_ate=mes_cobr&"/"&ult_dia&"/"&ano_form	
'	ano_data_ate = ano_form	
end if

	data_calc=dia&"/"&mes&"/"&ano_form	
	
if unidade_form="nulo" then
	sql_unidade =""
	unidade_form=""
	
	
	
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0
		
		NU_Unidade_Check=999999	
		unidade_check = 0	
		While not RS0.EOF	
			unidade_bd=RS0("NU_Unidade")
			if unidade_check = 0 then
				unidade_form = unidade_bd
				unidade_padrao=unidade_form
			else
				unidade_form = unidade_form&","&unidade_bd	
			end if
		unidade_check = unidade_check+1	
		RS0.MOVENEXT
		WEND		
else
	sql_unidade = " AND TB_Matriculas.NU_Unidade = "& unidade_form
	unidade_padrao = unidade_form
end if	

if isnull(curso_form)or curso_form="" then
	sql_curso =""
	curso_form=""
else
	if isnumeric(curso_form) then
		curso_form=curso_form*1
	 if curso_form=999990 then
		sql_curso =""
		curso_form=""	 
	 else
		sql_curso = " AND TB_Matriculas.CO_Curso = '"& curso_form&"'"
	end if	
   else
	 if curso_form="999990" then
		sql_curso =""
		curso_form=""	 
	 else
		sql_curso = " AND TB_Matriculas.CO_Curso = '"& curso_form&"'"
	end if	   
   end if	
end if	

if isnull(etapa_form)or etapa_form="" then
	sql_etapa =""
	etapa_form=""
else
	if isnumeric(etapa_form) then
		etapa_form=etapa_form*1
	 if etapa_form=999990 then
		sql_etapa =""
		etapa_form=""	 
	 else
		sql_etapa = " AND TB_Matriculas.CO_Etapa = '"& etapa_form&"'"
	end if	
   else
	 if etapa_form="999990" then
		sql_etapa =""
		etapa_form=""	 
	 else
		sql_etapa = " AND TB_Matriculas.CO_Etapa = '"& etapa_form&"'"
	end if	   
   end if	
end if	

if tipo_parcela="nulo" then
	sql_parcela =""
else
	sql_parcela =" AND NO_Lancamento = '"&tipo_parcela&"'"

end if	

'	Set RSP = Server.CreateObject("ADODB.Recordset")
'	SQLP ="SELECT VA_Mora, VA_Multa FROM TB_Correcao"
'	RSP.Open SQLP, CON0	
'	
'	IF EOF THEN
'		mora="nulo"
'		multa="nulo"
'	ELSE
'		mora=RSP("VA_Mora")
'		multa=RSP("VA_Multa")	
'	END IF	 		 	


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
	
'		
	Set RS2 = Server.CreateObject("ADODB.Recordset")
	SQL2 = "SELECT * FROM TB_Unidade WHERE NU_Unidade="&unidade_padrao
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
		
	if complemento_unidade=" " or complemento_unidade="" or isnull(complemento_unidade)then
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
		municipio_bairro=municipio_unidade
		municipio_unidade=GeraNomesNovaVersao("Mun",uf_unidade_municipio,municipio_unidade,variavel3,variavel4,variavel5,CON0,outro)
		if bairro_unidade="" or isnull(bairro_unidade)then
		else
			bairro_unidade=GeraNomesNovaVersao("Bai",uf_unidade_municipio,municipio_bairro,bairro_unidade,variavel4,variavel5,CON0,outro)					
			bairro_unidade=" - "&bairro_unidade
		end if					
									
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
			Text = "<center><i><b><font style=""font-size:18pt;"">Rela&ccedil;&atilde;o de Inadimplentes</font></b></i></center>"
			
			
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

			
		altura_unidade = altura_segundo_separador
		y_paginacao = altura_unidade		
	'================================================================================================================	
		

		NU_Unidade_Check=999999	

		geral_acumula_original = 0
		geral_acumula_multa = 0
		geral_acumula_mora = 0
		geral_acumula_corrigido = 0
					
		colunas_de_notas=12
		total_de_colunas=12					
		altura_medias=30	
			
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade Where NU_Unidade IN("&unidade_form&") order by NO_Abr"
		RS0.Open SQL0, CON0		
		
		Set param_materias = PDF.CreateParam	
		param_materias.Set "size=8" 			

	
		param_materias.Add "indenty=2;alignment=right;html=true"
		param_materias.Add "indentx=0"			
		
		While not RS0.EOF	
			nu_unidade=RS0("NU_Unidade")
			nome_unidade=RS0("NO_Unidade")
			
			altura_unidade = altura_unidade-10
			y_paginacao = y_paginacao-10
			if NU_Unidade_Check <> nu_unidade then
				
				IF NU_Unidade_Check<>999999	THEN
				
					Set Page = Page.NextPage	
					
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
						Text = "<center><i><b><font style=""font-size:18pt;"">Rela&ccedil;&atilde;o de Inadimplentes</font></b></i></center>"
						
						
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
									
						altura_unidade = altura_segundo_separador
				'================================================================================================================		
						altura_titulo = altura_unidade-10
						
						SET Param_per = Pdf.CreateParam("x="&margem&";y="&altura_titulo&"; height=30; width="&width_texto&"; alignment=center; size=12; color=#000000; html=true")
						Text_per = "<left><i><b><font style=""font-size:12pt;"">"&nome_unidade&" - Per&iacute;odo de "&data_inicio&" at&eacute; "&data_fim&"</font></b></i></left>"
						
						
						Do While Len(Text_per) > 0
							CharsPrinted = Page.Canvas.DrawText(Text_per, Param_per, Font )
						 
							If CharsPrinted = Len(Text_per) Then Exit Do
								SET Page = Page.NextPage
							Text_per = Right( Text_per, Len(Text_per) - CharsPrinted)
						Loop 	
									altura_unidade = altura_unidade-10
				else
					SET Param_per = Pdf.CreateParam("x="&margem&";y="&altura_unidade&"; height=30; width="&width_texto&"; alignment=center; size=12; color=#000000; html=true")
					Text_per = "<left><i><b><font style=""font-size:12pt;"">"&nome_unidade&" - Per&iacute;odo de "&data_inicio&" at&eacute; "&data_fim&"</font></b></i></left>"
					
					
					Do While Len(Text_per) > 0
						CharsPrinted = Page.Canvas.DrawText(Text_per, Param_per, Font )
					 
						If CharsPrinted = Len(Text_per) Then Exit Do
							SET Page = Page.NextPage
						Text_per = Right( Text_per, Len(Text_per) - CharsPrinted)
					Loop 																

				END IF			
							
							
				NU_Unidade_Check = nu_unidade
												
			y_segunda_tabela=altura_unidade-25	
			altura_unidade = altura_unidade-25
			y_paginacao = y_segunda_tabela-30					
			Set param_table2 = Pdf.CreateParam("width="&area_utilizavel&"; height=30; rows=2; cols=12; border=0; cellborder=0.1; cellspacing=0; x="&margem&"; y="&y_segunda_tabela&"; MaxHeight=400")

			Set Notas_Tit = Doc.CreateTable(param_table2)
			Notas_Tit.Font = Font				
			largura_colunas=(area_utilizavel-50-210)/colunas_de_notas		
			
			With Notas_Tit.Rows(1)
			   .Cells(1).Width = 45
			   .Cells(2).Width = 145	
			   .Cells(3).Width = 145
			   .Cells(4).Width = 55			             
			   .Cells(5).Width = 50
			   .Cells(6).Width = 40			             
			   .Cells(7).Width = 40
			   .Cells(8).Width = 50
			   .Cells(9).Width = 125
			   .Cells(10).Width = 30			   			             
			   .Cells(11).Width = 35
			   .Cells(12).Width = 35			             

			End With
			
				Notas_Tit(1, 1).RowSpan = 2	
				Notas_Tit(1, 2).RowSpan = 2		
				Notas_Tit(1, 3).RowSpan = 2		
				Notas_Tit(1, 4).RowSpan = 2		
				Notas_Tit(1, 5).RowSpan = 2		
				Notas_Tit(1, 6).RowSpan = 2		
				Notas_Tit(1, 7).RowSpan = 2
				Notas_Tit(1, 8).RowSpan = 2		
				Notas_Tit(1, 9).RowSpan = 2		
				Notas_Tit(1, 10).RowSpan = 2	
				Notas_Tit(1, 11).RowSpan = 2	
				Notas_Tit(1, 12).RowSpan = 2	
																				
				Notas_Tit(1, 1).AddText "<div align=""center"">Matr&iacute;cula</div>", "size=9;indenty=7; html=true", Font 
				Notas_Tit(1, 2).AddText "<div align=""center"">Nome</div>", "size=9;alignment=center; indenty=7;html=true", Font 
				Notas_Tit(1, 3).AddText "<div align=""center"">Respons&aacute;vel Financeiro</div>", "size=9;alignment=center; indenty=7;html=true", Font 
				Notas_Tit(1, 4).AddText "<div align=""center"">Data de Vencimento</div>", "size=9;alignment=center; indenty=2;html=true", Font 	
				Notas_Tit(1, 5).AddText "<div align=""center"">Valor Original</div>", "size=9;alignment=center; indenty=2;html=true", Font 					
				Notas_Tit(1, 6).AddText "<div align=""center"">Multa</div>", "size=9; indenty=7;html=true", Font 
				Notas_Tit(1, 7).AddText "<div align=""center"">Corre&ccedil;&atilde;o</div>", "size=9;alignment=center; indenty=7;html=true", Font 
				Notas_Tit(1, 8).AddText "<div align=""center"">Valor Corrigido</div>", "size=9;alignment=center; indenty=2;html=true", Font 			
				Notas_Tit(1, 9).AddText "<div align=""center"">Tipo de Lan&ccedil;amento</div>", "size=9;alignment=center; indenty=2;html=true", Font 
				Notas_Tit(1, 10).AddText "<div align=""center"">Curso</div>", "size=9;alignment=center; indenty=7;html=true", Font 
				Notas_Tit(1, 11).AddText "<div align=""center"">Etapa</div>", "size=9; indenty=7;html=true", Font 
				Notas_Tit(1, 12).AddText "<div align=""center"">Turma</div>", "size=9; indenty=7;html=true", Font 
				
				altura_unidade = altura_unidade - 30	
				y_paginacao = y_paginacao-30	
				linha=2
				
			end if


			Set Rs_ordena = Server.CreateObject ( "ADODB.RecordSet" )
			'O método Append recebe 3 parâmetros:
			'Nome do campo, Tipo, Tamanho (opcional)
			'O tipo pertence à um DataTypeEnum, e você pode conferir os tipos em
			'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/ado270/htm/mdcstdatatypeenum.asp
			'200 -> VarChar (String), 7 -> Data, 139 -> Numeric, 6 -> currency
			Rs_ordena.Fields.Append "matricula", 139, 10
			Rs_ordena.Fields.Append "nome_aluno", 200, 255
			Rs_ordena.Fields.Append "nome_responsavel", 200, 255
			Rs_ordena.Fields.Append "vencimento", 7
			Rs_ordena.Fields.Append "val_original", 6
			Rs_ordena.Fields.Append "multa", 6
			Rs_ordena.Fields.Append "mora", 6
			Rs_ordena.Fields.Append "val_corrigido", 6				
			Rs_ordena.Fields.Append "tipo_lancamento", 200, 255
			Rs_ordena.Fields.Append "unidade", 200, 255
			Rs_ordena.Fields.Append "curso", 200, 255
			Rs_ordena.Fields.Append "etapa", 200, 255
			Rs_ordena.Fields.Append "turma", 200, 255
										
			Rs_ordena.Open		
		
			Set RSA = Server.CreateObject("ADODB.Recordset")
			CONEXAOA = "Select TB_Matriculas.CO_Matricula, TB_Matriculas.CO_Situacao, TB_Matriculas.NU_Unidade, TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma, TB_Alunos.NO_Aluno, TB_Alunos.TP_Resp_Fin from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND NU_Unidade = "&nu_unidade&sql_curso&sql_etapa&" AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula order by TB_Matriculas.NU_Unidade ASC, TB_Matriculas.CO_Curso ASC, TB_Matriculas.CO_Etapa ASC, TB_Matriculas.CO_Turma ASC, TB_Alunos.NO_Aluno ASC"		
			Set RSA = CON1.Execute(CONEXAOA)
		
			vetor_matriculas="" 
			nu_seq_aluno=0
			nu_chamada_conta=1					

			While Not RSA.EOF
					
				nu_seq_aluno=nu_seq_aluno+1
				nu_matricula = RSA("CO_Matricula")		
				nome_aluno= RSA("NO_Aluno")	
				'co_situacao = RSA("CO_Situacao")	
				tp_resp_fin= RSA("TP_Resp_Fin")							
				unidade_aluno =	RSA("NU_Unidade")	
				curso_aluno =	RSA("CO_Curso")
				etapa_aluno =	RSA("CO_Etapa")
				turma_aluno =	RSA("CO_Turma")		
				
				Set RSc = Server.CreateObject("ADODB.Recordset")
				SQLc = "SELECT NO_Contato,CO_CPF_PFisica, TX_EMail FROM TB_Contatos where CO_Matricula = "& nu_matricula &" AND TP_Contato = '"& tp_resp_fin&"'"
				RSc.Open SQLc, CON6		
				
				If RSc.EOF then
					nome_resp ="Nome n&atilde;o cadastrado para o Respons&aacute;vel Financeiro"
				else
					nome_resp = RSc("NO_Contato")
				end if					
				nome_aluno=replace_latin_char(nome_aluno,"html")				
				nome_resp=replace_latin_char(nome_resp,"html")	
				
				
				Set RSM = Server.CreateObject("ADODB.Recordset")
				SQLM ="SELECT DA_Vencimento, VA_Compromisso, NO_Lancamento FROM TB_Posicao where DA_Realizado is NULL AND (DA_Vencimento BETWEEN #"&data_de&"# AND #"&data_ate&"#) AND CO_Matricula_Escola ="& nu_matricula &sql_parcela&" ORDER BY Mes"

				RSM.Open SQLM, CON7	
				
				if RSM.EOF then
				
				else
						
					While not RSM.EOF 
						vencimento= RSM("DA_Vencimento")
						val_original= RSM("VA_Compromisso")
						tipo_lancamento= RSM("NO_Lancamento")				

'						qtd_dias = DateDiff("d",vencimento,data_calc)	
												
'						if multa = "nulo" then
'							val_multa = 0				
'						else
'							val_multa = val_original*(multa/100)						
'						end if	
						
						val_multa = CalculaMulta(vencimento, data_calc, val_original)
						
'						if mora = "nulo" then
'							val_mora = 0
'						else
'							val_mora = val_original*(mora/100)*qtd_dias
'							val_mora = val_mora*100
'							val_mora = INT(val_mora)
'							val_mora = val_mora/100
'						end if			
						
						val_mora = CalculaMora(vencimento, data_calc, val_original)
						val_corrigido = val_original + val_multa + val_mora															
					
						Rs_ordena.AddNew
						Rs_ordena.Fields("matricula").Value = nu_matricula
						Rs_ordena.Fields("nome_aluno").Value = nome_aluno
						Rs_ordena.Fields("nome_responsavel").Value = nome_resp
						Rs_ordena.Fields("vencimento").Value = vencimento			
						Rs_ordena.Fields("val_original").Value = val_original
						Rs_ordena.Fields("multa").Value = val_multa					
						Rs_ordena.Fields("mora").Value = val_mora
						Rs_ordena.Fields("val_corrigido").Value = val_corrigido		
						Rs_ordena.Fields("tipo_lancamento").Value = tipo_lancamento						
						Rs_ordena.Fields("unidade").Value = unidade_aluno
						Rs_ordena.Fields("curso").Value = curso_aluno
						Rs_ordena.Fields("etapa").Value = etapa_aluno
						Rs_ordena.Fields("turma").Value = turma_aluno

					RSM.MOVENEXT
					WEND
				End if			
			nu_chamada_conta=nu_chamada_conta+1		
			RSA.MoveNext
			Wend 
			Rs_ordena.Sort = "unidade ASC, curso ASC, etapa ASC, turma ASC, nome_aluno ASC, vencimento ASC, tipo_lancamento ASC"		
			
			nome_aluno="nulo"		
			conta_inadimplencia_aluno = 0		
			While not Rs_ordena.EOF				

				if nome_aluno<> Rs_ordena.Fields("nome_aluno") and nome_aluno<>"nulo" and conta_inadimplencia_aluno>1 then															
					linha=linha+1						
					Set Row = Notas_Tit.Rows.Add(13) ' row height	
					altura_unidade = altura_unidade-13					
					y_paginacao=y_paginacao-13	
					param_materias.Add "expand=true" 												
					Notas_Tit(linha, 1).ColSpan = 3						
					Notas_Tit(linha, 1).AddText "<div align=""center"">Total do Aluno</div>", param_materias
					Notas_Tit(linha, 5).AddText "<div align=""right"">"&aluno_acumulado_original&"&nbsp;</div>", param_materias																						
					Notas_Tit(linha, 6).AddText "<div align=""right"">"&aluno_acumulado_multa&"&nbsp;</div>", param_materias		
					Notas_Tit(linha, 7).AddText "<div align=""right"">"&aluno_acumulado_mora&"&nbsp;</div>", param_materias	 	
					Notas_Tit(linha, 8).AddText "<div align=""right"">"&aluno_acumulado_corrigido&"&nbsp;</div>", param_materias	 									
							
				end if	
													

				Set Row = Notas_Tit.Rows.Add(13) ' row height									
	
				linha=linha+1
				altura_unidade = altura_unidade-13					
				param_materias.Add "expand=true" 	
				
				no_curso = GeraNomesNovaVersao("CA",Rs_ordena.Fields("curso").Value,variavel2,variavel3,variavel4,variavel5,CON0,outro)
				'no_etapa = GeraNomesNovaVersao("E",Rs_ordena.Fields("curso").Value,Rs_ordena.Fields("etapa").Value,variavel3,variavel4,variavel5,CON0,outro)
				
				
				if nome_aluno<> Rs_ordena.Fields("nome_aluno") then		
					Notas_Tit(linha, 1).AddText "<div align=""center"">"&Rs_ordena.Fields("matricula")&"</div>", param_materias
					Notas_Tit(linha, 2).AddText "<div align=""center"">"&Rs_ordena.Fields("nome_aluno")&"</div>", param_materias	
					Notas_Tit(linha, 3).AddText "<div align=""center"">"&Rs_ordena.Fields("nome_responsavel")&"</div>", param_materias	
					aluno_acumula_original = 0
					aluno_acumula_multa = 0
					aluno_acumula_mora = 0
					aluno_acumula_corrigido = 0	
					conta_inadimplencia_aluno = 0						
					
				end if		
				conta_inadimplencia_aluno = conta_inadimplencia_aluno+1									
				Notas_Tit(linha, 4).AddText "<div align=""center"">"&Rs_ordena.Fields("vencimento")&"</div>", param_materias	
				Notas_Tit(linha, 5).AddText "<div align=""right"">"&formatnumber(Rs_ordena.Fields("val_original"),2)&"&nbsp;</div>", param_materias																						
				Notas_Tit(linha, 6).AddText "<div align=""right"">"&formatnumber(Rs_ordena.Fields("multa"),2)&"&nbsp;</div>", param_materias		
				Notas_Tit(linha, 7).AddText "<div align=""right"">"&formatnumber(Rs_ordena.Fields("mora"),2)&"&nbsp;</div>", param_materias	 	
				Notas_Tit(linha, 8).AddText "<div align=""right"">"&formatnumber(Rs_ordena.Fields("val_corrigido"),2)&"&nbsp;</div>", param_materias	 	
				Notas_Tit(linha, 9).AddText "<div align=""center"">"&Rs_ordena.Fields("tipo_lancamento")&"</div>", param_materias	 
				Notas_Tit(linha, 10).AddText "<div align=""center"">"&no_curso&"</div>", param_materias		
				Notas_Tit(linha, 11).AddText "<div align=""center"">"&Rs_ordena.Fields("etapa")&"</div>", param_materias			
				Notas_Tit(linha, 12).AddText "<div align=""center"">"&Rs_ordena.Fields("turma")	&"</div>", param_materias	
		
				nome_aluno = Rs_ordena.Fields("nome_aluno")

				y_paginacao=y_paginacao-13					
				total_linhas=total_linhas*1	
				total_linhas=total_linhas+1		

				aluno_acumula_original = aluno_acumula_original + Rs_ordena.Fields("val_original")
				aluno_acumula_multa = aluno_acumula_multa + Rs_ordena.Fields("multa")
				aluno_acumula_mora = aluno_acumula_mora + Rs_ordena.Fields("mora")
				aluno_acumula_corrigido = aluno_acumula_corrigido +  Rs_ordena.Fields("val_corrigido")					
				
				aluno_acumulado_original = formatnumber(aluno_acumula_original,2)
				aluno_acumulado_multa = formatnumber(aluno_acumula_multa,2)
				aluno_acumulado_mora = formatnumber(aluno_acumula_mora,2)
				aluno_acumulado_corrigido = formatnumber(aluno_acumula_corrigido,2)
				
						
				unidade_acumula_original = unidade_acumula_original + Rs_ordena.Fields("val_original")
				unidade_acumula_multa = unidade_acumula_multa + Rs_ordena.Fields("multa")
				unidade_acumula_mora = unidade_acumula_mora + Rs_ordena.Fields("mora")
				unidade_acumula_corrigido = unidade_acumula_corrigido + Rs_ordena.Fields("val_corrigido")			
						
			
	
			Rs_ordena.MOVENEXT
		WEND				
			limite=0
			Paginacao = 0	
					
			if conta_inadimplencia_aluno>1 then	
				linha=linha+1						
				Set Row = Notas_Tit.Rows.Add(13) ' row height	
				altura_unidade = altura_unidade-13					
				y_paginacao=y_paginacao-13				
				Notas_Tit(linha, 1).ColSpan = 3						
				Notas_Tit(linha, 1).AddText "<div align=""center"">Total do Aluno</div>", param_materias
				Notas_Tit(linha, 5).AddText "<div align=""right"">"&aluno_acumulado_original&"&nbsp;</div>", param_materias																						
				Notas_Tit(linha, 6).AddText "<div align=""right"">"&aluno_acumulado_multa&"&nbsp;</div>", param_materias		
				Notas_Tit(linha, 7).AddText "<div align=""right"">"&aluno_acumulado_mora&"&nbsp;</div>", param_materias	 	
				Notas_Tit(linha, 8).AddText "<div align=""right"">"&aluno_acumulado_corrigido&"&nbsp;</div>", param_materias	
			end if								

			unidade_acumulado_original = formatnumber(unidade_acumula_original,2)
			unidade_acumulado_multa = formatnumber(unidade_acumula_multa,2)
			unidade_acumulado_mora = formatnumber(unidade_acumula_mora,2)
			unidade_acumulado_corrigido = formatnumber(unidade_acumula_corrigido,2)	
	
			linha=linha+1						
			Set Row = Notas_Tit.Rows.Add(13) ' row height	
			altura_unidade = altura_unidade-13					
			y_paginacao=y_paginacao-13	
			'param_materias.Add "expand=true" 												
			Notas_Tit(linha, 1).ColSpan = 3						
			Notas_Tit(linha, 1).AddText "<div align=""center"">Total da Unidade "&nome_unidade&"</div>", param_materias
			Notas_Tit(linha, 5).AddText "<div align=""right"">"&unidade_acumulado_original&"&nbsp;</div>", param_materias																						
			Notas_Tit(linha, 6).AddText "<div align=""right"">"&unidade_acumulado_multa&"&nbsp;</div>", param_materias		
			Notas_Tit(linha, 7).AddText "<div align=""right"">"&unidade_acumulado_mora&"&nbsp;</div>", param_materias	 	
			Notas_Tit(linha, 8).AddText "<div align=""right"">"&unidade_acumulado_corrigido&"&nbsp;</div>", param_materias	 		
			
			geral_acumula_original = geral_acumula_original + unidade_acumulado_original
			geral_acumula_multa = geral_acumula_multa + unidade_acumulado_multa
			geral_acumula_mora = geral_acumula_mora + unidade_acumulado_mora
			geral_acumula_corrigido = geral_acumula_corrigido + unidade_acumulado_corrigido					
			
			aluno_acumula_original = 0
			aluno_acumula_multa = 0
			aluno_acumula_mora = 0
			aluno_acumula_corrigido = 0					
			
			unidade_acumula_original = 0
			unidade_acumula_multa = 0
			unidade_acumula_mora = 0
			unidade_acumula_corrigido = 0					
						
				Do While True
					limite=limite+1
					Paginacao = Paginacao+1
				   LastRow = Page.Canvas.DrawTable( Notas_Tit, param_table2 )
		
					if LastRow >= Notas_Tit.Rows.Count Then 
						Exit Do ' entire table displayed
					else
					
						 SET Param_Relatorio = Pdf.CreateParam("x="&margem&";y="&margem&"; height=50; width="&area_utilizavel&"; alignment=left; size=8; color=#000000")
						
						Relatorio = "SWD030 - Sistema Web Diretor"
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


						y_paginacao = 100
						param_table2.Add( "RowTo=2; RowFrom=1" ) ' Row 1 is header.
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
						Text = "<center><i><b><font style=""font-size:18pt;"">Rela&ccedil;&atilde;o de Inadimplentes</font></b></i></center>"
						
						
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
									
						altura_unidade = altura_segundo_separador
				'================================================================================================================		
						altura_titulo = altura_unidade-10
						
						SET Param_per = Pdf.CreateParam("x="&margem&";y="&altura_titulo&"; height=30; width="&width_texto&"; alignment=center; size=12; color=#000000; html=true")
						Text_per = "<left><i><b><font style=""font-size:12pt;"">"&nome_unidade&" - Per&iacute;odo de "&data_inicio&" at&eacute; "&data_fim&"</font></b></i></left>"
						
						
						Do While Len(Text_per) > 0
							CharsPrinted = Page.Canvas.DrawText(Text_per, Param_per, Font )
						 
							If CharsPrinted = Len(Text_per) Then Exit Do
								SET Page = Page.NextPage
							Text_per = Right( Text_per, Len(Text_per) - CharsPrinted)
						Loop 				
					
					end if

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
			
		RS0.MOVENEXT
		WEND	

	
		

		Set param_total = Pdf.CreateParam("width="&area_utilizavel&"; height=20; rows=1; cols=12; border=0; cellborder=0.5; cellspacing=0;")
		Set Table = Doc.CreateTable(param_total)
		Table.Font = Font
		
		largura_colunas=(area_utilizavel-50-210)/colunas_de_notas		
		
		With Table.Rows(1)
		   .Cells(1).Width = 45
		   .Cells(2).Width = 145	
		   .Cells(3).Width = 145
		   .Cells(4).Width = 55			             
		   .Cells(5).Width = 50
		   .Cells(6).Width = 40			             
		   .Cells(7).Width = 40
		   .Cells(8).Width = 50
		   .Cells(9).Width = 125
		   .Cells(10).Width = 30			   			             
		   .Cells(11).Width = 35
		   .Cells(12).Width = 35			             

		End With
		
		geral_acumulado_original = formatnumber(geral_acumula_original,2)
		geral_acumulado_multa = formatnumber(geral_acumula_multa,2)
		geral_acumulado_mora = formatnumber(geral_acumula_mora,2)
		geral_acumulado_corrigido = formatnumber(geral_acumula_corrigido,2)			

		Table(1, 1).ColSpan = 3			
		y_paginacao = y_paginacao-5
		Table(1, 1).AddText "<center><b>Total Geral:</b></center>", "size=8; indenty=4;html=true", Font 
		Table(1, 5).AddText "<div align=""right"">"&geral_acumulado_original&"&nbsp;</div>", "size=8; indenty=4;html=true",  Font																						
		Table(1, 6).AddText "<div align=""right"">"&geral_acumulado_multa&"&nbsp;</div>", "size=8; indenty=4;html=true",  Font		
		Table(1, 7).AddText "<div align=""right"">"&geral_acumulado_mora&"&nbsp;</div>", "size=8; indenty=4;html=true",  Font	 	
		Table(1, 8).AddText "<div align=""right"">"&geral_acumulado_corrigido&"&nbsp;</div>", "size=8; indenty=4;html=true",  Font			
		Page.Canvas.DrawTable Table, "x="&margem&", y="&y_paginacao&"" 	
		y_paginacao = y_paginacao-20			
				

arquivo="SWD030"
Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>
