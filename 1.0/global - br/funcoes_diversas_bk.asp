<%
Function calcula_idade(y,m,d)
ano = y
mes = m
dia = d

data= dia&"-"&mes&"-"&ano
intervalo = DateDiff("d", data , now )

calcula_idade = int(intervalo/365.25)

End Function

Function arredonda(valor,opcao,qtd_decimais,outro)

if opcao="mat" then
		decimo = valor - Int(valor)
			If decimo >= 0.5 Then
				nota_arredondada = Int(valor) + 1
				valor=nota_arredondada
			else
				nota_arredondada = Int(valor)
				valor=nota_arredondada						
			End If			
		arredonda = formatNumber(valor,qtd_decimais)	
		
		
elseif opcao="mat_dez"	then
		valor=valor*10
		decimo = valor - Int(valor)
			If decimo >= 0.5 Then
				nota_arredondada = Int(valor) + 1
				valor=nota_arredondada
			else
				nota_arredondada = Int(valor)
				valor=nota_arredondada						
			End If	
		valor=valor/10			
		arredonda = formatNumber(valor,qtd_decimais)
		
		
elseif	opcao="quarto" then
		decimo = valor - Int(valor)
		If decimo > 0.5 Then
			nota_arredondada = Int(valor) + 1
			valor=nota_arredondada
		elseIf decimo >= 0.25 Then
			nota_arredondada = Int(valor) + 0.5
			valor=nota_arredondada
		else
			nota_arredondada = Int(valor)
			valor=nota_arredondada											
		End If			
		arredonda = formatNumber(valor,1)	
		
elseif	opcao="quarto_dez" then

		valor=valor*10
		decimo = valor - Int(valor)
		If decimo > 0.5 Then
			nota_arredondada = Int(valor) + 1
			valor=nota_arredondada
		elseIf decimo >= 0.25 Then
			nota_arredondada = Int(valor) + 0.5
			valor=nota_arredondada
		else
			nota_arredondada = Int(valor)
			valor=nota_arredondada											
		End If			
		valor=valor/10
		arredonda = formatNumber(valor,1)	
end if
End Function

'Funзгo de Busca o Nome da Unidade, Curso, Etapa, Turma, Perнodo e Disciplina
'===================================================================================================

Function verifica_nome(variavel_1,variavel_2,variavel_3,variavel_4,variavel_5,CON0,tipo_busca, detalhe_busca)
'tipo_busca
'u,c,e,t,p,d
'detalhe_busca
'f=nome completo ou a= nome abreviado
	if tipo_busca="u" then
	
		Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Unidade where NU_Unidade="&variavel_1
		RSu.Open SQLu, CON0
		
		if detalhe_busca="f" then
			verifica_nome=RSu("NO_Sede")
		else
			verifica_nome=RSu("NO_Abr")	
		end if
			
	elseif tipo_busca="c" then
	
		Set RSc = Server.CreateObject("ADODB.Recordset")
		SQLc = "SELECT * FROM TB_Curso where CO_Curso='"&variavel_2&"'"
		RSc.Open SQLc, CON0
		
		if detalhe_busca="f" then
			verifica_nome=RSc("NO_Curso")
		else
			verifica_nome=RSc("NO_Abreviado_Curso")	
		end if
	
	elseif tipo_busca="e" then
		Set RSe = Server.CreateObject("ADODB.Recordset")
		SQLe = "SELECT * FROM TB_Etapa where CO_Curso='"&variavel_2&"' AND CO_Etapa='"&variavel_3&"'"
		RSe.Open SQLe, CON0

		if detalhe_busca="f" then
			verifica_nome=RSe("NO_Etapa")
		else
			verifica_nome=RSe("NO_Etapa")	
		end if
		
	elseif tipo_busca="t" then
	
	elseif tipo_busca="p" then	

	elseif tipo_busca="d" then

	end if

end function			





'Funзгo de Busca
'===================================================================================================
Function busca_por_nome(query,CONEXAO,tipo_busca)
'tipo_busca: alun=aluno, prof=professor

	'Converte caracteres que nгo sгo vбlidos em uma URL e os transformamem equivalentes para URL
	strProcura = Server.URLEncode(query)
	'Como nossa pesquisa serб por "mъltiplas palavras" (aqui vocк pode alterar ao seu gosto)
	'й necessбrio trocar o sinal de (=) pelo (%) que й usado com o LIKE na string SQL
	strProcura = replace(strProcura,"+"," ")
	strProcura = replace(strProcura,"%27","ґ")
	strProcura = replace(strProcura,"%27","'")
	strProcura = replace(strProcura,"%C0,","А")
	strProcura = replace(strProcura,"%C1","Б")
	strProcura = replace(strProcura,"%C2","В")
	strProcura = replace(strProcura,"%C3","Г")
	strProcura = replace(strProcura,"%C9","Й")
	strProcura = replace(strProcura,"%CA","К")
	strProcura = replace(strProcura,"%CD","Н")
	strProcura = replace(strProcura,"%D3","У")
	strProcura = replace(strProcura,"%D4","Ф")
	strProcura = replace(strProcura,"%D5","Х")
	strProcura = replace(strProcura,"%DA","Ъ")
	strProcura = replace(strProcura,"%DC","Ь")	
	strProcura = replace(strProcura,"%E1","а")
	strProcura = replace(strProcura,"%E1","б")
	strProcura = replace(strProcura,"%E2","в")
	strProcura = replace(strProcura,"%E3","г")
	strProcura = replace(strProcura,"%E7","з")
	strProcura = replace(strProcura,"%E9","й")
	strProcura = replace(strProcura,"%EA","к")
	strProcura = replace(strProcura,"%ED","н")
	strProcura = replace(strProcura,"%F3","у")
	strProcura = replace(strProcura,"F4","ф")
	strProcura = replace(strProcura,"F5","х")
	strProcura = replace(strProcura,"%FA","ъ")
	strProcura = replace(strProcura,"%FC","ь")

IF tipo_busca="alun" THEN
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Alunos where NO_Aluno like '%"& strProcura & "%' order BY NO_Aluno"
	RS.Open SQL, CONEXAO		

	check_aluno=0
	WHile Not RS.EOF
		cod = RS("CO_Matricula")
		if check_aluno=0 then
			vetor_busca=cod		
		ELSE
			vetor_busca=vetor_busca&"#!#"&cod
		END IF
	check_aluno=check_aluno+1
	RS.MOVENEXT
	Wend
ELSEif tipo_busca="prof" THEN

		Set RS = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT * FROM TB_Professor where NO_Professor like '%"& strProcura & "%' order BY NO_Professor"
		RS.Open SQL, CONEXAO

	check_professor=0
	WHile Not RS.EOF
		cod = RS("CO_Professor")
		if check_professor=0 then
			vetor_busca=cod		
		ELSE
			vetor_busca=vetor_busca&"#!#"&cod
		END IF
	check_professor=check_professor+1
	RS.MOVENEXT
	Wend
END IF
busca_por_nome=vetor_busca	
End Function

Function replace_latin_char(variavel,tipo_replace)

	if tipo_replace="html" then
		strReplacement = variavel	
		strReplacement = replace(strReplacement,"А,","&Agrave;")
		strReplacement = replace(strReplacement,"Б","&Aacute;")
		strReplacement = replace(strReplacement,"В","&Acirc;")
		strReplacement = replace(strReplacement,"Г","&Atilde;")
		strReplacement = replace(strReplacement,"Й","&Eacute;")
		strReplacement = replace(strReplacement,"К","&Ecirc;")
		strReplacement = replace(strReplacement,"Н","&Iacute;")
		strReplacement = replace(strReplacement,"У","&Oacute;")
		strReplacement = replace(strReplacement,"Ф","&Ocirc;")
		strReplacement = replace(strReplacement,"Х","&Otilde;")
		strReplacement = replace(strReplacement,"Ъ","&Uacute;")
		strReplacement = replace(strReplacement,"Ь","&Uuml;")	
		strReplacement = replace(strReplacement,"а","&agrave;")
		strReplacement = replace(strReplacement,"б","&aacute;")
		strReplacement = replace(strReplacement,"в","&acirc;")
		strReplacement = replace(strReplacement,"г","&atilde;")
		strReplacement = replace(strReplacement,"з","&ccedil;")
		strReplacement = replace(strReplacement,"й","&eacute;")
		strReplacement = replace(strReplacement,"к","&ecirc;")
		strReplacement = replace(strReplacement,"н","&iacute;")
		strReplacement = replace(strReplacement,"у","&oacute;")
		strReplacement = replace(strReplacement,"ф","&ocirc;")
		strReplacement = replace(strReplacement,"х","&otilde;")
		strReplacement = replace(strReplacement,"ъ","&uacute;")
		strReplacement = replace(strReplacement,"ь","&uuml;")			
	elseif tipo_replace="url" then
		strReplacement = Server.URLEncode(variavel)
		strReplacement = replace(strReplacement,"+"," ")
		strReplacement = replace(strReplacement,"%27","ґ")
		strReplacement = replace(strReplacement,"%27","'")
		strReplacement = replace(strReplacement,"%C0,","А")
		strReplacement = replace(strReplacement,"%C1","Б")
		strReplacement = replace(strReplacement,"%C2","В")
		strReplacement = replace(strReplacement,"%C3","Г")
		strReplacement = replace(strReplacement,"%C9","Й")
		strReplacement = replace(strReplacement,"%CA","К")
		strReplacement = replace(strReplacement,"%CD","Н")
		strReplacement = replace(strReplacement,"%D3","У")
		strReplacement = replace(strReplacement,"%D4","Ф")
		strReplacement = replace(strReplacement,"%D5","Х")
		strReplacement = replace(strReplacement,"%DA","Ъ")
		strReplacement = replace(strReplacement,"%DC","Ь")	
		strReplacement = replace(strReplacement,"%E1","а")
		strReplacement = replace(strReplacement,"%E1","б")
		strReplacement = replace(strReplacement,"%E2","в")
		strReplacement = replace(strReplacement,"%E3","г")
		strReplacement = replace(strReplacement,"%E7","з")
		strReplacement = replace(strReplacement,"%E9","й")
		strReplacement = replace(strReplacement,"%EA","к")
		strReplacement = replace(strReplacement,"%ED","н")
		strReplacement = replace(strReplacement,"%F3","у")
		strReplacement = replace(strReplacement,"F4","ф")
		strReplacement = replace(strReplacement,"F5","х")
		strReplacement = replace(strReplacement,"%FA","ъ")
		strReplacement = replace(strReplacement,"%FC","ь")
	end if
replace_latin_char=strReplacement
end function	


Function verifica_ano_letivo(variavel_1,variavel_2,variavel_3,variavel_4,variavel_5,CONEXAO,tipo_busca, detalhe_busca)
'tipo_busca
'lib=busca ano letivo em aberto se nгo existir retorna o maior
'con=consulta se o ano letivo estб aberto

if tipo_busca="lib" then

	Set RSano = Server.CreateObject("ADODB.Recordset")
	SQLano = "SELECT * FROM TB_Ano_Letivo where ST_Ano_Letivo='L'"
	RSano.Open SQLano, CONEXAO

	if RSano.EOF then
		Set RSano = Server.CreateObject("ADODB.Recordset")
		SQLano = "SELECT MAX(NU_Ano_Letivo) AS ano_letivo FROM TB_Ano_Letivo"
		RSano.Open SQLano, conexao
			
		verifica_ano_letivo=RSano("ano_letivo")
	else
		verifica_ano_letivo=RSano("NU_Ano_Letivo")
	end if
elseif tipo_busca="con" then	

	Set RSano = Server.CreateObject("ADODB.Recordset")
	SQLano = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&variavel_1&"'"
	RSano.Open SQLano, CONEXAO

	if RSano.EOF then
		verifica_ano_letivo="ERR#!#9713"
	else
		verifica_ano_letivo=RSano("ST_Ano_Letivo")
	end if
else	
	verifica_ano_letivo="ERR#!#"	
end if	
end function	

%>