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

if isnumeric(valor) then
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
			If decimo >= 0.75 Then
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
			If decimo >= 0.75 Then
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
end if
End Function

'Função de Busca o Nome da Unidade, Curso, Etapa, Turma, Período e Disciplina
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





'Função de Busca
'===================================================================================================
Function busca_por_nome(query,CONEXAO,tipo_busca)
'tipo_busca: alun=aluno, prof=professor

	'Converte caracteres que não são válidos em uma URL e os transformamem equivalentes para URL
	strProcura = Server.URLEncode(query)
	'Como nossa pesquisa será por "múltiplas palavras" (aqui você pode alterar ao seu gosto)
	'é necessário trocar o sinal de (=) pelo (%) que é usado com o LIKE na string SQL
	strProcura = replace(strProcura,"+"," ")
	strProcura = replace(strProcura,"%27","´")
	strProcura = replace(strProcura,"%27","'")
	strProcura = replace(strProcura,"%C0,","À")
	strProcura = replace(strProcura,"%C1","Á")
	strProcura = replace(strProcura,"%C2","Â")
	strProcura = replace(strProcura,"%C3","Ã")
	strProcura = replace(strProcura,"%C9","É")
	strProcura = replace(strProcura,"%CA","Ê")
	strProcura = replace(strProcura,"%CD","Í")
	strProcura = replace(strProcura,"%D3","Ó")
	strProcura = replace(strProcura,"%D4","Ô")
	strProcura = replace(strProcura,"%D5","Õ")
	strProcura = replace(strProcura,"%DA","Ú")
	strProcura = replace(strProcura,"%DC","Ü")	
	strProcura = replace(strProcura,"%E1","à")
	strProcura = replace(strProcura,"%E1","á")
	strProcura = replace(strProcura,"%E2","â")
	strProcura = replace(strProcura,"%E3","ã")
	strProcura = replace(strProcura,"%E7","ç")
	strProcura = replace(strProcura,"%E9","é")
	strProcura = replace(strProcura,"%EA","ê")
	strProcura = replace(strProcura,"%ED","í")
	strProcura = replace(strProcura,"%F3","ó")
	strProcura = replace(strProcura,"F4","ô")
	strProcura = replace(strProcura,"F5","õ")
	strProcura = replace(strProcura,"%FA","ú")
	strProcura = replace(strProcura,"%FC","ü")

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
	elseif tipo_replace="url" then
		strReplacement = Server.URLEncode(variavel)
		strReplacement = replace(strReplacement,"+"," ")
		strReplacement = replace(strReplacement,"%27","´")
		strReplacement = replace(strReplacement,"%27","'")
		strReplacement = replace(strReplacement,"%C0,","À")
		strReplacement = replace(strReplacement,"%C1","Á")
		strReplacement = replace(strReplacement,"%C2","Â")
		strReplacement = replace(strReplacement,"%C3","Ã")
		strReplacement = replace(strReplacement,"%C9","É")
		strReplacement = replace(strReplacement,"%CA","Ê")
		strReplacement = replace(strReplacement,"%CD","Í")
		strReplacement = replace(strReplacement,"%D3","Ó")
		strReplacement = replace(strReplacement,"%D4","Ô")
		strReplacement = replace(strReplacement,"%D5","Õ")
		strReplacement = replace(strReplacement,"%DA","Ú")
		strReplacement = replace(strReplacement,"%DC","Ü")	
		strReplacement = replace(strReplacement,"%E1","à")
		strReplacement = replace(strReplacement,"%E1","á")
		strReplacement = replace(strReplacement,"%E2","â")
		strReplacement = replace(strReplacement,"%E3","ã")
		strReplacement = replace(strReplacement,"%E7","ç")
		strReplacement = replace(strReplacement,"%E9","é")
		strReplacement = replace(strReplacement,"%EA","ê")
		strReplacement = replace(strReplacement,"%ED","í")
		strReplacement = replace(strReplacement,"%F3","ó")
		strReplacement = replace(strReplacement,"F4","ô")
		strReplacement = replace(strReplacement,"F5","õ")
		strReplacement = replace(strReplacement,"%FA","ú")
		strReplacement = replace(strReplacement,"%FC","ü")
	end if
replace_latin_char=strReplacement
end function	


Function verifica_ano_letivo(variavel_1,variavel_2,variavel_3,variavel_4,variavel_5,CONEXAO,tipo_busca, detalhe_busca)
'tipo_busca
'lib=busca ano letivo em aberto se não existir retorna o maior
'con=consulta se o ano letivo está aberto

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



Function Base64Encode(inData) 
Const Base64 = "1234567890ABCDEFGHIJKLMNOPQRSTUVXYZWabcdefghijklmnopqrstuvxyzw+/"



	Dim cOut, sOut, I 
	
	
	
	'For each group of 3 bytes 
	For I = 1 To Len(inData) Step 3 
		Dim nGroup, pOut, sGroup 
		
		
		
		'Create one long from this 3 bytes. 
		nGroup = &H10000 * Asc(Mid(inData, I, 1)) + _ 
		&H100 * MyASC(Mid(inData, I + 1, 1)) + MyASC(Mid(inData, I + 2, 1)) 
		
		
		
		'Oct splits the long to 8 groups with 3 bits 
		nGroup = Oct(nGroup) 
		
		
		
		'Add leading zeros 
		nGroup = String(8 - Len(nGroup), "0") & nGroup 
		
		
		
		'Convert to base64 
		pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _ 
		Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _ 
		Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _ 
		Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1) 
		
		
		
		'Add the part to output string 
		sOut = sOut + pOut 
		
		
		
		'Add a new line for each 76 chars in dest (76*3/4 = 57) 
		If (I + 2) Mod 57 = 0 Then sOut = sOut + vbCrLf 
	Next 
	
	Select Case Len(inData) Mod 3 
		Case 1: '8 bit final 
			sOut = Left(sOut, Len(sOut) - 2) + "@@" 
		Case 2: '16 bit final 
			sOut = Left(sOut, Len(sOut) - 1) + "@" 
	End Select 
Base64Encode = sOut 
End Function 



Function MyASC(OneChar) 
	If OneChar = "" Then 
		MyASC = 0 
	Else 
		MyASC = Asc(OneChar) 
	end if	
End Function 


Function Base64Decode(Byval base64String) 
	Const Base64 = "1234567890ABCDEFGHIJKLMNOPQRSTUVXYZWabcdefghijklmnopqrstuvxyzw+/" 
	Dim dataLength, sOut, groupBegin 
	
	
	
	'remove white spaces, if any 
	base64String = Replace(base64String, vbCrLf, "") 
	base64String = Replace(base64String, vbTab, "") 
	base64String = Replace(base64String, " ", "") 
	
	
	
	'The source must consists from groups with len of 4 chars 
	dataLength = Len(base64String) 
	If dataLength Mod 4 <> 0 Then 
		Err.Raise 1, "Base64Decode", "Bad Base64 string." 
		Exit Function 
	End If 
	
	
	
	
	' Now decode each group: 
	For groupBegin = 1 To dataLength Step 4 
		Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut 
		' Each data group encodes up To 3 actual bytes. 
		numDataBytes = 3 
		nGroup = 0 
		
		
		
		For CharCounter = 0 To 3 
			' Convert each character into 6 bits of data, And add it To 
			' an integer For temporary storage. If a character is a '=', there 
			' is one fewer data byte. (There can only be a maximum of 2 '=' In 
			' the whole string.) 
			
			
			
			thisChar = Mid(base64String, groupBegin + CharCounter, 1) 
			
			
			
			If thisChar = "@" Then 
			numDataBytes = numDataBytes - 1 
			thisData = 0 
			Else 
			thisData = Instr(Base64, thisChar) - 1 
			End If 
			If thisData = -1 Then 
			Err.Raise 2, "Base64Decode", "Bad character In Base64 string." 
			Exit Function 
			End If 
			
			
			
			nGroup = 64 * nGroup + thisData 
		Next 
		
		
		
		'Hex splits the long to 6 groups with 4 bits 
		nGroup = Hex(nGroup) 
		
		
		
		'Add leading zeros 
		nGroup = String(6 - Len(nGroup), "0") & nGroup 
		
		
		
		'Convert the 3 byte hex integer (6 chars) to 3 characters 
		pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _ 
		Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _ 
		Chr(CByte("&H" & Mid(nGroup, 5, 2))) 
		
		
		
		'add numDataBytes characters to out string 
		sOut = sOut & Left(pOut, numDataBytes) 
	Next 
	
	
	
	Base64Decode = sOut 
End Function 



%>