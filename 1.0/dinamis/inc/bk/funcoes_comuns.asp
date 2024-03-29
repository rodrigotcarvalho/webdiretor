<%
Function calcula_idade(y,m,d)
ano = y
mes = m
dia = d

data= dia&"-"&mes&"-"&ano
intervalo = DateDiff("d", data , now )

calcula_idade = int(intervalo/365.25)

End Function

Function Rpad (sValue, sPadchar, iLength)
  Rpad = sValue & string(iLength - Len(sValue), sPadchar)
End Function

Function Lpad (sValue, sPadchar, iLength)
  Lpad = string(iLength - Len(sValue),sPadchar) & sValue
End Function

Function SpacePad (sValue, sPadchar, iLength, PadSide)
	wrk_tamanho = Len(sValue)
	wrk_string = sValue
	for i = wrk_tamanho to iLength
		if PadSide = "R" then
			wrk_string = wrk_string&sPadchar
		else
			wrk_string = sPadchar&wrk_string	

		end if	
	next
	SpacePad = wrk_string
End Function

Function pad_zeros(p_valor, p_qtd_zeros)
'p_qtd_zeros - N�mero de zeros para prefixar o valor.

  valor = p_valor
  If Len(valor) < p_qtd_zeros Then
    Do Until Len(valor) = p_qtd_zeros
      valor = "0" & valor
    Loop
  End If
  pad_zeros = valor
End Function

Function SomaData(p_data, p_validade)
	diaaDoAno =DatePart("y", p_data) 
	dataExpiracao = diaaDoAno+p_validade
	ano = DatePart("yyyy", now) 	
	if ano mod 4 = 0 then
		diaMax = 366
	else
		diaMax = 355	
	end if
	
	if dataExpiracao > diaMax then
		dataExpiracao = dataExpiracao - diaMax
		ano = ano+1
	end if
	diaExpiracao = DAY(CDate(dataExpiracao))
	diaExpiracao = pad_zeros(diaExpiracao, 2)
	
	mesExpiracao = Month(CDate(dataExpiracao))
	mesExpiracao = pad_zeros(mesExpiracao, 2)

SomaData = diaExpiracao&"/"&mesExpiracao&"/"&ano
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


Function verifica_nome(variavel_1,variavel_2,variavel_3,variavel_4,variavel_5,CONEXAO,tipo_busca, detalhe_busca)
'Busca o Nome da Unidade, Curso, Etapa, Turma, Per�odo e Disciplina
'===================================================================================================
'tipo_busca
'u,c,e,t,p,d
'detalhe_busca
'f=nome completo ou a= nome abreviado
	if tipo_busca="u" then
	
		Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Unidade where NU_Unidade="&variavel_1
		RSu.Open SQLu, CONEXAO
		
		if detalhe_busca="f" then
			verifica_nome=RSu("NO_Sede")
		else
			verifica_nome=RSu("NO_Abr")	
		end if
			
	elseif tipo_busca="c" then
	
		Set RSc = Server.CreateObject("ADODB.Recordset")
		SQLc = "SELECT * FROM TB_Curso where CO_Curso='"&variavel_2&"'"
		RSc.Open SQLc, CONEXAO
		
		if detalhe_busca="f" then
			verifica_nome=RSc("NO_Curso")
		else
			verifica_nome=RSc("NO_Abreviado_Curso")	
		end if
	
	elseif tipo_busca="e" then
		Set RSe = Server.CreateObject("ADODB.Recordset")
		SQLe = "SELECT * FROM TB_Etapa where CO_Curso='"&variavel_2&"' AND CO_Etapa='"&variavel_3&"'"
		RSe.Open SQLe, CONEXAO

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





'Fun��o de Busca
'===================================================================================================
Function busca_por_nome(query,CONEXAO,tipo_busca)
'tipo_busca: alun=aluno, prof=professor

	'Converte caracteres que n�o s�o v�lidos em uma URL e os transformamem equivalentes para URL
	strProcura = Server.URLEncode(query)
	'Como nossa pesquisa ser� por "m�ltiplas palavras" (aqui voc� pode alterar ao seu gosto)
	'� necess�rio trocar o sinal de (=) pelo (%) que � usado com o LIKE na string SQL
	strProcura = replace(strProcura,"+"," ")
	strProcura = replace(strProcura,"%27","�")
	strProcura = replace(strProcura,"%27","'")
	strProcura = replace(strProcura,"%C0,","�")
	strProcura = replace(strProcura,"%C1","�")
	strProcura = replace(strProcura,"%C2","�")
	strProcura = replace(strProcura,"%C3","�")
	strProcura = replace(strProcura,"%C9","�")
	strProcura = replace(strProcura,"%CA","�")
	strProcura = replace(strProcura,"%CD","�")
	strProcura = replace(strProcura,"%D3","�")
	strProcura = replace(strProcura,"%D4","�")
	strProcura = replace(strProcura,"%D5","�")
	strProcura = replace(strProcura,"%DA","�")
	strProcura = replace(strProcura,"%DC","�")	
	strProcura = replace(strProcura,"%E1","�")
	strProcura = replace(strProcura,"%E1","�")
	strProcura = replace(strProcura,"%E2","�")
	strProcura = replace(strProcura,"%E3","�")
	strProcura = replace(strProcura,"%E7","�")
	strProcura = replace(strProcura,"%E9","�")
	strProcura = replace(strProcura,"%EA","�")
	strProcura = replace(strProcura,"%ED","�")
	strProcura = replace(strProcura,"%F3","�")
	strProcura = replace(strProcura,"F4","�")
	strProcura = replace(strProcura,"F5","�")
	strProcura = replace(strProcura,"%FA","�")
	strProcura = replace(strProcura,"%FC","�")

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
		strReplacement = replace(strReplacement,"�,","&Agrave;")
		strReplacement = replace(strReplacement,"�","&Aacute;")
		strReplacement = replace(strReplacement,"�","&Acirc;")
		strReplacement = replace(strReplacement,"�","&Atilde;")
		strReplacement = replace(strReplacement,"�","&Eacute;")
		strReplacement = replace(strReplacement,"�","&Ecirc;")
		strReplacement = replace(strReplacement,"�","&Iacute;")
		strReplacement = replace(strReplacement,"�","&Oacute;")
		strReplacement = replace(strReplacement,"�","&Ocirc;")
		strReplacement = replace(strReplacement,"�","&Otilde;")
		strReplacement = replace(strReplacement,"�","&Uacute;")
		strReplacement = replace(strReplacement,"�","&Uuml;")	
		strReplacement = replace(strReplacement,"�","&agrave;")
		strReplacement = replace(strReplacement,"�","&aacute;")
		strReplacement = replace(strReplacement,"�","&acirc;")
		strReplacement = replace(strReplacement,"�","&atilde;")
		strReplacement = replace(strReplacement,"�","&ccedil;")
		strReplacement = replace(strReplacement,"�","&eacute;")
		strReplacement = replace(strReplacement,"�","&ecirc;")
		strReplacement = replace(strReplacement,"�","&iacute;")
		strReplacement = replace(strReplacement,"�","&oacute;")
		strReplacement = replace(strReplacement,"�","&ocirc;")
		strReplacement = replace(strReplacement,"�","&otilde;")
		strReplacement = replace(strReplacement,"�","&uacute;")
		strReplacement = replace(strReplacement,"�","&uuml;")			
	elseif tipo_replace="url" then
		strReplacement = Server.URLEncode(variavel)
		strReplacement = replace(strReplacement,"+"," ")
		strReplacement = replace(strReplacement,"%27","�")
		strReplacement = replace(strReplacement,"%27","'")
		strReplacement = replace(strReplacement,"%C0,","�")
		strReplacement = replace(strReplacement,"%C1","�")
		strReplacement = replace(strReplacement,"%C2","�")
		strReplacement = replace(strReplacement,"%C3","�")
		strReplacement = replace(strReplacement,"%C9","�")
		strReplacement = replace(strReplacement,"%CA","�")
		strReplacement = replace(strReplacement,"%CD","�")
		strReplacement = replace(strReplacement,"%D3","�")
		strReplacement = replace(strReplacement,"%D4","�")
		strReplacement = replace(strReplacement,"%D5","�")
		strReplacement = replace(strReplacement,"%DA","�")
		strReplacement = replace(strReplacement,"%DC","�")	
		strReplacement = replace(strReplacement,"%E1","�")
		strReplacement = replace(strReplacement,"%E1","�")
		strReplacement = replace(strReplacement,"%E2","�")
		strReplacement = replace(strReplacement,"%E3","�")
		strReplacement = replace(strReplacement,"%E7","�")
		strReplacement = replace(strReplacement,"%E9","�")
		strReplacement = replace(strReplacement,"%EA","�")
		strReplacement = replace(strReplacement,"%ED","�")
		strReplacement = replace(strReplacement,"%F3","�")
		strReplacement = replace(strReplacement,"F4","�")
		strReplacement = replace(strReplacement,"F5","�")
		strReplacement = replace(strReplacement,"%FA","�")
		strReplacement = replace(strReplacement,"%FC","�")
	end if
replace_latin_char=strReplacement
end function	


Function verifica_ano_letivo(variavel_1, variavel_2, CONEXAO, tipo_busca, detalhe_busca)
'tipo_busca
'lib=busca ano letivo em aberto se n�o existir retorna o maior
'situac=consulta se o ano letivo est� aberto

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
elseif tipo_busca="situac" then	

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




' common consts
'Const TypeBinary = 1
TypeBinary = 1
'Const 
ForReading = 1
ForWriting = 2
ForAppending = 8

'private function readBytes(file)
function readBytes(file)
  dim inStream
  ' ADODB stream object used
  set inStream = CreateObject("ADODB.Stream")
  ' open with no arguments makes the stream an empty container
  inStream.Open
  inStream.type= TypeBinary
  inStream.LoadFromFile(file)
  readBytes = inStream.Read()
end function

'private function encodeBase64(bytes)
function encodeBase64(bytes)
  dim DM, EL
  Set DM = CreateObject("Microsoft.XMLDOM")
  ' Create temporary node with Base64 data type
  Set EL = DM.createElement("tmp")
  EL.DataType = "bin.base64"
  ' Set bytes, get encoded String
  EL.NodeTypedValue = bytes
  encodeBase64 = EL.Text
end function

'private function decodeBase64(base64)
function decodeBase64(base64)
  dim DM, EL
  Set DM = CreateObject("Microsoft.XMLDOM")
  ' Create temporary node with Base64 data type
  Set EL = DM.createElement("tmp")
  EL.DataType = "bin.base64"
  ' Set encoded String, get bytes
  EL.Text = base64
  decodeBase64 = EL.NodeTypedValue
end function

'private Sub writeBytes(file, bytes)
Sub writeBytes(file, bytes)
  Dim binaryStream
  Set binaryStream = CreateObject("ADODB.Stream")
  binaryStream.Type = TypeBinary
  'Open the stream and write binary data
  binaryStream.Open
  binaryStream.Write bytes
  'Save binary data to disk
  binaryStream.SaveToFile file, ForWriting
End Sub

Function MyASC(OneChar) 
	If OneChar = "" Then 
		MyASC = 0 
	Else 
		MyASC = Asc(OneChar) 
	end if	
End Function 


'Function Base64Encode(inData) 
'Const Base64 = "1234567890ABCDEFGHIJKLMNOPQRSTUVXYZWabcdefghijklmnopqrstuvxyzw+/"
'
'
'
'	Dim cOut, sOut, I 
'	
'	
'	
'	'For each group of 3 bytes 
'	For I = 1 To Len(inData) Step 3 
'		Dim nGroup, pOut, sGroup 
'		
'		
'		
'		'Create one long from this 3 bytes. 
'		nGroup = &H10000 * Asc(Mid(inData, I, 1)) + _ 
'		&H100 * MyASC(Mid(inData, I + 1, 1)) + MyASC(Mid(inData, I + 2, 1)) 
'		
'		
'		
'		'Oct splits the long to 8 groups with 3 bits 
'		nGroup = Oct(nGroup) 
'		
'		
'		
'		'Add leading zeros 
'		nGroup = String(8 - Len(nGroup), "0") & nGroup 
'		
'		
'		
'		'Convert to base64 
'		pOut = Mid(Base64, CLng("&o" & Mid(nGroup, 1, 2)) + 1, 1) + _ 
'		Mid(Base64, CLng("&o" & Mid(nGroup, 3, 2)) + 1, 1) + _ 
'		Mid(Base64, CLng("&o" & Mid(nGroup, 5, 2)) + 1, 1) + _ 
'		Mid(Base64, CLng("&o" & Mid(nGroup, 7, 2)) + 1, 1) 
'		
'		
'		
'		'Add the part to output string 
'		sOut = sOut + pOut 
'		
'		
'		
'		'Add a new line for each 76 chars in dest (76*3/4 = 57) 
'		If (I + 2) Mod 57 = 0 Then sOut = sOut + vbCrLf 
'	Next 
'	
'	Select Case Len(inData) Mod 3 
'		Case 1: '8 bit final 
'			sOut = Left(sOut, Len(sOut) - 2) + "@@" 
'		Case 2: '16 bit final 
'			sOut = Left(sOut, Len(sOut) - 1) + "@" 
'	End Select 
'Base64Encode = sOut 
'End Function 


'Function Base64Decode(Byval base64String) 
'	Const Base64 = "1234567890ABCDEFGHIJKLMNOPQRSTUVXYZWabcdefghijklmnopqrstuvxyzw+/" 
'	Dim dataLength, sOut, groupBegin 
'	
'	
'	
'	'remove white spaces, if any 
'	base64String = Replace(base64String, vbCrLf, "") 
'	base64String = Replace(base64String, vbTab, "") 
'	base64String = Replace(base64String, " ", "") 
'	
'	
'	
'	'The source must consists from groups with len of 4 chars 
'	dataLength = Len(base64String) 
'	If dataLength Mod 4 <> 0 Then 
'		Err.Raise 1, "Base64Decode", "Bad Base64 string." 
'		Exit Function 
'	End If 
'	
'	
'	
'	
'	' Now decode each group: 
'	For groupBegin = 1 To dataLength Step 4 
'		Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut 
'		' Each data group encodes up To 3 actual bytes. 
'		numDataBytes = 3 
'		nGroup = 0 
'		
'		
'		
'		For CharCounter = 0 To 3 
'			' Convert each character into 6 bits of data, And add it To 
'			' an integer For temporary storage. If a character is a '=', there 
'			' is one fewer data byte. (There can only be a maximum of 2 '=' In 
'			' the whole string.) 
'			
'			
'			
'			thisChar = Mid(base64String, groupBegin + CharCounter, 1) 
'			
'			
'			
'			If thisChar = "@" Then 
'			numDataBytes = numDataBytes - 1 
'			thisData = 0 
'			Else 
'			thisData = Instr(Base64, thisChar) - 1 
'			End If 
'			If thisData = -1 Then 
'			Err.Raise 2, "Base64Decode", "Bad character In Base64 string." 
'			Exit Function 
'			End If 
'			
'			
'			
'			nGroup = 64 * nGroup + thisData 
'		Next 
'		
'		
'		
'		'Hex splits the long to 6 groups with 4 bits 
'		nGroup = Hex(nGroup) 
'		
'		
'		
'		'Add leading zeros 
'		nGroup = String(6 - Len(nGroup), "0") & nGroup 
'		
'		
'		
'		'Convert the 3 byte hex integer (6 chars) to 3 characters 
'		pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _ 
'		Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _ 
'		Chr(CByte("&H" & Mid(nGroup, 5, 2))) 
'		
'		
'		
'		'add numDataBytes characters to out string 
'		sOut = sOut & Left(pOut, numDataBytes) 
'	Next 
'	
'	
'	
'	Base64Decode = sOut 
'End Function 

Function CompactaDB(CaminhoDB, Access97)

Dim fso, Engine, strCaminhoDB
strCaminhoDB = left(CaminhoDB,instrrev(CaminhoDB,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(CaminhoDB) Then
      Set Engine = CreateObject("JRO.JetEngine")
      If Access97 = "S" Then
           Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CaminhoDB, _
           "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strCaminhoDB & "Compacta.mdb;" _
           & "Jet OLEDB:Engine Type=4"
      Else
           Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & CaminhoDB, _
           "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strCaminhoDB & "Compacta.mdb"
      End If
      fso.CopyFile strCaminhoDB & "Compacta.mdb",CaminhoDB
      fso.DeleteFile(strCaminhoDB & "Compacta.mdb")
      Set fso = nothing
      Set Engine = nothing
      'CompactaDB = "Seu banco de dados, " & CaminhoDB & ", foi compactado com sucesso" & vbCrLf
	  CompactaDB ="OK"
Else
	  CompactaDB ="ER"
      'CompactaDB = "O Caminho ou o banco de dados n�o foi localizado.Tente outra vez..." & vbCrLf
End If

End Function

Function Nome_Mes(valor,tipo,ini_maiuscula,outro)

if isnumeric(valor) then
	if tipo = "completo" then
		if valor = 1 then
			if ini_maiuscula = "S" then
				nome_apurado = "Janeiro"
			else
				nome_apurado = "janeiro"		
			end if	
		elseif valor = 2 then
			if ini_maiuscula = "S" then
				nome_apurado = "Fevereiro"
			else
				nome_apurado = "fevereiro"		
			end if			
		elseif valor = 3 then
			if ini_maiuscula = "S" then
				nome_apurado = "Mar&ccedil;o"
			else
				nome_apurado = "mar&ccedil;o"		
			end if			
		elseif valor = 4 then
			if ini_maiuscula = "S" then
				nome_apurado = "Abril"
			else
				nome_apurado = "abril"		
			end if			
		elseif valor = 5 then
			if ini_maiuscula = "S" then
				nome_apurado = "Maio"
			else
				nome_apurado = "maio"		
			end if			
		elseif valor = 6 then
			if ini_maiuscula = "S" then
				nome_apurado = "Junho"
			else
				nome_apurado = "junho"		
			end if			
		elseif valor = 7 then
			if ini_maiuscula = "S" then
				nome_apurado = "Julho"
			else
				nome_apurado = "julho"		
			end if			
		elseif valor = 8 then
			if ini_maiuscula = "S" then
				nome_apurado = "Agosto"
			else
				nome_apurado = "agosto"		
			end if			
		elseif valor = 9 then
			if ini_maiuscula = "S" then
				nome_apurado = "Setembro"
			else
				nome_apurado = "setembro"		
			end if			
		elseif valor = 10 then
			if ini_maiuscula = "S" then
				nome_apurado = "Outubro"
			else
				nome_apurado = "outubro"		
			end if			
		elseif valor = 11 then	
			if ini_maiuscula = "S" then
				nome_apurado = "Novembro"
			else
				nome_apurado = "novembro"		
			end if			
		elseif valor = 12 then	
			if ini_maiuscula = "S" then
				nome_apurado = "Dezembro"
			else
				nome_apurado = "dezembro"		
			end if				
		end if															
	else
		if valor = 1 then
			if ini_maiuscula = "S" then
				nome_apurado = "Jan"
			else
				nome_apurado = "jan"		
			end if	
		elseif valor = 2 then
			if ini_maiuscula = "S" then
				nome_apurado = "Fev"
			else
				nome_apurado = "fev"		
			end if			
		elseif valor = 3 then
			if ini_maiuscula = "S" then
				nome_apurado = "Mar"
			else
				nome_apurado = "mar"		
			end if			
		elseif valor = 4 then
			if ini_maiuscula = "S" then
				nome_apurado = "Abr"
			else
				nome_apurado = "abr"		
			end if			
		elseif valor = 5 then
			if ini_maiuscula = "S" then
				nome_apurado = "Mai"
			else
				nome_apurado = "mai"		
			end if			
		elseif valor = 6 then
			if ini_maiuscula = "S" then
				nome_apurado = "Jun"
			else
				nome_apurado = "jun"		
			end if			
		elseif valor = 7 then
			if ini_maiuscula = "S" then
				nome_apurado = "Jul"
			else
				nome_apurado = "jul"		
			end if			
		elseif valor = 8 then
			if ini_maiuscula = "S" then
				nome_apurado = "Ago"
			else
				nome_apurado = "ago"		
			end if			
		elseif valor = 9 then
			if ini_maiuscula = "S" then
				nome_apurado = "Set"
			else
				nome_apurado = "set"		
			end if			
		elseif valor = 10 then
			if ini_maiuscula = "S" then
				nome_apurado = "Out"
			else
				nome_apurado = "out"		
			end if			
		elseif valor = 11 then	
			if ini_maiuscula = "S" then
				nome_apurado = "Nov"
			else
				nome_apurado = "nov"		
			end if			
		elseif valor = 12 then	
			if ini_maiuscula = "S" then
				nome_apurado = "Dez"
			else
				nome_apurado = "dez"		
			end if	
		end if		
	end if
else
	nome_apurado = "ERRO"
end if
Nome_Mes = nome_apurado
End Function



%>