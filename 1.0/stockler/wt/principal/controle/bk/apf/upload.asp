<!--#include file="../../../../inc/caminhos.asp"-->
<%
ano_letivo=request.QueryString("al")
opt=request.QueryString("opt")
Server.ScriptTimeout = 1800 'valor em segundos
nvg=session("nvg")
session("nvg")=nvg

	Set conexao = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	conexao.Open ABRIR

	Set RSano = Server.CreateObject("ADODB.Recordset")
	SQLano = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&ano_letivo&"'"
	RSano.Open SQLano, conexao		

	situacao_ano=RSano("ST_Ano_Letivo")
	
	if situacao_ano="L" then
		ano_vigente=ano_letivo
	else	
		ano_vigente=DatePart("yyyy", now)
	end if	
				
Set upl = Server.CreateObject("SoftArtisans.FileUp")
upl.Path = CAMINHO_tp
file1 = upl.Form("FILE1").ShortFileName 
arquivo_nome = upl.Form("FILE1").name
    If file1 = "" Then
		response.Redirect("index.asp?nvg="&nvg&"&opt=err1")
 	Elseif file1 <> "POSICAOWEB.txt" and file1 <> "BOLETOWEB.txt" Then
		response.Redirect("index.asp?nvg="&nvg&"&opt=err2")	
	else
		file1 = file1
		upl.Form("FILE1").Save

		response.Redirect("insert.asp?nvg="&nvg&"&opt=a1")	
    End If

									
Set upl = Nothing 	


'Dim Contador, Tamanho
'Dim ConteudoBinario, ConteudoTexto
'Dim Delimitador, Posicao1, Posicao2
'Dim ArquivoNome, ArquivoConteudo, PastaDestino
'Dim objFSO, objArquivo
'
'PastaDestino = CAMINHO_tp
'
''Determina o tamanho do conteúdo
'Tamanho = Request.TotalBytes
'
''Obtém o conteúdo no formato binário
'ConteudoBinario = Request.BinaryRead(Tamanho)
'
''Transforma o conteúdo binário em string
'For Contador = 1 To Tamanho
'  ConteudoTexto = ConteudoTexto & Chr(AscB(MidB(ConteudoBinario, Contador, 1)))
'Next 
'
''Determina o delimitador de campos
'Delimitador = Left(ConteudoTexto, InStr(ConteudoTexto, vbCrLf) - 1)
'
''Percorre a String procurando os campos
''identifica os arquivo e grava no disco
'Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
'
'Posicao1 = InStr(ConteudoTexto, Delimitador) + Len(Delimitar)
'
'do while True
''response.Flush()
''     Response.write ("1 "&Contador&"-"&Posicao1 & "\" &Posicao2&"'<br>")   
'  ArquivoNome = ""
'  Posicao1 = InStr(Posicao1, ConteudoTexto, "filename=")
'  if Posicao1 = 0 then
'    exit do
'  else
'		'Determina o nome do arquivo
'		Posicao1 = Posicao1 + 10
'		Posicao2 = InStr(Posicao1, ConteudoTexto, """")
''		 Response.write ("2 "&Contador&"-"&Posicao1 & "\" &Posicao2&"'<br>")   
'		For contador = (Posicao2 - 1) to Posicao1 step -1
'		if Mid(ConteudoTexto, Contador, 1) <> "\" then '"
'		  ArquivoNome = Mid(ConteudoTexto, Contador, 1) & ArquivoNome
'		else
'		  exit for
'		end if
'		next
'		
'		'Determina o conteúdo do arquivo
'		Posicao1 = InStr(Posicao1, ConteudoTexto, vbCrLf & vbCrLf) + 4
'		Posicao2 = InStr(Posicao1, ConteudoTexto, Delimitador) - 2
'		ArquivoConteudo = Mid(ConteudoTexto, Posicao1, (Posicao2 - Posicao1 + 1))
''		 Response.write ("Arquivo '" & ArquivoNome&"'<br>")
'		'Grava o arquivo
'		if ArquivoNome <> "" then
'			if ArquivoNome = "POSICAOWEB.txt" or ArquivoNome = "BOLETOWEB.txt" then
'				 Set objArquivo = objFSO.CreateTextFile(PastaDestino & "\" & ArquivoNome, true)
'				 objArquivo.WriteLine ArquivoConteudo
'				 objArquivo.Close
'						
'				' Response.write "Arquivo " & PastaDestino & "\" & _
'				 'ArquivoNome & " gravado com sucesso!<br>"
'				 Set objArquivo = nothing
'				response.Redirect("insert.asp?nvg="&nvg&"&opt=a1")						 
'			else
'				response.Redirect("index.asp?nvg="&nvg&"&opt=err2")		
'			end if	
'		else
'			response.Redirect("index.asp?nvg="&nvg&"&opt=err1")					 
'		end if
'	end if
'Loop
'Set objFSO = nothing


'response.Redirect("index.asp?opt=ok")					
%>