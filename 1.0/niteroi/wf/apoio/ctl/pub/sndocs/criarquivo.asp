<html>
<title></title>
<body bgcolor="#FFFFFF">
<!--#include file="../../../../../inc/caminhos.asp"-->
<%
opt=request.QueryString("opt")
ambiente_escola=request.QueryString("env")
if transicao = "S" then
	area="wd"
	link="http://www.simplynet.com.br/"&area&"/"&ambiente_escola
else
	if left(ambiente_escola,5)= "teste" then
		area="wdteste"
		link="http://www.simplynet.com.br/"&area&"/"&ambiente_escola
	else
		area="wd"
		link="http://www.simplynet.com.br/wd/"&ambiente_escola
	end if
end if

if opt="i" then
	ano_letivo_wf=session("ano_letivo_wf")
	tipo_arquivo=session("tipo_arquivo_upl")
elseif opt="e" then
	ano_letivo_wf=request.QueryString("al")
	tipo_arquivo=request.QueryString("tp")
end if

%>

<!--#include file="connect_arquivo.asp"-->

<%Set dir = CreateObject("Scripting.FileSystemObject") 


nome_pasta = caminho_arquivo

response.Write(nome_pasta)

set FSO = server.createObject("Scripting.FileSystemObject")

Set Rs_ordena = Server.CreateObject ( "ADODB.RecordSet" )

Set pasta = FSO.GetFolder(nome_pasta)



Set arquivos = pasta.Files

' Nome do documento XML de saida
arquivo_xml= tipo_arquivo&".xml"

' cria um arquivo usando o file system object
set fso = createobject("scripting.filesystemobject")

' cria o arquivo texto no disco com opção de sobrescrever o arquivo existente
Set act = fso.CreateTextFile(server.mappath(arquivo_xml), true)

' cabecalho do XML
act.WriteLine("<?xml version=""1.0"" encoding=""ISO-8859-1""?>")
act.WriteLine("<arquivos>")

for each arquivo in arquivos
nome_arquivo =arquivo.Name 

act.WriteLine("<nome>" & nome_arquivo & "</nome>" )



next

' fecha a tag 
act.WriteLine("</arquivos>")

' fecha o objeto xml
act.close

' Escreve um link para o arquivo xml criado
'response.write("<a href=1.xml>S</a>")


if opt="i" then
url=link&"/wf/apoio/ctl/pub/upload.asp?opt=f&arq="&Session("arquivos")&"&upl="&Session("upl_total")&"&tp="&tipo_arquivo					
elseif opt="e" then
url=link&"/wf/apoio/ctl/pub/docs.asp?opt=ok1&pagina=1&v=s"
else
response.Write("ERRO!")
response.end()					
end if
response.redirect(url)
%>
</body>
</html>