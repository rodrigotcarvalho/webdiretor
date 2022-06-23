<html>
<title></title>
<body bgcolor="#FFFFFF">
<%
opt=request.QueryString("opt")
if opt="i" then
ano_letivo_wf=session("ano_letivo_wf") 
tipo_arquivo=session("tipo_arquivo_upl")
elseif opt="e" then
ano_letivo_wf=request.QueryString("al")
tipo_arquivo=request.QueryString("tp")
end if
%>

<!--#include file="caminhos.asp"-->

<%Set dir = CreateObject("Scripting.FileSystemObject") 



nome_pasta = caminho_arquivo

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
response.write("<a href=1.xml>S</a>")

if opt="i" then
url="http://www.simplynet.com.br/sjohn/wf/apoio/ctl/pub/upload.asp?opt=f&arq="&Session("arquivos")&"&upl="&Session("upl_total")					
else opt="e" then
url="http://www.simplynet.com.br/sjohn/wf/apoio/ctl/pub/docs.asp?opt=ok1&pagina=1&v=s")
else
response.Write("ERRO!")
response.end()					
end if
response.redirect(url)
%>
</body>
</html>