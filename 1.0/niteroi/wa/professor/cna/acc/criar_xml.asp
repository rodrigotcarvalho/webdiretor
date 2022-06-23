<!--#include file="../../../../inc/caminhos.asp"-->
<%
nvg=request.QueryString("nvg")
%>


<%
if left(ambiente_escola,5)="teste" then
	pasta_ambiente="wdteste"
else
	pasta_ambiente="wd"
end if

Set dir = CreateObject("Scripting.FileSystemObject") 


nome_pasta = "e\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\web\"&pasta_ambiente&"\"&ambiente_escola&"\img\fotos\aluno\"
'nome_pasta = "e:\home\simplynet\Web\wdteste\testemraythe\wa\professor\cna\acc\"
'response.write(nome_pasta)
set FSO = server.createObject("Scripting.FileSystemObject")

Set Rs_ordena = Server.CreateObject ( "ADODB.RecordSet" )

Set pasta = FSO.GetFolder(nome_pasta)



Set arquivos = pasta.Files

' Nome do documento XML de saida
arquivo_xml= "alunos.xml"

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

response.redirect("index.asp?nvg="&nvg&"&opt=xml")

%>
