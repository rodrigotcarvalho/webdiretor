<!--#include file="caminhos.asp"-->
<%
Response.Buffer = True  

ano_letivo=request.QueryString("al")
tipo_arquivo=request.QueryString("tp")
%>



<%Set dir = CreateObject("Scripting.FileSystemObject") 



nome_pasta = caminho_arquivo

set FSO = server.createObject("Scripting.FileSystemObject")

Set Rs_ordena = Server.CreateObject ( "ADODB.RecordSet" )

Set pasta = FSO.GetFolder(nome_pasta)



Set arquivos = pasta.Files

for each arquivo in arquivos
nome_arquivo =arquivo.Name 
Response.cookies("arquivos").item("nome") = nome_arquivo
next

Response.redirect("lecookie.asp"
%>

