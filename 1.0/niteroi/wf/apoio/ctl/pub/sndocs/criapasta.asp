<%
opt=request.QueryString("opt")
ano_letivo_wf=request.QueryString("al")
prox_pasta=request.QueryString("mp")
ambiente_escola=request.QueryString("env")

%>

<!--#include file="connect_arquivo.asp"-->
<!--#include file="../../../../../inc/caminhos.asp"-->
<%



'cria a pasta
directory= caminho_pasta&prox_pasta
'response.Write(directory)
Set fso = CreateObject("Scripting.FileSystemObject") 
fso.createfolder(directory)

'cria o xml
set FSOx = server.createObject("Scripting.FileSystemObject")
Set Rs_ordena = Server.CreateObject ( "ADODB.RecordSet" )
Set pasta = FSOx.GetFolder(directory)



Set arquivos = pasta.Files
arquivo_xml= prox_pasta&".xml"
set fsoc = createobject("scripting.filesystemobject")
Set act = fsoc.CreateTextFile(server.mappath(arquivo_xml), true)

' cabecalho do XML
act.WriteLine("<?xml version=""1.0"" encoding=""ISO-8859-1""?>")
act.WriteLine("<arquivos>")
act.WriteLine("</arquivos>")
act.close
	
if transicao = "S" then
	area="wd"
	url="http://www.simplynet.com.br/wd/"&ambiente_escola&"/wf/apoio/ctl/pub/novo_tp_doc.asp?opt=ok"
else
	if left(ambiente_escola,5)= "teste" then
		area="wdteste"
		url="http://www.simplynet.com.br/wd/"&ambiente_escola&"/wf/apoio/ctl/pub/novo_tp_doc.asp?opt=ok"	
	else
		area="wd"
		url="http://www.webdiretor.com.br/"&ambiente_escola&"/wf/apoio/ctl/pub/novo_tp_doc.asp?opt=ok"
	end if
end if		
response.redirect(url)
%>
