<%
opt=request.QueryString("opt")
ano_letivo_wf=request.QueryString("al")
prox_pasta=request.QueryString("mp")
ambiente_escola=request.QueryString("env")
if left(ambiente_escola,5)= "teste" then
	area="wdteste"
else
	area="wd"
end if		

%>
<!--#include file="../../../../../inc/caminhos.asp"--> 
<%



'cria a pasta
directory= caminho_pasta&prox_pasta

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
 	url="http://simplynet2.tempsite.ws/"&area&"/"&ambiente_escola&"/wf/apoio/ctl/pub/novo_tp_doc.asp?opt=ok"
else
	url="http://www.simplynet.com.br/"&area&"/"&ambiente_escola&"/wf/apoio/ctl/pub/novo_tp_doc.asp?opt=ok"	
end if			

response.redirect(url)
%>
