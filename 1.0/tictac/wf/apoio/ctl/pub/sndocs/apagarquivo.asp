<%
tipo_arquivo=request.form("tipo_arquivo")
ano_letivo_wf=request.form("ano_letivo_wf")
ambiente_escola=request.querystring("env")
%>
<!--#include file="connect_arquivo.asp"-->
<%
exclui_doc=request.form("exclui_doc")

vertorExclui = split(exclui_doc,", ")
conta_ocorr=0
for i =0 to ubound(vertorExclui)
exclui_arq=replace(vertorExclui(i),"#$#", ",")
exclui_arq=replace(exclui_arq, "#virgespaco#" , ", ")
co_doc = exclui_arq
		
SET FSO = Server.CreateObject("Scripting.FileSystemObject")
Path = caminho_arquivo

arquivo = Path & co_doc

FSO.deletefile(arquivo) 
next		
response.Redirect("criarquivo.asp?opt=e&al="&ano_letivo_wf&"&tp="&tipo_arquivo&"&env="&ambiente_escola)
%>