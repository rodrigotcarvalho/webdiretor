<%
tipo_arquivo=request.form("tipo_arquivo")
ano_letivo=request.form("ano_letivo")
%>

<%
exclui_doc=request.form("exclui_doc")

vertorExclui = split(exclui_doc,", ")
conta_ocorr=0
for i =0 to ubound(vertorExclui)
exclui_arq=replace(vertorExclui(i),"#$#", ",")
co_doc = exclui_arq
		
SET FSO = Server.CreateObject("Scripting.FileSystemObject")
Path = caminho_arquivo

arquivo = Path & co_doc

FSO.deletefile(arquivo) 
next		
response.Redirect("criarquivo.asp?opt=e&al="&ano_letivo&"&tp="&tipo_arquivo)
%>