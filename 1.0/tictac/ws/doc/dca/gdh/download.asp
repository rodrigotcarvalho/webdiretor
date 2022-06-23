<!--#include file="../../../../inc/caminhos.asp"-->
<%
Server.ScriptTimeout = 900 'valor em segundos
nome_arq=request.QueryString("opt")
ano_letivo = session("ano_letivo") 
chave=session("chave")
session("chave")=chave

caminho_completo=caminho_gera_mov&nome_arq

'** instancia o objeto FileUp
Set oFileUp = Server.CreateObject("SoftArtisans.FileUp") 

'** método de abertura do arquivo
Response.ContentType = "application/x-msdownload" 

'** nome que o arquivo terá ao ser baixado, neste caso, tiramos o código.
Response.AddHeader "Content-Disposition", "attachment;filename=""" & nome_arq & """" 

'** obtem o tamanho do arquivo para a barra de progresso do navegador
Set oFM = Server.CreateObject("SoftArtisans.FileManager")
Set oFile = oFM.GetFile(caminho_completo)
Response.AddHeader "Content-Size", oFile.Size 

'** baixa arquivo
oFileUp.TransferFile caminho_completo 

'** destrói objetos
Set oFile = Nothing
Set oFM = Nothing
Set oFileUp = Nothing 

	dim fs
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	'fs.CreateTextFile(caminho_completo,True)
	if fs.FileExists(caminho_completo) then
	  fs.DeleteFile(caminho_completo)
	end if

'response.Redirect("index.asp?nvg="&chave&"&fl="&arquivo&"&opt=ok")

%>
<%If Err.number<>0 then
errnumb = Err.number
errdesc = Err.Description
lsPath = Request.ServerVariables("SCRIPT_NAME")
arPath = Split(lsPath, "/")
GetFileName =arPath(UBound(arPath,1))
passos = 0
for way=0 to UBound(arPath,1)
passos=passos+1
next
seleciona1=passos-2
pasta=arPath(seleciona1)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>