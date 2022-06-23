<!--#include file="../../../../../inc/caminhos.asp"-->
<%
Server.ScriptTimeout = 1200 'valor em segundos

ordem=request.QueryString("opt")
ano_letivo_wf = request.QueryString("al")
tipo_arquivo=request.QueryString("ta")
arquivo=request.QueryString("na")
%>
<!--#include file="connect_arquivo.asp"-->
<%

caminho_completo=caminho_arquivo&arquivo

'Response.ContentType = "application/unknown"

'Response.AddHeader "Content-Disposition","attachment; filename="&arquivo


'Set objStream = Server.CreateObject("ADODB.Stream")

'objStream.Open

'objStream.Type = 1

'objStream.LoadFromFile caminho_completo

'download = objStream.Read

'Response.BinaryWrite download





Response.buffer = false
 
'** instancia o objeto FileUp
Set oFileUp = Server.CreateObject("SoftArtisans.FileUp") 
 
'** arquivo e caminho completo do arquivo a ser baixado
caminho = caminho_completo
arquivo = arquivo
 
'** método de abertura do arquivo
Response.ContentType = "application/x-msdownload" 
 
'** nome que o arquivo terá ao ser baixado, neste caso, tiramos o código.
Response.AddHeader "Content-Disposition", "attachment;filename=""" & arquivo & """" 
 
'** obtem o tamanho do arquivo para a barra de progresso do navegador
Set oFM = Server.CreateObject("SoftArtisans.FileManager")
Set oFile = oFM.GetFile(caminho)
Response.AddHeader "Content-Size", oFile.Size 

'** baixa arquivo
oFileUp.TransferFile caminho 
 
'** destrói objetos
Set oFile = Nothing
Set oFM = Nothing
Set oFileUp = Nothing 
 

%>

