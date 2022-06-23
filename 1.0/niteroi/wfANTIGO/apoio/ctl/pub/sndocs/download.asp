<%
ordem=request.QueryString("opt")
ano_letivo_wf = request.QueryString("al")
session("tipo_arquivo")=request.QueryString("ta")
nome_arq=request.QueryString("na")
%>
 <!--#include file="../../../../../inc/caminhos.asp"-->
<%

caminho_completo=caminho_arquivo&nome_arq
Arquivo=nome_arq

'response.Write(caminho_completo)
'response.End()
'** instancia o objeto FileUp
Set oFileUp = Server.CreateObject("SoftArtisans.FileUp") 

'** arquivo e caminho completo do arquivo a ser baixado
caminho = caminho_completo
arquivo = Arquivo

'** m�todo de abertura do arquivo
Response.ContentType = "application/x-msdownload" 

'** nome que o arquivo ter� ao ser baixado, neste caso, tiramos o c�digo.
Response.AddHeader "Content-Disposition", "attachment;filename=""" & arquivo & """" 

'** obtem o tamanho do arquivo para a barra de progresso do navegador
Set oFM = Server.CreateObject("SoftArtisans.FileManager")
Set oFile = oFM.GetFile(caminho)
Response.AddHeader "Content-Size", oFile.Size 

'** baixa arquivo
oFileUp.TransferFile caminho 

'** destr�i objetos
Set oFile = Nothing
Set oFM = Nothing
Set oFileUp = Nothing 

%> 

%>

