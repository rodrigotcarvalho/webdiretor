<%
ordem=request.QueryString("opt")
ano_letivo = request.QueryString("al")
tipo_arquivo=request.QueryString("ta")
nome_arq=request.QueryString("na")
%>

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

%>

