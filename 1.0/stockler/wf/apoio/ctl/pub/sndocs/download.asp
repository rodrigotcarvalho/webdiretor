<!--#include file="../../../../../inc/caminhos.asp"-->
<%
Server.ScriptTimeout = 1200 'valor em segundos

ordem=request.QueryString("opt")
ano_letivo_wf = request.QueryString("al")
tipo_arquivo=request.QueryString("ta")
nome_arq=request.QueryString("na")
%>
<!--#include file="../../../../../inc/caminhos.asp"-->
<%

caminho_completo=caminho_arquivo&arquivo

Response.ContentType = "application/unknown"

Response.AddHeader "Content-Disposition","attachment; filename="&arquivo


Set objStream = Server.CreateObject("ADODB.Stream")

objStream.Open

objStream.Type = 1

objStream.LoadFromFile caminho_completo

download = objStream.Read

Response.BinaryWrite download
%>

