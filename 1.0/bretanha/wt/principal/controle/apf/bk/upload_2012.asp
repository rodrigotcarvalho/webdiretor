<!--#include file="../../../../inc/caminhos.asp"-->
<%
ano_letivo=request.QueryString("al")
opt=request.QueryString("opt")
Server.ScriptTimeout = 1800 'valor em segundos
nvg=session("nvg")
session("nvg")=nvg

	Set conexao = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	conexao.Open ABRIR

	Set RSano = Server.CreateObject("ADODB.Recordset")
	SQLano = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&ano_letivo&"'"
	RSano.Open SQLano, conexao		

	situacao_ano=RSano("ST_Ano_Letivo")
	
	if situacao_ano="L" then
		ano_vigente=ano_letivo
	else	
		ano_vigente=DatePart("yyyy", now)
	end if	
				
Set Upload = Server.CreateObject("Persits.Upload")
' we use memory uploads, so we must limit file size
Upload.SetMaxSize 10000000, True '10mb

' Save to memory. Path parameter is omitted
Upload.Save

' Check whether a file was selected
Set File = Upload.Files("FILE1")
If Not File Is Nothing Then
   ' Obtain file name
   Filename = file.Filename
   
   if Filename <> "POSICAOWEB.txt" and Filename <> "BOLETOWEB.txt" Then
		response.Redirect("index.asp?nvg="&nvg&"&opt=err2")	
   else
   ' check if file exists in c:\upload under this name
   'If Upload.FileExists("c:\upload\" & filename ) Then
   '   Response.Write "File with this name already exists."
   'Else
      ' otherwise save file
      File.SaveAs CAMINHO_tp & File.Filename
   '   Response.Write "File saved as " & File.Path
   'End If
   		if Filename = "POSICAOWEB.txt" then
			response.Redirect("insert.asp?nvg="&nvg&"&opt=a1")	
		else
			response.Redirect("insert.asp?nvg="&nvg&"&opt=a2")			
		end if	
	end if		
Else ' file not selected
   response.Redirect("index.asp?nvg="&nvg&"&opt=err1")
End If								
%>