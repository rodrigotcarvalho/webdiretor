<!--#include file="../../../../inc/caminhos.asp"-->
<%
ano_letivo=request.QueryString("al")
opt=request.QueryString("opt")
chave=session("nvg")
			
Set upl = Server.CreateObject("SoftArtisans.FileUp")
upl.Path = CAMINHO_upload
					
file1 = upl.Form("FILE1").ShortFileName 

lcase_arq=LCase(file1)

Session("arquivo") = file1
					
    If file1 = "" Then
		response.Redirect("index.asp?opt=err1&nvg="&chave)
 	Elseif lcase_arq <> "resultados.xls" Then
		response.Redirect("index.asp?opt=err2&nvg="&chave)	
	else
		file1 = file1
		upl.Form("FILE1").Save
    End If

file1_nom=file1	
 
Session("upl_total") = upl.TotalBytes
Session("ano_letivo") =ano_letivo										
Set upl = Nothing 	

response.Redirect("le_excel.asp")					
%>