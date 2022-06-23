<!--#include file="../../../../inc/caminhos.asp"-->
<%
ano_letivo=request.QueryString("al")
opt=request.QueryString("opt")

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
				
Set upl = Server.CreateObject("SoftArtisans.FileUp")
upl.Path = caminho_bd
					
file1 = upl.Form("FILE1").ShortFileName 

    If file1 = "" Then
		response.Redirect("index.asp?opt=err1")
 	Elseif file1 <> "Posicao.mdb" Then
		response.Redirect("index.asp?opt=err2")	
	else
		file1 = file1
		upl.Form("FILE1").Save
    End If

file1_nom=file1	
		
Session("arquivo") = file1_nom
					
					'upl.Save 
Session("upl_total") = upl.TotalBytes
Session("ano_letivo") =ano_letivo										
Set upl = Nothing 	

response.Redirect("index.asp?opt=ok")					
%>