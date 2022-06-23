 <%
opt=request.QueryString("opt")
if opt ="f" then
if session("tipo_arquivo_upl")="0" then

response.Redirect("http://www.simplynet.com.br/bretanha/wf/apoio/ctl/pub/upload.asp?opt=err1")
else
					
					Set upl = Server.CreateObject("SoftArtisans.FileUp")
 upl.Path = "e:\home\bretanha\dados\BD\"&session("ano_letivo")&"\docs\"&session("tipo_arquivo_upl")

					
file1 = upl.Form("FILE1").ShortFileName 
file2 = upl.Form("FILE2").ShortFileName 
file3 = upl.Form("FILE3").ShortFileName 
file4 = upl.Form("FILE4").ShortFileName 
file5 = upl.Form("FILE5").ShortFileName 
contarq=0

    If Not file1 = "" Then
  file1 = file1
  upl.Form("FILE1").Save
contarq=contarq+1  
    Else
	um="no"
    End If
    
	If Not file2 = "" Then
  file2 = file2
  upl.Form("FILE2").Save
contarq=contarq+1    
    Else
	dois="no"
    End If	
	
	    If Not file3 = "" Then
  file3 = file3 
  upl.Form("FILE3").Save
contarq=contarq+1    
    Else
	tres="no"
    End If
	
	    If Not file4 = "" Then
  file4 = file4
  upl.Form("FILE4").Save
contarq=contarq+1    
    Else
	quatro="no"
    End If
	
	    If Not file5 = "" Then
  file5 = file5
  upl.Form("FILE5").Save
contarq=contarq+1    
    Else
	cinco="no"
    End If
if um="no" and dois="no" and tres="no" and quatro="no" and cinco="no" then
response.Redirect("http://www.simplynet.com.br/bretanha/wf/apoio/ctl/pub/upload.asp?opt=err")					
end if
file1_nom=file1	
if contarq>1 and um<>"no" And dois<>"no" then
file2_nom=", "&file2
else
file2_nom=file2
end if
if contarq>1 and (um<>"no" or dois<>"no") And tres<>"no" then
file3_nom=", "&file3
else
file3_nom=file3
end if
if contarq>1 and (um<>"no" or dois<>"no" or tres<>"no") And quatro<>"no" then
file4_nom=", "&file4 
else
file4_nom=file4 
end if
if contarq>1 and (um<>"no" or dois<>"no" or tres<>"no" or quatro<>"no") And cinco<>"no" then
file5_nom=", "&file5
else
file5_nom=file5
end if
		
Session("arquivos") = file1_nom&file2_nom&file3_nom&file4_nom&file5_nom
					
					'upl.Save 
                    Session("upl_total") = upl.TotalBytes
					
response.Redirect("criarquivo.asp?opt=i")					
                    Set upl = Nothing 
					
end if
end if					
%>