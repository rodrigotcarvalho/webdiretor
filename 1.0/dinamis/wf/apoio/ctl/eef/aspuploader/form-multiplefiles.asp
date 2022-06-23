<%@ Language="VBScript" %>
<!-- #include file="aspuploader/include_aspuploader.asp" -->
<!--#include file="../../../../../inc/caminhos.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<title>
		Form - Multiple files upload
	</title>
	<link href="demo.css" rel="stylesheet" type="text/css" />   
	
	<script type="text/javascript">
	function CuteWebUI_AjaxUploader_OnPostback() {
		//submit the form after the file have been uploaded:
		document.forms[0].submit();
	}
	</script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" /></head>
<body>
	
	<table width="988" border="0" align="right" cellpadding="0" cellspacing="0">
	  <tr>
	    <td><div class="demo">
	  <form id="form1" method="POST">
        <%
			Dim uploader
			Set uploader=new AspUploader
			uploader.MaxSizeKB=10240
			uploader.Name="myuploader"
			uploader.InsertText="Anexar arquivo"
			uploader.MultipleFilesUpload=true
			%>
<%=uploader.GetString() %><br/>
</form>
	  <%

If Request.Form("myuploader")&""<>"" Then

	Dim list,i
	
	'Gets the GUID List of the files based on uploader name 
	list=Split(Request.Form("myuploader"),"/")
	
	Response.Write("<div class='tb_subtit'><br>Arquivos anexados:</div><br>")	
	For i=0 to Ubound(list)
		if i>0 then
			Response.Write("<hr/>")
		end if
		Dim mvcfile
		
		'get the uploaded file based on GUID
		Set mvcfile=uploader.GetUploadedFile(list(i))

		Response.Write("<div class='form_dado_texto'><!--<input name='arquivo_"&i&"' type='checkbox' class='borda' id='dest' value='"&mvcfile.FileName&"' Checked onClick=MM_callJS('desanexa(this.value)') >-->&nbsp;&nbsp;")
		Response.Write(mvcfile.FileName&" ("&mvcfile.FileSize&"Kb)")
		Response.Write("</div>")
		'Copys the uploaded file to a new location.    
        mvcfile.CopyTo(CAMINHO_upload)            
        'Moves the uploaded file to a new location.    
        mvcfile.MoveTo(CAMINHO_upload)   
		
		if i=0 then
			arquivos_anexados=mvcfile.FileName
		else
			arquivos_anexados=arquivos_anexados&"#!#"&mvcfile.FileName
		end if
	Next
	
		Session("arquivos_anexados")=arquivos_anexados

End If

%>
			
	</div></td>
      </tr>
</table>
</body>
</html>
