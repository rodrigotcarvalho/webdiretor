<%@ Language="VBScript" %>
<!-- #include file="aspuploader/include_aspuploader.asp" -->
<%

Dim uploader,mvcfile
Set uploader=new AspUploader
Dim list,i

If Request.Form("myuploader")&""<>"" Then


	
	'Gets the GUID List of the files based on uploader name 
	list=Split(Request.Form("myuploader"),"/")
	
	Response.Write("<div class='tb_subtit'>Arquivos anexados:</div><br>")	
	For i=0 to Ubound(list)
		if i>0 then
			Response.Write("<hr/>")
		end if

		
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

If Request.QueryString("download")&""<>"" Then
	Set mvcfile=uploader.GetUploadedFile(Request.QueryString("download"))
	Response.ContentType="application/oct-stream"
	Response.AddHeader "Content-Disposition","attachment; filename=""" & mvcfile.FileName & """"
	Dim data,stream
	Set stream=CreateObject("ADODB.Stream")
	stream.Mode=3
	stream.Type=1
	stream.Open()
	stream.LoadFromFile(mvcfile.FilePath)
	Dim ws,cs
	ws=0
	while ws<stream.Size
		cs=stream.Size-ws
		If cs>1000 Then
			cs=1000
		End If
		data=stream.Read(cs)
		Response.BinaryWrite(data)
		Response.Flush()
		ws=ws+cs
	wend
	
	stream.Close()
	Response.End()
End If

If Request.Form("delete")&""<>"" Then
	Set mvcfile=uploader.GetUploadedFile(Request.Form("delete"))
	Dim fso
	Set fso=CreateObject("Scripting.FileSystemObject")
	fso.DeleteFile(mvcfile.FilePath)
	Response.Write("OK")
	Response.End()
End If

If Request.Form("guidlist")&""<>"" Then

	list=Split(Request.Form("guidlist"),"/")
	Response.Write("[")
	For i=0 to Ubound(list)
		if i>0 then
			Response.Write(",")
		end if
		Set mvcfile=uploader.GetUploadedFile(list(i))
		Response.Write("{")
		Response.Write("FileGuid:'" & mvcfile.FileGuid & "'")
		Response.Write(",")
		Response.Write("FileSize:'" & mvcfile.FileSize & "'")
		Response.Write(",")
		Response.Write("FileName:'" & mvcfile.FileName & "'")
		Response.Write("}")
	Next
	Response.Write("]")
End If

Response.End()

%>
