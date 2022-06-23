<html>
 <head>
 <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
 <title>Exemplo Locaweb</title>
 <style type="text/css">
  <!--
 .texto {
 font-family: Tahoma, Verdana, Geneva, Arial, Helvetica, sans-serif;
 font-size: 12px; color: #666666; text-decoration: none;
 }
 h4 {
 font-family: Tahoma, Verdana, Geneva, Arial, Helvetica, sans-serif;
 font-size: 16px; color: #666666; font-weight: bold; text-decoration: none;
 }
 input {
 font-family: Tahoma, Verdana, Geneva, Arial, Helvetica, sans-serif;
 font-size: 10px; color: #666666; font-weight: bold; text-decoration: none;
 background-color: #E8E8E8;
 }
 file {
 font-family: Tahoma, Verdana, Geneva, Arial, Helvetica, sans-serif;
 font-size: 10px; color: #666666; font-weight: bold; text-decoration: none;
 background-color: #E8E8E8;
 }
 textarea {
 font-family: Tahoma, Verdana, Geneva, Arial, Helvetica, sans-serif; 
 font-size: 10px; color: #666666; font-weight: bold; text-decoration: none;
 background-color: #E8E8E8;
 }
 -->
 </style>
 </head>
 <%
 v_situacao = " disabled" 'variavel que habilita os campos do e-mail
 Select Case Request.QueryString("acao") 'Verifica parametro acao para executar determinado script
    Case "upload" 'caso a acao seja upload, executa script do SaFileUp
        Set obj_Upload = Server.CreateObject("SoftArtisans.FileUp")
        obj_Upload.Path = Server.MapPath("./upload") 'local onde será gravado o arquivo
        obj_Upload.Form("File").Save
        Session("arquivo") = obj_Upload.Form("File").ServerName 'recupera o nome do arquivo no servidor
        Response.Write "<script>alert('Total de Bytes Enviados: " & obj_Upload.TotalBytes & "')</script>"
        Set obj_Upload = Nothing
        v_situacao = "" 'habilita os campos pra enviar o e-mail
        v_foco = " onLoad=""document.frm_email.txt_nome_rem.focus();""" 'coloca o cursor no campo do form de e-mail
    Case "email" 'caso a acao seja email, executa script do CDOSYS
        Set objCDOSYSMail = Server.CreateObject("CDO.Message")
        Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration") 'objeto de configuração do CDO
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
        objCDOSYSCon.Fields.update
        Set objCDOSYSMail.Configuration = objCDOSYSCon
        objCDOSYSMail.From = Trim(Request.Form("txt_nome_rem")) & "<" & Trim(Request.Form("txt_email_rem")) & ">"
        objCDOSYSMail.To = Trim(Request.Form("txt_nome_para")) & "<" & Trim(Request.Form("txt_email_para")) & ">"
        objCDOSYSMail.AddAttachment(Session("Arquivo")) 'anexa o arquivo
        objCDOSYSMail.Subject = Request.Form("txt_assunto")
        objCDOSYSMail.TextBody = Request.Form("txt_corpo")
        objCDOSYSMail.Send 'envia o e-mail com o anexo
        Set objCDOSYSMail = Nothing
        Set objCDOSYSCon = Nothing
        Response.Write "<script>alert('E-mail enviado com Sucesso!')</script>"
        Session("Arquivo") = ""
 End Select
 %>
 <body<%=v_foco%>>
 <div align="center" class="texto">
 <h4><strong>SaFileUp + CDOSYS</strong></h4>
 <p>Neste exemplo faremos o upload de um arquivo usando o componente SaFileUp para anexar em 
um e-mail que será enviado pelo componente CDOSYS.</p>

 <form action="" method="post" enctype="multipart/form-data" name="frm_upload" id="frm_upload">
 <table width="500" border="1" cellspacing="0" cellpadding="2">
 <tr><th width="390" scope="col"><div align="center"><input name="file" type="file" size="40"></div></th>
 <th width="96" scope="col"><input name="Upload" type="submit" id="Upload" value="Upload"></th>
 </tr></table></form>

 <form action="" method="post" name="frm_email" id="frm_email">
 <table width="500" border="1" cellspacing="0" cellpadding="2"><tr><th colspan="4">Remetente</th></tr>
 <tr><th width="55">Nome:</th><td width="181"><div align="left">
 <input name="txt_nome_rem" type="text" id="txt_nome_rem" size="30"<%=v_situacao%>>
 </div></td><th width="55">Email:</th><td>
 
 <div align="left">
 <input name="txt_email_rem" type="text" id="txt_email_rem" size="30"<%=v_situacao%>>
 </div></td></tr><tr><th colspan="4">Destinatário</th></tr>
 <tr><th>Nome:</th><td>

 <div align="left">
 <input name="txt_nome_para" type="text" id="txt_nome_para" size="30"<%=v_situacao%>>
 </div></td><th>Email:</th><td>

 <div align="left">
 <input name="txt_email_para" type="text" id="txt_email_para" size="30"<%=v_situacao%>>
 </div></td></tr><tr><th>Assunto:</th><td>

 <div align="left">
 <input name="txt_assunto" type="text" id="txt_assunto" size="30"<%=v_situacao%>>
 </div></td><th>Arquivo:</th><td>

 <div align="left">
 <input name="txt_arquivo" type="text" disabled id="txt_arquivo" value="<%=Session("arquivo")%>" size="30">
 </div></td></tr><tr><th colspan="4">Mensagem</th>
 </tr><tr><td colspan="4">

 <div align="center">
 <textarea name="txt_corpo" cols="75" rows="5" id="txt_corpo"<%=v_situacao%>></textarea>
 </div></td></tr><tr><td colspan="4">

 <div align="right">
 <input name="Enviar" type="submit" id="Enviar" value="Enviar"<%=v_situacao%>>
 </div></td></tr></table></form>
 </div>

 </body>
</html>

