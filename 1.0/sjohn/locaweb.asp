<%On Error Resume Next%>
<!--#include file="inc/caminhos.asp"-->
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>Untitled Document</title>
<link href="estilos.css" rel="stylesheet" type="text/css">
<style type="text/css">
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
</style>
</head>

<body>
<div class="urgente" >
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p align="center"><strong>Estamos tendo problemas com nossos servidores. Por favor, acesse novamente em alguns minutos.</strong></p>
  <p align="center">&nbsp;</p>
  <p align="center">&nbsp;</p>
  <p>&nbsp;</p>
</div>
</body>
</html>
<%
Set objCDOSYSMail = Server.CreateObject("CDO.Message")
        Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration") 'objeto de configuração do CDO
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
        objCDOSYSCon.Fields.update
        Set objCDOSYSMail.Configuration = objCDOSYSCon
        objCDOSYSMail.From = "suportewebdiretorsjohn@simplynetcloud.educacao.ws"
        objCDOSYSMail.To = "osmarpio@openlink.com.br"
        'objCDOSYSMail.AddAttachment(Session("Arquivo")) 'anexa o arquivo
        objCDOSYSMail.Subject = "Erro no servidor"
        objCDOSYSMail.TextBody = "Foi identificado erro de memória no servidor da Locaweb. Favor reiniciar o pool"
        objCDOSYSMail.Send 'envia o e-mail com o anexo
        Set objCDOSYSMail = Nothing
        Set objCDOSYSCon = Nothing
%>