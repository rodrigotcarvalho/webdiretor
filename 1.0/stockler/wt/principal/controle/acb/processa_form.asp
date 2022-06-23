<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>
<%
concatena_contrato = request.form("num_contrato")
vetor_concatena1 = split(concatena_contrato,", ")
vetor_concatena2 = split(vetor_concatena1(0),"-")
matric=vetor_concatena2(0)

vetor_concatena3 = split(vetor_concatena2(1),"$")
ano_contrato=vetor_concatena3(0)
contrato = vetor_concatena3(1)


 if Request.Form("acao") = "Cancelar" then
	response.redirect("confirma.asp?cc="&concatena_contrato)
  elseif Request.Form("acao") = "Alterar Contrato" then
	response.redirect("alterar_contrato.asp?ac="&ano_contrato&"&nc="&contrato&"&mc="&matric)
  elseif Request.Form("acao") = "Alterar Bolsa" then
	response.redirect("alterar_bolsa.asp?ac="&ano_contrato&"&nc="&contrato&"&mc="&matric)
  end if

%>
<body>
</body>
</html>
