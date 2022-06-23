<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>
<%
opcao = request.QueryString("opt")
dataAula = request.QueryString("P_DATA_AULA")
if opcao="a" then
	url = "alterar.asp?acao=a&P_DATA_AULA="&dataAula
else
	url = "confirmar.asp"
end if
%>
<body onLoad="javascript:top.location='<%response.Write(url)%>'">
</body>
</html>
