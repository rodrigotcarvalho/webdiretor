<%

dados = request.form("vencimento")
dados = replace(dados, ", ","$!$")
cod_cons = request.form("cod")

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Web Família</title>
<script language="JavaScript" type="text/JavaScript">
<!--

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
  window.history.go(-1)
}
//-->
</script>
</head>

<body onload="MM_openBrWindow('boleto_itau.asp?c=<%=cod_cons%>&amp;vc=<%=dados%>','','status=yes,scrollbars=yes,resizable=yes,width=800,height=500')">
</body>
</html>
