<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>
<%
curso = request.Form("curso")
unidade = request.Form("unidade")
co_etapa = request.Form("etapa")
turma = request.Form("turma")
periodo = request.Form("periodo")
dr= request.Form("dr")
mr= request.Form("mr")
ar= request.Form("ar")
motivo= request.Form("motivo")
obrigatorio=curso&"_"&unidade&"_"&co_etapa&"_"&turma&"_"&periodo&"_"&dr&"_"&mr&"_"&ar&"_"&motivo
obr=unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&periodo&"_"&ano_letivo
session("obrigatorio")=obrigatorio
%>
<body onLoad="MM_openBrWindow('imprime.asp','','status=yes,menubar=yes,width=1030,height=500,top=50,left=50')">





































































































































































































































































</body>
</html>
<%response.redirect("mapa.asp?or=02&opt=gera")%>

