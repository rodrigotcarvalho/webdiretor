<%On Error Resume Next%>
<!--#include file="caminhos.asp"-->
<%
chave = request.QueryString("nvg")

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
 call navegacao (CON,chave,4)
navega=Session("caminho")%>
<!--#include file="funcoes.asp"-->
<html>
<head>
<title>Web Acad&ecirc;mico</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="../js/global.js"></script>
<script language="JavaScript" type="text/JavaScript"></script>
<link href="../estilos.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="../img/fundo.gif" onLoad="MM_preloadImages('../img/baner_export_r2_c2_f2.gif','../img/baner_export_r2_c2_f4.gif','../img/baner_export_r2_c2_f3.gif','../img/baner_export_r4_c2_f2.gif','../img/baner_export_r4_c2_f4.gif','../img/baner_export_r4_c2_f3.gif','../img/baner_export_r2_c5_f2.gif','../img/baner_export_r2_c5_f4.gif','../img/baner_export_r2_c5_f3.gif','../img/baner_export_r3_c4_f2.gif','../img/baner_export_r3_c4_f4.gif','../img/baner_export_r3_c4_f3.gif');">
<%call cabecalho(1)
%>
        
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr> 
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> </td>
  </tr>
  <tr> 
    <td height="600" valign="top"> 
      <table width="1000" height="10" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td><div align="center">
              <%call mensagens(1,9700,1,0) %>
            </div></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
<%If Err.number<>0 then
errnumb = Err.number
errdesc = Err.Description
lsPath = Request.ServerVariables("SCRIPT_NAME")
arPath = Split(lsPath, "/")
GetFileName =arPath(UBound(arPath,1))
passos = 0
for way=0 to UBound(arPath,1)
passos=passos+1
next
seleciona1=passos-2
pasta=arPath(seleciona)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("erro.asp")
end if
%>