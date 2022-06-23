<%On Error Resume Next%>
<!--#include file="funcoes.asp"-->
<!--#include file="caminhos.asp"-->
<%

opt=request.QueryString("opt")

if opt="al" then

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
ano_info= request.Form("ano_letivo")
ano_info_2 =Split(ano_info, "-")
nivel=ano_info_2(0)
thisPath=ano_info_2(1)
ano_letivo=ano_info_2(2)
session("ano_letivo")=ano_letivo
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&ano_letivo&"'"
		RS0.Open SQL0, CON
		
		sit_an=RS0("ST_Ano_Letivo")
If sit_an="B" then
session("trava")="s"
%>
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
        
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr>
            
    <td height="600" valign="top"> 
      <table width="1000" height="10" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
    <td><div align="center"><%call mensagens(1,9703,0,0) %></div></td>
  </tr>
</table>
    </td>
</tr>
  <tr>
          <td height="40" valign="top"><img src="../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
<%
else
session("trava")="n"
response.redirect ("../inicio.asp")
end if

'response.redirect (thisPath)


call GravaLog ("WR-AUT-AUT-AAL",ano_letivo)

elseif opt="sa" then
sistema=request.form("sistema")
arPath = Split(sistema, "-")
nivel=arPath(0)
chave=arPath(1)
pasta=arPath(2)
session("chave")=chave
session("sistema_local")=chave
response.redirect("../"&pasta)

elseif opt="ar"then
	rapido=request.form("rapido")
	response.redirect("../"&rapido)
end if
	%>
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
response.redirect("../inc/erro.asp")
end if
%>
	