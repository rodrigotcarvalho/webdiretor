<%On Error Resume Next%>
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/caminhos.asp"-->

<!--#include file="../inc/caminhos.asp"-->


<%
opt = request.QueryString("opt")
cod= request.QueryString("cod")
ano_letivo = session("ano_letivo")
volta=request.QueryString("volta")


id0 = " > Emitir Lista de Reunião"
id1 = " > Seleciona Unidade"


    	Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2

	%>
<html>
<head>
<title>Web Acad&ecirc;mico</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<% EscreveFuncaoJavaScriptCurso ( CON0 ) %>
<script type="text/javascript" src="../js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_nbGroup(event, grpName) { //v6.0
  var i,img,nbArr,args=MM_nbGroup.arguments;
  if (event == "init" && args.length > 2) {
    if ((img = MM_findObj(args[2])) != null && !img.MM_init) {
      img.MM_init = true; img.MM_up = args[3]; img.MM_dn = img.src;
      if ((nbArr = document[grpName]) == null) nbArr = document[grpName] = new Array();
      nbArr[nbArr.length] = img;
      for (i=4; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
        if (!img.MM_up) img.MM_up = img.src;
        img.src = img.MM_dn = args[i+1];
        nbArr[nbArr.length] = img;
    } }
  } else if (event == "over") {
    document.MM_nbOver = nbArr = new Array();
    for (i=1; i < args.length-1; i+=3) if ((img = MM_findObj(args[i])) != null) {
      if (!img.MM_up) img.MM_up = img.src;
      img.src = (img.MM_dn && args[i+2]) ? args[i+2] : ((args[i+1])? args[i+1] : img.MM_up);
      nbArr[nbArr.length] = img;
    }
  } else if (event == "out" ) {
    for (i=0; i < document.MM_nbOver.length; i++) {
      img = document.MM_nbOver[i]; img.src = (img.MM_dn) ? img.MM_dn : img.MM_up; }
  } else if (event == "down") {
    nbArr = document[grpName];
    if (nbArr)
      for (i=0; i < nbArr.length; i++) { img=nbArr[i]; img.src = img.MM_up; img.MM_dn = 0; }
    document[grpName] = nbArr = new Array();
    for (i=2; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
      if (!img.MM_up) img.MM_up = img.src;
      img.src = img.MM_dn = (args[i+1])? args[i+1] : img.MM_up;
      nbArr[nbArr.length] = img;
  } }
}

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
function submitform()  
{
   var f=document.forms[0]; 
      f.submit(); 
}  
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<table width="1000" border="0" align="center" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td><div align="center"><table width="1000" border="0">
  <tr> 
    <td><font color="#FFFF33" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="../inicio.asp" class="caminho">Web 
      Acad&ecirc;mico</a> 
      <%
	  response.Write(origem&id0&id1)%>
 </font></td>
  </tr>
</table>
<br>
<table width="1000" border="0" cellspacing="0">
  <tr> 
    <td width="219" valign="top">
<table width="100%" border="0" cellspacing="0">
<%if opt = "ok" then%> 
       <tr>
          <td>      
<%
		call mensagens(17,2,0)
	elseif opt = "ok2" then
		call mensagens(20,2,0) 
		
%>
</td>
        </tr>
<% 	end if 
%>      <tr>
          <td>
<%	call mensagens(22,0,0) 
	call ultimo(0) 
%>		  		  
		  </td>
        </tr>
      </table>
      
    </td>
    <td width="785" valign="top"> 	  
      <form name="alteracao" method="post" action="tabelas2.asp?or=01">	  
        <table width="770" border="0" align="right" cellspacing="0" class="tb_corpo"
>
          <tr class="tb_tit"
> 
            <td width="653" height="15" class="tb_tit"
><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Ano 
              Letivo:</strong></font><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%response.Write(ano_letivo)
%>
              </font> </td>
          </tr>
          <tr class="tb_tit"
> 
            <td height="15" bgcolor="#FFFFFF"> </td>
          </tr>
          <tr class="tb_tit"
> 
            <td height="15" class="tb_tit"
><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
              Segmento</strong></font></td>
          </tr>
          <tr> 
            <td><table width="770" border="0" cellspacing="0">
                <tr> 
                  <td width="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                  <td width="194" class="tb_subtit"> <div align="center"><font color="#FF6600"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">UNIDADE 
                      </font></strong></font></div></td>
                  <td width="192" class="tb_subtit"> <div align="center"><font color="#FF6600"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">CURSO 
                      </font></strong></font></div></td>
                  <td width="192" class="tb_subtit"> <div align="center"><font color="#FF6600"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">ETAPA 
                      </font></strong></font></div></td>
                  <td width="192" class="tb_subtit"> <div align="center"><font color="#FF6600"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">TURMA</font></strong></font></div></td>
                </tr>
                <tr> 
                  <td width="10">&nbsp; </td>
                  <td width="194"> <div align="center"> 
                      <select name="unidade" class="borda" onChange="javascript:atualiza_curso(this.form);">
                        <option value="0"></option>
                        <%		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0
While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")
%>
                        <option value="<%response.Write(NU_Unidade)%>"> 
                        <%response.Write(NO_Abr)%>
                        </option>
                        <%RS0.MOVENEXT
WEND
%>
                      </select>
                    </div></td>
                  <td width="192"> <div align="center"> 
                      <select name="curso" class="borda" onChange="MM_callJS('submitform()')">
                        <option value="0"></option>
                      </select>
                    </div></td>
                  <td width="192"> <div align="center"> </div></td>
                  <td width="192"> <div align="center"> </div></td>
                </tr>
                <tr> 
                  <td height="15" colspan="4" bgcolor="#FFFFFF"> </td>
                  <td width="90"> </td>
                </tr>
                <tr> 
                  <td colspan="4" class="tb_tit"
><div align="center"><strong></strong></div></td>
                  <td width="90" class="tb_tit"
>&nbsp;</td>
                </tr>
              </table></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td bgcolor="#FFFFFF">&nbsp; </td>
          </tr>
        </table>
        </form>
	
		</td>
  </tr>
</table>
<p>&nbsp;</p>
<table width="1000" border="0" cellspacing="0">
  <tr>
    <td width="238">&nbsp;</td>
    <td width="770" class="tb_voltar"
><font color="#669999" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="../alunos.asp" target="_parent" class="voltar1">&lt; 
      Voltar para o menu Alunos</a></strong></font></td>
  </tr>
</table>
<p align="center">&nbsp;</p></div></td>
  </tr>
</table>

</body>
<script type="text/javascript">
<!--
  initInputHighlightScript();
//-->
</script>

</html>
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
pasta=arPath(seleciona1)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>