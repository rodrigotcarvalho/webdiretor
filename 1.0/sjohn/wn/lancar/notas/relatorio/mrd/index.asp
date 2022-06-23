<%On Error Resume Next%>
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/caminhos.asp"-->

<!--#include file="../inc/caminhos.asp"-->


<%
opt = request.QueryString("opt")
ano_letivo = request.QueryString("ano")
co_prof= session("co_prof")




id0 = " > Mapa de Resultados"


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

if z="2" then	
			
else				

end if

%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Acadêmico</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="file:../img/mm_menu.js"></script>
<%call EscreveFuncaoJavaScriptCurso (CON0,CON1,co_prof) %>
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
function submitano()  
{
   var f=document.forms[0]; 
      f.submit(); 
}
function submitsistema()  
{
   var f=document.forms[1]; 
      f.submit(); 
}
function submitrapido()  
{
   var f=document.forms[2]; 
      f.submit(); 
}  
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
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
//-->
</script>
<link href="../estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../img/menu_r1_c2_f3.gif','../img/menu_r1_c2_f2.gif','../img/menu_r1_c2_f4.gif','../img/menu_r1_c4_f3.gif','../img/menu_r1_c4_f2.gif','../img/menu_r1_c4_f4.gif','../img/menu_r1_c6_f3.gif','../img/menu_r1_c6_f2.gif','../img/menu_r1_c6_f4.gif','../img/menu_r1_c8_f3.gif','../img/menu_r1_c8_f2.gif','../img/menu_r1_c8_f4.gif','../img/menu_direita_r2_c1_f3.gif','../img/menu_direita_r2_c1_f2.gif','../img/menu_direita_r2_c1_f4.gif','../img/menu_direita_r4_c1_f3.gif','../img/menu_direita_r4_c1_f2.gif','../img/menu_direita_r4_c1_f4.gif','../img/menu_direita_r6_c1_f3.gif','../img/menu_direita_r6_c1_f2.gif','../img/menu_direita_r6_c1_f4.gif')">
<% call cabecalho(nivel)
	  %>
<table width="1000" height="670" border="0" align="center" cellspacing="0" bgcolor="#FFFFFF">
  <tr> 
    <td valign="top"> <div align="center"> 
        <table width="1000" border="0" class="tb_caminho">
          <tr> 
            <td><font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="../inicio.asp" class="caminho">Web 
              Notas</a> 
              <%
	  response.Write(origem&id0)%>
              </font></td>
          </tr>
        </table>
        <br>
        <table width="1000" border="0" cellspacing="0">
          <tr> 
            <td width="219" valign="top"> <div align="center"> 
                <table width="100%" border="0" cellspacing="0">
                  <%
if opt = "ok" then%>
                  <tr> 
                    <td> 
                      <%
		call mensagens(17,2,0)
%>
                    </td>
                  </tr>
                  <%
	elseif opt = "ok2" then
%>
                  <tr> 
                    <td> 
                      <%
		call mensagens(20,2,0) 
%>
                    </td>
                  </tr>
                  <%		
	elseif opt = "cln" then
%>
                  <tr> 
                    <td> 
                      <%	
	call mensagens(8,2,0)	
%>
                    </td>
                  </tr>
                  <% 	end if 

 Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Da_Aula Where CO_Professor = "&co_prof
		RS1.Open SQL1, CON1
		
 if RS1.EOF THEN%>
                  <tr> 
                    <td> 
                      <%
		call mensagens(28,1,0)
%>
                    </td>
                  </tr>
                  <%else%>
                  <tr> 
                    <td> 
                      <%	call mensagens(22,0,0) 
	call ultimo(0) 

end if
%>
                    </td>
                  </tr>
                </table>
              </div></td>
            <td width="770" valign="top"> <form name="alteracao" method="post" action="tabelas2.asp?or=01">
                <table width="770" border="0" align="right" cellspacing="0" class="tb_corpo">
                  <tr class="tb_tit"> 
                    <td width="653" height="15" class="tb_tit"><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Ano 
                      Letivo:</strong></font><font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% =ano_letivo%>">
                      <%response.Write(ano_letivo)
%>
                      <input name="co_prof" type="hidden" id="co_prof" value="<% = co_prof %>">
                      </font> </td>
                  </tr>
                  <tr class="tb_tit"> 
                    <td height="15" bgcolor="#FFFFFF"> </td>
                  </tr>
                  <tr class="tb_tit"> 
                    <td height="15" class="tb_tit"><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                      Grade de aulas</strong></font></td>
                  </tr>
                  <tr> 
                    <td><table width="770" border="0" cellspacing="0">
                        <tr> 
                          <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                          <td class="tb_subtit"> <div align="center"><font color="#FF6600"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">UNIDADE 
                              </font></strong></font></div></td>
                          <td class="tb_subtit"> <div align="center"><font color="#FF6600"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">CURSO 
                              </font></strong></font></div></td>
                          <td class="tb_subtit"> <div align="center"><font color="#FF6600"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">ETAPA 
                              </font></strong></font></div></td>
                          <td class="tb_subtit"> <div align="center"><font color="#FF6600"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">TURMA 
                              </font></strong></font></div></td>
                          <td width="107" class="tb_subtit"><div align="center"><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>DISCIPLINA</strong></font></div></td>
                        </tr>
                        <% if RS1.EOF THEN %>
                        <tr> 
                          <td width="7">&nbsp; </td>
                          <td width="200"></td>
                          <td width="56"></td>
                          <td width="143"></td>
                          <td width="103">&nbsp;</td>
                          <td>&nbsp;</td>
                        </tr>
                        <%else%>
                        <tr> 
                          <td width="7">&nbsp; </td>
                          <td width="200"> <div align="center"> 
                              <select name="unidade" class="borda" onChange="javascript:atualiza_curso(this.form);">
                                <option value="0"></option>
                                <%		

NU_Unidade_Check=999999		
u=0					
While not RS1.EOF

NU_Unidade = RS1("NU_Unidade")

if NU_Unidade = NU_Unidade_Check then
RS1.MOVENEXT
else
if u < 2 then
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade Where NU_Unidade= "&NU_Unidade&" order by NO_Abr"
		RS0.Open SQL0, CON0
u=u+1
NO_Abr = RS0("NO_Abr")
%>
                                <option value="<%response.Write(NU_Unidade)%>"> 
                                <%response.Write(NO_Abr)%>
                                </option>
                                <%
else
RS1.MOVENEXT
end if
NU_Unidade_Check = NU_Unidade
RS1.MOVENEXT
end if
WEND
%>
                              </select>
                            </div></td>
                          <td width="56"> <div align="center"> 
                              <select name="curso" class="borda" onChange="MM_callJS('submitform()')">
                                <option value="0"></option>
                              </select>
                            </div></td>
                          <td width="143"> <div align="center"> </div></td>
                          <td width="103">&nbsp;</td>
                          <td>&nbsp;</td>
                        </tr>
                        <%end if %>
                        <tr> 
                          <td height="15" colspan="4" bgcolor="#FFFFFF"> </td>
                          <td width="103"></td>
                          <td width="107"></td>
                        </tr>
                        <tr> 
                          <td colspan="4" class="tb_tit"><div align="center"><strong></strong></div></td>
                          <td width="103" class="tb_tit">&nbsp;</td>
                          <td width="107" class="tb_tit">&nbsp;</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr bgcolor="#FFFFFF"> 
                    <td bgcolor="#FFFFFF">&nbsp; </td>
                  </tr>
                </table>
              </form></td>
          </tr>
        </table>
        <table width="1000" border="0" cellspacing="0">
          <tr> 
            <td width="219">&nbsp;</td>
            <td width="770" class="tb_voltar"><font color="#FF9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="../inicio.asp" class="voltar1">&lt; 
              Voltar para a p&aacute;gina inicial</a></strong></font></td>
          </tr>
        </table>
      </div></td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
</body>
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
response.redirect("../inc/erro.asp")
end if
%>