<%On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->







<%

nivel=4

opt=request.QueryString("opt")
orig=request.QueryString("or")

autoriza=Session("autoriza")
Session("autoriza")=autoriza
if autoriza="con" or  autoriza="in" or autoriza="ex" then
check_autoriza="con"
elseif autoriza="no" then
response.redirect("../../../../novologin.asp?opt=04")
end if

permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("chave")
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

if opt="gera" then
curso = request.Form("curso")
unidade = request.Form("unidade")
co_etapa = request.Form("etapa")
turma = request.Form("turma")
periodo = request.Form("periodo")
dr= request.Form("dr")
mr= request.Form("mr")
ar= request.Form("ar")
motivo= request.Form("motivo")
obrigatorio=curso&"_"&unidade&"_"&co_etapa&"_"&turma&"_"&periodo&"_"&dr&"_"&mr&"_"&ar
obr=obrigatorio
end if



if co_etapa = "f0"then
co_etapa=0
elseif co_etapa = "f1" or co_etapa = "m1"then
co_etapa=1
elseif co_etapa = "f2" or co_etapa = "m2"then
co_etapa = 2
elseif co_etapa = "f3" or co_etapa = "m3"then
co_etapa = 3
elseif co_etapa = "f4" then
co_etapa = 4
elseif co_etapa = "f5" then
co_etapa = 5
elseif co_etapa = "f6" then
co_etapa = 6
elseif co_etapa = "f7" then
co_etapa = 7
elseif co_etapa = "f8" then
co_etapa = 8
elseif co_etapa = "f55" then
co_etapa = 55
elseif co_etapa = "f66" then
co_etapa = 66
elseif co_etapa = "f77" then
co_etapa = 77
end if

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2
		
		Set CON4 = Server.CreateObject("ADODB.Connection")
		ABRIR4 = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4
		

 call navegacao (CON,chave,nivel)
navega=Session("caminho")

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Abreviado_Curso")


	%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<% EscreveFuncaoJavaScriptCurso ( CON0 ) %>
<script type="text/javascript" src="../../cna/js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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
function submitfuncao()  
{
   var f=document.forms[0]; 
      f.submit(); 
	  
}  function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif"  bgcolor="#FFFFFF">
  <tr>                    
            <td height="10" class="tb_caminho"><font class="style-caminho"> 
              <%
	  response.Write(navega)

%>
              </font>
	</td>
  </tr>
  <%if check_autoriza="con" then%>
  <tr>                   
    <td height="10"> 
      <% call mensagens(4,9701,1,0)%>
    </td>
    </tr> 
<%else%>	
  <%if opt = "ok" then%>
  <tr>                   
    <td height="10"> 
      <% call mensagens(4,602,2,0)%>
    </td>
    </tr>
	<% end if %>				
    <tr>                   
    <td height="10"> 
      <%	call mensagens(4,636,0,0) %>
    </td>
                </tr>
<%end if%>												  				  


          <tr> 
            <td valign="top">
                <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">
                  <tr class="tb_tit"> 
                    
          <td height="15" class="tb_tit">Grade de Aulas 
            <input name="co_grupo" type="hidden" id="co_grupo" value="<% = co_grupo %>"> 
          </td>
                  </tr>
                  <tr> 
                    <td><table width="100%" border="0" cellspacing="0">
              <tr> 
                <td width="8">&nbsp;</td>
                <td width="498" class="tb_subtit"> <div align="center">UNIDADE 
                  </div></td>
                <td width="498" class="tb_subtit"> <div align="center">CURSO </div></td>
                <td width="498" class="tb_subtit"> <div align="center">ETAPA </div></td>
                <td width="498" class="tb_subtit"> <div align="center">TURMA </div></td>
              </tr>
              <tr> 
                <td width="8"> </td>
                <td width="498"> <div align="center"> <font class="form_dado_texto"> 
                    <%response.Write(no_unidade)%>
                    </font></div></td>
                <td width="498"> <div align="center"> <font class="form_dado_texto"> 
                    <%
response.Write(no_curso)%>
                    </font></div></td>
                <td width="498"> <div align="center"> <font class="form_dado_texto"> 
                    <%

		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' and CO_Curso ='"& curso &"'"  
		RS3.Open SQL3, CON0
		
if RS3.EOF THEN
no_etapa="sem etapa"
else
no_etapa=RS3("NO_Etapa")
end if
response.Write(no_etapa)%>
                    </font></div></td>
                <td width="498"> <div align="center"> <font class="form_dado_texto"> 
                    <%response.Write(turma)%>
                    </font></div></td>
              </tr>
            </table></td>
                  </tr>
                  <tr bgcolor="#FFFFFF"> 
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>
                  <tr> 
                    
          <td>
      <form name="inclusao" method="post" action="consulta_turma_cp2.asp?opt=vt&c=<%= curso%>&u=<%= unidade %>" onSubmit="return checksubmit()">	  
		  			
			  <table width="95%" border="0" align="center" cellspacing="0" bgcolor="#FFFFFF">
                <tr> 
                  <td width="17%" class="tb_subtit"> <div align="right">Data da Reuni&atilde;o:</div></td>
                  <td width="83%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%response.Write(dr)%>
                    / 
                    <%response.Write(mr)%>
                    / 
                    <%response.Write(ar)%>
                    </font></td>
                </tr>
                <tr> 
                  <td class="tb_subtit"> <div align="right">Motivo:</div></td>
                  <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%response.Write(motivo)
					Session("motivo")=motivo
					%>
                    </font></td>
                </tr>
                <tr> 
                  <td colspan="2"><div align="center">
                      <input name="Submit" type="submit" class="borda_bot" value="Voltar">
                    </div></td>
                </tr>
              </table>
        </form>
			  </td>
              </tr>
            </table>
					
</td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
<%call GravaLog (chave,obr)%>
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