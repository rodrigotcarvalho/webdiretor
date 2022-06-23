<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->

<!--#include file="../../../../inc/funcoes2.asp"-->




<%
opt = request.QueryString("opt")

ano_letivo = request.QueryString("ano")
co_usr = session("co_user")
autoriza = session("autoriza")
session("autoriza")=autoriza
nivel=4

'response.Write(autoriza)
		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1

    	Set CON_WF = Server.CreateObject("ADODB.Connection") 
		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_WF.Open ABRIR_WF


if opt="a" then
nvg = session("chave")
co_usr = session("co_user")
chave=nvg
session("chave")=chave

nivel=4
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)


controle=request.form("wf")		
co_apr1=request.form("apr1")
co_apr2=request.form("apr2")
co_apr3=request.form("apr3")
co_apr4=request.form("apr4")
co_apr5=request.form("apr5")
'co_apr6=request.form("apr6")
co_prova1=request.form("prova1")
co_prova2=request.form("prova2")
co_prova3=request.form("prova3")
co_prova4=request.form("prova4")
co_prova5=request.form("prova5")
'co_prova6=request.form("prova6")

outro=controle&co_apr1&co_prova1&co_apr2&co_prova2&co_apr3&co_prova3&co_apr4&co_prova4
'sql_atualiza= "UPDATE TB_Controle SET CO_controle ='"&controle&"', CO_apr1='"&co_apr1&"', CO_prova1 ='"&co_prova1&"', CO_apr2 ='"&co_apr2&"', CO_prova2 ='"&co_prova2&"', "
'sql_atualiza=sql_atualiza&"CO_apr3='"&co_apr3&"', CO_prova3 ='"& co_prova3 &"', CO_apr4 ='"& co_apr4 &"', CO_prova4='"& co_prova4 &"', CO_apr5 ='"& co_apr5 &"',CO_prova5='"& co_prova5 &"', CO_apr6 ='"& co_apr6 &"',CO_prova6='"& co_prova6 &"'"

sql_atualiza= "UPDATE TB_Controle SET CO_controle ='"&controle&"', CO_apr1='"&co_apr1&"', CO_prova1 ='"&co_prova1&"', CO_apr2 ='"&co_apr2&"', CO_prova2 ='"&co_prova2&"', "
sql_atualiza=sql_atualiza&"CO_apr3='"&co_apr3&"', CO_prova3 ='"& co_prova3 &"', CO_apr4 ='"& co_apr4 &"', CO_prova4='"& co_prova4 &"', CO_apr5 ='"& co_apr5 &"',CO_prova5='"& co_prova5 &"'"
Set RS4 = CON_WF.Execute(sql_atualiza)

else
nvg = request.QueryString("nvg")
chave=nvg
session("chave")=chave
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)
end if	

ano_info=nivel&"-"&chave&"-"&ano_letivo


 call navegacao (CON,chave,nivel)
navega=Session("caminho")	

 Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Autoriz_Usuario_Grupo Where CO_Usuario = "&co_usr
		RS2.Open SQL2, CON
		
if RS2.EOF then

else		
co_grupo=RS2("CO_Grupo")
End if
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="file:../../../../img/mm_menu.js"></script>

<script type="text/javascript" src="../../../../js/global.js"></script>
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
function submitforminterno()  
{
   var f=document.forms[3]; 
      f.submit(); 
	  
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
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../../img/menu_r1_c2_f3.gif','../../../../img/menu_r1_c2_f2.gif','../../../../img/menu_r1_c2_f4.gif','../../../../img/menu_r1_c4_f3.gif','../../../../img/menu_r1_c4_f2.gif','../../../../img/menu_r1_c4_f4.gif','../../../../img/menu_r1_c6_f3.gif','../../../../img/menu_r1_c6_f2.gif','../../../../img/menu_r1_c6_f4.gif','../../../../img/menu_r1_c8_f3.gif','../../../../img/menu_r1_c8_f2.gif','../../../../img/menu_r1_c8_f4.gif','../../../../img/menu_direita_r2_c1_f3.gif','../../../../img/menu_direita_r2_c1_f2.gif','../../../../img/menu_direita_r2_c1_f4.gif','../../../../img/menu_direita_r4_c1_f3.gif','../../../../img/menu_direita_r4_c1_f2.gif','../../../../img/menu_direita_r4_c1_f4.gif','../../../../img/menu_direita_r6_c1_f3.gif','../../../../img/menu_direita_r6_c1_f2.gif','../../../../img/menu_direita_r6_c1_f4.gif')">
<% call cabecalho (nivel)
	  %>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
                    
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
	  </td>
	  </tr>
<%if autoriza=1 then%>
  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(nivel,9701,0,0)
%>
    </td>
                  </tr>
<%end if%> 	  
      <%
if opt = "a" then%>
  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(nivel,9705,2,0)
%>
    </td>
                  </tr>
                  <% 	end if 

%>                  <tr> 
                    
    <td height="10"> 
      <%	call mensagens(nivel,9704,0,0) 
	  
		Set RS_WF = Server.CreateObject("ADODB.Recordset")
		SQL_WF = "SELECT * FROM TB_Controle"
		RS_WF.Open SQL_WF, CON_WF
		
controle=RS_WF("CO_controle")		
co_apr1=RS_WF("CO_apr1")
co_apr2=RS_WF("CO_apr2")
co_apr3=RS_WF("CO_apr3")
co_apr4=RS_WF("CO_apr4")
co_apr5=RS_WF("CO_apr5")
'co_apr6=RS_WF("CO_apr6")
co_prova1=RS_WF("CO_prova1")
co_prova2=RS_WF("CO_prova2")
co_prova3=RS_WF("CO_prova3")
co_prova4=RS_WF("CO_prova4")
co_prova5=RS_WF("CO_prova5")		
'co_prova6=RS_WF("CO_prova6")		

			  
%>
</td></tr>
<tr>

            <td valign="top"> <FORM METHOD="POST" ACTION="index.asp?opt=a">
                
        <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td width="653" height="15" class="tb_tit">Status do Web Fam&iacute;lia
<input name="co_grupo" type="hidden" id="co_grupo" value="<% = co_grupo %>"></td>
          </tr>
          <tr> 
            <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="408"><div align="right"><font class="form_dado_texto">Web 
                      Fam&iacute;lia </font></div></td>
                  <td width="6"><div align="center"><font class="form_dado_texto"> 
                      </font></div></td>
                  <td width="125"> 
                    <% if controle="L" then
				  %>
                    <input name="wf" type="radio"  value="L" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="wf" value="L" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> No AR</font></td>
                  <td width="461"> 
                    <% if controle="D" then%>
                    <input type="radio" name="wf" value="D" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="wf" value="D" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> FORA do AR</font></td>
                </tr>
                <tr> 
                  <td height="15" colspan="4">&nbsp;</td>
                </tr>
                <tr> 
                  <td height="15" colspan="4" class="tb_tit">Status do Aproveitamento 
                    Escolar 
                    <input name="co_grupo" type="hidden" id="co_grupo" value="<% = co_grupo %>"></td>
                </tr>
                <tr> 
                  <td><div align="right"><font class="form_dado_texto">Testes 
                      Per&iacute;odo 1 </font></div></td>
                  <td>&nbsp;</td>
                  <td> 
                    <% if co_apr1="L" then%>
                    <input type="radio" name="apr1" value="L" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="apr1" value="L" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Liberado </font></td>
                  <td> 
                    <% if co_apr1="D" then%>
                    <input type="radio" name="apr1" value="D" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="apr1" value="D" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Bloqueado</font></td>
                </tr>
                <tr> 
                  <td><div align="right"><font class="form_dado_texto">Provas 
                      Per&iacute;odo 1</font></div></td>
                  <td>&nbsp;</td>
                  <td> 
                    <% if co_prova1="L" then%>
                    <input type="radio" name="prova1" value="L" class="borda" checked>	
                    <%else%>
                    <input type="radio" name="prova1" value="L" class="borda" >	
                    <%end if%>
                    <font class="form_dado_texto"> Liberado </font></td>
                  <td> 
                    <% if co_prova1="D" then%>
                    <input type="radio" name="prova1" value="D" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="prova1" value="D" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Bloqueado</font></td>
                </tr>
                <tr> 
                  <td><div align="right"><font class="form_dado_texto">Testes 
                      Per&iacute;odo 2</font></div></td>
                  <td>&nbsp;</td>
                  <td> 
                    <% if co_apr2="L" then%>
                    <input type="radio" name="apr2" value="L" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="apr2" value="L" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Liberado </font></td>
                  <td> 
                    <% if co_apr2="D" then%>
                    <input type="radio" name="apr2" value="D" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="apr2" value="D" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Bloqueado</font></td>
                </tr>
                <tr> 
                  <td><div align="right"><font class="form_dado_texto">Provas 
                      Per&iacute;odo 2</font></div></td>
                  <td>&nbsp;</td>
                  <td> 
                    <% if co_prova2="L" then%>
                    <input type="radio" name="prova2" value="L" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="prova2" value="L" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Liberado </font></td>
                  <td> 
                    <% if co_prova2="D" then%>
                    <input type="radio" name="prova2" value="D" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="prova2" value="D" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Bloqueado</font></td>
                </tr>
                <tr> 
                  <td><div align="right"><font class="form_dado_texto">Testes 
                      Per&iacute;odo 3</font></div></td>
                  <td>&nbsp;</td>
                  <td> 
                    <% if co_apr3="L" then%>
                    <input type="radio" name="apr3" value="L" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="apr3" value="L" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Liberado </font></td>
                  <td> 
                    <% if co_apr3="D" then%>
                    <input type="radio" name="apr3" value="D" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="apr3" value="D" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Bloqueado</font></td>
                </tr>
                <tr> 
                  <td><div align="right"><font class="form_dado_texto">Provas 
                      Per&iacute;odo 3</font></div></td>
                  <td>&nbsp;</td>
                  <td> 
                    <% if co_prova3="L" then%>
                    <input type="radio" name="prova3" value="L" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="prova3" value="L" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Liberado </font></td>
                  <td> 
                    <% if co_prova3="D" then%>
                    <input type="radio" name="prova3" value="D" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="prova3" value="D" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Bloqueado</font></td>
                </tr>
                <tr> 
                  <td><div align="right"><font class="form_dado_texto">Testes 
                      Per&iacute;odo 4</font></div></td>
                  <td>&nbsp;</td>
                  <td> 
                    <% if co_apr4="L" then%>
                    <input type="radio" name="apr4" value="L" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="apr4" value="L" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Liberado </font></td>
                  <td> 
                    <% if co_apr4="D" then%>
                    <input type="radio" name="apr4" value="D" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="apr4" value="D" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Bloqueado</font></td>
                </tr>
                <tr> 
                  <td><div align="right"><font class="form_dado_texto">Provas 
                      Per&iacute;odo 4</font></div></td>
                  <td>&nbsp;</td>
                  <td> 
                    <% if co_prova4="L" then%>
                    <input type="radio" name="prova4" value="L" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="prova4" value="L" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Liberado </font></td>
                  <td> 
                    <% if co_prova4="D" then%>
                    <input type="radio" name="prova4" value="D" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="prova4" value="D" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Bloqueado</font></td>
                </tr>
                <tr> 
                  <td><div align="right"><font class="form_dado_texto">Testes 
                      Per&iacute;odo 5</font></div></td>
                  <td>&nbsp;</td>
                  <td> 
                    <% if co_apr5="L" then%>
                    <input type="radio" name="apr5" value="L" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="apr5" value="L" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Liberado </font></td>
                  <td> 
                    <% if co_apr5="D" then%>
                    <input type="radio" name="apr5" value="D" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="apr5" value="D" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Bloqueado</font></td>
                </tr>
                <tr> 
                  <td><div align="right"><font class="form_dado_texto">Provas 
                      Per&iacute;odo 5</font></div></td>
                  <td>&nbsp;</td>
                  <td> 
                    <% if co_prova5="L" then%>
                    <input type="radio" name="prova5" value="L" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="prova5" value="L" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Liberado </font></td>
                  <td> 
                    <% if co_prova5="D" then%>
                    <input type="radio" name="prova5" value="D" class="borda" checked> 
                    <%else%>
                    <input type="radio" name="prova5" value="D" class="borda" > 
                    <%end if%>
                    <font class="form_dado_texto"> Bloqueado</font></td>
                </tr>
                <tr> 
                  <td colspan="4"><hr></td>
                </tr>
                <tr> 
                  <td colspan="4"><div align="center"> 
                      <%if autoriza=1 then%>
                      <%else%>
                      <INPUT TYPE=SUBMIT VALUE="Alterar" class="botao_prosseguir">
                      <%end if%>
                    </div></td>
                </tr>
              </table>
 </td>
          </tr>
        </table>
              </form></td>
          </tr>
		  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
        </table>

</body>
</html>
<%
if opt="a" then
			call GravaLog (chave,outro)
end if
If Err.number<>0 then
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