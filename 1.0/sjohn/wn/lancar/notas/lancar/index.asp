<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->



<%
opt = request.QueryString("opt")
ori	= request.QueryString("ori")
if isnull(ori) or ori="" then
ori=0
end if

ano_letivo = session("ano_letivo")
co_usr = session("co_user")
nivel=4
if opt="cln" then
nvg=session("chave")
chave=nvg
else
nvg = request.QueryString("nvg")
chave=nvg
session("chave")=chave
end if
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano_info=nivel&"-"&chave&"-"&ano_letivo

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1

    	Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2= "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2	


 call navegacao (CON,chave,nivel)
navega=Session("caminho")	

 Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Professor Where CO_Usuario = "&co_usr
		RS2.Open SQL2, CON2
		
if RS2.EOF then
Response.Write("Usuário não é Professor!")
else		
co_prof=RS2("CO_Professor")
End if
Session("co_prof")=co_prof
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
function submitfuncao()  
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
                         <script>
function createXMLHTTP()
            {
                        try
                        {
                                   ajax = new ActiveXObject("Microsoft.XMLHTTP");
                        }
                        catch(e)
                        {
                                   try
                                   {
                                               ajax = new ActiveXObject("Msxml2.XMLHTTP");
                                               alert(ajax);
                                   }
                                   catch(ex)
                                   {
                                               try
                                               {
                                                           ajax = new XMLHttpRequest();
                                               }
                                               catch(exc)
                                               {
                                                            alert("Esse browser não tem recursos para uso do Ajax");
                                                            ajax = null;
                                               }
                                   }
                                   return ajax;
                        }
           
           
               var arrSignatures = ["MSXML2.XMLHTTP.5.0", "MSXML2.XMLHTTP.4.0",
               "MSXML2.XMLHTTP.3.0", "MSXML2.XMLHTTP",
               "Microsoft.XMLHTTP"];
               for (var i=0; i < arrSignatures.length; i++) {
                                                                          try {
                                                                                                             var oRequest = new ActiveXObject(arrSignatures[i]);
                                                                                                             return oRequest;
                                                                          } catch (oError) {
                                                                          }
                                      }
           
                                      throw new Error("MSXML is not installed on your system.");
                        }                                
						
						
						 function recuperarCurso(uTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa_wn.asp?opt=c", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
document.all.divEtapa.innerHTML ="<select class=select_style></select>"
document.all.divTurma.innerHTML = "<select class=select_style></select>"
document.all.divDisc.innerHTML = "<select class=select_style></select>"
document.all.divPeriodo.innerHTML = "<select class=select_style></select>"
//recuperarEtapa()
                                                           }
                                               }

                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa_wn.asp?opt=e", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select class=select_style></select>"
document.all.divDisc.innerHTML = "<select class=select_style></select>"
document.all.divPeriodo.innerHTML = "<select class=select_style></select>"
//recuperarTurma()
                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa_wn.asp?opt=t3", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t
document.all.divDisc.innerHTML = "<select class=select_style></select>"
document.all.divPeriodo.innerHTML = "<select class=select_style></select>"
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
function recuperarDisciplina(eTipo,co_prof)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa_wn.asp?opt=d2", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_d= oHTTPRequest.responseText;
resultado_d = resultado_d.replace(/\+/g," ")
resultado_d = unescape(resultado_d)
document.all.divDisc.innerHTML = resultado_d
document.all.divPeriodo.innerHTML = "<select class=select_style></select>"
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo + "&pr_pub=" +co_prof);
                                   }
function recuperarPeriodo(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa_wn.asp?opt=p", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_p= oHTTPRequest.responseText;
resultado_p = resultado_p.replace(/\+/g," ")
resultado_p = unescape(resultado_p)
document.all.divPeriodo.innerHTML = resultado_p
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }								   								   

 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
                        </script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../../img/menu_r1_c2_f3.gif','../../../../img/menu_r1_c2_f2.gif','../../../../img/menu_r1_c2_f4.gif','../../../../img/menu_r1_c4_f3.gif','../../../../img/menu_r1_c4_f2.gif','../../../../img/menu_r1_c4_f4.gif','../../../../img/menu_r1_c6_f3.gif','../../../../img/menu_r1_c6_f2.gif','../../../../img/menu_r1_c6_f4.gif','../../../../img/menu_r1_c8_f3.gif','../../../../img/menu_r1_c8_f2.gif','../../../../img/menu_r1_c8_f4.gif','../../../../img/menu_direita_r2_c1_f3.gif','../../../../img/menu_direita_r2_c1_f2.gif','../../../../img/menu_direita_r2_c1_f4.gif','../../../../img/menu_direita_r4_c1_f3.gif','../../../../img/menu_direita_r4_c1_f2.gif','../../../../img/menu_direita_r4_c1_f4.gif','../../../../img/menu_direita_r6_c1_f3.gif','../../../../img/menu_direita_r6_c1_f2.gif','../../../../img/menu_direita_r6_c1_f4.gif')">
<% call cabecalho (nivel)
	  %>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
                    
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
	  </td>
	  </tr>
      <%
if opt = "ok" then%>
  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(17,2,0)
%>
    </td>
                  </tr>
                  <%
	elseif opt = "ok2" then
%>
                  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(20,2,0) 
%>
    </td>
                  </tr>
                  <%		
	elseif opt = "cln" then
%>
                  <tr> 
                    
    <td height="10"> 
      <%	
	call mensagens(nivel,623,2,0)	
%>
    </td>
                  </tr>
                  <% 	end if 

 Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS1.Open SQL1, CON0
		
 if RS1.EOF THEN%>
                  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(4,619,1,0)
%>
    </td>
                  </tr>
                  <%else%>
                  <tr> 
                    
    <td height="10"> 
      <%	call mensagens(4,1,0,0) 
	  
	  
	call ultimo(0) 
%>
</td></tr>
<%

end if
%><tr>

            <td valign="top"> <form name="alteracao" method="post" action="notas.asp">
                
        <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td width="653" height="15" class="tb_tit">Grade de Aulas</td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="16%" class="tb_subtit"> <div align="center">UNIDADE 
                    </div></td>
                  <td width="16%" class="tb_subtit"> <div align="center">CURSO 
                    </div></td>
                  <td width="16%" class="tb_subtit"> <div align="center">ETAPA 
                    </div></td>
                  <td width="16%" class="tb_subtit"> <div align="center">TURMA 
                    </div></td>
                  <td class="tb_subtit"> <div align="center">DISCIPLINA</div></td>
                  <td class="tb_subtit"> <div align="center">PER&Iacute;ODO</div></td>
                </tr>
                <% 'if RS1.EOF THEN %>
                <%'else%>
                <tr>
                  <td> <div align="center"> 
                     <select name="unidade" class="select_style" onchange="recuperarCurso(this.value)">
                        <option value="0"></option> 
                        <%		
		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0
		
While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")

		Set RSUn_Prof = Server.CreateObject("ADODB.Recordset")
		SQLUn_Prof = "SELECT * FROM TB_Da_Aula Where CO_Professor="&co_prof&" AND NU_Unidade="& NU_Unidade &" ORDER BY NU_Unidade"
		RSUn_Prof.Open SQLUn_Prof, CON1

if RSUn_Prof.eof then		
RS0.MOVENEXT
else		
	if NU_Unidade =session("unidades") and ori=1 or ori="1" then
	%>
							<option value="<%response.Write(NU_Unidade)%>" selected> 
							<%response.Write(NO_Abr)%>
							</option> 
	<%		
	else
	%>
							<option value="<%response.Write(NU_Unidade)%>"> 
							<%response.Write(NO_Abr)%>
							</option>
							<%
	end if
RS0.MOVENEXT
end if
WEND
%>
                      </select>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divCurso"> 
					  <%if ori<>1 or ori<>"1" then%>
                        <select class="select_style">
                        </select>
					<%else
					%>
                     <select name="curso" class="select_style" onchange="recuperarEtapa(this.value)">
                        <option value="999990" selected></option> 
                        <%		

								Set RS0 = Server.CreateObject("ADODB.Recordset")
								SQLUn_Prof = "SELECT DISTINCT CO_Curso FROM TB_Da_Aula Where CO_Professor="&co_prof&" AND NU_Unidade="& session("unidades")
								RS0.Open SQLUn_Prof, CON1
								
								
						While not RS0.EOF
						CO_Curso = RS0("CO_Curso")
						
								Set RS0a = Server.CreateObject("ADODB.Recordset")
								SQL0a = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
								RS0a.Open SQL0a, CON0
								
						NO_Curso = RS0a("NO_Abreviado_Curso")
						CO_Curso=CO_Curso*1
						session("grau") =session("grau") *1
							if CO_Curso =session("grau") then
														%>
								<option value="<%response.Write(CO_Curso)%>" selected> 
								<%response.Write(NO_Curso)%>
								</option>
							<%
							else		
								%>
								<option value="<%response.Write(CO_Curso)%>"> 
								<%response.Write(NO_Curso)%>
								</option>
							<%end if
						RS0.MOVENEXT
						WEND
						%>
                      </select>					
					<%end if%>	
                      </div>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divEtapa"> 
					  <%if ori<>1 or ori<>"1" then%>
                        <select class="select_style">
                        </select>
						<%else%>
                     <select name="etapa" class="select_style" onchange="recuperarTurma(this.value)">
                        <option value="999990" selected></option>
                        <%		
		
		
										Set RS0b = Server.CreateObject("ADODB.Recordset")
										SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Da_Aula Where CO_Professor="&co_prof&" AND NU_Unidade="& session("unidades") &" AND CO_Curso='"&session("grau")&"'"
										RS0b.Open SQL0b, CON1
										
											
								While not RS0b.EOF
								CO_Etapa = RS0b("CO_Etapa")
								
										Set RS0c = Server.CreateObject("ADODB.Recordset")
										SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("grau")&"' AND CO_Etapa='"&CO_Etapa&"'"
										RS0c.Open SQL0c, CON0
										
								NO_Etapa = RS0c("NO_Etapa")
								
								if CO_Etapa =session("serie") then%>
                        <option value="<%response.Write(CO_Etapa)%>" selected> 
                        <%response.Write(NO_Etapa)%>
                        </option>
					<%										
								else%>
                        <option value="<%response.Write(CO_Etapa)%>"> 
                        <%response.Write(NO_Etapa)%>
                        </option>
					<%end if
					RS0b.MOVENEXT
					WEND
					%>
                      </select>						
						<%end if%>	
                      </div>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divTurma"> 
					  <%if ori<>1 or ori<>"1" then%>
                        <select class="select_style">
                        </select>
						<%else%>
						<select name="turma" class="select_style" onChange="MM_callJS('recuperarDisciplina()')">
												<option value="999990" selected></option>
							<%
								Set RS3 = Server.CreateObject("ADODB.Recordset")
								SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Da_Aula where CO_Professor="&co_prof&" AND NU_Unidade="&session("unidades")&"AND CO_Curso='"&session("grau")&"' AND CO_Etapa='" & session("serie") & "' order by CO_Turma" 
								RS3.Open SQL3, CON1						
						
						while not RS3.EOF
						co_turma= RS3("CO_Turma")
						 if co_turma =session("turma") then%>
						<option value="<%=co_turma%>" selected> 
							<%response.Write(co_turma)%>
						</option>
							<%						
						else
						 %>
						<option value="<%=co_turma%>"> 
							<%response.Write(co_turma)%>
						</option>
							<%end if
						RS3.MOVENEXT
						WEND
						%>
							
						 </select>						
						<%end if%>
                      </div>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divDisc"> 
					  <%if ori<>1 or ori<>"1" then%>
                        <select class="select_style">
                        </select>
						<%else
							Set RSG = Server.CreateObject("ADODB.Recordset")
							SQLG = "SELECT DISTINCT CO_Materia_Principal FROM TB_Da_Aula where CO_Professor ="& co_prof&" and CO_Etapa = '"&session("serie") &"' AND NU_Unidade = "&session("unidades")&" and CO_Curso = '"&session("grau") &"' order by CO_Materia_Principal"
							RSG.Open SQLG, CON1
							
					IF RSG.EOF THEN
					
					RESPONSE.Write("Sem disciplinas cadastradas. Procure seu Coordenador.")
					
					ELSE
					%>
										  <select name="mat_prin" class="select_style" onChange="MM_callJS('recuperarPeriodo()')">
											<option value="999999" selected></option>
											<%
					while not RSG.EOF
					co_mat_prin= RSG("CO_Materia_Principal")
					
							Set RS7 = Server.CreateObject("ADODB.Recordset")
							SQL7 = "SELECT * FROM TB_Materia where CO_Materia ='"& co_mat_prin &"'"
							RS7.Open SQL7, CON0
							
							no_mat_prin= RS7("NO_Materia")
								if co_mat_prin=session("co_materia") then	%>
								<option value="<%=co_mat_prin%>" selected> 
								<%response.Write(no_mat_prin)%>						
								</option>
								<%else%>
								<option value="<%=co_mat_prin%>"> 
								<%response.Write(no_mat_prin)%>						
								</option>
								<%end if
					
					RSG.MOVENEXT
					WEND
					END IF
					%>
                      </select>						
						<%end if%>
                      </div>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divPeriodo"> 
					  <%if ori<>1 or ori<>"1" then%>
                        <select class="select_style">
                        </select>
						<%else%>
							<select name="periodo" class="select_style" id="periodo" onChange="MM_callJS('submitfuncao()')">
							  <option value="0" selected></option>
							  <%
							Set RS4 = Server.CreateObject("ADODB.Recordset")
							SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
							RS4.Open SQL4, CON0
							
							while not RS4.EOF
							NU_Periodo =  RS4("NU_Periodo")
							NO_Periodo= RS4("NO_Periodo")
							NU_Periodo=NU_Periodo*1
							session("periodo")=session("periodo")*1
							if NU_Periodo =session("periodo") then%>
							  <option value="<%=NU_Periodo%>" selected> 
							  <%response.Write(NO_Periodo)%>
							  </option>
							  <%else%>
							  <option value="<%=NU_Periodo%>"> 
							  <%response.Write(NO_Periodo)%>
							  </option>
							  <%
							end if
							RS4.MOVENEXT
							WEND%>
                                    </select>						
						<%end if%>
                      </div>
                    </div></td>
                </tr>
                <%'end if %>
                <tr> 
                  <td width="16%" height="15" bgcolor="#FFFFFF"></td>
                  <td width="16%" height="15" bgcolor="#FFFFFF"></td>
                  <td width="16%" height="15" bgcolor="#FFFFFF"></td>
                  <td width="16%" height="15" bgcolor="#FFFFFF"></td>
                  <td width="18%" bgcolor="#FFFFFF"></td>
                  <td width="16%" bgcolor="#FFFFFF"></td>
                </tr>
              </table></td>
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
'response.redirect("../../../../inc/erro.asp")
end if
%>