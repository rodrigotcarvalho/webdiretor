<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->

<!--#include file="../../../../inc/caminhos.asp"-->
<% 
session("nvg")=""
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
nvg=request.QueryString("nvg")
opt = request.QueryString("opt")
ori = request.QueryString("ori")
recria_at = SESSION("recria_at")
recria_att = SESSION("recria_att")

chave=nvg
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
if opt="" or isnull("opt") then
opt="sel"
else
opt=opt
if opt="ok" then
cod_cons = request.QueryString("cod_cons")
co_usr_prof = request.QueryString("co_usr_prof")
end if
end if

unidade_pesquisa=SESSION("unidade_pesquisa")
curso_pesquisa=SESSION("curso_pesquisa")
etapa_pesquisa=SESSION("etapa_pesquisa")
turma_pesquisa=SESSION("turma_pesquisa")

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
 call VerificaAcesso (CON,chave,nivel)
autoriza=Session("autoriza")

 call navegacao (CON,chave,nivel)
navega=Session("caminho")

%>
<html>
<head>
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../../../js/mm_menu.js"></script>
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

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}

function MM_showHideLayers() { //v6.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
    obj.visibility=v; }
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
<script>
<!--
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
                                               oHTTPRequest.open("post", "executa.asp?ori=alt&opt=c", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
document.all.divEtapa.innerHTML ="<select name=etapa class=borda><option value=999990 selected></option></select>"
document.all.divTurma.innerHTML = "<select name=turma class=borda><option value=999990 selected></option></select>"
                                                           }
                                               }
                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?ori=alt&opt=e", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select name=turma class=borda><option value=999990 selected></option></select>"

                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?ori=alt&opt=t", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t																	   
                                                           }
                                               }
                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }  
						 function recuperarAT(eTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?ori=tr", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.at.innerHTML = resultado_t																	   
                                                           }
                                               }
                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   } 
						 function recuperarATT(eTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?ori=tr", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.at.innerHTML = ""
document.all.mt.innerHTML = ""																	   
                                                           }
                                               }
                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   } 								   
						 function recuperarAT(eTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?ori=tr", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.at.innerHTML = resultado_t
document.all.mt.innerHTML = ""																	   
                                                           }
                                               }
                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   } 
						 function recuperarMT(eTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?ori=tr", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.at.innerHTML = ""
document.all.mt.innerHTML = resultado_t																	   
                                                           }
                                               }
                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
function checksubmit()
{
 var el = document.forms[3].elements;
 for(var i = 0 ; i < el.length ; ++i) {
  if(el[i].type == "radio") {
   var radiogroup = el[el[i].name]; // get the whole set of radio buttons.
   var itemchecked = false;
   for(var j = 0 ; j < radiogroup.length ; ++j) {
    if(radiogroup[j].checked) {
	 itemchecked = true;
	 break;
	}
   }
   if(!itemchecked) { 
 //   alert("Por favor selecione uma opção de busca para o campo "+el[i].name+".");
	 alert("Por favor selecione uma opção de busca.");
    if(el[i].focus)
     el[i].focus();
	return false;
   }
  }
 }
  return true
}
function disableForm1 (f) {
 
document.enturma.recria_att.disabled  = false;
document.enturma.recria_at.disabled  = true;
}
function disableForm2 (f) {
 
document.enturma.recria_att.disabled  = true;
document.enturma.recria_at.disabled  = false;
}
function disableForm3 (f) {
 
document.enturma.recria_att.disabled  = true;
document.enturma.recria_at.disabled  = true;
}
function addLoadEvent(func) {
  var oldonload = window.onload;
  if (typeof window.onload != 'function') {
    window.onload = func;
  } else {
    window.onload = function() {
      if (oldonload) {
        oldonload();
      }
      func();
    }
  }
}

addLoadEvent(function() {
<%if opt="ok" then %>
	disableForm1 ('s');
<% elseif opt="ok2" then %>
	disableForm2 ('s');
<%elseif opt="ok3" or ori=2 then %>
	disableForm3 ('s');
<%end if%>	
});
//-->
</script>
</head>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
            
    <td height="10" colspan="5" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
    </td>
          </tr>
<%if opt="ok" or opt="ok2" or opt="ok3" then%>
		            <tr> 
            
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,419,2,0) %>
    </td>
          </tr>
<%end if%>		  
		            <tr> 
            
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,414,0,0) %>
    </td>
          </tr>

                <tr class="tb_corpo"> 
                  
    <td height="10" colspan="5" class="tb_tit">Escolha o tipo de Enturma&ccedil;&atilde;o</td>
                </tr>
                <tr> 
                  
    <td colspan="5" valign="top"><form name="enturma" id="enturma" method="post" action="altera.asp" onSubmit="return checksubmit()">
        <table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="2%"> 
              <%if opt="ok2" or opt="ok3" then%>
              <input type="radio" name="enturma" value="att" onClick="recuperarATT(this.value);disableForm1('s')" > 
              <%else%>
              <input type="radio" name="enturma" value="att" onClick="recuperarATT(this.value);disableForm1('s')" checked>	
              <%end if%>
            </td>
            <td width="25%"><font class="form_dado_texto">Automaticamente para 
              todas as Turmas</font></td>
            <td width="2%"> 
              <%if opt="ok" and recria_att="s"then%>
              <input name="recria_att" type="Checkbox" id="recria_att" value="s" checked> 
              <%else%>
              <input name="recria_att" type="Checkbox" id="recria_att" value="s">	
              <%end if%>
            </td>
            <td width="71%"><font class="form_dado_texto">Recriar o n&uacute;mero 
              da chamada</font></td>
          </tr>
          <tr> 
            <td> 
              <%if opt="ok2" then%>
              <input type="radio" name="enturma" value="at" onClick="recuperarAT(this.value);disableForm2('s')" checked> 
              <%else%>
              <input type="radio" name="enturma" value="at" onClick="recuperarAT(this.value);disableForm2('s')">	
              <%end if%>
            </td>
            <td><font class="form_dado_texto">Automaticamente por Turma</font></td>
            <td> 
              <%if opt="ok2" and recria_at="s"then%>
              <input name="recria_at" type="Checkbox" id="recria_at" value="s" checked > 
              <%else%>
              <input name="recria_at" type="Checkbox" id="recria_at" value="s">	
              <%end if%>
            </td>
            <td><font class="form_dado_texto">Recriar o n&uacute;mero da chamada</font></td>
          </tr>
          <tr> 
            <td></td>
            <td height="0" colspan="3"> <div id="at"> 
                <%if opt="ok2" then%>
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td width="25%"> <div align="center"><font class="form_dado_texto">Unidade</font> 
                      </div></td>
                    <td width="25%"> <div align="center"><font class="form_dado_texto">Curso</font> 
                      </div></td>
                    <td width="25%"> <div align="center"><font class="form_dado_texto">Etapa</font> 
                      </div></td>
                    <td width="25%"> <div align="center"><font class="form_dado_texto">Turma 
                        </font> </div></td>
                  </tr>
                  <tr valign="top"> 
                    <td width="25%" height="10"> <div align="center"> 
                        <select name="unidade" class="borda" onchange="recuperarCurso(this.value)">
                          <%if unidade_pesquisa="999990" then%>
                          <option value="999990" selected></option>
                          <%end if		

		Set RS0u = Server.CreateObject("ADODB.Recordset")
		SQL0u = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0u.Open SQL0u, CON0
NU_Unidade_Check=999999		
While not RS0u.EOF
NU_Unidade = RS0u("NU_Unidade")
NO_Abr = RS0u("NO_Abr")
if NU_Unidade = NU_Unidade_Check then
RS0u.MOVENEXT		
else
unidade_pesquisa=unidade_pesquisa*1
NU_Unidade=NU_Unidade*1
if NU_Unidade = unidade_pesquisa then
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
NU_Unidade_Check = NU_Unidade
RS0u.MOVENEXT
end if
WEND
%>
                        </select>
                      </div></td>
                    <td width="25%" height="10" align="left"> <div align="center" id="divCurso"> 
                        <select name="curso" class="borda" onchange="recuperarEtapa(this.value)">
                          <%if curso_pesquisa="999990" then%>
                          <option value="999990" selected></option>
                          <%end if
		Set RS0ue = Server.CreateObject("ADODB.Recordset")
		SQL0ue = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade_pesquisa
		RS0ue.Open SQL0ue, CON0
		
	
CO_Curso_check="999999"		
While not RS0ue.EOF
CO_Curso = RS0ue("CO_Curso")

if CO_Curso = CO_Curso_check then
RS0ue.MOVENEXT		
else

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
		RS0c.Open SQL0c, CON0
		
NO_Curso = RS0c("NO_Curso")		

if CO_Curso = curso_pesquisa then
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
                          <%						
end if

CO_Curso_check = CO_Curso
RS0ue.MOVENEXT
end if
WEND
%>
                        </select>
                      </div></td>
                    <td width="25%" height="10" align="center"> <div align="center" id="divEtapa"> 
                        <select name="etapa" class="borda" onchange="recuperarTurma(this.value)">
                          <%if etapa_pesquisa="999990" then%>
                          <option value="999990" selected></option>
                          <%end if						

		Set RS0e = Server.CreateObject("ADODB.Recordset")
		SQL0e = "SELECT * FROM TB_Unidade_Possui_Etapas where CO_Curso ='"& curso_pesquisa &"' AND NU_Unidade="& unidade_pesquisa 
		RS0e.Open SQL0e, CON0
			

while not RS0e.EOF
co_etapa= RS0e("CO_Etapa")

		Set RS3e = Server.CreateObject("ADODB.Recordset")
		SQL3e = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' And CO_Curso ='"& curso_pesquisa &"'" 
		RS3e.Open SQL3e, CON0
		

no_etapa=RS3e("NO_Etapa")

if co_etapa = etapa_pesquisa then
%>
                          <option value="<%response.Write(co_etapa)%>" selected> 
                          <%response.Write(no_etapa)%>
                          </option>
                          <%
else
%>
                          <option value="<%=co_etapa%>"> 
                          <%response.Write(no_etapa)%>
                          </option>
                          <%						
end if
RS0e.MOVENEXT
WEND
%>
                        </select>
                      </div></td>
                    <td width="25%" height="10" align="center"> <div align="center" id="divTurma"> 
                        <select name="turma" class="borda">
                          <%if turma_pesquisa="999990" then%>
                          <option value="999990" selected></option>
                          <%end if						
	
		Set RS0t = Server.CreateObject("ADODB.Recordset")
		SQL0t = "SELECT * FROM TB_Turma where NU_Unidade="&unidade_pesquisa&"AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='" & etapa_pesquisa & "' order by CO_Turma" 
		RS0t.Open SQL0t, CON0						
co_turma_check=9999990
while not RS0t.EOF
co_turma= RS0t("CO_Turma")

if co_turma = co_turma_check then
RS0t.MOVENEXT
else

if co_turma = turma_pesquisa then
capacidade= RS0t("NU_Capacidade")
%>
                          <option value="<%response.Write(co_turma)%>" selected> 
                          <%response.Write(co_turma)%>
                          </option>
                          <%
else
%>
                          <option value="<%=co_turma%>"> 
                          <%response.Write(co_turma)%>
                          </option>
                          <%						
end if

co_turma_check = co_turma
RS0t.MOVENEXT
end if
WEND
%>
                        </select>
                      </div></td>
                  </tr>
                </table>
                <%end if%>
              </div></td>
          </tr>
          <tr> 
            <td> 
              <%if opt="ok3" or ori=2 then%>
              <input type="radio" name="enturma" value="mt" onClick="recuperarMT(this.value);disableForm3('s')" checked> 
              <%else%>
              <input type="radio" name="enturma" value="mt" onClick="recuperarMT(this.value);disableForm3('s')">	
              <%end if%>
            </td>
            <td><font class="form_dado_texto">Manualmente por Turma</font></td>
            <td colspan="2"><font class="form_dado_texto">&nbsp; </font></td>
          </tr>
          <tr> 
            <td height="0"></td>
            <td height="0" colspan="3"> <div id="mt"> 
                <%if opt="ok3" or ori=2 then%>
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td width="25%"> <div align="center"><font class="form_dado_texto">Unidade</font> 
                      </div></td>
                    <td width="25%"> <div align="center"><font class="form_dado_texto">Curso</font> 
                      </div></td>
                    <td width="25%"> <div align="center"><font class="form_dado_texto">Etapa</font> 
                      </div></td>
                    <td width="25%"> <div align="center"><font class="form_dado_texto">Turma 
                        </font> </div></td>
                  </tr>
                  <tr valign="top"> 
                    <td width="25%" height="10"> <div align="center"> 
                        <select name="unidade" class="borda" onchange="recuperarCurso(this.value)">
                          <%if unidade_pesquisa="999990" then%>
                          <option value="999990" selected></option>
                          <%end if		

		Set RS0u = Server.CreateObject("ADODB.Recordset")
		SQL0u = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0u.Open SQL0u, CON0
NU_Unidade_Check=999999		
While not RS0u.EOF
NU_Unidade = RS0u("NU_Unidade")
NO_Abr = RS0u("NO_Abr")
if NU_Unidade = NU_Unidade_Check then
RS0u.MOVENEXT		
else
unidade_pesquisa=unidade_pesquisa*1
NU_Unidade=NU_Unidade*1
if NU_Unidade = unidade_pesquisa then
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
NU_Unidade_Check = NU_Unidade
RS0u.MOVENEXT
end if
WEND
%>
                        </select>
                      </div></td>
                    <td width="25%" height="10" align="left"> <div align="center" id="divCurso"> 
                        <select name="curso" class="borda" onchange="recuperarEtapa(this.value)">
                          <%if curso_pesquisa="999990" then%>
                          <option value="999990" selected></option>
                          <%end if
		Set RS0ue = Server.CreateObject("ADODB.Recordset")
		SQL0ue = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade_pesquisa
		RS0ue.Open SQL0ue, CON0
		
	
CO_Curso_check="999999"		
While not RS0ue.EOF
CO_Curso = RS0ue("CO_Curso")

if CO_Curso = CO_Curso_check then
RS0ue.MOVENEXT		
else

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
		RS0c.Open SQL0c, CON0
		
NO_Curso = RS0c("NO_Curso")		

if CO_Curso = curso_pesquisa then
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
                          <%						
end if

CO_Curso_check = CO_Curso
RS0ue.MOVENEXT
end if
WEND
%>
                        </select>
                      </div></td>
                    <td width="25%" height="10" align="center"> <div align="center" id="divEtapa"> 
                        <select name="etapa" class="borda" onchange="recuperarTurma(this.value)">
                          <%if etapa_pesquisa="999990" then%>
                          <option value="999990" selected></option>
                          <%end if						

		Set RS0e = Server.CreateObject("ADODB.Recordset")
		SQL0e = "SELECT * FROM TB_Unidade_Possui_Etapas where CO_Curso ='"& curso_pesquisa &"' AND NU_Unidade="& unidade_pesquisa 
		RS0e.Open SQL0e, CON0
			

while not RS0e.EOF
co_etapa= RS0e("CO_Etapa")

		Set RS3e = Server.CreateObject("ADODB.Recordset")
		SQL3e = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' And CO_Curso ='"& curso_pesquisa &"'" 
		RS3e.Open SQL3e, CON0
		

no_etapa=RS3e("NO_Etapa")

if co_etapa = etapa_pesquisa then
%>
                          <option value="<%response.Write(co_etapa)%>" selected> 
                          <%response.Write(no_etapa)%>
                          </option>
                          <%
else
%>
                          <option value="<%=co_etapa%>"> 
                          <%response.Write(no_etapa)%>
                          </option>
                          <%						
end if
RS0e.MOVENEXT
WEND
%>
                        </select>
                      </div></td>
                    <td width="25%" height="10" align="center"> <div align="center" id="divTurma"> 
                        <select name="turma" class="borda">
                          <%if turma_pesquisa="999990" then%>
                          <option value="999990" selected></option>
                          <%end if						

	
		Set RS0t = Server.CreateObject("ADODB.Recordset")
		SQL0t = "SELECT * FROM TB_Turma where NU_Unidade="&unidade_pesquisa&"AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='" & etapa_pesquisa & "' order by CO_Turma" 
		RS0t.Open SQL0t, CON0						
co_turma_check=9999990
while not RS0t.EOF
co_turma= RS0t("CO_Turma")

if co_turma = co_turma_check then
RS0t.MOVENEXT
else

if co_turma = turma_pesquisa then
capacidade= RS0t("NU_Capacidade")
%>
                          <option value="<%response.Write(co_turma)%>" selected> 
                          <%response.Write(co_turma)%>
                          </option>
                          <%
else
%>
                          <option value="<%=co_turma%>"> 
                          <%response.Write(co_turma)%>
                          </option>
                          <%						
end if

co_turma_check = co_turma
RS0t.MOVENEXT
end if
WEND
%>
                        </select>
                      </div></td>
                  </tr>
                </table>
                <%end if%>
              </div></td>
          </tr>
          <tr>
            <td height="0"></td>
            <td height="0" colspan="3">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td colspan="3"><hr></td>
                </tr>
                <tr> 
                  <td width="33%"><div align="center"> 
                      <input name="SUBMIT5" type=button class="borda_bot3" onClick="MM_goToURL('parent','../../../index.asp?nvg=WS');return document.MM_returnValue" value="Voltar">
                    </div></td>
                  <td width="34%">&nbsp;</td>
                  <td width="33%"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      <input name="Button22" type="submit" class="borda_bot2"value="Confirmar">
                      </font></div></td>
                </tr>
              </table></td>
          </tr>
        </table>
      </form> </td>
                </tr>				
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

</body>

</html>
<%
	SESSION("recria_att")=""
	SESSION("recria_at")=""	
	SESSION("unidade_pesquisa")=""
	SESSION("curso_pesquisa")=""
	SESSION("etapa_pesquisa")=""
	SESSION("turma_pesquisa")=""
	
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
pasta=arPath(seleciona1)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>