<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%
opt = request.QueryString("opt")

ano_letivo = session("ano_letivo")
co_usr = session("co_user")
nivel=4
if opt="cln" then
nvg=session("nvg")
else
nvg = request.QueryString("nvg")
session("nvg")=nvg
end if
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano_info=nivel&"-"&nvg&"-"&ano_letivo

if opt="acc" then
	unidade = request.Form("unidade")	
	curso = request.Form("curso")
	co_etapa = request.Form("etapa")
	turma = request.Form("turma")
	periodo = request.Form("periodo")
	acumulado = request.Form("acumulado")
	qto_falta = request.Form("qto_falta")	
	ano_letivo = session("ano_letivo")
	
	obr_mapa=unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma&"$!$"&periodo&"$!$"&acumulado&"$!$"&qto_falta&"$!$"&ano_letivo	
	obr_log=unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&periodo&"_"&ano_letivo
	
	call GravaLog (nvg,obr_log)

	onload="onLoad=""MM_openBrWindow('mapa.asp?obr="&obr_mapa&"','')"""

elseif opt="xml" then
'essa opção é só para o script não entrar em loop na criação do XML
else


response.redirect("criar_xml.asp?nvg="&nvg)

end if


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
call navegacao (CON,nvg,nivel)
navega=Session("caminho")	
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
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=c", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
document.all.divEtapa.innerHTML ="<select name='etapa' class='select_style' id='etapa'><option value='999990' selected>           </option></select>"
document.all.divTurma.innerHTML = "<select name='turma' class='select_style' id='turma'><option value='999990' selected>           </option></select>"
document.all.divPeriodo.innerHTML = "<select name='periodo' class='select_style' id='periodo'><option value='0' selected>           </option></select>"
//recuperarEtapa()
                                                           }
                                               }
 
                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }
 
 
						 function recuperarEtapa(cTipo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=e", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select name='turma' class='select_style' id='turma'><option value='999990' selected>           </option></select>"
document.all.divPeriodo.innerHTML = "<select name='periodo' class='select_style' id='periodo'><option value='0' selected>           </option></select>"
//recuperarTurma()
                                                           }
                                               }
 
                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }
 
 
						 function recuperarTurma(eTipo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t4", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t
document.all.divPeriodo.innerHTML = "<select name='periodo' class='select_style' id='periodo'><option value='0' selected>           </option></select>"
																	   
                                                           }
                                               }
 
                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
function recuperarPeriodo(eTipo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=p1", true);
 
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


function habilita_campo(){
		document.getElementById('qto_falta1').disabled   = false;
		document.getElementById('qto_falta2').disabled   = false;		
		document.getElementById('qto_falta1').checked   = true;
		document.getElementById('qto_falta2').checked   = false;		    
}

function desabilita_campo(){
		document.getElementById('qto_falta1').disabled   = true;
		document.getElementById('qto_falta2').disabled   = true;	   
		document.getElementById('qto_falta1').checked   = false;
		document.getElementById('qto_falta2').checked   = true;		
}



function checksubmit()
{
  if (document.busca.unidade.value == "999990" || document.busca.curso.value == "999990" || document.busca.etapa.value == "999990" || document.busca.turma.value == "999990" || document.busca.periodo.value == "0")
  { alert("É necessário preencher pelo menos Unidade, Curso, Etapa, Turma e Periodo!")
	var combo = document.getElementById("unidade");
	combo.options[0].selected = "true";
	var combo2 = document.getElementById("curso");
	combo2.options[0].selected = "true";	
	var combo3 = document.getElementById("etapa");
	combo3.options[0].selected = "true";	
	var combo4 = document.getElementById("turma");
	combo4.options[0].selected = "true";	
	var combo5 = document.getElementById("periodo");
	combo5.options[0].selected = "true";		
    return false
  }  
   
  return true
}

function MM_openBrWindow(theURL,winName) { //v2.0
var largura=screen.availWidth-5
var altura=screen.availHeight-60
url=theURL+"$!$"+largura+"$!$"+altura
  window.open(url,winName,'status=yes,scrollbars=yes,resizable=yes,width='+largura+',height='+altura+',top=0,left=0,bReplace=true');
}

</script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" <%response.Write(onload)%>>
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
  <tr>                   
    <td height="10"> 
      <%
	  if autoriza="no" then
	  	call mensagens(4,9700,1,0) 	  
	  else
	  	call mensagens(nivel,1,0,0) 
	  end if
	  call ultimo(0) %>
    </td>
                  </tr> 
<tr>

            <td valign="top"> <form name="busca" method="post" action="index.asp?nvg=WA-PF-CN-ACC&opt=acc" onSubmit="return checksubmit()">
                
        <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td width="653" height="15" class="tb_tit">Grade de Aulas</td>
          </tr>
          <tr> 
            <td>
            <%if opt="acc" then%>
<table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="14%" class="tb_subtit"> 
                    <div align="center">UNIDADE 
                    </div></td>
                  <td width="14%" class="tb_subtit"> 
                    <div align="center">CURSO 
                    </div></td>
                  <td width="14%" class="tb_subtit"> 
                    <div align="center">ETAPA 
                    </div></td>
                  <td width="14%" class="tb_subtit"> 
                    <div align="center">TURMA 
                    </div></td>
                  <td width="14%" class="tb_subtit"> 
                  <div align="center">PER&Iacute;ODO</div></td>
                  <td width="14%" class="tb_subtit"><div align="center">ACUMULADO</div></td>
                  <td width="16%" class="tb_subtit"><div align="center">QUANTO FALTA</div></td>
                </tr>
                <% 'if RS1.EOF THEN %>
                <%'else%>
                <tr> 
                  <td width="14%"> 
                    <div align="center"> 
                      <select name="unidade" class="select_style" onChange="recuperarCurso(this.value)">
                        <%		
			Set RS0 = Server.CreateObject("ADODB.Recordset")
			SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
			RS0.Open SQL0, CON0
While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")
unidade=unidade*1
NU_Unidade=NU_Unidade*1
if NU_Unidade=unidade then
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
WEND
%>
                      </select>
                  </div></td>
                  <td width="14%"> 
                    <div align="center"> 
                      <div id="divCurso"> 
<select name="curso" class="select_style" onChange="recuperarEtapa(this.value)">
                          <%		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT Distinct CO_Curso FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade
		RS0.Open SQL0, CON0
		
While not RS0.EOF
CO_Curso = RS0("CO_Curso")

		Set RS0a = Server.CreateObject("ADODB.Recordset")
		SQL0a = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
		RS0a.Open SQL0a, CON0
		
NO_Curso = RS0a("NO_Abreviado_Curso")		

if CO_Curso=curso then
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
RS0.MOVENEXT
WEND
%>
                        </select>
                      </div>
                  </div></td>
                  <td width="14%"> 
                    <div align="center"> 
                      <div id="divEtapa"> 
                        <select name="etapa" class="select_style" onChange="recuperarTurma(this.value)">
                          <%		

		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"'"
		RS0b.Open SQL0b, CON0
		
		
While not RS0b.EOF
Etapa = RS0b("CO_Etapa")


		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&curso&"' AND CO_Etapa='"&Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
if Etapa=co_etapa then
%>
                          <option value="<%response.Write(Etapa)%>" selected> 
                          <%response.Write(NO_Etapa)%>
                          </option>
                          <%
else
%>
                          <option value="<%response.Write(Etapa)%>"> 
                          <%response.Write(NO_Etapa)%>
                          </option>
                          <%

end if
RS0b.MOVENEXT
WEND
%>
                        </select>
                      </div>
                  </div></td>
                  <td width="14%"> 
                    <div align="center"> 
                      <div id="divTurma"> 
                        <select name="turma" class="select_style" onChange="MM_callJS('submitfuncao()')">
                          <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Turma where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"' AND CO_Etapa='" & co_etapa & "' order by CO_Turma" 
		RS3.Open SQL3, CON0						

while not RS3.EOF
co_turma= RS3("CO_Turma")

if co_turma=turma then
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
co_turma_check = co_turma
end if
RS3.MOVENEXT
WEND
%>
                        </select>
                      </div>
                  </div></td>
                  <td width="14%"> 
                    <div align="center"> 
                      <div id="divPeriodo"> 
                        <select name="periodo" class="select_style" id="periodo">
                          <option value="999990"></option>
                          <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")
NU_Periodo=NU_Periodo*1
periodo=periodo*1
if NU_Periodo=periodo then
%>
                          <option value="<%=NU_Periodo%>" selected> 
                          <%response.Write(NO_Periodo)%>
                          </option>
                          <%
else
%>
                          <option value="<%=NU_Periodo%>"> 
                          <%response.Write(NO_Periodo)%>
                          </option>
                          <%
end if
RS4.MOVENEXT
WEND%>
                        </select>
                      </div>
                  </div></td>
                  <td width="14%"><table width="50%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20%"><div align="left">
                      <% if acumulado="s" then%>
                      <input name="acumulado" type="radio" id="acumulado" value="s" onClick="javascript:habilita_campo();" checked>
                      <%else%>
                      <input name="acumulado" type="radio" id="acumulado" value="s" onClick="javascript:habilita_campo();">
                      <%end if%>
                      </div></td>
                      <td width="25%" class="form_dado_texto"><div align="center">S</div></td>
                      <td width="20%"><div align="left">
                      <% if acumulado="s" then%>
                      <input name="acumulado" type="radio" id="acumulado" value="n"  onClick="javascript:desabilita_campo();">
                      <%else%>
                      <input name="acumulado" type="radio" id="acumulado" value="n" checked onClick="javascript:desabilita_campo();">
                      <%end if%>
                      </div></td>
                      <td width="27%" class="form_dado_texto"><div align="center">N</div></td>
                    </tr>
                  </table></td>
                  <td width="16%"><table width="50%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20%"><div align="left">
                      <% if qto_falta="s" then%>
                      <input name="qto_falta" type="radio" id="qto_falta1" value="s" checked>
                      <%elseif acumulado="s" then%>
                      <input name="qto_falta" type="radio" id="qto_falta1" value="s">
                      <%else%>
                      <input name="qto_falta" type="radio" id="qto_falta1" value="s" disabled>
                      <%end if%>                      
						</div></td>
                      <td width="25%" class="form_dado_texto"><div align="center">S</div></td>
                      <td width="20%"><div align="left">
                      <% if qto_falta="s" then%>
                      <input name="qto_falta" type="radio" id="qto_falta2" value="n">
                      <%elseif acumulado="s" then%>
                      <input name="qto_falta" type="radio" id="qto_falta2" value="n" checked>
                      <%else%>
                      <input name="qto_falta" type="radio" id="qto_falta2" value="n" checked disabled>
                      <%end if%>                       
						</div></td>
                      <td width="27%" class="form_dado_texto"><div align="center">N</div></td>
                    </tr>
                  </table></td>
                </tr>
                <%'end if %>
                <tr> 
                  <td height="15" colspan="7" bgcolor="#FFFFFF"><hr></td>
                </tr>
                <tr>
                  <td height="15" colspan="7" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0">
                    <tr>
                      <td width="33%"><div align="center"></div></td>
                      <td width="34%"><div align="center"></div></td>
                      <td width="33%"><div align="center"><font size="3" face="Courier New, Courier, mono">
                        <input type="submit" name="Submit2" value="Prosseguir" class="botao_prosseguir">
                      </font></div></td>
                    </tr>
                  </table></td>
                </tr>
            </table>            
			<%else%>
<table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="14%" class="tb_subtit"> 
                    <div align="center">UNIDADE 
                    </div></td>
                  <td width="14%" class="tb_subtit"> 
                    <div align="center">CURSO 
                    </div></td>
                  <td width="14%" class="tb_subtit"> 
                    <div align="center">ETAPA 
                    </div></td>
                  <td width="14%" class="tb_subtit"> 
                    <div align="center">TURMA 
                    </div></td>
                  <td width="14%" class="tb_subtit"> 
                  <div align="center">PER&Iacute;ODO</div></td>
                  <td width="14%" class="tb_subtit"><div align="center">ACUMULADO</div></td>
                  <td width="16%" class="tb_subtit"><div align="center">QUANTO FALTA</div></td>
                </tr>
                <% 'if RS1.EOF THEN %>
                <%'else%>
                <tr> 
                  <td width="14%"> 
                    <div align="center"> 
                      <select name="unidade" id="unidade" class="select_style" onChange="recuperarCurso(this.value)">
                        <option value="999990"></option>
                        <%		
		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0
NU_Unidade_Check=999999		
While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")

if NU_Unidade = NU_Unidade_Check then
RS0.MOVENEXT		
else
%>
                        <option value="<%response.Write(NU_Unidade)%>"> 
                        <%response.Write(NO_Abr)%>
                        </option>
                        <%

NU_Unidade_Check = NU_Unidade
RS0.MOVENEXT
end if
WEND
%>
                      </select>
                  </div></td>
                  <td width="14%"> 
                    <div align="center"> 
                      <div id="divCurso"> 
                        <select name="curso" class="select_style" id="curso">
                        <option value="999990" selected> 
                        </option>                        
                        </select>
                      </div>
                  </div></td>
                  <td width="14%"> 
                    <div align="center"> 
                      <div id="divEtapa"> 
                        <select name="etapa" class="select_style" id="etapa">
                        <option value="999990" selected> 
                        </option>                        
                        </select>
                      </div>
                  </div></td>
                  <td width="14%"> 
                    <div align="center"> 
                      <div id="divTurma"> 
                        <select name="turma" class="select_style" id="turma">
                        <option value="999990" selected> 
                        </option>                        
                        </select>
                      </div>
                  </div></td>
                  <td width="14%"> 
                    <div align="center"> 
                      <div id="divPeriodo"> 
                        <select name="periodo" class="select_style" id="periodo">
                        <option value="0" selected> 
                        </option>                        
                        </select>
                      </div>
                  </div></td>
                  <td width="14%"><table width="50%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20%"><div align="left"><input name="acumulado" type="radio" id="acumulado" value="s" onClick="javascript:habilita_campo();"></div></td>
                      <td width="25%" class="form_dado_texto"><div align="center">S</div></td>
                      <td width="20%"><div align="left"><input name="acumulado" type="radio" id="acumulado" value="n" checked onClick="javascript:desabilita_campo();"> </div></td>
                      <td width="27%" class="form_dado_texto"><div align="center">N</div></td>
                    </tr>
                  </table></td>
                  <td width="16%"><table width="50%" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="20%"><div align="left"><input name="qto_falta" type="radio" id="qto_falta1" value="s" disabled></div></td>
                      <td width="25%" class="form_dado_texto"><div align="center">S</div></td>
                      <td width="20%"><div align="left"><input name="qto_falta" type="radio" id="qto_falta2" value="n" checked disabled></div></td>
                      <td width="27%" class="form_dado_texto"><div align="center">N</div></td>
                    </tr>
                  </table></td>
                </tr>
                <%'end if %>
                <tr> 
                  <td height="15" colspan="7" bgcolor="#FFFFFF"><hr></td>
                </tr>
                <tr>
                  <td height="15" colspan="7" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0">
                    <tr>
                      <td width="33%"><div align="center"></div></td>
                      <td width="34%"><div align="center"></div></td>
                      <td width="33%"><div align="center"><font size="3" face="Courier New, Courier, mono">
                        <input type="submit" name="Submit2" value="Prosseguir" class="botao_prosseguir">
                      </font></div></td>
                    </tr>
                  </table></td>
                </tr>
            </table>            
            <%end if%>
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