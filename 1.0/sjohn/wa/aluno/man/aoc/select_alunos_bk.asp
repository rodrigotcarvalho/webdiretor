<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<% 
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")


chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
volta=request.QueryString("vt")


if volta="s" then
	obr = request.QueryString("obr")
	obr_split=split(obr,"$!$")

	unidade=obr_split(0)
	curso=obr_split(1)
	co_etapa=obr_split(2)
	turma=obr_split(3)
	
	opt = request.QueryString("opt")	
	matrics = request.QueryString("cod")	
	
else
	unidade=request.form("unidade")
	curso=request.form("curso")
	if isnull(curso) or curso=""  then
		curso=session("c_aoc")
	end if	
	
	co_etapa=request.form("etapa")
	if isnull(co_etapa) then		
		co_etapa=session("e_aoc")
	end if	
	
	turma=request.form("turma")	
	if isnull(turma) then				
		turma=session("t_aoc")
	end if	
end if
obr=unidade&"#!#"&curso&"#!#"&co_etapa&"#!#"&turma
if unidade=999990 or isnull(unidade) or unidade="" then
	sql_unidade=""
	sql_curso=""
	sql_co_etapa=""		
	sql_turma=""	
else
	sql_unidade="TB_Matriculas.NU_Unidade = "& unidade&" AND "
	
	teste_curso = isnumeric(curso)
	
	if teste_curso= TRUE then		
		if curso=999990 or isnull(curso) or curso="" then
			sql_curso=""
			sql_co_etapa=""	
			sql_turma=""	
		else
			sql_curso="TB_Matriculas.CO_Curso = '"& curso &"' AND "
			
			teste_co_etapa = isnumeric(co_etapa)
			
			if teste_co_etapa= TRUE then	
				if co_etapa=999990 or isnull(co_etapa) or co_etapa="" then
					sql_co_etapa=""
					sql_turma=""	
				else
					sql_co_etapa="TB_Matriculas.CO_Etapa = '"& co_etapa &"' AND "
					if turma="999990" or isnull(turma) or turma="" then
						sql_turma=""
					else
						sql_turma="TB_Matriculas.CO_Turma = '"& turma &"' AND "
					end if			
				end if
			else
				if co_etapa="999990" or co_etapa="" then
					sql_co_etapa=""
					sql_turma=""	
				else
					sql_co_etapa="TB_Matriculas.CO_Etapa = '"& co_etapa &"' AND "
					if turma="999990" or isnull(turma) or turma="" then
						sql_turma=""
					else
						sql_turma="TB_Matriculas.CO_Turma = '"& turma &"' AND "
					end if			
				end if
			end if				
		end if	
	else
		if curso="999990" or isnull(curso) or curso="" then
			sql_curso=""
			sql_co_etapa=""	
			sql_turma=""	
		else
			sql_curso="TB_Matriculas.CO_Curso = '"& curso &"' AND "
			
			teste_co_etapa = isnumeric(co_etapa)
			
			if teste_co_etapa= TRUE then	
				if co_etapa=999990 or isnull(co_etapa) or co_etapa="" then
					sql_co_etapa=""
					sql_turma=""	
				else
					sql_co_etapa="TB_Matriculas.CO_Etapa = '"& co_etapa &"' AND "
					if turma="999990" or isnull(turma) or turma="" then
						sql_turma=""
					else
						sql_turma="TB_Matriculas.CO_Turma = '"& turma &"' AND "
					end if			
				end if
			else
				if co_etapa="999990" or co_etapa="" then
					sql_co_etapa=""
					sql_turma=""	
				else
					sql_co_etapa="TB_Matriculas.CO_Etapa = '"& co_etapa &"' AND "
					if turma="999990" or isnull(turma) or turma="" then
						sql_turma=""
					else
						sql_turma="TB_Matriculas.CO_Turma = '"& turma &"' AND "
					end if			
				end if
			end if				
		end if	
	
	end if
end if

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
var currentlyActiveInputRef = false;
var currentlyActiveInputClassName = false;

function highlightActiveInput() {
  if(currentlyActiveInputRef) {
    currentlyActiveInputRef.className = currentlyActiveInputClassName;
  }
  currentlyActiveInputClassName = this.className;
  this.className = 'inputHighlighted';
  currentlyActiveInputRef = this;
}

function blurActiveInput() {
  this.className = currentlyActiveInputClassName;
}

function initInputHighlightScript() {
  var tags = ['INPUT','TEXTAREA'];
  for(tagCounter=0;tagCounter<tags.length;tagCounter++){
    var inputs = document.getElementsByTagName(tags[tagCounter]);
    for(var no=0;no<inputs.length;no++){
      if(inputs[no].className && inputs[no].className=='doNotHighlightThisInput')continue;
      if(inputs[no].tagName.toLowerCase()=='textarea' || (inputs[no].tagName.toLowerCase()=='input' && inputs[no].type.toLowerCase()=='text')){
        inputs[no].onfocus = highlightActiveInput;
        inputs[no].onblur = blurActiveInput;
      }
    }
  }
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
function checksubmit()
{
  if (document.busca.busca1.value != "" && document.busca.busca2.value != "")
  {    alert("Por favor digite SOMENTE uma opção de busca!")
    document.busca.busca1.focus()
    return false
  }
    if (document.busca.busca1.value == "" && document.busca.busca2.value == "")
  {    alert("Por favor digite uma opção de busca!")
    document.busca.busca1.focus()
    return false
  }
  return true
}

function checksubmit_alunos()
{
 	if ( document.seleciona_alunos.alunos.checked == false )
    {	alert ( "Por favor selecione pelo menos um aluno!" );
	 return false		
    }
  return true
}



function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
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
//Nessa função o formulário a ser enviadopor essa função é o segundo
//function submitfuncao()  
//{
//   var f=document.forms[3]; 
//      f.submit(); 
//}
function submitfuncao()  
{
   var f=document.forms[4]; 
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
//recuperarTurma()
                                                           }
                                               }
 
                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }
 
 
						 function recuperarTurma(eTipo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t5", true);
 
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

						 function gravarTurma(tTipo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t6", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma2.innerHTML = resultado_t																   
                                                           }
                                               }
 
                                               oHTTPRequest.send("t_pub=" + tTipo);
                                   }

var checkflag = "false";
function check(field) {
if (checkflag == "false") {
for (i = 0; i < field.length; i++) {
field[i].checked = true;}
checkflag = "true";
return "Desmarcar Todos"; }
else {
for (i = 0; i < field.length; i++) {
field[i].checked = false; }
checkflag = "false";
return "Marcar Todos"; }
}
</script>

</head>

<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<%call cabecalho(nivel)
%>
<table width="1000" height="685" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
            
    <td height="10" colspan="5" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
    </td>
  </tr>
<% if opt="ok1" then%>    
  <tr> 
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,312,2,0) %>
    </td>
			  </tr>	
<%elseif opt="err1" then%>  
  <tr>           
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,319,1,cod) %>
    </td>
	</tr>  
<%elseif opt="err1" then%>  
  <tr>           
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,320,1,cod) %>
    </td>
	</tr>    	
<%end if%> 
	 <tr>             
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,321,0,0) %>
    </td>
	  </tr>                         
  <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,9706,0,0) %>
    </td>
			  </tr>			  
        <form action="index.asp?opt=list&nvg=<%=chave%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
          <tr class="tb_tit"> 
            
      <td height="10" colspan="5">Preencha um dos campos abaixo</td>
          </tr>
          <TR>
		  
      <td height="26" valign="top"> 
        <table width="1000" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            
            <td width="147"  height="10"> 
              <div align="right"><font class="form_dado_texto"> Matr&iacute;cula:</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                </strong></font></div></td>
            
            <td width="62" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="busca1" type="text" class="textInput" id="busca1" size="12">
              </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font></td>
            
            <td width="147" height="10"> 
              <div align="right"><font class="form_dado_texto"> Nome: </font></div></td>
            
            <td width="392" height="10" ><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="busca2" type="text" class="textInput" id="busca2" size="55" maxlength="50">
              </font></td>
            
            <td width="250" height="10"><div align="center">
              <input name="Submit" type="submit" class="botao_prosseguir" id="Submit" value="Procurar">
              </div> </td>
          </tr>
		  </table>
		  </td>
		  </TR>
      </form>
      <tr>    
      	<td height="10"><hr> 
	 	</td>
  </tr>
<form name="alteracao" method="post" action="select_alunos.asp">      
      <tr>    
      	<td valign="top"> 
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
                </tr>
                <tr> 
                  <td width="14%"> 
                    <div align="center"> 
                      <select name="unidade" class="select_style" onChange="recuperarCurso(this.value)">
                        <option value="999990"></option>                            
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
                        <option value="999990"></option>      
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
                        <option value="999990"></option>                              
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
                        <select name="turma" class="select_style" onChange="gravarTurma(this.value)">
                        <option value="999990"></option>                              
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
                </tr>
                <tr> 
                  <td height="15" colspan="4" bgcolor="#FFFFFF"><hr></td>
                </tr>
                <tr>
                  <td height="15" colspan="4" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0">
                    <tr>
                      <td width="25%"><div align="center"><div id="divTurma2"></div></div></td>
                      <td width="25%"><div align="center"></div></td>
                      <td width="25%">&nbsp;</td>
                      <td width="25%"><div align="center"><font size="3" face="Courier New, Courier, mono">
                        <input type="submit" name="Submit2" value="Prosseguir" class="botao_prosseguir">
                      </font></div></td>
                    </tr>
                  </table></td>
                </tr>
            </table>        
	 	</td>
    </tr>  
  </form>         
      <tr>    
      	<td height="10"><hr> 
	 	</td>
  </tr>
                    
<%


		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = 	"Select TB_Matriculas.CO_Matricula, TB_Matriculas.NU_Chamada, TB_Alunos.NO_Aluno, TB_Matriculas.NU_Unidade,TB_Matriculas.CO_Curso,TB_Matriculas.CO_Etapa,TB_Matriculas.CO_Turma from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Situacao='C' AND "&sql_unidade&sql_curso&sql_co_etapa&sql_turma&" TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula order by TB_Matriculas.NU_Unidade,TB_Matriculas.CO_Curso,TB_Matriculas.CO_Etapa,TB_Matriculas.CO_Turma,TB_Matriculas.NU_Chamada"
		RS.Open SQL, CON1
%>

 <tr class="tb_corpo"> 
                  
    <td height="10" colspan="5" class="tb_tit">Lista de alunos que atendem as condições informadas</td>
                </tr>
<form name="seleciona_alunos" method="post" action="incluir_multi.asp" onSubmit="return checksubmit_alunos()">                  
    <tr> 
                  
    <td colspan="5" valign="top"> 
    <table width="1000" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="20"><input type="checkbox" name="todos" class="borda" value="" onClick="this.value=check(this.form.alunos)"></td>
    <td width="60" class="tb_subtit"><div align="center">Chamada</div></td>
    <td width="500" class="tb_subtit">&nbsp;Nome<input name="obr" type="hidden" value="<%response.write(obr)%>"></td>
    <td width="105" height="10" class="tb_subtit"><div align="center">Unidade</div></td>
    <td width="115" height="10" class="tb_subtit"><div align="center">Curso</div></td>
    <td width="100" height="10" class="tb_subtit"><div align="center"> Etapa</div></td>
    <td width="95" height="10" class="tb_subtit"><div align="center">Turma </div></td>
    </tr>
        <%
check=0		
WHile Not RS.EOF
	nu_matricula = RS("CO_Matricula")
	no_aluno= RS("NO_Aluno")			
	nu_chamada = RS("NU_Chamada")
	unidade_aluno = RS("NU_Unidade")
	curso_aluno = RS("CO_Curso")
	co_etapa_aluno = RS("CO_Etapa")
	turma_aluno = RS("CO_Turma")
	
call GeraNomes("PORT",unidade_aluno,curso_aluno,co_etapa_aluno,CON0)
no_unidade = session("no_unidades")
no_curso = session("no_grau")
no_etapa = session("no_serie")	
	
	 if check mod 2 =0 then
		cor = "tb_fundo_linha_par" 
	 else 
		cor ="tb_fundo_linha_impar"
	 end if 
 		
%>		
  <tr class="<%=cor%>">
    <td width="20" ><input type="checkbox" name="alunos" id="alunos" class="borda" value="<%=nu_matricula%>"></td>
    <td width="60"><div align="center"><%response.Write(nu_chamada)%></div></td>
    <td width="500">&nbsp;<%response.Write(no_aluno)%></td>
    <td width="105"><div align="center"><%response.Write(no_unidade)%></div></td>
    <td width="115"><div align="center"><%response.Write(no_curso)%></div></td>
    <td width="100"><div align="center"><%response.Write(no_etapa)%></div></td>
    <td width="95"><div align="center"><%response.Write(turma_aluno)%></div></td>
  </tr>
	
 <% 
check=check+1
RS.Movenext
Wend
%>
  <tr class="tb_fundo_linha_par">
    <td colspan="7">
    <table width="1000" border="0" align="center" cellspacing="0">
                      <tr>
                        <td colspan="3"><hr></td>
                      </tr>
                      <tr> 
                        <td width="375">&nbsp;</td>
                        <td width="375">&nbsp;</td>
                        <td width="250"> 
                          <div align="center"> 
                            <input name="Submit" type="submit" class="botao_prosseguir" value="Incluir">
                          </div>
                        </td>
                </tr>
    </table>
    </td>
  </tr>	
</table>
</td>
                </tr>
  </form>                
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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