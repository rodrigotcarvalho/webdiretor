<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<% 
session("nvg")=""
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
nvg=request.QueryString("nvg")
opt = request.QueryString("opt")
session("nvg")=nvg
ano_info=nivel&"-"&nvg&"-"&ano_letivo

if opt="" or isnull("opt") or opt="err1" or opt="err2" then
display="select"
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
		
call VerificaAcesso (CON,nvg,nivel)
autoriza=Session("autoriza")

call navegacao (CON,nvg,nivel)
navega=Session("caminho")

if opt="search" then
	busca1=request.form("busca1") 
	busca2=request.form("busca2")
	unidade=request.form("unidade") 
	curso=request.form("curso")
	co_etapa=request.form("etapa") 
	turma=request.form("turma")	
	valor3=session("valor3")
	valor4=session("valor4")
	valor5=session("valor5")
	valor6=session("valor6")
	valor7=session("valor7")	
	
	
		
	if unidade=999990 or unidade="999990" then		
		if busca1 ="" then
			query = busca2
			mensagem=304
		elseif busca2 ="" then
			query = busca1 
			mensagem=303
		end if 
	
		teste = IsNumeric(query)
		if teste = TRUE Then
	  
			Set RS = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT * FROM TB_Alunos where CO_Matricula = "& query
			RS.Open SQL, CON1
				
			if RS.EOF Then
				display="reselect"
			else	
				response.Redirect("../../../../relatorios/swd056.asp?ori=ebe&opt=01&cod_cons="&query&"")	
			end if
		
		ELSE	
			busca=busca_por_nome(query,CON1,"alun")
			alunos_encontrados = split(busca, "#!#" )
			
			if ubound(alunos_encontrados)=-1 then
				display="reselect"
			elseif ubound(alunos_encontrados)=0 then
				cod_cons=alunos_encontrados(0)
				response.Redirect("../../../../relatorios/swd056.asp?ori=ebe&opt=01&cod_cons="&cod_cons&"")	
			else
				display="list"
			end if
		END IF
	else
		if curso="" or isnull(curso) then
			curso=valor4
		end if
		if co_etapa="" or isnull(co_etapa) then
			co_etapa=valor5
		end if
		if turma="" or isnull(turma) then
			turma=valor6
		end if
		if periodo="" or isnull(periodo) then
			periodo=valor7
		end if
			
		obr=unidade&"_"&curso&"_"&co_etapa&"_"&turma
		response.Redirect("../../../../relatorios/swd056.asp?ori=ebe&opt=02&obr="&obr)			
	end if	
elseif opt="listall" then
	display="listall"
end if
%>
<html>
<head>
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
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
/*  if (document.busca.busca1.value != "" && document.busca.busca2.value != "")
  {    alert("Por favor digite SOMENTE uma opção de busca!")
    document.busca.busca1.value = "";
	document.busca.busca2.value = "";    
    document.busca.busca1.focus()
    return false
  }
  if (document.busca.busca1.value != "" && document.busca.unidade.value != "999990")
  {    alert("Por favor digite SOMENTE uma opção de busca!")
    document.busca.busca1.value = "";
	document.busca.busca2.value = "";   
	var combo = document.getElementById("unidade");
	combo.options[0].selected = "true";
	//document.busca.unidade.selectedIndex = "999990";  
    document.busca.busca1.focus()
    return false
  }
  if (document.busca.unidade.value != "999990" && document.busca.busca2.value != "")
  {    alert("Por favor digite SOMENTE uma opção de busca!")
    document.busca.busca1.value = "";
	document.busca.busca2.value = "";   
	var combo = document.getElementById("unidade");
	combo.options[0].selected = "true";
    document.busca.busca1.focus()
    return false
  }  
    if (document.busca.busca1.value == "" && document.busca.busca2.value == "" && document.busca.unidade.value == 999990)
  {    alert("Por favor digite uma opção de busca!")
    document.busca.busca1.focus()
    return false
  }*/
  if (document.busca.unidade.value == 999990 )
  {    alert("Por favor selecione uma unidade!")
	var combo = document.getElementById("unidade");
	combo.options[0].selected = "true";
    return false
  }    
  return true
}

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}function submitfuncao()  
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
						
						
/*						 function recuperarCurso(uTipo)
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
document.all.divEtapa.innerHTML ="<select class=select_style></select>"
document.all.divTurma.innerHTML = "<select class=select_style></select>"
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
document.all.divTurma.innerHTML = "<select class=select_style></select>"
//recuperarTurma()
                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t", true);

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
                                   }*/
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
document.all.divEtapa.innerHTML ="<select class=select_style id=etapa><option value='99990' selected></option></select>"
document.all.divTurma.innerHTML = "<select class=select_style id=turma><option value='99990' selected></option></select>"
GuardaUnidade(uTipo)
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
document.all.divTurma.innerHTML = "<select class=select_style id=turma><option value='99990' selected></option></select>"
GuardaCurso(cTipo)
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
GuardaEtapa(eTipo)																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
	 function GuardaUnidade(u)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../../../../../global/guarda_valores_digitados.asp?opt=valor3", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }

	
		   oHTTPRequest.send("valor3=" + u);
		}
	 function GuardaCurso(c)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../../../../../global/guarda_valores_digitados.asp?opt=valor4", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor4=" + c);
		}
function GuardaEtapa(e)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../../../../../global/guarda_valores_digitados.asp?opt=valor5", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor5=" + e);
		}			
		

	 function gravarTurma(t)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../../../../../global/guarda_valores_digitados.asp?opt=valor6", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor6=" + t);
		}	
 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
                        </script>
</head>
<% 
if opt="listall" or opt="list" then
onload=""
else
onload=""
'onload="onLoad=MM_callJS('document.busca.busca1.focus()')"
end if%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" <%response.Write(onload)%>>
         <form action="index.asp?opt=search&nvg=<%=nvg%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()">              	  

<%call cabecalho(nivel)
%>
<table width="1000" height="670" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr>             
    <td width="1000" height="10" colspan="5" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
    </td>
          </tr>
<%if opt="err1" then%>
          <tr> 
            
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,700,1,0) %>
    </td>
          </tr> 
<%elseif opt="err2" then%>
          <tr> 
            
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,701,1,0) %>
    </td>
          </tr> 
<%
end if
if display="listall" then%>
          <tr> 
            
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,1,0,0) %>
    </td>
          </tr>          
<%elseif display="reselect" then%>
            <tr>              
    <td height="10" colspan="5"> 
      <%call mensagens(nivel,mensagem,1,0) %>
    </td>
			   </tr>          
<%
end if
if display="select" or display="reselect" or display="list"  then%>
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,300,0,0) %>
    </td>
			  </tr>	
          <tr class="tb_tit">             
      <td height="10" colspan="5">Preencha um dos campos abaixo</td>
          </tr>
          <tr>                   
            <td height="10" colspan="5" valign="top">  
  <table width="100%" border="0" cellspacing="0">
    <tr> 
      <td width="250" class="tb_subtit"> 
        <div align="center">UNIDADE 
          </div></td>
      <td width="250" class="tb_subtit"> 
        <div align="center">CURSO 
          </div></td>
      <td width="250" class="tb_subtit"> 
        <div align="center">ETAPA 
          </div></td>
      <td width="250" class="tb_subtit"> 
        <div align="center">TURMA 
          </div></td>
      </tr>
    <tr> 
      <td width="250"> 
        <div align="center"> 
          <select name="unidade" class="select_style" id="unidade" onChange="recuperarCurso(this.value)">
            <option value="999990" selected></option>
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
      <td width="250"> 
        <div align="center"> 
          <div id="divCurso"> 
            <select class="select_style">
              </select>
            </div>
          </div></td>
      <td width="250"> 
        <div align="center"> 
          <div id="divEtapa"> 
            <select class="select_style">
              </select>
            </div>
          </div></td>
      <td width="250"> 
        <div align="center"> 
          <div id="divTurma"> 
            <select class="select_style">
              </select>
            </div>
          </div></td>
      </tr>
    <tr>
      <td height="15" colspan="4" bgcolor="#FFFFFF"><hr></td>
      </tr>
    <tr> 
      <td width="250" height="15" bgcolor="#FFFFFF"></td>
      <td width="250" height="15" bgcolor="#FFFFFF"></td>
      <td width="250" height="15" bgcolor="#FFFFFF"></td>
      <td width="250" height="15" bgcolor="#FFFFFF"><font size="2" face="Arial, Helvetica, sans-serif">
        <div align="center"><input name="Submit" type="submit" class="botao_prosseguir" id="Submit" value="Procurar"></div>
        </font></td>
      </tr>
    </table>
              
            </td>
          </tr>
 
                <tr>                   
    <td colspan="5" valign="top"> 
    </td>
  </tr>        
<%
	if display="list" then
	%>
                <tr>                   
    <td height="10" colspan="5" valign="top"> 
    <hr>
    </td>
  </tr>       
					<tr class="tb_corpo">                   
		<td height="10" colspan="5" class="tb_tit">Alunos Encontrados</td>
					</tr>
					<tr> 
					  
		<td colspan="5" valign="top">
         <ul> 
       <%	for i =0 to ubound(alunos_encontrados)
		cod_cons=alunos_encontrados(i)
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula="&cod_cons
		RS.Open SQL, CON1

		nome = RS("NO_Aluno")
		'ativo = RS("IN_Ativo_Escola")
		ativo= "True" 
			if ativo = "True" then
			Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=ativos href=../../../../relatorios/swd025.asp?ori=ebe&opt=01&cod_cons="&cod_cons&" >"&nome&"</a></font></li>")
			else
			Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=inativos href=../../../../relatorios/swd025.asp?ori=ebe&opt=01&cod_cons="&cod_cons&">"&nome&"</a></font></li>")
			end if
		NEXT			
	%>
		  </ul>
          </td>
                </tr>  
<%
	END IF
elseif display="listall" then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos Order BY NO_Aluno"
		RS.Open SQL, CON1
%>

                <tr class="tb_corpo"> 
                  
    <td height="10" colspan="5" class="tb_tit">Lista de completa de Alunos</td>
                </tr>
                <tr> 
                  
    <td colspan="5" valign="top"> 
      <ul>
        <%
	WHile Not RS.EOF
	nome = RS("NO_Aluno")
	cod = RS("CO_Matricula")
	ativo = RS("IN_Ativo_Escola")
		if ativo = "True" then
		Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=ativos href=../../../../relatorios/swd025.asp?ori=ebe&opt=01&cod_cons="&cod&" >"&nome&"</a></font></li>")
		else
		Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=inativos href=../../../../relatorios/swd025.asp?ori=ebe&opt=01&cod_cons="&cod&">"&nome&"</a></font></li>")
		end if
	RS.Movenext
	Wend
%>
      </ul></td>
                </tr>
<%end if %>                
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
  </form>  
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