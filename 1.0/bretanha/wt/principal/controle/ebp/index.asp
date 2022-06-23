<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/boletos.asp"-->
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

	'response.Redirect("../../../../relatorios/gera_boleto.asp?tp=EBP&ucet="&ucet&"&de="&de&"&ate="&ate&"&r="&restricao)




if opt="" or isnull("opt") or opt="err1" or opt="err2" or opt="err3" or opt="err4" then
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
    opcao=request.form("opcao")
	if opcao = "mat" then
		busca1=request.form("busca1") 
		busca2=request.form("busca2")
		de=request.form("data_de") 
		ate=request.form("data_ate")		
		restricao=request.form("restricao_mat")			
	else
		unidade=request.form("unidade") 
		curso=request.form("curso")
		co_etapa=request.form("etapa") 
		turma=request.form("turma")	
		mes=request.form("mes") 	
		restricao=request.form("restricao_ucet")
	end if	
	
	if opcao = "mat" then		
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
				boletoGerado = GeraBoletos(4, null, query, null, null, de, ate, "EBP", restricao)		
				'response.Redirect("../../../../relatorios/gera_boleto.asp?tp=EBP&c="&query&"&de="&de&"&ate="&ate&"&r="&restricao)	
			end if
		
		ELSE
			busca=busca_por_nome(query,CON1,"alun")
			alunos_encontrados = split(busca, "#!#" )
			
			if ubound(alunos_encontrados)=-1 then
				display="reselect"
			elseif ubound(alunos_encontrados)=0 then
				cod_cons=alunos_encontrados(0)
				boletoGerado = GeraBoletos(4, null, cod_cons, null, null, de, ate, "EBP", restricao)								
				'response.Redirect("../../../../relatorios/gera_boleto.asp?tp=EBP&c="&cod_cons&"&de="&de&"&ate="&ate&"&r="&restricao)	
			else
				display="list"
			end if
		END IF
	else
		ucet=unidade&"_"&curso&"_"&co_etapa&"_"&turma
		boletoGerado = GeraBoletos(4, null, null, mes, ucet, null, null, "EBP", restricao)
					response.Write(boletoGerado)
					response.Flush()
					
		if boletoGerado = "S" then
			response.Write("OK")
			response.End()
		end if		
		'response.Redirect("index.asp?nvg=WT-PR-CR-EBP&opt=ok&ucet="&ucet&"&de="&de&"&ate="&ate&"&r="&restricao)					
		
	end if	
elseif opt="ok" then
    ucet = request.QueryString("ucet")
	de = request.QueryString("de")
	ate = request.QueryString("ate")
	restricao = request.QueryString("r")
	boletoGerado = GeraBoletos(ucet,de,ate,restricao)
	if boletoGerado = "S" then
		response.Write("OK")
	end if
	'response.Redirect("../../../../relatorios/gera_boleto.asp?tp=EBP&ucet="&ucet&"&de="&de&"&ate="&ate&"&r="&restricao)
	display="select"		
elseif opt="listall" then
	display="listall"
end if
%>
<html>
<head>
<title>Web Diretor</title>
<meta charset="UTF-8">
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
if (document.getElementById('radio_mat').checked == false && document.getElementById('radio_ucet').checked == false) {
	alert("Por favor selecione uma opção de busca!")
    return false	
}
if (document.getElementById('radio_mat').checked == true)
  {    
	if (document.getElementById('busca1').value == "" && document.getElementById('busca2').value == "")
	  {    alert("Por favor digite pelo menos uma opção de busca!")  
		document.getElementById('busca1').focus()
		return false
	  }  
  
	if (document.getElementById('busca1').value != "" && document.getElementById('busca2').value != "")
	  {    alert("Por favor digite SOMENTE uma opção de busca!")
		document.getElementById('busca1').value = "";
		document.getElementById('busca2').value = "";    
		document.getElementById('busca1').focus()
		return false
	  }
	  
	if (document.getElementById('data_de').value == "" )
	  {    alert("Por favor selecione uma data inicial!")  
		document.getElementById('data_de').focus()
		return false
	  }  	
	if (document.getElementById('data_ate').value == "" )
	  {    alert("Por favor selecione uma data final!")  
		document.getElementById('data_ate').focus()
		return false
	  }  
	if (document.getElementById('data_de').value>document.getElementById('data_ate').value )
	  {    alert("A data final deve ser maior ou igual a data inicial!")  
		document.getElementById('data_ate').focus()
		return false
	  }  	  		    
}

if (document.getElementById('radio_ucet').checked == true)
  {    
	if (document.getElementById('unidade').options[0].selected == true)
	  {    alert("Por favor selecione pelo menos uma unidade!")
		var combo = document.getElementById("unidade");
		return false
	  }
	  //if (document.getElementById('unidade').options[0].selected == false) { 
//		  MM_showHideLayers('carregando','','show','carregando_fundo','','show')  
//	  }	  
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
                                   }

 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
                        </script>
</head>

<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" <%response.Write(onload)%>>

<%call cabecalho(nivel)
%>
         <form action="index.asp?opt=search&nvg=<%=nvg%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()">     
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
      <%call mensagens(nivel,9716,1,0) %>
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
      <%call mensagens(nivel,9715,0,0) %>
    </td>
			  </tr>	
         	  
          <tr class="tb_tit">             
      <td height="10" colspan="5">Preencha um dos campos abaixo</td>
          </tr>
          <TR>
      <td height="10" valign="top"> 
<div id="carregando"  align="center" style="position:absolute;  top: 200px; width:1000px; z-index: 4; height: 150px; visibility: hidden;">
				  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="75" height="75" vspace="80" title="Carregando">
				    <param name="movie" value="../../../../img/carregando.swf">
				    <param name="quality" value="high">
				    <param name="wmode" value="transparent">
				    <embed src="../../../../img/carregando.swf" width="75" height="75" vspace="80" quality="high" wmode="transparent" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash"></embed>
			      </object>
	    </div>              
		<div id="carregando_fundo" align="center" style="position:absolute; width:1000px; z-index: 3; height: 150px; visibility: hidden; background-color:#FFF; top: 250px;  filter: Alpha(Opacity=90, FinishOpacity=100, Style=0, StartX=0, StartY=100, FinishX=100, FinishY=100);">
			 </div> 
        <table width="1000" border="0" cellpadding="0" cellspacing="0">
          <tr class="form_dado_texto">          
            <td width="204" height="20" align="right">Selecione uma opção para busca: </td><td width="31" height="20" align="center" valign="middle"> <input name="opcao" type="radio" id ="radio_mat" value="mat" onClick="document.getElementById('mat').style.display='block';document.getElementById('envio').style.display='block';document.getElementById('ucet').style.display='none';" ></td>
            <td width="72" valign="middle">Matr&iacute;cula            </td>
            <td width="22" valign="middle"><input name="opcao" id ="radio_ucet" type="radio" value="ucet" onClick="document.getElementById('ucet').style.display='block';document.getElementById('envio').style.display='block';document.getElementById('mat').style.display='none';" ></td>
            <td width="648" valign="middle">Unidade, Curso, Etapa, Turma</td>
            <td width="23" height="20"> </td></tr>
          <tr class="form_dado_texto">
            <td height="20" colspan="6" align="right"><hr></td>
          </tr>
        </table>
                         
             <div id="mat" style="display: none;">  
               <table width="1000" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td class="tb_subtit"><div align="center">MATRÍCULA</div></td>
            <td class="tb_subtit"><div align="center">NOME </div></td>
            <td class="tb_subtit"><div align="center">DE </div></td>
            <td align="center" class="tb_subtit">ATÉ</td>
            <td class="tb_subtit"><div align="center">TIPO</div></td>
          </tr>
          <tr>          
            <td width="104"  height="10" align="center"><font size="2" face="Arial, Helvetica, sans-serif">
              <input name="busca1" type="text" class="textInput" id="busca1" size="12">
            </font></td>
            
            <td width="323" height="10" align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font><font size="2" face="Arial, Helvetica, sans-serif">
              <input name="busca2" type="text" class="textInput" id="busca2" size="55" maxlength="50">
              </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font></td>
            
            <td width="172" height="10" align="center">
              <input name="data_de" type="date" class="textInput" id="data_de" max="<%response.Write(ano_letivo)%>-12-31" min="<%response.Write(ano_letivo)%>-01-01" value="<%response.Write(ano_letivo)%>-01-01"></td>
            
            <td width="236" height="10" align="center" ><input name="data_ate" type="date" class="textInput" id="data_ate" max="<%response.Write(ano_letivo)%>-12-31"  min="<%response.Write(ano_letivo)%>-01-01" value="<%response.Write(ano_letivo)%>-12-31"></td>
            
            <td width="165" height="10" align="center"><select name="restricao_mat" id="restricao_mat" class="select_style">
                    <option value="M" selected>Mensalidade com Servi&ccedil;os</option>
                    <option value="S">Servi&ccedil;os</option>              
              </select></td>
          </tr>
		  </table>
   		</div>   
             <div id="ucet" style="display: none;">  
                <table border="0" cellpadding="0" cellspacing="0"><tr><td width="250" class="tb_subtit"> 
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
                  <td width="250" align="center" class="tb_subtit">MÊS</td>
                  <td width="250" class="tb_subtit"><div align="center">TIPO</div></td>
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
                  <td width="250" align="center">                  
                  <select name="mes" id="mes" class="select_style">
                    <option value="1" selected>Janeiro</option>
                    <option value="2">Fevereiro</option>       
                    <option value="3">Mar&ccedil;o</option>  
                    <option value="4">Abril</option>  
                    <option value="5">Maio</option>  
                    <option value="6">Junho</option>  
                    <option value="7">Julho</option>
                    <option value="8">Agosto</option>
                    <option value="9">Setembro</option>
                    <option value="10">Outubro</option>
                    <option value="11">Novembro</option>
                    <option value="12">Dezembro</option>
                    </select></td>
                  <td width="250" align="center">
                  <select name="restricao_ucet" id="restricao_ucet" class="select_style">
                    <option value="M" selected>Mensalidade com Servi&ccedil;os</option>
                    <option value="S">Servi&ccedil;os</option>              
                  </select></td></tr></table>
    		</div>                  
		  </td>
		  </TR>
           <tr>                   
    <td height="10" colspan="5" valign="top"> 
    <div id="envio" style="display: none;"> 
<table width="100%" border="0" cellspacing="0">
                <tr>
                  <td height="15" colspan="7" bgcolor="#FFFFFF"><hr></td>
                </tr>
                <tr> 
                  <td width="250" height="15" bgcolor="#FFFFFF"></td>
                  <td width="250" height="15" bgcolor="#FFFFFF"></td>
                  <td width="250" height="15" bgcolor="#FFFFFF"></td>
                  <td width="250" height="15" bgcolor="#FFFFFF"></td>
                  <td width="250" bgcolor="#FFFFFF">&nbsp;</td>
                  <td width="250" bgcolor="#FFFFFF">&nbsp;</td>
                  <td width="250" bgcolor="#FFFFFF"><font size="2" face="Arial, Helvetica, sans-serif">
                    <div align="center"><input name="Submit" type="submit" class="botao_prosseguir" id="Submit" value="Procurar"></div>
                  </font></td>
                </tr>
              </table>
	</div>
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
			Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=ativos href=../../../../relatorios/gera_boleto.asp?tp=EBP&c="&cod_cons&"&de="&de&"&ate="&ate&"&r="&restricao&">"&nome&"</a></font></li>")
			else
			Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=inativos href=../../../../relatorios/gera_boleto.asp?tp=EBP&c="&cod_cons&"&de="&de&"&ate="&ate&"&r="&restricao&">"&nome&"</a></font></li>")
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
		Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=ativos href=../../../../relatorios/gera_boleto.asp?tp=EBP&c="&cod_cons&"&de="&de&"&ate="&ate&"&r="&restricao&" >"&nome&"</a></font></li>")
		else
		Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=inativos href=../../../../relatorios/gera_boleto.asp?tp=EBP&c="&cod_cons&"&de="&de&"&ate="&ate&"&r="&restricao&"</a></font></li>")
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