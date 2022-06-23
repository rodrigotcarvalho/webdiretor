<%'On Error Resume Next
%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes7.asp"-->
<!--#include file="../../../../inc/boletos.asp"-->
<!--#include file="../../../../inc/funcoes_email.asp"-->
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
onLoad = "onLoad=""CarregarUCET()"""
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
		
		Set CON6 = Server.CreateObject("ADODB.Connection") 
		ABRIR6 = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON6.Open ABRIR6	
		
		Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4 = "DBQ="& CAMINHO_pf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4			
		
'		Set CON_WF = Server.CreateObject("ADODB.Connection") 
'		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
'		CON_WF.Open ABRIR_WF			
		
call VerificaAcesso (CON,nvg,nivel)
autoriza=Session("autoriza")

call navegacao (CON,nvg,nivel)
navega=Session("caminho")

if opt="search" then
    display="select"
    onLoad = "onLoad=""CarregarUCET();CarregarResult()"""
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
		mes_selecionado=request.form("select_mes") 	
		restricao=request.form("restricao_ucet")
	end if	
	
	botao_submit = request.form("Submit")
	
	if botao_submit = "Enviar Email" then
		vetor_matric = request.form("matric")
		matriculas = split(vetor_matric,",")
		
		for mm =0 to ubound(matriculas)
			enviado=""
			enviado = email_anexo(mes_selecionado, matriculas(mm), "N")
			'response.Write(matriculas(mm)&" "&enviado&"<BR>")
		next

		'response.Write("OK")
	end if
	
response.Charset="UTF-8"

'Trecho comentado abaixo se refere ao código da função Emitir Bloqueto por Período
	
	'if opcao = "mat" then		
'		if busca1 ="" then
'			query = busca2
'			mensagem=304
'		elseif busca2 ="" then
'			query = busca1 
'			mensagem=303
'		end if 
'	
'		teste = IsNumeric(query)
'		if teste = TRUE Then
'	  
'			Set RS = Server.CreateObject("ADODB.Recordset")
'			SQL = "SELECT * FROM TB_Alunos where CO_Matricula = "& query
'			RS.Open SQL, CON1
'				
'			if RS.EOF Then
'				display="reselect"
'			else
'				boletoGerado = GeraBoletos(4, null, query, null, null, de, ate, "EBP", restricao)		
'				'response.Redirect("../../../../relatorios/gera_boleto.asp?tp=EBP&c="&query&"&de="&de&"&ate="&ate&"&r="&restricao)	
'			end if
'		
'		ELSE
'			busca=busca_por_nome(query,CON1,"alun")
'			alunos_encontrados = split(busca, "#!#" )
'			
'			if ubound(alunos_encontrados)=-1 then
'				display="reselect"
'			elseif ubound(alunos_encontrados)=0 then
'				cod_cons=alunos_encontrados(0)
'				boletoGerado = GeraBoletos(4, null, cod_cons, null, null, de, ate, "EBP", restricao)								
'				'response.Redirect("../../../../relatorios/gera_boleto.asp?tp=EBP&c="&cod_cons&"&de="&de&"&ate="&ate&"&r="&restricao)	
'			else
'				display="list"
'			end if
'		END IF
'	else
'		ucet=unidade&"_"&curso&"_"&co_etapa&"_"&turma
'		boletoGerado = GeraBoletos(4, null, null, mes, ucet, null, null, "EBP", restricao)
'					response.Write(boletoGerado)
'					response.Flush()
'					
'		if boletoGerado = "S" then
'			response.Write("OK")
'			response.End()
'		end if		
'		'response.Redirect("index.asp?nvg=WT-PR-CR-EBP&opt=ok&ucet="&ucet&"&de="&de&"&ate="&ate&"&r="&restricao)					
'		
'	end if	
'==========================================================================================================================
'Fim do código da função Emitir Bloqueto por Período
%>
<%

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
	if (document.getElementById('unidade').value == "999990")
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


 function CarregarUCET() { 
  document.getElementById('ucet').style.display='block';
  document.getElementById('envio').style.display='block';
  document.getElementById('mat').style.display='none';
}  

 function CarregarResult() { 
  document.getElementById('Result').style.display='block';
;
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
if enviado = "S" then
%>
          <tr> 
            
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,9714,2,0) %>
    </td>
          </tr> 
<%end if
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
            <td width="204" height="20" align="right">Selecione uma op&ccedil;&atilde;o para busca: </td><td width="31" height="20" align="center" valign="middle"> <input name="opcao" type="radio" disabled="disabled" id ="radio_mat" onClick="document.getElementById('mat').style.display='block';document.getElementById('envio').style.display='block';document.getElementById('ucet').style.display='none';" value="mat" ></td>
            <td width="72" valign="middle">Matr&iacute;cula            </td>
            <td width="22" valign="middle"><input name="opcao" type="radio" disabled="disabled" id ="radio_ucet" onClick="document.getElementById('ucet').style.display='block';document.getElementById('envio').style.display='block';document.getElementById('mat').style.display='none';" value="ucet" checked="checked" ></td>
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
                  <td width="250" align="center" class="tb_subtit">M&Ecirc;S</td>
                  <td width="250" class="tb_subtit"><div align="center">TIPO</div></td>
                </tr>
                <tr> 
                  <td width="250"> 
                    <div align="center"> 
                      <select name="unidade" class="select_style" id="unidade" onChange="recuperarCurso(this.value)">
                       
                       <%if isnull("unidade") or unidade = "" then%>
                        <option value="999990" selected></option>
                        <%	end if		
		
		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0
		
While not RS0.EOF
	NU_Unidade = RS0("NU_Unidade")
	NO_Abr = RS0("NO_Abr")
	unidade=unidade*1
	NU_Unidade=NU_Unidade*1
	
	if unidade=NU_Unidade then
		selected="selected"
	else
		selected=""	
	end if	
%>
                        <option value="<%response.Write(NU_Unidade)%>" <%response.Write(selected)%>> 
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
<%if isnull("curso") or curso = "" then%>
                        <select class="select_style">
                        </select>
<%else%>

                        
                         <select name="curso" class="borda" onChange="recuperarEtapa(this.value)">
<%
			Set RS0ue = Server.CreateObject("ADODB.Recordset")
			SQL0ue = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade
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
			
	NO_Curso = RS0c("NO_Abreviado_Curso")		
	

	if curso=CO_Curso then
		selected="selected"
	else
		selected=""	
	end if	
	%>
      <option value="<%response.Write(CO_Curso)%>" <%response.Write(selected)%>> 
      <%response.Write(NO_Curso)%>
      </option>
      <%
	
	CO_Curso_check = CO_Curso
	RS0ue.MOVENEXT
	end if
	WEND
end if
%>
                        </select>
                        
                        
                      </div>
                    </div></td>
                  <td width="250"> 
                    <div align="center"> 
                      <div id="divEtapa"> 
<%if isnull("co_etapa") or co_etapa = "" then%>
                        <select class="select_style">
                        </select>
<%else%>
                        <select name="etapa" class="borda" onChange="recuperarTurma(this.value)">
                          <%

		Set RS0e = Server.CreateObject("ADODB.Recordset")
		SQL0e = "SELECT * FROM TB_Unidade_Possui_Etapas where CO_Curso ='"& curso &"' AND NU_Unidade="& unidade 
		RS0e.Open SQL0e, CON0
			

while not RS0e.EOF
etapa= RS0e("CO_Etapa")

		Set RS3e = Server.CreateObject("ADODB.Recordset")
		SQL3e = "SELECT * FROM TB_Etapa where CO_Etapa ='"& etapa &"' And CO_Curso ='"& curso &"'" 
		RS3e.Open SQL3e, CON0
		

no_etapa=RS3e("NO_Etapa")

	if co_etapa=etapa then
		selected="selected"
	else
		selected=""	
	end if	
%>
                          <option value="<%response.Write(etapa)%>" <%response.Write(selected)%>> 
                          <%response.Write(no_etapa)%>
                          </option>
                          <%						
RS0e.MOVENEXT
WEND
%>
</select>
<%
end if
%>
                      </div>
                    </div></td>
                  <td width="250"> 
                    <div align="center"> 
                      <div id="divTurma"> 
                      
<%if isnull("turma") or turma = "" then%>                      
                        <select class="select_style">
                        </select>
                        
<%else
%>
                        <select name="turma" class="borda" onChange="recuperarTudo(this.value)">
                          <%
	
		Set RS0t = Server.CreateObject("ADODB.Recordset")
		SQL0t = "SELECT * FROM TB_Turma where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"' AND CO_Etapa='" & co_etapa & "' order by CO_Turma" 
		RS0t.Open SQL0t, CON0						
co_turma_check=9999990
while not RS0t.EOF
co_turma= RS0t("CO_Turma")

if co_turma = co_turma_check then
RS0t.MOVENEXT
else

	if co_turma=turma then
		selected="selected"
	else
		selected=""	
	end if	
%>
                          <option value="<%response.Write(co_turma)%>" <%response.Write(selected)%>> 
                          <%response.Write(co_turma)%>
                          </option>
                          <%						
end if

co_turma_check = co_turma
RS0t.MOVENEXT

WEND
%>
                        </select>
                        <%end if%>                        
                        
                        
                      </div>
                    </div></td>
                  <td width="250" align="center">               
                  <select name="select_mes" id="select_mes" class="select_style">
                  <% 
				  if mes_selecionado="" or isnull(mes_selecionado) then
				  	selected1 =""
					selected2 =""
					selected3 =""
					selected4 =""
					selected5 =""
					selected6 =""
					selected7 =""
					selected8 =""
					selected9 =""
					selected10 =""
					selected11 =""
					selected12 =""
				  else
					  mes_selecionado=mes_selecionado*1 
					  if mes_selecionado=1 then
					  	selected1 ="selected"
					  elseif mes_selecionado=2 then
					  	selected2 ="selected"	
					  elseif mes_selecionado=3 then
					  	selected3 ="selected"	
					  elseif mes_selecionado=4 then
					  	selected4 ="selected"	
					  elseif mes_selecionado=5 then
					  	selected5 ="selected"	
					  elseif mes_selecionado=6 then
					  	selected6 ="selected"	
					  elseif mes_selecionado=7 then
					  	selected7 ="selected"	
					  elseif mes_selecionado=8 then
					  	selected8 ="selected"	
					  elseif mes_selecionado=9 then
					  	selected9 ="selected"	
					  elseif mes_selecionado=10 then
					  	selected10 ="selected"	
					  elseif mes_selecionado=11 then
					  	selected11 ="selected"	
					  elseif mes_selecionado=12 then
					  	selected12 ="selected"																																																																		
					  end if
				  end if
				  %> 
                    <option value="1" <%response.Write(selected1)%>>Janeiro</option>
                    <option value="2" <%response.Write(selected2)%>>Fevereiro</option>       
                    <option value="3" <%response.Write(selected3)%>>Mar&ccedil;o</option>  
                    <option value="4" <%response.Write(selected4)%>>Abril</option>  
                    <option value="5" <%response.Write(selected5)%>>Maio</option>  
                    <option value="6" <%response.Write(selected6)%>>Junho</option>  
                    <option value="7" <%response.Write(selected7)%>>Julho</option>
                    <option value="8" <%response.Write(selected8)%>>Agosto</option>
                    <option value="9" <%response.Write(selected9)%>>Setembro</option>
                    <option value="10" <%response.Write(selected10)%>>Outubro</option>
                    <option value="11" <%response.Write(selected11)%>>Novembro</option>
                    <option value="12" <%response.Write(selected12)%>>Dezembro</option>
                    </select></td>
                  <td width="250" align="center">
                  <select name="restricao_ucet" id="restricao_ucet" class="select_style" disabled>
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
                  <td colspan="6">&nbsp;</td>
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
    <div id="Result" style="display: none;">	
<table width="1000" border="0" cellspacing="0" cellpadding="0">
  <tbody>
    <tr>
      <th width="50"scope="col" class="tb_subtit" align="left"><input type="checkbox" name="todos" value="Marcar Todos" onclick="this.value=check(this.form.matric)"></th>
      <th width="100" scope="col" class="tb_subtit">Matr&iacute;cula</th>
      <th width="350" scope="col" class="tb_subtit" align="left">Nome do Aluno</th>
      <th width="350" scope="col" class="tb_subtit" align="left">Nome do Respons&aacute;vel Financeiro</th>
      <th width="250" scope="col" class="tb_subtit" align="left">E-mail do do Respons&aacute;vel Financeiro</th>
    </tr>	
<%	
	alunos_vetor=alunos_turma(ano_letivo,unidade,curso,co_etapa,turma,"nome")	
if alunos_vetor="nulo" then
%>
    <tr>
      <td colspan="5" class="form_corpo">Não foram encontrados alunos para as condições especificadas</td>
    </tr>

<%
else	
	alunos=split(alunos_vetor,"#$#")
	total_email=0



	for a=0 to ubound(alunos)
		aluno=split(alunos(a),"#!#")	
		co_matric=aluno(0)	
		
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4= "SELECT * FROM TB_Posicao WHERE VA_Realizado=0 AND NO_Lancamento='Mensalidade' AND CO_Matricula_Escola ="& co_matric &" AND Mes = "&mes_selecionado
		RS4.Open SQL4, CON4			
		
		if not RS4.EOF then
		
			Set RSa = Server.CreateObject("ADODB.Recordset")
			SQLa = "SELECT TB_Alunos.NO_Aluno, TB_Alunos.TP_Resp_Fin, TB_Alunos.IN_Sexo, TB_Matriculas.NU_Unidade, TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma FROM TB_Alunos, TB_Matriculas where TB_Matriculas.CO_Matricula = TB_Alunos.CO_Matricula AND TB_Matriculas.CO_Matricula = "& co_matric &" AND TB_Matriculas.NU_Ano = "& ano_letivo
			RSa.Open SQLa, CON1
			
	
			
			nome_aluno = RSa("NO_Aluno")
			tp_resp_fin = RSa("TP_Resp_Fin")
			in_sexo = RSa("IN_Sexo")		
			nu_unidade = RSa("NU_Unidade")
			co_curso = RSa("CO_Curso")
			co_etapa = RSa("CO_Etapa")
			co_turma = RSa("CO_Turma")
			
			if in_sexo = "F" then
				designacao = "o"
			else
				designacao = "a"		
			end if	
		
			
			Set RSc = Server.CreateObject("ADODB.Recordset")
			SQLc = "SELECT NO_Contato,CO_CPF_PFisica, TX_EMail FROM TB_Contatos where CO_Matricula = "& co_matric &" AND TP_Contato = '"& tp_resp_fin&"'"
			RSc.Open SQLc, CON6		
			
			If RSc.EOF then
				nome_responsavel ="<font color = #FF0000> O(A) "&tp_resp_fin&" d"&designacao&" alun"&designacao&" não est&aacute; cadastrado.</font>"
			else
				nome_responsavel = RSc("NO_Contato")
				cpf_resp = RSc("CO_CPF_PFisica")
				email_resp =RSc("TX_EMail")
				
				if isnull(nome_responsavel) or nome_responsavel="" then
					nome_responsavel ="<font color = #FF0000>Nome não cadastrado para o(a) "&tp_resp_fin&" d"&designacao&" alun"&designacao&".</font>"
				end if	
						
				if cpf_resp = "" or isnull(cpf_resp) then
				
				else
					cpf_resp = replace(cpf_resp,"-","")
					cpf_resp = replace(cpf_resp,".","")				
				end if
				
				if isnull(email_resp) or email_resp="" then
					email_resp ="<font color = #FF0000>Email não cadastrado para o respons&aacute;vel financeiro</font>"
				end if				
			end if		
		
		
		
	
				'Set RSF = Server.CreateObject("ADODB.Recordset")
'				'sqlF = "select CO_Usuario from TB_RespxAluno where CO_Aluno="&co_matric&" and (TP_Resp='F' or TP_Resp='P')"
'				sqlF = "select CO_Usuario from TB_RespxAluno where CO_Aluno="&co_matric&" and TP_Resp='F'"	
'				set RSF = CON_wf.Execute (sqlF)		
'				
'				if RSF.EOF then
'					nome_responsavel = "<font color = #FF0000>Respons&aacute;vel Financeiro não cadastrado</font>"
'				else
'				resp=RSF("CO_Usuario")
'				
'					Set RSFM = Server.CreateObject("ADODB.Recordset")
'					sqlFM = "select NO_Usuario, TX_EMail_Usuario,IN_Aut_email from TB_Usuario where CO_Usuario="&resp
'					set RSFM = CON_wf.Execute (sqlFM)					
'	
'					if RSFM.EOF then
'					
'					else	
'						nome_responsavel = 	RSFM("NO_Usuario")				
'						email_resp=RSFM("TX_EMail_Usuario")
'						aut_resp=RSFM("IN_Aut_email")
'						
'						if aut_resp=TRUE then
'							if total_email=0 then
'								publico=email_resp
'					
'							else
'								publico=publico&","&email_resp
'							end if
'							total_email=total_email+1	
'						end if	
'					end if	
'				end if	
%>
    <tr>
      <td class="form_dado_texto"><input type="checkbox" name="matric" value="<%response.write(co_matric)%>"></td>
      <td class="form_dado_texto" align="center"><%response.write(co_matric)%></td>
      <td class="form_dado_texto"><%response.write(aluno(2))%></td>
      <td class="form_dado_texto"><%response.write(resp&" "&nome_responsavel)%></td>
      <td class="form_dado_texto"><%response.write(email_resp)%></td>
    </tr>


<%		End IF
	Next
end if	
	%>	
    <tr>
      <td colspan="5" class="form_corpo"><hr/></td>
    </tr>     				  
    <tr>
      <td class="form_corpo">&nbsp;</td>    
      <td class="form_corpo">&nbsp;</td>   
      <td class="form_corpo">&nbsp;</td>   
      <td class="form_corpo">&nbsp;</td>                    
      <td class="form_corpo"><center><input type="Submit" class="botao_prosseguir" id="SubmitEmail" name="Submit" value="Enviar Email"></center></td>
    </tr>    


  </tbody>
</table>
</div>
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