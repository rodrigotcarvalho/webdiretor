<%	'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%
opt = request.QueryString("opt")

ano_letivo_wf = session("ano_letivo_wf")
session("ano_letivo_wf")=ano_letivo_wf
co_usr = session("co_user")
nivel=4

Session("dia_de")=""
Session("dia_de")=""
Session("dia_ate")=""
Session("mes_ate")=""
Session("unidade")=""
Session("curso")=""
Session("etapa")=""
Session("turma")=""
Session("arquivos_desanexados")="nulo" 

if isnull(Session("arquivos_anexados")) then

elseif (Session("arquivos_anexados")<>"nulo" and Session("arquivos_anexados")<>"") then


	SET FSO = Server.CreateObject("Scripting.FileSystemObject")
	
	Set pasta = FSO.GetFolder(CAMINHO_upload)
	Set arquivos = pasta.Files

	for each apagarquivo in arquivos
		data_arquivo =apagarquivo.DateLastModified
		nome_arquivo =apagarquivo.Name
		'response.Write(DatePart("n",Now())&"<BR>")	
		hora=DatePart("h",Now())
		min=DatePart("n",Now())
		hora_arquivo=DatePart("h",data_arquivo)
		min_arquivo=DatePart("n",data_arquivo)
		if (hora_arquivo<hora and min>30) or (min-min_arquivo>30) then
			FSO.deletefile(apagarquivo) 
		end if	
	next
	anexos=split(Session("arquivos_anexados"),"#!#")
	for atch=0 to ubound(anexos)	
		arquivo = CAMINHO_upload & anexos(atch)
		FSO.deletefile(arquivo) 
	Next	
	Session("arquivos_anexados")="nulo" 
end if	


nvg = request.QueryString("nvg")
chave=nvg
session("chave")=chave
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano_info=nivel&"-"&chave&"-"&ano_letivo_wf



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
                         <script>
// A fun��o abaixo pega a vers�o mais nova do xmlhttp do IE e verifica se � Firefox. Funciona nos dois.
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
                                                            alert("Esse browser n�o tem recursos para uso do Ajax");
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
// Cria��o do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicita��o HTTP. O primeiro par�metro informa o m�todo post/get
// O segundo par�metro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicita��o s�ncrona, o par�metro deve ser false
                                               oHTTPRequest.open("post", "executa.asp?opt=c", true);
// Para solicita��es utilizando o m�todo post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A fun��o abaixo � executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto j� completou a solicita��o
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto � gerado no arquivo executa.asp e colocado no div
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
document.all.divEtapa.innerHTML ="<select name='etapa' class=select_style><option value='nulo' selected></option></select>"
document.all.divTurma.innerHTML = "<select name='turma' class=select_style><option value='nulo' selected></option></select>"
//recuperarEtapa()
                                                           }
                                               }
// Abaixo � enviada a solicita��o. Note que a configura��o
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {
// Cria��o do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicita��o HTTP. O primeiro par�metro informa o m�todo post/get
// O segundo par�metro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicita��o s�ncrona, o par�metro deve ser false
                                               oHTTPRequest.open("post", "executa.asp?opt=e", true);
// Para solicita��es utilizando o m�todo post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A fun��o abaixo � executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto j� completou a solicita��o
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto � gerado no arquivo executa.asp e colocado no div
                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select  name='turma' class=select_style><option value='nulo' selected></option></select>"
//recuperarTurma()
                                                           }
                                               }
// Abaixo � enviada a solicita��o. Note que a configura��o
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {
// Cria��o do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicita��o HTTP. O primeiro par�metro informa o m�todo post/get
// O segundo par�metro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicita��o s�ncrona, o par�metro deve ser false
                                               oHTTPRequest.open("post", "executa.asp?opt=t", true);
// Para solicita��es utilizando o m�todo post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A fun��o abaixo � executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto j� completou a solicita��o
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto � gerado no arquivo executa.asp e colocado no div
                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t																	   
                                                           }
                                               }
// Abaixo � enviada a solicita��o. Note que a configura��o
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
						 function recuperarMensagem(mTipo)
                                   {
// Cria��o do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicita��o HTTP. O primeiro par�metro informa o m�todo post/get
// O segundo par�metro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicita��o s�ncrona, o par�metro deve ser false
                                               oHTTPRequest.open("post", "executa.asp?opt=msg", true);
// Para solicita��es utilizando o m�todo post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A fun��o abaixo � executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto j� completou a solicita��o
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto � gerado no arquivo executa.asp e colocado no div
                                                                       var resultado_M= oHTTPRequest.responseText;
resultado_M = resultado_M.replace(/\+/g," ")
resultado_M = unescape(resultado_M)
document.all.divMensagem.innerHTML = resultado_M																	   
                                                           }
                                               }
// Abaixo � enviada a solicita��o. Note que a configura��o
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("m_pub=" + mTipo);
                                   }								   
                        
	   
						 function desanexa(aTipo)
                                   {
// Cria��o do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicita��o HTTP. O primeiro par�metro informa o m�todo post/get
// O segundo par�metro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicita��o s�ncrona, o par�metro deve ser false
                                               oHTTPRequest.open("post", "executa.asp?opt=danx", true);
// Para solicita��es utilizando o m�todo post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A fun��o abaixo � executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto j� completou a solicita��o
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto � gerado no arquivo executa.asp e colocado no div
//                                                                      var resultado_A= oHTTPRequest.responseText;
//resultado_A = resultado_A.replace(/\+/g," ")
//resultado_A = unescape(resultado_A)
//document.all.divAnexo.innerHTML = resultado_A																	   
                                                           }
                                               }
// Abaixo � enviada a solicita��o. Note que a configura��o
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("a_pub=" + aTipo);
                                   }
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
                         </script>
						 
 <%
		Set RS1t = Server.CreateObject("ADODB.Recordset")
		SQL1t = "SELECT count('CO_Email') as total FROM TB_Email_Mensagem where CO_Email<=5"
		RS1t.Open SQL1t, CON0
		if RS1t("total") =1 then		
			Set RS1c = Server.CreateObject("ADODB.Recordset")
			SQL1c = "SELECT CO_Email FROM TB_Email_Mensagem where CO_Email<=5"
			RS1c.Open SQL1c, CON0	
			
				
			onload = ";recuperarMensagem("&RS1c("CO_Email")&")"
		else
			onload = ""		
		end if			
	%>
                        
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
</head>

<body leftmargin="0"  topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../../img/menu_r1_c2_f3.gif','../../../../img/menu_r1_c2_f2.gif','../../../../img/menu_r1_c2_f4.gif','../../../../img/menu_r1_c4_f3.gif','../../../../img/menu_r1_c4_f2.gif','../../../../img/menu_r1_c4_f4.gif','../../../../img/menu_r1_c6_f3.gif','../../../../img/menu_r1_c6_f2.gif','../../../../img/menu_r1_c6_f4.gif','../../../../img/menu_r1_c8_f3.gif','../../../../img/menu_r1_c8_f2.gif','../../../../img/menu_r1_c8_f4.gif','../../../../img/menu_direita_r2_c1_f3.gif','../../../../img/menu_direita_r2_c1_f2.gif','../../../../img/menu_direita_r2_c1_f4.gif','../../../../img/menu_direita_r4_c1_f3.gif','../../../../img/menu_direita_r4_c1_f2.gif','../../../../img/menu_direita_r4_c1_f4.gif','../../../../img/menu_direita_r6_c1_f3.gif','../../../../img/menu_direita_r6_c1_f2.gif','../../../../img/menu_direita_r6_c1_f4.gif')<%response.Write(onload)%>">
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
<%if opt="ok" then%>      
      <tr>                
    <td height="10"> 
      <%	call mensagens(4,709,2,0) 
%>
	</td>
	</tr> 
<%end if%>         
      <tr>                
    <td height="10"> 
      <%	call mensagens(4,9706,0,0) 
%>
	</td>
	</tr>
<tr>

    <td valign="top"> 
		<%
mes = DatePart("m", now) 
dia = DatePart("d", now) 



dia=dia*1
mes=mes*1
%>				

                
      <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
        <tr class="tb_tit"> 
          <td width="653" height="15" class="tb_tit">Informe os crit&eacute;rios 
              para pesquisa 
            </td>
        </tr>
        <tr> 
          <td valign="top"><FORM name="formulario" METHOD="POST" ACTION="email.asp"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="338" valign="top"><table width="97%" border="0" align="right" cellpadding="0" cellspacing="0">
                      <tr>
                        <td class="tb_subtit">Informe o Assunto:<input name="co_grupo" type="hidden" id="co_grupo" value="<% = co_grupo %>"></td>
                      </tr>
                      <tr>
                        <td><select name="assunto" class="select_style_fixo_1" id="assunto" >
                          <%
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Email_Assunto where CO_Assunto=1 order by CO_Assunto"
		RS1.Open SQL1, CON0

if RS1.eof then %>

<%else
	while not RS1.EOF
		co_assunto=RS1("CO_Assunto")
		assunto=RS1("TX_Titulo_Assunto")
		assunto_padrao=RS1("IN_Assunto_Padrao")
		
		if assunto_padrao=TRUE then
			assunto_selected="SELECTED"
		ELSE
			assunto_selected=""
		END IF	
		
		%>
							  <option value="<%response.Write(co_assunto)%>" <%response.Write(assunto_selected)%>>
								<%response.Write(assunto)%>
								</option>
							  <%

	RS1.movenext
	Wend
end if	



%>
                        </select></td>
                      </tr>
                      <tr>
                        <td height="5"></td>
                      </tr>
                      <tr>
                        <td><span class="tb_subtit">Com c&oacute;pia para:</span></td>
                      </tr>
                      <tr>
                        <td><select name="cc" class="select_style_fixo_1" id="cc" >
                          <%
'		Set RS1 = Server.CreateObject("ADODB.Recordset")
'		SQL1 = "SELECT Login FROM TB_Operador"
'		RS1.Open SQL1, CON
'qtd_cc=0
'while not RS1.EOF
'	email=RS1("Login")
'	
'	if qtd_cc=0then
'		cc_selected="SELECTED"
'		qtd_cc=1
'	ELSE
'		cc_selected=""
'	END IF	
'	
'	%>
<!--                          <option value="<%response.Write(email)%>" <%response.Write(cc_selected)%>>
                            <%response.Write(email)%>
                            </option>
-->                         <%
'RS1.movenext
'Wend

		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT DISTINCT(TX_email_WF) FROM TB_Etapa"
		RS2.Open SQL2, CON0


while not RS2.EOF
	email=RS2("TX_email_WF")
	
	%>
                          <option value="<%response.Write(email)%>">
                            <%response.Write(email)%>
                            </option>
                          <%
RS2.movenext
Wend
%>
                          </select></td>
                      </tr>
                      <tr>
                        <td>&nbsp;</td>
                      </tr>
                      <tr>
                        <td></td>
                      </tr>
                      <tr>
                        <td></td>
                      </tr>
                    </table></td>
                    <td width="662" valign="top"><table width="100%" border="0" align="right" cellpadding="0" cellspacing="0">
                      <tr>
                        <td class="tb_subtit">&nbsp;Informe o conte&uacute;do da mensagem:</td>
                      </tr>
                      <tr>
                        <td><select name="mensagem" class="select_style_fixo_1" id="mensagem" onChange="recuperarMensagem(this.value)">
                          <%					
		Set RS1b = Server.CreateObject("ADODB.Recordset")
		SQL1b = "SELECT * FROM TB_Email_Mensagem where CO_Email<=5 order by CO_Email"
		RS1b.Open SQL1b, CON0		

if RS1b.EOF then

else
	while not RS1b.EOF
		co_email=RS1b("CO_Email")
		co_msg=RS1b("NO_Email")
		msg=RS1b("TX_Conteudo_Email")
		msg_padrao=RS1b("IN_Email_Padrao")
		
'		if co_msg<10 then
'			espacador="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
'		elseif co_msg<100 then
'			espacador="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"	
'		else	
			espacador="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"		
'		end if	
		msg_combo="Mensagem "&co_msg&espacador
		
		if msg_padrao=TRUE then
			msg_selected="SELECTED"
			conteudo_email=msg
		ELSE
			msg_selected=""
		END IF	
		
		%>
							  <option value="<%response.Write(co_email)%>" <%response.Write(msg_selected)%>>
								<%response.Write(msg_combo)%>
								</option>
							  <%
	
	RS1b.movenext
	Wend
end if	
%>
                        </select></td>
                      </tr>
                      <tr>
                        <td height="5"></td>
                      </tr>
                      <tr>
                        <td><div id="divMensagem">
                          <textarea name="msg" cols="125" rows="8" class="borda"><%response.write(conteudo_email)%>
                        </textarea>
                        </div></td>
                      </tr>
                    </table></td>
                  </tr>
                  <tr>
                    <td colspan="2"></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td colspan="2">&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td colspan="4" class="tb_subtit"><table width="988" border="0" align="right" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="180" class="tb_subtit">Selecione os Destinat&aacute;rios:</td>
<!--                     <td width="25"><input name="dest" type="checkbox" class="borda" id="dest" value="a"></td>
                   <td width="60" class="form_dado_texto">Alunos</td>-->
                    <td width="25"><input name="dest" type="checkbox" class="borda" id="dest" value="r"></td>
                    <td width="100" class="form_dado_texto">Respons&aacute;veis</td>
                    <td width="25"><input name="dest" type="checkbox" class="borda" id="dest" value="i"></td>
                    <td class="form_dado_texto">Contatos</td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td width="247" class="tb_subtit"><div align="center">UNIDADE </div></td>
                <td width="247" class="tb_subtit"><div align="center">CURSO </div></td>
                <td width="247" class="tb_subtit"><div align="center">ETAPA </div></td>
                <td width="247" class="tb_subtit"><div align="center">TURMA </div></td>
              </tr>
              <tr>
                <td width="247"><div align="center">
                  <select name="unidade" class="select_style" onChange="recuperarCurso(this.value)">
                    <option value="nulo" selected></option>
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
                <td width="247"><div align="center">
                  <div id="divCurso">
                    <select name="curso" class="select_style">
                    <option value="nulo" selected></option>                    
                    </select>
                  </div>
                </div></td>
                <td width="247"><div align="center">
                  <div id="divEtapa">
                    <select name="etapa" class="select_style">
                    <option value="nulo" selected></option>                    
                    </select>
                  </div>
                </div></td>
                <td width="247"><div align="center">
                  <div id="divTurma">
                    <select name="turma" class="select_style">
                    <option value="nulo" selected></option>                    
                    </select>
                  </div>
                </div></td>
              </tr>
            </table>
</form>
          </td>
        </tr>
        <tr>
          <td valign="top">	<table width="988" border="0" align="right" cellpadding="0" cellspacing="0">
  <tr>
    <td>   
</td>
  </tr>
</table>

 </td>
        </tr>
        <tr>
          <td valign="top">&nbsp;</td>
        </tr>
        <tr>
          <td valign="top"><hr width="1000"></td>
        </tr>
        <tr>
          <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="33%">&nbsp;</td>
                <td width="34%">&nbsp;</td>
                <td width="33%"><div align="center"><input name="SUBMIT" type=SUBMIT class="botao_prosseguir" onClick="MM_callJS('submitforminterno()')" value="Enviar"  ></div></td>
              </tr>
          </table></td>
        </tr>
<!--        <tr>
          <td valign="top"><iframe src="aspuploader/form-multiplefiles.asp" frameborder ="0" width="100%" height="500" scrolling="no"> </iframe></td>
        </tr>-->
        <tr>
          <td valign="top">&nbsp;</td>
        </tr>
        <tr>
          <td valign="top">&nbsp;</td>
        </tr>
      </table>
      </td>
          </tr>
		  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
        </table>
 <script type="text/javascript">
    function setUnidade(p_val){

      document.forms[3].unidade.options[1].selected = "true";
      recuperarCurso(p_val);

    } 
  setUnidade(1);
 </script> 
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
response.redirect("../../../../inc/erro.asp")
end if
%>