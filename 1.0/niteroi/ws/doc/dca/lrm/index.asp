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

unidade_form=request.form("unidade")


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

	response.Redirect("../../../../relatorios/swd057.asp?un="&unidade_form)
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
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
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
  if (document.busca.unidade.value == "nulo")
  {    alert("Por favor selecione uma unidade!")
	var combo = document.getElementById("unidade");
	combo.options[0].selected = "true";
    document.busca.unidade.focus()
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
 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
</script>
</head>
<body background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
         <form action="index.asp?opt=search&nvg=<%response.Write(nvg)%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()">   
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
<%if opt="search" then%>
            <tr>              
    <td height="10" colspan="5"> 
      <%call mensagens(nivel,mensagem,1,0) %>
    </td>
			   </tr>          
<%
end if
%>
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,710,0,0) %>
    </td>
			  </tr>	
           	  
          <tr class="tb_tit">             
      <td height="10" colspan="5">Selecione a Unidade</td>
          </tr>
          <tr>                   
            <td height="10" colspan="5" valign="top">  
  <table width="100%" border="0" cellspacing="0">
<!--    <tr> 
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
-->    <tr> 
      <td width="250">&nbsp;</td>
      <td width="250">        <div align="right" class="form_dado_texto">Unidade:
          </div></td>
      <td width="250"><div align="left">
        <select name="unidade" class="select_style" id="unidade">
          <option value="nulo" selected></option>
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
      <td width="250">&nbsp;</td>
      </tr>
    <tr>
      <td height="15" colspan="4" bgcolor="#FFFFFF"><hr></td>
      </tr>
    <tr> 
      <td width="250" height="15" bgcolor="#FFFFFF"><font size="2" face="Arial, Helvetica, sans-serif">
        <div align="center">
          <input name="Submit2" type="button" class="botao_cancelar" id="Submit2" onClick="MM_goToURL('parent','../../../index.asp?nvg=WS');return document.MM_returnValue" value="Cancelar">

        </div>
      </font></td>
      <td width="250" height="15" bgcolor="#FFFFFF"></td>
      <td width="250" height="15" bgcolor="#FFFFFF"></td>
      <td width="250" height="15" bgcolor="#FFFFFF"><font size="2" face="Arial, Helvetica, sans-serif">
        <div align="center"><input name="Submit" type="submit" class="botao_prosseguir" id="Submit" value="Prosseguir"></div>
        </font></td>
      </tr>
    </table>
              
            </td>
          </tr>
  
                <tr>                   
    <td colspan="5" valign="top"> 
    </td>
  </tr>             
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.gif" width="1000" height="40"></td>
  </tr>
</table>
  </form> 
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