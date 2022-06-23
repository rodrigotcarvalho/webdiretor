<%On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->

<!--#include file="../../../../inc/caminhos.asp"-->

<!--#include file="../../../../inc/funcoes2.asp"-->

<%

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
nvg=request.QueryString("nvg")
opt = request.QueryString("opt")
ori = request.QueryString("ori")
chave=nvg
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
trava = session("trava") 
cod_cons = request.QueryString("cod_cons")
		
		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
				Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR		

 call VerificaAcesso (CON,chave,nivel)
autoriza=Session("autoriza")




 call navegacao (CON,chave,nivel)
navega=Session("caminho")

if ori="2" or ori="3" then	
		
		
codigo = request.querystring("cod")
nome_prof = request.querystring("nome")
apelido = request.querystring("apelido")
sexo = request.querystring("sexo")
nasce= request.querystring("nasce")		
rua = request.querystring("rua")
numero = request.querystring("numero")
complemento = request.querystring("complemento")
bairro= request.querystring("bairro")
municipio= request.querystring("ciddom")
pais = request.querystring("pais")
uf= request.querystring("estadodom")
cep = request.querystring("cep")
telefone = request.querystring("telefones")
uf_natural = request.querystring("estadonat")
nacionalidade = request.querystring("nacionalidade")
natural = request.querystring("cidadenat")
email = request.querystring("email")

ativo = request.querystring("ativo")		

if ativo="sim" then
	ativo="True"
else
	ativo="False"
End if	

pais = pais*1
nacionalidade = nacionalidade*1
municipio = municipio*1
bairro = bairro*1
natural = natural*1		
elseif ori="01" then				

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Professor WHERE CO_Professor ="& cod_cons
		RS.Open SQL, CON1

cod_cons = RS("CO_Professor")
nome_prof = RS("NO_Professor")
apelido = RS("NO_Apelido_Professor")
sexo = RS("IN_Sexo")
nascimento = RS("DA_Nascimento")

vetor_nascimento = Split(nascimento,"/")  
dia = vetor_nascimento(0)
mes = vetor_nascimento(1)
ano = vetor_nascimento(2)

if dia<10 then 
dia = "0"&dia
end if

if mes<10 then
mes = "0"&mes
end if
dia_a = dia
mes_a = mes
ano_a = ano

nasce = dia&"/"&mes&"/"&ano

rua = RS("NO_Logradouro")
numero = RS("NU_Logradouro")
complemento = RS("TX_Complemento_Logradouro")
bairro= RS("CO_Bairro")
municipio= RS("CO_Municipio")
pais = RS("CO_Pais")
uf= RS("SG_UF")
cep = RS("CO_CEP")
telefone = RS("NUS_Telefones")
uf_natural = RS("SG_Estado_Natural")
nacionalidade = RS("CO_Nacionalidade")
natural = RS("CO_Municipio_Natural")
email = RS("TX_EMail")
ativo = RS("IN_Ativo_Escola")
co_usr_prof = RS("CO_Usuario")

'response.Write(ativo)
else
natural = 6001
pais = 10
uf = "RJ"
municipio = 6001
uf_natural = "RJ"
nacionalidade = 1
natural = 6001
end if

if isnull(pais) then 
pais = 10
end if

if isnull(uf) then 
uf = "RJ"
end if

if isnull(municipio) then 
municipio = 6001
end if

if isnull(uf_natural) then 
uf_natural = "RJ"
end if

if isnull(nacionalidade) then 
nacionalidade = 1
end if

if rua = "nulo" then 
rua = ""
end if

if complemento = "nulo" then 
complemento = ""
end if

	%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../../../../js/global.js"></script>
<script type="text/javascript" src="../../../../js/atualiza_select.js"></script>
<script language="JavaScript" type="text/JavaScript">
function formatar(src, mask)
{
  var i = src.value.length;
  var saida = mask.substring(0,1);
  var texto = mask.substring(i)
if (texto.substring(0,1) != saida)
  {
        src.value += texto.substring(0,1);
  }
}
function check_date(field){
var checkstr = "0123456789";
var DateField = field;
var Datevalue = "";
var DateTemp = "";
var seperator = ".";
var day;
var month;
var year;
var leap = 0;
var err = 0;
var i;
   err = 0;
   DateValue = DateField.value;
   /* Delete all chars except 0..9 */
   for (i = 0; i < DateValue.length; i++) {
	  if (checkstr.indexOf(DateValue.substr(i,1)) >= 0) {
	     DateTemp = DateTemp + DateValue.substr(i,1);
	  }
   }
   DateValue = DateTemp;
   /* Always change date to 8 digits - string*/
   /* if year is entered as 2-digit / always assume 20xx */
   if (DateValue.length == 6) {
      DateValue = DateValue.substr(0,4) + '20' + DateValue.substr(4,2); }
   if (DateValue.length != 8) {
      err = 19;}
   /* year is wrong if year = 0000 */
   year = DateValue.substr(4,4);
   if (year == 0) {
      err = 20;
   }
   /* Validation of month*/
   month = DateValue.substr(2,2);
   if ((month < 1) || (month > 12)) {
      err = 21;
   }
   /* Validation of day*/
   day = DateValue.substr(0,2);
   if (day < 1) {
     err = 22;
   }
   /* Validation leap-year / february / day */
   if ((year % 4 == 0) || (year % 100 == 0) || (year % 400 == 0)) {
      leap = 1;
   }
   if ((month == 2) && (leap == 1) && (day > 29)) {
      err = 23;
   }
   if ((month == 2) && (leap != 1) && (day > 28)) {
      err = 24;
   }
   /* Validation of other months */
   if ((day > 31) && ((month == "01") || (month == "03") || (month == "05") || (month == "07") || (month == "08") || (month == "10") || (month == "12"))) {
      err = 25;
   }
   if ((day > 30) && ((month == "04") || (month == "06") || (month == "09") || (month == "11"))) {
      err = 26;
   }
   /* if 00 ist entered, no error, deleting the entry */
   if ((day == 0) && (month == 0) && (year == 00)) {
      err = 0; day = ""; month = ""; year = ""; seperator = "";
   }
   /* if no error, write the completed date to Input-Field (e.g. 13.12.2001) */
   if (err == 0) {
      DateField.value = day + seperator + month + seperator + year;
   }
   /* Error-message if err != 0 */
   else {
      alert("Date is incorrect!");
      DateField.select();
	  DateField.focus();
   }
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
}  function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
function checksubmit()
{
  if (document.inclusao.nome.value == "")
  {    alert("Por favor, digite um nome para o professor!")
    document.inclusao.nome.focus()
    return false
  }
erro=0;
        hoje = new Date();
         anoAtual = hoje.getFullYear();
         barras = inclusao.nasce.value.split("/");
         if (barras.length == 3){
                   dia = barras[0];
                   mes = barras[1];
                   ano = barras[2];
                   resultado = (!isNaN(dia) && (dia > 0) && (dia < 32)) && (!isNaN(mes) && (mes > 0) && (mes < 13)) && (!isNaN(ano) && (ano.length == 4) && (ano <= anoAtual && ano >= 1900));
                   if (!resultado) {
                             alert("Formato de data invalido!");
                             inclusao.nasce.focus();
                             return false;
                   }
         } else {
                   alert("Formato de data invalido!");
                   inclusao.nasce.focus();
                   return false;
         }
  if (document.inclusao.sexo.value == "0")
  {    alert("Por favor, escolha o sexo do professor!")
    document.inclusao.sexo.focus()
    return false
  }   
  if (document.inclusao.rua.value == "")
  {    alert("Por favor, digite a rua onde o professor reside!")
    document.inclusao.rua.focus()
    return false
  }    

erro=0;

         barras = inclusao.cep.value.split("-");
         if (barras.length == 2){
                   cep0= barras[0];
                   cep1 = barras[1];
                   resultado = (!isNaN(dia) && (cep0 > 10000) && (cep0 < 999999)) && (!isNaN(mes) && (cep1 >= 0) && (cep1 < 999));
                   if (!resultado) {
                             alert("Formato do CEP invalido!");
                             inclusao.cep.focus();
                             return false;
                   }
         } else {
                   alert("Formato do CEP invalido!");
                   inclusao.cep.focus();
                   return false;
         }

  if (document.inclusao.telefones.value == "")
  {    alert("Por favor, digite pelo menos um telefone para contato com o professor!")
    document.inclusao.telefones.focus()
    return false
  }                  	     
  return true

}
function MM_showHideLayers() { //v6.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
    obj.visibility=v; }
}

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

						 function recuperarCidNat(estadonat)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/cid_bairro.asp?opt=c&o=n&f=n", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_cid_nat  = oHTTPRequest.responseText;
resultado_cid_nat = resultado_cid_nat.replace(/\+/g," ")
resultado_cid_nat = unescape(resultado_cid_nat)
document.all.cid_nat.innerHTML =resultado_cid_nat
                                                           }
                                               }
                                               oHTTPRequest.send("c_pub=" + estadonat);
                                   }
						 function recuperarCidRes(estadores)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/cid_bairro.asp?opt=c&o=r&f=n", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_cid_res  = oHTTPRequest.responseText;
resultado_cid_res = resultado_cid_res.replace(/\+/g," ")
resultado_cid_res = unescape(resultado_cid_res)
document.all.cid_res.innerHTML =resultado_cid_res
document.all.bairro_res.innerHTML ="<select class=select_style></select>"
                                                           }
                                               }
                                               oHTTPRequest.send("c_pub=" + estadores);
                                   }

						 function recuperarBairroRes(estadores,cidres)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/cid_bairro.asp?opt=b&o=r&f=n", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_bairro_res  = oHTTPRequest.responseText;
resultado_bairro_res = resultado_bairro_res.replace(/\+/g," ")
resultado_bairro_res = unescape(resultado_bairro_res)
document.all.bairro_res.innerHTML =resultado_bairro_res
                                                           }
                                               }
                                               oHTTPRequest.send("c_pub=" + estadores +"&b_pub=" + cidres);
                                   }
function centraliza(w,h){
//o 120 e o 16 se referem ao tamanho di cabeçalho do navegador e a barra de rolagem respectivamente
    x = parseInt((screen.width - w - 16)/2);
    y = parseInt((screen.height - h - 120)/2);
   //alert(x + '\n' + y);
    document.getElementById('alinha').style.left = x;
    document.getElementById('alinha').style.top = y;
	
//	alert('w '+x +' h '+ y)
}								   
//-->
</script>
</head> 
<%call cabecalho(nivel)%>
<%if erro ="dt" then%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('document.alteracao.nasce.focus()');" >
<div id="fundo" style="position:absolute; left:0px; top:0px; width:100%; height:100%; z-index:1; background-color: #000000; layer-background-color: #000000; border: 1px none #000000; visibility: hidden;" class="transparente" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')"></div>
<div id="alinha" style="position:absolute; width:400px; visibility: hidden; z-index: 2; left: 326px; height: 520px;" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')"> 
  <table width="100%" border="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr> 
      <td width="478"> <div align="right"> <span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')">fechar</a></font></span></div></td>
      <td width="20"><div align="right"><span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')"><img src="../../../../img/fecha.gif" width="20" border="0"></a></font></span></div></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center"><img src="../../../../img/fotos/professor/<%response.Write(cod_cons)%>.jpg" height="500"></div></td>
    </tr>
    <tr>
      <td colspan="2"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
          <%response.Write(nome_prof)%>
          </font></div></td>
    </tr>
  </table>
</div>
<%elseif erro ="nb" then%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" background="../../../../img/fundo.gif" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('document.alteracao.numero.focus()');" >
<div id="fundo" style="position:absolute; left:0px; top:0px; width:100%; height:100%; z-index:1; background-color: #000000; layer-background-color: #000000; border: 1px none #000000; visibility: hidden;" class="transparente" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')"></div>
<div id="alinha" style="position:absolute; width:400px; visibility: hidden; z-index: 2; left: 326px; height: 520px;" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')"> 
  <table width="100%" border="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr> 
      <td width="478"> <div align="right"> <span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')">fechar</a></font></span></div></td>
      <td width="20"><div align="right"><span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')"><img src="../../../../img/fecha.gif" width="20" border="0"></a></font></span></div></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center"><img src="../../../../img/fotos/professor/<%response.Write(cod_cons)%>.jpg" height="500"></div></td>
    </tr>
    <tr>
      <td colspan="2"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
          <%response.Write(nome_prof)%>
          </font></div></td>
    </tr>
  </table>
</div>
<%elseif erro ="cp" then%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" background="../../../../img/fundo.gif" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('document.alteracao.cep.focus()');" >

<div id="fundo" style="position:absolute; left:0px; top:0px; width:100%; height:100%; z-index:1; background-color: #000000; layer-background-color: #000000; border: 1px none #000000; visibility: hidden;" class="transparente" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')"></div>
<div id="alinha" style="position:absolute; width:400px; visibility: hidden; z-index: 2; left: 326px; height: 520px;" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')"> 
  <table width="100%" border="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr> 
      <td width="478"> <div align="right"> <span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')">fechar</a></font></span></div></td>
      <td width="20"><div align="right"><span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')"><img src="../../../../img/fecha.gif" width="20" border="0"></a></font></span></div></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center"><img src="../../../../img/fotos/professor/<%response.Write(cod_cons)%>.jpg" height="500"></div></td>
    </tr>
    <tr>
      <td colspan="2"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
          <%response.Write(nome_prof)%>
          </font></div></td>
    </tr>
  </table>
</div>
<%else
if ori=02 then
focus="inclusao"
elseif ori=01 then
focus="alteracao"
end if
%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" background="../../../../img/fundo.gif" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('document.<%response.Write(focus)%>.nome.focus()');" >

<div id="fundo" style="position:absolute; left:0px; top:0px; width:100%; height:100%; z-index:1; background-color: #000000; layer-background-color: #000000; border: 1px none #000000; visibility: hidden;" class="transparente" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')"></div>
<div id="alinha" style="position:absolute; width:500px; visibility: hidden; z-index: 2; height: 536px;" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')"> 
  <table width="100%" border="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr> 
      <td width="478"> <div align="right"> <span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')">fechar</a></font></span></div></td>
      <td width="20"><div align="right"><span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')"><img src="../../../../img/fecha.gif" width="20" border="0"></a></font></span></div></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center" ><img src="../../../../img/fotos/professor/<%response.Write(cod_cons)%>.jpg" height="500"></div></td>
    </tr>
    <tr>
      <td colspan="2"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
          <%response.Write(nome_prof)%>
          </font></div></td>
    </tr>
  </table>
</div>
<%end if %>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> </td>
  </tr>
<%if opt="ok" then%>	
	  <tr> 
    <td width="1000" height="10"> 
      <%
	  	call mensagens(4,616,2,0) 
		%>
    </td>
  </tr>
<%end if%>

  <tr> 
    <td width="1000" height="10"> 
      <%
	  if autoriza="0" then
	  	call mensagens(4,9700,1,0) 	  
	  elseif autoriza="1" then
	  	call mensagens(4,9701,0,0) 	  
	  else
	  	call mensagens(4,614,0,0) 
	  end if%>
    </td>

  </tr>

  <tr> 
    <td valign="top"> 
            <%	  if autoriza="no" then			
		elseif ori="02" then		
%>
            <form action="bd.asp?opt=inc&nvg=<%=nvg%>" method="post" name="inclusao" id="inclusao" onSubmit="return checksubmit()">
              
        <table width="1000" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td colspan="6" class="tb_tit">Dados Pessoais do Professor </td>
          </tr>
          <tr> 
            <td width="151" height="20" class="tb_corpo"> <div align="right"><font class="form_dado_texto">C&oacute;digo: 
                </font></div></td>
            <td height="20" colspan="5" class="tb_corpo"> <font class="form_dado_texto"> 
              <input name="cod_cons" type="hidden" class="textInput" id="cod_cons" value="<%=cod_cons%>" size="4">
              <font class="form_corpo"> 
              <%RESPONSE.Write(cod_cons)%>
              </font> 
              <input name="tp" type="hidden" id="tp" value="L">
              <input name="acesso" type="hidden" id="acesso" value="2">
              </font></td>
          </tr>
          <tr class="tb_corpo"> 
            <td height="20"><div align="right"><font class="form_dado_texto">Nome: 
                </font></div></td>
            <td height="20" colspan="5"> <font class="form_dado_texto"> 
              <input name="nome" type="text" class="select_style" id="nome" size="75" maxlength="50">
              </font></td>
          </tr>
          <tr class="tb_corpo"> 
            <td height="20"> <div align="right"><font class="form_dado_texto">Apelido:</font></div></td>
            <td width="221" height="20"> <font class="form_dado_texto"> 
              <input name="apelido" type="text" class="textInput" id="apelido" size="20" maxlength="15">
              </font></td>
            <td width="137" height="20"> <div align="right"><font class="form_dado_texto">&nbsp;Data 
                de Nascimento:</font></div></td>
            <td height="20" colspan="3"> <font class="form_dado_texto"> 
              <input name="nasce" type="text" class="textInput" id="nasce" size="12" maxlength="10" onKeyup="formatar(this, '##/##/####')">
              </font></td>
          </tr>
          <tr class="tb_corpo"> 
            <td height="20"> <div align="right"><font class="form_dado_texto">Sexo:</font></div></td>
            <td height="20"> <font class="form_dado_texto"> 
              <select name="sexo" class="select_style" id="select">
                <%if sexo = "M" then%>
                <option value="0"></option>
                <option value="M" selected>Masculino</option>
                <option value="F">Feminino</option>
                <%elseif sexo = "F" then%>
                <option value="0"></option>
                <option value="M">Masculino</option>
                <option value="F" selected>Feminino</option>
                <%else%>
                <option value="0" selected></option>
                <option value="M">Masculino</option>
                <option value="F">Feminino</option>
                <%End IF%>
              </select>
              </font></td>
            <td height="20"> <div align="right"><font class="form_dado_texto">Pa&iacute;s:</font></div></td>
            <td height="20" colspan="3"> <font class="form_dado_texto"> 
              <select name="pais" class="select_style" id="select">
                <%				
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Paises order by NO_Pais"
		RS1.Open SQL1, CON0
		
while not RS1.EOF						
CO_Pais= RS1("CO_Pais")
NO_Pais= RS1("NO_Pais")

if CO_Pais = pais then
%>
                <option value="<%=CO_Pais%>" selected> 
                <% =NO_Pais%>
                </option>
                <%else%>
                <option value="<%=CO_Pais%>"> 
                <% =NO_Pais%>
                </option>
                <%
end if				
RS1.MOVENEXT
WEND
%>
              </select>
              </font></td>
          </tr>
          <tr class="tb_corpo"
> 
            <td height="20"> <div align="right"><font class="form_dado_texto">Nacionalidade: 
                </font></div></td>
            <td height="20"> <div align="left"> <font class="form_dado_texto"> 
                <select name="nacionalidade" class="select_style" id="select">
                  <%				
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Nacionalidades order by TX_Nacionalidade"
		RS0.Open SQL0, CON0
		
while not RS0.EOF						
CO_Nacionalidade= RS0("CO_Nacionalidade")
TX_Nacionalidade= RS0("TX_Nacionalidade")

if CO_Nacionalidade = nacionalidade then
%>
                  <option value="<%=CO_Nacionalidade%>" selected> 
                  <% =TX_Nacionalidade%>
                  </option>
                  <%else%>
                  <option value="<%=CO_Nacionalidade%>"> 
                  <% =TX_Nacionalidade%>
                  </option>
                  <%
end if						
RS0.MOVENEXT
WEND
%>
                </select>
                </font></div>
              <div align="left"> </div></td>
            <td height="20"> <div align="right"><font class="form_dado_texto">Natural 
                do estado:</font></div></td>
            <td height="20" colspan="3"> <font class="form_dado_texto"> 
              <select name="estadonat" class="select_style" onChange="recuperarCidNat(this.value)">
                <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
SG_UF= RS2("SG_UF")
NO_UF= RS2("NO_UF")
if isnull(uf_natural) then
uf_natural="RJ"
end if
if SG_UF = uf_natural then
%>
                <option value="<%=SG_UF%>" selected> 
                <% =NO_UF%>
                </option>
                <%else%>
                <option value="<%=SG_UF%>"> 
                <% =NO_UF%>
                </option>
                <%end if						
RS2.MOVENEXT
WEND
%>
              </select>
              </font> <div align="right"></div></td>
          </tr>
          <tr class="tb_corpo"
> 
            <td width="151" height="20"> <div align="right"><font class="form_dado_texto">Natural 
                da cidade:</font></div></td>
            <td height="20" colspan="5"> <div id="cid_nat"> 
                <select name="cidadenat" class="select_style" id="select">
                  <option value="0"></option>
                  <%
Set RS2m = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Municipios WHERE SG_UF='RJ' order by NO_Municipio"
		RS2m.Open SQL2m, CON0
		
while not RS2m.EOF						
SG_UF= RS2m("CO_Municipio")
NO_UF= RS2m("NO_Municipio")

if isnull(natural) then
natural=6001
end if

if SG_UF = natural then
%>
                  <option value="<%=SG_UF%>" selected> 
                  <% =NO_UF%>
                  </option>
                  <% else %>
                  <option value="<%=SG_UF%>"> 
                  <% =NO_UF%>
                  </option>
                  <%
end if	
RS2m.MOVENEXT
WEND
%>
                </select>
              </div></td>
          </tr>
          <tr class="tb_corpo"
> 
            <td height="20">&nbsp;</td>
            <td height="20">&nbsp;</td>
            <td height="20">&nbsp;</td>
            <td height="20" colspan="3">&nbsp;</td>
          </tr>
          <tr class="tb_corpo"> 
            <td  class="tb_tit" colspan="6">Dados Residenciais do Professor</td>
          </tr>
          <tr class="tb_corpo"> 
            <td height="20"> <div align="right"><font class="form_dado_texto">Logradouro:</font></div></td>
            <td height="20" colspan="3"><font class="form_dado_texto"> 
              <input name="rua" type="text" class="textInput" id="rua3" size="75" maxlength="50">
              </font></td>
            <td width="67" height="20"> <div align="right"><font class="form_dado_texto">N&uacute;mero:</font></div></td>
            <td width="297" height="20"> <font class="form_dado_texto"> 
              <input name="numero" type="text" class="textInput" id="numero" size="11" maxlength="6">
              &nbsp; </font></td>
          </tr>
          <tr class="tb_corpo"
> 
            <td height="20"> <div align="right"><font class="form_dado_texto">Complemento:</font></div></td>
            <td height="20"> <font class="form_dado_texto"> 
              <input name="complemento" type="text" class="textInput" id="complemento" size="30" maxlength="35">
              </font></td>
            <td height="20"> <div align="right"><font class="form_dado_texto">CEP:</font></div></td>
            <td width="127" height="20"> <font class="form_dado_texto"> 
              <input name="cep" type="text" class="textInput" id="cep" size="11" maxlength="9" onKeyup="formatar(this, '#####-###')">
              </font></td>
            <td height="20"> <div align="right"><font class="form_dado_texto">Estado:</font></div></td>
            <td height="20"> <font class="form_dado_texto"> 
              <select name="estadodom" class="select_style" onChange="recuperarCidRes(this.value)">
                <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
SG_UF= RS2("SG_UF")
NO_UF= RS2("NO_UF")

if SG_UF = uf then
%>
                <option value="<%=SG_UF%>" selected> 
                <% =NO_UF%>
                </option>
                <% else %>
                <option value="<%=SG_UF%>"> 
                <% =NO_UF%>
                </option>
                <%
end if	
RS2.MOVENEXT
WEND
%>
              </select>
              &nbsp; </font></td>
          </tr>
          <tr class="tb_corpo"> 
            <td height="20"> <div align="right"><font class="form_dado_texto">Cidade:</font></div></td>
            <td height="20"> <div id="cid_res"> 
                <select name="ciddom" class="select_style" id="ciddom" onChange="recuperarBairroRes(estadodom.value,this.value)">
                  <option value="0"></option>
                  <%
Set RS2m = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Municipios WHERE SG_UF='RJ' order by NO_Municipio"
		RS2m.Open SQL2m, CON0
		
while not RS2m.EOF						
SG_UF= RS2m("CO_Municipio")
NO_UF= RS2m("NO_Municipio")

if SG_UF = 6001 then
%>
                  <option value="<%=SG_UF%>" selected> 
                  <% =NO_UF%>
                  </option>
                  <% else %>
                  <option value="<%=SG_UF%>"> 
                  <% =NO_UF%>
                  </option>
                  <%
end if	
RS2m.MOVENEXT
WEND
%>
                </select>
              </div></td>
            <td height="20"> <div align="right"><font class="form_dado_texto">Bairro:</font></div></td>
            <td height="20" colspan="3"> <font class="form_dado_texto"> <div id="bairro_res">	
                <select name="bairro" class="select_style" id="bairro">
                  <%
Set RS2b = Server.CreateObject("ADODB.Recordset")
		SQL2b = "SELECT * FROM TB_Bairros WHERE CO_Municipio=6001 AND SG_UF='RJ' order by NO_Bairro"
		RS2b.Open SQL2b, CON0
		
while not RS2b.EOF						
SG_UF= RS2b("CO_Bairro")
NO_UF= RS2b("NO_Bairro")


%>
                  <option value="<%=SG_UF%>"> 
                  <% =NO_UF%>
                  </option>
                  <%

RS2b.MOVENEXT
WEND
%>
                </select>
              </div></td>
          </tr>
          <tr class="tb_corpo"> 
            <td height="20"
> <div align="right"><font class="form_dado_texto">Telefones de Contato:</font></div></td>
            <td height="20" colspan="5"> <font class="form_dado_texto"> 
              <input name="telefones" type="text" class="textInput" id="telefones2" size="75" maxlength="50">
              </font></td>
          </tr>
          <tr class="tb_corpo"
> 
            <td height="20"> <div align="right"><font class="form_dado_texto">Endere&ccedil;o 
                Eletr&ocirc;nico:</font></div></td>
            <td height="20" colspan="5"> <font class="form_dado_texto"> 
              <input name="email" type="text" class="textInput" id="email3" size="75" maxlength="50">
              </font></td>
          </tr>
          <tr class="tb_corpo"
> 
            <td height="20">&nbsp;</td>
            <td height="20"> <div align="right"><font class="form_dado_texto">O 
                Professor est&aacute; ativo nesta escola</font></div></td>
            <td height="20"> <font class="form_dado_texto"> 
              <select name="ativo" class="select_style" id="select">
                <option value="sim" selected>Sim</option>
                <option value="nao">N&atilde;o</option>
              </select>
              </font></td>
            <td height="20" colspan="3">&nbsp;</td>
          </tr>
          <tr class="tb_corpo"
>
            <td height="30" colspan="6"><hr></td>
          </tr>
          <tr class="tb_corpo"
> 
            <td height="30" colspan="6"><div align="center">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td width="33%">
<div align="center"></div></td>
                    <td width="34%"> 
                      <div align="center"></div></td>
                    <td width="33%">
<div align="center"> 
                        <input name="Submit22" type="submit" class="botao_prosseguir" id="Submit23" value="Confirmar">
                      </div></td>
                  </tr>
                </table>
              </div></td>
          </tr>
        </table>
            </form>
<%		elseif ori="01" then
%>
            <form name="alteracao" method="post" action="bd.asp?opt=alt&nvg=<%=nvg%>" id="alteracao" onSubmit="return checksubmit()">
              
        <table width="1000" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td colspan="7" class="tb_tit">Dados Pessoais do Professor </td>
          </tr>
          <tr> 
            <td width="146" height="20" class="tb_corpo"> <div align="right"><font class="form_dado_texto">C&oacute;digo: 
                </font></div></td>
            <td height="20" colspan="5" class="tb_corpo"><font class="form_corpo"> 
              <%RESPONSE.Write(cod_cons)%>
              </font> <input name="cod_cons" type="hidden" class="textInput" id="cod2" value="<%=cod_cons%>" size="4"> 
              <input name="tp" type="hidden" id="tp2" value="L"> <input name="acesso" type="hidden" id="acesso2" value="2"> 
              <input name="co_usr_prof" type="hidden" id="co_usr_prof" value="<% =co_usr_prof%>"> 
            </td>
            <td width="130" rowspan="7" valign="top" class="tb_corpo"> <table height="110" border="3" align="right" cellpadding="0" cellspacing="0" bordercolor="#EEEEEE">
                <tr> 
                  <td><div align="center"><a href="#" onClick="centraliza(500,536);MM_showHideLayers('fundo','','show','alinha','','show')"><img src="../../../../img/fotos/professor/<% =cod_cons %>.jpg" alt="" height="110" border="0"></a></div></td>
                </tr>
                <tr> 
                  <td height="15" bgcolor="#EEEEEE"> <div align="center"><a href="#" onClick="centraliza(500,536);MM_showHideLayers('fundo','','show','alinha','','show')"><img src="../../../../img/clique.gif" width="85" height="13" border="0"></a></div></td>
                </tr>
              </table></td>
          </tr>
          <tr class="tb_corpo"> 
            <td width="146" height="20"> <div align="right"><font class="form_dado_texto">Nome: 
                </font></div></td>
            <td height="20" colspan="5"> <input name="nome" type="text" class="textInput" id="nome" value="<%response.Write(nome_prof)%>" size="75" maxlength="50"> 
            </td>
          </tr>
          <tr class="tb_corpo"> 
            <td width="146" height="20"> <div align="right"><font class="form_dado_texto">Apelido:</font></div></td>
            <td width="208" height="20"> <input name="apelido" type="text" class="textInput" id="apelido2" value="<%response.Write(apelido)%>" size="20" maxlength="15"> 
            </td>
            <td width="123" height="20"> <div align="right">&nbsp;<font class="form_dado_texto">Data 
                de Nascimento:</font></div></td>
            <td height="20" colspan="3"> <input name="nasce" type="text" class="textInput" id="nasce2" onKeyup="formatar(this, '##/##/####')" value="<%response.Write(nasce)%>" size="12" maxlength="10"> 
              <font class="form_dado_texto"> &nbsp;-&nbsp; 
              <%
					call aniversario(ano_a,mes_a,dia_a) %>
              </font></td>
          </tr>
          <tr class="tb_corpo"> 
            <td width="146" height="20"> <div align="right"><font class="form_dado_texto">Sexo:</font></div></td>
            <td height="20"> <select name="sexo" class="select_style" id="select5">
                <%if sexo = "M" then%>
                <option value="0"></option>
                <option value="M" selected>Masculino</option>
                <option value="F">Feminino</option>
                <%elseif sexo = "F" then%>
                <option value="0"></option>
                <option value="M">Masculino</option>
                <option value="F" selected>Feminino</option>
                <%else%>
                <option value="0" selected></option>
                <option value="M">Masculino</option>
                <option value="F">Feminino</option>
                <%End IF%>
              </select> </td>
            <td height="20"> <div align="right"><font class="form_dado_texto">Pa&iacute;s:</font></div></td>
            <td height="20" colspan="3"> <select name="pais" class="select_style" id="select15">
                <%				
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Paises order by NO_Pais"
		RS1.Open SQL1, CON0

response.Write("tp = "&pais)		
while not RS1.EOF						
CO_Pais= RS1("CO_Pais")
NO_Pais= RS1("NO_Pais")

if CO_Pais = pais then
%>
                <option value="<%=CO_Pais%>" selected> 
                <% =NO_Pais%>
                </option>
                <%else%>
                <option value="<%=CO_Pais%>"> 
                <% =NO_Pais%>
                </option>
                <%
end if				
RS1.MOVENEXT
WEND
%>
              </select> </td>
          </tr>
          <tr class="tb_corpo"
> 
            <td width="146" height="20"> <div align="right"><font class="form_dado_texto">Nacionalidade:</font> 
              </div></td>
            <td height="20"> <div align="left"> 
                <select name="nacionalidade" class="select_style" id="select16">
                  <%				
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Nacionalidades order by TX_Nacionalidade"
		RS0.Open SQL0, CON0
		
while not RS0.EOF						
CO_Nacionalidade= RS0("CO_Nacionalidade")
TX_Nacionalidade= RS0("TX_Nacionalidade")

if CO_Nacionalidade = nacionalidade then
%>
                  <option value="<%=CO_Nacionalidade%>" selected> 
                  <% =TX_Nacionalidade%>
                  </option>
                  <%else%>
                  <option value="<%=CO_Nacionalidade%>"> 
                  <% =TX_Nacionalidade%>
                  </option>
                  <%
end if						
RS0.MOVENEXT
WEND
%>
                </select>
              </div>
              <div align="left"> </div></td>
            <td height="20"> <div align="right"><font class="form_dado_texto">Natural 
                do estado:</font></div></td>
            <td height="20" colspan="3"> <select name="estadonat" class="select_style" onChange="recuperarCidNat(this.value)">
                <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
SG_UF= RS2("SG_UF")
NO_UF= RS2("NO_UF")



if SG_UF = uf_natural then
%>
                <option value="<%=SG_UF%>" selected> 
                <% =NO_UF%>
                </option>
                <% else %>
                <option value="<%=SG_UF%>"> 
                <% =NO_UF%>
                </option>
                <%
end if						
RS2.MOVENEXT
WEND
%>
              </select> <div align="right"></div></td>
          </tr>
          <tr class="tb_corpo"
> 
            <td width="146" height="20"> <div align="right"><font class="form_dado_texto">Natural 
                da cidade:</font></div></td>
            <td height="20" colspan="4"> <div id="cid_nat"> 
                <select name="cidadenat" class="select_style" id="cidadenat">
                  <option value="0"></option>
                  <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Municipios where SG_UF = '"& uf_natural &"'  order by NO_Municipio"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
CO_Municipio= RS2("CO_Municipio")
NO_Municipio= RS2("NO_Municipio")

if CO_Municipio = natural then
%>
                  <option value="<%=CO_Municipio%>" selected> 
                  <% =NO_Municipio%>
                  </option>
                  <% else %>
                  <option value="<%=CO_Municipio%>"> 
                  <% =NO_Municipio%>
                  </option>
                  <%
end if						
RS2.MOVENEXT
WEND
%>
                </select>
              </div></td>
            <td width="181" height="20">&nbsp;</td>
          </tr>
          <tr class="tb_corpo"
> 
            <td height="20" colspan="6">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="7"  class="tb_tit">Dados Residenciais do Professor</td>
          </tr>
          <tr class="tb_corpo"> 
            <td height="20"> <div align="right"><font class="form_dado_texto">Logradouro:</font></div></td>
            <td height="20" colspan="3"> <input name="rua" type="text" class="textInput" id="rua3" value="<%response.Write(rua)%>" size="75" maxlength="50"></td>
            <td width="78" height="20"> <div align="right"><font class="form_dado_texto">N&uacute;mero:</font></div></td>
            <td height="20"> <input name="numero" type="text" class="textInput" id="numero" value="<%response.Write(numero)%>" size="11" maxlength="6"> 
              &nbsp; </td>
            <td width="130" valign="top" class="tb_corpo">&nbsp;</td>
          </tr>
          <tr class="tb_corpo"
> 
            <td height="20"> <div align="right"><font class="form_dado_texto">Complemento:</font></div></td>
            <td height="20"> <input name="complemento" type="text" class="textInput" id="complemento" value="<%response.Write(complemento)%>" size="30" maxlength="35"> 
            </td>
            <td height="20"> <div align="right"><font class="form_dado_texto">CEP:</font></div></td>
            <td width="134" height="20"> <input name="cep" type="text" class="textInput" id="cep" onKeyup="formatar(this, '#####-###')" value="<%response.Write(cep)%>" size="11" maxlength="9"> 
            </td>
            <td height="20"> <div align="right"><font class="form_dado_texto">Estado:</font></div></td>
            <td height="20" colspan="2"> <select name="estadodom" class="textInput" onChange="recuperarCidRes(this.value)">
                <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
SG_UF= RS2("SG_UF")
NO_UF= RS2("NO_UF")

if SG_UF = uf then
%>
                <option value="<%=SG_UF%>" selected> 
                <% =NO_UF%>
                </option>
                <% else %>
                <option value="<%=SG_UF%>"> 
                <% =NO_UF%>
                </option>
                <%
end if	
RS2.MOVENEXT
WEND
%>
              </select> &nbsp; </td>
          </tr>
          <tr class="tb_corpo"
> 
            <td height="20"> <div align="right"><font class="form_dado_texto">Cidade:</font></div></td>
            <td height="20"> <div id="cid_res"> 
                <select name="ciddom" class="select_style" onChange="recuperarBairroRes(estadodom.value,this.value)">
                  <%				
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Municipios where SG_UF = '"& uf &"' order by NO_Municipio"
		RS3.Open SQL3, CON0
		
while not RS3.EOF						
CO_Municipio= RS3("CO_Municipio")
NO_Municipio= RS3("NO_Municipio")

if CO_Municipio = municipio then
%>
                  <option value="<%=CO_Municipio%>" selected> 
                  <% =NO_Municipio%>
                  </option>
                  <% else %>
                  <option value="<%=CO_Municipio%>"> 
                  <% =NO_Municipio%>
                  </option>
                  <%
end if						
RS3.MOVENEXT
WEND
%>
                </select>
              </div></td>
            <td height="20"> <div align="right"><font class="form_dado_texto">Bairro:</font></div></td>
            <td height="20" colspan="4"> <div id="bairro_res"> 
                <select name="bairro" class="select_style">
                  <option value="0" selected> </option>
                  <%				
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Bairros where CO_Municipio = "&municipio&" order by NO_Bairro"
		RS4.Open SQL4, CON0
		
while not RS4.EOF						
CO_Bairro= RS4("CO_Bairro")

NO_Bairro= RS4("NO_Bairro")


if CO_Bairro = bairro then
%>
                  <option value="<%=CO_Bairro%>" selected> 
                  <% =NO_Bairro%>
                  </option>
                  <% else %>
                  <option value="<%=CO_Bairro%>"> 
                  <% =NO_Bairro%>
                  </option>
                  <%
end if	
RS4.MOVENEXT
WEND
%>
                </select>
              </div></td>
          </tr>
          <tr class="tb_corpo"> 
            <td height="20"
> <div align="right"><font class="form_dado_texto">Telefones de Contato:</font></div></td>
            <td height="20" colspan="6"> <input name="telefones" type="text" class="textInput" id="telefones2" value="<%response.Write(telefone)%>" size="75" maxlength="50"> 
            </td>
          </tr>
          <tr class="tb_corpo"
> 
            <td height="20"> <div align="right"><font class="form_dado_texto">Endere&ccedil;o 
                Eletr&ocirc;nico:</font></div></td>
            <td height="20" colspan="6"> <input name="email" type="text" class="textInput" id="email3" value="<%response.Write(email)%>" size="75" maxlength="50"> 
            </td>
          </tr>
          <tr class="tb_corpo"
> 
            <td height="20">&nbsp;</td>
            <td height="20"> <div align="right"><font class="form_dado_texto">O 
                Professor est&aacute; ativo nesta escola</font></div></td>
            <td height="20"> 
            <select name="ativo" class="select_style" id="select17">
                <%if ativo = "False" then%>
                <option value="sim">Sim</option>
                <option value="nao" selected>N&atilde;o</option>
                <% else %>
                <option value="sim" selected>Sim</option>
                <option value="nao">N&atilde;o</option>
                <% END IF%>
              </select> </td>
            <td height="20" colspan="4">&nbsp;</td>
          </tr>
          <tr class="tb_corpo"
>
            <td height="30" colspan="7"><hr></td>
          </tr>
          <tr class="tb_corpo"
> 
            <td height="30" colspan="7"><div align="center">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td width="33%"><div align="center"></div></td>
                    <td width="34%"><div align="center"></div></td>
                    <td width="33%"><div align="center">
					<%
					if (autoriza="1" or autoriza="0") then
					else
					%> 
                        <input name="Submit2" type="submit" class="botao_prosseguir" id="Submit24" value="Confirmar">
					<%end if%> 	
                      </div></td>
                  </tr>
                </table>
                
              </div></td>
          </tr>
        </table>			
            </form>
      <%end if%>
    </td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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