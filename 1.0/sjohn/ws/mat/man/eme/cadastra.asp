<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<% 
chave= request.QueryString("nvg")
session("nvg")=chave
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo")
ano_letivo_real = ano_letivo
sistema_local=session("sistema_local")
ori = request.QueryString("or")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
cod= request.QueryString("cod")	

z = request.QueryString("z")
erro = request.QueryString("e")
vindo = request.QueryString("vd")
obr = request.QueryString("o")
session("qtd_familiares")=0
session("pai_ok")="n"
session("mae_ok")="n"

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"

		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")		

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		




Call LimpaVetor2

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
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
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
  {    alert("Por favor, escolha o sexo do aluno!")
    document.inclusao.sexo.focus()
    return false
  }   
  if (document.inclusao.rua.value == "")
  {    alert("Por favor, digite a rua onde o aluno reside!")
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
                                               oHTTPRequest.open("post", "executa.asp?opt=v", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.dadesc.innerHTML =resultado_c
                                                           }
                                               }
                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }
						 function recuperarOrigem(oTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?opt=o", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_o  = oHTTPRequest.responseText;
resultado_o = resultado_o.replace(/\+/g," ")
resultado_o = unescape(resultado_o)
document.all.dadesc.innerHTML =resultado_o
                                                           }
                                               }
                                               oHTTPRequest.send("o_pub=" + oTipo);
                                   }

						 function recuperarCidNat(estadonat)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "cid_bairro.asp?opt=c&o=n", true);
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
                                               oHTTPRequest.open("post", "cid_bairro.asp?opt=c&o=r", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_cid_res  = oHTTPRequest.responseText;
resultado_cid_res = resultado_cid_res.replace(/\+/g," ")
resultado_cid_res = unescape(resultado_cid_res)
document.all.cid_res.innerHTML =resultado_cid_res
document.all.bairro_res.innerHTML ="<select class=borda></select>"
                                                           }
                                               }
                                               oHTTPRequest.send("c_pub=" + estadores);
                                   }
						 function recuperarCidCom(estadocom)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "cid_bairro.asp?opt=c&o=c", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_cid_com  = oHTTPRequest.responseText;
resultado_cid_com = resultado_cid_com.replace(/\+/g," ")
resultado_cid_com = unescape(resultado_cid_com)
document.all.cid_com.innerHTML =resultado_cid_com
document.all.bairro_com.innerHTML ="<select class=borda></select>"
                                                           }
                                               }
                                               oHTTPRequest.send("c_pub=" + estadocom);
                                   }
						 function recuperarBairroRes(cidres)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "cid_bairro.asp?opt=b&o=r", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_bairro_res  = oHTTPRequest.responseText;
resultado_bairro_res = resultado_bairro_res.replace(/\+/g," ")
resultado_bairro_res = unescape(resultado_bairro_res)
document.all.bairro_res.innerHTML =resultado_bairro_res
                                                           }
                                               }
                                               oHTTPRequest.send("b_pub=" + cidres);
                                   }
						 function recuperarBairroCom(cidcom)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "cid_bairro.asp?opt=b&o=c", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_bairro_com  = oHTTPRequest.responseText;
resultado_bairro_com = resultado_bairro_com.replace(/\+/g," ")
resultado_bairro_com = unescape(resultado_bairro_com)
document.all.bairro_com.innerHTML =resultado_bairro_com
                                                           }
                                               }
                                               oHTTPRequest.send("b_pub=" + cidcom);
                                   }

							 function recuperarPai(nome,tp)
                                   {
//pai = pai.replace(" ",/\+/g)
nome = escape(nome)
								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "pai_mae.asp", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_pai  = oHTTPRequest.responseText;
resultado_pai = resultado_pai.replace(/\+/g," ")
resultado_pai = unescape(resultado_pai)
document.all.pai_div.innerHTML =resultado_pai
                                                           }
                                               }
                                               oHTTPRequest.send("nome_pub=" + nome+ "&tp_familiares=" + tp);											   
                                   }
							 function recuperarMae(nome,tp)
                                   {
//pai = pai.replace(" ",/\+/g)
nome = escape(nome)
								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "pai_mae.asp", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_mae  = oHTTPRequest.responseText;
resultado_mae = resultado_mae.replace(/\+/g," ")
resultado_mae = unescape(resultado_mae)
document.all.mae_div.innerHTML =resultado_mae
                                                           }
                                               }
                                               oHTTPRequest.send("nome_pub=" + nome+ "&tp_familiares=" + tp);											   
                                   }
							 function recuperarFamiliares(nome,qtd,tp)
                                   {
//pai = pai.replace(" ",/\+/g)
nome = escape(nome)
								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "familiares.asp", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_familiares  = oHTTPRequest.responseText;
resultado_familiares = resultado_familiares.replace(/\+/g," ")
resultado_familiares = unescape(resultado_familiares)
document.all.familiares.innerHTML =resultado_familiares
                                                           }
                                               }
                                               oHTTPRequest.send("nome_pub=" + nome+ "&qtd_pub=" + qtd+ "&tp_familiares=" + tp);											   
                                   }								   								   							   								   								   								   

								   								   								   
								   								   								   
//-->
</script>
</head>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<table width="1002" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
            
    <td width="1000" height="10" colspan="5" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
    </td>
  </tr>
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,402,0,0) %>
    </td>
  </tr>			  
        <form action="cadastro.asp?opt=list&or=01" method="post" name="inclusao" id="inclusao" onSubmit="return checksubmit()">
          <tr>
      <td valign="top"> 
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
          <tr> 
            <td width="841" class="tb_tit"
>Dados Pessoais</td>
            <td width="151" class="tb_tit"
> </td>
            <td width="2" class="tb_tit"
></td>
          </tr>
          <tr> 
            <td height="10"> <font class="form_corpo"> 
              <input name="tp" type="hidden" id="tp2" value="P">
              </font> <table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="17%" height="10"><font class="form_dado_texto">Matr&iacute;cula</font></td>
                  <td width="2%"><div align="center">:</div></td>
                  <td width="26%" height="10">&nbsp;</td>
                  <td height="10"><font class="form_dado_texto">Nome:</font></td>
                  <td><div align="center">:</div></td>
                  <td height="10"><input name="textfield" type="text" class="borda" size="50"></td>
                </tr>
                <tr> 
                  <td height="10"> <div align="left"><font class="form_dado_texto"> 
                      Apelido</font></div></td>
                  <td><div align="center">:</div></td>
                  <td height="10"><input name="textfield243" type="text" class="borda" size="30"></td>
                  <td width="17%" height="10"> <div align="left"><font class="form_dado_texto"> 
                      Data de Nascimento</font></div></td>
                  <td width="2%"><div align="center">:</div></td>
                  <td height="10"><input name="nasce" type="text" class="borda" id="nasce2" size="12" maxlength="10" onKeyup="formatar(this, '##/##/####')"></td>
                </tr>
                <tr> 
                  <td height="10"> <div align="left"><font class="form_dado_texto"> 
                      Sexo</font></div></td>
                  <td><div align="center">:</div></td>
                  <td height="10"> <select name="sexo" class="borda" id="select14">
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
                    </select> &nbsp;</td>
                  <td height="10"> <div align="left"><font class="form_dado_texto"> 
                      Pa&iacute;s de Origem</font></div></td>
                  <td><div align="center">:</div></td>
                  <td height="10"><select name="pais" class="borda" id="select6">
                      <%				
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Paises order by NO_Pais"
		RS1.Open SQL1, CON0
		
while not RS1.EOF						
CO_Pais= RS1("CO_Pais")
NO_Pais= RS1("NO_Pais")
CO_Pais=CO_Pais*1
if CO_Pais = 10 then
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
                    </select></td>
                </tr>
                <tr> 
                  <td height="10"> <div align="left"><font class="form_dado_texto"> 
                      Nacionalidade</font></div></td>
                  <td><div align="center">:</div></td>
                  <td height="10"><select name="nacionalidade" class="borda" id="nacionalidade">
                      <%				
		Set RS_nacional= Server.CreateObject("ADODB.Recordset")
		SQL_nacional = "SELECT * FROM TB_Nacionalidades order by TX_Nacionalidade"
		RS_nacional.Open SQL_nacional, CON0
		
while not RS_nacional.EOF						
co_nacional= RS_nacional("CO_Nacionalidade")
no_nacional= RS_nacional("TX_Nacionalidade")
if co_nacional = 1 then
%>
                     <option value="<%=co_nacional%>" selected> 
                      <% =no_nacional%>
                      </option>
                      <%else%>
                      <option value="<%=co_nacional%>"> 
                      <% =no_nacional%>
                      </option>
                      <%end if						
RS_nacional.MOVENEXT
WEND
%>
                    </select></td>
                  <td height="10"> <div align="left"><font class="form_dado_texto"> 
                      Natural do Estado</font></div></td>
                  <td><div align="center">:</div></td>
                  <td height="10"><font class="form_corpo">
                    <select name="estadonat" class="borda" onChange="recuperarCidNat(this.value)">
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
if SG_UF = "RJ" then
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
                    </font></td>
                </tr>
                <tr> 
                  <td height="10"> <div align="left"><font class="form_dado_texto"> 
                      Natural da Cidade</font></div></td>
                  <td><div align="center">:</div></td>
                  <td height="10">
				  <div id="cid_nat">
				  <select name="cidnat" class="borda" id="select">
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
                    </select></div></td>
                  <td height="10"> <div align="left"><font class="form_dado_texto"> 
                      Religi&atilde;o</font></div></td>
                  <td><div align="center">:</div></td>
                  <td height="10">
                    <select name="religiao" class="borda" id="religiao">
                      <option value="0"></option>					
                      <%				
		Set RS_re = Server.CreateObject("ADODB.Recordset")
		SQL_re = "SELECT * FROM TB_Religiao order by TX_Descricao_Religiao"
		RS_re.Open SQL_re, CON0
		
while not RS_re.EOF						
co_relig= RS_re("CO_Religiao")
no_relig= RS_re("TX_Descricao_Religiao")
'if co_relig = 1 then
%>
                     <!--   <option value="<%'=co_relig%>" selected> 
                      <% '=no_relig%>
                      </option>-->
                      <%'else%>
                      <option value="<%=co_relig%>"> 
                      <% =no_relig%>
                      </option>
                      <%'end if						
RS_re.MOVENEXT
WEND
%>
                    </select>				  
				  
				  </td>
                </tr>
                <tr> 
                  <td height="10"> <div align="left"><font class="form_dado_texto"> 
                      Cor / Ra&ccedil;a</font></div></td>
                  <td><div align="center">:</div></td>
                  <td height="10">
				  <select name="cor_raca" class="borda" id="cor_raca">
				                        <option value="0"></option>
                      <%				
		Set RS_cor_raca = Server.CreateObject("ADODB.Recordset")
		SQL_cor_raca = "SELECT * FROM TB_Raca order by TX_Descricao_Raca"
		RS_cor_raca.Open SQL_cor_raca, CON0
		
while not RS_cor_raca.EOF						
co_cor_raca= RS_cor_raca("CO_Raca")
no_cor_raca= RS_cor_raca("TX_Descricao_Raca")
'if co_cor_raca = 1 then
%>
                     <!--   <option value="<%'=co_cor_raca%>" selected> 
                      <% '=no_cor_raca%>
                      </option>-->
                      <%'else%>
                      <option value="<%=co_cor_raca%>"> 
                      <% =no_cor_raca%>
                      </option>
                      <%'end if						
RS_cor_raca.MOVENEXT
WEND
%>
                    </select></td>
                  <td height="10"> <font class="form_dado_texto"> Ocupa&ccedil;&atilde;o</font></td>
                  <td><div align="center">:</div></td>
                  <td width="36%" height="10"><font class="form_corpo">
                    <select name="ocupacao" class="borda" id="ocupacao">
                      <%				
		Set RS_oc = Server.CreateObject("ADODB.Recordset")
		SQL_oc = "SELECT * FROM TB_Ocupacoes order by NO_Ocupacao"
		RS_oc.Open SQL_oc, CON0
		
while not RS_oc.EOF						
co_ocup= RS_oc("CO_Ocupacao")
no_ocup= RS_oc("NO_Ocupacao")
if co_ocup = 1 then
%>
                      <option value="<%=co_ocup%>" selected> 
                      <% =no_ocup%>
                      </option>
                      <%else%>
                      <option value="<%=co_ocup%>"> 
                      <% =no_ocup%>
                      </option>
                      <%end if						
RS_oc.MOVENEXT
WEND
%>
                    </select>
                    </font></td>
                </tr>
                <tr> 
                  <td height="10"><font class="form_dado_texto">Identidade</font></td>
                  <td><div align="center">:</div></td>
                  <td height="10"><input name="textfield22" type="text" class="borda" size="15"></td>
                  <td height="10"><font class="form_dado_texto">Tipo - Data de 
                    Emiss&atilde;o </font></td>
                  <td><div align="center">:</div></td>
                  <td height="10"><input name="textfield2352" type="text" class="borda" size="15">
                    - 
                    <input name="nasce2" type="text" class="borda" id="nasce" size="12" maxlength="10" onKeyup="formatar(this, '##/##/####')"></td>
                </tr>
                <tr> 
                  <td height="10"><div align="left"><font class="form_dado_texto"> 
                      CPF</font></div></td>
                  <td><div align="center">:</div></td>
                  <td height="10"><input name="textfield2" type="text" class="borda" size="15"></td>
                  <td height="10"><div align="left"><font class="form_dado_texto"> 
                      Empresa onde trabalha</font></div></td>
                  <td><div align="center">:</div></td>
                  <td height="10"><input name="textfield24522" type="text" class="borda" size="30"></td>
                </tr>
                <tr> 
                  <td height="10"><font class="form_dado_texto">E-mail</font></td>
                  <td><div align="center">:</div></td>
                  <td height="10"><input name="textfield244" type="text" class="borda" size="30"></td>
                  <td height="10">&nbsp;</td>
                  <td><div align="center"></div></td>
                  <td height="10">&nbsp;</td>
                </tr>
                <tr> 
                  <td height="10"><font class="form_dado_texto">Login Orkut</font></td>
                  <td><div align="center">:</div></td>
                  <td height="10"><input name="textfield245" type="text" class="borda" size="30"></td>
                  <td height="10"><font class="form_dado_texto">Login Messenger</font></td>
                  <td><div align="center">:</div></td>
                  <td height="10"><input name="textfield2452" type="text" class="borda" size="30"></td>
                </tr>
                <tr> 
                  <td height="10"> <div align="left"><font class="form_dado_texto">Telefones 
                      de Contato</font></div></td>
                  <td><div align="center">:</div></td>
                  <td height="10"><div align="left"><font class="form_corpo"></font> 
                      <font class="form_corpo"> 
                      <input name="textfield246" type="text" class="borda" size="30">
                      </font></div></td>
                  <td height="10"> <div align="left"><font class="form_dado_texto">Escrita</font></div></td>
                  <td><div align="center">:</div></td>
                  <td height="10"><select name="desteridade" id="desteridade" class="borda">
                      <option value="S">Destro</option>
                      <option value="N">Canhoto</option>
                    </select></td>
                </tr>
              </table></td>
            <td valign="top">&nbsp;</td>
            <td valign="top">&nbsp;</td>
          </tr>         
          <tr>
            <td colspan="3">
<div id="dadesc">		
			    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td class="tb_tit"
>Copiar dados do Aluno</td>
                  </tr>
                  <tr> 
                    <td> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr bgcolor="#FFFFFF" background="../../../../img/fundo_interno.gif"> 
                          <td width="5%"  height="10"> <div align="left"><font class="form_dado_texto"> 
                              Matr&iacute;cula:</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                              </strong></font></div></td>
                          <td width="10%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
                            </font><font size="2" face="Arial, Helvetica, sans-serif"> 
                            <input name="busca1" type="text" class="borda" id="busca1" size="12">
                            </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
                            </font></td>
                          <td width="37%" height="10"> <div align="right"><font class="form_dado_texto"> 
                              </font></div></td>
                          <td width="2%" height="10" ><font size="2" face="Arial, Helvetica, sans-serif">&nbsp; 
                            </font></td>
                          <td width="46%" height="10"><font size="2" face="Arial, Helvetica, sans-serif"> 
                            <input name="Button" type="button" class="borda_bot" id="Submit" value="Carregar" onClick="recuperarCurso(busca1.value)">
                            </font> </td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr class="tb_corpo"> 
                    <td class="tb_tit"
>Endere&ccedil;o Residencial</td>
                  </tr>
                  <tr class="tb_corpo"> 
                    <td height="10"> <table width="100%" border="0" cellspacing="0">
                        <tr> 
                          <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Logradouro</font></div></td>
                          <td width="13" class="tb_corpo"
> <div align="left">:</div></td>
                          <td width="217" height="10" class="tb_corpo"
><font class="form_corpo"> 
                            <input name="textfield2432" type="text" class="borda" size="30">
                            </font></td>
                          <td width="140" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
                          <td width="19" class="tb_corpo"
> <div align="center">:</div></td>
                          <td width="206" class="tb_corpo"
><font class="form_corpo"> <font class="form_corpo"> <font class="form_corpo"> 
                            <input name="nasce22" type="text" class="borda" id="nasce222" size="12" maxlength="10" onKeyUp="formatar(this, '##/##/####')">
                            </font></font></font></td>
                          <td width="90" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
                          <td width="11" class="tb_corpo"
> <div align="center">:</div></td>
                          <td width="139" height="10" class="tb_corpo"
> <div align="left"><font class="form_corpo"> </font> <font class="form_corpo"> 
                              <input name="nasce25" type="text" class="borda" id="nasce25" size="12" maxlength="10" onKeyup="formatar(this, '##/##/####')">
                              </font></div></td>
                        </tr>
                        <tr> 
                          <td width="145" height="21" class="tb_corpo"
><font class="form_dado_texto">Bairro</font></td>
                          <td width="13" class="tb_corpo"
> <div align="left">:</div></td>
                          <td width="217" height="21" class="tb_corpo"
><font class="form_corpo"> 
                            <div id="bairro_res"> 
                              <select name="bairrores" class="borda" id="bairrores">
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
                            </div>
                            </font></td>
                          <td width="140" height="21" class="tb_corpo"
><font class="form_dado_texto">Cidade</font></td>
                          <td width="19" class="tb_corpo"
> <div align="center">:</div></td>
                          <td width="206" class="tb_corpo"
><font class="form_corpo"> 
                            <div id="cid_res"> 
                              <select name="cidres" class="borda" id="select10" onChange="recuperarBairroRes(this.value)">
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
                            </div>
                            </font></td>
                          <td width="90" class="tb_corpo"
><font class="form_dado_texto">Estado</font></td>
                          <td width="11" class="tb_corpo"
> <div align="center">:</div></td>
                          <td width="139" height="21" class="tb_corpo"
><font class="form_corpo"> 
                            <select name="estadores" class="borda" id="estadores" onChange="recuperarCidRes(this.value)">
                              <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
SG_UF= RS2("SG_UF")
NO_UF= RS2("NO_UF")

if SG_UF = "RJ" then
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
                            </font></td>
                        </tr>
                        <tr> 
                          <td width="145" height="10" class="tb_corpo"
><font class="form_dado_texto">CEP</font></td>
                          <td width="13" class="tb_corpo"
> <div align="left">:</div></td>
                          <td width="217" height="10" class="tb_corpo"
><font class="form_dado_texto"> 
                            <input name="cep" type="text" class="borda" id="cep" size="11" maxlength="9" onKeyup="formatar(this, '#####-###')">
                            </font></td>
                          <td width="140" height="10" class="tb_corpo"
>&nbsp;</td>
                          <td width="19" class="tb_corpo"
>&nbsp;</td>
                          <td width="206" class="tb_corpo"
>&nbsp;</td>
                          <td width="90" class="tb_corpo"
>&nbsp;</td>
                          <td width="11" class="tb_corpo"
> <div align="center"></div></td>
                          <td width="139" height="10" class="tb_corpo"
>&nbsp;</td>
                        </tr>
                        <tr> 
                          <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">Telefones deste endere&ccedil;o</font></div></td>
                          <td width="13" class="tb_corpo"
> <div align="left">:</div></td>
                          <td height="10" colspan="2" class="tb_corpo"
><font class="form_corpo"> 
                            <input name="telefones" type="text" class="borda" id="telefones" size="50" maxlength="50">
                            </font> <div align="left"></div></td>
                          <td width="19" class="tb_corpo"
> <div align="center">:</div></td>
                          <td width="206" class="tb_corpo"
>&nbsp;</td>
                          <td width="90" class="tb_corpo"
>&nbsp;</td>
                          <td width="11" class="tb_corpo"
> <div align="center"></div></td>
                          <td width="139" height="10" class="tb_corpo"
>&nbsp;</td>
                        </tr>
                        <tr> 
                          <td height="10" colspan="9" class="tb_tit"
><div align="left">Endere&ccedil;o Comercial </div></td>
                        </tr>
                        <tr> 
                          <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Logradouro</font></div></td>
                          <td width="13" class="tb_corpo"
> <div align="left">:</div></td>
                          <td width="217" height="10" class="tb_corpo"
><font class="form_corpo"> 
                            <input name="textfield2433" type="text" class="borda" size="30">
                            </font></td>
                          <td width="140" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
                          <td width="19" class="tb_corpo"
> <div align="center">:</div></td>
                          <td width="206" class="tb_corpo"
><font class="form_corpo"> 
                            <input name="nasce23" type="text" class="borda" id="nasce23" size="12" maxlength="10" onKeyup="formatar(this, '##/##/####')">
                            </font></td>
                          <td width="90" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
                          <td class="tb_corpo"
><div align="center">:</div></td>
                          <td width="139" height="10" class="tb_corpo"
> <div align="left"><font class="form_corpo"> </font> <font class="form_corpo"> 
                              <input name="nasce24" type="text" class="borda" id="nasce24" size="12" maxlength="10" onKeyup="formatar(this, '##/##/####')">
                              </font></div></td>
                        </tr>
                        <tr class="tb_corpo"
> 
                          <td width="145" height="26"><font class="form_dado_texto">Bairro</font></td>
                          <td width="13"> <div align="left">:</div></td>
                          <td width="217" height="26"><font class="form_corpo"> 
                            <div id="bairro_com"> 
                              <select name="bairrocom" class="borda" id="bairro">
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
                            </div>
                            </font></td>
                          <td width="140" height="26"><font class="form_dado_texto">Cidade</font></td>
                          <td width="19"> <div align="center">:</div></td>
                          <td width="206"> <div id="cid_com"> 
                              <select name="cidcom" class="borda" id="select10" onChange="recuperarBairroCom(this.value)">
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
                          <td width="90"><font class="form_dado_texto">Estado</font></td>
                          <td><div align="center">:</div></td>
                          <td width="139" height="26"><font class="form_corpo"> 
                            <select name="estadocom" class="borda" id="estadocom" onChange="recuperarCidCom(this.value)">
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
if SG_UF = "RJ" then
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
                            </font></td>
                        </tr>
                        <tr class="tb_corpo"
> 
                          <td width="145" height="26"><font class="form_dado_texto">CEP</font></td>
                          <td width="13"> <div align="left">:</div></td>
                          <td width="217" height="26"><font class="form_dado_texto"> 
                            <input name="cepcom" type="text" class="borda" id="cepcom" size="11" maxlength="9" onKeyup="formatar(this, '#####-###')">
                            </font></td>
                          <td width="140" height="26">&nbsp;</td>
                          <td width="19">&nbsp;</td>
                          <td width="206">&nbsp;</td>
                          <td width="90">&nbsp;</td>
                          <td><div align="center"></div></td>
                          <td width="139" height="26">&nbsp;</td>
                        </tr>
                        <tr class="tb_corpo"
> 
                          <td width="145" height="28"> <div align="left"><font class="form_dado_texto">Telefones 
                              deste endere&ccedil;o:</font></div></td>
                          <td width="13"> <div align="left">:</div></td>
                          <td height="28" colspan="2"><font class="form_corpo"> 
                            <input name="telefones" type="text" class="borda" id="telefones" size="50" maxlength="50">
                            </font> <div align="left"></div></td>
                          <td width="19"> <div align="center"></div></td>
                          <td width="206">&nbsp;</td>
                          <td width="90">&nbsp;</td>
                          <td><div align="center"></div></td>
                          <td width="139" height="28">&nbsp;</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr class="tb_corpo"> 
                    <td class="tb_tit"
>Filia&ccedil;&atilde;o</td>
                  </tr>
                  <tr class="tb_corpo"> 
                    <td><table width="100%" border="0" cellspacing="0">
                        <tr> 
                          <td width="145" height="26"> <div align="left"><font class="form_dado_texto"> 
                              Pai</font></div></td>
                          <td width="13"> <div align="left">:</div></td>
                          <td width="217" height="26"><font class="form_corpo"> 
                            <input name="pai" type="text" class="borda" size="30" onBlur="recuperarPai(this.value,'p');recuperarFamiliares('nulo','1','o')">
                            </font></td>
                          <td width="140" height="26"> <div align="left"><font class="form_dado_texto"> 
                              Falecido</font></div></td>
                          <td width="19"> <div align="center"><font class="form_dado_texto">?</font></div></td>
                          <td width="206" height="26"><font class="form_corpo"> 
                            <select name="pai_falecido" class="borda">
                              <option value="n">N&atilde;o</option>
                              <option value="s">Sim</option>
                            </select>
                            </font></td>
                          <td width="90" height="26"> <div align="left"><font class="form_dado_texto"> 
                              Situa&ccedil;&atilde;o dos Pais</font></div></td>
                          <td width="11"> <div align="center">:</div></td>
                          <td width="139" height="26"><select name="sit_pais" class="borda" id="sit_pais">
                              <option value=0></option>
                              <%				
		Set RS_ec = Server.CreateObject("ADODB.Recordset")
		SQL_ec = "SELECT * FROM TB_Estado_Civil order by CO_Estado_Civil"
		RS_ec.Open SQL_ec, CON0
		
while not RS_ec.EOF						
co_ec= RS_ec("CO_Estado_Civil")
no_ec= RS_ec("TX_Estado_Civil")

if co_ec=5 then
%>
                              <option value="<%=co_ec%>" selected> 
                              <% =no_ec%>
                              </option>
                              <%
else							  
%>
                              <option value="<%=co_ec%>"> 
                              <% =no_ec%>
                              </option>
                              <%
end if							  						
RS_ec.MOVENEXT
WEND
%>
                            </select></td>
                        </tr>
                        <tr> 
                          <td width="145" height="10"> <div align="left"><font class="form_dado_texto"> 
                              M&atilde;e</font></div></td>
                          <td width="13"> <div align="left">: </div></td>
                          <td width="217" height="10"><font class="form_corpo"> 
                            <input name="mae" type="text" class="borda" size="30" onBlur="recuperarMae(this.value,'m');recuperarFamiliares('nulo','1','o')">
                            </font></td>
                          <td width="140" height="10"> <div align="left"><font class="form_dado_texto"> 
                              Falecida</font></div></td>
                          <td width="19"> <div align="center"><font class="form_dado_texto">?</font></div></td>
                          <td width="206" height="10"><font class="form_corpo"> 
                            <select name="mae_falecido" class="borda">
                              <option value="n">N&atilde;o</option>
                              <option value="s">Sim</option>
                            </select>
                            </font></td>
                          <td width="90" height="10"> <div align="left"><font class="form_dado_texto"> 
                              </font></div></td>
                          <td width="11"> <div align="center"></div></td>
                          <td width="139" height="10"><font class="form_dado_texto">&nbsp; 
                            </font></td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr class="tb_corpo"> 
                    <td width="145" class="tb_tit"
>Familiares<input name="qtd_familiares" type="hidden" value="0"></td>
                  </tr>
                  <tr class="tb_corpo"> 
                    <td> <div id="pai_div"> 
                      </div>
					   <div id="mae_div"> 
                      </div>
					   <div id="familiares"> 
                      </div></td>
                  </tr>
                </table>
</div>
			  
			  </td>
          </tr>
        </table></td></tr>
</form>
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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
response.redirect("../../../../inc/erro.asp")
end if
%>