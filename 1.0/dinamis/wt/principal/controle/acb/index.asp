<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%
opt = request.QueryString("opt")

ano_letivo = session("ano_letivo")

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


nvg = request.QueryString("nvg")
chave=nvg
session("chave")=chave
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano_info=nivel&"-"&chave&"-"&ano_letivo



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
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function checksubmit()
{
// if (document.formulario.tipo_doc.value == "0")
//  {    alert("Por favor selecione um tipo de Documento!")
//   document.formulario.tipo_doc.focus()
//    return false
// }
// 
//  if (document.formulario.status.value == "nulo")
//  {    alert("Por favor selecione um Status!")
//   document.formulario.status.focus()
//    return false
// }
 
 
  return true
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
function submitforminterno()  
{
   var f=document.forms[3]; 
      f.submit(); 
	  
}
//-->
</script>
<script>
<!--

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
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=c", true);
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
document.all.divEtapa.innerHTML ="<select class=borda></select>"
document.all.divTurma.innerHTML = "<select class=borda></select>"
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
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=e", true);
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
document.all.divTurma.innerHTML = "<select class=borda></select>"
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
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t", true);
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
						 function recuperarLink(dTipo)
                                   {
// Cria��o do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicita��o HTTP. O primeiro par�metro informa o m�todo post/get
// O segundo par�metro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicita��o s�ncrona, o par�metro deve ser false
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=l", true);
// Para solicita��es utilizando o m�todo post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A fun��o abaixo � executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto j� completou a solicita��o
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto � gerado no arquivo executa.asp e colocado no div
                                                                       var resultado_l= oHTTPRequest.responseText;
resultado_l = resultado_l.replace(/\+/g," ")
resultado_l = unescape(resultado_l)
document.all.divLink.innerHTML = resultado_l																	   
                                                           }
                                               }
// Abaixo � enviada a solicita��o. Note que a configura��o
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("d_pub=" + dTipo);
                                   }	
	var CID = null
	
	function clickHandler(){
		clearTimeout(CID)
		CID = setTimeout(getSuggestions, 333);
	}
	
	function getSuggestions(){		
	var input = escape(document.getElementById("busca2").value);
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/sugestoes.asp", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_busca  = oHTTPRequest.responseText;
resultado_busca = resultado_busca.replace(/\+/g," ")
resultado_busca = unescape(resultado_busca)
document.all.suggs.innerHTML =resultado_busca

                                                           }
                                               }

                                               oHTTPRequest.send("input=" + input);
                                   }
								   

	function limpa_sugestoes(){		

	document.all.suggs.innerHTML =""


                                   }
								   
//-->
</script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<link href="../../../../suggestions.css" rel="stylesheet" type="text/css" />
</head>

<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../../img/menu_r1_c2_f3.gif','../../../../img/menu_r1_c2_f2.gif','../../../../img/menu_r1_c2_f4.gif','../../../../img/menu_r1_c4_f3.gif','../../../../img/menu_r1_c4_f2.gif','../../../../img/menu_r1_c4_f4.gif','../../../../img/menu_r1_c6_f3.gif','../../../../img/menu_r1_c6_f2.gif','../../../../img/menu_r1_c6_f4.gif','../../../../img/menu_r1_c8_f3.gif','../../../../img/menu_r1_c8_f2.gif','../../../../img/menu_r1_c8_f4.gif','../../../../img/menu_direita_r2_c1_f3.gif','../../../../img/menu_direita_r2_c1_f2.gif','../../../../img/menu_direita_r2_c1_f4.gif','../../../../img/menu_direita_r4_c1_f3.gif','../../../../img/menu_direita_r4_c1_f2.gif','../../../../img/menu_direita_r4_c1_f4.gif','../../../../img/menu_direita_r6_c1_f3.gif','../../../../img/menu_direita_r6_c1_f2.gif','../../../../img/menu_direita_r6_c1_f4.gif')">
<FORM name="formulario" METHOD="POST" ACTION="contratos.asp?pagina=1&v=n">
<% call cabecalho_novo (nivel)
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
                <tr> 
                    
    <td height="10"> 
      <%	call mensagens(4,9706,0,0) 
	  
	  
%>
</td></tr>
<tr>

            <td valign="top"> 
<%
ano = DatePart("yyyy", now) 
mes = DatePart("m", now) 
dia = DatePart("d", now) 

dia=dia*1
mes=mes*1
%>				
                
        <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td width="653" height="15" class="tb_tit">Informe os crit&eacute;rios 
              para pesquisa 
              <input name="co_grupo" type="hidden" id="co_grupo" value="<% = co_grupo %>"></td>
          </tr>
          <tr> 
            <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td colspan="4"><table width="1000" border="0" cellpadding="0" cellspacing="0">
                          <tr>
                            <td width="147"  height="10"></td>
                            <td width="68" height="10"></td>
                            <td width="141" height="10"></td>
                            <td width="304" height="10" ></td>
                            <td width="340" rowspan="3" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                              <tr class="tb_subtit">
                                <td align="center">Somente Ativos</td>
                                <td align="center">Somente Cancelados</td>
                                <td align="center">Sem Parcelas</td>
                                <td align="center">Somente Bolsistas</td>
                              </tr>
                              <tr>
                                <td align="center"><input name="ativos" type="checkbox" id="ativos" value="s" checked></td>
                                <td align="center"><input type="checkbox" name="cancelados" id="cancelados" value="s"></td>
                                <td align="center"><input name="sem_parcelas" type="checkbox" id="sem_parcelas" value="s"></td>
                                <td align="center"><input name="so_bolsistas" type="checkbox" id="so_bolsistas" value="s"></td>
                              </tr>
                            </table></td>
                          </tr>
                          <tr>
                            <td width="147"  height="15"><div align="right"><font class="form_dado_texto"> Matr&iacute;cula:</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> </strong></font></div></td>
                            <td width="68" height="35" rowspan="2" valign="top">
                              <input name="busca1" type="text" class="textInput" id="busca1" size="12">
                           </td>
                            <td width="141" height="15"><div align="right"><font class="form_dado_texto"> Nome: </font></div></td>
                            <td width="304" height="35" rowspan="2" valign="top" >
                              <input name="busca2" type="text" class="textInput" id="busca2" size="55" maxlength="50" autocomplete="off" onKeyUp="clickHandler()"  /><div id="suggs"></div>
                            </td>
<div id="usersList"></div>                            
                            </tr>
                          <tr>
                            <td  height="20">&nbsp;</td>
                            <td height="20">&nbsp;</td>
                            </tr>
                        </table></td>
                      </tr>
                      <tr> 
                        <td width="12%" class="tb_subtit"><div align="center"> N&uacute;mero de Contrato</div></td>
                        <td width="46%" class="tb_subtit"><div align="center">Per&iacute;odo</div></td>
                        <td width="20%" align="center" class="tb_subtit">Tipo de Bolsa</td>
                        <td width="22%" class="tb_subtit"><div align="center">Desconto</div></td>
                      </tr>
                      <tr> 
                        <td width="12%" align="center"><font size="2" face="Arial, Helvetica, sans-serif">
                          <input name="contrato" type="text" class="textInput" id="contrato" size="12">
                        </font></td>
                        <td width="46%"><div align="center" class="form_dado_texto">
                         
  <select name="dia_de" id="dia_de" class="borda">
    <option value="1" selected>01</option>
    <option value="2">02</option>
    <option value="3">03</option>
    <option value="4">04</option>
    <option value="5">05</option>
    <option value="6">06</option>
    <option value="7">07</option>
    <option value="8">08</option>
    <option value="9">09</option>
    <option value="10">10</option>
    <option value="11">11</option>
    <option value="12">12</option>
    <option value="13">13</option>
    <option value="14">14</option>
    <option value="15">15</option>
    <option value="16">16</option>
    <option value="17">17</option>
    <option value="18">18</option>
    <option value="19">19</option>
    <option value="20">20</option>
    <option value="21">21</option>
    <option value="22">22</option>
    <option value="23">23</option>
    <option value="24">24</option>
    <option value="25">25</option>
    <option value="26">26</option>
    <option value="27">27</option>
    <option value="28">28</option>
    <option value="29">29</option>
    <option value="30">30</option>
    <option value="31">31</option>
  </select>
                          /
  <select name="mes_de" id="mes_de" class="borda">
    <option value="1" selected>janeiro</option>
    <option value="2">fevereiro</option>
    <option value="3">mar&ccedil;o</option>
    <option value="4">abril</option>
    <option value="5">maio</option>
    <option value="6">junho</option>
    <option value="7">julho</option>
    <option value="8">agosto</option>
    <option value="9">setembro</option>
    <option value="10">outubro</option>
    <option value="11">novembro</option>
    <option value="12" >dezembro</option>
  </select>
                          /
 <select name="ano_de" id="ano_de" class="borda">
 <%

		
		Set RSano = Server.CreateObject("ADODB.Recordset")
		SQLano = "SELECT * FROM TB_Ano_Letivo order by NU_Ano_Letivo"
		RSano.Open SQLano, CON	
		
	while not RSano.eof
		
		ald = RSano("NU_Ano_Letivo")
		ald=ald*1
		ano_letivo=ano_letivo*1		
 		if ald=ano_letivo then
			selected="selected"
		else
			selected=""
		end if				
 %>
     <option value="<%Response.Write(ald)%>" <%response.Write(selected)%>><%Response.Write(ald)%></option>
  <%RSano.movenext
  	wend%>
  </select>     
                          at&eacute; <select name="dia_ate" id="dia_ate" class="borda">
                            <% 
							 For i =1 to 31
							 dia=dia*1
							 if dia=i then 
								if dia<10 then
								dia="0"&dia
								end if
							 %>
                            <option value="<%response.Write(i)%>" selected>
                              <%response.Write(dia)%>
                              </option>
                            <% else
							  	i_cod=i
								if i<10 then
								
								i="0"&i
								end if
							%>
                            <option value="<%response.Write(i_cod)%>">
                              <%response.Write(i)%>
                              </option>
                            <% end if 
							next
							%>
                            </select>
                          /
  <select name="mes_ate" id="mes_ate" class="borda">
    <%mes=mes*1
								if mes="1" or mes=1 then%>
    <option value="1" selected>janeiro</option>
    <% else%>
    <option value="1">janeiro</option>
    <%end if
								if mes="2" or mes=2 then%>
    <option value="2" selected>fevereiro</option>
    <% else%>
    <option value="2">fevereiro</option>
    <%end if
								if mes="3" or mes=3 then%>
    <option value="3" selected>mar&ccedil;o</option>
    <% else%>
    <option value="3">mar&ccedil;o</option>
    <%end if
								if mes="4" or mes=4 then%>
    <option value="4" selected>abril</option>
    <% else%>
    <option value="4">abril</option>
    <%end if
								if mes="5" or mes=5 then%>
    <option value="5" selected>maio</option>
    <% else%>
    <option value="5">maio</option>
    <%end if
								if mes="6" or mes=6 then%>
    <option value="6" selected>junho</option>
    <% else%>
    <option value="6">junho</option>
    <%end if
								if mes="7" or mes=7 then%>
    <option value="7" selected>julho</option>
    <% else%>
    <option value="7">julho</option>
    <%end if%>
    <%if mes="8" or mes=8 then%>
    <option value="8" selected>agosto</option>
    <% else%>
    <option value="8">agosto</option>
    <%end if
								if mes="9" or mes=9 then%>
    <option value="9" selected>setembro</option>
    <% else%>
    <option value="9">setembro</option>
    <%end if
								if mes="10" or mes=10 then%>
    <option value="10" selected>outubro</option>
    <% else%>
    <option value="10">outubro</option>
    <%end if
								if mes="11" or mes=11 then%>
    <option value="11" selected>novembro</option>
    <% else%>
    <option value="11">novembro</option>
    <%end if
								if mes="12" or mes=12 then%>
    <option value="12" selected>dezembro</option>
    <% else%>
    <option value="12">dezembro</option>
    <%end if%>
  </select>
                          /
 <select name="ano_ate" id="ano_ate" class="borda">
 <% 		
 		Set RSano = Server.CreateObject("ADODB.Recordset")
		SQLano = "SELECT * FROM TB_Ano_Letivo order by NU_Ano_Letivo"
		RSano.Open SQLano, CON
		
	while not RSano.eof
		
		ala = RSano("NU_Ano_Letivo")	 
		ala=ala*1
		ano_letivo=ano_letivo*1
		if ala=ano_letivo then
			selected="selected"
		else
			selected=""
		end if	
 %>
     <option value="<%Response.Write(ala)%>" <%response.Write(selected)%>><%Response.Write(ala)%></option>
  <%RSano.movenext
  	wend%>
  </select>     
                        </div></td>
                        <td width="20%" align="center"><label>
                          <select name="bolsa" id="bolsa"  class="borda">
                            <option value="nulo"></option>                          
                          <%
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Tipo_Bolsa Order By NO_Bolsa"
	RS.Open SQL, CON0	
	
	while not RS.EOF
		co_bolsa=RS("CO_Bolsa")		
		nome_bolsa=RS("NO_Bolsa")
%>
                            <option value="<%response.Write(co_bolsa)%>">
                              <%response.Write(nome_bolsa)%>
                              </option>

<%		
	RS.MOVENEXT
	WEND	
						  %>
                          </select>
                        </label></td>
                        <td width="22%"><div align="center" class="form_dado_texto">
                          <select name="desconto_de" id="desconto_de" class="borda">
<%For dd=0 to 101
	if dd=0 then
		selected="selected"
	else
		selected=""
	end if		
%>
                            <option value="<%response.Write(dd)%>" <%response.Write(selected)%>>
                              <%response.Write(dd)%>%
                              </option>

<%
dd=dd+4
next%>
                          </select>
                          at&eacute;
                          <select name="desconto_ate" id="desconto_ate" class="borda">
<%For da=0 to 101
	if da>=100 then
		selected="selected"
	else
		selected=""
	end if		
%>
                            <option value="<%response.Write(da)%>" <%response.Write(selected)%>>
                              <%response.Write(da)%>%
                              </option>

<%
da=da+4
next%>
                          </select>
                        </div></td>
                      </tr>
                  </table></td>
                </tr>
                <tr>
                  <td valign="top">&nbsp;</td>
                  <td >&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td width="220" class="tb_subtit"> <div align="center">Unidade</div></td>
                  <td width="280" class="tb_subtit"> <div align="center">Curso</div></td>
                  <td width="253" class="tb_subtit"> <div align="center">Etapa</div></td>
                  <td width="247" class="tb_subtit"> <div align="center">Turma</div></td>
                </tr>
                <tr> 
                  <td width="220"> <div align="center"> 
                      <select name="unidade" class="borda" onChange="recuperarCurso(this.value)">
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
                  <td width="280"> <div align="center"> 
                      <div id="divCurso"> 
                        <select class="borda">
                        </select>
                      </div>
                    </div></td>
                  <td width="253"> <div align="center"> 
                      <div id="divEtapa"> 
                        <select class="borda">
                        </select>
                      </div>
                    </div></td>
                  <td width="247"> <div align="center"> 
                      <div id="divTurma"> 
                        <select class="borda">
                        </select>
                      </div>
                    </div></td>                                                                
                </tr>
                <tr> 
                  <td colspan="4"><hr width="1000"></td>
                </tr>
                <tr> 
                  <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="33%"> <div align="center"> </div></td>
                        <td width="34%"> <div align="center"></div> <div align="center"></div></td>
                        <td width="33%"> <div align="center"> 
                          <input name="SUBMIT" type=SUBMIT class="botao_prosseguir" value="Prosseguir">
                        </div></td>
                      </tr>
                    </table></td>
                </tr>
              </table>
 </td>
          </tr>
        </table>
</td>
          </tr>
		  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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