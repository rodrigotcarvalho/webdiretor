<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%
Session.LCID = 1046
nivel=4
ori = request.QueryString("or")
opt= request.QueryString("opt")
permissao = session("permissao") 
ano_letivo = Session("ano_letivo")
sistema_local=session("sistema_local")
trava=session("trava")
nvg=session("nvg")
session("nvg")=nvg


if Request.QueryString("pagina")="" then
      intpagina = 1
tp_ocor=request.form("tp_ocor")
qtd_ocor=request.form("qtd_ocor")

curso = request.Form("curso")
unidade = request.Form("unidade")
co_etapa = request.Form("etapa")
turma= request.Form("turma")

dia_de= request.form("dia_de")
mes_de= request.form("mes_de")

dia_ate= request.form("dia_ate")
mes_ate= request.form("mes_ate")

Session("tp_ocor")=tp_ocor
Session("qtd_ocor")=qtd_ocor
Session("curso")=curso
Session("unidade")=unidade
Session("co_etapa")=co_etapa
Session("turma")=turma
Session("dia_de")=dia_de
Session("mes_de")=mes_de
Session("dia_ate")=dia_ate
Session("mes_ate")=mes_ate
else
tp_ocor=Session("tp_ocor")
qtd_ocor=Session("qtd_ocor")
curso=Session("curso")
unidade=Session("unidade")
co_etapa=Session("co_etapa")
turma=Session("turma")
dia_de=Session("dia_de")
mes_de=Session("mes_de")
dia_ate=Session("dia_ate")
mes_ate=Session("mes_ate")
Session("tp_ocor")=tp_ocor
Session("qtd_ocor")=qtd_ocor
Session("curso")=curso
Session("unidade")=unidade
Session("co_etapa")=co_etapa
Session("turma")=turma
Session("dia_de")=dia_de
Session("mes_de")=mes_de
Session("dia_ate")=dia_ate
Session("mes_ate")=mes_ate
end if

obr=tp_ocor&"_"&qtd_ocor&"_"&unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&dia_de&"_"&mes_de&"_"&dia_ate&"_"&mes_ate

data_de=mes_de&"/"&dia_de&"/"&ano_letivo


dia_de=dia_de*1
mes_de=mes_de*1

if dia_de<10 then
dia_de="0"&dia_de
end if
if mes_de<10 then
mes_de="0"&mes_de
end if


data_inicio=dia_de&"/"&mes_de&"/"&ano_letivo

data_ate=mes_ate&"/"&dia_ate&"/"&ano_letivo

dia_ate=dia_ate*1
mes_ate=mes_ate*1

if dia_ate<10 then
dia_ate="0"&dia_ate
end if
if mes_ate<10 then
mes_ate="0"&mes_ate
end if

data_fim=dia_ate&"/"&mes_ate&"/"&ano_letivo



		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON1 = Server.CreateObject("ADODB.Connection")
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2
		
		Set CON3 = Server.CreateObject("ADODB.Connection") 
		ABRIR3 = "DBQ="& CAMINHO_o & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3		
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CONp = Server.CreateObject("ADODB.Connection") 
		ABRIRp = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONp.Open ABRIRp		
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0	

 call navegacao (CON,nvg,nivel)
navega=Session("caminho")	

if unidade="999990" or unidade="" or isnull(unidade) then
	SQL_ALUNOS="NULO"
else	
	SQL_ALUNOS= "Select * from TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND  NU_Unidade = "& unidade
		
	if curso="999990" or curso="" or isnull(curso) then
		SQL_CURSO=""
	else
		SQL_CURSO=" AND CO_Curso = '"& curso &"'"
	end if

	if co_etapa="999990" or co_etapa="" or isnull(co_etapa) then
		SQL_ETAPA=""
	else
		SQL_ETAPA=" AND CO_Etapa = '"& co_etapa &"'"
	end if

	if turma="999990" or turma="" or isnull(turma) then
		SQL_TURMA=""
	else
		SQL_TURMA=" AND CO_Turma = '"& turma &"' "
	end if

SQL_ALUNOS= SQL_ALUNOS&SQL_CURSO&SQL_ETAPA&SQL_TURMA&" order by NU_Chamada"
end if


if tp_ocor=999999 or tp_ocor="999999" then
	SQL_TP_OCORRENCIAS=""
ELSE
	SQL_TP_OCORRENCIAS="CO_Ocorrencia ="& tp_ocor&" AND"
end if
if qtd_ocor=0 or qtd_ocor="0" then
	SQL_QTD_OCORRENCIAS=""
else
	qtd_ocor=qtd_ocor*1
	minimo_ocorrencia=qtd_ocor-1
	SQL_QTD_OCORRENCIAS="HAVING COUNT(*)>"&minimo_ocorrencia
end if
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
// Criação do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicitação HTTP. O primeiro parâmetro informa o método post/get
// O segundo parâmetro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicitação síncrona, o parâmetro deve ser false
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=c", true);
// Para solicitações utilizando o método post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A função abaixo é executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto já completou a solicitação
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto é gerado no arquivo executa.asp e colocado no div
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
document.all.divEtapa.innerHTML ="<select class=select_style></select>"
document.all.divTurma.innerHTML = "<select class=select_style></select>"
//recuperarEtapa()
                                                           }
                                               }
// Abaixo é enviada a solicitação. Note que a configuração
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {
// Criação do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicitação HTTP. O primeiro parâmetro informa o método post/get
// O segundo parâmetro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicitação síncrona, o parâmetro deve ser false
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=e", true);
// Para solicitações utilizando o método post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A função abaixo é executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto já completou a solicitação
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto é gerado no arquivo executa.asp e colocado no div
                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select class=select_style></select>"
//recuperarTurma()
                                                           }
                                               }
// Abaixo é enviada a solicitação. Note que a configuração
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {
// Criação do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicitação HTTP. O primeiro parâmetro informa o método post/get
// O segundo parâmetro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicitação síncrona, o parâmetro deve ser false
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t", true);
// Para solicitações utilizando o método post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A função abaixo é executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto já completou a solicitação
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto é gerado no arquivo executa.asp e colocado no div
                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t																	   
                                                           }
                                               }
// Abaixo é enviada a solicitação. Note que a configuração
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}								   
                         </script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../../img/menu_r1_c2_f3.gif','../../../../img/menu_r1_c2_f2.gif','../../../../img/menu_r1_c2_f4.gif','../../../../img/menu_r1_c4_f3.gif','../../../../img/menu_r1_c4_f2.gif','../../../../img/menu_r1_c4_f4.gif','../../../../img/menu_r1_c6_f3.gif','../../../../img/menu_r1_c6_f2.gif','../../../../img/menu_r1_c6_f4.gif','../../../../img/menu_r1_c8_f3.gif','../../../../img/menu_r1_c8_f2.gif','../../../../img/menu_r1_c8_f4.gif','../../../../img/menu_direita_r2_c1_f3.gif','../../../../img/menu_direita_r2_c1_f2.gif','../../../../img/menu_direita_r2_c1_f4.gif','../../../../img/menu_direita_r4_c1_f3.gif','../../../../img/menu_direita_r4_c1_f2.gif','../../../../img/menu_direita_r4_c1_f4.gif','../../../../img/menu_direita_r6_c1_f3.gif','../../../../img/menu_direita_r6_c1_f2.gif','../../../../img/menu_direita_r6_c1_f4.gif')">
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
                <tr> 
                    
    <td height="10"> 
      <%	call mensagens(4,636,0,0) %>
</td></tr>
<tr>
                    
    <td height="10"> 
      <%	call mensagens(4,9706,0,0) 
	  
	  
%>
</td></tr>

<tr>

            <td valign="top"> 
		<%
mes = DatePart("m", now) 
dia = DatePart("d", now) 



dia=dia*1
mes=mes*1
%>	
<FORM name="formulario" METHOD="POST" ACTION="altera.asp?ori=1">
                
        <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td width="653" height="15" class="tb_tit">Informe os crit&eacute;rios 
              para pesquisa 
              <input name="co_grupo" type="hidden" id="co_grupo" value="<% = co_grupo %>"></td>
          </tr>
          <tr> 
            <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="247" class="tb_subtit"> 
                    <div align="center">Tipo de Ocorr&ecirc;ncia</div></td>
                  <td width="253" class="tb_subtit">
<div align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Per&iacute;odo da 
                      Ocorr&ecirc;ncia</div></td>
                  <td class="tb_subtit">&nbsp;</td>
                  <td width="253" class="tb_subtit">Quantidade m&iacute;nima de ocorr&ecirc;ncia</td>
                </tr>
                <tr> 
                  <td width="247"> 
                    <div align="center"><font class="form_dado_texto"> 
                      <select name="tp_ocor" class="textInput" id="tp_ocor">
                      <option value="999999" selected>Selecione um tipo de ocorrência</option>
                      <%
 
 		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Tipo_Ocorrencia order by NO_Ocorrencia"
		RS1.Open SQL1, CON0

While not RS1.EOF
co_ocorrencia=RS1("CO_Ocorrencia")
no_ocorrencia=RS1("NO_Ocorrencia")
tp_ocor=tp_ocor*1
co_ocorrencia=co_ocorrencia*1
if co_ocorrencia=tp_ocor then
%>
                      <option value="<%=co_ocorrencia%>" selected> 
                      <%Response.Write(no_ocorrencia)%>
                      </option>
                      <%
else
%>
                      <option value="<%=co_ocorrencia%>"> 
                      <%Response.Write(no_ocorrencia)%>
                      </option>
                      <%
end if
RS1.Movenext
WEND
%>
                    </select> 
                      </font></div></td>					  
                  <td colspan="2"><div align="center"><font class="form_dado_texto"> 
                   <select name="dia_de" id="dia_de" class="select_style">
                       <% 
							 For i =1 to 31
							 dia_de=dia_de*1
							 if dia_de=i then 
								if dia_de<10 then
								dia_de="0"&dia_de
								end if
							 %>
                              <option value="<%response.Write(dia_de)%>" selected> 
                              <%response.Write(dia_de)%>
                              </option>
                              <% else
									if i<10 then
									i="0"&i
								end if
							%>
                              <option value="<%response.Write(i)%>"> 
                              <%response.Write(i)%>
                              </option>
                              <% end if 
							next
							%>
                      </select>
                      / 
                      <select name="mes_de" id="mes_de" class="select_style">
                              <%mes_de=mes_de*1
								if mes_de="1" or mes_de=1 then%>
                              <option value="1" selected>janeiro</option>
                              <% else%>
                              <option value="1">janeiro</option>
                              <%end if
								if mes_de="2" or mes_de=2 then%>
                              <option value="2" selected>fevereiro</option>
                              <% else%>
                              <option value="2">fevereiro</option>
                              <%end if
								if mes_de="3" or mes_de=3 then%>
                              <option value="3" selected>mar&ccedil;o</option>
                              <% else%>
                              <option value="3">mar&ccedil;o</option>
                              <%end if
								if mes_de="4" or mes_de=4 then%>
                              <option value="4" selected>abril</option>
                              <% else%>
                              <option value="4">abril</option>
                              <%end if
								if mes_de="5" or mes_de=5 then%>
                              <option value="5" selected>maio</option>
                              <% else%>
                              <option value="5">maio</option>
                              <%end if
								if mes_de="6" or mes_de=6 then%>
                              <option value="6" selected>junho</option>
                              <% else%>
                              <option value="6">junho</option>
                              <%end if
								if mes_de="7" or mes_de=7 then%>
                              <option value="7" selected>julho</option>
                              <% else%>
                              <option value="7">julho</option>
                              <%end if%>
                              <%if mes_de="8" or mes_de=8 then%>
                              <option value="8" selected>agosto</option>
                              <% else%>
                              <option value="8">agosto</option>
                              <%end if
								if mes_de="9" or mes_de=9 then%>
                              <option value="9" selected>setembro</option>
                              <% else%>
                              <option value="9">setembro</option>
                              <%end if
								if mes_de="10" or mes_de=10 then%>
                              <option value="10" selected>outubro</option>
                              <% else%>
                              <option value="10">outubro</option>
                              <%end if
								if mes_de="11" or mes_de=11 then%>
                              <option value="11" selected>novembro</option>
                              <% else%>
                              <option value="11">novembro</option>
                              <%end if
								if mes_de="12" or mes_de=12 then%>
                              <option value="12" selected>dezembro</option>
                              <% else%>
                              <option value="12">dezembro</option>
                              <%end if%>
                      </select>
                      / 
                      <%response.write(ano_letivo)%>
                      at&eacute; 
                        <select name="dia_ate" id="dia_ate" class="select_style">
                       <% 
							 For i =1 to 31
							 dia_ate=dia_ate*1
							 if dia_ate=i then 
								if dia_ate<10 then
								dia_ate="0"&dia_ate
								end if
							 %>
                              <option value="<%response.Write(dia_ate)%>" selected> 
                              <%response.Write(dia_ate)%>
                              </option>
                              <% else
								if i<10 then
								i="0"&i
								end if
							%>
                              <option value="<%response.Write(i)%>"> 
                              <%response.Write(i)%>
                              </option>
                              <% end if 
							next
							%>
                      </select>
                      / 
                      <select name="mes_ate" id="mes_ate" class="select_style">
                              <%mes_ate=mes_ate*1
								if mes_ate="1" or mes_ate=1 then%>
                              <option value="1" selected>janeiro</option>
                              <% else%>
                              <option value="1">janeiro</option>
                              <%end if
								if mes_ate="2" or mes_ate=2 then%>
                              <option value="2" selected>fevereiro</option>
                              <% else%>
                              <option value="2">fevereiro</option>
                              <%end if
								if mes_ate="3" or mes_ate=3 then%>
                              <option value="3" selected>mar&ccedil;o</option>
                              <% else%>
                              <option value="3">mar&ccedil;o</option>
                              <%end if
								if mes_ate="4" or mes_ate=4 then%>
                              <option value="4" selected>abril</option>
                              <% else%>
                              <option value="4">abril</option>
                              <%end if
								if mes_ate="5" or mes_ate=5 then%>
                              <option value="5" selected>maio</option>
                              <% else%>
                              <option value="5">maio</option>
                              <%end if
								if mes_ate="6" or mes_ate=6 then%>
                              <option value="6" selected>junho</option>
                              <% else%>
                              <option value="6">junho</option>
                              <%end if
								if mes_ate="7" or mes_ate=7 then%>
                              <option value="7" selected>julho</option>
                              <% else%>
                              <option value="7">julho</option>
                              <%end if%>
                              <%if mes_ate="8" or mes_ate=8 then%>
                              <option value="8" selected>agosto</option>
                              <% else%>
                              <option value="8">agosto</option>
                              <%end if
								if mes_ate="9" or mes_ate=9 then%>
                              <option value="9" selected>setembro</option>
                              <% else%>
                              <option value="9">setembro</option>
                              <%end if
								if mes_ate="10" or mes_ate=10 then%>
                              <option value="10" selected>outubro</option>
                              <% else%>
                              <option value="10">outubro</option>
                              <%end if
								if mes_ate="11" or mes_ate=11 then%>
                              <option value="11" selected>novembro</option>
                              <% else%>
                              <option value="11">novembro</option>
                              <%end if
								if mes_ate="12" or mes_ate=12 then%>
                              <option value="12" selected>dezembro</option>
                              <% else%>
                              <option value="12">dezembro</option>
                              <%end if%>
                      </select>
                      / 
                      <%response.write(ano_letivo)%>
                      </font></div></td>
 
                  <td width="253"><select name="qtd_ocor" id="qtd_ocor" class="select_style">
                  <%qtd_ocor=qtd_ocor*1
				  if qtd_ocor=0 then
				  %>
                        <option value="0" selected>00</option>
                  <%else%>
                  		
                  <%end if 
				  if qtd_ocor=1 then
				  %>
                        <option value="1" selected>01</option>
                  <%else%>
                  		<option value="1">01</option>
                  <%end if    
				  if qtd_ocor=2 then
				  %>
                        <option value="2" selected>02</option>
                  <%else%>
                  		<option value="2">02</option>
                  <%end if                                              
				  if qtd_ocor=3 then
				  %>
                        <option value="3" selected>03</option>
                  <%else%>
                  		<option value="3">03</option>
                  <%end if    
 				  if qtd_ocor=4 then
				  %>
                        <option value="4" selected>04</option>
                  <%else%>
                  		<option value="4">04</option>
                  <%end if  
				  if qtd_ocor=5 then
				  %>
                        <option value="5" selected>05</option>
                  <%else%>
                  		<option value="5">05</option>
                  <%end if                         
                   if qtd_ocor=6 then
				  %>
                        <option value="6" selected>06</option>
                  <%else%>
                  		<option value="6">06</option>
                  <%end if
                   if qtd_ocor=7 then
				  %>
                        <option value="7" selected>07</option>
                  <%else%>
                  		<option value="7">07</option>
                  <%end if
				  if qtd_ocor=8 then
				  %>
                        <option value="8" selected>08</option>
                  <%else%>
                  		<option value="8">08</option>
                  <%end if
				  if qtd_ocor=9 then
				  %>
                        <option value="9" selected>09</option>
                  <%else%>
                  		<option value="9">09</option>
                  <%end if
					for i=10 to 100					
						if qtd_ocor=i then
						%>     
							<option value="<%response.Write(i)%>" selected><%response.Write(i)%></option>
						<%ELSE%>     
							<option value="<%response.Write(i)%>"><%response.Write(i)%></option>
					   <%end if
				   next%>                         
                  </select></td>
                </tr>
                <tr> 
                  <td width="247">&nbsp;</td>
                  <td colspan="2">&nbsp;</td>
                  <td width="253">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="247" class="tb_subtit"> 
                    <div align="center">UNIDADE 
                    </div></td>
                  <td width="253" class="tb_subtit"> 
                    <div align="center">CURSO 
                    </div></td>
                  <td width="247" class="tb_subtit"> <div align="center">ETAPA 
                    </div></td>
                  <td width="253" class="tb_subtit"> 
                    <div align="center">TURMA 
                    </div></td>
                </tr>
                <tr>
                  <td background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF"><div align="center">
                    <select name="unidade" class="select_style" onChange="recuperarCurso(this.value)">
					<% if unidade=999990 or unidade="999990" then%>
                         <option value="999990" selected></option>
                      <%else%>
                         <option value="999990"></option>					  
					  <%end if	
					Set RS6 = Server.CreateObject("ADODB.Recordset")
					SQL6 = "SELECT * FROM TB_Unidade order by NO_Abr"
					RS6.Open SQL6, CON0
					While not RS6.EOF
					NU_Unidade = RS6("NU_Unidade")
					NO_Abr = RS6("NO_Abr")
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
					RS6.MOVENEXT
					WEND
					%>
                    </select>
                  </div></td>
                  <td background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF"><div align="center">
                    <div id="divCurso">
				<% if unidade=999990 or unidade="999990" then%>
                    <select class="select_style">
                    </select>                
				<%else%>
                      <select name="curso" class="select_style" onChange="recuperarEtapa(this.value)">
					<% if curso=999990 or curso="999990" then%>
                         <option value="999990" selected></option>
                      <%else%>
                         <option value="999990"></option>					  
					  <%end if	                      		
					Set RS7 = Server.CreateObject("ADODB.Recordset")
					SQL7 = "SELECT Distinct CO_Curso FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade
					RS7.Open SQL7, CON0
					
					While not RS7.EOF
						CO_Curso = RS7("CO_Curso")
				
						Set RS7a = Server.CreateObject("ADODB.Recordset")
						SQL7a = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
						RS7a.Open SQL7a, CON0
				
						NO_Curso = RS7a("NO_Abreviado_Curso")		
		
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
						RS7.MOVENEXT
						WEND
						%>
                      </select>
                    <%end if%>  
                    </div>
                  </div></td>
                  <td background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF"><div align="center">
                    <div id="divEtapa">
				<% if (unidade=999990 or unidade="999990") or isnull(curso) then%>
                    <select class="select_style">
                    </select>                
				<%else%>                    
                      <select name="etapa" class="select_style" onChange="recuperarTurma(this.value)">  
					<% if co_etapa=999990 or co_etapa="999990" then%>
                         <option value="999990" selected></option>
                      <%else%>
                         <option value="999990"></option>					  
					  <%end if	      	
					Set RS0b = Server.CreateObject("ADODB.Recordset")
					SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"'"
					RS0b.Open SQL0b, CON0
			
					While not RS0b.EOF
					Etapa = RS0b("CO_Etapa")
					
					
							Set RS8a = Server.CreateObject("ADODB.Recordset")
							SQL8a = "SELECT * FROM TB_Etapa where CO_Curso='"&curso&"' AND CO_Etapa='"&Etapa&"'"
							RS8a.Open SQL8a, CON0
							
					NO_Etapa = RS8a("NO_Etapa")		
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
                    <%end if%>  
                    </div>
                  </div></td>
                  <td><div id="divTurma" align="center"> 
				<% if unidade=999990 or unidade="999990" then%>
                    <select class="select_style">
                    </select>                
				<%else%>                  
                  <select name="turma" class="select_style">
                        <option value="999990" selected></option> 
				<%
                    Set RS9 = Server.CreateObject("ADODB.Recordset")
                    SQL9 = "SELECT DISTINCT CO_Turma FROM TB_Turma where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"' AND CO_Etapa='" & co_etapa & "' order by CO_Turma" 
                    RS9.Open SQL9, CON0
                                                
					while not RS9.EOF
					co_turma= RS9("CO_Turma")
						if co_turma = turma then
						 %>
						<option value="<%=co_turma%>" selected> 
							<%response.Write(co_turma)%>
						</option>
							<%
						else	
						 %>
						<option value="<%=co_turma%>"> 
							<%response.Write(co_turma)%>
						</option>
							<%
						end if
					RS9.MOVENEXT
					WEND
					%>   
 				</select>
                <%end if%>
 </div></td>
                </tr>
                <tr> 
                  <td colspan="4"><hr width="1000"></td>
                </tr>
                <tr> 
                  <td><div align="center"> </div></td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td><div align="center">
                      <input name="SUBMIT" type=SUBMIT class="botao_prosseguir" value="Prosseguir">
                    </div></td>
                </tr>
                <tr> 
                  <td colspan="4">&nbsp;</td>
                </tr>                
                <tr>
                  <td colspan="4">
<%	

if SQL_ALUNOS="NULO" then
	SQL_MATRICULAS="" 
else

nu_chamada_check = 1
	Set RSA = Server.CreateObject("ADODB.Recordset")
	CONEXAOA = SQL_ALUNOS
	Set RSA = CON1.Execute(CONEXAOA)
	vetor_matriculas="" 
	While Not RSA.EOF
		nu_matricula = RSA("CO_Matricula")
		nu_chamada = RSA("NU_Chamada")
			
		if nu_chamada_check = 1 and nu_chamada=nu_chamada_check then
			vetor_matriculas=nu_matricula
		elseif nu_chamada_check = 1 then
			while nu_chamada_check < nu_chamada
				nu_chamada_check=nu_chamada_check+1
			wend 
			vetor_matriculas=nu_matricula
		else
			vetor_matriculas=vetor_matriculas&","&nu_matricula
		end if
		nu_chamada_check=nu_chamada_check+1		
	RSA.MoveNext
	Wend 
	SQL_MATRICULAS="CO_Matricula IN("& vetor_matriculas&") AND" 		
end if	
	
%>	                  
                  
                  <table width="1000" border="0" cellspacing="0" cellpadding="0">

<%
		Set RSo = Server.CreateObject("ADODB.Recordset")
		SQLo = "SELECT  CO_Matricula, CO_Assunto, CO_Ocorrencia, COUNT(*) AS Num_Ocorrencias FROM TB_Ocorrencia_Aluno WHERE "&SQL_MATRICULAS&" "&SQL_TP_OCORRENCIAS&" (DA_Ocorrencia BETWEEN #"&data_de&"# AND #"&data_ate&"#) GROUP BY CO_Matricula, CO_Assunto, CO_Ocorrencia "&SQL_QTD_OCORRENCIAS&" order BY COUNT(*) Desc"
		RSo.Open SQLo, CON3, 3, 3

    if cint(Request.QueryString("pagina"))<1 then
	intpagina = 1
    else
		if cint(Request.QueryString("pagina"))>RSo.PageCount then  
	    intpagina = RSo.PageCount
        else
    	intpagina = Request.QueryString("pagina")
		end if
    end if   
	

check=2
IF RSo.eof then
sem_link=1
%> 
  <tr>
    <th colspan="9"> <font class="form_dado_texto">Não foram encontradas ocorrências que atendessem aos critérios informados.</font></th>
  </tr>
<% 
else
sem_link=0
RSo.PageSize = 30
RSo.AbsolutePage = intpagina
%> 
  <tr>
    <th width="110" scope="col" class="tb_subtit"><div align="center">Unidade</div></th>
    <th width="70" scope="col" class="tb_subtit"><div align="center">Curso</div></th>
    <th width="70" scope="col" class="tb_subtit"><div align="center">Etapa</div></th>
    <th width="70" scope="col" class="tb_subtit"><div align="center">Turma</div></th>
    <th width="50" scope="col" class="tb_subtit">Chamada</th>
    <th width="50" scope="col" class="tb_subtit">Matr&iacute;cula</th>
    <th width="300" scope="col" class="tb_subtit">Nome</th>
    <th width="250" scope="col" class="tb_subtit">Ocorr&ecirc;ncia</th>
    <th width="20" scope="col" class="tb_subtit">Qtd</th>
  </tr>
<% 
	intrec = 0
	WHILE intrec<RSo.PageSize and NOT RSo.eof
	nu_matric=RSo("CO_Matricula")
	tipo_assunto=RSo("CO_Assunto")
	tipo_ocorrencia=RSo("CO_Ocorrencia")
	num_ocorrencia=RSo("Num_Ocorrencias")
	
			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL2 = "SELECT  * FROM TB_Tipo_Ocorrencia WHERE CO_Assunto= '"&tipo_assunto&"' AND CO_Ocorrencia="&tipo_ocorrencia&""
			RS2.Open SQL2, CON0
	
	nome_ocorrencia=RS2("NO_Ocorrencia")
	
			Set RS3 = Server.CreateObject("ADODB.Recordset")
			SQL3= "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& nu_matric
			RS3.Open SQL3, CON1
	IF RS3.EOF then
	no_unidade = ""
	no_curso = ""
	no_etapa = ""
	else	
	unidade= RS3("NU_Unidade")
	curso= RS3("CO_Curso")
	etapa= RS3("CO_Etapa")
	turma= RS3("CO_Turma")
	nu_chamada= RS3("NU_Chamada")
	
	call GeraNomes("PORT",unidade,curso,etapa,CON0)
	no_unidade = session("no_unidades")
	no_etapa = session("no_serie")
	
			Set RS5 = Server.CreateObject("ADODB.Recordset")
			Sql5= "SELECT * FROM TB_Curso where CO_Curso = '"& curso &"'"
			Set RS5= CON0.Execute ( Sql5 ) 
		IF RS5.eof THEN
			no_curso=""
		ELSE
			no_curso= RS5("NO_Abreviado_Curso")
		END IF
		
	end if
			
			Set RS4 = Server.CreateObject("ADODB.Recordset")
			SQL4= "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& nu_matric
			RS4.Open SQL4, CON1
			
	if RS4.EOF then
	nome_aluno = "Sem nome cadastrado"
	else
	nome_aluno = RS4("NO_Aluno")
	end if		
	
	if check mod 2 =0 then
		cor = "tb_fundo_linha_par" 
	else
		cor ="tb_fundo_linha_impar"
	end if 
	
	%>  
	  <tr class="<%=cor%>">
		<td width="110"><div align="center"><%response.Write(no_unidade)%></div></td>
		<td width="70"><div align="center"><%response.Write(no_curso)%></div></td>
		<td width="70"><div align="center"><%response.Write(no_etapa)%></div></td>
		<td width="70"><div align="center"><%response.Write(turma)%></div></td>
		<td width="50"><div align="center"><%response.Write(nu_chamada)%></div></td>
		<td width="50"><div align="center"><%response.Write(nu_matric)%></div></td>
		<td width="300"><div align="center"><%response.Write(nome_aluno)%></div></td>
		<td width="250"><div align="center"><%response.Write(nome_ocorrencia)%></div></td>
		<td width="20"><div align="center"><%response.Write(num_ocorrencia)%></div></td>
	  </tr>
	<%
	intrec = intrec + 1
	check=check+1
	RSo.MOVENEXT
	wend
END iF	%>
	  <tr class="<%=cor%>">
	    <td colspan="9">
        <table width="100%">
                <tr>
          <td><div align="center">
		  <%for i=1 to RSo.PageCount
		  intpagina=intpagina*1
			  if i=intpagina then%>
			  <font class="form_dado_texto"><%response.Write(intpagina)%></font>
			  <%else%>
			   <a href="altera.asp?pagina=<%=response.Write(i)%>" class="linkPaginacao"><%response.Write(i)%></a> 
			  <%end if
		  next
		  %></div></td>
        </tr>    
        <tr> 
          <td class="tb_tit"><div align="center">
              <%
			  
if sem_link=0 then
	%>&nbsp;<%		  
			    if intpagina>1 then
    %>
              <a href="altera.asp?pagina=<%=intpagina-1%>" class="linktres">Anterior</a> 
              <%
    end if
    if StrComp(intpagina,RSo.PageCount)<>0 then  
    %>
              <a href="altera.asp?pagina=<%=intpagina + 1%>" class="linktres">Próximo</a> 
              <%
    end if
else	
	%>&nbsp;<%
end if	

    %>
            </div></td>
        </tr>
 <%
     RSo.close
    Set RSo = Nothing
%>       
      </table>  
        
        </td>
	    </tr>
</table>
</td>
                </tr>
              </table>
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
<%

call GravaLog (nvg,outro)

If Err.number<>0 then
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