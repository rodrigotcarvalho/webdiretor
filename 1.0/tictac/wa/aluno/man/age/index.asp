<%On Error Resume Next%>
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

chave=nvg
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
if opt="" or isnull("opt") then
	opt="sel"
else
	opt=opt
	if opt="ok" then
		cod_cons = request.QueryString("cod_cons")
		co_usr_prof = request.QueryString("co_usr_prof")
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
//    if (document.busca.busca1.value == "" && document.busca.busca2.value == "")
//  {    alert("Por favor digite uma opção de busca!")
//    document.busca.busca1.focus()
//    return false
//  }
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
   var f=document.forms[3]; 
      f.submit(); 
}
//-->
</script>
</head>
<% if opt="listall" or opt="list" then%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%else %>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('document.busca.busca1.focus()')">
<%end if %>
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
        <%if opt="sel" then%>
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,300,0,0) %>
    </td>
			  </tr>			  
        <form action="index.asp?opt=list&nvg=<%=nvg%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
          <tr class="tb_tit"> 
            
      <td height="10" colspan="5">Preencha um dos campos abaixo</td>
          </tr>
          <TR>
		  
      <td height="26" valign="top"> 
        <table width="1000" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            
            <td width="150"  height="10"> 
              <div align="right"><font class="form_dado_texto"> Matr&iacute;cula:</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                </strong></font></div></td>
            
            <td width="50" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="busca1" type="text" class="textInput" id="busca1" size="12">
              </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font></td>
            
            <td width="150" height="10"> 
              <div align="right"><font class="form_dado_texto"> Nome: </font></div></td>
            
            <td width="500" height="10" ><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="busca2" type="text" class="textInput" id="busca2" size="55" maxlength="50">
              </font></td>
            
            <td width="150" height="10">&nbsp;</td>
          </tr>
          <tr>
            <td  height="10" colspan="5">&nbsp;</td>
          </tr>
          <tr>
            <td  height="10" colspan="5"><table width="1000" border="0" cellspacing="0" cellpadding="0" height="46">
              <tr class="tb_subtit">
                <td width="500" height="13">Per&iacute;odo</td>
                <td width="300">Status da Entrevista</td>
                <td width="200"><div align="center">Ordenado por:</div></td>
                </tr>
              <tr>
                <td width="500" height="23" class="form_dado_texto"> De</font>&nbsp;
                  <select name="dia_de" id="dia_de" class="select_style">
<%
aa = DatePart("yyyy", now)
mm = DatePart("m", now) 
dd = DatePart("d", now) 
	for dia= 1 to 31
		if dd=dia then
			dia_selected = "selected"
		else
			dia_selected = ""	
		end if	
		
		if dia<10 then
			dia_txt="0"&dia
		else	
			dia_txt=dia		
		end if		
	
%>                  
                    <option value="<%response.Write(dia)%>" <%response.Write(dia_selected)%>><%response.Write(dia_txt)%></option>
<%next%>
                  </select>
                  /
                  <select name="mes_de" id="mes_de" class="select_style">                  
 <%

	for mes= 1 to 12
		if mm=mes then
			mes_selected = "selected"
		else
			mes_selected = ""	
		end if	
		
		Select case mes
		
			case 1
			mes_txt="janeiro"
			
			case 2
			mes_txt="fevereiro"
			
			case 3
			mes_txt="mar&ccedil;o"
			
			case 4
			mes_txt="abril"
			
			case 5
			mes_txt="maio"
	
			case 6
			mes_txt="junho"
			
			case 7
			mes_txt="julho"
			
			case 8
			mes_txt="agosto"		
			
			case 9
			mes_txt="setembro"
	
			case 10
			mes_txt="outubro"
			
			case 11
			mes_txt="novembro"
			
			case 12
			mes_txt="dezembro"				
		end select			
	
%>                     

                    <option value="<%response.Write(mes)%>" <%response.Write(mes_selected)%>><%response.Write(mes_txt)%></option>
<%next%>
                  </select>
                  /
                  <select name="ano_de" class="select_style" id="ano_de">
                    <%
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Ano_Letivo order by NU_Ano_Letivo"
		RS0.Open SQL0, CON
		while not RS0.EOF 
		ano_bd=RS0("NU_Ano_Letivo")
		
				ano_letivo=ano_letivo*1
		ano_bd=ano_bd*1

				if ano_letivo=ano_bd then%>
                    <option value="<%=ano_bd%>" selected><%=ano_bd%></option>
                    <%else%>
                    <option value="<%=ano_bd%>"><%=ano_bd%></option>
                    <%end if
		RS0.MOVENEXT
		WEND 		
			prox_ano= ano_bd+1	
				%>
					            <option value="<%=prox_ano%>"><%=prox_ano%></option>    
                  </select>
                  <!-- &agrave;s 
                    <select name="hora_de" id="select6" class="select_style">
                      <option value="00" selected>00</option>
                      <option value="1" >01</option>
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
                    </select>
                    : 
                    <select name="min_de" id="select8" class="select_style">
                      <option value="00" selected>00</option>
                      <option value="1">01</option>
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
                      <option value="32">32</option>
                      <option value="33">33</option>
                      <option value="34">34</option>
                      <option value="35">35</option>
                      <option value="36">36</option>
                      <option value="37">37</option>
                      <option value="38">38</option>
                      <option value="39">39</option>
                      <option value="40">40</option>
                      <option value="41">41</option>
                      <option value="42">42</option>
                      <option value="43">43</option>
                      <option value="44">44</option>
                      <option value="45">45</option>
                      <option value="46">46</option>
                      <option value="47">47</option>
                      <option value="48">48</option>
                      <option value="49">49</option>
                      <option value="50">50</option>
                      <option value="51">51</option>
                      <option value="52">52</option>
                      <option value="53">53</option>
                      <option value="54">54</option>
                      <option value="55">55</option>
                      <option value="56">56</option>
                      <option value="57">57</option>
                      <option value="58">58</option>
                      <option value="59">59</option>
                    </select>-->
                  At&eacute;
                  <select name="dia_ate" id="select4" class="select_style">
                    <option value="1">01</option>
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
                    <option value="23">22</option>
                    <option value="24">24</option>
                    <option value="25">25</option>
                    <option value="26">26</option>
                    <option value="27">27</option>
                    <option value="28">28</option>
                    <option value="29">29</option>
                    <option value="30">30</option>
                    <option value="31" selected>31</option>
                  </select>
                  /
                  <select name="mes_ate" id="select5" class="select_style">
                    <option value="1">janeiro</option>
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
                    <option value="12" selected>dezembro</option>
                  </select>
                  /
                  <select name="ano_ate" class="select_style" id="ano_ate" >
                    <%
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Ano_Letivo order by NU_Ano_Letivo"
		RS0.Open SQL0, CON
		while not RS0.EOF 
		ano_bd=RS0("NU_Ano_Letivo")

		
		ano_letivo=ano_letivo*1
		ano_bd=ano_bd*1
				if ano_letivo=ano_bd then%>
                    <option value="<%=ano_bd%>" selected><%=ano_bd%></option>
                    <%else%>
                    <option value="<%=ano_bd%>"><%=ano_bd%></option>
                    <%end if
		RS0.MOVENEXT
		WEND 		
			prox_ano= ano_bd+1	
				%>
					            <option value="<%=prox_ano%>"><%=prox_ano%></option>    
                  </select>
                  <!--&agrave;s 
                    <select name="hora_ate" id="select11" class="select_style">
                      <option value="00">00</option>
                      <option value="1">01</option>
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
                      <option value="23" selected>23</option>
                    </select>
                    : 
                    <select name="min_ate" id="select10" class="select_style">
                      <option value="0">00</option>
                      <option value="1">01</option>
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
                      <option value="32">32</option>
                      <option value="33">33</option>
                      <option value="34">34</option>
                      <option value="35">35</option>
                      <option value="36">36</option>
                      <option value="37">37</option>
                      <option value="38">38</option>
                      <option value="39">39</option>
                      <option value="40">40</option>
                      <option value="41">41</option>
                      <option value="42">42</option>
                      <option value="43">43</option>
                      <option value="44">44</option>
                      <option value="45">45</option>
                      <option value="46">46</option>
                      <option value="47">47</option>
                      <option value="48">48</option>
                      <option value="49">49</option>
                      <option value="50">50</option>
                      <option value="51">51</option>
                      <option value="52">52</option>
                      <option value="53">53</option>
                      <option value="54">54</option>
                      <option value="55">55</option>
                      <option value="56">56</option>
                      <option value="57">57</option>
                      <option value="58">58</option>
                      <option value="59" selected>59</option>
                    </select>
                    --></td>
                <td width="300" class="form_dado_texto"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                  <tr class="form_dado_texto">
                    <td width="25"><input name="status" type="checkbox" id="status" value="1"></td>
                    <td width="142"> Atendidas </td>
                    <td width="25"><input name="status" type="checkbox" id="status" value="3"></td>
                    <td width="142">Pendentes </td>
                    <td width="25"><input name="status" type="checkbox" id="status" value="2"></td>
                    <td width="141">Canceladas</td>
                    </tr>
                </table></td>
                <td width="200" align="center" class="tb_corpo">
                  <select name="ordem" class="select_style">
                        <% if ordem="dt" then%>
                        <option value="dt" selected>Data/Hora</option>
                        <%else%>
                        <option value="dt" >Data/Hora</option>
                        <%end if%>
                        <% if ordem="mt" then%>
                        <option value="mt" selected>Matr&iacute;cula</option>
                        <%else%>
                        <option value="mt" >Matr&iacute;cula</option>
                        <%end if%>
                        <% if ordem="al" then%>
                        <option value="al" selected>Nome Aluno</option>
                        <%else%>
                        <option value="al" >Nome Aluno</option>
                        <%end if%>                                                
                        <% if ordem="en" then%>
                        <option value="en" selected>Tipo de Entrevista</option>
                        <%else%>
                        <option value="en" >Tipo de Entrevista</option>
                        <%end if%>
                        <% if ordem="pr" then%>
                        <option value="pr" selected>Participantes</option>
                        <%else%>
                        <option value="pr" >Participantes</option>
                        <%end if%>
                        <% if ordem="at" then%>
                        <option value="at" selected>Atendido por</option>
                        <%else%>
                        <option value="at" >Atendido por</option>
                        <%end if%>
                        <% if ordem="st" then%>
                        <option value="st" selected>Status</option>
                        <%else%>
                        <option value="st" >Status</option>
                        <%end if%>                        
                      </select></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td  height="10" colspan="5"><div align="center">
              <hr>
            <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> </font></div></td>
          </tr>
          <tr>
            <td  height="10" colspan="5"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="33%">&nbsp;</td>
                <td width="34%">&nbsp;</td>
                <td width="33%" align="center"><font size="2" face="Arial, Helvetica, sans-serif">
                  <input name="Submit4" type="submit" class="botao_prosseguir" id="Submit3" value="Procurar">
                </font></td>
              </tr>
            </table></td>
          </tr>
		  </table>
		  </td>
		  </TR>
      </form>
	   <tr> 
            
      <td > 
	  </td>
          </tr>
 <%elseif opt="list" then
	    busca1=request.form("busca1") 
	    busca2=request.form("busca2")
		status_entrevista = request.form("status")
		if status_entrevista = "" or isnull(status_entrevista) then
		
		else
			session("status_entrevista")=	status_entrevista
		end if  
		
		dia_de = request.form("dia_de")
		if dia_de = "" or isnull(dia_de) then
		
		else
			session("dia_de")=	dia_de
		end if  		
		
		mes_de = request.form("mes_de")
		if mes_de = "" or isnull(mes_de) then
		
		else
			session("mes_de")=	mes_de
		end if  
		
		ano_de = request.form("ano_de")
		if ano_de = "" or isnull(ano_de) then
		
		else
			session("ano_de")=	ano_de
		end if  		

'		h_de= request.form("hora_de")
'		min_de= request.form("min_de")
'		
'		
'		hora_de=h_de&":"&min_de
		
'		data_inicio=dia_de&"/"&mes_de&"/"&ano_de	

		dia_ate = request.form("dia_ate")
		if dia_ate = "" or isnull(dia_ate) then
		
		else
			session("dia_ate")=	dia_ate
		end if  		
		
		mes_ate = request.form("mes_ate")
		if mes_ate = "" or isnull(mes_ate) then
		
		else
			session("mes_ate")=	mes_ate
		end if  
		
		ano_ate = request.form("ano_ate")
		if ano_ate = "" or isnull(ano_ate) then
		
		else
			session("ano_ate")=	ano_ate
		end if  		

		ordem = request.form("ordem")
		if ordem = "" or isnull(ordem) then
		
		else
			session("ordem")=	ordem
		end if  			

'		hora_ate= request.form("hora_ate")
'		min_ate= request.form("min_ate")	

	    
	  if busca1 ="" then
		  query = busca2
		  mensagem=304
	  elseif busca2 ="" then
		  query = busca1 
		  mensagem=303
	  end if 
	  if isnull(query) or query= "" then
		response.Redirect("resumo.asp?or=01&cod_cons=0")
	  else
		  if IsNumeric(query) Then
	  
			Set RS = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT * FROM TB_Alunos where CO_Matricula = "& query
			RS.Open SQL, CON1
			
			if RS.EOF Then
			%>
						<tr> 
						  
				<td height="10" colspan="5"> 
				  <%call mensagens(nivel,mensagem,1,0) %>
				</td>
						   </tr>
						<tr> 
						  
				<td height="10" colspan="5"> 
				  <%call mensagens(nivel,300,0,0) %>
				</td>
						  </tr>
					<form action="index.asp?opt=list&nvg=<%=nvg%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
							  <tr class="tb_tit"> 
								
				  <td height="10" colspan="5">Preencha um dos campos abaixo</td>
							  </tr>
							  <tr>
				  <TD height="26" valign="top"> 
					<Table width="1000" border="0" cellpadding="0" cellspacing="0">
					  <tr> 
								
						<td width="150"  height="10"> 
						  <div align="right"><font class="form_dado_texto"> Matr&iacute;cula: 
							</font></div></td>
								
						<td width="50"  height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
						  </font><font size="2" face="Arial, Helvetica, sans-serif"> 
						  <input name="busca1" type="text" class="textInput" id="busca1" size="12">
						  </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
						  </font></td>
								
						<td width="150" height="10"> 
						  <div align="right"><font class="form_dado_texto"> Nome: </font></div></td>
								
						<td width="500"  height="10" ><font size="2" face="Arial, Helvetica, sans-serif"> 
						  <input name="busca2" type="text" class="textInput" id="busca2" size="55" maxlength="50">
						  </font></td>
								
						<td width="150" height="10" >&nbsp;</td>
							  </tr>
					  <tr>
					    <td  height="10" colspan="5">&nbsp;</td>
				      </tr>
					  <tr>
					    <td  height="10" colspan="5"><table width="1000" border="0" cellspacing="0" cellpadding="0" height="46">
					      <tr class="tb_subtit">
					        <td width="500" height="13">Per&iacute;odo</td>
					        <td width="300">Status da Entrevista</td>
					        <td width="200"><div align="center">Ordenado por:</div></td>
				          </tr>
					      <tr>
					        <td width="500" height="23" class="form_dado_texto"> De</font>&nbsp;
                  <select name="dia_de" id="dia_de" class="select_style">
<%
aa = DatePart("yyyy", now)
mm = DatePart("m", now) 
dd = DatePart("d", now) 
	for dia= 1 to 31
		if dd=dia then
			dia_selected = "selected"
		else
			dia_selected = ""	
		end if	
		
		if dia<10 then
			dia_txt="0"&dia
		else	
			dia_txt=dia		
		end if		
	
%>                  
                    <option value="<%response.Write(dia)%>" <%response.Write(dia_selected)%>><%response.Write(dia_txt)%></option>
<%next%>
                  </select>
                  /
                  <select name="mes_de" id="mes_de" class="select_style">                  
 <%

	for mes= 1 to 12
		if mm=mes then
			mes_selected = "selected"
		else
			mes_selected = ""	
		end if	
		
		Select case mes
		
			case 1
			mes_txt="janeiro"
			
			case 2
			mes_txt="fevereiro"
			
			case 3
			mes_txt="mar&ccedil;o"
			
			case 4
			mes_txt="abril"
			
			case 5
			mes_txt="maio"
	
			case 6
			mes_txt="junho"
			
			case 7
			mes_txt="julho"
			
			case 8
			mes_txt="agosto"		
			
			case 9
			mes_txt="setembro"
	
			case 10
			mes_txt="outubro"
			
			case 11
			mes_txt="novembro"
			
			case 12
			mes_txt="dezembro"				
		end select			
	
%>                     

                    <option value="<%response.Write(mes)%>" <%response.Write(mes_selected)%>><%response.Write(mes_txt)%></option>
<%next%>
                  </select>
					          /
					          <select name="ano_de" class="select_style" id="ano_de">
					            <%
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Ano_Letivo order by NU_Ano_Letivo"
		RS0.Open SQL0, CON
		while not RS0.EOF 
		ano_bd=RS0("NU_Ano_Letivo")
		
				ano_letivo=ano_letivo*1
		ano_bd=ano_bd*1

				if ano_letivo=ano_bd then%>
					            <option value="<%=ano_bd%>" selected><%=ano_bd%></option>
					            <%else%>
					            <option value="<%=ano_bd%>"><%=ano_bd%></option>
					            <%end if
		RS0.MOVENEXT
		WEND 	
			prox_ano= ano_bd+1	
				%>
					            <option value="<%=prox_ano%>"><%=prox_ano%></option>                
				              </select>
					          <!-- &agrave;s 
                    <select name="hora_de" id="select6" class="select_style">
                      <option value="00" selected>00</option>
                      <option value="1" >01</option>
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
                    </select>
                    : 
                    <select name="min_de" id="select8" class="select_style">
                      <option value="00" selected>00</option>
                      <option value="1">01</option>
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
                      <option value="32">32</option>
                      <option value="33">33</option>
                      <option value="34">34</option>
                      <option value="35">35</option>
                      <option value="36">36</option>
                      <option value="37">37</option>
                      <option value="38">38</option>
                      <option value="39">39</option>
                      <option value="40">40</option>
                      <option value="41">41</option>
                      <option value="42">42</option>
                      <option value="43">43</option>
                      <option value="44">44</option>
                      <option value="45">45</option>
                      <option value="46">46</option>
                      <option value="47">47</option>
                      <option value="48">48</option>
                      <option value="49">49</option>
                      <option value="50">50</option>
                      <option value="51">51</option>
                      <option value="52">52</option>
                      <option value="53">53</option>
                      <option value="54">54</option>
                      <option value="55">55</option>
                      <option value="56">56</option>
                      <option value="57">57</option>
                      <option value="58">58</option>
                      <option value="59">59</option>
                    </select>-->
					          At&eacute;
					          <select name="dia_ate" id="dia_ate" class="select_style">
					            <option value="1">01</option>
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
					            <option value="23">22</option>
					            <option value="24">24</option>
					            <option value="25">25</option>
					            <option value="26">26</option>
					            <option value="27">27</option>
					            <option value="28">28</option>
					            <option value="29">29</option>
					            <option value="30">30</option>
					            <option value="31" selected>31</option>
				              </select>
					          /
					          <select name="mes_ate" id="mes_ate" class="select_style">
					            <option value="1">janeiro</option>
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
					            <option value="12" selected>dezembro</option>
				              </select>
					          /
					          <select name="ano_ate" class="select_style" id="ano_ate" >
					            <%
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Ano_Letivo order by NU_Ano_Letivo"
		RS0.Open SQL0, CON
		while not RS0.EOF 
		ano_bd=RS0("NU_Ano_Letivo")

		
		ano_letivo=ano_letivo*1
		ano_bd=ano_bd*1
				if ano_letivo=ano_bd then%>
					            <option value="<%=ano_bd%>" selected><%=ano_bd%></option>
					            <%else%>
					            <option value="<%=ano_bd%>"><%=ano_bd%></option>
					            <%end if
		RS0.MOVENEXT
		WEND 		
			prox_ano= ano_bd+1	
				%>
					            <option value="<%=prox_ano%>"><%=prox_ano%></option>         
				              </select>
				            <!--&agrave;s 
                    <select name="hora_ate" id="select11" class="select_style">
                      <option value="00">00</option>
                      <option value="1">01</option>
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
                      <option value="23" selected>23</option>
                    </select>
                    : 
                    <select name="min_ate" id="select10" class="select_style">
                      <option value="0">00</option>
                      <option value="1">01</option>
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
                      <option value="32">32</option>
                      <option value="33">33</option>
                      <option value="34">34</option>
                      <option value="35">35</option>
                      <option value="36">36</option>
                      <option value="37">37</option>
                      <option value="38">38</option>
                      <option value="39">39</option>
                      <option value="40">40</option>
                      <option value="41">41</option>
                      <option value="42">42</option>
                      <option value="43">43</option>
                      <option value="44">44</option>
                      <option value="45">45</option>
                      <option value="46">46</option>
                      <option value="47">47</option>
                      <option value="48">48</option>
                      <option value="49">49</option>
                      <option value="50">50</option>
                      <option value="51">51</option>
                      <option value="52">52</option>
                      <option value="53">53</option>
                      <option value="54">54</option>
                      <option value="55">55</option>
                      <option value="56">56</option>
                      <option value="57">57</option>
                      <option value="58">58</option>
                      <option value="59" selected>59</option>
                    </select>
                    --></td>
					        <td width="300" class="form_dado_texto"><table width="100%" border="0" cellpadding="0" cellspacing="0">
					          <tr class="form_dado_texto">
					            <td width="25"><input name="status" type="checkbox" id="status" value="1"></td>
					            <td width="142"> Atendidas </td>
					            <td width="25"><input name="status" type="checkbox" id="status" value="3"></td>
					            <td width="142">Pendentes </td>
					            <td width="25"><input name="status" type="checkbox" id="status" value="2"></td>
					            <td width="141">Canceladas</td>
				              </tr>
					          </table></td>
					        <td width="200" align="center" class="tb_corpo"><select name="ordem" class="select_style">
                        <% if ordem="dt" then%>
                        <option value="dt" selected>Data/Hora</option>
                        <%else%>
                        <option value="dt" >Data/Hora</option>
                        <%end if%>
                        <% if ordem="mt" then%>
                        <option value="mt" selected>Matr&iacute;cula</option>
                        <%else%>
                        <option value="mt" >Matr&iacute;cula</option>
                        <%end if%>
                        <% if ordem="al" then%>
                        <option value="al" selected>Nome Aluno</option>
                        <%else%>
                        <option value="al" >Nome Aluno</option>
                        <%end if%>                                                
                        <% if ordem="en" then%>
                        <option value="en" selected>Tipo de Entrevista</option>
                        <%else%>
                        <option value="en" >Tipo de Entrevista</option>
                        <%end if%>
                        <% if ordem="pr" then%>
                        <option value="pr" selected>Participantes</option>
                        <%else%>
                        <option value="pr" >Participantes</option>
                        <%end if%>
                        <% if ordem="at" then%>
                        <option value="at" selected>Atendido por</option>
                        <%else%>
                        <option value="at" >Atendido por</option>
                        <%end if%>
                        <% if ordem="st" then%>
                        <option value="st" selected>Status</option>
                        <%else%>
                        <option value="st" >Status</option>
                        <%end if%>                        
                      </select></td>
				          </tr>
					      </table></td>
				      </tr>
					  <tr>
					    <td  height="10" colspan="5"><div align="center">
					      <hr>
					      <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> </font></div></td>
				      </tr>
					  <tr>
					    <td  height="10" colspan="5"><table width="100%" border="0" cellspacing="0" cellpadding="0">
					      <tr>
					        <td width="33%">&nbsp;</td>
					        <td width="34%">&nbsp;</td>
					        <td width="33%" align="center"><font size="2" face="Arial, Helvetica, sans-serif">
					          <input name="Submit4" type="submit" class="botao_prosseguir" id="Submit4" value="Procurar">
					          </font></td>
				          </tr>
					      </table></td>
				      </tr>
				    </Table>
				  </TD>
				</tr>
					</form>
			 <tr>             
				  <td > 
				  </td>
					  </tr>
			<%ELSE		
				 response.Redirect("altera.asp?or=01&cod_cons="&query&"")
			  END IF
		  ELSE
			
			'Converte caracteres que não são válidos em uma URL e os transformamem equivalentes para URL
			strProcura = Server.URLEncode(request("busca2"))
			'Como nossa pesquisa será por "múltiplas palavras" (aqui você pode alterar ao seu gosto)
			'é necessário trocar o sinal de (=) pelo (%) que é usado com o LIKE na string SQL
			strProcura = replace(strProcura,"+"," ")
			strProcura = replace(strProcura,"%27","´")
			strProcura = replace(strProcura,"%27","'")
				
			strProcura = replace(strProcura,"%C0,","À")
			strProcura = replace(strProcura,"%C1","Á")
			strProcura = replace(strProcura,"%C2","Â")
			strProcura = replace(strProcura,"%C3","Ã")
			strProcura = replace(strProcura,"%C9","É")
			strProcura = replace(strProcura,"%CA","Ê")
			strProcura = replace(strProcura,"%CD","Í")
			strProcura = replace(strProcura,"%D3","Ó")
			strProcura = replace(strProcura,"%D4","Ô")
			strProcura = replace(strProcura,"%D5","Õ")
			strProcura = replace(strProcura,"%DA","Ú")
			strProcura = replace(strProcura,"%DC","Ü")
			
			strProcura = replace(strProcura,"%E1","à")
			strProcura = replace(strProcura,"%E1","á")
			strProcura = replace(strProcura,"%E2","â")
			strProcura = replace(strProcura,"%E3","ã")
			strProcura = replace(strProcura,"%E7","ç")
			strProcura = replace(strProcura,"%E9","é")
			strProcura = replace(strProcura,"%EA","ê")
			strProcura = replace(strProcura,"%ED","í")
			strProcura = replace(strProcura,"%F3","ó")
			strProcura = replace(strProcura,"F4","ô")
			strProcura = replace(strProcura,"F5","õ")
			strProcura = replace(strProcura,"%FA","ú")
			strProcura = replace(strProcura,"%FC","ü")
			
			
			Set RS = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT * FROM TB_Alunos where NO_Aluno like '%"& strProcura & "%' order BY NO_Aluno"
			RS.Open SQL, CON1		
			
			
			WHile Not RS.EOF
				cod = RS("CO_Matricula")
				nome = RS("NO_Aluno")
				Valor_Vetor = cod
				
				cod = RS("CO_Matricula")
				'Chama a function que ira incluir um valor para o vetor
				Call Incluir_Vetor2
			
			RS.Movenext
			Wend
				
			
			Call VisualizaValoresVetor2
		END IF
	END IF		
elseif opt="listall" then
		
	NO_Aluno = request.Form("NO_Aluno")


		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos Order BY NO_Aluno"
		RS.Open SQL, CON1
%>
		  <tr> 
			
	<td height="10" colspan="5" valign="top"> 
	  <%call mensagens(nivel,1,0,0) %>
	</td>
		  </tr>
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
			Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=ativos href=altera.asp?or=01&cod_cons="&cod&" >"&nome&"</a></font></li>")
		else
			Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=inativos href=altera.asp?or=01&cod_cons="&cod&">"&nome&"</a></font></li>")
		end if
	RS.Movenext
	Wend
end if 
%>
      </ul></td>
                </tr>
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