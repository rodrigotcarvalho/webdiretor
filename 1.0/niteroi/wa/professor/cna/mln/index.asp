<%On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->

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
Session("data_consulta")=""
Session("hora_consulta")=""


		
		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2	

 call VerificaAcesso (CON,chave,nivel)
autoriza=Session("autoriza")

 call navegacao (CON,chave,nivel)
navega=Session("caminho")


	%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
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
function submitfuncao()  
{
   var f=document.forms[0]; 
      f.submit(); 
	  
}  function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif"leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<form name="form1" method="post" action="monitora.asp?opt=1&nvg=<%=nvg%>" onSubmit="return checksubmit()">
<%call cabecalho(nivel)%>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
                  <tr>                    
            
    <td height="10" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> </td>
  </tr>
  <tr>                   
    <td height="10"> 
      <%
	  if autoriza="no" then
	  	call mensagens(4,9700,1,0) 	  
	  else
	  	call mensagens(4,617,0,0) 
	  end if%>
    </td>
                  </tr>				  
		  				  				  

  <tr> 
    <td valign="top">

<table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">

          <tr> 
            <td> 
              <%	  if autoriza="no" then			
		else
ano_slct = DatePart("yyyy", now)
mes_slct = DatePart("m", now) 
dia_slct = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

hora = hora*1
min = min*1		
%>
        
      <table width="1000" border="0" cellspacing="0">
        <tr> 
                <td valign="top"> 
                  <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo"
>
                    <tr> 
                      <td class="tb_tit">Período de Monitoramento de Notas</td>
                </tr>
                <tr> 
                  <td><table width="100%" border="0" align="right" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td class="tb_subtit"> <div align="center">DATA</div></td>
                              <td class="tb_subtit"> <div align="center">HORA</div></td>
                            </tr>
                            <tr> 
                              <td width="43%"><div align="center"> <font class="form_dado_texto"> 
                              <input name="dia_mnl" type="hidden" id="dia_mnl" value="<%=dia_slct%>">
                              <input name="mes_mnl" type="hidden" id="mes_mnl" value="<%=mes_slct%>">
                              <input name="ano_mnl" type="hidden" id="ano_mnl" value="<%=ano_slct%>">
                              <%response.Write(dia_slct)%>
                              / 
                              <%
					  if mes_slct< 10 then
					  mes_slct="0"&mes_slct
					  end if 
					  
					  response.Write(mes_slct)%>
                              / 
                              <%response.Write(ano_slct)%></font>

                                  </div></td>
                              <td width="57%"><div align="center"> 
                                  <select name="hora_mnl" class="select_style">
                                    <% if hora =0 or hora =24  then%>
                                    <option value="0" selected>0</option>
                                    <% else%>
                                    <option value="0">00</option>
                                    <% end if %>
                                    <% if hora =1 then%>
                                    <option value="1" selected>01</option>
                                    <% else%>
                                    <option value="1">01</option>
                                    <% end if %>
                                    <% if hora =2 then%>
                                    <option value="2" selected>02</option>
                                    <% else%>
                                    <option value="2">02</option>
                                    <% end if %>
                                    <% if hora =3 then%>
                                    <option value="3" selected>03</option>
                                    <% else%>
                                    <option value="3">03</option>
                                    <% end if %>
                                    <% if hora =4 then%>
                                    <option value="4" selected>04</option>
                                    <% else%>
                                    <option value="4">04</option>
                                    <% end if %>
                                    <% if hora =5 then%>
                                    <option value="5" selected>05</option>
                                    <% else%>
                                    <option value="5">05</option>
                                    <% end if %>
                                    <% if hora =6 then%>
                                    <option value="6" selected>06</option>
                                    <% else%>
                                    <option value="6">06</option>
                                    <% end if %>
                                    <% if hora =7 then%>
                                    <option value="7" selected>07</option>
                                    <% else%>
                                    <option value="7">07</option>
                                    <% end if %>
                                    <% if hora =8 then%>
                                    <option value="8" selected>08</option>
                                    <% else%>
                                    <option value="8">08</option>
                                    <% end if %>
                                    <% if hora =9 then%>
                                    <option value="9" selected>09</option>
                                    <% else%>
                                    <option value="9">09</option>
                                    <% end if %>
                                    <% if hora =10 then%>
                                    <option value="10" selected>10</option>
                                    <% else%>
                                    <option value="10">10</option>
                                    <% end if %>
                                    <% if hora =11 then%>
                                    <option value="11" selected>11</option>
                                    <% else%>
                                    <option value="11">11</option>
                                    <% end if %>
                                    <% if hora =12 then%>
                                    <option value="12" selected>12</option>
                                    <% else%>
                                    <option value="12">12</option>
                                    <% end if %>
                                    <% if hora =13 then%>
                                    <option value="13" selected>13</option>
                                    <% else%>
                                    <option value="13">13</option>
                                    <% end if %>
                                    <% if hora =14 then%>
                                    <option value="14" selected>14</option>
                                    <% else%>
                                    <option value="14">14</option>
                                    <% end if %>
                                    <% if hora =15 then%>
                                    <option value="15" selected>15</option>
                                    <% else%>
                                    <option value="15">15</option>
                                    <% end if %>
                                    <% if hora =16 then%>
                                    <option value="16" selected>16</option>
                                    <% else%>
                                    <option value="16">16</option>
                                    <% end if %>
                                    <% if hora =17 then%>
                                    <option value="17" selected>17</option>
                                    <% else%>
                                    <option value="17">17</option>
                                    <% end if %>
                                    <% if hora =18 then%>
                                    <option value="18" selected>18</option>
                                    <% else%>
                                    <option value="18">18</option>
                                    <% end if %>
                                    <% if hora =19 then%>
                                    <option value="19" selected>19</option>
                                    <% else%>
                                    <option value="19">19</option>
                                    <% end if %>
                                    <% if hora =20 then%>
                                    <option value="20" selected>20</option>
                                    <% else%>
                                    <option value="20">20</option>
                                    <% end if %>
                                    <% if hora =21 then%>
                                    <option value="21" selected>21</option>
                                    <% else%>
                                    <option value="21">21</option>
                                    <% end if %>
                                    <% if hora =22 then%>
                                    <option value="22" selected>22</option>
                                    <% else%>
                                    <option value="22">22</option>
                                    <% end if %>
                                    <% if hora =23 then%>
                                    <option value="23" selected>23</option>
                                    <% else%>
                                    <option value="23">23</option>
                                    <% end if %>
                                  </select>
                                  : 
                                  <select name="min_mnl" class="select_style">
                                    <% if min =00 then%>
                                    <option value="00" selected>00</option>
                                    <% else%>
                                    <option value="00">00</option>
                                    <% end if %>
                                    <%if min =1 then%>
                                    <option value="1" selected>01</option>
                                    <% else%>
                                    <option value="1">01</option>
                                    <% end if %>
                                    <% if min =2 then%>
                                    <option value="2" selected>02</option>
                                    <% else%>
                                    <option value="2">02</option>
                                    <% end if %>
                                    <% if min =3 then%>
                                    <option value="3" selected>03</option>
                                    <% else%>
                                    <option value="3">03</option>
                                    <% end if %>
                                    <% if min =4 then%>
                                    <option value="4" selected>04</option>
                                    <% else%>
                                    <option value="4">04</option>
                                    <% end if %>
                                    <% if min =5 then%>
                                    <option value="5" selected>05</option>
                                    <% else%>
                                    <option value="5">05</option>
                                    <% end if %>
                                    <% if min =6 then%>
                                    <option value="6" selected>06</option>
                                    <% else%>
                                    <option value="6">06</option>
                                    <% end if %>
                                    <% if min =7 then%>
                                    <option value="7" selected>07</option>
                                    <% else%>
                                    <option value="7">07</option>
                                    <% end if %>
                                    <% if min =8 then%>
                                    <option value="8" selected>08</option>
                                    <% else%>
                                    <option value="8">08</option>
                                    <% end if %>
                                    <% if min =9 then%>
                                    <option value="9" selected>09</option>
                                    <% else%>
                                    <option value="9">09</option>
                                    <% end if %>
                                    <% if min =10 then%>
                                    <option value="10" selected>10</option>
                                    <% else%>
                                    <option value="10">10</option>
                                    <% end if %>
                                    <% if min =11 then%>
                                    <option value="11" selected>11</option>
                                    <% else%>
                                    <option value="11">11</option>
                                    <% end if %>
                                    <% if min =12 then%>
                                    <option value="12" selected>12</option>
                                    <% else%>
                                    <option value="12">12</option>
                                    <% end if %>
                                    <% if min =13 then%>
                                    <option value="13" selected>13</option>
                                    <% else%>
                                    <option value="13">13</option>
                                    <% end if %>
                                    <% if min =14 then%>
                                    <option value="14" selected>14</option>
                                    <% else%>
                                    <option value="14">14</option>
                                    <% end if %>
                                    <% if min =15 then%>
                                    <option value="15" selected>15</option>
                                    <% else%>
                                    <option value="15">15</option>
                                    <% end if %>
                                    <% if min =16 then%>
                                    <option value="16" selected>16</option>
                                    <% else%>
                                    <option value="16">16</option>
                                    <% end if %>
                                    <% if min =17 then%>
                                    <option value="17" selected>17</option>
                                    <% else%>
                                    <option value="17">17</option>
                                    <% end if %>
                                    <% if min =18 then%>
                                    <option value="18" selected>18</option>
                                    <% else%>
                                    <option value="18">18</option>
                                    <% end if %>
                                    <% if min =19 then%>
                                    <option value="19" selected>19</option>
                                    <% else%>
                                    <option value="19">19</option>
                                    <% end if %>
                                    <% if min =20 then%>
                                    <option value="20" selected>20</option>
                                    <% else%>
                                    <option value="20">20</option>
                                    <% end if %>
                                    <% if min =21 then%>
                                    <option value="21" selected>21</option>
                                    <% else%>
                                    <option value="21">21</option>
                                    <% end if %>
                                    <% if min =22 then%>
                                    <option value="22" selected>22</option>
                                    <% else%>
                                    <option value="22">22</option>
                                    <% end if %>
                                    <% if min =23 then%>
                                    <option value="23" selected>23</option>
                                    <% else%>
                                    <option value="23">23</option>
                                    <% end if %>
                                    <% if min =24 then%>
                                    <option value="24" selected>24</option>
                                    <% else%>
                                    <option value="24">24</option>
                                    <% end if %>
                                    <% if min =25 then%>
                                    <option value="25" selected>25</option>
                                    <% else%>
                                    <option value="25">25</option>
                                    <% end if %>
                                    <% if min =26 then%>
                                    <option value="26" selected>26</option>
                                    <% else%>
                                    <option value="26">26</option>
                                    <% end if %>
                                    <% if min =27 then%>
                                    <option value="27" selected>27</option>
                                    <% else%>
                                    <option value="27">27</option>
                                    <% end if %>
                                    <% if min =28 then%>
                                    <option value="28" selected>28</option>
                                    <% else%>
                                    <option value="28">28</option>
                                    <% end if %>
                                    <% if min =29 then%>
                                    <option value="29" selected>29</option>
                                    <% else%>
                                    <option value="29">29</option>
                                    <% end if %>
                                    <% if min =30 then%>
                                    <option value="30" selected>30</option>
                                    <% else%>
                                    <option value="30">30</option>
                                    <% end if %>
                                    <% if min =31 then%>
                                    <option value="31" selected>31</option>
                                    <% else%>
                                    <option value="31">31</option>
                                    <% end if %>
                                    <% if min =32 then%>
                                    <option value="32" selected>32</option>
                                    <% else%>
                                    <option value="32">32</option>
                                    <% end if %>
                                    <% if min =33 then%>
                                    <option value="33" selected>33</option>
                                    <% else%>
                                    <option value="33">33</option>
                                    <% end if %>
                                    <% if min =34 then%>
                                    <option value="34" selected>34</option>
                                    <% else%>
                                    <option value="34">34</option>
                                    <% end if %>
                                    <% if min =35 then%>
                                    <option value="35" selected>35</option>
                                    <% else%>
                                    <option value="35">35</option>
                                    <% end if %>
                                    <% if min =36 then%>
                                    <option value="36" selected>36</option>
                                    <% else%>
                                    <option value="36">36</option>
                                    <% end if %>
                                    <% if min =37 then%>
                                    <option value="37" selected>37</option>
                                    <% else%>
                                    <option value="37">37</option>
                                    <% end if %>
                                    <% if min =38 then%>
                                    <option value="38" selected>38</option>
                                    <% else%>
                                    <option value="38">38</option>
                                    <% end if %>
                                    <% if min =39 then%>
                                    <option value="39" selected>39</option>
                                    <% else%>
                                    <option value="39">39</option>
                                    <% end if %>
                                    <% if min =40 then%>
                                    <option value="40" selected>40</option>
                                    <% else%>
                                    <option value="40">40</option>
                                    <% end if %>
                                    <% if min =41 then%>
                                    <option value="41" selected>41</option>
                                    <% else%>
                                    <option value="41">41</option>
                                    <% end if %>
                                    <% if min =42 then%>
                                    <option value="42" selected>42</option>
                                    <% else%>
                                    <option value="42">42</option>
                                    <% end if %>
                                    <% if min =43 then%>
                                    <option value="43" selected>43</option>
                                    <% else%>
                                    <option value="43">43</option>
                                    <% end if %>
                                    <% if min =44 then%>
                                    <option value="44" selected>44</option>
                                    <% else%>
                                    <option value="44">44</option>
                                    <% end if %>
                                    <% if min =45 then%>
                                    <option value="45" selected>45</option>
                                    <% else%>
                                    <option value="45">45</option>
                                    <% end if %>
                                    <% if min =46 then%>
                                    <option value="46" selected>46</option>
                                    <% else%>
                                    <option value="46">46</option>
                                    <% end if %>
                                    <% if min =47 then%>
                                    <option value="47" selected>47</option>
                                    <% else%>
                                    <option value="47">47</option>
                                    <% end if %>
                                    <% if min =48 then%>
                                    <option value="48" selected>48</option>
                                    <% else%>
                                    <option value="48">48</option>
                                    <% end if %>
                                    <% if min =49 then%>
                                    <option value="49" selected>49</option>
                                    <% else%>
                                    <option value="49">49</option>
                                    <% end if %>
                                    <% if min =50 then%>
                                    <option value="50" selected>50</option>
                                    <% else%>
                                    <option value="50">50</option>
                                    <% end if %>
                                    <% if min =51 then%>
                                    <option value="51" selected>51</option>
                                    <% else%>
                                    <option value="51">51</option>
                                    <% end if %>
                                    <% if min =52 then%>
                                    <option value="52" selected>52</option>
                                    <% else%>
                                    <option value="52">52</option>
                                    <% end if %>
                                    <% if min =53 then%>
                                    <option value="53" selected>53</option>
                                    <% else%>
                                    <option value="53">53</option>
                                    <% end if %>
                                    <% if min =54 then%>
                                    <option value="54" selected>54</option>
                                    <% else%>
                                    <option value="54">54</option>
                                    <% end if %>
                                    <% if min =55 then%>
                                    <option value="55" selected>55</option>
                                    <% else%>
                                    <option value="55">55</option>
                                    <% end if %>
                                    <% if min =56 then%>
                                    <option value="56" selected>56</option>
                                    <% else%>
                                    <option value="56">56</option>
                                    <% end if %>
                                    <% if min =57 then%>
                                    <option value="57" selected>57</option>
                                    <% else%>
                                    <option value="57">57</option>
                                    <% end if %>
                                    <% if min =58 then%>
                                    <option value="58" selected>58</option>
                                    <% else%>
                                    <option value="58">58</option>
                                    <% end if %>
                                    <% if min =59 then%>
                                    <option value="59" selected>59</option>
                                    <% else%>
                                    <option value="59">59</option>
                                    <% end if %>
                                  </select>
                                </div></td>
                            </tr>
                            <tr> 
                              <td>&nbsp;</td>
                              <td>&nbsp;</td>
                            </tr>
                            <tr> 
                              <td><div align="center"> 
                                  <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','../../../index.asp');return document.MM_returnValue" value="Cancelar">
                                </div></td>
                              <td><div align="center"> 
                                  <input name="Submit2" type="submit" class="botao_prosseguir" value="Iniciar">
                                </div></td>
                            </tr>
                          </table>
                    </td>
                </tr>
              </table>
		        </td>
        </tr>
      </table>
      </div> 
      <%end if 
%>
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