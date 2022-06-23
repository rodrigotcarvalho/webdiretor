<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->


<%
opt = request.QueryString("opt")

ano_letivo = session("ano_letivo")
ano_atual = DatePart("yyyy", now)
co_usr = session("co_user")
nivel=4
if opt="ok" then
fl=request.QueryString("fl")
unidade = session("unidade_trabalho")
curso = session("curso_trabalho")
etapa = session("etapa_trabalho")
periodo = session("periodo_trabalho")
dia_de= session("dia_de_trabalho")
mes_de = session("mes_de_trabalho")
dia_ate = session("dia_ate_trabalho")
mes_ate = session("mes_ate_trabalho")
unidade=unidade*1
periodo=periodo*1
dia_de=dia_de*1
mes_de=mes_de*1
dia_ate=dia_ate*1
mes_ate=mes_ate*1
end if
nvg = request.QueryString("nvg")
chave=nvg
session("chave")=chave
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano_info=nivel&"-"&chave&"-"&ano_letivo

trava=session("trava")

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1

    	Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2= "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2	


 call navegacao (CON,chave,nivel)
navega=Session("caminho")

 Set RS2a = Server.CreateObject("ADODB.Recordset")
		SQL2a = "SELECT * FROM TB_Autoriz_Usuario_Grupo Where CO_Usuario = "&co_usr
		RS2.Open SQL2a, CON
		
if RS2a.EOF then

else		
co_grupo=RS2a("CO_Grupo")
End if	

		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Professor Where CO_Usuario = "&co_usr
		RS2.Open SQL2, CON2
		
'if RS2.EOF then
'Response.Write("Usuário não é Professor!")
'else		
''co_prof=RS2("CO_Professor")
'End if
%>
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
function submitfuncao()  
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
document.all.divPeriodo.innerHTML = "<select class=select_style></select>"
//recuperarEtapa()
                                                           }
                                               }

                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=e4", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divPeriodo.innerHTML = "<select class=select_style></select>"
//recuperarTurma()
                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarPeriodo()
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=p1", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_p= oHTTPRequest.responseText;
resultado_p = resultado_p.replace(/\+/g," ")
resultado_p = unescape(resultado_p)
document.all.divPeriodo.innerHTML = resultado_p
																	   
                                                           }
                                               }

                                               oHTTPRequest.send();
                                   }

 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
                        </script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<% call cabecalho (nivel)
	  %>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
                    
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
	  </td>
	  </tr>                   
<%If opt="ok" then%>
 <tr> 
    <td height="10"> 
      <%
		call mensagens(nivel,639,2,0)
%>
    </td>
                  </tr>
 <% 	end if 
%><%If trava="s" then%>
 <tr> 
    <td height="10"> 
      <%
		call mensagens(nivel,9701,0,0)
%>
    </td>
                  </tr>
 <% 	else%>
         <tr> 
                    
    <td height="10"> 
      <%	call mensagens(4,1,0,0)  
	  
%>
</td></tr>
 <% 	end if 
%>
<tr>

            <td valign="top"> <form name="alteracao" method="post" action="altera.asp">
                
        <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td width="653" height="15" class="tb_tit">Grade de Aulas</td>
          </tr>
          <tr> 
            <td>
			<table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="250" class="tb_subtit"> <div align="center">UNIDADE 
                    </div></td>
                  <td width="250" class="tb_subtit"> <div align="center">CURSO 
                    </div></td>
                  <td width="250" class="tb_subtit"> <div align="center">ETAPA 
                    </div></td>
                  <td width="250" class="tb_subtit"> <div align="center">Per&iacute;odo</div></td>
                </tr>
                <tr> 
                  <td width="250"> <div align="center"> 
                      <select name="unidade" class="select_style" onChange="recuperarCurso(this.value)">
					  <%if opt<>"ok" then%>
                        <option value="999990" selected></option>
						<%else%>
                        <option value="999990"></option>				
                        <%end if
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0
While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")

if NU_Unidade=unidade and opt="ok" then
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
RS0.MOVENEXT
WEND
%>
                      </select>
                    </div></td>
                  <td width="250"> <div align="center"> 
                      <div id="divCurso"> 
					  <%
					  if opt<>"ok" then%>
                        <select class="select_style">
                        </select>
					<%else%>
                   <select name="curso" class="select_style" onChange="recuperarEtapa(this.value)">
                        <option value="999990"></option> 
                        <%		

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT DISTINCT CO_Curso FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade
		RS0.Open SQL0, CON0	
		
While not RS0.EOF
CO_Curso = RS0("CO_Curso")

		Set RS0a = Server.CreateObject("ADODB.Recordset")
		SQL0a = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
		RS0a.Open SQL0a, CON0
		
NO_Curso = RS0a("NO_Abreviado_Curso")
if CO_Curso=curso and opt="ok" then
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
RS0.MOVENEXT
WEND
%>
                     </select> 
					
					<%end if%>	
                      </div>
                    </div></td>
                  <td width="250"> <div align="center"> 
                      <div id="divEtapa"> 
					  <%if opt<>"ok" then%>
                        <select class="select_style">
                        </select>
					<%else%>
                      <select name="etapa" class="select_style" onChange="recuperarPeriodo()">
                        <option value="999990"></option>
                        <%		
		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"'"
		RS0b.Open SQL0b, CON0
				
While not RS0b.EOF
CO_Etapa = RS0b("CO_Etapa")

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&curso&"' AND CO_Etapa='"&CO_Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")

if CO_Etapa=etapa then
%>
                        <option value="<%response.Write(CO_Etapa)%>" selected> 
                        <%response.Write(NO_Etapa)%>
                        </option>
                        <%
else		
%>
                        <option value="<%response.Write(CO_Etapa)%>"> 
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
                  <td width="250"> <div align="center"> 
                      <div id="divPeriodo"> 
					  <%if opt<>"ok" then%>
                        <select class="select_style">
                        </select>
					<%else%>
<select name="periodo" class="select_style" id="periodo">
                                      <option value="999990"></option>
                                      <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")
if NU_Periodo=periodo and opt="ok" then
%>
                                      <option value="<%=NU_Periodo%>" selected> 
                                      <%response.Write(NO_Periodo)%>
                                      </option>
                                      <%
else
%>
                                      <option value="<%=NU_Periodo%>"> 
                                      <%response.Write(NO_Periodo)%>
                                      </option>
                                      <%
end if
RS4.MOVENEXT
WEND%>
                                    </select>					
					<%end if%>
                      </div>
                    </div>
					</td>
                </tr>
                <tr> 
                  <td width="250" height="15" bgcolor="#FFFFFF"></td>
                  <td width="250" height="15" bgcolor="#FFFFFF"></td>
                  <td width="250" height="15" bgcolor="#FFFFFF"></td>
                  <td width="250" height="15" bgcolor="#FFFFFF"></td>
                </tr>
                <tr> 
                  <td height="15" colspan="4" bgcolor="#FFFFFF"><table width="1000" border="0" align="right" cellpadding="0" cellspacing="0">
                      <tr> 
                        <td class="tb_tit"
>Voc&ecirc; deseja que sejam utilizados os dados atualizados entre quais datas?</td>
                      </tr>
                      <tr> 
                        <td><div align="center"><font class="form_dado_texto">de 
                            <select name="dia_de" id="dia_de" class="select_style">
							<% if opt="ok" then
								for i= 1 to 31
									if i<10 then
										dia="0"&i
									else
										dia=i
									end if
									if i=dia_de then
									Response.Write("<option value="&i&" selected>"&dia&"</option>")
									else
									Response.Write("<option value="&i&">"&dia&"</option>")									
									end if
								next
							else%>
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
                              <option value="23">22</option>
                              <option value="24">24</option>
                              <option value="25">25</option>
                              <option value="26">26</option>
                              <option value="27">27</option>
                              <option value="28">28</option>
                              <option value="29">29</option>
                              <option value="30">30</option>
                              <option value="31">31</option>
							<%end if%>							  							  
                            </select>
                            / 
                            <select name="mes_de" id="mes_de" class="select_style">
							<% 
							if opt="ok" then
								if mes_de=1 then%>
								  <option value="1" selected>janeiro</option>
								<% else%> 
								  <option value="1" >janeiro</option>
								<%end if%>  
								<% if mes_de=2 then%>
								  <option value="2" selected>fevereiro</option>
								<% else%> 							
								  <option value="2">fevereiro</option>
								<%end if%>
								<% if mes_de=3 then%>
								  <option value="3" selected>mar&ccedil;o</option>
								<% else%> 							   
								  <option value="3">mar&ccedil;o</option>
								<%end if%>
								<% if mes_de=4 then%>
								  <option value="4" selected>abril</option>
								<% else%> 							                              
								  <option value="4">abril</option>
								<%end if%>                                
								<% if mes_de=5 then%>
								  <option value="5" selected>maio</option>
								<% else%> 
								  <option value="5">maio</option>
								<%end if%>
								<% if mes_de=6 then%>
								  <option value="6" selected>junho</option>
								<% else%> 							  
								  <option value="6">junho</option>
								<%end if%>  
								<% if mes_de=7 then%>
								  <option value="7" selected>julho</option>
								<% else%> 							
								  <option value="7">julho</option>
								<%end if%>
								<% if mes_de=8 then%>
								  <option value="8" selected>agosto</option>
								<% else%> 							  
								  <option value="8">agosto</option>
								<%end if%>
								<% if mes_de=9 then%>
								  <option value="9" selected>setembro</option>
								<% else%> 							  
								  <option value="9">setembro</option>
								<%end if%>
								<% if mes_de=10 then%>
								  <option value="10" selected>outubro</option>
								<% else%> 							  
								  <option value="10">outubro</option>
								<%end if%>
								<% if mes_de=11 then%>
								  <option value="11" selected>novembro</option>
								<% else%> 							  
								  <option value="11">novembro</option>
								<%end if%>
								<% if mes_de=12 then%>
								  <option value="12" selected>dezembro</option>
								<% else%> 							  
								  <option value="12">dezembro</option>
								<%end if
							else	
								%>  
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
							<%end if%>
                            </select>
                            / 
                            <%response.write(ano_letivo)%>
                            at&eacute; 
                            <select name="dia_ate" id="dia_ate" class="select_style">
							<% if opt="ok" then
								for i= 1 to 31
									if i<10 then
										dia="0"&i
									else
										dia=i
									end if
									if i=dia_ate then
									Response.Write("<option value="&i&" selected>"&dia&"</option>")
									else
									Response.Write("<option value="&i&">"&dia&"</option>")									
									end if
								next
							else%>
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
                              <option value="23">22</option>
                              <option value="24">24</option>
                              <option value="25">25</option>
                              <option value="26">26</option>
                              <option value="27">27</option>
                              <option value="28">28</option>
                              <option value="29">29</option>
                              <option value="30">30</option>
                              <option value="31" selected>31</option>
							<%end if%>
                            </select>
                            / 
                            <select name="mes_ate" id="mes_ate" class="select_style">
							<% 
							if opt="ok" then
								if mes_ate=1 then%>
								  <option value="1" selected>janeiro</option>
								<% else%> 
								  <option value="1" >janeiro</option>
								<%end if%>  
								<% if mes_ate=2 then%>
								  <option value="2" selected>fevereiro</option>
								<% else%> 							
								  <option value="2">fevereiro</option>
								<%end if%>
								<% if mes_ate=3 then%>
								  <option value="3" selected>mar&ccedil;o</option>
								<% else%> 							   
								  <option value="3">mar&ccedil;o</option>
								<%end if%>
								<% if mes_ate=4 then%>
								  <option value="4" selected>abril</option>
								<% else%> 							                              
								  <option value="4">abril</option>
								<%end if%>                                
								<% if mes_ate=5 then%>
								  <option value="5" selected>maio</option>
								<% else%> 
								  <option value="5">maio</option>
								<%end if%>
								<% if mes_ate=6 then%>
								  <option value="6" selected>junho</option>
								<% else%> 							  
								  <option value="6">junho</option>
								<%end if%>  
								<% if mes_ate=7 then%>
								  <option value="7" selected>julho</option>
								<% else%> 							
								  <option value="7">julho</option>
								<%end if%>
								<% if mes_ate=8 then%>
								  <option value="8" selected>agosto</option>
								<% else%> 							  
								  <option value="8">agosto</option>
								<%end if%>
								<% if mes_ate=9 then%>
								  <option value="9" selected>setembro</option>
								<% else%> 							  
								  <option value="9">setembro</option>
								<%end if%>
								<% if mes_ate=10 then%>
								  <option value="10" selected>outubro</option>
								<% else%> 							  
								  <option value="10">outubro</option>
								<%end if%>
								<% if mes_ate=11 then%>
								  <option value="11" selected>novembro</option>
								<% else%> 							  
								  <option value="11">novembro</option>
								<%end if%>
								<% if mes_ate=12 then%>
								  <option value="12" selected>dezembro</option>
								<% else%> 							  
								  <option value="12">dezembro</option>
								<%end if
							else	
								%>  
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
						<%end if%>
                            </select>
                            / 
                            <%response.write(ano_atual)%>
                            </font></div></td>
                      </tr>
                      <tr> 
                        <td><hr></td>
                      </tr>
                      <tr> 
                        <td><div align="center"> 
                            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td width="33%">&nbsp;</td>
                                <td width="34%">&nbsp;</td>
                                <td width="33%">
<div align="center"><%If trava="s" then
						else%>
                                    <input type="submit" name="Submit" value="Salvar" class="botao_prosseguir">
                     <%end if%>             </div></td>
                              </tr>
                            </table>
                          </div></td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
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
'response.redirect("../../../../inc/erro.asp")
end if
%>