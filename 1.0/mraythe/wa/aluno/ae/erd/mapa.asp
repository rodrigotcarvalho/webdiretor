<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->
<!--#include file="../../../../inc/funcoes7.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/media.asp"-->


<%opt = REQUEST.QueryString("opt")

autoriza=Session("autoriza")
Session("autoriza")=autoriza

permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("chave")
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
nivel=4

if opt= "vt" then
obr= REQUEST.QueryString("obr")
dados= split(obr, "_" )
unidade= dados(0)
curso= dados(1)
co_etapa= dados(2)
turma = dados(3)
periodo= dados(4)
mediainformada = dados(5)

else
	curso = request.Form("curso")
	unidade = request.Form("unidade")
	co_etapa = request.Form("etapa")
	turma = request.Form("turma")
	periodo = request.Form("periodo")
	mediainformada = request.Form("mediainformada")
	if isnull(curso) or curso = ""  then
		curso = session("valor4")
	end if	
	if isnull(co_etapa) or co_etapa = ""  then
		co_etapa = session("valor5")
	end if	
	if isnull(turma) or turma = "" then
		turma = session("valor6")
	end if	
	if isnull(periodo) or periodo = ""  then
		periodo = session("valor7")
	end if	
	if isnull(mediainformada) or mediainformada = "" then
		mediainformada = session("valor8")
	end if	
end if
ano_letivo = session("ano_letivo")
m_cons="VA_Media3"

obr=unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&periodo&"_"&mediainformada
'response.Write(obr)

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

 call navegacao (CON,chave,nivel)
navega=Session("caminho")
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2
		
		Set CONa = Server.CreateObject("ADODB.Connection") 
		ABRIRa = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONa.Open ABRIRa


		Set RSFIL = Server.CreateObject("ADODB.Recordset")
		SQLFIL = "SELECT * FROM TB_Da_Aula where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"'" 
		RSFIL.Open SQLFIL, CON2
if 	RSFIL.eof then
			response.Write("ERRO -  NU_Unidade ="& unidade &" AND CO_Curso ="& curso &" AND CO_Etapa ="& co_etapa &"  N�o cadastrado em TB_Da_Aula!" )
response.end()
else	
	notaFIL=RSFIL("TP_Nota")
	
	if notaFIL ="TB_NOTA_A" then
	CAMINHOn = CAMINHO_na
	
	elseif notaFIL="TB_NOTA_B" then
		CAMINHOn = CAMINHO_nb
	
	elseif notaFIL ="TB_NOTA_C" then
			CAMINHOn = CAMINHO_nc
	else
			response.Write("ERRO")
	end if
end if

		Set CONn = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONn.Open ABRIRn

		Set CON4 = Server.CreateObject("ADODB.Connection")
		ABRIR4 = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Abreviado_Curso")



		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' and CO_Curso ='"& curso &"'"  
		RS3.Open SQL3, CON0
		
if RS3.EOF THEN
no_etapa="sem etapa"
else
no_etapa=RS3("NO_Etapa")
end if

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
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
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
document.all.divPeriodo.innerHTML = "<select class=select_style></select>"
document.all.divMedia.innerHTML = "<select class=select_style></select>"
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
document.all.divPeriodo.innerHTML = "<select class=select_style></select>"
document.all.divMedia.innerHTML = "<select class=select_style></select>"
GuardaCurso(cTipo);
                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t7", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t
GuardaEtapa(eTipo);
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
						 function recuperarPeriodo(pTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=p5", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_p= oHTTPRequest.responseText;
resultado_p = resultado_p.replace(/\+/g," ")
resultado_p = unescape(resultado_p)
document.all.divPeriodo.innerHTML = resultado_p
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("p_pub=" + pTipo);
                                   }								   
						 function recuperarMedia(mTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=mi2", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_m= oHTTPRequest.responseText;
resultado_m = resultado_m.replace(/\+/g," ")
resultado_m = unescape(resultado_m)
document.all.divMedia.innerHTML = resultado_m
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("m_pub=" + mTipo);
                                   }
	 function GuardaUnidade(u)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../../../../../global/guarda_valores_digitados.asp?opt=valor3", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor3=" + u);
		}
	 function GuardaCurso(c)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../../../../../global/guarda_valores_digitados.asp?opt=valor4", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor4=" + c);
		}
function GuardaEtapa(e)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../../../../../global/guarda_valores_digitados.asp?opt=valor5", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor5=" + e);
		}			
		

	 function gravarTurma(t)
		{ 
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../../../../../global/guarda_valores_digitados.asp?opt=valor6", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor6=" + t);
		}			
	 function GuardaPeriodo(p)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../../../../../global/guarda_valores_digitados.asp?opt=valor7", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor7=" + p);
		}	
		
	 function GuardaMedia(m)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../../../../../global/guarda_valores_digitados.asp?opt=valor8", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
			   if (oHTTPRequest.readyState==4){
			 	  MM_callJS('submitfuncao()')
			   }
								   }
	
		   oHTTPRequest.send("valor8=" + m);

		}					
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
    <td height="10" colspan="7" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> </td>
  </tr>
  <tr> 
    <td height="10" colspan="7"> 
      <%	call mensagens(nivel,18,0,0) 

%>
    </td>
  </tr>
  <form name="inclusao" method="post" action="mapa.asp">
    <tr class="tb_tit"> 
      <td height="15" colspan="6" class="tb_tit"> Segmento</td>
    </tr>
    <tr> 
      <td width="166" height="10" class="tb_subtit"> <div align="center">UNIDADE 
        </div></td>
      <td width="166" height="10" class="tb_subtit"> <div align="center">CURSO 
        </div></td>
      <td width="166" height="10" class="tb_subtit"> <div align="center">ETAPA 
        </div></td>
      <td width="166" height="10" class="tb_subtit"> <div align="center">TURMA</div></td>
      <td width="167" class="tb_subtit"><div align="center">PER&Iacute;ODO</div></td>
      <td width="166" class="tb_subtit"><div align="center">M&Eacute;DIA INFORMADA</div></td>
    </tr>
    <tr>
      <td width="166"> <div align="center"> 
          <select name="unidade" class="select_style" onChange="recuperarCurso(this.value)">
            <%		
			Set RS0 = Server.CreateObject("ADODB.Recordset")
			SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
			RS0.Open SQL0, CON0
While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")
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
RS0.MOVENEXT
WEND
%>
          </select>
        </div></td>
      <td width="166"> <div align="center"> 
          <div id="divCurso"> 
            <select name="curso" class="select_style" onChange="recuperarEtapa(this.value)">
              <%		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT Distinct CO_Curso FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade
		RS0.Open SQL0, CON0
		
While not RS0.EOF
CO_Curso = RS0("CO_Curso")

		Set RS0a = Server.CreateObject("ADODB.Recordset")
		SQL0a = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
		RS0a.Open SQL0a, CON0
		
NO_Curso = RS0a("NO_Abreviado_Curso")		

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
RS0.MOVENEXT
WEND
%>
            </select>
          </div>
        </div></td>
      <td width="166"> <div align="center"> 
          <div id="divEtapa"> 
            <select name="etapa" class="select_style" onChange="recuperarTurma(this.value)">
              <%		

		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"'"
		RS0b.Open SQL0b, CON0
		
		
While not RS0b.EOF
Etapa = RS0b("CO_Etapa")


		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&curso&"' AND CO_Etapa='"&Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
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
          </div>
        </div></td>
      <td width="166" height="10"> <div id="divTurma" align="center"> 
        <select name="turma" class="select_style" onChange="MM_callJS('recuperarPeriodo()')">
          <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT DISTINCT CO_Turma FROM TB_Turma where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"' AND CO_Etapa='" & co_etapa & "' order by CO_Turma" 
		RS4.Open SQL4, CON0						

while not RS4.EOF
co_turma= RS4("CO_Turma")

if co_turma=turma then
%>
          <option value="<%response.Write(co_turma)%>" selected> 
            <%response.Write(co_turma)%>
            </option>
          <%
else
%>
          <option value="<%=co_turma%>"> 
            <%response.Write(co_turma)%>
            </option>
          <%
co_turma_check = co_turma
end if
RS4.MOVENEXT
WEND
%>
          </select>
      </div></td>
      <td width="167"><div align="center">
        <div id="divPeriodo">
<select name="periodo" class="select_style" id="periodo" onChange="MM_callJS('recuperarMedia()')">
                                      <%
		Set RSp = Server.CreateObject("ADODB.Recordset")
		SQLp = "SELECT * FROM TB_Periodo order by NU_Periodo"
		RSp.Open SQLp, CON0
periodo=periodo*1
while not RSp.EOF
NU_Periodo =  RSp("NU_Periodo")
NO_Periodo= RSp("NO_Periodo")

if periodo=NU_Periodo then
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
RSp.MOVENEXT
WEND%>
                                    </select>									

        </div>
      </div></td>
      <td width="166"><div align="center">
        <div id="divMedia">
  <select name="mediainformada" class="select_style" id="mediainformada" onChange="MM_callJS('submitfuncao()')">
  <option value="0" selected></option>
  <%opcoes=0
mediainformada=mediainformada*1
while opcoes<10.5 

if mediainformada=opcoes then
%>
    <option value="<%=opcoes%>" selected> 
      <%response.Write(opcoes)%>
      </option>
  <%
else
%>
    <option value="<%=opcoes%>"> 
      <%response.Write(opcoes)%>
      </option>
  <%
end if
opcoes=opcoes+0.5
WEND%>
    </select>	
          </div>
      </div></td>
    </tr>
  </form>
  <tr> 
    <td colspan="6" valign="top"> <table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" class="tb_corpo" >
        <tr>
          <td><hr></td>
        </tr>
        <tr> 
          <td> 
            <%
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
		RS5.Open SQL5, CON0
co_materia_check=1
while not RS5.EOF
	co_mat_fil= RS5("CO_Materia")	
	mae= RS5("IN_MAE")
	fil= RS5("IN_FIL")
	in_co= RS5("IN_CO")
	peso= RS5("NU_Peso")
		
	if (mae=FALSE and fil=FALSE and in_co=TRUE) then	
		
	else			
		if co_materia_check=1 then
			vetor_materia=co_mat_fil
		else
			vetor_materia=vetor_materia&"#!#"&co_mat_fil
		end if
		co_materia_check=co_materia_check+1	
	end if			
			
RS5.MOVENEXT
wend

vetor_materia_exibe=programa_aula(vetor_materia, unidade, curso, co_etapa, "nulo")


		Set RSt0 = Server.CreateObject("ADODB.Recordset")
		SQLt0 = "SELECT * FROM TB_Aluno_Esta_Turma where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma='"&turma&"'"
		RSt0.Open SQLt0, CONa

	co_matric_alunos_check=1
	while not RSt0.EOF
	co_matricula= RSt0("CO_Matricula")
	
		if co_matric_alunos_check=1 then
			co_matric_alunos=co_matricula
		else
			co_matric_alunos=co_matric_alunos&","&co_matricula
		end if
	co_matric_alunos_check=co_matric_alunos_check+1
	RSt0.MOVENEXT
	wend
		
		
operacoes_vetor="menor#!#maior#!#nulo"
operacoes=Split(operacoes_vetor,"#!#")
'response.Write(co_matric_alunos&"-"&notaFIL)
for o=0 to ubound(operacoes)
	if operacoes(o)="menor" then
		operacoes_nome="Menor"
	elseif operacoes(o)="maior" then
		operacoes_nome="Maior/Igual"
	elseif operacoes(o)="nulo" then
		operacoes_nome="Sem nota"
	end if		
	if o=0 then
		vetor_linha_quadro=operacoes_nome
	else
			
		vetor_linha_quadro=vetor_linha_quadro&operacoes_nome
	end if
	if operacoes(o)="nulo" then
		vetor_monta_quadro=conta_medias(unidade, curso, co_etapa, turma, periodo, co_matric_alunos, vetor_materia, CAMINHOn, notaFIL, m_cons, mediainformada, operacoes(o), outro, "nulo")
	else
		vetor_monta_quadro=conta_medias(unidade, curso, co_etapa, turma, periodo, co_matric_alunos, vetor_materia, CAMINHOn, notaFIL, m_cons, mediainformada, operacoes(o), outro, "media_turma")
	end if
vetor_linha_quadro=vetor_linha_quadro&"#!#"&vetor_monta_quadro
next
'response.Write(vetor_linha_quadro&"<BR>")
'response.end()
turmas=Split(vetor_turma,"#!#")
'Para retirar o �ltimo "#$#". Se n�o fizer isso o Ubound considera 1 elemento a mais no vetor.
linhas=Split(vetor_linha_quadro,"#$#")
For x=0 to ubound(linhas)-1
	if x=0 then
		vetor_linha_quadro=linhas(x)
	else
		vetor_linha_quadro=vetor_linha_quadro&"#$#"&linhas(x)
	end if
next
'linhas=Split(vetor_linha_quadro,"#$#")
	
co_materia_exibe=Split(vetor_materia_exibe,"#!#")

if ubound(co_materia_exibe)>15 then
	tamanho_coluna=35
else
	tamanho_coluna=50
end if
largura_tabela=(tamanho_coluna*ubound(co_materia_exibe))+50+70
%>
<table width="<%response.Write(largura_tabela)%>" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="70">&nbsp;</td>
<%For j=0 to ubound(co_materia_exibe)%>
    <th class="tb_tit" width="50"><%response.Write(co_materia_exibe(j))%></th>
<%next%>    
  </tr>
<%For k=0 to ubound(linhas)%>
  <tr>
	<%
	colunas=Split(linhas(k),"#!#")
	For m=0 to ubound(colunas)
		if m=0 then
		%>   
			<th class="tb_subtit" width="70"><%response.Write(colunas(m))%></th> 
		<%else
			if colunas(0)="Menor" then
				operacoes_nome="menor"
			elseif colunas(0)="Maior/Igual" then
				operacoes_nome="maior"
			elseif colunas(0)="Sem nota" then
				operacoes_nome="nulo"
			end if	
		
		j_corresp_m=m-1
		faixa=operacoes_nome&"_"&co_materia_exibe(j_corresp_m)
		%>    

			<td class="form_dado_texto"><div align="center"><a href="detalhar.asp?opt=grafico&fx=<%response.Write(faixa)%>&obr=<%response.Write(obr)%>&order=d"><%response.Write(colunas(m))%></a></div></th>
		<%
		end if
	next%>      
  </tr>
<%
next

session("faixas")=vetor_linha_quadro
session("faixas")=replace(session("faixas"),",",".")
session("categorias")=vetor_materia_exibe

%>  
</table>


</td>
        </tr> 
        <tr> 
         <td> 
<DIV align="center">
<iframe src ="iframe.asp" frameborder ="0" width="1000" height="400" align="middle"> </iframe>
</DIV>         
         </td>
        </tr> 
      </table></td>
  </tr>
  <tr> 
    <td height="40" colspan="7" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

</body>
<%

call GravaLog (chave,obr)
%>
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
pasta=arPath(seleciona)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>