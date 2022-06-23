<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->
<!--#include file="../../../../../global/funcoes_diversas.asp" -->
<%
opt = request.QueryString("opt")
alunos_recebidos = REQUEST.QueryString("con")
obr = request.QueryString("obr")
faixa_origem = request.QueryString("fx")
faixa_analise= split(faixa_origem, "_" )
order = request.QueryString("order")
nivel=4


autoriza=Session("autoriza")
Session("autoriza")=autoriza

permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("chave")
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo



if opt="grafico" then
	dados= split(obr, "_" )
	unidade= dados(0)
	curso= dados(1)
	co_etapa= dados(2)
	turma= dados(3)	
	periodo= dados(4)
	mediainformada= dados(5)
else
	curso = request.Form("curso")
	unidade = request.Form("unidade")
	co_etapa = request.Form("etapa")
	turma = request.Form("turma")
	periodo = request.Form("periodo")
	mediainformada = request.Form("mediainformada")
	
end if

ano_letivo = session("ano_letivo")
obr=unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&periodo&"_"&mediainformada


if periodo=1 then
	m_cons="VA_Mc1"
elseif periodo=2 then
	m_cons="VA_Mc2"
elseif periodo=3 then
	m_cons="VA_Mc3"
elseif periodo=4 then
	m_cons="VA_Mfinal"
elseif periodo=5 then
	m_cons="VA_Media3"
elseif periodo=6 then
	m_cons="VA_Media3"
end if


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CONa = Server.CreateObject("ADODB.Connection") 
		ABRIRa = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONa.Open ABRIRa		
		
		Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT


		Set RSFIL = Server.CreateObject("ADODB.Recordset")
		SQLFIL = "SELECT * FROM TB_Da_Aula where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"'" 
		RSFIL.Open SQLFIL, CON2
		if 	RSFIL.eof then
					response.Write("ERRO -  NU_Unidade ="& unidade &" AND CO_Curso ="& curso &" AND CO_Etapa ="& co_etapa &"  Não cadastrado em TB_Da_Aula!" )
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

 call navegacao (CON,chave,nivel)
navega=Session("caminho")
		
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
function checksubmit()
{
  if (document.inclusao.etapa.value == "")
  {    alert("Por favor, selecione uma etapa!")
    document.inclusao.etapa.focus()
    return false
  }
  if (document.inclusao.turma.value == "")
  {    alert("Por favor, selecione uma turma!")
    document.inclusao.turma.focus()
return false
}
  if (document.inclusao.mat_prin.value == "0")
  {    alert("Por favor, selecione uma disciplina!")
    document.inclusao.mat_prin.focus()
    return false
  }   
  if (document.inclusao.tabela.value == "")
  {    alert("Por favor, selecione uma tabela!")
    document.inclusao.tabela.focus()
    return false
  }                 	     
  return true
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
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
//recuperarTurma()
                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t4", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
						 function recuperarPeriodo(pTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=p3", true);

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

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=mi", true);

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
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif"leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" background="../../../../img/fundo_interno.gif" align="center" cellspacing="0" bgcolor="#FFFFFF">
  <tr>                    
            <td height="10" class="tb_caminho"> <font class="style-caminho">
              <%
	  response.Write(navega)

%>
              </font>
	</td>
  </tr>             <tr> 
                  
    <td height="10"> 
      <%
	call mensagens(nivel,18,0,0) 
%>
    </td>
                </tr>
                <tr> 
                  
    <td valign="top"> 
      <form name="inclusao" method="post" action="detalhar.asp?opt=detalhe&fx=<%response.Write(faixa_origem)%>" onSubmit="return checksubmit()">
                <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
                  <tr class="tb_tit"
> 
                    <td width="653" height="15" class="tb_tit"
>Segmento 
                    <input name="alunos_enviados" type="hidden" id="alunos_enviados" value="<%response.write(alunos_enviados)%>"></td>
                  </tr>
                  <tr> 
                    
            <td><table width="998" border="0" cellspacing="0">
                <tr> 
                  <td width="166" class="tb_subtit"> <div align="center">UNIDADE 
                    </div></td>
                  <td width="166" class="tb_subtit"> <div align="center">CURSO 
                    </div></td>
                  <td width="166" class="tb_subtit"> <div align="center">ETAPA 
                  </div></td>
                  <td width="166"  class="tb_subtit"><div align="center">TURMA</div></td>
                  <td width="167" class="tb_subtit"> <div align="center">PER&Iacute;ODO</div></td>
                  <td width="166" class="tb_subtit"><div align="center">M&Eacute;DIA INFORMADA</div></td>
                </tr>
                <tr>
                  <td width="166" > 
                    <div align="center"> 
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
                  <td width="166" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF"> 
                    <div align="center"> 
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
                  <td width="166" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF"> 
                    <div align="center"> 
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
                  <td width="166" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF"><div id="divTurma" align="center">
                    <select name="turma" class="select_style" onChange="recuperarPeriodo(this.value)">
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
                  <td width="167"> <div id="divPeriodo" align="center"> 
                      <select name="periodo" class="select_style" id="periodo" onChange="recuperarMedia(this.value)">
                        <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")
periodo=periodo*1
NU_Periodo=NU_Periodo*1
%>
                        <% if NU_Periodo=periodo then%>
                        <option value="<%=NU_Periodo%>" selected><%=NO_Periodo%></option>
                        <%else%>
                        <option value="<%=NU_Periodo%>"><%=NO_Periodo%></option>
                        <%end if%>
                        <%RS4.MOVENEXT
WEND%>
                      </select>
                    </div></td>
                  <td width="166" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF"><div align="center">
                    <div id="divMedia">
                      <select name="mediainformada" class="select_style" id="mediainformada" onChange="MM_callJS('submitfuncao()')">
                        <option value="0" selected></option>
                        <%opcoes=0
mediainformada=mediainformada*1
while opcoes<10.1 

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
                <tr>
                  <td colspan="6" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF"><hr></td>
                </tr>
              </table></td>
                  </tr>
                  <tr> 
                    
            <td align="center" valign="top">
            <%
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
		RS5.Open SQL5, CON0
co_materia_check=1
while not RS5.EOF
	co_mat_fil= RS5("CO_Materia")				
	if co_materia_check=1 then
		vetor_materia=co_mat_fil
	else
		vetor_materia=vetor_materia&"#!#"&co_mat_fil
	end if
	co_materia_check=co_materia_check+1			
			
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
		operacoes_nome="Abaixo"
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
	mediainformada=mediainformada*10	
	if operacoes(o)="nulo" then
		vetor_quadro=conta_medias(unidade, curso, co_etapa, turma, periodo, co_matric_alunos, vetor_materia, CAMINHOn, notaFIL, m_cons, mediainformada, operacoes(o), faixa_analise(1), "nulo")
		co_alunos=Session("aluno_nulo")
	else

		vetor_quadro=conta_medias(unidade, curso, co_etapa, turma, periodo, co_matric_alunos, vetor_materia, CAMINHOn, notaFIL, m_cons, mediainformada, operacoes(o), outro, "media_turma")
	end if
	mediainformada=mediainformada/10	
vetor_linha_quadro=vetor_linha_quadro&"#!#"&vetor_quadro
next

turmas=Split(vetor_turma,"#!#")
'Para retirar o último "#$#". Se não fizer isso o Ubound considera 1 elemento a mais no vetor.
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

largura_tabela=(50*ubound(co_materia_exibe))+50+70
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
			if colunas(0)="Abaixo" then
				operacoes_nome="menor"
			elseif colunas(0)="Maior/Igual" then
				operacoes_nome="maior"
			elseif colunas(0)="Sem nota" then
				operacoes_nome="nulo"
			end if	
			'response.Write(faixa_analise(0)&"="&operacoes_nome &"and"& faixa_analise(1)&"="&co_materia_exibe(j_corresp_m)&"<BR>")

		j_corresp_m=m-1
		faixa=operacoes_nome&"_"&co_materia_exibe(j_corresp_m)
			if faixa_analise(0)=operacoes_nome and faixa_analise(1)=co_materia_exibe(j_corresp_m) then
			%>    
	
				<td class="form_dado_texto"><div align="center"><%response.Write(colunas(m))%></div></th>
			<%
			else
			%>    
	
				<td class="form_dado_texto"><div align="center"><a href="detalhar.asp?opt=grafico&fx=<%response.Write(faixa)%>&obr=<%response.Write(obr)%>&order=d"><%response.Write(colunas(m))%></a></div></th>
			<%
			end if
		end if
	next%>      
  </tr>
<%
next

session("faixas")=vetor_linha_quadro
session("categorias")=vetor_materia_exibe

%>  
</table>
</td>
                  </tr>
                  <tr>
                    <td align="center" valign="top">&nbsp;</td>
                  </tr>
                  <tr>
                    <td align="center" valign="top">
<% 
if faixa_analise(0)="nulo" then
	'é montado no conta_medias()
	co_alunos=Session("aluno_nulo")
	ordenacao="nome_aluno"
else
	mediainformada=mediainformada*10	
	if faixa_analise(0)="menor" then
		operador=m_cons&"<"&mediainformada
	elseif faixa_analise(0)="maior" then
		operador=m_cons&">="&mediainformada
	end if

'response.Write("SELECT * FROM "&notaFIL&" where CO_Matricula IN("& co_matric_alunos &") AND CO_Materia ='"& faixa_analise(1)&"' And "&operador&"  And NU_Periodo="&periodo)
	aluno= split(co_matric_alunos,",")
	conta_aluno=0	
	for n=0 to ubound(aluno)		
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM "&notaFIL&" where CO_Matricula ="& aluno(n) &" AND CO_Materia ='"& faixa_analise(1)&"' And "&operador 
		RS3.Open SQL3, CONn
		if RS3.EOF then	
		else	
		
			if periodo=1 then
				m_cons="VA_Mc1"
			elseif periodo=2 then
				m_cons="VA_Mc2"
			elseif periodo=3 then
				m_cons="VA_Mc3"
			elseif periodo=4 then
				m_cons="VA_Mfinal"
			elseif periodo=5 then
				m_cons="VA_Media3"
			elseif periodo=6 then
				m_cons="VA_Media3"
			end if		
			media_aluno=RS3(m_cons)
			if media_aluno="" or isnull(media_aluno) then
			else
				'media_aluno=media_aluno/10
				media_aluno=formatnumber(media_aluno,0)
			end if	
			if conta_aluno=0 then
				aluno_notas=aluno(n)
				medias_encontradas=media_aluno
			else
				aluno_notas=aluno_notas&"#!#"&aluno(n)
				medias_encontradas=medias_encontradas&"#!#"&media_aluno
			end if	
			conta_aluno=conta_aluno+1				
		end if
	next				
	co_alunos=aluno_notas
	medias_ordena=split(medias_encontradas,"#!#")	
	ordenacao="media_aluno"	
end if
aluno_exibe=split(co_alunos,"#!#")

if ubound(aluno_exibe)=-1 then
 %>
 <div align="center"><font class="form_corpo"> Não existem alunos nessa faixa de notas</font></div>
<%else

%>
                  
        <table width="695" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <th scope="col" width="30" class="tb_subtit">Nº</th>   
            <th width="555" align="left" class="tb_subtit" scope="col">Nome</th>
            <th width="15" align="left" class="tb_subtit" scope="col">
           <%if order="d"then%>
           <img src="../../../../img/decres01.gif" width="12" height="25">
           <%else%>
            <a href="detalhar.asp?opt=grafico&fx=<%response.Write(faixa_origem)%>&obr=<%response.Write(obr)%>&order=d"><img src="../../../../img/decres02.gif" width="12" height="25" border="0"></a>
            <%end if%>
            </th>
            <th width="15" align="left" class="tb_subtit" scope="col">
            <%if order="d"then%>
            <a href="detalhar.asp?opt=grafico&fx=<%response.Write(faixa_origem)%>&obr=<%response.Write(obr)%>&order=c"><img src="../../../../img/cres02.gif" width="12" height="25" border="0"></a>           
			<%else%>
           <img src="../../../../img/cres01.gif" width="12" height="25"> 
            <%end if%>
            </th>
            <th scope="col" width="80" class="tb_subtit">Média-Geral</th>
          </tr>
          <%
 		Set Rsordena = Server.CreateObject ( "ADODB.RecordSet" )     
		Rsordena.Fields.Append "nu_chamada", 3
		Rsordena.Fields.Append "nome_aluno", 200, 255
		Rsordena.Fields.Append "media_aluno", 5
		Rsordena.Open
	    check = 1 
        for k=0 to ubound(aluno_exibe)  
        
			Set RSnome = Server.CreateObject("ADODB.Recordset")
			SQLnome = "SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where NU_Ano="&ano_letivo&" AND TB_Matriculas.CO_Matricula ="& aluno_exibe(k) 
			RSnome.Open SQLnome, CON4
                
			turma= RSnome("CO_Turma")
			nu_chamada = RSnome("NU_Chamada")
			nome_aluno = RSnome("NO_Aluno")
	
			Rsordena.AddNew
			Rsordena.Fields("nu_chamada").Value = nu_chamada
			Rsordena.Fields("nome_aluno").Value = nome_aluno
			if faixa_analise(0)<>"nulo" then
				media_exibe=formatnumber(medias_ordena(k),0)			
				Rsordena.Fields("media_aluno").Value = media_exibe
			end if
		next
		
		if order="d" then
			ordenar="DESC"
		else
			ordenar="ASC"	
		END IF		
		Rsordena.Sort = ordenacao&" "&ordenar
		
		Rsordena.MoveFirst
		While Not Rsordena.EoF        
            if check mod 2 =0 then
                cor = "tb_fundo_linha_par" 
            else
                cor ="tb_fundo_linha_impar"
            end if
          %> 
            <tr class="<%response.Write(cor)%>">
            <td><div align="center"><%response.Write(Rsordena.fields("nu_chamada").value)%></div></td>
            <td colspan="3"><%response.Write(Rsordena.fields("nome_aluno").value)%></td>
            <td><div align="center"><%response.Write(formatnumber(Rsordena.Fields("media_aluno").Value/10,1))%></div></td>
            </tr>
        <%
        check=check+1
		Rsordena.MoveNext
		Wend
        %> 
     </table>
<%
Rsordena.Close
Set Rsordena = Nothing
end if%>     
     </td></tr>
                  <tr>
                    <td align="center" valign="top">&nbsp;</td>
                  </tr>     
                  <tr>
                    <td align="center" valign="top"><hr></td>
                  </tr>
                  <tr>
                    <td align="center" valign="top"><table width="1000" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="33%" align="center"><input type="button" name="Voltar" id="Voltar" value="Voltar" class="botao_cancelar" onClick="MM_goToURL('parent','mapa.asp?opt=vt&obr=<%response.Write(obr)%>');return document.MM_returnValue" ></td>
                        <td width="34%" align="center">&nbsp;</td>
                        <td width="33%" align="center">&nbsp;</td>
                      </tr>
                    </table></td>
                  </tr>
        </table>
              </form></td>
  </tr>
  <tr>
    <td height="40"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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