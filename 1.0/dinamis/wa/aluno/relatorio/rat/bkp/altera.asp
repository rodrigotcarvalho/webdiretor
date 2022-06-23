<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/connect_l.asp"-->
<!--#include file="../../../../inc/connect_g.asp"-->
<!--#include file="../../../../inc/connect_pr.asp"-->
<!--#include file="../../../../inc/connect_p.asp"-->
<!--#include file="../../../../inc/connect_a.asp"-->
<!--#include file="../../../../inc/connect_ct.asp"-->
<!--#include file="../../../../inc/connect_al.asp"-->

<%

nivel=4

opt=request.QueryString("opt")
orig=request.QueryString("ori")

autoriza=Session("autoriza")
Session("autoriza")=autoriza
if autoriza="con" or  autoriza="in" or autoriza="ex" then
check_autoriza="con"
elseif autoriza="no" then
response.redirect("../../../../novologin.asp?opt=04")
end if

permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("chave")
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

if opt="vt" or opt="ok" then
curso = request.querystring("curso")
unidade = request.querystring("unidade")
co_etapa = request.querystring("etapa")
turma = request.querystring("turma")
else
curso = request.Form("curso")
unidade = request.Form("unidade")
co_etapa = request.Form("etapa")
turma = request.Form("turma")
end if

obr=unidade&"-"&curso&"-"&co_etapa&"-"&turma


if co_etapa = "f0"then
co_etapa=0
elseif co_etapa = "f1" or co_etapa = "m1"then
co_etapa=1
elseif co_etapa = "f2" or co_etapa = "m2"then
co_etapa = 2
elseif co_etapa = "f3" or co_etapa = "m3"then
co_etapa = 3
elseif co_etapa = "f4" then
co_etapa = 4
elseif co_etapa = "f5" then
co_etapa = 5
elseif co_etapa = "f6" then
co_etapa = 6
elseif co_etapa = "f7" then
co_etapa = 7
elseif co_etapa = "f8" then
co_etapa = 8
elseif co_etapa = "f55" then
co_etapa = 55
elseif co_etapa = "f66" then
co_etapa = 66
elseif co_etapa = "f77" then
co_etapa = 77
end if

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2
		
		Set CON4 = Server.CreateObject("ADODB.Connection")
		ABRIR4 = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4
		
		Set CON5 = Server.CreateObject("ADODB.Connection") 
		ABRIR5 = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON5.Open ABRIR5
		
		Set CON6 = Server.CreateObject("ADODB.Connection") 
		ABRIR6 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON6.Open ABRIR6
 call navegacao (CON,chave,nivel)
navega=Session("caminho")

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Abreviado_Curso")


	%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../../cna/js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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
//recuperarTurma()
                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t2", true);

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

//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif"  bgcolor="#FFFFFF">
  <tr>                    
            <td height="10" class="tb_caminho"><font class="style-caminho"> 
              <%
	  response.Write(navega)

%>
              </font>
	</td>
  </tr>
  <%if check_autoriza="con" then%>
  <tr>                   
    <td height="10"> 
      <% call mensagens(4,9701,1,0)%>
    </td>
    </tr> 
<%else%>	
  <%if opt = "ok" then%>
  <tr>                   
    <td height="10"> 
      <% call mensagens(4,602,2,0)%>
    </td>
    </tr>
	<% end if %>				
    <tr>                   
    <td height="10"> 
      <%	call mensagens(4,636,0,0) %>
    </td>
                </tr>
<%end if%>												  				  


          <tr> 
            <td valign="top">
                <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">
                  <tr class="tb_tit"> 
                    
          <td height="15" class="tb_tit">Grade de Aulas 
             <input name="co_grupo" type="hidden" id="co_grupo" value="<% = co_grupo %>"> 
          </td>
                  </tr>
                  <tr> 
                    <td>
      <form name="inclusao" method="post" action="altera.asp">				
				
              <table width="1000" border="0" cellspacing="0">
                <tr> 
                  <td width="25%" class="tb_subtit"> 
                    <div align="center">UNIDADE </div></td>
                  <td width="25%" class="tb_subtit"> 
                    <div align="center">CURSO </div></td>
                  <td width="25%" class="tb_subtit"> 
                    <div align="center">ETAPA </div></td>
                  <td width="25%" class="tb_subtit"> 
                    <div align="center">TURMA </div></td>
                </tr>
                <tr> 
                  <td width="25%"> 
                    <div align="center">
                      <select name="unidade" class="select_style" onchange="recuperarCurso(this.value)">
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
                  <td width="25%"> 
                    <div align="center"> 
                      <div id="divCurso"> 
                        <select name="curso" class="select_style" onchange="recuperarEtapa(this.value)">
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
                  <td width="25%"> 
                    <div align="center"> 
                      <div id="divEtapa"> 
                        <select name="etapa" class="select_style" onchange="recuperarTurma(this.value)">
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
                  <td width="25%"> 
                    <div align="center"> 
                      <div id="divTurma"> 
                        <select name="turma" class="select_style" onChange="MM_callJS('submitfuncao()')">
                          <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Turma where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"' AND CO_Etapa='" & co_etapa & "' order by CO_Turma" 
		RS3.Open SQL3, CON0						

while not RS3.EOF
co_turma= RS3("CO_Turma")

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
RS3.MOVENEXT
WEND
%>
                        </select>
                      </div>
                    </div></td>
                </tr>
              </table>
			  </FORM>
			  </td>
                  </tr>
                  <tr> 
                    <td>
                  <table width="1000" border="0" align="right" cellspacing="0" bordercolor="#000000">
              <tr class="tb_subtit"> 
                <td width="100"> <div align="center">Matr&iacute;cula</div></td>
                <td width="350"> 
                  <div align="left">Nome do Aluno </div></td>
                <td width="50"> 
                  <div align="center">N&ordm;</div></td>
                <td width="350"> <div align="left"> Respons&aacute;vel Pedag&oacute;gico</div></td>
                <td width="150" height="40"> <div align="center">Tel de Contato</div></td>
              </tr>
              <%  check = 2
nu_chamada_check = 1

	Set RSA = Server.CreateObject("ADODB.Recordset")
	CONEXAOA = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
	Set RSA = CON4.Execute(CONEXAOA)
 
 While Not RSA.EOF
nu_matricula = RSA("CO_Matricula")
nu_chamada = RSA("NU_Chamada")
alunos=alunos+1

  		Set RSA2 = Server.CreateObject("ADODB.Recordset")
		CONEXAOA2 = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
		Set RSA2 = CON4.Execute(CONEXAOA2)
  		NO_Aluno= RSA2("NO_Aluno")
		
		Set RSA3 = Server.CreateObject("ADODB.Recordset")
		CONEXAOA3 = "Select * from TB_Alunos WHERE CO_Matricula = "& nu_matricula
		Set RSA3 = CON6.Execute(CONEXAOA3)
tp_respp= RSA3("TP_Resp_Ped")
sx= RSA3("IN_Sexo")
if sx = "F" then
sxF=sxf+1
ELSE
sxM=sxM+1
end if

		Set RSA5 = Server.CreateObject("ADODB.Recordset")
		CONEXAOA5 = "Select * from TB_Contatos WHERE CO_Matricula = "& nu_matricula&" AND TP_Contato='"&tp_respp&"'"
		Set RSA5 = CON5.Execute(CONEXAOA5)

if RSA5.EOF then
		no_respp= "RESPONSÁVEL PEDAGÓGICO NÃO CADASTRADO"
		tel_respp= ""
else
no_respp= RSA5("NO_Contato")
tel_respp= RSA5("NU_Telefones")
	if isnull(no_respp) or no_respp="" then
		no_respp= "NOME DO RESPONSÁVEL PEDAGÓGICO EM BRANCO"
	end if
		if isnull(tel_respp) or tel_respp="" then
		tel_respp= ""
	end if

end if		


 if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if
  
if nu_chamada=nu_chamada_check then
nu_chamada_check=nu_chamada_check+1%>
              <tr> 
                <td width="100" class="<%=cor%>"> <div align="center"> <font class="form_dado_texto"> 
                    <%response.Write(nu_matricula)%>
                    </font></div></td>
                <td width="350"  class="<%=cor%>"> 
                  <div align="left"> <font class="form_dado_texto"> 
                    <%response.Write(NO_Aluno)%>
                    </font></div></td>
                <td width="50" class="<%=cor%>"> 
                  <div align="center"> <font class="form_dado_texto"> 
                    <%response.Write(nu_chamada)%>
                    </font></div></td>
                <td width="350"  class="<%=cor%>"><font class="form_dado_texto"> 
                  &nbsp; 
                  <%response.Write(no_respp)%>
                  </font></td>
                <td width="150"  class="<%=cor%>"> <div align="center"><font class="form_dado_texto"> 
                    <%response.Write(tel_respp)%>
                    &nbsp;</font></div></td>
              </tr>
              <% 
else
While nu_chamada>nu_chamada_check
%>
              <tr> 
                <td width="100" bgcolor="#E4E4E4"> <div align="center"> <font class="form_dado_texto"> 
                    </font></div></td>
                <td width="350" bordercolor="#000000" bgcolor="#E4E4E4"  > 
                  <div align="left"><font class="form_dado_texto"> 
                    &nbsp;</font></div></td>
                <td width="50" bgcolor="#E4E4E4"  > 
                  <div align="center"> <font class="form_dado_texto"> 
                    </font></div></td>
                <td width="350" bordercolor="#000000" bgcolor="#E4E4E4"  ><font class="form_dado_texto"> 
                  &nbsp;</font></td>
                <td width="150" bordercolor="#000000" bgcolor="#E4E4E4"  > <div align="left"><font class="form_dado_texto"> 
                    <strong> &nbsp;</strong></font></div></td>
              </tr>
              <%
nu_chamada_check=nu_chamada_check+1	 
wend	
%>
              <tr> 
                <td width="100"  class="<%=cor%>"> <div align="center"> <font class="form_dado_texto"> 
                    <%response.Write(nu_matricula)%>
                    </font></div></td>
                <td width="350"  class="<%=cor%>"> 
                  <div align="left"> <font class="form_dado_texto"> 
                    <%response.Write(NO_Aluno)%>
                    </font></div></td>
                <td width="50"  class="<%=cor%>"> 
                  <div align="center"> <font class="form_dado_texto"> 
                    <%response.Write(nu_chamada)%>
                    </font></div></td>
                <td width="350"  class="<%=cor%>"><font class="form_dado_texto"> 
                  &nbsp;&nbsp; 
                  <%response.Write(no_respp)%>
                  </font></td>
                <td width="150"  class="<%=cor%>"> <div align="center"><font class="form_dado_texto"> 
                    <%response.Write(tel_respp)%>
                    &nbsp;&nbsp;</font></div></td>
              </tr>
              <%
 nu_chamada_check=nu_chamada_check+1	  
end if

	check = check+1
  RSA.MoveNext
  Wend 
%>
              <tr bgcolor="#FFFFFF"> 
                <td width="100">&nbsp;</td>
                <td width="350">&nbsp;</td>
                <td width="50">&nbsp;</td>
                <td width="350">&nbsp;</td>
                <td width="150">&nbsp;</td>
              </tr>
              <tr> 
                <td colspan="5"  class="tb_subtit"><table width="360" border="0" cellspacing="0">
                    <tr> 
                      <td width="160" class="tb_subtit"><font class="form_dado_texto"> 
                        Total de Alunos: 
                        <%response.Write(alunos)%>
                        </font></td>
                      <td width="100" class="tb_subtit"><font class="form_dado_texto"> 
                        Eles: 
                        <%response.Write(sxM)%>
                        </font></td>
                      <td width="100" class="tb_subtit"><font class="form_dado_texto"> 
                        Elas: 
                        <%response.Write(sxF)%>
                        </font></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
              </tr>
            </table>
					
</td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
<%call GravaLog (chave,obr)%>
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