<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../../global/tabelas_escolas.asp"-->
<%opt = REQUEST.QueryString("opt")
volta= REQUEST.QueryString("volta")

cod=session("codigo")
escola=session("escola")

if opt="direto" then
ano_letivo = request.querystring("ano")
curso = request.querystring("curso")
unidade = request.querystring("unidade")
grade = request.querystring("grade")
turma = "sem turma"
elseif volta="1" then
ano_letivo = request.querystring("ano")
curso = request.querystring("curso")
unidade = request.querystring("unidade")
co_etapa = request.querystring("etapa")
turma = request.querystring("turma")
periodo = request.querystring("periodo")
else
ano_letivo = session("ano_letivo")
curso = request.Form("curso")
unidade = request.Form("unidade")
co_etapa = request.Form("etapa")
turma = request.Form("turma")
periodo = request.Form("periodo")
avaliacao_form = request.Form("avaliacoes")
end if
nivel=4

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
		ABRIR4 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4
		
		Set CON_wr = Server.CreateObject("ADODB.Connection") 
		ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_wr.Open ABRIR_wr		
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")

		Set RSTB = Server.CreateObject("ADODB.Recordset")
		CONEXAOTB = "Select * from TB_Da_Aula WHERE NU_Unidade="& unidade &" AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		Set RSTB = CON2.Execute(CONEXAOTB)
		
nota= RSTB("TP_Nota")

		
if nota = "TB_NOTA_A" Then		
		CAMINHOn = CAMINHO_na
		opcao="A"
elseif nota = "TB_NOTA_B" Then
		CAMINHOn = CAMINHO_nb
		opcao="B"		
elseif nota = "TB_NOTA_C" Then
		CAMINHOn = CAMINHO_nc
		opcao="C"		
elseif nota ="TB_NOTA_D" then
		CAMINHOn = CAMINHO_nd
		opcao="D"		
elseif nota ="TB_NOTA_E" then
		CAMINHOn = CAMINHO_ne
		opcao="E"		
elseif nota ="TB_NOTA_F" then
		CAMINHOn = CAMINHO_nf
		opcao="F"		
elseif nota ="TB_NOTA_K" then
		CAMINHOn = CAMINHO_nk
		opcao="K"	
elseif nota ="TB_NOTA_L" then
		CAMINHOn = CAMINHO_nl
		opcao="L"		
elseif nota ="TB_NOTA_M" then
		CAMINHOn = CAMINHO_nm
		opcao="M"													
end if
obr=unidade&"-"&curso&"-"&co_etapa&"-"&turma&"-"&periodo&"-"&avaliacao_form&"-"&nota

		Set CON3 = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3

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

<script type="text/javascript" src="../js/global.js"></script>
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
}  function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
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
document.all.divAvaliacao.innerHTML = "<select class=select_style></select>"
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
document.all.divAvaliacao.innerHTML = "<select class=select_style></select>"
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
document.all.divPeriodo.innerHTML = "<select class=select_style></select>"
document.all.divAvaliacao.innerHTML = "<select class=select_style></select>"
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
							function recuperarPeriodo(pTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=p2", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_p= oHTTPRequest.responseText;
resultado_p = resultado_p.replace(/\+/g," ")
resultado_p = unescape(resultado_p)
document.all.divPeriodo.innerHTML = resultado_p
document.all.divAvaliacao.innerHTML = "<select class=select_style></select>"
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("p_pub=" + pTipo);
                                   }								   
						 function recuperarAvaliacoes(avTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=av2", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_a= oHTTPRequest.responseText;
resultado_a = resultado_a.replace(/\+/g," ")
resultado_a = unescape(resultado_a)
document.all.divAvaliacao.innerHTML = resultado_a
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("av_pub=" + avTipo);
                                   }
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
 
</head> 
<body link="#CC9900" vlink="#CC9900" background="../../../../img/fundo.gif" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr> 
            
    <td height="10" class="tb_caminho">
<font class="style-caminho"> 
                <%
	  response.Write(navega)%>
	  </font>
</td>
          </tr>
        <tr> 
          
    <td height="10" valign="top"> 
      <div align="left"> <strong> 
        <%	call mensagens(nivel,636,0,0) 

%>
        </strong></div></td>
                </tr>
<tr>             
    <td height="10" valign="top"> 
      <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo" dwcopytype="CopyTableCell"
>


                  <tr class="tb_tit"
> 
                    
            <td height="15" class="tb_tit"
>Grade de Aulas </td>
                  </tr>          
                </table>
    </td>
                  </tr> 
<tr>             
    <td height="41" valign="top"> 
      <form name="inclusao" method="post" action="mapa.asp?or=02" onSubmit="return checksubmit()">
			
        <table width="100%" border="0" cellspacing="0">
          <tr> 
            <td width="166" class="tb_subtit"> <div align="center">UNIDADE </div></td>
            <td width="166" class="tb_subtit"> <div align="center">CURSO </div></td>
            <td width="166" class="tb_subtit"> <div align="center">ETAPA </div></td>
            <td width="166" class="tb_subtit"> <div align="center">TURMA</div></td>
            <td width="167" class="tb_subtit"> <div align="center">PER&Iacute;ODO</div></td>
            <td width="166" class="tb_subtit"> <div align="center">AVALIA&Ccedil;&Atilde;O</div></td>
          </tr>
          <tr>
            <td> <div align="center"> 
                <select name="unidade" class="select_style" onChange="recuperarCurso(this.value)">
                  <option value="999990"></option>
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
            <td> <div id="divCurso" align="center"> 
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
              </div></td>
            <td> <div id="divEtapa" align="center"> 
                  <select name="etapa" class="select_style" onChange="recuperarTurma(this.value)">
                    <option value="999990"></option>
                    <%		
		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"'"
		RS0b.Open SQL0b, CON0
				
While not RS0b.EOF
COD_Etapa = RS0b("CO_Etapa")

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&curso&"' AND CO_Etapa='"&COD_Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")

if COD_Etapa=co_etapa then
%>
                    <option value="<%response.Write(COD_Etapa)%>" selected> 
                    <%response.Write(NO_Etapa)%>
                    </option>
                    <%
else		
%>
                    <option value="<%response.Write(COD_Etapa)%>"> 
                    <%response.Write(NO_Etapa)%>
                    </option>
                    <%
end if						
RS0b.MOVENEXT
WEND
%>
                  </select>
              </div></td>
            <td width="166"> <div id="divTurma" align="center"><font class="form_dado_texto"> 
<select name="turma" class="select_style" onChange="MM_callJS('recuperarPeriodo()')">
                        <option value="999990" selected></option>
    <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Turma where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"' AND CO_Etapa='" & session("e_pub") & "' order by CO_Turma" 
		RS3.Open SQL3, CON0						

while not RS3.EOF
co_turma= RS3("CO_Turma")
if co_turma=turma then
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
RS3.MOVENEXT
WEND
%>
    
 </select>
                </font></div></td>
            <td width="167"> <div id="divPeriodo" align="center"> 
                <select name="periodo" class="select_style" id="select" onChange="MM_callJS('recuperarAvaliacoes()')">
                  <option value="999990"></option>
                  <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")
periodo=periodo*1
NU_Periodo=NU_Periodo*1
if NU_Periodo=periodo then
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
              </div></td>
            <td width="166"> <div id="divAvaliacao" align="center"> 
                <select name="avaliacoes" class="select_style" id="avaliacoes" onChange="MM_callJS('submitfuncao()')">
                  <option value="999990"></option>
                  <%

	dados_tabela=verifica_dados_tabela(CAMINHOn,opcao,outro)
	dados_separados=split(dados_tabela,"#$#")
	ln_nom_cols=dados_separados(4)
	nm_vars=dados_separados(5)
	nm_bd=dados_separados(6)
	avaliacoes_nomes=split(ln_nom_cols,"#!#")
	verifica_avaliacoes=split(nm_vars,"#!#")
	avaliacoes=split(nm_bd,"#!#")
	
for i=2 to UBOUND(avaliacoes_nomes)
	j=i-2
	'if avaliacoes(j)="CALCULADO" or verifica_avaliacoes(j)="media_teste" or verifica_avaliacoes(j)="media_prova" or verifica_avaliacoes(j)="media1" or verifica_avaliacoes(j)="media2" or verifica_avaliacoes(j)="media3" then
	if avaliacoes(j)="CALCULADO" or verifica_avaliacoes(j)="media_teste" or verifica_avaliacoes(j)="rs" or verifica_avaliacoes(j)="rb"  then	
	elseif avaliacoes(j)=avaliacao_form then
%>
                  <option value="<%response.Write(avaliacoes(j))%>" selected> 
                  <%response.Write(avaliacoes_nomes(i))%>
                  </option>
                  <%
	else
	%>
                  <option value="<%response.Write(avaliacoes(j))%>"> 
                  <%response.Write(avaliacoes_nomes(i))%>
                  </option>
					  <%
	end if	
NEXT
%>
                </select>   
</div></td>
          </tr>
        </table>
              </form></td>
          </tr>
          <tr> 
            
    <td valign="top" bordercolor="#FFFFCC" bgcolor="#FFFFFF"> 
      <%
campo_check =avaliacao_form


		Set RSNN = Server.CreateObject("ADODB.Recordset")
		CONEXAONN = "Select CO_Materia  from TB_Programa_Aula WHERE CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa&"' order by NU_Ordem_Boletim"

		Set RSNN = CON0.Execute(CONEXAONN)


materia_nome_check="vazio"
nome_nota="vazio"
i=0
largura = 0
While not RSNN.eof
'i=i+1
materia_nome= RSNN("CO_Materia")
'			response.Write(materia_nome&"<br>")
'Response.Flush()

if materia_nome=materia_nome_check then
RSNN.movenext
else
If Not IsArray(nome_nota) Then 
nome_nota = Array()
End if
'If InStr(Join(nome_nota), materia_nome) = 0 Then
ReDim preserve nome_nota(UBound(nome_nota)+1)
nome_nota(Ubound(nome_nota)) = materia_nome
largura=largura+35
i=i+1
materia_nome_check=materia_nome


end if
'end if
RSNN.movenext
wend
larg=1008-(largura/i)
		

		
%>
      <table width="1000" border="0" align="right" cellspacing="0">
        <tr> 
          <td width="17" class="tb_subtit"> <div align="right"><strong>N&ordm;</strong></div></td>
          <td width="larg"  class="tb_subtit"> <div align="center"><strong>Nome</strong></div></td>
          <%For k=0 To ubound(nome_nota)%>
          <td width="40" class="tb_subtit"> <div align="center"><strong> 
              <% response.Write(nome_nota(k))%>
              </strong></div></td>
          <%
Next%>
        </tr>
        <%  check = 2
nu_chamada_check = 1

	Set RSA = Server.CreateObject("ADODB.Recordset")
	CONEXAOA = "Select * from TB_Matriculas where NU_Ano="& ano_letivo &" AND NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
	Set RSA = CON4.Execute(CONEXAOA)
	

 
 While Not RSA.EOF
nu_matricula = RSA("CO_Matricula")
nu_chamada = RSA("NU_Chamada")

  		Set RSA2 = Server.CreateObject("ADODB.Recordset")
		CONEXAOA2 = "Select * from TB_Alunos WHERE CO_Matricula = "& nu_matricula
		Set RSA2 = CON4.Execute(CONEXAOA2)
  		NO_Aluno= RSA2("NO_Aluno")

 if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if
  
if nu_chamada=nu_chamada_check then
nu_chamada_check=nu_chamada_check+1%>
        <tr class="<%=cor%>"> 
          <td width="17"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
              <%response.Write(nu_chamada)%>
              </strong></font></div></td>
          <td width="200"> <div align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
              <%response.Write(NO_Aluno)%>
              </strong></font></div></td>
          <%For k=0 To ubound(nome_nota)
 		Set RS3 = Server.CreateObject("ADODB.Recordset")
		CONEXAO3 = "Select "& campo_check & " from "& nota & " WHERE NU_Periodo = "& periodo &" And CO_Materia = '"& nome_nota(k) &"' And CO_Matricula = "& nu_matricula
		Set RS3 = CON3.Execute(CONEXAO3)
  		
if RS3.EOF Then
%>
          <td> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1"></font></font></div></td>
          <%else
nota_materia= RS3(""&campo_check&"")
if isnumeric(nota_materia) then
'nota_materia=nota_materia/10
nota_materia = formatNumber(nota_materia,0)  
end if   
%>
          <td width="40"> 
            <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%response.Write(nota_materia)%>
              </font></div></td>
          <%end IF

NEXT%>
        </tr>
        <% 
else
While nu_chamada>nu_chamada_check
%>
        <tr bgcolor="#E4E4E4"> 
          <td width="17" > <div align="center"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%response.Write(nu_chamada_check)%>
              </font></strong></div></td>
          <td width="200"> <div align="left"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              </font></strong></div></td>
          <%For k=0 To ubound(nome_nota)%>
          <td> <div align="center"></div></td>
          <%

NEXT
%>
        </tr>
        <%
nu_chamada_check=nu_chamada_check+1	 
wend	
%>
        <tr class="<%=cor%>"> 
          <td width="17"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
              <%response.Write(nu_chamada)%>
              </strong></font></div></td>
          <td width="200"> <div align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
              <%response.Write(NO_Aluno)%>
              </strong></font></div></td>
          <%For k=0 To ubound(nome_nota)

  		Set RS3 = Server.CreateObject("ADODB.Recordset")
		CONEXAO3 = "Select "& campo_check & " from "& nota & " WHERE NU_Periodo = "& periodo &" And CO_Materia = '"& nome_nota(k) &"' And CO_Matricula = "& nu_matricula
		Set RS3 = CON3.Execute(CONEXAO3)
		
if RS3.EOF Then
%>
          <td> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1"></font></font></div></td>
          <%else
nota_materia= RS3(""& campo_check & "")

if isnumeric(nota_materia) then
'nota_materia=nota_materia/10
nota_materia = formatNumber(nota_materia,0)  
end if    
%>
          <td width="40"> 
            <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1"> 
              <%response.Write(nota_materia)%>
              </font></font></div></td>
          <%end IF

NEXT%>
        </tr>
        <%
 nu_chamada_check=nu_chamada_check+1	  
end if

	check = check+1
  RSA.MoveNext
  Wend 
%>
      </table>
		
    </td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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
pasta=arPath(seleciona1)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>