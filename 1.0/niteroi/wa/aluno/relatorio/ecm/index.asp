<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/parametros.asp"-->


<%
opt = request.QueryString("opt")

ano_letivo = session("ano_letivo")
co_usr = session("co_user")
nivel=4
if opt="cln" or opt="gera" then
	nvg=session("chave")
	unidade=request.form("unidade")
	curso=request.form("curso")
	co_etapa=request.form("etapa")
	turma=request.form("turma")	
	periodo=request.form("media")
	tp_ord=request.form("tp_ord")
	
	response.Redirect("../../../../relatorios/swd070.asp?inf="&unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma&"$!$"&periodo&"$!$"&tp_ord)

else
	nvg = request.QueryString("nvg")
	chave=nvg
	session("chave")=chave
end if
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

    	Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2= "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2	


 call navegacao (CON,chave,nivel)
navega=Session("caminho")	

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
MM_reloadPage(true);function submitfuncao()  
{
   var f=document.forms[0]; f.submit(); 
	  
}
function checksubmit()
{
  if (document.alteracao.media.value == "nulo")
  {    alert("Por favor selecione uma média para gerar o relatório!")
    document.alteracao.media.focus()
    return false
  }
  return true
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
//recuperarEtapa()
                                                           }
                                               }

                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=e10", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
//recuperarTurma()
                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }
						   								   

 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
                        </script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../../img/menu_r1_c2_f3.gif','../../../../img/menu_r1_c2_f2.gif','../../../../img/menu_r1_c2_f4.gif','../../../../img/menu_r1_c4_f3.gif','../../../../img/menu_r1_c4_f2.gif','../../../../img/menu_r1_c4_f4.gif','../../../../img/menu_r1_c6_f3.gif','../../../../img/menu_r1_c6_f2.gif','../../../../img/menu_r1_c6_f4.gif','../../../../img/menu_r1_c8_f3.gif','../../../../img/menu_r1_c8_f2.gif','../../../../img/menu_r1_c8_f4.gif','../../../../img/menu_direita_r2_c1_f3.gif','../../../../img/menu_direita_r2_c1_f2.gif','../../../../img/menu_direita_r2_c1_f4.gif','../../../../img/menu_direita_r4_c1_f3.gif','../../../../img/menu_direita_r4_c1_f2.gif','../../../../img/menu_direita_r4_c1_f4.gif','../../../../img/menu_direita_r6_c1_f3.gif','../../../../img/menu_direita_r6_c1_f2.gif','../../../../img/menu_direita_r6_c1_f4.gif')">
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
      <%
if opt = "ok" then%>
  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(17,2,0)
%>
    </td>
                  </tr>
                  <%
	elseif opt = "ok2" then
%>
                  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(20,2,0) 
%>
    </td>
                  </tr>
                  <% 	end if 

 Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS1.Open SQL1, CON0
		
 if RS1.EOF THEN%>
                  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(4,619,1,0)
%>
    </td>
                  </tr>
                  <%else%>
                  <tr> 
                    
    <td height="10"> 
      <%	call mensagens(4,25,0,0) 
	  
	  
	call ultimo(0) 
%>
</td></tr>
<%

end if
%><tr>

            <td valign="top"> <form name="alteracao" method="post" action="index.asp?opt=gera" onSubmit="return checksubmit()">
                
        <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td width="653" height="15" class="tb_tit">Grade de Aulas</td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="200" class="tb_subtit"> 
                    <div align="center">UNIDADE 
                    </div></td>
                  <td width="200" class="tb_subtit"> 
                    <div align="center">CURSO 
                    </div></td>
                  <td width="200" class="tb_subtit"> 
                  	<div align="center">ETAPA 
                  		</div></td>
                  <td width="200" class="tb_subtit"><div align="center">M&Eacute;DIA</div></td>
                  <td width="200" class="tb_subtit"><div align="center"> ORDENA&Ccedil;&Atilde;O</div></td>
                  </tr>
                <% 'if RS1.EOF THEN %>
                <%'else%>
                <tr> 
                  <td width="200"> 
                    <div align="center"> 
                      <select name="unidade" class="select_style" onChange="recuperarCurso(this.value)">
                        <option value="nulo"></option>
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
                  <td width="200"> 
                    <div align="center"> 
                      <div id="divCurso"> 
                        <select class="select_style">
                        </select>
                      </div>
                    </div></td>
                  <td width="200"> 
                  	<div align="center"> 
                  		<div id="divEtapa"> 
                  			<select class="select_style">
                  				</select>
                  			</div>
                  		</div></td>
                  <td width="200" align="center"><select name="media" class="select_style">
                        <option value="nulo"></option>                  
                  <%	
		Set RSp = Server.CreateObject("ADODB.Recordset")
		SQLp = "SELECT * FROM TB_Periodo where NU_periodo <5 order by NU_Periodo"
		RSp.Open SQLp, CON0
		
While not RSp.EOF
num_media=RSp("NU_Periodo")
nome_media=RSp("NO_Periodo")
%>
                        <option value="<%response.Write(num_media)%>"><%response.Write(nome_media)%></option>    

<%
RSp.MOVENEXT
WEND

nome_medias = funcao_vetor_medias("I", "MA", outro)
nomes_medias = split(nome_medias, "#!#")

for sm = 0 to ubound(nomes_medias)
%>
                        <option value="<%response.Write("MA")%>"><%response.Write(nomes_medias(sm))%></option>  

<%Next
%>		
                  				</select></td>
                  <td width="200" align="center"><table width="199" border="0" cellspacing="0" cellpadding="0">
                  	<tr>
                  		<td width="25"><input name="tp_ord" type="radio" class="borda" id="tp_ord" value="C"></td>
                  		<td width="56" class="form_dado_texto">Crescente</td>
                  		<td width="37" align="right"><input name="tp_ord" type="radio" class="borda" id="tp_ord" value="D" checked></td>
                  		<td width="81" class="form_dado_texto">Decrescente</td>
                  		</tr>
               	  </table></td>
                  </tr>
                <%'end if %>
                <tr> 
                  <td height="15" colspan="5" bgcolor="#FFFFFF"><hr></td>
                </tr>
                <tr>
                  <td height="15" colspan="5" bgcolor="#FFFFFF"><table width="1000" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td width="33%">&nbsp;</td>
                                <td width="34%">&nbsp;</td>
                                <td width="33%"> 
                                  <div align="center">
                                    <input name="Submit" type="submit" class="botao_prosseguir" value="Criar Relat&oacute;rio">
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