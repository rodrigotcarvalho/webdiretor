<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<%
opt = request.QueryString("opt")

ano_letivo = session("ano_letivo")
co_usr = session("co_user")
nivel=4
nvg = request.QueryString("nvg")
session("nvg")=nvg
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

	Set RS = Server.CreateObject("ADODB.Recordset")
	CONEXAO = "Select * from TB_Situacao_Aluno order by CO_Situacao"
	Set RS = CON0.Execute(CONEXAO)
	
	conta=0
	While not RS.EOF
	co_situac_aluno=RS("CO_Situacao")
	nom_situac_aluno=RS("TX_Descricao_Situacao")
		if co_situac_aluno = "C" or co_situac_aluno = "P" or co_situac_aluno = "T" or co_situac_aluno = "E" THEN 
			conta=conta+1
			if conta=1 then
				vetor_co_situac=co_situac_aluno	
				vetor_nom_situac=nom_situac_aluno
			else
				vetor_co_situac=vetor_co_situac&"#!#"&co_situac_aluno
				vetor_nom_situac=vetor_nom_situac&"#!#"&nom_situac_aluno
			end if
		end if	
	RS.MOVENEXT
	WEND
co_situacoes=split(vetor_co_situac,"#!#")		
situacoes=split(vetor_nom_situac,"#!#")	
	
call navegacao (CON,nvg,nivel)
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

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=c2", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
//recuperarEtapa()
                                                           }
                                               }

                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }								   								   

 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
function checkSubmit()
{
  if (document.grafico.tema.value == "nulo")
  {    alert("Por favor selecione um Tema Estatístico!")
    document.grafico.tema.focus()
    return false
  }
  return true
}

function stAba(conteudo)
	{
		this.conteudo = conteudo;
	}
 
	var arAbas = new Array();
	arAbas[0] = new stAba('qaa');
	arAbas[1] = new stAba('qpr');
	arAbas[2] = new stAba('qde');	
 
	function AlternarConteudo(conteudo)
	{
		for (i=0;i<arAbas.length;i++)
		{
			c = document.getElementById(arAbas[i].conteudo)
			c.style.display = 'none';
		}
		c = document.getElementById(conteudo)
		c.style.display = '';
	}
                        </script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../../img/menu_r1_c2_f3.gif','../../../../img/menu_r1_c2_f2.gif','../../../../img/menu_r1_c2_f4.gif','../../../../img/menu_r1_c4_f3.gif','../../../../img/menu_r1_c4_f2.gif','../../../../img/menu_r1_c4_f4.gif','../../../../img/menu_r1_c6_f3.gif','../../../../img/menu_r1_c6_f2.gif','../../../../img/menu_r1_c6_f4.gif','../../../../img/menu_r1_c8_f3.gif','../../../../img/menu_r1_c8_f2.gif','../../../../img/menu_r1_c8_f4.gif','../../../../img/menu_direita_r2_c1_f3.gif','../../../../img/menu_direita_r2_c1_f2.gif','../../../../img/menu_direita_r2_c1_f4.gif','../../../../img/menu_direita_r4_c1_f3.gif','../../../../img/menu_direita_r4_c1_f2.gif','../../../../img/menu_direita_r4_c1_f4.gif','../../../../img/menu_direita_r6_c1_f3.gif','../../../../img/menu_direita_r6_c1_f2.gif','../../../../img/menu_direita_r6_c1_f4.gif')">
<% call cabecalho (nivel)
	  %>
<table width="1006" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
                    
    <td width="1004" height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
	  </td>
	  </tr>
<tr> 
                    
    <td height="10"> 
      <%	call mensagens(4,0,0,0) 
%>
</td></tr>
<tr>

            <td valign="top">
        <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">
          <tr> 
            <td width="1000"><form name="grafico" method="post" action="grafico.asp" onSubmit="return checkSubmit()">
            <table width="1000" border="0" cellspacing="0" cellpadding="0">
              <tr class="tb_subtit">
                <td class="tb_tit">Preencha os campos abaixo</td>
                <td class="tb_tit">&nbsp;</td>
                <td class="tb_tit">&nbsp;</td>
                <td class="tb_tit">&nbsp;</td>
              </tr>
              <tr class="tb_subtit">
                <td class="tb_subtit"><div align="right">Tema Estat&iacute;stico</div></td>
                <td class="tb_subtit">&nbsp;
                  <select name="tema" class="select_style" id="select" onChange="AlternarConteudo(this.value)">
                    <option value="nulo" selected></option>
                    <option value="qaa">Quanto ao Aluno</option>
                    <option value="qpr">Quanto aos Pais e Respons&aacute;veis </option>
                    <option value="qde">Quanto a Distribui&ccedil;&atilde;o na Escola</option>
                  </select>
                </td>
                <td class="tb_subtit"><div align="right">UNIDADE</div></td>
                <td class="tb_subtit"><div align="left">&nbsp;
                  <select name="unidade" class="borda">
                    <option value="999990" selected></option>
                    <%		
		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0

		While not RS0.EOF
		NU_Unidade = RS0("NU_Unidade")
		NO_Abr = RS0("NO_Abr")
		
		%>
                    <option value="<%response.Write(NU_Unidade)%>">
                      <%response.Write(NO_Abr)%>
                      </option>
                    <%
		RS0.MOVENEXT
		WEND
		%>
                  </select>
                </div></td>
                </tr>
              <tr>
                <td colspan="4">
<div id="qaa" style="display: none">                
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="50%" class="form_dado_texto"><div align="right">Por G&ecirc;nero</div></td>
                <td width="50%"><label>
                  <input name="qqa" type="radio" class="option_button" id="qqa" value="gen" checked>
                </label></td>
              </tr>
              <tr>
                <td width="50%" class="form_dado_texto"><div align="right">Por Idade</div></td>
                <td width="50%"><input name="qqa" type="radio" class="option_button" id="qqa" value="idade"></td>
              </tr>
              <tr>
                <td width="50%" class="form_dado_texto"><div align="right">Pelo Ano de Nascimento</div></td>
                <td width="50%"><input name="qqa" type="radio" class="option_button" id="qqa" value="ano"></td>
              </tr>
              <tr>
                <td class="form_dado_texto"><div align="right">Pelo M&ecirc;s de Nascimento</div></td>
                <td><input name="qqa" type="radio" class="option_button" id="qqa" value="mes"></td>
              </tr>
              </table>
</div>   
<div id="qpr" style="display: none">                
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td class="form_dado_texto"><div align="right"> Por Estado Civil&nbsp; </div></td>
                <td width="50%"><label>
                  <input name="qpr" type="radio" class="option_button" id="qpr" value="estc" checked>
                </label></td>
              </tr>
              <tr>
                <td class="form_dado_texto"><div align="right"> Por Resp. Financeiros </div></td>
                <td width="50%"><input name="qpr" type="radio" class="option_button" id="qpr" value="respf"></td>
              </tr>
              <tr>
                <td class="form_dado_texto"><div align="right"> Por Resp. Pedag&oacute;gicos </div></td>
                <td width="50%"><input name="qpr" type="radio" class="option_button" id="qpr" value="respp"></td>
              </tr>
              </table>
</div>
<div id="qde" style="display: none">                
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td class="form_dado_texto"><div align="right"> Por Unidade </div></td>
                <td width="50%"><label>
                  <input name="qde" type="radio" class="option_button" id="qde" value="unidd" checked>
                </label></td>
              </tr>
              <tr>
                <td class="form_dado_texto"><div align="right"> Por Curso </div></td>
                <td width="50%"><input name="qde" type="radio" class="option_button" id="qde" value="curso"></td>
              </tr>
              <tr>
                <td class="form_dado_texto"><div align="right"> Por Etapas&nbsp; </div></td>
                <td width="50%"><input name="qde" type="radio" class="option_button" id="qde" value="etapa"></td>
              </tr>
              </table>
</div>           
                
                
                </td>
              </tr>
              <tr>
                <td colspan="4"><hr></td>
                </tr>
              <tr>
                <td width="20%">&nbsp;</td>
                <td width="9%">&nbsp;</td>
                <td width="13%">&nbsp;</td>
                <td width="13%"><font size="2" face="Arial, Helvetica, sans-serif">
                  <div align="center">
                    <input name="Submit" type="submit" class="botao_prosseguir" id="Submit5" value="Prosseguir">
                    </div>
                  </font></td>
              </tr>
            </table></form></td>
          </tr>
        </table>
</td>
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