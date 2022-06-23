<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->

<%

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
nvg=request.QueryString("nvg")
opt = request.QueryString("opt")
ori = request.QueryString("ori")
chave=nvg
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
trava = session("trava") 

cod_cons = request.QueryString("cod_cons")
		
		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON9 = Server.CreateObject("ADODB.Connection") 
		ABRIR9 = "DBQ="& CAMINHO_ax & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON9.Open ABRIR9		

 call VerificaAcesso (CON,chave,nivel)
autoriza=Session("autoriza")




 call navegacao (CON,chave,nivel)
navega=Session("caminho")

if ori="2" or ori="3" then	
		
	
elseif ori="01" then				

	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Item WHERE CO_Item ="& cod_cons
	RS.Open SQL, CON9

	cod_cons = RS("CO_Item")
	nome_item  = RS("NO_Item")
	apelido = RS("NO_Apelido_Item")
	tipo_peso = RS("CO_Tipo_Peso")
	minimo = RS("QT_Estoque_Minimo")
	alerta = RS("QV_Estoque_Minimo")
	grupo_bd = RS("CO_Grupo")
	observacoes= RS("TX_Observacoes")

end if

	%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../../../../js/global.js"></script>
<script type="text/javascript" src="../../../../js/atualiza_select.js"></script>
<script language="JavaScript" type="text/JavaScript">
function formatar(src, mask)
{
  var i = src.value.length;
  var saida = mask.substring(0,1);
  var texto = mask.substring(i)
if (texto.substring(0,1) != saida)
  {
        src.value += texto.substring(0,1);
  }
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
function checksubmit()
{
  if (document.inclusao.nome.value == "")
  {    alert("Por favor, digite um nome para o Item!")
    document.inclusao.nome.focus()
    return false
  }

  return true

}
function MM_showHideLayers() { //v6.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
    obj.visibility=v; }
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

						 function recuperarCidNat(estadonat)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/cid_bairro.asp?opt=c&o=n&f=n", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_cid_nat  = oHTTPRequest.responseText;
resultado_cid_nat = resultado_cid_nat.replace(/\+/g," ")
resultado_cid_nat = unescape(resultado_cid_nat)
document.all.cid_nat.innerHTML =resultado_cid_nat
                                                           }
                                               }
                                               oHTTPRequest.send("c_pub=" + estadonat);
                                   }
						 function recuperarCidRes(estadores)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/cid_bairro.asp?opt=c&o=r&f=n", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_cid_res  = oHTTPRequest.responseText;
resultado_cid_res = resultado_cid_res.replace(/\+/g," ")
resultado_cid_res = unescape(resultado_cid_res)
document.all.cid_res.innerHTML =resultado_cid_res
document.all.bairro_res.innerHTML ="<select class=select_style></select>"
                                                           }
                                               }
                                               oHTTPRequest.send("c_pub=" + estadores);
                                   }

						 function recuperarBairroRes(estadores,cidres)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/cid_bairro.asp?opt=b&o=r&f=n", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_bairro_res  = oHTTPRequest.responseText;
resultado_bairro_res = resultado_bairro_res.replace(/\+/g," ")
resultado_bairro_res = unescape(resultado_bairro_res)
document.all.bairro_res.innerHTML =resultado_bairro_res
                                                           }
                                               }
                                               oHTTPRequest.send("c_pub=" + estadores +"&b_pub=" + cidres);
                                   }
function centraliza(w,h){
//o 120 e o 16 se referem ao tamanho di cabeçalho do navegador e a barra de rolagem respectivamente
    x = parseInt((screen.width - w - 16)/2);
    y = parseInt((screen.height - h - 120)/2);
   //alert(x + '\n' + y);
    document.getElementById('alinha').style.left = x;
    document.getElementById('alinha').style.top = y;
	
//	alert('w '+x +' h '+ y)
}								   
//-->
</script>
</head> 
<%call cabecalho(nivel)%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" background="../../../../img/fundo.gif" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('document.inclusao.nome.focus()');" >
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> </td>
  </tr>
<%if opt="ok" then%>	
	  <tr> 
    <td width="1000" height="10"> 
      <%
	  	call mensagens(4,813,2,0) 
		%>
    </td>
  </tr>
<%end if%>

  <tr> 
    <td width="1000" height="10"> 
      <%
	  if autoriza="no" then
	  	call mensagens(4,9700,1,0) 	  
	  elseif autoriza="1" then
	  	call mensagens(4,9701,0,0) 	  
	  else
	  	call mensagens(4,812,0,0) 
	  end if%>
    </td>

  </tr>

  <tr> 
    <td valign="top"> 
            <%	 
		 if autoriza="no" then			
		elseif ori="02" then	
			action = "bd.asp?opt=inc&nvg="&nvg
		elseif ori="01" then
			action = "bd.asp?opt=alt&nvg="&nvg
		end if		
%>
            <form action="<%response.Write(action)%>" method="post" name="inclusao" id="inclusao" onSubmit="return checksubmit()">
              
        <table width="1000" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td colspan="6" class="tb_tit">Dados Pessoais do Fornecedor</td>
          </tr>
          <tr> 
            <td width="143" height="20" class="tb_corpo"> <div align="right"><font class="form_dado_texto">C&oacute;digo: 
                </font></div></td>
            <td height="20" colspan="5" class="tb_corpo"> <font class="form_dado_texto"> 
              <input name="cod_cons" type="hidden" class="textInput" id="cod_cons" value="<%=cod_cons%>" size="4">
              <font class="form_corpo"> 
              <%RESPONSE.Write(cod_cons)%>
              </font> 
              <input name="tp" type="hidden" id="tp" value="L">
              <input name="acesso" type="hidden" id="acesso" value="2">
              </font></td>
          </tr>
          <tr class="tb_corpo"> 
            <td height="20"><div align="right"><font class="form_dado_texto">Nome: 
                </font></div></td>
            <td height="20" colspan="5"> <font class="form_dado_texto"> 
              <input name="nome" type="text" class="select_style" id="nome" value="<%response.write(nome_item)%>" size="75" maxlength="50">
              </font></td>
          </tr>
          <tr class="tb_corpo"> 
            <td height="20"> <div align="right"><font class="form_dado_texto">Apelido:</font></div></td>
            <td width="119" height="20"> <font class="form_dado_texto"> 
              <input name="apelido" type="text" class="textInput" id="apelido" value="<%response.write(apelido)%>"  size="20" maxlength="15">
              </font></td>
            <td width="282" height="20"><div align="right">&nbsp;<font class="form_dado_texto">Tipo de Medida:</font></div></td>
            <td height="20" colspan="3"><font class="form_dado_texto">
              <select name="tipo_peso" class="select_style" id="tipo_peso">
<%              
if isnull(tipo_peso) or tipo_peso="" then
%>
    <option value="nulo" selected></option>     
<%
else
	if tipo_peso = "CX" then
		selected_Cx = "selected"
	else
		selected_Cx = ""
	end if	
	if tipo_peso = "UN" then
		selected_U = "selected"
	else
		selected_U = ""
	end if
	if tipo_peso = "M" then
		selected_M = "selected"
	else
		selected_M = ""
	end if
	if tipo_peso = "CM" then
		selected_Cm = "selected"
	else
		selected_Cm = ""
	end if	
	if tipo_peso = "KG" then
		selected_Kg = "selected"
	else
		selected_Kg = ""
	end if
	if tipo_peso = "G" then
		selected_G = "selected"
	else
		selected_G = ""
	end if
	if tipo_peso = "L" then
		selected_L = "selected"
	else
		selected_L = ""
	end if
			
end if
	
%> 
             <option value="Cx" <%response.Write(selected_Cx)%>>
                  CX
                </option>              
         
             <option value="UN" <%response.Write(selected_U)%>>
                  UN
                </option>
             <option value="M" <%response.Write(selected_M)%>>
                  M
                </option>  
             <option value="CM" <%response.Write(selected_Cm)%>>
                  CM
                </option>  
             <option value="KG" <%response.Write(selected_Kg)%>>
                  KG
                </option>   
             <option value="G" <%response.Write(selected_G)%>>
                  G
                </option>    
             <option value="L" <%response.Write(selected_L)%>>
                  L
                </option>                                                                          
              </select>
            </font></td>
          </tr>
          <tr class="tb_corpo"> 
            <td height="20"> <div align="right"><font class="form_dado_texto">Grupo:</font></div></td>
            <td height="20"> <div id="cid_res"> 
                <select name="grupo" class="select_style" id="grupo">
                  <%
Set RS2m = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Grupo order by NO_Grupo"
		RS2m.Open SQL2m, CON9
		
while not RS2m.EOF						
co_grupo= RS2m("CO_Grupo")
no_grupo= RS2m("NO_Grupo")

if isnull(grupo_bd) or grupo_bd="" then
%>
                  <option value="0" selected></option>
<%
else

grupo_bd = grupo_bd*1
co_grupo = co_grupo*1
	if grupo_bd = co_grupo then
		selected_grupo = "selected"
	else
		selected_grupo = ""
	end if
end if
	
%>


                  <option value="<%=co_grupo%>" <%response.Write(selected_grupo)%>> 
                  <% =no_grupo%>
                  <%

RS2m.MOVENEXT
WEND
%>
                </select>
              </div></td>
            <td height="20"> <div align="right"><font class="form_dado_texto">Quantidade M&iacute;nima em Estoque:</font></div></td>
            <td width="101" height="20"> <font class="form_dado_texto">
              <input name="minimo" type="text" class="textInput" id="minimo" value="<%response.write(minimo)%>" size="6" maxlength="5">
            </font><font class="form_dado_texto"> <div id="bairro_res"></div></td>
            <td width="102" height="20"><div align="right"><font class="form_dado_texto">Aviso Estoque:</font></div></td>
            <td width="253" height="20"><font class="form_dado_texto">
              <input name="alerta" type="text" class="textInput" id="alerta" value="<%response.write(alerta)%>" size="6" maxlength="5">
            </font></td>
          </tr>
          <tr class="tb_corpo">
            <td height="20" valign="top"
><div align="right"><font class="form_dado_texto">Observa&ccedil;&otilde;es:</font></div></td>
            <td height="20" colspan="5"><font class="form_dado_texto">
              <textarea name="observacoes" cols="100" rows="5" class="textInput" id="observacoes"><%response.write(observacoes)%></textarea>
            </font></td>
          </tr>
          <tr class="tb_corpo"
>
            <td height="30" colspan="6"><hr></td>
          </tr>
          <tr class="tb_corpo"
> 
            <td height="30" colspan="6"><div align="center">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td width="33%">
<div align="center"></div></td>
                    <td width="34%"> 
                      <div align="center"></div></td>
                    <td width="33%">
<div align="center"> 
                        <input name="Submit22" type="submit" class="botao_prosseguir" id="Submit23" value="Confirmar">
                      </div></td>
                  </tr>
                </table>
              </div></td>
          </tr>
        </table>
            </form>

    </td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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