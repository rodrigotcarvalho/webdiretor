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


chave=nvg
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
trava = session("trava") 

	cod_grupo = request.form("grupo")
	exibicao = request.form("exibicao")

		
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
//  if (document.inclusao.nome.value == "")
//  {    alert("Por favor, digite um nome para o fornecedor!")
//    document.inclusao.nome.focus()
//    return false
//  }
      	     
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
                                                            alert("Esse browser n?o tem recursos para uso do Ajax");
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
//o 120 e o 16 se referem ao tamanho di cabe?alho do navegador e a barra de rolagem respectivamente
    x = parseInt((screen.width - w - 16)/2);
    y = parseInt((screen.height - h - 120)/2);
   //alert(x + '\n' + y);
    document.getElementById('alinha').style.left = x;
    document.getElementById('alinha').style.top = y;
	
//	alert('w '+x +' h '+ y)
}								   
//-->
</script>
<script language="javascript">
function mudar_cor_focus(celula){
   celula.style.backgroundColor="#D8FF9D"

}
function mudar_cor_blur_par(celula){
   celula.style.backgroundColor="#FFFFFF"
} 
function mudar_cor_blur_impar(celula){
   celula.style.backgroundColor="#FFFFE1"
} 
function mudar_cor_blur_erro(celula){
   celula.style.backgroundColor="#CC0000"
}    
    function keyPressed(TB, e, max_right, max_bottom)  
    { 
        if (e.keyCode == 40 || e.keyCode == 13) { // arrow down 
            if (TB.split("c")[0] < max_bottom) 
            document.getElementById(eval(TB.split("c")[0] + '+1') + 'c' + TB.split("c")[1]).focus(); 
            if (TB.split("c")[0] == max_bottom) 
            document.getElementById(1 + 'c' + TB.split("c")[1]).focus();


        } 
  
        if (e.keyCode == 38) { // arrow up 
            if(TB.split("c")[0] > 1) 
            document.getElementById(eval(TB.split("c")[0] + '-1') + 'c' + TB.split("c")[1]).focus(); 
            if (TB.split("c")[0] == 1) 
            document.getElementById(max_bottom + 'c' + TB.split("c")[1]).focus(); 
		
        } 
  
        if (e.keyCode == 37) { // arrow left 
            if(TB.split("c")[1] > 1) 
            document.getElementById(TB.split("c")[0] + 'c' + eval(TB.split("c")[1] + '-1')).focus();             
            if (TB.split("c")[1] == 1) 
            document.getElementById(TB.split("c")[0] + 'c' + max_right).focus(); 

		}   
  
        if (e.keyCode == 39) { // arrow right 
            if(TB.split("c")[1] < max_right) 
            document.getElementById(TB.split("c")[0] + 'c' + eval(TB.split("c")[1] + '+1')).focus();  
            if (TB.split("c")[1] == max_right) 
            document.getElementById(TB.split("c")[0] + 'c' + 1).focus(); 

		}                  
    }
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
</script> 
</head> 
<%call cabecalho(nivel)%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" background="../../../../img/fundo.gif" topmargin="0" marginwidth="0" marginheight="0">
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> </td>
  </tr>
  <tr> 
    <td width="1000" height="10"> 
      <%
	  if autoriza="no" then
	  	call mensagens(4,9700,1,0) 	  
	  elseif autoriza="1" then
	  	call mensagens(4,9701,0,0) 	  
	  else
	  	call mensagens(4,9714,0,0) 
	  end if%>
    </td>

  </tr>
  <tr>
    <td height="10" class="tb_tit">Informe os crit&eacute;rios 
    para pesquisa</td>
  </tr>
  <tr>
    <td height="10">
<FORM name="formulario" METHOD="POST" ACTION="altera.asp?nvg=<%response.Write(nvg)%>">    
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="tb_subtit"><div align="center">Grupo</div></td>
        <td width="50%" align="center" class="tb_subtit"><div align="center">Exibi&ccedil;&atilde;o</div></td>
      </tr>
      <tr>
        <td width="50%"><div align="center">
          <select name="grupo" class="select_style" id="grupo">
            <option value="nulo" selected></option>
            <%      Set RS = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Grupo order by NO_Grupo"
		RS.Open SQL2m, CON9
		
while not RS.EOF						
co_grupo= RS("CO_Grupo")
no_grupo= RS("NO_Grupo")

if isnull(co_grupo) or co_grupo="" then
%>
            <%
else
	if cod_grupo<>"nulo" then
		cod_grupo = cod_grupo*1
		co_grupo = co_grupo*1
		if cod_grupo = co_grupo then
			selected = "selected"
		else
			selected = ""
		end if
	end if	
end if
	
%>
            <option value="<%=co_grupo%>" <%response.Write(selected)%>>
              <% =no_grupo%>
              <%

RS.MOVENEXT
WEND
%>
            </select>
        </div></td>
        <td width="50%" align="center" class="<%response.Write(classe)%>"><div align="center"><font class="form_dado_texto">
          <select name="exibicao" class="select_style">
          <%	
		    if exibicao = "tds" then
				selected_tds = "selected"
				selected_min = ""
			else
				selected_tds = ""
				selected_min = "selected"
			end if%>
            <option value="tds" <%response.Write(selected_tds)%>>Todos</option>
            <option value="min" <%response.Write(selected_min)%>>Apenas os com estoque m&iacute;nimo</option>
          </select>
        </font></div></td>
      </tr>
      <tr>
        <td colspan="2">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="2"><hr width="1000"></td>
      </tr>
      <tr>
        <td colspan="2" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="33%">&nbsp;</td>
            <td width="34%">&nbsp;</td>
            <td width="33%" align="center"><input name="SUBMIT" type=SUBMIT class="botao_prosseguir" value="Prosseguir"></td>
          </tr>
        </table></td>
      </tr>
    </table>
    </form>
    </td>
  </tr>
  <tr>
    <td height="10">&nbsp;</td>
  </tr>

  <tr> 
    <td valign="top">             
              
        <table width="1000" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td colspan="9" class="tb_tit">Dados do Invent&aacute;rio 
              <input name="cod_grupo" type="hidden" class="textInput" id="cod_grupo" value="<%response.Write(cod_grupo)%>" size="4"></td>
          </tr>
        <tr class="tb_subtit">
          <td width="50" align="center">&nbsp;</td>
            <td width="100" height="20" align="center">C&oacute;digo</td>
            <td width="200" height="20" align="center">Nome</td>
            <td width="140" height="20" align="center">Apelido</td>
            <td width="140" height="20" align="center">Tipo de Medida</td>
            <td width="140" height="20" align="center">Grupo</td>
            <td width="90" align="center">Quantidade M&iacute;nima em Estoque</td>
            <td width="90" height="20" align="center">Quantidade Atual</td>
            <td width="50" align="center">Situa&ccedil;&atilde;o</td>
          </tr>          
          <%	
		  
	Set RSI = Server.CreateObject("ADODB.Recordset")
    if cod_grupo = "nulo" then
		SQL = "SELECT count(CO_Item) as TOTAL FROM TB_Item"
	else	
		nome_grupo = GeraNomesNovaVersao("GRP_ITEM",cod_grupo,variavel2,variavel3,variavel4,variavel5,CON9,outro)	
		SQL = "SELECT count(CO_Item) as TOTAL FROM TB_Item WHERE CO_Grupo ="& cod_grupo
	end if
	RSI.Open SQL, CON9		  

	if RSI.EOF then
		total = 0	
	else
		total = RSI("TOTAL")		
	end if
	
	if total = 0 then
%>
           <tr>
            <td colspan="9" align="center" class="form_dado_texto" >&nbsp; </td>
           </tr>         
                  
          <tr>
            <td colspan="9" align="center" class="form_dado_texto" > Nenhum item cadastrado para o grupo <%response.Write(nome_grupo)%></td>
           </tr>
<%	end if

	Set RS = Server.CreateObject("ADODB.Recordset")
    if cod_grupo = "nulo" then
		SQL = "SELECT  * FROM TB_Item ORDER BY NO_Item"
	else	
		SQL = "SELECT  * FROM TB_Item WHERE CO_Grupo ="& cod_grupo&" ORDER BY NO_Item"
	end if
	RS.Open SQL, CON9	
		
	linha = 0
	
	while not RS.EOF
		cod_cons = RS("CO_Item")
		nome  = RS("NO_Item")
		apelido = RS("NO_Apelido_Item")
		tipo_peso = RS("CO_Tipo_Peso")
		minimo = RS("QT_Estoque_Minimo")
		estoque = RS("QT_Atual")
		alerta = RS("QV_Estoque_Minimo")		
		linha = linha+1
		if linha mod 2 =0 then
			classe = "tb_fundo_linha_par" 
			onblur="mudar_cor_blur_par"
		else 
			classe ="tb_fundo_linha_impar"
			onblur="mudar_cor_blur_impar"
		end if 		
		
		  if isnull(minimo) or minimo = "" then
			minimo_estoque = 0		  
		  else
			minimo_estoque = minimo		  
		  end if	
		  
		  if isnull(estoque) or estoque = "" then
			estoque_atual = 0		  
		  else
			estoque_atual = estoque		  
		  end if	
		   
		  if isnull(alerta) or alerta = "" then
			alerta = 2		  
		  else
			alerta = alerta		  
		  end if	
		  estoque_atual = estoque_atual*1	   
		  testa_alerta = alerta*minimo	  
		  if estoque_atual<testa_alerta then
			cor = "FF0000"
			escreve = "Comprar"					
		  else
			cor = "0000FF"
			escreve = "OK"	
		  end if		
		
		if (exibicao="min" and escreve = "Comprar") or exibicao="tds" then
%>
          
                  
          <tr class="<%response.Write(classe)%>" id="<%response.Write("celula"&linha)%>">
            <td width="50" align="center"><%						  
			  if escreve = "Comprar" then
				%>
					<img src="../../../../img/red_arrow_right.gif" width="25" height="18">				
			  <%end if%>
            </td>
            <td width="100" height="20" align="center"><%RESPONSE.Write(cod_cons)%><input name="cod_cons" type="hidden" class="textInput" id="cod_cons" value="<%response.Write(cod_cons)%>" size="4"></td>
            <td width="200" height="20" align="center"><a href="#" onClick="MM_openBrWindow('historico.asp?cod_item=<%response.write(cod_cons)%>','Historico','menubar=yes,scrollbars=yes,width=650,height=400')"><%RESPONSE.Write(nome)%>
            </a></td>
            <td width="140" height="20" align="center"><%RESPONSE.Write(apelido)%></td>
            <td width="140" height="20" align="center"><%RESPONSE.Write(tipo_peso)%></td>
            <td width="140" height="20" align="center"><%RESPONSE.Write(nome_grupo)%></td>
            <td width="90" height="20" align="center"><%RESPONSE.Write(minimo)%></td>
          <td width="90" height="20" align="center"><%RESPONSE.Write(estoque)%></td
          >
          <td width="50" align="center"><%
		  		  			  
			
			  response.Write("<font color="&cor&">"&escreve&" </font>")			  
			  
		  
		  %></td
          >
          </tr>
          <%
		  end if
	  RS.MOVENEXT
	  WEND%>
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
response.redirect("../../../../inc/erro.asp")
end if
%>