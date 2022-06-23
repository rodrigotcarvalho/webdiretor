<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->

<%

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
nvg = session("nvg")
session("nvg")=nvg
opt = request.QueryString("opt")

ano_info=nivel&"-"&chave&"-"&ano_letivo
trava = session("trava") 

if opt="ok1" or opt="ok2" or opt="vt" or opt="ok3" then	
	cod_np = session("cod_np")
	dia_de = session("dia_de")
	mes_de = session("mes_de")
	ano_de = session("ano_de")
	dia_ate = session("dia_ate")
	mes_ate = session("mes_ate")
	ano_ate = session("ano_ate")
	situacao = session("situacao")	
else
	cod_np = request.form("np")
	dia_de = request.form("dia_de")
	mes_de = request.form("mes_de")
	ano_de = request.form("ano_de")	
	
	dia_ate = request.form("dia_ate")
	mes_ate = request.form("mes_ate")	
	ano_ate = request.form("ano_ate")	
	situacao = request.form("situacao")						
end if	


session("cod_np") = cod_np
session("dia_de") = dia_de
session("mes_de") = mes_de
session("ano_de") = ano_de
session("dia_ate") = dia_ate
session("mes_ate") = mes_ate
session("ano_ate") = ano_ate	
session("situacao") = situacao

dados_msg = cod_np&"$!$"&dia_de&"$!$"&mes_de&"$!$"&ano_de&"$!$"&dia_ate&"$!$"&mes_ate&"$!$"&ano_ate&"$!$"&situacao

'Select case ordem
'
'case "dt"
'ordena="DA_NotaF"
'
'case "nf"
'ordena="NU_NotaF"
'
'case "fr"
'ordena="NO_Fornecedor"
'
'case "vn"
'ordena="VA_NotaF"
'
'
'end select	

data_de=mes_de&"/"&dia_de&"/"&ano_de


data_ate=mes_ate&"/"&dia_ate&"/"&ano_ate


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON9 = Server.CreateObject("ADODB.Connection") 
		ABRIR9 = "DBQ="& CAMINHO_ax & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON9.Open ABRIR9		

 call VerificaAcesso (CON,nvg,nivel)
autoriza=Session("autoriza")




 call navegacao (CON,nvg,nivel)
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
var checkflag = "false";
function check(field) {
if (checkflag == "false") {
for (i = 0; i < field.length; i++) {
field[i].checked = true;}
checkflag = "true";
return "Desmarcar Todos"; }
else {
for (i = 0; i < field.length; i++) {
field[i].checked = false; }
checkflag = "false";
return "Marcar Todos"; }
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
  if (document.inclusao.nome_grupo.value == "")
  {    alert("Por favor, digite um nome para o Grupo!")
    document.inclusao.nome_grupo.focus()
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
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
function checkTheBox() {
   var chk = document.getElementsByName('nota_fiscal')
    var len = chk.length

    for(i=0;i<len;i++)
    {
         if(chk[i].checked){
        return true;
          }
    }
	alert("Pelo menos uma nota fiscal deve ser selecionada!")		
    return false;
    }	
</script> 
</head> 
<%call cabecalho(nivel)%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" background="../../../../img/fundo.gif" topmargin="0" marginwidth="0" marginheight="0" onLoad="document.inclusao.np.focus()">
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> </td>
  </tr>
<%if opt="ok1" then%>	
	  <tr> 
    <td width="1000" height="10"> 
      <%
	  	call mensagens(4,831,2,0) 
		%>
    </td>
  </tr>
<%elseif opt="ok2" then%>	
	  <tr> 
    <td width="1000" height="10"> 
      <%
	  	call mensagens(4,830,2,0) 
		%>
    </td>
  </tr>  
<%elseif opt="ok3" then%>	
	  <tr> 
    <td width="1000" height="10"> 
      <%
	  	call mensagens(4,829,2,0) 
		%>
    </td>
  </tr>    
<%end if
%>
  <tr> 
    <td width="1000" height="10"> 
      <%
	  if autoriza="no" then
	  	call mensagens(4,9700,1,0) 	  
	  elseif autoriza="1" then
	  	call mensagens(4,9701,0,0) 	  
	  else
	  	call mensagens(4,9704,0,0) 
	  end if%>
    </td>

  </tr>
<%	  if autoriza<>"no" then
	  %>
  <tr> 
    <td width="1000" height="10"> 
      <%	 
	  	call mensagens(4,645,0,"R23") 
	%>
    </td>
  </tr>    
	<%  end if%>      

  <tr> 
    <td height="100" valign="top"> 
            <%	 
		 if autoriza="no" then			
		else		
			action = "resumo.asp?nvg="&nvg
		end if		
%>
            <form action="<%response.Write(action)%>" method="post" name="inclusao" id="inclusao">
              
        <table width="1000" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td colspan="3" class="tb_tit">Crit&eacute;rios Informados</td>
          </tr>
        <tr class="tb_subtit">
            <td width="25%" height="20" align="center">N&uacute;mero do Pedido</td>
            <td width="50%" height="20" align="center">Per&iacute;odo Solicitado</td>
            <td width="25%" align="center"><div align="center">Situa&ccedil;&atilde;o do Pedido:</div></td>
          </tr>          
	  <tr class="<%response.Write(classe)%>" id="<%response.Write("celula"&linha)%>">
					<td width="25%" height="20" align="center"><span class="form_dado_texto">
				    <input name="np" type="text" class="textInput" id="np" Value ="<%response.Write(cod_np)%>" size="20" maxlength="30">
					</span></td>
					<td width="50%" height="20" align="center"><span class="form_dado_texto">
					  <select name="dia_de" id="dia_de" class="select_style">
					    <% 
							 For i =1 to 31
							 dia_de=dia_de*1
							 if dia_de=i then 
								if dia_de<10 then
								dia_de="0"&dia_de
								end if
							 %>
					    <option value="<%response.Write(dia_de)%>" selected> 
				        <%response.Write(dia_de)%>
				        </option>
					    <% else
									if i<10 then
									i="0"&i
								end if
							%>
					    <option value="<%response.Write(i)%>"> 
				        <%response.Write(i)%>
				        </option>
					    <% end if 
							next
							%>
				    </select>
                    / 
                    <select name="mes_de" id="mes_de" class="select_style">
                        <%mes_de=mes_de*1
								if mes_de="1" or mes_de=1 then%>
                        <option value="1" selected>janeiro</option>
                        <% else%>
                        <option value="1">janeiro</option>
                        <%end if
								if mes_de="2" or mes_de=2 then%>
                        <option value="2" selected>fevereiro</option>
                        <% else%>
                        <option value="2">fevereiro</option>
                        <%end if
								if mes_de="3" or mes_de=3 then%>
                        <option value="3" selected>mar&ccedil;o</option>
                        <% else%>
                        <option value="3">mar&ccedil;o</option>
                        <%end if
								if mes_de="4" or mes_de=4 then%>
                        <option value="4" selected>abril</option>
                        <% else%>
                        <option value="4">abril</option>
                        <%end if
								if mes_de="5" or mes_de=5 then%>
                        <option value="5" selected>maio</option>
                        <% else%>
                        <option value="5">maio</option>
                        <%end if
								if mes_de="6" or mes_de=6 then%>
                        <option value="6" selected>junho</option>
                        <% else%>
                        <option value="6">junho</option>
                        <%end if
								if mes_de="7" or mes_de=7 then%>
                        <option value="7" selected>julho</option>
                        <% else%>
                        <option value="7">julho</option>
                        <%end if%>
                        <%if mes_de="8" or mes_de=8 then%>
                        <option value="8" selected>agosto</option>
                        <% else%>
                        <option value="8">agosto</option>
                        <%end if
								if mes_de="9" or mes_de=9 then%>
                        <option value="9" selected>setembro</option>
                        <% else%>
                        <option value="9">setembro</option>
                        <%end if
								if mes_de="10" or mes_de=10 then%>
                        <option value="10" selected>outubro</option>
                        <% else%>
                        <option value="10">outubro</option>
                        <%end if
								if mes_de="11" or mes_de=11 then%>
                        <option value="11" selected>novembro</option>
                        <% else%>
                        <option value="11">novembro</option>
                        <%end if
								if mes_de="12" or mes_de=12 then%>
                        <option value="12" selected>dezembro</option>
                        <% else%>
                        <option value="12">dezembro</option>
                        <%end if%>
                    </select>
                    / 
                    <%response.write(ano_letivo)%>
                    <font class="form_dado_texto">
                    <input name="ano_de" type="hidden" id="ano_de" value="<%response.write(ano_letivo)%>">
                    </font>                    at&eacute; 
                      <select name="dia_ate" id="dia_ate" class="select_style">
                        <% 
							 For i =1 to 31
							 dia_ate=dia_ate*1
							 if dia_ate=i then 
								if dia_ate<10 then
								dia_ate="0"&dia_ate
								end if
							 %>
                          <option value="<%response.Write(dia_ate)%>" selected> 
                          <%response.Write(dia_ate)%>
                          </option>
                          <% else
								if i<10 then
								i="0"&i
								end if
							%>
                          <option value="<%response.Write(i)%>"> 
                          <%response.Write(i)%>
                          </option>
                          <% end if 
							next
							%>
                    </select>
                    / 
                    <select name="mes_ate" id="mes_ate" class="select_style">
                        <%mes_ate=mes_ate*1
								if mes_ate="1" or mes_ate=1 then%>
                        <option value="1" selected>janeiro</option>
                        <% else%>
                        <option value="1">janeiro</option>
                        <%end if
								if mes_ate="2" or mes_ate=2 then%>
                        <option value="2" selected>fevereiro</option>
                        <% else%>
                        <option value="2">fevereiro</option>
                        <%end if
								if mes_ate="3" or mes_ate=3 then%>
                        <option value="3" selected>mar&ccedil;o</option>
                        <% else%>
                        <option value="3">mar&ccedil;o</option>
                        <%end if
								if mes_ate="4" or mes_ate=4 then%>
                        <option value="4" selected>abril</option>
                        <% else%>
                        <option value="4">abril</option>
                        <%end if
								if mes_ate="5" or mes_ate=5 then%>
                        <option value="5" selected>maio</option>
                        <% else%>
                        <option value="5">maio</option>
                        <%end if
								if mes_ate="6" or mes_ate=6 then%>
                        <option value="6" selected>junho</option>
                        <% else%>
                        <option value="6">junho</option>
                        <%end if
								if mes_ate="7" or mes_ate=7 then%>
                        <option value="7" selected>julho</option>
                        <% else%>
                        <option value="7">julho</option>
                        <%end if%>
                        <%if mes_ate="8" or mes_ate=8 then%>
                        <option value="8" selected>agosto</option>
                        <% else%>
                        <option value="8">agosto</option>
                        <%end if
								if mes_ate="9" or mes_ate=9 then%>
                        <option value="9" selected>setembro</option>
                        <% else%>
                        <option value="9">setembro</option>
                        <%end if
								if mes_ate="10" or mes_ate=10 then%>
                        <option value="10" selected>outubro</option>
                        <% else%>
                        <option value="10">outubro</option>
                        <%end if
								if mes_ate="11" or mes_ate=11 then%>
                        <option value="11" selected>novembro</option>
                        <% else%>
                        <option value="11">novembro</option>
                        <%end if
								if mes_ate="12" or mes_ate=12 then%>
                        <option value="12" selected>dezembro</option>
                        <% else%>
                        <option value="12">dezembro</option>
                        <%end if%>
                    </select>
                    / 
                    <%response.write(ano_letivo)%>
					<input name="ano_ate" type="hidden" id="ano_ate" value="<%response.write(ano_letivo)%>">
					</span></td>
					<td width="25%" align="center"><div align="center"><font class="form_dado_texto">
                    <select name="situacao" class="select_style">
			          <% if situacao="td" then%>                    
                      <option value="td" selected>Todas</option>
					    <%else%>
				      <option value="td" >Todas</option>
					    <%end if%>                     
				        <% if situacao="at" then%>                    
                        <option value="at" selected>Atendido</option>
					    <%else%>
					    <option value="at" >Atendido</option>
					    <%end if%>     
				        <% if situacao="pt" then%>                    
                        <option value="pt" selected>Pendente</option>
					    <%else%>
                      	<option value="pt" >Pendente</option>
					    <%end if%> 
				        <% if situacao="cd" then%>                    
                        <option value="cd" selected>Cancelado</option>
					    <%else%>
                        <option value="cd" >Cancelado</option>
					    <%end if%>                                                                  


                    </select>
                  					  
				    </font></div></td>
		  </tr>	
	
          
          <tr class="tb_corpo"
>
            <td height="30" colspan="3"><hr></td>
          </tr>
          <tr class="tb_corpo"
> 
            <td height="30" colspan="3"><div align="center">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td width="33%">
<div align="center"></div></td>
                    <td width="34%"> 
                      <div align="center"></div></td>
                    <td width="33%">
<div align="center"> 
                        <input name="Submit" type="submit" class="botao_prosseguir" id="Submit" value="Prosseguir">
                      </div></td>
                  </tr>
                </table>
              </div></td>
          </tr>
          </table>
            </form>
    </td>
  </tr>
  <tr><td valign="top">
  <form action="redireciona.asp" method="post" name="redireciona" id="redireciona" onSubmit="return checkTheBox()">
  <table border="0" cellpadding="0" cellspacing="0"><tr class="tb_corpo">
            <td height="10" colspan="2" class="tb_tit"
>Movimenta&ccedil;&otilde;es de Estoque Detalhadas</td>
          </tr>
          <tr class="tb_corpo"
>
            <td height="30" colspan="2"><table width="1000" border="0" cellspacing="0" cellpadding="0">
              <tr class="tb_subtit">
                <td width="20" height="10"><input type="checkbox" name="todos" class="borda" value="" onClick="this.value=check(this.form.nota_fiscal)"></td>
                <td width="60" align="center"> N&ordm; Pedido </td>
                <td width="90" align="center"> Data do Pedido </td>
                <td width="100"><div align="left"> Projeto <font class="form_dado_texto">
                  <input name="cod" type="hidden" id="cod" value="<%=codigo%>">
                  <input name="data_de" type="hidden" class="textInput" id="data_de"  value="<%response.Write(data_de)%>" size="75" maxlength="50">
                  <input name="data_inicio" type="hidden" class="textInput" id="data_inicio"  value="<%response.Write(data_inicio)%>" size="75" maxlength="50">
                  <input name="data_ate" type="hidden" class="textInput" id="data_ate"  value="<%response.Write(data_ate)%>" size="75" maxlength="50">
                  </font><font class="form_dado_texto">
                    <input name="data_fim" type="hidden" class="textInput" id="data_fim"  value="<%response.Write(data_fim)%>" size="75" maxlength="50">
                  </font>
                  <input type="hidden" name="dados_msg" id="dados_msg">
                </div></td>
                <td width="145" align="center"> Unidade </td>
                <td width="50" align="center"> Curso </td>
                <td width="140" align="center"> Etapa </td>
                <td width="50" align="center"> Turma </td>
                <td width="165" align="center"> Solicitado por </td>
                <td width="90" align="center"> Situa&ccedil;&atilde;o </td>
                <td width="91" align="center"> Atendido em </td>
              </tr>
              <tr>
                <td colspan="11"><hr width="1000"></td>
              </tr>
              <%			
				

if cod_np="" or isnull(cod_np) then
	sql_cod = ""
else
	sql_cod = "NU_Pedido ="& cod_np&" AND "
end if
if situacao="td" then
	sql_st = ""
elseif situacao="at" then
	sql_st = "ST_Pedido = 'A' AND "
elseif situacao="pt" then
	sql_st = "ST_Pedido = 'P' AND "
elseif situacao="cd" then 
	sql_st = "ST_Pedido = 'C' AND "
end if
	
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Mov_Estoque, TB_Projeto WHERE TB_Mov_Estoque.CO_Projeto = TB_Projeto.CO_Projeto AND "&sql_cod&sql_st&"(TB_Mov_Estoque.DA_Pedido BETWEEN #"&data_de&"# AND #"&data_ate&"#) order by TB_Mov_Estoque.DA_Pedido, TB_Mov_Estoque.NU_Pedido"		
		RS.Open SQL, CON9

if RS.EOF	then	
%>
              <tr>
                <td width="20">&nbsp;</td>
                <td colspan="10" align="center" class="form_dado_texto">Nenhuma movimenta&ccedil;&atilde;o encontrada para os crit&eacute;rios informados</td>
                </tr>
              <%else
check = 2
WHILE not RS.EOF
  if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if 
  
seq_pd=RS("NU_Pedido")
da_pd=RS("DA_Pedido")
no_projeto=RS("NO_Projeto")
unidade=RS("NU_Unidade")
curso=RS("CO_Curso")
etapa=RS("CO_Etapa")
turma=RS("CO_Turma")
observacao=RS("TX_Observa")
situacao=RS("ST_Pedido")
da_atendido=RS("DA_Atendido")
co_usu_reg=RS("CO_Usuario")

'RESPONSE.Write(unidade&"<br>")
'		Set RS = Server.CreateObject("ADODB.Recordset")
'		SQL = "SELECT * FROM TB_Unidade where NU_Unidade = "& unidade
'		RS.Open SQL, CON0
'	
'		IF RS.eof THEN
'			no_unidade= ""
'		ELSE
'			no_unidade= RS("NO_Unidade")
'		END IF
'
'
'		Set RS = Server.CreateObject("ADODB.Recordset")
'		SQL = "SELECT * FROM TB_Curso where CO_Curso = '"& curso &"'"
'		RS.Open SQL, CON0
'	
'		IF RS.eof THEN
'			no_curso= ""
'		ELSE
'			no_curso= RS("NO_Abreviado_Curso")
'		END IF		
'		
'		Set RS = Server.CreateObject("ADODB.Recordset")
'		SQL = "SELECT * FROM TB_Etapa where CO_Curso = '"& curso &"' and CO_Etapa = '"& etapa &"'"
'		RS.Open SQL, CON0
'	
'		IF RS.eof THEN
'			no_etapa= ""
'		ELSE
'			no_etapa= RS("NO_Etapa")
'		END IF			

no_unidade = GeraNomesNovaVersao("U",unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)
no_curso = GeraNomesNovaVersao("CA",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro)
no_etapa = GeraNomesNovaVersao("E",curso,etapa,variavel3,variavel4,variavel5,CON0,outro)


Select case situacao

case "A"
st_movim="Atendido"

case "P"
st_movim="Pendente"

case "C"
st_movim="Cancelado"

end select	

		
if co_usu_reg="" or isnull(co_usu_reg) then
	no_registrador=""
else

		Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& co_usu_reg
		RSu.Open SQLu, CON

	IF RSu.EOF then
		no_registrador=""
	else
		no_registrador=RSu("NO_Usuario")
	end if		
end if			


optobr=seq_pd&"?"&da_pd

Session("data_de")=data_de
Session("data_inicio")=data_inicio
Session("data_ate")=data_ate
Session("data_fim")=data_fim


data_split= Split(da_pd,"/")
dia=data_split(0)
mes=data_split(1)
ano=data_split(2)


dia=dia*1

mes=mes*1
hora=hora*1
min=min*1

if dia<10 then
dia="0"&dia
end if
if mes<10 then
mes="0"&mes
end if
if hora<10 then
hora="0"&hora
end if
if min<10 then
min="0"&min
end if
da_show=dia&"/"&mes&"/"&ano


if da_atendido="" or isnull(da_atendido) then
 da_atendido_show=""
else
	data_split= Split(da_atendido,"/")
	dia=data_split(0)
	mes=data_split(1)
	ano=data_split(2)
	
	
	dia=dia*1
	
	mes=mes*1
	hora=hora*1
	min=min*1
	
	if dia<10 then
	dia="0"&dia
	end if
	if mes<10 then
	mes="0"&mes
	end if
	if hora<10 then
	hora="0"&hora
	end if
	if min<10 then
	min="0"&min
	end if
	da_atendido_show=dia&"/"&mes&"/"&ano
end if
%>
              <tr class="<%=cor%>">
                <td width="20"><input name="pedido" type="checkbox" class="borda" id="nota_fiscal" value="<%=optobr%>"></td>
                <td width="60" align="center">&nbsp;
                  <%response.Write(seq_pd)%>
                  <div align="center"></div>
                  <div align="left"></div></td>
                <td width="90" align="center">
                  <%response.Write(da_show)%>
                </td>
                <td width="100"><%response.Write(no_projeto)%>
                  </td>
                <td width="145" align="center"><%response.Write(no_unidade)%></td>
                <td width="50" align="center"><%response.Write(no_curso)%></td>
                <td width="140" align="center"><%response.Write(no_etapa)%></td>
                <td width="50" align="center"><%response.Write(turma)%></td>
                <td width="165" align="center"><%response.Write(no_registrador)%>
                  </td>
                <td width="90"><div align="center">
                  <%response.Write(st_movim)%>
                </div></td>
                <td width="91"><div align="center">
                  <%response.Write(da_atendido_show)%>
                </div></td>
              </tr>
              <%check = check+1
RS.Movenext
WEND 

END IF%>   
              <tr class="<%=cor%>">
                <td colspan="11"><div align="center"> </div>
                  <div align="left"></div>
                  <div align="left"></div>
                  <div align="left">
                    <hr width="1000">
                  </div></td>
              </tr>
           
              <tr class="<%=cor%>">
                <td colspan="11" align="center" valign="top"><table width="1000" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="25%" align="center"><%if trava="n" then%>
                        <input name="botao_prosseguir" type="submit" class="botao_prosseguir" onClick="MM_goToURL('parent','inclui.asp');return document.MM_returnValue" value="Novo Pedido">
                      <% end if%></td>
                    <td width="25%" align="center"><%if trava="n" then%>
                        <input name="submit" type="submit" class="botao_prosseguir" value="Alterar">
                      <% end if%></td>
                    <td width="25%" align="center"><%if trava="n" then%>
                        <input name="submit" type="submit" class="botao_excluir" value="Cancelar">
                      <% end if%></td>
                    <td width="25%" align="center"><%if trava="n" then%>
                        <input name="botao_imprimir" type="submit" class="botao_cancelar" value="Imprimir">
                      <% end if%></td>
                  </tr>
                </table></td>
              </tr>

            </table></td>
  </tr></table>            </form></td></tr>
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