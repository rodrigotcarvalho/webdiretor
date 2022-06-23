<%On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<% 
session("nvg")=""
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
nvg=request.QueryString("nvg")
opt = request.QueryString("opt")

chave=nvg
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
if opt="" or isnull("opt") then
	opt="sel"
else
	opt=opt
	if opt="ok" then
		cod_cons = request.QueryString("cod_cons")
		co_usr_prof = request.QueryString("co_usr_prof")
	end if
end if


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
 call VerificaAcesso (CON,chave,nivel)
autoriza=Session("autoriza")

 call navegacao (CON,chave,nivel)
navega=Session("caminho")

Call LimpaVetor2

%>
<html>
<head>
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../../../js/mm_menu.js"></script>
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
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}

function MM_showHideLayers() { //v6.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
    obj.visibility=v; }
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function checksubmit()
{
  if (document.busca.busca1.value != "" && document.busca.busca2.value != "")
  {    alert("Por favor digite SOMENTE uma opção de busca!")
    document.busca.busca1.focus()
    return false
  }
    if (document.busca.busca1.value == "" && document.busca.busca2.value == "")
  {    alert("Por favor digite uma opção de busca!")
    document.busca.busca1.focus()
    return false
  }
  return true
}

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
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
</head>
<% if opt="listall" or opt="list" then%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%else %>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('document.busca.busca1.focus()')">
<%end if %>
<%call cabecalho(nivel)
%>
<table width="1000" height="685" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
            
    <td height="10" colspan="5" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
    </td>
          </tr>
        <%if opt="sel" then%>
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,300,0,0) %>
    </td>
			  </tr>			  
        <form action="index.asp?opt=list&nvg=<%=nvg%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
          <tr class="tb_tit"> 
            
      <td height="10" colspan="5">Preencha um dos campos abaixo</td>
          </tr>
          <TR>
		  
      <td height="92" valign="top"> 
        <table width="1000" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            
            <td width="77" height="10" align="center" class="tb_subtit">Matr&iacute;cula 
            </td>
            
            <td width="313" height="10" ><span class="tb_subtit">Nome</span></td>
            <td width="340" class="tb_subtit">Per&iacute;odo 
              da Entrevista</td>
            <td width="270" class="tb_subtit">Status da Entrevista</td>
          </tr>
          <tr>
            <td height="10" align="center"><font size="2" face="Arial, Helvetica, sans-serif">
              <input name="busca4" type="text" class="textInput" id="busca4" size="12">
            </font></td>
            <td height="10" ><font size="2" face="Arial, Helvetica, sans-serif">
              <input name="busca3" type="text" class="textInput" id="busca3" size="55" maxlength="50">
            </font></td>
            <td><font class="form_dado_texto">
              <select name="dia_de" id="select" class="select_style">
                <% 
							 For i =1 to 31
							 dia=dia*1
							 if dia=i then 
								if dia<10 then
								dia="0"&dia
								end if
							 %>
                <option value="<%response.Write(i)%>" selected>
                  <%response.Write(dia)%>
                </option>
                <% else
							  	i_cod=i
								if i<10 then
								
								i="0"&i
								end if
							%>
                <option value="<%response.Write(i_cod)%>">
                  <%response.Write(i)%>
                </option>
                <% end if 
							next
							%>
              </select>
              /
  <select name="mes_de" id="select2" class="select_style">
    <%mes=mes*1
								if mes="1" or mes=1 then%>
    <option value="1" selected>janeiro</option>
    <% else%>
    <option value="1">janeiro</option>
    <%end if
								if mes="2" or mes=2 then%>
    <option value="2" selected>fevereiro</option>
    <% else%>
    <option value="2">fevereiro</option>
    <%end if
								if mes="3" or mes=3 then%>
    <option value="3" selected>mar&ccedil;o</option>
    <% else%>
    <option value="3">mar&ccedil;o</option>
    <%end if
								if mes="4" or mes=4 then%>
    <option value="4" selected>abril</option>
    <% else%>
    <option value="4">abril</option>
    <%end if
								if mes="5" or mes=5 then%>
    <option value="5" selected>maio</option>
    <% else%>
    <option value="5">maio</option>
    <%end if
								if mes="6" or mes=6 then%>
    <option value="6" selected>junho</option>
    <% else%>
    <option value="6">junho</option>
    <%end if
								if mes="7" or mes=7 then%>
    <option value="7" selected>julho</option>
    <% else%>
    <option value="7">julho</option>
    <%end if%>
    <%if mes="8" or mes=8 then%>
    <option value="8" selected>agosto</option>
    <% else%>
    <option value="8">agosto</option>
    <%end if
								if mes="9" or mes=9 then%>
    <option value="9" selected>setembro</option>
    <% else%>
    <option value="9">setembro</option>
    <%end if
								if mes="10" or mes=10 then%>
    <option value="10" selected>outubro</option>
    <% else%>
    <option value="10">outubro</option>
    <%end if
								if mes="11" or mes=11 then%>
    <option value="11" selected>novembro</option>
    <% else%>
    <option value="11">novembro</option>
    <%end if
								if mes="12" or mes=12 then%>
    <option value="12" selected>dezembro</option>
    <% else%>
    <option value="12">dezembro</option>
    <%end if%>
  </select>
              /
  <%response.write(ano_letivo)%>
              at&eacute;
  <select name="dia_ate" id="select3" class="select_style">
    <option value="1">01</option>
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
    <option value="22">22</option>
    <option value="23">23</option>
    <option value="24">24</option>
    <option value="25">25</option>
    <option value="26">26</option>
    <option value="27">27</option>
    <option value="28">28</option>
    <option value="29">29</option>
    <option value="30">30</option>
    <option value="31" selected>31</option>
  </select>
              /
  <select name="mes_ate" id="select4" class="select_style">
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
  </select>
              /
  <%response.write(ano_letivo)%>
            </font></td>
            <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr class="form_dado_texto">
                <td><input name="status" type="checkbox" id="status" value="A"></td>
                <td> Atendidas </td>
                <td><input name="status" type="checkbox" id="status" value="P"></td>
                <td>Pendentes </td>
                <td><input name="status" type="checkbox" id="status" value="C"></td>
                <td>Canceladas</td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td height="10" align="center">&nbsp;</td>
            <td height="10" >&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td height="10" colspan="4" align="center" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td colspan="4"><hr></td>
                </tr>
              <tr>
                <td width="25%"><div align="center"> </div></td>
                <td width="25%"><div align="center"> </div></td>
                <td width="25%"><div align="center"> </div></td>
                <td width="25%"><div align="center">
                  <input name="SUBMIT2" type=SUBMIT class="botao_prosseguir" value="Prosseguir">
                </div></td>
              </tr>
            </table></td>
          </tr>
		  </table>
		  </td>
		  </TR>
      </form>
	   <tr> 
            
      <td > 
	  </td>
          </tr>
        <%elseif opt="list" then
  busca1=request.form("busca1") 
  busca2=request.form("busca2")
  if busca1 ="" then
  query = busca2
  mensagem=304
  elseif busca2 ="" then
  query = busca1 
mensagem=303
  end if 
  
  teste = IsNumeric(query)
  if teste = TRUE Then
  
  		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos where CO_Matricula = "& query
		RS.Open SQL, CON1
		
if RS.EOF Then
%>
            <tr> 
              
    <td height="10" colspan="5"> 
      <%call mensagens(nivel,mensagem,1,0) %>
    </td>
			   </tr>
            <tr> 
              
    <td height="10" colspan="5"> 
      <%call mensagens(nivel,300,0,0) %>
    </td>
			  </tr>
        <form action="index.asp?opt=list&nvg=<%=nvg%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
                  <tr class="tb_tit"> 
                    
      <td height="10" colspan="5">Preencha um dos campos abaixo</td>
                  </tr>
                  <tr>
      <TD height="26" valign="top"><table width="1000" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="77" height="10" align="center" class="tb_subtit">Matr&iacute;cula </td>
          <td width="313" height="10" ><span class="tb_subtit">Nome</span></td>
          <td width="340" class="tb_subtit">Per&iacute;odo 
            da Entrevista</td>
          <td width="270" class="tb_subtit">Status da Entrevista</td>
        </tr>
        <tr>
          <td height="10" align="center"><font size="2" face="Arial, Helvetica, sans-serif">
            <input name="busca5" type="text" class="textInput" id="busca5" size="12">
          </font></td>
          <td height="10" ><font size="2" face="Arial, Helvetica, sans-serif">
            <input name="busca5" type="text" class="textInput" id="busca6" size="55" maxlength="50">
          </font></td>
          <td><font class="form_dado_texto">
            <select name="dia_de2" id="dia_de" class="select_style">
              <% 
							 For i =1 to 31
							 dia=dia*1
							 if dia=i then 
								if dia<10 then
								dia="0"&dia
								end if
							 %>
              <option value="<%response.Write(i)%>" selected>
                <%response.Write(dia)%>
                </option>
              <% else
							  	i_cod=i
								if i<10 then
								
								i="0"&i
								end if
							%>
              <option value="<%response.Write(i_cod)%>">
                <%response.Write(i)%>
                </option>
              <% end if 
							next
							%>
            </select>
            /
            <select name="mes_de2" id="mes_de" class="select_style">
              <%mes=mes*1
								if mes="1" or mes=1 then%>
              <option value="1" selected>janeiro</option>
              <% else%>
              <option value="1">janeiro</option>
              <%end if
								if mes="2" or mes=2 then%>
              <option value="2" selected>fevereiro</option>
              <% else%>
              <option value="2">fevereiro</option>
              <%end if
								if mes="3" or mes=3 then%>
              <option value="3" selected>mar&ccedil;o</option>
              <% else%>
              <option value="3">mar&ccedil;o</option>
              <%end if
								if mes="4" or mes=4 then%>
              <option value="4" selected>abril</option>
              <% else%>
              <option value="4">abril</option>
              <%end if
								if mes="5" or mes=5 then%>
              <option value="5" selected>maio</option>
              <% else%>
              <option value="5">maio</option>
              <%end if
								if mes="6" or mes=6 then%>
              <option value="6" selected>junho</option>
              <% else%>
              <option value="6">junho</option>
              <%end if
								if mes="7" or mes=7 then%>
              <option value="7" selected>julho</option>
              <% else%>
              <option value="7">julho</option>
              <%end if%>
              <%if mes="8" or mes=8 then%>
              <option value="8" selected>agosto</option>
              <% else%>
              <option value="8">agosto</option>
              <%end if
								if mes="9" or mes=9 then%>
              <option value="9" selected>setembro</option>
              <% else%>
              <option value="9">setembro</option>
              <%end if
								if mes="10" or mes=10 then%>
              <option value="10" selected>outubro</option>
              <% else%>
              <option value="10">outubro</option>
              <%end if
								if mes="11" or mes=11 then%>
              <option value="11" selected>novembro</option>
              <% else%>
              <option value="11">novembro</option>
              <%end if
								if mes="12" or mes=12 then%>
              <option value="12" selected>dezembro</option>
              <% else%>
              <option value="12">dezembro</option>
              <%end if%>
            </select>
            /
            <%response.write(ano_letivo)%>
            at&eacute;
            <select name="dia_ate2" id="dia_ate" class="select_style">
              <option value="1">01</option>
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
              <option value="22">22</option>
              <option value="23">23</option>
              <option value="24">24</option>
              <option value="25">25</option>
              <option value="26">26</option>
              <option value="27">27</option>
              <option value="28">28</option>
              <option value="29">29</option>
              <option value="30">30</option>
              <option value="31" selected>31</option>
            </select>
            /
            <select name="mes_ate2" id="mes_ate" class="select_style">
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
            </select>
            /
            <%response.write(ano_letivo)%>
          </font></td>
          <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr class="form_dado_texto">
              <td><input name="status2" type="checkbox" id="status2" value="A"></td>
              <td> Atendidas </td>
              <td><input name="status2" type="checkbox" id="status2" value="P"></td>
              <td>Pendentes </td>
              <td><input name="status2" type="checkbox" id="status2" value="C"></td>
              <td>Canceladas</td>
            </tr>
          </table></td>
        </tr>
        <tr>
          <td height="10" align="center">&nbsp;</td>
          <td height="10" >&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td height="10" colspan="4" align="center" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td colspan="4"><hr></td>
            </tr>
            <tr>
              <td width="25%"><div align="center"> </div></td>
              <td width="25%"><div align="center"> </div></td>
              <td width="25%"><div align="center"> </div></td>
              <td width="25%"><div align="center">
                <input name="SUBMIT" type=SUBMIT class="botao_prosseguir" value="Prosseguir">
              </div></td>
            </tr>
          </table></td>
        </tr>
      </table> 
        
      </TD>
    </tr>
        </form>
 <tr>             
      <td > 
	  </td>
          </tr>
        <%ELSE		
  response.Redirect("docs.asp?or=01&cod_cons="&query&"&pagina=1&v=n")
END IF
  ELSE

'Converte caracteres que não são válidos em uma URL e os transformamem equivalentes para URL
strProcura = Server.URLEncode(request("busca2"))
'Como nossa pesquisa será por "múltiplas palavras" (aqui você pode alterar ao seu gosto)
'é necessário trocar o sinal de (=) pelo (%) que é usado com o LIKE na string SQL
strProcura = replace(strProcura,"+"," ")
strProcura = replace(strProcura,"%27","´")
strProcura = replace(strProcura,"%27","'")
	
strProcura = replace(strProcura,"%C0,","À")
strProcura = replace(strProcura,"%C1","Á")
strProcura = replace(strProcura,"%C2","Â")
strProcura = replace(strProcura,"%C3","Ã")
strProcura = replace(strProcura,"%C9","É")
strProcura = replace(strProcura,"%CA","Ê")
strProcura = replace(strProcura,"%CD","Í")
strProcura = replace(strProcura,"%D3","Ó")
strProcura = replace(strProcura,"%D4","Ô")
strProcura = replace(strProcura,"%D5","Õ")
strProcura = replace(strProcura,"%DA","Ú")
strProcura = replace(strProcura,"%DC","Ü")

strProcura = replace(strProcura,"%E1","à")
strProcura = replace(strProcura,"%E1","á")
strProcura = replace(strProcura,"%E2","â")
strProcura = replace(strProcura,"%E3","ã")
strProcura = replace(strProcura,"%E7","ç")
strProcura = replace(strProcura,"%E9","é")
strProcura = replace(strProcura,"%EA","ê")
strProcura = replace(strProcura,"%ED","í")
strProcura = replace(strProcura,"%F3","ó")
strProcura = replace(strProcura,"F4","ô")
strProcura = replace(strProcura,"F5","õ")
strProcura = replace(strProcura,"%FA","ú")
strProcura = replace(strProcura,"%FC","ü")


		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos where NO_Aluno like '%"& strProcura & "%' order BY NO_Aluno"
		RS.Open SQL, CON1		


WHile Not RS.EOF
cod = RS("CO_Matricula")
nome = RS("NO_Aluno")
Valor_Vetor = cod

cod = RS("CO_Matricula")
'Chama a function que ira incluir um valor para o vetor
Call Incluir_Vetor2

RS.Movenext
Wend
	

Call VisualizaValoresVetor2
END IF
elseif opt="listall" then

	NO_Aluno = request.Form("NO_Aluno")


		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos Order BY NO_Aluno"
		RS.Open SQL, CON1
%>
          <tr> 
            
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,1,0,0) %>
    </td>
          </tr>
                <tr class="tb_corpo"> 
                  
    <td height="10" colspan="5" class="tb_tit">Lista de completa de Alunos</td>
                </tr>
                <tr> 
                  
    <td colspan="5" valign="top"> 
      <ul>
        <%
WHile Not RS.EOF
nome = RS("NO_Aluno")
cod = RS("CO_Matricula")
ativo = RS("IN_Ativo_Escola")
if ativo = "True" then
Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=ativos href=docs.asp?or=01&cod_cons="&cod&"&pagina=1&v=n>"&nome&"</a></font></li>")
else
Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=inativos href=docs.asp?or=01&cod_cons="&cod&"&pagina=1&v=n>"&nome&"</a></font></li>")
end if
RS.Movenext
Wend
end if 
%>
      </ul></td>
                </tr>
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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