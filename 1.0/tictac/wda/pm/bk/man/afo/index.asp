<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->

<%
session("nvg")=""
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
if opt="" or isnull("opt") then
	opt="sel"
else
	if opt="ok" then
		cod_cons = request.QueryString("cod_cons")
		co_usr_prof = request.QueryString("co_usr_prof")
		tx_login=request.QueryString("tx_login")
	end if
end if
		
		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON9 = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_ax & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON9.Open ABRIR		

	call VerificaAcesso (CON,chave,nivel)
	autoriza=Session("autoriza")

	call navegacao (CON,chave,nivel)
	navega=Session("caminho")

Call LimpaVetor
	%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../../../../js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
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

<%if opt="list" or opt="listall" then
	onLoad=""
else
	onLoad="onLoad=""MM_callJS('document.busca.busca1.focus()')"""
end if
%>
//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" <%response.Write(onLoad)%>>
<%call cabecalho(nivel)%>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr>                    
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> </td>
  </tr>
<%if opt="err2" then%>	
	 <tr>                   
    <td height="10"> 
      <%
		if ori=01 then
	  	call mensagens(4,801,1,"N") 
	  elseif ori=02 then
	  	call mensagens(4,801,1,"C") 		
	  end if%>
    </td>				  				  
 </tr>
<%end if%>  
<%if opt="ok" then%>
  <tr>                   
    <td height="10"> 
      <%
	  if autoriza="no" then
	  	call mensagens(4,9700,1,0) 	  
	  else
	  	call mensagens(4,802,2,0) 
	  end if%>
    </td>
  <tr>                   
    <td height="10"> 
      <%
	  if autoriza="no" then
	  	call mensagens(4,9700,1,0) 	 	  
	  elseif autoriza="1" then
	  	call mensagens(4,800,0,"N") 
	  else
	  	call mensagens(4,800,0,"A") 
	  end if%>
    </td>
                  </tr>				  
<%elseif opt="sel" or opt="ok" or opt="err2" then%>
  <tr>                   
    <td height="10">
      <%
	  if autoriza="no" then
	  	call mensagens(4,9700,1,0) 	  
	  elseif autoriza="1" then
	  	call mensagens(4,800,0,"N") 
	  else
	  	call mensagens(4,800,0,"A") 
	  end if%>
    </td>
<%end if%>

                  </tr>
				  				  				  

  <tr> 
    <td valign="top">

<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="tb_corpo">
        <tr> 
            <td> 
              <%	  if autoriza="no" then			
		else
		
		if opt="sel" or opt="ok" or opt="err2" then%>
        <form action="index.asp?opt=list&nvg=<%=nvg%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
          <table width="1000" border="0" cellspacing="0">
            <tr> 

                  <td height="70" valign="top"> 
                    <table width="1000" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo"
>
                      <tr class="tb_tit"
> 
                    <td colspan="8"><font class="tb_tit">Preencha 
                      um dos campos abaixo</font></td>
                  </tr>
                  <tr> 
                        <td height="30"> 
                          <div align="right"><font class="form_dado_texto">C&oacute;digo 
                            : </font></div></td>
                        <td height="30">
                          <input name="busca1" type="text" class="textInput" id="busca1" size="12"></td>
                        <td width="100" height="30"> 
                          <div align="right">
                            <p><font class="form_dado_texto">Nome :</font></p>
                          </div></td>
                        <td height="30" colspan="4"><font size="2" face="Arial, Helvetica, sans-serif"> 
                          <input name="busca2" type="text" class="textInput" id="busca2" size="55" maxlength="50">
                          </font></td>
                        <td><font size="2" face="Arial, Helvetica, sans-serif"> 
                          <input name="Submit" type="submit" class="botao_prosseguir" id="Submit" value="Procurar">
                          </font> </td>
                  </tr>
                </table></td>
            </tr>
          </table>
        </form>
    </td>				  				  
                  </tr>		
        <%elseif opt="list" then
  busca1=request.form("busca1") 
  busca2=request.form("busca2")
  if busca1 ="" then
	  query = busca2
	  ori=01
  elseif busca2 ="" then
	  query = busca1 
	  ori=02
  end if 
  
teste = IsNumeric(query)
if teste = TRUE Then
  
  		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Fornecedor where CO_Fornecedor = "& query
		RS.Open SQL, CON9
		
		if RS.EOF Then
		'  response.Redirect("resultado.asp?or=02&cod_cons="&query&"&nvg="&nvg)
		%>
		  <tr>                   
			<td height="10"> 
			  <%
			  if autoriza="no" then
				call mensagens(4,9700,1,0) 	  
			  elseif autoriza="1" then
				call mensagens(4,800,0,"N") 
			  else
				call mensagens(4,800,0,"A") 
			  end if%>
			</td>
		 <tr>                   
			<td height="10"> 
			  <%
			  if ori=01 then
				call mensagens(4,801,1,"N") 
			  elseif ori=02 then
				call mensagens(4,801,1,"C") 	
			  end if%>
			</td>				  				  
		 </tr>
														  
		
		  <tr> 
			<td valign="top">
		
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="tb_corpo">
				<tr> 
					<td> 
				<form action="index.asp?opt=list&nvg=<%=nvg%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
				  <table width="1000" border="0" cellspacing="0">
					<tr> 
		
						  <td height="70" valign="top"> 
							<table width="1000" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo"
		>
							  <tr class="tb_tit"
		> 
							<td colspan="8"><font class="tb_tit">Preencha 
							  um dos campos abaixo</font></td>
						  </tr>
						  <tr> 
								<td height="30"> 
								  <div align="right"><font class="form_dado_texto">C&oacute;digo 
									: </font></div></td>
								<td height="30">
								  <input name="busca1" type="text" class="textInput" id="busca1" size="12"></td>
								<td width="100" height="30"> 
								  <div align="right">
									<p><font class="form_dado_texto">Nome :</font></p>
								  </div></td>
								<td height="30" colspan="4"><font size="2" face="Arial, Helvetica, sans-serif"> 
								  <input name="busca2" type="text" class="textInput" id="busca2" size="55" maxlength="50">
								  </font></td>
								<td><font size="2" face="Arial, Helvetica, sans-serif"> 
								  <input name="Submit" type="submit" class="botao_prosseguir" id="Submit" value="Procurar">
								  </font> </td>
						  </tr>
						</table></td>
					</tr>
				  </table>
				</form>
			</td>				  				  
		 </tr>
		</table>
			</td>				  				  
		 </tr>
		<%ELSE		
              response.Redirect("altera.asp?ori=01&cod_cons="&query&"&nvg="&nvg)
        END IF
ELSE

'Converte caracteres que não são válidos em uma URL e os transformamem equivalentes para URL
strProcura = Server.URLEncode(request("busca2"))
'Como nossa pesquisa será por "múltiplas palavras" (aqui você pode alterar ao seu gosto)
'é necessário trocar o sinal de (=) pelo (%) que é usado com o LIKE na string SQL
strProcura = replace(strProcura,"+"," ")
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
strProcura = replace(strProcura,"%E9","é")
strProcura = replace(strProcura,"%EA","ê")
strProcura = replace(strProcura,"%ED","í")
strProcura = replace(strProcura,"%F3","ó")
strProcura = replace(strProcura,"F4","ô")
strProcura = replace(strProcura,"F5","õ")
strProcura = replace(strProcura,"%FA","ú")
strProcura = replace(strProcura,"%FC","ü")


	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Fornecedor where NO_Fornecedor like '%"& strProcura & "%' order BY NO_Fornecedor"
	RS.Open SQL, CON9		

	if RS.EOF then
		response.redirect("index.asp?nvg="&chave&"&ori=01&opt=err2")
	else	
		WHile Not RS.EOF
		nome = RS("NO_Fornecedor")
		Valor_Vetor = nome
		
		cod_cons = RS("CO_Fornecedor")
		'Chama a function que ira incluir um valor para o vetor
		Call Incluir_Vetor
		
		RS.Movenext
		Wend
	end if		
	
	Call VisualizaValoresVetor4("F")
END IF
elseif opt="listall" then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Fornecedor Order BY NO_Fornecedor"
		RS.Open SQL, CON9
%>
              <tr> 
                <td valign="top">         
            <table width="1000" border="0" cellpadding="0" cellspacing="0">
              <tr> 
                <td valign="top"> 
                  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="tb_corpo"
>
                    <tr> 
                <td class="tb_tit"
>Lista de completa de Fornecedores</td>
              </tr>
  <tr>                   
    <td height="10"> 
      <%
	  if autoriza="no" then
	  	call mensagens(4,9700,1,0) 	  
	  else
	  	call mensagens(4,803,0,0) 
	  end if%>
    </td>
                  </tr>			  
              <tr> 
                <td> <ul>
                    <%
WHile Not RS.EOF
	nome = RS("NO_Fornecedor")
	cod_cons = RS("CO_Fornecedor")
	ativo = RS("IN_Ativo")
	if ativo = "True" then
		Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=ativos href=altera.asp?ori=01&cod_cons="&cod_cons&"&nvg="&nvg&" >"&nome&"</a></font></li>")
	else
		Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=inativos href=altera.asp?ori=01&cod_cons="&cod_cons&"&nvg="&nvg&">"&nome&"</a></font></li>")
	end if
RS.Movenext
Wend
%>
                  </ul></td>
              </tr>
            </table>
                </td>
        </tr>
      </table>
      </div> 
      <%end if 
	  end if%>
            </td>
        </table>        
     </td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
</body>
<%if opt<>"list" and opt<>"listall" then%>
<script type="text/javascript">
<!--
  initInputHighlightScript();
//-->
<%end if%>
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