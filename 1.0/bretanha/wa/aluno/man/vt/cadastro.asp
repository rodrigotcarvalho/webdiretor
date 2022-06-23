<%On Error Resume Next%>
<!--#include file="../inc/funcoes.asp"-->



<!--#include file="../inc/caminhos.asp"-->
<% 
opt= request.QueryString("opt")
ano_letivo= request.form("ano_letivo")
session("ano_letivo") = ano_letivo
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.addHeader "pragma","no-cache"
Response.addHeader "cache-control","private"
Response.CacheControl = "no-cache"
Response.Buffer = True


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		

id = " > Consultar Cadastro"
Call LimpaVetor2

%>
<html>
<head>
<title>Web Acad&ecirc;mico</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../js/mm_menu.js"></script>
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
//-->
</script>
</head>
<% if opt="listall" or opt="list" then%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%else %>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('document.busca.busca1.focus()')">
<%end if %>
<%call cabecalho(nivel)
%>
<table width="1000" border="0" align="center" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td><div align="center"><table width="1000" border="0">
  <tr> 
    <td><font color="#FFFF33" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="../inicio.asp" class="caminho">Web 
      Acad&ecirc;mico</a> 
      <%response.Write(origem&id)%>
 </font></td>
  </tr>
</table>
<br>
<%if opt="sel" then%>
<form action="cadastro.asp?opt=list&or=01" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
  <table width="1000" border="0" cellspacing="0">
    <tr>
      <td width="220" valign="top"> 
      <%call mensagens(1000,0,0) %>
      </td>
      <td width="788" height="70" valign="top"> 
        <table width="770" border="0" align="right" cellspacing="0" class="tb_corpo"
>
          <tr class="tb_tit"
> 
            <td colspan="8"><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Dados 
              Pessoais do Aluno</strong></font></td>
          </tr>
          <tr> 
            <td width="88" height="30"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Matr&iacute;cula: 
                </strong></font><font color="#CC9900"><strong></strong></font></div></td>
            <td width="60" height="30"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font><font size="2" face="Arial, Helvetica, sans-serif">
              <input name="busca1" type="text" class="textInput" id="busca12" size="5">
              </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font></td>
            <td width="35" height="30"><div align="right"><font color="#CC9900"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Nome: 
                </font></strong></font></div></td>
            <td width="410" height="30" colspan="4"><font size="2" face="Arial, Helvetica, sans-serif">
              <input name="busca2" type="text" class="textInput" id="busca2" size="55" maxlength="50">
              </font></td>
            <td width="161"><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="Submit" type="submit" class="borda_bot" id="Submit" value="Procurar">
              </font> </td>
          </tr>
        </table></td>
    </tr>
  </table>
</form>

<%elseif opt="list" then
  busca1=request.form("busca1") 
  busca2=request.form("busca2")
  if busca1 ="" then
  query = busca2
  elseif busca2 ="" then
  query = busca1 

  end if 
  
  teste = IsNumeric(query)
  if teste = TRUE Then
  
  		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos where CO_Matricula = "& query
		RS.Open SQL, CON
		
if RS.EOF Then
%>

<form action="cadastro.asp?opt=list&or=01" method="post" name="busca" id="busca" onSubmit="return checksubmit()"> 
  <table width="1000" border="0" cellspacing="0">
    <tr> 
      <td width="220"> 
        <%call mensagens(1003,1,0) %>
        <br>
      </td>
      <td width="788" rowspan="2" valign="top"> <table width="770" border="0" align="right" cellspacing="0" class="tb_corpo"
>
          <tr class="tb_tit"
> 
            <td colspan="8"><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Dados 
              Pessoais do Aluno</strong></font></td>
          </tr>
          <tr> 
            <td width="88" height="30"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Matr&iacute;cula: 
                </strong></font><font color="#CC9900"><strong></strong></font></div></td>
            <td width="60" height="30"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="busca1" type="text" class="textInput" id="busca1" size="5">
              </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font></td>
            <td width="35" height="30"><div align="right"><font color="#CC9900"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Nome: 
                </font></strong></font></div></td>
            <td width="410" height="30" colspan="4"><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="busca2" type="text" class="textInput" id="busca2" size="55" maxlength="50">
              </font></td>
            <td width="161"><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="Submit3" type="submit" class="borda_bot" id="Submit2" value="Procurar">
              </font> </td>
          </tr>
        </table></td>
    </tr>
    <tr>
      <td>
        <%call mensagens(0,0,0) %>
      </td>
    </tr>
  </table>
</form>
<%ELSE		
  response.Redirect("altera.asp?or=01&cod="&query&"")
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
		SQL = "SELECT * FROM TB_Alunos where NO_Aluno like '%"& strProcura & "%' order BY NO_Aluno"
		RS.Open SQL, CON		

WHile Not RS.EOF
nome = RS("NO_Aluno")
Valor_Vetor = nome

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
		RS.Open SQL, CON
%>
<table width="1000" border="0" cellspacing="0">
  <tr> 
    <td width="220" valign="top"> 
      <%call mensagens(1,0,0) %>
    </td>
    <td width="788" valign="top"> 
        <table width="770" border="0" align="right" cellspacing="0" class="tb_corpo"
>
        <tr class="tb_corpo"
> 
            
          <td class="tb_tit"
><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Lista 
            de completa de Alunos</strong></font></td>
          </tr>
          <tr> 
            <td> <ul>
              <%
WHile Not RS.EOF
nome = RS("NO_Aluno")
cod = RS("CO_Matricula")
ativo = RS("IN_Ativo_Escola")
if ativo = "True" then
Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=ativos href=altera.asp?or=01&cod="&cod&" >"&nome&"</a></font></li>")
else
Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=inativos href=altera.asp?or=01&cod="&cod&">"&nome&"</a></font></li>")
end if
RS.Movenext
Wend
%></ul>
            </td>
          </tr>
        </table>
      </td>
  </tr>
</table>
</div>
<%end if %>
<p>&nbsp;</p><table width="1000" border="0" cellspacing="0">
  <tr> 
    <td width="238">&nbsp;</td>
    <td width="770" class="tb_voltar"
><font color="#669999" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="../alunos.asp" class="voltar1">&lt; 
      Voltar para o menu Alunos</a></strong></font></td>
  </tr>
</table>
<p align="center">&nbsp;</p></div></td>
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
pasta=arPath(seleciona)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>