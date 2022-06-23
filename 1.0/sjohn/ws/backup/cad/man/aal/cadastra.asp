<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<%
opt=request.QueryString("opt")

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR		
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0	

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1

campo="CO_Matricula"		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT Max(CO_Matricula) AS COD FROM TB_Alunos"
		RS.Open SQL, CON1
'WITH (READPAST)				
codigo = RS("COD")
codigo=codigo*1
codigo=codigo+1


nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo")
ano_letivo_real = ano_letivo
sistema_local=session("sistema_local")
ori = request.QueryString("or")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
cod= request.QueryString("cod")
vinc_erro= request.QueryString("v")


if ori="s" then
nome_aluno=Session("nome_cadastrar")
else
nome_aluno=""
end if
Call LimpaVetor2
 call navegacao (CON,chave,nivel)
navega=Session("caminho")

			ano = DatePart("yyyy", now)
			mes = DatePart("m", now) 
			dia = DatePart("d", now) 
da_cadastro =dia&"/"& mes &"/"& ano
%>

<html>
<head>
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../../../js/mm_menu.js"></script>
<script type="text/javascript" src="../../../../js/global.js"></script>
<script language="JavaScript">
 window.history.forward(1);
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

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
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

	
function checksubmit()
{
  if (document.inclusao.nome.value == "")
  {    alert("Por favor, digite um nome para o aluno!")
    document.inclusao.nome.focus()
    return false
  }
  return true

}						   								   								   								   								   								   								   								   
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</head>

<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="document.inclusao.nome.focus()">
<%call cabecalho(nivel)
%>
<table width="1002" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
            
    <td width="1000" height="10" colspan="5" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
    </td>
  </tr> 
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,411,0,0) %>
    </td>
  </tr>			  
        <form action="bd.asp?opt=i" method="post" name="inclusao" id="inclusao" onSubmit="return checksubmit()">
          <tr>
      <td valign="top"> 
        <table width="1000" border="0" align="right" cellspacing="0"  class="tb_corpo">
          <tr> 
            <td width="841" class="tb_tit"
>Dados Pessoais</td>
            <td width="11" class="tb_tit"
> </td>
            <td width="149" class="tb_tit"
></td>
          </tr>
          <tr> 
            <td colspan="3"><table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Matr&iacute;cula</font></div></td>
                  <td width="13" class="tb_corpo"
> <div align="left">:</div></td>
                  <td width="217" height="10" class="tb_corpo"
><font class="form_corpo"> 
                    <input name="codigo" type="hidden" class="borda" id="codigo" size="50" value="<%response.Write(codigo)%>">
                    <font class="form_dado_texto"> 
                    <%response.Write(codigo)%>
                    </font> </font></td>
                  <td width="140" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">Nome:</font></div></td>
                  <td width="19" height="10"><div align="center">:</div></td>
                  <td width="286" height="10"><font class="form_corpo"><font class="form_corpo"> 
                    <input name="nome" type="text" class="borda" id="nome" value="<%response.Write(nome_aluno)%>" size="50" maxlength="50">
                    </font></font></td>
                  <td width="11" height="10">&nbsp;</td>
                  <td width="149" height="10">&nbsp;</td>
                </tr>
              </table></td>
          </tr>
        </table></td></tr>
		          <tr> 
            <td colspan="3"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr class="tb_corpo"> 
                  <td colspan="3"><hr></td>
                </tr>
                <tr> 
                  <td width="33%"><div align="center"> 
                      <input type="button" name="Submit2" value="Voltar" class="borda_bot3" onClick="MM_goToURL('parent','index.asp?nvg=WS-CA-MA-AAL')">
                    </div></td>
                  <td width="34%">&nbsp;</td>
                  <td width="33%"> <div align="center"> 
                      <input type="submit" name="Submit" value="Confirmar" class="borda_bot">
                    </div></td>
                </tr>
              </table></td>
          </tr>
</form>
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.gif" width="1000" height="40"></td>
  </tr>
</table>
<div id="bd_familiar"> </div>
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