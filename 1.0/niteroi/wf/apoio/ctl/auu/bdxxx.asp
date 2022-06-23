<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<% 
opt= request.QueryString("opt")
nivel=4
permissao = session("permissao") 
ano_letivo_wf = session("ano_letivo_wf") 
ano_letivo_wf_real = ano_letivo_wf
sistema_local=session("sistema_local")
ori = request.QueryString("ori")
chave=session("chave")
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo_wf
cod= request.QueryString("cod_cons")	

z = request.QueryString("z")
erro = request.QueryString("e")
vindo = request.QueryString("vd")
obr = request.QueryString("o")



		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")		

	
		Set CON_WF = Server.CreateObject("ADODB.Connection") 
		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_WF.Open ABRIR_WF
		
		
  		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Usuario where CO_Usuario = "& cod
		RS.Open SQL, CON_WF
		
		
codigo = RS("CO_Usuario")
nome_prof = RS("NO_Usuario")
email = RS("TX_EMail_Usuario")
cadastro = RS("DA_Cadastro")
data_acesso = RS("DA_Ult_Acesso")
hora_acesso = RS("HO_ult_Acesso")
num_acesso = RS("NU_Acesso")
situacao = RS("ST_Usuario")
if situacao="L" then
situacao= "Liberado"
label="Bloqueia usuário"
st="Blq"
st2="B"
else
situacao= "Bloqueado"
label="Desbloqueia usuário"
st="Lib"
st2="L"
end if
		
		
if opt="bl" then

			strSQL3= "UPDATE TB_Usuario SET ST_Usuario= '"&st2&"' WHERE CO_Usuario = "& cod
			set tabela3 = CON_WF.Execute (strSQL3)


			call GravaLog (chave,cod&st)

response.Redirect("altera.asp?cod_cons="&cod&"&opt=ok&nvg="&chave)
elseif opt="rs" then

			strSQL3= "UPDATE TB_Usuario SET Senha= '"&cod&"' AND ST_Usuario= 'L' WHERE CO_Usuario = "& cod
			set tabela3 = CON_WF.Execute (strSQL3)


			call GravaLog (chave,cod&"Sen")

response.Redirect("altera.asp?cod_cons="&cod&"&opt=ok2&nvg="&chave)
else		
		
		
		
		
			



if cadastro="" or isnull(cadastro) then
else
vetor_cadastro = Split(cadastro,"/")  
dia_c = vetor_cadastro(0)
mes_c = vetor_cadastro(1)
ano_c = vetor_cadastro(2)

if dia_c<10 then 
dia_c = "0"&dia_c
end if

if mes_c<10 then
mes_c = "0"&mes_c
end if


cadastro = dia_c&"/"&mes_c&"/"&ano_c
end if
if data_acesso="" or isnull(data_acesso) then
else
data_a = Split(data_acesso,"/")  
dia_a = data_a(0)
mes_a = data_a(1)
ano_a = data_a(2)

if dia_a<10 then 
dia_a = "0"&dia_a
end if

if mes_a<10 then
mes_a = "0"&mes_a
end if

ult_data = dia_a&"/"&mes_a&"/"&ano_a
end if
if hora_acesso="" or isnull(hora_acesso) then
else
hora = Split(hora_acesso,":")  
h = hora(0)
m = hora(1)

if m<10 then
m = "0"&m
end if

ult_hora=h&":"&m
end if	

Call LimpaVetor3

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

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</head>

<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<%call cabecalho(nivel)
%>
<table width="1002" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
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
      <%call mensagens(nivel,22,1,0)
	  
	   %>
    </td>
  </tr>			  
        <form action="bd.asp?opt=rs&cod_cons=<%=codigo%>" method="post" name="busca" id="busca">
          <tr>
      <td valign="top"> 
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
          <tr> 
            <td width="841" class="tb_tit"
>Dados do Usuário</td>
          </tr>
          <tr> 
            <td height="10"> <font class="form_dado_texto">&nbsp; </font> <table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="17%" height="10"><font class="form_dado_texto">Usu&aacute;rio</font></td>
                  <td width="2%"><div align="center">:</div></td>
                  <td width="26%" height="10"><font class="form_dado_texto"> 
                    <input name="cod" type="hidden" value="<%=cod%>">
                    <%response.Write(codigo)%>
                    </font></td>
                  <td height="10"><font class="form_dado_texto">Nome:</font></td>
                  <td><div align="center">:</div></td>
                  <td width="36%" height="10"><font class="form_dado_texto"> 
                    <%response.Write(nome_prof)%>
                    </font></td>
                </tr>
                <tr> 
                  <td height="10"> <div align="left"><font class="form_dado_texto"> 
                      Email</font></div></td>
                  <td><div align="center">:</div></td>
                  <td height="10"><font class="form_dado_texto"> 
                    <%response.Write(email)%>
                    &nbsp; </font></td>
                  <td width="17%" height="10"> <div align="left"><font class="form_dado_texto"> 
                      Data de Cadastro</font></div></td>
                  <td width="2%"><div align="center">:</div></td>
                  <td height="10"><font class="form_dado_texto"> 
                    <%response.Write(cadastro)%>
                    </font></td>
                </tr>
                <tr> 
                  <td height="10"> <div align="left"><font class="form_dado_texto"> 
                      &Uacute;ltimo Acesso</font></div></td>
                  <td><div align="center">:</div></td>
                  <td height="10"><font class="form_dado_texto"> 
                    <%
					  response.Write(ult_data&", "&ult_hora)%>
                    </font></td>
                  <td height="10"> <div align="left"><font class="form_dado_texto"> 
                      N&uacute;mero de Acessos</font></div></td>
                  <td><div align="center">:</div></td>
                  <td height="10"><font class="form_dado_texto"> 
                    <%response.Write(num_acesso)%>
                    </font></td>
                </tr>
                <tr> 
                  <td height="10"> <div align="left"><font class="form_dado_texto"> 
                      Situa&ccedil;&atilde;o</font></div></td>
                  <td><div align="center">:</div></td>
                  <td height="10"> <font class="form_dado_texto"> 
                    <%response.Write(situacao)%>
                    </font> </td>
                  <td height="10"> <div align="left"><font class="form_dado_texto"> 
                      </font></div></td>
                  <td><div align="center"></div></td>
                  <td height="10"><font class="form_dado_texto">&nbsp; </font></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td><hr></td>
          </tr>
          <tr>
            <td><table width="100%" border="0" align="center" cellspacing="0">
                <tr> 
                  <td width="33%"> 
                    <div align="center"> 
                      <input name="alterar" type="submit" class="botao_cancelar" id="alterar" onClick="MM_goToURL('parent','altera.asp?cod_cons=<%=cod%>&nvg=<%=chave%>');return document.MM_returnValue" value="Cancelar">
                    </div></td>
                  <td width="34%">&nbsp;</td>
                  <td width="33%">
<div align="center"> 
                      <input name="Submit" type="submit" class="botao_prosseguir" value="Confirmar">
                    </div></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td class="tb_tit"
>&nbsp;</td>
          </tr>
        </table></td></tr>
</form>
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

</body>

</html>
<%end if%>
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