<!--#include file="inc/caminhos.asp"-->

<%
lg = session("lg")
senha= session("senha")


	Set conexao_ctl = Server.CreateObject("ADODB.Connection") 
	ABRIR_ctl = "DBQ="& CAMINHOctl & ";Driver={Microsoft Access Driver (*.mdb)}"
	conexao_ctl.Open ABRIR_ctl
	
	consulta_ctl = "select * from TB_Controle"
	set tabela_ctl = conexao_ctl.Execute (consulta_ctl)

controle=tabela_ctl("CO_controle")

if controle= "D" then
response.Redirect("manutencao.asp")
end if

opt = request.QueryString("opt")
if opt="" or isnull(opt) then
	opt=999999
end if

select case opt
case 00
	msg="conex&atilde;o foi encerrada por estar inativa a mais de 10 minutos. Digite novamente seu login e senha para ter acesso ao Sistema."
	tipo="e"
case 01
	msg="O campo Usu�rio � obrigat�rio!"
	tipo="e"
	lg = ""
	pas= session("senha")
	ti = session("ti")
	mail_prim=session("mail_prim")
case 02
	msg="O campo Senha � obrigat�rio!"
	tipo="e"
	lg = session("lg")
	pas= ""
	ti = session("ti")
	mail_prim=session("mail_prim")
case 03
	msg="Digitar o c�digo da figura � obrigat�rio!"
	tipo="e"
	lg = session("lg")
	pas= session("senha")
	pas1= session("pas1")
	pas2= session("pas2")
	ti = ""
	mail_prim=session("mail_prim")
case 04
	msg="O Usu�rio "&lg&" n�o existe!"
	tipo="e"
	lg = ""
	pas= session("senha")
	ti = session("ti")
	mail_prim=session("mail_prim")
case 05
	msg="Usu�rio n�o autorizado."
	tipo="e"
	lg = ""
	pas= session("senha")
	ti = session("ti")
	mail_prim=session("mail_prim")
case 06
	msg="Senha Incorreta!"
	tipo="e"
	lg = session("lg")
	pas= ""
	ti = session("ti")
	mail_prim=session("mail_prim")
case 07
	msg="Tempo de digita��o do c�digo da figura excedido. Tente novamente."
	tipo="e"
	lg = session("lg")
	pas= session("senha")
	mail_prim=session("mail_prim")
	ti = ""
case 08
	msg="C�digo da figura Incorreto!"
	tipo="e"
	lg = session("lg")
	pas= session("senha")
	pas1= session("pas1")
	pas2= session("pas2")
	mail_prim=session("mail_prim")
ti = ""
case 09
	tipo="prim"
	lg = session("lg")
	pas= session("senha")
	mail_prim=session("mail_prim")
case 10
	msg="O campo Nova Senha deve ser preenchido!"
	tipo="e"
	lg = session("lg")
	pas= session("senha")
	mail_prim=session("mail_prim")
case 11
	msg="O campo Confirme a Senha deve ser preenchido!"
	tipo="e"
	lg = session("lg")
	pas= session("senha")
	pas1= session("pas1")
	mail_prim=session("mail_prim")
case 12
	msg="a senha digitada no campo Nova Senha n�o pode ser igual a Senha j� cadastrada."
	tipo="e"
	lg = session("lg")
	pas= session("senha")
	pas1= ""
	pas2= ""
	mail_prim=session("mail_prim")
case 13
	msg="a senha digitada no campo Confirme a Senha n�o � igual a digitada no campo Nova Senha. Para que a senha seja alterada � necess�rio que estas sejam iguais."
	tipo="e"
	lg = session("lg")
	pas= session("senha")
	pas1= ""
	pas2= ""
	mail_prim=session("mail_prim")
case 14
	msg="Usu�rio j� cadastrado!"
	tipo="e"
	lg = session("lg")
	pas= session("senha")
	pas1= ""
	pas2= ""
	mail_prim=session("mail_prim")
case 999998
'aparece quando o bot�o novo c�digo � clicado
	lg = session("valor1")
	pas= session("valor2")
	pas1= session("valor3")
	pas2= session("valor4")
	mail_prim=session("valor5")
case else
	session.Contents.RemoveAll()
end select
'response.Write("6>>>"&senha)
'RESPONSE.END
if tipo="e" then
cor = "#FF0000"
end if

caracter = Array("1","2","3","4","5","6","7","8","9","q","w","e","r","t","y","i","p","a","s","d","f","g","h","j","k","l","z","x","c","v","b","n","m")

	Randomize
	For i = 1 to 5
		gerarNumeros = gerarNumeros &"-"& caracter(ubound(caracter) * Rnd)
	Next

codigo=split(gerarNumeros,"-")
codigo1=codigo(1)
codigo2=codigo(2)
codigo3=codigo(3)
codigo4=codigo(4)
codigo5=codigo(5)

codigo_seguranca = codigo1&codigo2&codigo3&codigo4&codigo5
'session("codigo_seguranca") = codigo1&codigo2&codigo3&codigo4&codigo5

'imagem = Array("10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25")

imagem = Array("12","15","17","18","19","21","24")


	Randomize
		gerarfundo = gerarfundo & imagem(ubound(imagem) * Rnd)



imagem_seguranca = gerarfundo

session("codigo_seguranca")=codigo_seguranca
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Web Fam&iacute;lia</title>
<style type="text/css">
<!--
body {
	background-image: url(img/grade-fundo.gif);
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<link href="estilo.css" rel="stylesheet" type="text/css" />
<script type="text/JavaScript">
<!--
function FocusNoForm() 
{ 
//formlogin.nome.value="testes"; 
<%if opt=02 or opt=06 then%>
login.senha.focus(); 
<%elseif opt=03 or opt=07 or opt=08  then%>
login.texto_imagem.focus(); 
<%elseif opt=09 or opt=10 or opt=12 then%>
login.pas1.focus(); 
<%elseif opt=11 then%>
login.pas2.focus(); 
<%else%>
login.login.focus(); 
<%end if%>
} 
function submitcodigo()  
{
   var f=document.forms[0]; 
      f.submit(); 
}
function submitlogin()  
{
   var f=document.forms[1]; 
   
      f.submit(); 
}
//function redirect(){
//setTimeout("submitcodigo()",5000);
//}

function checksubmit()
{

//if (document.login.senha.value == document.login.pas1.value)
//  {    alert("A Nova Senha deve ser diferente da atual")
 //   document.login.pas1.focus()
//    return false
//  }

pas1 = document.login.pas1.value;
     if (pas1.length < 6)
  {    alert("O valor do campo Nova Senha deve possuir pelo menos 6 caracteres")
    document.login.pas1.focus()
    return false
  }

pas2 = document.login.pas2.value;
     if (pas2.length < 6)
  {    alert("O valor do campo Confirme a Senha deve possuir pelo menos 6 caracteres")
    document.login.pas2.focus()
    return false
  }
 if( document.login.email.length==0 || ((document.login.email.length != 0) && (document.login.email.value.indexOf('@')==-1 || document.login.email.value.indexOf('.')==-1)))
	{
		alert( "Preencha campo E-MAIL corretamente!" );
			document.login.email.focus();
				return false;
	} 
if(pas2 != pas1){
alert("a senha digitada no campo confirma��o n�o � igual a digitada no campo senha. Para que a senha seja alterada � necess�rio que estas sejam iguais");
document.login.pas1.focus()
return false
	}
	
senha = document.login.senha.value;
if(pas1 == senha){
alert("a nova senha n�o pode ser igual a senha atual.");
document.login.pas1.focus()
return false
	}
	else {
var f=document.forms[1]; 
f.submit(); 
}
  
//  return true

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
									alert("Esse browser n�o tem recursos para uso do Ajax");
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
						
						
	 function GuardaLogin(login)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../global/guarda_valores_digitados.asp?opt=valor1", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor1=" + login);
		}
	
	
	 function GuardaSenha(senha)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../global/guarda_valores_digitados.asp?opt=valor2", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor2=" + senha);
		}
	 function GuardaNovaSenha(nvsenha)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../global/guarda_valores_digitados.asp?opt=valor3", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor3=" + nvsenha);
		}
	 function GuardaConfirmaSenha(cfmsenha)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../global/guarda_valores_digitados.asp?opt=valor4", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor4=" + cfmsenha);
		}
	 function GuardaEmail(email)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../global/guarda_valores_digitados.asp?opt=valor5", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor5=" + email);
		}		

 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
</script>
</head>

<body onLoad="MM_preloadImages('img/botao_retornar/botao_retornar_f3.gif','img/botao_retornar/botao_retornar_f2.gif','img/botao_retornar/botao_retornar_f4.gif');FocusNoForm()">
<!-- <form name="form1" method="post" action="primeiro.asp">
</form> -->
<form action="conecta_primeiro.asp" method="post" name="login" id="login" autocomplete="OFF" onSubmit="return checksubmit()">  
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="10">&nbsp;</td>
  </tr>
  <tr>
    <td><table width="784" height="634" border="0" align="center" cellpadding="0" cellspacing="0" background="img/prim_acess.png">
  <!--DWLayoutTable-->
  <tr> 
    <td width="92" height="271">&nbsp;</td>
    <td width="14">&nbsp;</td>
    <td width="100">&nbsp;</td>
    <td width="53">&nbsp;</td>
    <td width="146">&nbsp;</td>
    <td width="379"><input name="log" type="hidden" id="log" value="prim" />
	              <input name="codigo_seguranca" type="hidden" id="codigo_seguranca" value="<%response.Write(codigo_seguranca)%>"></td>
  </tr>
  <tr>
    <td height="100" valign="top">&nbsp;</td>
    <td height="100" valign="top">&nbsp;</td>
    <td height="100" colspan="3" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="20"><img src="img/usuario.png" width="47" height="15"></td>
        <td height="20"><%if opt=999999 then%>
          <label>
            <input name="login" type="text" class="textbox" id="login"  size="25" onKeyUp="GuardaLogin(this.value)">
            </label>
          <%else%>
          <label>
            <input name="login" type="text" class="textbox" id="login"  value="<%RESPONSE.Write(lg)%>" size="25" onKeyUp="GuardaLogin(this.value)">
            </label>
          <%end if%></td>
        </tr>
      <tr>
        <td width="37%" height="20"><img src="img/senha.png" width="39" height="15"></td>
        <td width="63%" height="20"><input name="senha" type="password" class="textbox" id="senha" value="<%RESPONSE.Write(pas)%>" size="25" onKeyUp="GuardaSenha(this.value)"></td>
        </tr>
      <tr>
        <td width="37%" height="20"><img src="img/nova_senha.png" width="68" height="15"></td>
        <td height="20"><input name="pas1" type="password" class="textbox" id="pas1" value="<%RESPONSE.Write(pas1)%>" size="25" onKeyUp="GuardaNovaSenha(this.value)"></td>
        </tr>
      <tr>
        <td width="37%" height="20"><img src="img/confirme_senha.png" width="98" height="15"></td>
        <td height="20"><input name="pas2" type="password" class="textbox" id="pas2" value="<%RESPONSE.Write(pas2)%>" size="25" onKeyUp="GuardaConfirmaSenha(this.value)"></td>
        </tr>
      <tr>
        <td height="20"><img src="img/email.png" width="36" height="15"></td>
        <td height="20"><input name="email" type="text" class="textbox" size="25" value="<%RESPONSE.Write(mail_prim)%>" onKeyUp="GuardaEmail(this.value)" /></td>
        </tr>
      </table></td>
    <td height="100" valign="top"></td>
  </tr>
  <tr> 
    <td height="22" align="right"><input name="autorizo" type="checkbox" value="ok" checked /></td>
    <td align="right" valign="top"><!--DWLayoutEmptyCell-->&nbsp;</td>
    <td colspan="3" rowspan="2" valign="top"><img src="img/autorizo.png" width="238" height="43"></td>
    <td></td>
  </tr>
  <tr> 
    <td height="21">&nbsp;</td>
    <td>&nbsp;</td>
    <td></td>
  </tr>
  <tr> 
    <td height="20">&nbsp;</td>
    <td height="20">&nbsp;</td>
    <td height="20" colspan="2"><img src="img/codigo_figura.png" width="94" height="15"></td>
    <td height="20" valign="top"><input name="texto_imagem" type="text" class="textbox" id="texto_imagem" value="<%=ti%>" size="14"></td>
    <td height="20"></td>
  </tr>
  <tr> 
    <td height="81" colspan="5" align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="18%">&nbsp;</td>
        <td width="82%" align="center"><table width="205" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="205" height="81" background="img/seguranca/<%= imagem_seguranca %>.gif"><img src="img/seguranca/<%=codigo1%>.gif" width="40" height="40"><img src="img/seguranca/<%=codigo2%>.gif" width="40" height="40"><img src="img/seguranca/<%=codigo3%>.gif" width="40" height="40"><img src="img/seguranca/<%=codigo4%>.gif" width="40" height="40"><img src="img/seguranca/<%=codigo5%>.gif" width="40" height="40"></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
    <td height="81"><p>&nbsp;</p></td>
  </tr>
  <tr>
    <td height="92" align="left" valign="top"><!--DWLayoutEmptyCell-->&nbsp;</td>
    <td height="92" align="left" valign="top"><!--DWLayoutEmptyCell-->&nbsp;</td>
    <td height="92" colspan="3" align="left" valign="top"><table width="84%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td align="right"><a href="primeiro.asp?opt=999998"><img src="img/nv_cod.png" width="102" height="15" border="0"></a></td>
      </tr>
    </table></td>
    <td height="92" align="left" valign="middle"><table width="94%" height="46" border="0" cellpadding="0" cellspacing="0">
          <tr> 
            <td width="50%"><div align="center"><a href="default.asp" target="_top" onClick="MM_nbGroup('down','navbar1','botao_retornar','img/botao_retornar/botao_retornar_f3.gif',1);" onMouseOver="MM_nbGroup('over','botao_retornar','img/botao_retornar/botao_retornar_f2.gif','img/botao_retornar/botao_retornar_f4.gif',1);" onMouseOut="MM_nbGroup('out');"><img src="img/botao_retornar/botao_retornar.gif" alt="" name="botao_retornar" width="100" height="30" border="0" id="botao_retornar" /></a></div></td>
            <td width="50%"><div align="center"> 
                <input name="Submit" type="submit" class="confirmar" value="          ">
              </div></td>
          </tr>
        </table> </td>
  </tr>
</table></td>
  </tr>
</table>

<p>
  <table width="500" height="40" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr bgcolor="<%=cor%>">
                  <td><div align="center"> <font color="#FFFFFF"><strong><font size="1" face="Arial, Helvetica, sans-serif"> 
                <%response.Write(msg)%>
                </font></strong></font></div></td>
                </tr>
              </table></p>
			  </form>
</body>
</html>
