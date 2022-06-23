<!--#include file="inc/caminhos.asp"-->
<%
opt = request.QueryString("opt")
escola = request.form("escola")

	Set CON = Server.CreateObject("ADODB.Connection")
 	ABRIR = "DBQ="& CAMINHO_wf& ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR

	Set CON_ctrle = Server.CreateObject("ADODB.Connection") 
	ABRIR_ctrle = "DBQ="& CAMINHO_ctrle & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_ctrle.Open ABRIR_ctrle
	
	Set CON_wr = Server.CreateObject("ADODB.Connection") 
	ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_wr.Open ABRIR_wr


if opt="mail" then
	lg = request.form("login")
	mail = request.form("mail")
	codigo_seguranca=session("codigo_seguranca")
	texto_imagem =request.form("texto_imagem")
	texto_imagem=LCase(texto_imagem)
	
	IF lg="" or isnull(lg) then
		session("ti")=ti
		session("mail")=mail
		response.Redirect("envia_senha.asp?opt=01")
	elseIF mail="" or isnull(mail) then
		session("lg")=lg 
		endereco= ""
		session("ti")=ti
		response.Redirect("envia_senha.asp?opt=09")
	elseIF texto_imagem="" or isnull(texto_imagem) then
		session("lg")=lg 
		session("mail")=mail		
		session("ti")=""		
		endereco= ""
		response.Redirect("envia_senha.asp?opt=03")
	end if
		session("lg")=lg 
		session("mail")=mail		
		session("ti")=ti
		
	if codigo_seguranca ="" then
		response.Redirect("envia_senha.asp?opt=07")
	elseif codigo_seguranca <> texto_imagem then
		response.Redirect("envia_senha.asp?opt=08")
	end if
	

	SQL = "select * from TB_Usuario where CO_Usuario = " & lg
	set RS = CON.Execute (SQL)
	
	IF RS.EOF then
	response.Redirect("envia_senha.asp?opt=04")
	end if	
	nome =Rs("NO_Usuario")
	email =Rs("TX_EMail_Usuario")
	senha =Rs("Senha")
	autorizo =Rs("IN_Aut_email")
	
	Set RS1 = Server.CreateObject("ADODB.Recordset")
	consulta1 = "select * from Email where CO_Escola="&escola
	set RS1 = CON_wr.Execute (consulta1)	
	
	Set RS2 = Server.CreateObject("ADODB.Recordset")
	consulta2 = "select * from Login where CO_Escola="&escola
	set RS2 = CON_wr.Execute (consulta2)	
		
	'mail_suporte=RS1("WebFamilia")
	mail_CC=RS1("Mail_Simplynet")		


	'if email=mail AND autorizo= TRUE then

'	Set objCDOSYSMail = Server.CreateObject("CDO.Message")
'	Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
'	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
'	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")= 25
'	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
'	objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30
'	objCDOSYSCon.Fields.update
'	Set objCDOSYSMail.Configuration = objCDOSYSCon
'	objCDOSYSMail.From = email_suporte_escola
'	objCDOSYSMail.To = email
'	objCDOSYSMail.BCC =""
'	'objCDOSYSMail.AddAttachment("e:\home\simplynet\dados\liessin\"&arquivo&"")
'	objCDOSYSMail.Subject = "Informações Web Diretor Família"
'	objCDOSYSMail.HtmlBody = "<font size=2 face=Arial, Verdana, Courier New, Courier, mono>Prezado(a) "&nome&"<BR><BR><BR> Conforme sua solicitação lembramos que o usuário "&lg&" possui a senha "&senha&".<BR><BR><BR> Atenciosamente,<BR><BR><BR> Administração "&nome_da_escola&"</font>"
	
	response.Write("<font size=2 face=Arial, Verdana, Courier New, Courier, mono>Prezado(a) "&nome&"<BR><BR><BR> Conforme sua solicitação lembramos que o usuário "&lg&" possui a senha "&senha&".<BR><BR><BR> Atenciosamente,<BR><BR><BR> Administração "&nome_da_escola&"</font>")
	response.End()
	'objCDOSYSMail.Send
	
	
	Set objCDOSYSMail = Nothing
	Set objCDOSYSCon = Nothing 
	
	
		response.Redirect("envia_senha.asp?opt=ok")
	'else	
	'	response.Redirect("envia_senha.asp?opt=10")
	'end if	

elseif opt="ok" then
	response.Redirect("default.asp?opt=09")
else
lg = session("lg")

	
	consulta_ctl = "select * from TB_Controle"
	set tabela_ctl = CON.Execute (consulta_ctl)

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
	msg="O campo Usuário é obrigatório!"
	tipo="e"
	lg = ""
	mail= session("mail")	
	ti = session("ti")
case 02
	msg="O campo Senha é obrigatório!"
	tipo="e"
	lg = session("lg")
	pas= ""
	ti = session("ti")
case 03
	msg="Digitar o código da figura é obrigatório!"
	tipo="e"
	lg = session("lg")
	pas= session("senha")
	ti = ""
case 04
	msg="O Usuário "&lg&" não existe!"
	tipo="e"
	lg = ""
	mail= ""	
	ti = ""
case 05
	msg="Usuário não autorizado."
	tipo="e"
	lg = ""
	mail= ""	
	ti = ""
case 06
	msg="Senha Incorreta!"
	tipo="e"
	lg = session("lg")
	pas= ""
	ti = session("ti")
case 07
	msg="Tempo de digitação do código da figura excedido. Tente novamente."
	tipo="e"
	lg = session("lg")
	mail= session("mail")	
	ti = ""
case 08
	msg="Código da figura Incorreto!"
	tipo="e"
	lg = session("lg")
	mail= session("senha")
	ti = ""
case 09
	msg="Campo email é obrigatório"
	tipo="e"
	lg = session("lg")
	ti = session("ti")
case 10
	msg="O endereço de e-mail digitado não coincide com o cadastrado na escola. Tente novamente ou entre em contato com a escola para obter uma nova senha."
	tipo="e"
	lg = session("lg")
	mail= session("mail")	
	ti = session("ti")	
case 999998
'aparece quando o botão novo código é clicado
	lg = session("valor1")
	mail= session("valor2")
case else
session.Contents.RemoveAll()
end select

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


session("codigo_seguranca") = codigo1&codigo2&codigo3&codigo4&codigo5

'imagem = Array("10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25")

imagem = Array("12","15","17","18","19","21","24")


	Randomize
		gerarfundo = gerarfundo & imagem(ubound(imagem) * Rnd)



imagem_seguranca = gerarfundo
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
function redirect(){
setTimeout("submitcodigo(form1)",600000);
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
						
						
	 function GuardaLogin(login)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../global/guarda_valores_digitados.asp?opt=valor1", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor1=" + login);
		}
	
	
	 function GuardaMail(mail)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../global/guarda_valores_digitados.asp?opt=valor2", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor2=" + mail);
		}

 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
</script>

</head>

<body onLoad="FocusNoForm();MM_preloadImages('img/botao_retornar/botao_retornar_f3.gif','img/botao_retornar/botao_retornar_f2.gif','img/botao_retornar/botao_retornar_f4.gif')">
<form name="form1" method="post" action="default.asp">
</form>
<form action="envia_senha.asp?opt=mail" method="post" name="login" id="login">
  <table width="1000" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td height="10">&nbsp;</td>
    </tr>
    <tr>
      <td><table width="801" height="560" border="0" align="center" background="img/esqueci_senha.png">
        <!--DWLayoutTable-->
        <tr>
          <td height="265" colspan="2">&nbsp;</td>
        </tr>
        <tr valign="top">
          <td width="426" rowspan="2" align="center"><table width="83%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="16%">&nbsp;</td>
              <td width="71%"><table width="334" height="196" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="81" valign="bottom"><table width="334" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="25"><table width="81%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td width="41%" align="right"><img src="img/usuario.png" width="47" height="15"></td>
                          <td width="59%" align="right">
						  <%if opt=999999 then%>
                            <input name="login" type="text" class="textbox" id="login" size="25" onKeyUp="GuardaLogin(this.value)">
                            <%else%>
                            <input name="login" type="text" class="textbox" id="login" value="<%response.Write(lg)%>" size="25" onKeyUp="GuardaLogin(this.value)">
                            <%end if%></td>
                        </tr>
                      </table></td>
                    </tr>
                    <tr>
                      <td height="25"><div align="left">
                        <table width="81%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="39%" align="right"><img src="img/email.png" width="36" height="15"></td>
                            <td width="61%" align="right"><input name="mail" type="text" class="textbox" id="mail" value="<%response.Write(mail)%>" size="25" onKeyUp="GuardaMail(this.value)"></td>
                          </tr>
                        </table>
                      </div></td>
                    </tr>
                    <tr>
                      <td height="25"><div align="left">
                        <table width="81%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="59%" align="right"><img src="img/codigo_figura.png" width="94" height="15"></td>
                            <td width="41%" align="right"><input name="texto_imagem" type="text" class="textbox" id="texto_imagem" value="<%response.Write(ti)%>" size="15"></td>
                            </tr>
                          </table>
                        </div></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td width="334" height="81"><table width="205" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="205" height="81" background="img/seguranca/<%= imagem_seguranca %>.gif"><img src="img/seguranca/<%=codigo1%>.gif" alt="cod" width="40" height="40"><img src="img/seguranca/<%=codigo2%>.gif" alt="cod" width="40" height="40"><img src="img/seguranca/<%=codigo3%>.gif" alt="cod" width="40" height="40"><img src="img/seguranca/<%=codigo4%>.gif" alt="cod" width="40" height="40"><img src="img/seguranca/<%=codigo5%>.gif" alt="cod" width="40" height="40"></td>
                      </tr>
                    </table></td>
                </tr>
                <tr>
                  <td height="18" valign="top"><div align="center">
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="55%" height="15">&nbsp;</td>
                        <td width="45%" valign="top"><a href="envia_senha.asp?opt=999998"><img src="img/nv_cod.png" alt="" width="102" height="15" border="0"></a></td>
                      </tr>
                    </table>
                  </div></td>
                </tr>
              </table></td>
              <td width="13%">&nbsp;</td>
            </tr>
          </table></td>
          <td width="365" height="135" align="center"><!--DWLayoutEmptyCell-->&nbsp;</td>
        </tr>
        <tr>
          <td height="77" align="left" valign="bottom">&nbsp;&nbsp;
            <input name="escola" type="hidden" id="escola" value="8">
            <table width="312" border="0" align="left" cellpadding="0" cellspacing="0">
              <tr>
                <td width="50%"><div align="center"><a href="default.asp" target="_top" onClick="MM_nbGroup('down','navbar1','botao_retornar','img/botao_retornar/botao_retornar_f3.gif',1);" onMouseOver="MM_nbGroup('over','botao_retornar','img/botao_retornar/botao_retornar_f2.gif','img/botao_retornar/botao_retornar_f4.gif',1);" onMouseOut="MM_nbGroup('out');"><img src="img/botao_retornar/botao_retornar.gif" alt="" name="botao_retornar" width="100" height="30" border="0" id="botao_retornar" /></a></div></td>
                <td width="50%"><div align="center">
                  <input name="Submit" type="submit" class="confirmar" value="          ">
                  </div></td>
                </tr>
            </table></td>
        </tr>
        <tr>
          <td align="center"><!--DWLayoutEmptyCell-->&nbsp;</td>
          <td height="73" align="left" valign="bottom"><!--DWLayoutEmptyCell-->&nbsp;</td>
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
  </table>
</p>
</form>
</body>
</html>
<%end if%>