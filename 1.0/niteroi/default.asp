<!--#include file="inc/caminhos.asp"-->
<%
lg = session("lg")


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
	msg="O campo Usuário é obrigatório!"
	tipo="e"
	lg = ""
	pas= session("senha")
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
	pas= session("senha")
	ti = session("ti")
case 05
	msg="Usuário não autorizado."
	tipo="e"
	lg = ""
	pas= session("senha")
	ti = session("ti")
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
	pas= session("senha")
	ti = ""
case 08
	msg="Código da figura Incorreto!"
	tipo="e"
	lg = session("lg")
	pas= session("senha")
	ti = ""
case 999998
'aparece quando o botão novo código é clicado
	lg = session("valor1")
	pas= session("valor2")
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

imagem = Array("10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25")

	Randomize
		gerarfundo = gerarfundo & imagem(ubound(imagem) * Rnd)



imagem_seguranca = gerarfundo
%>

<html>
<head>
<title>Web Diretor</title>
<link href="estilos.css" rel="stylesheet" type="text/css">
<script language="JavaScript"> 
function FocusNoForm() 
{ 
//formlogin.nome.value="testes"; 
<%if opt=02 or opt=06 then%>
login.senha.focus(); 
<%elseif opt=03 or opt=07 or opt=08 or opt=999998 then%>
login.texto_imagem.focus(); 
<%else%>
login.login.focus(); 
<%end if%>
} 
</script> 
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<SCRIPT LANGUAGE="JAVASCRIPT" TYPE="TEXT/JAVASCRIPT">
function redirect(){
setTimeout("submitform(form1)",600000);
}
</SCRIPT>
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
	
	
	 function GuardaSenha(senha)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../global/guarda_valores_digitados.asp?opt=valor2", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor2=" + senha);
		}

 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
</script>
<script language="JavaScript">
 window.history.forward(1);
</script>
</head>

<body bgcolor="#FFFFFF" topmargin="20" marginheight="20" onLoad="FocusNoForm()">
<form name="form1" method="post" action="default.asp">
</form>
<form action="conecta.asp" method="post" name="login" id="login">
          
  <table width="801" height="535" border="0" align="center" cellpadding="0" cellspacing="0" background="img/login.png">
    <tr> 
              <td width="36" height="235">&nbsp;</td>
              <td height="235" colspan="2">&nbsp;</td>
    </tr>
            <tr> 
              <td height="298">&nbsp;</td>
              <td width="377" height="298" valign="top">&nbsp;</td>
              <td width="388" height="298" valign="top"><table width="339" height="244" border="0" align="left" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="339" height="40"><table width="334" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="334" height="20" align="right"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td height="25"><table width="81%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td width="43%" height="25" align="right"><img src="img/usuario.png" alt="" width="47" height="15"></td>
                              <td width="57%" height="25" align="right"><%if opt=999999 then%>
                                <input name="login" type="text" class="textInput" id="login" size="25"  onKeyUp="GuardaLogin(this.value)">
                                <%else%>
                                <input name="login" type="text" class="textInput" id="login" value="<%=lg%>" size="25" onKeyUp="GuardaLogin(this.value)">
                              <%end if%></td>
                            </tr>
                          </table></td>
                        </tr>
                        <tr>
                          <td width="334" height="25"><table width="81%" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td width="41%" height="25" align="right"><img src="img/senha.png" alt="" width="39" height="15"></td>
                              <td width="59%" height="25" align="right"><input name="senha" type="password" class="textInput" id="senha" value="<%=pas%>" size="25" onKeyUp="GuardaSenha(this.value)"></td>
                            </tr>
                          </table></td>
                        </tr>
                        <tr>
                          <td height="25"><table width="270" border="0" cellspacing="0" cellpadding="0">
                            <tr>
                              <td width="61%" height="25" align="right"><img src="img/codigo_figura.png" alt="" width="94" height="15"></td>
                              <td width="39%" height="25" align="right"><input name="texto_imagem" type="text" class="textInput" id="texto_imagem" value="<%=ti%>" size="15"></td>
                            </tr>
                          </table></td>
                        </tr>
                      </table></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="81"><table width="205" border="0" align="center" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="205" height="81" background="img/seguranca/<%= imagem_seguranca %>.gif"><img src="img/seguranca/<%=codigo1%>.gif" alt="cod" width="40" height="40"><img src="img/seguranca/<%=codigo2%>.gif" alt="cod" width="40" height="40"><img src="img/seguranca/<%=codigo3%>.gif" alt="cod" width="40" height="40"><img src="img/seguranca/<%=codigo4%>.gif" alt="cod" width="40" height="40"><img src="img/seguranca/<%=codigo5%>.gif" alt="cod" width="40" height="40"></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="20" valign="top"><div align="center">
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="54%">&nbsp;</td>
                        <td width="46%"><a href="default.asp?opt=999998"><img src="img/nv_cod.png" alt="" width="102" height="15" border="0"></a></td>
                      </tr>
                    </table>
                  </div></td>
                </tr>
                <tr>
                  <td height="29"><div align="center">
                    <input name="escola" type="hidden" id="escola" value="8">
                    <input name="Enviar" type="image" src="img/botao_autenticar.gif" alt="autenticar" width="130" height="30" border="0">
                  </div></td>
                </tr>
                <tr>
                  <td height="18" valign="top"><table width="337" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="337" height="13"><div align="center"></div></td>
                    </tr>
                  </table></td>
                </tr>
              </table></td>
            </tr>		
  </table>
	 <% if tipo="e" then%>    
  <table width="600" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr> 
      <td>&nbsp;</td>
    </tr>
    <tr> 
        <td><table width="550" height="40" border="1" align="center" cellpadding="0" cellspacing="1" bordercolor="#FF0000">
        <tr bgcolor="<%=cor%>"> 
            <td><div align="center"> <font color="#FFFFFF"><strong><font size="1" face="Arial, Helvetica, sans-serif"> 
                <%response.Write(msg)%>
                </font></strong></font></div></td>
          </tr>
      </table></td>
    </tr>
  </table>
  <%end if%>
</form>
</body>
</html>
