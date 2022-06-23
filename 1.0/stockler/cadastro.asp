<%On Error Resume Next%>


<!--#include file="inc/funcoes.asp"-->
<!--#include file="inc/funcoes5.asp"-->

<%
opt=request.QueryString("opt")
obr = request.QueryString("obr")
nivel=0
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local="WR"
chave=session("chave")
ano_info=nivel&"-"&chave&"-"&ano_letivo
co_usr = session("co_user")

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

 call navegacao (CON,chave,nivel)
navega=Session("caminho")

%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="estilos.css" type="text/css">
<script language="JavaScript">
 window.history.forward(1);
</script>
<script language="JavaScript" type="text/JavaScript">
<!--


function MM_popupMsg(msg) { //v1.0
  alert(msg);
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
<script>
function valid(){
var pas1 = document.cadastro.pas1.value
//var userLength = document.theform.user.value.length
var pas2 = document.cadastro.pas2.value
//var passLength = document.theform.pass.value.length
if(pas2 != pas1){
alert("a senha digitada no campo confirmação não é igual a digitada no campo senha. Para que o senha seja alterada é necessário que estas sejam iguais");
document.cadastro.pas1.focus()
return false
	}
	if(pas2 == pas1){
return true
	}
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
</script>
<script language="JavaScript"> 
function FocusNoForm() 
{ 
//formlogin.nome.value="testes"; 
<%if opt="sh" then%>
cadastro.pas1.focus(); 
<%elseif opt="ml" or opt="err1" then%>
cadastro.email.focus(); 
<%else%>
cadastro.login.focus(); 
<%end if%>
} 
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
</script> 

</head>

<body leftmargin="0" topmargin="0" marginwidth="0" background="img/fundo.gif" marginheight="0" onLoad="FocusNoForm()">
<%
 call cabecalho (nivel) %>
 
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega&" > Alterar Senha")

%>
      </font></td>
  </tr> 
    <td height="10">
<table width="100%" border="0" cellspacing="0">
        <% if opt = "lg" then%>
        <tr> 
          <td> 
            <%
		call mensagens(0,7,0,0)
%>
          </td>
        </tr>
	        <% elseif opt = "err0" then%>
        <tr> 
          <td> 		  
            <%
		call mensagens(0,13,1,0)
%>
          </td>
        </tr>
        <tr> 
          <td> 
            <%
		call mensagens(0,7,0,0)
%>
          </td>
        </tr>						
        <% elseif opt = "sh" then%>
        <tr> 
          <td> 
            <%
		call mensagens(0,8,0,0)
%>
          </td>
        </tr>
        <% elseif opt = "ml" then%>
        <tr> 
          <td> 
            <%
		call mensagens(0,14,0,0)
%>
          </td>
        </tr>
	        <% elseif opt = "err1" then%>
        <tr> 
          <td> 		  
            <%
		call mensagens(0,15,1,0)
%>
          </td>
        </tr>
        <tr> 
          <td> 
            <%
		call mensagens(0,14,0,0)
%>		
		
        <%else%>
        <tr> 
          <td> 
            <%	call mensagens(0,114,0,0) 

%>
          </td>
        </tr>
        <% end if%>
        <tr> 
          <td> </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td valign="top"> 
      <table width="1000" border="0" cellspacing="0">
        <tr> 
          <td width="1000" valign="top"> <div align="right"> 
              <%If opt="lg" or opt="err0" then
		call alterads(0,0,0,0,co_usr)
elseif opt="sh" then
		call alterads(1,0,0,0,co_usr)
elseif opt="ml" or opt="err1" then
		call alterads(2,0,0,0,co_usr)
elseif opt="cadastrar" then
		login = request.form("login")
		pass = request.form("pas1")
		mail = request.form("email")
		call alterads(99,login,pass,mail,co_usr)
end if %>
            </div></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td height="40"><img src="img/rodape.jpg" width="1000" height="40"></td>
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
response.redirect("inc/erro.asp")
end if
%>