<%On Error Resume Next%>
<!--#include file="inc/funcoes.asp"-->



<%

opt = request.QueryString("opt")
escola= session("escola")
sistema_local="WR"
if opt="mail" then
sender = request.form("email")
nome = request.form("nome")
tipo = request.form("tipo")
if tipo = "Duvida" then
tipo = "Dúvida"
elseif tipo = "Solicitacao" then
tipo = "Solicitação"
elseif tipo = "Sugestao" then
tipo = "Sugestão"
elseif tipo = "Reclamacao" then
tipo = "Reclamação"
else
tipo = tipo
end if
assunto = request.form("assunto")
mensagem_rec = request.form("mensagem")

mensagem = "De: "&nome&", e-mail: "&sender&", Tipo: "&tipo&"."&mensagem_rec

	Set CON_wr = Server.CreateObject("ADODB.Connection") 
	ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_wr.Open ABRIR_wr
	
	Set RS1 = Server.CreateObject("ADODB.Recordset")
	consulta1 = "select * from Email where CO_Escola="&escola
	set RS1 = CON_wr.Execute (consulta1)
	
	mail_suporte=RS1("Suporte")
	mail_CC=RS1("Mail_Simplynet")


Set objCDO = Server.CreateObject("CDONTS.NewMail")
objCDO.From = mail_suporte
objCDO.To = mail_suporte
objCDO.BCC = mail_CC
objCDO.Subject = assunto
objCDO.Body = mensagem
objCDO.Send()
Set objCDO = Nothing

response.Redirect("faleconosco.asp?opt=ok")
else

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Usuario where CO_Usuario="&session("co_user")
		RS.Open SQL, CON
		
nome_user=RS("NO_Usuario")
email_user=RS("Email_Usuario")
		
%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="js/mm_menu.js"></script>
<script language="JavaScript">
 window.history.forward(1);
</script>
<script language="JavaScript">
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
function checksubmit()
{
  var obj = eval("document.forms[3].email");
  var txt = obj.value;
  if ((txt.length == 0)||((txt.length != 0) && ((txt.indexOf("@") < 1) || (txt.indexOf('.') < 7))))
  {
    alert('Email inválido');
	obj.focus();
	return false
  }
 if (document.form1.tipo.value == "0")
  {    alert("Por favor selecione um tipo de e-mail!")
   document.form1.tipo.focus()
    return false
 }
//aula = document.busca.aula.value;
//    if (aula.length > 3)
//  {    alert("O valor do campo Aula deve possuir menos que 3 caracteres")
//    document.busca.aula.focus()
//    return false
//  }
    if (document.form1.assunto.value == "")
  {    alert("Por favor digite um assunto para a Notícia!")
    document.form1.assunto.focus()
    return false
  }
      if (document.form1.mensagem.value == "")
  {    alert("O campo Mensagem não pode estar em branco!")
    document.form1.mensagem.focus()
    return false
  }
  return true
}
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
</script>
<link href="estilos.css" rel="stylesheet" type="text/css"></head>
<body link="#CC9900" background="img/fundo.gif" vlink="#CC9900" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"> 
<%call cabecalho(0)

nome = session("nome")
%>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
    <td height="10" valign="top">
<table width="100%" border="0" align="left" class="tb_caminho">
        <tr> 
            
          <td><font class="style-caminho">Voc&ecirc; 
            est&aacute; em <a href="inicio.asp" class="caminho">Web Diretor</a> > Fale Conosco</font> </font></td>
        </tr>
</table></td>
  </tr>
  <tr> 
    <td valign="top"><table width="100%" border="0" align="right" cellspacing="0">
        <tr> 
          <td class="tb_corpo"> <table width="100%" border="0" cellspacing="0">
              <%if opt = "ok" then%>
              <tr> 
                <td> 
                  <%
		call mensagens(0,6,2,0)
%>
                </td>
              </tr>
              <%end if%>
              <tr> 
                <td> 
                  <%
call mensagens(0,5,0,0) 
%>
                </td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td
><form name="form1" method="post" action="faleconosco.asp?opt=mail" onSubmit="return checksubmit()">
              <table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="333"> <div align="right"><font class="form_dado_texto"> Digite 
                      seu email:</font></div></td>
                  <td width="661"><font class="form_dado_texto"> 
				  <% if email_user="" or isnull(email_user) then %>
				  	<input name="email" type="text" class="select_style" id="email" size="75">
				  <%else
				   response.Write(email_user)%>
					<input name="email" type="hidden" value="<%=email_user%>">
				  <%end if%></font>
				  </td>
                </tr>
                <tr> 
                  <td> <div align="right"><font class="form_dado_texto"> Nome 
                      do usu&aacute;rio:</font></div></td>
                  <td><font class="form_dado_texto"> 
                    <% response.Write(nome_user)%>
                    <input name="nome" type="hidden" id="nome" value="<%=nome_user%>">
                    </font></td>
                </tr>
                <tr> 
                  <td> <div align="right"><font class="form_dado_texto"> Tipo 
                      de email:</font></div></td>
                  <td><select name="tipo" class="select_style" id="tipo">
                      <option value="0" selected></option>
                      <option value="Elogio">Elogio</option>
                      <option value="Duvida">D&uacute;vida</option>
                      <option value="Solicitacao">Solicita&ccedil;&atilde;o</option>
                      <option value="Sugestao">Sugest&atilde;o</option>
                      <option value="Reclamacao">Reclama&ccedil;&atilde;o</option>
                      <option value="Outros">Outros</option>
                    </select></td>
                </tr>
                <tr> 
                  <td> <div align="right"><font class="form_dado_texto"> Assunto:</font></div></td>
                  <td><input name="assunto" type="text" class="textInput" id="assunto" size="75"></td>
                </tr>
                <tr> 
                  <td valign="top"> <div align="right"><font class="form_dado_texto"> Mensagem:</font></div></td>
                  <td><textarea name="mensagem" cols="75" rows="15" wrap="VIRTUAL" class="textInput" id="mensagem"></textarea></td>
                </tr>
                <tr> 
                  <td><div align="right"></div></td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="2"><table width="500" border="0" align="center" cellspacing="0">
                      <tr> 
                        <td width="50%"> <div align="center"> 
                            <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','inicio.asp');return document.MM_returnValue" value="Cancelar">
                          </div></td>
                        <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
                            <input type="submit" name="Submit" value="Confirmar" class="botao_prosseguir">
                            </font></div></td>
                      </tr>
                    </table></td>
                </tr>
              </table>
            </form></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="40" valign="top"><img src="img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

</body>
</html>
<%end if %>
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
response.redirect("inc/erro.asp")
end if
%>