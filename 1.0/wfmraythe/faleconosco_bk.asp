
<!--#include file="inc/caminhos.asp"-->

<!--#include file="inc/caminhos.asp"-->
<!--#include file="inc/funcoes.asp"-->

<%

tp=session("tp")
nome = session("nome") 
co_user = session("co_user")
escola = session("escola")
co_aluno=Session("aluno_selecionado")

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0 
	
	Set CON = Server.CreateObject("ADODB.Connection")
 	ABRIR = "DBQ="& CAMINHO_wf& ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1
	
	Set CON_wr = Server.CreateObject("ADODB.Connection") 
	ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_wr.Open ABRIR_wr


opt=request.QueryString("opt")

if opt="ok" then
mensagem="MM_popupMsg('Mensagem enviada com sucesso!')"
end if

	SQL2 = "select * from TB_Usuario where CO_Usuario = " & co_user 
	set RS2 = CON.Execute (SQL2)
	
nome= RS2("NO_Usuario")
mail_db= RS2("TX_EMail_Usuario")


	SQL3 = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& co_aluno 
	set RS3 = CON1.Execute (SQL3)

nu_unidade= RS3("NU_Unidade")
co_curso= RS3("CO_Curso")
co_etapa= RS3("CO_Etapa")
co_turma= RS3("CO_Turma")

	SQL2 = "select * from TB_Etapa where CO_Curso = '" & co_curso &"' AND CO_Etapa='"&co_etapa&"'"
	set RS2 = CON0.Execute (SQL2)


	Set RS1 = Server.CreateObject("ADODB.Recordset")
	consulta1 = "select * from Email where CO_Escola="&escola
	set RS1 = CON_wr.Execute (consulta1)
	
	'mail_suporte=RS1("WebFamilia")
	'mail_CC=RS1("Mail_Simplynet")
	mail_CC=RS1("Mail_Rodan")
	
if RS1.EOF then
response.Write("<font=arial size=2>Endereço de e-mail da escola não cadastrado. Favor contatar a coordenação.<br> <a href=javascript:window.history.go(-1)>voltar</a></font>")
response.end()
else
	mail_escola=RS1("WebFamilia")
	
	if isnull(mail_escola) or mail_escola="" then
		response.Write("<font face=arial size=2>Endereço de e-mail da escola não cadastrado. Favor contatar a coordenação.<br> <a href=javascript:window.history.go(-1)>voltar</a></font>")
		response.end()	
	end if

end if

if isnull(mail_db) or mail_db="" then
	if isnull(mail_form) or mail_form="" then
	sender=mail_suporte
	else
	sender=mail_form
	end if
else
sender=mail_db
end if

if opt="mail" then
mail_form = request.form("email")
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
	'SQL3 = "select * from TB_Operador" 
	'set RS3 = CON.Execute (SQL3)
	
'mail_escola=RS3("Login")	
'mail_escola=RS1("Mail_RoDan")
'mail_suporte=RS1("Mail_RoDan")
'mail_CC=RS1("Mail_RoDan")

Set objCDO = Server.CreateObject("CDONTS.NewMail")
objCDO.From = sender
objCDO.To = mail_escola
objCDO.BCC = mail_CC
objCDO.Subject = assunto
objCDO.Body = mensagem
objCDO.Send()
Set objCDO = Nothing

response.Redirect("faleconosco.asp?opt=ok")
else
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Web Família</title>
<link href="estilo.css" rel="stylesheet" type="text/css" />
<script type="text/JavaScript">
<!--
function MM_popupMsg(msg) { //v1.0
  alert(msg);
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
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
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

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
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
  var obj = eval("document.forms[0].email");
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
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
function FocusNoForm() 
{ 
//formlogin.nome.value="testes"; 
<%if mail="" or isnull(mail) then%>
form1.email.focus(); 
<%else%>
form1.tipo.focus(); 
<%end if%>
} 
//-->
</script>
</head>

<body onLoad="MM_preloadImages(<%response.Write(swapload)%>);<%response.Write(mensagem)%>;FocusNoForm()"><%if tp="R" and opt="sa" then%>
<%end if
%>

<table width="1000" height="1039" border="0" align="center" cellpadding="0" cellspacing="0" class="tabela_principal">
  <%
			ano = DatePart("yyyy", now)
			mes = DatePart("m", now) 
			dia = DatePart("d", now) 
			hora = DatePart("h", now) 
			min = DatePart("n", now) 
select case mes
 case 1 
 mes = "janeiro"
 case 2 
 mes = "fevereiro"
 case 3 
 mes = "março"
 case 4
 mes = "abril"
 case 5
 mes = "maio"
 case 6 
 mes = "junho"
 case 7
 mes = "julho"
 case 8 
 mes = "agosto"
 case 9 
 mes = "setembro"
 case 10 
 mes = "outubro"
 case 11 
 mes = "novembro"
 case 12 
 mes = "dezembro"
end select

data = dia &" / "& mes &" / "& ano
data= FormatDateTime(data,1) 			

			horario = hora & ":"& min%>
  <tr>
    <td height="998"><table width="200" height="998" border="0" cellpadding="0" cellspacing="0">
        <!--DWLayoutTable-->
        <tr> 
          <td height="130" colspan="3">
            <%call cabecalho(nivel)%>
          </td>
        </tr>
        <tr class="tabela_menu"> 
          <td width="172" height="144" rowspan="4" valign="top" class="tabela_menu"><p>&nbsp;</p>
            <% call menu_lateral(nivel)%>
            <p>&nbsp;</p></td>
          <td width="640" height="12" nowrap="nowrap"><p class="style1">&nbsp;&nbsp;Ol&aacute; 
              <span class="style2">
              <%response.Write(nome)%>
              </span> , &uacute;ltimo acesso dia 
              <% Response.Write(session("dia_t")) %>
              &agrave;s 
              <% Response.Write(session("hora_t")) %>
            </p></td>
          <td width="188"><p align="right" class="style1"> 
              <%response.Write(data)%>
            </p></td>
        </tr>
        <tr class="tabela_menu"> 
          <td height="5" colspan="2"><p><img src="img/linha-pontilhada_grande.gif" alt="" width="828" height="5" /></p></td>
        </tr>
      <tr class="tabela_menu">
        <td height="19" colspan="2">&nbsp;</td>
      </tr>		
        <tr class="tabela_menu"> 
          <td height="832" colspan="2" valign="top"> <p><img src="img/fale_conosco.jpg" width="700" height="30"></p>
            <table width="80%" border="0" align="center" cellspacing="0"  class="tabela_principal">
              <tr> 
                <td class="tb_corpo"></td>
              </tr>
              <tr> 
                <td
><form name="form1" method="post" action="faleconosco.asp?opt=mail" onSubmit="return checksubmit()">
                    <table width="700" border="0" align="center" cellspacing="0">
                      <tr> 
                        <td width="183" > <div align="right"><font class="style3">Seu 
                            email:</font></div></td>
                        <td width="513"><div align="left"><font class="style1">
						<% if mail_db="" or isnull(mail_db) then
						%><input name="email" type="text" class="borda" id="email" size="75">
						<%			
						else
						response.Write(mail_db)%>
						<input name="email" type="hidden" value="<%=mail_db%>">
						<%end if%></font></div>
						</td>
                      </tr>
                      <tr> 
                        <td > <div align="right"><font class="style3">Nome do 
                            usu&aacute;rio:</font></div></td>
                        <td><div align="left"><font class="style1"> 
                          <% response.Write(nome)%>
                          <input name="nome" type="hidden" id="nome" value="<%=nome%>">
                          </font></div></td>
                      </tr>
                      <tr> 
                        <td > <div align="right"><font class="style3">Tipo de 
                            email:</font></div></td>
                        <td><select name="tipo" class="borda" id="tipo">
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
                        <td > <div align="right"><font class="style3">Assunto:</font></div></td>
                        <td><input name="assunto" type="text" class="borda" id="assunto" size="75"></td>
                      </tr>
                      <tr> 
                        <td valign="top"  > <div align="right"><font class="style3">Mensagem:</font></div></td>
                        <td><textarea name="mensagem" cols="72" rows="15" wrap="VIRTUAL" class="borda" id="mensagem"></textarea></td>
                      </tr>
                      <tr> 
                        <td ><div align="right"></div></td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr> 
                        <td colspan="2" ><table width="100%" border="0" align="left" cellspacing="0">
                            <tr> 
                              <td> <div align="center"> </div>
                                <div align="center"><font size="3" face="Courier New, Courier, mono"> 
                                  <input type="submit" name="Submit" value="Confirmar" class="borda_bot2">
                                  </font></div></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table>
                  </form></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td width="1000"><img src="img/rodape.jpg" width="1000" height="40" /></td>
  </tr>
</table>
</body>
</html>
<%end if%>
