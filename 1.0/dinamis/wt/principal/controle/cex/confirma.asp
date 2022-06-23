<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->



<%
e_vencimento = request.QueryString("opt")
e_cod = request.QueryString("c")
ano_letivo = session("ano_letivo")

co_usr = session("co_user")
nivel=4

chave=session("chave")
session("chave")=chave

nvg_split=split(chave,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano_info=nivel&"-"&chave&"-"&ano_letivo


	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
		
	Set CON2 = Server.CreateObject("ADODB.Connection") 
	ABRIR2 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON2.Open ABRIR2	
	
	Set CON6 = Server.CreateObject("ADODB.Connection") 
	ABRIR6 = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON6.Open ABRIR6		
	
	Set CON7 = Server.CreateObject("ADODB.Connection") 
	ABRIR7 = "DBQ="& CAMINHO_pf & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON7.Open ABRIR7		
	
	Set CON8 = Server.CreateObject("ADODB.Connection") 
	ABRIR8 = "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON8.Open ABRIR8							

 call navegacao (CON,chave,nivel)
navega=Session("caminho")	

		Set RSa = Server.CreateObject("ADODB.Recordset")
		SQLa = "SELECT TB_Alunos.NO_Aluno, TB_Alunos.TP_Resp_Fin, TB_Alunos.IN_Sexo, TB_Matriculas.NU_Unidade, TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma FROM TB_Alunos, TB_Matriculas where TB_Matriculas.CO_Matricula = TB_Alunos.CO_Matricula AND TB_Matriculas.CO_Matricula = "& e_cod &" AND TB_Matriculas.NU_Ano = "& ano_letivo
		RSa.Open SQLa, CON2
		

		
		nome_aluno = RSa("NO_Aluno")
		tp_resp_fin = RSa("TP_Resp_Fin")
		in_sexo = RSa("IN_Sexo")		
		nu_unidade = RSa("NU_Unidade")
		co_curso = RSa("CO_Curso")
		co_etapa = RSa("CO_Etapa")
		co_turma = RSa("CO_Turma")
		
	
		
		Set RSc = Server.CreateObject("ADODB.Recordset")
		SQLc = "SELECT NO_Contato,CO_CPF_PFisica, TX_EMail FROM TB_Contatos where CO_Matricula = "& e_cod &" AND TP_Contato = '"& tp_resp_fin&"'"
		RSc.Open SQLc, CON6		
		
		If RSc.EOF then
			nome_resp ="Nome não cadastrado para o "&tp_resp_fin
		else
			nome_resp = RSc("NO_Contato")
			cpf_resp = RSc("CO_CPF_PFisica")
			email_resp =RSc("TX_EMail")
					
			if cpf_resp = "" or isnull(cpf_resp) then
			
			else
				cpf_resp = replace(cpf_resp,"-","")
				cpf_resp = replace(cpf_resp,".","")				
			end if
			
			if isnull(email_resp) or email_resp="" then
				email_resp ="Email não cadastrado"
			end if		
				
		end if		
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="file:../../../../img/mm_menu.js"></script>
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
//-->
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
						
						
						 function recuperarCurso(uTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=c", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
document.all.divEtapa.innerHTML ="<select class=select_style></select>"
document.all.divTurma.innerHTML = "<select class=select_style></select>"
//recuperarEtapa()
                                                           }
                                               }

                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=e", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select class=select_style></select>"
//recuperarTurma()
                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }

 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
var checkflag = "false";
function check(field) {
if (checkflag == "false") {
for (i = 0; i < field.length; i++) {
field[i].checked = true;}
checkflag = "true";
return "Desmarcar Todos"; }
else {
for (i = 0; i < field.length; i++) {
field[i].checked = false; }
checkflag = "false";
return "Marcar Todos"; }
}
                        </script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../../img/menu_r1_c2_f3.gif','../../../../img/menu_r1_c2_f2.gif','../../../../img/menu_r1_c2_f4.gif','../../../../img/menu_r1_c4_f3.gif','../../../../img/menu_r1_c4_f2.gif','../../../../img/menu_r1_c4_f4.gif','../../../../img/menu_r1_c6_f3.gif','../../../../img/menu_r1_c6_f2.gif','../../../../img/menu_r1_c6_f4.gif','../../../../img/menu_r1_c8_f3.gif','../../../../img/menu_r1_c8_f2.gif','../../../../img/menu_r1_c8_f4.gif','../../../../img/menu_direita_r2_c1_f3.gif','../../../../img/menu_direita_r2_c1_f2.gif','../../../../img/menu_direita_r2_c1_f4.gif','../../../../img/menu_direita_r4_c1_f3.gif','../../../../img/menu_direita_r4_c1_f2.gif','../../../../img/menu_direita_r4_c1_f4.gif','../../../../img/menu_direita_r6_c1_f3.gif','../../../../img/menu_direita_r6_c1_f2.gif','../../../../img/menu_direita_r6_c1_f4.gif')">
<% call cabecalho (nivel)
	  %>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
                    
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
	  </td>
	  </tr>
                  <tr> 
                    
    <td height="10"> 
      <%	call mensagens(4,914,0,email_resp) 
	  
	 
%>
</td></tr>
<tr>

            <td valign="top"> 
                
        <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td width="653" height="15" class="tb_tit">Confirmar Envio</td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr valign="top">
		<td><form name="alteracao" method="post" action="email_anexo.asp?opt=<%response.Write(e_vencimento)%>&c=<%response.Write(e_cod)%>"><table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td>
    <%		nome_meses=GeraNomesNovaVersao("MES",e_vencimento,variavel2,variavel3,variavel4,variavel5,CON0,outro)
		meses = 1
		ck_mail = e_cod&"#!#"&email_resp&"#!#"&nome_resp&"#!#"&nome_aluno&"#!#"&meses&"#!#"&nome_meses&"#!#"&in_sexo		

		Set RSa = Server.CreateObject("ADODB.Recordset")
		SQLa = "SELECT TX_Titulo_Assunto FROM TB_Email_Assunto where CO_Assunto = 20"
		RSa.Open SQLa, CON0
		
		assunto_padrao=RSa("TX_Titulo_Assunto")

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT TX_Conteudo_Email FROM TB_Email_Mensagem where CO_Email = 20"
		RS0.Open SQL0, CON0																						
	
		if RS0.EOF then
			mensagem_padrao = ""
		else
			mensagem_padrao = RS0("TX_Conteudo_Email")	
		end if	

	dados_form = split(ck_mail,", ")

	for e = 0 to ubound(dados_form)
		dados_mensagem = split(dados_form(e),"#!#")
		co_matricula = dados_mensagem(0)	
		end_email=dados_mensagem(1)		
		resp_fin=dados_mensagem(2)
		nome_aluno = dados_mensagem(3)		
		meses = dados_mensagem(4)
		desc_meses = replace(dados_mensagem(5),"-",", ")					
		in_sexo=dados_mensagem(6)							

		if InStr(end_email,"@")=0 then
		else 	
			assunto=assunto_padrao&" - Ref: "&	co_matricula
			
			mensagem = "Prezado(a) Senhor(a) "&resp_fin&", respons&aacute;vel "
			if in_sexo = "F" then
				mensagem=mensagem&"pela aluna "
			else
				mensagem=mensagem&"pelo aluno "	
			end if	
			
			mensagem=mensagem&nome_aluno&", matr&iacute;cula "&co_matricula&Chr(13)&Chr(13)	
			
			mensagem=mensagem&"Segue em anexo a 2&ordf; Via do Bloqueto banc&aacute;rio da parcela referente ao m&ecirc;s de "&desc_meses&" do ano letivo de "&ano_letivo&Chr(13)&Chr(13)	 

		'	mensagem=mensagem&"Qualquer d&uacute;vida por favor entre em contato com a secretaria da escola D&iacute;namis."&Chr(13)&Chr(13)	  
			
			'If meses = 1 then
'				mensagem=mensagem&"Parcela em aberto do m&ecirc;s de "&desc_meses&" "	
'			else
'				mensagem=mensagem&"Parcelas em aberto dos meses: "&desc_meses&" "			
'			end if		
'				
'			mensagem=mensagem&"do ano letivo de "&session("ano_letivo")&Chr(13)&Chr(13)
			mensagem=Replace(mensagem&mensagem_padrao,Chr(13),"<BR>")
								
%>
    
    
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
			  <td height="30" colspan="3" valign="top" class="form_dado_texto">&nbsp;</td>
			  </tr>
			<tr>
				<td width="8%" height="30" valign="top" class="tb_subtit">Para
					<input name="ck_email" type="hidden" id="ck_email" value="<%response.write(ck_email)%>"></td>
				<td width="0%" valign="top" class="form_dado_texto">&nbsp;</td>
				<td width="92%" valign="top" class="form_dado_texto"><%response.Write(end_email)%></td>
			</tr>
			<tr>
				<td height="30" valign="top" class="tb_subtit">Assunto
					<input name="tipo_email" type="hidden" id="tipo_email" value="<%response.write(tipo_email)%>"></td>
				<td valign="top" class="form_dado_texto">&nbsp;</td>
				<td valign="top" class="form_dado_texto"><%response.Write(assunto)%></td>
			</tr>
			<tr>
				<td height="30" valign="top" class="tb_subtit">Mensagem</td>
				<td valign="top" class="form_dado_texto">&nbsp;</td>
				<td valign="top" class="form_dado_texto">
<%response.Write(replace(mensagem, Chr(13),"<BR>"))%></td>
			</tr>
			<tr>
				<td height="30">&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
		</table>
<%
		end if		
	next	
%>        
        </td></tr><tr>
                <tr>
                	<td height="15" bgcolor="#FFFFFF"><hr></td>
                	</tr>
                	<td height="15" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                		<tr>
                			<td width="33%" align="center"><input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','altera.asp?opt=vt&cod_cons=<%response.Write(e_cod)%>');" value="Voltar"></td>
                			<td width="34%">&nbsp;</td>
                			<td width="33%" align="center"><input name="button" type="submit" class="botao_prosseguir" id="button" value="Enviar"></td>
                			</tr>
                		</table></td>
                				</table></form></td>
	</tr>
</table>
</td>
          </tr>
        </table>
              </td>
          </tr>
		  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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
'response.redirect("../../../../inc/erro.asp")
end if
%>