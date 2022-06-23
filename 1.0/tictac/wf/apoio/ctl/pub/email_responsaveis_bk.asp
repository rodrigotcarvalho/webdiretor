<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes7.asp"-->
<% 

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

opt = request.QueryString("opt")
cod = session("cod_url")
session("cod_url")=cod
ori = request.QueryString("ori")
opt_a = request.QueryString("opt_a")
obr = request.QueryString("obr")
pasta = session("tipo_arquivo")
nome = 	session("tit_doc")
dados_msg=ori&"$$$"&obr&"$$$"&opt_a
envia_mensagem = "N"

Set CON = Server.CreateObject("ADODB.Connection") 
ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
CON.Open ABRIR

Set CON0 = Server.CreateObject("ADODB.Connection") 
ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
CON0.Open ABRIR0	

Set CON1 = Server.CreateObject("ADODB.Connection") 
ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
CON1.Open ABRIR1

Set CON2 = Server.CreateObject("ADODB.Connection") 
ABRIR2 = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
CON2.Open ABRIR2	

Set CON_WF = Server.CreateObject("ADODB.Connection") 
ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
CON_WF.Open ABRIR_WF	

if opt = "sim" then
	alunos=split(cod,",")
	total_email=0
	conta_destinatarios=0
	vetor_destinatarios=""
	
	
		Set RS_doc = Server.CreateObject("ADODB.Recordset")
		SQL_doc = "SELECT * FROM TB_Tipo_Pasta_Doc where CO_Pasta_Doc = "&pasta
		RS_doc.Open SQL_doc, CON0

		if RS_doc.eof then
		  nom_tp_doc = ""
		else
		  nom_tp_doc=RS_doc("NO_Pasta")	
		end if  	
		
		nome = nom_tp_doc&" - "&nome
			
	for a=0 to ubound(alunos)
		co_matric=alunos(a)
		
		Set RSF = Server.CreateObject("ADODB.Recordset")
		'sqlF = "select CO_Usuario from TB_RespxAluno where CO_Aluno="&co_matric&" and (TP_Resp='F' or TP_Resp='P')"
		sqlF = "select CO_Usuario from TB_RespxAluno where CO_Aluno="&co_matric&" and TP_Resp='P'"
		set RSF = CON_wf.Execute (sqlF)		
		'response.Write("P "&sqlF&"<BR>")		
		while not RSF.EOF 
			resp=RSF("CO_Usuario")
		
			Set RSFM = Server.CreateObject("ADODB.Recordset")
			sqlFM = "select TX_EMail_Usuario,IN_Aut_email, ST_Usuario from TB_Usuario where CO_Usuario="&resp			
			set RSFM = CON_wf.Execute (sqlFM)					
			
				'response.Write(sqlFM&"<BR>")				
		
'			Set RS = Server.CreateObject("ADODB.Recordset")
'			SQL = "SELECT TP_Resp_Fin, TP_Resp_Ped FROM TB_Alunos where CO_Matricula = "&co_matric
'			RS.Open SQL, CON1
'			
'			'resp_fin  = RS("TP_Resp_Fin")
'			resp_ped  = RS("TP_Resp_Ped")					
'			
'			Set RSc = Server.CreateObject("ADODB.Recordset")
'			'SQLc = "SELECT DISTINCT(TX_EMail) FROM TB_Contatos where CO_Matricula = "&co_matric&" AND TP_Contato IN ('PAI', 'MAE', '"&resp_fin&"', '"&resp_ped&"')"
'			SQLc = "SELECT DISTINCT(TX_EMail) FROM TB_Contatos where CO_Matricula = "&co_matric&" AND TP_Contato IN ('PAI', 'MAE', '"&resp_ped&"')"
'			RSc.Open SQLc, CON2					
'			
'			while not RSc.EOF 
'				email_contato=RSc("TX_EMail")
			if not RSFM.EOF	then
				email_contato=RSFM("TX_EMail_Usuario")
				aut_resp=RSFM("IN_Aut_email")
				situacao_resp=RSFM("ST_Usuario")										
				'response.Write(email_contato&"-"&aut_resp&"-"&situacao_resp&"<BR>")									
				if isnull(email_contato) or email_contato="" or InStr(email_contato,"@")=0 then		
				else
					email_contato=Replace(email_contato, " ", "")	
				
					if aut_resp=TRUE and situacao_resp="L" then		
						envia_mensagem = "S"
						if total_email=0 then
							vetor_destinatarios=co_matric&"#!#"&email_contato				
						else
							vetor_destinatarios=vetor_destinatarios&"#$#"&co_matric&"#!#"&email_contato				
						end if	
						total_email=total_email+1		
					end if									
				end if	
			end if							
'			RSc.MOVENEXT
'			wend	
		RSF.MOVENEXT
		wend												
	next					
			
'			response.Write(envia_mensagem&"-"&vetor_destinatarios)
'			response.End()		

	if envia_mensagem = "S" then	
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Email_Assunto where CO_Assunto=7"
		RS1.Open SQL1, CON0

		if RS1.eof then 
		else

			assunto=RS1("TX_Titulo_Assunto")	
					
			publico=split(vetor_destinatarios,"#$#")	

			for e= 0 to ubound(publico)	
				dados_aluno=split(publico(e),"#!#")						
				co_aluno=dados_aluno(0)							
				destinatario=dados_aluno(1)
				
				Set RS1b = Server.CreateObject("ADODB.Recordset")
				SQL1b = "SELECT * FROM TB_Email_Mensagem where CO_Email=7"
				RS1b.Open SQL1b, CON0					
				
				mensagem=RS1b("TX_Conteudo_Email")															


				personalizado = nome
				mensagem=replace (mensagem, "XXX", personalizado)

				'destinatario = "webdiretor@gmail.com"	
						
				Set objCDOSYSMail = Server.CreateObject("CDO.Message")
				Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration") 'objeto de configuração do CDO
				objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
				objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
				objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 90
				objCDOSYSCon.Fields.update
				Set objCDOSYSMail.Configuration = objCDOSYSCon
				objCDOSYSMail.From = email_suporte_escola
				objCDOSYSMail.To = destinatario
				'objCDOSYSMail.Cc = ""
				'objCDOSYSMail.Bcc = email_teste		
				'objCDOSYSMail.Bcc = "webdiretor@gmail.com"	
				'objCDOSYSMail.Bcc = "osmarpio@openlink.com.br"
				'objCDOSYSMail.Bcc = destinatario	
				objCDOSYSMail.Subject = assunto
				objCDOSYSMail.TextBody = mensagem
				objCDOSYSMail.Send 
				Set objCDOSYSMail = Nothing
				Set objCDOSYSCon = Nothing
		
				If Err <> 0 Then
					erro = "<b><font color='red'> Erro ao enviar a mensagem.</font></b><br>"
					erro = erro & "<b>Erro.Description:</b> " & Err.Description & "<br>"
					erro = erro & "<b>Erro.Number:</b> "      & Err.Number & "<br>"
					erro = erro & "<b>Erro.Source:</b> "      & Err.Source & "<br>"
					response.write erro
					response.End()
				End if						
			next							
		end if													
	end if
end if
if opt <> "ask" then
	if ori = "i" then
		response.Redirect("incluir.asp?opt=ok")
	elseif ori = "a" then
		response.Redirect("alterar.asp?opt=ok&c="&obr)
	end if   
end if

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
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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
//-->
</script>
</head>

<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
            
    <td height="10" colspan="5" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
    </td>
          </tr>
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,9714,0,dados_msg) %>
    </td>
			  </tr>			  
          <tr>
      <td valign="top">&nbsp;</td>
    </tr>
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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
response.redirect("../../../../inc/erro.asp")
end if
%>