<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%
opt = request.QueryString("opt")
nvg = session("chave")
co_usr = session("co_user")
Session("arquivos")=request.QueryString("arq")
Session("upl_total")=request.QueryString("upl")
tipo_arquivo_upload = request.QueryString("tp")
ano_letivo_wf = Session("ano_letivo_wf")
chave=nvg
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo_wf
nivel=4
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano_info=nivel&"-"&chave&"-"&ano_letivo_wf

dia_de= Session("dia_de")
mes_de= Session("mes_de")
dia_ate=Session("dia_ate")
mes_ate=Session("mes_ate")
unidade=Session("unidade")
curso=Session("curso")
etapa=Session("etapa")
turma=Session("turma")
tit=Session("tit")
check_status=Session("check_status")
tp_doc=session("tipo_arquivo")


Session("dia_de")=dia_de
Session("mes_de")=mes_de
Session("dia_ate")=dia_ate
Session("mes_ate")=mes_ate
Session("unidade")=unidade
Session("curso")=curso
Session("etapa")=etapa
Session("turma")=turma
Session("tit")=tit
Session("check_status")=check_status
session("tipo_arquivo") =tp_doc

if transicao = "S" then
	area="wd"
	site_escola="www.simplynet.com.br/wd/"&ambiente_escola&"/wf/apoio/ctl/pub/"
else
	if left(ambiente_escola,5)= "teste" then
		area="wdteste"
		site_escola="www.simplynet.com.br/"&area&"/"&ambiente_escola&"/wf/apoio/ctl/pub"
	else
		area="wd"
		site_escola="www.simplynet.com.br/wd/"&ambiente_escola&"/wf/apoio/ctl/pub"
	end if
end if

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1

    	Set CON_WF = Server.CreateObject("ADODB.Connection") 
		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_WF.Open ABRIR_WF	


 call navegacao (CON,chave,nivel)
navega=Session("caminho")	

 Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Autoriz_Usuario_Grupo Where CO_Usuario = "&co_usr
		RS2.Open SQL2, CON
		
if RS2.EOF then

else		
co_grupo=RS2("CO_Grupo")
End if

ano_expira = DatePart("yyyy", now) 
mes_expira = DatePart("m", now) 
dia_expira = DatePart("d", now) 


if ano_letivo_wf=ano then
	data_expira=dia_expira&"/"&mes_expira&"/"&ano_expira
else

	data_expira="01/01/"&ano_letivo_wf
end if	

		Set RS_doc = Server.CreateObject("ADODB.Recordset")
		SQL_doc = "SELECT * FROM TB_Tipo_Pasta_Doc where ((DA_Expira NOT BETWEEN #01/01/1900# AND #"&data_expira&"#) AND IN_Expira= TRUE) or IN_Expira= FALSE order by NO_Pasta Asc"		
		RS_doc.Open SQL_doc, CON0

		if RS_doc.eof then
		else
			qtd=0
			while not RS_doc.eof
				cod_tp_doc=RS_doc("CO_Pasta_Doc")
				nom_tp_doc=RS_doc("NO_Pasta")	
				strReplacement = Server.URLEncode(nom_tp_doc)
				strReplacement = replace(strReplacement,"+"," ")
				strReplacement = replace(strReplacement,"%27","´")
				strReplacement = replace(strReplacement,"%27","'")
				strReplacement = replace(strReplacement,"%C0,","À")
				strReplacement = replace(strReplacement,"%C1","Á")
				strReplacement = replace(strReplacement,"%C2","Â")
				strReplacement = replace(strReplacement,"%C3","Ã")
				strReplacement = replace(strReplacement,"%C9","É")
				strReplacement = replace(strReplacement,"%CA","Ê")
				strReplacement = replace(strReplacement,"%CD","Í")
				strReplacement = replace(strReplacement,"%D3","Ó")
				strReplacement = replace(strReplacement,"%D4","Ô")
				strReplacement = replace(strReplacement,"%D5","Õ")
				strReplacement = replace(strReplacement,"%DA","Ú")
				strReplacement = replace(strReplacement,"%DC","Ü")	
				strReplacement = replace(strReplacement,"%E1","à")
				strReplacement = replace(strReplacement,"%E1","á")
				strReplacement = replace(strReplacement,"%E2","â")
				strReplacement = replace(strReplacement,"%E3","ã")
				strReplacement = replace(strReplacement,"%E7","ç")
				strReplacement = replace(strReplacement,"%E9","é")
				strReplacement = replace(strReplacement,"%EA","ê")
				strReplacement = replace(strReplacement,"%ED","í")
				strReplacement = replace(strReplacement,"%F3","ó")
				strReplacement = replace(strReplacement,"F4","ô")
				strReplacement = replace(strReplacement,"F5","õ")
				strReplacement = replace(strReplacement,"%FA","ú")
				nom_tp_doc = replace(strReplacement,"%FC","ü")				
				if qtd=0 then
					vetor_tp_doc=cod_tp_doc&"!$!"&nom_tp_doc
				else
					vetor_tp_doc=vetor_tp_doc&"$!$"&cod_tp_doc&"!$!"&nom_tp_doc			
				end if
			qtd=qtd+1	
			RS_doc.movenext
			WEND
		end if	
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="file:../../../../img/mm_menu.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
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
function submitforminterno()  
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
<!--

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>								   
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../../img/menu_r1_c2_f3.gif','../../../../img/menu_r1_c2_f2.gif','../../../../img/menu_r1_c2_f4.gif','../../../../img/menu_r1_c4_f3.gif','../../../../img/menu_r1_c4_f2.gif','../../../../img/menu_r1_c4_f4.gif','../../../../img/menu_r1_c6_f3.gif','../../../../img/menu_r1_c6_f2.gif','../../../../img/menu_r1_c6_f4.gif','../../../../img/menu_r1_c8_f3.gif','../../../../img/menu_r1_c8_f2.gif','../../../../img/menu_r1_c8_f4.gif','../../../../img/menu_direita_r2_c1_f3.gif','../../../../img/menu_direita_r2_c1_f2.gif','../../../../img/menu_direita_r2_c1_f4.gif','../../../../img/menu_direita_r4_c1_f3.gif','../../../../img/menu_direita_r4_c1_f2.gif','../../../../img/menu_direita_r4_c1_f4.gif','../../../../img/menu_direita_r6_c1_f3.gif','../../../../img/menu_direita_r6_c1_f2.gif','../../../../img/menu_direita_r6_c1_f4.gif')">
<% call cabecalho (nivel)
	  %>
<table width="1000" height="613" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> </td>
  </tr>
<%				
if opt = "f" then%>
  <tr> 
    <td height="10"> 
      <%
		call mensagens(nivel,52,2,0)
%>
    </td>
  </tr>
  <% 	end if 

if opt = "err" then%>
  <tr> 
    <td height="10"> 
      <%
		call mensagens(nivel,53,1,0)
%>
    </td>
  </tr>
  <%elseif opt = "err1" then%>
  <tr> 
    <td height="10"> 
      <%
		call mensagens(nivel,62,1,0)
%>
    </td>
  </tr>
  <% 	end if 

%>
  <tr> 
    <td height="10"> 
      <%	call mensagens(nivel,51,0,0) 
	  
Session("file1") =""
Session("file2") =""
Session("file3") = ""
Session("file4") = ""
Session("file5") = "" 
Session("upl_total") = ""
%>
    </td>
  </tr>
  <tr class="tb_tit"> 
    <td height="15" class="tb_tit">Publique os arquivos 
      <input name="co_grupo" type="hidden" id="co_grupo" value="<% = co_grupo %>"></td>
  </tr>
  <tr> 
    <td valign="top"> <iframe scrolling="no" src ="http://<%response.Write(site_escola)%>/sndocs/envia.asp?opt=<%response.Write(ano_letivo_wf)%>&tp=<%response.Write(vetor_tp_doc)%>&env=<%response.Write(ambiente_escola)%>&tp_sel=<%response.Write(tipo_arquivo_upload)%>" frameborder ="0" width="100%" height="170"> 
      </iframe> </td>
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
response.redirect("../../../../inc/erro.asp")
end if
%>