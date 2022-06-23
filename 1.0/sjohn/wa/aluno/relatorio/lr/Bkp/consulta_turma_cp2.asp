<%On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->

<%
nivel=4

opt=request.QueryString("opt")

autoriza=Session("autoriza")
Session("autoriza")=autoriza

permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
if opt="vt" then
curso = request.QueryString("c")
unidade = request.QueryString("u")
else
curso = request.Form("curso")
unidade = request.Form("unidade")
end if


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CONG = Server.CreateObject("ADODB.Connection") 
		ABRIRG = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONG.Open ABRIRG		


		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Abreviado_Curso")

 call navegacao (CON,chave,nivel)
navega=Session("caminho")
	%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<% call EscreveFuncaoJavaScriptTurma (CONG,CON0) %>
<script type="text/javascript" src="../js/global.js"></script>
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
}  function checksubmit()
{
  if (document.inclusao.etapa.value == "999999")
  {    alert("Por favor, selecione uma etapa!")
    document.inclusao.etapa.focus()
    return false
  }         	     
  return true
}
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
<script language="JavaScript">
function atualiza_nota (form) {
switch (form.etapa.options[form.etapa.selectedIndex].value) {
case 'f0':
form.campo.length=0;
form.campo.options[0] = new Option('','999999');
form.campo.options[1] = new Option('N1','N1');
form.campo.options[2] = new Option('N2','N2');
form.campo.options[3] = new Option('N3','N3');
form.campo.options[4] = new Option('BÔN','BON');
form.campo.options[5] = new Option('REC','REC');
break;

case 'f1':
form.campo.length=0;
form.campo.options[0] = new Option('','999999');
form.campo.options[1] = new Option('N1','N1');
form.campo.options[2] = new Option('N2','N2');
form.campo.options[3] = new Option('N3','N3');
form.campo.options[4] = new Option('BÔN','BON');
form.campo.options[5] = new Option('REC','REC');
break;

case 'f2':
form.campo.length=0;
form.campo.options[0] = new Option('','999999');
form.campo.options[1] = new Option('N1','N1');
form.campo.options[2] = new Option('N2','N2');
form.campo.options[3] = new Option('N3','N3');
form.campo.options[4] = new Option('BÔN','BON');
form.campo.options[5] = new Option('REC','REC');
break;

case 'f3':
form.campo.length=0;
form.campo.options[0] = new Option('','999999');
form.campo.options[1] = new Option('N1','N1');
form.campo.options[2] = new Option('N2','N2');
form.campo.options[3] = new Option('N3','N3');
form.campo.options[4] = new Option('BÔN','BON');
form.campo.options[5] = new Option('REC','REC');
break;

case 'f4':
form.campo.length=0;
form.campo.options[0] = new Option('','999999');
form.campo.options[1] = new Option('N1','N1');
form.campo.options[2] = new Option('N2','N2');
form.campo.options[3] = new Option('N3','N3');
form.campo.options[4] = new Option('BÔN','BON');
form.campo.options[5] = new Option('REC','REC');
break;

case 'f5':
form.campo.length=0;
form.campo.options[0] = new Option('','999999');
form.campo.options[1] = new Option('A1','A1');
form.campo.options[2] = new Option('A2','A2');
form.campo.options[3] = new Option('B1','B1');
form.campo.options[4] = new Option('B2','B2');
form.campo.options[5] = new Option('AV1','AV1');
form.campo.options[6] = new Option('AV2','AV2');
form.campo.options[7] = new Option('N3','N3');
form.campo.options[8] = new Option('N4','N4');
form.campo.options[9] = new Option('BÔN','BON');
form.campo.options[10] = new Option('REC','REC');
break;

case 'f6':
form.campo.length=0;
form.campo.options[0] = new Option('','999999');
form.campo.options[1] = new Option('A1','A1');
form.campo.options[2] = new Option('A2','A2');
form.campo.options[3] = new Option('B1','B1');
form.campo.options[4] = new Option('B2','B2');
form.campo.options[5] = new Option('AV1','AV1');
form.campo.options[6] = new Option('AV2','AV2');
form.campo.options[7] = new Option('N3','N3');
form.campo.options[8] = new Option('N4','N4');
form.campo.options[9] = new Option('BÔN','BON');
form.campo.options[10] = new Option('REC','REC');
break;

case 'f7':
form.campo.length=0;
form.campo.options[0] = new Option('','999999');
form.campo.options[1] = new Option('A1','A1');
form.campo.options[2] = new Option('A2','A2');
form.campo.options[3] = new Option('B1','B1');
form.campo.options[4] = new Option('B2','B2');
form.campo.options[5] = new Option('AV1','AV1');
form.campo.options[6] = new Option('AV2','AV2');
form.campo.options[7] = new Option('N3','N3');
form.campo.options[8] = new Option('N4','N4');
form.campo.options[9] = new Option('BÔN','BON');
form.campo.options[10] = new Option('REC','REC');
break;

case 'f8':
form.campo.length=0;
form.campo.options[0] = new Option('','999999');
form.campo.options[1] = new Option('TES','TES');
form.campo.options[2] = new Option('PRO','PRO');
form.campo.options[3] = new Option('N3','N3');
form.campo.options[4] = new Option('BÔN','BON');
form.campo.options[5] = new Option('REC','REC');
break;

case 'f55':
form.campo.length=0;
form.campo.options[0] = new Option('','999999');
form.campo.options[1] = new Option('A1','A1');
form.campo.options[2] = new Option('A2','A2');
form.campo.options[3] = new Option('B1','B1');
form.campo.options[4] = new Option('B2','B2');
form.campo.options[5] = new Option('AV1','AV1');
form.campo.options[6] = new Option('AV2','AV2');
form.campo.options[7] = new Option('N3','N3');
form.campo.options[8] = new Option('N4','N4');
form.campo.options[9] = new Option('BÔN','BON');
form.campo.options[10] = new Option('REC','REC');
break;

case 'f66':
form.campo.length=0;
form.campo.options[0] = new Option('','999999');
form.campo.options[1] = new Option('A1','A1');
form.campo.options[2] = new Option('A2','A2');
form.campo.options[3] = new Option('B1','B1');
form.campo.options[4] = new Option('B2','B2');
form.campo.options[5] = new Option('AV1','AV1');
form.campo.options[6] = new Option('AV2','AV2');
form.campo.options[7] = new Option('N3','N3');
form.campo.options[8] = new Option('N4','N4');
form.campo.options[9] = new Option('BÔN','BON');
form.campo.options[10] = new Option('REC','REC');
break;

case 'f77':
form.campo.length=0;
form.campo.options[0] = new Option('','999999');
form.campo.options[1] = new Option('A1','A1');
form.campo.options[2] = new Option('A2','A2');
form.campo.options[3] = new Option('B1','B1');
form.campo.options[4] = new Option('B2','B2');
form.campo.options[5] = new Option('AV1','AV1');
form.campo.options[6] = new Option('AV2','AV2');
form.campo.options[7] = new Option('N3','N3');
form.campo.options[8] = new Option('N4','N4');
form.campo.options[9] = new Option('BÔN','BON');
form.campo.options[10] = new Option('REC','REC');
break;

case 'f88':
form.campo.length=0;
form.campo.options[0] = new Option('','999999');
form.campo.options[1] = new Option('TES','TES');
form.campo.options[2] = new Option('PRO','PRO');
form.campo.options[3] = new Option('N3','N3');
form.campo.options[4] = new Option('BÔN','BON');
form.campo.options[5] = new Option('REC','REC');
break;


case 'm1':
form.campo.length=0;
form.campo.options[0] = new Option('','999999');
form.campo.options[1] = new Option('TES','TES');
form.campo.options[2] = new Option('PRO','PRO');
form.campo.options[3] = new Option('N3','N3');
form.campo.options[4] = new Option('BÔN','BON');
form.campo.options[5] = new Option('REC','REC');
break;

case 'm2':
form.campo.length=0;
form.campo.options[0] = new Option('','999999');
form.campo.options[1] = new Option('TES','TES');
form.campo.options[2] = new Option('PRO','PRO');
form.campo.options[3] = new Option('N3','N3');
form.campo.options[4] = new Option('BÔN','BON');
form.campo.options[5] = new Option('REC','REC');
break;

case 'm3':
form.campo.length=0;
form.campo.options[0] = new Option('','999999');
form.campo.options[1] = new Option('TES','TES');
form.campo.options[2] = new Option('PRO','PRO');
form.campo.options[3] = new Option('N3','N3');
form.campo.options[4] = new Option('BÔN','BON');
form.campo.options[5] = new Option('REC','REC');
break;}}
</script>   
</head> 
<body link="#CC9900" vlink="#CC9900" background="../../../../img/fundo.gif" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr>                    
            <td height="10" class="tb_caminho"> <font class="style-caminho">
              <%
	  response.Write(navega)

%>
              </font>
	</td>
  </tr>
<%If opt="err2" then%>
                <tr>                   
    <td height="10"> 
      <%	call mensagens(4,4,1,0) 

%>
    </td>
	</tr>
<%end if%>
                <tr>                   
    <td height="10"> 
      <%	call mensagens(4,2,0,0) 

%>
    </td>
                </tr>				  				  


          <tr> 
            <td valign="top"> 
              <form name="inclusao" method="post" action="mapa.asp?opt=gera" onSubmit="return checksubmit()">
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>


                  <tr class="tb_tit"
> 
                    
            <td height="15" class="tb_tit"
>Grade de Aulas </td>
                  </tr>
                  <tr> 
                    <td><table width="1000" border="0" cellspacing="0">
                <tr> 
                  <td width="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                  <td width="248" class="tb_subtit"> 
                    <div align="center">UNIDADE 
                    </div></td>
                  <td width="247" class="tb_subtit"> 
                    <div align="center">CURSO 
                    </div></td>
                  <td width="248" class="tb_subtit"> 
                    <div align="center">ETAPA 
                    </div></td>
                  <td width="247" class="tb_subtit"> 
                    <div align="center">TURMA</div></td>
                </tr>
                <tr> 
                  <td width="10"> </td>
                  <td width="248"> 
                    <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_unidade)%>
                      <input name="unidade" type="hidden" id="unidade" value="<% = unidade %>">
                      </font></div></td>
                  <td width="247"> 
                    <div align="center"> <font class="form_dado_texto"> 
                      <%
response.Write(no_curso)%>
                      <input type="hidden" name="curso" value="<% = curso %>">
                      </font></div></td>
                  <td width="248"> 
                    <div align="center"> <font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      <%
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Unidade_Possui_Etapas where CO_Curso ='"& curso &"' AND NU_Unidade="& unidade 
		RS2.Open SQL2, CON0
		
%>
                      <select name="etapa" class="borda" onChange="javascript:atualizaturma(this.form);javascript:atualiza_nota(this.form);">
                        <option value="999999" selected></option>
                        <%while not RS2.EOF
co_etapa= RS2("CO_Etapa")

		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' And CO_Curso ='"& curso &"'" 
		RS3.Open SQL3, CON0
		


no_etapa=RS3("NO_Etapa")

 %>
                        <option value="<%=co_etapa%>"> 
                        <%response.Write(no_etapa)%>
                        </option>
                        <%RS2.MOVENEXT
WEND
%>
                      </select>
                      </font></div></td>
                  <td width="247"> 
                    <div align="center"> 
                      <select name="turma" class="borda">
                        <option value="999999" selected></option>
                      </select>
                    </div></td>
                </tr>
              </table></td>
                  </tr>
                  <tr bgcolor="#FFFFFF"> 
                    <td bgcolor="#FFFFFF"><table width="95%" border="0" align="center" cellspacing="0" bgcolor="#FFFFFF">
                <tr> 
                  <td width="17%" class="tb_subtit"> <div align="right"> Data da Reuni&atilde;o:</div></td>
                  <td width="83%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <input name="dr" type="text" class="borda" id="dr" size="3">
                    / 
                    <select name="mr" class="borda" id="mr">
                      <option value="Janeiro">Janeiro</option>
                      <option value="Fevereiro">Fevereiro</option>
                      <option value="Mar&ccedil;o">Mar&ccedil;o</option>
                      <option value="Abril">Abril</option>
                      <option value="Maio">Maio</option>
                      <option value="Junho">Junho</option>
                      <option value="Julho">Julho</option>
                      <option value="Agosto">Agosto</option>
                      <option value="Setembro">Setembro</option>
                      <option value="Outubro">Outubro</option>
                      <option value="Novembro">Novembro</option>
                      <option value="Dezembro">Dezembro</option>
                    </select>
                    / 
                    <input name="ar" type="text" class="borda" id="ar" size="5">
                    </font></td>
                </tr>
                <tr> 
                  <td class="tb_subtit"> <div align="right"> Motivo:</div></td>
                  <td><font class="form_dado_texto"> 
                    <input name="motivo" type="text" class="borda" id="motivo" size="100">
                    </font></td>
                </tr>
                <tr> 
                  <td colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                </tr>
                <tr> 
                  <td colspan="2"><div align="center"> 
                      <input name="Submit" type="submit" class="borda_bot2" value="Criar Relat&oacute;rio">
                    </div></td>
                </tr>
              </table></td>
                  </tr>
                </table>
              </form></td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

</body>
<script type="text/javascript">
<!--
  initInputHighlightScript();
//-->
</script>
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