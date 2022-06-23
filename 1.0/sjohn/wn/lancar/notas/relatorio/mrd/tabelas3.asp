<%On Error Resume Next%>
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/caminhos.asp"-->

<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/caminhos.asp"-->


<%opt = request.QueryString("opt")
co_usr = session("co_usr")
if opt = "err1" OR opt = "err2" OR opt = "err3"then
dd=request.querystring("dd")
dados = split(dd,"_")
ano_letivo = dados(0)
curso = dados(1)
unidade = dados(2)
co_prof = dados(3)
else
ano_letivo = request.form("ano_letivo")
co_prof = request.Form("co_prof")
curso = request.Form("curso")
unidade = request.Form("unidade")
etapa= request.form("etapa")
turma= request.form("turma")
end if
session("ano_letivo")=ano_letivo

'if curso = "0" then

'response.Redirect("direciona.asp?opt=direto&curso=0&or=02&ano="& ano_letivo&"&unidade="& unidade&"&grade="& grade&"")


'else


id0 = " > <a href='tabelas.asp?or=02&ano="&ano_letivo&"' class='linkum' target='_parent'>Mapa de Resultados</a>"
id1 = " > Selecionando Turma"



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


	%>
<html>
<head>
<title>Lan&ccedil;ar Notas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../estilos.css" rel="stylesheet" type="text/css">
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
function checksubmit()
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
</head> 
<body link="#6699CC" vlink="#6699CC" alink="#6699CC" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<% call cabecalho(nivel) 

%>
<table width="1000" height="670" border="0" align="center" cellspacing="0" bgcolor="#FFFFFF">
  <tr> 
    <td valign="top"> <div align="center"> 
        <table width="1000" border="0" class="tb_caminho">
          <tr> 
            <td><font color="#FFFF33" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="../inicio.asp" target="_parent" class="caminho">Web 
              Notas</a> 
              <%

	  response.Write(origem&id0&id1)
%>
              </font></td>
          </tr>
        </table>
        <br>
        <table width="1000" border="0" cellspacing="0">
          <tr> 
            <td width="219" valign="top"> <table width="100%" border="0" cellspacing="0">
                <%if opt = "err1" then%>
                <tr> 
                  <td> 
                    <%
		call mensagens(231,1,0)
%>
                  </td>
                </tr>
                <%end if%>
                <%if opt = "err2" then%>
                <tr> 
                  <td> 
                    <%
		call mensagens(232,1,0)
%>
                  </td>
                </tr>
                <%end if%>
                <%if opt = "err3" then%>
                <tr> 
                  <td> 
                    <%
		call mensagens(233,1,0)
%>
                  </td>
                </tr>
                <%end if%>
                <tr> 
                  <td> 
                    <%	call mensagens(14,0,0) 

%>
                  </td>
                </tr>
                <tr> 
                  <td> 
                    <%
'	call ultimo(0) 
%>
                  </td>
                </tr>
              </table></td>
            <td width="770" valign="top"> <form name="inclusao" method="post" action="mapa.asp?or=02" onSubmit="return checksubmit()">
                <table width="770" border="0" align="right" cellspacing="0" class="tb_corpo">
                  <tr class="tb_tit"> 
                    <td height="15" class="tb_tit"><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Ano 
                      Letivo:</strong></font><font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      <%response.Write(ano_letivo)

%>
                      <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% =ano_letivo%>">
                      <input name="co_prof" type="hidden" id="co_prof" value="<% = co_prof%>">
                      <input name="co_usr" type="hidden" id="co_usr" value="<%=co_usr%>">
                      </font> </td>
                  </tr>
                  <tr class="tb_tit"> 
                    <td height="15" bgcolor="#FFFFFF"> </td>
                  </tr>
                  <tr class="tb_tit"> 
                    <td height="15" class="tb_tit"><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                      Grade de aulas</strong></font></td>
                  </tr>
                  <tr> 
                    <td><table width="770" border="0" cellspacing="0">
                        <tr> 
                          <td width="9"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                          <td width="127" class="tb_subtit"> <div align="center"><font color="#FF6600"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">UNIDADE 
                              </font></strong></font></div></td>
                          <td width="92" class="tb_subtit"> <div align="center"><font color="#FF6600"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">CURSO 
                              </font></strong></font></div></td>
                          <td width="160" class="tb_subtit"> <div align="center"><font color="#FF6600"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">ETAPA 
                              </font></strong></font></div></td>
                          <td width="40" class="tb_subtit"> <div align="center"><font color="#FF6600"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">TURMA 
                              </font></strong></font></div></td>
                          <td width="197" class="tb_subtit"><div align="center"><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>DISCIPLINA</strong></font></div></td>
                        </tr>
                        <tr> 
                          <td> </td>
                          <td> <div align="center"> <font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                              <%response.Write(no_unidade)%>
                              <input name="unidade" type="hidden" id="unidade" value="<% = unidade %>">
                              </font></div></td>
                          <td> <div align="center"> <font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                              <%
response.Write(no_curso)%>
                              <input type="hidden" name="curso" value="<% = curso %>">
                              </font></div></td>
                          <td> <div align="center"> 
                              <input name="etapa" type="hidden" id="etapa" value="<%=etapa%>">
                              <font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                              <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& etapa &"'" 
		RS3.Open SQL3, CON0
no_etapa=RS3("NO_Etapa")
response.Write(no_etapa)%>
                              </font></div></td>
                          <td> <div align="center"> 
                              <input name="turma" type="hidden" id="turma" value="<%=turma%>">
                              <font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                              <%
response.Write(turma)%>
                              </font></div></td>
                          <td> <div align="center"> <font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                              <%
if curso = 0 then
mat_prin = ""
response.Write("sem disciplina  <input type='hidden' name='mat_prin' value='9999990'>")

else
                       
		Set RSG = Server.CreateObject("ADODB.Recordset")
		SQLG = "SELECT * FROM TB_Da_Aula where CO_Professor ="& co_prof&" and CO_Etapa = '"&etapa&"' AND NU_Unidade = "&unidade&" and CO_Curso = '"&curso&"' order by CO_Materia_Principal"
		RSG.Open SQLG, CONG
		
IF RSG.EOF THEN

RESPONSE.Write("Sem disciplinas cadastradas. Procure seu Coordenador.")

ELSE%>
                              <select name="mat" class="borda" onChange="MM_callJS('submitform()')">
                                <option value="999999" selected></option>
                                <%
co_mat_check="9999990"
while not RSG.EOF
co_mat= RSG("CO_Materia_Principal")
if co_mat = co_mat_check then
RSG.MOVENEXT
else
		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Materia where CO_Materia ='"& co_mat &"'"
		RS7.Open SQL7, CON0
		
		no_mat= RS7("NO_Materia")
%>
                                <option value="<%=co_mat%>"> 
                                <%response.Write(no_mat)%>
                                </option>
                                <%
co_mat_check = co_mat
RSG.MOVENEXT
end if
WEND%>
                              </select>
                              <%end if 
  end if %>
                              </font> </div></td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr bgcolor="#FFFFFF"> 
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>
                </table>
              </form></td>
          </tr>
        </table>
        <table width="1000" border="0" cellspacing="0">
          <tr> 
            <td width="219">&nbsp;</td>
            <td width="770" class="tb_voltar"><font color="#669999" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="tabelas.asp?or=01&ano=<%=ano_letivo%>" target="_parent" class="voltar1">&lt; 
              Voltar para Lan&ccedil;ar Notas</a></strong></font></td>
          </tr>
        </table>
      </div></td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
</body>
<script type="text/javascript">
<!--
  initInputHighlightScript();
//-->
</script>
<%

'end if
%>
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
response.redirect("../inc/erro.asp")
end if
%>