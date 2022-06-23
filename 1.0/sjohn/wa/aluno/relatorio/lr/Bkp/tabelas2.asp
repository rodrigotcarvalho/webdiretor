<%On Error Resume Next%>
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/caminhos.asp"-->

<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/caminhos.asp"-->


<%opt = request.QueryString("opt")

if opt = "err2" then

curso = request.querystring("curso")
unidade = request.querystring("unidade")
grade = request.querystring("grade")
elseif opt = "vt" then

curso = request.querystring("curso")
unidade = request.querystring("unidade")

else

codigo = request.form("cod")
nome_prof = request.form("nome_prof")
co_usr_prof = request.form("co_usr_prof")
curso = request.Form("curso")
unidade = request.Form("unidade")
grade = request.Form("grade")
end if
ano_letivo = session("ano_letivo")

'if curso = "0" then

'response.Redirect("consulta_turma_cp3.asp?opt=direto&curso=0&or=02&unidade="& unidade&"&grade="& grade&"")


'else



id0 = " > <a href='tabelas.asp?or=01&volta=1' class='caminho' target='_self'>Enitir Lista de Reunião</a>"
id1 = " > <a href='tabelas.asp?or=01&volta=1' class='caminho' target='_self'>Seleciona Unidade</a>"
id2 = " > Seleciona Etapa"




		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0


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
<title>Web Acad&ecirc;mico</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
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
function submitform()  
{
   var f=document.forms[0]; 
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
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<table width="1000" border="0" align="center" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td><div align="center"><table width="1000" border="0">
  <tr> 
    <td><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="../inicio.asp" target="_parent" class="caminho">Web 
      Acad&ecirc;mico</a> 
      <%

	  response.Write(origem&id0&id1&id2)
%>
      
      </strong> </font></td>
  </tr>
</table>
<br>
<table width="1000" border="0" cellspacing="0">
  <tr> 
    <td width="219" valign="top">
<table width="100%" border="0" cellspacing="0">
<%if opt = "err2" then%>
<tr>
          <td>     
<%
		call mensagens(23,1,0)
%>
</td>
        </tr>        
<%end if%><tr>
          <td>   		 
<%	call mensagens(140,0,0) 

%>
</td>
        </tr>
        <tr>
          <td>
<%
	call ultimo(0) 
%>		  		  
		  </td>
        </tr>
      </table>
      
    </td>
    <td width="785" valign="top"> 	  
      <form name="inclusao" method="post" action="mapa.asp?or=01&opt=gera" onSubmit="return checksubmit()">
        <table width="770" border="0" align="right" cellspacing="0" class="tb_corpo"
>
          <tr class="tb_tit"
> 
            <td width="653" height="15" class="tb_tit"
><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Ano 
              Letivo:</strong></font><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%response.Write(ano_letivo)
%>
              <input name="ano_letivo" type="hidden" id="ano_letivo" value="<%=ano_letivo%>">
              </font> </td>
          </tr>
          <tr class="tb_tit"
> 
            <td height="15" bgcolor="#FFFFFF"> </td>
          </tr>
          <tr class="tb_tit"
> 
            <td height="15" class="tb_tit"
><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
              Segmento</strong></font></td>
          </tr>
          <tr> 
            <td><table width="770" border="0" cellspacing="0">
                <tr> 
                  <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                  <td class="tb_subtit"> <div align="center"><font color="#FF6600"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">UNIDADE 
                      </font></strong></font></div></td>
                  <td class="tb_subtit"> <div align="center"><font color="#FF6600"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">CURSO 
                      </font></strong></font></div></td>
                  <td class="tb_subtit"> <div align="center"><font color="#FF6600"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">ETAPA 
                      </font></strong></font></div></td>
                  <td class="tb_subtit"> <div align="center"><font color="#FF6600"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">TURMA 
                      </font></strong></font></div></td>
                </tr>
                <tr> 
                  <td width="10"> </td>
                  <td width="300"> <div align="center"> <font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      <%response.Write(no_unidade)%>
                      <input name="unidade" type="hidden" id="unidade" value="<% = unidade %>">
                      </font></div></td>
                  <td width="70"> <div align="center"> <font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      <%
response.Write(no_curso)%>
                      <input type="hidden" name="curso" value="<% = curso %>">
                      </font></div></td>
                  <td width="300"> <div align="center"> <font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      <%
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Unidade_Possui_Etapas where CO_Curso ='"& curso &"' AND NU_Unidade="& unidade 
		RS2.Open SQL2, CON0
		
'co_etapa= RS2("CO_Etapa")
'if curso = 0 then
'etapa = ""

'response.Write("sem etapa  <input type='hidden' name='etapa' value="& etapa &">")

'else
%>
                      <select name="etapa" class="borda">
                        <option value="999999" selected></option>
                        <%while not RS2.EOF
co_etapa= RS2("CO_Etapa")

		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' and CO_Curso ='"& curso &"'"  
		RS3.Open SQL3, CON0
		
'if RS3.EOF THEN
'no_etapa="SEM ETAPA"
%>
                        <%

'else
no_etapa=RS3("NO_Etapa")

 %>
                        <option value="<%=co_etapa%>"> 
                        <%response.Write(no_etapa)%>
                        </option>
                        <%RS2.MOVENEXT
WEND
%>
                      </select>
                      <%'end if %>
                      </font></div></td>
                  <td width="90"> <div align="center"> 
                      <select name="turma" class="borda">
                        <option value="0" selected></option>
                        <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Turma where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND NU_Unidade="& unidade 
		RS4.Open SQL4, CON0

while not RS4.EOF
CO_Turma= RS4("CO_Turma")%>
                        <option value="<%=CO_Turma%>"> 
                        <%response.Write(CO_Turma)%>
                        </option>
                        <%RS4.MOVENEXT
WEND%>
                      </select>
                    </div></td>
                </tr>
              </table></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td bgcolor="#FFFFFF">
			<table width="95%" border="0" align="center" cellspacing="0" bgcolor="#FFFFFF">
                <tr> 
                  <td width="17%"> 
                    <div align="right"><font color="#FF6600"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
                      da Reuni&atilde;o:</font></strong></font></div></td>
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
                  <td> 
                    <div align="right"><font color="#FF6600"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Motivo:</font></strong></font></div></td>
                  <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
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
              </table>
			
			</td>
          </tr>
        </table>
</form>
	
		</td>
  </tr>
</table>
<p>&nbsp;</p>
<table width="1000" border="0" cellspacing="0">
  <tr>
    <td width="238">&nbsp;</td>
    <td width="770" class="tb_voltar"
><font color="#669999" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="../alunos.asp" target="_parent" class="voltar1">&lt; 
      Voltar para o menu Alunos</a></strong></font></td>
  </tr>
</table>
<p align="center">&nbsp;</p></div></td>
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
response.redirect("../../../../inc/erro.asp")
end if
%>