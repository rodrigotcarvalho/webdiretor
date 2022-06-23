<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->

<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->


<%
nivel=4
nvg = session("chave")
opt=request.QueryString("opt")
cod_cons= request.QueryString("cod_cons")

ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=nvg
session("chave")=chave


cod_prof = request.form("cod_prof")
nome_prof = request.form("nome_prof")
co_usr_prof = request.form("co_usr_prof")
curso = request.Form("curso")
unidade = request.Form("unidade")
co_etapa= request.Form("etapa")
turma= request.Form("turma")
mat_prin = request.form("mat_prin")
tabela = request.Form("tabela")
coordenador= request.Form("coordenador")
grade = request.Form("grade")



		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0

 call navegacao (CON,chave,nivel)
navega=Session("caminho")
	%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../../../../js/global.js"></script>
<script language="JavaScript">
 window.history.forward(1);
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
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
        <td height="10" class="tb_caminho"><font class="style-caminho">
      <%
	  response.Write(navega)
	  %>
      </font></td>
	  </tr>
<%
if grade= "" then

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Abreviado_Curso")
%>
          <tr> 
            
    <td height="10"> 
      <%
	call mensagens(nivel,633,0,0) 
%>
    </td>
          </tr>
          <tr> 
            <td width="770" valign="top"> <form action="bd.asp?opt=inc" method="post" name="alteracao">
                
        <table width="1000" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo"
>
          <tr class="tb_tit"
> 
            <td width="653" height="15" class="tb_tit"
>Professor</td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="9%" height="30" class="tb_subtit"> <div align="right">C&Oacute;DIGO: 
                    </div></td>
                  <td width="11%" height="30"> <font class="form_dado_texto"> 
                    <input name="cod_prof" type="hidden" id="cod_prof" value="<%=cod_prof%>">
                    <%response.Write(cod_prof)%>
                    <input name="tp" type="hidden" id="tp" value="P">
                    <input name="acesso" type="hidden" id="acesso" value="2">
                    <input name="nome_prof" type="hidden" id="nome_prof" value="<% =nome_prof%>">
                    <input name="co_usr_prof" type="hidden" id="co_usr_prof" value="<% =co_usr_prof%>">
                    </font></td>
                  <td width="6%" height="30" class="tb_subtit"> <div align="right" >NOME: 
                    </div></td>
                  <td width="74%" height="30"> <font class="form_dado_texto"> 
                    <%response.Write(nome_prof)%>
                    </font> </td>
                </tr>
              </table></td>
          </tr>
          <tr class="tb_tit"
> 
            <td height="15" class="tb_tit"
>Grade de Aulas</td>
          </tr>
          <tr> 
            <td><table width="1000" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="13"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                  <td width="100" class="tb_subtit"> <div align="center"><strong>UNIDADE 
                      </strong></div></td>
                  <td width="100" class="tb_subtit"> <div align="center"><strong>CURSO 
                      </strong></div></td>
                  <td width="100" class="tb_subtit"> <div align="center"><strong>ETAPA 
                      </strong></div></td>
                  <td width="141" class="tb_subtit"> <div align="center"><strong>TURMA 
                      </strong></div></td>
                  <td width="202" class="tb_subtit"> <div align="center"><strong>DISCIPLINA</strong></div></td>
                  <td width="141" class="tb_subtit"> <div align="center"><strong>MODELO</strong></div></td>
                  <td width="203" class="tb_subtit"> <div align="center"><strong>COORDENADOR 
                      </strong></div></td>
                </tr>
                <tr> 
                  <td width="13"> </td>
                  <td width="100"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_unidade)%>
                      <input name="unidade" type="hidden" id="unidade" value="<% = unidade %>">
                      </font></div></td>
                  <td width="100"> <div align="center"> <font class="form_dado_texto"> 
                      <%
response.Write(no_curso)%>
                      <input type="hidden" name="curso" value="<% = curso %>">
                      </font></div></td>
                  <td width="100"> <div align="center"> <font class="form_dado_texto"> 
                      <%

		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' and CO_Curso ='"& curso &"'"  
		RS3.Open SQL3, CON0
		
if RS3.EOF THEN
no_etapa="sem etapa"
else
no_etapa=RS3("NO_Etapa")
end if
response.Write(no_etapa)%>
                      <input name="etapa" type="hidden" id="etapa" value="<% = co_etapa %>">
                      </font></div></td>
                  <td width="141"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(turma)%>
                      <input name="turma" type="hidden" id="turma" value="<% = turma%>">
                      </font></div></td>
                  <td width="202"> <div align="center"><font class="form_dado_texto"> 
                      <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Materia where CO_Materia ='"& mat_prin &"'" 
		RS4.Open SQL4, CON0
		
if RS4.EOF THEN
no_mat_prin="sem disciplina"
else
no_mat_prin=RS4("NO_Materia")
end if
response.Write(no_mat_prin)
%>
                      <input name="mat_prin" type="hidden" id="mat_prin" value="<% = mat_prin%>">
                      </font> </div></td>
                  <td width="141"> <div align="center"><font class="form_dado_texto"> 
                      <%
select case tabela
case "TB_NOTA_A" 
response.Write("Modelo A")
case "TB_NOTA_B" 
response.Write("Modelo B")
case "TB_NOTA_C"
response.Write("Modelo C")
end select

%>
                      <input name="tabela" type="hidden" id="tabela" value="<% = tabela%>">
                      </font> </div></td>
                  <td width="203"> <div align="center"><font class="form_dado_texto"> 
                      <%
		Set RS8 = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT * FROM TB_Usuario where CO_Usuario ="& coordenador 
				RS8.Open SQL8, CON
		
no_coordenador = RS8("NO_Usuario")							  
					  response.Write(no_coordenador)%>
                      <input name="coordenador" type="hidden" id="coordenador" value="<% = coordenador%>">
                      </font> </div></td>
                </tr>
              </table></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td bgcolor="#FFFFFF"> <table width="500" border="0" align="center" cellspacing="0">
                <tr> 
                  <td width="25%"> <div align="center"> 
                      <input name="alterar" type="submit" class="botao_cancelar" id="alterar" onClick="MM_goToURL('parent','altera.asp?cod_cons=<%=cod_prof%>&amp;nvg=<%=nvg%>');return document.MM_returnValue" value="Cancelar">
                    </div></td>
                  <td width="25%"><div align="center"> 
                      <input name="Submit" type="submit" class="botao_prosseguir" value="Confirmar">
                    </div></td>
                </tr>
              </table></td>
          </tr>
        </table>
              </form></td>
          </tr>
<%
else

session("cod_prof")= cod_prof
session("nome_prof")= nome_prof
session("co_usr_prof")= co_usr_prof
session("curso")= curso
session("unidade")= unidade
session("grade")= grade
%>
          <tr> 
            
    <td width="219" height="10" valign="top"> 
      <%
	call mensagens(nivel,634,0,0) 
%>
    </td>
                </tr>
                <tr> 
            <td width="770" valign="top"> <form action="bd.asp?opt=exc" method="post" name="alteracao">
                
              
        <table width="1000" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo"
>
          <tr class="tb_tit"
> 
            <td width="653" height="15" class="tb_tit"
>Professor</td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="9%" height="30" class="tb_subtit"> <div align="right">C&Oacute;DIGO: 
                    </div></td>
                  <td width="11%" height="30"> <font class="form_dado_texto"> 
                    <input name="cod_prof" type="hidden" id="cod_prof" value="<%=cod_prof%>">
                    <%response.Write(cod_prof)%>
                    <input name="tp" type="hidden" id="tp" value="P">
                    <input name="acesso" type="hidden" id="acesso" value="2">
                    <input name="nome_prof" type="hidden" id="nome_prof" value="<% =nome_prof%>">
                    <input name="co_usr_prof" type="hidden" id="co_usr_prof" value="<% =co_usr_prof%>">
                    </font></td>
                  <td width="6%" height="30" class="tb_subtit"> <div align="right" >NOME: 
                    </div></td>
                  <td width="74%" height="30"> <font class="form_dado_texto"> 
                    <%response.Write(nome_prof)%>
                    </font> </td>
                </tr>
              </table></td>
          </tr>
          <tr class="tb_tit"
> 
            <td height="15" class="tb_tit"
>Grade de Aulas</td>
          </tr>
          <tr> 
            <td><table width="1000" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="13"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                  <td width="100" class="tb_subtit"> <div align="center"> UNIDADE 
                    </div></td>
                  <td width="100" class="tb_subtit"> <div align="center"> CURSO 
                    </div></td>
                  <td width="100" class="tb_subtit"> <div align="center"> ETAPA 
                    </div></td>
                  <td width="141" class="tb_subtit"> <div align="center">TURMA 
                    </div></td>
                  <td width="202" class="tb_subtit"> <div align="center">DISCIPLINA</div></td>
                  <td width="141" class="tb_subtit"> <div align="center"> MODELO</div></td>
                  <td width="203" class="tb_subtit"> <div align="center">COORDENADOR 
                    </div><input name="grade" type="hidden" value="<%=grade%>"> </td>
                </tr>
                <%

vertorExclui = split(grade,", ")
for i =0 to ubound(vertorExclui)

exclui = split(vertorExclui(i),"-")

unidade = exclui(0)
curso= exclui(1)
co_etapa= exclui(2)
turma= exclui(3)
mat_prin= exclui(4)
mat_fil= exclui(5)
tabela= exclui(6)
coordenador= exclui(7)


		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Abreviado_Curso")
%>
                <tr> 
                  <td width="13"> 
                  </td>
                  <td width="100"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_unidade)%>
                      <input name="unidade" type="hidden" id="unidade" value="<% = unidade %>">
                      </font></div></td>
                  <td width="100"> <div align="center"> <font class="form_dado_texto"> 
                      <%
response.Write(no_curso)%>
                      <input type="hidden" name="curso" value="<% = curso %>">
                      </font></div></td>
                  <td width="100"> <div align="center"> <font class="form_dado_texto"> 
                      <%

		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' and CO_Curso ='"& curso &"'"  
		RS3.Open SQL3, CON0
		
if RS3.EOF THEN
no_etapa="sem etapa"
else
no_etapa=RS3("NO_Etapa")
end if
response.Write(no_etapa)%>
                      <input name="etapa" type="hidden" id="etapa" value="<% = co_etapa %>">
                      </font></div></td>
                  <td width="141"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(turma)%>
                      <input name="turma" type="hidden" id="turma" value="<% = turma%>">
                      </font></div></td>
                  <td width="202"> <div align="center"> <font class="form_dado_texto"> 
                      <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Materia where CO_Materia ='"& mat_prin &"'" 
		RS4.Open SQL4, CON0
		
if RS4.EOF THEN
no_mat_prin="sem disciplina"
else
no_mat_prin=RS4("NO_Materia")
end if
response.Write(no_mat_prin)%>
                      <input name="mat_prin" type="hidden" id="mat_prin" value="<% = mat_prin%>">
                      </font> </div></td>
                  <td width="141"> <div align="center"> <font class="form_dado_texto"> 
                      <%
select case tabela
case "TB_NOTA_A" 
response.Write("Modelo A")
case "TB_NOTA_B" 
response.Write("Modelo B")
case "TB_NOTA_C"
response.Write("Modelo C")
end select

%>
                      <input name="tabela" type="hidden" id="tabela" value="<% = tabela%>">
                      </font> </div></td>
                  <td width="203"> <div align="center"> <font class="form_dado_texto"> 
                      <%
		Set RS8 = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT * FROM TB_Usuario where CO_Usuario ="& coordenador 
				RS8.Open SQL8, CON
		
no_coordenador = RS8("NO_Usuario")							  
					  response.Write(no_coordenador)%>
                      <input name="coordenador" type="hidden" id="coordenador" value="<% = coordenador%>">
                      </font> </div></td>
                </tr>
                <%
next
%>
              </table></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td bgcolor="#FFFFFF"> <table width="500" border="0" align="center" cellspacing="0">
                <tr> 
                  <td width="25%"> <div align="center"> 
                      <input name="alterar" type="submit" class="botao_cancelar" id="alterar" onClick="MM_goToURL('parent','altera.asp?cod_cons=<%=cod_prof%>&nvg=<%=nvg%>');return document.MM_returnValue" value="Cancelar">
                    </div></td>
                  <td width="25%"><div align="center"> 
                      <input name="Submit" type="submit" class="botao_prosseguir" value="Confirmar">
                    </div></td>
                </tr>
              </table></td>
          </tr>
        </table>
              </form></td>
  </tr>

<%end if
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