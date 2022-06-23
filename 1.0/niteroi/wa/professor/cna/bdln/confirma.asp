<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%
nivel=4
opt=request.QueryString("opt")

autoriza=Session("autoriza")
Session("autoriza")=autoriza
if autoriza="con" or  autoriza="in" or autoriza="ex" then
check_autoriza="con"
elseif autoriza="no" then
response.redirect("../../../../novologin.asp?opt=04")
end if

permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("chave")
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

if autoriza="con" or  autoriza="in" or autoriza="ex" then
check_autoriza="con"
elseif autoriza="no" then
response.redirect("../../../../novologin.asp?opt=04")
end if

nota=request.querystring("nt")

curso = request.querystring("c")
unidade = request.querystring("u")
grade = request.querystring("g")
turma = request.querystring("t")
co_etapa = request.querystring("e")
co_professor= request.querystring("pr")
co_mat_prin= request.querystring("d")
periodo= request.querystring("P")


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2

 call navegacao (CON,chave,nivel)
navega=Session("caminho")

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Curso")

		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' and CO_Curso ='"& curso &"'"  
		RS3.Open SQL3, CON0

no_etapa=RS3("NO_Etapa")

		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Materia where CO_Materia_Principal='"& co_mat_prin &"'"
		RS7.Open SQL7, CON0

		Set RS8 = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT * FROM TB_Materia where CO_Materia='"& co_mat_prin &"'"
		RS8.Open SQL8, CON0

		no_mat= RS8("NO_Materia")

		if RS7.EOF Then						
		co_mat_fil = co_mat_prin
		no_mat_prin = no_mat
		else		
		co_mat_fil= RS7("CO_Materia")
		end if
		
no_materia=no_mat

	%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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
MM_reloadPage(true);function checksubmit()
{
  if (document.inclusao.etapa.value == "")
  {    alert("Por favor, selecione uma etapa!")
    document.inclusao.etapa.focus()
    return false
  }
  if (document.inclusao.turma.value == "")
  {    alert("Por favor, selecione uma turma!")
    document.inclusao.turma.focus()
return false
}
  if (document.inclusao.mat_prin.value == "0")
  {    alert("Por favor, selecione uma disciplina!")
    document.inclusao.mat_prin.focus()
    return false
  }   
  if (document.inclusao.tabela.value == "")
  {    alert("Por favor, selecione uma tabela!")
    document.inclusao.tabela.focus()
    return false
  }                 	     
  return true
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
<body background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%if opt="blq" then%>
<form name="inclusao" method="post" action="bd.asp?opt=blq" onSubmit="return checksubmit()">	
<%elseif opt="dblq" then%>
<form name="inclusao" method="post" action="bd.asp?opt=dblq" onSubmit="return checksubmit()">
<%end if%> 
<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" align="center" background="../../../../img/fundo_interno.gif"  cellspacing="0"  bgcolor="#FFFFFF">
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
      <%	call mensagens(4,601,0,0) %>
    </td>
                </tr>
				
   <tr> 
            <td valign="top">            
        <table width="100%" border="0" align="center" cellspacing="0" class="tb_corpo"
>
          <tr class="tb_tit"
> 
            <td height="15" colspan="2" class="tb_tit"
><strong>Grade de aulas</strong> <input name="nota" type="hidden" id="nota" value="<%=nota%>"> 
              <input name="opt" type="hidden" id="opt2" value="<%=opt%>"> <input name="ano" type="hidden" id="ano2" value="<%=ano_letivo%>"> 
              &nbsp; &nbsp; </td>
          </tr>
          <tr> 
            <td colspan="2"><table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="8"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                  <td width="298" class="tb_subtit"> <div align="center">UNIDADE 
                    </div></td>
                  <td width="298" class="tb_subtit"> <div align="center">CURSO 
                    </div></td>
                  <td width="298" class="tb_subtit"> <div align="center">ETAPA 
                    </div></td>
                  <td width="298" class="tb_subtit"> <div align="center">TURMA 
                    </div></td>
                </tr>
                <tr> 
                  <td width="8"> </td>
                  <td width="298"> <div align="center"> <font class="form_dado_texto">   
                      <%response.Write(no_unidade)%>
                      <input name="unidade" type="hidden" id="unidade" value="<% = unidade %>">
                      </font></div></td>
                  <td width="298"> <div align="center"> <font class="form_dado_texto">   
                      <%
response.Write(no_curso)%>
                      <input type="hidden" name="curso" value="<% = curso %>">
                      </font></div></td>
                  <td width="298"> <div align="center"> <font class="form_dado_texto">   
                      <%


response.Write(no_etapa)%>
                      <input name="etapa" type="hidden" id="etapa" value="<% = co_etapa %>">
                      </font></div></td>
                  <td width="298"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(turma)%>
                      </font><font class="form_dado_texto"> 
                      <input name="turma" type="hidden" id="turma" value="<% = turma%>">
                      </font><font class="form_dado_texto"> </font><font class="form_dado_texto"> 
                      </font></div></td>
                </tr>
              </table></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2" bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="2"><table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="8"> </td>
                  <td width="450" class="tb_subtit"> <div align="center">DISCIPLINA</div></td>
                  <td width="450" class="tb_subtit"> <div align="center">PROFESSOR</div></td>
                  <td width="92" class="tb_subtit"> <div align="center">PER&Iacute;ODO</div></td>
                </tr>
                <tr> 
                  <td width="8"> </td>
                  <td width="450"> <div align="center"><font class="form_dado_texto"> 
                      <%



		response.Write(no_mat_prin)
%>
                      <input name="mat" type="hidden" id="mat" value="<%=co_mat_prin%>">
                      </font></div></td>
                  <td width="450" bgcolor="<%=cor%>"> <div align="center"> 
                      <%

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Professor where CO_Professor="& co_professor
		RS1.Open SQL1, CON1
			
if RS1.EOF then		
				

response.Write("nome em branco<br>")
else
no_prof= RS1("NO_Professor")
ativo = RS1("IN_Ativo_Escola")
if ativo = "True" then
Response.Write("<font size=1 face=Verdana, Arial, Helvetica, sans-serif><a class=ativos href=../professores/altera.asp?or=02&cod="&co_professor&" target=_parent>"&no_prof&"</a></font><br>")
else
Response.Write("<font size=1 face=Verdana, Arial, Helvetica, sans-serif><a class=inativos href=../professores/altera.asp?or=02&cod="&co_professor&" target=_parent>"&no_prof&"</a></font><br>")
end if
end if

%>
                      <input name="prof" type="hidden" id="prof" value="<%=co_professor%>">
                    </div></td>
                  <td width="92"> <div align="center"><font class="form_dado_texto"> 
                      <input name="periodo" type="hidden" id="periodo" value="<%=periodo%>">
                      <%
		Set RSp = Server.CreateObject("ADODB.Recordset")
		SQLp = "SELECT * FROM TB_Periodo where NU_Periodo="& periodo 
		RSp.Open SQLp, CON0
		
		no_periodo=RSp("NO_Periodo")
		
				  response.Write(no_periodo)
				 ' END IF
%>
                      </font></div></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td width="653"> <div align="center"> 
                <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','altera.asp?opt=vt&unidade=<% = unidade %>&curso=<% = curso %>&etapa=<% = co_etapa %>&turma=<% = turma%>');return document.MM_returnValue" value="Cancelar">
              </div></td>
            <td width="653"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
                <input type="submit" name="Submit" value="Confirmar" class="botao_prosseguir">
                </font></div></td>
          </tr>
        </table>        
     </td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
</form>
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