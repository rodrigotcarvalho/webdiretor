<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes3.asp"-->


<%opt = REQUEST.QueryString("opt")
obr = request.QueryString("o")
nivel=4
Server.ScriptTimeout = 600 'valor em segundos
opt=request.QueryString("opt")

autoriza=Session("autoriza")
Session("autoriza")=autoriza

permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

if opt="direto" then

curso = request.querystring("curso")
unidade = request.querystring("unidade")
grade = request.querystring("grade")
turma = "sem turma"

elseif opt= "vt" then
dados= split(obr, "_" )
unidade= dados(0)
curso= dados(1)
co_etapa= dados(2)
turma= dados(3)


else

curso = request.Form("curso")
unidade = request.Form("unidade")
co_etapa = request.Form("etapa")
turma = request.Form("turma")
end if
ano_letivo = session("ano_letivo")
obr=unidade&"_"&curso&"_"&co_etapa&"_"&turma
inf=unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma
if co_etapa = "f0"then
co_etapa=0
elseif co_etapa = "f1" or co_etapa = "m1"then
co_etapa=1
elseif co_etapa = "f2" or co_etapa = "m2"then
co_etapa = 2
elseif co_etapa = "f3" or co_etapa = "m3"then
co_etapa = 3
elseif co_etapa = "f4" then
co_etapa = 4
elseif co_etapa = "f5" then
co_etapa = 5
elseif co_etapa = "f6" then
co_etapa = 6
elseif co_etapa = "f7" then
co_etapa = 7
elseif co_etapa = "f8" then
co_etapa = 8
elseif co_etapa = "f55" then
co_etapa = 55
elseif co_etapa = "f66" then
co_etapa = 66
elseif co_etapa = "f77" then
co_etapa = 77
end if

if co_etapa="999999" then 
response.Redirect("tabelas2.asp?opt=err2&or=02&curso="&curso&"&unidade="&unidade&"")
else

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Abreviado_Curso")

		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' and CO_Curso ='"& curso &"'"  
		RS3.Open SQL3, CON0

 call navegacao (CON,chave,nivel)
navega=Session("caminho")
		
if RS3.EOF THEN
no_etapa="sem etapa"
else
no_etapa=RS3("NO_Etapa")
end if
	%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">

<script type="text/javascript" src="../../../../js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
function MM_showHideLayers() { //v6.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
    obj.visibility=v; }
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
//function submitano()  
//{
//   var f=document.forms[0]; 
//      f.submit(); 
//}
//function submitsistema()  
//{
//   var f=document.forms[1]; 
//      f.submit(); 
//}
//function submitrapido()  
//{
//   var f=document.forms[2]; 
//      f.submit(); 
//}  
//function submitfuncao()  
{
   var f=document.forms[0]; 
      f.submit(); 
} 
function checksubmit()
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
<%		Set RS5count = Server.CreateObject("ADODB.Recordset")
		SQL5count = "SELECT * FROM TB_Matriculas where NU_Ano="& ano_letivo &" AND NU_Unidade="& unidade &" AND CO_Curso='"& curso &"' AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"' ORDER BY NU_Chamada"
		RS5count.Open SQL5count, CON1
z=2
while not RS5count.EOF
matricula= RS5count("CO_Matricula")
%>
function centraliza_<%response.Write(matricula)%>(w,h){
//o 120 e o 16 se referem ao tamanho di cabeçalho do navegador e a barra de rolagem respectivamente
    x = parseInt((screen.width - w - 16)/2);
    y = parseInt((screen.height - h - 120)/2);
   //alert(x + '\n' + y);
    document.getElementById('c<%=matricula%>t').style.left = x;
    document.getElementById('c<%=matricula%>t').style.top = y;
	
//	alert('w '+x +' h '+ y)
}
<%
z=z+1
RS5count.MOVENEXT
WEND
%>
//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif"leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<div id="fundo" style="position:absolute; left:0px; top:0px; width:100%; height:1200; z-index:1; background-color: #000000; layer-background-color: #000000; border: 1px none #000000; visibility: hidden;" class="transparente"></div>
<%
		Set RS5count = Server.CreateObject("ADODB.Recordset")
		SQL5count = "SELECT * FROM TB_Matriculas where NU_Ano="& ano_letivo &" AND NU_Unidade="& unidade &" AND CO_Curso='"& curso &"' AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"' ORDER BY NU_Chamada"
		RS5count.Open SQL5count, CON1
z=2
while not RS5count.EOF
matricula= RS5count("CO_Matricula")

		Set RS8count = Server.CreateObject("ADODB.Recordset")
		SQL8count = "SELECT * FROM TB_Alunos where CO_Matricula="& matricula 
		RS8count.Open SQL8count, CON1

		no_al= RS8count("NO_Aluno")
%>
<div id="c<%=matricula%>t" style="position:absolute; width:500px; visibility: hidden; z-index: <%=z%>; height: 536px;"> 
  <table width="100%" border="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr> 
      <td width="478"> <div align="right"> <span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','c<%=matricula%>t','','hide')">fechar</a></font></span></div></td>
      <td width="20"><div align="right"><span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','c<%=matricula%>t','','hide')"><img src="../../../../img/fecha.gif" width="20" border="0"></a></font></span></div></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center" ><img src="../../../../img/fotos/aluno/<%response.Write(matricula)%>.jpg" height="500"></div></td>
    </tr>
    <tr>
      <td colspan="2"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
          <%response.Write(no_al)%>
          </font></div></td>
    </tr>
  </table>
</div>

<%

z=z+1
RS5count.MOVENEXT
WEND
%>

<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr>                    
            <td height="10" class="tb_caminho"> <font class="style-caminho">
              <%
	  response.Write(navega)

%>
              </font>
	</td>
  </tr>             <tr> 
                  
    <td height="10"> 
      <%
	call mensagens(nivel,305,0,0) 
%>
    </td>
                </tr>
    <td valign="top"> 
      <form name="inclusao" method="post" action="mapa.asp" onSubmit="return checksubmit()">
                
        <table width="1000" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo"
>
          <tr class="tb_tit"
> 
            <td width="653" height="15" class="tb_tit"
>Segmento </td>
          </tr>
          <tr> 
            <td><table width="1000" border="0" cellspacing="0">
                <tr> 
                  <td width="250" class="tb_subtit"> 
                    <div align="center">UNIDADE </div></td>
                  <td width="250" class="tb_subtit"> 
                    <div align="center">CURSO </div></td>
                  <td width="250" class="tb_subtit"> 
                    <div align="center">ETAPA </div></td>
                  <td width="250" class="tb_subtit"> 
                    <div align="center">TURMA </div></td>
                </tr>
                <tr> 
                  <td width="250"> 
                    <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_unidade)%>
                      </font></div></td>
                  <td width="250"> 
                    <div align="center"> <font class="form_dado_texto"> 
                      <%
response.Write(no_curso)%>
                      </font></div></td>
                  <td width="250">
<div align="center"> <font class="form_dado_texto"> 
                      <%


response.Write(no_etapa)%>
                      <input name="etapa" type="hidden" id="etapa" value="<% = co_etapa %>">
                      </font></div></td>
                  <td width="250"> 
                    <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(turma)%>
                      <input name="turma" type="hidden" id="turma" value="<% = turma%>">
                      </font></div></td>
                </tr>
              </table></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
          <tr> 
            <td> <table width="1000" border="0" cellspacing="0">
                <tr> 
                  <%
				
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Matriculas where NU_Ano="& ano_letivo &" AND NU_Unidade="& unidade &" AND CO_Curso='"& curso &"' AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"' ORDER BY NU_Chamada"
		RS5.Open SQL5, CON1
contafotos=0
check=2
while not RS5.EOF
chamada= RS5("NU_Chamada")
matricula= RS5("CO_Matricula")
rematricula= RS5("DA_Rematricula")
situacao= RS5("CO_Situacao")
' if check mod 6 =0 then
'  cor = "tb_fundo_linha_par" 
' else cor ="tb_fundo_linha_impar"
'  end if
		Set RS8 = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT * FROM TB_Alunos where CO_Matricula="& matricula 
		RS8.Open SQL8, CON1

		no_al= RS8("NO_Aluno")
		

if contafotos mod 6 = 0 then
%>
                </tr>
                <tr > 
                  <td width="2"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                  <td > <div align="center"> 
                      <table width="145" height="110" border="0" cellspacing="0" bgcolor="#EEEEEE">
                        <tr> 
                          <td height="110" bgcolor="#EEEEEE"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="centraliza_<%response.Write(matricula)%>(500,536);MM_showHideLayers('fundo','','show','c<%=matricula%>t','','show')"><img src="../../../../img/fotos/aluno/<% =matricula %>.jpg" alt="" height="110" border="0"></a></font></div></td>
                        </tr>
                        <tr> 
                          <td bgcolor="#EEEEEE"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                              <a href="../alunos/altera.asp?ori=01&vd=crmt&cod_cons=<% =matricula %>&o=<%=obr%>" class='linkdois'>
                              <%		response.Write(no_al)%>
                              </a> </font></div></td>
                        </tr>
                      </table>
                    </div></td>
                  <%
contafotos=contafotos+1
check=check+1
RS5.MOVENEXT
else%>
                  <td width="2"></td>
                  <td> <div align="center"> 
                      <table width="145" height="110" border="0" cellspacing="0" bgcolor="#EEEEEE">
                        <tr> 
                          <td height="110" bgcolor="#EEEEEE"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="centraliza_<%response.Write(matricula)%>(500,536);MM_showHideLayers('fundo','','show','c<%=matricula%>t','','show')"><img src="../../../../img/fotos/aluno/<% =matricula %>.jpg" alt="" height="110" border="0"></a></font></div></td>
                        </tr>
                        <tr> 
                          <td bgcolor="#EEEEEE"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                              <a href="../alunos/altera.asp?ori=01&vd=crmt&cod_cons=<% =matricula %>&o=<%=obr%>" class='linkdois'>
                              <%		response.Write(no_al)%>
                              </a> </font></div></td>
                        </tr>
                      </table>
                    </div></td>
                  <%

contafotos=contafotos+1
check=check+1
RS5.MOVENEXT
end if
WEND
%>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td colspan="7"><hr></td>
          </tr>
          <tr> 
            <td colspan="7"><div align="center">
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="34%"><div align="center">
                        <input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','altera.asp?ori=01&opt=vt&o=<%=obr%>');return document.MM_returnValue" value="Voltar">
                      </div></td>
                    <td width="34%">&nbsp;</td>
                    <td width="34%"> <div align="center">
                      <input name="bt2" type="button" class="botao_prosseguir" id="bt2" onClick="MM_goToURL('parent','../../../../relatorios/swd100.asp?inf=<%=inf%>');return document.MM_returnValue" value="Imprimir">
                    </div></td>
                  </tr>
                </table>
              </div></td>
          </tr>
        </table>
        </form>
</td>
  </tr>
  <tr>
    <td height="40"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

</body>
<%end if 
call GravaLog (chave,obr)
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
pasta=arPath(seleciona)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>