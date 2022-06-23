<%On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->



<!--#include file="../../../../inc/funcoes2.asp"-->


<%

nivel=4


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

chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

curso = request.Form("curso")
unidade = request.Form("unidade")
co_etapa = request.Form("etapa")
turma = request.Form("turma")
periodo = request.Form("periodo")
avaliacao = request.Form("campo")

if co_etapa="999999" then 
response.Redirect("consulta_turma_cp2.asp?opt=err2&c="&curso&"&u="&unidade&"")
else

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
		
		Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4 = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Abreviado_Curso")

		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Da_Aula where NU_Unidade="& unidade &" AND CO_Curso='"& curso &"' AND CO_Etapa='"& co_etapa &"'"
		RS2.Open SQL2, CON2

tp_nota= RS2("TP_Nota")

if tp_nota="TB_NOTA_A" then
CAMINHO_n = CAMINHO_na
elseif tp_nota="TB_NOTA_B" then
CAMINHO_n = CAMINHO_nb
elseif tp_nota="TB_NOTA_C" then
CAMINHO_n = CAMINHO_nc
end if
		Set CON3 = Server.CreateObject("ADODB.Connection") 
		ABRIR3 = "DBQ="& CAMINHO_n & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3
	%>
<html>
<head>
<title>Web Acad&ecirc;mico</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<% EscreveFuncaoJavaScriptCurso ( CON0 ) %>
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
}  function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}

//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900" background="../../../../img/fundo.gif" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<table width="1000" height="670" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
                  <tr>                    
            <td height="10" class="tb_caminho"> <font color="#FFFF33" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%
	  response.Write(navega)

%>
              </font>
	</td>
  </tr>
  <%if check_autoriza="con" then%>
 <tr>                   
    <td height="10"> 
      <% call mensagens(4,9701,1,0)%>
    </td>
    </tr> 
<%else%>					
    <tr>                   
    <td height="10"> 
      <%	call mensagens(4,603,0,0) %>
    </td>
                </tr>
<%end if%>												  				  


          <tr> 
            <td valign="top">
<form name="inclusao" method="post" action="consulta_turma_cp3.asp?or=02" onSubmit="return checksubmit()">
                <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">
                  <tr class="tb_tit"> 
                    
            <td height="15" class="tb_tit">Grade de Aulas </td>
                  </tr>
                  <tr> 
                    
          <td><table width="100%" border="0" cellspacing="0">
              <tr> 
                <td width="8"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                <td class="tb_subtit"> <div align="center">UNIDADE </div></td>
                <td class="tb_subtit"> <div align="center">CURSO </div></td>
                <td class="tb_subtit"> <div align="center">ETAPA </div></td>
                <td class="tb_subtit"><div align="center">PER&Iacute;ODO</div></td>
                <td class="tb_subtit"> <div align="center">AVALIA&Ccedil;&Atilde;O</div></td>
              </tr>
              <tr> 
                <td width="8"> </td>
                <td> <div align="center"> <font class="form_dado_texto"> 
                    <%response.Write(no_unidade)%>
                    <input name="unidade" type="hidden" id="unidade" value="<% = unidade %>">
                    </font></div></td>
                <td> <div align="center"> <font class="form_dado_texto"> 
                    <%
response.Write(no_curso)%>
                    <input type="hidden" name="curso" value="<% = curso %>">
                    </font></div></td>
                <td><div align="center"> <font class="form_dado_texto"> 
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
                <td><div align="center"><font class="form_dado_texto"> 
                    <input name="periodo" type="hidden" id="periodo" value="<%=periodo%>">
                        <option value="0" selected></option>
                        <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo where NU_Periodo="&periodo
		RS4.Open SQL4, CON0

NO_Periodo=RS4("NO_Periodo")
                        response.Write(NO_Periodo)%>
                    </font></div></td>
                <td> <div align="center"> <font class="form_dado_texto"> </font></font><font class="form_dado_texto"> 
                      <select name="campo" class="borda" id="campo" onChange="MM_callJS('submitfuncao()')">
                        <option value="999999" ></option>
                      <%
Select case periodo

case 1
if avaliacao="Apr1_P1" then
%>
                      <option value="Apr1_P1" selected>Apr1</option>
                      <%else%>
                      <option value="Apr1_P1">Apr1</option>
                      <%end if
if avaliacao="Apr2_P1" then
%>
                      <option value="Apr2_P1" selected>Apr2</option>
                      <%else%>
                      <option value="Apr2_P1">Apr2</option>
                      <%end if
if avaliacao="Apr3_P1" then
%>
                      <option value="Apr3_P1" selected>Apr3</option>
                      <%else%>
                      <option value="Apr3_P1" >Apr3</option>
                      <%end if
if avaliacao="Apr4_P1" then
%>
                      <option value="Apr4_P1" selected>Apr4</option>
                      <%else%>
                      <option value="Apr4_P1" >Apr4</option>
                      <%end if
if avaliacao="Apr5_P1" then
%>
                      <option value="Apr5_P1" selected>Apr5</option>
                      <%else%>
                      <option value="Apr5_P1" >Apr5</option>
                      <%end if
if avaliacao="Apr6_P1" then
%>
                      <option value="Apr6_P1" selected>Apr6</option>
                      <%else%>
                      <option value="Apr6_P1" >Apr6</option>
                      <%end if
if avaliacao="Apr7_P1" then
%>
                      <option value="Apr7_P1" selected>Tec1</option>
                      <%else%>
                      <option value="Apr7_P1" >Tec1</option>
                      <%end if
if avaliacao="Apr8_P1" then
%>
                      <option value="Apr8_P1" selected>Tec2</option>
                      <%else%>
                      <option value="Apr8_P1" >Tec2</option>
                      <%end if					  
if avaliacao="VA_Pr1" then
%>
                      <option value="VA_Pr1" selected>Pr1</option>
                      <%else%>
                      <option value="VA_Pr1" >Pr1</option>
                      <%end if
if avaliacao="VA_Te1" then
%>
                      <option value="VA_Te1" selected>Pr2</option>
                      <%else%>
                      <option value="VA_Te1" >Pr2</option>
                      <%end if
if avaliacao="VA_Bon1" then
%>
                      <option value="VA_Bon1" selected>Bon</option>
                      <%else%>
                      <option value="VA_Bon1" >Bon</option>
                      <%end if
'/////////////////////////////////////////////////////////////////////////////////////////////		  				 
case 2
if avaliacao="Apr1_P2" then
%>
                      <option value="Apr1_P2" selected>Apr1</option>
                      <%else%>
                      <option value="Apr1_P2">Apr1</option>
                      <%end if
if avaliacao="Apr2_P2" then
%>
                      <option value="Apr2_P2" selected>Apr2</option>
                      <%else%>
                      <option value="Apr2_P2">Apr2</option>
                      <%end if
if avaliacao="Apr3_P2" then
%>
                      <option value="Apr3_P2" selected>Apr3</option>
                      <%else%>
                      <option value="Apr3_P2" >Apr3</option>
                      <%end if
if avaliacao="Apr4_P2" then
%>
                      <option value="Apr4_P2" selected>Apr4</option>
                      <%else%>
                      <option value="Apr4_P2" >Apr4</option>
                      <%end if
if avaliacao="Apr5_P2" then
%>
                      <option value="Apr5_P2" selected>Apr5</option>
                      <%else%>
                      <option value="Apr5_P2" >Apr5</option>
                      <%end if
if avaliacao="Apr6_P2" then
%>
                      <option value="Apr6_P2" selected>Apr6</option>
                      <%else%>
                      <option value="Apr6_P2" >Apr6</option>
                      <%end if
if avaliacao="Apr7_P2" then
%>
                      <option value="Apr7_P2" selected>Tec1</option>
                      <%else%>
                      <option value="Apr7_P2" >Tec1</option>
                      <%end if
if avaliacao="Apr8_P2" then
%>
                      <option value="Apr8_P2" selected>Tec2</option>
                      <%else%>
                      <option value="Apr9_P2" >Tec2</option>
                      <%end if					  
if avaliacao="VA_Pr2" then
%>
                      <option value="VA_Pr2" selected>Pr1</option>
                      <%else%>
                      <option value="VA_Pr2" >Pr1</option>
                      <%end if
if avaliacao="VA_Te2" then
%>
                      <option value="VA_Te2" selected>Pr2</option>
                      <%else%>
                      <option value="VA_Te2" >Pr2</option>
                      <%end if					  
if avaliacao="VA_Bon2" then
%>
                      <option value="VA_Bon2" selected>Bon</option>
                      <%else%>
                      <option value="VA_Bon2" >Bon</option>
                      <%end if

'/////////////////////////////////////////////////////////////////////////////////////////////		  				 
case 3
if avaliacao="Apr1_P3" then
%>
                      <option value="Apr1_P3" selected>Apr1</option>
                      <%else%>
                      <option value="Apr1_P3">Apr1</option>
                      <%end if
if avaliacao="Apr2_P3" then
%>
                      <option value="Apr2_P3" selected>Apr2</option>
                      <%else%>
                      <option value="Apr2_P3">Apr2</option>
                      <%end if
if avaliacao="Apr3_P3" then
%>
                      <option value="Apr3_P3" selected>Apr3</option>
                      <%else%>
                      <option value="Apr3_P3" >Apr3</option>
                      <%end if
if avaliacao="Apr4_P3" then
%>
                      <option value="Apr4_P3" selected>Apr4</option>
                      <%else%>
                      <option value="Apr4_P3" >Apr4</option>
                      <%end if
if avaliacao="Apr5_P3" then
%>
                      <option value="Apr5_P3" selected>Apr5</option>
                      <%else%>
                      <option value="Apr5_P3" >Apr5</option>
                      <%end if
if avaliacao="Apr6_P3" then
%>
                      <option value="Apr6_P3" selected>Apr6</option>
                      <%else%>
                      <option value="Apr6_P3" >Apr6</option>
                      <%end if
if avaliacao="Apr7_P3" then
%>
                      <option value="Apr7_P3" selected>Tec1</option>
                      <%else%>
                      <option value="Apr7_P3" >Tec1</option>
                      <%end if
if avaliacao="Apr8_P3" then
%>
                      <option value="Apr8_P3" selected>Tec2</option>
                      <%else%>
                      <option value="Apr8_P3" >Tec2</option>
                      <%end if					  
if avaliacao="VA_Pr3" then
%>
                      <option value="VA_Pr3" selected>Pr1</option>
                      <%else%>
                      <option value="VA_Pr3" >Pr1</option>
                      <%end if
if avaliacao="VA_Te3" then
%>
                      <option value="VA_Te3" selected>Pr2</option>
                      <%else%>
                      <option value="VA_Te3" >Pr2</option>
                      <%end if
if avaliacao="VA_Bon3" then
%>
                      <option value="VA_Bon3" selected>Bon</option>
                      <%else%>
                      <option value="VA_Bon3" >Bon</option>
                      <%end if
'/////////////////////////////////////////////////////////////////////////////////////////////		  				 
case 4
if avaliacao="Apr1_EC" then
%>
                      <option value="Apr1_EC" selected>Apr1</option>
                      <%else%>
                      <option value="Apr1_EC">Apr1</option>
                      <%end if
if avaliacao="Apr2_EC" then
%>
                      <option value="Apr2_EC" selected>Apr2</option>
                      <%else%>
                      <option value="Apr2_EC">Apr2</option>
                      <%end if
if avaliacao="Apr3_P3" then
%>
                      <option value="Apr3_P3" selected>Apr3</option>
                      <%else%>
                      <option value="Apr3_P3" >Apr3</option>
                      <%end if
if avaliacao="Apr4_EC" then
%>
                      <option value="Apr4_EC" selected>Apr4</option>
                      <%else%>
                      <option value="Apr4_EC" >Apr4</option>
                      <%end if
if avaliacao="Apr5_EC" then
%>
                      <option value="Apr5_EC" selected>Apr5</option>
                      <%else%>
                      <option value="Apr5_EC" >Apr5</option>
                      <%end if
if avaliacao="Apr6_EC" then
%>
                      <option value="Apr6_EC" selected>Apr6</option>
                      <%else%>
                      <option value="Apr6_EC" >Apr6</option>
                      <%end if
if avaliacao="Apr7_EC" then
%>
                      <option value="Apr7_EC" selected>Tec1</option>
                      <%else%>
                      <option value="Apr7_EC" >Tec1</option>
                      <%end if
if avaliacao="Apr8_EC" then
%>
                      <option value="Apr8_EC" selected>Tec2</option>
                      <%else%>
                      <option value="Apr8_EC" >Tec2</option>
                      <%end if					  					  
if avaliacao="VA_Pr4" then
%>
                      <option value="VA_Pr4" selected>Pr1</option>
                      <%else%>
                      <option value="VA_Pr4" >Pr1</option>
                      <%end if
end select
%>
                    </select>
                    </font><font class="form_dado_texto"> </font></div></td>
              </tr>
            </table></td>
                  </tr>
                  <tr bgcolor="#FFFFFF"> 
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>
                  <tr> 
                    
          <td><table width="100%" border="0" cellspacing="0">
                <tr> 
                <td width="4"> </td>
                  <td width="241" class="tb_subtit"> 
                    <div align="center">DISCIPLINA</div></td>
                  <td width="76" class="tb_subtit">
<div align="center">TURMA</div></td>
                  <td width="37" class="tb_subtit">
<div align="center">N&ordm;</div></td>
                  <td width="360" class="tb_subtit"> 
                    <div align="center">ALUNO</div></td>
              </tr>
              <%

  		Set RS11 = Server.CreateObject("ADODB.Recordset")
		SQL11 = "SELECT * FROM TB_Aluno_Esta_Turma where NU_Unidade="& unidade &" AND CO_Curso='"& curso &"' AND CO_Etapa='"& co_etapa &"' order by CO_Turma,NU_Chamada"
		RS11.Open SQL11, CON4
check=2		
while not RS11.EOF
 if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if						
CO_Matricula = RS11("CO_Matricula")
turma= RS11("CO_Turma")
NU_Chamada = RS11("NU_Chamada")
  
    	Set RS12 = Server.CreateObject("ADODB.Recordset")
		SQL12 = "SELECT * FROM TB_Aluno WHERE CO_Matricula = "&CO_Matricula &""
		RS12.Open SQL12, CON4
 NO_Aluno = RS12("NO_Aluno") 
  				
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
		RS5.Open SQL5, CON0

while not RS5.EOF
co_mat_prin= RS5("CO_Materia")
mae= RS5("IN_MAE")
fil= RS5("IN_FIL")
nu_peso= RS5("NU_Peso")
in_co= RS5("IN_CO")

if mae=TRUE AND fil=true AND in_co=false AND isnull(nu_peso) then
RS5.MOVENEXT
elseif mae=TRUE AND fil=false AND in_co=true AND isnull(nu_peso) then
RS5.MOVENEXT
else


  
  
		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Materia where CO_Materia_Principal='"& co_mat_prin &"'"
		RS7.Open SQL7, CON0

		Set RS8 = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT * FROM TB_Materia where CO_Materia='"& co_mat_prin &"'"
		RS8.Open SQL8, CON0

		no_mat= RS8("NO_Materia")

		if RS7.EOF Then						
		co_mat_fil = co_mat_prin
		else		
		co_mat_fil= RS7("CO_Materia")
		end if

'response.Write("SQL10 = SELECT "&avaliacao&" as campo_check FROM "& tp_nota &" where CO_Matricula = "&CO_Matricula &" And CO_Materia='"& co_mat_fil&"'")
  
  		Set RS10 = Server.CreateObject("ADODB.Recordset")
		SQL10 = "SELECT "&avaliacao&" as campo_check FROM "& tp_nota &" where CO_Matricula = "&CO_Matricula &" And CO_Materia='"& co_mat_fil&"'"
		RS10.Open SQL10, CON3	
' And "& campo_check &" is null 
if RS10.EOF THEN
%>
              <tr> 
                <td width="4"> </td>
                  <td width="241" class="<%=cor%>"> 
                    <div align="center"><font class="form_dado_texto">  
                    <%

		response.Write(no_mat)
%>
                    </font></div></td>
                  <td width="76" class="<%=cor%>"> 
                    <div align="center"><font class="form_dado_texto">  
                    <%response.Write(turma)%>
                    </font></font></div></td>
                  <td width="37" class="<%=cor%>"> 
                    <div align="center"> <font class="form_dado_texto">  
                    <%response.Write(NU_Chamada)%>
                    </font> </font></div></td>
                  <td width="360" class="<%=cor%>"> 
                    <div align="center"><font class="form_dado_texto">  
                    <%response.Write(NO_Aluno) %>
                    </font></font></div></td>
              </tr>
              <%				
else
while not RS10.EOF 
IF RS10("campo_check")="" or isnull(RS10("campo_check")) or RS10("campo_check")=0 then
%>
              <tr> 
                <td width="4"> </td>
                  <td width="241" class="<%=cor%>"> 
                    <div align="center"><font class="form_dado_texto">  
                    <%

		response.Write(no_mat)
%>
                    </font></div></td>
                  <td width="76" class="<%=cor%>"> 
                    <div align="center"><font class="form_dado_texto">  
                    <%response.Write(turma)%>
                    </font></font></div></td>
                  <td width="37" class="<%=cor%>"> 
                    <div align="center"> <font class="form_dado_texto">  
                    <%response.Write(NU_Chamada)%>
                    </font> </font></div></td>
                  <td width="360" class="<%=cor%>"> 
                    <div align="center"><font class="form_dado_texto">  
                    <%response.Write(NO_Aluno) %>
                    </font></font></div></td>
              </tr>
              <%
RS10.MOVENEXT

else

RS10.MOVENEXT

end if

WEND
end if
end if
RS5.MOVENEXT
WEND
check=check+1
RS11.MOVENEXT
WEND
call GravaLog (chave,avaliacao)
%>
            </table></td>
          </tr>
        </table>
</form>
</td>
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
<%end if%>
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