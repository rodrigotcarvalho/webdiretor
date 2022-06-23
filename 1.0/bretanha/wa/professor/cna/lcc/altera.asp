<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/bd_grade.asp"-->

<%
opt = request.QueryString("opt")
ori = request.QueryString("ori")
ano_letivo = session("ano_letivo")

co_usr = session("co_user")
nivel=4

chave=session("nvg")
session("nvg")=chave

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
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1

if opt="ok" then
	vetor_cod_cons = session("vetor_cod_cons")
	ori = session("ori")
	obr = session("obr")	
else
	if ori=01 or ori = "01" or ori=02 or ori = "02" then
		vetor_cod_cons = request.QueryString("cod_cons")
		obr = vetor_cod_cons		
			
	else
		obr = request.form("obr")
		vetor_cod_cons = request.form("alunos")			
	end if	
end if

'if ori=01 or ori = "01" or ori=02 or ori = "02" then	
'	if ori=01 or ori = "01" then
'		matric_exibe = vetor_cod_cons
'	end if
'	
'	if ori=02 or ori = "02" then 
'		Set RSNOME = Server.CreateObject("ADODB.Recordset")
'		SQLNOME = 	"Select NO_Aluno from TB_Alunos WHERE CO_Matricula ="&vetor_cod_cons
'		RSNOME.Open SQLNOME, CON1		
'		
'		nome_exibe = RS("NO_Aluno")	
'	end if
'
'else
'	dados_selecao = split(obr,"$!$")
'	unidade_exibe = dados_selecao(0)
'	curso_exibe = dados_selecao(1)
'	etapa_exibe = dados_selecao(2)
'	turma_exibe = dados_selecao(3)			
'end if	


session("vetor_cod_cons") = ""
session("obr") = ""	
session("ori") = ""	



 call navegacao (CON,chave,nivel)
navega=Session("caminho")	

 'Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Professor Where CO_Usuario = "&co_usr
		RS2.Open SQL2, CON2
		
'if RS2.EOF then
'Response.Write("Usuário não é Professor!")
'else		
''co_prof=RS2("CO_Professor")
'End if
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
function checksubmit()
{
  if (document.busca.busca1.value != "" && document.busca.busca2.value != "")
  {    alert("Por favor digite SOMENTE uma opção de busca!")
    document.busca.busca1.focus()
    return false
  }
    if (document.busca.busca1.value == "" && document.busca.busca2.value == "")
  {    alert("Por favor digite uma opção de busca!")
    document.busca.busca1.focus()
    return false
  }
  return true
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
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
document.all.divEtapa.innerHTML ="<select name='etapa' class='select_style' id='etapa'><option value='999990' selected>           </option></select>"
document.all.divTurma.innerHTML = "<select name='turma' class='select_style' id='turma'><option value='999990' selected>           </option></select>"
GuardaUnidade(uTipo)
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
document.all.divTurma.innerHTML = "<select name='turma' class='select_style' id='turma'><option value='999990' selected>           </option></select>"
GuardaCurso(cTipo)
                                                           }
                                               }
 
                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }
 
 
						 function recuperarTurma(eTipo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t5", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t	
GuardaEtapa(eTipo)																		   
                                                           }
                                               }
 
                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }

						 
	 function GuardaUnidade(u)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../../../../../global/guarda_valores_digitados.asp?opt=valor3", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }

	
		   oHTTPRequest.send("valor3=" + u);
		}
	 function GuardaCurso(c)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../../../../../global/guarda_valores_digitados.asp?opt=valor4", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor4=" + c);
		}
function GuardaEtapa(e)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../../../../../global/guarda_valores_digitados.asp?opt=valor5", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor5=" + e);
		}			
		

	 function gravarTurma(t)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../../../../../global/guarda_valores_digitados.asp?opt=valor6", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor6=" + t);
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
      <%
if opt = "ok" then%>
  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(4,720,2,0)
%>
    </td>
                  </tr>
<%end if%>                  
<tr>
                    
    <td height="10"> 
      <%
		call mensagens(4,9708,0,0)
%>
    </td>
                  </tr>
                <tr>

            <td valign="top"> 
                
        <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit">
            <td width="653" height="15" class="tb_tit">Aluno Selecionado</td>
          </tr>
          <tr>
            <td><form action="bd.asp" method="post"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="tb_subtit">
                <td width="70" align="center">Chamada</td>
                <td><div align="center">Matr&iacute;cula
                <input name="vetor_cod_cons" type="hidden" value="<%response.write(vetor_cod_cons)%>">
                
                </div></td>
                <td>&nbsp;Nome
                  <input name="obr" type="hidden" value="<%response.write(obr)%>"></td>
                <td height="10"><div align="center">Unidade</div></td>
                <td height="10"><div align="center">Curso</div></td>
                <td height="10"><div align="center"> Etapa</div></td>
                <td height="10"><div align="center">Turma </div></td>
                <td><div align="center">Disciplina</div></td>
                <td align="center"> Per&iacute;odo</td>

                </tr>
<%   
cod_cons = split(vetor_cod_cons,", ")
check=2
For vcc=0 to ubound(cod_cons)
	nu_matric = cod_cons(vcc)
		if check mod 2 =0 then
			classe = "tb_fundo_linha_par" 
		else 
			classe ="tb_fundo_linha_impar"
		end if 	
	
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "Select TB_Matriculas.NU_Chamada, TB_Alunos.NO_Aluno, TB_Matriculas.NU_Unidade,TB_Matriculas.CO_Curso,TB_Matriculas.CO_Etapa,TB_Matriculas.CO_Turma from TB_Matriculas, TB_Alunos WHERE TB_Matriculas.NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Situacao='C' AND TB_Matriculas.CO_Matricula="&nu_matric&"  AND TB_Matriculas.CO_Matricula=TB_Alunos.CO_Matricula"
	'SQL = "Select * from TB_Alunos WHERE CO_Matricula="&nu_matric	
	RS.Open SQL, CON1	
		
	no_aluno= RS("NO_Aluno")			
	nu_chamada = RS("NU_Chamada")
	unidade_aluno = RS("NU_Unidade")
	curso_aluno = RS("CO_Curso")
	co_etapa_aluno = RS("CO_Etapa")
	turma_aluno = RS("CO_Turma")	
	
call GeraNomes("PORT",unidade_aluno,curso_aluno,co_etapa_aluno,CON0)
no_unidade = session("no_unidades")
no_curso = session("no_grau")
no_etapa = session("no_serie")		
	
	'response.Write(nu_chamada&"=")	
	val_bonus = BonusMediaAnual(nu_matric, null)	
	val_bonus = val_bonus*1
	
	'response.End()
	if val_bonus=0 then
		selected_0 = "selected"
		selected_5 = ""		
	elseif val_bonus=5 then
		selected_0 = ""
		selected_5 = "selected"		
	end if
	%>
             
              <tr>
                <td width="70" align="center" class="<%response.Write(classe)%>"><%response.Write(nu_chamada)%></td>
                <td width="75" align="center" class="<%response.Write(classe)%>"><%response.Write(nu_matric)%></td>
                <td width="350" align="center" class="<%response.Write(classe)%>"><%response.Write(no_aluno)%></td>
                <td width="70" class="<%response.Write(classe)%>">
                  <%response.Write(no_unidade)%>
                </td>
                <td width="70" align="center" class="<%response.Write(classe)%>">
                  <%response.Write(no_curso)%>
                </div></td>
                <td width="70" align="center" class="<%response.Write(classe)%>">
                  <%response.Write(no_etapa)%>
                </td>
                <td width="70" align="center" class="<%response.Write(classe)%>">
                  <%response.Write(turma_aluno)%>
                </td>
                 <td width="167"><div align="center">
                    <div id="divDisciplina">
                     
<%		


		Set RSG = Server.CreateObject("ADODB.Recordset")
		SQLG = "SELECT DISTINCT CO_Materia FROM TB_Programa_Aula where IN_MAE = True And CO_Etapa = '"&co_etapa_aluno&"' AND CO_Curso = '"&curso_aluno&"' order by CO_Materia"
		RSG.Open SQLG, CON0
		
IF RSG.EOF THEN

RESPONSE.Write("Sem disciplinas cadastradas. Procure seu Coordenador.")


ELSE
%>
                      <select name="mat_prin" class="select_style">
                        <%
while not RSG.EOF
co_mat_prin= RSG("CO_Materia")

		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Materia where CO_Materia ='"& co_mat_prin &"'"
		RS7.Open SQL7, CON0
		
		no_mat_prin= RS7("NO_Materia")
		
	if co_mat_prin=codigo_mat_prin then
	%>
							<option value="<%=co_mat_prin%>" selected> 
							<%response.Write(no_mat_prin)%>						
							</option>
							  <%
	else
	%>
							<option value="<%=co_mat_prin%>"> 
							<%response.Write(no_mat_prin)%>						
							</option>
							<%
	end if		


RSG.MOVENEXT
WEND
END IF%>
                      </select>
                    </div>
                  </div></td>
                  <td width="167"><div align="center">
                    <div id="divPeriodo">
                      <select name="periodo" class="select_style" id="periodo">                    
                        <option value="F" selected>M&eacute;dia Anual</option>
                        <option value="R">Recupera&ccedil;&atilde;o Final</option>
                      </select>
                    </div>
                  </div></td>
                </tr>

<%
check=check+1
Next%> 
              <tr class="form_dado_texto">
                <td colspan="10" align="center"><hr></td>
                </tr>
              <tr class="form_dado_texto">
                <td colspan="10" align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="33%"><div align="center">
                      <input name="Submit2" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','index.asp?nvg=<%response.Write(chave)%>');return document.MM_returnValue" value="Voltar">
                       
                    </div></td>
                    <td width="34%">&nbsp;</td>
                    <td width="33%"><div align="center">
                      <input type="submit" name="Submit5" value="Prosseguir" class="botao_prosseguir">
                    </div></td>
                  </tr>
                </table></td>
                </tr>                             
            </table></form></td>
          </tr>
        </table>
              </td>
          </tr>
		  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
		  <tr>
		    <td height="40" valign="top">&nbsp;</td>
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