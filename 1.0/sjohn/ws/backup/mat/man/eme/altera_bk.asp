<%'On Error Resume Next%>
<% Response.Charset="ISO-8859-1" %>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->


<%
opt= request.QueryString("opt")
aluno_novo_dados="n"

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo")
ano_letivo_real = ano_letivo
sistema_local=session("sistema_local")
ori = request.QueryString("or")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
cod= request.QueryString("cod")	

z = request.QueryString("z")
erro = request.QueryString("e")
vindo = request.QueryString("vd")
obr = request.QueryString("o")




nvg = session("chave")
chave=nvg
session("chave")=chave
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano_info=nivel&"-"&chave&"-"&ano_letivo

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1

		Set CON_al = Server.CreateObject("ADODB.Connection") 
		ABRIR_al = "DBQ="& CAMINHOa& ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_al.Open ABRIR_al

		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0


 call navegacao (CON,chave,nivel)
navega=Session("caminho")	

 Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Autoriz_Usuario_Grupo Where CO_Usuario = "&cod
		RS2.Open SQL2, CON
		
if RS2.EOF then

else		
co_grupo=RS2("CO_Grupo")
End if
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod
		RS.Open SQL, CON1
		
		
codigo = RS("CO_Matricula")
nome_prof = RS("NO_Aluno")
col_origem= RS("NO_Colegio_Origem")
cursada= RS("NO_Serie_Cursada")
uf_cursada= RS("SG_UF_Cursada")
cid_cursada= RS("CO_Municipio_Cursada")
resp_fin= RS("TP_Resp_Fin")
resp_ped= RS("TP_Resp_Ped")
entrada= RS("DA_Entrada_Escola")
cadastro= RS("DA_Cadastro")

		Set RS_mat = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Matriculas WHERE NU_Ano="&ano_letivo&" AND CO_Matricula ="& cod
		RS_mat.Open SQL, CON1

if RS_mat.EOF Then
aluno_novo="s"

		Set RS_aluno_novo = Server.CreateObject("ADODB.Recordset")
		SQL_aluno_novo= "SELECT * FROM TB_Matriculas WHERE CO_Matricula ="& cod &" order by NU_Ano"
		RS_aluno_novo.Open SQL_aluno_novo, CON1
		
while not RS_aluno_novo.EOF
aluno_novo_dados="s"
unidade= RS_aluno_novo("NU_Unidade")
curso= RS_aluno_novo("CO_Curso")
etapa= RS_aluno_novo("CO_Etapa")
etapa_ck=etapa
turma= RS_aluno_novo("CO_Turma")
RS_aluno_novo.movenext
wend		


else
aluno_novo="n"
ano_aluno = RS_mat("NU_Ano")

rematricula_atual = RS_mat("DA_Rematricula")
unidade= RS_mat("NU_Unidade")
curso= RS_mat("CO_Curso")
etapa= RS_mat("CO_Etapa")
turma= RS_mat("CO_Turma")
situacao_atual= RS_mat("CO_Situacao")

end if

if aluno_novo_dados="s" or aluno_novo="n" then
'		response.Write "SELECT * FROM TB_Unidade_Possui_Etapas WHERE NU_Unidade ="& unidade&" AND CO_Curso='"&curso&"' AND CO_Etapa='"&etapa&"'"


		Set RS_prox = Server.CreateObject("ADODB.Recordset")
		SQL_prox = "SELECT * FROM TB_Unidade_Possui_Etapas WHERE NU_Unidade ="& unidade&" AND CO_Curso='"&curso&"' AND CO_Etapa='"&etapa&"'"
		RS_prox.Open SQL_prox, CON0

prox_unidade= RS_prox("Prox_NU_Unidade")
prox_curso= RS_prox("Prox_CO_Curso")
prox_etapa= RS_prox("Prox_CO_Etapa")

		'response.Write "SELECT * FROM TB_Unidade_Possui_Etapas WHERE NU_Unidade ="& unidade&" AND CO_Curso='"&curso&"' AND CO_Etapa='"&etapa&"'"


call GeraNomes("PORT",unidade,curso,etapa,CON0)
no_unidade = session("no_unidades")
no_grau = session("no_grau")
no_etapa = session("no_serie")
end if

		Set RSCONTA = Server.CreateObject("ADODB.Recordset")
		SQLA = "SELECT * FROM TB_Contatos WHERE TP_Contato ='ALUNO' And CO_Matricula ="& cod
		RSCONTA.Open SQLA, CONCONT



		Set RSano = Server.CreateObject("ADODB.Recordset")
		SQLano = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo ='"&ano_letivo&"'"
		RSano.Open SQLano, CON

situac_mat_post=RSano("ST_Mat_Post_Autoriz")




if situac_mat_post=TRUE then
ano_que_vem=RSano("NU_Ano_Posterior")
ano_letivo_combo=ano_que_vem
unidade_combo=prox_unidade
curso_combo=prox_curso
etapa_combo=prox_etapa
etapa_ck=prox_etapa
pre_matricula="aberta"
else
ano_que_vem=ano_letivo
if aluno_novo="s" then
ano_letivo_combo=ano_letivo
unidade_combo=prox_unidade
curso_combo=prox_curso
etapa_combo=prox_etapa
etapa_ck=prox_etapa
ano_letivo_combo="combo"
pre_matricula="aberta"

if unidade_combo="" or isnull(unidade_combo) then
unidade_combo=unidade
end if
if curso_combo="" or isnull(curso_combo) then
curso_combo=curso
end if
if etapa_combo="" or isnull(etapa_combo) then
etapa_combo=etapa
end if

else
if situacao_atual="P" or situacao_atual="L" then
unidade_combo=unidade
curso_combo=curso
etapa_combo=etapa
etapa_ck=etapa
ano_letivo_combo=ano_que_vem
pre_matricula="aberta"


elseif situacao_atual<>"P" and situacao_atual<>"L"then
pre_matricula="fechada"
end if
end if
end if


'response.Write(situac_mat_post&"-"&aluno_novo&">>"&curso_combo)
'if situac_mat_post=FALSE and aluno_novo="s" then
'ano_letivo_combo=ano_letivo
'elseif situac_mat_post=FALSE and aluno_novo="n" and (situacao_atual="P" or situacao_atual="L") then
'unidade_combo=unidade
'curso_combo=curso
'etapa_combo=etapa
'etapa_ck=etapa
'ano_letivo_combo=ano_que_vem
'pre_matricula="aberta"

'elseif situac_mat_post=FALSE and aluno_novo="n" and (situacao_atual<>"P" and situacao_atual<>"L")then
'pre_matricula="fechada"
'elseif situac_mat_post=FALSE and aluno_novo="s" then
'unidade_combo=prox_unidade
'curso_combo=prox_curso
'etapa_combo=prox_etapa
'etapa_ck=prox_etapa
'ano_letivo_combo="combo"
'pre_matricula="aberta"
'elseif curso_combo=2 and etapa_combo=3 then
'unidade_combo=prox_unidade
'curso_combo=prox_curso
'etapa_combo=prox_etapa
'etapa_ck=prox_etapa
'ano_letivo_combo="combo"
'pre_matricula="aberta"
'end if
session("c_pub")=curso_combo

Select case etapa_combo
case "0M1"
etapa_altera=901
case "0M2"
etapa_altera=902
case "JD1"
etapa_altera=903
case "JD2"
etapa_altera=904
case "JD3"
etapa_altera=905
case else
etapa_altera=etapa_combo
end select

session("e_pub")=etapa_combo

session("c_ck")=curso_combo
session("e_ck")=etapa_ck
session("t_ck")=turma




Call LimpaVetor2

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
<script type="text/javascript" src="../../../../js/global.js"></script>
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

function checksubmit()
{

  if (document.formulario.curso.value == "999990")
  {    alert("Por favor selecione um curso!")
    document.formulario.curso.focus()
    return false
  }
  if (document.formulario.etapa.value == "999990")
  {    alert("Por favor selecione uma etapa!")
    document.formulario.etapa.focus()
    return false
  }
  if (document.formulario.turma.value == "999990")
  {    alert("Por favor selecione uma turma!")
    document.formulario.turma.focus()
    return false
  }
  return true
}
//-->
</script>
                         
<script>
<!--
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
                                               oHTTPRequest.open("post", "executa.asp?opt=c", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
document.all.divEtapa.innerHTML ="<select name=etapa class=borda><option value=999990 selected></option></select>"
document.all.divTurma.innerHTML = "<select name=turma class=borda><option value=999990 selected></option></select>"
                                                           }
                                               }
                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?opt=e", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select name=turma class=borda><option value=999990 selected></option></select>"

                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?opt=t", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t																	   
                                                           }
                                               }
                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }    
								   
//==========================================================================================================================================
						 function recuperarCursoLoad(uTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?opt=c", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
                                                           }
                                               }
                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapaLoad(cTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?opt=e", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e

                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurmaLoad(eTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?opt=t", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t																	   
                                                           }
                                               }
                                               oHTTPRequest.send("e_load=" + eTipo);
                                   } 								                       

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0" <%if (aluno_novo="s" and aluno_novo_dados="s") or aluno_novo="n" then%> onLoad="recuperarCursoLoad(<%response.Write(unidade_combo)%>);recuperarEtapaLoad(<%response.Write(curso_combo)%>);recuperarTurmaLoad(<%response.Write(etapa_altera)%>)" <%end if%>>
<% call cabecalho (nivel)
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
 <%if opt="ok" then%> 
              <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,401,2,0) %>
    </td>
  </tr>
 <%elseif opt="ok1" then%> 
              <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,9709,2,0) %>
    </td>
  </tr>  
 <%end if%> 
  <%if pre_matricula="fechada" then%> 
              <tr>         
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,403,1,0) %>
    </td>
  </tr>
 <%end if%> 
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,302,0,0) 
	  %>
    </td>
  </tr>
<tr>

            <td valign="top"> 
			
<FORM name="formulario" METHOD="POST" ACTION="bd.asp?opt=i" onSubmit="return checksubmit()">
        <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
          <tr> 
            <td width="841" class="tb_tit"
>Principais Dados do Aluno</td>
            <td width="151" class="tb_tit"
> </td>
            <td width="2" class="tb_tit"
></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td height="10" colspan="3"> <table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="14%" height="10"><font class="form_dado_texto">Matr&iacute;cula</font></td>
                  <td width="2%"><div align="center">:</div></td>
                  <td width="22%" height="10"><font class="form_corpo"> 
                    <input name="cod" type="hidden" id="cod" value="<%=codigo%>">
                    <%response.Write(codigo)%>
                    <input name="pre_matricula" type="hidden" id="acesso3" value="<%=pre_matricula%>">
                    <input name="co_usr_prof2" type="hidden" id="co_usr_prof3" value="<% =co_usr_prof%>">
                    <input name="tp2" type="hidden" id="tp" value="P">
                    </font></td>
                  <td width="15%" height="10"><font class="form_dado_texto">Nome</font></td>
                  <td width="1%"> <div align="center">:</div></td>
                  <td width="46%" height="10"><font class="form_corpo"> 
                    <%response.Write(nome_prof)%>
                    </font></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td colspan="3" class="tb_tit"
>Familiares</td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="3"> <table width="100%" border="0" cellspacing="0">
                <tr bgcolor="#FFFFFF" class="tb_corpo"
> 
                  <td width="16%" height="10"> <div align="left"><font class="form_dado_texto">Rela&ccedil;&atilde;o</font></div></td>
                  <td width="68%" height="10"> <font class="form_dado_texto">Nome</font></td>
                  <td width="16%"><div align="center"><font class="form_dado_texto"> 
                      Vinculado a Matr&iacute;cula</font></div></td>
                </tr>
                <%

		Set RSCONTPR = Server.CreateObject("ADODB.Recordset")
		SQLCONTPR = "SELECT * FROM TB_Tipo_Contatos Order by NU_Prioridade_Combo"
		RSCONTPR.Open SQLCONTPR, CON0
		
while not RSCONTPR.EOF
tp_resp = RSCONTPR("TP_Contato")
no_tp_resp = RSCONTPR("TX_Descricao")

		Set RSCONT = Server.CreateObject("ADODB.Recordset")
		SQLCONT = "SELECT * FROM TB_Contatos WHERE CO_Matricula ="& cod &" And TP_Contato = '"&tp_resp&"'"
		RSCONT.Open SQLCONT, CONCONT

if RSCONT.EOF then
RSCONTPR.MOVENEXT
else
resp= RSCONT("NO_Contato")

%>
                <tr bgcolor="#FFFFFF" class="tb_corpo"
> 
                  <td width="16%" height="10"> <div align="left"><font class="form_corpo"> 
                      <%response.Write(no_tp_resp)%>
                      </font></div></td>
                  <td width="68%" height="10"> <font class="form_corpo"> 
                    <!--  <a href="contatos.asp?or=01&cod=<%=cod%>&tp=<%=tp_resp%>&vd=<%=vindo%>&o=<%=obr%>" class="ativos">-->
                    <%response.Write(resp)%>
                    <!--  </a>  -->
                    </font></td>
                  <td width="16%"><div align="center"></div></td>
                </tr>
                <%RSCONTPR.MOVENEXT
end if
WEND
%>
              </table></td>
          </tr>
          <tr> 
            <td colspan="3" class="tb_tit">Respons&aacute;veis</td>
          </tr>
          <tr> 
            <td colspan="3"><table width="100%" border="0" cellspacing="0">
                <tr bgcolor="#FFFFFF" class="tb_corpo"
> 
                  <td width="14%" height="10"> <div align="left"><font class="form_dado_texto"> 
                      Pedag&oacute;gico</font></div></td>
                  <td width="2%"> <div align="center">:</div></td>
                  <td width="22%" height="10"> <select name="resp_ped" class="borda" id="resp_ped">
                      <%

		Set RSCONTPR = Server.CreateObject("ADODB.Recordset")
		SQLCONTPR = "SELECT * FROM TB_Tipo_Contatos Order by NU_Prioridade_Combo"
		RSCONTPR.Open SQLCONTPR, CON0
		
while not RSCONTPR.EOF
tp_resp = RSCONTPR("TP_Contato")
		Set RSCONT = Server.CreateObject("ADODB.Recordset")
		SQLCONT = "SELECT * FROM TB_Contatos WHERE CO_Matricula ="& cod &" And TP_Contato = '"&tp_resp&"'"
		RSCONT.Open SQLCONT, CONCONT

if RSCONT.EOF then
RSCONTPR.MOVENEXT
else
resp= RSCONT("NO_Contato")
if resp_ped=tp_resp then
%>
                      <option value="<%response.Write(tp_resp)%>" selected> 
                      <%response.Write(resp)%>
                      </option>
                      <%else%>
                      <option value="<%response.Write(tp_resp)%>"> 
                      <%response.Write(resp)%>
                      </option>
                      <%
end if
RSCONTPR.MOVENEXT
end if
WEND
%>
                    </select></td>
                  <td width="15%" height="10"> <div align="left"><font class="form_dado_texto"> 
                      Financeiro</font></div></td>
                  <td width="1%"> <div align="center">:</div></td>
                  <td width="46%" height="10"> <select name="resp_fin" class="borda" id="resp_ped">
                      <%
		Set RSCONTPR = Server.CreateObject("ADODB.Recordset")
		SQLCONTPR = "SELECT * FROM TB_Tipo_Contatos Order by NU_Prioridade_Combo"
		RSCONTPR.Open SQLCONTPR, CON0
		
while not RSCONTPR.EOF
tp_resp = RSCONTPR("TP_Contato")
		Set RSCONT = Server.CreateObject("ADODB.Recordset")
		SQLCONT = "SELECT * FROM TB_Contatos WHERE CO_Matricula ="& cod &" And TP_Contato = '"&tp_resp&"'"
		RSCONT.Open SQLCONT, CONCONT

if RSCONT.EOF then
RSCONTPR.MOVENEXT
else
resp= RSCONT("NO_Contato")
if resp_fin=tp_resp then
%>
                      <option value="<%response.Write(tp_resp)%>" selected> 
                      <%response.Write(resp)%>
                      </option>
                      <%else%>
                      <option value="<%response.Write(tp_resp)%>"> 
                      <%response.Write(resp)%>
                      </option>
                      <%
end if
RSCONTPR.MOVENEXT
end if
WEND
%>
                    </select></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td colspan="3" class="tb_tit"
>Dados Escolares </td>
          </tr>
          <tr> 
            <td colspan="3"><table width="100%" border="0" cellspacing="0">
                <tr class="tb_corpo"
> 
                  <td height="10" colspan="2"> <div align="left"><font class="form_dado_texto"> 
                      Col&eacute;gio de Origem</font></div></td>
                  <td><div align="center">:</div></td>
                  <td colspan="3"><font class="form_corpo"> 
                    <%response.Write(col_origem)%>
                    </font></td>
                  <td height="10" colspan="9">&nbsp;</td>
                </tr>
                <tr class="tb_corpo"
> 
                  <td height="10" colspan="2"><font class="form_dado_texto">Etapa 
                    cursada </font></td>
                  <td><div align="center">:</div></td>
                  <td colspan="3"><font class="form_corpo"> 
                    <%response.Write(cursada)%>
                    </font></td>
                  <td height="10"><font class="form_dado_texto">Local</font></td>
                  <td><div align="center">:</div></td>
                  <td height="10" colspan="7"><font class="form_corpo"> 
                    <%
					IF ISNULL(cid_cursada) THEN					
					response.Write(cid_cursada)
					ELSEIF ISNULL(uf_cursada) THEN
					response.Write(uf_cursada)					
					ELSE					
					response.Write(cid_cursada&"/"&uf_cursada)					
END IF					
					%>
                    </font></td>
                </tr>
                <tr class="tb_corpo"
> 
                  <td height="10" colspan="2"> <div align="left"><font class="form_dado_texto">Data 
                      do Cadastro</font></div></td>
                  <td><div align="center">:</div></td>
                  <td colspan="3"><font class="form_corpo"> 
                    <%response.Write(cadastro)%>
                    </font></td>
                  <td height="10"> <div align="left"><font class="form_dado_texto">Data 
                      de Entrada na Escola</font></div></td>
                  <td width="5"><div align="center">:</div></td>
                  <td height="10" colspan="7"><font class="form_corpo"> 
                    <%response.Write(entrada)%>
                    </font></td>
                </tr>
                <tr class="tb_corpo"
> 
                  <td width="71" height="26"><div align="center"><font class="form_dado_texto"> 
                      </font><font class="form_dado_texto">Ano Letivo </font></div></td>
                  <td width="62" height="26"><div align="center"><font class="form_dado_texto">Matr&iacute;cula</font></div></td>
                  <td width="13"><div align="center"></div></td>
                  <td width="10"><div align="center"></div></td>
                  <td width="88"><div align="center"><font class="form_dado_texto">Cancelamento</font></div></td>
                  <td width="118"><div align="center"><font class="form_dado_texto">Situa&ccedil;&atilde;o</font></div></td>
                  <td width="154"><div align="center"><font class="form_dado_texto">Unidade</font></div></td>
                  <td><div align="center"></div></td>
                  <td width="2">&nbsp;</td>
                  <td width="173"><font class="form_dado_texto">Curso</font></td>
                  <td width="79" height="26"> <div align="center"><font class="form_dado_texto"> 
                      Etapa</font></div></td>
                  <td width="2" height="26">&nbsp;</td>
                  <td width="97"><div align="center"><font class="form_dado_texto">Turma 
                      </font></div></td>
                  <td width="71" height="26"> <div align="center"><font class="form_dado_texto"> 
                      Chamada</font></div></td>
                  <td width="23" height="26">&nbsp;</td>
                </tr>
                <%

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Matriculas WHERE CO_Matricula ="& cod
		RS.Open SQL, CON1

while not RS.EOF
ano_aluno = RS("NU_Ano")

rematricula = RS("DA_Rematricula")
situacao = RS("CO_Situacao")
encerramento= RS("DA_Encerramento")
unidade_dadesc= RS("NU_Unidade")
curso_dadesc= RS("CO_Curso")
etapa_dadesc= RS("CO_Etapa")
turma_dadesc= RS("CO_Turma")
cham= RS("NU_Chamada")

call GeraNomes("PORT",unidade_dadesc,curso_dadesc,etapa_dadesc,CON0)
no_unidade_dadesc = session("no_unidades")
no_curso_dadesc = session("no_grau")
no_etapa_dadesc = session("no_serie")


%>
                <tr class="tb_corpo"
			  
> 
                  <td height="10"> <div align="center"><font class="form_corpo"> 
                      <%response.Write(ano_aluno)%>
                      </font></div></td>
                  <td height="10"><div align="center"><font class="form_corpo"> 
                      <%response.Write(rematricula)%>
                      </font></div></td>
                  <td width="13"><div align="center"></div></td>
                  <td width="10"><div align="center"></div></td>
                  <td width="88"><div align="center"><font class="form_corpo"> 
                      <%response.Write(encerramento)%>
                      </font></div></td>
                  <td width="118"><div align="center"><font class="form_corpo"> 
                      <%
					
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
		RSCONTST.Open SQLCONTST, CON0
	
if RSCONTST.EOF then
else
							
				no_situacao = RSCONTST("TX_Descricao_Situacao")	
					response.Write(no_situacao)
end if					
					%>
                      </font></div></td>
                  <td height="10"><div align="center"><font class="form_corpo"> 
                      <%response.Write(no_unidade_dadesc)%>
                      </font></div></td>
                  <td><div align="center"></div></td>
                  <td>&nbsp;</td>
                  <td><div align="left"><font class="form_corpo"> 
                      <%response.Write(no_curso_dadesc)%>
                      </font></div></td>
                  <td height="10"> <div align="center"><font class="form_corpo"> 
                      <%response.Write(no_etapa_dadesc)%>
                      </font></div></td>
                  <td height="10">&nbsp;</td>
                  <td><div align="center"><font class="form_corpo"> 
                      <%response.Write(turma_dadesc)%>
                      </font></div></td>
                  <td height="10"> <div align="center"><font class="form_corpo"> 
                      <%response.Write(cham)%>
                      </font></div></td>
                  <td height="10">&nbsp;</td>
                </tr>
                <%RS.MOVENEXT
WEND
%>
              </table></td>
          </tr>
          <tr> 
            <td valign="top" colspan="3"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <%if pre_matricula="fechada" then
matricula="no"
else
%>
                <tr> 
                  <td colspan="12" class="tb_tit">Caracter&iacute;sticas da Matr&iacute;cula</td>
                </tr>
                <tr valign="top" bgcolor="#FFFFFF"> 
                  <td colspan="12"> <table width="100%" border="0" cellspacing="0">
                      <tr bgcolor="#FFFFFF" class="tb_corpo"> 
                        <td width="14%" height="10"><font class="form_dado_texto">Tipo 
                          de Matr&iacute;cula</font></td>
                        <td width="2%"> <div align="center">:</div></td>
                        <td width="22%"><font class="form_corpo"> Por Curso</font></td>
                        <td width="15%" height="10">&nbsp;</td>
                        <td width="1%" height="10"> <div align="center"></div></td>
                        <td width="46%" height="10">&nbsp;</td>
                      </tr>
                      <tr class="tb_corpo"
> 
                        <td height="10" colspan="6" class="tb_subtit">Matr&iacute;cula 
                          sugerida <input name="sit_nova_mat" type="hidden" id="sit_nova_mat" value="P"> 
                          <font class="form_corpo"> 
                          <input name="aluno_novo" type="hidden" id="aluno_novo2" value="<%response.Write(aluno_novo)%>">
                          <input name="ano_letivo" type="hidden" id="ano_letivo_mat2" value="<%response.Write(ano_letivo_combo)%>">
                          <input name="ult_ano_aluno" type="hidden" id="ult_ano_aluno" value="<%response.Write(ano_aluno)%>">
                          </font> </td>
                      </tr>
                    </table>
                <tr> 
                  <td width="74" height="26" bgcolor="#FFFFFF" class="tb_corpo"> 
                    <div align="center"><font class="form_dado_texto"> </font><font class="form_dado_texto">Ano 
                      Letivo </font></div></td>
                  <td width="65" height="26" bgcolor="#FFFFFF" class="tb_corpo"> 
                    <div align="center"><font class="form_dado_texto">Matr&iacute;cula</font></div></td>
                  <td width="14" bgcolor="#FFFFFF" class="tb_corpo"> <div align="center"></div></td>
                  <td width="10" bgcolor="#FFFFFF" class="tb_corpo"> <div align="center"></div></td>
                  <td width="97" bgcolor="#FFFFFF" class="tb_corpo">&nbsp;</td>
                  <td width="123" bgcolor="#FFFFFF" class="tb_corpo"> <div align="center"><font class="form_dado_texto">Situa&ccedil;&atilde;o</font></div></td>
                  <td width="163"> <div align="center"><font class="form_dado_texto">Unidade</font> 
                    </div></td>
                  <td width="173"> <div align="left"><font class="form_dado_texto">Curso</font> 
                    </div></td>
                  <td width="82"> <div align="center"><font class="form_dado_texto">Etapa</font> 
                    </div></td>
                  <td width="121"> <div align="center"><font class="form_dado_texto">Turma 
                      </font> </div></td>
                  <td width="54">&nbsp;</td>
                  <td width="24">&nbsp;</td>
                </tr>
                <tr valign="top"> 
                  <td height="40" bgcolor="#FFFFFF" class="tb_corpo"> 
                    <div align="center"><font class="form_corpo"> 
                      <%
					  if ano_letivo_combo="combo" then
					  
					  %>
                      <select name="ano_letivo_mat" class="borda" id="ano_letivo_mat">
                        <%
		ano_letivo=ano_letivo*1				
		ano_letivo_anterior=ano_letivo-1	
		ano_letivo_posterior=ano_letivo+1	
		
		Set RS0combo = Server.CreateObject("ADODB.Recordset")
		SQL0combo = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo >='"&ano_letivo_anterior&"' order by NU_Ano_Letivo"
		RS0combo.Open SQL0combo, CON
		while not RS0combo.EOF 
		ano_bd=RS0combo("NU_Ano_Posterior")
		ano_bd=ano_bd*1
				if ano_letivo_posterior=ano_bd then%>
                        <option value="<%=ano_bd%>" selected><%=ano_bd%></option>
                        <%else%>
                        <option value="<%=ano_bd%>"><%=ano_bd%></option>
                        <%end if
		RS0combo.MOVENEXT
		WEND 		
				%>
                      </select>
                      <%
					  else%>
<input name="ano_letivo_mat" type="hidden" id="ano_letivo_mat2" value="<%response.Write(ano_letivo_combo)%>">					  
					 <% response.Write(ano_letivo_combo)%>
                      <%end if%>
                      </font></div></td>
                  <td width="65" height="40" bgcolor="#FFFFFF" class="tb_corpo"> 
                    <div align="center"><font class="form_corpo"> 
                      <%
ano = DatePart("yyyy", now)
mes_de = DatePart("m", now) 
dia_de = DatePart("d", now) 

if dia_de<10 then
dia_de="0"&dia_de
end if

if mes_de<10 then
mes_de="0"&mes_de
end if
					
data_cadastro=dia_de&"/"&mes_de&"/"&ano						
							
							response.Write(data_cadastro)%>
                      </font></div></td>
                  <td width="14" height="40" bgcolor="#FFFFFF" class="tb_corpo"> 
                    <div align="center"></div></td>
                  <td width="10" height="40" bgcolor="#FFFFFF" class="tb_corpo"> 
                    <div align="center"></div></td>
                  <td width="97" height="40" bgcolor="#FFFFFF" class="tb_corpo">&nbsp;</td>
                  <td width="123" height="40" bgcolor="#FFFFFF" class="tb_corpo"> 
                    <div align="center"><font class="form_corpo"> 
                      Pr&eacute;-matriculado </font></div></td>
                  <td width="163" height="40"> 
                    <div align="center"> 
                      <select name="unidade" class="borda" onChange="recuperarCurso(this.value)">
                        <%		
						
if aluno_novo="s" and aluno_novo_dados="n" then%>

 <option value="999990" selected></option>
<%
end if						

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0
NU_Unidade_Check=999999		
While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")
if NU_Unidade = NU_Unidade_Check then
RS0.MOVENEXT		
unidade=unidade*1
NU_Unidade=NU_Unidade*1		
elseif NU_Unidade = unidade_dadesc then
%>
                        <option value="<%response.Write(NU_Unidade)%>" selected> 
                        <%response.Write(NO_Abr)%>
                        </option>
                        <%
NU_Unidade_Check = NU_Unidade
RS0.MOVENEXT							  
else%>
                        <option value="<%response.Write(NU_Unidade)%>"> 
                        <%response.Write(NO_Abr)%>
                        </option>
                        <%

NU_Unidade_Check = NU_Unidade
RS0.MOVENEXT
end if
WEND
%>
                      </select>
                    </div></td>
                  <td width="173" height="40" align="left"> 
                    <div id="divCurso"></div></td>
                  <td width="82" height="40" align="center"> 
                    <div id="divEtapa"></div></td>
                  <td width="121" height="40" align="center"> 
                    <div id="divTurma"></div></td>
                  <td width="54" height="40">&nbsp;</td>
                  <td width="24" height="40">&nbsp;</td>
                </tr>
                <%end if%>
                <tr> 
                  <td colspan="12" class="tb_tit">Documentos Entregues</td>
                </tr>
                <tr> 
                  <td colspan="12"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="33%"><div align="center"><font class="form_dado_texto">Documento 
                            </font></div></td>
                        <td width="34%"><div align="center"><font class="form_dado_texto">Situa&ccedil;&atilde;o 
                            </font></div></td>
                        <td width="33%"><div align="center"><font class="form_dado_texto">Data 
                            </font></div></td>
                      </tr>
                      <%
		Set RSdt = Server.CreateObject("ADODB.Recordset")
		SQLdt = "SELECT * FROM TB_Documentos_Matricula order by NO_Documento"
		RSdt.Open SQLdt, CON0

while not RSdt.EOF
co_doc_mat=RSdt("CO_Documento")
no_doc_mat=RSdt("NO_Documento")


		Set RSde = Server.CreateObject("ADODB.Recordset")
		SQLde = "SELECT * FROM TB_Documentos_Entregues where CO_Documento='"&co_doc_mat&"' And CO_Matricula="&cod
		RSde.Open SQLde, CON0

IF RSde.EOF then
%>
                      <tr> 
                        <td width="33%"><div align="center"><font class="form_corpo"> 
                            <%response.Write(no_doc_mat)%>
                            </font></div></td>
                        <td width="34%"><div align="center"> 
                            <table width="50%" border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr> 
                                <td width="8%"><input name="<%response.Write(co_doc_mat)%>" type="radio" value="S"></td>
                                <td width="38%"><font class="form_corpo"> Entregue 
                                  </font></td>
                                <td width="7%"><input type="radio" name="<%response.Write(co_doc_mat)%>" value="N" checked></td>
                                <td width="47%"><font class="form_corpo"> Pendente 
                                  </font></td>
                              </tr>
                            </table>
                          </div></td>
                        <td width="33%"><div align="center"><font class="form_corpo"> 
                            </font></div></td>
                      </tr>
                      <%else
data_ent=RSde("DA_Entrega_Documento")
%>
                      <tr> 
                        <td width="33%"><div align="center"><font class="form_corpo"> 
                            <%response.Write(no_doc_mat)%>
                            </font></div></td>
                        <td width="34%"><div align="center"> 
                            <table width="50%" border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr> 
                                <td width="8%"><input name="<%response.Write(co_doc_mat)%>" type="radio" value="S" checked></td>
                                <td width="38%"><font class="form_corpo"> Entregue 
                                  </font></td>
                                <td width="7%"><input type="radio" name="<%response.Write(co_doc_mat)%>" value="N"></td>
                                <td width="47%"><font class="form_corpo"> Pendente 
                                  </font></td>
                              </tr>
                            </table>
                          </div></td>
                        <td width="33%"><div align="center"><font class="form_corpo"> 
                            <%response.Write(data_ent)%>
                            </font></div></td>
                      </tr>
                      <%
end if
RSdt.Movenext
wend
%>
                    </table></td>
                </tr>
                <tr> 
                  <td colspan="12"><hr width="1000"></td>
                </tr>
                <tr> 
                  <td colspan="12"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="25%"> <div align="center"> 
                            <input name="SUBMIT5" type=button class="borda_bot3" onClick="MM_goToURL('parent','index.asp?nvg=WS-MA-MA-EME');return document.MM_returnValue" value="Voltar">
                          </div></td>
                        <td width="25%"> <div align="center"> </div></td>
                        <td width="25%"> <div align="center"> </div></td>
                        <td width="25%"> <div align="center">
						<%if pre_matricula="fechada" then 
						else%>
                            <input name="SUBMIT" type=SUBMIT class="borda_bot2" value="Confirmar">
					<%end if%>		
                          </div></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td valign="top" colspan="3">&nbsp;</td>
          </tr>
        </table>
      </form></td>
          </tr>
		  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.gif" width="1000" height="40"></td>
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