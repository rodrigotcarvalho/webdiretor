<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<% 
opt=request.QueryString("opt")
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

enturma = request.form("enturma")
recria_att = request.form("recria_att")
recria_at = request.form("recria_at")

SESSION("recria_att")=recria_att
SESSION("recria_at")=recria_at


if opt="err1" then
unidade_pesquisa=SESSION("unidade_pesquisa")
curso_pesquisa=SESSION("curso_pesquisa")
etapa_pesquisa=SESSION("etapa_pesquisa")
turma_pesquisa=SESSION("turma_pesquisa")
matriculas_digitadas=Session("GuardaMatriculas")
chamadas_digitadas=Session("GuardaChamadasDigitadas")
else
	if enturma="att" then
	SESSION("recria_att")=recria_att
	SESSION("recria_at")=""	
	response.Redirect("bd.asp?et="&enturma)
	elseif enturma="at" then
	unidade_pesquisa=request.Form("unidade")
	curso_pesquisa=request.Form("curso")
	etapa_pesquisa=request.Form("etapa")
	turma_pesquisa=request.Form("turma")
	SESSION("recria_att")=""
	SESSION("recria_at")=recria_at	
	SESSION("unidade_pesquisa")=unidade_pesquisa
	SESSION("curso_pesquisa")=curso_pesquisa
	SESSION("etapa_pesquisa")=etapa_pesquisa
	SESSION("turma_pesquisa")=turma_pesquisa
	response.Redirect("bd.asp?et="&enturma)
	else
	unidade_pesquisa=request.Form("unidade")
	curso_pesquisa=request.Form("curso")
	etapa_pesquisa=request.Form("etapa")
	turma_pesquisa=request.Form("turma")
	SESSION("recria_att")=""
	SESSION("recria_at")=""	
	SESSION("unidade_pesquisa")=unidade_pesquisa
	SESSION("curso_pesquisa")=curso_pesquisa
	SESSION("etapa_pesquisa")=etapa_pesquisa
	SESSION("turma_pesquisa")=turma_pesquisa
	end if
end if


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")		

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2		

		Set CON1_aux = Server.CreateObject("ADODB.Connection") 
		ABRIR1_aux = "DBQ="& CAMINHO_al_aux & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1_aux.Open ABRIR1_aux					
	
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		



%>
<html>
<head>
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../../../js/mm_menu.js"></script>
<script type="text/javascript" src="../../../../js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
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
                                               oHTTPRequest.open("post", "executa.asp?ori=alt&opt=c", true);
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
                                               oHTTPRequest.open("post", "executa.asp?ori=alt&opt=e", true);
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
                                               oHTTPRequest.open("post", "executa.asp?ori=alt&opt=t2", true);
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

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}

//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
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

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
							   
</head>

<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="recuperarCursoLoad(<%response.Write(unidade)%>);recuperarEtapaLoad(<%response.Write(curso)%>);recuperarTurmaLoad(<%response.Write(etapa)%>);recuperarChamadaLoad('<%response.Write(turma)%>')">
<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
            
    <td height="10" colspan="5" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
    </td>
          </tr>
 <%if opt="err1" then%>
             <tr> 
         
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,417,1,0) %>
    </td>
			  </tr>
 <%end if
 if unidade_pesquisa="999990" or curso_pesquisa="999990" or etapa_pesquisa="999990" or turma_pesquisa="999990" then%>	
             <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,421,1,0) %>
    </td>
			  </tr>
 <%end if%>			  	  	  
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,418,0,0) %>
    </td>
			  </tr>			  
          <tr>
      <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td colspan="4"> 
			<form action="altera.asp" method="post" name="busca" id="busca">
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr class="tb_tit"> 
                  <td width="25%"> <div align="center">Unidade </div></td>
                  <td width="25%"> <div align="center">Curso </div></td>
                  <td width="25%"> <div align="center">Etapa </div></td>
                  <td width="25%"> <div align="center">Turma </div></td>
                </tr>
                <tr valign="top"> 
                  <td width="25%" height="10"> <div align="center"> 
                      <select name="unidade" class="borda" onChange="recuperarCurso(this.value)">
					  <%if unidade_pesquisa="999990" then%>
					   <option value="999990" selected></option>
                        <%end if		

		Set RS0u = Server.CreateObject("ADODB.Recordset")
		SQL0u = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0u.Open SQL0u, CON0
NU_Unidade_Check=999999		
While not RS0u.EOF
NU_Unidade = RS0u("NU_Unidade")
NO_Abr = RS0u("NO_Abr")
if NU_Unidade = NU_Unidade_Check then
RS0u.MOVENEXT		
else
unidade_pesquisa=unidade_pesquisa*1
NU_Unidade=NU_Unidade*1
if NU_Unidade = unidade_pesquisa then
%>
                        <option value="<%response.Write(NU_Unidade)%>" selected> 
                        <%response.Write(NO_Abr)%>
                        </option>
                        <%
else
%>
                        <option value="<%response.Write(NU_Unidade)%>"> 
                        <%response.Write(NO_Abr)%>
                        </option>
                        <%						
end if
NU_Unidade_Check = NU_Unidade
RS0u.MOVENEXT
end if
WEND
%>
                      </select>
                    </div></td>
                  <td width="25%" height="10" align="left"> <div align="center" id="divCurso"> 
                      <select name="curso" class="borda" onChange="recuperarEtapa(this.value)">
					  <%if curso_pesquisa="999990" then%>
					   <option value="999990" selected></option>
                        <%end if
		Set RS0ue = Server.CreateObject("ADODB.Recordset")
		SQL0ue = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade_pesquisa
		RS0ue.Open SQL0ue, CON0
		
	
CO_Curso_check="999999"		
While not RS0ue.EOF
CO_Curso = RS0ue("CO_Curso")

if CO_Curso = CO_Curso_check then
RS0ue.MOVENEXT		
else

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
		RS0c.Open SQL0c, CON0
		
NO_Curso = RS0c("NO_Curso")		

if CO_Curso = curso_pesquisa then
%>
                        <option value="<%response.Write(CO_Curso)%>" selected> 
                        <%response.Write(NO_Curso)%>
                        </option>
                        <%
else
%>
                        <option value="<%response.Write(CO_Curso)%>"> 
                        <%response.Write(NO_Curso)%>
                        </option>
                        <%						
end if

CO_Curso_check = CO_Curso
RS0ue.MOVENEXT
end if
WEND
%>
                      </select>
                    </div></td>
                  <td width="25%" height="10" align="center"> <div align="center" id="divEtapa"> 
                      <select name="etapa" class="borda" onChange="recuperarTurma(this.value)">
					  <%if etapa_pesquisa="999990" then%>
					   <option value="999990" selected></option>
                        <%end if						

		Set RS0e = Server.CreateObject("ADODB.Recordset")
		SQL0e = "SELECT * FROM TB_Unidade_Possui_Etapas where CO_Curso ='"& curso_pesquisa &"' AND NU_Unidade="& unidade_pesquisa 
		RS0e.Open SQL0e, CON0
			

while not RS0e.EOF
co_etapa= RS0e("CO_Etapa")

		Set RS3e = Server.CreateObject("ADODB.Recordset")
		SQL3e = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' And CO_Curso ='"& curso_pesquisa &"'" 
		RS3e.Open SQL3e, CON0
		

no_etapa=RS3e("NO_Etapa")

if co_etapa = etapa_pesquisa then
%>
                        <option value="<%response.Write(co_etapa)%>" selected> 
                        <%response.Write(no_etapa)%>
                        </option>
                        <%
else
%>
                        <option value="<%=co_etapa%>"> 
                        <%response.Write(no_etapa)%>
                        </option>
                        <%						
end if
RS0e.MOVENEXT
WEND
%>
                      </select>
                    </div></td>
                  <td width="25%" height="10" align="center"> <div align="center" id="divTurma"> 
                      <select name="turma" class="borda" onChange="submitfuncao()">
					  <%if turma_pesquisa="999990" then%>
					   <option value="999990" selected></option>
                        <%end if
	
		Set RS0t = Server.CreateObject("ADODB.Recordset")
		SQL0t = "SELECT * FROM TB_Turma where NU_Unidade="&unidade_pesquisa&"AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='" & etapa_pesquisa & "' order by CO_Turma" 
		RS0t.Open SQL0t, CON0						
co_turma_check=9999990
while not RS0t.EOF
co_turma= RS0t("CO_Turma")

if co_turma = co_turma_check then
RS0t.MOVENEXT
else

if co_turma = turma_pesquisa then
capacidade= RS0t("NU_Capacidade")
%>
                        <option value="<%response.Write(co_turma)%>" selected> 
                        <%response.Write(co_turma)%>
                        </option>
                        <%
else
%>
                        <option value="<%=co_turma%>"> 
                        <%response.Write(co_turma)%>
                        </option>
                        <%						
end if

co_turma_check = co_turma
RS0t.MOVENEXT
end if
WEND
%>
                      </select>
                    </div></td>
                </tr>
              </table>
		</form>	  
			  </td>
          </tr>
        <form action="bd.asp?et=mt" method="post" name="busca" id="busca" onSubmit="return checksubmit()">		  
          <tr class="tb_subtit"> 
            <td width="100"> <div align="center">Matricula</div></td>
            <td width="600"> <div align="left">Nome</div></td>
            <td width="200"> <div align="center">Situa&ccedil;&atilde;o</div></td>
            <td width="100"> <div align="center">Chamada </div></td>
          </tr>
          <%
'				SQL1 = "SELECT * FROM TB_Matriculas WHERE NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='"&etapa_pesquisa&"' AND CO_Turma='"&turma_pesquisa&"' AND NU_Ano="&ano_letivo
'response.Write(SQL1)
'response.End()
n=-1	
				Set RS1 = Server.CreateObject("ADODB.Recordset")
				SQL1 = "SELECT * FROM TB_Matriculas WHERE CO_Situacao<>'P' AND NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='"&etapa_pesquisa&"' AND CO_Turma='"&turma_pesquisa&"' AND NU_Ano="&ano_letivo&" order by NU_Chamada"
				RS1.Open SQL1, CON1	
				
		IF RS1.EOF THEN		
		ELSE%>
          <%
			WHile Not RS1.EOF
			n=n+1
			cod = RS1("CO_Matricula")
			chamada = RS1("NU_Chamada")
			matricula = RS1("CO_Matricula")			
			data_matricula = RS1("DA_Rematricula")
			situacao = RS1("CO_Situacao")
			
					Set RSsit = Server.CreateObject("ADODB.Recordset")
					SQLsit = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
					RSsit.Open SQLsit, CON0
						
			no_situac=RSsit("TX_Descricao_Situacao")
			
					Set RS2 = Server.CreateObject("ADODB.Recordset")
					SQL2 = "SELECT * FROM TB_Alunos WHERE CO_Matricula="&cod
					RS2.Open SQL2, CON1
			
			nome = RS2("NO_Aluno")
			%>
          <tr> 
            <td width="100"> <div align="center"><font class="form_dado_texto"> 
                <%Response.Write(matricula)%>
                </font> </div></td>
            <td width="600"> <div align="left"><font class="form_dado_texto"> 
                <%Response.Write(nome)%>
                </font> </div></td>
            <td width="200"> <div align="center"><font class="form_dado_texto">
			 <%nome_situac="situac_"&matricula%>
                <input name="<%response.Write(nome_situac)%>" type="hidden" id="<%response.Write(nome_situac)%>" value="<%response.Write(situacao)%>">               
                <%Response.Write(no_situac)%>
                </font> </div></td>
            <td width="100"> <div align="center"><font class="form_dado_texto"> 
                <%nome_chamada="chamada_"&matricula%>
                <select name="<%response.Write(nome_chamada)%>" class="borda">
                  <%for i=1 to capacidade

if opt="err1" then
chamada_digitada=chamadas_digitadas(n)
chamada_digitada=chamada_digitada*1
	if i=chamada_digitada then
	%>
					  <option value="<%response.Write(i)%>" selected> 
					  <%response.Write(i)%>
					  </option>
					  <%
	else
	%>
					  <option value="<%response.Write(i)%>"> 
					  <%response.Write(i)%>
					  </option>
					  <%
	end if
else
	if i=chamada then
	%>
					  <option value="<%response.Write(i)%>" selected> 
					  <%response.Write(i)%>
					  </option>
					  <%
	else
	%>
					  <option value="<%response.Write(i)%>"> 
					  <%response.Write(i)%>
					  </option>
					  <%
	end if
end if
next%>
                </select>
                </font> </div></td>
          </tr>
          <%
			RS1.Movenext
			Wend
		end if	  			
				Set RS1b = Server.CreateObject("ADODB.Recordset")
				SQL1b = "SELECT * FROM TB_Matriculas WHERE CO_Situacao='P' AND NU_Unidade="&unidade_pesquisa&" AND CO_Curso='"&curso_pesquisa&"' AND CO_Etapa='"&etapa_pesquisa&"' AND CO_Turma='"&turma_pesquisa&"' AND NU_Ano="&ano_letivo&" order by DA_Rematricula"
				RS1b.Open SQL1b, CON1			
		IF RS1b.EOF THEN		
		ELSE%>
          <%
			WHile Not RS1b.EOF
			n=n+1
			cod = RS1b("CO_Matricula")
			chamada = RS1b("NU_Chamada")
			matricula = RS1b("CO_Matricula")			
			data_matricula = RS1b("DA_Rematricula")
			situacao = RS1b("CO_Situacao")
			
					Set RSsit = Server.CreateObject("ADODB.Recordset")
					SQLsit = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
					RSsit.Open SQLsit, CON0
						
			no_situac=RSsit("TX_Descricao_Situacao")
			
					Set RS2 = Server.CreateObject("ADODB.Recordset")
					SQL2 = "SELECT * FROM TB_Alunos WHERE CO_Matricula="&cod
					RS2.Open SQL2, CON1
			
			nome = RS2("NO_Aluno")
			%>
          <tr> 
            <td width="100"> <div align="center"><font class="form_dado_texto"> 
                <%Response.Write(matricula)%>
                </font> </div></td>
            <td width="600"> <div align="left"><font class="form_dado_texto"> 
                <%Response.Write(nome)%>
                </font> </div></td>
            <td width="200"> <div align="center"><font class="form_dado_texto">
			 <%nome_situac="situac_"&matricula%>
                <input name="<%response.Write(nome_situac)%>" type="hidden" id="<%response.Write(nome_situac)%>" value="<%response.Write(situacao)%>"> 			 
                <%Response.Write(no_situac)%>
                </font> </div></td>
            <td width="100"> <div align="center"><font class="form_dado_texto"> 
                <%nome_chamada="chamada_"&matricula%>
                <select name="<%response.Write(nome_chamada)%>" class="borda">
                  <%for i=1 to capacidade

if opt="err1" then
chamada_digitada=chamadas_digitadas(n)
chamada_digitada=chamada_digitada*1
	if i=chamada_digitada then
	%>
					  <option value="<%response.Write(i)%>" selected> 
					  <%response.Write(i)%>
					  </option>
					  <%
	else
	%>
					  <option value="<%response.Write(i)%>"> 
					  <%response.Write(i)%>
					  </option>
					  <%
	end if
else
	if i=chamada then
	%>
					  <option value="<%response.Write(i)%>" selected> 
					  <%response.Write(i)%>
					  </option>
					  <%
	else
	%>
					  <option value="<%response.Write(i)%>"> 
					  <%response.Write(i)%>
					  </option>
					  <%
	end if
end if
next%>
                </select>

                </font> </div></td>
          </tr>
          <%
			RS1b.Movenext
			Wend		  
		end if	  			
%>
          <tr> 
            <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td colspan="3"><hr></td>
                </tr>
                <tr> 
                  <td width="33%"><div align="center">
                      <input name="SUBMIT5" type=button class="borda_bot3" onClick="MM_goToURL('parent','index.asp?ori=2&nvg=WS-MA-MA-ETA');return document.MM_returnValue" value="Voltar">
                    </div></td>
                  <td width="34%">&nbsp;</td>
                  <td width="33%">
<div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      <input name="Button22" type="submit" class="borda_bot2"value="Confirmar">
                      </font></div></td>
                </tr>
              </table></td>
          </tr></form>		 
        </table></td>
    </tr>
 
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

</body>
<script type="text/javascript">
<!--
//  initInputHighlightScript();
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