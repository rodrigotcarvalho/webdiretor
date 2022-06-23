<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->

<%
'Session.LCID = 1046
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
nvg=request.QueryString("nvg")
opt = request.QueryString("opt")
obr = request.QueryString("obr")
chave=nvg
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo


if opt="1" then
			call GravaLog (chave,"Ativado")
end if

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CONL = Server.CreateObject("ADODB.Connection") 
		ABRIRL = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&CAMINHO_log&";"
		CONL.Open ABRIRL
		
if Request.QueryString("pagina")="" then
    intpagina = 1
	unidade=request.Form("unidade")
	curso=request.Form("curso")
	etapa=request.Form("etapa")
	turma=request.Form("turma")
	mat_prin=request.Form("mat_prin")	  
	periodo=request.Form("periodo")	
	ano_de = request.form("ano_de")
	mes_de = request.form("mes_de")
	dia_de = request.form("dia_de")
	ano_ate = request.form("ano_ate")
	mes_ate = request.form("mes_ate")
	dia_ate = request.form("dia_ate")
	
	obr = unidade&"$!$"&curso&"$!$"&etapa&"$!$"&turma&"$!$"&mat_prin&"$!$"&periodo&"$!$"&ano_de&"$!$"&mes_de&"$!$"&dia_de&"$!$"&ano_ate&"$!$"&mes_ate&"$!$"&dia_ate


else
	dados_obr=split(obr,"$!$")
	unidade=dados_obr(0)
	curso=dados_obr(1)
	etapa=dados_obr(2)
	turma=dados_obr(3)
	mat_prin=dados_obr(4)
	periodo=dados_obr(5)
	ano_de = dados_obr(6)
	mes_de = dados_obr(7)
	dia_de = dados_obr(8)
	ano_ate = dados_obr(9)
	mes_ate = dados_obr(10)
	dia_ate = dados_obr(11)

end if
'response.Write(obr)
if periodo>0 then
	sql_periodo="P:"&periodo&","
else
	sql_periodo="%"	
end if	

unidade=unidade*1
if unidade = 999990 then
	if sql_periodo="" then
		sql_restr=""	
	else
		sql_restr=" AND (TB_Log_Ocorrencias.TX_Descricao like '"&sql_periodo&"%')"
	end if	
else	
	if curso = 999990 then 	
		sql_restr=" AND (TB_Log_Ocorrencias.TX_Descricao like '"&sql_periodo&"%U:"&unidade&",%')"
	else
		if isnumeric(etapa) then
			etapa=etapa*1
		end if	
		if isnumeric(mat_prin) then
			mat_prin=mat_prin*1
			if mat_prin<>999999 then
				sql_disciplina="D:"&mat_prin&","
			else
				sql_disciplina=""	
			end if	
		else
			if mat_prin<>"999999" then
				sql_disciplina="D:"&mat_prin&","
			else
				sql_disciplina=""	
			end if			
		end if	
		if isnumeric(turma) then
			turma=turma*1
		end if				
		if turma<>999990 then
			sql_turma="T:"&turma
		else
			sql_turma=""	
		end if			
		if etapa = 999990 then 	
			sql_restr=" AND (TB_Log_Ocorrencias.TX_Descricao like '"&sql_periodo&sql_disciplina&"%U:"&unidade&",C:"&curso&","&sql_turma&"%')"
		else
			sql_restr=" AND (TB_Log_Ocorrencias.TX_Descricao like '"&sql_periodo&sql_disciplina&"%U:"&unidade&",C:"&curso&",E:"&etapa&","&sql_turma&"%')"		
		end if		
	end if
end if			
data_de=mes_de&"/"&dia_de&"/"&ano_de
data_ate=mes_ate&"/"&dia_ate&"/"&ano_ate
		Set RSL= Server.CreateObject("ADODB.Recordset")
SQLL = "SELECT * FROM TB_Log_Ocorrencias WHERE ((TB_Log_Ocorrencias.CO_Sistema='WN' AND TB_Log_Ocorrencias.CO_Modulo='LN' AND TB_Log_Ocorrencias.CO_Setor='LN' AND TB_Log_Ocorrencias.CO_Funcao='LAN') OR (TB_Log_Ocorrencias.CO_Sistema='WA' AND TB_Log_Ocorrencias.CO_Modulo='PF' AND TB_Log_Ocorrencias.CO_Setor='CN' AND TB_Log_Ocorrencias.CO_Funcao='MNL'))"&sql_restr&" AND (TB_Log_Ocorrencias.DA_Ult_Acesso Between #"&data_de&"# and #"&data_ate&"#)  order by DA_Ult_Acesso,HO_ult_Acesso"		
'SQLL = "SELECT * FROM TB_Log_Ocorrencias WHERE (((TB_Log_Ocorrencias.CO_Modulo)='LN') AND ((TB_Log_Ocorrencias.CO_Funcao)='LAN') AND ((TB_Log_Ocorrencias.DA_Ult_Acesso)=#"&data_consulta&"#) AND ((TB_Log_Ocorrencias.HO_ult_Acesso)>=#"&data_consulta&" "&hora_consulta&"#)) OR (((TB_Log_Ocorrencias.CO_Modulo)='PF') AND ((TB_Log_Ocorrencias.CO_Funcao)='MNL') AND ((TB_Log_Ocorrencias.DA_Ult_Acesso)=#"&data_consulta&"#) AND ((TB_Log_Ocorrencias.HO_ult_Acesso)>=#"&data_consulta&" "&hora_consulta&"#))order by DA_Ult_Acesso,HO_ult_Acesso"	
'response.Write(SQLL)
		RSL.Open SQLL, CONL, 3, 3
		
	
 RSL.PageSize = 50
 
if Request.QueryString("pagina")="" then
      intpagina = 1
else
    if cint(Request.QueryString("pagina"))<1 then
	intpagina = 1
    else
		if cint(Request.QueryString("pagina"))>RSL.PageCount then  
	    intpagina = RSL.PageCount
        else
    	intpagina = Request.QueryString("pagina")
		end if
    end if   
 end if   

 
 call VerificaAcesso (CON,chave,nivel)
autoriza=Session("autoriza")

 call navegacao (CON,chave,nivel)
navega=Session("caminho")


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
document.all.divDisciplina.innerHTML = "<select name='mat_prin' class='select_style' id='periodo'><option value='0' selected>           </option></select>"
//recuperarEtapa()
                                                           }
                                               }
 
                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }
 
 
						 function recuperarEtapa(cTipo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=e10", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select name='turma' class='select_style' id='turma'><option value='999990' selected>           </option></select>"
document.all.divDisciplina.innerHTML = "<select name='mat_prin' class='select_style' id='periodo'><option value='0' selected>           </option></select>"
//recuperarTurma()
                                                           }
                                               }
 
                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }
 
 
						 function recuperarTurma(eTipo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t4", true);
 
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
function recuperarPeriodo(eTipo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=p1", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                       var resultado_p= oHTTPRequest.responseText;
resultado_p = resultado_p.replace(/\+/g," ")
resultado_p = unescape(resultado_p)
document.all.divPeriodo.innerHTML = resultado_p
																	   
                                                           }
                                               }
 
                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }		
function recuperarDisciplina(cTipo, eTipo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=d5", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                       var resultado_d= oHTTPRequest.responseText;
resultado_d = resultado_d.replace(/\+/g," ")
resultado_d = unescape(resultado_d)
document.all.divDisciplina.innerHTML = resultado_d
																	   
                                                           }
                                               }
 
                                               oHTTPRequest.send("c_pub=" + cTipo +"&e_pub=" + eTipo);
                                   }									   

//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900"  background="../../../../img/fundo.gif" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)%>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">		  				  				  
                  <tr>                    
            
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> </td>
  </tr>
   <tr>                   
    <td height="10"> 
      <%
	  if autoriza="no" then
	  	call mensagens(4,9700,1,0) 	  
	  else
	  	call mensagens(4,618,0,0) 
	  end if%>
    </td>
                  </tr> 
  <tr> 
    <td valign="top">

<table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
        <tr> 
          <td> 
            <%	  
		if autoriza="no" then			
		else

%>
            <table width="1000" border="0" cellspacing="0">
              <tr> 
                <td valign="top"> 
                  <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo"
>
                    <tr> 
                      <td class="tb_tit">Monitorando Notas</td>
                    </tr>
                    <tr> 
                      <td> <form name="form1" method="post" action="monitora.asp?nvg=<%response.Write(nvg)%>&opt=1" onSubmit="return checksubmit()">
                          <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0">
                            <tr>
                            	<td width="830" height="15" bgcolor="#FFFFFF"><table width="1000" border="0" cellspacing="0" cellpadding="0">
                            		<tr>
                            			<td class="tb_subtit"><div align="center">&nbsp;</div></td>
                            			</tr>
                            		<tr>
                            			<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                            				<tr>
                            					<td width="166" class="tb_subtit"><div align="center">UNIDADE </div></td>
                            					<td width="166" class="tb_subtit"><div align="center">CURSO </div></td>
                            					<td width="166" class="tb_subtit"><div align="center">ETAPA </div></td>
                            					<td width="166" class="tb_subtit"><div align="center">TURMA </div></td>
                            					<td width="166" class="tb_subtit"><div align="center">DISCIPLINA</div></td>
                            					<td width="166" class="tb_subtit"><div align="center">PER&Iacute;ODO</div></td>
                            					</tr>
                            				<tr>
                            					<td width="166"><div align="center">
                            						<select name="unidade" id="unidade" class="select_style" onChange="recuperarCurso(this.value)">
                            							<option value="999990"></option>
                            							<%		
		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT DISTINCT NU_Unidade,NO_Abr FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0

While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")
unidade=unidade*1
NU_Unidade=NU_Unidade*1
if NU_Unidade=unidade then
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

RS0.MOVENEXT
WEND
%>
                            							</select>
                            						</div></td>
                            					<td width="166"><div align="center">
                            						<div id="divCurso">
                            							<select name="curso" class="select_style" onChange="recuperarEtapa(this.value)">
                            								<option value="999990"></option>
                            								<%		
		Set RSc = Server.CreateObject("ADODB.Recordset")
		SQLc = "SELECT Distinct CO_Curso FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade
		RSc.Open SQLc, CON0
		
While not RSc.EOF
CO_Curso = RSc("CO_Curso")

		Set RS0a = Server.CreateObject("ADODB.Recordset")
		SQL0a = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
		RS0a.Open SQL0a, CON0
		
NO_Curso = RS0a("NO_Abreviado_Curso")		

if CO_Curso=curso then
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
RSc.MOVENEXT
WEND
%>
                            								</select>
                            							</div>
                            						</div></td>
                            					<td width="166"><div align="center">
                            						<div id="divEtapa">
                            							<select name="etapa" class="select_style" onChange="recuperarTurma(this.value)">
                            								<option value="999990"></option>
                            								<%		

		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"'"
		RS0b.Open SQL0b, CON0
		
		
While not RS0b.EOF
co_etapa = RS0b("CO_Etapa")


		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&curso&"' AND CO_Etapa='"&co_etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
if isnumeric(etapa) then
	etapa=etapa*1
end if
if isnumeric(co_etapa) then
	co_etapa=co_etapa*1
end if
if co_etapa=etapa then
%>
                            								<option value="<%response.Write(co_etapa)%>" selected>
                            									<%response.Write(NO_Etapa)%>
                            									</option>
                            								<%
else
%>
                            								<option value="<%response.Write(co_etapa)%>">
                            									<%response.Write(NO_Etapa)%>
                            									</option>
                            								<%

end if
RS0b.MOVENEXT
WEND
%>
                            								</select>
                            							</div>
                            						</div></td>
                            					<td width="166"><div align="center">
                            						<div id="divTurma">
                            							<select name="turma" class="select_style" onChange="gravarTurma(this.value)">
                            								<option value="999990"></option>
                            								<%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Turma where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"' AND CO_Etapa='" & etapa & "' order by CO_Turma" 
		RS3.Open SQL3, CON0						

while not RS3.EOF
co_turma= RS3("CO_Turma")
if isnumeric(turma) then
	turma=turma*1
end if
if isnumeric(co_turma) then
	co_turma=co_turma*1
end if
if co_turma=turma then
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
co_turma_check = co_turma
end if
RS3.MOVENEXT
WEND
%>
                            								</select>
                            							</div>
                            						</div></td>
                            					<td width="166"><div align="center">
                            						<div id="divDisciplina">
                            							<select name="mat_prin" class="select_style">
                            								<option value="999999" selected></option>
                            								<%
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso&"' order by NU_Ordem_Boletim "
		RS5.Open SQL5, CON0

'response.Write(SQL5)

while not RS5.EOF
co_mat_prin= RS5("CO_Materia")


		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_mat_prin &"'"
		RS7.Open SQL7, CON0
		
	if RS7.eof then
		Set RS7b = Server.CreateObject("ADODB.Recordset")
		SQL7b = "SELECT * FROM TB_Materia where CO_Materia ='"& co_mat_prin &"'"
		RS7b.Open SQL7b, CON0		
		
		no_mat_prin= RS7b("NO_Materia")
	else
		no_mat_prin= RS7("NO_Materia")
	end if
	
if mat_prin= co_mat_prin then
 mat_selected="selected" 
else
 mat_selected="" 
end if 	
	%>
                            								<option value="<%response.Write(co_mat_prin)%>" <%response.Write(mat_selected)%>>
                            									<%response.Write(no_mat_prin)%>
                            									</option>
                            								<%	
	RS5.MOVENEXT	
WEND%>
                            								</select>
                            							</div>
                            						</div></td>
                            					<td width="166"><div align="center">
                            						<div id="divPeriodo">
                            							<select name="periodo" class="select_style" id="periodo">
                            								<option value="0" selected></option>
                            								<%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")
NU_Periodo=NU_Periodo*1
periodo=periodo*1
if NU_Periodo= periodo then
 per_selected="selected" 
else
 per_selected="" 
end if 	
%>
                            								<option value="<%=NU_Periodo%>" <%response.Write(per_selected)%>>
                            									<%response.Write(NO_Periodo)%>
                            									</option>
                            								<%RS4.MOVENEXT
WEND%>
                            								</select>
                            							</div>
                            						</div></td>
                            					</tr>
                            				</table></td>
                            			</tr>
                            		<tr>
                            			<td>&nbsp;</td>
                            			</tr>
                            		<tr>
                            			<td><div align="center"> <font class="form_dado_texto">De
                            				<select name="dia_de" class="select_style">
                            					<% for d=1 to 31
									dia_de=dia_de*1
									d=d*1
                					if dia_de = d then
										dia_selected="selected"
									else
										dia_selected=""									
									end if
									if d<10 then
										dia_exibe="0"&d
									else
										dia_exibe=d										
									end if %>
                            					<option value="<%response.Write(d)%>" <%response.Write(dia_selected)%>>
                            						<%response.Write(dia_exibe)%>
                            						</option>
                            					<%next%>
                            					</select>
                            				/
                            				<select name="mes_de" id="mes_de" class="select_style">
                            					<%mes_de=mes_de*1
								if mes_de="1" or mes_de=1 then%>
                            					<option value="1" selected>janeiro</option>
                            					<% else%>
                            					<option value="1">janeiro</option>
                            					<%end if
								if mes_de="2" or mes_de=2 then%>
                            					<option value="2" selected>fevereiro</option>
                            					<% else%>
                            					<option value="2">fevereiro</option>
                            					<%end if
								if mes_de="3" or mes_de=3 then%>
                            					<option value="3" selected>mar&ccedil;o</option>
                            					<% else%>
                            					<option value="3">mar&ccedil;o</option>
                            					<%end if
								if mes_de="4" or mes_de=4 then%>
                            					<option value="4" selected>abril</option>
                            					<% else%>
                            					<option value="4">abril</option>
                            					<%end if
								if mes_de="5" or mes_de=5 then%>
                            					<option value="5" selected>maio</option>
                            					<% else%>
                            					<option value="5">maio</option>
                            					<%end if
								if mes_de="6" or mes_de=6 then%>
                            					<option value="6" selected>junho</option>
                            					<% else%>
                            					<option value="6">junho</option>
                            					<%end if
								if mes_de="7" or mes_de=7 then%>
                            					<option value="7" selected>julho</option>
                            					<% else%>
                            					<option value="7">julho</option>
                            					<%end if%>
                            					<%if mes_de="8" or mes_de=8 then%>
                            					<option value="8" selected>agosto</option>
                            					<% else%>
                            					<option value="8">agosto</option>
                            					<%end if
								if mes_de="9" or mes_de=9 then%>
                            					<option value="9" selected>setembro</option>
                            					<% else%>
                            					<option value="9">setembro</option>
                            					<%end if
								if mes_de="10" or mes_de=10 then%>
                            					<option value="10" selected>outubro</option>
                            					<% else%>
                            					<option value="10">outubro</option>
                            					<%end if
								if mes_de="11" or mes_de=11 then%>
                            					<option value="11" selected>novembro</option>
                            					<% else%>
                            					<option value="11">novembro</option>
                            					<%end if
								if mes_de="12" or mes_de=12 then%>
                            					<option value="12" selected>dezembro</option>
                            					<% else%>
                            					<option value="12">dezembro</option>
                            					<%end if%>
                            					</select>
                            				/
                            				<select name="ano_de" class="select_style" id="ano_de">
                            					<%
							Set RSa = Server.CreateObject("ADODB.Recordset")
							SQLa = "SELECT * FROM TB_Ano_Letivo order by NU_Ano_Letivo"
							RSa.Open SQLa, CON
							while not RSa.EOF 
								ano_bd=RSa("NU_Ano_Letivo")
								ano_de=ano_de*1
								ano_bd=ano_bd*1

								if ano_de=ano_bd then	%>
                            					<option value="<%response.Write(ano_bd)%>" selected><%response.Write(ano_bd)%></option>
                            					<%else%>
                            					<option value="<%response.Write(ano_bd)%>"><%response.Write(ano_bd)%></option>
                            					<%end if
							RSa.MOVENEXT
							WEND 		
				%>
                            					</select>
                            				at&eacute;
                            				<select name="dia_ate" id="dia_ate" class="select_style">
                            					<% for d=1 to 31
									dia_ate=dia_ate*1
									d=d*1
                					if dia_ate = d then
										dia_selected="selected"
									else
										dia_selected=""									
									end if
									if d<10 then
										dia_exibe="0"&d
									else
										dia_exibe=d										
									end if %>
                            					<option value="<%response.Write(d)%>" <%response.Write(dia_selected)%>>
                            						<%response.Write(dia_exibe)%>
                            						</option>
                            					<%next%>
                            					</select>
                            				/
                            				<select name="mes_ate" id="mes_ate" class="select_style">
                            					<%mes_ate=mes_ate*1
								if mes_ate="1" or mes_ate=1 then%>
                            					<option value="1" selected>janeiro</option>
                            					<% else%>
                            					<option value="1">janeiro</option>
                            					<%end if
								if mes_ate="2" or mes_ate=2 then%>
                            					<option value="2" selected>fevereiro</option>
                            					<% else%>
                            					<option value="2">fevereiro</option>
                            					<%end if
								if mes_ate="3" or mes_ate=3 then%>
                            					<option value="3" selected>mar&ccedil;o</option>
                            					<% else%>
                            					<option value="3">mar&ccedil;o</option>
                            					<%end if
								if mes_ate="4" or mes_ate=4 then%>
                            					<option value="4" selected>abril</option>
                            					<% else%>
                            					<option value="4">abril</option>
                            					<%end if
								if mes_ate="5" or mes_ate=5 then%>
                            					<option value="5" selected>maio</option>
                            					<% else%>
                            					<option value="5">maio</option>
                            					<%end if
								if mes_ate="6" or mes_ate=6 then%>
                            					<option value="6" selected>junho</option>
                            					<% else%>
                            					<option value="6">junho</option>
                            					<%end if
								if mes_ate="7" or mes_ate=7 then%>
                            					<option value="7" selected>julho</option>
                            					<% else%>
                            					<option value="7">julho</option>
                            					<%end if%>
                            					<%if mes_ate="8" or mes_ate=8 then%>
                            					<option value="8" selected>agosto</option>
                            					<% else%>
                            					<option value="8">agosto</option>
                            					<%end if
								if mes_ate="9" or mes_ate=9 then%>
                            					<option value="9" selected>setembro</option>
                            					<% else%>
                            					<option value="9">setembro</option>
                            					<%end if
								if mes_ate="10" or mes_ate=10 then%>
                            					<option value="10" selected>outubro</option>
                            					<% else%>
                            					<option value="10">outubro</option>
                            					<%end if
								if mes_ate="11" or mes_ate=11 then%>
                            					<option value="11" selected>novembro</option>
                            					<% else%>
                            					<option value="11">novembro</option>
                            					<%end if
								if mes_ate="12" or mes_ate=12 then%>
                            					<option value="12" selected>dezembro</option>
                            					<% else%>
                            					<option value="12">dezembro</option>
                            					<%end if%>
                            					</select>
                            				/
                            				<select name="ano_ate" class="select_style" id="ano_ate">
                            					<%
							Set RSb = Server.CreateObject("ADODB.Recordset")
							SQLb = "SELECT * FROM TB_Ano_Letivo order by NU_Ano_Letivo"
							RSb.Open SQLb, CON
							while not RSb.EOF 
								ano_bd=RSb("NU_Ano_Letivo")
							
								ano_ate=ano_ate*1
								ano_bd=ano_bd*1

								if ano_letivo=ano_bd then%>
                            					<option value="<%=ano_bd%>" selected><%=ano_bd%></option>
                            					<%else%>
                            					<option value="<%=ano_bd%>"><%=ano_bd%></option>
                            					<%end if
							RSb.MOVENEXT
							WEND 		
							%>
                            					</select>
                            				</font></div></td>
                            			</tr>
                            		<tr>
                            			<td><hr></td>
                            			</tr>
                            		<tr>
                            			<td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                            				<tr>
                            					<td width="33%"><div align="center"></div></td>
                            					<td width="34%"><div align="center"></div></td>
                            					<td width="33%"><div align="center"><font size="3" face="Courier New, Courier, mono">
                            						<input type="submit" name="Submit2" value="Prosseguir" class="botao_prosseguir">
                            						</font></div></td>
                            					</tr>
                            				</table></td>
                            			</tr>
                            		</table></td>
                            	</tr>
                            <tr> 
                            	<td width="328%" height="5"> </td>
                            	</tr>
                            <tr> 
                              <td> <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td width="70" class="tb_subtit"> 
                                      <div align="center"><strong>Unidade</strong></div></td>
                                    <td width="70" class="tb_subtit"> 
                                      <div align="center"><strong>Curso</strong></div></td>
                                    <td width="70" class="tb_subtit"> 
                                      <div align="center"><strong>Etapa</strong></div></td>
                                    <td width="70" class="tb_subtit"> 
                                      <div align="center"><strong>Turma</strong></div></td>
                                    <td width="125" class="tb_subtit"> 
                                      <div align="center"><strong>Per&iacute;odo</strong></div></td>
                                    <td width="125" class="tb_subtit">
<div align="center"><strong>Disciplina</strong></div></td>
                                    <td width="280" class="tb_subtit"> 
                                      <div align="center"><strong>Planilha 
                                        modificada por</strong></div></td>
                                    <td width="60" class="tb_subtit"> 
                                      <div align="center"><strong>Dia</strong></div></td>
                                    <td width="60" class="tb_subtit"> 
                                      <div align="center"><strong>Hora</strong></div></td>
                                  </tr>
                                  <% 
	IF RSL.EOF then
		intpagina=1
		sem_link=1
	%>
	<tr> 
								  <td colspan="9"><div align="center"> <font class="form_dado_texto"> Sem Movimento</font></div></td></tr>
	<%
	else
		sem_link=0
		RSL.AbsolutePage = intpagina
		intrec = 0

		While intrec<RSL.PageSize and not RSL.EOF
			ln_dt = RSL("DA_Ult_Acesso")
			ln_hr = RSL("HO_ult_Acesso")
			
	
			mnl_ln_dt = split(ln_dt,"/")
			mnl_dia = mnl_ln_dt(0)
			mnl_m = mnl_ln_dt(1)
			mnl_a = mnl_ln_dt(2)
			
			mnl_dia = mnl_dia*1
			mnl_m = mnl_m*1
			mnl_a = mnl_a*1
			
			mnl_ln_hr = split(ln_hr,":")
			mnl_h = mnl_ln_hr(0)
			mnl_mn = mnl_ln_hr(1)
			mnl_h = mnl_h*1
			mnl_mn = mnl_mn*1
		
			  if mnl_m< 10 then
				  mnl_m_wrt="0"&mnl_m
			  else
				  mnl_m_wrt = mnl_m
			  end if 
		
			  if mnl_mn< 10 then
				  mnl_mn_wrt="0"&mnl_mn
			  else
				  mnl_mn_wrt = mnl_mn
			  end if 
		
		
			ln_dt = mnl_dia&"/"&mnl_m_wrt&"/"&mnl_a
			ln_hr = mnl_h&":"&mnl_mn_wrt
			
			usr_grv = RSL("CO_Usuario")
			desc = RSL("TX_Descricao")
			
			
			mnl_desc = split(desc,",")
			
			mnl_p = mnl_desc(0)
			mnl_p_dado = split(mnl_p,":")
			mnl_p_dado_tx = mnl_p_dado(1)
			
			mnl_d = mnl_desc(1)
			mnl_d_dado = split(mnl_d,":")
			mnl_d_dado_tx = mnl_d_dado(1)
			
			mnl_u = mnl_desc(2)
			mnl_u_dado = split(mnl_u,":")
			mnl_u_dado_tx = mnl_u_dado(1)
			
			mnl_c = mnl_desc(3)
			mnl_c_dado = split(mnl_c,":")
			mnl_c_dado_tx = mnl_c_dado(1)
			
			mnl_e = mnl_desc(4)
			mnl_e_dado = split(mnl_e,":")
			mnl_e_dado_tx = mnl_e_dado(1)
			
			mnl_t = mnl_desc(5)
			mnl_t_dado = split(mnl_t,":")
			mnl_t_dado_tx = mnl_t_dado(1)
		
			Set RS0 = Server.CreateObject("ADODB.Recordset")
			SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& mnl_u_dado_tx 
			RS0.Open SQL0, CON0
				
			no_unidade = RS0("NO_Unidade")
		
			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& mnl_c_dado_tx &"'"
			RS1.Open SQL1, CON0
				
			no_curso = RS1("NO_Abreviado_Curso")
		
			Set RS3 = Server.CreateObject("ADODB.Recordset")
			SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& mnl_e_dado_tx &"' AND CO_Curso ='"& mnl_c_dado_tx &"'"
			RS3.Open SQL3, CON0
				
			if RS3.EOF THEN
				no_etapa="sem etapa"
			else
				no_etapa=RS3("NO_Etapa")
			end if
		
			Set RS7 = Server.CreateObject("ADODB.Recordset")
			SQL7 = "SELECT * FROM TB_Materia where CO_Materia_Principal='"& mnl_d_dado_tx &"'"
			RS7.Open SQL7, CON0
	
			Set RS8 = Server.CreateObject("ADODB.Recordset")
			SQL8 = "SELECT * FROM TB_Materia where CO_Materia='"& mnl_d_dado_tx &"'"
			RS8.Open SQL8, CON0
	
				no_mat= RS8("NO_Materia")
	
			if RS7.EOF Then						
				co_mat_fil = co_mat_prin
				no_mat_prin = no_mat
			else		
				co_mat_fil= RS7("CO_Materia")
			end if
			no_materia=no_mat
		
			SQL9 = "select * from TB_Usuario where CO_Usuario = " & usr_grv & ""
			set RS9 = CON.Execute (SQL9)
		
		nom_prof=RS9("NO_Usuario")
		%>
										  <tr> 
											<td width="70">
		<div align="center"> <font class="form_dado_texto"> 
												<% response.Write(no_unidade)%>
												</font> </div></td>
											<td width="70">
		<div align="center"> <font class="form_dado_texto"> 
												<% response.Write(no_curso)%>
												</font> </div></td>
											<td width="70">
		<div align="center"> <font class="form_dado_texto"> 
												<% response.Write(no_etapa)%>
												</font> </div></td>
											<td width="70">
		<div align="center"> <font class="form_dado_texto"> 
												<% response.Write(mnl_t_dado_tx)%>
												</font> </div></td>
											<td width="125"> 
											  <div align="center"> <font class="form_dado_texto"> 
												<% 
									
				Set RS4 = Server.CreateObject("ADODB.Recordset")
				SQL4 = "SELECT * FROM TB_Periodo where NU_Periodo="&mnl_p_dado_tx
				RS4.Open SQL4, CON0
		
		
		NO_Periodo= RS4("NO_Periodo")
		response.Write(NO_Periodo)%>
											  </font> </div></td>
											<td width="125"> 
											  <div align="center"> <font class="form_dado_texto"> 
												<% response.Write(no_materia)%>
												</font> </div></td>
											<td width="280"> 
											  <div align="center"> <font class="form_dado_texto"> 
												<% response.Write(nom_prof)%>
												</font> </div></td>
											<td width="60">
		<div align="center"> <font class="form_dado_texto"> 
												<% response.Write(ln_dt)%>
												</font> </div></td>
											<td width="60"> <div align="center"><font class="form_dado_texto"> 
												<% response.Write(ln_hr)%>
												</font> </div></td>
										  </tr>
										  <%
			intrec = intrec + 1
		RSL.Movenext
		Wend
	End if
	
	%>
									</table></td>
                            </tr>
                          </table>
                        </form>
                        <p>&nbsp;</p></td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table></div>
            <%end if 
%>
          </td>
        </tr>
        <tr>
          <td><div align="center">
		  <%for i=1 to RSL.PageCount
		  intpagina=intpagina*1
			  if i=intpagina then%>
			  <font class="form_dado_texto"><%response.Write(intpagina)%></font>
			  <%else%>
			   <a href="monitora.asp?pagina=<%=response.Write(i)%>&nvg=<%=nvg%>&p=<%=periodo%>&obr=<%response.Write(obr)%>&opt=pg" class="linkPaginacao"><%response.Write(i)%></a> 
			  <%end if
		  next
		  %></div></td>
        </tr>
        <tr> 
          <td class="tb_tit"><div align="center">
              <%
if sem_link=0 then
	%>&nbsp;<%		  
			    if intpagina>1 then
    %>
              <a href="monitora.asp?pagina=<%=intpagina-1%>&nvg=<%=nvg%>&obr=<%response.Write(obr)%>" class="linktres">Anterior</a> 
              <%
    end if
    if StrComp(intpagina,RSL.PageCount)<>0 then  
    %>
              <a href="monitora.asp?pagina=<%=intpagina + 1%>&nvg=<%=nvg%>&obr=<%response.Write(obr)%>" class="linktres">Próximo</a> 
              <%
    end if
else	
	%>&nbsp;<%
end if	
    RSL.close
    Set RSL = Nothing
    %>
            </div></td>
        </tr>
      </table>        
      </form></td>
  </tr>
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