<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->



<%
opt = request.QueryString("opt")

ano_letivo = session("ano_letivo")
session("ano_letivo") = ano_letivo
co_usr = session("co_user")
nivel=4

chave=session("chave")
session("chave")=chave

nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano_info=nivel&"-"&chave&"-"&ano_letivo

if opt = "err1" or opt = "ok" or opt = "vt" then
	unidade_form = session("unidade_form")
	curso_form	= session("curso_form")
	etapa_form	= session("etapa_form")
	turma_form	= session("turma_form")
	ordenacao	= session("ordenacao")
	notificados	= session("notificados")
	dia_de_form		= session("dia_de_form")
	mes_de_form		= session("mes_de_form")
	dia_ate_form	= session("dia_ate_form")
	mes_ate_form	= session("mes_ate_form")

else
	unidade_form = request.form("unidade")
	curso_form	= request.form("curso")
	etapa_form	= request.form("etapa")
	turma_form	= request.form("turma")
	ordenacao	= request.form("ordenacao")
	notificados	= request.form("notificados")
	dia_de_form		= request.form("dia_de")
	mes_de_form		= request.form("mes_de")
	dia_ate_form	= request.form("dia_ate")
	mes_ate_form	= request.form("mes_ate")
	
	session("unidade_form") = unidade_form 
	session("curso_form") = curso_form	
	session("etapa_form") = etapa_form
	session("turma_form") = turma_form
	session("ordenacao") = ordenacao
	session("notificados") = notificados
	session("dia_de_form") = dia_de_form
	session("mes_de_form") = mes_de_form
	session("dia_ate_form") = dia_ate_form
	session("mes_ate_form") = mes_ate_form
end if

ano_letivo = session("ano_letivo") 

data_de=mes_de_form&"/"&dia_de_form&"/"&ano_letivo
data_ate=mes_ate_form&"/"&dia_ate_form&"/"&ano_letivo

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

if notificados = "nulo" then
	SQL_POSICAO ="SELECT CO_Matricula_Escola, COUNT(Mes) as MESES FROM TB_Posicao where DA_Realizado is NULL AND (DA_Vencimento BETWEEN #"&data_de&"# AND #"&data_ate&"#) GROUP BY CO_Matricula_Escola order by CO_Matricula_Escola" 
Else
	SQL_POSICAO ="SELECT CO_Matricula_Escola, COUNT(Mes) as MESES FROM TB_Posicao where DA_Realizado is NULL AND (DA_Vencimento BETWEEN #"&data_de&"# AND #"&data_ate&"#) AND CO_Matricula_Escola IN (SELECT CO_Matricula_Escola from TB_Email_Enviado where CO_Email = "&notificados&" order by CO_Matricula_Escola) GROUP BY CO_Matricula_Escola" 
end if	

mensagem_eof="N&atilde;o existem alunos inadimplentes para os crit&eacute;rios informados!"

'if mes<10 then
'meswrt="0"&mes
'else
'meswrt=mes
'end if
'if min<10 then
'minwrt="0"&min
'else
'minwrt=min
'end if

data = mes &"/"& dia &"/"& ano
data_compara=data

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2
		
		Set CON6 = Server.CreateObject("ADODB.Connection") 
		ABRIR6 = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON6.Open ABRIR6			

		Set CON7 = Server.CreateObject("ADODB.Connection") 
		ABRIR7 = "DBQ="& CAMINHO_pf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON7.Open ABRIR7	

		Set CON8 = Server.CreateObject("ADODB.Connection") 
		ABRIR8 = "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON8.Open ABRIR8			

 call navegacao (CON,chave,nivel)
navega=Session("caminho")	


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
document.all.divEtapa.innerHTML ="<select class=select_style></select>"
document.all.divTurma.innerHTML = "<select class=select_style></select>"
//recuperarEtapa()
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
document.all.divTurma.innerHTML = "<select class=select_style></select>"
//recuperarTurma()
                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t", true);

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

 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
var checkflag = "false";
function check(field) {
if (checkflag == "false") {
for (i = 0; i < field.length; i++) {
field[i].checked = true;}
checkflag = "true";
return "Desmarcar Todos"; }
else {
for (i = 0; i < field.length; i++) {
field[i].checked = false; }
checkflag = "false";
return "Marcar Todos"; }
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
		call mensagens(nivel,912,2,0)	
%>
    </td>
                  </tr>
                  <%		
	elseif opt = "err1" then
%>
                  <tr> 
                    
    <td height="10"> 
      <%	
	call mensagens(nivel,910,1,0)	
%>
    </td>
                  </tr>
                  <% 	end if 

 Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS1.Open SQL1, CON0
		
 if RS1.EOF THEN%>
                  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(4,619,1,0)
%>
    </td>
                  </tr>
                  <%else%>
                  <tr> 
                    
    <td height="10"> 
      <%	call mensagens(4,1,0,0) 
	  
	 
%>
</td></tr>
<%

end if
%><tr>

            <td valign="top"> 
                
        <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td width="653" height="15" class="tb_tit">Grade de Aulas</td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td valign="top"><form name="alteracao" method="post" action="msgs.asp"><table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="250" class="tb_subtit"> 
                    <div align="center">UNIDADE 
                    </div></td>
                  <td width="250" class="tb_subtit"> 
                    <div align="center">CURSO
                    </div></td>
                  <td width="250" class="tb_subtit"> 
                    <div align="center">ETAPA 
                    </div></td>
                  <td width="250" class="tb_subtit"> 
                    <div align="center">TURMA 
                    </div></td>
                </tr>

                <tr> 
                  <td width="250"> 
                    <div align="center"> 
                      <select name="unidade" class="select_style" onChange="recuperarCurso(this.value)">
                        <option value="999990"></option>
                        <%		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0
While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")
unidade_form=unidade_form*1
NU_Unidade=NU_Unidade*1
if unidade_form = NU_Unidade then
	un_selected="selected"
else
	un_selected=""
end if
%>
                        <option value="<%response.Write(NU_Unidade)%>" <%response.Write(un_selected)%>> 
                        <%response.Write(NO_Abr)%>
                        </option>
                        <%RS0.MOVENEXT
WEND
%>
                      </select>
                    </div></td>
                  <td width="250"> 
                    <div align="center"> 
                      <div id="divCurso"> 
                      <select name="curso" class="select_style" id="curso" onChange="recuperarEtapa(this.value)">
                        <%		
	if unidade_form<>999990 then
	%>
	                        <option value="999990" selected></option>
	<%						
			Set RS0 = Server.CreateObject("ADODB.Recordset")
			SQL0 = "SELECT DISTINCT CO_Curso FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade_form
			RS0.Open SQL0, CON0
			
			
	While not RS0.EOF
	CO_Curso = RS0("CO_Curso")
	
			Set RS0a = Server.CreateObject("ADODB.Recordset")
			SQL0a = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
			RS0a.Open SQL0a, CON0
			
	NO_Curso = RS0a("NO_Abreviado_Curso")	
	
	if isnumeric(curso_form) then
		if curso_form = CO_Curso then
			cs_selected="selected"
		else
			cs_selected=""
		end if
	end if	
%>
                        <option value="<%response.Write(CO_Curso)%>" <%response.Write(cs_selected)%>> 
                        <%response.Write(NO_Curso)%>
                        </option>
                        <%
	RS0.MOVENEXT
	WEND
End if
%>
                      </select>
                      </div>
                    </div></td>
                  <td width="250"> 
                    <div align="center"> 
                      <div id="divEtapa">
                      <select name="etapa" class="select_style" onChange="recuperarTurma(this.value)">
					  
                        <%		
	if unidade_form=999990 or isnull(curso_form)  then
	
	elseif isnumeric(curso_form) then
		curso_form=curso_form*1
		if curso_form<>999990 then
		%>					  
                        <option value="999990" selected></option>
                        <%		
				Set RS0b = Server.CreateObject("ADODB.Recordset")
				SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade_form&"AND CO_Curso='"&curso_form&"'"
				RS0b.Open SQL0b, CON0
						
			While not RS0b.EOF
				CO_Etapa = RS0b("CO_Etapa")
				if isnumeric(etapa_form) then
					if etapa_form = CO_Etapa then
						et_selected="selected"
					else
						et_selected=""
					end if
				end if	
					Set RS0c = Server.CreateObject("ADODB.Recordset")
					SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&curso_form&"' AND CO_Etapa='"&CO_Etapa&"'"
					RS0c.Open SQL0c, CON0
					
			NO_Etapa = RS0c("NO_Etapa")		
			%>
									<option value="<%response.Write(CO_Etapa)%>" <%response.Write(et_selected)%>> 
									<%response.Write(NO_Etapa)%>
									</option>
									<%
			RS0b.MOVENEXT
			WEND
		end if
end if		
%>
                      </select>
                      </div>
                    </div></td>
                  <td width="250"> 
                    <div align="center"> 
                      <div id="divTurma">
                      	<select name="turma" class="select_style">
                      		
                      		<%if unidade_form=999990 or isnull(curso_form) or isnull(etapa_form) then
							Else%>
                      		<option value="999990" selected></option>
                      		<%
								Set RS3 = Server.CreateObject("ADODB.Recordset")
								SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Turma where NU_Unidade="&unidade_form&"AND CO_Curso='"&curso_form&"' AND CO_Etapa='" & etapa_form & "' order by CO_Turma" 
								RS3.Open SQL3, CON0						
						
						while not RS3.EOF
						co_turma= RS3("CO_Turma")
						
						if turma_form = co_turma then
							tm_selected="selected"
						else
							tm_selected=""
						end if						
						 %>
                      		<option value="<%=co_turma%>" <%response.Write(tm_selected)%>> 
                      			<%response.Write(co_turma)%>
                      			</option>
                      		<%
						RS3.MOVENEXT
						WEND
					end if
			
						%>
                      		
                      		</select>
                      </div>
                    </div></td>
                </tr>
                <%'end if %>
                <tr> 
                	<td width="250" height="15" bgcolor="#FFFFFF"></td>
                	<td width="250" height="15" bgcolor="#FFFFFF"></td>
                	<td width="250" height="15" bgcolor="#FFFFFF"></td>
                	<td width="250" height="15" bgcolor="#FFFFFF"></td>
                	</tr>
                <tr>
                	<td class="tb_subtit"><div align="center">ORDENA&Ccedil;&Atilde;O  </div></td>
                	<td class="tb_subtit"><div align="center">SOMENTE OS J&Aacute; NOTIFICADOS</div></td>
                	<td colspan="2" class="tb_subtit"><div align="center">PER&Iacute;ODO</div></td>
                	</tr>
                <tr>
                	<td height="15" align="center" bgcolor="#FFFFFF"><select name="ordenacao" class="select_style">
					<% if ordenacao = "UCET" then
						ucet_selected ="selected"
						nomea_selected =""
						nomer_selected =""
						qtd_selected =""																		
					elseif ordenacao = "NOMEA" then
						ucet_selected =""
						nomea_selected ="selected"
						nomer_selected =""
						qtd_selected =""						
					elseif ordenacao = "NOMER" then
						ucet_selected =""
						nomea_selected =""
						nomer_selected ="selected"
						qtd_selected =""						
					elseif ordenacao = "QTD" then
						ucet_selected =""
						nomea_selected =""
						nomer_selected =""
						qtd_selected ="selected"						
					end if%>
															
							<option value="UCET" <%response.Write(ucet_selected)%>>U/C/E/T</option>
							<option value="NOMEA" <%response.Write(nomea_selected)%>>Nome do Aluno</option>
							<option value="NOMER" <%response.Write(nomer_selected)%>>Nome do Responsável</option>
							<option value="QTD" <%response.Write(qtd_selected)%>>Quantidade de Meses inadimplente</option>																																																																										
							</select>
					</td>
                	<td height="15" align="center" bgcolor="#FFFFFF">
					<select name="notificados" class="select_style">
																		<option value="nulo" selected></option>
						<%
						Set RS3 = Server.CreateObject("ADODB.Recordset")
						SQL3 = "SELECT distinct(CO_Email) as email FROM TB_Email_Enviado"
						RS3.Open SQL3, CON7		
																								
						while not RS3.EOF
							
							co_assunto = RS3("email")
							
							Set RS0 = Server.CreateObject("ADODB.Recordset")
							SQL0 = "SELECT TX_Titulo_Assunto FROM TB_Email_Assunto where CO_Assunto = "&co_assunto
							RS0.Open SQL0, CON0																						
						
							if RS0.EOF then
							
							else
							assunto = RS0("TX_Titulo_Assunto")							
								
							%> 
								<option value="<%response.Write(co_assunto)%>"><%response.Write(assunto)%></option>
							<% end if
						%>
						
						<%RS3.MOVENEXT
						WEND%>																																																								
						</select></td>
                	<td height="15" colspan="2" align="center" bgcolor="#FFFFFF"><font class="form_dado_texto">
                		<select name="dia_de" id="dia_de" class="select_style">			
                			<% 
							 For i =1 to 31
							  	i_cod=i
								if i<10 then								
									i="0"&i
								end if
								i_cod=i_cod*1
								dia_de_form=dia_de_form*1
							if dia_de_form=i_cod then
								dia_selected="selected"
							else
								dia_selected=""							
							end if
							%>
                			<option value="<%response.Write(i_cod)%>" <%response.Write(dia_selected)%>>
                				<%response.Write(i)%>
                				</option>
                			<%
							next
							%>
                			</select>
/
<select name="mes_de" id="mes_de" class="select_style">
	<%mes_de_form=mes_de_form*1
								if mes_de_form="1" or mes_now=1 then%>
	<option value="1" selected>janeiro</option>
	<% else%>
	<option value="1">janeiro</option>
	<%end if
								if mes_de_form="2" or mes_de_form=2 then%>
	<option value="2" selected>fevereiro</option>
	<% else%>
	<option value="2">fevereiro</option>
	<%end if
								if mes_de_form="3" or mes_de_form=3 then%>
	<option value="3" selected>mar&ccedil;o</option>
	<% else%>
	<option value="3">mar&ccedil;o</option>
	<%end if
								if mes_de_form="4" or mes_de_form=4 then%>
	<option value="4" selected>abril</option>
	<% else%>
	<option value="4">abril</option>
	<%end if
								if mes_de_form="5" or mes_de_form=5 then%>
	<option value="5" selected>maio</option>
	<% else%>
	<option value="5">maio</option>
	<%end if
								if mes_de_form="6" or mes_de_form=6 then%>
	<option value="6" selected>junho</option>
	<% else%>
	<option value="6">junho</option>
	<%end if
								if mes_de_form="7" or mes_de_form=7 then%>
	<option value="7" selected>julho</option>
	<% else%>
	<option value="7">julho</option>
	<%end if%>
	<%if mes_de_form="8" or mes_de_form=8 then%>
	<option value="8" selected>agosto</option>
	<% else%>
	<option value="8">agosto</option>
	<%end if
								if mes_de_form="9" or mes_de_form=9 then%>
	<option value="9" selected>setembro</option>
	<% else%>
	<option value="9">setembro</option>
	<%end if
								if mes_de_form="10" or mes_de_form=10 then%>
	<option value="10" selected>outubro</option>
	<% else%>
	<option value="10">outubro</option>
	<%end if
								if mes_de_form="11" or mes_de_form=11 then%>
	<option value="11" selected>novembro</option>
	<% else%>
	<option value="11">novembro</option>
	<%end if
								if mes_de_form="12" or mes_de_form=12 then%>
	<option value="12" selected>dezembro</option>
	<% else%>
	<option value="12">dezembro</option>
	<%end if%>

</select>

/
<%response.write(ano_letivo)%>
at&eacute;
<select name="dia_ate" id="dia_ate" class="select_style">
	<% 
	 For i =1 to 31
		i_cod=i*1
		dia_ate_form=dia_ate_form*1
		
		if i_cod=dia_ate_form then
			selected="selected"
		else
			selected=""		
		end if	
		
		if i<10 then								
			i_char="0"&i
		else
			i_char=i
		end if

	%>
	<option value="<%response.Write(i_cod)%>" <%response.Write(selected)%>><%response.Write(i_char)%></option>
	<%next%>
</select>
/
<select name="mes_ate" id="mes_ate" class="select_style">
	<%mes_ate_form=mes_ate_form*1
								if mes_ate_form="1" or mes_ate_form=1 then%>
	<option value="1" selected>janeiro</option>
	<% else%>
	<option value="1">janeiro</option>
	<%end if
								if mes_ate_form="2" or mes_ate_form=2 then%>
	<option value="2" selected>fevereiro</option>
	<% else%>
	<option value="2">fevereiro</option>
	<%end if
								if mes_ate_form="3" or mes_ate_form=3 then%>
	<option value="3" selected>mar&ccedil;o</option>
	<% else%>
	<option value="3">mar&ccedil;o</option>
	<%end if
								if mes_ate_form="4" or mes_ate_form=4 then%>
	<option value="4" selected>abril</option>
	<% else%>
	<option value="4">abril</option>
	<%end if
								if mes_ate_form="5" or mes_ate_form=5 then%>
	<option value="5" selected>maio</option>
	<% else%>
	<option value="5">maio</option>
	<%end if
								if mes_ate_form="6" or mes_ate_form=6 then%>
	<option value="6" selected>junho</option>
	<% else%>
	<option value="6">junho</option>
	<%end if
								if mes_ate_form="7" or mes_ate_form=7 then%>
	<option value="7" selected>julho</option>
	<% else%>
	<option value="7">julho</option>
	<%end if%>
	<%if mes_ate_form="8" or mes_ate_form=8 then%>
	<option value="8" selected>agosto</option>
	<% else%>
	<option value="8">agosto</option>
	<%end if
								if mes_ate_form="9" or mes_ate_form=9 then%>
	<option value="9" selected>setembro</option>
	<% else%>
	<option value="9">setembro</option>
	<%end if
								if mes_ate_form="10" or mes_ate_form=10 then%>
	<option value="10" selected>outubro</option>
	<% else%>
	<option value="10">outubro</option>
	<%end if
								if mes_ate_form="11" or mes_ate_form=11 then%>
	<option value="11" selected>novembro</option>
	<% else%>
	<option value="11">novembro</option>
	<%end if
								if mes_ate_form="12" or mes_ate_form=12 then%>
	<option value="12" selected>dezembro</option>
	<% else%>
	<option value="12">dezembro</option>
	<%end if%>
</select>
/
<%response.write(ano_letivo)%>
					</font></td>
                	</tr>
                <tr>
                	<td height="15" colspan="4" bgcolor="#FFFFFF"><hr></td>
                	</tr>
                <tr>
                	<td height="15" colspan="4" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                		<tr>
                			<td width="33%">&nbsp;</td>
                			<td width="34%">&nbsp;</td>
                			<td width="33%" align="center"><input name="button" type="submit" class="botao_prosseguir" id="button" value="Prosseguir"></td>
                			</tr>
                		</table></td>
                	</tr>
                </table></form></td>
	</tr>
	<tr valign="top">
		<td><form name="alteracao" method="post" action="confirma.asp"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                		<tr>
                			<td class="tb_tit" colspan="11">Dados retornados</td>
                			</tr>
                		<tr>
                			<td></td>
                			</tr>
                				<tr>
                					<td width="20" class="tb_subtit"><div align="center">
                						<input type="checkbox" name="todos" class="borda" value="" onClick="this.value=check(this.form.ck_email)">
                						</div></td>
                					<td width="80" class="tb_subtit"><div align="center">Matr&iacute;cula</div></td>
                					<td width="265" class="tb_subtit"><div align="left">&nbsp; Nome do Aluno</div></td>
                					<td width="235" class="tb_subtit"><div align="left"> Nome / Email&nbsp; do Responsavel Financeiro </div></td>
                					<td width="160" class="tb_subtit"><div align="center"> Meses em aberto </div></td>
                					<td width="40" class="tb_subtit"><div align="center">&nbsp;Ult. Msg/Data </div></td>
                					<td width="50" class="tb_subtit"><div align="center">Un</div></td>
                					<td width="50" class="tb_subtit"><div align="center">Curso </div></td>
                					<td width="50" class="tb_subtit"><div align="center">Etapa</div></td>
                					<td width="50" class="tb_subtit"><div align="center">Turma</div></td>
                					</tr>
                				<tr class="<%response.write(cor)%>">
                					<td colspan="10"><hr width="1000"></td>
                					</tr>
<%	Set RSP = Server.CreateObject("ADODB.Recordset")
	SQLP =SQL_POSICAO 
	RSP.Open SQLP, CON7
	
	IF RSP.EOF THEN
	%>									
									
                				<tr class="tb_fundo_linha_par">
                					<td colspan="11" valign="top"><div align="center"><font class="style1">
                						<%
										response.Write(mensagem_eof)%>
                						</font></div></td>
                					</tr>
<%
ELSE
	Set Rs_ordena = Server.CreateObject ( "ADODB.RecordSet" )
	

	'O método Append recebe 3 parâmetros:
	'Nome do campo, Tipo, Tamanho (opcional)
	'O tipo pertence à um DataTypeEnum, e você pode conferir os tipos em
	'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/ado270/htm/mdcstdatatypeenum.asp
	'200 -> VarChar (String), 7 -> Data, 139 -> Numeric
	Rs_ordena.Fields.Append "matricula", 139, 10
	Rs_ordena.Fields.Append "nome_aluno", 200, 255
	Rs_ordena.Fields.Append "nome_responsavel", 200, 255
	Rs_ordena.Fields.Append "email_responsavel", 200, 255	
	Rs_ordena.Fields.Append "meses", 139, 10
	Rs_ordena.Fields.Append "meses_desc", 200, 255	
	Rs_ordena.Fields.Append "tipo_msg", 139, 10					
	Rs_ordena.Fields.Append "data", 7
	Rs_ordena.Fields.Append "unidade", 200, 255
	Rs_ordena.Fields.Append "curso", 200, 255
	Rs_ordena.Fields.Append "etapa", 200, 255
	Rs_ordena.Fields.Append "turma", 200, 255
	Rs_ordena.Fields.Append "in_sexo", 200, 1	

	Rs_ordena.Open

	check=0
	While not RSP.EOF	
		co_matricula = RSP("CO_Matricula_Escola")
		meses = RSP("MESES")						
		
		
		Set RSPa = Server.CreateObject("ADODB.Recordset")
		SQLPa = "SELECT CO_Email, DT_Envio FROM TB_Email_Enviado where CO_Matricula_Escola = "& co_matricula
		RSPa.Open SQLPa, CON7		
		
		if RSPa.EOF then
			tipo_msg=""
			data_envio=""
		else
			tipo_msg=RSPa("CO_Email")
			data_envio=RSPa("DT_Envio")
		end if
		
		Set RSa = Server.CreateObject("ADODB.Recordset")
		SQLa = "SELECT TB_Alunos.NO_Aluno, TB_Alunos.TP_Resp_Fin, TB_Alunos.IN_Sexo, TB_Matriculas.NU_Unidade, TB_Matriculas.CO_Curso, TB_Matriculas.CO_Etapa, TB_Matriculas.CO_Turma FROM TB_Alunos, TB_Matriculas where TB_Matriculas.CO_Matricula = TB_Alunos.CO_Matricula AND TB_Matriculas.CO_Matricula = "& co_matricula &" AND TB_Matriculas.NU_Ano = "& ano_letivo
		RSa.Open SQLa, CON2
		
		nome_aluno = RSa("NO_Aluno")
		tp_resp_fin = RSa("TP_Resp_Fin")
		in_sexo = RSa("IN_Sexo")		
		nu_unidade = RSa("NU_Unidade")
		co_curso = RSa("CO_Curso")
		co_etapa = RSa("CO_Etapa")
		co_turma = RSa("CO_Turma")


		grava_aluno="S"
		if isnull(unidade_form) then

		else
			if isnumeric(nu_unidade) then
				nu_unidade=nu_unidade*1			
			end if	
		 	if isnumeric(unidade_form) then
		
				unidade_form=unidade_form*1
				 if (unidade_form=999990 or unidade_form=nu_unidade)  then
					 
				 else
		 
					grava_aluno = "N"
				 end if			
			else

				 if (unidade_form="999990" or unidade_form=nu_unidade) then
				 
				 else
					grava_aluno = "N"
				 end if		
			end if
		end if
'response.Write(">"&curso_form)
'response.Write(grava_aluno)
		if isnull(curso_form) or curso_form="" then

		else
		 	if isnumeric(curso_form) then
				if isnumeric(co_curso) then
					co_curso=co_curso*1
				end if				
				curso_form=curso_form*1									
				 if (curso_form=999990 or curso_form=co_curso) then
				 
				 else
					grava_aluno = "N"
				 end if
			else
				 if (curso_form="999990" or curso_form=co_curso) then
				 
				 else
					grava_aluno = "N"
				 end if			
			end if
		end if
'response.Write(grava_aluno)
		if isnull(etapa_form) or etapa_form="" then

		else
		 	if isnumeric(etapa_form) then
				if isnumeric(co_etapa) then
					co_etapa=co_etapa*1
				end if			
				etapa_form=etapa_form*1			
				 if (etapa_form=999990 or etapa_form=co_etapa) then
				 
				 else
					grava_aluno = "N"
				 end if
			else
				 if (etapa_form="999990" or etapa_form=co_etapa) then
				 
				 else
					grava_aluno = "N"
				 end if			
			end if
		end if	
'response.Write(grava_aluno)						
		if isnull(turma_form) or turma_form="" then

		else
		 	if isnumeric(turma_form) then
				if isnumeric(co_turma) then
					co_turma=co_turma*1
				end if				
				turma_form=turma_form*1
				 if (turma_form=999990 or turma_form=co_turma) then
				 
				 else
					grava_aluno = "N"
				 end if
			else
				 if (turma_form="999990" or turma_form=co_turma) then
				 
				 else
					grava_aluno = "N"
				 end if			
			end if
		end if
'response.Write(grava_aluno)
		if grava_aluno="S" then		
	
			Set RSM = Server.CreateObject("ADODB.Recordset")
			SQLM ="SELECT Mes FROM TB_Posicao where DA_Realizado is NULL AND (DA_Vencimento BETWEEN #"&data_de&"# AND #"&data_ate&"#) AND CO_Matricula_Escola ="& co_matricula&" ORDER BY  Mes"
			RSM.Open SQLM, CON7	
			
			conta_mes=0
			While not RSM.EOF
				num_meses = RSM("Mes")	
				nome_meses=GeraNomesNovaVersao("MES_ABR",num_meses,variavel2,variavel3,variavel4,variavel5,CON0,outro)
				IF conta_mes=0 then
					desc_mes = nome_meses
				else
					desc_mes = desc_mes&", "&nome_meses			
				end if	
				
				conta_mes=conta_mes+1
			RSM.MOVENEXT
			WEND	
			substitui = split(desc_mes,", ")
			
			for s =0 to ubound(substitui)
				if s=ubound(substitui) and s>0 then
					desc_mes = desc_mes&" e "&substitui(s)
				elseif s=0 then	
					desc_mes = substitui(s)	
				else	
					desc_mes = desc_mes&", "&substitui(s)							
				end if			
			Next
					
			
		
		Set RSc = Server.CreateObject("ADODB.Recordset")
		SQLc = "SELECT NO_Contato,CO_CPF_PFisica FROM TB_Contatos where CO_Matricula = "& co_matricula &" AND TP_Contato = '"& tp_resp_fin&"'"
		RSc.Open SQLc, CON6		
		
		If RSc.EOF then
			nome_resp ="Nome não cadastrado para o "&tp_resp_fin
		else
			nome_resp = RSc("NO_Contato")
			cpf_resp = RSc("CO_CPF_PFisica")
			
			if cpf_resp = "" or isnull(cpf_resp) then
			
			else
				cpf_resp = replace(cpf_resp,"-","")
				cpf_resp = replace(cpf_resp,".","")				
			end if
			
			Set RSe = Server.CreateObject("ADODB.Recordset")
			SQLe = "SELECT TX_EMail_Usuario FROM TB_Usuario where CO_Usuario = "& cpf_resp
			RSe.Open SQLe, CON8		
			
			If RSe.EOF then
				email_resp ="Email não cadastrado no Webfamília"
			else	
				email_resp =RSe("TX_EMail_Usuario")
				if isnull(email_resp) or email_resp="" then
					email_resp ="Email não cadastrado no Webfamília"
				end if					
			end if					
		end if	
		
		
			Rs_ordena.AddNew
			Rs_ordena.Fields("matricula").Value = co_matricula
			Rs_ordena.Fields("nome_aluno").Value = nome_aluno
			Rs_ordena.Fields("nome_responsavel").Value = nome_resp
			Rs_ordena.Fields("email_responsavel").Value = email_resp			
			Rs_ordena.Fields("meses").Value = meses
			Rs_ordena.Fields("meses_desc").Value = desc_mes		
			Rs_ordena.Fields("tipo_msg").Value = tipo_msg					
			Rs_ordena.Fields("data").Value = data_envio
			Rs_ordena.Fields("unidade").Value = nu_unidade
			Rs_ordena.Fields("curso").Value = co_curso
			Rs_ordena.Fields("etapa").Value = co_etapa
			Rs_ordena.Fields("turma").Value = co_turma
			Rs_ordena.Fields("in_sexo").Value = in_sexo			
		end if	
	
	
		RSP.MOVENEXT
	WEND	


 if Rs_ordena.RecordCount=0 then
 %>
		<tr class="tb_fundo_linha_par">
			<td colspan="11" valign="top"><div align="center"><font class="style1">
				<%response.Write(mensagem_eof)%>
				</font></div></td>
			</tr>
<%else	
	if ordenacao="UCET" then
	 Rs_ordena.Sort = "unidade ASC, curso ASC, etapa ASC, turma ASC"	
	elseif ordenacao="NOMEA" then
	 Rs_ordena.Sort = "nome_aluno ASC"	
	elseif ordenacao="NOMER" then
	 Rs_ordena.Sort = "nome_responsavel ASC"	
	elseif ordenacao="QTD" then
	 Rs_ordena.Sort = "meses DESC"	
	end if
	 Rs_ordena.MoveFirst
								
 
	While not Rs_ordena.EOF		
		if check MOD 2 = 0 then
			cor = "tb_fundo_linha_par"
		else
			cor = "tb_fundo_linha_impar"
		end if

		no_unidade = GeraNomesNovaVersao("U",Rs_ordena.Fields("unidade").Value,variavel2,variavel3,variavel4,variavel5,CON0,outro)
		no_curso = GeraNomesNovaVersao("CA",Rs_ordena.Fields("curso").Value,variavel2,variavel3,variavel4,variavel5,CON0,outro)
		no_etapa = GeraNomesNovaVersao("E",Rs_ordena.Fields("curso").Value,Rs_ordena.Fields("etapa").Value,variavel3,variavel4,variavel5,CON0,outro)
		no_turma = 	Rs_ordena.Fields("turma").Value		
		
		meses_desc_ck_mail=replace(Rs_ordena.Fields("meses_desc").Value,", ", "-")
		
		msg_data = Rs_ordena.Fields("tipo_msg").Value&"<BR>"&Rs_ordena.Fields("data").Value 
		
		
		ck_mail = Rs_ordena.Fields("matricula").Value&"#!#"&Rs_ordena.Fields("email_responsavel").Value&"#!#"&Rs_ordena.Fields("nome_responsavel").Value&"#!#"&Rs_ordena.Fields("nome_aluno").Value&"#!#"&Rs_ordena.Fields("meses").Value&"#!#"&meses_desc_ck_mail&"#!#"&Rs_ordena.Fields("in_sexo").Value
		
%>							
									<tr valign="top" class="<%response.write(cor)%>">
                					<td width="20"><div align="center"><font class="form_dado_texto">
            							<%if InStr(Rs_ordena.Fields("email_responsavel").Value,"@") = 0 or isnull(InStr(Rs_ordena.Fields("email_responsavel").Value,"@") ) then%>
										&nbsp;
										<%else%>
                						<input name="ck_email" type="checkbox" class="borda" value="<%response.Write(ck_mail)%>">
                						<%end if%>
                						</font></div></td>
                					<td width="80"><div align="center">
                						<%response.Write(Rs_ordena.Fields("matricula").Value)%>
                						</div></td>
                					<td width="265"><div align="left"> &nbsp;
         								<%response.Write(Rs_ordena.Fields("nome_aluno").Value)%>
                						</div></td>
                					<td width="235">
                						<div align="left"><%response.Write(Rs_ordena.Fields("nome_responsavel").Value)%><BR><%response.Write(Rs_ordena.Fields("email_responsavel").Value)%></div></td>
                					<td width="160" align="center"><%response.Write(Rs_ordena.Fields("meses_desc").Value)%></td>
                					<td width="40" align="center"><%response.Write(msg_data)%></td>
                					<td width="50"><div align="center">
                						<%response.Write(no_unidade)%>
                						</div></td>
                					<td width="50"><div align="center">
                						<%response.Write(no_curso)%>
                						</div></td>
                					<td width="50"><div align="center">
                						<%response.Write(no_etapa)%>
                						</div></td>
                					<td width="50"><div align="center">
                						<%response.Write(no_turma)%>
                						</div></td>
                					</tr>
									
<%
			check = check+1
			Rs_ordena.MOVENEXT
		WEND
		%>
		<tr valign="top" class="<%response.write(cor)%>">
		<td colspan="10"><table width="100%" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td colspan="3"><hr></td></tr>
							<tr>
				<td colspan="3">
				
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td width="17%" class="tb_subtit">Enviar o Email de Cobran&ccedil;a </td>
						<td width="83%"><select name="tipo_email" class="select_style" id="tipo_email">
						
						  <%Set RSea = Server.CreateObject("ADODB.Recordset")
							SQLea = "SELECT * FROM TB_Email_Assunto"

							RSea.Open SQLea, CON0																						

							While not RSea.EOF 
								co_assunto = RSea("CO_Assunto")		
								assunto = RSea("TX_Titulo_Assunto")		
								assunto_padrao = RSea("IN_Assunto_Padrao")														
								if assunto_padrao = TRUE THEN
							%> 
								<option value="<%response.Write(co_assunto)%>" selected><%response.Write(assunto)%></option>
								<%ELSE								
							%> 
								<option value="<%response.Write(co_assunto)%>"><%response.Write(assunto)%></option>
							<%	END IF 
							RSea.Movenext
							wend%>					
						</select></td>
					</tr>
				</table></td>
				</tr>
			<tr>
				<td colspan="3"><hr></td>
				</tr>
			<tr>
				<td width="33%">&nbsp;</td>
				<td width="33%">&nbsp;</td>
				<td width="33%" align="center"><input name="button2" type="submit" class="botao_prosseguir" id="button2" value="Prosseguir"></td>
			</tr>
		</table></td>
		</tr>	
	<%			
	END IF								
END IF

%>
                				</table></form></td>
	</tr>
</table>
</td>
          </tr>
        </table>
              </td>
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
'response.redirect("../../../../inc/erro.asp")
end if
%>