<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%

nivel=4

opt=request.QueryString("opt")
orig=request.QueryString("or")

trava=session("trava")

permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("chave")
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

if opt="vt" or opt="ok" then
	curso = request.querystring("curso")
	unidade = request.querystring("unidade")
	co_etapa = request.querystring("etapa")
	turma = request.querystring("turma")
	periodo = request.querystring("periodo")
	comando = request.querystring("comando")
else
	curso = request.Form("curso")
	unidade = request.Form("unidade")
	co_etapa = request.Form("etapa")
	turma = request.Form("turma")
	periodo = request.Form("periodo")
	comando = request.Form("comando")
	obr=unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&periodo&"_"&comando	
	periodo=periodo*1
	if periodo <> 0 then
	 response.Redirect("bd2.asp?opt="&obr)
	end if	
end if

obr=unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&periodo&"_"&comando



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
		
no_curso = RS1("NO_Abreviado_Curso")


	%>
<html>
<head>
<title>Web Diretor</title>
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
                                                            alert("Esse browser n�o tem recursos para uso do Ajax");
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

//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif"  bgcolor="#FFFFFF">
  <tr>                    
            <td height="10" class="tb_caminho"><font class="style-caminho"> 
              <%
	  response.Write(navega)

%>
              </font>
	</td>
  </tr>	
  <%if opt = "ok" then%>
  <tr>                   
    <td height="10"> 
      <% call mensagens(4,602,2,0)%>
    </td>
    </tr>
	<% end if %>
<%if trava="s" then%>
  <tr>                   
    <td height="10"> 
      <% call mensagens(4,9701,0,0)%>
    </td>
    </tr> 
<%else%>					
    <tr>                   
    <td height="10"> 
      <%	call mensagens(4,600,0,0) %>
    </td>
    </tr>
<%end if%>												  				  


          <tr> 
            <td valign="top">
                <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">
                  <tr class="tb_tit"> 
                    
          <td height="15" class="tb_tit">Grade de Aulas 
            <input name="co_grupo" type="hidden" id="co_grupo" value="<% = co_grupo %>"> 
          </td>
                  </tr>
                  <tr> 
                    <td><form name="alteracao" method="post" action="altera.asp"><table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="250" class="tb_subtit"> <div align="center">UNIDADE 
                    </div></td>
                  <td width="250" class="tb_subtit"> <div align="center">CURSO 
                    </div></td>
                  <td width="250" class="tb_subtit"> <div align="center">ETAPA 
                    </div></td>
                  <td width="250" class="tb_subtit"> <div align="center">TURMA 
                    </div></td>
                </tr>
                <tr> 
                  <td> <div align="center"> 
                      <select name="unidade" class="select_style" onChange="recuperarCurso(this.value)">
                        <%		
			Set RS0 = Server.CreateObject("ADODB.Recordset")
			SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
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
                  <td> <div align="center"> 
                      <div id="divCurso"> 
                        <select name="curso" class="select_style" onChange="recuperarEtapa(this.value)">
                          <%		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT Distinct CO_Curso FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade
		RS0.Open SQL0, CON0
		
While not RS0.EOF
CO_Curso = RS0("CO_Curso")

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
RS0.MOVENEXT
WEND
%>
                        </select>
                      </div>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divEtapa"> 
                        <select name="etapa" class="select_style" onChange="recuperarTurma(this.value)">
                          <%		

		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"'"
		RS0b.Open SQL0b, CON0
		
		
While not RS0b.EOF
Etapa = RS0b("CO_Etapa")


		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&curso&"' AND CO_Etapa='"&Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
if Etapa=co_etapa then
%>
                          <option value="<%response.Write(Etapa)%>" selected> 
                          <%response.Write(NO_Etapa)%>
                          </option>
                          <%
else
%>
                          <option value="<%response.Write(Etapa)%>"> 
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
                  <td> <div align="center"> 
                      <div id="divTurma"> 
                        <select name="turma" class="select_style" onChange="MM_callJS('submitfuncao()')">
                          <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Turma where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"' AND CO_Etapa='" & co_etapa & "' order by CO_Turma" 
		RS3.Open SQL3, CON0						

while not RS3.EOF
co_turma= RS3("CO_Turma")

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
                </tr>
                <tr>
                	<td>&nbsp;</td>
                	<td>&nbsp;</td>
                	<td>&nbsp;</td>
                	<td>&nbsp;</td>
                	</tr>
                <tr class="tb_tit">
                	<td>Crit&eacute;rios</td>
                	<td>&nbsp;</td>
                	<td>&nbsp;</td>
                	<td>&nbsp;</td>
                	</tr>
                <tr>
                	<td height="18" class="tb_subtit"><div align="center">PER&Iacute;ODO</div></td>
                	<td class="tb_subtit"><div align="center">COMANDO</div></td>
                	<td class="tb_subtit">&nbsp;</td>
                	<td class="tb_subtit">&nbsp;</td>
                	</tr>
                <tr>
                	<td align="center"><select name="periodo" class="select_style">
                		<option value="0" selected></option>
                		<%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")
NU_Periodo=NU_Periodo*1
periodo = periodo*1
if NU_Periodo= periodo then
	periodo_selected = "selected"
else
	periodo_selected=""
end if

%>
                		<option value="<%response.Write(NU_Periodo)%>" <%response.Write(periodo_selected)%>>
                			<%response.Write(NO_Periodo)%>
                			</option>
                		<%RS4.MOVENEXT
WEND%>
                		</select></td>
                	<td align="center"><select name="comando" class="select_style">
<%if comando= "blq" then
	blq_selected = "selected"
	dblq_selected=""	
else
	blq_selected=""
	dblq_selected = "selected"	
end if
%>					
                		<option value="blq" <%response.Write(blq_selected)%>>Bloquear</option>
                		<option value="dblq" <%response.Write(dblq_selected)%>>Desbloquear</option>
                		</select></td>
                	<td>&nbsp;</td>
                	<td>&nbsp;</td>
                	</tr>
                <tr>
                	<td height="15" colspan="4" bgcolor="#FFFFFF"><hr></td>
                	</tr>
                <tr>
                	<td height="15" colspan="4" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                		<tr>
                			<td width="34%">&nbsp;</td>
                			<td width="33%">&nbsp;</td>
                			<td width="33%" align="center"><input type="submit" name="Submit2" value="Prosseguir" class="botao_prosseguir"></td>
                			</tr>
                		</table></td>
                	</tr>
                	</table>
                    </form></td>
                  </tr>
                  <tr bgcolor="#FFFFFF"> 
                    <td bgcolor="#FFFFFF">&nbsp;</td>
                  </tr>
                  <tr> 
                    <td><table width="100%" border="0" cellspacing="0">
              <tr> 
                <td width="250" class="tb_subtit"> <div align="center"><strong>DISCIPLINA</strong></div></td>
                <td width="450" class="tb_subtit"> 
                  <div align="center"><strong>PROFESSOR</strong></div></td>
                  <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
		RS4.Open SQL4, CON0
		conta_periodo=0
		while not RS4.eof 
			periodo= RS4("NU_Periodo")
			sigla_periodo= RS4("SG_Periodo") 
			
			if conta_periodo=0 then
				vetor_periodo=periodo		
			else
				vetor_periodo=vetor_periodo&"#!#"&periodo			
			end if	
			
	if periodo = "1" then
		periodo_SQL="ST_Per_1"
	elseif periodo = "2" then
		periodo_SQL="ST_Per_2"
	elseif periodo = "3" then
		periodo_SQL="ST_Per_3"
	elseif periodo= "4" then
		periodo_SQL="ST_Per_4"
	elseif periodo = "5" then
		periodo_SQL="ST_Per_5"
	elseif periodo = "6" then
		periodo_SQL="ST_Per_6"
	end if	

	Set RS_b = Server.CreateObject("ADODB.Recordset")
	SQL_b = "SELECT "&periodo_SQL&" FROM TB_Da_Aula where NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
	RS_b.Open SQL_b, CON2		
	
	bloqueado=0
	desbloqueado=0	
	while not RS_b.EOF
	periodo_bloqueado=RS_b(periodo_SQL)
	
		if periodo_bloqueado="x" then	
			bloqueado=bloqueado+1
		else
			desbloqueado=desbloqueado+1	
		end if		
	RS_b.MOVENEXT
	WEND
	
	if desbloqueado=0 then	
		opt="fdblq"	
	else
		opt="fblq"
	end if		
	 %>
                  <td class="tb_subtit"> <div align="center"> 
                      <%response.Write("<a href=bd.asp?opt="&opt&"&nt="&nota&"&u="&unidade&"&c="&curso&"&e="&co_etapa&"&t="&turma&"&p="&periodo&">"&ucase(sigla_periodo)&"</a>")%>
                    </div></td>                  
   		<%
		conta_periodo=conta_periodo+1		
		RS4.MOVENEXT
		WEND
		width=formatnumber(300/conta_periodo,0)
		%>               
<!--                <td width="50" class="tb_subtit"> <div align="center"><strong>B1</strong></div></td>
                <td width="50" class="tb_subtit"> <div align="center"><strong>B2</strong></div></td>
                <td width="50" class="tb_subtit"> <div align="center"><strong>B3</strong></div></td>
                <td width="50" class="tb_subtit"> <div align="center"><strong>B4</strong></div></td>
                <td width="50" class="tb_subtit"> <div align="center"><strong>REC</strong></div></td>
-->              </tr>
              <%
				



  				
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
		RS5.Open SQL5, CON0
check=2
while not RS5.EOF
co_mat_prin= RS5("CO_Materia")

 if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if%>
              <tr> 
                <td width="250" class="<%=cor%>"> <div align="center"><font class="form_dado_texto"> 
                    <%

		Set RS8 = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT * FROM TB_Materia where CO_Materia='"& co_mat_prin &"'"
		RS8.Open SQL8, CON0

		no_mat= RS8("NO_Materia")	
		co_mat_fil= RS8("CO_Materia")


		response.Write(no_mat)
%>
                    </font></div></td>
                <td width="450" class="<%=cor%>"> 
                  <div align="center"><font class="form_dado_texto"> 
                    <%

		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Da_Aula where CO_Materia_Principal='"& co_mat_fil &"'AND NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		RS2.Open SQL2, CON2
		
if RS2.EOF then
		response.Write("n�o cadastrado<br>")		

else		
while not RS2.EOF					  

co_professor = RS2("CO_Professor")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Professor where CO_Professor="& co_professor
		RS1.Open SQL1, CON1
			
if RS1.EOF then		
				

response.Write("nome em branco<br>")
else
no_prof= RS1("NO_Professor")
ativo = RS1("IN_Ativo_Escola")
	if ativo = "True" then
	Response.Write("<font size=1 face=Verdana, Arial, Helvetica, sans-serif><a class=ativos href=../../man/professores/altera.asp?ori=01&opt=ln&cod_cons="&co_professor&"&nvg=WA-PF-MA-APR target=_parent>"&no_prof&"</a></font><br>")
	else
	Response.Write("<font size=1 face=Verdana, Arial, Helvetica, sans-serif><a class=inativos href=../../man/professores/altera.asp?ori=01&opt=ln&cod_cons="&co_professor&"&nvg=WA-PF-MA-APR target=_parent>"&no_prof&"</a></font><br>")
	end if
end if
		
RS2.MOVENEXT
WEND
end if
%>
                    </font></div></td>
<%
periodo_colunas=split(vetor_periodo,"#!#")
for per=0 to ubound(periodo_colunas)

	if periodo_colunas(per) = "1" then
		periodo_SQL="ST_Per_1"
	elseif periodo_colunas(per) = "2" then
		periodo_SQL="ST_Per_2"
	elseif periodo_colunas(per) = "3" then
		periodo_SQL="ST_Per_3"
	elseif periodo_colunas(per) = "4" then
		periodo_SQL="ST_Per_4"
	elseif periodo_colunas(per) = "5" then
		periodo_SQL="ST_Per_5"
	elseif periodo_colunas(per) = "6" then
		periodo_SQL="ST_Per_6"
	end if	

%>
                <td width="<%response.Write(width)%>" class="<%=cor%>"> 
                  <%
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Da_Aula where CO_Materia_Principal='"& co_mat_prin &"'AND NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		RS2.Open SQL2, CON2
		
	if RS2.EOF then
			
	
	else		
		while not RS2.EOF					  

			co_professor = RS2("CO_Professor")					

			Set RS4a = Server.CreateObject("ADODB.Recordset")
			SQL4a = "SELECT * FROM TB_Da_Aula where CO_Professor="& co_professor &"AND CO_Materia_Principal='"& co_mat_prin &"'AND NU_Unidade="& unidade &"AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
			RS4a.Open SQL4a, CON2
	
	
	
			if RS4a.EOF then
	
			else	

					p = RS4a(periodo_SQL)
					nota = RS4a("TP_Nota")
					if p = "x" then
						if trava="s" then%>
						  <div align="center"><img src="../../../../img/s.gif" width="8" height="8" border="0"></div>
					  <%else%>
						  <div align="center"><a href="bd.asp?opt=dblq&nt=<%response.Write(nota)%>&u=<% response.Write(unidade) %>&c=<% response.Write(curso) %>&e=<% response.Write(co_etapa) %>&t=<%response.Write(turma)%>&d=<%response.Write(co_mat_prin)%>&pr=<%response.Write(co_professor)%>&p=<% response.Write(periodo_colunas(per)) %>"><img src="../../../../img/s.gif" width="8" height="8" border="0"></a></div>
					  <%end if%>
					  <%
					else
						if trava="s" then%>
							  <div align="center"><img src="../../../../img/n.gif" width="8" height="8" border="0"></div>
						  <%else%>
							  <div align="center"><a href="bd.asp?opt=blq&nt=<%response.Write(nota)%>&u=<% response.Write(unidade) %>&c=<% response.Write(curso) %>&e=<% response.Write(co_etapa) %>&t=<%response.Write(turma)%>&d=<%response.Write(co_mat_prin)%>&pr=<%response.Write(co_professor)%>&p=<% response.Write(periodo_colunas(per)) %>"><img src="../../../../img/n.gif" width="8" height="8" border="0"></a></div>
						  <%end if%>
						  <%
						end if
					end if
		RS2.MOVENEXT
		WEND
	end if	
%>
                </td>
<%next%>                
              </tr>
              <%
check=check+1
RS5.MOVENEXT
WEND
%>
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
<%call GravaLog (chave,"0")%>
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