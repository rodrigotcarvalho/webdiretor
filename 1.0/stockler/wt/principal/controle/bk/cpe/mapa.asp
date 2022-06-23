<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes_comuns.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes7.asp"-->
<!--#include file="../../../../inc/graficos.asp"-->
<%
opt = REQUEST.QueryString("opt")
obr = request.QueryString("obr")
nivel=4

autoriza=Session("autoriza")
Session("autoriza")=autoriza

permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("chave")
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

if opt= "vt" then
dados= split(obr, "_" )
unidade= dados(0)
curso= dados(1)
co_etapa= dados(2)
tipo_parcela = dados(3)


else
unidade = request.Form("unidade")
curso = request.Form("curso")
co_etapa = request.Form("etapa")
tipo_parcela = request.Form("tipo")
end if

ano_letivo = session("ano_letivo")
obr=unidade&"_"&curso&"_"&co_etapa&"_"&tipo_parcela&"_"&ano_letivo


m_cons="VA_Media3"


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4

		Set CON7 = Server.CreateObject("ADODB.Connection") 
		ABRIR7 = "DBQ="& CAMINHO_pf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON7.Open ABRIR7			


if unidade="nulo" then
	no_unidade="sem unidade"
	sql_unidade = ""	
	sql_curso = ""
	sql_etapa = ""	
else	
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
	no_unidade = RS0("NO_Unidade")
	sql_unidade =" AND NU_Unidade="& unidade

	if isnull(curso) or curso="" or curso="999990" then
		no_etapa="sem curso"
		sql_curso = ""
		sql_etapa = ""	
	else
		if curso=999990 then
			no_etapa="sem curso"
			sql_curso = ""	
			sql_etapa = ""				
		else
			Set RS1 = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
			RS1.Open SQL1, CON0
			sql_curso = " AND CO_Curso ='"& curso &"'"		
			no_curso = RS1("NO_Abreviado_Curso")
	
			if isnull(co_etapa) or co_etapa="" or co_etapa="999990" then
				no_etapa="sem etapa"	
				sql_etapa = ""	
			else
				if isnumeric(co_etapa) then
					if co_etapa=999990 then
						no_etapa="sem etapa"	
						sql_etapa = ""				
					else
						Set RS3 = Server.CreateObject("ADODB.Recordset")
						SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' and CO_Curso ='"& curso &"'"  
						RS3.Open SQL3, CON0
									
						if RS3.EOF THEN
							no_etapa="sem etapa"
						else
							no_etapa=RS3("NO_Etapa")
							sql_etapa = " AND CO_Etapa ='"& co_etapa &"'"	
						end if
					end if	
				else
					no_etapa="sem etapa"	
					sql_etapa = ""				
				end if	
			end if	
		end if	
	end if
end if	

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

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
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
document.all.divEtapa.innerHTML ="<select class=select_style></select>"
//recuperarEtapa()
                                                           }
                                               }

                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=e9", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
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
document.all.divPeriodo.innerHTML = "<select class=select_style></select>"
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
function recuperarDisciplina(eTipo,co_prof)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=d2", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_d= oHTTPRequest.responseText;
resultado_d = resultado_d.replace(/\+/g," ")
resultado_d = unescape(resultado_d)
document.all.divDisc.innerHTML = resultado_d
document.all.divPeriodo.innerHTML = "<select class=select_style></select>"
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo + "&pr_pub=" +co_prof);
                                   }
function recuperarPeriodo(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=p", true);

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
 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}								   
//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif"leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" background="../../../../img/fundo_interno.gif" align="center" cellspacing="0" bgcolor="#FFFFFF">
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
	call mensagens(nivel,18,0,0) 
%>
    </td>
                </tr>
                <tr> 
                  
    <td valign="top"> 
      <form name="inclusao" method="post" action="mapa.asp">
                <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
                  <tr class="tb_tit"
> 
                    <td width="653" height="15" class="tb_tit"
>Segmento </td>
                  </tr>
                  <tr> 
                    
            <td><table width="998" border="0" cellspacing="0">
                <tr> 
                  <td width="25%" class="tb_subtit"> <div align="center">UNIDADE 
                    </div></td>
                  <td width="25%" class="tb_subtit"> <div align="center">CURSO 
                    </div></td>
                  <td width="25%" class="tb_subtit"> <div align="center">ETAPA 
                  	</div></td>
                  <td width="25%" class="tb_subtit"><div align="center">TIPO DE LAN&Ccedil;AMENTO</div></td>
                  </tr>
                <tr>
                  <td background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF"> 
                    <div align="center"> 
                      <select name="unidade" class="select_style" onChange="recuperarCurso(this.value)">
<%
if unidade="nulo" then
%>
                        <option value="nulo" selected> 
                        </option>
                        <%					  
end if	
			Set RS0 = Server.CreateObject("ADODB.Recordset")
			SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
			RS0.Open SQL0, CON0
While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")
if isnumeric(unidade) then
	unidade=unidade*1
end if
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
                  <td background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF"> 
                    <div align="center"> 
                      <div id="divCurso"> 
                        <select name="curso" class="select_style" onChange="recuperarEtapa(this.value)">
<%
if unidade<>"nulo" then
%>
				
	                        <option value="999990" selected></option>						
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
	end if
%>
                        </select>
                      </div>
                    </div></td>
                  <td background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF"> 
                  	<div align="center"> 
                  		<div id="divEtapa"> 
                  			<select name="etapa" class="select_style" onChange="recuperarPeriodo(this.value)">
<%
if unidade<>"nulo" then
%>							
                  				<option value="999990" selected></option>							
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
end if	
%>
                  				</select>
                  			</div>
                  		</div></td>
                  <td width="25%"><div align="center">
                  	<select name="tipo" class="select_style">
                  		<option value="nulo" selected></option>
                  		<%Set RSP = Server.CreateObject("ADODB.Recordset")
	SQLP = "SELECT DISTINCT NO_Lancamento FROM TB_Posicao"
	RSP.Open SQLP, CON7
	
		While not RSP.EOF 		
		
		if tipo_parcela = RSP("NO_Lancamento") then
			selected = "selected"
		else
			selected = ""		
		end if	
		%>
                  		<option value="<%response.Write(RSP("NO_Lancamento"))%>"  <%response.Write(selected)%>>
                  			<%response.Write(RSP("NO_Lancamento"))%>
                  			</option>
                  		<%RSP.Movenext 
		Wend		
		%>
                  		</select>
                  	</div></td>
                  </tr>
                <tr>
                	<td colspan="4" bgcolor="#FFFFFF"><hr></td>
                	</tr>
                <tr>
                	<td height="15" colspan="4" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                		<tr>
                			<td width="33%" height="15" bgcolor="#FFFFFF"></td>
                			<td width="34%" height="15" bgcolor="#FFFFFF"></td>
                			<td width="33%" height="15" align="center" bgcolor="#FFFFFF"><font size="2" face="Arial, Helvetica, sans-serif">
                				<input name="Submit4" type="submit" class="botao_prosseguir" id="Submit4" value="Procurar">
                				</font></td>
                			</tr>
                		</table></td>
                	</tr>
              </table></td>
                  </tr>
                  <tr> 
                    
            <td align="center" valign="top">
              <%
if tipo_parcela = "nulo" then
	sql_tipo_parcela = ""
else
	sql_tipo_parcela = " AND NO_Lancamento = '"&tipo_parcela&"'"
end if

co_mat_cons=split(vetor_materias,"#!#")
tp_mat_cons=split(vetor_tipo_materia,"#!#")

		Set RSt0 = Server.CreateObject("ADODB.Recordset")
		SQLt0 = "SELECT * FROM TB_Matriculas where NU_Ano ="& ano_letivo &sql_unidade&sql_curso&sql_etapa&" And CO_Situacao='C' order by NU_Chamada"
		RSt0.Open SQLt0, CON4
		
nu_media_check = 1
while not RSt0.EOF
	nu_matricula = RSt0("CO_Matricula")
	media_numerador=0
	media_denominador=0
	
		if nu_aluno_check = 1 then
			vetor_aluno=nu_matricula
		else
			vetor_aluno=vetor_aluno&", "&nu_matricula
		end if
		nu_aluno_check=nu_aluno_check+1			
	RSt0.MOVENEXT	

wend

	Set RSM = Server.CreateObject("ADODB.Recordset")
	SQLM = "SELECT Min(Mes) as MenorMes, Max(Mes) as MaiorMes FROM TB_Posicao where CO_Matricula_Escola IN( "& vetor_aluno &")"&sql_tipo_parcela
	RSM.Open SQLM, CON7
	
	if RSM.EOF then
	else
		menor_mes=RSM("MenorMes")
		maior_mes=RSM("MaiorMes")
	end if
alunos = split(vetor_aluno, ", ")
vetor_receita_jan = 0
vetor_realizado_jan = 0
vetor_juros_jan = 0
vetor_inadimplencia_jan= 0

vetor_receita_fev = 0
vetor_realizado_fev = 0
vetor_juros_fev = 0
vetor_inadimplencia_fev= 0

vetor_receita_mar = 0
vetor_realizado_mar = 0
vetor_juros_mar = 0
vetor_inadimplencia_mar= 0

vetor_receita_abr = 0
vetor_realizado_abr = 0
vetor_juros_abr = 0
vetor_inadimplencia_abr= 0

vetor_receita_mai = 0
vetor_realizado_mai = 0
vetor_juros_mai = 0
vetor_inadimplencia_mai= 0

vetor_receita_jun = 0
vetor_realizado_jun = 0
vetor_juros_jun = 0
vetor_inadimplencia_jun= 0

vetor_receita_jul = 0
vetor_realizado_jul = 0
vetor_juros_jul = 0
vetor_inadimplencia_jul= 0

vetor_receita_ago = 0
vetor_realizado_ago = 0
vetor_juros_ago = 0
vetor_inadimplencia_ago= 0

vetor_receita_set = 0
vetor_realizado_set = 0
vetor_juros_set = 0
vetor_inadimplencia_set= 0

vetor_receita_out = 0
vetor_realizado_out = 0
vetor_juros_out = 0
vetor_inadimplencia_out= 0

vetor_receita_nov = 0
vetor_realizado_nov = 0
vetor_juros_nov = 0
vetor_inadimplencia_nov= 0

vetor_receita_dez = 0
vetor_realizado_dez = 0
vetor_juros_dez = 0
vetor_inadimplencia_dez= 0


'for al=0 to ubound(alunos)


	Set RSP = Server.CreateObject("ADODB.Recordset")
	'SQLP = "SELECT Mes, SUM(VA_Compromisso) as Receita, SUM(VA_Realizado) as Realizado FROM TB_Posicao where CO_Matricula_Escola = "& alunos(al) &"  "&sql_tipo_parcela&"GROUP BY Mes order by Mes"
	SQLP = "SELECT Mes, SUM(VA_Compromisso) as Receita, SUM(VA_Realizado) as Realizado FROM TB_Posicao where CO_Matricula_Escola IN("& vetor_aluno &")"&sql_tipo_parcela& " GROUP BY Mes order by Mes"	
	'response.Write(SQLP)
	RSP.Open SQLP, CON7
	
	if RSP.EOF then
		gera_grafico = "N"
	else	
		While not RSP.EOF 
			gera_grafico = "S"
			mes       = RSP("Mes")		
			if menor_mes <> 1 then
				menor_mes = 1
				mes_incluido=1
				while mes_incluido<=12
					receita = 0
					realizado = 0
					inclui_realizado = 0
					juros = 0
					inclui_juros = 0
					nome_apurado = Nome_Mes(mes_incluido,"ABRV", "S", 0)	
					if isnumeric(mes_incluido) then
						if mes_incluido = 1 then	
							vetor_receita_jan = vetor_receita_jan+receita
							vetor_realizado_jan = vetor_realizado_jan+inclui_realizado
							vetor_juros_jan = vetor_juros_jan+inclui_juros
							vetor_inadimplencia_jan= vetor_inadimplencia_jan+inclui_inadimplencia							
						elseif mes_incluido = 2 then	
							vetor_receita_fev = vetor_receita_fev+receita
							vetor_realizado_fev = vetor_realizado_fev+inclui_realizado
							vetor_juros_fev = vetor_juros_fev+inclui_juros
							vetor_inadimplencia_fev= vetor_inadimplencia_fev+inclui_inadimplencia						
						elseif mes_incluido = 3 then	
							vetor_receita_mar = vetor_receita_mar+receita
							vetor_realizado_mar = vetor_realizado_mar+inclui_realizado
							vetor_juros_mar = vetor_juros_mar+inclui_juros
							vetor_inadimplencia_mar= vetor_inadimplencia_mar+inclui_inadimplencia						
						elseif mes_incluido = 4 then	
							vetor_receita_abr = vetor_receita_abr+receita
							vetor_realizado_abr = vetor_realizado_abr+inclui_realizado
							vetor_juros_abr = vetor_juros_abr+inclui_juros
							vetor_inadimplencia_abr= vetor_inadimplencia_abr+inclui_inadimplencia					
						elseif mes_incluido = 5 then	
							vetor_receita_mai = vetor_receita_mai+receita
							vetor_realizado_mai = vetor_realizado_mai+inclui_realizado
							vetor_juros_mai = vetor_juros_mai+inclui_juros
							vetor_inadimplencia_mai= vetor_inadimplencia_mai+inclui_inadimplencia					
						elseif mes_incluido = 6 then	
							vetor_receita_jun = vetor_receita_jun+receita
							vetor_realizado_jun = vetor_realizado_jun+inclui_realizado
							vetor_juros_jun = vetor_juros_jun+inclui_juros
							vetor_inadimplencia_jun= vetor_inadimplencia_jun+inclui_inadimplencia					
						elseif mes_incluido = 7 then	
							vetor_receita_jul = vetor_receita_jul+receita
							vetor_realizado_jul = vetor_realizado_jul+inclui_realizado
							vetor_juros_jul = vetor_juros_jul+inclui_juros
							vetor_inadimplencia_jul= vetor_inadimplencia_jul+inclui_inadimplencia				
						elseif mes_incluido = 8 then	
							vetor_receita_ago = vetor_receita_ago+receita
							vetor_realizado_ago = vetor_realizado_ago+inclui_realizado
							vetor_juros_ago = vetor_juros_ago+inclui_juros
							vetor_inadimplencia_ago= vetor_inadimplencia_ago+inclui_inadimplencia					
						elseif mes_incluido = 9 then	
							vetor_receita_set = vetor_receita_set+receita
							vetor_realizado_set = vetor_realizado_set+inclui_realizado
							vetor_juros_set = vetor_juros_set+inclui_juros
							vetor_inadimplencia_set= vetor_inadimplencia_set+inclui_inadimplencia					
						elseif mes_incluido = 10 then	
							vetor_receita_out = vetor_receita_out+receita
							vetor_realizado_out = vetor_realizado_out+inclui_realizado
							vetor_juros_out = vetor_juros_out+inclui_juros
							vetor_inadimplencia_out= vetor_inadimplencia_out+inclui_inadimplencia					
						elseif mes_incluido = 11 then	
							vetor_receita_nov = vetor_receita_nov+receita
							vetor_realizado_nov = vetor_realizado_nov+inclui_realizado
							vetor_juros_nov = vetor_juros_nov+inclui_juros
							vetor_inadimplencia_nov= vetor_inadimplencia_nov+inclui_inadimplencia					
						elseif mes_incluido = 12 then	
							vetor_receita_dez = vetor_receita_dez+receita
							vetor_realizado_dez = vetor_realizado_dez+inclui_realizado
							vetor_juros_dez = vetor_juros_dez+inclui_juros
							vetor_inadimplencia_dez= vetor_inadimplencia_dez+inclui_inadimplencia																																													
						end if					
					end if	
		
					If Not IsArray(vetor_categorias) Then 
						vetor_categorias = Array() 
					End if	
					If InStr(Join(vetor_categorias), nome_apurado) = 0 Then
						ReDim preserve vetor_categorias(UBound(vetor_categorias)+1)	
						vetor_categorias(Ubound(vetor_categorias )) = nome_apurado	
					end if	
				mes_incluido = mes_incluido+1
				wend								
			end if	

			receita   = RSP("Receita")
			realizado = RSP("Realizado")		
			nome_apurado = Nome_Mes(mes	,"ABRV", "S", 0)				
			if isnumeric(receita) then
		
			else
				receita = 0
			end if	
			
			if isnumeric(realizado) then
		
			else
				realizado = 0
			end if		
			
			if  receita < realizado  then
				inclui_realizado = receita
			else
				inclui_realizado = realizado	
			end if	
			juros =realizado-receita
			
			if juros > 0 then
				inclui_juros = juros
			else
				inclui_juros = 0	
			end if	
		
	'if mes = 12 then		
	'		response.Write("mes: "&mes)
	'		response.Write("<BR>")		
	'		response.Write("receita: "&receita)
	'		response.Write("<BR>")		
	'		response.Write("realizado: "&realizado)	
	'		response.Write("<BR>")			
	'		response.Write("inclui_juros: "&inclui_juros)
	'		response.Write("<BR>")		
	'		response.Write("inclui_inadimplencia: "&inclui_inadimplencia)
	'		response.Write("<BR>")	
	'		response.Write("vetor_receita_dez: "&vetor_receita_dez)
	'		response.Write("<BR>")	
	'		response.Write("vetor_realizado_dez: "&vetor_realizado_dez)
	'		response.Write("<BR>")	
	'		response.Write("vetor_juros_dez: "&vetor_juros_dez)
	'		response.Write("<BR>")	
	'		response.Write("vetor_inadimplencia_dez: "&vetor_inadimplencia_dez)
	'		response.Write("<BR>")	
	'		response.Write("<BR>")										
	'end if								
			if isnumeric(mes) then
				if mes = 1 then	
					vetor_receita_jan = vetor_receita_jan+receita
					vetor_realizado_jan = vetor_realizado_jan+inclui_realizado
					vetor_juros_jan = vetor_juros_jan+inclui_juros
					vetor_inadimplencia_jan= vetor_inadimplencia_jan+inclui_inadimplencia							
				elseif mes = 2 then	
					vetor_receita_fev = vetor_receita_fev+receita
					vetor_realizado_fev = vetor_realizado_fev+inclui_realizado
					vetor_juros_fev = vetor_juros_fev+inclui_juros
					vetor_inadimplencia_fev= vetor_inadimplencia_fev+inclui_inadimplencia						
				elseif mes = 3 then	
					vetor_receita_mar = vetor_receita_mar+receita
					vetor_realizado_mar = vetor_realizado_mar+inclui_realizado
					vetor_juros_mar = vetor_juros_mar+inclui_juros
					vetor_inadimplencia_mar= vetor_inadimplencia_mar+inclui_inadimplencia						
				elseif mes = 4 then	
					vetor_receita_abr = vetor_receita_abr+receita
					vetor_realizado_abr = vetor_realizado_abr+inclui_realizado
					vetor_juros_abr = vetor_juros_abr+inclui_juros
					vetor_inadimplencia_abr= vetor_inadimplencia_abr+inclui_inadimplencia					
				elseif mes = 5 then	
					vetor_receita_mai = vetor_receita_mai+receita
					vetor_realizado_mai = vetor_realizado_mai+inclui_realizado
					vetor_juros_mai = vetor_juros_mai+inclui_juros
					vetor_inadimplencia_mai= vetor_inadimplencia_mai+inclui_inadimplencia					
				elseif mes = 6 then	
					vetor_receita_jun = vetor_receita_jun+receita
					vetor_realizado_jun = vetor_realizado_jun+inclui_realizado
					vetor_juros_jun = vetor_juros_jun+inclui_juros
					vetor_inadimplencia_jun= vetor_inadimplencia_jun+inclui_inadimplencia					
				elseif mes = 7 then	
					vetor_receita_jul = vetor_receita_jul+receita
					vetor_realizado_jul = vetor_realizado_jul+inclui_realizado
					vetor_juros_jul = vetor_juros_jul+inclui_juros
					vetor_inadimplencia_jul= vetor_inadimplencia_jul+inclui_inadimplencia				
				elseif mes = 8 then	
					vetor_receita_ago = vetor_receita_ago+receita
					vetor_realizado_ago = vetor_realizado_ago+inclui_realizado
					vetor_juros_ago = vetor_juros_ago+inclui_juros
					vetor_inadimplencia_ago= vetor_inadimplencia_ago+inclui_inadimplencia					
				elseif mes = 9 then	
					vetor_receita_set = vetor_receita_set+receita
					vetor_realizado_set = vetor_realizado_set+inclui_realizado
					vetor_juros_set = vetor_juros_set+inclui_juros
					vetor_inadimplencia_set= vetor_inadimplencia_set+inclui_inadimplencia					
				elseif mes = 10 then	
					vetor_receita_out = vetor_receita_out+receita
					vetor_realizado_out = vetor_realizado_out+inclui_realizado
					vetor_juros_out = vetor_juros_out+inclui_juros
					vetor_inadimplencia_out= vetor_inadimplencia_out+inclui_inadimplencia					
				elseif mes = 11 then	
					vetor_receita_nov = vetor_receita_nov+receita
					vetor_realizado_nov = vetor_realizado_nov+inclui_realizado
					vetor_juros_nov = vetor_juros_nov+inclui_juros
					vetor_inadimplencia_nov= vetor_inadimplencia_nov+inclui_inadimplencia					
				elseif mes = 12 then	
					vetor_receita_dez = vetor_receita_dez+receita
					vetor_realizado_dez = vetor_realizado_dez+inclui_realizado
					vetor_juros_dez = vetor_juros_dez+inclui_juros
					vetor_inadimplencia_dez= vetor_inadimplencia_dez+inclui_inadimplencia																																													
				end if					
			end if	
			'if tipo_parcela = "nulo" then
				If Not IsArray(vetor_categorias) Then 
					vetor_categorias = Array() 
				End if	
				If InStr(Join(vetor_categorias), nome_apurado) = 0 Then
					ReDim preserve vetor_categorias(UBound(vetor_categorias)+1)	
					vetor_categorias(Ubound(vetor_categorias )) = nome_apurado	
				end if	
			'end if			
		
		RSP.MOVENEXT
		WEND	
		
		if maior_mes <> 12 then
			maior_mes = 12
			mes_incluido=1	
			
			while mes_incluido<=maior_mes
			
				receita = 0
				realizado = 0
				inclui_realizado = 0
				juros = 0
				inclui_juros = 0
				nome_apurado = Nome_Mes(mes_incluido	,"ABRV", "S", 0)	
				if isnumeric(mes_incluido) then
					if mes_incluido = 1 then	
						vetor_receita_jan = vetor_receita_jan+receita
						vetor_realizado_jan = vetor_realizado_jan+inclui_realizado
						vetor_juros_jan = vetor_juros_jan+inclui_juros
						vetor_inadimplencia_jan= vetor_inadimplencia_jan+inclui_inadimplencia							
					elseif mes_incluido = 2 then	
						vetor_receita_fev = vetor_receita_fev+receita
						vetor_realizado_fev = vetor_realizado_fev+inclui_realizado
						vetor_juros_fev = vetor_juros_fev+inclui_juros
						vetor_inadimplencia_fev= vetor_inadimplencia_fev+inclui_inadimplencia						
					elseif mes_incluido = 3 then	
						vetor_receita_mar = vetor_receita_mar+receita
						vetor_realizado_mar = vetor_realizado_mar+inclui_realizado
						vetor_juros_mar = vetor_juros_mar+inclui_juros
						vetor_inadimplencia_mar= vetor_inadimplencia_mar+inclui_inadimplencia						
					elseif mes_incluido = 4 then	
						vetor_receita_abr = vetor_receita_abr+receita
						vetor_realizado_abr = vetor_realizado_abr+inclui_realizado
						vetor_juros_abr = vetor_juros_abr+inclui_juros
						vetor_inadimplencia_abr= vetor_inadimplencia_abr+inclui_inadimplencia					
					elseif mes_incluido = 5 then	
						vetor_receita_mai = vetor_receita_mai+receita
						vetor_realizado_mai = vetor_realizado_mai+inclui_realizado
						vetor_juros_mai = vetor_juros_mai+inclui_juros
						vetor_inadimplencia_mai= vetor_inadimplencia_mai+inclui_inadimplencia					
					elseif mes_incluido = 6 then	
						vetor_receita_jun = vetor_receita_jun+receita
						vetor_realizado_jun = vetor_realizado_jun+inclui_realizado
						vetor_juros_jun = vetor_juros_jun+inclui_juros
						vetor_inadimplencia_jun= vetor_inadimplencia_jun+inclui_inadimplencia					
					elseif mes_incluido = 7 then	
						vetor_receita_jul = vetor_receita_jul+receita
						vetor_realizado_jul = vetor_realizado_jul+inclui_realizado
						vetor_juros_jul = vetor_juros_jul+inclui_juros
						vetor_inadimplencia_jul= vetor_inadimplencia_jul+inclui_inadimplencia				
					elseif mes_incluido = 8 then	
						vetor_receita_ago = vetor_receita_ago+receita
						vetor_realizado_ago = vetor_realizado_ago+inclui_realizado
						vetor_juros_ago = vetor_juros_ago+inclui_juros
						vetor_inadimplencia_ago= vetor_inadimplencia_ago+inclui_inadimplencia					
					elseif mes_incluido = 9 then	
						vetor_receita_set = vetor_receita_set+receita
						vetor_realizado_set = vetor_realizado_set+inclui_realizado
						vetor_juros_set = vetor_juros_set+inclui_juros
						vetor_inadimplencia_set= vetor_inadimplencia_set+inclui_inadimplencia					
					elseif mes_incluido = 10 then	
						vetor_receita_out = vetor_receita_out+receita
						vetor_realizado_out = vetor_realizado_out+inclui_realizado
						vetor_juros_out = vetor_juros_out+inclui_juros
						vetor_inadimplencia_out= vetor_inadimplencia_out+inclui_inadimplencia					
					elseif mes_incluido = 11 then	
						vetor_receita_nov = vetor_receita_nov+receita
						vetor_realizado_nov = vetor_realizado_nov+inclui_realizado
						vetor_juros_nov = vetor_juros_nov+inclui_juros
						vetor_inadimplencia_nov= vetor_inadimplencia_nov+inclui_inadimplencia					
					elseif mes_incluido = 12 then	
						vetor_receita_dez = vetor_receita_dez+receita
						vetor_realizado_dez = vetor_realizado_dez+inclui_realizado
						vetor_juros_dez = vetor_juros_dez+inclui_juros
						vetor_inadimplencia_dez= vetor_inadimplencia_dez+inclui_inadimplencia																																													
					end if					
				end if	
		
				If Not IsArray(vetor_categorias) Then 
					vetor_categorias = Array() 
				End if	
				If InStr(Join(vetor_categorias), nome_apurado) = 0 Then
					ReDim preserve vetor_categorias(UBound(vetor_categorias)+1)	
					vetor_categorias(Ubound(vetor_categorias )) = nome_apurado	
				end if	
			mes_incluido = mes_incluido+1
			wend								
		end if			
	end if	
'			for n = 0 to ubound(vetor_categorias)
'				response.Write(vetor_categorias(n)&"<BR>")
'			next
if isnumeric(menor_mes) then
	if vetor_receita_jan < 0 then
	else
		vetor_receita = vetor_receita_jan
		vetor_realizado = vetor_realizado_jan
		vetor_juros = vetor_juros_jan
		inclui_inadimplencia = 1	
		if vetor_receita_jan<>0 then
			inadimplencia = formatnumber(1-(vetor_realizado_jan/vetor_receita_jan),4)
			
			if inadimplencia > 0 then
				inclui_inadimplencia = inadimplencia
			else
				inclui_inadimplencia = 0	
			end if	
		else
			inclui_inadimplencia = 0				
		end if			
		vetor_inadimplencia_jan=formatnumber(inclui_inadimplencia*100,2)	
		vetor_inadimplencia= vetor_inadimplencia_jan		
	end if	
	
	if 	vetor_receita_fev < 0 then
	else
		inclui_inadimplencia = 1	
		if vetor_receita_fev<>0 then
			inadimplencia = formatnumber(1-(vetor_realizado_fev/vetor_receita_fev),4)
			
			if inadimplencia > 0 then
				inclui_inadimplencia = inadimplencia
			else
				inclui_inadimplencia = 0	
			end if	
		else
			inclui_inadimplencia = 0				
		end if		
		vetor_inadimplencia_fev=formatnumber(inclui_inadimplencia*100,2)			
'		if menor_mes = 2 then
'			vetor_receita = vetor_receita_fev
'			vetor_realizado = vetor_realizado_fev
'			vetor_juros = vetor_juros_fev
'			vetor_inadimplencia= vetor_inadimplencia_fev			
'		else
			vetor_receita = vetor_receita&"#!#"&vetor_receita_fev
			vetor_realizado = vetor_realizado&"#!#"&vetor_realizado_fev
			vetor_juros = vetor_juros&"#!#"&vetor_juros_fev 
			vetor_inadimplencia= vetor_inadimplencia&"#!#"&vetor_inadimplencia_fev	
'		end if	
	end if	
		
	if 	vetor_receita_mar < 0 then
	else
		inclui_inadimplencia = 1	
		if vetor_receita_mar<>0 then
			inadimplencia = formatnumber(1-(vetor_realizado_mar/vetor_receita_mar),4)
			
			if inadimplencia > 0 then
				inclui_inadimplencia = inadimplencia
			else
				inclui_inadimplencia = 0	
			end if	
		else
			inclui_inadimplencia = 0				
		end if		
		vetor_inadimplencia_mar=formatnumber(inclui_inadimplencia*100,2)			
'		if menor_mes = 3 then
'			vetor_receita = vetor_receita_mar
'			vetor_realizado = vetor_realizado_mar
'			vetor_juros = vetor_juros_mar
'			vetor_inadimplencia= vetor_inadimplencia_mar			
'		else
			vetor_receita = vetor_receita&"#!#"&vetor_receita_mar
			vetor_realizado = vetor_realizado&"#!#"&vetor_realizado_mar
			vetor_juros = vetor_juros&"#!#"&vetor_juros_mar 
			vetor_inadimplencia= vetor_inadimplencia&"#!#"&vetor_inadimplencia_mar	
'		end if		
	end if		
						
	if 	vetor_receita_abr < 0 then
	else
		inclui_inadimplencia = 1	
		if vetor_receita_abr<>0 then
			inadimplencia = formatnumber(1-(vetor_realizado_abr/vetor_receita_abr),4)
			
			if inadimplencia > 0 then
				inclui_inadimplencia = inadimplencia
			else
				inclui_inadimplencia = 0	
			end if	
		else
			inclui_inadimplencia = 0				
		end if		
		vetor_inadimplencia_abr=formatnumber(inclui_inadimplencia*100,2)			
'		if menor_mes = 4 then
'			vetor_receita = vetor_receita_abr
'			vetor_realizado = vetor_realizado_abr
'			vetor_juros = vetor_juros_abr
'			vetor_inadimplencia= vetor_inadimplencia_abr			
'		else
			vetor_receita = vetor_receita&"#!#"&vetor_receita_abr
			vetor_realizado = vetor_realizado&"#!#"&vetor_realizado_abr
			vetor_juros = vetor_juros&"#!#"&vetor_juros_abr 
			vetor_inadimplencia= vetor_inadimplencia&"#!#"&vetor_inadimplencia_abr	
'		end if		
	end if		
			
	if 	vetor_receita_mai < 0 then
	else
		inclui_inadimplencia = 1	
		if vetor_receita_mai<>0 then
			inadimplencia = formatnumber(1-(vetor_realizado_mai/vetor_receita_mai),4)
			
			if inadimplencia > 0 then
				inclui_inadimplencia = inadimplencia
			else
				inclui_inadimplencia = 0	
			end if	
		else
			inclui_inadimplencia = 0				
		end if		
		vetor_inadimplencia_mai=formatnumber(inclui_inadimplencia*100,2)			
'		if menor_mes = 5 then
'			vetor_receita = vetor_receita_mai
'			vetor_realizado = vetor_realizado_mai
'			vetor_juros = vetor_juros_mai
'			vetor_inadimplencia= vetor_inadimplencia_mai			
'		else
			vetor_receita = vetor_receita&"#!#"&vetor_receita_mai
			vetor_realizado = vetor_realizado&"#!#"&vetor_realizado_mai
			vetor_juros = vetor_juros&"#!#"&vetor_juros_mai 
			vetor_inadimplencia= vetor_inadimplencia&"#!#"&vetor_inadimplencia_mai	
	'	end if		
	end if		
	
	if 	vetor_receita_jun < 0 then
	else
		inclui_inadimplencia = 1	
		if vetor_receita_jun<>0 then
			inadimplencia = formatnumber(1-(vetor_realizado_jun/vetor_receita_jun),4)
			
			if inadimplencia > 0 then
				inclui_inadimplencia = inadimplencia
			else
				inclui_inadimplencia = 0	
			end if	
		else
			inclui_inadimplencia = 0				
		end if		
		vetor_inadimplencia_jun=formatnumber(inclui_inadimplencia*100,2)		
'		if menor_mes = 6 then
'			vetor_receita = vetor_receita_jun
'			vetor_realizado = vetor_realizado_jun
'			vetor_juros = vetor_juros_jun
'			vetor_inadimplencia= vetor_inadimplencia_jun			
'		else
			vetor_receita = vetor_receita&"#!#"&vetor_receita_jun
			vetor_realizado = vetor_realizado&"#!#"&vetor_realizado_jun
			vetor_juros = vetor_juros&"#!#"&vetor_juros_jun 
			vetor_inadimplencia= vetor_inadimplencia&"#!#"&vetor_inadimplencia_jun	
	'	end if		
	end if	
	
	if 	vetor_receita_jul < 0 then
	else
		inclui_inadimplencia = 1	
		if vetor_receita_jul<>0 then
			inadimplencia = formatnumber(1-(vetor_realizado_jul/vetor_receita_jul),4)
			
			if inadimplencia > 0 then
				inclui_inadimplencia = inadimplencia
			else
				inclui_inadimplencia = 0	
			end if	
		else
			inclui_inadimplencia = 0				
		end if		
		vetor_inadimplencia_jul=formatnumber(inclui_inadimplencia*100,2)			
'		if menor_mes = 7 then
'			vetor_receita = vetor_receita_jul
'			vetor_realizado = vetor_realizado_jul
'			vetor_juros = vetor_juros_jul
'			vetor_inadimplencia= vetor_inadimplencia_jul			
'		else
			vetor_receita = vetor_receita&"#!#"&vetor_receita_jul
			vetor_realizado = vetor_realizado&"#!#"&vetor_realizado_jul
			vetor_juros = vetor_juros&"#!#"&vetor_juros_jul 
			vetor_inadimplencia= vetor_inadimplencia&"#!#"&vetor_inadimplencia_jul	
	'	end if		
	end if	
	
	if 	vetor_receita_ago < 0 then
	else
		inclui_inadimplencia = 1	
		if vetor_receita_ago<>0 then
			inadimplencia = formatnumber(1-(vetor_realizado_ago/vetor_receita_ago),4)
			
			if inadimplencia > 0 then
				inclui_inadimplencia = inadimplencia
			else
				inclui_inadimplencia = 0	
			end if	
		else
			inclui_inadimplencia = 0				
		end if		
		vetor_inadimplencia_ago=formatnumber(inclui_inadimplencia*100,2)		
'		if menor_mes = 8 then
'			vetor_receita = vetor_receita_ago
'			vetor_realizado = vetor_realizado_ago
'			vetor_juros = vetor_juros_ago
'			vetor_inadimplencia= vetor_inadimplencia_ago			
'		else
			vetor_receita = vetor_receita&"#!#"&vetor_receita_ago
			vetor_realizado = vetor_realizado&"#!#"&vetor_realizado_ago
			vetor_juros = vetor_juros&"#!#"&vetor_juros_ago 
			vetor_inadimplencia= vetor_inadimplencia&"#!#"&vetor_inadimplencia_ago	
'		end if		
	end if	
	
	if 	vetor_receita_set < 0 then
	else
		inclui_inadimplencia = 1	
		if vetor_receita_set<>0 then
			inadimplencia = formatnumber(1-(vetor_realizado_set/vetor_receita_set),4)
			
			if inadimplencia > 0 then
				inclui_inadimplencia = inadimplencia
			else
				inclui_inadimplencia = 0	
			end if	
		else
			inclui_inadimplencia = 0				
		end if		
		vetor_inadimplencia_set=formatnumber(inclui_inadimplencia*100,2)			
'		if menor_mes = 9 then
'			vetor_receita = vetor_receita_set
'			vetor_realizado = vetor_realizado_set
'			vetor_juros = vetor_juros_set
'			vetor_inadimplencia= vetor_inadimplencia_set			
'		else
			vetor_receita = vetor_receita&"#!#"&vetor_receita_set
			vetor_realizado = vetor_realizado&"#!#"&vetor_realizado_set
			vetor_juros = vetor_juros&"#!#"&vetor_juros_set 
			vetor_inadimplencia= vetor_inadimplencia&"#!#"&vetor_inadimplencia_set	
		'end if		
	end if	
	
	if 	vetor_receita_out < 0 then
	else
		inclui_inadimplencia = 1	
		if vetor_receita_out<>0 then
			inadimplencia = formatnumber(1-(vetor_realizado_out/vetor_receita_out),4)
			
			if inadimplencia > 0 then
				inclui_inadimplencia = inadimplencia
			else
				inclui_inadimplencia = 0	
			end if	
		else
			inclui_inadimplencia = 0				
		end if		
		vetor_inadimplencia_out=formatnumber(inclui_inadimplencia*100,2)			
'		if menor_mes < 10 then
'			vetor_receita = vetor_receita_out
'			vetor_realizado = vetor_realizado_out
'			vetor_juros = vetor_juros_out
'			vetor_inadimplencia= vetor_inadimplencia_out			
'		else
			vetor_receita = vetor_receita&"#!#"&vetor_receita_out
			vetor_realizado = vetor_realizado&"#!#"&vetor_realizado_out
			vetor_juros = vetor_juros&"#!#"&vetor_juros_out 
			vetor_inadimplencia= vetor_inadimplencia&"#!#"&vetor_inadimplencia_out	
		'end if			
	end if	
	
	if 	vetor_receita_nov < 0 then
	else
		inclui_inadimplencia = 1	
		if vetor_receita_nov<>0 then
			inadimplencia = formatnumber(1-(vetor_realizado_nov/vetor_receita_nov),4)
			
			if inadimplencia > 0 then
				inclui_inadimplencia = inadimplencia
			else
				inclui_inadimplencia = 0	
			end if	
		else
			inclui_inadimplencia = 0				
		end if		
		vetor_inadimplencia_nov=formatnumber(inclui_inadimplencia*100,2)		
'		if menor_mes = 11 then
'			vetor_receita = vetor_receita_nov
'			vetor_realizado = vetor_realizado_nov
'			vetor_juros = vetor_juros_nov
'			vetor_inadimplencia= vetor_inadimplencia_nov			
'		else
			vetor_receita = vetor_receita&"#!#"&vetor_receita_nov
			vetor_realizado = vetor_realizado&"#!#"&vetor_realizado_nov
			vetor_juros = vetor_juros&"#!#"&vetor_juros_nov 
			vetor_inadimplencia= vetor_inadimplencia&"#!#"&vetor_inadimplencia_nov	
		'end if				
	end if	
	
	if 	vetor_receita_dez < 0 then
	else
		inclui_inadimplencia = 1	
		if vetor_receita_dez<>0 then
			inadimplencia = formatnumber(1-(vetor_realizado_dez/vetor_receita_dez),4)
			
			if inadimplencia > 0 then
				inclui_inadimplencia = inadimplencia
			else
				inclui_inadimplencia = 0	
			end if	
		else
			inclui_inadimplencia = 0				
		end if		
		vetor_inadimplencia_dez=formatnumber(inclui_inadimplencia*100,2)	
'		if menor_mes = 12 then
'			vetor_receita = vetor_receita_dez
'			vetor_realizado = vetor_realizado_dez
'			vetor_juros = vetor_juros_dez
'			vetor_inadimplencia= vetor_inadimplencia_dez			
'		else
			vetor_receita = vetor_receita&"#!#"&vetor_receita_dez
			vetor_realizado = vetor_realizado&"#!#"&vetor_realizado_dez
			vetor_juros = vetor_juros&"#!#"&vetor_juros_dez 
			vetor_inadimplencia= vetor_inadimplencia&"#!#"&vetor_inadimplencia_dez	
		'end if		
	end if	
else
	session("faixas")=""
end if	


session("faixas")=vetor_receita&"#$#"&vetor_realizado&"#$#"&vetor_juros
session("legendas") = "Receita Prevista#!#Receita Realizada#!#Juros"


'session("faixas")=replace(session("faixas"),",",".")
'response.Write(session("faixas"))
'response.End()
if gera_grafico = "S" then
	for cat=0 to ubound(vetor_categorias)
		if cat=0 then
			categorias = vetor_categorias(cat)	
		else	
			categorias = categorias&"#!#"&vetor_categorias(cat)
		end if
	next	

session("categorias")=categorias
conta_registros = ubound(vetor_categorias)+1
width_table = (conta_registros*73)+120
response.Write(vetor_faixas)
%>
<table width="<%response.Write(width_table)%>" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>    
  <th width="120" class="tb_tit" scope="col">Meses</th>
<%
faixas=session("faixas")
categorias=session("categorias")

classes=split(categorias,"#!#")
'response.Write(ubound(classes))
'linhas=Split(faixas,"#$#")

mes_atual = DatePart("m", now) 


for y=0 to ubound(classes)
	'response.Write(classes(y)&"<BR>")
	nomes = classes(y)
	faixa=y
	%>
	<th width="73" class="tb_subtit" scope="col">	
	<%if y+1> mes_atual then
	 response.Write(nomes)
	else
	%>
<!--    <th scope="col" class="tb_subtit"><a href="detalhar.asp?opt=grafico&fx=<%response.Write(faixa)%>&obr=<%response.Write(obr)%>&order=d"><%response.Write(nomes)%></a></th>	-->
    <a href="../../../../relatorios/swd030.asp?obr=<%response.Write(obr)%>&opt=<%response.Write(nomes)%>" class="setor" target="_self"><%response.Write(nomes)%></a>
	
	<%end if%></th>
<%next%>
  </tr>
  <tr>
     <td width="120" align="center" class="tb_tit">Receita Prevista</td>  
 <%
valores=split(vetor_receita,"#!#")
for i=0 to ubound(valores)
	if isnumeric(valores(i)) then
		valor_1=formatnumber(valores(i),2)
	else
		valor_1=""	
	end if
%>
    <td width="73" align="right" class="form_corpo_reduzido"><%response.Write(valor_1)%></td>
<%
next
%>  
  </tr>
  <tr>
<td width="120" align="center" class="tb_tit">Receita Realizada</td>  
 <%
valores_rlzd=split(vetor_realizado,"#!#")
for r=0 to ubound(valores_rlzd)
	if isnumeric(valores_rlzd(r)) then
		valor_2=formatnumber(valores_rlzd(r),2)
	else
		valor_2=""	
	end if
%>
  	<td width="73" align="right" class="form_corpo_reduzido"><%response.Write(valor_2)%></td>
<%
next
%>  	
  	</tr>
  <tr>
<td width="120" align="center" class="tb_tit">Juros</td>  
 <%
valores_jrs=split(vetor_juros,"#!#")
for s=0 to ubound(valores_jrs)
	if isnumeric(valores_jrs(s)) then
		valor_3=formatnumber(valores_jrs(s),2)
	else
		valor_3=""		
	end if
%>
  	<td width="73" align="right" class="form_corpo_reduzido"><%response.Write(valor_3)%></td>
<%
next
%>  	
  	</tr>
  <tr>
  	<td width="120" align="center" class="tb_tit"><strong><a href="../../../../relatorios/swd030.asp?obr=<%response.Write(obr)%>&opt=p" class="modulo" target="_self">Inadimpl&ecirc;ncia (%)</a></strong></td>
 <%
valores_indmplc=split(vetor_inadimplencia,"#!#")
for t=0 to ubound(valores_indmplc)
	if isnumeric(valores_indmplc(t)) then
		valor_4=formatnumber(valores_indmplc(t),2)
	else
		valor_4=""		
	end if
%>
  	<td width="73" align="center" class="form_corpo_reduzido"><%response.Write(valor_4)%></td>
<%	
next
%>  	
  	</tr>
</table>
				<%end if%>
<DIV align="center">
<%if gera_grafico = "S" then%>
<iframe src ="iframe.asp" frameborder ="0" width="990" height="400" align="middle"> </iframe>
<%else%>
<span class="form_dado_texto">N&atilde;o existem informações disponíveis no momento!
</span>
<%end if%>
</DIV>
</td>
                  </tr>
                </table>

              </form></td>
  </tr>
  <tr>
    <td height="40"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

</body>
<%
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