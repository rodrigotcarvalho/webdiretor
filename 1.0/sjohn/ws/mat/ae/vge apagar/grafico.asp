<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->


<%
opt = request.QueryString("opt")
situacao_aluno= request.form("situacao_aluno")
ano_letivo = session("ano_letivo")
co_usr = session("co_user")
nivel=4
nvg = session("nvg")
session("nvg")=nvg
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano_info=nivel&"-"&chave&"-"&ano_letivo

'if opt="un" then
'	classe_tit_un="tb_tit"
'	classe_corpo_un="tb_fundo_linha_par"
'	classe_tit_cs="transparencia_tit"
'	classe_corpo_cs="transparencia_corpo"
'	classe_tit_et="transparencia_tit"
'	classe_corpo_et="transparencia_corpo"
'elseif opt="cs" then
'	unidade= request.form("unidade")
'	classe_tit_un="transparencia_tit"
'	classe_corpo_un="transparencia_corpo"
'	classe_tit_cs="tb_tit"
'	classe_corpo_cs="tb_fundo_linha_par"
'	classe_tit_et="transparencia_tit"
'	classe_corpo_et="transparencia_corpo"
'elseif opt="et" then
'	unidade= request.form("unidade")
'	curso= request.form("curso")
'	classe_tit_un="transparencia_tit"
'	classe_corpo_un="transparencia_corpo"
'	classe_tit_cs="transparencia_tit"
'	classe_corpo_cs="transparencia_corpo"
'	classe_tit_et="tb_tit"
'	classe_corpo_et="tb_fundo_linha_par"
'else
'	classe_tit_un="tb_tit"
'	classe_corpo_un="tb_fundo_linha_par"
'	classe_tit_cs="tb_tit"
'	classe_corpo_cs="tb_fundo_linha_par"
'	classe_tit_et="tb_tit"
'	classe_corpo_et="tb_fundo_linha_par"
'end if

	tema=request.form("tema")
	qqa=request.form("qqa")
	qpr=request.form("qpr")
	qde=request.form("qde")
	unidade=request.form("unidade")

	qqa_select=""
	qpr_select=""
	qde_select=""
	gen_checked=""	
	idade_checked=""		
	ano_checked=""		
	mes_checked=""
	estc_checked=""			
	respf_checked=""	
	respp_checked=""	
	unidd_checked=""
	curso_checked=""				
	etapa_checked=""	

	if tema="qqa" then					
	'if qqa="gen" or qqa="idade" or qqa="ano" or qqa="mes" then
		onload=";AlternarConteudo('qqa')"
		qqa_select="SELECTED"
		
		if qqa="gen" then
			gen_checked="CHECKED"
			tipo_grafico="pizza"
		elseif qqa="idade" then
			idade_checked="CHECKED"	
			tipo_grafico="barra"				
		elseif qqa="ano" then
			ano_checked="CHECKED"
			tipo_grafico="barra"							
		elseif qqa="mes" then
			mes_checked="CHECKED"
			tipo_grafico="barra"					
		end if		
				
	elseif tema="qpr" then	
	'elseif qpr="estc" or qpr="respf" or qpr="respp" then
		onload=";AlternarConteudo('qpr')"
		qpr_select="SELECTED"
				
		if qpr="estc" then
			estc_checked="CHECKED"
			tipo_grafico="pizza"			
		elseif qpr="respf" then
			respf_checked="CHECKED"	
			tipo_grafico="barra"					
		elseif qpr="respp" then
			respp_checked="CHECKED"	
			tipo_grafico="barra"							
		end if				
	elseif tema="qde" then	
	'elseif qde="unidd" or qde="curso" or qde="etapa" then
		onload=";AlternarConteudo('qde')"
		qde_select="SELECTED"
				
		if qde="unidd" then
			unidd_checked="CHECKED"
			tipo_grafico="pizza"			
		elseif qde="curso" then
			curso_checked="CHECKED"
			tipo_grafico="torre"					
		elseif qde="etapa" then
			etapa_checked="CHECKED"	
			tipo_grafico="torre"						
		end if				
	end if
	
	
	

	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
		
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0

	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1	

	Set RS = Server.CreateObject("ADODB.Recordset")
	CONEXAO = "Select * from TB_Situacao_Aluno order by CO_Situacao"
	Set RS = CON0.Execute(CONEXAO)
	
	conta=0
	While not RS.EOF
	co_situac_aluno=RS("CO_Situacao")
	nom_situac_aluno=RS("TX_Descricao_Situacao")
		if co_situac_aluno = "C" or co_situac_aluno = "P" or co_situac_aluno = "T" or co_situac_aluno = "E" THEN 
			conta=conta+1
			if conta=1 then
				vetor_co_situac=co_situac_aluno	
				vetor_nom_situac=nom_situac_aluno
			else
				vetor_co_situac=vetor_co_situac&"#!#"&co_situac_aluno
				vetor_nom_situac=vetor_nom_situac&"#!#"&nom_situac_aluno
			end if
		end if	
	RS.MOVENEXT
	WEND
co_situacoes=split(vetor_co_situac,"#!#")		
situacoes=split(vetor_nom_situac,"#!#")	
	
call navegacao (CON,nvg,nivel)
navega=Session("caminho")	

	if opt="un" then
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0
		nu_unidades_check = 1
		nu_alunos_check = 1
		While not RS0.EOF
		qtd_aluno=0
		unidade_tabela = RS0("NU_Unidade")
		NO_Abr = RS0("NO_Abr")	
			Set RSt0 = Server.CreateObject("ADODB.Recordset")
			SQLt0 ="SELECT Count(CO_Matricula) as Qtd_Alunos FROM TB_Matriculas where NU_Ano ="& ano_letivo&" AND NU_Unidade ="& unidade_tabela&" AND CO_Situacao='"&situacao_aluno&"'"
			RSt0.Open SQLt0, CON1	
			qtd_aluno=RSt0("Qtd_Alunos")

				if nu_alunos_check = 1 then
					vetor_faixas=qtd_aluno
				else
					vetor_faixas=vetor_faixas&"#!#"&qtd_aluno
				end if
			nu_alunos_check=nu_alunos_check+1	

				if nu_unidades_check = 1 then
					vetor_categorias=NO_Abr
				else
					vetor_categorias=vetor_categorias&"#!#"&NO_Abr
				end if
			nu_unidades_check=nu_unidades_check+1	
			
		RS0.MOVENEXT
		WEND
	elseif opt="cs" then
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT DISTINCT CO_Curso FROM TB_Unidade_Possui_Etapas WHERE NU_Unidade= "&unidade&" order by CO_Curso"
		RS0.Open SQL0, CON0
		nu_curso_check = 1
		nu_alunos_check = 1
		While not RS0.EOF
		qtd_aluno=0
		curso_tabela = RS0("CO_Curso")
		
			Set RS0a = Server.CreateObject("ADODB.Recordset")
			SQL0a = "SELECT * FROM TB_Curso WHERE CO_Curso= '"&curso_tabela&"'"
			RS0a.Open SQL0a, CON0		
			NO_Abr = RS0a("NO_Abreviado_Curso")	
			
			Set RSt0 = Server.CreateObject("ADODB.Recordset")		
			SQLt0 ="SELECT Count(CO_Matricula) as Qtd_Alunos FROM TB_Matriculas where NU_Ano ="& ano_letivo&" AND NU_Unidade ="& unidade&" AND CO_Curso ='"& curso_tabela &"' AND CO_Situacao='"&situacao_aluno&"'"
			RSt0.Open SQLt0, CON1	
			qtd_aluno=RSt0("Qtd_Alunos")

				if nu_alunos_check = 1 then
					vetor_faixas=qtd_aluno
				else
					vetor_faixas=vetor_faixas&"#!#"&qtd_aluno
				end if
			nu_alunos_check=nu_alunos_check+1	

				if nu_curso_check = 1 then
					vetor_categorias=NO_Abr
				else
					vetor_categorias=vetor_categorias&"#!#"&NO_Abr
				end if
			nu_curso_check=nu_curso_check+1	
			
		RS0.MOVENEXT
		WEND
	
	elseif opt="et" then
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas WHERE NU_Unidade= "&unidade&" AND CO_Curso ='"& curso &"' order by CO_Etapa"
		RS0.Open SQL0, CON0
		nu_etapa_check = 1
		nu_alunos_check = 1
		While not RS0.EOF
		qtd_aluno=0
		co_etapa_tabela = RS0("CO_Etapa")
		
			Set RS0a = Server.CreateObject("ADODB.Recordset")
			SQL0a = "SELECT * FROM TB_Etapa WHERE CO_Curso= '"&curso&"' AND CO_Etapa='"&co_etapa_tabela&"'"
			RS0a.Open SQL0a, CON0		
			NO_Abr = RS0a("NO_Etapa")	
			
			Set RSt0 = Server.CreateObject("ADODB.Recordset")		
			SQLt0 ="SELECT Count(CO_Matricula) as Qtd_Alunos FROM TB_Matriculas where NU_Ano ="& ano_letivo&" AND NU_Unidade ="& unidade&" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa_tabela &"' AND CO_Situacao='"&situacao_aluno&"'"
			RSt0.Open SQLt0, CON1	
			qtd_aluno=RSt0("Qtd_Alunos")

				if nu_alunos_check = 1 then
					vetor_faixas=qtd_aluno
				else
					vetor_faixas=vetor_faixas&"#!#"&qtd_aluno
				end if
			nu_alunos_check=nu_alunos_check+1	

				if nu_etapa_check = 1 then
					vetor_categorias=NO_Abr
				else
					vetor_categorias=vetor_categorias&"#!#"&NO_Abr
				end if
			nu_etapa_check=nu_etapa_check+1	
			
		RS0.MOVENEXT
		WEND	
	
	
	
	
		sql="SELECT Count(CO_Matricula) as Qtd_Alunos FROM TB_Matriculas where NU_Unidade ="& unidade&" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa_tabela &"' AND CO_Situacao='"&situacao_aluno&"'"
	end if

session("faixas")=vetor_faixas
session("categorias")=vetor_categorias
'response.Write(	vetor_faixas&"_"&vetor_categorias)	
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
 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
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

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=c2", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
//recuperarEtapa()
                                                           }
                                               }

                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }								   								   

 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
function checkSubmit()
{
  if (document.grafico.tema.value == "nulo")
  {    alert("Por favor selecione um Tema Estat�stico!")
    document.grafico.tema.focus()
    return false
  }
  return true
}

function stAba(conteudo)
	{
		this.conteudo = conteudo;
	}
 
	var arAbas = new Array();
	arAbas[0] = new stAba('qaa');
	arAbas[1] = new stAba('qpr');
	arAbas[2] = new stAba('qde');	
 
	function AlternarConteudo(conteudo)
	{
		for (i=0;i<arAbas.length;i++)
		{
			c = document.getElementById(arAbas[i].conteudo)
			c.style.display = 'none';
		}
		c = document.getElementById(conteudo)
		c.style.display = '';
	}
</script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../../img/menu_r1_c2_f3.gif','../../../../img/menu_r1_c2_f2.gif','../../../../img/menu_r1_c2_f4.gif','../../../../img/menu_r1_c4_f3.gif','../../../../img/menu_r1_c4_f2.gif','../../../../img/menu_r1_c4_f4.gif','../../../../img/menu_r1_c6_f3.gif','../../../../img/menu_r1_c6_f2.gif','../../../../img/menu_r1_c6_f4.gif','../../../../img/menu_r1_c8_f3.gif','../../../../img/menu_r1_c8_f2.gif','../../../../img/menu_r1_c8_f4.gif','../../../../img/menu_direita_r2_c1_f3.gif','../../../../img/menu_direita_r2_c1_f2.gif','../../../../img/menu_direita_r2_c1_f4.gif','../../../../img/menu_direita_r4_c1_f3.gif','../../../../img/menu_direita_r4_c1_f2.gif','../../../../img/menu_direita_r4_c1_f4.gif','../../../../img/menu_direita_r6_c1_f3.gif','../../../../img/menu_direita_r6_c1_f2.gif','../../../../img/menu_direita_r6_c1_f4.gif')<%response.Write(onload)%>">
<% call cabecalho (nivel)
	  %>
<table width="1003" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
                    
    <td width="1001" height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
	  </td>
	  </tr>
<tr>               
    <td height="10"> 
      <%	call mensagens(4,708,0,0) 
%>
</td></tr>
<tr>

       <td width="1000" valign="top">                
        <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">
          <tr> 
            <td width="653" class="<%response.Write(classe_corpo_un)%>"><form name="grafico" method="post" action="grafico.asp" onSubmit="return checkSubmit()">
            <table width="1000" border="0" cellspacing="0" cellpadding="0">
              <tr class="tb_subtit">
                <td class="tb_tit">Preencha os campos abaixo</td>
                <td class="tb_tit">&nbsp;</td>
                <td class="tb_tit">&nbsp;</td>
                <td class="tb_tit">&nbsp;</td>
              </tr>
              <tr class="tb_subtit">
                <td class="tb_subtit"><div align="right">Tema Estat&iacute;stico</div></td>
                <td class="tb_subtit">&nbsp;
                  <select name="tema" class="select_style" id="select" onChange="AlternarConteudo(this.value)">
                    <option value="qaa" <%response.Write(qqa_select)%>>Quanto ao Aluno</option>
                    <option value="qpr" <%response.Write(qpr_select)%>>Quanto aos Pais e Respons&aacute;veis </option>
                    <option value="qde" <%response.Write(qde_select)%>>Quanto a Distribui&ccedil;&atilde;o na Escola</option>
                  </select>
                </td>
                <td class="tb_subtit"><div align="right">UNIDADE</div></td>
                <td class="tb_subtit"><div align="left">&nbsp;
                  <select name="unidade" class="borda">
                    <option value="999990"></option>
                    <%		
		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0

		While not RS0.EOF
		NU_Unidade = RS0("NU_Unidade")
		NO_Abr = RS0("NO_Abr")
		
		unidade=unidade*1
		NU_Unidade=NU_Unidade*1
		
		if unidade=NU_Unidade then
			unidade_selected="SELECTED"
		else
			unidade_selected=""		
		end if	
		%>
                    <option value="<%response.Write(NU_Unidade)%>" <%response.Write(unidade_selected)%>>
                      <%response.Write(NO_Abr)%>
                      </option>
                    <%
		RS0.MOVENEXT
		WEND
		%>
                  </select>
                </div></td>
                </tr>
              <tr>
                <td colspan="4">                
<div id="qaa" style="display: none">                
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="50%" class="form_dado_texto"><div align="right">Por G&ecirc;nero</div></td>
                <td width="50%"><label>
                  <input name="qqa" type="radio" class="option_button" id="qqa" value="gen" <% response.Write(gen_checked)%>>
                </label></td>
              </tr>
              <tr>
                <td width="50%" class="form_dado_texto"><div align="right">Por Idade</div></td>
                <td width="50%"><input name="qqa" type="radio" class="option_button" id="qqa" value="idade" <% response.Write(idade_checked)%>></td>
              </tr>
              <tr>
                <td width="50%" class="form_dado_texto"><div align="right">Pelo Ano de Nascimento</div></td>
                <td width="50%"><input name="qqa" type="radio" class="option_button" id="qqa" value="ano" <% response.Write(ano_checked)%>></td>
              </tr>
              <tr>
                <td class="form_dado_texto"><div align="right">Pelo M&ecirc;s de Nascimento</div></td>
                <td><input name="qqa" type="radio" class="option_button" id="qqa" value="mes" <% response.Write(mes_checked)%>></td>
              </tr>
              </table>
</div>   
<div id="qpr" style="display: none">                
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td class="form_dado_texto"><div align="right"> Por Estado Civil&nbsp; </div></td>
                <td width="50%"><label>
                  <input name="qpr" type="radio" class="option_button" id="qpr" value="estc" <% response.Write(estc_checked)%>>
                </label></td>
              </tr>
              <tr>
                <td class="form_dado_texto"><div align="right"> Por Resp. Financeiros </div></td>
                <td width="50%"><input name="qpr" type="radio" class="option_button" id="qpr" value="respf" <% response.Write(respf_checked)%>></td>
              </tr>
              <tr>
                <td class="form_dado_texto"><div align="right"> Por Resp. Pedag&oacute;gicos </div></td>
                <td width="50%"><input name="qpr" type="radio" class="option_button" id="qpr" value="respp" <% response.Write(respp_checked)%>></td>
              </tr>
              </table>
</div>
<div id="qde" style="display: none">                
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td class="form_dado_texto"><div align="right"> Por Unidade </div></td>
                <td width="50%"><label>
                  <input name="qde" type="radio" class="option_button" id="qde" value="unidd" <% response.Write(unidd_checked)%>>
                </label></td>
              </tr>
              <tr>
                <td class="form_dado_texto"><div align="right"> Por Curso </div></td>
                <td width="50%"><input name="qde" type="radio" class="option_button" id="qde" value="curso" <% response.Write(curso_checked)%>></td>
              </tr>
              <tr>
                <td class="form_dado_texto"><div align="right"> Por Etapas&nbsp; </div></td>
                <td width="50%"><input name="qde" type="radio" class="option_button" id="qde" value="etapa" <% response.Write(etapa_checked)%>></td>
              </tr>
              </table>
</div>           
                
                
                </td>
              </tr>
              <tr>
                <td colspan="4"><hr></td>
                </tr>
              <tr>
                <td width="20%">&nbsp;</td>
                <td width="9%">&nbsp;</td>
                <td width="13%">&nbsp;</td>
                <td width="13%"><font size="2" face="Arial, Helvetica, sans-serif">
                  <div align="center">
                    <input name="Submit" type="submit" class="botao_prosseguir" id="Submit5" value="Prosseguir">
                    </div>
                  </font></td>
              </tr>
            </table></form></td>
          </tr>
          <tr>
            <td><hr></td>
          </tr>
           <tr>
            <td height="262">
            <DIV align="center">
<iframe src ="iframe.asp?opt=<%response.Write(tipo_grafico)%>" frameborder ="0" width="1000" height="400" align="middle"> </iframe>
</DIV>
            </td>
          </tr>
        </table>
</td>
          </tr>
		  <tr>
    <td width="1000" height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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