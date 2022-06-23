<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
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


Set CON = Server.CreateObject("ADODB.Connection") 
ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
CON.Open ABRIR
	
Set CON0 = Server.CreateObject("ADODB.Connection") 
ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
CON0.Open ABRIR0

Set CON1 = Server.CreateObject("ADODB.Connection") 
ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
CON1.Open ABRIR1	

qde=request.form("qde")
unidade=request.form("unidade")

unidd_checked=""
curso_checked=""				
etapa_checked=""
	
unidade=unidade*1			
if unidade<>999990 then		
	unidade_selecionada="s"	
	Set RS0 = Server.CreateObject("ADODB.Recordset")
	SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="&unidade
	RS0.Open SQL0, CON0

	no_unidade = RS0("NO_Abr")

	SQL0a = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&" order by CO_Curso"	
	SQL0b = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&" order by CO_Curso,CO_Etapa"		
else

	SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"	
	SQL0a = "SELECT * FROM TB_Unidade_Possui_Etapas order by CO_Curso"	
	SQL0b = "SELECT * FROM TB_Unidade_Possui_Etapas order by CO_Curso,CO_Etapa"	
	
	unidade_selecionada="n"		
end if			
				
if qde="unidd" then
	unidd_checked="CHECKED"
	tipo_grafico="pizza"	

	Set RS0a = Server.CreateObject("ADODB.Recordset")	
	RS0a.Open SQL0, CON0

		nu_unidades_check = 1
		nu_alunos_check = 1
		
		While not RS0a.EOF
			qtd_aluno=0
			unidade_tabela = RS0a("NU_Unidade")
			NO_Abr= RS0a("NO_Abr")	
		
			if 	unidade_selecionada="s"	 then
				titulo = "Distribuição de alunos por unidade na unidade "&NO_Abr		
			else
				titulo = "Distribuição de alunos por unidade"
			end if
			
			Set RSt0 = Server.CreateObject("ADODB.Recordset")
			SQLt0 ="SELECT Count(CO_Matricula) as Qtd_Alunos FROM TB_Matriculas where NU_Ano ="& ano_letivo&" AND NU_Unidade ="& unidade_tabela&" AND CO_Situacao='C'"
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
			
		RS0a.MOVENEXT
		WEND
	
	
			
elseif qde="curso" then
	curso_checked="CHECKED"
	tipo_grafico="coluna_agrupada"	
	
	Set RS0a = Server.CreateObject("ADODB.Recordset")	
	RS0a.Open SQL0a, CON0
		
	nu_alunos_check = 1
			
	While not RS0a.EOF
	qtd_aluno=0
	curso_tabela = RS0a("CO_Curso")
	
	Set RS2 = Server.CreateObject("ADODB.Recordset")
	SQL2 = "SELECT * FROM TB_Curso where CO_Curso='"&curso_tabela&"' order by CO_Curso"
	RS2.Open SQL2, CON0				
	
	NO_Abr = RS2("NO_Abreviado_Curso")	

			
		Set RSt0 = Server.CreateObject("ADODB.Recordset")
		if unidade_selecionada="s"then
			titulo = "Distribuição de alunos por curso na unidade "&no_unidade		
			SQLt0 ="SELECT Count(CO_Matricula) as Qtd_Alunos FROM TB_Matriculas where NU_Ano ="& ano_letivo&" AND NU_Unidade ="& unidade&"  AND CO_Curso ='"& curso_tabela &"' AND CO_Situacao='C'"
		else
			titulo = "Distribuição de alunos por curso"
			SQLt0 ="SELECT Count(CO_Matricula) as Qtd_Alunos FROM TB_Matriculas where NU_Ano ="& ano_letivo&" AND CO_Curso ='"& curso_tabela &"' AND CO_Situacao='C'"			
		end if	
		RSt0.Open SQLt0, CON1	
		
		if RSt0.EOF then
			qtd_aluno=0			
		else
			qtd_aluno=RSt0("Qtd_Alunos")
			if qtd_aluno="" or isnull(qtd_aluno) then
				qtd_aluno=0
			end if
		end if
		
		If Not IsArray(faixas) Then faixas = Array() End if

		If Not IsArray(categorias) Then categorias = Array() End if
		If InStr(Join(categorias), NO_Abr) = 0 Then
			ReDim preserve categorias(UBound(categorias)+1)
			categorias(Ubound(categorias)) = NO_Abr	
			ReDim preserve faixas(UBound(faixas)+1)
			faixas(Ubound(faixas)) = qtd_aluno	
		end if		
							
	RS0a.MOVENEXT
	WEND
	
	for va=0 to ubound(faixas)
			if nu_alunos_check = 1 then
				vetor_faixas=faixas(va)
				vetor_categorias=categorias(va)&"-"&faixas(va)
			else
				vetor_faixas=vetor_faixas&"#!#"&faixas(va)
				vetor_categorias=vetor_categorias&"#!#"&categorias(va)&"-"&faixas(va)					
			end if
		nu_alunos_check=nu_alunos_check+1	
	Next	

					
elseif qde="etapa" then
	etapa_checked="CHECKED"	
	tipo_grafico="coluna_agrupada"	
	
	Set RS0a = Server.CreateObject("ADODB.Recordset")	
	RS0a.Open SQL0b, CON0

	nu_alunos_check = 1
			
	While not RS0a.EOF
	qtd_aluno=0
	etapa_tabela = RS0a("CO_Etapa")
	
	Set RS2 = Server.CreateObject("ADODB.Recordset")
	SQL2 = "SELECT * FROM TB_Etapa where CO_Etapa='"&etapa_tabela&"' order by CO_Etapa"
	RS2.Open SQL2, CON0				
	
	NO_Abr = RS2("NO_Etapa")	
	
		Set RSt0 = Server.CreateObject("ADODB.Recordset")
		if unidade_selecionada="s"then
			titulo = "Distribuição de alunos por etapa na unidade "&no_unidade			
			SQLt0 ="SELECT Count(CO_Matricula) as Qtd_Alunos FROM TB_Matriculas where NU_Ano ="& ano_letivo&" AND NU_Unidade ="& unidade&"  AND CO_Etapa ='"& etapa_tabela &"' AND CO_Situacao='C'"
		else
			titulo = "Distribuição de alunos por etapa"		
			SQLt0 ="SELECT Count(CO_Matricula) as Qtd_Alunos FROM TB_Matriculas where NU_Ano ="& ano_letivo&" AND CO_Etapa ='"& etapa_tabela &"' AND CO_Situacao='C'"			
		end if	
		RSt0.Open SQLt0, CON1	
		
		if RSt0.EOF then
			qtd_aluno=0			
		else
			qtd_aluno=RSt0("Qtd_Alunos")
			if qtd_aluno="" or isnull(qtd_aluno) then
				qtd_aluno=0
			end if
		end if
		
		If Not IsArray(faixas) Then faixas = Array() End if

		If Not IsArray(categorias) Then categorias = Array() End if
		If InStr(Join(categorias), NO_Abr) = 0 Then
			ReDim preserve categorias(UBound(categorias)+1)
			categorias(Ubound(categorias)) = NO_Abr	
			ReDim preserve faixas(UBound(faixas)+1)
			faixas(Ubound(faixas)) = qtd_aluno	
		end if		
							
	RS0a.MOVENEXT
	WEND
	
	for va=0 to ubound(faixas)
			if nu_alunos_check = 1 then
				vetor_faixas=faixas(va)
				vetor_categorias=categorias(va)&"-"&faixas(va)
			else
				vetor_faixas=vetor_faixas&"#!#"&faixas(va)
				vetor_categorias=vetor_categorias&"#!#"&categorias(va)&"-"&faixas(va)					
			end if
		nu_alunos_check=nu_alunos_check+1	
	Next	
		
		While not RS0a.EOF
		qtd_aluno=0
		etapa_tabela = RS0a("CO_Etapa")

		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Etapa where CO_Etapa='"&etapa_tabela&"' order by CO_Etapa"
		RS2.Open SQL2, CON0				
		
		NO_Abr = RS2("NO_Etapa")	
		
			Set RSt0 = Server.CreateObject("ADODB.Recordset")
			SQLt0 ="SELECT Count(CO_Matricula) as Qtd_Alunos FROM TB_Matriculas where NU_Ano ="& ano_letivo&" AND CO_Etapa ='"& etapa_tabela &"' AND CO_Situacao='C'"
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
			
		RS0a.MOVENEXT
		WEND
						
end if				

	
call navegacao (CON,nvg,nivel)
navega=Session("caminho")	

session("faixas")=vetor_faixas
session("categorias")=vetor_categorias
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
function submitfuncao()  
{
   var f=document.forms[0]; 
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
  {    alert("Por favor selecione um Tema Estatístico!")
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
            <td width="653" class="<%response.Write(classe_corpo_un)%>"><form name="grafico" method="post" action="grafico.asp">
            <table width="1000" border="0" cellspacing="0" cellpadding="0">
              <tr class="tb_subtit">
                <td class="tb_tit">Preencha os campos abaixo</td>
                <td class="tb_tit">&nbsp;</td>
                <td class="tb_tit">&nbsp;</td>
                <td class="tb_tit">&nbsp;</td>
              </tr>
               <tr>
                <td colspan="4">                
           
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="tb_subtit">
                <td class="form_dado_texto"><div align="right" class="tb_subtit">UNIDADE</div></td>
                <td>&nbsp;
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
                  </select></td>
              </tr>
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
          <tr class="form_corpo">
            <td><div align="center"><%response.Write(titulo)%></div></td>
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