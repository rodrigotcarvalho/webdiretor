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
	
	Set CON2 = Server.CreateObject("ADODB.Connection") 
	ABRIR2 = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON2.Open ABRIR2		

call navegacao (CON,nvg,nivel)
navega=Session("caminho")	




qqa=request.form("qqa")
unidade=request.form("unidade")

gen_checked=""	
idade_checked=""		
ano_checked=""		
mes_checked=""

unidade=unidade*1			
if unidade<>999990 then		
	unidade_selecionada="s"	
	Set RS0 = Server.CreateObject("ADODB.Recordset")
	SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="&unidade
	RS0.Open SQL0, CON0

	no_unidade = RS0("NO_Abr")
else
	unidade_selecionada="n"		
end if	
			
if qqa="gen" then
	gen_checked="CHECKED"
	tipo_grafico="pizza"


	if unidade_selecionada="n" then
		titulo = "Distribuição de alunos na escola por g&ecirc;nero"
		sql="SELECT TB_Matriculas.NU_Ano, TB_Matriculas.CO_Situacao, TB_Alunos.IN_Sexo, Count(TB_Matriculas.CO_Matricula) AS ContarDeCO_Matricula "
		sql=sql&"FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula = TB_Matriculas.CO_Matricula GROUP BY TB_Matriculas.NU_Ano, "
		sql=sql&"TB_Matriculas.CO_Situacao, TB_Alunos.IN_Sexo HAVING (((TB_Matriculas.NU_Ano)="&ano_letivo&") AND ((TB_Matriculas.CO_Situacao)='C'))"			
	else
		titulo = "Distribuição de alunos na unidade "&no_unidade&" por g&ecirc;nero"
		sql="SELECT TB_Matriculas.NU_Ano, TB_Matriculas.CO_Situacao, TB_Matriculas.NU_Unidade, TB_Alunos.IN_Sexo, Count(TB_Matriculas.CO_Matricula) "
		sql=sql&"AS ContarDeCO_Matricula FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula = TB_Matriculas.CO_Matricula GROUP BY "
		sql=sql&"TB_Matriculas.NU_Ano, TB_Matriculas.NU_Unidade, TB_Matriculas.CO_Situacao, TB_Alunos.IN_Sexo HAVING "			
		sql=sql&"(((TB_Matriculas.NU_Ano)="&ano_letivo&") AND ((TB_Matriculas.CO_Situacao)='C') AND ((TB_Matriculas.NU_Unidade)="&unidade&"))"	
	end if				

	Set RS = Server.CreateObject("ADODB.Recordset")
	CONEXAO = sql
	Set RS = CON1.Execute(CONEXAO)

	conta=1
	Do While NOT Rs.EOF
			sex=Rs.Fields("IN_Sexo").value
			qtd=Rs.Fields("ContarDeCO_Matricula").value		
			if conta = 1 then
				vetor_categorias=sex&"-"&qtd
				vetor_faixas=qtd
			else
				vetor_categorias=vetor_categorias&"#!#"&sex&"-"&qtd
				vetor_faixas=vetor_faixas&"#!#"&qtd
			end if
		conta=conta+1			
		RS.MOVENEXT
	LOOP	
	
elseif qqa="idade" then
	idade_checked="CHECKED"	
	tipo_grafico="barra"	
	
	
	if unidade_selecionada="n" then
		titulo = "Distribuição de alunos na escola por idade"
		
			Set RSa = Server.CreateObject("ADODB.Recordset")
			CONEXAOa = "SELECT * from TB_Matriculas where NU_Ano="&ano_letivo&" AND CO_Situacao='C'"
			Set RSa = CON1.Execute(CONEXAOa)			
			conta_matric=1
			While NOT RSa.EOF
					co_aluno=RSa("CO_Matricula")	
					if conta_matric = 1 then
						vetor_alunos=co_aluno
					else
						vetor_alunos=vetor_alunos&", "&co_aluno
					end if
				conta_matric=conta_matric+1			
				RSa.MOVENEXT
			WEND	
		
		sql="SELECT TB_Contatos.TP_Contato, Int((Date()-([DA_Nascimento_Contato]))/365) AS Expr1, " 
		sql=sql&"Count(TB_Contatos.CO_Matricula) AS ContarDeCO_Matricula FROM TB_Contatos where CO_Matricula in("&vetor_alunos&") GROUP BY "
		sql=sql&"TB_Contatos.TP_Contato, Int((Date()-([DA_Nascimento_Contato]))/365) HAVING "
		sql=sql&"(((TB_Contatos.TP_Contato)='ALUNO'))order by Int((Date()-([DA_Nascimento_Contato]))/365) Desc"
	else
		titulo = "Distribuição de alunos na unidade "&no_unidade&" por idade"

		
		Set RSa = Server.CreateObject("ADODB.Recordset")
		CONEXAOa = "SELECT * from TB_Matriculas where NU_Ano="&ano_letivo&" AND CO_Situacao='C' AND NU_Unidade="&unidade
		Set RSa = CON1.Execute(CONEXAOa)			
		conta_matric=1
		While NOT RSa.EOF
				co_aluno=RSa("CO_Matricula")	
				if conta_matric = 1 then
					vetor_alunos=co_aluno
				else
					vetor_alunos=vetor_alunos&", "&co_aluno
				end if
			conta_matric=conta_matric+1			
			RSa.MOVENEXT
		WEND	
		
		sql="SELECT TB_Contatos.TP_Contato, Int((Date()-([DA_Nascimento_Contato]))/365) AS Expr1, " 
		sql=sql&"Count(TB_Contatos.CO_Matricula) AS ContarDeCO_Matricula FROM TB_Contatos where CO_Matricula in("&vetor_alunos&") GROUP BY "
		sql=sql&"TB_Contatos.TP_Contato, Int((Date()-([DA_Nascimento_Contato]))/365) HAVING "
		sql=sql&"(((TB_Contatos.TP_Contato)='ALUNO')) order by Int((Date()-([DA_Nascimento_Contato]))/365) Desc"
	end if		
	
	Set RS = Server.CreateObject("ADODB.Recordset")
	CONEXAO = sql
	Set RS = CON2.Execute(CONEXAO)

	conta=1
	Do While NOT Rs.EOF
			idd=Rs.Fields("Expr1").value
			qtd=Rs.Fields("ContarDeCO_Matricula").value		
			if conta = 1 then
				vetor_categorias=idd
				vetor_faixas=qtd
			else
				vetor_categorias=vetor_categorias&"#!#"&idd
				vetor_faixas=vetor_faixas&"#!#"&qtd
			end if
		conta=conta+1			
		RS.MOVENEXT
	LOOP		
	
elseif qqa="ano" then
	ano_checked="CHECKED"
	tipo_grafico="barra"	
	
	
	if unidade_selecionada="n" then
		titulo = "Distribuição de alunos na escola por Ano de Nascimento"
		
			Set RSa = Server.CreateObject("ADODB.Recordset")
			CONEXAOa = "SELECT * from TB_Matriculas where NU_Ano="&ano_letivo&" AND CO_Situacao='C'"
			Set RSa = CON1.Execute(CONEXAOa)			
			conta_matric=1
			While NOT RSa.EOF
					co_aluno=RSa("CO_Matricula")	
					if conta_matric = 1 then
						vetor_alunos=co_aluno
					else
						vetor_alunos=vetor_alunos&", "&co_aluno
					end if
				conta_matric=conta_matric+1			
				RSa.MOVENEXT
			WEND	
		
		sql="SELECT TB_Contatos.TP_Contato, Year([DA_Nascimento_Contato]) AS Expr1, " 
		sql=sql&"Count(TB_Contatos.CO_Matricula) AS ContarDeCO_Matricula FROM TB_Contatos where CO_Matricula in("&vetor_alunos&") GROUP BY "
		sql=sql&"TB_Contatos.TP_Contato, Year([DA_Nascimento_Contato]) HAVING "
		sql=sql&"(((TB_Contatos.TP_Contato)='ALUNO'))order by Year([DA_Nascimento_Contato]) Desc"
	else
		titulo = "Distribuição de alunos na unidade "&no_unidade&" por Ano de Nascimento"

		
		Set RSa = Server.CreateObject("ADODB.Recordset")
		CONEXAOa = "SELECT * from TB_Matriculas where NU_Ano="&ano_letivo&" AND CO_Situacao='C' AND NU_Unidade="&unidade
		Set RSa = CON1.Execute(CONEXAOa)			
		conta_matric=1
		While NOT RSa.EOF
				co_aluno=RSa("CO_Matricula")	
				if conta_matric = 1 then
					vetor_alunos=co_aluno
				else
					vetor_alunos=vetor_alunos&", "&co_aluno
				end if
			conta_matric=conta_matric+1			
			RSa.MOVENEXT
		WEND	
		
		sql="SELECT TB_Contatos.TP_Contato, Year([DA_Nascimento_Contato]) AS Expr1, " 
		sql=sql&"Count(TB_Contatos.CO_Matricula) AS ContarDeCO_Matricula FROM TB_Contatos where CO_Matricula in("&vetor_alunos&") GROUP BY "
		sql=sql&"TB_Contatos.TP_Contato, Year([DA_Nascimento_Contato]) HAVING "
		sql=sql&"(((TB_Contatos.TP_Contato)='ALUNO')) order by Year([DA_Nascimento_Contato]) Desc"
	end if		
	
	Set RS = Server.CreateObject("ADODB.Recordset")
	CONEXAO = sql
	Set RS = CON2.Execute(CONEXAO)

	conta=1
	Do While NOT Rs.EOF
			ano=Rs.Fields("Expr1").value
			qtd=Rs.Fields("ContarDeCO_Matricula").value		
			if conta = 1 then
				vetor_categorias=ano
				vetor_faixas=qtd
			else
				vetor_categorias=vetor_categorias&"#!#"&ano
				vetor_faixas=vetor_faixas&"#!#"&qtd
			end if
		conta=conta+1			
		RS.MOVENEXT
	LOOP	
					
elseif qqa="mes" then
	mes_checked="CHECKED"
	tipo_grafico="barra"	

	
	if unidade_selecionada="n" then
		titulo = "Distribuição de alunos na escola por M&ecirc;s de Nascimento"
		
			Set RSa = Server.CreateObject("ADODB.Recordset")
			CONEXAOa = "SELECT * from TB_Matriculas where NU_Ano="&ano_letivo&" AND CO_Situacao='C'"
			Set RSa = CON1.Execute(CONEXAOa)			
			conta_matric=1
			While NOT RSa.EOF
					co_aluno=RSa("CO_Matricula")	
					if conta_matric = 1 then
						vetor_alunos=co_aluno
					else
						vetor_alunos=vetor_alunos&", "&co_aluno
					end if
				conta_matric=conta_matric+1			
				RSa.MOVENEXT
			WEND	
		
		sql="SELECT TB_Contatos.TP_Contato, Month([DA_Nascimento_Contato]) AS Expr1, " 
		sql=sql&"Count(TB_Contatos.CO_Matricula) AS ContarDeCO_Matricula FROM TB_Contatos where CO_Matricula in("&vetor_alunos&") GROUP BY "
		sql=sql&"TB_Contatos.TP_Contato, Month([DA_Nascimento_Contato]) HAVING "
		sql=sql&"(((TB_Contatos.TP_Contato)='ALUNO'))order by Month([DA_Nascimento_Contato]) Desc"
	else
		titulo = "Distribuição de alunos na unidade "&no_unidade&" por M&ecirc;s de Nascimento"

		
		Set RSa = Server.CreateObject("ADODB.Recordset")
		CONEXAOa = "SELECT * from TB_Matriculas where NU_Ano="&ano_letivo&" AND CO_Situacao='C' AND NU_Unidade="&unidade
		Set RSa = CON1.Execute(CONEXAOa)			
		conta_matric=1
		While NOT RSa.EOF
				co_aluno=RSa("CO_Matricula")	
				if conta_matric = 1 then
					vetor_alunos=co_aluno
				else
					vetor_alunos=vetor_alunos&", "&co_aluno
				end if
			conta_matric=conta_matric+1			
			RSa.MOVENEXT
		WEND	
		
		sql="SELECT TB_Contatos.TP_Contato, Month([DA_Nascimento_Contato]) AS Expr1, " 
		sql=sql&"Count(TB_Contatos.CO_Matricula) AS ContarDeCO_Matricula FROM TB_Contatos where CO_Matricula in("&vetor_alunos&") GROUP BY "
		sql=sql&"TB_Contatos.TP_Contato, Month([DA_Nascimento_Contato]) HAVING "
		sql=sql&"(((TB_Contatos.TP_Contato)='ALUNO')) order by Month([DA_Nascimento_Contato]) Desc"
	end if		
	
	Set RS = Server.CreateObject("ADODB.Recordset")
	CONEXAO = sql
	Set RS = CON2.Execute(CONEXAO)

	conta=1
	Do While NOT Rs.EOF
			mes=Rs.Fields("Expr1").value
			qtd=Rs.Fields("ContarDeCO_Matricula").value	
			
			select case mes
				case 1
					no_mes="Jan"
				case 2
					no_mes="Fev"				
				case 3
					no_mes="Mar"			
				case 4
					no_mes="Abr"				
				case 5
					no_mes="Mai"
				case 6
					no_mes="Jun"
				case 7
					no_mes="Jul"				
				case 8
					no_mes="Ago"			
				case 9
					no_mes="Set"				
				case 10
					no_mes="Out"
				case 11
					no_mes="Nov"				
				case 12
					no_mes="Dez"
			end select
					
			if conta = 1 then
				vetor_categorias=no_mes
				vetor_faixas=qtd
			else
				vetor_categorias=vetor_categorias&"#!#"&no_mes
				vetor_faixas=vetor_faixas&"#!#"&qtd
			end if
		conta=conta+1			
		RS.MOVENEXT
	LOOP						
end if		




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
      <%	call mensagens(4,712,0,0) 
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
  <div id="qaa">                
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
        <td width="50%" height="15" class="form_dado_texto"><div align="right">Por G&ecirc;nero</div></td>
        <td width="50%" height="15">
          <input name="qqa" type="radio" class="option_button" id="qqa" value="gen" <% response.Write(gen_checked)%>>
          </td>
      </tr>
      <tr>
        <td width="50%" height="15" class="form_dado_texto"><div align="right">Por Idade</div></td>
        <td width="50%" height="15"><input name="qqa" type="radio" class="option_button" id="qqa" value="idade" <% response.Write(idade_checked)%>></td>
        </tr>
      <tr>
        <td width="50%" height="15" class="form_dado_texto"><div align="right">Pelo Ano de Nascimento</div></td>
        <td width="50%" height="15"><input name="qqa" type="radio" class="option_button" id="qqa" value="ano" <% response.Write(ano_checked)%>></td>
        </tr>
      <tr>
        <td height="15" class="form_dado_texto"><div align="right">Pelo M&ecirc;s de Nascimento</div></td>
        <td height="15"><input name="qqa" type="radio" class="option_button" id="qqa" value="mes" <% response.Write(mes_checked)%>></td>
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