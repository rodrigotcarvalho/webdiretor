<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->

<% 
session("nvg")=""
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
nvg=request.QueryString("nvg")
opt = request.QueryString("opt")
session("nvg")=nvg
ano_info=nivel&"-"&nvg&"-"&ano_letivo



		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
call VerificaAcesso (CON,nvg,nivel)
autoriza=Session("autoriza")

call navegacao (CON,nvg,nivel)
navega=Session("caminho")


if opt="" or isnull(opt) then
	display="select"
	
	onload="onLoad=""redimensiona();MM_showHideLayers('divTabela','','hide');AlternarMensagem('divMensagem2');MM_callJS('document.busca.busca1.focus()')"""
	obr_mapa=unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma&"$!$"&periodo&"$!$"&acumulado&"$!$"&qto_falta&"$!$"&ano_letivo	
	dados_msg="t$$$"&no_unidade&"#!#"&no_curso&"#!#"&no_etapa&"#!#"&data_grav&"#!#"&hora_grav&"$$$"&obr_mapa	

elseif opt="err1" or opt="err2" then
	display="select"
	
	onload="onLoad=""redimensiona();MM_showHideLayers('divTabela','','hide');AlternarMensagem('divMensagem3');MM_callJS('document.busca.busca1.focus()')"""

	obr_mapa=unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma&"$!$"&periodo&"$!$"&acumulado&"$!$"&qto_falta&"$!$"&ano_letivo	
	dados_msg="t$$$"&no_unidade&"#!#"&no_curso&"#!#"&no_etapa&"#!#"&data_grav&"#!#"&hora_grav&"$$$"&obr_mapa	

elseif opt="err707" or opt="err9713" then
	display="bloq"
	
	onload="onLoad=""redimensiona();MM_showHideLayers('divTabela','','hide');MM_callJS('document.busca.busca1.focus()')"""
	
	obr_mapa=unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma&"$!$"&periodo&"$!$"&acumulado&"$!$"&qto_falta&"$!$"&ano_letivo	
	dados_msg=request.QueryString("obr")
elseif opt="acc" then
	display="ask2"

	dados_msg=request.QueryString("obr")
	
	separa_dados=split(dados_msg,"$$$")
	tipo_busca=separa_dados(0)	
	dados_funcao=split(separa_dados(1),"$!$")

	unidade = dados_funcao(0)
	curso = dados_funcao(1)
	co_etapa = dados_funcao(2)
	turma = dados_funcao(3)
	periodo_form = dados_funcao(4)	
	

	obr_mapa=unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma&"$!$"&periodo_form&"$!$"&ano_letivo&"$!$ficha"		
	obr_log=unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&periodo&"_"&ano_letivo
	

	
	if tipo="a" then
		dados_msg=tipo_busca&"$!$"&SESSION("aluno_boletim")&"$$$"&no_unidade&"#!#"&no_curso&"#!#"&no_etapa&"#!#"&data_grav&"#!#"&hora_grav&"$$$"&obr_mapa	
	else
		dados_msg=tipo_busca&"$$$"&no_unidade&"#!#"&no_curso&"#!#"&no_etapa&"#!#"&data_grav&"#!#"&hora_grav&"$$$"&obr_mapa	
	end if


	onload="onLoad=""redimensiona();AlternarMensagem('divMensagem3')"""
	
elseif opt="acc_mult" then
	display="ask2"

	dados_msg=request.QueryString("obr")
	
	separa_dados=split(dados_msg,"$$$")
	tipo_busca=separa_dados(0)	
	dados_funcao=split(separa_dados(1),"$!$")

	unidade = dados_funcao(0)
	curso = dados_funcao(1)
	co_etapa = dados_funcao(2)
	turma = dados_funcao(3)
	periodo_form = dados_funcao(4)	
	

	obr_mapa=unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma&"$!$"&periodo_form&"$!$"&ano_letivo&"$!$ficha"		
	obr_log=unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&periodo&"_"&ano_letivo
	

	
	if tipo="a" then
		dados_msg=tipo_busca&"$!$"&SESSION("aluno_boletim")&"$$$"&no_unidade&"#!#"&no_curso&"#!#"&no_etapa&"#!#"&data_grav&"#!#"&hora_grav&"$$$"&obr_mapa	
	else
		dados_msg=tipo_busca&"$$$"&no_unidade&"#!#"&no_curso&"#!#"&no_etapa&"#!#"&data_grav&"#!#"&hora_grav&"$$$"&obr_mapa	
	end if


	onload="onLoad=""MM_showHideLayers('divTabela','','show');AlternarMensagem('divMensagem4')"""	
	
elseif opt="ask" then

	obr=request.QueryString("obr")
	divide_obr=split(obr,"$$$")
	tipo_obr=split(divide_obr(0),"$!$")
	tipo=tipo_obr(0)
		
	dados_obr=split(divide_obr(1),"$!$")	
	unidade = dados_obr(0)
	curso = dados_obr(1)
	co_etapa = dados_obr(2)
	turma = dados_obr(3)
	periodo_form = dados_obr(4)
	
	if tipo="a" then
		cod_cons=tipo_obr(1)	
		display="ask1"
		onload="onLoad=""MM_showHideLayers('divTabela','','hide');AlternarMensagem('divMensagem1')"""
	else
		display="ask2"		
		onload="onLoad=""MM_showHideLayers('divTabela','','hide');AlternarMensagem('divMensagem1')"""
	end if
	obr_mapa=unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma&"$!$"&periodo&"$!$"&ano_letivo&"$!$ficha"		
	obr_log=unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&periodo&"_"&ano_letivo
	
	dados_msg=session ("dados_msg")

	
'else
'
'	onload="onLoad=""redimensiona();MM_showHideLayers('divTabela','','hide');AlternarMensagem('divMensagem2')"""

elseif opt="search" then
	busca1=request.form("busca1") 
	busca2=request.form("busca2")
	unidade=request.form("unidade") 
	curso=request.form("curso")
	co_etapa=request.form("etapa") 
	turma=request.form("turma")	
	periodo=request.form("periodo")	
	valor3=session("valor3")
	valor4=session("valor4")
	valor5=session("valor5")
	valor6=session("valor6")
	valor7=session("valor7")	

'response.Write(valor1&"-"&valor2&"-"&valor3&"-"&valor4&"-"&valor5&"-"&valor6&"-"&valor7)
'response.end()

session("WS-DO-DPA-EFI-periodo")=periodo

	if unidade=999990 or unidade="999990" then		
		if busca1 ="" then
			query = busca2
			mensagem=304
		elseif busca2 ="" then
			query = busca1 
			mensagem=303
		end if 
	
		teste = IsNumeric(query)
		if teste = TRUE Then
	  
			Set RS = Server.CreateObject("ADODB.Recordset")
			SQL ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& query &" and TB_Matriculas.NU_Ano="&ano_letivo
			RS.Open SQL, CON1
				
			if RS.EOF Then
				display="reselect"
			else	
				obr=query&"$!$"&periodo
				response.Redirect("avalia.asp?opt=1&obr="&obr)					
				'response.Redirect("../../../../relatorios/swd025.asp?ori=ebe&opt=01&cod_cons="&query&"&prd="&periodo)	
			end if
		
		ELSE	

			busca=busca_por_nome(query,CON1,"alun")
			alunos_encontrados = split(busca, "#!#" )
			
			if ubound(alunos_encontrados)=-1 then
				display="reselect"
			elseif ubound(alunos_encontrados)=0 then
				cod_cons=alunos_encontrados(0)
				obr=cod_cons&"$!$"&periodo
				response.Redirect("avalia.asp?opt=1&obr="&obr)		
				'response.Redirect("../../../../relatorios/swd025.asp?ori=ebe&opt=01&cod_cons="&cod_cons&"&prd="&periodo)	
			else
				display="list"
			end if
		END IF
	else
		if curso="" or isnull(curso) then
			curso=valor4
		end if
		if co_etapa="" or isnull(co_etapa) then
			co_etapa=valor5
		end if
		if turma="" or isnull(turma) then
			turma=valor6
		end if
		if periodo="" or isnull(periodo) then
			periodo=valor7
		end if
	
		obr=unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma&"$!$"&periodo
		response.Redirect("avalia.asp?opt=2&obr="&obr)		
		'response.Redirect("../../../../relatorios/swd025.asp?ori=ebe&opt=02&obr="&obr)			
	end if	
	
	onload="onLoad=""redimensiona();MM_showHideLayers('divTabela','','hide');AlternarMensagem('divMensagem2');MM_callJS('document.busca.busca1.focus()')"""	
	obr_mapa=unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma&"$!$"&periodo&"$!$"&acumulado&"$!$"&qto_falta&"$!$"&ano_letivo	
	dados_msg="t$$$"&no_unidade&"#!#"&no_curso&"#!#"&no_etapa&"#!#"&data_grav&"#!#"&hora_grav&"$$$"&obr_mapa	
elseif opt="listall" then
	display="listall"
	
	onload="onLoad=""redimensiona();MM_showHideLayers('divTabela','','hide');AlternarMensagem('divMensagem2')"""
	
	obr_mapa=unidade&"$!$"&curso&"$!$"&co_etapa&"$!$"&turma&"$!$"&periodo&"$!$"&acumulado&"$!$"&qto_falta&"$!$"&ano_letivo	
	dados_msg="t$$$"&no_unidade&"#!#"&no_curso&"#!#"&no_etapa&"#!#"&data_grav&"#!#"&hora_grav&"$$$"&obr_mapa			
end if

	tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")		

%>
<html>
<head>
<title>Web Diretor</title>
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
var currentlyActiveInputRef = false;
var currentlyActiveInputClassName = false;

function highlightActiveInput() {
  if(currentlyActiveInputRef) {
    currentlyActiveInputRef.className = currentlyActiveInputClassName;
  }
  currentlyActiveInputClassName = this.className;
  this.className = 'inputHighlighted';
  currentlyActiveInputRef = this;
}

function blurActiveInput() {
  this.className = currentlyActiveInputClassName;
}

function initInputHighlightScript() {
  var tags = ['INPUT','TEXTAREA'];
  for(tagCounter=0;tagCounter<tags.length;tagCounter++){
    var inputs = document.getElementsByTagName(tags[tagCounter]);
    for(var no=0;no<inputs.length;no++){
      if(inputs[no].className && inputs[no].className=='doNotHighlightThisInput')continue;
      if(inputs[no].tagName.toLowerCase()=='textarea' || (inputs[no].tagName.toLowerCase()=='input' && inputs[no].type.toLowerCase()=='text')){
        inputs[no].onfocus = highlightActiveInput;
        inputs[no].onblur = blurActiveInput;
      }
    }
  }
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

function MM_showHideLayers() { //v6.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
    obj.visibility=v; }
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function checksubmit()
{
  if (document.busca.busca1.value != "" && document.busca.busca2.value != "")
  {    alert("Por favor digite SOMENTE uma opção de busca!")
    document.busca.busca1.value = "";
	document.busca.busca2.value = "";    
    document.busca.busca1.focus()
    return false
  }
  if (document.busca.busca1.value != "" && document.busca.unidade.value != "999990")
  {    alert("Por favor digite SOMENTE uma opção de busca!")
    document.busca.busca1.value = "";
	document.busca.busca2.value = "";   
	var combo = document.getElementById("unidade");
	combo.options[0].selected = "true";
	var combo1 = document.getElementById("curso");
	combo1.options[0].selected = "true";
	var combo2 = document.getElementById("etapa");
	combo2.options[0].selected = "true";
	var combo3 = document.getElementById("turma");
	combo3.options[0].selected = "true";
	var combo4 = document.getElementById("periodo");
	combo4.options[0].selected = "true";	
    document.busca.busca1.focus()
    return false
  }
  if (document.busca.unidade.value != "999990" && document.busca.busca2.value != "")
  {    alert("Por favor digite SOMENTE uma opção de busca!")
    document.busca.busca1.value = "";
	document.busca.busca2.value = "";   
	var combo = document.getElementById("unidade");
	combo.options[0].selected = "true";
    document.busca.busca1.focus()
    return false
  }  
    if (document.busca.busca1.value == "" && document.busca.busca2.value == "" && document.busca.unidade.value == 999990)
  {    alert("Por favor digite uma opção de busca!")
    document.busca.busca1.focus()
    return false
  }
  if (document.busca.busca1.value == "" && document.busca.busca2.value == "" && document.busca.curso.value == "999990")
  {    alert("É necessário preencher pelo menos Unidade, Curso e Etapa!")
	var combo = document.getElementById("unidade");
	combo.options[0].selected = "true";
    document.busca.busca1.focus()
    return false
  }  
  if (document.busca.busca1.value == "" && document.busca.busca2.value == "" && document.busca.etapa.value == "999990")
  {    alert("É necessário preencher pelo menos Unidade, Curso e Etapa!")
	var combo = document.getElementById("unidade");
	combo.options[0].selected = "true";
	var combo2 = document.getElementById("curso");
	combo2.options[0].selected = "true";	
    document.busca.busca1.focus()
    return false
  }  
  MM_showHideLayers('carregando','','show','carregando_fundo','','show')      
  return true
}

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}function submitfuncao()  
{
   var f=document.forms[4]; 
      f.submit(); 
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
document.all.divEtapa.innerHTML ="<select class=select_style id=etapa><option value='99990' selected></option></select>"
document.all.divTurma.innerHTML = "<select class=select_style id=turma><option value='99990' selected></option></select>"
GuardaUnidade(uTipo)
                                                           }
                                               }

                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=e7", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select class=select_style id=turma><option value='99990' selected></option></select>"
GuardaCurso(cTipo)
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
GuardaEtapa(eTipo)																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
								   
	function recuperarPeriodo(tTipo)
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
gravarTurma(tTipo)																		   
                                                           }
                                               }

                                               oHTTPRequest.send("t_pub=" + tTipo);
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
	 function GuardaPeriodo(p)
		{
	
		   var oHTTPRequest = createXMLHTTP();
	
		   oHTTPRequest.open("post", "../../../../../global/guarda_valores_digitados.asp?opt=valor7", true);
	
		   oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
	
		   oHTTPRequest.onreadystatechange=function() {
	
								   }
	
		   oHTTPRequest.send("valor7=" + p);
		}								   
 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}

function MM_showHideLayers() { //v9.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) 
  with (document) if (getElementById && ((obj=getElementById(args[i]))!=null)) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
    obj.visibility=v; }
}
function redimensiona(){
//o 120 e se refere ao tamanho de cabeçalho do navegador
    y = parseInt((screen.availHeight - 120 - 135 - 70 - 40));
    document.getElementById('carregando_fundo').style.height = y;
}
function go_there()
{
// var where_to= confirm("<%'response.Write(javascript)%>");
// if (where_to== true)
// {

   window.location="gera_pdf.asp?obr=<%response.Write(obr_mapa)%>";
// }
// else
// {
//  window.location="<%'response.Write("avalia.asp?opt=rgnrt")%>";
//  }

}
var timeout         = 5000;
var closetimer		= 0;

function mclose()
{
	div1 = document.getElementById("carregando");
	div2 = document.getElementById("carregando_fundo");	
	div1.style.visibility = 'hidden';
	div2.style.visibility = 'hidden';	
}


function mclosetime()
{
	closetimer = window.setTimeout(mclose, timeout);
}
function mensagem(conteudo)
	{
		this.conteudo = conteudo;
	}
 
	var arAbas = new Array();
	arAbas[0] = new mensagem('divMensagem1');
	arAbas[1] = new mensagem('divMensagem2');
	arAbas[2] = new mensagem('divMensagem3');	
	arAbas[3] = new mensagem('divMensagem4');		
 
	function AlternarMensagem(conteudo)
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
<body background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" <%response.Write(onload)%>>
<form action="index.asp?opt=search&nvg=<%=nvg%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()">

<%call cabecalho(nivel)
%>


<table width="1000" height="670" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr>             
    <td width="1000" height="10" colspan="5" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
    </td>
          </tr>
<%if opt="err1" then%>
          <tr> 
            
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,700,1,0) %>
    </td>
          </tr> 
<%elseif opt="err2" then%>
          <tr> 
            
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,701,1,0) %>
    </td>
          </tr> 
<%elseif opt="err707" then%>
          <tr> 
            
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,707,1,dados_msg) %>
    </td>
         </tr>
<%elseif opt="err9713" then%>
          <tr> 
            
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,9713,1,dados_msg) %>
    </td>
         </tr>
<%end if
if display="listall" then%>
          <tr> 
            
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,1,0,0) %>
    </td>
          </tr>          
<%elseif display="reselect" then%>
            <tr>              
    <td height="10" colspan="5"> 
      <%call mensagens(nivel,mensagem,1,0) %>
    </td>
			   </tr>          
<%
end if
if display="select" or display="reselect" or display="list" or display="ask1" or display="ask2" then%>
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
    <div id="divMensagem1" style="display: none">
      <%
	  if autoriza="no" then
	  	call mensagens(nivel,9700,1,0) 	  
	  else
  	  	call mensagens(nivel,656,0,dados_msg)
	  end if
	  call ultimo(0) %>
      </div>
    <div id="divMensagem2" style="display: none">
      <%
	  if autoriza="no" then
	  	call mensagens(nivel,9700,1,0) 	  
	  else
	  '	call mensagens(nivel,1,0,0) 
		call mensagens(nivel,300,0,0) 
	  end if
	  call ultimo(0) %>
      </div>   
    <div id="divMensagem3" style="display: none">
      <%
	  if autoriza="no" then
	  	call mensagens(nivel,9700,1,0) 	  
	  else
	  	call mensagens(nivel,657,0,dados_msg) 
	  end if
	  call ultimo(0) %>
      </div>  
    <div id="divMensagem4" style="display: none">
      <%
	  if autoriza="no" then
	  	call mensagens(nivel,9700,1,0) 	  
	  else
	  	call mensagens(nivel,662,0,dados_msg) 
	  end if
	  call ultimo(0) %>
      </div>               
    </td>
			  </tr>	
<%end if%>
<%if display="select" or display="reselect" or display="list" or display="bloq" then%>                      	  
          <tr class="tb_tit">             
      <td height="10" colspan="5">Preencha um dos campos abaixo</td>
          </tr>
          
          <TR>
      <td height="10" valign="top"> 

                <table width="1000" border="0" cellpadding="0" cellspacing="0">
          <tr>          
            <td width="150"  height="10"> 
              <div align="right"><font class="form_dado_texto"> 
                Matr&iacute;cula:</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
            </strong></font></div></td>
            
            <td width="50" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="busca1" type="text" class="textInput" id="busca1" size="12">
              </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font></td>
            
            <td width="150" height="10"> 
              <div align="right"><font class="form_dado_texto"> Nome: </font></div></td>
            
            <td width="500" height="10" ><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="busca2" type="text" class="textInput" id="busca2" size="55" maxlength="50">
              </font></td>
            
            <td width="150" height="10">&nbsp;</td>
          </tr>
		  </table>
		  </td>
		  </TR>
           <tr>                   
    <td height="10" colspan="5" valign="top">  
<table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="200" class="tb_subtit"> 
                    <div align="center">UNIDADE 
                    </div></td>
                  <td width="200" class="tb_subtit"> 
                    <div align="center">CURSO 
                    </div></td>
                  <td width="200" class="tb_subtit"> 
                    <div align="center">ETAPA 
                    </div></td>
                  <td width="200" class="tb_subtit"> 
                    <div align="center">TURMA 
                    </div></td>
                  <td width="200" class="tb_subtit"><div align="center">PER&Iacute;ODO</div></td>
                </tr>
                <tr> 
                  <td width="200"> 
                    <div align="center"> 
                      <select name="unidade" class="select_style" id="unidade" onChange="recuperarCurso(this.value)">
                        <option value="999990" selected></option>
                        <%		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0
While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")
%>
                        <option value="<%response.Write(NU_Unidade)%>"> 
                        <%response.Write(NO_Abr)%>
                        </option>
                        <%RS0.MOVENEXT
WEND
%>
                      </select>
                    </div></td>
                  <td width="200"> 
                    <div align="center"> 
                      <div id="divCurso"> 
                        <select name="curso" class="select_style" id="curso">
                        <option value="999990" selected> 
                        </option>                        
                        </select>
                      </div>
                    </div></td>
                  <td width="200"> 
                    <div align="center"> 
                      <div id="divEtapa"> 
                        <select name="etapa" class="select_style" id="etapa">
                        <option value="999990" selected> 
                        </option>                          
                        </select>
                      </div>
                    </div></td>
                  <td width="200"> 
                    <div align="center"> 
                      <div id="divTurma"> 
                        <select name="turma" class="select_style" id="turma">
                        <option value="999990" selected> 
                        </option>                           
                        </select>
                      </div>
                    </div></td>
                  <td width="200"><div align="center">
                    <div id="divPeriodo">
                       <select name="periodo" class="select_style" id="periodo" onChange="GuardaPeriodo(this.value)">
                                      <option value="0" selected></option>
                                      <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")%>
                                      <option value="<%=NU_Periodo%>"> 
                                      <%response.Write(NO_Periodo)%>
                                      </option>
                                      <%RS4.MOVENEXT
WEND%>
                                    </select>
                    </div>
                  </div></td>
                </tr>
                <tr>
                  <td height="15" colspan="5" bgcolor="#FFFFFF"><hr></td>
                </tr>
                <tr> 
                  <td height="15" colspan="5" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="33%">&nbsp;</td>
                      <td width="34%">&nbsp;</td>
                      <td width="33%"><font size="2" face="Arial, Helvetica, sans-serif">
                        <div align="center">
                          <input name="Submit" type="submit" class="botao_prosseguir" id="Submit" value="Procurar">
                      </div>
                      </font></td>
                    </tr>
                  </table></td>
                </tr>
              </table>

    </td>
  </tr>
<%
	if display="list" then
	%>
                <tr>                   
    <td height="10" colspan="5" valign="top"> 
    <hr>
    </td>
  </tr>       
					<tr class="tb_corpo">                   
		<td height="10" colspan="5" class="tb_tit">Alunos Encontrados</td>
					</tr>
					<tr> 
					  
		<td colspan="5" valign="top">
         <ul> 
       <%	for i =0 to ubound(alunos_encontrados)
		cod_cons=alunos_encontrados(i)
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula="&cod_cons
		RS.Open SQL, CON1

		nome = RS("NO_Aluno")
		'ativo = RS("IN_Ativo_Escola")
		ativo= "True" 
			if ativo = "True" then
			Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=ativos href=avalia.asp?ori=ebe&opt=1&obr="&cod_cons&"$!$"&session("WS-DO-DPA-EFI-periodo")&">"&nome&"</a></font></li>")
			else
			Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=inativos href=avalia.asp?ori=ebe&opt=1&obr="&cod_cons&"$!$"&session("WS-DO-DPA-EFI-periodo")&">"&nome&"</a></font></li>")

			end if
		NEXT			
	%>
		  </ul>
          </td>
     </tr>  
<%
	END IF     
	
elseif display="ask1" then%>                      	  
 
      	  
          <tr class="tb_tit">             
      <td height="10" colspan="5">Preencha um dos campos abaixo</td>
          </tr>
          <TR>
      <td height="10" valign="top">      
            <div id="divTabela" style="visibility: hidden;">                  
<table width="1000" border="0" cellspacing="0">
         <TR>
      <td height="10" colspan="5" valign="top"> 

                <table width="1000" border="0" cellpadding="0" cellspacing="0">
          <tr>          
            <td width="150"  height="10"> 
              <div align="right"><font class="form_dado_texto"> 
                Matr&iacute;cula:</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
            </strong></font></div></td>
            
            <td width="50" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="busca1" type="text" class="textInput" id="busca1" value="<%response.Write(cod_cons)%>" size="12">
              </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font></td>
            
            <td width="150" height="10"> 
              <div align="right"><font class="form_dado_texto"> Nome: </font></div></td>
            
            <td width="500" height="10" ><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="busca2" type="text" class="textInput" id="busca2" size="55" maxlength="50">
              </font></td>
            
            <td width="150" height="10">&nbsp;</td>
          </tr>
		  </table>
		  </td>
		  </TR>
                <tr> 
                  <td width="200" class="tb_subtit"> 
                    <div align="center">UNIDADE 
                    </div></td>
                  <td width="200" class="tb_subtit"> 
                    <div align="center">CURSO 
                    </div></td>
                  <td width="200" class="tb_subtit"> 
                    <div align="center">ETAPA 
                    </div></td>
                  <td width="200" class="tb_subtit"> 
                    <div align="center">TURMA 
                    </div></td>
                  <td width="200" class="tb_subtit"> 
                  <div align="center">PER&Iacute;ODO</div></td>
                </tr>
                <tr> 
                  <td width="200"> 
                    <div align="center"> 
                      <select name="unidade" class="select_style" id="unidade" onChange="recuperarCurso(this.value)">
                        <option value="999990" selected></option>
                        <%		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0
While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")
%>
                        <option value="<%response.Write(NU_Unidade)%>"> 
                        <%response.Write(NO_Abr)%>
                        </option>
                        <%RS0.MOVENEXT
WEND
%>
                      </select>
                  </div></td>
                  <td width="200"> 
                    <div align="center"> 
                      <div id="divCurso"> 
                        <select name="curso" class="select_style" id="curso">
                        <option value="999990" selected> 
                        </option>                        
                        </select>
                      </div>
                  </div></td>
                  <td width="200"> 
                    <div align="center"> 
                      <div id="divEtapa"> 
                        <select name="etapa" class="select_style" id="etapa">
                        <option value="999990" selected> 
                        </option>                          
                        </select>
                      </div>
                  </div></td>
                  <td width="200"> 
                    <div align="center"> 
                      <div id="divTurma"> 
                        <select name="turma" class="select_style" id="turma">
                        <option value="999990" selected> 
                        </option>                           
                        </select>
                      </div>
                  </div></td>
                  <td width="200"> 
                    <div align="center"> 
                      <div id="divPeriodo"> 
<select name="periodo" class="select_style" id="periodo" onChange="GuardaPeriodo(this.value)">
                          <option value="0"></option>
                          <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo WHERE TP_Modelo='"&tp_modelo&"' order by NU_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")
NU_Periodo=NU_Periodo*1
periodo_form=periodo_form*1
if NU_Periodo=periodo_form then
%>
                          <option value="<%=NU_Periodo%>" selected> 
                          <%response.Write(NO_Periodo)%>
                          </option>
                          <%
else
%>
                          <option value="<%=NU_Periodo%>"> 
                          <%response.Write(NO_Periodo)%>
                          </option>
                          <%
end if
RS4.MOVENEXT
WEND%>
                        </select>
                      </div>
                  </div></td>
                    </tr>
                <tr> 
                  <td height="15" colspan="5" bgcolor="#FFFFFF"><hr></td>
                </tr>
                <tr>
                  <td height="15" colspan="5" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0">
                    <tr>
                      <td width="33%"><div align="center"></div></td>
                      <td width="34%"><div align="center"></div></td>
                      <td width="33%"><div align="center"><font size="3" face="Courier New, Courier, mono">
                        <input type="submit" name="Submit2" value="Prosseguir" class="botao_prosseguir">
                      </font></div></td>
                    </tr>
                  </table></td>
                </tr>
            </table>   

    </div>
    </td>
  </tr>                  
 
 
 
 
 
 
 
 
 
<%
	if display="list" then
	%>
                <tr>                   
    <td height="10" colspan="5" valign="top"> 
    <hr>
    </td>
  </tr>       
					<tr class="tb_corpo">                   
		<td height="10" colspan="5" class="tb_tit">Alunos Encontrados</td>
					</tr>
					<tr> 
					  
		<td colspan="5" valign="top">
         <ul> 
       <%	for i =0 to ubound(alunos_encontrados)
		cod_cons=alunos_encontrados(i)
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula="&cod_cons
		RS.Open SQL, CON1

		nome = RS("NO_Aluno")
		'ativo = RS("IN_Ativo_Escola")
		ativo= "True" 
			if ativo = "True" then
			Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=ativos href=avalia.asp?ori=ebe&opt=1&obr="&cod_cons&"$!$"&session("WS-DO-DPA-EFI-periodo")&">"&nome&"</a></font></li>")
			else
			Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=inativos href=avalia.asp?ori=ebe&opt=1&obr="&cod_cons&"$!$"&session("WS-DO-DPA-EFI-periodo")&">"&nome&"</a></font></li>")
			end if
		NEXT			
	%>
		  </ul>
          </td>
     </tr>  
<%
	END IF   	 

elseif display="ask2" then%>
      	  
          <tr class="tb_tit">             
      <td height="10" colspan="5">Preencha um dos campos abaixo</td>
          </tr>
          <TR>
      <td height="10" valign="top">      
            <div id="divTabela" style="visibility: hidden;">                  
<table width="1000" border="0" cellspacing="0">
         <TR>
      <td height="10" colspan="5" valign="top"> 

                <table width="1000" border="0" cellpadding="0" cellspacing="0">
          <tr>          
            <td width="150"  height="10"> 
              <div align="right"><font class="form_dado_texto"> 
                Matr&iacute;cula:</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
            </strong></font></div></td>
            
            <td width="50" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="busca1" type="text" class="textInput" id="busca1" size="12">
              </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
              </font></td>
            
            <td width="150" height="10"> 
              <div align="right"><font class="form_dado_texto"> Nome: </font></div></td>
            
            <td width="500" height="10" ><font size="2" face="Arial, Helvetica, sans-serif"> 
              <input name="busca2" type="text" class="textInput" id="busca2" size="55" maxlength="50">
              </font></td>
            
            <td width="150" height="10">&nbsp;</td>
          </tr>
		  </table>
		  </td>
		  </TR>
                <tr> 
                  <td width="200" class="tb_subtit"> 
                    <div align="center">UNIDADE 
                    </div></td>
                  <td width="200" class="tb_subtit"> 
                    <div align="center">CURSO 
                    </div></td>
                  <td width="200" class="tb_subtit"> 
                    <div align="center">ETAPA 
                    </div></td>
                  <td width="200" class="tb_subtit"> 
                    <div align="center">TURMA 
                    </div></td>
                  <td width="200" class="tb_subtit"> 
                  <div align="center">PER&Iacute;ODO</div></td>
                </tr>
                <tr> 
                  <td width="200"> 
                    <div align="center"> 
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
                  <td width="200"> 
                    <div align="center"> 
                      <div id="divCurso"> 
                        <select name="curso" class="select_style" onChange="recuperarEtapa(this.value)">
                            <option value="999990" selected>   
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
                  <td width="200"> 
                    <div align="center"> 
                      <div id="divEtapa"> 
                        <select name="etapa" class="select_style" onChange="recuperarTurma(this.value)">
                            <option value="999990" selected>                           
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
                  <td width="200"> 
                    <div align="center"> 
                      <div id="divTurma"> 
                        <select name="turma" class="select_style" onChange="recuperarPeriodo()">
                            <option value="999990" selected>                         
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
                  <td width="200"> 
                    <div align="center"> 
                      <div id="divPeriodo"> 
                        <select name="periodo" class="select_style" id="periodo" onChange="GuardaPeriodo(this.value)">
                          <option value="0"></option>
                          <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo WHERE TP_Modelo='"&tp_modelo&"' order by NU_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")
NU_Periodo=NU_Periodo*1
periodo_form=periodo_form*1
if NU_Periodo=periodo_form then
%>
                          <option value="<%=NU_Periodo%>" selected> 
                          <%response.Write(NO_Periodo)%>
                          </option>
                          <%
else
%>
                          <option value="<%=NU_Periodo%>"> 
                          <%response.Write(NO_Periodo)%>
                          </option>
                          <%
end if
RS4.MOVENEXT
WEND%>
                        </select>
                      </div>
                  </div></td>
                    </tr>
                <tr> 
                  <td height="15" colspan="5" bgcolor="#FFFFFF"><hr></td>
                </tr>
                <tr>
                  <td height="15" colspan="5" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0">
                    <tr>
                      <td width="33%"><div align="center"></div></td>
                      <td width="34%"><div align="center"></div></td>
                      <td width="33%"><div align="center"><font size="3" face="Courier New, Courier, mono">
                        <input type="submit" name="Submit2" value="Prosseguir" class="botao_prosseguir">
                      </font></div></td>
                    </tr>
                  </table></td>
                </tr>
            </table>   

    </div>
    </td>
  </tr>                    


<%
elseif display="listall" then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos Order BY NO_Aluno"
		RS.Open SQL, CON1
%>

                <tr class="tb_corpo"> 
                  
    <td height="10" colspan="5" class="tb_tit">Lista de completa de Alunos</td>
                </tr>
                <tr> 
                  
    <td colspan="5" valign="top"> 
      <ul>
        <%
	WHile Not RS.EOF
	nome = RS("NO_Aluno")
	cod_cons = RS("CO_Matricula")
	ativo = RS("IN_Ativo_Escola")
		if ativo = "True" then
		Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=ativos href=avalia.asp?ori=ebe&opt=1&obr="&cod_cons&"$!$"&session("WS-DO-DPA-EFI-periodo")&">"&nome&"</a></font></li>")
		else
		Response.Write("<li><font size=2 face=Arial, Helvetica, sans-serif><a class=inativos href=avalia.asp?ori=ebe&opt=1&obr="&cod_cons&"$!$"&session("WS-DO-DPA-EFI-periodo")&">"&nome&"</a></font></li>")
		end if
	RS.Movenext
	Wend
%>
      </ul></td>
                </tr>
<%end if %> 

                    
                <tr>         
    <div id="carregando"  align="center" style="position:absolute;  top: 200px; width:1000px; z-index: 4; height: 150px; visibility: hidden;">
				  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="75" height="75" vspace="80" title="Carregando">
				    <param name="movie" value="../../../../img/carregando.swf">
				    <param name="quality" value="high">
				    <param name="wmode" value="transparent">
				    <embed src="../../../../img/carregando.swf" width="75" height="75" vspace="80" quality="high" wmode="transparent" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash"></embed>
			      </object>
			 </div>              
				<div id="carregando_fundo" align="center" style="position:absolute; width:1000px; z-index: 3; height: 150px; visibility: hidden; background-color:#FFF; top: 250px;  filter: Alpha(Opacity=90, FinishOpacity=100, Style=0, StartX=0, StartY=100, FinishX=100, FinishY=100);">
			 </div>   
  </tr>   
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
</form>
</body>
<script type="text/javascript">
<!--
  initInputHighlightScript();
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