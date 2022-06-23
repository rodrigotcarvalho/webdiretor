<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<% 

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&nvg&"-"&ano_letivo
obr_acc=Session("obr_AAC")
incl_acc=Session("incl_AAC")
Session("obr_AAC")=obr_acc
Session("incl_AAC")=incl_acc
opt=request.QueryString("opt")
historico=request.QueryString("cod")
res=request.QueryString("res")
dados_historico = 	split(historico,"$!$")

ano_hist = dados_historico(0)
seq_hist = dados_historico(1)
matric_hist	= dados_historico(2)


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON7 = Server.CreateObject("ADODB.Connection") 
		ABRIR7 = "DBQ="& CAMINHO_h & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON7.Open ABRIR7		
		
call VerificaAcesso (CON,chave,nivel)
autoriza=Session("autoriza")

call navegacao (CON,chave,nivel)
navega=Session("caminho")
%>
<html>
<head>
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<link rel="stylesheet" href="../../../../jquery-ui.css" />                        
  <script src="http://code.jquery.com/jquery-1.9.1.js"></script>
  <script src="http://code.jquery.com/ui/1.10.3/jquery-ui.js"></script>

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
<% ano_atual = DATEPART("yyyy", now) %>
<script language="JavaScript" type="text/JavaScript">
<!--
function checksubmit()
{


	if (document.busca.ano_hist_form.value == "")
  {    alert("Por favor digite um ano letivo") 
    document.busca.ano_hist_form.focus()
    return false
  }
  
	if (document.busca.ano_hist_form.value < 1970 ||document.busca.ano_hist_form.value > <%response.Write(ano_atual)%> || isNaN(document.busca.ano_hist_form.value))
  {    alert("Por favor digite um ano letivo válido") 
    document.busca.ano_hist_form.focus()
    return false
  }  
  return true
}

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
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
						
						
						 function recuperarSegmento(tTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=s", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                    var resultado_t  = oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divSegmento.innerHTML =resultado_t

                                                           }
                                               }

                                               oHTTPRequest.send("t_pub=" + tTipo);
                                   }

								function recuperarEstabelecimento(tTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=est", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                    var resultado_t  = oHTTPRequest.responseText;
resultado_est = resultado_t.replace(/\+/g," ")
recuperarEstabelecimento = unescape(resultado_est)


                                                           }
                                               }

                                               oHTTPRequest.send("t_pub=" + tTipo);
                                   }
								   
function carregaEsquema(ETipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=esq", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                    var resultado_esq  = oHTTPRequest.responseText;
resultado_esq = resultado_esq.replace(/\+/g," ")
resultado_esq = unescape(resultado_esq)
document.all.divEsquema.innerHTML =resultado_esq

                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + ETipo);
                                   }								   
function gravaEsquema(Nome,Modelo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=gesq&inclui=S", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                    var resultado_gesq  = oHTTPRequest.responseText;
resultado_gesq = resultado_gesq.replace(/\+/g," ")
resultado_gesq = unescape(resultado_gesq)
document.all.divComboEsquema.innerHTML =resultado_gesq

                                                           }
                                               }

                                               oHTTPRequest.send("n_pub=" + Nome + "&m_pub=" + Modelo);
                                   }			
function excluiEsquema(Nome,Modelo)
                                   {
									   
	if(confirm('Confirma a exclusão do modelo selecionado')){

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=gesq&inclui=N", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                    var resultado_eesq  = oHTTPRequest.responseText;
resultado_eesq = resultado_eesq.replace(/\+/g," ")
resultado_eesq = unescape(resultado_eesq)
document.all.divComboEsquema.innerHTML =resultado_eesq

                                                           }
                                               }

                                               oHTTPRequest.send("n_pub=" + Nome + "&m_pub=" + Modelo);
	}
                                   }									   					   
 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
                        </script>

  <script>
<%	
 		Set RSE = Server.CreateObject("ADODB.Recordset")
		SQLE = "SELECT distinct NO_Materia FROM TB_Historico_Nota order by NO_Materia"
		RSE.Open SQLE, CON7
		
	vetor_materia=""
	conta_materia=0
	while not RSE.EOF
		nome_materia = replace(RSE("NO_Materia"),"""","'")
		if isnull(nome_materia) or nome_materia="" or nome_materia=" " or nome_materia="	" then
		
		else			
			if conta_materia = 0 then
				vetor_materia=""""&nome_materia&""""
			else
				vetor_materia=vetor_materia&","""&nome_materia&""""		
			end if
			conta_materia = conta_materia+1	
		end if				
	RSE.MOVENEXT
	WEND		

  		Set RSE = Server.CreateObject("ADODB.Recordset")
		SQLE = "SELECT distinct NO_Escola FROM TB_Historico_Ano order by NO_Escola"
		RSE.Open SQLE, CON7
		
	vetor_estabelecimento=""
	conta_estabelecimentos=0
	while not RSE.EOF
		nome_estabelecimento = replace(RSE("NO_Escola"),"""","'")
		if isnull(nome_estabelecimento) or nome_estabelecimento="" or nome_estabelecimento=" " or nome_estabelecimento="	"  then
		
		else			
			if conta_estabelecimentos = 0 then
				vetor_estabelecimento=""""&nome_estabelecimento&""""
			else
				vetor_estabelecimento=vetor_estabelecimento&","""&nome_estabelecimento&""""		
			end if
		conta_estabelecimentos = conta_estabelecimentos+1	
		end if			
	RSE.MOVENEXT
	WEND	
	
  		Set RSM = Server.CreateObject("ADODB.Recordset")
		SQLM = "SELECT distinct NO_Municipio FROM TB_Historico_Ano order by NO_Municipio"
		RSM.Open SQLM, CON7
		
	vetor_municipios=""
	conta_municipios=0
	while not RSM.EOF
		nome_municipio = replace(RSM("NO_Municipio"),"""","'")
		if isnull(nome_municipio) or nome_municipio="" or nome_municipio=" " or nome_municipio="	" then
		
		else	
			if conta_municipios = 0 then
				vetor_municipios=""""&nome_municipio&""""
			else
				vetor_municipios=vetor_municipios&","""&nome_municipio&""""		
			end if
			conta_municipios = conta_municipios+1				
		end if					
	RSM.MOVENEXT
	WEND		
%>
  $(function() {
    var estabelecimentos = [<%response.Write(vetor_estabelecimento)%>];
    var municipios = [<%response.Write(vetor_municipios)%>];		
    var materias = [<%response.Write(vetor_materia)%>];	

    $( "#estabelecimento_form" ).autocomplete({
      source: estabelecimentos
    });
	
    $( "#municipio_form").autocomplete({
      source: municipios
    });	
	
	$( "#disciplina_1" ).autocomplete({
      source: materias
    });	
								 
  });
function putFocusOn(campo)
{
	  var focal = document.getElementById(campo);
	focal.focus(); 
}


 
function changeImage(img){

	document.getElementById(img).innerHTML ='<a id="close_'+img+'" href="#" class="remove"><img src="../../../../img/close.png" alt="Excluir Item" width="20" height="20" border = "0" ></a>';

}   
function verifica_habilitacao(id, valor, alvo){
  if (id="carrega_modelo" && valor=="nulo"){
	  document.getElementById(alvo).disabled   = true
  } else {
	  document.getElementById(alvo).disabled   = false	  
  };	
}
function existeNomeEsquema(nomeEsquema) {
var itens;

var s = document.getElementById('carrega_modelo');

	for (index = 0; index < s.options.length; index++) {
		itens = s.options[index].value;			
		itens = itens.replace("SUPASUP", "ª");
		itens = itens.replace("SUPOSUP", "º");		
		itens = itens.replace("23A23", "Á");
		itens = itens.replace("23E23", "É");
		itens = itens.replace("23I23", "Í");
		itens = itens.replace("23O23", "Ó");
		itens = itens.replace("23U23", "Ú");
		itens = itens.replace("23C23", "Ç");
		itens = itens.replace("45A45", "Ã");
		itens = itens.replace("45N45", "Ñ");
		itens = itens.replace("45O45", "Õ");
		itens = itens.replace("78A78", "Â");
		itens = itens.replace("78E78", "Ê");
		itens = itens.replace("78O78", "Ô");
		itens = itens.toUpperCase();
		nomeEsquema = nomeEsquema.toUpperCase()
		if (nomeEsquema == itens) {
				return true;
		}
	} 
		return false;
}
	
function capturaEsquema(nomeEsquema) {
var container, inputs, index, vetorEsquema, contador, disciplina;

if (existeNomeEsquema(nomeEsquema)){
	if (!confirm('Já existe modelo com esse nome. Ao prosseguir ele será atualizado. Deseja continuar?')) { 	
		return false;
	}	
}
contador=0;
// Get the container element
container = document.getElementById('divEsquema');

// Find its child `input` elements
inputs = container.getElementsByTagName('input');
	for (index = 0; index < inputs.length; index++) {
		if (inputs[index].id.substring(0,11)=="disciplina_"){		
			contador++;
			disciplina = escape(inputs[index].value)
			if (contador==1){				
				vetorEsquema = disciplina;			
			} else {						
				vetorEsquema = vetorEsquema+','+disciplina;								
			}
		}
	}
	//alert(vetorEsquema);
	gravaEsquema(nomeEsquema,vetorEsquema);
	alert("Modelo gravado com sucesso!");	
}
  </script>  
<script>

      $(document).ready(function(){  
        $(document).on('click', 'a.add', function(){ 
		  var i = $("#itens_criados").val();
		  var j = $("#qtd_itens").val();	 	  
		  i++; 
		  j++; 
		  $("#itens_criados").val(i);		
		  $("#qtd_itens").val(j);			  		  
		  var listaMaterias= [<%response.Write(vetor_materia)%>];	  
          var row = "<tr><td width='80' height='25' align='right' class='form_corpo'><input name='num_linha' type='hidden' id='num_linha' value='"+i+"'>Disciplina:&nbsp;</td><td width='380' class='form_corpo'><span class='ui-widget'><input name='disciplina_"+i+"' class='textInput' type='text' id='disciplina_"+i+"' size='50'></span></td><td width='100' align='right' class='form_corpo'>Carga-hor&aacute;ria:&nbsp;</td><td width='40' class='form_corpo'><input name='carga_form_"+i+"' type='text' class='textInput' id='carga_form_"+i+"' size='6' maxlength='4'></td><td width='80' align='right' class='form_corpo'>Frequ&ecirc;ncia:&nbsp;</td><td width='40' class='form_corpo'><input name='frequencia_form_"+i+"' type='text' class='textInput' id='frequencia_form_"+i+"' size='6' maxlength='4'></td><td width='50' align='right' class='form_corpo'>Nota:&nbsp;</td><td width='40' class='form_corpo'><input name='nota_form_"+i+"' type='text' class='textInput' id='nota_form_"+i+"' size='6' maxlength='4'></td><td width='80' align='right' class='form_corpo'>Aprovado:&nbsp;</td><td width='60' class='form_corpo'><select name='aprovado_"+i+"' class='select_style'><option value='S' selected='selected'>Sim</option><option value='N'>N&atilde;o</option></select></td><td width='50' align='left' class='form_corpo'><div id='"+i+"'><a href='javascript:void(0);' class='add'><img src='../../../../img/add.png' width='20' height='20'  alt='Adicionar Disciplina' border='0'></a></div></td></tr>"	  	  
          $(this).closest('table').append(row);
		  $("#disciplina_"+i+"").autocomplete({source:listaMaterias});
	  	  changeImage(i-1);  
		  $("#disciplina_"+i+"").focus();
        });  
        $(document).on('click', 'a.remove', function(){  
          $(this).closest('tr').remove();  
		  var i = $("#itens_criados").val();		  
		  var j = $("#qtd_itens").val();
		  j--;  
		  $("#qtd_itens").val(j);
		  if (i>1){
		  	$("#disciplina_"+i+"").focus();		 
		  }
        });  
      });  
	  
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
	  
    </script>                         
</head>
<% 
if display<>"list" then
	'onload="onLoad=MM_callJS('document.busca.busca1.focus()')"
end if%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" <%response.Write(onload)%>>

<%call cabecalho(nivel)
%>
<table width="1000" height="670" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr>             
    <td width="1000" height="10" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
    </td>
          </tr>
<%if res="ok" then
if opt="inc" then
	num_msg = 418
else
	num_msg = 419
end if	
%>          
            <tr> 
              
    <td height="10" valign="top"> 
      <%call mensagens(nivel,num_msg,2,0) %>
    </td>
			  </tr>	
<%end if%>              
            <tr> 
              
    <td height="10" valign="top"> 
      <%call mensagens(nivel,300,0,0) %>
    </td>
			  </tr>	
         <form action="bd.asp?opt=<%=opt%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()">    
          <tr class="tb_tit">             
      <td height="10">Aluno</td>
          </tr>  
          <tr>
<td height="10" valign="top" class="form_dado_texto"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="form_dado_texto">
          <tr>
<%
	Set RSs = Server.CreateObject("ADODB.Recordset")
	SQL_s ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& matric_hist&"  and TB_Matriculas.NU_Ano in (SELECT max(NU_Ano) FROM TB_Matriculas where CO_Matricula ="& matric_hist&")"	
	Set RSs = CON1.Execute(SQL_s)

	nome_aluno=RSs("NO_Aluno")
%>          
            <td width="170" height="22" align="right" class="form_corpo"><input name="dados_historico" type="hidden" id="dados_historico" value="<%response.Write(historico)%>">
            Matrícula:&nbsp;</td>
            <td width="150" height="22" class="form_dado_texto"><%response.Write(matric_hist)%></td>   
            <td width="100" height="22" align="right" class="form_corpo">Nome:&nbsp;</td>
            <td width="580" height="22" class="form_dado_texto"><%response.Write(nome_aluno)%></td>    
          </tr>
          <tr>
            <td width="170" height="22" align="right" class="form_corpo">&nbsp;</td>
            <td width="150" height="22" valign="top" class="form_dado_texto">&nbsp;</td>   
            <td width="100" height="22" align="right" class="form_corpo">&nbsp;</td>
            <td width="580" height="22" valign="top" class="form_dado_texto">&nbsp;</td>    
          </tr>           </table>  </td></tr>                              	  
          <tr class="tb_tit">             
      <td height="10">Resumo</td>
          </tr>
          <TR>
      <td height="10" valign="top" class="form_dado_texto"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="form_dado_texto">
          <tr>
<%
if opt = "alt" then
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Historico_Ano where CO_Matricula = "& matric_hist &" AND DA_Ano = "&ano_hist&" AND NU_Seq = "&seq_hist
	RS.Open SQL, CON7	
	
	if RS.EOF then
		ano_hist = ""
		seq_hist = ""	
	else
		tipo_curso = RS("TP_Curso")
		co_seg = RS("CO_Seg")
		no_escola = RS("NO_Escola")
		pais_escola = RS("NO_Pais")	
		uf_escola = RS("SG_UF")	
		cidade_escola = RS("NO_Municipio")		
		situac_hist = RS("IN_Aprovado")
		obs_hist = RS("TX_Observacoes")
		tp_reg_hist = RS("TP_Registro")
		data_hist = RS("DT_Registro")
		ch_total = RS("NU_Carga_Horaria_Total")
		fq_total = RS("TX_Frequencia_Total")
	end if
else
	ano_hist = ""
	seq_hist = ""
	pais_escola = "Brasil"	
	uf_escola = "RJ"
	cidade_escola = "Rio de Janeiro"
	situac_hist = 1	
end if
%>              
            <td width="170" height="25" align="right" class="form_corpo">Ano Letivo:&nbsp;</td>
            <td width="70" height="25" valign="top">
            <input name="ano_hist_form" type="text" id="ano_hist_form" class="textInput" size="6" maxlength="4" value="<%response.Write(ano_hist)%>"></td>
            <td width="70" height="25" align="right" class="form_corpo">Curso:&nbsp;</td>
            <td width="240" height="25" valign="top"><select name="tipo_curso" class="select_style" id="tipo_curso" onChange="recuperarSegmento(this.value)">  
<option value="nulo" selected></option>                  
                        <%		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Tipo_Curso order by NU_Ordem"
		RS0.Open SQL0, CON7
		
While not RS0.EOF
	tipo_curso_bd = RS0("TC_Curso")
	no_abrv_curso = RS0("NO_Curso")
	if tipo_curso_bd = tipo_curso then
		selected_curso = "selected"
	else
		selected_curso=""	
	end if
%>
                        <option value="<%response.Write(tipo_curso_bd)%>" <%response.Write(selected_curso)%>> 
                        <%response.Write(no_abrv_curso)%>
                        </option>
                        <%RS0.MOVENEXT
WEND
%>                        
                        </select></td>
            <td width="130" height="25" align="right" class="form_corpo">Etapa:&nbsp;</td>
            <td height="25" colspan="3" valign="top"><div id="divSegmento">
            <select name="co_seg"  class="select_style" id="co_seg">
<%if opt = "alt" then                   
				Set RS0 = Server.CreateObject("ADODB.Recordset")
				SQL0 = "SELECT * FROM TB_Segmento where TP_Curso='"&tipo_curso&"' order by NU_Ordem"
				RS0.Open SQL0, CON7
			
			
				While not RS0.EOF
				
				co_seqmento = RS0("CO_Seg")		
				no_seqmento = RS0("NO_Abreviado_Curso")	
				if co_seg = co_seqmento then
					selected_seg = "selected"
				else
					selected_seg=""	
				end if								
				%>
                <option value="<%response.Write(co_seqmento)%>" <%response.Write(selected_seg)%>> 
                <%response.Write(no_seqmento)%>
                </option>
                <%
				RS0.MOVENEXT
				WEND
End if				                   
 %>
            </select></div></td>
          </tr>
          <tr>
            <td width="170" height="25" align="right" class="form_corpo">Estabelecimento de Ensino:&nbsp;</td>
            <td height="25" colspan="3" valign="top">
            <div class="ui-widget">
 			<input name="estabelecimento_form" class="textInput" type="text" id="estabelecimento_form" size="50" value="<%response.Write(no_escola)%>">  
</div>      </td>
            <td width="130" height="25" align="right" class="form_corpo">Pa&iacute;s:&nbsp;</td>
            <td width="150" height="25">
            <input name="pais_form" type="text" class="textInput" id="pais_form" value="<%response.Write(pais_escola)%>"></td>
            <td width="50" height="25" align="right"><span class="form_corpo">UF:&nbsp;</span></td>
            <td width="120" height="25" valign="top"><select name="uf_form" class="select_style">
            
            <% if uf_escola = "" or isnull(uf_escola) then
					selected_uf = "selected"
				end if
			%>
              <option value="nulo" <%response.Write(selected_uf)%>></option>		
                <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
	SG_UF= RS2("SG_UF")
	NO_UF= RS2("NO_UF")
	if SG_UF = uf_escola then
%>
                <option value="<%=SG_UF%>" selected>
                  <% =SG_UF%>
                </option>
              <%else%>
              <option value="<%=SG_UF%>">
                <% =SG_UF%>
                </option>
              <%end if						
RS2.MOVENEXT
WEND
%>
            </select></td>
          </tr>
          <tr>
            <td width="170" height="25" align="right" class="form_corpo">Munic&iacute;pio:&nbsp;</td>
            <td height="25" colspan="3" valign="top">
            <div class="ui-widget">
            <input name="municipio_form" class="textInput" type="text" id="municipio_form" size="50" value="<%response.Write(cidade_escola)%>">
            </div>              
            </td>
            <td width="130" height="25" align="right" class="form_corpo">Resultado Final:</td>
            <td height="25" colspan="3" valign="top"><select name="resultado_final_form" class="select_style" id="resultado_final_form">
            
            <% if situac_hist = "" or isnull(situac_hist) then
					selected_res_fin = "selected"
				end if
			%>
              <option value="nulo" <%response.Write(selected_res_fin)%>></option>
              <%		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Resultado_Final order by TP_Resultado"
		RS0.Open SQL0, CON7
		
While not RS0.EOF
	tp_resultado = RS0("TP_Resultado")
	no_resultado = RS0("NO_Resultado")
	situac_hist = situac_hist*1
	tp_resultado=tp_resultado*1
	if situac_hist = tp_resultado then
		selected_situac = "selected"
	else
		selected_situac=""	
	end if		
%>
              <option value="<%response.Write(tp_resultado)%>" <%response.Write(selected_situac)%>>
                <%response.Write(no_resultado)%>
                </option>
              <%RS0.MOVENEXT
WEND
%>
            </select></td>
          </tr>
          <tr>
            <td width="170" height="25" align="right" class="form_corpo">Observa&ccedil;&otilde;es:&nbsp;</td>
            <td height="70" colspan="7" rowspan="2" valign="top">
              <textarea name="observacoes_form" cols="160" rows="6" id="observacoes_form"><%response.Write(obs_hist)%></textarea>
            </td>
          </tr>
          <tr>
            <td width="170" height="70">&nbsp;</td>
          </tr>
        </table></td>
		  </TR>
           <tr>
             <td height="10" class="tb_tit">Disciplinas</td>                   
           </tr>  
                 <tr>
                  <td height="40" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="17%" height="25" align="right"><span class="form_corpo">Carga Hor&aacute;ria Total:&nbsp;</span></td>
                      <td height="25" colspan="2">
                      <input name="carga_total_form" type="text" class="textInput" id="carga_total_form" size="6" maxlength="4" value="<%response.Write(ch_total)%>"></td>
                      <td width="16%" height="25" align="right"><span class="form_corpo">Frequ&ecirc;ncia Total:&nbsp;</span></td>
                      <td width="24%" height="25">
                        <input name="frequencia_total_form" type="text" class="textInput" id="frequencia_total_form" size="6" maxlength="4" value="<%response.Write(fq_total)%>">
                      </td>
                    </tr>
                    <tr>
                      <td width="17%" height="25" align="right"><span class="form_corpo">Carrega Campos do Modelo:&nbsp;</span></td>
                      <td width="37%" height="25"><div id="divComboEsquema"><select name="carrega_modelo" class="select_style" id="carrega_modelo" onChange="javascript:verifica_habilitacao(this.id,this.value,'bt_carrega_esquema');javascript:verifica_habilitacao(this.id,this.value,'bt_exclui_esquema');">
                        <option value="nulo" selected></option>
                        <%		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT NO_Esquema FROM TB_Historico_Esquema_Disciplinas group by NO_Esquema order by NO_Esquema"
		RS0.Open SQL0, CON7
		
While not RS0.EOF
no_esquema = RS0("NO_Esquema")
modelo = ucase(no_esquema)
modelo = replace(modelo, "ª","SUPASUP")
modelo = replace(modelo, "º","SUPOSUP")
modelo = replace(modelo, "Á","23A23")
modelo = replace(modelo, "É","23E23")
modelo = replace(modelo, "Í","23I23")
modelo = replace(modelo, "Ó","23O23")
modelo = replace(modelo, "Ú","23U23")
modelo = replace(modelo, "Ç","23C23")
modelo = replace(modelo, "Ã","45A45")
modelo = replace(modelo, "Ñ","45N45")
modelo = replace(modelo, "Õ","45O45")
modelo = replace(modelo, "Â","78A78")
modelo = replace(modelo, "Ê","78E78")
modelo = replace(modelo, "Ô","78O78")
%>
                        <option value="<%response.Write(modelo)%>">
                          <%response.Write(no_esquema)%>
                        </option>
                        <%RS0.MOVENEXT
WEND
%>
                      </select>&nbsp;&nbsp;
                      <input name="bt_carrega_esquema" type="button" class="botao_prosseguir" id="bt_carrega_esquema" value="Carregar" disabled onClick="carregaEsquema(carrega_modelo.value)">
                      <input name="bt_exclui_esquema" type="button" class="botao_excluir" id="bt_exclui_esquema" value="Excluir" disabled onClick="excluiEsquema(carrega_modelo.value,'nulo')">
                      </div></td>
                      <td height="25" colspan="2" align="right"><span class="form_corpo">Grava um novo modelo com o nome:&nbsp;</span></td>
                      <td width="24%" height="25"><input name="nome_novo_esquema" class="textInput" type="text" id="nome_novo_esquema" size="15" maxlength="10">
                        &nbsp;&nbsp;
                        <input name="bt_grava_esquema" type="button" class="botao_prosseguir" id="bt_grava_esquema" value="Salvar" onClick="capturaEsquema(nome_novo_esquema.value)">
                      </span></td>
                    </tr>
                    <tr>
                      <td colspan="5" align="right"><hr></td>
                    </tr>
                    </table>
                 </td>
                </tr>
                <tr>
                  <td valign="top"><div id="divEsquema"><table width="1000" border="0" cellspacing="0" cellpadding="0">
  <tr class="form_corpo">
    <td height="25" align="right"><table id="tblInnerHTML" width="100%" border="0" cellspacing="0" cellpadding="0">
    <%if opt="inc" then
		total_registros = 1
	%>
      <tr>
        <td width="80" height="25" align="right" class="form_corpo"><input name="num_linha" type="hidden" id="num_linha" value="1">
          Disciplina:&nbsp;</td>
        <td width="380" class="form_corpo"><span class="ui-widget">
          <input name="disciplina_1" class="textInput" type="text" id="disciplina_1" size="50">
        </span></td>
        <td width="100" align="right" class="form_corpo">Carga-hor&aacute;ria:&nbsp;</td>
        <td width="40" class="form_corpo"><input name="carga_form_1" type="text" class="textInput" id="carga_form_1" size="6" maxlength="4"></td>
        <td width="80" align="right" class="form_corpo">Frequ&ecirc;ncia:&nbsp;</td>
        <td width="40" class="form_corpo"><input name="frequencia_form_1" type="text" class="textInput" id="frequencia_form_1" size="6" maxlength="4"></td>
        <td width="50" align="right" class="form_corpo">Nota:&nbsp;</td>
        <td width="40" class="form_corpo"><input name="nota_form_1" type="text" class="textInput" id="nota_form_1" size="6" maxlength="4"></td>
        <td width="80" align="right" class="form_corpo">Aprovado:&nbsp;</td>
        <td width="60" class="form_corpo"><select name="aprovado_1" class="select_style">
          <option value="S" selected="selected">Sim</option>
          <option value="N">N&atilde;o</option>
        </select></td>
        <td width="50" align="left" class="form_corpo"><div id="1"><a href="javascript:void(0);"  class="add" ><img src="../../../../img/add.png" width="20" height="20"  alt="Adicionar Disciplina" border="0"></a></div></td>
        </tr>
       <%else
	   
			Set RSL = Server.CreateObject("ADODB.Recordset")
			SQLL = "SELECT * FROM TB_Historico_Nota where CO_Matricula = "& matric_hist &" AND DA_Ano = "&ano_hist&" AND NU_Seq = "&seq_hist
			RSL.Open SQLL, CON7	
			
			if RSL.EOF then
				ano_hist = ""
				seq_hist = ""	
				vetor_historico = "1$!$"&indice&"#!##!##!##!##!#"
				indice = 1					
			else
				indice = 0
				while not RSL.EOF
					indice = indice+1				 
					hist_disciplina = RSL("NO_Materia")
					hist_carga = RSL("NU_Carga_Horaria")
					hist_freq = RSL("TX_Frequencia")
					hist_nota = RSL("VA_Nota")
					hist_apr = RSL("IN_Aprovado")
					if hist_apr = TRUE then
						hist_apr = "S"
					else
						hist_apr = "N"						
					end if
					
					if indice=1 then
						vetor_historico = indice&"#!#"&hist_disciplina&"#!#"&hist_carga&"#!#"&hist_freq&"#!#"&hist_nota&"#!#"&hist_apr
					else
						vetor_historico = vetor_historico&"$!$"&indice&"#!#"&hist_disciplina&"#!#"&hist_carga&"#!#"&hist_freq&"#!#"&hist_nota&"#!#"&hist_apr
					end if
				RSL.MOVENEXT
				WEND	
				total_registros = indice
			end if	  
			dados_historico = split(vetor_historico, "$!$") 
	   		FOR h=0 to ubound(dados_historico)
				dados_disciplina = split(dados_historico(h), "#!#") 	
				linha = dados_disciplina(0)	
				nome_disc = dados_disciplina(1)	
				carga_disc = dados_disciplina(2)	
				freq_disc = dados_disciplina(3)	
				nota_disc = dados_disciplina(4)	
				apr_disc = dados_disciplina(5)	
				h=h*1
				validador=h+1
				indice = indice*1
				if validador <> indice then
					img_bt="close.png"  	
					alt="Excluir Disciplina"	
					classe = "remove"	
				else 
					img_bt="add.png"  
					alt="Adicionar Disciplina"
					classe = "add"	
				end if								
	   %>
       
      <tr>
        <td width="80" height="25" align="right" class="form_corpo"><input name="num_linha" type="hidden" id="num_linha" value="<%response.Write(linha)%>">
          Disciplina:&nbsp;</td>
        <td width="380" class="form_corpo"><span class="ui-widget">
          <input name="disciplina_<%response.Write(linha)%>" class="textInput" type="text" id="disciplina_<%response.Write(linha)%>" value="<%response.Write(nome_disc)%>" size="50">
        </span></td>
        <td width="100" align="right" class="form_corpo">Carga-hor&aacute;ria:&nbsp;</td>
        <td width="40" class="form_corpo"><input name="carga_form_<%response.Write(linha)%>" type="text" class="textInput" id="carga_form_<%response.Write(linha)%>" value="<%response.Write(carga_disc)%>" size="6" maxlength="4"></td>
        <td width="80" align="right" class="form_corpo">Frequ&ecirc;ncia:&nbsp;</td>
        <td width="40" class="form_corpo"><input name="frequencia_form_<%response.Write(linha)%>" type="text" class="textInput" id="frequencia_form_<%response.Write(linha)%>" value="<%response.Write(freq_disc)%>" size="6" maxlength="4"></td>
        <td width="50" align="right" class="form_corpo">Nota:&nbsp;</td>
        <td width="40" class="form_corpo"><input name="nota_form_<%response.Write(linha)%>" type="text" class="textInput" id="nota_form_<%response.Write(linha)%>" value="<%response.Write(nota_disc)%>" size="6" maxlength="4"></td>
        <td width="80" align="right" class="form_corpo">Aprovado:&nbsp;</td>
        <td width="60" class="form_corpo"><select name="aprovado_<%response.Write(linha)%>" class="select_style">
        <%  if apr_disc="S" then
			 selected_apr = "selected"
			 selected_rpr = ""
			else
			 selected_apr = ""	
			 selected_rpr = "selected"		
			end if
		
		%>
          <option value="S" <%response.Write(selected_apr)%>>Sim</option>
          <option value="N" <%response.Write(selected_rpr)%>>N&atilde;o</option>
        </select></td>
        <td width="50" align="left" class="form_corpo"><div id="<%response.Write(linha)%>"><a href="javascript:void(0);"  class="<%response.Write(classe)%>" ><img src="../../../../img/<%response.Write(img_bt)%>" width="20" height="20"  alt="<%response.Write(alt)%>" border="0"></a></div></td>
        </tr>
       
       <%
	   		Next
	   end if%> 
    </table> </td>
  </tr>
  <tr class="form_corpo">
    <td width="1000" align="right"><hr></td>
  </tr>
  <tr class="form_corpo">
    <td height="45" align="right"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="20%"><input name="qtd_itens" type="hidden" id="qtd_itens" value="<%response.Write(total_registros)%>">
          <input name="itens_criados" type="hidden" id="itens_criados" value="<%response.Write(total_registros)%>"></td>
        <td width="20%" align="center"><input type="button" name="button" id="button" class="botao_cancelar" value="Cancelar" onClick="MM_goToURL('parent','resumo.asp?voltar=S');"/></td>
        <td width="20%">&nbsp;</td>
        <td width="20%" align="center"><input type="submit" name="Submit" id="button" class="botao_prosseguir" value="Salvar" /></td>
        <td width="20%"></td>
      </tr>
    </table></td>
  </tr>
                  </table></div> 
                 </td>
                </tr></form> 
                <tr>
                  <td valign="top">&nbsp;</td>
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