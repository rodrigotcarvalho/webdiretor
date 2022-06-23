<%	'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/utils.asp"-->
<%
opt = request.QueryString("opt")

ano_letivo_wf = session("ano_letivo_wf")
session("ano_letivo_wf")=ano_letivo_wf
co_usr = session("co_user")
nivel=4

Session("dia_de")=""
Session("dia_de")=""
Session("dia_ate")=""
Session("mes_ate")=""
Session("unidade")=""
Session("curso")=""
Session("etapa")=""
Session("turma")=""
Session("arquivos_desanexados")="nulo" 

if transicao = "S" then
 area="wd"
 url="http://simplynet2.tempsite.ws/wd/"&ambiente_escola&"/anexos/"
else	
	if left(ambiente_escola,5) = "teste" then
		url = "http://www.simplynet.com.br/wdteste/"&ambiente_escola&"/anexos/"
	else
		url = "http://www.webdiretor.com.br/"&ambiente_escola&"/anexos/"
	end if	
end if	
if isnull(Session("arquivos_anexados")) then

elseif (Session("arquivos_anexados")<>"nulo" and Session("arquivos_anexados")<>"") then


	SET FSO = Server.CreateObject("Scripting.FileSystemObject")
	
	Set pasta = FSO.GetFolder(CAMINHO_upload)
	Set arquivos = pasta.Files

	for each apagarquivo in arquivos
		data_arquivo =apagarquivo.DateLastModified
		nome_arquivo =apagarquivo.Name
		'response.Write(DatePart("n",Now())&"<BR>")	
		hora=DatePart("h",Now())
		min=DatePart("n",Now())
		hora_arquivo=DatePart("h",data_arquivo)
		min_arquivo=DatePart("n",data_arquivo)
		if (hora_arquivo<hora and min>30) or (min-min_arquivo>30) then
			FSO.deletefile(apagarquivo) 
		end if	
	next
	anexos=split(Session("arquivos_anexados"),"#!#")
	for atch=0 to ubound(anexos)	
		arquivo = CAMINHO_upload & anexos(atch)
		FSO.deletefile(arquivo) 
	Next	
	Session("arquivos_anexados")="nulo" 
end if	


nvg = request.QueryString("nvg")
chave=nvg
session("chave")=chave
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano_info=nivel&"-"&chave&"-"&ano_letivo_wf



		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1

    	Set CON_WF = Server.CreateObject("ADODB.Connection") 
		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_WF.Open ABRIR_WF	


 call navegacao (CON,chave,nivel)
navega=Session("caminho")	

 Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Autoriz_Usuario_Grupo Where CO_Usuario = "&co_usr
		RS2.Open SQL2, CON
		
if RS2.EOF then

else		
co_grupo=RS2("CO_Grupo")
End if
%>


<!DOCTYPE html>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
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
function submitforminterno()  
{	
   var f=document.forms[3]; 

   if (
	document.getElementById("dest1").checked == false &&
	document.getElementById("dest2").checked == false &&
	document.getElementById("dest3").checked == false)
	{

		alert ('Favor selecionar o tipo de destinatário');
		return false;
	} 
	//console.log('unidade');			
	if (document.getElementById("unidade").options[0].selected) {

		alert ('É necessário selecionar a unidade');
		return false;
	}
	//console.log('curso');	
	//if (document.getElementById("curso").options[0].selected) {
//		alert ('É necessário selecionar o curso');
//		return false;
//	}
	//console.log('etapa');	
	//if (document.getElementById("etapa").options[0].selected) {
//		alert ('É necessário selecionar a etapa');
//		return false;
//	}
	//console.log('turma');
	//if (document.getElementById("turma").options[0].selected) {
//		alert ('É necessário selecionar a turma');
//		return false;
//	}

   div2form('formulario');   
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
// A função abaixo pega a versão mais nova do xmlhttp do IE e verifica se é Firefox. Funciona nos dois.
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
// Criação do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicitação HTTP. O primeiro parâmetro informa o método post/get
// O segundo parâmetro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicitação síncrona, o parâmetro deve ser false
                                               oHTTPRequest.open("post", "executa.asp?opt=c", true);
// Para solicitações utilizando o método post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A função abaixo é executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto já completou a solicitação
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto é gerado no arquivo executa.asp e colocado no div
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
document.all.divEtapa.innerHTML ="<select name='etapa' id = 'etapa' class=select_style><option value='nulo' selected></option></select>"
document.all.divTurma.innerHTML = "<select name='turma' id = 'turma' class=select_style><option value='nulo' selected></option></select>"
//recuperarEtapa()
                                                           }
                                               }
// Abaixo é enviada a solicitação. Note que a configuração
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {
// Criação do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicitação HTTP. O primeiro parâmetro informa o método post/get
// O segundo parâmetro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicitação síncrona, o parâmetro deve ser false
                                               oHTTPRequest.open("post", "executa.asp?opt=e", true);
// Para solicitações utilizando o método post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A função abaixo é executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto já completou a solicitação
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto é gerado no arquivo executa.asp e colocado no div
                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select  name='turma' id = 'turma' class=select_style><option value='nulo' selected></option></select>"
//recuperarTurma()
                                                           }
                                               }
// Abaixo é enviada a solicitação. Note que a configuração
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {
// Criação do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicitação HTTP. O primeiro parâmetro informa o método post/get
// O segundo parâmetro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicitação síncrona, o parâmetro deve ser false
                                               oHTTPRequest.open("post", "executa.asp?opt=t", true);
// Para solicitações utilizando o método post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A função abaixo é executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto já completou a solicitação
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto é gerado no arquivo executa.asp e colocado no div
                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t																	   
                                                           }
                                               }
// Abaixo é enviada a solicitação. Note que a configuração
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
						 function recuperarMensagem(mTipo)
                                   {
// Criação do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicitação HTTP. O primeiro parâmetro informa o método post/get
// O segundo parâmetro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicitação síncrona, o parâmetro deve ser false
                                               oHTTPRequest.open("post", "executa.asp?opt=msg", true);
// Para solicitações utilizando o método post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A função abaixo é executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto já completou a solicitação
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto é gerado no arquivo executa.asp e colocado no div
                                                                       var resultado_M= oHTTPRequest.responseText;
resultado_M = resultado_M.replace(/\+/g," ")
resultado_M = unescape(resultado_M)
document.all.divMensagem.innerHTML = resultado_M																	   
                                                           }
                                               }
// Abaixo é enviada a solicitação. Note que a configuração
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("m_pub=" + mTipo);
                                   }								   
                        
	   
						 function desanexa(aTipo)
                                   {
// Criação do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicitação HTTP. O primeiro parâmetro informa o método post/get
// O segundo parâmetro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicitação síncrona, o parâmetro deve ser false
                                               oHTTPRequest.open("post", "executa.asp?opt=danx", true);
// Para solicitações utilizando o método post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A função abaixo é executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto já completou a solicitação
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto é gerado no arquivo executa.asp e colocado no div
//                                                                      var resultado_A= oHTTPRequest.responseText;
//resultado_A = resultado_A.replace(/\+/g," ")
//resultado_A = unescape(resultado_A)
//document.all.divAnexo.innerHTML = resultado_A																	   
                                                           }
                                               }
// Abaixo é enviada a solicitação. Note que a configuração
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("a_pub=" + aTipo);
                                   }
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
                         </script>
<script>
function allowDrop(ev) {
    ev.preventDefault();
}

function drag(ev) {
    ev.dataTransfer.setData("img", ev.target.id);
}

function drop(ev) {
    ev.preventDefault();
    var data = ev.dataTransfer.getData("img");
	var imagem = document.getElementById(data)
	var divMensagem = document.getElementById(divMensagem)	
	imagem.style.height = '100%';
	var imgSrc = imagem.src;	
	imagem.src = imgSrc.replace("../../../../anexos/","<%response.Write(url)%>");
    ev.target.appendChild(imagem);
}

function colarImagem(data) {
	
  	//var imagem = "".concat("<%response.Write(url)%>","'"+data+"'")
var imagem = document.createElement("img");
imagem.setAttribute("name", "imgAnexo");
imagem.setAttribute("src", "<%response.Write(url)%>"+data);	
	//alert("<%response.Write(url)%>"+data);
	var divMensagem = document.getElementById("divMensagem")	
	//imagem.style.height = '100%';
	//var imgSrc = imagem.src;	
	//imagem.src = imgSrc.replace("../../../../anexos/","<%'response.Write(url)%>");
    divMensagem.appendChild(imagem);
}
 function div2form(id){
        var form=document.getElementById(id);
//        if(!form){
//            return;
//        }
//        var divs=document.getElementsByName(id+'div')
//        var i, ndivs=divs.length;
//        for(i=0;i<ndivs;i++){
//            if(document.getElementById('textarea'+divs[i].id)){
//               document.getElementById('textarea'+divs[i].id).value=divs[i].innerHTML; 
//            } else {
//                var texta=document.createElement('TEXTAREA');
//                texta.name=divs[i].id;
//                texta.id='textarea'+divs[i].id;
//                texta.value=divs[i].innerHTML;      
//                texta.style.display='none';
//                form.appendChild(texta);                
//            }                  
//        }   
        	   var divs=document.getElementById("divMensagem")
               var texta=document.createElement("textarea");
                texta.name="msg";
                texta.id="msg";
                texta.value=divs.innerHTML;      
                texta.style.display='none';
                form.appendChild(texta);  
				    
    }
	
function anexaForm(id, arquivo,nome){
//guardar o total de anexos para encaminhar ao loop do programa de envio de email	
var qtd = document.getElementById("qtdAnexos").value;
qtd = qtd*1;
qtd = qtd+1;
document.getElementById("qtdAnexos").value = qtd;

// span que representa a linha da tabela
var span = document.createElement("span");
span.setAttribute("name", "spanAnexo"+qtd);
span.setAttribute("id", "spanAnexo"+qtd);
span.setAttribute("style", "style=display:inline-block");

// span que representa a primeira célular da tabela
var spanCell1 = document.createElement("span");
spanCell1.setAttribute("name", "spanCell1Anexo"+qtd);
spanCell1.setAttribute("id", "spanCell1Anexo"+qtd);
spanCell1.setAttribute("style", "display:table-cell;width:20px");

// span que representa a segunda célular da tabela
var spanCell2 = document.createElement("span");
spanCell2.setAttribute("name", "spanCell2Anexo"+qtd);
spanCell2.setAttribute("id", "spanCell2Anexo"+qtd);
spanCell2.setAttribute("style", "display:table-cell;vertical-align: middle;width:340px;font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 10px;color: #336699");

// span que representa a terceira célular da tabela
var spanCell3 = document.createElement("span");
spanCell3.setAttribute("name", "spanCell3Anexo"+qtd);
spanCell3.setAttribute("id", "spanCell3Anexo"+qtd);
spanCell3.setAttribute("style", "display:table-cell;width:20px");

// input que guarda o nome do arquivo que deve ser anexado
var input = document.createElement("input");
input.setAttribute("type", "hidden");
input.setAttribute("name", "anexo"+qtd);
input.setAttribute("value", arquivo);

// anexa a linha na div de anexos
document.getElementById("divAnexos").appendChild(span);
// anexa a imagem para excluir o anexo
var idSpan = "spanAnexo"+qtd

document.getElementById(idSpan).appendChild(spanCell1);
// anexa o input
document.getElementById("spanCell1Anexo"+qtd).innerHTML="<a href=# onClick=desAnexaForm('"+id+"','spanAnexo"+qtd+"')><img src=../../../../img/fecha.gif width=20 height=16 /></a>";
// anexa a segunda coluna
document.getElementById("spanAnexo"+qtd).appendChild(spanCell2);
// anexa o nome que é exibido para o usuário
document.getElementById("spanCell2Anexo"+qtd).innerHTML=nome;
// anexa a terceira coluna
document.getElementById("spanAnexo"+qtd).appendChild(spanCell3);
// anexa a primeira coluna
document.getElementById("spanCell3Anexo"+qtd).appendChild(input);


document.getElementById(id).setAttribute("style", "display:none");



}	
function desAnexaForm(id, anexo){
var qtd = document.getElementById("qtdAnexos").value;

document.getElementById(id).setAttribute("style", "display:inline-block;");

var d = document.getElementById(anexo)
d.parentNode.removeChild( d );

qtd = qtd*1;
document.getElementById("qtdAnexos").value = qtd-1;
}	

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
</script>
                        
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0" >
<% call cabecalho (nivel)
	  %>
<form name="formulario" id="formulario" METHOD="POST" ACTION="email.asp">      
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
                    
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
	  </td>
	  </tr>
<%if opt="ok" then%>      
      <tr>                
    <td height="10"> 
      <%	call mensagens(4,709,2,0) 
%>
	</td>
	</tr> 
<%elseif opt="ok2" then%>      
      <tr>                
    <td height="10"> 
      <%	call mensagens(4,52,2,0) 
%>
	</td>
	</tr>     
<%end if%>         
      <tr>                
    <td height="10"> 
      <%	call mensagens(4,9706,0,0) 
%>
	</td>
	</tr>
<tr>

    <td valign="top"> 
		<%
mes = DatePart("m", now) 
dia = DatePart("d", now) 



dia=dia*1
mes=mes*1
%>				
	  
<table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
        <tr class="tb_tit"> 
          <td width="653" height="15" class="tb_tit">Informe os crit&eacute;rios 
              para pesquisa 
            </td>
        </tr>
        <tr> 
          <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="338" valign="top"><table width="97%" border="0" align="right" cellpadding="0" cellspacing="0">
                      <tr>
                        <td class="tb_subtit">Informe o Assunto:<input name="co_grupo" type="hidden" id="co_grupo" value="<% = co_grupo %>"></td>
                      </tr>
                      <tr>
                        <td><select name="assunto" class="select_style_fixo_1" id="assunto" >
                          <%
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Email_Assunto order by CO_Assunto"
		RS1.Open SQL1, CON0

if RS1.eof then %>

<%else
	while not RS1.EOF
		co_assunto=RS1("CO_Assunto")
		assunto=RS1("TX_Titulo_Assunto")
		assunto_padrao=RS1("IN_Assunto_Padrao")
		
		if assunto_padrao=TRUE then
			assunto_selected="SELECTED"
		ELSE
			assunto_selected=""
		END IF	
		
		%>
							  <option value="<%response.Write(co_assunto)%>" <%response.Write(assunto_selected)%>>
								<%response.Write(assunto)%>
								</option>
							  <%

	RS1.movenext
	Wend
end if	



%>
                        </select></td>
                      </tr>
                      <tr>
                        <td height="5"></td>
                      </tr>
                      <tr>
                        <td><span class="tb_subtit">Com c&oacute;pia para:</span></td>
                      </tr>
                      <tr>
                        <td><select name="cc" class="select_style_fixo_1" id="cc" >
                          <%
'		Set RS1 = Server.CreateObject("ADODB.Recordset")
'		SQL1 = "SELECT Login FROM TB_Operador"
'		RS1.Open SQL1, CON
'qtd_cc=0
'while not RS1.EOF
'	email=RS1("Login")
'	
'	if qtd_cc=0then
'		cc_selected="SELECTED"
'		qtd_cc=1
'	ELSE
'		cc_selected=""
'	END IF	
'	
'	%>
<!--                          <option value="<%response.Write(email)%>" <%response.Write(cc_selected)%>>
                            <%response.Write(email)%>
                            </option>
-->                         <%
'RS1.movenext
'Wend

		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT DISTINCT(TX_email_WF) FROM TB_Etapa"
		RS2.Open SQL2, CON0


while not RS2.EOF
	email=RS2("TX_email_WF")
	
	%>
                          <option value="<%response.Write(email)%>">
                            <%response.Write(email)%>
                            </option>
                          <%
RS2.movenext
Wend
%>
                          </select></td>
                      </tr>
                      <tr>
                        <td>&nbsp;</td>
                      </tr>   
                      <tr>
                        <td class="tb_subtit">Arquivos disponíveis para envio</td>
                      </tr>
                      <tr>
                        <td>
                        <div class="form_dado_texto" id="arquivos" style="overflow-x: visible; overflow-y: auto;height:300px;">
						<%
hora = DatePart("h", now) 
min = DatePart("n", now)
seg= DatePart("s", now) 
								
' Cria uma objeto
Set objXMLDoc = Server.CreateObject("MSXML2.DOMDocument.6.0")
 
' Indicamos que o download em segundo plano não é permitido
objXMLDoc.async = False
 
objXMLDoc.load(Server.MapPath("..\..\..\..\anexos\anexos.xml"))

For Each xmlProduct In objXMLDoc.documentElement.selectNodes("arquivo")
     Dim arquivo : arquivo = xmlProduct.selectSingleNode("file").text   
     Dim nome : nome = xmlProduct.selectSingleNode("nome").text   
	 Dim formato : formato = xmlProduct.selectSingleNode("formato").text 
	 
	 nome = pontinhos(nome,35,"S")
	 if totalImgs mod 2 = 0 then
	 	bgColor = "#E8F2EE"
	else
	 	bgColor = "#FFFFFF"	 
	 end if	
	 %>
<span id="arquivo<% response.write(totalImgs)%>" style="display:inline-block;background-color:<%response.Write(bgColor)%>" >
<span style="display:table-cell;width:240px;vertical-align: middle;" ><%Response.Write Server.HTMLEncode(nome) %></span>
    <span style="display:table-cell;width:35px">
        <a href="#" onClick="anexaForm('arquivo<% response.write(totalImgs)%>','<%Response.Write (arquivo) %>','<%Response.Write Server.HTMLEncode(nome) %>')">
            <img src="../../../../img/icones/mail_attachment.png" width="25" height="25" alt=""/>
        </a>
    </span>
    <span style="display:table-cell;width:35px">
		<% if formato="IMG" then%>
        <a href="#" onClick="colarImagem('<%Response.Write (arquivo) %>')"><img src="../../../../img/icones/mail_open.png" width="25" height="25" alt="" /></a>
        <% end if %>
    </span>
</span>
<% totalImgs = totalImgs+1
				
Next 					
						%>
</div> 
<div id="btUpload"><input type="button" class="botao_prosseguir" value="Enviar arquivos" style="width:50%;display: block;margin-left: auto;margin-right: auto" onClick="MM_goToURL('parent','upload.asp');"></div>
</td>
                      </tr>                                                               
                    </table></td>
                    <td width="662" valign="top"><table width="100%" border="0" align="right" cellpadding="0" cellspacing="0">
                      <tr>
                        <td class="tb_subtit">&nbsp;Informe o conte&uacute;do da mensagem:</td>
                      </tr>
                      <tr>
                        <td><select name="mensagem" class="select_style_fixo_1" id="mensagem" onChange="recuperarMensagem(this.value)">
                          <%
		Set RS1b = Server.CreateObject("ADODB.Recordset")
		SQL1b = "SELECT * FROM TB_Email_Mensagem order by CO_Email"
		RS1b.Open SQL1b, CON0

if RS1b.EOF then

else
	while not RS1b.EOF
		co_email=RS1b("CO_Email")
		co_msg=RS1b("NO_Email")
		msg=RS1b("TX_Conteudo_Email")
		msg_padrao=RS1b("IN_Email_Padrao")
		
'		if co_msg<10 then
'			espacador="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
'		elseif co_msg<100 then
'			espacador="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"	
'		else	
			espacador="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"		
'		end if	
		msg_combo="Mensagem "&co_msg&espacador
		
		if msg_padrao=TRUE then
			msg_selected="SELECTED"
			conteudo_email=msg
		ELSE
			msg_selected=""
		END IF	
		
		%>
							  <option value="<%response.Write(co_email)%>" <%response.Write(msg_selected)%>>
								<%response.Write(msg_combo)%>
								</option>
							  <%
	
	RS1b.movenext
	Wend
end if	
%>
                        </select></td>
                      </tr>
                      <tr>
                        <td height="5"></td>
                      </tr>
                      <tr>
                        <td><div id="divMensagem" draggable="true" contenteditable="true" ondrop="drop(event)" ondragover="allowDrop(event)"><%response.write(replace(conteudo_email,chr(13),"<BR>"))%></div></td>
                      </tr>
                      <tr>
                        <td class="tb_subtit">Arquivos anexados</td>
                      </tr>
                      <tr>
                        <td><div id="divAnexos" style="overflow-x: visible; overflow-y: auto; height:200px;">
                          <input type="hidden" name="qtdAnexos" id="qtdAnexos" value="0">
                        </div></td>
                      </tr>                                            
                    </table></td>
                  </tr>
                  <tr>
                    <td colspan="2"></td>
                  </tr>
                </table></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td colspan="2">&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td colspan="4" class="tb_subtit"><table width="988" border="0" align="right" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="180" class="tb_subtit">Selecione os Destinat&aacute;rios:</td>
<!--                     <td width="25"><input name="dest" type="checkbox" class="borda" id="dest" value="a"></td>
                   <td width="60" class="form_dado_texto">Alunos</td>-->
                    <td width="25"><input name="dest" type="checkbox" class="borda" id="dest1" value="r"></td>
                    <td width="100" class="form_dado_texto">Respons&aacute;veis</td>
                    <td width="25"><input name="dest" type="checkbox" class="borda" id="dest2" value="i"></td>
                    <td width="65" class="form_dado_texto">Contatos</td>
                    <td width="24"><input name="dest" type="checkbox" class="borda" id="dest3" value="p" checked="checked"></td>
                    <td width="569" class="form_dado_texto">Professores</td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td width="247" class="tb_subtit"><div align="center">UNIDADE </div></td>
                <td width="247" class="tb_subtit"><div align="center">CURSO </div></td>
                <td width="247" class="tb_subtit"><div align="center">ETAPA </div></td>
                <td width="247" class="tb_subtit"><div align="center">TURMA </div></td>
              </tr>
              <tr>
                <td width="247"><div align="center">
                  <select name="unidade" id = "unidade" class="select_style" onChange="recuperarCurso(this.value)">
                    <option value="nulo" selected></option>
                    <%		

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0.Open SQL0, CON0
NU_Unidade_Check=999999		
While not RS0.EOF
NU_Unidade = RS0("NU_Unidade")
NO_Abr = RS0("NO_Abr")
if NU_Unidade = NU_Unidade_Check then
RS0.MOVENEXT		
else
%>
                    <option value="<%response.Write(NU_Unidade)%>">
                      <%response.Write(NO_Abr)%>
                      </option>
                    <%

NU_Unidade_Check = NU_Unidade
RS0.MOVENEXT
end if
WEND
%>
                  </select>
                </div></td>
                <td width="247"><div align="center">
                  <div id="divCurso">
                    <select name="curso" id="curso" class="select_style">
                    <option value="nulo" selected></option>                    
                    </select>
                  </div>
                </div></td>
                <td width="247"><div align="center">
                  <div id="divEtapa">
                    <select name="etapa" id="etapa" class="select_style">
                    <option value="nulo" selected></option>                    
                    </select>
                  </div>
                </div></td>
                <td width="247"><div align="center">
                  <div id="divTurma">
                    <select name="turma" id="turma" class="select_style">
                    <option value="nulo" selected></option>                    
                    </select>
                  </div>
                </div></td>
              </tr>
            </table>

          </td>
        </tr>
        <tr>
          <td valign="top">	<table width="988" border="0" align="right" cellpadding="0" cellspacing="0">
  <tr>
    <td>   
</td>
  </tr>
</table>

 </td>
        </tr>
        <tr>
          <td valign="top">&nbsp;</td>
        </tr>
        <tr>
          <td valign="top"><hr width="1000"></td>
        </tr>
        <tr>
          <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="33%">&nbsp;</td>
                <td width="34%">&nbsp;</td>
                <td width="33%"><div align="center"><input name="SUBMIT" type="BUTTON" class="botao_prosseguir" onClick="MM_callJS('submitforminterno()')" value="Enviar"  ></div></td>
              </tr>
          </table></td>
        </tr>
<!--        <tr>
          <td valign="top"><iframe src="aspuploader/form-multiplefiles.asp" frameborder ="0" width="100%" height="500" scrolling="no"> </iframe></td>
        </tr>-->
        <tr>
          <td valign="top">&nbsp;</td>
        </tr>
        <tr>
          <td valign="top">&nbsp;</td>
        </tr>
      </table>     
      </td>
          </tr>
		  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
        </table>
</form> 
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