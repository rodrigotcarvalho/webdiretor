<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->

<!--#include file="../../../../inc/funcoes2.asp"-->




<%
opt = request.QueryString("opt")


co_usr = session("co_user")
nivel=4

dia_de= Session("dia_de")
mes_de= Session("dia_de")
dia_ate=Session("dia_ate")
mes_ate=Session("mes_ate")
unidade=Session("unidade")
curso=Session("curso")
etapa=Session("etapa")
turma=Session("turma")
tit=Session("tit")
check_status=Session("check_status")

Session("dia_de")=dia_de
Session("dia_de")=mes_de
Session("dia_ate")=dia_ate
Session("mes_ate")=mes_ate
Session("unidade")=unidade
Session("curso")=curso
Session("etapa")=etapa
Session("turma")=turma
Session("tit")=tit
Session("check_status")=check_status


nvg = session("chave")
chave=nvg
session("chave")=chave
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


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="file:../../../../img/mm_menu.js"></script>
<script type="text/javascript" src="../../../../js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
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

function checksubmit()
{
 if (document.formulario.tipo_doc.value == "0")
  {    alert("Por favor selecione um Tipo de Calendário!")
   document.formulario.tipo_doc.focus()
    return false
 }
  if (document.formulario.status.value == "nulo")
  {    alert("Por favor selecione um Tipo de Evento!")
   document.formulario.status.focus()
    return false
 }

    if (document.formulario.tit.value == "")
  {    alert("Por favor digite um Nome para o Evento!")
    document.formulario.tit.focus()
    return false
  }

dia_de = document.formulario.dia_de.value;
mes_de = document.formulario.mes_de.value;
ano_de = document.formulario.ano_de.value;
dia_ate = document.formulario.dia_ate.value;
mes_ate = document.formulario.mes_ate.value;
ano_ate = document.formulario.ano_ate.value;

data1=dia_de+"/"+mes_de+"/"+ano_de
data2=dia_ate+"/"+mes_ate+"/"+ano_ate
if ( parseInt( data2.split( "/" )[2].toString() + data2.split( "/" )[1].toString() + data2.split( "/" )[0].toString() ) < parseInt( data1.split( "/" )[2].toString() + data1.split( "/" )[1].toString() + data1.split( "/" )[0].toString() ) )
{
  alert( "A data final deve ser maior que a data de inicial!" )
  document.formulario.dia_ate.focus()
  return false
}
  return true
}
//-->
</script>
                         
<script>
<!--
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
                                               oHTTPRequest.open("post", "executa.asp?opt=c", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
document.all.divEtapa.innerHTML ="<select class=borda></select>"
document.all.divTurma.innerHTML = "<select class=borda></select>"
                                                           }
                                               }
                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?opt=e", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select class=borda></select>"

                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?opt=t", true);
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


//--------------------------------------------------------------------------------------------------------------
								   
						 function recuperarCurso1(uTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa_associar.asp?opt=c&nv=1", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_c1  = oHTTPRequest.responseText;
resultado_c1 = resultado_c1.replace(/\+/g," ")
resultado_c1 = unescape(resultado_c1)
document.all.divCurso1.innerHTML =resultado_c1
document.all.divEtapa1.innerHTML ="<select class=borda></select>"
document.all.divTurma1.innerHTML = "<select class=borda></select>"
                                                           }
                                               }
                                               oHTTPRequest.send("u_pub1=" + uTipo);
                                   }


						 function recuperarEtapa1(cTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa_associar.asp?opt=e&nv=1", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                      var resultado_e1= oHTTPRequest.responseText;
																	   
resultado_e1 = resultado_e1.replace(/\+/g," ")
resultado_e1 = unescape(resultado_e1)
document.all.divEtapa1.innerHTML =resultado_e1
document.all.divTurma1.innerHTML = "<select class=borda></select>"

                                                           }
                                               }

                                               oHTTPRequest.send("c_pub1=" + cTipo);
                                   }


						 function recuperarTurma1(eTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa_associar.asp?opt=t&nv=1", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                       var resultado_t1= oHTTPRequest.responseText;
resultado_t1 = resultado_t1.replace(/\+/g," ")
resultado_t1 = unescape(resultado_t1)
document.all.divTurma1.innerHTML = resultado_t1																	   
                                                           }
                                               }
                                               oHTTPRequest.send("e_pub1=" + eTipo);
                                   }  
//--------------------------------------------------------------------------------------------------------------								   
								   
								   						 function recuperarCurso2(uTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa_associar.asp?opt=c&nv=2", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_c2  = oHTTPRequest.responseText;
resultado_c2 = resultado_c2.replace(/\+/g," ")
resultado_c2 = unescape(resultado_c2)
document.all.divCurso2.innerHTML =resultado_c2
document.all.divEtapa2.innerHTML ="<select class=borda></select>"
document.all.divTurma2.innerHTML = "<select class=borda></select>"
                                                           }
                                               }
                                               oHTTPRequest.send("u_pub2=" + uTipo);
                                   }


						 function recuperarEtapa2(cTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa_associar.asp?opt=e&nv=2", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                      var resultado_e2= oHTTPRequest.responseText;
																	   
resultado_e2 = resultado_e2.replace(/\+/g," ")
resultado_e2 = unescape(resultado_e2)
document.all.divEtapa2.innerHTML =resultado_e2
document.all.divTurma2.innerHTML = "<select class=borda></select>"

                                                           }
                                               }

                                               oHTTPRequest.send("c_pub2=" + cTipo);
                                   }


						 function recuperarTurma2(eTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa_associar.asp?opt=t&nv=2", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                       var resultado_t2= oHTTPRequest.responseText;
resultado_t2 = resultado_t2.replace(/\+/g," ")
resultado_t2 = unescape(resultado_t2)
document.all.divTurma2.innerHTML = resultado_t2																	   
                                                           }
                                               }
                                               oHTTPRequest.send("e_pub2=" + eTipo);
                                   }      								   
//--------------------------------------------------------------------------------------------------------------								   
								   
								   						 function recuperarCurso3(uTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa_associar.asp?opt=c&nv=3", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_c3  = oHTTPRequest.responseText;
resultado_c3 = resultado_c3.replace(/\+/g," ")
resultado_c3 = unescape(resultado_c3)
document.all.divCurso3.innerHTML =resultado_c3
document.all.divEtapa3.innerHTML ="<select class=borda></select>"
document.all.divTurma3.innerHTML = "<select class=borda></select>"
                                                           }
                                               }
                                               oHTTPRequest.send("u_pub3=" + uTipo);
                                   }


						 function recuperarEtapa3(cTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa_associar.asp?opt=e&nv=3", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                      var resultado_e3= oHTTPRequest.responseText;
																	   
resultado_e3 = resultado_e3.replace(/\+/g," ")
resultado_e3 = unescape(resultado_e3)
document.all.divEtapa3.innerHTML =resultado_e3
document.all.divTurma3.innerHTML = "<select class=borda></select>"

                                                           }
                                               }

                                               oHTTPRequest.send("c_pub3=" + cTipo);
                                   }


						 function recuperarTurma3(eTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa_associar.asp?opt=t&nv=3", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                       var resultado_t3= oHTTPRequest.responseText;
resultado_t3 = resultado_t3.replace(/\+/g," ")
resultado_t3 = unescape(resultado_t3)
document.all.divTurma3.innerHTML = resultado_t3																	   
                                                           }
                                               }
                                               oHTTPRequest.send("e_pub3=" + eTipo);
                                   }   
								   
//--------------------------------------------------------------------------------------------------------------								   
								   
								   						 function recuperarCurso4(uTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa_associar.asp?opt=c&nv=4", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_c4  = oHTTPRequest.responseText;
resultado_c4 = resultado_c4.replace(/\+/g," ")
resultado_c4 = unescape(resultado_c4)
document.all.divCurso4.innerHTML =resultado_c4
document.all.divEtapa4.innerHTML ="<select class=borda></select>"
document.all.divTurma4.innerHTML = "<select class=borda></select>"
                                                           }
                                               }
                                               oHTTPRequest.send("u_pub4=" + uTipo);
                                   }


						 function recuperarEtapa4(cTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa_associar.asp?opt=e&nv=4", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                      var resultado_e4= oHTTPRequest.responseText;
																	   
resultado_e4 = resultado_e4.replace(/\+/g," ")
resultado_e4 = unescape(resultado_e4)
document.all.divEtapa4.innerHTML =resultado_e4
document.all.divTurma4.innerHTML = "<select class=borda></select>"

                                                           }
                                               }

                                               oHTTPRequest.send("c_pub4=" + cTipo);
                                   }


						 function recuperarTurma4(eTipo)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa_associar.asp?opt=t&nv=4", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                       var resultado_t4= oHTTPRequest.responseText;
resultado_t4 = resultado_t4.replace(/\+/g," ")
resultado_t4 = unescape(resultado_t4)
document.all.divTurma4.innerHTML = resultado_t4																	   
                                                           }
                                               }
                                               oHTTPRequest.send("e_pub4=" + eTipo);
                                   }

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../../img/menu_r1_c2_f3.gif','../../../../img/menu_r1_c2_f2.gif','../../../../img/menu_r1_c2_f4.gif','../../../../img/menu_r1_c4_f3.gif','../../../../img/menu_r1_c4_f2.gif','../../../../img/menu_r1_c4_f4.gif','../../../../img/menu_r1_c6_f3.gif','../../../../img/menu_r1_c6_f2.gif','../../../../img/menu_r1_c6_f4.gif','../../../../img/menu_r1_c8_f3.gif','../../../../img/menu_r1_c8_f2.gif','../../../../img/menu_r1_c8_f4.gif','../../../../img/menu_direita_r2_c1_f3.gif','../../../../img/menu_direita_r2_c1_f2.gif','../../../../img/menu_direita_r2_c1_f4.gif','../../../../img/menu_direita_r4_c1_f3.gif','../../../../img/menu_direita_r4_c1_f2.gif','../../../../img/menu_direita_r4_c1_f4.gif','../../../../img/menu_direita_r6_c1_f3.gif','../../../../img/menu_direita_r6_c1_f2.gif','../../../../img/menu_direita_r6_c1_f4.gif')">
<% call cabecalho (nivel)
	  %>
<table width="1000" height="669" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
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
      <%	call mensagens(4,68,2,0) 
	  
	  
%>
</td></tr>
<%end if%>	  
	  
                <tr> 
                    
    <td height="10"> 
      <%	call mensagens(4,67,0,0) 
	  
	  
%>
</td></tr>
<tr>

            <td valign="top"> 
			
<FORM name="formulario" METHOD="POST" ACTION="bd.asp?opt=i" onSubmit="return checksubmit()">
                
        <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td width="653" height="15" class="tb_tit">Incluir Evento 
              <input name="co_grupo" type="hidden" id="co_grupo" value="<% = co_grupo %>"></td>
          </tr>
          <tr> 
            <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="220" class="tb_subtit"> 
                          <div align="center">Tipo 
                            de Calend&aacute;rio</div></td>
                        <td width="220" class="tb_subtit"> 
                          <div align="center">Tipo 
                            de Eventos</div></td>
                        <td width="220" class="tb_subtit"> 
                          <div align="center">Data 
                            da Inicial</div></td>
                        <td width="220" class="tb_subtit">
<div align="center">Data 
                            Final</div></td>
                        <td width="80" class="tb_subtit">
<div align="center">Repete</div></td>
                      </tr>
                      <tr> 
                        <td width="220"> 
                          <div align="center"><font class="form_dado_texto"> 
                            <select name="tipo_doc" class="borda" id="tipo_doc" >
                              <option value="0"></option>
                              <%
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Tipo_Calendario  order by NU_Prioridade_Combo"
		RS1.Open SQL1, CON0

while not RS1.EOF
tp_noticia=RS1("TP_Calen")
descricao=RS1("TX_Descricao")

%>
                              <option value="<%=tp_noticia%>"><%=descricao%></option>
                              <%
RS1.movenext
Wend
%>
                            </select>
                            </font></div></td>
                        <td width="220"> 
                          <div align="center"> 
                            <select name="status" class="borda">
                              <option value="nulo" selected></option>
                              <%
		Set RS1e = Server.CreateObject("ADODB.Recordset")
		SQL1e = "SELECT * FROM TB_Tipo_Eventos order by NU_Prioridade_Combo"
		RS1e.Open SQL1e, CON0

while not RS1e.EOF
tp_evento=RS1e("TP_Evento")
descricao_evento=RS1e("TX_Descricao")

%>
                              <option value="<%=tp_evento%>"><%=descricao_evento%></option>
                              <%
RS1e.movenext
Wend
%>
                            </select>
                          </div></td>
                        <td width="220">
<div align="center"><font class="form_dado_texto"> 
                            <select name="dia_de" id="dia_de" class="borda">
                              <% 
							 For i =1 to 31
							 dia=dia*1
							 if dia=i then 
								if dia<10 then
								dia="0"&dia
								end if
							 %>
                              <option value="<%response.Write(i)%>" selected> 
                              <%response.Write(dia)%>
                              </option>
                              <% else
								if i<10 then
								i="0"&i
								end if
							%>
                              <option value="<%response.Write(i)%>"> 
                              <%response.Write(i)%>
                              </option>
                              <% end if 
							next
							%>
                            </select>
                            / 
                            <select name="mes_de" id="mes_de" class="borda">
                              <%mes=mes*1
								if mes="1" or mes=1 then%>
                              <option value="1" selected>janeiro</option>
                              <% else%>
                              <option value="1">janeiro</option>
                              <%end if
								if mes="2" or mes=2 then%>
                              <option value="2" selected>fevereiro</option>
                              <% else%>
                              <option value="2">fevereiro</option>
                              <%end if
								if mes="3" or mes=3 then%>
                              <option value="3" selected>mar&ccedil;o</option>
                              <% else%>
                              <option value="3">mar&ccedil;o</option>
                              <%end if
								if mes="4" or mes=4 then%>
                              <option value="4" selected>abril</option>
                              <% else%>
                              <option value="4">abril</option>
                              <%end if
								if mes="5" or mes=5 then%>
                              <option value="5" selected>maio</option>
                              <% else%>
                              <option value="5">maio</option>
                              <%end if
								if mes="6" or mes=6 then%>
                              <option value="6" selected>junho</option>
                              <% else%>
                              <option value="6">junho</option>
                              <%end if
								if mes="7" or mes=7 then%>
                              <option value="7" selected>julho</option>
                              <% else%>
                              <option value="7">julho</option>
                              <%end if%>
                              <%if mes="8" or mes=8 then%>
                              <option value="8" selected>agosto</option>
                              <% else%>
                              <option value="8">agosto</option>
                              <%end if
								if mes="9" or mes=9 then%>
                              <option value="9" selected>setembro</option>
                              <% else%>
                              <option value="9">setembro</option>
                              <%end if
								if mes="10" or mes=10 then%>
                              <option value="10" selected>outubro</option>
                              <% else%>
                              <option value="10">outubro</option>
                              <%end if
								if mes="11" or mes=11 then%>
                              <option value="11" selected>novembro</option>
                              <% else%>
                              <option value="11">novembro</option>
                              <%end if
								if mes="12" or mes=12 then%>
                              <option value="12" selected>dezembro</option>
                              <% else%>
                              <option value="12">dezembro</option>
                              <%end if%>
                            </select>
                            / 
                            <%response.write(ano_letivo)%><input name="ano_de" type="hidden" value="<%response.write(ano_letivo)%>">
                            </font></div></td>
                        <td width="220">
<div align="center"><font class="form_dado_texto"> 
                            <select name="dia_ate" id="dia_ate" class="borda">
                              <option value="1">01</option>
                              <option value="2">02</option>
                              <option value="3">03</option>
                              <option value="4">04</option>
                              <option value="5">05</option>
                              <option value="6">06</option>
                              <option value="7">07</option>
                              <option value="8">08</option>
                              <option value="9">09</option>
                              <option value="10">10</option>
                              <option value="11">11</option>
                              <option value="12">12</option>
                              <option value="13">13</option>
                              <option value="14">14</option>
                              <option value="15">15</option>
                              <option value="16">16</option>
                              <option value="17">17</option>
                              <option value="18">18</option>
                              <option value="19">19</option>
                              <option value="20">20</option>
                              <option value="21">21</option>
                              <option value="22">22</option>
                              <option value="23">23</option>
                              <option value="24">24</option>
                              <option value="25">25</option>
                              <option value="26">26</option>
                              <option value="27">27</option>
                              <option value="28">28</option>
                              <option value="29">29</option>
                              <option value="30">30</option>
                              <option value="31" selected>31</option>
                            </select>
                            / 
                            <select name="mes_ate" id="mes_ate" class="borda">
                              <option value="1">janeiro</option>
                              <option value="2">fevereiro</option>
                              <option value="3">mar&ccedil;o</option>
                              <option value="4">abril</option>
                              <option value="5">maio</option>
                              <option value="6">junho</option>
                              <option value="7">julho</option>
                              <option value="8">agosto</option>
                              <option value="9">setembro</option>
                              <option value="10">outubro</option>
                              <option value="11">novembro</option>
                              <option value="12" selected>dezembro</option>
                            </select>
                            / 
                            <select name="ano_ate" class="borda">
                              <%
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Ano_Letivo order by NU_Ano_Letivo"
		RS0.Open SQL0, CON
		while not RS0.EOF 
		ano_bd=RS0("NU_Ano_Letivo")
		if ano_bd=Session("ano_letivo") then
		%>
                              <option value="<%=ano_bd%>" selected><%=ano_bd%></option>		
<% else%>
                              <option value="<%=ano_bd%>"><%=ano_bd%></option>
                              <%
		end if					  
		RS0.MOVENEXT
		WEND 		
				%>
                            </select>
                            </font></div></td>
                        <td width="80">
<div align="center"> 
                            <select name="repete" id="repete" class="borda">
                              <option value="1">Sim</option>
                              <option value="0" selected>N&atilde;o</option>
                            </select>
                          </div></td>
                      </tr>
                      <tr> 
                        <td width="220">&nbsp;</td>
                        <td width="220">&nbsp;</td>
                        <td width="220">&nbsp;</td>
                        <td width="220">&nbsp;</td>
                        <td width="80">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="220" class="tb_subtit">&nbsp;</td>
                        <td width="220" class="tb_subtit">Nome</td>
                        <td width="220" class="tb_subtit">&nbsp;</td>
                        <td width="220" class="tb_subtit">&nbsp;</td>
                        <td width="80" class="tb_subtit">&nbsp;</td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td colspan="4"><input name="tit" type="text" class="borda" id="tit4" size="150"></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td width="247">&nbsp;</td>
                  <td colspan="2">&nbsp;</td>
                  <td width="247">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="4" class="tb_tit">Este evento s&oacute; poder&aacute; 
                    ser acessado pelo segmento abaixo:</td>
                </tr>
                <tr> 
                  <td class="tb_subtit"> <div align="center">UNIDADE </div></td>
                  <td width="247" class="tb_subtit"> <div align="center">CURSO 
                    </div></td>
                  <td width="247" class="tb_subtit"> <div align="center">ETAPA 
                    </div></td>
                  <td class="tb_subtit"> <div align="center">TURMA </div></td>
                </tr>
                <tr> 
                  <td> <div align="center"> 
                      <select name="unidade" class="borda" onChange="recuperarCurso(this.value)">
                        <option value="999990" selected></option>
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
                  <td> <div align="center"> 
                      <div id="divCurso"> 
                        <select name="curso" class="borda" id="curso">
                          <option value="999990" selected></option>
                        </select>
                      </div>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divEtapa"> 
                        <select name="etapa" class="borda" id="etapa">
                          <option value="999990" selected></option>
                        </select>
                      </div>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divTurma"> 
                        <select name="turma" class="borda" id="turma">
                          <option value="999990" selected></option>
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
                <tr> 
                  <td colspan="4">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="4" class="tb_tit">Este evento tamb&eacute;m poder&aacute; 
                    ser acessado por estes outros segmentos abaixo:</td>
                </tr>
                <tr> 
                  <td> <div align="center"> 
                      <select name="unidade1" class="borda" id="unidade1" onChange="recuperarCurso1(this.value)">
                        <option value="999990" selected></option>
                        <%		

		Set RS0u = Server.CreateObject("ADODB.Recordset")
		SQL0u = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0u.Open SQL0u, CON0
NU_Unidade_Check=999999		
While not RS0u.EOF
NU_Unidade = RS0u("NU_Unidade")
NO_Abr = RS0u("NO_Abr")
if NU_Unidade = NU_Unidade_Check then
RS0u.MOVENEXT		
else
%>
                        <option value="<%response.Write(NU_Unidade)%>"> 
                        <%response.Write(NO_Abr)%>
                        </option>
                        <%						

NU_Unidade_Check = NU_Unidade
RS0u.MOVENEXT
end if
WEND
%>
                      </select>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divCurso1"> 
                        <select name="curso1" class="borda" id="curso1" onChange="recuperarEtapa1(this.value)">
                          <option value="999990" selected></option>
                        </select>
                      </div>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divEtapa1"> 
                        <select name="etapa1" class="borda" id="etapa1" onChange="recuperarTurma1(this.value)">
                          <option value="999990" selected></option>
                        </select>
                      </div>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divTurma1"> 
                        <select name="turma1" class="borda" id="turma1">
                          <option value="999990" selected></option>
                        </select>
                      </div>
                    </div></td>
                </tr>
                <tr> 
                  <td> <div align="center"> 
                      <select name="unidade2" class="borda" onChange="recuperarCurso2(this.value)">
                        <option value="999990" selected></option>
                        <%		

		Set RS0u = Server.CreateObject("ADODB.Recordset")
		SQL0u = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0u.Open SQL0u, CON0
NU_Unidade_Check=999999		
While not RS0u.EOF
NU_Unidade = RS0u("NU_Unidade")
NO_Abr = RS0u("NO_Abr")
if NU_Unidade = NU_Unidade_Check then
RS0u.MOVENEXT		
else
%>
                        <option value="<%response.Write(NU_Unidade)%>"> 
                        <%response.Write(NO_Abr)%>
                        </option>
                        <%						

NU_Unidade_Check = NU_Unidade
RS0u.MOVENEXT
end if
WEND
%>
                      </select>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divCurso2"> 
                        <select name="curso2" class="borda" id="curso2" onChange="recuperarEtapa2(this.value)">
                          <option value="999990" selected></option>
                        </select>
                      </div>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divEtapa2"> 
                        <select name="etapa2" class="borda" id="etapa2" onChange="recuperarTurma2(this.value)">
                          <option value="999990" selected></option>
                        </select>
                      </div>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divTurma2"> 
                        <select name="turma2" class="borda" id="turma2">
                          <option value="999990" selected></option>
                        </select>
                      </div>
                    </div></td>
                </tr>
                <tr> 
                  <td> <div align="center"> 
                      <select name="unidade3" class="borda" onChange="recuperarCurso3(this.value)">
                        <option value="999990" selected></option>
                        <%		

		Set RS0u = Server.CreateObject("ADODB.Recordset")
		SQL0u = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0u.Open SQL0u, CON0
NU_Unidade_Check=999999		
While not RS0u.EOF
NU_Unidade = RS0u("NU_Unidade")
NO_Abr = RS0u("NO_Abr")
if NU_Unidade = NU_Unidade_Check then
RS0u.MOVENEXT		
else
%>
                        <option value="<%response.Write(NU_Unidade)%>"> 
                        <%response.Write(NO_Abr)%>
                        </option>
                        <%						

NU_Unidade_Check = NU_Unidade
RS0u.MOVENEXT
end if
WEND
%>
                      </select>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divCurso3"> 
                        <select name="curso3" class="borda" id="curso3" onChange="recuperarEtapa3(this.value)">
                          <option value="999990" selected></option>
                        </select>
                      </div>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divEtapa3"> 
                        <select name="etapa3" class="borda" id="etapa3" onChange="recuperarTurma3(this.value)">
                          <option value="999990" selected></option>
                        </select>
                      </div>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divTurma3"> 
                        <select name="turma3" class="borda" id="turma">
                          <option value="999990" selected></option>
                        </select>
                      </div>
                    </div></td>
                </tr>
                <tr> 
                  <td> <div align="center"> 
                      <select name="unidade4" class="borda" onChange="recuperarCurso4(this.value)">
                        <option value="999990" selected></option>
                        <%		

		Set RS0u = Server.CreateObject("ADODB.Recordset")
		SQL0u = "SELECT * FROM TB_Unidade order by NO_Abr"
		RS0u.Open SQL0u, CON0
NU_Unidade_Check=999999		
While not RS0u.EOF
NU_Unidade = RS0u("NU_Unidade")
NO_Abr = RS0u("NO_Abr")
if NU_Unidade = NU_Unidade_Check then
RS0u.MOVENEXT		
else
%>
                        <option value="<%response.Write(NU_Unidade)%>"> 
                        <%response.Write(NO_Abr)%>
                        </option>
                        <%						

NU_Unidade_Check = NU_Unidade
RS0u.MOVENEXT
end if
WEND
%>
                      </select>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divCurso4"> 
                        <select name="curso4" class="borda" id="curso4" onChange="recuperarEtapa4(this.value)">
                          <option value="999990" selected></option>
                        </select>
                      </div>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divEtapa4"> 
                        <select name="etapa4" class="borda" id="etapa4" onChange="recuperarTurma4(this.value)">
                          <option value="999990" selected></option>
                        </select>
                      </div>
                    </div></td>
                  <td> <div align="center"> 
                      <div id="divTurma4"> 
                        <select name="turma4" class="borda" id="turma4">
                          <option value="999990" selected></option>
                        </select>
                      </div>
                    </div></td>
                </tr>
                <tr> 
                  <td colspan="4"><hr width="1000"></td>
                </tr>
                <tr> 
                  <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="25%"> <div align="center"> 
                            <input name="SUBMIT5" type=button class="botao_cancelar" onClick="MM_goToURL('parent','docs.asp?pagina=1&v=s');return document.MM_returnValue" value="Voltar">
                        </div></td>
                        <td width="25%"> <div align="center"> </div></td>
                        <td width="25%"> <div align="center"> </div></td>
                        <td width="25%"> <div align="center"> 
                            <input name="SUBMIT" type=SUBMIT class="botao_prosseguir" value="Incluir">
                        </div></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
              </table>
 </td>
          </tr>
        </table>
              </form></td>
          </tr>
		  <tr>
    <td height="59" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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