<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes3.asp"-->
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../../../estilos.css" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--

function MM_popupMsg(msg) { //v1.0
  alert(msg);
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
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
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresiz!=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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
   var f=document.forms[3]; 
      f.submit(); 
} 

function mudar_cor_focus(celula){
   celula.style.backgroundColor="#D8FF9D"

}
function mudar_cor_blur_par(celula){
   celula.style.backgroundColor="#FFFFFF"
} 
function mudar_cor_blur_impar(celula){
   celula.style.backgroundColor="#FFFFE1"
} 
function mudar_cor_blur_erro(celula){
   celula.style.backgroundColor="#CC0000"
}  
function checksubmit()
{
// if (document.nota.pt.value == "")
//  {    alert("Por favor digite um peso para os Testes!")
//    document.nota.pt.focus()
//    return false
//  }
//  if (isNaN(document.nota.pt.value))
//  {    alert("O peso dos Testes deve ser um número!")
//    document.nota.pt.focus()
//    return false
//  }  
//    if (document.nota.pp.value == "")
//  {    alert("Por favor digite um peso para as Provas!")
//    document.nota.pp.focus()
//    return false
//  }
//  if (isNaN(document.nota.pp.value))
//  {    alert("O peso das Provas deve ser um número!")
//    document.nota.pp.focus()
//    return false
//  }
  return true
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
<script language="javascript"> 
  
    function keyPressed(TB, e, max_right, max_bottom)  
    { 
        if (e.keyCode == 40 || e.keyCode == 13) { // arrow down 
            if (TB.split("c")[0] < max_bottom) 
            document.getElementById(eval(TB.split("c")[0] + '+1') + 'c' + TB.split("c")[1]).focus(); 
            if (TB.split("c")[0] == max_bottom) 
            document.getElementById(1 + 'c' + TB.split("c")[1]).focus();


        } 
  
        if (e.keyCode == 38) { // arrow up 
            if(TB.split("c")[0] > 1) 
            document.getElementById(eval(TB.split("c")[0] + '-1') + 'c' + TB.split("c")[1]).focus(); 
            if (TB.split("c")[0] == 1) 
            document.getElementById(max_bottom + 'c' + TB.split("c")[1]).focus(); 
		
        } 
  
        if (e.keyCode == 37) { // arrow left 
            if(TB.split("c")[1] > 1) 
            document.getElementById(TB.split("c")[0] + 'c' + eval(TB.split("c")[1] + '-1')).focus();             
            if (TB.split("c")[1] == 1) 
            document.getElementById(TB.split("c")[0] + 'c' + max_right).focus(); 

		}   
  
        if (e.keyCode == 39) { // arrow right 
            if(TB.split("c")[1] < max_right) 
            document.getElementById(TB.split("c")[0] + 'c' + eval(TB.split("c")[1] + '+1')).focus();  
            if (TB.split("c")[1] == max_right) 
            document.getElementById(TB.split("c")[0] + 'c' + 1).focus(); 

		}                  
    } 
  
</script> 
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%

obr=session("obr")
nota = session("nota")
vetor_obr=split(obr,"$!$")
co_materia = vetor_obr(0)
unidades = vetor_obr(1)
grau = vetor_obr(2)
serie = vetor_obr(3)
turma = vetor_obr(4)
periodo = vetor_obr(5)
ano_letivo= session("ano_letivo")

co_usr = session("co_user")
grupo=session("grupo")
autoriza=session("autoriza")
coordenador = session("coordenador")

co_usr=co_usr*1
coordenador=coordenador*1
autoriza=autoriza*1
co_prof = session("co_prof")
bancoPauta = session("bancoPauta")
CAMINHOn = session("caminhoBancoPauta")

	if opt="cln" then
		if bancoPauta ="Pauta_A" then

		Call pauta(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_prof,"A","cln",0)
		else
			if bancoPauta="Pauta_B" then
			Call pauta(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_prof,"B","cln",0)
			else
				if bancoPauta ="Pauta_C" then
				Call pauta(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_prof,"C","cln",0)
				else
				response.Write("ERRO")
				End if
			end if
		end if
	
	ELSEIF ((co_usr=coordenador and autoriza=5) AND trava<>"s") or (grupo<>"COO" and autoriza=5 AND trava<>"s") then
		if bancoPauta ="Pauta_A" then
		Call pauta(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_prof,"A","edit",0)
		else
			if bancoPauta="Pauta_B" then
			Call pauta(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_prof,"B","edit",0)
			else
				if bancoPauta ="Pauta_C" then
				Call pauta(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_prof,"C","edit",0)
				else
				response.Write("ERRO")
				End if
			end if
		end if	
	else
		if bancoPauta ="Pauta_A" then
		Call pauta(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_prof,"A","blq",0)
		else
			if bancoPauta="Pauta_B" then
			Call pauta(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_prof,"B","blq",0)
			else
				if bancoPauta ="Pauta_C" then
				Call pauta(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_prof,"C","blq",0)
				else
				response.Write("ERRO")
				End if
			end if
			
		end if
	end if	
	
 %>
</body>
</html>
