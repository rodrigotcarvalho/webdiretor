<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<% 
FUNCTION URLDecode(str)
'// This function:
'// - decodes any utf-8 encoded characters into unicode characters eg. (%C3%A5 = å)
'// - replaces any plus sign separators with a space character
'//
'// IMPORTANT:
'// Your webpage must use the UTF-8 character set. Easiest method is to use this META tag:
'// <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
'//
    Dim objScript
    Set objScript = Server.CreateObject("ScriptControl")
    objScript.Language = "JavaScript"
    URLDecode = objScript.Eval("decodeURIComponent(""" & str & """.replace(/\+/g,"" ""))")
    Set objScript = NOTHING
END FUNCTION


opt = request.QueryString("opt")
nvg = session("chave")
chave=nvg
session("chave")=chave
session("nvg")=nvg
nivel=4
if opt="F"  then
	if transicao = "S" then
	 area="wd"
	 site_escola="simplynet2.tempsite.ws/wd/"&ambiente_escola&"/wf/apoio/ctl/pub/"
	else
		if left(ambiente_escola,5)= "teste" then
			area="wdteste"
			link="http://www.simplynet.com.br/"&area&"/"&ambiente_escola
		else
			area="wd"
			link="http://www.webdiretor.com.br/"&ambiente_escola
		end if	
	end if
	Set upl = Server.CreateObject("Persits.Upload")
	caminho_download = Server.MapPath("../../../../anexos")
	caminho_upload = caminho_download&"/original"
	'response.Write(caminho_upload)			
	contarq = upl.Save(caminho_upload)
	if contarq = 0 then
		response.Write("IsEmpty")		
	end if	
	'response.Write(caminho_upload)						
	Session("upl_total") = upl.TotalBytes	
	Set dir = CreateObject("Scripting.FileSystemObject") 


	nome_pasta = caminho_upload
	
	set FSO = server.createObject("Scripting.FileSystemObject")
	
	Set Rs_ordena = Server.CreateObject ( "ADODB.RecordSet" )
	
	Set pasta = FSO.GetFolder(nome_pasta)
	
	
	
	Set arquivos = pasta.Files
	
	' Nome do documento XML de saida
	 arquivo_xml= "..\..\..\..\anexos\anexos.xml"
	
	' cria um arquivo usando o file system object
	set fso = createobject("scripting.filesystemobject")
	
	' cria o arquivo texto no disco com opção de sobrescrever o arquivo existente
	'response.Write(server.mappath(arquivo_xml))
	Set act = fso.CreateTextFile(server.mappath(arquivo_xml), true)
	
	' cabecalho do XML
	act.WriteLine("<?xml version=""1.0"" encoding=""ISO-8859-1""?>")
	act.WriteLine("<arquivos>")
	
	for each arquivo in arquivos
		nome_arquivo =arquivo.Name 
			'nome_arquivo = replace(nome_arquivo,"%2E", ".")			
		vetor_nome = split(nome_arquivo,".")
		'Renomear arquivo

		
		if ucase(vetor_nome(1)) <> "TMP" and ucase(vetor_nome(1)) <> "XML"  then 	
			'Renomear o arquivo		
			nome_alterado = vetor_nome(0)
				strReplacement = replace(nome_alterado,"+","_e_")
				strReplacement = replace(strReplacement," ","_")
				strReplacement = replace(strReplacement,"&","_e_")	
				strReplacement = replace(strReplacement,"-","")												
				strReplacement = replace(strReplacement,"´","")
				strReplacement = replace(strReplacement,"'","")
				strReplacement = replace(strReplacement,"Á","A")
				strReplacement = replace(strReplacement,"À","A")
				strReplacement = replace(strReplacement,"Â","A")
				strReplacement = replace(strReplacement,"Ã","A")
				strReplacement = replace(strReplacement,"Ç","C")			
				strReplacement = replace(strReplacement,"É","E")
				strReplacement = replace(strReplacement,"Ê","E")
				strReplacement = replace(strReplacement,"Í","I")
				strReplacement = replace(strReplacement,"Ó","O")
				strReplacement = replace(strReplacement,"Ô","O")
				strReplacement = replace(strReplacement,"Õ","O")
				strReplacement = replace(strReplacement,"Ú","U")
				strReplacement = replace(strReplacement,"Ü","U")	
				strReplacement = replace(strReplacement,"á","a")
				strReplacement = replace(strReplacement,"à","a")
				strReplacement = replace(strReplacement,"â","a")
				strReplacement = replace(strReplacement,"ã","a")
				strReplacement = replace(strReplacement,"ç","c")
				strReplacement = replace(strReplacement,"é","e")
				strReplacement = replace(strReplacement,"ê","e")
				strReplacement = replace(strReplacement,"í","i")
				strReplacement = replace(strReplacement,"ó","o")
				strReplacement = replace(strReplacement,"ô","o")
				strReplacement = replace(strReplacement,"õ","o")
				strReplacement = replace(strReplacement,"ú","u")
				strReplacement = replace(strReplacement,"ª","")
				strReplacement = replace(strReplacement,"º","")
				strReplacement = replace(strReplacement,"Âº","")	
				strReplacement = replace(strReplacement,"Ãº","")												
				nome_alterado = replace(strReplacement,"ü","u")
			nome_alterado = nome_alterado&"."&vetor_nome(1)				
			
			
			
			
			'nome_alterado = replace(nome_alterado,"%2E", ".")	
			'copiar para pasta de acesso	
			fso.CopyFile  caminho_upload&"\"&nome_arquivo,  caminho_download&"\"&nome_alterado		
			
			act.WriteLine("<arquivo>" )
			act.WriteLine("<file>" & nome_alterado & "</file>" )
			act.WriteLine("<nome>" & vetor_nome(0) & "</nome>" )
			act.WriteLine("<fileOriginal>" & nome_arquivo & "</fileOriginal>" )			
			if ucase(vetor_nome(1)) = "GIF" or ucase(vetor_nome(1)) = "JPG" or ucase(vetor_nome(1)) = "JPEG" or ucase(vetor_nome(1)) = "PNG" or ucase(vetor_nome(1)) = "BMP" then
				formato = "IMG"
			else
				formato = ucase(vetor_nome(1))	
			end if			
			act.WriteLine("<formato>" & formato & "</formato>" )		
			act.WriteLine("</arquivo>" )	
		end if			
	next
	' fecha a tag 
	act.WriteLine("</arquivos>")
	' fecha o objeto xml
	act.close
	Set upl = Nothing 
	response.Redirect("index.asp?nvg=WF-AS-CO-EEF&opt=ok2")
end if


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0

    	Set CON_WF = Server.CreateObject("ADODB.Connection") 
		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_WF.Open ABRIR_WF		
		
 call VerificaAcesso (CON,chave,nivel)
autoriza=Session("autoriza")

 call navegacao (CON,chave,nivel)
navega=Session("caminho")

Call LimpaVetor3

%>
<html>
<head>
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
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
    document.busca.busca1.focus()
    return false
  }
    if (document.busca.busca1.value == "" && document.busca.busca2.value == "")
  {    alert("Por favor digite uma opção de busca!")
    document.busca.busca1.focus()
    return false
  }
  return true
}

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
</head>

<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
            
    <td width="20" height="10" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
    </td>
          </tr>
 <%if opt = "err" then%>
  <tr> 
    <td height="10"> 
      <%
		call mensagens(nivel,53,1,0)
%>
    </td>
  </tr> 
  <% end if%>  
   <tr>             
    <td height="10" valign="top"> 
      <%call mensagens(nivel,80,0,0) %>
    </td>
			  </tr>			  
        <form action="upload.asp?opt=F" ENCTYPE="multipart/form-data" method="post">

        
                <tr class="tb_corpo"> 
                  
    <td height="10" class="tb_tit">Envio de arquivos para anexo</td>
                </tr>
                <tr> 
                  
    <td valign="top">
    <span id="arquivo1" style="display:block;" >    
    <span style="display:table-cell;width:1000px;vertical-align: middle;" ><INPUT TYPE="FILE" SIZE=60 NAME="FILE1" class="borda" style="width:40%;display: block;margin-left: auto;margin-right: auto"></span>
    </span>
    <span id="arquivo2" style="display:block;" >    
    <span style="display:table-cell;width:1000px;vertical-align: middle;" ><INPUT TYPE="FILE" SIZE=60 NAME="FILE2" class="borda" style="width:40%;display: block;margin-left: auto;margin-right: auto"></span>
    </span>
    <span id="arquivo3" style="display:block;" >    
    <span style="display:table-cell;width:1000px;vertical-align: middle;" ><INPUT TYPE="FILE" SIZE=60 NAME="FILE3" class="borda" style="width:40%;display: block;margin-left: auto;margin-right: auto"></span>
    </span>
    <span id="arquivo4" style="display:block;" >    
    <span style="display:table-cell;width:1000px;vertical-align: middle;" ><INPUT TYPE="FILE" SIZE=60 NAME="FILE4" class="borda" style="width:40%;display: block;margin-left: auto;margin-right: auto"></span>
    </span>
    <span id="arquivo5" style="display:block;" >    
    <span style="display:table-cell;width:1000px;vertical-align: middle;" ><INPUT TYPE="FILE" SIZE=60 NAME="FILE5" class="borda" style="width:40%;display: block;margin-left: auto;margin-right: auto"></span>
    </span>                
    <span id="botao" style="display:block;" >  
    <span style="display:table-cell;width:1000px;vertical-align: middle;" ><input type="submit" class="botao_prosseguir" value="Enviar arquivos" style="width:10%;display: block;margin-left: auto;margin-right: auto" >
    </span>
    </span></td>
                </tr>
      </form>                
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