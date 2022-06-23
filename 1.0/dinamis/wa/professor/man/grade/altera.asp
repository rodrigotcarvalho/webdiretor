<%On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->

<!--#include file="../../../../inc/caminhos.asp"-->


<!--#include file="../../../../inc/funcoes2.asp"-->



<%
nivel=4
nvg = session("nvg")
opt=request.QueryString("opt")
cod_cons= request.QueryString("cod_cons")
autoriza=Session("autoriza")
Session("autoriza")=autoriza
trava=session("trava")
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=nvg
session("chave")=chave


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2

		Set CON3 = Server.CreateObject("ADODB.Connection") 
		ABRIR3 = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3		
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")
		

if z="2" then	
		
		
codigo = request.querystring("cod")
nome_prof = request.querystring("nome")
apelido = request.querystring("apelido")
sexo = request.querystring("sexo")
nasce= request.querystring("nasce")		
rua = request.querystring("rua")
numero = request.querystring("numero")
complemento = request.querystring("complemento")
bairro= request.querystring("bairro")
municipio= request.querystring("ciddom")
pais = request.querystring("pais")
uf= request.querystring("estadodom")
cep = request.querystring("cep")
telefone = request.querystring("telefones")
uf_natural = request.querystring("estadonat")
nacionalidade = request.querystring("nacionalidade")
natural = request.querystring("cidadenat")
email = request.querystring("email")
ativo = request.querystring("ativo")		
		
pais = pais*1
nacionalidade = nacionalidade*1
municipio = municipio*1
bairro = bairro*1
natural = natural*1		
else				

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Professor WHERE CO_Professor ="& cod_cons
		RS.Open SQL, CON3

cod_prof = RS("CO_Professor")
nome_prof = RS("NO_Professor")
co_usr_prof = RS("CO_Usuario")
end if

	%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../../../../js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
var checkflag = "false";
function check(field) {
if (checkflag == "false") {
for (i = 0; i < field.length; i++) {
field[i].checked = true;}
checkflag = "true";
return "Desmarcar Todos"; }
else {
for (i = 0; i < field.length; i++) {
field[i].checked = false; }
checkflag = "false";
return "Marcar Todos"; }
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
	  
} function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
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
document.all.divEtapa.innerHTML ="<select class=select_style></select>"
document.all.divTurma.innerHTML = "<select class=select_style></select>"
document.all.divDisc.innerHTML = "<select class=select_style></select>"
//recuperarEtapa()
                                                           }
                                               }

                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=e2", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select class=select_style></select>"
document.all.divDisc.innerHTML = "<select class=select_style></select>"
//recuperarTurma()
                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t", true);

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
						 function recuperarDisciplina(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=d", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_d= oHTTPRequest.responseText;
resultado_d = resultado_d.replace(/\+/g," ")
resultado_d = unescape(resultado_d)
document.all.divDisc.innerHTML = resultado_d																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }

								   								   								   
                        </script>
</head> 
<body link="#CC9900" vlink="#CC9900" alink="#CC9900"  background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
            
    <td height="10" class="tb_caminho"><font class="style-caminho">
      <%
	  response.Write(navega)
	  %>
      </font></td>
          </tr>
      <%if opt = "ok" or  opt = "ok2" or opt = "err1" then%>		  
          <tr> 
    <td height="10"> 
      <%if opt = "ok" then
		call mensagens(nivel,628,2,0)
	elseif opt = "ok2" then
		call mensagens(nivel,629,2,0) 
	elseif opt = "err1" then
		call mensagens(nivel,630,1,0) 		
end if		
%>
    </td>
          </tr>
<%end if%>
<% IF trava="s" then%>
            <tr>     
    <td height="10" valign="top"> 
      <%	call mensagens(nivel,9701,0,0) %>
</td>
            </tr> 		 
<%else%>		  
          <tr> 
    <td height="10"> 
      <%	call mensagens(nivel,627,0,0) %>
</td>
          </tr>
<%end if%>		  
          <tr> 
            <td  valign="top"> <form name="alteracao" method="post" action="confirma.asp">
                
        <table width="1000" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo"
>
          <tr class="tb_tit"
> 
            <td height="15" class="tb_tit"
>Professor</td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="9%" height="30" class="tb_subtit"> <div align="right">C&oacute;digo 
                      : </div></td>
                  <td width="11%" height="30"> <font class="form_dado_texto"> 
                    <input name="cod_prof" type="hidden" id="cod_prof" value="<%=cod_prof%>">
                    <%response.Write(cod_prof)%>
                    <input name="tp" type="hidden" id="tp" value="P">
                    <input name="acesso" type="hidden" id="acesso" value="2">
                    <input name="nome_prof" type="hidden" id="nome_prof" value="<% =nome_prof%>">
                    <input name="co_usr_prof" type="hidden" id="co_usr_prof" value="<% =co_usr_prof%>">
                    </font></td>
                  <td width="6%" height="30" class="tb_subtit">
                    <div align="right" >Nome : </div>
                    </td>
                  <td width="74%" height="30"> <font class="form_dado_texto"> 
                    <%response.Write(nome_prof)%>
                    </font> </td>
                </tr>
              </table></td>
          </tr>
          <tr class="tb_tit"
> 
            <td height="15" class="tb_tit"
>Grade de Aulas</td>
          </tr>
          <tr> 
            <td><table width="1000" border="0" cellspacing="0">
                <tr> 
                  <td width="13"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                  <td width="120" class="tb_subtit"> 
                    <div align="center">UNIDADE 
                    </div></td>
                  <td width="100" class="tb_subtit"> <div align="center">CURSO 
                    </div></td>
                  <td width="120" class="tb_subtit"> 
                    <div align="center">ETAPA 
                    </div></td>
                  <td width="141" class="tb_subtit"> <div align="center">TURMA 
                    </div></td>
                  <td width="202" class="tb_subtit"> <div align="center">DISCIPLINA</div></td>
                  <td width="100" class="tb_subtit"> <div align="center">MODELO</div></td>
                  <td width="203" class="tb_subtit"> <div align="center">COORDENADOR 
                    </div></td>
                </tr>
                <%
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Da_Aula where CO_Professor ="& cod_prof 
		RS1.Open SQL1, CON2
		

if RS1.EOF THEN
ELSE

while not RS1.EOF
cod_prof = RS1("CO_Professor")
curso = RS1("CO_Curso")
unidade = RS1("NU_Unidade")
co_etapa= RS1("CO_Etapa")
turma= RS1("CO_Turma")
mat_prin = RS1("CO_Materia_Principal")
mat_fil = RS1("CO_Materia")
tabela = RS1("TP_Nota")
coordenador= RS1("CO_Cord")
		
		valor = unidade&"-"&curso&"-"&co_etapa&"-"&turma&"-"&mat_prin&"-"&mat_fil&"-"&tabela&"-"&coordenador


	Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RSu.Open SQLu, CON0
		
no_unidade = RSu("NO_Unidade")

		Set RSc = Server.CreateObject("ADODB.Recordset")
		SQLc = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RSc.Open SQLc, CON0
		
no_curso = RSc("NO_Abreviado_Curso")


%>
                <tr> 
                  <td width="13"> 
                    <%if trava="n" then%>
                    <input name="grade" type="checkbox" class="borda" value="<% = valor %>"> 
                    <% end if%>
                  </td>
                  <td width="120"> 
                    <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_unidade)%>
                      </font></div></td>
                  <td width="100"> <div align="center"> <font class="form_dado_texto"> 
                      <%
response.Write(no_curso)%>
                      </font></div></td>
                  <td width="120"> 
                    <div align="center"> <font class="form_dado_texto"> 
                      <%

		Set RSe = Server.CreateObject("ADODB.Recordset")
		SQLe = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' and CO_Curso ='"& curso &"'"  
		RSe.Open SQLe, CON0
		
if RSe.EOF THEN
no_etapa="sem etapa"
else
no_etapa=RSe("NO_Etapa")
end if
response.Write(no_etapa)%>
                      </font></div></td>
                  <td width="141"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(turma)%>
                      </font></div></td>
                  <td width="202"> <div align="center"><font class="form_dado_texto"> 
                      <%
		Set RSm = Server.CreateObject("ADODB.Recordset")
		SQLm = "SELECT * FROM TB_Materia where CO_Materia ='"& mat_prin &"'" 
		RSm.Open SQLm, CON0
		
if RSm.EOF THEN
no_mat_prin="sem disciplina"
else
no_mat_prin=RSm("NO_Materia")
end if
response.Write(no_mat_prin)%>
                      </font> </div></td>
                  <td width="100"> <div align="center"><font class="form_dado_texto"> 
                      <%
select case tabela
case "TB_NOTA_A" 
response.Write("Modelo A")
case "TB_NOTA_B" 
response.Write("Modelo B")
case "TB_NOTA_C"
response.Write("Modelo C")
end select

%>
                      </font> </div></td>
                  <td width="203"> <div align="center"><font class="form_dado_texto"> 
                      <%
		Set RS8 = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT * FROM TB_Usuario where CO_Usuario ="& coordenador 
				RS8.Open SQL8, CON
		
no_coordenador = RS8("NO_Usuario")							  
					  response.Write(no_coordenador)%>
                      </font> </div></td>
                </tr>
                <%RS1.MOVENEXT
WEND
END IF				
if trava="n" then%>
                <tr> 
                  <td width="13">&nbsp; </td>
                  <td width="120"> 
                    <div align="center"> 

                      <select name="unidade" class="select_style" onchange="recuperarCurso(this.value)">
                        <option value="0"></option>
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
                  <td width="100"> 
                    <div align="center"> 
                      <div id="divCurso"> 
                        <select class="select_style">
                        </select>
                      </div>
                    </div>
                  </td>
                  <td width="120"> 
                    <div align="center"> 
                      <div id="divEtapa"> 
                        <select class="select_style">
                        </select>
                      </div>
                    </div></td>
                  <td width="141"> <div align="center"> 
                      <div id="divTurma"> 
                        <select class="select_style">
                        </select>
                      </div>
                    </div></td>
                  <td width="202"> <div align="center"> 
				  <div id="divDisc"> 
                        <select class="select_style">
                        </select>
                      </div></div></td>
                  <td width="100">
				  <div align="center"> 
                      <select name="tabela" class="select_style">
                        <option value="999999" selected></option>
                        <option value="TB_NOTA_A">Modelo A </option>
                        <option value="TB_NOTA_B">Modelo B</option>
                        <option value="TB_NOTA_C">Modelo C </option>
                      </select>
                    </div></td>
                  <td width="203">                    <div align="center"> 
                       <select name="coordenador" class="select_style" onChange="MM_callJS('submitforminterno()')">
                       <option value="0" selected></option>
                        <%
		Set RS8 = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT CO_Cord FROM TB_Da_Aula order BY CO_Cord"
		RS8.Open SQL8, CON2				
		
cod_coor_check= 999999		
while not RS8.EOF								
cod_coor = 	RS8("CO_Cord")
if cod_coor = cod_coor_check then
RS8.MOVENEXT	
else		
		Set RS9 = Server.CreateObject("ADODB.Recordset")
		SQL9 = "SELECT * FROM TB_Usuario where CO_Usuario ="&cod_coor
		RS9.Open SQL9, CON
				
		no_coor= RS9("NO_Usuario")
		%>
                        <option value="<%=cod_coor%>">  
                        <%response.Write(no_coor)%>
                       </option> 
                        <%
cod_coor_check = cod_coor
RS8.MOVENEXT
end if
WEND
%>
                     </select> 
                    </div></td>
                </tr>
<% end if%>				
                <tr> 
                  <td width="13" height="15"> </td>
                  <td width="120" height="15"></td>
                  <td width="100" height="15"></td>
                  <td width="120" height="15"></td>
                  <td width="141"> </td>
                  <td width="202"> </td>
                  <td width="100"> </td>
                  <td width="203"> </td>
                </tr>
                <tr> 
                  <td class="tb_tit"
><div align="center"><strong></strong></div></td>
                  <td width="120" class="tb_tit"
>&nbsp;</td>
                  <td class="tb_tit"
>&nbsp;</td>
                  <td width="120" class="tb_tit"
>&nbsp;</td>
                  <td class="tb_tit"
>&nbsp;</td>
                  <td class="tb_tit"
>&nbsp;</td>
                  <td width="100" class="tb_tit"
>&nbsp;</td>
                  <td class="tb_tit"
>&nbsp;</td>
                </tr>
              </table></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td bgcolor="#FFFFFF"> <table width="500" border="0" align="center" cellspacing="0">
                <tr> 
                  <td width="50%">
  				  <%if trava="n" then%>
				  <div align="center"> 
                      <input type=button class="botao_prosseguir" onClick="this.value=check(this.form.grade)" value="Marcar Todos">
                    </div><% end if%>
                    <div align="center"> </div></td>
                  <td width="50%">
  				  <%if trava="n" then%>
				  <div align="center"> 
                      <input name="Submit" type="submit" class="botao_excluir" value="Excluir">
                    </div><% end if%>
					</td>
                </tr>
              </table></td>
          </tr>
        </table>
              </form></td>
          </tr>
  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

</body>
<script type="text/javascript">
<!--
  initInputHighlightScript();
//-->
</script>
<%'end if %>
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