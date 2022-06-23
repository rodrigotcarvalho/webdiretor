<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%
opt = request.QueryString("opt")

ano_letivo_wf = Session("ano_letivo_wf")
co_usr = session("co_user")
nivel=4
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



nvg = request.QueryString("nvg")
chave=nvg
session("chave")=chave
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano = DatePart("yyyy", now) 
mes = DatePart("m", now) 
dia = DatePart("d", now) 


if ano_letivo_wf=ano then
	data_expira=dia&"/"&mes&"/"&ano
else

	data_expira="31/12/"&ano_letivo_wf
end if	

ano_info=nivel&"-"&chave&"-"&ano_letivo_wf


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
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
function checksubmit()
{
 if (document.form.pasta.value == "nulo")
  {    alert("Por favor selecione uma pasta!")
   document.form.pasta.focus()
    return false
 }
 
  return true
}
//-->
</script>
<script>
<!--

// A fun��o abaixo pega a vers�o mais nova do xmlhttp do IE e verifica se � Firefox. Funciona nos dois.
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
// Cria��o do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicita��o HTTP. O primeiro par�metro informa o m�todo post/get
// O segundo par�metro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicita��o s�ncrona, o par�metro deve ser false
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=c", true);
// Para solicita��es utilizando o m�todo post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A fun��o abaixo � executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto j� completou a solicita��o
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto � gerado no arquivo executa.asp e colocado no div
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
document.all.divEtapa.innerHTML ="<select class=borda></select>"
document.all.divTurma.innerHTML = "<select class=borda></select>"
//recuperarEtapa()
                                                           }
                                               }
// Abaixo � enviada a solicita��o. Note que a configura��o
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {
// Cria��o do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicita��o HTTP. O primeiro par�metro informa o m�todo post/get
// O segundo par�metro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicita��o s�ncrona, o par�metro deve ser false
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=e", true);
// Para solicita��es utilizando o m�todo post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A fun��o abaixo � executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto j� completou a solicita��o
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto � gerado no arquivo executa.asp e colocado no div
                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select class=borda></select>"
//recuperarTurma()
                                                           }
                                               }
// Abaixo � enviada a solicita��o. Note que a configura��o
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {
// Cria��o do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicita��o HTTP. O primeiro par�metro informa o m�todo post/get
// O segundo par�metro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicita��o s�ncrona, o par�metro deve ser false
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t", true);
// Para solicita��es utilizando o m�todo post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A fun��o abaixo � executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto j� completou a solicita��o
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto � gerado no arquivo executa.asp e colocado no div
                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t																	   
                                                           }
                                               }
// Abaixo � enviada a solicita��o. Note que a configura��o
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("e_pub=" + eTipo);

                                   }
//-->
</script>

<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../../img/menu_r1_c2_f3.gif','../../../../img/menu_r1_c2_f2.gif','../../../../img/menu_r1_c2_f4.gif','../../../../img/menu_r1_c4_f3.gif','../../../../img/menu_r1_c4_f2.gif','../../../../img/menu_r1_c4_f4.gif','../../../../img/menu_r1_c6_f3.gif','../../../../img/menu_r1_c6_f2.gif','../../../../img/menu_r1_c6_f4.gif','../../../../img/menu_r1_c8_f3.gif','../../../../img/menu_r1_c8_f2.gif','../../../../img/menu_r1_c8_f4.gif','../../../../img/menu_direita_r2_c1_f3.gif','../../../../img/menu_direita_r2_c1_f2.gif','../../../../img/menu_direita_r2_c1_f4.gif','../../../../img/menu_direita_r4_c1_f3.gif','../../../../img/menu_direita_r4_c1_f2.gif','../../../../img/menu_direita_r4_c1_f4.gif','../../../../img/menu_direita_r6_c1_f3.gif','../../../../img/menu_direita_r6_c1_f2.gif','../../../../img/menu_direita_r6_c1_f4.gif')">
<% call cabecalho (nivel)
	  %>
	  
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
                    
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
	  </td>
	  </tr>
      <%
if opt = "a" then%>
  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(nivel,9705,2,0)
%>
    </td>
                  </tr>
                  <% 	end if 

%>                  <tr> 
                    
    <td height="10"> 
      <%	call mensagens(nivel,9704,0,0) 
	  				  
%>
</td></tr>
<tr>

            <td valign="top"> 
<form action="docs.asp" method="post" name="form" onSubmit="return checksubmit()">              
        <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
          <tr>
            <td valign="top" class="tb_tit">Verificar Documentos Publicados</td>
          </tr>
          <tr>
            <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                  	<td width="200" class="tb_subtit"><div align="center">UNIDADE </div></td>
                  	<td width="200" class="tb_subtit"><div align="center">CURSO </div></td>
                  	<td width="200" class="tb_subtit"><div align="center">ETAPA </div></td>
                  	<td width="200" class="tb_subtit"><div align="center">TURMA </div></td>
                  	<td width="200" class="tb_subtit"><div align="center">PASTAS PUBLICADAS</div></td>
                    </tr>
                  <tr>
                  	<td width="200"><div align="center">
                  		<select name="unidade" class="borda" onChange="recuperarCurso(this.value)">
                  			<option value="999990"></option>
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
                  	<td width="200"><div align="center">
                  		<div id="divCurso">
                  			<select class="borda">
                  				</select>
                  			</div>
                  		</div></td>
                  	<td width="200"><div align="center">
                  		<div id="divEtapa">
                  			<select class="borda">
                  				</select>
                  			</div>
                  		</div></td>
                  	<td width="200"><div align="center">
                  		<div id="divTurma">
                  			<select class="borda">
                  				</select>
                  			</div>
                  		</div></td>
                  	<td width="200"><div align="center">
                  		<%
		
		Set RSt = Server.CreateObject("ADODB.Recordset")
		SQLt = "SELECT COUNT(CO_Pasta_Doc) AS TOTAL_PASTAS FROM TB_Tipo_Pasta_Doc where ((DA_Expira NOT BETWEEN #01/01/1900# AND #"&data_expira&"#) AND IN_Expira= TRUE) or IN_Expira= FALSE"
		RST.Open SQLt, CON0
		
		Set RS_doc = Server.CreateObject("ADODB.Recordset")
		SQL_doc = "SELECT CO_Pasta_Doc, NO_Pasta FROM TB_Tipo_Pasta_Doc where ((DA_Expira NOT BETWEEN #01/01/1900# AND #"&data_expira&"#) AND IN_Expira= TRUE) or IN_Expira= FALSE order by NO_Pasta Asc"
		RS_doc.Open SQL_doc, CON0

		if RS_doc.eof then%>
                  		<font class="style1"> 
                  			<%response.Write("<br><br><br><br><br>N�o existem documentos cadastrados!")%>
                  			</font>
                  		<%ELSE%>		 					
                  		<select name="pasta" class="borda" id="pasta">
                  			<%
				conta_registros=0
				WHILE NOT RS_doc.eof
				total_pastas = RST("TOTAL_PASTAS")
				cod_tp_doc = RS_doc("CO_Pasta_Doc")		
				nom_tp_doc = RS_doc("NO_Pasta")	
								
				total_pastas=total_pastas*1
				if total_pastas = 1 then
					selected="selected"
				else
					selected=""
					if conta_registros=0 then
						response.Write("<option value=""nulo"" selected=""selected""></option>")
					end if					
				end if					
				%>
                  			<option value="<%response.Write(cod_tp_doc)%>" <%response.Write(selected)%>><%response.Write(nom_tp_doc)%></option>
                  			<%
				conta_registros=conta_registros+1	
				RS_doc.MOVENEXT
				WEND
				%>
                  			</select>
                  		<%END if%>						
                  		</div></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td>
                  <hr>
                    </td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td width="33%">&nbsp;</td>
		<td width="34%">&nbsp;</td>
		<td width="33%" align="center"><input name="button" type="submit" class="botao_prosseguir" id="button" value="Prosseguir"></td>
	</tr>
</table>
</td>
              </tr>
            </table></td>
          </tr>
        </table></form>
              </td>
  </tr>
		  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
        </table>
 <script type="text/javascript">
    function setUnidade(p_val){

      document.forms[3].unidade.options[1].selected = "true";
      recuperarCurso(p_val);

    } 
  setUnidade(1);
 </script> 
</body>
</html>
<%
if opt="a" then
			call GravaLog (chave,outro)
end if
If Err.number<>0 then
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