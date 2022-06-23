<!--#include file="../../../../inc/caminhos.asp"-->
<% 
opt=request.QueryString("opt")
if opt="f" then
		Set upl = Server.CreateObject("SoftArtisans.FileUp")
		response.Write(CAMINHO_upload)
		 upl.Path = CAMINHO_upload
		
							
		file1 = upl.Form("FILE").ShortFileName 

	
	
		If Not file1 = "" Then
			file1 = file1
			upl.Form("FILE").Save
			contarq=contarq+1  
		Else
			um="no"
		End If
					
		Set upl = Nothing 
						

else


%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/JavaScript">
function upload()  
{
   var f=document.forms[0]; 
      f.submit(); 
	  
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
document.all.divEtapa.innerHTML ="<select class=borda></select>"
document.all.divTurma.innerHTML = "<select class=borda></select>"
//recuperarEtapa()
                                                           }
                                               }
// Abaixo é enviada a solicitação. Note que a configuração
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


	   
						 function incluirAnexo(aTipo)
                                   {
// Criação do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicitação HTTP. O primeiro parâmetro informa o método post/get
// O segundo parâmetro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicitação síncrona, o parâmetro deve ser false
                                               oHTTPRequest.open("post", "executa.asp?opt=anx", true);
// Para solicitações utilizando o método post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A função abaixo é executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto já completou a solicitação
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto é gerado no arquivo executa.asp e colocado no div
                                                                       var resultado_A= oHTTPRequest.responseText;
resultado_A = resultado_A.replace(/\+/g," ")
resultado_A = unescape(resultado_A)
document.all.divAnexo.innerHTML = resultado_A																	   
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
                        
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0;
	margin-right: 0px;
	margin-bottom: 0px;
	background-color: #FFF;
}
-->
</style></head>

<body>

<table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
  <tr>
    <td width="653" valign="top"><a href="#" onClick="MM_callJS('incluirAnexo()')" >Anexo</a></td>
  </tr>
  <tr>
    <td valign="top"><FORM name="form_upload" id="form_upload" METHOD="POST" ENCTYPE="multipart/form-data" ACTION="upload.asp?opt=f" target="_parent">
      <div id="divAnexo"></div>
    </FORM></td>
  </tr>
</table>
</body>
</html>
<%end if%>
