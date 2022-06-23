<%ano_letivo=request.QueryString("opt")

session("ano_letivo_upload")=ano_letivo%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
<!--
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
						
						
						 function gravarTipo(tpTipo)
                                   {
// Criação do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicitação HTTP. O primeiro parâmetro informa o método post/get
// O segundo parâmetro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicitação síncrona, o parâmetro deve ser false
                                               oHTTPRequest.open("post", "grava.asp?opt=t", true);
// Para solicitações utilizando o método post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A função abaixo é executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto já completou a solicitação
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto é gerado no arquivo executa.asp e colocado no div
                                                                       document.all.divCurso.innerHTML = oHTTPRequest.responseText;
                                                           }
                                               }
// Abaixo é enviada a solicitação. Note que a configuração
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("tp_pub=" + tpTipo);
                                   }

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>							   
<link href="estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../../img/menu_r1_c2_f3.gif','../../../../img/menu_r1_c2_f2.gif','../../../../img/menu_r1_c2_f4.gif','../../../../img/menu_r1_c4_f3.gif','../../../../img/menu_r1_c4_f2.gif','../../../../img/menu_r1_c4_f4.gif','../../../../img/menu_r1_c6_f3.gif','../../../../img/menu_r1_c6_f2.gif','../../../../img/menu_r1_c6_f4.gif','../../../../img/menu_r1_c8_f3.gif','../../../../img/menu_r1_c8_f2.gif','../../../../img/menu_r1_c8_f4.gif','../../../../img/menu_direita_r2_c1_f3.gif','../../../../img/menu_direita_r2_c1_f2.gif','../../../../img/menu_direita_r2_c1_f4.gif','../../../../img/menu_direita_r4_c1_f3.gif','../../../../img/menu_direita_r4_c1_f2.gif','../../../../img/menu_direita_r4_c1_f4.gif','../../../../img/menu_direita_r6_c1_f3.gif','../../../../img/menu_direita_r6_c1_f2.gif','../../../../img/menu_direita_r6_c1_f4.gif')">
<table width="1000" height="170" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
    <td width="1000" valign="top"> <table width="1000" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
        <tr> 
          <td width="1000"><FORM METHOD="POST" ENCTYPE="multipart/form-data" ACTION="grava_arquivo.asp?opt=f&al=<%=ano_letivo%>" target="_parent">
              <table width="1000" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
                <tr bgcolor="#FFFFFF" background="../../../../img/fundo_interno.gif"> 
                  <td height="20"> <div align="right"><font class="form_dado_texto"> 
                      Tipo de Documento: </font></div></td>
                  <td height="20"> <select name="tipo_doc" class="borda" onChange="gravarTipo(this.value)">
                      <%if session("tipo_arquivo_upl")="1" then%>
                      <option value="0"></option>
                      <option value="1" selected>Avisos e Circulares</option>
                      <option value="2">Provas e Gabaritos</option>
                      <%elseif session("tipo_arquivo_upl")="2" then%>
                      <option value="0"></option>
                      <option value="1">Avisos e Circulares</option>
                      <option value="2" selected>Provas e Gabaritos</option>
					  <%else
					  session("tipo_arquivo_upl")="0" %>
                      <option value="0" selected></option>
                      <option value="1">Avisos e Circulares</option>
                      <option value="2">Provas e Gabaritos</option>
                      <%end IF%>
                    </select> </td>
                </tr>
                <tr> 
                  <td  width="350"> <div align="right"><font class="form_dado_texto">Arquivo 
                      1: </font></div></td>
                  <td width="650"> <INPUT TYPE=FILE SIZE=60 NAME="FILE1" class="borda"></td>
                </tr>
                <tr> 
                  <td width="350"> <div align="right"><font class="form_dado_texto">Arquivo 
                      2: </font></div></td>
                  <td width="650"> <INPUT TYPE=FILE SIZE=60 NAME="FILE2" class="borda"></td>
                </tr>
                <tr> 
                  <td width="350"> <div align="right"><font class="form_dado_texto">Arquivo 
                      3: </font></div></td>
                  <td width="650"> <INPUT TYPE=FILE SIZE=60 NAME="FILE3" class="borda"></td>
                </tr>
                <tr> 
                  <td width="350"> <div align="right"><font class="form_dado_texto">Arquivo 
                      4: </font></div></td>
                  <td width="650"> <INPUT TYPE=FILE SIZE=60 NAME="FILE4" class="borda"></td>
                </tr>
                <tr> 
                  <td width="350"> <div align="right"><font class="form_dado_texto">Arquivo 
                      5: </font></div></td>
                  <td width="650"> <INPUT TYPE=FILE SIZE=60 NAME="FILE5" class="borda"></td>
                </tr>
                <tr> 
                  <td colspan="2"><hr width="1000"></td>
                </tr>
                <tr> 
                  <td colspan="2"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td width="33%"> <div align="center"> 
                            <input name="SUBMIT5" type=button class="borda_bot3" onClick="MM_goToURL('parent','http://www.simplynet.com.br/stockler/wf/apoio/ctl/pub/docs.asp?opt=f&pagina=1&v=s');return document.MM_returnValue" value="Voltar">
                          </div></td>
                        <td width="34%"> <div align="center"></div></td>
                        <td width="33%"> <div align="center"> 
                            <input name="SUBMIT" type=SUBMIT class="borda_bot2" value="Upload!">
							
                          </div></td>
                      </tr>
                    </table></td>
                </tr>
              </table>
            </form></td>
        </tr>
      </table></td>
  </tr>
</table>

</body>
</html>