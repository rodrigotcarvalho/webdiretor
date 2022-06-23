<!--#include file="../../inc/caminhos.asp"-->
<!--#include file="../../inc/funcoes.asp"-->
<%nivel=2
tp=session("tp")
nome = session("nome") 
co_user = session("co_user")
cod= Session("aluno_selecionado")
Session("aluno_selecionado") = cod

	Set CON = Server.CreateObject("ADODB.Connection")
 	ABRIR = "DBQ="& CAMINHO_wf& ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set RSAL = Server.CreateObject("ADODB.Recordset")
	SQLAL = "SELECT ST_Matricula FROM TB_Alunos_Autorizados where CO_Matricula = "&cod
	RSAL.Open SQLAL, CON

	if not RSAL.EOF THEN
		if RSAL("ST_Matricula") = "L" then

			wrk_mensagem = "&nbsp;&nbsp;&nbsp;Próximos passos: <br>&nbsp;<br>"  
			wrk_mensagem = wrk_mensagem&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1) O responsável financeiro deverá imprimir o Contrato e o Adendo, assim como o boleto e " 
			'wrk_mensagem = wrk_mensagem&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; efetuando seu pagamento na rede bancária.<br>&nbsp;<br>" 	
			wrk_mensagem = wrk_mensagem&" efetuar o pagamento na rede bancária.<br>&nbsp;<br>"					
			wrk_mensagem = wrk_mensagem&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2) Encaminhar à Secretaria da Unidade de estudo de seu filho o Contrato e o Adendo devidamente assinados pelo responsável financeiro. <br>&nbsp;<br>"  								
			
			wrk_mensagem = wrk_mensagem&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;3) Após identificar o pagamento bancário, a escola deverá enviar, dentro de um prazo de cinco dias, o Kit matrícula contendo (agenda, calendário e material), contendo os Contratos" 

			'wrk_mensagem = wrk_mensagem&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; assinados pela Escola, assim como a relação de livros e cadernos (Ens. Fundamental II), efetivando, assim, a matrícula para 2016." 	
			wrk_mensagem = wrk_mensagem&" assinados pela Escola, assim como a relação de livros e cadernos (Ens. Fundamental II), efetivando, assim, a matrícula para "&Session("ano_letivo")+1&"." 			
					
		else

			wrk_mensagem = "&nbsp;&nbsp;&nbsp;Prezado Sr. Responsável: <br>&nbsp;<br>" 
			wrk_mensagem = wrk_mensagem&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Para dar continuidade ao processo de renovação de matrícula de seu filho, solicitamos " 	
			if escola = "dinamis" then			
				'wrk_mensagem = wrk_mensagem&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; seu comparecimento ao Centro Administrativo (Rua Muniz Barreto, 460 - Botafogo), <br>" 
				wrk_mensagem = wrk_mensagem&" seu comparecimento ao Centro Administrativo (Rua Muniz Barreto, 460 - Botafogo), " 					
			else
				'wrk_mensagem = wrk_mensagem&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; seu comparecimento à secretaria da Unidade Stockler(Rua General Rabelo, 56 - Gávea), <br>" 
				wrk_mensagem = wrk_mensagem&" seu comparecimento à secretaria da Unidade Stockler(Rua General Rabelo, 56 - Gávea), " 											
			end if	
			'wrk_mensagem = wrk_mensagem&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; para verificar a existência de pendências financeiras. <br>&nbsp;<br>" 
			wrk_mensagem = wrk_mensagem&" para verificar a existência de pendências financeiras. <br>&nbsp;<br>" 				
			wrk_mensagem = wrk_mensagem&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Atenciosamente <br>&nbsp;<br>" 	
			wrk_mensagem = wrk_mensagem&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; a Direção " 
			
		end if
	
	else
		wrk_mensagem = "Aluno não localizado"
	end if		

  
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Web Família</title>
<link href="../../estilo.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_popupMsg(msg) { //v1.0
  alert(msg);
}

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
						
						
						 function GravaRematricula()
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "grava_rematricula.asp", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
alert(resultado_c)
                                                           }
                                               }
 
                                               oHTTPRequest.send("cod_aluno=" + <%response.Write(cod)%>);
                                   }
 
 

//-->
</script>

</head>

<body >
<table width="1000" height="1038" border="0" align="center" cellpadding="0" cellspacing="0" class="tabela_principal">
  <%
			ano = DatePart("yyyy", now)
			mes = DatePart("m", now) 
			dia = DatePart("d", now) 
			hora = DatePart("h", now) 
			min = DatePart("n", now) 
select case mes
 case 1 
 mes = "janeiro"
 case 2 
 mes = "fevereiro"
 case 3 
 mes = "março"
 case 4
 mes = "abril"
 case 5
 mes = "maio"
 case 6 
 mes = "junho"
 case 7
 mes = "julho"
 case 8 
 mes = "agosto"
 case 9 
 mes = "setembro"
 case 10 
 mes = "outubro"
 case 11 
 mes = "novembro"
 case 12 
 mes = "dezembro"
end select

data = dia &" / "& mes &" / "& ano
data= FormatDateTime(data,1) 			

			horario = hora & ":"& min%>
  <tr>
    <td height="998"><table width="200" height="998" border="0" cellpadding="0" cellspacing="0">
          <!--DWLayoutTable-->
                  <tr valign="bottom"> 
          <td height="120" colspan="3"> 
              <%call cabecalho(nivel)%>
            </td>
          </tr>
          <tr class="tabela_menu"> 
            <td width="172" height="144" rowspan="4" valign="top" class="tabela_menu"><p>&nbsp;</p>
              <% call menu_lateral(nivel)%>
              <p>&nbsp;</p></td>
            <td width="640" height="12" nowrap="nowrap"><p class="style1">&nbsp;&nbsp;Ol&aacute; 
                <span class="style2">
                <%response.Write(nome)%>
                </span> , &uacute;ltimo acesso dia 
                <% Response.Write(session("dia_t")) %>
                &agrave;s 
                <% Response.Write(session("hora_t")) %>
              </p></td>
            <td width="188"><p align="right" class="style1"> 
                <%response.Write(data)%>
              </p></td>
          </tr>
          <tr class="tabela_menu"> 
            <td height="5" colspan="2"><p><img src="../../img/linha-pontilhada_grande.gif" alt="" width="828" height="5" /></p></td>
          </tr>
      <tr class="tabela_menu">
        <td height="19" colspan="2">&nbsp;</td>
      </tr>		  
          <tr class="tabela_menu"> 
            
          <td height="832" colspan="2" valign="top"><img src="../../img/rematricula.jpg" width="700" height="30"> 
            <table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="100%" valign="top">
                      <table width="100%" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td height="100" colspan="5" class="style1"><%Response.Write(wrk_mensagem)%></td>
                  </tr>
                  <tr>
                    <td width="7%">&nbsp;</td>
                    <td width="40%">&nbsp;</td>
                    <td width="6%">&nbsp;</td>
                    <td width="40%">&nbsp;</td>
                    <td width="7%">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="7%">&nbsp;</td>
                    <td width="40%">
                    <%if RSAL("ST_Matricula") = "L" then%>
                    <table width="200" height="20" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr> <%mes=1%>
                              <td class="tb_tit"><div align="center"><a href="../../relatorios/gera_boleto.asp?c=<%=cod%>&amp;tp=rematricula"  class="impressao" onClick="javascript:GravaRematricula();">Emitir boleto</a></div></td>
                            </tr>
                          </table></td>
                    <td width="6%">&nbsp;</td>
                    <td width="40%"><table width="200" height="20" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td class="tb_tit"><div align="center"><a href="seleciona.asp" class="impressao">Emitir contrato</a></div></td>
                            </tr>
                          </table>
                      <%end if%>    
                          </td>
                    <td width="7%">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="7%" height="500">&nbsp;</td>
                    <td width="40%" height="500">&nbsp;</td>
                    <td width="6%" height="500">&nbsp;</td>
                    <td width="40%" height="500">&nbsp;</td>
                    <td width="7%" height="500">&nbsp;</td>
                  </tr>
                </table>
                   </td>
                </tr>			
              </table>
            </td>
          </tr>
        </table></td>
  </tr>
  <tr>
    <td width="100%" height="74" valign="top"><img src="../../img/rodape.jpg" width="1000" height="74" /></td>
  </tr>
</table>
</body>
</html>