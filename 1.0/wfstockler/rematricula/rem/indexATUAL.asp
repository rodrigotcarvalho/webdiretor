<!--#include file="../../inc/caminhos.asp"-->
<!--#include file="../../inc/funcoes.asp"-->
<%nivel=2
tp=session("tp")
nome = session("nome") 
co_user = session("co_user")
cod= Session("aluno_selecionado")
Session("aluno_selecionado") = cod
opt= request.QueryString("opt")
versao_contrato_adendo = request.QueryString("tp_contrato")

	Set CON = Server.CreateObject("ADODB.Connection")
 	ABRIR = "DBQ="& CAMINHO_wf& ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set RSAL = Server.CreateObject("ADODB.Recordset")
	SQLAL = "SELECT ST_Matricula FROM TB_Alunos_Autorizados where CO_Matricula = "&cod
	RSAL.Open SQLAL, CON
	
	SQL2 = "select * from TB_Usuario where CO_Usuario = " & cod 
	set RS2 = CON.Execute (SQL2)
	
nome_aluno= RS2("NO_Usuario")	
	

	if not RSAL.EOF THEN
		if session("ano_letivo") >=2020 then
			wrk_mensagem = "ATENÇÃO:<br>&nbsp;<br>"
						
			
			wrk_mensagem = wrk_mensagem&"1) Abaixo, temos duas colunas onde estão disponíveis as duas propostas de planos de pagamento. Escolha a opção 1 para pagamento da 1ª Parcela de Mensalidade de forma única, à vista (Outubro/20) ; Escolha a opção 2 para pagamento da 1ª Parcela de Mensalidade em 2 vezes ( Outubro/20 e Janeiro/21);<br>&nbsp;<br>"

wrk_mensagem = wrk_mensagem&"2)      logo após a escolha do plano de pagamento, solicitamos que clique primeiramente no botão emitir boleto e, após a sua impressão, clicar no botão assinar contrato;<br>&nbsp;<br>"

wrk_mensagem = wrk_mensagem&"3)      após clicar no botão assinar contrato, a empresa certificadora enviará uma mensagem para o seu e-mail e solicitará algumas informações básicas para serem comparadas com o nosso banco de dados;<br>&nbsp;<br>"

wrk_mensagem = wrk_mensagem&"4)      após a confirmação das informações, o sistema enviará um novo e-mail com um código TOKEN a ser digitado para a validação e assinatura do contrato;<br>&nbsp;<br>"

wrk_mensagem = wrk_mensagem&"5)     IMPORTANTE: a matrícula só será confirmada após o pagamento da 1ª parcela e o recebimento do contrato assinado digitalmente pela Escola, via e-mail.<br>&nbsp;<br>"
		else
	
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
 
 <%
 onload =""
 if opt="ok" then

	response.Write(versao_contrato_adendo)
  onload = "GravaRematricula();"
 end if%>
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>

</head>

<body onLoad="<%response.Write(onload)%>">
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
                    <td height="20" colspan="5" class="tb_tit">
                        Dados Escolares</td>
                  </tr>
                <tr>
                    <td height="33" colspan="5" class="style1">
                    <table width="100%" border="0" cellspacing="0">
                      <tr> 
                        <td width="19%" height="10"> <div align="right"><font class="style3"> 
                            Matr&iacute;cula: </font></div></td>
                        <td width="9%" height="10"><font class="style1"> 
                          <input name="cod" type="hidden" value="<%=cod%>">
                          <%response.Write(cod)%>
                          </font></td>
                        <td width="6%" height="10"> <div align="right"><font class="style3"> 
                            Nome: </font></div></td>
                        <td width="66%" height="10"><font class="style1"> 
                          <%response.Write(nome_aluno)%>
                          <input name="nome2" type="hidden" class="textInput" id="nome2"  value="<%response.Write(nome_aluno)%>" size="75" maxlength="50">
                          &nbsp;</font></td>
                      </tr>
                    </table>
                    </td>
                  </tr>                                        
                  <tr>
                    <td height="100" colspan="5" class="style1"><blockquote>
                      <p>
                        <%Response.Write(wrk_mensagem)%>
                      </p>
                    </blockquote></td>
                  </tr>

                  <%
    if opt<>"ok" then
		Set RSrem = Server.CreateObject("ADODB.Recordset")
		SQLrem = "SELECT * FROM TB_Aunos_Rematriculados where CO_Matricula_Escola="&cod
		RSrem.Open SQLrem, CON
	
		if RSrem.EOF then
			versao_contrato_adendo = "nulo"
		else
			versao_contrato_adendo = RSrem("TP_Contrato")		
		end if
	end if	
				  
				  if RSAL("ST_Matricula") = "L" then
				  
				  		if session("ano_letivo")>=2020 then
							
							if opt="ok" or versao_contrato_adendo<>"nulo" then
						
						%>
                                              <tr>
                        <td width="7%">&nbsp;</td>
                        <td colspan="3" align="center" class="tb_tit">Op&ccedil;&atilde;o <%response.write(versao_contrato_adendo)%></td>
                        <td width="7%">&nbsp;</td>
                      </tr>
                      <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                          <tr>
                            <td width="7%" height="50">&nbsp;</td>
                            <td height="50" colspan="3" align="center">
                            
                            <strong>
                            
                            </strong>
        <input name="boleto" type="submit" class="bt_contrato" id="boleto" value="Emitir Boleto" onClick="MM_goToURL('parent','../../relatorios/gera_boleto.asp?c=<%=cod%>&amp;tp=rematricula&amp;contrato=<%response.write(versao_contrato_adendo)%>');return document.MM_returnValue;"><br></td>
                            <td width="7%" height="50">&nbsp;</td>
                          </tr>

                       <%ELSE%> 
                      <tr>
                        <td width="7%">&nbsp;</td>
                        <td width="40%" align="center" class="tb_tit">Op&ccedil;&atilde;o 1</td>
                        <td width="6%" align="center">&nbsp;</td>
                        <td width="40%" align="center" class="tb_tit">Op&ccedil;&atilde;o 2</td>
                        <td width="7%">&nbsp;</td>
                      </tr>
                      <tr>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>                        
                          <tr>
                            <td height="50">&nbsp;</td>
                            <td height="50" align="center"><input name="assinar1" type="submit" class="bt_contrato" id="assinar1" value="Assinar Contrato" onClick="MM_goToURL('parent','Seleciona.asp?opt=1');return document.MM_returnValue"></td>
                            <td height="50">&nbsp;</td>
                            <td height="50" align="center"><input name="assinar2" type="submit" class="bt_contrato" id="assinar2" value="Assinar Contrato" onClick="MM_goToURL('parent','Seleciona.asp?opt=2');return document.MM_returnValue"></td>
                            <td height="50">&nbsp;</td>
                          </tr>
                          <tr>
                            <td width="7%" height="50">&nbsp;</td>
                            <td width="40%" height="50" align="center">
                            
                            <strong>
                            
                            </strong>
        <input name="boleto1" type="submit" class="bt_contrato" id="boleto1" value="Emitir Boleto" onClick="MM_goToURL('parent','../../relatorios/gera_boleto.asp?c=<%=cod%>&amp;tp=rematricula&amp;contrato=1');return document.MM_returnValue;"><br></td>
                            <td width="6%" height="50">&nbsp;</td>
                            <td width="40%" height="50" align="center"><strong> </strong>
                              <input name="boleto2" type="submit" class="bt_contrato" id="boleto2" value="Emitir Boleto"  onClick="MM_goToURL('parent','../../relatorios/gera_boleto.asp?c=<%=cod%>&amp;tp=rematricula&amp;contrato=2');return document.MM_returnValue;">
                            <br></td>
                            <td width="7%" height="50">&nbsp;</td>
                          </tr>
                      <%
					  	END IF
					  else%>
                      <tr>
                        <td width="7%">&nbsp;</td>
                        <td width="40%"><br>
                        <table width="200" height="20" border="0" align="center" cellpadding="0" cellspacing="0">
                                <tr> <%mes=1%>
                                  <td class="tb_tit"><div align="center"><a href="../../relatorios/gera_boleto.asp?c=<%=cod%>&amp;tp=rematricula"  class="impressao" onClick="javascript:GravaRematricula();">Emitir boleto</a></div></td>
                                </tr>
                        </table></td>
                        <td width="6%">&nbsp;</td>
                        <td width="40%"><table width="200" height="20" border="0" align="center" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td class="tb_tit"><div align="center">
                                   <%if opt="ok" then%>
                                <font class="impressao">Emitir contrato</font>
                                 <%else%> 
                                <a href="seleciona.asp" class="impressao">Emitir contrato</a>                                                        
                                 <%end if%>
                                  </div></td>
                                </tr>
                              </table>
     
                        </td>
                        <td width="7%">&nbsp;</td>
                      </tr>
                  <%	end if
				  end if%>                     
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
        </table>
      <p class="confirmar">&nbsp;</p>
      <p class="tb_tit">&nbsp;</p></td>
  </tr>
  <tr>
    <td width="100%" height="74" valign="top"><img src="../../img/rodape.jpg" width="1000" height="74" /></td>
  </tr>
</table>
</body>
</html>