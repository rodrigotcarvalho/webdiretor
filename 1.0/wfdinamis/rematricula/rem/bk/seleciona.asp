<!--#include file="../../inc/caminhos.asp"-->
<!--#include file="../../inc/funcoes.asp"-->
<!--#include file="../../inc/bd_webfamilia.asp"-->
<!--#include file="../../inc/bd_alunos.asp"-->
<!--#include file="../../inc/bd_contato.asp"-->
<%nivel=2
tp=session("tp")
nome = session("nome") 
co_user = session("co_user")
cod= Session("aluno_selecionado")
Session("aluno_selecionado") = cod
ano_letivo = session("ano_letivo")
session("ano_letivo") = ano_letivo
opt= request.QueryString("opt")

	Set CON = Server.CreateObject("ADODB.Connection")
 	ABRIR = "DBQ="& CAMINHO_wf& ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR

	SQL2 = "select * from TB_Usuario where CO_Usuario = " & cod 
	set RS2 = CON.Execute (SQL2)

nome_aluno= RS2("NO_Usuario")	

vetorUcet = buscaUCET(cod,ano_letivo)
ucet = split(vetorUcet,"#!#")
modeloAdendo = modeloContratoAdendo(ucet(0),ucet(1),ucet(2),ucet(3),"A")

tipo_resp_fin = buscaTipoResponsavelFinanceiro(cod)	
vetorContato = buscaContato (cod, tipo_resp_fin)
dadosContato = split(vetorContato, "#!#")
dadosContato = split(vetorContato, "#!#")
nomeRespFin = Server.HTMLEncode(dadosContato(2))
emailRespFin  = dadosContato(8)

if modeloAdendo = "X" then
	response.redirect("../../relatorios/contratoAdendoPDF.asp?c="&cod&"&t=C")
end if
if vetorUcet <> "" then
	ucet = split(vetorUcet,"#!#")
	modeloContrato = modeloContratoAdendo(ucet(0),ucet(1),ucet(2),ucet(3),"C")
	
	modeloAdendo = modeloContratoAdendo(ucet(0),ucet(1),ucet(2),ucet(3),"A")
	vetor_modeloContrato = split(modeloContrato, "#!#")
	vetor_modeloAdendo = split(vetor_modeloAdendo, "#!#")
	
	'if para verificar se existe adendo
	if ubound(vetor_modeloAdendo)>=0 then
		if vetor_modeloAdendo(0) = "X" and ubound(vetor_modeloContrato)=0 then
			response.redirect("../../../../relatorios/contratoAdendoPDF.asp?c="&cod&"&t=C&dr=D")
		elseif modeloAdendo = "" then
			response.redirect("index.asp?opt=err2&cod="&cod)	
		end if
	end if	
else
		response.redirect("index.asp?opt=err1&cod="&cod)	
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
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
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
            <% if session("ano_letivo")>=2020 then%>
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
                          <%response.Write(Server.HTMLEncode(nome_aluno))%>
                          <input name="nome2" type="hidden" class="textInput" id="nome2"  value="<%response.Write(nome_aluno)%>" size="75" maxlength="50">
                          &nbsp;</font></td>
                      </tr>
                    </table>
                    </td>
                  </tr>                    
					<tr>
                    <td height="100" colspan="5" class="style1" style="padding:15"><span class="style2"><strong>Atencão!</strong></span>&nbsp;Sua escolha é a opção<strong> <%=opt%></strong>.<br>&nbsp;<br>

Após sua confirmação a empresa certificadora e a escola farão todo o controle das assinaturas e enviarão a seguir o contrato na integra para o email <%response.Write(emailRespFin)%> para assinatura de <%response.Write(nomeRespFin)%>.<br>&nbsp;<br>

Ao receber o email Clique no link visualizar para assinar, leia o contrato e ao final clique no botão assinar.<br>&nbsp;<br>

<strong>Confirma a assinatura do contrato <%=opt%>?</strong></td>
                  </tr> 
				   <tr>
                   	<td height="100" colspan="5" align="center" class="style1">
                    <input name="assinar1" type="submit" class="bt_contrato" id="assinar1" value="Assinar Contrato" onClick="MM_goToURL('parent','../../relatorios/contratoAdendoPDF.asp?c=<%=cod%>&amp;modelo=<%response.Write(vetor_modeloContrato(opt-1))%>&amp;t=C&amp;dr=R&amp;v=<%response.Write(opt)%>');return document.MM_returnValue">
                    </td>
              </tr>                             
            
            <%else%>
              <tr> 
                  <td width="100%" valign="top">
                      <table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <tr>
                            <td height="100" colspan="5" class="style1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Selecione o item que deseja imprimir</td>
                          </tr>
                          <tr>
                            <td>&nbsp;</td>
                            <td class="style1">
                                           <% for con = 0 to ubound(vetor_modeloContrato)%>
                            <ul><li><a href="../../relatorios/contratoAdendoPDF.asp?c=<%=cod%>&amp;modelo=<%response.Write(vetor_modeloContrato(con))%>&amp;t=C&amp;dr=R"><%response.Write(vetor_modeloContrato(con))%></a></li></ul>
                            
							<% next 
                                                    
                            for aden = 0 to ubound(vetor_modeloAdendo)%>
                            <ul><li><a href="../../relatorios/contratoAdendoPDF.asp?c=<%=cod%>&amp;modelo=<%response.Write(vetor_modeloAdendo(con))%>&amp;t=A&amp;dr=R"><%response.Write(vetor_modeloAdendo(aden))%></a></li></ul>
                            <%next%>
                    
                            </td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                            <td>&nbsp;</td>
                          </tr>
                          <tr>
                            <td height="500">&nbsp;</td>
                            <td height="500">&nbsp;</td>
                            <td height="500">&nbsp;</td>
                            <td height="500">&nbsp;</td>
                            <td height="500">&nbsp;</td>
                          </tr>
                	</table>
                   </td>
                </tr>	
                <%end if%>		
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
