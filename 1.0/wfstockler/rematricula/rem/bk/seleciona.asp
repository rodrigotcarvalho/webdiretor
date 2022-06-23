<!--#include file="../../inc/caminhos.asp"-->
<!--#include file="../../inc/funcoes.asp"-->
<!--#include file="../../inc/bd_webfamilia.asp"-->
<!--#include file="../../inc/bd_alunos.asp"-->
<%nivel=2
tp=session("tp")
nome = session("nome") 
co_user = session("co_user")
cod= Session("aluno_selecionado")
Session("aluno_selecionado") = cod
ano_letivo = session("ano_letivo")
session("ano_letivo") = ano_letivo

vetorUcet = buscaUCET(cod,ano_letivo)
ucet = split(vetorUcet,"#!#")
modeloAdendo = modeloContratoAdendo(ucet(0),ucet(1),ucet(2),ucet(3),"A")

if modeloAdendo = "X" then
	response.redirect("../../relatorios/contratoAdendoPDF.asp?c="&cod&"&t=C")
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
                    <td height="100" colspan="5" class="style1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Selecione o item que deseja imprimir</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td class="style1"><ul><li><a href="../../relatorios/contratoAdendoPDF.asp?c=<%=cod%>&amp;t=C">Contrato</a></li><li><a href="../../relatorios/contratoAdendoPDF.asp?c=<%=cod%>&amp;t=A">Adendo</a></li></ul></td>
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
