<!--#include file="caminhos.asp"-->
<%

sistema_local=session("sistema_local")

Select Case sistema_local

case "WR"
sistema_nome="Web Diretor"

case "WA"
sistema_nome="Web Acad�mico"

case "WN"
sistema_nome="Web Professor"

case "WS"
sistema_nome="Web Secretaria"

case "WT"
sistema_nome="Web Tesouraria"

case "WM"
sistema_nome="Web Marketing"

case "WD"
sistema_nome="Web Diretoria"

case "WI"
sistema_nome="Web Inform�tica"

case "WF"
sistema_nome="Web Fam�lia"
end select

%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<table width="650" height="300" border="3" align="center" cellspacing="1" bordercolor="#EEEEEE">
  <tr>
    <td bgcolor="#FFE8E8">
<p align="center"><font size="2" face="Arial, Helvetica, sans-serif"><img src="../img/pare.gif" width="28" height="25"></font></p>
      <p align="center"><font size="2" face="Arial, Helvetica, sans-serif"><strong>Houve 
        um erro na tentativa de acessar os arquivos do sistema. </strong></font></p>
      <p align="center"><font size="2" face="Arial, Helvetica, sans-serif"><a href="../inicio.asp">Tente 
        novamente realizar a opera��o desejada</a>. </font></p>
      <p align="left"><font size="2" face="Arial, Helvetica, sans-serif">Caso 
        o problema volte a ocorrer, por favor, anote as informa��es abaixo e entre 
        em contato com Simply Net Informa��o e Tecnologia LTDA.Pelo telefone 2232-5541 
        / E-Mail: <a href="mailto:suporte@simplynet.com.br">suporte@simplynet.com.br</a> 
        informando os dados fornecidos abaixo: </font></p>
      <p align="left"><font size="2" face="Arial, Helvetica, sans-serif">Sistema: 
        <%response.Write(nome_da_escola)%> - <%response.Write(sistema_nome)%><br>
         
        <%
		
errnumb=session("errnumb")
errdesc=session("errdesc")
errfile=session("errfile")


Response.Write "O n�mero do erro �: " & errnumb & "<BR>"
Response.Write "A descri��o fornecida �: " & errdesc& "<BR>"
Response.Write "O erro ocorreu no Arquivo: " & errfile& "<BR>"


Set objErr = Nothing
%>
        </font></p>
      </td>
  </tr>
</table>
<p>&nbsp; </p>
</body>
</html>
