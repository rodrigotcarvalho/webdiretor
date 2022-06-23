<!--#include file="caminhos.asp"-->
<%
escola=session("escola")

	Set CON_wr = Server.CreateObject("ADODB.Connection") 
	ABRIR_wr = "DBQ="& CAMINHO_ctrle & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_wr.Open ABRIR_wr
	
	Set RS1 = Server.CreateObject("ADODB.Recordset")
	consulta1 = "select * from TB_Cliente where CO_Cliente="&escola
	set RS1 = CON_wr.Execute (consulta1)
		
if RS1.EOF then
nome_escola=""
else
nome_escola=RS1("NO_Cliente")
end if
sistema_local=session("sistema_local")

Select Case sistema_local

case "WR"
sistema_nome="Web Diretor"

case "WA"
sistema_nome="Web Acadêmico"

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
sistema_nome="Web Informática"

case "WF"
sistema_nome="Web Família"
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
        novamente realizar a operação desejada</a>. </font></p>
      <p align="left"><font size="2" face="Arial, Helvetica, sans-serif">Caso 
        o problema volte a ocorrer, por favor, anote as informações abaixo e entre 
        em contato com Simply Net Informação e Tecnologia LTDA.Pelo telefone 2232-5541 
        / E-Mail: <a href="mailto:suporte@simplynet.com.br">suporte@simplynet.com.br</a> 
        informando os dados fornecidos abaixo: </font></p>
      <p align="left"><font size="2" face="Arial, Helvetica, sans-serif">Sistema: 
        <%'response.Write(nome_escola)%> - <%response.Write(sistema_nome)%><br>
         
        <%
		
errnumb=session("errnumb")
errdesc=session("errdesc")
errfile=session("errfile")


Response.Write "O número do erro é: " & errnumb & "<BR>"
Response.Write "A descrição fornecida é: " & errdesc& "<BR>"
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
