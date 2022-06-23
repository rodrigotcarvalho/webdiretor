<%
ano_letivo = session("ano_letivo")
session("ano_letivo") = ano_letivo

ordenacao= request.QueryString("obr")

response.Redirect("../../../../relatorios/swd310.asp?opt="&ordenacao) 
%>