<%
opt=request.querystring("opt")
if opt="t" then
session("tipo_arquivo_upl")=Request.Form("tp_pub")
end if
%>