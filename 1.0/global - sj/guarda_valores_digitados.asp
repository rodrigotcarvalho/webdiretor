<%
opt=request.querystring("opt")
if opt="valor1" then
	valor1=Request.Form("valor1")
	session("valor1")=valor1
elseif opt="valor2" then
	valor2=Request.Form("valor2")
	session("valor2")=valor2
elseif opt="valor3" then
	valor3=Request.Form("valor3")
	session("valor3")=valor3
elseif opt="valor4" then
	valor4=Request.Form("valor4")
	session("valor4")=valor4
elseif opt="valor5" then
	valor5=Request.Form("valor5")
	session("valor5")=valor5
elseif opt="valor6" then
	valor6=Request.Form("valor6")
	session("valor6")=valor6
elseif opt="valor7" then
	valor7=Request.Form("valor7")
	session("valor7")=valor7
elseif opt="valor8" then
	valor8=Request.Form("valor8")
	session("valor8")=valor8
elseif opt="valor9" then
	valor9=Request.Form("valor9")
	session("valor9")=valor9
elseif opt="valor10" then
	valor10=Request.Form("valor10")
	session("valor10")=valor10		
end if

%>
