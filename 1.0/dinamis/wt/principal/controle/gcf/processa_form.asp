<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>
<%
if Request.Form("acao") = "Confirmar" then
	selecao = request.form("selecao")
	tp_compromisso  = request.form("tp_compromisso")
	compromissos  = request.form("compromissos")
	gera_prd_contrato = request.form("gera_prd_contrato")
	mes_de = request.form("mes_de")
	mes_ate = request.form("mes_ate")
	acao_compromissos = request.form("acao_compromissos")
	dia_vencimento = request.form("dia_vencimento")
	dia_util = request.form("dia_util")
	compromisso = request.form("compromisso")

	concatena_incluir  = selecao&"$"&tp_compromisso&"$"&compromissos&"$"&gera_prd_contrato&"$"&mes_de&"$"&mes_ate&"$"&dia_vencimento&"$"&dia_util&"$"&compromisso&"$"&acao_compromissos
elseif Request.Form("acao") = "Incluir" then
	concatena_incluir = request.form("selecao")
else
	concatena_compromissos = request.form("compromissos")
	
	vetor_concatena1 = split(concatena_compromissos,", ")
	if ubound(vetor_concatena1)>-1 then
		vetor_concatena2 = split(vetor_concatena1(0),"-")
		matric=vetor_concatena2(0)
		
		vetor_concatena3 = split(vetor_concatena2(1),"$")
		ano_contrato=vetor_concatena3(0)
		contrato = vetor_concatena3(1)
	end if
end if


 if Request.Form("acao") = "Excluir" then
	response.redirect("confirma.asp?opt=e&cc="&concatena_compromissos)
  elseif Request.Form("acao") = "Alterar" then
	response.redirect("alterar_contrato.asp?ac="&ano_contrato&"&nc="&contrato&"&mc="&matric)
  elseif Request.Form("acao") = "Incluir" then
	response.redirect("incluir.asp?ci="&concatena_incluir)
  elseif Request.Form("acao") = "Confirmar" then
	response.redirect("confirma.asp?opt=i&cc="&concatena_incluir)  
  end if

%>
<body>
</body>
</html>
