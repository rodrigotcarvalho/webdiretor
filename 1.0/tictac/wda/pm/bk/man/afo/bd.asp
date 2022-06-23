<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"--> 
<!--#include file="../../../../inc/funcoes2.asp"-->
      <html>
      <head>
      <title>Web Diretor</title>
      <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
      <script language="JavaScript" type="text/JavaScript">
<!--
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
      </head>
      <% 

opt = request.QueryString("opt")
nvg=request.QueryString("nvg")
ano_letivo = session("ano_letivo") 
chave=nvg
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
sistema_local=session("sistema_local")
session("sistema_local")=sistema_local
acesso= request.Form("acesso")
cod_cons = request.Form("cod_cons")
nome = request.Form("nome")
apelido = request.Form("apelido")
rua = request.Form("rua")
numero = request.Form("numero")
complemento = request.Form("complemento")
cep = request.Form("cep")
estado = request.Form("estado")
cidade = request.Form("cidade")
bairro = request.Form("bairro")
telefones = request.Form("telefones")
cnpj = request.Form("cnpj")
email = request.Form("email")
contatos = request.Form("contatos")
ativo = request.Form("ativo")
'response.Write(ativo&"<BR>")

if ativo="sim" then
	situacao= TRUE
ELSE
	situacao= FALSE
END IF

		Set CON9 = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_ax & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON9.Open ABRIR

if cep="" or isnull(cep) then
	cep="0-0"
end if

cep_left = LEFT(cep,5)
cep_right = RIGHT(cep,3)
cep_format = cep_left&"-"&cep_right
ceplit = split(cep_format,"-")
cepin=ceplit(0)
cepfim=ceplit(1)
			
		
If IsNumeric(cepin) Then
	If IsNumeric(cepfim) Then
		valicp = 0
	else
		valicp = 1
	end if
else
	valicp = 1
end if	


if numero = "" then
	numero = 0
elseif IsNumeric(numero) Then
	valinm = 0
else
	valinm = 1
end if



if valinm = 1 then
	if opt = "inc" then
		response.redirect("altera.asp?ori=02&nvg="&nvg&"&z=2&e=nb&cod_cons="&cod_cons&"&nome="&nome&"&apelido="&apelido&"&rua="&rua&"&numero="&numero&"&complemento="&complemento&"&cep="&cep&"&estado="&estado&"&cidade="&cidade&"&bairro="&bairro&"&telefones="&telefones&"&cnpj="&cnpj&"&email="&email&"&contatos="&contatos&"&ativo="&ativo)
	elseif opt = "alt" then
		response.redirect("altera.asp?ori=02&nvg="&nvg&"&z=3&e=nb&cod_cons="&cod_cons&"&nome="&nome&"&apelido="&apelido&"&rua="&rua&"&numero="&numero&"&complemento="&complemento&"&cep="&cep&"&estado="&estado&"&cidade="&cidade&"&bairro="&bairro&"&telefones="&telefones&"&cnpj="&cnpj&"&email="&email&"&contatos="&contatos&"&ativo="&ativo)
	end if
elseif valicp = 1 then
	if opt = "inc" then
		response.redirect("altera.asp?ori=02&nvg="&nvg&"&z=2&e=cp&cod_cons="&cod_cons&"&nome="&nome&"&apelido="&apelido&"&rua="&rua&"&numero="&numero&"&complemento="&complemento&"&cep="&cep&"&estado="&estado&"&cidade="&cidade&"&bairro="&bairro&"&telefones="&telefones&"&cnpj="&cnpj&"&email="&email&"&contatos="&contatos&"&ativo="&ativo)
	elseif opt = "alt" then
		response.redirect("altera.asp?ori=02&nvg="&nvg&"&z=3&e=cp&cod_cons="&cod_cons&"&nome="&nome&"&apelido="&apelido&"&rua="&rua&"&numero="&numero&"&complemento="&complemento&"&cep="&cep&"&estado="&estado&"&cidade="&cidade&"&bairro="&bairro&"&telefones="&telefones&"&cnpj="&cnpj&"&email="&email&"&contatos="&contatos&"&ativo="&ativo)
	end if
else
	if opt = "inc" then
		
		if cep="0-0" then
			cep=NULL
		end if	
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT MAX(CO_Fornecedor) as Max_CO_Fornecedor FROM TB_Fornecedor"
		RS.Open SQL, CON9
		
		if RS.EOF then
			cod_cons = 1
		else
			ult_codigo = RS("Max_CO_Fornecedor")
			cod_cons = ult_codigo+1
		end if	
		
		
		Set RS = server.createobject("adodb.recordset")
		RS.open "TB_Fornecedor", CON9, 2, 2 'which table do you want open
		
		RS.addnew
		RS("CO_Fornecedor") = cod_cons
		RS("NO_Fornecedor") = nome
		RS("NO_Apelido_Fornecedor") = apelido
		RS("NO_Logradouro") = rua
		RS("NU_Logradouro") = numero
		RS("TX_Complemento_Logradouro") = complemento
		RS("CO_Bairro") = bairro
		RS("CO_Municipio") = cidade
		RS("SG_UF") = estado
		RS("CO_CEP") = cep_format
		RS("NUS_Telefones") = telefones
		RS("CO_CNPJ") = cnpj
		RS("TX_EMail") = email	
		RS("NO_Contatos") = contatos
		RS("IN_Ativo") = situacao
		RS.update
		  
		set RS=nothing
		
		
		call GravaLog (nvg,"Fornecedor de código "&cod_cons&" incluído")
		
		response.Redirect("index.asp?ori=02&opt=ok&nvg="&nvg&"&cod_cons="& cod_cons &"&co_usr_prof="&co_usr_prof&"&tx_login="&login)
		
	elseif opt = "alt" then		
		if cep="0-0" then
			cep=NULL
		end if
		
		sql_atualiza = "UPDATE TB_Fornecedor SET [NO_Fornecedor] = '"&nome&"',"
		sql_atualiza = sql_atualiza&"	[NO_Apelido_Fornecedor] = '"&apelido&"',"
		sql_atualiza = sql_atualiza&"	[NO_Logradouro] = '"&rua&"',"
		sql_atualiza = sql_atualiza&"	[NU_Logradouro] = "&numero&","
		sql_atualiza = sql_atualiza&"	[TX_Complemento_Logradouro] = '"&complemento&"',"
		sql_atualiza = sql_atualiza&"	[CO_Bairro] = "&bairro&","
		sql_atualiza = sql_atualiza&"	[CO_Municipio] = "&cidade&"," 
		sql_atualiza = sql_atualiza&"	[SG_UF] = '"&estado&"'," 
		sql_atualiza = sql_atualiza&"	[CO_CEP] = '"&cep_format&"',"
		sql_atualiza = sql_atualiza&"	[NUS_Telefones] = '"&telefones&"',"
		sql_atualiza = sql_atualiza&"	[CO_CNPJ] = '"&cnpj&"',"
		sql_atualiza = sql_atualiza&"	[TX_EMail]= '"&email&"', "
		sql_atualiza = sql_atualiza&"	[NO_Contatos]= '"&contatos&"', "	
		sql_atualiza = sql_atualiza&"	[IN_Ativo] ="&situacao
		sql_atualiza = sql_atualiza&" WHERE [CO_Fornecedor] = "& cod_cons 
		'response.Write(sql_atualiza)
		'response.End()
		Set RSup = CON9.Execute(sql_atualiza)
		
		'response.End()
		call GravaLog (nvg,"Dados cadastrais do fornecedor de código "&cod_cons&" alteradoa")
		
		response.Redirect("altera.asp?ori=01&nvg="&nvg&"&opt=ok&cod_cons="& cod_cons &"")
	
	end if

end if


%>
</html><%If Err.number<>0 then
errnumb = Err.number
errdesc = Err.Description
lsPath = Request.ServerVariables("SCRIPT_NAME")
arPath = Split(lsPath, "/")
GetFileName =arPath(UBound(arPath,1))
passos = 0
for way=0 to UBound(arPath,1)
passos=passos+1
next
seleciona1=passos-2
pasta=arPath(seleciona)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>