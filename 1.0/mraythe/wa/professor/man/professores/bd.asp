<%On Error Resume Next%>


<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->

<!--#include file="../../../../inc/caminhos.asp"-->
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

cod_cons = request.Form("cod_cons")
'tp= request.Form("tp")
acesso= request.Form("acesso")
nome = request.Form("nome")
apelido = request.Form("apelido")
nasce = request.Form("nasce")
sexo = request.Form("sexo")
pais = request.Form("pais")
nacionalidade = request.Form("nacionalidade")
estadonat = request.Form("estadonat")
cidadenat = request.Form("cidadenat")
rua = request.Form("rua")
numero = request.Form("numero")
complemento = request.Form("complemento")
cep = request.Form("cep")
estadodom = request.Form("estadodom")
ciddom = request.Form("ciddom")
bairro = request.Form("bairro")
telefones = request.Form("telefones")
email = request.Form("email")
ativo = request.Form("ativo")
'response.Write(ativo&"<BR>")

if ativo="sim" then
situacao_professor=-1
situacao_usuario="L"
ELSE
situacao_professor=0
situacao_usuario="B"
END IF
'response.Write(situacao_professor&"-"&situacao_usuario&"<BR>")


nascimento = split (nasce,"/")

dia = nascimento(0)
mes = nascimento(1)
ano = nascimento(2)


If IsNumeric(dia) Then
	If IsNumeric(mes) Then
		If IsNumeric(ano) Then

if (ano > 1900)or(ano < 2000) then       

	select case mes 
              
			case 01,03,05,07,08,10,12 
				if  (dia <= 31)  then
				valilidata = 0
				else 
				validata =1
          		end if			           
            case 04,06,09,11 
				if  (dia <= 30)  then
				valilidata = 0
				else 
				validata =1
          		end if                       
			case 02                   
' Validando ano Bissexto / fevereiro / dia */                    
			if ((ano/4 = 0) or (ano/100 = 0) or (ano/400 = 0)) then
				bissexto = 1               

				if ((bissexto = 1) and (dia <= 29))  then                   
				valilidata = 0                 
				else
				validata =1
          		end if   
				if ((bissexto <> 1) and (dia <= 28))  then               
				valilidata = 0
				else
				validata =1
          		end if   
			end if
			case else	
				validata =1		
			end select		
	else
	validata =1
    end if 
Else
validata =1
end if 
Else
validata =1
end if 
Else
validata =1
end if 

if cep="" or isnull(cep) then
cep="0-0"
end if
ceplit = split(cep,"-")
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
			
if apelido = "" then
apelido = ""
end if

if rua = "" then
rua = "nulo"
end if

if numero = "" then
numero = 0
elseif IsNumeric(numero) Then
valinm = 0
else
valinm = 1
end if

if complemento = "" then
complemento = "nulo"
end if

if bairro = "" then
bairro = 0
end if



if validata = 1 then
if opt = "inc" then
response.redirect("altera.asp?ori=02&nvg="&nvg&"&z=2&e=dt&cod_cons="&cod_cons&"&nome="&nome&"&apelido="&apelido&"&nasce="&nasce&"&sexo="&sexo&"&pais="&pais&"&nacionalidade="&nacionalidade&"&estadonat="&estadonat&"&cidadenat="&cidadenat&"&rua="&rua&"&numero="&numero&"&complemento="&complemento&"&cep="&cep&"&estadodom="&estadodom&"&ciddom="&ciddom&"&bairro="&bairro&"&telefones="&telefones&"&email="&email&"&ativo="&ativo&"&")
elseif opt = "alt" then
response.redirect("altera.asp?ori=02&nvg="&nvg&"&z=3&e=dt&cod_cons="&cod_cons&"&nome="&nome&"&apelido="&apelido&"&nasce="&nasce&"&sexo="&sexo&"&pais="&pais&"&nacionalidade="&nacionalidade&"&estadonat="&estadonat&"&cidadenat="&cidadenat&"&rua="&rua&"&numero="&numero&"&complemento="&complemento&"&cep="&cep&"&estadodom="&estadodom&"&ciddom="&ciddom&"&bairro="&bairro&"&telefones="&telefones&"&email="&email&"&ativo="&ativo&"&")
end if
elseif valinm = 1 then
if opt = "inc" then
response.redirect("altera.asp?ori=02&nvg="&nvg&"&z=2&e=nb&cod_cons="&cod_cons&"&nome="&nome&"&apelido="&apelido&"&nasce="&nasce&"&sexo="&sexo&"&pais="&pais&"&nacionalidade="&nacionalidade&"&estadonat="&estadonat&"&cidadenat="&cidadenat&"&rua="&rua&"&numero="&numero&"&complemento="&complemento&"&cep="&cep&"&estadodom="&estadodom&"&ciddom="&ciddom&"&bairro="&bairro&"&telefones="&telefones&"&email="&email&"&ativo="&situacao_professor&"&")
elseif opt = "alt" then
response.redirect("altera.asp?ori=02&nvg="&nvg&"&z=3&e=nb&cod="&cod_cons&"&nome="&nome&"&apelido="&apelido&"&nasce="&nasce&"&sexo="&sexo&"&pais="&pais&"&nacionalidade="&nacionalidade&"&estadonat="&estadonat&"&cidadenat="&cidadenat&"&rua="&rua&"&numero="&numero&"&complemento="&complemento&"&cep="&cep&"&estadodom="&estadodom&"&ciddom="&ciddom&"&bairro="&bairro&"&telefones="&telefones&"&email="&email&"&ativo="&ativo&"&")
end if
elseif valicp = 1 then
if opt = "inc" then
response.redirect("altera.asp?ori=02&nvg="&nvg&"&z=2&e=cp&cod_cons="&cod_cons&"&nome="&nome&"&apelido="&apelido&"&nasce="&nasce&"&sexo="&sexo&"&pais="&pais&"&nacionalidade="&nacionalidade&"&estadonat="&estadonat&"&cidadenat="&cidadenat&"&rua="&rua&"&numero="&numero&"&complemento="&complemento&"&cep="&cep&"&estadodom="&estadodom&"&ciddom="&ciddom&"&bairro="&bairro&"&telefones="&telefones&"&email="&email&"&ativo="&ativo&"&")
elseif opt = "alt" then
response.redirect("altera.asp?ori=02&nvg="&nvg&"&z=3&e=cp&cod_cons="&cod_cons&"&nome="&nome&"&apelido="&apelido&"&nasce="&nasce&"&sexo="&sexo&"&pais="&pais&"&nacionalidade="&nacionalidade&"&estadonat="&estadonat&"&cidadenat="&cidadenat&"&rua="&rua&"&numero="&numero&"&complemento="&complemento&"&cep="&cep&"&estadodom="&estadodom&"&ciddom="&ciddom&"&bairro="&bairro&"&telefones="&telefones&"&email="&email&"&ativo="&ativo&"&")
end if
else

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CONwr = Server.CreateObject("ADODB.Connection") 
		ABRIRwr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONwr.Open ABRIRwr		


if opt = "inc" then


	call ultimo(0) 
	cod_cons= session("codigo_u")
	call ultimo(1) 	
	co_usr_prof = session("codigo_u2")
	
escola=session("escola")

		Set RSwr = Server.CreateObject("ADODB.Recordset")
		SQLwr = "SELECT * FROM TB_Ano_Letivo WHERE NU_Ano_Letivo ='"& ano_letivo&"'"
'response.Write(SQLwr&"-")	
		RSwr.Open SQLwr, CON1
no_escola=RSwr("TX_String_Usuario")
'response.Write(no_escola&"-")

login=no_escola&co_usr_prof
'response.Write(login)
'response.End()

if cep="0-0" then
cep=""
end if



Set RS1 = server.createobject("adodb.recordset")
RS1.open "TB_Usuario", CON1, 2, 2 'which table do you want open


'response.Write(">>>"&cod_cons)

RS1.addnew
RS1("CO_Usuario") = co_usr_prof
RS1("NO_Usuario") = nome
RS1("Login") = login
RS1("Email_Usuario") = email
RS1("Senha") = co_usr_prof
RS1("ST_Usuario") = situacao_usuario
RS1.update
  
set RS1=nothing



Set RS = server.createobject("adodb.recordset")
RS.open "TB_Professor", CON, 2, 2 'which table do you want open

RS.addnew
RS("CO_Professor") = cod_cons
RS("NO_Professor") = nome
RS("NO_Apelido_Professor") = apelido
RS("IN_Sexo") = sexo
RS("DA_Nascimento") = nasce
RS("NO_Logradouro") = rua
RS("NU_Logradouro") = numero
RS("TX_Complemento_Logradouro") = complemento
RS("CO_Bairro") = bairro
RS("CO_Municipio") = ciddom
RS("CO_Pais") = pais
RS("SG_UF") = estadodom
RS("CO_CEP") = cep
RS("NUS_Telefones") = telefones
RS("SG_Estado_Natural") = estadonat
RS("CO_Nacionalidade") = nacionalidade
RS("CO_Municipio_Natural") = cidadenat
RS("TX_EMail") = email
RS("IN_Ativo_Escola") = situacao_professor
RS("CO_Usuario") = co_usr_prof
RS.update
  
set RS=nothing

Set RS1a = server.createobject("adodb.recordset")
RS1a.open "TB_Autoriz_Usuario_Grupo", CON1, 2, 2 'which table do you want open
RS1a.addnew
RS1a("CO_Usuario") = co_usr_prof
RS1a("CO_Grupo") = "PRO"
RS1a.update
  
set RS1a=nothing


call GravaLog (nvg,"Professor de código de usuário "&co_usr_prof&" incluído")

response.Redirect("index.asp?ori=02&opt=ok&nvg="&nvg&"&cod_cons="& cod_cons &"&co_usr_prof="&co_usr_prof&"&tx_login="&login)

elseif opt = "alt" then

co_usr_prof = request.Form("co_usr_prof")

'if email = "" then
'email = co_usr_prof&"@email.com.br"
'end if

if cep="0-0" then
cep="00000-000"
end if


sql_atualiza = "UPDATE TB_Professor SET [NO_Professor] = '"&nome&"',"
sql_atualiza = sql_atualiza&"	[NO_Apelido_Professor] = '"&apelido&"',"
sql_atualiza = sql_atualiza&"	[IN_Sexo] = '"&sexo&"',"
sql_atualiza = sql_atualiza&"	[DA_Nascimento] = #"&nasce&"#," 
sql_atualiza = sql_atualiza&"	[NO_Logradouro] = '"&rua&"',"
sql_atualiza = sql_atualiza&"	[NU_Logradouro] = "&numero&","
sql_atualiza = sql_atualiza&"	[TX_Complemento_Logradouro] = '"&complemento&"',"
sql_atualiza = sql_atualiza&"	[CO_Bairro] = "&bairro&","
sql_atualiza = sql_atualiza&"	[CO_Municipio] = "&ciddom&"," 
sql_atualiza = sql_atualiza&"	[CO_Pais] = "&pais&"," 
sql_atualiza = sql_atualiza&"	[SG_UF] = '"&estadodom&"'," 
sql_atualiza = sql_atualiza&"	[CO_CEP] = '"&cep&"',"
sql_atualiza = sql_atualiza&"	[NUS_Telefones] = '"&telefones&"',"
sql_atualiza = sql_atualiza&"	[SG_Estado_Natural] = '"&estadonat&"',"
sql_atualiza = sql_atualiza&"	[CO_Nacionalidade] = "&nacionalidade&","
sql_atualiza = sql_atualiza&"	[CO_Municipio_Natural] = "&cidadenat&","
sql_atualiza = sql_atualiza&"	[TX_EMail]= '"&email&"', "
sql_atualiza = sql_atualiza&"	[IN_Ativo_Escola] ="&situacao_professor&","
sql_atualiza = sql_atualiza&"	[CO_Usuario]= "&co_usr_prof
sql_atualiza = sql_atualiza&" WHERE [CO_Professor] = "& cod_cons 
response.Write(sql_atualiza)
'response.End()
Set RSup = CON.Execute(sql_atualiza)


sql_atualiza2 = "UPDATE TB_Usuario SET NO_Usuario = '"&nome&"', Email_Usuario = '"&email&"', ST_Usuario = '"&situacao_usuario&"' WHERE CO_Usuario = "& co_usr_prof 
Set RSup2 = CON1.Execute(sql_atualiza2)
'response.Write("3"& sql_atualiza2)
'response.End()
call GravaLog (nvg,"Dados cadastrais do professor de código de usuário "&co_usr_prof&" alterado")

response.Redirect("altera.asp?ori=01&nvg="&nvg&"&opt=ok&cod_cons="& cod_cons &"")

end if

end if


%>

</html>
<%If Err.number<>0 then
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