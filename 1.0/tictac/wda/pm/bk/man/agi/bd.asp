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
tipo_peso = request.Form("tipo_peso")
minimo = request.Form("minimo")
grupo_form = request.Form("grupo")
observacoes = request.Form("observacoes")


		Set CON9 = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_ax & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON9.Open ABRIR



	if opt = "inc" then
				
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT MAX(CO_Item) as Max_CO_Item FROM TB_Item"
		RS.Open SQL, CON9
		
		if RS.EOF then
			cod_cons = 1
		else
			ult_codigo = RS("Max_CO_Item")
			cod_cons = ult_codigo+1
		end if	
		
		
		
		Set RS = server.createobject("adodb.recordset")
		RS.open "TB_Item", CON9, 2, 2 'which table do you want open
		
		RS.addnew
		RS("CO_Item") = cod_cons
		RS("NO_Item") = nome
		RS("NO_Apelido_Item") = apelido
		RS("CO_Tipo_Peso") = tipo_peso
		RS("QT_Estoque_Minimo") = minimo
		RS("CO_Grupo") = grupo_form
		RS("TX_Observacoes") = observacoes
		RS.update
		  
		set RS=nothing
		
		
		call GravaLog (nvg,"Fornecedor de código "&cod_cons&" incluído")
		
		response.Redirect("index.asp?ori=02&opt=ok&nvg="&nvg&"&cod_cons="& cod_cons )
		
	elseif opt = "alt" then		
		
		sql_atualiza = "UPDATE TB_Item SET [NO_Item] = '"&nome&"',"
		sql_atualiza = sql_atualiza&"	[NO_Apelido_Item] = '"&apelido&"',"
		sql_atualiza = sql_atualiza&"	[CO_Tipo_Peso] = '"&tipo_peso&"',"
		sql_atualiza = sql_atualiza&"	[QT_Estoque_Minimo] = "&minimo&","
		sql_atualiza = sql_atualiza&"	[CO_Grupo] = "&grupo_form&","
		sql_atualiza = sql_atualiza&"	[TX_Observacoes] = '"&observacoes&"'"
		sql_atualiza = sql_atualiza&" WHERE [CO_Item] = "& cod_cons 
		'response.Write(sql_atualiza)
		'response.End()
		Set RSup = CON9.Execute(sql_atualiza)
		
		'response.End()
		call GravaLog (nvg,"Dados cadastrais do Item de código "&cod_cons&" alterados")
		
		response.Redirect("altera.asp?ori=01&nvg="&nvg&"&opt=ok&cod_cons="& cod_cons &"")
	
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