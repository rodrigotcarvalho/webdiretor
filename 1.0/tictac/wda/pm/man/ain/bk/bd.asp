<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
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

cod_grupo = request.Form("cod_grupo")




	Set CON9 = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO_ax & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON9.Open ABRIR
		
    Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Item WHERE CO_Grupo ="& cod_grupo&" ORDER BY NO_Item"
	RS.Open SQL, CON9
	
	while not RS.EOF
		cod_cons = RS("CO_Item")		

		minimo = request.Form("minimo_"&cod_cons)
		estoque = request.Form("estoque_"&cod_cons)	
		alerta = request.Form("alerta_"&cod_cons)
	
		if isnull(minimo) or minimo = "" then
			minimo = 0
		end if	
		
		if isnull(estoque) or estoque = "" then
			estoque = 0
		end if						
		
		if isnull(alerta) or alerta = "" then
			alerta = 2
		end if
	
		sql_atualiza = "UPDATE TB_Item SET [QT_Estoque_Minimo] = "&minimo&","
		sql_atualiza = sql_atualiza&"	[QT_Atual] = "&estoque&","
		sql_atualiza = sql_atualiza&"	[QV_Estoque_Minimo] = "&alerta		
		sql_atualiza = sql_atualiza&" WHERE [CO_Item] = "& cod_cons 
		response.Write(sql_atualiza)

		Set RSup = CON9.Execute(sql_atualiza)
		
		'response.End()
		call GravaLog (nvg,"Quantidades do Item de código "&cod_cons&" alteradas")

	  RS.MOVENEXT
	  WEND		
		'response.End()	  

	

		response.Redirect("altera.asp?nvg="&nvg&"&opt=ok&grupo="& cod_grupo &"")


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