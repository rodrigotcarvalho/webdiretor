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

cod_cons = request.Form("cod_cons")
nome_projeto = request.Form("nome_projeto")
co_etapa = request.Form("etapa")



	Set CON9 = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO_ax & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON9.Open ABRIR

	cod_cons = cod_cons*1

	if cod_cons = 0 then
	
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT MAX(CO_Projeto) as Max_CO_Projeto FROM TB_Projeto"
		RS.Open SQL, CON9
		
		if RS.EOF then
			cod_cons = 1
		else
			ult_codigo = RS("Max_CO_Projeto")
			cod_cons = ult_codigo+1
		end if	
		
		
		
		Set RS = server.createobject("adodb.recordset")
		RS.open "TB_Projeto", CON9, 2, 2 'which table do you want open
		
		RS.addnew
		RS("CO_Projeto") = cod_cons
		RS("NO_Projeto") = nome_projeto
		RS("CO_Etapa") = co_etapa		
		RS.update
		  
		set RS=nothing
		
		
		call GravaLog (nvg,"Projeto de código "&cod_cons&" incluído")	
		opt ="ok1"		
	else
	
		sql_atualiza = "UPDATE TB_Projeto SET [NO_Projeto] = '"&nome_projeto&"', [CO_Etapa] = '"&co_etapa&"'"
		sql_atualiza = sql_atualiza&" WHERE [CO_Projeto] = "& cod_cons 
		Set RSup = CON9.Execute(sql_atualiza)

		'response.End()
			call GravaLog (nvg,"Projeto de código "&cod_cons&" alterado")
			opt ="ok2"
	end if  

	

		response.Redirect("altera.asp?nvg="&nvg&"&opt="&opt&"&projeto="& cod_cons&"")


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