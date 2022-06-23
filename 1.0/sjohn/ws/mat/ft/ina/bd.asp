<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<%
chave=session("nvg")
session("nvg")=chave
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo")
ano_letivo_real = ano_letivo
sistema_local=session("sistema_local")
opt=request.querystring("opt")

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1		
		
			ano = DatePart("yyyy", now)
			mes = DatePart("m", now) 
			dia = DatePart("d", now) 
			da_encerramento =dia&"/"& mes &"/"& ano

cod=request.form("cod")
situacao=request.form("situacao")			
motivo=request.form("motivo")		

Set RSALUNO_aux_bd2 = server.createobject("adodb.recordset")
sql_atualiza_al= "UPDATE TB_Matriculas SET CO_Situacao ='"& situacao &"', DA_Encerramento=#"& da_encerramento &"#, DS_Motivo ='"& motivo &"' WHERE CO_Matricula = "& cod&" AND NU_Ano="&ano_letivo
Set RSALUNO_aux_bd2 = CON1.Execute(sql_atualiza_al)

outro=cod&"/"&situacao

			call GravaLog (chave,outro)

response.Write(sql_atualiza_al)
response.redirect("altera.asp?opt=ok&cod_cons="&cod)
%>