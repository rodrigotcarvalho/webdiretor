



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

		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHOa &";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2	
		
		
unidade_veio=request.form("unidade_veio")
curso_veio=request.form("curso_veio")
etapa_veio=request.form("etapa_veio")
turma_veio=request.form("turma_veio")
cham_veio=request.form("cham_veio")

cod=request.form("cod")
unidade=request.form("unidade")
curso=request.form("curso")			
etapa=request.form("etapa")
turma=request.form("turma")
chamada=request.form("chamada")		
dia_remaneja=request.form("dia_remaneja")
mes_remaneja=request.form("mes_remaneja")
ano_remaneja=request.form("ano_remaneja")
motivo=request.form("motivo")

data_remaneja = dia_remaneja &"/"& mes_remaneja &"/"& ano_remaneja

Set RSALUNO_aux_bd2 = server.createobject("adodb.recordset")
sql_atualiza_al= "UPDATE TB_Aluno_Esta_Turma SET NU_Unidade_Rm ="& unidade &", CO_Curso_Rm='"& curso &"', CO_Etapa_Rm ='"& etapa &"', CO_Turma_Rm ='"& turma &"', DA_Remanejamento =#"& data_remaneja &"#,DS_Motivo='"&motivo&"' WHERE CO_Matricula = "& cod &" AND NU_Unidade ="& unidade_veio &" AND CO_Curso='"& curso_veio &"' AND CO_Etapa ='"& etapa_veio &"' AND CO_Turma ='"& turma_veio &"'"
Set RSALUNO_aux_bd2 = CON2.Execute(sql_atualiza_al)

Set RS = server.createobject("adodb.recordset")
RS.open "TB_Aluno_Esta_Turma", CON2, 2, 2 'which table do you want open
RS.addnew
  RS("CO_Matricula") = cod
  RS("NU_Unidade") = unidade
  RS("CO_Curso") = curso
  RS("CO_Etapa") = etapa
  RS("CO_Turma") = turma
  RS("NU_Chamada") = chamada
  RS.update
 
set RS=nothing


Set RSALUNO_aux_bd1 = server.createobject("adodb.recordset")
sql_atualiza_aluno= "UPDATE TB_Matriculas SET NU_Unidade ="& unidade &", CO_Curso='"& curso &"', CO_Etapa ='"& etapa &"', CO_Turma ='"& turma &"', NU_Chamada ="& chamada &" WHERE CO_Matricula = "& cod &" AND NU_Ano="&ano_letivo
Set RSALUNO_aux_bd1 = CON1.Execute(sql_atualiza_aluno)


outro=cod&"-U:"&unidade_veio&",C:"&curso_veio&",E:"&etapa_veio&",T:"&turma_veio&",CH:"&cham_veio&",U:"&unidade&",C:"&curso&",E:"&etapa&",T:"&turma&",CH:"&chamada

			call GravaLog (chave,outro)


response.Write(sql_atualiza_al)
response.redirect("altera.asp?opt=ok&cod_cons="&cod)
%>