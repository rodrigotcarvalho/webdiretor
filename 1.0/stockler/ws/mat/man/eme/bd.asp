<%'On Error Resume Next%>
<!--#include file="../../../../inc/connect_a.asp"-->
<!--#include file="../../../../inc/connect_al.asp"-->
<!--#include file="../../../../inc/connect_pr.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<%
chave=session("nvg")
session("nvg")=chave


		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1

		Set CON_al = Server.CreateObject("ADODB.Connection") 
		ABRIR_al = "DBQ="& CAMINHOa& ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_al.Open ABRIR_al


ano_letivo = session("ano_letivo")
co_usr = session("co_user")

cod=request.form("cod")
resp_ped=request.form("resp_ped")
resp_fin=request.form("resp_fin")
ano_letivo_mat=request.form("ano_letivo_mat")
sit_nova_mat=request.form("sit_nova_mat")
unidade=request.form("unidade")
curso=request.form("curso")
etapa=request.form("etapa")
turma=request.form("turma")
ult_ano_aluno=request.form("ult_ano_aluno")
pre_matricula=request.form("pre_matricula")

aluno_novo=request.form("aluno_novo")


		Set RSdt = Server.CreateObject("ADODB.Recordset")
		SQLdt = "SELECT * FROM TB_Documentos_Matricula order by NO_Documento"
		RSdt.Open SQLdt, CON0

If Not IsArray(co_doc_mat) Then co_doc_mat = Array() End if
'If Not IsArray(doc) Then doc = Array() End if
i=0
while not RSdt.EOF
If InStr(Join(co_doc_mat), RSdt("CO_Documento")) = 0 Then
ReDim preserve co_doc_mat(UBound(co_doc_mat)+1)
teste=request.form(RSdt("CO_Documento"))
doc_inclui=RSdt("CO_Documento")&"|"&teste

co_doc_mat(Ubound(co_doc_mat)) = doc_inclui

end if
i=i+1
RSdt.Movenext
wend


ano = DatePart("yyyy", now)
mes_de = DatePart("m", now) 
dia_de = DatePart("d", now) 

if dia_de<10 then
dia_de="0"&dia_de
end if

if mes_de<10 then
mes_de="0"&mes_de
end if
					
data_cadastro=mes_de&"/"&dia_de&"/"&ano

if aluno_novo="s" then
Set RS = server.createobject("adodb.recordset")
if pre_matricula="aberta" then
RS.open "TB_Matriculas", CON1, 2, 2 'which table do you want open
RS.addnew
  RS("CO_Matricula") = cod
  RS("CO_Situacao") = sit_nova_mat
  RS("NU_Ano") = ano_letivo
  RS("NU_Unidade") = unidade
  RS("CO_Curso") = curso
  RS("CO_Etapa") = etapa
  RS("CO_Turma") = co_turma
  RS("DA_Rematricula") = data_cadastro
RS.update
end if
sql_resp="UPDATE TB_Alunos SET TP_Resp_Fin='"&resp_fin&"', TP_Resp_Ped='"&resp_ped&"' where CO_Matricula="&cod
Set RS_resp = CON1.Execute(sql_resp)
  
set RS=nothing

For j=0 to Ubound(co_doc_mat)
docs=split(co_doc_mat(j),"|")
tipo_doc=docs(0)
doc_entregue=docs(1)
if doc_entregue="S" then
Set RS_doc = server.createobject("adodb.recordset")
RS_doc.open "TB_Documentos_Entregues", CON0, 2, 2 'which table do you want open
RS_doc.addnew
RS_doc("CO_Matricula") = cod
RS_doc("CO_Documento") = tipo_doc
RS_doc("DA_Entrega_Documento") = data_cadastro
RS_doc.update
set RS_doc=nothing
end if
next


outro= cod
call GravaLog (chave,outro)
response.redirect("altera.asp?or=01&cod="&cod&"&opt=ok")

else
if pre_matricula="aberta" then
sql_atualiza= "UPDATE TB_Matriculas SET CO_Situacao='"&sit_nova_mat&"', NU_Unidade="&unidade&", CO_Curso='"&curso&"', CO_Etapa='" & etapa & "', CO_Turma='" & turma & "', DA_Rematricula =#"& data_cadastro &"#, Ult_NU_Ano="&ult_ano_aluno&" WHERE CO_Matricula = "&cod&" AND NU_Ano="&ano_letivo

Set RS2 = CON1.Execute(sql_atualiza)
end if

sql_resp="UPDATE TB_Alunos SET TP_Resp_Fin='"&resp_fin&"', TP_Resp_Ped='"&resp_ped&"' where CO_Matricula="&cod
Set RS_resp = CON1.Execute(sql_resp)

For j=0 to Ubound(co_doc_mat)
docs=split(co_doc_mat(j),"|")
tipo_doc=docs(0)
doc_entregue=docs(1)
if doc_entregue="S" then

		Set RSde = Server.CreateObject("ADODB.Recordset")
		SQLde = "SELECT * FROM TB_Documentos_Entregues where CO_Documento='"&tipo_doc&"' And CO_Matricula="&cod
		RSde.Open SQLde, CON0

IF RSde.EOF then
		Set RS_doc = Server.CreateObject("ADODB.Recordset")
RS_doc.open "TB_Documentos_Entregues", CON0, 2, 2 'which table do you want open
RS_doc.addnew
RS_doc("CO_Matricula") = cod
RS_doc("CO_Documento") = tipo_doc
RS_doc("DA_Entrega_Documento") = data_cadastro
RS_doc.update
set RS_doc=nothing
end if

else
		Set RSde = Server.CreateObject("ADODB.Recordset")
		SQLde = "SELECT * FROM TB_Documentos_Entregues where CO_Documento='"&tipo_doc&"' And CO_Matricula="&cod
		RSde.Open SQLde, CON0

IF RSde.EOF then
else
		Set RS_doc = Server.CreateObject("ADODB.Recordset")
SQL_doc = "DELETE * FROM TB_Documentos_Entregues where CO_Documento='"&tipo_doc&"' And CO_Matricula="&cod
		RS_doc.Open SQL_doc, CON0
end if
end if
next


outro= cod
call GravaLog (chave,outro)

response.redirect("altera.asp?or=01&cod="&cod&"&opt=ok1")
end if



%>