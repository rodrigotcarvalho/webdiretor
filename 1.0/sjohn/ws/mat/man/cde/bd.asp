<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<%
chave=session("nvg")
session("nvg")=chave
ori= request.QueryString("ori")
pagina=Request.QueryString("pagina")

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

cod_cons=request.form("cod_cons")
unidade=request.form("unidade")
curso=request.form("curso")
etapa=request.form("etapa")
turma=request.form("turma")


		Set RSdt = Server.CreateObject("ADODB.Recordset")
		SQLdt = "SELECT * FROM TB_Documentos_Matricula order by NO_Documento"
		RSdt.Open SQLdt, CON0

If Not IsArray(co_doc_mat) Then co_doc_mat = Array() End if

i=0
while not RSdt.EOF
	If InStr(Join(co_doc_mat), RSdt("CO_Documento")) = 0 Then
	ReDim preserve co_doc_mat(UBound(co_doc_mat)+1)
	
	'busca a situação do doc no form	
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

For j=0 to Ubound(co_doc_mat)
docs=split(co_doc_mat(j),"|")
tipo_doc=docs(0)
doc_entregue=docs(1)
if doc_entregue="S" then

		Set RSde = Server.CreateObject("ADODB.Recordset")
		SQLde = "SELECT * FROM TB_Documentos_Entregues where CO_Documento='"&tipo_doc&"' And CO_Matricula="&cod_cons
		RSde.Open SQLde, CON0

		IF RSde.EOF then
			Set RS_doc = Server.CreateObject("ADODB.Recordset")
			RS_doc.open "TB_Documentos_Entregues", CON0, 2, 2 'which table do you want open
			RS_doc.addnew
			RS_doc("CO_Matricula") = cod_cons
			RS_doc("CO_Documento") = tipo_doc
			RS_doc("DA_Entrega_Documento") = data_cadastro
			RS_doc.update
			set RS_doc=nothing
		else
'			Set RS_doc2 = server.createobject("adodb.recordset")
'			sql_doc= "UPDATE TB_Documentos_Entregues SET CO_Documento='"& tipo_doc &"', DA_Entrega_Documento =#"& data_cadastro &"# WHERE CO_Matricula = "& cod
'			
'			response.Write(sql_doc)
'			Set RS_doc2 = CON0.Execute(sql_doc)

		end if

else
		Set RSde = Server.CreateObject("ADODB.Recordset")
		SQLde = "SELECT * FROM TB_Documentos_Entregues where CO_Documento='"&tipo_doc&"' And CO_Matricula="&cod_cons
		RSde.Open SQLde, CON0

		IF RSde.EOF then
		else
				Set RS_doc = Server.CreateObject("ADODB.Recordset")
				SQL_doc = "DELETE * FROM TB_Documentos_Entregues where CO_Documento='"&tipo_doc&"' And CO_Matricula="&cod_cons
				RS_doc.Open SQL_doc, CON0
		end if
end if
next

outro= ""
call GravaLog (chave,outro)
if ori="3" then
response.redirect("altera.asp?ori=3&cod_cons="&cod_cons&"&pagina="&pagina&"&opt=ok1")
else
response.redirect("altera.asp?ori=1&cod_cons="&cod_cons&"&opt=ok1")
end if
%>