<%'On Error Resume Next%>
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
textfield=request.form("textfield")
textfield2=request.form("textfield2")
textfield3=request.form("textfield3")
textfield4=request.form("textfield4")


Set CONt = Server.CreateObject("ADODB.Connection") 
ABRIRt = "DBQ="& CAMINHO_t & ";Driver={Microsoft Access Driver (*.mdb)}"
CONt.Open ABRIRt	

		Set RSdt = Server.CreateObject("ADODB.Recordset")
		SQLdt = "SELECT * FROM TB_OBS_Ficha where CO_Matricula="&cod_cons
		RSdt.Open SQLdt, CONt


		IF RSdt.EOF then
			Set RS_doc = Server.CreateObject("ADODB.Recordset")
			RS_doc.open "TB_OBS_Ficha", CONt, 2, 2 'which table do you want open
			RS_doc.addnew
			RS_doc("CO_Matricula") = cod_cons
			RS_doc("TX_OBS_LIN1") = textfield
			RS_doc("TX_OBS_LIN2") = textfield2
			RS_doc("TX_OBS_LIN3") = textfield3
			RS_doc("TX_OBS_LIN4") = textfield4						
			RS_doc.update
			set RS_doc=nothing
		else
			Set RS_doc2 = server.createobject("adodb.recordset")
			sql_doc= "UPDATE TB_OBS_Ficha SET TX_OBS_LIN1='"& textfield &"', TX_OBS_LIN2='"& textfield2 &"', TX_OBS_LIN3='"& textfield3 &"', TX_OBS_LIN4='"& textfield4 &"' where CO_Matricula="&cod_cons			
			Set RS_doc2 = CONt.Execute(sql_doc)

		end if



outro= ""
call GravaLog (chave,outro)
if ori="3" then
response.redirect("altera.asp?ori=3&cod_cons="&cod_cons&"&pagina="&pagina&"&opt=ok1")
else
response.redirect("altera.asp?ori=1&cod_cons="&cod_cons&"&opt=ok1")
end if
%>