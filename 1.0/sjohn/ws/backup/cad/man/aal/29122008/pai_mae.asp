<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->

<%
'opt= request.querystring("opt")

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0

		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CONCONT_aux = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT_aux = "DBQ="& CAMINHO_ct_aux & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT_aux.Open ABRIRCONT_aux
		
		Set CON1_aux = Server.CreateObject("ADODB.Connection") 
		ABRIR1_aux = "DBQ="& CAMINHO_al_aux & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1_aux.Open ABRIR1_aux		

tp= request.form("tp_familiares")
nome= request.form("nome_pub")
qld= request.form("qld_pub")
cod= request.form("cod_pub")

response.Write(nome&"-"&tp&"-"&qld&"-"&cod)

if tp="p" then
form_campo="pai"
tp_nome="Pai"
cod_familiar="PAI"
bd="NO_Pai"
	if session("pai_ok")="s" then
	session("pai_ok")=session("pai_ok")
	else
	session("pai_ok")="s"
	end if
session("mae_ok")=session("mae_ok")
elseif tp="m" then
form_campo="mae"
tp_nome="Mãe"
cod_familiar="MAE"
bd="NO_Mae"
	if session("mae_ok")="s" then
	session("mae_ok")=session("mae_ok")
	else
	session("mae_ok")="s"
	end if
session("pai_ok")=session("pai_ok")
end if

if qld="n" then
	Set RSCONTATO_aux_bd = server.createobject("adodb.recordset")
	RSCONTATO_aux_bd.open "TBI_Contatos", CONCONT_aux, 2, 2 'which table do you want open
	RSCONTATO_aux_bd.addnew
	RSCONTATO_aux_bd("CO_Matricula")=cod
	RSCONTATO_aux_bd("TP_Contato")=cod_familiar
	RSCONTATO_aux_bd("NO_Contato")=nome
	RSCONTATO_aux_bd.update
	set RSCONTATO_aux_bd=nothing
	
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TBI_Alunos WHERE CO_Matricula ="& cod
		RS.Open SQL, CON1_aux
		
		if RS.EOF then
			
			Set RSALUNO_aux_bd = server.createobject("adodb.recordset")
			RSALUNO_aux_bd.open "TBI_Alunos", CON1_aux, 2, 2
			RSALUNO_aux_bd.addnew
			RSALUNO_aux_bd("CO_Matricula")=cod
			RSALUNO_aux_bd(""&bd&"")=nome							  
			RSALUNO_aux_bd.update
			set RSALUNO_aux_bd=nothing
		
		else	
			Set RSALUNO_aux_bd2 = server.createobject("adodb.recordset")
			sql_atualiza_al= "UPDATE TBI_Alunos SET "&bd&" ='"& nome &"' WHERE CO_Matricula = "& cod
			Set RSALUNO_aux_bd2 = CON1_aux.Execute(sql_atualiza_al)
		end if	
	
else
	Set RSCONTATO_aux_bd2 = server.createobject("adodb.recordset")
	sql_atualiza= "UPDATE TBI_Contatos SET NO_Contato = '"&nome&"' WHERE CO_Matricula = "& cod &" AND TP_Contato = '"& cod_familiar &"'"
	Set RSCONTATO_aux_bd2 = CONCONT_aux.Execute(sql_atualiza)
		
	Set RSALUNO_aux_bd2 = server.createobject("adodb.recordset")
	sql_atualiza_al= "UPDATE TBI_Alunos SET "&bd&" ='"& nome &"' WHERE CO_Matricula = "& cod
	Set RSALUNO_aux_bd2 = CON1_aux.Execute(sql_atualiza_al)

end if
%>
<input name="<%response.Write(form_campo)%>" type="text" class="borda" onBlur="recuperarMae(this.value,'<%response.Write(tp)%>','s','<%response.Write(cod)%>')" value="<%response.Write(mae)%>" onKeyDown="return KeyTest()" size="30" maxlength="50">
