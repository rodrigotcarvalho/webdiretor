<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<%
chave=session("nvg")
session("nvg")=chave
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo")
ano_letivo_real = ano_letivo
sistema_local=session("sistema_local")
opt=request.querystring("opt")

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
		
if opt="af" then				
variavel=request.form("variavel_pub")
cod=request.form("cod_pub")
cod_familiar=request.form("foco_pub")
tp_vinc=request.form("tp_vinc_pub")
cod_vinc=request.form("cod_vinc_pub")
bd=request.form("bd_pub")

variavel=unescape(variavel)

if (isnull(cod_vinc) or cod_vinc="") and (isnull(tp_vinc) or tp_vinc="") then
tp_vinc="NULL"
cod_vinc="NULL"
SQLAA_aux_tx="SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod
else
SQLAA_aux_tx="SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_vinc&"' and CO_Matricula ="&cod_vinc
end if

'response.Write(SQLAA_aux_tx)

		Set RSCONTATO_aux = Server.CreateObject("ADODB.Recordset")
		SQLAA_aux= SQLAA_aux_tx
		RSCONTATO_aux.Open SQLAA_aux, CONCONT_aux

if RSCONTATO_aux.EOF then

if (isnull(cod_vinc) or cod_vinc="" or cod_vinc="NULL") and (isnull(tp_vinc) or tp_vinc="" or tp_vinc="NULL") then
tp_vinc=null
cod_vinc=null
end if

Set RSCONTATO_aux_bd = server.createobject("adodb.recordset")

RSCONTATO_aux_bd.open "TBI_Contatos", CONCONT_aux, 2, 2 'which table do you want open
RSCONTATO_aux_bd.addnew
RSCONTATO_aux_bd("CO_Matricula")=cod
RSCONTATO_aux_bd("TP_Contato")=cod_familiar
RSCONTATO_aux_bd(""&bd&"")=variavel
RSCONTATO_aux_bd("CO_Matricula_Vinc")=cod_vinc
RSCONTATO_aux_bd("TP_Contato_Vinc")=tp_vinc

  RSCONTATO_aux_bd.update
  
set RSCONTATO_aux_bd=nothing

else	


	if isnumeric(variavel) then
		if isnull(variavel) or variavel="" then
			sql=bd&"=NULL"
		else	
			sql=bd&"="&variavel
		end if
	else
		if bd="DA_Nascimento_Contato" or bd="CO_DERG_PFisica" then
			if isnull(variavel) or variavel="" then
				sql=bd&"=NULL"
			else
				vetor_nascimento = Split(variavel,"/")  
				dia_n = vetor_nascimento(0)
				mes_n = vetor_nascimento(1)
				ano_n = vetor_nascimento(2)	

				dia_a = dia_n
				mes_a = mes_n
				ano_a = ano_n
		
				variavel = mes_n&"/"&dia_n&"/"&ano_n
				sql=bd&"=#"&variavel&"#"
			end if
		elseif bd="ID_Res_Aluno" then
			if variavel="s" then
				sql=bd&"=TRUE"
			else
				sql=bd&"=FALSE"
			end if
		elseif isnull(variavel) or variavel="" then
			sql=bd&"=NULL"
		else	
			sql=bd&"='"&variavel&"'"
		end if
	end if
	
if (isnull(cod_vinc) or cod_vinc="NULL" or cod_vinc="") and (isnull(tp_vinc) or tp_vinc="NULL" or tp_vinc="") then
sql_atualiza_tx="UPDATE TBI_Contatos SET "&sql&" WHERE CO_Matricula = "& cod &" AND TP_Contato = '"& cod_familiar &"'"
sql_atualiza_tx1="1UPDATE TBI_Contatos SET "&sql&" WHERE CO_Matricula = "& cod &" AND TP_Contato = '"& cod_familiar &"'"
else
sql_atualiza_tx="UPDATE TBI_Contatos SET "&sql&" WHERE CO_Matricula = "& cod_vinc &" AND TP_Contato = '"& tp_vinc &"'"
sql_atualiza_tx1="2UPDATE TBI_Contatos SET "&sql&" WHERE CO_Matricula = "& cod_vinc &" AND TP_Contato = '"& tp_vinc &"'"
end if		

Set RSCONTATO_aux_bd2 = server.createobject("adodb.recordset")
sql_atualiza= sql_atualiza_tx
Set RSCONTATO_aux_bd2 = CONCONT_aux.Execute(sql_atualiza)

response.Write(Server.URLEncode(sql_atualiza_tx1))

'if do RSCONTATO_aux.EOF
END IF


'else do opt====================================================================================================================
elseif opt="ef" then
ordem_familiares=request.Form("ord_pub")
qtd_tipo_familiares=request.Form("qtd_tp_pub")
cod=request.Form("cod_pub")
foco=request.Form("foco_pub")

		Set RSCONTATO = Server.CreateObject("ADODB.Recordset")
		SQLAA= "DELETE * FROM TBI_Contatos WHERE TP_Contato='"&foco&"' and CO_Matricula ="&cod
		RSCONTATO.Open SQLAA, CONCONT_aux


'else do opt====================================================================================================================
elseif opt="re" then
variavel=request.form("variavel_pub")
bd=request.form("bd_pub")
valor_resp=request.form("valor_resp_pub")
tipo_resp=request.form("tipo_resp_pub")
cod=request.form("cod_pub")

if bd="TP_Resp_Fin" or bd="TP_Resp_Ped" then

		Set RS_aux = Server.CreateObject("ADODB.Recordset")
		SQL_aux = "SELECT * FROM TBI_Alunos WHERE CO_Matricula ="& cod
		RS_aux.Open SQL_aux, CON1_aux
		
if RS_aux.EOF Then
Set RSALUNO_aux_bd = server.createobject("adodb.recordset")
RSALUNO_aux_bd.open "TBI_Alunos", CON1_aux, 2, 2
RSALUNO_aux_bd.addnew
RSALUNO_aux_bd("CO_Matricula")=cod
RSALUNO_aux_bd(""&bd&"")=variavel							  
RSALUNO_aux_bd.update
  
set RSALUNO_aux_bd=nothing

else

Set RSALUNO_aux_bd2 = server.createobject("adodb.recordset")
sql_atualiza_al= "UPDATE TBI_Alunos SET "&bd&" ='"& variavel &"' WHERE CO_Matricula = "& cod

Set RSALUNO_aux_bd2 = CON1_aux.Execute(sql_atualiza_al)
end if

elseif bd="ID_Familia" then
Set RSCONTATO_aux_bd3 = server.createobject("adodb.recordset")
sql_atualiza3= "UPDATE TBI_Contatos SET ID_Familia = '"&variavel&"' WHERE CO_Matricula = "& cod &" AND TP_Contato = '"& valor_resp &"'"
Set RSCONTATO_aux_bd3 = CONCONT_aux.Execute(sql_atualiza3)

elseif bd="ID_End_Bloqueto" and tipo_resp="TP_Resp_Fin" then
Set RSCONTATO_aux_bd4 = server.createobject("adodb.recordset")
sql_atualiza4= "UPDATE TBI_Contatos SET ID_End_Bloqueto ='"& variavel &"' WHERE CO_Matricula = "& cod &" AND TP_Contato = '"& valor_resp &"'"
Set RSCONTATO_aux_bd4 = CONCONT_aux.Execute(sql_atualiza4)

elseif bd="ID_End_Bloqueto" and tipo_resp="TP_Resp_Ped" then
Set RSCONTATO_aux_bd5 = server.createobject("adodb.recordset")
sql_atualiza5= "UPDATE TBI_Contatos SET ID_End_Bloqueto ='"& variavel &"' WHERE CO_Matricula = "& cod &" AND TP_Contato = '"& valor_resp &"'"
Set RSCONTATO_aux_bd5 = CONCONT_aux.Execute(sql_atualiza5)
end if		
END IF		
%>