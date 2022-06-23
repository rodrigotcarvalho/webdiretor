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
cod=request.form("cod_consulta")
nome_familiar_aux=request.form("nome_familiar")
nasce_fam_aux=request.form("nasce_fam")
cod_familiar_aux=request.form("cod_familiar")
ocupacao_fam_aux=request.form("ocupacao_fam")
empresa_fam_aux=request.form("trabalho_fam")
email_fam_aux=request.form("email_fam")
cpf_fam_aux=request.form("cpf_fam")
rg_fam_aux=request.form("id_fam")
emitido_fam_aux=request.form("tipo_id_fam")
emissao_fam_aux=request.form("nasce2_fam")
tel_fam_aux=request.form("tel_fam")
rua_res_fam_aux=request.form("rua_res_fam")
num_res_fam_aux=request.form("num_res_fam")
comp_res_fam_aux=request.form("comp_res_fam")
uf_res_fam_aux=request.form("estadores_fam")
cid_res_fam_aux=request.form("cidres_fam")
bairro_res_fam_aux=request.form("bairrores_fam")
cep_res_fam_aux=request.form("cep_fam")
tel_res_fam_aux=request.form("tel_res_fam")
id_res_fam_aux=request.form("mes_end")
rua_com_fam_aux=request.form("rua_com_fam")
num_com_fam_aux=request.form("num_com_fam")
comp_com_fam_aux=request.form("comp_com_fam")
uf_com_fam_aux=request.form("estadocom_fam")
cid_com_fam_aux=request.form("cidcom_fam")
bairro_com_aux_fam=request.form("bairrocom_fam")
cep_com_fam_aux=request.form("cepcom_fam")
tel_com_fam_aux=request.form("tel_com_fam")
tp_vinc_familiar_aux =request.form("tp_vinc_familiar_aux")
co_vinc_familiar_aux =request.form("co_vinc_familiar_aux")

'response.Write(">>"&id_res_fam_aux)

if isnull(nasce_fam_aux) or nasce_fam_aux="" then
else
vetor_nascimento = Split(nasce_fam_aux,"/")  
dia_n = vetor_nascimento(0)
mes_n = vetor_nascimento(1)
ano_n = vetor_nascimento(2)

dia_a = dia_n
mes_a = mes_n
ano_a = ano_n

nasce_fam_aux = mes_n&"/"&dia_n&"/"&ano_n
end if

if isnull(emissao_fam_aux) or emissao_fam_aux="" then
else
vetor_nascimento = Split(emissao_fam_aux,"/")  
dia_n = vetor_nascimento(0)
mes_n = vetor_nascimento(1)
ano_n = vetor_nascimento(2)

dia_a = dia_n
mes_a = mes_n
ano_a = ano_n

emissao_fam_aux = mes_n&"/"&dia_n&"/"&ano_n
end if

if id_res_fam_aux="s" then
id_res_fam_aux=TRUE
else
id_res_fam_aux=FALSE
end if

		Set RSCONTATO_aux = Server.CreateObject("ADODB.Recordset")
		SQLAA_aux= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar_aux&"' and CO_Matricula ="&cod
		RSCONTATO_aux.Open SQLAA_aux, CONCONT_aux


		response.Write("SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar_aux&"' and CO_Matricula ="&cod)

if RSCONTATO_aux.EOF then

response.Write("TADA")
if isnull(nasce_fam_aux) or nasce_fam_aux="" then
nasce_fam_aux =NULL
end if

if isnull(emissao_fam_aux) or emissao_fam_aux="" then
emissao_fam_aux=NULL
end if

if isnull(ocupacao_fam_aux) or ocupacao_fam_aux="" then
ocupacao_fam_aux =NULL
end if

if isnull(num_res_fam_aux) or num_res_fam_aux="" then
num_res_fam_aux =NULL
end if

if isnull(bairro_res_fam_aux) or bairro_res_fam_aux="" then
bairro_res_fam_aux =NULL
end if

if isnull(cid_res_fam_aux) or cid_res_fam_aux="" then
cid_res_fam_aux =NULL
end if

if isnull(num_com_fam_aux) or num_com_fam_aux="" then
num_com_fam_aux =NULL
end if

if isnull(cid_com_fam_aux) or cid_com_fam_aux="" then
cid_com_fam_aux=NULL
end if

if isnull(bairro_com_fam_aux) or bairro_com_fam_aux="" then
bairro_com_fam_aux =NULL
end if

if isnull(co_vinc_familiar_aux) or co_vinc_familiar_aux="" then
co_vinc_familiar_aux =NULL
end if

if isnull(tp_vinc_familiar_aux) or tp_vinc_familiar_aux="" then
tp_vinc_familiar_aux =NULL
end if

Set RSCONTATO_aux_bd = server.createobject("adodb.recordset")

RSCONTATO_aux_bd.open "TBI_Contatos", CONCONT_aux, 2, 2 'which table do you want open
RSCONTATO_aux_bd.addnew
RSCONTATO_aux_bd("CO_Matricula")=cod
RSCONTATO_aux_bd("TP_Contato")=cod_familiar_aux
RSCONTATO_aux_bd("NO_Contato")=nome_familiar_aux
RSCONTATO_aux_bd("DA_Nascimento_Contato")=nasce_fam_aux
RSCONTATO_aux_bd("CO_CPF_PFisica")=cpf_fam_aux
RSCONTATO_aux_bd("CO_RG_PFisica")=rg_fam_aux
RSCONTATO_aux_bd("CO_OERG_PFisica")=emitido_fam_aux
RSCONTATO_aux_bd("CO_DERG_PFisica")=emissao_fam_aux
RSCONTATO_aux_bd("TX_EMail")=email_fam_aux
RSCONTATO_aux_bd("CO_Ocupacao")=ocupacao_fam_aux
RSCONTATO_aux_bd("NO_Empresa")=empresa_fam_aux
RSCONTATO_aux_bd("NU_Telefones")=tel_fam_aux
RSCONTATO_aux_bd("ID_Res_Aluno")=id_res_fam_aux
if id_res_fam_aux=FALSE then
RSCONTATO_aux_bd("NO_Logradouro_Res")=rua_res_fam_aux
RSCONTATO_aux_bd("NU_Logradouro_Res")=num_res_fam_aux
RSCONTATO_aux_bd("TX_Complemento_Logradouro_Res")=comp_res_fam_aux
RSCONTATO_aux_bd("CO_Bairro_Res")=bairro_res_fam_aux
RSCONTATO_aux_bd("CO_Municipio_Res")=cid_res_fam_aux
RSCONTATO_aux_bd("SG_UF_Res")=uf_res_fam_aux
RSCONTATO_aux_bd("CO_CEP_Res")=cep_res_fam_aux
RSCONTATO_aux_bd("NU_Telefones_Res")=tel_res_fam_aux
end if
RSCONTATO_aux_bd("NO_Logradouro_Com")=rua_com_fam_aux
RSCONTATO_aux_bd("NU_Logradouro_Com")=num_com_fam_aux
RSCONTATO_aux_bd("TX_Complemento_Logradouro_Com")=comp_com_fam_aux
RSCONTATO_aux_bd("CO_Bairro_Com")=bairro_com_fam_aux
RSCONTATO_aux_bd("CO_Municipio_Com")=cid_com_fam_aux
RSCONTATO_aux_bd("SG_UF_Com")=uf_com_fam_aux
RSCONTATO_aux_bd("CO_CEP_Com")=cep_com_fam_aux
RSCONTATO_aux_bd("NU_Telefones_Com")=tel_com_fam_aux
RSCONTATO_aux_bd("CO_Matricula_Vinc")=co_vinc_familiar_aux
RSCONTATO_aux_bd("TP_Contato_Vinc")=tp_vinc_familiar_aux

  RSCONTATO_aux_bd.update
  
set RSCONTATO_aux_bd=nothing


else	
'response.Write("TADA-OK")
'response.Write("Nascimento="&nasce_familiar_aux)	
if isnull(nasce_fam_aux) or nasce_fam_aux="" then
sql_nasce="DA_Nascimento_Contato =NULL"
else
sql_nasce="DA_Nascimento_Contato =#"& nasce_fam_aux &"#"
end if

if isnull(emissao_fam_aux) or emissao_fam_aux="" then
sql_emissao="CO_DERG_PFisica =NULL"
else
sql_emissao="CO_DERG_PFisica =#"& emissao_fam_aux &"#"
end if

if isnull(ocupacao_fam_aux) or ocupacao_fam_aux="" then
sql_ocupacao="CO_Ocupacao =NULL"
else
sql_ocupacao="CO_Ocupacao ="& ocupacao_fam_aux &""
end if

if isnull(num_res_fam_aux) or num_res_fam_aux="" then
sql_num_res="NU_Logradouro_Res =NULL"
else
sql_num_res="NU_Logradouro_Res ="& num_res_fam_aux &""
end if

if isnull(bairro_res_fam_aux) or bairro_res_fam_aux="" then
sql_bairro_res=" CO_Bairro_Res =NULL"
else
sql_bairro_res=" CO_Bairro_Res ="& bairro_res_fam_aux &""
end if

if isnull(cid_res_fam_aux) or cid_res_fam_aux="" then
sql_cid_res=" CO_Municipio_Res =NULL"
else
sql_cid_res=" CO_Municipio_Res ="& cid_res_fam_aux &""
end if

if isnull(num_com_fam_aux) or num_com_fam_aux="" then
sql_num_com="NU_Logradouro_Com =NULL"
else
sql_num_com="NU_Logradouro_Com ="& num_com_fam_aux &""
end if

if isnull(cid_com_fam_aux) or cid_com_fam_aux="" then
sql_cid_com=" CO_Municipio_Com =NULL"
else
sql_cid_com=" CO_Municipio_Com ="& cid_com_fam_aux &""
end if

if isnull(bairro_com_fam_aux) or bairro_com_fam_aux="" then
sql_bairro_com=" CO_Bairro_Com =NULL"
else
sql_bairro_com=" CO_Bairro_Com ="& bairro_com_fam_aux &""
end if

if isnull(co_vinc_familiar_aux) or co_vinc_familiar_aux="" then
sql_vinc="CO_Matricula_Vinc =NULL"
else
sql_vinc="CO_Matricula_Vinc ="& co_vinc_familiar_aux &""
end if

Set RSCONTATO_aux_bd2 = server.createobject("adodb.recordset")
if (isnull(co_vinc_familiar_aux) or co_vinc_familiar_aux="NULL" or co_vinc_familiar_aux="") and (isnull(tp_vinc_familiar_aux) or tp_vinc_familiar_aux="NULL" or tp_vinc_familiar_aux="") then
sql_atualiza= "UPDATE TBI_Contatos SET NO_Contato = '"&nome_familiar_aux&"', "& sql_nasce &", CO_CPF_PFisica ='"& cpf_fam_aux &"', CO_RG_PFisica ='"& rg_fam_aux &"', CO_OERG_PFisica ='"& emitido_fam_aux&"', "& sql_emissao &", TX_EMail ='"& email_fam_aux &"', "&sql_ocupacao&", NO_Empresa ='"& empresa_fam_aux &"', NU_Telefones ='"& tel_fam_aux&"', ID_Res_Aluno = "&id_res_fam_aux&", "
if id_res_fam_aux=FALSE then
sql_atualiza=sql_atualiza&"NO_Logradouro_Res ='"& rua_res_fam_aux &"', "& sql_num_res&", TX_Complemento_Logradouro_Res = '"&comp_res_fam_aux&"', "& sql_bairro_res &", "& sql_cid_res &", SG_UF_Res ='"& uf_res_fam_aux &"', CO_CEP_Res ='"& cep_res_fam_aux &"', NU_Telefones_Res ='"& tel_res_fam_aux&"', "
END IF
sql_atualiza=sql_atualiza&"NO_Logradouro_Com = '"&rua_com_fam_aux&"', "& sql_num_com&", TX_Complemento_Logradouro_Com= '"&comp_com_fam_aux&"',"&sql_bairro_com&", "& sql_cid_com &", SG_UF_Com ='"& uf_com_fam_aux &"', "
sql_atualiza=sql_atualiza&"CO_CEP_Com='"& cep_com_fam_aux &"', NU_Telefones_Com ='"& tel_com_fam_aux&"', "&sql_vinc&", TP_Contato_Vinc ='"& tp_vinc_familiar_aux&"' WHERE CO_Matricula = "& cod &" AND TP_Contato = '"& cod_familiar_aux &"'"

else
sql_atualiza= "UPDATE TBI_Contatos SET NO_Contato = '"&nome_familiar_aux&"', "& sql_nasce &", CO_CPF_PFisica ='"& cpf_fam_aux &"', CO_RG_PFisica ='"& rg_fam_aux &"', CO_OERG_PFisica ='"& emitido_fam_aux&"', "& sql_emissao &", TX_EMail ='"& email_fam_aux &"', "&sql_ocupacao&", NO_Empresa ='"& empresa_fam_aux &"', NU_Telefones ='"& tel_fam_aux&"', ID_Res_Aluno = "&id_res_fam_aux&", "
if id_res_fam_aux=FALSE then
sql_atualiza=sql_atualiza&"NO_Logradouro_Res ='"& rua_res_fam_aux &"', "& sql_num_res&", TX_Complemento_Logradouro_Res = '"&comp_res_fam_aux&"', "& sql_bairro_res &", "& sql_cid_res &", SG_UF_Res ='"& uf_res_fam_aux &"', CO_CEP_Res ='"& cep_res_fam_aux &"', NU_Telefones_Res ='"& tel_res_fam_aux&"', "
END IF
sql_atualiza=sql_atualiza&"NO_Logradouro_Com = '"&rua_com_fam_aux&"', "& sql_num_com&", TX_Complemento_Logradouro_Com= '"&comp_com_fam_aux&"',"&sql_bairro_com&", "& sql_cid_com &", SG_UF_Com ='"& uf_com_fam_aux &"', "
sql_atualiza=sql_atualiza&"CO_CEP_Com='"& cep_com_fam_aux &"', NU_Telefones_Com ='"& tel_com_fam_aux&"' WHERE CO_Matricula = "& co_vinc_familiar_aux &" AND TP_Contato = '"& tp_vinc_familiar_aux &"'"
end if

'sql_atualiza= "UPDATE TBI_Contatos SET NU_Telefones ='"& tel_familiar_aux&"', ID_Res_Aluno = "&id_res_familiar_aux&", NO_Logradouro_Res ='"& rua_res_familiar_aux &"', "& sql_num_res&", TX_Complemento_Logradouro_Res = '"&comp_res_familiar_aux&"', "& sql_bairro_res &", "& sql_cid_res &", SG_UF_Res ='"& uf_res_familiar_aux &"' WHERE CO_Matricula = "& cod &" AND TP_Contato = '"& cod_familiar &"'"


response.Write(sql_atualiza)

Set RSCONTATO_aux_bd2 = CONCONT_aux.Execute(sql_atualiza)
'if do RSCONTATO_aux.EOF

'response.Write(Server.URLEncode(sql_atualiza_tx1))

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

		Set RSCONTATO = Server.CreateObject("ADODB.Recordset")
		SQLAA= "DELETE * FROM TBI_Contatos WHERE TP_Contato_Vinc='"&foco&"' and CO_Matricula_Vinc ="&cod
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