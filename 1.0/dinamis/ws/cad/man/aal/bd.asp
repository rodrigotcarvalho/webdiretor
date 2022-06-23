<!--#include file="../../../../inc/connect_ct.asp"-->
<!--#include file="../../../../inc/connect_ct_aux.asp"-->
<!--#include file="../../../../inc/connect_pr.asp"-->
<!--#include file="../../../../inc/connect_al_aux.asp"-->
<!--#include file="../../../../inc/connect_al.asp"-->
<%
chave=session("nvg")
session("nvg")=chave
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo")
ano_letivo_real = ano_letivo
sistema_local=session("sistema_local")
opt=request.querystring("opt")

			ano = DatePart("yyyy", now)
			mes = DatePart("m", now) 
			dia = DatePart("d", now) 
			hora = DatePart("h", now) 
			min = DatePart("n", now) 

			data_cadastro = ano&"/"&mes&"/"&dia 

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0

		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1		
		
		Set CONCONT_aux = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT_aux = "DBQ="& CAMINHO_ct_aux & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT_aux.Open ABRIRCONT_aux
		
		Set CON1_aux = Server.CreateObject("ADODB.Connection") 
		ABRIR1_aux = "DBQ="& CAMINHO_al_aux & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1_aux.Open ABRIR1_aux

if opt="i" then

if isnull(da_entrada) or da_entrada="" then
da_entrada =NULL
end if

			ano = DatePart("yyyy", now)
			mes = DatePart("m", now) 
			dia = DatePart("d", now) 
			da_cadastro =dia&"/"& mes &"/"& ano
'			da_cadastro =mes&"/"&dia&"/"& ano
response.Write(da_cadastro)
nome_aluno=request.form("nome")
'cod=request.form("codigo")
'Caso alguém crie um aluno entre uma tela e outra eu capturo isso com o select abaixo.

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT Max(CO_Matricula) AS COD FROM TB_Alunos"
		RS.Open SQL, CON1
			
cod = RS("COD")
cod=cod*1
cod=cod+1
		
Set RSALUNO_bd = server.createobject("adodb.recordset")
RSALUNO_bd.open "TB_Alunos", CON1, 2, 2
RSALUNO_bd.addnew
RSALUNO_bd("CO_Matricula")=cod
RSALUNO_bd("RA_Aluno")=cod
RSALUNO_bd("NO_Aluno")=nome_aluno
RSALUNO_bd("DA_Cadastro")=data_cadastro
RSALUNO_bd("CO_Matricula_Vinc")=NULL
RSALUNO_bd.update
  
set RSALUNO_bd=nothing

Set RSCONTATO_bd = server.createobject("adodb.recordset")
RSCONTATO_bd.open "TB_Contatos", CONCONT, 2, 2 'which table do you want open
RSCONTATO_bd.addnew
RSCONTATO_bd("CO_Matricula")=cod
RSCONTATO_bd("TP_Contato")="ALUNO"
RSCONTATO_bd("NO_Contato")=nome_aluno
RSCONTATO_bd("CO_Matricula_Vinc")=NULL
RSCONTATO_bd("TP_Contato_Vinc")=NULL
RSCONTATO_bd.update
  
set RSCONTATO_bd=nothing

response.redirect("altera.asp?opt=ok2&cod_cons="&cod)			

elseif opt="a" then

'Grava dados preenchidos no form para tabelas temporárias==========================================================================

'familiar na tela
cod=request.form("cod")
nome_familiar=request.form("nome_familiar")
nasce_fam=request.form("nasce_fam")
cod_ultimo_familiar=request.form("cod_familiar")
ocupacao_fam=request.form("ocupacao_fam")
trabalho_fam=request.form("trabalho_fam")
email_fam=request.form("email_fam")
cpf_fam=request.form("cpf_fam")
rg_fam=request.form("id_fam")
emitido_fam=request.form("tipo_id_fam")
emissao_fam=request.form("nasce2_fam")
tel_cont_fam=request.form("tel_fam")
rua_res_fam=request.form("rua_res_fam")
num_res_fam=request.form("num_res_fam")
comp_res_fam=request.form("comp_res_fam")
uf_res_fam=request.form("estadores_fam")
cid_res_fam=request.form("cidres_fam")
bairro_res_fam=request.form("bairrores_fam")
cep_fam=request.form("cep_fam")
tel_res_fam=request.form("tel_res_fam")
id_res_fam=request.form("id_res_fam_aux")
rua_com_fam=request.form("rua_com_fam")
num_com_fam=request.form("num_com_fam")
comp_com_fam=request.form("comp_com_fam")
uf_com_fam=request.form("estadocom_fam")
cid_com_fam=request.form("cidcom_fam")
bairro_com_fam=request.form("bairrocom_fam")
cepcom_fam=request.form("cepcom_fam")
tel_com_fam=request.form("tel_com_fam")
tp_vinc_familiar_aux =request.form("tp_vinc_familiar_aux")
co_vinc_familiar_aux =request.form("co_vinc_familiar_aux")

'responsáveis financeiro e pedagógico
responsavel_financeiro=request.form("rf")  
responsavel_pedagogico=request.form("rp")   
id_familia=request.form("id_familia")
bloq=request.form("bloq")
circ =request.form("circ")




			if id_res_fam="s" then
				id_res_fam="TRUE"
			else
				id_res_fam="FALSE"
			end if

if isnull(nasce_fam) or nasce_fam="" then
else
vetor_nascimento = Split(nasce_fam,"/")  
dia_n = vetor_nascimento(0)
mes_n = vetor_nascimento(1)
ano_n = vetor_nascimento(2)

dia_a = dia_n
mes_a = mes_n
ano_a = ano_n

nasce_fam = dia_n&"/"&mes_n&"/"&ano_n
'nasce_fam = mes_n&"/"&dia_n&"/"&ano_n
end if



if isnull(emissao_fam) or emissao_fam="" then
else
vetor_nasce2 = Split(emissao_fam,"/")  
dia_e = vetor_nasce2(0)
mes_e = vetor_nasce2(1)
ano_e = vetor_nasce2(2)

dia_a = dia_e
mes_a = mes_e
ano_a = ano_e
emissao_fam = dia_e&"/"&mes_e&"/"&ano_e
'emissao_fam = mes_e&"/"&dia_e&"/"&ano_e
end if



if isnull(cod_ultimo_familiar) or cod_ultimo_familiar="" then
else

			Set RSCONTATO_aux = Server.CreateObject("ADODB.Recordset")
			SQLAA_aux= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_ultimo_familiar&"' and CO_Matricula ="&cod
			RSCONTATO_aux.Open SQLAA_aux, CONCONT_aux
			
			
	
	if RSCONTATO_aux.EOF then
	
		if isnull(nasce_fam) or nasce_fam="" then
		nasce_fam =NULL
		end if
		
		if isnull(emissao_fam) or emissao_fam="" then
		emissao_fam=NULL
		end if
		
		if isnull(ocupacao_fam) or ocupacao_fam="" then
		ocupacao_fam =NULL
		end if
		
		if isnull(num_res_fam) or num_res_fam="" then
		num_res_fam =NULL
		end if
		
		if isnull(bairrores_fam) or bairrores_fam="" then
		bairrores_fam =NULL
		end if
		
		if isnull(cidres_fam) or cidres_fam="" then
		cidres_fam =NULL
		end if
		
		if isnull(num_com_fam) or num_com_fam="" then
		num_com_fam =NULL
		end if
		
		if isnull(cidcom_fam) or cidcom_fam="" then
		cidcom_fam=NULL
		end if
		
		if isnull(bairrocom_fam) or bairrocom_fam="" then
		bairrocom_fam =NULL
		end if
	
		if isnull(co_vinc_familiar_aux) or co_vinc_familiar_aux="" then
		co_vinc_familiar_aux =NULL
		familiar_tela_vinculado="n"
		else 
		familiar_tela_vinculado="s"
		end if
	
		if isnull(tp_vinc_familiar_aux) or tp_vinc_familiar_aux="" then
		tp_vinc_familiar_aux =NULL
		familiar_tela_vinculado="n"
		else 
		familiar_tela_vinculado="s"
		end if
	

	
		if familiar_tela_vinculado="n" then
				Set RSCONTATO_aux_bd = server.createobject("adodb.recordset")
				RSCONTATO_aux_bd.open "TBI_Contatos", CONCONT_aux, 2, 2 'which table do you want open
				RSCONTATO_aux_bd.addnew
				RSCONTATO_aux_bd("CO_Matricula")=cod
				RSCONTATO_aux_bd("TP_Contato")=cod_ultimo_familiar
				RSCONTATO_aux_bd("NO_Contato")=nome_familiar
				RSCONTATO_aux_bd("DA_Nascimento_Contato")=nasce_fam
				RSCONTATO_aux_bd("CO_CPF_PFisica")=cpf_fam
				RSCONTATO_aux_bd("CO_RG_PFisica")=rg_fam
				RSCONTATO_aux_bd("CO_OERG_PFisica")=emitido_fam
				RSCONTATO_aux_bd("CO_DERG_PFisica")=emissao_fam
				RSCONTATO_aux_bd("TX_EMail")=email_fam
				RSCONTATO_aux_bd("CO_Ocupacao")=ocupacao_fam
				RSCONTATO_aux_bd("NO_Empresa")=trabalho_fam
				RSCONTATO_aux_bd("NU_Telefones")=tel_cont_fam
				RSCONTATO_aux_bd("ID_Res_Aluno")=id_res_fam
				if id_res_fam="FALSE" then
				RSCONTATO_aux_bd("NO_Logradouro_Res")=rua_res_fam
				RSCONTATO_aux_bd("NU_Logradouro_Res")=num_res_fam
				RSCONTATO_aux_bd("TX_Complemento_Logradouro_Res")=comp_res_fam
				RSCONTATO_aux_bd("CO_Bairro_Res")=bairro_res_fam
				RSCONTATO_aux_bd("CO_Municipio_Res")=cid_res_fam
				RSCONTATO_aux_bd("SG_UF_Res")=uf_res_fam
				RSCONTATO_aux_bd("CO_CEP_Res")=cep_fam
				RSCONTATO_aux_bd("NU_Telefones_Res")=tel_res_fam
				END IF
				RSCONTATO_aux_bd("NO_Logradouro_Com")=rua_com_fam
				RSCONTATO_aux_bd("NU_Logradouro_Com")=num_com_fam
				RSCONTATO_aux_bd("TX_Complemento_Logradouro_Com")=comp_com_fam
				RSCONTATO_aux_bd("CO_Bairro_Com")=bairro_com_fam
				RSCONTATO_aux_bd("CO_Municipio_Com")=cid_com_fam
				RSCONTATO_aux_bd("SG_UF_Com")=uf_com_fam
				RSCONTATO_aux_bd("CO_CEP_Com")=cepcom_fam
				RSCONTATO_aux_bd("NU_Telefones_Com")=tel_com_fam
				RSCONTATO_aux_bd.update		
				set RSCONTATO_aux_bd=nothing
		else
		

				Set RSCONTATO_aux_bd = server.createobject("adodb.recordset")
				RSCONTATO_aux_bd.open "TBI_Contatos", CONCONT_aux, 2, 2 'which table do you want open
				RSCONTATO_aux_bd.addnew
				RSCONTATO_aux_bd("CO_Matricula")=cod
				RSCONTATO_aux_bd("TP_Contato")=cod_ultimo_familiar
				RSCONTATO_aux_bd("CO_Matricula_Vinc")=co_vinc_familiar_aux
				RSCONTATO_aux_bd("TP_Contato_Vinc")=tp_vinc_familiar_aux
				RSCONTATO_aux_bd.update		
				set RSCONTATO_aux_bd=nothing
				
				Set RSCONTATO_aux = Server.CreateObject("ADODB.Recordset")
				SQLAA_aux= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_vinc_familiar_aux&"' and CO_Matricula ="&co_vinc_familiar_aux
				RSCONTATO_aux.Open SQLAA_aux, CONCONT_aux
				
				if 	RSCONTATO_aux.eof then	
					Set RSCONTATO_aux_bd = server.createobject("adodb.recordset")
					RSCONTATO_aux_bd.open "TBI_Contatos", CONCONT_aux, 2, 2 'which table do you want open
					RSCONTATO_aux_bd.addnew
					RSCONTATO_aux_bd("CO_Matricula")=co_vinc_familiar_aux
					RSCONTATO_aux_bd("TP_Contato")=tp_vinc_familiar_aux
					RSCONTATO_aux_bd("NO_Contato")=nome_familiar
					RSCONTATO_aux_bd("DA_Nascimento_Contato")=nasce_fam
					RSCONTATO_aux_bd("CO_CPF_PFisica")=cpf_fam
					RSCONTATO_aux_bd("CO_RG_PFisica")=rg_fam
					RSCONTATO_aux_bd("CO_OERG_PFisica")=emitido_fam
					RSCONTATO_aux_bd("CO_DERG_PFisica")=emissao_fam
					RSCONTATO_aux_bd("TX_EMail")=email_fam
					RSCONTATO_aux_bd("CO_Ocupacao")=ocupacao_fam
					RSCONTATO_aux_bd("NO_Empresa")=trabalho_fam
					RSCONTATO_aux_bd("NU_Telefones")=tel_cont_fam
					RSCONTATO_aux_bd("ID_Res_Aluno")=id_res_fam
					if id_res_fam="FALSE" then
					RSCONTATO_aux_bd("NO_Logradouro_Res")=rua_res_fam
					RSCONTATO_aux_bd("NU_Logradouro_Res")=num_res_fam
					RSCONTATO_aux_bd("TX_Complemento_Logradouro_Res")=comp_res_fam
					RSCONTATO_aux_bd("CO_Bairro_Res")=bairro_res_fam
					RSCONTATO_aux_bd("CO_Municipio_Res")=cid_res_fam
					RSCONTATO_aux_bd("SG_UF_Res")=uf_res_fam
					RSCONTATO_aux_bd("CO_CEP_Res")=cep_fam
					RSCONTATO_aux_bd("NU_Telefones_Res")=tel_res_fam
					END IF
					RSCONTATO_aux_bd("NO_Logradouro_Com")=rua_com_fam
					RSCONTATO_aux_bd("NU_Logradouro_Com")=num_com_fam
					RSCONTATO_aux_bd("TX_Complemento_Logradouro_Com")=comp_com_fam
					RSCONTATO_aux_bd("CO_Bairro_Com")=bairro_com_fam
					RSCONTATO_aux_bd("CO_Municipio_Com")=cid_com_fam
					RSCONTATO_aux_bd("SG_UF_Com")=uf_com_fam
					RSCONTATO_aux_bd("CO_CEP_Com")=cepcom_fam
					RSCONTATO_aux_bd("NU_Telefones_Com")=tel_com_fam
					RSCONTATO_aux_bd.update		
					set RSCONTATO_aux_bd=nothing
				else
					Set RSCONTATO_aux_bd2 = server.createobject("adodb.recordset")
					sql_atualiza= "UPDATE TBI_Contatos SET NO_Contato = '"&nome_familiar_aux&"', "& sql_nasce &", CO_CPF_PFisica ='"& cpf_familiar_aux &"', CO_RG_PFisica ='"& rg_familiar_aux &"', CO_OERG_PFisica ='"& emitido_familiar_aux&"', "& sql_emissao &", TX_EMail ='"& email_familiar_aux &"', "&sql_ocupacao&", NO_Empresa ='"& empresa_familiar_aux &"', "
					sql_atualiza=sql_atualiza&"NU_Telefones ='"& tel_familiar_aux&"', ID_Res_Aluno = "&mes_end&", NO_Logradouro_Res ='"& rua_res_familiar_aux &"', "& sql_num_res&", TX_Complemento_Logradouro_Res = '"&comp_res_familiar_aux&"', "& sql_bairro_res &", "& sql_cid_res &", SG_UF_Res ='"& uf_res_familiar_aux &"', "
					sql_atualiza=sql_atualiza&"CO_CEP_Res ='"& cep_res_familiar_aux &"', NU_Telefones_Res ='"& tel_res_familiar_aux&"', NO_Logradouro_Com = '"&rua_com_familiar_aux&"', "& sql_num_com&", TX_Complemento_Logradouro_Com= '"&comp_com_familiar_aux&"',"&sql_bairro_com&", "& sql_cid_com &", SG_UF_Com ='"& uf_com_familiar_aux &"', "
					sql_atualiza=sql_atualiza&"CO_CEP_Com='"& cep_com_familiar_aux &"', NU_Telefones_Com ='"& tel_com_familiar_aux&"' WHERE CO_Matricula = "& co_vinc_familiar_aux &" AND TP_Contato = '"& tp_vinc_familiar_aux &"'"
					Set RSCONTATO_aux_bd2 = CONCONT_aux.Execute(sql_atualiza)	
		
				
				end if
				
		end if		
	'response.Write("ERRO")
	
	
	else	

		
		if isnull(nasce_fam) or nasce_fam="" then
		sql_nasce="DA_Nascimento_Contato =NULL"
		else
		sql_nasce="DA_Nascimento_Contato =#"& nasce_fam &"#"
		end if
		
		if isnull(emissao_fam) or emissao_fam="" then
		sql_emissao="CO_DERG_PFisica =NULL"
		else
		sql_emissao="CO_DERG_PFisica =#"& emissao_fam &"#"
		end if
		
		if isnull(ocupacao_fam) or ocupacao_fam="" then
		sql_ocupacao="CO_Ocupacao =NULL"
		else
		sql_ocupacao="CO_Ocupacao ="& ocupacao_fam &""
		end if
		
		if isnull(num_res_fam) or num_res_fam="" then
		sql_num_res="NU_Logradouro_Res =NULL"
		else
		sql_num_res="NU_Logradouro_Res ="& num_res_fam &""
		end if
		
		if isnull(bairro_res_fam) or bairro_res_fam="" then
		sql_bairro_res=" CO_Bairro_Res =NULL"
		else
		sql_bairro_res=" CO_Bairro_Res ="& bairro_res_fam &""
		end if
		
		if isnull(cid_res_fam) or cid_res_fam="" then
		sql_cid_res=" CO_Municipio_Res =NULL"
		else
		sql_cid_res=" CO_Municipio_Res ="& cid_res_fam &""
		end if
		
		if isnull(num_com_fam) or num_com_fam="" then
		sql_num_com="NU_Logradouro_Com =NULL"
		else
		sql_num_com="NU_Logradouro_Com ="& num_com_fam &""
		end if
		
		if isnull(cid_com_fam) or cid_com_fam="" then
		sql_cid_com=" CO_Municipio_Com =NULL"
		else
		sql_cid_com=" CO_Municipio_Com ="& cid_com_fam &""
		end if
		
		if isnull(bairro_com_fam) or bairro_com_fam="" then
		sql_bairro_com=" CO_Bairro_Com =NULL"
		else
		sql_bairro_com=" CO_Bairro_Com ="& bairro_com_fam &""
		end if
		
		if isnull(co_vinc_familiar_aux) or co_vinc_familiar_aux="" then
		sql_vinc="CO_Matricula_Vinc =NULL"
		familiar_tela_vinculado="n"
		else
		sql_vinc="CO_Matricula_Vinc ="& co_vinc_familiar_aux &""
		familiar_tela_vinculado="s"
		end if
		
		if (isnull(co_vinc_familiar_aux) or co_vinc_familiar_aux="NULL" or cod_vinc="") and (isnull(tp_vinc_familiar_aux) or tp_vinc_familiar_aux="NULL" or tp_vinc_familiar_aux="") then
			Set RSCONTATO_aux_bd2 = server.createobject("adodb.recordset")
			sql_atualiza= "UPDATE TBI_Contatos SET NO_Contato = '"&nome_familiar&"', "& sql_nasce &", CO_CPF_PFisica ='"& cpf_fam &"', CO_RG_PFisica ='"& rg_fam&"', CO_OERG_PFisica ='"& emitido_fam&"', "& sql_emissao &", TX_EMail ='"& email_fam &"', "&sql_ocupacao&", NO_Empresa ='"& trabalho_fam &"', NU_Telefones ='"& tel_cont_fam&"', ID_Res_Aluno = "&id_res_fam&", "
				if id_res_fam="FALSE" then
				sql_atualiza=sql_atualiza&"NO_Logradouro_Res ='"& rua_res_fam &"', "& sql_num_res&", TX_Complemento_Logradouro_Res = '"&comp_res_fam&"', "& sql_bairro_res &", "& sql_cid_res &", SG_UF_Res ='"& uf_res_fam &"', CO_CEP_Res ='"& cep_fam &"', NU_Telefones_Res ='"& tel_res_fam&"', "
				END IF
			sql_atualiza=sql_atualiza&"NO_Logradouro_Com = '"&rua_com_fam&"', "& sql_num_com&", TX_Complemento_Logradouro_Com= '"&comp_com_fam&"',"&sql_bairro_com&", "& sql_cid_com &", SG_UF_Com ='"& uf_com_fam &"', "
			sql_atualiza=sql_atualiza&"CO_CEP_Com='"& cepcom_fam &"', CO_Matricula_Vinc =NULL, TP_Contato_Vinc =NULL, NU_Telefones_Com ='"& tel_com_fam&"' WHERE CO_Matricula = "& cod &" AND TP_Contato = '"& cod_ultimo_familiar &"'"
			response.Write(sql_atualiza)
			Set RSCONTATO_aux_bd2 = CONCONT_aux.Execute(sql_atualiza)
		
		else
		
			Set RSCONTATO_aux_bd2 = server.createobject("adodb.recordset")
			sql_atualiza= "UPDATE TBI_Contatos SET NO_Contato = '"&nome_familiar&"', "& sql_nasce &", CO_CPF_PFisica ='"& cpf_fam &"', CO_RG_PFisica ='"& rg_fam&"', CO_OERG_PFisica ='"& emitido_fam&"', "& sql_emissao &", TX_EMail ='"& email_fam &"', "&sql_ocupacao&", NO_Empresa ='"& trabalho_fam &"', NU_Telefones ='"& tel_cont_fam&"', ID_Res_Aluno = "&id_res_fam&", "
				if id_res_fam="FALSE" then
				sql_atualiza=sql_atualiza&"NO_Logradouro_Res ='"& rua_res_fam &"', "& sql_num_res&", TX_Complemento_Logradouro_Res = '"&comp_res_fam&"', "& sql_bairro_res &", "& sql_cid_res &", SG_UF_Res ='"& uf_res_fam &"', CO_CEP_Res ='"& cep_fam &"', NU_Telefones_Res ='"& tel_res_fam&"', "
				END IF
			sql_atualiza=sql_atualiza&"NO_Logradouro_Com = '"&rua_com_fam&"', "& sql_num_com&", TX_Complemento_Logradouro_Com= '"&comp_com_fam&"',"&sql_bairro_com&", "& sql_cid_com &", SG_UF_Com ='"& uf_com_fam &"', "
			sql_atualiza=sql_atualiza&"CO_CEP_Com='"& cepcom_fam &"', CO_Matricula_Vinc =NULL, TP_Contato_Vinc =NULL, NU_Telefones_Com ='"& tel_com_fam&"' WHERE CO_Matricula = "& co_vinc_familiar_aux &" AND TP_Contato = '"& tp_vinc_familiar_aux &"'"
			Set RSCONTATO_aux_bd2 = CONCONT_aux.Execute(sql_atualiza)
	'response.Write("DELETE * FROM TBI_Contatos WHERE TP_Contato='"&cod_ultimo_familiar&"' and CO_Matricula ="&cod)
	'response.end()
					Set RSCONTATO_aux_delete = Server.CreateObject("ADODB.Recordset")
					SQLAC_delete= "DELETE * FROM TBI_Contatos WHERE TP_Contato='"&cod_ultimo_familiar&"' and CO_Matricula ="&cod
					RSCONTATO_aux_delete.Open SQLAC_delete, CONCONT_aux
			
				Set RSCONTATO_aux_bd = server.createobject("adodb.recordset")
				RSCONTATO_aux_bd.open "TBI_Contatos", CONCONT_aux, 2, 2 'which table do you want open
				RSCONTATO_aux_bd.addnew
				RSCONTATO_aux_bd("CO_Matricula")=cod
				RSCONTATO_aux_bd("TP_Contato")=cod_ultimo_familiar
				RSCONTATO_aux_bd("CO_Matricula_Vinc")=co_vinc_familiar_aux
				RSCONTATO_aux_bd("TP_Contato_Vinc")=tp_vinc_familiar_aux
				RSCONTATO_aux_bd.update		
				set RSCONTATO_aux_bd=nothing
				
		end if	
	
	END IF

'end do if isnull(cod_ultimo_familiar)
END IF




cod=request.form("codigo")
 nome_aluno=request.form("nome")
 apelido=request.form("apelido")
 sexo=request.form("sexo")
 nasce=request.form("nasce")
 nacionalidade=request.form("nacionalidade")
 cidnat=request.form("cidnat")
 cor_raca=request.form("cor_raca")
 rg=request.form("rg")
 cpf=request.form("cpf")
 email=request.form("email")    
 orkut=request.form("orkut")
 tel=request.form("tel")
 pais=request.form("pais")
 estadonat=request.form("estadonat")
 religiao=request.form("religiao")
 ocupacao=request.form("ocupacao")
 tipo_id=request.form("tipo_id")
 nasce2=request.form("nasce2")
 empresa=request.form("trabalho")
 messenger=request.form("messenger") 
 desteridade=request.form("desteridade")
 vinculado=request.form("vinculado")
 rua_res=request.form("rua_res")
 num_res=request.form("num_res")
 comp_res=request.form("comp_res")
 bairro_res=request.form("bairrores")
 cid_res=request.form("cidres")
 uf_res=request.form("estadores")
 cep=request.form("cep")
 tel_res=request.form("tel_res")
 rua_com=request.form("rua_com")    
 num_com=request.form("num_com")
 comp_com=request.form("comp_com")
 bairro_com=request.form("bairrocom")
 cid_com=request.form("cidcom")
 uf_com=request.form("estadocom")
 cep_com=request.form("cep_com")
 tel_com=request.form("tel_com")
 pai=request.form("pai")
 mae=request.form("mae")
 pai_falecido=request.form("pai_falecido")  
 mae_falecido=request.form("mae_falecido")    
 sit_pais=request.form("sit_pais")

col_or=request.form("col_or")   
et_curs=request.form("et_curs")   
uf_curs=request.form("estadocurs")   
cid_curs=request.form("cid_curs")   
da_entrada=request.form("da_entrada")   
da_cadastro =request.form("da_cadastro")

'if isnull(da_entrada) or da_entrada="" then
'da_entrada =NULL
'end if

if isnull(da_cadastro) or da_cadastro="" then
			ano = DatePart("yyyy", now)
			mes = DatePart("m", now) 
			dia = DatePart("d", now) 
			'hora = DatePart("h", now) 
			'min = DatePart("n", now) 
			da_cadastro =dia&"/"& mes &"/"& ano
'			da_cadastro =mes&"/"&dia &"/"& ano			
end if

'comentado por último===============================================================================================================	
'cod_familiar="ALUNO"
'==============================================================================================================	


if id_res="s" then
id_res=TRUE
else
id_res=FALSE
end if

if pai_falecido="s" then
pai_falecido=TRUE
else
pai_falecido=FALSE
end if

if mae_falecido="s" then
mae_falecido=TRUE
else
mae_falecido=FALSE
end if


'dados do aluno

		Set RSCONTATO_aux_delete = Server.CreateObject("ADODB.Recordset")
		SQLAC_delete= "DELETE * FROM TBI_Contatos WHERE TP_Contato='ALUNO' and CO_Matricula ="&cod
		RSCONTATO_aux_delete.Open SQLAC_delete, CONCONT_aux


		Set RSCONTATO_aux = Server.CreateObject("ADODB.Recordset")
		SQLAA_aux= "SELECT * FROM TBI_Contatos WHERE TP_Contato='ALUNO' and CO_Matricula ="&cod
		RSCONTATO_aux.Open SQLAA_aux, CONCONT_aux

if isnull(nasce) or nasce="" then
nasce =NULL
end if

if isnull(nasce2) or nasce2="" then
nasce2=NULL
end if

if isnull(ocupacao) or ocupacao="" then
ocupacao =NULL
end if

if isnull(num_res) or num_res="" then
num_res =NULL
end if

if isnull(bairro_res) or bairro_res="" then
bairro_res =NULL
end if

if isnull(cid_res) or cid_res="" then
cid_res =NULL
end if

if isnull(num_com) or num_com="" then
num_com =NULL
end if

if isnull(cid_com) or cid_com="" then
cid_com=NULL
end if

if isnull(bairro_com) or bairro_com="" then
bairro_com =NULL
end if

if isnull(da_entrada) or da_entrada="" then
da_entrada =NULL
end if


Set RSCONTATO_aux_bd = server.createobject("adodb.recordset")
	RSCONTATO_aux_bd.open "TBI_Contatos", CONCONT_aux, 2, 2 'which table do you want open
	RSCONTATO_aux_bd.addnew
	RSCONTATO_aux_bd("CO_Matricula")=cod
	RSCONTATO_aux_bd("TP_Contato")="ALUNO"
	RSCONTATO_aux_bd("NO_Contato")=nome_aluno
	RSCONTATO_aux_bd("DA_Nascimento_Contato")=nasce
	RSCONTATO_aux_bd("CO_CPF_PFisica")=cpf
	RSCONTATO_aux_bd("CO_RG_PFisica")=rg
	RSCONTATO_aux_bd("CO_OERG_PFisica")=tipo_id
	RSCONTATO_aux_bd("CO_DERG_PFisica")=nasce2
	RSCONTATO_aux_bd("TX_EMail")=email
	RSCONTATO_aux_bd("CO_Ocupacao")=ocupacao
	RSCONTATO_aux_bd("NO_Empresa")=empresa
	RSCONTATO_aux_bd("NU_Telefones")=tel
	RSCONTATO_aux_bd("ID_Res_Aluno")=id_res
	if isnull(vinculado) or vinculado="" then
		RSCONTATO_aux_bd("NO_Logradouro_Res")=rua_res
		RSCONTATO_aux_bd("NU_Logradouro_Res")=num_res
		RSCONTATO_aux_bd("TX_Complemento_Logradouro_Res")=comp_res
		RSCONTATO_aux_bd("CO_Bairro_Res")=bairro_res
		RSCONTATO_aux_bd("CO_Municipio_Res")=cid_res
		RSCONTATO_aux_bd("SG_UF_Res")=uf_res
		RSCONTATO_aux_bd("CO_CEP_Res")=cep
		RSCONTATO_aux_bd("NU_Telefones_Res")=tel_res
		RSCONTATO_aux_bd("NO_Logradouro_Com")=rua_com
		RSCONTATO_aux_bd("NU_Logradouro_Com")=num_com
		RSCONTATO_aux_bd("TX_Complemento_Logradouro_Com")=comp_com
		RSCONTATO_aux_bd("CO_Bairro_Com")=bairro_com
		RSCONTATO_aux_bd("CO_Municipio_Com")=cid_com
		RSCONTATO_aux_bd("SG_UF_Com")=uf_com
		RSCONTATO_aux_bd("CO_CEP_Com")=cep_com
		RSCONTATO_aux_bd("NU_Telefones_Com")=tel_com
	else
		RSCONTATO_aux_bd("CO_Matricula_Vinc")=vinculado
		RSCONTATO_aux_bd("TP_Contato_Vinc")="ALUNO"
	end if
	RSCONTATO_aux_bd.update
	set RSCONTATO_aux_bd=nothing
	
	if isnull(vinculado) or vinculado="" then
	else
		if isnull(num_res) or num_res="" then
		sql_num_res="NU_Logradouro_Res =NULL"
		else
		sql_num_res="NU_Logradouro_Res ="& num_res &""
		end if
		
		if isnull(bairro_res) or bairro_res="" then
		sql_bairro_res=" CO_Bairro_Res =NULL"
		else
		sql_bairro_res=" CO_Bairro_Res ="& bairro_res &""
		end if
		
		if isnull(cid_res) or cid_res="" then
		sql_cid_res=" CO_Municipio_Res =NULL"
		else
		sql_cid_res=" CO_Municipio_Res ="& cid_res &""
		end if
		
		if isnull(num_com) or num_com="" then
		sql_num_com="NU_Logradouro_Com =NULL"
		else
		sql_num_com="NU_Logradouro_Com ="& num_com &""
		end if
		
		if isnull(bairro_com) or bairro_com="" then
		sql_bairro_com=" CO_Bairro_Com =NULL"
		else
		sql_bairro_com=" CO_Bairro_Com ="& bairro_com &""
		end if
		
		if isnull(cid_com) or cid_com="" then
		sql_cid_com=" CO_Municipio_Com =NULL"
		else
		sql_cid_com=" CO_Municipio_Com ="& cid_com &""
		end if
					
			Set RSCONTATO_aux_bd2 = server.createobject("adodb.recordset")
			sql_atualiza= "UPDATE TBI_Contatos SET NO_Logradouro_Res ='"& rua_res &"', "& sql_num_res&", TX_Complemento_Logradouro_Res = '"&comp_res&"', "& sql_bairro_res &", "& sql_cid_res &", SG_UF_Res ='"& uf_res &"', "
			sql_atualiza=sql_atualiza&"CO_CEP_Res ='"& cep &"', NU_Telefones_Res ='"& tel_res&"', NO_Logradouro_Com = '"&rua_com&"', "& sql_num_com&", TX_Complemento_Logradouro_Com= '"&comp_com&"',"&sql_bairro_com&", "& sql_cid_com &", SG_UF_Com ='"& uf_com &"', "
			sql_atualiza=sql_atualiza&"CO_CEP_Com='"& cep_com &"', NU_Telefones_Com ='"& tel_com&"' WHERE CO_Matricula = "& vinculado &" AND TP_Contato = 'ALUNO'"
			Set RSCONTATO_aux_bd2 = CONCONT_aux.Execute(sql_atualiza)

	end if


		Set RSCONTATO_aux_delete = Server.CreateObject("ADODB.Recordset")
		SQLAA_delete= "DELETE * FROM TBI_Alunos WHERE CO_Matricula ="&cod
		RSCONTATO_aux_delete.Open SQLAA_delete, CON1_aux

		Set RS_aux = Server.CreateObject("ADODB.Recordset")
		SQL_aux = "SELECT * FROM TBI_Alunos WHERE CO_Matricula ="& cod
		RS_aux.Open SQL_aux, CON1_aux
		
if RS_aux.EOF Then

Set RSALUNO_aux_bd = server.createobject("adodb.recordset")
RSALUNO_aux_bd.open "TBI_Alunos", CON1_aux, 2, 2
RSALUNO_aux_bd.addnew
RSALUNO_aux_bd("CO_Matricula")=cod
RSALUNO_aux_bd("RA_Aluno")=cod
RSALUNO_aux_bd("NO_Aluno")=nome_aluno
RSALUNO_aux_bd("NO_Apelido")=apelido
RSALUNO_aux_bd("IN_Sexo")=sexo
RSALUNO_aux_bd("IN_Desteridade")=desteridade
RSALUNO_aux_bd("CO_Nacionalidade")=nacionalidade
RSALUNO_aux_bd("CO_Pais_Natural")=pais
RSALUNO_aux_bd("SG_UF_Natural")=estadonat
RSALUNO_aux_bd("CO_Municipio_Natural")=cidnat
RSALUNO_aux_bd("TX_MSN")=messenger
RSALUNO_aux_bd("TX_ORKUT")=orkut
RSALUNO_aux_bd("CO_Raca")=cor_raca
RSALUNO_aux_bd("CO_Religiao")=religiao
if isnull(vinculado) or vinculado="" then
RSALUNO_aux_bd("NO_Pai")=pai
RSALUNO_aux_bd("NO_Mae")=mae
RSALUNO_aux_bd("IN_Pai_Falecido")=pai_falecido
RSALUNO_aux_bd("IN_Mae_Falecida")=mae_falecido
RSALUNO_aux_bd("CO_Estado_Civil")=sit_pais
RSALUNO_aux_bd("TP_Resp_Fin")=responsavel_financeiro							  
RSALUNO_aux_bd("TP_Resp_Ped")=responsavel_pedagogico							  
end if
RSALUNO_aux_bd("NO_Colegio_Origem")= col_or
RSALUNO_aux_bd("NO_Serie_Cursada")=et_curs
RSALUNO_aux_bd("SG_UF_Cursada")=uf_curs
RSALUNO_aux_bd("CO_Municipio_Cursada")=cid_curs
RSALUNO_aux_bd("DA_Entrada_Escola")=da_entrada
RSALUNO_aux_bd("DA_Cadastro")=da_cadastro
RSALUNO_aux_bd.update

set RSALUNO_aux_bd=nothing
end if


if isnull(vinculado) or vinculado="" then

	Set RSCONTATO_aux_bd3 = server.createobject("adodb.recordset")
	sql_atualiza3= "UPDATE TBI_Contatos SET ID_Familia = '"&id_familia&"', ID_End_Bloqueto ='"& bloq &"' WHERE CO_Matricula = "& cod &" AND TP_Contato = '"& responsavel_financeiro &"'"
	Set RSCONTATO_aux_bd3 = CONCONT_aux.Execute(sql_atualiza3)
	
	Set RSCONTATO_aux_bd4 = server.createobject("adodb.recordset")
	sql_atualiza4= "UPDATE TBI_Contatos SET ID_End_Bloqueto ='"& circ &"' WHERE CO_Matricula = "& cod &" AND TP_Contato = '"& responsavel_pedagogico &"'"
	Set RSCONTATO_aux_bd4 = CONCONT_aux.Execute(sql_atualiza4)
else


Set RSALUNO_aux_bd2 = server.createobject("adodb.recordset")
sql_atualiza_al_vinc= "UPDATE TBI_Alunos SET NO_Pai ='"& pai &"', NO_Mae ='"& mae &"', IN_Pai_Falecido ="& pai_falecido &", IN_Mae_Falecida ="& mae_falecido &", "
sql_atualiza_al_vinc=sql_atualiza_al_vinc&"CO_Estado_Civil ='"& sit_pais &"' WHERE CO_Matricula = "& vinculado
Set RSALUNO_aux_bd2 = CON1_aux.Execute(sql_atualiza_al_vinc)


	Set RSCONTATO_aux_bd3 = server.createobject("adodb.recordset")
	sql_atualiza3= "UPDATE TBI_Contatos SET ID_Familia = '"&id_familia&"', ID_End_Bloqueto ='"& bloq &"' WHERE CO_Matricula = "& vinculado &" AND TP_Contato = '"& responsavel_financeiro &"'"
	Set RSCONTATO_aux_bd3 = CONCONT_aux.Execute(sql_atualiza3)
	
	Set RSCONTATO_aux_bd4 = server.createobject("adodb.recordset")
	sql_atualiza4= "UPDATE TBI_Contatos SET ID_End_Bloqueto ='"& circ &"' WHERE CO_Matricula = "& vinculado &" AND TP_Contato = '"& responsavel_pedagogico &"'"
	Set RSCONTATO_aux_bd4 = CONCONT_aux.Execute(sql_atualiza4)
end if



'Grava dados definitivos============================================================================================================


' seleciona os responsáveis financeiros e pedagógicos

if isnull(vinculado) or vinculado="" then
cod_consulta=cod
else
cod_consulta=vinculado
end if

		Set RSRESPS_aux = Server.CreateObject("ADODB.Recordset")
		SQLARESPS_aux= "SELECT * FROM TBI_Alunos WHERE CO_Matricula ="&cod_consulta
		RSRESPS_aux.Open SQLARESPS_aux, CON1_aux
		
IF RSRESPS_aux.EOF THEN
response.redirect("altera.asp?opt=err1&cod_cons="&cod&"&v="&vinculado)
else	
	
tp_resp_fin=RSRESPS_aux("TP_Resp_Fin")
tp_resp_ped=RSRESPS_aux("TP_Resp_Ped")

if isnull(tp_resp_fin) or tp_resp_fin="" then
response.redirect("altera.asp?opt=err2&cod_cons="&cod&"&v="&vinculado)
end if
if isnull(tp_resp_ped) or tp_resp_ped="" then
response.redirect("altera.asp?opt=err3&cod_cons="&cod&"&v="&vinculado)
end if
end if

		Set RSCONTATO_aux = Server.CreateObject("ADODB.Recordset")
		SQLAA_aux= "SELECT * FROM TBI_Contatos WHERE CO_Matricula ="&cod
		RSCONTATO_aux.Open SQLAA_aux, CONCONT_aux
		
While not RSCONTATO_aux.EOF
cod_familiar=RSCONTATO_aux("TP_Contato")
nome_familiar_aux=RSCONTATO_aux("NO_Contato")
nasce_familiar_aux=RSCONTATO_aux("DA_Nascimento_Contato")
cpf_familiar_aux=RSCONTATO_aux("CO_CPF_PFisica")
rg_familiar_aux=RSCONTATO_aux("CO_RG_PFisica")
emitido_familiar_aux=RSCONTATO_aux("CO_OERG_PFisica")
emissao_familiar_aux=RSCONTATO_aux("CO_DERG_PFisica")
email_familiar_aux=RSCONTATO_aux("TX_EMail")
ocupacao_familiar_aux=RSCONTATO_aux("CO_Ocupacao")
empresa_familiar_aux=RSCONTATO_aux("NO_Empresa")
tel_familiar_aux=RSCONTATO_aux("NU_Telefones")
id_res_familiar_aux=RSCONTATO_aux("ID_Res_Aluno")
id_familia_aux=RSCONTATO_aux("ID_Familia")
id_end_bloq_aux=RSCONTATO_aux("ID_End_Bloqueto")
rua_res_familiar_aux=RSCONTATO_aux("NO_Logradouro_Res")
num_res_familiar_aux=RSCONTATO_aux("NU_Logradouro_Res")
comp_res_familiar_aux=RSCONTATO_aux("TX_Complemento_Logradouro_Res")
bairro_res_familiar_aux=RSCONTATO_aux("CO_Bairro_Res")
cid_res_familiar_aux=RSCONTATO_aux("CO_Municipio_Res")
uf_res_familiar_aux=RSCONTATO_aux("SG_UF_Res")
cep_res_familiar_aux=RSCONTATO_aux("CO_CEP_Res")
tel_res_familiar_aux=RSCONTATO_aux("NU_Telefones_Res")
rua_com_familiar_aux=RSCONTATO_aux("NO_Logradouro_Com")
num_com_familiar_aux=RSCONTATO_aux("NU_Logradouro_Com")
comp_com_familiar_aux=RSCONTATO_aux("TX_Complemento_Logradouro_Com")
bairro_com_familiar_aux=RSCONTATO_aux("CO_Bairro_Com")
cid_com_familiar_aux=RSCONTATO_aux("CO_Municipio_Com")
uf_com_familiar_aux=RSCONTATO_aux("SG_UF_Com")
cep_com_familiar_aux=RSCONTATO_aux("CO_CEP_Com")
tel_com_familiar_aux=RSCONTATO_aux("NU_Telefones_Com")
co_vinc_familiar_aux=RSCONTATO_aux("CO_Matricula_Vinc")
tp_vinc_familiar_aux=RSCONTATO_aux("TP_Contato_Vinc")

familiares_ficaram=familiares_ficaram&"#$#"&cod_familiar

	'response.Write("<br> - "&cod_familiar&" - "&emissao_familiar_aux	)
	'response.end()

'response.Write("n="&nome_familiar_aux&"t="&tel_familiar_aux&"e"&rua_res_familiar_aux)
'response.end()
	if isnull(vinculado) or vinculado="" then
		if cod_familiar=tp_resp_fin then
			'if isnull(nome_familiar_aux) or nome_familiar_aux="" or isnull(cpf_familiar_aux) or cpf_familiar_aux="" or isnull(rg_familiar_aux) or rg_familiar_aux="" or isnull(emitido_familiar_aux) or emitido_familiar_aux="" or isnull(email_familiar_aux) or email_familiar_aux="" or isnull(empresa_familiar_aux) or empresa_familiar_aux="" or isnull(tel_familiar_aux) or tel_familiar_aux="" or isnull(rua_res_familiar_aux) or rua_res_familiar_aux="" or isnull(comp_res_familiar_aux) or comp_res_familiar_aux="" or isnull(uf_res_familiar_aux) or uf_res_familiar_aux="" or isnull(cep_res_familiar_aux) or cep_res_familiar_aux="" or isnull(tel_res_familiar_aux) or tel_res_familiar_aux="" or isnull(rua_com_familiar_aux) or rua_com_familiar_aux="" or isnull(comp_com_familiar_aux) or comp_com_familiar_aux="" or isnull(uf_com_familiar_aux) or uf_com_familiar_aux="" or isnull(cep_com_familiar_aux) or cep_com_familiar_aux="" or isnull(tel_com_familiar_aux) or tel_com_familiar_aux="" then
			if isnull(nome_familiar_aux) or nome_familiar_aux="" or isnull(tel_familiar_aux) or tel_familiar_aux="" or isnull(rua_res_familiar_aux) or rua_res_familiar_aux=""then
			response.redirect("altera.asp?opt=err4&cod_cons="&cod&"&v="&vinculado)
			end if
		elseif cod_familiar=tp_resp_ped then
			'if isnull(nome_familiar_aux) or nome_familiar_aux="" or isnull(cpf_familiar_aux) or cpf_familiar_aux="" or isnull(rg_familiar_aux) or rg_familiar_aux="" or isnull(emitido_familiar_aux) or emitido_familiar_aux="" or isnull(email_familiar_aux) or email_familiar_aux="" or isnull(empresa_familiar_aux) or empresa_familiar_aux="" or isnull(tel_familiar_aux) or tel_familiar_aux="" or isnull(rua_res_familiar_aux) or rua_res_familiar_aux="" or isnull(comp_res_familiar_aux) or comp_res_familiar_aux="" or isnull(uf_res_familiar_aux) or uf_res_familiar_aux="" or isnull(cep_res_familiar_aux) or cep_res_familiar_aux="" or isnull(tel_res_familiar_aux) or tel_res_familiar_aux="" or isnull(rua_com_familiar_aux) or rua_com_familiar_aux="" or isnull(comp_com_familiar_aux) or comp_com_familiar_aux="" or isnull(uf_com_familiar_aux) or uf_com_familiar_aux="" or isnull(cep_com_familiar_aux) or cep_com_familiar_aux="" or isnull(tel_com_familiar_aux) or tel_com_familiar_aux="" then
			if isnull(nome_familiar_aux) or nome_familiar_aux="" or isnull(tel_familiar_aux) or tel_familiar_aux="" or isnull(rua_res_familiar_aux) or rua_res_familiar_aux=""then
			response.redirect("altera.asp?opt=err5&cod_cons="&cod&"&v="&vinculado)
			end if
		end if
	end if


		Set RSCONTATO_aux_delete = Server.CreateObject("ADODB.Recordset")
		SQLAC_delete= "DELETE * FROM TB_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod
		RSCONTATO_aux_delete.Open SQLAC_delete, CONCONT

	
		Set RSCONTATO = Server.CreateObject("ADODB.Recordset")
		SQLAA= "SELECT * FROM TB_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod
		RSCONTATO.Open SQLAA, CONCONT

if RSCONTATO.EOF then
'	if isnull(vinculado) or vinculado="" then
'		if cod_familiar=tp_resp_fin then
'		response.redirect("altera.asp?opt=err4&cod="&cod&"&v="&vinculado)
'		elseif cod_familiar=tp_resp_ped then
'		response.redirect("altera.asp?opt=err5&cod="&cod&"&v="&vinculado)
'		end if	
'	end if
if isnull(nasce_familiar_aux) or nasce_familiar_aux="" then
nasce_familiar_aux =NULL
end if

if isnull(emissao_familiar_aux) or emissao_familiar_aux="" then
emissao_familiar_aux=NULL
end if

if isnull(ocupacao_familiar_aux) or ocupacao_familiar_aux="" or ocupacao_familiar_aux=0  then
ocupacao_familiar_aux =NULL
end if

if isnull(num_res_familiar_aux) or num_res_familiar_aux="" then
num_res_familiar_aux =NULL
end if

if isnull(bairro_res_familiar_aux) or bairro_res_familiar_aux="" or bairro_res_familiar_aux=0 then
bairro_res_familiar_aux =NULL
end if

if isnull(cid_res_familiar_aux) or cid_res_familiar_aux="" or cid_res_familiar_aux=0 then
cid_res_familiar_aux =NULL
end if

if isnull(num_com_familiar_aux) or num_com_familiar_aux="" then
num_com_familiar_aux =NULL
end if

if isnull(cid_com_familiar_aux) or cid_com_familiar_aux="" or cid_com_familiar_aux=0  then
cid_com_familiar_aux=NULL
end if

if isnull(bairro_com_familiar_aux) or bairro_com_familiar_aux="" or bairro_com_familiar_aux=0 then
bairro_com_familiar_aux =NULL
end if

if isnull(co_vinc_familiar_aux) or co_vinc_familiar_aux="" then
co_vinc_familiar_aux =NULL
end if



Set RSCONTATO_bd = server.createobject("adodb.recordset")
RSCONTATO_bd.open "TB_Contatos", CONCONT, 2, 2 'which table do you want open
RSCONTATO_bd.addnew
RSCONTATO_bd("CO_Matricula")=cod
RSCONTATO_bd("TP_Contato")=cod_familiar
RSCONTATO_bd("NO_Contato")=nome_familiar_aux
RSCONTATO_bd("DA_Nascimento_Contato")=nasce_familiar_aux
RSCONTATO_bd("CO_CPF_PFisica")=cpf_familiar_aux
RSCONTATO_bd("CO_RG_PFisica")=rg_familiar_aux
RSCONTATO_bd("CO_OERG_PFisica")=emitido_familiar_aux
RSCONTATO_bd("CO_DERG_PFisica")=emissao_familiar_aux
RSCONTATO_bd("TX_EMail")=email_familiar_aux
RSCONTATO_bd("CO_Ocupacao")=ocupacao_familiar_aux
RSCONTATO_bd("NO_Empresa")=empresa_familiar_aux
RSCONTATO_bd("NU_Telefones")=tel_familiar_aux
RSCONTATO_bd("ID_Res_Aluno")=id_res_familiar_aux
RSCONTATO_bd("ID_Familia")=id_familia_aux
RSCONTATO_bd("ID_End_Bloqueto")=id_end_bloq_aux
RSCONTATO_bd("NO_Logradouro_Res")=rua_res_familiar_aux
RSCONTATO_bd("NU_Logradouro_Res")=num_res_familiar_aux
RSCONTATO_bd("TX_Complemento_Logradouro_Res")=comp_res_familiar_aux
RSCONTATO_bd("CO_Bairro_Res")=bairro_res_familiar_aux
RSCONTATO_bd("CO_Municipio_Res")=cid_res_familiar_aux
RSCONTATO_bd("SG_UF_Res")=uf_res_familiar_aux
RSCONTATO_bd("CO_CEP_Res")=cep_res_familiar_aux
RSCONTATO_bd("NU_Telefones_Res")=tel_res_familiar_aux
RSCONTATO_bd("NO_Logradouro_Com")=rua_com_familiar_aux
RSCONTATO_bd("NU_Logradouro_Com")=num_com_familiar_aux
RSCONTATO_bd("TX_Complemento_Logradouro_Com")=comp_com_familiar_aux
RSCONTATO_bd("CO_Bairro_Com")=bairro_com_familiar_aux
RSCONTATO_bd("CO_Municipio_Com")=cid_com_familiar_aux
RSCONTATO_bd("SG_UF_Com")=uf_com_familiar_aux
RSCONTATO_bd("CO_CEP_Com")=cep_com_familiar_aux
RSCONTATO_bd("NU_Telefones_Com")=tel_com_familiar_aux
RSCONTATO_bd("CO_Matricula_Vinc")=co_vinc_familiar_aux
RSCONTATO_bd("TP_Contato_Vinc")=tp_vinc_familiar_aux

  RSCONTATO_bd.update
  
set RSCONTATO_bd=nothing

'inclui familiares em alunos vinculados
'==============================================================================================================================
		Set RSCONTATO_aux_vinc = Server.CreateObject("ADODB.Recordset")
		SQLAA_aux_vinc= "SELECT * FROM TB_Contatos WHERE TP_Contato='ALUNO' and CO_Matricula_Vinc ="&cod
		RSCONTATO_aux_vinc.Open SQLAA_aux_vinc, CONCONT

'response.Write("2 "&SQLAA_aux_vinc)
'response.End()

if RSCONTATO_aux_vinc.eof then

else
		while not RSCONTATO_aux_vinc.eof	
		cod_vinculado=RSCONTATO_aux_vinc("CO_Matricula")

		if cod_familiar="ALUNO" then
		RSCONTATO_aux_vinc.movenext
		else
			Set RSCONTATO_aux_delete = Server.CreateObject("ADODB.Recordset")
			SQLAC_delete= "DELETE * FROM TB_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod_vinculado
			RSCONTATO_aux_delete.Open SQLAC_delete, CONCONT
				
				Set RSCONTATO_bd = server.createobject("adodb.recordset")
				RSCONTATO_bd.open "TB_Contatos", CONCONT, 2, 2 'which table do you want open
				RSCONTATO_bd.addnew
				RSCONTATO_bd("CO_Matricula")=cod_vinculado
				RSCONTATO_bd("TP_Contato")=cod_familiar
				RSCONTATO_bd("CO_Matricula_Vinc")=cod
				RSCONTATO_bd("TP_Contato_Vinc")=cod_familiar
				
				  RSCONTATO_bd.update
				  
				set RSCONTATO_bd=nothing			

		
		RSCONTATO_aux_vinc.movenext
		end if
		wend
End if
'==============================================================================================================================
	if isnull(co_vinc_familiar_aux) or co_vinc_familiar_aux="" then
	else
	
			Set RSCONTATO_vinc_familiar_aux = Server.CreateObject("ADODB.Recordset")
			SQLAA_vinc_familiar_aux= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_vinc_familiar_aux&"' and CO_Matricula ="&co_vinc_familiar_aux
			RSCONTATO_vinc_familiar_aux.Open SQLAA_vinc_familiar_aux, CONCONT_aux
	'response.Write "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_vinc_familiar_aux&"' and CO_Matricula ="&co_vinc_familiar_aux

	cod_familiar_vinc=RSCONTATO_vinc_familiar_aux("TP_Contato")
	nome_familiar_vinc=RSCONTATO_vinc_familiar_aux("NO_Contato")
	nasce_familiar_vinc=RSCONTATO_vinc_familiar_aux("DA_Nascimento_Contato")
	cpf_familiar_vinc=RSCONTATO_vinc_familiar_aux("CO_CPF_PFisica")
	rg_familiar_vinc=RSCONTATO_vinc_familiar_aux("CO_RG_PFisica")
	emitido_familiar_vinc=RSCONTATO_vinc_familiar_aux("CO_OERG_PFisica")
	emissao_familiar_vinc=RSCONTATO_vinc_familiar_aux("CO_DERG_PFisica")
	email_familiar_vinc=RSCONTATO_vinc_familiar_aux("TX_EMail")
	ocupacao_familiar_vinc=RSCONTATO_vinc_familiar_aux("CO_Ocupacao")
	empresa_familiar_vinc=RSCONTATO_vinc_familiar_aux("NO_Empresa")
	tel_familiar_vinc=RSCONTATO_vinc_familiar_aux("NU_Telefones")
	id_res_familiar_vinc=RSCONTATO_vinc_familiar_aux("ID_Res_Aluno")
	id_familia_aux=RSCONTATO_vinc_familiar_aux("ID_Familia")
	id_end_bloq_aux=RSCONTATO_vinc_familiar_aux("ID_End_Bloqueto")
	rua_res_familiar_vinc=RSCONTATO_vinc_familiar_aux("NO_Logradouro_Res")
	num_res_familiar_vinc=RSCONTATO_vinc_familiar_aux("NU_Logradouro_Res")
	comp_res_familiar_vinc=RSCONTATO_vinc_familiar_aux("TX_Complemento_Logradouro_Res")
	bairro_res_familiar_vinc=RSCONTATO_vinc_familiar_aux("CO_Bairro_Res")
	cid_res_familiar_vinc=RSCONTATO_vinc_familiar_aux("CO_Municipio_Res")
	uf_res_familiar_vinc=RSCONTATO_vinc_familiar_aux("SG_UF_Res")
	cep_res_familiar_vinc=RSCONTATO_vinc_familiar_aux("CO_CEP_Res")
	tel_res_familiar_vinc=RSCONTATO_vinc_familiar_aux("NU_Telefones_Res")
	rua_com_familiar_vinc=RSCONTATO_vinc_familiar_aux("NO_Logradouro_Com")
	num_com_familiar_vinc=RSCONTATO_vinc_familiar_aux("NU_Logradouro_Com")
	comp_com_familiar_vinc=RSCONTATO_vinc_familiar_aux("TX_Complemento_Logradouro_Com")
	bairro_com_familiar_vinc=RSCONTATO_vinc_familiar_aux("CO_Bairro_Com")
	cid_com_familiar_vinc=RSCONTATO_vinc_familiar_aux("CO_Municipio_Com")
	uf_com_familiar_vinc=RSCONTATO_vinc_familiar_aux("SG_UF_Com")
	cep_com_familiar_vinc=RSCONTATO_vinc_familiar_aux("CO_CEP_Com")
	tel_com_familiar_vinc=RSCONTATO_vinc_familiar_aux("NU_Telefones_Com")
	co_vinc_familiar_vinc=RSCONTATO_vinc_familiar_aux("CO_Matricula_Vinc")
	tp_vinc_familiar_vinc=RSCONTATO_vinc_familiar_aux("TP_Contato_Vinc")		
	
	if isnull(nasce_familiar_vinc) or nasce_familiar_vinc="" then
		'if cod_familiar=tp_resp_fin then
		'response.redirect("altera.asp?opt=err4&cod="&cod&"&v="&vinculado)
		'elseif cod_familiar=tp_resp_ped then
		'response.redirect("altera.asp?opt=err5&cod="&cod&"&v="&vinculado)
		'end if
	sql_nasce="DA_Nascimento_Contato =NULL"
	else
	sql_nasce="DA_Nascimento_Contato =#"& nasce &"#"
	end if
	
	if isnull(emissao_familiar_vinc) or emissao_familiar_vinc="" then
		'if cod_familiar=tp_resp_fin then
		'response.redirect("altera.asp?opt=err4&cod="&cod&"&v="&vinculado)
		'elseif cod_familiar=tp_resp_ped then
		'response.redirect("altera.asp?opt=err5&cod="&cod&"&v="&vinculado)
		'end if
	sql_emissao="CO_DERG_PFisica =NULL"
	else
	sql_emissao="CO_DERG_PFisica =#"& emissao_familiar_vinc &"#"
	end if
	
	if isnull(ocupacao_familiar_vinc) or ocupacao_familiar_vinc="" then
		'if cod_familiar=tp_resp_fin then
		'response.redirect("altera.asp?opt=err4&cod="&cod&"&v="&vinculado)
		'elseif cod_familiar=tp_resp_ped then
		'response.redirect("altera.asp?opt=err5&cod="&cod&"&v="&vinculado)
		'end if
	sql_ocupacao="CO_Ocupacao =NULL"
	else
	sql_ocupacao="CO_Ocupacao ="& ocupacao_familiar_vinc &""
	end if
	
	if isnull(num_res_familiar_vinc) or num_res_familiar_vinc="" then
	sql_num_res="NU_Logradouro_Res =NULL"
	else
	sql_num_res="NU_Logradouro_Res ="& num_res_familiar_vinc &""
	end if
	
	if isnull(bairro_res_familiar_vinc) or bairro_res_familiar_vinc="" then
	sql_bairro_res=" CO_Bairro_Res =NULL"
	else
	sql_bairro_res=" CO_Bairro_Res ="& bairro_res_familiar_vinc &""
	end if
	
	if isnull(cid_res_familiar_vinc) or cid_res_familiar_vinc="" then
	sql_cid_res=" CO_Municipio_Res =NULL"
	else
	sql_cid_res=" CO_Municipio_Res ="& cid_res_familiar_vinc &""
	end if
	
	if isnull(num_com_familiar_vinc) or num_com_familiar_vinc="" then
	sql_num_com="NU_Logradouro_Com =NULL"
	else
	sql_num_com="NU_Logradouro_Com ="& num_com_familiar_vinc &""
	end if
	
	if isnull(cid_com_familiar_vinc) or cid_com_familiar_vinc="" then
	sql_cid_com=" CO_Municipio_Com =NULL"
	else
	sql_cid_com=" CO_Municipio_Com ="& cid_com_familiar_vinc &""
	end if
	
	if isnull(bairro_com_familiar_vinc) or bairro_com_familiar_vinc="" then
	sql_bairro_com=" CO_Bairro_Com =NULL"
	else
	sql_bairro_com=" CO_Bairro_Com ="& bairro_com_familiar_vinc &""
	end if
	
	if isnull(co_vinc_familiar_vinc) or co_vinc_familiar_vinc="" then
	sql_vinc="CO_Matricula_Vinc =NULL"
	else
	sql_vinc="CO_Matricula_Vinc ="& co_vinc_familiar_vinc &""
	end if
	
	Set RSCONTATO_bd2 = server.createobject("adodb.recordset")
	sql_atualiza= "UPDATE TB_Contatos SET NO_Contato = '"&nome_familiar_vinc&"', "& sql_nasce &", CO_CPF_PFisica ='"& cpf_familiar_vinc &"', CO_RG_PFisica ='"& rg_familiar_vinc &"', CO_OERG_PFisica ='"& emitido_familiar_vinc&"', "& sql_emissao &", TX_EMail ='"& email_familiar_vinc &"', "&sql_ocupacao&", NO_Empresa ='"& empresa_familiar_vinc &"', "
	sql_atualiza=sql_atualiza&"NU_Telefones ='"& tel_familiar_vinc&"', ID_Res_Aluno = "&id_res_familiar_vinc&", ID_Familia='"& id_familia_aux &"', ID_End_Bloqueto='"& id_end_bloq_aux &"', NO_Logradouro_Res ='"& rua_res_familiar_vinc &"', "& sql_num_res&", TX_Complemento_Logradouro_Res = '"&comp_res_familiar_vinc&"', "& sql_bairro_res &", "& sql_cid_res &", SG_UF_Res ='"& uf_res_familiar_vinc &"', "
	sql_atualiza=sql_atualiza&"CO_CEP_Res ='"& cep_res_familiar_vinc &"', NU_Telefones_Res ='"& tel_res_familiar_vinc&"', NO_Logradouro_Com = '"&rua_com_familiar_vinc&"', "& sql_num_com&", TX_Complemento_Logradouro_Com= '"&comp_com_familiar_vinc&"',"&sql_bairro_com&", "& sql_cid_com &", SG_UF_Com ='"& uf_com_familiar_vinc &"', "
	sql_atualiza=sql_atualiza&"CO_CEP_Com='"& cep_com_familiar_vinc &"', NU_Telefones_Com ='"& tel_com_familiar_vinc&"', "&sql_vinc&", TP_Contato_Vinc ='"& tp_vinc_familiar_vinc&"' WHERE CO_Matricula = "& co_vinc_familiar_aux &" AND TP_Contato = '"& tp_vinc_familiar_aux &"'"
	
	'response.Write("<BR>TADA"&sql_atualiza&"<BR><BR>")

	Set RSCONTATO_bd2 = CONCONT.Execute(sql_atualiza)
	'if isnull(co_vinc_familiar_aux) or co_vinc_familiar_aux="" then
	end if
'if do RSCONTATO.EOF
END IF
RSCONTATO_aux.MOVENEXT
WEND
	'response.End()
		Set RS_aux = Server.CreateObject("ADODB.Recordset")
		SQL_aux = "SELECT * FROM TBI_Alunos WHERE CO_Matricula ="& cod
		RS_aux.Open SQL_aux, CON1_aux
		
codigo_aux = RS_aux("CO_Matricula")
nome_aluno_aux = RS_aux("NO_Aluno")
apelido_aux= RS_aux("NO_Apelido")
desteridade_aux= RS_aux("IN_Desteridade")
nacionalidade_aux= RS_aux("CO_Nacionalidade")
sexo_aux=RS_aux("IN_Sexo")
pai_aux= RS_aux("NO_Pai")
mae_aux= RS_aux("NO_Mae")
pai_fal_aux= RS_aux("IN_Pai_Falecido")
mae_fal_aux= RS_aux("IN_Mae_Falecida")
uf_nat_aux = RS_aux("SG_UF_Natural")
cid_nat_aux = RS_aux("CO_Municipio_Natural")
resp_fin_aux= RS_aux("TP_Resp_Fin")
resp_ped_aux= RS_aux("TP_Resp_Ped")
pais_aux= RS_aux("CO_Pais_Natural")
msn_aux= RS_aux("TX_MSN")
orkut_aux= RS_aux("TX_ORKUT")
religiao_aux= RS_aux("CO_Religiao")
cor_raca_aux= RS_aux("CO_Raca")
sit_pais_aux= RS_aux("CO_Estado_Civil")
col_or_aux= RS_aux("NO_Colegio_Origem")
et_curs_aux= RS_aux("NO_Serie_Cursada")
uf_curs_aux= RS_aux("SG_UF_Cursada")
cid_curs_aux= RS_aux("CO_Municipio_Cursada")
da_cadastro_aux= RS_aux("DA_Cadastro")
da_entrada_aux= RS_aux("DA_Entrada_Escola")


		Set RSCONTATO_aux_delete = Server.CreateObject("ADODB.Recordset")
		SQLAA_delete= "DELETE * FROM TB_Alunos WHERE CO_Matricula ="&cod
		RSCONTATO_aux_delete.Open SQLAA_delete, CON1

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod
		RS.Open SQL, CON1
		
if RS.EOF Then

Set RSALUNO_bd = server.createobject("adodb.recordset")
RSALUNO_bd.open "TB_Alunos", CON1, 2, 2
RSALUNO_bd.addnew
RSALUNO_bd("CO_Matricula")=codigo_aux
RSALUNO_bd("RA_Aluno")=codigo_aux
RSALUNO_bd("NO_Aluno")=nome_aluno_aux
RSALUNO_bd("NO_Apelido")=apelido_aux
RSALUNO_bd("IN_Sexo")=sexo_aux
RSALUNO_bd("IN_Desteridade")=desteridade_aux
RSALUNO_bd("CO_Nacionalidade")=nacionalidade_aux
RSALUNO_bd("CO_Pais_Natural")=pais_aux
RSALUNO_bd("SG_UF_Natural")=uf_nat_aux
RSALUNO_bd("CO_Municipio_Natural")=cid_nat_aux
RSALUNO_bd("TX_MSN")=msn_aux
RSALUNO_bd("TX_ORKUT")=orkut_aux
RSALUNO_bd("CO_Raca")=cor_raca_aux
RSALUNO_bd("CO_Religiao")=religiao_aux
RSALUNO_bd("NO_Pai")=pai_aux
RSALUNO_bd("NO_Mae")=mae_aux
RSALUNO_bd("IN_Pai_Falecido")=pai_fal_aux
RSALUNO_bd("IN_Mae_Falecida")=mae_fal_aux
RSALUNO_bd("CO_Estado_Civil")=sit_pais_aux
RSALUNO_bd("TP_Resp_Fin")=resp_fin_aux							  
RSALUNO_bd("TP_Resp_Ped")=resp_ped_aux
RSALUNO_bd("NO_Colegio_Origem")= col_or_aux
RSALUNO_bd("NO_Serie_Cursada")=et_curs_aux
RSALUNO_bd("SG_UF_Cursada")=uf_curs_aux
RSALUNO_bd("CO_Municipio_Cursada")=cid_curs_aux
RSALUNO_bd("DA_Entrada_Escola")=da_entrada_aux
RSALUNO_bd("DA_Cadastro")=da_cadastro_aux
RSALUNO_bd.update
  
set RSALUNO_bd=nothing

if isnull(vinculado) or vinculado="" then
	Set RSCONTATO_aux_bd3 = server.createobject("adodb.recordset")
	sql_atualiza3= "UPDATE TB_Contatos SET ID_Familia = '"&id_familia&"', ID_End_Bloqueto ='"& bloq &"' WHERE CO_Matricula = "& cod &" AND TP_Contato = '"& responsavel_financeiro &"'"
	Set RSCONTATO_aux_bd3 = CONCONT.Execute(sql_atualiza3)
	
	Set RSCONTATO_aux_bd4 = server.createobject("adodb.recordset")
	sql_atualiza4= "UPDATE TB_Contatos SET ID_End_Bloqueto ='"& circ &"' WHERE CO_Matricula = "& cod &" AND TP_Contato = '"& responsavel_pedagogico &"'"
	Set RSCONTATO_aux_bd4 = CONCONT.Execute(sql_atualiza4)
else
		Set RS_aux = Server.CreateObject("ADODB.Recordset")
		SQL_aux = "SELECT * FROM TBI_Alunos WHERE CO_Matricula ="& vinculado
		RS_aux.Open SQL_aux, CON1_aux
		
pai_aux_vinc= RS_aux("NO_Pai")
mae_aux_vinc= RS_aux("NO_Mae")
pai_fal_aux_vinc= RS_aux("IN_Pai_Falecido")
mae_fal_aux_vinc= RS_aux("IN_Mae_Falecida")
sit_pais_aux_vinc= RS_aux("CO_Estado_Civil")

Set RSALUNO_aux_bd2 = server.createobject("adodb.recordset")
sql_atualiza_al_vinc= "UPDATE TB_Alunos SET NO_Pai ='"& pai_aux_vinc &"', NO_Mae ='"& mae_aux_vinc &"', IN_Pai_Falecido ="& pai_fal_aux_vinc &", IN_Mae_Falecida ="& mae_fal_aux_vinc &", "
sql_atualiza_al_vinc=sql_atualiza_al_vinc&"CO_Estado_Civil ='"& sit_pais_aux_vinc &"' WHERE CO_Matricula = "& vinculado
Set RSALUNO_aux_bd2 = CON1.Execute(sql_atualiza_al_vinc)

'response.Write("<BR>2>>"&sql_atualiza_al_vinc)
'response.end()

	Set RSCONTATO_aux_bd3 = server.createobject("adodb.recordset")
	sql_atualiza3= "UPDATE TB_Contatos SET ID_Familia = '"&id_familia&"', ID_End_Bloqueto ='"& bloq &"' WHERE CO_Matricula = "& vinculado &" AND TP_Contato = '"& responsavel_financeiro &"'"
	Set RSCONTATO_aux_bd3 = CONCONT.Execute(sql_atualiza3)
	
	Set RSCONTATO_aux_bd4 = server.createobject("adodb.recordset")
	sql_atualiza4= "UPDATE TB_Contatos SET ID_End_Bloqueto ='"& circ &"' WHERE CO_Matricula = "& vinculado &" AND TP_Contato = '"& responsavel_pedagogico &"'"
	Set RSCONTATO_aux_bd4 = CONCONT.Execute(sql_atualiza4)
end if

END IF	

'Apaga dados que sobraram das tabelas definitivas==========================================================================

familiares_verifica_ficaram=split(familiares_ficaram,"#$#")
familiar_presente=familiares_verifica_ficaram(i)

		Set RSCONTATO_verifica = Server.CreateObject("ADODB.Recordset")
		SQLAv= "SELECT * FROM TB_Contatos WHERE CO_Matricula ="&cod
		RSCONTATO_verifica.Open SQLAv, CONCONT

while not RSCONTATO_verifica.eof
cod_familiar_verifica=RSCONTATO_verifica("TP_Contato")
apaga="s"
for i=1 to ubound(familiares_verifica_ficaram)
familiar_presente=familiares_verifica_ficaram(i)
if cod_familiar_verifica=familiar_presente then		
apaga="n"
response.Write(familiar_presente&" - "& apaga)
end if
next
	
	
if apaga="s" then
	Set RSCONTATO_delete = Server.CreateObject("ADODB.Recordset")
	SQLAD_delete= "DELETE * FROM TB_Contatos WHERE TP_Contato='"&cod_familiar_verifica&"' AND CO_Matricula ="&cod
	RSCONTATO_delete.Open SQLAD_delete, CONCONT
end if
RSCONTATO_verifica.movenext
wend	
'if do opt
END IF
response.redirect("altera.asp?opt=ok&cod_cons="&cod)
%>