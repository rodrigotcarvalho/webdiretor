<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<%
nivel=4
opt=request.querystring("opt")

qtd_tipo_familiares=request.Form("qtd_tp_pub")
foco=request.Form("foco_pub")
cod=request.form("cod_pub")
ordem_familiares=request.Form("ord_pub")
if opt="cpf" then
cpf_cons=request.Form("cpf_pub")
'ordem_familiares=Session("ordem_familiares")
'Session("ordem_familiares")=ordem_familiares
'else
ordem_familiares=request.Form("ord_pub")
'Session("ordem_familiares")=ordem_familiares
end if
dados=Server.URLEncode(ordem_familiares)&"#sep#"&qtd_tipo_familiares&"#sep#"&foco&"#sep#"&cod
'response.Write(">>"&dados)
larg=1000

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0

		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CONCONT_aux = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT_aux = "DBQ="& CAMINHO_ct_aux & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT_aux.Open ABRIRCONT_aux
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON1_aux = Server.CreateObject("ADODB.Connection") 
		ABRIR1_aux = "DBQ="& CAMINHO_al_aux & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1_aux.Open ABRIR1_aux


if opt="zero" then
' copia os dados da tabela para a tabela auxiliar
familiares = Split(ordem_familiares, "##")

'Apaga dados preenchidos nas tabelas tempor�rias==========================================================================

		Set RSCONTATO_aux_delete = Server.CreateObject("ADODB.Recordset")
		SQLAA_delete= "DELETE * FROM TBI_Alunos WHERE CO_Matricula ="&cod
		RSCONTATO_aux_delete.Open SQLAA_delete, CON1_aux
		
		Set RSCONTATO_aux_delete = Server.CreateObject("ADODB.Recordset")
		SQLAC_delete= "DELETE * FROM TBI_Contatos WHERE CO_Matricula ="&cod
		RSCONTATO_aux_delete.Open SQLAC_delete, CONCONT_aux
			
'Grava dados da BD para tabelas tempor�rias==========================================================================
		


		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod
		RS.Open SQL, CON1
		

		if RS.EOF Then
			else
			resp_fin= RS("TP_Resp_Fin")
			resp_ped= RS("TP_Resp_Ped")
			pai_fal= RS("IN_Pai_Falecido")
			mae_fal= RS("IN_Mae_Falecida")

			Set RS_aux = Server.CreateObject("ADODB.Recordset")
			SQL_aux = "SELECT * FROM TBI_Alunos WHERE CO_Matricula ="& cod
			RS_aux.Open SQL_aux, CON1_aux

			if RS_aux.EOF Then
				Set RSALUNO_aux_bd = server.createobject("adodb.recordset")
				RSALUNO_aux_bd.open "TBI_Alunos", CON1_aux, 2, 2
				RSALUNO_aux_bd.addnew
				RSALUNO_aux_bd("CO_Matricula")=cod
				RSALUNO_aux_bd("TP_Resp_Fin")=resp_fin							  
				RSALUNO_aux_bd("TP_Resp_Ped")=resp_ped
				RSALUNO_aux_bd("IN_Pai_Falecido")=pai_fal							  
				RSALUNO_aux_bd("IN_Mae_Falecida")=mae_fal				
				RSALUNO_aux_bd.update	
				set RSALUNO_aux_bd=nothing
			else
				Set RSALUNO_aux_bd2 = server.createobject("adodb.recordset")
				sql_atualiza_al= "UPDATE TBI_Alunos SET IN_Pai_Falecido="&pai_fal&", IN_Mae_Falecida="&mae_fal&", TP_Resp_Fin ='"& resp_fin &"', TP_Resp_Ped ='"& resp_ped &"' WHERE CO_Matricula = "& cod
				Set RSALUNO_aux_bd2 = CON1_aux.Execute(sql_atualiza_al)
			end if

		end if
		
	for i=1 to ubound(familiares)
		cod_nome_familiar=familiares(i)
		cod_nome = Split(cod_nome_familiar, "!!")
		cod_familiar=cod_nome(0)
		nome_familiar=cod_nome(1)		

		Set RSCONTATO = Server.CreateObject("ADODB.Recordset")
		SQLAA= "SELECT * FROM TB_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod
		RSCONTATO.Open SQLAA, CONCONT
		
		if RSCONTATO.EOF then
		else
			nome_familiar_aux=RSCONTATO("NO_Contato")
			nasce_familiar_aux=RSCONTATO("DA_Nascimento_Contato")
			cpf_familiar_aux=RSCONTATO("CO_CPF_PFisica")
			rg_familiar_aux=RSCONTATO("CO_RG_PFisica")
			emitido_familiar_aux=RSCONTATO("CO_OERG_PFisica")
			emissao_familiar_aux=RSCONTATO("CO_DERG_PFisica")
			email_familiar_aux=RSCONTATO("TX_EMail")
			ocupacao_familiar_aux=RSCONTATO("CO_Ocupacao")
			empresa_familiar_aux=RSCONTATO("NO_Empresa")
			tel_familiar_aux=RSCONTATO("NU_Telefones")
			id_res_familiar_aux=RSCONTATO("ID_Res_Aluno")
			id_familia_aux=RSCONTATO("ID_Familia")
			id_end_bloq_aux=RSCONTATO("ID_End_Bloqueto")
			rua_res_familiar_aux=RSCONTATO("NO_Logradouro_Res")
			num_res_familiar_aux=RSCONTATO("NU_Logradouro_Res")
			comp_res_familiar_aux=RSCONTATO("TX_Complemento_Logradouro_Res")
			bairro_res_familiar_aux=RSCONTATO("CO_Bairro_Res")
			cid_res_familiar_aux=RSCONTATO("CO_Municipio_Res")
			uf_res_familiar_aux=RSCONTATO("SG_UF_Res")
			cep_res_familiar_aux=RSCONTATO("CO_CEP_Res")
			tel_res_familiar_aux=RSCONTATO("NU_Telefones_Res")
			rua_com_familiar_aux=RSCONTATO("NO_Logradouro_Com")
			num_com_familiar_aux=RSCONTATO("NU_Logradouro_Com")
			comp_com_familiar_aux=RSCONTATO("TX_Complemento_Logradouro_Com")
			bairro_com_familiar_aux=RSCONTATO("CO_Bairro_Com")
			cid_com_familiar_aux=RSCONTATO("CO_Municipio_Com")
			uf_com_familiar_aux=RSCONTATO("SG_UF_Com")
			cep_com_familiar_aux=RSCONTATO("CO_CEP_Com")
			tel_com_familiar_aux=RSCONTATO("NU_Telefones_Com")
			co_vinc_familiar_aux=RSCONTATO("CO_Matricula_Vinc")
			tp_vinc_familiar_aux=RSCONTATO("TP_Contato_Vinc")

			Set RSCONTATO_aux = Server.CreateObject("ADODB.Recordset")
			SQLAA_aux= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod
			RSCONTATO_aux.Open SQLAA_aux, CONCONT_aux

			if RSCONTATO_aux.EOF then
				Set RSCONTATO_aux_bd = server.createobject("adodb.recordset")

				RSCONTATO_aux_bd.open "TBI_Contatos", CONCONT_aux, 2, 2 'which table do you want open
				RSCONTATO_aux_bd.addnew
				RSCONTATO_aux_bd("CO_Matricula")=cod
				RSCONTATO_aux_bd("TP_Contato")=cod_familiar
				RSCONTATO_aux_bd("NO_Contato")=nome_familiar_aux
				RSCONTATO_aux_bd("DA_Nascimento_Contato")=nasce_familiar_aux
				RSCONTATO_aux_bd("CO_CPF_PFisica")=cpf_familiar_aux
				RSCONTATO_aux_bd("CO_RG_PFisica")=rg_familiar_aux
				RSCONTATO_aux_bd("CO_OERG_PFisica")=emitido_familiar_aux
				RSCONTATO_aux_bd("CO_DERG_PFisica")=emissao_familiar_aux
				RSCONTATO_aux_bd("TX_EMail")=email_familiar_aux
				RSCONTATO_aux_bd("CO_Ocupacao")=ocupacao_familiar_aux
				RSCONTATO_aux_bd("NO_Empresa")=empresa_familiar_aux
				RSCONTATO_aux_bd("NU_Telefones")=tel_familiar_aux
				RSCONTATO_aux_bd("ID_Res_Aluno")=id_res_familiar_aux
				RSCONTATO_aux_bd("ID_Familia")=id_familia_aux
				RSCONTATO_aux_bd("ID_End_Bloqueto")=id_end_bloq_aux
				RSCONTATO_aux_bd("NO_Logradouro_Res")=rua_res_familiar_aux
				RSCONTATO_aux_bd("NU_Logradouro_Res")=num_res_familiar_aux
				RSCONTATO_aux_bd("TX_Complemento_Logradouro_Res")=comp_res_familiar_aux
				RSCONTATO_aux_bd("CO_Bairro_Res")=bairro_res_familiar_aux
				RSCONTATO_aux_bd("CO_Municipio_Res")=cid_res_familiar_aux
				RSCONTATO_aux_bd("SG_UF_Res")=uf_res_familiar_aux
				RSCONTATO_aux_bd("CO_CEP_Res")=cep_res_familiar_aux
				RSCONTATO_aux_bd("NU_Telefones_Res")=tel_res_familiar_aux
				RSCONTATO_aux_bd("NO_Logradouro_Com")=rua_com_familiar_aux
				RSCONTATO_aux_bd("NU_Logradouro_Com")=num_com_familiar_aux
				RSCONTATO_aux_bd("TX_Complemento_Logradouro_Com")=comp_com_familiar_aux
				RSCONTATO_aux_bd("CO_Bairro_Com")=bairro_com_familiar_aux
				RSCONTATO_aux_bd("CO_Municipio_Com")=cid_com_familiar_aux
				RSCONTATO_aux_bd("SG_UF_Com")=uf_com_familiar_aux
				RSCONTATO_aux_bd("CO_CEP_Com")=cep_com_familiar_aux
				RSCONTATO_aux_bd("NU_Telefones_Com")=tel_com_familiar_aux
				RSCONTATO_aux_bd("CO_Matricula_Vinc")=co_vinc_familiar_aux
				RSCONTATO_aux_bd("TP_Contato_Vinc")=tp_vinc_familiar_aux
				RSCONTATO_aux_bd.update
				set RSCONTATO_aux_bd=nothing

			else

				if isnull(nasce_familiar_aux) or nasce_familiar_aux="" then
				sql_nasce="DA_Nascimento_Contato =NULL"
				else
				sql_nasce="DA_Nascimento_Contato =#"& nasce_familiar_aux &"#"
				end if

				if isnull(emissao_familiar_aux) or emissao_familiar_aux="" then
				sql_emissao="CO_DERG_PFisica =NULL"
				else
				sql_emissao="CO_DERG_PFisica =#"& emissao_familiar_aux &"#"
				end if

				if isnull(ocupacao_familiar_aux) or ocupacao_familiar_aux="" then
				sql_ocupacao="CO_Ocupacao =NULL"
				else
				sql_ocupacao="CO_Ocupacao ="& ocupacao_familiar_aux &""
				end if

				if isnull(num_res_familiar_aux) or num_res_familiar_aux="" then
				sql_num_res="NU_Logradouro_Res =NULL"
				else
				sql_num_res="NU_Logradouro_Res ="& num_res_familiar_aux &""
				end if

				if isnull(bairro_res_familiar_aux) or bairro_res_familiar_aux="" then
				sql_bairro_res=" CO_Bairro_Res =NULL"
				else
				sql_bairro_res=" CO_Bairro_Res ="& bairro_res_familiar_aux &""
				end if

				if isnull(cid_res_familiar_aux) or cid_res_familiar_aux="" then
				sql_cid_res=" CO_Municipio_Res =NULL"
				else
				sql_cid_res=" CO_Municipio_Res ="& cid_res_familiar_aux &""
				end if

				if isnull(num_com_familiar_aux) or num_com_familiar_aux="" then
				sql_num_com="NU_Logradouro_Com =NULL"
				else
				sql_num_com="NU_Logradouro_Com ="& num_com_familiar_aux &""
				end if

				if isnull(cid_com_familiar_aux) or cid_com_familiar_aux="" then
				sql_cid_com=" CO_Municipio_Com =NULL"
				else
				sql_cid_com=" CO_Municipio_Com ="& cid_com_familiar_aux &""
				end if

				if isnull(bairro_com_familiar_aux) or bairro_com_familiar_aux="" then
				sql_bairro_com=" CO_Bairro_Com =NULL"
				else
				sql_bairro_com=" CO_Bairro_Com ="& bairro_com_familiar_aux &""
				end if

				if isnull(co_vinc_familiar_aux) or co_vinc_familiar_aux="" then
				sql_vinc="CO_Matricula_Vinc =NULL"
				else
				sql_vinc="CO_Matricula_Vinc ="& co_vinc_familiar_aux &""
				end if
				
				Set RSCONTATO_aux_bd2 = server.createobject("adodb.recordset")
				sql_atualiza= "UPDATE TBI_Contatos SET NO_Contato = '"&nome_familiar_aux&"', "& sql_nasce &", CO_CPF_PFisica ='"& cpf_familiar_aux &"', CO_RG_PFisica ='"& rg_familiar_aux &"', CO_OERG_PFisica ='"& emitido_familiar_aux&"', "& sql_emissao &", TX_EMail ='"& email_familiar_aux &"', "&sql_ocupacao&", NO_Empresa ='"& empresa_familiar_aux &"', "
				sql_atualiza=sql_atualiza&"NU_Telefones ='"& tel_familiar_aux&"', ID_Res_Aluno = "&id_res_familiar_aux&", NO_Logradouro_Res ='"& rua_res_familiar_aux &"', "& sql_num_res&", TX_Complemento_Logradouro_Res = '"&comp_res_familiar_aux&"', "& sql_bairro_res &", "& sql_cid_res &", SG_UF_Res ='"& uf_res_familiar_aux &"', "
				sql_atualiza=sql_atualiza&"CO_CEP_Res ='"& cep_res_familiar_aux &"', NU_Telefones_Res ='"& tel_res_familiar_aux&"', NO_Logradouro_Com = '"&rua_com_familiar_aux&"', "& sql_num_com&", TX_Complemento_Logradouro_Com= '"&comp_com_familiar_aux&"',"&sql_bairro_com&", "& sql_cid_com &", SG_UF_Com ='"& uf_com_familiar_aux &"', "
				sql_atualiza=sql_atualiza&"CO_CEP_Com='"& cep_com_familiar_aux &"', NU_Telefones_Com ='"& tel_com_familiar_aux&"', "&sql_vinc&", TP_Contato_Vinc ='"& tp_vinc_familiar_aux&"' WHERE CO_Matricula = "& cod &" AND TP_Contato = '"& cod_familiar &"'"
				Set RSCONTATO_aux_bd2 = CONCONT_aux.Execute(sql_atualiza)
				
			end if
		end if
	next
elseif opt="cpf" then
' copia os dados da tabela para a tabela auxiliar

'response.Write(">>>"&ordem_familiares)
familiares = Split(ordem_familiares, "##")

'Apaga dados preenchidos nas tabelas tempor�rias==========================================================================

		'Set RSCONTATO_aux_delete = Server.CreateObject("ADODB.Recordset")
		'SQLAA_delete= "DELETE * FROM TBI_Contatos WHERE CO_Matricula ="&cod
		'RSCONTATO_aux_delete.Open SQLAA_delete, CONCONT_aux
		
'Grava dados da BD para tabelas tempor�rias==========================================================================
		
	for i=1 to ubound(familiares)
	cod_nome_familiar=familiares(i)
	cod_nome = Split(cod_nome_familiar, "!!")
	cod_familiar=cod_nome(0)
	nome_familiar=cod_nome(1)

		Set RSCONTATO = Server.CreateObject("ADODB.Recordset")
		SQLAA= "SELECT * FROM TB_Contatos WHERE CO_CPF_PFisica='"&cpf_cons&"'"
		RSCONTATO.Open SQLAA, CONCONT
		
		if RSCONTATO.EOF then
			recupera_valores="n"
		else
			while not RSCONTATO.EOF
				co_vinc_familiar_aux=RSCONTATO("CO_Matricula")
				tp_vinc_familiar_aux=RSCONTATO("TP_Contato")	

				if tp_vinc_familiar_aux=foco and cod_aux=cod then
					recupera_valores="n"
					RSCONTATO.movenext
				else
					recupera_valores="s"	

					if cod_familiar=foco then
						
						Set RSCONTATO_aux = Server.CreateObject("ADODB.Recordset")
						SQLAA_aux= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod
						RSCONTATO_aux.Open SQLAA_aux, CONCONT_aux
						
						if RSCONTATO_aux.EOF then
								Set RSCONTATO_aux_bd = server.createobject("adodb.recordset")
								RSCONTATO_aux_bd.open "TBI_Contatos", CONCONT_aux, 2, 2 'which table do you want open
								RSCONTATO_aux_bd.addnew
								RSCONTATO_aux_bd("CO_Matricula")=cod
								RSCONTATO_aux_bd("TP_Contato")=foco
								RSCONTATO_aux_bd("NO_Contato")=""								
								RSCONTATO_aux_bd("CO_Matricula_Vinc")=co_vinc_familiar_aux
								RSCONTATO_aux_bd("TP_Contato_Vinc")=tp_vinc_familiar_aux
								RSCONTATO_aux_bd.update
								set RSCONTATO_aux_bd=nothing

							else
								Set RSCONTATO_aux_bd2 = server.createobject("adodb.recordset")
								sql_atualiza= "UPDATE TBI_Contatos SET CO_Matricula_Vinc="&co_vinc_familiar_aux&", NO_Contato='', TP_Contato_Vinc ='"& tp_vinc_familiar_aux&"' WHERE CO_Matricula = "& cod &" AND TP_Contato = '"& cod_familiar &"'"
								Set RSCONTATO_aux_bd2 = CONCONT_aux.Execute(sql_atualiza)
							end if
						end if
					RSCONTATO.movenext
					'end if do if tp_contato_aux=foco and cod_aux=cod then
				end if
			wend
		end if
	next
'end if do if opt="zero"
end if
%>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
                    <td>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>

          <%
		  
		Set RS_aux = Server.CreateObject("ADODB.Recordset")
		SQL_aux = "SELECT * FROM TBI_Alunos WHERE CO_Matricula ="& cod
		RS_aux.Open SQL_aux, CON1_aux

resp_fin= RS_aux("TP_Resp_Fin")
resp_ped= RS_aux("TP_Resp_Ped")
pai_fal= RS_aux("IN_Pai_Falecido")
mae_fal= RS_aux("IN_Mae_Falecida")		  
		  
familiares = Split(ordem_familiares, "##")
for i=1 to ubound(familiares)
cod_nome_familiar=familiares(i)
cod_nome = Split(cod_nome_familiar, "!!")
cod_familiar=cod_nome(0)
nome_familiar=cod_nome(1)

tp_familiar_guarda=nome_familiar
if cod_familiar="AVOM" then
nome_familiar="Av�s M"
elseif cod_familiar="AVOP" then
nome_familiar="Av�s P"
else
nome_familiar=nome_familiar
end if
			
	javascript_recupera="recuperarFamiliares('"&Server.URLEncode(ordem_familiares)&"','"&qtd_tipo_familiares&"','"&Server.URLEncode(cod_familiar)&"','"&cod&"')"
	javascript_exclui="ConfirmaExcluirFamiliares('"&Server.URLEncode(ordem_familiares)&"','"&qtd_tipo_familiares&"','"&Server.URLEncode(cod_familiar)&"','"&cod&"')"
	javascript_inclui ="criaFamiliar('"&Server.URLEncode(ordem_familiares)&"','"&qtd_tipo_familiares&"',this.value,'"&cod&"')"
				  
		'response.Write("SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod)
		Set RSCONTATO = Server.CreateObject("ADODB.Recordset")
		SQLAA= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod
		RSCONTATO.Open SQLAA, CONCONT_aux

if RSCONTATO.EOF and cod_familiar<>foco then
combos_responsaveis=combos_responsaveis&"##"&cod_familiar&"!!"&nome_familiar
else

'response.Write(cod_familiar&"="&foco&" and opt="&opt&" then")

if cod_familiar="ALUNO" then

elseif cod_familiar=foco and opt<>"i" then

co_vinc_familiar_aux=RSCONTATO("CO_Matricula_Vinc")
tp_vinc_familiar_aux=RSCONTATO("TP_Contato_Vinc")
'response.Write("tada "&tp_vinc_familiar_aux&"-"&co_vinc_familiar_aux)
	if (isnull(tp_vinc_familiar_aux) or tp_vinc_familiar_aux="") and(isnull(co_vinc_familiar_aux) or co_vinc_familiar_aux="") then
	tp_familiar_exibe=tp_familiar_guarda
vinculado="n"		
	nascimento = RSCONTATO("DA_Nascimento_Contato")
		if isnull(nascimento) or nascimento="" then
		else
			vetor_nascimento = Split(nascimento,"/")  
			dia_n = vetor_nascimento(0)
			mes_n = vetor_nascimento(1)
			ano_n = vetor_nascimento(2)
	
			if dia_n<10 then 
			dia_n = "0"&dia_n
			end if
	
			if mes_n<10 then
			mes_n = "0"&mes_n
			end if
			dia_a = dia_n
			mes_a = mes_n
			ano_a = ano_n
			
			nasce = dia_n&"/"&mes_n&"/"&ano_n
		end if

	nome_contato = RSCONTATO("NO_Contato")
	rua_res = RSCONTATO("NO_Logradouro_Res")
	num_res = RSCONTATO("NU_Logradouro_Res")
	comp_res = RSCONTATO("TX_Complemento_Logradouro_Res")
	bairrores= RSCONTATO("CO_Bairro_Res")
	cidres= RSCONTATO("CO_Municipio_Res")
	estadores= RSCONTATO("SG_UF_Res")
	cep = RSCONTATO("CO_CEP_Res")
	tel_res = RSCONTATO("NU_Telefones_Res")
	tel = RSCONTATO("NU_Telefones")
	mail= RSCONTATO("TX_EMail")
	ocupacao= RSCONTATO("CO_Ocupacao")
	cpf= RSCONTATO("CO_CPF_PFisica")
	rg= RSCONTATO("CO_RG_PFisica")
	emitido= RSCONTATO("CO_OERG_PFisica")
	emissao= RSCONTATO("CO_DERG_PFisica")
	empresa= RSCONTATO("NO_Empresa")
	rua_com=RSCONTATO("NO_Logradouro_Com")
	num_com = RSCONTATO("NU_Logradouro_Com")
	comp_com = RSCONTATO("TX_Complemento_Logradouro_Com")
	bairrocom= RSCONTATO("CO_Bairro_Com")
	cidcom= RSCONTATO("CO_Municipio_Com")
	estadocom= RSCONTATO("SG_UF_Com")
	cepcom = RSCONTATO("CO_CEP_Com")
	tel_com = RSCONTATO("NU_Telefones_Com")
	mes_end = RSCONTATO("ID_Res_Aluno")

	else
	'	response.Write "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_vinc_familiar_aux&"' and CO_Matricula ="&co_vinc_familiar_aux
		vinculado="s"
		le ="n"
		while le ="n" 

			Set RSCONTATO_vinc = Server.CreateObject("ADODB.Recordset")
			SQLA_vinc= "SELECT * FROM TB_Contatos WHERE TP_Contato='"&tp_vinc_familiar_aux&"' and CO_Matricula ="&co_vinc_familiar_aux
			RSCONTATO_vinc.Open SQLA_vinc, CONCONT
			
			co_vinc_familiar_aux_mais_um=RSCONTATO_vinc("CO_Matricula_Vinc")
			tp_vinc_familiar_aux_mais_um=RSCONTATO_vinc("TP_Contato_Vinc")
	
			if (isnull(tp_vinc_familiar_aux_mais_um) or tp_vinc_familiar_aux_mais_um="") and (isnull(co_vinc_familiar_aux_mais_um) or co_vinc_familiar_aux_mais_um="") then
			le="s"
			
	
				tp_familiar_exibe=tp_familiar_guarda
				
				if RSCONTATO_vinc.EOF then
				else
					nascimento = RSCONTATO_vinc("DA_Nascimento_Contato")
					if isnull(nascimento) or nascimento="" then
					else
						vetor_nascimento = Split(nascimento,"/")  
						dia_n = vetor_nascimento(0)
						mes_n = vetor_nascimento(1)
						ano_n = vetor_nascimento(2)
						
						if dia_n<10 then 
						dia_n = "0"&dia_n
						end if
						
						if mes_n<10 then
						mes_n = "0"&mes_n
						end if
						dia_a = dia_n
						mes_a = mes_n
						ano_a = ano_n
						
						nasce = dia_n&"/"&mes_n&"/"&ano_n
					end if
				nome_contato = RSCONTATO_vinc("NO_Contato")
				rua_res = RSCONTATO_vinc("NO_Logradouro_Res")
				num_res = RSCONTATO_vinc("NU_Logradouro_Res")
				comp_res = RSCONTATO_vinc("TX_Complemento_Logradouro_Res")
				bairrores= RSCONTATO_vinc("CO_Bairro_Res")
				cidres= RSCONTATO_vinc("CO_Municipio_Res")
				estadores= RSCONTATO_vinc("SG_UF_Res")
				cep = RSCONTATO_vinc("CO_CEP_Res")
				tel_res = RSCONTATO_vinc("NU_Telefones_Res")
				tel = RSCONTATO_vinc("NU_Telefones")
				mail= RSCONTATO_vinc("TX_EMail")
				ocupacao= RSCONTATO_vinc("CO_Ocupacao")
				cpf= RSCONTATO_vinc("CO_CPF_PFisica")
				rg= RSCONTATO_vinc("CO_RG_PFisica")
				emitido= RSCONTATO_vinc("CO_OERG_PFisica")
				emissao= RSCONTATO_vinc("CO_DERG_PFisica")
				empresa= RSCONTATO_vinc("NO_Empresa")
				rua_com=RSCONTATO_vinc("NO_Logradouro_Com")
				num_com = RSCONTATO_vinc("NU_Logradouro_Com")
				comp_com = RSCONTATO_vinc("TX_Complemento_Logradouro_Com")
				bairrocom= RSCONTATO_vinc("CO_Bairro_Com")
				cidcom= RSCONTATO_vinc("CO_Municipio_Com")
				estadocom= RSCONTATO_vinc("SG_UF_Com")
				cepcom = RSCONTATO_vinc("CO_CEP_Com")
				tel_com = RSCONTATO_vinc("NU_Telefones_Com")
				mes_end = RSCONTATO_vinc("ID_Res_Aluno")
				
					nome_familiar_aux=RSCONTATO_vinc("NO_Contato")
					nasce_familiar_aux=RSCONTATO_vinc("DA_Nascimento_Contato")
					cpf_familiar_aux=RSCONTATO_vinc("CO_CPF_PFisica")
					rg_familiar_aux=RSCONTATO_vinc("CO_RG_PFisica")
					emitido_familiar_aux=RSCONTATO_vinc("CO_OERG_PFisica")
					emissao_familiar_aux=RSCONTATO_vinc("CO_DERG_PFisica")
					email_familiar_aux=RSCONTATO_vinc("TX_EMail")
					ocupacao_familiar_aux=RSCONTATO_vinc("CO_Ocupacao")
					empresa_familiar_aux=RSCONTATO_vinc("NO_Empresa")
					tel_familiar_aux=RSCONTATO_vinc("NU_Telefones")
					id_res_familiar_aux=RSCONTATO_vinc("ID_Res_Aluno")
					id_familia_aux=RSCONTATO_vinc("ID_Familia")
					id_end_bloq_aux=RSCONTATO_vinc("ID_End_Bloqueto")
					rua_res_familiar_aux=RSCONTATO_vinc("NO_Logradouro_Res")
					num_res_familiar_aux=RSCONTATO_vinc("NU_Logradouro_Res")
					comp_res_familiar_aux=RSCONTATO_vinc("TX_Complemento_Logradouro_Res")
					bairro_res_familiar_aux=RSCONTATO_vinc("CO_Bairro_Res")
					cid_res_familiar_aux=RSCONTATO_vinc("CO_Municipio_Res")
					uf_res_familiar_aux=RSCONTATO_vinc("SG_UF_Res")
					cep_res_familiar_aux=RSCONTATO_vinc("CO_CEP_Res")
					tel_res_familiar_aux=RSCONTATO_vinc("NU_Telefones_Res")
					rua_com_familiar_aux=RSCONTATO_vinc("NO_Logradouro_Com")
					num_com_familiar_aux=RSCONTATO_vinc("NU_Logradouro_Com")
					comp_com_familiar_aux=RSCONTATO_vinc("TX_Complemento_Logradouro_Com")
					bairro_com_familiar_aux=RSCONTATO_vinc("CO_Bairro_Com")
					cid_com_familiar_aux=RSCONTATO_vinc("CO_Municipio_Com")
					uf_com_familiar_aux=RSCONTATO_vinc("SG_UF_Com")
					cep_com_familiar_aux=RSCONTATO_vinc("CO_CEP_Com")
					tel_com_familiar_aux=RSCONTATO_vinc("NU_Telefones_Com")
		
				
					Set RSCONTATO_aux = Server.CreateObject("ADODB.Recordset")
					SQLAA_aux= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_vinc_familiar_aux&"' and CO_Matricula ="&co_vinc_familiar_aux
					RSCONTATO_aux.Open SQLAA_aux, CONCONT_aux
		
					if RSCONTATO_aux.EOF then
						Set RSCONTATO_aux_bd = server.createobject("adodb.recordset")
		
						RSCONTATO_aux_bd.open "TBI_Contatos", CONCONT_aux, 2, 2 'which table do you want open
						RSCONTATO_aux_bd.addnew
						RSCONTATO_aux_bd("CO_Matricula")=co_vinc_familiar_aux
						RSCONTATO_aux_bd("TP_Contato")=tp_vinc_familiar_aux
						RSCONTATO_aux_bd("NO_Contato")=nome_familiar_aux
						RSCONTATO_aux_bd("DA_Nascimento_Contato")=nasce_familiar_aux
						RSCONTATO_aux_bd("CO_CPF_PFisica")=cpf_familiar_aux
						RSCONTATO_aux_bd("CO_RG_PFisica")=rg_familiar_aux
						RSCONTATO_aux_bd("CO_OERG_PFisica")=emitido_familiar_aux
						RSCONTATO_aux_bd("CO_DERG_PFisica")=emissao_familiar_aux
						RSCONTATO_aux_bd("TX_EMail")=email_familiar_aux
						RSCONTATO_aux_bd("CO_Ocupacao")=ocupacao_familiar_aux
						RSCONTATO_aux_bd("NO_Empresa")=empresa_familiar_aux
						RSCONTATO_aux_bd("NU_Telefones")=tel_familiar_aux
						RSCONTATO_aux_bd("ID_Res_Aluno")=id_res_familiar_aux
						RSCONTATO_aux_bd("ID_Familia")=id_familia_aux
						RSCONTATO_aux_bd("ID_End_Bloqueto")=id_end_bloq_aux
						RSCONTATO_aux_bd("NO_Logradouro_Res")=rua_res_familiar_aux
						RSCONTATO_aux_bd("NU_Logradouro_Res")=num_res_familiar_aux
						RSCONTATO_aux_bd("TX_Complemento_Logradouro_Res")=comp_res_familiar_aux
						RSCONTATO_aux_bd("CO_Bairro_Res")=bairro_res_familiar_aux
						RSCONTATO_aux_bd("CO_Municipio_Res")=cid_res_familiar_aux
						RSCONTATO_aux_bd("SG_UF_Res")=uf_res_familiar_aux
						RSCONTATO_aux_bd("CO_CEP_Res")=cep_res_familiar_aux
						RSCONTATO_aux_bd("NU_Telefones_Res")=tel_res_familiar_aux
						RSCONTATO_aux_bd("NO_Logradouro_Com")=rua_com_familiar_aux
						RSCONTATO_aux_bd("NU_Logradouro_Com")=num_com_familiar_aux
						RSCONTATO_aux_bd("TX_Complemento_Logradouro_Com")=comp_com_familiar_aux
						RSCONTATO_aux_bd("CO_Bairro_Com")=bairro_com_familiar_aux
						RSCONTATO_aux_bd("CO_Municipio_Com")=cid_com_familiar_aux
						RSCONTATO_aux_bd("SG_UF_Com")=uf_com_familiar_aux
						RSCONTATO_aux_bd("CO_CEP_Com")=cep_com_familiar_aux
						RSCONTATO_aux_bd("NU_Telefones_Com")=tel_com_familiar_aux
						RSCONTATO_aux_bd.update
						set RSCONTATO_aux_bd=nothing
		
					else
		
						if isnull(nasce_familiar_aux) or nasce_familiar_aux="" then
						sql_nasce="DA_Nascimento_Contato =NULL"
						else
						sql_nasce="DA_Nascimento_Contato =#"& nasce_familiar_aux &"#"
						end if
		
						if isnull(emissao_familiar_aux) or emissao_familiar_aux="" then
						sql_emissao="CO_DERG_PFisica =NULL"
						else
						sql_emissao="CO_DERG_PFisica =#"& emissao_familiar_aux &"#"
						end if
		
						if isnull(ocupacao_familiar_aux) or ocupacao_familiar_aux="" then
						sql_ocupacao="CO_Ocupacao =NULL"
						else
						sql_ocupacao="CO_Ocupacao ="& ocupacao_familiar_aux &""
						end if
		
						if isnull(num_res_familiar_aux) or num_res_familiar_aux="" then
						sql_num_res="NU_Logradouro_Res =NULL"
						else
						sql_num_res="NU_Logradouro_Res ="& num_res_familiar_aux &""
						end if
		
						if isnull(bairro_res_familiar_aux) or bairro_res_familiar_aux="" then
						sql_bairro_res=" CO_Bairro_Res =NULL"
						else
						sql_bairro_res=" CO_Bairro_Res ="& bairro_res_familiar_aux &""
						end if
		
						if isnull(cid_res_familiar_aux) or cid_res_familiar_aux="" then
						sql_cid_res=" CO_Municipio_Res =NULL"
						else
						sql_cid_res=" CO_Municipio_Res ="& cid_res_familiar_aux &""
						end if
		
						if isnull(num_com_familiar_aux) or num_com_familiar_aux="" then
						sql_num_com="NU_Logradouro_Com =NULL"
						else
						sql_num_com="NU_Logradouro_Com ="& num_com_familiar_aux &""
						end if
		
						if isnull(cid_com_familiar_aux) or cid_com_familiar_aux="" then
						sql_cid_com=" CO_Municipio_Com =NULL"
						else
						sql_cid_com=" CO_Municipio_Com ="& cid_com_familiar_aux &""
						end if
		
						if isnull(bairro_com_familiar_aux) or bairro_com_familiar_aux="" then
						sql_bairro_com=" CO_Bairro_Com =NULL"
						else
						sql_bairro_com=" CO_Bairro_Com ="& bairro_com_familiar_aux &""
						end if
						
						Set RSCONTATO_aux_bd2 = server.createobject("adodb.recordset")
						sql_atualiza= "UPDATE TBI_Contatos SET NO_Contato = '"&nome_familiar_aux&"', "& sql_nasce &", CO_CPF_PFisica ='"& cpf_familiar_aux &"', CO_RG_PFisica ='"& rg_familiar_aux &"', CO_OERG_PFisica ='"& emitido_familiar_aux&"', "& sql_emissao &", TX_EMail ='"& email_familiar_aux &"', "&sql_ocupacao&", NO_Empresa ='"& empresa_familiar_aux &"', "
						sql_atualiza=sql_atualiza&"NU_Telefones ='"& tel_familiar_aux&"', ID_Res_Aluno = "&id_res_familiar_aux&", NO_Logradouro_Res ='"& rua_res_familiar_aux &"', "& sql_num_res&", TX_Complemento_Logradouro_Res = '"&comp_res_familiar_aux&"', "& sql_bairro_res &", "& sql_cid_res &", SG_UF_Res ='"& uf_res_familiar_aux &"', "
						sql_atualiza=sql_atualiza&"CO_CEP_Res ='"& cep_res_familiar_aux &"', NU_Telefones_Res ='"& tel_res_familiar_aux&"', NO_Logradouro_Com = '"&rua_com_familiar_aux&"', "& sql_num_com&", TX_Complemento_Logradouro_Com= '"&comp_com_familiar_aux&"',"&sql_bairro_com&", "& sql_cid_com &", SG_UF_Com ='"& uf_com_familiar_aux &"', "
						sql_atualiza=sql_atualiza&"CO_CEP_Com='"& cep_com_familiar_aux &"', NU_Telefones_Com ='"& tel_com_familiar_aux&"' WHERE CO_Matricula = "& co_vinc_familiar_aux &" AND TP_Contato = '"& tp_vinc_familiar_aux &"'"
						Set RSCONTATO_aux_bd2 = CONCONT_aux.Execute(sql_atualiza)	
					end if					
				end if
			else
				co_vinc_familiar_aux=co_vinc_familiar_aux_mais_um
				tp_vinc_familiar_aux=tp_vinc_familiar_aux_mais_um
			end if
		wend
	end if


larg=larg-70
%>
		          <td width="70"><div align="right">
              <input name="botao" class="aba_foco" type="button" id="botao" value="<%RESPONSE.Write(Server.URLEncode(nome_familiar))%>" onClick="<%response.Write(javascript_recupera)%>">
            </div></td>
<%
elseif opt="i" and cod_familiar=foco then
tp_familiar_exibe=tp_familiar_guarda
larg=larg-70
vinculado="n"
%>
		          <td width="70"><div align="right">
              <input name="botao" class="aba_foco" type="button" id="botao" value="<%RESPONSE.Write(Server.URLEncode(nome_familiar))%>" onClick="<%response.Write(javascript_recupera)%>">
            </div></td>
<%
else
co_vinc_familiar_aux=RSCONTATO("CO_Matricula_Vinc")
tp_vinc_familiar_aux=RSCONTATO("TP_Contato_Vinc")
	if (isnull(tp_vinc_familiar_aux) or tp_vinc_familiar_aux="") and(isnull(co_vinc_familiar_aux) or co_vinc_familiar_aux="") then
vinculado="n"
else
vinculado="s"
end if
larg=larg-70
%>		          <td width="55"><div align="right">
              <input name="botao" class="aba_sem_foco"  type="button" id="botao" value="<%RESPONSE.Write(Server.URLEncode(nome_familiar))%>" onClick="<%response.Write(javascript_recupera)%>">
            </div></td>
<% if (cod_familiar="PAI" and pai_fal = false) or (cod_familiar="MAE" and mae_fal = false) or vinculado="s" then%>
<%else%>
		          <td width="15"><div align="left">
              <input name="botao" class="aba_exclui" type="button" id="botao" value="X" onClick="<%response.Write(javascript_exclui)%>">			  
            </div></td>
<%
end if
end if
end if	  
Next
%> <td width="5"> </td> 
<%larg=larg-5%>                   
          <td width="<%RESPONSE.Write(larg)%>"> 
            <div align="left"> 
              <select name="tp1" class="borda" onChange="<%response.Write(javascript_inclui)%>">
                <option value="0" selected></option>
<%familiares = Split(combos_responsaveis, "##")
for i=1 to ubound(familiares)
cod_nome_familiar=familiares(i)
cod_nome = Split(cod_nome_familiar, "!!")
cod_familiar=cod_nome(0)
nome_familiar=cod_nome(1)

IF cod_familiar="PAI" or cod_familiar="MAE" or cod_familiar="ALUNO" then
else
%>
                                  <option value="<%response.Write(cod_familiar)%>"> 
                                  <%response.Write(Server.URLEncode(nome_familiar))%>
                                  </option>				
<%end if	  
Next
%>  				
              </select>
            </div></td>
        </tr>
      </table>
					</td>
                  </tr>
                  <tr> 
                    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td class="tabela_aba"><div id="conteudo">
              <table width="100%" border="0" cellspacing="0" dwcopytype="CopyTableRow">
<%if opt="e" then%>			  
                <tr> 
                  <td colspan="9">
                    <%call mensagens(nivel,405,1,dados) %>
                  </td>
                </tr>
<%elseif recupera_valores="s" then%>			  
                <tr> 
                  <td colspan="9">
                    <%call mensagens(nivel,406,1,dados) %>
                  </td>
                </tr>
<%end if%>							  
                <tr> 
                  <td colspan="9" class="tb_tit">Dados Pessoais</td>
                </tr>
                <tr> 
                  <td width="152" height="26"><font class="form_dado_texto">Nome</font></td>
                  <td width="17"><div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="186" height="26"><font class="form_corpo"> 
                    <%if foco="PAI" or foco="MAE" then%>
                    <div id="<%response.Write(foco)%>"> 
                      <%response.Write(Server.URLEncode(nome_contato))%>
                      <input name="nome_familiar" type="hidden" class="borda" id="nome_familiar" value="<%response.Write(Server.URLEncode(nome_contato))%>" size="30">
                    </div>
                    <%else%>
                    <input name="nome_familiar" type="text" class="borda" id="nome_familiar" value="<%response.Write(Server.URLEncode(nome_contato))%>" size="30" onBlur="ValidaNomeFamiliar(this.value)">
                    <%
					end if%>
                    </font></td>
                  <td width="175" height="26"><font class="form_dado_texto">Data 
                    de Nascimento</font></td>
                  <td width="18"> <div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="144" height="26"> <input name="nasce_fam" type="text" class="borda" id="nasce_fam" onKeyup="formatar(this,'##/##/####')" value="<%response.write(nasce)%>" size="12" maxlength="10" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'DA_Nascimento_Contato')"></td>
                  <td width="147" height="26"> <div align="left"><font class="form_dado_texto">Rela&ccedil;&atilde;o</font></div></td>
                  <td width="17"> <div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="149" height="26"> <div align="left"><font class="form_corpo"> 
                      <%response.write(Server.URLEncode(tp_familiar_exibe))%>
                      <input name="cod_familiar" type="hidden" class="borda" id="cod_familiar" value="<%response.Write(Server.URLEncode(foco))%>">
                      <input name="cod" type="hidden" value="<%response.Write(cod)%>">
                      <input name="ordem_familiares" type="hidden" value="<%response.Write(Server.URLEncode(ordem_familiares))%>">
                      <input name="qtd_tipo_familiares" type="hidden" value="<%response.Write(Server.URLEncode(qtd_tipo_familiares))%>">
					  <%if opt="i" then%>
                      <input type="hidden" name="tp_vinc_familiar_aux" value="">
                      <input type="hidden" name="co_vinc_familiar_aux" value="">
					   <%else%>
                      <input type="hidden" name="tp_vinc_familiar_aux" value="<%response.Write(tp_vinc_familiar_aux)%>">
                      <input type="hidden" name="co_vinc_familiar_aux" value="<%response.Write(co_vinc_familiar_aux)%>">
					  <%end if%>					  
                      </font></div></td>
                </tr>
                <tr> 
                  <td width="152" height="26"> <div align="left"><font class="form_dado_texto">Ocupa&ccedil;&atilde;o 
                      </font></div></td>
                  <td width="17"> <div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="186" height="26"><font class="form_corpo"> 
                    <select name="ocupacao_fam" class="borda" id="ocupacao_fam" onChange="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'CO_Ocupacao')">
                      <option value="0" > </option>
                      <%				
		Set RS_oc = Server.CreateObject("ADODB.Recordset")
		SQL_oc = "SELECT * FROM TB_Ocupacoes order by NO_Ocupacao"
		RS_oc.Open SQL_oc, CON0
		
while not RS_oc.EOF						
co_ocup= RS_oc("CO_Ocupacao")
no_ocup= RS_oc("NO_Ocupacao")
if co_ocup = ocupacao then
%>
                      <option value="<%=co_ocup%>" selected> 
                      <% response.Write(Server.URLEncode(no_ocup))%>
                      </option>
                      <%else%>
                      <option value="<%=co_ocup%>"> 
                      <% response.Write(Server.URLEncode(no_ocup))%>
                      </option>
                      <%end if						
RS_oc.MOVENEXT
WEND
%>
                    </select>
                    </font></td>
                  <td width="175" height="26"> <div align="left"><font class="form_dado_texto">Empresa 
                      onde trabalha </font></div></td>
                  <td width="18"> <div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="144" height="26"><font class="form_corpo"> 
                    <%if empresa="" or isnull(empresa) then%>
                  <input name="trabalho_fam" type="text" class="borda" id="trabalho_fam" size="24" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'DA_Nascimento_Contato')">
                    <%else%>
                    <input name="trabalho_fam" type="text" class="borda" id="trabalho_fam" value="<%response.write(Server.URLEncode(empresa))%>" size="24" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'NO_Empresa')">
                    <%end if%>
                    </font></td>
                  <td width="147" height="26"> <div align="left"><font class="form_dado_texto">E-mail 
                      </font></div></td>
                  <td width="17"> <div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="149" height="26"> <input name="email_fam" type="text" class="borda" id="email_fam" value="<%response.write(mail)%>" size="20" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'TX_EMail')"></td>
                </tr>
                <tr> 
                  <td width="152" height="10"> <div align="left"><font class="form_dado_texto">CPF 
                      </font></div></td>
                  <td width="17"> <div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="186" height="10"><font class="form_corpo"> 
				  <%'response.Write("ValidaCPFFamiliar(this.value,"&Server.URLEncode(ordem_familiares)&"','"&Server.URLEncode(qtd_tipo_familiares)&"','"&foco&"','"&cod&"')")%>
                    <input name="cpf_fam" type="text" class="borda" id="cpf_fam" onBlur="ValidaCPFFamiliar(this.value,ordem_familiares.value,qtd_tipo_familiares.value,cod_familiar.value,cod.value)"  onKeyup="formatar(this,'#########-##')" value="<%response.write(cpf)%>" size="15">
                    </font></td>
                  <td width="175" height="10"> <div align="left"><font class="form_dado_texto">Identidade 
                      </font></div></td>
                  <td width="18"> <div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="144" height="10"><font class="form_corpo"> 
                    <input name="id_fam" type="text" class="borda" id="id_fam" value="<%response.write(rg)%>" size="15" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'CO_RG_PFisica')">
                    </font> </td>
                  <td width="147" height="10"> <div align="left"><font class="form_dado_texto">Tipo 
                      - Data de Emiss&atilde;o </font></div></td>
                  <td width="17"> <div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="149" height="10"><font class="form_corpo"> 
                    <input name="tipo_id_fam" type="text" class="borda" id="tipo_id_fam" value="<%response.write(emitido)%>" size="10" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'CO_OERG_PFisica')">
                    - 
                    <input name="nasce2_fam" type="text" class="borda" id="nasce2_fam" onKeyUp="formatar(this, '##/##/####')" value="<%response.write(emissao)%>" size="11" maxlength="10" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'CO_DERG_PFisica')">
                    </font></td>
                </tr>
                <tr> 
                  <td width="152" height="10"><font class="form_dado_texto">Telefones 
                    de Contato</font></td>
                  <td width="17"><div align="center"><font class="form_dado_texto">:</font> 
                    </div></td>
                  <td width="186" height="10"><font class="form_corpo"> 
                    <input name="tel_fam" type="text" class="borda" id="tel" value="<%response.write(tel)%>" size="42" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'NU_Telefones')">
                    </font> </td>
                  <td width="175" height="10">
				  				  <%if foco="PAI" or foco="MAE" then%>
<%else%>				  
<marquee id="mqLooper1" loop="1" onStart="document.inclusao.nome_familiar.focus()"></marquee>
                    <%
					end if%>
</td>
                  <td width="18"> <div align="center"></div></td>
                  <td width="150" height="10">&nbsp;</td>
                  <td width="147" height="10">&nbsp;</td>
                  <td width="17"> <div align="center"></div></td>
                  <td width="149" height="10">&nbsp; </td>
                </tr>
                <tr> 
                  <td height="10" colspan="9"><div id="end"> 
                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr class="tb_corpo"> 
                          <td class="tb_tit"
>Endere&ccedil;o Residencial</td>
                        </tr>
                        <tr class="tb_corpo"> 
                          <td height="10"> <table width="100%" border="0" cellspacing="0">
                              <tr> 
                                <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Logradouro</font></div></td>
                                <td width="13" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                                <td width="217" height="10" class="tb_corpo"
><font class="form_corpo"> 
                                  <%if rua_res="" or isnull(rua_res) then%>
                             <input name="rua_res_fam" type="text" class="borda" id="rua_res_fam" size="30" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'NO_Logradouro_Res')"> 
                                  <%else%>
                                  <input name="rua_res_fam" type="text" class="borda" id="rua_res_fam" value="<%response.write(Server.URLEncode(rua_res))%>" size="30" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'NO_Logradouro_Res')">
                                  <%end if%>
                                  </font></td>
                                <td width="140" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
                                <td width="19" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                                <td width="196" class="tb_corpo"
><font class="form_corpo"> 
                                  <input name="num_res_fam" type="text" class="borda" id="num_res_fam"  value="<%response.write(num_res)%>" size="12" maxlength="10" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'NU_Logradouro_Res')">
                                  </font></td>
                                <td width="90" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
                                <td width="15" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                                <td width="149" height="10" class="tb_corpo"
> <div align="left"><font class="form_corpo"> </font> <font class="form_corpo"> 
                                    <%if comp_res="" or isnull(comp_res) then%>
                                   <input name="comp_res_fam" type="text" class="borda" id="comp_res_fam" size="20" maxlength="15" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'TX_Complemento_Logradouro_Res')">
                                    <%else%>
                                    <input name="comp_res_fam" type="text" class="borda" id="comp_res_fam" value="<%response.write(Server.URLEncode(comp_res))%>" size="20" maxlength="15" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'TX_Complemento_Logradouro_Res')">
                                    <%end if%>
                                    </font></div></td>
                              </tr>
                              <tr> 
                                <td width="145" height="21" class="tb_corpo"
><font class="form_dado_texto">Estado</font></td>
                                <td width="13" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                                <td width="217" height="21" class="tb_corpo"
><font class="form_corpo"> <font class="form_corpo"> 
<%response.Write("recuperarCidResFam(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value)")%>
                                  <select name="estadores_fam" class="borda" id="estadores_fam" onChange="recuperarCidResFam(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value);BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'SG_UF_Res')">
                                    <option value="0" > </option>
                                    <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
SG_UF= RS2("SG_UF")
NO_UF= RS2("NO_UF")

if SG_UF = estadores then
%>
                                    <option value="<%=SG_UF%>" selected> 
                                    <% response.Write(Server.URLEncode(NO_UF))%>
                                    </option>
                                    <% else %>
                                    <option value="<%=SG_UF%>"> 
                                    <% response.Write(Server.URLEncode(NO_UF))%>
                                    </option>
                                    <%
end if	
RS2.MOVENEXT
WEND
%>
                                  </select>
                                  </font> </font></td>
                                <td width="140" height="21" class="tb_corpo"
><font class="form_dado_texto">Cidade</font></td>
                                <td width="19" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                                <td width="196" class="tb_corpo"
><font class="form_corpo"> 
                                  <div id="cid_res_fam"> 
                                    <select name="cidres_fam" class="borda" id="cidres_fam" onChange="recuperarBairroResFam(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value);BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'CO_Municipio_Res')">

                                      <%
if isnull(estadores) or estadores="" then
%>
                                      <option value="0" selected> </option>
                                      <%else%>
                                    <option value="0" > </option>
<%																		  
Set RS2m = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Municipios WHERE SG_UF='"&estadores&"' order by NO_Municipio"
		RS2m.Open SQL2m, CON0
		
while not RS2m.EOF						
SG_UF= RS2m("CO_Municipio")
NO_UF= RS2m("NO_Municipio")

if SG_UF = cidres then
%>
                                      <option value="<%=SG_UF%>" selected> 
                                      <% response.Write(Server.URLEncode(NO_UF))%>
                                      </option>
                                      <% else %>
                                      <option value="<%=SG_UF%>"> 
                                      <% response.Write(Server.URLEncode(NO_UF))%>
                                      </option>
                                      <%
end if	
RS2m.MOVENEXT
WEND
end if
%>
                                    </select>
                                  </div>
                                  </font></td>
                                <td width="90" class="tb_corpo"
><font class="form_dado_texto">Bairro</font></td>
                                <td width="15" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                                <td width="149" height="21" class="tb_corpo"
> <div id="bairro_res_fam"><font class="form_corpo"> 
                                    <select name="bairrores_fam" class="borda" id="bairrores_fam" onChange="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'CO_Bairro_Res')">
                                      <%
if isnull(estadores) or estadores="" or isnull(cidres) or cidres="" then
%>
                                      <option value="0"> </option>
                                      <%else
Set RS2b = Server.CreateObject("ADODB.Recordset")
		SQL2b = "SELECT * FROM TB_Bairros WHERE CO_Municipio="&cidres&" AND SG_UF='"&estadores&"' order by NO_Bairro"
		RS2b.Open SQL2b, CON0
		
while not RS2b.EOF						
SG_UF= RS2b("CO_Bairro")
NO_UF= RS2b("NO_Bairro")
if SG_UF = bairrores then
%>
                                      <option value="<%=SG_UF%>" selected> 
                                      <% response.Write(Server.URLEncode(NO_UF))%>
                                      </option>
                                      <% else %>
                                      <option value="<%=SG_UF%>"> 
                                      <% response.Write(Server.URLEncode(NO_UF))%>
                                      </option>
                                      <%
end if	

RS2b.MOVENEXT
WEND
end if
%>
                                    </select>
                                    </font></div></td>
                              </tr>
                              <tr> 
                                <td width="145" height="10" class="tb_corpo"
><font class="form_dado_texto">CEP</font></td>
                                <td width="13" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                                <td width="217" height="10" class="tb_corpo"
><font class="form_dado_texto"> 
                                  <input name="cep_fam" type="text" class="borda" id="cep_fam" onKeyup="formatar(this, '#####-###')" value="<%response.write(cep)%>" size="11" maxlength="9" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'CO_CEP_Res')">
                                  </font></td>
                                <td width="140" height="10" class="tb_corpo"
>&nbsp;</td>
                                <td width="19" class="tb_corpo"
>&nbsp;</td>
                                <td width="196" class="tb_corpo"
>&nbsp;</td>
                                <td width="90" class="tb_corpo"
>&nbsp;</td>
                                <td width="15" class="tb_corpo"
>&nbsp; </td>
                                <td width="149" height="10" class="tb_corpo"
>&nbsp;</td>
                              </tr>
                              <tr> 
                                <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">Telefones deste endere&ccedil;o</font></div></td>
                                <td width="13" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                                <td height="10" colspan="2" class="tb_corpo"
><font class="form_corpo"> 
                                  <input name="tel_res_fam" type="text" class="borda" id="tel_res_fam" value="<%response.write(tel_res)%>" size="42" maxlength="50" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'NU_Telefones_Res')">
                                  </font> </td>
                                <td width="19" class="tb_corpo"
> <div align="center"></div></td>
                                <td width="196" class="tb_corpo"
>				
</td>
                                <td width="90" class="tb_corpo"
><font class="form_dado_texto">Mesmo endere&ccedil;o do aluno</font></td>
                                <td width="15" class="tb_corpo"
><font class="form_dado_texto">:</font> </td>
                                <td width="149" height="10" class="tb_corpo"
> 
                                  <%if mes_end="s" then%>
                                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                    <tr> 
                                      <td width="9%"><input type="radio" name="mes_end" value="s"  onClick="recuperarEnd('<%response.Write(cod)%>','<%response.Write(foco)%>');BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'ID_Res_Aluno')" checked></td>
                                      <td width="25%"><font class="form_corpo">Sim</font></td>
                                      <td width="5%"><input name="mes_end" type="radio"  onClick="recuperarOrigemEnd('<%response.Write(cod)%>','<%response.Write(foco)%>');BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'ID_Res_Aluno')" value="n"></td>
                                      <td width="61%"><font class="form_corpo">N&atilde;o 
                                        <input name="id_res_familiar_aux" type="hidden" value="s">
                                        </font></td>
                                    </tr>
                                  </table>
                                  <%else%>
                                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                    <tr> 
                                      <td width="9%"><input type="radio" name="mes_end" value="s"  onClick="recuperarEnd('<%response.Write(cod)%>','<%response.Write(foco)%>');BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'ID_Res_Aluno')"></td>
                                      <td width="25%"><font class="form_corpo">Sim</font></td>
                                      <td width="5%"><input name="mes_end" type="radio"  onClick="recuperarOrigemEnd('<%response.Write(cod)%>','<%response.Write(foco)%>');BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'ID_Res_Aluno')" value="n" checked></td>
                                      <td width="61%"><font class="form_corpo">N&atilde;o 
                                        <input name="id_res_familiar_aux" type="hidden" value="n">
                                        </font></td>
                                    </tr>
                                  </table>
                                  <%end if%>
                                </td>
                              </tr>
                            </table></td>
                        </tr>
                      </table>
                    </div>
                    <table width="100%" border="0" cellspacing="0" dwcopytype="CopyTableRow">
                      <tr> 
                        <td height="10" colspan="9" class="tb_tit"
><div align="left">Endere&ccedil;o Comercial </div></td>
                      </tr>
                      <tr> 
                        <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Logradouro</font></div></td>
                        <td width="13" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                        <td width="217" height="10" class="tb_corpo"
><font class="form_corpo">
                                  <%if rua_com="" or isnull(rua_com) then%>
                          <input name="rua_com_fam" type="text" class="borda" id="rua_com_fam" size="30" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'NO_Logradouro_Com')">
                                  <%else%>
                          <input name="rua_com_fam" type="text" class="borda" id="rua_com_fam" value="<%response.write(Server.URLEncode(rua_com))%>" size="30" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'NO_Logradouro_Com')">
                                  <%end if%>

 

                          </font></td>
                        <td width="140" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
                        <td width="19" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                        <td width="196" class="tb_corpo"
><font class="form_corpo"> 
                          <input name="num_com_fam" type="text" class="borda" id="num_com_fam" value="<%response.write(num_com)%>" size="12" maxlength="10" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'NU_Logradouro_Com')">
                          </font></td>
                        <td width="90" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
                        <td width="17" class="tb_corpo"
><div align="center"><font class="form_dado_texto">:</font></div></td>
                        <td width="149" height="10" class="tb_corpo"
> <div align="left"><font class="form_corpo"> </font> <font class="form_corpo"> 
                            <% if isnull(comp_com) or comp_com="" then%>
							<input name="comp_com_fam" type="text" class="borda" id="comp_com" size="20" maxlength="15" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'TX_Complemento_Logradouro_Com')">
							<%else%>
							<input name="comp_com_fam" type="text" class="borda" id="comp_com" value="<%response.write(Server.URLEncode(comp_com))%>" size="20" maxlength="15" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'TX_Complemento_Logradouro_Com')">
							<%end if%>
                            </font></div></td>
                      </tr>
                      <tr class="tb_corpo"
> 
                        <td width="145" height="26"><font class="form_dado_texto">Estado</font></td>
                        <td width="13"> <div align="center"><font class="form_dado_texto">:</font></div></td>
                        <td width="217" height="26"><font class="form_corpo"> 
                          <select name="estadocom_fam" class="borda" id="estadocom_fam" onChange="recuperarCidComFam(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value);BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'SG_UF_Com')">
                                    <option value="0" > </option>
                            <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
SG_UF= RS2("SG_UF")
NO_UF= RS2("NO_UF")
if isnull(uf_natural) then
uf_natural="RJ"
end if
if SG_UF = estadocom then
%>
                            <option value="<%=SG_UF%>" selected> 
                            <% response.Write(Server.URLEncode(NO_UF))%>
                            </option>
                            <%else%>
                            <option value="<%=SG_UF%>"> 
                            <% response.Write(Server.URLEncode(NO_UF))%>
                            </option>
                            <%end if						
RS2.MOVENEXT
WEND
%>
                          </select>
                          </font></td>
                        <td width="140" height="26"><font class="form_dado_texto">Cidade</font></td>
                        <td width="19"> <div align="center"><font class="form_dado_texto">:</font></div></td>
                        <td width="196"> <div id="cid_com_fam"> 
                            <select name="cidcom_fam" class="borda" id="cidcom_fam" onChange="recuperarBairroComFam(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value);BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'CO_Municipio_Com')">
                              <%
if isnull(estadocom) or estadocom="" then
%>
                              <option value="0" selected> </option>
                              <%else %>
                                    <option value="0" > </option>
<%																  
Set RS2m = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Municipios WHERE SG_UF='"&estadocom&"' order by NO_Municipio"
		RS2m.Open SQL2m, CON0
		
while not RS2m.EOF						
SG_UF= RS2m("CO_Municipio")
NO_UF= RS2m("NO_Municipio")

if SG_UF = cidcom then
%>
                              <option value="<%=SG_UF%>" selected> 
                              <% response.Write(Server.URLEncode(NO_UF))%>
                              </option>
                              <% else %>
                              <option value="<%=SG_UF%>"> 
                              <% response.Write(Server.URLEncode(NO_UF))%>
                              </option>
                              <%
end if	
RS2m.MOVENEXT
WEND
end if
%>
                            </select>
                          </div></td>
                        <td width="90"><font class="form_dado_texto">Bairro</font></td>
                        <td><div align="center"><font class="form_dado_texto">:</font></div></td>
                        <td width="149" height="26"> <div id="bairro_com_fam"><font class="form_corpo"> 
                            <select name="bairrocom_fam" class="borda" id="bairrocom_fam" onChange="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'CO_Bairro_Com')">
                              <%
if isnull(estadocom) or estadocom="" or isnull(cidcom) or cidcom="" then
%>
                              <option value="0" selected> </option>
                              <%else
Set RS2b = Server.CreateObject("ADODB.Recordset")
		SQL2b = "SELECT * FROM TB_Bairros WHERE CO_Municipio="&cidcom&" AND SG_UF='"&estadocom&"' order by NO_Bairro"
		RS2b.Open SQL2b, CON0
		
while not RS2b.EOF						
SG_UF= RS2b("CO_Bairro")
NO_UF= RS2b("NO_Bairro")
if SG_UF = bairrocom then
%>
                              <option value="<%=SG_UF%>" selected> 
                              <% response.Write(Server.URLEncode(NO_UF))%>
                              </option>
                              <% else %>
                              <option value="<%=SG_UF%>"> 
                              <% response.Write(Server.URLEncode(NO_UF))%>
                              </option>
                              <%
end if
RS2b.MOVENEXT
WEND
end if
%>
                            </select>
                            </font> </div></td>
                      </tr>
                      <tr class="tb_corpo"
> 
                        <td width="145" height="26"><font class="form_dado_texto">CEP</font></td>
                        <td width="13"> <div align="center"><font class="form_dado_texto">:</font></div></td>
                        <td width="217" height="26"><font class="form_dado_texto"> 
                          <input name="cepcom_fam" type="text" class="borda" id="cepcom_fam" onKeyup="formatar(this, '#####-###')" value="<%response.write(cepcom)%>" size="11" maxlength="9"  onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'CO_CEP_Com')">                          </font></td>
                        <td width="140" height="26">&nbsp;</td>
                        <td width="19">&nbsp;</td>
                        <td width="196">&nbsp;</td>
                        <td width="90">&nbsp;</td>
                        <td>&nbsp;</td>
                        <td width="149" height="26">&nbsp;</td>
                      </tr>
                      <tr class="tb_corpo"
> 
                        <td width="145" height="28"> <div align="left"><font class="form_dado_texto">Telefones 
                            deste endere&ccedil;o<font class="form_dado_texto">:</font></font></div></td>
                        <td width="13"> <div align="center"><font class="form_dado_texto">:</font></div></td>
                        <td height="28" colspan="2"><font class="form_corpo"> 
                          <input name="tel_com_fam" type="text" class="borda" id="tel_com_fam" value="<%response.write(tel_com)%>" size="42" maxlength="50" onBlur="BD_aux(this.value,cod.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'NU_Telefones_Com')">
                          </font> </td>
                        <td width="19">&nbsp; </td>
                        <td width="196">&nbsp;</td>
                        <td width="90">&nbsp;</td>
                        <td>&nbsp;</td>
                        <td width="149" height="28">&nbsp;</td>
                      </tr>
                    </table></td>
                </tr>
              </table>
</div></td>
  </tr>
</table>
</td>
                  </tr>
                </table>
<div id="responsaveis">				
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td height="10" colspan="9" class="tb_tit">Respons&aacute;veis</td>
    </tr>
    <tr> 
      <td width="144" height="10"><font class="form_dado_texto">Financeiro</font></td>
      <td width="12" height="10"><div align="center"><font class="form_dado_texto">:</font></div></td>
      <td width="217" height="10"> <select name="rf" class="borda" onChange="GravaResponsaveis(this.value,'TP_Resp_Fin',0,'TP_Resp_Fin','<%response.write(cod)%>')">
          <option value="0" ></option>
<%if opt="e" or opt="cpf" then

		Set RSCONTPR = Server.CreateObject("ADODB.Recordset")
		SQLCONTPR = "SELECT * FROM TB_Tipo_Contatos Order by NU_Prioridade_Combo"
		RSCONTPR.Open SQLCONTPR, CON0
total_tp_familiares=0		
while not RSCONTPR.EOF	  
cod_familiar = RSCONTPR("TP_Contato")

		Set RSRESP_PED = Server.CreateObject("ADODB.Recordset")
		SQLRESP_PED= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod
		RSRESP_PED.Open SQLRESP_PED, CONCONT_aux
		
if RSRESP_PED.EOF then
else
cod_vinc=RSRESP_PED("CO_Matricula_Vinc")
tp_familiar_vinc=RSRESP_PED("TP_Contato_Vinc")
nome_familiar=RSRESP_PED("NO_Contato")
if (isnull(cod_vinc) or cod_vinc="NULL" or cod_vinc="") and (isnull(tp_familiar_vinc) or tp_familiar_vinc="NULL" or tp_familiar_vinc="") then
else
		response.write "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_familiar_vinc&"' and CO_Matricula ="&cod_vinc

		Set RSRESP_PED_vinc = Server.CreateObject("ADODB.Recordset")
		SQLRESP_PED_vinc= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_familiar_vinc&"' and CO_Matricula ="&cod_vinc
		RSRESP_PED_vinc.Open SQLRESP_PED_vinc, CONCONT_aux						  

nome_familiar=RSRESP_PED_vinc("NO_Contato")
end if
						  

if cod_familiar=resp_ped then
id_familia=RSRESP_PED("ID_Familia")
id_end_bloq=RSRESP_PED("ID_End_Bloqueto")

						  %>
          <option value="<%response.Write(cod_familiar)%>" selected> 
          <%response.Write(Server.URLEncode(nome_familiar))%>
          </option>
          <%
else
						  %>
          <option value="<%response.Write(cod_familiar)%>" > 
          <%response.Write(Server.URLEncode(nome_familiar))%>
          </option>
          <%
end if
end if

RSCONTPR.MOVENEXT
WEND	  

else		  
familiares = Split(ordem_familiares, "##")
for i=1 to ubound(familiares)
cod_nome_familiar=familiares(i)
cod_nome = Split(cod_nome_familiar, "!!")
cod_familiar=cod_nome(0)
							  

		Set RSRESP_FIN = Server.CreateObject("ADODB.Recordset")
		SQLRESP_FIN= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod
		RSRESP_FIN.Open SQLRESP_FIN, CONCONT_aux
		
if RSRESP_FIN.EOF then
else
cod_vinc=RSRESP_FIN("CO_Matricula_Vinc")
tp_familiar_vinc=RSRESP_FIN("TP_Contato_Vinc")
nome_familiar=RSRESP_FIN("NO_Contato")
if (isnull(cod_vinc) or cod_vinc="NULL" or cod_vinc="") and (isnull(tp_familiar_vinc) or tp_familiar_vinc="NULL" or tp_familiar_vinc="") then
else
	'	response.write "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_familiar_vinc&"' and CO_Matricula ="&cod_vinc

		Set RSRESP_FIN_vinc = Server.CreateObject("ADODB.Recordset")
		SQLRESP_FIN_vinc= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_familiar_vinc&"' and CO_Matricula ="&cod_vinc
		RSRESP_FIN_vinc.Open SQLRESP_FIN_vinc, CONCONT_aux						  

nome_familiar=RSRESP_FIN_vinc("NO_Contato")
end if
		  if nome_familiar="" or isnull(nome_familiar) then
		  nome_familiar="Familiar "&tp_familiar_vinc&" sem nome cadastrado"
		  end if

if cod_familiar=resp_fin then
id_familia=RSRESP_FIN("ID_Familia")
id_end_bloq=RSRESP_FIN("ID_End_Bloqueto")

						  %>
          <option value="<%response.Write(cod_familiar)%>" selected> 
          <%
		  response.Write(Server.URLEncode(nome_familiar))%>
          </option>
          <%
else
						  %>
          <option value="<%response.Write(cod_familiar)%>" > 
          <%response.Write(Server.URLEncode(nome_familiar))%>
          </option>
          <%
end if
end if
Next
end if
%>
        </select> </td>
      <td width="140" height="10"><font class="form_dado_texto">Fam&iacute;lia</font></td>
      <td width="19" height="10"><div align="center"><font class="form_dado_texto">:</font></div></td>
      <td width="196" height="10"><input name="id_familia" type="text" class="borda" value="<%response.Write(id_familia)%>" id="rg2" size="30" onBlur="GravaResponsaveis(this.value,'ID_Familia',rf.value,'TP_Resp_Fin','<%response.write(cod)%>')"> 
      </td>
      <td width="90"><font class="form_dado_texto">End. Bloqueto </font></td>
      <td width="11"><div align="center"><font class="form_dado_texto">?</font></div></td>
      <td width="149"><select name="bloq" class="borda" id="bloq" onChange="GravaResponsaveis(this.value,'ID_End_Bloqueto',rf.value,'TP_Resp_Fin','<%response.write(cod)%>')">
          <%if id_end_bloq="R" then
						  %>
          <option value="R" selected> Residencial </option>
          <option value="C"> Comercial </option>
          <%
elseif id_end_bloq="C" then
						  %>
          <option value="R"> Residencial </option>
          <option value="C" selected> Comercial </option>
          <%else%>
          <option value="0" selected></option>
          <option value="R"> Residencial </option>
          <option value="C" > Comercial </option>
          <%
end if
%>
        </select> </td>
    </tr>
    <tr> 
      <td height="10"><font class="form_dado_texto">Pedag&oacute;gico</font></td>
      <td height="10"><div align="center"><font class="form_dado_texto">:</font></div></td>
      <td height="10"> 	  <select name="rp" class="borda" onChange="GravaResponsaveis(this.value,'TP_Resp_Ped',0,'TP_Resp_Ped','<%response.write(cod)%>')">
          <option value="0" ></option>
<%if opt="e" or opt="cpf" then

		Set RSCONTPR = Server.CreateObject("ADODB.Recordset")
		SQLCONTPR = "SELECT * FROM TB_Tipo_Contatos Order by NU_Prioridade_Combo"
		RSCONTPR.Open SQLCONTPR, CON0
total_tp_familiares=0		
while not RSCONTPR.EOF	  
cod_familiar = RSCONTPR("TP_Contato")

		Set RSRESP_PED = Server.CreateObject("ADODB.Recordset")
		SQLRESP_PED= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod
		RSRESP_PED.Open SQLRESP_PED, CONCONT_aux
		
if RSRESP_PED.EOF then
else
cod_vinc=RSRESP_PED("CO_Matricula_Vinc")
tp_familiar_vinc=RSRESP_PED("TP_Contato_Vinc")
nome_familiar=RSRESP_PED("NO_Contato")
if (isnull(cod_vinc) or cod_vinc="NULL" or cod_vinc="") and (isnull(tp_familiar_vinc) or tp_familiar_vinc="NULL" or tp_familiar_vinc="") then
else
		response.write "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_familiar_vinc&"' and CO_Matricula ="&cod_vinc

		Set RSRESP_PED_vinc = Server.CreateObject("ADODB.Recordset")
		SQLRESP_PED_vinc= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_familiar_vinc&"' and CO_Matricula ="&cod_vinc
		RSRESP_PED_vinc.Open SQLRESP_PED_vinc, CONCONT_aux						  

nome_familiar=RSRESP_PED_vinc("NO_Contato")
end if

		  if nome_familiar="" or isnull(nome_familiar) then
		  nome_familiar="Familiar "&tp_familiar_vinc&" sem nome cadastrado"
		  end if
						  

if cod_familiar=resp_ped then
id_familia=RSRESP_PED("ID_Familia")
id_end_bloq=RSRESP_PED("ID_End_Bloqueto")

						  %>
          <option value="<%response.Write(cod_familiar)%>" selected> 
          <%response.Write(Server.URLEncode(nome_familiar))%>
          </option>
          <%
else
						  %>
          <option value="<%response.Write(cod_familiar)%>" > 
          <%response.Write(Server.URLEncode(nome_familiar))%>
          </option>
          <%
end if
end if

RSCONTPR.MOVENEXT
WEND	  

else

familiares = Split(ordem_familiares, "##")
for i=1 to ubound(familiares)
cod_nome_familiar=familiares(i)
cod_nome = Split(cod_nome_familiar, "!!")
cod_familiar=cod_nome(0)
							  

		Set RSRESP_PED = Server.CreateObject("ADODB.Recordset")
		SQLRESP_PED= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod
		RSRESP_PED.Open SQLRESP_PED, CONCONT_aux
		
if RSRESP_PED.EOF then
else
cod_vinc=RSRESP_PED("CO_Matricula_Vinc")
tp_familiar_vinc=RSRESP_PED("TP_Contato_Vinc")
nome_familiar=RSRESP_PED("NO_Contato")
if (isnull(cod_vinc) or cod_vinc="NULL" or cod_vinc="") and (isnull(tp_familiar_vinc) or tp_familiar_vinc="NULL" or tp_familiar_vinc="") then
else
		response.write "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_familiar_vinc&"' and CO_Matricula ="&cod_vinc

		Set RSRESP_PED_vinc = Server.CreateObject("ADODB.Recordset")
		SQLRESP_PED_vinc= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_familiar_vinc&"' and CO_Matricula ="&cod_vinc
		RSRESP_PED_vinc.Open SQLRESP_PED_vinc, CONCONT_aux						  

nome_familiar=RSRESP_PED_vinc("NO_Contato")
end if
						  

if cod_familiar=resp_ped then
id_familia=RSRESP_PED("ID_Familia")
id_end_bloq=RSRESP_PED("ID_End_Bloqueto")

						  %>
          <option value="<%response.Write(cod_familiar)%>" selected> 
          <%response.Write(Server.URLEncode(nome_familiar))%>
          </option>
          <%
else
						  %>
          <option value="<%response.Write(cod_familiar)%>" > 
          <%response.Write(Server.URLEncode(nome_familiar))%>
          </option>
          <%
end if
end if
Next

end if%>
        </select> </td>
      <td height="10">&nbsp;</td>
      <td height="10">&nbsp;</td>
      <td width="196" height="10">&nbsp;</td>
      <td width="90" height="10"><font class="form_dado_texto">End. Circular</font></td>
      <td width="11" height="10"><div align="center"><font class="form_dado_texto">?</font></div></td>
      <td width="149" height="10"><select name="circ" class="borda" id="circ" onChange="GravaResponsaveis(this.value,'ID_End_Bloqueto',rp.value,'TP_Resp_Ped','<%response.write(cod)%>')">
          <%if id_end_bloq="R" then
						  %>
          <option value="R" selected> Residencial </option>
          <option value="C"> Comercial </option>
          <%
elseif id_end_bloq="C" then
						  %>
          <option value="R"> Residencial </option>
          <option value="C" selected> Comercial </option>
          <%else%>
          <option value="0" selected></option>
          <option value="R"> Residencial </option>
          <option value="C"> Comercial </option>
          <%
end if
%>
        </select></td>
    </tr>
  </table>
 </div>				