<!--#include file="../../../../inc/connect_al.asp"-->
<!--#include file="../../../../inc/connect_al_aux.asp"-->
<!--#include file="../../../../inc/connect_ct.asp"-->
<!--#include file="../../../../inc/connect_ct_aux.asp"-->
<!--#include file="../../../../inc/connect_pr.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<%
nivel=4
opt=request.querystring("opt")

qtd_tipo_familiares=request.Form("qtd_tp_pub")
foco=request.Form("foco_pub")
cod=request.form("cod_pub")
ordem_familiares=request.Form("ord_pub")
cod_consulta=request.form("cod_vinc_pub")
' abre_campos serve para determinar se existema familiares cadastrados e mostrar os campos relativos aos familiares
'abre_campos="n"
'response.Write(ordem_familiares&">>"&cod&">>"&foco&">>"&qtd_tipo_familiares&">>"&cod_consulta)
		'response.Write (foco&" SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula "&opt&" "&cod&" <BR>")

if opt="cpf" then
cpf_cons=request.Form("cpf_pub")
ordem_familiares=request.Form("ord_pub")
end if

'tem outra determinação da variável dados dentro da opt=cpf mais abaixo
dados=Server.URLEncode(ordem_familiares)&"#sep#"&qtd_tipo_familiares&"#sep#"&foco&"#sep#"&cod_consulta&"#sep#"&cod
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

'Apaga dados preenchidos nas tabelas temporárias==========================================================================
'if cod_consulta=cod or isnull(cod_consulta) or cod_consulta="" then
'cod_consulta=cod
'	aluno_vinculado="n"
'else
'	aluno_vinculado="s"
'end if
		
		Set RSCONTATO_aux_delete = Server.CreateObject("ADODB.Recordset")
		SQLAA_delete= "DELETE * FROM TBI_Alunos WHERE CO_Matricula ="&cod
		RSCONTATO_aux_delete.Open SQLAA_delete, CON1_aux
		
		Set RSCONTATO_aux_delete = Server.CreateObject("ADODB.Recordset")
		SQLAC_delete= "DELETE * FROM TBI_Contatos WHERE CO_Matricula ="&cod
		RSCONTATO_aux_delete.Open SQLAC_delete, CONCONT_aux

			
'Grava dados da BD para tabelas temporárias==========================================================================
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod
		RS.Open SQL, CON1

		if RS.EOF Then
		else
		
		Set RSCONTATO = Server.CreateObject("ADODB.Recordset")
		SQLAA= "SELECT * FROM TB_Contatos WHERE TP_Contato='ALUNO' and CO_Matricula ="&cod
		RSCONTATO.Open SQLAA, CONCONT
		
		codigo_aluno_vinculado=RSCONTATO("CO_Matricula_Vinc")
		tp_vinc_aluno=RSCONTATO("TP_Contato_Vinc")
			if (isnull(tp_vinc_aluno) or tp_vinc_aluno="") and(isnull(codigo_aluno_vinculado) or codigo_aluno_vinculado="") then
				aluno_vinculado="n"
			else
				aluno_vinculado="s"
				Set RSCONTATO_aux_delete = Server.CreateObject("ADODB.Recordset")
				SQLAA_delete= "DELETE * FROM TBI_Alunos WHERE CO_Matricula ="&cod_consulta
				RSCONTATO_aux_delete.Open SQLAA_delete, CON1_aux
				
				Set RSCONTATO_aux_delete = Server.CreateObject("ADODB.Recordset")
				SQLAC_delete= "DELETE * FROM TBI_Contatos WHERE CO_Matricula ="&cod_consulta
				RSCONTATO_aux_delete.Open SQLAC_delete, CONCONT_aux
			end if		
		
		
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
			end if

		end if					
				
		
	conta_familiares_aluno=0	
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
		conta_familiares_aluno=conta_familiares_aluno+1
		
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
			mes_end=RSCONTATO("ID_Res_Aluno")
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
				RSCONTATO_aux_bd("ID_Res_Aluno")=mes_end
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
			end if
			
			if (isnull(tp_vinc_familiar_aux) or tp_vinc_familiar_aux="") and(isnull(co_vinc_familiar_aux) or co_vinc_familiar_aux="") then
			
			else

				
				'response.Write "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_vinc_familiar_aux&"' and CO_Matricula ="&co_vinc_familiar_aux
'				foco_vinculado="s"

				le ="n"
				while le ="n" 
		
					Set RSCONTATO_vinc = Server.CreateObject("ADODB.Recordset")
					SQLA_vinc= "SELECT * FROM TB_Contatos WHERE TP_Contato='"&tp_vinc_familiar_aux&"' and CO_Matricula ="&co_vinc_familiar_aux
					RSCONTATO_vinc.Open SQLA_vinc, CONCONT
					
					
					co_vinc_familiar_aux_mais_um=RSCONTATO_vinc("CO_Matricula_Vinc")
					tp_vinc_familiar_aux_mais_um=RSCONTATO_vinc("TP_Contato_Vinc")
'			response.Write("linha 241<br>")
'			response.Write(SQLA_vinc)
'			response.Write("<br>")
'			response.Write(co_vinc_familiar_aux_mais_um)
'			response.Write("<br>")
'			response.Write(tp_vinc_familiar_aux_mais_um)												
'			response.end()			
		
					if (isnull(tp_vinc_familiar_aux_mais_um) or tp_vinc_familiar_aux_mais_um="") and (isnull(co_vinc_familiar_aux_mais_um) or co_vinc_familiar_aux_mais_um="") then
					le="s"
						co_vinculo_familiar= co_vinc_familiar_aux
						tp_vinculo_familiar=tp_vinc_familiar_aux
						tp_familiar_exibe=tp_familiar_guarda
						
						if RSCONTATO_vinc.EOF then
						else
							nasce_familiar_aux=RSCONTATO_vinc("DA_Nascimento_Contato")
'							if isnull(nasce_familiar_aux) or nasce_familiar_aux="" then
'							else
'								vetor_nascimento = Split(nasce_familiar_aux,"/")  
'								dia_n = vetor_nascimento(0)
'								mes_n = vetor_nascimento(1)
'								ano_n = vetor_nascimento(2)
'								
'								if dia_n<10 then 
'								dia_n = "0"&dia_n
'								end if
'								
'								if mes_n<10 then
'								mes_n = "0"&mes_n
'								end if
'								dia_a = dia_n
'								mes_a = mes_n
'								ano_a = ano_n
'								
'								nasce = dia_n&"/"&mes_n&"/"&ano_n
'							end if
'						nome_contato = RSCONTATO_vinc("NO_Contato")
'						rua_res = RSCONTATO_vinc("NO_Logradouro_Res")
'						num_res = RSCONTATO_vinc("NU_Logradouro_Res")
'						comp_res = RSCONTATO_vinc("TX_Complemento_Logradouro_Res")
'						bairrores= RSCONTATO_vinc("CO_Bairro_Res")
'						cidres= RSCONTATO_vinc("CO_Municipio_Res")
'						estadores= RSCONTATO_vinc("SG_UF_Res")
'						cep = RSCONTATO_vinc("CO_CEP_Res")
'						tel_res = RSCONTATO_vinc("NU_Telefones_Res")
'						tel = RSCONTATO_vinc("NU_Telefones")
'						mail= RSCONTATO_vinc("TX_EMail")
'						ocupacao= RSCONTATO_vinc("CO_Ocupacao")
'						cpf= RSCONTATO_vinc("CO_CPF_PFisica")
'						rg= RSCONTATO_vinc("CO_RG_PFisica")
'						emitido= RSCONTATO_vinc("CO_OERG_PFisica")
'						emissao= RSCONTATO_vinc("CO_DERG_PFisica")
'						empresa= RSCONTATO_vinc("NO_Empresa")
'						rua_com=RSCONTATO_vinc("NO_Logradouro_Com")
'						num_com = RSCONTATO_vinc("NU_Logradouro_Com")
'						comp_com = RSCONTATO_vinc("TX_Complemento_Logradouro_Com")
'						bairrocom= RSCONTATO_vinc("CO_Bairro_Com")
'						cidcom= RSCONTATO_vinc("CO_Municipio_Com")
'						estadocom= RSCONTATO_vinc("SG_UF_Com")
'						cepcom = RSCONTATO_vinc("CO_CEP_Com")
'						tel_com = RSCONTATO_vinc("NU_Telefones_Com")
'						mes_end = RSCONTATO_vinc("ID_Res_Aluno")
'						
							nome_familiar_aux=RSCONTATO_vinc("NO_Contato")
							cpf_familiar_aux=RSCONTATO_vinc("CO_CPF_PFisica")
							rg_familiar_aux=RSCONTATO_vinc("CO_RG_PFisica")
							emitido_familiar_aux=RSCONTATO_vinc("CO_OERG_PFisica")
							emissao_familiar_aux=RSCONTATO_vinc("CO_DERG_PFisica")
							email_familiar_aux=RSCONTATO_vinc("TX_EMail")
							ocupacao_familiar_aux=RSCONTATO_vinc("CO_Ocupacao")
							empresa_familiar_aux=RSCONTATO_vinc("NO_Empresa")
							tel_familiar_aux=RSCONTATO_vinc("NU_Telefones")
							mes_end=RSCONTATO_vinc("ID_Res_Aluno")
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
									RSCONTATO_aux_bd("ID_Res_Aluno")=mes_end
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
									sql_atualiza=sql_atualiza&"NU_Telefones ='"& tel_familiar_aux&"', ID_Res_Aluno = "&mes_end&", NO_Logradouro_Res ='"& rua_res_familiar_aux &"', "& sql_num_res&", TX_Complemento_Logradouro_Res = '"&comp_res_familiar_aux&"', "& sql_bairro_res &", "& sql_cid_res &", SG_UF_Res ='"& uf_res_familiar_aux &"', "
									sql_atualiza=sql_atualiza&"CO_CEP_Res ='"& cep_res_familiar_aux &"', NU_Telefones_Res ='"& tel_res_familiar_aux&"', NO_Logradouro_Com = '"&rua_com_familiar_aux&"', "& sql_num_com&", TX_Complemento_Logradouro_Com= '"&comp_com_familiar_aux&"',"&sql_bairro_com&", "& sql_cid_com &", SG_UF_Com ='"& uf_com_familiar_aux &"', "
									sql_atualiza=sql_atualiza&"CO_CEP_Com='"& cep_com_familiar_aux &"', NU_Telefones_Com ='"& tel_com_familiar_aux&"' WHERE CO_Matricula = "& co_vinc_familiar_aux &" AND TP_Contato = '"& tp_vinc_familiar_aux &"'"
									Set RSCONTATO_aux_bd2 = CONCONT_aux.Execute(sql_atualiza)	
								end if					
						end if
						'para quando eu localizar o vinculado final e trabalhar com os valores corretos
						co_vinc_familiar_aux=co_vinc_familiar_aux_mais_um
						tp_vinc_familiar_aux=tp_vinc_familiar_aux_mais_um
					else
						co_vinc_familiar_aux=co_vinc_familiar_aux_mais_um
						tp_vinc_familiar_aux=tp_vinc_familiar_aux_mais_um
					end if
				wend
	
			end if		
		end if
	next	
	
if aluno_vinculado="s" then


		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod_consulta
		RS.Open SQL, CON1
		

		if RS.EOF Then
			else
			resp_fin= RS("TP_Resp_Fin")
			resp_ped= RS("TP_Resp_Ped")
			pai_fal= RS("IN_Pai_Falecido")
			mae_fal= RS("IN_Mae_Falecida")

			Set RS_aux = Server.CreateObject("ADODB.Recordset")
			SQL_aux = "SELECT * FROM TBI_Alunos WHERE CO_Matricula ="& cod_consulta
			RS_aux.Open SQL_aux, CON1_aux

			if RS_aux.EOF Then
				Set RSALUNO_aux_bd = server.createobject("adodb.recordset")
				RSALUNO_aux_bd.open "TBI_Alunos", CON1_aux, 2, 2
				RSALUNO_aux_bd.addnew
				RSALUNO_aux_bd("CO_Matricula")=cod_consulta
				RSALUNO_aux_bd("TP_Resp_Fin")=resp_fin							  
				RSALUNO_aux_bd("TP_Resp_Ped")=resp_ped
				RSALUNO_aux_bd("IN_Pai_Falecido")=pai_fal							  
				RSALUNO_aux_bd("IN_Mae_Falecida")=mae_fal				
				RSALUNO_aux_bd.update	
				set RSALUNO_aux_bd=nothing
			end if

		end if
		
	for i=1 to ubound(familiares)
		cod_nome_familiar=familiares(i)
		cod_nome = Split(cod_nome_familiar, "!!")
		cod_familiar=cod_nome(0)
		nome_familiar=cod_nome(1)		

		Set RSCONTATO = Server.CreateObject("ADODB.Recordset")
		SQLAA= "SELECT * FROM TB_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod_consulta
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
			mes_end=RSCONTATO("ID_Res_Aluno")
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
			SQLAA_aux= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod_consulta
			RSCONTATO_aux.Open SQLAA_aux, CONCONT_aux

			if RSCONTATO_aux.EOF then
				Set RSCONTATO_aux_bd = server.createobject("adodb.recordset")

				RSCONTATO_aux_bd.open "TBI_Contatos", CONCONT_aux, 2, 2 'which table do you want open
				RSCONTATO_aux_bd.addnew
				RSCONTATO_aux_bd("CO_Matricula")=cod_consulta
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
				RSCONTATO_aux_bd("ID_Res_Aluno")=mes_end
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
				
			end if
		end if
	next
end if
	
elseif opt="cpf" then
' copia os dados da tabela para a tabela auxiliar
familiares = Split(ordem_familiares, "##")

'Apaga dados preenchidos nas tabelas temporárias==========================================================================

		'Set RSCONTATO_aux_delete = Server.CreateObject("ADODB.Recordset")
		'SQLAA_delete= "DELETE * FROM TBI_Contatos WHERE CO_Matricula ="&cod
		'RSCONTATO_aux_delete.Open SQLAA_delete, CONCONT_aux
		
'Grava dados da BD para tabelas temporárias==========================================================================
		
	'for i=1 to ubound(familiares)
	'cod_nome_familiar=familiares(i)
	'cod_nome = Split(cod_nome_familiar, "!!")
	'cod_familiar=cod_nome(0)
	'nome_familiar=cod_nome(1)
	
	'response.Write("SELECT * FROM TB_Contatos WHERE CO_CPF_PFisica='"&cpf_cons&"'")

		Set RSCONTATO = Server.CreateObject("ADODB.Recordset")
		SQLAA= "SELECT * FROM TB_Contatos WHERE CO_CPF_PFisica='"&cpf_cons&"'"
		RSCONTATO.Open SQLAA, CONCONT
		
		if RSCONTATO.EOF then
			recupera_valores="n"
			'response.Write("recupera_valores")
		else
			'while not RSCONTATO.EOF
				co_vinc_familiar_aux=RSCONTATO("CO_Matricula")
				tp_vinc_familiar_aux=RSCONTATO("TP_Contato")	

				'if tp_vinc_familiar_aux=foco and co_vinc_familiar_aux=cod_consulta then
					'recupera_valores="n"
					'RSCONTATO.movenext
				'else
					recupera_valores="s"	

					'if cod_familiar=foco then
						'response.Write("SELECT * FROM TBI_Contatos WHERE TP_Contato='"&foco&"' and CO_Matricula ="&cod&" - recupera_valores")
						Set RSCONTATO_aux = Server.CreateObject("ADODB.Recordset")
						SQLAA_aux= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&foco&"' and CO_Matricula ="&cod
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
								Set RSCONTATO_aux_delete = Server.CreateObject("ADODB.Recordset")
								SQLAA_delete= "DELETE * FROM TBI_Contatos WHERE TP_Contato='"&foco&"' and CO_Matricula ="&cod
								RSCONTATO_aux_delete.Open SQLAA_delete, CONCONT_aux					
						
								Set RSCONTATO_aux_bd2 = server.createobject("adodb.recordset")
								sql_atualiza= "UPDATE TBI_Contatos SET CO_Matricula_Vinc="&co_vinc_familiar_aux&", NO_Contato='', TP_Contato_Vinc ='"& tp_vinc_familiar_aux&"' WHERE CO_Matricula = "& cod &" AND TP_Contato = '"& cod_familiar &"'"
								Set RSCONTATO_aux_bd2 = CONCONT_aux.Execute(sql_atualiza)
						end if
						
						
' Copiando os dados do proprietário do CPF para a tabela temporária
						
						Set RSCONTATO = Server.CreateObject("ADODB.Recordset")
						SQLAA= "SELECT * FROM TB_Contatos WHERE TP_Contato='"&tp_vinc_familiar_aux&"' and CO_Matricula ="&co_vinc_familiar_aux
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
							mes_end=RSCONTATO("ID_Res_Aluno")
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
								RSCONTATO_aux_bd("ID_Res_Aluno")=mes_end
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
								RSCONTATO_aux_bd("CO_Matricula_Vinc")=NULL
								RSCONTATO_aux_bd("TP_Contato_Vinc")=""
								RSCONTATO_aux_bd.update
								set RSCONTATO_aux_bd=nothing
'Não tem else pois se já está cadastrado na tabela term]mporária é por que alguém está alterando os dados e assim utilizarei os dados mais atuais								
							end if
						end if
						
						
						
					'end if
				'RSCONTATO.movenext
				'end if
			'wend
		end if
	'next
'tem outra determinação da variável dados no início do código para as outras opt
dados=Server.URLEncode(ordem_familiares)&"#sep#"&qtd_tipo_familiares&"#sep#"&foco&"#sep#"&co_vinc_familiar_aux&"#sep#"&cod

'end if do if opt="zero"
end if

%>
<table width="100%" height="10" border="0" cellpadding="0" cellspacing="0">
  <tr> 
      <td>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
                    <td>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>

          <%
'Erro nessa matrícula quando carrega o CPF
		'response.Write "SELECT * FROM TBI_Alunos WHERE CO_Matricula ="& cod_consulta
		  
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
nome_familiar="Avós M"
elseif cod_familiar="AVOP" then
nome_familiar="Avós P"
else
nome_familiar=nome_familiar
end if

sit_vinculado="n"

		
	javascript_recupera="BdFamiliar(cod.value,nome_familiar.value,nasce_fam.value,ocupacao_fam.value,trabalho_fam.value,email_fam.value,cpf_fam.value,id_fam.value,tipo_id_fam.value,nasce2_fam.value,tel_fam.value,rua_res_fam.value,num_res_fam.value,comp_res_fam.value,estadores_fam.value,cidres_fam.value,bairrores_fam.value,cep_fam.value,tel_res_fam.value,rua_com_fam.value,num_com_fam.value,comp_com_fam.value,estadocom_fam.value,cidcom_fam.value,bairrocom_fam.value,cepcom_fam.value,tel_com_fam.value,cod_familiar.value,id_res_fam_aux.value,aluno_vinculado.value,co_vinc_familiar_aux.value,tp_vinc_familiar_aux.value);recuperarFamiliares('"&Server.URLEncode(ordem_familiares)&"','"&qtd_tipo_familiares&"','"&Server.URLEncode(cod_familiar)&"','"&cod_consulta&"','"&cod&"')"
	javascript_inclui ="BdFamiliar(cod.value,nome_familiar.value,nasce_fam.value,ocupacao_fam.value,trabalho_fam.value,email_fam.value,cpf_fam.value,id_fam.value,tipo_id_fam.value,nasce2_fam.value,tel_fam.value,rua_res_fam.value,num_res_fam.value,comp_res_fam.value,estadores_fam.value,cidres_fam.value,bairrores_fam.value,cep_fam.value,tel_res_fam.value,rua_com_fam.value,num_com_fam.value,comp_com_fam.value,estadocom_fam.value,cidcom_fam.value,bairrocom_fam.value,cepcom_fam.value,tel_com_fam.value,cod_familiar.value,id_res_fam_aux.value,aluno_vinculado.value,co_vinc_familiar_aux.value,tp_vinc_familiar_aux.value);criaFamiliar('"&Server.URLEncode(ordem_familiares)&"','"&qtd_tipo_familiares&"',this.value,'"&cod_consulta&"','"&cod&"')"
	javascript_primeiro_familiar ="criaFamiliar('"&Server.URLEncode(ordem_familiares)&"','"&qtd_tipo_familiares&"',this.value,'"&cod_consulta&"','"&cod&"')"
'Lê os familiares do Aluno e vê se este familiar é vinculado a outro aluno por isso é cod e não cod_consulta

			'response.Write (foco&" SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula "&cod&"="&opt&" <BR>")


		Set RSCONTATO = Server.CreateObject("ADODB.Recordset")
		SQLAA= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod
		RSCONTATO.Open SQLAA, CONCONT_aux

if RSCONTATO.EOF and cod_familiar<>foco then
combos_responsaveis=combos_responsaveis&"##"&cod_familiar&"!!"&nome_familiar
else

'response.Write(cod_familiar&"="&foco&" and opt="&opt&" then<br>")

if cod_familiar="ALUNO" then

		codigo_aluno_vinculado=RSCONTATO("CO_Matricula_Vinc")
		tp_vinc_aluno=RSCONTATO("TP_Contato_Vinc")
			if (isnull(tp_vinc_aluno) or tp_vinc_aluno="") and(isnull(codigo_aluno_vinculado) or codigo_aluno_vinculado="") then
			aluno_vinculado="n"
			else
			aluno_vinculado="s"
			end if		
		
else
'abre_campos="s"

	if RSCONTATO.EOF and opt<>"i"then 
	else


		if cod_familiar=foco and opt<>"i" then
		
		co_vinc_familiar_aux=RSCONTATO("CO_Matricula_Vinc")
		tp_vinc_familiar_aux=RSCONTATO("TP_Contato_Vinc")
		

		
			if (isnull(tp_vinc_familiar_aux) or tp_vinc_familiar_aux="") and(isnull(co_vinc_familiar_aux) or co_vinc_familiar_aux="") then
			tp_familiar_exibe=tp_familiar_guarda
			foco_vinculado="n"		
			nascimento = RSCONTATO("DA_Nascimento_Contato")
			co_vinculo_familiar=""
			'response.Write(">>"&nascimento)
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
			id_familia=RSCONTATO("ID_Familia")
			id_end_bloq=RSCONTATO("ID_End_Bloqueto")			
		
			
			if mes_end=TRUE then
				mes_end="s"
				else
				mes_end="n"
				end if
			else
			
				foco_vinculado="s"
				le ="n"
				while le ="n" 

				'response.Write "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_vinc_familiar_aux&"' and CO_Matricula ="&co_vinc_familiar_aux
				
					Set RSCONTATO_vinc = Server.CreateObject("ADODB.Recordset")
					SQLA_vinc= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_vinc_familiar_aux&"' and CO_Matricula ="&co_vinc_familiar_aux
					RSCONTATO_vinc.Open SQLA_vinc, CONCONT_aux
					
					co_vinc_familiar_aux_mais_um=RSCONTATO_vinc("CO_Matricula_Vinc")
					tp_vinc_familiar_aux_mais_um=RSCONTATO_vinc("TP_Contato_Vinc")
			
		
					if (isnull(tp_vinc_familiar_aux_mais_um) or tp_vinc_familiar_aux_mais_um="") and (isnull(co_vinc_familiar_aux_mais_um) or co_vinc_familiar_aux_mais_um="") then
					le="s"
						co_vinculo_familiar= co_vinc_familiar_aux
						tp_vinculo_familiar=tp_vinc_familiar_aux
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
						
							
							nome_contato=RSCONTATO_vinc("NO_Contato")
							cpf=RSCONTATO_vinc("CO_CPF_PFisica")
							rg=RSCONTATO_vinc("CO_RG_PFisica")
							emitido=RSCONTATO_vinc("CO_OERG_PFisica")
							emissao=RSCONTATO_vinc("CO_DERG_PFisica")
							mail=RSCONTATO_vinc("TX_EMail")
							ocupacao=RSCONTATO_vinc("CO_Ocupacao")
							empresa=RSCONTATO_vinc("NO_Empresa")
							tel=RSCONTATO_vinc("NU_Telefones")
							mes_end=RSCONTATO_vinc("ID_Res_Aluno")
							id_familia=RSCONTATO_vinc("ID_Familia")
							id_end_bloq=RSCONTATO_vinc("ID_End_Bloqueto")
							rua_res=RSCONTATO_vinc("NO_Logradouro_Res")
							num_res=RSCONTATO_vinc("NU_Logradouro_Res")
							comp_res=RSCONTATO_vinc("TX_Complemento_Logradouro_Res")
							bairrores=RSCONTATO_vinc("CO_Bairro_Res")
							cidres=RSCONTATO_vinc("CO_Municipio_Res")
							estadores=RSCONTATO_vinc("SG_UF_Res")
							cep=RSCONTATO_vinc("CO_CEP_Res")
							tel_res=RSCONTATO_vinc("NU_Telefones_Res")
							rua_com=RSCONTATO_vinc("NO_Logradouro_Com")
							num_com=RSCONTATO_vinc("NU_Logradouro_Com")
							comp_com=RSCONTATO_vinc("TX_Complemento_Logradouro_Com")
							bairrocom=RSCONTATO_vinc("CO_Bairro_Com")
							cidcom=RSCONTATO_vinc("CO_Municipio_Com")
							estadocom=RSCONTATO_vinc("SG_UF_Com")
							cepcom=RSCONTATO_vinc("CO_CEP_Com")
							tel_com=RSCONTATO_vinc("NU_Telefones_Com")
					
			if mes_end=TRUE then
				mes_end="s"
				else
				mes_end="n"
			end if
					
							
												
						end if
						'para quando eu localizar o vinculado final e trabalhar com os valores corretos
						co_vinc_familiar_aux=co_vinc_familiar_aux_mais_um
						tp_vinc_familiar_aux=tp_vinc_familiar_aux_mais_um
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
		sit_vinculado="n"
		%>
						  <td width="70"><div align="right">
					  <input name="botao" class="aba_foco" type="button" id="botao" value="<%RESPONSE.Write(Server.URLEncode(nome_familiar))%>" onClick="<%response.Write(javascript_recupera)%>">
					</div></td>
		<%
		else
			Set RSCONTATO_vinc = Server.CreateObject("ADODB.Recordset")
			SQLA_vinc= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod
			RSCONTATO_vinc.Open SQLA_vinc, CONCONT_aux
				
			co_vinc_familiar_verifica_vinculo=RSCONTATO_vinc("CO_Matricula_Vinc")
			cod_consulta=cod_consulta*1
			co_vinc_familiar_verifica_vinculo=co_vinc_familiar_verifica_vinculo*1
			codigo_aluno_vinculado=codigo_aluno_vinculado*1
		'response.Write(cod_familiar&"-"&codigo_aluno_vinculado&"="&co_vinc_familiar_verifica_vinculo&">>"&aluno_vinculado&"<br>")
		'response.Write(cod_familiar&"-"&cod_consulta&"="&co_vinc_familiar_verifica_vinculo)		
			if isnull(co_vinc_familiar_verifica_vinculo) or co_vinc_familiar_verifica_vinculo="" then
			sit_vinculado="n"
			elseif aluno_vinculado="s" and codigo_aluno_vinculado=co_vinc_familiar_verifica_vinculo then			
			sit_vinculado="s"
			elseif cod_consulta=co_vinc_familiar_verifica_vinculo then
			sit_vinculado="s"
			else
			sit_vinculado="n"
			end if
		larg=larg-70
		%>		          <td width="55"><div align="right">
					  <input name="botao" class="aba_sem_foco"  type="button" id="botao" value="<%RESPONSE.Write(Server.URLEncode(nome_familiar))%>" onClick="<%response.Write(javascript_recupera)%>">
					</div></td>
			<% 


						if (cod_familiar="PAI" and pai_fal = false) or (cod_familiar="MAE" and mae_fal = false) or sit_vinculado="s" then%>
			<%else%>
							  <td width="15"><div align="left">
						  <input name="botao" class="aba_exclui" type="button" id="botao" value="X" onClick="<%response.Write("BdFamiliar(cod.value,nome_familiar.value,nasce_fam.value,ocupacao_fam.value,trabalho_fam.value,email_fam.value,cpf_fam.value,id_fam.value,tipo_id_fam.value,nasce2_fam.value,tel_fam.value,rua_res_fam.value,num_res_fam.value,comp_res_fam.value,estadores_fam.value,cidres_fam.value,bairrores_fam.value,cep_fam.value,tel_res_fam.value,rua_com_fam.value,num_com_fam.value,comp_com_fam.value,estadocom_fam.value,cidcom_fam.value,bairrocom_fam.value,cepcom_fam.value,tel_com_fam.value,cod_familiar.value,id_res_fam_aux.value,aluno_vinculado.value,co_vinc_familiar_aux.value,tp_vinc_familiar_aux.value);ConfirmaExcluirFamiliares('"&Server.URLEncode(ordem_familiares)&"','"&qtd_tipo_familiares&"','"&Server.URLEncode(cod_familiar)&"','"&cod&"')")%>">			  
						</div></td>
			<%
			end if
		end if
	end if	
	end if 
end if 	 
Next
%> <td width="5"><marquee id="mqLooper1" loop="1" onStart="habilita()"></marquee> </td> 
<%larg=larg-5

	conta_familiares_aluno=0	
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
		conta_familiares_aluno=conta_familiares_aluno+1
		end if
	next


if conta_familiares_aluno=1 and aluno_vinculado<>"s" and opt<>"i" then
foco="nulo"
end if

if foco="nulo" then
javascript_combo=javascript_inclui
javascript_combo=javascript_primeiro_familiar
else
javascript_combo=javascript_inclui
end if
%>                   
          <td width="<%RESPONSE.Write(larg)%>"> 
            <div align="left"> 
              <select name="tp1" class="borda" onChange="<%response.Write(javascript_combo)%>">
                <option value="0" selected></option>
<%familiares = Split(combos_responsaveis, "##")
for i=1 to ubound(familiares)
cod_nome_familiar=familiares(i)
cod_nome = Split(cod_nome_familiar, "!!")
cod_familiar=cod_nome(0)
nome_familiar=cod_nome(1)

IF cod_familiar="ALUNO" then
else
%>
                                  <option value="<%response.Write(cod_familiar)%>"> 
                                  <%response.Write(Server.URLEncode(nome_familiar))%>
                                  </option>				
<%end if	  
Next
%>  				
              </select>
            </div>			
			</td>
        </tr>
      </table>
					</td>
                  </tr>
                  <tr> 
                    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
                <td class="tabela_aba"> 
                  <div id="conteudo">
<%
'foco é definido como nulo em executa.asp ou na linha 214
if foco<>"nulo" then%>	
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
                        <td colspan="9" class="tb_tit">Dados Pessoais 
                        </td>
                      </tr>
<%'if abre_campos="s" then%>					  
                      <tr> 
                        <td width="145" height="26"><font class="form_dado_texto">Nome</font></td>
                        <td width="13"> <div align="left"><font class="form_dado_texto">:</font></div></td>
                        <td width="217" height="26"><font class="form_corpo"> 
                          <%if foco="PAI" or foco="MAE" then%>
                          <div id="<%response.Write(foco)%>"> 
                          <input name="nome_familiar" type="text" class="borda" id="nome_familiar" onBlur="ValidaNomeFamiliar(this.value)" value="<%response.Write(Server.URLEncode(nome_contato))%>" size="30" maxlength="60">
                          </div>
                          <%else%>
                          <input name="nome_familiar" type="text" class="borda" id="nome_familiar" onBlur="ValidaNomeFamiliar(this.value)" value="<%response.Write(Server.URLEncode(nome_contato))%>" size="30" maxlength="60">
                          <%
					end if%>
                          </font></td>
                        <td width="140" height="26"><font class="form_dado_texto">Data 
                          de Nascimento</font></td>
                        <td width="19"> 
                          <div align="center"><font class="form_dado_texto">:</font></div></td>
                        <td width="196" height="26"> 
                          <input name="nasce_fam" type="text" class="borda" id="nasce_fam" onKeyup="formatar(this,'##/##/####')" value="<%response.write(nasce)%>" size="12" maxlength="10" onBlur="ValidaDataNasce(this.value)"></td>
                        <td width="90" height="26"> 
                          <div align="left"><font class="form_dado_texto">Rela&ccedil;&atilde;o</font></div></td>
                        <td width="11"> 
                          <div align="center"><font class="form_dado_texto">:</font></div></td>
                        <td width="149" height="26"> <div align="left"><font class="form_corpo"> 
                            <%response.write(Server.URLEncode(tp_familiar_exibe))%>
                            <input name="cod_familiar" type="hidden" class="borda" id="cod_familiar" value="<%response.Write(Server.URLEncode(foco))%>">
                            <input name="ordem_familiares" type="hidden" value="<%response.Write(Server.URLEncode(ordem_familiares))%>">
                            <input name="qtd_tipo_familiares" type="hidden" value="<%response.Write(Server.URLEncode(qtd_tipo_familiares))%>">
                            <input name="aluno_vinculado" type="hidden" value="<%response.Write(aluno_vinculado)%>">
                            <%if opt="i" or foco_vinculado="n" then%>
                            <input name="cod" type="hidden" value="<%response.Write(cod)%>">							
                            <input name="cod_consulta" type="hidden" value="<%response.Write(cod)%>">
                            <%'response.Write("TADA2"&cod)%>
                            <input type="hidden" name="tp_vinc_familiar_aux" value="">
                            <input type="hidden" name="co_vinc_familiar_aux" value="">
                            <%else%>
							<input name="cod" type="hidden" value="<%response.Write(cod)%>"><%'response.Write("TADA "&cod)%>
                            <input name="cod_consulta" type="hidden" value="<%response.Write(co_vinculo_familiar)%>"><%'response.Write(" TADA2 "&co_vinculo_familiar)%>
                            <%'response.Write("TADA - "&cod_consulta)%>
                            <input type="hidden" name="tp_vinc_familiar_aux" value="<%response.Write(tp_vinculo_familiar)%>">
                            <%'response.Write(Server.URLEncode(">>"&tp_vinc_familiar_aux))%>
                            <input type="hidden" name="co_vinc_familiar_aux" value="<%response.Write(co_vinculo_familiar)%>"><%'response.Write(" TADA4 "&tp_vinculo_familiar&"//"&co_vinculo_familiar)%>
                            <%'response.Write("TADA"&co_vinc_familiar_aux)%>
                            <%end if%>
                            </font></div></td>
                      </tr>
                      <tr> 
                        <td width="145" height="26"> <div align="left"><font class="form_dado_texto">Ocupa&ccedil;&atilde;o 
                            </font></div></td>
                        <td width="13"> <div align="left"><font class="form_dado_texto">:</font></div></td>
                        <td width="217" height="26"><font class="form_corpo"> 						
                          <select name="ocupacao_fam" class="borda" id="ocupacao_fam">
                            <option value="0" > </option>
                            <%
							
											
		Set RS_oc = Server.CreateObject("ADODB.Recordset")
		SQL_oc = "SELECT * FROM TB_Ocupacoes order by NO_Ocupacao"
		RS_oc.Open SQL_oc, CON0
		
while not RS_oc.EOF						
co_ocup= RS_oc("CO_Ocupacao")
no_ocup= RS_oc("NO_Ocupacao")
co_ocup=co_ocup*1
ocupacao=ocupacao*1
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
                        <td width="140" height="26"> 
                          <div align="left"><font class="form_dado_texto">Empresa 
                            onde trabalha </font></div></td>
                        <td width="19"> 
                          <div align="center"><font class="form_dado_texto">:</font></div></td>
                        <td width="196" height="26"><font class="form_corpo"> 
                          <%if empresa="" or isnull(empresa) then%>
                          <input name="trabalho_fam" type="text" class="borda" id="trabalho_fam" size="24" maxlength="40" >
                          <%else%>
                          <input name="trabalho_fam" type="text" class="borda" id="trabalho_fam" value="<%response.write(Server.URLEncode(empresa))%>" size="24" maxlength="40">
                          <%end if%>
                          </font></td>
                        <td width="90" height="26"> 
                          <div align="left"><font class="form_dado_texto">E-mail 
                            </font></div></td>
                        <td width="11"> 
                          <div align="center"><font class="form_dado_texto">:</font></div></td>
                        <td width="149" height="26"> <input name="email_fam" type="text" class="borda" id="email_fam" value="<%response.write(mail)%>" size="20" maxlength="50" ></td>
                      </tr>
                      <tr> 
                        <td width="145" height="10"> <div align="left"><font class="form_dado_texto">CPF 
                            </font></div></td>
                        <td width="13"> <div align="left"><font class="form_dado_texto">:</font></div></td>
                        <td width="217" height="10"><font class="form_corpo"> 
                          <%'response.Write("ValidaCPFFamiliar(this.value,"&Server.URLEncode(ordem_familiares)&"','"&Server.URLEncode(qtd_tipo_familiares)&"','"&foco&"','"&cod&"')")%>
                          <input name="cpf_fam" type="text" class="borda" id="cpf_fam" onBlur="ValidaCPFFamiliar(this.value,ordem_familiares.value,qtd_tipo_familiares.value,cod_familiar.value,cod.value)"  onKeyup="formatar(this,'#########-##')" value="<%response.write(cpf)%>" size="15" maxlength="15">
                          </font></td>
                        <td width="140" height="10"> 
                          <div align="left"><font class="form_dado_texto">Identidade 
                            </font></div></td>
                        <td width="19"> 
                          <div align="center"><font class="form_dado_texto">:</font></div></td>
                        <td width="196" height="10"><font class="form_corpo"> 
                          <input name="id_fam" type="text" class="borda" id="id_fam" value="<%response.write(rg)%>" size="15" maxlength="15">
                          </font> </td>
                        <td width="90" height="10"> 
                          <div align="left"><font class="form_dado_texto">Tipo 
                            - Data de Emiss&atilde;o </font></div></td>
                        <td width="11"> 
                          <div align="center"><font class="form_dado_texto">:</font></div></td>
                        <td width="149" height="10"><font class="form_corpo"> 
                          <input name="tipo_id_fam" type="text" class="borda" id="tipo_id_fam" value="<%response.write(emitido)%>" size="10" maxlength="15" >
                          - 
                          <input name="nasce2_fam" type="text" class="borda" id="nasce2_fam" onKeyUp="formatar(this, '##/##/####')" value="<%response.write(emissao)%>" size="11" maxlength="10" onBlur="ValidaDataEmissao(this.value)">
                          </font></td>
                      </tr>
                      <tr> 
                        <td width="145" height="10"><font class="form_dado_texto">Telefones 
                          de Contato</font></td>
                        <td width="13"> <div align="left"><font class="form_dado_texto">:</font> 
                          </div></td>
                        <td width="217" height="10"><font class="form_corpo"> 
                          <input name="tel_fam" type="text" class="borda" id="tel" value="<%response.write(tel)%>" size="42" maxlength="100">
                          </font> </td>
                        <td width="140" height="10"> 
                        </td>
                        <td width="19"> 
                          <div align="center"></div></td>
                        <td width="196" height="10">&nbsp;</td>
                        <td width="90" height="10">&nbsp;</td>
                        <td width="11"> 
                          <div align="center"></div></td>
                        <td width="149" height="10">&nbsp; </td>
                      </tr>
                      <tr> 
                        <td colspan="9"> <div id="end"> 
                            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                              <tr class="tb_corpo"> 
                                <td class="tb_tit"
>Endere&ccedil;o Residencial</td>
                              </tr>
                              <tr class="tb_corpo"> 
                                <td> 
<% if mes_end="s" then		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod
		RS.Open SQL, CON1
		
		
codigo = RS("CO_Matricula")
nome_aluno = RS("NO_Aluno")

sexo = RS("IN_Sexo")

		Set RSCONTA = Server.CreateObject("ADODB.Recordset")
		SQLA = "SELECT * FROM TB_Contatos WHERE TP_Contato ='ALUNO' And CO_Matricula ="& cod
		RSCONTA.Open SQLA, CONCONT

co_vinc_endereco=RSCONTA("CO_Matricula_Vinc")		

if isnull(co_vinc_endereco) or co_vinc_endereco="" then
rua_res = RSCONTA("NO_Logradouro_Res")
num_res = RSCONTA("NU_Logradouro_Res")
comp_res = RSCONTA("TX_Complemento_Logradouro_Res")
bairrores= RSCONTA("CO_Bairro_Res")
cidres= RSCONTA("CO_Municipio_Res")
estadores= RSCONTA("SG_UF_Res")
cep = RSCONTA("CO_CEP_Res")
tel_res = RSCONTA("NU_Telefones_Res")
tel = RSCONTA("NU_Telefones")
empresa= RSCONTA("NO_Empresa")
rua_com=RSCONTA("NO_Logradouro_Com")
num_com = RSCONTA("NU_Logradouro_Com")
comp_com = RSCONTA("TX_Complemento_Logradouro_Com")
bairrocom= RSCONTA("CO_Bairro_Com")
cidcom= RSCONTA("CO_Municipio_Com")
estadocom= RSCONTA("SG_UF_Com")
cepcom = RSCONTA("CO_CEP_Com")
tel_com = RSCONTA("NU_Telefones_Com")

else
		Set RSCONTA = Server.CreateObject("ADODB.Recordset")
		SQLA = "SELECT * FROM TB_Contatos WHERE TP_Contato ='ALUNO' And CO_Matricula ="& co_vinc_endereco
		RSCONTA.Open SQLA, CONCONT

rua_res = RSCONTA("NO_Logradouro_Res")
num_res = RSCONTA("NU_Logradouro_Res")
comp_res = RSCONTA("TX_Complemento_Logradouro_Res")
bairrores= RSCONTA("CO_Bairro_Res")
cidres= RSCONTA("CO_Municipio_Res")
estadores= RSCONTA("SG_UF_Res")
cep = RSCONTA("CO_CEP_Res")
tel_res = RSCONTA("NU_Telefones_Res")
tel = RSCONTA("NU_Telefones")
empresa= RSCONTA("NO_Empresa")
rua_com=RSCONTA("NO_Logradouro_Com")
num_com = RSCONTA("NU_Logradouro_Com")
comp_com = RSCONTA("TX_Complemento_Logradouro_Com")
bairrocom= RSCONTA("CO_Bairro_Com")
cidcom= RSCONTA("CO_Municipio_Com")
estadocom= RSCONTA("SG_UF_Com")
cepcom = RSCONTA("CO_CEP_Com")
tel_com = RSCONTA("NU_Telefones_Com")
end if

session("id_res_familiar")="s"


if isnull(pais) then 
pais = 10
end if

if isnull(estadores) then 
estadores = "RJ"
end if

if isnull(cidres) then 
cidres = 6001
end if

if isnull(estadonat) then 
estadonat = "RJ"
end if

if isnull(nacionalidade) then 
nacionalidade = 1
end if

if isnull(cidnat) then 
cidnat = 6001
end if

if comp_res = "nulo" then 
comp_res = ""
end if

if isnull(cid_cursada) then 
cid_cursada = 6001
end if

if isnull(uf_cursada) then 
uf_cursada = "RJ"
end if


cep5= lEFT(cep, 5)
cep3= Right(cep, 3)


cep=cep5&"-"&cep3

cep5c= lEFT(cepcom, 5)
cep3c= Right(cepcom, 3)


cepcom=cep5c&"-"&cep3c


%>
                                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                    <tr class="tb_corpo"> 
                                      <td height="10"> <table width="100%" border="0" cellspacing="0">
                                          <tr> 
                                            <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Logradouro</font></div></td>
                                            <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                                            <td width="217" height="10" class="tb_corpo"
><font class="form_corpo"> 
                                              <%if isnull(rua_res) or rua_res="" then%>
                                              <input name="rua_res_fam" type="hidden" class="borda" id="rua_res_fam" value="" size="30">
                                              <%else
           response.write(Server.URLEncode(rua_res))%>
                                              <input name="rua_res_fam" type="hidden" class="borda" id="rua_res_fam" value="<%response.write(Server.URLEncode(rua_res))%>" size="30">
                                              <%end if%>
                                              </font></td>
                                            <td width="140" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
                                            <td width="19" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                                            <td width="196" class="tb_corpo"
><font class="form_corpo"> 
                                              <%response.write(num_res)%>
                                              <input name="num_res_fam" type="hidden" class="borda" id="num_res_fam"  value="<%response.write(num_res)%>" size="12" maxlength="10">
                                              </font></td>
                                            <td width="90" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
                                            <td width="11" class="tb_corpo"
> 
                                              <div align="center"><font class="form_dado_texto">:</font></div></td>
                                            <td width="149" height="10" class="tb_corpo"
> 
                                              <div align="left"><font class="form_corpo"> 
                                                <%if isnull(comp_res) or comp_res="" then%>
                                                <input name="comp_res_fam" type="hidden" class="borda" id="comp_res" value="" size="12" maxlength="10">
                                                <%else
response.write(Server.URLEncode(comp_res))%>
                                                <input name="comp_res_fam" type="hidden" class="borda" id="comp_res" value="<%response.write(Server.URLEncode(comp_res))%>" size="12" maxlength="10">
                                                <%end if%>
                                                </font></div></td>
                                          </tr>
                                          <tr> 
                                            <td width="145" height="21" class="tb_corpo"
><font class="form_dado_texto">Estado</font></td>
                                            <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                                            <td width="217" height="21" class="tb_corpo"
><font class="form_corpo"> 
                                              <%
if isnull(estadores)or estadores="" then
%>
                                              <input name="estadores_fam" type="hidden" class="borda" id="estadores_fam" value="RJ" size="12" maxlength="10">
                                              <%
else			  				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF where SG_UF='"&estadores&"'"
		RS2.Open SQL2, CON0

NO_UF= RS2("NO_UF")

response.Write(Server.URLEncode(NO_UF))
%>
                                              <input name="estadores_fam" type="hidden" class="borda" id="estadores_fam" value="<%response.write(Server.URLEncode(estadores))%>" size="12" maxlength="10">
                                              <%
END IF
%>
                                              </font></td>
                                            <td width="140" height="21" class="tb_corpo"
><font class="form_dado_texto">Cidade</font></td>
                                            <td width="19" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                                            <td width="196" class="tb_corpo"
><font class="form_corpo"> 
                                              <%
if isnull(estadores)or estadores="" or cidres="" or isnull(cidres) then
%>
                                              <input name="cidres_fam" type="hidden" class="borda" id="cidres_fam" value="6001" size="12" maxlength="10">
                                              <%
else	
Set RS2m = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Municipios WHERE SG_UF='"&estadores&"' And CO_Municipio="&cidres
		RS2m.Open SQL2m, CON0
		
NO_UF= RS2m("NO_Municipio")
response.Write(Server.URLEncode(NO_UF))
%>
                                              <input name="cidres_fam" type="hidden" class="borda" id="cidres_fam" value="<%response.write(Server.URLEncode(cidres))%>" size="12" maxlength="10">
                                              <%
END IF
%>
                                              </font></td>
                                            <td width="90" class="tb_corpo"
><font class="form_dado_texto">Bairro</font></td>
                                            <td width="11" class="tb_corpo"
> 
                                              <div align="center"><font class="form_dado_texto">:</font></div></td>
                                            <td width="149" height="21" class="tb_corpo"
><font class="form_corpo"> 
                                              <%
if isnull(estadores)or estadores="" or cidres="" or isnull(cidres) or bairrores="" or isnull(bairrores)then
%>
                                              <input name="bairrores_fam" type="hidden" class="borda" id="bairrores_fam" value="6001" size="12" maxlength="10">
                                              <%
else	
Set RS2b = Server.CreateObject("ADODB.Recordset")
		SQL2b = "SELECT * FROM TB_Bairros WHERE CO_Bairro="&bairrores&" AND CO_Municipio="&cidres&" AND SG_UF='"&estadores&"' order by NO_Bairro"
		RS2b.Open SQL2b, CON0
CO_UF= RS2b("CO_Bairro")
NO_UF= RS2b("NO_Bairro")
response.Write(Server.URLEncode(NO_UF))
%>
                                              <input name="bairrores_fam" type="hidden" class="borda" id="bairrores_fam" value="<%response.write(Server.URLEncode(CO_UF))%>" size="12" maxlength="10">
                                              <%
END IF
%>
                                              </font></td>
                                          </tr>
                                          <tr> 
                                            <td width="145" height="10" class="tb_corpo"
><font class="form_dado_texto">CEP</font></td>
                                            <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                                            <td width="217" height="10" class="tb_corpo"
><font class="form_dado_texto"> 
                                              <%response.write(cep)%>
                                              <input name="cep_fam" type="hidden" class="borda" id="cep_fam" value="<%response.write(cep)%>" size="11" maxlength="9">
                                              </font></td>
                                            <td width="140" height="10" class="tb_corpo"
>&nbsp;</td>
                                            <td width="19" class="tb_corpo"
>&nbsp;</td>
                                            <td width="196" class="tb_corpo"
>&nbsp;</td>
                                            <td width="90" class="tb_corpo"
>&nbsp;</td>
                                            <td width="11" class="tb_corpo"
>&nbsp; </td>
                                            <td width="149" height="10" class="tb_corpo"
>&nbsp;</td>
                                          </tr>
                                          <tr> 
                                            <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">Telefones deste endere&ccedil;o</font></div></td>
                                            <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                                            <td height="10" colspan="2" class="tb_corpo"
><font class="form_corpo"> 
                                              <%response.write(tel_res)%>
                                              <input name="tel_res_fam" type="hidden" class="borda" id="tel_res_fam" value="<%response.write(tel_res)%>" size="50" maxlength="50">
                                              </font> </td>
                                            <td width="19" class="tb_corpo"
> <div align="center"></div></td>
                                            <td width="196" class="tb_corpo"
>&nbsp;</td>
                                            <td width="90" class="tb_corpo"
><font class="form_dado_texto">Mesmo endere&ccedil;o do aluno</font></td>
                                            <td width="11" class="tb_corpo"
>
<div align="center"><font class="form_dado_texto">:</font> </div></td>
                                            <td width="149" height="10" class="tb_corpo"
> 
                                              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                                <tr> 
                                                  <td width="9%"><input type="radio" name="mes_end" id="mes_end" value="s"  onClick="recuperarEnd('<%response.Write(cod)%>','<%response.Write(foco)%>')" checked></td>
                                                  <td width="25%"><font class="form_corpo">Sim</font></td>
                                                  <td width="5%"><input name="mes_end" id="mes_end"  type="radio"  onClick="recuperarOrigemEnd('<%response.Write(cod)%>','<%response.Write(foco)%>')" value="n"></td>
                                                  <td width="61%"><font class="form_corpo">N&atilde;o 
                                                    <input name="id_res_fam_aux" type="hidden" id="id_res_fam_aux" value="s">
                                                    </font></td>
                                                </tr>
                                              </table></td>
                                          </tr>
                                        </table></td>
                                    </tr>
                                  </table>
                                  <%else%>
                                  <table width="100%" border="0" cellspacing="0">
                                    <tr> 
                                      <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Logradouro</font></div></td>
                                      <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                                      <td width="217" height="10" class="tb_corpo"
><font class="form_corpo"> 
                                        <%if rua_res="" or isnull(rua_res) then%>
                                        <input name="rua_res_fam" type="text" class="borda" id="rua_res_fam" size="30" maxlength="60">
                                        <%else%>
                                        <input name="rua_res_fam" type="text" class="borda" id="rua_res_fam" value="<%response.write(Server.URLEncode(rua_res))%>" size="30" maxlength="60">
                                        <%end if%>
                                        </font></td>
                                      <td width="140" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
                                      <td width="19" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                                      <td width="196" class="tb_corpo"
><font class="form_corpo"> 
                                        <input name="num_res_fam" type="text" class="borda" id="num_res_fam"  value="<%response.write(num_res)%>" size="12" maxlength="10" onBlur="ValidaNumResFam(this.value)">
                                        </font></td>
                                      <td width="90" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
                                      <td width="11" class="tb_corpo"
> 
                                        <div align="center"><font class="form_dado_texto">:</font></div></td>
                                      <td width="149" height="10" class="tb_corpo"
> <div align="left"><font class="form_corpo"> </font> <font class="form_corpo"> 
                                          <%if comp_res="" or isnull(comp_res) then%>
                                          <input name="comp_res_fam" type="text" class="borda" id="comp_res_fam" size="20" maxlength="30">
                                          <%else%>
                                          <input name="comp_res_fam" type="text" class="borda" id="comp_res_fam" value="<%response.write(Server.URLEncode(comp_res))%>" size="20" maxlength="30">
                                          <%end if%>
                                          </font></div></td>
                                    </tr>
                                    <tr> 
                                      <td width="145" height="21" class="tb_corpo"
><font class="form_dado_texto">Estado</font></td>
                                      <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                                      <td width="217" height="21" class="tb_corpo"
><font class="form_corpo"> <font class="form_corpo"> 
                                        <select name="estadores_fam" class="borda" id="estadores_fam" onChange="recuperarCidResFam(this.value)">
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
                                          <select name="cidres_fam" class="borda" id="cidres_fam" onChange="recuperarBairroResFam(estadores_fam.value,this.value)">
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
                                      <td width="11" class="tb_corpo"
> 
                                        <div align="center"><font class="form_dado_texto">:</font></div></td>
                                      <td width="149" height="21" class="tb_corpo"
> <div id="bairro_res_fam"><font class="form_corpo"> 
                                          <select name="bairrores_fam" class="borda" id="bairrores_fam">
                                            <%
if isnull(estadores) or estadores="" or isnull(cidres) or cidres="" then
%>
                                            <option value="0"> </option>
                                            <%else
Set RS2b = Server.CreateObject("ADODB.Recordset")
		SQL2b = "SELECT * FROM TB_Bairros WHERE CO_Municipio="&cidres&" AND SG_UF='"&estadores&"' order by NO_Bairro"
		RS2b.Open SQL2b, CON0

IF RS2b.EOF then
%>
                                            <option value="0"> 
                                            <% response.Write(Server.URLEncode("Bairros não cadastrados"))%>
                                            </option>
                                            <%else	
		
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
end if
%>
                                          </select>
                                          </font></div></td>
                                    </tr>
                                    <tr> 
                                      <td width="145" height="10" class="tb_corpo"
><font class="form_dado_texto">CEP</font></td>
                                      <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                                      <td width="217" height="10" class="tb_corpo"
><font class="form_dado_texto"> 
                                        <input name="cep_fam" type="text" class="borda" id="cep_fam" onKeyup="formatar(this, '#####-###')" value="<%response.write(cep)%>" size="11" maxlength="9" onBlur="ValidaCepResFam(this.value)">
                                        </font></td>
                                      <td width="140" height="10" class="tb_corpo"
>&nbsp;</td>
                                      <td width="19" class="tb_corpo"
>&nbsp;</td>
                                      <td width="196" class="tb_corpo"
>&nbsp;</td>
                                      <td width="90" class="tb_corpo"
>&nbsp;</td>
                                      <td width="11" class="tb_corpo"
>&nbsp; </td>
                                      <td width="149" height="10" class="tb_corpo"
>&nbsp;</td>
                                    </tr>
                                    <tr> 
                                      <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">Telefones deste endere&ccedil;o</font></div></td>
                                      <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                                      <td height="10" colspan="2" class="tb_corpo"
><font class="form_corpo"> 
                                        <input name="tel_res_fam" type="text" class="borda" id="tel_res_fam" value="<%response.write(tel_res)%>" size="42" maxlength="100">
                                        </font> </td>
                                      <td width="19" class="tb_corpo"
> <div align="center"></div></td>
                                      <td width="196" class="tb_corpo"
> </td>
                                      <td width="90" class="tb_corpo"
><font class="form_dado_texto">Mesmo endere&ccedil;o do aluno</font></td>
                                      <td width="11" class="tb_corpo"
><div align="center"><font class="form_dado_texto">:</font> </div></td>
                                      <td width="149" height="10" class="tb_corpo"
> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                          <tr> 
                                            <td width="9%"><input type="radio" name="mes_end" value="s"  onClick="recuperarEnd('<%response.Write(cod_consulta)%>','<%response.Write(foco)%>')"></td>
                                            <td width="25%"><font class="form_corpo">Sim</font></td>
                                            <td width="5%"><input name="mes_end" type="radio"  onClick="recuperarOrigemEnd('<%response.Write(cod_consulta)%>','<%response.Write(foco)%>')" value="n" checked></td>
                                            <td width="61%"><font class="form_corpo">N&atilde;o 
                                              <input name="id_res_fam_aux" type="hidden" id="id_res_fam_aux" value="n">
                                              </font></td>
                                          </tr>
                                        </table></td>
                                    </tr>
                                  </table>
                                  <%end if%>
                                </td>
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
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                              <td width="217" height="10" class="tb_corpo"
><font class="form_corpo"> 
                                <%if rua_com="" or isnull(rua_com) then%>
                                <input name="rua_com_fam" type="text" class="borda" id="rua_com_fam" size="30" maxlength="60">
                                <%else%>
                                <input name="rua_com_fam" type="text" class="borda" id="rua_com_fam" value="<%response.write(Server.URLEncode(rua_com))%>" size="30" maxlength="60">
                                <%end if%>
                                </font></td>
                              <td width="140" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
                              <td width="19" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                              <td width="196" class="tb_corpo"
><font class="form_corpo"> 
                                <input name="num_com_fam" type="text" class="borda" id="num_com_fam" value="<%response.write(num_com)%>" size="12" maxlength="10" onBlur="ValidaNumComFam(this.value)">
                                </font></td>
                              <td width="90" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
                              <td width="11" class="tb_corpo"
>
<div align="center"><font class="form_dado_texto">:</font></div></td>
                              <td width="149" height="10" class="tb_corpo"
> <div align="left"><font class="form_corpo"> </font> <font class="form_corpo"> 
                                  <% if isnull(comp_com) or comp_com="" then%>
                                  <input name="comp_com_fam" type="text" class="borda" id="comp_com" size="20" maxlength="30">
                                  <%else%>
                                  <input name="comp_com_fam" type="text" class="borda" id="comp_com" value="<%response.write(Server.URLEncode(comp_com))%>" size="20" maxlength="30">
                                  <%end if%>
                                  </font></div></td>
                            </tr>
                            <tr class="tb_corpo"
> 
                              <td width="145" height="26"><font class="form_dado_texto">Estado</font></td>
                              <td width="13"> <div align="left"><font class="form_dado_texto">:</font></div></td>
                              <td width="217" height="26"><font class="form_corpo"> 
                                <select name="estadocom_fam" class="borda" id="estadocom_fam" onChange="recuperarCidComFam(this.value)">
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
                                  <select name="cidcom_fam" class="borda" id="cidcom_fam" onChange="recuperarBairroComFam(estadocom_fam.value,this.value)">
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
                              <td width="11">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                              <td width="149" height="26"> <div id="bairro_com_fam"><font class="form_corpo"> 
                                  <select name="bairrocom_fam" class="borda" id="bairrocom_fam">
                                    <%
if isnull(estadocom) or estadocom="" or isnull(cidcom) or cidcom="" then
%>
                                    <option value="0" selected> </option>
                                    <%else
Set RS2b = Server.CreateObject("ADODB.Recordset")
		SQL2b = "SELECT * FROM TB_Bairros WHERE CO_Municipio="&cidcom&" AND SG_UF='"&estadocom&"' order by NO_Bairro"
		RS2b.Open SQL2b, CON0

IF RS2b.EOF then
%>
                                    <option value="0"> 
                                    <% response.Write(Server.URLEncode("Bairros não cadastrados"))%>
                                    </option>
                                    <%else	
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
end if
%>
                                  </select>
                                  </font> </div></td>
                            </tr>
                            <tr class="tb_corpo"
> 
                              <td width="145" height="26"><font class="form_dado_texto">CEP</font></td>
                              <td width="13"> <div align="left"><font class="form_dado_texto">:</font></div></td>
                              <td width="217" height="26"><font class="form_dado_texto"> 
                                <input name="cepcom_fam" type="text" class="borda" id="cepcom_fam" onKeyup="formatar(this, '#####-###')" value="<%response.write(cepcom)%>" size="11" maxlength="9" onBlur="ValidaCepComFam(this.value)">
                                </font></td>
                              <td width="140" height="26">&nbsp;</td>
                              <td width="19">&nbsp;</td>
                              <td width="196">&nbsp;</td>
                              <td width="90">&nbsp;</td>
                              <td width="11">&nbsp;</td>
                              <td width="149" height="26">&nbsp;</td>
                            </tr>
                            <tr class="tb_corpo"
> 
                              <td width="145" height="28"> <div align="left"><font class="form_dado_texto">Telefones 
                                  deste endere&ccedil;o<font class="form_dado_texto">:</font></font></div></td>
                              <td width="13"> <div align="left"><font class="form_dado_texto">:</font></div></td>
                              <td height="28" colspan="2"><font class="form_corpo"> 
                                <input name="tel_com_fam" type="text" class="borda" id="tel_com_fam" value="<%response.write(tel_com)%>" size="42" maxlength="100">
                                </font> </td>
                              <td width="19">&nbsp; </td>
                              <td width="196">&nbsp;</td>
                              <td width="90">&nbsp;</td>
                              <td width="11">&nbsp;</td>
                              <td width="149" height="28">&nbsp;</td>
                            </tr>
                          </table></td>
                      </tr>
                    </table>
<%end if%>
</div></td>
  </tr>
</table>
</td>
                  </tr>
                </table>
</td>
                  </tr>
<%
'end if do if abre_campos
 'end if%>				  
<div id="responsaveis">
<% if aluno_vinculado="s" then%>
    <tr> 
      <td valign="bottom"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="10" colspan="9" class="tb_tit">Respons&aacute;veis</td>
          </tr>
          <tr> 
            <td width="145" height="10"><font class="form_dado_texto">Financeiro</font></td>
            <td width="13" height="10"> <div align="left"><font class="form_dado_texto">:</font></div></td>
            <td width="217" height="10"> <div align="left"><font class="form_dado_texto"> 
                <%

		Set RSRESPs = Server.CreateObject("ADODB.Recordset")
		SQLs = "SELECT * FROM TBI_Alunos WHERE CO_Matricula ="&codigo_aluno_vinculado
		RSRESPs.Open SQLs, CON1_AUX
		

if RSRESPs.EOF then
else
			resp_fin= RSRESPs("TP_Resp_Fin")

			co_vinc_familiar_aux=cod_consulta

		le ="n"
		while le ="n" 

		'response.Write("<br>SELECT * FROM TBI_Contatos WHERE TP_Contato='"&resp_fin&"' and CO_Matricula ="&co_vinc_familiar_aux)
		Set RSRESP_PED_vinc = Server.CreateObject("ADODB.Recordset")
		SQLRESP_PED_vinc= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&resp_fin&"' and CO_Matricula ="&codigo_aluno_vinculado
		RSRESP_PED_vinc.Open SQLRESP_PED_vinc, CONCONT_aux
			
			co_vinc_familiar_aux_mais_um=RSRESP_PED_vinc("CO_Matricula_Vinc")
			tp_vinc_familiar_aux_mais_um=RSRESP_PED_vinc("TP_Contato_Vinc")
			id_familia=RSRESP_PED_vinc("ID_Familia")
			id_end_bloq=RSRESP_PED_vinc("ID_End_Bloqueto")			
	
			if (isnull(tp_vinc_familiar_aux_mais_um) or tp_vinc_familiar_aux_mais_um="") and (isnull(co_vinc_familiar_aux_mais_um) or co_vinc_familiar_aux_mais_um="") then
			le="s"
			nome_familiar=RSRESP_PED_vinc("NO_Contato")
			'response.Write(" - '"&nome_familiar&"'")
			else
			co_vinc_familiar_aux=co_vinc_familiar_aux_mais_um
			resp_fin=tp_vinc_familiar_aux_mais_um
			end if
		wend 	
end if				  
		if nome_familiar="" or isnull(nome_familiar) then
		nome_familiar="Familiar "&resp_fin&" sem nome cadastrado"
		end if


response.Write(Server.URLEncode(nome_familiar))
%>
                <input name="rf" type="hidden" id="rf" value="<%response.Write(resp_fin)%>">
                </font></div></td>
            <td width="140" height="10"><font class="form_dado_texto">Fam&iacute;lia</font></td>
            <td width="19" height="10"><div align="center"><font class="form_dado_texto">:</font></div></td>
            <td width="196" height="10"><div align="left"><font class="form_dado_texto"> 
                <%response.Write(id_familia)%>
                <input name="id_familia" type="hidden" id="id_familia" value="<%response.Write(id_familia)%>">
                </font></div></td>
            <td width="90"><font class="form_dado_texto">End. Bloqueto </font></td>
            <td width="11"><div align="center"><font class="form_dado_texto">?</font></div></td>
            <td width="149"><div align="left"><font class="form_dado_texto"> 
                <%if id_end_bloq="R" then
						  %>
                Residencial 
                <%
elseif id_end_bloq="C" then%>
                <input name="bloq" type="hidden" id="bloq" value="<%response.Write(id_end_bloq)%>">
                Comercial 
                <%else
end if
%>
                </font></div></td>
          </tr>
          <tr> 
            <td width="145" height="10"><font class="form_dado_texto">Pedag&oacute;gico</font></td>
            <td width="13" height="10"> <div align="left"><font class="form_dado_texto">:</font></div></td>
            <td width="217" height="10"> 
              <div align="left"><font class="form_dado_texto"> 
                <%

		Set RSRESPs = Server.CreateObject("ADODB.Recordset")
		SQLs = "SELECT * FROM TBI_Alunos WHERE CO_Matricula ="&codigo_aluno_vinculado
		RSRESPs.Open SQLs, CON1_AUX
		

if RSRESPs.EOF then
else
			resp_ped= RSRESPs("TP_Resp_Ped")
			tp_vinc_familiar_aux_mais_um=resp_fin

		le ="n"
		while le ="n" 

		'response.Write("<br>SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_familiar_vinc&"' and CO_Matricula ="&cod_vinc)
		Set RSRESP_PED_vinc = Server.CreateObject("ADODB.Recordset")
		SQLRESP_PED_vinc= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&resp_ped&"' and CO_Matricula ="&codigo_aluno_vinculado
		RSRESP_PED_vinc.Open SQLRESP_PED_vinc, CONCONT_aux
			
			co_vinc_familiar_aux_mais_um=RSRESP_PED_vinc("CO_Matricula_Vinc")
			tp_vinc_familiar_aux_mais_um=RSRESP_PED_vinc("TP_Contato_Vinc")
			id_end_circ=RSRESP_PED_vinc("ID_End_Bloqueto")
	
			if (isnull(tp_vinc_familiar_aux_mais_um) or tp_vinc_familiar_aux_mais_um="") and (isnull(co_vinc_familiar_aux_mais_um) or co_vinc_familiar_aux_mais_um="") then
			le="s"
			nome_familiar=RSRESP_PED_vinc("NO_Contato")
			'response.Write(" - '"&nome_familiar&"'")
			else
			co_vinc_familiar_aux=co_vinc_familiar_aux_mais_um
			resp_ped=tp_vinc_familiar_aux_mais_um
			end if
		wend 	
end if				  
		if nome_familiar="" or isnull(nome_familiar) then
		nome_familiar="Familiar "&resp_ped&" sem nome cadastrado"
		end if

response.Write(Server.URLEncode(nome_familiar))
%>
                <input name="rp" type="hidden" id="rp" value="<%response.Write(resp_ped)%>">
                </font></div></td>
            <td width="140" height="10">&nbsp;</td>
            <td width="19" height="10">&nbsp;</td>
            <td width="196" height="10">&nbsp;</td>
            <td width="90" height="10"><font class="form_dado_texto">End. Circular</font></td>
            <td width="11" height="10"><div align="center"><font class="form_dado_texto">?</font></div></td>
            <td width="149" height="10"><div align="left"><font class="form_dado_texto"> 
                <%if id_end_circ="R" then%>
                Residencial 
                <%
elseif id_end_circ="C" then%>
                <input name="circ" type="hidden" id="circ" value="<%response.Write(id_end_circ)%>">
                Comercial 
                <%else
end if
%>
                </font></div></td>
          </tr>
        </table>
</td>
</tr>  
<% else %>
    <tr> 
      <td valign="bottom"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="10" colspan="9" class="tb_tit">Respons&aacute;veis</td>
          </tr>
          <tr> 
            <td width="145" height="10"><font class="form_dado_texto">Financeiro</font></td>
            <td width="13" height="10"> <div align="left"><font class="form_dado_texto">:</font></div></td>
            <td width="217" height="10"> <select name="rf" class="borda" onChange="GravaResponsaveis(this.value,'TP_Resp_Fin',0,'TP_Resp_Fin','<%response.write(cod_consulta)%>')">
                <option value="0" ></option>
                <%

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
		le ="n"
		while le ="n" 

		'response.Write("<br>SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_familiar_vinc&"' and CO_Matricula ="&cod_vinc)
		Set RSRESP_PED_vinc = Server.CreateObject("ADODB.Recordset")
		SQLRESP_PED_vinc= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_familiar_vinc&"' and CO_Matricula ="&cod_vinc
		RSRESP_PED_vinc.Open SQLRESP_PED_vinc, CONCONT_aux
			
			co_vinc_familiar_aux_mais_um=RSRESP_PED_vinc("CO_Matricula_Vinc")
			tp_vinc_familiar_aux_mais_um=RSRESP_PED_vinc("TP_Contato_Vinc")
	
			if (isnull(tp_vinc_familiar_aux_mais_um) or tp_vinc_familiar_aux_mais_um="") and (isnull(co_vinc_familiar_aux_mais_um) or co_vinc_familiar_aux_mais_um="") then
			le="s"
			nome_familiar=RSRESP_PED_vinc("NO_Contato")
			response.Write(" - '"&nome_familiar&"'")
			else
			cod_vinc=co_vinc_familiar_aux_mais_um
			tp_familiar_vinc=tp_vinc_familiar_aux_mais_um
			end if
		wend 	
end if				  
		if nome_familiar="" or isnull(nome_familiar) then
		nome_familiar="Familiar "&tp_familiar_vinc&" sem nome cadastrado"
		end if
		
if cod_familiar=resp_fin then
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
%>
              </select> </td>
            <td width="140" height="10"><font class="form_dado_texto">Fam&iacute;lia</font></td>
            <td width="19" height="10"><div align="center"><font class="form_dado_texto">:</font></div></td>
            <td width="196" height="10"><input name="id_familia" type="text" class="borda" id="rg2" onBlur="GravaResponsaveis(this.value,'ID_Familia',rf.value,'TP_Resp_Fin','<%response.write(cod_consulta)%>')" value="<%response.Write(id_familia)%>" size="30" maxlength="50"> 
            </td>
            <td width="90"><font class="form_dado_texto">End. Bloqueto </font></td>
            <td width="11"><div align="center"><font class="form_dado_texto">?</font></div></td>
            <td width="149"><select name="bloq" class="borda" id="bloq" onChange="GravaResponsaveis(this.value,'ID_End_Bloqueto',rf.value,'TP_Resp_Fin','<%response.write(cod_consulta)%>')">
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
            <td width="145" height="10"><font class="form_dado_texto">Pedag&oacute;gico</font></td>
            <td width="13" height="10"> <div align="left"><font class="form_dado_texto">:</font></div></td>
            <td width="217" height="10"> 
              <select name="rp" class="borda" onChange="GravaResponsaveis(this.value,'TP_Resp_Ped',0,'TP_Resp_Ped','<%response.write(cod)%>')">
                <option value="0" ></option>
                <%

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
		le ="n"
		while le ="n" 

		Set RSRESP_PED_vinc = Server.CreateObject("ADODB.Recordset")
		SQLRESP_PED_vinc= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_familiar_vinc&"' and CO_Matricula ="&cod_vinc
		RSRESP_PED_vinc.Open SQLRESP_PED_vinc, CONCONT_aux
			
			co_vinc_familiar_aux_mais_um=RSRESP_PED_vinc("CO_Matricula_Vinc")
			tp_vinc_familiar_aux_mais_um=RSRESP_PED_vinc("TP_Contato_Vinc")
	
			if (isnull(tp_vinc_familiar_aux_mais_um) or tp_vinc_familiar_aux_mais_um="") and (isnull(co_vinc_familiar_aux_mais_um) or co_vinc_familiar_aux_mais_um="") then
			le="s"
			nome_familiar=RSRESP_PED_vinc("NO_Contato")
			else
			cod_vinc=co_vinc_familiar_aux_mais_um
			tp_familiar_vinc=tp_vinc_familiar_aux_mais_um
			end if
		wend 
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
%>
              </select> </td>
            <td width="140" height="10">&nbsp;</td>
            <td width="19" height="10">&nbsp;</td>
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
</td>
</tr> 
 <%end if%>
</div>
</table>				