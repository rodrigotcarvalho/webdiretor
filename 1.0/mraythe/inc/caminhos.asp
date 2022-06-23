<%
transicao = "N"
ano_letivo = session("ano_letivo") 
ano_vigente = session("ano_vigente")
tipo_arquivo= session("tipo_arquivo") 
escola="mraythe"
session("escola") = escola
ambiente_escola="mraythe"
site_escola="http://www.maria-raythe.com.br/"
nome_da_escola="Col&eacute;gio Maria Raythe"
email_suporte_escola="suportewebdiretormraythe@simplynet.com.br"
email_financeiro = "tesouraria@maria-raythe.com.br"

		CAMINHO = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\Cadastro\Logins.mdb"
		
' Apagar o caminho abaixo assim que desvincular as funções com TB_Aluno_esta_Turma
		CAMINHOa = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\AlunoxTurma.mdb"		
'======================================================================================		
		CAMINHO_al = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\Cadastro\Alunos.mdb"		
		CAMINHO_b = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Boletim.mdb"
		CAMINHO_bl = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Bloqueto.mdb"
		CAMINHO_ct = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\Cadastro\Contatos.mdb"
		CAMINHOctl = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\Cadastro\Controle.mdb"
		CAMINHO_g = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Grade.mdb"
		CAMINHO_h = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\Cadastro\Historico.mdb"		
		CAMINHO_log = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\Cadastro\Log.mdb"
		CAMINHO_msg = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Mensagem.mdb"
		CAMINHO_na = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_A.mdb"
		CAMINHO_nb = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_B.mdb"
		CAMINHO_nc = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_C.mdb"
		CAMINHO_nd = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_D.mdb"
		CAMINHO_ne = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_E.mdb"		
		CAMINHO_o = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Ocorrencias.mdb"
		CAMINHO_p = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\Cadastro\Professor.mdb"
		CAMINHO_mca = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Materia_Lecionada_A.mdb"		
		CAMINHO_mcb = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Materia_Lecionada_B.mdb"	
		CAMINHO_mcc = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Materia_Lecionada_C.mdb"					
		CAMINHO_pta = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Pauta_A.mdb"		
		CAMINHO_ptb = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Pauta_B.mdb"	
		CAMINHO_ptc = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Pauta_C.mdb"					
		CAMINHO_pf = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Posicao.mdb"
		CAMINHO_pr = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Parametros.mdb"
		if ano_letivo = ano_vigente then
			CAMINHO_wf = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\Cadastro\WebFamilia.mdb"
		else
			CAMINHO_wf = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\WebFamilia.mdb"			
		end if	
		caminho_arquivo="e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\docs\"&tipo_arquivo&"\"		
		CAMINHO_ctrle = "e:\home\simplynetcloud2e1\Dados\webdiretor\Controle.mdb"
		caminho_gera_mov = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"
		caminho_bd = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_vigente&"\"
		CAMINHO_wr = "e:\home\simplynetcloud2e1\Dados\webdiretor\WebDiretor.mdb"
		CAMINHO_t = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Temp.mdb"	
		CAMINHO_tp = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\Temp\"
%>