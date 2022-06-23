<%
ano_letivo = session("ano_letivo") 
ano_vigente = session("ano_vigente")
tipo_arquivo= session("tipo_arquivo") 
ambiente_escola="bretanha"
transicao = "S"

		CAMINHO = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\Cadastro\Logins.mdb"
		
' Apagar o caminho abaixo assim que desvincular as funções com TB_Aluno_esta_Turma
		CAMINHOa = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\AlunoxTurma.mdb"		
'======================================================================================		
		CAMINHO_al = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\Cadastro\Alunos.mdb"		
		CAMINHO_b = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Boletim.mdb"
		CAMINHO_bl = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Bloqueto.mdb"
		CAMINHO_ct = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\Cadastro\Contatos.mdb"
		CAMINHOctl = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\Cadastro\Controle.mdb"
		CAMINHO_g = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Grade.mdb"
		CAMINHO_log = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\Cadastro\Log.mdb"
		CAMINHO_mca = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Materia_Lecionada_A.mdb"		
		CAMINHO_mcb = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Materia_Lecionada_B.mdb"	
		CAMINHO_mcc = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Materia_Lecionada_C.mdb"					
		CAMINHO_pta = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Pauta_A.mdb"		
		CAMINHO_ptb = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Pauta_B.mdb"	
		CAMINHO_ptc = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Pauta_C.mdb"			
		CAMINHO_msg = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Mensagem.mdb"
		CAMINHO_na = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_A.mdb"
		CAMINHO_nb = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_B.mdb"
		CAMINHO_nc = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_C.mdb"
		CAMINHO_nd = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_D.mdb"
		CAMINHO_ne = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_E.mdb"
		CAMINHO_o = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Ocorrencias.mdb"
		CAMINHO_p = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\Cadastro\Professor.mdb"
		CAMINHO_pf = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Posicao.mdb"
		CAMINHO_pr = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Parametros.mdb"
		CAMINHO_wf = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\Cadastro\WebFamilia.mdb"
		caminho_arquivo="e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\docs\"&tipo_arquivo&"\"		
		CAMINHO_ctrle = "e:\home\simplynetcloud3e1\Dados\webdiretor\Controle.mdb"
        caminho_gera_mov = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"
		caminho_bd = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_vigente&"\"
		CAMINHO_wr = "e:\home\simplynetcloud3e1\Dados\webdiretor\WebDiretor.mdb"
		CAMINHO_t = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Temp.mdb"	
		CAMINHO_tp = "e:\home\simplynetcloud3e1\Dados\"&ambiente_escola&"\Temp\"
%>