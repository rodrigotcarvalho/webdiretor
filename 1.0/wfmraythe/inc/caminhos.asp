<%
transicao = "S"
ano_letivo = session("ano_letivo") 
ano_vigente = session("ano_vigente")
tipo_arquivo= session("tipo_arquivo") 
ambiente_escola="mraythe"
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
		CAMINHO_log = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\Cadastro\Log.mdb"
		CAMINHO_na = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_A.mdb"
		CAMINHO_nb = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_B.mdb"
		CAMINHO_nc = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_C.mdb"
		CAMINHO_o = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Ocorrencias.mdb"
		CAMINHO_p = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\Cadastro\Professor.mdb"
		CAMINHO_pf = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Posicao.mdb"
		CAMINHO_pr = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Parametros.mdb"
		CAMINHO_wf = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\Cadastro\WebFamilia.mdb"
		caminho_arquivo="e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\docs\"&tipo_arquivo&"\"		
		CAMINHO_ctrle = "e:\home\simplynetcloud2e1\Dados\webdiretor\Controle.mdb"
		caminho_gera_mov = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"
		caminho_bd = "e:\home\simplynetcloud2e1\Dados\"&ambiente_escola&"\BD\"&ano_vigente&"\"
		CAMINHO_wr = "e:\home\simplynetcloud2e1\Dados\webdiretor\WebDiretor.mdb"

%>