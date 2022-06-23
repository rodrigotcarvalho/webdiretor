<%
transicao = "S"
ano_letivo = session("ano_letivo") 
ano_vigente = session("ano_vigente")
tipo_arquivo= session("tipo_arquivo") 
ambiente_escola="tictac"
nome_da_escola="Tic Tic Tac Educaчуo Infantil"
email_suporte_escola="suportewebdiretortictictac@webdiretor.com.br"
site_escola="www.tictictac.com.br"
email_financeiro="suportewebdiretortictictac@webdiretor.com.br"

		ACAMINHO = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\Temp\Auxiliares.mdb"	
		
		CAMINHO = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\Cadastro\Logins.mdb"
	
' Apagar o caminho abaixo assim que desvincular as funчѕes com TB_Aluno_esta_Turma
		CAMINHOa = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_letivo&"\AlunoxTurma.mdb"		
'======================================================================================		
		CAMINHO_al = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\Cadastro\Alunos.mdb"		
		CAMINHO_ax = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Almoxarifado.mdb"		
		CAMINHO_b = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Boletim.mdb"
		CAMINHO_bl = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Bloqueto.mdb"
		CAMINHO_ct = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\Cadastro\Contatos.mdb"
		CAMINHOctl = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\Cadastro\Controle.mdb"
		CAMINHO_e = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Entrevistas.mdb"		
		CAMINHO_ei = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\Cadastro\Entrevistas_Inicial.mdb"		
		CAMINHO_g = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Grade.mdb"
		CAMINHO_log = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\Cadastro\Log.mdb"
		CAMINHO_msg = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Mensagem.mdb"		
		CAMINHO_na = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_A.mdb"
		CAMINHO_nb = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_B.mdb"
		CAMINHO_nc = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_C.mdb"
		CAMINHO_nd = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_D.mdb"
		CAMINHO_ne = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_E.mdb"	
		CAMINHO_nf = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_F.mdb"			
		CAMINHO_o = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Ocorrencias.mdb"
		CAMINHO_p = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\Cadastro\Professor.mdb"
		CAMINHO_pf = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Posicao.mdb"
		CAMINHO_pr = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Parametros.mdb"
		CAMINHO_wf = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\Cadastro\WebFamilia.mdb"		
		caminho_arquivo="e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_letivo&"\docs\"&tipo_arquivo&"\"		
		CAMINHO_ctrle = "e:\home\simplynetcloud1e1\dados\webdiretor\Controle.mdb"
		caminho_gera_mov = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"
		caminho_bd = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_vigente&"\"
		CAMINHO_wr = "e:\home\simplynetcloud1e1\dados\webdiretor\WebDiretor.mdb"
		CAMINHO_t = "e:\home\simplynetcloud1e1\dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Temp.mdb"
		CAMINHO_tp = "e:\home\simplynetcloud1e1\Dados\"&ambiente_escola&"\Temp\"
	
%>