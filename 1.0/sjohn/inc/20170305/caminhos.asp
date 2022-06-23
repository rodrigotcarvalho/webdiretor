<%
transicao = "S"
ano_letivo = session("ano_letivo") 
ano_vigente = session("ano_vigente")
ano_letivo_prog_aula = 2013
tipo_arquivo= session("tipo_arquivo") 
ambiente_escola="sjohn"
site_escola="www.saintjohn.g12.br"

'Function caminhos(param_ano_banco, param_ambiente,CAMINHO, CAMINHOa,CAMINHO_al,CAMINHO_b,CAMINHO_bl,CAMINHO_ct,CAMINHOctl,CAMINHO_g,CAMINHO_log,CAMINHO_h,CAMINHO_msg,CAMINHO_na,CAMINHO_nb,CAMINHO_nc,CAMINHO_nd,CAMINHO_ne,CAMINHO_nf,CAMINHO_nk,CAMINHO_nv,CAMINHO_nw,CAMINHO_o,CAMINHO_p,CAMINHO_pf,CAMINHO_pr,CAMINHO_wf,caminho_arquivo,CAMINHO_ctrle,caminho_gera_mov,caminho_bd,CAMINHO_wr,CAMINHO_upload,CAMINHO_t)
' if isnull(param_ano_banco) then
 	ano_banco = session("ano_letivo") 
' else
' 	ano_banco = param_ano_banco
' end if	
' 
' if isnull(param_ambiente) then
 	ambiente = ambiente_escola
' else
' 	ambiente = param_ambiente
' end if	 

 
		CAMINHO = "e:\home\simplynetcloud\dados\"&ambiente&"\Cadastro\Logins.mdb"
		
' Apagar o caminho abaixo assim que desvincular as funções com TB_Aluno_esta_Turma
		CAMINHOa = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\AlunoxTurma.mdb"		
'======================================================================================		
		CAMINHO_al = "e:\home\simplynetcloud\dados\"&ambiente&"\Cadastro\Alunos.mdb"		
		CAMINHO_b = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\Boletim.mdb"
		CAMINHO_bl = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\Bloqueto.mdb"
		CAMINHO_ct = "e:\home\simplynetcloud\dados\"&ambiente&"\Cadastro\Contatos.mdb"
		CAMINHOctl = "e:\home\simplynetcloud\dados\"&ambiente&"\Cadastro\Controle.mdb"
		CAMINHO_g = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\Grade.mdb"
		CAMINHO_log = "e:\home\simplynetcloud\dados\"&ambiente&"\Cadastro\Log.mdb"
		CAMINHO_h = "e:\home\simplynetcloud\dados\"&ambiente&"\Cadastro\Historico.mdb"		
		CAMINHO_msg = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\Mensagem.mdb"			
		CAMINHO_na = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\Modelo_A.mdb"
		CAMINHO_nb = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\Modelo_B.mdb"
		CAMINHO_nc = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\Modelo_C.mdb"
		CAMINHO_nd = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\Modelo_D.mdb"
		CAMINHO_ne = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\Modelo_E.mdb"
		CAMINHO_nf = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\Modelo_F.mdb"
		CAMINHO_nk = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\Modelo_K.mdb"		
		CAMINHO_nv = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\Modelo_V.mdb"				
		CAMINHO_nw = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\Modelo_W.mdb"				
		CAMINHO_o = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\Ocorrencias.mdb"
		CAMINHO_p = "e:\home\simplynetcloud\dados\"&ambiente&"\Cadastro\Professor.mdb"
		CAMINHO_pf = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\Posicao.mdb"
		CAMINHO_pr = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\Parametros.mdb"
		CAMINHO_wf = "e:\home\simplynetcloud\dados\"&ambiente&"\Cadastro\WebFamilia.mdb"
		caminho_arquivo="e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\docs\"&tipo_arquivo&"\"		
		CAMINHO_ctrle = "e:\home\simplynetcloud\dados\webdiretor\Controle.mdb"
		caminho_gera_mov = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"
		caminho_bd = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_vigente&"\"
		CAMINHO_wr = "e:\home\simplynetcloud\dados\webdiretor\WebDiretor.mdb"
		CAMINHO_upload = "e:\home\simplynetcloud\dados\"&ambiente&"\temp\"
		CAMINHO_t = "e:\home\simplynetcloud\dados\"&ambiente&"\BD\"&ano_banco&"\Temp.mdb"	

'end Function	

'call caminhos(null, null,CAMINHO, CAMINHOa,CAMINHO_al,CAMINHO_b,CAMINHO_bl,CAMINHO_ct,CAMINHOctl,CAMINHO_g,CAMINHO_log,CAMINHO_h,CAMINHO_msg,CAMINHO_na,CAMINHO_nb,CAMINHO_nc,CAMINHO_nd,CAMINHO_ne,CAMINHO_nf,CAMINHO_nk,CAMINHO_nv,CAMINHO_nw,CAMINHO_o,CAMINHO_p,CAMINHO_pf,CAMINHO_pr,CAMINHO_wf,caminho_arquivo,CAMINHO_ctrle,caminho_gera_mov,caminho_bd,CAMINHO_wr,CAMINHO_upload,CAMINHO_t)	

%>