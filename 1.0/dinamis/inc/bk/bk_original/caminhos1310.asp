<%
ano_letivo = session("ano_letivo") 
ano_vigente = session("ano_vigente")
tipo_arquivo= session("tipo_arquivo") 
escola="dinamis" 
ambiente_escola="testedinamis"
site_escola="www.dinamis.com.br"
transicao = "S"

CAMINHO = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\Cadastro\Logins.mdb"
' Apagar o caminho abaixo assim que desvincular as funções com TB_Aluno_esta_Turma
CAMINHOa = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\AlunoxTurma.mdb"		
'======================================================================================		
CAMINHO_al = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\Cadastro\Alunos.mdb"		
CAMINHO_b = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Boletim.mdb"
CAMINHO_bl = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Bloqueto.mdb"
CAMINHO_contrato = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\BD\Contrato.mdb"
CAMINHO_ct = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\Cadastro\Contatos.mdb"
CAMINHOctl = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\Cadastro\Controle.mdb"
CAMINHO_g = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Grade.mdb"
CAMINHO_log = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\Cadastro\Log.mdb"
CAMINHO_msg = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Mensagem.mdb"		
CAMINHO_na = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_A.mdb"
CAMINHO_nb = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_B.mdb"
CAMINHO_nc = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Modelo_C.mdb"
CAMINHO_o = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Ocorrencias.mdb"
CAMINHO_p = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\Cadastro\Professor.mdb"
CAMINHO_pf = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Posicao.mdb"
CAMINHO_pr = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Parametros.mdb"
CAMINHO_wf = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\Cadastro\WebFamilia.mdb"
caminho_arquivo="\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\docs\"&tipo_arquivo&"\"		
		CAMINHO_ctrle = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\webdiretor\Controle.mdb"
caminho_gera_mov = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\BD\"
caminho_bd = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\BD\"&ano_vigente&"\"
		CAMINHO_wr = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\webdiretor\WebDiretor.mdb"
CAMINHO_t = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\BD\"&ano_letivo&"\Temp.mdb"
CAMINHO_tp = "\\windows-pd-0001.fs.locaweb.com.br\WNFS-0001\simplynet2\Dados\"&ambiente_escola&"\Temp\"




%>