<% function buscaContato (P_CO_MATRIC, P_TIPO_CONTATO)

 
	Set conC = Server.CreateObject("ADODB.Connection") 
	ABRIRC = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
	conC.Open ABRIRC
 

	Set RSRC = Server.CreateObject("ADODB.Recordset")
	SQLRC = "SELECT * FROM TB_Contatos where CO_Matricula ="&P_CO_MATRIC&" AND TP_Contato = '"&P_TIPO_CONTATO&"'"
    RSRC.Open SQLRC, conC

IF NOT RSRC.eof THEN
    contato = RSRC("CO_Matricula")&"#!#"&RSRC("TP_Contato")&"#!#"&RSRC("NO_Contato")&"#!#"&RSRC("DA_Nascimento_Contato")&"#!#"&RSRC("CO_CPF_PFisica")
    contato = contato&"#!#"&RSRC("CO_RG_PFisica")&"#!#"&RSRC("CO_OERG_PFisica")&"#!#"&RSRC("CO_DERG_PFisica")&"#!#"&RSRC("TX_EMail")&"#!#"&RSRC("CO_Ocupacao")
    contato = contato&"#!#"&RSRC("NO_Empresa")&"#!#"&RSRC("NU_Telefones")&"#!#"&RSRC("ID_Res_Aluno")&"#!#"&RSRC("NO_Logradouro_Res")&"#!#"&RSRC("NU_Logradouro_Res")
    contato = contato&"#!#"&RSRC("TX_Complemento_Logradouro_Res")&"#!#"&RSRC("CO_Bairro_Res")&"#!#"&RSRC("CO_Municipio_Res")&"#!#"&RSRC("SG_UF_Res")&"#!#"&RSRC("CO_CEP_Res")
    contato = contato&"#!#"&RSRC("NU_Telefones_Res")&"#!#"&RSRC("NO_Logradouro_Com")&"#!#"&RSRC("NU_Logradouro_Com")&"#!#"&RSRC("TX_Complemento_Logradouro_Com")
    contato = contato&"#!#"&RSRC("CO_Bairro_Com")&"#!#"&RSRC("CO_Municipio_Com")&"#!#"&RSRC("SG_UF_Com")&"#!#"&RSRC("CO_CEP_Com")&"#!#"&RSRC("NU_Telefones_Com")
else
    contato = "#!#"&"#!#"&"#!#"&"#!#"
    contato = contato&"#!#"&"#!#"&"#!#"&"#!#"&"#!#"
    contato = contato&"#!#"&"#!#"&"#!#"&"#!#"&"#!#"
    contato = contato&"#!#"&"#!#"&"#!#"&"#!#"&"#!#"
    contato = contato&"#!#"&"#!#"&"#!#"&"#!#"
    contato = contato&"#!#"&"#!#"&"#!#"&"#!#"&"#!#"
END IF

buscaContato = contato
end function
 %>

