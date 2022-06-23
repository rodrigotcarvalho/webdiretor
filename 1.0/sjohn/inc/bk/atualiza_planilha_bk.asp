<%'On Error Resume Next%>
<!--#include file="caminhos.asp"-->
<!--#include file="funcoes.asp"--> 

<%		
function atualiza_planilha(opt,unidade, curso, etapa, turma, periodo, p_cod_materia, p_matric, p_valor, outro)

valor_recebido = p_valor
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 
data = dia &"/"& mes &"/"& ano
horario = hora & ":"& min

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1

		Set CONg = Server.CreateObject("ADODB.Connection") 
		ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONg.Open ABRIRg	

tb_nota = tabela_notas(CONg, unidade, curso, etapa, turma, periodo, p_co_materia, outro)

CAMINHOn = caminho_notas(CONg, tb_nota, outro)	



    dados_tabela=verifica_dados_tabela(CAMINHOn,opt,outro)
	dados_separados=split(dados_tabela,"#$#")
	tb=dados_separados(0)


		Set CONOTA = Server.CreateObject("ADODB.Connection") 
		ABRIROTA  = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONOTA.Open ABRIROTA 

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0

Set RSR = Server.CreateObject("ADODB.Recordset")
'SQLR = "SELECT * FROM TB_Programa_Aula WHERE CO_Curso = '"&curso&"' AND CO_Etapa = '"&etapa&"' AND (((('"&opt&"'='LBS') or ('"&opt&"'='LSM') or '"&opt&"'='LBG' AND (CO_Materia='"&p_cod_materia&" or "&p_cod_materia&" = "")) and TP_Disciplina = 'R') OR ('"&opt&"'='LBA' AND CO_Materia='"&p_cod_materia&"')) order by NU_Ordem_Boletim"
SQLR = "SELECT * FROM TB_Programa_Aula WHERE CO_Curso = '"&curso&"' AND CO_Etapa = '"&etapa&"' AND (((('"&opt&"'='LBS') or '"&opt&"'='LBG' AND (CO_Materia='"&p_cod_materia&"' or '"&p_cod_materia&"' = '')) and TP_Disciplina = 'R') OR ('"&opt&"'='LBA' AND CO_Materia='"&p_cod_materia&"')) order by NU_Ordem_Boletim"

RSR.Open SQLR, CON0
'response.write(SQLR&"<BR>")
'p_matric=p_matric*1
'if p_matric = 32129 then
    
'  response.end()

'end if

while not RSR.EOF

co_materia = RSR("CO_Materia")
ehMae = RSR("IN_MAE")
		
	Set RSMP = Server.CreateObject("ADODB.Recordset")
	SQLMP = "SELECT * FROM TB_Materia WHERE CO_Materia='"&co_materia&"'"
	RSMP.Open SQLMP, CON0
	
	if RSMP.EOF then
		mat_princ=co_materia
	else
		materia_mae=RSMP("CO_Materia_Principal")
		if isnull(materia_mae) or materia_mae="" then
			mat_princ=co_materia					
		else
			mat_princ=materia_mae		
		end if	
	end if
    if ehMae then
        zera_media="N"	
    else
        if materia_mae_anterior = mat_princ then
           zera_media="S"	
        else
           zera_media="N"
           materia_mae_anterior=mat_princ	
        end if
    end if
    valor_gravacao = valor_recebido
    curso=curso*1
    if zera_media="S" and isnumeric(valor_recebido)  and curso = 2 then
        valor_gravacao=0
    end if
	 
'response.Write( valor_recebido&"=<BR>")

		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL_N = "Select * from "& tb_nota &" WHERE CO_Matricula = "&p_matric&" AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
		'response.Write(SQL_N&"<BR>")		
		Set RS3 = CONOTA.Execute(SQL_N)	



'response.Write("Atualizando planilha de notas "&co_materia&" "&RSR("TP_Disciplina")&"<BR>")
		while not RS3.EOF			
			
					
			
				if tb_nota="TB_NOTA_L" then
					va_t1 = RS3("VA_Teste1") 
					va_t2 = RS3("VA_Teste2")
					va_t3 = RS3("VA_Teste3") 
					va_t4 = RS3("VA_Teste4")
					va_p1 = RS3("VA_Prova1")
					va_p2 = RS3("VA_Prova2")				
					va_sim = RS3("VA_Sim")	
					va_bat = RS3("VA_Bat")
					va_bon = RS3("VA_Bonus")
					va_rec = RS3("VA_Rec")	
					
					if opt="LBA" or opt="LBG" then					
						va_bat = valor_gravacao
					elseif opt="LBS" then		
						va_bsi = valor_gravacao								
					else
						'LSM não utiliza mais essa função e sim a atualiza_planilha_simulado mais abaixo					
						if valor_gravacao=0 then	
							va_sim = ""
						end if						
					end if											
				
					fail = calcula_medias_L(CONOTA, curso, etapa, co_materia, mat_princ, periodo, p_matric, va_t1, va_t2, va_t3, va_t4, va_p1, va_p2, va_sim, va_bat, va_bon, va_rec)
				else
					va_av1 = RS3("VA_Av1") 
					va_av2 = RS3("VA_Av2")
					va_av3 = RS3("VA_Av3") 
					va_av4 = RS3("VA_Av4")
					va_av5 = RS3("VA_Av5")
					va_sim = RS3("VA_Sim")	
					va_bat = RS3("VA_Bat")						
					va_bsi = RS3("VA_Bsi")				
					va_bon = RS3("VA_Bonus")
					va_rec = RS3("VA_Rec")		
					
					if opt="LBA" or opt="LBG" then					
						va_bat = valor_gravacao
					elseif opt="LBS" then		
						va_bsi = valor_gravacao								
					else	
						'LSM não utiliza mais essa função e sim a atualiza_planilha_simulado mais abaixo
						if valor_gravacao=0 then	
							va_sim = ""
						end if					
					end if																
	
					fail = calcula_medias_M(CONOTA, curso, etapa, co_materia, mat_princ, periodo, p_matric, va_av1, va_av2, va_av3, va_av4, va_av5, va_sim, va_bat, va_bsi, va_bon, va_rec)
				end if				
			
			if fail = 1 then
				response.Write("ERRO "&tb_nota)
				response.End()
			end if
			
		RS3.Movenext
		wend	
RSR.Movenext
wend	
'response.End()	
end function
















function atualiza_planilha_simulado(opt,unidade, curso, etapa, turma, periodo, p_areadoconhecimento, p_matric, p_valor, outro)

Server.ScriptTimeout = 600 'valor em segundos

valor_recebido = p_valor
'response.Write( p_valor&"=-<BR>")
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 
data = dia &"/"& mes &"/"& ano
horario = hora & ":"& min

	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1

	Set CONg = Server.CreateObject("ADODB.Connection") 
	ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONg.Open ABRIRg	

	tb_nota = tabela_notas(CONg, unidade, curso, etapa, turma, periodo, p_co_materia, outro)
	
	CAMINHOn = caminho_notas(CONg, tb_nota, outro)	



    dados_tabela=verifica_dados_tabela(CAMINHOn,opt,outro)
	dados_separados=split(dados_tabela,"#$#")
	tb=dados_separados(0)



	Set CONOTA = Server.CreateObject("ADODB.Connection") 
	ABRIROTA  = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONOTA.Open ABRIROTA 

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0

	Set RSAC = Server.CreateObject("ADODB.Recordset")		
	CONEXAOAC = "SELECT TB_Area_ConhecimentoxTB_Programa_Aula.[CO_Materia] AS MATERIA FROM TB_Area_ConhecimentoxTB_Programa_Aula WHERE TB_Area_ConhecimentoxTB_Programa_Aula.[CO_Curso]='"&curso&"' AND TB_Area_ConhecimentoxTB_Programa_Aula.[CO_Etapa]='"&etapa&"' AND (TB_Area_ConhecimentoxTB_Programa_Aula.[TP_Area]="&p_areadoconhecimento&" or "&p_areadoconhecimento&" = 0) ORDER BY TB_Area_ConhecimentoxTB_Programa_Aula.[CO_Materia]"
	Set RSAC = CON0.Execute(CONEXAOAC)		

'response.Write(CONEXAOAC&"<BR>")
				
	while not RSAC.EOF
	
	co_materia = RSAC("MATERIA")
	
	
		Set RSR = Server.CreateObject("ADODB.Recordset")
		SQLR = "SELECT * FROM TB_Programa_Aula WHERE CO_Curso = '"&curso&"' AND CO_Etapa = '"&etapa&"' and CO_Materia='"&co_materia&"' order by NU_Ordem_Boletim"
		RSR.Open SQLR, CON0
		
'		ehMae = RSR("IN_MAE")
'		ehFilha = RSR("IN_FIL")	
				
		Set RSMP = Server.CreateObject("ADODB.Recordset")
		SQLMP = "SELECT * FROM TB_Materia WHERE CO_Materia='"&co_materia&"'"
		RSMP.Open SQLMP, CON0
		
		if not RSMP.EOF then
			materia_mae=RSMP("CO_Materia_Principal")
			if isnull(materia_mae) or materia_mae="" then
				'Se não existe matéria principal então a matéria principal é ela própria			
				mat_princ=co_materia					
			else
				mat_princ=materia_mae		
			end if		

			Set RSFIL = Server.CreateObject("ADODB.Recordset")
			SQLFIL = "SELECT * FROM TB_Materia WHERE CO_Materia_Principal='"&mat_princ&"'"
			
			RSFIL.Open SQLFIL, CON0			


			if RSFIL.EOF then
				vetor_materia_simulado = mat_princ&"#!#"&mat_princ&"#!#"&valor_recebido
			else
				conta_filhas = 0	
							
				while not RSFIL.EOF 	
					materia_fil=RSFIL("CO_Materia")
					'response.Write(materia_fil&"<BR>")	
					if conta_filhas = 0 then
						if materia_fil = co_materia then		
							vetor_materia_simulado = mat_princ&"#!#"&materia_fil&"#!#"&valor_recebido
						else
							vetor_materia_simulado = mat_princ&"#!#"&materia_fil&"#!#0"
						end if	
					else
						if materia_fil = co_materia then		
							vetor_materia_simulado = vetor_materia_simulado&"#$#"&mat_princ&"#!#"&materia_fil&"#!#"&valor_recebido
						else
							vetor_materia_simulado = vetor_materia_simulado&"#$#"&mat_princ&"#!#"&materia_fil&"#!#0"
						end if										
					end if	
					conta_filhas=conta_filhas+1							
				RSFIL.Movenext
				wend
			end if				
		end if
		
		
		
'		if ehMae then
'			zera_media="N"	
'		else
'			if materia_mae_anterior = mat_princ then
'			   zera_media="S"	
'			else
'			   zera_media="N"
'			   materia_mae_anterior=mat_princ	
'			end if
'		end if
'		valor_gravacao = valor_recebido
'		
'		curso=curso*1
'		if zera_media="S" and isnumeric(valor_recebido)  and curso = 2 then
'			valor_gravacao=0
'		end if
	
		 
'	response.Write( vetor_materia_simulado&"=<BR>")
	
		percorre_disciplinas = split(vetor_materia_simulado,"#$#")
		for p =0 to ubound(percorre_disciplinas)
		valores_disciplinas = split(percorre_disciplinas(p),"#!#")
			for v=0 to ubound(valores_disciplinas)
				mat_principal = valores_disciplinas(0)
				co_materia_fil = valores_disciplinas(1)
				valor_gravacao = valores_disciplinas(2)
			
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL_N = "Select * from "& tb_nota &" WHERE CO_Matricula = "&p_matric&" AND CO_Materia_Principal = '"& mat_principal &"' AND CO_Materia = '"& co_materia_fil &"' AND NU_Periodo="&periodo
				'response.Write("2 "&SQL_N&"<BR>")		
				Set RS3 = CONOTA.Execute(SQL_N)	
		
		
		
		'response.Write("Atualizando planilha de notas "&co_materia&" "&RSR("TP_Disciplina")&"<BR>")
				while not RS3.EOF											
					
					if tb_nota="TB_NOTA_L" then
						va_t1 = RS3("VA_Teste1") 
						va_t2 = RS3("VA_Teste2")
						va_t3 = RS3("VA_Teste3") 
						va_t4 = RS3("VA_Teste4")
						va_p1 = RS3("VA_Prova1")
						va_p2 = RS3("VA_Prova2")				
						va_sim = valor_gravacao		
						va_bat = RS3("VA_Bat")
						va_bon = RS3("VA_Bonus")
						va_rec = RS3("VA_Rec")	
									
		
						fail = calcula_medias_L(CONOTA, curso, etapa, co_materia_fil, mat_principal, periodo, p_matric, va_t1, va_t2, va_t3, va_t4, va_p1, va_p2, va_sim, va_bat, va_bon, va_rec)
					else
						va_av1 = RS3("VA_Av1") 
						va_av2 = RS3("VA_Av2")
						va_av3 = RS3("VA_Av3") 
						va_av4 = RS3("VA_Av4")
						va_av5 = RS3("VA_Av5")
						va_sim = ""
						va_bat = RS3("VA_Bat")						
						va_bsi = RS3("VA_Bsi")				
						va_bon = RS3("VA_Bonus")
						va_rec = RS3("VA_Rec")			
					   
					   etapa=etapa*1
					   if etapa < 3 then
							va_sim = valor_gravacao		
						else
							va_av1 = valor_gravacao				
						end if		
			
						fail = calcula_medias_M(CONOTA, curso, etapa, co_materia_fil, mat_principal, periodo, p_matric, va_av1, va_av2, va_av3, va_av4, va_av5, va_sim, va_bat, va_bsi, va_bon, va_rec)
					end if				
					
					if fail = 1 then
						response.Write("ERRO "&tb_nota)
						response.End()
					end if
					
				RS3.Movenext
				wend	
			Next
		Next
	RSAC.Movenext
	wend	

'response.End()


end function






















function calcula_medias_L(CONEXAO_BD, curso, etapa, co_materia, co_materia_pr, periodo, nu_matricula, va_t1, va_t2, va_t3, va_t4, va_p1, va_p2, va_simul, va_bat_coord, va_bon, va_rec)
fail = 0

'////////////////////////////////////////////////////////////////
'pesos (por enquanto essa verificação não é usada)
	
	if va_pt="" or isnull(va_pt) then
		va_pt = 1
		'p_va_pt="vazio"
		teste_va_pt= true
	else
		teste_va_pt = isnumeric(va_pt)
	end if
	
	if va_pp="" or isnull(va_pp) then
		if etapa <>"V" then
			etapa=etapa*1
				if etapa=3 then
					va_pp = 4
				else	
					va_pp = 2		
				end if
				'p_va_pp="vazio"
				teste_va_pp= true
		else
			va_pp = 2
			teste_va_pp= true	
		end if		
	else
		teste_va_pp = isnumeric(va_pp)
	end if
	
	
	if teste_va_pt=true and teste_va_pp=true then
		va_pt=va_pt*1
		va_pp=va_pp*1
	
	else
		fail = 1 
		erro = "pt"
		url = 0&"_"&sum_p&"_"&erro
		grava = "no"
	end if

'///////////////////////////////////////////////////////////////////////////

'TESTES
	s_va_t=0
	if va_t1="" or isnull(va_t1) then
		va_t1=NULL		
		s_va_t1=0
		soma_t1=0	
		t1_lancado="no"			
	else
		teste_va_t1 = isnumeric(va_t1)
		if teste_va_t1= true then					
			va_t1=va_t1*1			
			if va_t1 =<100 then
				IF Int(va_t1)=va_t1 THEN
					s_va_t1=1
					soma_t1=va_t1
					t1_lancado="sim"													
				ELSE	
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "t1"
						matric_Erro=i
						url = nu_matricula&"_"&va_t1&"_"&erro
						grava = "no"
					end if					
				end if																				
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "t1"
					matric_Erro=i
					url = nu_matricula&"_"&va_t1&"_"&erro
					grava = "no"
				end if					
			end if				
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "t1"
				url = nu_matricula&"_"&va_t1&"_"&erro
				grava = "no"
			end if
		end if
	end if

	if va_t2="" or isnull(va_t2) then
		va_t2=NULL		
		s_va_t2=0
		soma_t2=0	
		t2_lancado="no"			
	else
		teste_va_t2 = isnumeric(va_t2)
		if teste_va_t2= true then					
			va_t2=va_t2*1			
			if va_t2 =<100 then
				IF Int(va_t2)=va_t2 THEN
					s_va_t2=1
					soma_t2=va_t2
					t2_lancado="sim"													
				ELSE	
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "t2"
						matric_Erro=i
						url = nu_matricula&"_"&va_t2&"_"&erro
						grava = "no"
					end if					
				end if																				
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "t2"
					matric_Erro=i
					url = nu_matricula&"_"&va_t2&"_"&erro
					grava = "no"
				end if					
			end if				
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "t2"
				url = nu_matricula&"_"&va_t2&"_"&erro
				grava = "no"
			end if
		end if
	end if
	
	if va_t3="" or isnull(va_t3) then
		va_t3=NULL		
		s_va_t3=0
		soma_t3=0	
		t3_lancado="no"			
	else
		teste_va_t3 = isnumeric(va_t3)
		if teste_va_t3= true then					
			va_t3=va_t3*1			
			if va_t3 =<100 then
				IF Int(va_t3)=va_t3 THEN
					s_va_t3=1
					soma_t3=va_t3
					t3_lancado="sim"													
				ELSE	
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "t3"
						matric_Erro=i
						url = nu_matricula&"_"&va_t3&"_"&erro
						grava = "no"
					end if					
				end if																				
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "t3"
					matric_Erro=i
					url = nu_matricula&"_"&va_t3&"_"&erro
					grava = "no"
				end if					
			end if				
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "t3"
				url = nu_matricula&"_"&va_t3&"_"&erro
				grava = "no"
			end if
		end if
	end if	

	if va_t4="" or isnull(va_t4) then
		va_t4=NULL		
		s_va_t4=0
		soma_t4=0	
		t4_lancado="no"			
	else
		teste_va_t4 = isnumeric(va_t4)
		if teste_va_t4= true then					
			va_t4=va_t4*1			
			if va_t4 =<100 then
				IF Int(va_t4)=va_t4 THEN
					s_va_t4=1
					soma_t4=va_t4
					t4_lancado="sim"													
				ELSE	
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "t4"
						matric_Erro=i
						url = nu_matricula&"_"&va_t4&"_"&erro
						grava = "no"
					end if					
				end if																				
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "t4"
					matric_Erro=i
					url = nu_matricula&"_"&va_t4&"_"&erro
					grava = "no"
				end if					
			end if				
		else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "t4"
					url = nu_matricula&"_"&va_t4&"_"&erro
					grava = "no"
				end if
		end if
	end if
	
'//////////////////////////////////////////////////////////////////////
'Notas	
	s_va_p=0	
	if va_p1="" or isnull(va_p1) then
		va_p1=NULL		
		s_va_p1=0
		soma_p1=0	
		p1_lancado="no"			
	else
		teste_va_p1 = isnumeric(va_p1)
		if teste_va_p1= true then					
			va_p1=va_p1*1			
			if va_p1 =<100 then
				IF Int(va_p1)=va_p1 THEN
					s_va_p1=1
					soma_p1=va_p1
					p1_lancado="sim"													
				ELSE	
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "p1"
						matric_Erro=i
						url = nu_matricula&"_"&va_p1&"_"&erro
						grava = "no"
					end if					
				end if																				
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "p1"
					matric_Erro=i
					url = nu_matricula&"_"&va_p1&"_"&erro
					grava = "no"
				end if					
			end if				
		else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "p1"
					url = nu_matricula&"_"&va_p1&"_"&erro
					grava = "no"
				end if
		end if
	end if
'response.Write(i&"-"&nu_matricula&"-"&va_apr7 &">"& va_v_apr7&"<BR>")



	if va_p2="" or isnull(va_p2) or periodo>=4 then
		va_p2=NULL
		s_va_p2=0		
		soma_p2=0	
		p2_lancado="no"	
	else
		teste_va_p2 = isnumeric(va_p2)
		if teste_va_p2= true then					
			va_p2=va_p2*1			
			if va_p2 =<100 then
				IF Int(va_p2)=va_p2 THEN
					s_va_p2=1
					soma_p2=va_p2	
					p2_lancado="sim"					
				ELSE	
					if  fail = 1 then
						grava = "no"
					else					
						fail = 1 
						erro = "p2"
						url = nu_matricula&"_"&va_p2&"_"&erro
						grava = "no"
					end if					
				end if															
			else
				fail = 1 
				erro = "p2"
				url = nu_matricula&"_"&va_sim&"_"&erro
				grava = "no"
			end if				
		else
			fail = 1 
			erro = "p2"
			url = nu_matricula&"_"&va_sim&"_"&erro
			grava = "no"
		end if
	end if
	
			
	if va_bon="" or isnull(va_bon) then
		va_bon=NULL		
		s_va_bon=0
	else
		teste_va_bon = isnumeric(va_bon) 
		if teste_va_bon = true then
			if va_bon=<100 then
				va_bon=va_bon*1
				IF Int(va_bon)=va_bon THEN
					s_va_bon=va_bon													
				ELSE						
					if  fail = 1 then
						grava = "no"
					else												
						fail = 1 
						erro = "bon"
						url = nu_matricula&"_"&va_bon&"_"&erro
						grava = "no"
					end if					
				end if								
			else
				if  fail = 1 then
					grava = "no"
				else											
					fail = 1 
					erro = "bon"
					url = nu_matricula&"_"&va_bon&"_"&erro
					grava = "no"
				end if			
			end if
		else
			if  fail = 1 then
				grava = "no"
			else											
				fail = 1 
				erro = "bon"
				url = nu_matricula&"_"&va_bon&"_"&erro
				grava = "no"
			end if
		end if
	end if
	if va_rec="" or isnull(va_rec) then
		va_rec=NULL		
		s_va_rec=0		
	else
		periodo=periodo*1
		if periodo=1 or periodo=2 then 
	
			teste_va_rec = isnumeric(va_rec) 
			if teste_va_rec = true then
				if va_rec=<100 then
					va_rec=va_rec*1
					IF Int(va_rec)=va_rec THEN
						s_va_rec=va_rec						
					ELSE	
						if  fail = 1 then
							grava = "no"
						else					
							fail = 1 
							erro = "rec"
							url = nu_matricula&"_"&va_rec&"_"&erro
							grava = "no"
						end if					
					end if								
				else
					if  fail = 1 then
						grava = "no"
					else					
						fail = 1 
						erro = "rec"
						url = nu_matricula&"_"&va_rec&"_"&erro
						grava = "no"
					end if							
				end if
	
			else
				if  fail = 1 then
					grava = "no"
				else					
					fail = 1 
					erro = "rec"
					url = nu_matricula&"_"&va_rec&"_"&erro
					grava = "no"
				end if
			end if
		else
			if  fail = 1 then
				grava = "no"
			else					
				fail = 1 
				erro = "recp"
				url = nu_matricula&"_"&va_rec&"_"&erro
				grava = "no"
			end if
		end if			
	end if	
	
	
	if va_simul="" or isnull(va_simul) then
		simulado="no"
		va_simul=NULL		
		soma_va_simul=0
		s_simul=0					
	else
		simulado="ok"		
		soma_va_simul=va_simul
		s_simul=1	
    end if
	
	if va_bat_coord="" or isnull(va_bat_coord) then
		atualidade="no"
		va_bat_coord=NULL		
		soma_va_bat_coord=0
		s_bat=0					
	else
		atualidade="ok"		
		soma_va_bat_coord=va_bat_coord
		s_bat=1	
    end if	
'/////////////////////////////////////////////////////////////////////////
'Médias


	if grava = "ok" then

		'if sim_lancado="sim" and va_simul="S" THEN
'		if va_simul="S" THEN
'			replica_sim=ReplicaInformacoes(unidade, curso, etapa, turma, nu_matricula, periodo, "TB_Nota_K", "VA_Sim", va_sim)
'		end if
'		
'		'if bat_coord_lancado="sim" and va_rb="S" THEN
'		if va_rb="S" THEN		
'			replica_sim=ReplicaInformacoes(unidade, curso, etapa, turma, nu_matricula, periodo, "TB_Nota_K", "VA_bat_coord", va_bat_coord)
'		end if	
	
	
		soma_t1=soma_t1*1
		soma_t2=soma_t2*1
		soma_t3=soma_t3*1
		soma_t4=soma_t4*1
		soma_p1=soma_p1*1
		soma_p2=soma_p2*1
			
		s_va_t1=s_va_t1*1
		s_va_t2=s_va_t2*1
		s_va_t3=s_va_t3*1
		s_va_t4=s_va_t4*1
		s_va_p1=s_va_p1*1
		s_va_p2=s_va_p2*1	
			
		s_va_t=s_va_t1+s_va_t2+s_va_t3+s_va_t4
		s_va_p=s_va_p1+s_va_p2

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0	


		eh_disciplina_regular = disciplina_regular(co_materia, curso, etapa, CON0)
		if s_va_t>0 then
			media_teste="ok"		
			va_md_teste=(soma_t1+soma_t2+soma_t3+soma_t4)/s_va_t
			decimo = va_md_teste - Int(va_md_teste)
			If decimo >= 0.5 Then
				nota_arredondada = Int(va_md_teste) + 1
				va_md_teste=nota_arredondada
			Else
				nota_arredondada = Int(va_md_teste)
				va_md_teste=nota_arredondada					
			End If
			'mt=mt/10			
			va_md_teste = formatNumber(va_md_teste,0)
			soma_media_teste = va_md_teste	
			s_mt=1						
		else
			media_teste="no"			
			va_md_teste =NULL		
			soma_media_teste = 0
			s_mt=0						
		end if
		
		media_prova="no"			
		va_md_prova =NULL
		soma_media_prova = 0
		s_mp = 0
		
		if periodo<4 then
			'if s_va_p=2 or (eh_disciplina_regular="N" and s_va_p1=1) then
			if s_va_p1=1 then			
			    media_prova="ok"
				va_md_prova =(soma_p1+soma_p2)/s_va_p 
				decimo = va_md_prova - Int(va_md_prova)
				If decimo >= 0.5 Then
					nota_arredondada = Int(va_md_prova) + 1
					va_md_prova=nota_arredondada
				Else
					nota_arredondada = Int(va_md_prova)
					va_md_prova=nota_arredondada					
				End If
				'mt=mt/10			
				va_md_prova = formatNumber(va_md_prova,0)	
				soma_media_prova = va_md_prova		
				s_mp = 1				
			end if
		else
			if s_va_p1=1 then	
				va_md_prova = soma_p1
				soma_media_prova = va_md_prova			
				media_prova="ok"
				s_mp = 1	
			end if	
		end if
		
		m1=null		
		soma_media_teste=soma_media_teste*1
		soma_media_prova=soma_media_prova*1
		soma_va_simul = soma_va_simul*1
		s_simul=s_simul*1
		s_mt=s_mt*1
		s_mp=s_mp*1		
		'response.write(periodo&" "&soma_p1&" "&s_va_p1&" "&eh_disciplina_regular&" "&media_prova)
		'response.end()
		if eh_disciplina_regular="S" then	
			'Retirada a obrigatoriedade do simulado em 16/04/2020	
			'if (periodo<4 and media_teste="ok" and media_prova="ok" and simulado="ok") or (periodo>=4 and media_prova="ok") then
			if (periodo<4 and media_teste="ok" and media_prova="ok") or (periodo>=4 and media_prova="ok") then					

				m1=(soma_media_teste+soma_media_prova+soma_va_simul)/(s_simul+s_mt+s_mp)

				decimo = m1 - Int(m1)
				If decimo >= 0.5 Then
					nota_arredondada = Int(m1) + 1
					m1=nota_arredondada
				Else
					nota_arredondada = Int(m1)
					m1=nota_arredondada					
				End If
				'm1=m1/10	
		
				m1 = formatNumber(m1,0)	
			else
				m1=null	
			END IF
		ELSE
		'response.write(s_va_p1&"=>="&s_simul&"+"&s_mt&"+"&s_mp&"<BR>")
		   if s_va_p1>0 then
				m1=(soma_media_teste+soma_media_prova+soma_va_simul)/(s_simul+s_mt+s_mp)

				decimo = m1 - Int(m1)
				If decimo >= 0.5 Then
					nota_arredondada = Int(m1) + 1
					m1=nota_arredondada
				Else
					nota_arredondada = Int(m1)
					m1=nota_arredondada					
				End If
				'm1=m1/10	
		
				m1 = formatNumber(m1,0)	
			end if			

		END IF

		if m1>100 then
			m1=100
		end if	

		if isnull(m1) or m1="" then
			m2=NULL
			m3=NULL	
		else
			soma_va_bat_coord=soma_va_bat_coord*1					
			if isnull(va_bon) or va_bon="" then

			m2=m1+soma_va_bat_coord		
			else
				m1=m1*1		
				va_bon=va_bon*1
				m2=m1+va_bon+soma_va_bat_coord
				
				if m2>100 then
					m2=100
				end if
				'm2=m2*10
				decimo = m2 - Int(m2)
				If decimo >= 0.5 Then
					nota_arredondada = Int(m2) + 1
					m2=nota_arredondada
				Else
					nota_arredondada = Int(m2)
					m2=nota_arredondada					
				End If
				'm2=m2/10				
				m2 = formatNumber(m2,0)
			end if

		if m2>100 then
			m2=100
		end if				
		
		if va_rec="" or isnull(va_rec) then
			va_rec=NULL		
			s_va_rec=0		
'		else
'			if session("ano_letivo")>=2017 then				
'				if isnumeric(va_rec) and isnumeric(m2) then			
'					if m2>=70 then
'				
'						if  fail = 1 then
'							grava = "no"
'						else	
'											
'							fail = 1 
'							erro = "rec70"
'							url = nu_matricula&"_"&va_rec&"_"&erro
'							grava = "no"
'						end if					
'					end if	
'				end if											
'			end if
		end if
	
		if s_va_rec=0 or (m2>70 and session("ano_letivo")>=2017) then
				m3=m2						
			else					
				m2=m2*1
				va_rec=va_rec*1
				m3_temp=(m2+va_rec)/2
				
				if m3_temp > m2 then
					m3=m3_temp
	
					if m3>70 then
						m3=70
					end if
				
				else
					m3=m2
				end if
				
		
				
				'm3=m3*10
				decimo = m3 - Int(m3)
					If decimo >= 0.5 Then
						nota_arredondada = Int(m3) + 1
						m3=nota_arredondada
					Else
						nota_arredondada = Int(m3)
						m3=nota_arredondada					
					End If
				'm3=m3/10			
				m3 = formatNumber(m3,0)		
			end if
		end if

		if m3>100 then
			m3=100
		end if		

	Set RS0 = Server.CreateObject("ADODB.Recordset")
	CONEXAO0 = "Select * from TB_Nota_L WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
	Set RS0 = CONEXAO_BD.Execute(CONEXAO0)
	
	If RS0.EOF THEN	
			
		'response.Write("4"&turma &"/"&co_materia_pr)
		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Nota_L", CONEXAO_BD, 2, 2 'which table do you want open
		RS.addnew
		
			RS("CO_Matricula") = nu_matricula
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Periodo") = periodo	
			RS("VA_Teste1")=va_t1
			RS("VA_Teste2")=va_t2
			RS("VA_Teste3")=va_t3
			RS("VA_Teste4")=va_t4
			RS("MD_Teste") = va_md_teste
			RS("VA_Prova1")=va_p1
			RS("VA_Prova2")=va_p2
			RS("MD_Prova") = va_md_prova
			RS("VA_Sim")=va_simul
			RS("VA_Media1")=m1					
			RS("VA_Bat")=va_bat_coord	
			RS("VA_Bonus")=va_bon
			RS("VA_Media2")=m2
			RS("VA_Rec")=va_rec
			RS("VA_Media3")=m3
			RS("DA_Ult_Acesso") = data
			RS("HO_ult_Acesso") = horario
			RS("CO_Usuario")= co_usr
	
		RS.update
		set RS=nothing
		
	else
		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		CONEXAO0 = "DELETE * from TB_Nota_L WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
		Set RS0 = CONEXAO_BD.Execute(CONEXAO0)

		Set RS = server.createobject("adodb.recordset")		
		RS.open "TB_Nota_L", CONEXAO_BD, 2, 2 'which table do you want open
		RS.addnew
		
			RS("CO_Matricula") = nu_matricula
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Periodo") = periodo	
			RS("VA_Teste1")=va_t1
			RS("VA_Teste2")=va_t2
			RS("VA_Teste3")=va_t3
			RS("VA_Teste4")=va_t4
			RS("MD_Teste") = va_md_teste
			RS("VA_Prova1")=va_p1
			RS("VA_Prova2")=va_p2
			RS("MD_Prova") = va_md_prova
			RS("VA_Sim")=va_simul
			RS("VA_Media1")=m1					
			RS("VA_Bat")=va_bat_coord	
			RS("VA_Bonus")=va_bon
			RS("VA_Media2")=m2
			RS("VA_Rec")=va_rec
			RS("VA_Media3")=m3
			RS("DA_Ult_Acesso") = data
			RS("HO_ult_Acesso") = horario
			RS("CO_Usuario")= co_usr
		
		RS.update
		set RS=nothing	
			
	end if
	if fail = 1 then
	
	response.Redirect(endereco_origem&"notas.asp?or=01&opt=err6&hp=err_"&url&"&obr="&obr) 

	END IF	
end if

end function	























function calcula_medias_M(CONEXAO_BD, curso, etapa, co_materia, co_materia_pr, periodo, nu_matricula, va_av1, va_av2, va_av3, va_av4, va_av5, va_sim, va_bat, va_bsi, va_bon, va_rec)

fail = 0

	if isnull(va_bsi) or va_bsi="" then
		va_bsi=NULL
		soma_bsi=0
	else
		soma_bsi=va_bsi	
	end if

'////////////////////////////////////////////////////////////////
'pesos (por enquanto essa verificação não é usada)
	
	if va_pt="" or isnull(va_pt) then
		va_pt = 1
		'p_va_pt="vazio"
		teste_va_pt= true
	else
		teste_va_pt = isnumeric(va_pt)
	end if
	
	if va_pp="" or isnull(va_pp) then
		if etapa <>"V" then
			etapa=etapa*1
				if etapa=3 then
					va_pp = 4
				else	
					va_pp = 2		
				end if
				'p_va_pp="vazio"
				teste_va_pp= true
		else
			va_pp = 2
			teste_va_pp= true	
		end if		
	else
		teste_va_pp = isnumeric(va_pp)
	end if
	
	
	if teste_va_pt=true and teste_va_pp=true then
		va_pt=va_pt*1
		va_pp=va_pp*1
	
	else
		fail = 1 
		erro = "pt"
		url = 0&"_"&sum_p&"_"&erro
		grava = "no"
	end if

'///////////////////////////////////////////////////////////////////////////

'TESTES
	s_va_t=0
	if va_av1="" or isnull(va_av1) then
		va_av1=NULL		
		s_va_av1=0
		soma_av1=0	
		av1_lancado="no"			
	else
		teste_va_av1 = isnumeric(va_av1)
		if teste_va_av1= true then					
			va_av1=va_av1*1			
			if va_av1 =<100 then
				IF Int(va_av1)=va_av1 THEN
					s_va_av1=1
					soma_av1=va_av1
					av1_lancado="sim"													
				ELSE	
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "av1"
						matric_Erro=i
						url = nu_matricula&"_"&va_av1&"_"&erro
						grava = "no"
					end if					
				end if																				
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "av1"
					matric_Erro=i
					url = nu_matricula&"_"&va_av1&"_"&erro
					grava = "no"
				end if					
			end if				
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "av1"
				url = nu_matricula&"_"&va_av1&"_"&erro
				grava = "no"
			end if
		end if
	end if

	if va_av2="" or isnull(va_av2) then
		va_av2=NULL		
		s_va_av2=0
		soma_av2=0	
		av2_lancado="no"			
	else
		teste_va_av2 = isnumeric(va_av2)
		if teste_va_av2= true then					
			va_av2=va_av2*1			
			if va_av2 =<100 then
				IF Int(va_av2)=va_av2 THEN
					s_va_av2=1
					soma_av2=va_av2
					av2_lancado="sim"													
				ELSE	
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "av2"
						matric_Erro=i
						url = nu_matricula&"_"&va_av2&"_"&erro
						grava = "no"
					end if					
				end if																				
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "av2"
					matric_Erro=i
					url = nu_matricula&"_"&va_av2&"_"&erro
					grava = "no"
				end if					
			end if				
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "av2"
				url = nu_matricula&"_"&va_av2&"_"&erro
				grava = "no"
			end if
		end if
	end if
	
	if va_av3="" or isnull(va_av3) then
		va_av3=NULL		
		s_va_av3=0
		soma_av3=0	
		av3_lancado="no"			
	else
		teste_va_av3 = isnumeric(va_av3)
		if teste_va_av3= true then					
			va_av3=va_av3*1			
			if va_av3 =<100 then
				IF Int(va_av3)=va_av3 THEN
					s_va_av3=1
					soma_av3=va_av3
					av3_lancado="sim"													
				ELSE	
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "av3"
						matric_Erro=i
						url = nu_matricula&"_"&va_av3&"_"&erro
						grava = "no"
					end if					
				end if																				
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "av3"
					matric_Erro=i
					url = nu_matricula&"_"&va_av3&"_"&erro
					grava = "no"
				end if					
			end if				
		else
			if  fail = 1 then
				grava = "no"
			else
				fail = 1 
				erro = "av3"
				url = nu_matricula&"_"&va_av3&"_"&erro
				grava = "no"
			end if
		end if
	end if	

	if va_av4="" or isnull(va_av4) then
		va_av4=NULL		
		s_va_av4=0
		soma_av4=0	
		av4_lancado="no"			
	else
		teste_va_av4 = isnumeric(va_av4)
		if teste_va_av4= true then					
			va_av4=va_av4*1			
			if va_av4 =<100 then
				IF Int(va_av4)=va_av4 THEN
					s_va_av4=1
					soma_av4=va_av4
					av4_lancado="sim"													
				ELSE	
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "av4"
						matric_Erro=i
						url = nu_matricula&"_"&va_av4&"_"&erro
						grava = "no"
					end if					
				end if																				
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "av4"
					matric_Erro=i
					url = nu_matricula&"_"&va_av4&"_"&erro
					grava = "no"
				end if					
			end if				
		else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "av4"
					url = nu_matricula&"_"&va_av4&"_"&erro
					grava = "no"
				end if
		end if
	end if
	
	if va_av5="" or isnull(va_av5) then
		va_av5=NULL		
		s_va_av5=0
		soma_av5=0	
		av5_lancado="no"			
	else
		teste_va_av5 = isnumeric(va_av5)
		if teste_va_av5= true then					
			va_av5=va_av5*1			
			if va_av5 =<100 then
				IF Int(va_av5)=va_av5 THEN
					s_va_av5=1
					soma_av5=va_av5
					av5_lancado="sim"													
				ELSE	
					if  fail = 1 then
						grava = "no"
					else
						fail = 1 
						erro = "av5"
						matric_Erro=i
						url = nu_matricula&"_"&va_av5&"_"&erro
						grava = "no"
					end if					
				end if																				
			else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "av5"
					matric_Erro=i
					url = nu_matricula&"_"&va_av5&"_"&erro
					grava = "no"
				end if					
			end if				
		else
				if  fail = 1 then
					grava = "no"
				else
					fail = 1 
					erro = "av5"
					url = nu_matricula&"_"&va_av5&"_"&erro
					grava = "no"
				end if
		end if
	end if

'//////////////////////////////////////////////////////////////////////
'Notas
	s_va_p=0
	if va_sim="" or isnull(va_sim) then
		va_sim=NULL
		s_va_sim=0		
		soma_sim=0	
		sim_lancado="no"	
	else
		teste_va_sim = isnumeric(va_sim)
		if teste_va_sim= true then					
			va_sim=va_sim*1			
			if va_sim =<100 then
				IF Int(va_sim)=va_sim THEN
					s_va_sim=1
					soma_sim=va_sim	
					sim_lancado="sim"					
				ELSE	
					if  fail = 1 then
						grava = "no"
					else					
						fail = 1 
						erro = "sim"
						url = nu_matricula&"_"&va_sim&"_"&erro
						grava = "no"
					end if					
				end if															
			else
				fail = 1 
				erro = "simv"
				url = nu_matricula&"_"&va_sim&"_"&erro
				grava = "no"
			end if				
		else
			fail = 1 
			erro = "sim"
			url = nu_matricula&"_"&va_sim&"_"&erro
			grava = "no"
		end if
	end if
	
	if va_bat="" or isnull(va_bat) then
		va_bat=NULL
		s_va_bat=0		
		soma_bat=0	
		bat_lancado="no"	
	else
		teste_va_bat = isnumeric(va_bat)
		if teste_va_bat= true then					
			va_bat=va_bat*1			
			if va_bat =<5 then
				IF Int(va_bat)=va_bat THEN
					s_va_bat=1
					soma_bat=va_bat	
					bat_lancado="sim"					
				ELSE	
					if  fail = 1 then
						grava = "no"
					else					
						fail = 1 
						erro = "bat"
						url = nu_matricula&"_"&va_bat&"_"&erro
						grava = "no"
					end if					
				end if															
			else
				fail = 1 
				erro = "batv"
				url = nu_matricula&"_"&va_bat&"_"&erro
				grava = "no"
			end if				
		else
			fail = 1 
			erro = "bat"
			url = nu_matricula&"_"&va_bat&"_"&erro
			grava = "no"
		end if
	end if
		
	if va_bon="" or isnull(va_bon) then
		va_bon=NULL		
		s_va_bon=0
	else
		teste_va_bon = isnumeric(va_bon) 
		if teste_va_bon = true then
			if va_bon=<100 then
				va_bon=va_bon*1
				IF Int(va_bon)=va_bon THEN
					s_va_bon=va_bon													
				ELSE						
					if  fail = 1 then
						grava = "no"
					else												
						fail = 1 
						erro = "bon"
						url = nu_matricula&"_"&va_bon&"_"&erro
						grava = "no"
					end if					
				end if								
			else
				if  fail = 1 then
					grava = "no"
				else											
					fail = 1 
					erro = "bon"
					url = nu_matricula&"_"&va_bon&"_"&erro
					grava = "no"
				end if			
			end if
		else
			if  fail = 1 then
				grava = "no"
			else											
				fail = 1 
				erro = "bon"
				url = nu_matricula&"_"&va_bon&"_"&erro
				grava = "no"
			end if
		end if
	end if
	if va_rec="" or isnull(va_rec) then
		va_rec=NULL		
		s_va_rec=0		
	else
		periodo=periodo*1
		if periodo=1 or periodo=2 then 
	
			teste_va_rec = isnumeric(va_rec) 
			if teste_va_rec = true then
				if va_rec=<100 then
					va_rec=va_rec*1
					IF Int(va_rec)=va_rec THEN
						s_va_rec=va_rec						
					ELSE	
						if  fail = 1 then
							grava = "no"
						else					
							fail = 1 
							erro = "rec"
							url = nu_matricula&"_"&va_rec&"_"&erro
							grava = "no"
						end if					
					end if								
				else
					if  fail = 1 then
						grava = "no"
					else					
						fail = 1 
						erro = "rec"
						url = nu_matricula&"_"&va_rec&"_"&erro
						grava = "no"
					end if							
				end if
	
			else
				if  fail = 1 then
					grava = "no"
				else					
					fail = 1 
					erro = "rec"
					url = nu_matricula&"_"&va_rec&"_"&erro
					grava = "no"
				end if
			end if
		else
			if  fail = 1 then
				grava = "no"
			else					
				fail = 1 
				erro = "recp"
				url = nu_matricula&"_"&va_rec&"_"&erro
				grava = "no"
			end if
		end if			
	end if	

'/////////////////////////////////////////////////////////////////////////
'Médias


	if grava = "ok" then

	
	
		soma_av1=soma_av1*1
		soma_av2=soma_av2*1
		soma_av3=soma_av3*1
		soma_av4=soma_av4*1
		soma_av5=soma_av5*1
		soma_sim=soma_sim*1		
	
		s_va_av1=s_va_av1*1
		s_va_av2=s_va_av2*1
		s_va_av3=s_va_av3*1
		s_va_av4=s_va_av4*1
		s_va_av5=s_va_av5*1
		s_va_sim=s_va_sim*1		
		
		s_va_t=s_va_av1+s_va_av2+s_va_av3+s_va_av4+s_va_av5+s_va_sim

		if periodo>3 and av1_lancado<>"no" then
			media_av="ok"		
			mav=soma_av1	
			mav = formatNumber(mav,0)			
			va_av2=NULL	
			va_av3=NULL	
			va_av4=NULL	
			va_av5=NULL					
						
		else
			if (etapa<3 and s_va_av1=1 and s_va_av2=1) or (etapa=3 and s_va_av2=1) then	
				media_av="ok"				
				mav=(soma_av1+soma_av2+soma_av3+soma_av4+soma_av5+soma_sim)/s_va_t
			
			'mt=mt*10
				decimo = mav - Int(mav)
				If decimo >= 0.5 Then
					nota_arredondada = Int(mav) + 1
					mav=nota_arredondada
				Else
					nota_arredondada = Int(mav)
					mav=nota_arredondada					
				End If
				'mt=mt/10			
				mav = formatNumber(mav,0)	
			else
				media_av="no"
				mav=NULL						
			end if			
		end if
		
		soma_bat=soma_bat*1
		soma_bsi=soma_bsi*1
		s_va_bat=s_va_bat*1
		
		m1=NULL				
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		

		if disciplina_regular(co_materia, curso, etapa, CON0)= "S" then	
			if media_av="ok" and (sim_lancado="sim" or etapa=3 or periodo>3)  then
				mav=mav*1
				soma_bat=soma_bat*1		
				soma_bsi=soma_bsi*1
				m1=mav+soma_bat+soma_bsi
				if m1>100 then
					m1=100
				end if	
				decimo = m1 - Int(m1)
				If decimo >= 0.5 Then
					nota_arredondada = Int(m1) + 1
					m1=nota_arredondada
				Else
					nota_arredondada = Int(m1)
					m1=nota_arredondada					
				End If
				'm1=m1/10	
		
				m1 = formatNumber(m1,0)		
			END IF
		else
			if media_av="ok" then
				mav=mav*1
				soma_bat=soma_bat*1		
				soma_bsi=soma_bsi*1
				m1=mav+soma_bat+soma_bsi
				if m1>100 then
					m1=100
				end if	
				decimo = m1 - Int(m1)
				If decimo >= 0.5 Then
					nota_arredondada = Int(m1) + 1
					m1=nota_arredondada
				Else
					nota_arredondada = Int(m1)
					m1=nota_arredondada					
				End If
				'm1=m1/10	
		
				m1 = formatNumber(m1,0)
			end if			
		end if	

		if m1>100 then
			m1=100
		end if
	
		if isnull(m1) or m1="" then
			m2=NULL
			m3=NULL	
		else		
			if isnull(va_bon) or va_bon="" then
			m2=m1		
			else
				m1=m1*1		
				va_bon=va_bon*1
				m2=m1+va_bon
				
				if m2>100 then
					m2=100
				end if
				'm2=m2*10
				decimo = m2 - Int(m2)
				If decimo >= 0.5 Then
					nota_arredondada = Int(m2) + 1
					m2=nota_arredondada
				Else
					nota_arredondada = Int(m2)
					m2=nota_arredondada					
				End If
				'm2=m2/10				
				m2 = formatNumber(m2,0)
			end if

			if m2>100 then
				m2=100
			end if			
			

		if va_rec="" or isnull(va_rec) then
			va_rec=NULL		
			s_va_rec=0		
		else
'			if session("ano_letivo")>=2017 then
'				if isnumeric(va_rec) and isnumeric(m2) then
'					if m2>=70 then
'						if  fail = 1 then
'							grava = "no"
'						else					
'							fail = 1 
'							erro = "rec70"
'							url = nu_matricula&"_"&va_rec&"_"&erro
'							grava = "no"
'						end if					
'					end if	
'				end if											
'			end if
		end if
		
		if s_va_rec=0 or (m2>70 and session("ano_letivo")>=2017) then
				m3=m2						
			else					
				m2=m2*1
				if m2<70 then
					va_rec=va_rec*1
					m3_temp=(m2+va_rec)/2
					
					if m3_temp > m2 then
						m3=m3_temp
		
						if m3>70 then
							m3=70
						end if
					
					else
						m3=m2
					end if
					
			
					
					'm3=m3*10
					decimo = m3 - Int(m3)
						If decimo >= 0.5 Then
							nota_arredondada = Int(m3) + 1
							m3=nota_arredondada
						Else
							nota_arredondada = Int(m3)
							m3=nota_arredondada					
						End If
					'm3=m3/10			
					m3 = formatNumber(m3,0)		
				else
					m3=m2	
				end if	
			end if			
		end if

		if m3>100 then
			m3=100
		end if			
	

	Set RS0 = Server.CreateObject("ADODB.Recordset")
	CONEXAO0 = "Select * from TB_Nota_M WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
	Set RS0 = CONEXAO_BD.Execute(CONEXAO0)
	
	If RS0.EOF THEN	
		Set RS = server.createobject("adodb.recordset")
		
		RS.open "TB_Nota_M", CONEXAO_BD, 2, 2 'which table do you want open
		RS.addnew
	'response.Write(nu_matricula&"-"&co_materia_pr&"-"&co_materia&"-"&periodo&"<BR>")		
			RS("CO_Matricula") = nu_matricula
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Periodo") = periodo	
			RS("VA_Av1")=va_av1
			RS("VA_Av2")=va_av2
			RS("VA_Av3")=va_av3
			RS("VA_Av4")=va_av4
			RS("VA_Av5")=va_av5
			RS("VA_Mav")=mav
			RS("VA_Sim")=va_sim
			RS("VA_Bat")=va_bat	
			RS("VA_Bsi")=va_bsi	
			RS("VA_Media1")=m1
			RS("VA_Bonus")=va_bon
			RS("VA_Media2")=m2
			RS("VA_Rec")=va_rec
			RS("VA_Media3")=m3
			RS("DA_Ult_Acesso") = data
			RS("HO_ult_Acesso") = horario
			RS("CO_Usuario")= co_usr
		
		RS.update
		set RS=nothing
		
	else
		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		CONEXAO0 = "DELETE * from TB_Nota_M WHERE CO_Matricula = "& nu_matricula &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
		Set RS0 = CONEXAO_BD.Execute(CONEXAO0)

		Set RS = server.createobject("adodb.recordset")		
		RS.open "TB_Nota_M", CONEXAO_BD, 2, 2 'which table do you want open
		RS.addnew
		
			RS("CO_Matricula") = nu_matricula
			RS("CO_Materia_Principal") = co_materia_pr
			RS("CO_Materia") = co_materia
			RS("NU_Periodo") = periodo	
			RS("VA_Av1")=va_av1
			RS("VA_Av2")=va_av2
			RS("VA_Av3")=va_av3
			RS("VA_Av4")=va_av4
			RS("VA_Av5")=va_av5
			RS("VA_Mav")=mav
			RS("VA_Sim")=va_sim
			RS("VA_Bat")=va_bat	
			RS("VA_Bsi")=va_bsi				
			RS("VA_Media1")=m1
			RS("VA_Bonus")=va_bon
			RS("VA_Media2")=m2
			RS("VA_Rec")=va_rec
			RS("VA_Media3")=m3
			RS("DA_Ult_Acesso") = data
			RS("HO_ult_Acesso") = horario
			RS("CO_Usuario")= co_usr
		
		RS.update
		set RS=nothing		
	end if
	if fail = 1 then

	response.Redirect(endereco_origem&"notas.asp?or=01&opt=err6&hp=err_"&url&"&obr="&obr) 

	END IF	
end if
'response.Write(url&"<BR>")
calcula_medias_M = fail
end function		
%>