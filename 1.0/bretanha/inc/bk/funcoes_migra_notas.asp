<!--#include file="caminhos.asp"-->
<!--#include file="atualiza_bda.asp"-->
<%
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 
data = dia &"/"& mes &"/"& ano
horario = hora & ":"& min
		
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0			
	
	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
		
	Set CONA = Server.CreateObject("ADODB.Connection") 
	ABRIRA = "DBQ="& CAMINHO_na & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONA.Open ABRIRA
	
	Set CONS = Server.CreateObject("ADODB.Connection") 
	ABRIRS = "DBQ="& CAMINHO_ns & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONS.Open ABRIRS	
		
	Set CONG = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONG.Open ABRIR
		
	Set CON_wr = Server.CreateObject("ADODB.Connection") 
	ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_wr.Open ABRIR_wr			
		
		
function atualiza_disciplina_mae(p_vetor_matricula, p_curso, p_etapa, p_co_materia_pr, p_periodo, p_dat, p_hora, p_cod_usr)
	
	atualiza_disciplina_mae = "N"

	co_matric_vetor = split(p_vetor_matricula,",")
	
	for ind_co_matric_vetor = 0 to ubound(co_matric_vetor)
		if co_matric_vetor(ind_co_matric_vetor) <> "" and not isnull(co_matric_vetor(ind_co_matric_vetor)) then
			Set RS5a = Server.CreateObject("ADODB.Recordset")
			SQL5a = "SELECT * FROM TB_Programa_Subs where CO_Etapa ='"& p_etapa &"' AND CO_Curso ='"& p_curso &"' AND CO_Materia_Principal ='"& p_co_materia_pr &"'"
			RS5a.Open SQL5a, CON0
			
			wrk_va_faltas ="" 
			wrk_va_pt="" 
			wrk_va_pp="" 
			wrk_va_t1="" 
			wrk_va_t2="" 
			wrk_va_t3="" 
			wrk_va_p1="" 
			wrk_va_p2="" 
			wrk_va_p3="" 
			wrk_va_bon="" 
			wrk_va_rec="" 	
			
			todas_mt_subs_lancadas = "S"	
			todas_mp_subs_lancadas = "S"
			acumula_faltas=0	
			conta_subs = 0	
			conta_mt = 0	
			conta_mp = 0						
			While NOT RS5a.EOF 
				co_materia_loop = RS5a("CO_Materia_Filha")		
				in_faltas = RS5a("IN_Faltas")
				po_teste = RS5a("PO_Teste")
				po_prova = RS5a("PO_Prova")				
				in_bonus = RS5a("IN_Bonus")
				rec_semestral = RS5a("IN_Rec_Semestral")	
				
				Set RSS = Server.CreateObject("ADODB.Recordset")
				CONEXAOS = "Select * from TB_Nota_S WHERE CO_Matricula ="&co_matric_vetor(ind_co_matric_vetor) &" AND CO_Materia_Principal = '"& p_co_materia_pr &"' AND CO_Materia = '"& co_materia_loop &"' AND NU_Periodo="&p_periodo&" Order by CO_Matricula, CO_Materia"	
				'response.Write(CONEXAOS&"<BR>")
				Set RSS = CONS.Execute(CONEXAOS)
				

					
				WHILE NOT RSS.EOF
					co_materia_loop = RSS("CO_Materia")				
					wrk_nu_matricula = RSS("CO_Matricula")
					wrk_va_pt = RSS("PE_Teste")				
					wrk_va_pp = RSS("PE_Prova")					
					wrk_mt = RSS("MD_Teste")	
					wrk_mp = RSS("MD_Prova")
						
					if isnumeric(wrk_mt) then
						conta_mt = conta_mt+1	
					end if					
																	
					if isnumeric(wrk_mp) then
						conta_mp = conta_mp+1
					end if					
									
					if po_teste = 1 then
						wrk_va_t1 = wrk_mt				
					elseif po_teste = 2 then
						wrk_va_t2 = wrk_mt
					elseif po_teste = 3 then				
						wrk_va_t3 = wrk_mt
					end if
					
					if po_prova = 1 then
						wrk_va_p1 = wrk_mp				
					elseif po_prova = 2 then
						wrk_va_p2 = wrk_mp
					elseif po_prova = 3 then				
						wrk_va_p3 = wrk_mp
					end if				

					if in_faltas then
						wrk_va_faltas = RSS("NU_Faltas")
						if wrk_va_faltas="" or isnull(wrk_va_faltas) then
							wrk_va_faltas=0
						end if
						acumula_faltas = acumula_faltas+wrk_va_faltas											
					end if	
					
					if in_bonus then
						wrk_va_bon = RSS("VA_Bonus")
					end if	
					
					if rec_semestral then
						wrk_va_rec = RSS("VA_Rec")			
					end if	
			
		'if wrk_nu_matricula = 20150114 then
	'response.Write(wrk_nu_matricula&"; "&p_co_materia_pr&"; "&co_materia_loop&"; "&po_teste&"; "&po_prova&"; "&wrk_va_t1&"; "&wrk_va_t2&"; "&wrk_va_t3&"; "&wrk_va_p1&"; "&wrk_va_p2&"; "&wrk_va_p3&"; "&wrk_va_faltas&"; "&wrk_va_bon&"; "&wrk_va_rec&"<BR>")
	'end if
																	
				RSS.MOVENEXT
				WEND	
			conta_subs = conta_subs+1
			RS5a.MOVENEXT
			WEND
			
			
			if conta_subs <> conta_mt then
				todas_mt_subs_lancadas = "N"	
			end if	
			
			if conta_subs <> conta_mp then
				todas_mp_subs_lancadas = "N"	
			end if	
	'response.Write(wrk_nu_matricula&" "&acumula_faltas&"<BR>")
	'response.Write(	conta_subs&" "&conta_mt&" "&conta_mp&" "&todas_mt_subs_lancadas&" "&todas_mp_subs_lancadas&"<BR>")
			if fail <> 1  then	
				gravou = Grava_BDA(wrk_nu_matricula, p_co_materia_pr, p_co_materia_pr, p_periodo, acumula_faltas, wrk_va_pt, wrk_va_pp, wrk_va_t1, wrk_va_t2, wrk_va_t3, wrk_va_p1, wrk_va_p2, wrk_va_p3, wrk_va_bon, wrk_va_rec, p_dat, p_hora, p_cod_usr, todas_mt_subs_lancadas, todas_mp_subs_lancadas)
				
				if gravou <>"S" then
					fail = 1
				END IF	
			end if	
		end if		
	NEXT	
	'response.Write("gravou "&gravou)
	'response.End()
	if fail = 1 then
		atualiza_disciplina_mae = gravou
	else
		atualiza_disciplina_mae = "S"		
	END IF		
end function	

Function comunica_disc_mae(p_unidade, p_curso, p_etapa, p_co_prof, p_co_materia_pr, p_periodo, p_nota)

	p_periodo = p_periodo*1

		Set RS5a = Server.CreateObject("ADODB.Recordset")
		SQL5a = "SELECT * FROM TB_Programa_Subs where CO_Etapa ='"& p_co_prof &"' AND CO_Curso ='"& p_curso &"' AND CO_Materia_Principal ='"& p_co_materia_pr &"''"
		RS5a.Open SQL5a, CON0
		
	bloqueia = "S"	
		
	While NOT RS5a.EOF 
	    co_materia_loop = RS5a("CO_Materia_Filha")

		consulta = "select * from TB_Da_Aula_Subs where CO_Professor="&p_co_prof&" AND NU_Unidade="&p_unidade&" AND CO_Curso='"&p_curso&"' AND CO_Etapa='"&p_etapa&"' AND CO_Materia_Principal='"&p_co_materia_pr&"' AND CO_Materia = '"&co_materia_loop&"'"
		set RS = CONG.Execute (consulta)
			
		if p_periodo=1 then		
			bloqueado = RS("ST_Per_1")
		elseif p_periodo=2 then		
			bloqueado = RS("ST_Per_2")
		elseif p_periodo=3 then		
			bloqueado = RS("ST_Per_3")						
		elseif p_periodo=4 then		
			bloqueado = RS("ST_Per_4")	
		elseif p_periodo=5 then		
			bloqueado = RS("ST_Per_5")	
		elseif p_periodo=6 then		
			bloqueado = RS("ST_Per_6")	
		end if		
		
		if bloqueado <>"x" then
			bloqueia = "N"			
		end if							
		
	RS5a.MOVENEXT
	WEND

	if bloqueia = "S" then
		escola= session("escola")
		
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		consulta1 = "select * from Email where CO_Escola="&escola
		set RS1 = CON_wr.Execute (consulta1)
		
		'response.Write(consulta1)
		
		mail_suporte=RS1("Suporte")
		mail_CC=RS1("Mail_Simplynet")
		mail_rodan=RS1("Mail_Rodan")
			
		consulta1 = "select * from TB_Usuario where CO_Usuario="&co_usr
		set RS1 = CON.Execute (consulta1)
		
		nome = RS1("NO_Usuario")
	
		consulta2a = "select * from TB_Unidade where NU_Unidade="&p_unidade
		set RS2a = CON0.Execute (consulta2a)
		
		no_unidades = RS2a("NO_Unidade")
	
		consulta2b = "select * from TB_Etapa where CO_Curso='"&p_curso&"' AND CO_Etapa='"&p_etapa&"'"
		set RS2b = CON0.Execute (consulta2b)
		
		no_serie = RS2b("NO_Etapa")
	
		consulta2c = "select * from TB_Curso where CO_Curso='"&p_curso&"'"
		set RS2c = CON0.Execute (consulta2c)
		
		no_grau = RS2c("NO_Curso")
	
		consulta = "select * from TB_Da_Aula where CO_Professor="&p_co_prof&" AND NU_Unidade="&p_unidade&" AND CO_Curso='"&p_curso&"' AND CO_Etapa='"&p_etapa&"' AND CO_Materia_Principal='"&co_materia&"'"
		set RS = CONG.Execute (consulta)
		
		coord = RS("CO_Cord")
	
		consulta_mail = "select * from TB_Usuario where CO_Usuario="&coord
		set RS_mail = CON.Execute (consulta_mail)
	
		if RS_mail.EOF then
			mail = "webdiretor@gmail.com"
			'response.write("<font class=form_corpo>Não é possível enviar a mensagem, pois o Coordenador não possui e-mail cadastrado.<br><a href=javascript:window.history.go(-1)>voltar</a></font>")
			'response.end()	
		else
			mail = RS_mail("Email_Usuario")
			'mail = mail_rodan
			if isnull(mail) or mail="" then
				response.write("<font class=form_corpo>Não é possível enviar a mensagem, pois o Coordenador codígo "&coord&" não possui e-mail cadastrado.<br><a href=javascript:window.history.go(-1)>voltar</a></font>")
				response.end()
			end if
		end if
		
		st= "x"
		p_periodo = p_periodo*1
		
		if p_periodo=1 then
			sql_atualiza= "UPDATE TB_Da_Aula SET ST_Per_1 = '"&st&"' WHERE CO_Professor="& p_co_prof &" AND CO_Materia_Principal='"& co_materia &"' AND NU_Unidade="& p_unidade &" AND CO_Curso='"& p_curso &"' AND CO_Etapa='"& p_etapa &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& p_nota &"'"
		elseif p_periodo=2 then
			sql_atualiza= "UPDATE TB_Da_Aula SET ST_Per_2 = '"&st&"' WHERE CO_Professor="& p_co_prof &" AND CO_Materia_Principal='"& co_materia &"' AND NU_Unidade="& p_unidade &" AND CO_Curso='"& p_curso &"' AND CO_Etapa='"& p_etapa &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& p_nota& "'"
		elseif p_periodo =3 then
			sql_atualiza= "UPDATE TB_Da_Aula SET ST_Per_3 = '"&st&"' WHERE CO_Professor="& p_co_prof &" AND CO_Materia_Principal='"& co_materia &"' AND NU_Unidade="& p_unidade &" AND CO_Curso='"& p_curso &"' AND CO_Etapa='"& p_etapa &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& p_nota &"'"
		elseif p_periodo =4 then
			sql_atualiza= "UPDATE TB_Da_Aula SET ST_Per_4 = '"&st&"' WHERE CO_Professor="& p_co_prof &" AND CO_Materia_Principal='"& co_materia &"' AND NU_Unidade="& p_unidade &" AND CO_Curso='"& p_curso &"' AND CO_Etapa='"& p_etapa &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& p_nota &"'"
		elseif p_periodo =5 then
			sql_atualiza= "UPDATE TB_Da_Aula SET ST_Per_5 = '"&st&"' WHERE CO_Professor="& p_co_prof &" AND CO_Materia_Principal='"& co_materia &"' AND NU_Unidade="& p_unidade &" AND CO_Curso='"& p_curso &"' AND CO_Etapa='"& p_etapa &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& p_nota &"'"
		elseif p_periodo =6 then
			sql_atualiza= "UPDATE TB_Da_Aula SET ST_Per_6 = '"&st&"' WHERE CO_Professor="& p_co_prof &" AND CO_Materia_Principal='"& co_materia &"' AND NU_Unidade="& p_unidade &" AND CO_Curso='"& p_curso &"' AND CO_Etapa='"& p_etapa &"' AND CO_Turma='"& turma &"' AND TP_Nota='"& p_nota &"'"
		end if
		
		Set RS2 = CONG.Execute(sql_atualiza)
		
		
		mensagem="O(A) Professor(a) "& nome &" lançou todas as notas de "& NO_Materia &" do "& no_serie &" do "& no_grau &", unidade: "& no_unidades &", turma "& turma&" do Periodo "& p_periodo&""
		'Dim objCDO
		Set objCDO = Server.CreateObject("CDONTS.NewMail")
		objCDO.From = mail_suporte
		objCDO.To = mail
		objCDO.CC = ""
		objCDO.BCC = mail_CC
		objCDO.Subject = "Confirmação do Lançamento de Notas através do Sistema Web Diretor"
		objCDO.Body = mensagem
		objCDO.Send()
		Set objCDO = Nothing
	End if	
	comunica_disc_mae="S"
end function
	
%>        