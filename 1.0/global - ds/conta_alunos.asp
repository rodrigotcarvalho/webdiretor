<%Function contalunos (CAMINHO,nu_ano,unidade,curso,etapa,turma,situacao_alunos)

		
		Set CON_A = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_A.Open ABRIR
		
if situacao_alunos="TODOS" then	
	sql_situacao_alunos=""
else	
	sql_situacao_alunos="AND CO_Situacao='"&situacao_alunos&"'"
end if	

if unidade="" or isnull(unidade) then	
	sql_unidade=""
else	
	sql_unidade=" AND NU_Unidade = "& unidade
	
	if curso="" or isnull(curso) then	
		sql_curso=""
	else	
		sql_curso=" AND CO_Curso = '"& curso &"'"
		
		if etapa="" or isnull(etapa) then	
			sql_etapa=""
		else	
			sql_etapa=" AND CO_Etapa = '"& etapa &"'"

			if turma="" or isnull(turma) then	
				sql_turma=""
			else	
				sql_turma=" AND CO_Turma = '"& turma &"'"
			end if	
		end if	
	end if		
end if	
	 
				
Set RS = Server.CreateObject("ADODB.Recordset")
SQL_A = "Select COUNT(CO_Matricula) AS Qtd_Alunos from TB_Matriculas WHERE NU_Ano="&nu_ano&sql_situacao_alunos&sql_unidade&sql_curso&sql_etapa&sql_turma
Set RS = CON_A.Execute(SQL_A)

if RS.EOF then
qtd_alunos=0
else
qtd_alunos=RS("Qtd_Alunos")
end if
contalunos=qtd_alunos
end function
%>