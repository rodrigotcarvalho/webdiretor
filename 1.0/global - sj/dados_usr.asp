<%
function busca_dados(ano_letivo,cod_cons,caminho_cons,caminho_contato,tipo,tp_contato)
'valores para tipo
'a=aluno
'p=professor
'f=contato
'u=usuário
'r=responsáveis

	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& caminho_cons & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	if caminho_contato<>"nulo" then
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& caminho_contato & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1	
	end if		

	if tipo="a" then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod_cons
		RS.Open SQL, CON
	
		if RS.EOF then
			resultado = "ERR1"
		else			
			nome = RS("NO_Aluno")			
			sexo = RS("IN_Sexo")
			apelido= RS("NO_Apelido")
			desteridade= RS("IN_Desteridade")
			nacionalidade= RS("CO_Nacionalidade")
			pai= RS("NO_Pai")
			mae= RS("NO_Mae")					
			pai_fal= RS("IN_Pai_Falecido")
			mae_fal= RS("IN_Mae_Falecida")
			pais= RS("CO_Pais_Natural")
			uf_natural = RS("SG_UF_Natural")
			natural = RS("CO_Municipio_Natural")
			resp_fin= RS("TP_Resp_Fin")
			resp_ped= RS("TP_Resp_Ped")
			msn= RS("TX_MSN")
			orkut= RS("TX_ORKUT")
			religiao= RS("CO_Religiao")
			raca= RS("CO_Raca")
			entrada= RS("DA_Entrada_Escola")
			cadastro= RS("DA_Cadastro")
			col_origem= RS("NO_Colegio_Origem")
			cursada= RS("NO_Serie_Cursada")
			uf_cursada= RS("SG_UF_Cursada")
			cid_cursada= RS("CO_Municipio_Cursada")
			co_estado_civil= RS("CO_Estado_Civil")			

			if isnull(nacionalidade) then 
				nacionalidade = 1
			end if

			if isnull(pais) then 
				pais = 10
			end if
			
			if isnull(uf_natural) then 
				uf_natural = "RJ"
			end if
			
			if isnull(natural) then 
				natural = 6001
			end if	
			
			if pai_fal = false then
				pai_fal = "Não"
			else
				pai_fal = "Sim"
			end if
			
			if mae_fal = false then
				mae_fal = "Não"
			else
				mae_fal = "Sim"
			end if
			
			if desteridade = "S" then
				desteridade = "Destro"
			else
				desteridade = "Canhoto"
			end if
			
			if isnull(cid_cursada) then 
				cid_cursada = 6001
			end if
			
			if isnull(uf_cursada) then 
				uf_cursada = "RJ"
			end if
						
			resultado = nome&"#!#"&sexo&"#!#"&apelido&"#!#"&desteridade&"#!#"&nacionalidade&"#!#"&pai&"#!#"&mae
			resultado = resultado&"#!#"&pai_fal&"#!#"&mae_fal&"#!#"&pais&"#!#"&uf_natural&"#!#"&natural
			resultado = resultado&"#!#"&resp_fin&"#!#"&resp_ped&"#!#"&msn&"#!#"&orkut&"#!#"&religiao&"#!#"&raca
			resultado = resultado&"#!#"&entrada&"#!#"&cadastro&"#!#"&col_origem&"#!#"&cursada&"#!#"&uf_cursada
			resultado = resultado&"#!#"&cid_cursada&"#!#"&co_estado_civil
						
		end if
				
		Set RSm = Server.CreateObject("ADODB.Recordset")
		SQLm = "SELECT * FROM TB_Matriculas WHERE CO_Matricula ="& cod_cons&" AND NU_Ano="&ano_letivo
		RSm.Open SQLm, CON		

		if RSm.EOF then
			resultado = "ERR2"
		else			
			data_rematricula = RSm("DA_Rematricula")			
			situac_matricula = RSm("CO_Situacao")
			encer_matricula= RSm("DA_Encerramento")
			unidade_matricula= RSm("NU_Unidade")
			curso_matricula= RSm("CO_Curso")
			etapa_matricula= RSm("CO_Etapa")
			turma_matricula= RSm("CO_Turma")					
			nu_cham_matricula = RSm("NU_Chamada")		

			resultado = resultado&"#!#"&data_rematricula&"#!#"&situac_matricula&"#!#"&encer_matricula&"#!#"&unidade_matricula&"#!#"&curso_matricula
			resultado = resultado&"#!#"&etapa_matricula&"#!#"&turma_matricula&"#!#"&nu_cham_matricula
						
		end if


		Set RSCONTA = Server.CreateObject("ADODB.Recordset")
		SQLA = "SELECT * FROM TB_Contatos WHERE TP_Contato ='"&tp_contato&"' And CO_Matricula ="& cod_cons
		RSCONTA.Open SQLA, CON1
	
		if RSCONTA.EOF then
			resultado = "ERR1"
		else
			nascimento = RSCONTA("DA_Nascimento_Contato")	
			vetor_nascimento = Split(nascimento,"/")  
			dia_n = vetor_nascimento(0)
			mes_n = vetor_nascimento(1)
			ano_n = vetor_nascimento(2)
		
			if dia_n<10 then 
				dia_n = "0"&dia_n
			end if
			
			if mes_n<10 then
				mes_n = "0"&mes_n
			end if
			dia_a = dia_n
			mes_a = mes_n
			ano_a = ano_n
			
			nasce = dia_n&"/"&mes_n&"/"&ano_n
			cpf= RSCONTA("CO_CPF_PFisica")
			rg= RSCONTA("CO_RG_PFisica")
			emitido= RSCONTA("CO_OERG_PFisica")
			emissao= RSCONTA("CO_DERG_PFisica")
			mail= RSCONTA("TX_EMail")
			ocupacao= RSCONTA("CO_Ocupacao")
			rua = RSCONTA("NO_Logradouro_Res")
			numero = RSCONTA("NU_Logradouro_Res")
			complemento = RSCONTA("TX_Complemento_Logradouro_Res")
			bairro= RSCONTA("CO_Bairro_Res")
			municipio= RSCONTA("CO_Municipio_Res")
			uf= RSCONTA("SG_UF_Res")
			cep = RSCONTA("CO_CEP_Res")
			telefone = RSCONTA("NU_Telefones_Res")
			tel_cont = RSCONTA("NU_Telefones")
		
			empresa= RSCONTA("NO_Empresa")
			rua2=RSCONTA("NO_Logradouro_Com")
			numero2 = RSCONTA("NU_Logradouro_Com")
			complemento2 = RSCONTA("TX_Complemento_Logradouro_Com")
			bairro2= RSCONTA("CO_Bairro_Com")
			municipio2= RSCONTA("CO_Municipio_Com")
			uf2= RSCONTA("SG_UF_Com")
			cep2 = RSCONTA("CO_CEP_Com")
			telefone2 = RSCONTA("NU_Telefones_Com")		
			
			if isnull(uf) then 
				uf = "RJ"
			end if
			
			if isnull(municipio) then 
				municipio = 6001
			end if
					
			if complemento = "nulo" or isnull(complemento)  then 
				complemento = ""
			end if
		
		
			cep5= left(cep,5)
			cep3= right(cep,3)					
			cep=cep5&"-"&cep3
		
		
			cep5= left(cep2,5)
			cep3= right(cep2,3)					
			cep2=cep5&"-"&cep3
			if tipo="a" then
				if resultado = "ERR1" then
					resultado = "ERR1"
				else
					resultado = resultado&"#!#"&nasce&"#!#"&cpf&"#!#"&rg&"#!#"&emitido&"#!#"&emissao&"#!#"&mail
					resultado = resultado&"#!#"&ocupacao&"#!#"&rua&"#!#"&numero&"#!#"&complemento&"#!#"&bairro
					resultado = resultado&"#!#"&municipio&"#!#"&uf&"#!#"&cep&"#!#"&telefone&"#!#"&tel_cont
					resultado = resultado&"#!#"&empresa&"#!#"&rua2&"#!#"&numero2&"#!#"&complemento2&"#!#"&bairro2
					resultado = resultado&"#!#"&municipio2&"#!#"&uf2&"#!#"&cep2&"#!#"&telefone2
				end if
			else
			resultado = nasce&"#!#"&cpf&"#!#"&rg&"#!#"&emitido&"#!#"&emissao&"#!#"&mail&"#!#"&ocupacao&"#!#"&rua
			resultado = resultado&"#!#"&numero&"#!#"&complemento&"#!#"&bairro&"#!#"&municipio&"#!#"&uf&"#!#"&cep
			resultado = resultado&"#!#"&telefone&"#!#"&tel_cont&"#!#"&empresa&"#!#"&rua2&"#!#"&numero2
			resultado = resultado&"#!#"&complemento2&"#!#"&bairro2&"#!#"&municipio2&"#!#"&uf2&"#!#"&cep2
			resultado = resultado&"#!#"&telefone2
			end if
		end if
		
		RSCONTA.close			
		set RSCONTA = nothing
	  
		CON1.close
		set CON1 = nothing	
		
	elseif tipo="r" then	
	
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& cod_cons
		RS.Open SQL, CON
	
		if RS.EOF then
			resultado = "ERR1"
		else
		tipo_usuario=RS("TP_Usuario")
			Set RS1 = Server.CreateObject("ADODB.Recordset")		

			if	tipo_usuario="A" then
				SQL1 = "SELECT * FROM TB_RespxAluno WHERE CO_Aluno ="& cod_cons		
				busca="CO_Usuario"
				RS1.Open SQL1, CON	
			'elseif	tipo_usuario="R" then
				'SQL1 = "SELECT * FROM TB_RespxAluno WHERE CO_Usuario ="& cod_cons	
				'busca="CO_Aluno"
				'resultado = cod_cons			
			'end if				
				
				if RS1.EOF then
					resultado = "ERR2"
				else	
					conta_resultados=0
					WHILE NOT RS1.EOF 
						result_sql=RS1(""&busca&"")
						if conta_resultados=0 then
							resultado = result_sql
						else
							resultado =resultado&"#!#"&result_sql
						end if	
						conta_resultados=conta_resultados+1
						RS1.MOVENEXT
					WEND										
				end if				
			elseif	tipo_usuario="R" then
				resultado = cod_cons			
			end if
		end if	
		RS1.close			
		set RS1 = nothing
	else
	
	
	end if

	RS.close			
	set RS = nothing
	
	CON.close
	set CON = nothing
	
busca_dados=resultado
end function
%>