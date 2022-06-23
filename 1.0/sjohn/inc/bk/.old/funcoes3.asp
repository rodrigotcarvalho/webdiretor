<!--#include file="caminhos.asp"-->
<!--#include file="parametros.asp"-->
<!--#include file="funcoes6.asp"-->
<!--#include file="../../global/conta_alunos.asp"-->
<!--#include file="../../global/tabelas_escolas.asp"-->
<%Function notas (CAMINHO_al,CAMINHOn,unidade,curso,etapa,turma,co_materia,periodo,ano_letivo,co_usr,opcao,subopcao,outro)									
chave = session("chave")
session("chave")=chave
split_chave=split(chave,"-")
sistema_origem=split_chave(0)
funcao_origem=split_chave(3)

if sistema_origem="WN" then
	endereco_origem="../wn/lancar/notas/lancar/"
elseif sistema_origem="WA" then	
	if funcao_origem="EPN" then
		endereco_origem="../wa/professor/relatorio/epn/"
	else
		endereco_origem="../wa/professor/cna/notas/"
	end if
end if	

		Set CON_N = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIR3
		
		'S� usado no caso da tabela V que utiliza tamb�m a F (2013), a K(>2013 e <2016) e a M(>2015)
		'--------------------------------------------------------------------
		Set CON_NF = Server.CreateObject("ADODB.Connection")
		if ano_letivo>=2017 then
			ABRIR3F = "DBQ="& CAMINHO_nm & ";Driver={Microsoft Access Driver (*.mdb)}"		
		elseif ano_letivo>=2014 then
			ABRIR3F = "DBQ="& CAMINHO_nk & ";Driver={Microsoft Access Driver (*.mdb)}"
		else
			ABRIR3F = "DBQ="& CAMINHO_nf & ";Driver={Microsoft Access Driver (*.mdb)}"			
		end if
		CON_NF.Open ABRIR3F		
		'--------------------------------------------------------------------
			
		Set CON_wr = Server.CreateObject("ADODB.Connection") 
		ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_wr.Open ABRIR_wr
		
		Set CON_0 = Server.CreateObject("ADODB.Connection") 
		ABRIR_0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_0.Open ABRIR_0
		
		Set CON_AL = Server.CreateObject("ADODB.Connection") 
		ABRIR_AL = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_AL.Open ABRIR_AL
		
		Set CONg = Server.CreateObject("ADODB.Connection") 
		ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONg.Open ABRIRg		

linha_tabela=1
ntvmla0= 59
ntvmlb0= 59
ntvmlc0= 69
ntvmla=ntvmla0
ntvmlb=ntvmlb0
ntvmlc=ntvmlc0

ntvmla2 = formatNumber(ntvmla0,1)
ntvmlb2 = formatNumber(ntvmlb0,1)
ntvmlc2 = formatNumber(ntvmlc0,1)
qtd_alunos=contalunos(CAMINHO_al,ano_letivo,unidade,curso,etapa,turma,"C")

 		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL_0 = "Select * from TB_Materia WHERE CO_Materia = '"& co_materia &"'"
		Set RS0 = CON_0.Execute(SQL_0)

mat_princ=RS0("CO_Materia_Principal")

if mat_princ="" or isnull(mat_princ) then
	mat_princ=co_materia
end if


		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' AND CO_Materia_Principal = '"& co_materia &"'"
		Set RS = CONg.Execute(CONEXAO)
if not RS.EOF then
	ST_Per_1 = RS("ST_Per_1")
	ST_Per_2 = RS("ST_Per_2")
	ST_Per_3 = RS("ST_Per_3")
	ST_Per_4 = RS("ST_Per_4")
	ST_Per_5 = RS("ST_Per_5")
	ST_Per_6 = RS("ST_Per_6")
end if
'response.Write(">"&opcao&"<")

'0 tb&"#$#"&
'1 ln_pesos_cols&"#$#"&
'2 ln_pesos_vars&"#$#"&
'3 nm_pesos_vars&"#$#"&
'4 ln_nom_cols&"#$#"&
'5 nm_vars&"#$#"&
'6 nm_bd&"#$#"&
'7 vars_calc&"#$#"&
'8 action&"#$#"&
'9 notas_a_lancar&"#$#"&
'10 gera_pdf&"#$#"&
'11 ln_bol_av_cols&"#$#"&
'12 ln_bol_av_span&"#$#"&
'13 nm_bol_av_vars&"#$#"&
'14 ln_bol_av_vars&"#$#"&
'15 vars_bol_av_calc&"#$#"&
'16 legenda_bol_av&"#$#"&
'17 exibe_apr_pr_bol_av

dados_tabela=verifica_dados_tabela(CAMINHOn,opcao,outro)
	dados_separados=split(dados_tabela,"#$#")
	tb=dados_separados(0)
	ln_pesos_cols=dados_separados(1)
	ln_pesos_vars=dados_separados(2)
	nm_pesos_vars=dados_separados(3)
	ln_nom_cols=dados_separados(4)
	nm_vars=dados_separados(5)
	nm_bd=dados_separados(6)
	action=dados_separados(8)
	notas_a_lancar=dados_separados(9)

	linha_pesos=split(ln_pesos_cols,"#!#")
	linha_pesos_variaveis=split(ln_pesos_vars,"#!#")
	nome_pesos_variaveis=split(nm_pesos_vars,"#!#")
	linha_nome_colunas=split(ln_nom_cols,"#!#")
	nome_variaveis=split(nm_vars,"#!#")
	variaveis_bd=split(nm_bd,"#!#")

if subopcao="cln" then
	comunica="s" 
	opt="?opt=cln&obr="&obr&"&nota="&tb
	tipo="hidden"

elseif subopcao="imp" then
	comunica="s" 
	opt=""
	tipo="hidden"
elseif subopcao="blq" then
	comunica="s" 
	opt=""
	tipo="hidden"
else
	comunica="n" 
	opt=""
	tipo="text"
end if

if subopcao="imp" Then
	classe_peso = "tabelaTit"
	classe_subtit = "tabelaTit"

elseif errou="pt" or errou="pp" Then
	classe_peso = "tb_fundo_linha_erro"
else
	classe_peso = "tb_fundo_linha_peso"
	classe_subtit = "tb_subtit"
end if


qtd_colunas=UBOUND(linha_nome_colunas)+1
width_num_cham="20"
width_nom_aluno="340"
width_else=(1000-20-340)/(qtd_colunas-2)

%>
<form action="<%response.Write(action&opt)%>" name="nota" method="post" onSubmit="return checksubmit()">
<table width="1000" border="0" cellspacing="0" cellpadding="0">
<%if ubound(linha_pesos)>-1 then %>
  <tr> 
 <% 
 for i= 0 to ubound(linha_pesos)
 		if i=0 then
			width=width_num_cham
			align="center"
		elseif i=1 then	
			width=width_nom_aluno
			align="left"			
		else
			width=width_else
			align="center"
		end if		

		if linha_pesos(i)="PESO" then
			linha_pesos_variaveis=split(ln_pesos_vars,"#!#")
			nome_pesos_variaveis=split(nm_pesos_vars,"#!#")		

			Set RStemp = Server.CreateObject("ADODB.Recordset")
			SQL_temp = "Select * from TB_Matriculas WHERE NU_Ano= "& ano_letivo &" AND CO_Situacao='C' AND NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
			Set RStemp = CON_AL.Execute(SQL_temp)
			
						
				nu_matriculatemp = RStemp("CO_Matricula")
			
				Set RSpeso = Server.CreateObject("ADODB.Recordset")
				SQL_peso = "Select "&linha_pesos_variaveis(i)&" from "& tb &" WHERE CO_Matricula = "& nu_matriculatemp & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				Set RSpeso = CON_N.Execute(SQL_peso)
					
				if RSpeso.EOF then
					if tb= "TB_NOTA_F" then
						if periodo<=4 then
							if linha_pesos_variaveis(i) = "PE_Teste" then
								valor_peso = 1
							elseif linha_pesos_variaveis(i) = "PE_Prova1" then
								valor_peso = 4
							else
								if periodo<4 then
									valor_peso = 5	
								else
						   			valor_peso =""										
								end if						 
							end if
						 else
						   valor_peso =""	
						 end if	
					end if				
				else	
					valor_peso=RSpeso(""&linha_pesos_variaveis(i)&"")
					if (valor_peso = "" or isnull(valor_peso)) and tb= "TB_NOTA_F" then
						if periodo<=4 then
							if linha_pesos_variaveis(i) = "PE_Teste" then
								valor_peso = 1
							elseif linha_pesos_variaveis(i) = "PE_Prova1" then
								valor_peso = 4
							else
								if periodo<4 then
									valor_peso = 5	
								else
						   			valor_peso =""										
								end if						 
							end if
						 else
						   valor_peso =""	
						 end if	
					end if
				end if		
				IF comunica="s" THEN		
					linha_pesos(i)=valor_peso&"<input name="&nome_pesos_variaveis(i)&" type=""hidden"" id="&nome_pesos_variaveis(i)&" class=""peso"" value="&valor_peso&">"	
				else	
					linha_pesos(i)="<input name="&nome_pesos_variaveis(i)&" type="&tipo&" id="&nome_pesos_variaveis(i)&" class=""peso"" value="&valor_peso&">"
				end if	
			
		end if				
 %>
    <td width="<%response.Write(width)%>" class="<%response.Write(classe_peso)%>"><div align="<%response.Write(align)%>"><%response.Write(linha_pesos(i))%></div></td>
<%	next%>
</tr>
<%end if%>
  <tr> 
 <% for j= 0 to ubound(linha_nome_colunas)
 		if j=0 then
			width=width_num_cham
			align="center"
		elseif j=1 then	
			width=width_nom_aluno
			align="left"			
		else
			width=width_else
			align="center"
		end if							
 %>
    <td width="<%response.Write(width)%>" class="<%response.Write(classe_subtit)%>"><div align="<%response.Write(align)%>"><%response.Write(linha_nome_colunas(j))%></div></td>
<%	next%>  
  </tr>
  <%
check = 2
nu_chamada_ckq = 0

Set RS = Server.CreateObject("ADODB.Recordset")
SQL_A = "Select * from TB_Matriculas WHERE NU_Ano="&ano_letivo&" AND NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
Set RS = CON_AL.Execute(SQL_A)


While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")
nu_matricula=nu_matricula*1
if isnumeric(calc) then
calc=calc*1
end if
	
	if subopcao="imp" Then
		classe = "tabela"
		classe_td_imp= " class = 'tabela'"
	elseif nu_matricula = calc then
		classe = "tb_fundo_linha_erro"
		onblur="mudar_cor_blur_erro"	
		classe_td_imp= ""	  	   
	else
		if check mod 2 =0 then
			classe = "tb_fundo_linha_par" 
			onblur="mudar_cor_blur_par"
		else 
			classe ="tb_fundo_linha_impar"
			onblur="mudar_cor_blur_impar"
		end if 
		classe_td_imp= ""		
	end if

	Set RSs = Server.CreateObject("ADODB.Recordset")
	SQL_s ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& nu_matricula&" and TB_Matriculas.NU_Ano="&ano_letivo
	Set RSs = CON_AL.Execute(SQL_s)

	if RSs.EOF then
	%>
			<tr>
					<td width="20" class="<%response.Write(classe)%>">&nbsp;
					</td>
					<td width="340" class="<%response.Write(classe)%>">Matr�cula <%response.Write(nu_matricula)%> cadastrada em TB_Matriculas sem correspond�ncia em TB_Alunos</td>               
						
						<%for m= 0 to ubound(nome_variaveis)
							width=width_else
							align="center"
					 %>
							<td width="<%response.Write(width)%>" class="<%response.Write(classe)%>">&nbsp;</td>
				 		 <%next%>
				  </tr>
	<%else
		situac=RSs("CO_Situacao")
		nome_aluno=RSs("NO_Aluno")	
	'Verificando se algum aluno mudou de turma e inserindo uma linha cinza para o lugar do aluno
			if (nu_chamada_ckq <>nu_chamada - 1) then
				teste_nu_chamada = nu_chamada-nu_chamada_ckq
				teste_nu_chamada=teste_nu_chamada-1
				
				classe_anterior=classe
				if subopcao="imp" Then
					classe = "tabela"
				else	
					classe="tb_fundo_linha_falta"
				end if
		
				for k=1 to teste_nu_chamada 
					nu_chamada_ckq=nu_chamada_ckq+1
					nu_chamada_falta=nu_chamada_ckq
				%>
				  <tr>
					<td width="20" class="<%response.Write(classe)%>">
					  <input name="nu_chamada_<%response.Write(nu_chamada_falta)%>" class="borda_edit" type="hidden" value="<%response.Write(nu_chamada_falta)%>"> 
					  <%response.Write(nu_chamada_falta)%>
					  <input name="nu_matricula_<%response.Write(nu_chamada_falta)%>" type="hidden" value="falta"> 
					</td>
					<td width="340" class="<%response.Write(classe)%>">&nbsp;</td>               
						
						<%for m= 0 to ubound(nome_variaveis)
							width=width_else
							align="center"
							nome_campo=nome_variaveis(m)&"_"&nu_chamada
							conteudo="&nbsp;"
					 %>
							<td width="<%response.Write(width)%>" class="<%response.Write(classe)%>">
								<div align="<%response.Write(align)%>">
									<%response.Write(conteudo)%> 
								</div>
					</td>
				 		 <%next%>
				  </tr>
					<%
					next
	'Inserindo o aluno seguinte aos que mudaram de turma
					nu_chamada_ckq=nu_chamada_ckq+1				
					inativo="N"
					if situac<>"C" then
						if subopcao="imp" Then
							classe = "tabela"
						else
							classe="tb_fundo_linha_falta"
						end if	
							valor="falta"
							nome_aluno=nome_aluno&" - Aluno Inativo"
							inativo="S"
					end if			
					%>   
			<tr class="<%response.Write(classe_anterior)%>" id="<%response.Write("celula"&nu_chamada)%>">
			<td width="20" <%response.Write(classe_td_imp)%>>
				<input name="nu_chamada_<%response.Write(nu_chamada)%>" class="borda_edit" type="hidden" value="<%response.Write(nu_chamada)%>"> 
				<%response.Write(nu_chamada)%>
				<input name="nu_matricula_<%response.Write(nu_chamada)%>" type="hidden" value='<%response.Write(nu_matricula)%>'> 
			</td>
			<td width="340" <%response.Write(classe_td_imp)%>><%response.Write(nome_aluno)%></td>               
				 <% 
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				Set RS3 = CON_N.Execute(SQL_N)			 
				coluna=0	 
				 for n= 0 to ubound(nome_variaveis)
					width=width_else
					align="center"
					nome_campo=nome_variaveis(n)&"_"&nu_chamada
				
					if RS3.EOF then 
						valor=""
					else
						nu_matricula=nu_matricula*1
						if isnumeric(calc) then
						calc=calc*1
						end if					
						if opt="err6" and nu_matricula = calc then	
							if errou=nome_variaveis(n) then
								valor=qerrou	
							else
								valor=Session(nome_variaveis(n))
							end if
						else	
							if tb = "TB_NOTA_V" and (variaveis_bd(n) = "VA_Media1" or variaveis_bd(n) = "VA_Bonus" or variaveis_bd(n) = "VA_Media2" or variaveis_bd(n) = "VA_Rec" or variaveis_bd(n) = "VA_Media3") then
								Set RS3F = Server.CreateObject("ADODB.Recordset")
								ano_letivo=ano_letivo*1
								if ano_letivo=2013 then
									SQL_NF = "Select * from TB_NOTA_F WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
								elseif ano_letivo<2017 then
									SQL_NF = "Select * from TB_NOTA_K WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
								else
									SQL_NF = "Select * from TB_NOTA_M WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo									
								end if										
								Set RS3F = CON_NF.Execute(SQL_NF)	
								if RS3F.eof then
									valor=""
								else
									valor=RS3F(""&variaveis_bd(n)&"")															
								end if
								
							elseif nome_variaveis(n) <> "rs" and nome_variaveis(n) <> "rb" then  															
								valor=RS3(""&variaveis_bd(n)&"")
							end if	
						end if							
					end if
					
					if (valor="" or isnull(valor)) and subopcao="imp" then
						coluna=coluna+1	
						conteudo="&nbsp;"			
					else
						if (nome_variaveis(n) = "rs" or nome_variaveis(n)="rb") and inativo="N" then
							tipo_form = "checkbox"
							valor = "S"
						else
							tipo_form = tipo
						end if
					
						if nome_variaveis(n)="media_teste" or nome_variaveis(n)="media_prova" or nome_variaveis(n)="media1" or nome_variaveis(n)="media2" or nome_variaveis(n)="media3" then
							coluna=coluna	
							conteudo=valor
						elseif nome_variaveis(n)="simul_coord"  or nome_variaveis(n)="bsi_coord" or nome_variaveis(n)="bat_coord" or situac<>"C" then
							coluna=coluna	
							conteudo=valor&"<input name='"&nome_campo&"' type='hidden' id='"&linha_tabela&"c"&coluna&"' value='"&valor&"'>"	
						else
							coluna=coluna+1
							if comunica="s" or subopcao="blq" then
							
								if nome_variaveis(n) = "rs" or nome_variaveis(n)="rb" then
									conteudo="&nbsp;"
								else
	'						if comunica="s" or (periodo_bloqueado="s" and sistema_origem="WN") then
									conteudo=valor&"<input name='"&nome_campo&"' type='"&tipo_form&"' id='"&linha_tabela&"c"&coluna&"' onFocus=""mudar_cor_focus(celula"&nu_chamada&");javascript:this.form."&nome_campo&".select();"" onBlur="&onblur&"(celula"&nu_chamada&") value='"&valor&"' class=""nota"" size=""4"" maxlength=""3"" onkeydown=""keyPressed(this.id,event,"&notas_a_lancar&","&qtd_alunos&")"">"	
								end if	
							else
								conteudo="<input name='"&nome_campo&"' type='"&tipo_form&"' id='"&linha_tabela&"c"&coluna&"' onFocus=""mudar_cor_focus(celula"&nu_chamada&");javascript:this.form."&nome_campo&".select();"" onBlur="&onblur&"(celula"&nu_chamada&") value='"&valor&"' class=""nota"" size=""4"" maxlength=""3"" onkeydown=""keyPressed(this.id,event,"&notas_a_lancar&","&qtd_alunos&")"">"										
							end if
						end if	
					end if	
					'conteudo=n
			 %>
				<td width="<%response.Write(width)%>" <%response.Write(classe_td_imp)%>>
					<div align="<%response.Write(align)%>">
						<%response.Write(conteudo)%> 
					</div>
				 </td>
			  <%	next  
			  %>
			  </tr>              
				<%
	'Se os n�meros de chamada estiverem completos. Se n�o faltar aluno na turma.
			ELSE	
					inativo="N"
					if situac<>"C" then
						if subopcao="imp" Then
							classe = "tabela"
						else
							classe="tb_fundo_linha_falta"
						end if	
							valor="falta"
							nome_aluno=nome_aluno&" - Aluno Inativo"
							inativo="S"
					end if			
					nu_chamada_ckq=nu_chamada_ckq+1
					%>   
			<tr class="<%response.Write(classe)%>" id="<%response.Write("celula"&nu_chamada)%>">
			<td width="20" <%response.Write(classe_td_imp)%>>
				<input name="nu_chamada_<%response.Write(nu_chamada)%>" class="borda_edit" type="hidden" value="<%response.Write(nu_chamada)%>"> 
				<%response.Write(nu_chamada)%>
				<input name="nu_matricula_<%response.Write(nu_chamada)%>" type="hidden" value='<%response.Write(nu_matricula)%>'> 
			</td>
			<td width="340" <%response.Write(classe_td_imp)%>><%response.Write(nome_aluno)%></td>               
				 <% 
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				Set RS3 = CON_N.Execute(SQL_N)			 
				coluna=0	 
				 for n= 0 to ubound(nome_variaveis)
					width=width_else
					align="center"
					nome_campo=nome_variaveis(n)&"_"&nu_chamada
				
					if RS3.EOF then 
						valor=""
					else
						nu_matricula=nu_matricula*1
						if isnumeric(calc) then
						calc=calc*1
						end if					
						if opt="err6" and nu_matricula = calc then	
							if errou=nome_variaveis(n) then
								valor=qerrou	
							else
								valor=Session(nome_variaveis(n))
							end if
						else	
							if tb = "TB_NOTA_V" and (variaveis_bd(n) = "VA_Media1" or variaveis_bd(n) = "VA_Bonus" or variaveis_bd(n) = "VA_Media2" or variaveis_bd(n) = "VA_Rec" or variaveis_bd(n) = "VA_Media3") then
								Set RS3F = Server.CreateObject("ADODB.Recordset")
								ano_letivo=ano_letivo*1
								if ano_letivo=2013 then
									SQL_NF = "Select * from TB_NOTA_F WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
								elseif ano_letivo<2017 then
									SQL_NF = "Select * from TB_NOTA_K WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
								else
									SQL_NF = "Select * from TB_NOTA_M WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo								
								end if										
								Set RS3F = CON_NF.Execute(SQL_NF)	
								if RS3F.eof then
									valor=""
								else
									valor=RS3F(""&variaveis_bd(n)&"")													
								end if								

							elseif nome_variaveis(n) <> "rs" and nome_variaveis(n) <> "rb"  then  										
								valor=RS3(""&variaveis_bd(n)&"")
							end if	
						end if							
					end if
					
					if (valor="" or isnull(valor)) and subopcao="imp" then
						coluna=coluna+1	
						conteudo="&nbsp;"			
					else
						if (nome_variaveis(n) = "rs" or nome_variaveis(n)="rb") and inativo="N" then
							tipo_form = "checkbox"
							valor="S"
						else
							tipo_form = tipo
						end if
					
					
						if nome_variaveis(n)="media_teste" or nome_variaveis(n)="media_prova" or nome_variaveis(n)="media1" or nome_variaveis(n)="media2" or nome_variaveis(n)="media3" then
							coluna=coluna	
							conteudo=valor
						elseif nome_variaveis(n)="simul_coord" or nome_variaveis(n)="bsi_coord" or nome_variaveis(n)="bat_coord" or situac<>"C" then
							coluna=coluna	
							conteudo=valor&"<input name='"&nome_campo&"' type='hidden' id='"&linha_tabela&"c"&coluna&"' value='"&valor&"'>"	
						else
							coluna=coluna+1
							if comunica="s" or subopcao="blq" then
	'						if comunica="s" or (periodo_bloqueado="s" and sistema_origem="WN") then
								if nome_variaveis(n) = "rs" or nome_variaveis(n)="rb" then
									conteudo="&nbsp;"
								else
	'						if comunica="s" or (periodo_bloqueado="s" and sistema_origem="WN") then
									conteudo=valor&"<input name='"&nome_campo&"' type='"&tipo_form&"' id='"&linha_tabela&"c"&coluna&"' onFocus=""mudar_cor_focus(celula"&nu_chamada&");javascript:this.form."&nome_campo&".select();"" onBlur="&onblur&"(celula"&nu_chamada&") value='"&valor&"' class=""nota"" size=""4"" maxlength=""3"" onkeydown=""keyPressed(this.id,event,"&notas_a_lancar&","&qtd_alunos&")"">"	
								end if		
							else
								conteudo="<input name='"&nome_campo&"' type='"&tipo_form&"' id='"&linha_tabela&"c"&coluna&"' onFocus=""mudar_cor_focus(celula"&nu_chamada&");javascript:this.form."&nome_campo&".select();"" onBlur="&onblur&"(celula"&nu_chamada&") value='"&valor&"' class=""nota"" size=""4"" maxlength=""3"" onkeydown=""keyPressed(this.id,event,"&notas_a_lancar&","&qtd_alunos&")"">"										
							end if
						end if	
					end if	
					'conteudo=n
			 %>
				<td width="<%response.Write(width)%>" <%response.Write(classe_td_imp)%>>
					<div align="<%response.Write(align)%>">
						<%response.Write(conteudo)%> 
					</div>
				 </td>
			  <%	next  
			  %>
			  </tr>
			  <%			
			END IF			              
		if situac<>"C" then
			linha_tabela=linha_tabela
		else
		
			linha_tabela=linha_tabela+1
		end if
 	
	END IF	
max=nu_chamada
	check = check+1 
RS.MoveNext
Wend 
session("max")=max

if subopcao="imp" then
else
%>
    <tr> 
      <td colspan="<%response.Write(qtd_colunas)%>" class="tb_subtit_lanca_notas">
     <%	  
	if funcao_origem="EPN" or subopcao="blq" then
	%>
				<table width="100%" border="0" cellspacing="0">
		  <tr>
			<td colspan="3">
			  <hr>
			</td>
		  </tr>	 
			<tr> 
			<td width="33%">
			<div align="center">
				<input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','altera.asp');return document.MM_returnValue" value="Voltar">
			  </div>
			  </td>
			  <td width="34%"> <div align="center">
				</div></td>
			  <td width="33%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
	
				  </font></div></td>
			</tr>
		  </table>
	<%
'	elseif  periodo_bloqueado="s" and sistema_origem="WN" then
'	 %>
<!--			<table width="100%" border="0" cellspacing="0">
		  <tr>
			<td colspan="3">
			  <hr>
			</td>
		  </tr>	 
			<tr> 
			<td width="33%">
			<div align="center">
				<input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','altera.asp');return document.MM_returnValue" value="Voltar">
			  </div>
			  </td>
			  <td width="34%"> <div align="center">
				</div></td>
			  <td width="33%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
	
				  </font></div></td>
			</tr>
		  </table>
-->	<%elseif comunica="s" then%>
		 <table width="100%" border="0" cellspacing="0">
		  <tr>
			<td colspan="3">
			  <hr>
			</td>
		  </tr>	 
			<tr> 
			<td width="33%"><div align="center">
							<input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','notas.asp?d=<%response.Write(co_materia)%>&pr=<%response.Write(co_prof)%>&p=<%response.Write(periodo)%>');return document.MM_returnValue" value="Voltar">
						  </div></td>
			  <td width="34%"> <div align="center"> 
				  <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?ori=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" value="Cancelar">
				</div></td>
			  <td width="33%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
				  <input type="submit" name="Submit" value="Confirmar" class="botao_prosseguir">
				  <input name="unidade" type="hidden" id="unidade" value="<%=unidades%>">
				  <input name="curso" type="hidden" id="curso" value="<%=grau%>">
				  <input name="etapa" type="hidden" id="etapa" value="<%=serie%>">
				  <input name="turma" type="hidden" id="turma" value="<%=turma%>">
				  <input name="co_materia" type="hidden" id="co_materia" value="<%= co_materia%>">
				  <input name="periodo" type="hidden" id="periodo" value="<%= periodo%>">
				  <input name="co_prof" type="hidden" id="co_prof" value="<% = co_prof%>">
				  <input name="max" type="hidden" id="max" value="<% =max%>">
				  <input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>">
				  <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>">
				  </font></div></td>
			</tr>
		  </table>
	<%else%>            
		  <table width="100%" border="0" align="center" cellspacing="0">
			  <tr> 
				<td colspan="3"><hr></td>
			  </tr>
			  <tr> 
				<td width="33%"><div align="center">             
					<input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','altera.asp');return document.MM_returnValue" value="Voltar">
				  </div></td>
				<td width="34%"><div align="center"> 
					<!--<input name="Submit" type="button" class="botao_prosseguir_comunicar" onClick="MM_goToURL('parent','notas.asp?or=01&opt=cln&obr=<%=obr%>');return document.MM_returnValue" value="Comunicar ao Coordenador T&eacute;rmino da Planilha">-->&nbsp;
				  </div></td>
				<td width="33%"> <div align="center"> 
					<input type="submit" name="Submit2" value="Salvar" class="botao_prosseguir">
					<input name="unidade" type="hidden" id="unidade" value="<%=unidade%>">
					<input name="curso" type="hidden" id="curso" value="<%=curso%>">
					<input name="etapa" type="hidden" id="etapa" value="<%=etapa%>">
					<input name="turma" type="hidden" id="turma" value="<%=turma%>">
					<input name="co_materia" type="hidden" id="co_materia" value="<%= co_materia%>">
					<input name="periodo" type="hidden" id="periodo" value="<%= periodo%>">
					<input name="co_prof" type="hidden" id="co_prof" value="<% = co_prof%>">
					<input name="max" type="hidden" id="max" value="<% =max%>">
					<input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>">
					<input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>">
				  </div></td>
			  </tr>
			</table>
	<%end if
end if%>        
</td>
</tr>
</table>
</form>        
<%end function

Function bonus_e_simulados (unidade,curso,etapa,turma,co_materia,periodo,ano_letivo,co_usr,opcao,subopcao,outro)									
chave = session("chave")
session("chave")=chave
split_chave=split(chave,"-")
sistema_origem=split_chave(0)
funcao_origem=split_chave(3)

url_retorno = "index.asp?nvg="&chave

if sistema_origem="WN" then
	endereco_origem="../wn/lancar/notas/lancar/"
elseif sistema_origem="WA" then	
	if funcao_origem="EPN" then
		endereco_origem="../wa/professor/relatorio/epn/"
	else
		endereco_origem="../wa/professor/cna/notas/"
	end if
end if	


		
		'S� usado no caso da tabela V que utiliza tamb�m a F (2013) e a K(>2013)
		'--------------------------------------------------------------------
		Set CON_NF = Server.CreateObject("ADODB.Connection")
		if ano_letivo>=2017 then
			ABRIR3F = "DBQ="& CAMINHO_nm & ";Driver={Microsoft Access Driver (*.mdb)}"		
		elseif ano_letivo>=2014 then
			ABRIR3F = "DBQ="& CAMINHO_nk & ";Driver={Microsoft Access Driver (*.mdb)}"
		else
			ABRIR3F = "DBQ="& CAMINHO_nf & ";Driver={Microsoft Access Driver (*.mdb)}"			
		end if
		CON_NF.Open ABRIR3F		
		'--------------------------------------------------------------------
			
		Set CON_wr = Server.CreateObject("ADODB.Connection") 
		ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_wr.Open ABRIR_wr
		
		Set CON_0 = Server.CreateObject("ADODB.Connection") 
		ABRIR_0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_0.Open ABRIR_0
		
		Set CON_AL = Server.CreateObject("ADODB.Connection") 
		ABRIR_AL = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_AL.Open ABRIR_AL
		
		Set CONg = Server.CreateObject("ADODB.Connection") 
		ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONg.Open ABRIRg	
		
tb_nota = tabela_notas(CONg, unidade, curso, etapa, turma, periodo, co_materia, outro)

if funcao_origem <> "LBA"then
	DisciplinaEhRegular = "S"
else

    Set RSder = Server.CreateObject("ADODB.Recordset")
	SQLder = "SELECT * FROM TB_Programa_Aula WHERE CO_Curso='"& curso &"' AND CO_Etapa = '"&etapa&"' and CO_Materia='"& co_materia&"' and TP_Disciplina = 'R' "	
	RSder.Open SQLder, CON_0	
	
	if RSder.EOF then
		DisciplinaEhRegular = "N"
	else
		DisciplinaEhRegular = "S"	
	end if
end if
'if (UCASE(tb_nota)=UCASE("TB_Nota_L") or UCASE(tb_nota)=UCASE("TB_Nota_M" ) ) and DisciplinaEhRegular = "S" then
'Retirado em 06/06/2017 #116 - Sistema Trimestral - Função NOVA - Lançar Bônus de Atualidade Por Disciplina  histórico de 29/05/2017
if (UCASE(tb_nota)=UCASE("TB_Nota_L") or UCASE(tb_nota)=UCASE("TB_Nota_M" ) )  then



CAMINHOn = caminho_notas(CONg, tb_nota, outro)	


linha_tabela=1
ntvmla0= 59
ntvmlb0= 59
ntvmlc0= 69
ntvmla=ntvmla0
ntvmlb=ntvmlb0
ntvmlc=ntvmlc0

ntvmla2 = formatNumber(ntvmla0,1)
ntvmlb2 = formatNumber(ntvmlb0,1)
ntvmlc2 = formatNumber(ntvmlc0,1)
qtd_alunos=contalunos(CAMINHO_al,ano_letivo,unidade,curso,etapa,turma,"C")

if not (co_materia="" or isnull(co_materia)) then
 		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL_0 = "Select * from TB_Materia WHERE CO_Materia = '"& co_materia &"'"
		Set RS0 = CON_0.Execute(SQL_0)

	mat_princ=RS0("CO_Materia_Principal")
	
	if mat_princ="" or isnull(mat_princ) then
		mat_princ=co_materia
	end if



		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' AND CO_Materia_Principal = '"& co_materia &"'"
		Set RS = CONg.Execute(CONEXAO)

if not RS.EOF then
	ST_Per_1 = RS("ST_Per_1")
	ST_Per_2 = RS("ST_Per_2")
	ST_Per_3 = RS("ST_Per_3")
	ST_Per_4 = RS("ST_Per_4")
	ST_Per_5 = RS("ST_Per_5")
	ST_Per_6 = RS("ST_Per_6")
end if	
	
	'response.Write(">"&opcao&"<")
	
	'0 tb&"#$#"&
	'1 ln_pesos_cols&"#$#"&
	'2 ln_pesos_vars&"#$#"&
	'3 nm_pesos_vars&"#$#"&
	'4 ln_nom_cols&"#$#"&
	'5 nm_vars&"#$#"&
	'6 nm_bd&"#$#"&
	'7 vars_calc&"#$#"&
	'8 action&"#$#"&
	'9 notas_a_lancar&"#$#"&
	'10 gera_pdf&"#$#"&
	'11 ln_bol_av_cols&"#$#"&
	'12 ln_bol_av_span&"#$#"&
	'13 nm_bol_av_vars&"#$#"&
	'14 ln_bol_av_vars&"#$#"&
	'15 vars_bol_av_calc&"#$#"&
	'16 legenda_bol_av&"#$#"&
	'17 exibe_apr_pr_bol_av

end if

dados_tabela=verifica_dados_tabela(CAMINHOn,opcao,outro)
	dados_separados=split(dados_tabela,"#$#")
	tb=dados_separados(0)
	ln_pesos_cols=dados_separados(1)
	ln_pesos_vars=dados_separados(2)
	nm_pesos_vars=dados_separados(3)
	ln_nom_cols=dados_separados(4)
	nm_vars=dados_separados(5)
	nm_bd=dados_separados(6)
	action=dados_separados(8)
	notas_a_lancar=dados_separados(9)
	

		Set CON_N = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIR3
	

	linha_pesos=split(ln_pesos_cols,"#!#")
	linha_pesos_variaveis=split(ln_pesos_vars,"#!#")
	nome_pesos_variaveis=split(nm_pesos_vars,"#!#")
	linha_nome_colunas=split(ln_nom_cols,"#!#")
	nome_variaveis=split(nm_vars,"#!#")
	variaveis_bd=split(nm_bd,"#!#")

if subopcao="cln" then
	comunica="s" 
	opt="?opt=cln&obr="&obr&"&nota="&tb
	tipo="hidden"

elseif subopcao="imp" then
	comunica="s" 
	opt=""
	tipo="hidden"
elseif subopcao="blq" then
	comunica="s" 
	opt=""
	tipo="hidden"
else
	comunica="n" 
	opt=""
	tipo="text"
end if

if subopcao="imp" Then
	classe_peso = "tabelaTit"
	classe_subtit = "tabelaTit"

elseif errou="pt" or errou="pp" Then
	classe_peso = "tb_fundo_linha_erro"
else
	classe_peso = "tb_fundo_linha_peso"
	classe_subtit = "tb_subtit"
end if


qtd_colunas=UBOUND(linha_nome_colunas)+1
width_num_cham="20"
width_nom_aluno="340"
width_else=(1000-20-340)/(qtd_colunas-2)
val_default_lgb = parametros_gerais(unidade,curso,etapa,turma,co_materia,"default_lbg",outro)
val_default_lbs = parametros_gerais(unidade,curso,etapa,turma,co_materia,"default_lbs",outro)

%>

<form action="<%response.Write(action&opt)%>" name="nota" method="post" onSubmit="return checksubmit()">
  <table width="1000" border="0" cellspacing="0" cellpadding="0">
    <%if ubound(linha_pesos)>-1 then %>
    <tr>
      <% 
 for i= 0 to ubound(linha_pesos)
 		if i=0 then
			width=width_num_cham
			align="center"
		elseif i=1 then	
			width=width_nom_aluno
			align="left"			
		else
			width=width_else
			align="center"
		end if		

		if linha_pesos(i)="PESO" then
			linha_pesos_variaveis=split(ln_pesos_vars,"#!#")
			nome_pesos_variaveis=split(nm_pesos_vars,"#!#")		

			Set RStemp = Server.CreateObject("ADODB.Recordset")
			SQL_temp = "Select * from TB_Matriculas WHERE NU_Ano= "& ano_letivo &" AND CO_Situacao='C' AND NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
			Set RStemp = CON_AL.Execute(SQL_temp)
			
						
				nu_matriculatemp = RStemp("CO_Matricula")
			
				Set RSpeso = Server.CreateObject("ADODB.Recordset")
				SQL_peso = "Select "&linha_pesos_variaveis(i)&" from "& tb &" WHERE CO_Matricula = "& nu_matriculatemp & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				Set RSpeso = CON_N.Execute(SQL_peso)
					
				if RSpeso.EOF then
					if tb= "TB_NOTA_F" then
						if periodo<=4 then
							if linha_pesos_variaveis(i) = "PE_Teste" then
								valor_peso = 1
							elseif linha_pesos_variaveis(i) = "PE_Prova1" then
								valor_peso = 4
							else
								if periodo<4 then
									valor_peso = 5	
								else
						   			valor_peso =""										
								end if						 
							end if
						 else
						   valor_peso =""	
						 end if	
					end if				
				else	
					valor_peso=RSpeso(""&linha_pesos_variaveis(i)&"")
					if (valor_peso = "" or isnull(valor_peso)) and tb= "TB_NOTA_F" then
						if periodo<=4 then
							if linha_pesos_variaveis(i) = "PE_Teste" then
								valor_peso = 1
							elseif linha_pesos_variaveis(i) = "PE_Prova1" then
								valor_peso = 4
							else
								if periodo<4 then
									valor_peso = 5	
								else
						   			valor_peso =""										
								end if						 
							end if
						 else
						   valor_peso =""	
						 end if	
					end if
				end if		
				IF comunica="s" THEN		
					linha_pesos(i)=valor_peso&"<input name="&nome_pesos_variaveis(i)&" type=""hidden"" id="&nome_pesos_variaveis(i)&" class=""peso"" value="&valor_peso&">"	
				else	
					linha_pesos(i)="<input name="&nome_pesos_variaveis(i)&" type="&tipo&" id="&nome_pesos_variaveis(i)&" class=""peso"" value="&valor_peso&">"
				end if	
			
		end if				
 %>
      <td width="<%response.Write(width)%>" class="<%response.Write(classe_peso)%>"><div align="<%response.Write(align)%>">
          <%response.Write(linha_pesos(i))%>
        </div></td>
      <%	next%>
    </tr>
    <%end if%>
    <tr>
      <% for j= 0 to ubound(linha_nome_colunas)
 		if j=0 then
			width=width_num_cham
			align="center"
		elseif j=1 then	
			width=width_nom_aluno
			align="left"			
		else
			width=width_else
			align="center"
		end if							
 %>
      <td width="<%response.Write(width)%>" class="<%response.Write(classe_subtit)%>"><div align="<%response.Write(align)%>">
          <%response.Write(linha_nome_colunas(j))%>
        </div></td>
      <%	next%>
    </tr>
    <%
check = 2
nu_chamada_ckq = 0

Set RS = Server.CreateObject("ADODB.Recordset")
SQL_A = "Select * from TB_Matriculas WHERE NU_Ano="&ano_letivo&" AND NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
Set RS = CON_AL.Execute(SQL_A)


While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")
nu_matricula=nu_matricula*1
if isnumeric(calc) then
calc=calc*1
end if
onchange="""validaDefault('$$$$','!!!!');"""
	if subopcao="imp" Then
		classe = "tabela"
		classe_td_imp= " class = 'tabela'"
	elseif nu_matricula = calc then
		classe = "tb_fundo_linha_erro"
		onblur="""mudar_cor_blur_erro(####)"""	
		classe_td_imp= ""	  	   
	else
		if check mod 2 =0 then
			classe = "tb_fundo_linha_par" 
			onblur="""mudar_cor_blur_par(####)"""
		else 
			classe ="tb_fundo_linha_impar"
			onblur="""mudar_cor_blur_impar(####)"""
		end if 
		classe_td_imp= ""		
	end if

	Set RSs = Server.CreateObject("ADODB.Recordset")
	SQL_s ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& nu_matricula&" and TB_Matriculas.NU_Ano="&ano_letivo
	Set RSs = CON_AL.Execute(SQL_s)

	if RSs.EOF then
	%>
    <tr>
      <td width="20" class="<%response.Write(classe)%>">&nbsp;</td>
      <td width="340" class="<%response.Write(classe)%>">Matr�cula
        <%response.Write(nu_matricula)%>
        cadastrada em TB_Matriculas sem correspond�ncia em TB_Alunos</td>
      <%for m= 0 to ubound(nome_variaveis)
							width=width_else
							align="center"
					 %>
      <td width="<%response.Write(width)%>" class="<%response.Write(classe)%>">&nbsp;</td>
      <%next%>
    </tr>
    <%else
		situac=RSs("CO_Situacao")
		nome_aluno=RSs("NO_Aluno")	
		onchange = replace(onchange,"$$$$",linha_tabela&"c1") 
		onchange = replace(onchange,"!!!!",linha_tabela&"c2") 		 							
		onblur = replace(onblur,"####","celula"&nu_chamada)		
	'Verificando se algum aluno mudou de turma e inserindo uma linha cinza para o lugar do aluno
			if (nu_chamada_ckq <>nu_chamada - 1) then
				teste_nu_chamada = nu_chamada-nu_chamada_ckq
				teste_nu_chamada=teste_nu_chamada-1
				
				classe_anterior=classe
				if subopcao="imp" Then
					classe = "tabela"
				else	
					classe="tb_fundo_linha_falta"
				end if

				for k=1 to teste_nu_chamada 
					nu_chamada_ckq=nu_chamada_ckq+1
					nu_chamada_falta=nu_chamada_ckq
				%>
    <tr>
      <td width="20" class="<%response.Write(classe)%>"><input name="nu_chamada_<%response.Write(nu_chamada_falta)%>" class="borda_edit" type="hidden" value="<%response.Write(nu_chamada_falta)%>">
        <%response.Write(nu_chamada_falta)%>
        <input name="nu_matricula_<%response.Write(nu_chamada_falta)%>" type="hidden" value="falta"></td>
      <td width="340" class="<%response.Write(classe)%>">&nbsp;</td>
      <%for m= 0 to ubound(nome_variaveis)
							width=width_else
							align="center"
							nome_campo=nome_variaveis(m)&"_"&nu_chamada
							conteudo="&nbsp;"
					 %>
      <td width="<%response.Write(width)%>" class="<%response.Write(classe)%>"><div align="<%response.Write(align)%>">
          <%response.Write(conteudo)%>
        </div></td>
      <%next%>
    </tr>
    <%
					next
	'Inserindo o aluno seguinte aos que mudaram de turma
					nu_chamada_ckq=nu_chamada_ckq+1				
					inativo="N"
					if situac<>"C" then
						if subopcao="imp" Then
							classe = "tabela"
						else
							classe="tb_fundo_linha_falta"
						end if	
							valor="falta"
							nome_aluno=nome_aluno&" - Aluno Inativo"
							inativo="S"
					end if			
					%>
    <tr class="<%response.Write(classe_anterior)%>" id="<%response.Write("celula"&nu_chamada)%>">
      <td width="20" <%response.Write(classe_td_imp)%>><input name="nu_chamada_<%response.Write(nu_chamada)%>" class="borda_edit" type="hidden" value="<%response.Write(nu_chamada)%>">
        <%response.Write(nu_chamada)%>
        <input name="nu_matricula_<%response.Write(nu_chamada)%>" type="hidden" value='<%response.Write(nu_matricula)%>'></td>
      <td width="340" <%response.Write(classe_td_imp)%>><%response.Write(nome_aluno)%></td>
      <% 		valor=""
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				if opcao = "LSM" or opcao = "LBS" or opcao = "LBG" then
					SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & " AND NU_Periodo="&periodo				
				else
					SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				end if
			
				Set RS3 = CON_N.Execute(SQL_N)			 
				coluna=0	 
				 for n= 0 to ubound(nome_variaveis)
					width=width_else
					align="center"
					nome_campo=nome_variaveis(n)&"_"&nu_chamada
					onfocus = "javascript:this.form."&nome_campo&".select();"""
					checked=""
				
					if RS3.EOF then 
						valor=""
						if nome_variaveis(n) = "DEFAULT_LBG"  then  															
								valor=val_default_lgb
						elseif nome_variaveis(n) = "DEFAULT_LBS"  then  															
							    valor=val_default_lbs						
						end if			
					else
						nu_matricula=nu_matricula*1
						if isnumeric(calc) then
						calc=calc*1
						end if					
						if opt="err6" and nu_matricula = calc then	
							if errou=nome_variaveis(n) then
								valor=qerrou	
							else
								valor=Session(nome_variaveis(n))
							end if
						else	
							if tb = "TB_NOTA_V" and (variaveis_bd(n) = "VA_Media1" or variaveis_bd(n) = "VA_Bonus" or variaveis_bd(n) = "VA_Media2" or variaveis_bd(n) = "VA_Rec" or variaveis_bd(n) = "VA_Media3") then
								Set RS3F = Server.CreateObject("ADODB.Recordset")
								ano_letivo=ano_letivo*1
								if ano_letivo=2013 then
									SQL_NF = "Select * from TB_NOTA_F WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
								elseif ano_letivo<2017 then
									SQL_NF = "Select * from TB_NOTA_K WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
								else
									SQL_NF = "Select * from TB_NOTA_M WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo	
								end if											
								Set RS3F = CON_NF.Execute(SQL_NF)	
								if RS3F.eof then
									valor=""
								else
									valor=RS3F(""&variaveis_bd(n)&"")															
								end if
							elseif nome_variaveis(n) = "DEFAULT_LBG"  then  															
								valor=val_default_lgb
							elseif nome_variaveis(n) = "DEFAULT_LBS"  then  															
									valor=val_default_lbs										
							elseif nome_variaveis(n) = "CK_LBG" or nome_variaveis(n) = "CK_LBS" then  																					
							elseif nome_variaveis(n) <> "rs" and nome_variaveis(n) <> "rb" then 
								valor=RS3(""&variaveis_bd(n)&"")															
							end if	
						end if							
					end if
					
					if (valor="" or isnull(valor)) and subopcao="imp" then
						coluna=coluna+1	
						conteudo="&nbsp;"			
					else
						if (nome_variaveis(n) = "rs" or nome_variaveis(n)="rb") and inativo="N" then
							tipo_form = "checkbox"
							valor = "S"
						elseif nome_variaveis(n) = "CK_LBG" or nome_variaveis(n) = "CK_LBS" then 
							tipo_form = "checkbox"	
							onfocus = """"
							if nome_variaveis(n) = "CK_LBG" then
							    val_default = val_default_lgb
								checked="onclick=""javascript:this.form.val_bat_"&nu_chamada&".value="&val_default&";"""				
							else
							    val_default = val_default_lbs						
								checked="onclick=""javascript:this.form.val_bsi_"&nu_chamada&".value="&val_default&";"""								
							end if	
							if isnumeric(valor) then
								valor=valor*1	
								val_default=val_default*1		
								if valor = val_default then
									checked = "checked "&checked
								end if	
							end if		
			
						else
							tipo_form = tipo
						end if
					
						if nome_variaveis(n) = "DEFAULT_LBS" or nome_variaveis(n) = "DEFAULT_LBG" or nome_variaveis(n)="media_teste" or nome_variaveis(n)="media_prova" or nome_variaveis(n)="media1" or nome_variaveis(n)="media2" or nome_variaveis(n)="media3" then
							coluna=coluna
							if inativo<>"S"	then
								conteudo=valor
							else
								conteudo="&nbsp;"
							end if
						elseif situac<>"C" then
							coluna=coluna	
							conteudo=valor&"<input name='"&nome_campo&"' type='hidden' id='"&linha_tabela&"c"&coluna&"' value='"&valor&"'>"	
						else
							coluna=coluna+1
							if comunica="s" or subopcao="blq" then
							
								if nome_variaveis(n) = "rs" or nome_variaveis(n)="rb" then
									conteudo="&nbsp;"
								else
	'						if comunica="s" or (periodo_bloqueado="s" and sistema_origem="WN") then
									conteudo=valor&"<input name='"&nome_campo&"' type='"&tipo_form&"' id='"&linha_tabela&"c"&coluna&"' onFocus=""mudar_cor_focus(celula"&nu_chamada&");"&onfocus&" onBlur="&onblur&" onChange="&onchange&" value='"&valor&"' class=""nota"" size=""4"" maxlength=""3"" onkeydown=""keyPressed(this.id,event,"&notas_a_lancar&","&qtd_alunos&")"" "&checked&">"	
								end if	
							else
								conteudo="<input name='"&nome_campo&"' type='"&tipo_form&"' id='"&linha_tabela&"c"&coluna&"' onFocus=""mudar_cor_focus(celula"&nu_chamada&");"&onfocus&" onBlur="&onblur&" onChange="&onchange&" value='"&valor&"' class=""nota"" size=""4"" maxlength=""3"" onkeydown=""keyPressed(this.id,event,"&notas_a_lancar&","&qtd_alunos&")"" "&checked&">"										
							end if
						end if	
					end if	
					'conteudo=n
			 %>
      <td width="<%response.Write(width)%>" <%response.Write(classe_td_imp)%>><div align="<%response.Write(align)%>">
          <%response.Write(conteudo)%>
        </div></td>
      <%	next  
			  %>
    </tr>
    <%
	'Se os n�meros de chamada estiverem completos. Se n�o faltar aluno na turma.
			ELSE	
					inativo="N"
					if situac<>"C" then
						if subopcao="imp" Then
							classe = "tabela"
						else
							classe="tb_fundo_linha_falta"
						end if	
							valor="falta"
							nome_aluno=nome_aluno&" - Aluno Inativo"
							inativo="S"
					end if			
					nu_chamada_ckq=nu_chamada_ckq+1
					%>
    <tr class="<%response.Write(classe)%>" id="<%response.Write("celula"&nu_chamada)%>">
      <td width="20" <%response.Write(classe_td_imp)%>><input name="nu_chamada_<%response.Write(nu_chamada)%>" class="borda_edit" type="hidden" value="<%response.Write(nu_chamada)%>">
        <%response.Write(nu_chamada)%>
        <input name="nu_matricula_<%response.Write(nu_chamada)%>" type="hidden" value='<%response.Write(nu_matricula)%>'></td>
      <td width="340" <%response.Write(classe_td_imp)%>><%response.Write(nome_aluno)%></td>
      <% 		valor=""
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				if opcao = "LSM" or opcao = "LBS" or opcao = "LBG" then
					SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & " AND NU_Periodo="&periodo				
				else
					SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				end if

				Set RS3 = CON_N.Execute(SQL_N)			 
				coluna=0	 
				 for n= 0 to ubound(nome_variaveis)
					width=width_else
					align="center"
					nome_campo=nome_variaveis(n)&"_"&nu_chamada
					onfocus = "javascript:this.form."&nome_campo&".select();"""					
					checked=""
					val_default = ""
					if RS3.EOF then 
						valor=""
						if nome_variaveis(n) = "DEFAULT_LBG"  then  															
								valor=val_default_lgb
						elseif nome_variaveis(n) = "DEFAULT_LBS"  then  															
								valor=val_default_lbs		
						end if		
					else
						nu_matricula=nu_matricula*1
						if isnumeric(calc) then
						calc=calc*1
						end if					
						if opt="err6" and nu_matricula = calc then	
							if errou=nome_variaveis(n) then
								valor=qerrou	
							else
								valor=Session(nome_variaveis(n))
							end if
						else	
							if tb = "TB_NOTA_V" and (variaveis_bd(n) = "VA_Media1" or variaveis_bd(n) = "VA_Bonus" or variaveis_bd(n) = "VA_Media2" or variaveis_bd(n) = "VA_Rec" or variaveis_bd(n) = "VA_Media3") then
								Set RS3F = Server.CreateObject("ADODB.Recordset")
								ano_letivo=ano_letivo*1
								if ano_letivo=2013 then
									SQL_NF = "Select * from TB_NOTA_F WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
								elseif ano_letivo<2017 then
									SQL_NF = "Select * from TB_NOTA_K WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
								else
									SQL_NF = "Select * from TB_NOTA_M WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo	
								end if								
								Set RS3F = CON_NF.Execute(SQL_NF)	
								if RS3F.eof then
									valor=""
								else
									valor=RS3F(""&variaveis_bd(n)&"")													
								end if	
							elseif nome_variaveis(n) = "DEFAULT_LBG"  then  															
								valor=val_default_lgb
							elseif nome_variaveis(n) = "DEFAULT_LBS"  then  															
									valor=val_default_lbs										
							elseif nome_variaveis(n) = "CK_LBG" or nome_variaveis(n) = "CK_LBS" then  										
							elseif nome_variaveis(n) <> "rs" and nome_variaveis(n) <> "rb"  then  										
								valor=RS3(""&variaveis_bd(n)&"")
							end if	
						end if							
					end if
					
					if (valor="" or isnull(valor)) and subopcao="imp" then
						coluna=coluna+1	
						conteudo="&nbsp;"			
					else
						if (nome_variaveis(n) = "rs" or nome_variaveis(n)="rb") and inativo="N" then
							tipo_form = "checkbox"
							valor="S"
						elseif nome_variaveis(n) = "CK_LBG" or nome_variaveis(n) = "CK_LBS" then 
							tipo_form = "checkbox"
					        onfocus = """"							
							if nome_variaveis(n) = "CK_LBG" then
							    val_default = val_default_lgb
								checked="onclick=""javascript:this.form.val_bat_"&nu_chamada&".value="&val_default&";"""			
							else
							    val_default = val_default_lbs								
								checked="onclick=""javascript:this.form.val_bsi_"&nu_chamada&".value="&val_default&";"""								
							end if	
							if isnumeric(valor) then
								valor=valor*1	
								val_default=val_default*1	

								if valor = val_default then
									checked = "checked "&checked
								end if	
							end if	
					
						else
							tipo_form = tipo
						end if
					
					
						if nome_variaveis(n) = "DEFAULT_LBS" or nome_variaveis(n) = "DEFAULT_LBG" or nome_variaveis(n)="media_teste" or nome_variaveis(n)="media_prova" or nome_variaveis(n)="media1" or nome_variaveis(n)="media2" or nome_variaveis(n)="media3" then
							coluna=coluna	
							if inativo<>"S"	then
								conteudo=valor
							else
								conteudo="&nbsp;"
							end if
						elseif situac<>"C" then
							coluna=coluna	
							conteudo=valor&"<input name='"&nome_campo&"' type='hidden' id='"&linha_tabela&"c"&coluna&"' value='"&valor&"'>"	
						else
							coluna=coluna+1
							if comunica="s" or subopcao="blq" then
	'						if comunica="s" or (periodo_bloqueado="s" and sistema_origem="WN") then
								if nome_variaveis(n) = "rs" or nome_variaveis(n)="rb" then
									conteudo="&nbsp;"
								else
	'						if comunica="s" or (periodo_bloqueado="s" and sistema_origem="WN") then
									conteudo=valor&"<input name='"&nome_campo&"' type='"&tipo_form&"' id='"&linha_tabela&"c"&coluna&"' onFocus=""mudar_cor_focus(celula"&nu_chamada&");"&onfocus&" onBlur="&onblur&" onChange="&onchange&" value='"&valor&"' class=""nota"" size=""4"" maxlength=""3"" onkeydown=""keyPressed(this.id,event,"&notas_a_lancar&","&qtd_alunos&")"" "&checked&">"	
								end if		
							else
								conteudo="<input name='"&nome_campo&"' type='"&tipo_form&"' id='"&linha_tabela&"c"&coluna&"' onFocus=""mudar_cor_focus(celula"&nu_chamada&");"&onfocus&" onBlur="&onblur&" onChange="&onchange&" value='"&valor&"' class=""nota"" size=""4"" maxlength=""3"" onkeydown=""keyPressed(this.id,event,"&notas_a_lancar&","&qtd_alunos&")"" "&checked&">"										
							end if
						end if	
					end if	
					'conteudo=n
			 %>
      <td width="<%response.Write(width)%>" <%response.Write(classe_td_imp)%>><div align="<%response.Write(align)%>">
          <%response.Write(conteudo)%>
        </div></td>
      <%	next  
			  %>
    </tr>
    <%			
			END IF			              
		if situac<>"C" then
			linha_tabela=linha_tabela
		else
		
			linha_tabela=linha_tabela+1
		end if
 	
	END IF	
max=nu_chamada
	check = check+1 
RS.MoveNext
Wend 
session("max")=max

if subopcao="imp" then
else
%>
    <tr>
      <td colspan="<%response.Write(qtd_colunas)%>" class="tb_subtit_lanca_notas"><%	  
	if funcao_origem="EPN" or subopcao="blq" then
	%>
        <table width="100%" border="0" cellspacing="0">
          <tr>
            <td colspan="3"><hr></td>
          </tr>
          <tr>
            <td width="33%"><div align="center">
                <input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','<%response.Write(url_retorno)%>');return document.MM_returnValue" value="Voltar">
              </div></td>
            <td width="34%"><div align="center"> </div></td>
            <td width="33%"><div align="center"><font size="3" face="Courier New, Courier, mono"> </font></div></td>
          </tr>
        </table>
        <%
'	elseif  periodo_bloqueado="s" and sistema_origem="WN" then
'	 %>
        
        <!--			<table width="100%" border="0" cellspacing="0">
		  <tr>
			<td colspan="3">
			  <hr>
			</td>
		  </tr>	 
			<tr> 
			<td width="33%">
			<div align="center">
				<input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','altera.asp');return document.MM_returnValue" value="Voltar">
			  </div>
			  </td>
			  <td width="34%"> <div align="center">
				</div></td>
			  <td width="33%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
	
				  </font></div></td>
			</tr>
		  </table>
-->
        <%elseif comunica="s" then%>
        <table width="100%" border="0" cellspacing="0">
          <tr>
            <td colspan="3"><hr></td>
          </tr>
          <tr>
            <td width="33%"><div align="center">
                <input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','<%response.Write(url_retorno)%>');return document.MM_returnValue" value="Voltar">
              </div></td>
            <td width="34%"><div align="center">
                <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?ori=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" value="Cancelar">
              </div></td>
            <td width="33%"><div align="center"><font size="3" face="Courier New, Courier, mono">
                <input type="submit" name="Submit" value="Confirmar" class="botao_prosseguir">
                <input name="unidade" type="hidden" id="unidade" value="<%=unidade%>">
                <input name="curso" type="hidden" id="curso" value="<%=curso%>">
                <input name="etapa" type="hidden" id="etapa" value="<%=etapa%>">
                <input name="turma" type="hidden" id="turma" value="<%=turma%>">
                <input name="co_materia" type="hidden" id="co_materia" value="<%= co_materia%>">
                <input name="periodo" type="hidden" id="periodo" value="<%= periodo%>">
                <input name="co_prof" type="hidden" id="co_prof" value="<% = co_prof%>">
                <input name="max" type="hidden" id="max" value="<% =max%>">
                <input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>">
                <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>">
                </font></div></td>
          </tr>
        </table>
        <%else%>
        <table width="100%" border="0" align="center" cellspacing="0">
          <tr>
            <td colspan="3"><hr></td>
          </tr>
          <tr>
            <td width="33%"><div align="center">
                <input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','<%response.Write(url_retorno)%>');return document.MM_returnValue" value="Voltar">
              </div></td>
            <td width="34%"><div align="center"> 
               <input name="submit" type="submit" class="botao_prosseguir_comunicar" id="bt" value="Atualizar planilha de notas"> 
               
<!--               onClick="MM_goToURL('parent','<%response.Write(url_retorno)%>');return document.MM_returnValue"-->
               </div></td>
            <td width="33%"><div align="center">
                <input type="submit" name="submit" value="Salvar" class="botao_prosseguir">
                <input name="unidade" type="hidden" id="unidade" value="<%=unidade%>">
                <input name="curso" type="hidden" id="curso" value="<%=curso%>">
                <input name="etapa" type="hidden" id="etapa" value="<%=etapa%>">
                <input name="turma" type="hidden" id="turma" value="<%=turma%>">
                <input name="co_materia" type="hidden" id="co_materia" value="<%= co_materia%>">
                <input name="periodo" type="hidden" id="periodo" value="<%= periodo%>">
                <input name="co_prof" type="hidden" id="co_prof" value="<% = co_prof%>">
                <input name="max" type="hidden" id="max" value="<% =max%>">
                <input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>">
                <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>">
              </div></td>
          </tr>
        </table>
        <%end if
end if%></td>
    </tr>
  </table>
</form>
<%else%>
<table width="100%" border="0" cellspacing="0">
          <tr>
            <td colspan="3"><div align="center" class="form_corpo" >Fun&ccedil;&atilde;o n&atilde;o dispon&iacute;vel para essa turma ou disciplina</div></td>
          </tr>
          <tr>
            <td colspan="3"><hr></td>
          </tr>
          <tr>
            <td width="33%"><div align="center">
                <input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','<%response.Write(url_retorno)%>');return document.MM_returnValue" value="Voltar">
              </div></td>
            <td width="34%"><div align="center"> </div></td>
            <td width="33%"><div align="center"><font size="3" face="Courier New, Courier, mono"> </font></div></td>
          </tr>
        </table>
<%end if
end function%>
