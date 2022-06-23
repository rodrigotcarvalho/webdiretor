<!--#include file="caminhos.asp"-->
<!--#include file="bd_parametros.asp"-->
<!--#include file="bd_pauta.asp"-->
<!--#include file="../../global/conta_alunos.asp"-->
<!--#include file="../../global/tabelas_escolas.asp"-->
<!--#include file="../../global/notas_calculos_diversos.asp"-->
<%
Function notas (CAMINHO_al,CAMINHOn,unidade,curso,etapa,turma,co_materia,periodo,ano_letivo,co_usr,opcao,subopcao,outro)

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
			
'		Set CON_wr = Server.CreateObject("ADODB.Connection") 
'		ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
'		CON_wr.Open ABRIR_wr
		
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



 		Set RSapr = Server.CreateObject("ADODB.Recordset")
		SQLapr = "Select * from TB_Regras_Aprovacao WHERE CO_Curso = '"& curso &"' AND CO_Etapa='"&etapa&"'"
		Set RSapr = CON_0.Execute(SQLapr)
		
		if RSapr.EOF then
			ntvml=0
		else
			ntvml= RSapr("NU_Valor_M1")
		end if
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

ST_Per_1 = RS("ST_Per_1")
ST_Per_2 = RS("ST_Per_2")
ST_Per_3 = RS("ST_Per_3")
ST_Per_4 = RS("ST_Per_4")
ST_Per_5 = RS("ST_Per_5")
ST_Per_6 = RS("ST_Per_6")

if ano_letivo<2011 then
	if outro=4 then
		outro="<2011-4"
	else
		outro="<2011"	
	end if	
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
	vars_calc=dados_separados(7)
	action=dados_separados(8)
	notas_a_lancar=dados_separados(9)

	linha_pesos=split(ln_pesos_cols,"#!#")
	linha_pesos_variaveis=split(ln_pesos_vars,"#!#")
	nome_pesos_variaveis=split(nm_pesos_vars,"#!#")
	linha_nome_colunas=split(ln_nom_cols,"#!#")
	nome_variaveis=split(nm_vars,"#!#")
	variaveis_bd=split(nm_bd,"#!#")
	calcula_variavel=split(vars_calc,"#!#")

if subopcao="cln" then
	action="comunicar.asp"
	comunica="s" 
	opt="?opt=cln&obr="&obr
	tipo="hidden"
response.Redirect(action&opt&"&nota="&tb)
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
width_else=(1000-width_num_cham-width_nom_aluno)/(qtd_colunas-2)

%>

<form action="<%response.Write(action&opt)%>" name="nota" method="post" onSubmit="return checksubmit()">
  <table width="1000" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <% for i= 0 to ubound(linha_pesos)
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
			SQL_temp = "Select * from TB_Matriculas WHERE NU_Ano="&ano_letivo&" AND NU_Unidade = "& unidade &" AND NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
			Set RStemp = CON_AL.Execute(SQL_temp)
			
						
				nu_matriculatemp = RStemp("CO_Matricula")
			
				Set RSpeso = Server.CreateObject("ADODB.Recordset")
				SQL_peso = "Select "&linha_pesos_variaveis(i)&" from "& tb &" WHERE CO_Matricula = "& nu_matriculatemp & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				Set RSpeso = CON_N.Execute(SQL_peso)
					
				if RSpeso.EOF then
				else	
					valor_peso=RSpeso(""&linha_pesos_variaveis(i)&"")
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

errante=errante*1
nu_matricula=nu_matricula*1	


	if subopcao="imp" Then
		classe = "tabela"
		classe_td_imp= " class = 'tabela'"
	elseif nu_matricula = errante then
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
      <td width="<%response.Write(width_num_cham)%>" class="<%response.Write(classe)%>">&nbsp;</td>
      <td width="<%response.Write(width_nom_aluno)%>" class="<%response.Write(classe)%>">Matrícula
        <%response.Write(nu_matricula)%>
        cadastrada em TB_Matriculas sem correspondência em TB_Alunos</td>
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
				'response.write(teste_nu_chamada&"="&nu_chamada&"-"&nu_chamada_ckq)
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
      <td width="<%response.Write(width_num_cham)%>" class="<%response.Write(classe)%>"><input name="nu_chamada_<%response.Write(nu_chamada_falta)%>" class="borda_edit" type="hidden" value="<%response.Write(nu_chamada_falta)%>">
        <%response.Write(nu_chamada_falta)%>
        <input name="nu_matricula_<%response.Write(nu_chamada_falta)%>" type="hidden" value='falta'></td>
      <td width="<%response.Write(width_nom_aluno)%>" class="<%response.Write(classe)%>">&nbsp;</td>
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
					if situac<>"C" then
						if subopcao="imp" Then
							classe = "tabela"
						else
							classe="tb_fundo_linha_falta"
						end if	
							valor="falta"
							nome_aluno=nome_aluno&" - Aluno Inativo"
					end if			
					%>
    <tr class="<%response.Write(classe_anterior)%>" id="<%response.Write("celula"&nu_chamada)%>">
      <td width="<%response.Write(width_num_cham)%>" <%response.Write(classe_td_imp)%>><input name="nu_chamada_<%response.Write(nu_chamada)%>" class="borda_edit" type="hidden" value="<%response.Write(nu_chamada)%>">
        <%response.Write(nu_chamada)%>
        <input name="nu_matricula_<%response.Write(nu_chamada)%>" type="hidden" value='<%response.Write(nu_matricula)%>'></td>
      <td width="<%response.Write(width_nom_aluno)%>" <%response.Write(classe_td_imp)%>><%response.Write(nome_aluno)%></td>
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
						errante=errante*1
						nu_matricula=nu_matricula*1																	
						if valido="n" and nu_matricula = errante then
							if errou=nome_variaveis(n) then
								valor=qerrou	
							else
								valor=Session(nome_variaveis(n))
							end if
						else	
							if variaveis_bd(n)="CALCULADO" then
								valor="&nbsp;"
								'Nesse caso o valor é calculado pela função calcular_nota chamada mais abaixo
							else
								valor=RS3(""&variaveis_bd(n)&"")
							end if
						end if							
					end if
					
					if (valor="" or isnull(valor)) and subopcao="imp" then
						coluna=coluna+1	
						conteudo="&nbsp;"			
					else
						if nome_variaveis(n)="nav" or nome_variaveis(n)="media_teste" or nome_variaveis(n)="media_prova" or nome_variaveis(n)="media1" or nome_variaveis(n)="media2" or nome_variaveis(n)="media3" then
							coluna=coluna	
							conteudo=valor
						elseif calcula_variavel(n)="CALC1" then
							coluna=coluna
							valor=calcular_nota(calcula_variavel(n),CAMINHOn,tb,nu_matricula,mat_princ,co_materia,periodo)
							conteudo=valor		
						elseif situac<>"C" then
							coluna=coluna	
							conteudo=valor&"<input name='"&nome_campo&"' type='hidden' id='"&linha_tabela&"c"&coluna&"' value='"&valor&"'>"	
						else
							coluna=coluna+1
							if nu_matricula = errante then

								conteudo="<input name='"&nome_campo&"' type='"&tipo&"' id='"&linha_tabela&"c"&coluna&"' onFocus=""javascript:this.form."&nome_campo&".select();"" value='"&valor&"' class=""nota"" size=""4"" maxlength=""3"" onkeydown=""keyPressed(this.id,event,"&notas_a_lancar&","&qtd_alunos&")"">"								
							
							elseif comunica="s" or subopcao="blq" then
								conteudo=valor&"<input name='"&nome_campo&"' type='"&tipo&"' id='"&linha_tabela&"c"&coluna&"' onFocus=""mudar_cor_focus(celula"&nu_chamada&");javascript:this.form."&nome_campo&".select();"" onBlur="&onblur&"(celula"&nu_chamada&") value='"&valor&"' class=""nota"" size=""4"" maxlength=""3"" onkeydown=""keyPressed(this.id,event,"&notas_a_lancar&","&qtd_alunos&")"">"		

							else
								conteudo="<input name='"&nome_campo&"' type='"&tipo&"' id='"&linha_tabela&"c"&coluna&"' onFocus=""mudar_cor_focus(celula"&nu_chamada&");javascript:this.form."&nome_campo&".select();"" onBlur="&onblur&"(celula"&nu_chamada&") value='"&valor&"' class=""nota"" size=""4"" maxlength=""3"" onkeydown=""keyPressed(this.id,event,"&notas_a_lancar&","&qtd_alunos&")"">"										
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

	'Se os números de chamada estiverem completos. Se não faltar aluno na turma.
			ELSE	
					if situac<>"C" then
						if subopcao="imp" Then
							classe = "tabela"
						else
							classe="tb_fundo_linha_falta"
						end if	
							valor="falta"
							nome_aluno=nome_aluno&" - Aluno Inativo"
					end if			
					nu_chamada_ckq=nu_chamada_ckq+1
					%>
    <tr class="<%response.Write(classe)%>" id="<%response.Write("celula"&nu_chamada)%>">
      <td width="<%response.Write(width_num_cham)%>" <%response.Write(classe_td_imp)%>><input name="nu_chamada_<%response.Write(nu_chamada)%>" class="borda_edit" type="hidden" value="<%response.Write(nu_chamada)%>">
        <%response.Write(nu_chamada)%>
        <input name="nu_matricula_<%response.Write(nu_chamada)%>" type="hidden" value='<%response.Write(nu_matricula)%>'></td>
      <td width="<%response.Write(width_nom_aluno)%>" <%response.Write(classe_td_imp)%>><%response.Write(nome_aluno)%></td>
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
'response.Write(errou&"="&nome_variaveis(n))					
						errante=errante*1
						nu_matricula=nu_matricula*1						
						if valido="n" and nu_matricula = errante then	
							if errou=nome_variaveis(n) then
								valor=qerrou	
							else
								valor=Session(nome_variaveis(n))
							end if
						else	
							if variaveis_bd(n)="CALCULADO" then
								valor="&nbsp;"
								'Nesse caso o valor é calculado pela função calcular_nota chamada mais abaixo
							else
								valor=RS3(""&variaveis_bd(n)&"")
							end if
						end if							
					end if
					
					if (valor="" or isnull(valor)) and subopcao="imp" then
						coluna=coluna+1	
						conteudo="&nbsp;"			
					else
						if nome_variaveis(n)="nav" or nome_variaveis(n)="media_teste" or nome_variaveis(n)="media_prova" or nome_variaveis(n)="media1" or nome_variaveis(n)="media2" or nome_variaveis(n)="media3" then
							coluna=coluna	
							conteudo=valor
						elseif calcula_variavel(n)="CALC1" then
							coluna=coluna
							valor=calcular_nota(calcula_variavel(n),CAMINHOn,tb,nu_matricula,mat_princ,co_materia,periodo)
							conteudo=valor		
						elseif situac<>"C" then
							coluna=coluna	
							conteudo=valor&"<input name='"&nome_campo&"' type='hidden' id='"&linha_tabela&"c"&coluna&"' value='"&valor&"'>"	
						else
							coluna=coluna+1
							if nu_matricula = errante then

								conteudo="<input name='"&nome_campo&"' type='"&tipo&"' id='"&linha_tabela&"c"&coluna&"' onFocus=""javascript:this.form."&nome_campo&".select();"" value='"&valor&"' class=""nota"" size=""4"" maxlength=""3"" onkeydown=""keyPressed(this.id,event,"&notas_a_lancar&","&qtd_alunos&")"">"								
							
							elseif comunica="s" or subopcao="blq" then
								conteudo=valor&"<input name='"&nome_campo&"' type='"&tipo&"' id='"&linha_tabela&"c"&coluna&"' onFocus=""mudar_cor_focus(celula"&nu_chamada&");javascript:this.form."&nome_campo&".select();"" onBlur="&onblur&"(celula"&nu_chamada&") value='"&valor&"' class=""nota"" size=""4"" maxlength=""3"" onkeydown=""keyPressed(this.id,event,"&notas_a_lancar&","&qtd_alunos&")"">"		
							else
								conteudo="<input name='"&nome_campo&"' type='"&tipo&"' id='"&linha_tabela&"c"&coluna&"' onFocus=""mudar_cor_focus(celula"&nu_chamada&");javascript:this.form."&nome_campo&".select();"" onBlur="&onblur&"(celula"&nu_chamada&") value='"&valor&"' class=""nota"" size=""4"" maxlength=""3"" onkeydown=""keyPressed(this.id,event,"&notas_a_lancar&","&qtd_alunos&")"">"										
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
                <input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','altera.asp');return document.MM_returnValue" value="Voltar">
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
                <input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','notas.asp?opt=vt&amp;obr=<%=obr%>');return document.MM_returnValue" value="Voltar">
              </div></td>
            <td width="34%"><div align="center">
                <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?ori=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" value="Cancelar">
              </div></td>
            <td width="33%"><div align="center"><font size="3" face="Courier New, Courier, mono">
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
                <input name="Submit" type="button" class="botao_prosseguir_comunicar"  onClick = "if (! confirm('Ao comunicar o término não será mais possível lançar notas nessa planilha. Deseja continuar?')) { return false; };MM_goToURL('parent','notas.asp?or=01&opt=cln&obr=<%=obr%>');return document.MM_returnValue;" value="Comunicar ao Coordenador T&eacute;rmino da Planilha">
              </div></td>
            <td width="33%"><div align="center">
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
end if%></td>
    </tr>
  </table>
</form>
<%end function




Function boletim_avaliacao (CAMINHO_al,CAMINHOn,unidade,curso,etapa,turma,co_materia,periodo,ano_letivo,co_aluno,co_usr,opcao,subopcao,bloqueia_notas,bloqueia_alterado_por,bloqueia_data_alt,outro)

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

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON_N = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIR3
			
'		Set CON_wr = Server.CreateObject("ADODB.Connection") 
'		ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
'		CON_wr.Open ABRIR_wr
		
		Set CON_0 = Server.CreateObject("ADODB.Connection") 
		ABRIR_0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_0.Open ABRIR_0
		
		Set CON_AL = Server.CreateObject("ADODB.Connection") 
		ABRIR_AL = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_AL.Open ABRIR_AL
		
		Set CONg = Server.CreateObject("ADODB.Connection") 
		ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONg.Open ABRIRg		
		
    	Set CON_WF = Server.CreateObject("ADODB.Connection") 
		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_WF.Open ABRIR_WF		


 		Set RSapr = Server.CreateObject("ADODB.Recordset")
		SQLapr = "Select * from TB_Regras_Aprovacao WHERE CO_Curso = '"& curso &"' AND CO_Etapa='"&etapa&"'"
		Set RSapr = CON_0.Execute(SQLapr)
		
		if RSapr.EOF then
			ntvml=0
		else
			ntvml= RSapr("NU_Valor_M1")
		end if

if ano_letivo<2011 then
	if outro=4 then
		outro="<2011-4"
	else
		outro="<2011"	
	end if	
end if



	dados_tabela=verifica_dados_tabela(CAMINHOn,opcao,outro)
	dados_separados=split(dados_tabela,"#$#")
	tb=dados_separados(0)
	ln_bol_av_cols=dados_separados(11)
	ln_bol_av_span=dados_separados(12)
	nm_bol_av_vars=dados_separados(13)
	ln_bol_av_vars=dados_separados(14)
	vars_bol_av_calc=dados_separados(15)
	legenda=dados_separados(16)
	exibe_apr_pr=dados_separados(17)

	nome_variaveis=split(nm_bol_av_vars,"#!#")
	variaveis_bd=split(ln_bol_av_vars,"#!#")
	calcula_variavel=split(vars_bol_av_calc,"#!#")
	qtd_linhas=split(ln_bol_av_cols,"#!!#")
	linha_nome_colunas=split(qtd_linhas(0),"#!#") 
	linha_span=split(ln_bol_av_span,"#!#")
	exibe_notas=split(exibe_apr_pr,"#!#"	)
	
if subopcao="WA" then
	classe_subtit = "tb_subtit"
	showapr="s"
	showprova="s"	
	tamanho_tabela="990"	
	width_nom_disc="130"
	width_alt_por="230"
	width_data_alt="90" 		
elseif subopcao="WAI" then
	classe_subtit = "tabelaTit"	
	showapr="s"
	showprova="s"	
	tamanho_tabela="990"	
	width_nom_disc="130"
	width_alt_por="230"
	width_data_alt="90" 	
elseif subopcao="WF" then	
	classe_subtit = "style3"
	tamanho_tabela="820"	
	width_nom_disc="130"
	width_alt_por="200"
	width_data_alt="120" 	
	Set RS4 = Server.CreateObject("ADODB.Recordset")
	SQL_WF = "SELECT * FROM TB_Autoriza_WF WHERE NU_Unidade="&unidade&" AND CO_Curso='"&curso&"' and CO_Etapa='"&etapa&"'"
	RS4.Open SQL_WF, CON_WF	
	
	co_apr1=RS4("CO_apr1")
	co_apr2=RS4("CO_apr2")
	co_apr3=RS4("CO_apr3")
	co_apr4=RS4("CO_apr4")
	co_apr5=RS4("CO_apr5")
	'co_apr6=RS4("CO_apr6")
	co_prova1=RS4("CO_prova1")
	co_prova2=RS4("CO_prova2")
	co_prova3=RS4("CO_prova3")
	co_prova4=RS4("CO_prova4")	
	co_prova5=RS4("CO_prova5")
	'co_prova6=RS4("CO_prova6")	
		
	if periodo=1 then		
		if co_apr1="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova1="D"then
		showprova="n"
		else 
		showprova="s"
		end if
	elseif periodo=2 then		
		if co_apr2="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova2="D"then
		showprova="n"
		else 
		showprova="s"
		end if					
	elseif periodo=3 then		
		if co_apr3="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova3="D"then
		showprova="n"
		else 
		showprova="s"
		end if
	elseif periodo=4 then		
		if co_apr4="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova4="D"then
		showprova="n"
		else 
		showprova="s"
		end if
	elseif periodo=5 then		
		if co_apr5="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova5="D"then
		showprova="n"
		else 
		showprova="s"
		end if	
	elseif periodo=6 then		
		if co_apr6="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova6="D"then
		showprova="n"
		else 
		showprova="s"
		end if					
	end if	

elseif subopcao="WFI" then	
	classe_subtit = "tabelaTit"
	tamanho_tabela="990"	
	width_nom_disc="130"
	width_alt_por="200"
	width_data_alt="120" 	
	Set RS4 = Server.CreateObject("ADODB.Recordset")
	SQL_WF = "SELECT * FROM TB_Autoriza_WF WHERE NU_Unidade="&unidade&" AND CO_Curso='"&curso&"' and CO_Etapa='"&etapa&"'"
	RS4.Open SQL_WF, CON_WF
	
	co_apr1=RS4("CO_apr1")
	co_apr2=RS4("CO_apr2")
	co_apr3=RS4("CO_apr3")
	co_apr4=RS4("CO_apr4")
	co_apr5=RS4("CO_apr5")
	'co_apr6=RS4("CO_apr6")
	co_prova1=RS4("CO_prova1")
	co_prova2=RS4("CO_prova2")
	co_prova3=RS4("CO_prova3")
	co_prova4=RS4("CO_prova4")	
	co_prova5=RS4("CO_prova5")
	'co_prova6=RS4("CO_prova6")	
		
	if periodo=1 then		
		if co_apr1="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova1="D"then
		showprova="n"
		else 
		showprova="s"
		end if
	elseif periodo=2 then		
		if co_apr2="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova2="D"then
		showprova="n"
		else 
		showprova="s"
		end if					
	elseif periodo=3 then		
		if co_apr3="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova3="D"then
		showprova="n"
		else 
		showprova="s"
		end if
	elseif periodo=4 then		
		if co_apr4="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova4="D"then
		showprova="n"
		else 
		showprova="s"
		end if
	elseif periodo=5 then		
		if co_apr5="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova5="D"then
		showprova="n"
		else 
		showprova="s"
		end if	
	elseif periodo=6 then		
		if co_apr6="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova6="D"then
		showprova="n"
		else 
		showprova="s"
		end if					
	end if		
end if

%>
<table width="<%response.Write(tamanho_tabela)%>" border="0" cellspacing="0" cellpadding="0" align="center">
  <% 


	qtd_colunas=UBOUND(nome_variaveis)+1
	
if (tb="TB_NOTA_B" or tb="TB_NOTA_C") then
'acrescenta as coluna de peso  para o colspan da legenda
	if ano_letivo>=2011 then
	 qtd_colunas=qtd_colunas+4
	end if 
end if	
		
if bloqueia_alterado_por="s" and bloqueia_data_alt="n" then
	
	width_else=(1000-width_nom_disc-width_data_alt)/qtd_colunas

'acrescenta a coluna Disciplinas e Data/Hora para o colspan da legenda
	total_colunas=qtd_colunas+2	
	
elseif bloqueia_alterado_por="n" and bloqueia_data_alt="s" then

	width_else=(1000-width_nom_disc-width_alt_por)/qtd_colunas

'acrescenta a coluna Disciplinas e Alterado Por para o colspan da legenda
	total_colunas=qtd_colunas+2
	
elseif bloqueia_alterado_por="s" and bloqueia_data_alt="s" then

	width_else=(1000-width_nom_disc)/qtd_colunas	

'acrescenta a coluna Disciplinas para o colspan da legenda
	total_colunas=qtd_colunas+1
		
else

	width_else=(1000-width_nom_disc-width_alt_por-width_data_alt)/qtd_colunas

'acrescenta a coluna Disciplinas, Alterado Por e Data/Hora para o colspan da legenda
	total_colunas=qtd_colunas+3
		
end if	

%>
  <tr>
    <%

for j = 0 to ubound(linha_nome_colunas)
 
	if j=0 then
		width=width_nom_disc
		align="left"			
	else
		width=width_else
		align="center"
	end if	
	
	if linha_span(j)= "ROWSPAN" then
		span="rowspan=""2"""
	elseif linha_span(j)= "COLSPAN" then
		span="colspan=""2"""
		width=width*2
	ELSE
		span=""
	end if	
	
	if bloqueia_alterado_por="s" then	
		if linha_nome_colunas(j)="Alterado por" then
			exibe_coluna_alterado="n"
		else
			exibe_coluna_alterado="s"
		end if
	else
		exibe_coluna_alterado="s"
		if linha_nome_colunas(j)="Alterado por" then
			width=width_alt_por
		end if				
	end if	
	
	if bloqueia_data_alt="s" then				
		if linha_nome_colunas(j)="Data/Hora" then
			exibe_coluna_data="n"
		else
			exibe_coluna_data="s"			
		end if
	else
		exibe_coluna_data="s"
		if linha_nome_colunas(j)="Data/Hora" then
			width=width_data_alt			
		end if		
	end if			
		

	if (linha_nome_colunas(j)="Alterado por" and exibe_coluna_alterado="n") or (linha_nome_colunas(j)="Data/Hora" and exibe_coluna_data="n") then
	elseif (tb="TB_NOTA_B" or tb="TB_NOTA_C") AND LEFT(linha_nome_colunas(j),2)="TR" AND ano_letivo>=2011 then	
	 %>
    <td width="<%response.Write(width)%>" <%response.Write(span)%> class="<%response.Write(classe_subtit)%>"><div align="<%response.Write(align)%>">PESO</div></td>
    <td width="<%response.Write(width)%>" <%response.Write(span)%> class="<%response.Write(classe_subtit)%>"><div align="<%response.Write(align)%>">
        <%response.Write(linha_nome_colunas(j))%>
      </div></td>
    <% 
	else
	 %>
    <td width="<%response.Write(width)%>" <%response.Write(span)%> class="<%response.Write(classe_subtit)%>"><div align="<%response.Write(align)%>">
        <%response.Write(linha_nome_colunas(j))%>
      </div></td>
    <%
	end if
next%>
  </tr>
  <% 
if ubound(qtd_linhas)>0	then 
linha2_nome_colunas=split(qtd_linhas(1),"#!#") 
%>
  <tr>
    <%
	for j2= 0 to ubound(linha2_nome_colunas)
 
		if j2=0 then
			width=width_else
			align="center"			
		else
			width=width_else
			align="center"
		end if	
	
	 %>
    <td width="<%response.Write(width)%>"  class="<%response.Write(classe_subtit)%>"><div align="<%response.Write(align)%>">
        <%response.Write(linha2_nome_colunas(j2))%>
      </div></td>
    <%
	next%>
  </tr>
  <%
end if

check = 2
nu_chamada_ckq = 0

Set RS = Server.CreateObject("ADODB.Recordset")
SQL_A = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
RS.Open SQL_A, CON0

While Not RS.EOF

	nu_matricula = co_aluno 
	co_materia=RS("CO_Materia")
	mae=RS("IN_MAE")
	fil=RS("IN_FIL")
	in_co=RS("IN_CO")
	nu_peso=RS("NU_Peso")
	ordem=RS("NU_Ordem_Boletim")

		Set RS1a = Server.CreateObject("ADODB.Recordset")
		SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia='"&co_materia&"'"
		RS1a.Open SQL1a, CON0
		
	mat_princ=RS1a("CO_Materia_Principal")
	no_materia=RS1a("NO_Materia")


	if mat_princ="" or isnull(mat_princ) then
		mat_princ=co_materia
	end if


	if subopcao="WAI" or subopcao="WFI" Then
		classe = "tabela"
		classe_td_imp= " class = 'tabela'"	  	   
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
    <td width="<%response.Write(width_nom_disc)%>" class="<%response.Write(classe)%>">Matrícula
      <%response.Write(nu_matricula)%>
      cadastrada em TB_Matriculas sem correspondência em TB_Alunos</td>
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
		if situac<>"C" then
			if subopcao="WAI" Then
				classe = "tabela"
			else
				classe="tb_fundo_linha_falta"
			end if	
				valor="falta"
				nome_aluno=nome_aluno&" - Aluno Inativo"
		end if			
		nu_chamada_ckq=nu_chamada_ckq+1
				%>
  <tr class="<%response.Write(classe)%>" id="<%response.Write("celula"&nu_chamada)%>">
    <td width="<%response.Write(width_nom_disc)%>" <%response.Write(classe_td_imp)%>><%response.Write(no_materia)%></td>
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
				ps=n+1
				valor="&nbsp;"
				valor_peso="&nbsp;"
			else
				ps=n+1
				if ano_letivo>=2011 then
					if linha_nome_colunas(ps)="TR1" then
						valor_peso=RS3("VA_Pt1")
					elseif linha_nome_colunas(ps)="TR2" then
						valor_peso=RS3("VA_Pt2")	
					elseif linha_nome_colunas(ps)="TR3" then
						valor_peso=RS3("VA_Pt3")	
					elseif linha_nome_colunas(ps)="TR4" then
						valor_peso=RS3("VA_Pt4")	
					end if
					if valor_peso="" or isnull(valor_peso) then
						valor_peso="&nbsp;"					
					end if
				end if	
				if variaveis_bd(n)="CALCULADO" then
					valor="&nbsp;"
					'Nesse caso o valor é calculado pela função calcular_nota chamada mais abaixo
				else
					valor=RS3(""&variaveis_bd(n)&"")
				end if						
			end if
				
			if valor="" or isnull(valor) then
				conteudo="&nbsp;"			
			else
				if calcula_variavel(n)="CALC1" then
					conteudo=calcular_nota(calcula_variavel(n),CAMINHOn,tb,nu_matricula,mat_princ,co_materia,periodo)	
				else
					conteudo=valor									
				end if	
			end if	
			
			if (tb="TB_NOTA_B" or tb="TB_NOTA_C") AND LEFT(linha_nome_colunas(ps),2)="TR" AND ano_letivo>=2011 then	
		 %>
    <td width="<%response.Write(width)%>" <%response.Write(classe_td_imp)%>><div align="<%response.Write(align)%>">
        <%
                if exibe_notas(n)="0" or (exibe_notas(n)="A" and showapr="s") or (exibe_notas(n)="P" and showprova="s") or (exibe_notas(n)="M" and showapr="s" and showprova="s") then		
                    response.Write(valor_peso)
                else
                    response.Write("&nbsp;") 
                end if
                        %>
      </div></td>
    <%end if%>
    <td width="<%response.Write(width)%>" <%response.Write(classe_td_imp)%>><div align="<%response.Write(align)%>">
        <%
			if exibe_notas(n)="0" or (exibe_notas(n)="A" and showapr="s") or (exibe_notas(n)="P" and showprova="s") or (exibe_notas(n)="M" and showapr="s" and showprova="s") then		
				response.Write(conteudo)
			else
				response.Write("&nbsp;")
			end if
					%>
      </div></td>
    <%	next  
		if exibe_coluna_alterado="s" then 	
			width=width_alt_por
				if RS3.EOF then 
					usuario_grav=""
				else		
					usuario_grav=RS3("CO_Usuario")	
				end if	
				if usuario_grav="" or isnull(usuario_grav) then
					no_usuario="&nbsp;"
				else
						Set RS_pro = Server.CreateObject("ADODB.Recordset")
						SQL_pro = "SELECT * FROM TB_Usuario WHERE CO_Usuario="& usuario_grav
						RS_pro.Open SQL_pro, CON
				
					if RS_pro.EOF then
						no_usuario="&nbsp;"
					else
						no_usuario=RS_pro("NO_Usuario")
					end if
				end if								
												
	%>
    <td width="<%response.Write(width)%>" <%response.Write(classe_td_imp)%>><div align="<%response.Write(align)%>">
        <%
					if showapr="s" or showprova="s" then
						response.Write(no_usuario)
					else
						response.Write("&nbsp;")
					end if
			%>
      </div></td>
    <%
		end if
		if exibe_coluna_data="s" then 
			width=width_data_alt	
				if RS3.EOF then 
					data_grav=""
					hora_grav=""
				else		
					data_grav=RS3("DA_Ult_Acesso")
					hora_grav=RS3("HO_ult_Acesso")		
				end if			
	
				if data_grav="" or isnull(data_grav) then
					
					data_grav="&nbsp;"
				
				else
				
					dados_dtd= split(data_grav, "/" )
					dia_de= dados_dtd(0)
					mes_de= dados_dtd(1)
					ano_de= dados_dtd(2)
					dia_de=dia_de*1
					mes_de=mes_de*1

					if dia_de<10 then
						dia_de="0"&dia_de
					end if
								
					if mes_de<10 then
						mes_de="0"&mes_de
					end if
					data_format=dia_de&"/"&mes_de&"/"&ano_de
					
				end if				 
				if hora_grav="" or isnull(hora_grav) then
				
					hora_grav="&nbsp;"
				
				else
					dados_hrd= split(hora_grav, ":" )
					h_de= dados_hrd(0)
					min_de= dados_hrd(1)
					h_de=h_de*1
					min_de=min_de*1
					
				
					if h_de<10 then
						h_de="0"&h_de
					end if
					if min_de<10 then
						min_de="0"&min_de
					end if	
					hora_format=h_de&":"&min_de			
					
				end if
				
				if data_grav="&nbsp;" and hora_grav<>"&nbsp;"then
					data_exibe=data_grav	
				elseif data_grav<>"&nbsp;" and hora_grav="&nbsp;"then		
					data_exibe=hora_format	
				elseif data_grav="&nbsp;" and hora_grav="&nbsp;"then
					data_exibe="&nbsp;"	
				else				
					data_exibe=data_format&", "&hora_format
				end if
		 %>
    <td width="<%response.Write(width)%>" <%response.Write(classe_td_imp)%>><div align="<%response.Write(align)%>">
        <%
					if showapr="s" or showprova="s" then
						response.Write(data_exibe)
					else
						response.Write("&nbsp;")
					end if					
					%>
      </div></td>
  </tr>
  <%
		 end if
	END IF	
	check = check+1 
RS.MoveNext
Wend 
%>
  <tr>
    <td colspan="<%response.Write(total_colunas)%>" class="linhaTopoR"><%response.Write(legenda)%></td>
  </tr>
</table>
</form>
<%end function

Function pauta (CAMINHO_al,CAMINHOn,unidade,curso,etapa,turma,co_materia,periodo,ano_letivo,co_prof,opcao,subopcao,qtd_datas,outro)

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
		endereco_origem="../wa/professor/cna/mpe/"
	end if
end if	


		Set CON_N = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIR3
			
'		Set CON_wr = Server.CreateObject("ADODB.Connection") 
'		ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
'		CON_wr.Open ABRIR_wr
		
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



 		Set RSapr = Server.CreateObject("ADODB.Recordset")
		SQLapr = "Select * from TB_Regras_Aprovacao WHERE CO_Curso = '"& curso &"' AND CO_Etapa='"&etapa&"'"
		Set RSapr = CON_0.Execute(SQLapr)
		
		if RSapr.EOF then
			ntvml=0
		else
			ntvml= RSapr("NU_Valor_M1")
		end if
qtd_alunos=contalunos(CAMINHO_al,ano_letivo,unidade,curso,etapa,turma,"C")

 		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL_0 = "Select * from TB_Materia WHERE CO_Materia = '"& co_materia &"'"
		Set RS0 = CON_0.Execute(SQL_0)
if RS0.eof then
	mat_princ=co_materia
else
mat_princ=RS0("CO_Materia_Principal")

	if mat_princ="" or isnull(mat_princ) then
		mat_princ=co_materia
	end if
end if


		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' AND CO_Materia_Principal = '"& co_materia &"'"
		Set RS = CONg.Execute(CONEXAO)

ST_Per_1 = RS("ST_Per_1")
ST_Per_2 = RS("ST_Per_2")
ST_Per_3 = RS("ST_Per_3")
ST_Per_4 = RS("ST_Per_4")
ST_Per_5 = RS("ST_Per_5")
ST_Per_6 = RS("ST_Per_6")

if ano_letivo<2011 then
	if outro=4 then
		outro="<2011-4"
	else
		outro="<2011"	
	end if	
end if

datas_formatado = buscaDataPauta(CAMINHOn, co_prof, unidade,curso,etapa,turma, mat_princ, co_materia, periodo, p_Vetor_Datas_Consulta, qtd_datas)
'response.Write(datas_formatado)

dados_tabela=verifica_dados_tabela(CAMINHOn,opcao,outro)
	dados_separados=split(dados_tabela,"#$#")
	tb=dados_separados(0)
'	ln_pesos_cols=dados_separados(1)
'	ln_pesos_vars=dados_separados(2)
'	nm_pesos_vars=dados_separados(3)
'	ln_nom_cols=dados_separados(4)
	nm_vars=datas_formatado
	'ln_nom_cols = "N&ordm;#!#Nome#!#"&nm_vars	
	ln_nom_cols = nm_vars		
'	nm_bd=dados_separados(6)
'	vars_calc=dados_separados(7)
'	action=dados_separados(8)
'	notas_a_lancar=dados_separados(9)
'
'	linha_pesos=split(ln_pesos_cols,"#!#")
'	linha_pesos_variaveis=split(ln_pesos_vars,"#!#")
'	nome_pesos_variaveis=split(nm_pesos_vars,"#!#")
	linha_nome_colunas=split(ln_nom_cols,"#!#")
	nome_variaveis=split(nm_vars,"#!#")
'	variaveis_bd=split(nm_bd,"#!#")
'	calcula_variavel=split(vars_calc,"#!#")
	datas_pauta=split(p_Vetor_Datas_Consulta,"#!#")


if subopcao="cln" then
	action="comunicar.asp"
	comunica="s" 
	opt="?opt=cln&obr="&obr
	tipo="hidden"
'response.Redirect(action&opt&"&nota="&tb)
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
width_num_cham=20
width_nom_aluno=340
width_else=40
tableWidth=width_num_cham+width_nom_aluno+(width_else*UBOUND(linha_nome_colunas))


%>
<table id="cabecalhoHorizontal">
  <thead>
    <tr>
      <% for j= 0 to ubound(linha_nome_colunas)
			valor = linha_nome_colunas(j)
			seq_pauta = buscaSeqDataPauta(CAMINHOn, datas_pauta(j), co_prof, unidade,curso,etapa,turma, mat_princ, co_materia, periodo, outro)		
			vetorSeqAula=buscaSeqAula(CAMINHOn,seq_pauta, datas_pauta(j), outro)
			seq_aula=split(vetorSeqAula,"#!#")
			colspan=ubound(seq_aula)-1
			if colspan<0 then
				colspan=0 
			end if
			width=80+(40*colspan)	
			data_form=replace(datas_pauta(j),"/",".")
			valor = "<input name=""data_form"" type=""checkbox"" value="""&data_form&""" />"&linha_nome_colunas(j)				


					
 %>
      <td style="height:40px; min-width: <%response.Write(width)%>px;width: <%response.Write(width)%>px;max-width: <%response.Write(width)%>px;" class="<%response.Write(classe_subtit)%>" align="center"><%response.Write(valor)%></td>
      <%	next
		if qtd_datas=366 then
		%>
      <td style="min-width: 200px;" class="<%response.Write(classe_subtit)%>" align="center"><%response.Write("Total de Faltas")%></td>
      <%end if%>
    </tr>
    <tr>
      <% 
	  
for j= 0 to ubound(linha_nome_colunas)
		align="center"
		seq_pauta = buscaSeqDataPauta(CAMINHOn, datas_pauta(j), co_prof, unidade,curso,etapa,turma, mat_princ, co_materia, periodo, outro)		
	
		vetorSeqAula=buscaSeqAula(CAMINHOn,seq_pauta, datas_pauta(j), outro)
		seq_aula=split(vetorSeqAula,"#!#")	
		colspan=ubound(seq_aula)-1
		if colspan<=1 then
			width=80			
		else
			width=80+(40*colspan)	
		end if	
%>
      <td  class="<%response.Write(classe_subtit)%>" align="center"><table border="0" cellpadding="0" cellspacing="0">
          <tr>
            <%
 
       
	  for col=0 to ubound(seq_aula) 			
			nome_tempo = buscaTempoAula(CAMINHOn, seq_pauta, seq_aula(col), outro)

						
	  %>
            <td style="height:20px; min-width: 40px;width: <%response.Write(width)%>px;max-width: <%response.Write(width)%>px;" class="<%response.Write(classe_subtit)%>" align="center"><%response.Write(nome_tempo)%></td>
            <%next
%>
          </tr>
        </table></td>
      <%
      next
	  		if qtd_datas=366 then
	  %>
      <td style="min-width: 200px;" class="<%response.Write(classe_subtit)%>"  align="center">&nbsp;</td>
      <%end if%>
    </tr>
  </thead>
</table>
<table id="cabecalhoVertical">
  <thead>
    <tr>
      <td style="max-width: 20px;"  class="<%response.Write(classe_subtit)%>">N&ordm;</td>
      <td style="min-width: 316px;width: 316px;max-width: 316px;" class="<%response.Write(classe_subtit)%>">Nome</td>
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
wrkTotalFaltas=0
errante=errante*1
nu_matricula=nu_matricula*1	


	if subopcao="imp" Then
		classe = "tabela"
		classe_td_imp= " class = 'tabela'"
	elseif nu_matricula = errante then
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
      <td style="max-width: 20px;"  class="<%response.Write(classe)%>">&nbsp;</td>
      <td style="min-width: 316px;width: 316px;max-width: 316px;"   class="<%response.Write(classe)%>">Matrícula
        <%response.Write(nu_matricula)%>
        cadastrada em TB_Matriculas sem correspondência em TB_Alunos</td>
    </tr>
    <%else
		situac=RSs("CO_Situacao")
		nome_aluno=RSs("NO_Aluno")	
	'Verificando se algum aluno mudou de turma e inserindo uma linha cinza para o lugar do aluno
			if (nu_chamada_ckq <>nu_chamada - 1) then
				teste_nu_chamada = nu_chamada-nu_chamada_ckq
				teste_nu_chamada=teste_nu_chamada-1
				'response.write(teste_nu_chamada&"="&nu_chamada&"-"&nu_chamada_ckq)
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
      <td style="max-width: 20px;"  class="<%response.Write(classe)%>"><input name="nu_chamada_<%response.Write(nu_chamada_falta)%>" class="borda_edit" type="hidden" value="<%response.Write(nu_chamada_falta)%>" />
        <%response.Write(nu_chamada_falta)%>
        <input name="nu_matricula_<%response.Write(nu_chamada_falta)%>" type="hidden" value='falta' /></td>
      <td style="min-width: 316px;width: 316px;max-width: 316px;"   class="<%response.Write(classe)%>">&nbsp;</td>
    </tr>
    <%				
					next
	'Inserindo o aluno seguinte aos que mudaram de turma
					nu_chamada_ckq=nu_chamada_ckq+1		
					if situac<>"C" then
						if subopcao="imp" Then
							classe = "tabela"
						else
							classe="tb_fundo_linha_falta"
						end if	
							valor="falta"
							nome_aluno=nome_aluno&" - Aluno Inativo"
					end if			
					%>
    <tr class="<%response.Write(classe_anterior)%>" id="<%response.Write("celula"&nu_chamada)%>">
      <td style="max-width: 20px;"  <%response.Write(classe_td_imp)%>><input name="nu_chamada_<%response.Write(nu_chamada)%>" class="borda_edit" type="hidden" value="<%response.Write(nu_chamada)%>" />
        <%response.Write(nu_chamada)%>
        <input name="nu_matricula_<%response.Write(nu_chamada)%>" type="hidden" value='<%response.Write(nu_matricula)%>' /></td>
      <td style="min-width: 316px;width: 316px;max-width: 316px;"   <%response.Write(classe_td_imp)%>><%response.Write(nome_aluno)%></td>
    </tr>
    <%

	'Se os números de chamada estiverem completos. Se não faltar aluno na turma.
			ELSE	
					if situac<>"C" then
						if subopcao="imp" Then
							classe = "tabela"
						else
							classe="tb_fundo_linha_falta"
						end if	
							valor="falta"
							nome_aluno=nome_aluno&" - Aluno Inativo"
					end if			
					nu_chamada_ckq=nu_chamada_ckq+1
					%>
    <tr class="<%response.Write(classe)%>" id="<%response.Write("celula"&nu_chamada)%>">
      <td style="max-width: 20px;" <%response.Write(classe_td_imp)%>><input name="nu_chamada_<%response.Write(nu_chamada)%>" class="borda_edit" type="hidden" value="<%response.Write(nu_chamada)%>" />
        <%response.Write(nu_chamada)%>
        <input name="nu_matricula_<%response.Write(nu_chamada)%>" type="hidden" value='<%response.Write(nu_matricula)%>' /></td>
      <td style="min-width: 316px;width: 316px;max-width: 316px;" <%response.Write(classe_td_imp)%>><%response.Write(nome_aluno)%></td>
    </tr>
    <%			
			END IF	%>
    <%            		              
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

%>
  </thead>
</table>
<table id="dados">
  <tbody>
    <tr>
      <%for c= 0 to ubound(linha_nome_colunas)
		align="center"
		seq_pauta = buscaSeqDataPauta(CAMINHOn, datas_pauta(c), co_prof, unidade,curso,etapa,turma, mat_princ, co_materia, periodo, outro)		
		vetorSeqAula=buscaSeqAula(CAMINHOn,seq_pauta, datas_pauta(c), outro)
		seq_aula=split(vetorSeqAula,"#!#")	

		colspan=ubound(seq_aula)-1
		if colspan<1 then
			width=80
			widthMin = round(width/(ubound(seq_aula)+1))
		else
			width=80+(40*colspan)	
			widthMin = round(width/(ubound(seq_aula)+1))					
		end if		
		if ubound(seq_aula)+1 > 3 then
			widthMin=widthMin-2
		else
			widthMin=widthMin-1		
		end if
		widthMin = replace(widthMin,",",".")
		
	  for col=0 to ubound(seq_aula) 	
				
			 %>
      <td style="min-width: <%response.Write(widthMin)%>px;width: <%response.Write(widthMin)%>px;max-width: <%response.Write(width)%>px;" align="center">&nbsp;</td>
      <%	next  
		next
		if qtd_datas=366 then
			  %>
      <td style="min-width: 198px;" >&nbsp;</td>
      <%end if%>
    </tr>
    <%
nu_chamada_ckq = 0
Set RS = Server.CreateObject("ADODB.Recordset")
SQL_A = "Select * from TB_Matriculas WHERE NU_Ano="&ano_letivo&" AND NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
Set RS = CON_AL.Execute(SQL_A)


While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")
wrkTotalFaltas=0
errante=errante*1
nu_matricula=nu_matricula*1	


	if subopcao="imp" Then
		classe = "tabela"
		classe_td_imp= " class = 'tabela'"
	elseif nu_matricula = errante then
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
      <%for m= 0 to ubound(nome_variaveis)
					 %>
      <td style="max-width: 55px;" class="<%response.Write(classe)%>">&nbsp;</td>
      <%next%>
      <td style="min-width: 198px;" class="<%response.Write(classe)%>">&nbsp;</td>
    </tr>
    <%else
		situac=RSs("CO_Situacao")
		nome_aluno=RSs("NO_Aluno")	
	'Verificando se algum aluno mudou de turma e inserindo uma linha cinza para o lugar do aluno
			if (nu_chamada_ckq <>nu_chamada - 1) then
				teste_nu_chamada = nu_chamada-nu_chamada_ckq
				teste_nu_chamada=teste_nu_chamada-1
				'response.write(teste_nu_chamada&"="&nu_chamada&"-"&nu_chamada_ckq)
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
      <%for m= 0 to ubound(linha_nome_colunas)
							align="center"
							nome_campo=nome_variaveis(m)&"_"&nu_chamada
							conteudo="&nbsp;"
					 %>
      <td style="max-width: 55px;" class="<%response.Write(classe)%>" align="center"><%response.Write(conteudo)%></td>
      <%next
	  if qtd_datas=366 then	
	  %>      
      <td style="min-width: 198px;" class="<%response.Write(classe)%>" align="center"><%response.Write(wrkTotalFaltas)%></td>
   		<%end if%>
    </tr>
    <%				
					next
	'Inserindo o aluno seguinte aos que mudaram de turma
					nu_chamada_ckq=nu_chamada_ckq+1		
					if situac<>"C" then
						if subopcao="imp" Then
							classe = "tabela"
						else
							classe="tb_fundo_linha_falta"
						end if	
							valor="falta"
							nome_aluno=nome_aluno&" - Aluno Inativo"
					end if			
					%>
    <tr class="<%response.Write(classe_anterior)%>" id="<%response.Write("celula"&nu_chamada)%>">
      <% 
	for c= 0 to ubound(linha_nome_colunas)
		align="center"
		seq_pauta = buscaSeqDataPauta(CAMINHOn, datas_pauta(c), co_prof, unidade,curso,etapa,turma, mat_princ, co_materia, periodo, outro)		
		vetorSeqAula=buscaSeqAula(CAMINHOn,seq_pauta, datas_pauta(c), outro)
		seq_aula=split(vetorSeqAula,"#!#")	
		
		colspan=ubound(seq_aula)-1
		if colspan<1 then
			width=80
			widthMin = round(width/(ubound(seq_aula)+1))
		else
			width=80+(40*colspan)	
			widthMin = round(width/(ubound(seq_aula)+1))					
		end if	
		'width = width-1	
		if ubound(seq_aula)+1 > 3 then
			widthMin=widthMin-2
		else
			widthMin=widthMin-1		
		end if
		widthMin = replace(widthMin,",",".")
							
	  for col=0 to ubound(seq_aula) 	
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL_N = "Select * from TB_Pauta_Faltas WHERE CO_Matricula = "& nu_matricula & " AND NU_Pauta = "& seq_pauta &" AND NU_Seq = "& seq_aula(col) 
				Set RS3 = CON_N.Execute(SQL_N)			 
				coluna=0	 
				
					align="center"
					
					if RS3.EOF then 
						conteudo="P"
						classe = "texto_azul"
					else															
						conteudo="F"																		
						classe = "texto_vermelho"
						wrkTotalFaltas=wrkTotalFaltas+1
					end if	
			 %>
      <td style="min-width: <%response.Write(widthMin)%>px;width: <%response.Write(widthMin)%>px;max-width: <%response.Write(width)%>px;" class="<%response.Write(classe)%>" align="center"><%response.Write(conteudo)%></td>
      <%	next  
		next
				
		classe = "texto_azul"	
		if qtd_datas=366 then	
			if wrkTotalFaltas>0 then
				classe = "texto_vermelho"
			end if		
			  %>
      <td style="min-width: 198px;" class="<%response.Write(classe)%>" align="center"><%response.Write(wrkTotalFaltas)%></td>
      <%end if%>
    </tr>
    <%

	'Se os números de chamada estiverem completos. Se não faltar aluno na turma.
			ELSE	
					if situac<>"C" then
						if subopcao="imp" Then
							classe = "tabela"
						else
							classe="tb_fundo_linha_falta"
						end if	
							valor="falta"
							nome_aluno=nome_aluno&" - Aluno Inativo"
					end if			
					nu_chamada_ckq=nu_chamada_ckq+1
					%>
    <tr class="<%response.Write(classe)%>" id="<%response.Write("celula"&nu_chamada)%>">
      <% 
	for c= 0 to ubound(linha_nome_colunas) 
		align="center"
		seq_pauta = buscaSeqDataPauta(CAMINHOn, datas_pauta(c), co_prof, unidade,curso,etapa,turma, mat_princ, co_materia, periodo, outro)		
		vetorSeqAula=buscaSeqAula(CAMINHOn,seq_pauta, datas_pauta(c), outro)
		seq_aula=split(vetorSeqAula,"#!#")	
		
		colspan=ubound(seq_aula)-1
		if colspan<1 then
			width=80
			widthMin = round(width/(ubound(seq_aula)+1))
		else
			width=80+(40*colspan)	
			widthMin = round(width/(ubound(seq_aula)+1))					
		end if
		'width = width-1	
		if ubound(seq_aula)+1 > 3 then
			widthMin=widthMin-2
		else
			widthMin=widthMin-1		
		end if	
		widthMin = replace(widthMin,",",".")	
	  for col=0 to ubound(seq_aula) 	
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL_N = "Select * from TB_Pauta_Faltas WHERE CO_Matricula = "& nu_matricula & " AND NU_Pauta = "& seq_pauta &" AND NU_Seq = "& seq_aula(col) 
				Set RS3 = CON_N.Execute(SQL_N)			 
				coluna=0	 
				
					align="center"
				
					if RS3.EOF then 
						conteudo="P"
						classe = "texto_azul"
					else															
						conteudo="F"																		
						classe = "texto_vermelho"
						wrkTotalFaltas=wrkTotalFaltas+1
					end if	

			 %>
      <td style="min-width: <%response.Write(widthMin)%>px;width: <%response.Write(widthMin)%>px;max-width: <%response.Write(width)%>px;"  class="<%response.Write(classe)%>" align="center"><%response.Write(conteudo)%></td>
      <%	next  
		next

		if qtd_datas=366 then

		  wrkSomaFaltas = 0
			for c= 0 to ubound(linha_nome_colunas)	

				  wrkTotalFaltas = TotalFaltas(CAMINHOn,  datas_pauta(c), nu_matricula, co_prof,  unidade,curso,etapa,turma, mat_princ, co_materia, periodo,  outro)
				  wrkSomaFaltas = wrkTotalFaltas+wrkSomaFaltas
			Next	  	
			
		if wrkSomaFaltas>0 then
			classe = "texto_vermelho"
		else
			classe = "texto_azul"				
		end if				  
		  
			  %>
      <td style="min-width: 198px;" class="<%response.Write(classe)%>" align="center"><%		  
		  response.Write(wrkSomaFaltas)%></td>
      <%end if%>
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
%>
  </tbody>
</table>
<%end function
%>
