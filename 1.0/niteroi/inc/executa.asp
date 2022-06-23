<%On Error Resume Next%>
<!--#include file="caminhos.asp"-->
<!--#include file="parametros.asp"-->
<!--#include file="funcoes6.asp"-->

<%
opt=request.querystring("opt")
nome = session("nome") 
acesso = session("acesso")
co_usr = session("co_user")
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
grupo=session("grupo")
escola=session("escola")
chave=session("chave")
escola= session("escola")
		
this_file = Request.ServerVariables("SCRIPT_NAME")
arPath = Split(this_Path, "/")

nivel=4


if nome = "" or acesso = "" or co_usr = "" or permissao = "" or ano_letivo = "" or chave = "" or isnull(chave) then
	if nivel=0 then
	response.Redirect("default.asp?opt=00")
	elseif nivel=1 then
	response.Redirect("../default.asp?opt=00")
	elseif nivel=2 then
	response.Redirect("../../default.asp?opt=00")
	elseif nivel=3 then
	response.Redirect("../../../default.asp?opt=00")
	elseif nivel=4 then
	response.Redirect("../../../../default.asp?opt=00")
	end if
end if	

		Set CON_AL = Server.CreateObject("ADODB.Connection") 
		ABRIR_AL = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_AL.Open ABRIR_AL
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON3 = Server.CreateObject("ADODB.Connection") 
		ABRIR3 = "DBQ="& CAMINHO_o & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3	
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT	

    	Set CON_WF = Server.CreateObject("ADODB.Connection") 
		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_WF.Open ABRIR_WF


		Set CONg = Server.CreateObject("ADODB.Connection") 
		ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONg.Open ABRIRg

if opt="imgtb" then

matric=Request.Form("matr_tb_pub")

vetor_fotos=Session("vetor_fotos")

	Set RSs = Server.CreateObject("ADODB.Recordset")
	SQL_s ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& matric
	Set RSs = CON_AL.Execute(SQL_s)

	nome_aluno=RSs("NO_Aluno")	

'vetor_arquivos=split(vetor_fotos,",")

for i = 0 to ubound(vetor_fotos)
nome_arquivo =vetor_fotos(i)
nome_jpg=matric&".jpg"
lowercase=lcase(nome_arquivo)
'response.Write(nome_jpg&"-"&lowercase&"-"&mostra_img)
	if nome_jpg=lowercase then
		mostra_img="OK"
	elseif mostra_img<>"OK" then
		mostra_img="NO"
	end if
Session("mostra_foto") = mostra_img	
next	

if mostra_img="OK" then

%>
<a href="#" title="<% response.Write(Server.URLEncode(nome_aluno)) %>" onClick="centraliza(500,536);MM_showHideLayers('fundo','','show','alinha','','show');mclosetime()"><img src="../../../../img/fotos/aluno/<% response.Write(matric) %>.jpg" alt="<% response.Write(Server.URLEncode(nome_aluno)) %>" width="50" height="60" border="0"></a>

<%end if%>




<%elseif opt="img" then
num_cham=Request.Form("num_cham_pub")
co_matric=Request.Form("matric_pub")
periodo_exibe=Request.Form("periodo_pub")

nom_periodo=split(periodo_exibe,"#!#")

	parametros_chamada_jscript="celula"&num_cham
	for b=1 to ubound(nom_periodo) 
		parametros_chamada_jscript=parametros_chamada_jscript&",celula"&num_cham&"p"&b
	next

	Set RSs = Server.CreateObject("ADODB.Recordset")
	SQL_s ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& co_matric
	Set RSs = CON_AL.Execute(SQL_s)

	nom_aluno=RSs("NO_Aluno")	
	
%>
   <div id="alinha" style="position:absolute; width:500px; z-index: 3; height: 536px; visibility: hidden;"> 
    <table border="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr> 
          <td height="16"> 
            <div align="right"> <span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide');focar('<%response.Write(num_cham&"c2")%>');mudar_cor_focus(<%response.Write(parametros_chamada_jscript)%>)">fechar</a>&nbsp;<a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide');focar('<%response.Write(num_cham&"c2")%>');mudar_cor_focus(<%response.Write(parametros_chamada_jscript)%>)"><img src="../../../../img/fecha.gif" width="20" height="16" border="0" align="absbottom"></a></font></span></div></td>
        </tr>
        <tr> 
          <td><div align="center" ><img src="../../../../img/fotos/aluno/<% Response.Write(co_matric) %>.jpg" height="500"></div></td>
        </tr>
        <tr>
          <td height="20">
    <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <% Response.Write(Server.URLEncode(nom_aluno)) %>
              </font></div></td>
        </tr>
      </table>
    </div>



<%elseif opt="nt" then

co_matric=Request.Form("matric_pub")
CAMINHOn=Request.Form("caminho_pub")
opcao=Request.Form("opcao_pub")
periodo=Request.Form("outro_pub")
ano_letivo=Request.Form("ano_pub")
wrk_materia=Request.Form("materia_pub")
curso=Request.Form("c_pub")
co_etapa=Request.Form("e_pub")
larg_max=Request.Form("larg_max_pub")

tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")

if periodo<>0 then
	sigla_periodo=periodos(periodo,tp_modelo,"sigla")	
end if
Set CON_N = Server.CreateObject("ADODB.Connection")
ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
CON_N.Open ABRIR3

Set RS = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& wrk_materia &"'"
RS.Open SQL, CON0

mae= RS("IN_MAE")
fil= RS("IN_FIL")
in_co= RS("IN_CO")
peso= RS("NU_Peso")

if (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) or (mae=TRUE and fil=TRUE and in_co=FALSE) or (mae=TRUE and fil=FALSE and in_co=TRUE) then

else

	Set RS1 = Server.CreateObject("ADODB.Recordset")
	SQL1 = "SELECT * FROM TB_Materia where CO_Materia ='"& wrk_materia &"' order by NU_Ordem_Boletim"
	RS1.Open SQL1, CON0

	if RS1.EOF then
		response.Write("ERRO na sele&ccedil;&atilde;o de mat&eacute;ria - Funcoes 6 - ln 965")
	else
		wrk_mat_princ=RS1("CO_Materia_Principal")
		if wrk_mat_princ="" or isnull(wrk_mat_princ) or wrk_mat_princ= " " then
			wrk_co_materia=wrk_materia
			wrk_mat_princ=wrk_materia
		else
			wrk_co_materia=wrk_materia
		end if
		
'CAMINHOn = CAMINHOn.replace(/\+/g," ")
'CAMINHOn = unescape(CAMINHOn)
'response.Write(CAMINHOn)
	
	dados_tabela=verifica_dados_tabela(CAMINHOn,opcao,periodo)
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

qtd_colunas=UBOUND(linha_nome_colunas)+1
width_num_cham="20"
width_nom_aluno="40"
width_else=(larg_max-width_num_cham-width_nom_aluno)/(qtd_colunas-2)	
%>
<table width="<%response.Write(larg_max)%>" border="0" cellspacing="0" cellpadding="0">
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
						
				Set RSpeso = Server.CreateObject("ADODB.Recordset")
				SQL_peso = "Select "&linha_pesos_variaveis(i)&" from "& tb &" WHERE CO_Matricula = "& co_matric & " AND CO_Materia_Principal = '"& wrk_mat_princ &"' AND CO_Materia = '"& wrk_co_materia &"' AND NU_Periodo="&periodo
				Set RSpeso = CON_N.Execute(SQL_peso)
					
				if RSpeso.EOF then
				else	
					valor_peso=RSpeso(""&linha_pesos_variaveis(i)&"")
				end if		
				IF comunica="s" THEN		
			'		linha_pesos(i)=valor_peso&"<input name="&nome_pesos_variaveis(i)&" type=""hidden"" id="&nome_pesos_variaveis(i)&" class=""peso"" value="&valor_peso&">"	
					linha_pesos(i)=valor_peso
				else	
			'		linha_pesos(i)="<input name="&nome_pesos_variaveis(i)&" type="&tipo&" id="&nome_pesos_variaveis(i)&" class=""peso"" value="&valor_peso&">"
					linha_pesos(i)=valor_peso
				end if	
			
		end if				
 %>
    <td width="<%response.Write(width)%>" class="<%response.Write(classe_subtit)%>"><div align="<%response.Write(align)%>"><%response.Write(linha_pesos(i))%></div></td>
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
	if linha_nome_colunas(j)="N&ordm;" then
		cabecalho="Per"
	elseif linha_nome_colunas(j)="Nome" then
		cabecalho="Disciplina"
	else
		cabecalho=linha_nome_colunas(j)
	end if	
 %>
    <td width="<%response.Write(width)%>" class="<%response.Write(classe_subtit)%>"><div align="<%response.Write(align)%>"><%response.Write(cabecalho)%></div></td>
<%	next%>  
  </tr>
  <%
	if periodo = 0 then
	qtd_cols_this_if=ubound(linha_nome_colunas)+1
	%>
        <tr>
            <td width="<%response.Write(width)%>" colspan="<%response.Write(qtd_cols_this_if)%>" class="<%response.Write(classe)%>"><div align="center">N&atilde;o Dispon&iacute;vel para esse per&iacute;odo.</div></td>                                      						
         </tr>
	<%
    else
		Set RSs = Server.CreateObject("ADODB.Recordset")
		SQL_s ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& co_matric&" and TB_Matriculas.NU_Ano="&ano_letivo
		Set RSs = CON_AL.Execute(SQL_s)
	
		if RSs.EOF then
		%>
				<tr>
					<td width="<%response.Write(width)%>"><div align="center"><%response.Write(sigla_periodo)%></div></td>
					<td width="<%response.Write(width)%>"><%response.Write(wrk_co_materia)%></td>                                      						
					<%for m= 0 to ubound(nome_variaveis)
						width=width_else
						align="center"
				 %>
						<td width="<%response.Write(width)%>" class="<%response.Write(classe)%>">&nbsp;</td>
					 <%next%>
				  </tr>
		 <%else		
			%>   
				<tr>
					<td width="<%response.Write(width)%>"><div align="center"><%response.Write(sigla_periodo)%></div></td>
					<td width="<%response.Write(width)%>" ><%response.Write(wrk_co_materia)%></td>               
					 <% 
					Set RS3 = Server.CreateObject("ADODB.Recordset")
					SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& co_matric & " AND CO_Materia_Principal = '"& wrk_mat_princ &"' AND CO_Materia = '"& wrk_co_materia &"' AND NU_Periodo="&periodo
					Set RS3 = CON_N.Execute(SQL_N)			 
					coluna=0	 
					 for n= 0 to ubound(nome_variaveis)
						width=width_else
						align="center"
						if RS3.EOF then 
							valor=""
						else
							if variaveis_bd(n)="CALCULADO" then
								valor="&nbsp;"
								'Nesse caso o valor é calculado pela função calcular_nota chamada mais abaixo
							else
								valor=RS3(""&variaveis_bd(n)&"")
							end if					
						end if
						
						if (valor="" or isnull(valor)) then
							coluna=coluna+1	
							conteudo="&nbsp;"			
						else
							if nome_variaveis(n)="faltas" then
								coluna=coluna+1
								conteudo=formatnumber(valor,0)		
							elseif nome_variaveis(n)="nav" or nome_variaveis(n)="media_teste" or nome_variaveis(n)="media_prova" or nome_variaveis(n)="media1" or nome_variaveis(n)="media2" or nome_variaveis(n)="media3" or situac<>"C" then
								coluna=coluna	
								conteudo=formatnumber(valor,1)
							elseif calcula_variavel(n)="CALC1" and situac="C" then
								coluna=coluna
								valor=calcular_nota(calcula_variavel(n),CAMINHOn,tb,nu_matricula,mat_princ,co_materia,periodo)
								conteudo=formatnumber(valor,1)	
							else
								coluna=coluna+1
								conteudo=formatnumber(valor,1)			
							end if	
						end if	
						'conteudo=n
				 %>
					<td width="<%response.Write(width)%>">
						<div align="<%response.Write(align)%>">
							<%response.Write(conteudo)%> 
						</div>
					 </td>
				  <%	next  
				  %>
				  </tr>
				  <%			
			end if	
		end if	
	END IF	
END IF	
%>     
</td>
</tr>
</table>
<%elseif opt="ntzoom" then

co_matric=Request.Form("matric_pub")
CAMINHOn=Request.Form("caminho_pub2")
opcao=Request.Form("opcao_pub")
periodo=Request.Form("outro_pub")
ano_letivo=Request.Form("ano_pub")
wrk_materia=Request.Form("materia_pub")
curso=Request.Form("c_pub")
co_etapa=Request.Form("e_pub")
larg_max=Request.Form("larg_max_pub")
CAMINHOn=replace(CAMINHOn,"$b$","\")
CAMINHOn=replace(CAMINHOn,"$u$","_")

tp_modelo=tipo_divisao_ano(curso,co_etapa,"tp_modelo")

if periodo<>0 then
	sigla_periodo=periodos(periodo,tp_modelo,"sigla")	
end if
Set CON_N = Server.CreateObject("ADODB.Connection")
ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
CON_N.Open ABRIR3

Set RS = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& wrk_materia &"'"
RS.Open SQL, CON0

mae= RS("IN_MAE")
fil= RS("IN_FIL")
in_co= RS("IN_CO")
peso= RS("NU_Peso")



if (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) or (mae=TRUE and fil=TRUE and in_co=FALSE) or (mae=TRUE and fil=FALSE and in_co=TRUE) then

else

	Set RS1 = Server.CreateObject("ADODB.Recordset")
	SQL1 = "SELECT * FROM TB_Materia where CO_Materia ='"& wrk_materia &"' order by NU_Ordem_Boletim"
	RS1.Open SQL1, CON0
	
	if RS1.EOF then
		response.Write("ERRO na sele&ccedil;&atilde;o de mat&eacute;ria - Funcoes 6 - ln 965")
	else
		wrk_mat_princ=RS1("CO_Materia_Principal")
		if wrk_mat_princ="" or isnull(wrk_mat_princ) or wrk_mat_princ= " " then
			wrk_co_materia=wrk_materia
			wrk_mat_princ=wrk_materia
		else
			wrk_co_materia=wrk_materia
		end if
		

	dados_tabela=verifica_dados_tabela(CAMINHOn,opcao,periodo)
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
	
qtd_colunas=UBOUND(linha_nome_colunas)+1
width_num_cham="20"
width_nom_aluno="40"
width_else=(larg_max-width_num_cham-width_nom_aluno)/(qtd_colunas-2)	
%>
<table width="<%response.Write(larg_max)%>" border="0" cellspacing="0" cellpadding="0">
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
						
				Set RSpeso = Server.CreateObject("ADODB.Recordset")
				SQL_peso = "Select "&linha_pesos_variaveis(i)&" from "& tb &" WHERE CO_Matricula = "& co_matric & " AND CO_Materia_Principal = '"& wrk_mat_princ &"' AND CO_Materia = '"& wrk_co_materia &"' AND NU_Periodo="&periodo
				Set RSpeso = CON_N.Execute(SQL_peso)
					
				if RSpeso.EOF then
				else	
					valor_peso=RSpeso(""&linha_pesos_variaveis(i)&"")
				end if		
				IF comunica="s" THEN		
'					linha_pesos(i)=valor_peso&"<input name="&nome_pesos_variaveis(i)&" type=""hidden"" id="&nome_pesos_variaveis(i)&" class=""peso"" value="&valor_peso&">"	
					linha_pesos(i)=valor_peso
				else	
				'	linha_pesos(i)="<input name="&nome_pesos_variaveis(i)&" type="&tipo&" id="&nome_pesos_variaveis(i)&" class=""peso"" value="&valor_peso&">"
					linha_pesos(i)=valor_peso
				end if	
			
		end if				
 %>
    <td width="<%response.Write(width)%>" class="zoom_label"><div align="<%response.Write(align)%>"><%response.Write(linha_pesos(i))%></div></td>
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
	if linha_nome_colunas(j)="N&ordm;" then
		cabecalho="Per"
	elseif linha_nome_colunas(j)="Nome" then
		cabecalho="Disciplina"
	else
		cabecalho=linha_nome_colunas(j)
	end if	
 %>
    <td width="<%response.Write(width)%>" class="zoom_tit"><div align="<%response.Write(align)%>"><%response.Write(cabecalho)%></div></td>
<%	next%>  
  </tr>
  <%
	if periodo = 0 then
	qtd_cols_this_if=ubound(linha_nome_colunas)+1
	%>
        <tr>
            <td width="<%response.Write(width)%>" colspan="<%response.Write(qtd_cols_this_if)%>" class="zoom_texto"><div align="center">N&atilde;o Dispon&iacute;vel para esse per&iacute;odo.</div></td>                                      						
         </tr>
	<%
    else
		Set RSs = Server.CreateObject("ADODB.Recordset")
		SQL_s ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& co_matric&" and TB_Matriculas.NU_Ano="&ano_letivo
		Set RSs = CON_AL.Execute(SQL_s)
	
		if RSs.EOF then
		%>
				<tr>
					<td width="<%response.Write(width)%>" class="zoom_texto"><div align="center"><%response.Write(sigla_periodo)%></div></td>
					<td width="<%response.Write(width)%>" class="zoom_texto"><%response.Write(wrk_co_materia)%></td>                                      						
					<%for m= 0 to ubound(nome_variaveis)
						width=width_else
						align="center"
				 %>
						<td width="<%response.Write(width)%>" class="zoom_texto">&nbsp;</td>
					 <%next%>
				  </tr>
		 <%else		
			%>   
				<tr>
					<td width="<%response.Write(width)%>" class="zoom_texto"><div align="center"><%response.Write(sigla_periodo)%></div></td>
					<td width="<%response.Write(width)%>" class="zoom_texto"><%response.Write(wrk_co_materia)%></td>               
					 <% 
					Set RS3 = Server.CreateObject("ADODB.Recordset")
					SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& co_matric & " AND CO_Materia_Principal = '"& wrk_mat_princ &"' AND CO_Materia = '"& wrk_co_materia &"' AND NU_Periodo="&periodo
					Set RS3 = CON_N.Execute(SQL_N)			 
					coluna=0	 
					 for n= 0 to ubound(nome_variaveis)
						width=width_else
						align="center"
						if RS3.EOF then 
							valor=""
						else
							if variaveis_bd(n)="CALCULADO" then
								valor="&nbsp;"
								'Nesse caso o valor é calculado pela função calcular_nota chamada mais abaixo
							else
								valor=RS3(""&variaveis_bd(n)&"")
							end if					
						end if
						
						if (valor="" or isnull(valor)) then
							coluna=coluna+1	
							conteudo="&nbsp;"			
						else
							if nome_variaveis(n)="faltas" then
								coluna=coluna+1
								conteudo=formatnumber(valor,0)		
							elseif nome_variaveis(n)="nav" or nome_variaveis(n)="media_teste" or nome_variaveis(n)="media_prova" or nome_variaveis(n)="media1" or nome_variaveis(n)="media2" or nome_variaveis(n)="media3" or situac<>"C" then
								coluna=coluna	
								conteudo=formatnumber(valor,1)
							elseif calcula_variavel(n)="CALC1" and situac="C" then
								coluna=coluna
								valor=calcular_nota(calcula_variavel(n),CAMINHOn,tb,nu_matricula,mat_princ,co_materia,periodo)
								conteudo=formatnumber(valor,1)	
							else
								coluna=coluna+1
								conteudo=formatnumber(valor,1)			
							end if	
						end if	
						'conteudo=n
				 %>
					<td width="<%response.Write(width)%>" class="zoom_texto">
						<div align="<%response.Write(align)%>">
							<%response.Write(conteudo)%> 
						</div>
					 </td>
				  <%	next  
				  %>
				  </tr>
				  <%			
			end if	
		end if	
	END IF	
END IF	
%>     
</td>
</tr>
</table>
<%elseif opt="ocr" then
co_matric=Request.Form("matric_pub")
		
	Set RS4 = Server.CreateObject("ADODB.Recordset")
	SQL4 = "select * from TB_Ocorrencia_Aluno where CO_Matricula = " & co_matric &" Order BY DA_Ocorrencia DESC,HO_Ocorrencia"
	set RS4 = CON3.Execute (SQL4)
	if RS4.EOF then
	%>
	<div align="left"><%Response.Write(Server.URLEncode("&nbsp;&nbsp;Última Ocorrencia Registrada:&nbsp;Sem Ocorrências"))%></div>
	<%else
	data_ocor=RS4("DA_Ocorrencia")
	ho_ocor=RS4("HO_Ocorrencia")
	assunto=RS4("CO_Assunto")
	co_ocr=RS4("CO_Ocorrencia")
	prof=RS4("CO_Professor")
	mat_ocr=RS4("NO_Materia")
	obs=RS4("TX_Observa")	
	
	Set RSto = Server.CreateObject("ADODB.Recordset")
	SQLto = "SELECT * FROM TB_Tipo_Ocorrencia Where CO_Ocorrencia="&co_ocr
	RSto.Open SQLto, CON0	
	no_ocorrencia=RSto("NO_Ocorrencia")
	%>
	<div align="left"><%Response.Write(Server.URLEncode("&nbsp;&nbsp;Última Ocorrencia Registrada:&nbsp;"&data_ocor&", referente a: "&no_ocorrencia))%></div>
	<%end if

elseif opt="nm" then
co_matric=Request.Form("matric_pub")
		
	Set RSs = Server.CreateObject("ADODB.Recordset")
	SQL_s ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& co_matric&" and TB_Matriculas.NU_Ano="&ano_letivo

	Set RSs = CON_AL.Execute(SQL_s)


	nome_aluno=RSs("NO_Aluno")
	
	if RSs.EOF then
	%>
	<div align="left"><%Response.Write(Server.URLEncode("executa.asp>1200-Erro na busca do nome"))%></div>
	<%else
		Set RSCONTA = Server.CreateObject("ADODB.Recordset")
		SQLA = "SELECT * FROM TB_Contatos WHERE TP_Contato ='ALUNO' And CO_Matricula ="& co_matric
		RSCONTA.Open SQLA, CONCONT

		if RSCONTA.EOF then
			nascimento="0/0/0"
		else
			nascimento = RSCONTA("DA_Nascimento_Contato")
		end if
			vetor_nascimento = Split(nascimento,"/")  
			dia = vetor_nascimento(0)
			mes = vetor_nascimento(1)
			ano = vetor_nascimento(2)
			
			data= dia&"-"&mes&"-"&ano
			intervalo = DateDiff("d", data , now )
			
			idade = int(intervalo/365.25)

	%>
	<div align="left"><%Response.Write("&nbsp;&nbsp;"&Server.URLEncode(nome_aluno)&"&nbsp;-&nbsp;"&idade&" anos")%></div>
	<%end if
	
elseif opt="ctrl" then
session("u_pub")=session("u_pub")
session("c_pub")=Request.Form("c_pub")

			

		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")&" AND CO_Curso='"&session("c_pub")&"'"
		'response.Write(SQL0b&"<BR>")
		RS0b.Open SQL0b, CON0
				
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<%
While not RS0b.EOF
co_etapa = RS0b("CO_Etapa")

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("c_pub")&"' AND CO_Etapa='"&co_etapa&"'"
		'response.Write(SQL0c&"<BR>")		
		RS0c.Open SQL0c, CON0
		
		no_etapa = RS0c("NO_Etapa")	
		tp_modelo = RS0c("TP_Modelo")		
	
	
		Set RS_WF = Server.CreateObject("ADODB.Recordset")
		SQL_WF = "SELECT * FROM TB_Autoriza_WF WHERE NU_Unidade="&session("u_pub")&" AND CO_Curso='"&session("c_pub")&"' and CO_Etapa='"&co_etapa&"'"
		'response.Write(SQL_WF&"<BR>")				
		RS_WF.Open SQL_WF, CON_WF
					
		co_apr1=RS_WF("CO_apr1")
		co_apr2=RS_WF("CO_apr2")
		co_apr3=RS_WF("CO_apr3")
		co_apr4=RS_WF("CO_apr4")
		co_apr5=RS_WF("CO_apr5")
		co_apr6=RS_WF("CO_apr6")
		co_apr7=RS_WF("CO_apr7")					
		co_prova1=RS_WF("CO_prova1")
		co_prova2=RS_WF("CO_prova2")
		co_prova3=RS_WF("CO_prova3")
		co_prova4=RS_WF("CO_prova4")
		co_prova5=RS_WF("CO_prova5")		
		co_prova6=RS_WF("CO_prova6")	
		co_prova7=RS_WF("CO_prova7")
	
	
%>
<tr><td class="tb_tit"><%response.Write(Server.URLEncode(no_etapa))%></td></tr>	
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                    <% 
					
					Set RSP = Server.CreateObject("ADODB.Recordset")
					SQLP = "SELECT Distinct(NU_Periodo),NO_Periodo FROM TB_Periodo where TP_Modelo='"&tp_modelo&"' order by NU_Periodo"
		'response.Write(SQLP&"<BR>")						
					RSP.Open SQLP, CON0
							
					conta_periodo=0
					while not RSP.EOF				
						nome_periodo=RSP("NO_Periodo")
					%>
                      <td colspan="2" class="tb_subtit"><div align="center"><%response.Write(Server.URLEncode(nome_periodo))%></div></td>
					<%
						conta_periodo=conta_periodo+1
						RSP.MOVENEXT
					WEND
					%>
						
                      </tr>
                    <tr>
                    <% for cp=1 to conta_periodo %>
                      <td width="71" class="tb_subtit"><div align="center">Testes 
                      </div></td>
                      <td width="72" class="tb_subtit"><div align="center">Provas 
                      </div></td>
					<%next%>
                      </tr>
                    <tr>
                    <% 
										
					for cp=1 to conta_periodo 	
										
						if cp=1 then						
							if co_apr1="L" then
								apr_checked_lib ="checked"
								apr_checked_blq =""														
							else
								apr_checked_lib =""
								apr_checked_blq ="checked"																	
							end if
							if co_prova1="L" then
								pr_checked_lib ="checked"
								pr_checked_blq =""														
							else
								pr_checked_lib =""
								pr_checked_blq ="checked"																	
							end if			
											
						elseif cp=2 then
							if co_apr2="L" then
								apr_checked_lib	="checked"
								apr_checked_blq	=""														
							else
								apr_checked_lib	=""
								apr_checked_blq	="checked"																	
							end if
							if co_prova2="L" then
								pr_checked_lib	="checked"
								pr_checked_blq	=""														
							else
								pr_checked_lib	=""
								pr_checked_blq	="checked"																	
							end if												
						elseif cp=3 then
							if co_apr3="L" then
								apr_checked_lib	="checked"
								apr_checked_blq	=""														
							else
								apr_checked_lib	=""
								apr_checked_blq	="checked"																	
							end if
							if co_prova3="L" then
								pr_checked_lib	="checked"
								pr_checked_blq	=""														
							else
								pr_checked_lib	=""
								pr_checked_blq	="checked"																	
							end if	
					
						elseif cp=4 then
							if co_apr4="L" then
								apr_checked_lib	="checked"
								apr_checked_blq	=""														
							else
								apr_checked_lib	=""
								apr_checked_blq	="checked"																	
							end if
							if co_prova4="L" then
								pr_checked_lib	="checked"
								pr_checked_blq	=""														
							else
								pr_checked_lib	=""
								pr_checked_blq	="checked"																	
							end if					
				
						elseif cp=5 then
							if co_apr5="L" then
								apr_checked_lib	="checked"
								apr_checked_blq	=""														
							else
								apr_checked_lib	=""
								apr_checked_blq	="checked"																	
							end if
							if co_prova5="L" then
								pr_checked_lib	="checked"
								pr_checked_blq	=""														
							else
								pr_checked_lib	=""
								pr_checked_blq	="checked"																	
							end if						
						elseif cp=6 then
							if co_apr6="L" then
								apr_checked_lib	="checked"
								apr_checked_blq	=""														
							else
								apr_checked_lib	=""
								apr_checked_blq	="checked"																	
							end if
							if co_prova6="L" then
								pr_checked_lib	="checked"
								pr_checked_blq	=""														
							else
								pr_checked_lib	=""
								pr_checked_blq	="checked"																	
							end if		
						elseif cp=7 then
							if co_apr7="L" then
								apr_checked_lib	="checked"
								apr_checked_blq	=""														
							else
								apr_checked_lib	=""
								apr_checked_blq	="checked"																	
							end if
							if co_prova7="L" then
								pr_checked_lib	="checked"
								pr_checked_blq	=""														
							else
								pr_checked_lib	=""
								pr_checked_blq	="checked"																	
							end if	
						end if												
					%>                    
                      <td width="71"><div align="center">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td>
                            <input type="radio" name="<%response.Write(co_etapa&"_t_"&cp)%>" id="<%response.Write(co_etapa&"_t_"&cp)%>" value="D" class="borda" <%response.Write(apr_checked_blq)%> onclick="GravaControle('<%response.Write(co_etapa&"_t_"&cp)%>',this.value)"/></td>
                            <td class="form_dado_texto">Bloq</td>
                          </tr>
                          <tr>
                            <td><input type="radio" name="<%response.Write(co_etapa&"_t_"&cp)%>" id="<%response.Write(co_etapa&"_t_"&cp)%>" value="L" class="borda" <%response.Write(apr_checked_lib)%> onclick="GravaControle('<%response.Write(co_etapa&"_t_"&cp)%>',this.value)"/></td>
                            <td class="form_dado_texto">Lib</td>
                          </tr>
                        </table>
                      </div></td>
                      <td width="72"><div align="center">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td><input type="radio" name="<%response.Write(co_etapa&"_p_"&cp)%>" id="<%response.Write(co_etapa&"_p_"&cp)%>" value="D" class="borda" <%response.Write(pr_checked_blq)%> onclick="GravaControle('<%response.Write(co_etapa&"_p_"&cp)%>',this.value)"/></td>
                            <td class="form_dado_texto">Bloq</td>
                          </tr>
                          <tr>
                            <td><input type="radio" name="<%response.Write(co_etapa&"_p_"&cp)%>" id="<%response.Write(co_etapa&"_p_"&cp)%>" value="L" class="borda" <%response.Write(pr_checked_lib)%> onclick="GravaControle('<%response.Write(co_etapa&"_p_"&cp)%>',this.value)"/></td>
                            <td class="form_dado_texto">Lib</td>
                          </tr>
                        </table>
                      </div></td>                   
                      <%
					  next
					  %>

                    </tr>
                    
                  </table>               
                  </td>
  </tr>	
<%
RS0b.MOVENEXT
WEND
%>    
</table>

<%


elseif Left(opt, 1)="c" then
	unidade_altera=Request.Form("u_pub")
	if unidade_altera="" or isnull(unidade_altera) then
		session("u_pub")=session("u_pub")
	else
		session("u_aoc")=unidade_altera	
		session("u_pub")=unidade_altera
		session("c_pub")=session("c_pub")
		session("e_pub")=session("e_pub")
		session("t_pub")=session("t_pub")
	end if
	if opt="c" then
		onchange="onchange=""recuperarEtapa(this.value)"""
	elseif opt="c2" then
		onchange=""	
	elseif opt="c3" then
		onchange="onchange=""recuperarControle(this.value)"""	
	end if	
%>
                      <select name="curso" class="select_style" id="curso" <%response.Write(onchange)%>>
                        <option value="999990" selected></option>
                        <%		

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT DISTINCT CO_Curso FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")
		RS0.Open SQL0, CON0
		
		
While not RS0.EOF
CO_Curso = RS0("CO_Curso")

		Set RS0a = Server.CreateObject("ADODB.Recordset")
		SQL0a = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
		RS0a.Open SQL0a, CON0
		
NO_Curso = RS0a("NO_Abreviado_Curso")		
%>
                        <option value="<%response.Write(CO_Curso)%>"> 
                        <%response.Write(Server.URLEncode(NO_Curso))%>
                        </option>
                        <%
RS0.MOVENEXT
WEND
%>
                      </select>
               
<%elseif Left(opt, 1)="e" then
	
	curso_altera=Request.Form("c_pub")
	if curso_altera="" or isnull(curso_altera) then
		session("c_pub")=session("c_pub")
	else
		session("u_pub")=session("u_pub")
		session("c_pub")=curso_altera
		session("e_pub")=session("e_pub")
		session("t_pub")=session("t_pub")
	end if
	if opt="e" then
		onchange="onchange=""recuperarTurma(this.value)"""
	elseif opt="e2" then
		onchange="onchange=""recuperarTurma(this.value);recuperarDisciplina(this.value)"""
	elseif opt="e3" then
		onchange="onchange=""recuperarPeriodo(this.value);recuperarAvaliacoes(this.value)"""	
	elseif opt="e4" then		
		onchange="onchange=""recuperarPeriodo(this.value)"""
	elseif opt="e5" then		
		onchange="onchange=""recuperarDisciplina(curso.value,this.value)"""
	elseif opt="e6" then		
		onchange="onchange=""MM_callJS('submitfuncao()')"""	
	elseif opt="e7" then		
		onchange="onchange=""recuperarTurma(this.value),recuperarPeriodo(this.value)"""		
	elseif opt="e8" then				
		onchange="onchange=""recuperarTurma(this.value),recuperarPeriodo(this.value)"""	
	elseif opt="e9" then				
		onchange=""	
	end if	
%>
                      <select name="etapa" class="select_style" <%response.Write(onchange)%>>
                        <option value="999990" selected></option>
                        <%		
		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"'"
		RS0b.Open SQL0b, CON0
				
While not RS0b.EOF
CO_Etapa = RS0b("CO_Etapa")

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("c_pub")&"' AND CO_Etapa='"&CO_Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
%>
                        <option value="<%response.Write(CO_Etapa)%>"> 
                        <%response.Write(Server.URLEncode(NO_Etapa))%>
                        </option>
                        <%
RS0b.MOVENEXT
WEND
%>
                      </select>

<%elseif Left(opt, 1)="t" then

	etapa_altera=Request.Form("e_pub")
	if etapa_altera="" or isnull(etapa_altera) then
		session("c_pub")=session("c_pub")
	else
		session("u_pub")=session("u_pub")
		session("c_pub")=session("c_pub")
		session("e_pub")=etapa_altera
		session("e_aoc")=etapa_altera
		session("t_pub")=session("t_pub")
	end if
	if opt="t" then
		onchange=""
	elseif opt="t2" then
		onchange="onchange=""MM_callJS('submitfuncao()')"""
	elseif opt="t3" then
		onchange="onchange=""MM_callJS('recuperarDisciplina()')"""	
	elseif opt="t4" then		
		onchange="onchange=""recuperarPeriodo(this.value)"""
	elseif opt="t5" then		
		onchange="onchange=""gravarTurma(this.value)"""
	elseif opt="t6" then		
		' Só serve para gravar o valor de turma para a função atualizar ocorrências (AOC)
		turma_aoc=Request.Form("t_pub")
		session("t_aoc")=turma_aoc	
	end if	

	if opt<>"t6" then
%>
    <select name="turma" class="select_style" <%response.Write(onchange)%>>
                            <option value="999990" selected></option>
        <%
            Set RS3 = Server.CreateObject("ADODB.Recordset")
            SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Turma where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"' AND CO_Etapa='" & session("e_pub") & "' order by CO_Turma" 
            RS3.Open SQL3, CON0						
    
    while not RS3.EOF
    co_turma= RS3("CO_Turma")
     %>
    <option value="<%=co_turma%>"> 
        <%response.Write(Server.URLEncode(co_turma))%>
    </option>
        <%
    RS3.MOVENEXT
    WEND
    %>
        
     </select>
<%	end if
elseif Left(opt, 1)="d" then

	curso_altera=Request.Form("c_pub")
	if curso_altera="" or isnull(curso_altera) then
		session("c_pub")=session("c_pub")
	else
		session("u_pub")=session("u_pub")
		session("c_pub")=curso_altera
		session("e_pub")=session("e_pub")
		session("t_pub")=session("t_pub")
	end if

	if opt="d" then
		onchange=""	
		combo="s"
	elseif opt="d2" then
		prof_altera=Session("co_prof")
		onchange="onchange=""MM_callJS('recuperarPeriodo()')"""
	elseif opt="d3" then
		onchange="onchange=""MM_callJS('recuperarPeriodo()')"""	
		combo="s"		
	elseif opt="d4" then
		onchange="onchange=""MM_callJS('submitfuncao()')"""	
		combo="s"
	end if	

	if opt="d2" then
		Set RSG = Server.CreateObject("ADODB.Recordset")
		SQLG = "SELECT DISTINCT CO_Materia_Principal FROM TB_Da_Aula where CO_Professor ="& prof_altera&" and CO_Etapa = '"&session("e_pub") &"' AND NU_Unidade = "&session("u_pub")&" and CO_Curso = '"&session("c_pub") &"' order by CO_Materia_Principal"
		RSG.Open SQLG, CONg
		
		IF RSG.EOF THEN	
			RESPONSE.Write("Sem disciplinas cadastradas. Procure seu Coordenador.")
			combo="n"
		else
			combo="s"
		end if
	end if
	
	if combo="s" then
%>
                      <select name="mat_prin" class="select_style" <%response.Write(onchange)%>>
                        <option value="999999" selected></option>
                        <%
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& session("e_pub") &"' AND CO_Curso ='"& session("c_pub") &"' order by NU_Ordem_Boletim "
		RS5.Open SQL5, CON0

while not RS5.EOF
co_mat_prin= RS5("CO_Materia")


		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Materia where CO_Materia ='"& co_mat_prin &"'"
		RS7.Open SQL7, CON0
		
		no_mat_prin= RS7("NO_Materia")
%>
                        <option value="<%=co_mat_prin%>"> 
                        <%response.Write(Server.URLEncode(no_mat_prin))%>						
                        </option>
                        <%

			RS5.MOVENEXT
			WEND%>
                      </select>


<%	end if

elseif Left(opt, 1)="p" then
	etapa_altera=Request.Form("e_pub")
	turma_altera=Request.Form("t_pub")	
	if etapa_altera="" or isnull(etapa_altera) then
		session("e_pub")=session("e_pub")
	else
		session("e_pub")=etapa_altera
	END IF
	if turma_altera="" or isnull(turma_altera) then
		session("t_pub")=session("t_pub")
	else
		session("t_pub")=turma_altera
	END IF	
	session("u_pub")=session("u_pub")
	session("c_pub")=session("c_pub")

	if opt="p" then
		onchange="onchange=""MM_callJS('submitfuncao()')"""	
	elseif opt="p1" then
		onchange="onchange=""GuardaPeriodo(this.value)"""	
	elseif opt="p2" then
		onchange="onchange=""MM_callJS('recuperarAvaliacoes()')"""		
	elseif opt="p3" then
		onchange="onchange=""MM_callJS('recuperarMedia()')"""	
	end if	
		tp_modelo=tipo_divisao_ano(session("c_pub"),session("e_pub"),"tp_modelo")	
%>
<select name="periodo" class="select_style" id="periodo" <%response.Write(onchange)%>>
                                      <option value="0" selected></option>
                                      <%
									  
									  
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo where TP_Modelo='"&tp_modelo&"' order by NU_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")%>
                                      <option value="<%=NU_Periodo%>"> 
                                      <%response.Write(Server.URLEncode(NO_Periodo))%>
                                      </option>
                                      <%RS4.MOVENEXT
WEND%>
                                    </select>
									
<%
elseif Left(opt, 2)="av" then
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")

	onchange="onchange=""MM_callJS('submitfuncao()')"""	

%>																		
<select name="avaliacoes" class="select_style" id="avaliacoes" <%response.Write(onchange)%>>
                                      <option value="999990"></option>
<%	

		Set RSTB = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& session("u_pub") &" AND CO_Curso = '"& session("c_pub") &"' AND CO_Etapa = '"& session("e_pub")  &"'"
		Set RSTB = CONg.Execute(CONEXAO)

nota = RSTB("TP_Nota")

if nota ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na
opcao="A"
else
	if nota="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb
	opcao="B"
	else
		if nota ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc
		opcao="C"
		else
			if nota ="TB_NOTA_D" then
			CAMINHOn = CAMINHO_nd
			opcao="D"
			else
				if nota ="TB_NOTA_E" then
				CAMINHOn = CAMINHO_ne
				opcao="E"
				else
				response.Write("ERRO")
				End if
			End if
		End if
	end if
end if	
	ln_nom_cols=verifica_dados_tabela(opcao,"nome_cols",outro)	
	nm_vars=verifica_dados_tabela(opcao,"wrk_var",outro)
	nm_bd=verifica_dados_tabela(opcao,"bd_var",outro)	
	avaliacoes_nomes=split(ln_nom_cols,"#!#")
	verifica_avaliacoes=split(nm_vars,"#!#")
	avaliacoes=split(nm_bd,"#!#")

for i=3 to UBOUND(avaliacoes_nomes)
	j=i-2
	if avaliacoes(j)="CALCULADO" or verifica_avaliacoes(j)="media_teste" or verifica_avaliacoes(j)="media_prova" or verifica_avaliacoes(j)="media1" or verifica_avaliacoes(j)="media2" or verifica_avaliacoes(j)="media3" then
	else
%>
                                      <option value="<%response.Write(avaliacoes(j))%>"> 
                                      <%response.Write(avaliacoes_nomes(i))%>
                                      </option>
<%
	end if
NEXT
%>
                                    </select> 
<%
elseif opt="mi" then
'etapa_altera=Request.Form("e_pub")
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")

%>
<select name="mediainformada" class="select_style" id="mediainformada" onChange="MM_callJS('submitfuncao()')">
<option value="0" selected></option>
<%i=0
while i<101 %>
                                      <option value="<%=i%>"> 
                                      <%response.Write(i)%>
                                      </option>
<%i=i+5
WEND%>
                                    </select>	

<%
else

session("t_pub")=Request.Form("t_pub")
end if%>
<%If Err.number<>0 then
errnumb = Err.number
errdesc = Err.Description
lsPath = Request.ServerVariables("SCRIPT_NAME")
arPath = Split(lsPath, "/")
GetFileName =arPath(UBound(arPath,1))
passos = 0
for way=0 to UBound(arPath,1)
passos=passos+1
next
seleciona1=passos-2
pasta=arPath(seleciona1)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
'response.redirect("../../../../inc/erro.asp")
end if
%>
