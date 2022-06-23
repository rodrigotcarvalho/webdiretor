<!--#include file="caminhos.asp"-->
<!--#include file="funcoes.asp"-->
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
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON_wr = Server.CreateObject("ADODB.Connection") 
		ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_wr.Open ABRIR_wr		

		Set CONg = Server.CreateObject("ADODB.Connection") 
		ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONg.Open ABRIRg
%>
<div align="center">
<%if opt="c" then
prof_altera=Session("co_prof")
Session("co_prof")=prof_altera
unidade_altera=Request.Form("u_pub")
if unidade_altera="" or isnull(unidade_altera) then
session("u_pub")=session("u_pub")
else
session("u_pub")=unidade_altera
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if
%>
                      <select name="curso" class="select_style" onchange="recuperarEtapa(this.value)">
                        <option value="999990" selected></option>
                        <%		

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQLUn_Prof = "SELECT DISTINCT CO_Curso FROM TB_Da_Aula Where CO_Professor="&prof_altera&" AND NU_Unidade="& unidade_altera
		RS0.Open SQLUn_Prof, CONg
		
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
<%elseif opt="e" then
prof_altera=Session("co_prof")
Session("co_prof")=prof_altera
curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if

%>
                      <select name="etapa" class="select_style" onchange="recuperarTurma(this.value)">
                        <option value="999990" selected></option>
                        <%		
		
		
		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Da_Aula Where CO_Professor="&prof_altera&" AND NU_Unidade="& session("u_pub") &" AND CO_Curso='"&session("c_pub")&"'"
		RS0b.Open SQL0b, CONg
		
			
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
<%
'Essa combo da etapa chama também a rotina da disciplina
elseif opt="e2" then
prof_altera=Session("co_prof")
Session("co_prof")=prof_altera
curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if

%>
                      <select name="etapa" class="select_style" onchange="recuperarTurma(this.value);recuperarDisciplina(this.value)">
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
<%'Essa combo da etapa chama também a rotina da avaliação
elseif opt="e3" then
prof_altera=Session("co_prof")
Session("co_prof")=prof_altera
curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if

%>
                      <select name="etapa" class="select_style" onchange="recuperarPeriodo(this.value);recuperarAvaliacoes(this.value)">
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
<%'Essa combo da etapa chama também a rotina do Período
elseif opt="e4" then
prof_altera=Session("co_prof")
Session("co_prof")=prof_altera
curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if

%>
                      <select name="etapa" class="select_style" onchange="recuperarPeriodo()">
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
<%elseif opt="t" then
prof_altera=Session("co_prof")
Session("co_prof")=prof_altera
etapa_altera=Request.Form("e_pub")
if etapa_altera="" or isnull(etapa_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=etapa_altera
session("t_pub")=session("t_pub")
end if
%>
<select name="turma" class="select_style">
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
 <%
'Essa combo da turma chama também a submete o form 
 elseif opt="t2" then
prof_altera=Session("co_prof")
Session("co_prof")=prof_altera
etapa_altera=Request.Form("e_pub")
if etapa_altera="" or isnull(etapa_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=etapa_altera
session("t_pub")=session("t_pub")
end if
%>
<select name="turma" class="select_style" onChange="MM_callJS('submitfuncao()')">
                        <option value="999990" selected></option>
    <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Da_Aula where CO_Professor="&prof_altera&" AND NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"' AND CO_Etapa='" & session("e_pub") & "' order by CO_Turma" 
		RS3.Open SQL3, CONg						

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
  <%
'Essa combo da turma chama também a submete o form 
 elseif opt="t3" then
prof_altera=Session("co_prof")
Session("co_prof")=prof_altera
etapa_altera=Request.Form("e_pub")
if etapa_altera="" or isnull(etapa_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=etapa_altera
session("t_pub")=session("t_pub")
end if
%>
<select name="turma" class="select_style" onChange="MM_callJS('recuperarDisciplina()')">
                        <option value="999990" selected></option>
    <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Da_Aula where CO_Professor="&prof_altera&" AND NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"' AND CO_Etapa='" & session("e_pub") & "' order by CO_Turma" 
		RS3.Open SQL3, CONg						

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
<%
elseif opt="d" then
prof_altera=Session("co_prof")
Session("co_prof")=prof_altera
curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if

%>
                      <select name="mat_prin" class="select_style">
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
<%
elseif opt="d2" then
prof_altera=Session("co_prof")
Session("co_prof")=prof_altera
curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if
wrk_vetor_co_materia = ""
%>
<select name="mat_prin" class="select_style" onChange="MM_callJS('recuperarPeriodo()')">
                        <option value="999999" selected></option>
  <%                      
		Set RSG = Server.CreateObject("ADODB.Recordset")
		SQLG = "SELECT DISTINCT CO_Materia_Principal FROM TB_Da_Aula where CO_Professor ="& prof_altera&" and CO_Etapa = '"&session("e_pub") &"' AND NU_Unidade = "&session("u_pub")&" and CO_Curso = '"&session("c_pub") &"' order by CO_Materia_Principal"
		RSG.Open SQLG, CONg
		
IF RSG.EOF THEN

RESPONSE.Write("Sem disciplinas cadastradas. Procure seu Coordenador.")

ELSE
		co_mat_prin= RSG("CO_Materia_Principal")
							Set RS5a = Server.CreateObject("ADODB.Recordset")
							SQL5a = "SELECT * FROM TB_Da_Aula_Subs where CO_Professor ="& prof_altera&" AND NU_Unidade = "&session("u_pub")&" AND CO_Etapa ='"& session("e_pub") &"' AND CO_Curso ='"& session("c_pub")&"' AND UCASE(CO_Materia_Principal) ='"& co_mat_prin &"'"
							RS5a.Open SQL5a, CONg
										
							if RS5a.EOF then
								
								wrk_vetor_co_materia = co_mat_prin

							else
								conta_wrk_vetor_co_materia=0
								mat_check=""
								while not RS5a.EOF
									co_mat_sub= RS5a("CO_Materia")	
									if conta_wrk_vetor_co_materia =0  then
										wrk_vetor_co_materia = co_mat_sub
										
									else
										if mat_check<>co_mat_sub then
											wrk_vetor_co_materia = wrk_vetor_co_materia&"#!#"&co_mat_sub	
										end if									
									end if	
								mat_check = co_mat_sub	
								conta_wrk_vetor_co_materia=conta_wrk_vetor_co_materia+1	
								RS5a.MOVENEXT
								WEND	
	
							end if	
							
							split_grid = split(wrk_vetor_co_materia,"#!#")
							for s=0 to ubound(split_grid)							
								
								Set RS8 = Server.CreateObject("ADODB.Recordset")
								SQL8 = "SELECT * FROM TB_Materia where CO_Materia='"& split_grid(s) &"'"
								RS8.Open SQL8, CON0
							
								no_mat= RS8("NO_Materia")								
								
	


%>

                       <option value="<%=split_grid(s) %>"> 
                        <%response.Write(Server.URLEncode(no_mat))%>						
                        </option>
                        <%

						Next
End if						
%>
                      </select>					  
<%
elseif opt="d3" then
prof_altera=Session("co_prof")
Session("co_prof")=prof_altera
curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if

%>
                    <select name="mat_prin" class="select_style" onChange="MM_callJS('submitfuncao()')">
                        <option value="999999" selected></option> 
                        <%
		Set RSG = Server.CreateObject("ADODB.Recordset")
		SQLG = "SELECT DISTINCT CO_Materia_Principal FROM TB_Da_Aula where CO_Professor ="& prof_altera&" and CO_Etapa = '"&session("e_pub") &"' AND NU_Unidade = "&session("u_pub")&" and CO_Curso = '"&session("c_pub") &"' order by CO_Materia_Principal"
		RSG.Open SQLG, CONg

while not RSG.EOF
co_mat_prin= RSG("CO_Materia_Principal")


		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Materia where CO_Materia ='"& co_mat_prin &"'"
		RS7.Open SQL7, CON0
		
		no_mat_prin= RS7("NO_Materia")
%>
                        <option value="<%=co_mat_prin%>"> 
                        <%response.Write(Server.URLEncode(no_mat_prin))%>						
                        </option>
                        <%

RSG.MOVENEXT
WEND%>
                      </select>

<%
elseif opt="p" then
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")

%>
<select name="periodo" class="select_style" id="periodo" onChange="MM_callJS('submitfuncao()')">
                                      <option value="0" selected></option>
                                      <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
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
elseif opt="p1" then
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")

%>
<select name="periodo" class="select_style" id="periodo">
                                      <option value="0" selected></option>
                                      <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
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
elseif opt="av" then
etapa_altera=Request.Form("e_pub")
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=etapa_altera
session("t_pub")=session("t_pub")

%>																		
<select name="avaliacoes" class="select_style" id="avaliacoes" onChange="MM_callJS('submitfuncao()')">
                                      <option value="999990"></option> 
<%

		Set RSTB = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& session("u_pub") &" AND CO_Curso = '"& session("c_pub") &"' AND CO_Etapa = '"& session("e_pub")  &"'"
		Set RSTB = CONg.Execute(CONEXAO)

response.Write(CONEXAO)

nota = RSTB("TP_Nota")

		Set RSAV = Server.CreateObject("ADODB.Recordset")
		SQLAV = "SELECT * FROM Avaliacoes where CO_Escola="&escola
		RSAV.Open SQLAV, CON_wr

avaliacoes = RSAV(nota)
avaliacoes_nomes = RSAV(nota&"_Nome")

avaliacao=SPLIT(avaliacoes,"#!#")
avaliacao_nome=SPLIT(avaliacoes_nomes,"#!#")
for i=0 to UBOUND(avaliacao)
%>
                                      <option value="<%=avaliacao(i)%>"> 
                                      <%response.Write(avaliacao_nome(i))%>
                                      </option>
<%
NEXT
%>
                                    </select> 
<%
else

session("t_pub")=Request.Form("t_pub")
%>
<%end if%>
</div>
