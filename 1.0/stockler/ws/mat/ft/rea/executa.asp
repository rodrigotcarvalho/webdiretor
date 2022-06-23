<!--#include file="../../../../inc/connect_pr.asp"-->
<!--#include file="../../../../inc/connect_al.asp"-->
<!--#include file="../../../../inc/connect_a.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%
opt=request.querystring("opt")
ori=request.querystring("ori")

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
	
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1		
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2

if opt="c" then
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
<select name="curso" class="borda" onchange="recuperarEtapa(this.value)">
 <option value="999990" selected></option>
                        <%		

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")
		RS0.Open SQL0, CON0
		
	
CO_Curso_check="999999"		
While not RS0.EOF
CO_Curso = RS0("CO_Curso")

if CO_Curso = CO_Curso_check then
RS0.MOVENEXT		
else

		Set RS0a = Server.CreateObject("ADODB.Recordset")
		SQL0a = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
		RS0a.Open SQL0a, CON0
		
NO_Curso = RS0a("NO_Curso")

if CO_Curso = session("c_ck") and ori="load" then
%>
                        <option value="<%response.Write(Server.URLEncode(CO_Curso))%>" selected> 
                        <%response.Write(Server.URLEncode(NO_Curso))%>
                        </option>
                        <%
CO_Curso_check = CO_Curso
RS0.MOVENEXT
else								
%>
                        <option value="<%response.Write(Server.URLEncode(CO_Curso))%>"> 
                        <%response.Write(Server.URLEncode(NO_Curso))%>
                        </option>
                        <%

CO_Curso_check = CO_Curso
RS0.MOVENEXT
end if
end if
WEND
%>
</select>
<%call GeraNomes("PORT",session("u_pub"),1,1,CON0)
no_unidade_alterada = session("no_unidades")
%>
<input name="unidade_alterada" type="hidden" id="unidade_alterada" value="<%response.Write(Server.URLEncode(no_unidade_alterada))%>">
<%elseif opt="e" then
curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if

%><select name="etapa" class="borda" onchange="recuperarTurma(this.value)">
 <option value="999990" selected></option>
                        <%		

		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"'"
		RS0b.Open SQL0b, CON0
		
	
CO_Etapa_check="999999"		
While not RS0b.EOF
CO_Etapa = RS0b("CO_Etapa")

if CO_Etapa = CO_Etapa_check then
RS0b.MOVENEXT		
else

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("c_pub")&"' AND CO_Etapa='"&CO_Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")
if CO_Etapa = session("e_ck") and ori="load" then
%>
                        <option value="<%response.Write(Server.URLEncode(CO_Etapa))%>" selected> 
                        <%response.Write(Server.URLEncode(NO_Etapa))%>
                        </option>
                        <%
CO_Etapa_check = CO_Etapa
RS0b.MOVENEXT
else								
%>
                        <option value="<%response.Write(Server.URLEncode(CO_Etapa))%>"> 
                        <%response.Write(Server.URLEncode(NO_Etapa))%>
                        </option>
                        <%

CO_Etapa_check = CO_Etapa
RS0b.MOVENEXT
end if
end if
WEND
%>
</select>
<%call GeraNomes("PORT",session("u_pub"),session("c_pub"),1,CON0)
no_curso_alterada = session("no_grau")

%>
<input name="curso_alterada" type="hidden" id="curso_alterada" value="<%response.Write(Server.URLEncode(no_curso_alterada))%>">
<%elseif opt="t" then
etapa_altera=Request.Form("e_pub")
'tenho que selecionar case e depois selecionar pois senão dá conflito quando carrega a primeira vez no arquivo altera.asp
Select case etapa_altera
case "0M1"
etapa_altera=901
case "0M2"
etapa_altera=902
case "JD1"
etapa_altera=903
case "JD2"
etapa_altera=904
case "JD3"
etapa_altera=905
end select


Select case etapa_altera
case 901
etapa_altera="0M1"
case 902
etapa_altera="0M2"
case 903
etapa_altera="JD1"
case 904
etapa_altera="JD2"
case 905
etapa_altera="JD3"
end select


if etapa_altera="" or isnull(etapa_altera) then
session("c_pub")=session("c_pub")
else
session("e_pub")=etapa_altera
session("t_pub")=session("t_pub")
end if


%><select name="turma" class="borda" onchange="recuperarChamada(this.value)">
 <option value="999990" selected></option> 
    <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Turma where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"' AND CO_Etapa='" & session("e_pub") & "' order by CO_Turma" 
		RS3.Open SQL3, CON0						
co_turma_check=9999990
while not RS3.EOF
co_turma= RS3("CO_Turma")
'cap_turma= RS3("NU_Capacidade")

		Set RS_al = Server.CreateObject("ADODB.Recordset")
		SQL_al = "SELECT COUNT(CO_Matricula) AS alunos_turma FROM TB_Matriculas where NU_Ano="&ano_letivo&" AND NU_Unidade="&session("u_pub")&" AND CO_Curso='"&session("c_pub")&"' AND CO_Etapa='" & session("e_pub") & "'  AND CO_Turma='" & co_turma & "'" 
		RS_al.Open SQL_al, CON1
								
alunos_turma=RS_al("alunos_turma")

cap_turma=cap_turma*1
alunos_turma=alunos_turma*1
vagas_turma=cap_turma-alunos_turma

'texto_turma= co_turma&" - "&vagas_turma
texto_turma= co_turma

if co_turma = co_turma_check then
RS3.MOVENEXT
elseif co_turma = session("t_ck") and ori="load" then
 %>
<option value="<%=response.Write(Server.URLEncode(co_turma))%>" selected> 
    <%response.Write(Server.URLEncode(texto_turma))%>
    </option> 
    <%
co_turma_check = co_turma
RS3.MOVENEXT
else
 %>
<option value="<%=response.Write(Server.URLEncode(co_turma))%>"> 
    <%response.Write(Server.URLEncode(texto_turma))%>
    </option> 
    <%
co_turma_check = co_turma
RS3.MOVENEXT
end if
WEND
%></select>
<%call GeraNomes("PORT",session("u_pub"),session("c_pub"),session("e_pub"),CON0)
no_etapa_alterada = session("no_serie")
%>
<input name="etapa_alterada" type="hidden" id="etapa_alterada" value="<%response.Write(Server.URLEncode(no_etapa_alterada))%>">
<%elseif opt="ch" then
turma_altera=request.form("t_pub")
if turma_altera="" or isnull(turma_altera) then
session("t_pub")=session("t_pub")
else
session("t_pub")=turma_altera
end if
session("ch_ck")=session("ch_ck")
%>
                    <select name="chamada" id="chamada" class="borda"  > 					 
<%nu_chamada_ckq=0
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Aluno_Esta_Turma where NU_Unidade="&session("u_pub")&" AND CO_Curso='"&session("c_pub")&"' AND CO_Etapa='" & session("e_pub") & "' AND CO_Turma='" & session("t_pub")& "' order by NU_Chamada"
		RS4.Open SQL4, CON2
		
while not RS4.EOF
nu_chamada=RS4("NU_Chamada")
session("ch_ck")=session("ch_ck")*1
nu_chamada=nu_chamada*1

 if (nu_chamada_ckq <>nu_chamada - 1) then
  teste_nu_chamada = nu_chamada-nu_chamada_ckq
  teste_nu_chamada=teste_nu_chamada-1
 for k=1 to teste_nu_chamada 
nu_chamada_falta=nu_chamada_ckq+1
	 %>	
	 
 	 <option value="<%=response.Write(nu_chamada_falta)%>"> 
    <%response.Write(nu_chamada_falta)%>
    </option>  
    <%
nu_chamada_ckq=nu_chamada_falta
next
 nu_chamada_ckq=nu_chamada	
RS4.MOVENEXT
else
 nu_chamada_ckq=nu_chamada
RS4.MOVENEXT
end if
wend
if RS4.EOF then
nu_chamada=nu_chamada*1
ultima_chamada=nu_chamada+1
%>
 	 <option value="<%=response.Write(ultima_chamada)%>" selected> 
    <%response.Write(ultima_chamada)%>
    </option>  	
<%end if	
 %>				  
       </select> 
<%else

session("t_pub")=Request.Form("t_pub")
%>					  
<%end if%>
</div>
