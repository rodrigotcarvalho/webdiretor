<!--#include file="../../../../inc/caminhos.asp"-->
<%
opt=request.querystring("opt")
nv=request.querystring("nv")

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
%>
<div align="center">
<%if opt="c" then

if nv="1" then
unidade_altera1=Request.Form("u_pub1")
if unidade_altera1="" or isnull(unidade_altera1) then
session("u_pub1")=session("u_pub1")
else
session("u_pub1")=unidade_altera1
session("c_pub1")=session("c_pub1")
session("e_pub1")=session("e_pub1")
session("t_pub1")=session("t_pub1")
end if
%>
                      <select name="curso1" class="borda" onchange="recuperarEtapa1(this.value)">
<%elseif nv="2" then
unidade_altera2=Request.Form("u_pub2")
if unidade_altera2="" or isnull(unidade_altera2) then
session("u_pub2")=session("u_pub2")
else
session("u_pub2")=unidade_altera2
session("c_pub2")=session("c_pub2")
session("e_pub2")=session("e_pub2")
session("t_pub2")=session("t_pub2")
end if
%>
                      <select name="curso2" class="borda" onchange="recuperarEtapa2(this.value)">
<%elseif nv="3" then
unidade_altera3=Request.Form("u_pub3")
if unidade_altera3="" or isnull(unidade_altera3) then
session("u_pub3")=session("u_pub3")
else
session("u_pub3")=unidade_altera3
session("c_pub3")=session("c_pub3")
session("e_pub3")=session("e_pub3")
session("t_pub3")=session("t_pub3")
end if
%>
                      <select name="curso3" class="borda" onchange="recuperarEtapa3(this.value)">
<%elseif nv="4" then
unidade_altera4=Request.Form("u_pub4")
if unidade_altera4="" or isnull(unidade_altera4) then
session("u_pub4")=session("u_pub4")
else
session("u_pub4")=unidade_altera4
session("c_pub4")=session("c_pub4")
session("e_pub4")=session("e_pub4")
session("t_pub4")=session("t_pub4")
end if
%>
                      <select name="curso4" class="borda" onchange="recuperarEtapa4(this.value)">
<%end if%>					  
                        <option value="nulo" selected></option>
                        <%		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
if nv="1" then
		SQL0 = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub1")
elseif nv="2" then
		SQL0 = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub2")
elseif nv="3" then
		SQL0 = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub3")
elseif nv="4" then
		SQL0 = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub4")
end if		
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
		
NO_Curso = RS0a("NO_Abreviado_Curso")		
%>
                        <option value="<%response.Write(CO_Curso)%>"> 
                        <%response.Write(Server.URLEncode(NO_Curso))%>
                        </option>
                        <%

CO_Curso_check = CO_Curso
RS0.MOVENEXT
end if
WEND
%>
                      </select>
<%elseif opt="e" then

if nv="1" then
curso_altera1=Request.Form("c_pub1")
if curso_altera1="" or isnull(curso_altera1) then
session("c_pub1")=session("c_pub1")
else
session("c_pub1")=curso_altera1
session("e_pub1")=session("e_pub1")
session("t_pub1")=session("t_pub1")
end if

%>
                      <select name="etapa1" class="borda" onchange="recuperarTurma1(this.value)">
<%elseif nv="2" then

curso_altera2=Request.Form("c_pub2")
if curso_altera2="" or isnull(curso_altera2) then
session("c_pub2")=session("c_pub2")
else
session("c_pub2")=curso_altera2
session("e_pub2")=session("e_pub2")
session("t_pub2")=session("t_pub2")
end if
%>
                      <select name="etapa2" class="borda" onchange="recuperarTurma2(this.value)">
<%elseif nv="3" then
curso_altera3=Request.Form("c_pub3")
if curso_altera3="" or isnull(curso_altera3) then
session("c_pub3")=session("c_pub3")
else
session("c_pub3")=curso_altera3
session("e_pub3")=session("e_pub3")
session("t_pub3")=session("t_pub3")
end if

%>
                      <select name="etapa3" class="borda" onchange="recuperarTurma3(this.value)">
<%elseif nv="4" then
curso_altera4=Request.Form("c_pub4")
if curso_altera4="" or isnull(curso_altera4) then
session("c_pub4")=session("c_pub4")
else
session("c_pub4")=curso_altera4
session("e_pub4")=session("e_pub4")
session("t_pub4")=session("t_pub4")
end if
%>
                      <select name="etapa4" class="borda" onchange="recuperarTurma4(this.value)">
<%end if%>

                        <option value="nulo" selected></option>
                        <%		

		Set RS0b = Server.CreateObject("ADODB.Recordset")
if nv="1" then
		SQL0b = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub1")&"AND CO_Curso='"&session("c_pub1")&"'"
elseif nv="2" then
		SQL0b = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub2")&"AND CO_Curso='"&session("c_pub2")&"'"
elseif nv="3" then
		SQL0b = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub3")&"AND CO_Curso='"&session("c_pub3")&"'"
elseif nv="4" then
		SQL0b = "SELECT * FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub4")&"AND CO_Curso='"&session("c_pub4")&"'"
end if			
		RS0b.Open SQL0b, CON0
		
	
CO_Etapa_check="999999"		
While not RS0b.EOF
CO_Etapa = RS0b("CO_Etapa")

if CO_Etapa = CO_Etapa_check then
RS0b.MOVENEXT		
else

		Set RS0c = Server.CreateObject("ADODB.Recordset")
if nv="1" then
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("c_pub1")&"' AND CO_Etapa='"&CO_Etapa&"'"
elseif nv="2" then
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("c_pub2")&"' AND CO_Etapa='"&CO_Etapa&"'"
elseif nv="3" then
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("c_pub3")&"' AND CO_Etapa='"&CO_Etapa&"'"
elseif nv="4" then
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("c_pub4")&"' AND CO_Etapa='"&CO_Etapa&"'"
end if			
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
%>
                        <option value="<%response.Write(CO_Etapa)%>"> 
                        <%response.Write(Server.URLEncode(NO_Etapa))%>
                        </option>
                        <%

CO_Etapa_check = CO_Etapa
RS0b.MOVENEXT
end if
WEND
%>
                      </select>

<%elseif opt="t" then

if nv="1" then
etapa_altera1=Request.Form("e_pub1")
if etapa_altera1="" or isnull(etapa_altera1) then
session("c_pub1")=session("c_pub1")
else
session("e_pub1")=etapa_altera1
session("t_pub1")=session("t_pub1")
end if
%>
						<select name="turma1" class="borda">
<%elseif nv="2" then
etapa_altera2=Request.Form("e_pub2")
if etapa_altera2="" or isnull(etapa_altera2) then
session("c_pub2")=session("c_pub2")
else
session("e_pub2")=etapa_altera2
session("t_pub2")=session("t_pub2")
end if
%>
						<select name="turma2" class="borda">
<%elseif nv="3" then
etapa_altera3=Request.Form("e_pub3")
if etapa_altera3="" or isnull(etapa_altera3) then
session("c_pub3")=session("c_pub3")
else
session("e_pub3")=etapa_altera3
session("t_pub3")=session("t_pub3")
end if
%>
						<select name="turma3" class="borda">
<%elseif nv="4" then
etapa_altera4=Request.Form("e_pub4")
if etapa_altera4="" or isnull(etapa_altera4) then
session("c_pub4")=session("c_pub4")
else
session("e_pub4")=etapa_altera4
session("t_pub4")=session("t_pub4")
end if
%>
						<select name="turma4" class="borda">
<%end if%>

                        <option value="nulo" selected></option>
    <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
if nv="1" then
		SQL3 = "SELECT * FROM TB_Turma where NU_Unidade="&session("u_pub1")&"AND CO_Curso='"&session("c_pub1")&"' AND CO_Etapa='" & session("e_pub1") & "' order by CO_Turma" 
elseif nv="2" then
		SQL3 = "SELECT * FROM TB_Turma where NU_Unidade="&session("u_pub2")&"AND CO_Curso='"&session("c_pub2")&"' AND CO_Etapa='" & session("e_pub2") & "' order by CO_Turma" 
elseif nv="3" then
		SQL3 = "SELECT * FROM TB_Turma where NU_Unidade="&session("u_pub3")&"AND CO_Curso='"&session("c_pub3")&"' AND CO_Etapa='" & session("e_pub3") & "' order by CO_Turma" 
elseif nv="4" then
		SQL3 = "SELECT * FROM TB_Turma where NU_Unidade="&session("u_pub4")&"AND CO_Curso='"&session("c_pub4")&"' AND CO_Etapa='" & session("e_pub4") & "' order by CO_Turma" 
end if		
		RS3.Open SQL3, CON0						
co_turma_check=999990
while not RS3.EOF
co_turma= RS3("CO_Turma")

if co_turma = co_turma_check then
RS3.MOVENEXT
else


 %>
    <option value="<%=co_turma%>"> 
    <%response.Write(Server.URLEncode(co_turma))%>
    </option>
    <%
co_turma_check = co_turma
RS3.MOVENEXT
end if
WEND
%>
    
  </select>
<%else

session("t_pub")=Request.Form("t_pub")
%>
<%end if%>
</div>
