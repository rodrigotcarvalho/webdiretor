
<%
opt=request.querystring("opt")


		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
%>
<div align="center">
<%if opt="c" then
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

curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if

%>
                      <select name="etapa" class="borda" onchange="recuperarTurma(this.value)">
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

etapa_altera=Request.Form("e_pub")
if etapa_altera="" or isnull(etapa_altera) then
session("c_pub")=session("c_pub")
else
session("e_pub")=etapa_altera
session("t_pub")=session("t_pub")
end if
%>
  <select name="turma" class="borda">
                        <option value="999990" selected></option>
    <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Turma where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"' AND CO_Etapa='" & session("e_pub") & "' order by CO_Turma" 
		RS3.Open SQL3, CON0						
co_turma_check=9999990
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
