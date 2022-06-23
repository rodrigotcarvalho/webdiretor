<!--#include file="caminhos.asp"-->
<!--#include file="parametros.asp"-->
<!--#include file="../../global/tabelas_escolas.asp"-->
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
		Set CON_wr = Server.CreateObject("ADODB.Connection") 
		ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_wr.Open ABRIR_wr		

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

%>
<div align="center">
<%if opt="ctrl" then
session("u_pub")=session("u_pub")
session("c_pub")=Request.Form("c_pub")

			

		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")&" AND CO_Curso='"&session("c_pub")&"'"
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
		'tp_modelo = RS0c("TP_Modelo")		
	
	
		Set RS_WF = Server.CreateObject("ADODB.Recordset")
		SQL_WF = "SELECT * FROM TB_Autoriza_WF WHERE NU_Unidade="&session("u_pub")&" AND CO_Curso='"&session("c_pub")&"' and CO_Etapa='"&co_etapa&"'"				
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
					tp_modelo=tipo_divisao_ano(session("c_pub"),co_etapa,"tp_modelo",session("ano_letivo")) 				
					Set RSP = Server.CreateObject("ADODB.Recordset")
					'SQLP = "SELECT Distinct(NU_Periodo),NO_Periodo FROM TB_Periodo where TP_Modelo='"&tp_modelo&"' order by NU_Periodo"
					SQLP = "SELECT Distinct(NU_Periodo),NO_Periodo FROM TB_Periodo order by NU_Periodo"
					response.Write(SQLP&"<BR>")						
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
<%elseif Left(opt, 1)="c" then
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
<%elseif opt="e" then

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
<%
'Essa combo da etapa chama também a rotina da disciplina
elseif opt="e2" then

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
<%
'Essa combo da etapa chama também a rotina da disciplina
elseif opt="e5" then

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
                      <select name="etapa" class="select_style" onchange="recuperarDisciplina(curso.value,this.value)">
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
<%
'Essa combo da etapa chama também a rotina da disciplina
elseif opt="e6" then

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
                      <select name="etapa" class="select_style" onchange="MM_callJS('submitfuncao()')">
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
 elseif opt="t3" then

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
  <%elseif opt="t4" then

etapa_altera=Request.Form("e_pub")
if etapa_altera="" or isnull(etapa_altera) then
session("c_pub")=session("c_pub")
else
session("e_pub")=etapa_altera
session("t_pub")=session("t_pub")
end if
%>
 <select name="turma" class="select_style" onchange="recuperarPeriodo()">
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
  <%

' Só serve para gravar o valor de turma para as funções atualizar ocorrências (AOC), Emitir Ficha e Emitir Boletim do Web Secretaria
elseif opt="t5" then

etapa_altera=Request.Form("e_pub")
if etapa_altera="" or isnull(etapa_altera) then
session("c_pub")=session("c_pub")
else
session("e_pub")=etapa_altera
session("e_aoc")=etapa_altera
session("t_pub")=session("t_pub")
end if
%>
 <select name="turma" class="select_style" onchange="gravarTurma(this.value)">
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
 
<%
' Só serve para gravar o valor de turma para a função atualizar ocorrências (AOC)
elseif opt="t6" then

turma_aoc=Request.Form("t_pub")
session("t_aoc")=turma_aoc

elseif opt="d" then

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
elseif opt="d" then

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
curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if


		Set RSG = Server.CreateObject("ADODB.Recordset")
		SQLG = "SELECT DISTINCT CO_Materia_Principal FROM TB_Da_Aula where CO_Professor ="& prof_altera&" and CO_Etapa = '"&session("e_pub") &"' AND NU_Unidade = "&session("u_pub")&" and CO_Curso = '"&session("c_pub") &"' order by CO_Materia_Principal"
		RSG.Open SQLG, CONg
		
IF RSG.EOF THEN

RESPONSE.Write("Sem disciplinas cadastradas. Procure seu Coordenador.")

ELSE
%>
                      <select name="mat_prin" class="select_style" onChange="MM_callJS('recuperarPeriodo()')">
                        <option value="999999" selected></option>
                        <%
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
WEND
END IF
%>
                      </select>
<%
elseif opt="d3" then
etapa_altera=Request.Form("etapa_pub")
curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=etapa_altera
session("t_pub")=session("t_pub")
end if

%>
                       <select name="mat_prin" class="select_style" onChange="MM_callJS('recuperarPeriodo()')">
                        <option value="999999" selected></option> 
                        <%
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& session("e_pub") &"' AND CO_Curso ='"& session("c_pub") &"' order by NU_Ordem_Boletim "
		RS5.Open SQL5, CON0

response.Write(SQL5)

while not RS5.EOF
co_mat_prin= RS5("CO_Materia")


		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_mat_prin &"'"
		RS7.Open SQL7, CON0
		
if RS7.eof then
		Set RS7b = Server.CreateObject("ADODB.Recordset")
		SQL7b = "SELECT * FROM TB_Materia where CO_Materia ='"& co_mat_prin &"'"
		RS7b.Open SQL7b, CON0		
		
		no_mat_prin= RS7b("NO_Materia")
%>
                        <option value="<%=co_mat_prin%>"> 
                        <%response.Write(Server.URLEncode(no_mat_prin))%>						
                        </option>
                        <%
RS5.MOVENEXT						
else
RS5.MOVENEXT
end if
WEND%>

                      </select>					  					  
<%
elseif opt="d4" then

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
elseif opt="p2" then
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")

%>
<select name="periodo" class="select_style" id="periodo" onChange="MM_callJS('recuperarAvaliacoes()')">
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
elseif opt="p3" then
'etapa_altera=Request.Form("e_pub")
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")

%>
<select name="periodo" class="select_style" id="periodo" onchange="recuperarMedia()">
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
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")

%>																		
<select name="avaliacoes" class="select_style" id="avaliacoes" onChange="MM_callJS('submitfuncao()')">
                                      <option value="999990"></option>
<%

		Set RSTB = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& session("u_pub") &" AND CO_Curso = '"& session("c_pub") &"' AND CO_Etapa = '"& session("e_pub")  &"'"
		Set RSTB = CONg.Execute(CONEXAO)

'response.Write(CONEXAO)

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
		response.Write("ERRO")
		End if
	end if
end if	

dados_tabela=verifica_dados_tabela(CAMINHOn,opcao,outro)
	dados_separados=split(dados_tabela,"#$#")
	ln_nom_cols=dados_separados(4)
	nm_vars=dados_separados(5)
	nm_bd=dados_separados(6)
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
elseif opt="av2" then

session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")

%>																		
<select name="avaliacoes" class="select_style" id="avaliacoes" onChange="MM_callJS('submitfuncao()')">
                                      <option value="999990"></option> 
<%

		Set RSTB = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& session("u_pub") &" AND CO_Curso = '"& session("c_pub") &"' AND CO_Etapa = '"& session("e_pub")  &"'"
		Set RSTB = CONg.Execute(CONEXAO)

'response.Write(CONEXAO)

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
		response.Write("ERRO")
		End if
	end if
end if	

dados_tabela=verifica_dados_tabela(CAMINHOn,opcao,outro)
	dados_separados=split(dados_tabela,"#$#")
	ln_nom_cols=dados_separados(4)
	nm_vars=dados_separados(5)
	nm_bd=dados_separados(6)
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
</div>
