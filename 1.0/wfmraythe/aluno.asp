<!--#include file="inc/caminhos.asp"-->
<!--#include file="inc/funcoes.asp"-->
<!--#include file="inc/funcoes2.asp"-->
<%


 	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
ano_pesquisa = Request.Form("e_pub")	
Session("aluno_selecionado") = ""

 	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO_wf&";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1
	
	Set CON3 = Server.CreateObject("ADODB.Connection") 
	ABRIR3 = "DBQ="& CAMINHO_o & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON3.Open ABRIR3	

	Set CON7 = Server.CreateObject("ADODB.Connection") 
	ABRIR7 = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON7.Open ABRIR7

tp=session("tp")
nome = session("nome") 
co_user = session("co_user")

alunos_vetor=0
	if tp="R" then			
		tipo="Responsável"
		SQL = "select * from TB_RespxAluno where CO_Usuario = " & co_user &" ORDER BY CO_Aluno"
		set RS = CON.Execute (SQL)
		'quantos=RS("quantos")
		quantos=0

		While not RS.EOF
			alunos = RS("CO_Aluno")
			'if quantos=0 then
			'	Session("aluno_selecionado")=alunos
			'end if
			alunos_vetor=alunos_vetor&"?"&alunos
			quantos=quantos+1
		RS.MOVENEXT
		WEND				
		if opt="as" then
		alunos=request.form("co_aluno")
		Session("aluno_selecionado")=alunos
		end if
		if opt="ad" or opt="sa" then
		alunos=Session("aluno_selecionado")
		Session("aluno_selecionado")=alunos
		end if		
	elseif tp="A" then
		tipo="Aluno"
		alunos_vetor="0?"&co_user		
		Session("aluno_selecionado")=co_user	
	end if
%>
<table width="800" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="10">&nbsp;</td>
                  <td height="30" colspan="7" valign="top"><font class="style3">O 
                    Web Fam&iacute;lia disponibilizar&aacute; acesso &agrave;s 
                    informa&ccedil;&otilde;es do aluno que estiver selecionado 
                    abaixo no seguinte ano letivo:
										<% 
					'ano_vigente = session("ano_vigente")	
					min_ano_letivo = session("min_ano_letivo")
					max_ano_letivo = session("max_ano_letivo")
					
					ano_letivo = ano_letivo*1
					'ano_vigente = ano_vigente*1					
					min_ano_letivo = min_ano_letivo*1					
					max_ano_letivo = max_ano_letivo*1	
					

'					if min_ano_letivo < ano_vigente then
'						if ano_letivo=ano_vigente then
'							'forçar a exibição de apenas 2 anos
'							min_ano_letivo=ano_letivo-1
'						else
'							min_ano_letivo=ano_letivo
'						end if	
'					end if
					
					if min_ano_letivo = max_ano_letivo then
						response.Write(min_ano_letivo)
					else	
					
					
						%>                    
						
						<select name="select_ano_letivo" class="textbox" onChange="atualizarAnoLetivo(this.value);">
						<%
						while min_ano_letivo <= max_ano_letivo 
							if min_ano_letivo=ano_letivo then
								selected = "Selected"
							else
								selected = ""
							end if
						%>
						<option value="<%response.write(min_ano_letivo)%>" <%response.write(selected)%>><%response.write(min_ano_letivo)%></option>
						<%
						min_ano_letivo=min_ano_letivo+1
						wend%>
						
						</select>
                    <%end if%></font></td>
                </tr>
                <tr> 
                  <td width="10">&nbsp;</td>
                  <td width="20">&nbsp;</td>
                  <td width="80"> <div align="center"><font class="style3">MATR&Iacute;CULA</font></div></td>
                  <td width="245"><font class="style3">NOME</font></td>
                  <td width="80"> <div align="center"><font class="style3">UNIDADE</font></div></td>
                  <td width="195"><font class="style3">CURSO</font></td>
                  <td width="60"> <div align="center"><font class="style3">TURMA</font></div></td>
                  <td width="109"><div align="center"><font class="style3">NASCIMENTO</font></div></td>
                </tr>
                <%
 	

vetor = split(alunos_vetor,"?")
for i =1 to ubound(vetor)
co_aluno= vetor(i)

	SQL2 = "select * from TB_Alunos where CO_Matricula = " & co_aluno 
	set RS2 = CON1.Execute (SQL2)
	
nome_aluno= RS2("NO_Aluno")

		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Contatos where CO_Matricula = " & co_aluno &" AND TP_Contato='ALUNO'"
		RS7.Open SQL7, CON7

nascimento = RS7("DA_Nascimento_Contato")

dados_dtd= split(nascimento, "/" )
dia_de= dados_dtd(0)
mes_de= dados_dtd(1)
ano_de= dados_dtd(2)
dia_de=dia_de*1
if dia_de<10 then
dia_de="0"&dia_de
end if
mes_de=mes_de*1
if mes_de<10 then
mes_de="0"&mes_de
end if

nascimento=dia_de&"/"&mes_de&"/"&ano_de

	SQL3 = "select * from TB_Matriculas where NU_Ano="& ano_pesquisa &" AND CO_Matricula = " & co_aluno 
	set RS3 = CON1.Execute (SQL3)

	if not RS3.EOF then
		nu_unidade= RS3("NU_Unidade")
		co_curso= RS3("CO_Curso")
		co_etapa= RS3("CO_Etapa")
		co_turma= RS3("CO_Turma")
		
		call GeraNomes("PORT",nu_unidade,co_curso,co_etapa,CON0)
		no_unidade = session("no_unidade")
		no_curso = session("no_curso")
		no_etapa = session("no_etapa")
		prep_curso=session("prep_curso")
		local= no_etapa&" "&prep_curso&" "&no_curso
		%>
                <tr> 
                  <td width="10">&nbsp;</td>
                  <td> 
                    <%
					if isnull(Session("aluno_selecionado")) or Session("aluno_selecionado")="" then
						Session("aluno_selecionado") = co_aluno
					end if						
										
					alunos=Session("aluno_selecionado")
					co_aluno=co_aluno*1
					alunos=alunos*1
					if co_aluno=alunos then
						unidade_documentos=nu_unidade
						curso_documentos=co_curso
						etapa_documentos=co_etapa
						turma_documentos=co_turma					
					%>
					<input name="co_aluno" type="radio" onClick="MM_callJS('submit()')" value="<%=co_aluno%>" checked> 
					<%else%>
					<input type="radio" name="co_aluno" onClick="MM_callJS('submit()')" value="<%=co_aluno%>">	
					<%end if%>
                  </td>
                  <td><div align="center"><font class="style1"> 
                      <%response.write(co_aluno)%>
                      </font> </div></td>
                  <td> <font class="style1"> 
                    <%response.write(Server.URLEncode(nome_aluno))%>
                    </font> </td>
                  <td><div align="center"><font class="style1"> 
                      <%response.write(Server.URLEncode(no_unidade))%>
                      </font></div></td>
                  <td><font class="style1"> 
                    <%response.write(Server.URLEncode(local))%>
                    </font></td>
                  <td><div align="center"><font class="style1"> 
                      <%response.write(Server.URLEncode(co_turma))%>
                      </font></div></td>
                  <td><div align="center"><font class="style1"> 
                      <%
					  nascimento=FormatDateTime(nascimento,2)					  			  
					  response.write(nascimento)%>
                      </font></div></td>
                </tr>

                <%
	end if
NEXT
%>
                <tr>
                  <td></td>
                  <td height="40" colspan="7" valign="bottom">
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <th colspan="2" scope="row"><hr>                      </th>
                      </tr>
<%
	Set RS4 = Server.CreateObject("ADODB.Recordset")
	SQL4 = "select * from TB_Ocorrencia_Aluno where CO_Matricula = " & alunos &" Order BY DA_Ocorrencia DESC,HO_Ocorrencia"
	set RS4 = CON3.Execute (SQL4)
if RS4.EOF then
%>
                    <tr>
                     <th width="190" scope="row"><div align="right"><font class="style3"> Última Ocorrencia Registrada:</font></div></th>
                      <td width="610" scope="value"><div align="left"><font class="style1">&nbsp;Sem Ocorrências</font></div></th>
                    </tr>
<%else
data_ocor=RS4("DA_Ocorrencia")
data_ocor=FormatDateTime(data_ocor,2)
%>
                    <tr>
                      <th width="190" scope="row"><div align="right"><font class="style3"> Última Ocorrencia Registrada:</font></div></th>
                      <td width="610" scope="value"><div align="left"><font class="style1"> &nbsp;<%response.Write(data_ocor)%></font></div></td>
                    </tr>
<%end if
		Set RS_doc = Server.CreateObject("ADODB.Recordset")
		SQL_doc = "SELECT * FROM TB_Documentos where TP_Doc= 1 AND (((Unidade='"&unidade_documentos&"') AND (Curso='"&curso_documentos&"') AND  (Etapa='"&etapa_documentos&"') AND (Turma='"&turma_documentos&"')) OR ((Unidade='"&unidade_documentos&"') AND (Curso Is Null) AND (Etapa Is Null) AND (Turma Is Null)) OR ((Unidade='"&unidade_documentos&"') AND (Curso='"&curso_documentos&"') AND (Etapa Is Null) AND (Turma Is Null)) OR ((Unidade='"&unidade_documentos&"') AND (Curso='"&curso_documentos&"') AND  (Etapa='"&etapa_documentos&"') AND (Turma Is Null)) OR  ((Unidade Is Null) AND (Curso Is Null) AND  (Etapa Is Null) AND (Turma Is Null))) order by DA_Doc Desc"
		RS_doc.Open SQL_doc, CON
if RS_doc.eof then
%>
                      <tr class="<%response.write(cor)%>"> 
                      <th width="190" scope="row"><div align="right"><font class="style3">Último Informe Escolar:</font></div> </th>
                        <td width="610" scope="value"> <div align="Left"><font class="style1"> 
                          &nbsp;Sem Publicações.
                          </font></div></td>
                      </tr>

<%else
tipo_arquivo=RS_doc("TP_Doc")
tit1=RS_doc("TI1_Doc")
data_pub=RS_doc("DA_Doc")

select case tipo_arquivo
case 1
nome_tipo_arquivo="Circulares"
case 2
nome_tipo_arquivo="Avaliações e Gabaritos"
case 3
nome_tipo_arquivo="Reunião de Pais"
end select

if data_pub="" or isnull(data_pub) then
else			
dados_dtd= split(data_pub, "/" )
dia_de= dados_dtd(0)
mes_de= dados_dtd(1)
ano_de= dados_dtd(2)
end if


if dia_de<10 then
dia_de="0"&dia_de
end if
if mes_de<10 then
mes_de="0"&mes_de
end if

data_inicio=dia_de&"/"&mes_de&"/"&ano_de
data_inicio=FormatDateTime(data_inicio,2)
%>                    
                    <tr>
                      <th width="190" scope="row"><div align="right"><font class="style3">Último Informe Escolar:</font></div> </th>
                      <td width="610" scope="value"><font class="style1"> &nbsp;<%response.Write(nome_tipo_arquivo &" - "&tit1 &", publicado em "&data_inicio)%></font></td>
                    </tr>
<%end if%>                    
                  </table></td>
                </tr>
                <tr>
				  <td width="10"></td>
                  <td height="40" colspan="7" valign="bottom"> 
                    <%if quantos>1 then%>
                    <font class="style3"> ATEN&Ccedil;&Atilde;O ! </font><font class="style3">Caso 
                    queira obter informa&ccedil;&otilde;es de outro aluno ou outro ano letivo volte 
                    a P&Aacute;GINA INICIAL e fa&ccedil;a nova sele&ccedil;&atilde;o.</font> 
                    <%end if%>
                  </td>
                </tr>
              </table>
              