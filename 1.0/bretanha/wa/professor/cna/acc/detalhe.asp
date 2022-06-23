<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../../global/dados_usr.asp"-->
<!--#include file="../../../../../global/funcoes_diversas.asp"-->
<!--#include file="../../../../inc/funcoes7.asp"-->
<%
cod_cons=Request.Form("matric_pub")
ano_letivo=Request.Form("ano_pub")
unidade=Request.Form("u_pub")
curso=Request.Form("c_pub")
co_etapa=Request.Form("e_pub")
turma=Request.Form("t_pub")
vet_co_materia_detalhe=Request.Form("materia_pub")
caminho_nota=Request.Form("caminho_pub")
tb_nota=Request.Form("tb_nt")
acumulado=Request.Form("acum_pub")
qto_falta=Request.Form("qf_pub")
parametros_chamada_jscript=Request.Form("prmtr_pub")
dados_notas_detalhe=session("dados_notas_detalhe")
width_tabela=session("width_pub")
javascript_periodo=Request.Form("nom_per_pub")
num_periodo_detalhe=Request.Form("num_per_pub")
caminho_cons=CAMINHO_al
caminho_contato=CAMINHO_ct
caminho_ajax=replace(caminho_nota,"\","$b$")
caminho_ajax=replace(caminho_ajax,"_","$u$")


if tb_nota ="TB_NOTA_A" then
	opcao="A"
elseif tb_nota="TB_NOTA_B" then
	opcao="B"		
elseif tb_nota ="TB_NOTA_C" then
	opcao="C"
elseif tb_nota ="TB_NOTA_D" then
	opcao="D"			
elseif tb_nota ="TB_NOTA_E" then
	opcao="E"					
else
	response.Write("ERRO")
end if	

'dados_notas_detalhe=javascript_periodo&"$!$"&num_periodo_detalhe&"$!$"&periodo_m1&"$!$"&periodo_m2&"$!$"&periodo_m3&"$!$"&ntazl&"$!$"&ntvml&"$!$"&999&"$!$"&peso_m2_m1&"$!$"&peso_m2_m2&"$!$"&peso_m3_m1&"$!$"&peso_m3_m2&"$!$"&peso_m3_m3

dados_notas=split(dados_notas_detalhe,"$!$")
periodo_m1=dados_notas(0)
periodo_m2=dados_notas(1)
periodo_m3=dados_notas(2)
ntazl=dados_notas(3)
ntvml=dados_notas(4)
peso_m2_m1=dados_notas(6)
peso_m2_m2=dados_notas(7)
peso_m3_m1=dados_notas(8)
peso_m3_m2=dados_notas(9)
peso_m3_m3=dados_notas(10)

nom_periodo=split(javascript_periodo,"#!#")
num_periodo=split(num_periodo_detalhe,"#!#")	

'Matricula,Nome, Idade,Número de chamada, Status
dados_aluno=busca_dados(ano_letivo,cod_cons,caminho_cons,caminho_contato,"a","ALUNO")
dados=split(dados_aluno,"#!#")
nome=dados(0)
situac=dados(26)
num_cham=dados(32)
nascimento=dados(33)
vetor_nascimento = Split(nascimento,"/")  
dia_a = vetor_nascimento(0)
mes_a = vetor_nascimento(1)
ano_a = vetor_nascimento(2)
idade=calcula_idade(ano_a,mes_a,dia_a) 

	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR

	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
		
	Set CON3 = Server.CreateObject("ADODB.Connection") 
	ABRIR3 = "DBQ="& CAMINHO_o & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON3.Open ABRIR3	
	
	Set CON4 = Server.CreateObject("ADODB.Connection") 
	ABRIR4 = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON4.Open ABRIR4		
	
	Set RS = Server.CreateObject("ADODB.Recordset")
	CONEXAO = "Select * from TB_Situacao_Aluno WHERE CO_Situacao = '"& situac &"'"
	Set RS = CON0.Execute(CONEXAO)
	
	Set CONt = Server.CreateObject("ADODB.Connection") 
	ABRIRt = "DBQ="& CAMINHO_t & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONt.Open ABRIRt	
	
	
	if RS.EOF then
		nome_situac="detalhe.asp#98 - Erro na busca por TX_Descricao_Situacao."	
	else
		nome_situac=RS("TX_Descricao_Situacao")
	end if	


	vet_co_materia= split(vet_co_materia_detalhe,"#!#")	

	qtd_colunas=ubound(vet_co_materia)+1
	colspan_notas=qtd_colunas+1
	width_scroll=20
	width_periodo=30
	width_tb_dados_turma=width_nu_chamada+width_nome+width_periodo+width_lupa-50
	width_else=width_tabela-225
	width_else_notas=(width_tabela-width_periodo)/qtd_colunas
	
	cor_nota_vml="#FF0000"	
	cor_nota_azl="#0000FF"	
	cor_nota_prt="#000000"	
	cor_nota_vrd="#006600"		

mostra_img=Session("mostra_foto")

nom_unidade=verifica_nome(unidade,0,0,0,0,CON0,"u", "f")
nom_curso=verifica_nome(unidade,curso,0,0,0,CON0,"c", "f")
nom_etapa=verifica_nome(unidade,curso,co_etapa,0,0,CON0,"e", "f")
nom_turma=turma
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="<%response.Write(width_tabela)%>" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="225">&nbsp;</td>
        <td width="<%response.Write(width_else)%>"><div align="right"> <span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<a href="#" onClick="MM_showHideLayers('fundo','','hide','lupa','','hide');focar('<%response.Write(num_cham&"c2")%>');mudar_cor_focus(<%response.Write(parametros_chamada_jscript)%>);limpalupa();"><img src="../../../../img/fecha.gif" width="20" height="16" border="0" align="absbottom"></a></font></span></div></td>
      </tr>
      <tr>
        <td width="225" height="300" rowspan="2" valign="top"><%
if mostra_img="OK" then
%><img src="../../../../img/fotos/aluno/<% response.Write(cod_cons)%>.jpg" width="225" height="300"><% end if%></td>
        <td width="<%response.Write(width_else)%>" height="187" align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="66" height="20" class="zoom_tit"><div align="right">N&ordm;:</div></td>
            <td width="74" height="20" class="zoom_texto">&nbsp;
<% response.Write(num_cham)%></td>
            <td width="70" height="20" class="zoom_tit">&nbsp;</td>
            <td height="20" colspan="3" class="zoom_texto">&nbsp;</td>
          </tr>
          <tr>
            <td height="20" class="zoom_tit"><div align="right">Matr&iacute;cula:</div></td>
            <td height="20" colspan="5" class="zoom_texto">&nbsp;
              <% response.Write(cod_cons)%></td>
          </tr>
          <tr>
            <td height="20" class="zoom_tit"><div align="right">Nome:</div></td>
            <td height="20" colspan="5" class="zoom_texto">&nbsp;
              <% response.Write(Server.URLEncode(nome))%></td>
            </tr>
          <tr>
            <td width="66" height="20" class="zoom_tit"><div align="right">Status:</div></td>
            <td height="20" colspan="5" class="zoom_texto">&nbsp;
              <% response.Write(Server.URLEncode(nome_situac))%></td>
            </tr>
          <tr>
            <td height="20" class="zoom_tit"><div align="right">Idade:</div></td>
            <td height="20" colspan="3" class="zoom_texto">&nbsp;
              <% response.Write(idade)%>
Anos</td>
            <td width="80" height="20" class="zoom_texto">&nbsp;</td>
            <td width="56" height="20" class="zoom_texto">&nbsp;</td>
          </tr>
          <tr>
            <td height="20" class="zoom_tit"><div align="right">Unidade:</div></td>
            <td height="20" colspan="5" class="zoom_texto">&nbsp;
              <% response.Write(Server.URLEncode(nom_unidade))%></td>
            </tr>
          <tr>
            <td height="20" class="zoom_tit"><div align="right">Curso:</div></td>
            <td height="20" colspan="5" class="zoom_texto">&nbsp;
              <% response.Write(Server.URLEncode(nom_curso))%></td>
            </tr>
          <tr>
            <td height="20" class="zoom_tit"><div align="right">Etapa:</div></td>
            <td height="20" colspan="5" class="zoom_texto">&nbsp;
              <% response.Write(Server.URLEncode(nom_etapa))%></td>
            </tr>
          <tr>
            <td height="20" class="zoom_tit"><div align="right">Turma:</div></td>
            <td height="20" colspan="5" class="zoom_texto">&nbsp;
              <% response.Write(Server.URLEncode(nom_turma))%></td>
            </tr>
        </table>
          </td>
      </tr>
      <tr>
        <td height="110" valign="bottom"><div align="center"><input name="button" type="button" class="botao_cancelar" id="button" value="Retornar" onClick="MM_showHideLayers('fundo','','hide','lupa','','hide');focar('<%response.Write(num_cham&"c2")%>');mudar_cor_focus(<%response.Write(parametros_chamada_jscript)%>);limpalupa();"/><br /><font size="1">&nbsp;</font></div>
	         
        </td>
      </tr>
    </table></td>
  </tr>
  <tr><td>
  <table width="<%response.Write(width_tabela)%>" border="0" align="left" cellpadding="0" cellspacing="0">
    <thead>
      <tr><td height="30" colspan="<%response.Write(colspan_notas)%>" valign="bottom" class="zoom_label"><div align="left">Notas</div></td></tr>
      <tr>   
        <td width="<%response.Write(width_periodo)%>"  class="zoom_tit"><div align="center">Per</div></td>
        <%for m=0 to ubound(vet_co_materia)%>
        <td width="<%response.Write(width_else_notas)%>" class="zoom_tit"><div align="center"><%response.Write(vet_co_materia(m))%></div></td>
  <%	next%>
  </tr>   
  <%

	cor = "tb_fundo_linha_par" 
	
	Set RS1 = Server.CreateObject("ADODB.Recordset")
	SQL1 = "Select * from TB_Mapao_Notas WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"' and CO_Matricula="&cod_cons&" ORDER BY NU_Seq_Per"
	Set RS1 = CONt.Execute(SQL1)	
	
	while not RS1.EOF			
	
		conta_notas=1
		vetor_nota_exibe=""
		seq_per=RS1("NU_Seq_Per")
		no_exibe_per=RS1("CO_Per")
		periodo_real=RS1("NU_Seq_Per_Real")
		for conta_notas=1 to ubound(vet_co_materia)+1
			if conta_notas<10 then
				campo="CO_0"&conta_notas
			else
				campo="CO_"&conta_notas			
			end if
			
			val_nota=RS1(campo)
			if conta_notas=1 then
				vetor_nota_exibe=val_nota
			else	
				vetor_nota_exibe=vetor_nota_exibe&"#!#"&val_nota
			end if
			vetor_nota_separa=vetor_nota_exibe	
		next
%> 
      <tr class="<%response.Write(cor)%>">         
        
        <td width="<%response.Write(width_periodo)%>" class="zoom_texto"><div align="center"><%response.Write(no_exibe_per)%></div></td>            
        
        <%For dsc=0 to ubound(vet_co_materia)	%>	
        <td width="<%response.Write(width_else)%>" class="zoom_texto" onFocus="recuperarNotaZoom(<% response.Write(width_tabela)%>,<% response.Write(cod_cons)%>,<% response.Write(ano_letivo)%>,<% response.Write(curso)%>,<% response.Write(co_etapa)%>,'<%response.Write(vet_co_materia(dsc))%>','<% response.Write(caminho_ajax)%>','<% response.Write(opcao)%>',<% response.Write(periodo_real)%>)"><div align="center">
          <%
		  
			vetor_nota=split(vetor_nota_exibe,"#!#")		  
		 
			media=vetor_nota(dsc)
			teste = isnumeric(media)			
			if teste=false then
				response.Write("&nbsp;")
			else	
				media=media*1	
				ntazl=ntazl*1
				ntvml=ntvml*1
				if media>=ntazl then	
					response.Write("<font color="&cor_nota_prt&">"&formatnumber(media,1)&"</font>")				
				elseif media>=ntvml then	
					response.Write("<font color="&cor_nota_azl&">"&formatnumber(media,1)&"</font>")
				else	
					response.Write("<font color="&cor_nota_vml&">"&formatnumber(media,1)&"</font>")	
				end if	
			end if	
'			teste = isnumeric(media)			
'			if teste=false then
'				response.Write("&nbsp;")
'			else	
'				media=media*1	
'				ntazl=ntazl*1
'				ntvml=ntvml*1
'
'				if (nom_periodo(n)="QF1" or nom_periodo(n)="QF2" or nom_periodo(n)="QF3") and media=0 then
'					response.Write("<font color="&cor_nota_prt&"></font>")
'				elseif (nom_periodo(n)="QF1" or nom_periodo(n)="QF2" or nom_periodo(n)="QF3") then
'					response.Write("<font color="&cor_nota_vrd&">"&formatnumber(media,1)&"</font>")
'				else
'					if media>=ntazl then	
'						response.Write("<font color="&cor_nota_prt&">"&formatnumber(media,1)&"</font>")				
'					elseif media>=ntvml then	
'						response.Write("<font color="&cor_nota_azl&">"&formatnumber(media,1)&"</font>")
'					else	
'						response.Write("<font color="&cor_nota_vml&">"&formatnumber(media,1)&"</font>")	
'					end if	
'				end if	
'			end if	
			
			%>
          </div></td>
        <%NEXT%>      
        </tr> 
      <%RS1.MOVENEXT
		WEND%>              
      </thead>
  </table>  </td></tr>
  <tr>
    <td height="30" class="zoom_label" valign="bottom">
Avalia&ccedil;&otilde;es
    </td>
  </tr>  
  <tr>
    <td height="20">
    <div id="div_avaliacoes_zoom"></div>
    </td>
  </tr>
  <tr>
    <td>
                <table width="<%response.Write(width_tabela)%>" border="0" align="left" cellpadding="0" cellspacing="0">
                  <tr class="<%response.write(cor)%>">
                    <td height="30" colspan="6" valign="bottom" class="zoom_label">Ocorr&ecirc;ncias</td>
                  </tr>                
                  <%
	Set RS3 = Server.CreateObject("ADODB.Recordset")
	SQL3 = "select * from TB_Ocorrencia_Aluno where CO_Matricula = " & cod_cons &" Order BY DA_Ocorrencia DESC,HO_Ocorrencia"
	set RS3 = CON3.Execute (SQL3)
if RS3.EOF then
%>

                  <tr class="<%response.write(cor)%>"> 
                    <td colspan="6"><div align="center"><font class="zoom_texto">N&atilde;o 
                        h&aacute; ocorr&ecirc;ncias para esse aluno</font></div></td>
                  </tr>
                  <%
else	
check=2
%>


                <tr class="zoom_tit"> 
                  <td width="180"> <div align="left">Data / Hora</div></td>
                  <td width="290"> <div align="left">Ocorr&ecirc;ncia</div></td>
                  <td> <div align="left">Professor</div></td>
                  <td> <div align="left">Disciplina</div></td>
                  <td><div align="center">Aula</div></td>
                  <td> <div align="center">Atendido por</div></td>
                </tr>




<%
While not RS3.EOF 

 if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if
  
data_ocor=RS3("DA_Ocorrencia")
hora_ocor=RS3("HO_Ocorrencia")
assunto=RS3("CO_Assunto")
ocorrencia=RS3("CO_Ocorrencia")
nu_aula=RS3("NU_Aula")
co_prof=RS3("CO_Professor")
materia=RS3("NO_Materia")
observa=RS3("TX_Observa")
cod_usr_ocorrencia=RS3("CO_Usuario")

dados_dtd= split(data_ocor, "/" )
dia_de= dados_dtd(0)
mes_de= dados_dtd(1)
ano_de= dados_dtd(2)

if dia_de<10 then
dia_de="0"&dia_de
end if
if mes_de<10 then
mes_de="0"&mes_de
end if

data_ocor=dia_de&"/"&mes_de&"/"&ano_de


	Set RS4 = Server.CreateObject("ADODB.Recordset")
	SQL4 = "select * from TB_Tipo_Assunto where CO_Assunto = '" & assunto &"'"
	set RS4 = CON0.Execute (SQL4)
	
no_assunto=RS4("NO_Assunto")	

	Set RS5 = Server.CreateObject("ADODB.Recordset")
	SQL5 = "select * from TB_Tipo_Ocorrencia where CO_Assunto = '" & assunto &"' AND CO_Ocorrencia="&ocorrencia
	set RS5 = CON0.Execute (SQL5)
if	RS5.EOF then
no_ocorrencia=""
else	
no_ocorrencia=RS5("NO_Ocorrencia")
end if

if co_prof="" or isnull(co_prof) then
	no_prof=""
else
		Set RS6 = Server.CreateObject("ADODB.Recordset")
		SQL6 = "select * from TB_Professor where CO_Professor = " & co_prof
		set RS6 = CON4.Execute (SQL6)
	
	if RS6.eof then	
	no_prof=""
	else
	no_prof=RS6("NO_Professor")
	end if
end if	

if materia="" or isnull(materia) then
	no_materia=""
else
	Set RS7 = Server.CreateObject("ADODB.Recordset")
	SQL7 = "select * from TB_Materia where CO_Materia = '" & materia &"'"
	set RS7 = CON0.Execute (SQL7)


	if RS7.eof then	
	no_materia=""
	else
	no_materia=RS7("NO_Materia")
	end if		
end if	

if cod_usr_ocorrencia="" or isnull(cod_usr_ocorrencia) then
	no_atendido=""
else
		Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& cod_usr_ocorrencia
		RSu.Open SQLu, CON

	IF RSu.EOF then
	else
		no_atendido=RSu("NO_Usuario")
	end if
end if

%>
                  <tr class="<%response.write(cor)%>">
                    <td width="180" class="zoom_texto"><div align="center">
                    <%response.Write(data_ocor&" &agrave;s "&hora_ocor)%></div></td>
                    <td width="290" class="zoom_texto"><div align="left">
                      <%response.Write(Server.URLEncode(no_ocorrencia))%>
                    </div></td>
                    <td class="zoom_texto"><div align="left">
                    <%response.Write(Server.URLEncode(no_prof))%></div></td>
                    <td class="zoom_texto"><div align="left"> 
                    <%response.Write(Server.URLEncode(no_materia))%></div></td>
                    <td class="zoom_texto"><div align="center">
                    <%response.Write(nu_aula)%></div></td>
                    <td class="zoom_texto"><div align="center">
                      <%response.Write(Server.URLEncode(no_atendido))%>
                    </div></td>
                  </tr>
<!--                  <tr class="<%response.write(cor)%>"> 
                    <td><div align="right" class="zoom_tit"><strong>Data e Hora:</strong></div></td>
                    <td class="zoom_texto"><div align="left">&nbsp; 
                      <%response.Write(data_ocor&" &agrave;s "&hora_ocor)%></div>
                    </td>
                    <td><div align="right" class="zoom_tit"><strong>Assunto:</strong></div></td>
                    <td class="zoom_texto"><div align="left">&nbsp; 
                      <%response.Write(Server.URLEncode(no_assunto))%></div>
                    </td>
                    <td><div align="right" class="zoom_tit"><strong>Ocorr&ecirc;ncia:</strong></div></td>
                    <td class="zoom_texto"><div align="left">&nbsp;
                      <%response.Write(Server.URLEncode(no_ocorrencia))%>
                    </div></td>
                  </tr>
                  <tr class="<%response.write(cor)%>"> 
                    <td width="10%"> <div align="right" class="zoom_tit"><strong>Aula:</strong></div></td>
                    <td width="16%" class="zoom_texto"><div align="left">&nbsp; 
                      <%response.Write(nu_aula)%></div>
                    </td>
                    <td width="15%"><div align="right" class="zoom_tit"><strong>Professor:</strong></div></td>
                    <td width="20%" class="zoom_texto"><div align="left">&nbsp; 
                      <%response.Write(Server.URLEncode(no_prof))%></div>
                    </td>
                    <td width="16%"><div align="right" class="zoom_tit"><strong>Disciplina:</strong></div></td>
                    <td width="23%" class="zoom_texto"><div align="left">&nbsp; 
                      <%response.Write(Server.URLEncode(no_materia))%></div>
                    </td>
                  </tr>
                  <tr class="<%response.write(cor)%>"> 
                    <td width="10%"> <div align="right" class="zoom_tit"><strong>Observa&ccedil;&atilde;o:</strong></div></td>
                    <td colspan="5" class="zoom_texto"><div align="left">&nbsp; 
                      <%response.Write(Server.URLEncode(observa))%></div>
                    </td>
                  </tr>
                  <tr class="<%response.write(cor)%>"> 
                    <td height="10" colspan="6" class="<%response.write(cor)%>">&nbsp;</td>
                  </tr>
-->                  <%
check=check+1				  
RS3.MOVENEXT
WEND
end if
%>
                </table>
</td>
  </tr>
</table>
</body>
</html>
