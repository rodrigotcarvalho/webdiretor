<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/banner.asp"-->
<%


Function cabecalho (nivel)
Session.LCID = 1046 
tp=session("tp")
nome = session("nome") 
acesso = session("acesso")
co_user = session("co_user")
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
escola=session("escola")
chave=session("chave")
dia_t=session("dia_t") 
hora_t=session("hora_t")
				alunos=Session("aluno_selecionado")
				Session("aluno_selecionado")=alunos
		
this_file = Request.ServerVariables("SCRIPT_NAME")
arPath = Split(this_Path, "/")


if nome = "" or acesso = "" or co_user = "" or ano_letivo = "" then
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
else
session("escola")=escola
session("nome") = nome
session("acesso") = acesso
session("co_user") = co_user
session("tp") = tp
session("ano_letivo") = ano_letivo
session("permissao") = permissao
session("sistema_local")=sistema_local
session("chave")=chave
session("grupo") = grupo
session("dia_t") = dia_t
session("hora_t") = hora_t
end if
call banner(nivel,this_file,sistema_local,nome,permissao,ano_letivo)
end function




Function navegacao (Conexao,chave, nivel)
session("chave")=chave
Select case nivel

case 0
origem ="Voc&ecirc; est&aacute; em Web Diretor"
case 1
chavearray=split(chave,"-")
sistema=chavearray(0)
		Set RSc1 = Server.CreateObject("ADODB.Recordset")
		SQLc1 = "SELECT * FROM TB_Sistema where CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc1.Open SQLc1, Conexao

sistema_nome=RSc1("TX_Descricao")
link_sistema=RSc1("CO_Pasta")

origem = "Voc&ecirc; est&aacute; em <a href='../inicio.asp' class='caminho' target='_self'>Web Diretor</a> > "&sistema_nome
case 2

chavearray=split(chave,"-")
sistema=chavearray(0)
modulo=chavearray(1)



		Set RSc1 = Server.CreateObject("ADODB.Recordset")
		SQLc1 = "SELECT * FROM TB_Sistema where CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc1.Open SQLc1, Conexao
		
		sistema_nome=RSc1("TX_Descricao")
		link_sistema=RSc1("CO_Pasta")



		Set RSc2 = Server.CreateObject("ADODB.Recordset")
		SQLc2 = "SELECT * FROM TB_Modulo where CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc2.Open SQLc2, Conexao

		modulo_nome=RSc2("TX_Descricao")
		link_modulo=RSc2("CO_Pasta")
	
	
origem = "Voc&ecirc; est&aacute; em <a href='../../inicio.asp' class='caminho' target='_self'>Web Diretor</a> > <a href='../../"&link_sistema&"/index.asp?nvg="&sistema&"' class='caminho' target='_self'>"&sistema_nome&"</a> > <a href='../"&link_modulo&"/index.asp?nvg="&chave&"' class='caminho' target='_self'>"&modulo_nome&"</a>"
		
case 3
chavearray=split(chave,"-")
sistema=chavearray(0)
modulo=chavearray(1)
setor=chavearray(2)
		Set RSc1 = Server.CreateObject("ADODB.Recordset")
		SQLc1 = "SELECT * FROM TB_Sistema where CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc1.Open SQLc1, Conexao
		
		sistema_nome=RSc1("TX_Descricao")
		link_sistema=RSc1("CO_Pasta")

		Set RSc2 = Server.CreateObject("ADODB.Recordset")
		SQLc2 = "SELECT * FROM TB_Modulo where CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc2.Open SQLc2, Conexao

		modulo_nome=RSc2("TX_Descricao")
		link_modulo=RSc2("CO_Pasta")
		
		Set RSc3 = Server.CreateObject("ADODB.Recordset")
		SQLc3 = "SELECT * FROM TB_Setor where CO_Setor = '"&setor&"' and CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc3.Open SQLc3, Conexao

		setor_nome=RSc3("TX_Descricao")
		link_setor=RSc3("CO_Pasta")

origem = "Voc&ecirc; est&aacute; em <a href='../../../inicio.asp' class='caminho' target='_self'>Web Diretor</a> > <a href='../../../"&link_sistema&"/index.asp?nvg="&sistema&"' class='caminho' target='_self'>"&sistema_nome&"</a> > <a href='../../"&link_modulo&"/index.asp?nvg="&sistema&"-"&modulo&"' class='caminho' target='_self'>"&modulo_nome&"</a> > <a href='../"&link_setor&"/index.asp?nvg="&chave&"' class='caminho' target='_self'>"&setor_nome&"</a>"

case 4
chavearray=split(chave,"-")
sistema=chavearray(0)
modulo=chavearray(1)
setor=chavearray(2)
funcao=chavearray(3)

grupo=session("grupo")
negado=request.querystring("neg")
ano_letivo = session("ano_letivo") 


		Set RSac = Server.CreateObject("ADODB.Recordset")
		SQLac = "SELECT * FROM TB_Autoriz_Grupo_Funcao where CO_Funcao = '"&funcao&"' and CO_Setor = '"&setor&"' and CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' and CO_Grupo= '"&grupo&"'"
		RSac.Open SQLac, Conexao
		
		Set RSal = Server.CreateObject("ADODB.Recordset")
		SQLal = "SELECT * FROM TB_Ano_Letivo where NU_Ano_Letivo='"&ano_letivo&"'"
		RSal.Open SQLal, Conexao
		
		sit_an=RSal("ST_Ano_Letivo")
		
		autoriza=RSac("TP_Acesso")
		if autoriza="0" and negado<>"1" then
		nvg=sistema&"-"&modulo&"-"&setor&"-"&funcao
		response.Redirect("../../../../inc/negado.asp?nvg="&nvg&"&neg=1")
		elseif autoriza="1" then
		session("trava")="s"
		elseif autoriza="5"  AND sit_an="L"then
		session("trava")="n"
		end if
		
		session("autoriza")=autoriza

		Set RSc1 = Server.CreateObject("ADODB.Recordset")
		SQLc1 = "SELECT * FROM TB_Sistema where CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc1.Open SQLc1, Conexao
		
		sistema_nome=RSc1("TX_Descricao")
		link_sistema=RSc1("CO_Pasta")

		Set RSc2 = Server.CreateObject("ADODB.Recordset")
		SQLc2 = "SELECT * FROM TB_Modulo where CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc2.Open SQLc2, Conexao

		modulo_nome=RSc2("TX_Descricao")
		link_modulo=RSc2("CO_Pasta")
		
		Set RSc3 = Server.CreateObject("ADODB.Recordset")
		SQLc3 = "SELECT * FROM TB_Setor where CO_Setor = '"&setor&"' and CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc3.Open SQLc3, Conexao

		setor_nome=RSc3("TX_Descricao")
		link_setor=RSc3("CO_Pasta")
		
		Set RSc4 = Server.CreateObject("ADODB.Recordset")
		SQLc4 = "SELECT * FROM TB_Funcao where CO_Funcao = '"&funcao&"' and CO_Setor = '"&setor&"' and CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc4.Open SQLc4, Conexao

		funcao_nome=RSc4("TX_Descricao")
		link_funcao=RSc4("CO_Pasta")

if negado="1" then
origem = "Voc&ecirc; est&aacute; em <a href='../inicio.asp' class='caminho' target='_self'>Web Diretor</a> > <a href='../"&link_sistema&"/index.asp?nvg="&sistema&"' class='caminho' target='_self'>"&sistema_nome&"</a> > <a href='../"&link_sistema&"/"&link_modulo&"/index.asp?nvg="&sistema&"-"&modulo&"'class='caminho' target='_self'>"&modulo_nome&"</a> > <a href='../"&link_sistema&"/"&link_modulo&"/"&link_setor&"/index.asp?nvg="&sistema&"-"&modulo&"-"&setor&"' class='caminho' target='_self'>"&setor_nome&"</a> > <a href='../"&link_sistema&"/"&link_modulo&"/"&link_setor&"/"&link_funcao&"/index.asp?nvg="&chave&"' class='caminho' target='_self'>"&funcao_nome&"</a>"

else
origem = "Voc&ecirc; est&aacute; em <a href='../../../../inicio.asp' class='caminho' target='_self'>Web Diretor</a> > <a href='../../../../"&link_sistema&"/index.asp?nvg="&sistema&"' class='caminho' target='_self'>"&sistema_nome&"</a> > <a href='../../../"&link_modulo&"/index.asp?nvg="&sistema&"-"&modulo&"'class='caminho' target='_self'>"&modulo_nome&"</a> > <a href='../../"&link_setor&"/index.asp?nvg="&sistema&"-"&modulo&"-"&setor&"' class='caminho' target='_self'>"&setor_nome&"</a> > <a href='../"&link_funcao&"/index.asp?nvg="&chave&"' class='caminho' target='_self'>"&funcao_nome&"</a>"
end if
		
end select

Session("caminho")=origem
'session("chave")=chave
chave=session("chave")
end function

Function menu_lateral (nivel)
tp=session("tp")

	if tp="R" then
		exibe_baq = "S"
	else	
		nu_unidade_baq = session("nu_unidade_baq")
		co_curso_baq = session("co_curso_baq")
		co_etapa_baq = session("co_etapa_baq")
		co_turma_baq = session("co_turma_baq")
		session("nu_unidade_baq") = nu_unidade_baq
		session("co_curso_baq") = co_curso_baq
		session("co_etapa_baq") = co_etapa_baq
		session("co_turma_baq") = co_turma_baq		
		
		co_curso_baq = co_curso_baq*1

		if isnumeric(co_etapa_baq) then
			co_etapa_baq=co_etapa_baq*1
		end if	
	
		if (co_curso_baq = 1 and co_etapa_baq< 6) then
			exibe_baq = "N"		
		else
			exibe_baq = "S"
		end if
		
	end if	

Select case nivel
case 0
%>
<table width="170" style="BORDER-TOP: #a8adb0 1px solid; BORDER-RIGHT: #a8adb0 1px solid; BORDER-LEFT: #a8adb0 1px solid; BORDER-BOTTOM: #a8adb0 1px solid">
  <tr> 
    <td width="164" height="24" class=menud  style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menue'" 
                onmouseout="this.className='menuf'"><div align="center"> 
        <% if tp="R" then%>
        <a href="inicio.asp" class="menu_lista">P&aacute;gina Inicial</a> 
        <%else%>
        <a href="inicio.asp" class="menu_lista">P&aacute;gina Inicial</a> 
        <%end if%>
      </div></td>
  </tr>
  <%IF tp="R" THEN%> 
  <tr> 
    <td height="24" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid"  bgcolor="#006666">
<div align="center"><font class="menu_lista">Coordena&ccedil;&atilde;o de Ensino</font></div></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="img/menu_seta.gif" width="19" height="15" /><a href="coordenacao/disciplinares/index.asp" class="menu_sublista">Ocorrências</a></td>
  </tr>
  <%end if%>
  <tr> 
    <td height="24" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid"  bgcolor="#006666"> 
      <div align="center"><font class="menu_lista">Aproveitamento Escolar</font></div></td>
  </tr>
  <!--
<tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="img/menu_seta.gif" width="19" height="15" /><a href="resultados/avprog/index.asp" class="menu_sublista">Avalia&ccedil;&otilde;es 
      Progressivas </a></td>
  </tr>
  
-->
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="img/menu_seta.gif" width="19" height="15" /><a href="resultados/boletim/index.asp" class="menu_sublista">Boletim</a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="img/menu_seta.gif" width="19" height="15" /><a href="resultados/avprog/index.asp" class="menu_sublista">Avaliações 
      Parciais</a></td>
  </tr>
  <% if exibe_baq = "S" then%>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="img/menu_seta.gif" width="19" height="15" /><a href="resultados/baq/index.asp" class="menu_sublista">Boletim de Av. Qualit.</a></td>
  </tr> 
<% end if%>   
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="img/menu_seta.gif" width="19" height="15" /><a href="resultados/graficos/index.asp" class="menu_sublista">Gráficos Comparativos</a></td>
  </tr>   
  <tr> 
<!--    <td height="24" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid"  bgcolor="#006666">
<div align="center"><font class="menu_lista">Informe Escolar</font></div></td>-->
    <td height="24" class=menud  style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menue'" onmouseout="this.className='menuf'"><div align="center"><a href="docs/index.asp" class="menu_lista">Informe Escolar</a></div></td>
  </tr>
 <!--   <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="img/menu_seta.gif" width="19" height="15" /><a href="docs/avcirc/index.asp" class="menu_sublista">Circulares</a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="img/menu_seta.gif" width="19" height="15" /><a href="docs/provgab/index.asp" class="menu_sublista">Avalia&ccedil;&otilde;es 
      e Gabaritos </a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="img/menu_seta.gif" width="19" height="15" /><a href="docs/reuniao/index.asp" class="menu_sublista">Reuni&atilde;o 
      de Pais </a></td>
  </tr>   
 <tr> 
 <td height="24" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid"  bgcolor="#006666">
<div align="center"><font class="menu_lista">Agenda do Ano Letivo</font></div></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="img/menu_seta.gif" width="19" height="15" /><a href="agenda/provas/index.asp" class="menu_sublista">Provas
      </a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="img/menu_seta.gif" width="19" height="15" /><a href="agenda/reunioes/index.asp" class="menu_sublista">Reuni&otilde;es</a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="img/menu_seta.gif" width="19" height="15" /><a href="agenda/eventos/index.asp" class="menu_sublista">Eventos
      </a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="img/menu_seta.gif" width="19" height="15" /><a href="agenda/feriados/index.asp" class="menu_sublista">Feriados
      e Recessos</a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="img/menu_seta.gif" width="19" height="15" /><a href="agenda/geral/index.asp" class="menu_sublista">Geral
      </a></td> 
  </tr> -->
  <tr> 
    <td height="24" class=menud  style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menue'" 
                onmouseout="this.className='menuf'"><div align="center"><a href="noticias/index.asp" class="menu_lista">Not&iacute;cias</a></div></td>
  </tr>
 <%  IF tp="R" and DateDiff("d", session("dt_exibe_pos_fin"), data_hoje)>=0 THEN%>
  <tr> 
    <td height="24" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid"  bgcolor="#006666">
<div align="center"><font class="menu_lista">Posi&ccedil;&atilde;o Financeira</font></div></td>
  </tr>
<!--  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="img/menu_seta.gif" width="19" height="15" /><a href="posfin/extrato/index.asp" class="menu_sublista">Extrato</a></td>
  </tr> -->
   <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="img/menu_seta.gif" width="19" height="15" /><a href="posfin/segvia/index.asp" class="menu_sublista"><!--2ª 
      Via de -->Boleto</a></td>
  </tr> 
<%END IF%>  	  
  <tr> 
    <td height="24" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid"  bgcolor="#006666">
<div align="center"><font class="menu_lista">Seguran&ccedil;a</font></div></td>
  </tr>
  <tr> 
    <td height="11" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="img/menu_seta.gif" width="19" height="15" /><a href="seguranca/senha_mail/index.asp" class="menu_sublista">Alterar 
      Senha </a></td>
  </tr>
</table>
<%case 1%>
<table width="170" style="BORDER-TOP: #a8adb0 1px solid; BORDER-RIGHT: #a8adb0 1px solid; BORDER-LEFT: #a8adb0 1px solid; BORDER-BOTTOM: #a8adb0 1px solid">
  <tr> 
    <td width="164" height="24" class=menud  style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menue'" 
                onmouseout="this.className='menuf'"><div align="center"> 
        <% if tp="R" then%>
        <a href="../inicio.asp?opt=ad" class="menu_lista">P&aacute;gina Inicial</a> 
        <%else%>
        <a href="../inicio.asp" class="menu_lista">P&aacute;gina Inicial</a> 
        <%end if%>
      </div></td>
  </tr>
  <%IF tp="R" THEN%> 
  <tr> 
    <td height="24" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid"  bgcolor="#006666">
<div align="center"><font class="menu_lista">Coordena&ccedil;&atilde;o de Ensino</font></div></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../img/menu_seta.gif" width="19" height="15" /><a href="../coordenacao/disciplinares/index.asp" class="menu_sublista">Ocorr&ecirc;ncias</a></td>
  </tr>
  <%end if%>
  <tr> 
    <td height="24" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid"  bgcolor="#006666">
<div align="center"><font class="menu_lista">Aproveitamento Escolar</font></div></td>
  </tr>
  <!--  
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../img/menu_seta.gif" width="19" height="15" /><a href="../resultados/avprog/index.asp" class="menu_sublista">Avalia&ccedil;&otilde;es 
      Progressivas </a></td>
  </tr>
-->
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../img/menu_seta.gif" width="19" height="15" /><a href="../resultados/boletim/index.asp" class="menu_sublista">Boletim</a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../img/menu_seta.gif" width="19" height="15" /><a href="../resultados/avprog/index.asp" class="menu_sublista">Avalia&ccedil;&otilde;es 
      Parciais</a></td>
  </tr>
  <% if exibe_baq = "S" then%>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../img/menu_seta.gif" width="19" height="15" /><a href="../resultados/baq/index.asp" class="menu_sublista">Boletim de Av. Qualit.</a></td>
  </tr> 
<% end if%>    
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../img/menu_seta.gif" width="19" height="15" /><a href="../resultados/graficos/index.asp" class="menu_sublista">Gráficos Comparativos</a></td>
  </tr>    
  <tr> 
<!--    <td height="24" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid"  bgcolor="#006666">
<div align="center"><font class="menu_lista">Informe Escolar</font></div></td>
-->  
    <td height="24" class=menud  style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menue'" onmouseout="this.className='menuf'"><div align="center"><a href="../docs/index.asp" class="menu_lista">Informe Escolar</a></div></td>
</tr>
<!--  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../img/menu_seta.gif" width="19" height="15" /><a href="../docs/avcirc/index.asp" class="menu_sublista">Circulares</a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../img/menu_seta.gif" width="19" height="15" /><a href="../docs/provgab/index.asp" class="menu_sublista">Avalia&ccedil;&otilde;es 
      e Gabaritos </a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../img/menu_seta.gif" width="19" height="15" /><a href="../docs/reuniao/index.asp" class="menu_sublista">Reuni&atilde;o 
      de Pais </a></td>
  </tr>   
  <tr> 
  <tr> 
 <td height="24" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid"  bgcolor="#006666">
<div align="center"><font class="menu_lista">Agenda do Ano Letivo</font></div></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../img/menu_seta.gif" width="19" height="15" /><a href="../agenda/provas/index.asp" class="menu_sublista">Provas
      </a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../img/menu_seta.gif" width="19" height="15" /><a href="../agenda/reunioes/index.asp" class="menu_sublista">Reuni&otilde;es</a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../img/menu_seta.gif" width="19" height="15" /><a href="../agenda/eventos/index.asp" class="menu_sublista">Eventos
      </a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../img/menu_seta.gif" width="19" height="15" /><a href="../agenda/feriados/index.asp" class="menu_sublista">Feriados
      e Recessos</a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../img/menu_seta.gif" width="19" height="15" /><a href="../agenda/geral/index.asp" class="menu_sublista">Geral
      </a></td> 
  </tr> -->
  </tr>
  <tr> 
    <td height="24" class=menud  style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menue'" 
                onmouseout="this.className='menuf'"><div align="center"><a href="../noticias/index.asp" class="menu_lista">Not&iacute;cias</a></div></td>
  </tr>
 <%  IF tp="R" and DateDiff("d", session("dt_exibe_pos_fin"), data_hoje)>=0 THEN%>
  <TR>  
    <td height="24" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid"  bgcolor="#006666">
<div align="center"><font class="menu_lista">Posi&ccedil;&atilde;o Financeira</font></div></td>
  </tr>
<!--  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../img/menu_seta.gif" width="19" height="15" /><a href="../posfin/extrato/index.asp" class="menu_sublista">Extrato</a></td>
  </tr> -->
   <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../img/menu_seta.gif" width="19" height="15" /><a href="../posfin/segvia/index.asp" class="menu_sublista"><!--2ª 
      Via de -->Boleto</a></td>
</tr> 
<%END IF%>	  
  <tr> 
    <td height="24" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid"  bgcolor="#006666">
<div align="center"><font class="menu_lista">Seguran&ccedil;a</font></div></td>
  </tr>
  <tr> 
    <td height="11" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><span class="style1"><img src="../img/menu_seta.gif" width="19" height="15" /><a href="../seguranca/senha_mail/index.asp" class="menu_sublista">Alterar 
      Senha </a></span></td>
  </tr>
</table>
<%case 2%>
<table width="170" style="BORDER-TOP: #a8adb0 1px solid; BORDER-RIGHT: #a8adb0 1px solid; BORDER-LEFT: #a8adb0 1px solid; BORDER-BOTTOM: #a8adb0 1px solid">
  <tr> 
    <td width="164" height="24" class=menud  style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menue'" 
                onmouseout="this.className='menuf'"><div align="center"> 
        <% if tp="R" then%>
        <a href="../../inicio.asp?opt=ad" class="menu_lista">P&aacute;gina Inicial</a> 
        <%else%>
        <a href="../../inicio.asp" class="menu_lista">P&aacute;gina Inicial</a> 
        <%end if%>
      </div></td>
  </tr>
  <%IF tp="R" THEN%> 
  <tr> 
    <td height="24" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid"  bgcolor="#006666">
<div align="center"><font class="menu_lista">Coordena&ccedil;&atilde;o de Ensino</font></div></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../../img/menu_seta.gif" width="19" height="15" /><a href="../../coordenacao/disciplinares/index.asp" class="menu_sublista">Ocorr&ecirc;ncias</a></td>
  </tr>
  <%end if%>
  <tr> 
    <td height="24" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid"  bgcolor="#006666">
<div align="center"><font class="menu_lista">Aproveitamento Escolar</font></div></td>
  </tr>
  <!--  
    <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../../img/menu_seta.gif" width="19" height="15" /><a href="../../resultados/avprog/index.asp" class="menu_sublista">Avalia&ccedil;&otilde;es 
      Progressivas </a></td>
  </tr>
-->
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../../img/menu_seta.gif" width="19" height="15" /><a href="../../resultados/boletim/index.asp" class="menu_sublista">Boletim</a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../../img/menu_seta.gif" width="19" height="15" /><a href="../../resultados/avprog/index.asp" class="menu_sublista">Avalia&ccedil;&otilde;es 
      Parciais</a></td>
  </tr>
  <% if exibe_baq = "S" then%>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../../img/menu_seta.gif" width="19" height="15" /><a href="../../resultados/baq/index.asp" class="menu_sublista">Boletim de Av. Qualit.</a></td>
  </tr> 
<% end if%>    
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../../img/menu_seta.gif" width="19" height="15" /><a href="../../resultados/graficos/index.asp" class="menu_sublista">Gráficos Comparativos</a></td>
  </tr>    
  <tr> 
<!--    <td height="24" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid"  bgcolor="#006666">
<div align="center"><font class="menu_lista">Informe Escolar</font></div></td>-->
    <td height="24" class=menud  style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menue'" onmouseout="this.className='menuf'"><div align="center"><a href="../../docs/index.asp" class="menu_lista">Informe Escolar</a></div></td>
  </tr>
<!--  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../../img/menu_seta.gif" width="19" height="15" /><a href="../../docs/avcirc/index.asp" class="menu_sublista">Circulares</a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../../img/menu_seta.gif" width="19" height="15" /><a href="../../docs/provgab/index.asp" class="menu_sublista">Avalia&ccedil;&otilde;es 
      e Gabaritos </a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../../img/menu_seta.gif" width="19" height="15" /><a href="../../docs/reuniao/index.asp" class="menu_sublista">Reuni&atilde;o 
      de Pais </a></td>
  </tr>  
  <tr> 
 <td height="24" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid"  bgcolor="#006666">
<div align="center"><font class="menu_lista">Agenda do Ano Letivo</font></div></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../../img/menu_seta.gif" width="19" height="15" /><a href="../../agenda/provas/index.asp" class="menu_sublista">Provas
      </a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../../img/menu_seta.gif" width="19" height="15" /><a href="../../agenda/reunioes/index.asp" class="menu_sublista">Reuni&otilde;es</a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../../img/menu_seta.gif" width="19" height="15" /><a href="../../agenda/eventos/index.asp" class="menu_sublista">Eventos
      </a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../../img/menu_seta.gif" width="19" height="15" /><a href="../../agenda/feriados/index.asp" class="menu_sublista">Feriados
      e Recessos</a></td>
  </tr>
  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../../img/menu_seta.gif" width="19" height="15" /><a href="../../agenda/geral/index.asp" class="menu_sublista">Geral
      </a></td> 
  </tr> -->
  <tr> 
    <td height="24" class=menud  style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menue'" 
                onmouseout="this.className='menuf'"><div align="center"><a href="../../noticias/index.asp" class="menu_lista">Not&iacute;cias</a></div></td>
  </tr>
  <%  IF tp="R" and DateDiff("d", session("dt_exibe_pos_fin"), data_hoje)>=0 THEN%>
  <tr>  
    <td height="24" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid"  bgcolor="#006666">
<div align="center"><font class="menu_lista">Posi&ccedil;&atilde;o Financeira</font></div></td>
  </tr>
<!--  <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../../img/menu_seta.gif" width="19" height="15" /><a href="../../posfin/extrato/index.asp" class="menu_sublista">Extrato</a></td>
  </tr> -->
   <tr> 
    <td height="24" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><img src="../../img/menu_seta.gif" width="19" height="15" /><a href="../../posfin/segvia/index.asp" class="menu_sublista"><!--2ª 
      Via de -->Boleto</a></td>
  </tr>	   
<%end if%>	  
  <tr> 
    <td height="24" style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid"  bgcolor="#006666">
<div align="center"><font class="menu_lista">Seguran&ccedil;a</font></div></td>
  </tr>
  <tr> 
    <td height="11" class=menua style="BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #ffffff 1px solid; BORDER-BOTTOM: #999999 1px solid" onmouseover="this.className='menub'" 
                onmouseout="this.className='menuc'"><span class="style1"><img src="../../img/menu_seta.gif" width="19" height="15" /><a href="../../seguranca/senha_mail/index.asp" class="menu_sublista">Alterar 
      Senha </a></span></td>
  </tr>
</table>
<%end select
end function

FUNCTION linkFuncao(Conexao,sistema,modulo,setor,funcao,nivel)


		Set RSc1 = Server.CreateObject("ADODB.Recordset")
		SQLc1 = "SELECT * FROM TB_Sistema where CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc1.Open SQLc1, Conexao
		
		sistema_nome=RSc1("TX_Descricao")
		link_sistema=RSc1("CO_Pasta")

		Set RSc2 = Server.CreateObject("ADODB.Recordset")
		SQLc2 = "SELECT * FROM TB_Modulo where CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc2.Open SQLc2, Conexao

		modulo_nome=RSc2("TX_Descricao")
		link_modulo=RSc2("CO_Pasta")
		
		Set RSc3 = Server.CreateObject("ADODB.Recordset")
		SQLc3 = "SELECT * FROM TB_Setor where CO_Setor = '"&setor&"' and CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc3.Open SQLc3, Conexao

		setor_nome=RSc3("TX_Descricao")
		link_setor=RSc3("CO_Pasta")
		
		Set RSc4 = Server.CreateObject("ADODB.Recordset")
		SQLc4 = "SELECT * FROM TB_Funcao where CO_Funcao = '"&funcao&"' and CO_Setor = '"&setor&"' and CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"' order by NU_Pos"
		RSc4.Open SQLc4, Conexao

		funcao_nome=RSc4("TX_Descricao")
		link_funcao=RSc4("CO_Pasta")


link_funcao="../../../../"&link_sistema&"/"&link_modulo&"/"&link_setor&"/"&link_funcao
session("link_funcao")=link_funcao
end function
'///////////////////////////////////////////////    MENSAGENS     //////////////////////////////////////////////////////////////////////////////


FUNCTION mensagens(nivel,msg,tab,cod)
escola=Session("escola")

SELECT CASE msg
'Mensagens Gerais de 0 a 299
case 0
wrt = "Escolha uma das opções abaixo"

case 1
wrt = "Selecione uma unidade e um curso "

case 2
wrt = "Selecione uma etapa e uma turma."

case 3
wrt = "Selecione uma etapa, uma turma, um período e uma avaliação."

case 4
wrt = "Para consultar é necessário selecionar uma etapa!"

case 5
wrt = "Esta função permite você fazer contato com a equipe técnica que realiza a manutenção do sistema Web Diretor. Utilize sempre que possível este canal para nos transmitir alguma informação relevante sobre o funcionamento desse produto. Obrigado pela sua atenção!"

case 6
wrt = "Mensagem enviada."

case 7
wrt = "Escolha um novo usuário."

case 8
wrt = "Escolha uma nova senha."

case 9
wrt = "Usuário alterado com sucesso."

case 10
wrt = "Senha alterada com sucesso."

case 11
wrt = "Selecione uma disciplina e um período."

case 12
wrt = "E-mail alterado com sucesso."

case 13
wrt = "Usuário já existe!"

case 14
wrt = "Digite seu novo endereço de correio eletrônico"

case 15
wrt = "Endereço de correio eletrônico já existe!"

case 16
wrt = "Selecione uma etapa, uma turma e um período."

case 17
wrt = "Selecione uma etapa e um período."

case 18
wrt = "Gráfico comparativo."

case 19
wrt = "Selecione uma etapa, uma disciplina e um período."

case 20
wrt = "Selecione uma etapa"

'alunos de 300 a 599
case 300
wrt = "Para consultar os dados do Aluno digite a Matrícula ou Nome e clique no bot&atilde;o Procurar."

' listagem de alunos

case 301
wrt = "Escolha um Aluno para consultar o cadastro."

case 302
wrt = "Verifique os dados do Aluno."

case 303
wrt = "Não foi encontrado nenhum Aluno com este código."

' erro na busca por nome
case 304
wrt = "Não foi encontrado nenhum Aluno com este nome."

case 305
wrt = "Lista de alunos associados a turma abaixo."

case 306
wrt = "Verifique os dados dos familiares."

case 307
wrt = "Selecione uma unidade e um mês."

case 308
wrt = "Comparar Turma por Média Geral."

case 309
wrt = "Verifique os dados do Aluno e escolha uma disciplina e um período."


'professores de 600 a 899

case 600
wrt =  "Os Professores em vermelho estão inativos. A mensagem 'não cadastrado' indica que não existe professor associado àquela disciplina naquela turma"
wrt = wrt &"<br>A mensagem 'nome em branco' indica que o nome do professor não está registrado no cadastro. Para bloquear a planilha clique na letra 'N' do período escolhido"

case 601
wrt = "Confirma o " 
if opt="blq" then
wrt= wrt &"BLOQUEIO"
else
wrt= wrt &"DESBLOQUEIO"
end if
wrt= wrt &" das notas do trimestre "&periodo&" de "&no_materia&", Unidade:"&no_unidade&" - "&no_etapa&" do "&no_curso&" Turma "&turma&""

case 602
if orig=01 then
act= "bloqueada"
elseif orig=02 then
act= "desbloqueada"
end if

wrt = "Planilha "&act&" com sucesso!"

case 603
wrt = "Avaliações não lançadas"

case 604
wrt = "Para consultar a Grade de aulas digite o C&oacute;digo ou Nome de um Professor e clique no bot&atilde;o Procurar."
wrt = wrt &"<br>Se preferir obter uma lista completa de TODOS os professores clique <a href='index.asp?opt=listall&nvg="&nvg&"' class='linkum'>aqui</a>"

case 605
wrt = "Não foi encontrado nenhum professor com este código."

case 606
wrt = "Escolha um professor para consultar a Grade de Aulas. Os Professores em vermelho estão inativos."

case 607
wrt = "Para atualizar os dados do Professor digite o C&oacute;digo ou Nome e clique no bot&atilde;o Procurar."
wrt = wrt &"Se preferir adicionar um NOVO professor clique <a href='grade_cp1.asp?or=02&nvg="&nvg&"' class='linkum'>aqui</a>."
wrt = wrt &"<BR>Se preferir obter uma lista completa de TODOS os professores clique <a href='index.asp?opt=listall&nvg="&nvg&"' class='linkum'>aqui</a>"

case 608
wrt = "Confirme o professor para consultar a Grade de Aulas."

case 609
wrt = "O período relacionado pela letra 'S' indica que a planilha está Bloqueada e 'N' que está Desbloqueada."

case 610
wrt = "Não foi encontrado nenhum professor com este código."

case 611
wrt = "Não foi encontrado nenhum professor com este nome."

case 612
wrt = "Escolha um professor para atualizar o cadastro. Os Professores em vermelho estão inativos."

case 613
wrt = "Confirme se é o professor correto para atualizar o cadastro."

case 614
wrt = "Preencha cuidadosamente os dados do Professor e click no bot&atilde;o CONFIRMAR para atualizar o cadastro"

case 615
wrt = "Professor código "&cod_cons&" e usuário "&escola&co_usr_prof&" incluído com sucesso!"

case 616
wrt = "Dados do Professor código "&cod_cons&" alterados com sucesso!"

case 617
wrt = "Selecione a Data e a Hora as quais você deseja iniciar o monitoramento de notas e clique em iniciar."

case 618
mes_mnl=mes_mnl*1
min_mnl=min_mnl*1
			  if mes_mnl< 10 then
			  mes_wrt="0"&mes_mnl
			  else
			  mes_wrt=mes_mnl					  
			  end if 
					  
			  if min_mnl< 10 then
			  min_wrt="0"&min_mnl
			  else
			  min_wrt=min_mnl					  
			  end if 
wrt = "Inicio da monitoração a partir do dia "&dia_mnl&"/"&mes_wrt&"/"&ano_mnl&" as "&hora_mnl&":"&min_wrt&" Dados atualizados a cada minuto."

case 619
wrt = "Não foram encontradas turmas cadastradas para você. Entre em contato com o seu coordenador."


case 620
if errou="pv1" or errou="pv2" or errou="pv3" or errou="pv4" or errou="pv5" or errou="pv6" Then
wrt = "Valor inválido para o campo  "&errado
elseif errou="sp" Then
wrt = "Soma dos Pesos maior que 10"
elseif errou="pt" Then
wrt = "Um dos pesos tem valor inválido"
elseif errou="pr1pr2" Then
wrt = "Soma das Pr's maior que 10"
else
wrt = "Valor inválido para o campo  "&errado&"  do número de chamada <b>"&errante&"</b>"
end if

' erro na busca por código
case 621
wrt = "Você está " 
if opt="cln" then
wrt= wrt &"comunicando"
else
wrt= wrt &"lançando"
end if
wrt= wrt &" notas do trimestre "&periodo&" de "&no_materia&", Unidade:"&no_unidades&" - "&no_serie&" do "&no_grau&" Turma "&turma&""

case 622
wrt = "Notas lançadas com sucesso."

case 623
wrt = "Comunicado efetuado!"

case 624
wrt = "Estas notas j&aacute; foram lan&ccedil;adas.Para alter&aacute;-las pe&ccedil;a autoriza&ccedil;&atilde;o ao coordenador"

case 625
wrt = "Escolha um Coordenador para consultar os Professores sob sua coordenação."

case 626
wrt = "Os Professores em vermelho estão inativos. A mensagem 'não cadastrado'indica que não existe professor associado àquela disciplina naquela turma"
wrt = wrt &"<br>A mensagem 'nome em branco' indica que o nome do professor não está registrado no cadastro"

case 627
wrt = "Para excluir, selecione uma ou mais disciplinas e clique em excluir.<br>Para incluir uma nova disciplina na Grade de Aulas, selecione uma unidade e um curso."

case 628
wrt = "Disciplina incluída com sucesso"

case 629
wrt = "Disciplina excluída com sucesso"

case 630
wrt = "Não é possível marcar uma disciplina na Grade de Aulas e selecionar uma unidade e um curso ao mesmo tempo.<br>Por favor selecione somente disciplina(s) para excluir ou selecione uma unidade para incluir uma nova disciplina na Grade de Aulas"

case 631
wrt = "Selecione uma disciplina, um modelo e um coordenador."

case 632
wrt = "Para atualizar é necessário selecionar uma disciplina,um modelo e um coordenador"

case 633
wrt = "Verifique os dados preenchidos e clique no botão Confirmar para continuar a inclusão ou no botão Alterar para voltar e modificar algum dado."


case 634
wrt = "Verifique as disciplinas selecionadas e clique no botão confirmar para Excluir ou no botão Cancelar para voltar e modificar algum dado."

case 635
wrt = "Professores que não comunicaram."

case 636
wrt = "Para imprimir clique <a class='linkum' href='#' onClick=MM_openBrWindow('imprime.asp?or=01&obr="&obr&"&p=p','','status=yes,menubar=yes,scrollbars=yes,resizable=yes,width=1030,height=500,top=50,left=50')>aqui</a>."

case 637
wrt = "Escolha um professor e um período."

case 638
wrt =  "Os Professores em vermelho estão inativos. A mensagem 'não cadastrado' indica que não existe professor associado àquela disciplina naquela turma"
wrt = wrt &"<br>A mensagem 'nome em branco' indica que o nome do professor não está registrado no cadastro. Clique no nome da disciplina para ver o mapa de resultado."

case 639
wrt = "Arquivo "& fl &" enviado com sucesso."

case 640
wrt = "Atenção! Estas notas j&aacute; foram lan&ccedil;adas pelo professor."



'Mensagens de sistema de 9700 a 9999
case 9700
wrt = "Acesso não permitido a esta função!"

case 9701
wrt = "Acesso permitido somente para consulta!"

case 9702
wrt = "Para imprimir clique <a class='linkum' href='#' onClick=MM_openBrWindow('imprime.asp?or=01&obr="&obr&"&p=p','','status=yes,menubar=yes,scrollbars=yes,resizable=yes,width=1030,height=500,top=50,left=50')>aqui</a>."

case 9703
wrt = "Aten&ccedil;&atilde;o! Ano Letivo est&aacute; Finalizado. As fun&ccedil;&otilde;es s&oacute; poder&atilde;o ser consultadas!<a href=../inicio.asp><img src=../img/ok.gif align=absbottom></a>"
end select




SELECT CASE tab


' primeira tela
case 0

%>
<table width="1000" height="52" border="3" align="center" cellpadding="0" cellspacing="0" bordercolor="#EEEEEE" class="aviso1">
  <tr> 
            
    <td height="46"> <div align="center"> 
      <%SELECT CASE nivel
				case 0%>
      <img src="img/atencao.gif" width="23" height="25" align="absmiddle"> 
      <%case 1%>
      <img src="../img/atencao.gif" width="23" height="25" align="absmiddle"> 
      <%		case 2%>
      <img src="../../img/atencao.gif" width="23" height="25" align="absmiddle"> 
      <%		case 3%>
      <img src="../../../img/atencao.gif" width="23" height="25" align="absmiddle"> 
      <%		case 4%>
      <img src="../../../../img/atencao.gif" width="23" height="25" align="absmiddle"> 
      <%end select%>
		  <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>
          <%response.Write(wrt)%>
          </strong></font></div>
                </div></td>
          </tr>
        </table>
		
<%
' erro
case 1
%>
<table width="1000" height="30" border="3" align="center" cellpadding="0" cellspacing="0" bordercolor="#EEEEEE" bgcolor="#FFE8E8" class="aviso2">
  <tr> 
            <td> <div align="center"> 
                <p>
		<%SELECT CASE nivel
				case 0%>
				<img src="img/pare.gif" width="28" height="25" align="absmiddle">
		<%case 1%>
				<img src="../img/pare.gif" width="28" height="25" align="absmiddle">
		<%case 2%>
				<img src="../../img/pare.gif" width="28" height="25" align="absmiddle">
		<%case 3%>
				<img src="../../../img/pare.gif" width="28" height="25" align="absmiddle">
		<%case 4%>
				<img src="../../../../img/pare.gif" width="28" height="25" align="absmiddle">												
		<%end select%>
                <font color="#CC0000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><%response.Write(wrt)%></strong></font> 
                </p>
              </div></td>
          </tr>
        </table>
<%
' inclusão / alteração de dados
case 2
%>
<table width="1000" height="30" border="3" align="center" cellpadding="0" cellspacing="0" bordercolor="#EEEEEE" bgcolor="#F2F9EE">
  <tr> 
            <td> <div align="center"> 
        <p>
		<%SELECT CASE nivel
						case 0%>
				<img src="img/atencao2.gif" width="23" height="25" align="absmiddle">
		<%case 1%>
		<img src="../img/atencao2.gif" width="23" height="25" align="absmiddle"> 
		<%case 2%>
				<img src="../../img/atencao2.gif" width="23" height="25" align="absmiddle">
		<%case 3%>
				<img src="../../../img/atencao2.gif" width="23" height="25" align="absmiddle">
		<%case 4%>
				<img src="../../../../img/atencao2.gif" width="23" height="25" align="absmiddle">		
		<%end select%>		
          <font color="#CC0000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
          <%response.Write(wrt)%>
          </strong></font> </p>
              </div></td>
          </tr>
        </table>
<%
end select

End Function

' Verifica Acesso
Function VerificaAcesso (CON,chave,nivel)
'0 - Sem Acesso, 1 - Só Consulta , 2 - Só Inclui,  3 - Só Altera, 4 - Só Exclui e  5 - Acesso Completo
chavearray=split(chave,"-")
sistema=chavearray(0)
modulo=chavearray(1)
setor=chavearray(2)
funcao=chavearray(3)
grupo=session("grupo")
		
		Set RSac = Server.CreateObject("ADODB.Recordset")
		SQLac = "SELECT * FROM TB_Autoriz_Grupo_Funcao where CO_Grupo = '"&grupo&"' and CO_Funcao = '"&funcao&"' and CO_Setor = '"&setor&"' and CO_Modulo = '"&modulo&"' and CO_Sistema = '"&sistema&"'"
		RSac.Open SQLac, CON

		funcao_acesso=RSac("TP_Acesso")

Select case funcao_acesso
case 0
autoriza="no"

case 1
autoriza="con"

case 2
autoriza="in"

case 3
autoriza="al"

case 4
autoriza="ex"

case 5
autoriza="full"
end select

Session("autoriza")=autoriza
End Function

'///////////////////////////////////////////////    Grava LOG  //////////////////////////////////////////////////////////////
Function GravaLog (nvg,outro)

'onde = Split(nvg, "-")
'stm=onde(0)
'mdl=onde(1)
'str=onde(2)
'fc=onde(3)

	co_usr = session("co_user")
	tp_entrada=nvg
	
	hora = DatePart("h", now) 
	min = DatePart("n", now) 
	dia = DatePart("d", now) 
	mes = DatePart("m", now) 
	ano = DatePart("yyyy", now)

if dia<10 then 
dia = "0"&dia
end if

if mes<10 then
mes = "0"&mes
end if

if hora<10 then 
hora = "0"&hora
end if

if min<10 then
min = "0"&min
end if	
	 
	gravahora= hora&":"&min
	gravadata= dia&"/"&mes&"/"&ano

tx_desc = outro

		Set CONL = Server.CreateObject("ADODB.Connection") 
		ABRIRL = "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONL.Open ABRIRL	

Set RSL = server.createobject("adodb.recordset")

RSL.open "TB_Log_Ocorrencias", CONL, 2, 2 'which table do you want open

RSL.addnew
RSL("TP_entrada") = tp_entrada
RSL("CO_Usuario") = co_usr
RSL("HO_ult_Acesso") = gravahora
RSL("DA_Ult_Acesso") = gravadata

RSL.update
  
set RSL=nothing

end function





Function regra_aprovacao (unidade,curso,etapa,turma,total_nota,m1_aluno,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2)

if total_nota>=4 then
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set RSra = Server.CreateObject("ADODB.Recordset")
	SQLra = "SELECT * FROM TB_Regras_Aprovacao where CO_Curso = '"&curso&"' and CO_Etapa = '"&etapa&"'"
	RSra.Open SQLra, CON0	
			
	valor_m1=RSra("NU_Valor_M1")
	res1_1=RSra("NO_Expr_Ma_Igual_M1")
	m1_maior_igual=RSra("NU_Int_Me_Ma_Igual_M1")
	m1_menor=RSra("NU_Int_Me_Me_M1")
	res1_2=RSra("NO_Expr_Int_M1_V")
	res1_3=RSra("NO_Expr_Int_M1_F")
	peso_m2_m1=RSra("NU_Peso_Media_M2_M1")
	peso_m2_m2=RSra("NU_Peso_Media_M2_M2")
	
	valor_m2=RSra("NU_Valor_M2")
	res2_1=RSra("NO_Expr_Ma_Igual_M2")
	m2_maior_igual=RSra("NU_Int_Me_Ma_Igual_M2")
	m2_menor=RSra("NU_Int_Me_Me_M2")
	res2_2=RSra("NO_Expr_Int_M2_V")
	res2_3=RSra("NO_Expr_Int_M2_F")
	peso_m3_m1=RSra("NU_Peso_Media_M3_M1")
	peso_m3_m2=RSra("NU_Peso_Media_M3_M2")
	peso_m3_m3=RSra("NU_Peso_Media_M3_M3")
	
	valor_m3=RSra("NU_Valor_M3")
	res3_1=RSra("NO_Expr_Ma_Igual_M3")
	m3_maior_igual=RSra("NU_Int_Me_Ma_Igual_M3")
	m3_menor=RSra("NU_Int_Me_Me_M3")
	res3_2=RSra("NO_Expr_Int_M3_V")
	res3_3=RSra("NO_Expr_Int_M3_F")
		
	m1_aluno=m1_aluno*1	
	valor_m1=valor_m1*1
	
	
	if m1_aluno >= valor_m1 then
	Session("resultado_1")=res1_1
	elseif (m1_aluno >= m1_maior_igual) and (m1_aluno <m1_menor) then
	Session("resultado_1")=res1_2
	else
	Session("resultado_1")=res1_3
	end if
	
	if Session("resultado_1")=res1_1 then
	else
	
				if nota_aux_m2_1="&nbsp;" then
				m2_aluno="&nbsp;"
				else					
				m1_aluno_peso=m1_aluno*peso_m2_m1
				nota_aux_m2_1_peso=nota_aux_m2_1*peso_m2_m2
				m2_aluno=(m1_aluno_peso+nota_aux_m2_1_peso)/(peso_m2_m1+peso_m2_m2)
							decimo = m2_aluno - Int(m2_aluno)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m2_aluno) + 1
									m2_aluno=nota_arredondada
								Else
									nota_arredondada = Int(m2_aluno)
									m2_aluno=nota_arredondada					
								End If
							m2_aluno = formatNumber(m2_aluno,0)
				end if
		
		if m2_aluno<>"&nbsp;" then		
		m2_aluno=m2_aluno*1
		valor_m2=valor_m2*1	
		end if		
		'response.write(m2_aluno &">="& valor_m2&"<BR>" )
		if m2_aluno >= valor_m2 and nota_aux_m2_1<>"&nbsp;" then
		Session("resultado_2")=res2_1
		elseif (m2_aluno >= m2_maior_igual) and (m2_aluno <m2_menor) and nota_aux_m2_1<>"&nbsp;" then
		Session("resultado_2")=res2_2
		elseif nota_aux_m2_1<>"&nbsp;" then
		Session("resultado_2")=res2_3
		else
		Session("resultado_2")="&nbsp;"
		end if
	
		if Session("resultado_2")=res2_1 then
		else
				if nota_aux_m3_1="&nbsp;" then
				m3_aluno="&nbsp;"
				else
				m1_aluno_peso=m1_aluno*peso_m3_m1					
				m2_aluno_peso=m2_aluno*peso_m3_m2
				nota_aux_m3_1_peso=nota_aux_m3_1*peso_m3_m3
				m3_aluno=(m1_aluno_peso+m2_aluno_peso+nota_aux_m3_1_peso)/(peso_m3_m1+peso_m3_m2+peso_m3_m3)
	
							decimo = m3_aluno - Int(m3_aluno)
								If decimo >= 0.5 Then
									nota_arredondada = Int(m3_aluno) + 1
									m3_aluno=nota_arredondada
								Else
									nota_arredondada = Int(m3_aluno)
									m3_aluno=nota_arredondada					
								End If
							m3_aluno = formatNumber(m3_aluno,0)
				end if	
			if m3_aluno<>"&nbsp;" then
			m3_aluno=m3_aluno*1
			valor_m3=valor_m3*1
			end if
			
			if m3_aluno >= valor_m3 and nota_aux_m3_1<>"&nbsp;" then
			Session("resultado_3")=res3_1
			elseif (m3_aluno >= m3_maior_igual) and (m3_aluno <m3_menor) and nota_aux_m3_1<>"&nbsp;" then
			Session("resultado_3")=res3_2
			elseif nota_aux_m3_1<>"&nbsp;" then
			Session("resultado_3")=res3_3
			else
			Session("resultado_3")="&nbsp;"
			end if
		end if	
	end if
	Session("M2")=m2_aluno
	Session("M3")=m3_aluno
else
	Session("resultado_1")="&nbsp;"
	Session("resultado_2")="&nbsp;"
	Session("resultado_3")="&nbsp;"
	Session("M2")="&nbsp;"
	Session("M3")="&nbsp;"	
end if		
end function
%>
