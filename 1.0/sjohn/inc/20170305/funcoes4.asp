<!--#include file="caminhos.asp"-->
<!--#include file="../../global/conta_alunos.asp"-->
<%Function faltas (CAMINHOa,CAMINHOn,unidade,curso,etapa,turma,co_materia,periodo,ano_letivo,co_usr,opcao,erro)

		Set CON_N = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIR3
		
		Set CON_A = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_A.Open ABRIR
		
		Set CON_wr = Server.CreateObject("ADODB.Connection") 
		ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_wr.Open ABRIR_wr
		
ntvmla0= 59
ntvmlb0= 59
ntvmlc0= 69
ntvmla=ntvmla0
ntvmlb=ntvmlb0
ntvmlc=ntvmlc0

ntvmla2 = formatNumber(ntvmla0,1)
ntvmlb2 = formatNumber(ntvmlb0,1)
ntvmlc2 = formatNumber(ntvmlc0,1)

notas_a_lancar = 4
'qtd_alunos=contalunos(CAMINHO_al,ano_letivo,unidade,curso,etapa,turma,"C")
qtd_alunos=contalunos(CAMINHO_al,ano_letivo,unidade,curso,etapa,turma,"TODOS")

trava=session("trava")
select case opcao

case 1
tb="TB_Frequencia_Periodo"

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso&"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)

calc = calc*1


		

if errou="pt" or errou="pp" Then
	classe_peso = "tb_fundo_linha_erro"
else
	classe_peso = "tb_fundo_linha_peso"
end if

%> 
 <form action="bda.asp?opt=cln" name="nota" method="post" onSubmit="return checksubmit()">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20" class="tb_subtit"> <div align="center">N&ordm;</div></td>
    <td width="780" class="tb_subtit"> 
      <div align="left">Nome</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B1</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B2</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B3</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B4</div></td>
  </tr>
  <%
check = 2
nu_chamada_ckq = 0
linha_tabela=1
While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")
	
	if nu_matricula = calc then
	  cor = "tb_fundo_linha_erro" 
	 onblur="mudar_cor_blur_erro"	  
	else
	 if check mod 2 =0 then
	  cor = "tb_fundo_linha_par" 
	  onblur="mudar_cor_blur_par"
	 else cor ="tb_fundo_linha_impar"
	 onblur="mudar_cor_blur_impar"
	  end if 
	end if


if (nu_chamada_ckq <>nu_chamada - 1) then
	teste_nu_chamada = nu_chamada-nu_chamada_ckq
	teste_nu_chamada=teste_nu_chamada-1
	
			for k=1 to teste_nu_chamada 
				nu_chamada_falta=nu_chamada_ckq+1
		
		%>
  <tr> 
    <td width="20" class="tb_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"></td>
    <td width="780" class="tb_fundo_linha_falta"><input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="falta" />
    &nbsp;</td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%  
				nu_chamada_ckq=nu_chamada_falta
			next	
		nu_chamada_ckq=nu_chamada	
			
				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
				Set RS2 = CON_A.Execute(SQL_A)
				
			NO_Aluno= RS2("NO_Aluno")
		
		
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
				Set RS3 = CON_N.Execute(SQL_N)
			
		if RS3.EOF then 
		 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780"  > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if
		%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
		end if
else
nu_chamada_ckq=nu_chamada

			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
			Set RS2 = CON_A.Execute(SQL_A)
			
		NO_Aluno= RS2("NO_Aluno")
	
			Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
			Set RS3 = CON_N.Execute(SQL_N)
		
		if RS3.EOF then 
	 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if		
						%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
	
	end if	
end if
check = check+1
linha_tabela=linha_tabela+1
max=nu_chamada
RS.MoveNext
Wend 
session("max")=max
%>
  <tr> 
    <td  colspan ="6"> 
      <%	  
if opt="cln" and trava="n" then
 %>
      <table width="100%" border="0" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"> <div align="center"> 
              <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?or=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Cancelar">
            </div></td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidades%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=grau%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=serie%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
              </font></div></td>
        </tr>
      </table>
      <%elseif trava="n" then%>
      <table width="100%" border="0" align="center" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"><div align="center"> 
            </div></td>
          <td width="50%"> <div align="center"> 
              <input type="submit" name="Submit2" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidade%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=curso%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=etapa%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
            </div></td>
        </tr>
      </table>
      <%end if%>
    </td>
  </tr>
</table>
<%case 2
tb="TB_Frequencia_Periodo"

 		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso&"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)

calc = calc*1


		

if errou="pt" or errou="pp" Then
	classe_peso = "tb_fundo_linha_erro"
else
	classe_peso = "tb_fundo_linha_peso"
end if

%> 
 <form action="bdb.asp?opt=cln" name="nota" method="post" onSubmit="return checksubmit()">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20" class="tb_subtit"> <div align="center">N&ordm;</div></td>
    <td width="780" class="tb_subtit"> 
      <div align="left">Nome</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B1</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B2</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B3</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B4</div></td>
  </tr>
  <%
check = 2
nu_chamada_ckq = 0
linha_tabela=1
While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")
	
	if nu_matricula = calc then
	  cor = "tb_fundo_linha_erro" 
	 onblur="mudar_cor_blur_erro"	  
	else
	 if check mod 2 =0 then
	  cor = "tb_fundo_linha_par" 
	  onblur="mudar_cor_blur_par"
	 else cor ="tb_fundo_linha_impar"
	 onblur="mudar_cor_blur_impar"
	  end if 
	end if


if (nu_chamada_ckq <>nu_chamada - 1) then
	teste_nu_chamada = nu_chamada-nu_chamada_ckq
	teste_nu_chamada=teste_nu_chamada-1
	
			for k=1 to teste_nu_chamada 
				nu_chamada_falta=nu_chamada_ckq+1
		
		%>
  <tr> 
    <td width="20" class="tb_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"></td>
    <td width="780" class="tb_fundo_linha_falta"><input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="falta" />
    &nbsp;</td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%  
				nu_chamada_ckq=nu_chamada_falta
			next	
		nu_chamada_ckq=nu_chamada	
			
				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
				Set RS2 = CON_A.Execute(SQL_A)
				
			NO_Aluno= RS2("NO_Aluno")
		
		
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
				Set RS3 = CON_N.Execute(SQL_N)
			
		if RS3.EOF then 
		 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780"  > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if
		%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
		end if
else
nu_chamada_ckq=nu_chamada

			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
			Set RS2 = CON_A.Execute(SQL_A)
			
		NO_Aluno= RS2("NO_Aluno")
	
			Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
			Set RS3 = CON_N.Execute(SQL_N)
		
		if RS3.EOF then 
	 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if		
						%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
	
	end if	
end if
check = check+1
linha_tabela=linha_tabela+1
max=nu_chamada
RS.MoveNext
Wend 
session("max")=max
%>
  <tr> 
    <td  colspan ="6"> 
      <%	  
if opt="cln" and trava="n" then
 %>
      <table width="100%" border="0" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	         <tr> 
          <td width="50%"> <div align="center"> 
              <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?or=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Cancelar">
            </div></td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidades%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=grau%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=serie%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
              </font></div></td>
        </tr>
      </table>
      <%elseif trava="n" then%>
      <table width="100%" border="0" align="center" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 
        <tr> 
          <td width="50%"><div align="center"> </div></td>
          <td width="50%"> <div align="center"> 
              <input type="submit" name="Submit2" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidade%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=curso%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=etapa%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
            </div></td>
        </tr>
      </table>
      <%end if%>
    </td>
  </tr>
</table>
<%case 3
tb="TB_Frequencia_Periodo"
 
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso&"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)

calc = calc*1


		

if errou="pt" or errou="pp" Then
	classe_peso = "tb_fundo_linha_erro"
else
	classe_peso = "tb_fundo_linha_peso"
end if

%> 
 <form action="bdc.asp?opt=cln" name="nota" method="post" onSubmit="return checksubmit()">
 <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20" class="tb_subtit"> <div align="center">N&ordm;</div></td>
    <td width="780" class="tb_subtit"> 
      <div align="left">Nome</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B1</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B2</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B3</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B4</div></td>
  </tr>
  <%
check = 2
nu_chamada_ckq = 0
linha_tabela=1
While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")
	
	if nu_matricula = calc then
	  cor = "tb_fundo_linha_erro" 
	 onblur="mudar_cor_blur_erro"	  
	else
	 if check mod 2 =0 then
	  cor = "tb_fundo_linha_par" 
	  onblur="mudar_cor_blur_par"
	 else cor ="tb_fundo_linha_impar"
	 onblur="mudar_cor_blur_impar"
	  end if 
	end if


if (nu_chamada_ckq <>nu_chamada - 1) then
	teste_nu_chamada = nu_chamada-nu_chamada_ckq
	teste_nu_chamada=teste_nu_chamada-1
	
			for k=1 to teste_nu_chamada 
				nu_chamada_falta=nu_chamada_ckq+1
		
		%>
  <tr> 
    <td width="20" class="tb_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"></td>
    <td width="780" class="tb_fundo_linha_falta"><input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="falta" />
    &nbsp;</td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%  
				nu_chamada_ckq=nu_chamada_falta
			next	
		nu_chamada_ckq=nu_chamada	
			
				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
				Set RS2 = CON_A.Execute(SQL_A)
				
			NO_Aluno= RS2("NO_Aluno")
		
		
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
				Set RS3 = CON_N.Execute(SQL_N)
			
		if RS3.EOF then 
		 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780"  > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if
		%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
		end if
else
nu_chamada_ckq=nu_chamada

			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
			Set RS2 = CON_A.Execute(SQL_A)
			
		NO_Aluno= RS2("NO_Aluno")
	
			Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
			Set RS3 = CON_N.Execute(SQL_N)
		
		if RS3.EOF then 
	 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
		end if				%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
	
	end if	
end if
check = check+1
max=nu_chamada
linha_tabela=linha_tabela+1
RS.MoveNext
Wend 
session("max")=max
%>
  <tr> 
    <td  colspan ="6"> 
      <%	  
if opt="cln" and trava="n" then
 %>
      <table width="100%" border="0" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 
        <tr> 
          <td width="50%"> <div align="center"> 
              <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?or=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Cancelar">
            </div></td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidades%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=grau%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=serie%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
              </font></div></td>
        </tr>
      </table>
      <%elseif trava="n" then%>
      <table width="100%" border="0" align="center" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 
        <tr> 
          <td width="50%"><div align="center"> </div></td>
          <td width="50%"> <div align="center"> 
              <input type="submit" name="Submit2" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidade%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=curso%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=etapa%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
            </div></td>
        </tr>
      </table>
      <%end if%>
    </td>
  </tr>
</table> 
<%




case 5
tb="TB_Frequencia_Periodo"

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso&"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)

calc = calc*1


		

if errou="pt" or errou="pp" Then
	classe_peso = "tb_fundo_linha_erro"
else
	classe_peso = "tb_fundo_linha_peso"
end if

%> 
 <form action="bde.asp?opt=cln" name="nota" method="post" onSubmit="return checksubmit()">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20" class="tb_subtit"> <div align="center">N&ordm;</div></td>
    <td width="780" class="tb_subtit"> 
      <div align="left">Nome</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B1</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B2</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B3</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B4</div></td>
  </tr>
  <%
check = 2
nu_chamada_ckq = 0
linha_tabela=1
While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")
	
	if nu_matricula = calc then
	  cor = "tb_fundo_linha_erro" 
	 onblur="mudar_cor_blur_erro"	  
	else
	 if check mod 2 =0 then
	  cor = "tb_fundo_linha_par" 
	  onblur="mudar_cor_blur_par"
	 else cor ="tb_fundo_linha_impar"
	 onblur="mudar_cor_blur_impar"
	  end if 
	end if


if (nu_chamada_ckq <>nu_chamada - 1) then
	teste_nu_chamada = nu_chamada-nu_chamada_ckq
	teste_nu_chamada=teste_nu_chamada-1
	
			for k=1 to teste_nu_chamada 
				nu_chamada_falta=nu_chamada_ckq+1
		
		%>
  <tr> 
    <td width="20" class="tb_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"></td>
    <td width="780" class="tb_fundo_linha_falta"><input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="falta" />
    &nbsp;</td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%  
				nu_chamada_ckq=nu_chamada_falta
			next	
		nu_chamada_ckq=nu_chamada	
			
				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
				Set RS2 = CON_A.Execute(SQL_A)
				
			NO_Aluno= RS2("NO_Aluno")
		
		
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
				Set RS3 = CON_N.Execute(SQL_N)
			
		if RS3.EOF then 
		 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780"  > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if
		%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
		end if
else
nu_chamada_ckq=nu_chamada

			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
			Set RS2 = CON_A.Execute(SQL_A)
			
		NO_Aluno= RS2("NO_Aluno")
	
			Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
			Set RS3 = CON_N.Execute(SQL_N)
		
		if RS3.EOF then 
	 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if		
						%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
	
	end if	
end if
check = check+1
max=nu_chamada
linha_tabela=linha_tabela+1
RS.MoveNext
Wend 
session("max")=max
%>
  <tr> 
    <td  colspan ="6"> 
      <%	  
if opt="cln" and trava="n" then
 %>
      <table width="100%" border="0" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"> <div align="center"> 
              <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?or=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Cancelar">
            </div></td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidades%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=grau%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=serie%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
              </font></div></td>
        </tr>
      </table>
      <%elseif trava="n" then%>
      <table width="100%" border="0" align="center" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"><div align="center"> 
            </div></td>
          <td width="50%"> <div align="center"> 
              <input type="submit" name="Submit2" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidade%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=curso%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=etapa%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
            </div></td>
        </tr>
      </table>
      <%end if%>
    </td>
  </tr>
</table>


<%case 6
tb="TB_Frequencia_Periodo"

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso&"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)

calc = calc*1


		

if errou="pt" or errou="pp" Then
	classe_peso = "tb_fundo_linha_erro"
else
	classe_peso = "tb_fundo_linha_peso"
end if

%> 
 <form action="bdf.asp?opt=cln" name="nota" method="post" onSubmit="return checksubmit()">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20" class="tb_subtit"> <div align="center">N&ordm;</div></td>
    <td width="780" class="tb_subtit"> 
      <div align="left">Nome</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B1</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B2</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B3</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B4</div></td>
  </tr>
  <%
check = 2
nu_chamada_ckq = 0
linha_tabela=1
While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")
	
	if nu_matricula = calc then
	  cor = "tb_fundo_linha_erro" 
	 onblur="mudar_cor_blur_erro"	  
	else
	 if check mod 2 =0 then
	  cor = "tb_fundo_linha_par" 
	  onblur="mudar_cor_blur_par"
	 else cor ="tb_fundo_linha_impar"
	 onblur="mudar_cor_blur_impar"
	  end if 
	end if


if (nu_chamada_ckq <>nu_chamada - 1) then
	teste_nu_chamada = nu_chamada-nu_chamada_ckq
	teste_nu_chamada=teste_nu_chamada-1
	
			for k=1 to teste_nu_chamada 
				nu_chamada_falta=nu_chamada_ckq+1
		
		%>
  <tr> 
    <td width="20" class="tb_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"></td>
    <td width="780" class="tb_fundo_linha_falta"><input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="falta" />
    &nbsp;</td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%  
				nu_chamada_ckq=nu_chamada_falta
			next	
		nu_chamada_ckq=nu_chamada	
			
				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
				Set RS2 = CON_A.Execute(SQL_A)
				
			NO_Aluno= RS2("NO_Aluno")
		
		
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
				Set RS3 = CON_N.Execute(SQL_N)
			
		if RS3.EOF then 
		 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780"  > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if
		%>

  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
		end if
else
nu_chamada_ckq=nu_chamada

			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
			Set RS2 = CON_A.Execute(SQL_A)
			
		NO_Aluno= RS2("NO_Aluno")
	
			Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
			Set RS3 = CON_N.Execute(SQL_N)
		
		if RS3.EOF then 
	 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50" >

<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if		
						%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
	
	end if	
end if
check = check+1
max=nu_chamada
linha_tabela=linha_tabela+1
RS.MoveNext
Wend 
session("max")=max
%>
  <tr> 
    <td  colspan ="6"> 
      <%	  
if opt="cln" and trava="n" then
 %>
      <table width="100%" border="0" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"> <div align="center"> 
              <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?or=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Cancelar">
            </div></td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidades%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=grau%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=serie%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
              </font></div></td>
        </tr>
      </table>
      <%elseif trava="n" then%>
      <table width="100%" border="0" align="center" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"><div align="center"> 
            </div></td>
          <td width="50%"> <div align="center"> 
              <input type="submit" name="Submit2" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidade%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=curso%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=etapa%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
            </div></td>
        </tr>
      </table>
      <%end if%>
    </td>
  </tr>
</table>
<%
case 7
tb="TB_Frequencia_Periodo"

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso&"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)

calc = calc*1


		

if errou="pt" or errou="pp" Then
	classe_peso = "tb_fundo_linha_erro"
else
	classe_peso = "tb_fundo_linha_peso"
end if

%> 
 <form action="bdk.asp?opt=cln" name="nota" method="post" onSubmit="return checksubmit()">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20" class="tb_subtit"> <div align="center">N&ordm;</div></td>
    <td width="780" class="tb_subtit"> 
      <div align="left">Nome</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B1</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B2</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B3</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B4</div></td>
  </tr>
  <%
check = 2
nu_chamada_ckq = 0
linha_tabela=1
While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")
	
	if nu_matricula = calc then
	  cor = "tb_fundo_linha_erro" 
	 onblur="mudar_cor_blur_erro"	  
	else
	 if check mod 2 =0 then
	  cor = "tb_fundo_linha_par" 
	  onblur="mudar_cor_blur_par"
	 else cor ="tb_fundo_linha_impar"
	 onblur="mudar_cor_blur_impar"
	  end if 
	end if


if (nu_chamada_ckq <>nu_chamada - 1) then
	teste_nu_chamada = nu_chamada-nu_chamada_ckq
	teste_nu_chamada=teste_nu_chamada-1
	
			for k=1 to teste_nu_chamada 
				nu_chamada_falta=nu_chamada_ckq+1
		
		%>
  <tr> 
    <td width="20" class="tb_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"></td>
    <td width="780" class="tb_fundo_linha_falta"><input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="falta" />
    &nbsp;</td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%  
				nu_chamada_ckq=nu_chamada_falta
			next	
		nu_chamada_ckq=nu_chamada	
			
				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
				Set RS2 = CON_A.Execute(SQL_A)
				
			NO_Aluno= RS2("NO_Aluno")
		
		
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
				Set RS3 = CON_N.Execute(SQL_N)
			
		if RS3.EOF then 
		 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780"  > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if
		%>

  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
		end if
else
nu_chamada_ckq=nu_chamada

			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
			Set RS2 = CON_A.Execute(SQL_A)
			
		NO_Aluno= RS2("NO_Aluno")
	
			Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
			Set RS3 = CON_N.Execute(SQL_N)
		
		if RS3.EOF then 
	 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50" >

<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if		
						%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
	
	end if	
end if
check = check+1
linha_tabela=linha_tabela+1
max=nu_chamada
RS.MoveNext
Wend 
session("max")=max
%>
  <tr> 
    <td  colspan ="6"> 
      <%	  
if opt="cln" and trava="n" then
 %>
      <table width="100%" border="0" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"> <div align="center"> 
              <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?or=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Cancelar">
            </div></td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidades%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=grau%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=serie%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
              </font></div></td>
        </tr>
      </table>
      <%elseif trava="n" then%>
      <table width="100%" border="0" align="center" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"><div align="center"> 
            </div></td>
          <td width="50%"> <div align="center"> 
              <input type="submit" name="Submit2" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidade%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=curso%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=etapa%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
            </div></td>
        </tr>
      </table>
      <%end if%>
    </td>
  </tr>
</table>





<%case 11
tb="TB_Frequencia_Periodo"

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso&"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)

calc = calc*1


		

if errou="pt" or errou="pp" Then
	classe_peso = "tb_fundo_linha_erro"
else
	classe_peso = "tb_fundo_linha_peso"
end if

%> 
 <form action="bda.asp" name="nota" method="post" onSubmit="return checksubmit()">
 
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20" class="tb_subtit"> <div align="center">N&ordm;</div></td>
    <td width="780" class="tb_subtit"> 
      <div align="left">Nome</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B1</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B2</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B3</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B4</div></td>
  </tr>
  <%
check = 2
nu_chamada_ckq = 0
linha_tabela=1
While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")
	
	if nu_matricula = calc then
	  cor = "tb_fundo_linha_erro" 
	 onblur="mudar_cor_blur_erro"	  
	else
	 if check mod 2 =0 then
	  cor = "tb_fundo_linha_par" 
	  onblur="mudar_cor_blur_par"
	 else cor ="tb_fundo_linha_impar"
	 onblur="mudar_cor_blur_impar"
	  end if 
	end if


if (nu_chamada_ckq <>nu_chamada - 1) then
	teste_nu_chamada = nu_chamada-nu_chamada_ckq
	teste_nu_chamada=teste_nu_chamada-1
	
			for k=1 to teste_nu_chamada 
				nu_chamada_falta=nu_chamada_ckq+1
		
		%>
  <tr> 
    <td width="20" class="tb_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"></td>
    <td width="780" class="tb_fundo_linha_falta"><input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="falta" />&nbsp;</td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%  
				nu_chamada_ckq=nu_chamada_falta
			next	
		nu_chamada_ckq=nu_chamada	
			
				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
				Set RS2 = CON_A.Execute(SQL_A)
				
			NO_Aluno= RS2("NO_Aluno")
		
		
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
				Set RS3 = CON_N.Execute(SQL_N)
			
		if RS3.EOF then 
		 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780"  > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if
		%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
		end if
else
nu_chamada_ckq=nu_chamada

			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
			Set RS2 = CON_A.Execute(SQL_A)
			
		NO_Aluno= RS2("NO_Aluno")
	
			Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
			Set RS3 = CON_N.Execute(SQL_N)
		
		if RS3.EOF then 
	 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if							%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
	
	end if	
end if
check = check+1
max=nu_chamada
linha_tabela=linha_tabela+1
RS.MoveNext
Wend 
session("max")=max
%>
  <tr> 
    <td  colspan ="6"> 
      <%	  
if opt="cln" then
 %>
      <table width="100%" border="0" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"> <div align="center"> 
              <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?or=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Cancelar">
            </div></td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidades%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=grau%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=serie%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
              </font></div></td>
        </tr>
      </table>
      <%else%>
      <table width="100%" border="0" align="center" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"><div align="center"> </div></td>
          <td width="50%"> <div align="center"> 
              <input type="submit" name="Submit2" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidade%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=curso%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=etapa%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
            </div></td>
        </tr>
      </table>
      <%end if%>
    </td>
  </tr>
</table>

<%case 21
tb="TB_Frequencia_Periodo"

 		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso&"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)

calc = calc*1


		

if errou="pt" or errou="pp" Then
	classe_peso = "tb_fundo_linha_erro"
else
	classe_peso = "tb_fundo_linha_peso"
end if

%> 
 <form action="bdb.asp" name="nota" method="post" onSubmit="return checksubmit()">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20" class="tb_subtit"> <div align="center">N&ordm;</div></td>
    <td width="780" class="tb_subtit"> 
      <div align="left">Nome</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B1</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B2</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B3</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B4</div></td>
  </tr>
  <%
check = 2
nu_chamada_ckq = 0
linha_tabela=1
While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")
	
	if nu_matricula = calc then
	  cor = "tb_fundo_linha_erro" 
	 onblur="mudar_cor_blur_erro"	  
	else
	 if check mod 2 =0 then
	  cor = "tb_fundo_linha_par" 
	  onblur="mudar_cor_blur_par"
	 else cor ="tb_fundo_linha_impar"
	 onblur="mudar_cor_blur_impar"
	  end if 
	end if


if (nu_chamada_ckq <>nu_chamada - 1) then
	teste_nu_chamada = nu_chamada-nu_chamada_ckq
	teste_nu_chamada=teste_nu_chamada-1
	
			for k=1 to teste_nu_chamada 
				nu_chamada_falta=nu_chamada_ckq+1
		
		%>
  <tr> 
    <td width="20" class="tb_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"></td>
    <td width="780" class="tb_fundo_linha_falta"><input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="falta" />&nbsp;</td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%  
				nu_chamada_ckq=nu_chamada_falta
			next	
		nu_chamada_ckq=nu_chamada	
			
				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
				Set RS2 = CON_A.Execute(SQL_A)
				
			NO_Aluno= RS2("NO_Aluno")
		
		
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
				Set RS3 = CON_N.Execute(SQL_N)
			
		if RS3.EOF then 
		 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780"  > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if
		%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
		end if
else
nu_chamada_ckq=nu_chamada

			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
			Set RS2 = CON_A.Execute(SQL_A)
			
		NO_Aluno= RS2("NO_Aluno")
	
			Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
			Set RS3 = CON_N.Execute(SQL_N)
		
		if RS3.EOF then 
	 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if							%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
	
	end if	
end if
check = check+1
max=nu_chamada
linha_tabela=linha_tabela+1
RS.MoveNext
Wend 
session("max")=max
%>
  <tr> 
    <td  colspan ="6"> 
      <%	  
if opt="cln" then
 %>
      <table width="100%" border="0" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"> <div align="center"> 
              <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?or=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Cancelar">
            </div></td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidades%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=grau%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=serie%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
              </font></div></td>
        </tr>
      </table>
      <%else%>
      <table width="100%" border="0" align="center" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"><div align="center"> </div></td>
          <td width="50%"> <div align="center"> 
              <input type="submit" name="Submit2" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidade%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=curso%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=etapa%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
            </div></td>
        </tr>
      </table>
      <%end if%>
    </td>
  </tr>
</table>
<%case 31
tb="TB_Frequencia_Periodo"
 
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso&"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)

calc = calc*1


		

if errou="pt" or errou="pp" Then
	classe_peso = "tb_fundo_linha_erro"
else
	classe_peso = "tb_fundo_linha_peso"
end if

%> 
 <form action="bdc.asp" name="nota" method="post" onSubmit="return checksubmit()">
 <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20" class="tb_subtit"> <div align="center">N&ordm;</div></td>
    <td width="780" class="tb_subtit"> 
      <div align="left">Nome</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B1</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B2</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B3</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B4</div></td>
  </tr>
  <%
check = 2
nu_chamada_ckq = 0
linha_tabela=1
While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")
	
	if nu_matricula = calc then
	  cor = "tb_fundo_linha_erro" 
	 onblur="mudar_cor_blur_erro"	  
	else
	 if check mod 2 =0 then
	  cor = "tb_fundo_linha_par" 
	  onblur="mudar_cor_blur_par"
	 else cor ="tb_fundo_linha_impar"
	 onblur="mudar_cor_blur_impar"
	  end if 
	end if


if (nu_chamada_ckq <>nu_chamada - 1) then
	teste_nu_chamada = nu_chamada-nu_chamada_ckq
	teste_nu_chamada=teste_nu_chamada-1
	
			for k=1 to teste_nu_chamada 
				nu_chamada_falta=nu_chamada_ckq+1
		
		%>
  <tr> 
    <td width="20" class="tb_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"></td>
    <td width="780" class="tb_fundo_linha_falta"><input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="falta" />&nbsp;</td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%  
				nu_chamada_ckq=nu_chamada_falta
			next	
		nu_chamada_ckq=nu_chamada	
			
				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
				Set RS2 = CON_A.Execute(SQL_A)
				
			NO_Aluno= RS2("NO_Aluno")
		
		
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
				Set RS3 = CON_N.Execute(SQL_N)
			
		if RS3.EOF then 
		 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780"  > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if
		%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
		end if
else
nu_chamada_ckq=nu_chamada

			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
			Set RS2 = CON_A.Execute(SQL_A)
			
		NO_Aluno= RS2("NO_Aluno")
	
			Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
			Set RS3 = CON_N.Execute(SQL_N)
		
		if RS3.EOF then 
	 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if						
						%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
	
	end if	
end if
check = check+1
max=nu_chamada
linha_tabela=linha_tabela+1
RS.MoveNext
Wend 
session("max")=max
%>
  <tr> 
    <td  colspan ="6"> 
      <%	  
if opt="cln" then
 %>
      <table width="100%" border="0" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"> <div align="center"> 
              <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?or=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Cancelar">
            </div></td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidades%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=grau%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=serie%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
              </font></div></td>
        </tr>
      </table>
      <%else%>
      <table width="100%" border="0" align="center" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"><div align="center"> </div></td>
          <td width="50%"> <div align="center"> 
              <input type="submit" name="Submit2" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidade%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=curso%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=etapa%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
             
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
            </div></td>
        </tr>
      </table>
      <%end if%>
    </td>
  </tr>
</table>








<%case 51
tb="TB_Frequencia_Periodo"

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso&"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)

calc = calc*1


		

if errou="pt" or errou="pp" Then
	classe_peso = "tb_fundo_linha_erro"
else
	classe_peso = "tb_fundo_linha_peso"
end if

%> 
 <form action="bde.asp" name="nota" method="post" onSubmit="return checksubmit()">
 
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20" class="tb_subtit"> <div align="center">N&ordm;</div></td>
    <td width="780" class="tb_subtit"> 
      <div align="left">Nome</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B1</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B2</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B3</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B4</div></td>
  </tr>
  <%
check = 2
nu_chamada_ckq = 0
linha_tabela=1
While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")
	
	if nu_matricula = calc then
	  cor = "tb_fundo_linha_erro" 
	 onblur="mudar_cor_blur_erro"	  
	else
	 if check mod 2 =0 then
	  cor = "tb_fundo_linha_par" 
	  onblur="mudar_cor_blur_par"
	 else cor ="tb_fundo_linha_impar"
	 onblur="mudar_cor_blur_impar"
	  end if 
	end if


if (nu_chamada_ckq <>nu_chamada - 1) then
	teste_nu_chamada = nu_chamada-nu_chamada_ckq
	teste_nu_chamada=teste_nu_chamada-1
	
			for k=1 to teste_nu_chamada 
				nu_chamada_falta=nu_chamada_ckq+1
		
		%>
  <tr> 
    <td width="20" class="tb_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"></td>
    <td width="780" class="tb_fundo_linha_falta"><input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="falta" />&nbsp;</td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%  
				nu_chamada_ckq=nu_chamada_falta
			next	
		nu_chamada_ckq=nu_chamada	
			
				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
				Set RS2 = CON_A.Execute(SQL_A)
				
			NO_Aluno= RS2("NO_Aluno")
		
		
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
				Set RS3 = CON_N.Execute(SQL_N)
			
		if RS3.EOF then 
		 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780"  > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if
		%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
		end if
else
nu_chamada_ckq=nu_chamada

			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
			Set RS2 = CON_A.Execute(SQL_A)
			
		NO_Aluno= RS2("NO_Aluno")
	
			Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
			Set RS3 = CON_N.Execute(SQL_N)
		
		if RS3.EOF then 
	 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if							%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
	
	end if	
end if
check = check+1
max=nu_chamada
linha_tabela=linha_tabela+1
RS.MoveNext
Wend 
session("max")=max
%>
  <tr> 
    <td  colspan ="6"> 
      <%	  
if opt="cln" then
 %>
      <table width="100%" border="0" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"> <div align="center"> 
              <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?or=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Cancelar">
            </div></td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidades%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=grau%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=serie%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
              </font></div></td>
        </tr>
      </table>
      <%else%>
      <table width="100%" border="0" align="center" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"><div align="center"> </div></td>
          <td width="50%"> <div align="center"> 
              <input type="submit" name="Submit2" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidade%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=curso%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=etapa%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
            </div></td>
        </tr>
      </table>
      <%end if%>
    </td>
  </tr>
</table>  
<%case 61
tb="TB_Frequencia_Periodo"

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso&"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)

calc = calc*1


		

if errou="pt" or errou="pp" Then
	classe_peso = "tb_fundo_linha_erro"
else
	classe_peso = "tb_fundo_linha_peso"
end if

%> 
 <form action="bdf.asp" name="nota" method="post" onSubmit="return checksubmit()">
 
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20" class="tb_subtit"> <div align="center">N&ordm;</div></td>
    <td width="780" class="tb_subtit"> 
      <div align="left">Nome</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B1</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B2</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B3</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B4</div></td>
  </tr>
  <%
check = 2
nu_chamada_ckq = 0
linha_tabela=1
While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")
	
	if nu_matricula = calc then
	  cor = "tb_fundo_linha_erro" 
	 onblur="mudar_cor_blur_erro"	  
	else
	 if check mod 2 =0 then
	  cor = "tb_fundo_linha_par" 
	  onblur="mudar_cor_blur_par"
	 else cor ="tb_fundo_linha_impar"
	 onblur="mudar_cor_blur_impar"
	  end if 
	end if


if (nu_chamada_ckq <>nu_chamada - 1) then
	teste_nu_chamada = nu_chamada-nu_chamada_ckq
	teste_nu_chamada=teste_nu_chamada-1
	
			for k=1 to teste_nu_chamada 
				nu_chamada_falta=nu_chamada_ckq+1
		
		%>
  <tr> 
    <td width="20" class="tb_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"></td>
    <td width="780" class="tb_fundo_linha_falta"><input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="falta" />&nbsp;</td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%  
				nu_chamada_ckq=nu_chamada_falta
			next	
		nu_chamada_ckq=nu_chamada	
			
				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
				Set RS2 = CON_A.Execute(SQL_A)
				
			NO_Aluno= RS2("NO_Aluno")
		
		
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
				Set RS3 = CON_N.Execute(SQL_N)
			
		if RS3.EOF then 
		 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780"  > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if
		%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
		end if
else
nu_chamada_ckq=nu_chamada

			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
			Set RS2 = CON_A.Execute(SQL_A)
			
		NO_Aluno= RS2("NO_Aluno")
	
			Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
			Set RS3 = CON_N.Execute(SQL_N)
		
		if RS3.EOF then 
	 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if							%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
	
	end if	
end if
check = check+1
max=nu_chamada
linha_tabela=linha_tabela+1
RS.MoveNext
Wend 
session("max")=max
%>
  <tr> 
    <td  colspan ="6"> 
      <%	  
if opt="cln" then
 %>
      <table width="100%" border="0" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"> <div align="center"> 
              <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?or=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Cancelar">
            </div></td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidades%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=grau%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=serie%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
              </font></div></td>
        </tr>
      </table>
      <%else%>
      <table width="100%" border="0" align="center" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"><div align="center"> </div></td>
          <td width="50%"> <div align="center"> 
              <input type="submit" name="Submit2" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidade%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=curso%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=etapa%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
            </div></td>
        </tr>
      </table>
      <%end if%>
    </td>
  </tr>
</table> 

<%
case 71
tb="TB_Frequencia_Periodo"

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso&"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)

calc = calc*1


		

if errou="pt" or errou="pp" Then
	classe_peso = "tb_fundo_linha_erro"
else
	classe_peso = "tb_fundo_linha_peso"
end if

%> 
 <form action="bdk.asp" name="nota" method="post" onSubmit="return checksubmit()">
 
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20" class="tb_subtit"> <div align="center">N&ordm;</div></td>
    <td width="780" class="tb_subtit"> 
      <div align="left">Nome</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B1</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B2</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B3</div></td>
    <td width="50" class="tb_subtit"> 
      <div align="center">B4</div></td>
  </tr>
  <%
check = 2
nu_chamada_ckq = 0
linha_tabela=1
While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")
	
	if nu_matricula = calc then
	  cor = "tb_fundo_linha_erro" 
	 onblur="mudar_cor_blur_erro"	  
	else
	 if check mod 2 =0 then
	  cor = "tb_fundo_linha_par" 
	  onblur="mudar_cor_blur_par"
	 else cor ="tb_fundo_linha_impar"
	 onblur="mudar_cor_blur_impar"
	  end if 
	end if


if (nu_chamada_ckq <>nu_chamada - 1) then
	teste_nu_chamada = nu_chamada-nu_chamada_ckq
	teste_nu_chamada=teste_nu_chamada-1
	
			for k=1 to teste_nu_chamada 
				nu_chamada_falta=nu_chamada_ckq+1
		
		%>
  <tr> 
    <td width="20" class="tb_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"></td>
    <td width="780" class="tb_fundo_linha_falta"><input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="falta" />&nbsp;</td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%  
				nu_chamada_ckq=nu_chamada_falta
			next	
		nu_chamada_ckq=nu_chamada	
			
				Set RS2 = Server.CreateObject("ADODB.Recordset")
				SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
				Set RS2 = CON_A.Execute(SQL_A)
				
			NO_Aluno= RS2("NO_Aluno")
		
		
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
				Set RS3 = CON_N.Execute(SQL_N)
			
		if RS3.EOF then 
		 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780"  > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if
		%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
		end if
else
nu_chamada_ckq=nu_chamada

			Set RS2 = Server.CreateObject("ADODB.Recordset")
			SQL_A = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
			Set RS2 = CON_A.Execute(SQL_A)
			
		NO_Aluno= RS2("NO_Aluno")
	
			Set RS3 = Server.CreateObject("ADODB.Recordset")
				'SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & "AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula
			Set RS3 = CON_N.Execute(SQL_N)
		
		if RS3.EOF then 
	 %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>">
      </div></td>
  </tr>
  <%else
		va_f1=RS3("NU_Faltas_P1")
		va_f2=RS3("NU_Faltas_P2")
		va_f3=RS3("NU_Faltas_P3")
		va_f4=RS3("NU_Faltas_P4")
		
			if opt="err6" and nu_matricula = calc then
					select case errou
					case "f1"
							va_f1 = qerrou
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")
							va_f5=Session("va_f5")
							va_f6=Session("va_f6")							
					case "f2"
							va_f1=Session("va_f1")
							va_f2 = qerrou
							va_f3=Session("va_f3")
							va_f4=Session("va_f4")											
					case "f3"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3 = qerrou
							va_f4=Session("va_f4")
					case "f4"
							va_f1=Session("va_f1")
							va_f2=Session("va_f2")
							va_f3=Session("va_f3")
							va_f4 = qerrou				
					end select
			end if							%>
  <tr class="<%=cor%>" id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" > 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" > 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c1")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c2")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c3")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write(linha_tabela&"c4")%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
	
	end if	
end if
check = check+1
max=nu_chamada
linha_tabela=linha_tabela+1
RS.MoveNext
Wend 
session("max")=max
%>
  <tr> 
    <td  colspan ="6"> 
      <%	  
if opt="cln" then
 %>
      <table width="100%" border="0" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"> <div align="center"> 
              <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?or=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Cancelar">
            </div></td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidades%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=grau%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=serie%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
              </font></div></td>
        </tr>
      </table>
      <%else%>
      <table width="100%" border="0" align="center" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%"><div align="center"> </div></td>
          <td width="50%"> <div align="center"> 
              <input type="submit" name="Submit2" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=unidade%>">
              <input name="curso" type="hidden" id="curso" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=curso%>">
              <input name="etapa" type="hidden" id="etapa" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=etapa%>">
              <input name="turma" type="hidden" id="turma" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" onkeydown="keyPressed(this.id,event,<% response.Write(notas_a_lancar)%>,<% response.Write(qtd_alunos)%>)" value="<% = ano_letivo%>">
            </div></td>
        </tr>
      </table>
      <%end if%>
    </td>
  </tr>
</table> 

<%
end select
%>
</FORM>
<%
End Function
Function av_quali (CAMINHOa,CAMINHOn,unidade,curso,etapa,turma,co_materia,periodo,ano_letivo,co_usr,opcao,erro)

		Set CON_N = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIR3
		
		Set CON_N = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIR3		
		
		Set CON_A = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_A.Open ABRIR
		
		Set CON_G = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_G.Open ABRIR			
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0	
		
		Set RS8 = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT * FROM TB_Materia where CO_Materia='"& co_materia &"'"
		RS8.Open SQL8, CON0

		no_mat= RS8("NO_Materia")	
		co_materia_pr= RS8("CO_Materia_Principal")
		
if Isnull(co_materia_pr) then
	co_materia_pr= co_materia
end if		
		
		Set RS9 = Server.CreateObject("ADODB.Recordset")
		SQL9 = "SELECT * FROM TB_Da_Aula where NU_Unidade = "&unidade&" AND CO_Curso = '"&curso&"' AND CO_Etapa = '"&etapa&"' AND CO_Turma = '"&turma&"' AND CO_Materia_Principal='"& co_materia &"' AND CO_Materia='"& co_materia_pr &"'"
'		response.Write(SQL9)
		RS9.Open SQL9, CON_G			

co_prof = RS9("CO_Professor")
opcao_exibicao = opcao
dados_periodo =  periodos(periodo, "num")
total_periodo = split(dados_periodo,"#!#") 
notas_a_lancar = ubound(total_periodo)-2

qtd_alunos=contalunos(CAMINHOa,ano_letivo,unidade,curso,etapa,turma,"TODOS")

trava=session("trava")

%>
 <form action="../../../../inc/bdw.asp" name="av_quali" method="post" onSubmit="return checksubmit()">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" class="tb_subtit">N&ordm;</td>
    <td align="left" class="tb_subtit">Nome</td>
    <%for i=0 to notas_a_lancar
		sigla_periodo =  periodos(total_periodo(i), "sigla")
	%>
    <td align="center" class="tb_subtit"><%response.Write(sigla_periodo)%></td>
    <%next%>
  </tr>
    <tr class="form_dado_texto">
    <td>&nbsp;</td>
    <td><input name="unidade" type="hidden" id="unidade" value="<%response.Write(unidade)%>">
              <input name="curso" type="hidden" id="curso"  value="<%response.Write(curso)%>">
              <input name="etapa" type="hidden" id="etapa"  value="<%response.Write(etapa)%>">
              <input name="turma" type="hidden" id="turma" value="<%response.Write(turma)%>">   
              <input name="co_mat_prin" type="hidden" id="turma" value="<%response.Write(co_materia_pr)%>">  
              <input name="co_mat" type="hidden" id="turma" value="<%response.Write(co_materia)%>">                
              <input name="max" type="hidden" id="max" value="<% response.Write(qtd_alunos)%>">
              <input name="co_usr" type="hidden" id="co_usr" value="<% response.Write(co_usr)%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" value="<%response.Write(session("ano_letivo"))%>"></td>
    <%for l=0 to notas_a_lancar
	%>    
    <td><table width="160" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="25" align="center" class="form_dado_texto">I</td>
            <td width="25" align="center" class="form_dado_texto">R</td>
            <td width="25" align="center" class="form_dado_texto">B</td>
            <td width="25" align="center" class="form_dado_texto">MB</td>
            <td width="60" align="center" valign="middle">&nbsp;</td>
      </tr></table></td>
      <%next%>
  </tr>
  <%
  vetor_matrics = alunos_turma(session("ano_letivo"),unidade,curso,etapa,turma,"num")
  
  vetor_alunos = split(vetor_matrics,"#$#") 

check = 2
nu_chamada_ckq = 0
linha_tabela=1
	

  for n = 0 to ubound(vetor_alunos) 
    dados_alunos = split(vetor_alunos(n),"#!#") 
  
	 if check mod 2 =0 then
	    cor = "tb_fundo_linha_par" 
	    onblur="mudar_cor_blur_par"
	 else 
	 	cor ="tb_fundo_linha_impar"
	 	onblur="mudar_cor_blur_impar"		
	 end if   
	 if dados_alunos(3) <> "C" then
		opcao = "disabled='disabled'"
		nome_aluno = dados_alunos(2)
		on_click="N"		
	 else
		select case opcao_exibicao
		case "C"
			opcao = "disabled='disabled'"
			on_click="N"				
		case "E"	
			opcao = ""
			on_click="S"
		end select	
		nome_aluno = dados_alunos(2)
	 end if
	 nu_chamada_ckq = nu_chamada_ckq*1
	 dados_alunos(1) = dados_alunos(1)*1

	 nu_chamada_anterior = dados_alunos(1)  - 1
	 num_chamada_ckq_seguinte = nu_chamada_ckq+1
'	 if num_chamada_ckq_seguinte<>dados_alunos(1) then	 	 
'	 	for f=num_chamada_ckq_seguinte to nu_chamada_anterior
%>
<!--  <tr class="tb_fundo_linha_falta" id="<%'response.Write("celula"&num_chamada_ckq_seguinte)%>">
    <td><%'response.Write(num_chamada_ckq_seguinte)%></td>
    <td>&nbsp;</td>
    <%'for j=0 to notas_a_lancar
	
	%><td>&nbsp;
    </td>
    <%'next%>
  </tr>
-->		
<%		
'		next
'	end if	
  %>
  <tr class="<%=cor%>" id="<%response.Write("celula"&dados_alunos(1))%>">
    <td><%response.Write(dados_alunos(1))%></td>
    <td><%response.Write(nome_aluno)%></td>
    <%for j=0 to notas_a_lancar
	
		identificacao = "av_n"&dados_alunos(1)&"_p"&total_periodo(j)
		if on_click="S" then
			onclick="limpa_option_button('"&identificacao&"')"
			limpa_bt="<a href=""#"" onclick="""&onclick&"""><img src=""../../../../img/botao-limpar.gif"" width=""40"" height=""16"" alt=""Limpar avalia&ccedil;&atilde;o"" border = ""0"" /></a>"			
		else
			onclick=""
			limpa_bt="<img src='../../../../img/botao-limpar.gif' width='40' height='16' alt='Limpar avalia&ccedil;&atilde;o' border = '0' />"					
		end if	
	%>
    	<td>
        <table width="160" border="0" cellspacing="0" cellpadding="0">
<!--          <tr>
            <td align="center" class="form_dado_texto">I</td>
            <td align="center" class="form_dado_texto">R</td>
            <td align="center" class="form_dado_texto">B</td>
            <td align="center" class="form_dado_texto">MB</td>
            <td width="60" align="center" valign="middle">&nbsp;</td>
          </tr>-->
          <tr>
       <%
	   
	   	if j=0 then
			wrk_bd_nota_per = "VA_Ava1"
		elseif j=1 then
			wrk_bd_nota_per = "VA_Ava2"
		elseif j=2 then 			
			wrk_bd_nota_per = "VA_Ava3"
		elseif j=3 then 			 	
			wrk_bd_nota_per = "VA_Ava4"
		end if	
	   
		Set RSC = Server.CreateObject("ADODB.Recordset")
		SQLC = "Select * from TB_Nota_W WHERE CO_Matricula = "& dados_alunos(0) &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"'"
		Set RSC = CON_N.Execute(SQLC)	
		
		IF RSC.eof THEN  
			wrk_checked_I="" 
			wrk_checked_R="" 
			wrk_checked_B="" 
			wrk_checked_MB="" 
        else	
			wrk_checked_I="" 
			wrk_checked_R="" 
			wrk_checked_B="" 
			wrk_checked_MB="" 		
			if RSC(wrk_bd_nota_per) = "I" then					
				wrk_checked_I="checked"
			elseif RSC(wrk_bd_nota_per) = "R" then	
				wrk_checked_R="checked" 
			elseif RSC(wrk_bd_nota_per) = "B" then	
				wrk_checked_B="checked" 	
			elseif RSC(wrk_bd_nota_per) = "M" then	
				wrk_checked_MB="checked" 													 
			end if										
	   end if
	   %>
        
            <td width="25" align="center" ><input name="<%response.Write(identificacao)%>" id="<%response.Write(identificacao)%>" type="radio"  value="I" <%response.Write(opcao)%> <%response.Write(wrk_checked_I)%> onFocus="mudar_cor_focus(celula<%response.Write(dados_alunos(1))%>);" onBlur="<%response.Write(onblur)%>(celula<%response.Write(dados_alunos(1))%>)" /></td>
            <td width="25" align="center"><input name="<%response.Write(identificacao)%>" id="<%response.Write(identificacao)%>" type="radio"  value="R" <%response.Write(opcao)%> <%response.Write(wrk_checked_R)%> onFocus="mudar_cor_focus(celula<%response.Write(dados_alunos(1))%>);" onBlur="<%response.Write(onblur)%>(celula<%response.Write(dados_alunos(1))%>)" /></td>
            <td width="25" align="center"><input name="<%response.Write(identificacao)%>" id="<%response.Write(identificacao)%>" type="radio"  value="B" <%response.Write(opcao)%> <%response.Write(wrk_checked_B)%> onFocus="mudar_cor_focus(celula<%response.Write(dados_alunos(1))%>);" onBlur="<%response.Write(onblur)%>(celula<%response.Write(dados_alunos(1))%>)" /></td>
            <td width="25" align="center"><input name="<%response.Write(identificacao)%>" id="<%response.Write(identificacao)%>" type="radio"  value="M" <%response.Write(opcao)%> <%response.Write(wrk_checked_MB)%>  onFocus="mudar_cor_focus(celula<%response.Write(dados_alunos(1))%>);" onBlur="<%response.Write(onblur)%>(celula<%response.Write(dados_alunos(1))%>)" /></td>
            <td width="60" align="center" valign="middle"><%response.Write(limpa_bt)%></td>
          </tr>
        </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr></tr>
        </table>
    </td>
    <%next%>
  </tr>
  <%
  check=check+1
  next
  
  colspan = notas_a_lancar+3
  %>
<tr>
    <td colspan="<%response.Write(colspan)%>"><table width="100%" border="0" cellspacing="0">
	  <tr>
		<td colspan="2">
		  <hr>
		</td>
	  </tr>	 	  
        <tr> 
          <td width="50%" align="center"><input name="co_prof" type="hidden" id="turma" value="<%response.Write(co_prof)%>">   
<% if erro>0 then
	url = "index.asp?nvg=WN-LN-LN-LAQ"
else
	url = "altera.asp"
end if%>          
             <input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','<%response.Write(url)%>');return document.MM_returnValue" value="Voltar">
            </td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit2" value="Salvar" class="botao_prosseguir">
          </font></div></td>
        </tr>
      </table></td>
    </tr>  
</table>
</form>
<%
		
End Function		

%>