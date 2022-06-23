<!--#include file="caminhos.asp"-->
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
    <td width="20" height="40" class="tb_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" value="falta"> 
    </td>
    <td width="780" class="tb_fundo_linha_falta">&nbsp;</td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> <%response.Write(va_f1)%>
        <input name="<%response.Write("f1_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f1_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> <%response.Write(va_f2)%>
        <input name="<%response.Write("f2_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f2_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> <%response.Write(va_f3)%>
        <input name="<%response.Write("f3_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f3_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"><%response.Write(va_f4)%> 
        <input name="<%response.Write("f4_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f4_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota">
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
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"><%response.Write(va_f1)%> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>">
      </div></td>
    <td width="50"  >
<div align="center"><%response.Write(va_f2)%> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>">
      </div></td>
    <td width="50"  >
<div align="center"> <%response.Write(va_f3)%>
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>">
      </div></td>
    <td width="50"  >
<div align="center"><%response.Write(va_f4)%> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>">
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
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> <%response.Write(va_f1)%>
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="hidden" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> <%response.Write(va_f2)%>
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="hidden" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"><%response.Write(va_f3)%> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="hidden" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"><%response.Write(va_f4)%> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="hidden" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota">
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
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> <%response.Write(va_f1)%>
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>">
      </div></td>
    <td width="50" >
<div align="center"> <%response.Write(va_f2)%>
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>">
      </div></td>
    <td width="50" >
<div align="center"><%response.Write(va_f3)%> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>">
      </div></td>
    <td width="50" >
<div align="center"><%response.Write(va_f4)%>
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>">
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
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"><%response.Write(va_f1)%> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="hidden" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> <%response.Write(va_f2)%>

        <input name="<%response.Write("f2_"&nu_chamada)%>" type="hidden" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> <%response.Write(va_f3)%>
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="hidden" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"><%response.Write(va_f4)%> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="hidden" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
	
	end if	
end if
check = check+1
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
              <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?or=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" value="Cancelar">
            </div></td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" value="<%=unidades%>">
              <input name="curso" type="hidden" id="curso" value="<%=grau%>">
              <input name="etapa" type="hidden" id="etapa" value="<%=serie%>">
              <input name="turma" type="hidden" id="turma" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>">
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
              <input type="submit" name="Submit2" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" value="<%=unidade%>">
              <input name="curso" type="hidden" id="curso" value="<%=curso%>">
              <input name="etapa" type="hidden" id="etapa" value="<%=etapa%>">
              <input name="turma" type="hidden" id="turma" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>">
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
    <td width="20" height="40" class="tb_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" value="falta"> 
    </td>
    <td width="780" class="tb_fundo_linha_falta">&nbsp;</td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> <%response.Write(va_f1)%>
        <input name="<%response.Write("f1_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f1_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> <%response.Write(va_f2)%>
        <input name="<%response.Write("f2_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f2_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> <%response.Write(va_f3)%>
        <input name="<%response.Write("f3_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f3_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"><%response.Write(va_f4)%> 
        <input name="<%response.Write("f4_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f4_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota">
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
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"><%response.Write(va_f1)%> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>">
      </div></td>
    <td width="50"  >
<div align="center"><%response.Write(va_f2)%> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>">
      </div></td>
    <td width="50"  >
<div align="center"><%response.Write(va_f3)%> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>">
      </div></td>
    <td width="50"  >
<div align="center"><%response.Write(va_f4)%> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>">
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
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> <%response.Write(va_f1)%>
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="hidden" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"><%response.Write(va_f2)%>
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="hidden" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"><%response.Write(va_f3)%> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="hidden" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"><%response.Write(va_f4)%> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="hidden" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota">
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
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> <%response.Write(va_f1)%>
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>">
      </div></td>
    <td width="50" >
<div align="center"> <%response.Write(va_f2)%>
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>">
      </div></td>
    <td width="50" >
<div align="center"> <%response.Write(va_f3)%>
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>">
      </div></td>
    <td width="50" >
<div align="center"><%response.Write(va_f4)%> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>">
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
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"><%response.Write(va_f1)%> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="hidden" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"><%response.Write(va_f2)%> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="hidden" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"><%response.Write(va_f3)%> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="hidden" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"><%response.Write(va_f4)%> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="hidden" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
	
	end if	
end if
check = check+1
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
              <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?or=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" value="Cancelar">
            </div></td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" value="<%=unidades%>">
              <input name="curso" type="hidden" id="curso" value="<%=grau%>">
              <input name="etapa" type="hidden" id="etapa" value="<%=serie%>">
              <input name="turma" type="hidden" id="turma" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>">
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
              <input type="submit" name="Submit2" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" value="<%=unidade%>">
              <input name="curso" type="hidden" id="curso" value="<%=curso%>">
              <input name="etapa" type="hidden" id="etapa" value="<%=etapa%>">
              <input name="turma" type="hidden" id="turma" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>">
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
    <td width="20" height="40" class="tb_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" value="falta"> 
    </td>
    <td width="780" class="tb_fundo_linha_falta">&nbsp;</td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"><%response.Write(va_f1)%> 
        <input name="<%response.Write("f1_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f1_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"><%response.Write(va_f2)%> 
        <input name="<%response.Write("f2_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f2_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> <%response.Write(va_f3)%>
        <input name="<%response.Write("f3_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f3_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"><%response.Write(va_f4)%> 
        <input name="<%response.Write("f4_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f4_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota">
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
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"><%response.Write(va_f1)%> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>">
      </div></td>
    <td width="50"  >
<div align="center"><%response.Write(va_f2)%> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>">
      </div></td>
    <td width="50"  >
<div align="center"><%response.Write(va_f3)%> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>">
      </div></td>
    <td width="50"  >
<div align="center"><%response.Write(va_f4)%> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>">
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
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"><%response.Write(va_f1)%> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="hidden" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"><%response.Write(va_f2)%> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="hidden" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"><%response.Write(va_f3)%> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="hidden" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"><%response.Write(va_f4)%> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="hidden" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota">
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
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> <%response.Write(va_f1)%>
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>">
      </div></td>
    <td width="50" >
<div align="center"> <%response.Write(va_f2)%>
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>">
      </div></td>
    <td width="50" >
<div align="center"> <%response.Write(va_f3)%>
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>">
      </div></td>
    <td width="50" >
<div align="center"><%response.Write(va_f4)%> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="hidden" class="nota" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>">
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
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> <%response.Write(va_f1)%>
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="hidden" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> <%response.Write(va_f2)%>
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="hidden" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> <%response.Write(va_f3)%>
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="hidden" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> <%response.Write(va_f4)%>
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="hidden" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
	
	end if	
end if
check = check+1
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
              <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?or=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" value="Cancelar">
            </div></td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" value="<%=unidades%>">
              <input name="curso" type="hidden" id="curso" value="<%=grau%>">
              <input name="etapa" type="hidden" id="etapa" value="<%=serie%>">
              <input name="turma" type="hidden" id="turma" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>">
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
              <input type="submit" name="Submit2" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" value="<%=unidade%>">
              <input name="curso" type="hidden" id="curso" value="<%=curso%>">
              <input name="etapa" type="hidden" id="etapa" value="<%=etapa%>">
              <input name="turma" type="hidden" id="turma" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>">
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
    <td width="20" height="40" class="tb_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" value="falta"> 
    </td>
    <td width="780" class="tb_fundo_linha_falta">&nbsp;</td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f1_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f2_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f3_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f4_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota">
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
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>">
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
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota">
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
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>">
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
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
	
	end if	
end if
check = check+1
max=nu_chamada
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
              <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?or=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" value="Cancelar">
            </div></td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" value="<%=unidades%>">
              <input name="curso" type="hidden" id="curso" value="<%=grau%>">
              <input name="etapa" type="hidden" id="etapa" value="<%=serie%>">
              <input name="turma" type="hidden" id="turma" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>">
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
              <input type="submit" name="Submit2" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" value="<%=unidade%>">
              <input name="curso" type="hidden" id="curso" value="<%=curso%>">
              <input name="etapa" type="hidden" id="etapa" value="<%=etapa%>">
              <input name="turma" type="hidden" id="turma" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>">
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
    <td width="20" height="40" class="tb_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" value="falta"> 
    </td>
    <td width="780" class="tb_fundo_linha_falta">&nbsp;</td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f1_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f2_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f3_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f4_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota">
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
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>">
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
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota">
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
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>">
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
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
	
	end if	
end if
check = check+1
max=nu_chamada
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
              <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?or=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" value="Cancelar">
            </div></td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" value="<%=unidades%>">
              <input name="curso" type="hidden" id="curso" value="<%=grau%>">
              <input name="etapa" type="hidden" id="etapa" value="<%=serie%>">
              <input name="turma" type="hidden" id="turma" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>">
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
              <input type="submit" name="Submit2" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" value="<%=unidade%>">
              <input name="curso" type="hidden" id="curso" value="<%=curso%>">
              <input name="etapa" type="hidden" id="etapa" value="<%=etapa%>">
              <input name="turma" type="hidden" id="turma" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>">
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
    <td width="20" height="40" class="tb_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" value="falta"> 
    </td>
    <td width="780" class="tb_fundo_linha_falta">&nbsp;</td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f1_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f2_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f3_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" class="tb_fundo_linha_falta">
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada_falta)%>" type="hidden" id="<%response.Write("f4_"&nu_chamada_falta)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota">
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
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>">
      </div></td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>">
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
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota">
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
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50"  >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" class="nota" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>">
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
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="text" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="text" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="text" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota">
      </div></td>
    <td width="50" >
<div align="center"> 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="text" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota">
      </div></td>
  </tr>
  <%
	
	end if	
end if
check = check+1
max=nu_chamada
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
              <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?or=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" value="Cancelar">
            </div></td>
          <td width="50%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
              <input type="submit" name="Submit" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" value="<%=unidades%>">
              <input name="curso" type="hidden" id="curso" value="<%=grau%>">
              <input name="etapa" type="hidden" id="etapa" value="<%=serie%>">
              <input name="turma" type="hidden" id="turma" value="<%=turma%>">
              
              <input name="max" type="hidden" id="max" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>">
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
              <input type="submit" name="Submit2" value="Confirmar" class="botao_prosseguir">
              <input name="unidade" type="hidden" id="unidade" value="<%=unidade%>">
              <input name="curso" type="hidden" id="curso" value="<%=curso%>">
              <input name="etapa" type="hidden" id="etapa" value="<%=etapa%>">
              <input name="turma" type="hidden" id="turma" value="<%=turma%>">
             
              <input name="max" type="hidden" id="max" value="<% =max%>">
              <input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>">
              <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>">
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
%>