<%On Error Resume Next
%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<html>
<head>
<title>Web Acad&ecirc;mico</title>
<meta http-equiv="Content-Type" content="Hidden/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../../../estilos.css" type="text/css">

<script type="Hidden/javascript" src="../js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
<style>
<!--table
	{mso-displayed-decimal-separator:"\,";
	mso-displayed-thousand-separator:"\.";}
@page
	{margin:.98in .79in .98in .79in;
	mso-header-margin:.49in;
	mso-footer-margin:.49in;
	mso-page-orientation:landscape;}
-->
</style>
</head>



<%opt = REQUEST.QueryString("obr")
p = REQUEST.QueryString("p")
'obr=split(opt,"_")

'unidade = obr(0)
'curso = obr(1)
'co_etapa = obr(2)
'turma = obr(3)
'periodo = obr(4)

alunos=0
sxF=0
sxM=0

obr = split(opt,"?")

unidade = obr(0)
curso = obr(1)
co_etapa = obr(2)
turma = obr(3)
ano_letivo = obr(4)

obr=unidade&"?"&curso&"?"&co_etapa&"?"&turma&"?"&ano_letivo



alunos=0
sxF=0
sxM=0

ano_letivo = session("ano_letivo")
'obr=unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&periodo&"_"&ano_letivo&"_"&d&"_"&p

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

if mes<10 then
meswrt="0"&mes
else
meswrt=mes
end if
if min<10 then
minwrt="0"&min
else
minwrt=min
end if

data = dia &"/"& meswrt &"/"& ano
horario = hora & ":"& minwrt

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2
		
		Set CON5 = Server.CreateObject("ADODB.Connection") 
		ABRIR5 = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON5.Open ABRIR5
		
		Set CON6 = Server.CreateObject("ADODB.Connection") 
		ABRIR6 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON6.Open ABRIR6

		
nota= "TB_NOTA_C"
tb="TB_Frequencia_Periodo"
		
if nota = "TB_NOTA_A" Then		
		CAMINHOn = CAMINHO_na
elseif nota = "TB_NOTA_B" Then
		CAMINHOn = CAMINHO_nb
elseif nota = "TB_NOTA_C" Then
		CAMINHOn = CAMINHO_nc
end if

		Set CON_N = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIR3

		Set CON_A = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_A.Open ABRIR

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")
un_endereco = RS0("NO_Logradouro")
un_complemento = RS0("TX_Complemento_Logradouro")
un_numero = RS0("NU_Logradouro")
un_bairro = RS0("CO_Bairro")
un_cidade = RS0("CO_Municipio")
un_uf = RS0("SG_UF")
un_tel = RS0("NUS_Telefones")
un_email = RS0("TX_EMail")
un_cep = RS0("CO_CEP")
un_ato = RS0("TX_Ato_Autorizativo")
un_cnpj = RS0("CO_CGC")


if un_ato="" or isnull(un_ato) then
separador1=0
else
separador1=1
end if

if un_complemento="" or isnull(un_complemento) then
separador2=0
else
separador2=1
end if
if un_email="" or isnull(un_email) then
separador3=0
else
separador3=1
end if

cep = un_cep/1000
cep3=Int(cep)
cep4= cep-cep3

cep4=cep4*1000
cep4 = int(cep4)

if cep4 = 0 then
cep4="000"
elseif cep4<10 then
cep4="00"&cep4
elseif cep4>=10 And cep4<100 then
cep4="0"&cep4
end if

un_cep=cep3&"-"&cep4


		Set RS11 = Server.CreateObject("ADODB.Recordset")
		SQL11 = "SELECT * FROM TB_Municipios WHERE SG_UF ='"& un_uf &"' AND CO_Municipio = "&un_cidade
		RS11.Open SQL11, CON0

cidade= RS11("NO_Municipio")

		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Bairros WHERE CO_Bairro ="& un_bairro &"AND SG_UF ='"& un_uf&"' AND CO_Municipio = "&un_cidade
		RS4.Open SQL4, CON0
if RS4.EOF then
bairro = ""
else
bairro= RS4("NO_Bairro")
end if

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Curso")



		Set RS3e = Server.CreateObject("ADODB.Recordset")
		SQL3e = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' And CO_Curso ='"& curso &"'" 
		RS3e.Open SQL3e, CON0
		
if RS3e.EOF THEN
no_etapa="sem etapa"
else
no_etapa=RS3e("NO_Etapa")
end if

				
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)

if p="p" then
	%>
<body link="#6699CC" vlink="#6699CC" alink="#6699CC" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('window.print()')"> 
<%else
%>
<body link="#6699CC" vlink="#6699CC" alink="#6699CC" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('window.print()');Layer1.style.filter='progid:DXImageTransform.Microsoft.BasicImage(rotation=1)'"> 
<%end if%>
<table width="950" border="0" align="center" cellspacing="0" class="tb_corpo">
    <tr> 
      
    <td width="203" height="15" bgcolor="#FFFFFF"><img src="../../../../img/logo_preto.gif"> 
    </td>
      <td width="741" bgcolor="#FFFFFF"><table width="100%" border="0" align="right" cellspacing="0">
          <tr> 
            <td width="29%" rowspan="2"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>UNIDADE 
              <%
			no_unidade= ucase(no_unidade)
			response.Write(" "&no_unidade)
			%>
              </strong></font></td>
            <td width="71%" height="8" class="linhaBaixo">&nbsp; </td>
          </tr>
          <tr> 
            <td height="5" ></td>
          </tr>
          <tr> 
            <td colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">	
              <%response.Write("CNPJ: "&un_cnpj)
			if separador1=1then
			response.Write(" - "&un_ato)
			end if
			%>
              </font></td>
          </tr>
          <tr> 
            <td colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%response.Write(un_endereco&", "&un_numero)
			if separador2=1then
			response.Write(" - "&un_complemento&" - "&bairro&" - "&un_cep)
			else
			response.Write(" - "&bairro&" - "&un_cep)
			end if
			%>
              </font></td>
          </tr>
          <tr> 
            <td colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%response.Write(cidade&" - "&un_uf)%>
              </font></td>
          </tr>
          <tr> 
            <td colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%response.Write("Tel: "&un_tel)
			if separador1=1then
			response.Write(" - E-mail: "&un_email)
			end if
			%>
              </font></td>
          </tr>
          <tr> 
            <td colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
               </font></td>
          </tr>
          <tr> 
            
          <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>FREQ&Uuml;&Ecirc;NCIA 
            DE ALUNOS</strong></font></td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td height="12" colspan="2" bgcolor="#EEEEEE" class="linhaTopoL"><table width="900" border="0" cellspacing="0">
          <tr class="tabela"> 
            <td width="77" height="12" align="right" bgcolor="#EEEEEE" > <div align="right"> 
                <strong>Ano Letivo:</strong></div></td>
            <td width="195" height="12" bgcolor="#EEEEEE" > 
              <%response.Write(ano_letivo)%>
            </td>
            <td width="49">&nbsp;</td>
            <td width="195">&nbsp;</td>
            <td width="56">&nbsp;</td>
            <td width="316">&nbsp;</td>
          </tr>
          <tr class="tabela"> 
            <td width="77" height="12" bgcolor="#EEEEEE"> <div align="right"><strong>Curso:</strong></div></td>
            <td width="195" height="12" bgcolor="#EEEEEE"> 
              <%
response.Write(no_curso)%>
            </td>
            <td width="49" height="12" bgcolor="#EEEEEE"> <div align="right"> 
                <strong>Etapa:</strong></div></td>
            <td width="195"> 
              <%response.Write(no_etapa)%>
            </td>
            <td width="56"><div align="right"><strong>Per&iacute;odo:</strong></div></td>
            <td width="316">
              <% response.Write(no_per)%>
            </td>
          </tr>
          <tr class="tabela"> 
            <td width="77" height="12" bgcolor="#EEEEEE"> <div align="right"><strong><strong>Turma:</strong></strong></div></td>
            <td width="195" height="12" bgcolor="#EEEEEE"> 
              <%response.Write(turma)%>
            </td>
            <td width="49" height="12" bgcolor="#EEEEEE"><strong>Disciplina:</strong></td>
            <td width="195">
              <% response.Write(no_mat)%>
              &nbsp;</td>
            <td width="56">&nbsp;</td>
            <td width="316">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr> 
      <td height="10" colspan="2" bgcolor="#EEEEEE"> </td>
    </tr>
    <tr> 
      
    <td colspan="2">
<%
tb="TB_Frequencia_Periodo"
 
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL_A = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso&"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
		Set RS = CON_A.Execute(SQL_A)

calc = calc*1


%> 
 <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="20" class="tabelaTit"> <div align="center">&nbsp;N&ordm;</div></td>
    <td width="780" class="tabelaTit"> 
      <div align="left">Nome</div></td>
   <%if session("ano_letivo")>=2017 then%> 
    <td width="50" class="tabelaTit"> 
      <div align="center">&nbsp;TRI1</div></td>
    <td width="50" class="tabelaTit"> 
      <div align="center">&nbsp;TRI2</div></td>
    <td width="50" class="tabelaTit"> 
      <div align="center">&nbsp;TRI3</div></td>
   <%else%>      
    <td width="50" class="tabelaTit"> 
      <div align="center">&nbsp;B1</div></td>
    <td width="50" class="tabelaTit"> 
      <div align="center">&nbsp;B2</div></td>
    <td width="50" class="tabelaTit"> 
      <div align="center">&nbsp;B3</div></td>
    <td width="50" class="tabelaTit"> 
      <div align="center">&nbsp;B4</div></td>
   <%end if%>        
  </tr>
  <%
check = 2
nu_chamada_ckq = 0

While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")
	
if (nu_chamada_ckq <>nu_chamada - 1) then
	teste_nu_chamada = nu_chamada-nu_chamada_ckq
	teste_nu_chamada=teste_nu_chamada-1
	
			for k=1 to teste_nu_chamada 
				nu_chamada_falta=nu_chamada_ckq+1
		
		%>
  <tr> 
    <td width="20" class="tabela_fundo_linha_falta"> <input name="nu_chamada_<%=nu_chamada_falta %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>">&nbsp;</td>
    <td width="780" class="tabela_fundo_linha_falta"><input name="nu_matricula_<%=nu_chamada_falta %>" type="hidden" value="falta">&nbsp;</td>
    <td width="50" class="tabela_fundo_linha_falta">&nbsp;</td>
    <td width="50" class="tabela_fundo_linha_falta">&nbsp;</td>
    <td width="50" class="tabela_fundo_linha_falta">&nbsp;</td>
   <%if session("ano_letivo")<2017 then%>      
    <td width="50" class="tabela_fundo_linha_falta">&nbsp;</td>
     <%end if%> 
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
  <tr id="<%response.Write("celula"&NU_Chamada)%>"> 
    <td width="20" class="tabela"> 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780"  class="tabela"> 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
        <td width="50"  class="tabela">
<div align="center">&nbsp; 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="Hidden" class="nota" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>"><%response.Write("&nbsp;"&va_f1)%>
      </div></td>
        <td width="50"  class="tabela">
<div align="center">&nbsp; 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="Hidden" class="nota" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>"><%response.Write("&nbsp;"&va_f2)%>
      </div></td>
        <td width="50"  class="tabela">
<div align="center">&nbsp; 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="Hidden" class="nota" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>"><%response.Write("&nbsp;"&va_f3)%>
      </div></td>
   <%if session("ano_letivo")<2017 then%>      
        <td width="50"  class="tabela">
<div align="center">&nbsp; 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="Hidden" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota"><%response.Write("&nbsp;"&va_f4)%>
      </div></td>
     <%end if%>       
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
    <td width="20" class="tabela"> 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" class="tabela"> 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
        <td width="50"  class="tabela">
<div align="center">&nbsp; 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="Hidden" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota"><%response.Write("&nbsp;"&va_f1)%>
      </div></td>
        <td width="50"  class="tabela">
<div align="center">&nbsp; 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="Hidden" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota"><%response.Write("&nbsp;"&va_f2)%>
      </div></td>
        <td width="50"  class="tabela">
<div align="center">&nbsp; 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="Hidden" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota"><%response.Write("&nbsp;"&va_f3)%>
      </div></td>
   <%if session("ano_letivo")<2017 then%>      
        <td width="50"  class="tabela">
<div align="center">&nbsp; 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="Hidden" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota"><%response.Write("&nbsp;"&va_f4)%>
      </div></td>
     <%end if%>  
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
    <td width="20" class="tabela"> 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" class="tabela"> 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" type="hidden" value="<%=nu_matricula%>"> 
    </td>
        <td width="50"  class="tabela">
<div align="center">&nbsp; 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="Hidden" class="nota" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>"><%response.Write("&nbsp;"&va_f1)%>
      </div></td>
        <td width="50"  class="tabela">
<div align="center">&nbsp; 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="Hidden" class="nota" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>"><%response.Write("&nbsp;"&va_f2)%>
      </div></td>
        <td width="50"  class="tabela">
<div align="center">&nbsp; 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="Hidden" class="nota" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>"><%response.Write("&nbsp;"&va_f3)%>
      </div></td>
   <%if session("ano_letivo")<2017 then%>      
        <td width="50"  class="tabela">
<div align="center">&nbsp; 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="Hidden" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota"><%response.Write("&nbsp;"&va_f4)%>
      </div></td>
     <%end if%>  
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
    <td width="20" class="tabela"> 
      <%response.Write(NU_Chamada)%>
    </td>
    <td width="780" class="tabela"> 
      <%response.Write(NO_Aluno)%>
      <input name="nu_chamada_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_chamada%>"> 
      <input name="nu_matricula_<%=nu_chamada %>" class="borda_edit" type="hidden" value="<%=nu_matricula%>"> 
    </td>
        <td width="50"  class="tabela">
<div align="center">&nbsp; 
        <input name="<%response.Write("f1_"&nu_chamada)%>" type="Hidden" id="<%response.Write("f1_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f1%>" class="nota"><%response.Write("&nbsp;"&va_f1)%>
      </div></td>
        <td width="50"  class="tabela">
<div align="center">&nbsp; 
        <input name="<%response.Write("f2_"&nu_chamada)%>" type="Hidden" id="<%response.Write("f2_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f2%>" class="nota"><%response.Write("&nbsp;"&va_f2)%>
      </div></td>
        <td width="50"  class="tabela">
<div align="center">&nbsp; 
        <input name="<%response.Write("f3_"&nu_chamada)%>" type="Hidden" id="<%response.Write("f3_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f3%>" class="nota"><%response.Write("&nbsp;"&va_f3)%>
      </div></td>
   <%if session("ano_letivo")<2017 then%>      
        <td width="50"  class="tabela">
<div align="center">&nbsp; 
        <input name="<%response.Write("f4_"&nu_chamada)%>" type="Hidden" id="<%response.Write("f4_"&nu_chamada)%>" onFocus="mudar_cor_focus(<%response.Write("celula"&NU_Chamada)%>)" onBlur="<%response.Write(onblur&"(celula"&NU_Chamada&")")%>" value="<%=va_f4%>" class="nota"><%response.Write("&nbsp;"&va_f4)%>
      </div></td>
     <%end if%>  
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
</table>
	</td>
    <tr> 
      <td class="linhaTopoL"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Sistema 
        Diretor - WEB ACAD&Ecirc;MICO</font></td>
      <td class="linhaTopoR"> <div align="right"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif">Impresso 
          em 
          <%response.Write(data &" �s "&horario)%>
          </font> </div></td>
    </tr>
  </table>

</body>
</html>
<%If Err.number<>0 then
errnumb = Err.number
errdesc = Err.Description
lsPath = Request.ServerVariables("SCRIPT_NAME")
arPath = Split(lsPath, "/")
GetFileName =arPath(UBound(arPath,1))
passos = 0
for way=0 to UBound(arPath,1)
passos=passos+1
next
seleciona1=passos-2
pasta=arPath(seleciona1)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>