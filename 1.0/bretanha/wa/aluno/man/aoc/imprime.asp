<%'On Error Resume Next%>
<html>
<head>
<title>Web Acad&ecirc;mico</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">

<script type="text/javascript" src="../../../../js/global.js"></script>
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
<!--#include file="../../../../inc/funcoes.asp"-->


<!--#include file="../../../../inc/caminhos.asp"-->







<%opt = REQUEST.QueryString("obr")
ano_letivo = session("ano_letivo") 
dados= split(opt, "?" )
cod= dados(0)
ordem= dados(1)
tp_ocor= dados(2)
dia_de= dados(3)
mes_de= dados(4)
ano_de= dados(5)
h_de= dados(6)
min_de= dados(7)
dia_ate= dados(8)
mes_ate= dados(9)
ano_ate= dados(10)
h_ate= dados(11)
min_ate= dados(12)

data_de=mes_de&"/"&dia_de&"/"&ano_de


dia_de=dia_de*1
mes_de=mes_de*1
h_de=h_de*1
min_de=min_de*1

if dia_de<10 then
dia_de="0"&dia_de
end if
if mes_de<10 then
mes_de="0"&mes_de
end if
if h_de<10 then
h_de="0"&h_de
end if
if min_de<10 then
min_de="0"&min_de
end if

hora_de=h_de&":"&min_de

data_inicio=dia_de&"/"&mes_de&"/"&ano_de&", "&hora_de

data_ate=mes_ate&"/"&dia_ate&"/"&ano_ate

dia_ate=dia_ate*1
mes_ate=mes_ate*1
h_ate=h_ate*1
min_ate=min_ate*1

if dia_ate<10 then
dia_ate="0"&dia_ate
end if
if mes_ate<10 then
mes_ate="0"&mes_ate
end if
if h_ate<10 then
h_ate="0"&h_ate
end if
if min_ate<10 then
min_ate="0"&min_ate
end if




hora_ate=h_ate&":"&min_ate	
data_fim=dia_ate&"/"&mes_ate&"/"&ano_ate&", "&hora_ate

Select case ordem

case "dt"
ordena="DA_Ocorrencia,HO_Ocorrencia"

case "oc"
ordena="CO_Ocorrencia"

case "pr"
ordena="CO_Professor"

case "di"
ordena="NO_Materia"

case "au"
ordena="NU_Aula"

case "at"
ordena="CO_Usuario"


end select


		Set CON10 = Server.CreateObject("ADODB.Connection") 
		ABRIR10 = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON10.Open ABRIR10

		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2
		
		Set CONp = Server.CreateObject("ADODB.Connection") 
		ABRIRp = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONp.Open ABRIRp

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON_al = Server.CreateObject("ADODB.Connection") 
		ABRIR_al = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_al.Open ABRIR_al
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1= "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON3 = Server.CreateObject("ADODB.Connection") 
		ABRIR3 = "DBQ="& CAMINHO_o & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod
		RS.Open SQL, CON_al

codigo = RS("CO_Matricula")
nome_prof = RS("NO_Aluno")


		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod
		RS.Open SQL, CON_al
cod=opt

ano_aluno = RS("NU_Ano")
rematricula = RS("DA_Rematricula")
situacao = RS("CO_Situacao")
encerramento= RS("DA_Encerramento")
unidade= RS("NU_Unidade")
curso= RS("CO_Curso")
co_etapa= RS("CO_Etapa")
turma= RS("CO_Turma")
cham= RS("NU_Chamada")

'ano_letivo = session("ano_letivo")
'obr=unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&periodo&"_"&ano_letivo




		Set CON4 = Server.CreateObject("ADODB.Connection")
		ABRIR4 = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4

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



		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' and CO_Curso ='"& curso &"'"  
		RS3.Open SQL3, CON0
		
if RS3.EOF THEN
no_etapa="sem etapa"
else
no_etapa=RS3("NO_Etapa")
end if

	%>



<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('window.print()')"> 
<br>
<table width="950" border="0" align="center" cellspacing="0" class="tb_corpo"
>
  <tr> 
    <td width="203" height="15" bgcolor="#FFFFFF"><img src="../../../../img/logo_preto.gif" width="175" height="130"> 
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
          <td colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            &nbsp; </font></td>
        </tr>
        <tr> 
          <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>OCORR&Ecirc;NCIAS 
            DO ALUNO</strong></font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="12" colspan="2" bgcolor="#EEEEEE" class="linhaTopoL"><table width="912" border="0" cellspacing="0">
        <tr class="tabela"> 
          <td width="80" height="12" align="right" bgcolor="#EEEEEE" > <div align="right"> 
              <strong>Ano Letivo:</strong></div></td>
          <td width="200" height="12" bgcolor="#EEEEEE" > <%response.Write(ano_letivo)%> </td>
          <td width="40"><div align="right"><strong>Etapa:</strong></div></td>
          <td width="200">
            <%response.Write(no_etapa)%>
          </td>
          <td width="97" bgcolor="#EEEEEE"><div align="right"><strong> Matr&iacute;cula:</strong></div></td>
          <td width="442" bgcolor="#EEEEEE"> <%response.Write(codigo)%> </td>
        </tr>
        <tr class="tabela"> 
          <td width="80" height="12" bgcolor="#EEEEEE"> <div align="right"><strong>Curso:</strong></div></td>
          <td width="200" height="12" bgcolor="#EEEEEE"> <%
response.Write(no_curso)%> </td>
          <td width="40" height="12" bgcolor="#EEEEEE"> <div align="right"><strong>Ocorr&ecirc;ncia:</strong> 
            </div></td>
          <td width="200"><%IF tp_ocor="999999" or tp_ocor=999999 then
		  response.Write("Todas")
		  else
		  
		  					   	Set RSto = Server.CreateObject("ADODB.Recordset")
						SQLto = "SELECT * FROM TB_Tipo_Ocorrencia WHERE CO_Ocorrencia ="& tp_ocor
						RSto.Open SQLto, CON0
						no_ocorrencia=RSto("NO_Ocorrencia")
					  response.Write(no_ocorrencia)
					  
					  end if%>&nbsp; </td>
          <td height="12" bgcolor="#EEEEEE"><div align="right"><strong>Nome: </strong></div></td>
          <td height="12" bgcolor="#EEEEEE"> <%response.Write(nome_prof)%> </td>
        </tr>
        <tr class="tabela"> 
          <td width="80" height="12" bgcolor="#EEEEEE"> <div align="right"><strong><strong>Turma:</strong></strong></div></td>
          <td width="200" height="12" bgcolor="#EEEEEE"> <%response.Write(turma)%> </font></td>
          <td width="40" height="12" bgcolor="#EEEEEE"></td>
          <td width="200">&nbsp;</td>
          <td width="80">&nbsp;</td>
          <td width="300">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="10" colspan="2" bgcolor="#EEEEEE"> </td>
  </tr>
  <tr> 
    <td colspan="2"> <table width="1000" border="0" cellspacing="0" bordercolor="#000000" bgcolor="#FFFFFF">
        <tr > 
          <td height="10" colspan="4" class="tabelaTit"><div align="left">Ocorr&ecirc;ncias Resumidas</div></td>
        </tr>
        <tr > 
          <td class="tabelaTit" width="30">&nbsp;</td>
          <td class="tabelaTit" width="594" height="10"><div align="left">Ocorr&ecirc;ncia</div></td>
          <td class="tabelaTit" width="78" height="10"><div align="center">Quantidade</div></td>
          <td  class="tabelaTit" width="308">&nbsp;</td>
        </tr>
        <%
		Set RSo = Server.CreateObject("ADODB.Recordset")
if tp_ocor=999999 or tp_ocor="999999" or tp_ocor="" or isnull(tp_ocor) then

		SQLo = "SELECT * FROM TB_Ocorrencia_Aluno WHERE CO_Matricula ="& codigo&" AND (DA_Ocorrencia BETWEEN #"&data_de&"# AND #"&data_ate&"#) order BY CO_Ocorrencia"

else
		SQLo = "SELECT * FROM TB_Ocorrencia_Aluno WHERE CO_Matricula ="& codigo&" AND CO_Ocorrencia ="& tp_ocor&" AND (DA_Ocorrencia BETWEEN #"&data_de&"# AND #"&data_ate&"#) order BY CO_Ocorrencia"

end if
		RSo.Open SQLo, CON3

if RSo.EOF	then

	
%>
        <tr> 
          <td width="30" class="tabela">&nbsp;</td>
          <td class="tabela"><div align="left"> Nenhuma ocorrência cadastrada para este Aluno</div></td>
          <td class="tabela"><div align="center"></div></td>
          <td class="tabela">&nbsp;</td>
        </tr>
        <%else
check = 2
count_ocorr=0
acum_ocorr=0
co_ocorr_check="nada"
WHILE not RSo.EOF
  if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if 
  
co_ocorrencia=RSo("CO_Ocorrencia")
da_ocorrencia=RSo("DA_Ocorrencia")
ho_ocorrencia=RSo("HO_Ocorrencia")

IF co_ocorr_check=co_ocorrencia then
RSo.Movenext
else
co_ocorr_check=co_ocorrencia
Set RSco = Server.CreateObject("ADODB.Recordset")
if co_ocorrencia="" or isnull(co_ocorrencia) then

		SQLco = "SELECT COUNT(DA_Ocorrencia) AS CT FROM TB_Ocorrencia_Aluno WHERE CO_Matricula ="& codigo&" AND ISNULL(CO_Ocorrencia) AND(DA_Ocorrencia BETWEEN #"&data_de&"# AND #"&data_ate&"#)"

else
		SQLco = "SELECT COUNT(CO_Ocorrencia) AS CT FROM TB_Ocorrencia_Aluno WHERE CO_Matricula ="& codigo&" AND CO_Ocorrencia ="& co_ocorrencia&" AND (DA_Ocorrencia BETWEEN #"&data_de&"# AND #"&data_ate&"#)"

end if
		RSco.Open SQLco, CON3
		
count_ocorr=RSco("CT")

acum_ocorr=acum_ocorr+count_ocorr

if co_ocorrencia="" or ISNULL(co_ocorrencia) OR co_ocorrencia="999999" OR co_ocorrencia=999999 OR co_ocorrencia="0" OR co_ocorrencia=0 THEN
ELSE
 		Set RSto = Server.CreateObject("ADODB.Recordset")
		SQLto = "SELECT * FROM TB_Tipo_Ocorrencia WHERE CO_Ocorrencia ="& co_ocorrencia
		RSto.Open SQLto, CON0
no_ocorrencia=RSto("NO_Ocorrencia")

END IF					  
%>
        <tr > 
          <td width="30" class="tabela">&nbsp;</td>
          <td height="15" class="tabela"> <div align="left"> 
              <%response.Write(no_ocorrencia)%>
            </div></td>
          <td class="tabela"><div align="center"> 
              <%response.Write(count_ocorr)%>
            </div></td>
          <td  class="tabela">&nbsp;</td>
        </tr>
        <%

check = check+1
RSo.Movenext
end if
WEND
END IF%>
        <tr> 
          <td width="30"  class="tabelaTit">&nbsp;</td>
          <td height="15"  class="tabelaTit"><div align="left">Total</div></td>
          <td  class="tabelaTit"><div align="center"> 
              <%response.Write(acum_ocorr)%>
            </div></td>
          <td  class="tabela">&nbsp;</td>
        </tr>
        <tr> 
          <td height="15"  class="tabelaTit" colspan="4"><table width="1000" border="0" cellspacing="0" bordercolor="#000000" bgcolor="#FFFFFF">
              <tr> 
                <td height="10" colspan="7" class="tabelaTit"><div align="left">Ocorr&ecirc;ncias 
                    Detalhadas</div></td>
              </tr>
              <tr> 
                <td width="30" height="10"  class="tabelaTit">&nbsp; </td>
                <td width="130"  class="tabelaTit"> <div align="left">Data / Hora</div></td>
                <td width="305"  class="tabelaTit"> <div align="left">Ocorr&ecirc;ncia<font class="form_dado_texto"> 
                    </font></div></td>
                <td width="255"  class="tabelaTit"> <div align="left">Professor<font class="form_dado_texto"> 
                    <input name="cod" type="hidden" id="cod2" value="<%=codigo%>">
                    <input name="nome" type="hidden" class="textInput" id="nome"  value="<%response.Write(nome_prof)%>" size="75" maxlength="50">
                    <input name="tp_ocor" type="hidden" class="textInput" id="tp_ocor3"  value="<%response.Write(tp_ocor)%>" size="75" maxlength="50">
                    <input name="data_de" type="hidden" class="textInput" id="tp_ocor3"  value="<%response.Write(data_de)%>" size="75" maxlength="50">
                    <input name="hora_de" type="hidden" class="textInput" id="tp_ocor3"  value="<%response.Write(hora_de)%>" size="75" maxlength="50">
                    <input name="data_inicio" type="hidden" class="textInput" id="tp_ocor3"  value="<%response.Write(data_inicio)%>" size="75" maxlength="50">
                    <input name="data_ate" type="hidden" class="textInput" id="tp_ocor3"  value="<%response.Write(data_ate)%>" size="75" maxlength="50">
                    </font></div></td>
                <td width="160" class="tabelaTit"> <div align="left">Disciplina<font class="form_dado_texto"> 
                    <input name="hora_ate" type="hidden" class="textInput" id="hora_ate2"  value="<%response.Write(hora_ate)%>" size="75" maxlength="50">
                    <input name="data_fim" type="hidden" class="textInput" id="hora_ate2"  value="<%response.Write(data_fim)%>" size="75" maxlength="50">
                    </font></div></td>
                <td width="40"  class="tabelaTit"> <div align="center">Aula</div></td>
                <td width="250"  class="tabelaTit"> <div align="center">Atendido 
                    por</div></td>
              </tr>
              <%			
				
		Set RSo = Server.CreateObject("ADODB.Recordset")
if tp_ocor=999999 or tp_ocor="999999" then
		SQLo = "SELECT * FROM TB_Ocorrencia_Aluno WHERE CO_Matricula ="& codigo&" AND (DA_Ocorrencia BETWEEN #"&data_de&"# AND #"&data_ate&"#) order BY "&ordena&""

else
		SQLo = "SELECT * FROM TB_Ocorrencia_Aluno WHERE CO_Matricula ="& codigo&" AND CO_Ocorrencia ="& tp_ocor&" AND (DA_Ocorrencia BETWEEN #"&data_de&"# AND #"&data_ate&"#) order BY "&ordena&""

end if
		RSo.Open SQLo, CON3

if RSo.EOF	then	
%>
              <tr> 
                <td width="30"  class="tabela">&nbsp;</td>
                <td width="130"  class="tabela"> <div align="left"></div></td>
                <td width="305"  class="tabela"> <div align="left"> Nenhuma ocorrência 
                    cadastrada para este Aluno</div></td>
                <td width="255"  class="tabela"> <div align="left">&nbsp;</div></td>
                <td width="160"  class="tabela"> <div align="left">&nbsp;</div></td>
                <td width="40"  class="tabela">&nbsp;</td>
                <td width="250"  class="tabela">&nbsp;</td>
              </tr>
              <%else
check = 2
WHILE not RSo.EOF
  if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if 
  
co_ocorrencia=RSo("CO_Ocorrencia")
da_ocorrencia=RSo("DA_Ocorrencia")
ho_ocorrencia=RSo("HO_Ocorrencia")
ass_ocorrencia=RSo("CO_Assunto")
au_ocorrencia=RSo("NU_Aula")
cp_ocorrencia=RSo("CO_Professor")
di_ocorrencia=RSo("NO_Materia")
ob_ocorrencia=RSo("TX_Observa")
cu_ocorrencia=RSo("CO_Usuario")

if di_ocorrencia="" or isnull(di_ocorrencia) then
no_materia=""
else

 		Set RSnomat = Server.CreateObject("ADODB.Recordset")
		SQLnomat = "SELECT * FROM TB_Materia Where CO_Materia='"&di_ocorrencia&"'"
		RSnomat.Open SQLnomat, CON0

no_materia=RSnomat("NO_Materia")
end if

'IF co_ocorr_check=co_ocorrencia then
'RSo.Movenext
'else
'co_ocorr_check=co_ocorrencia
'Set RSco = Server.CreateObject("ADODB.Recordset")
if co_ocorrencia="" or ISNULL(co_ocorrencia) then
no_ocorrencia=""

'		SQLco = "SELECT COUNT(CO_Ocorrencia) AS CT FROM TB_Ocorrencia_Aluno WHERE CO_Matricula ="& cod&" AND CO_Ocorrencia="&co_ocorrencia&" AND (DA_Ocorrencia BETWEEN #"&data_de&"# AND #"&data_ate&"#)"

else
'		SQLco = "SELECT COUNT(CO_Ocorrencia) AS CT  FROM TB_Ocorrencia_Aluno WHERE CO_Matricula ="& cod&" AND CO_Ocorrencia ="& tp_ocor&" AND (DA_Ocorrencia BETWEEN #"&data_de&"# AND #"&data_ate&"#)"

'end if
'		RSco.Open SQLco, CON3
		
'count_ocor=RSco("ct")
 
 		Set RSto = Server.CreateObject("ADODB.Recordset")
		SQLto = "SELECT * FROM TB_Tipo_Ocorrencia WHERE CO_Ocorrencia ="& co_ocorrencia
		RSto.Open SQLto, CON0
no_ocorrencia=RSto("NO_Ocorrencia")

end if

if cp_ocorrencia="" or isnull(cp_ocorrencia)or cp_ocorrencia="999999" or cp_ocorrencia=999999  then
no_professor=""
else


		Set RSp = Server.CreateObject("ADODB.Recordset")
		SQLp = "SELECT * FROM TB_Professor WHERE CO_Professor ="& cp_ocorrencia
		RSp.Open SQLp, CONp
		
IF RSp.EOF then
else
co_professor=RSp("CO_Usuario")
end if

		Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& co_professor
		RSu.Open SQLu, CON

IF RSu.EOF then
no_professor=""
else
no_professor=RSu("NO_Usuario")
end if		
end if
			
if cu_ocorrencia="" or isnull(cu_ocorrencia) then
no_atendido=""
else
		Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& cu_ocorrencia
		RSu.Open SQLu, CON

IF RSu.EOF then
else
no_atendido=RSu("NO_Usuario")
end if
		
end if
obr=cod&"?"&da_ocorrencia&"?"&ho_ocorrencia&"?"&co_ocorrencia&"?PED"
Session("tp_ocor")=tp_ocor
Session("data_de")=data_de
Session("hora_de")=hora_de
Session("data_inicio")=data_inicio
Session("data_ate")=data_ate
Session("hora_ate")=hora_ate
Session("data_fim")=data_fim


data_split= Split(da_ocorrencia,"/")
dia=data_split(0)
mes=data_split(1)
ano=data_split(2)

hora_split= Split(ho_ocorrencia,":")
hora=hora_split(0)
min=hora_split(1)

dia=dia*1
mes=mes*1
hora=hora*1
min=min*1

if dia<10 then
dia="0"&dia
end if
if mes<10 then
mes="0"&mes
end if
if hora<10 then
hora="0"&hora
end if
if min<10 then
min="0"&min
end if
da_show=dia&"/"&mes&"/"&ano
hora_show=hora&":"&min
%>
              <tr> 
                <td width="30"  class="tabela">&nbsp; </td>
                <td width="130"  class="tabela"> 
                  <%response.Write(da_show&", "&hora_show)%>
                  <div align="left"></div>
                  <div align="left"></div></td>
                <td width="305"  class="tabela"> <div align="left"> 
                    <%response.Write(no_ocorrencia)%>
                    &nbsp; </div></td>
                <td width="255"  class="tabela"> 
                  <%response.Write(no_professor)%>
                  &nbsp; <div align="left"></div></td>
                <td width="160"  class="tabela"> 
                  <%response.Write(no_materia)%>
                  &nbsp; <div align="left"></div></td>
                <td width="40"  class="tabela"> <div align="center"> 
                    <%response.Write(au_ocorrencia)%>
                    &nbsp; </div></td>
                <td width="250"  class="tabela"> <div align="center"> 
                    <%response.Write(no_atendido)%>
                    &nbsp; </div></td>
              </tr>
              <%check = check+1
RSo.Movenext
'end if
WEND%>
              <tr> 
                <td colspan="7"  class="tabela"> <div align="center"> </div>
                  <div align="left"></div>
                  <div align="left"></div>
                  <div align="left"> 
                    <hr width="1000" size="1">
                  </div></td>
              </tr>
              <tr class="tabela"> 
                <td colspan="7">&nbsp;</td>
              </tr>
              <%

END IF%>
            </table></td>
        </tr>
      </table>
      
      
    </td>
  <tr> 
    <td class="linhaTopoL"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Sistema 
      Diretor - WEB ACAD&Ecirc;MICO</font></td>
    <td class="linhaTopoR"> <div align="right"> 
        <font size="1" face="Verdana, Arial, Helvetica, sans-serif">Impresso 
          em 
          <%
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
		  
		  
		  response.Write(data &" às "&horario)%>
          </font>
      </div></td>
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