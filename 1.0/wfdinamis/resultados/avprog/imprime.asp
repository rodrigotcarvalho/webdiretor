<%'On Error Resume Next%>
<html>
<head>
<title>Web Fam&iacute;lia</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../estilo.css" rel="stylesheet" type="text/css">

<script type="text/javascript" src="../../js/global.js"></script>
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
<!--#include file="../../inc/caminhos.asp"-->
<!--#include file="../../inc/funcoes.asp"-->
<!--#include file="../../inc/funcoes2.asp"-->


<%opt = REQUEST.QueryString("obr")
dados_opt= split(opt, "?" )
cod= dados_opt(0)
periodo_check= dados_opt(1)

 	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0

 	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1
	
	Set CON2 = Server.CreateObject("ADODB.Connection") 
	ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON2.Open ABRIR2		

	SQL2 = "select * from TB_Usuario where CO_Usuario = " & cod 
	set RS2 = CON.Execute (SQL2)
	
nome_aluno= RS2("NO_Usuario")

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod
		RS.Open SQL, CON1


ano_aluno = RS("NU_Ano")
rematricula = RS("DA_Rematricula")
situacao = RS("CO_Situacao")
encerramento= RS("DA_Encerramento")
unidade= RS("NU_Unidade")
curso= RS("CO_Curso")
etapa= RS("CO_Etapa")
turma= RS("CO_Turma")
cham= RS("NU_Chamada")

ano_letivo = session("ano_letivo")
obr=unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&periodo_check&"_"&ano_letivo

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



		Set RS_tb = Server.CreateObject("ADODB.Recordset")
		SQL_tb = "SELECT * FROM TB_Da_Aula WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso&"' AND CO_Etapa ='"& etapa&"' AND CO_Turma ='"& turma&"'"
		RS_tb.Open SQL_tb, CON2

notaFIL=RS_tb("TP_Nota")



if notaFIL ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na

elseif notaFIL="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb

elseif notaFIL ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc
else
		response.Write("ERRO")
end if	

		Set CON3 = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3

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

cep= lEFT(un_cep, 5)
cep3= Right(un_cep, 3)

un_cep=cep&"-"&cep


		Set RS11 = Server.CreateObject("ADODB.Recordset")
		SQL11 = "SELECT * FROM TB_Municipios WHERE SG_UF ='"& un_uf &"' AND CO_Municipio = "&un_cidade
		RS11.Open SQL11, CON0

cidade= RS11("NO_Municipio")

		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Bairros WHERE CO_Bairro ="& un_bairro &"AND SG_UF ='"& un_uf&"' AND CO_Municipio = "&un_cidade
		RS4.Open SQL4, CON0
if RS4.EOF then
bairro = "&nbsp;"
else
bairro= RS4("NO_Bairro")
end if


		Set RSPER = Server.CreateObject("ADODB.Recordset")
		SQLPER = "SELECT * FROM TB_Periodo where NU_Periodo="&periodo_check
		RSPER.Open SQLPER, CON0

		no_periodo=RSPER("NO_Periodo")


		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Curso")



		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& etapa &"' and CO_Curso ='"& curso &"'"  
		RS3.Open SQL3, CON0
		
if RS3.EOF THEN
no_etapa="sem etapa"
else
no_etapa=RS3("NO_Etapa")
end if

if periodo_check=2 then
width=170
else
width=200
end if
	%>



<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('window.print()')"> 
<br>
<table width="950" border="0" align="center" cellspacing="0" class="tb_corpo"
>
  <tr> 
    <td width="122" height="15" bgcolor="#FFFFFF"><div align="center"><img src="../../img/logo_preto.gif"> 
      </div></td>
    <td width="824" bgcolor="#FFFFFF"><table width="100%" border="0" align="right" cellspacing="0">
        <tr> 
          <td width="29%" rowspan="2"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>UNIDADE 
            <%
			no_unidade= ucase(no_unidade)
			response.Write("&nbsp;"&no_unidade)
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
          <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>AVALIA&Ccedil;&Otilde;ES 
            PROGRESSIVAS</strong></font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="12" colspan="2" bgcolor="#EEEEEE" class="linhaTopoL"><table width="912" border="0" cellspacing="0">
        <tr class="tabela"> 
          <td width="78" height="12" align="right" bgcolor="#EEEEEE" > <div align="right"> 
              <strong>Ano Letivo:</strong></div></td>
          <td width="168" height="12" bgcolor="#EEEEEE" > 
            <%response.Write(ano_letivo)%>
          </td>
          <td height="12" bgcolor="#EEEEEE"> <div align="right"> <strong>Etapa:</strong></div></td>
          <td> 
            <%response.Write(no_etapa)%>
          </td>
          <td width="80" bgcolor="#EEEEEE"><div align="right"><strong> Matr&iacute;cula:</strong></div></td>
          <td width="308" bgcolor="#EEEEEE"> 
            <%response.Write(cod)%>
          </td>
        </tr>
        <tr class="tabela"> 
          <td width="78" height="12" bgcolor="#EEEEEE"> <div align="right"><strong>Curso:</strong></div></td>
          <td width="168" height="12" bgcolor="#EEEEEE"> 
            <%
response.Write(no_curso)%>
          </td>
          <td width="68" height="12" bgcolor="#EEEEEE"> <div align="right"> <strong>Per&iacute;odo:</strong></div></td>
          <td width="198"> 
            <%response.Write(no_periodo)%>
          </td>
          <td height="12" bgcolor="#EEEEEE"><div align="right"><strong>Nome: </strong></div></td>
          <td height="12" bgcolor="#EEEEEE"> 
            <%response.Write(nome_aluno)%>
          </td>
        </tr>
        <tr class="tabela"> 
          <td width="78" height="12" bgcolor="#EEEEEE"> <div align="right"><strong><strong>Turma:</strong></strong></div></td>
          <td width="168" height="12" bgcolor="#EEEEEE"> 
            <%response.Write(turma)%></font>
            </td>
          <td width="68" height="12" bgcolor="#EEEEEE"></td>
          <td width="198">&nbsp;</td>
          <td width="80">&nbsp;</td>
          <td width="308">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="10" colspan="2" bgcolor="#EEEEEE"> </td>
  </tr>
  <tr> 
    <td colspan="2"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="<%response.Write(width)%>" rowspan="2" class="tabela" 
> <div align="left"><strong>Disciplina</strong></div></td>
          <td width="30" rowspan="2" class="tabela" 
> <div align="center">F</div></td>
          <td colspan="2" class="tabela" 
><div align="center">APR1</div></td>
          <td colspan="2" class="tabela" 
> <div align="center">APR2</div></td>
          <td colspan="2" class="tabela" 
> <div align="center">APR3</div></td>
          <td colspan="2" class="tabela" 
> <div align="center">APR4</div></td>
          <td colspan="2" class="tabela" 
> <div align="center">APR5</div></td>
          <td colspan="2" class="tabela" 
> <div align="center">APR6</div></td>
          <td colspan="2" class="tabela" 
> <div align="center">TEC1</div></td>
          <td colspan="2" class="tabela" 
> <div align="center">TEC2</div></td>
          <td width="30" rowspan="2" class="tabela" 
> <div align="center">SAPR</div></td>
          <td width="30" rowspan="2" class="tabela" 
> <div align="center">PR</div></td>
          <td width="30" rowspan="2" class="tabela" 
> <div align="center">MP</div></td>
<% if periodo_check=2 then%>
                          <td width="30" rowspan="2" class="tabela"
> <div align="center">EC1</div></td>
<%end if%>
<!--           <td width="200" rowspan="2" class="tabela" 
> <div align="center">Alterado em</div></td>--> 
        </tr>
        <tr> 
          <td width="30" class="tabela" 
> <div align="center">N</div></td>
          <td width="30" class="tabela" 
> <div align="center">P</div></td>
          <td width="30" class="tabela" 
> <div align="center">N</div></td>
          <td width="30" class="tabela" 
> <div align="center">P</div></td>
          <td width="30" class="tabela" 
> <div align="center">N</div></td>
          <td width="30" class="tabela" 
> <div align="center">P</div></td>
          <td width="30" class="tabela" 
> <div align="center">N</div></td>
          <td width="30" class="tabela" 
> <div align="center">P</div></td>
          <td width="30" class="tabela" 
> <div align="center">N</div></td>
          <td width="30" class="tabela" 
> <div align="center">P</div></td>
          <td width="30" class="tabela" 
> <div align="center">N</div></td>
          <td width="30" class="tabela" 
> <div align="center">P</div></td>
          <td width="30" class="tabela" 
> <div align="center">N</div></td>
          <td width="30" class="tabela" 
> <div align="center">P</div></td>
          <td width="30" class="tabela" 
> <div align="center">N</div></td>
          <td width="30" class="tabela" 
> <div align="center">P</div></td>
        </tr>
        <%
rec_lancado="sim"

		Set RSprog = Server.CreateObject("ADODB.Recordset")
		SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
		RSprog.Open SQLprog, CON0

check=2
	
while not RSprog.EOF

	materia=RSprog("CO_Materia")
	mae=RSprog("IN_MAE")
	fil=RSprog("IN_FIL")
	in_co=RSprog("IN_CO")
	nu_peso=RSprog("NU_Peso")
	ordem=RSprog("NU_Ordem_Boletim")

		Set RS1a = Server.CreateObject("ADODB.Recordset")
		SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
		RS1a.Open SQL1a, CON0
		
no_materia=RS1a("NO_Materia")

 if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if

		
		
		Set CON_N = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIRn

'for periodofil=1 to 4


		
		Set RSnFIL = Server.CreateObject("ADODB.Recordset")
	  	Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"'"
		Set RS3 = CON_N.Execute(SQL_N)


		Set RSPESO = Server.CreateObject("ADODB.Recordset")
		SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&periodo_check
		RSPESO.Open SQLPESO, CON0



if RS3.EOF then
va_apr1="&nbsp;"
va_apr2="&nbsp;"
va_apr3="&nbsp;"
va_apr4="&nbsp;"
va_apr5="&nbsp;"
va_apr6="&nbsp;"
va_apr7="&nbsp;"
va_apr8="&nbsp;"
va_v_apr1="&nbsp;"
va_v_apr2="&nbsp;"
va_v_apr3="&nbsp;"
va_v_apr4="&nbsp;"
va_v_apr5="&nbsp;"
va_v_apr6="&nbsp;"
va_v_apr7="&nbsp;"
va_v_apr8="&nbsp;"
va_sapr="&nbsp;"
va_pr="&nbsp;"
va_te="&nbsp;"
va_bon="&nbsp;"
va_me="&nbsp;"
va_mc="&nbsp;"
va_faltas="&nbsp;"
data_grav="nulo"
hora_grav="nulo"		
else
if periodo_check=1 then
va_apr1=RS3("Apr1_P1")
va_apr2=RS3("Apr2_P1")
va_apr3=RS3("Apr3_P1")
va_apr4=RS3("Apr4_P1")
va_apr5=RS3("Apr5_P1")
va_apr6=RS3("Apr6_P1")
va_apr7=RS3("Apr7_P1")
va_apr8=RS3("Apr8_P1")
va_v_apr1=RS3("V_Apr1_P1")
va_v_apr2=RS3("V_Apr2_P1")
va_v_apr3=RS3("V_Apr3_P1")
va_v_apr4=RS3("V_Apr4_P1")
va_v_apr5=RS3("V_Apr5_P1")
va_v_apr6=RS3("V_Apr6_P1")
va_v_apr7=RS3("V_Apr7_P1")
va_v_apr8=RS3("V_Apr8_P1")
va_sapr=RS3("VA_Sapr1")
va_pr=RS3("VA_Pr1")
va_te=RS3("VA_Te1")
va_bon=RS3("VA_Bon1")
va_me=RS3("VA_Me1")
va_mc=RS3("VA_Mc1")
va_faltas=RS3("NU_Faltas_P1")
data_grav=RS3("DA_Ult_Acesso")
hora_grav=RS3("HO_ult_Acesso")

va_pr=va_pr*1
va_te=va_te*1
pr=va_pr+va_te
elseif periodo_check=2 then
va_apr1=RS3("Apr1_P2")
va_apr2=RS3("Apr2_P2")
va_apr3=RS3("Apr3_P2")
va_apr4=RS3("Apr4_P2")
va_apr5=RS3("Apr5_P2")
va_apr6=RS3("Apr6_P2")
va_apr7=RS3("Apr7_P2")
va_apr8=RS3("Apr8_P2")
va_v_apr1=RS3("V_Apr1_P2")
va_v_apr2=RS3("V_Apr2_P2")
va_v_apr3=RS3("V_Apr3_P2")
va_v_apr4=RS3("V_Apr4_P2")
va_v_apr5=RS3("V_Apr5_P2")
va_v_apr6=RS3("V_Apr6_P2")
va_v_apr7=RS3("V_Apr7_P2")
va_v_apr8=RS3("V_Apr8_P2")
va_sapr=RS3("VA_Sapr2")
va_pr=RS3("VA_Pr2")
va_te=RS3("VA_Te2")
va_bon=RS3("VA_Bon2")
va_me=RS3("VA_Me2")
va_mc=RS3("VA_Mc2")
va_faltas=RS3("NU_Faltas_P2")
data_grav=RS3("DA_Ult_Acesso")
hora_grav=RS3("HO_ult_Acesso")

va_pr=va_pr*1
va_te=va_te*1
pr=va_pr+va_te
elseif periodo_check=3 then
va_apr1=RS3("Apr1_P3")
va_apr2=RS3("Apr2_P3")
va_apr3=RS3("Apr3_P3")
va_apr4=RS3("Apr4_P3")
va_apr5=RS3("Apr5_P3")
va_apr6=RS3("Apr6_P3")
va_apr7=RS3("Apr7_P3")
va_apr8=RS3("Apr8_P3")
va_v_apr1=RS3("V_Apr1_P3")
va_v_apr2=RS3("V_Apr2_P3")
va_v_apr3=RS3("V_Apr3_P3")
va_v_apr4=RS3("V_Apr4_P3")
va_v_apr5=RS3("V_Apr5_P3")
va_v_apr6=RS3("V_Apr6_P3")
va_v_apr7=RS3("V_Apr7_P3")
va_v_apr8=RS3("V_Apr8_P3")
va_sapr=RS3("VA_Sapr3")
va_pr=RS3("VA_Pr3")
va_te=RS3("VA_Te3")
va_bon=RS3("VA_Bon3")
va_me=RS3("VA_Me3")
va_mc=RS3("VA_Mc3")
va_faltas=RS3("NU_Faltas_P3")
data_grav=RS3("DA_Ult_Acesso")
hora_grav=RS3("HO_ult_Acesso")

va_pr=va_pr*1
va_te=va_te*1
pr=va_pr+va_te
elseif periodo_check=4 then
va_apr1=RS3("Apr1_EC")
va_apr2=RS3("Apr2_EC")
va_apr3=RS3("Apr3_EC")
va_apr4=RS3("Apr4_EC")
va_apr5=RS3("Apr5_EC")
va_apr6=RS3("Apr6_EC")
va_apr7=RS3("Apr7_EC")
va_apr8=RS3("Apr8_EC")
va_v_apr1=RS3("V_Apr1_EC")
va_v_apr2=RS3("V_Apr2_EC")
va_v_apr3=RS3("V_Apr3_EC")
va_v_apr4=RS3("V_Apr4_EC")
va_v_apr5=RS3("V_Apr5_EC")
va_v_apr6=RS3("V_Apr6_EC")
va_v_apr7=RS3("V_Apr7_EC")
va_v_apr8=RS3("V_Apr8_EC")
va_sapr=RS3("VA_Sapr_EC")
va_pr=RS3("VA_Pr4")
va_me=RS3("VA_Me_EC")
va_mc=RS3("VA_Mfinal")
data_grav=RS3("DA_Ult_Acesso")
hora_grav=RS3("HO_ult_Acesso")

pr=va_pr
end if

if va_me="" or isnull(va_me) or va_me="&nbsp;" then
else
va_me=va_me/10
'	decimo = va_me - Int(va_me)
'		If decimo >= 0.5 Then
'			nota_arredondada = Int(va_me) + 1
'			va_me=nota_arredondada
'		Else
'			nota_arredondada = Int(va_me)
'			va_me=nota_arredondada					
'		End If
	va_me = formatNumber(va_me,1)
end if	
	
		
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL_WF = "SELECT * FROM TB_Autoriza_WF WHERE NU_Unidade="&unidade&" AND CO_Curso='"&curso&"' and CO_Etapa='"&etapa&"'"
		RS4.Open SQL_WF, CON	
	

co_apr1=RS4("CO_apr1")
co_apr2=RS4("CO_apr2")
co_apr3=RS4("CO_apr3")
co_apr4=RS4("CO_apr4")
co_prova1=RS4("CO_prova1")
co_prova2=RS4("CO_prova2")
co_prova3=RS4("CO_prova3")
co_prova4=RS4("CO_prova4")	
		
if periodo_check=1 then		
		if co_apr1="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova1="D"then
		showprova="n"
		else 
		showprova="s"
		end if
elseif periodo_check=2 then		
		if co_apr2="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova2="D"then
		showprova="n"
		else 
		showprova="s"
		end if					
elseif periodo_check=3 then		
		if co_apr3="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova3="D"then
		showprova="n"
		else 
		showprova="s"
		end if
elseif periodo_check=4 then		
		if co_apr4="D"then
		showapr="n"
		else 
		showapr="s"
		end if
		if co_prova4="D"then
		showprova="n"
		else 
		showprova="s"
		end if
end if	
				
		
if hora_grav="nulo" then
hora_de=""
else
dados_hrd= split(hora_grav, ":" )
h_de= dados_hrd(0)
min_de= dados_hrd(1)
h_de=h_de*1
min_de=min_de*1


	if h_de<10 then
	h_de="0"&h_de
	end if
	if min_de<10 then
	min_de="0"&min_de
	end if	
	hora_de=h_de&":"&min_de
				
end if		
					
if data_grav="nulo"	then
data_inicio=""
else
		
dados_dtd= split(data_grav, "/" )
dia_de= dados_dtd(0)
mes_de= dados_dtd(1)
ano_de= dados_dtd(2)
dia_de=dia_de*1
mes_de=mes_de*1
if dia_de<10 then
dia_de="0"&dia_de
end if
if mes_de<10 then
mes_de="0"&mes_de
end if
data_inicio=dia_de&"/"&mes_de&"/"&ano_de&", "&hora_de
end if

				
end if

if (isnull(va_apr1) OR va_apr1="&nbsp;") and (ISNULL(va_apr2) OR va_apr2="&nbsp;")and (ISNULL(va_apr3) OR va_apr3="&nbsp;")and (ISNULL(va_apr4) OR va_apr4="&nbsp;")and (ISNULL(va_apr5)  OR va_apr5="&nbsp;")and (ISNULL(va_apr6) OR  va_apr6="&nbsp;") and (ISNULL(va_apr7) OR va_apr7="&nbsp;")and (ISNULL(va_apr8) OR va_apr8="&nbsp;")and (ISNULL(va_sapr) OR va_sapr="&nbsp;")  then
data_inicio=""
va_faltas=""
else
		if (va_apr1=0 OR va_apr1="0") and (va_apr2=0 OR va_apr2="0")and (va_apr3=0 OR va_apr3="0")and (va_apr4=0 OR va_apr4="0")and (va_apr5=0 OR va_apr5="0")and (va_apr6=0 OR va_apr6="0") and (va_apr7=0 OR va_apr7="0") and (va_apr8=0 OR va_apr8="0")and (va_sapr=0 OR va_sapr="0")  then
		data_inicio=""
		va_faltas=""
		end if
end if

%>
        <tr> 
          <td width="<%response.Write(width)%>" class="tabela" 
> 
            <%response.Write(no_materia)%>
          </td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
							if showprova="n" AND showapr="n" then
			  response.Write("&nbsp;")							
							else			  
			  response.Write(va_faltas&"&nbsp;")
			  end if
			  %>
            </div></td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write(va_apr1)
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="30" class="tabela" 
><div align="center"> 
              <%
							if showapr="s" then					
							response.Write(va_v_apr1)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write(va_apr2)
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write(va_v_apr2)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write(va_apr3)
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write(va_v_apr3)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write(va_apr4)
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write(va_v_apr4)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write(va_apr5)
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write(va_v_apr5)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write(va_apr6)
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write(va_v_apr6)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write(va_apr7)
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write(va_v_apr7)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write(va_apr8)
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write(va_v_apr8)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
			  				if showapr="s" then
	if va_sapr="&nbsp;" or isnull(va_sapr) then
								response.Write("&nbsp;")
	else						
	va_sapr = formatNumber(va_sapr,1)
	end if												
							response.Write(va_sapr&"&nbsp;")
							else
							response.Write("&nbsp;")
							end if
%>
            </div></td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
			if showprova="s" then
	if pr="" or isnull(pr) then
				response.Write("&nbsp;")	
	else							
	pr = formatNumber(pr,1)												
							response.Write(pr&"&nbsp;")
	end if		
			else
				response.Write("&nbsp;")
			end if			  
			  %>
            </div></td>
          <td width="30" class="tabela" 
> <div align="center"> 
              <%
			if showprova="s" then
			response.Write(va_me&"&nbsp;")			
			else			
			response.Write("&nbsp;")
			end if			
			%>
            </div></td>
<% if periodo_check=2 then%>
                          <td width="30" class="tabela" > 
						  <div align="center">
							<%
							if showprova="s" then
							response.Write(va_bon)							
							else
			response.Write("&nbsp;")												
							end if%></div>
							</td>
<%end if%>			
 <!--          <td width="200" class="tabela" 
> <div align="center"> 
              <%
							if showprova="n" AND showapr="n" then
			  response.Write("&nbsp;")							
							else			  
			  response.Write(data_inicio)
			  end if
			  %>
            </div></td>--> 
        </tr>
        <%check=check+1
RSprog.MOVENEXT
wend%>
      </table></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
<%if periodo_check=2 then%>							  
                          <td height="20" colspan="2" 
> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> F-Faltas , N-Nota Apr, P-Peso Apr SAPR–Média 
                              das Aprs, PR-Prova, MP–Média Período e ECE1-1ª Etapa Complementar de Estudos</font></div></td>							  
<%else%>							
                          <td height="20" colspan="2" 
> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> F-Faltas , N-Nota Apr, P-Peso Apr SAPR–Média 
                              das Aprs, PR-Prova e MP–Média Período</font></div></td>
<%end if%>  
  </tr>
  <tr> 
    <td colspan="2" class="linhaTopoL"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Sistema 
            Diretor - WEB FAMILIA</font> </td>
          <td><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Impresso 
              em 
              <%response.Write(data &" às "&horario)%>
              </font></div></td>
        </tr>
      </table>
      <div align="right"> </div></td>
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
response.redirect("../../inc/erro.asp")
end if
%>