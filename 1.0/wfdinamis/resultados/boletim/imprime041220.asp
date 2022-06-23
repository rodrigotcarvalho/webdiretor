<%On Error Resume Next%>
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


'		Set RSPER = Server.CreateObject("ADODB.Recordset")
'		SQLPER = "SELECT * FROM TB_Periodo where NU_Periodo="&periodo_check
'		RSPER.Open SQLPER, CON0

'		no_periodo=RSPER("NO_Periodo")


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

	%>



<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('window.print()')"> 
<br>
<table width="950" border="0" align="center" cellspacing="0" class="tb_corpo"
>
  <tr> 
    <td width="122" height="15" bgcolor="#FFFFFF"> <div align="center"><img src="../../img/logo_preto.gif"> 
      </div></td>
    <td width="828" bgcolor="#FFFFFF"><table width="100%" border="0" align="right" cellspacing="0">
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
          <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>BOLETIM 
            N&Atilde;O OFICIAL</strong></font></td>
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
          <td width="265" rowspan="2" class="tabela" 
> <div align="left"><strong>Disciplina</strong></div></td>
          <td colspan="5" class="tabela" 
> <div align="center">TRIMESTRE 1</div></td>
          <td colspan="5" class="tabela" 
><div align="center">TRIMESTRE 2 </div></td>
          <td colspan="5" class="tabela" 
><div align="center">TRIMESTRE 3 </div></td>
          <td width="35" rowspan="2" class="tabela" 
> <div align="center">RA</div></td>
          <td colspan="4" class="tabela" 
><div align="center">ETAPA COMPLEMENTAR</div></td>
          <td width="35" rowspan="2" class="tabela" 
> <div align="center">RV</div></td>
        </tr>
        <tr> 
          <td width="35" class="tabela" 
><div align="center">SAPR</div></td>
          <td width="35" class="tabela" 
> <div align="center">PR</div></td>
          <td width="35" class="tabela" 
> <div align="center">MP</div></td>
          <td width="35" class="tabela" 
> <div align="center">MC</div></td>
          <td width="35" class="tabela" 
> <div align="center">F</div></td>
          <td class="tabela" 
><div align="center">SAPR</div></td>
          <td class="tabela" 
> <div align="center">PR</div></td>
          <td class="tabela" 
> <div align="center">MP</div></td>
          <td class="tabela" 
> <div align="center">MC*</div></td>
          <td class="tabela" 
> <div align="center">F</div></td>
          <td class="tabela" 
><div align="center">SAPR</div></td>
          <td class="tabela" 
> <div align="center">PR</div></td>
          <td class="tabela" 
> <div align="center">MP</div></td>
          <td class="tabela" 
> <div align="center">MC</div></td>
          <td class="tabela" 
> <div align="center">F</div></td>
          <td class="tabela" 
><div align="center">SAPR</div></td>
          <td class="tabela" 
> <div align="center">PR</div></td>
          <td class="tabela" 
> <div align="center">MP</div></td>
          <td class="tabela" 
> <div align="center">MC</div></td>
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
va_sapr1="&nbsp;"
va_pr1="&nbsp;"
va_te1="&nbsp;"
pr1="&nbsp;"
va_me1="&nbsp;"
va_mc1="&nbsp;"
va_faltas1="&nbsp;"
va_sapr2="&nbsp;"
va_pr2="&nbsp;"
va_te2="&nbsp;"
pr2="&nbsp;"
va_me2="&nbsp;"
va_mc2="&nbsp;"
va_faltas2="&nbsp;"
va_sapr3="&nbsp;"
va_pr3="&nbsp;"
va_te3="&nbsp;"
pr3="&nbsp;"
va_me3="&nbsp;"
va_mc3="&nbsp;"
va_faltas3="&nbsp;"
va_sapr4="&nbsp;"
va_pr4="&nbsp;"
va_me4="&nbsp;"
va_mc4="&nbsp;"
	
else
va_sapr1=RS3("VA_Sapr1")
va_pr1=RS3("VA_Pr1")
va_te1=RS3("VA_Te1")
va_me1=RS3("VA_Me1")
va_mc1=RS3("VA_Mc1")
va_faltas1=RS3("NU_Faltas_P1")
va_pr1=va_pr1*1
va_te1=va_te1*1
pr1=va_pr1+va_te1

va_sapr2=RS3("VA_Sapr2")
va_pr2=RS3("VA_Pr2")
va_te2=RS3("VA_Te2")
va_me2=RS3("VA_Me2")
va_mc2=RS3("VA_Mc2")
va_faltas2=RS3("NU_Faltas_P2")

va_pr2=va_pr2*1
va_te2=va_te2*1
pr2=va_pr2+va_te2

va_sapr3=RS3("VA_Sapr3")
va_pr3=RS3("VA_Pr3")
va_te3=RS3("VA_Te3")
va_me3=RS3("VA_Me3")
va_mc3=RS3("VA_Mc3")
va_faltas3=RS3("NU_Faltas_P3")

va_pr3=va_pr3*1
va_te3=va_te3*1
pr3=va_pr3+va_te3

va_sapr4=RS3("VA_Sapr_EC")
va_pr4=RS3("VA_Pr4")
va_me4=RS3("VA_Me_EC")
va_mc4=RS3("VA_Mfinal")

pr4=va_pr

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
		
'if periodo_check=1 then		
		if co_apr1="D"then
		showapr1="n"
		else 
		showapr1="s"
		end if
		if co_prova1="D"then
		showprova1="n"
		else 
		showprova1="s"
		end if
'elseif periodo_check=2 then		
		if co_apr2="D"then
		showapr2="n"
		else 
		showapr2="s"
		end if
		if co_prova2="D"then
		showprova2="n"
		else 
		showprova2="s"
		end if					
'elseif periodo_check=3 then		
		if co_apr3="D"then
		showapr3="n"
		else 
		showapr3="s"
		end if
		if co_prova3="D"then
		showprova3="n"
		else 
		showprova3="s"
		end if
'elseif periodo_check=4 then		
		if co_apr4="D"then
		showapr4="n"
		else 
		showapr4="s"
		end if
		if co_prova4="D"then
		showprova4="n"
		else 
		showprova4="s"
		end if
'end if											
		
		
				
end if

if va_me1="&nbsp;" or isnull(va_me1)then
else
va_me1=va_me1/10
'	decimo = va_me1 - Int(va_me1)
'		If decimo >= 0.5 Then
'			nota_arredondada = Int(va_me1) + 1
'			va_me1=nota_arredondada
'		Else
'			nota_arredondada = Int(va_me1)
'			va_me1=nota_arredondada					
'		End If
	va_me1 = formatNumber(va_me1,1)
end if	
	
if va_mc1="&nbsp;" or isnull(va_mc1)then
else	
va_mc1=va_mc1/10
'	decimo = va_mc1 - Int(va_mc1)
'		If decimo >= 0.5 Then
'			nota_arredondada = Int(va_mc1) + 1
'			va_mc1=nota_arredondada
'		Else
'			nota_arredondada = Int(va_mc1)
'			va_mc1=nota_arredondada					
'		End If
	va_mc1 = formatNumber(va_mc1,1)
end if	
if va_me2="&nbsp;" or isnull(va_me2)then
else
va_me2=va_me2/10
'	decimo = va_me2 - Int(va_me2)
'		If decimo >= 0.5 Then
'			nota_arredondada = Int(va_me2) + 1
'			va_me2=nota_arredondada
'		Else
'			nota_arredondada = Int(va_me2)
'			va_me2=nota_arredondada					
'		End If
	va_me2 = formatNumber(va_me2,1)
end if	
	
if va_mc2="&nbsp;" or isnull(va_mc2)then
else		
va_mc2=va_mc2/10
'	decimo = va_mc2 - Int(va_mc2)
'		If decimo >= 0.5 Then
'			nota_arredondada = Int(va_mc2) + 1
'			va_mc2=nota_arredondada
'		Else
'			nota_arredondada = Int(va_mc2)
'			va_mc2=nota_arredondada					
'		End If
	va_mc2 = formatNumber(va_mc2,1)		
end if	
if va_me3="&nbsp;" or isnull(va_me3)then
else
va_me3=va_me3/10
'	decimo = va_me3 - Int(va_me3)
'		If decimo >= 0.5 Then
'			nota_arredondada = Int(va_me3) + 1
'			va_me3=nota_arredondada
'		Else
'			nota_arredondada = Int(va_me3)
'			va_me3=nota_arredondada					
'		End If
	va_me3 = formatNumber(va_me3,1)
end if	
	
if va_mc3="&nbsp;" or isnull(va_mc3)then
else		
va_mc3=va_mc3/10
'	decimo = va_mc3 - Int(va_mc3)
'		If decimo >= 0.5 Then
'			nota_arredondada = Int(va_mc3) + 1
'			va_mc3=nota_arredondada
'		Else
'			nota_arredondada = Int(va_mc3)
'			va_mc3=nota_arredondada					
'		End If
	va_mc3 = formatNumber(va_mc3,1)
end if	
if va_me4="&nbsp;" or isnull(va_me4)then
else	
va_me4=va_me4/10
'	decimo = va_me4 - Int(va_me4)
'		If decimo >= 0.5 Then
'			nota_arredondada = Int(va_me4) + 1
'			va_me4=nota_arredondada
'		Else
'			nota_arredondada = Int(va_me4)
'			va_me4=nota_arredondada					
'		End If
	va_me4 = formatNumber(va_me4,1)
end if	
	
if va_mc4="&nbsp;" or isnull(va_mc4)then
else		
va_mc4=va_mc4/10
'	decimo = va_mc4 - Int(va_mc4)
'		If decimo >= 0.5 Then
'			nota_arredondada = Int(va_mc4) + 1
'			va_mc4=nota_arredondada
'		Else
'			nota_arredondada = Int(va_mc4)
'			va_mc4=nota_arredondada					
'		End If
	va_mc4 = formatNumber(va_mc4,1)	
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
          <td width="265" class="tabela" 
> 
            <%response.Write(no_materia)%>
          </td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
							if showapr1="s" and showprova1="s" then
	if va_sapr1="&nbsp;" or isnull(va_sapr1) then
	else						
	va_sapr1 = formatNumber(va_sapr1,1)
	end if												
							response.Write("&nbsp;"&va_sapr1)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
		if showapr1="s" and showprova1="s" then
			if pr1="&nbsp;" or isnull(pr1) OR pr1="" then
				response.Write("&nbsp;")
			else							
				pr1 = formatNumber(pr1,1)												
				response.Write("&nbsp;"&pr1)
			end if							
		else
			response.Write("&nbsp;")							
		end if
							%>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
							if showapr1="s" and showprova1="s" then					
							response.Write("&nbsp;"&va_me1)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
							if showapr1="s" and showprova1="s" then					
							response.Write("&nbsp;"&va_mc1)
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
							if showapr1="s" and showprova1="s" then
							response.Write("&nbsp;"&va_faltas1)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
							if showapr2="s" and showprova2="s" then
	if va_sapr2="&nbsp;" or isnull(va_sapr2) then
								response.Write("&nbsp;")
	else							
	va_sapr2 = formatNumber(va_sapr2,1)												
							response.Write("&nbsp;"&va_sapr2)
	end if							
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
							if showapr2="s" and showprova2="s" then
	if pr2="&nbsp;" or isnull(pr2) then
								response.Write("&nbsp;")
	else							
	pr2 = formatNumber(pr2,1)												
							response.Write("&nbsp;"&pr2)
	end if							
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
							if showapr2="s" and showprova2="s" then					
							response.Write("&nbsp;"&va_me2)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
							if showapr2="s" and showprova2="s" then					
							response.Write("&nbsp;"&va_mc2)
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
							if showapr2="s" and showprova2="s" then
							response.Write("&nbsp;"&va_faltas2)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
							if showapr3="s" and showprova3="s" then
	if va_sapr3="&nbsp;" or isnull(va_sapr3) then
								response.Write("&nbsp;")
	else							
	va_sapr3 = formatNumber(va_sapr3,1)												
							response.Write("&nbsp;"&va_sapr3)
	end if		
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
							if showapr3="s" and showprova3="s" then
	if pr3="&nbsp;" or isnull(pr3) then
								response.Write("&nbsp;")
	else							
	pr3 = formatNumber(pr3,1)												
							response.Write("&nbsp;"&pr3)
	end if								

							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
							if showapr3="s" and showprova3="s" then					
							response.Write("&nbsp;"&va_me3)
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
							if showapr3="s" and showprova3="s" then					
							response.Write("&nbsp;"&va_mc3)
							else
							response.Write("&nbsp;")
							end if
							%>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
							if showapr3="s" and showprova3="s" then
							response.Write("&nbsp;"&va_faltas3)
							else
							response.Write("&nbsp;")							
							end if
							%>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
					if showapr3="s" and showprova3="s" then	
						if va_mc3="&nbsp;" or isnull(va_mc3) then
							response.Write(va_mc3)
						else
							if va_mc3 < 7 then					
							response.Write("ECE")
							resultado1="ece"
							else
							response.Write("APR")
							resultado1="apr"							
							end if
						end if	
					else
					response.Write("&nbsp;")							
					end if							
							%>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
			  if resultado1="ece" then 			  
							if showapr4="s" and showprova4="s" then
	if va_sapr4="&nbsp;" or isnull(va_sapr4) then
								response.Write("&nbsp;")
	else							
	va_sapr4 = formatNumber(va_sapr4,1)
							response.Write("&nbsp;"&va_sapr4)
	end if													

							else
							response.Write("&nbsp;")
							end if
			else
				response.Write("&nbsp;")
			end if
%>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
			  if resultado1="ece" then 
			if showapr4="s" and showprova4="s" then
	if pr4="&nbsp;" or isnull(pr4) OR pr4="" then
								response.Write("&nbsp;")
	else							
	pr4 = formatNumber(pr4,1)								
				response.Write("&nbsp;"&pr4)
	end if								
			

			else
				response.Write("&nbsp;")
			end if
			else
				response.Write("&nbsp;")
			end if			  
			  %>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
			  if resultado1="ece" then 			  
			if showapr4="s" and showprova4="s" then					
			response.Write("&nbsp;"&va_me4)			
			else			
			response.Write("&nbsp;")
			end if
			else
				response.Write("&nbsp;")
			end if			
			%>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
			  if resultado1="ece" then 			  
			if showapr4="s" and showprova4="s"then					
				response.Write("&nbsp;"&va_mc4)							
			else			  
			response.Write("&nbsp;")
			end if
			else
				response.Write("&nbsp;")
			end if
			  %>
            </div></td>
          <td width="35" class="tabela" 
> <div align="center"> 
              <%
			  if resultado1="ece" then 			  
					if showapr4="s" and showprova4="s" then	
						if va_mc4="&nbsp;" or isnull(va_mc4) then
							response.Write("&nbsp;")
						else
							if va_mc4 < 5 then					
							response.Write("REP")
							else
							response.Write("APR")
							end if
						end if	
					else							
					response.Write("&nbsp;")							
					end if
			else
				response.Write("&nbsp;")
			end if					
							%>
            </div></td>
        </tr>
        <%check=check+1
RSprog.MOVENEXT
wend%>
      </table></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td colspan="2"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Sapr–Média 
        das Aprs, PR-Prova, MP-Média do Período, MC-Média Acumulada, F-Faltas, 
        RA-Resultado Anual, RF-Resultado Final</font></div></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td colspan="2"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">* Esta nota est&aacute; sujeita a altera&ccedil;&otilde;es pela 1&ordf; 
                          Etapa Complementar de Estudos (Vide o Boletim de Avalia&ccedil;&otilde;es 
                          do 2&ordm; Trimestre).</font></div></td>
  </tr>
  <tr> 
    <td colspan="2" class="linhaTopoL"> <div align="right"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Sistema 
              Diretor - WEB FAMILIA</font> </td>
            <td><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Impresso 
                em 
                <%response.Write(data &" às "&horario)%>
                </font></div></td>
          </tr>
        </table>
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
response.redirect("../../inc/erro.asp")
end if
%>