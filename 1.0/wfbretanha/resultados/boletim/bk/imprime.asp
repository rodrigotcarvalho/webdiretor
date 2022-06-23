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
<!--#include file="../../inc/funcoes6.asp"-->

<%opt = REQUEST.QueryString("obr")
dados_opt= split(opt, "?" )
cod= dados_opt(0)
'periodo_check= dados_opt(1)

 	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0

 	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO_wf&";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1
	
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2		

		Set RSal = Server.CreateObject("ADODB.Recordset")
		SQLal = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod
		RSal.Open SQLal, CON1

nome_aluno= RSal("NO_Aluno")

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
if RS_tb.eof then
%>
                    <div align="center"><font class="style1"> <%response.Write("<br><br><br><br><br>Não existe Boletim para este aluno!")%></div>
                      </div>
                    <%
else

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
    <td width="122" height="15" bgcolor="#FFFFFF">
<div align="center"><img src="../../img/logo_preto.gif"> </div></td>
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
          <td width="68" height="12" bgcolor="#EEEEEE"> <div align="right"> <!-- <strong>Per&iacute;odo:</strong> --></div></td>
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
            <%response.Write(turma)%>
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
    <td colspan="2"> </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td colspan="2"><div align="right"> 

                    <%		Set RS_tb = Server.CreateObject("ADODB.Recordset")
		SQL_tb = "SELECT * FROM TB_Da_Aula WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso&"' AND CO_Etapa ='"& etapa&"' AND CO_Turma ='"& turma&"'"
		RS_tb.Open SQL_tb, CON2
if RS_tb.eof then
%>
                    <div align="center"> <%response.Write("<br><br><br><br><br>Não existe Boletim para este aluno!")%>
        </div>
                    <%
else
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
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL_WF = "SELECT * FROM TB_Autoriza_WF WHERE NU_Unidade="&unidade&" AND CO_Curso='"&curso&"' and CO_Etapa='"&etapa&"'"
		RS4.Open SQL_WF, CON
	
co_apr1=RS4("CO_apr1")
co_apr2=RS4("CO_apr2")
co_apr3=RS4("CO_apr3")
co_apr4=RS4("CO_apr4")
co_apr5=RS4("CO_apr5")
co_apr6=RS4("CO_apr6")
co_prova1=RS4("CO_prova1")
co_prova2=RS4("CO_prova2")
co_prova3=RS4("CO_prova3")
co_prova4=RS4("CO_prova4")
co_prova5=RS4("CO_prova5")
co_prova6=RS4("CO_prova6")	

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
		
		if co_apr5="D"then
		showapr5="n"
		else 
		showapr5="s"
		end if
		if co_prova5="D"then
		showprova5="n"
		else 
		showprova5="s"
		end if	

'if notaFIL="TB_NOTA_A" or notaFIL="TB_NOTA_B" or notaFIL="TB_NOTA_C" then			
	%>
                    
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="320" rowspan="2" class="tabela"> 
                        <div align="left"><strong>Disciplina</strong></div></td>
                      <td colspan="12" class="tabela"> <div align="center">Aproveitamento</div></td>
                      <td colspan="4" class="tabela"><div align="center">Freq&uuml;&ecirc;ncia 
                          (Faltas)</div></td>
                    </tr>
                    <tr> 
                      <td width="68" class="tabela"> <div align="center">BIM 
                          1</div></td>
                      <td width="68" class="tabela"> <div align="center">BIM 
                          2</div></td>
                      <td width="68" class="tabela"><div align="center">M&eacute;dia 
                          Sem 1</div></td>
                      <td width="68" class="tabela"><div align="center">Recup 
                          Sem</div></td>
                      <td width="68" class="tabela"><div align="center">M&eacute;dia 
                          Sem 2</div></td>
                      <td width="68" class="tabela"> <div align="center">BIM 
                          3</div></td>
                      <td width="68" class="tabela"> <div align="center">BIM 
                          4</div></td>
                      <td width="68" class="tabela"><div align="center">M&eacute;dia 
                          Sem 3</div></td>
                      <td width="68" class="tabela"> <div align="center">M&eacute;dia 
                          Anual</div></td>
                      <td width="68" class="tabela"> <div align="center">Recup 
                          Final </div></td>
                      <td width="68" class="tabela"><div align="center">M&eacute;dia 
                        Final</div></td>
                      <td width="68" class="tabela"><div align="center">Result</div></td>
                      <td width="68" class="tabela"><div align="center">BIM 
                          1</div></td>
                      <td width="68" class="tabela"> <div align="center">BIM 
                          2</div></td>
                      <td width="68" class="tabela"> <div align="center">BIM 
                          3</div></td>
                      <td width="68" class="tabela"> <div align="center">BIM 
                          4</div></td>
                    </tr>
                    <%
			rec_lancado="sim"
		
			Set RSprog = Server.CreateObject("ADODB.Recordset")
			SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
			RSprog.Open SQLprog, CON0
		
			check=2
				
			while not RSprog.EOF
			
					dividendo1=0
					divisor1=0
					dividendo2=0
					divisor2=0
					dividendo3=0
					divisor3=0
					dividendo4=0
					divisor4=0
					dividendo_ma=0
					divisor_ma=0
					dividendo5=0
					divisor5=0
					dividendo_mf=0
					divisor_mf=0
					dividendo6=0
					divisor6=0
					dividendo_rec=0
					divisor_rec=0
					res1="&nbsp;"
					res2="&nbsp;"
					res3="&nbsp;"
					m2="&nbsp;"
					m3="&nbsp;"
															
			verifica="ok"
			
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
																

				'response.Write(materia&" "&mae&" "&in_co&"<br>")
				
				if mae=TRUE and fil=FALSE and in_co=FALSE and isnull(nu_peso) then
				f1_ac=0
				f2_ac=0
				f3_ac=0
				f4_ac=0
				mb1_ac=0
				mb2_ac=0
				mb3_ac=0
				mb4_ac=0
				divisor_mb1_ac=0
				divisor_mb2_ac=0
				divisor_mb3_ac=0
				divisor_mb4_ac=0								
				ms1_ac=0
				ms2_ac=0
				ms3_ac=0
				ms35_ac=0
				ms36_ac=0								
				ma_ac=0
				peso_ac=0
				ordem2=ordem+1
				tentativas=0
								
				
			
						Set RS1a = Server.CreateObject("ADODB.Recordset")
						SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
						RS1a.Open SQL1a, CON0
							
						no_materia=RS1a("NO_Materia")
					
						if check mod 2 =0 then
						cor = "tabela" 
						cor2 = "tabela" 				
						else 
						cor ="tabela"
						cor2 = "tabela" 
						end if
					
							
							
						Set CON_N = Server.CreateObject("ADODB.Connection") 
						ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
						CON_N.Open ABRIRn
					

						for periodofil=1 to 5
											
								
							Set RSnFIL = Server.CreateObject("ADODB.Recordset")
							Set RS3 = Server.CreateObject("ADODB.Recordset")
							SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' AND NU_Periodo="&periodofil
							Set RS3 = CON_N.Execute(SQL_N)
						
							if RS3.EOF then
								if periodofil=1 then
								f1="&nbsp;"
								va_m31="&nbsp;"
								elseif periodofil=2 then
								f2="&nbsp;"
								va_m32="&nbsp;"
								va_rec_sem="&nbsp;"
								elseif periodofil=3 then
								f3="&nbsp;"
								va_m33="&nbsp;"
								elseif periodofil=4 then
								f4="&nbsp;"
								va_m34="&nbsp;"
								elseif periodofil=5 then
								va_m35="&nbsp;"
								elseif periodofil=6 then
								va_m36="&nbsp;"
								end if	
							else
								if periodofil=1 then
								f1=RS3("NU_Faltas")
								va_m31=RS3("VA_Media3")
								elseif periodofil=2 then
								f2=RS3("NU_Faltas")
								va_m32=RS3("VA_Media3")
								va_rec_sem=RS3("VA_Rec")
								elseif periodofil=3 then
								f3=RS3("NU_Faltas")
								va_m33=RS3("VA_Media3")
								elseif periodofil=4 then
								f4=RS3("NU_Faltas")
								va_m34=RS3("VA_Media3")
								elseif periodofil=5 then
								va_m35=RS3("VA_Media3")
								elseif periodofil=6 then
								va_m36=RS3("VA_Media3")
								end if
							end if
						NEXT
							
							if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
							dividendo1=0
							divisor1=0
							else
							dividendo1=va_m31
							divisor1=1
							end if	
							
							if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
							dividendo2=0
							divisor2=0
							else
							dividendo2=va_m32
							divisor2=1
							end if
							
							if isnull(va_rec_sem) or va_rec_sem="&nbsp;"  or va_rec_sem="" then
							dividendorec=0
							divisorrec=0
							else
							dividendorec=va_rec_sem
							divisorrec=1
							end if
							
							if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
							dividendo3=0
							divisor3=0
							else
							dividendo3=va_m33
							divisor3=1
							end if
							
							if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
							dividendo4=0
							divisor4=0
							else
							dividendo4=va_m34
							divisor4=1
							end if														
							if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
							nota_aux_m2_1="&nbsp;"
							'dividendo5=0
							'divisor5=0
							else
							nota_aux_m2_1=va_m35
							'dividendo5=va_m35
							'divisor5=1
							end if

'							showapr1="s"
'							showprova1="s"
'							showapr2="s"
'							showprova2="s"
'							showapr3="s"
'							showprova3="s"
'							showapr4="s"
'							showprova4="s"

							dividendo_ms1=dividendo1+dividendo2
							divisor_ms1=divisor1+divisor2
													
							if divisor_ms1<2 then
							ms1="&nbsp;"
							dividendoms1=0
							divisorms1=0
							else
							ms1=dividendo_ms1/divisor_ms1
								decimo = ms1 - Int(ms1)
								If decimo >= 0.5 Then
									nota_arredondada = Int(ms1) + 1
									ms1=nota_arredondada
								'elseIf decimo >= 0.25 Then
								'	nota_arredondada = Int(ms1) + 0.5
								'	ms1=nota_arredondada
								else
									nota_arredondada = Int(ms1)
									ms1=nota_arredondada											
								End If			
							ms1 = formatNumber(ms1,0)
							dividendoms1=ms1
							divisorms1=1
							end if
							

							
							if divisorrec=0 then
								ms2=ms1
								if ms2="&nbsp;" then
									dividendoms2=0
									divisorms2=0
									dividendo_anual_ms2=0
									divisor_anual_ms2=0
								else
									dividendoms2=ms2
									divisorms2=1						
									dividendo_anual_ms2=ms2
									divisor_anual_ms2=1
								end if
							else
								dividendo_ms2=dividendoms1+dividendorec
								divisor_ms2=divisorms1+divisorrec
																						
								ms2=dividendo_ms2/divisor_ms2
'response.Write(ms2&"+"&dividendoms1&"+"&divisor_ms2)
ms2=ms2*1	
ms1=ms1*1							
								if ms2<ms1 then
									ms2=ms1								
								end if
									decimo = ms2 - Int(ms2)
									If decimo >= 0.5 Then
										nota_arredondada = Int(ms2) + 1
										ms2=nota_arredondada
									'elseIf decimo >= 0.25 Then
									'	nota_arredondada = Int(ms2) + 0.5
									'	ms2=nota_arredondada
									else
										nota_arredondada = Int(ms2)
										ms2=nota_arredondada											
									End If
								ms2 = formatNumber(ms2,0)																
								dividendo_anual_ms2=ms2
								divisor_anual_ms2=1
							end if
							
							dividendo_ms3=dividendo3+dividendo4
							divisor_ms3=divisor3+divisor4
							
							if divisor_ms3<2 then
							ms3="&nbsp;"
							dividendo_anual_ms3=0
							divisor_anual_ms3=0					
							else
							ms3=dividendo_ms3/divisor_ms3
								decimo = ms3 - Int(ms3)
								If decimo >= 0.5 Then
									nota_arredondada = Int(ms3) + 1
									ms3=nota_arredondada
								'elseIf decimo >= 0.25 Then
								'	nota_arredondada = Int(ms3) + 0.5
								'	ms3=nota_arredondada
								else
									nota_arredondada = Int(ms3)
									ms3=nota_arredondada											
								End If
								ms3 = formatNumber(ms3,0)								
							dividendo_anual_ms3=ms3
							divisor_anual_ms3=1						
							end if					
								dividendo_anual_ms2=dividendo_anual_ms2*1
								dividendo_anual_ms3=dividendo_anual_ms3*1
								divisor_anual_ms2=divisor_anual_ms2*1
								divisor_anual_ms3=divisor_anual_ms3*1		
							dividendo_ma=dividendo_anual_ms2+dividendo_anual_ms3
							divisor_ma=divisor_anual_ms2+divisor_anual_ms3
							
							'response.Write(dividendo_ma&"<<")
							
							if divisor_ma<2 then
							ma="&nbsp;"
							else
							ma=dividendo_ma/divisor_ma
							end if
							
							if ma="&nbsp;" then
							else
							'mf=mf/10
								decimo = ma - Int(ma)
		
									If decimo >= 0.5 Then
										nota_arredondada = Int(ma) + 1
										ma=nota_arredondada
									'elseIf decimo >= 0.25 Then
									'	nota_arredondada = Int(ma) + 0.5
									'	ma=nota_arredondada								
									Else
										nota_arredondada = Int(ma)
										ma=nota_arredondada					
									End If
									
								ma = formatNumber(ma,0)
							end if

		
							if ma="&nbsp;" then
								media_final="&nbsp;"
								resultado_final="&nbsp;"							
							else	
							
								if nota_aux_m2_1="&nbsp;" then
									tipo_calculo="anual"
								else
									tipo_calculo="final"
								end if	
								if isnumeric(ma) then	
									ma=ma/10
								end if	
								
								if isnumeric(nota_aux_m2_1) then	
									nota_aux_m2_1=nota_aux_m2_1/10
								end if		
	
								resultado=regra_aprovacao(curso,etapa,ma,nota_aux_m2_1,"&nbsp;","&nbsp;","&nbsp;",tipo_calculo)
								resultado_aluno = split(resultado,"#!#")
								
								media_final=resultado_aluno(0)
								if isnumeric(media_final) then
									media_final=media_final*10
								end if								
								resultado_final=resultado_aluno(1)								
							end if
							
							
'response.Write(Session("resultado_1") &" - "& Session("resultado_2") &" - "& Session("resultado_3")&"<<")
			%>
                    <tr> 
                      <td width="320" class="<%response.Write(cor)%>"> 
                        <%response.Write(no_materia)
						  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
						  %>
                      </td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                          <%
							if showapr1="s" and showprova1="s" then																	
							response.Write(va_m31)
							else
							response.Write("&nbsp;")							
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                          <%
							if showapr2="s" and showprova2="s" then												
							response.Write(va_m32)
							else
							response.Write("&nbsp;")													
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
                          <%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" then					
							response.Write(ms1)
							else
							response.Write("&nbsp;")							
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
                          <%
							if showapr2="s" and showprova2="s" then				
							response.Write(va_rec_sem)
							else
							response.Write("&nbsp;")							
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
                          <%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" then					
							response.Write(ms2)
							else
							response.Write("&nbsp;")							
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                          <%
							if showapr3="s" and showprova3="s" then				
							response.Write(va_m33)
							else
							response.Write("&nbsp;")							
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                          <%
							if showapr4="s" and showprova4="s" then					
							response.Write(va_m34)
							else
							response.Write("&nbsp;")							
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                          <%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then					
							response.Write(ms3)
							else
							response.Write("&nbsp;")							
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                          <%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then	
							response.Write(ma)
							else
							response.Write("&nbsp;")
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                          <%
							if showapr5="s" and showprova5="s" then												
							response.Write(va_m35)
							else
							response.Write("&nbsp;")							
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">
                        <%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then
								if (showapr5="s" and showprova5="s") or (showapr5="n" and showprova5="n" and nota_aux_m2_1="&nbsp;") then										
									response.Write(media_final)
								else
									response.Write("&nbsp;")
								end if
							end if	
							%>
                      </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                        <%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then
								if (showapr5="s" and showprova5="s") or (showapr5="n" and showprova5="n" and nota_aux_m2_1="&nbsp;") then										
									response.Write(resultado_final)
								else
									response.Write("&nbsp;")
								end if
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
                          <%
							response.Write("&nbsp;"&f1)
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
                          <%
							response.Write("&nbsp;"&f2)
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
                          <%
							response.Write("&nbsp;"&f3)

							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
                          <%
							response.Write("&nbsp;"&f4)
							%>
                        </div></td>
                    </tr>
                    <%
					check=check+1
					else%>
                    <tr> 
                      <td width="320" class="<%response.Write(cor)%>"> 
                        <%response.Write(no_materia)
						  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
						  %>
                      </td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>">&nbsp;</td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
                        </div></td>
                    </tr>
                    <%
					check=check+1
					end if	
			RSprog.MOVENEXT
			wend
		
			%>
                    <tr valign="bottom"> 
                      <td height="20"> <div align="right"> </div></td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td>&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                    </tr>
                  </table>
<%
end if					
					%>
</div></td>
  </tr>
  <tr> 
    <td colspan="2" class="linhaTopoL">
<div align="right"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Sistema 
              Diretor - WEB FAM&Iacute;LIA</font></td>
            <td><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Impresso 
                em 
                <%response.Write(data &" às "&horario)%>
                </font></div></td>
          </tr>
        </table>
        
      </div></td>
  </tr>
</table>
<%end if%>
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