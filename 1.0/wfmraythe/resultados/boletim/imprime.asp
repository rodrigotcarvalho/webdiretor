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
'periodo_check= dados_opt(1)

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
          <td width="68" height="12" bgcolor="#EEEEEE"> <div align="right"> <strong><!-- Per&iacute;odo: --></strong></div></td>
          <td width="198"><!--  
            <%response.Write(no_periodo)%>
 -->          </td>
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
    <td colspan="2">  
<%		Set RS_tb = Server.CreateObject("ADODB.Recordset")
		SQL_tb = "SELECT * FROM TB_Da_Aula WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso&"' AND CO_Etapa ='"& etapa&"' AND CO_Turma ='"& turma&"'"
		RS_tb.Open SQL_tb, CON2
if RS_tb.eof then
%>
                  <div align="center"><font class="form_corpo">
                    <%response.Write("<br><br><br><br><br>Não existe Boletim para este aluno!")%>
                    </font> </div>
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
	%><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="339" rowspan="2" class="tabela"> 
                        <div align="left"><strong>Disciplina</strong></div></td>
                      <td colspan="7" class="tabela"> <div align="center"></div>
                        <div align="center">Aproveitamento</div></td>
                      <td width="1"> <div align="center"></div></td>
                      <td colspan="4" class="tabela"><div align="center">Freq&uuml;&ecirc;ncia 
                          (Faltas):</div></td>
                    </tr>
                    <tr> 
                      <td width="60" class="tabela"> 
                        <div align="center">PA1</div></td>
                      <td width="60" class="tabela"> 
                        <div align="center">PA2</div></td>
                      <td width="60" class="tabela"> 
                        <div align="center">PA3</div></td>
                      <td width="60" class="tabela"> 
                        <div align="center">TOTAL</div></td>
                      <td width="60" class="tabela"> 
                        <div align="center">4&ordf; 
                          aval<br>
                          p.2</div></td>
                      <td width="60" class="tabela"> 
                        <div align="center">TOTAL</div></td>
                      <td width="60" class="tabela"> 
                        <div align="center">M&eacute;dia 
                          Final</div></td>
                      <td width="1">&nbsp;</td>
                      <td width="60" class="tabela"> 
                        <div align="center">PA1</div></td>
                      <td width="60" class="tabela"> 
                        <div align="center">PA2</div></td>
                      <td width="60" class="tabela"> 
                        <div align="center">PA3</div></td>
                      <td width="60" class="tabela"> 
                        <div align="center">TOTAL</div></td>
                    </tr>
                    <%
			rec_lancado="sim"
		
			Set RSprog = Server.CreateObject("ADODB.Recordset")
			SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' AND IN_MAE=TRUE order by NU_Ordem_Boletim "
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
				cor = "tabela" 
				cor2 = "tabela" 				
				else 
				cor ="tabela"
				cor2 = "tabela" 
				end if
			
					
					
				Set CON_N = Server.CreateObject("ADODB.Connection") 
				ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
				CON_N.Open ABRIRn
			
				if (mae=TRUE and fil=FALSE and in_co=FALSE and isnull(nu_peso)) or (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(nu_peso)) or (mae=TRUE and fil=TRUE and in_co=FALSE) then				
			
					for periodofil=1 to 4
					
						Set RS3 = Server.CreateObject("ADODB.Recordset")
						SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' AND CO_Materia_Principal = '"& materia &"' AND NU_Periodo="&periodofil
						Set RS3 = CON_N.Execute(SQL_N)
					
					
					
						if RS3.EOF then
							if periodofil=1 then
							f1="&nbsp;"
							va_m31="&nbsp;"
							elseif periodofil=2 then
							f2="&nbsp;"
							va_m32="&nbsp;"
							elseif periodofil=3 then
							f3="&nbsp;"
							va_m33="&nbsp;"
							elseif periodofil=4 then
							f4="&nbsp;"
							va_m34="&nbsp;"
							end if	
						else
							if periodofil=1 then
								f1=RS3("NU_Faltas")						
								va_m31=RS3("VA_Media3")
							elseif periodofil=2 then
								f2=RS3("NU_Faltas")						
								va_m32=RS3("VA_Media3")
							elseif periodofil=3 then
								f3=RS3("NU_Faltas")						
								va_m33=RS3("VA_Media3")
							elseif periodofil=4 then
								f4=RS3("NU_Faltas")
								va_m34=RS3("VA_Media3")
							end if
						end if
					NEXT		
				else
					Set RS1a = Server.CreateObject("ADODB.Recordset")
					SQL1a = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& materia &"' order by NU_Ordem_Boletim"
					RS1a.Open SQL1a, CON0
						
					if RS1a.EOF then
					else
					co_materia_fil_check=1 
					peso_acumula=0
					va_m31_acumula=0
					va_m32_acumula=0
					va_m33_acumula=0
					va_m34_acumula=0
					sem_nota1="n"
					sem_nota2="n"
					sem_nota3="n"
					sem_nota4="n"					
						while not RS1a.EOF
							co_mat_fil= RS1a("CO_Materia")
							
							Set RSp2 = Server.CreateObject("ADODB.Recordset")
							SQLp2 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia = '"& co_mat_fil &"' order by NU_Ordem_Boletim"
							RSp2.Open SQLp2, CON0	
													
							nu_peso_fil=RSp2("NU_Peso")	
										
							peso_acumula=peso_acumula+nu_peso_fil
							
							for periodofil=1 to 4
														
								Set RS3 = Server.CreateObject("ADODB.Recordset")
								SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& co_mat_fil &"' AND CO_Materia_Principal = '"& materia &"' AND NU_Periodo="&periodofil
								Set RS3 = CON_N.Execute(SQL_N)						
						

								if RS3.EOF then
									if periodofil=1 then
										f1="&nbsp;"
										va_m31_temp="&nbsp;"
									elseif periodofil=2 then
										f2="&nbsp;"
										va_m32_temp="&nbsp;"
									elseif periodofil=3 then
										f3="&nbsp;"
										va_m33_temp="&nbsp;"
									elseif periodofil=4 then
										f4="&nbsp;"
										va_m34_temp="&nbsp;"
									end if	
								else
									if periodofil=1 then
										f1=RS3("NU_Faltas")						
										va_m31_temp=RS3("VA_Media3")
									elseif periodofil=2 then
										f2=RS3("NU_Faltas")						
										va_m32_temp=RS3("VA_Media3")
									elseif periodofil=3 then
										f3=RS3("NU_Faltas")						
										va_m33_temp=RS3("VA_Media3")
									elseif periodofil=4 then
										f4=RS3("NU_Faltas")
										va_m34_temp=RS3("VA_Media3")
									end if
								end if
							next	
							if isnull(va_m31_temp) or va_m31_temp="&nbsp;"  or va_m31_temp="" then
								sem_nota1="s"
							else
								va_m31_acumula=va_m31_acumula+va_m31_temp								
							end if	
							
							if isnull(va_m32_temp) or va_m32_temp="&nbsp;" or va_m32_temp="" then
								sem_nota2="s"
							else
								va_m32_acumula=va_m32_acumula+va_m32_temp	
							end if
							
							if isnull(va_m33_temp) or va_m33_temp="&nbsp;" or va_m33_temp="" then
								sem_nota3="s"
							else
								va_m33_acumula=va_m33_acumula+va_m33_temp	
							end if
							
							if isnull(va_m34_temp) or va_m34_temp="&nbsp;" or va_m34_temp="" then
								sem_nota4="s"
							else
								va_m34_acumula=va_m34_acumula+va_m34_temp	
							end if
						RS1a.MOVENEXT
						wend
						if sem_nota1="s" then
							va_m31="&nbsp;"
						else	
							va_m31=va_m31_acumula/peso_acumula
							va_m31=va_m31*10
								decimo = va_m31 - Int(va_m31)
								If decimo >= 0.5 Then
									nota_arredondada = Int(va_m31) + 1
									va_m31=nota_arredondada
								else
									nota_arredondada = Int(va_m31)
									va_m31=nota_arredondada											
								End If
							va_m31=va_m31/10	
							va_m31 = formatNumber(va_m31,1)									
						end if	
						
						if sem_nota2="s" then
							va_m32="&nbsp;"
						else	
							va_m32=va_m32_acumula/peso_acumula
							va_m32=va_m32*10
								decimo = va_m32 - Int(va_m32)
								If decimo >= 0.5 Then
									nota_arredondada = Int(va_m32) + 1
									va_m32=nota_arredondada
								else
									nota_arredondada = Int(va_m32)
									va_m32=nota_arredondada											
								End If
							va_m32=va_m32/10	
							va_m32 = formatNumber(va_m32,1)									
						end if

						if sem_nota3="s" then
							va_m33="&nbsp;"
						else	
							va_m33=va_m33_acumula/peso_acumula
							va_m33=va_m33*10
								decimo = va_m33 - Int(va_m33)
								If decimo >= 0.5 Then
									nota_arredondada = Int(va_m33) + 1
									va_m33=nota_arredondada
								else
									nota_arredondada = Int(va_m33)
									va_m33=nota_arredondada											
								End If
							va_m33=va_m33/10	
							va_m33 = formatNumber(va_m33,1)								
						end if
						
						if sem_nota4="s" then
							va_m34="&nbsp;"
						else	
							va_m34=va_m34_acumula/peso_acumula
							va_m34=va_m34*10
								decimo = va_m34 - Int(va_m34)
								If decimo >= 0.5 Then
									nota_arredondada = Int(va_m34) + 1
									va_m34=nota_arredondada
								else
									nota_arredondada = Int(va_m34)
									va_m34=nota_arredondada											
								End If
							va_m34=va_m34/10	
							va_m34 = formatNumber(va_m34,1)									
						end if							
					end if									
				end if					
					if isnull(f1) or f1="&nbsp;"  or f1="" then
					soma_f1=0
					f1="&nbsp;"
					else
					soma_f1=f1
					end if

					if isnull(f2) or f2="&nbsp;"  or f2="" then
					soma_f2=0
					f2="&nbsp;" 
					else
					soma_f2=f2
					end if
					
					if isnull(f3) or f3="&nbsp;"  or f3="" then
					soma_f3=0
					f3="&nbsp;"
					else
					soma_f3=f3
					end if					
					
					soma_f=soma_f1+soma_f2+soma_f3
					
					if isnull(va_m31) or va_m31="&nbsp;" or va_m31="" then
					dividendo1=0
					divisor1=0
					va_m31="&nbsp;" 
					else
					dividendo1=va_m31
					divisor1=1
					end if	
					
					if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
					dividendo2=0
					divisor2=0
					va_m32="&nbsp;" 
					else
					dividendo2=va_m32
					divisor2=1
					end if
					
					if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
					dividendo3=0
					divisor3=0
					va_m33="&nbsp;"
					else
					dividendo3=va_m33
					divisor3=1
					end if
					
					if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
					nota_aux_m2_1="&nbsp;"
					dividendo4=0
					divisor4=0
					va_m34="&nbsp;"
					else
					nota_aux_m2_1=va_m34
					dividendo4=va_m34
					divisor4=1
					end if
					dividendo1=dividendo1*1	
					dividendo2=dividendo2*1	
					dividendo3=dividendo3*1	
					divisor1=divisor1*1
					divisor2=divisor2*1
					divisor3=divisor3*1										
					dividendo_ma=dividendo1+dividendo2+dividendo3
					divisor_ma=divisor1+divisor2+divisor3
					divisor_m3=divisor1+divisor2+divisor3+(divisor4*2)				
					'response.Write(dividendo_ma&"<<")
					
					if divisor_ma<3 then
					ma="&nbsp;"
					else
					ma=dividendo_ma
					end if
					
					if ma="&nbsp;" then
					else
					ma=ma*10
							decimo = ma - Int(ma)
							If decimo >= 0.5 Then
								nota_arredondada = Int(ma) + 1
								ma=nota_arredondada
							else
								nota_arredondada = Int(ma)
								ma=nota_arredondada											
							End If
					ma=ma/10	
						ma = formatNumber(ma,1)										
					end if

					
					if ma="&nbsp;" then
					dividendo_m2=0
					divisor_m2=0
					else
					dividendo_m2=ma+(dividendo4*2)
					divisor_m2=1
					end if
					
					if divisor_m2=0 then
					m2="&nbsp;"
					else
					'response.Write(mf&"="&dividendo_mf&"/"&divisor_mf)
					m2=dividendo_m2
					end if
					
					if m2="&nbsp;" then
					else
					m3=m2/divisor_m3
					m3=m3*10
						decimo = m3 - Int(m3)
							If decimo >= 0.5 Then
								nota_arredondada = Int(m3) + 1
								m3=nota_arredondada
							else
								nota_arredondada = Int(m3)
								m3=nota_arredondada					
							End If
					m3=m3/10		
						m3 = formatNumber(m3,1)															
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
			%>
                    <tr> 
                      <td width="339" class="<%response.Write(cor)%>"> 
                        <%response.Write(no_materia)
						  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
						  %>
                      </td>
                      <td width="60" class="<%response.Write(cor)%>"> 
                        <div align="center">
                          <%
							if showapr1="s" and showprova1="s" then																	
								response.Write(va_m31)
							else
								response.Write("&nbsp;")									
							end if	
							%>
                        </div></td>
                      <td width="60" class="<%response.Write(cor)%>"> 
                        <div align="center">
                          <%
							if showapr2="s" and showprova2="s" then																	
								response.Write(va_m32)
							else
								response.Write("&nbsp;")									
							end if	
							%>
                        </div></td>
                      <td width="60" class="<%response.Write(cor)%>"> 
                        <div align="center">
                          <%
							if showapr3="s" and showprova3="s" then																	
								response.Write(va_m33)
							else
								response.Write("&nbsp;")									
							end if	
							%>
                        </div></td>
                      <td width="60" class="<%response.Write(cor)%>"> 
                        <div align="center">
                            <%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" then
								response.Write(ma)
							else
								response.Write("&nbsp;")	
							end if
							%>
                        </div></td>
                      <td width="60" class="<%response.Write(cor)%>"> 
                        <div align="center">
                          <%
							if showapr4="s" and showprova4="s" then																	
								response.Write(va_m34)
							else
								response.Write("&nbsp;")									
							end if	
							%>
                        </div></td>
                      <td width="60" class="<%response.Write(cor)%>"> 
                        <div align="center">
                            <%
							'if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s"  and showapr4="s" and showprova4="s" then
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" then							
								response.Write(m2)
							else
								response.Write("&nbsp;")	
							end if
							%>
                        </div></td>
                      <td width="60" class="<%response.Write(cor)%>"> 
                        <div align="center"> 
                            <%
							'if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s"  and showapr4="s" and showprova4="s" then
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" then		
								response.Write(m3)
							else
								response.Write("&nbsp;")	
							end if
							%>
                        </div></td>
                      <td width="1">&nbsp;</td>
                      <td width="60" class="<%response.Write(cor)%>"> 
                        <div align="center">&nbsp; 
                          <%													
							response.Write(f1)	
							%>
                        </div></td>
                      <td width="60" class="<%response.Write(cor)%>"> 
                        <div align="center">&nbsp; 
                          <%													
							response.Write(f2)	
							%>
                        </div></td>
                      <td width="60" class="<%response.Write(cor)%>">
<div align="center">&nbsp; 
                          <%													
							response.Write(f3)	
							%>
                        </div></td>
                      <td width="60" class="<%response.Write(cor)%>">
<div align="center">&nbsp; 
                          <%													
							response.Write(soma_f)	
							%>
                        </div></td>
                    </tr>
                    <%
			check=check+1
			RSprog.MOVENEXT
			wend
end if		
			%>
                  </table>	
        </div>
      </div></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td colspan="2"><div align="right"></div></td>
  </tr>
  <tr> 
    <td colspan="2" class="linhaTopoL">
<div align="right"> 
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