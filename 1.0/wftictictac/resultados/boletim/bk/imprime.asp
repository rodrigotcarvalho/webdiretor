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
<!--#include file="../../inc/connect_wf.asp"-->
<!--#include file="../../inc/connect_al.asp"-->
<!--#include file="../../inc/connect_g.asp"-->
<!--#include file="../../inc/connect_n.asp"-->
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
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
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
    <td colspan="2"><%		Set RS_tb = Server.CreateObject("ADODB.Recordset")
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
		SQL4 = "SELECT * FROM TB_Controle"
		RS4.Open SQL4, CON
	
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
		if co_apr6="D"then
		showapr6="n"
		else 
		showapr6="s"
		end if
		if co_prova6="D"then
		showprova6="n"
		else 
		showprova6="s"
		end if				

'if notaFIL="TB_NOTA_A" or notaFIL="TB_NOTA_B" or notaFIL="TB_NOTA_C" then			
	%>
                    
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="252" rowspan="2" class="tabela"> <div align="left"><strong>Disciplina</strong></div></td>
                      <td colspan="13" class="tabela"> <div align="center">Aproveitamento</div></td>
                      <td colspan="4" class="tabela"><div align="center">Freq&uuml;&ecirc;ncia 
                          (Faltas)</div></td>
                    </tr>
                    <tr> 
                      <td width="68" class="tabela"> <div align="center">BIM 
                          1</div></td>
                      <td width="68" class="tabela"> <div align="center">BIM 
                          2</div></td>
                      <td width="68" class="tabela"><div align="center">REC 
                          PARAL</div></td>
                      <td width="68" class="tabela"><div align="center">BIM 
                          1 * </div></td>
                      <td width="68" class="tabela"><div align="center">BIM 
                          2 * </div></td>
                      <td width="68" class="tabela"> <div align="center">BIM 
                          3</div></td>
                      <td width="68" class="tabela"> <div align="center">BIM 
                          4</div></td>
                      <td width="68" class="tabela"><div align="center">SOMA 
                          PER</div></td>
                      <td width="68" class="tabela"> <div align="center">MEDIA 
                          ANUAL</div></td>
                      <td width="68" class="tabela"><div align="center">PROVA RECUP 
                      FINAL</div></td>
                      <td width="68" class="tabela"> <div align="center">M&Eacute;DIA RECUP 
                          FINAL</div></td>
                      <td width="68" class="tabela"> <div align="center">PROVA 
                          FINAL</div></td>
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
						
						if check mod 2 =0 then
						cor = "tabela" 
						cor2 = "tabela" 				
						else 
						cor ="tabela"
						cor2 = "tabela" 
						end if										

				'response.Write(materia&" "&mae&" "&in_co&"<br>")
				
				if mae=TRUE and in_co=TRUE then
				f1_ac=0
				f2_ac=0
				f3_ac=0
				f4_ac=0
				mb1_ac=0
				mb2_ac=0
				mrec_ac=0
				mb3_ac=0
				mb4_ac=0
				divisor_mb1_ac=0
				divisor_mb2_ac=0
				divisor_mrec_ac=0
				divisor_mb3_ac=0
				divisor_mb4_ac=0
				divisor_mb5_ac=0
				divisor_mb6_ac=0																			
				ms1_ac=0
				ms2_ac=0
				ms3_ac=0
				ms35_ac=0
				ms36_ac=0								
				ma_ac=0
				peso_ac=0
				m2_ac=0
				m3_ac=0				
				ordem2=ordem+1
				tentativas=0
				conta_filhas=0
				While verifica="ok"
				
				Set RSprog2 = Server.CreateObject("ADODB.Recordset")
				SQLprog2 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' And NU_Ordem_Boletim="&ordem2
				RSprog2.Open SQLprog2, CON0
				
					if RSprog2.EOF then
						ordem2=ordem2+1	
						tentativas=tentativas+1
						if tentativas>3 then
						verifica="no"
						end if
					else
					
						materia=RSprog2("CO_Materia")
						mae=RSprog2("IN_MAE")
						fil=RSprog2("IN_FIL")
						in_co=RSprog2("IN_CO")
						nu_peso_fil=RSprog2("NU_Peso")
						
							'response.Write(materia&"-"&mae&"-"&fil&"-"&in_co&"-"&nu_peso_fil&"-"&peso_ac&"<BR>")						
							'response.Write(f1_ac&"-"&f2_ac&"-"&f3_ac&"-"&f4_ac&"-"&ms1_ac&"-"&ms2_ac&"-"&ms3_ac&" <BR>")									
					
						if mae=false AND fil =false AND in_co=True then
						conta_filhas=conta_filhas+1
						ordem2=ordem2+1
						verifica="ok"		
						peso_ac=peso_ac+nu_peso_fil		
								
								
							Set CON_N = Server.CreateObject("ADODB.Connection") 
							ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
							CON_N.Open ABRIRn
						
							for periodofil=1 to 6				
									
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
								
								if periodofil=1 then							
									if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
									dividendo1=0
									divisor1=0
									mb1_ac=mb1_ac
									divisor_mb1_ac=divisor_mb1_ac
									else
									dividendo1=va_m31
									divisor1=1
									'response.Write(materia&"-"&mb1_ac&"="&mb1_ac&"+("&va_m31&"*"&nu_peso_fil&")<BR>")
									mb1_ac=mb1_ac+(va_m31*nu_peso_fil)
									divisor_mb1_ac=divisor_mb1_ac+1
									end if	
								elseif periodofil=2 then
									if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
									dividendo2=0
									divisor2=0
									mb2_ac=mb2_ac
									divisor_mb2_ac=divisor_mb2_ac
									else
									dividendo2=va_m32
									divisor2=1
									mb2_ac=mb2_ac+(va_m32*nu_peso_fil)
									divisor_mb2_ac=divisor_mb2_ac+1
									end if
								
									if isnull(va_rec_sem) or va_rec_sem="&nbsp;"  or va_rec_sem="" then
									dividendorec=0
									divisorrec=0
									mrec_ac=mrec_ac
									divisor_mrec_ac=divisor_mrec_ac									
									else
									dividendorec=va_rec_sem
									divisorrec=1
									mrec_ac=mrec_ac+(va_rec_sem*nu_peso_fil)
									divisor_mrec_ac=divisor_mrec_ac+1									
									end if
									'response.Write("-"&divisor_mrec_ac&"-")
								elseif periodofil=3 then								
									if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
									dividendo3=0
									divisor3=0
									mb3_ac=mb3_ac
									divisor_mb3_ac=divisor_mb3_ac
									else
									dividendo3=va_m33
									divisor3=1
									mb3_ac=mb3_ac+(va_m33*nu_peso_fil)
									divisor_mb3_ac=divisor_mb3_ac+1
									end if								
								elseif periodofil=4 then								
									if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
									dividendo4=0
									divisor4=0
									mb4_ac=mb4_ac
									divisor_mb4_ac=divisor_mb4_ac
									else
									dividendo4=va_m34
									divisor4=1
									mb4_ac=mb4_ac+(va_m34*nu_peso_fil)
									divisor_mb4_ac=divisor_mb4_ac+1
									end if
								elseif periodofil=5 then	
									if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
										ms35_ac=ms35_ac
										divisor_mb5_ac=divisor_mb5_ac									
									else
										ms35_ac=ms35_ac+(va_m35*nu_peso_fil)
										divisor_mb5_ac=divisor_mb5_ac+1				
									end if	
								elseif periodofil=6 then								
									if isnull(va_m36) or va_m36="&nbsp;" or va_m36="" then
										ms36_ac=ms36_ac
										divisor_mb6_ac=divisor_mb6_ac																	
									else
										ms36_ac=ms36_ac+(va_m36*nu_peso_fil)
										divisor_mb6_ac=divisor_mb6_ac+1																	
									end if									
								end if
								
								
								
								'response.Write("N :"&va_m31&"-"&va_m32&"-"&va_m33&"-"&va_m34&"-"&va_m35&"-"&va_m36&" <BR>")
								'response.Write("D :"&dividendo1&"-"&dividendo2&"-"&divisor1&"-"&divisor2&"-"&dividendo_ms1&"-"&divisor_ms1&" <BR>")
								dividendo_ms1=dividendo1+dividendo2
								divisor_ms1=divisor1+divisor2
															


							NEXT
							
							if divisor1=1 and divisorrec=0 then
								ms1="&nbsp;"
								dividendoms1=dividendo1
								divisorms1=1
							elseif divisor1=0 and divisorrec=0 then
								ms1="&nbsp;"							
							else
								if dividendo1<7 and dividendo1<dividendorec then
									ms1=dividendorec
								else
									ms1=dividendo1
								end if
								ms1 = formatNumber(ms1,1)
								dividendoms1=ms1
								divisorms1=1
							end if
							
							if divisor2=1 and divisorrec=0 then
								ms2="&nbsp;"
								dividendoms2=dividendo2
								divisorms2=1							
							elseif divisor2=0 and divisorrec=0 then
								ms2="&nbsp;"
							else
								if dividendo2<7 and dividendo2<dividendorec then
									ms2=dividendorec
								else
									ms2=dividendo2
								end if
								ms2 = formatNumber(ms2,1)									
							dividendoms2=ms2
							divisorms2=1
							end if							
							
							
								if ms1="&nbsp;" then
								ms1_ac=ms1_ac
								else
								ms1_ac=ms1_ac+(ms1*nu_peso_fil)				
								end if
								
								
								if ms2="&nbsp;" then
								ms2_ac=ms2_ac
								else
								ms2_ac=ms2_ac+(ms2*nu_peso_fil)				
								end if
								

											
								'dividendo_ma=dividendoms2+dividendoms3
								'divisor_ma=divisorms2+divisorms3
								
								'response.Write(">>"&dividendo_ma&"<<")
								
								if divisor_ma<2 then
								ma="&nbsp;"
								else
								ma=dividendo_ma/divisor_ma
								end if
								
								if ma="&nbsp;" then
								'ma_ac=ma_ac
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
								
								else								
								nota_aux_m2_1=ms35_ac/peso_ac
								nota_aux_m3_1=ms36_ac/peso_ac									
													
									resultados_apurados=regra_aprovacao (curso,etapa,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,"wfboletim")
									resultados_aluno=split(resultados_apurados,"#!#")
									res1=resultados_aluno(1)
									res2=resultados_aluno(3)
									res3=resultados_aluno(5)
									m2=resultados_aluno(2)
									m3=resultados_aluno(4)
									m2_ac=m2_ac+(m2*nu_peso_fil)
									m3_ac=m3_ac+(m3*nu_peso_fil)
									ma = formatNumber(ma,1)
								end if

								
								if f1="&nbsp;" or isnull(f1) then
								f1_ac=f1_ac
								else
								f1_ac=f1_ac+f1				
								end if
								
								if f2="&nbsp;" or isnull(f2)  then
								f2_ac=f2_ac
								else
								f2_ac=f2_ac+f2				
								end if
	
								if f3="&nbsp;" or isnull(f3)  then
								f3_ac=f3_ac
								else
								f3_ac=f3_ac+f3				
								end if
								
								if f4="&nbsp;" or isnull(f4) then
								f4_ac=f4_ac
								else
								f4_ac=f4_ac+f4				
								end if
																
						else
						verifica="no"
						end if
					end if									
				wend
				f1=f1_ac
				f2=f2_ac										
				f3=f3_ac
				f4=f4_ac
				
				if peso_ac=0 then
				mb1="&nbsp;"				
				mb2="&nbsp;"
				mb3="&nbsp;"
				mb4="&nbsp;"
				
				ms1="&nbsp;"
				ms2="&nbsp;"
				ms3="&nbsp;"								
				else
				divisor_mb1_ac=divisor_mb1_ac*1
				divisor_mb2_ac=divisor_mb2_ac*1
				divisor_mb3_ac=divisor_mb3_ac*1
				divisor_mb4_ac=divisor_mb4_ac*1
				divisor_mrec_ac=divisor_mrec_ac*1																				
				conta_filhas=conta_filhas*1				
				
					if divisor_mb1_ac<conta_filhas then
					mb1="&nbsp;"
					somamb1=0
					validamb1=0					
					else
					mb1=mb1_ac/peso_ac
						decimo = mb1 - Int(mb1)					
						If decimo >= 0.75 Then
							nota_arredondada = Int(mb1) + 1
							mb1=nota_arredondada
						elseIf decimo >= 0.25 Then
							nota_arredondada = Int(mb1) + 0.5
							mb1=nota_arredondada
						else
							nota_arredondada = Int(mb1)
							mb1=nota_arredondada											
						End If			
						mb1 = formatNumber(mb1,1)					
						somamb1=mb1
						validamb1=1															
					end if
					
					if divisor_mb2_ac<conta_filhas then
					mb2="&nbsp;"
						somamb2=0
						validamb2=0											
					else
					mb2=mb2_ac/peso_ac
						decimo = mb2 - Int(mb2)					
						If decimo >= 0.75 Then
							nota_arredondada = Int(mb2) + 1
							mb2=nota_arredondada
						elseIf decimo >= 0.25 Then
							nota_arredondada = Int(mb2) + 0.5
							mb2=nota_arredondada
						else
							nota_arredondada = Int(mb2)
							mb2=nota_arredondada											
						End If			
						mb2 = formatNumber(mb2,1)					
						somamb2=mb2
						validamb2=1															
					end if					

					if divisor_mb3_ac<conta_filhas then
					mb3="&nbsp;"
						somamb3=0
						validamb3=0											
					else
					mb3=mb3_ac/peso_ac
						decimo = mb3 - Int(mb3)					
						If decimo >= 0.75 Then
							nota_arredondada = Int(mb3) + 1
							mb3=nota_arredondada
						elseIf decimo >= 0.25 Then
							nota_arredondada = Int(mb3) + 0.5
							mb3=nota_arredondada
						else
							nota_arredondada = Int(mb3)
							mb3=nota_arredondada											
						End If			
						mb3 = formatNumber(mb3,1)
						somamb3=mb3
						validamb3=1										
					end if
				
					if divisor_mb4_ac<conta_filhas then
					mb4="&nbsp;"
						somamb4=0
						validamb4=0											
					else
					mb4=mb4_ac/peso_ac
						decimo = mb4 - Int(mb4)					
						If decimo >= 0.75 Then
							nota_arredondada = Int(mb4) + 1
							mb4=nota_arredondada
						elseIf decimo >= 0.25 Then
							nota_arredondada = Int(mb4) + 0.5
							mb4=nota_arredondada
						else
							nota_arredondada = Int(mb4)
							mb4=nota_arredondada											
						End If			
						mb4 = formatNumber(mb4,1)					
						somamb4=mb4
						validamb4=1						
					end if				

				'response.Write(divisor_mrec_ac&"-"&conta_filhas)
				divisor_mrec_ac=divisor_mrec_ac*1
				conta_filhas=conta_filhas*1
				somamb1=somamb1*1
				somamb2=somamb2*1
				somamb3=somamb3*1
				somamb4=somamb4*1				
				if divisor_mrec_ac<conta_filhas then
					ms1="&nbsp;"
					ms2="&nbsp;"
				else
					va_rec_sem=mrec_ac/peso_ac
					decimo = va_rec_sem - Int(va_rec_sem)						
						If decimo >= 0.75 Then
							nota_arredondada = Int(va_rec_sem) + 1
							va_rec_sem=nota_arredondada
						elseIf decimo >= 0.25 Then
							nota_arredondada = Int(va_rec_sem) + 0.5
							va_rec_sem=nota_arredondada
						else
							nota_arredondada = Int(va_rec_sem)
							va_rec_sem=nota_arredondada											
						End If			
					va_rec_sem = formatNumber(va_rec_sem,1)
										
					somamb1=somamb1*1	
					va_rec_sem=va_rec_sem*1						
					if somamb1<7 and somamb1<va_rec_sem then
						teste_ms1=(somamb1+va_rec_sem)/2
						if teste_ms1>somamb1 then
							ms1=teste_ms1
							decimo = ms1 - Int(ms1)
							If decimo >= 0.75 Then
								nota_arredondada = Int(ms1) + 1
								ms1=nota_arredondada
							elseIf decimo >= 0.25 Then
								nota_arredondada = Int(ms1) + 0.5
								ms1=nota_arredondada
							else
								nota_arredondada = Int(ms1)
								ms1=nota_arredondada											
							End If			
							ms1 = formatNumber(ms1,1)							
						else
							ms1=somamb1
						end if
					else
						ms1=somamb1
						decimo = ms1 - Int(ms1)
						If decimo >= 0.75 Then
							nota_arredondada = Int(ms1) + 1
							ms1=nota_arredondada
						elseIf decimo >= 0.25 Then
							nota_arredondada = Int(ms1) + 0.5
							ms1=nota_arredondada
						else
							nota_arredondada = Int(ms1)
							ms1=nota_arredondada											
						End If			
						ms1 = formatNumber(ms1,1)									
					end if					
					
					somamb2=somamb2*1		
					va_rec_sem=va_rec_sem*1						
					if somamb2<7 and somamb2<va_rec_sem then
						teste_ms2=(somamb2+va_rec_sem)/2
						if teste_ms2>somamb2 then
							ms2=teste_ms2
							decimo = ms2 - Int(ms2)
							If decimo >= 0.75 Then
								nota_arredondada = Int(ms2) + 1
								ms2=nota_arredondada
							elseIf decimo >= 0.25 Then
								nota_arredondada = Int(ms2) + 0.5
								ms2=nota_arredondada
							else
								nota_arredondada = Int(ms2)
								ms2=nota_arredondada											
							End If			
							ms2 = formatNumber(ms2,1)							
						else
							ms2=somamb2
						end if
					else
						ms2=somamb2
						decimo = ms2 - Int(ms2)
						If decimo >= 0.75 Then
							nota_arredondada = Int(ms2) + 1
							ms2=nota_arredondada
						elseIf decimo >= 0.25 Then
							nota_arredondada = Int(ms2) + 0.5
							ms2=nota_arredondada
						else
							nota_arredondada = Int(ms2)
							ms2=nota_arredondada											
						End If			
						ms2 = formatNumber(ms2,1)									
					end if																													
					somamb1=ms1
					somamb2=ms2										
					va_rec_sem = formatNumber(va_rec_sem,1)
				end if
				
				'end if
				'response.Write(dividendoms1&"+"&dividendoms2&"+"&dividendo3&"+"&dividendo4)
				somamb1=somamb1*1
				somamb2=somamb2*1
				somamb3=somamb3*1
				somamb4=somamb4*1	
				validamb1=validamb1*1
				validamb2=validamb2*1
				validamb3=validamb3*1
				validamb4=validamb4*1																										
				dividendo_ms3=somamb1+somamb2+somamb3+somamb4
				divisor_ms3=validamb1+validamb2+validamb3+validamb4
				'response.Write(validamb1&"+"&validamb2&"+"&validamb3&"+"&validamb4)
				if divisor_ms3<4 then
				ms3="&nbsp;"
				dividendoms3=0
				divisorms3=0					
				else
					ms3=dividendo_ms3
					ms3 = formatNumber(ms3,1)								
				dividendoms3=ms3
				divisorms3=1						
				end if					
				dividendoms3=dividendoms3*1		
				dividendo_ma=dividendoms3
				divisor_ma=divisor_ms3

'response.Write(divisor_mb5_ac&"-"&ms35_ac)
					if divisor_mb5_ac<conta_filhas then
						mb5="&nbsp;"
					else						
						mb5=ms35_ac/divisor_mb5_ac
						decimo = mb5 - Int(mb5)					
						If decimo >= 0.75 Then
							nota_arredondada = Int(mb5) + 1
							mb5=nota_arredondada
						elseIf decimo >= 0.25 Then
							nota_arredondada = Int(mb5) + 0.5
							mb5=nota_arredondada
						else
							nota_arredondada = Int(mb5)
							mb5=nota_arredondada											
						End If			
						mb5 = formatNumber(mb5,1)						
					end if
					
					if divisor_mb6_ac<conta_filhas then
						mb6="&nbsp;"
					else						
						mb6=ms36_ac/divisor_mb6_ac
						decimo = mb6 - Int(mb6)					
						If decimo >= 0.75 Then
							nota_arredondada = Int(mb6) + 1
							mb6=nota_arredondada
						elseIf decimo >= 0.25 Then
							nota_arredondada = Int(mb6) + 0.5
							mb6=nota_arredondada
						else
							nota_arredondada = Int(mb6)
							mb6=nota_arredondada											
						End If			
						mb6 = formatNumber(mb6,1)						
					end if

										
				if divisor_ms3<4 then
					ma="&nbsp;"
				else
					ma=dividendoms3/divisor_ms3								
					nota_aux_m2_1=mb5
					nota_aux_m3_1=mb6									
				'response.Write(nota_aux_m2_1&"-"&nota_aux_m3_1&"-"&res1)	
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
						
					resultados_apurados=regra_aprovacao (curso,etapa,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,"wfboletim")
					resultados_aluno=split(resultados_apurados,"#!#")
					res1=resultados_aluno(1)
					res2=resultados_aluno(3)
					res3=resultados_aluno(5)
					m2=resultados_aluno(2)
					m3=resultados_aluno(4)						
					ma = formatNumber(ma,1)					
				end if					

							

				end if
			%>
                    <tr> 
                      <td width="252" class="<%response.Write(cor)%>"> 
                        <%response.Write(no_materia)
						  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
						  %>
                      </td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center">&nbsp;
                          <%
							if showapr1="s" and showprova1="s" then					
							response.Write(mb1)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                          <%
							if showapr2="s" and showprova2="s"  then					
							response.Write(mb2)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">&nbsp; 
                          <%
							if showapr2="s" and showprova2="s"  then					
							response.Write(va_rec_sem)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">&nbsp; 
                          <%
							if showapr1="s" and showprova1="s" then					
							response.Write(ms1)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">&nbsp; 
                          <%
							if showapr2="s" and showprova2="s"  then					
							response.Write(ms2)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                          <%
							if showapr3="s" and showprova3="s" then					
							response.Write(mb3)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                          <%
							if showapr4="s" and showprova4="s"  then					
							response.Write(mb4)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                          <%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then					
							response.Write(ms3)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                          <%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then					
							response.Write(ma)
							else
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">&nbsp;
                        <%
							if showapr5="s" and showprova5="s" then												
							response.Write(mb5)
							end if
							%>
                      </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                          <%
							if showapr5="s" and showprova5="s" then												
							response.Write(m2)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center">&nbsp;
                          <%
							if showapr6="s" and showprova6="s" then												
							response.Write(mb6)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                          <%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then							  
								if divisor_ms3<4 then
								response.Write("&nbsp;")							
								else
									if nota_aux_m3_1="&nbsp;" then
										if nota_aux_m2_1="&nbsp;" then					
										response.Write(res1)
										else
										response.Write(res2)
										end if
									else
									response.Write(res3)
									end if
								end if	
							else
								response.Write("&nbsp;")
							end if																
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center">&nbsp; 
                          <%
							response.Write(f1)
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center">&nbsp; 
                          <%
							response.Write(f2)
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center">&nbsp; 
                          <%
							response.Write(f3)
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center">&nbsp; 
                          <%
							response.Write(f4)
							%>
                        </div></td>
                    </tr>
                    <%				
				
				check=check+1
				elseif mae=false AND fil =false AND in_co=True then
				
				else
			
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
					
						for periodofil=1 to 6
						
												'Set RS_media_turma = Server.CreateObject("ADODB.Recordset")
							'SQL_media_turma = "Select AVG(VA_Media3) as media_turma from "& notaFIL &" WHERE CO_Matricula in ("& alunos_turma &") AND CO_Materia = '"& materia &"' AND NU_Periodo="&periodofil
							'Set RS_media_turma = CON_N.Execute(SQL_media_turma)
							
						'		if periodofil=1 then
						'		va_mt1=RS_media_turma("media_turma")
						'		elseif periodofil=2 then
						'		va_mt2=RS_media_turma("media_turma")
						'		elseif periodofil=3 then
						'		va_mt3=RS_media_turma("media_turma")
						'		elseif periodofil=4 then
						'		va_mt4=RS_media_turma("media_turma")
						'		elseif periodofil=5 then
						'		va_mt5=RS_media_turma("media_turma")
						'		elseif periodofil=6 then
						'		va_mt6=RS_media_turma("media_turma")
						'		end if				
						
						
								
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
							
							if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
							dividendo1=0
							divisor1=0
							else
							dividendo1=va_m31
							divisor1=1
							va_m31 = formatNumber(va_m31,1)							
							end if	
							
							if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
							dividendo2=0
							divisor2=0
							else
							dividendo2=va_m32
							divisor2=1
							va_m32 = formatNumber(va_m32,1)							
							end if
							
							if isnull(va_rec_sem) or va_rec_sem="&nbsp;"  or va_rec_sem="" then
							dividendorec=0
							divisorrec=0
							else
							dividendorec=va_rec_sem
							divisorrec=1
							va_rec_sem = formatNumber(va_rec_sem,1)							
							end if
							
							if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
							dividendo3=0
							divisor3=0
							else
							dividendo3=va_m33
							divisor3=1
							va_m33 = formatNumber(va_m33,1)							
							end if
							
							if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
							dividendo4=0
							divisor4=0
							else
							dividendo4=va_m34
							divisor4=1
							va_m34 = formatNumber(va_m34,1)							
							end if
							'dividendo_ms1=dividendo1+dividendo2
							'divisor_ms1=divisor1+divisor2
							
							
							
							'response.Write(va_m35&"<br>")														
							if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
							nota_aux_m2_1="&nbsp;"
							'dividendo5=0
							'divisor5=0
							else
							nota_aux_m2_1=va_m35
							'dividendo5=va_m35
							'divisor5=1
							va_m35 = formatNumber(va_m35,1)							
							end if
							
							
							if isnull(va_m36) or va_m36="&nbsp;" or va_m36="" then
							nota_aux_m3_1="&nbsp;"
							'dividendo6=0
							'divisor6=0
							else
							nota_aux_m3_1=va_m36
							'dividendo6=va_m36
							'divisor6=1
							va_m36 = formatNumber(va_m36,1)
							end if
							
									
											
							'end if
						NEXT
						
							'if divisor_ms1<2 then
							'ms1="&nbsp;"
							'dividendoms1=0
							'divisorms1=0
							'else
							'ms1=dividendo_ms1/divisor_ms1
							'	decimo = ms1 - Int(ms1)
							'	If decimo >= 0.75 Then
							'		nota_arredondada = Int(ms1) + 1
							'		ms1=nota_arredondada
							'	elseIf decimo >= 0.25 Then
							'		nota_arredondada = Int(ms1) + 0.5
							'		ms1=nota_arredondada
							'	else
							'		nota_arredondada = Int(ms1)
							'		ms1=nota_arredondada											
							'	End If			

							if divisor1=1 and divisorrec=0 then
								ms1="&nbsp;"
								dividendoms1=dividendo1
								divisorms1=1
							elseif divisor1=0 and divisorrec=0 then
								ms1="&nbsp;"							
							else
								dividendo1=dividendo1*1	
								dividendorec=dividendorec*1							
								if dividendo1<7 and dividendo1<dividendorec then
									teste_ms1=(dividendo1+dividendorec)/2
									if teste_ms1>dividendo1 then
										ms1=teste_ms1
									else
										ms1=dividendo1
									end if
										
								else
									ms1=dividendo1
								end if
								decimo = ms1 - Int(ms1)
								If decimo >= 0.75 Then
									nota_arredondada = Int(ms1) + 1
									ms1=nota_arredondada
								elseIf decimo >= 0.25 Then
									nota_arredondada = Int(ms1) + 0.5
									ms1=nota_arredondada
								else
									nota_arredondada = Int(ms1)
									ms1=nota_arredondada											
								End If			
								ms1 = formatNumber(ms1,1)
								dividendoms1=ms1
								divisorms1=1
							end if
							
							
							

							'end if
							
							'dividendo_ms2=dividendoms1+dividendorec
							'divisor_ms2=divisorms1+divisorrec
							
							'if dividendorec=0 then
							'ms2=ms1
							'	if ms2="&nbsp;" then
							'	dividendoms2=0
							'	divisorms2=0
							'	else
							'	dividendoms2=ms2
							'	divisorms2=1						
							'	end if
							'else
							'	ms2=dividendo_ms2/divisor_ms2
							'	if ms2<dividendorec then
							'	ms2=dividendorec
							'	end if
							'		decimo = ms2 - Int(ms2)
							'		If decimo >= 0.75 Then
							'			nota_arredondada = Int(ms2) + 1
							'			ms2=nota_arredondada
							'		elseIf decimo >= 0.25 Then
							'			nota_arredondada = Int(ms2) + 0.5
							'			ms2=nota_arredondada
							'		else
							'			nota_arredondada = Int(ms2)
							'			ms2=nota_arredondada											
							'		End If
							if divisor2=1 and divisorrec=0 then
								ms2="&nbsp;"
								dividendoms2=dividendo2
								divisorms2=1							
							elseif divisor2=0 and divisorrec=0 then
								ms2="&nbsp;"
							else
								dividendo2=dividendo2*1
								dividendorec=dividendorec*1							
								if dividendo2<7 and dividendo2<dividendorec then
									teste_ms2=(dividendo2+dividendorec)/2
									'response.Write(teste_ms2)
									if teste_ms2>dividendo2 then
										ms2=teste_ms2
									else
										ms2=dividendo2
									end if
								else
									ms2=dividendo2
								end if
								decimo = ms2 - Int(ms2)
								If decimo >= 0.75 Then
									nota_arredondada = Int(ms2) + 1
									ms2=nota_arredondada
								elseIf decimo >= 0.25 Then
									nota_arredondada = Int(ms2) + 0.5
									ms2=nota_arredondada
								else
									nota_arredondada = Int(ms2)
									ms2=nota_arredondada											
								End If			
								ms2 = formatNumber(ms2,1)									
							dividendoms2=ms2
							divisorms2=1
							end if							
							
							'end if
							'response.Write(dividendoms1&"+"&dividendoms2&"+"&dividendo3&"+"&dividendo4)
							dividendoms1=dividendoms1*1
							dividendoms2=dividendoms2*1
							dividendo3=dividendo3*1
							dividendo4=dividendo4*1					
							dividendo_ms3=dividendoms1+dividendoms2+dividendo3+dividendo4
							divisor_ms3=divisorms1+divisorms2+divisor3+divisor4
							'response.Write(divisorms1&"+"&divisorms2&"+"&divisor3&"+"&divisor4)
							if divisor_ms3<4 then
							ms3="&nbsp;"
							dividendoms3=0
							divisorms3=0					
							else
							ms3=dividendo_ms3
								'decimo = ms3 - Int(ms3)
								'If decimo >= 0.75 Then
								'	nota_arredondada = Int(ms3) + 1
								'	ms3=nota_arredondada
								'elseIf decimo >= 0.25 Then
								'	nota_arredondada = Int(ms3) + 0.5
								'	ms3=nota_arredondada
								'else
								'	nota_arredondada = Int(ms3)
								'	ms3=nota_arredondada											
								'End If
								ms3 = formatNumber(ms3,1)								
							dividendoms3=ms3
							divisorms3=1						
							end if					
							dividendoms3=dividendoms3*1		
							dividendo_ma=dividendoms3
							divisor_ma=divisor_ms3
							
							'response.Write(dividendo_ma&"<<")
							
							if divisor_ma<4 then
							ma="&nbsp;"
							else
							ma=dividendo_ma/divisor_ma
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
							else	
															
								resultados_apurados=regra_aprovacao (curso,etapa,ma,nota_aux_m2_1,nota_aux_m2_2,nota_aux_m3_1,nota_aux_m3_2,"wfboletim")
								resultados_aluno=split(resultados_apurados,"#!#")
								res1=resultados_aluno(1)
								res2=resultados_aluno(3)
								res3=resultados_aluno(5)
								m2=resultados_aluno(2)
								m3=resultados_aluno(4)
								ma = formatNumber(ma,1)
							end if
							
															
'response.Write(Session("resultado_1") &" - "& Session("resultado_2") &" - "& Session("resultado_3")&"<<")
			%>
                    <tr> 
                      <td width="252" class="<%response.Write(cor)%>"> 
                        <%response.Write(no_materia)
						  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
						  %>
                      </td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                          <%
							if showapr1="s" and showprova1="s" then																	
							response.Write(va_m31)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                          <%
							if showapr2="s" and showprova2="s"  then												
							response.Write(va_m32)						
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">&nbsp; 
                          <%
							if showapr2="s" and showprova2="s"  then				
							response.Write(va_rec_sem)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">&nbsp; 
                          <%
							if showapr1="s" and showprova1="s"  then					
							response.Write(ms1)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">&nbsp; 
                          <%
							if showapr2="s" and showprova2="s"  then					
							response.Write(ms2)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                          <%
							if showapr3="s" and showprova3="s" then					
							response.Write(va_m33)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                          <%
							if showapr4="s" and showprova4="s"  then					
							response.Write(va_m34)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center">&nbsp;
                          <%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then					
							response.Write(ms3)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                          <%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then					
							response.Write(ma)
							else
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center">&nbsp;
                        <%
							if showapr5="s" and showprova5="s" then												
							response.Write(va_m35)
							end if
							%>
                      </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                          <%
							if showapr5="s" and showprova5="s" then												
							response.Write(m2)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                          <%
							if showapr6="s" and showprova6="s" then												
							response.Write(va_m36)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center">&nbsp; 
                          <%
							if showapr1="s" and showprova1="s" and showapr2="s" and showprova2="s" and showapr3="s" and showprova3="s" and showapr4="s" and showprova4="s" then							  
								if nota_aux_m3_1="&nbsp;"then
									if nota_aux_m2_1="&nbsp;"then					
									response.Write(res1)
									else
									response.Write(res2)
									end if
								else
								response.Write(res3)
								end if
							else
							response.Write("&nbsp;")
							end if								
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center">&nbsp; 
                          <%
							response.Write(f1)
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center">&nbsp; 
                          <%
							response.Write(f2)
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center">&nbsp; 
                          <%
							response.Write(f3)
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center">&nbsp; 
                          <%
							response.Write(f4)
							%>
                        </div></td>
                    </tr>
                    <%
					check=check+1
					end if	
			RSprog.MOVENEXT
			wend
			
			'Set RSF = Server.CreateObject("ADODB.Recordset")
			'SQL_N = "Select * from TB_Frequencia_Periodo WHERE CO_Matricula = "& cod
			'Set RSF = CON_N.Execute(SQL_N)
			
			'if RSF.eof THEN
			'f1="&nbsp;"
			'f2="&nbsp;"
			'f3="&nbsp;"
			'f4="&nbsp;"			
			'else	
			'f1=RSF("NU_Faltas_P1")
			'f2=RSF("NU_Faltas_P2")
			'f3=RSF("NU_Faltas_P3")
			'f4=RSF("NU_Faltas_P4")		
			'END IF		
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
                      <td>&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
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
    </td>
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