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
periodo_check= dados_opt(1)
notaFIL= dados_opt(2)

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

if notaFIL ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na

elseif notaFIL="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb

elseif notaFIL ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc

elseif notaFIL ="TB_NOTA_E" then
		CAMINHOn = CAMINHO_ne		
		
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

un_cep=cep&"-"&cep3


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
periodo_check=periodo_check*1		
if periodo_check=1 then		
		if co_apr1="D"then
		showapr_ctrl="n"
		else 
		showapr_ctrl="s"
		end if
		if co_prova1="D"then
		showprova_ctrl="n"
		else 
		showprova_ctrl="s"
		end if
elseif periodo_check=2 then		
		if co_apr2="D"then
		showapr_ctrl="n"
		else 
		showapr_ctrl="s"
		end if
		if co_prova2="D"then
		showprova_ctrl="n"
		else 
		showprova_ctrl="s"
		end if					
elseif periodo_check=3 then		
		if co_apr3="D"then
		showapr_ctrl="n"
		else 
		showapr_ctrl="s"
		end if
		if co_prova3="D"then
		showprova_ctrl="n"
		else 
		showprova_ctrl="s"
		end if
elseif periodo_check=4 then		
		if co_apr4="D"then
		showapr_ctrl="n"
		else 
		showapr_ctrl="s"
		end if
		if co_prova4="D"then
		showprova_ctrl="n"
		else 
		showprova_ctrl="s"
		end if
elseif periodo_check=5 then		
		if co_apr5="D"then
		showapr_ctrl="n"
		else 
		showapr_ctrl="s"
		end if
		if co_prova5="D"then
		showprova_ctrl="n"
		else 
		showprova_ctrl="s"
		end if
elseif periodo_check=6 then		
		if co_apr6="D"then
		showapr_ctrl="n"
		else 
		showapr_ctrl="s"
		end if
		if co_prova6="D"then
		showprova_ctrl="n"
		else 
		showprova_ctrl="s"
		end if		
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
            PARCIAIS</strong></font></td>
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
    <td colspan="2">
<%

		Set CON_N = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIRn

if notaFIL="TB_NOTA_A" then
minimo_recuperacao= 60
%>
                      
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="230" class="tabela"> 
            <div align="left">Disciplina</div></td>
          <td width="35"  class="tabela"> <div align="center">T1</div></td>
          <td width="35"  class="tabela"> <div align="center">T2</div></td>
          <td width="35"  class="tabela"> <div align="center">T3</div></td>
          <td width="35"  class="tabela"> <div align="center">T4</div></td>
          <td width="35" class="tabela"> <div align="center">MT</div></td>
<!--          <td width="31" class="tabela">&nbsp;</td>-->
          <td width="35"  class="tabela"> <div align="center">PR1</div></td>
          <td width="35"  class="tabela"> <div align="center">PR2</div></td>
          <td width="35"  class="tabela"> <div align="center">PR3</div></td>
          <td width="35" class="tabela"> <div align="center"> MP</div></td>
<!--          <td width="31" class="tabela">&nbsp;</td>-->
          <td width="35" class="tabela"> <div align="center">M1</div></td>
          <td width="35"  class="tabela"> <div align="center">Bon</div></td>
          <td width="35"  class="tabela"> <div align="center">M2</div></td>
          <td width="35"  class="tabela"> <div align="center">Rec</div></td>
          <td width="35"  class="tabela"> <div align="center">M3</div></td>
          <td width="230" class="tabela"> 
            <div align="center">Data/Hora</div></td>
        </tr>
<!--        <tr>
          <td width="225" class="tabela">&nbsp;</td> 
          <td width="31"  class="tabela"> <div align="center">N</div></td>
          <td width="31"  class="tabela"> <div align="center">N</div></td>
          <td width="31"  class="tabela"> <div align="center">N</div></td>
          <td width="31"  class="tabela"> <div align="center">N</div></td>
          <td width="62" class="tabela"> <div align="center">M</div></td>
          <td width="31" class="tabela"> <div align="center">P</div></td>
          <td width="31"  class="tabela"> <div align="center">N</div></td>
          <td width="31"  class="tabela"> <div align="center">N</div></td>
          <td width="31"  class="tabela"> <div align="center">N</div></td>
          <td width="62" class="tabela"> <div align="center"> M</div></td>
          <td width="31" class="tabela"> <div align="center">P</div></td>
          <td width="31" class="tabela"> <div align="center">M</div></td>
          <td width="31"  class="tabela"> <div align="center">N</div></td>
          <td width="31"  class="tabela"> <div align="center">M</div></td>
          <td width="31"  class="tabela"> <div align="center">N</div></td>
          <td width="31"  class="tabela"> <div align="center">M</div></td>
          <td width="191" class="tabela">&nbsp;</td>
        </tr>-->
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
  cor = "tabela" 
 else cor ="tabela"
  end if
	
		Set RSnFIL = Server.CreateObject("ADODB.Recordset")
	  	Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' and NU_Periodo="&periodo_check
		Set RS3 = CON_N.Execute(SQL_N)
if RS3.EOF then
		va_pt="&nbsp;"
		va_pp="&nbsp;"
		va_t1="&nbsp;"
		va_t2="&nbsp;"
		va_t3="&nbsp;"
		va_t4="&nbsp;"
		va_mt="&nbsp;"
		va_p1="&nbsp;"
		va_p2="&nbsp;"
		va_p3="&nbsp;"
		va_mp="&nbsp;"
		va_m1="&nbsp;"
		va_bon="&nbsp;"
		va_m2="&nbsp;"
		va_rec="&nbsp;"
		va_m3="&nbsp;"
		data_grav="&nbsp;"
		hora_grav="&nbsp;"
					
else
		va_pt=RS3("PE_Teste")
		va_pp=RS3("PE_Prova")
		va_t1=RS3("VA_Teste1")
		va_t2=RS3("VA_Teste2")
		va_t3=RS3("VA_Teste3")
		va_t4=RS3("VA_Teste4")
		va_mt=RS3("MD_Teste")
		va_p1=RS3("VA_Prova1")
		va_p2=RS3("VA_Prova2")
		va_p3=RS3("VA_Prova3")
		va_mp=RS3("MD_Prova")
		va_m1=RS3("VA_Media1")
		va_bon=RS3("VA_Bonus")
		va_m2=RS3("VA_Media2")
		va_rec=RS3("VA_Rec")
		va_m3=RS3("VA_Media3")
		data_grav=RS3("DA_Ult_Acesso")
		hora_grav=RS3("HO_ult_Acesso")
		
end if

if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then										
	showapr="n"
	showprova="n"
else
		showapr=showapr_ctrl
		showprova=showprova_ctrl
end if	
		
if hora_grav="&nbsp;" then
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
					
if data_grav="&nbsp;" then
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


'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
'data_inicio=""
'va_faltas=""
'		end if


%>
        <tr class="<%response.Write(cor)%>"> 
          <td class="tabela" width="230"> 
            <%response.Write("&nbsp;"&no_materia)%>
          </td>
          <td class="tabela" width="35"> <div align="center"> 
              <%
							if showapr="s" then							
							response.Write("&nbsp;"&va_t1)
							else
							response.Write("&nbsp;")
							End IF							
							%>
            </div></td>
          <td class="tabela" width="35"> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write("&nbsp;"&va_t2)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="35"
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write("&nbsp;"&va_t3)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="35"
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write("&nbsp;"&va_t4)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="35"
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write("&nbsp;"&va_mt)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
<!--          <td class="tabela" width="31"
> <div align="center"> 
              <%
							if showapr="s" then					
							'response.Write("&nbsp;"&va_pt)
							else
							'response.Write("&nbsp;")
							End IF
							%>
            </div></td>-->
          <td class="tabela" width="35"
> <div align="center"> 
              <%
							if showprova="s" then					
							response.Write("&nbsp;"&va_p1)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="35"
> <div align="center"> 
              <%
							if showprova="s" then					
							response.Write("&nbsp;"&va_p2)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="35"
> <div align="center"> 
              <%
							if showprova="s" then					
							response.Write("&nbsp;"&va_p3)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="35"
> <div align="center"> 
              <%
							if showprova="s" then					
							response.Write("&nbsp;"&va_mp)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
<!--          <td class="tabela" width="31"
> <div align="center"> 
              <%
							if showprova="s" then					
							'response.Write("&nbsp;"&va_pp)
							else
							'response.Write("&nbsp;")
							End IF
							%>
            </div></td>-->
          <td class="tabela" width="35"
> <div align="center"> 
              <%
							if showprova="s" AND showapr="s" then					
							response.Write("&nbsp;"&va_m1)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="35"
> <div align="center"> 
              <%
							if showprova="s" AND showapr="s" then					
							response.Write("&nbsp;"&va_bon)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="35"
> <div align="center"> 
              <%
							if showprova="s" AND showapr="s" then					
							response.Write("&nbsp;"&va_m2)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="35"
> <div align="center"> 
              <%
							if showprova="s" AND showapr="s" then					
							response.Write("&nbsp;"&va_rec)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="35"
> <div align="center"> 
              <%
							if showprova="s" AND showapr="s" then					
							response.Write("&nbsp;"&va_m3)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="230"
> <div align="center"> 
            <%
							if showprova="s" AND showapr="s" then							
							response.Write("&nbsp;"&data_inicio)
							else
							response.Write("&nbsp;")
							End IF
							%></font>
            </td>
        </tr>
        <%check=check+1
RSprog.MOVENEXT
wend

%>
        <tr valign="bottom"> 
          <td height="20" colspan="22"  class="tabela" 
> <div align="right"><font class="tabela_rodape"> T-Teste, MT–Média dos Testes, 
              PR-Prova, MP–Média das Provas, N-Nota e M-M&eacute;dia</font></div></td>
        </tr>
      </table>
      <%
elseif notaFIL="TB_NOTA_B" or notaFIL="TB_NOTA_E" then
%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="235" class="tabela"> 
            <div align="left">Disciplina</div></td>
          <td width="40"  class="tabela"> <div align="center">T1</div></td>
          <td width="40"  class="tabela"> <div align="center">T2</div></td>
          <td width="40" class="tabela"> <div align="center">MT</div></td>
<!--          <td width="37" class="tabela">&nbsp;</td>-->
          <td width="40"  class="tabela"> <div align="center">PR1</div></td>
          <td width="40"  class="tabela"> <div align="center">S</div></td>
          <td width="40"  class="tabela"> <div align="center">PR2</div></td>
          <td width="40" class="tabela"> <div align="center"> MP</div></td>
<!--          <td width="37" class="tabela">&nbsp;</td>-->
          <td width="40" class="tabela"> <div align="center">M1</div></td>
          <td width="40"  class="tabela"> <div align="center">Bon</div></td>
          <td width="40"  class="tabela"> <div align="center">M2</div></td>
          <td width="40"  class="tabela"> <div align="center">Rec</div></td>
          <td width="40"  class="tabela"> <div align="center">M3</div></td>
          <td width="235" class="tabela"> 
            <div align="center">Data/Hora</div></td>
        </tr>
<!--        <tr>
          <td width="225" class="tabela">&nbsp;</td> 
          <td width="36"  class="tabela"> <div align="center">N</div></td>
          <td width="36"  class="tabela"> <div align="center">N</div></td>
          <td width="74" class="tabela"> <div align="center">M</div></td>
          <td width="37" class="tabela"> <div align="center">P</div></td>
          <td width="36"  class="tabela"> <div align="center">N</div></td>
          <td width="36"  class="tabela"> <div align="center">N</div></td>
          <td width="36"  class="tabela"> <div align="center">N</div></td>
          <td width="74" class="tabela"> <div align="center"> M</div></td>
          <td width="37" class="tabela"> <div align="center">P</div></td>
          <td width="37" class="tabela"> <div align="center">M</div></td>
          <td width="36"  class="tabela"> <div align="center">N</div></td>
          <td width="37"  class="tabela"> <div align="center">M</div></td>
          <td width="36"  class="tabela"> <div align="center">N</div></td>
          <td width="37"  class="tabela"> <div align="center">M</div></td>
          <td width="191" class="tabela">&nbsp;</td>
        </tr>-->
        <%
		rec_lancado="sim"
		
				Set RSprog = Server.CreateObject("ADODB.Recordset")
				SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
				RSprog.Open SQLprog, CON0
		
		check=1
			
		while not RSprog.EOF
		
				materia=RSprog("CO_Materia")
				mae=RSprog("IN_MAE")
				fil=RSprog("IN_FIL")
				in_co=RSprog("IN_CO")
				nu_peso=RSprog("NU_Peso")
				ordem=RSprog("NU_Ordem_Boletim")
				
			peso_acumula=0
			m1_ac=0
			m2_ac=0			
			m3_ac=0
			m1_exibe=0
			m2_exibe=0
			m3_exibe=0
					
			if mae=TRUE THEN
			
			check=check+1
			
					Set RS1a = Server.CreateObject("ADODB.Recordset")
					SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia_Principal='"&materia&"'"
					RS1a.Open SQL1a, CON0
					
			if RS1a.EOF then
			
						Set RS1b = Server.CreateObject("ADODB.Recordset")
						SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
						RS1b.Open SQL1b, CON0
						
						no_materia=RS1b("NO_Materia")
						
 if check mod 2 =0 then
  cor = "tabela" 
 else cor ="tabela"
  end if
								Set RSnFIL = Server.CreateObject("ADODB.Recordset")
								Set RS3 = Server.CreateObject("ADODB.Recordset")
								SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' and NU_Periodo="&periodo_check
								Set RS3 = CON_N.Execute(SQL_N)
						if RS3.EOF then
								va_pt="&nbsp;"
								va_pp="&nbsp;"
								va_t1="&nbsp;"
								va_t2="&nbsp;"
								va_mt="&nbsp;"
								va_p1="&nbsp;"
								va_p2="&nbsp;"
								va_p3="&nbsp;"
								va_mp="&nbsp;"
								va_m1="&nbsp;"
								va_bon="&nbsp;"
								va_m2="&nbsp;"
								va_rec="&nbsp;"
								va_m3="&nbsp;"
								data_grav="&nbsp;"
								hora_grav="&nbsp;"
											
						else
								va_pt=RS3("PE_Teste")
								va_pp=RS3("PE_Prova")
								va_t1=RS3("VA_Teste1")
								va_t2=RS3("VA_Teste2")
								va_mt=RS3("MD_Teste")
								va_p1=RS3("VA_Prova1")
								va_p2=RS3("VA_Simul")
								va_p3=RS3("VA_Prova2")
								va_mp=RS3("MD_Prova")
								va_m1=RS3("VA_Media1")
								va_bon=RS3("VA_Bonus")
								va_m2=RS3("VA_Media2")
								va_rec=RS3("VA_Rec")
								va_m3=RS3("VA_Media3")
								data_grav=RS3("DA_Ult_Acesso")
								hora_grav=RS3("HO_ult_Acesso")
								
						end if
						
if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then										
	showapr="n"
	showprova="n"
else
		showapr=showapr_ctrl
		showprova=showprova_ctrl
end if															
								
						if hora_grav="&nbsp;" then
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
											
						if data_grav="&nbsp;" then
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
						
						
						'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
						'data_inicio=""
						'va_faltas=""
						'		end if
						
						
						%>
        <tr class="<%response.Write(cor)%>"> 
          <td class="tabela"  width="235"> 
            <%response.Write("&nbsp;"&no_materia)%>
          </td>
          <td class="tabela"  width="40"> <div align="center"> 
              <%
													if showapr="s" then							
													response.Write("&nbsp;"&va_t1)
													else
							response.Write("&nbsp;")
							End IF							
													%>
            </div></td>
          <td class="tabela"  width="40"> <div align="center"> 
              <%
													if showapr="s" then					
													response.Write("&nbsp;"&va_t2)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showapr="s" then					
													response.Write("&nbsp;"&va_mt)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
<!--          <td class="tabela"  width="37"
						> <div align="center"> 
              <%
													if showapr="s" then					
													'response.Write("&nbsp;"&va_pt)
													else
							'response.Write("&nbsp;")
							End IF
													%>
            </div></td>-->
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showprova="s" then					
													response.Write("&nbsp;"&va_p1)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showprova="s" then					
													response.Write("&nbsp;"&va_p2)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showprova="s" then					
													response.Write("&nbsp;"&va_p3)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showprova="s" and materia<>"LP" then					
													response.Write("&nbsp;"&va_mp)
													else
													response.Write("&nbsp;")
													End IF
													%>
            </div></td>
<!--          <td class="tabela"  width="37"
						> <div align="center"> 
              <%
													if showprova="s" then					
													'response.Write("&nbsp;"&va_pp)
													else
							'response.Write("&nbsp;")
							End IF
													%>
            </div></td>-->
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showprova="s" AND showapr="s" then					
													response.Write("&nbsp;"&va_m1)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showprova="s" AND showapr="s" then					
													response.Write("&nbsp;"&va_bon)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showprova="s" AND showapr="s" then					
													response.Write("&nbsp;"&va_m2)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showprova="s" AND showapr="s" then					
													response.Write("&nbsp;"&va_rec)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showprova="s" AND showapr="s" then					
													response.Write("&nbsp;"&va_m3)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
          <td class="tabela"  width="235"
						> 
            <div align="center"> 
              <%
													if showprova="s" AND showapr="s" then							
													response.Write("&nbsp;"&data_inicio)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
        </tr>
        <%
			else

			
			
 if check mod 2 =0 then
  cor = "tabela" 
 else cor ="tabela"
  end if
			
						Set RS1b = Server.CreateObject("ADODB.Recordset")
						SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
						RS1b.Open SQL1b, CON0
						
				no_materia=RS1b("NO_Materia")
					
						Set RSnFIL = Server.CreateObject("ADODB.Recordset")
						Set RS3 = Server.CreateObject("ADODB.Recordset")
						SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' and NU_Periodo="&periodo_check
						Set RS3 = CON_N.Execute(SQL_N)
				if RS3.EOF then
						va_pt="&nbsp;"
						va_pp="&nbsp;"
						va_t1="&nbsp;"
						va_t2="&nbsp;"
						va_mt="&nbsp;"
						va_p1="&nbsp;"
						va_p2="&nbsp;"
						va_p3="&nbsp;"
						va_mp="&nbsp;"
						va_m1="&nbsp;"
						va_bon="&nbsp;"
						va_m2="&nbsp;"
						va_rec="&nbsp;"
						va_m3="&nbsp;"
						data_grav="&nbsp;"
						hora_grav="&nbsp;"
									
				else
						va_pt=RS3("PE_Teste")
						va_pp=RS3("PE_Prova")
						va_t1=RS3("VA_Teste1")
						va_t2=RS3("VA_Teste2")
						va_mt=RS3("MD_Teste")
						va_p1=RS3("VA_Prova1")
						va_p2=RS3("VA_Simul")
						va_p3=RS3("VA_Prova2")
						va_mp=RS3("MD_Prova")
						va_m1=RS3("VA_Media1")
						va_bon=RS3("VA_Bonus")
						va_m2=RS3("VA_Media2")
						va_rec=RS3("VA_Rec")
						va_m3=RS3("VA_Media3")
						data_grav=RS3("DA_Ult_Acesso")
						hora_grav=RS3("HO_ult_Acesso")
						
				end if
if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then										
	showapr="n"
	showprova="n"
else
		showapr=showapr_ctrl
		showprova=showprova_ctrl
end if
						
				if hora_grav="&nbsp;" then
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
									
				if data_grav="&nbsp;" then
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
				
				
				'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
				'data_inicio=""
				'va_faltas=""
				'		end if
				
				
				%>
        <tr class="<%response.Write(cor)%>"> 
          <td class="tabela"  width="235"> 
            <%response.Write("&nbsp;"&no_materia)%>
          </td>
          <td class="tabela"  width="40"> <div align="center"> 
              <%
											if showapr="s" then							
											response.Write("&nbsp;"&va_t1)
											else
							response.Write("&nbsp;")
							End IF							
											%>
            </div></td>
          <td class="tabela"  width="40"> <div align="center"> 
              <%
											if showapr="s" then					
											response.Write("&nbsp;"&va_t2)
											else
							response.Write("&nbsp;")
							End IF
											%>
            </div></td>
          <td class="tabela"  width="40"
				> <div align="center"> 
              <%
											if showapr="s" then					
											response.Write("&nbsp;"&va_mt)
											else
							response.Write("&nbsp;")
							End IF
											%>
            </div></td>
<!--          <td class="tabela"  width="37"
				> <div align="center"> 
              <%
											if showapr="s" then					
											'response.Write("&nbsp;"&va_pt)
											else
							'response.Write("&nbsp;")
							End IF
											%>
            </div></td>-->
          <td class="tabela"  width="40"
				> <div align="center"> 
              <%
											if showprova="s" then					
											response.Write("&nbsp;"&va_p1)
											else
							response.Write("&nbsp;")
							End IF
											%>
            </div></td>
          <td class="tabela"  width="40"
				> <div align="center"> 
              <%
											if showprova="s" then					
											response.Write("&nbsp;"&va_p2)
											else
							response.Write("&nbsp;")
							End IF
											%>
            </div></td>
          <td class="tabela"  width="40"
				> <div align="center"> 
              <%
											if showprova="s" then					
											response.Write("&nbsp;"&va_p3)
											else
							response.Write("&nbsp;")
							End IF
											%>
            </div></td>
          <td class="tabela"  width="40"
				> <div align="center"> 
              <%
											if showprova="s" and materia<>"LP" then					
											response.Write("&nbsp;"&va_mp)
											else
											response.Write("&nbsp;")
											End IF
											%>
            </div></td>
<!--          <td class="tabela"  width="37"
				> <div align="center"> 
              <%
											if showprova="s" then					
											'response.Write("&nbsp;"&va_pp)
											else
							'response.Write("&nbsp;")
							End IF
											%>
            </div></td>-->
          <td class="tabela"  width="40"
				> <div align="center"> 
              <%
											if showprova="s" AND showapr="s" then					
											response.Write("&nbsp;"&va_m1)
											else
							response.Write("&nbsp;")
							End IF
											%>
            </div></td>
          <td class="tabela"  width="40"
				> <div align="center"> 
              <%
											if showprova="s" AND showapr="s" then					
											response.Write("&nbsp;"&va_bon)
											else
							response.Write("&nbsp;")
							End IF
											%>
            </div></td>
          <td class="tabela"  width="40"
				> <div align="center"> 
              <%
											if showprova="s" AND showapr="s" then					
											response.Write("&nbsp;"&va_m2)
											else
							response.Write("&nbsp;")
							End IF
											%>
            </div></td>
          <td class="tabela"  width="40"
				> <div align="center"> 
              <%
											if showprova="s" AND showapr="s" then					
											response.Write("&nbsp;"&va_rec)
											else
							response.Write("&nbsp;")
							End IF
											%>
            </div></td>
          <td class="tabela"  width="40"
				> <div align="center"> 
              <%
											if showprova="s" AND showapr="s" then					
											response.Write("&nbsp;"&va_m3)
											else
							response.Write("&nbsp;")
							End IF
											%>
            </div></td>
          <td class="tabela"  width="235"
				> 
            <div align="center"> 
              <%
											if showprova="s" AND showapr="s" then							
											response.Write("&nbsp;"&data_inicio)
											else
							response.Write("&nbsp;")
							End IF
											%>
            </div></td>
        </tr>
        <%
			peso_acumula=0
			acumula_m1=0
			m2_ac=0			
			m3_ac=0
			m1_exibe=0
			m2_exibe=0
			m3_exibe=0
			
			
				while not RS1a.EOF
				
						materia_fil=RS1a("CO_Materia")
					
								Set RS1b = Server.CreateObject("ADODB.Recordset")
								SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia_fil&"'"
								RS1b.Open SQL1b, CON0
								
						no_materia_fil=RS1b("NO_Materia")
						
						Set RSpa = Server.CreateObject("ADODB.Recordset")
						SQLpa= "SELECT * FROM TB_Programa_Aula where CO_Materia='"&materia_fil&"' AND CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"'"
						RSpa.Open SQLpa, CON0
												
						nu_peso_fil=RSpa("NU_Peso")						
							
								Set RSnFIL = Server.CreateObject("ADODB.Recordset")
								Set RS3 = Server.CreateObject("ADODB.Recordset")
								SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia_Principal = '"& materia &"' AND CO_Materia = '"& materia_fil &"' and NU_Periodo="&periodo_check
								Set RS3 = CON_N.Execute(SQL_N)
						if RS3.EOF then
								va_pt="&nbsp;"
								va_pp="&nbsp;"
								va_t1="&nbsp;"
								va_t2="&nbsp;"
								va_mt="&nbsp;"
								va_p1="&nbsp;"
								va_p2="&nbsp;"
								va_p3="&nbsp;"
								va_mp="&nbsp;"
								va_m1=0
								va_bon="&nbsp;"
								va_m2=0
								va_rec="&nbsp;"
								va_m3=0
								data_grav="&nbsp;"
								hora_grav="&nbsp;"
											
						else
								va_pt=RS3("PE_Teste")
								va_pp=RS3("PE_Prova")
								va_t1=RS3("VA_Teste1")
								va_t2=RS3("VA_Teste2")
								va_mt=RS3("MD_Teste")
								va_p1=RS3("VA_Prova1")
								va_p2=RS3("VA_Simul")
								va_p3=RS3("VA_Prova2")
								va_mp=RS3("MD_Prova")
								va_m1=RS3("VA_Media1")
								va_bon=RS3("VA_Bonus")
								va_m2=RS3("VA_Media2")
								va_rec=RS3("VA_Rec")
								va_m3=RS3("VA_Media3")
								data_grav=RS3("DA_Ult_Acesso")
								hora_grav=RS3("HO_ult_Acesso")
								
						end if
						if isnull(va_m1) or va_m1="" then
						va_m1=0
						end if

						if isnull(va_m2) or va_m2="" then
						va_m2=0
						end if
						
						if isnull(va_m3) or va_m3="" then
						va_m3=0
						end if												
						
							peso_acumula=peso_acumula+nu_peso_fil
							m1_ac=m1_ac+(va_m1*nu_peso_fil)	
							m2_ac=m2_ac+(va_m2*nu_peso_fil)
							m3_ac=m3_ac+(va_m3*nu_peso_fil)
																										
								
						if hora_grav="&nbsp;" then
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
											
						if data_grav="&nbsp;" then
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
						
						
						'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
						'data_inicio=""
						'va_faltas=""
						'		end if
						
						
						%>
        <tr class="<%response.Write(cor)%>"> 
          <td class="tabela"  width="235">&nbsp;&nbsp;&nbsp; 
            <%response.Write("&nbsp;"&no_materia_fil)%>
          </td>
          <td class="tabela"  width="40"> <div align="center"> 
              <%
													if showapr="s" then							
													response.Write("&nbsp;"&va_t1)
													else
							response.Write("&nbsp;")
							End IF							
													%>
            </div></td>
          <td class="tabela"  width="40"> <div align="center"> 
              <%
													if showapr="s" then					
													response.Write("&nbsp;"&va_t2)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showapr="s" then					
													response.Write("&nbsp;"&va_mt)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
<!--          <td class="tabela"  width="37"
						> <div align="center"> 
              <%
													if showapr="s" then					
													'response.Write("&nbsp;"&va_pt)
													else
							'response.Write("&nbsp;")
							End IF
													%>
            </div></td>-->
          <td class="tabela"  width="40"
						> <div align="center">&nbsp;</div></td>
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showprova="s" then					
													response.Write("&nbsp;"&va_p2)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showprova="s" then					
													response.Write("&nbsp;"&va_p3)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showprova="s" and materia<>"LP" then					
													response.Write("&nbsp;"&va_mp)
													else
													response.Write("&nbsp;")
													End IF
													%>
            </div></td>
<!--          <td class="tabela"  width="37"
						> <div align="center"> 
              <%
													if showprova="s" then					
													'response.Write("&nbsp;"&va_pp)
													else
							'response.Write("&nbsp;")
							End IF
													%>
            </div></td>-->
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showprova="s" AND showapr="s" then					
													response.Write("&nbsp;"&va_m1)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showprova="s" AND showapr="s" then					
													response.Write("&nbsp;"&va_bon)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showprova="s" AND showapr="s" then					
													response.Write("&nbsp;"&va_m2)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showprova="s" AND showapr="s" then					
													response.Write("&nbsp;"&va_rec)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
          <td class="tabela"  width="40"
						> <div align="center"> 
              <%
													if showprova="s" AND showapr="s" then					
													response.Write("&nbsp;"&va_m3)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
          <td class="tabela"  width="235"
						> 
            <div align="center"> 
              <%
													if showprova="s" AND showapr="s" then							
													response.Write("&nbsp;"&data_inicio)
													else
							response.Write("&nbsp;")
							End IF
													%>
            </div></td>
        </tr>
        <%
				RS1a.movenext
				wend
							m1_exibe=m1_ac/peso_acumula
							m2_exibe=m2_ac/peso_acumula
							m3_exibe=m3_ac/peso_acumula
							
								decimo = m1_exibe - Int(m1_exibe)
									If decimo >= 0.5 Then
										nota_arredondada = Int(m1_exibe) + 1
										m1_exibe=nota_arredondada
									Else
										nota_arredondada = Int(m1_exibe)
										m1_exibe=nota_arredondada					
									End If
								m1_exibe= formatNumber(m1_exibe,0)
								
								decimo = m2_exibe - Int(m2_exibe)
									If decimo >= 0.5 Then
										nota_arredondada = Int(m2_exibe) + 1
										m2_exibe=nota_arredondada
									Else
										nota_arredondada = Int(m2_exibe)
										m2_exibe=nota_arredondada					
									End If
								m2_exibe= formatNumber(m2_exibe,0)
								
								decimo = m3_exibe - Int(m3_exibe)
									If decimo >= 0.5 Then
										nota_arredondada = Int(m3_exibe) + 1
										m3_exibe=nota_arredondada
									Else
										nota_arredondada = Int(m3_exibe)
										m3_exibe=nota_arredondada					
									End If
								m3_exibe= formatNumber(m3_exibe,0)							
				
				%>
        <tr class="tabela"> 
          <td class="tabela"  width="235">&nbsp;&nbsp;&nbsp; M&eacute;dia</td>
          <td class="tabela"  width="40"> <div align="center">&nbsp;</div></td>
          <td class="tabela"  width="40"> <div align="center"> &nbsp;</div></td>
          <td class="tabela"  width="40"
				> <div align="center">&nbsp;</div></td>
<!--          <td class="tabela"  width="37"
				> <div align="center">&nbsp;</div></td>-->
          <td class="tabela"  width="40"
				> <div align="center">&nbsp; </div></td>
          <td class="tabela"  width="40"
				> <div align="center">&nbsp; </div></td>
          <td class="tabela"  width="40"
				> <div align="center">&nbsp; </div></td>
          <td class="tabela"  width="40"
				> <div align="center">&nbsp; </div></td>
<!--          <td class="tabela"  width="37"
				> <div align="center"> &nbsp;</div></td>-->
          <td class="tabela"  width="40"
				> <div align="center"> 
                            <%
							if showprova="s" AND showapr="s" then
							response.Write("&nbsp;"&m1_exibe)
							else
							response.Write("&nbsp;")
							end if%>
            </div></td>
          <td class="tabela"  width="40"
				> <div align="center">&nbsp; </div></td>
          <td class="tabela"  width="40"
				> <div align="center"> 
                            <%
							if showprova="s" AND showapr="s" then
							response.Write("&nbsp;"&m2_exibe)
							else
							response.Write("&nbsp;")
							end if%>
            </div></td>
          <td class="tabela"  width="40"
				> <div align="center"> &nbsp;</div></td>
          <td class="tabela"  width="40"
				> <div align="center"> 
                            <%
							if showprova="s" AND showapr="s" then
							response.Write("&nbsp;"&m3_exibe)
							else
							response.Write("&nbsp;")
							end if%>
            </div></td>
          <td class="tabela"  width="235"
				> 
            <div align="center">&nbsp; </div></td>
        </tr>
        <%
			end if
			end if

		RSprog.MOVENEXT
		wend
		%>
        <tr valign="bottom"> 
          <td  height="20" colspan="20" class="tabela" 
		> <div align="right"><font class="tabela_rodape"> 
            <%
													if showprova="s" then					
													'response.Write("&nbsp;"&va_p1)
													else
							'response.Write("&nbsp;")
							End IF
													%>
                    <% if etapa=6 or etapa=7 or etapa=8 or etapa=9 then
				Response.Write("T-Teste, PR-Prova, S - Simulado, MT=(T1+T2)/2, MP=(PR1+S), M3=((MTx1)+(MPx2))/3. <br>Para a Disciplina Português PR2 = Redação e M3=((MTx1)+(MPx2)+(PR2x2))/5.")			
		else
			Response.Write("T-Teste, MT–Média dos Testes, PR-Prova, MP–Média das Provas e M-M&eacute;dia")
		End if%></font></div></td>
        </tr>
      </table>
      <%
elseif notaFIL="TB_NOTA_C" then
%>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="241" class="tabela"> 
            <div align="left">Disciplina</div></td>
          <td width="36"  class="tabela"> <div align="center">T1</div></td>
          <td width="36"  class="tabela"> <div align="center">T2</div></td>
          <td width="36"  class="tabela"> <div align="center">T3</div></td>
          <td width="36"  class="tabela"> <div align="center">T4</div></td>
          <td width="36" class="tabela"> <div align="center">MT</div></td>
<!--          <td width="33" class="tabela">&nbsp;</td>-->
          <td width="36"  class="tabela"> <div align="center">PR1</div></td>
          <td width="36"  class="tabela"> <div align="center">PR2</div></td>
          <td width="36" class="tabela"> <div align="center"> MP</div></td>
<!--          <td width="33" class="tabela">&nbsp;</td>-->
          <td width="36" class="tabela"> <div align="center">M1</div></td>
          <td width="36"  class="tabela"> <div align="center">Bon</div></td>
          <td width="36"  class="tabela"> <div align="center">M2</div></td>
          <td width="36"  class="tabela"> <div align="center">Rec</div></td>
          <td width="36"  class="tabela"> <div align="center">M3</div></td>
          <td width="241" class="tabela"> 
          <div align="center">Data/Hora</div></td>
        </tr>
       <!-- <tr>
          <td width="225" class="tabela">&nbsp;</td> 
          <td width="33"  class="tabela"> <div align="center">N</div></td>
          <td width="33"  class="tabela"> <div align="center">N</div></td>
          <td width="33"  class="tabela"> <div align="center">N</div></td>
          <td width="33"  class="tabela"> <div align="center">N</div></td>
          <td width="66" class="tabela"> <div align="center">M</div></td>
          <td width="33" class="tabela"> <div align="center">P</div></td>
          <td width="33"  class="tabela"> <div align="center">N</div></td>
          <td width="33"  class="tabela"> <div align="center">N</div></td>
          <td width="66" class="tabela"> <div align="center"> M</div></td>
          <td width="33" class="tabela"> <div align="center">P</div></td>
          <td width="33" class="tabela"> <div align="center">M</div></td>
          <td width="33"  class="tabela"> <div align="center">N</div></td>
          <td width="33"  class="tabela"> <div align="center">M</div></td>
          <td width="33"  class="tabela"> <div align="center">N</div></td>
          <td width="33"  class="tabela"> <div align="center">M</div></td>
          <td width="191" class="tabela">&nbsp;</td>
        </tr>-->
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
  cor = "tabela" 
 else cor ="tabela"
  end if
	
		Set RSnFIL = Server.CreateObject("ADODB.Recordset")
	  	Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' and NU_Periodo="&periodo_check
		Set RS3 = CON_N.Execute(SQL_N)
if RS3.EOF then
		va_pt="&nbsp;"
		va_pp="&nbsp;"
		va_t1="&nbsp;"
		va_t2="&nbsp;"
		va_t3="&nbsp;"
		va_t4="&nbsp;"
		va_mt="&nbsp;"
		va_p1="&nbsp;"
		va_p2="&nbsp;"
		va_mp="&nbsp;"
		va_m1="&nbsp;"
		va_bon="&nbsp;"
		va_m2="&nbsp;"
		va_rec="&nbsp;"
		va_m3="&nbsp;"
		data_grav="&nbsp;"
		hora_grav="&nbsp;"
					
else
		va_pt=RS3("PE_Teste")
		va_pp=RS3("PE_Prova")
		va_t1=RS3("VA_Teste1")
		va_t2=RS3("VA_Teste2")
		va_t3=RS3("VA_Teste3")
		va_t4=RS3("VA_Teste4")
		va_mt=RS3("MD_Teste")
		va_p1=RS3("VA_Prova1")
		va_p2=RS3("VA_Prova2")
		va_mp=RS3("MD_Prova")
		va_m1=RS3("VA_Media1")
		va_bon=RS3("VA_Bonus")
		va_m2=RS3("VA_Media2")
		va_rec=RS3("VA_Rec")
		va_m3=RS3("VA_Media3")
		data_grav=RS3("DA_Ult_Acesso")
		hora_grav=RS3("HO_ult_Acesso")
		
end if

if curso=1 and etapa<6 and (materia="ARTC" or materia="EART" or materia="EFIS" or materia="INGL") then										
	showapr="n"
	showprova="n"
else
		showapr=showapr_ctrl
		showprova=showprova_ctrl
end if									
		
if hora_grav="&nbsp;" then
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
					
if data_grav="&nbsp;" then
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


'		if (va_apr1=0 OR va_apr1="0" OR va_apr1="")and (va_apr2=0 OR va_apr2="0"OR va_apr2="")and (va_apr3=0 OR va_apr3="0" OR va_apr3="")and (va_apr4=0 OR va_apr4="0" OR va_apr4="")and (va_apr5=0 OR va_apr5="0" OR va_apr5="")and (va_apr6=0 OR va_apr6="0" OR va_apr6="")and (va_apr7=0 OR va_apr7="0" OR va_apr7="")and (va_apr8=0 OR va_apr8="0" OR va_apr8="")and (va_sapr=0 OR va_sapr="0" OR va_sapr="" OR ISNULL(va_sapr))  then
'data_inicio=""
'va_faltas=""
'		end if


%>
        <tr class="<%response.Write(cor)%>"> 
          <td class="tabela" width="241"> 
            <%response.Write("&nbsp;"&no_materia)%>
          </td>
          <td class="tabela" width="36"> <div align="center"> 
              <%
							if showapr="s" then							
							response.Write("&nbsp;"&va_t1)
							else
							response.Write("&nbsp;")
							End IF							
							%>
            </div></td>
          <td class="tabela" width="36"> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write("&nbsp;"&va_t2)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="36"
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write("&nbsp;"&va_t3)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="36"
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write("&nbsp;"&va_t4)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="36"
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write("&nbsp;"&va_mt)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="36"
> <div align="center"> 
              <%
							if showapr="s" then					
							response.Write("&nbsp;"&va_pt)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
<!--          <td class="tabela" width="33"
> <div align="center"> 
              <%
							if showprova="s" then					
							'response.Write("&nbsp;"&va_p1)
							else
							'response.Write("&nbsp;")
							End IF
							%>
            </div></td>-->
          <td class="tabela" width="36"
> <div align="center"> 
              <%
							if showprova="s" then					
							response.Write("&nbsp;"&va_p2)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="36"
> <div align="center"> 
              <%
							if showprova="s" then					
							response.Write("&nbsp;"&va_mp)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
<!--          <td class="tabela" width="33"
> <div align="center"> 
              <%
							if showprova="s" then					
							'response.Write("&nbsp;"&va_pp)
							else
							'response.Write("&nbsp;")
							End IF
							%>
            </div></td>-->
          <td class="tabela" width="36"
> <div align="center"> 
              <%
							if showprova="s" AND showapr="s" then					
							response.Write("&nbsp;"&va_m1)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="36"
> <div align="center"> 
              <%
							if showprova="s" AND showapr="s" then					
							response.Write("&nbsp;"&va_bon)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="36"
> <div align="center"> 
              <%
							if showprova="s" AND showapr="s" then					
							response.Write("&nbsp;"&va_m2)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="36"
> <div align="center"> 
              <%
							if showprova="s" AND showapr="s" then					
							response.Write("&nbsp;"&va_rec)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="36"
> <div align="center"> 
              <%
							if showprova="s" AND showapr="s" then					
							response.Write("&nbsp;"&va_m3)
							else
							response.Write("&nbsp;")
							End IF
							%>
            </div></td>
          <td class="tabela" width="241"
> <div align="center"> 
            <%
							if showprova="s" AND showapr="s" then							
							response.Write("&nbsp;"&data_inicio)
							else
							response.Write("&nbsp;")
							End IF
							%></font>
          </td>
        </tr>
        <%check=check+1
RSprog.MOVENEXT
wend

%>
        <tr valign="bottom"> 
          <td height="20" colspan="21" class="tabela" 
> <div align="right"><font class="tabela_rodape"> T-Teste, MT–Soma dos Testes, PR-Prova, MP=(PR1+PR2)/2 e M3=((MTx1)+(MPx1))/2</font></div></td>
        </tr>
      </table>
<%
end if%>					  	
	 </td>
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