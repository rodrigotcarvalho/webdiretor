<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/boletim.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<html>
<head>
<title>Web Diretor</title>
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


<%opt = REQUEST.QueryString("obr")

nta= REQUEST.QueryString("n")
obr=split(opt,"_")
co_materia= obr(0)
unidade = obr(1)
curso = obr(2)
co_etapa = obr(3)
turma = obr(4)
co_prof = obr(6)



		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0

	no_materia= GeraNomes("D",co_materia,variavel2,variavel3,variavel4,variavel5,CON0,outro) 	
	no_unidade= GeraNomes("U",unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)
	no_curso=GeraNomes("C",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro) 	
	no_etapa=GeraNomes("E",curso,co_etapa,variavel3,variavel4,variavel5,CON0,outro) 	



ano_letivo = session("ano_letivo")
obr=unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&periodo&"_"&ano_letivo

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
		

		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2


		Set RSTB = Server.CreateObject("ADODB.Recordset")
		CONEXAOTB = "Select * from TB_Da_Aula WHERE NU_Unidade="& unidade &" AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		Set RSTB = CON2.Execute(CONEXAOTB)
		
tb_nota= RSTB("TP_Nota")

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

cep3=left(un_cep,5)
cep4 = right(un_cep,3)

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
<table width="1000" border="0" align="center" cellspacing="0" class="tb_corpo"
>
  <tr> 
    <td width="203" height="15" bgcolor="#FFFFFF"><div align="center"><img src="../../../../img/logo_preto.gif"> 
      </div></td>
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
          <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>MAP&Atilde;O 
            DE RESULTADO POR DISCIPLINA</strong></font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="12" colspan="2" bgcolor="#EEEEEE" class="linhaTopoL"><table width="900" border="0" cellspacing="0">
        <tr class="tabela"> 
          <td width="80" height="12" align="right" bgcolor="#EEEEEE" > 
            <div align="right"> <strong>Ano 
              Letivo:</strong></div></td>
          <td width="180" height="12" bgcolor="#EEEEEE" > 
            <%response.Write(ano_letivo)%>
            </td>
          <td width="60">&nbsp;</td>
          <td width="200">&nbsp;</td>
          <td width="380">&nbsp;</td>
        </tr>
        <tr class="tabela"> 
          <td width="80" height="12" bgcolor="#EEEEEE"> 
            <div align="right"><strong>Curso:</strong></div></td>
          <td width="180" height="12" bgcolor="#EEEEEE"> 
            <%
response.Write(no_curso)%>
</td>
          <td width="60" height="12" bgcolor="#EEEEEE"> 
            <div align="right">
<strong>Etapa:</strong></div></td>
          <td width="200">
            <%response.Write(no_etapa)%>
</td>
          <td width="380">&nbsp;</td>
        </tr>
        <tr class="tabela"> 
          <td width="80" height="12" bgcolor="#EEEEEE"> 
            <div align="right"><strong><strong>Turma:</strong></strong></div></td>
          <td width="180" height="12" bgcolor="#EEEEEE"> 
            <%response.Write(turma)%>
            </td>
          <td width="60" height="12" bgcolor="#EEEEEE"> 
            <div align="right"><strong>Disciplina:</strong></div></td>
          <td width="200">
            <%response.Write(no_materia)%>
            </td>
          <td width="380">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="10" colspan="2" bgcolor="#EEEEEE"> </td>
  </tr>
  <tr> 
    <td colspan="2"><%


		Set RSNN = Server.CreateObject("ADODB.Recordset")
		CONEXAONN = "Select * from TB_Programa_Aula WHERE CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa&"' order by NU_Ordem_Boletim"
		Set RSNN = CON0.Execute(CONEXAONN)


	if tb_nota ="TB_NOTA_A" then
	CAMINHOn = CAMINHO_na
	
	elseif tb_nota="TB_NOTA_B" then
		CAMINHOn = CAMINHO_nb
	
	elseif tb_nota ="TB_NOTA_C" then
			CAMINHOn = CAMINHO_nc
	else
			response.Write("ERRO")
	end if
	

	call boletim_escolar (unidade,curso,co_etapa,turma,CAMINHOn,tb_nota,co_materia,"MRDI")


%></td>
  <tr> 
    <td class="linhaTopoL"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Sistema 
      Diretor - WEB ACAD&Ecirc;MICO</font></td>
    <td class="linhaTopoR"> <div align="right"> 
        <font size="1" face="Verdana, Arial, Helvetica, sans-serif">Impresso 
          em 
          <%response.Write(data &" às "&horario)%>
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
pasta=arPath(seleciona)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>