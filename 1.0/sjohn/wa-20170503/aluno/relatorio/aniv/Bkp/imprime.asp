<%On Error Resume Next%>
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
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->

<%opt = request.QueryString("obr")
obr = split(opt,"_")

unidade = obr(0)
ma = obr(1)




		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0

		Set CONa = Server.CreateObject("ADODB.Connection") 
		ABRIRa = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONa.Open ABRIRa
		
		Set CON7 = Server.CreateObject("ADODB.Connection") 
		ABRIR7 = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON7.Open ABRIR7

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


	%>

<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('window.print()')"> 
<br>
<table width="950" border="0" align="center" cellspacing="0" class="tb_corpo"
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
          <td colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            &nbsp; </font></td>
        </tr>
        <tr> 
          <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>LISTA 
            DE ANIVERSARIANTES DO M&Ecirc;S</strong></font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="12" colspan="2" bgcolor="#EEEEEE" class="linhaTopoL"><table width="900" border="0" cellspacing="0">
        <tr class="tabela"> 
          <td width="80" height="12" align="right" bgcolor="#EEEEEE" > 
            <div align="right"> <strong>Ano 
              Letivo:</strong></div></td>
          <td width="200" height="12" bgcolor="#EEEEEE" >
            <%response.Write(ano_letivo)%>
            </td>
          <td width="40">&nbsp;</td>
          <td width="200">&nbsp;</td>
          <td width="380">&nbsp;</td>
        </tr>
        <tr class="tabela"> 
          <td width="80" height="12" bgcolor="#EEEEEE"> 
            <div align="right"><strong>M&ecirc;s:</strong></div></td>
          <td width="200" height="12" bgcolor="#EEEEEE">
            <%
select case ma
 case 1 
 mes_a = "janeiro"
 case 2 
 mes_a = "fevereiro"
 case 3 
 mes_a = "março"
 case 4
 mes_a = "abril"
 case 5
 mes_a = "maio"
 case 6 
 mes_a = "junho"
 case 7
 mes_a = "julho"
 case 8 
 mes_a = "agosto"
 case 9 
 mes_a = "setembro"
 case 10 
 mes_a = "outubro"
 case 11 
 mes_a = "novembro"
 case 12 
 mes_a = "dezembro"
end select					  
response.Write(mes_a)%>
</td>
          <td width="40" height="12" bgcolor="#EEEEEE"> 
            <div align="right"> </div></td>
          <td width="200">&nbsp; </td>
          <td width="380">&nbsp;</td>
        </tr>
        <tr class="tabela"> 
          <td width="80" height="12" bgcolor="#EEEEEE"> 
            <div align="right"><strong></strong></div></td>
          <td width="200" height="12" bgcolor="#EEEEEE"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
            </font><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
          <td width="40" height="12" bgcolor="#EEEEEE"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
          <td width="200">&nbsp;</td>
          <td width="380">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="10" colspan="2" bgcolor="#EEEEEE"> </td>
  </tr>
  <tr> 
    <td colspan="2">
      <table width="950" border="0" align="left" cellspacing="0" bordercolor="#000000">
        <tr> 
          <td width="60" class="tabelaTit"> 
            <div align="center">Anivers&aacute;rio</div></td>
          <td width="420" class="tabelaTit"> 
            <div align="center"><strong>Nome</strong></div></td>
          <td width="98" height="40" class="tabelaTit"> 
            <div align="center">Far&aacute;</div></td>
          <td width="98" class="tabelaTit"> 
            <div align="center">Matr&iacute;cula</div></td>
          <td width="98" class="tabelaTit"> 
            <div align="center">Curso</div></td>
          <td width="98" class="tabelaTit"> 
            <div align="center">Etapa</div></td>
          <td width="98" class="tabelaTit"> 
            <div align="center">Turma</div></td>
        </tr>
        <%
		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Contatos where TP_Contato='ALUNO' order by day(DA_Nascimento_Contato),Year(DA_Nascimento_Contato)"
		RS7.Open SQL7, CON7

While not RS7.EOF
cod_al=RS7("CO_Matricula")	  
'matriculas=matriculas&"_"&mat

		Set RSa0 = Server.CreateObject("ADODB.Recordset")
		SQLa0 = "SELECT * FROM TB_Matriculas where NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod_al &" AND NU_Unidade="&unidade
		RSa0.Open SQLa0, CONa
		
		Set RSa = Server.CreateObject("ADODB.Recordset")
		SQLa = "SELECT * FROM TB_Alunos where CO_Matricula ="& cod_al
		RSa.Open SQLa, CONa		

IF RSa0.EOF	Then
RS7.Movenext
Else
nascimento = RS7("DA_Nascimento_Contato")		
if nascimento="" or isnull(nascimento) then
RS7.Movenext
else
nasceu=split(nascimento,"/")
if nasceu(1)= ma then
nome_al=RSa("NO_Aluno")


		Set RSa1 = Server.CreateObject("ADODB.Recordset")
		SQLa1 = "SELECT * FROM TB_Matriculas where NU_Ano="& ano_letivo &" AND CO_Matricula="&cod_al
		RSa1.Open SQLa1, CONa
		
curso=RSa1("CO_Curso")
etp =RSa1("CO_Etapa")
tm =RSa1("CO_Turma")	
		
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
			
no_curso = RS1("NO_Curso")
		
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Etapa where CO_Curso ='"& curso &"' AND CO_Etapa ='"& etp &"'"
		RS2.Open SQL2, CON0
		
no_etp = RS2("NO_Etapa")
%>
        <tr> 
          <td width="60" class="tabela"> 
            <div align="center"> 
              <%response.Write(nasceu(0))%>
            </div></td>
          <td width="420"  class="tabela"> 
            <div align="left"> 
              <%response.Write(nome_al)%>
            </div></td>
          <td width="98"  class="tabela"> 
            <div align="center"> 
              <%call aniversario(nasceu(2),nasceu(1),nasceu(0)) %>
            </div></td>
          <td width="98"  class="tabela"> 
            <div align="center"> 
              <%response.Write(cod_al)%>
            </div></td>
          <td width="98"  class="tabela"> 
            <div align="center"> 
              <%response.Write(no_curso)%>
            </div></td>
          <td width="98"  class="tabela"> 
            <div align="center"> 
              <%response.Write(no_etp)%>
            </div></td>
          <td width="98"  class="tabela"> 
            <div align="center"> 
              <%response.Write(tm)%>
            </div></td>
        </tr>
        <%
RS7.Movenext
Else
RS7.Movenext
end if
end if
end if
Wend
%>
      </table></td>
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