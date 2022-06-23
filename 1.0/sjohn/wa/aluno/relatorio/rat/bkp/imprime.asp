<%On Error Resume Next%>
<html>
<head>
<title>Web Acad&ecirc;mico</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">

<script type="text/javascript" src="../js/global.js"></script>
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
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->

<%opt = REQUEST.QueryString("obr")

'obr=split(opt,"_")

'unidade = obr(0)
'curso = obr(1)
'co_etapa = obr(2)
'turma = obr(3)
'periodo = obr(4)

alunos=0
sxF=0
sxM=0

obr = split(opt,"-")


unidade = obr(0)
curso = obr(1)
co_etapa = obr(2)
turma = obr(3)
if co_etapa = "f0"then
co_etapa=0
elseif co_etapa = "f1" or co_etapa = "m1"then
co_etapa=1
elseif co_etapa = "f2" or co_etapa = "m2"then
co_etapa = 2
elseif co_etapa = "f3" or co_etapa = "m3"then
co_etapa = 3
elseif co_etapa = "f4" then
co_etapa = 4
elseif co_etapa = "f5" then
co_etapa = 5
elseif co_etapa = "f6" then
co_etapa = 6
elseif co_etapa = "f7" then
co_etapa = 7
elseif co_etapa = "f8" then
co_etapa = 8
elseif co_etapa = "f55" then
co_etapa = 55
elseif co_etapa = "f66" then
co_etapa = 66
elseif co_etapa = "f77" then
co_etapa = 77
end if
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
          <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>RELA&Ccedil;&Atilde;O 
            DE ALUNOS POR TURMA</strong></font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="12" colspan="2" bgcolor="#EEEEEE" class="linhaTopoL"><table width="900" border="0" cellspacing="0">
        <tr class="tabela"> 
          <td width="78" height="12" align="right" bgcolor="#EEEEEE" > <div align="right"> 
              <strong>Ano Letivo:</strong></div></td>
          <td width="198" height="12" bgcolor="#EEEEEE" > <%response.Write(ano_letivo)%> </td>
          <td width="38">&nbsp;</td>
          <td width="198">&nbsp;</td>
          <td width="86">&nbsp;</td>
          <td width="290">&nbsp;</td>
        </tr>
        <tr class="tabela"> 
          <td width="78" height="12" bgcolor="#EEEEEE"> <div align="right"><strong>Curso:</strong></div></td>
          <td width="198" height="12" bgcolor="#EEEEEE"> <%
response.Write(no_curso)%> </td>
          <td width="38" height="12" bgcolor="#EEEEEE"> <div align="right"> <strong>Etapa:</strong></div></td>
          <td width="198"> <%response.Write(no_etapa)%> </td>
          <td width="86">&nbsp;</td>
          <td width="290">&nbsp;</td>
        </tr>
        <tr class="tabela"> 
          <td width="78" height="12" bgcolor="#EEEEEE"> <div align="right"><strong><strong>Turma:</strong></strong></div></td>
          <td width="198" height="12" bgcolor="#EEEEEE"> <%response.Write(turma)%> </td>
          <td width="38" height="12" bgcolor="#EEEEEE"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
          <td width="198">&nbsp;</td>
          <td width="86">&nbsp;</td>
          <td width="290">&nbsp;</td>
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
          <td width="70" class="tabelaTit"> <div align="center">Matr&iacute;cula</div></td>
          <td width="250" class="tabelaTit"> <div align="left"><strong>Nome do 
              Aluno</strong></div></td>
          <td width="30" class="tabelaTit"> <div align="center">N&ordm;</div></td>
          <td width="300" class="tabelaTit"><div align="left">Respons&aacute;vel 
              Pedag&oacute;gico</div></td>
          <td width="300" height="40" class="tabelaTit"> <div align="center">Telefones 
              de Contato do Respons&aacute;vel Pedag&oacute;gico</div></td>
        </tr>
        <%  check = 2
nu_chamada_check = 1

	Set RSA = Server.CreateObject("ADODB.Recordset")
	CONEXAOA = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
	Set RSA = CON4.Execute(CONEXAOA)
 
 While Not RSA.EOF
nu_matricula = RSA("CO_Matricula")
nu_chamada = RSA("NU_Chamada")
alunos=alunos+1

  		Set RSA2 = Server.CreateObject("ADODB.Recordset")
		CONEXAOA2 = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
		Set RSA2 = CON4.Execute(CONEXAOA2)
  		NO_Aluno= RSA2("NO_Aluno")
		
		Set RSA3 = Server.CreateObject("ADODB.Recordset")
		CONEXAOA3 = "Select * from TB_Alunos WHERE CO_Matricula = "& nu_matricula
		Set RSA3 = CON6.Execute(CONEXAOA3)
tp_respp= RSA3("TP_Resp_Ped")
sx= RSA3("IN_Sexo")
if sx = "F" then
sxF=sxf+1
ELSE
sxM=sxM+1
end if
		
		Set RSA5 = Server.CreateObject("ADODB.Recordset")
		CONEXAOA5 = "Select * from TB_Contatos WHERE CO_Matricula = "& nu_matricula&" AND TP_Contato='"&tp_respp&"'"
		Set RSA5 = CON5.Execute(CONEXAOA5)
		
		no_respp= RSA5("NO_Contato")
		tel_respp= RSA5("NU_Telefones")

 if check mod 2 =0 then
  cor = "#F8FAFC" 
  else cor ="#F1F5FA"
  end if
  
if nu_chamada=nu_chamada_check then
nu_chamada_check=nu_chamada_check+1%>
        <tr> 
          <td width="70" class="tabela"> <div align="center"> 
              <%response.Write(nu_matricula)%>
            </div></td>
          <td width="250"  class="tabela"> <div align="left"> 
              <%response.Write(NO_Aluno)%>
            </div></td>
          <td width="30" class="tabela"> <div align="center"> 
              <%response.Write(nu_chamada)%>
            </div></td>
          <td width="300"  class="tabela"><%response.Write(no_respp)%> &nbsp;</td>
          <td width="300"  class="tabela"> <div align="center"> 
              <%response.Write(tel_respp)%>
              &nbsp;</div></td>
        </tr>
        <% 
else
While nu_chamada>nu_chamada_check
%>
        <tr> 
          <td width="70" bgcolor="#E4E4E4"  class="tabela"> <div align="center">&nbsp; 
            </div></td>
          <td width="250" bordercolor="#000000" bgcolor="#E4E4E4"  class="tabela"> 
            <div align="left">&nbsp;</div></td>
          <td width="30" bgcolor="#E4E4E4"  class="tabela"> <div align="center">&nbsp; 
            </div></td>
          <td width="300" bordercolor="#000000" bgcolor="#E4E4E4"  class="tabela">&nbsp;</td>
          <td width="300" bordercolor="#000000" bgcolor="#E4E4E4"  class="tabela"> 
            <div align="left"><strong> &nbsp;</strong></div></td>
        </tr>
        <%
nu_chamada_check=nu_chamada_check+1	 
wend	
%>
        <tr> 
          <td width="70"  class="tabela"> <div align="center"> 
              <%response.Write(nu_matricula)%>
            </div></td>
          <td width="250"  class="tabela"> <div align="left"> 
              <%response.Write(NO_Aluno)%>
            </div></td>
          <td width="30"  class="tabela"> <div align="center"> 
              <%response.Write(nu_chamada)%>
            </div></td>
          <td width="300"  class="tabela"><%response.Write(no_respp)%> &nbsp;</td>
          <td width="300"  class="tabela"> <div align="center"> 
              <%response.Write(tel_respp)%>
              &nbsp;</div></td>
        </tr>
        <%
 nu_chamada_check=nu_chamada_check+1	  
end if

	check = check+1
  RSA.MoveNext
  Wend 
%>
        <tr>
          <td colspan="5"  class="imprime">&nbsp;</td>
        </tr>
        <tr> 
          <td colspan="5"  class="tabela"><table width="350" border="0" cellspacing="0">
              <tr class="imprime"> 
                <td width="40%">Total de Alunos: 
                  <%response.Write(alunos)%> </td>
                <td width="30%">Eles: 
                  <%response.Write(sxM)%> </td>
                <td width="30%"> Elas: 
                  <%response.Write(sxF)%> </td>
              </tr>
            </table></td>
        </tr>
      </table></td>
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