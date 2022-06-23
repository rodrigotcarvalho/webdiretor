<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/parametros.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes7.asp"-->
<!--#include file="../../../../inc/boletim.asp"-->
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

<%'opt = REQUEST.QueryString("obr")
'dados_opt= split(opt, "?" )
cod= REQUEST.QueryString("obr")
'periodo_check= dados_opt(1)

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
<% if unidade = 1 then
	logo = "logo_niteroi_preto.gif"
else
	logo = "logo_arariboia_preto.gif"
end if%>    
<div align="center"><img src="../../../../img/<%response.Write(logo)%>"> </div></td>
    <td width="828" bgcolor="#FFFFFF"><table width="100%" border="0" align="right" cellspacing="0">
        <tr> 
          <td width="74%" rowspan="2"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>UNIDADE 
            <%
			no_unidade= ucase(no_unidade)
			response.Write("&nbsp;"&no_unidade)
			%>
            </strong></font></td>
          <td width="26%" height="8" class="linhaBaixo">&nbsp; </td>
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
    <td colspan="2"> </td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td colspan="2"><div align="right">
                    <%	
tb_nota=tabela_nota(ano_letivo,unidade,curso,etapa,turma,"tb",0)
caminho_nota=tabela_nota(ano_letivo,unidade,curso,etapa,turma,"cam",0)
if tb_nota="erro" or caminho_nota ="erro" then
%>
%>
                    <div align="center"> <%response.Write("<br><br><br><br><br>Não existe Boletim para este aluno!")%>
        </div>
                    <%
else
	call boletim_escolar (unidade,curso,etapa,turma,caminho_nota,tb_nota,cod,"WAI")

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
              Diretor - WEB ACAD&Ecirc;MICO</font></td>
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
response.redirect("../../../../inc/erro.asp")
end if
%>