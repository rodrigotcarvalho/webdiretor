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
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../../global/tabelas_escolas.asp"-->
<%obr = REQUEST.QueryString("obr")
obr=split(obr,"-")

unidade = obr(0)
curso = obr(1)
co_etapa = obr(2)
turma = obr(3)
periodo = obr(4)
avaliacao_form=obr(5)
nota=obr(6)
campo_check=avaliacao_form
escola=session("escola")

if nota = "TB_NOTA_A" Then		
		CAMINHOn = CAMINHO_na
		opcao="A"		
elseif nota = "TB_NOTA_B" Then
		CAMINHOn = CAMINHO_nb
		opcao="B"		
elseif nota = "TB_NOTA_C" Then
		CAMINHOn = CAMINHO_nc
		opcao="C"		
end if

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


		Set CON3 = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3

		Set CON4 = Server.CreateObject("ADODB.Connection")
		ABRIR4 = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4


	dados_tabela=verifica_dados_tabela(CAMINHOn,opcao,outro)
	dados_separados=split(dados_tabela,"#$#")
	ln_nom_cols=dados_separados(4)
	nm_vars=dados_separados(5)
	nm_bd=dados_separados(6)
	avaliacoes_nomes=split(ln_nom_cols,"#!#")
	verifica_avaliacoes=split(nm_vars,"#!#")
	avaliacoes=split(nm_bd,"#!#")

for i=3 to UBOUND(avaliacoes_nomes)
	j=i-2
	if avaliacoes(j)=avaliacao_form then
		nome_avaliacao=avaliacoes_nomes(i)
	end if
next




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
          <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>MAP&Atilde;O 
            DE AVALIA&Ccedil;&Atilde;O</strong></font></td>
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
          <td width="180" height="12" bgcolor="#EEEEEE"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.Write(turma)%>
            </font><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
          <td width="60" height="12" bgcolor="#EEEEEE"><div align="right"><strong>Avalia&ccedil;&atilde;o:</strong></div></td>
          <td width="200">
            <%response.Write(nome_avaliacao)%>
            </td>
          <td width="380">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="10" colspan="2" bgcolor="#EEEEEE"> </td>
  </tr>
  <tr> 
    <td colspan="2"> 
      <%

		Set RSNN = Server.CreateObject("ADODB.Recordset")
		CONEXAONN = "Select CO_Materia from TB_Programa_Aula WHERE CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa&"' order by NU_Ordem_Boletim"

		Set RSNN = CON0.Execute(CONEXAONN)
		
materia_nome_check="vazio"
nome_nota="vazio"
i=0
largura = 0
While not RSNN.eof
materia_nome= RSNN("CO_Materia")

if materia_nome=materia_nome_check then
RSNN.movenext
else

If Not IsArray(nome_nota) Then 
nome_nota = Array()
End if
' O if abaixo dá erro quando o vetor possui uma matéria chamada EFIS e verifica se outra chamada FIS existe no vetor
'If InStr(Join(nome_nota), materia_nome) = 0 Then
ReDim preserve nome_nota(UBound(nome_nota)+1)
nome_nota(Ubound(nome_nota)) = materia_nome
largura=largura+35

i=i+1
materia_nome_check=materia_nome

RSNN.movenext
end if
'end if
wend
larg=1008-(largura/i)

%>
      <table width="950" border="0" align="left" cellspacing="0" bordercolor="#000000">
        <tr> 
          <td width="17" class="tabelaTit"> <div align="center">N&ordm;</div></td>
          <td width="<%=larg%>" class="tabelaTit"> 
            <div align="center"><strong>Nome</strong></div></td>
          <td width="30" height="40" class="tabelaTit"> 
            <div align="center">Per</div></td>
          <%For k=0 To ubound(nome_nota)%>
          <td width="31" class="tabelaTit"> 
<% response.Write(nome_nota(k))%></td>
          <%
Next%>
        </tr>
        <%  check = 2
nu_chamada_check = 1

	Set RSA = Server.CreateObject("ADODB.Recordset")

	CONEXAOA = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
	Set RSA = CON4.Execute(CONEXAOA)
 
 While Not RSA.EOF
nu_matricula = RSA("CO_Matricula")
nu_chamada = RSA("NU_Chamada")

  		Set RSA2 = Server.CreateObject("ADODB.Recordset")
		CONEXAOA2 = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
		Set RSA2 = CON4.Execute(CONEXAOA2)
  		NO_Aluno= RSA2("NO_Aluno")

 if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if
  
if nu_chamada=nu_chamada_check then
nu_chamada_check=nu_chamada_check+1%>
        <tr> 
          <td width="17" class="tabela"> 
            <div align="center">
              <%response.Write(nu_chamada)%>
             </div></td>
          <td width="283"  class="tabela"> 
            <div align="left">
              <%response.Write(NO_Aluno)%>
</div></td>
          <td width="30"  class="tabela"> 
            <div align="center">
              <%

			Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo where NU_Periodo="&periodo
		RS4.Open SQL4, CON0


NO_Periodo= RS4("SG_Periodo")
		
							  
					  response.Write(NO_Periodo)%>
</div></td>
          <%For k=0 To ubound(nome_nota)

  		Set RS3 = Server.CreateObject("ADODB.Recordset")
		CONEXAO3 = "Select "& campo_check & " from "& nota & " WHERE NU_Periodo = "& periodo &" And CO_Materia = '"& nome_nota(k) &"' And CO_Matricula = "& nu_matricula
		Set RS3 = CON3.Execute(CONEXAO3)
  		
if RS3.EOF Then
%>
          <td width="31" bordercolor="#000000" class="tabelaTit"> 
            <div align="center">&nbsp;</div></td>
          <%else
nota_materia= RS3(""&campo_check&"")
if isnumeric(nota_materia) then
'nota_materia=nota_materia/10
nota_materia = formatNumber(nota_materia,1)   
end if
%>
          <td width="31" class="tabela"> 
  <%if  nota_materia="" or isnull(nota_materia) then%>
            <div align="center">&nbsp;</div>
<%else%>
            <div align="center">
              <%response.Write(nota_materia)%>
</div>
<%		  end IF%>
</td>
          <%end IF


NEXT%>
        </tr>
        <% 
else
While nu_chamada>nu_chamada_check
%>
        <tr> 
          <td width="17" bgcolor="#E4E4E4" class="tabela"> 
            <div align="center">
              <%response.Write(nu_chamada_check)%>
</div></td>
          <td width="283" bordercolor="#000000" bgcolor="#E4E4E4" class="tabela"> 
            <div align="left">&nbsp;</div></td>
          <td width="30" bordercolor="#000000" bgcolor="#E4E4E4" class="tabela"> 
            <div align="left"><strong>
              &nbsp;</strong></div></td>
          <%For k=0 To ubound(nome_nota)%>
          <td width="31" bordercolor="#000000" bgcolor="#E4E4E4" class="tabelaTit"> 
            <div align="center">&nbsp;</div></td>
          <%

NEXT
%>
        </tr>
        <%
nu_chamada_check=nu_chamada_check+1	 
wend	
%>
        <tr> 
          <td width="17" class="tabela"> 
            <div align="center">
              <%response.Write(nu_chamada)%>
</div></td>
          <td width="283" class="tabela"> 
            <div align="left">
              <%response.Write(NO_Aluno)%>
</div></td>
          <td width="30" class="tabela"> 
            <div align="center">
              <%

			Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo where NU_Periodo="&periodo
		RS4.Open SQL4, CON0


NO_Periodo= RS4("SG_Periodo")
		
							  
					  response.Write(NO_Periodo)%>
</div></td>
          <%
				  
		For k=0 To ubound(nome_nota)

  		Set RS3 = Server.CreateObject("ADODB.Recordset")
		CONEXAO3 = "Select "& campo_check & " from "& nota & " WHERE NU_Periodo = "& periodo &" And CO_Materia = '"& nome_nota(k) &"' And CO_Matricula = "& nu_matricula
		Set RS3 = CON3.Execute(CONEXAO3)
  		
if RS3.EOF Then
%>
          <td width="31" bordercolor="#000000" class="tabelaTit"> 
            <div align="center">&nbsp;</div></td>
          <%else
nota_materia= RS3(""&campo_check&"")
if isnumeric(nota_materia) then
'nota_materia=nota_materia/10
nota_materia = formatNumber(nota_materia,1) 
end if
  %>
          <td width="31" class="tabela"> 
  <%if  nota_materia="" or isnull(nota_materia) then%>
            <div align="center">&nbsp;</div>
<%else%>
            <div align="center">
              <%response.Write(nota_materia)%>
</div>
<%
  end IF
  %>
  </td>
  <%
  
end IF

NEXT%>
        </tr>
        <%
 nu_chamada_check=nu_chamada_check+1	  
end if

	check = check+1
  RSA.MoveNext
  Wend 
%>
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
pasta=arPath(seleciona1)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>