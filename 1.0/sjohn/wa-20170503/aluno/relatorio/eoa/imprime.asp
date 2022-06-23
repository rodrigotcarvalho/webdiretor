<%'On Error Resume Next%>
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
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->

<%opt = REQUEST.QueryString("obr")
p = REQUEST.QueryString("p")
obr=split(opt,"_")

tp_ocor=obr(0)
qtd_ocor=obr(1)
unidade=obr(2)
curso=obr(3)
co_etapa=obr(4)
turma=obr(5)
dia_de=obr(6)
mes_de=obr(7)
dia_ate=obr(8)
mes_ate=obr(9)
ordenacao=obr(10)

if ordenacao = "nome" then
	nome_ordenacao="Nome"
elseif ordenacao = "ucet" then
	nome_ordenacao="Unidade/Curso/Etapa/Turma"
else
	nome_ordenacao="Matrícula"					  					  
end if


ano_letivo = session("ano_letivo")
obr=tp_ocor&"_"&qtd_ocor&"_"&unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&dia_de&"_"&mes_de&"_"&dia_ate&"_"&mes_ate

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

opt=tp_ocor&"_"&qtd_ocor&"_"&unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&dia_de&"_"&mes_de&"_"&dia_ate&"_"&mes_ate

data_de=mes_de&"/"&dia_de&"/"&ano_letivo


dia_de=dia_de*1
mes_de=mes_de*1

if dia_de<10 then
dia_de="0"&dia_de
end if
if mes_de<10 then
mes_de="0"&mes_de
end if


data_inicio=dia_de&"/"&mes_de&"/"&ano_letivo

data_ate=mes_ate&"/"&dia_ate&"/"&ano_letivo

dia_ate=dia_ate*1
mes_ate=mes_ate*1

if dia_ate<10 then
dia_ate="0"&dia_ate
end if
if mes_ate<10 then
mes_ate="0"&mes_ate
end if

data_fim=dia_ate&"/"&mes_ate&"/"&ano_letivo



		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON1 = Server.CreateObject("ADODB.Connection")
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2
		
		Set CON3 = Server.CreateObject("ADODB.Connection") 
		ABRIR3 = "DBQ="& CAMINHO_o & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3		
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CONp = Server.CreateObject("ADODB.Connection") 
		ABRIRp = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONp.Open ABRIRp		
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0	

if unidade="999990" or unidade="" or isnull(unidade) then
	SQL_ALUNOS="NULO"
	no_unidade ="TODAS UNIDADES"
	separador1=0
	separador2=0
	separador3=0
	endereco_completo=""
	no_curso = "Todos"	
	no_etapa= "Todas"	
	nome_turma= "Todas"	
else

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
	no_unidade = "UNIDADE "&RS0("NO_Unidade")
	un_endereco = RS0("NO_Logradouro")
	un_complemento = RS0("TX_Complemento_Logradouro")
	un_numero = RS0("NU_Logradouro")
	un_bairro = RS0("CO_Bairro")
	un_cidade = RS0("CO_Municipio")
	un_uf = RS0("SG_UF")
	un_tel = "Tel: "&RS0("NUS_Telefones")
	un_email = " - E-mail: "&RS0("TX_EMail")
	un_cep = RS0("CO_CEP")
	un_ato = " - "&RS0("TX_Ato_Autorizativo")
	un_cnpj = "CNPJ: "&RS0("CO_CGC")

	cid_estado=cidade&" - "&un_uf
	
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

			if separador2=1then
				endereco_completo=un_endereco&", "&un_numero&" - "&un_complemento&" - "&bairro&" - "&un_cep
			else
				endereco_completo=un_endereco&", "&un_numero&" - "&" - "&bairro&" - "&un_cep
			end if


		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Municipios WHERE SG_UF ='"& un_uf &"' AND CO_Municipio = "&un_cidade
		RS1.Open SQL1, CON0

	cidade= RS1("NO_Municipio")

		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Bairros WHERE CO_Bairro ="& un_bairro &"AND SG_UF ='"& un_uf&"' AND CO_Municipio = "&un_cidade
		RS4.Open SQL4, CON0
	if RS4.EOF then
	bairro = ""
	else
	bairro= RS4("NO_Bairro")
	end if



	
	SQL_ALUNOS= "Select * from TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND NU_Unidade = "& unidade
		
	if curso="999990" or curso="" or isnull(curso) then
		SQL_CURSO=""
		no_curso = "Todos"		
	else
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS2.Open SQL2, CON0
		
		no_curso = RS2("NO_Curso")	
			
		SQL_CURSO=" AND CO_Curso = '"& curso &"'"
	end if

	if co_etapa="999990" or co_etapa="" or isnull(co_etapa) then
		SQL_ETAPA=""
		no_etapa="Todas"
	else

		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' and CO_Curso ='"& curso &"'"  
		RS3.Open SQL3, CON0
		
		if RS3.EOF THEN
		no_etapa="sem etapa cadastrada"
		else
		no_etapa=RS3("NO_Etapa")
		end if
	
		SQL_ETAPA=" AND CO_Etapa = '"& co_etapa &"'"
	end if

	if turma="999990" or turma="" or isnull(turma) then
		SQL_TURMA=""
		nome_turma= "Todas"
	else
		SQL_TURMA=" AND CO_Turma = '"& turma &"' "
		nome_turma=turma
	end if

SQL_ALUNOS= SQL_ALUNOS&SQL_CURSO&SQL_ETAPA&SQL_TURMA&" order by NU_Chamada"
end if

if tp_ocor=999999 or tp_ocor="999999" then
	SQL_TP_OCORRENCIAS=""
ELSE
	SQL_TP_OCORRENCIAS="CO_Ocorrencia ="& tp_ocor&" AND"
end if
if qtd_ocor=0 or qtd_ocor="0" then
	SQL_QTD_OCORRENCIAS=""
else
	qtd_ocor=qtd_ocor*1
	minimo_ocorrencia=qtd_ocor-1
	SQL_QTD_OCORRENCIAS="HAVING COUNT(*)>"&minimo_ocorrencia
end if

%>

<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('window.print()')"> 
<div id="Layer1" style="position:absolute; left:21px; top:21px; width:210px; height:228px; z-index:1"> 
  <table width="950" border="0" align="center" cellspacing="0" class="tb_corpo"
>
  <tr> 
      <td width="203" height="15" bgcolor="#FFFFFF"><div align="center"><img src="../../../../img/logo_preto.gif"> 
        </div></td>
    <td width="741" bgcolor="#FFFFFF"><table width="100%" border="0" align="right" cellspacing="0">
        <tr> 
          <td width="40%" rowspan="2"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>
            <%
			no_unidade= ucase(no_unidade)
			response.Write(" "&no_unidade)
			%>
            </strong></font></td>
          <td width="60%" height="8" class="linhaBaixo">&nbsp; </td>
        </tr>
        <tr> 
          <td height="5" ></td>
        </tr>
        <tr> 
          <td colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">	
            <%response.Write(un_cnpj)
			if separador1=1then
			response.Write(un_ato)
			end if
			%>
            </font></td>
        </tr>
        <tr> 
          <td colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.Write(endereco_completo)
			%>
            </font></td>
        </tr>
        <tr> 
          <td colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.Write(cid_estado)%>
            </font></td>
        </tr>
        <tr> 
          <td colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.Write(un_tel)
			if separador1=1then
			response.Write(un_email)
			end if
			%>
            </font></td>
        </tr>
        <tr> 
          <td colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
             </font></td>
        </tr>
        <tr> 
          <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>OCORR&Ecirc;NCIAS DO ALUNO</strong></font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="12" colspan="2" bgcolor="#EEEEEE" class="linhaTopoL"><table width="950" border="0" cellspacing="0">
        <tr class="tabela"> 
          <td width="150" height="12" align="right" bgcolor="#EEEEEE" > 
            <div align="right"> <strong>Tipo de Ocorrência</strong></div></td>
          <td width="80" height="12" bgcolor="#EEEEEE" >
                      <% 

if tp_ocor=999999 or tp_ocor="999999" then
	no_ocorrencia="Todas"
ELSE
 		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Tipo_Ocorrencia WHERE CO_Ocorrencia="&tp_ocor
		RS5.Open SQL5, CON0

	no_ocorrencia=RS5("NO_Ocorrencia")
end if
Response.Write(no_ocorrencia)%>                  
            </td>
          <td width="250">            <div align="right">
<strong>Per&iacute;odo da 
            Ocorr&ecirc;ncia:</strong></div></td>
          <td width="140">De <%Response.Write(data_inicio)%>  at&eacute; <%Response.Write(data_ate)%> </td>
          <td width="250"><strong><div align="right">Quantidade m&iacute;nima de ocorr&ecirc;ncia:</div></strong></td>
          <td width="80"><% response.Write(qtd_ocor)%></td>
        </tr>
        <tr class="tabela"> 
          <td width="150" height="12" bgcolor="#EEEEEE"> 
            <div align="right"><strong>Curso:</strong></div></td>
          <td width="80" height="12" bgcolor="#EEEEEE">
            <%
response.Write(no_curso)%>
</td>
          <td width="250" height="12" bgcolor="#EEEEEE"> 
            <div align="right">
<strong>Etapa:</strong></div></td>
          <td width="140">
            <%response.Write(no_etapa)%>
</td>
          <td width="250"><strong><div align="right">Turma:</div></strong></td>
          <td width="80"><%response.Write(nome_turma)%></td>
        </tr>
        <tr class="tabela">
        	<td height="12" bgcolor="#EEEEEE"><div align="right"><strong>Ordena&ccedil;&atilde;o:</strong></div></td>
        	<td height="12" bgcolor="#EEEEEE"><%response.Write(nome_ordenacao)%></td>
        	<td height="12" bgcolor="#EEEEEE">&nbsp;</td>
        	<td>&nbsp;</td>
        	<td>&nbsp;</td>
        	<td>&nbsp;</td>
        	</tr>
      </table></td>
  </tr>
  <tr> 
    <td height="10" colspan="2" bgcolor="#EEEEEE"> </td>
  </tr>
  <tr> 
    <td colspan="2">
 <%	

if SQL_ALUNOS="NULO" then
	SQL_MATRICULAS="" 
else

nu_chamada_check = 1
	Set RSA = Server.CreateObject("ADODB.Recordset")
	CONEXAOA = SQL_ALUNOS
	Set RSA = CON1.Execute(CONEXAOA)
	vetor_matriculas="" 
	While Not RSA.EOF
		nu_matricula = RSA("CO_Matricula")
		nu_chamada = RSA("NU_Chamada")
		if nu_chamada_check = 1 and nu_chamada=nu_chamada_check then
			vetor_matriculas=nu_matricula
		elseif nu_chamada_check = 1 then
			while nu_chamada_check < nu_chamada
				nu_chamada_check=nu_chamada_check+1
			wend 
			vetor_matriculas=nu_matricula
		else
			vetor_matriculas=vetor_matriculas&","&nu_matricula
		end if
		nu_chamada_check=nu_chamada_check+1		
	RSA.MoveNext
	Wend 
	SQL_MATRICULAS="CO_Matricula IN("& vetor_matriculas&") AND" 		
end if		
%>	                  
                  
    <table width="950" border="0" align="left" cellspacing="0" bordercolor="#000000">

<%
		Set RSo = Server.CreateObject("ADODB.Recordset")
		SQLo = "SELECT  CO_Matricula, CO_Assunto, CO_Ocorrencia, COUNT(*) AS Num_Ocorrencias FROM TB_Ocorrencia_Aluno WHERE "&SQL_MATRICULAS&" "&SQL_TP_OCORRENCIAS&" (DA_Ocorrencia BETWEEN #"&data_de&"# AND #"&data_ate&"#) GROUP BY CO_Matricula, CO_Assunto, CO_Ocorrencia "&SQL_QTD_OCORRENCIAS
		RSo.Open SQLo, CON3


IF RSo.eof then
%> 
  <tr>
    <th colspan="9"> <font class="form_dado_texto">Não foram encontradas ocorrências que atendessem aos critérios informados.</font></th>
  </tr>
<% 
else
	Set Rs_ordena = Server.CreateObject ( "ADODB.RecordSet" )
	'Vamos adicionar 2 campos nesse recordset!
	'O método Append recebe 3 parâmetros:
	'Nome do campo, Tipo, Tamanho (opcional)
	'O tipo pertence à um DataTypeEnum, e você pode conferir os tipos em
	'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/ado270/htm/mdcstdatatypeenum.asp
	'200 -> VarChar (String), 7 -> Data, 139 -> Numeric
	Rs_ordena.Fields.Append "nu_matric", 139, 10
	Rs_ordena.Fields.Append "nu_chamada", 139, 10	
	Rs_ordena.Fields.Append "nome", 200, 255
	Rs_ordena.Fields.Append "tipo_assunto", 200, 255
	Rs_ordena.Fields.Append "tipo_ocorrencia", 200, 255
	Rs_ordena.Fields.Append "nome_ocorrencia", 200, 255	
	Rs_ordena.Fields.Append "num_ocorrencia", 139, 10
	Rs_ordena.Fields.Append "unidade", 200, 255
	Rs_ordena.Fields.Append "curso", 200, 255
	Rs_ordena.Fields.Append "etapa", 200, 255
	Rs_ordena.Fields.Append "turma", 200, 255
	Rs_ordena.Fields.Append "status", 200, 255
	
	'Vamos abrir o Recordset!
	Rs_ordena.Open
	
	WHILE NOT RSo.eof
		nu_matric=RSo("CO_Matricula")	
		tipo_assunto=RSo("CO_Assunto")
		tipo_ocorrencia=RSo("CO_Ocorrencia")
		num_ocorrencia=RSo("Num_Ocorrencias")
	
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT  * FROM TB_Tipo_Ocorrencia WHERE CO_Assunto= '"&tipo_assunto&"' AND CO_Ocorrencia="&tipo_ocorrencia&""
		RS2.Open SQL2, CON0
	
		nome_ocorrencia=RS2("NO_Ocorrencia")
	
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3= "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& nu_matric
		RS3.Open SQL3, CON1
		
		IF RS3.EOF then
			no_unidade = ""
			no_curso = ""
			no_etapa = ""
		else	
			unidade= RS3("NU_Unidade")
			curso= RS3("CO_Curso")
			etapa= RS3("CO_Etapa")
			turma= RS3("CO_Turma")
			nu_chamada= RS3("NU_Chamada")
		
			call GeraNomes("PORT",unidade,curso,etapa,CON0)
			no_unidade = session("no_unidades")
			no_etapa = session("no_serie")
	
			Set RS5 = Server.CreateObject("ADODB.Recordset")
			Sql5= "SELECT * FROM TB_Curso where CO_Curso = '"& curso &"'"
			Set RS5= CON0.Execute(Sql5) 
			
			IF RS5.eof THEN
				no_curso=""
			ELSE
				no_curso= RS5("NO_Abreviado_Curso")
			END IF
		end if
			
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4= "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& nu_matric
		RS4.Open SQL4, CON1
		
		if RS4.EOF then
			nome_aluno = "Sem nome cadastrado"
		else
			nome_aluno = RS4("NO_Aluno")
		end if	
		
		Rs_ordena.AddNew
		Rs_ordena.Fields("nu_matric").Value = nu_matric		
		Rs_ordena.Fields("nu_chamada").Value = nu_chamada
		Rs_ordena.Fields("nome").Value = nome_aluno
		Rs_ordena.Fields("tipo_assunto").Value = tipo_assunto
		Rs_ordena.Fields("tipo_ocorrencia").Value = tipo_ocorrencia
		Rs_ordena.Fields("nome_ocorrencia").Value = nome_ocorrencia
		Rs_ordena.Fields("num_ocorrencia").Value = num_ocorrencia
		Rs_ordena.Fields("unidade").Value = no_unidade
		Rs_ordena.Fields("curso").Value = no_curso
		Rs_ordena.Fields("etapa").Value = no_etapa
		Rs_ordena.Fields("turma").Value = turma
		
	RSo.MOVENEXT
	wend	
		
	if ordenacao = "nome" then
		Rs_ordena.Sort = "nome ASC"		
	elseif ordenacao = "ucet" then
		Rs_ordena.Sort = "unidade ASC, curso ASC, etapa ASC, turma ASC, nu_chamada ASC"		
	else
		Rs_ordena.Sort = "nu_matric ASC"					  					  
	end if	


	sem_link=0
	Rs_ordena.PageSize = 30

	if Request.QueryString("pagina")="" then
		  intpagina = 1
		  Rs_ordena.MoveFirst
	else
		if cint(Request.QueryString("pagina"))<1 then
			intpagina = 1
		else
			if cint(Request.QueryString("pagina"))>Rs_ordena.PageCount then  
				intpagina = Rs_ordena.PageCount
			else
				intpagina = Request.QueryString("pagina")
			end if
		end if   
	end if   
	
	Rs_ordena.AbsolutePage = intpagina
	intrec = 0
	check=2

%> 
  <tr>
    <th width="110" scope="col" class="tabelaTit"><div align="center">Unidade</div></th>
    <th width="70" scope="col" class="tabelaTit"><div align="center">Curso</div></th>
    <th width="70" scope="col" class="tabelaTit"><div align="center">Etapa</div></th>
    <th width="70" scope="col" class="tabelaTit"><div align="center">Turma</div></th>
    <th width="50" scope="col" class="tabelaTit">Chamada</th>
    <th width="50" scope="col" class="tabelaTit">Matr&iacute;cula</th>
    <th width="300" scope="col" class="tabelaTit">Nome</th>
    <th width="250" scope="col" class="tabelaTit">Ocorr&ecirc;ncia</th>
    <th width="20" scope="col" class="tabelaTit">Qtd</th>
  </tr>
<% 

	While intrec<Rs_ordena.PageSize and Not Rs_ordena.EoF	
	
	if check mod 2 =0 then
		cor = "tb_fundo_linha_par" 
	else
		cor ="tb_fundo_linha_impar"
	end if 
	no_unidade = Rs_ordena.Fields("unidade").Value
	no_curso = Rs_ordena.Fields("curso").Value
	no_etapa = Rs_ordena.Fields("etapa").Value
	turma = Rs_ordena.Fields("turma").Value
	nu_chamada = Rs_ordena.Fields("nu_chamada").Value
	nu_matric = Rs_ordena.Fields("nu_matric").Value
	nome_aluno = Rs_ordena.Fields("nome").Value
	nome_ocorrencia = Rs_ordena.Fields("nome_ocorrencia").Value
	num_ocorrencia = Rs_ordena.Fields("num_ocorrencia").Value
	
	%>  
	  <tr class="tabela">
		<td width="110" class="tabela"><div align="center"><%response.Write(no_unidade)%></div></td>
		<td width="70" class="tabela"><div align="center"><%response.Write(no_curso)%></div></td>
		<td width="70" class="tabela"><div align="center"><%response.Write(no_etapa)%></div></td>
		<td width="70" class="tabela"><div align="center"><%response.Write(turma)%></div></td>
		<td width="50" class="tabela"><div align="center"><%response.Write(nu_chamada)%></div></td>
		<td width="50" class="tabela"><div align="center"><%response.Write(nu_matric)%></div></td>
		<td width="300" class="tabela"><div align="center"><%response.Write(nome_aluno)%></div></td>
		<td width="250" class="tabela"><div align="center"><%response.Write(nome_ocorrencia)%></div></td>
		<td width="20" class="tabela"><div align="center"><%response.Write(num_ocorrencia)%></div></td>
	  </tr>
	<%
		intrec=intrec+1
		check=check+1				
	Rs_ordena.movenext
	Wend
END iF	
     RSo.close
    Set RSo = Nothing
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
</table></div>
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