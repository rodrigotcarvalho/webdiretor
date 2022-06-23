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




<!--#include file="../../../../inc/funcoes4.asp"-->

<%opt = REQUEST.QueryString("obr")
p = REQUEST.QueryString("p")
obr=split(opt,"_")

unidade = obr(0)
curso = obr(1)
co_etapa = obr(2)
turma = obr(3)
periodo = obr(4)

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

minimo_recuperacao=50

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2


		Set RSTB = Server.CreateObject("ADODB.Recordset")
		CONEXAOTB = "Select * from TB_Da_Aula WHERE NU_Unidade="& unidade &" AND CO_Curso='"& curso &"'AND CO_Etapa='"& co_etapa &"'AND CO_Turma='"& turma &"'"
		Set RSTB = CON2.Execute(CONEXAOTB)
		
nota= RSTB("TP_Nota")

		
if nota = "TB_NOTA_A" Then		
		CAMINHOn = CAMINHO_na
elseif nota = "TB_NOTA_B" Then
		CAMINHOn = CAMINHO_nb
elseif nota = "TB_NOTA_C" Then
		CAMINHOn = CAMINHO_nc
end if

		Set CON3 = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3

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
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"'" 
		RS3.Open SQL3, CON0
		
if RS3.EOF THEN
no_etapa="sem etapa"
else
no_etapa=RS3("NO_Etapa")
end if




if p="999" then
	%>
<body link="#6699CC" vlink="#6699CC" alink="#6699CC" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('window.print()');Layer1.style.filter='progid:DXImageTransform.Microsoft.BasicImage(rotation=1)'"> 
<%else
%>
<body link="#6699CC" vlink="#6699CC" alink="#6699CC" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('window.print()')"> 
<%end if%>
<br>
<div id="Layer1" style="position:absolute; left:21px; top:21px; width:210px; height:228px; z-index:1"> 
  <table width="950" border="0" align="center" cellspacing="0" bgcolor="#F8FAFC">
  <tr> 
    <td width="203" height="15" bgcolor="#FFFFFF"><img src="../../../../img/logo_preto.gif" width="175" height="130"> 
    </td>
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
            <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>MAP&Atilde;O 
              DE M&Eacute;DIAS FINAIS</strong></font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="12" colspan="2" bgcolor="#EEEEEE" class="linhaTopoL"><table width="900" border="0" cellspacing="0">
        <tr class="tabela"> 
          <td width="79" height="12" align="right" bgcolor="#EEEEEE" > 
            <div align="right"> <strong>Ano 
              Letivo:</strong></div></td>
          <td width="197" height="12" bgcolor="#EEEEEE" >
            <%response.Write(ano_letivo)%>
            </td>
          <td width="41">&nbsp;</td>
          <td width="195">&nbsp;</td>
          <td width="378">&nbsp;</td>
        </tr>
        <tr class="tabela"> 
          <td width="79" height="12" bgcolor="#EEEEEE"> 
            <div align="right"><strong>Curso:</strong></div></td>
          <td width="197" height="12" bgcolor="#EEEEEE">
            <%
response.Write(no_curso)%>
</td>
          <td width="41" height="12" bgcolor="#EEEEEE"> 
            <div align="right">
<strong>Etapa:</strong></div></td>
          <td width="195">
            <%response.Write(no_etapa)%>
</td>
          <td width="378">&nbsp;</td>
        </tr>
        <tr class="tabela"> 
          <td width="79" height="12" bgcolor="#EEEEEE"> 
            <div align="right"><strong><strong>Turma:</strong></strong></div></td>
          <td width="197" height="12" bgcolor="#EEEEEE">
            <%response.Write(turma)%>
            &nbsp;</td>
          <td width="41" height="12" bgcolor="#EEEEEE"><div align="right"></div></td>
            <td width="195">&nbsp;</td>
          <td width="378">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="10" colspan="2" bgcolor="#EEEEEE"> </td>
  </tr>
  <tr> 
    <td colspan="2"> <%
Set RSNN = Server.CreateObject("ADODB.Recordset")
        CONEXAONN = "Select * from TB_Programa_Aula WHERE CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa&"' order by NU_Ordem_Boletim"
        Set RSNN = CON0.Execute(CONEXAONN)
 
materia_nome_check="vazio"
nome_nota="vazio"
i=0
largura = 0
While not RSNN.eof
materia_nome= RSNN("CO_Materia")
    mae=RSNN("IN_MAE")
    fil=RSNN("IN_FIL")
    in_co=RSNN("IN_CO")
    nu_peso=RSNN("NU_Peso")
	ordem=RSNN("NU_Ordem_Boletim")
'response.Write(materia_nome&"-"&ordem)
if mae=TRUE AND fil=true AND in_co=false AND isnull(nu_peso) then

If Not IsArray(tipo_mae) Then
tipo_mae = Array()
End if
tipo = 4
ReDim preserve tipo_mae(UBound(tipo_mae)+1)
tipo_mae(Ubound(tipo_mae)) = tipo

If Not IsArray(ordem_mae) Then
ordem_mae = Array()
End if
ReDim preserve ordem_mae(UBound(ordem_mae)+1)
ordem_mae(Ubound(ordem_mae)) = ordem

If Not IsArray(nome_nota) Then
nome_nota = Array()
End if
If InStr(Join(nome_nota), materia_nome) = 0 Then
ReDim preserve nome_nota(UBound(nome_nota)+1)
nome_nota(Ubound(nome_nota)) = materia_nome
largura=largura+35

i=i+1

RSNN.movenext
else
RSNN.movenext
end if
 

' sub do anterior
elseif mae=false AND fil=true AND in_co=false then

RSNN.movenext


'MCAL


elseif mae=TRUE AND fil=false AND in_co=true AND isnull(nu_peso) then

If Not IsArray(tipo_mae) Then
tipo_mae = Array()
End if
tipo = 2
ReDim preserve tipo_mae(UBound(tipo_mae)+1)
tipo_mae(Ubound(tipo_mae)) = tipo

If Not IsArray(ordem_mae) Then
ordem_mae = Array()
End if
ReDim preserve ordem_mae(UBound(ordem_mae)+1)
ordem_mae(Ubound(ordem_mae)) = ordem

If Not IsArray(nome_nota) Then
nome_nota = Array()
End if
If InStr(Join(nome_nota), materia_nome) = 0 Then
ReDim preserve nome_nota(UBound(nome_nota)+1)
nome_nota(Ubound(nome_nota)) = materia_nome
largura=largura+35

i=i+1
RSNN.movenext
else
RSNN.movenext
end if


'sub do anterior - MATE 1 E MATE2
elseif mae=false AND fil =false AND in_co=True AND isnull(nu_peso) then

RSNN.movenext

elseif mae=TRUE AND fil=false AND in_co=false AND isnull(nu_peso) then

If Not IsArray(tipo_mae) Then
tipo_mae = Array()
End if
tipo = 3
ReDim preserve tipo_mae(UBound(tipo_mae)+1)
tipo_mae(Ubound(tipo_mae)) = tipo

If Not IsArray(ordem_mae) Then
ordem_mae = Array()
End if
ReDim preserve ordem_mae(UBound(ordem_mae)+1)
ordem_mae(Ubound(ordem_mae)) = ordem

If Not IsArray(nome_nota) Then
nome_nota = Array()
End if
If InStr(Join(nome_nota), materia_nome) = 0 Then
ReDim preserve nome_nota(UBound(nome_nota)+1)
nome_nota(Ubound(nome_nota)) = materia_nome
largura=largura+35

i=i+1


 RSNN.movenext
 else
RSNN.movenext
end if

'se não for nenhum
elseif mae=TRUE AND fil=TRUE AND in_co=false then

If Not IsArray(tipo_mae) Then
tipo_mae = Array()
End if
tipo = 1
ReDim preserve tipo_mae(UBound(tipo_mae)+1)
tipo_mae(Ubound(tipo_mae)) = tipo

If Not IsArray(ordem_mae) Then
ordem_mae = Array()
End if
ReDim preserve ordem_mae(UBound(ordem_mae)+1)
ordem_mae(Ubound(ordem_mae)) = ordem

If Not IsArray(nome_nota) Then
nome_nota = Array()
End if
If InStr(Join(nome_nota), materia_nome) = 0 Then
ReDim preserve nome_nota(UBound(nome_nota)+1)
nome_nota(Ubound(nome_nota)) = materia_nome
largura=largura+35

i=i+1

RSNN.movenext
else
RSNN.movenext
end if
end if
wend
larg=950-17-30-largura

' response.Write(larg&"_"&i)%> 
        <table width="950" border="0" align="left" cellspacing="0" bordercolor="#000000">
          <tr> 
            <td width="17" height="40" class="tabelaTit"> <div align="center">N&ordm;</div></td>
            <td width="<%=larg%>" class="tabelaTit"> <div align="center"><strong>Nome</strong></div></td>
            <%For k=0 To ubound(nome_nota)%>
            <td width="31" class="tabelaTit"> 
              <% response.Write(nome_nota(k))%>
            </td>
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
medias = Array()

  		Set RSA2 = Server.CreateObject("ADODB.Recordset")
		CONEXAOA2 = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
		Set RSA2 = CON4.Execute(CONEXAOA2)
  		NO_Aluno= RSA2("NO_Aluno")

 if check mod 2 =0 then
  cor = "#F8FAFC" 
  else cor ="#F1F5FA"
  end if
  
if nu_chamada=nu_chamada_check then
nu_chamada_check=nu_chamada_check+1%>
          <tr> 
            <td rowspan="2" width="17" class="tabela"> <div align="center"> 
                <%response.Write(nu_chamada)%>
              </div></td>
            <td rowspan="2" width="283"  class="tabela"> <div align="left"> 
                <%response.Write(NO_Aluno)%>
              </div></td>
            <%For k=0 To ubound(nome_nota)
%>
            <td class="tabela"> <div align="center">
                <%tipo=tipo_mae(k)
			  ordem2=ordem_mae(k)
			  materia=nome_nota(k)
			 'response.Write(">>"&ordem2&"-")
			  call Calc_Med_An_Fin(nu_matricula,unidade,curso,co_etapa,turma,materia,ordem2,tipo)
somamp=session("somampAn")
mamp=session("medFin")

If Not IsArray(medias) Then
medias = Array()
End if
ReDim preserve medias(UBound(medias)+1)
medias(Ubound(medias)) = mamp

response.Write(mamp)%>
              </div></td>
<%

NEXT%>
       </tr>		<tr class="tabela">
<%For k=0 To ubound(nome_nota)
%>
<td class="tabela"> <div align="center"><strong>
              <%
mamp=medias(k)
mamp=mamp*1
minimo_recuperacao=minimo_recuperacao*1
If mamp >= minimo_recuperacao then
res="APR"
else
res="REP"
END IF		  
response.Write(res)%>
</strong></div>
		</td>
<%NEXT%>
		</tr>
          <% 
else
While nu_chamada>nu_chamada_check
%>
          <tr> 
            <td width="17" bgcolor="#E4E4E4"  class="tabela"> <div align="center"> 
                <%response.Write(nu_chamada_check)%>
              </div></td>
            <td width="283" bordercolor="#000000" bgcolor="#E4E4E4"  class="tabela"> 
              <div align="left">&nbsp;</div></td>
            <%For k=0 To ubound(nome_nota)%>
            <td width="31" bordercolor="#000000" bgcolor="#E4E4E4"  class="tabelaTit"> 
              <div align="center">&nbsp;</div></td>
            <%

NEXT
%>
          </tr>
          <%
nu_chamada_check=nu_chamada_check+1	 
wend	
%>
          <tr class="tabela">  
            <td width="17" rowspan="2"  class="tabela"> <div align="center"> 
                <%response.Write(nu_chamada)%>
              </div></td>
            <td width="283" rowspan="2"  class="tabela"> <div align="left"> 
                <%response.Write(NO_Aluno)%>
              </div></td>
            <%
				  
		For k=0 To ubound(nome_nota)

%>
            <td class="tabela"> <div align="center">
                <%tipo=tipo_mae(k)
			  ordem2=ordem_mae(k)
			  materia=nome_nota(k)
			 'response.Write(">>"&ordem2&"-")
			  call Calc_Med_An_Fin(nu_matricula,unidade,curso,co_etapa,turma,materia,ordem2,tipo)

mamp=session("medFin")

If Not IsArray(medias) Then
medias = Array()
End if
ReDim preserve medias(UBound(medias)+1)
medias(Ubound(medias)) = mamp

response.Write(mamp)%>
              </div></td>
            <%

NEXT%>
       </tr>		<tr class="tabela">
<%For k=0 To ubound(nome_nota)
%>
<td class="tabela"> <div align="center"><strong>
              <%
mamp=medias(k)
mamp=mamp*1
minimo_recuperacao=minimo_recuperacao*1
If mamp >= minimo_recuperacao then
res="APR"
else
res="REP"
END IF		  
response.Write(res)%>
              </strong></div>
		</td>
<%NEXT%>
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