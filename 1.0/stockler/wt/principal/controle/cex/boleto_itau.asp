<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 60 'valor em segundos

%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->
<html>
<head>
<STYLE type=text/css>
.ti { FONT: 9px Arial, Helvetica, sans-serif }
.ct { FONT: 9px Arial Narrow; COLOR: black /* navy*/ }
.cn { FONT: 9px Arial; COLOR: black }
.cp { FONT: bold 11px Arial; COLOR: black }
.ld { FONT: bold 15px Arial; COLOR: #000000 }
.bc { FONT: bold 18px Arial; COLOR: #000000 }
.wp	{ FONT: bold 11px Arial; COLOR: black }
</STYLE>

</head>
<%

' Response.Expires = 60
'Response.Expiresabsolute = Now() - 1
'Response.AddHeader "pragma","no-cache"
'Response.AddHeader "cache-control","private"
'Response.CacheControl = "no-cache"
'dados = request.QueryString("vc")
dados = request.QueryString("opt")
'dados = replace(dados,"$!$", ", ")
cod_cons = request.QueryString("c")
vetor_meses = split(dados,", ")
'vetor_meses = 
ano_letivo = session("ano_letivo")

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

if dia<10 then
	dia="0"&dia
end if

if mes<10 then
	mes="0"&mes
end if

data_documento = dia&"/"&mes&"/"&ano

'**************************
FUNCTION linhadigitavel(codigobarras)
'**************************
cmplivre=mid(codigobarras,20,25)
campo1=left(codigobarras,4)&mid(cmplivre,1,5)
campo1=campo1&calcdig10(campo1)
campo1=mid(campo1,1,5)&"."&mid(campo1,6,5)

campo2=mid(cmplivre,6,10)
campo2=campo2&calcdig10(campo2)
campo2=mid(campo2,1,5)&"."&mid(campo2,6,6)

campo3=mid(cmplivre,16,10)
campo3=campo3&calcdig10(campo3)
campo3=mid(campo3,1,5)&"."&mid(campo3,6,6)

campo4=mid(codigobarras,5,1)

campo5=int(mid(codigobarras,6,14))

if campo5=0 then
	campo5="000"
end if

linhadigitavel=campo1&"&nbsp;&nbsp;"&campo2&"&nbsp;&nbsp;"&campo3&"&nbsp;&nbsp;"&campo4&"&nbsp;&nbsp;"&campo5
'*************************
END FUNCTION
'*************************




'valortal=CALCdig10("11513024791005193100033")
'response.write valortal

'**************************
FUNCTION CALCDIG10(cadeia)
'**************************
	mult=(len(cadeia) mod 2) 
	mult=mult+1
	total=0
	for pos=1 to len(cadeia)
		res= mid(cadeia, pos, 1) * mult
		if res>9 then
			res=int(res/10) + (res mod 10)
		end if
		total=total+res
		if mult=2 then
			mult=1
		else
			mult=2
		end if
	next
	total=((10-(total mod 10)) mod 10 )
	CALCDIG10=total
'*************************
END FUNCTION
'*************************




'valortal1=CALCdig11("0339000000000103581481302647800076960003348",9,0)
'response.write valortal1

'**************************
FUNCTION CALCDIG11(cadeia,limitesup,lflag)
'**************************
mult=1 + (len(cadeia) mod (limitesup-1))
if mult=1 then
	mult=limitesup
end if
total=0
for pos=1 to len(cadeia)
	total=total+(mid(cadeia,pos,1) * mult)
	mult=mult-1
	if mult=1 then
		mult=limitesup
	end if
Next
nresto=(total mod 11)
if lflag = 1 then
	calcdig11=nresto
else
	if nresto=0 or nresto=1 or nresto=10 then
		ndig=1
	else
		ndig=11 - nresto	
	end if
	calcdig11=ndig
end if

'*************************
END FUNCTION
'*************************



'**************************
'FUNCTION fatorvencimento(vencimento)
''**************************
'
'if len(vencimento)<8 then
'   fatorvencimento="0000"
'else
'   fatorvencimento=datevalue(""&vencimento&"")-datevalue("1997/10/07")
'end if
'
''*************************
'END FUNCTION
'*************************




'**************************
FUNCTION codbar(banco,moeda,vencimento,valor,carteira,nossonumero,dvnossonumero,agencia,conta,dvagconta)
'**************************

strcodbar=banco&moeda&vencimento&valor&carteira&nossonumero&dvnossonumero&agencia&conta&dvagconta&"000"
dv3=calcdig11(strcodbar,9,0)
codbar=banco&moeda&dv3&vencimento&valor&carteira&nossonumero&dvnossonumero&agencia&conta&dvagconta&"000"
'*************************
END FUNCTION
'*************************

'**************************
Sub WBarCode( Valor )
'**************************

Dim f, f1, f2, i
Dim texto
Const fino = 1
Const largo = 3
Const altura = 50
Dim BarCodes(99)

if isempty(BarCodes(0)) then
  BarCodes(0) = "00110"
  BarCodes(1) = "10001"
  BarCodes(2) = "01001"
  BarCodes(3) = "11000"
  BarCodes(4) = "00101"
  BarCodes(5) = "10100"
  BarCodes(6) = "01100"
  BarCodes(7) = "00011"
  BarCodes(8) = "10010"
  BarCodes(9) = "01010"
  for f1 = 9 to 0 step -1
    for f2 = 9 to 0 Step -1
      f = f1 * 10 + f2
      texto = ""
      for i = 1 To 5
        texto = texto & mid(BarCodes(f1), i, 1) + mid(BarCodes(f2), i, 1)
      next
      BarCodes(f) = texto
    next
  next
end if

'Desenho da barra


' Guarda inicial
%>


<img src=../../../../img/boleto_itau/2.gif width=<%=fino%> height=<%=altura%> border=0><img 
src=../../../../img/boleto_itau/1.gif width=<%=fino%> height=<%=altura%> border=0><img 
src=../../../../img/boleto_itau/2.gif width=<%=fino%> height=<%=altura%> border=0><img 
src=../../../../img/boleto_itau/1.gif width=<%=fino%> height=<%=altura%> border=0><img 

<%
texto = valor
if len( texto ) mod 2 <> 0 then
  texto = "0" & texto
end if


' Draw dos dados
do while len(texto) > 0
  i = cint( left( texto, 2) )
  texto = right( texto, len( texto ) - 2)
  f = BarCodes(i)
  for i = 1 to 10 step 2
    if mid(f, i, 1) = "0" then
      f1 = fino
    else
      f1 = largo
    end if
    %>
    src=../../../../img/boleto_itau/2.gif width=<%=f1%> height=<%=altura%> border=0><img 
    <%
    if mid(f, i + 1, 1) = "0" Then
      f2 = fino
    else
      f2 = largo
    end if
    %>
    src=../../../../img/boleto_itau/1.gif width=<%=f2%> height=<%=altura%> border=0><img 
    <%
  next
loop

' Draw guarda final
%>
src=../../../../img/boleto_itau/2.gif width=<%=largo%> height=<%=altura%> border=0><img 
src=../../../../img/boleto_itau/1.gif width=<%=fino%> height=<%=altura%> border=0><img 
src=../../../../img/boleto_itau/2.gif width=<%=1%> height=<%=altura%> border=0>

<%
'**************************
end sub
'**************************

SDIG=""
CDIG=""
LDIG=""
NOSSONUMERO=""
Dim atab(99)

	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR

	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1
			
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0	

	Set CONBL = Server.CreateObject("ADODB.Connection") 
	ABRIRBL = "DBQ="& CAMINHO_bl & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONBL.Open ABRIRBL
	
	Set CON4 = Server.CreateObject("ADODB.Connection") 
	ABRIR4 = "DBQ="& CAMINHO_pf & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON4.Open ABRIR4	
	
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod_cons
	RS.Open SQL, CON1
	
	nome_aluno = RS("NO_Aluno")
	sexo_aluno = RS("IN_Sexo")
	
	nome_aluno=replace_latin_char(nome_aluno,"html")	
	
	if sexo_aluno="F" then
		desinencia="a"
	else
		desinencia="o"
	end if


	Set RS1 = Server.CreateObject("ADODB.Recordset")
	SQL1 = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod_cons
	RS1.Open SQL1, CON1
	
	if RS1.EOF then
		response.redirect("index.asp?nvg="&nvg&"&opt=err1")
	else
	
		ano_aluno = RS1("NU_Ano")
		rematricula = RS1("DA_Rematricula")
		situacao = RS1("CO_Situacao")
		encerramento= RS1("DA_Encerramento")
		unidade= RS1("NU_Unidade")
		'curso= RS1("CO_Curso")
		'etapa= RS1("CO_Etapa")
		'turma= RS1("CO_Turma")
		cham= RS1("NU_Chamada")
			
		call GeraNomes("PORT",unidade,curso,etapa,CON0)
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Unidade WHERE NU_Unidade="& unidade
		RS2.Open SQL2, CON0
		
		if RS2.EOF then
			no_unidade = ""
			co_cnpj = ""
		else				
			no_unidade = RS2("TX_Imp_Cabecalho")	
			co_cnpj = RS2("CO_CGC")			
		end if
'		no_curso= session("no_grau")
'		no_etapa = session("no_serie")
				
		vetor_cnpj=SPLIT(co_cnpj,"/")
		if ubound(vetor_cnpj)>0 then
			if vetor_cnpj(1)<0 then
				vetor_cnpj(1)=vetor_cnpj(1)*10
			end if
			exibe_cnpj="CNPJ: "&vetor_cnpj(0)&"/"&vetor_cnpj(1)
		end if					
						
'		Set RS3 = Server.CreateObject("ADODB.Recordset")
'		SQL3 = "SELECT * FROM TB_Curso WHERE CO_Curso='"& curso &"'"
'		RS3.Open SQL3, CON0
'		
'		no_abrv_curso = RS3("NO_Abreviado_Curso")
'		co_concordancia_curso = RS3("CO_Conc")	
		
		'no_unidade = unidade&" - "&no_unidade
		'no_curso= no_curso&" - "&no_etapa	
	end if	
for n=0 to ubound(vetor_meses)
	Set RS4 = Server.CreateObject("ADODB.Recordset")
	SQL4= "SELECT * FROM TB_Posicao WHERE VA_Realizado=0 AND CO_Matricula_Escola ="& cod_cons &" AND Mes = "&vetor_meses(n)
	RS4.Open SQL4, CON4	

	if RS4.EOF then
		response.redirect("index.asp?nvg="&nvg&"&opt=err2")
	else
		vencimento=RS4("DA_Vencimento")
		nu_cota=RS4("NU_Cota")
		
		vetor_vencimento = split(vencimento, "/")
		
	    dia_vencimento = vetor_vencimento(0)*1
    	mes_vencimento = vetor_vencimento(1)*1 

		if dia_vencimento<10 then
			dia_vencimento="0"&dia_vencimento		
		end if
		
		if mes_vencimento<10 then
			mes_vencimento="0"&mes_vencimento				
		end if

		vencimento = dia_vencimento&"/"&mes_vencimento&"/"&vetor_vencimento(2)
		vencimento_inicial = vetor_vencimento(1)&"/"&vetor_vencimento(0)&"/"&vetor_vencimento(2)
		
		if ((((vetor_vencimento(1) = 1 or vetor_vencimento(1) = 3 or vetor_vencimento(1) = 5 or vetor_vencimento(1) = 7 or vetor_vencimento(1) = 8 or vetor_vencimento(1) = 10 or vetor_vencimento(1) = 12) and vetor_vencimento(0) = 31)   or   (vetor_vencimento(1) = 4 or vetor_vencimento(1) = 6 or vetor_vencimento(1) = 9 or vetor_vencimento(1) = 11) and vetor_vencimento(0) = 30)) then
			dia_vencimento = 1
			mes_vencimento = vetor_vencimento(1)+1
		elseif ((vetor_vencimento(1) = 2 and (vetor_vencimento(2) MOD 4 = 0) and  vetor_vencimento(0) = 29) or (vetor_vencimento(1) = 2 and  vetor_vencimento(0) = 28)) then
			dia_vencimento = 1
			mes_vencimento = vetor_vencimento(1)+1				
		else
			dia_vencimento = vetor_vencimento(0)+1
			mes_vencimento = vetor_vencimento(1)
		end if	
		
		if ((vetor_vencimento(1) = 12) and vetor_vencimento(0) = 31) then
			ano_vencimento = vetor_vencimento(2)+1			
		else
			ano_vencimento = vetor_vencimento(2)				
		end if 
		vencimento_fim = mes_vencimento&"/"&dia_vencimento&"/"&ano_vencimento
	end if
	
	Set RS1 = Server.CreateObject("ADODB.Recordset")
	SQL1 = "SELECT * FROM TB_Bloqueto WHERE DA_Vencimento >=#"& vencimento_inicial &"# and DA_Vencimento <#"& vencimento_fim &"#  AND CO_Matricula_Escola ="& cod_cons
	RS1.Open SQL1, CONBL
	
	if RS1.EOF then	
		response.Write("ERRO! Valores não encontrados em TB_Bloqueto")	
		response.End()
	else	
		nu_carne=RS1("NU_Bloqueto")
		nosso_numero = RS1("CO_Nosso_Numero")
		va_inicial = RS1("VA_Inicial")
		cod_superior=RS1("CO_Superior")				
		cod_barras =RS1("CO_Barras")
		turma =RS1("CO_Turma")
		no_cedente =RS1("NO_Cedente")
		co_agencia =RS1("CO_Agencia")
		co_conta =RS1("CO_Conta")
		da_process =RS1("DA_Processamento")
		msg01 =RS1("TX_Msg_01")
		msg02 =RS1("TX_Msg_02")
		msg03 =RS1("TX_Msg_03")
		end_rua =RS1("NO_Logradouro_Empresa")
		end_num =RS1("NU_Logradouro_Empresa")
		end_comp =RS1("TX_Complemento_Logradouro_Empresa")
		end_bairro =RS1("NO_Bairro_Empresa")
		end_cid =RS1("NO_Cidade_Empresa")
		end_uf =RS1("SG_UF_Empresa")
		end_cep =RS1("CO_CEP_Empresa")
		no_curso=RS1("NO_Grau")
		no_etapa=RS1("NO_Serie")
		cpf_responsavel=RS1("CO_CPF")
		no_responsavel=RS1("NO_Responsavel")		
	end if	

	
 



'********************************
' CONSTANTES
'********************************	

cons_banco="341"
cons_dvbanco="7"
cons_agencia=co_agencia
cons_conta=co_conta
cons_carteira="175"
cons_moeda="9"
cons_especie="R$"
cons_cedente=no_cedente
'cons_dadoscedente=no_cedente&".<br>&Av. Afonso Arinos de Melo Fanco, 397 / 1404<br>Barra da Tijuca<br>22631-455&nbsp;- Rio de Janeiro - RJ<br>Telefone / Fax: +55 (21) 3086-0080<br>E-mail: suporte@rea.com.br"

'********************************
' VARIÁVEIS 
'********************************

var_sacado=no_responsavel
var_endereco=end_rua &", "& end_num&"/ "&end_comp
var_bairro=end_bairro
var_cidade=end_cid
var_estado=end_uf
var_cep=end_cep
var_cpfcnpj="00.000.000/0000-00"

var_nossonumero=nosso_numero
var_datadocumento=data_documento
var_datavencimento=vencimento
var_valordocumento=va_inicial
var_numerodoc=nu_cota
var_instrucoes=msg01&"<BR>"&msg02&"<BR>"&msg03&"<BR> Após o vencimento cobrar MULTA de 2% do valor principal<BR> Atraso superior a 30 dias acrescentar JUROS de 1% ao mês" 
var_observacoes="<B> Linha 1<BR> Linha 2<BR> Linha 3<BR> Linha 4<BR> Linha 5<BR> Linha 6<BR></b>" 



'********************************
' INICIO DO CÁLCULO
'********************************

'dvnossonumero=calcdig10(cons_agencia&cons_conta&cons_carteira&var_nossonumero)
'dvagconta=calcdig10(cons_agencia&cons_conta)


valordia=date()
var_data=Day(valordia) & "/" & Month(valordia) & "/" & YEAR(valordia)

valorvalor1=var_valordocumento
valorvalor2=replace(valorvalor1,",","")
valorvalor2=replace(valorvalor2,".","")
valorvalor3=len(valorvalor2)
valorvalor4=10-valorvalor3
var_valor= String(""&valorvalor4&"","0") & (""&valorvalor2&"")
if valorvalor1=0 then
   var_valor=""
end if

'var_fatorvencimento=fatorvencimento(""&var_datavencimento&"")
'if var_fatorvencimento="0000" then
'   var_datavencimento="Contra Apresentação"
'end if



'var_codigobarras=codbar(""&cons_banco&"",""&cons_moeda&"",""&var_fatorvencimento&"",""&var_valor&"",""&cons_carteira&"",""&var_nossonumero&"",""&dvnossonumero&"",""&cons_agencia&"",""&cons_conta&"",""&dvagconta&"")
'var_linhadigitavel=linhadigitavel(""&var_codigobarras&"")
var_codigobarras=cod_barras
var_linhadigitavel=cod_superior

%>


<SCRIPT language=JavaScript>
var da = (document.all) ? 1 : 0;
var pr = (window.print) ? 1 : 0;
var mac = (navigator.userAgent.indexOf("Mac") != -1); 

function x86(){
if (pr) // NS4, IE5
//window.print()
	printPage()
else if (da && !mac) // IE4 (Windows)
vbx86()
else // outros browsers
alert("Desculpe seu browser não suporta esta função. Por favor utilize a barra de trabalho para imprimir a página.");
return false;}
if (da && !pr && !mac) with (document) {
writeln('<OBJECT ID="WB" WIDTH="0" HEIGHT="0" CLASSID="clsid:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>');
writeln('<' + 'SCRIPT LANGUAGE="VBScript">');
writeln('Sub window_onunload');
writeln('  On Error Resume Next');
writeln('  Set WB = nothing');
writeln('End Sub');
writeln('Sub vbx86');
writeln('  OLECMDID_PRINT = 6');
writeln('  OLECMDEXECOPT_DONTPROMPTUSER = 2');
writeln('  OLECMDEXECOPT_PROMPTUSER = 1');
writeln('  On Error Resume Next');
writeln('  WB.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER');
writeln('End Sub');
writeln('<' + '/SCRIPT>');}

//function printPage(){
//document.getElementById('print_button').style.display='none';
//window.print();
//document.getElementById('print_button').style.display='block';
//}
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
</SCRIPT>

<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('window.print()')"> 

<CENTER>
<TABLE WIDTH="660" CELLSPACING=0 CELLPADDING=0 BORDER=0>
  	<TR>
		<TD class=cp VALIGN=BOTTOM WIDTH=225>&nbsp;</TD>
		<TD ALIGN=RIGHT VALIGN=BOTTOM><FONT class=ld><B>RECIBO DO ALUNO</B></FONT></TD>
	</TR>
</TABLE>
<TABLE WIDTH="660" CELLSPACING=0 CELLPADDING=0 BORDER=0>
	<TR>
		<TD width="139" align=center><p><img src="../../../../img/logo_boleto.gif" width="62" height="70" /><br />
			<font class="cp">
			<%response.Write(no_unidade)%>
			</font><br />
		<font class="cn">
				<%response.Write(exibe_cnpj)%>
			</font></p></TD>
		<TD width="521" align=right><table width="100%" border="1" cellspacing="0" cellpadding="1">
			<tr>
				<td colspan="4"><font class="ct">Aluno</font><br />
					<font class="cp">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=nome_aluno%></font></td>
			</tr>
			<tr>
				<td width="20%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
							<td><font class="ct">Matr&iacute;cula</font></td>
						</tr>
						<tr>
							<td align="center"><font align="center" class="cn">&nbsp;</font><font class="cp"><%=cod_cons%></font></td>
						</tr>
					</table></td>
				<td width="20%"><table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><font class="ct">Turma</font></td>
						</tr>
					<tr>
							<td align="center"><font align="center" class="cn">&nbsp;<%=turma%></font></td>
						</tr>
				</table>
					</td>
				<td width="20%">&nbsp;</td>
				<td width="20%" bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td align="left"><font class="ct">Vencimento</font></td>
					</tr>
					<tr>
						<td align="center"><font class="cp"><%=var_datavencimento%></font></td>
					</tr>
				</table></td>
			</tr>
			<tr>
				<td width="20%" height="10" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><font class="ct">N&ordm; Cota</font></td>
						</tr>
					<tr>
						<td align="center"><font align="center" class="cn">&nbsp;</font><font class="cn"><%=var_numerodoc%></font></td>
						</tr>
				</table></td>
				<td width="20%" height="10" valign="top">
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td><font class="ct">N&ordm; Carne</font></td>
						</tr>
						<tr>
							<td align="center"><font align="center" class="cn">&nbsp;</font><font align="center" class="cn"><%=nu_carne%></font></td>
						</tr>
					</table>
					</td>
				<td width="20%" height="10" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
							<td><font class="ct">Nosso N&uacute;mero</font></td>
						</tr>
					<tr>
							<td align="center"><font align="center" class="cn">&nbsp;</font><font class="cn"><%=var_nossonumero%></font></td>
						</tr>
				</table>
					</td>
				<td width="20%" height="10" valign="top" bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td align="left"><font class="ct">(=) Valor Cobrado</font></td>
					</tr>
					<tr>
						<td align="center"><font class="cp"><%=formatcurrency(var_valordocumento)%></font></td>
					</tr>
				</table></td>
			</tr>
		</table></TD>
	</TR>
	<TR>
		<TD colspan="2" align=right><FONT class=ct>Autentica&ccedil;&atilde;o Mec&acirc;nica<b class="wp"> - Ficha de Compensação</b></FONT><BR><BR><BR></TD>
		</TR>
</TABLE>

<img src="../../../../img/boleto_itau/corte.gif" border=0 width="660"><br><br>

<TABLE WIDTH="660" BORDER=0 CELLSPACING=0 CELLPADDING=0>
  <tr>
		<td class=cp width=150><div align="left" class="ld"><img src="../../../../img/boleto_itau/logobanco.gif" width="29" height="26"><strong>Banco Ita&uacute; S.A</strong>.</div></td>
  		<td width=3 valign="bottom"><img height=22 src="../../../../img/boleto_itau/barra.gif" width=2 border=0></td>
	  	<td class=cpt  width=58 valign="bottom"><div align="center"><font class="bc"><%=cons_banco%>-<%=cons_dvbanco%></font></div></td>
  		<td width=3 valign="bottom"><img height=22 src="../../../../img/boleto_itau/barra.gif" width=2 border=0></td>
	  	<td class=ld align=right width=453 valign="bottom"><span class='ld'><p align="right">&nbsp;<%=var_linhadigitavel%></span></td>
  </tr>
</TABLE>
<TABLE WIDTH="660" BORDER=1 CELLSPACING=0 CELLPADDING=1>
  <TR>
			<TD COLSPAN=5 WIDTH=500>
					<FONT class=ct>Local de Pagamento</FONT><BR>
					<FONT class=cp>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ATÉ O VENCIMENTO PAGÁVEL EM QUALQUER BANCO</FONT>
			</TD>
			<TD width=170 bgcolor="#CCCCCC">
			   <TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0>
				     <TR><TD align=left><FONT class=ct>Vencimento</FONT></TD></TR>
						 <TR><TD align=center><FONT class=cp><%=var_datavencimento%></FONT></TD></TR>
				 </TABLE>
			</TD>
	</TR>
	<TR>
			<TD COLSPAN=5 WIDTH=500><FONT class=ct>Cedente</FONT><BR><FONT class=cn>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=cons_cedente%></FONT></TD>
			<TD width=170>
			   <TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0>
				     <TR><TD align=left><FONT class=ct>Ag&ecirc;ncia / C&oacute;digo Cedente</FONT></TD></TR>
						 <TR><TD align=center><FONT class=cn><%=cons_agencia%>&nbsp;/&nbsp;<%=cons_conta%><!---<%=dvagconta%>--></FONT></TD></TR>
				 </TABLE>
			</TD>
	</TR>
	<TR>
			<TD valign=top>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><font class="ct">Data Documento</font></td>
					</tr>
					<tr>
						<td align="center"><font align="center" class="cn">&nbsp;</font><font class="cn"><%=var_datadocumento%></font></td>
					</tr>
			</table></TD>
			<TD valign=top>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><font class="ct">N&uacute;mero Documento</font></td>
					</tr>
					<tr>
						<td align="center"><font align="center" class="cn">&nbsp;</font><font class="cn"><%=var_numerodoc%></font></td>
					</tr>
			</table></TD>
			<TD valign=top><table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td><font class="ct">Tipo Docu.</font></td>
				</tr>
				<tr>
					<td align="center"><font class="cn">DP</font></td>
				</tr>
			</table></TD>
			<TD valign=top>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><font class="ct">Aceite</font></td>
					</tr>
					<tr>
						<td align="center"><font class="cn">N</font></td>
					</tr>
			</table></TD>
			<TD valign=top>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><font class="ct">Data Processamento</font></td>
					</tr>
					<tr>
						<td align="center"><font align="center" class="cn">&nbsp;</font><font class="cn"><%=var_data%></font></td>
					</tr>
			</table></TD>
			<TD width=170>
			   <TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0>
				     <TR><TD align=left><FONT class=ct>Nosso Número</FONT></TD></TR>
						 <TR><TD align=center><FONT class=cn><%=var_nossonumero%></FONT></TD></TR>
				 </TABLE>
			</TD>
	</TR>
	<TR>
			<TD valign=top><table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><font class="ct">Uso Banco</font></td>
					</tr>
					<tr>
						<td align="center"><font align="center" class="cn">&nbsp;</font></td>
					</tr>
				</table>
			</TD>
			<TD valign=top>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><font class="ct">Carteira</font></td>
					</tr>
					<tr>
						<td align="center"><font align="center" class="cn">&nbsp;</font><font class="cn"><%=cons_carteira%></font></td>
					</tr>
			</table></TD>
			<TD valign=top>
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><font class="ct">Moeda</font></td>
					</tr>
					<tr>
						<td align="center"><font align="center" class="cn">&nbsp;</font><font class="cn"><%=cons_especie%></font></td>
					</tr>
			</table></TD>
			<TD valign=top><table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><font class="ct">Quantidade</font></td>
					</tr>
					<tr>
						<td align="center"><font align="center" class="cn">&nbsp;</font></td>
					</tr>
				</table>
			</TD>
			<TD valign=top><table width="100%" border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td><font class="ct">Valor</font></td>
					</tr>
					<tr>
						<td align="center"><font align="center" class="cn">&nbsp;</font></td>
					</tr>
				</table>
			</TD>
			<TD width=170 bgcolor="#CCCCCC">
			   <TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0>
				     <TR><TD align=left><FONT class=ct>Valor do Documento</FONT></TD></TR>
						 <TR><TD align=center><FONT class=cp><%=formatcurrency(var_valordocumento)%></FONT></TD></TR>
				 </TABLE>
			</TD>
	</TR>
	<TR>
			<TH COLSPAN=5 ROWSPAN=5 valign=top align=LEFT ><FONT class=ct>Instru&ccedil;&otilde;es</FONT><BR>
 				<TABLE WIDTH="475" ALIGN=RIGHT CELLSPACING=0 CELLPADDING=0 BORDER=0>
					<TR>
						<TD valign=top align=left>
							<FONT class=cn>
								<%=var_instrucoes%>
							</FONT>
						</TD>
					</TR>
				</TABLE>
			</TH>
			<TD WIDTH=170><FONT class=ct>(-) Desconto / Abatimento</FONT><BR><FONT class=cn>&nbsp;</FONT></TD>
	</TR>
	<TR>
			<TD WIDTH=170><FONT class=ct>(-) Outras Deduções</FONT><BR><FONT class=cn>&nbsp;</FONT></TD>
	</TR>
	<TR>
		<td width="170"><font class="ct">(+) Mora / Multa</font><br />
			<font class="cn">&nbsp;</font></td>
		</TR>
	<TR>
			<TD WIDTH=170><FONT class=ct>(+) Outros Acr&eacute;scimos</FONT><BR><FONT class=cn>&nbsp;</FONT></TD>
	</TR>
	<TR>
			<TD WIDTH=170 bgcolor="#CCCCCC"><FONT class=ct>(=) Valor Cobrado</FONT><BR><FONT class=cn>&nbsp;</FONT></TD>
	</TR>
	<TR>
			<TD COLSPAN=6 valign=top>
					<FONT class=ct>Sacado</FONT><BR>
						<TABLE WIDTH="640" ALIGN=RIGHT CELLSPACING=0 CELLPADDING=0 BORDER=0>
							<TR>
								<TD width="419" align=left valign=top>
									<FONT class=cn>
										<b><%=var_sacado%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CPF:&nbsp;<%=cpf_responsavel%></b><BR>
										<%=var_endereco%><BR>
										<%=var_bairro%><BR>
					<%=var_cep%>&nbsp;-&nbsp;<%=var_cidade%>&nbsp;-&nbsp;<%=var_estado%><BR>
										<!--<%=var_cpfcnpj%>--><BR>
									</FONT>
	 							</TD>
								<TD width="221" align=left valign=top><table width="100%" border="0" cellspacing="0" cellpadding="0">
									<tr>
										<td width="56%" height="5"><FONT class=cn>Matr&iacute;cula : <B><%response.Write(cod_cons)%></B></FONT></td>
										<td width="44%" height="5"></td>
									</tr>
									<tr>
										<td height="5"><FONT class=cn><%response.Write(no_curso)%></FONT></td>
										<td rowspan="2"><FONT class=cp><%
										vetor_nome = SPLIT(nome_aluno, " ")
										response.Write(vetor_nome(0))%></FONT></td>
									</tr>
									<tr>
										<td height="5"><FONT class=cn><%response.Write(no_etapa)%></FONT></td>
									</tr>
									<tr>
										<td height="5"><FONT class=cn><%response.Write(turma)%></FONT></td>
										<td height="5"></td>
									</tr>
								</table></TD>
							</TR>
						</TABLE>
			</TD>
	</TR>
</TABLE>
<TABLE WIDTH="660" CELLSPACING=0 CELLPADDING=0 BORDER=0>
	<TR>
			<TD class=ct align=right>
  				<div align="right">Autenticação Mecânica - <b class="cp">Ficha de Compensação</b></div>
			</TD>
	</TR>
	<TR>
			<TD align=left>
					<%
						call wbarcode(var_codigobarras)
					%>
			</TD>
	</TR>
	<TR>
		<TD height="25" align=left valign="bottom"><img src="../../../../img/boleto_itau/corte.gif" width="660" height="16" border="0" /></TD>
	</TR>
	<TR>
		<TD align=center><FONT class=cn>Recortar na linha pontilhada abaixo do código de barras</FONT></TD>
	</TR>
</TABLE>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="380">&nbsp;</td>
	</tr>
</table>
<BR>

<%next%>

<!--<table width='640' cellspacing=5 cellpadding=0 border=0>
	<tr>
			<form name='forma'>
						<td align="center"><input type="button" id="print_button" value=' Imprimir Boleto' onClick='x86()' name='print_button'></td>
			</form>
	</tr>
</table>-->

</CENTER>
</body>
</HTML>
