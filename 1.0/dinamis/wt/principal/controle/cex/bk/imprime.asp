<%'On Error Resume Next%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">

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
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->

<%opt = REQUEST.QueryString("obr")
dados_opt= split(opt, "?" )
cod= dados_opt(0)
periodo_check= dados_opt(1)
mes_parcela = REQUEST.QueryString("p")
if mes_parcela = "nulo" then
	sql_mes = "SELECT * FROM TB_Posicao WHERE CO_Matricula_Escola ="& cod 
else
	sql_mes = "SELECT * FROM TB_Posicao WHERE CO_Matricula_Escola ="& cod &" AND Mes = "&mes_parcela
end if

 	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0

 	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1
	
		Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4 = "DBQ="& CAMINHO_pf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4		

	SQL2 = "select * from TB_Alunos where CO_Matricula = " & cod 
	set RS2 = CON1.Execute (SQL2)
	
nome_aluno= RS2("NO_Aluno")

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod
		RS.Open SQL, CON1


		if RS.EOF then
			existe = "N"
		else
			existe = "S"				
			ano_aluno = RS("NU_Ano")
			rematricula = RS("DA_Rematricula")
			situacao = RS("CO_Situacao")
			encerramento= RS("DA_Encerramento")
			unidade= RS("NU_Unidade")
			curso= RS("CO_Curso")
			etapa= RS("CO_Etapa")
			turma= RS("CO_Turma")
			cham= RS("NU_Chamada")
							  
			call GeraNomes("PORT",unidade,curso,etapa,CON0)
			no_unidade = session("no_unidades")
			no_curso = session("no_grau")
			no_etapa= session("no_serie")
		


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
data_compara=data
horario = hora & ":"& minwrt

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("TX_Imp_Cabecalho")
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
end if	
	%>



<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('window.print()')"> 
<br>
<table width="950" border="0" align="center" cellspacing="0" class="tb_corpo"
>
  <tr> 
    <td width="122" height="15" bgcolor="#FFFFFF">
<div align="center"><img src="../../../../img/logo_preto.gif" width="100" height="101"> 
      </div></td>
    <td width="828" bgcolor="#FFFFFF"><table width="100%" border="0" align="right" cellspacing="0">
        <tr> 
          <td width="29%" rowspan="2"><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong>
          	<%
			no_unidade= ucase(no_unidade)
			response.Write(no_unidade)
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
          <td colspan="2"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>EXTRATO 
            FINANCEIRO </strong></font></td>
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
          <td width="68" height="12" bgcolor="#EEEEEE"> <div align="right"> </div></td>
          <td width="198">&nbsp; </td>
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
    <td colspan="2"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="100" class="tabela" 
><div align="center">DATA VENCIMENTO</div></td>
          <td width="232" class="tabela" 
> 
            <div align="center">SERVI&Ccedil;O</div></td>
          <td width="166" align="right" class="tabela" 
> 
           VALOR A PAGAR&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td width="167" align="right" class="tabela" 
> 
            VALOR PAGO&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
          <td width="166" class="tabela" 
> 
            <div align="center">DATA PAGAMENTO</div></td>
          <td width="166" class="tabela" 
> 
            <div align="center">SITUA&Ccedil;&Atilde;O</div></td>
        </tr>
<%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4= sql_mes
		RS4.Open SQL4, CON4
		
if 	existe = "N"then
%>
                <tr> 
                  <td colspan="6" class="tabela"> <div align="center"><br>
                      <br>
                      <br>
                      <br>
                      <br>
                      Este aluno não está ativo neste ano letivo.<br>
                      <br>
                      <br>
                      <br>
                      </div></td>
                </tr>	
		
<%
elseif RS4.EOF then
compromisso=""
da_vencimento=""
realizado=""
da_realizado=""
nome_lanc=""
da_vencimento_show=""
da_realizado_show=""		
		%>
		                      <tr> 
                        <td colspan="6" class="tabela" > <div align="center">
                            Informações não disponíveis para esse aluno. Informe 
                            a secretaria da Escola.</div></td>
                      </tr>
 <%else
check = 1
compromisso_total=0
realizado_total=0
da_vencimento_check = "01/01/1900"
while not RS4.EOF
	compromisso=RS4("VA_Compromisso")
	da_vencimento=RS4("DA_Vencimento")
	realizado=RS4("VA_Realizado") 
	da_realizado=RS4("DA_Realizado")
	nome_lanc=RS4("NO_Lancamento")

	if da_vencimento_check<>da_vencimento then
		check=check+1
		da_vencimento_check = da_vencimento	
	end if		
	
	if isnull(compromisso) or compromisso="" then
		compromisso=0
	end if	
	if isnull(realizado) or realizado="" then
		realizado=0
	end if		
	
	compromisso_total=compromisso_total+compromisso
	realizado_total=realizado_total+realizado	
	
	if realizado = 0 OR ISNULL(realizado) then
		realizado="&nbsp;"
	else
		if realizado<compromisso then
			situacao="Parcela Paga**"
		else
			situacao="Parcela Paga"
		end if
		realizado=FormatNumber(realizado)
	end if

	venc_split=split(da_vencimento,"/")
	dia_venc=venc_split(0)
	mes_venc=venc_split(1)
	ano_venc=venc_split(2)
	venc=mes_venc&"/"&dia_venc&"/"&ano_venc

	if dia_venc<10 then
		dia_venc="0"&dia_venc
	else
		dia_venc=dia_venc
	end if
	
	if mes_venc<10 then
		mes_venc="0"&mes_venc
	else
		mes_venc=mes_venc
	end if
	
	da_vencimento_show=dia_venc&"/"&mes_venc&"/"&ano_venc
	
	venc=replace(da_vencimento,"/","$wxg$adn$")
	
	if isnull(da_realizado) then
		d_diff=DateDiff("d",data_compara,da_vencimento)
		situacao="Parcela Não Paga"
		da_realizado_show=""
	else
		real_split=split(da_realizado,"/")
		dia_real=real_split(0)
		mes_real=real_split(1)
		ano_real=real_split(2)
		real=mes_real&"/"&dia_real&"/"&ano_real
		
		if dia_real<10 then
			dia_real="0"&dia_real
		else
			dia_real=dia_real
		end if
		
		if mes_real<10 then
			mes_real="0"&mes_real
		else
			mes_real=mes_real
		end if
	
		da_realizado_show=dia_real&"/"&mes_real&"/"&ano_real
		
		d_diff=DateDiff("d",da_realizado,da_vencimento)
		
	end if
	if isnull(da_realizado) and d_diff<0 then
	  situacao="Parcela Vencida"
	end if

 if check mod 2 =0 then
  cor = "tb_fundo_linha_par_impr" 
  else cor ="tb_fundo_linha_impar_impr"
  end if  
%>		
        <tr> 
          <td width="100" class="<%response.Write(cor)%>" 
>             <div align="center"> 
            <% 
			if da_vencimento_show="" or isnull(da_vencimento_show) then
			response.Write("&nbsp;")
			else
			response.Write(da_vencimento_show)
			end if
			%></div>
          </td>
          <td width="232" class="<%response.Write(cor)%>" 
> 
            <div align="center"> 
            <% 
			if nome_lanc="" or isnull(nome_lanc) then
			response.Write("&nbsp;")
			else
			response.Write(nome_lanc)
			end if
			%>
            </div></td>
          <td width="166" align="right" class="<%response.Write(cor)%>" 
> 
            <% 
			if compromisso="" or isnull(compromisso) then
			response.Write("&nbsp;")
			else
			response.Write(FormatNumber(compromisso)&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
			end if
			%>
            </td>
          <td width="166" align="right" class="<%response.Write(cor)%>" 
> 
                            <% response.Write(realizado&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")%>
            </td>
          <td width="166" class="<%response.Write(cor)%>" 
> 
            <div align="center"> 
            <% 
			if da_realizado_show="" or isnull(da_realizado_show) then
			response.Write("&nbsp;")
			else
			response.Write(da_realizado_show)
			end if
			%>
            </div></td>
          <td width="166" class="<%response.Write(cor)%>" 
> 
            <div align="center"> 
            <% 
			if situacao="" or isnull(situacao) then
			response.Write("&nbsp;")
			else
			response.Write(situacao)
			end if
			%>
            </div></td>
        </tr>
        <%
RS4.MOVENEXT
WEND
END IF

 if check mod 2 =0 then
  cor = "tb_fundo_linha_par_impr" 
  else cor ="tb_fundo_linha_impar_impr"
  end if  
  %>
                <tr class="<% = cor %>"> 
                  <td width="100" align="center"><b>Total</b></td>
                  <td width="232" align="center">&nbsp;</td>
                  <td width="166" align="right"><b><%response.Write(FormatCurrency(compromisso_total)&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")%></b></td>
                  <td width="166" align="right"><b><%response.Write(FormatCurrency(realizado_total)&"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")%></b></td>
                  <td width="166" align="center">&nbsp;</td>
                  <td width="166" align="center">&nbsp;</td>
                </tr>
      </table></td>
  </tr>
  <tr bgcolor="#FFFFFF"> 
    <td height="10" colspan="2"></td>
  </tr>
  <tr> 
    <td colspan="2" class="linhaTopoL">
<div align="right"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Sistema 
              Diretor - WEB DIRETOR</font> </td>
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