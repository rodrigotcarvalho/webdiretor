<%On Error Resume Next%>
<!--#include file="../inc/funcoes.asp"-->





<%
cod= request.QueryString("cod")
tp_r = request.QueryString("tp")


id0 = " > <a href='cadastro.asp?opt=sel&or=01' class='caminho'>Atualizar Cadastro</a>"
id = " > <a href='altera.asp?opt=vt&or=01&cod="&cod&"' class='caminho'>Consultar Aluno</a>"
id1 = " > Consultar Contato"


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Contatos WHERE TP_Contato = '"&tp_r&"' AND CO_Matricula ="& cod
		RS.Open SQL, CONCONT
		
		
		Set RSCONTPR = Server.CreateObject("ADODB.Recordset")
		SQLCONTPR = "SELECT * FROM TB_Tipo_Contatos WHERE TP_Contato = '"&tp_r&"'"
		RSCONTPR.Open SQLCONTPR, CON0
		
no_tp_resp = RSCONTPR("TX_Descricao")


nome_prof = RS("NO_Contato")


nascimento = RS("DA_Nascimento_Contato")

if isnull(nascimento) then
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
dia_a = dia
mes_a = mes
ano_a = ano
nasce=nascimento
else

vetor_nascimento = Split(nascimento,"/")  
dia_n = vetor_nascimento(0)
mes_n = vetor_nascimento(1)
ano_n = vetor_nascimento(2)

if dia<10 then 
dia = "0"&dia
end if

if mes<10 then
mes = "0"&mes
end if
dia_a = dia_n
mes_a = mes_n
ano_a = ano_n

nasce = dia_n&"/"&mes_n&"/"&ano_n
end if

cpf= RS("CO_CPF_PFisica")
rg= RS("CO_RG_PFisica")
emitido= RS("CO_OERG_PFisica")
emissao = RS("CO_DERG_PFisica")
profissao = RS("NO_Profissao")
empresa = RS("NO_Empresa")
rua=RS("NO_Logradouro_Res")
numero = RS("NU_Logradouro_Res")
complemento = RS("TX_Complemento_Logradouro_Res")
bairro= RS("CO_Bairro_Res")
municipio= RS("CO_Municipio_Res")
uf= RS("SG_UF_Res")
cep = RS("CO_CEP_Res")
telefone = RS("NU_Telefones_Res")
rua2=RS("NO_Logradouro_Com")
numero2 = RS("NU_Logradouro_Com")
complemento2 = RS("TX_Complemento_Logradouro_Com")
bairro2= RS("CO_Bairro_Com")
municipio2= RS("CO_Municipio_Com")
uf2= RS("SG_UF_Com")
cep2 = RS("CO_CEP_Com")
telefone2 = RS("NU_Telefones_Com")
mail= RS("TX_EMail")

if isnull(pais) then 
pais = 10
end if

if isnull(uf) then 
uf = "RJ"
end if

if isnull(municipio) then 
municipio = 6001
end if

if isnull(uf_natural) then 
uf_natural = "RJ"
end if

if isnull(nacionalidade) then 
nacionalidade = 1
end if

if isnull(natural) then 
natural = 6001
end if

if complemento = "nulo" then 
complemento = ""
end if

if pai_fal = false then
pai_fal = "Não"
else
pai_fal = "Sim"
end if

if mae_fal = false then
mae_fal = "Não"
else
mae_fal = "Sim"
end if


		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Bairros WHERE CO_Bairro ="& bairro &"AND SG_UF ='"& uf&"' AND CO_Municipio = "&municipio
		RS0.Open SQL0, CON0
if RS0.EOF then
bairro = ""
else
bairro= RS0("NO_Bairro")
end if
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Bairros WHERE CO_Bairro ="& bairro2 &"AND SG_UF ='"& uf2&"' AND CO_Municipio = "&municipio2
		RS1.Open SQL1, CON0

bairro2= RS1("NO_Bairro")


		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Municipios WHERE SG_UF ='"& uf&"' AND CO_Municipio = "&municipio
		RS2.Open SQL2, CON0

municipio= RS2("NO_Municipio")

		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Municipios WHERE SG_UF ='"& uf2&"' AND CO_Municipio = "&municipio2
		RS3.Open SQL3, CON0

municipio2= RS3("NO_Municipio")

		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_UF WHERE SG_UF ='"& uf&"'"
		RS4.Open SQL4, CON0

uf = RS4("NO_UF")

		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_UF WHERE SG_UF ='"& uf2&"'"
		RS5.Open SQL5, CON0

uf2 = RS5("NO_UF")

cep = cep/1000
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

cep=cep3&"-"&cep4

cep2 = cep2/1000
cep32=Int(cep2)
cep42= cep2-cep32

cep42=cep42*1000
cep42 = int(cep42)

if cep42 = 0 then
cep42="000"
elseif cep42<10 then
cep42="00"&cep42
elseif cep42>=10 And cep42<100 then
cep42="0"&cep42
end if

cep2=cep32&"-"&cep42



	%>
<html>
<head>
<title>Web Acad&ecirc;mico</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../js/mm_menu.js"></script>
<script type="text/javascript" src="../js/atualiza_select.js"></script>
<script type="text/javascript" src="../js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function checksubmit()
{
  if (document.alteracao.nome.value == "")
  {    alert("Por favor, digite um nome para o professor!")
    document.alteracao.nome.focus()
    return false
  }
//  if (document.alteracao.nasce.value == "")
//  {    alert("Por favor, digite a data de nascimento do professor!")
//    document.alteracao.nasce.focus()
//return false
//}
erro=0;
        hoje = new Date();
         anoAtual = hoje.getFullYear();
         barras = alteracao.nasce.value.split("/");
         if (barras.length == 3){
                   dia = barras[0];
                   mes = barras[1];
                   ano = barras[2];
                   resultado = (!isNaN(dia) && (dia > 0) && (dia < 32)) && (!isNaN(mes) && (mes > 0) && (mes < 13)) && (!isNaN(ano) && (ano.length == 4) && (ano <= anoAtual && ano >= 1900));
                   if (!resultado) {
                             alert("Formato de data invalido!");
                             alteracao.nasce.focus();
                             return false;
                   }
         } else {
                   alert("Formato de data invalido!");
                   alteracao.nasce.focus();
                   return false;
         }
  if (document.alteracao.sexo.value == "0")
  {    alert("Por favor, escolha o sexo do professor!")
    document.alteracao.sexo.focus()
    return false
  }   
  if (document.alteracao.rua.value == "")
  {    alert("Por favor, digite a rua onde o professor reside!")
    document.alteracao.rua.focus()
    return false
  }    
erro=0;

         barras = alteracao.cep.value.split("-");
         if (barras.length == 2){
                   cep0= barras[0];
                   cep1 = barras[1];
                   resultado = (!isNaN(dia) && (cep0 > 10000) && (cep0 < 999999)) && (!isNaN(mes) && (cep1 >= 0) && (cep1 < 999));
                   if (!resultado) {
                             alert("Formato do CEP invalido!");
                             alteracao.cep.focus();
                             return false;
                   }
         } else {
                   alert("Formato de cep invalido!");
                   alteracao.cep.focus();
                   return false;
         }
  if (document.alteracao.telefones.value == "")
  {    alert("Por favor, digite pelo menos um telefone para contato com o professor!")
    document.alteracao.telefones.focus()
    return false
  }                  	     
  return true
}

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('document.alteracao.nome.focus()');alinhamento()" onresize="alinhamento()">
<%call cabecalho(nivel)
%>
<table width="1000" border="0" align="center" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td><div align="center">table width="1000" border="0">
  <tr> 
    <td><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="../inicio.asp" class="caminho">Web 
      Acad&ecirc;mico</a> 
      <%response.Write(origem&id0&id&id1)%>
      
      </strong> </font></td>
  </tr>
</table>
<br>
<table width="1000" border="0" cellspacing="0">
  <tr> 
    <td width="219" valign="top">
<table width="100%" border="0" cellspacing="0">
        <tr>
          <td>      
<%
if opt = "ok" then
	call mensagens(6,2,cod)
elseif erro = "dt" then
	call mensagens(7,1,0)
elseif erro = "nb" then
	call mensagens(8,1,0)	
elseif erro = "cp" then
	call mensagens(9,1,0)	
end if
%>
</td>
        </tr>
        <tr>
          <td>
<%
	call mensagens(1005,0,0) 

%>		  		  
		  </td>
        </tr>
      </table>
      
    </td>
    <td width="785" valign="top"> 	  
        <table width="770" border="0" align="right" cellspacing="0" class="tb_corpo"
>
        <tr> 
          <td width="653" class="tb_tit"
><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Dados 
            Pessoais</strong></font></td>
          <td width="113" class="tb_tit"
> </td>
        </tr>
        <tr> 
          <td height="10" colspan="2"> <table width="100%" border="0" cellspacing="0">
              <tr> 
                <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nome: 
                    </strong></font></div></td>
                <td height="10" colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(nome_prof)%>
                  </font> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                <td width="13%" height="10"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
                    </font><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Rela&ccedil;&atilde;o:</strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    </font></div></td>
                <td width="19%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(no_tp_resp)%>
                  </font></td>
              </tr>
              <tr> 
                <td width="19%" height="10"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
                    de Nascimento: </strong></font></div></td>
                <td width="27%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(nasce)%>
                  </font> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <input name="nasce" type="hidden" class="textInput" id="nasce4" value="<%response.Write(nasce)%>" size="12" maxlength="10">
                  <strong> &nbsp;-&nbsp; <font color="#CC9900"> 
                  <%
					call aniversario(ano_a,mes_a,dia_a) %>
                  </font> </strong></font></td>
                <td width="22%" height="10">
				</td>
                <td height="10" colspan="2">
				</td>
              </tr>
              <tr> 
                <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Profiss&atilde;o:</strong></font></div></td>
                <td height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <% response.Write(profissao)%>
                  </font></td>
                <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Empresa 
                    onde trabalha:</strong></font></div></td>
                <td height="10" colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(empresa)%>
                  </font></td>
              </tr>
              <tr> 
                <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Endere&ccedil;o 
                    Eletr&ocirc;nico:</strong></font></div></td>
                <td height="10" colspan="4"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(mail)%>
                  </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
                  </font></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="10" colspan="2"> <table width="100%" border="0" cellspacing="0">
              <tr> 
                <td height="10" width="13%"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>CPF:</strong></font></div></td>
                <td height="10" width="18%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(cpf)%>
                  </font></td>
                <td width="18%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                    Identidade: </strong></font></div></td>
                <td width="30%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(rg)%>
                  &nbsp;</font></td>
                <td width="12%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipo:</strong></font></div></td>
                <td width="12%" height="10"> <div align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    <%response.Write(emitido)%>
                    </font><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                    </strong></font></div></td>
                <td width="12%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data:</strong></font></div></td>
                <td width="12%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(emissao)%>
                  </font></td>
              </tr>
              <tr> 
                <td height="10" colspan="2"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefones 
                    de Contato:</strong></font></div></td>
                <td height="10" colspan="6"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(tel_cont)%>
                  </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="10" bgcolor="#FFFFFF">&nbsp;</td>
          <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
        <tr> 
          <td colspan="2" class="tb_tit"
><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Endere&ccedil;o 
            Residencial</strong></font></td>
        </tr>
        <tr> 
          <td height="10" colspan="2"> <table width="100%" height="10" border="0" cellspacing="0">
              <tr class="tb_corpo"
> 
                <td width="11%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Logradouro:</strong></font></div></td>
                <td width="60%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(rua)%>
                  <input name="rua" type="hidden" class="textInput" id="rua4" value="<%response.Write(rua)%>" size="75" maxlength="50">
                  </font></td>
                <td width="13%" height="10"> <div align="right"></div>
                  <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>N&uacute;mero:</strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    </font></div></td>
                <td width="16%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(numero)%>
                  <input name="numero" type="hidden" class="textInput" id="numero4" value="<%response.Write(numero)%>" size="11" maxlength="6">
                  &nbsp; </font> </td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="10" colspan="2"> <table width="100%" height="10" border="0" cellspacing="0">
              <tr class="tb_corpo"
> 
                <td width="12%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Complemento:</strong></font></div></td>
                <td width="36%" height="10"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(complemento)%>
                  </font> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <input name="complemento" type="hidden" class="textInput" id="complemento4" value="<%response.Write(complemento)%>" size="45" maxlength="30">
                  </font></td>
                <td width="6%" height="10"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
                    </font><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>CEP:</strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
                    </font></div></td>
                <td width="13%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(cep)%>
                  <input name="cep" type="hidden" class="textInput" id="cep7" value="<%response.Write(cep)%>" size="11" maxlength="9" onKeyup="formatar(this, '#####-###')">
                  </font></td>
                <td width="7%" height="10">
                  <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Estado:</strong></font> 
                  </div></td>
                <td width="26%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(uf)%>
                  </font> </td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="10" colspan="2"> <table width="100%" height="10" border="0" cellspacing="0">
              <tr class="tb_corpo"
> 
                <td width="16%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cidade:</strong></font></div></td>
                <td width="29%" height="10"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(municipio)%>
                  </font></td>
                <td width="11%" height="10">
                  <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro:</strong></font> 
                  </div></td>
                <td width="44%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(bairro)%></font></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="10" colspan="2"> <table width="100%" height="10" border="0" cellspacing="0">
              <tr class="tb_corpo"
> 
                <td width="20%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefones 
                    deste endere&ccedil;o:</strong></font></div></td>
                <td width="80%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(telefone)%>
                  <input name="telefones" type="hidden" class="textInput" id="telefones4" value="<%response.Write(telefone)%>" size="75" maxlength="50">
                  </font></td>
              </tr>
            </table></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
          <td height="10" colspan="2">&nbsp; </td>
        </tr>
        <tr> 
          <td colspan="2" class="tb_tit"
><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Endere&ccedil;o 
            Comercial </strong></font></td>
        </tr>
        <tr> 
          <td height="10" colspan="2"> <table width="100%" border="0" cellspacing="0">
              <tr class="tb_corpo"
> 
                <td width="11%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Logradouro:</strong></font></div></td>
                <td width="60%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(rua2)%>
                  <input name="rua" type="hidden" class="textInput" id="rua4" value="<%response.Write(rua2)%>" size="75" maxlength="50">
                  </font></td>
                <td width="13%" height="10"> <div align="right"></div>
                  <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>N&uacute;mero:</strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    </font></div></td>
                <td width="16%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(numero2)%>
                  <input name="numero" type="hidden" class="textInput" id="numero4" value="<%response.Write(numero2)%>" size="11" maxlength="6">
                  &nbsp; </font> </td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="10" colspan="2"> <table width="100%" border="0" cellspacing="0">
              <tr class="tb_corpo"
> 
                <td width="12%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Complemento:</strong></font></div></td>
                <td width="36%" height="10"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(complemento2)%>
                  </font> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <input name="complemento" type="hidden" class="textInput" id="complemento4" value="<%response.Write(complemento2)%>" size="45" maxlength="30">
                  </font></td>
                <td width="6%" height="10"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
                    </font><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>CEP:</strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
                    </font></div></td>
                <td width="13%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(cep2)%>
                  <input name="cep" type="hidden" class="textInput" id="cep7" value="<%response.Write(cep2)%>" size="11" maxlength="9" onKeyup="formatar(this, '#####-###')">
                  </font></td>
                <td width="7%" height="10"> <div align="right"></div>
                  <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Estado:</strong></font> 
                  </div></td>
                <td width="26%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(uf2)%>
                  </font> </td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="10" colspan="2"> <table width="100%" border="0" cellspacing="0">
              <tr class="tb_corpo"
> 
                <td width="16%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cidade:</strong></font></div></td>
                <td width="29%" height="10"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(municipio2)%>
                  </font></td>
                <td width="11%" height="10"> <div align="right"></div>
                  <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro:</strong></font> 
                  </div></td>
                <td width="44%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(bairro2)%>
                  </font></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="10" colspan="2"> <table width="100%" border="0" cellspacing="0">
              <tr class="tb_corpo"
> 
                <td width="20%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefones 
                    deste endere&ccedil;o:</strong></font></div></td>
                <td width="80%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%response.Write(telefone2)%>
                  <input name="telefones" type="hidden" class="textInput" id="telefones4" value="<%response.Write(telefone2)%>" size="75" maxlength="50">
                  </font></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td height="10" colspan="2"> <table width="100%" border="0" cellspacing="0">
              <tr class="tb_corpo"
> 
                <td bgcolor="#FFFFFF"> <div align="right"></div></td>
              </tr>
            </table></td>
        </tr>
        <tr> 
          <td colspan="2" class="tb_tit"
>&nbsp;</td>
        </tr>
        <tr> 
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
        <tr> 
          <td colspan="2" bgcolor="#FFFFFF"> <table width="500" border="0" align="center" cellspacing="0">
              <tr> 
                <td width="50%"> <div align="center"> </div></td>
              </tr>
            </table></td>
        </tr>
      </table>
</td>
  </tr>
</table>
<p>&nbsp;</p>
<table width="1000" border="0" cellspacing="0">
  <tr>
    <td width="238">&nbsp;</td>
    <td width="770" class="tb_voltar"
><font color="#669999" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="altera.asp?opt=vt&or=01&cod=<%=cod%>" class="voltar1">&lt; 
      Voltar para Consultar Aluno</a></strong></font></td>
  </tr>
</table>
<p align="center">&nbsp;</p></div></td>
  </tr>
</table>
<
</body>
<script type="text/javascript">
<!--
  initInputHighlightScript();
//-->
</script>

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