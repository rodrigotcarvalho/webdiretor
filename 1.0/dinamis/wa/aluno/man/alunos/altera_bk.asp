<%On Error Resume Next%>
<!--#include file="../inc/funcoes.asp"-->





<%
cod= request.QueryString("cod")	
opt = request.QueryString("opt")
z = request.QueryString("z")
erro = request.QueryString("e")
vindo = request.QueryString("vd")
obr = request.QueryString("o")

if vindo="crmt" then
dados= split(obr, "_" )
unidade= dados(0)
curso= dados(1)
co_etapa= dados(2)
turma= dados(3)
obr=unidade&"_"&curso&"_"&co_etapa&"_"&turma
end if

id0 = " > <a href='cadastro.asp?opt=sel&or=01' class='caminho'>Selecionar Aluno</a>"
id = " > Consultar Aluno"
idc = " > <a href='../vt/tabelas.asp?or=02&volta=1' class='caminho' target='_self'>Verificar Turmas</a>"
idc1 = " > <a href='../vt/tabelas.asp?or=02&volta=1' class='caminho' target='_self'>Seleciona Unidade</a>"
idc2 = " > <a href='../vt/tabelas2.asp?opt=vt&or=02&curso="&curso&"&unidade="&unidade&"' class='caminho' target='_self'>Seleciona Etapa</a>"
idc3 = " > <a href='../vt/consulta_turma_cp3.asp?or=01&opt=vt&o="&obr&"' class='caminho' target='_self'>Visualizando</a>"
idc4 = " > <a href='../vt/carometro.asp?or=01&opt=vt&o="&obr&"' class='caminho' target='_self'>Carômetro</a>"


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
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod
		RS.Open SQL, CON
		
		
codigo = RS("CO_Matricula")
nome_prof = RS("NO_Aluno")

sexo = RS("IN_Sexo")

		Set RSCONTA = Server.CreateObject("ADODB.Recordset")
		SQLA = "SELECT * FROM TB_Contatos WHERE TP_Contato ='ALUNO' And CO_Matricula ="& cod
		RSCONTA.Open SQLA, CONCONT

if RSCONTA.EOF
nascimento="0/0/0"
else
nascimento = RSCONTA("DA_Nascimento_Contato")
end if
vetor_nascimento = Split(nascimento,"/")  
dia_n = vetor_nascimento(0)
mes_n = vetor_nascimento(1)
ano_n = vetor_nascimento(2)

if dia_n<10 then 
dia_n = "0"&dia_n
end if

if mes_n<10 then
mes_n = "0"&mes_n
end if
dia_a = dia_n
mes_a = mes_n
ano_a = ano_n

nasce = dia_n&"/"&mes_n&"/"&ano_n

apelido= RS("NO_Apelido")
desteridade= RS("IN_Desteridade")
nacionalidade= RS("CO_Nacionalidade")
rua = RSCONTA("NO_Logradouro_Res")
numero = RSCONTA("NU_Logradouro_Res")
complemento = RSCONTA("TX_Complemento_Logradouro_Res")
bairro= RSCONTA("CO_Bairro_Res")
municipio= RSCONTA("CO_Municipio_Res")
pai= RS("NO_Pai")
mae= RS("NO_Mae")
pai_fal= RS("IN_Pai_Falecido")
mae_fal= RS("IN_Mae_Falecida")
uf= RSCONTA("SG_UF_Res")
cep = RSCONTA("CO_CEP_Res")
telefone = RSCONTA("NU_Telefones_Res")
tel_cont = RSCONTA("NU_Telefones")
uf_natural = RS("SG_UF_Natural")
natural = RS("CO_Municipio_Natural")
resp_fin= RS("TP_Resp_Fin")
resp_ped= RS("TP_Resp_Ped")
mail= RSCONTA("TX_EMail")
pais= RS("CO_Pais_Natural")
ocupacao= RSCONTA("CO_Ocupacao")
msn= RS("TX_MSN")
orkut= RS("TX_ORKUT")
religiao= RS("CO_Religiao")
raca= RS("CO_Raca")
entrada= RS("DA_Entrada_Escola")
cadastro= RS("DA_Cadastro")
col_origem= RS("NO_Colegio_Origem")
cursada= RS("NO_Serie_Cursada")
uf_cursada= RS("SG_UF_Cursada")
cid_cursada= RS("CO_Municipio_Cursada")
co_estado_civil= RS("CO_Estado_Civil")
cpf= RSCONTA("CO_CPF_PFisica")
rg= RSCONTA("CO_RG_PFisica")
emitido= RSCONTA("CO_OERG_PFisica")
emissao= RSCONTA("CO_DERG_PFisica")
empresa= RSCONTA("NO_Empresa")
rua2=RSCONTA("NO_Logradouro_Com")
numero2 = RSCONTA("NU_Logradouro_Com")
complemento2 = RSCONTA("TX_Complemento_Logradouro_Com")
bairro2= RSCONTA("CO_Bairro_Com")
municipio2= RSCONTA("CO_Municipio_Com")
uf2= RSCONTA("SG_UF_Com")
cep2 = RSCONTA("CO_CEP_Com")
telefone2 = RSCONTA("NU_Telefones_Com")

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Matriculas WHERE CO_Matricula ="& cod
		RS.Open SQL, CON


ano_aluno = RS("NU_Ano")
rematricula = RS("DA_Rematricula")
situacao = RS("CO_Situacao")
encerramento= RS("DA_Encerramento")
unidade= RS("NU_Unidade")
curso= RS("CO_Curso")
etapa= RS("CO_Etapa")
turma= RS("CO_Turma")
cham= RS("NU_Chamada")



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

if desteridade = "S" then
desteridade = "Destro"
else
desteridade = "Canhoto"
end if

if isnull(cid_cursada) then 
cid_cursada = 6001
end if

if isnull(uf_cursada) then 
uf_cursada = "RJ"
end if


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


		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Religiao WHERE CO_Religiao ="& religiao
		RS0.Open SQL0, CON0

religiao = RS0("TX_Descricao_Religiao")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Raca WHERE CO_Raca ="& raca
		RS1.Open SQL1, CON0

raca = RS1("TX_Descricao_Raca")

		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Paises WHERE CO_Pais ="& pais
		RS2.Open SQL2, CON0

pais = RS2("NO_Pais")

		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Nacionalidades WHERE CO_Nacionalidade ="& nacionalidade
		RS3.Open SQL3, CON0

nacionalidade = RS3("TX_Nacionalidade")



		Set RS6 = Server.CreateObject("ADODB.Recordset")
		SQL6 = "SELECT * FROM TB_Municipios WHERE SG_UF ='"& uf_natural&"' AND CO_Municipio = "&natural
		RS6.Open SQL6, CON0

natural= RS6("NO_Municipio")

		Set RS8 = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT * FROM TB_UF WHERE SG_UF ='"& uf_natural&"'" 
		RS8.Open SQL8, CON0

uf_natural= RS8("NO_UF")

		Set RS9 = Server.CreateObject("ADODB.Recordset")
		SQL9 = "SELECT * FROM TB_Ocupacoes WHERE CO_Ocupacao ="& ocupacao
		RS9.Open SQL9, CON0

ocupacao= RS9("NO_Ocupacao")


		Set RS10 = Server.CreateObject("ADODB.Recordset")
		SQL10 = "SELECT * FROM TB_Estado_Civil WHERE CO_Estado_Civil ='"& co_estado_civil&"'"
		RS10.Open SQL10, CON0

estado_civil= RS10("TX_Estado_Civil")

		Set RS11 = Server.CreateObject("ADODB.Recordset")
		SQL11 = "SELECT * FROM TB_Municipios WHERE SG_UF ='"& uf_cursada &"' AND CO_Municipio = "&cid_cursada
		RS11.Open SQL11, CON0

cid_cursada= RS11("NO_Municipio")

		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Bairros WHERE CO_Bairro ="& bairro &"AND SG_UF ='"& uf&"' AND CO_Municipio = "&municipio
		RS4.Open SQL4, CON0
if RS4.EOF then
bairro = ""
else
bairro= RS4("NO_Bairro")
end if

if isnull(uf2)or isnull(municipio2) then

else
		Set RS4a = Server.CreateObject("ADODB.Recordset")
		SQL4a = "SELECT * FROM TB_Bairros WHERE CO_Bairro ="& bairro2 &"AND SG_UF ='"& uf2&"' AND CO_Municipio = "&municipio2
		RS4a.Open SQL4a, CON0

if RS4a.EOF then
bairro2 = ""
else
bairro2= RS4a("NO_Bairro")
end if

		Set RS5a = Server.CreateObject("ADODB.Recordset")
		SQL5a = "SELECT * FROM TB_Municipios WHERE SG_UF ='"& uf2&"' AND CO_Municipio = "&municipio2
		RS5a.Open SQL5a, CON0

if RS5a.EOF then
municipio2 = ""
else
municipio2= RS5a("NO_Municipio")
end if



		Set RS7a = Server.CreateObject("ADODB.Recordset")
		SQL7a = "SELECT * FROM TB_UF WHERE SG_UF ='"& uf2&"'"
		RS7a.Open SQL7a, CON0

if RS7a.EOF then
uf2 = ""
else
uf2 = RS7a("NO_UF")
end if
end if
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Municipios WHERE SG_UF ='"& uf&"' AND CO_Municipio = "&municipio
		RS5.Open SQL5, CON0
if RS5.EOF then
municipio = ""
else
municipio= RS5("NO_Municipio")
end if

		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_UF WHERE SG_UF ='"& uf&"'"
		RS7.Open SQL7, CON0

if RS7.EOF then
uf = ""
else
uf = RS7("NO_UF")
end if

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

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

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

function MM_showHideLayers() { //v6.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
    obj.visibility=v; }
}
//-->
</script>
</head> 
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('document.alteracao.nome.focus()');alinhamento()" onresize="alinhamento()">
<%call cabecalho(nivel)
%>
<div id="fundo" style="position:absolute; left:0px; top:0px; width:100%; height:100%; z-index:1; background-color: #000000; layer-background-color: #000000; border: 1px none #000000; visibility: hidden;" class="transparente"></div>
<div id="alinha" style="position:absolute; width:400px; visibility: hidden; z-index: 2; left: 326px; height: 520px;"> 
<table width="1000" border="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr> 
      <td width="478"> <div align="right"> <span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')">fechar</a></font></span></div></td>
      <td width="20"><div align="right"><span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')"><img src="../img/fecha.gif" width="20" border="0"></a></font></span></div></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center" ><img src="../img/fotos/aluno/<%response.Write(cod)%>.jpg" height="500"></div></td>
    </tr>
    <tr>
      <td colspan="2"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
          <%response.Write(nome_prof)%>
          </font></div></td>
    </tr>
  </table>
</div>
<table width="1000" height="670" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
    <td><div align="center"> 
        <table width="1000" border="0" class="tb_caminho">
          <tr> 
            <td><font color="#FFFF33" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="../inicio.asp" class="caminho">Web 
              Acad&ecirc;mico</a> 
              <%	  
if vindo="crmt" then
response.Write(origem&idc&idc1&idc2&idc3&idc4&id)
elseif vindo="vt" then
response.Write(origem&idc&idc1&idc2&idc3&id)
else
response.Write(origem&id0&id)
end if	  
	  %>
              </font></td>
          </tr>
        </table>
        <br>
        <table width="1000" border="0" cellspacing="0">
          <tr> 
            <td width="219" valign="top"> <table width="100%" border="0" cellspacing="0">
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
	call mensagens(1002,0,0) 

%>
                  </td>
                </tr>
              </table></td>
            <td width="770" valign="top"> 
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
                  <td height="10"> <table width="100%" border="0" cellspacing="0">
                      <tr> 
                        <td width="19%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Matr&iacute;cula: 
                            </strong></font></div></td>
                        <td width="9%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <input name="cod" type="hidden" value="<%=codigo%>">
                          <%response.Write(codigo)%>
                          <input name="tp" type="hidden" id="tp" value="P">
                          <input name="acesso" type="hidden" id="acesso" value="2">
                          <input name="co_usr_prof" type="hidden" id="co_usr_prof" value="<% =co_usr_prof%>">
                          </font></td>
                        <td width="6%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nome: 
                            </strong></font></div></td>
                        <td height="10" colspan="4"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(nome_prof)%>
                          <input name="nome2" type="hidden" class="textInput" id="nome2"  value="<%response.Write(nome_prof)%>" size="75" maxlength="50">
                          &nbsp; </font></td>
                      </tr>
                      <tr> 
                        <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Apelido:</strong></font></div></td>
                        <td height="10" colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(apelido)%>
                          &nbsp; </font></td>
                        <td width="23%" height="10"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
                            de Nascimento: </strong></font></div></td>
                        <td height="10" colspan="3"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(nasce)%>
                          </font> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <input name="nasce2" type="hidden" class="textInput" id="nasce" value="<%response.Write(nasce)%>" size="12" maxlength="10">
                          <strong> &nbsp;-&nbsp; <font color="#CC9900"> 
                          <%
					call aniversario(ano_a,mes_a,dia_a) %>
                          </font></strong>&nbsp; </font></td>
                      </tr>
                      <tr> 
                        <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Sexo:</strong></font></div></td>
                        <td height="10" colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%if sexo = "M" then
					  response.Write("Masculino")
						else
					  response.Write("Feminino")
                    End IF%>
                          </font></td>
                        <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Pa&iacute;s 
                            de Origem:</strong></font></div></td>
                        <td height="10" colspan="3"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(pais)%>
                          </font></td>
                      </tr>
                      <tr> 
                        <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                            Nacionalidade:</strong></font></div></td>
                        <td height="10" colspan="2"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif">	
                          <%response.Write(nacionalidade)%>
                          </font> </td>
                        <td height="10"> <div align="right"></div>
                          <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#CC9900">Natural 
                            do Estado:</font> </strong></font></div></td>
                        <td height="10" colspan="3"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(uf_natural)%>
                          </font></td>
                      </tr>
                      <tr> 
                        <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Natural 
                            da cidade:</strong></font></div></td>
                        <td height="10" colspan="2"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(natural)%>
                          </font> </td>
                        <td height="10"> <div align="right"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#CC9900">Religi&atilde;o:</font></strong></font> 
                          </div></td>
                        <td height="10" colspan="3"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(religiao)%>
                          </font></td>
                      </tr>
                      <tr> 
                        <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cor 
                            / Ra&ccedil;a: </strong> </font></div></td>
                        <td height="10" colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(raca)%>
                          </font> <div align="right"></div></td>
                        <td height="10"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
                            </font><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Ocupa&ccedil;&atilde;o: 
                            </strong> </font></div></td>
                        <td width="9%" height="10"> <div align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            <%response.Write(ocupacao)%>
                            </font></div></td>
                        <td width="9%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Escrita:</strong></font></div></td>
                        <td width="25%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(desteridade)%>
                          </font></td>
                      </tr>
                      <tr> 
                        <td height="10"><div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>CPF:</strong></font></div></td>
                        <td height="10" colspan="2"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div>
                          <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(cpf)%>
                          </font></td>
                        <td height="10"><div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Empresa 
                            onde trabalha:</strong></font></div></td>
                        <td height="10" colspan="3"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(empresa)%>
                          </font></td>
                      </tr>
                      <tr> 
                        <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                            Identidade: </strong></font></div></td>
                        <td height="10" colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(rg)%>
                          &nbsp;</font> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"></font></div></td>
                        <td height="10"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            </font><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Tipo: 
                            </strong></font></div></td>
                        <td height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(emitido)%>
                          </font></td>
                        <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data:</strong></font></div></td>
                        <td height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(emissao)%>
                          </font></td>
                      </tr>
                      <tr> 
                        <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Endere&ccedil;o 
                            Eletr&ocirc;nico:</strong></font></div></td>
                        <td height="10" colspan="6"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(mail)%>
                          &nbsp; &nbsp;</font></td>
                      </tr>
                      <tr> 
                        <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Login 
                            Orkut:</strong></font></div></td>
                        <td height="10" colspan="3"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(orkut)%>
                          &nbsp; </font></td>
                        <td height="10" colspan="2"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Login 
                            Messenger:</strong></font></div></td>
                        <td height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(msn)%>
                          &nbsp;</font></td>
                      </tr>
                      <tr> 
                        <td height="10" colspan="2"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefones 
                            de Contato:</strong></font></div></td>
                        <td height="10" colspan="5"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(tel_cont)%>
                          &nbsp;</font></td>
                      </tr>
                    </table></td>
                  <td valign="top"> <table height="110" border="3" align="right" cellspacing="0" bordercolor="#EEEEEE">
                      <tr> 
                        <td><div align="center"><a href="#"><img src="../img/fotos/aluno/<% =codigo %>.jpg" alt="" height="110" border="0"></a></div></td>
                      </tr>
                      <tr> 
                        <td height="15" bgcolor="#EEEEEE"> <div align="center"><a href="#" onClick="MM_showHideLayers('fundo','','show','alinha','','show')"><img src="../img/clique.gif" width="85" height="13" border="0"></a></div></td>
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
                  <td height="10" colspan="2"> <table width="100%" border="0" cellspacing="0">
                      <tr> 
                        <td width="15%" height="10" class="tb_corpo"
> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Logradouro:</strong></font></div></td>
                        <td height="10" colspan="3" class="tb_corpo"
><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(rua)%>
                          <input name="rua" type="hidden" class="textInput" id="rua4" value="<%response.Write(rua)%>" size="75" maxlength="50">
                          </font></td>
                        <td width="9%" height="10" class="tb_corpo"
> <div align="right"></div>
                          <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>N&uacute;mero:</strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            </font></div></td>
                        <td width="22%" height="10" class="tb_corpo"
><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(numero)%>
                          <input name="numero" type="hidden" class="textInput" id="numero4" value="<%response.Write(numero)%>" size="11" maxlength="6">
                          &nbsp; </font> </td>
                      </tr>
                      <tr class="tb_corpo"
> 
                        <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Complemento:</strong></font></div></td>
                        <td height="10"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(complemento)%>
                          </font> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <input name="complemento" type="hidden" class="textInput" id="complemento4" value="<%response.Write(complemento)%>" size="45" maxlength="30">
                          &nbsp; </font></td>
                        <td width="4%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>CEP:</strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            </font></div></td>
                        <td width="13%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(cep)%>
                          <input name="cep" type="hidden" class="textInput" id="cep7" value="<%response.Write(cep)%>" size="11" maxlength="9" onKeyup="formatar(this, '#####-###')">
                          </font></td>
                        <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Estado:</strong></font> 
                          </div></td>
                        <td height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(uf)%>
                          </font> </td>
                      </tr>
                      <tr> 
                        <td height="10" class="tb_corpo"
> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cidade:</strong></font></div></td>
                        <td height="10" colspan="2" class="tb_corpo"
> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(municipio)%>
                          </font></td>
                        <td height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td height="10" class="tb_corpo"
> <div align="right"></div>
                          <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro:</strong></font> 
                          </div></td>
                        <td height="10" class="tb_corpo"
><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(bairro)%>
                          </font></td>
                      </tr>
                      <tr> 
                        <td height="10" colspan="6" class="tb_corpo"
><div align="left"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefones 
                            deste endere&ccedil;o:</strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            <%response.Write(telefone)%>
                            <input name="telefones2" type="hidden" class="textInput" id="telefones" value="<%response.Write(telefone)%>" size="75" maxlength="50">
                            </font></div></td>
                      </tr>
                      <tr> 
                        <td height="10" colspan="6" class="tb_corpo"
>&nbsp;</td>
                      </tr>
                      <tr> 
                        <td height="10" colspan="6" class="tb_tit"
><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Endere&ccedil;o 
                          Comercial </strong></font></td>
                      </tr>
                      <tr> 
                        <td height="10" class="tb_corpo"
> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Logradouro:</strong></font></div></td>
                        <td height="10" colspan="3" class="tb_corpo"
><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(rua2)%>
                          <input name="rua2" type="hidden" class="textInput" id="rua" value="<%response.Write(rua2)%>" size="75" maxlength="50">
                          </font></td>
                        <td height="10" class="tb_corpo"
> <div align="right"></div>
                          <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>N&uacute;mero:</strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            </font></div></td>
                        <td height="10" class="tb_corpo"
><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(numero2)%>
                          <input name="numero2" type="hidden" class="textInput" id="numero" value="<%response.Write(numero2)%>" size="11" maxlength="6">
                          &nbsp; </font> </td>
                      </tr>
                      <tr class="tb_corpo"
> 
                        <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Complemento:</strong></font></div></td>
                        <td height="10"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(complemento2)%>
                          </font> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <input name="complemento2" type="hidden" class="textInput" id="complemento" value="<%response.Write(complemento2)%>" size="45" maxlength="30">
                          </font></td>
                        <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>CEP:</strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            </font></div></td>
                        <td height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(cep2)%>
                          <input name="cep2" type="hidden" class="textInput" id="cep" value="<%response.Write(cep2)%>" size="11" maxlength="9" onKeyup="formatar(this, '#####-###')">
                          </font></td>
                        <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Estado:</strong></font> 
                          </div></td>
                        <td height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(uf2)%>
                          </font> </td>
                      </tr>
                      <tr> 
                        <td height="10" class="tb_corpo"
> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cidade:</strong></font></div></td>
                        <td height="10" colspan="2" class="tb_corpo"
> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(municipio2)%>
                          </font></td>
                        <td height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td height="10" class="tb_corpo"
> <div align="right"></div>
                          <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Bairro:</strong></font> 
                          </div></td>
                        <td height="10" class="tb_corpo"
><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(bairro2)%>
                          </font></td>
                      </tr>
                      <tr> 
                        <td height="10" colspan="6" class="tb_corpo"
><div align="left"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Telefones 
                            deste endere&ccedil;o:</strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            <%response.Write(telefone2)%>
                            <input name="telefones3" type="hidden" class="textInput" id="telefones2" value="<%response.Write(telefone2)%>" size="75" maxlength="50">
                            </font></div></td>
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
><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Filia&ccedil;&atilde;o</strong></font></td>
                </tr>
                <tr> 
                  <td colspan="2"><table width="100%" border="0" cellspacing="0">
                      <tr> 
                        <td width="4%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Pai:</strong></font></div></td>
                        <td width="46%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(pai)%>
                          </font></td>
                        <td width="8%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Falecido?</strong></font></div></td>
                        <td width="7%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(pai_fal)%>
                          </font></td>
                        <td width="14%" height="10"> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><font color="#CC9900">Situa&ccedil;&atilde;o 
                            dos Pais</font></strong>:</font></div></td>
                        <td width="21%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%
				
				
				response.Write(estado_civil)
				
				
				%>
                          </font></td>
                      </tr>
                      <tr> 
                        <td width="4%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>M&atilde;e:</strong></font></div></td>
                        <td height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(mae)%>
                          </font></td>
                        <td height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Falecida?</strong></font></div></td>
                        <td height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(mae_fal)%>
                          </font></td>
                        <td height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                        <td height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td bgcolor="#FFFFFF">&nbsp;</td>
                  <td bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="2" class="tb_tit"
><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Respons&aacute;veis</strong></font></td>
                </tr>
                <tr> 
                  <td colspan="2"><table width="100%" border="0" cellspacing="0">
                      <tr class="tb_corpo"
> 
                        <td width="10%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Pedag&oacute;gico:</strong></font></div></td>
                        <td width="40%" height="10"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%
		Set RSCONTP = Server.CreateObject("ADODB.Recordset")
		SQLP = "SELECT * FROM TB_Contatos WHERE TP_Contato ='"&resp_ped&"' And CO_Matricula ="& cod
		RSCONTP.Open SQLP, CONCONT
		
		resp_ped = RSCONTP("NO_Contato")
				  
				  response.Write(resp_ped)%>
                          </font></td>
                        <td width="10%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Financeiro:</strong></font></div></td>
                        <td width="40%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%
		Set RSCONTF = Server.CreateObject("ADODB.Recordset")
		SQLF = "SELECT * FROM TB_Contatos WHERE TP_Contato ='"&resp_fin&"' And CO_Matricula ="& cod
		RSCONTF.Open SQLF, CONCONT
		
		resp_fin = RSCONTF("NO_Contato")
				  
				  
				  response.Write(resp_fin)%>
                          </font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="2" class="tb_tit"
><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Dados 
                    Detalhados dos Familiares</strong></font></td>
                </tr>
                <tr> 
                  <td colspan="2"><table width="100%" border="0" cellspacing="0">
                      <%

		Set RSCONTPR = Server.CreateObject("ADODB.Recordset")
		SQLCONTPR = "SELECT * FROM TB_Tipo_Contatos Order by NU_Prioridade_Combo"
		RSCONTPR.Open SQLCONTPR, CON0
		
while not RSCONTPR.EOF
tp_resp = RSCONTPR("TP_Contato")
no_tp_resp = RSCONTPR("TX_Descricao")

		Set RSCONT = Server.CreateObject("ADODB.Recordset")
		SQLCONT = "SELECT * FROM TB_Contatos WHERE CO_Matricula ="& cod &" And TP_Contato = '"&tp_resp&"'"
		RSCONT.Open SQLCONT, CONCONT

if RSCONT.EOF then
RSCONTPR.MOVENEXT
else
resp= RSCONT("NO_Contato")

%>
                      <tr class="tb_corpo"
> 
                        <td width="20%" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                            <%response.Write(no_tp_resp)%>
                            :</strong></font></div></td>
                        <td width="80%" height="10"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <a href="contatos.asp?or=01&cod=<%=cod%>&tp=<%=tp_resp%>&vd=<%=vindo%>&o=<%=obr%>" class="ativos"> 
                          <%response.Write(resp)%>
                          </a> </font> <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            </font></div></td>
                      </tr>
                      <%RSCONTPR.MOVENEXT
end if
WEND
%>
                    </table></td>
                </tr>
                <tr> 
                  <td bgcolor="#FFFFFF">&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="2" class="tb_tit"
><font color="#FF6600" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Dados 
                    Escolares </strong></font></td>
                </tr>
                <tr> 
                  <td colspan="2"><table width="100%" border="0" cellspacing="0">
                      <tr class="tb_corpo"
> 
                        <td height="10" colspan="3"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nome 
                            do Col&eacute;gio de Origem:</strong></font></div></td>
                        <td height="10" colspan="6"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(col_origem)%>
                          </font></td>
                      </tr>
                      <tr class="tb_corpo"
> 
                        <td height="10" colspan="2"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Etapa 
                            cursada: </strong></font></div></td>
                        <td height="10" colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(cursada)%>
                          </font></td>
                        <td width="113" height="10"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Local:</strong></font></div></td>
                        <td height="10" colspan="4"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(cid_cursada&"/"&uf_cursada)%>
                          </font></td>
                      </tr>
                      <tr class="tb_corpo"
> 
                        <td height="10" colspan="2"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
                            do Cadastro: </strong></font></div></td>
                        <td height="10" colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(cadastro)%>
                          </font></td>
                        <td height="10" colspan="3"> <div align="right"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Data 
                            de Entrada na Escola: </strong></font></div></td>
                        <td height="10" colspan="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                          <%response.Write(entrada)%>
                          </font></td>
                      </tr>
                      <tr class="tb_corpo"
> 
                        <td width="33" height="10"> <div align="center"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                            <%
call GeraNomes("PORT",unidade,curso,etapa,CON0)
no_unidades = session("no_unidades")
no_grau = session("no_grau")
no_serie = session("no_serie")
%>
                            Ano</strong></font></div></td>
                        <td width="81" height="10"> <div align="center"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Matr&iacute;cula</strong></font></div></td>
                        <td width="75" height="10"> <div align="center"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Cancelamento</strong></font></div></td>
                        <td width="86" height="10"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            <strong><font color="#CC9900">Situa&ccedil;&atilde;o</font></strong></font></div></td>
                        <td width="113" height="10"> <div align="center"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Unidade</strong></font></div></td>
                        <td width="133" height="10"> <div align="center"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Curso</strong></font></div></td>
                        <td width="85" height="10"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"> 
                            <strong><font color="#CC9900" size="1"><strong>Etapa</strong></font></strong></font></div></td>
                        <td width="90" height="10"> <div align="center"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Turma</strong></font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            </font></div></td>
                        <td width="54" height="10"> <div align="center"><font color="#CC9900" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Chamada</strong></font></div></td>
                      </tr>
                      <tr class="tb_corpo"
> 
                        <td width="33" height="10"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            <%response.Write(ano_aluno)%>
                            </font> </div></td>
                        <td width="81" height="10"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            <%response.Write(rematricula)%>
                            </font></div></td>
                        <td width="75" height="10"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            <%response.Write(encerramento)%>
                            </font></div></td>
                        <td width="86" height="10"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            <%
					
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
		RSCONTST.Open SQLCONTST, CON0
							
				no_situacao = RSCONTST("TX_Descricao_Situacao")	
					response.Write(no_situacao)%>
                            </font></div></td>
                        <td width="113" height="10"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            <%response.Write(no_unidades)%>
                            </font></div></td>
                        <td width="133" height="10"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            <%response.Write(no_grau)%>
                            </font></div></td>
                        <td width="85" height="10"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            <%response.Write(no_serie)%>
                            </font></div></td>
                        <td width="90" height="10"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            <%response.Write(turma)%>
                            </font></div></td>
                        <td width="54" height="10"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            <%response.Write(cham)%>
                            </font></div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td bgcolor="#FFFFFF">&nbsp;</td>
                  <td>&nbsp;</td>
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
              </table></td>
          </tr>
        </table>
        <table width="1000" border="0" cellspacing="0">
          <tr> 
            <td width="219">&nbsp;</td>
            <td width="770" class="tb_voltar"
><font color="#669999" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="../alunos.asp" class="voltar1">&lt; 
              Voltar para o menu Alunos</a></strong></font></td>
          </tr>
        </table>
        
      </div></td>
  </tr>
  <tr>
    <td><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

</body>
<script type="text/javascript">
<!--
  initInputHighlightScript();
//-->
</script>
<%call GravaLog (3,18,codigo,"0")%>
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