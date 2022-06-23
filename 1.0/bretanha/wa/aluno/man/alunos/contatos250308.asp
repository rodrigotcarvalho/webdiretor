<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->



<!--#include file="../../../../inc/caminhos.asp"-->
<% 
opt= request.QueryString("opt")
tp_r= request.QueryString("tp")
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
ori = request.QueryString("or")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
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

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")		

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod
		RS.Open SQL, CON1
		
		Set RSCONTPR = Server.CreateObject("ADODB.Recordset")
		SQLCONTPR = "SELECT * FROM TB_Tipo_Contatos WHERE TP_Contato = '"&tp_r&"'"
		RSCONTPR.Open SQLCONTPR, CON0
		

		
codigo = RS("CO_Matricula")
no_tp_resp = RSCONTPR("TX_Descricao")




sexo = RS("IN_Sexo")



		Set RSCONTA = Server.CreateObject("ADODB.Recordset")
		SQLA = "SELECT * FROM TB_Contatos WHERE TP_Contato ='"&tp_r&"' And CO_Matricula ="& cod
		RSCONTA.Open SQLA, CONCONT



if RSCONTA.EOF then
nascimento="1/1/1900"
else
nascimento = RSCONTA("DA_Nascimento_Contato")
end if
if isnull(nascimento) then
nascimento="1/1/1900"
else
nascimento=nascimento
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
nome_prof = RSCONTA("NO_Contato")
apelido= RS("NO_Apelido")
desteridade= RS("IN_Desteridade")
nacionalidade= RS("CO_Nacionalidade")

cpf= RSCONTA("CO_CPF_PFisica")
rg= RSCONTA("CO_RG_PFisica")
emitido= RSCONTA("CO_OERG_PFisica")
emissao = RSCONTA("CO_DERG_PFisica")
profissao = RSCONTA("CO_Ocupacao")
empresa = RSCONTA("NO_Empresa")
rua=RSCONTA("NO_Logradouro_Res")
numero = RSCONTA("NU_Logradouro_Res")
complemento = RSCONTA("TX_Complemento_Logradouro_Res")
bairro= RSCONTA("CO_Bairro_Res")
municipio= RSCONTA("CO_Municipio_Res")
uf= RSCONTA("SG_UF_Res")
cep = RSCONTA("CO_CEP_Res")
telefone = RSCONTA("NU_Telefones_Res")
rua2=RSCONTA("NO_Logradouro_Com")
numero2 = RSCONTA("NU_Logradouro_Com")
complemento2 = RSCONTA("TX_Complemento_Logradouro_Com")
bairro2= RSCONTA("CO_Bairro_Com")
municipio2= RSCONTA("CO_Municipio_Com")
uf2= RSCONTA("SG_UF_Com")
cep2 = RSCONTA("CO_CEP_Com")
telefone2 = RSCONTA("NU_Telefones_Com")
mail= RSCONTA("TX_EMail")
tel_cont = RSCONTA("NU_Telefones")

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Matriculas WHERE CO_Matricula ="& cod
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

'if isnull(religiao) then
'else

'		Set RS0 = Server.CreateObject("ADODB.Recordset")
'		SQL0 = "SELECT * FROM TB_Religiao WHERE CO_Religiao ="& religiao
'		RS0.Open SQL0, CON0

'religiao = RS0("TX_Descricao_Religiao")
'end if
'if isnull(raca) then
'else
'		Set RS1 = Server.CreateObject("ADODB.Recordset")
'		SQL1 = "SELECT * FROM TB_Raca WHERE CO_Raca ="& raca
'		RS1.Open SQL1, CON0
'
'raca = RS1("TX_Descricao_Raca")
'end if
'if isnull(pais) then
'else
'		Set RS2 = Server.CreateObject("ADODB.Recordset")
'		SQL2 = "SELECT * FROM TB_Paises WHERE CO_Pais ="& pais
'		RS2.Open SQL2, CON0

'pais = RS2("NO_Pais")
'end if
'if isnull(nacionalidade) then
'else
'		Set RS3 = Server.CreateObject("ADODB.Recordset")
'		SQL3 = "SELECT * FROM TB_Nacionalidades WHERE CO_Nacionalidade ="& nacionalidade
'		RS3.Open SQL3, CON0

'nacionalidade = RS3("TX_Nacionalidade")
'end if

'if isnull(uf_natural) then
'else
'		Set RS6 = Server.CreateObject("ADODB.Recordset")
'		SQL6 = "SELECT * FROM TB_Municipios WHERE SG_UF ='"& uf_natural&"' AND CO_Municipio = "&natural
'		RS6.Open SQL6, CON0

'natural= RS6("NO_Municipio")
'end if
'if isnull(uf_natural) then
'else
'		Set RS8 = Server.CreateObject("ADODB.Recordset")
'		SQL8 = "SELECT * FROM TB_UF WHERE SG_UF ='"& uf_natural&"'" 
'		RS8.Open SQL8, CON0

'uf_natural= RS8("NO_UF")
'end if
'if isnull(ocupacao) then
'else
'		Set RS9 = Server.CreateObject("ADODB.Recordset")
'		SQL9 = "SELECT * FROM TB_Ocupacoes WHERE CO_Ocupacao ="& ocupacao
'		RS9.Open SQL9, CON0

'ocupacao= RS9("NO_Ocupacao")
'end if
'if isnull(co_estado_civil) then
'else
'response.Write("SQL10 = SELECT * FROM TB_Estado_Civil WHERE CO_Estado_Civil ='"& co_estado_civil&"'")

'		Set RS10 = Server.CreateObject("ADODB.Recordset")
'		SQL10 = "SELECT * FROM TB_Estado_Civil WHERE CO_Estado_Civil ='"& co_estado_civil&"'"
'		RS10.Open SQL10, CON0

'estado_civil= RS10("TX_Estado_Civil")
'end if
'if isnull(cid_cursada) then
'else
'		Set RS11 = Server.CreateObject("ADODB.Recordset")
'		SQL11 = "SELECT * FROM TB_Municipios WHERE SG_UF ='"& uf_cursada &"' AND CO_Municipio = "&cid_cursada
'		RS11.Open SQL11, CON0

'cid_cursada= RS11("NO_Municipio")
'end if
if isnull(uf)or isnull(bairro)or isnull(municipio) then
else
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Bairros WHERE CO_Bairro ="& bairro &"AND SG_UF ='"& uf&"' AND CO_Municipio = "&municipio
		RS4.Open SQL4, CON0
if RS4.EOF then
bairro = ""
else
bairro= RS4("NO_Bairro")
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

		


Call LimpaVetor2

%>
<html>
<head>
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../../../js/mm_menu.js"></script>
<script type="text/javascript" src="../../../../js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_nbGroup(event, grpName) { //v6.0
  var i,img,nbArr,args=MM_nbGroup.arguments;
  if (event == "init" && args.length > 2) {
    if ((img = MM_findObj(args[2])) != null && !img.MM_init) {
      img.MM_init = true; img.MM_up = args[3]; img.MM_dn = img.src;
      if ((nbArr = document[grpName]) == null) nbArr = document[grpName] = new Array();
      nbArr[nbArr.length] = img;
      for (i=4; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
        if (!img.MM_up) img.MM_up = img.src;
        img.src = img.MM_dn = args[i+1];
        nbArr[nbArr.length] = img;
    } }
  } else if (event == "over") {
    document.MM_nbOver = nbArr = new Array();
    for (i=1; i < args.length-1; i+=3) if ((img = MM_findObj(args[i])) != null) {
      if (!img.MM_up) img.MM_up = img.src;
      img.src = (img.MM_dn && args[i+2]) ? args[i+2] : ((args[i+1])? args[i+1] : img.MM_up);
      nbArr[nbArr.length] = img;
    }
  } else if (event == "out" ) {
    for (i=0; i < document.MM_nbOver.length; i++) {
      img = document.MM_nbOver[i]; img.src = (img.MM_dn) ? img.MM_dn : img.MM_up; }
  } else if (event == "down") {
    nbArr = document[grpName];
    if (nbArr)
      for (i=0; i < nbArr.length; i++) { img=nbArr[i]; img.src = img.MM_up; img.MM_dn = 0; }
    document[grpName] = nbArr = new Array();
    for (i=2; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
      if (!img.MM_up) img.MM_up = img.src;
      img.src = img.MM_dn = (args[i+1])? args[i+1] : img.MM_up;
      nbArr[nbArr.length] = img;
  } }
}

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

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
<script language="JavaScript" type="text/JavaScript">
<!--
function checksubmit()
{
  if (document.busca.busca1.value != "" && document.busca.busca2.value != "")
  {    alert("Por favor digite SOMENTE uma opção de busca!")
    document.busca.busca1.focus()
    return false
  }
    if (document.busca.busca1.value == "" && document.busca.busca2.value == "")
  {    alert("Por favor digite uma opção de busca!")
    document.busca.busca1.focus()
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
<% if opt="listall" or opt="list" then%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%else %>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_callJS('document.busca.busca1.focus()')">
<%end if %>
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

<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
            
    <td height="10" colspan="5" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
    </td>
          </tr>
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,306,0,0) %>
    </td>
			  </tr>			  
        <form action="cadastro.asp?opt=list&or=01" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
          <tr>
      <td height="412" valign="top"> 
        <table width="1000" height="400" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo"
>
          <tr> 
                  
            <td width="653" height="10" class="tb_tit"
>Dados Pessoais</td>
                  <td width="113" class="tb_tit"
> </td>
                </tr>
                <tr> 
                  <td height="10" colspan="2"> <table width="100%" border="0" cellspacing="0">
                      <tr> 
                        <td height="10"> <div align="right"><font class="form_dado_texto"> 
              Nome: </font></div></td>
                        <td height="10" colspan="2"><font class="form_corpo">  
                          <%response.Write(nome_prof)%>
                          </font> <div align="right"><font class="form_dado_texto"> </font></div></td>
                        <td width="13%" height="10"><div align="right"><font class="form_corpo"> &nbsp; 
              </font><font class="form_dado_texto"> Rela&ccedil;&atilde;o:</font><font class="form_corpo">  
              </font></div></td>
                        <td width="19%"><font class="form_corpo">  
                          <%response.Write(no_tp_resp)%>
                          </font></td>
                      </tr>
                      <tr> 
                        <td width="19%" height="10"> <div align="right"><font class="form_corpo"> &nbsp;</font><font class="form_dado_texto"> 
              Data de Nascimento: </font></div></td>
                        <td width="27%" height="10"><font class="form_corpo">  
                          <%response.Write(nasce)%>
                          </font> <font class="form_corpo">  
                          <input name="nasce" type="hidden" class="textInput" id="nasce4" value="<%response.Write(nasce)%>" size="12" maxlength="10">
            &nbsp;-&nbsp; <font class="form_corpo">  
            <%
					call aniversario(ano_a,mes_a,dia_a) %>
            </font> </font></td>
                        <td width="22%" height="10"> </td>
                        <td height="10" colspan="2"> </td>
                      </tr>
                      <tr> 
                        <td height="10"> <div align="right"><font class="form_dado_texto"> 
              Ocupa&ccedil;&atilde;o:</font></div></td>
                        <td height="10"><font class="form_corpo">  
                          <% response.Write(profissao)%>
                          </font></td>
                        <td height="10"> <div align="right"><font class="form_dado_texto"> 
              Empresa onde trabalha:</font></div></td>
                        <td height="10" colspan="2"><font class="form_corpo">  
                          <%response.Write(empresa)%>
                          </font></td>
                      </tr>
                      <tr> 
                        <td height="10"> <div align="right"><font class="form_dado_texto"> 
              Endere&ccedil;o Eletr&ocirc;nico:</font></div></td>
                        <td height="10" colspan="4"><font class="form_corpo">  
                          <%response.Write(mail)%>
                          </font><font class="form_corpo"> &nbsp; 
                          </font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="10" colspan="2"> <table width="100%" border="0" cellspacing="0">
                      <tr> 
                        <td height="10" width="13%"> <div align="right"><font class="form_dado_texto"> 
              CPF:</font></div></td>
                        <td height="10" width="18%"><font class="form_corpo">  
                          <%response.Write(cpf)%>
                          </font></td>
                        <td width="18%" height="10"> <div align="right"><font class="form_dado_texto"> 
              Identidade: </font></div></td>
                        <td width="30%" height="10"><font class="form_corpo">  
                          <%response.Write(rg)%>
                          &nbsp;</font></td>
                        <td width="12%" height="10"> <div align="right"><font class="form_dado_texto"> 
              Tipo:</font></div></td>
                        <td width="12%" height="10"> <div align="left"><font class="form_corpo">  
              <%response.Write(emitido)%>
              </font><font class="form_dado_texto"> </font></div></td>
                        <td width="12%" height="10"> <div align="right"><font class="form_dado_texto"> 
              Data:</font></div></td>
                        <td width="12%" height="10"><font class="form_corpo">  
                          <%response.Write(emissao)%>
                          </font></td>
                      </tr>
                      <tr> 
                        <td height="10" colspan="8"> <div align="left"><font class="form_dado_texto"> 
              Telefones de Contato:</font><font class="form_corpo">  
              <%response.Write(tel_cont)%>
              </font><font class="form_corpo"> &nbsp;</font></div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="10" bgcolor="#FFFFFF">&nbsp;</td>
                  <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                </tr>
                <tr> 
                  
            <td height="10" colspan="2" class="tb_tit"
>Endere&ccedil;o Residencial</td>
                </tr>
                <tr> 
                  <td height="10" colspan="2"> <table width="100%" height="10" border="0" cellspacing="0">
                      <tr class="tb_corpo"
> 
                        <td width="11%" height="10"> <div align="right"><font class="form_dado_texto"> 
              Logradouro:</font></div></td>
                        <td width="60%" height="10"><font class="form_corpo">  
                          <%response.Write(rua)%>
                          <input name="rua" type="hidden" class="textInput" id="rua4" value="<%response.Write(rua)%>" size="75" maxlength="50">
                          </font></td>
                        <td width="13%" height="10"> 
<div align="right"><font class="form_dado_texto"> N&uacute;mero:</font><font class="form_corpo">  
              </font></div></td>
                        <td width="16%" height="10"><font class="form_corpo">  
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
                        <td width="12%" height="10"> <div align="right"><font class="form_dado_texto"> 
              Complemento:</font></div></td>
                        <td width="36%" height="10"> <font class="form_corpo">  
                          <%response.Write(complemento)%>
                          </font> <font class="form_corpo">  
                          <input name="complemento" type="hidden" class="textInput" id="complemento4" value="<%response.Write(complemento)%>" size="45" maxlength="30">
                          </font></td>
                        <td width="6%" height="10"> <div align="right"><font class="form_corpo"> &nbsp; 
              </font><font class="form_dado_texto"> CEP:</font><font class="form_corpo"> &nbsp; 
              </font></div></td>
                        <td width="13%" height="10"><font class="form_corpo">  
                          <%response.Write(cep)%>
                          <input name="cep" type="hidden" class="textInput" id="cep7" value="<%response.Write(cep)%>" size="11" maxlength="9" onKeyup="formatar(this, '#####-###')">
                          </font></td>
                        <td width="7%" height="10"> <div align="right"><font class="form_dado_texto"> 
              Estado:</font> </div></td>
                        <td width="26%" height="10"><font class="form_corpo">  
                          <%response.Write(uf)%>
                          </font> </td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="10" colspan="2"> <table width="100%" height="10" border="0" cellspacing="0">
                      <tr class="tb_corpo"
> 
                        <td width="16%" height="10"> <div align="right"><font class="form_dado_texto"> 
              Cidade:</font></div></td>
                        <td width="29%" height="10"> <font class="form_corpo">  
                          <%response.Write(municipio)%>
                          </font></td>
                        <td width="11%" height="10"> <div align="right"><font class="form_dado_texto"> 
              Bairro:</font> </div></td>
                        <td width="44%" height="10"><font class="form_corpo">  
                          <%response.Write(bairro)%>
                          </font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="10" colspan="2"> <table width="100%" height="10" border="0" cellspacing="0">
                      <tr class="tb_corpo"
> 
                        <td width="20%" height="10"> <div align="right"><font class="form_dado_texto"> 
              Telefones deste endere&ccedil;o:</font></div></td>
                        <td width="80%" height="10"><font class="form_corpo">  
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
                  
            <td height="10" colspan="2" class="tb_tit"
>Endere&ccedil;o Comercial </td>
                </tr>
                <tr> 
                  <td height="10" colspan="2"> <table width="100%" border="0" cellspacing="0">
                      <tr class="tb_corpo"
> 
                        <td width="11%" height="10"> <div align="right"><font class="form_dado_texto"> 
              Logradouro:</font></div></td>
                        <td width="60%" height="10"><font class="form_corpo">  
                          <%response.Write(rua2)%>
                          <input name="rua" type="hidden" class="textInput" id="rua4" value="<%response.Write(rua2)%>" size="75" maxlength="50">
                          </font></td>
                        <td width="13%" height="10"> 
<div align="right"><font class="form_dado_texto"> N&uacute;mero:</font><font class="form_corpo">  
              </font></div></td>
                        <td width="16%" height="10"><font class="form_corpo">  
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
                        <td width="12%" height="10"> <div align="right"><font class="form_dado_texto"> 
              Complemento:</font></div></td>
                        <td width="36%" height="10"> <font class="form_corpo">  
                          <%response.Write(complemento2)%>
                          </font> <font class="form_corpo">  
                          <input name="complemento" type="hidden" class="textInput" id="complemento4" value="<%response.Write(complemento2)%>" size="45" maxlength="30">
                          </font></td>
                        <td width="6%" height="10"> <div align="right"><font class="form_corpo"> &nbsp; 
              </font><font class="form_dado_texto"> CEP:</font><font class="form_corpo"> &nbsp; 
              </font></div></td>
                        <td width="13%" height="10"><font class="form_corpo">  
                          <%response.Write(cep2)%>
                          <input name="cep" type="hidden" class="textInput" id="cep7" value="<%response.Write(cep2)%>" size="11" maxlength="9" onKeyup="formatar(this, '#####-###')">
                          </font></td>
                        <td width="7%" height="10"> 
<div align="right"><font class="form_dado_texto"> Estado:</font> </div></td>
                        <td width="26%" height="10"><font class="form_corpo">  
                          <%response.Write(uf2)%>
                          </font> </td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="10" colspan="2"> <table width="100%" border="0" cellspacing="0">
                      <tr class="tb_corpo"
> 
                        <td width="16%" height="10"> <div align="right"><font class="form_dado_texto"> 
              Cidade:</font></div></td>
                        <td width="29%" height="10"> <font class="form_corpo">  
                          <%response.Write(municipio2)%>
                          </font></td>
                        <td width="11%" height="10"> 
<div align="right"><font class="form_dado_texto"> Bairro:</font> </div></td>
                        <td width="44%" height="10"><font class="form_corpo">  
                          <%response.Write(bairro2)%>
                          </font></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td height="10" colspan="2"> <table width="100%" border="0" cellspacing="0">
                      <tr class="tb_corpo"
> 
                        <td width="20%" height="10"> <div align="right"><font class="form_dado_texto"> 
              Telefones deste endere&ccedil;o:</font></div></td>
                        <td width="80%" height="10"><font class="form_corpo">  
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
                        
          <td bgcolor="#FFFFFF">&nbsp; </td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  
            <td height="10" colspan="2" class="tb_tit"
>&nbsp;</td>
                </tr>
              </table></td></tr>
</form>
          <tr> 
            
    <td height="190" valign="top"> 
      <table width="1000" border="0" cellspacing="0">
          <tr> 
            
          <td width="219">&nbsp;</td>
            
          <td width="770" class="tb_voltar"
><font color="#669999" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="altera.asp?opt=vt&cod=<%=cod%>" class="voltar1">&lt; 
            Voltar para Consultar Aluno</a></strong></font></td>
          </tr>
        </table>
    </td>
  </tr>
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

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
pasta=arPath(seleciona1)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("../../../../inc/erro.asp")
end if
%>