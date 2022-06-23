<%On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->



<!--#include file="../../../../inc/caminhos.asp"-->
<% 
opt= request.QueryString("opt")
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo")
ano_letivo_real = ano_letivo
sistema_local=session("sistema_local")
ori = request.QueryString("ori")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
cod= request.QueryString("cod_cons")	

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
		
		
codigo = RS("CO_Matricula")
nome_prof = RS("NO_Aluno")

sexo = RS("IN_Sexo")

		Set RSCONTA = Server.CreateObject("ADODB.Recordset")
		SQLA = "SELECT * FROM TB_Contatos WHERE TP_Contato ='ALUNO' And CO_Matricula ="& cod
		RSCONTA.Open SQLA, CONCONT

if RSCONTA.EOF then
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

if religiao="" or isnull(religiao) or religiao=0 then
else
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Religiao WHERE CO_Religiao ="& religiao
		RS0.Open SQL0, CON0

if RS0.EOF then
else
religiao = RS0("TX_Descricao_Religiao")
end if
end if

if raca="" or isnull(raca) or raca=0 then
else
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Raca WHERE CO_Raca ="& raca
		RS1.Open SQL1, CON0
		
if RS1.EOF then
else
raca = RS1("TX_Descricao_Raca")
end if
end if

if pais="" or isnull(pais) or pais=0 then
else
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Paises WHERE CO_Pais ="& pais
		RS2.Open SQL2, CON0
		
if RS2.EOF then
else
pais = RS2("NO_Pais")
end if
end if


if nacionalidade="" or isnull(nacionalidade) or nacionalidade=0 then
else
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Nacionalidades WHERE CO_Nacionalidade ="& nacionalidade
		RS3.Open SQL3, CON0

if RS3.EOF then
else
nacionalidade = RS3("TX_Nacionalidade")
end if
end if


if uf_natural="" or isnull(uf_natural) or natural="" or isnull(natural) or natural=0 then
else
		Set RS6 = Server.CreateObject("ADODB.Recordset")
		SQL6 = "SELECT * FROM TB_Municipios WHERE SG_UF ='"& uf_natural&"' AND CO_Municipio = "&natural
		RS6.Open SQL6, CON0

if RS6.EOF then
else
natural= RS6("NO_Municipio")
end if
end if

if uf_natural="" or isnull(uf_natural) then
else
		Set RS8 = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT * FROM TB_UF WHERE SG_UF ='"& uf_natural&"'" 
		RS8.Open SQL8, CON0

if RS8.EOF then
else
uf_natural= RS8("NO_UF")
end if
end if


if ocupacao="" or isnull(ocupacao) or ocupacao=0 then
else
		Set RS9 = Server.CreateObject("ADODB.Recordset")
		SQL9 = "SELECT * FROM TB_Ocupacoes WHERE CO_Ocupacao ="& ocupacao
		RS9.Open SQL9, CON0

if RS9.EOF then
else
ocupacao= RS9("NO_Ocupacao")
end if
end if

if co_estado_civil="" or isnull(co_estado_civil) then
else
		Set RS10 = Server.CreateObject("ADODB.Recordset")
		SQL10 = "SELECT * FROM TB_Estado_Civil WHERE CO_Estado_Civil ='"& co_estado_civil&"'"
		RS10.Open SQL10, CON0

if RS10.EOF then
else
estado_civil= RS10("TX_Estado_Civil")
end if
end if

if uf_cursada="" or isnull(uf_cursada) or cid_cursada="" or isnull(cid_cursada) or cid_cursada=0 then
else
		Set RS11 = Server.CreateObject("ADODB.Recordset")
		SQL11 = "SELECT * FROM TB_Municipios WHERE SG_UF ='"& uf_cursada &"' AND CO_Municipio = "&cid_cursada
		RS11.Open SQL11, CON0
		
if RS11.EOF then
else
cid_cursada= RS11("NO_Municipio")
end if
end if

if uf="" or isnull(uf) or municipio="" or isnull(municipio) or municipio=0 then
else

		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Bairros WHERE CO_Bairro ="& bairro &"AND SG_UF ='"& uf&"' AND CO_Municipio = "&municipio
		RS4.Open SQL4, CON0
if RS4.EOF then
bairro = ""
else
bairro= RS4("NO_Bairro")
end if
end if

if isnull(uf2)or isnull(municipio2) or municipio2=0 then

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
function submitano()  
{
   var f=document.forms[0]; 
      f.submit(); 
}
function submitsistema()  
{
   var f=document.forms[1]; 
      f.submit(); 
}
function submitrapido()  
{
   var f=document.forms[2]; 
      f.submit(); 
}  
function submitfuncao()  
{
   var f=document.forms[3]; 
      f.submit(); 
} 

function centraliza(w,h){
//o 120 e o 16 se referem ao tamanho di cabeçalho do navegador e a barra de rolagem respectivamente
    x = parseInt((screen.width - w - 16)/2);
    y = parseInt((screen.height - h - 120)/2);
   //alert(x + '\n' + y);
    document.getElementById('alinha').style.left = x;
    document.getElementById('alinha').style.top = y;
	
//	alert('w '+x +' h '+ y)
}

//-->
</script>
</head>
<% if opt="listall" or opt="list" then%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%else %>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%end if %>
<%call cabecalho(nivel)
%>
<div id="fundo" style="position:absolute; left:0px; top:0px; width:100%; height:1200; z-index:1; background-color: #000000; layer-background-color: #000000; border: 1px none #000000; visibility: hidden;" class="transparente"></div>
<div id="alinha" style="position:absolute; width:500px; z-index: 2; height: 536px; visibility: hidden;"> 
  <table width="500" border="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr> 
      <td width="478" height="16"> 
        <div align="right"> <span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')">fechar</a></font></span></div></td>
      <td width="20"><div align="right"><span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide')"><img src="../../../../img/fecha.gif" width="20" height="16" border="0"></a></font></span></div></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center" ><img src="../../../../img/fotos/aluno/<% =codigo %>.jpg" height="500"></div></td>
    </tr>
    <tr>
      <td height="20" colspan="2">
<div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
          <%response.Write(nome_prof)%>
          </font></div></td>
    </tr>
  </table>
</div>

<table width="1002" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
            
    <td width="1000" height="10" colspan="5" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
    </td>
  </tr>
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,302,0,0) %>
    </td>
  </tr>			  
        <form action="cadastro.asp?opt=list&or=01" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
          <tr><td><table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
        <tr> 
                  
          <td width="841" class="tb_tit"
>Dados Pessoais</td>
                  <td width="151" class="tb_tit"
> </td>
                  <td width="2" class="tb_tit"
></td>
        </tr>
                <tr> 
                  <td height="10"> <font class="form_corpo">
                    <input name="tp" type="hidden" id="tp2" value="P">
                  </font>
                    <table width="100%" border="0" cellspacing="0">
              
              <tr>
                <td width="17%" height="10"><font class="form_dado_texto">Matr&iacute;cula</font></td>
                <td width="2%"><div align="center">:</div></td>
                <td width="26%" height="10"><font class="form_corpo">
                  <input name="cod" type="hidden" value="<%=codigo%>">
                  <%response.Write(codigo)%>
                  <input name="acesso" type="hidden" id="acesso" value="2">
                  <input name="co_usr_prof" type="hidden" id="co_usr_prof" value="<% =co_usr_prof%>">
                </font></td>
                <td height="10"><font class="form_dado_texto">Nome:</font></td>
                <td><div align="center">:</div></td>
                <td height="10"><font class="form_corpo">
                  <%response.Write(nome_prof)%>
                  <input name="nome" type="hidden" class="textInput" id="nome"  value="<%response.Write(nome_prof)%>" size="75" maxlength="50">
                </font></td>
              </tr>
              
              <tr> 
                <td height="10"> <div align="left"><font class="form_dado_texto"> 
                  Apelido</font></div></td>
                
                <td><div align="center">:</div></td>
                <td height="10"><font class="form_corpo"> 
            <%response.Write(apelido)%>
            &nbsp; </font></td>
                <td width="17%" height="10"> <div align="left"><font class="form_dado_texto"> Data de Nascimento</font></div></td>
                
                <td width="2%"><div align="center">:</div></td>
                <td height="10"><font class="form_corpo">
                  <%response.Write(nasce)%>
                  <input name="nasce2" type="hidden" class="textInput" id="nasce" value="<%response.Write(nasce)%>" size="12" maxlength="10"> 
                  - 
                  <%
					call aniversario(ano_a,mes_a,dia_a) %>
                </font></td>
              </tr>
              <tr> 
                <td height="10"> <div align="left"><font class="form_dado_texto"> 
                  Sexo</font></div></td>
                
                <td><div align="center">:</div></td>
                <td height="10"><font class="form_corpo"> 
            <%if sexo = "M" then
					  response.Write("Masculino")
						else
					  response.Write("Feminino")
                    End IF%>
            </font></td>
                <td height="10"> <div align="left"><font class="form_dado_texto"> Pa&iacute;s de Origem</font></div></td>
                
                <td><div align="center">:</div></td>
                <td height="10"><font class="form_corpo"> 
            <%response.Write(pais)%>
            </font></td>
              </tr>
              <tr> 
                <td height="10"> <div align="left"><font class="form_dado_texto"> 
                  Nacionalidade</font></div></td>
                
                <td><div align="center">:</div></td>
                <td height="10"> <font class="form_corpo"> 
            <%response.Write(nacionalidade)%>
            </font> </td>
                <td height="10"> 
                  <div align="left"><font class="form_dado_texto"> Natural do Estado</font></div></td>
                
                <td><div align="center">:</div></td>
                <td height="10"><font class="form_corpo"> 
            <%response.Write(uf_natural)%>
            </font></td>
              </tr>
              <tr> 
                <td height="10"> <div align="left"><font class="form_dado_texto"> 
                  Natural da cidade</font></div></td>
                
                <td><div align="center">:</div></td>
                <td height="10"> <font class="form_corpo"> 
            <%response.Write(natural)%>
            </font> </td>
                <td height="10">  <div align="left"><font class="form_dado_texto"> Religi&atilde;o</font></div></td>
                
                <td><div align="center">:</div></td>
                <td height="10"><font class="form_corpo"> 
            <%response.Write(religiao)%>
            </font></td>
              </tr>
              <tr> 
                <td height="10"> <div align="left"><font class="form_dado_texto"> Cor / Ra&ccedil;a</font></div></td>
                
                <td><div align="center">:</div></td>
                <td height="10"><font class="form_corpo"> 
            <%response.Write(raca)%>
            </font> </td>
                <td height="10"> <font class="form_dado_texto">  Ocupa&ccedil;&atilde;o</font></td>
                <td><div align="center">:</div></td>
                <td width="36%" height="10"> <div align="left"><font class="form_corpo"> 
            <%response.Write(ocupacao)%>
              </font></div></td>
                </tr>
              <tr>
                <td height="10"><font class="form_dado_texto">Identidade</font></td>
                <td><div align="center">:</div></td>
                <td height="10"><font class="form_corpo">
                  <%response.Write(rg)%>
                </font></td>
                <td height="10"><font class="form_dado_texto">Tipo - Data de Emiss&atilde;o </font></td>
                <td><div align="center">:</div></td>
                <td height="10"><font class="form_corpo">
                  <%response.Write(emitido)%>
                - 
                <%response.Write(emissao)%>
                </font></td>
              </tr>
              <tr> 
                <td height="10"><div align="left"><font class="form_dado_texto"> 
                  CPF</font></div></td>
                
                <td><div align="center">:</div></td>
                <td height="10"> <font class="form_corpo"> 
            <%response.Write(cpf)%>
            </font></td>
                <td height="10"><div align="left"><font class="form_dado_texto"> 
                  Empresa onde trabalha</font></div></td>
                
                <td><div align="center">:</div></td>
                <td height="10"><font class="form_corpo"> 
            <%response.Write(empresa)%>
            </font></td>
              </tr>
              
              <tr>
                <td height="10"><font class="form_dado_texto">Endere&ccedil;o Eletr&ocirc;nico</font></td>
                <td><div align="center">:</div></td>
                <td height="10"><font class="form_corpo">
                  <%response.Write(mail)%>
                </font></td>
                <td height="10">&nbsp;</td>
                <td><div align="center"></div></td>
                <td height="10">&nbsp;</td>
                </tr>
              <tr>
                <td height="10"><font class="form_dado_texto">Login Orkut</font></td>
                <td><div align="center">:</div></td>
                <td height="10"><font class="form_corpo">
                  <%response.Write(orkut)%>
                </font></td>
                <td height="10"><font class="form_dado_texto">Login Messenger</font></td>
                <td><div align="center">:</div></td>
                <td height="10"><font class="form_corpo">
                  <%response.Write(msn)%>
                </font></td>
                </tr>
              <tr> 
                <td height="10"> <div align="left"><font class="form_dado_texto">Telefones de Contato</font></div></td>
                
                <td><div align="center">:</div></td>
                <td height="10"><div align="left"><font class="form_corpo"></font> <font class="form_corpo">
                  <%response.Write(tel_cont)%>
                </font></div></td>
                <td height="10"> <div align="left"><font class="form_dado_texto">Escrita</font></div></td>
                
                <td><div align="center">:</div></td>
                <td height="10"><font class="form_corpo">
                  <%response.Write(desteridade)%>
                </font></td>
                </tr>
            </table></td>
                  <td valign="top"><table height="110" border="3" align="right" cellspacing="0" bordercolor="#EEEEEE">
                    
                    
                    <tr>
                      <td><div align="center"><a href="#" onClick="centraliza(500,536);MM_showHideLayers('fundo','','show','alinha','','show')"><img src="../../../../img/fotos/aluno/<% =codigo %>.jpg" alt="" width="133" height="167" border="0"></a></div></td>
                    </tr>
                    <tr>
                      <td height="15" bgcolor="#EEEEEE"><div align="center"><a href="#" onClick="centraliza(500,536);MM_showHideLayers('fundo','','show','alinha','','show')"><img src="../../../../img/clique.gif" width="85" height="13" border="0"></a></div></td>
                    </tr>
                  </table></td>
                  <td valign="top">&nbsp;</td>
                </tr>
                
                <tr> 
                  
          <td colspan="3" class="tb_tit"
>Endere&ccedil;o Residencial</td>
                </tr>
                <tr> 
                  <td height="10" colspan="3"> <table width="100%" border="0" cellspacing="0">
              <tr> 
                <td width="14%" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Logradouro</font></div></td>
                <td width="2%" class="tb_corpo"
><div align="center">:</div></td>
                <td height="10" class="tb_corpo"
><font class="form_corpo"> 
            <%response.Write(rua)%>
                  <input name="rua" type="hidden" class="textInput" id="rua4" value="<%response.Write(rua)%>" size="75" maxlength="50">
                  </font></td>
                <td width="15%" height="10" class="tb_corpo"
> 
                  <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
                <td width="1%" class="tb_corpo"
><div align="center">:</div></td>
                <td width="15%" class="tb_corpo"
><font class="form_corpo">
                  <font class="form_corpo">
                  <%response.Write(numero)%>
                  </font>
                  <input name="numero" type="hidden" class="textInput" id="numero4" value="<%response.Write(numero)%>" size="11" maxlength="6">
                </font></td>
                <td width="15%" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
                <td width="1%" class="tb_corpo"
><div align="center">:</div></td>
                <td width="15%" height="10" class="tb_corpo"
><div align="left"><font class="form_corpo"> </font> <font class="form_corpo">
                  <%response.Write(complemento)%>
                </font></div></td>
              </tr>
              
              <tr>
                <td height="21" class="tb_corpo"
><font class="form_dado_texto">Bairro</font></td>
                <td class="tb_corpo"
><div align="center">:</div></td>
                <td height="21" class="tb_corpo"
><font class="form_corpo">
                  <%response.Write(bairro)%>
                </font></td>
                <td height="21" class="tb_corpo"
><font class="form_dado_texto">Cidade</font></td>
                <td class="tb_corpo"
><div align="center">:</div></td>
                <td class="tb_corpo"
><font class="form_corpo">
                  <%response.Write(municipio)%>
                </font></td>
                <td class="tb_corpo"
><font class="form_dado_texto">Estado</font></td>
                <td class="tb_corpo"
><div align="center">:</div></td>
                <td height="21" class="tb_corpo"
><font class="form_corpo">
                  <%response.Write(uf)%>
                </font></td>
              </tr>
              <tr>
                <td height="10" class="tb_corpo"
><font class="form_dado_texto">CEP</font></td>
                <td class="tb_corpo"
><div align="center">:</div></td>
                <td height="10" class="tb_corpo"
><font class="form_corpo">
                  <%response.Write(cep)%>
                  <input name="cep2" type="hidden" class="textInput" id="cep" value="<%response.Write(cep2)%>" size="11" maxlength="9" onKeyup="formatar(this, '#####-###')">
                </font></td>
                <td height="10" class="tb_corpo"
>&nbsp;</td>
                <td class="tb_corpo"
>&nbsp;</td>
                <td class="tb_corpo"
>&nbsp;</td>
                <td class="tb_corpo"
>&nbsp;</td>
                <td class="tb_corpo"
><div align="center"></div></td>
                <td height="10" class="tb_corpo"
>&nbsp;</td>
              </tr>
              <tr> 
                <td height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">Telefones deste endere&ccedil;o</font></div></td>
                <td class="tb_corpo"
><div align="center">:</div></td>
                <td height="10" class="tb_corpo"
><font class="form_corpo">
                  <%response.Write(telefone)%>
                  <input name="telefones2" type="hidden" class="textInput" id="telefones" value="<%response.Write(telefone)%>" size="75" maxlength="50">
                </font></td>
                <td height="10" class="tb_corpo"
> 
                  <div align="left"></div></td>
                <td class="tb_corpo"
><div align="center">:</div></td>
                <td class="tb_corpo"
>&nbsp;</td>
                <td class="tb_corpo"
>&nbsp;</td>
                <td class="tb_corpo"
><div align="center"></div></td>
                <td height="10" class="tb_corpo"
>&nbsp;</td>
              </tr>
              
              <tr> 
                <td height="10" colspan="9" class="tb_tit"
><div align="left">Endere&ccedil;o Comercial </div></td>
              </tr>
              <tr> 
                <td height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Logradouro</font></div></td>
                <td class="tb_corpo"
><div align="center">:</div></td>
                <td height="10" class="tb_corpo"
><font class="form_corpo"> 
            <%response.Write(rua2)%>
                  <input name="rua2" type="hidden" class="textInput" id="rua" value="<%response.Write(rua2)%>" size="75" maxlength="50">
                  </font></td>
                <td height="10" class="tb_corpo"
> 
                  <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
                <td class="tb_corpo"
><div align="center">:</div></td>
                <td class="tb_corpo"
><font class="form_corpo">
                  <%response.Write(numero2)%>
                  <font class="form_corpo">
                  <input name="numero2" type="hidden" class="textInput" id="numero" value="<%response.Write(numero2)%>" size="11" maxlength="6">
                  </font></font></td>
                <td class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
                <td class="tb_corpo"
><div align="center">:</div></td>
                <td height="10" class="tb_corpo"
><div align="left"><font class="form_corpo"> </font> <font class="form_corpo">
                  <%response.Write(complemento2)%>
                  </font><font class="form_dado_texto">
                    <input name="complemento2" type="hidden" class="textInput" id="complemento" value="<%response.Write(complemento2)%>" size="45" maxlength="30">
                  </font></div></td>
              </tr>
              <tr class="tb_corpo"
>
                <td height="26"><font class="form_dado_texto">Bairro</font></td>
                <td><div align="center">:</div></td>
                <td height="26"><font class="form_corpo">
                  <%response.Write(bairro2)%>
                </font></td>
                <td height="26"><font class="form_dado_texto">Cidade</font></td>
                <td><div align="center">:</div></td>
                <td><font class="form_corpo">
                  <%response.Write(municipio2)%>
                </font></td>
                <td><font class="form_dado_texto">Estado</font></td>
                <td><div align="center">:</div></td>
                <td height="26"><font class="form_corpo">
                  <%response.Write(uf2)%>
                </font></td>
              </tr>
              <tr class="tb_corpo"
>
                <td height="26"><font class="form_dado_texto">CEP</font></td>
                <td><div align="center">:</div></td>
                <td height="26"><font class="form_corpo">
                  <%response.Write(cep2)%>
                  <input name="cep22" type="hidden" class="textInput" id="cep2" value="<%response.Write(cep2)%>" size="11" maxlength="9" onKeyup="formatar(this, '#####-###')">
                </font></td>
                <td height="26">&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td><div align="center"></div></td>
                <td height="26">&nbsp;</td>
              </tr>
              <tr class="tb_corpo"
> 
                <td height="28"> <div align="left"><font class="form_dado_texto">Telefones deste endere&ccedil;o:</font></div></td>
                <td><div align="center">:</div></td>
                <td width="22%" height="28"><font class="form_corpo">
                  <%response.Write(telefone2)%>
                  <input name="telefones3" type="hidden" class="textInput" id="telefones2" value="<%response.Write(telefone2)%>" size="75" maxlength="50">
                </font></td>
                <td height="28"> <div align="left"></div></td>
                <td><div align="center"></div></td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
                <td><div align="center"></div></td>
                <td height="28">&nbsp;</td>
              </tr>
              
              
            </table></td>
                </tr>
                
                <tr> 
                  
          <td colspan="3" class="tb_tit"
>Filia&ccedil;&atilde;o</td>
                </tr>
                <tr> 
                  <td colspan="3"><table width="100%" border="0" cellspacing="0">
              <tr> 
                <td width="14%" height="26"> <div align="left"><font class="form_dado_texto"> 
                  Pai</font></div></td>
                <td width="2%"><div align="center">:</div></td>
                <td width="22%" height="26"><font class="form_corpo"> 
            <%response.Write(pai)%>
                  </font></td>
                <td width="15%" height="26"> <div align="left"><font class="form_dado_texto"> 
                  Falecido</font></div></td>
                <td width="1%"><div align="center"><font class="form_dado_texto">?</font></div></td>
                <td width="15%" height="26"><font class="form_corpo"> 
            <%response.Write(pai_fal)%>
                  </font></td>
                <td width="15%" height="26"> <div align="left"><font class="form_dado_texto"> Situa&ccedil;&atilde;o dos Pais</font></div></td>
                <td width="1%"><div align="center">:</div></td>
                <td width="15%" height="26"><font class="form_corpo"> 
            <%
				
				
				response.Write(estado_civil)
				
				
				%>
                  </font></td>
              </tr>
              <tr> 
                <td width="14%" height="10"> <div align="left"><font class="form_dado_texto"> M&atilde;e</font></div></td>
                <td width="2%"><div align="center">: </div></td>
                <td height="10"><font class="form_corpo"> 
            <%response.Write(mae)%>
                  </font></td>
                <td height="10"> <div align="left"><font class="form_dado_texto"> 
                  Falecida</font></div></td>
                <td><div align="center"><font class="form_dado_texto">?</font></div></td>
                <td height="10"><font class="form_corpo"> 
            <%response.Write(mae_fal)%>
                  </font></td>
                <td height="10"><div align="left"><font class="form_dado_texto"> </font></div></td>
                <td><div align="center"></div></td>
                <td height="10"><font class="form_dado_texto">&nbsp; </font></td>
              </tr>
            </table></td>
                </tr>
                
                <tr> 
                  
          <td colspan="3" class="tb_tit">Respons&aacute;veis</td>
                </tr>
                <tr> 
                  <td colspan="3"><table width="100%" border="0" cellspacing="0">
              <tr class="tb_corpo"
> 
                <td width="14%" height="10"> <div align="left"><font class="form_dado_texto"> Pedag&oacute;gico</font></div></td>
                <td width="2%"><div align="center">:</div></td>
                <td width="22%" height="10"> <font class="form_corpo"> 
            <%
		Set RSCONTP = Server.CreateObject("ADODB.Recordset")
		SQLP = "SELECT * FROM TB_Contatos WHERE TP_Contato ='"&resp_ped&"' And CO_Matricula ="& cod
		RSCONTP.Open SQLP, CONCONT
		
		resp_ped = RSCONTP("NO_Contato")
				  
				  response.Write(resp_ped)%>
                  </font></td>
                <td width="15%" height="10"> <div align="left"><font class="form_dado_texto"> 
                  Financeiro</font></div></td>
                <td width="1%"><div align="center">:</div></td>
                <td width="46%" height="10"><font class="form_corpo"> 
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
                  
          <td colspan="3" class="tb_tit"
>Familiares</td>
                </tr>
                <tr> 
                  <td colspan="3"><table width="100%" border="0" cellspacing="0">
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
                        
                <td width="14%" height="10"> 
                  <div align="left"><font class="form_dado_texto"> 
                    <%response.Write(no_tp_resp)%>
                  </font></div></td>
                        
                <td width="2%"><div align="center">:</div></td>
                <td width="84%" height="10"> <font class="form_corpo"> <a href="contatos.asp?or=01&cod=<%=cod%>&tp=<%=tp_resp%>&vd=<%=vindo%>&o=<%=obr%>" class="ativos"> 
                  <%response.Write(resp)%>
                  </a> </font> </td>
                      </tr>
                      <%RSCONTPR.MOVENEXT
end if
WEND
%>
                    </table></td>
                </tr>
                
                <tr> 
                  
          <td colspan="3" class="tb_tit"
>Dados Escolares </td>
                </tr>
                <tr> 
                  <td colspan="3"><table width="100%" border="0" cellspacing="0">
              <tr class="tb_corpo"
> 
                <td height="10" colspan="2"> <div align="left"><font class="form_dado_texto"> Col&eacute;gio de Origem</font></div></td>
                <td><div align="center">:</div></td>
                <td colspan="3"><font class="form_corpo">
                  <%response.Write(col_origem)%>
                </font></td>
                <td height="10" colspan="9">&nbsp;</td>
              </tr>
              <tr class="tb_corpo"
>
                <td height="10" colspan="2"><font class="form_dado_texto">Etapa cursada </font></td>
                <td><div align="center">:</div></td>
                <td colspan="3"><font class="form_corpo">
                  <%response.Write(cursada)%>
                </font></td>
                <td height="10"><font class="form_dado_texto">Local</font></td>
                <td><div align="center">:</div></td>
                <td height="10" colspan="7"><font class="form_corpo">
                  <%response.Write(cid_cursada&"/"&uf_cursada)%>
                </font></td>
              </tr>
              <tr class="tb_corpo"
> 
                <td height="10" colspan="2"> <div align="left"><font class="form_dado_texto">Data do Cadastro</font></div></td>
                <td><div align="center">:</div></td>
                <td colspan="3"><font class="form_corpo">
                  <%response.Write(cadastro)%>
                </font></td>
                <td height="10"> <div align="left"><font class="form_dado_texto">Data de Entrada na Escola</font></div></td>
                <td width="5"><div align="center">:</div></td>
                <td height="10" colspan="7"><font class="form_corpo">
                  <%response.Write(entrada)%>
                </font></td>
              </tr>
              
              <tr class="tb_corpo"
> 
                <td width="71" height="26"><div align="center"><font class="form_dado_texto">

                </font><font class="form_dado_texto">Ano Letivo </font></div></td>
                <td width="62" height="26"><div align="center"><font class="form_dado_texto">Matr&iacute;cula</font></div></td>
                <td width="13"><div align="center"></div></td>
                <td width="10"><div align="center"></div></td>
                <td width="88"><div align="center"><font class="form_dado_texto">Cancelamento</font></div></td>
                <td width="118"><div align="center"><font class="form_dado_texto">Situa&ccedil;&atilde;o</font></div></td>
                <td width="154"><div align="center"><font class="form_dado_texto">Unidade</font></div></td>
                <td><div align="center"></div></td>
                <td width="2">&nbsp;</td>
                <td width="173"><font class="form_dado_texto">Curso</font></td>
                <td width="79" height="26"> <div align="center"><font class="form_dado_texto"> 
              Etapa</font></div></td>
                <td width="2" height="26">&nbsp;</td>
                <td width="97"><div align="center"><font class="form_dado_texto">Turma </font></div></td>
                <td width="71" height="26"> <div align="center"><font class="form_dado_texto"> 
              Chamada</font></div></td>
                <td width="23" height="26">&nbsp;</td>
              			  
			  </tr>
                  <%

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Matriculas WHERE CO_Matricula ="& cod
		RS.Open SQL, CON1

while not RS.EOF
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
no_unidades = session("no_unidades")
no_grau = session("no_grau")
no_serie = session("no_serie")

%>			   
              <tr class="tb_corpo"
			  
> 
                <td height="10"> <div align="center"><font class="form_corpo"> 
            <%response.Write(ano_aluno)%>
                    </font></div></td>
                <td height="10"><div align="center"><font class="form_corpo">
                  <%response.Write(rematricula)%>
                </font></div></td>
                <td width="13"><div align="center"></div></td>
                <td width="10"><div align="center"></div></td>
                <td width="88"><div align="center"><font class="form_corpo">
                  <%response.Write(encerramento)%>
                </font></div></td>
                <td width="118"><div align="center"><font class="form_corpo">
                  <%
					
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
		RSCONTST.Open SQLCONTST, CON0
							
				no_situacao = RSCONTST("TX_Descricao_Situacao")	
					response.Write(no_situacao)%>
                </font></div></td>
                <td height="10"><div align="center"><font class="form_corpo">
                  <%response.Write(no_unidades)%>
                </font></div></td>
                <td><div align="center"></div></td>
                <td>&nbsp;</td>
                <td><div align="left"><font class="form_corpo">
                  <%response.Write(no_grau)%>
                </font></div></td>
                <td height="10"> <div align="center"><font class="form_corpo">
                  <%response.Write(no_serie)%>
                </font></div></td>
                <td height="10">&nbsp;</td>
                <td><div align="center"><font class="form_corpo">
                  <%response.Write(turma)%>
                </font></div></td>
                <td height="10"> <div align="center"><font class="form_corpo">
                  <%response.Write(cham)%>
                </font></div></td>
                <td height="10">&nbsp;</td>
              </tr>
                      <%RS.MOVENEXT
WEND
%>			  
            </table></td>
                </tr>
                
                <tr> 
                  <td colspan="3" class="tb_tit"
>&nbsp;</td>
                </tr>

              </table></td></tr>
</form>
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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