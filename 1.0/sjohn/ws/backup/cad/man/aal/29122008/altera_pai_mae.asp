<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<%
opt=request.QueryString("opt")

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR		
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0	


		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON1_aux = Server.CreateObject("ADODB.Connection") 
		ABRIR1_aux = "DBQ="& CAMINHO_al_aux & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1_aux.Open ABRIR1_aux		
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CONCONT_aux = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT_aux = "DBQ="& CAMINHO_ct_aux & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT_aux.Open ABRIRCONT_aux	



nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo")
ano_letivo_real = ano_letivo
sistema_local=session("sistema_local")
ori = request.QueryString("or")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
cod= request.QueryString("cod_cons")
vinc_erro= request.QueryString("v")

		Set RSCONTPR = Server.CreateObject("ADODB.Recordset")
		SQLCONTPR = "SELECT * FROM TB_Tipo_Contatos Order by NU_Prioridade_Combo"
		RSCONTPR.Open SQLCONTPR, CON0
total_tp_familiares=0		
while not RSCONTPR.EOF
tp_resp = RSCONTPR("TP_Contato")
no_tp_resp = RSCONTPR("TX_Descricao")
ordem_familiares=ordem_familiares&"##"&tp_resp&"!!"&no_tp_resp
total_tp_familiares=total_tp_familiares+1
if total_tp_familiares=1 then
foco_default="nulo"
end if
RSCONTPR.MOVENEXT
WEND

if opt="err1" or opt="err2" or opt="err3" or opt="err4" or opt="err5" then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TBI_Alunos WHERE CO_Matricula ="& cod
		RS.Open SQL, CON1_aux
				
		Set RSCONTA = Server.CreateObject("ADODB.Recordset")
		SQLA = "SELECT * FROM TBI_Contatos WHERE TP_Contato ='ALUNO' And CO_Matricula ="& cod
		RSCONTA.Open SQLA, CONCONT_aux		
else		
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod
		RS.Open SQL, CON1
		
		Set RSCONTA = Server.CreateObject("ADODB.Recordset")
		SQLA = "SELECT * FROM TB_Contatos WHERE TP_Contato ='ALUNO' And CO_Matricula ="& cod
		RSCONTA.Open SQLA, CONCONT				
end if		
		
codigo = RS("CO_Matricula")
nome_aluno = RS("NO_Aluno")

sexo = RS("IN_Sexo")


		
'	nascimento = RSCONTA("DA_Nascimento_Contato")		
'response.Write(">>"&nascimento)
if RSCONTA.EOF then
nascimento="0/0/0"
else
	if isnull(nascimento) or nascimento="" then
	nascimento="0/0/0"
	else
	nascimento = RSCONTA("DA_Nascimento_Contato")
	end if
end if
	nascimento = RSCONTA("DA_Nascimento_Contato")

	if isnull(nascimento) or nascimento="" then
		else
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

'response.Write(">>"&CAMINHO_ct)
if nasce="00/00/0" then 
nasce =""
end if
	end if

apelido= RS("NO_Apelido")
desteridade= RS("IN_Desteridade")
nacionalidade= RS("CO_Nacionalidade")
estadonat = RS("SG_UF_Natural")
cidnat = RS("CO_Municipio_Natural")
pais= RS("CO_Pais_Natural")
msn= RS("TX_MSN")
orkut= RS("TX_ORKUT")
religiao= RS("CO_Religiao")
cor_raca= RS("CO_Raca")

col_or= RS("NO_Colegio_Origem")
et_curs= RS("NO_Serie_Cursada")
uf_curs= RS("SG_UF_Cursada")
cid_curs= RS("CO_Municipio_Cursada")
da_cadastro= RS("DA_Cadastro")
da_entrada= RS("DA_Entrada_Escola")
tel = RSCONTA("NU_Telefones")
mail= RSCONTA("TX_EMail")
ocupacao= RSCONTA("CO_Ocupacao")
cpf= RSCONTA("CO_CPF_PFisica")
rg= RSCONTA("CO_RG_PFisica")
emitido= RSCONTA("CO_OERG_PFisica")
emissao= RSCONTA("CO_DERG_PFisica")
empresa= RSCONTA("NO_Empresa")
vinculado=RSCONTA("CO_Matricula_Vinc")

if isnull(vinculado) or vinculado="" then
pai= RS("NO_Pai")
mae= RS("NO_Mae")
pai_fal= RS("IN_Pai_Falecido")
mae_fal= RS("IN_Mae_Falecida")
resp_fin= RS("TP_Resp_Fin")
resp_ped= RS("TP_Resp_Ped")
sit_pais= RS("CO_Estado_Civil")


rua_res = RSCONTA("NO_Logradouro_Res")
num_res = RSCONTA("NU_Logradouro_Res")
comp_res = RSCONTA("TX_Complemento_Logradouro_Res")
bairrores= RSCONTA("CO_Bairro_Res")
cidres= RSCONTA("CO_Municipio_Res")
estadores= RSCONTA("SG_UF_Res")
cep = RSCONTA("CO_CEP_Res")
tel_res = RSCONTA("NU_Telefones_Res")
rua_com=RSCONTA("NO_Logradouro_Com")
num_com = RSCONTA("NU_Logradouro_Com")
comp_com = RSCONTA("TX_Complemento_Logradouro_Com")
bairrocom= RSCONTA("CO_Bairro_Com")
cidcom= RSCONTA("CO_Municipio_Com")
estadocom= RSCONTA("SG_UF_Com")
cepcom = RSCONTA("CO_CEP_Com")
tel_com = RSCONTA("NU_Telefones_Com")
else

		Set RS_vinc = Server.CreateObject("ADODB.Recordset")
		SQL_vinc = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& vinculado
		RS_vinc.Open SQL_vinc, CON1
		
pai= RS_vinc("NO_Pai")
mae= RS_vinc("NO_Mae")
pai_fal= RS_vinc("IN_Pai_Falecido")
mae_fal= RS_vinc("IN_Mae_Falecida")
resp_fin= RS_vinc("TP_Resp_Fin")
resp_ped= RS_vinc("TP_Resp_Ped")
sit_pais= RS_vinc("CO_Estado_Civil")		

		Set RSCONTA = Server.CreateObject("ADODB.Recordset")
		SQLA = "SELECT * FROM TB_Contatos WHERE TP_Contato ='ALUNO' And CO_Matricula ="& vinculado
		RSCONTA.Open SQLA, CONCONT
nome_vinculado = RSCONTA("NO_Contato")		
rua_res = RSCONTA("NO_Logradouro_Res")
num_res = RSCONTA("NU_Logradouro_Res")
comp_res = RSCONTA("TX_Complemento_Logradouro_Res")
bairrores= RSCONTA("CO_Bairro_Res")
cidres= RSCONTA("CO_Municipio_Res")
estadores= RSCONTA("SG_UF_Res")
cep = RSCONTA("CO_CEP_Res")
tel_res = RSCONTA("NU_Telefones_Res")
rua_com=RSCONTA("NO_Logradouro_Com")
num_com = RSCONTA("NU_Logradouro_Com")
comp_com = RSCONTA("TX_Complemento_Logradouro_Com")
bairrocom= RSCONTA("CO_Bairro_Com")
cidcom= RSCONTA("CO_Municipio_Com")
estadocom= RSCONTA("SG_UF_Com")
cepcom = RSCONTA("CO_CEP_Com")
tel_com = RSCONTA("NU_Telefones_Com")
end if

if isnull(vinc_erro) or vinc_erro="" then
cod_erro=cod
else
cod_erro=vinc_erro
end if
		Set RSCONTPR = Server.CreateObject("ADODB.Recordset")
		SQLCONTPR = "SELECT * FROM TB_Tipo_Contatos Order by NU_Prioridade_Combo"
		RSCONTPR.Open SQLCONTPR, CON0
foco=foco_default		
familiares=0
while not RSCONTPR.EOF
tp_resp = RSCONTPR("TP_Contato")
no_tp_resp = RSCONTPR("TX_Descricao")
		
		Set RSCONTACONT = Server.CreateObject("ADODB.Recordset")
		SQLAC = "SELECT * FROM TB_Contatos WHERE TP_Contato='"&tp_resp&"' and CO_Matricula ="& cod
		RSCONTACONT.Open SQLAC, CONCONT

if RSCONTACONT.EOF then
foco=foco
else

	tipo_contato=RSCONTACONT("TP_Contato")
		if familiares=0 then
		foco=tipo_contato
		end if
		
		if tipo_contato="PAI" OR tipo_contato="MAE" OR tipo_contato="ALUNO" then
		else
		nome_contato=RSCONTACONT("NO_Contato")
		end if
	familiares=familiares+1		
end if
RSCONTPR.MOVENEXT
WEND


if opt="err4" then
		Set RS_aux = Server.CreateObject("ADODB.Recordset")
		SQL_aux = "SELECT * FROM TBI_Alunos WHERE CO_Matricula ="& cod_erro
		RS_aux.Open SQL_aux, CON1_aux

resp_fin= RS_aux("TP_Resp_Fin")
resp_ped= RS_aux("TP_Resp_Ped")

foco=resp_fin
elseif opt="err5" then

		Set RS_aux = Server.CreateObject("ADODB.Recordset")
		SQL_aux = "SELECT * FROM TBI_Alunos WHERE CO_Matricula ="& cod_erro
		RS_aux.Open SQL_aux, CON1_aux

resp_fin= RS_aux("TP_Resp_Fin")
resp_ped= RS_aux("TP_Resp_Ped")
foco=resp_ped
else
foco=foco
end if

Call LimpaVetor2
 call navegacao (CON,chave,nivel)
navega=Session("caminho")

if isnull(pais) then 
pais = 10
end if

if isnull(estadores) then 
estadores = "RJ"
end if

if isnull(cidres) then 
cidres = 6001
end if

if isnull(estadonat) then 
estadonat = "RJ"
end if

if isnull(nacionalidade) then 
nacionalidade = 1
end if

if isnull(cidnat) then 
cidnat = 6001
end if

if comp_res = "nulo" then 
comp_res = ""
end if

if isnull(cid_cursada) then 
cid_cursada = 6001
end if

if isnull(uf_cursada) then 
uf_cursada = "RJ"
end if

if isnull(cep) or cep="" then
cep=""
else
cep5= lEFT(cep, 5)
cep3= Right(cep, 3)
cep=cep5&"-"&cep3
end if

if isnull(cepcom) or cepcom="" then
cepcom=""
else
cep5c= lEFT(cepcom, 5)
cep3c= Right(cepcom, 3)
cepcom=cep5c&"-"&cep3c
end if



				Set RSCONTATO_aux = Server.CreateObject("ADODB.Recordset")
				SQLAA_aux= "SELECT * FROM TBI_Contatos WHERE TP_Contato='PAI' and CO_Matricula ="&cod
				RSCONTATO_aux.Open SQLAA_aux, CONCONT_aux
if RSCONTATO_aux.eof then
pai_cadastrado="n"
else
pai_cadastrado="s"
end if				
				
				
				Set RSCONTATO_aux = Server.CreateObject("ADODB.Recordset")
				SQLAA_aux= "SELECT * FROM TBI_Contatos WHERE TP_Contato='MAE' and CO_Matricula ="&cod
				RSCONTATO_aux.Open SQLAA_aux, CONCONT_aux
				
if RSCONTATO_aux.eof then
mae_cadastrado="n"
else
mae_cadastrado="s"
end if					
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
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
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

function formatar(src, mask)
{
  var i = src.value.length;
  var saida = mask.substring(0,1);
  var texto = mask.substring(i)
if (texto.substring(0,1) != saida)
  {
        src.value += texto.substring(0,1);
  }
}
function ValidaCPF() {
		cpf = document.inclusao.cpf.value;
if (document.inclusao.cpf.value != "") {		
		erro = new String;
			//	document.write(cpf)
		cpf_split=cpf.split("-")
		cpf_split_0=cpf_split[0]
		cpf_split_1=cpf_split[1]				
		cpf=cpf_split_0+cpf_split_1
		if (cpf.length < 11) erro += "Sao necessarios 11 digitos para verificacao do CPF do Aluno! \n\n"; 
		var nonNumbers = /\D/;
		if (nonNumbers.test(cpf)) erro += "A verificacao de CPF do Aluno suporta apenas numeros! \n\n";	
		if (cpf == "00000000000" || cpf == "11111111111" || cpf == "22222222222" || cpf == "33333333333" || cpf == "44444444444" || cpf == "55555555555" || cpf == "66666666666" || cpf == "77777777777" || cpf == "88888888888" || cpf == "99999999999"){
			  erro += "Numero de CPF do Aluno invalido!"
		}
		var a = [];
		var b = new Number;
		var c = 11;
		for (i=0; i<11; i++){
			a[i] = cpf.charAt(i);
			if (i <  9) b += (a[i] *  --c);
		}
		if ((x = b % 11) < 2) { a[9] = 0 } else { a[9] = 11-x }
		b = 0;
		c = 11;
		for (y=0; y<10; y++) b += (a[y] *  c--); 
		if ((x = b % 11) < 2) { a[10] = 0; } else { a[10] = 11-x; }
		status = a[9] + ""+ a[10]
		if ((cpf.charAt(9) != a[9]) || (cpf.charAt(10) != a[10])){
			erro +="Digito verificador do CPF do Aluno com problema!";
		}
		if (erro.length > 0){
			alert(erro);
			document.inclusao.cpf.focus()
			return false;
		}
	}
		return true;
	}
function PopUp(){	
var answer = confirm ("CPF digitado possui dados previamente cadastrados. Esses dados serão aproveitados")
if (!answer){ 
alert ("Não é possível utilizar esse CPF sem reaproveitar os dados!")
//document.inclusao.nome_familiar.focus()

//	Return false
}else{
//alert ("OK!")
//	 return TRUE
  } 
  } 
function ValidaNomeFamiliar() {
avisado=false
		tp_vinc = unescape(document.inclusao.tp_vinc_familiar_aux.value)
		co_vinc = unescape(document.inclusao.co_vinc_familiar_aux.value)
		foco = document.inclusao.cod_familiar.value
		cod = document.inclusao.cod_consulta.value
		nome=document.inclusao.nome_familiar.value
			if (!avisado){ 
  if (document.inclusao.nome_familiar.value == "")
  {    alert("Por favor, digite um nome para o Familiar!")
    document.inclusao.nome_familiar.focus()
    return false
  }
  			avisado=true 
         setTimeout('avisado=false',10000)
		 }
//	BD_aux(nome,cod,foco,tp_vinc,co_vinc,'NO_Contato')			 
//    return true
}	 

function ValidaDataNasce() {
//	if (document.inclusao.nasce.value == "")
//    {    alert("Por favor, digite a data de nascimento do aluno!")
//    document.inclusao.nasce.focus()
//    return false
// } else{
 erro=0;
       	 hoje = new Date();
         anoAtual = hoje.getFullYear();
         barras = document.inclusao.nasce_fam.value.split("/");
         if (barras.length == 3){
                   dia = barras[0];
                   mes = barras[1];
                   ano = barras[2];
                   resultado = (!isNaN(dia) && (dia > 0) && (dia < 32)) && (!isNaN(mes) && (mes > 0) && (mes < 13)) && (!isNaN(ano) && (ano.length == 4) && (ano <= anoAtual && ano >= 1900));
                   if (!resultado) {
                             alert("Data de nascimento do familiar é inválida!");
                             inclusao.nasce_fam.focus();
                             return false;
                   }
         } else {
                   alert("Formato da data de nascimento do familiar invalido!");
                   inclusao.nasce_fam.focus();
                   return false;
         }
//	}
}
function ValidaDataEmissao() {
//	if (document.inclusao.nasce.value == "")
//    {    alert("Por favor, digite a data de nascimento do aluno!")
//    document.inclusao.nasce.focus()
//    return false
// } else{
 erro=0;
       	 hoje = new Date();
         anoAtual = hoje.getFullYear();
         barras = document.inclusao.nasce2_fam.value.split("/");
         if (barras.length == 3){
                   dia = barras[0];
                   mes = barras[1];
                   ano = barras[2];
                   resultado = (!isNaN(dia) && (dia > 0) && (dia < 32)) && (!isNaN(mes) && (mes > 0) && (mes < 13)) && (!isNaN(ano) && (ano.length == 4) && (ano <= anoAtual && ano >= 1900));
                   if (!resultado) {
                             alert("Data de emissão da identidade do familiar é inválida!");
                             inclusao.nasce2_fam.focus();
                             return false;
                   }
         } else {
                   alert("Formato da data de emissão da identidade do familiar invalido!");
                   inclusao.nasce2_fam.focus();
                   return false;
         }
//	}
}
function ValidaCPFFamiliar() {
		cpf = document.inclusao.cpf_fam.value;
		ord = unescape(document.inclusao.ordem_familiares.value)
		qtd_tp = unescape(document.inclusao.qtd_tipo_familiares.value)
		foco = document.inclusao.cod_familiar.value
		co_vinc = unescape(document.inclusao.cod_consulta.value)
		cod = document.inclusao.codigo.value
if (document.inclusao.cpf_fam.value != "") {		
		erro = new String;
			//	document.write(cpf)
		cpf_consulta=cpf	
		cpf_split=cpf.split("-")
		cpf_split_0=cpf_split[0]
		cpf_split_1=cpf_split[1]				
		cpf=cpf_split_0+cpf_split_1

		if (cpf.length < 11) erro += "Sao necessarios 11 digitos para verificacao do CPF do Familiar! \n\n"; 
		var nonNumbers = /\D/;
		if (nonNumbers.test(cpf)) erro += "A verificacao de CPF do Familiar suporta apenas numeros! \n\n";	
		if (cpf == "00000000000" || cpf == "11111111111" || cpf == "22222222222" || cpf == "33333333333" || cpf == "44444444444" || cpf == "55555555555" || cpf == "66666666666" || cpf == "77777777777" || cpf == "88888888888" || cpf == "99999999999"){
			  erro += "Numero de CPF do Familiar invalido!"
		}
		var a = [];
		var b = new Number;
		var c = 11;
		for (i=0; i<11; i++){
			a[i] = cpf.charAt(i);
			if (i <  9) b += (a[i] *  --c);
		}
		if ((x = b % 11) < 2) { a[9] = 0 } else { a[9] = 11-x }
		b = 0;
		c = 11;
		for (y=0; y<10; y++) b += (a[y] *  c--); 
		if ((x = b % 11) < 2) { a[10] = 0; } else { a[10] = 11-x; }
		status = a[9] + ""+ a[10]
		if ((cpf.charAt(9) != a[9]) || (cpf.charAt(10) != a[10])){
			erro +="Digito verificador com problema no CPF do Familiar!";
		}
		if (erro.length > 0){
			alert(erro);
			document.inclusao.cpf_fam.focus()
			return false;
		}
 VerificaCPFFamiliar(cpf_consulta,ord,qtd_tp,foco,co_vinc,cod)
 }		
//		return true;
	}	
	
	function ValidaNumResFam() {
 	num_res= document.inclusao.num_res_fam.value
	resultado =(!isNaN(num_res))
	        if (!resultado) {
                alert("O número do endereço residencial do familiar não pode conter letras. Para isso utilize o campo complemento")
				document.inclusao.num_res_fam.focus()
              return false;
                }	
}
	function ValidaNumComFam() {
 	num_com= document.inclusao.num_com_fam.value
	resultado =(!isNaN(num_com))
	        if (!resultado) {
                alert("O número do endereço comercial do familiar não pode conter letras. Para isso utilize o campo complemento")
				document.inclusao.num_com_fam.focus()
              return false;
                }	
}	
	function ValidaCepResFam() {
algarismos = document.inclusao.cep_fam.value  
  if (document.inclusao.cep_fam.value != "")
 { 
         barras = document.inclusao.cep_fam.value.split("-");
         if ((barras.length == 2) && (algarismos.length == 9))
		 {
                   cep0= barras[0];
                   cep1 = barras[1];
                   resultado = (!isNaN(cep0) && (cep0 > 10000) && (cep0 < 999999)) && (!isNaN(cep1) && (cep1 >= 0) && (cep1 < 999));
                   if (!resultado) {
                             alert("Formato do CEP do endereço residencial do familiar é invalido!");
                             inclusao.cep_fam.focus();
                             return false;
                   }
         } else {
                   alert("Formato do CEP do endereço residencial do familiar é invalido!");
                   inclusao.cep_fam.focus();
                  return false;
         }
	  }	
}	
	function ValidaCepComFam() {
algarismos = document.inclusao.cepcom_fam.value  
  if (document.inclusao.cepcom_fam.value != "")
 { 
         barras = document.inclusao.cepcom_fam.value.split("-");
         if ((barras.length == 2) && (algarismos.length == 9))
		 {
                   cep0= barras[0];
                   cep1 = barras[1];
                   resultado = (!isNaN(cep0) && (cep0 > 10000) && (cep0 < 999999)) && (!isNaN(cep1) && (cep1 >= 0) && (cep1 < 999));
                   if (!resultado) {
                             alert("Formato do CEP do endereço comercial do familiar é invalido!");
                             inclusao.cepcom_fam.focus();
                             return false;
                   }
         } else {
                   alert("Formato do CEP do endereço comercial do familiar é invalido!");
                   inclusao.cepcom_fam.focus();
                  return false;
         }
	  }	
}	
function checksubmit()
{
  if (document.inclusao.nome.value == "")
  {    alert("Por favor, digite um nome para o aluno!")
    document.inclusao.nome.focus()
    return false
  }
	if (document.inclusao.nasce.value == "")
    {    alert("Por favor, digite a data de nascimento do aluno!")
    document.inclusao.nasce.focus()
    return false
 } else{
 erro=0;
        hoje = new Date();
         anoAtual = hoje.getFullYear();
         barras = document.inclusao.nasce.value.split("/");
         if (barras.length == 3){
                   dia = barras[0];
                   mes = barras[1];
                   ano = barras[2];
                   resultado = (!isNaN(dia) && (dia > 0) && (dia < 32)) && (!isNaN(mes) && (mes > 0) && (mes < 13)) && (!isNaN(ano) && (ano.length == 4) && (ano <= anoAtual && ano >= 1900));
                   if (!resultado) {
                             alert("Data de nascimento do aluno é invalida!");
                             inclusao.nasce.focus();
                             return false;
                   }
         } else {
                   alert("Formato de data invalido!");
                   inclusao.nasce.focus();
                   return false;
         }
	}	 	 		 
  if (document.inclusao.sexo.value == "0")
  {    alert("Por favor, escolha o sexo do aluno!")
    document.inclusao.sexo.focus()
    return false
  }
  if (document.inclusao.tel.value == "")
  {    alert("Por favor, digite pelo menos um telefone para contato com o aluno!")
    document.inclusao.tel.focus()
    return false
  }
  if (document.inclusao.rua_res.value == "")
  {    alert("Por favor, digite a rua onde o aluno reside!")
    document.inclusao.rua_res.focus()
    return false
  } 
 	num_res= document.inclusao.num_res.value
 if (num_res != ""){ 
//     alert("Por favor, digite o número da residência do aluno!")
//    document.inclusao.num_res.focus()
//    return false
//	}else{
	resultado =(!isNaN(num_res))
	        if (!resultado) {
                alert("O número do endereço residencial do aluno não pode conter letras. Para isso utilize o campo complemento")
				document.inclusao.num_res.focus()
              return false;
                }	
  } 
 	num_com= document.inclusao.num_com.value
 if (num_com != ""){ 
//     alert("Por favor, digite o número da residência do aluno!")
//    document.inclusao.num_res.focus()
//    return false
//	}else{
	resultado =(!isNaN(num_com))
	        if (!resultado) {
                alert("O número do endereço comercial do aluno não pode conter letras. Para isso utilize o campo complemento")
				document.inclusao.num_com.focus()
              return false;
                }	
  }      

erro=0;
algarismos = document.inclusao.cep.value  
  if (document.inclusao.cep.value != "")
 { 
//     alert("Por favor, digite o CEP da residência do aluno!")
//   document.inclusao.cep.focus()
//    return false
//  } else  {   
//  	 
         barras = document.inclusao.cep.value.split("-");
         if ((barras.length == 2) && (algarismos.length == 9))
		 {
                   cep0= barras[0];
                   cep1 = barras[1];
                   resultado = (!isNaN(cep0) && (cep0 > 10000) && (cep0 < 999999)) && (!isNaN(cep1) && (cep1 >= 0) && (cep1 < 999));
                   if (!resultado) {
                             alert("Formato do CEP do endereço residencial do aluno é invalido!");
                             inclusao.cep.focus();
                             return false;
                   }
         } else {
                   alert("Formato do CEP do endereço residencial do aluno é invalido!");
                   inclusao.cep.focus();
                  return false;
         }
	  }
	algarismos = document.inclusao.cep_com.value  
   if (document.inclusao.cep_com.value != "")  
   {   
  		 
         barras = document.inclusao.cep_com.value.split("-");
         if ((barras.length == 2) && (algarismos.length == 9)){
                   cep0= barras[0];
                   cep1 = barras[1];
                   resultado = (!isNaN(cep0) && (cep0 > 10000) && (cep0 < 999999)) && (!isNaN(cep1) && (cep1 >= 0) && (cep1 < 999));
                   if (!resultado) {
                             alert("Formato do CEP do endereço comercial do aluno é invalido!");
                             inclusao.cep_com.focus();
                             return false;
                   }
         } else {
                   alert("Formato do CEP do endereço comercial do aluno é invalido!");
                   inclusao.cep_com.focus();
                   return false;
         }		 
   }
  if (document.inclusao.pai.value == "")
  {    alert("Por favor, digite o nome do pai do aluno!")
    document.inclusao.pai.focus()
    return false
  }   
  if (document.inclusao.mae.value == "")
  {    alert("Por favor, digite o nome da mãe do aluno!")
    document.inclusao.mae.focus()
    return false
  } 
  return true

}

function createXMLHTTP()
            {
                        try
                        {
                                   ajax = new ActiveXObject("Microsoft.XMLHTTP");
                        }
                        catch(e)
                        {
                                   try
                                   {
                                               ajax = new ActiveXObject("Msxml2.XMLHTTP");
                                               alert(ajax);
                                   }
                                   catch(ex)
                                   {
                                               try
                                               {
                                                           ajax = new XMLHttpRequest();
                                               }
                                               catch(exc)
                                               {
                                                            alert("Esse browser não tem recursos para uso do Ajax");
                                                            ajax = null;
                                               }
                                   }
                                   return ajax;
                        }
           
           
               var arrSignatures = ["MSXML2.XMLHTTP.5.0", "MSXML2.XMLHTTP.4.0",
               "MSXML2.XMLHTTP.3.0", "MSXML2.XMLHTTP",
               "Microsoft.XMLHTTP"];
               for (var i=0; i < arrSignatures.length; i++) {
                                                                          try {
                                                                                                             var oRequest = new ActiveXObject(arrSignatures[i]);
                                                                                                             return oRequest;
                                                                          } catch (oError) {
                                                                          }
                                      }
           
                                      throw new Error("MSXML is not installed on your system.");
                        }        
						                        
						
						
						 function VincularAluno(vinc,aluno)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?opt=v", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.dadesc.innerHTML =resultado_c
                                                           }
                                               }
                                               oHTTPRequest.send("vinc_pub=" + vinc +"&aluno_pub=" + aluno);
                                   }
						 function DesvincularAluno(vinc,aluno)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?opt=d", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_o  = oHTTPRequest.responseText;
resultado_o = resultado_o.replace(/\+/g," ")
resultado_o = unescape(resultado_o)
document.all.dadesc.innerHTML =resultado_o
                                                           }
                                               }
                                               oHTTPRequest.send("vinc_pub=" + vinc +"&aluno_pub=" + aluno);
                                   }								   
						 function recuperarEnd(cod,foco)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?opt=r", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_e  = oHTTPRequest.responseText;
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.end.innerHTML =resultado_e
                                                           }
                                               }
                                               oHTTPRequest.send("cod_pub=" + cod+"&foco_pub=" + foco);
                                   }								   

						 function recuperarOrigemEnd(cod,foco)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "executa.asp?opt=oe", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_oe  = oHTTPRequest.responseText;
resultado_oe = resultado_oe.replace(/\+/g," ")
resultado_oe = unescape(resultado_oe)
document.all.end.innerHTML =resultado_oe
                                                           }
                                               }
                                               oHTTPRequest.send("cod_pub=" + cod+"&foco_pub=" + foco);
                                   }								   

						 function recuperarCidNat(estadonat)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "cid_bairro.asp?opt=c&o=n&f=n", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_cid_nat  = oHTTPRequest.responseText;
resultado_cid_nat = resultado_cid_nat.replace(/\+/g," ")
resultado_cid_nat = unescape(resultado_cid_nat)
document.all.cid_nat.innerHTML =resultado_cid_nat
                                                           }
                                               }
                                               oHTTPRequest.send("c_pub=" + estadonat);
                                   }
						 function recuperarCidRes(estadores)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "cid_bairro.asp?opt=c&o=r&f=n", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_cid_res  = oHTTPRequest.responseText;
resultado_cid_res = resultado_cid_res.replace(/\+/g," ")
resultado_cid_res = unescape(resultado_cid_res)
document.all.cid_res.innerHTML =resultado_cid_res
document.all.bairro_res.innerHTML ="<select class=borda></select>"
                                                           }
                                               }
                                               oHTTPRequest.send("c_pub=" + estadores);
                                   }
						 function recuperarCidCom(estadocom)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "cid_bairro.asp?opt=c&o=c&f=n", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_cid_com  = oHTTPRequest.responseText;
resultado_cid_com = resultado_cid_com.replace(/\+/g," ")
resultado_cid_com = unescape(resultado_cid_com)
document.all.cid_com.innerHTML =resultado_cid_com
document.all.bairro_com.innerHTML ="<select class=borda></select>"
                                                           }
                                               }
                                               oHTTPRequest.send("c_pub=" + estadocom);
                                   }
						 function recuperarBairroRes(estadores,cidres)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "cid_bairro.asp?opt=b&o=r&f=n", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_bairro_res  = oHTTPRequest.responseText;
resultado_bairro_res = resultado_bairro_res.replace(/\+/g," ")
resultado_bairro_res = unescape(resultado_bairro_res)
document.all.bairro_res.innerHTML =resultado_bairro_res
                                                           }
                                               }
                                               oHTTPRequest.send("c_pub=" + estadores +"&b_pub=" + cidres);
                                   }
						 function recuperarBairroCom(estadocom,cidcom)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "cid_bairro.asp?opt=b&o=c&f=n", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_bairro_com  = oHTTPRequest.responseText;
resultado_bairro_com = resultado_bairro_com.replace(/\+/g," ")
resultado_bairro_com = unescape(resultado_bairro_com)
document.all.bairro_com.innerHTML =resultado_bairro_com
                                                           }
                                               }
                                               oHTTPRequest.send("c_pub=" + estadocom +"&b_pub=" + cidcom);
                                   }
						 function recuperarCidResFam(estadores)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "cid_bairro.asp?opt=c&o=r&f=s", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_cid_res_fam  = oHTTPRequest.responseText;
resultado_cid_res_fam = resultado_cid_res_fam.replace(/\+/g," ")
resultado_cid_res_fam = unescape(resultado_cid_res_fam)
document.all.cid_res_fam.innerHTML =resultado_cid_res_fam
document.all.bairro_res_fam.innerHTML ="<select class=borda></select>"
                                                           }
                                               }											   
											   
                                               oHTTPRequest.send("c_pub=" + estadores);
                                   }
						 function recuperarCidComFam(estadocom)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "cid_bairro.asp?opt=c&o=c&f=s", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_cid_com_fam  = oHTTPRequest.responseText;
resultado_cid_com_fam = resultado_cid_com_fam.replace(/\+/g," ")
resultado_cid_com_fam = unescape(resultado_cid_com_fam)
document.all.cid_com_fam.innerHTML =resultado_cid_com_fam
document.all.bairro_com_fam.innerHTML ="<select class=borda></select>"
                                                           }
                                               }
                                               oHTTPRequest.send("c_pub=" + estadocom);
                                   }
						 function recuperarCidCurs(estadocurs)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "cid_bairro.asp?opt=c&o=e&f=e", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_cid_com  = oHTTPRequest.responseText;
resultado_cid_com = resultado_cid_com.replace(/\+/g," ")
resultado_cid_com = unescape(resultado_cid_com)
document.all.cid_com2.innerHTML =resultado_cid_com
                                                           }
                                               }
                                               oHTTPRequest.send("c_pub=" + estadocurs);
                                   }								   
						 function recuperarBairroResFam(estadores,cidres)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "cid_bairro.asp?opt=b&o=r&f=s", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_bairro_res_fam  = oHTTPRequest.responseText;
resultado_bairro_res_fam = resultado_bairro_res_fam.replace(/\+/g," ")
resultado_bairro_res_fam = unescape(resultado_bairro_res_fam)
document.all.bairro_res_fam.innerHTML =resultado_bairro_res_fam
                                                           }
                                               }
                                               oHTTPRequest.send("c_pub=" + estadores +"&b_pub=" + cidres);
                                   }
						 function recuperarBairroComFam(estadocom,cidcom)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "cid_bairro.asp?opt=b&o=c&f=s", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_bairro_com_fam  = oHTTPRequest.responseText;
resultado_bairro_com_fam = resultado_bairro_com_fam.replace(/\+/g," ")
resultado_bairro_com_fam = unescape(resultado_bairro_com_fam)
document.all.bairro_com_fam.innerHTML =resultado_bairro_com_fam
                                                           }
                                               }
                                               oHTTPRequest.send("c_pub=" + estadocom +"&b_pub=" + cidcom);
                                   }								   

							 function recuperarPai(nome,tp,qld,cod)
                                   {
nome = escape(nome)
								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "pai_mae.asp", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_pai  = oHTTPRequest.responseText;
//resultado_pai = resultado_pai.replace(/\+/g," ")
//resultado_pai = unescape(resultado_pai)
//document.all.bd_familiar.innerHTML =resultado_pai
//document.all.nome_pai.innerHTML =resultado_pai
recuperarFamiliares('<%response.Write(ordem_familiares)%>','<%response.Write(total_tp_familiares)%>','PAI','<%response.Write(cod)%>','<%response.Write(cod)%>')
                                                           }
                                               }
                                               oHTTPRequest.send("nome_pub=" + nome+ "&tp_familiares=" + tp+ "&qld_pub=" + qld+ "&cod_pub=" + cod);											   
                                   }
							 function recuperarMae(nome,tp,qld,cod)
                                   {

nome = escape(nome)
								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "pai_mae.asp", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_mae  = oHTTPRequest.responseText;
//resultado_mae = resultado_mae.replace(/\+/g," ")
//resultado_mae = unescape(resultado_mae)
//document.all.bd_familiar.innerHTML =resultado_mae
//document.all.nome_mae.innerHTML =resultado_mae
recuperarFamiliares('<%response.Write(ordem_familiares)%>','<%response.Write(total_tp_familiares)%>','MAE','<%response.Write(cod)%>','<%response.Write(cod)%>')
                                                           }
                                               }
                                               oHTTPRequest.send("nome_pub=" + nome+ "&tp_familiares=" + tp+ "&qld_pub=" + qld+ "&cod_pub=" + cod);											   
                                   }
						function GravaFamiliares(ord,qtd_tp,foco,cod_vinc,cod)
                                   {
ord = escape(ord)
								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "familiares.asp?opt=zero", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_familiares  = oHTTPRequest.responseText;
resultado_familiares = resultado_familiares.replace(/\+/g," ")
resultado_familiares = unescape(resultado_familiares)
document.all.familiares.innerHTML =resultado_familiares
                                                           }
                                               }
                                               oHTTPRequest.send("ord_pub=" + ord +"&qtd_tp_pub=" + qtd_tp +"&foco_pub=" +foco+"&cod_vinc_pub=" +cod_vinc+"&cod_pub=" +cod);											   
                                   }								   
							 function recuperarFamiliares(ord,qtd_tp,foco,cod_vinc,cod)
                                   {
ord = escape(ord)
								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "familiares.asp", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_familiares  = oHTTPRequest.responseText;
resultado_familiares = resultado_familiares.replace(/\+/g," ")
resultado_familiares = unescape(resultado_familiares)
document.all.familiares.innerHTML =resultado_familiares
                                                           }
                                               }
                                               oHTTPRequest.send("ord_pub=" + ord +"&qtd_tp_pub=" + qtd_tp +"&foco_pub=" +foco+"&cod_vinc_pub=" +cod_vinc+"&cod_pub=" +cod);											   
                                   }
							 function criaFamiliar(ord,qtd_tp,foco,cod_vinc,cod)
                                   {

ord = escape(ord)
								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "familiares.asp?opt=i", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_cria_familiar  = oHTTPRequest.responseText;
resultado_cria_familiar = resultado_cria_familiar.replace(/\+/g," ")
resultado_cria_familiar = unescape(resultado_cria_familiar)
document.all.familiares.innerHTML =resultado_cria_familiar


                                                           }
                                               }
                                               oHTTPRequest.send("ord_pub=" + ord +"&qtd_tp_pub=" + qtd_tp +"&foco_pub=" +foco+"&cod_vinc_pub=" +cod_vinc+"&cod_pub=" +cod);											   
                                   }
							 function BdFamiliar(cod_aluno,nome_familiar,nasce_fam,ocupacao_fam,trabalho_fam,email_fam,cpf_fam,id_fam,tipo_id_fam,nasce2_fam,tel_fam,rua_res_fam,num_res_fam,comp_res_fam,estadores_fam,cidres_fam,bairrores_fam,cep_fam,tel_res_fam,rua_com_fam,num_com_fam,comp_com_fam,estadocom_fam,cidcom_fam,bairrocom_fam,cepcom_fam,tel_com_fam,cod_familiar,mes_end,aluno_vinculado,co_vinc_familiar_aux,tp_vinc_familiar_aux)
                                   {

nome_familiar = escape(nome_familiar)
trabalho_fam = escape(trabalho_fam)
email_fam = escape(email_fam)
tipo_id_fam = escape(tipo_id_fam)
rua_res_fam = escape(rua_res_fam)
comp_res_fam = escape(comp_res_fam)
rua_com_fam = escape(rua_com_fam)
comp_com_fam = escape(comp_com_fam)

								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "bd_aux.asp?opt=af", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_bd_familiar  = oHTTPRequest.responseText;
//resultado_bd_familiar = resultado_bd_familiar.replace(/\+/g," ")
//resultado_bd_familiar = unescape(resultado_bd_familiar)
//document.all.bd_familiar.innerHTML =resultado_bd_familiar
                                                           }
                                               }
                                        oHTTPRequest.send("cod_consulta=" + cod_aluno +"&nome_familiar=" + nome_familiar +"&nasce_fam=" +nasce_fam+"&ocupacao_fam=" +ocupacao_fam+"&trabalho_fam=" + trabalho_fam +"&email_fam=" +email_fam+"&cpf_fam=" +cpf_fam+"&id_fam=" + id_fam +"&tipo_id_fam="+ tipo_id_fam +"&nasce2_fam=" +nasce2_fam+"&tel_fam=" +tel_fam+"&rua_res_fam=" + rua_res_fam +"&num_res_fam=" +num_res_fam+"&comp_res_fam=" +comp_res_fam+"&estadores_fam=" + estadores_fam +"&cidres_fam=" +cidres_fam+"&bairrores_fam=" +bairrores_fam+"&cep_fam=" + cep_fam +"&tel_res_fam=" +tel_res_fam+"&rua_com_fam=" + rua_com_fam +"&num_com_fam=" +num_com_fam+"&comp_com_fam=" +comp_com_fam+"&estadocom_fam=" + estadocom_fam +"&cidcom_fam=" +cidcom_fam+"&bairrocom_fam=" +bairrocom_fam+"&cepcom_fam=" + cepcom_fam +"&tel_com_fam=" +tel_com_fam+"&cod_familiar=" +cod_familiar+"&mes_end=" +mes_end+"&aluno_vinculado=" +aluno_vinculado+"&co_vinc_familiar_aux=" +co_vinc_familiar_aux+"&tp_vinc_familiar_aux=" +tp_vinc_familiar_aux);											   
                                   }
// function IncluiFamiliar(cod_aluno,nome_familiar,nasce_fam,ocupacao_fam,trabalho_fam,email_fam,cpf_fam,id_fam,tipo_id_fam,nasce2_fam,tel_fam,rua_res_fam,num_res_fam,comp_res_fam,estadores_fam,cidres_fam,bairrores_fam,cep_fam,tel_res_fam,rua_com_fam,num_com_fam,comp_com_fam,estadocom_fam,cidcom_fam,bairrocom_fam,cepcom_fam,tel_com_fam,cod_familiar,mes_end,ord,qtd_tp,foco)
//
//                                   {
//
//
//nome_familiar = escape(nome_familiar)
//trabalho_fam = escape(trabalho_fam)
//email_fam = escape(email_fam)
//tipo_id_fam = escape(tipo_id_fam)
//rua_res_fam = escape(rua_res_fam)
//comp_res_fam = escape(comp_res_fam)
//rua_com_fam = escape(rua_com_fam)
//comp_com_fam = escape(comp_com_fam)

//ord = escape(ord)

								   
//                                               var oHTTPRequest = createXMLHTTP();
//                                               oHTTPRequest.open("post", "bd_aux.asp?opt=if", true);
//                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
//                                               oHTTPRequest.onreadystatechange=function() {
//                                                           if (oHTTPRequest.readyState==4){
//                                                                    var resultado_bd_familiar  = oHTTPRequest.responseText;
//resultado_bd_familiar = resultado_bd_familiar.replace(/\+/g," ")
//resultado_bd_familiar = unescape(resultado_bd_familiar)
//document.all.familiares.innerHTML =resultado_bd_familiar
//                                                           }
//                                               }
//                                               oHTTPRequest.send("cod_aluno=" + cod_aluno +"&nome_familiar=" + nome_familiar +"&nasce_fam=" +nasce_fam+"&ocupacao_fam=" +ocupacao_fam+"&trabalho_fam=" + trabalho_fam +"&email_fam=" +email_fam+"&cpf_fam=" +cpf_fam+"&id_fam=" + id_fam +"&tipo_id_fam="+ tipo_id_fam +"&nasce2_fam=" +nasce2_fam+"&tel_fam=" +tel_fam+"&rua_res_fam=" + rua_res_fam +"&num_res_fam=" +num_res_fam+"&comp_res_fam=" +comp_res_fam+"&estadores_fam=" + estadores_fam +"&cidres_fam=" +cidres_fam+"&bairrores_fam=" +bairrores_fam+"&cep_fam=" + cep_fam +"&tel_res_fam=" +tel_res_fam+"&rua_com_fam=" + rua_com_fam +"&num_com_fam=" +num_com_fam+"&comp_com_fam=" +comp_com_fam+"&estadocom_fam=" + estadocom_fam +"&cidcom_fam=" +cidcom_fam+"&bairrocom_fam=" +bairrocom_fam+"&cepcom_fam=" + cepcom_fam +"&tel_com_fam=" +tel_com_fam+"&cod_familiar=" +cod_familiar+"&mes_end=" +mes_end+"&ord_pub=" + ord +"&qtd_tp_pub=" + qtd_tp +"&foco_pub=" +foco);											   
								
//								   }

							 function ConfirmaExcluirFamiliares(ord,qtd_tp,foco,cod)
                                   {
ord = escape(ord)
								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "familiares.asp?opt=e", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_familiares  = oHTTPRequest.responseText;
resultado_familiares = resultado_familiares.replace(/\+/g," ")
resultado_familiares = unescape(resultado_familiares)
document.all.familiares.innerHTML =resultado_familiares
                                                           }
                                               }
                                               oHTTPRequest.send("ord_pub=" + ord +"&qtd_tp_pub=" + qtd_tp +"&foco_pub=" +foco+"&cod_pub=" +cod);											   
                                   }
							 function ExcluiFamiliares(ord,qtd_tp,foco,cod)
                                   {
ord = escape(ord)
								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "bd_aux.asp?opt=ef", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
//                                                                    var resultado_familiares  = oHTTPRequest.responseText;
//resultado_familiares = resultado_familiares.replace(/\+/g," ")
//resultado_familiares = unescape(resultado_familiares)
//document.all.familiares.innerHTML =resultado_familiares
recuperarFamiliares('<%response.Write(ordem_familiares)%>','<%response.Write(total_tp_familiares)%>','<%response.Write(foco)%>','<%response.Write(cod)%>','<%response.Write(cod)%>')

                                                           }
                                               }
                                               oHTTPRequest.send("ord_pub=" + ord +"&qtd_tp_pub=" + qtd_tp +"&foco_pub=" +foco+"&cod_pub=" +cod);											   
                                   }

								   			function GravaResponsaveis(variavel,bd,valor_resp,tipo_resp,cod)
                                   {

								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "bd_aux.asp?opt=re", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_responsaveis  = oHTTPRequest.responseText;
//resultado_responsaveis = resultado_responsaveis.replace(/\+/g," ")
//resultado_responsaveis = unescape(resultado_responsaveis)
//document.all.bd_familiar.innerHTML =resultado_responsaveis
                                                           }
                                               }
                                               oHTTPRequest.send("variavel_pub=" + variavel +"&bd_pub=" + bd +"&valor_resp_pub=" + valor_resp +"&tipo_resp_pub=" + tipo_resp +"&cod_pub=" +cod);											   
                                   }
								   
								   
								   
								   
function VerificaCPFFamiliar(cpf,ord,qtd_tp,foco,cod_vinc,cod)
 {
ord = escape(ord)

                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "familiares.asp?opt=cpf", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_vinculado  = oHTTPRequest.responseText;
resultado_vinculado = resultado_vinculado.replace(/\+/g," ")
resultado_vinculado = unescape(resultado_vinculado)
document.all.familiares.innerHTML =resultado_vinculado
                                                           }
                                               }
                                               oHTTPRequest.send("cpf_pub=" + cpf +"&ord_pub=" + ord +"&qtd_tp_pub=" + qtd_tp +"&foco_pub=" +foco+"&cod_vinc_pub=" + cod_vinc+"&cod_pub=" + cod);											   
                                   }
//function BD_aux(variavel,cod,foco,tp_vinc,cod_vinc,bd)
// {
//variavel = escape(variavel)
//
//                                               var oHTTPRequest = createXMLHTTP();
//                                               oHTTPRequest.open("post", "bd_aux.asp?opt=af", true);
//                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
//                                               oHTTPRequest.onreadystatechange=function() {
//                                                          if (oHTTPRequest.readyState==4){
//                                                                   var resultado_vinculado  = oHTTPRequest.responseText;
// resultado_vinculado = resultado_vinculado.replace(/\+/g," ")
// resultado_vinculado = unescape(resultado_vinculado)
// document.all.bd_familiar.innerHTML =resultado_vinculado
//                                                            }
//                                              }
//                                               oHTTPRequest.send("variavel_pub=" + variavel +"&cod_pub=" + cod +"&foco_pub=" + foco +"&tp_vinc_pub=" +tp_vinc+"&cod_vinc_pub=" + cod_vinc+"&bd_pub="+bd);											   
//                                   }								   								   								   								   								   
								   								   								   
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
<script language="javascript">

//desabilita o TAB

function KeyTest() {
	if (event.keyCode==9) {
		return false;
	}
}

</script>
</head>
<%
if opt="err1" or opt="err2" or opt="err3" or opt="err4" or opt="err5" then

	if isnull(vinculado) or vinculado="" then
		if isnull(vinc_erro) or vinc_erro="" then
		onload="recuperarFamiliares('"&ordem_familiares&"','"&total_tp_familiares&"','"&foco&"','"&cod&"','"&cod&"')"
		else
		onload="recuperarCurso("&vinc_erro&","&cod&")"
		end if
	else
		if isnull(vinc_erro) or vinc_erro="" then
		onload="recuperarCurso("&vinculado&","&cod&")"
		else
		onload="recuperarCurso("&vinc_erro&","&cod&")"
		end if
	end if
else
	if isnull(vinculado) or vinculado="" then
	onload="GravaFamiliares('"&ordem_familiares&"','"&total_tp_familiares&"','"&foco&"','"&cod&"','"&cod&"')"
	else
	onload="GravaFamiliares('"&ordem_familiares&"','"&total_tp_familiares&"','"&foco&"','"&vinculado&"','"&cod&"')"
	end if
end if
%>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onload="<%response.Write(onload)%>">
<%call cabecalho(nivel)
%>
<table width="1002" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
            
    <td width="1000" height="10" colspan="5" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
    </td>
  </tr> 
  <%if opt="ok" then%>
             <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,9709,2,0) %>
    </td>
  </tr>
     <%
	 end if
	 'if opt="ok2" then%>
      <!--        <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%'call mensagens(nivel,412,0,0) %>
    </td>
  </tr> --> 
   <%
   'end if
   if opt="err1" then%>
             <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,9710,1,0) %>
    </td>
  </tr> 
   <%
   end if
   if opt="err2" then%>
             <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,407,1,0) %>
    </td>
  </tr> 
   <%
   end if
   if opt="err3" then%>
             <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,408,1,0) %>
    </td>
  </tr> 
   <%
   end if
   if opt="err4" then%>
             <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,409,1,0) %>
    </td>
  </tr>    
   <%
   end if
   if opt="err5" then%>
             <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,410,1,0) %>
    </td>
  </tr>       
  <%end if%>
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,402,0,0) %>
    </td>
  </tr>			  
        <form action="bd.asp?opt=a" method="post" name="inclusao" id="inclusao" onSubmit="return checksubmit()">
          <tr>
      <td valign="top"> 
        <table width="1000" border="0" align="right" cellspacing="0"  class="tb_corpo">
          <tr> 
            <td width="841" class="tb_tit"
>Dados Pessoais</td>
            <td width="151" class="tb_tit"
> </td>
            <td width="2" class="tb_tit"
></td>
          </tr>
          <tr> 
            <td height="10"> <font class="form_corpo">&nbsp; </font> <table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="145" height="10"><font class="form_dado_texto">Matr&iacute;cula</font></td>
                  <td width="13">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="217" height="10">
<input name="codigo" type="hidden" class="borda" id="nome" size="50" value="<%response.Write(codigo)%>"> 
                    <font class="form_dado_texto"> 
                    <%response.Write(codigo)%>
                    </font> </td>
                  <td width="140" height="10"><font class="form_dado_texto">Nome:</font></td>
                  <td width="19">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td height="10"><input name="nome" type="text" class="borda" id="nome" value="<%response.Write(nome_aluno)%>" size="50" maxlength="60"></td>
                </tr>
                <tr> 
                  <td width="145" height="10"> 
                    <div align="left"><font class="form_dado_texto"> 
                      Apelido</font></div></td>
                  <td width="13">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="217" height="10">
<input name="apelido" type="text" class="borda" id="apelido" value="<%response.Write(apelido)%>" size="30" maxlength="15"></td>
                  <td width="140" height="10"> 
                    <div align="left"><font class="form_dado_texto"> 
                      Data de Nascimento</font></div></td>
                  <td width="19">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td height="10"><input name="nasce" type="text" class="borda" id="nasce" value="<%response.Write(nasce)%>" size="12" maxlength="10" onKeyup="formatar(this, '##/##/####')"></td>
                </tr>
                <tr> 
                  <td width="145" height="10"> 
                    <div align="left"><font class="form_dado_texto"> 
                      Sexo</font></div></td>
                  <td width="13">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="217" height="10"> 
                    <select name="sexo" class="borda" id="select14">
                      <%if sexo = "M" then%>
                      <option value="0"></option>
                      <option value="M" selected>Masculino</option>
                      <option value="F">Feminino</option>
                      <%elseif sexo = "F" then%>
                      <option value="0"></option>
                      <option value="M">Masculino</option>
                      <option value="F" selected>Feminino</option>
                      <%else%>
                      <option value="0" selected></option>
                      <option value="M">Masculino</option>
                      <option value="F">Feminino</option>
                      <%End IF%>
                    </select> &nbsp;</td>
                  <td width="140" height="10"> 
                    <div align="left"><font class="form_dado_texto"> 
                      Pa&iacute;s de Origem</font></div></td>
                  <td width="19">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td height="10"><select name="pais" class="borda" id="select6">
                      <%				
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Paises order by NO_Pais"
		RS1.Open SQL1, CON0
		
while not RS1.EOF						
CO_Pais= RS1("CO_Pais")
NO_Pais= RS1("NO_Pais")
CO_Pais=CO_Pais*1
if CO_Pais = pais then
%>
                      <option value="<%=CO_Pais%>" selected> 
                      <% =NO_Pais%>
                      </option>
                      <%else%>
                      <option value="<%=CO_Pais%>"> 
                      <% =NO_Pais%>
                      </option>
                      <%
end if				
RS1.MOVENEXT
WEND
%>
                    </select></td>
                </tr>
                <tr> 
                  <td width="145" height="10"> 
                    <div align="left"><font class="form_dado_texto"> 
                      Nacionalidade</font></div></td>
                  <td width="13">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="217" height="10">
<select name="nacionalidade" class="borda" id="nacionalidade">
                      <%				
		Set RS_nacional= Server.CreateObject("ADODB.Recordset")
		SQL_nacional = "SELECT * FROM TB_Nacionalidades order by TX_Nacionalidade"
		RS_nacional.Open SQL_nacional, CON0
		
while not RS_nacional.EOF						
co_nacional= RS_nacional("CO_Nacionalidade")
no_nacional= RS_nacional("TX_Nacionalidade")
if co_nacional = nacionalidade then
%>
                      <option value="<%=co_nacional%>" selected> 
                      <% =no_nacional%>
                      </option>
                      <%else%>
                      <option value="<%=co_nacional%>"> 
                      <% =no_nacional%>
                      </option>
                      <%end if						
RS_nacional.MOVENEXT
WEND
%>
                    </select></td>
                  <td width="140" height="10"> 
                    <div align="left"><font class="form_dado_texto"> 
                      Natural do Estado</font></div></td>
                  <td width="19">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td height="10"><font class="form_corpo"> 
                    <select name="estadonat" class="borda" onChange="recuperarCidNat(this.value)">
                      <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
SG_UF= RS2("SG_UF")
NO_UF= RS2("NO_UF")
if isnull(estadonat) then
estadonat="RJ"
end if
if SG_UF = estadonat then
%>
                      <option value="<%=SG_UF%>" selected> 
                      <% =NO_UF%>
                      </option>
                      <%else%>
                      <option value="<%=SG_UF%>"> 
                      <% =NO_UF%>
                      </option>
                      <%end if						
RS2.MOVENEXT
WEND
%>
                    </select>
                    </font></td>
                </tr>
                <tr> 
                  <td width="145" height="10"> 
                    <div align="left"><font class="form_dado_texto"> 
                      Natural da Cidade</font></div></td>
                  <td width="13">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="217" height="10"> 
                    <div id="cid_nat"> 
                      <select name="cidnat" class="borda" id="select">
                        <%
Set RS2m = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Municipios WHERE SG_UF='"&estadonat&"' order by NO_Municipio"
		RS2m.Open SQL2m, CON0
		
while not RS2m.EOF						
SG_UF= RS2m("CO_Municipio")
NO_UF= RS2m("NO_Municipio")
if isnull(cidnat) or cidnat="" then
cidnat=6001
end if
if SG_UF = cidnat then
%>
                        <option value="<%=SG_UF%>" selected> 
                        <% =NO_UF%>
                        </option>
                        <% else %>
                        <option value="<%=SG_UF%>"> 
                        <% =NO_UF%>
                        </option>
                        <%
end if	
RS2m.MOVENEXT
WEND
%>
                      </select>
                    </div></td>
                  <td width="140" height="10"> 
                    <div align="left"><font class="form_dado_texto"> 
                      Religi&atilde;o</font></div></td>
                  <td width="19">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td height="10"> <select name="religiao" class="borda" id="religiao">
                      <option value="0"></option>
                      <%				
		Set RS_re = Server.CreateObject("ADODB.Recordset")
		SQL_re = "SELECT * FROM TB_Religiao order by TX_Descricao_Religiao"
		RS_re.Open SQL_re, CON0
		
while not RS_re.EOF						
co_relig= RS_re("CO_Religiao")
no_relig= RS_re("TX_Descricao_Religiao")
co_relig=co_relig*1
religiao=religiao*1
if co_relig = religiao then
%>
                      <option value="<%=co_relig%>" selected> 
                      <% =no_relig%>
                      </option>
                      <%else%>
                      <option value="<%=co_relig%>"> 
                      <% =no_relig%>
                      </option>
                      <%end if						
RS_re.MOVENEXT
WEND
%>
                    </select> </td>
                </tr>
                <tr> 
                  <td width="145" height="10"> 
                    <div align="left"><font class="form_dado_texto"> 
                      Cor / Ra&ccedil;a</font></div></td>
                  <td width="13">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="217" height="10"> 
                    <select name="cor_raca" class="borda" id="cor_raca">
                      <option value="0"></option>
                      <%				
		Set RS_cor_raca = Server.CreateObject("ADODB.Recordset")
		SQL_cor_raca = "SELECT * FROM TB_Raca order by TX_Descricao_Raca"
		RS_cor_raca.Open SQL_cor_raca, CON0
		
while not RS_cor_raca.EOF						
co_cor_raca= RS_cor_raca("CO_Raca")
no_cor_raca= RS_cor_raca("TX_Descricao_Raca")
co_cor_raca=co_cor_raca*1
cor_raca=cor_raca*1
if co_cor_raca = cor_raca then
%>
                      <option value="<%=co_cor_raca%>" selected> 
                      <% =no_cor_raca%>
                      </option>
                      <%else%>
                      <option value="<%=co_cor_raca%>"> 
                      <% =no_cor_raca%>
                      </option>
                      <%end if						
RS_cor_raca.MOVENEXT
WEND
%>
                    </select></td>
                  <td width="140" height="10"> <font class="form_dado_texto"> 
                    Ocupa&ccedil;&atilde;o</font></td>
                  <td width="19">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="36%" height="10"><font class="form_corpo"> 
                    <select name="ocupacao" class="borda" id="ocupacao">
                      <%				
		Set RS_oc = Server.CreateObject("ADODB.Recordset")
		SQL_oc = "SELECT * FROM TB_Ocupacoes order by NO_Ocupacao"
		RS_oc.Open SQL_oc, CON0
		
while not RS_oc.EOF						
co_ocup= RS_oc("CO_Ocupacao")
no_ocup= RS_oc("NO_Ocupacao")
if co_ocup = ocupacao then
%>
                      <option value="<%=co_ocup%>" selected> 
                      <% =no_ocup%>
                      </option>
                      <%else%>
                      <option value="<%=co_ocup%>"> 
                      <% =no_ocup%>
                      </option>
                      <%end if						
RS_oc.MOVENEXT
WEND
%>
                    </select>
                    </font></td>
                </tr>
                <tr> 
                  <td width="145" height="10"><font class="form_dado_texto">Identidade</font></td>
                  <td width="13">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="217" height="10">
<input name="rg" type="text" class="borda" id="rg" value="<%response.Write(rg)%>" size="15" maxlength="15"></td>
                  <td width="140" height="10"><font class="form_dado_texto">Tipo 
                    - Data de Emiss&atilde;o </font></td>
                  <td width="19">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td height="10"><input name="tipo_id" type="text" class="borda" id="tipo_id" value="<%response.Write(emitido)%>" size="15" maxlength="15">
                    - 
                    <input name="nasce2" type="text" class="borda" id="nasce2" size="12" maxlength="10" value="<%response.Write(emissao)%>" onKeyup="formatar(this, '##/##/####')"></td>
                </tr>
                <tr> 
                  <td width="145" height="10">
<div align="left"><font class="form_dado_texto"> 
                      CPF</font></div></td>
                  <td width="13">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="217" height="10">
<input name="cpf" type="text" class="borda" id="cpf" onBlur="ValidaCPF(this.value)"  onKeyup="formatar(this, '#########-##')" value="<%response.Write(cpf)%>" size="15" maxlength="15"></td>
                  <td width="140" height="10">
<div align="left"><font class="form_dado_texto"> 
                      Empresa onde trabalha</font></div></td>
                  <td width="19">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td height="10"><input name="trabalho" type="text" class="borda" id="trabalho" value="<%response.Write(empresa)%>" size="30" maxlength="40"></td>
                </tr>
                <tr> 
                  <td width="145" height="10"><font class="form_dado_texto">E-mail</font></td>
                  <td width="13">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="217" height="10">
<input name="email" type="text" class="borda" id="email" value="<%response.Write(mail)%>" size="30" maxlength="50"></td>
                  <td width="140" height="10">&nbsp;</td>
                  <td width="19">
<div align="center"></div></td>
                  <td height="10">&nbsp;</td>
                </tr>
                <tr> 
                  <td width="145" height="10"><font class="form_dado_texto">Login 
                    Orkut</font></td>
                  <td width="13">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="217" height="10">
<input name="orkut" type="text" class="borda" id="orkut" value="<%response.Write(orkut)%>" size="30" maxlength="50"></td>
                  <td width="140" height="10"><font class="form_dado_texto">Login 
                    Messenger</font></td>
                  <td width="19">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td height="10"><input name="messenger" type="text" class="borda" id="messenger" value="<%response.Write(msn)%>" size="30" maxlength="50"></td>
                </tr>
                <tr> 
                  <td width="145" height="10"> 
                    <div align="left"><font class="form_dado_texto">Telefones 
                      de Contato</font></div></td>
                  <td width="13">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td width="217" height="10">
<div align="left"><font class="form_corpo"></font> 
                      <font class="form_corpo"> 
                      <input name="tel" type="text" class="borda" id="tel" value="<%response.Write(tel)%>" size="42" maxlength="100">
                      </font></div></td>
                  <td width="140" height="10"> 
                    <div align="left"><font class="form_dado_texto">Escrita</font></div></td>
                  <td width="19">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                  <td height="10"><select name="desteridade" id="desteridade" class="borda">
                      <%if desteridade = "S" then%>
                      <option value="S" selected>Destro</option>
                      <option value="N">Canhoto</option>
                      <%elseif desteridade = "N" then%>
                      <option value="S">Destro</option>
                      <option value="N" selected>Canhoto</option>
					  <%else%>
                      <option value="0"></option>					  
                      <option value="S">Destro</option>
                      <option value="N">Canhoto</option>
                      <%End IF%>
                    </select></td>
                </tr>
              </table></td>
            <td valign="top">&nbsp;</td>
            <td valign="top">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="3"> <div id="dadesc"> 
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td class="tb_tit"
>Vincular dados do Aluno</td>
                  </tr>
                  <tr> 
                    <td> <% 
vinculado=vinculado*1
codigo=codigo*1
if vinculado=codigo or isnull(vinculado) or vinculado="" then%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr bgcolor="#FFFFFF" background="../../../../img/fundo_interno.gif"> 
                          <td width="5%"  height="10"> <div align="left"><font class="form_dado_texto"> 
                              Matr&iacute;cula:</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                              </strong></font></div></td>
                          <td width="10%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
                            </font><font size="2" face="Arial, Helvetica, sans-serif"> 
                            <input name="vinculado_novo" type="text" class="borda" id="vinculado_novo" size="12">
                            </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
                            </font></td>
                          <td width="37%" height="10"> <div align="right"><font class="form_dado_texto"> 
                              </font></div></td>
                          <td width="2%" height="10" ><font size="2" face="Arial, Helvetica, sans-serif">&nbsp; 
                            </font></td>
                          <td width="46%" height="10"><font size="2" face="Arial, Helvetica, sans-serif"> 
                            <input name="Button" type="button" class="borda_bot" id="Submit" value="Vincular" onClick="VincularAluno(vinculado_novo.value,codigo.value)">
                            </font> </td>
                        </tr>
                      </table>
<%else%>			
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr bgcolor="#FFFFFF" background="../../../../img/fundo_interno.gif"> 
                          <td width="5%"  height="10"> <div align="left"><font class="form_dado_texto"> 
                              Matr&iacute;cula:</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                              </strong></font></div></td>
                  <td width="10%" height="10"><div align="left"><font class="form_dado_texto"> 
			  
              <%response.Write(vinculado)%><input name="vinculado" type="hidden" class="borda" id="vinculado" value="<%response.Write(vinculado)%>" size="12">
                              </font></div>
                    </td>
                  
          <td width="37%" height="10"> <div align="left"><font class="form_dado_texto"> 
              <%response.Write(nome_vinculado)%>
              </font></div></td>
                  <td width="2%" height="10" ><font size="2" face="Arial, Helvetica, sans-serif">&nbsp; 
                    </font></td>
                  <td width="46%" height="10"><font size="2" face="Arial, Helvetica, sans-serif">
                    <input type="button" name="Button" value="Desvincular"  class="borda_bot3" onClick="DesvincularAluno(vinculado.value,codigo.value)">
                    </font> </td>
                </tr>
              </table>
<%end if%>	</td>
                  </tr>
                  <tr class="tb_corpo"> 
                    <td class="tb_tit"
>Endere&ccedil;o Residencial</td>
                  </tr>
                  <tr class="tb_corpo"> 
                    <td height="10"> <table width="100%" border="0" cellspacing="0">
                        <tr> 
                          <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Logradouro</font></div></td>
                          <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                          <td width="217" height="10" class="tb_corpo"
><font class="form_corpo"> 
                            <input name="rua_res" type="text" class="borda" id="rua_res" value="<%response.Write(rua_res)%>" size="30" maxlength="60">
                            </font></td>
                          <td width="140" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
                          <td width="19" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                          <td width="196" class="tb_corpo"
><font class="form_corpo"> <font class="form_corpo"> <font class="form_corpo"> 
                            <input name="num_res" type="text" class="borda" id="num_res" value="<%response.Write(num_res)%>" size="12" maxlength="10">
                            </font></font></font></td>
                          <td width="90" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
                          <td width="11" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                          <td width="149" height="10" class="tb_corpo"
> <div align="left"><font class="form_corpo"> </font> <font class="form_corpo"> 
                              <input name="comp_res" type="text" class="borda" id="comp_res" value="<%response.Write(comp_res)%>" size="20" maxlength="30">
                              </font></div></td>
                        </tr>
                        <tr> 
                          <td width="145" height="21" class="tb_corpo"
><font class="form_dado_texto">Estado</font></td>
                          <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                          <td width="217" height="21" class="tb_corpo"
><font class="form_corpo"> <font class="form_corpo"> 
<select name="estadores" class="borda" id="estadores" onChange="recuperarCidRes(this.value)">
                              <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
SG_UF= RS2("SG_UF")
NO_UF= RS2("NO_UF")

if isnull(estadores) or estadores="" then
estadores="RJ"
end if

if SG_UF = estadores then
%>
                              <option value="<%=SG_UF%>" selected> 
                              <% =NO_UF%>
                              </option>
                              <% else %>
                              <option value="<%=SG_UF%>"> 
                              <% =NO_UF%>
                              </option>
                              <%
end if	
RS2.MOVENEXT
WEND
%>
                            </select>
                            </font> </font></td>
                          <td width="140" height="21" class="tb_corpo"
><font class="form_dado_texto">Cidade</font></td>
                          <td width="19" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                          <td width="196" class="tb_corpo"
><font class="form_corpo"> 
                            <div id="cid_res"> 
                              <select name="cidres" class="borda" id="select10" onChange="recuperarBairroRes(estadores.value,this.value)">
                                <%
Set RS2m = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Municipios WHERE SG_UF='"&estadores&"' order by NO_Municipio"
		RS2m.Open SQL2m, CON0
		
while not RS2m.EOF						
SG_UF= RS2m("CO_Municipio")
NO_UF= RS2m("NO_Municipio")
if isnull(cidres) or cidres="" then
cidres=6001
end if
if SG_UF = cidres then
%>
                                <option value="<%=SG_UF%>" selected> 
                                <% =NO_UF%>
                                </option>
                                <% else %>
                                <option value="<%=SG_UF%>"> 
                                <% =NO_UF%>
                                </option>
                                <%
end if	
RS2m.MOVENEXT
WEND
%>
                              </select>
                            </div>
                            </font></td>
                          <td width="90" class="tb_corpo"
><font class="form_dado_texto">Bairro</font></td>
                          <td width="11" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                          <td width="149" height="21" class="tb_corpo"
>  <div id="bairro_res"><font class="form_corpo"> 
                              <select name="bairrores" class="borda" id="select3">
                                <%
Set RS2b = Server.CreateObject("ADODB.Recordset")
		SQL2b = "SELECT * FROM TB_Bairros WHERE CO_Municipio="&cidres&" AND SG_UF='"&estadores&"' order by NO_Bairro"
		RS2b.Open SQL2b, CON0
		
while not RS2b.EOF						
SG_UF= RS2b("CO_Bairro")
NO_UF= RS2b("NO_Bairro")
if SG_UF=bairrores then
%>
                                <option value="<%=SG_UF%>" selected> 
                                <% =NO_UF%>
                                </option>
                                <%else
%>
                                <option value="<%=SG_UF%>"> 
                                <% =NO_UF%>
                                </option>
                                <%
end if
RS2b.MOVENEXT
WEND
%>
                              </select>
                              </font></div></td>
                        </tr>
                        <tr> 
                          <td width="145" height="10" class="tb_corpo"
><font class="form_dado_texto">CEP</font></td>
                          <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                          <td width="217" height="10" class="tb_corpo"
><font class="form_dado_texto"> 
                            <input name="cep" type="text" class="borda" id="cep" onKeyup="formatar(this, '#####-###')" value="<%response.Write(cep)%>" size="11" maxlength="9">
                            </font></td>
                          <td width="140" height="10" class="tb_corpo"
>&nbsp;</td>
                          <td width="19" class="tb_corpo"
>&nbsp;</td>
                          <td width="196" class="tb_corpo"
>&nbsp;</td>
                          <td width="90" class="tb_corpo"
>&nbsp;</td>
                          <td width="11" class="tb_corpo"
> <div align="center"></div></td>
                          <td width="149" height="10" class="tb_corpo"
>&nbsp;</td>
                        </tr>
                        <tr> 
                          <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">Telefones deste endere&ccedil;o</font></div></td>
                          <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                          <td height="10" colspan="2" class="tb_corpo"
><font class="form_corpo"> 
                            <input name="tel_res" type="text" class="borda" id="tel_res" value="<%response.Write(tel_res)%>" size="42" maxlength="50">
                            </font> <div align="left"></div></td>
                          <td width="19" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                          <td width="196" class="tb_corpo"
>&nbsp;</td>
                          <td width="90" class="tb_corpo"
>&nbsp;</td>
                          <td width="11" class="tb_corpo"
> <div align="center"></div></td>
                          <td width="149" height="10" class="tb_corpo"
>&nbsp;</td>
                        </tr>
                        <tr> 
                          <td height="10" colspan="9" class="tb_tit"
><div align="left">Endere&ccedil;o Comercial </div></td>
                        </tr>
                        <tr> 
                          <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Logradouro</font></div></td>
                          <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                          <td width="217" height="10" class="tb_corpo"
><font class="form_corpo"> 
                            <input name="rua_com" type="text" class="borda" id="rua_com" value="<%response.Write(rua_com)%>" size="30" maxlength="60">
                            </font></td>
                          <td width="140" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
                          <td width="19" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                          <td width="196" class="tb_corpo"
><font class="form_corpo"> 
                            <input name="num_com" type="text" class="borda" id="num_com" value="<%response.Write(num_com)%>" size="12" maxlength="10">
                            </font></td>
                          <td width="90" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
                          <td class="tb_corpo"
><div align="center"><font class="form_dado_texto">:</font></div></td>
                          <td width="149" height="10" class="tb_corpo"
> <div align="left"><font class="form_corpo"> </font> <font class="form_corpo"> 
                              <input name="comp_com" type="text" class="borda" id="comp_com"  value="<%response.Write(comp_com)%>" size="20" maxlength="30">
                              </font></div></td>
                        </tr>
                        <tr class="tb_corpo"
> 
                          <td width="145" height="26"><font class="form_dado_texto">Estado</font></td>
                          <td width="13"> <div align="left"><font class="form_dado_texto">:</font></div></td>
                          <td width="217" height="26"><font class="form_corpo"> 
 <font class="form_corpo"> 
                            <select name="estadocom" class="borda" id="select2" onChange="recuperarCidCom(this.value)">
                              <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
SG_UF= RS2("SG_UF")
NO_UF= RS2("NO_UF")
if isnull(estadocom) or estadocom=""  then
estadocom="RJ"
end if
if SG_UF = estadocom then
%>
                              <option value="<%=SG_UF%>" selected> 
                              <% =NO_UF%>
                              </option>
                              <%else%>
                              <option value="<%=SG_UF%>"> 
                              <% =NO_UF%>
                              </option>
                              <%end if						
RS2.MOVENEXT
WEND
%>
                            </select>
                            </font> </font></td>
                          <td width="140" height="26"><font class="form_dado_texto">Cidade</font></td>
                          <td width="19"> <div align="center"><font class="form_dado_texto">:</font></div></td>
                          <td width="196"> <div id="cid_com"> 
                              <select name="cidcom" class="borda" id="select10" onChange="recuperarBairroCom(estadocom.value,this.value)">
                                <%
Set RS2m = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Municipios WHERE SG_UF='"&estadocom&"' order by NO_Municipio"
		RS2m.Open SQL2m, CON0
		
while not RS2m.EOF						
SG_UF= RS2m("CO_Municipio")
NO_UF= RS2m("NO_Municipio")
if isnull(cidcom) or cidcom="" then
cidcom=6001
end if
if SG_UF = cidcom then
%>
                                <option value="<%=SG_UF%>" selected> 
                                <% =NO_UF%>
                                </option>
                                <% else %>
                                <option value="<%=SG_UF%>"> 
                                <% =NO_UF%>
                                </option>
                                <%
end if	
RS2m.MOVENEXT
WEND
%>
                              </select>
                            </div></td>
                          <td width="90"><font class="form_dado_texto">Bairro</font></td>
                          <td><div align="center"><font class="form_dado_texto">:</font></div></td>
                          <td width="149" height="26"> <div id="bairro_com"><font class="form_corpo"> 
                              <select name="bairrocom" class="borda" id="bairro">
                                <%
Set RS2b = Server.CreateObject("ADODB.Recordset")
		SQL2b = "SELECT * FROM TB_Bairros WHERE CO_Municipio="&cidcom&" AND SG_UF='"&estadocom&"' order by NO_Bairro"
		RS2b.Open SQL2b, CON0
		
while not RS2b.EOF						
SG_UF= RS2b("CO_Bairro")
NO_UF= RS2b("NO_Bairro")

if SG_UF=bairrocom then
%>
                                <option value="<%=SG_UF%>" selected> 
                                <% =NO_UF%>
                                </option>
                                <%else
%>
                                <option value="<%=SG_UF%>"> 
                                <% =NO_UF%>
                                </option>
                                <%
end if

RS2b.MOVENEXT
WEND
%>
                              </select>
                              </font></div></td>
                        </tr>
                        <tr class="tb_corpo"
> 
                          <td width="145" height="26"><font class="form_dado_texto">CEP</font></td>
                          <td width="13"> <div align="left"><font class="form_dado_texto">:</font></div></td>
                          <td width="217" height="26"><font class="form_dado_texto"> 
                            <input name="cep_com" type="text" class="borda" id="cep_com" onKeyup="formatar(this, '#####-###')" value="<%response.Write(cepcom)%>" size="11" maxlength="9">
                            </font></td>
                          <td width="140" height="26">&nbsp;</td>
                          <td width="19">&nbsp;</td>
                          <td width="196">&nbsp;</td>
                          <td width="90">&nbsp;</td>
                          <td><div align="center"></div></td>
                          <td width="149" height="26">&nbsp;</td>
                        </tr>
                        <tr class="tb_corpo"
> 
                          <td width="145" height="28"> <div align="left"><font class="form_dado_texto">Telefones 
                              deste endere&ccedil;o:</font></div></td>
                          <td width="13"> <div align="left"><font class="form_dado_texto">:</font></div></td>
                          <td height="28" colspan="2"><font class="form_corpo"> 
                            <input name="tel_com" type="text" class="borda" id="tel_com" value="<%response.Write(tel_com)%>" size="42" maxlength="50">
                            </font> <div align="left"></div></td>
                          <td width="19"> <div align="center"></div></td>
                          <td width="196">&nbsp;</td>
                          <td width="90">&nbsp;</td>
                          <td><div align="center"></div></td>
                          <td width="149" height="28">&nbsp;</td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr class="tb_corpo"> 
                    <td class="tb_tit"
>Filia&ccedil;&atilde;o</td>
                  </tr>
                  <tr class="tb_corpo"> 
                    <td><table width="100%" border="0" cellspacing="0">
                        <tr> 
                          <td width="145" height="26"> <div align="left"><font class="form_dado_texto"> 
                              Pai</font></div></td>
                          <td width="13"> <div align="left"><font class="form_dado_texto">:</font></div></td>
                          <td width="217" height="26"><font class="form_corpo"> 
							<div id="nome_pai">  
                            <input name="pai" type="text" class="borda" onBlur="recuperarPai(this.value,'p','<%response.Write(pai_cadastrado)%>','<%response.Write(cod)%>')" value="<%response.Write(pai)%>" onKeyDown="return KeyTest()" size="30" maxlength="50">
                           </div>	 </font></td>
                          <td width="140" height="26"> <div align="left"><font class="form_dado_texto"> 
                              Falecido</font></div></td>
                          <td width="19"> <div align="center"><font class="form_dado_texto">?</font></div></td>
                          <td width="196" height="26"><font class="form_corpo"> 
                            <select name="pai_falecido" class="borda">
                              <% if pai_fal = false then%>
                              <option value="n"selected>N&atilde;o</option>
                              <option value="s">Sim</option>
                              <%else%>
                              <option value="n">N&atilde;o</option>
                              <option value="s" selected>Sim</option>
                              <%end if%>
                            </select>
                            </font></td>
                          <td width="90" height="26"> <div align="left"><font class="form_dado_texto"> 
                              Situa&ccedil;&atilde;o dos Pais</font></div></td>
                          <td width="11"> <div align="center"><font class="form_dado_texto">:</font></div></td>
                          <td width="149" height="26"> <select name="sit_pais" class="borda" id="sit_pais">
                              <option value=0></option>
                              <%				
		Set RS_ec = Server.CreateObject("ADODB.Recordset")
		SQL_ec = "SELECT * FROM TB_Estado_Civil order by CO_Estado_Civil"
		RS_ec.Open SQL_ec, CON0
		
while not RS_ec.EOF						
co_ec= RS_ec("CO_Estado_Civil")
no_ec= RS_ec("TX_Estado_Civil")

if co_ec=sit_pais then
%>
                              <option value="<%=co_ec%>" selected> 
                              <% =no_ec%>
                              </option>
                              <%
else							  
%>
                              <option value="<%=co_ec%>"> 
                              <% =no_ec%>
                              </option>
                              <%
end if							  						
RS_ec.MOVENEXT
WEND
%>
                            </select></td>
                        </tr>
                        <tr> 
                          <td width="145" height="10"> <div align="left"><font class="form_dado_texto"> 
                              M&atilde;e</font></div></td>
                          <td width="13"> <div align="left"><font class="form_dado_texto">:</font></div></td>
                          <td width="217" height="10"><font class="form_corpo"> 
                            <div id="nome_mae">  
							<input name="mae" type="text" class="borda" onBlur="recuperarMae(this.value,'m','<%response.Write(mae_cadastrado)%>','<%response.Write(cod)%>')" value="<%response.Write(mae)%>" onKeyDown="return KeyTest()" size="30" maxlength="50">
                            </div></font></td>
                          <td width="140" height="10"> <div align="left"><font class="form_dado_texto"> 
                              Falecida</font></div></td>
                          <td width="19"> <div align="center"><font class="form_dado_texto">?</font></div></td>
                          <td width="196" height="10"><font class="form_corpo"> 
                            <select name="mae_falecido" class="borda">
                              <% if mae_fal = false then%>
                              <option value="n"selected>N&atilde;o</option>
                              <option value="s">Sim</option>
                              <%else%>
                              <option value="n">N&atilde;o</option>
                              <option value="s" selected>Sim</option>
                              <%end if%>
                            </select>
                            </font></td>
                          <td width="90" height="10"> <div align="left"><font class="form_dado_texto"> 
                              </font></div></td>
                          <td width="11"> <div align="center"></div></td>
                          <td width="149" height="10"><font class="form_dado_texto">&nbsp; 
                            </font></td>
                        </tr>
                      </table></td>
                  </tr>
                  <tr class="tb_corpo"> 
                    <td width="145" class="tb_tit"
>Familiares</td>
                  </tr>
                  <tr class="tb_corpo"> 
                    <td> <div id="familiares"> </div></td>
                  </tr>
                  <tr class="tb_corpo"> 
                    <td><div id="responsaveis"> </div></td>
                  </tr>
                  <tr class="tb_corpo"> 
                    <td></td>
                  </tr>
                  <tr class="tb_corpo"> 
                    <td></td>
                  </tr>
                </table>
              </div></td>
          </tr>
          <tr>
            <td colspan="3"><table width="100%" border="0" cellspacing="0">
                <tr> 
                          <td height="10" colspan="9" class="tb_tit"
><div align="left">Dados Escolares</div></td>
                        </tr>
                        <tr> 
                          
                  <td width="145" height="10" class="tb_corpo"
> 
                    <div align="left"><font class="form_dado_texto"> Col&eacute;gio de Origem</font></div></td>
                          
                  <td width="13" class="tb_corpo"
> 
                    <div align="left"><font class="form_dado_texto">:</font></div></td>
                          
                  <td width="217" height="10" class="tb_corpo"
><font class="form_corpo"> 
                    <input name="col_or" type="text" class="borda" id="col_or" value="<%response.Write(col_or)%>" size="42" maxlength="40">
                            </font></td>
                          
                  <td width="140" height="10" class="tb_corpo"
> 
                    <div align="left"><font class="form_dado_texto">Etapa cursada </font></div></td>
                          <td width="18" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                          
                  <td width="230" class="tb_corpo"
><font class="form_corpo"><font class="form_corpo"> 
                    <input name="et_curs" type="text" class="borda" id="et_curs" value="<%response.Write(et_curs)%>" size="42" maxlength="50">
                            </font></font></td>
                          
                  <td width="56" class="tb_corpo"
>&nbsp;</td>
                          
                  <td width="11" class="tb_corpo"
>
<div align="center"></div></td>
                          
                  <td width="149" height="10" class="tb_corpo"
> 
                    <div align="left"><font class="form_corpo"> </font> <font class="form_corpo"> 
                              </font></div></td>
                        </tr>
                        <tr class="tb_corpo"
> 
                          
                  <td width="145" height="26"><font class="form_dado_texto">Estado</font></td>
                          
                  <td width="13"> 
                    <div align="left"><font class="form_dado_texto">:</font></div></td>
                          
                  <td width="217" height="26"><font class="form_corpo"> <font class="form_corpo"><font class="form_corpo"> 
                    <select name="estadocurs" class="borda" id="select5" onChange="recuperarCidCurs(this.value)">
                              <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
SG_UF= RS2("SG_UF")
NO_UF= RS2("NO_UF")
if isnull(uf_curs) then
uf_curs="RJ"
end if
if SG_UF = uf_curs then
%>
                              <option value="<%=SG_UF%>" selected> 
                              <% =NO_UF%>
                              </option>
                              <%else%>
                              <option value="<%=SG_UF%>"> 
                              <% =NO_UF%>
                              </option>
                              <%end if						
RS2.MOVENEXT
WEND
%>
                            </select>
                            </font> </font> </font></td>
                          
                  <td width="140" height="26"><font class="form_dado_texto">Cidade</font></td>
                          <td width="18"> <div align="center"><font class="form_dado_texto">:</font></div></td>
                          
                  <td width="230">
<div id="cid_com2"> <font class="form_corpo"><font class="form_corpo"> 
                              <select name="cid_curs" class="borda" id="select7">
                                <%
Set RS2m = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Municipios WHERE SG_UF='"&uf_curs&"' order by NO_Municipio"
		RS2m.Open SQL2m, CON0
		
while not RS2m.EOF						
SG_UF= RS2m("CO_Municipio")
NO_UF= RS2m("NO_Municipio")

if SG_UF = cid_curs then
%>
                                <option value="<%=SG_UF%>" selected> 
                                <% =NO_UF%>
                                </option>
                                <% else %>
                                <option value="<%=SG_UF%>"> 
                                <% =NO_UF%>
                                </option>
                                <%
end if	
RS2m.MOVENEXT
WEND
%>
                              </select>
                              </font></font></div></td>
                          
                  <td width="56">&nbsp;</td>
                          
                  <td width="11">
<div align="center"></div></td>
                          
                  <td width="149" height="26"> </td>
                        </tr>
                        <tr class="tb_corpo"
> 
                          
                  <td width="145" height="26"><font class="form_dado_texto">Data 
                    do Cadastro</font></td>
                          
                  <td width="13"> 
                    <div align="left"><font class="form_dado_texto">:</font></div></td>
                          
                  <td width="217" height="26"><font class="form_dado_texto"> 
                    <%response.Write(da_cadastro)%>
                    <input name="da_cadastro" type="hidden" id="da_cadastro" value="<%response.Write(da_cadastro)%>">
                    </font></td>
                          
                  <td width="140" height="26"><font class="form_dado_texto">Data 
                    de Entrada na Escola</font></td>
                          <td width="18"> <div align="center"><font class="form_dado_texto">:</font></div></td>
                          
                  <td width="230">
<input name="da_entrada" type="text" class="borda" id="da_entrada" value="<%response.Write(da_entrada)%>" size="12" maxlength="10" onKeyup="formatar(this, '##/##/####')"></td>
                          
                  <td width="56">&nbsp;</td>
                          
                  <td width="11">
<div align="center"></div></td>
                          
                  <td width="149" height="26">&nbsp;</td>
                        </tr>
                      </table></td>
          </tr>
        </table></td></tr>
		          <tr> 
            <td colspan="3"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr class="tb_corpo"> 
                  <td colspan="3"><hr></td>
                </tr>
                <tr> 
                  <td width="33%"><div align="center"> 
                      <input type="button" name="Submit2" value="Voltar" class="borda_bot3" onClick="MM_goToURL('parent','index.asp?nvg=WS-CA-MA-AAL')">
                    </div></td>
                  <td width="34%">&nbsp;</td>
                  <td width="33%"> <div align="center"> 
                      <input type="submit" name="Submit" value="Confirmar" class="borda_bot">
                    </div></td>
                </tr>
              </table></td>
          </tr>
</form>
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.gif" width="1000" height="40"></td>
  </tr>
</table>
<div id="bd_familiar"></div>
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