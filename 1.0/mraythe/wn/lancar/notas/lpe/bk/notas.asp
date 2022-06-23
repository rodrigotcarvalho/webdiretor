<%'On Error Resume Next%>
<!--#include file="../../../../../global/mensagens.asp" -->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/parametros.asp"-->
<!--#include file="../../../../inc/utils.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes3.asp"-->


<%
Server.ScriptTimeout = 1800 'valor em segundos
opt=request.QueryString("opt")
obr = request.QueryString("obr")
totalExibe = request.QueryString("exb")
exibe_tudo = "N"
if isnull(totalExibe) or totalExibe="" then
 totalExibe = 366
 exibe_tudo = "S" 
end if
session("totalExibe")=totalExibe
autoriza=session("autoriza")
grupo_usuario=session("grupo_usuario") 
nvg = session("chave")
ano_letivo = request.QueryString("ano")
co_usr = session("co_user")
grupo=session("grupo")
chave=nvg
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
nivel=4
trava=session("trava")
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CONer= Server.CreateObject("ADODB.Connection") 
		ABRIRer = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONer.Open ABRIRer		

		Set CONg = Server.CreateObject("ADODB.Connection") 
		ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONg.Open ABRIRg		
		
		Set CON0= Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		


if opt = "err" then
'response.write("session(obr) = "&session("obr"))
	dados= split(session("obr"), "$!$")
	co_materia = dados(0)
	unidade= dados(1)
	curso= dados(2)
	etapa= dados(3)
	turma= dados(4)
	periodo = dados(5)
	ano_letivo = dados(6)
	co_prof = dados(7)
	co_usr = session("co_usr")
	
	errante=0
	valido="s"
	javascript=""
elseif opt="ok" or opt = "vt" then

	unidade= session("unidades")
	curso= session("grau")
	etapa= session("serie")
	turma= session("turma")
	co_materia = session("co_materia")
	periodo = session("periodo")
	co_prof = session("co_prof")
	co_usr = session("co_usr")
	tb = session("nota")	
	ano_letivo = session("ano_letivo")
	voltaDireto = session("voltaDireto")				
	
else
	grava_nota=session("grava_nota")
	web_professor = request.form("wn")
	if web_professor="S" then
		unidade= request.form("unidade")
		curso= request.form("curso")
		etapa= request.form("etapa")
		turma= request.form("turma")
		co_materia = request.form("mat_prin")
		periodo = request.form("periodo")
		co_prof = Session("co_prof")
		voltaDireto = "S"							
	else

		dados= split(grava_nota, "?" )
		unidade= dados(0)
		curso= dados(1)
		etapa= dados(2)
		turma= dados(3)
		co_materia = request.querystring("d")
		periodo = request.querystring("p")
		co_prof = request.querystring("pr")
		voltaDireto = session("voltaDireto")	
				
	end if
	co_usr = session("co_usr")
	errante=0
	valido="s"
	javascript=""
end if
grava_nota= unidade&"?"&curso&"?"&etapa&"?"&turma
session("grava_nota")=grava_nota
session("voltaDireto")= voltaDireto

session("co_materia")=co_materia
session("unidades")=unidade
session("grau")=curso
session("serie")=etapa
session("turma")=turma
session("periodo")=periodo
session("co_prof") = co_prof 


obr=co_materia&"$!$"&unidade&"$!$"&curso&"$!$"&etapa&"$!$"&turma&"$!$"&periodo&"$!$"&ano_letivo&"$!$"&co_prof
session("obr")=obr

		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"'"	
		Set RS = CONg.Execute(CONEXAO)


if RS.EOF then
response.Write("<div align=center><font size=2 face=Courier New, Courier, mono  color=#990000><b>Esta turma não está disponível no momento</b></font><br")
response.Write("<font size=2 face=Courier New, Courier, mono  color=#990000><a href=javascript:window.history.go(-1)>voltar</a></font></div>")

else
nota = RS("TP_Nota")
coordenador = RS("CO_Cord")
end if
session("obr")=obr
session("nota")=nota
session("coordenador")=coordenador
 call navegacao (CON,chave,nivel)
navega=Session("caminho")

%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../../../estilos.css" type="text/css">
<script src="http://code.jquery.com/jquery-latest.min.js" type="text/javascript"></script>
<script type="text/javascript" language="javascript">

$(document).ready(function () {

    // Ao mover a barra de rolagem da tabela, mover seus cabecalhos e o 'versus'

    $("div#tabela").scroll(function () {

        $('div#tabela #cabecalhoHorizontal, #versus').css('top', $(this).scrollTop());

        $('div#tabela #cabecalhoVertical, #versus').css('left', $(this).scrollLeft());

    });

});

</script>
<script language="JavaScript" type="text/JavaScript">
<!--

function MM_popupMsg(msg) { //v1.0
  alert(msg);
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
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
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresiz!=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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

function mudar_cor_focus(celula){
   celula.style.backgroundColor="#D8FF9D"

}
function mudar_cor_blur_par(celula){
   celula.style.backgroundColor="#FFFFFF"
} 
function mudar_cor_blur_impar(celula){
   celula.style.backgroundColor="#FFFFE1"
} 
function mudar_cor_blur_erro(celula){
   celula.style.backgroundColor="#CC0000"
}  
function checksubmit()
{
// if (document.nota.pt.value == "")
//  {    alert("Por favor digite um peso para os Testes!")
//    document.nota.pt.focus()
//    return false
//  }
//  if (isNaN(document.nota.pt.value))
//  {    alert("O peso dos Testes deve ser um número!")
//    document.nota.pt.focus()
//    return false
//  }  
//    if (document.nota.pp.value == "")
//  {    alert("Por favor digite um peso para as Provas!")
//    document.nota.pp.focus()
//    return false
//  }
//  if (isNaN(document.nota.pp.value))
//  {    alert("O peso das Provas deve ser um número!")
//    document.nota.pp.focus()
//    return false
//  }
  return true
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
<script language="javascript"> 
  
    function keyPressed(TB, e, max_right, max_bottom)  
    { 
        if (e.keyCode == 40 || e.keyCode == 13) { // arrow down 
            if (TB.split("c")[0] < max_bottom) 
            document.getElementById(eval(TB.split("c")[0] + '+1') + 'c' + TB.split("c")[1]).focus(); 
            if (TB.split("c")[0] == max_bottom) 
            document.getElementById(1 + 'c' + TB.split("c")[1]).focus();


        } 
  
        if (e.keyCode == 38) { // arrow up 
            if(TB.split("c")[0] > 1) 
            document.getElementById(eval(TB.split("c")[0] + '-1') + 'c' + TB.split("c")[1]).focus(); 
            if (TB.split("c")[0] == 1) 
            document.getElementById(max_bottom + 'c' + TB.split("c")[1]).focus(); 
		
        } 
  
        if (e.keyCode == 37) { // arrow left 
            if(TB.split("c")[1] > 1) 
            document.getElementById(TB.split("c")[0] + 'c' + eval(TB.split("c")[1] + '-1')).focus();             
            if (TB.split("c")[1] == 1) 
            document.getElementById(TB.split("c")[0] + 'c' + max_right).focus(); 

		}   
  
        if (e.keyCode == 39) { // arrow right 
            if(TB.split("c")[1] < max_right) 
            document.getElementById(TB.split("c")[0] + 'c' + eval(TB.split("c")[1] + '+1')).focus();  
            if (TB.split("c")[1] == max_right) 
            document.getElementById(TB.split("c")[0] + 'c' + 1).focus(); 

		}                  
    } 
function scrollTo(hash) {
    location.hash = "#" + hash;
}
function show(){
 //alert("cheked the button - worked");
  scrollTo('topo');
  document.getElementById("carregando").style.visibility= 'visible' ;
  document.getElementById("tabela").style.visibility='hidden';
 // document.getElementById("carregando_fundo").style.visibility= 'visible' ; 
  document.getElementById("MSG1").style.display= 'none' ;
  document.getElementById("MSG2").style.display= 'none' ;  
  document.getElementById("MSG3").style.display = 'block' ;    
  document.getElementById("MSG3").style.visibility= 'visible' ;   
  }  
</script> 
<style type="text/css">

	body

	{

		font-family:"Trebuchet MS", Arial, Helvetica, sans-serif;

		font-size:13px;

	}

	div#tabela

	{

		width: 990px;      /* Largura da minha tabela na tela */

		height: 600px;     /* Altura da minha tabela na tela */

		overflow: auto;    /* Barras de rolagem automáticas nos eixos X e Y */

		margin: 0 auto;    /* O 'auto' é para ficar no centro da tela */

		position:relative; /* Necessário para os cabecalhos fixos */

		top:0;             /* Necessário para os cabecalhos fixos */

		left:0;            /* Necessário para os cabecalhos fixos */

	}

	div#tabela table

	{

		border-collapse:collapse; /* Sem espaços entre as células */

	}

	div#tabela table td

	{

		font-size:12px;

		font-family:"Trebuchet MS", Arial, Helvetica, sans-serif;

		border:1px solid #d8d8d8;	

		/*width:70px;      /* Células precisam ter altura e largura fixas */

		min-width:20px;  /* Se você não colocar isso, as células menores que 70px vão ser diminuídas */

		max-width:70px; */

		height:30px;

		min-height:30px;

		max-height:30px;

	}

	div#tabela table#cabecalhoHorizontal td, 

	div#tabela table#cabecalhoVertical td 

	{

		background-color:buttonface;

	}

	div#tabela table#cabecalhoHorizontal

	{

		margin-left:352px;  /* 70px de largura do cabecalho vertical + 2 pixels das bordas do cabecalho */

		position:absolute; /* Posição variável em relação ao topo da div#tabela */

		top:0;             /* Posição inicial em relação ao topo da div#tabela */

		z-index:5;         /* Para ficar por cima da tabela de dados */

	}

	div#tabela table#cabecalhoHorizontal td 

	{

		text-align:center;

		vertical-align:middle;

	}

	div#tabela table#cabecalhoVertical

	{

		margin-top:64px;   /* 30px de altura do cabecalho horizontal + 2 pixels das bordas do cabecalho + 1px */

		position:absolute; /* Posição variável em relação a esquerda da div#tabela */

		left:0;            /* Posição inicial em relação a esquerda da div#tabela */

		z-index:5;         /* Para ficar por cima da tabela de dados */

	}

	div#tabela table#cabecalhoVertical td

	{

		white-space:nowrap; /* Não quebrar linhas */

		text-align:left;



		/* Aqui temos um problema: preciso de uma margem, mas a largura da margem é somada à largura

		 * da célula e por isso a largura extrapola o tamanho máximo definido (70px). 

		 * Por isso, aqui eu diminuo a largura para, somada à margem, ficar do tamanho certo.

		*/

/*		width:173px;

		min-width:15px;

		max-width:345px;*/

		padding-left:5px;

	}

	div#tabela table#dados

	{

		margin-top:63px;  /* 30px de altura do cabecalho horizontal + 2 pixels das bordas do cabecalho + 1 px*/

		margin-left:350px; /* 70px de largura do cabecalho vertical + 2 pixels das bordas do cabecalho */

		z-index:2;		  /* Menor que dos cabecalhos, para que fique por detrás deles */

	}

	div#tabela table#dados td

	{

		background:white;

		text-align:center;

	}

	

	/* Célula com o 'X', que virtualmente pertence ao cabecalho vertical e horizontal */

	div#tabela #versus

	{

 		display:inline-block;

		position:absolute;

		top:0;

		left:0;

		z-index:10;



		height:63px;

		line-height:63px;

		width:350px;

		min-width:71px;

		

		text-align:center;

		vertical-align:middle;

		border:1px solid #d8d8d8;

		background-color:#F4FAE8;

		color:#A1D16D;

	}

</style>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" background="../../../../img/fundo.gif" marginheight="0" <%response.Write(javascript)%>><a name="topo"></a>
<%

IF imprime="1"then
else
 call cabecalho (nivel) 
 end if%>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
    <td height="10" valign="top" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> </td>
  </tr>
    <%

	

call GeraNomes(co_materia,unidade,curso,etapa,CON0)

no_materia= session("no_materia")
no_unidade= session("no_unidades")
no_curso= session("no_grau")
no_etapa= session("no_serie")


nome_prof = session("nome_prof") 
tp=	session("tp")

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("m", now) 
data = dia &"/"& mes &"/"& ano
horario = hora & ":"& min
acesso_prof = session("acesso_prof")


		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE CO_Professor= "& co_prof &"AND NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' AND CO_Materia_Principal = '"& co_materia &"'"
		Set RS = CONg.Execute(CONEXAO)
periodo=periodo*1
if periodo=1 then
	ST_Per_1 = RS("ST_Per_1")
elseif periodo=2 then
	ST_Per_2 = RS("ST_Per_2")
elseif periodo=3 then
	ST_Per_3 = RS("ST_Per_3")
elseif periodo=4 then
	ST_Per_4 = RS("ST_Per_4")
elseif periodo=5 then
	ST_Per_5 = RS("ST_Per_5")
elseif periodo=6 then
	ST_Per_6 = RS("ST_Per_6")
end if
tp = session("tp")

planilha_notas = RS("TP_Nota")

bancoPauta = escolheBancoPauta(planilha_notas,p_subopcao,p_outro)
caminhoBancoPauta = verificaCaminhoBancoPauta(bancoPauta,p_subopcao,p_outro)
session("bancoPauta") = bancoPauta
session("caminhoBancoPauta") = caminhoBancoPauta

		Set RS8 = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT * FROM TB_Materia where CO_Materia='"& co_materia &"'"
		RS8.Open SQL8, CON0

		if RS8.EOF then
			response.Write(co_materia&" não possui nome cadastrado<br>")				
		else
			co_mat_prin= RS8("CO_Materia_Principal")
		end if
		
		if co_mat_prin ="" or isnull(co_mat_prin) then
			co_mat_prin=co_materia
		end if

		Set CONPauta = Server.CreateObject("ADODB.Connection") 
		ABRIRPauta = "DBQ="& caminhoBancoPauta & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONPauta.Open ABRIRPauta
		
		wrkQtdAulasLancadas = 0		
		Set RSP = Server.CreateObject("ADODB.Recordset")
		SQL = "Select TB_Pauta_Aula.NU_Pauta, TB_Pauta_Aula.NU_Seq, TB_Pauta_Aula.DT_Aula, TB_Pauta_Aula.NU_Tempo,TB_Pauta.NU_Dia_Previsto  from TB_Pauta INNER JOIN TB_Pauta_Aula on TB_Pauta.NU_Pauta=TB_Pauta_Aula.NU_Pauta WHERE CO_Professor  = "& co_prof &" AND CO_Materia_Principal = '"& co_mat_prin &"' AND CO_Materia = '"& co_materia &"' AND NU_Unidade  = "& unidade &" AND CO_Curso  = '"& curso &"' AND CO_Etapa  = '"& etapa &"' AND CO_Turma  = '"& turma &"' AND NU_Periodo = "& periodo		
		Set RSP = CONPauta.Execute(SQL)
		qtdPrevistas=NULL			
		while not RSP.EOF

			wrkQtdAulasLancadas = wrkQtdAulasLancadas+1	
			if isnull(qtdPrevistas) then
				qtdPrevistas = RSP("NU_Dia_Previsto")
			end if	
		RSP.MOVENEXT 
		Wend	


		if wrkQtdAulasLancadas = 0 then
			response.Redirect("alterar.asp?ini=S")
		end if	

%>
            <%if opt = "ok" then%>
            <tr>         
    <td height="10" valign="top"> 
      <%
		call mensagens(4,664,2,dados)		
%>
      <div align="center"></div></td>
            </tr>			
                        <%elseif opt= "err" then %>
            <tr> 
    <td height="10" valign="top"> 
      <%
	call mensagens(nivel,671,1,dados)
%>
</td>
            </tr>
            <%end if
%>
            <% IF trava="s" or (co_usr<>coordenador AND grupo="COO") then%>
            <tr>     
    <td height="10" valign="top"> 
      <%
	 	 call mensagens_escolas(ambiente_escola,nivel,9701,"inf",0,0,0)	
	  %>
</td>
            </tr>
		<% ELSEIF (autoriza=5 OR co_usr=coordenador) AND trava<>"s" AND ((periodo = 1 and ST_Per_1="x") OR (periodo = 2 and ST_Per_2="x") OR (periodo = 3 and ST_Per_3="x") OR (periodo = 4 and ST_Per_4="x") OR (periodo = 5 and ST_Per_5="x") OR (periodo = 6 and ST_Per_6="x")) then%>
            <tr>     
    <td height="10" valign="top"> 
      <%
	 	 call mensagens_escolas(ambiente_escola,nivel,640,"err",0,0,0)		  
	  %>
</td>
            </tr>


            <%elseif (periodo = 1 and ST_Per_1="x") OR (periodo = 2 and ST_Per_2="x") OR (periodo = 3 and ST_Per_3="x") OR (periodo = 4 and ST_Per_4="x") OR (periodo = 5 and ST_Per_5="x") OR (periodo = 6 and ST_Per_6="x") then%>
            <tr> 
    <td height="10" valign="top"> 
      <%
	 	 call mensagens_escolas(ambiente_escola,nivel,624,"inf",0,0,0)			
%>
</td>
            </tr>

            <% end if%>
<%if opt= "cln" then %>
            <tr> 
    <td height="10" valign="top"> 
      <%
	call mensagens_escolas(ambiente_escola,nivel,621,"inf",0,0,0)			
%>
</td>
            </tr>
            <% end if%>						
	            <tr> 
    <td height="10" valign="top"> 
    <DIV ID="MSG1">
      <%
		call mensagens(4,667,0,totalExibe)		  
%></DIV>
    </td>
<% if exibe_tudo = "N" then 
colwidth = "20%"
%>    
<tr> 
    <td height="10" valign="top"> 
    <DIV ID="MSG2">    
      <%
		call mensagens(4,669,0,totalExibe)		  
%>
</DIV>
    </td>    
            </tr>	
<%
else
colwidth = "25%"
end if%>           
<tr> 
    <td height="10" valign="top"> 

	<DIV ID="MSG3" style="display:none"><%	call mensagens(nivel,670,0,"T") %></DIV> 
    </td>    
            </tr>	     
 		
            <tr class="tb_tit"> 
              
    <td height="15" class="tb_tit">&nbsp;Grade de Aulas</td>
            </tr>
            <tr> 
    <td height="36" valign="top"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="230" class="tb_subtit"><div align="center"><strong>PER&Iacute;ODO </strong></div></td>
          <td width="145" class="tb_subtit"> 
            <div align="center"><strong>UNIDADE 
              </strong></div></td>
          <td width="145" class="tb_subtit"> 
            <div align="center"><strong>CURSO 
              </strong></div></td>
          <td width="145" class="tb_subtit"> 
            <div align="center"><strong>ETAPA 
              </strong></div></td>
          <td width="145" class="tb_subtit"> 
            <div align="center"><strong>TURMA 
              </strong></div></td>
          <td width="190" class="tb_subtit"> 
            <div align="center"><strong>DISCIPLINA</strong></div></td>
        </tr>
        <tr>
          <td width="230"><div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
            <%
	
			
		Set RSper = Server.CreateObject("ADODB.Recordset")
		SQLper = "SELECT * FROM TB_Periodo where NU_Periodo= "&periodo
		RSper.Open SQLper, CON0

NO_Periodo= RSper("NO_Periodo")
dataInicio = RSper("DA_Inicio_Periodo")
dataFim = RSper("DA_Fim_Periodo")

if isnull(dataInicio) or dataInicio="" then

else
	dataInicio = formata(dataInicio,"DD/MM/YYYY")
end if

if isnull(dataFim) or dataFim="" then

else
	dataFim = formata(dataFim,"DD/MM/YYYY")
end if

response.Write(NO_Periodo)%>
          </font></div></td>
          <td width="145"> 
            <div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%response.Write(no_unidade)%>
              </font></div></td>
          <td width="145"> 
            <div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%response.Write(no_curso)%>
              </font></div></td>
          <td width="145"> 
            <div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%
response.Write(no_etapa)%>
              </font></div></td>
          <td width="145"> 
            <div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%
response.Write(turma)%>
              </font></div></td>
          <td width="190"> 
            <div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%

response.Write(no_materia)%>
              </font> </div></td>
        </tr>
        <tr>
          <td width="230">&nbsp;</td>
          <td width="145">&nbsp;</td>
          <td width="145">&nbsp;</td>
          <td width="145">&nbsp;</td>
          <td width="145">&nbsp;</td>
          <td width="190">&nbsp;</td>
        </tr>
        <tr>
          <td width="230" align="center" class="form_dado_texto">In&iacute;cio: <%response.Write(dataInicio)%> Fim: <%response.Write(dataFim)%></td>
          <td colspan="2" align="center" class="form_dado_texto">Total de Aulas Previstas:<%
		  if not isnull(qtdPrevistas) then
			  response.Write(qtdPrevistas)
		  end if%></td>
          <td colspan="2" align="center" class="form_dado_texto">Total de Aulas Lan&ccedil;adas:
          <%response.Write(wrkQtdAulasLancadas)%></td>
          <td width="190" align="center" class="form_dado_texto">Legenda: P-Presen&ccedil;a, F-Falta</td>
        </tr>
        <tr>
          <td align="center" class="form_dado_texto">&nbsp;</td>
          <td>&nbsp;</td>
          <td class="form_dado_texto">&nbsp;</td>
          <td>&nbsp;</td>
          <td align="right">&nbsp;</td>
          <td align="right" class="form_dado_texto">&nbsp;</td>
        </tr>
      </table></td>
            </tr>
      <tr> 
        
    <td valign="top"> <div id="carregando"  align="center" style="position:absolute;  top: 300px; width:1000px; z-index: 4; height: 150px; visibility: hidden;">
				  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="75" height="75" vspace="80" title="Carregando">
				    <param name="movie" value="../../../../img/carregando.swf">
				    <param name="quality" value="high">
				    <param name="wmode" value="transparent">
				    <embed src="../../../../img/carregando.swf" width="75" height="75" vspace="80" quality="high" wmode="transparent" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash"></embed>
			      </object>
			 </div>              
<!--				<div id="carregando_fundo" align="center" style="position:absolute; width:1000px; z-index: 3; height: 150px; visibility: hidden; background-color:#FFF; top: 250px; filter: Alpha(Opacity=90, FinishOpacity=100, Style=0, StartX=0, StartY=100, FinishX=100, FinishY=100);">  </div> -->
<form action="redireciona.asp" name="pauta" method="post" onSubmit="return checksubmit()">    
<div id="tabela">
    <span id="versus">Aulas Lan&ccedil;adas</span>
<%

totalExibe=totalExibe*1

	if opt="cln" then
		if bancoPauta ="Pauta_A" then

		Call pauta(CAMINHO_al,caminhoBancoPauta,unidade,curso,etapa,turma,co_materia,periodo,ano_letivo,co_prof,"A","cln",totalExibe,0)
		else
			if bancoPauta="Pauta_B" then
			Call pauta(CAMINHO_al,caminhoBancoPauta,unidade,curso,etapa,turma,co_materia,periodo,ano_letivo,co_prof,"B","cln",totalExibe,0)
			else
				if bancoPauta ="Pauta_C" then
				Call pauta(CAMINHO_al,caminhoBancoPauta,unidade,curso,etapa,turma,co_materia,periodo,ano_letivo,co_prof,"C","cln",totalExibe,0)
				else
				response.Write("ERRO")
				End if
			end if
		end if
	
	ELSEIF ((co_usr=coordenador and autoriza=5) AND trava<>"s") or (grupo<>"COO" and autoriza=5 AND trava<>"s") then
		if bancoPauta ="Pauta_A" then
		Call pauta(CAMINHO_al,caminhoBancoPauta,unidade,curso,etapa,turma,co_materia,periodo,ano_letivo,co_prof,"A","edit",totalExibe,0)
		else
			if bancoPauta="Pauta_B" then
			Call pauta(CAMINHO_al,caminhoBancoPauta,unidade,curso,etapa,turma,co_materia,periodo,ano_letivo,co_prof,"B","edit",totalExibe,0)
			else
				if bancoPauta ="Pauta_C" then
				Call pauta(CAMINHO_al,caminhoBancoPauta,unidade,curso,etapa,turma,co_materia,periodo,ano_letivo,co_prof,"C","edit",totalExibe,0)
				else
				response.Write("ERRO")
				End if
			end if
		end if	
	else
		if bancoPauta ="Pauta_A" then
		Call pauta(CAMINHO_al,caminhoBancoPauta,unidade,curso,etapa,turma,co_materia,periodo,ano_letivo,co_prof,"A","blq",totalExibe,0)
		else
			if bancoPauta="Pauta_B" then
			Call pauta(CAMINHO_al,caminhoBancoPauta,unidade,curso,etapa,turma,co_materia,periodo,ano_letivo,co_prof,"B","blq",totalExibe,0)
			else
				if bancoPauta ="Pauta_C" then
				Call pauta(CAMINHO_al,caminhoBancoPauta,unidade,curso,etapa,turma,co_materia,periodo,ano_letivo,co_prof,"C","blq",totalExibe,0)
				else
				response.Write("ERRO")
				End if
			end if
			
		end if
	end if	
%>    
 




</div>
<table width="1000" border="0" align="center" cellspacing="0">
            <tr>
              <td colspan="5"><hr /></td>
            </tr>
            <tr>
              <td width="<%response.Write(width)%>%"><div align="center">
<% 
if voltaDireto="S" then
	urlVolta="index.asp?nvg=WN-LN-LN-LPE&ori=1"
else
	urlVolta="altera.asp"
end if

	exibe_bt_alterar = "S"
IF autoriza=5 AND co_usr=coordenador then
	exibe_bt_alterar = "S"
ELSEIF trava="s" OR ((periodo = 1 and ST_Per_1="x") OR (periodo = 2 and ST_Per_2="x") OR (periodo = 3 and ST_Per_3="x") OR (periodo = 4 and ST_Per_4="x") OR (periodo = 5 and ST_Per_5="x") OR (periodo = 6 and ST_Per_6="x")) then
	exibe_bt_alterar = "N"
ELSEIF autoriza<>5 then
	exibe_bt_alterar = "N"
END IF

%>             
                <input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','<%response.Write(urlVolta)%>');return document.MM_returnValue" value="Voltar" />
              </div></td>
              <% if exibe_tudo = "N" then %>  
              <td width="<%response.Write(colwidth)%>" align="center"><input name="tudo" type="button" class="botao_prosseguir" id="tudo"  onclick = "show();MM_goToURL('parent','notas.asp?d=<%= co_materia%>&pr=<% = co_prof%>&p=<%= periodo%>');return document.MM_returnValue;" value="Exibir Pauta Completa" /></td>
              <%end if%>
              <td width="<%response.Write(colwidth)%>"><div align="center">
              	<% if exibe_bt_alterar = "S" then %>
                <input name="incluir" type="button" class="botao_prosseguir" id="incluir"  onclick = "MM_goToURL('parent','alterar.asp');return document.MM_returnValue;" value="Incluir" />
                <%end if%>
              </div></td>
              <td width="<%response.Write(colwidth)%>" align="center">
              	<% if exibe_bt_alterar = "S" then %>
              	<input name="submit" type="submit" class="botao_cancelar" id="alterar"  value="Alterar" />
              	<%end if%>
              </td>
              <td width="<%response.Write(colwidth)%>"><div align="center">
              	<% if exibe_bt_alterar = "S" then %>
                <input type="submit" name="submit" value="Excluir" class="botao_excluir" id="Excluir" />
                <%end if%>
                <input name="unidade" type="hidden" id="unidade" value="<%=unidade%>" />
                <input name="curso" type="hidden" id="curso" value="<%=curso%>" />
                <input name="etapa" type="hidden" id="etapa" value="<%=etapa%>" />
                <input name="turma" type="hidden" id="turma" value="<%=turma%>" />
                <input name="co_materia" type="hidden" id="co_materia" value="<%= co_materia%>" />
                <input name="periodo" type="hidden" id="periodo" value="<%= periodo%>" />
                <input name="co_prof" type="hidden" id="co_prof" value="<% = co_prof%>" />
                <input name="max" type="hidden" id="max" value="<% =max%>" />
                <input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>" />
                <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>" />
              </div></td>
            </tr>
          </table>
</form>
    </td>
      </tr>
      <%	
Set RS = Nothing
Set RS2 = Nothing
Set RS3 = Nothing

%>
      <tr>      
    <td height="40" valign="top"> <img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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