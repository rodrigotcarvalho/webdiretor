<%'On Error Resume Next%>
<!--#include file="../../../../../global/mensagens.asp" -->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/parametros.asp"-->
<!--#include file="../../../../inc/utils.asp"-->
<!--#include file="../../../../inc/bd_parametros.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->



<%
opt=request.QueryString("opt")
obr = request.QueryString("obr")
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
			
		Set CON_AL = Server.CreateObject("ADODB.Connection") 
		ABRIR_AL = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_AL.Open ABRIR_AL		

		Set CONg = Server.CreateObject("ADODB.Connection") 
		ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONg.Open ABRIRg		
		
		Set CON0= Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		


if opt = "err6" then

	dados= split(obr, "$!$" )
	co_materia = dados(0)
	unidades= dados(1)
	grau= dados(2)
	serie= dados(3)
	turma= dados(4)
	periodo = dados(5)
	ano_letivo = dados(6)
	co_prof = dados(7)
	co_usr = session("co_usr")
	
	hp= request.QueryString("hp")
	alt= split(hp, "_" )
	errante=alt(1)
	qerrou= alt(2)
	errou = alt(3)
	
	valido="n"

	if errante=0 then
	
	else
	
	num_erro= split(errou, "$" )
	campo_errado=num_erro(0)
	
	
			Set RSer  = Server.CreateObject("ADODB.Recordset")
			SQL_er  = "Select * from TB_Matriculas WHERE CO_Matricula = "& errante
			Set RSer  = CON_AL.Execute(SQL_er )
			
	num_chamada_erro = RSer("NU_Chamada")
	local_form=campo_errado&"_"&num_chamada_erro
	javascript="onLoad='nota."&local_form&".focus();'"
	end if

elseif opt="ok" or  opt= "vt" then
	dados= split(obr, "$!$")
	co_materia = dados(0)
	unidades= dados(1)
	grau= dados(2)
	serie= dados(3)
	turma= dados(4)
	periodo = dados(5)
	ano_letivo = dados(6)
	co_prof = dados(7)
	co_usr = session("co_usr")
	
	errante=0
	valido="s"
	javascript=""
	
elseif opt="cln" then
	dados= split(obr, "$!$")
	co_materia = dados(0)
	unidades= dados(1)
	grau= dados(2)
	serie= dados(3)
	turma= dados(4)
	periodo = dados(5)
	ano_letivo = dados(6)
	co_prof = dados(7)
	co_usr = session("co_usr")
	
	errante=0
	valido="s"
	javascript=""
	
else

	unidades= session("unidades")
	grau= session("grau")
	serie= session("serie")
	turma= session("turma")
	co_materia = session("co_materia")
	periodo = session("periodo")
	co_prof = session("co_prof")
	co_usr = session("co_usr")
	tb = session("nota")	
	errante=0
	valido="s"
	javascript=""
end if

session("co_materia")=co_materia
session("unidades")=unidades
session("grau")=grau
session("serie")=serie
session("turma")=turma
session("periodo")=periodo
session("co_prof") = co_prof 
session("nota") = tb

if tb ="TB_NOTA_A" then
	CAMINHOn = CAMINHO_na
else
	if tb="TB_NOTA_B" then
		CAMINHOn = CAMINHO_nb
	else
		if tb ="TB_NOTA_C" then
			CAMINHOn = CAMINHO_nc
		else
			response.Write("ERRO")
		End if
	end if
end if


		Set RSper = Server.CreateObject("ADODB.Recordset")
		SQLper = "SELECT * FROM TB_Periodo where NU_Periodo= "&periodo
		RSper.Open SQLper, CON0

NO_Periodo= RSper("NO_Periodo")
dataInicio = RSper("DA_Inicio_Periodo")
dataFim = RSper("DA_Fim_Periodo")

vetorInicioPeriodo = split(dataInicio,"/")
diaInicial = vetorInicioPeriodo(0)
mesInicial = vetorInicioPeriodo(1)
anoInicial = vetorInicioPeriodo(2)

vetorFimPeriodo = split(dataFim,"/")
diaFinal = vetorFimPeriodo(0)
mesFinal = vetorFimPeriodo(1)
anoFinal = vetorFimPeriodo(2)

if isnull(dataInicio) or dataInicio="" then

else
	dataInicio = formata(dataInicio,"DD/MM/YYYY")
end if

if isnull(dataFim) or dataFim="" then

else
	dataFim = formata(dataFim,"DD/MM/YYYY")
end if



		Set CON_N = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIR3

obr=co_materia&"$!$"&unidades&"$!$"&grau&"$!$"&serie&"$!$"&turma&"$!$"&periodo&"$!$"&ano_letivo&"$!$"&co_prof


		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidades &" AND CO_Curso = '"& grau &"' AND CO_Etapa = '"& serie &"' AND CO_Turma = '"& turma &"'"
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

datas_periodo = diasPeriodo(periodo)
datas_formatado = diasPeriodoFormatado(periodo,", ","DD/MM/YYYY")
%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../../../estilos.css" type="text/css">
<script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jquery/1.3.2/jquery.min.js"></script> 
        <script type="text/javascript" src="http://ajax.googleapis.com/ajax/libs/jqueryui/1.7.1/jquery-ui.min.js"></script> 
        <link type="text/css" rel="stylesheet" href="http://ajax.googleapis.com/ajax/libs/jqueryui/1.7.1/themes/base/jquery-ui.css" />
<script>
       $(function() {
var dateMin = new Date();
        var weekDays = AddWeekDays(3);

        dateMin.setDate(dateMin.getDate() + weekDays);
		
        var natDays = [	
          [1, 1, 'uk'],			
<%
		Set RSF = Server.CreateObject("ADODB.Recordset")
		CONEXAOF = "Select * from TB_Feriados"
		Set RSF = CON0.Execute(CONEXAOF)

		while not RSF.EOF
			inicioFeriado = RSF("DA_Inicio")
			fimFeriado = RSF("DA_Termino")	
			vetorInicioFeriado= split(inicioFeriado,"/")
			diaInicialFeriado = vetorInicioFeriado(0)
			mesInicialFeriado = vetorInicioFeriado(1)
			anoInicialFeriado = vetorInicioFeriado(2)
			
			vetorFimFeriado = split(fimFeriado,"/")
			diaFinalFeriado = vetorFimFeriado(0)
			mesFinalFeriado = vetorFimFeriado(1)
			anoFinalFeriado = vetorFimFeriado(2)			
			if inicioFeriado=fimFeriado then
				response.Write("["&mesInicialFeriado&", "&diaInicialFeriado&", 'uk'],")
			else
				if mesInicialFeriado = mesFinalFeriado then	
					for dias = diaInicialFeriado to diaFinalFeriado
						response.Write("["&mesInicialFeriado&", "&dias&", 'uk'],")					
					next
				else
					if mesInicialFeriado = 2 then
						if anoInicialFeriado mod 4 =0 then
							limiteMensal = 29	
						else
							limiteMensal = 28	
						end if						
					elseif mesInicialFeriado = 4 or mesInicialFeriado = 6  or mesInicialFeriado = 9 or mesInicialFeriado = 11 then
						limiteMensal = 30					
					else
						limiteMensal = 31					
					end if
					for dias = diaInicialFeriado to limiteMensal
						response.Write("["&mesInicialFeriado&", "&dias&", 'uk'],")										
					next
					for dias = 1 to diaFinalFeriado
						response.Write("["&mesFinalFeriado&", "&dias&", 'uk'],")																			
					next				
				end if			
			end if		
		
		RSF.MOVENEXT
		WEND
		%>

          [12, 25, 'uk']
        ];

        function noWeekendsOrHolidays(date) {
            var noWeekend = $.datepicker.noWeekends(date);
            if (noWeekend[0]) {
                return nationalDays(date);
            } else {
                return noWeekend;
            }
        }
        function nationalDays(date) {
            for (i = 0; i < natDays.length; i++) {
                if (date.getMonth() == natDays[i][0] - 1 && date.getDate() == natDays[i][1]) {
                    return [false, natDays[i][2] + '_day'];
                }
            }
            return [true, ''];
        }
        function AddWeekDays(weekDaysToAdd) {
            var daysToAdd = 0
            var mydate = new Date()
            var day = mydate.getDay()
            weekDaysToAdd = weekDaysToAdd - (5 - day)
            if ((5 - day) < weekDaysToAdd || weekDaysToAdd == 1) {
                daysToAdd = (5 - day) + 2 + daysToAdd
            } else { // (5-day) >= weekDaysToAdd
                daysToAdd = (5 - day) + daysToAdd
            }
            while (weekDaysToAdd != 0) {
                var week = weekDaysToAdd - 5
                if (week > 0) {
                    daysToAdd = 7 + daysToAdd
                    weekDaysToAdd = weekDaysToAdd - 5
                } else { // week < 0
                    daysToAdd = (5 + week) + daysToAdd
                    weekDaysToAdd = weekDaysToAdd - (5 + week)
                }
            }

            return daysToAdd;
        }

    $( "#datepicker" ).datepicker(
        {
            inline: true,
            beforeShowDay: noWeekendsOrHolidays,
            //altField: '#dataLancamento',
            showOn: "focus",
            dateFormat: "dd/mm/yy",
            firstDay: 1,
            changeFirstDay: false,
			dayNamesMin: [ "Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sab" ],			
			monthNames: [ "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro" ],
			minDate: new Date(<%response.Write(anoInicial)%>, <%response.Write(mesInicial)%> -1, <%response.Write(diaInicial)%>),
			maxDate: new Date(<%response.Write(anoFinal)%>, <%response.Write(mesFinal)%> -1, <%response.Write(diaFinal)%>)
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
  
</script> 
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" background="../../../../img/fundo.gif" marginheight="0" <%response.Write(javascript)%>>
<%IF imprime="1"then
else
 call cabecalho (nivel) 
 end if%>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
    <td height="10" valign="top" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> </td>
  </tr>
    <%

	

call GeraNomes(co_materia,unidades,grau,serie,CON0)

no_materia= session("no_materia")
no_unidades= session("no_unidades")
no_grau= session("no_grau")
no_serie= session("no_serie")


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
		CONEXAO = "Select * from TB_Da_Aula WHERE CO_Professor= "& co_prof &"AND NU_Unidade = "& unidades &" AND CO_Curso = '"& grau &"' AND CO_Etapa = '"& serie &"' AND CO_Turma = '"& turma &"' AND CO_Materia_Principal = '"& co_materia &"'"
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
		
		Set RSP = Server.CreateObject("ADODB.Recordset")
		SQL = "Select * from TB_Pauta_Disciplina WHERE CO_Materia_Principal = '"& co_mat_prin &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo = "& periodo
		Set RSP = CONPauta.Execute(SQL)
		
		if RSP.EOF then
			wrkQtdAulasLancadas = 0
		else
			wrkQtdAulasLancadas = 0	
			for dts=1 to 80
				if dts<10 then
					wrkCampo = "DT_0"&dts
				else
					wrkCampo = "DT_"&dts				
				end if	
				wrkDataLancada = RSP(wrkCampo)		
			 	if wrkDataLancada ="" or isnull(wrkDataLancada) then
				else
					wrkQtdAulasLancadas=wrkQtdAulasLancadas+1
				end if
			 next
		end if

%>
            <%if opt = "ok" then%>
            <tr>         
    <td height="10" valign="top"> 
      <%
		call mensagens_escolas(ambiente_escola,nivel,622,"ok",0,0,0)		
%>
      <div align="center"></div></td>
            </tr>			
            <%elseif opt= "err6" then %>
            <tr> 
    <td height="10" valign="top"> 
      <%
	call mensagens_escolas(ambiente_escola,nivel,1000,"err",num_chamada_erro,errou,0)
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
      <%
	 	 	call mensagens_escolas(ambiente_escola,nivel,645,"inf",0,0,0)			  

%>
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
              <%response.Write(no_unidades)%>
              </font></div></td>
          <td width="145"> 
            <div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%response.Write(no_grau)%>
              </font></div></td>
          <td width="145"> 
            <div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%
response.Write(no_serie)%>
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
          <td colspan="4" align="center" class="form_dado_texto">Total de Aulas Lan&ccedil;adas: <%response.Write(wrkQtdAulasLancadas)%></td>
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
        
    <td valign="top"><form name="form1" method="post" action="">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td colspan="4" align="center"><table width="50%" border="0" cellspacing="0" cellpadding="0">
            <tr class="tb_tit">
              <td width="25%" align="center">Data da Aula:</td>
              <td width="25%" align="center">N&ordm; de Aulas</td>
              </tr>
            <tr>
              <td align="center"><input name="dataLancamento" type="text" id="datepicker" size="12"></td>
              <td align="center"><select name="qtdAulas" class="select_style" id="qtdAulas">
                <option value="1" selected="selected">1</option>
                <option value="2">2</option>
                <option value="3">3</option>
                <option value="4">4</option>
                <option value="5">5</option>
                </select></td>
              </tr>
          </table>          </td>
          </tr>
        <tr class="tb_tit">
          <td colspan="4" align="center"><hr></td>
          </tr>
        <tr class="tb_tit">
          <td width="100" rowspan="2" align="center">N&ordm;</td>
          <td width="700" rowspan="2" align="left">Nome</td>
          <td align="center">Frequ&ecirc;ncia</td>
          <td align="center">&nbsp;</td>
        </tr>
        <tr class="tb_tit">
          <td width="100" align="center"><label for="qtdAulas">Aula 1</label></td>
          <td width="100" align="center">&nbsp;</td>
        </tr>
      <%
check = 2
nu_chamada_ckq = 0

Set RS = Server.CreateObject("ADODB.Recordset")
SQL_A = "Select * from TB_Matriculas WHERE NU_Ano="&ano_letivo&" AND NU_Unidade = "& unidades &" AND CO_Curso = '"& grau &"' AND CO_Etapa = '"& serie &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
Set RS = CON_AL.Execute(SQL_A)


While Not RS.EOF
nu_matricula = RS("CO_Matricula")
session("matricula")=nu_matricula
nu_chamada = RS("NU_Chamada")

errante=errante*1
nu_matricula=nu_matricula*1	


	if subopcao="imp" Then
		classe = "tabela"
		classe_td_imp= " class = 'tabela'"
	elseif nu_matricula = errante then
		classe = "tb_fundo_linha_erro"
		onblur="mudar_cor_blur_erro"	
		classe_td_imp= ""	  	   
	else
		if check mod 2 =0 then
			classe = "tb_fundo_linha_par" 
			onblur="mudar_cor_blur_par"
		else 
			classe ="tb_fundo_linha_impar"
			onblur="mudar_cor_blur_impar"
		end if 
		classe_td_imp= ""		
	end if

	Set RSs = Server.CreateObject("ADODB.Recordset")
	SQL_s ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& nu_matricula&" and TB_Matriculas.NU_Ano="&ano_letivo
	Set RSs = CON_AL.Execute(SQL_s)

	if RSs.EOF then
	%>
      <tr>
        <td width="100" class="<%response.Write(classe)%>">&nbsp;</td>
        <td width="700" class="<%response.Write(classe)%>">Matrícula
          <%response.Write(nu_matricula)%>
          cadastrada em TB_Matriculas sem correspondência em TB_Alunos</td>
           <td width="100" align="center" class="<%response.Write(classe)%>">&nbsp;</td>
           <td width="100" align="center" class="<%response.Write(classe)%>">&nbsp;</td>
      </tr>
      <%else
		situac=RSs("CO_Situacao")
		nome_aluno=RSs("NO_Aluno")	
	'Verificando se algum aluno mudou de turma e inserindo uma linha cinza para o lugar do aluno
			if (nu_chamada_ckq <>nu_chamada - 1) then
				teste_nu_chamada = nu_chamada-nu_chamada_ckq
				teste_nu_chamada=teste_nu_chamada-1
				'response.write(teste_nu_chamada&"="&nu_chamada&"-"&nu_chamada_ckq)
				classe_anterior=classe
				if subopcao="imp" Then
					classe = "tabela"
				else	
					classe="tb_fundo_linha_falta"
				end if
		
				for k=1 to teste_nu_chamada 
					nu_chamada_ckq=nu_chamada_ckq+1
					nu_chamada_falta=nu_chamada_ckq
				%>
      <tr>
        <td width="100" align="center" class="<%response.Write(classe)%>"><input name="nu_chamada_<%response.Write(nu_chamada_falta)%>" class="borda_edit" type="hidden" value="<%response.Write(nu_chamada_falta)%>" />
          <%response.Write(nu_chamada_falta)%>
          <input name="nu_matricula_<%response.Write(nu_chamada_falta)%>" type="hidden" value='falta' /></td>
        <td width="700" class="<%response.Write(classe)%>">&nbsp;</td>
        <%
							width=width_else

							align="center"
							nome_campo="check_"&nu_chamada
							conteudo="&nbsp;"
					 %>
        <td width="100" align="center" class="<%response.Write(classe)%>"><div align="<%response.Write(align)%>">
          <%response.Write(conteudo)%>
        </div></td>
        <td width="100" align="center" class="<%response.Write(classe)%>">&nbsp;</td>
      </tr>
      <%				
					next
	'Inserindo o aluno seguinte aos que mudaram de turma
					nu_chamada_ckq=nu_chamada_ckq+1		
					if situac<>"C" then
						if subopcao="imp" Then
							classe = "tabela"
						else
							classe="tb_fundo_linha_falta"
						end if	
							valor="falta"
							nome_aluno=nome_aluno&" - Aluno Inativo"
					end if			
					%>
      <tr class="<%response.Write(classe_anterior)%>" id="<%response.Write("celula"&nu_chamada)%>">
        <td width="100" align="center" <%response.Write(classe_td_imp)%>><input name="nu_chamada_<%response.Write(nu_chamada)%>" class="borda_edit" type="hidden" value="<%response.Write(nu_chamada)%>" />
          <%response.Write(nu_chamada)%>
          <input name="nu_matricula_<%response.Write(nu_chamada)%>" type="hidden" value='<%response.Write(nu_matricula)%>' /></td>
        <td width="700" <%response.Write(classe_td_imp)%>><%response.Write(nome_aluno)%></td>
        <% 
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				Set RS3 = CON_N.Execute(SQL_N)			 
				coluna=0	 

					width=width_else
					align="center"
					nome_campo="check_"&nu_chamada
				
					if RS3.EOF then 
						valor=""
					else
						errante=errante*1
						nu_matricula=nu_matricula*1																																	

					end if	
					'conteudo=n
			 %>
        <td width="100" align="center" <%response.Write(classe_td_imp)%>><div align="<%response.Write(align)%>">
          <input name="<%response.Write(nome_campo)%>" type="checkbox" value="S">
        </div></td>
        <td width="100" align="center" <%response.Write(classe_td_imp)%>>&nbsp;</td>
      </tr>
      <%

	'Se os números de chamada estiverem completos. Se não faltar aluno na turma.
			ELSE	
					if situac<>"C" then
						if subopcao="imp" Then
							classe = "tabela"
						else
							classe="tb_fundo_linha_falta"
						end if	
							valor="falta"
							nome_aluno=nome_aluno&" - Aluno Inativo"
					end if			
					nu_chamada_ckq=nu_chamada_ckq+1
					%>
      <tr class="<%response.Write(classe)%>" id="<%response.Write("celula"&nu_chamada)%>">
        <td width="100" align="center" <%response.Write(classe_td_imp)%>><input name="nu_chamada_<%response.Write(nu_chamada)%>" class="borda_edit" type="hidden" value="<%response.Write(nu_chamada)%>" />
          <%response.Write(nu_chamada)%>
          <input name="nu_matricula_<%response.Write(nu_chamada)%>" type="hidden" value='<%response.Write(nu_matricula)%>' /></td>
        <td width="700" <%response.Write(classe_td_imp)%>><%response.Write(nome_aluno)%></td>
        <% 
				Set RS3 = Server.CreateObject("ADODB.Recordset")
				SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
				Set RS3 = CON_N.Execute(SQL_N)			 
				coluna=0	 

					width=width_else
					align="center"
					nome_campo="check_"&nu_chamada

			 %>
        <td width="100" align="center" <%response.Write(classe_td_imp)%>><div align="<%response.Write(align)%>">
          <input name="<%response.Write(nome_campo)%>" type="checkbox" value="S">
        </div></td>
        <td width="100" align="center" <%response.Write(classe_td_imp)%>>&nbsp;</td>

      </tr>
      <%			
			END IF			              
		if situac<>"C" then
			linha_tabela=linha_tabela
		else
		
			linha_tabela=linha_tabela+1
		end if
 	
	END IF	
max=nu_chamada
	check = check+1 
RS.MoveNext
Wend 
session("max")=max

%>
      </table>
    </form></td>
      </tr>
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