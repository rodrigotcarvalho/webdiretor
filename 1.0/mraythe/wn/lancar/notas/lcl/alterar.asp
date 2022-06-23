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
acao=request.QueryString("acao")
P_DATA_AULA = request.QueryString("P_DATA_AULA")
inicial = request.QueryString("ini")

if isnull(session("origem")) or session("origem") = "" then
	origemNav=request.QueryString("ori")
	session("origem") = origemNav
else
	origemNav=session("origem") 
	inicial = "S"
end if	

voltaDireto = session("voltaDireto")
session("voltaDireto") = voltaDireto


autoriza=session("autoriza")
grupo_usuario=session("grupo_usuario") 
nvg = session("chave")
ano_letivo = session("ano_letivo")
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
	unidade= dados(1)
	curso= dados(2)
	etapa= dados(3)
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
	
elseif opt="cln" then
	dados= split(obr, "$!$")
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
	
else

	unidade= session("unidades")
	curso= session("grau")
	etapa= session("serie")
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
session("unidades")=unidade
session("grau")=curso
session("serie")=etapa
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

if P_DATA_AULA="" or isnull(P_DATA_AULA) then

else
	P_DATA_AULA=replace(P_DATA_AULA,".","/")
	'if acao<>"a" then
		V_DATA_AULA=split(P_DATA_AULA,"/")
		P_DATA_AULA = V_DATA_AULA(1)&"/"&V_DATA_AULA(0)&"/"&V_DATA_AULA(2)	
		WRK_DATA_AULA = V_DATA_AULA(0)&"/"&V_DATA_AULA(1)&"/"&V_DATA_AULA(2)	
	'end if
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

obr=co_materia&"$!$"&unidade&"$!$"&curso&"$!$"&etapa&"$!$"&turma&"$!$"&periodo&"$!$"&ano_letivo&"$!$"&co_prof&"$!$"&tb


		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"'"
		Set RS = CONg.Execute(CONEXAO)


if RS.EOF then
response.Write("<div align=center><font size=2 face=Courier New, Courier, mono  color=#990000><b>Esta turma não está disponível no momento</b></font><br")
response.Write("<font size=2 face=Courier New, Courier, mono  color=#990000><a href=javascript:window.history.go(-1)>voltar</a></font></div>")

else
coordenador = RS("CO_Cord")
end if
session("obr")=obr

session("coordenador")=coordenador
 call navegacao (CON,chave,nivel)
navega=Session("caminho")

datas_periodo = diasPeriodo(periodo)
datas_formatado = diasPeriodoFormatado(periodo,", ","DD/MM/YYYY")

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

bancoPauta = escolheBancoPauta(planilha_notas,"M",p_outro)
caminhoBancoPauta = verificaCaminhoBancoPauta(bancoPauta,"M",p_outro)

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
		SQL = "Select COUNT(DT_Aula) as QTD_AULAS from TB_Materia_Lecionada WHERE CO_Professor  = "& co_prof &" AND CO_Materia_Principal = '"& co_mat_prin &"' AND CO_Materia = '"& co_materia &"' AND NU_Unidade  = "& unidade &" AND CO_Curso  = '"& curso &"' AND CO_Etapa  = '"& etapa &"' AND CO_Turma  = '"& turma &"' AND NU_Periodo = "& periodo
		Set RSP = CONPauta.Execute(SQL)				
		
		if RSP.EOF then
			wrkQtdAulasLancadas = 0		
		else
			wrkQtdAulasLancadas = RSP("QTD_AULAS")					
		end if	
		
		if P_DATA_AULA="" then
			data_Pauta = ""	
			tx_aula	= ""	
			le_tabelas="N"	
		else

			
			Set RSP = Server.CreateObject("ADODB.Recordset")
			SQL = "Select DT_Aula,TX_Aula,TX_Obs from TB_Materia_Lecionada WHERE DT_Aula = #"&P_DATA_AULA&"# AND CO_Professor  = "& co_prof &" AND CO_Materia_Principal = '"& co_mat_prin &"' AND CO_Materia = '"& co_materia &"' AND NU_Unidade  = "& unidade &" AND CO_Curso  = '"& curso &"' AND CO_Etapa  = '"& etapa &"' AND CO_Turma  = '"& turma &"' AND NU_Periodo = "& periodo

			Set RSP = CONPauta.Execute(SQL)
			
			data_Pauta = WRK_DATA_AULA
			
			if RSP.EOF  then
				le_tabelas="N"					
			else
				tx_aula = RSP("TX_Aula")
				tx_observacao = RSP("TX_Obs") 			
				le_tabelas="S"	
			end if
		end if
%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../../../estilos.css" type="text/css">
<script type="text/javascript" src="../../../../js/jquery.min.js"></script> 
<script type="text/javascript" src="../../../../js/jquery-ui.min.js"></script> 
<link type="text/css" rel="stylesheet" href="../../../../js/jquery-ui.css" />
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
				
					limiteMensal = qtdDiasMes(mesInicialFeriado,anoInicialFeriado)
'					if mesInicialFeriado = 2 then'						if anoInicialFeriado mod 4 =0 then
'							limiteMensal = 29	
'						else
'							limiteMensal = 28	
'						end if						
'					elseif mesInicialFeriado = 4 or mesInicialFeriado = 6  or mesInicialFeriado = 9 or mesInicialFeriado = 11 then
'						limiteMensal = 30					
'					else
'						limiteMensal = 31					
'					end if'
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
		$(document).ready(function()
		{	
						
			$("#enviar").click(function()
			{
				
	   if (document.myform.tx_aula.value == "")
	    {    alert("Por favor digite um conteúdo para a aula!")
    		document.myform.tx_aula.focus() 	
		} else {
				
				$("#myform").submit();
		}
			});
//			$("#submit2").click(function()
//			{
//				$("form[name='myForm']").submit(); 
//			});
//			$("#submit3").click(function()
//			{
//				$("form:first").submit();
//		 
//			});
//		 
//			$("#submit4").click(function()
//			{
//				$("#testForm").submit(function()
//				{
//				 alert('Form is submitting');
//				 return true;
//				});     
//				$("#testForm").submit(); //invoke form submission
//		 
			});
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
            altField: '#dataLancamentoForm',
            showOn: "focus",
            dateFormat: "dd/mm/yy",
            firstDay: 1,
            changeFirstDay: false,
			dayNamesMin: [ "Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sab" ],			
			monthNames: [ "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro" ],
			minDate: new Date(<%response.Write(anoInicial)%>, <%response.Write(mesInicial)%> -1, <%response.Write(diaInicial)%>),
			maxDate: new Date(<%response.Write(anoFinal)%>, <%response.Write(mesFinal)%> -1, <%response.Write(diaFinal)%>),
			defaultDate: new Date(<%response.Write(anoInicial)%>, <%response.Write(mesInicial)%> -1, <%response.Write(diaInicial)%>)

	   });
						 
  });
  
$(document).ready(function () {

     var myform = $('#myform'),
		  iter=2 


     $('#btnAddCol').change(function () {
			geraColunas('N');		 
//     	   var qtdAulasLancadas = $('#qtdAulasForm').val();	
//		   qtdAulasLancadas=qtdAulasLancadas*1; 	 
//           iter = qtdAulasLancadas+1;		   
//		   	 
//		   var totalCols = $('#btnAddCol').val();
//			$('#qtdAulasForm').val(totalCols);	   
//
//		   if (iter>2 && iter>totalCols) {
//			   var colMax = totalCols++;
//				for  (var e = iter; e > totalCols; e--){				
//					$("#blacklistgrid td:last-child").remove();
//				}
//			   iter = 2;			   
//		   } 
//		   
//		   totalCols--;		
//
//		  //i=1   
//		   for  (var i = qtdAulasLancadas ; i <= totalCols; i++){
//			   myform.find('tr').each(function(){
//			   var trow = $(this);
//				if($("tr", $(this).closest("table")).index(this) == 0){
//					 trow.append('<td width="10%" align="center">Frequ&ecirc;ncia</td>');		   			 
//				} else if($("tr", $(this).closest("table")).index(this) == 1){
//					 		trow.append('<td width="10%" align="center">Aula'+iter+'</td>');
//				 		}else{
//					 	trow.append('<td width="10%" align="center"><input type="checkbox" name="check_'+iter+'_'+$("tr", $(this).closest("table")).index(this)+'" Value="S"/></td>');
//				 		}
//			 	});
//			 iter += 1;
//		   }
     });
 });  
 
function verificaExistencia(){
	return true;
     }
  
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
function checkSubmit(e)
{
   if(e && e.keyCode == 13)
   {
	   if (document.myform.tx_aula.value == "")
	    {    alert("Por favor digite um conteúdo para a aula!")
    		document.myform.tx_aula.focus() 
		} else {
     	 document.forms[3].submit();
	}
   }
    
 
   
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
 <div onKeyPress="return checkSubmit(event)"/>
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
    <td height="10" valign="top" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> </td>
  </tr>
            <%if opt = "ok1" then%>
            <tr>         
    <td height="10" valign="top"> 
      <%
		call mensagens(4,665,2,dados)		
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
	 	 call mensagens(nivel,672,1,0)		  
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
	 	 	call mensagens_escolas(ambiente_escola,nivel,402,"inf",0,0,0)			  

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
          <td colspan="4" align="center" class="form_dado_texto">Total de datas Lan&ccedil;adas: 
            <%response.Write(wrkQtdAulasLancadas)%></td>
          <td width="190" align="center" class="form_dado_texto">&nbsp;</td>
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
        
    <td valign="top">
    <table width="100%" border="0" cellpadding="0" cellspacing="0">        
<form name="myform" method="post" id="myform" action="bd.asp?opt=i"><tr><td>       
    <table width="100%" border="0" cellpadding="0" cellspacing="0">        <tr>
          <td align="center"><table width="50%" border="0" cellspacing="0" cellpadding="0">
            <tr class="tb_tit">
              <td width="25%" align="center">Data da Aula:</td>
              </tr>
            <tr>
              <td align="center"><input name="dataLancamento" type="text" id="datepicker" size="12" value="<%response.Write(data_Pauta)%>" readonly align="middle" onBlur="javascript:verificaExistencia();"></td>
              </tr>
          </table>          </td>
          </tr>
        <tr>
          <td colspan="4" align="center">&nbsp;</td>
          </tr>
</table>
      </td></tr><tr><td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" id="blacklistgrid">
        <tr  class="tb_tit">
          <td width="700" align="center">Conte&uacute;do</td>         
          </tr>
        <tr class="form_dado_texto">
          <td width="700" align="center">
            <input name="obr" type="hidden" id="obr" value="<%response.Write(obr)%>" />                    
            <textarea name="tx_aula" id="tx_aula" cols="180" rows="15"><%response.Write(tx_aula)%></textarea></td>
         
          </tr>
        <tr class="form_dado_texto">
          <td align="center">&nbsp;</td>
        </tr>
        <tr  class="tb_tit">
          <td align="center">Observa&ccedil;&atilde;o</td>
        </tr>
        <tr class="form_dado_texto">
          <td align="center"><label for="tx_observacao"></label>
            <textarea name="tx_observacao" id="tx_observacao" cols="180" rows="5"><%response.Write(tx_observacao)%></textarea></td>
        </tr>

      </table>
      </td></tr><tr>      
    <td height="40" valign="top"> <hr>
</td>
      </tr><tr>      
    <td height="40" valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="33%" align="center">
    <%
    if inicial<>"S"  then
	
			url= "notas.asp?d="&co_materia&"&pr="&co_prof&"&p="&periodo
		else
			if isnull(origemNav) or origemNav = "" then
				url= "index.asp?nvg="&nvg
			else
				func = split(origemNav,"-")
			    if func(0) = "WA" then
					url= "../../../../wa/professor/cna/mpe/notas.asp?d="&co_materia&"&pr="&co_prof&"&exb=5&p="&periodo
				else
					url= "../../../../wn/lancar/notas/lpe/notas.asp?d="&co_materia&"&pr="&co_prof&"&exb=5&p="&periodo
				end if	
			end if	
		end if%>
    <input name="voltar" type="button" class="botao_cancelar" id="voltar" value="Voltar" onClick="MM_goToURL('parent','<%response.Write(url)%>');return document.MM_returnValue" ></td>
    <td width="33%" align="center">&nbsp;</td>
    <td width="33%" align="center">
        <%IF (autoriza=5 OR co_usr=coordenador) AND trava<>"s" AND ((periodo = 1 and ST_Per_1="x") OR (periodo = 2 and ST_Per_2="x") OR (periodo = 3 and ST_Per_3="x") OR (periodo = 4 and ST_Per_4="x") OR (periodo = 5 and ST_Per_5="x") OR (periodo = 6 and ST_Per_6="x")) then
    else
%>
        <input name="enviar" type="button" class="botao_prosseguir" id="enviar" value="Salvar">
        <%end if%>
    </td>
  </tr>
</table>

</td>
      </tr></form></table>
    </td>
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