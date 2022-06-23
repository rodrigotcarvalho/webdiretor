<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->

<%
opt = request.QueryString("opt")

ano_letivo = session("ano_letivo")
co_usr = session("co_user")
nivel=4

nvg = request.QueryString("nvg")
session("nvg")=nvg

nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 





ano_info=nivel&"-"&nvg&"-"&ano_letivo

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		
		
		Set conw = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		conw.Open ABRIR

		Set RSRA = Server.CreateObject("ADODB.Recordset")
		SQLRA = "SELECT * FROM TB_Ano_Letivo WHERE NU_Ano_Letivo = '"&session("ano_letivo")&"'"
        RSRA.Open SQLRA, conw
		
		data_hoje = RSRA("DT_Inicio_Rematricula")
		data_fim_ano = RSRA("DT_Final_Rematricula")
		data_bloqueto = RSRA("DT_Bloqueto_Rematricula")

if data_hoje = "" or isnull(data_hoje) then
	data_hoje = Date
end if

if data_fim_ano = "" or isnull(data_fim_ano) then
	data_fim_ano = "31/12/"&ano
end if	

if data_bloqueto = "" or isnull(data_bloqueto) then
	data_bloqueto = data_fim_ano
end if

anoInicial = ano
anoFinal = anoInicial+1
mesInicial = mes
mesFinal = 12
diaInicial = dia
diaFinal = 31



 call navegacao (CON,nvg,nivel)
navega=Session("caminho")	

%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="file:../../../../img/mm_menu.js"></script>
<script type="text/javascript" src="../../../../js/global.js"></script>
<script type="text/javascript" src="../../../../js/jquery.min.js"></script> 
<script type="text/javascript" src="../../../../js/jquery-ui.min.js"></script> 
<link type="text/css" rel="stylesheet" href="../../../../js/jquery-ui.css" />
<script language="JavaScript" type="text/JavaScript">
<!--
       $(function() {
		var dateMin = new Date();
        var weekDays = AddWeekDays(3);

        dateMin.setDate(dateMin.getDate() + weekDays);
		
        var natDays = [	
          [1, 1, 'uk'],			
<%
'		Set RSF = Server.CreateObject("ADODB.Recordset")
'		CONEXAOF = "Select * from TB_Feriados"
'		Set RSF = CON0.Execute(CONEXAOF)
'
'		while not RSF.EOF
'			inicioFeriado = RSF("DA_Inicio")
'			fimFeriado = RSF("DA_Termino")	
'			vetorInicioFeriado= split(inicioFeriado,"/")
'			diaInicialFeriado = vetorInicioFeriado(0)
'			mesInicialFeriado = vetorInicioFeriado(1)
'			anoInicialFeriado = vetorInicioFeriado(2)
'			
'			vetorFimFeriado = split(fimFeriado,"/")
'			diaFinalFeriado = vetorFimFeriado(0)
'			mesFinalFeriado = vetorFimFeriado(1)
'			anoFinalFeriado = vetorFimFeriado(2)			
'			if inicioFeriado=fimFeriado then
'				response.Write("["&mesInicialFeriado&", "&diaInicialFeriado&", 'uk'],")
'			else
'				if mesInicialFeriado = mesFinalFeriado then	
'					for dias = diaInicialFeriado to diaFinalFeriado
'						response.Write("["&mesInicialFeriado&", "&dias&", 'uk'],")					
'					next
'				else
'				
'					limiteMensal = qtdDiasMes(mesInicialFeriado,anoInicialFeriado)
'					'if mesInicialFeriado = 2 then
''						if anoInicialFeriado mod 4 =0 then
''							limiteMensal = 29	
''						else
''							limiteMensal = 28	
''						end if						
''					elseif mesInicialFeriado = 4 or mesInicialFeriado = 6  or mesInicialFeriado = 9 or mesInicialFeriado = 11 then
''						limiteMensal = 30					
''					else
''						limiteMensal = 31					
''					end if
'					for dias = diaInicialFeriado to limiteMensal
'						response.Write("["&mesInicialFeriado&", "&dias&", 'uk'],")										
'					next
'					for dias = 1 to diaFinalFeriado
'						response.Write("["&mesFinalFeriado&", "&dias&", 'uk'],")																			
'					next				
'				end if			
'			end if		
'		
'		RSF.MOVENEXT
'		WEND
'		%>

          [12, 25, 'uk']
        ];
		$(document).ready(function()
		{	
						
			$("#enviar").click(function()
			{
				
				$("#myform").submit();
		 
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

    $( "[readonly]" ).datepicker(
        {
            inline: true,
            //beforeShowDay: noWeekendsOrHolidays,
            altField: '#dataLancamentoForm',
            showOn: "focus",
            dateFormat: "dd/mm/yy",
            firstDay: 1,
            changeFirstDay: false,
			dayNamesMin: [ "Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sab" ],			
			monthNames: [ "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Juhol", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro" ],
			minDate: new Date(<%response.Write(anoInicial)%>, <%response.Write(mesInicial)%> -1, <%response.Write(diaInicial)%>),
			maxDate: new Date(<%response.Write(anoFinal)%>, <%response.Write(mesFinal)%> -1, <%response.Write(diaFinal)%>),
			defaultDate: new Date(<%response.Write(anoInicial)%>, <%response.Write(mesInicial)%> -1, <%response.Write(diaInicial)%>)

	   });
						 
  });
  
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
//-->
</script>
                         <script>
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
						
						
						 function recuperarCurso(uTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=c", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
document.all.divEtapa.innerHTML ="<select name='etapa' class='select_style' id='etapa'><option value='999990' selected>           </option></select>"
document.all.divTurma.innerHTML = "<select name='turma' class='select_style' id='turma'><option value='999990' selected>           </option></select>"
document.all.divPeriodo.innerHTML = "<select name='periodo' class='select_style' id='periodo'><option value='0' selected>           </option></select>"
//recuperarEtapa()
                                                           }
                                               }

                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=e", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select name='turma' class='select_style' id='turma'><option value='999990' selected>           </option></select>"
document.all.divPeriodo.innerHTML = "<select name='periodo' class='select_style' id='periodo'><option value='0' selected>           </option></select>"
//recuperarTurma()
                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t4", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t
document.all.divPeriodo.innerHTML = "<select name='periodo' class='select_style' id='periodo'><option value='0' selected>           </option></select>"
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
								   
function recuperarPeriodo(eTipo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=p1", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                       var resultado_p= oHTTPRequest.responseText;
resultado_p = resultado_p.replace(/\+/g," ")
resultado_p = unescape(resultado_p)
document.all.divPeriodo.innerHTML = resultado_p
																	   
                                                           }
                                               }
 
                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }									   

 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}

function checksubmit()
{
  if (document.busca.unidade.value == "999990" || document.busca.curso.value == "999990" || document.busca.etapa.value == "999990" || document.busca.turma.value == "999990" || document.busca.periodo.value == "0")
  { alert("É necessário preencher pelo menos Unidade, Curso, Etapa, Turma e Periodo!")
	var combo = document.getElementById("unidade");
	combo.options[0].selected = "true";
	var combo2 = document.getElementById("curso");
	combo2.options[0].selected = "true";	
	var combo3 = document.getElementById("etapa");
	combo3.options[0].selected = "true";	
	var combo4 = document.getElementById("turma");
	combo4.options[0].selected = "true";	
	var combo5 = document.getElementById("periodo");
	combo5.options[0].selected = "true";		
    return false
  }  
   
  return true
}
</script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<% call cabecalho (nivel)
	  %>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
                    
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
	  </td>
	  </tr>
      <%
if opt = "ok" then%>
  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(4,9709,2,0)
%>
    </td>
  </tr>
<%

end if
%>

                  <tr> 
                    
    <td height="10"> 
      <%	call mensagens(4,9706,0,0) 
%>
</td></tr>
<tr>

            <td valign="top"> <form name="busca" method="post" action="bd.asp" onSubmit="return checksubmit()">
                
        <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td width="653" height="15" class="tb_tit">Par&acirc;metros para Rematr&iacute;cula </td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="250" class="tb_subtit">&nbsp;</td>
                  <td width="250" align="center" class="tb_subtit">DATA DE IN&Iacute;CIO</td>
                  <td width="250" align="center" class="tb_subtit">DATA DE FIM</td>
                  <td width="250" align="center" class="tb_subtit">DATA DO BLOQUETO</td>
                  <td width="250" class="tb_subtit">&nbsp;</td>
                </tr>
                <% 'if RS1.EOF THEN %>
                <%'else%>
                <tr> 
                  <td width="250">&nbsp;</td>
                  <td width="250" align="center"><input name="dataLancamentoInicio" type="text" id="datepickerInicio" size="12" value="<%response.Write(data_hoje)%>" readonly align="middle"></td>
                  <td width="250" align="center"><input name="dataLancamentoFim" type="text" id="datepickerFim" size="12" value="<%response.Write(data_fim_ano)%>" readonly align="middle"></td>
                  <td width="250" align="center"><input name="dataLancamentoBloqueto" type="text" id="datepickerBloqueto" size="12" value="<%response.Write(data_bloqueto)%>" readonly align="middle"></td>
                  <td width="250">&nbsp;</td>
                </tr>
                <%'end if %>
                <tr> 
                  <td height="15" colspan="5" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td colspan="3"><hr></td>
                      </tr>
                    <tr>
                      <td width="33%">&nbsp;</td>
                      <td width="34%">&nbsp;</td>
                      <td width="33%"><div align="center"><font size="3" face="Courier New, Courier, mono">
                        <input type="submit" name="Submit2" value="Prosseguir" class="botao_prosseguir">
                      </font></div></td>
                    </tr>
                  </table></td>
                </tr>
              </table></td>
          </tr>
        </table>
              </form></td>
          </tr>
		  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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
'response.redirect("../../../../inc/erro.asp")
end if
%>