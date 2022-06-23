<%' Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/bd_parametros.asp"-->

<%
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
data_calc=dia&"/"&mes&"/"&ano	

opt = request.QueryString("opt")
cod= request.QueryString("cod_cons")
cod_form = cod
ano_letivo = session("ano_letivo")
co_usr = session("co_user")
nivel=4

obr=cod&"?"&ano_letivo
Session("dia_de")=""
Session("dia_de")=""
Session("dia_ate")=""
Session("mes_ate")=""
Session("unidade")=""
Session("curso")=""
Session("etapa")=""
Session("turma")=""


	nvg=session("chave")
	session("chave")=nvg
	session("nvg")=nvg
	
mes_selecionado= request.QueryString("mes")
if mes_selecionado = "" or isnull(mes_selecionado) then
	mes_parcela = session("mes_extrato")
else
	mes_parcela = mes_selecionado
	session("mes_extrato") = mes_parcela
end if	

if mes_parcela = "" or isnull(mes_parcela) then
	mes_parcela = "nulo"
end if	

if mes_parcela = "nulo" then
	sql_mes = "SELECT * FROM TB_Posicao WHERE CO_Matricula_Escola ="& cod 
else
	sql_mes = "SELECT * FROM TB_Posicao WHERE CO_Matricula_Escola ="& cod &" AND Mes = "&mes_parcela
end if
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano_info=nivel&"-"&nvg&"-"&ano_letivo



		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
    	Set CON_WF = Server.CreateObject("ADODB.Connection") 
		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_WF.Open ABRIR_WF	

	
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1

	
	
		'response.Write "DBQ="& CAMINHO_pf & ";Driver={Microsoft Access Driver (*.mdb)}"
	
		Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4 = "DBQ="& CAMINHO_pf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4		
		
		Set CON6 = Server.CreateObject("ADODB.Connection") 
		ABRIR6 = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON6.Open ABRIR6			

 call navegacao (CON,nvg,nivel)
navega=Session("caminho")	

 Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Autoriz_Usuario_Grupo Where CO_Usuario = "&co_usr
		RS2.Open SQL2, CON
		
if RS2.EOF then

else		
	co_grupo=RS2("CO_Grupo")
End if

	

	SQL2 = "select * from TB_Alunos where CO_Matricula = " & cod 
	set RS2 = CON1.Execute (SQL2)
	
	nome_aluno= RS2("NO_Aluno")
	tp_resp_fin = RS2("TP_Resp_Fin")

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
		
		end if	


	
		
		Set RSc = Server.CreateObject("ADODB.Recordset")
		SQLc = "SELECT NO_Contato,CO_CPF_PFisica, TX_EMail FROM TB_Contatos where CO_Matricula = "& cod &" AND TP_Contato = '"& tp_resp_fin&"'"
		RSc.Open SQLc, CON6		
		existe_email = "N"		
		If RSc.EOF then
			nome_resp ="Nome não cadastrado para o "&tp_resp_fin
		else
			nome_resp = RSc("NO_Contato")
			cpf_resp = RSc("CO_CPF_PFisica")
			email_resp =RSc("TX_EMail")
					
			if cpf_resp = "" or isnull(cpf_resp) then
			
			else
				cpf_resp = replace(cpf_resp,"-","")
				cpf_resp = replace(cpf_resp,".","")				
			end if
			
			if isnull(email_resp) or email_resp="" then
				email_resp ="Email não cadastrado"
			else
				existe_email = "S"				
			end if		
				
		end if	

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
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script language="JavaScript" src="file:../../../../img/mm_menu.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
<script type="text/javascript" src="../../../../js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
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
function submitforminterno()  
{
   var f=document.forms[3]; 
      f.submit(); 
	  
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
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
<%
parametros_funcao_jscript="celula"
total_periodo=1
if  total_periodo>1 then
	for b=2 to total_periodo
		parametros_funcao_jscript=parametros_funcao_jscript&",celulap"&b
	next
end if
%>
function mudar_cor_focus(<%response.Write(parametros_funcao_jscript)%>){
   celula.style.backgroundColor="#D8FF9D";
<%
for pr=1 to maior_periodo
	if pr=2 then
   		response.Write("celulap2.style.backgroundColor=""#D8FF9D"";")
	elseif pr=3 then
   		response.Write("celulap3.style.backgroundColor=""#D8FF9D"";")
	elseif pr=4 then
   		response.Write("celulap4.style.backgroundColor=""#D8FF9D"";")
	elseif pr=5 then
   		response.Write("celulap5.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=6 then
   		response.Write("celulap6.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=7 then
   		response.Write("celulap7.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=8 then
   		response.Write("celulap8.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=9 then
   		response.Write("celulap9.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=10 then
   		response.Write("celulap10.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=11 then
   		response.Write("celulap11.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=12 then
   		response.Write("celulap12.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=13 then
   		response.Write("celulap13.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=14 then
   		response.Write("celulap14.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=15 then
   		response.Write("celulap15.style.backgroundColor=""#D8FF9D"";")
	elseif pr=16 then
   		response.Write("celulap16.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=17 then
   		response.Write("celulap17.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=18 then
   		response.Write("celulap18.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=19 then
   		response.Write("celulap19.style.backgroundColor=""#D8FF9D"";")	
	elseif pr=20 then
   		response.Write("celulap20.style.backgroundColor=""#D8FF9D"";")																						
	end if
next	
%>									 
}
function mudar_cor_blur_par(<%response.Write(parametros_funcao_jscript)%>){
   celula.style.backgroundColor="#FFFFFF";
<%
for pr=1 to maior_periodo
	if pr=2 then
   		response.Write("celulap2.style.backgroundColor=""#FFFFFF"";")
	elseif pr=3 then
   		response.Write("celulap3.style.backgroundColor=""#FFFFFF"";")
	elseif pr=4 then
   		response.Write("celulap4.style.backgroundColor=""#FFFFFF"";")
	elseif pr=5 then
   		response.Write("celulap5.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=6 then
   		response.Write("celulap6.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=7 then
   		response.Write("celulap7.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=8 then
   		response.Write("celulap8.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=9 then
   		response.Write("celulap9.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=10 then
   		response.Write("celulap10.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=11 then
   		response.Write("celulap11.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=12 then
   		response.Write("celulap12.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=13 then
   		response.Write("celulap13.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=14 then
   		response.Write("celulap14.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=15 then
   		response.Write("celulap15.style.backgroundColor=""#FFFFFF"";")
	elseif pr=16 then
   		response.Write("celulap16.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=17 then
   		response.Write("celulap17.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=18 then
   		response.Write("celulap18.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=19 then
   		response.Write("celulap19.style.backgroundColor=""#FFFFFF"";")	
	elseif pr=20 then
   		response.Write("celulap20.style.backgroundColor=""#FFFFFF"";")																						
	end if
next	
%>   
} 
function mudar_cor_blur_impar(<%response.Write(parametros_funcao_jscript)%>){
   celula.style.backgroundColor="#E9E9E9";
<%
for pr=1 to maior_periodo
	if pr=2 then
   		response.Write("celulap2.style.backgroundColor=""#E9E9E9"";")
	elseif pr=3 then
   		response.Write("celulap3.style.backgroundColor=""#E9E9E9"";")
	elseif pr=4 then
   		response.Write("celulap4.style.backgroundColor=""#E9E9E9"";")
	elseif pr=5 then
   		response.Write("celulap5.style.backgroundColor=""#E9E9E9"";")	
	elseif pr=6 then
   		response.Write("celulap6.style.backgroundColor=""#E9E9E9"";")	
	elseif pr=7 then
   		response.Write("celulap7.style.backgroundColor=""#E9E9E9"";")	
	elseif pr=8 then
   		response.Write("celulap8.style.backgroundColor=""#E9E9E9"";")	
	elseif pr=9 then
   		response.Write("celulap9.style.backgroundColor=""#E9E9E9"";")	
	elseif pr=10 then
   		response.Write("celulap10.style.backgroundColor=""#E9E9E9"";")	
	elseif pr=11 then
   		response.Write("celulap11.style.backgroundColor=""#E9E9E9"";")	
	elseif pr=12 then
   		response.Write("celulap12.style.backgroundColor=""#E9E9E9"";")	
	elseif pr=13 then
   		response.Write("celulap13.style.backgroundColor=""#E9E9E9"";")	
	elseif pr=14 then
   		response.Write("celulap14.style.backgroundColor=""#E9E9E9"";")	
	elseif pr=15 then
   		response.Write("celulap15.style.backgroundColor=""#E9E9E9"";")
	elseif pr=16 then
   		response.Write("celulap16.style.backgroundColor=""#E9E9E9"";")	
	elseif pr=17 then
   		response.Write("celulap17.style.backgroundColor=""#E9E9E9"";")	
	elseif pr=18 then
   		response.Write("celulap18.style.backgroundColor=""#E9E9E9"";")	
	elseif pr=19 then
   		response.Write("celulap19.style.backgroundColor=""#E9E9E9"";")	
	elseif pr=20 then
   		response.Write("celulap20.style.backgroundColor=""#E9E9E9"";")																						
	end if
next	
%>   
}
function focar_load() 
{ 
//mudar_cor_blur_par("celula1");

}
<%
if existe = "N" then
	onload=""
else
	onload="focar_load();"
end if
%>
//-->
</script>
  <link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0" onLoad="<%response.Write(onload)%>">
<% call cabecalho (nivel)
	  %>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
                    
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
	  </td>
	  </tr>   
<%if opt = "ok" then %>      
      <tr> 
    <td height="10"> 
      <%	call mensagens(nivel,709,2,0) 
%>
</td></tr> 
<%end if %>  
      <tr>                   
    <td height="10"> 
      <%	call mensagens(nivel,636,0,mes_parcela) 
%>
</td></tr>
<tr>

            <td valign="top"> 
<%
mes = DatePart("m", now) 
dia = DatePart("d", now) 



dia=dia*1
mes=mes*1
%>				
<form action="index.asp?opt=list&nvg=<%=nvg%>" method="post" name="busca" id="busca">
        <table width="100%" border="0" align="center" cellspacing="0" class="tb_corpo"
>
          <tr> 
            <td width="684" class="tb_tit"
>Dados Escolares</td>
            <td width="140" class="tb_tit"
> </td>
          </tr>
          <tr> 
            <td height="10"> <table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="19%" height="10"> <div align="right"><font class="form_dado_texto"> 
                      <strong>Matr&iacute;cula:</strong> </font></div></td>
                  <td width="9%" height="10" ><font class="form_dado_texto"> 
                    <input name="busca1" type="hidden" id="busca1" value="<% response.Write(cod_form)%>">
					<input name="busca2" type="hidden" class="textInput" id="busca2"  value="" size="75" maxlength="50">					
                    <%response.Write(cod)%>
                    </font></td>
                  <td width="6%" height="10"> <div align="right"><font class="form_dado_texto"> 
                      <strong>Nome: </strong></font></div></td>
                  <td width="33%" height="10"><font class="form_dado_texto"> 
                  	<%response.Write(nome_aluno)%>
                  	<input name="nome" type="hidden" class="textInput" id="nome"  value="<%response.Write(nome_aluno)%>" size="75" maxlength="50">
                  	&nbsp;</font></td>
                  <td ><div align="right"><font class="form_dado_texto"> M&ecirc;s: </font></div></td>
                  <td ><div align="center">
                  	<select name="mes" class="select_style">
					<% if mes_parcela = "nulo" then
							select_nulo="selected"
					   else
							select_nulo=""	
							mes_parcela= mes_parcela*1
							if mes_parcela = 1 then
								select_1="selected"
							else	
								select_1=""	
								if mes_parcela = 2 then
									select_2="selected"
								else	
									select_2=""		
									if mes_parcela = 3 then
										select_3="selected"
									else	
										select_3=""		
									end if		
										if mes_parcela = 4 then
											select_4="selected"
										else	
											select_4=""	
											if mes_parcela = 5 then
												select_5="selected"
											else	
												select_5=""	
												if mes_parcela = 6 then
													select_6="selected"
												else	
													select_6=""		
													if mes_parcela = 7 then
														select_7="selected"
													else	
														select_7=""		
														if mes_parcela = 8 then
															select_8="selected"
														else	
															select_8=""	
															if mes_parcela = 9 then
																select_9="selected"
															else	
																select_9=""		
																if mes_parcela = 10 then
																	select_10="selected"
																else	
																	select_10=""	
																	if mes_parcela = 11 then
																		select_11="selected"
																	else	
																		select_11=""	
																		if mes_parcela = 12 then
																			select_12="selected"
																		else	
																			select_12=""		
																		end if																				
																	end if																			
																end if																		
															end if																		
														end if															
													end if																											
												end if														
											end if														
										end if																		
								end if	 								
						end if	  
					 end if
						%>   	
                  		<option value="nulo" <%response.Write(select_nulo)%>>Todos</option>
                  		<option value="1" <%response.Write(select_1)%>>Janeiro</option>
                  		<option value="2" <%response.Write(select_2)%>>Fevereiro</option>
                  		<option value="3" <%response.Write(select_3)%>>Mar&ccedil;o</option>
                  		<option value="4" <%response.Write(select_4)%>>Abril</option>
                  		<option value="5" <%response.Write(select_5)%>>Maio</option>
                  		<option value="6" <%response.Write(select_6)%>>Junho</option>
                  		<option value="7" <%response.Write(select_7)%>>Julho</option>
                  		<option value="8" <%response.Write(select_8)%>>Agosto</option>
                  		<option value="9" <%response.Write(select_9)%>>Setembro</option>
                  		<option value="10" <%response.Write(select_10)%>>Outubro</option>
                  		<option value="11" <%response.Write(select_11)%>>Novembro</option>
                  		<option value="12" <%response.Write(select_12)%>>Dezembro</option>
                  		</select>
                  	</div></td>
                  </tr>
              </table></td>
            <td rowspan="2" valign="top"><div align="center"><font size="2" face="Arial, Helvetica, sans-serif">
            	<input name="Submit2" type="submit" class="botao_prosseguir" id="Submit3" value="Procurar">
            </font> </div></td>
          </tr>
          <tr> 
            <td height="10">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="2"><table width="100%" border="0" cellspacing="0">
                <tr class="tb_subtit"> 
                  <td width="34" height="10"> <div align="center"> 
                      ANO</div></td>
                  <td width="74" height="10"> <div align="center">MATR&Iacute;CULA</div></td>
                  <td width="96" height="10"> <div align="center">CANCELAMENTO</div></td>
                  <td width="83" height="10"> <div align="center"> SITUA&Ccedil;&Atilde;O</div></td>
                  <td width="81" height="10"> <div align="center">UNIDADE</div></td>
                  <td width="111" height="10"> <div align="center">CURSO</div></td>
                  <td width="63" height="10"> <div align="center"> ETAPA</div></td>
                  <td width="66" height="10"> <div align="center">TURMA</div></td>
                  <td width="60" height="10"> <div align="center">CHAMADA</div></td>
                </tr>
                <tr class="tb_corpo"
> 
                  <td width="34" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(ano_aluno)%>
                      </font></div></td>
                  <td width="74" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(rematricula)%>
                      </font></div></td>
                  <td width="96" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(encerramento)%>
                      </font></div></td>
                  <td width="83" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%
if existe = "S" then						
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
		RSCONTST.Open SQLCONTST, CON0
							
				no_situacao = RSCONTST("TX_Descricao_Situacao")	
end if					
					response.Write(no_situacao)%>
                      </font></div></td>
                  <td width="81" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_unidade)%>
                      </font></div></td>
                  <td width="111" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_curso)%>
                      </font></div></td>
                  <td width="63" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_etapa)%>
                      </font></div></td>
                  <td width="66" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(turma)%>
                      </font></div></td>
                  <td width="60" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(cham)%>
                      </font></div></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td height="10" colspan="2" class="tb_tit"
>Extrato</td>
          </tr>
          <tr> 
            <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr class="tb_subtit"> 
                  <td width="87"> 
                    <div align="center">DATA VENCIMENTO</div></td>
                  <td width="139"> 
                    <div align="center">SERVI&Ccedil;O</div></td>
                  <td width="105" align="right"> 
                    VALOR A PAGAR</td>
                  <td width="92" align="right"> MULTA</td>
                  <td width="104" align="right"> CORRE&Ccedil;&Atilde;O</td>
                  <td width="119" align="right"> VALOR CORRIGIDO</td>
                  <td width="105" align="right"> 
                    VALOR PAGO</td>
                  <td width="114"> 
                    <div align="center">DATA<br>
                  PAGAMENTO</div></td>
                  <td width="107"> 
                    <div align="center">SITUA&Ccedil;&Atilde;O</div></td>
                  <td width="26">&nbsp;</td>
                </tr>
                <%		
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4= sql_mes
		RS4.Open SQL4, CON4
		
if existe = "N" then
%>
                <tr> 
                  <td colspan="10"> <div align="center"><font class="form_dado_texto"> <br>
                      <br>
                      <br>
                      <br>
                      <br>
                      Este aluno não está ativo neste ano letivo.<br>
                      <br>
                      <br>
                      <br>
                      <br>
                      </font></div></td>
                </tr>	
		
<%
elseif RS4.EOF THEN
%>
                <tr> 
                  <td colspan="10"> <div align="center"><font class="form_dado_texto"> <br>
                      <br>
                      <br>
                      <br>
                      <br>
                      Não existem lançamentos financeiros para este aluno.<br>
                      <br>
                      <br>
                      <br>
                      <br>
                      </font></div></td>
                </tr>
                <%else
check = 1
linha = 1
compromisso_total=0
multa_total=0
mora_total=0
corrigido_total=0
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

	if realizado = 0 or isnull(realizado) then
	realizado=""
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
	dia_venc = dia_venc*1
	if dia_venc<10 then
	dia_venc="0"&dia_venc
	else
	dia_venc=dia_venc
	end if
	mes_venc = mes_venc*1
	if mes_venc<10 then
	mes_venc="0"&mes_venc
	else
	mes_venc=mes_venc
	end if
	
	da_vencimento_show=dia_venc&"/"&mes_venc&"/"&ano_venc
	p_vencimento = mes_venc
	venc=replace(da_vencimento,"/","$wxg$adn$")
	'RESPONSE.Write(data_compara&"<<")
	if isnull(da_realizado) then
		 d_diff=DateDiff("d",data_compara,da_vencimento)
		 if nome_lanc="Mensalidade" then
		 situacao="<a href=""../../../../relatorios/gera_boleto.asp?opt="&p_vencimento&"&c="&cod&""">Parcela Não Paga</a>"	  
		 'situacao="<a href=""#"" onclick = ""MM_openBrWindow('boleto_itau.asp?c="&cod&"&opt="&p_vencimento&"','','status=yes,scrollbars=yes,resizable=yes,width=800,height=500')"">Parcela Não Paga</a>"
		else
		  situacao="Parcela Não Paga"
		end if	
			
		da_realizado_show=""
	else
		real_split=split(da_realizado,"/")
		dia_real=real_split(0)
		mes_real=real_split(1)
		ano_real=real_split(2)
		real=mes_real&"/"&dia_real&"/"&ano_real
		dia_real = dia_real*1
		if dia_real<10 then
			dia_real="0"&dia_real
		else
			dia_real=dia_real
		end if
		mes_real=mes_real*1
		if mes_real<10 then
			mes_real="0"&mes_real
		else
			mes_real=mes_real
		end if
		
		da_realizado_show=dia_real&"/"&mes_real&"/"&ano_real
		
		d_diff=DateDiff("d",da_realizado,da_vencimento)
	
	end if
	'response.Write(da_vencimento&" - "&da_realizado&" = "&d_diff&"<BR>")
	if isnull(da_realizado) and d_diff<0 then
	  cor = "tb_fundo_linha_atraso" 
	  if nome_lanc="Mensalidade" then
	      situacao="<a href=""../../../../relatorios/gera_boleto.asp?opt="&p_vencimento&"&c="&cod&""">Parcela Vencida</a>"	  
		  'situacao="<a href=""#"" onclick = ""MM_openBrWindow('boleto_itau.asp?c="&cod&"&opt="&p_vencimento&"','','status=yes,scrollbars=yes,resizable=yes,width=800,height=500')"">Parcela Vencida</a>"
	  else
		  situacao="Parcela Vencida"
	  end if	    
	else
	 if check mod 2 =0 then
		cor = "tb_fundo_linha_par_extr" 
		onblur="mudar_cor_blur_par"		 
	 else 
		cor ="tb_fundo_linha_impar_extr"
		onblur="mudar_cor_blur_impar"	
	 end if  
	end if
	
	
%>
                <tr class="<% response.Write(cor)%>" id="<%response.Write("celula"&linha)%>" > 
                  <td width="87" onFocus="mudar_cor_focus(<%response.Write("celula"&linha)%>)" onBlur="<%response.Write(onblur)%>(<%response.Write("celula"&linha)%>)"> <div align="center"> 
                      <%' if situacao="aberto" and UCASE(nome_lanc) = "MENSALIDADE" then %>
<!--                                                  <a href="#" class="menu_sublista" onClick="MM_openBrWindow('../segvia/bloqueto.asp?c=<%=cod%>&amp;m=<%=mes_venc%>&amp;v=<%=venc%>&amp;opt=c','','width=700,height=100')"> 
                      <%' response.Write(da_vencimento_show)%>
                                                   </a> -->
  
                      <%'elseif situacao="em atraso" and UCASE(nome_lanc) = "MENSALIDADE" then %>
<!--                                                  <a href="#" class="menu_lista" onClick="MM_openBrWindow('../segvia/bloqueto.asp?c=<%=cod%>&amp;m=<%=mes_venc%>&amp;v=<%=venc%>&amp;opt=c','','width=700,height=100')"> 
 
                      <%' response.Write(da_vencimento_show)%>
                                                 </a> 
  -->
                      <%'else
 response.Write(da_vencimento_show)
'end if%>
                    </div></td>
                  <td width="139" onFocus="mudar_cor_focus(<%response.Write("celula"&linha)%>)" onBlur="<%response.Write(onblur)%>(<%response.Write("celula"&linha)%>)"> <div align="center"> 
                      <% response.Write(nome_lanc)%>
                    </div></td>
                  <td width="105" align="right" onFocus="mudar_cor_focus(<%response.Write("celula"&linha)%>)" onBlur="<%response.Write(onblur)%>(<%response.Write("celula"&linha)%>)">
                      <% response.Write(FormatNumber(compromisso))%>
                    </td>
                  <td width="92" align="right" onFocus="mudar_cor_focus(<%response.Write("celula"&linha)%>)" onBlur="<%response.Write(onblur)%>(<%response.Write("celula"&linha)%>)"><%
				  if isnull(da_realizado)then
				  	  val_multa = CalculaMulta(da_vencimento, data_calc, compromisso)
					  if val_multa>0 then
				  	  	response.Write(FormatNumber(val_multa))
					   end if	
				  end if
				  %></td>
                  <td width="104" align="right" onFocus="mudar_cor_focus(<%response.Write("celula"&linha)%>)" onBlur="<%response.Write(onblur)%>(<%response.Write("celula"&linha)%>)">
                  <%
				  if isnull(da_realizado) then
				  	  val_mora = CalculaMora(da_vencimento, data_calc, compromisso)		
					  if val_mora>0 then					  		  
				  	  	response.Write(FormatNumber(val_mora))
					  end if						  
				  end if
				  %></td>
                  <td width="119" align="right" onFocus="mudar_cor_focus(<%response.Write("celula"&linha)%>)" onBlur="<%response.Write(onblur)%>(<%response.Write("celula"&linha)%>)"><%
				 if val_multa>0 or val_mora>0 then				  
					  val_corrigido = compromisso+val_multa+val_mora
					  response.Write(FormatNumber(val_corrigido))
					  
					  multa_total=multa_total+val_multa
					  mora_total=mora_total+val_mora
					  corrigido_total=corrigido_total+val_corrigido
				 end if				  
				  %></td>
                  <td width="105" align="right" onFocus="mudar_cor_focus(<%response.Write("celula"&linha)%>)" onBlur="<%response.Write(onblur)%>(<%response.Write("celula"&linha)%>)">
                      <% response.Write(realizado)%>
                    </td>
                  <td width="114" onFocus="mudar_cor_focus(<%response.Write("celula"&linha)%>)" onBlur="<%response.Write(onblur)%>(<%response.Write("celula"&linha)%>)"> <div align="center"> 
                      <% response.Write(da_realizado_show)%>
                    </div></td>
                  <td width="107" onFocus="mudar_cor_focus(<%response.Write("celula"&linha)%>)" onBlur="<%response.Write(onblur)%>(<%response.Write("celula"&linha)%>)"> <div align="center"> 
                      <% response.Write(situacao)%>
                    </div></td>
                  <td width="26" onFocus="mudar_cor_focus(<%response.Write("celula"&linha)%>)" onBlur="<%response.Write(onblur)%>(<%response.Write("celula"&linha)%>)">
                  <%'and d_diff<0			  
				  if isnull(da_realizado)  and nome_lanc="Mensalidade" and existe_email = "S" then
				   %>
                  <a href="confirma.asp?opt=<%response.Write(p_vencimento)%>&c=<%response.Write(cod)%>"><img src="../../../../img/email.gif" alt="Envia boleto por e-mail" width="20" height="20"></a>
                  <% end if%>
                  </td>
                </tr>
                <%
linha = linha+1				
RS4.MOVENEXT
WEND
END IF
 if check mod 2 =0 then
  cor = "tb_fundo_linha_par_extr" 
  else cor ="tb_fundo_linha_impar_extr"
  end if  
%>
                <tr class="<% = cor %>"> 
                  <td width="87" align="center"><b>Total</b></td>
                  <td width="139" align="center">&nbsp;</td>
                  <td width="105" align="right"><b><%response.Write(FormatCurrency(compromisso_total))%></b></td>
                  <td width="92" align="right">
                  <b><%				  
					  if multa_total>0 then
					  response.Write(FormatCurrency(multa_total))
					  end if
				  %></b>
                  </td>
                  <td width="104" align="right"><b><%				  
					  if mora_total>0 then
					  response.Write(FormatCurrency(mora_total))
					  end if
				  %></b></td>
                  <td width="119" align="right"><b><%				  
					  if multa_total>0 or mora_total>0 then
					  response.Write(FormatCurrency(corrigido_total))
					  end if
				  %></b></td>
                  <td width="105" align="right"><b><%response.Write(FormatCurrency(realizado_total))%></b></td>
                  <td width="114" align="center">&nbsp;</td>
                  <td width="107" align="center">&nbsp;</td>
                  <td width="26" align="center">&nbsp;</td>
                </tr>
                <tr class="<% = cor %>"> 
                  <td colspan="10">&nbsp;</td>
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
response.redirect("../../../../inc/erro.asp")
end if
%>