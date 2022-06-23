<%	'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<% 
nivel=4
concatena_contrato = request.QueryString("cc")
opt = request.QueryString("opt")


permissao = session("permissao") 
ano_letivo_wf = session("ano_letivo_wf")
sistema_local=session("sistema_local")
nvg = session("chave")
chave=nvg
session("chave")=chave
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)


cod_form=Session("cod_form")
nome_form=Session("nome_form")
compromissos=Session("compromissos")
anuidade = Session("anuidade")
servicos = Session("servicos")		
aberto = Session("aberto")
liq_escola = Session("liq_escola")
liq_banco = Session("liq_banco")	
dia_de= Session("dia_de")
mes_de= Session("mes_de")
ano_de=Session("ano_de")
dia_ate=Session("dia_ate")
mes_ate=Session("mes_ate")
ano_ate=Session("ano_ate")
unidade=Session("unidade")
curso=Session("curso")
etapa=Session("etapa")
turma=session("turma")

Session("cod_form")=cod_form
Session("nome_form")=nome_form
Session("compromissos")=compromissos
Session("anuidade")=anuidade
Session("servicos")=servicos	
Session("aberto")=aberto
Session("liq_escola")=liq_escola
Session("liq_banco")=liq_banco
Session("dia_de")=dia_de
Session("mes_de")=mes_de
Session("ano_de")=ano_de
Session("dia_ate")=dia_ate
Session("mes_ate")=mes_ate
Session("ano_ate")=ano_ate
Session("unidade")=unidade
Session("curso")=curso
Session("etapa")=etapa
session("turma") =turma

data_de=mes_de&"/"&dia_de&"/"&ano_de
data_ate=mes_ate&"/"&dia_ate&"/"&ano_ate

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		
		
    	Set CON3 = Server.CreateObject("ADODB.Connection") 
		ABRIR3= "DBQ="& CAMINHO_cp & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3		
		
    	Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4= "DBQ="& CAMINHO_lq & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4			

		Set CON5 = Server.CreateObject("ADODB.Connection") 
		ABRIR5 = "DBQ="& CAMINHO_cr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON5.Open ABRIR5

		Set CONa = Server.CreateObject("ADODB.Connection") 
		ABRIRa = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONa.Open ABRIRa	

if opt="vi" or opt="i" then
	vetor_reenvio=concatena_contrato
	vetor_concatena=split(concatena_contrato,"$")
	tipo_consulta=vetor_concatena(0)
	if tipo_consulta = "a" then
		co_cons=vetor_concatena(1) 
	else
		info=split(vetor_concatena(1),"-")
		unidade_opt=info(0)
		curso_opt=info(1)
		etapa_opt=info(2)
		turma_opt=info(3)
	end if	
	tp_compromisso=vetor_concatena(2) 
	compromisso=vetor_concatena(3) 
	pergunta1=vetor_concatena(4) 
	mes_de=vetor_concatena(5) 		
	mes_ate=vetor_concatena(6) 	
	
	if opt = "i" then
		acao_compromissos = request.form("acao_compromissos")
		
		if acao_compromissos = "I" then
			response.Redirect("compromissos.asp?pagina=1&v=s")
		end if
	end if
	if tp_compromisso = "A" then
		nome_tp_compromisso = "Parcelas de Anuidade"
		tit_pergunta1="Gera Parcelas  para o período informado no contrato?"
		sql_compromissos = "TP_Compromisso = 'COTA' AND"
	else
		nome_tp_compromisso = "Parcelas de Serviços Adicionais"	
		tit_pergunta1="Sincroniza com as Parcelas de Anuidade?"
		sql_compromissos = "TP_Compromisso <> 'COTA' AND"
	end if		
	if pergunta1 = "S" then
		tx_pergunta1 = "Sim"
	else
		tx_pergunta1 = "Não"	
	end if
    
	if tp_compromisso = "A" and pergunta1 = "N" then
		select case mes_de
			case 1
				tx_mes_de="Janeiro"
			case 2
				tx_mes_de="Fevereiro"
			case 3
				tx_mes_de="Março"
			case 4
				tx_mes_de="Abril"
			case 5
				tx_mes_de="Maio"
			case 6
				tx_mes_de="Junho"
			case 7
				tx_mes_de="Julho"
			case 8
				tx_mes_de="Agosto"
			case 9
				tx_mes_de="Setembro"																								
			case 10
				tx_mes_de="Outubro"
			case 11
				tx_mes_de="Novembro"
			case 12
				tx_mes_de="Dezembro"	
		end select		
		
		select case mes_ate
			case 1
				tx_mes_ate="Janeiro"
			case 2
				tx_mes_ate="Fevereiro"
			case 3
				tx_mes_ate="Março"
			case 4
				tx_mes_ate="Abril"
			case 5
				tx_mes_ate="Maio"
			case 6
				tx_mes_ate="Junho"
			case 7
				tx_mes_ate="Julho"
			case 8
				tx_mes_ate="Agosto"
			case 9
				tx_mes_ate="Setembro"																								
			case 10
				tx_mes_ate="Outubro"
			case 11
				tx_mes_ate="Novembro"
			case 12
				tx_mes_ate="Dezembro"	
		end select			
		tx_pergunta2="De "&tx_mes_de&" até "&tx_mes_ate&"."
	end if	
	
	Set RSCp= Server.CreateObject("ADODB.Recordset")
	SQLC = "SELECT MIN(DA_Vencimento) as menor_data, MAX(DA_Vencimento) as maior_data FROM TB_Compromissos where "&sql_aluno&sql_compromissos&" (DA_Vencimento BETWEEN #"&data_de&"# AND #"&data_ate&"#)"
	RSCp.Open SQLC, CON3	
	
	if RSCp.EOF then
		response.Redirect("confirma.asp?opt=i&cc="&vetor_reenvio)
	else
		menor_data=RSCp("menor_data")
		maior_data=RSCp("maior_data")
	end if
		
else
	vetor_concatena = split(concatena_contrato,", ")
	tipo_consulta=""
	dados_msg = UBOUND(vetor_concatena)+1
end if

		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")		
	

if tipo_consulta = "a" then
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT a.NO_Aluno as NOME, m.NU_Unidade as UNI, m.CO_Curso as CUR, m.CO_Etapa as ETA, m.CO_Turma as TUR FROM TB_Alunos a, TB_Matriculas m where a.CO_Matricula = "& co_cons& " and a.CO_Matricula = m.CO_Matricula and m.NU_Ano = "&ano_letivo		
		RS.Open SQL, CONa	
		
		nome_cons = RS("NOME")		
		unidade_bd = RS("UNI")
		curso_bd = RS("CUR")
		etapa_bd = RS("ETA")
		turma_bd = RS("TUR")
		
		dados_msg =nome_cons 	
		sql_aluno = "CO_Matricula = "& cod_cons & " and "			

elseif tipo_consulta = "u" then

	
	if unidade_opt="999990" or unidade_opt="" or isnull(unidade_opt) then
		sql_un=""
		unidade_nome="Todas"
	else

		sql_un=" AND NU_Unidade= "&unidade_opt
	
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="&unidade_opt
		RS0.Open SQL0, CON0
		
		unidade_nome = RS0("NO_Unidade")
	end if
	
	if curso_opt="999990" or curso_opt="" or isnull(curso_opt) then
		sql_cu=""
		curso_nome="Todos"
	else
		sql_cu=" AND CO_Curso='"&curso_opt&"'"
		
		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Curso where CO_Curso='"&curso_opt&"'"
		RS0c.Open SQL0c, CON0
		
		curso_nome = RS0c("NO_Curso")
	end if
	
	if etapa_opt="999990" or etapa_opt="" or isnull(etapa_opt) then
		sql_et=""
		etapa_nome="Todas"
	else
		sql_et=" AND CO_Etapa='"&etapa_opt&"'"
	
		Set RS0e = Server.CreateObject("ADODB.Recordset")
		SQL0e = "SELECT * FROM TB_Etapa where CO_Curso='"&curso_opt&"' AND CO_Etapa='"&etapa_opt&"'"
		RS0e.Open SQL0e, CON0
		
		etapa_nome = RS0e("NO_Etapa")
	end if
	
	if turma_opt="999990" or turma_opt="" or isnull(turma_opt) then
		sql_tu=""
		turma_nome="Todas"
		turma_selecionada="n"		
	else
		sql_tu=" AND CO_Turma='"&turma_opt&"'"
		turma_nome=turma
		turma_selecionada="s"
	end if
	
	Set RSm = Server.CreateObject("ADODB.Recordset")
	SQLm = "SELECT CO_Matricula as MATRIC FROM TB_Matriculas where NU_Ano = "&ano_letivo&sql_un&sql_cu&sql_et&sql_tu 
	RSm.Open SQLm, CONa		
	
	conta=0
	if RSm.EOF then
		sql_aluno=""
	else
		while not RSm.EOF 
			matric_aluno=RSm("MATRIC")
			if conta=0	then
				vetor_matric=matric_aluno
			else
				vetor_matric=vetor_matric&", "&matric_aluno		
			end if
		conta=conta+1	
		RSm.MOVENEXT
		WEND	

		sql_aluno = "CO_Matricula IN ("& vetor_matric &") and "		
		
	end if	
	
	dados_msg ="Unidade: "&unidade_nome&", Curso: "&curso_nome&", Etapa: "&etapa_nome&" e Turma: "&turma_nome
end if
%>
<html>
<head>
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../../../js/mm_menu.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--

<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
<%if opt="vi" then%>
function ativa_bt(ativar)  
{
if (ativar =="S"){
	document.getElementById('confirmar').disabled   = false;	
	}	else {
	document.getElementById('confirmar').disabled   = true;			
	}
}
<%end if%>
//-->
</script>
</head>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%if opt="vi" then%>
<form action="confirma.asp?opt=i&cc=<%response.Write(vetor_reenvio)%>" method="post" name="busca" id="busca">
<%else%>
<form action="bd.asp?opt=<%response.Write(opt)%>" method="post" name="busca" id="busca">
<%end if%>
<%call cabecalho_novo(nivel)
%>
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
      <td valign="top"> 
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo">
        <%if opt = "vi" then
			tit="Compromissos a serem criados"
			
				dados_msg= dados_msg&"$$$"&menor_data&"$!$"&maior_data

			%>
			   <tr> 
				<td width="766" height="10" colspan="4" valign="top"> 
				  <%call mensagens(nivel,807,0,dados_msg) %>
				</td>
			  </tr>       
        <%			
		elseif opt = "i" then
		tit="Compromissos a serem criados"
			%>
			   <tr> 
				<td width="766" height="10" colspan="4" valign="top"> 
				  <%call mensagens(nivel,808,0,dados_msg) %>
				</td>
			  </tr>       
        <%
		else
		tit="Contratos a serem cancelados"
		%>
          <tr> 
            <td width="766" height="10" colspan="4" valign="top"> 
              <%call mensagens(nivel,803,0,dados_msg) %>
            </td>
          </tr>
        <%end if%>  
          <tr> 
            <td height="10" class="tb_tit"><%response.Write(tit)%></td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
<%
if opt="vi" then%>
                     <td colspan="10" align="center">
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr class="tb_subtit">
                        <td colspan="4" align="left">Selecione o tipo de a&ccedil;&atilde;o a ser tomada</td>
                        <%if tp_compromisso = "A" and pergunta1 = "N" then%>
                        <%end if%>
                      </tr>
                      <tr class="form_dado_texto">
                        <td width="12%" align="center">&nbsp;</td>
                        <td width="4%" align="center"><input name="acao_compromissos" type="radio" class="nota" id="acao_compromissos" value="M" onClick="ativa_bt('S')"></td>
                        <td width="1%" align="center">&nbsp;</td>                   
                        <td width="83%" align="left">Manter os compromissos existentes e gerar apenas para o per&iacute;odo que n&atilde;o coincidir com o que j&aacute; existe</td>                    
                      </tr>
                      <tr class="form_dado_texto">
                        <td align="center">&nbsp;</td>
                        <td align="center"><input name="acao_compromissos" type="radio" class="nota" id="acao_compromissos" value="S" onClick="ativa_bt('S')"></td>
                        <td align="center">&nbsp;</td>
                        <td align="left">Sobrescrever&nbsp; os compromissos existentes e n&atilde;o baixados </td>
                      </tr>
                      <tr class="form_dado_texto">
                        <td align="center">&nbsp;</td>
                        <td align="center"><input name="acao_compromissos" type="radio" class="nota" id="acao_compromissos" value="I" onClick="ativa_bt('S')"></td>
                        <td align="center">&nbsp;</td>
                        <td align="left"> Interromper a inclus&atilde;o de compromissos </td>
                      </tr>
                    </table>
          		</td>
                </tr>
<%
elseif opt = "i" then

	Set RSc = Server.CreateObject("ADODB.Recordset")
	SQLc = "SELECT * FROM TB_Tipo_Compromisso WHERE TP_Lancamento = '"&tp_compromisso&"' AND CO_Compromisso = '"&compromisso&"'"
	RSc.Open SQLc, CON3
	
	if RSc.EOF then
		nome_compromisso=""
	else
		nome_compromisso=RSc("NO_Compromisso")
	end if

%>  
              <tr>
                <td colspan="10" align="center">
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr class="tb_subtit">
                        <td align="center">Tipo Lançamento</td>
                        <td align="center">Tipo de Compromisso </td>
                        <td align="center"><%response.Write(tit_pergunta1)%></td>
                       <%if tp_compromisso = "A" and pergunta1 = "N" then%>
                        <td align="center">Gera parcelas para o período?</td>
                        <%end if%>
                      </tr>
                      <tr class="form_dado_texto">
                        <td align="center"><%response.Write(nome_tp_compromisso)%></td>
                        <td align="center"><%response.Write(nome_compromisso)%></td>
                        <td align="center"><%response.Write(tx_pergunta1)%></td>
                       <%if tp_compromisso = "A" and pergunta1 = "N" then%>                        
                        <td align="center"><%response.Write(tx_pergunta2)%></td>
                        <%end if%>                        
                      </tr>
                    </table>
          		</td>
                </tr>
<%elseif opt = "e" then%>            
              <tr class="tb_subtit">
                <td width="20" align="center"><div align="center"><input name="excluir_contratos" type="hidden" id="excluir_contratos" value="<%response.Write(concatena_contrato)%>"></div></td>
                <td align="center">N&uacute;mero</td>
                <td align="center">Matr&iacute;cula</td>
                <td align="center">Aluno</td>
                <td align="center">Data</td>
                <td align="center">Tipo de Compromisso</td>
                <td align="center">Parcela</td>
                <td align="center">Cota</td>
                <td align="center">Seq&uuml;&ecirc;ncia</td>
                <td align="center">Situa&ccedil;&atilde;o</td>
                </tr>
              <%

	FOR vc = 0 to ubound(vetor_concatena)
	
	
		vetor_temp = split(vetor_concatena(vc),"-")
		nu_matric = vetor_temp(0)		
		vetor_compromissos = split(vetor_temp(1),"$")
		ano_contrato = vetor_compromissos(0)	
		contratos = vetor_compromissos(1)
		co_compromissos = vetor_compromissos(2)
		nu_parcela = vetor_compromissos(3)
		nu_cota = vetor_compromissos(4)	
		nu_sequencial = vetor_compromissos(5)
					  
				  
		Set RSC= Server.CreateObject("ADODB.Recordset")
		SQLC = "SELECT * FROM TB_Compromissos where CO_Matricula = "&nu_matric&" AND NU_Ano_Letivo = "&ano_contrato&" AND NU_Contrato = "&contratos&" AND TP_Compromisso = '"&co_compromissos&"' AND NU_Parcela = "&nu_parcela&" AND NU_Cota = "&nu_cota&" AND NU_Sequencial = "&nu_sequencial
		RSC.Open SQLC, CON3
		
		if RSC.EOF then
	
		else
	
			check=2
			While Not RSC.EoF
			
			 if check mod 2 =0 then
				cor = "tb_fundo_linha_par" 
			 else 
				cor ="tb_fundo_linha_impar"
			 end if
			
				data_vencimento= RSC("DA_Vencimento")
				ano_contrato = RSC("NU_Ano_Letivo")
				nu_contrato = RSC("NU_Contrato")
				if nu_contrato<100000 then
					if nu_contrato<10000 then
						if nu_contrato<1000 then
							if nu_contrato<100 then
								if nu_contrato<10 then
									nu_contrato="00000"&nu_contrato							
								else
									nu_contrato="0000"&nu_contrato					
								end if						
							else
								nu_contrato="000"&nu_contrato					
							end if	
						else
							nu_contrato="00"&nu_contrato					
						end if
					else
						nu_contrato="0"&nu_contrato					
					end if
				end if	 
				
				
				concatena_contrato = ano_contrato&"/"&nu_contrato
				matricula_contrato = RSC("CO_Matricula")
				co_compromissos = RSC("TP_Compromisso")			
				nu_parcela = RSC("NU_Parcela")
				nu_cota = RSC("NU_Cota")		
				nu_sequencial = RSC("NU_Sequencial")	
					
				if isnull(co_compromissos) or co_compromissos= "" then
					no_compromisso=""
				else
					Set RScp = Server.CreateObject("ADODB.Recordset")
					SQLcp = "SELECT * FROM TB_Tipo_Compromisso WHERE CO_Compromisso='"&co_compromissos&"'"
					RScp.Open SQLcp, CON3
					
					no_compromisso=RScp("NO_Compromisso")	
				end if		
				
				Set RSL= Server.CreateObject("ADODB.Recordset")
				SQLL = "SELECT * FROM TB_Lancamento_Realizado where NU_Sequencial = "&nu_sequencial
				RSL.Open SQLL, CON4		
					
				if RSL.EOF then
					if aberto = "s" then
						exibe_contrato="s"
						nome_situacao="Em Aberto"
					else
						exibe_contrato="n"				
					end if	
				else
					exibe_contrato="s"
					co_situacao=RSL("IN_Lancamento_Bancario")
					if co_situacao=TRUE then
						nome_situacao="Liq. Banco"					
					else
						nome_situacao="Liq. Escola"				
					end if	
				end if				
				
				Set RSA = Server.CreateObject("ADODB.Recordset")	
				SQLA = "SELECT a.NO_Aluno as NOME, m.NU_Unidade as UNI, m.CO_Curso as CUR, m.CO_Etapa as ETA, m.CO_Turma as TUR FROM TB_Alunos a, TB_Matriculas m where a.CO_Matricula = "& matricula_contrato & " and a.CO_Matricula = m.CO_Matricula and m.NU_Ano = "&ano_contrato	
				RSA.Open SQLA, CONa		
					
				if RSA.EOF then
					nome_contrato = "N&atilde;o Informado"
					unidade_contrato = "N.I."
					curso_contrato = "N.I."
					etapa_contrato = "N.I."
					turma_contrato = "N.I."
				else
					nome_contrato = RSA("NOME")		
					co_unidade_contrato = RSA("UNI")
					co_curso_contrato = RSA("CUR")
					etapa_contrato = RSA("ETA")
					turma_contrato = RSA("TUR")	
					
					Set RS0 = Server.CreateObject("ADODB.Recordset")
					SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="&co_unidade_contrato
					RS0.Open SQL0, CON0
					
					unidade_contrato = RS0("NO_Abr")
			
					sql_cu="(Curso='"&curso&"' OR (Curso  is null)) AND"
					Set RS0c = Server.CreateObject("ADODB.Recordset")
					SQL0c = "SELECT * FROM TB_Curso where CO_Curso='"&co_curso_contrato&"'"
					RS0c.Open SQL0c, CON0
					
					curso_contrato = RS0c("NO_Abreviado_Curso")			
				end if			
		
		%>
					  <tr>
						<td class="<%response.Write(cor)%>"></td>
						<td align="center" class="<%response.Write(cor)%>"><%response.Write(concatena_contrato)%></td>
						<td align="center" class="<%response.Write(cor)%>"><%response.Write(matricula_contrato)%></td>
						<td align="center" class="<%response.Write(cor)%>"><%response.Write(nome_contrato)%></td>
						<td align="center" class="<%response.Write(cor)%>"><%response.Write(data_vencimento)%></td>
						<td align="center" class="<%response.Write(cor)%>"><%response.Write(no_compromisso)%></td>
						<td align="center" class="<%response.Write(cor)%>"><%response.Write(nu_parcela)%></td>
						<td align="center" class="<%response.Write(cor)%>"><%response.Write(nu_cota)%></td>
						<td align="center" class="<%response.Write(cor)%>"><%response.Write(nu_sequencial)%></td>
						<td align="center" class="<%response.Write(cor)%>"><%response.Write(nome_situacao)%></td>
					  </tr>
					  <% 		
				intrec=intrec+1
				check=check+1	
			RSC.MOVENEXT
			WEND
		end if
	NEXT
end if	
%>
              <tr>
                <td colspan="10" align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="50%"><hr></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td><div align="center"> 
                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr> 
                    <td width="33%"> <div align="center"> 
                    <%  if opt="vi" then
							ativa_incluir="disabled"
							url="incluir.asp?ci="&concatena_contrato							
						elseif opt="i" then
							ativa_incluir=""							
							url="incluir.asp?ci="&concatena_contrato
						else
							ativa_incluir=""							
							url="compromissos.asp?pagina=1&v=s"						
						end if
					%>
                        <input name="SUBMIT5" type=button class="botao_cancelar" onClick="MM_goToURL('parent','<%response.Write(url)%>');return document.MM_returnValue" value="Voltar">
                    </div></td>
                    <td width="34%"> <div align="center"> </div> <div align="center"> </div></td>
                    <td width="33%"> <div align="center"> 
                        <input name="Submit" type="submit" class="botao_prosseguir" id="confirmar" value="Confirmar" <%response.Write(ativa_incluir)%>>
                    </div></td>
                  </tr>
                  <tr>
                    <td width="33%">&nbsp;</td>
                    <td width="34%">&nbsp;</td>
                    <td width="33%">&nbsp;</td>
                  </tr>
                </table>
            <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> </font></div></td>
          </tr>
        </table></td>
    </tr>
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
</form>
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