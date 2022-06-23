<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style>
<script type="text/javascript">
<!--
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
						
								   
function ComboCompromissos()
                                   {
if (document.getElementById('anuidade').checked == false ) {
	p_anuidade = "F"} else{
	p_anuidade = "S";
	document.all.dPergunta1.innerHTML = "Gera Parcelas  para o período informado no contrato?";	
	document.all.dFormPergunta1.innerHTML = "<select name='gera_prd_contrato' class='select_style' id='gera_prd_contrato' onchange='ComboPergunta2(this.value)' disabled><option value='nulo' selected></option><option value='S'>Sim</option><option value='N'>Não</option></select>"	
	}

if (document.getElementById('servicos').checked == false ) {
	p_servicos = "F"} else{
	p_servicos = "S";
	document.all.dPergunta1.innerHTML = "Sincroniza com as Parcelas de Anuidade?";
	document.all.dFormPergunta1.innerHTML = "<select name='gera_prd_contrato' class='select_style' id='gera_prd_contrato' onchange='	ComboPergunta2(this.value)' disabled><option value='nulo' selected></option><option value='S'>Sim</option><option value='N'>Não</option></select>"	
	}									   
// Criação do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicitação HTTP. O primeiro parâmetro informa o método post/get
// O segundo parâmetro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicitação síncrona, o parâmetro deve ser false
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=cc", true);
// Para solicitações utilizando o método post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A função abaixo é executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto já completou a solicitação
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto é gerado no arquivo executa.asp e colocado no div
                                                                       var resultado_cc= oHTTPRequest.responseText;
resultado_cc = resultado_cc.replace(/\+/g," ")
resultado_cc = unescape(resultado_cc)
document.all.divCompromissos.innerHTML = resultado_cc																	   
                                                           }
                                               }
// Abaixo é enviada a solicitação. Note que a configuração
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("p_anuidade=" + p_anuidade + "&p_servicos = "+ p_servicos);
                                   }										   
								   
	function ComboPergunta2(valor)	
	{						   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=cp", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                       var resultado_cp= oHTTPRequest.responseText;

	if (valor=="N") {
		document.all.dPergunta2.innerHTML = "Gera parcelas para o período?";
	}else{
		document.all.dPergunta2.innerHTML = "";
		}
	resultado_cp = resultado_cp.replace(/\+/g," ")
	resultado_cp = unescape(resultado_cp)
	document.all.dFormPergunta2.innerHTML = resultado_cp
if (document.getElementById('anuidade').checked == false ) {	
	ComboPergunta3(valor)
	ComboPergunta4(valor)
	ComboPergunta5(valor)
}
//ativa_bt("S");
                                                           }
                                               }

                                               oHTTPRequest.send("p_pergunta=" + valor);
                                   }										   



	function ComboPergunta3(valor)	
	{						   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=cp3", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                       var resultado_cp3= oHTTPRequest.responseText;
if (valor=="N") {
	document.all.dPergunta3.innerHTML = "Selecione o Dia para vencimento";
}else{
	document.all.dPergunta3.innerHTML = "";
	}
resultado_cp3 = resultado_cp3.replace(/\+/g," ")
resultado_cp3 = unescape(resultado_cp3)
document.all.dFormPergunta3.innerHTML = resultado_cp3
//ativa_bt("S");
                                                           }
                                               }

                                               oHTTPRequest.send("p_pergunta=" + valor);
                                   }										   

	function ComboPergunta4(valor)	
	{						   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=cp4", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                       var resultado_cp4= oHTTPRequest.responseText;
if (valor=="N") {
	document.all.dPergunta4.innerHTML = "Caso o Dia para vencimento Seja não útil";
}else{
	document.all.dPergunta4.innerHTML = "";
	}
resultado_cp4 = resultado_cp4.replace(/\+/g," ")
resultado_cp4 = unescape(resultado_cp4)
document.all.dFormPergunta4.innerHTML = resultado_cp4
//ativa_bt("S");
                                                           }
                                               }

                                               oHTTPRequest.send("p_pergunta=" + valor);
                                   }										   
	function ComboPergunta5(valor)	
	{						   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=cp5", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                       var resultado_cp5= oHTTPRequest.responseText;


	resultado_cp5 = resultado_cp5.replace(/\+/g," ")
	resultado_cp5 = unescape(resultado_cp5)
if (valor=="N") {
	document.all.dPergunta5.innerHTML = "Valor do Compromisso";
	document.all.dFormPergunta5.innerHTML = resultado_cp5	
}else{
	document.all.dPergunta2.innerHTML = "Valor do Compromisso";
	document.all.dFormPergunta2.innerHTML = resultado_cp5	
	document.all.dPergunta5.innerHTML = "";
	document.all.dFormPergunta5.innerHTML = ""		
	}

                                                           }
                                               }

                                               oHTTPRequest.send("p_pergunta=" + valor + "&p_compromisso =" +document.getElementById('compromissos').value );
                                   }										   

function LimpaPerguntas()
{
	document.all.dPergunta2.innerHTML = "";
	document.all.dFormPergunta2.innerHTML = "";
	document.all.dPergunta3.innerHTML = "";
	document.all.dFormPergunta3.innerHTML = "";
	document.all.dPergunta4.innerHTML = "";
	document.all.dFormPergunta4.innerHTML = "";	
	document.all.dPergunta5.innerHTML = "";
	document.all.dFormPergunta5.innerHTML = "";	


}
function ativa_bt(ativar)  
{
if (ativar =="S"){
	document.getElementById('acao3').disabled   = false;	
	}	else {
	document.getElementById('acao3').disabled   = true;			
	}
}

function ativa_combo(ativar)  
{
if (ativar =="S"){
	document.getElementById('gera_prd_contrato').disabled   = false;	
	}	else {
	document.getElementById('gera_prd_contrato').disabled   = true;			
	}
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
function submitforminterno()  
{
   var f=document.forms[3]; 
      f.submit(); 
	  
}
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}

function checksubmit()
{
 if (document.getElementById('compromissos').value == "nulo")
  {    alert("É necessário selecionar um tipo de compromissos!")
	 document.form.compromissos.focus()
    return false
 } 
  return true
}
//-->
</script>
<script type="text/javascript">  
// Formata o campo valor
function formataValor(campo) {
	campo.value = filtraCampoValor(campo); 
	vr = campo.value;
	tam = vr.length;

	if ( tam <= 2 ){ 
 		campo.value = vr ; }
 	if ( (tam > 2) && (tam <= 5) ){
 		campo.value = vr.substr( 0, tam - 2 ) + ',' + vr.substr( tam - 2, tam ) ; }
 	if ( (tam >= 6) && (tam <= 8) ){
 		campo.value = vr.substr( 0, tam - 5 ) + '.' + vr.substr( tam - 5, 3 ) + ',' + vr.substr( tam - 2, tam ) ; }
 	if ( (tam >= 9) && (tam <= 11) ){
 		campo.value = vr.substr( 0, tam - 8 ) + '.' + vr.substr( tam - 8, 3 ) + '.' + vr.substr( tam - 5, 3 ) + ',' + vr.substr( tam - 2, tam ) ; }
 	if ( (tam >= 12) && (tam <= 14) ){
 		campo.value = vr.substr( 0, tam - 11 ) + '.' + vr.substr( tam - 11, 3 ) + '.' + vr.substr( tam - 8, 3 ) + '.' + vr.substr( tam - 5, 3 ) + ',' + vr.substr( tam - 2, tam ) ; }
 	if ( (tam >= 15) && (tam <= 18) ){
 		campo.value = vr.substr( 0, tam - 14 ) + '.' + vr.substr( tam - 14, 3 ) + '.' + vr.substr( tam - 11, 3 ) + '.' + vr.substr( tam - 8, 3 ) + '.' + vr.substr( tam - 5, 3 ) + ',' + vr.substr( tam - 2, tam ) ;}
 		
}
//limpa todos os caracteres especiais do campo solicitado
function filtraCampoValor(campo){
	var s = "";
	var cp = "";
	vr = campo.value;
	tam = vr.length;
	for (i = 0; i < tam ; i++) {  
		if (vr.substring(i,i + 1) >= "0" && vr.substring(i,i + 1) <= "9"){
		 	s = s + vr.substring(i,i + 1);}
	} 
	campo.value = s;
	return cp = campo.value
}
</script> 
</head>
<%
nivel=4
nvg = session("chave")
chave=nvg
session("chave")=chave

ano_letivo = session("ano_letivo")
opt=Request.QueryString("ci")

ativar_bt = "N"
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


		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CONa = Server.CreateObject("ADODB.Connection") 
		ABRIRa = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONa.Open ABRIRa
		
    	Set CON3 = Server.CreateObject("ADODB.Connection") 
		ABRIR3= "DBQ="& CAMINHO_cp & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON3.Open ABRIR3		
		
    	Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4= "DBQ="& CAMINHO_lq & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4				
		
tipo_inf=split(opt,"$")

if tipo_inf(0) = "a" then
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT a.NO_Aluno as NOME, m.NU_Unidade as UNI, m.CO_Curso as CUR, m.CO_Etapa as ETA, m.CO_Turma as TUR FROM TB_Alunos a, TB_Matriculas m where a.CO_Matricula = "& tipo_inf(1) & " and a.CO_Matricula = m.CO_Matricula and m.NU_Ano = "&ano_letivo		
		RS.Open SQL, CONa	
		
		nome_cons = RS("NOME")		
		unidade_bd = RS("UNI")
		curso_bd = RS("CUR")
		etapa_bd = RS("ETA")
		turma_bd = RS("TUR")
		
		dados_msg =nome_cons 	
		concatena_aluno = "a$"&tipo_inf(1)		
else
	info=split(tipo_inf(1),"-")
	unidade_opt=info(0)
	curso_opt=info(1)
	etapa_opt=info(2)
	turma_opt=info(3)
	
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
'	response.Write(SQLm)
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
	concatena_aluno = "u$"&unidade_opt&"-"&curso_opt&"-"&etapa_opt&"-"&turma_opt
end if

%>
<body <%response.Write(onload)%>>
<form name = "form" id="form" action="processa_form.asp" method="post" onSubmit="return checksubmit()">  
<% call cabecalho_novo (nivel)
	  %>    
<table width="1000" border="0" align="center" cellpadding="0" cellspacing="0" class="tb_corpo">
  <tr> 
                    
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
	  </td>
  <%if sem_aluno = "s" then%>   
    <tr>
    <td><%call mensagens(nivel,303,1,0) %></td>
  </tr>
 <%end if%>  
  <%if opt = "ok" then%>   
    <tr>
    <td><%call mensagens(nivel,809,2,0) %></td>
  </tr>
 <%end if%>   
  <tr>
    <td><%call mensagens(nivel,806,0,dados_msg) %></td>
  </tr>
  <tr>
    <td class="tb_tit">Gerar Compromissos</td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td colspan="6">
            <table width="1000" border="0" cellpadding="0" cellspacing="0">
            <% if tipo_inf(0) = "a" then%>
              <tr>
                <td width="220"  height="30" valign="middle"><div align="right"><font class="form_dado_texto"> Matr&iacute;cula:</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> </strong></font></div></td>
                <td width="280" height="30" valign="middle"><font class="form_dado_texto"><%RESPONSE.Write(tipo_inf(1))%></font></td>
                <td width="253" height="30" valign="middle"><div align="right"><font class="form_dado_texto"> Nome: </font></div></td>
                <td width="304" height="30" valign="middle" ><font class="form_dado_texto"><%RESPONSE.Write(nome_cons)%><input name="selecao" type="hidden" value="<%response.write(concatena_aluno)%>"></font></td>
                <div id="usersList"></div>
                </tr>
             <%else%>
      <tr>
        <td width="220" height="30" class="tb_subtit"><div align="center">Unidade</div></td>
        <td width="280" height="30" class="tb_subtit"><div align="center">Curso</div></td>
        <td width="253" height="30" class="tb_subtit"><div align="center">Etapa</div></td>
        <td width="304" height="30" class="tb_subtit"><div align="center">Turma</div></td>
        </tr>
      <tr>
        <td width="220" height="30"> <div align="center"> <font class="form_dado_texto"> 
          <%response.Write(unidade_nome)%>
          </font> </div></td>
        <td width="280" height="30"> <div align="center"> <font class="form_dado_texto"> 
          <%response.Write(curso_nome)%>
          </font> </div></td>
        <td width="253" height="30"> <div align="center"> <font class="form_dado_texto"> 
          <%response.Write(etapa_nome)%>
          </font> </div></td>
        <td width="304" height="30"> <div align="center"> <font class="form_dado_texto"> 
          <%response.Write(turma_nome)%><input name="selecao" type="hidden" value="<%response.write(concatena_aluno)%>">
          </font> </div></td>
        </tr>
             <%end if%>   
              </table></td>
            </tr>
          <tr>
            <td width="24%" height="30" align="center" class="tb_subtit"> Parcelas de Anuidade </td>
            <td width="46%" height="30" align="center" class="tb_subtit"> Parcelas de Servi&ccedil;os Adicionais </td>
            <td class="tb_subtit"><div align="center"> Tipo de Compromisso </div></td>
          </tr>
          <tr>
            <td align="center"><input name="tp_compromisso" type="radio" id="anuidade" value="A" checked onClick="ComboCompromissos();LimpaPerguntas();" ></td>
            <td align="center"><input name="tp_compromisso" type="radio" id="servicos" value="S" onClick="ComboCompromissos();ativa_combo('N');LimpaPerguntas();"></td>
            <td align="center"><div id="divCompromissos">
                          <select name="compromissos" class="select_style" id="compromissos" onChange="ativa_combo('S');">
                            <option value="nulo" selected></option>
                            <%
	Set RSc = Server.CreateObject("ADODB.Recordset")
	SQLc = "SELECT * FROM TB_Tipo_Compromisso WHERE TP_Lancamento = 'A' Order By NO_Compromisso"
	RSc.Open SQLc, CON3
	while not RSc.EOF
		co_compromisso=RSc("CO_Compromisso")		
		nome_compromisso=RSc("NO_Compromisso")
%>
                            <option value="<%response.Write(co_compromisso)%>">
                              <%response.Write(nome_compromisso)%>
                              </option>
                            <%		
	RSc.MOVENEXT
	WEND	
						  %>
                          </select>
                        </div></td>
            </tr>
          <tr>
            <td colspan="2" class="tb_subtit"><div id = "dPergunta1" align="center">Gera Parcelas  para o período informado no contrato?
            </div></td>
            <td width="46%" class="tb_subtit"><div id = "dPergunta2" align="center">
            </div></td>
            </tr>
          <tr>
            <td colspan="2" align="center"><div id = "dFormPergunta1" align="center"><select name='gera_prd_contrato' class='select_style' id='gera_prd_contrato' onchange='ComboPergunta2(this.value)' disabled><option value='nulo' selected></option><option value='S'>Sim</option><option value='N'>Não</option></select>
            </div></td>
            <td width="46%"><div align="center" class="form_dado_texto" id = "dFormPergunta2"></div></td>
            </tr>
          <tr>
          	<td colspan="3">
            <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <td class="tb_subtit"><div id = "dPergunta3" align="center"></div></td>
                <td class="tb_subtit"><div id = "dPergunta4" align="center"></div></td>            
                <td class="tb_subtit"><div id = "dPergunta5" align="center"></div></td>
                </tr>
              <tr class="form_dado_texto">
                <td align="center"><div id = "dFormPergunta3" align="center"></div></td>
                <td align="center"><div id = "dFormPergunta4" align="center"></div></td>
                <td align="center"><div id = "dFormPergunta5" align="center"></div></td>
                </tr>
               </table> 
            </td>
           </tr> 
          <tr>
            <td colspan="3" align="center"><div align="center">
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr class="tb_subtit">
                  <td colspan="4" align="left">Caso j&aacute; existam compromissos gerados, selecione o tipo de a&ccedil;&atilde;o a ser tomada</td>
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
                </table>
            </div></td>
            </tr>
          </table></td>
        </tr>
      <tr>
        <td colspan="5"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td colspan="4"><hr></td>
          </tr>
          <tr>
            <td width="33%"><div align="center">
              <input name="voltar" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','compromissos.asp?pagina=1&v=s');return document.MM_returnValue" value="Voltar">
            </div></td>
            <%	if ativar_bt = "s" then
					ativa_incluir=""					
				else
					ativa_incluir="disabled"				
				end if
				%>
            <td width="34%">&nbsp;</td>
            <td width="33%"><div align="center">
              <input name="acao" id="acao3" type="submit" class="botao_prosseguir" value="Confirmar" <%response.Write(ativa_incluir)%>>
            </div></td>
          </tr>
          <tr>
            <td width="33%">&nbsp;</td>
            <td width="34%">&nbsp;</td>
            <td width="33%">&nbsp;</td>
          </tr>
        </table></td>
      </tr>
      </table></td>
  </tr>
  <tr>
      <td><img src="../../../../img/rodape.jpg" alt="" width="1000" height="40"></td>
      </tr>  
</table>
</form>
</body>
</html>
