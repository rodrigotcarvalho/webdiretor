<%	'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<% 
opt=request.QueryString("opt")
matricula_contrato=request.QueryString("mc")
ano_contrato=request.QueryString("ac")
nu_contrato=request.QueryString("nc")
nivel=4



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
contrato=Session("contrato")
ativos = Session("ativos")
sem_parcelas = Session("sem_parcelas")
so_bolsistas = Session("so_bolsistas")
dia_de= Session("dia_de")
mes_de= Session("mes_de")
ano_de=Session("ano_de")
dia_ate=Session("dia_ate")
mes_ate=Session("mes_ate")
ano_ate=Session("ano_ate")
bolsa=Session("bolsa")
desconto_de=Session("desconto_de")
desconto_ate=session("desconto_ate")
unidade=Session("unidade")
curso=Session("curso")
etapa=Session("etapa")
turma=session("turma")
situacao=session("situacao")


Session("cod_form")=cod_form
Session("nome_form")=nome_form
Session("contrato")=contrato
Session("ativos")=ativos
Session("sem_parcelas")=sem_parcelas
Session("so_bolsistas")=so_bolsistas
Session("dia_de")=dia_de
Session("mes_de")=mes_de
Session("ano_de")=ano_de
Session("dia_ate")=dia_ate
Session("mes_ate")=mes_ate
Session("ano_ate")=ano_ate
Session("bolsa")=bolsa
Session("desconto_de")=desconto_de
session("desconto_ate") =desconto_ate
Session("unidade")=unidade
Session("curso")=curso
Session("etapa")=etapa
session("turma") =turma
session("situacao") =situacao


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")		


		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		


		Set CON5 = Server.CreateObject("ADODB.Connection") 
		ABRIR5 = "DBQ="& CAMINHO_cr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON5.Open ABRIR5

		Set CONa = Server.CreateObject("ADODB.Connection") 
		ABRIRa = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONa.Open ABRIRa	
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT					

    
		Set RSC= Server.CreateObject("ADODB.Recordset")
		SQLC = "SELECT * FROM TB_Contrato where NU_Contrato = "&nu_contrato&" order by NU_Ano_Letivo Desc,NU_Contrato"		
		RSC.Open SQLC, CON5
		
		if RSC.EOF then

		else
	
	check=2
	While Not RSC.EoF
	
	 if check mod 2 =0 then
		cor = "tb_fundo_linha_par" 
	 else 
		cor ="tb_fundo_linha_impar"
	 end if
	
		dt_contrato_bd = RSC("DT_Contrato")
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
		situac_contrato = RSC("ST_Contrato")
		plano_pagto_bd = RSC("CO_Plano_Pagamento")
		nu_parcelas_bd = RSC("NU_Parcelas")		
		vencimento_bd = RSC("DI_Prefencia")
		mes_inicio_parcelas = RSC("IN_Parcela")	
		resp_fin_prc = RSC("RP_Fina_Principal")
		resp_fin_alt = RSC("RP_Fina_Alter")	
		dia_util = RSC("TP_Vencimento")	
		proporcional = RSC("TP_Calculo")	

		reserva_vaga_bd = RSC("VA_Desconto_Reserva_Vaga")	
		anuidade_bd = RSC("VA_Desconto_Anuidade")
		dt_cancela_bd = RSC("DT_Cancela")
		
		if nu_parcelas_bd="" or isnull(nu_parcelas_bd) then
			nu_parcelas_bd =12 
		end if		
			
		if mes_inicio_parcelas="" or isnull(mes_inicio_parcelas) then
			mes_inicio_parcelas ="1/1/"&ano_contrato 
		end if	
		
		if dia_util="" or isnull(dia_util) then
			dia_util ="M"
		end if						
		
		if isnull(dt_contrato_bd) or dt_contrato_bd="" then
			dt_contrato = DatePart("d", now)&"/"&DatePart("m", now)&"/"&DatePart("yyyy", now)
		else
			dt_contrato = dt_contrato_bd
		end if	
		
		dados_data_contrato=split(dt_contrato,"/")
		dia_contrato = dados_data_contrato(0)
		if dia_contrato< 10 then
			dia_contrato = "0"&dia_contrato
		end if	
		mes_contrato = dados_data_contrato(1)
		if mes_contrato< 10 then
			mes_contrato = "0"&mes_contrato
		end if			
		dt_contrato = dia_contrato&"/"&mes_contrato&"/"&dados_data_contrato(2)
		
		if situac_contrato="A" then
			situac_contrato_nome="Ativo"		
		else
			situac_contrato_nome="Cancelado"			
		end if		
		
		if isnull(dt_cancela_bd) or dt_cancela_bd="" then
			data_cancelamento_padrao="0/0/0"	
		else
			data_cancelamento_padrao=dt_cancela_bd	
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
		
		Set RSB= Server.CreateObject("ADODB.Recordset")
		SQLB = "SELECT * FROM TB_Contrato_Bolsas where CO_Matricula ="&matricula_contrato&" AND NU_Ano_Letivo="&ano_contrato&" AND NU_Contrato = "&nu_contrato
		RSB.Open SQLB, CON5	
		
		if RSB.EOF then	
			bolsista="N&atilde;o"	
		else
			bolsista="Sim"			
		end if	
		
if plano_pagto_bd="" or isnull(plano_pagto_bd) then

			dados_inicio_parcelas=split(mes_inicio_parcelas,"/")
	
			if vencimento_bd< 10 then
				dia_parcela = "0"&vencimento_bd
			else
				dia_parcela	=vencimento_bd
			end if	
			mes_parcela= dados_inicio_parcelas(1)
			if mes_parcela< 10 then
				mes_parcela = "0"&mes_parcela
			end if			
			dt_inicio_parcelas = dia_parcela&"/"&mes_parcela&"/"&dados_inicio_parcelas(2)		
				
		else
			Set RSp = Server.CreateObject("ADODB.Recordset")
			SQLp = "SELECT * FROM TB_Plano_Pagamento WHERE NU_Ano_Letivo = "&ano_contrato&" AND CO_PlanoPG = '"&plano_pagto_bd&"'"
			RSp.Open SQLp, CON0
			nome_plano_pagto=RSp("NO_PlanoPG")	
			reserva_vaga_prmtro=RSp("VA_Reserva_Vaga")		
			anuidade_prmtro=RSp("VA_Anuidade")	
			vencimento_prmtro=RSp("DI_Vencimento")
			
			if vencimento_bd="" or isnull(vencimento_bd) then
				vencimento_bd =vencimento_prmtro 
			end if	
						
			dados_inicio_parcelas=split(mes_inicio_parcelas,"/")
	
			if vencimento_bd< 10 then
				dia_parcela = "0"&vencimento_bd
			else
				dia_parcela	=vencimento_bd
			end if	
			mes_parcela= dados_inicio_parcelas(1)
			if mes_parcela< 10 then
				mes_parcela = "0"&mes_parcela
			end if			
			dt_inicio_parcelas = dia_parcela&"/"&mes_parcela&"/"&dados_inicio_parcelas(2)		
				
			if isnull(reserva_vaga_bd) or reserva_vaga_bd=0 then
				if isnull(reserva_vaga_prmtro) or reserva_vaga_prmtro="" then
					reserva_vaga=reserva_vaga_prmtro
					val_reserva_vaga=reserva_vaga_prmtro					
				else
					val_reserva_vaga=reserva_vaga_prmtro
					reserva_vaga=Formatnumber(reserva_vaga_prmtro,2)		
				end if
			else
				val_reserva_vaga=reserva_vaga_bd
				reserva_vaga=Formatnumber(reserva_vaga_bd,2)		
			end if	

			if isnull(anuidade_bd) or anuidade_bd=0 then
				if isnull(anuidade_prmtro) or anuidade_prmtro="" then
					anuidade=anuidade_prmtro
					val_anuidade=anuidade_prmtro
					
				else
					val_anuidade=anuidade_prmtro				
					anuidade=Formatnumber(anuidade_prmtro,2)		
				end if	
			else
				val_anuidade=anuidade_bd						
				anuidade=Formatnumber(anuidade_bd,2)		
			end if			
			
		end if

		%>

<html>
<head>
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../../../../js/mm_menu.js"></script>
<script language="JavaScript" type="text/JavaScript">
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
						
								   
						 function altera_reserva_vaga(ano_pp,plano_pagto)
                                   {
								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=rv", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_rv  = oHTTPRequest.responseText;
resultado_rv = resultado_rv.replace(/\+/g," ")
resultado_rv = unescape(resultado_rv)
document.all.div_reserva_vaga.innerHTML ="R$"&resultado_rv
altera_anuidade(ano_pp,plano_pagto)
                                                         }
                                               }
                                               oHTTPRequest.send("ano_pub=" + ano_pp + "&pp_pub=" + plano_pagto);
                                   }
								   
						 function altera_anuidade(ano_pp,plano_pagto)
                                   {
								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=aa", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_aa  = oHTTPRequest.responseText;
resultado_aa = resultado_aa.replace(/\+/g," ")
resultado_aa = unescape(resultado_aa)
document.all.div_anuidade.innerHTML ="R$"&resultado_aa

                                                         }
                                               }
                                               oHTTPRequest.send("ano_pub=" + ano_pp + "&pp_pub=" + plano_pagto);
                                   }	
								   
						 function altera_qtd_parcelas(mes_inicio)
                                   {
								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=qp", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_qp  = oHTTPRequest.responseText;
resultado_qp = resultado_qp.replace(/\+/g," ")
resultado_qp = unescape(resultado_qp)
document.all.div_qtd_parcelas.innerHTML =resultado_qp

                                                         }
                                               }
                                               oHTTPRequest.send("q_pub=" + mes_inicio);
                                   }	
								   
								   
								   
								   
						 function calcula_parcela_referencia(proporcional,reserva_vaga,anuidade,nu_parcelas)
                                   {
								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=pr", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_pr  = oHTTPRequest.responseText;
resultado_pr = resultado_pr.replace(/\+/g," ")
resultado_pr = unescape(resultado_pr)
document.all.div_parcela_refcia.innerHTML =resultado_pr
//altera_proporcional(nu_parcelas)
//if (proporcional =='P') {
//	altera_qtd_parcelas(1)
//}

                                                         }
                                               }
                                               oHTTPRequest.send("p_pub=" + proporcional + "&r_pub=" + reserva_vaga + "&a_pub=" + anuidade +"&n_pub=" + nu_parcelas);
                                   }	
								   
						 function calcula_parcela_referencia_plano_pagto(proporcional,ano_pp,plano_pagto,nu_parcelas)
                                   {
								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=prpp", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_pr  = oHTTPRequest.responseText;
resultado_pr = resultado_pr.replace(/\+/g," ")
resultado_pr = unescape(resultado_pr)
document.all.div_parcela_refcia.innerHTML =resultado_pr
                                                         }
                                               }
                                               oHTTPRequest.send("p_pub=" + proporcional + "&ano_pub=" + ano_pp + "&pp_pub=" + plano_pagto +"&n_pub=" + nu_parcelas);
                                   }	
								   
						 function altera_proporcional(nu_parcelas)
                                   {
								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=pp", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_pp  = oHTTPRequest.responseText;
resultado_pp = resultado_pp.replace(/\+/g," ")
resultado_pp = unescape(resultado_pp)
document.all.div_proporcional.innerHTML =resultado_pp

                                                         }
                                               }
                                               oHTTPRequest.send("pp_pub=" + nu_parcelas);
                                   }								   
								   
								   
							 function CancelaContrato(situacao)
                                   {
								   
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=dc", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_dc  = oHTTPRequest.responseText;
resultado_dc = resultado_dc.replace(/\+/g," ")
resultado_dc = unescape(resultado_dc)
document.all.div_data_cancelamento.innerHTML =resultado_dc

                                                         }
                                               }
                                               oHTTPRequest.send("dc_pub=" + situacao);
                                   }		
								   
								   
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
//function checksubmit()
//{
//  if (document.busca.busca1.value != "" && document.busca.busca2.value != "")
//  {    alert("Por favor digite SOMENTE uma opção de busca!")
//    document.busca.busca1.focus()
//    return false
//  }
//    if (document.busca.busca1.value == "" && document.busca.busca2.value == "")
//  {    alert("Por favor digite uma opção de busca!")
//    document.busca.busca1.focus()
//    return false
//  }
//  return true
//}

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
//function ativa_campo(variavel){	
//	if (variavel=='bt_altera_ppagto') {
//		document.getElementById('bt_altera_ppagto').disabled   = false;
//	}
//}
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
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="calcula_parcela_referencia('<%response.Write(proporcional)%>','<%response.Write(val_reserva_vaga)%>','<%response.Write(val_anuidade)%>','<%response.Write(nu_parcelas_bd)%>');">
<form action="bd.asp?opt=c" method="post" name="busca" id="busca">
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
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
<%if opt="ok" then%>      
          <tr> 
            <td width="766" height="10" colspan="4" valign="top"> 
              <%call mensagens(nivel,9705,2,0) %>
            </td>
          </tr>
<%end if%>          
          <tr> 
            <td width="766" height="10" colspan="4" valign="top"> 
              <%call mensagens(nivel,9708,0,0) %>
            </td>
          </tr>
          <tr> 
            <td height="10" class="tb_tit"
> Alterar Contrato </td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr class="tb_subtit">
                <td width="100" align="center">Data</td>
                <td width="100" align="center">N&uacute;mero</td>
                <td width="100" align="center">Matr&iacute;cula</td>
                <td width="420" align="center">Aluno</td>
                <td width="50" align="center">Bolsista?</td>
                <td width="50" align="center">Un</td>
                <td width="50" align="center">Curso</td>
                <td width="50" align="center">Etapa</td>
                <td width="50" align="center">Turma</td>
                <td width="50" align="center">Situa&ccedil;&atilde;o</td>
              </tr>
              <tr>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(dt_contrato)%></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(concatena_contrato)%></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(matricula_contrato)%></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(nome_contrato)%></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(bolsista)%></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(unidade_contrato)%></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(curso_contrato)%></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(etapa_contrato)%></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(turma_contrato)%></td>
                <td align="center" class="<%response.Write(cor)%>"><%response.Write(situac_contrato_nome)%></td>
              </tr>
              <tr>
                <td colspan="10" align="center" class="<%response.Write(cor)%>"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="206" align="right"><input name="matricula" type="hidden" id="matricula" value="<%response.Write(matricula_contrato)%>"></td>
                    <td width="147"><input name="ano_contrato" type="hidden" id="ano_contrato" value="<%response.Write(ano_contrato)%>"></td>
                    <td width="153" align="left" class="form_dado_texto"><input name="contrato" type="hidden" id="contrato" value="<%response.Write(nu_contrato)%>"></td>
                    <td width="109">&nbsp;</td>
                    <td width="147">&nbsp;</td>
                    <td width="236">&nbsp;</td>
                  </tr>
                  <tr>
                    <td width="206" height="30" align="right" ><span class="form_dado_texto">Plano de Pagamento&nbsp;</span></td>
                    <td height="30" colspan="2">
                      <table border="0" cellpadding="0" cellspacing="0">
                        <tr>
  <td><div class="form_dado_texto" id="div_plano_pagto">	
    </div></td>
                          <td align="left" class="form_dado_texto">
                            <select name="plano_pagto" class="select_style" id="plano_pagto" onChange="altera_reserva_vaga('<%response.Write(ano_contrato)%>',this.value);calcula_parcela_referencia_plano_pagto(proporcional.value,<%response.Write(ano_contrato)%>,this.value,parcelas.value);">
                              <%
		Set RSpp = Server.CreateObject("ADODB.Recordset")
		SQLpp = "SELECT * FROM TB_Plano_Pagamento WHERE NU_Ano_Letivo = "&ano_contrato&" ORDER BY NO_PlanoPG"
		RSpp.Open SQLpp, CON0

		while not RSpp.EOF
			co_plano=RSpp("CO_PlanoPG")

			no_plano=RSpp("NO_PlanoPG")
			
			if plano_pagto_bd=co_plano then
				selected="SELECTED"
			else	
				selected=""
			end if		
			
			' Esse replace é necessário para que o javscript não retire o espaço dos códigos quando enviar para o AJAX
			co_plano=replace(co_plano," ","%20%")		
		%>
                              <option value="<%response.Write(co_plano)%>" <%response.Write(selected)%>>
                                <%response.Write(no_plano)%>
                                </option>
                              <%		RSpp.MOVENEXT
		WEND
%>
                              </select>&nbsp;</td>
  <!--                    <td>
                      <input name="button" type="button" class="botao_excluir" id="bt_altera_ppagto" value="Alterar Plano Pagto" disabled onClick="altera_ppagto('<%'response.Write(ano_contrato)%>',plano_pagto.value);">
                    </td> -->                   
                          </tr>
                        </table></td>
                    <td height="30" colspan="2" align="right"><strong><span class="form_dado_texto">Valor da Parcela de Refer&ecirc;ncia:</span></strong></td>
                    <td width="236" height="30"><div class="form_corpo" id="div_parcela_refcia"></div></td>
                  </tr>
                  <tr>
                    <td width="206" align="right"><span class="form_dado_texto">Reserva de Vaga&nbsp;</span></td>
                    <td width="147" class="form_dado_texto"><div id="div_reserva_vaga">R$
                      <input name="reserva_vaga" type="text" class="textInput" id="reserva_vaga" value="<%response.write(reserva_vaga)%>" size="15" autocomplete="off" onKeyUp="formataValor(this);" onBlur="calcula_parcela_referencia(proporcional.value,this.value,anuidade.value,parcelas.value);" >
                    </div></td>
                    <td width="153" align="right" class="form_dado_texto">Anuidade&nbsp;</td>
                    <td width="109" class="form_dado_texto"><div id="div_anuidade">R$<input name="anuidade" type="text" class="textInput" id="anuidade" value="<%response.write(anuidade)%>" size="15" onKeyUp="formataValor(this);"  onBlur="calcula_parcela_referencia(proporcional.value,reserva_vaga.value,this.value,parcelas.value);"></div></td>
                    <td width="147" align="right" class="form_dado_texto">Vencimento </td>
                    <td width="236"><select name="dia_vencimento" class="select_style" id="dia_vencimento">
                      <%for dv=1 to 31
					
					  	vencimento_bd=vencimento_bd*1
					  	if vencimento_bd=dv then
							selected="SELECTED"
						else	
							selected=""
						end if						
					%>
                      <option value="<%response.Write(dv)%>" <%response.Write(selected)%>>
                        <%response.Write(dv)%>
                        </option>
                      <%next%>
                    </select></td>
                    </tr>
                  <tr>
                    <td width="206" align="right" class="form_dado_texto">M&ecirc;s inicial das Parcelas&nbsp;</td>
                    <td width="147"><select name="mes_inicio" class="select_style" id="mes_inicio" onChange="altera_qtd_parcelas(this.value);">
                    <%for mi=1 to 12
						mes_inicio_parcelas=mes_inicio_parcelas*1
						select case mi
						case 1
							mes_nome="Janeiro"
							if mes_inicio_parcelas=1 then
								selected="selected"
							else
								selected=""							
							end if	
						case 2
							mes_nome="Fevereiro"
							if mes_inicio_parcelas=2 then
								selected="selected"
							else
								selected=""							
							end if								
						case 3
							mes_nome="Março"
							if mes_inicio_parcelas=3 then
								selected="selected"
							else
								selected=""							
							end if								
						case 4
							mes_nome="Abril"
							if mes_inicio_parcelas=4 then
								selected="selected"
							else
								selected=""							
							end if								
						case 5
							mes_nome="Maio"
							if mes_inicio_parcelas=5 then
								selected="selected"
							else
								selected=""							
							end if								
						case 6
							mes_nome="Junho"
							if mes_inicio_parcelas=6 then
								selected="selected"
							else
								selected=""							
							end if								
						case 7
							mes_nome="Julho"
							if mes_inicio_parcelas=7 then
								selected="selected"
							else
								selected=""							
							end if								
						case 8
							mes_nome="Agosto"
							if mes_inicio_parcelas=8 then
								selected="selected"
							else
								selected=""							
							end if								
						case 9
							mes_nome="Setembro"
							if mes_inicio_parcelas=9 then
								selected="selected"
							else
								selected=""							
							end if								
						case 10
							mes_nome="Outubro"
							if mes_inicio_parcelas=10 then
								selected="selected"
							else
								selected=""							
							end if								
						case 11
							mes_nome="Novembro"
							if mes_inicio_parcelas=11 then
								selected="selected"
							else
								selected=""							
							end if								
						case 12
							mes_nome="Dezembro"	
							if mes_inicio_parcelas=12 then
								selected="selected"
							else
								selected=""							
							end if								
						end Select																																									
					
					%>
                      <option value="<%response.Write(mi)%>" <%response.Write(selected)%>><%response.Write(mes_nome)%></option>
					<%next%>
                      </select></td>
                    <td width="153" align="right" class="form_dado_texto">Qtd de Parcelas&nbsp;</td>
                    <td width="109"><div id="div_qtd_parcelas"><select name="parcelas" class="select_style" id="parcelas" onChange="calcula_parcela_referencia(proporcional.value,reserva_vaga.value,anuidade.value,this.value);">
                      <%for p=1 to 12
					  	nu_parcelas_bd=nu_parcelas_bd*1
					  	if nu_parcelas_bd=p then
							selected="SELECTED"
						else	
							selected=""
						end if	
					  %>
                      <option value="<%response.Write(p)%>" <%response.Write(selected)%>>
                        <%response.Write(p)%>
                        </option>
                      <%next%>
                    </select></div></td>
                    <td width="147" align="right"><span class="form_dado_texto">Dia &Uacute;til</span></td>
                    <td width="236"><select name="dia_util" class="select_style" id="dia_util">
                      <% if dia_util = "M" then
					  		m_selected = "selected"
							a_selected=""
							p_selected=""
						elseif dia_util = "A" then
					  		m_selected = ""
							a_selected="selected"
							p_selected=""
						elseif dia_util = "P" then	
					  		m_selected = ""
							a_selected=""
							p_selected="selected"
						end if												
					  %>
                      <option value="M" <%response.Write(m_selected)%>>Mant&eacute;m Dia</option>
                      <option value="A" <%response.Write(a_selected)%>>Dia &Uacute;til Anterior</option>
                      <option value="P" <%response.Write(p_selected)%>>Dia &Uacute;til Posterior</option>
                    </select></td>
                  </tr>
                  <tr>
                    <td width="206" align="right" class="form_dado_texto">Respons&aacute;vel Financeiro Principal&nbsp;</td>
                    <td colspan="2">
                      <select name="rfp" class="select_style_fixo_1" id="rfp">
                        <% if resp_fin_prc = "" or isnull(resp_fin_prc) then%>		                      
                        <option value="nulo" selected></option>  
                        <%else%>
                        <option value="nulo"></option>
                        <%end if

		Set RSCONT = Server.CreateObject("ADODB.Recordset")
		SQLCONT = "SELECT * FROM TB_Contatos WHERE CO_Matricula ="& matricula_contrato&" ORDER BY TP_Contato"
		RSCONT.Open SQLCONT, CONCONT

		while not RSCONT.EOF
			tipo_contato=RSCONT("TP_Contato")
			nome_contato=RSCONT("NO_Contato")
			
			if resp_fin_prc=tipo_contato then
				selected = "selected"
			else
				selected = ""			
			end if	
		%>	
                        <option value="<%response.Write(tipo_contato)%>" <%response.Write(selected)%>><%response.Write(nome_contato&" ("&tipo_contato&")")%></option>	
                        <%		RSCONT.MOVENEXT
		WEND
%>
                        
                        </select>
                    </td>
                    <td colspan="2" align="right" class="form_dado_texto">Calcula Parcela Proporcional a 12 meses?&nbsp;</td>
                    <td width="236" class="form_dado_texto"><div id="div_proporcional"><select name="proporcional" class="select_style" id="proporcional" onChange="calcula_parcela_referencia(this.value,reserva_vaga.value,anuidade.value,parcelas.value);">
                      <% if proporcional = "P" or proporcional = "" or isnull(proporcional) then%>
                      <option value="P" selected>Sim</option>
                      <option value="A" >N&atilde;o</option>
                      <%else%>
                      <option value="P" >Sim</option>
                      <option value="A" selected>N&atilde;o</option>
                      <%end if%>
                    </select></div></td>
                  </tr>
                  <tr>
                    <td width="206" align="right" class="form_dado_texto">Respons&aacute;vel Financeiro Alternativo&nbsp;</td>
                    <td colspan="2"><select name="rfa" class="select_style_fixo_1" id="rfa">
                      <% if resp_fin_alt = "" or isnull(resp_fin_alt) then%>		                      
                      <option value="nulo" selected></option>  
                      <%else%>
                      <option value="nulo"></option>
                      <%end if                                                                            

		Set RSCONT = Server.CreateObject("ADODB.Recordset")
		SQLCONT = "SELECT * FROM TB_Contatos WHERE CO_Matricula ="& matricula_contrato&" ORDER BY TP_Contato"
		RSCONT.Open SQLCONT, CONCONT

		while not RSCONT.EOF
			tipo_contato=RSCONT("TP_Contato")
			nome_contato=RSCONT("NO_Contato")
			
			if resp_fin_alt=tipo_contato then
				selected = "selected"
			else
				selected = ""			
			end if				
		%>	
                      <option value="<%response.Write(tipo_contato)%>" <%response.Write(selected)%>><%response.Write(nome_contato&" ("&tipo_contato&")")%></option>	
                      <%		RSCONT.MOVENEXT
		WEND
%>                    
                    </select></td>
                    <td align="right">&nbsp;</td>
                    <td align="right"><span class="form_dado_texto">Data do Contrato&nbsp;</span></td>
                    <td width="236" class="form_dado_texto">
                    <select name="dia_contrato_frm" id="data_contrato_frm" class="select_style">
                      <% 
					  data_contrato=split(dt_contrato,"/")
							 For i =1 to 31
								data_contrato(0)=data_contrato(0)*1
								if data_contrato(0)=i then 
									if data_contrato(0)=0 then
										dd=""
									else
										if data_contrato(0)<10 then
											dd="0"&data_contrato(0)
										else
											dd=data_contrato(0)									
										end if
									end if	
									%>
									<option value="<%response.Write(i)%>" selected>
									<%response.Write(dd)%>
									</option>
								<% 
								else
									if i=0 then
										i_cod=""
									else
										if i<10 then									
											i_cod="0"&i
										else
											i_cod=i											
										end if
									end if	
								%>
                                    <option value="<%response.Write(i_cod)%>">
                                    <%response.Write(i_cod)%>
                                    </option>
								<% end if 
							next
							%>
                    </select>
                      /
                      
                      <select name="mes_contrato_frm" id="mes_contrato_frm" class="select_style">
                        <%data_contrato(1)=data_contrato(1)*1
							if data_contrato(1)=0 then%>
                        <option value="0" selected></option>
                        <% else%>
<!--                        <option value="0"></option>-->
                        <%end if						
								if data_contrato(1)=1 then%>
                        <option value="1" selected>janeiro</option>
                        <% else%>
                        <option value="1">janeiro</option>
                        <%end if
								if data_contrato(1)=2 then%>
                        <option value="2" selected>fevereiro</option>
                        <% else%>
                        <option value="2">fevereiro</option>
                        <%end if
								if data_contrato(1)=3 then%>
                        <option value="3" selected>mar&ccedil;o</option>
                        <% else%>
                        <option value="3">mar&ccedil;o</option>
                        <%end if
								if data_contrato(1)=4 then%>
                        <option value="4" selected>abril</option>
                        <% else%>
                        <option value="4">abril</option>
                        <%end if
								if data_contrato(1)=5 then%>
                        <option value="5" selected>maio</option>
                        <% else%>
                        <option value="5">maio</option>
                        <%end if
								if data_contrato(1)=6 then%>
                        <option value="6" selected>junho</option>
                        <% else%>
                        <option value="6">junho</option>
                        <%end if
								if data_contrato(1)=7 then%>
                        <option value="7" selected>julho</option>
                        <% else%>
                        <option value="7">julho</option>
                        <%end if%>
                        <%if data_contrato(1)=8 then%>
                        <option value="8" selected>agosto</option>
                        <% else%>
                        <option value="8">agosto</option>
                        <%end if
								if data_contrato(1)=9 then%>
                        <option value="9" selected>setembro</option>
                        <% else%>
                        <option value="9">setembro</option>
                        <%end if
								if data_contrato(1)=10 then%>
                        <option value="10" selected>outubro</option>
                        <% else%>
                        <option value="10">outubro</option>
                        <%end if
								if data_contrato(1)=11 then%>
                        <option value="11" selected>novembro</option>
                        <% else%>
                        <option value="11">novembro</option>
                        <%end if
								if data_contrato(1)=12 then%>
                        <option value="12" selected>dezembro</option>
                        <% else%>
                        <option value="12">dezembro</option>
                        <%end if%>
                      </select>
/
<select name="ano_contrato_frm" id="ano_contrato_frm" class="select_style">
<%						if data_contrato(2) = 0 then
 %>
                        <option value="0" SELECTED>
                          </option>
                        <%	
						else%>
<!--                        <option value="0">
                          </option>  -->                      
						<%				
						end if 
						For da =ano_letivo-1 to ano_letivo+1 
						data_contrato(2)=data_contrato(2)*1
							if da=data_contrato(2) then
								selected="selected"
							else
								selected=""
							end if		
 %>
                        <option value="<%Response.Write(da)%>" <%response.Write(selected)%>>
                          <%Response.Write(da)%>
                          </option>
                        <%NEXT%>
                      </select>                    </td>
                  </tr>
                  <tr>
                    <td width="206" align="right" class="form_dado_texto">&nbsp;</td>
                    <td width="147"><%'response.Write(dt_inicio_parcelas)%></td>
                    <td width="153" align="right" class="form_dado_texto">Situa&ccedil;&atilde;o&nbsp;</td>
                    <td width="109"><select name="situacao" class="select_style" id="situacao" onChange="CancelaContrato(this.value)">
                      <% 
                        if situac_contrato="A" then
                            ativo_contrato_select="Selected"	
                            canc_contrato_nome=""									
                        else
                            ativo_contrato_select=""							
                            canc_contrato_nome="Selected"			
                        end if		
					%>	                    
                      <option value="A" <%response.Write(ativo_contrato_select)%>>Ativo</option>
                      <option value="C" <%response.Write(canc_contrato_nome)%>>Cancelado</option>
                    </select></td>
                    <td align="right" class="form_dado_texto">Data de Cancelamento&nbsp;</td>
                    <td width="236"><div id="div_data_cancelamento">
                    <select name="dia_cancelamento" id="dia_cancelamento" class="select_style">
                      <% 
					  data_cancelamento=split(data_cancelamento_padrao,"/")
							 For i =0 to 31
								data_cancelamento(0)=data_cancelamento(0)*1
								if data_cancelamento(0)=i then 
									if data_cancelamento(0)=0 then
										dd=""
									else
										if data_cancelamento(0)<10 then
											dd="0"&data_cancelamento(0)
										else
											dd=data_cancelamento(0)									
										end if
									end if	
									%>
                      <option value="<%response.Write(i)%>" selected>
                        <%response.Write(dd)%>
                        </option>
                      <% 
								else
									if i=0 then
										i_cod=""
									else
										if i<10 then									
											i_cod="0"&i
										else
											i_cod=i											
										end if
									end if	
								%>
                      <option value="<%response.Write(i_cod)%>">
                        <%response.Write(i_cod)%>
                        </option>
                      <% end if 
							next
							%>
                    </select>
/
<select name="mes_cancelamento" id="mes_cancelamento" class="select_style">
  <%data_cancelamento(1)=data_cancelamento(1)*1
							if data_cancelamento(1)=0 then%>
  <option value="0" selected></option>
  <% else%>
  <!--                        <option value="0"></option>-->
  <%end if						
								if data_cancelamento(1)=1 then%>
  <option value="1" selected>janeiro</option>
  <% else%>
  <option value="1">janeiro</option>
  <%end if
								if data_cancelamento(1)=2 then%>
  <option value="2" selected>fevereiro</option>
  <% else%>
  <option value="2">fevereiro</option>
  <%end if
								if data_cancelamento(1)=3 then%>
  <option value="3" selected>mar&ccedil;o</option>
  <% else%>
  <option value="3">mar&ccedil;o</option>
  <%end if
								if data_cancelamento(1)=4 then%>
  <option value="4" selected>abril</option>
  <% else%>
  <option value="4">abril</option>
  <%end if
								if data_cancelamento(1)=5 then%>
  <option value="5" selected>maio</option>
  <% else%>
  <option value="5">maio</option>
  <%end if
								if data_cancelamento(1)=6 then%>
  <option value="6" selected>junho</option>
  <% else%>
  <option value="6">junho</option>
  <%end if
								if data_cancelamento(1)=7 then%>
  <option value="7" selected>julho</option>
  <% else%>
  <option value="7">julho</option>
  <%end if%>
  <%if data_cancelamento(1)=8 then%>
  <option value="8" selected>agosto</option>
  <% else%>
  <option value="8">agosto</option>
  <%end if
								if data_cancelamento(1)=9 then%>
  <option value="9" selected>setembro</option>
  <% else%>
  <option value="9">setembro</option>
  <%end if
								if data_cancelamento(1)=10 then%>
  <option value="10" selected>outubro</option>
  <% else%>
  <option value="10">outubro</option>
  <%end if
								if data_cancelamento(1)=11 then%>
  <option value="11" selected>novembro</option>
  <% else%>
  <option value="11">novembro</option>
  <%end if
								if data_cancelamento(1)=12 then%>
  <option value="12" selected>dezembro</option>
  <% else%>
  <option value="12">dezembro</option>
  <%end if%>
</select>
/
<select name="ano_cancelamento" id="ano_cancelamento" class="select_style">
  <%						if data_cancelamento(2) = 0 then
 %>
  <option value="0" SELECTED> </option>
  <%	
						else%>
  <!--                        <option value="0">
                          </option>  -->
  <%				
						end if 
						For da =ano_letivo-1 to ano_letivo+1 
						data_cancelamento(2)=data_cancelamento(2)*1
							if da=data_cancelamento(2) then
								selected="selected"
							else
								selected=""
							end if		
 %>
  <option value="<%Response.Write(da)%>" <%response.Write(selected)%>>
  <%Response.Write(da)%>
  </option>
  <%NEXT%>
</select></div></td>
                  </tr>
                </table></td>
                </tr>
              <% 		
		intrec=intrec+1
		check=check+1	
	RSC.MOVENEXT
	WEND
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
                        <input name="SUBMIT5" type=button class="botao_cancelar" onClick="MM_goToURL('parent','contratos.asp?pagina=1&v=s');return document.MM_returnValue" value="Voltar">
                    </div></td>
                    <td width="34%"> <div align="center"> </div> <div align="center"> </div></td>
                    <td width="33%"> <div align="center"> 
                        <input name="Submit" type="submit" class="botao_prosseguir" value="Confirmar">
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