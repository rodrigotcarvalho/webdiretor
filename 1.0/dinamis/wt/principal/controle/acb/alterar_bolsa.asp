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
ano_letivo = session("ano_letivo")
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
cancelados = Session("cancelados")		
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


Session("cod_form")=cod_form
Session("nome_form")=nome_form
Session("contrato")=contrato
Session("ativos")=ativos
Session("cancelados")=cancelados	
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
		
		Set RSa = Server.CreateObject("ADODB.Recordset")
		SQLa = "SELECT a.NO_Aluno as NOME, m.NU_Unidade as UNI, m.CO_Curso as CUR, m.CO_Etapa as ETA, m.CO_Turma as TUR FROM TB_Alunos a, TB_Matriculas m where a.CO_Matricula = "& matricula_contrato & " and a.CO_Matricula = m.CO_Matricula"
		RSa.Open SQLa, CONa		
	
		nome_cons = RSa("NOME")				

		Set RSC= Server.CreateObject("ADODB.Recordset")
		SQLC = "SELECT * FROM TB_Contrato c, TB_Contrato_Bolsas b where c.CO_Matricula = "& matricula_contrato & " and c.NU_Contrato = "&nu_contrato&" and c.NU_Ano_Letivo  = "& ano_contrato & "AND c.CO_Matricula = b.CO_Matricula and c.NU_Contrato = b.NU_Contrato AND c.NU_Ano_Letivo = b.NU_Ano_Letivo order by c.NU_Ano_Letivo Desc,c.NU_Contrato"
		RSC.Open SQLC, CON5
		
		if RSC.EOF then
			anula="S"
			anula_bolsa1 = "N"	
			anula_bolsa2 = "N"	
			anula_bolsa3 = "N"	
		else
			anula="N"
			aplicacao=RSC("AP_Bolsa")
			b1_bolsa=RSC("CO_Bolsa1")
			b1_desconto=RSC("VA_Desconto1")
			b1_vl_inic=RSC("VL_Inicio1")
			b1_vl_fim=RSC("VL_Fim1")
			b1_pc_inic=RSC("PC_Inicio1")
			b1_pc_fim=RSC("PC_Fim1")
			b1_ap_bolsa=RSC("AP_Bolsa1")
			b1_dt_conce=RSC("DT_Concessao1")
			b1_ob_bolsa=RSC("OB_Bolsa1")
			b2_bolsa=RSC("CO_Bolsa2")
			b2_desconto=RSC("VA_Desconto2")
			b2_vl_inic=RSC("VL_Inicio2")
			b2_vl_fim=RSC("VL_Fim2")
			b2_pc_inic=RSC("PC_Inicio2")
			b2_pc_fim=RSC("PC_Fim2")
			b2_ap_bolsa=RSC("AP_Bolsa2")
			b2_dt_conce=RSC("DT_Concessao2")
			b2_ob_bolsa=RSC("OB_Bolsa2")
			b3_bolsa=RSC("CO_Bolsa3")
			b3_desconto=RSC("VA_Desconto3")
			b3_vl_inic=RSC("VL_Inicio3")
			b3_vl_fim=RSC("VL_Fim3")
			b3_pc_inic=RSC("PC_Inicio3")
			b3_pc_fim=RSC("PC_Fim3")
			b3_ap_bolsa=RSC("AP_Bolsa3")
			b3_dt_conce=RSC("DT_Concessao3")
			b3_ob_bolsa=RSC("OB_Bolsa3")
			usuario_bolsa=RSC("CO_Usuario")
			
			if isnull(b1_bolsa) or b1_bolsa= "" then
				anula_bolsa1 = "S"
				b1_desconto = ""				
			else
				anula_bolsa1 = "N"						
			end if	
			
			if isnull(b2_bolsa) or b2_bolsa= "" then
				anula_bolsa2 = "S"
				b2_desconto = ""				
			else
				anula_bolsa2 = "N"										
			end if	
			
			if isnull(b3_bolsa) or b1_bolsa= "" then
				anula_bolsa3 = "S"
				b3_desconto = ""						
			else
				anula_bolsa3 = "N"								
			end if												
		end if				

hoje_ano = DatePart("yyyy", now) 
hoje_mes = DatePart("m", now) 
hoje_dia = DatePart("d", now) 

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
						
						
						 function recuperarTpDesconto(dTipo, bolsa)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=td", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_d  = oHTTPRequest.responseText;
resultado_d = resultado_d.replace(/\+/g," ")
resultado_d = unescape(resultado_d)
if (bolsa =="b1") {
	document.all.b1_tp_desconto.innerHTML =resultado_d
} else if (bolsa =="b2") {
	document.all.b2_tp_desconto.innerHTML =resultado_d
} else if (bolsa =="b3") {
	document.all.b3_tp_desconto.innerHTML =resultado_d	
}
                                                           }
                                               }
                                               oHTTPRequest.send("d_pub=" + dTipo + "&b_pub=" + bolsa);
                                   }

						 function recuperarValDesconto(vTipo, bolsa)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=vd", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_v  = oHTTPRequest.responseText;
resultado_v = resultado_v.replace(/\+/g," ")
resultado_v = unescape(resultado_v)
if (bolsa =="b1") {
	document.all.b1_val_desconto.innerHTML =resultado_v
} else if (bolsa =="b2") {
	document.all.b2_val_desconto.innerHTML =resultado_v
} else if (bolsa =="b3") {
	document.all.b3_val_desconto.innerHTML =resultado_v
}
recuperarIncidencia(vTipo, bolsa)

                                                           }
                                               }
                                               oHTTPRequest.send("v_pub=" + vTipo + "&b_pub=" + bolsa);
                                   }


			 function recuperarIncidencia(iTipo, bolsa)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=ri", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_v  = oHTTPRequest.responseText;
resultado_v = resultado_v.replace(/\+/g," ")
resultado_v = unescape(resultado_v)
if (bolsa =="b1") {
	document.all.aplica_b1.innerHTML =resultado_v
} else if (bolsa =="b2") {
	document.all.aplica_b2.innerHTML =resultado_v
} else if (bolsa =="b3") {
	document.all.aplica_b3.innerHTML =resultado_v
}
recuperarDataInicio(iTipo, bolsa)

                                                           }
                                               }
                                               oHTTPRequest.send("i_pub=" + iTipo + "&b_pub=" + bolsa);
                                   }

			 function recuperarDataInicio(iTipo, bolsa)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=di", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_v  = oHTTPRequest.responseText;
resultado_v = resultado_v.replace(/\+/g," ")
resultado_v = unescape(resultado_v)
if (bolsa =="b1") {
	document.all.div_data_inicio_b1.innerHTML =resultado_v
} else if (bolsa =="b2") {
	document.all.div_data_inicio_b2.innerHTML =resultado_v
} else if (bolsa =="b3") {
	document.all.div_data_inicio_b3.innerHTML =resultado_v
}
recuperarDataFim(iTipo, bolsa)

                                                           }
                                               }
                                               oHTTPRequest.send("i_pub=" + iTipo + "&b_pub=" + bolsa);
                                   }
			 function recuperarDataFim(iTipo, bolsa)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=df", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                       if (oHTTPRequest.readyState==4){
                                                                    var resultado_v  = oHTTPRequest.responseText;
resultado_v = resultado_v.replace(/\+/g," ")
resultado_v = unescape(resultado_v)
if (bolsa =="b1") {
	document.all.div_data_fim_b1.innerHTML =resultado_v
} else if (bolsa =="b2") {
	document.all.div_data_fim_b2.innerHTML =resultado_v
} else if (bolsa =="b3") {
	document.all.div_data_fim_b3.innerHTML =resultado_v
}
recuperarDataConcessao(bolsa)

                                                           }
                                               }
                                               oHTTPRequest.send("i_pub=" + iTipo + "&b_pub=" + bolsa);
                                   }
			 function recuperarDataConcessao(bolsa)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=dcb", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                       if (oHTTPRequest.readyState==4){
                                                                    var resultado_dc  = oHTTPRequest.responseText;
resultado_dc = resultado_dc.replace(/\+/g," ")
resultado_dc = unescape(resultado_dc)
if (bolsa =="b1") {
	document.all.div_dt_concessao_b1.innerHTML =resultado_dc;
	b1_habilita_campo();
} else if (bolsa =="b2") {
	document.all.div_dt_concessao_b2.innerHTML =resultado_dc;
	b2_habilita_campo();	
} else if (bolsa =="b3") {
	document.all.div_dt_concessao_b3.innerHTML =resultado_dc;
	b3_habilita_campo();	
}
limpa_prazo(bolsa,'q')


                                                           }
                                               }
                                               oHTTPRequest.send("b_pub=" + bolsa);
                                   }
								   

			 function recuperarAplicacaoBolsa(aplicacaoBolsa,bolsa1,bolsa2,bolsa3)
                                   {
                                               var oHTTPRequest = createXMLHTTP();
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=ab", true);
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
                                               oHTTPRequest.onreadystatechange=function() {
                                                           if (oHTTPRequest.readyState==4){
                                                                    var resultado_ab  = oHTTPRequest.responseText;
resultado_ab = resultado_ab.replace(/\+/g," ")
resultado_ab = unescape(resultado_ab)
if ((aplicacaoBolsa =="B") || (bolsa1=="nulo" && bolsa2=="nulo" && bolsa3=="nulo")){
	document.all.div_aplicacao.innerHTML =resultado_ab
}

                                                           }
                                               }
                                               oHTTPRequest.send("ab_pub=" + aplicacaoBolsa + "&b1_pub=" + bolsa1 + "&b2_pub=" + bolsa2 + "&b3_pub=" + bolsa3);
                                   }								   

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

function apaga_bolsa(bolsa){
	if (bolsa == 'b1') {
		document.getElementById('b1_tipo_bolsa').value   = 'nulo';	
		document.getElementById('b1_tipo_desconto').value   = 0;
		document.getElementById('b1_desconto').value   = '';
		document.getElementById('b1_prazo').value   = 'S';				
		document.getElementById('b1_dia_de').value   = 0;
		document.getElementById('b1_dia_ate').value   = 0;	
		document.getElementById('b1_mes_de').value   = 0;
		document.getElementById('b1_mes_ate').value   = 0;
		document.getElementById('b1_ano_de').value   = 0;
		document.getElementById('b1_ano_ate').value   = 0;			
		document.getElementById('b1_pi').value   = '';
		document.getElementById('b1_pf').value   = '';
		document.getElementById('b1_aplica_bolsa').value   = 'nulo';	
		document.getElementById('b1_observacao').value   = '';	
	} else if (bolsa == 'b2') {
		document.getElementById('b2_tipo_bolsa').value   = 'nulo';	
		document.getElementById('b2_tipo_desconto').value   = 0;
		document.getElementById('b2_desconto').value   = '';
		document.getElementById('b2_prazo').value   = 'S';				
		document.getElementById('b2_dia_de').value   = 0;
		document.getElementById('b2_dia_ate').value   = 0;	
		document.getElementById('b2_mes_de').value   = 0;
		document.getElementById('b2_mes_ate').value   = 0;
		document.getElementById('b2_ano_de').value   = 0;
		document.getElementById('b2_ano_ate').value   = 0;			
		document.getElementById('b2_pi').value   = '';
		document.getElementById('b2_pf').value   = '';
		document.getElementById('b2_aplica_bolsa').value   = 'nulo';	
		document.getElementById('b2_observacao').value   = '';	
	} else {
		document.getElementById('b3_tipo_bolsa').value   = 'nulo';	
		document.getElementById('b3_tipo_desconto').value   = 0;
		document.getElementById('b3_desconto').value   = '';
		document.getElementById('b3_prazo').value   = 'S';				
		document.getElementById('b3_dia_de').value   = 0;
		document.getElementById('b3_dia_ate').value   = 0;	
		document.getElementById('b3_mes_de').value   = 0;
		document.getElementById('b3_mes_ate').value   = 0;
		document.getElementById('b3_ano_de').value   = 0;
		document.getElementById('b3_ano_ate').value   = 0;			
		document.getElementById('b3_pi').value   = '';
		document.getElementById('b3_pf').value   = '';
		document.getElementById('b3_aplica_bolsa').value   = 'nulo';	
		document.getElementById('b3_observacao').value   = '';	
	}	
}
function b1_bloqueia(){	
		document.getElementById('b1_tipo_desconto').disabled   = true;
		document.getElementById('b1_desconto').disabled   = true;
		document.getElementById('b1_prazo').disabled   = true;				
		document.getElementById('b1_dia_de').disabled   = true;
		document.getElementById('b1_dia_ate').disabled   = true;	
		document.getElementById('b1_mes_de').disabled   = true;
		document.getElementById('b1_mes_ate').disabled   = true;
		document.getElementById('b1_ano_de').disabled   = true;
		document.getElementById('b1_ano_ate').disabled   = true;			
		document.getElementById('b1_pi').disabled   = true;
		document.getElementById('b1_pf').disabled   = true;
		document.getElementById('b1_aplica_bolsa').disabled   = true;	
		document.getElementById('b1_observacao').disabled   = true;			
}

function b2_bloqueia(){	
		document.getElementById('b2_tipo_desconto').disabled   = true;
		document.getElementById('b2_desconto').disabled   = true;	
		document.getElementById('b2_prazo').disabled   = true;				
		document.getElementById('b2_dia_de').disabled   = true;
		document.getElementById('b2_dia_ate').disabled   = true;	
		document.getElementById('b2_mes_de').disabled   = true;
		document.getElementById('b2_mes_ate').disabled   = true;
		document.getElementById('b2_ano_de').disabled   = true;
		document.getElementById('b2_ano_ate').disabled   = true;			
		document.getElementById('b2_pi').disabled   = true;
		document.getElementById('b2_pf').disabled   = true;
		document.getElementById('b2_aplica_bolsa').disabled   = true;	
		document.getElementById('b2_observacao').disabled   = true;			
}

function b3_bloqueia(){	
		document.getElementById('b3_tipo_desconto').disabled   = true;
		document.getElementById('b3_desconto').disabled   = true;	
		document.getElementById('b3_prazo').disabled   = true;				
		document.getElementById('b3_dia_de').disabled   = true;
		document.getElementById('b3_dia_ate').disabled   = true;	
		document.getElementById('b3_mes_de').disabled   = true;
		document.getElementById('b3_mes_ate').disabled   = true;
		document.getElementById('b3_ano_de').disabled   = true;
		document.getElementById('b3_ano_ate').disabled   = true;			
		document.getElementById('b3_pi').disabled   = true;
		document.getElementById('b3_pf').disabled   = true;
		document.getElementById('b3_aplica_bolsa').disabled   = true;	
		document.getElementById('b3_observacao').disabled   = true;			
}

function b1_desbloqueia(){	
		document.getElementById('b1_tipo_desconto').disabled   = false;
		document.getElementById('b1_desconto').disabled   = false;	
		document.getElementById('b1_prazo').disabled   = false;				
		document.getElementById('b1_dia_de').disabled   = false;
		document.getElementById('b1_dia_ate').disabled   = false;	
		document.getElementById('b1_mes_de').disabled   = false;
		document.getElementById('b1_mes_ate').disabled   = false;
		document.getElementById('b1_ano_de').disabled   = false;
		document.getElementById('b1_ano_ate').disabled   = false;			
		document.getElementById('b1_aplica_bolsa').disabled   = false;	
		document.getElementById('b1_observacao').disabled   = false;			
}

function b2_desbloqueia(){	
		document.getElementById('b2_tipo_desconto').disabled   = false;
		document.getElementById('b2_desconto').disabled   = false;	
		document.getElementById('b2_prazo').disabled   = false;			
		document.getElementById('b2_dia_de').disabled   = false;
		document.getElementById('b2_dia_ate').disabled   = false;	
		document.getElementById('b2_mes_de').disabled   = false;
		document.getElementById('b2_mes_ate').disabled   = false;
		document.getElementById('b2_ano_de').disabled   = false;
		document.getElementById('b2_ano_ate').disabled   = false;			
		document.getElementById('b2_aplica_bolsa').disabled   = false;	
		document.getElementById('b2_observacao').disabled   = false;			
}

function b3_desbloqueia(){	
		document.getElementById('b3_tipo_desconto').disabled   = false;
		document.getElementById('b3_desconto').disabled   = false;	
		document.getElementById('b3_prazo').disabled   = false;			
		document.getElementById('b3_dia_de').disabled   = false;
		document.getElementById('b3_dia_ate').disabled   = false;	
		document.getElementById('b3_mes_de').disabled   = false;
		document.getElementById('b3_mes_ate').disabled   = false;
		document.getElementById('b3_ano_de').disabled   = false;
		document.getElementById('b3_ano_ate').disabled   = false;			
		document.getElementById('b3_aplica_bolsa').disabled   = false;	
		document.getElementById('b3_observacao').disabled   = false;			
}

function b1_habilita_campo(){
		document.getElementById('b1_dia_de').disabled   = false;
		document.getElementById('b1_dia_ate').disabled   = false;	
		document.getElementById('b1_mes_de').disabled   = false;
		document.getElementById('b1_mes_ate').disabled   = false;
		document.getElementById('b1_ano_de').disabled   = false;
		document.getElementById('b1_ano_ate').disabled   = false;			
		document.getElementById('b1_pi').disabled   = true;
		document.getElementById('b1_pf').disabled   = true;	
}
		
function b1_desabilita_campo(){	   
		document.getElementById('b1_dia_de').disabled   = true;
		document.getElementById('b1_dia_ate').disabled   = true;	
		document.getElementById('b1_mes_de').disabled   = true;
		document.getElementById('b1_mes_ate').disabled   = true;
		document.getElementById('b1_ano_de').disabled   = true;
		document.getElementById('b1_ano_ate').disabled   = true;	
		document.getElementById('b1_pi').disabled   = false;
		document.getElementById('b1_pf').disabled   = false;			
}
function b2_habilita_campo(){
		document.getElementById('b2_dia_de').disabled   = false;
		document.getElementById('b2_dia_ate').disabled   = false;	
		document.getElementById('b2_mes_de').disabled   = false;
		document.getElementById('b2_mes_ate').disabled   = false;
		document.getElementById('b2_ano_de').disabled   = false;
		document.getElementById('b2_ano_ate').disabled   = false;			
		document.getElementById('b2_pi').disabled   = true;
		document.getElementById('b2_pf').disabled   = true;	
}

function b2_desabilita_campo(){	   
		document.getElementById('b2_dia_de').disabled   = true;
		document.getElementById('b2_dia_ate').disabled   = true;	
		document.getElementById('b2_mes_de').disabled   = true;
		document.getElementById('b2_mes_ate').disabled   = true;
		document.getElementById('b2_ano_de').disabled   = true;
		document.getElementById('b2_ano_ate').disabled   = true;	
		document.getElementById('b2_pi').disabled   = false;
		document.getElementById('b2_pf').disabled   = false;			
}
function b3_habilita_campo(){
		document.getElementById('b3_dia_de').disabled   = false;
		document.getElementById('b3_dia_ate').disabled   = false;	
		document.getElementById('b3_mes_de').disabled   = false;
		document.getElementById('b3_mes_ate').disabled   = false;
		document.getElementById('b3_ano_de').disabled   = false;
		document.getElementById('b3_ano_ate').disabled   = false;			
		document.getElementById('b3_pi').disabled   = true;
		document.getElementById('b3_pf').disabled   = true;	
}

function b3_desabilita_campo(){	   
		document.getElementById('b3_dia_de').disabled   = true;
		document.getElementById('b3_dia_ate').disabled   = true;	
		document.getElementById('b3_mes_de').disabled   = true;
		document.getElementById('b3_mes_ate').disabled   = true;
		document.getElementById('b3_ano_de').disabled   = true;
		document.getElementById('b3_ano_ate').disabled   = true;	
		document.getElementById('b3_pi').disabled   = false;
		document.getElementById('b3_pf').disabled   = false;			
}

function limpa_prazo(bolsa,limpa){	 

	if (limpa=='d') {
		if (bolsa=='b1'){
			document.getElementById('b1_prazoS').checked = false;
			document.getElementById('b1_prazoN').checked = true;				
			document.getElementById('b1_dia_de').value   = 0;
			document.getElementById('b1_dia_ate').value   = 0;	
			document.getElementById('b1_mes_de').value   = 0;
			document.getElementById('b1_mes_ate').value   = 0;
			document.getElementById('b1_ano_de').value   = 0;
			document.getElementById('b1_ano_ate').value   = 0;					
		} else if (bolsa=='b2'){
			document.getElementById('b2_prazoS').checked = false;
			document.getElementById('b2_prazoN').checked = true;				
			document.getElementById('b2_dia_de').value   = 0;
			document.getElementById('b2_dia_ate').value   = 0;	
			document.getElementById('b2_mes_de').value   = 0;
			document.getElementById('b2_mes_ate').value   = 0;
			document.getElementById('b2_ano_de').value   = 0;
			document.getElementById('b2_ano_ate').value   = 0;			
		} else {
			document.getElementById('b3_prazoS').checked = false;
			document.getElementById('b3_prazoN').checked = true;				
			document.getElementById('b3_dia_de').value   = 0;
			document.getElementById('b3_dia_ate').value   = 0;	
			document.getElementById('b3_mes_de').value   = 0;
			document.getElementById('b3_mes_ate').value   = 0;
			document.getElementById('b3_ano_de').value   = 0;
			document.getElementById('b3_ano_ate').value   = 0;			
		}
	}else if (limpa=='q') {
		if (bolsa=='b1'){		
			document.getElementById('b1_prazoS').checked = true;
			document.getElementById('b1_prazoN').checked = false;			
			document.getElementById('b1_dia_de').value   = <%response.Write(hoje_dia)%>;
			document.getElementById('b1_dia_ate').value   = 31;	
			document.getElementById('b1_mes_de').value   = <%response.Write(hoje_mes)%>;
			document.getElementById('b1_mes_ate').value   = 12;
			document.getElementById('b1_ano_de').value   = <%response.Write(ano_letivo)%>;
			document.getElementById('b1_ano_ate').value   = <%response.Write(ano_letivo)%>;		
			document.getElementById('b1_pi').value   = '';
			document.getElementById('b1_pf').value   = '';			
		} else if (bolsa=='b2'){
			document.getElementById('b2_prazoS').checked = true;
			document.getElementById('b2_prazoN').checked = false;				
			document.getElementById('b2_dia_de').value   = <%response.Write(hoje_dia)%>;
			document.getElementById('b2_dia_ate').value   = 31;	
			document.getElementById('b2_mes_de').value   = <%response.Write(hoje_mes)%>;
			document.getElementById('b2_mes_ate').value   = 12;
			document.getElementById('b2_ano_de').value   = <%response.Write(ano_letivo)%>;
			document.getElementById('b2_ano_ate').value   = <%response.Write(ano_letivo)%>;		
			document.getElementById('b2_pi').value   = '';
			document.getElementById('b2_pf').value   = '';			
		} else {
			document.getElementById('b3_prazoS').checked = true;
			document.getElementById('b3_prazoN').checked = false;				
			document.getElementById('b3_dia_de').value   = <%response.Write(hoje_dia)%>;
			document.getElementById('b3_dia_ate').value   = 31;	
			document.getElementById('b3_mes_de').value   = <%response.Write(hoje_mes)%>;
			document.getElementById('b3_mes_ate').value   = 12;
			document.getElementById('b3_ano_de').value   = <%response.Write(ano_letivo)%>;
			document.getElementById('b3_ano_ate').value   = <%response.Write(ano_letivo)%>;		
			document.getElementById('b3_pi').value   = '';
			document.getElementById('b3_pf').value   = '';			
		}	
			
	}
}
<% 
if anula_bolsa1 = "S" then
		onload = "b1_bloqueia()" 
		b1_prazo_data = "S"
		b1_prazo = "N"			
else
	if 	(b1_pc_inic=0 and  b1_pc_fim=0) or (isnull(b1_pc_inic) and isnull(b1_pc_fim)) then
		onload = "b1_habilita_campo()" 
		b1_prazo_data = "S"
		b1_prazo = "N"			
	else
		b1_prazo_data = "N"
		b1_prazo = "S"
		onload = "b1_desabilita_campo()" 
	end if
end if

if anula_bolsa2 = "S" then
		onload = onload&"; b2_bloqueia()" 
		b2_prazo_data = "S"
		b2_prazo = "N"			
else
	if 	b2_pc_inic=0 and  b2_pc_fim=0 or (isnull(b2_pc_inic) and isnull(b2_pc_fim)) then
		onload = onload&"; b2_habilita_campo()" 
		b2_prazo_data = "S"
		b2_prazo = "N"		
	else
		b2_prazo_data = "N"
		b2_prazo = "S"
		onload = onload&"; b2_desabilita_campo()" 
	end if
end if	

if anula_bolsa3 = "S" then
		onload = onload&"; b3_bloqueia()" 
		b3_prazo_data = "S"
		b3_prazo = "N"			
else
	if 	b3_pc_inic=0 and  b3_pc_fim=0 or (isnull(b3_pc_inic) and isnull(b3_pc_fim)) then
		onload = onload&"; b3_habilita_campo()" 
		b3_prazo_data = "S"
		b3_prazo = "N"		
	else
		b3_prazo_data = "N"
		b3_prazo = "S"
		onload = onload&"; b3_desabilita_campo()" 
	end if
end if	

%>
function limitText(limitField, limitCount, limitNum) {
	if (limitField.value.length > limitNum) {
		limitField.value = limitField.value.substring(0, limitNum);
	} else {
		limitCount.value = limitNum - limitField.value.length;
	}
}

//-->
</script>

</head>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="<% response.Write(onload)%>">
<form action="bd.asp?opt=b" method="post" name="busca" id="busca">
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
> Alterar Bolsa de Estudo </td>
          </tr>
          <tr>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td colspan="2" class="form_dado_texto"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr class="tb_subtit">
                    <td width="20%" align="center"> Matr&iacute;cula</td>
                    <td width="60%" align="center">Nome</td>
                    <td align="center">N&uacute;mero do Contrato</td>
                  </tr>
                  <tr class="form_dado_texto">
                    <td width="20%" align="center"><%response.Write(matricula_contrato)%><input name="matricula" type="hidden" id="matricula" value="<%response.Write(matricula_contrato)%>"><input name="ano_contrato" type="hidden" id="ano_contrato" value="<%response.Write(ano_contrato)%>"><input name="contrato" type="hidden" id="contrato" value="<%response.Write(nu_contrato)%>"></td>
                    <td width="60%" align="center"><%response.Write(nome_cons)%></td>
                    <td align="center"><%response.Write(ano_contrato)%>/<%
'					if nu_contrato<100000 then
'						if nu_contrato<10000 then
'							if nu_contrato<1000 then
'								if nu_contrato<100 then
'									if nu_contrato<10 then
'										nu_contrato="00000"&nu_contrato							
'									else
'										nu_contrato="0000"&nu_contrato					
'									end if						
'								else
'									nu_contrato="000"&nu_contrato					
'								end if	
'							else
'								nu_contrato="00"&nu_contrato					
'							end if
'						else
'							nu_contrato="0"&nu_contrato					
'						end if
'					end if	 
					
					response.Write(nu_contrato)%></td>
                  </tr>
                </table></td>
                </tr>
              <tr>
                <td colspan="2" class="form_dado_texto">&nbsp;</td>
                </tr>
              <tr>
                <td width="12%" class="form_dado_texto"> Aplica&ccedil;&atilde;o das Bolsas </td>
                <td width="88%"><div id="div_aplicacao"><select name="aplicacao_bolsa" id="aplicacao_bolsa" class="select_style">
                <% if anula="S" then
						z_select = ""				
						s_select = ""
						c_select = ""
						m_select = ""										
						b_select = "Selected"
					elseif aplicacao = "S" then
						z_select = ""					
						s_select = "Selected"
						c_select = ""
						m_select = ""										
						b_select = ""
					elseif aplicacao = "C" then
						z_select = ""					
						s_select = ""
						c_select = "Selected"
						m_select = ""										
						b_select = ""
					elseif aplicacao = "M" then
						z_select = ""					
						s_select = ""
						c_select = ""
						m_select = "Selected"										
						b_select = ""
					else
						z_select = "Selected"
						s_select = ""
						c_select = ""
						m_select = ""										
						b_select = ""						
					end if															
					%>
<!--                  <option value="nulo" <%response.Write(z_select)%>></option>                    
-->                  <option value="S" <%response.Write(s_select)%>>Somar</option>
                  <option value="C" <%response.Write(c_select)%>>Cascata</option>
                  <option value="M" <%response.Write(m_select)%>>Maior</option>
                  <option value="B" <%response.Write(b_select)%>>N&atilde;o Possui Bolsas</option>
                </select></div></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td class="tb_subtit">Bolsas Concedidas</td>
          </tr>
          <tr>
            <td><strong class="form_dado_texto">Bolsa 1</strong></td>
          </tr>
          <tr>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="15%" height="25" align="right"><span class="form_dado_texto"> Tipo de Bolsa </span></td>
                <td width="20%" height="25" align="left"><select name="b1_tipo_bolsa" class="select_style" onChange="b1_desbloqueia();recuperarTpDesconto(this.value,'b1');recuperarValDesconto(this.value,'b1');recuperarAplicacaoBolsa(aplicacao_bolsa.value,this.value,b2_tipo_bolsa.value,b3_tipo_bolsa.value);">
                  <option value="nulo" selected></option>                
                <%	
				Set RS = Server.CreateObject("ADODB.Recordset")
				SQL = "SELECT * FROM TB_Tipo_Bolsa order by NO_Bolsa"
				RS.Open SQL, CON0

								
				while not RS.EOF
					co_bolsa=RS("CO_Bolsa")
					no_bolsa=RS("NO_Bolsa")
	
					if b1_bolsa=co_bolsa then
						b1_bolsa_select = "SELECTED"
						b1_tp_desconto = RS("TP_Desconto")
						if b1_tp_desconto = "V" then
							b1_tp_desconto_nulo_select = ""							
							b1_tp_desconto_valor_select = "Selected"
							b1_tp_desconto_percent_select = ""									
						elseif b1_tp_desconto = "P" then
							b1_tp_desconto_nulo_select = ""									
							b1_tp_desconto_valor_select = ""						
							b1_tp_desconto_percent_select = "Selected"															
						else
							b1_tp_desconto_nulo_select = "Selected"									
							b1_tp_desconto_valor_select = ""						
							b1_tp_desconto_percent_select = ""																
						end if
					else				
						b1_bolsa_select = ""					
					end if	
				%>
				  <option value="<%response.Write(co_bolsa)%>"<%response.Write(b1_bolsa_select)%>><%response.Write(no_bolsa)%></option>
				<%
				RS.MOVENEXT
				WEND	
				%>  
                </select>
                  <label>
                    <input name="apagar_b1" type="button" class="botao_apagar" id="apagar_b1" value="Apagar Bolsa" onClick="apaga_bolsa('b1');recuperarAplicacaoBolsa(aplicacao_bolsa.value,b1_tipo_bolsa.value,b2_tipo_bolsa.value,b3_tipo_bolsa.value);">
                  </label></td>
                <td width="15%" height="25" align="right"><span class="form_dado_texto">Tipo de Desconto: </span></td>
                <td width="20%" height="25" align="left"><span class="form_dado_texto"><div id="b1_tp_desconto">
                <select name="b1_tipo_desconto" id="b1_tipo_desconto" class="select_style" >
                  <option value="nulo" <%response.Write(b1_tp_desconto_nulo_select)%>></option>
                 <option value="P"<%response.Write(b1_tp_desconto_percent_select)%>>Percentual</option>                  
                 <option value="V"<%response.Write(b1_tp_desconto_valor_select)%>>Valor</option>
                </select>                 
                  
                  </div></span></td>
                <td width="12%" height="25" align="right"><span class="form_dado_texto"> Desconto </span></td>
                <td width="20%" height="25" align="left"><span class="form_dado_texto"><div id="b1_val_desconto">
                    <input name="b1_desconto" type="text" class="textInput" id="b1_desconto" size="10" maxlength="8" value="<%response.Write(b1_desconto)%>" onFocus="this.select()">
                </div></span></td>
              </tr>
              <tr>
                <td width="15%" height="25" align="right"><span class="form_dado_texto">Validade da Bolsa  </span></td>
                <td colspan="3" rowspan="2" align="right" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="2%" rowspan="2" valign="top"><% if b1_prazo_data="S" then%>
                      <input name="b1_prazo" type="radio" id="b1_prazoS" value="s" onClick="javascript:b1_habilita_campo();limpa_prazo('b1','q');" checked>
                      <%else%>
                      <input name="b1_prazo" type="radio" id="b1_prazoS" value="s" onClick="javascript:b1_habilita_campo();limpa_prazo('b1','q');">
                      <%end if%></td>
                    <td width="12%" height="25" align="right" class="form_dado_texto">Data inicial </td>
                    <td width="30%" height="25" class="form_dado_texto"><div id="div_data_inicio_b1">
                    <%
					if b1_vl_inic = "" or isnull(b1_vl_inic) or anula_bolsa1 = "S" then
						b1_vl_inic = "0/0/0"
					end if	
					if b1_vl_fim = "" or isnull(b1_vl_fim) or anula_bolsa1 = "S" then
						b1_vl_fim = "0/0/0"
					end if						
						b1_data_de=split(b1_vl_inic,"/")
						b1_data_ate=split(b1_vl_fim,"/")
					%>
                    
                    
                    <select name="b1_dia_de" id="b1_dia_de" class="select_style">
                      <% 
							 For i =0 to 31
								b1_data_de(0)=b1_data_de(0)*1
								if b1_data_de(0)=i then 
									if b1_data_de(0)=0 then
										dd=""
									else
										if b1_data_de(0)<10 then
											dd="0"&b1_data_de(0)
										else
											dd=b1_data_de(0)									
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
                      
                      <select name="b1_mes_de" id="b1_mes_de" class="select_style">
                        <%b1_data_de(1)=b1_data_de(1)*1
							if b1_data_de(1)=0 then%>
                        <option value="0" selected></option>
                        <% else%>
                        <option value="0"></option>
                        <%end if						
								if b1_data_de(1)=1 then%>
                        <option value="1" selected>janeiro</option>
                        <% else%>
                        <option value="1">janeiro</option>
                        <%end if
								if b1_data_de(1)=2 then%>
                        <option value="2" selected>fevereiro</option>
                        <% else%>
                        <option value="2">fevereiro</option>
                        <%end if
								if b1_data_de(1)=3 then%>
                        <option value="3" selected>mar&ccedil;o</option>
                        <% else%>
                        <option value="3">mar&ccedil;o</option>
                        <%end if
								if b1_data_de(1)=4 then%>
                        <option value="4" selected>abril</option>
                        <% else%>
                        <option value="4">abril</option>
                        <%end if
								if b1_data_de(1)=5 then%>
                        <option value="5" selected>maio</option>
                        <% else%>
                        <option value="5">maio</option>
                        <%end if
								if b1_data_de(1)=6 then%>
                        <option value="6" selected>junho</option>
                        <% else%>
                        <option value="6">junho</option>
                        <%end if
								if b1_data_de(1)=7 then%>
                        <option value="7" selected>julho</option>
                        <% else%>
                        <option value="7">julho</option>
                        <%end if%>
                        <%if b1_data_de(1)=8 then%>
                        <option value="8" selected>agosto</option>
                        <% else%>
                        <option value="8">agosto</option>
                        <%end if
								if b1_data_de(1)=9 then%>
                        <option value="9" selected>setembro</option>
                        <% else%>
                        <option value="9">setembro</option>
                        <%end if
								if b1_data_de(1)=10 then%>
                        <option value="10" selected>outubro</option>
                        <% else%>
                        <option value="10">outubro</option>
                        <%end if
								if b1_data_de(1)=11 then%>
                        <option value="11" selected>novembro</option>
                        <% else%>
                        <option value="11">novembro</option>
                        <%end if
								if b1_data_de(1)=12 then%>
                        <option value="12" selected>dezembro</option>
                        <% else%>
                        <option value="12">dezembro</option>
                        <%end if%>
                      </select>
                      /
                      <select name="b1_ano_de" id="b1_ano_de" class="select_style">
<%						if b1_data_ate(2) = 0 then
 %>
                        <option value="0" SELECTED>
                          </option>
                        <%	
						else%>
                        <option value="0">
                          </option>                        
						<%				
						end if 
						For ald =ano_letivo-1 to ano_letivo+1 
						b1_data_de(2)=b1_data_de(2)*1
							if ald=b1_data_de(2) then
								selected="selected"
							else
								selected=""
							end if		
 %>
                        <option value="<%Response.Write(ald)%>" <%response.Write(selected)%>>
                          <%Response.Write(ald)%>
                          </option>
                        <%NEXT%>
                      </select></div></td>
                    <td width="2%" height="25" rowspan="2" valign="top" class="form_dado_texto"><% if b1_prazo="S" then%>
                      <input name="b1_prazo" type="radio" id="b1_prazoN" value="n" onClick="javascript:b1_desabilita_campo();limpa_prazo('b1','d');" checked>
                      <%else%>
                      <input name="b1_prazo" type="radio" id="b1_prazoN" value="n" onClick="javascript:b1_desabilita_campo();limpa_prazo('b1','d');">
                      <%end if%></td>
                    <td width="12%" height="25" align="right" class="form_dado_texto">Parcela Inicial  </td>
                    <td width="20%" height="25" class="form_dado_texto"><input name="b1_pi" type="text" class="textInput" id="b1_pi" size="4" maxlength="3" value="<%response.Write(b1_pc_inic)%>"></td>
                  </tr>
                  <tr>
                    <td width="12%" align="right" class="form_dado_texto">Data Final </td>
                    <td width="30%" class="form_dado_texto"><div id="div_data_fim_b1">
                      <select name="b1_dia_ate" id="b1_dia_ate" class="select_style">
                        <% 
							 For i =0 to 31
							 b1_data_ate(0)=b1_data_ate(0)*1
							 if i=b1_data_ate(0) then 
							 	if b1_data_ate(0)=0 then
									da=""
								else
									if b1_data_ate(0)<10 then
										da="0"&b1_data_ate(0)
									else
										da=b1_data_ate(0)									
									end if
								end if								 
							 %>
                        <option value="<%response.Write(i)%>" selected>
                          <%response.Write(da)%>
                          </option>
                        <% else
					  		if i=0 then
								i_cod=""
							else
							  	i_cod=i
								if i<10 then
								
								i="0"&i
								end if
							end if	
							%>
                        <option value="<%response.Write(i_cod)%>">
                          <%response.Write(i)%>
                          </option>
                        <% end if 
							next
							%>
                      </select>
                      /
<select name="b1_mes_ate" id="b1_mes_ate" class="select_style">
                        <%b1_data_ate(1)=b1_data_ate(1)*1
								if b1_data_ate(1)=0 then%>
                        <option value="0" selected></option>
                        <% else%>
                        <option value="0"></option>
                        <%end if						
								if b1_data_ate(1)=1 then%>
                        <option value="1" selected>janeiro</option>
                        <% else%>
                        <option value="1">janeiro</option>
                        <%end if
								if b1_data_ate(1)=2 then%>
                        <option value="2" selected>fevereiro</option>
                        <% else%>
                        <option value="2">fevereiro</option>
                        <%end if
								if b1_data_ate(1)=3 then%>
                        <option value="3" selected>mar&ccedil;o</option>
                        <% else%>
                        <option value="3">mar&ccedil;o</option>
                        <%end if
								if b1_data_ate(1)=4 then%>
                        <option value="4" selected>abril</option>
                        <% else%>
                        <option value="4">abril</option>
                        <%end if
								if b1_data_ate(1)=5 then%>
                        <option value="5" selected>maio</option>
                        <% else%>
                        <option value="5">maio</option>
                        <%end if
								if b1_data_ate(1)=6 then%>
                        <option value="6" selected>junho</option>
                        <% else%>
                        <option value="6">junho</option>
                        <%end if
								if b1_data_ate(1)=7 then%>
                        <option value="7" selected>julho</option>
                        <% else%>
                        <option value="7">julho</option>
                        <%end if%>
                        <%if b1_data_ate(1)=8 then%>
                        <option value="8" selected>agosto</option>
                        <% else%>
                        <option value="8">agosto</option>
                        <%end if
								if b1_data_ate(1)=9 then%>
                        <option value="9" selected>setembro</option>
                        <% else%>
                        <option value="9">setembro</option>
                        <%end if
								if b1_data_ate(1)=10 then%>
                        <option value="10" selected>outubro</option>
                        <% else%>
                        <option value="10">outubro</option>
                        <%end if
								if b1_data_ate(1)=11 then%>
                        <option value="11" selected>novembro</option>
                        <% else%>
                        <option value="11">novembro</option>
                        <%end if
								if b1_data_ate(1)=12 then%>
                        <option value="12" selected>dezembro</option>
                        <% else%>
                        <option value="12">dezembro</option>
                        <%end if%>
                      </select>

                      </select>
                      /
                      
                      <select name="b1_ano_ate" id="b1_ano_ate" class="select_style">
                        <% 
						if b1_data_ate(2) = 0 then
 %>
                        <option value="0" selected> </option>
                        <%	
						else%>
                        <option value="0"> </option>
                        <%					
						end if
						For ala =ano_letivo-1 to ano_letivo+1 
							b1_data_ate(2)=b1_data_ate(2)*1						
							if ala=b1_data_ate(2) then
								selected="selected"
							else
								selected=""
							end if		
 %>
                        <option value="<%Response.Write(ala)%>" <%response.Write(selected)%>>
                          <%Response.Write(ala)%>
                          </option>
                        <%NEXT%>
                      </select></div>
                      </td>
                    <td width="12%" align="right" class="form_dado_texto">Parcela Final </td>
                    <td width="20%" class="form_dado_texto"><input name="b1_pf" type="text" class="textInput" id="b1_pf" size="4" maxlength="3" value="<%response.Write(b1_pc_fim)%>"></td>
                  </tr>
                </table></td>
                <td width="12%" align="right" class="form_dado_texto">Incid&ecirc;ncia </td>
                <td width="20%" align="left"><div id="aplica_b1"><select name="b1_aplica_bolsa" id="b1_aplica_bolsa" class="select_style">
                <% if anula="S" or anula_bolsa1 = "S" then
						z_select = "Selected"				
						s_select = ""
						p_select = ""									
						a_select = ""
					elseif b1_ap_bolsa = "P" then
						z_select = ""					
						s_select = ""
						p_select = "Selected"
						a_select = ""							
					elseif b1_ap_bolsa = "S" then
						z_select = ""					
						s_select = "Selected"
						p_select = ""
						a_select = ""										
					elseif b1_ap_bolsa = "A" then
						z_select = ""
						s_select = ""
						p_select = ""	
						a_select = "Selected"												
					end if															
					%>
                  <option value="nulo" <%response.Write(z_select)%>></option>  
                  <option value="P" <%response.Write(p_select)%>>Parcela da Anuidade</option>                                    
                  <option value="S" <%response.Write(s_select)%>>Serviços</option>
                  <option value="M" <%response.Write(a_select)%>>Ambos</option>
                </select></div></td>
                </tr>
              <tr>
                <td width="15%" height="25" align="right">&nbsp;</td>
                <td width="12%" height="25" align="right"><input type="hidden" name="countdown" size="3" value="255"></td>
                <td width="20%" height="25" align="right">&nbsp;</td>
                </tr>
              <tr class="form_dado_texto">
                <td width="15%" align="right" valign="top">Observa&ccedil;&atilde;o  </td>
                <td colspan="3" align="left">
                  <textarea name="b1_observacao" cols="115" rows="3" class="textInput" id="b1_observacao" onKeyDown="limitText(this.form.b1_observacao,this.form.countdown,255);" onKeyUp="limitText(this.form.b1_observacao,this.form.countdown,255);"><%response.write(b1_ob_bolsa)%></textarea></td>
                <td width="12%" align="right">Data Concess&atilde;o</td>
                <td width="20%" align="left"><div id="div_dt_concessao_b1"><%response.write("&nbsp;"&b1_dt_conce)%></div><input name="b1_dt_conce" type="hidden" value="<%response.write(b1_dt_conce)%>"></td>
              </tr>
            </table></td>
          </tr>
          <tr> 
            <td><hr></td>
          </tr>
          <tr>
            <td><strong class="form_dado_texto">Bolsa 2</strong></td>
          </tr>
          <tr>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="15%" height="25" align="right"><span class="form_dado_texto"> Tipo de Bolsa </span></td>
                <td width="20%" height="25" align="left"><select name="b2_tipo_bolsa" class="select_style" onChange="b2_desbloqueia();recuperarTpDesconto(this.value,'b2');recuperarValDesconto(this.value,'b2');recuperarAplicacaoBolsa(aplicacao_bolsa.value,b1_tipo_bolsa.value,this.value,b3_tipo_bolsa.value);">
                  <option value="nulo" selected></option>
                  <%	
				Set RS = Server.CreateObject("ADODB.Recordset")
				SQL = "SELECT * FROM TB_Tipo_Bolsa order by NO_Bolsa"
				RS.Open SQL, CON0	
								
				while not RS.EOF
					co_bolsa=RS("CO_Bolsa")
					no_bolsa=RS("NO_Bolsa")
				
					if b2_bolsa=co_bolsa then
						b2_bolsa_select = "SELECTED"
						b2_tp_desconto = RS("TP_Desconto")
						if b2_tp_desconto = "V" then
							b2_tp_desconto_nulo_select = ""							
							b2_tp_desconto_valor_select = "Selected"
							b2_tp_desconto_percent_select = ""									
						elseif b2_tp_desconto = "P" then
							b2_tp_desconto_nulo_select = ""									
							b2_tp_desconto_valor_select = ""						
							b2_tp_desconto_percent_select = "Selected"															
						else
							b2_tp_desconto_nulo_select = "Selected"									
							b2_tp_desconto_valor_select = ""						
							b2_tp_desconto_percent_select = ""																
						end if
					else				
						b2_bolsa_select = ""					
					end if	
				%>
				  <option value="<%response.Write(co_bolsa)%>"<%response.Write(b2_bolsa_select)%>>
					<%response.Write(no_bolsa)%>
					</option>
				  <%
				RS.MOVENEXT
				WEND		
				%>
                </select>
                  <input name="apagar_b2" type="button" class="botao_apagar" id="apagar_b2" value="Apagar Bolsa" onClick="apaga_bolsa('b2');recuperarAplicacaoBolsa(aplicacao_bolsa.value,b1_tipo_bolsa.value,b2_tipo_bolsa.value,b3_tipo_bolsa.value);"></td>
                <td width="15%" height="25" align="right"><span class="form_dado_texto">Tipo de Desconto: </span></td>
                <td width="20%" height="25" align="left"><span class="form_dado_texto">
                <div id="b2_tp_desconto">
                <select name="b2_tipo_desconto" id="b2_tipo_desconto" class="select_style" >
                  <option value="nulo" <%response.Write(b2_tp_desconto_nulo_select)%>></option>
                 <option value="P" <%response.Write(b2_tp_desconto_percent_select)%>>Percentual</option>                  
                 <option value="V" <%response.Write(b2_tp_desconto_valor_select)%>>Valor</option>
                </select>  </div></span></td>
                <td width="12%" height="25" align="right"><span class="form_dado_texto"> Desconto </span></td>
                <td width="20%" height="25" align="left"><span class="form_dado_texto"><div id="b2_val_desconto">
                    <input name="b2_desconto" type="text" class="textInput" id="b2_desconto" size="10" maxlength="8" value="<%response.Write(b2_desconto)%>" onFocus="this.select()">
                </div></span></td>
              </tr>
              <tr>
                <td width="15%" height="25" align="right"><span class="form_dado_texto">Validade da Bolsa  </span></td>
                <td colspan="3" rowspan="2" align="right" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="2%" rowspan="2" valign="top"><% if b2_prazo_data="S" then%>
                      <input name="b2_prazo" type="radio" id="b2_prazoS" value="s" onClick="javascript:b2_habilita_campo();limpa_prazo('b2','q');" checked>
                      <%else%>
                      <input name="b2_prazo" type="radio" id="b2_prazoS" value="s" onClick="javascript:b2_habilita_campo();limpa_prazo('b2','q');">
                      <%end if%>
                      </td>
                    <td width="12%" height="25" align="right" class="form_dado_texto">Data inicial </td>
                    <td width="30%" height="25" class="form_dado_texto"><div id="div_data_inicio_b2">
                    <%
					if b2_vl_inic = "" or isnull(b2_vl_inic) or anula_bolsa2 = "S"	then
						b2_vl_inic = "0/0/0"
					end if	
					if b2_vl_fim = "" or isnull(b2_vl_fim) or anula_bolsa2 = "S" then
						b2_vl_fim = "0/0/0"
					end if							
					
						b2_data_de=split(b2_vl_inic,"/")
						b2_data_ate=split(b2_vl_fim,"/")
					%>
                    
                    
                    <select name="b2_dia_de" id="b2_dia_de" class="select_style">
                      <% 
							 For i =0 to 31
							 b2_data_de(0)=b2_data_de(0)*1
							 if b2_data_de(0)=i then 
							 	if b2_data_de(0)=0 then
									dd=""
								else
									if b2_data_de(0)<10 then
										dd="0"&b2_data_de(0)
									else
										dd=b2_data_de(0)									
									end if
								end if	
							 %>
                      <option value="<%response.Write(i)%>" selected>
                        <%response.Write(dd)%>
                        </option>
                      <% else
					  		if i=0 then
								i_cod=""
							else
							  	i_cod=i
								if i<10 then
								
								i="0"&i
								end if
							end if	
							%>
                      <option value="<%response.Write(i_cod)%>">
                        <%response.Write(i)%>
                        </option>
                      <% end if 
							next
							%>
                    </select>
                      /
                      <select name="b2_mes_de" id="b2_mes_de" class="select_style">
                        <%b2_data_de(1)=b2_data_de(1)*1
							if b2_data_de(1)=0 then%>
                        <option value="0" selected></option>
                        <% else%>
                        <option value="0"></option>
                        <%end if	
								if b2_data_de(1)=1 then%>
                        <option value="1" selected>janeiro</option>
                        <% else%>
                        <option value="1">janeiro</option>
                        <%end if
								if b2_data_de(1)=2 then%>
                        <option value="2" selected>fevereiro</option>
                        <% else%>
                        <option value="2">fevereiro</option>
                        <%end if
								if b2_data_de(1)=3 then%>
                        <option value="3" selected>mar&ccedil;o</option>
                        <% else%>
                        <option value="3">mar&ccedil;o</option>
                        <%end if
								if b2_data_de(1)=4 then%>
                        <option value="4" selected>abril</option>
                        <% else%>
                        <option value="4">abril</option>
                        <%end if
								if b2_data_de(1)=5 then%>
                        <option value="5" selected>maio</option>
                        <% else%>
                        <option value="5">maio</option>
                        <%end if
								if b2_data_de(1)=6 then%>
                        <option value="6" selected>junho</option>
                        <% else%>
                        <option value="6">junho</option>
                        <%end if
								if b2_data_de(1)=7 then%>
                        <option value="7" selected>julho</option>
                        <% else%>
                        <option value="7">julho</option>
                        <%end if%>
                        <%if b2_data_de(1)=8 then%>
                        <option value="8" selected>agosto</option>
                        <% else%>
                        <option value="8">agosto</option>
                        <%end if
								if b2_data_de(1)=9 then%>
                        <option value="9" selected>setembro</option>
                        <% else%>
                        <option value="9">setembro</option>
                        <%end if
								if b2_data_de(1)=10 then%>
                        <option value="10" selected>outubro</option>
                        <% else%>
                        <option value="10">outubro</option>
                        <%end if
								if b2_data_de(1)=11 then%>
                        <option value="11" selected>novembro</option>
                        <% else%>
                        <option value="11">novembro</option>
                        <%end if
								if b2_data_de(1)=12 then%>
                        <option value="12" selected>dezembro</option>
                        <% else%>
                        <option value="12">dezembro</option>
                        <%end if%>
                      </select>
                      /
                      <select name="b2_ano_de" id="b2_ano_de" class="select_style">
<%						if b2_data_ate(2) = 0 then
 %>
                        <option value="0" SELECTED>
                          </option>
                        <%	
						else%>
                        <option value="0">
                          </option>                        
						<%						
						end if                      
						For ald =ano_letivo-1 to ano_letivo+1
							b2_data_de(2)=b2_data_de(2)*1
							if ald=b2_data_de(2) then
								selected="selected"
							else
								selected=""
							end if		
 %>
                        <option value="<%Response.Write(ald)%>" <%response.Write(selected)%>>
                          <%Response.Write(ald)%>
                          </option>
                        <%NEXT%>
                      </select></div></td>
                    <td width="2%" height="25" rowspan="2" valign="top" class="form_dado_texto"><% if b2_prazo="S" then%>
                      <input name="b2_prazo" type="radio" id="b2_prazoN" value="n" onClick="javascript:b2_desabilita_campo();limpa_prazo('b2','d');" checked>
                      <%else%>
                      <input name="b2_prazo" type="radio" id="b2_prazoN" value="n" onClick="javascript:b2_desabilita_campo();limpa_prazo('b2','d');">
                      <%end if%></td>
                    <td width="12%" height="25" align="right" class="form_dado_texto">Parcela Inicial  </td>
                    <td width="20%" height="25" class="form_dado_texto"><input name="b2_pi" type="text" class="textInput" id="b2_pi" size="4" maxlength="3" value="<%response.Write(b2_pc_inic)%>"></td>
                  </tr>
                  <tr>
                    <td width="12%" align="right" class="form_dado_texto">Data Final </td>
                    <td width="30%" class="form_dado_texto"><div id="div_data_fim_b2">
                      <select name="b2_dia_ate" id="b2_dia_ate" class="select_style">
                        <% 
							 For i =0 to 31
							 b2_data_ate(0)=b2_data_ate(0)*1
							 if i=b2_data_ate(0) then 
							 	if b2_data_ate(0) =0 then
									da=""
								else
									if b2_data_ate(0)<10 then
										da="0"&b2_data_ate(0)
									else
										da=b2_data_ate(0)									
									end if
								end if	
							 %>
                        <option value="<%response.Write(i)%>" selected>
                          <%response.Write(da)%>
                          </option>
                        <% else
								if i=0 then
									i_cod=""
								else
									i_cod=i
									if i<10 then
									
									i="0"&i
									end if
								end if	
							%>
                        <option value="<%response.Write(i_cod)%>">
                          <%response.Write(i)%>
                          </option>
                        <% end if 
							next
							%>
                      </select>
/
<select name="b2_mes_ate" id="b2_mes_ate" class="select_style">
                        <%b2_data_ate(1)=b2_data_ate(1)*1
								if b2_data_ate(1)=0 then%>
                        <option value="0" selected></option>
                        <% else%>
                        <option value="0"></option>
                        <%end if	
								if b2_data_ate(1)=1 then%>
                        <option value="1" selected>janeiro</option>
                        <% else%>
                        <option value="1">janeiro</option>
                        <%end if
								if b2_data_ate(1)=2 then%>
                        <option value="2" selected>fevereiro</option>
                        <% else%>
                        <option value="2">fevereiro</option>
                        <%end if
								if b2_data_ate(1)=3 then%>
                        <option value="3" selected>mar&ccedil;o</option>
                        <% else%>
                        <option value="3">mar&ccedil;o</option>
                        <%end if
								if b2_data_ate(1)=4 then%>
                        <option value="4" selected>abril</option>
                        <% else%>
                        <option value="4">abril</option>
                        <%end if
								if b2_data_ate(1)=5 then%>
                        <option value="5" selected>maio</option>
                        <% else%>
                        <option value="5">maio</option>
                        <%end if
								if b2_data_ate(1)=6 then%>
                        <option value="6" selected>junho</option>
                        <% else%>
                        <option value="6">junho</option>
                        <%end if
								if b2_data_ate(1)=7 then%>
                        <option value="7" selected>julho</option>
                        <% else%>
                        <option value="7">julho</option>
                        <%end if%>
                        <%if b2_data_ate(1)=8 then%>
                        <option value="8" selected>agosto</option>
                        <% else%>
                        <option value="8">agosto</option>
                        <%end if
								if b2_data_ate(1)=9 then%>
                        <option value="9" selected>setembro</option>
                        <% else%>
                        <option value="9">setembro</option>
                        <%end if
								if b2_data_ate(1)=10 then%>
                        <option value="10" selected>outubro</option>
                        <% else%>
                        <option value="10">outubro</option>
                        <%end if
								if b2_data_ate(1)=11 then%>
                        <option value="11" selected>novembro</option>
                        <% else%>
                        <option value="11">novembro</option>
                        <%end if
								if b2_data_ate(1)=12 then%>
                        <option value="12" selected>dezembro</option>
                        <% else%>
                        <option value="12">dezembro</option>
                        <%end if%>
                      </select>

                      </select>
                      /                      
                      <select name="b2_ano_ate" id="b2_ano_ate" class="select_style">
                        <% 
						if b2_data_ate(2) = 0 then
 %>
                        <option value="0" selected> </option>
                        <%	
						else%>
                        <option value="0"> </option>
                        <%					
						end if                      
                        For ala =ano_letivo-1 to ano_letivo+1 
						b2_data_ate(2) =b2_data_ate(2)*1
						if ala=b2_data_ate(2) then
							selected="selected"
						else
							selected=""
						end if		
 %>
                        <option value="<%Response.Write(ala)%>" <%response.Write(selected)%>>
                          <%Response.Write(ala)%>
                          </option>
                        <%NEXT%>
                      </select></div></td>
                    <td width="12%" align="right" class="form_dado_texto">Parcela Final </td>
                    <td width="20%" class="form_dado_texto"><input name="b2_pf" type="text" class="textInput" id="b2_pf" size="4" maxlength="3" value="<%response.Write(b2_pc_fim)%>"></td>
                  </tr>
                </table></td>
                <td width="12%" align="right" class="form_dado_texto">Incid&ecirc;ncia </td>
                <td width="20%" align="left"><div id="aplica_b2"><select name="b2_aplica_bolsa" id="b2_aplica_bolsa" class="select_style">
                <% if anula="S" or anula_bolsa2 = "S" then
						z_select = "Selected"				
						s_select = ""
						p_select = ""									
						a_select = ""
					elseif b2_ap_bolsa = "P" then
						z_select = ""					
						s_select = ""
						p_select = "Selected"
						a_select = ""							
					elseif b2_ap_bolsa = "S" then
						z_select = ""					
						s_select = "Selected"
						p_select = ""
						a_select = ""										
					elseif b2_ap_bolsa = "A" then
						z_select = ""
						s_select = ""
						p_select = ""	
						a_select = "Selected"												
					end if															
					%>
                  <option value="nulo" <%response.Write(z_select)%>></option>  
                  <option value="P" <%response.Write(p_select)%>>Parcela da Anuidade</option>                                    
                  <option value="S" <%response.Write(s_select)%>>Serviços</option>
                  <option value="M" <%response.Write(a_select)%>>Ambos</option>
                </select></div></td>
                </tr>
              <tr>
                <td width="15%" height="25" align="right">&nbsp;</td>
                <td width="12%" height="25" align="right">&nbsp;</td>
                <td width="20%" height="25" align="right">&nbsp;</td>
                </tr>
              <tr class="form_dado_texto">
                <td width="15%" align="right" valign="top">Observa&ccedil;&atilde;o  </td>
                <td colspan="3" align="left"><label>
                  <textarea name="b2_observacao" cols="115" rows="3" class="textInput" id="b2_observacao" onKeyDown="limitText(this.form.b2_observacao,this.form.countdown,255);" onKeyUp="limitText(this.form.b2_observacao,this.form.countdown,255);"><%response.write(b2_ob_bolsa)%></textarea>
                </label></td>
                <td width="12%" align="right">Data Concess&atilde;o</td>
                <td width="20%" align="left"><div id="div_dt_concessao_b2"><%response.write("&nbsp;"&b2_dt_conce)%></div><input name="b2_dt_conce" type="hidden" value="<%response.write(b2_dt_conce)%>"></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td><hr></td>
          </tr>
          <tr>
            <td><strong class="form_dado_texto">Bolsa 3</strong></td>
          </tr>
          <tr>
            <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="15%" height="25" align="right"><span class="form_dado_texto"> Tipo de Bolsa </span></td>
                <td width="20%" height="25" align="left"><select name="b3_tipo_bolsa" class="select_style" onChange="b3_desbloqueia();recuperarTpDesconto(this.value,'b3');recuperarValDesconto(this.value,'b3');recuperarAplicacaoBolsa(aplicacao_bolsa.value,b1_tipo_bolsa.value,b2_tipo_bolsa.value,this.value);">
                  <option value="nulo" selected></option>                
                <%	
				Set RS = Server.CreateObject("ADODB.Recordset")
				SQL = "SELECT * FROM TB_Tipo_Bolsa order by NO_Bolsa"
				RS.Open SQL, CON0	
				

				while not RS.EOF
					co_bolsa=RS("CO_Bolsa")
					no_bolsa=RS("NO_Bolsa")
				
					if b3_bolsa=co_bolsa then
						b3_bolsa_select = "SELECTED"
						b3_tp_desconto = RS("TP_Desconto")
						if b3_tp_desconto = "V" then
							b3_tp_desconto_nulo_select = ""							
							b3_tp_desconto_valor_select = "Selected"
							b3_tp_desconto_percent_select = ""									
						elseif b3_tp_desconto = "P" then
							b3_tp_desconto_nulo_select = ""									
							b3_tp_desconto_valor_select = ""						
							b3_tp_desconto_percent_select = "Selected"															
						else
							b3_tp_desconto_nulo_select = "Selected"									
							b3_tp_desconto_valor_select = ""						
							b3_tp_desconto_percent_select = ""																
						end if
					else						
						b3_bolsa_select = ""					
					end if	
				%>
				  <option value="<%response.Write(co_bolsa)%>"<%response.Write(b3_bolsa_select)%>><%response.Write(no_bolsa)%></option>
				<%
				RS.MOVENEXT
				WEND
				%>  
                </select>
                  <input name="apagar_b3" type="button" class="botao_apagar" id="apagar_b3" value="Apagar Bolsa" onClick="apaga_bolsa('b3');recuperarAplicacaoBolsa(aplicacao_bolsa.value,b1_tipo_bolsa.value,b2_tipo_bolsa.value,b3_tipo_bolsa.value);"></td>
                <td width="15%" height="25" align="right"><span class="form_dado_texto">Tipo de Desconto: </span></td>
                <td width="20%" height="25" align="left"><span class="form_dado_texto"><div id="b3_tp_desconto"> 
                <select name="b3_tipo_desconto" id="b3_tipo_desconto"  class="select_style" >
                  <option value="nulo" <%response.Write(b3_tp_desconto_nulo_select)%>></option>
                 <option value="P" <%response.Write(b3_tp_desconto_percent_select)%>>Percentual</option>                  
                 <option value="V" <%response.Write(b3_tp_desconto_valor_select)%>>Valor</option>
                </select>  </div></span></td>
                <td width="12%" height="25" align="right"><span class="form_dado_texto"> Desconto </span></td>
                <td width="20%" height="25" align="left"><span class="form_dado_texto">
                <div id="b3_val_desconto">
                    <input name="b3_desconto" type="text" class="textInput" id="b3_desconto" size="10" maxlength="8" value="<%response.Write(b3_desconto)%>" onFocus="this.select()">
                </div></span></td>
              </tr>
              <tr>
                <td width="15%" height="25" align="right"><span class="form_dado_texto">Validade da Bolsa  </span></td>
                <td colspan="3" rowspan="2" align="right" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="2%" rowspan="2" valign="top"><% if b3_prazo_data="S" then%>
                      <input name="b3_prazo" type="radio" id="b3_prazoS" value="s" onClick="javascript:b3_habilita_campo();limpa_prazo('b3','q');" checked>
                      <%else%>
                      <input name="b3_prazo" type="radio" id="b3_prazoS" value="s" onClick="javascript:b3_habilita_campo();limpa_prazo('b3','q');">
                      <%end if%></td>
                    <td width="12%" height="25" align="right" class="form_dado_texto">Data inicial </td>
                    <td width="30%" height="25" class="form_dado_texto"><div id="div_data_inicio_b3">
                    <%
					if b3_vl_inic = "" or isnull(b3_vl_inic) or anula_bolsa3 = "S"	then
						b3_vl_inic = "0/0/0"
					end if	
					if b3_vl_fim = "" or isnull(b3_vl_fim) or anula_bolsa3 = "S" then
						b3_vl_fim = "0/0/0"
					end if						
						b3_data_de=split(b3_vl_inic,"/")
						b3_data_ate=split(b3_vl_fim,"/")
					%>
                    
                    
                    <select name="b3_dia_de" id="b3_dia_de" class="select_style">
                      <% 
							 For i =0 to 31
							 b3_data_de(0)=b3_data_de(0)*1
							 if b3_data_de(0)=i then 
							 	if b3_data_de(0)=0 then
									dd=""
								else
									if b3_data_de(0)<10 then
										dd="0"&b3_data_de(0)
									else
										dd=b3_data_de(0)									
									end if
								end if	
							 %>
                      <option value="<%response.Write(i)%>" selected>
                        <%response.Write(dd)%>
                        </option>
                      <% else
					  		if i=0 then
								i_cod=""
							else
							  	i_cod=i
								if i<10 then
								
								i="0"&i
								end if
							end if	
							%>
                      <option value="<%response.Write(i_cod)%>">
                        <%response.Write(i)%>
                        </option>
                      <% end if 
							next
							%>
                    </select>
                      /
                      <select name="b3_mes_de" id="b3_mes_de" class="select_style">
                        <%b3_data_de(1)=b3_data_de(1)*1
							if b3_data_de(1)=0 then%>
                        <option value="0" selected></option>
                        <% else%>
                        <option value="0"></option>
                        <%end if	
								if b3_data_de(1)=1 then%>
                        <option value="1" selected>janeiro</option>
                        <% else%>
                        <option value="1">janeiro</option>
                        <%end if
								if b3_data_de(1)=2 then%>
                        <option value="2" selected>fevereiro</option>
                        <% else%>
                        <option value="2">fevereiro</option>
                        <%end if
								if b3_data_de(1)=3 then%>
                        <option value="3" selected>mar&ccedil;o</option>
                        <% else%>
                        <option value="3">mar&ccedil;o</option>
                        <%end if
								if b3_data_de(1)=4 then%>
                        <option value="4" selected>abril</option>
                        <% else%>
                        <option value="4">abril</option>
                        <%end if
								if b3_data_de(1)=5 then%>
                        <option value="5" selected>maio</option>
                        <% else%>
                        <option value="5">maio</option>
                        <%end if
								if b3_data_de(1)=6 then%>
                        <option value="6" selected>junho</option>
                        <% else%>
                        <option value="6">junho</option>
                        <%end if
								if b3_data_de(1)=7 then%>
                        <option value="7" selected>julho</option>
                        <% else%>
                        <option value="7">julho</option>
                        <%end if%>
                        <%if b3_data_de(1)=8 then%>
                        <option value="8" selected>agosto</option>
                        <% else%>
                        <option value="8">agosto</option>
                        <%end if
								if b3_data_de(1)=9 then%>
                        <option value="9" selected>setembro</option>
                        <% else%>
                        <option value="9">setembro</option>
                        <%end if
								if b3_data_de(1)=10 then%>
                        <option value="10" selected>outubro</option>
                        <% else%>
                        <option value="10">outubro</option>
                        <%end if
								if b3_data_de(1)=11 then%>
                        <option value="11" selected>novembro</option>
                        <% else%>
                        <option value="11">novembro</option>
                        <%end if
								if b3_data_de(1)=12 then%>
                        <option value="12" selected>dezembro</option>
                        <% else%>
                        <option value="12">dezembro</option>
                        <%end if%>
                      </select>
                      /
                      <select name="b3_ano_de" id="b3_ano_de" class="select_style">
<%						if b3_data_ate(2) = 0 then
 %>
                        <option value="0" SELECTED>
                          </option>
                        <%	
						else
 %>
                        <option value="0">
                          </option>
                        <%												
						end if 
						 For ald =ano_letivo-1 to ano_letivo+1
						 	b3_data_de(2)=b3_data_de(2)*1
							if ald=b3_data_de(2) then
								selected="selected"
							else
								selected=""
							end if		
 %>
                        <option value="<%Response.Write(ald)%>" <%response.Write(selected)%>>
                          <%Response.Write(ald)%>
                          </option>
                        <%NEXT%>
                      </select></div></td>
                    <td width="2%" height="25" rowspan="2" valign="top" class="form_dado_texto"><% if b3_prazo="S" then%>
                      <input name="b3_prazo" type="radio" id="b3_prazoN" value="n" onClick="javascript:b3_desabilita_campo();limpa_prazo('b3','d');" checked>
                      <%else%>
                      <input name="b3_prazo" type="radio" id="b3_prazoN" value="n" onClick="javascript:b3_desabilita_campo();limpa_prazo('b3','d');">
                      <%end if%></td>
                    <td width="12%" height="25" align="right" class="form_dado_texto">Parcela Inicial  </td>
                    <td width="20%" height="25" class="form_dado_texto"><input name="b3_pi" type="text" class="textInput" id="b3_pi" size="4" maxlength="3" value="<%response.Write(b3_pc_inic)%>"></td>
                  </tr>
                  <tr>
                    <td width="12%" align="right" class="form_dado_texto">Data Final </td>
                    <td width="30%" class="form_dado_texto"><div id="div_data_fim_b3"><select name="b3_dia_ate" id="b3_dia_ate" class="select_style">
                      <% 
							 For i =0 to 31
							 b3_data_ate(0)=b3_data_ate(0)*1
							 if i=b3_data_ate(0) then 
							 	if b3_data_ate(0)=0 then
									da=""
								else
									if b3_data_ate(0)<10 then
										da="0"&b3_data_ate(0)
									else
										da=b3_data_ate(0)									
									end if
								end if		
							 %>
                      <option value="<%response.Write(i)%>" selected>
                        <%response.Write(da)%>
                        </option>
                      <% else
					  		if i=0 then
								i_cod=""
							else
							  	i_cod=i
								if i<10 then
								
								i="0"&i
								end if
							end if	
							%>
                      <option value="<%response.Write(i_cod)%>">
                        <%response.Write(i)%>
                        </option>
                      <% end if 
							next
							%>
                    </select>
                      /
                      <select name="b3_mes_ate" id="b3_mes_ate" class="select_style">
                        <%b3_data_ate(1)=b3_data_ate(1)*1
								if b3_data_ate(1)=0 then%>
                        <option value="0" selected></option>
                        <% else%>
                        <option value="0"></option>
                        <%end if	
								if b3_data_ate(1)=1 then%>
                        <option value="1" selected>janeiro</option>
                        <% else%>
                        <option value="1">janeiro</option>
                        <%end if
								if b3_data_ate(1)=2 then%>
                        <option value="2" selected>fevereiro</option>
                        <% else%>
                        <option value="2">fevereiro</option>
                        <%end if
								if b3_data_ate(1)=3 then%>
                        <option value="3" selected>mar&ccedil;o</option>
                        <% else%>
                        <option value="3">mar&ccedil;o</option>
                        <%end if
								if b3_data_ate(1)=4 then%>
                        <option value="4" selected>abril</option>
                        <% else%>
                        <option value="4">abril</option>
                        <%end if
								if b3_data_ate(1)=5 then%>
                        <option value="5" selected>maio</option>
                        <% else%>
                        <option value="5">maio</option>
                        <%end if
								if b3_data_ate(1)=6 then%>
                        <option value="6" selected>junho</option>
                        <% else%>
                        <option value="6">junho</option>
                        <%end if
								if b3_data_ate(1)=7 then%>
                        <option value="7" selected>julho</option>
                        <% else%>
                        <option value="7">julho</option>
                        <%end if%>
                        <%if b3_data_ate(1)=8 then%>
                        <option value="8" selected>agosto</option>
                        <% else%>
                        <option value="8">agosto</option>
                        <%end if
								if b3_data_ate(1)=9 then%>
                        <option value="9" selected>setembro</option>
                        <% else%>
                        <option value="9">setembro</option>
                        <%end if
								if b3_data_ate(1)=10 then%>
                        <option value="10" selected>outubro</option>
                        <% else%>
                        <option value="10">outubro</option>
                        <%end if
								if b3_data_ate(1)=11 then%>
                        <option value="11" selected>novembro</option>
                        <% else%>
                        <option value="11">novembro</option>
                        <%end if
								if b3_data_ate(1)=12 then%>
                        <option value="12" selected>dezembro</option>
                        <% else%>
                        <option value="12">dezembro</option>
                        <%end if%>
                      </select>

                      </select>
                      /
                      <select name="b3_ano_ate" id="b3_ano_ate" class="select_style">
                        <% 
						if b3_data_ate(2) = 0 then
 %>
                        <option value="0" SELECTED>
                          </option>
                        <%	
						else%>
                        <option value="0">
                          </option>                        
						<%				
						end if
						 For ala =ano_letivo-1 to ano_letivo+1
							b3_data_ate(2)=b3_data_ate(2)*1
							if ala=b3_data_ate(2) then
								selected="selected"
							else
								selected=""
							end if		
 %>
                        <option value="<%Response.Write(ala)%>" <%response.Write(selected)%>>
                          <%Response.Write(ala)%>
                          </option>
                        <%NEXT%>
                      </select></div></td>
                    <td width="12%" align="right" class="form_dado_texto">Parcela Final </td>
                    <td width="20%" class="form_dado_texto"><input name="b3_pf" type="text" class="textInput" id="b3_pf" size="4" maxlength="3" value="<%response.Write(b3_pc_fim)%>"></td>
                  </tr>
                </table></td>
                <td width="12%" align="right" class="form_dado_texto">Incid&ecirc;ncia </td>
                <td width="20%" align="left"><div id="aplica_b3"><select name="b3_aplica_bolsa" id="b3_aplica_bolsa" class="select_style">
                <% if anula="S" or anula_bolsa3 = "S" then
						z_select = "Selected"				
						s_select = ""
						p_select = ""									
						a_select = ""
					elseif b3_ap_bolsa = "P" then
						z_select = ""					
						s_select = ""
						p_select = "Selected"
						a_select = ""							
					elseif b3_ap_bolsa = "S" then
						z_select = ""					
						s_select = "Selected"
						p_select = ""
						a_select = ""										
					elseif b3_ap_bolsa = "A" then
						z_select = ""
						s_select = ""
						p_select = ""	
						a_select = "Selected"												
					end if															
					%>
                  <option value="nulo" <%response.Write(z_select)%>></option>  
                  <option value="P" <%response.Write(p_select)%>>Parcela da Anuidade</option>                                    
                  <option value="S" <%response.Write(s_select)%>>Serviços</option>
                  <option value="M" <%response.Write(a_select)%>>Ambos</option>
                </select></div></td>
                </tr>
              <tr>
                <td width="15%" height="25" align="right">&nbsp;</td>
                <td width="12%" height="25" align="right">&nbsp;</td>
                <td width="20%" height="25" align="right">&nbsp;</td>
                </tr>
              <tr class="form_dado_texto">
                <td width="15%" align="right" valign="top">Observa&ccedil;&atilde;o  </td>
                <td colspan="3" align="left"><textarea name="b3_observacao" cols="115" rows="3" class="textInput" id="b3_observacao" onKeyDown="limitText(this.form.b3_observacao,this.form.countdown,255);" onKeyUp="limitText(this.form.b3_observacao,this.form.countdown,255);"><%response.write(b3_ob_bolsa)%></textarea></td>
                <td width="12%" align="right">Data Concess&atilde;o</td>
                <td width="20%" align="left"><div id="div_dt_concessao_b3"><%response.write("&nbsp;"&b3_dt_conce)%></div><input name="b3_dt_conce" type="hidden" value="<%response.write(b3_dt_conce)%>"></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td><hr></td>
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