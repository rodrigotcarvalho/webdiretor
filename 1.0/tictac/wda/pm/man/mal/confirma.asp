<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<% 

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
opt = request.QueryString("opt")

chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

trava=session("trava")
pedido_form=request.querystring("cod")

cod_np = session("cod_np")
dia_de = session("dia_de")
mes_de = session("mes_de")
ano_de = session("ano_de")
dia_ate = session("dia_ate")
mes_ate = session("mes_ate")
ano_ate = session("ano_ate")
situacao = session("situacao")	

session("cod_np") = cod_np
session("dia_de") = dia_de
session("mes_de") = mes_de
session("ano_de") = ano_de
session("dia_ate") = dia_ate
session("mes_ate") = mes_ate
session("ano_ate") = ano_ate	
session("situacao") = situacao

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON99 = Server.CreateObject("ADODB.Connection") 
		ABRIR99 = "DBQ="& ACAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON99.Open ABRIR99			
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")		

		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		

		Set CON9 = Server.CreateObject("ADODB.Connection") 
		ABRIR9 = "DBQ="& CAMINHO_ax & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON9.Open ABRIR9	

		Set CON99 = Server.CreateObject("ADODB.Connection") 
		ABRIR99 = "DBQ="& ACAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON99.Open ABRIR99			


%>
<html>
<head>
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html;charset=ISO-8859-1">
<script language="JavaScript" src="../../../../js/mm_menu.js"></script>
<script type="text/javascript" src="../../../../js/global.js"></script>
      <script language="JavaScript" type="text/JavaScript">
<!--


function changeAction() {
    document.anf.action = "confirma.asp?opt=<%response.Write(opt)%>";
}
var checkflag = "false";
function check(field) {
if (checkflag == "false") {
for (i = 0; i < field.length; i++) {
field[i].checked = true;}
checkflag = "true";
return "Desmarcar Todos"; }
else {
for (i = 0; i < field.length; i++) {
field[i].checked = false; }
checkflag = "false";
return "Marcar Todos"; }
}
//-----------------------------------------------------

function ValidaNumero(val, campo) {

    if (isNaN(parseFloat(val))) {
		alert(unescape("O valor desse campo deve ser num%E9rico"));
		putFocusOn(campo);		
        return false;
		
     }

     return true

}
function valorFormat(fld) {
	var unidade = "R$";
	if(unidade.charAt(1) == 'U'){
	  milSep = ',';
	  decSep = '.';
	  }else{
	  milSep = '.';
	  decSep = ',';
		}  

	var key = '';
	var i = j = len = len2 = 0;
	var strCheck = '0123456789';//+decSep
	var aux = aux2 = '';
	len = fld.value.length;

	for(; i < len; i++)	
	if ((fld.value.charAt(i) != '0') && (fld.value.charAt(i) != decSep)) break;
	aux = '';
	for(;i < len; i++)
  if (strCheck.indexOf(fld.value.charAt(i))!=-1) aux += fld.value.charAt(i);
  aux += key;
  document.anf.aux_format.value = aux;
  len = aux.length;
  if (len > 2){
 	 aux2 = '';
 	 for (j = 0, i = len - 3; i >= 0; i--) {
    if (j == 3) {
   	 aux2 += milSep;
   	 j = 0;
    }
    aux2 += aux.charAt(i);
    j++;
 	 }
 	 fld.value = '';
 	 len2 = aux2.length; 	 
 	 for (i = len2 - 1; i >= 0; i--){
    fld.value += aux2.charAt(i);
 	 }
 	 fld.value += decSep + aux.substr(len - 2, len);
  }
}
function FormataNumero(valor) {
	var unidade = "R$";
	if(unidade.charAt(1) == 'U'){
	  milSep = ',';
	  decSep = '.';
	  }else{
	  milSep = '.';
	  decSep = ',';
		}  

	var key = '';
	var i = j = len = len2 = 0;
	var strCheck = '0123456789';//+decSep
	var aux = aux2 = '';
	len = valor.length;

	for(; i < len; i++)	
	if ((valor.value.charAt(i) != '0') && (valor.value.charAt(i) != decSep)) break;
	aux = '';
	for(;i < len; i++)
  if (strCheck.indexOf(valor.value.charAt(i))!=-1) aux += valor.value.charAt(i);
  aux += key;
  document.anf.aux_format.value = aux;
  len = aux.length;
  if (len > 2){
 	 aux2 = '';
 	 for (j = 0, i = len - 3; i >= 0; i--) {
    if (j == 3) {
   	 aux2 += milSep;
   	 j = 0;
    }
    aux2 += aux.charAt(i);
    j++;
 	 }
 	 fld.value = '';
 	 len2 = aux2.length; 	 
 	 for (i = len2 - 1; i >= 0; i--){
    fld.value += aux2.charAt(i);
 	 }
 	 fld.value += decSep + aux.substr(len - 2, len);
  }
}

function currencyFormat(fld) {
  valorFormat(fld);
  produto(fld);
}

function produto(fld)  
{ 		//var t,linha	;
//		t = fld.id.split("_");
//		linha = t[1];		
//		var fator1 = document.getElementById('quantidade_'+linha)		
//		var fator2 = document.getElementById('valor_'+linha)		
//		var resultado = document.getElementById('produto_'+linha)	
//		if (isNaN(parseFloat(fator1.value)) || isNaN(parseFloat(fator2.value))){
//			}else{
//			//invertendo do formato brasileiro para o americano
//			convertido1 = fator1.value.replace( '.', '' );  
//			convertido2 = fator2.value.replace( '.', '' );  				
//			convertido1 = convertido1.replace( ',', '.' );  
//			convertido2 = convertido2.replace( ',', '.' );  														
//			resultado.value = convertido1*convertido2;  
	
			soma();
//		}
}

function soma()  
{ 
//todas as linhas que foram criadas na tabela nessa sess�o incluindo as que foram exclu�das  		
		var total_itens = document.getElementById('itens_criados');
		
		var money = 0	
		var total = document.getElementById('total')
		
	for (var i=1;i<=total_itens.value;i++)
	{ 
			var produto = document.getElementById('produto_'+i)
			//produto = produto.value.replace( ',', '.' );  
			if (produto){
			//money = (produto.value*100)+money; 
			money = produto.value+money; 			
			}
			
	}		//var arredonda = money/100;  
			//total.value = arredonda.toFixed(2)
			
			//var valor_arredondado = money;
			//var valor_arredondado = Math.round(money)
			//var valor_arredondado = Math.ceil(money)
			//var valor_arredondado = Math.flor(money)						
			//total.value = valor_arredondado/100;  
		
}
function putFocusOn(campo)
{
	  var focal = document.getElementById(campo);
	focal.focus(); 
}
function deleteRow(ln)
{	
	//As linhas s�o contadas a partir do zero 
	//enquanto na tabela est�o numeradas a partir do 1
	var linha_da_tabela = ln-1
	//var linha_imagem = linha_da_tabela-1	
	//var identidade ="close_"+linha_imagem
	//alert(identidade)	
  var qtd_itens = document.getElementById('qtd_itens');	
  //linhas existentes na tabela atualmente
  var total_itens = document.getElementById('qtd_itens').value;	 	
  //todas as linhas que foram criadas na tabela nessa sess�o incluindo as que foram exclu�das  
  var itens_criados = document.getElementById('itens_criados').value;	 	
  var linha_a_apagar = linha_da_tabela-(itens_criados-total_itens)
	document.getElementById("tblInnerHTML").deleteRow(linha_a_apagar);
	//document.getElementById(identidade).onclick = function () { deleteRow(linha_a_apagar); ShowImage(linha_imagem);};
 
  var rowCount = total_itens-1;  

  qtd_itens.value = rowCount;  	
  
//changeAction();
soma(); 
}
function addRow()
{
	addRowInnerHTML('tblInnerHTML')
}
function addRowInnerHTML(tblId)
{
  var table = document.getElementById(tblId);	
  var tblBody = document.getElementById(tblId).tBodies[0];
  var newRow = tblBody.insertRow(-1);
  var qtd_itens = document.getElementById('qtd_itens');	
  var itens_criados = document.getElementById('itens_criados');	
  var total_itens = document.getElementById('itens_criados').value;	  
  var prox_item = (total_itens*1)+1;    
	
  itens_criados.value = prox_item;  
	
  var rowCount = table.rows.length;  

  qtd_itens.value = rowCount;  
  
  var newCell0 = newRow.insertCell(0); 
  newRow.className = 'form_dado_texto';  
  newCell0.align = 'right'  
  newCell0.width = 140
  newCell0.innerHTML = prox_item;
  
  var newCell1 = newRow.insertCell(1);
  newCell1.align = 'right'    
  newCell1.width = 140  
  newCell1.innerHTML = 'Item:';
  
  var newCell2 = newRow.insertCell(2);
  newCell2.align = 'left'    
  newCell2.width = 180    
  newCell2.innerHTML = '<select name="item_fornecedor_'+prox_item+'" class="select_style" id="item_fornecedor_'+prox_item+'"><option value="nulo" selected></option><%
		Set RS = Server.CreateObject("ADODB.Recordset")
		sql = "Select * From TB_Item order by NO_Item"
		RS.Open sql, CON9  
		
		while not RS.EOF 
		cod_item=RS("CO_Item")		
		nome_item=RS("NO_Item")		  
		
		%><option value="<%response.Write(cod_item)%>"><%response.Write(nome_item)%></option><%
		RS.MOVENEXT
		WEND 		
				%></select>';  
				
  var newCell3 = newRow.insertCell(3);
  newCell3.align = 'right' 
  newCell3.width = 180        
  newCell3.innerHTML = 'Quantidade:';		
  
  var newCell4 = newRow.insertCell(4);
  newCell4.align = 'left'    
  newCell4.width = 180   
  newCell4.innerHTML = '<input name="quantidade_'+prox_item+'" type="text" value="1" class="textInput" id="quantidade_'+prox_item+'" size="15" maxlength="15" onBlur="return (ValidaNumero(this.value, this.id));" onChange="return (produto(this));" onFocus="javascript:this.form.quantidade_'+prox_item+'.select();">';  	
  
//  var newCell5 = newRow.insertCell(5);
//  newCell5.align = 'right' 
//  newCell5.width = 184      
//  newCell5.innerHTML = 'Valor Unit&aacute;rio:';	  	
//  
//  var newCell6 = newRow.insertCell(6);
//  newCell6.align = 'left'   
//  newCell6.width = 184    
//  newCell6.innerHTML = '<input name="valor_'+prox_item+'" type="text" value="0" class="textInput" id="valor_'+prox_item+'" size="20" maxlength="20" onBlur="return (ValidaNumero(this.value, this.id));" onChange="return (produto(this));" onKeyDown="return (currencyFormat(this))" onKeyUp="return (currencyFormat(this))" onFocus="javascript:this.form.valor_'+prox_item+'.select();"><input name="aux_format" readonly type="hidden"><input name="produto_'+prox_item+'" type="hidden" id="produto_'+prox_item+'" value="0">';	 
  
  var newCell5 = newRow.insertCell(5);
  newCell5.align = 'left'  
  newCell5.width = 180     
  newCell5.innerHTML = '<div id="'+prox_item+'"><a href="#" onClick="addRow();changeImage('+prox_item+');"><img src="../../../../img/add.png" alt="Adicionar Item" width="20" height="20" border = "0"></a></div>';	  
  
 putFocusOn('item_fornecedor_'+prox_item);  
}
function changeImage(img){

	document.getElementById(img).innerHTML ='<a id="close_'+img+'" href="#" onClick="deleteRow('+img+');"><img src="../../../../img/close.png" alt="Excluir Item" width="20" height="20" border = "0" ></a>';

}  
function ShowImage(img){
	document.getElementById(img).style.visibility='visible';
}  
function hideImage(img){
	document.getElementById(img).style.visibility='hidden';
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
//-->
</script>
      <script language="JavaScript" type="text/JavaScript">
<!--
function checksubmit()
{
// if (document.anf.nota_fiscal.value == "")
//  {    alert("Por favor digite uma nota fiscal!")
//   document.anf.nota_fiscal.focus()
//    return false 
// }
 
 var sel_etapa = document.getElementById('etapa')
 if (sel_etapa.value == "nulo" || sel_etapa.value == "")
  {    alert("Por favor selecione uma etapa!")
   sel_etapa.focus()
	return false
 } 
 var sel_turma = document.getElementById('turma')
 if (sel_turma.value == 999990 || sel_turma.value == "999990" || sel_turma.value == "")
  {    alert("Por favor selecione uma turma!")
   sel_turma.focus()
	return false
 }  

var sel_projeto = document.getElementById('projeto')
 if (sel_projeto.value == "nulo" || sel_projeto.value == "")
  {    alert("Por favor selecione um projeto!")
   sel_projeto.focus()
	return false
 } 

 
//    if (document.anf.valor.value == 0 || document.anf.valor.value == "")
//  {    alert("Por favor digite um valor diferente de zero!")
//    document.anf.valor.focus()
//    return false
//  } 
 var total_itens = document.getElementById('itens_criados');
 var verifica_loop;
 if (total_itens.value == 1){
	 verifica_loop = 1
 } else{
	 verifica_loop = total_itens.value - 1	 
 }
 for (var i=1;i<=verifica_loop;i++)
 { 
		var item_fornecedor_i = document.getElementById('item_fornecedor_'+i)
		var quantidade_i = document.getElementById('quantidade_'+i)
		//var valor_i = document.getElementById('valor_'+i)				

		if (item_fornecedor_i){
		 if (item_fornecedor_i.value == "nulo")
		  {    alert("Por favor selecione um item!")
		   item_fornecedor_i.focus()
			return false
		 } 
		}
		
		if (quantidade_i){
		 if (quantidade_i.value == 0 || quantidade_i.value == "")
		  {    alert("Por favor digite uma quantidade diferente de zero!")
		   quantidade_i.focus()
			return false
		 } 
		}
		
//		if (valor_i){
//		 if (valor_i.value == 0 || valor_i.value == "")
//		  {    alert("Por favor digite um valor diferente de zero!")
//		   valor_i.focus()
//			return false
//		 } 
//		}				
		
	} 
 //aula = document.busca.aula.value;
//     if (aula.length > 3)
//  {    alert("O valor do campo Aula deve possuir menos que 3 caracteres")
//    document.busca.aula.focus()
//    return false
//  }

  return true
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
                                                            alert("Esse browser n�o tem recursos para uso do Ajax");
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
document.all.divEtapa.innerHTML ="<select class=select_style></select>"
document.all.divTurma.innerHTML = "<select class=select_style></select>"
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
document.all.divTurma.innerHTML = "<select class=select_style></select>"
//recuperarTurma()
                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
function recuperarProjeto(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=proj", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_p= oHTTPRequest.responseText;
resultado_p = resultado_p.replace(/\+/g," ")
resultado_p = unescape(resultado_p)
document.all.divProj.innerHTML = resultado_p
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }								   
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
      </script>
</head>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
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
<% if opt="exc"	then
exclui_pedido = pedido_form
%>		  
        <form action="bd.asp?opt=exc" method="post" name="busca" id="busca">
          <tr>
      <td valign="top"> 
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>            <tr> 
              
    <td width="766" height="10" colspan="4" valign="top"> 
      <%call mensagens(nivel,834,0,0) %>
    </td>
			  </tr>
          <tr> 
            <td height="10" class="tb_tit"
>Movimenta&ccedil;&otilde;es de Estoque a serem exclu&iacute;das</td>
          </tr>
          <tr> 
            <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
                <tr class="tb_subtit"> 
                  <td width="20" height="10"><div align="center">
                    <input name="exclui_pedido" type="hidden" id="exclui_pedido" value="<%response.write(exclui_pedido)%>">
                  </div></td>
                  <td align="center"> N&ordm; Pedido </td>
                  <td align="center"> Data do Pedido </td>
                  <td><div align="left"> Projeto <font class="form_dado_texto">
                    <input name="cod" type="hidden" id="cod" value="<%=codigo%>">
                    <input name="data_de" type="hidden" class="textInput" id="data_de"  value="<%response.Write(data_de)%>" size="75" maxlength="50">
                    <input name="data_inicio" type="hidden" class="textInput" id="data_inicio"  value="<%response.Write(data_inicio)%>" size="75" maxlength="50">
                    <input name="data_ate" type="hidden" class="textInput" id="data_ate"  value="<%response.Write(data_ate)%>" size="75" maxlength="50">
                    </font><font class="form_dado_texto">
                      <input name="data_fim" type="hidden" class="textInput" id="data_fim"  value="<%response.Write(data_fim)%>" size="75" maxlength="50">
                    </font></div></td>
                  <td align="center"> Unidade </td>
                  <td align="center"> Curso </td>
                  <td align="center"> Etapa </td>
                  <td align="center"> Turma </td>
                  <td align="center"> Solicitado por </td>
                  <td align="center"> Situa&ccedil;&atilde;o </td>
                  <td align="center"> Atendido em </td>
                </tr>
                <%
'response.Write(">>"&exclui_ocorrencia)				
check = 2	
exclui_pedido = replace(pedido_form,"$!$","/")		
vetorExclui = split(exclui_pedido,", ")
conta_ocorr=0
exibe_confimar = "S"
for i =0 to ubound(vetorExclui)

exclui = split(vetorExclui(i),"?")

'obr=cod&"?"&da_ocorrencia&"?"&ho_ocorrencia&"?"&co_ocorrencia
nu_pedido = exclui(0)
data_pedido= exclui(1)

				
dados_data=split(data_pedido,"/")
dia=dados_data(0)
mes=dados_data(1)
ano=dados_data(2)




data_pedido_cons=mes&"/"&dia&"/"&ano
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Mov_Estoque, TB_Projeto WHERE TB_Mov_Estoque.CO_Projeto = TB_Projeto.CO_Projeto AND NU_Pedido ="& nu_pedido&" AND (TB_Mov_Estoque.DA_Pedido BETWEEN #"&data_pedido_cons&"# AND #"&data_pedido_cons&"#)"		
		RS.Open SQL, CON9


  if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if 
  
seq_pd=RS("NU_Pedido")
da_pd=RS("DA_Pedido")
no_projeto=RS("NO_Projeto")
unidade=RS("NU_Unidade")
curso=RS("CO_Curso")
etapa=RS("CO_Etapa")
turma=RS("CO_Turma")
observacao=RS("TX_Observa")
situacao=RS("ST_Pedido")
da_atendido=RS("DA_Atendido")
co_usu_reg=RS("CO_Usuario")



no_unidade = GeraNomesNovaVersao("U",unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)
no_curso = GeraNomesNovaVersao("CA",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro)
no_etapa = GeraNomesNovaVersao("E",curso,etapa,variavel3,variavel4,variavel5,CON0,outro)


Select case situacao

case "A"
st_movim="Atendido"

case "P"
st_movim="Pendente"

case "C"
st_movim="Cancelado"
exibe_confimar = "N"
end select	

		
if co_usu_reg="" or isnull(co_usu_reg) then
	no_registrador=""
else

		Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& co_usu_reg
		RSu.Open SQLu, CON

	IF RSu.EOF then
		no_registrador=""
	else
		no_registrador=RSu("NO_Usuario")
	end if		
end if			


optobr=seq_pd&"?"&da_pd

Session("data_de")=data_de
Session("data_inicio")=data_inicio
Session("data_ate")=data_ate
Session("data_fim")=data_fim


data_split= Split(da_pd,"/")
dia=data_split(0)
mes=data_split(1)
ano=data_split(2)


dia=dia*1

mes=mes*1
hora=hora*1
min=min*1

if dia<10 then
dia="0"&dia
end if
if mes<10 then
mes="0"&mes
end if
if hora<10 then
hora="0"&hora
end if
if min<10 then
min="0"&min
end if
da_show=dia&"/"&mes&"/"&ano


if da_atendido="" or isnull(da_atendido) then
 da_atendido_show=""
else
	data_split= Split(da_atendido,"/")
	dia=data_split(0)
	mes=data_split(1)
	ano=data_split(2)
	
	
	dia=dia*1
	
	mes=mes*1
	hora=hora*1
	min=min*1
	
	if dia<10 then
	dia="0"&dia
	end if
	if mes<10 then
	mes="0"&mes
	end if
	if hora<10 then
	hora="0"&hora
	end if
	if min<10 then
	min="0"&min
	end if
	da_atendido_show=dia&"/"&mes&"/"&ano
end if


%>
                <tr class="<%=cor%>"> 
                  <td width="20">&nbsp;</td>
                  <td align="center">&nbsp;
                    <%response.Write(seq_pd)%>
                    <div align="center"></div>
                    <div align="left"></div></td>
                  <td align="center"><%response.Write(da_show)%></td>
                  <td><%response.Write(no_projeto)%></td>
                  <td align="center"><%response.Write(no_unidade)%></td>
                  <td align="center"><%response.Write(no_curso)%></td>
                  <td align="center"><%response.Write(no_etapa)%></td>
                  <td align="center"><%response.Write(turma)%></td>
                  <td align="center"><%response.Write(no_registrador)%></td>
                  <td><div align="center">
                    <%response.Write(st_movim)%>
                  </div></td>
                  <td><div align="center">
                    <%response.Write(da_atendido_show)%>
                  </div></td>
                </tr>
<%check = check+1
next
%>				
              </table></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></div></td>
          </tr>
          <tr bgcolor="#FFFFFF">
            <td><hr></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td><div align="center"> 
                <table width="1000" border="0" align="center" cellspacing="0">
                  <tr> 
                    <td width="391"> <div align="center"> 
                        <input name="alterar" type="submit" class="botao_cancelar" id="alterar" onClick="MM_goToURL('parent','resumo.asp?opt=vt');return document.MM_returnValue" value="Voltar">
                      </div></td>
                    <td width="391">&nbsp;</td>
                    <td width="218"> <div align="left"> 
                     <% IF exibe_confimar = "S" THEN%>
                        <input name="Submit" type="submit" class="botao_prosseguir" value="Confirmar">
                     <%END IF%>   
                      </div></td>
                  </tr>
                </table>
                <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> </font></div></td>
          </tr>
        </table></td>
    </tr>
</form>
<%else
' OPT = ALT
' OPT = ERRI
' OPT = ERRA

nota_fiscal_alt = replace(pedido_form,"$!$","/")		
vetorAltera = split(nota_fiscal_alt,"?")

cod_nf = vetorAltera(0)
data_nf= vetorAltera(1)
dados_msg = cod_nf&"?"&data_nf
				
dados_data=split(data_nf,"/")
dia=dados_data(0)
mes=dados_data(1)
ano=dados_data(2)




data_nf_cons=mes&"/"&dia&"/"&ano

	if opt="erri" or opt="erra" then
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TBA_Mov_Estoque WHERE NU_Pedido = "&cod_nf
		RS.Open SQL, CON99
  

	else	
		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Mov_Estoque WHERE TB_Mov_Estoque.NU_Pedido = "&cod_nf
		RS.Open SQL, CON9

	
  	end if
  
	seq_pd=RS("NU_Pedido")
	da_pd=RS("DA_Pedido")
	co_projeto=RS("CO_Projeto")	
	unidade=RS("NU_Unidade")
	curso=RS("CO_Curso")
	co_etapa=RS("CO_Etapa")
	turma=RS("CO_Turma")
	observacao=RS("TX_Observa")
	situacao=RS("ST_Pedido")
	da_atendido=RS("DA_Atendido")
	co_usu_reg=RS("CO_Usuario")	
  

  if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if 


data_split= Split(da_pd,"/")
dia_pd=data_split(0)
mes_pd=data_split(1)
ano_pd=data_split(2)


dia_pd=dia_pd*1
mes_pd=mes_pd*1
ano_pd=ano_pd*1


if dia_pd<10 then
dia_pd="0"&dia_pd
end if
if mes_pd<10 then
mes_pd="0"&mes_pd
end if

da_show=dia_pd&"/"&mes_pd&"/"&ano_pd



if opt="erra" or opt="alt" then
	action = "bd.asp?opt=alt"
elseif opt="erri" then
	action = "bd.asp?opt=inc"
end if

%>

          <tr>   
            <td valign="top">
        <form action="<%response.Write(action)%>" method="post" name="anf" id="anf" onSubmit="return checksubmit()">            
            <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
<%if opt="alt" then%>
            <tr> 
              
    <td height="10" colspan="8" valign="top"> 
      <%call mensagens(nivel,9708,0,0) %>
    </td>
			  </tr>   
            <tr> 
              
    <td height="10" colspan="8" valign="top"> 
      <%call mensagens(nivel,645,0,"R24") %>
    </td>
			  </tr> 
<%elseif opt="erra" or  opt="erri" then%>
            <tr> 
              
    <td height="10" colspan="8" valign="top"> 
      <%call mensagens(nivel,9999,1,chave) %>
    </td>
			  </tr> 
<%	

END IF%>                                 
                <tr>
                <td width="101" class="tb_tit"
>Pedido</td>
                <td colspan="7" class="tb_tit"
></td>
              </tr>
                <tr>
                <td width="101" height="10" align="right" class="form_dado_texto">N&ordm; Pedido:</td>
                <td width="127"  class="form_dado_texto"><%response.Write(seq_pd)%>
                  <input name="nu_pedido" type="hidden" id="nu_pedido" size="35" maxlength="30" value="<%response.Write(seq_pd)%>"><input name="situacao_pedido" type="hidden" id="situacao_pedido" size="35" maxlength="30" value="<%response.Write(situacao)%>"></td>
                <td width="154" align="right" class="form_dado_texto">Data do Pedido:</td>
                <td width="129" class="form_dado_texto">
                  <input name="dia_nf" type="hidden" id="dia_nf" size="35" maxlength="30" value="<%response.Write(dia_pd)%>">	
                  <input name="mes_nf" type="hidden" id="mes_nf" size="35" maxlength="30" value="<%response.Write(mes_pd)%>">                  <input name="ano_nf" type="hidden" id="ano_nf" size="35" maxlength="30" value="<%response.Write(ano_pd)%>">                  			
				<%response.Write(da_show)%><!--<select name="dia_nf" id="dia_nf" class="select_style">
                    <% 
							 For i =1 to 31
							 dia=dia*1
							 if dia=i then 
								if dia<10 then
								dia="0"&dia
								end if
							 %>
                    <option value="<%response.Write(i)%>" selected>
                  <%response.Write(dia)%>
                  </option>
                    <% else
								if i<10 then
								i="0"&i
								end if
							%>
                    <option value="<%response.Write(i)%>">
                  <%response.Write(i)%>
                  </option>
                    <% end if 
							next
							%>
                  </select>
                    /
                    <select name="mes_nf" id="mes_nf" class="select_style">
                    <%mes=mes*1
								if mes="1" or mes=1 then%>
                    <option value="1" selected>janeiro</option>
                    <% else%>
                    <option value="1">janeiro</option>
                    <%end if
								if mes="2" or mes=2 then%>
                    <option value="2" selected>fevereiro</option>
                    <% else%>
                    <option value="2">fevereiro</option>
                    <%end if
								if mes="3" or mes=3 then%>
                    <option value="3" selected>mar&ccedil;o</option>
                    <% else%>
                    <option value="3">mar&ccedil;o</option>
                    <%end if
								if mes="4" or mes=4 then%>
                    <option value="4" selected>abril</option>
                    <% else%>
                    <option value="4">abril</option>
                    <%end if
								if mes="5" or mes=5 then%>
                    <option value="5" selected>maio</option>
                    <% else%>
                    <option value="5">maio</option>
                    <%end if
								if mes="6" or mes=6 then%>
                    <option value="6" selected>junho</option>
                    <% else%>
                    <option value="6">junho</option>
                    <%end if
								if mes="7" or mes=7 then%>
                    <option value="7" selected>julho</option>
                    <% else%>
                    <option value="7">julho</option>
                    <%end if%>
                    <%if mes="8" or mes=8 then%>
                    <option value="8" selected>agosto</option>
                    <% else%>
                    <option value="8">agosto</option>
                    <%end if
								if mes="9" or mes=9 then%>
                    <option value="9" selected>setembro</option>
                    <% else%>
                    <option value="9">setembro</option>
                    <%end if
								if mes="10" or mes=10 then%>
                    <option value="10" selected>outubro</option>
                    <% else%>
                    <option value="10">outubro</option>
                    <%end if
								if mes="11" or mes=11 then%>
                    <option value="11" selected>novembro</option>
                    <% else%>
                    <option value="11">novembro</option>
                    <%end if
								if mes="12" or mes=12 then%>
                    <option value="12" selected>dezembro</option>
                    <% else%>
                    <option value="12">dezembro</option>
                    <%end if%>
                  </select>
                    /
                    <select name="ano_nf" class="select_style" id="ano_nf">
                    <%
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Ano_Letivo order by NU_Ano_Letivo"
		RS0.Open SQL0, CON
		while not RS0.EOF 
		ano_bd=RS0("NU_Ano_Letivo")
		
				ano_letivo=ano_letivo*1
		ano_bd=ano_bd*1

				if ano_letivo=ano_bd then%>
                    <option value="<%=ano_bd%>" selected><%=ano_bd%></option>
                    <%else%>
                    <option value="<%=ano_bd%>"><%=ano_bd%></option>
                    <%end if
		RS0.MOVENEXT
		WEND 		
				%>
                  </select>--></td>
                <td width="110" height="10" align="right" class="form_dado_texto">					<input type="hidden" name="unidade" id="unidade" value ="<%response.Write(unidade)%>">
                    <input type="hidden" name="curso" id="curso" value ="<%response.Write(curso)%>">Etapa:</td>
                <td width="192" class="form_dado_texto"><div id="divEtapa">
                        <select name="etapa" id="etapa" class="select_style" onChange="recuperarTurma(this.value);recuperarProjeto(this.value)">
                          <%		

		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"'"
		RS0b.Open SQL0b, CON0
		
		
While not RS0b.EOF
Etapa = RS0b("CO_Etapa")


		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&curso&"' AND CO_Etapa='"&Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
if Etapa=co_etapa then
%>
                          <option value="<%response.Write(Etapa)%>" selected>
                            <%response.Write(NO_Etapa)%>
                          </option>
                          <%
else
%>
                          <option value="<%response.Write(Etapa)%>">
                            <%response.Write(NO_Etapa)%>
                          </option>
                          <%

end if
RS0b.MOVENEXT
WEND
%>
                        </select>
                      </div></td>
                <td width="117" align="right" class="form_dado_texto">Turma:</td>
                <td width="56" class="form_dado_texto"><div id="divTurma">
                        <select name="turma" id="turma" class="select_style" onChange="MM_callJS('submitfuncao()')">
                          <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Turma where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"' AND CO_Etapa='" & co_etapa & "' order by CO_Turma" 
		RS3.Open SQL3, CON0						

while not RS3.EOF
co_turma= RS3("CO_Turma")

if co_turma=turma then
%>
                          <option value="<%response.Write(co_turma)%>" selected>
                            <%response.Write(co_turma)%>
                          </option>
                          <%
else
%>
                          <option value="<%=co_turma%>">
                            <%response.Write(co_turma)%>
                          </option>
                          <%
co_turma_check = co_turma
end if
RS3.MOVENEXT
WEND
%>
                        </select>
                      </div></td>
              </tr>
                <tr>
                  <td colspan="8"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="10" align="right" class="form_dado_texto">Projeto:</td>
                      <td colspan="3"><div id="divProj"><select name="projeto" class="select_style" id="projeto">
                    <%
		Set RS = Server.CreateObject("ADODB.Recordset")
		sql = "Select * From TB_Projeto order by NO_Projeto"
		RS.Open sql, CON9  
		
		while not RS.EOF 
		cod_for=RS("CO_Projeto")		
		nome_for=RS("NO_Projeto")
		
		if isnumeric(cod_for) then		
			cod_for = cod_for*1
		end if	
		if isnumeric(co_projeto) then		
			co_projeto	= co_projeto*1
		end if					  
		 
		 if co_projeto = cod_for then
		 	selected_projeto = "selected"
		 else
		 	selected_projeto = ""		 
		 end if
		%>
                    <option value="<%response.Write(cod_for)%>" <%response.Write(selected_projeto)%>>
                  <%response.Write(nome_for)%>
                  </option>
                    <%
		RS.MOVENEXT
		WEND 		
				%>
                  </select></div></td>
                    </tr>
                    <tr>
                      <td width="10%" height="10" align="right" class="form_dado_texto">&nbsp;</td>
                      <td colspan="3">&nbsp;</td>
                    </tr>
                    <tr>
                      <td height="10" align="right" valign="top" class="form_dado_texto">Observa&ccedil;&atilde;o:</td>
                      <td width="53%"><textarea name="obs" cols="80" rows="2" class="textInput" id="obs"><%response.write(observacao)%>
                      </textarea></td>
                      <td width="10%" height="10" align="right" valign="top" class="form_dado_texto">Solicitado por: </td>
                      <td width="27%" valign="top"><span class="form_dado_texto">
                        <select name="solicitado" class="select_style" id="solicitado">
                          <%
							Set RSU = Server.CreateObject("ADODB.Recordset")
							SQLU = "SELECT * FROM TB_Usuario Where CO_Usuario<>0 AND ST_Usuario='L' ORDER BY NO_Usuario"		
							RSU.Open SQLU, CON
							
							while not RSU.EOF
								cod_usuario = RSU("CO_Usuario")							
								nome_usuario = RSU("NO_Usuario")
								if cod_usuario ="" or isnull(cod_usuario) then
									selected_usuario = ""
								else
									if co_usu_reg = cod_usuario then
										selected_usuario = "selected"									
									else
										selected_usuario = ""									
									end if
								end if	
							%>
                          <option value="<%response.Write(cod_usuario)%>" <%response.Write(selected_usuario)%>>
                            <%response.Write(nome_usuario)%>
                          </option>
                          <%
							RSU.MOVENEXT
							WEND				  
						  %>
                        </select>
                      </span></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td colspan="8">&nbsp;</td>
                </tr>
                <tr>
                <td colspan="8" class="tb_tit"
>Composi&ccedil;&atilde;o do Pedido</td>
              </tr>
                <tr>
                <td colspan="8"><table id="tblInnerHTML" width="100%" border="0" cellspacing="0" cellpadding="0">
<% 		

if opt="erra" or  opt="erri" then
		Set RSC = Server.CreateObject("ADODB.Recordset")
		SQLC = "Select COUNT(CO_Item) as Total From TBA_Mov_Estoque_Item where NU_Pedido = "&seq_pd
		RSC.Open SQLC, CON99 
		
		if RSC.EOF then
			total = 1	
		else
			total = RSC("Total")		
		end if


		Set RSI = Server.CreateObject("ADODB.Recordset")
		SQLI = "Select CO_Item, QT_Solicitado From TBA_Mov_Estoque_Item where NU_Pedido ="& seq_pd&" GROUP BY CO_Item, QT_Solicitado"
		RSI.Open SQLI, CON99	  

else

		Set RSC = Server.CreateObject("ADODB.Recordset")
		SQLC = "Select COUNT(CO_Item) as Total From TB_Mov_Estoque_Item where NU_Pedido = "&seq_pd
		RSC.Open SQLC, CON9	 
		
		if RSC.EOF then
			total = 1	
		else
			total = RSC("Total")		
		end if


		Set RSI = Server.CreateObject("ADODB.Recordset")
		SQLI = "Select CO_Item, QT_Solicitado From TB_Mov_Estoque_Item where NU_Pedido ="& seq_pd&" GROUP BY CO_Item, QT_Solicitado"
		RSI.Open SQLI, CON9	  

end if		
		linhas = 0    
		soma=0
		While not RSI.EOF         
			linhas = linhas+1

			co_item = RSI("CO_Item")			
			quantidade_item = RSI("QT_Solicitado")

				
			if isnull(quantidade_item) or quantidade_item ="" then
				quantidade_item = 0
			end if

			produto = quantidade_item
			soma = soma+produto  					
%>		              
                    <tr>
                      <td width="140" align="right" class="form_dado_texto"><input name="num_linha" type="hidden" id="num_linha" value="<%response.Write(linhas)%>">
                      <%response.Write(linhas)%></td>
                    <td width="140" align="right" class="form_dado_texto">Item:</td>
                    <td width="180"><select name="item_fornecedor_<%response.Write(linhas)%>" class="select_style" id="item_fornecedor_<%response.Write(linhas)%>">
                        <option value="nulo" selected></option>
                        <%
		Set RS = Server.CreateObject("ADODB.Recordset")
		sql = "Select * From TB_Item order by NO_Item"
		RS.Open sql, CON9  
		
		while not RS.EOF 
		cod_item_bd=RS("CO_Item")		
		nome_item=RS("NO_Item")
		if isnumeric(cod_item_bd) then		
			cod_item_bd = cod_item_bd*1
		end if	
		if isnumeric(co_item) then		
			co_item	= co_item*1
		end if			

		if cod_item_bd = co_item then
		 	selected_item = "selected"		
		else
		 	selected_item = ""			
		end if  
		
		%>
                        <option value="<%response.Write(cod_item_bd)%>" <%response.Write(selected_item)%>>
                      <%response.Write(nome_item)%>
                      </option>
                        <%
		RS.MOVENEXT
		WEND 		
				%>
                      </select></td>
                    <td width="180" align="right" class="form_dado_texto">Quantidade:</td>
                    <td width="180"><input name="quantidade_<%response.Write(linhas)%>" type="text" class="textInput" id="quantidade_<%response.Write(linhas)%>" onBlur="return (ValidaNumero(this.value, this.id));" onChange="return (produto(this))" value="<%response.Write(quantidade_item)%>" size="15" maxlength="15" onFocus="javascript:this.form.quantidade_<%response.Write(linhas)%>.select();" ></td>                    
                    <td width="180">                                
                    <div id="<%response.Write(linhas)%>">
                    <% 
					total = total*1
					linhas = linhas*1
					if total = linhas then%>
                        <a href="#"  onClick="addRow();changeImage(<%response.Write(linhas)%>)"><img src="../../../../img/add.png" alt="Adicionar Item" width="20" height="20" border = "0" ></a>
                    <%else%>
                        <a id="close_<%response.Write(linhas)%>" href="#" onClick="deleteRow(<%response.Write(linhas)%>);"><img src="../../../../img/close.png" alt="Excluir Item" width="20" height="20" border = "0" ></a>                        
                    <% end if%>
                    </div></td>
                  </tr>
<% RSI.MOVENEXT
WEND

%>                  
                  </table></td>
                  
              </tr>
<!--                <tr bgcolor="#FFFFFF">
                <td colspan="7"><table id="tblInnerHTML" width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td colspan="8" align="right" class="form_dado_texto"><hr></td>
                    </tr>
                    <tr>
                      <td width="100" align="right" class="form_dado_texto">&nbsp;</td>
                    <td width="100" align="right" class="form_dado_texto">&nbsp;</td>
                    <td width="167">&nbsp;</td>
                    <td width="184" align="right" class="form_dado_texto">&nbsp;</td>
                    <td width="184">&nbsp;</td>
                    <td width="184" align="right" class="form_dado_texto">Total:</td>
                    <td width="184">
                  <input name="total" id="total" type="text" class="textInput" value="<%response.Write(formatnumber(soma,2))%>" size="20" maxlength="20" readonly></td>
                    <td width="184">&nbsp;</td>
                  </tr>
                  </table></td>
                                             
              </tr>-->
                <tr bgcolor="#FFFFFF">
                <td colspan="8">                  <input name="qtd_itens" type="hidden" id="qtd_itens" value="<%response.Write(linhas)%>">
                  <input name="itens_criados" type="hidden" id="itens_criados" value="<%response.Write(linhas)%>">
</td>
              </tr>
                <tr bgcolor="#FFFFFF">
                <td colspan="8"><div align="center">
                    <table width="1000" border="0" align="center" cellspacing="0">
                    <tr>
                        <td height="24" colspan="3"><hr></td>
                      </tr>
                    <tr>
                        <td width="33%"><div align="center">
                            <% if ori=2 or ori="2" then %>
                            <input name="alterar" type="submit" class="botao_cancelar" id="alterar" onClick="MM_goToURL('parent','index.asp?nvg=WA-AL-MA-AOC');return document.MM_returnValue" value="Voltar">
                            <% else%>
                            <input name="alterar" type="submit" class="botao_cancelar" id="alterar" onClick="MM_goToURL('parent','resumo.asp?opt=vt');return document.MM_returnValue" value="Voltar">
                            <%end if%>
                          </div></td>
                        <td width="34%">&nbsp;</td>
                        <td width="33%"><div align="center">
                            <input name="Submit" type="submit" class="botao_prosseguir" value="Confirmar">
                          </div></td>
                      </tr>
                  </table>
                    <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> </font></div></td>
              </tr>
              </table>       
               </form>
               </td>
          </tr>

<%
	if opt="erra" or  opt="erri" then
	
	
			D_chavearray=split(chave,"-")
			D_sistema=D_chavearray(0)
			D_modulo=D_chavearray(1)
			D_setor=D_chavearray(2)
			D_funcao=D_chavearray(3)
	
		Set RSD = Server.CreateObject("ADODB.Recordset")
		CONEXAOD = "DELETE * from TBA_Msg_Erro WHERE CO_Sistema = '"&D_sistema&"' and CO_Modulo = '"&D_modulo&"' and CO_Setor = '"&D_setor&"' and CO_Funcao = '"&D_funcao&"' and CO_Usuario = "&session("co_user")
		Set RSD = CON99.Execute(CONEXAOD)	
		
		Set RSD = Server.CreateObject("ADODB.Recordset")
		CONEXAOD = "DELETE * from TBA_Mov_Estoque WHERE NU_Pedido = "&cod_nf
		Set RSD = CON99.Execute(CONEXAOD)	
		
		Set RSD = Server.CreateObject("ADODB.Recordset")
		CONEXAOD = "DELETE * from TBA_Mov_Estoque_Item WHERE NU_Pedido = "&cod_nf
		Set RSD = CON99.Execute(CONEXAOD)	
	end if
opt=co_nf&"?"&da_nf

end if%>
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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