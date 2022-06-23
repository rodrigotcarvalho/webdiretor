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
nota_fiscal_form=request.querystring("cod")

cod_nf = session("cod_nf")
dia_de = session("dia_de")
mes_de = session("mes_de")
ano_de = session("ano_de")
dia_ate = session("dia_ate")
mes_ate = session("mes_ate")
ano_ate = session("ano_ate")

session("cod_nf") = cod_nf
session("dia_de") = dia_de
session("mes_de") = mes_de
session("ano_de") = ano_de
session("dia_ate") = dia_ate
session("mes_ate") = mes_ate
session("ano_ate") = ano_ate

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")		

		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		

		Set CON9 = Server.CreateObject("ADODB.Connection") 
		ABRIR9 = "DBQ="& CAMINHO_ax & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON9.Open ABRIR9	


%>
<html>
<head>
<title>Web Diretor</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
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
{ 		var t,linha	;
		t = fld.id.split("_");
		linha = t[1];		
		var fator1 = document.getElementById('quantidade_'+linha)		
		var fator2 = document.getElementById('valor_'+linha)		
		var resultado = document.getElementById('produto_'+linha)	
		if (isNaN(parseFloat(fator1.value)) || isNaN(parseFloat(fator2.value))){
			}else{
			//invertendo do formato brasileiro para o americano
			convertido1 = fator1.value.replace( '.', '' );  
			convertido2 = fator2.value.replace( '.', '' );  				
			convertido1 = convertido1.replace( ',', '.' );  
			convertido2 = convertido2.replace( ',', '.' );  														
			resultado.value = convertido1*convertido2;  
	
			soma();
		}
}

function soma()  
{ 
//todas as linhas que foram criadas na tabela nessa sessão incluindo as que foram excluídas  		
		var total_itens = document.getElementById('itens_criados');
		
		var money = 0	
		var total = document.getElementById('total')
		total.value = total.value.replace( ',', '.' );	
	
	for (var i=1;i<=total_itens.value;i++)
	{ 
			var produto = document.getElementById('produto_'+i)
			//produto = produto.value.replace( ',', '.' );  
			if (produto){
			money = (produto.value*100)+money; 
			}
			
	}		var arredonda = money/100;  
			total.value = arredonda.toFixed(2)
			
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
	//As linhas são contadas a partir do zero 
	//enquanto na tabela estão numeradas a partir do 1
	var linha_da_tabela = ln-1
	//var linha_imagem = linha_da_tabela-1	
	//var identidade ="close_"+linha_imagem
	//alert(identidade)	
  var qtd_itens = document.getElementById('qtd_itens');	
  //linhas existentes na tabela atualmente
  var total_itens = document.getElementById('qtd_itens').value;	 	
  //todas as linhas que foram criadas na tabela nessa sessão incluindo as que foram excluídas  
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
  newCell0.width = 100
  newCell0.innerHTML = prox_item;
  
  var newCell1 = newRow.insertCell(1);
  newCell1.align = 'right'    
  newCell1.width = 100  
  newCell1.innerHTML = 'Item:';
  
  var newCell2 = newRow.insertCell(2);
  newCell2.align = 'left'    
  newCell2.width = 167    
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
  newCell3.width = 184        
  newCell3.innerHTML = 'Quantidade:';		
  
  var newCell4 = newRow.insertCell(4);
  newCell4.align = 'left'    
  newCell4.width = 184   
  newCell4.innerHTML = '<input name="quantidade_'+prox_item+'" type="text" value="1" class="textInput" id="quantidade_'+prox_item+'" size="15" maxlength="15" onBlur="return (ValidaNumero(this.value, this.id));" onChange="return (produto(this));" onFocus="javascript:this.form.quantidade_'+prox_item+'.select();">';  	
  
  var newCell5 = newRow.insertCell(5);
  newCell5.align = 'right' 
  newCell5.width = 184      
  newCell5.innerHTML = 'Valor Unit&aacute;rio:';	  	
  
  var newCell6 = newRow.insertCell(6);
  newCell6.align = 'left'   
  newCell6.width = 184    
  newCell6.innerHTML = '<input name="valor_'+prox_item+'" type="text" value="0" class="textInput" id="valor_'+prox_item+'" size="20" maxlength="20" onBlur="return (ValidaNumero(this.value, this.id));" onChange="return (produto(this));" onKeyDown="return (currencyFormat(this))" onKeyUp="return (currencyFormat(this))" onFocus="javascript:this.form.valor_'+prox_item+'.select();"><input name="aux_format" readonly type="hidden"><input name="produto_'+prox_item+'" type="hidden" id="produto_'+prox_item+'" value="0">';	 
  
  var newCell7 = newRow.insertCell(7);
  newCell7.align = 'left'  
  newCell7.width = 184     
  newCell7.innerHTML = '<div id="'+prox_item+'"><a href="#" onClick="addRow();changeImage('+prox_item+');"><img src="../../../../img/add.png" alt="Adicionar Item" width="20" height="20" border = "0"></a></div>';	  
  
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
 if (document.anf.nota_fiscal.value == "")
  {    alert("Por favor digite uma nota fiscal!")
   document.anf.nota_fiscal.focus()
    return false 
 }
 
 if (document.anf.fornecedor.value == "nulo")
  {    alert("Por favor selecione um fornecedor!")
   document.anf.fornecedor.focus()
    return false
 } 
 
    if (document.anf.valor.value == 0 || document.anf.valor.value == "")
  {    alert("Por favor digite um valor diferente de zero!")
    document.anf.valor.focus()
    return false
  } 
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
		var valor_i = document.getElementById('valor_'+i)				

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
		
		if (valor_i){
		 if (valor_i.value == 0 || valor_i.value == "")
		  {    alert("Por favor digite um valor diferente de zero!")
		   valor_i.focus()
			return false
		 } 
		}				
		
	} 
 //aula = document.busca.aula.value;
//     if (aula.length > 3)
//  {    alert("O valor do campo Aula deve possuir menos que 3 caracteres")
//    document.busca.aula.focus()
//    return false
//  }

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
//-->
</script>
</head>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<% if opt="exc"	then%>
<form action="bd.asp?opt=exc" method="post" name="busca" id="busca">
<%else%>
<form action="bd.asp?opt=alt" method="post" name="anf" id="anf" onSubmit="return checksubmit()">
<%end if%>
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
exclui_nota_form = nota_fiscal_form
%>		  
          <tr>
      <td valign="top"> 
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>            <tr> 
              
    <td width="766" height="10" colspan="4" valign="top"> 
      <%call mensagens(nivel,828,0,0) %>
    </td>
			  </tr>
          <tr> 
            <td height="10" class="tb_tit"
>Notas Fiscais a serem exclu&iacute;das</td>
          </tr>
          <tr> 
            <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
                <tr class="tb_subtit"> 
                  <td width="20" height="10"><div align="center">
                    <input name="exclui_nf" type="hidden" id="exclui_nf" value="<%response.write(exclui_nota_form)%>">
                  </div></td>
                  <td><div align="left">Nota Fiscal</div></td>
                  <td align="center">Data Compra</td>
                  <td><div align="left">Fornecedor</div></td>
                  <td width="100" align="right">Valor da Nota</td>
                  <td width="220" align="center">Conferido por</td>
                  <td width="220" align="center">Registrado por</td>
                </tr>
                <%
'response.Write(">>"&exclui_ocorrencia)				
check = 2	
exclui_nota_fiscal = replace(nota_fiscal_form,"$!$","/")		
vetorExclui = split(exclui_nota_fiscal,", ")
conta_ocorr=0
for i =0 to ubound(vetorExclui)

exclui = split(vetorExclui(i),"?")

'obr=cod&"?"&da_ocorrencia&"?"&ho_ocorrencia&"?"&co_ocorrencia
cod_nf = exclui(0)
data_nf= exclui(1)

				
dados_data=split(data_nf,"/")
dia=dados_data(0)
mes=dados_data(1)
ano=dados_data(2)




data_nf_cons=mes&"/"&dia&"/"&ano

		
		Set RS = Server.CreateObject("ADODB.Recordset")
		 SQL = "SELECT * FROM TB_NFiscais_Compra, TB_Fornecedor WHERE TB_Fornecedor.CO_Fornecedor = TB_NFiscais_Compra.CO_Fornecedor AND NU_NotaF ='"& cod_nf&"' AND (DA_NotaF BETWEEN #"&data_nf_cons&"# AND #"&data_nf_cons&"#)"
		RS.Open SQL, CON9

  if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if 
  
co_nf=RS("NU_NotaF")
da_nf=RS("DA_NotaF")
co_fornecedor=RS("CO_Fornecedor")
valor_nf=RS("VA_NotaF")
observacao=RS("TX_Observa")
co_usu_conf=RS("CO_Usuario_Conf")
co_usu_reg=RS("CO_Usuario_Reg")


data_split= Split(da_nf,"/")
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

if co_fornecedor="" or isnull(co_fornecedor) then
	no_fornecedor=""
else

	Set RSnom = Server.CreateObject("ADODB.Recordset")
	SQLnom = "SELECT NO_Fornecedor FROM TB_Fornecedor Where CO_Fornecedor="&co_fornecedor
	RSnom.Open SQLnom, CON9
	
	if RSnom.EOF then
		no_fornecedor=""	
	else
		no_fornecedor=RSnom("NO_Fornecedor")
	end if	
end if


if co_usu_conf="" or isnull(co_usu_conf) then
	no_conferidor=""
else

		Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& co_usu_conf
		RSu.Open SQLu, CON

	IF RSu.EOF then
		no_conferidor=""
	else
		no_conferidor=RSu("NO_Usuario")
	end if		
end if
		
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

opt=co_nf&"?"&da_nf

%>
                <tr class="<%=cor%>"> 
                  <td width="20">&nbsp;</td>
                  <td><%response.Write(co_nf)%></td>
                  <td align="center">
                    <%response.Write(da_show)%>
                  </td>
                  <td><%response.Write(no_fornecedor)%>
                  <div align="left"></div></td>
					<td width="100" align="right"><%response.Write(formatnumber(valor_nf,2))%>
                  <td><div align="center">
                    <%response.Write(no_conferidor)%>
                  </div></td>
                  <td><div align="center">
                    <%response.Write(no_registrador)%>
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
                        <input name="Submit" type="submit" class="botao_prosseguir" value="Confirmar">
                      </div></td>
                  </tr>
                </table>
                <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> </font></div></td>
          </tr>
        </table></td>
    </tr>

<%else
' OPT = ALT

nota_fiscal_alt = replace(nota_fiscal_form,"$!$","/")		
vetorAltera = split(nota_fiscal_alt,"?")

cod_nf = vetorAltera(0)
data_nf= vetorAltera(1)
dados_msg = cod_nf&"?"&data_nf
				
dados_data=split(data_nf,"/")
dia=dados_data(0)
mes=dados_data(1)
ano=dados_data(2)




data_nf_cons=mes&"/"&dia&"/"&ano

		
		Set RS = Server.CreateObject("ADODB.Recordset")
		 SQL = "SELECT * FROM TB_NFiscais_Compra, TB_Fornecedor WHERE TB_Fornecedor.CO_Fornecedor = TB_NFiscais_Compra.CO_Fornecedor AND NU_NotaF ='"& cod_nf&"' AND (DA_NotaF BETWEEN #"&data_nf_cons&"# AND #"&data_nf_cons&"#)"
		RS.Open SQL, CON9

  if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if 
  
co_nf=RS("NU_NotaF")
da_nf=RS("DA_NotaF")
co_fornecedor=RS("CO_Fornecedor")
valor_nf=RS("VA_NotaF")
observacao=RS("TX_Observa")
co_usu_conf=RS("CO_Usuario_Conf")
co_usu_reg=RS("CO_Usuario_Reg")


data_split= Split(da_nf,"/")
dia=data_split(0)
mes=data_split(1)
ano=data_split(2)


dia=dia*1
mes=mes*1
ano=ano*1


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

if co_fornecedor="" or isnull(co_fornecedor) then
	no_fornecedor=""
else

	Set RSnom = Server.CreateObject("ADODB.Recordset")
	SQLnom = "SELECT NO_Fornecedor FROM TB_Fornecedor Where CO_Fornecedor="&co_fornecedor
	RSnom.Open SQLnom, CON9
	
	if RSnom.EOF then
		no_fornecedor=""	
	else
		no_fornecedor=RSnom("NO_Fornecedor")
	end if	
end if


if co_usu_conf="" or isnull(co_usu_conf) then
	no_conferidor=""
else

		Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& co_usu_conf
		RSu.Open SQLu, CON

	IF RSu.EOF then
		no_conferidor=""
	else
		no_conferidor=RSu("NO_Usuario")
	end if		
end if
		
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

opt=co_nf&"?"&da_nf




%>

          <tr>   
            <td valign="top"><table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
            <tr> 
              
    <td height="10" colspan="7" valign="top"> 
      <%call mensagens(nivel,9708,0,0) %>
    </td>
			  </tr>   
            <tr> 
              
    <td height="10" colspan="7" valign="top"> 
      <%call mensagens(nivel,645,0,"R22") %>
    </td>
			  </tr>                    
                <tr>
                <td colspan="7" class="tb_tit"
>Nota de Compra</td>
              </tr>
                <tr>
                <td width="91" height="10" align="right" class="form_dado_texto">Nota Fiscal:</td>
                <td width="219"  class="form_dado_texto"><%response.Write(co_nf)%>
                  <input name="nota_fiscal" type="hidden" id="nota_fiscal" size="35" maxlength="30" value="<%response.Write(co_nf)%>"></td>
                <td width="28" valign="top">&nbsp;</td>
                <td width="119" align="right" class="form_dado_texto">Data da Nota:</td>
                <td width="202" class="form_dado_texto">
                  <input name="dia_nf" type="hidden" id="dia_nf" size="35" maxlength="30" value="<%response.Write(dia)%>">	
                  <input name="mes_nf" type="hidden" id="mes_nf" size="35" maxlength="30" value="<%response.Write(mes)%>">                  <input name="ano_nf" type="hidden" id="ano_nf" size="35" maxlength="30" value="<%response.Write(ano)%>">                  			
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
                <td width="90" align="right" class="form_dado_texto">Valor:</td>
                <td width="239" align="left" class="form_dado_texto"><input name="valor" type="text" class="textInput" id="valor" size="35" maxlength="30" onKeyDown="return (valorFormat(this))" onKeyUp="return (valorFormat(this))" onBlur="return (ValidaNumero(this.value, this.id));" value="<%response.Write(formatnumber(valor_nf,2))%>"></td>
              </tr>              
                <tr>
                <td width="91" height="10" align="right" class="form_dado_texto">Fornecedor:</td>
                <td valign="top" bgcolor="#FFFFFF"><select name="fornecedor" class="select_style" id="fornecedor">
                    <%
		Set RS = Server.CreateObject("ADODB.Recordset")
		sql = "Select * From TB_Fornecedor order by NO_Fornecedor"
		RS.Open sql, CON9  
		
		while not RS.EOF 
		cod_for=RS("CO_Fornecedor")		
		nome_for=RS("NO_Fornecedor")
		
		if isnumeric(cod_for) then		
			cod_for = cod_for*1
		end if	
		if isnumeric(co_fornecedor) then		
			co_fornecedor	= co_fornecedor*1
		end if					  
		 
		 if co_fornecedor = cod_for then
		 	selected_fornecedor = "selected"
		 else
		 	selected_fornecedor = ""		 
		 end if
		%>
                    <option value="<%response.Write(cod_for)%>" <%response.Write(selected_fornecedor)%>>
                  <%response.Write(nome_for)%>
                  </option>
                    <%
		RS.MOVENEXT
		WEND 		
				%>
                  </select></td>
                <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                <td height="10" align="right" class="form_dado_texto">Conferido por:</td>
                <td colspan="3" valign="top" bgcolor="#FFFFFF"><select name="cp" class="select_style_fixo_1" id="cp" >
                          <option value="nulo" SELECTED>
                  </option>                
                          <%
		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT CO_Usuario,NO_Usuario FROM TB_Usuario Where CO_Usuario<>0 order by NO_Usuario"
		RS1.Open SQL1, CON

while not RS1.EOF
	co_usuario=RS1("CO_Usuario")
	no_usuario=RS1("NO_Usuario")
	
	co_usu_conf=co_usu_conf*1	
	co_usuario=co_usuario*1
	if co_usu_conf=co_usuario then
		cp_selected="SELECTED"
	ELSE
		cp_selected=""
	END IF	
	
	%>
                  <option value="<%response.Write(co_usuario)%>" <%response.Write(cp_selected)%>>
                            <%response.Write(no_usuario)%>
                            </option>
                  <%
RS1.movenext
Wend
%>
                  </select></td>
              </tr>
                <tr>
                <td colspan="7">&nbsp;</td>
              </tr>
                <tr>
                <td colspan="7" class="tb_tit"
>Composi&ccedil;&atilde;o da Nota Fiscal</td>
              </tr>
                <tr>
                <td colspan="7"><table id="tblInnerHTML" width="100%" border="0" cellspacing="0" cellpadding="0">
<% 		

		Set RSC = Server.CreateObject("ADODB.Recordset")
		SQLC = "Select COUNT(CO_Item) as Total From TB_NFiscais_Compra_Item where NU_NotaF ='"& cod_nf&"'"
		RSC.Open SQLC, CON9	 
		
		if RSC.EOF then
			total = 1	
		else
			total = RSC("Total")		
		end if


		Set RSI = Server.CreateObject("ADODB.Recordset")
		SQLI = "Select CO_Item, QT_Item, VA_Unitario From TB_NFiscais_Compra_Item where NU_NotaF ='"& cod_nf&"' GROUP BY CO_Item, QT_Item, VA_Unitario"
		RSI.Open SQLI, CON9	  
		
		linhas = 0    
		soma=0
		While not RSI.EOF         
			linhas = linhas+1

			co_item = RSI("CO_Item")			
			quantidade_item = RSI("QT_Item")
			valor_unit = RSI("VA_Unitario")	
				
			if isnull(quantidade_item) or quantidade_item ="" then
				quantidade_item = 0
			end if
			if isnull(valor_unit) or valor_unit ="" then
				valor_unit = 0
			end if	
			produto = quantidade_item*valor_unit
			soma = soma+produto  					
%>		              
                    <tr>
                      <td width="100" align="right" class="form_dado_texto"><input name="num_linha" type="hidden" id="num_linha" value="<%response.Write(linhas)%>">
                      <%response.Write(linhas)%></td>
                    <td width="100" align="right" class="form_dado_texto">Item:</td>
                    <td width="167"><select name="item_fornecedor_<%response.Write(linhas)%>" class="select_style" id="item_fornecedor_<%response.Write(linhas)%>">
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
                    <td width="184" align="right" class="form_dado_texto">Quantidade:</td>
                    <td width="184"><input name="quantidade_<%response.Write(linhas)%>" type="text" class="textInput" id="quantidade_<%response.Write(linhas)%>" onBlur="return (ValidaNumero(this.value, this.id));" onChange="return (produto(this))" value="<%response.Write(quantidade_item)%>" size="15" maxlength="15" onFocus="javascript:this.form.quantidade_<%response.Write(linhas)%>.select();" ></td>
                    <td width="184" align="right" class="form_dado_texto">Valor Unit&aacute;rio:</td>
                    <td width="184"><input name="valor_<%response.Write(linhas)%>" type="text" class="textInput" id="valor_<%response.Write(linhas)%>" onChange="return (produto(this))" onKeyDown="return (currencyFormat(this))" onKeyUp="return (currencyFormat(this))" onBlur="return (ValidaNumero(this.value, this.id));" value="<%response.Write(valor_unit)%>" size="20" maxlength="20" onFocus="javascript:this.form.valor_<%response.Write(linhas)%>.select();">
                      <input name="aux_format" readonly type="hidden">
                      <input name="produto_<%response.Write(linhas)%>" type="hidden" id="produto_<%response.Write(linhas)%>" value="<%response.Write(produto)%>"></td>
                    <td width="184">                                
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
                <tr bgcolor="#FFFFFF">
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
                  <input name="qtd_itens" type="hidden" id="qtd_itens" value="<%response.Write(linhas)%>">
                  <input name="itens_criados" type="hidden" id="itens_criados" value="<%response.Write(linhas)%>">
                  <input name="total" id="total" type="text" class="textInput" value="<%response.Write(formatnumber(soma,2))%>" size="20" maxlength="20" readonly></td>
                    <td width="184">&nbsp;</td>
                  </tr>
                  </table></td>
                                             
              </tr>
                <tr bgcolor="#FFFFFF">
                <td colspan="7"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></div></td>
              </tr>
                <tr bgcolor="#FFFFFF">
                <td colspan="7"><div align="center">
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
              </table></td>
          </tr>

<%end if%>
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