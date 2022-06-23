<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
      <% 

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
ori = request.QueryString("or")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

trava=session("trava")


obr=session("obr")
session("obr")=obr


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

mes_atual = DatePart("m", now) 
dia = DatePart("d", now) 

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
		

Call LimpaVetor2

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

function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
      </head>
      <body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="putFocusOn('nota_fiscal')">
      <%call cabecalho(nivel)
%>
              <form action="bd.asp?opt=inc" method="post" name="anf" id="anf" onSubmit="return checksubmit()">
      <table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
        <tr>
          <td height="10" colspan="5" class="tb_caminho"><font class="style-caminho">
            <%
	  response.Write(navega)

%>
            </font></td>
        </tr>
        <tr>
          <td height="10" colspan="5" valign="top"><%call mensagens(nivel,826,0,0) %></td>
        </tr>
          <tr>
            <td valign="top"><table width="1002" border="0" align="right" cellspacing="0" class="tb_corpo"
>
                <tr>
                <td width="100" class="tb_tit"
>Nota de Compra</td>
                <td colspan="4" class="tb_tit"
></td>
              </tr>
                <tr>
                <td width="100" height="10" align="right" class="form_dado_texto">Nota Fiscal:</td>
                <td width="232" valign="top"><input name="nota_fiscal" type="text" class="textInput" id="nota_fiscal" size="35" maxlength="30"></td>
                <td width="86" valign="top">&nbsp;</td>
                <td width="84" align="right"><span class="form_dado_texto">Data da Nota:</span></td>
                <td width="493" valign="top"><select name="dia_nf" id="dia_nf" class="select_style">
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
                    <%mes_atual=mes_atual*1
								if mes_atual="1" or mes_atual=1 then%>
                    <option value="1" selected>janeiro</option>
                    <% else%>
                    <option value="1">janeiro</option>
                    <%end if
								if mes_atual="2" or mes_atual=2 then%>
                    <option value="2" selected>fevereiro</option>
                    <% else%>
                    <option value="2">fevereiro</option>
                    <%end if
								if mes_atual="3" or mes_atual=3 then%>
                    <option value="3" selected>mar&ccedil;o</option>
                    <% else%>
                    <option value="3">mar&ccedil;o</option>
                    <%end if
								if mes_atual="4" or mes_atual=4 then%>
                    <option value="4" selected>abril</option>
                    <% else%>
                    <option value="4">abril</option>
                    <%end if
								if mes_atual="5" or mes_atual=5 then%>
                    <option value="5" selected>maio</option>
                    <% else%>
                    <option value="5">maio</option>
                    <%end if
								if mes_atual="6" or mes_atual=6 then%>
                    <option value="6" selected>junho</option>
                    <% else%>
                    <option value="6">junho</option>
                    <%end if
								if mes_atual="7" or mes_atual=7 then%>
                    <option value="7" selected>julho</option>
                    <% else%>
                    <option value="7">julho</option>
                    <%end if%>
                    <%if mes_atual="8" or mes_atual=8 then%>
                    <option value="8" selected>agosto</option>
                    <% else%>
                    <option value="8">agosto</option>
                    <%end if
								if mes_atual="9" or mes_atual=9 then%>
                    <option value="9" selected>setembro</option>
                    <% else%>
                    <option value="9">setembro</option>
                    <%end if
								if mes_atual="10" or mes_atual=10 then%>
                    <option value="10" selected>outubro</option>
                    <% else%>
                    <option value="10">outubro</option>
                    <%end if
								if mes_atual="11" or mes_atual=11 then%>
                    <option value="11" selected>novembro</option>
                    <% else%>
                    <option value="11">novembro</option>
                    <%end if
								if mes_atual="12" or mes_atual=12 then%>
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
                  </select></td>
              </tr>
                <tr>
                <td width="100" height="10" align="right" class="form_dado_texto">Fornecedor:</td>
                <td valign="top" bgcolor="#FFFFFF"><select name="fornecedor" class="select_style" id="fornecedor">
                    <option value="nulo" selected></option>
                    <%
		Set RS = Server.CreateObject("ADODB.Recordset")
		sql = "Select * From TB_Fornecedor order by NO_Fornecedor"
		RS.Open sql, CON9  
		
		while not RS.EOF 
		cod_for=RS("CO_Fornecedor")		
		nome_for=RS("NO_Fornecedor")		  
		
		%>
                    <option value="<%response.Write(cod_for)%>">
                  <%response.Write(nome_for)%>
                  </option>
                    <%
		RS.MOVENEXT
		WEND 		
				%>
                  </select></td>
                <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
                <td height="10" align="right" class="form_dado_texto">Valor:</td>
                <td valign="top" bgcolor="#FFFFFF"><input name="valor" type="text" class="textInput" id="valor" size="35" maxlength="30" onKeyDown="return (valorFormat(this))" onKeyUp="return (valorFormat(this))" onBlur="return (ValidaNumero(this.value, this.id));" ></td>
              </tr>
                <tr>
                <td colspan="5">&nbsp;</td>
              </tr>
                <tr>
                <td colspan="5" class="tb_tit"
>Composi&ccedil;&atilde;o da Nota Fiscal<input name="qtd_itens" type="hidden" id="qtd_itens" value="1">
                  <input name="itens_criados" type="hidden" id="itens_criados" value="1"></td>
              </tr>
                <tr>
                <td colspan="5"><table id="tblInnerHTML" width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="100" align="right" class="form_dado_texto"><input name="num_linha" type="hidden" id="num_linha" value="1">
                      1</td>
                    <td width="100" align="right" class="form_dado_texto">Item:</td>
                    <td width="167"><select name="item_fornecedor_1" class="select_style" id="item_fornecedor_1">
                        <option value="nulo" selected></option>
                        <%
		Set RS = Server.CreateObject("ADODB.Recordset")
		sql = "Select * From TB_Item order by NO_Item"
		RS.Open sql, CON9  
		
		while not RS.EOF 
		cod_item=RS("CO_Item")		
		nome_item=RS("NO_Item")		  
		
		%>
                        <option value="<%response.Write(cod_item)%>">
                      <%response.Write(nome_item)%>
                      </option>
                        <%
		RS.MOVENEXT
		WEND 		
				%>
                      </select></td>
                    <td width="184" align="right" class="form_dado_texto">Quantidade:</td>
                    <td width="184"><input name="quantidade_1" type="text" class="textInput" id="quantidade_1" onBlur="return (ValidaNumero(this.value, this.id));" onChange="return (produto(this))" value="1" size="15" maxlength="15" onFocus="javascript:this.form.quantidade_1.select();" ></td>
                    <td width="184" align="right" class="form_dado_texto">Valor Unit&aacute;rio:</td>
                    <td width="184"><input name="valor_1" type="text" class="textInput" id="valor_1" onChange="return (produto(this))" onKeyDown="return (currencyFormat(this))" onKeyUp="return (currencyFormat(this))" onBlur="return (ValidaNumero(this.value, this.id));" value="0" size="20" maxlength="20" onFocus="javascript:this.form.valor_1.select();">
                      <input name="aux_format" readonly type="hidden"><input name="produto_1" type="hidden" id="produto_1" value="0"></td>
                    <td width="184">                                
                    <div id="1"><a href="#"  onClick="addRow();changeImage(1)"><img src="../../../../img/add.png" alt="Adicionar Item" width="20" height="20" border = "0" ></a></div><!--onClick="addRow();hideImage(1);showImage('close_1');"       --></td>
                  </tr>
                  </table></td>
                  
              </tr>
                <tr bgcolor="#FFFFFF">
                <td colspan="5"><table id="tblInnerHTML" width="100%" border="0" cellspacing="0" cellpadding="0">
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
                    <td width="184"><input name="total" id="total" type="text" class="textInput" value="0" size="20" maxlength="20" readonly></td>
                    <td width="184">&nbsp;</td>
                  </tr>
                  </table></td>
                                             
              </tr>
                <tr bgcolor="#FFFFFF">
                <td colspan="5"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></div></td>
              </tr>
                <tr bgcolor="#FFFFFF">
                <td colspan="5"><div align="center">
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