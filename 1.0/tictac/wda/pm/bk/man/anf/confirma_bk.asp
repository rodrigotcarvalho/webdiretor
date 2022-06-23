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
exclui_nota_fiscal=request.querystring("cod")

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
	changeAction(); 
}

function currencyFormat(fld) {
  valorFormat(fld);
  produto(fld);
}
//function produto(linha)  
//{ 	
//		var convertido1, convertido2		
//		var fator1 = document.getElementById('quantidade_'+linha)		
//		var fator2 = document.getElementById('valor_'+linha)		
//		var resultado = document.getElementById('produto_'+linha)	
//		convertido1 = fator1.value.replace( ',', '.' );  
//		convertido2 = fator2.value.replace( ',', '.' );  												
//		resultado.value = convertido1*convertido2;  
//
//	soma();
//}
function produto(fld)  
{ 		var t,linha	;
		t = fld.id.split("_");
		linha = t[1];		
		var fator1 = document.getElementById('quantidade_'+linha)		
		var fator2 = document.getElementById('valor_'+linha)		
		var resultado = document.getElementById('produto_'+linha)	
		convertido1 = fator1.value.replace( ',', '.' );  
		convertido2 = fator2.value.replace( ',', '.' );  												
		resultado.value = convertido1*convertido2;  

	soma();
}
function soma()  
{ 

		var total_itens = document.getElementById('qtd_itens');
		
		var money = 0	
		var total = document.getElementById('total')
		
	for (var i=1;i<=total_itens.value;i++)
	{ 
			var produto = document.getElementById('produto_'+i)
			//produto = produto.value.replace( ',', '.' );  
			money = (produto.value*100)+money;  	
			
	}		
			total.value = money/100;  		
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
	var linha_a_apagar = ln-1
	var linha_imagem = linha_a_apagar-1	
	var identidade ="close_"+linha_imagem
	alert(identidade)	
	document.getElementById("tblInnerHTML").deleteRow(linha_a_apagar);
	document.getElementById(identidade).onclick = function () { deleteRow(linha_a_apagar); ShowImage(linha_imagem);};
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
  

  var rowCount = table.rows.length;  

  qtd_itens.value = rowCount;  
  
  var newCell0 = newRow.insertCell(0); 
  newRow.className = 'form_dado_texto';  
  newCell0.align = 'right'  
  newCell0.innerHTML = rowCount;
  
  var newCell1 = newRow.insertCell(1);
  newCell1.align = 'right'    
  newCell1.innerHTML = 'Item:';
  
  var newCell2 = newRow.insertCell(2);
  newCell2.align = 'left'    
  newCell2.innerHTML = '<select name="item_fornecedor_'+rowCount+'" class="select_style" id="item_fornecedor_'+rowCount+'"><option value="nulo" selected></option><%
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
  newCell3.innerHTML = 'Quantidade:';		
  
  var newCell4 = newRow.insertCell(4);
  newCell4.align = 'left'    
  newCell4.innerHTML = '<input name="quantidade_'+rowCount+'" type="text" value="1" class="textInput" id="quantidade_'+rowCount+'" size="15" maxlength="15" onChange="return (produto(this));" onFocus="javascript:this.form.quantidade_'+rowCount+'.select();">';  	
  
  var newCell5 = newRow.insertCell(5);
  newCell5.align = 'right'    
  newCell5.innerHTML = 'Valor Unit&aacute;rio:';	  	
  
  var newCell6 = newRow.insertCell(6);
  newCell6.align = 'left'    
  newCell6.innerHTML = '<input name="valor_'+rowCount+'" type="text" value="0" class="textInput" id="valor_'+rowCount+'" size="20" maxlength="20" onChange="return (produto(this));" onKeyDown="return (currencyFormat(this))" onKeyUp="return (currencyFormat(this))" onFocus="javascript:this.form.valor_'+rowCount+'.select();"><input name="aux_format" readonly type="hidden"><input name="produto_'+rowCount+'" type="hidden" id="produto_'+rowCount+'" value="0">';	 
  
  var newCell7 = newRow.insertCell(7);
  newCell7.align = 'left'    
  newCell7.innerHTML = '<div id="'+rowCount+'"><a href="#" onClick="addRow();hideImage('+rowCount+');"><img src="../../../../img/add.png" alt="Adicionar Item" width="20" height="20" border = "0"></a></div>';	  
  
 putFocusOn('item_fornecedor_'+rowCount);  
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
exclui_nota_form = exclui_nota_fiscal
%>		  
        <form action="bd.asp?opt=exc" method="post" name="busca" id="busca">
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
exclui_nota_fiscal = replace(exclui_nota_fiscal,"$!$","/")		
vertorExclui = split(exclui_nota_fiscal,", ")
conta_ocorr=0
for i =0 to ubound(vertorExclui)

exclui = split(vertorExclui(i),"?")

'obr=cod&"?"&da_ocorrencia&"?"&ho_ocorrencia&"?"&co_ocorrencia
cod_nf = exclui(0)
data_nf= exclui(1)

				
dados_data=split(data_nf,"/")
dia=dados_data(0)
mes=dados_data(1)
ano=dados_data(2)




data_nf_cons=mes&"/"&dia&"/"&ano

h=h*1
m=m*1

			
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
</form>
<%else

	if opt = "inc" then
		qtd_itens_original = request.Form("qtd_itens")	
		nota_fiscal = request.Form("nota_fiscal")	
		fornecedor = request.Form("fornecedor")			
		dia_nf = request.Form("dia_nf")	
		mes_nf = request.Form("mes_nf")	
		ano_nf = request.Form("ano_nf")	
		valor = request.Form("valor")													
		total = request.Form("total")	
	elseif opt = "alt" then	
	
	end if






%>
<form action="bd.asp?opt=inc" method="post" name="anf" id="busca" onSubmit="return checksubmit()">
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
                <td width="232" valign="top"><input name="nota_fiscal" type="text" class="textInput" id="nota_fiscal" size="35" maxlength="30" value="<%response.Write(nota_fiscal)%>"></td>
                <td width="86" valign="top">&nbsp;</td>
                <td width="84" align="right"><span class="form_dado_texto">Data da Nota:</span></td>
                <td width="493" valign="top"><select name="dia_nf" id="dia_nf" class="select_style">
                    <% 
							 For i =1 to 31
							 dia_nf=dia_nf*1
							 if dia_nf=i then 
								if dia_nf<10 then
								dia_nf="0"&dia_nf
								end if
							 %>
                    <option value="<%response.Write(i)%>" selected>
                  <%response.Write(dia_nf)%>
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
                    <%mes_nf=mes_nf*1
								if mes_nf="1" or mes_nf=1 then%>
                    <option value="1" selected>janeiro</option>
                    <% else%>
                    <option value="1">janeiro</option>
                    <%end if
								if mes_nf="2" or mes_nf=2 then%>
                    <option value="2" selected>fevereiro</option>
                    <% else%>
                    <option value="2">fevereiro</option>
                    <%end if
								if mes_nf="3" or mes_nf=3 then%>
                    <option value="3" selected>mar&ccedil;o</option>
                    <% else%>
                    <option value="3">mar&ccedil;o</option>
                    <%end if
								if mes_nf="4" or mes_nf=4 then%>
                    <option value="4" selected>abril</option>
                    <% else%>
                    <option value="4">abril</option>
                    <%end if
								if mes_nf="5" or mes_nf=5 then%>
                    <option value="5" selected>maio</option>
                    <% else%>
                    <option value="5">maio</option>
                    <%end if
								if mes_nf="6" or mes_nf=6 then%>
                    <option value="6" selected>junho</option>
                    <% else%>
                    <option value="6">junho</option>
                    <%end if
								if mes_nf="7" or mes_nf=7 then%>
                    <option value="7" selected>julho</option>
                    <% else%>
                    <option value="7">julho</option>
                    <%end if%>
                    <%if mes_nf="8" or mes_nf=8 then%>
                    <option value="8" selected>agosto</option>
                    <% else%>
                    <option value="8">agosto</option>
                    <%end if
								if mes_nf="9" or mes_nf=9 then%>
                    <option value="9" selected>setembro</option>
                    <% else%>
                    <option value="9">setembro</option>
                    <%end if
								if mes_nf="10" or mes_nf=10 then%>
                    <option value="10" selected>outubro</option>
                    <% else%>
                    <option value="10">outubro</option>
                    <%end if
								if mes_nf="11" or mes_nf=11 then%>
                    <option value="11" selected>novembro</option>
                    <% else%>
                    <option value="11">novembro</option>
                    <%end if
								if mes_nf="12" or mes_nf=12 then%>
                    <option value="12" selected>dezembro</option>
                    <% else%>
                    <option value="12">dezembro</option>
                    <%end if%>
                  </select>
                    /
                    <select name="ano_nf" class="select_style" id="ano_de">
                    <%
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Ano_Letivo order by NU_Ano_Letivo"
		RS0.Open SQL0, CON
		while not RS0.EOF 
			ano_bd=RS0("NU_Ano_Letivo")
			
			ano_nf=ano_nf*1
			ano_bd=ano_bd*1

				if ano_nf=ano_bd then%>
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
                    <% if fornecedor = "nulo" then
						selected_nulo = "selected"
					   else
						selected_nulo = ""					   
					   end if%>                  
                        <option value="nulo" <%response.Write(selected_nulo)%>></option>                
                    <%
		Set RS = Server.CreateObject("ADODB.Recordset")
		sql = "Select * From TB_Fornecedor order by NO_Fornecedor"
		RS.Open sql, CON9  
		
		while not RS.EOF 
			cod_for=RS("CO_Fornecedor")		
			nome_for=RS("NO_Fornecedor")	
			
			cod_for = cod_for*1
			if fornecedor <> "nulo" then
				fornecedor = fornecedor*1
				if cod_for = fornecedor then
					selected_fornecedor = "selected"
				else
					selected_fornecedor = ""					   
				end if	
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
                <td height="10" align="right" class="form_dado_texto">Valor:</td>
                <td valign="top" bgcolor="#FFFFFF"><input name="valor" type="text" class="textInput" id="valor" size="35" maxlength="30" onKeyDown="return (valorFormat(this))" onKeyUp="return (valorFormat(this))" value="<%response.Write(valor)%>"></td>
              </tr>
                <tr>
                <td colspan="5"><span class="form_dado_texto">
                  <input name="qtd_itens" type="hidden" id="qtd_itens" value="<%response.Write(qtd_itens_original)%>">
                <input name="aux_format" readonly type="hidden">
                </span></td>
              </tr>
                <tr>
                <td colspan="5" class="tb_tit"
>Composi&ccedil;&atilde;o da Nota Fiscal</td>
              </tr>
                <tr>
                <td colspan="5">
                <table id="tblInnerHTML" width="100%" border="0" cellspacing="0" cellpadding="0">
                <% for ln = 1 to qtd_itens_original 
				
						if opt = "inc" then			
							item_fornecedor = request.Form("item_fornecedor_"&ln)	
							quantidade = request.Form("quantidade_"&ln)	
							valor = request.Form("valor_"&ln)	
							produto = request.Form("produto_"&ln)															
						end if
				%>                
                    <tr>
                      <td width="100" align="right" class="form_dado_texto"><%response.Write(ln)%></td>
                    <td width="100" align="right" class="form_dado_texto">Item:</td>
                    <td width="167"><select name="item_fornecedor_<%response.Write(ln)%>" class="select_style" id="item_fornecedor_<%response.Write(ln)%>">
                    <% if item_fornecedor = "nulo" then
						selected_nulo = "selected"
					   else
						selected_nulo = ""					   
					   end if%>                  
                        <option value="nulo" <%response.Write(selected_nulo)%>></option>
                        <%
		Set RS = Server.CreateObject("ADODB.Recordset")
		sql = "Select * From TB_Item order by NO_Item"
		RS.Open sql, CON9  
		
		while not RS.EOF 
			cod_item=RS("CO_Item")		
			nome_item=RS("NO_Item")		  
			cod_item = cod_item*1
			if item_fornecedor <> "nulo" then			
				item_fornecedor = item_fornecedor*1
				if cod_item = item_fornecedor then
					selected_item = "selected"
				else
					selected_item = ""					   
				end if  
			else
				selected_item = ""					
			end if	  		
		%>
                        <option value="<%response.Write(cod_item)%>" <%response.Write(selected_item)%>>
                      <%response.Write(nome_item)%>
                      </option>
                        <%
		RS.MOVENEXT
		WEND 		
				%>
                      </select></td>
                    <td width="184" align="right" class="form_dado_texto">Quantidade:</td>
                    <td width="184"><input name="quantidade_<%response.Write(ln)%>" type="text" class="textInput" id="quantidade_<%response.Write(ln)%>" onChange="return (produto(this))" value="<%response.Write(quantidade)%>" size="15" maxlength="15" onFocus="javascript:this.form.quantidade_<%response.Write(ln)%>.select();" ></td>
                    <td width="184" align="right" class="form_dado_texto">Valor Unit&aacute;rio:</td>
                    <td width="184"><input name="valor_<%response.Write(ln)%>" type="text" class="textInput" id="valor_<%response.Write(ln)%>" onChange="return (produto(this))" onKeyDown="return (currencyFormat(this))" onKeyUp="return (currencyFormat(this))" value="<%response.Write(valor)%>" size="20" maxlength="20" onFocus="javascript:this.form.valor_<%response.Write(ln)%>.select();"><input name="produto_<%response.Write(ln)%>" type="hidden" id="produto_<%response.Write(ln)%>" value="<%response.Write(produto)%>"></td>
                    <td width="184">
                    <table width="50" border="0" cellpadding="0" cellspacing="0">
                    <tr><td align="center">
                    <%ln = ln*1
					  qtd_itens_original = qtd_itens_original*1
					  
					  if ln<>qtd_itens_original then
						visibilidade = "style=""visibility: hidden"""
						onclick = "deleteRow("&ln&");"
					  else
					  	visibilidade = ""
						onclick = "deleteRow("&ln&");ShowImage("&ln-1&");"						
					  end if	
					%>
                    <a id="close_<%response.Write(ln)%>" href="#" onClick="<%response.Write(onclick)%>">
                    <img src="../../../../img/close.png" alt="Excluir Item" width="20" height="20" border = "0" ></a></td>
                    <td align="center">                    
                    <div id="<%response.Write(ln)%>" <%response.Write(visibilidade)%>><a href="#" onClick="addRow();hideImage(<%response.Write(ln)%>);"><img src="../../../../img/add.png" alt="Adicionar Item" width="20" height="20" border = "0" ></a></div></td></tr></table></td>
                  </tr>
			<%next%>                    
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
                    <td width="184"><input name="total" id="total" type="text" class="textInput" value="<%response.Write(total)%>" size="20" maxlength="20" readonly></td>
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
                            <% if opt = "inc" then %>
                            <input name="alterar" type="submit" class="botao_cancelar" id="alterar" onClick="MM_goToURL('parent','inclui.asp');return document.MM_returnValue" value="Voltar">
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
        </form>
<%end if%>
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