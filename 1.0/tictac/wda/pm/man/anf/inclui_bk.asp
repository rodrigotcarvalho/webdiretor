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


<!--
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
function currencyFormat(fld) {
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
  document.frm2.teste.value = aux;
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

//function addRow() {
//    /* Declare variables */
//    var elements, templateRow, rowCount, row, className, newRow, element;
//    var i, s, t;
//    
//    /* Get and count all "tr" elements with class="row".    The last one will
//     * be serve as a template. */
//    if (!document.getElementsByTagName)
//        return false; /* DOM not supported */
//    elements = document.getElementsByTagName("tr");
//    templateRow = null;
//    rowCount = 0;
//    for (i = 0; i < elements.length; i++) {
//        row = elements.item(i);
//        
//        /* Get the "class" attribute of the row. */
//        className = null;
//        if (row.getAttribute)
//            className = row.getAttribute('class')
//        if (className == null && row.attributes) {    // MSIE 5
//            /* getAttribute('class') always returns null on MSIE 5, and
//             * row.attributes doesn't work on Firefox 1.0.    Go figure. */
//            className = row.attributes['class'];
//            if (className && typeof(className) == 'object' && className.value) {
//                // MSIE 6
//                className = className.value;
//            }
//        } 
//        
//        /* This is not one of the rows we're looking for.    Move along. */
//        if (className != "row_to_clone")
//            continue;
//        
//        /* This *is* a row we're looking for. */
//        templateRow = row;
//        rowCount++;
//    }
//    if (templateRow == null)
//        return false; /* Couldn't find a template row. */
//    
//    /* Make a copy of the template row */
//    newRow = templateRow.cloneNode(true);
//
//    /* Change the form variables e.g. price[x] -> price[rowCount] */
//    elements = newRow.getElementsByTagName("input");
//    for (i = 0; i < elements.length; i++) {
//        element = elements.item(i);
//        s = null;
//        s = element.getAttribute("name");
//        if (s == null)
//            continue;
//        t = s.split("[");
//        if (t.length < 2)
//            continue;
//        s = t[0] + "[" + rowCount.toString() + "]";
//        element.setAttribute("name", s);
//        element.value = "";
//    }
//    
//    /* Add the newly-created row to the table */
//    templateRow.parentNode.appendChild(newRow);
//    return true;
//}
		function addRow(tableID) {

			var table = document.getElementById(tableID);

			var rowCount = table.rows.length;
			var row = table.insertRow(rowCount);

//			var cell1 = row.insertCell(0);
//			var element1 = document.createElement("input");
//			element1.type = "checkbox";
//			element1.name="chkbox[]";
//			cell1.appendChild(element1);

			var cell1 = row.insertCell(0);
			cell1.innerHTML = rowCount + 1;

			var cell3 = row.insertCell(2);
			var element2 = document.createElement("input");
			element2.type = "text";
			element2.name = "txtbox[]";
			cell3.appendChild(element2);


		}

		function deleteRow(tableID) {
			try {
			var table = document.getElementById(tableID);
			var rowCount = table.rows.length;

			for(var i=0; i<rowCount; i++) {
				var row = table.rows[i];
				var chkbox = row.cells[0].childNodes[0];
				if(null != chkbox && true == chkbox.checked) {
					table.deleteRow(i);
					rowCount--;
					i--;
				}


			}
			}catch(e) {
				alert(e);
			}
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
 if (document.busca.tp_ocor.value == "999999")
  {    alert("Por favor selecione um tipo de ocorrência!")
   document.busca.tp_ocor.focus()
    return false
 }aula = document.busca.aula.value;
     if (aula.length > 3)
  {    alert("O valor do campo Aula deve possuir menos que 3 caracteres")
    document.busca.aula.focus()
    return false
  }
//    if (document.busca.observacao.value == "")
//  {    alert("Por favor digite uma observação!")
//    document.busca.observacao.focus()
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
      <% if opt="listall" or opt="list" then%>
      <body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
      <%else %>
      <body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
      <%end if %>
      <%call cabecalho(nivel)
%>
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
        <form action="confirma.asp?opt=inc" method="post" name="frm2" id="busca" onSubmit="return checksubmit()">
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
                <td width="493" valign="top"><select name="dia_de" id="dia_de" class="select_style">
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
                    <select name="mes_de" id="mes_de" class="select_style">
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
                    <select name="ano_de" class="select_style" id="ano_de">
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
                <td valign="top" bgcolor="#FFFFFF"><input name="valor" type="text" class="textInput" id="valor" size="35" maxlength="30" onKeyDown="return (currencyFormat(this))" onKeyUp="return (currencyFormat(this))"></td>
              </tr>
                <tr>
                <td colspan="5">&nbsp;</td>
              </tr>
                <tr>
                <td colspan="5" class="tb_tit"
>Composi&ccedil;&atilde;o da Nota Fiscal</td>
              </tr>
                <tr>
                <td colspan="5"><table width="100%" id="dataTable" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="100" align="right" class="form_dado_texto">1</td>
                    <td width="100" align="right" class="form_dado_texto">Item:</td>
                    <td width="167"><select name="item_fornecedor" class="select_style" id="item_fornecedor">
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
                    <td width="184"><input name="quantidade" type="text" class="textInput" id="quantidade" size="15" maxlength="15"></td>
                    <td width="184" align="right" class="form_dado_texto">Valor Unit&aacute;rio:</td>
                    <td width="184"><input name="valor" type="text" class="textInput" id="valor" size="20" maxlength="20" onKeyDown="return (currencyFormat(this))" onKeyUp="return (currencyFormat(this))">
                        <input name="teste" readonly type="hidden"></td>
                    <td width="184"><a href="#" onclick="addRow('dataTable')"><img src="../../../../img/add.png" alt="Adicionar Item" width="20" height="20" border = "0"></a></td>
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
                            <input name="Submit" type="submit" class="botao_prosseguir" value="Prosseguir">
                      </div></td>
                      </tr>
                  </table>
                    <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> </font></div></td>
              </tr>
              </table></td>
          </tr>
        </form>
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