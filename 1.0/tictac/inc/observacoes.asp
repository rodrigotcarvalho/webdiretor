<%'On Error Resume Next%>
<!--#include file="parametros.asp"-->
<!--#include file="caminhos.asp"-->
<!--#include file="funcoes.asp"-->
<!--#include file="funcoes2.asp"-->
      <!--#include file="funcoes7.asp"-->
      <%
Server.ScriptTimeout = 1800 'valor em segundos
%>
      <%
opt = request.QueryString("opt")

if opt = "ok" or opt = "ok1" then
	obr = request.QueryString("obr")
	
	dados = split(obr,"$!$")
	cod_cons = 	dados(1)
	co_materia = dados(0)
	periodo  = dados(6)	
	todos = dados(9)	 
else

	cod_cons = request.QueryString("cod_cons")
	co_materia = request.QueryString("obr")
	periodo = request.QueryString("prd")
end if	
autoriza=session("autoriza")
nvg = session("chave")
ano_letivo = session("ano_letivo")
co_usr = session("co_user")
grupo=session("grupo")
chave=nvg
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
nivel=1
trava=session("trava")
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

if sistema_local="WN" then
	endereco_origem="../wn/lancar/notas/lancar/"
elseif sistema_local="WA" then	
	if funcao="EPN" then
		endereco_origem="../wa/professor/relatorio/epn/"
	else
		endereco_origem="../wa/professor/cna/notas/"
	end if
end if	
	co_prof=session("co_prof")

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CONa= Server.CreateObject("ADODB.Connection") 
		ABRIRer = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONa.Open ABRIRer		

		Set CONg = Server.CreateObject("ADODB.Connection") 
		ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONg.Open ABRIRg	
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		
		
 call navegacao (CON,chave,4)	
navega=Session("caminho") 	

			Set RS = Server.CreateObject("ADODB.Recordset")
			SQL ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& cod_cons&" and TB_Matriculas.NU_Ano="&ano_letivo
			Set RS = CONa.Execute(SQL)

				nu_chamada = RS("NU_Chamada")
				nome_aluno=RS("NO_Aluno")					
				unidade = RS("NU_Unidade")
				curso = RS("CO_Curso")
				etapa = RS("CO_Etapa")
				turma = RS("CO_Turma")
	tb=tabela_nota(ano_letivo,unidade,curso,etapa,turma,"tb",0)		
	opcao=tabela_nota(ano_letivo,unidade,curso,etapa,turma,"opt",0)		
	cam=tabela_nota(ano_letivo,unidade,curso,etapa,turma,"cam",0)	
	coordenador = tabela_nota(ano_letivo,unidade,curso,etapa,turma,"coo",0)	
	
	Set CON_N = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& cam & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_N.Open ABRIR	
		
	action=verifica_dados_tabela(opcao,"action",outro)			
	
	action = Right(action, 7)

obr=co_materia&"$!$"&unidade&"$!$"&curso&"$!$"&etapa&"$!$"&turma&"$!$"&periodo&"$!$"&ano_letivo&"$!$"&co_prof&"$!$"&cod_cons	
%>
      <html>
      <head>
      <title>Web Diretor</title>
      <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
      <link rel="stylesheet" href="../estilos.css" type="text/css">
      <meta http-equiv="X-UA-Compatible" content="IE=edge" />
      <script language="JavaScript" type="text/JavaScript">
<!--

function MM_popupMsg(msg) { //v1.0
  alert(msg);
}
//-->
</script>
      <script language="JavaScript" type="text/JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
      <script language="JavaScript" type="text/JavaScript">
<!--
function setSelectionRange(input, selectionStart, selectionEnd) {
  if (input.setSelectionRange) {
    input.focus();
    input.setSelectionRange(selectionStart, selectionEnd);
  }
  else if (input.createTextRange) {
    var range = input.createTextRange();
    range.collapse(true);
    range.moveEnd('character', selectionEnd);
    range.moveStart('character', selectionStart);
    range.select();
  }
}

function replaceSelection (input, replaceString) {
	if (input.setSelectionRange) {
		var selectionStart = input.selectionStart;
		var selectionEnd = input.selectionEnd;
		input.value = input.value.substring(0, selectionStart)+ replaceString + input.value.substring(selectionEnd);
    
		if (selectionStart != selectionEnd){ 
			setSelectionRange(input, selectionStart, selectionStart + 	replaceString.length);
		}else{
			setSelectionRange(input, selectionStart + replaceString.length, selectionStart + replaceString.length);
		}

	}else if (document.selection) {
		var range = document.selection.createRange();

		if (range.parentElement() == input) {
			var isCollapsed = range.text == '';
			range.text = replaceString;

			 if (!isCollapsed)  {
				range.moveStart('character', -replaceString.length);
				range.select();
			}
		}
	}
}


// We are going to catch the TAB key so that we can use it, Hooray!
function catchTab(item,e){
	if(navigator.userAgent.match("Gecko")){
		c=e.which;
	}else{
		c=e.keyCode;
	}
	if(c==9){
		replaceSelection(item,String.fromCharCode(9));
		setTimeout("document.getElementById('"+item.id+"').focus();",0);	
		return false;
	}
		    
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

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresiz!=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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
  

function MM_callJS(jsStr) {
  return eval(jsStr)
}
function checkTheBox() {
   var chk = document.getElementsByName('todos')
    var len = chk.length

    for(i=0;i<len;i++)
    {
         if(chk[i].checked){
			 var r=confirm("Deseja salvar o conteúdo dessa avaliação para todos os alunos dessa turma?");
				if (r==true)
				  {
				  return true;
				  }
				else
				  {
				  //return false;
				  chk[i].checked = false;
				  return true;
				  }
			 

          }
    }	
    
    }	
//-->
var tam = 15; 
function mudaFonte(tipo,elemento){ 
var textarea = document.getElementById("elm1");
var textsize = textarea.cols;
//alert(textsize);
if (textsize>130){textsize = 130}
	if (tipo=="mais") { 
		if(tam<24) tam+=1;
		textsize = textsize-10
		createCookie("fonte",tam,365); } 
	else { 
		if(tam>10) tam-=1; 
		createCookie("fonte",tam,365); 
		textsize = textsize+10	
		} 
		textarea.style.fontSize = tam+"px"; 
		
	    textarea.cols = textsize;		
//alert(textsize);	
} 
function createCookie(name,value,days) { 
if (days) { var date = new Date(); 
	date.setTime(date.getTime()+(days*24*60*60*1000)); 
	var expires = "; expires="+date.toGMTString(); } 
else 
	var expires = "";
 document.cookie = name+"="+value+expires+"; path=/"; 
 } 
 
 function readCookie(name) { 
 var nameEQ = name + "="; 
 var ca = document.cookie.split(";"); 
 for(var i=0;i < ca.length;i++) { 
 	var c = ca[i]; while (c.charAt(0)==" ") c = c.substring(1,c.length); 
	if (c.indexOf(nameEQ) == 0) 
	return c.substring(nameEQ.length,c.length); 
	} return null; 
}



</script>

	<script type="text/javascript">
		// simple AJAX routine
		function createAjaxObj()
		{
			var httprequest = false;
			if (window.XMLHttpRequest)
			{
				// if Mozilla, Safari etc
				httprequest = new XMLHttpRequest();
				if (httprequest.overrideMimeType)
				{
					httprequest.overrideMimeType('text/xml');
				}
			}
			else if (window.ActiveXObject)
			{
				// if IE
				try
				{
					httprequest = new ActiveXObject("Msxml2.XMLHTTP");
				}
				catch (e)
				{
					try
					{
						httprequest = new ActiveXObject("Microsoft.XMLHTTP");
					}
					catch (e){}
				}
			}
			return httprequest;
		}

		var ajaxpack = {};
		ajaxpack.ajaxobj = createAjaxObj();
		ajaxpack.filetype = "txt";

		ajaxpack.postAjaxRequest = function(url, parameters, callbackfunc, filetype)
		{
			ajaxpack.ajaxobj = createAjaxObj(); //recreate ajax object to defeat cache problem in IE
			if (this.ajaxobj)
			{
				this.filetype = filetype;
				this.ajaxobj.onreadystatechange = callbackfunc;
				this.ajaxobj.open('POST', url, true);
				this.ajaxobj.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
				this.ajaxobj.setRequestHeader("Content-length", parameters.length);
				this.ajaxobj.setRequestHeader("Connection", "close");
				this.ajaxobj.send(parameters);
			}
		}
	</script>
<script type="text/javascript">
	window.onload = function()
	{
		setInterval(function() {
			var str = '';
			var poststr = '';
			var elems = document.avaliacao.elements;
			var incluiEComercial = 'N';			
			for (var i=0; i<elems.length; i++)
			{	str= escape(elems[i].value); 
		 
				while (str.indexOf("/")!=-1) 
				{ 
					str = str.replace("/","%2F"); 
				} 
								
				if (elems[i].name == "todos"){
					 if (elems[i].checked == true){ 					
						poststr += (i != 0) ? "&" : "";
						//poststr += encodeURI(elems[i].name) + "=" + encodeURI(elems[i].value);
						poststr += escape(elems[i].name) + "=" + str;	
						incluiEComercial = 'S';
					 }						
				} else {
					if (incluiEComercial == 'S'){
						poststr += (i != 0) ? "&" : "";
					}
					
					
					//poststr += encodeURI(elems[i].name) + "=" + encodeURI(elems[i].value);
					poststr += escape(elems[i].name) + "=" + str;					
					incluiEComercial = 'S';										
				}						
			}
			ajaxpack.postAjaxRequest('<%response.write(action)%>?opt=ajax', poststr, null, 'txt');
			alert('Avaliação salva automaticamente!'); 
		}, 600000);
	};
	
function submitAjax(){
	
			var str = '';
			var poststr = '';
			var elems = document.avaliacao.elements;
			var incluiEComercial = 'N';			
			for (var i=0; i<elems.length; i++)
			{	str= escape(elems[i].value); 
		 
				while (str.indexOf("/")!=-1) 
				{ 
					str = str.replace("/","%2F"); 
				} 
								
				if (elems[i].name == "todos"){
					 if (elems[i].checked == true){ 					
						poststr += (i != 0) ? "&" : "";
						//poststr += encodeURI(elems[i].name) + "=" + encodeURI(elems[i].value);
						poststr += escape(elems[i].name) + "=" + str;	
						incluiEComercial = 'S';
					 }						
				} else {
					if (incluiEComercial == 'S'){
						poststr += (i != 0) ? "&" : "";
					}
					
					
					//poststr += encodeURI(elems[i].name) + "=" + encodeURI(elems[i].value);
					poststr += escape(elems[i].name) + "=" + str;					
					incluiEComercial = 'S';										
				}						
			}
			ajaxpack.postAjaxRequest('<%response.write(action)%>?opt=ajax', poststr, null, 'txt');
			alert('Avaliação salva!'); 				
}
</script>
      </head>
      
      <body role="application" leftmargin="0" topmargin="0" marginwidth="0" background="../../../../img/fundo.gif" marginheight="0">
      <%

IF imprime="1"then
else
 call cabecalho (nivel) 
 end if
 

 %>
      <table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
        <tr>
          <td height="10" valign="top" class="tb_caminho"><font class="style-caminho">
            <%if sistema_local = "WA" then%>
            Voc&ecirc; est&aacute; em <a href='../inicio.asp' class='caminho' target='_self'>Web Diretor</a> > <a href='../wa/index.asp?nvg=WA' class='caminho' target='_self'>Web Acadêmico</a> > <a href='../wa/professor/index.asp?nvg=WA-PF'class='caminho' target='_self'>Professores</a> > <a href='../wa/professor/cna/index.asp?nvg=WA-PF-CN' class='caminho' target='_self'>Controle das Avaliações</a> > <a href='../wa/professor/cna/notas/index.asp?nvg=WA-PF-CN-MNL' class='caminho' target='_self'>Modificar Avaliações Lançadas</a>
            <%else%>
            Voc&ecirc; est&aacute; em <a href='../inicio.asp' class='caminho' target='_self'>Web Diretor</a> > <a href='../wn/index.asp?nvg=WN' class='caminho' target='_self'>Web Professor</a> > <a href='../wn/lancar/index.asp?nvg=WN-LN'class='caminho' target='_self'>Central de Provas</a> > <a href='../wn/lancar/notas/index.asp?nvg=WN-LN-LN' class='caminho' target='_self'>Avaliações dos Alunos</a> > <a href='../wn/lancar/notas/lancar/index.asp?nvg=WN-LN-LN-LAN' class='caminho' target='_self'>Registrar Avaliações</a>
            <%end if%>
            </font></td>
        </tr>
        <%


no_unidade= GeraNomes("U",unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)
no_curso=GeraNomes("C",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro) 	
no_etapa=GeraNomes("E",curso,etapa,variavel3,variavel4,variavel5,CON0,outro) 	
no_materia= GeraNomes("D",co_materia,variavel2,variavel3,variavel4,variavel5,CON0,outro) 	
tp_modelo=tipo_divisao_ano(curso,etapa,"tp_modelo")	


nome_prof = session("nome_prof") 
tp=	session("tp")

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("m", now) 
data = dia &"/"& mes &"/"& ano
horario = hora & ":"& min
acesso_prof = session("acesso_prof")



		
		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' AND CO_Materia_Principal = '"& co_materia &"'"
		Set RS = CONg.Execute(CONEXAO)
periodo=periodo*1
if periodo=1 then
	ST_Per_1 = RS("ST_Per_1")
elseif periodo=2 then
	ST_Per_2 = RS("ST_Per_2")
elseif periodo=3 then
	ST_Per_3 = RS("ST_Per_3")
elseif periodo=4 then
	ST_Per_4 = RS("ST_Per_4")
elseif periodo=5 then
	ST_Per_5 = RS("ST_Per_5")
elseif periodo=6 then
	ST_Per_6 = RS("ST_Per_6")
end if
tp = session("tp")
%>
        <%if opt = "ok" then%>
        <tr>
          <td height="10" valign="top"><%
		call mensagens(nivel,622,2,0)	
%>
            <div align="center"></div></td>
        </tr>
        <%elseif opt= "ok1" then %>
        <tr>
          <td height="10" valign="top"><%
	call mensagens(nivel,622,2,"T")
%></td>
        </tr>
        <%elseif opt= "err6" then %>
        <tr>
          <td height="10" valign="top"><%
	call mensagens(nivel,1000,1,0)
%></td>
        </tr>
        <%end if
%>
        <% IF trava="s" or (co_usr<>coordenador AND grupo="COO") then%>
        <tr>
          <td height="10" valign="top"><%
	 	 call mensagens(nivel,9701,0,0)
	  %></td>
        </tr>
        <% 
		ELSEIF (co_usr=coordenador) AND trava<>"s" AND ((periodo = 1 and ST_Per_1="x") OR (periodo = 2 and ST_Per_2="x") OR (periodo = 3 and ST_Per_3="x") OR (periodo = 4 and ST_Per_4="x") OR (periodo = 5 and ST_Per_5="x") OR (periodo = 6 and ST_Per_6="x")) then%>
        <tr>
          <td height="10" valign="top"><%
	 	 call mensagens(nivel,640,1,0)	  
	  %></td>
        </tr>
        <%elseif (periodo = 1 and ST_Per_1="x") OR (periodo = 2 and ST_Per_2="x") OR (periodo = 3 and ST_Per_3="x") OR (periodo = 4 and ST_Per_4="x") OR (periodo = 5 and ST_Per_5="x") OR (periodo = 6 and ST_Per_6="x") then%>
        <tr>
          <td height="10" valign="top"><%
	 	 call mensagens(nivel,624,1,0)
%></td>
        </tr>
        <% end if%>
        <%if opt= "cln" then %>
        <tr>
          <td height="10" valign="top"><%
	call mensagens(nivel,621,0,0)			
%></td>
        </tr>
        <% end if%>
        <tr>
          <td height="10" valign="top"><%
	 	 	call mensagens(nivel,645,0,"R16")			  

%></td>
        </tr>
        <tr class="tb_tit">
          <td height="15" class="tb_tit">&nbsp;Grade de Aulas</td>
        </tr>
        <tr>
          <td height="36" valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr>
              <td width="166" class="tb_subtit"><div align="center"><strong>UNIDADE </strong></div></td>
              <td width="166" class="tb_subtit"><div align="center"><strong>CURSO </strong></div></td>
              <td width="166" class="tb_subtit"><div align="center"><strong>ETAPA </strong></div></td>
              <td width="166" class="tb_subtit"><div align="center"><strong>TURMA </strong></div></td>
              <td width="170" class="tb_subtit"><div align="center"><strong>DISCIPLINA</strong></div></td>
              <td width="166" class="tb_subtit"><div align="center"><strong>PER&Iacute;ODO </strong></div></td>
            </tr>
              <tr>
              <td width="166"><div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
                <%response.Write(no_unidade)%>
                </font></div></td>
              <td width="166"><div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
                <%response.Write(no_curso)%>
                </font></div></td>
              <td width="166"><div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
                <%
response.Write(no_etapa)%>
                </font></div></td>
              <td width="166"><div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
                <%
response.Write(turma)%>
                </font></div></td>
              <td width="170"><div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
                <%

response.Write(no_materia)%>
                </font> </div></td>
              <td width="166"><div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif">
                <%
		Set RSper = Server.CreateObject("ADODB.Recordset")
		SQLper = "SELECT * FROM TB_Periodo where TP_Modelo='"&tp_modelo&"' AND NU_Periodo= "&periodo
		RSper.Open SQLper, CON0

NO_Periodo= RSper("NO_Periodo")
response.Write(NO_Periodo)%>
                </font> </div></td>
            </tr>
            </table></td>
        </tr>
        <tr>
          <td valign="top"><%
	
 %>
            <form name="avaliacao" method="post" action="<%response.Write(action)%>" onSubmit="return checkTheBox()">
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td valign="top"><table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                      <td colspan="3"><hr></td>
                    </tr>
                      <tr>
                      <td width="150" align="center" class="tb_subtit">CHAMADA</td>
                      <td width="150" align="center" class="tb_subtit">MATR&Iacute;CULA</td>
                      <td width="700" align="center" class="tb_subtit">NOME</td>
                    </tr>
                      <tr>
                      <td width="150" align="center" class="form_dado_texto"><%response.Write(nu_chamada)%></td>
                      <td width="150" align="center" class="form_dado_texto"><%response.Write(cod_cons)%></td>
                      <td width="700" align="center" class="form_dado_texto"><%response.Write(nome_aluno)%></td>
                    </tr>
                      <tr>
                      <td colspan="3" align="center" class="form_dado_texto">&nbsp;</td>
                    </tr>
                      <tr>
                      <td align="center" class="form_dado_texto"><%if todos = "S" then
		%>
                          <input name="todos" type="checkbox" id="todos" value="S" checked>
                          <%		
		else
		%>
                          <input name="todos" type="checkbox" id="todos" value="S">
                          <%end if%></td>
                      <td colspan="2" align="left" class="form_dado_texto">Gravar esta avalia&ccedil;&atilde;o para todos os alunos da turma</td>
                    </tr>
                    </table></td>
                </tr>
                <tr>
                  <td align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="94%" align="right"><a href="javascript:void(0);" Onclick="mudaFonte('mais'); return false"> <img src="../img/Zoom-Plus.png" border="0"></a> <a href="javascript:void(0);" Onclick="mudaFonte('menos'); return false"> <img src="../img/Zoom-Minus.png" border="0"></a></td>
                      <td width="6%">&nbsp;</td>
                    </tr>
                  </table>
                  </td>
                </tr>
                <tr>
                  <td align="center"><div align="center">
                      <%
		Set RSMT  = Server.CreateObject("ADODB.Recordset")
		SQL_MT  = "Select CO_Materia_Principal from TB_Materia WHERE CO_Materia = '"& co_materia&"'"
		Set RSMT  = CON0.Execute(SQL_MT)
		
co_materia_pr = RSMT("CO_Materia_Principal")
		
if Isnull(co_materia_pr) or co_materia_pr= "" then
	co_materia_pr= co_materia
else
	co_materia_pr = co_materia_pr
end if	
	
	
	Set RS3 = Server.CreateObject("ADODB.Recordset")
	SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& cod_cons & " AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
	Set RS3 = CON_N.Execute(SQL_N)			
	
	if RS3.EOF then
		avaliacao = ""
	else
		avaliacao = RS3("TX_Avalia")
	end if	

	periodo=periodo*1

if ((periodo = 1 and (isnull(ST_Per_1) or ST_Per_1<>"x" )) OR (periodo = 2 and (isnull(ST_Per_2) or ST_Per_2<>"x")) OR (periodo = 3 and ST_Per_3<>"x") OR (periodo = 4 and ST_Per_4<>"x") OR (periodo = 5 and ST_Per_5<>"x") OR (periodo = 6 and ST_Per_6<>"x")) AND (grupo="PRO" and autoriza=5 AND trava="n") or (grupo<>"PRO" and autoriza=5 AND trava="n") then
	botao_salvar = "S"	 %>
<textarea id="elm1" name="elm1" cols="180" rows="50" onKeyDown="return catchTab(this,event)"><%response.Write(avaliacao)%>
</textarea>    
                      
                      <%else
	botao_salvar = "N"	%><textarea id="elm1" name="elm1" cols="180" rows="50" disabled><%response.Write(avaliacao)%>
</textarea>

                      
                      <%end if%>
                    </div></td>
                </tr>
                <tr>
                  <td><hr></td>
                </tr>
                <tr>
                  <td><table width="100%" border="0" align="center" cellspacing="0">
                      <tr>
                      <td width="33%"><div align="center">
                          <input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','<%response.Write(endereco_origem)%>notas.asp?opt=vt&obr=<%response.Write(obr)%>');return document.MM_returnValue" value="Voltar">
                        </div></td>
                      <td width="34%"><div align="center"> 
                          <!--<input name="Submit" type="button" class="botao_prosseguir_comunicar" onClick="MM_goToURL('parent','notas.asp?or=01&opt=cln&obr=<%=obr%>');return document.MM_returnValue" value="Comunicar ao Coordenador T&eacute;rmino da Planilha">--> 
                        </div></td>
                      <td width="33%"><div align="center">
                          <% if botao_salvar = "S" then%>
                          <input type="button" name="button" value="Salvar" class="botao_prosseguir" onClick="submitAjax()">
                          <%end if%>
                          <input name="cod_cons" type="hidden" id="cod_cons" value="<%response.Write(cod_cons)%>">
                          <input name="unidade" type="hidden" id="unidade" value="<%response.Write(unidade)%>">
                          <input name="curso" type="hidden" id="curso" value="<%response.Write(curso)%>">
                          <input name="etapa" type="hidden" id="etapa" value="<%response.Write(etapa)%>">
                          <input name="turma" type="hidden" id="turma" value="<%response.Write(turma)%>">
                          <input name="co_materia" type="hidden" id="co_materia" value="<%response.Write(co_materia)%>">
                          <input name="periodo" type="hidden" id="periodo" value="<%response.Write(periodo)%>">
                          <input name="co_prof" type="hidden" id="co_prof" value="<%response.Write(co_prof)%>">
                          <input name="max" type="hidden" id="max" value="<%response.Write(max)%>">
                          <input name="co_usr" type="hidden" id="co_usr" value="<%response.Write(co_usr)%>">
                          <input name="ano_letivo" type="hidden" id="ano_letivo" value="<%response.Write(ano_letivo)%>">
                        </div></td>
                    </tr>
                    </table></td>
                </tr>
              </table>
            </form></td>
        </tr>
        <%	
Set RS = Nothing
Set RS2 = Nothing
Set RS3 = Nothing

%>
        <tr>
          <td height="40" valign="top"><img src="../img/rodape.jpg" width="1000" height="40"></td>
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
response.redirect("erro.asp")
end if
%>